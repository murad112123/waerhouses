from sqlcipher3 import dbapi2 as sqlite

db = None


def _dict_factory(cursor, row):
    d = {}
    for idx, col in enumerate(cursor.description):
        value = row[idx]
        key = col[0]
        if isinstance(value, int):
            d[key] = str(value)
        elif isinstance(value, float):
            d[key] = str(value)
        elif value:
            d[key] = value
        else:
            d[key] = ''
    return d


class Database:
    def __init__(self, database_path, password):
        self.connection = sqlite.connect(database_path)
        self.connection.execute('PRAGMA key=' + password)
        self.connection.execute('PRAGMA foreign_keys = ON')
        self.connection.row_factory = _dict_factory

    @staticmethod
    def open_database(password):
        global db
        db = Database('warehouses.db', password)

    def change_user_pass(self, username, password):
        self.connection.execute('UPDATE users SET pass = ? WHERE users.name = ?', (password, username))
        self.connection.commit()

    def is_user(self, username):
        return self.connection.execute("select id, permission from branches where code = ?", (username,)).fetchone()

    def count_row(self, table, r):
        if r == 1:
            return self.connection.execute(f'select count(*) as count from {table}').fetchone()['count']
        else:
            return self.connection.execute(f'select count(*) as count from {table} where code = ?', (r,)).fetchone()[
                'count']

    def count_quantity_branch(self, table, bid, mid):
        return self.connection.execute(f"select count(*) as count from {table} where b_id = ? and m_id = ?", (bid, mid)).fetchone()['count']

    def get_next_id(self, table):
        if self.connection.execute(f"select max(id)+1 as seq from {table}").fetchone()['seq'] == '':
            return 1
        return self.connection.execute(f"select max(id)+1 as seq from {table}").fetchone()['seq']

    def query_row(self, table, id):
        return self.connection.execute(f"select * from {table} where id = ?", (id, )).fetchone()

    def query_csp(self, table):
        return {e['id']: e['code'] for e in self.connection.execute(f'select id, code from {table}').fetchall()}

    def query_req(self):
        return {e['id']: e['code'] for e in self.connection.execute('select id, code from requests').fetchall()}

    def get_id_by_code(self, table, code):
        return self.connection.execute(f"select id from {table} where code = ?", (code,)).fetchone()['id']

    def get_id_by_mid(self, table, mid, bid):
        return self.connection.execute(f"select id from {table} where m_id = ? and branch_id = ?", (mid, bid)).fetchone()['id']

    def get_code_by_id(self, table, id):
        return self.connection.execute(f'select code from {table} where id = ?', (id,)).fetchone()['code']

    #
    # def count_table(self, table, id):
    #     return self.connection.execute(f"SELECT count(*) as count FROM {table} where id = '{id}'").fetchone()['count']
    #
    # def insert_table(self, table, dic):
    #     new_ids = [int(d['id']) for d in dic]
    #     placeholders = ", ".join("?" * len(new_ids))
    #     del_ids = self.connection.execute(f"SELECT id FROM {table} WHERE branch_id = {dic[0]['branch_id']} and id not in ({placeholders})", tuple(new_ids)).fetchall()
    #     for d in dic:
    #         if self.count_table(table, d['id']) == '1':
    #             self.update_row(table, d)
    #         else:
    #             self.insert_row(table, d)
    #     for d in del_ids:
    #         self.delete_row(table, d['id'])
    #
    def insert_row(self, table, row):
        def _insert(obj):
            columns = ', '.join(obj.keys())
            placeholders = ':' + ', :'.join(obj.keys())
            query = f'INSERT INTO {table} ({columns}) VALUES ({placeholders})'
            self.connection.execute(query, obj)
            self.connection.commit()

        if isinstance(row, dict):
            _insert(row)
        elif isinstance(row, list):
            for d in row:
                _insert(d)

    def update_row(self, table, row):
        def _update(obj):
            placeholders = ', '.join([f'{key}=:{key}' for key in obj.keys()])
            query = f"UPDATE {table} SET {placeholders} WHERE id = ?"
            li = list(obj.values())
            li.append(obj['id'])
            self.connection.execute(query, tuple(li))
            self.connection.commit()

        if isinstance(row, dict):
            _update(row)
        elif isinstance(row, list):
            for d in row:
                _update(d)

    def delete_row(self, table, id):
        self.connection.execute(f'delete from {table} where id = ?', (id,))
        self.connection.commit()

    def get_material_by_code(self, code):
        return self.connection.execute(f"select id, name, description, price, link from material where code = ?", (code, )).fetchone()

    def query_all_material(self, filter: dict, limit1, limit2):
        sql_cmd = "SELECT id, code, name, description, type, price from material"

        if filter:
            sql_cmd += " where "
            filter_cmd = []
            if 'code' in filter:
                filter['code'] = f'%{filter["code"]}%'
                filter_cmd.append(f'code like :code')
            if 'name' in filter:
                filter['name'] = f'%{filter["name"]}%'
                filter_cmd.append(f'name like :name')
            if 'type' in filter:
                filter['type'] = f'%{filter["type"]}%'
                filter_cmd.append(f'type like :type')

            sql_cmd += ' and '.join(filter_cmd)
            sql_cmd += f' limit {limit1}, {limit2}'
            return self.connection.execute(sql_cmd, filter).fetchall()
        else:
            sql_cmd += f' limit {limit1}, {limit2}'
            return self.connection.execute(sql_cmd).fetchall()

    def query_all_material_branch(self, table, filter: dict, limit1, limit2):
        sql_cmd = f"SELECT id, b_id, m_id, quantity, place from {table}"
        if filter:
            sql_cmd += " where "
            filter_cmd = []
            if 'b_id' in filter:
                filter_cmd.append(f'b_id =:b_id')
            if 'm_code' in filter:
                filter['m_code'] = f'%{filter["m_code"]}%'
                filter_cmd.append(f'm_id in (select id from material where code like :m_code)')

            sql_cmd += ' and '.join(filter_cmd)
            sql_cmd += f' limit {limit1}, {limit2}'
            return self.connection.execute(sql_cmd, filter).fetchall()
        else:
            sql_cmd += f' limit {limit1}, {limit2}'
            return self.connection.execute(sql_cmd).fetchall()

    # product
    # ###############################################################

    def query_all_product(self, filter1: dict, limit1, limit2):
        sql_cmd = "SELECT id, code, name, description, price, cost from product"

        if filter1:
            sql_cmd += " where "
            filter_cmd = []
            if 'code' in filter1:
                filter1['code'] = f'%{filter1["code"]}%'
                filter_cmd.append(f'code like :code')
            if 'name' in filter1:
                filter1['name'] = f'%{filter1["name"]}%'
                filter_cmd.append(f'name like :name')

            sql_cmd += ' and '.join(filter_cmd)
            sql_cmd += f' limit {limit1}, {limit2}'
            return self.connection.execute(sql_cmd, filter1).fetchall()
        else:
            sql_cmd += f' limit {limit1}, {limit2}'
            return self.connection.execute(sql_cmd).fetchall()

    # def query_all_cs(self, table, filter: dict, limit1, limit2):
    #     sql_cmd = f"SELECT id, code, name, phone, balance from {table}"
    #
    #     if filter:
    #         sql_cmd += " where "
    #         filter_cmd = []
    #         if 'code' in filter:
    #             filter['code'] = f'%{filter["code"]}%'
    #             filter_cmd.append(f'code like :code')
    #         if 'name' in filter:
    #             filter['name'] = f'%{filter["name"]}%'
    #             filter_cmd.append(f'name like :name')
    #
    #         sql_cmd += ' and '.join(filter_cmd)
    #         sql_cmd += f' limit {limit1}, {limit2}'
    #         return self.connection.execute(sql_cmd, filter).fetchall()
    #     else:
    #         sql_cmd += f' limit {limit1}, {limit2}'
    #         return self.connection.execute(sql_cmd).fetchall()
    #
    # # ################################################################
    #
    # def query_all_fm(self, filter: dict, limit1, limit2):
    #     sql_cmd = f"SELECT * from fund_movement"
    #
    #     if filter:
    #         sql_cmd += " where "
    #         filter_cmd = []
    #         if 'type' in filter:
    #             filter_cmd.append(f'type = :type')
    #         if 'owner' in filter:
    #             filter_cmd.append(f'owner = :owner')
    #         if 'date_from' in filter:
    #             if 'date_to' in filter:
    #                 filter_cmd.append(f'date between :date_from and :date_to')
    #             else:
    #                 filter_cmd.append(f'date = :date_from')
    #         if 'note' in filter:
    #             filter['note'] = f'%{filter["note"]}%'
    #             filter_cmd.append(f'note like :note')
    #         sql_cmd += ' and '.join(filter_cmd)
    #         sql_cmd += f' limit {limit1}, {limit2}'
    #         return self.connection.execute(sql_cmd, filter).fetchall()
    #     else:
    #         sql_cmd += f' limit {limit1}, {limit2}'
    #         return self.connection.execute(sql_cmd).fetchall()
    #
    # # ################################################################
    #
    def query_all_bill(self, bill_type, filter: dict, limit1, limit2):
        sql_cmd = f"SELECT * from {bill_type}"

        if filter:
            sql_cmd += " where "
            filter_cmd = []
            if 'code' in filter:
                filter['code'] = f'%{filter["code"]}%'
                filter_cmd.append(f'code like :code')
            if 'branch_id' in filter:
                filter_cmd.append(f'branch_id =:branch_id')
            if 'date_from' in filter:
                if 'date_to' in filter:
                    filter_cmd.append(f'date between :date_from and :date_to')
                else:
                    filter_cmd.append(f'date = :date_from')

            sql_cmd += ' and '.join(filter_cmd)
            sql_cmd += f' limit {limit1}, {limit2}'
            return self.connection.execute(sql_cmd, filter).fetchall()
        else:
            sql_cmd += f' limit {limit1}, {limit2}'
            return self.connection.execute(sql_cmd).fetchall()

    def get_product_material(self, table, p_id):
        return self.connection.execute(f"SELECT * FROM {table} WHERE p_id = ?", (p_id,)).fetchall()

    def get_order_requests(self, table, r_id):
        return self.connection.execute(f"SELECT * FROM {table} WHERE req_id = ?", (r_id,)).fetchall()

# def get_noti_pro1(self):
#     return self.connection.execute(f"SELECT code, name, quantity FROM product WHERE quantity <= less_quantity").fetchall()
#
# def get_noti_pro2(self):
#     return self.connection.execute(f"SELECT code, name, quantity FROM product WHERE quantity = 0").fetchall()
#
# def get_noti_cus(self, table):
#     return self.connection.execute(f"SELECT code, name, range_balance FROM {table} WHERE balance >= range_balance").fetchall()

# def get_box(self):
#     return self.connection.execute("SELECT * FROM box").fetchone()
#
# def exchange_dollar_turky(self, to, do, tu):
#     if to == 'do_tu':
#         self.connection.execute("UPDATE box SET dollar = dollar - ?, turky = turky + ?", (do, tu))
#     else:
#         self.connection.execute("UPDATE box SET dollar = dollar + ?, turky = turky - ?", (do, tu))
#     self.connection.commit()
