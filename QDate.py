import PyQt5.QtCore as QtCore

date_format = 'yyyy-MM-dd'


class QDate(QtCore.QDate):
    def __init__(self, date):
        date = QtCore.QDate().fromString(date, date_format) if date else QtCore.QDate(1900, 1, 1)
        super().__init__(date)

    @staticmethod
    def fromString(date) -> QtCore.QDate:
        return QtCore.QDate().fromString(date, date_format) if date else QtCore.QDate(1900, 1, 1)

    @staticmethod
    def toString(date: QtCore.QDate):
        return '' if (date.day() == 1 and date.month() == 1 and date.year() == 1900) else date.toString(date_format)
