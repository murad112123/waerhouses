#!/usr/bin/env python3

from os import path, walk, remove, makedirs
import string
import random
from shutil import copy2
from subprocess import run

import aes

RANDOM_CHARS = string.ascii_letters + string.digits
APP_DIR = path.dirname(path.realpath(__file__))
ASSETS_DIR = path.join(APP_DIR, 'assets')


def list_all_assets():
    return next(walk(ASSETS_DIR))[2]


def asset_exists(uid, asset_id):
    uid = str(uid)
    if asset_id and uid:
        asset_path = path.join(ASSETS_DIR, uid, asset_id)
        return path.exists(asset_path)


def asset_delete(uid, asset_id):
    uid = str(uid)
    if asset_id and uid and asset_exists(uid, asset_id):
        remove(get_asset_path(uid, asset_id))


def create_asset(uid, filepath, encrypt_key=None):
    uid = str(uid)
    if uid and not path.exists(path.join(ASSETS_DIR, uid)):
        makedirs(path.join(ASSETS_DIR, uid))

    if uid and path.exists(filepath):
        asset_id = path.basename(filepath) + ''.join(random.choice(RANDOM_CHARS) for i in range(16))
        asset_path = path.join(ASSETS_DIR, uid, asset_id)
        run(["mat2", "--inplace", filepath])
        if encrypt_key:
            aes.AESCipher(encrypt_key).encrypt_file(filepath, asset_path)
        else:
            copy2(filepath, asset_path)
        return asset_id


def get_asset_path(uid, asset_id):
    uid = str(uid)
    if uid and asset_id:
        return path.join(ASSETS_DIR, uid, asset_id)


def get_asset_file(uid, asset_id):
    uid = str(uid)
    if asset_id and asset_exists(uid, asset_id):
        asset_path = get_asset_path(uid, asset_id)
        return open(asset_path, 'rb')


def get_asset_data(uid, asset_id):
    uid = str(uid)
    if asset_id and asset_exists(uid, asset_id):
        with get_asset_file(uid, asset_id) as asset:
            return asset.read()


def decrypt_asset(uid, asset_id, key):
    uid = str(uid)
    if uid and asset_id:
        tmp = f'/tmp/prisons-{asset_id}-' + ''.join(random.choice(RANDOM_CHARS) for i in range(16))
        aes.AESCipher(key).decrypt_file(get_asset_path(uid, asset_id), tmp)
        return tmp
