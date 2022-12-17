 #! /usr/bin/env python3
import aes
pass_key = '123'
with open('config.json', 'r') as plain_config:
    aes_cipher = aes.AESCipher(pass_key)
    config = aes_cipher.encrypt(plain_config.read())
    with open('config.dat', 'w') as encrypted_config:
        encrypted_config.write(config)
