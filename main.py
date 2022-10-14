# pip install pycrypto
# RSA doc: https://cryptography.io/en/latest/hazmat/primitives/asymmetric/rsa/
# 脚本用处：
#   会使用填在 lianwei_key_pub 里的RSA公钥去加密
#   与它在同一路径下、已保存关闭、扩展名为.xlsx的excel文件里的 第一个sheet、叫为「明文手机号」的列
# 脚本输出：
#   会在excel


from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_public_key, load_pem_private_key
import base64
import os
import pandas as pd

test_key_pub = '''
-----BEGIN PUBLIC KEY-----
MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA5w/xioNB+WRleATGaxkY
xRwegK/IQT3L8VjDyyXl//ILi6UUixBjuMbDzM6J3n000j7OA7tcvYiFu/FkT5O+
ee6/NGQIw5Ty2wy5nDa3D2A1l6WrloH1Ra04ODuUJ2hh0HPp8TgxcNo/z+OI1LuO
I+NU37JxS8Y5FaSV6vDes+k5a9sfC0vosS9PHR/jM3AtzA3piBM+9mscnNpRlert
MlZOUzBKnaaTLUlWtDlhe2U+8nK2HJ+DelXrnynMX6KPxU+hyyhNmMZbYjOVp419
l8cSy/y/ibbd/WG3t2dtpCZqLgKawVK57V7f/uvnvHZzWFtbXyZdu4adQnu9f/RR
IQIDAQAB
-----END PUBLIC KEY-----
'''
test_key_pvt = '''
-----BEGIN PRIVATE KEY-----
MIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDnD/GKg0H5ZGV4
BMZrGRjFHB6Ar8hBPcvxWMPLJeX/8guLpRSLEGO4xsPMzonefTTSPs4Du1y9iIW7
8WRPk7557r80ZAjDlPLbDLmcNrcPYDWXpauWgfVFrTg4O5QnaGHQc+nxODFw2j/P
44jUu44j41TfsnFLxjkVpJXq8N6z6Tlr2x8LS+ixL08dH+MzcC3MDemIEz72axyc
2lGV6u0yVk5TMEqdppMtSVa0OWF7ZT7ycrYcn4N6VeufKcxfoo/FT6HLKE2Yxlti
M5WnjX2XxxLL/L+Jtt39Ybe3Z22kJmouAprBUrntXt/+6+e8dnNYW1tfJl27hp1C
e71/9FEhAgMBAAECggEAd8OCzemk25BXK7NF6SMT/K7LfKYgJPjT6Z+5tGVLZrNd
mp8RG9d96LtVp6VyPpklNMiK3dJSiobl0nmXJcwNkYPXJV+oTz+39SQOXNLbZaPX
g2fCnXt01w2mszbmXtqywGokWvxmW/kz1Bw6wxEH5sAhMOV75euzkO1DK4h31kXT
dNz3Uftiwrjnz2eEwrJ2oguPcAT7SjRDv/QRKwKyWCsfb40H+UJUrwV4bMHPpCbF
qw/6WprLisBlSFggwIz2gnf55h+S8Q135wNdpChWGAE9ZkWcZ17XvOsAIS9wsgGv
MoZ4R+lfsuWGXNZ8dFoK7ov8gzRf0H3T5p7Y6Cb4AQKBgQD006ihllykS3z8N2FX
ymL17NKW8yasrNkc7SyY2inqkI35AGvWAVJFKTvViI8+1VjpXVT2AXBqqaezVLAB
ZojBeja318AJQXNNLsXI2A2drIak4R+2x7EQwsd2V1G4EmNTPToHWio45Bf4En2O
yIXY2cQkM0TQ6CMbWGML6Q/SQQKBgQDxm3jqTb+Ix9oFXKBvlbPSA+q51uqKtkbq
RiY/nBR40SUCjBem8XytUxWnzMImhcZS1yylgZnW+VyjrWizKStbfWnV8PRV+KaK
8gMPCROukZzUJKyV/OylDgcExlFWVHtQAPOKpoSyxg8mjB5VIhp3JRckuoszAHVX
yIJXRF8G4QKBgFtGQbskZJt37TvWpbrmICjRRt2x/vwnYLYxEgxWYYQqqlNnvcxG
J9bS/ZSpWcYyIfi2rAMfHDsXzwbDju6mvFttZdL6Y4TP2t2uj1xGeCUNehEkQP+S
yUeXZmePPE8kw9T3oZe2HMGi//CjbB38UjI7Va2tU32S3evG8v4wwI5BAoGAIRoO
9/MNAd13xnJJXOBi0axNtYZ3fee5UZGo3eAxgdgNvQqalvnQ/iI6/lF0bDi50lG+
wTI/dI+XnKk+hgVm9lL5dCFeKIU3tCOyPZYdxzYWCY64wpfziC2i0omlTTGn728h
7uYfmiq+mqZp5XoVrCs9v397YNJ4QT2sde5dIqECgYEAnOX6tmqbD5wKiSE8WKZB
TysvaUq3awgvqw5r79J3/1HVCiVVjuuoy9ABzhx8zFaNwqZdJVKxZZyl5QWSNwUn
ZIfAADx+DHdP590i9Wk3zr6NwI1cJSSZGHA33lSkVJIE+RUXosg+xCLRBZk5VFN8
xWqRBBSPNnEJrWZCG2p8F0E=
-----END PRIVATE KEY-----
'''




lianwei_key_pub = '''
这里换成联蔚的公钥
'''



def get_public_key_instance_from_str(public_key_str):
    """
    :param public_key_str: 标准格式的RSA公钥，或者光一个key
    :return: 一个能用来加密的公钥instance
    """
    if len(public_key_str) < 20:
        raise ValueError('请填入公钥...')
    if 'BEGIN PUBLIC KEY' in public_key_str and 'END PUBLIC KEY' in public_key_str:
        pass
    else:
        public_key_str = '-----BEGIN PUBLIC KEY-----\n{key}\n-----END PUBLIC KEY-----'.format(key=public_key_str)
    print('使用公钥:\n{pub_key}\n'.format(pub_key=public_key_str))

    public_key_instance = load_pem_public_key(
        bytes(public_key_str, 'utf-8')
    )
    return public_key_instance


def rsa_encrypt(message, public_key_instance):
    """
    :param public_key_instance: 一个公钥instance
    :param message: 一串待加密的信息，长度不超过117字节，len(bytes(message, 'utf-8') <= 177
    :return: base64编码的密文
    """
    message = str(message)
    message_bytes = bytes(message, 'utf-8')
    message_bytes_len = len(bytes(message, 'utf-8'))
    if message_bytes_len <= 177:
        pass
    else:
        raise ValueError('待加密字符串超长! 应小于 177 bytes，得到: {got} bytes'.format(got=message_bytes_len))

    ciphertext_bytes = public_key_instance.encrypt(
        message_bytes,
        padding.OAEP(
            mgf=padding.MGF1(algorithm=hashes.SHA256()),
            algorithm=hashes.SHA256(),
            label=None
        )
    )
    ciphertext_base64 = base64.b64encode(ciphertext_bytes).decode()
    return ciphertext_base64


if __name__ == '__main__':
    # 单个字符串加密测试
    # public_key_ins = get_public_key_instance_from_str(test_key_pub)
    # ciphertext_base64 = rsa_encrypt('17723986404', public_key_instance)
    # print(ciphertext_base64)

    public_key_ins = get_public_key_instance_from_str(test_key_pub)
    # public_key_ins = get_public_key_instance_from_str(lianwei_key_pub)

    print('读取 {current_dir} 文件夹下所有xlsx文件...'.format(current_dir=os.getcwd()))
    for file_name in os.listdir():
        if file_name[-4:] == 'xlsx' and file_name[0] != '~':
            print('读取 {file} ...'.format(file=file_name))
            excel_df = pd.read_excel(io=file_name, sheet_name=0)

            try:
                excel_df['密文手机号'] = excel_df['明文手机号'].apply(rsa_encrypt, args=(public_key_ins,))
                print('文件 {file} ，找到明文手机号，总行数: {len}，为空行数: {empt}.'.format(file=file_name, len=len(excel_df['明文手机号']), empt=sum(excel_df['明文手机号'] == '')))
            except:
                print('文件 {file} 没有「密文手机号」列，跳过...'.format(file=file_name))

            excel_df.to_excel(excel_writer=file_name, sheet_name='Sheet1', header=True, index=False)
            print('带密文xlsx文件覆盖至: {file_path} ...'.format(file_path=os.sep.join([os.getcwd(), file_name])))


