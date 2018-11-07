# appid 已在配置中移除,请在参数 Bucket 中带上 appid。Bucket 由 bucketname-appid 组成
# 1. 设置用户配置, 包括 secretId，secretKey 以及 Region
# -*- coding=utf-8

# bucket = 	aic-1253948304

from qcloud_cos import CosConfig
from qcloud_cos import CosS3Client
import os
import re
import sys
import logging

logging.basicConfig(level=logging.INFO, stream=sys.stdout)

secret_id = 'AKIDWt55sQDlZ9VaPyrL7csra3GF2ZqyMWv6'      # 替换为用户的 secretId
secret_key = '4uWt1uYw4wXsxlJabMr53kBLxJPFp9D2'      # 替换为用户的 secretKey
region = 'ap-guangzhou'     # 替换为用户的 Region
token = None                # 使用临时密钥需要传入 Token，默认为空，可不填
scheme = 'https'            # 指定使用 http/https 协议来访问 COS，默认为 https，可不填
config = CosConfig(Region=region, SecretId=secret_id, SecretKey=secret_key, Token=token, Scheme=scheme)
# 2. 获取客户端对象
client = CosS3Client(config)

list_folder_response = client.list_objects(
    Bucket='aic-1253948304',
    # Delimiter='/',
    # Marker='string',
    MaxKeys=1000,
    Prefix='CorpImg/SL/8至11月516户个体企业抽查',
    # EncodingType='url'
)

for item in list_folder_response['Contents']:
    download_file_dict = client.get_object(
        Bucket='aic-1253948304',
        Key=item['Key']
    )
    e = re.match('CorpImg\/SL\/(\S+)\/(\S+)\/(\S+)_(\S+)_(\S+)-(\S+)--(\S+)\.jpg', item['Key'])
    # 将KEY分拆, e = elements
    # Example: CorpImg / SL / 8至11月516户个体企业抽查 / 20180905 / SL_广州市太雅皮具有限公司_20180905 - 1008 - -26394.jpg
    # 1 - 8至11月516户个体企业抽查
    # 2 - 20180905
    # 3 - SL
    # 4 - 广州市太雅皮具有限公司
    # 5 - 20180905
    # 6 - 1008
    # 7 - 26394

    # 建立行动名文件夹并进入
    try:
        os.mkdir(e[1])
    except FileExistsError:
        pass
    finally:
        os.chdir(e[1])

    # 建立企业名文件夹并进入
    try:
        os.mkdir(e[4])
    except FileExistsError:
        pass
    finally:
        os.chdir(e[4])

    # 将文件的BODY存入格式化的文件名中
    download_file_dict['Body'].get_stream_to_file(e[4] + '-' + e[5] + e[6] + '-' + e[7] + '.jpg')
    # 回到行动名文件夹
    os.chdir('..')
    # 回到程序所在文件夹
    os.chdir('..')
    # 本循环结束
