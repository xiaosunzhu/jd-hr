# coding=utf-8
from _ssl import SSLError
import os
import shutil
import urllib2
import zipfile

from base import encode_str, CURRENT_VERSION


__author__ = 'yijun.sun'

global null
global false
global true

null = None
false = False
true = True


def get_version_code(version_str):
    if '-' in version_str:
        version_nums_str = version_str[version_str.index('-') + 1:].split('.')
    else:
        version_nums_str = version_str.split('.')
    return int(version_nums_str[0]) * 10000 + int(version_nums_str[1]) * 100 + int(version_nums_str[2])


def process_result(content):
    new_version_code = get_version_code(content['tag_name'])
    current_version_code = get_version_code(CURRENT_VERSION)
    if new_version_code <= current_version_code:
        print(encode_str('已是最新版本，不需要更新'))
        return
    if new_version_code - current_version_code >= 100:
        print(encode_str('有重要的新版本：' + content['tag_name']))
    else:
        print(encode_str('有可用的新版本：' + content['tag_name']))
    enter = raw_input(encode_str('回车查看更新信息或输入n回车跳过:'))
    asset = content['assets'][0]
    if enter != 'n':
        print(encode_str('更新信息：\n' + content['body']))
        print(encode_str('软件大小：\n' + str(asset['size'] / 1024) + 'KB'))
    return asset['name'], asset['browser_download_url']


def request_to_github():
    try:
        request = urllib2.Request('https://api.github.com/repos/xiaosunzhu/jd-hr/releases/latest')
        request.add_header('Authorization', 'token 05a5a84f2f3f3099792a51307338cecfae29efdc')
        request.add_header('cache-control', 'no-cache')
        print(encode_str('检查更新中......'))
        response = urllib2.urlopen(request, timeout=7)
        return process_result(eval(response.read()))
    except urllib2.HTTPError, e:
        print(encode_str('检查更新发生错误，Github响应状态：' + str(e.code)))
    except urllib2.URLError, e:
        print(encode_str('检查更新连接服务失败：' + e.reason.message))
    except SSLError, e:
        print(encode_str('检查更新读取失败：' + e.reason.message))


def update(file_name, download_url):
    f = urllib2.urlopen(download_url)
    with open(file_name + '.temp', "wb") as temp_file:
        temp_file.write(f.read())
    zip_temp_file_name = file_name + '.temp'
    zip_file = zipfile.ZipFile(zip_temp_file_name, mode='r')
    dir_name = 'new_version'
    try:
        zip_file.extractall(dir_name)
        for file in zip_file.namelist():
            file_name = dir_name + '/' + file
            if not os.path.isfile(file_name):
                continue
            old_file_name = file[file.index('/') + 1:]
            if old_file_name == 'update.exe':
                continue
            if os.path.exists(old_file_name):
                os.remove(old_file_name)
            shutil.copyfile(file_name, old_file_name)
    finally:
        zip_file.close()
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
        if os.path.exists(zip_temp_file_name):
            os.remove(zip_temp_file_name)