# coding=utf-8

import sys
from time import sleep
import traceback

from base import encode_str, SelfException
from check_update import request_to_github, update


__author__ = 'yijun.sun'

reload(sys)
sys.setdefaultencoding("utf-8")

try:
    result = request_to_github()
    if result:
        print('')
        enter = raw_input(encode_str('回车进行更新或输入n回车退出:'))
        if enter != 'n':
            print(encode_str('正在更新，请耐心等待......'))
            update(result[0], result[1])
except Exception, e:
    print(encode_str('程序异常！ ') + str(e.message))
    if not isinstance(e, SelfException):
        sleep(0.2)
        print('')
        traceback.print_exc()
finally:
    sleep(0.6)
    raw_input(encode_str('键入回车退出程序'))