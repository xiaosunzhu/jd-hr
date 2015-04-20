# coding=utf-8

__author__ = 'Administrator'

SYSTEM_ENCODING = 'GBK'

CURRENT_VERSION = 'release-0.3.2'


def encode_str(string):
    return string.encode(SYSTEM_ENCODING)


class SelfException(Exception):
    def __init__(self, msg):
        Exception.__init__(self, msg)

