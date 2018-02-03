#!/usr/bin/env python
# -*- coding:utf-8 -*-

import ConfigParser


def getConfig(section, key):
    try:
        config = ConfigParser.ConfigParser()
        config.read("default.conf")
        return config.get(section, key)
    except Exception, e:
        return None


if __name__ == '__main__':
    print getConfig('excelFilePath', 'firstFile')
    print getConfig('excelFilePath', 'secondFile')