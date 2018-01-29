#!/usr/bin/env python
# -*- coding:utf-8 -*-

import ConfigParser


def getConfig(section, key):
    config = ConfigParser.ConfigParser()
    config.read("default.conf")
    return config.get(section, key)


if __name__ == '__main__':
    print getConfig('excelFilePath', 'firstFile')
    print getConfig('excelFilePath', 'secondFile')