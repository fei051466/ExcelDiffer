#!/usr/bin/env python
# -*- coding:utf-8 -*-


def diff(dataBefore, dataAfter):
    """
    计算数据修改前后的变化
    先考虑最简单的情况：只有行变化
    实现一个只能计算行增删情况的diff算法
    :param dataBefore: 二维数组数据
    :param dataAfter: 二维数组数据
    :return: 变化描述，包含行增删，列增删和单元增删信息
    """
    dataBeforeT = transfer(dataBefore)
    dataAfterT = transfer(dataAfter)
    indexBefore, indexAfter = longgestCommonSubsequence(dataBeforeT, dataAfterT)
    print indexBefore, indexAfter
    info = []
    i = 0
    j = 0
    bLen = len(dataBeforeT)
    aLen = len(dataAfterT)
    bIndex = 0
    aIndex = 0
    bIndexLen = len(indexBefore)
    aIndexLen = len(indexAfter)
    while i < bLen or j < aLen:
        if bIndex == bIndexLen:
            info.append("a%d" % (j + 1))
            j += 1
        elif aIndex == aIndexLen:
            info.append("d%d" % (i + 1))
            i += 1
        elif indexBefore[bIndex] == i and indexAfter[aIndex] == j:
            info.append("s%d" % (j + 1))
            i += 1
            j += 1
            bIndex += 1
            aIndex += 1
        elif indexBefore[bIndex] == i and indexAfter[aIndex] > j:
            info.append("a%d" % (j + 1))
            j += 1
        elif indexBefore[aIndex] > i and indexAfter[aIndex] == j:
            info.append("d%d" % (i + 1))
            i += 1
        else:
            info.append("c%d" % j)
            i += 1
            j += 1
    return info


def transfer(data):
    """
    将二维数组转化为一维数组
    直接将每一行的元数据简单合并为一个字符串
    :param data: 二维数组数据
    :return: 一维数组数据
    """
    return ["".join(map(str, row)) for row in data]


def longgestCommonSubsequence(a, b):
    aLen = len(a)
    bLen = len(b)
    dp = [[0 for i in range(bLen + 1)] for j in range(aLen + 1)]
    path = [[0 for i in range(bLen + 1)] for j in range(aLen + 1)]
    for i in range(aLen):
        for j in range(bLen):
            if equal(a[i], b[j]):
                dp[i + 1][j + 1] = dp[i][j] + 1
                path[i + 1][j + 1] = 1
            elif dp[i + 1][j] > dp[i][j + 1]:
                dp[i + 1][j + 1] = dp[i + 1][j]
                path[i + 1][j + 1] = -1
            else:
                dp[i + 1][j + 1] = dp[i][j + 1]
                path[i + 1][j + 1] = -2
    return calcSubsequenceIndex(path, a, b, aLen, bLen)


def equal(a, b):
    """
    计算两个数据是否相等
    使用==判断
    :param a: 数据a
    :param b: 数据b
    :return: 数据a与b是否相等
    """
    return a == b


def calcSubsequenceIndex(path, a, b, aIndex, bIndex):
    if aIndex == 0 or bIndex == 0:
        return [], []
    if path[aIndex][bIndex] == 1:
        s1, s2 = calcSubsequenceIndex(path, a, b, aIndex - 1, bIndex - 1)
        s1.append(aIndex - 1)
        s2.append(bIndex - 1)
        return s1, s2
    elif path[aIndex][bIndex] == -1:
        return calcSubsequenceIndex(path, a, b, aIndex, bIndex - 1)
    else:
        return calcSubsequenceIndex(path, a, b, aIndex - 1, bIndex)


if __name__ == '__main__':
    data1 = [[1,2,3], [2,3,4]]
    print transfer(data1)
    data2 = [[2,3,4]]
    diff(data1, data2)

