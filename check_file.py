#!/usr/bin/env python3
# -*- coding: utf-8 -*-

def check_file():
    try:
        with open('app.py', 'r', encoding='utf-8') as f:
            lines = f.readlines()
            print("文件前10行:")
            for i, line in enumerate(lines[:10], 1):
                print(f"{i:2d}: {repr(line.rstrip())}")
    except Exception as e:
        print(f"读取文件时出错: {e}")

if __name__ == '__main__':
    check_file() 