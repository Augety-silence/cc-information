# -*- coding: utf-8 -*-
"""
使用 requests 下载 Excel 中 K~T 列的所有链接文件
文件命名：H列_I列 + 原始后缀
下载到对应 10 个自定义文件夹（文件夹加序号前缀）
自动处理重复文件名
"""

import os
import time
import pandas as pd
import requests
from tqdm import tqdm

# ========== 路径设置 ==========
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "cc_collection.xlsx")
SAVE_DIR = os.path.join(BASE_DIR, "downloads_sorted")

# 自定义文件夹名称（取前五个字）
FOLDER_NAMES = [
    "绿皮载光阴",
    "田埂践农辛",
    "善举润心田",
    "学子辩真知",
    "笺札越流年",
    "异域皆共鸣",
    "灯烛破夜暗",
    "运动强体魄",
    "文艺润生活",
    "星光映童眸"
]

# 创建带序号前缀的目录
os.makedirs(SAVE_DIR, exist_ok=True)
PREFIXED_FOLDERS = []
for idx, name in enumerate(FOLDER_NAMES, start=1):
    folder_name = f"{idx:02d}_{name}"
    PREFIXED_FOLDERS.append(folder_name)
    os.makedirs(os.path.join(SAVE_DIR, folder_name), exist_ok=True)

# ========== 下载函数 ==========
def download_file(url, folder_name, h_val, i_val):
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 13_4) "
                          "AppleWebKit/537.36 (KHTML, like Gecko) "
                          "Chrome/114.0.0.0 Safari/537.36"
        }
        resp = requests.get(url, headers=headers, timeout=30, allow_redirects=True)
        if resp.status_code != 200:
            print(f"下载失败 {url} 状态码: {resp.status_code}")
            return

        # 获取文件后缀
        if 'content-disposition' in resp.headers and '.' in resp.headers['content-disposition']:
            ext = os.path.splitext(resp.headers['content-disposition'].split('filename=')[-1].strip('"'))[1]
        else:
            ext = os.path.splitext(url.split('?')[0])[1]

        if not ext:
            ext = ".dat"  # 默认后缀

        new_name = f"{h_val}_{i_val}{ext}"
        save_path = os.path.join(SAVE_DIR, folder_name, new_name)

        # 防止重复文件名覆盖
        base, ext2 = os.path.splitext(save_path)
        count = 1
        while os.path.exists(save_path):
            save_path = f"{base}({count}){ext2}"
            count += 1

        # 写入二进制文件
        with open(save_path, 'wb') as f:
            f.write(resp.content)
        print(f"已保存到 {save_path}")
        time.sleep(0.5)

    except Exception as e:
        print(f"下载异常 {url}: {e}")


# ========== 主程序 ==========
def main():
    if not os.path.exists(EXCEL_PATH):
        print(f"Excel 文件未找到：{EXCEL_PATH}")
        return

    df = pd.read_excel(EXCEL_PATH, engine="openpyxl")

    for idx, row in tqdm(df.iterrows(), total=len(df), desc="处理进度"):
        h_val = str(row.iloc[7]) if len(row) > 7 else ""
        i_val = str(row.iloc[8]) if len(row) > 8 else ""

        for col_offset, folder_name in enumerate(PREFIXED_FOLDERS):  # K→第1文件夹 ... T→第10文件夹
            col_idx = 10 + col_offset
            if col_idx >= len(row):
                continue
            link = str(row.iloc[col_idx])
            if not (isinstance(link, str) and link.startswith("http")):
                continue

            download_file(link, folder_name, h_val, i_val)

    print("所有文件下载完成")


if __name__ == "__main__":
    main()