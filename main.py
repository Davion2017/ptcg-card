import os
import shutil
import time
from io import BytesIO

import requests
from PIL import Image
from docx import Document
from docx.shared import Cm


def get_img_urls():
    """
    读取文本中的图片编号
    :return: 返回图片编号生成的url列表
    """
    # 图片下载链接模板
    img_url_template = 'https://asia.pokemon-card.com/hk/card-img/hk%s.png'
    urls = []
    with open('./card.txt', 'r', encoding='utf-8') as file:
        # 读取文件内容
        lines = file.readlines()
        for line in lines:
            if line.startswith('#') or not line.strip():
                continue
            line = line.strip()
            id = line.zfill(8)
            urls.append(f"{img_url_template % id}")
    return urls


def download_img_data(url: str):
    """
    下载图片数据
    :param url: 图片url地址
    :return: 图片的byte数据
    """
    try:
        res = requests.get(url)
        return BytesIO(res.content)
    except:
        print(f'图片下载失败:{url}')
        return None


def insert_img_into_doc(img_path, run):
    """
    写入4张图片到word中
    :param img_path: 图片地址
    :param run: word文档的段落操作符
    :return: None
    """
    run.add_picture(img_path, width=Cm(8.9), height=Cm(6.37))
    run.add_text('  ')
    run.add_picture(img_path, width=Cm(8.9), height=Cm(6.37))
    run.add_break()
    run.add_picture(img_path, width=Cm(8.9), height=Cm(6.37))
    run.add_text('  ')
    run.add_picture(img_path, width=Cm(8.9), height=Cm(6.37))
    run.add_break()


if __name__ == '__main__':
    # 创建存储图片的临时文件夹
    os.makedirs('./tmp', exist_ok=True)
    # 从模板中复制一个新word，用时间命名
    tmp_doc_name = time.strftime('%Y%m%d%H%M%S', time.localtime()) + '.docx'
    shutil.copy('./template.docx', tmp_doc_name)
    # 打开新word，获取当前段落的操作符
    doc = Document(tmp_doc_name)
    paragraph = doc.paragraphs[0]
    run = paragraph.add_run()
    # 获取文本中的图片id对应的下载地址列表
    img_urls = get_img_urls()
    for i in range(len(img_urls)):
        # 下载图片数据
        img_data = download_img_data(img_urls[i])
        if img_data:
            with Image.open(img_data) as img:
                # 图片旋转90度
                rotate_data = img.rotate(90, expand=True)
                save_path = f'./tmp/{img_urls[i][-12:]}'
                rotate_data.save(save_path)
                # 写入到word中
                insert_img_into_doc(save_path, run)
    # 保存word文档
    doc.save(tmp_doc_name)