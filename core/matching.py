# -*- coding: utf-8 -*-
"""
纯逻辑工具（core.matching）

与界面、OCR 无关、可独立测试的函数：编号年份提取、文件夹匹配、路径防碰撞。
均不读取全局配置或文件系统（存在性检查可通过参数注入），便于单元测试。
"""

import os
import re


def extract_year_from_project_id(project_id, prefix_length=4, century_prefix='20'):
    """从项目编号中提取两位年份后缀；无法提取时返回 None。

    规则：跳过编号前 prefix_length 个字符得到数字段；
    若数字段以世纪前缀开头则取其后的两位，否则取数字段前两位。
    """
    if not project_id:
        return None
    number_part = project_id[prefix_length:] if len(project_id) > prefix_length else project_id
    if len(number_part) >= 4 and number_part[:2] == century_prefix:
        return number_part[2:4]
    if len(number_part) >= 2:
        return number_part[:2]
    return None


def folder_matches(folder_name, project_id, prefix_match=True, word_boundary_match=True):
    """判断项目文件夹名是否命中编号（均不区分大小写）。

    - prefix_match：文件夹名以编号开头；
    - word_boundary_match：文件夹名包含完整编号（按非单词字符边界）。
    """
    if prefix_match and folder_name.lower().startswith(project_id.lower()):
        return True
    if word_boundary_match:
        pattern = r'(?:^|\W)' + re.escape(project_id) + r'(?:$|\W)'
        return re.search(pattern, folder_name, re.IGNORECASE) is not None
    return False


def unique_path(path, exists=os.path.exists):
    """对完整路径做防碰撞处理：已存在同名文件时在扩展名前追加 _1、_2…"""
    if not exists(path):
        return path
    base, ext = os.path.splitext(path)
    counter = 1
    while exists(f"{base}_{counter}{ext}"):
        counter += 1
    return f"{base}_{counter}{ext}"


def unique_target_path(folder, filename_base, ext='.pdf', exists=os.path.exists):
    """在 folder 内为 filename_base 生成不冲突的目标路径（冲突时追加 _1、_2…）"""
    target_path = os.path.join(folder, f"{filename_base}{ext}")
    counter = 1
    while exists(target_path):
        target_path = os.path.join(folder, f"{filename_base}_{counter}{ext}")
        counter += 1
    return target_path
