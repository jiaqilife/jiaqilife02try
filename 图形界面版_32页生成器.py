#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
图形界面版32页PPT生成器 - Web部署兼容版本
"""

# 🚨 Critical: Disable ALL GUI backends to prevent libtk8.6.so error
import os
import sys
os.environ['MPLBACKEND'] = 'Agg'  # Disable matplotlib GUI backend
os.environ['DISPLAY'] = ''        # Disable X11 display
os.environ['QT_QPA_PLATFORM'] = 'offscreen'  # Disable Qt GUI
os.environ['SDL_VIDEODRIVER'] = 'dummy'      # Disable SDL video

# Disable pandas plotting backends that might trigger tkinter
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='.*')

import json
from datetime import datetime
from pathlib import Path
import re

# 导入库
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE_TYPE
import pandas as pd  # 用于Excel读取

# 配置文件路径
CONFIG_FILE = "gemba_config.json"

def load_config():
    """加载配置文件"""
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception as e:
        print(f"加载配置失败: {e}")
    
    # 默认配置
    return {
        "last_ppt_folder": os.path.expanduser("~/Desktop"),
        "last_zip_folder": os.path.expanduser("~/Desktop"),
        "last_ppt_file": "",
        "last_zip_file": ""
    }

def save_config(config):
    """保存配置文件"""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        print("配置已保存")
    except Exception as e:
        print(f"保存配置失败: {e}")

def select_files_web_compatible(ppt_file=None, zip_file=None, output_folder=None):
    """Web兼容的文件处理函数 - 不使用GUI对话框"""
    print("Web模式：使用参数传递的文件路径")
    
    # 检查必需的文件参数
    if not ppt_file:
        print("错误: 未提供PPT模板文件路径")
        return None, None, None
        
    if not zip_file:
        print("错误: 未提供ZIP压缩包文件路径")
        return None, None, None
        
    if not output_folder:
        output_folder = "."  # 默认使用当前目录
        print("使用默认输出文件夹: 当前目录")
    
    # 验证文件存在性
    if not os.path.exists(ppt_file):
        print(f"错误: PPT模板文件不存在: {ppt_file}")
        return None, None, None
        
    if not os.path.exists(zip_file):
        print(f"错误: ZIP文件不存在: {zip_file}")
        return None, None, None
    
    print(f"文件验证成功:")
    print(f"PPT模板: {os.path.basename(ppt_file)}")
    print(f"ZIP文件: {os.path.basename(zip_file)}")
    print(f"输出位置: {output_folder}")
    
    return ppt_file, zip_file, output_folder

# 保持原函数名但重定向到新的Web兼容版本
def select_files():
    """向后兼容的包装函数"""
    print("警告: 原GUI版本已禁用，请使用Web界面或直接调用generate_ppt_with_user_files")
    return None, None, None

def read_excel_data(excel_path):
    """从Excel文件动态读取数据，替代硬编码数据"""
    try:
        print(f"正在读取Excel文件: {excel_path}")
        
        # 使用pandas读取Excel文件
        df = pd.read_excel(excel_path)
        print(f"Excel文件读取成功，共 {len(df)} 行数据")
        
        # 数据验证：检查必需的列是否存在
        required_columns = ["问题发现区域", "发现人", "问题收集", "问题分类"]
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print(f"警告：Excel文件缺少必需列: {missing_columns}")
            # 使用默认值填充缺失列
            for col in missing_columns:
                df[col] = "未知"
        
        # 过滤空行和无效数据
        df_cleaned = df.dropna(subset=["问题收集"]).copy()
        print(f"清理后有效数据: {len(df_cleaned)} 行")
        
        # 转换为标准格式
        data_list = []
        for _, row in df_cleaned.iterrows():
            data_row = {
                "问题发现区域": str(row.get("问题发现区域", "未知")).strip(),
                "发现人": str(row.get("发现人", "未知")).strip(),
                "问题收集": str(row.get("问题收集", "")).strip(),
                "问题分类": str(row.get("问题分类", "Others")).strip()
            }
            # 只添加非空的问题记录
            if data_row["问题收集"]:
                data_list.append(data_row)
        
        print(f"最终处理数据: {len(data_list)} 行")
        return data_list
        
    except Exception as e:
        print(f"读取Excel文件失败: {e}")
        print("使用备用硬编码数据...")
        # 发生错误时返回原有的硬编码数据作为备用
        return get_all_31_rows_backup()

def get_all_31_rows_backup():
    """备用硬编码数据函数（原get_all_31_rows重命名）"""
    return [
        {"问题发现区域": "包装", "发现人": "谢佳", "问题收集": "码垛机器人旁边漏雨", "问题分类": "5S"},
        {"问题发现区域": "成品库、空柄库", "发现人": "谢佳", "问题收集": "成品库虚线还要有", "问题分类": "5S"},
        {"问题发现区域": "成品库、空柄库", "发现人": "谢佳", "问题收集": "板台里放了箱子，要分开", "问题分类": "5S"},
        {"问题发现区域": "成品库、空柄库", "发现人": "谢佳", "问题收集": "主路不放木箱", "问题分类": "5S"},
        {"问题发现区域": "成品库、空柄库", "发现人": "谢佳", "问题收集": "这个区域少放料", "问题分类": "5S"},
        {"问题发现区域": "公共区域", "发现人": "谢佳", "问题收集": "餐厅区域，信息公布栏，过期信息", "问题分类": "Others"},
        {"问题发现区域": "公共区域", "发现人": "谢佳", "问题收集": "二期餐厅外面空调挂机铁板锈严重", "问题分类": "5S"},
        {"问题发现区域": "装配", "发现人": "谢佳", "问题收集": "立牌子，调试中", "问题分类": "5S"},
        {"问题发现区域": "装配", "发现人": "谢佳", "问题收集": "UV贴纸区域，无关物料不能放在现场", "问题分类": "5S"},
        {"问题发现区域": "钳子冷锻", "发现人": "谢佳", "问题收集": "办公室上面的玻璃要擦", "问题分类": "5S"},
        {"问题发现区域": "钳子冷锻", "发现人": "谢佳", "问题收集": "油，管子，清理 刷漆", "问题分类": "5S"},
        {"问题发现区域": "钳子冷锻", "发现人": "谢佳", "问题收集": "玻璃需要擦", "问题分类": "5S"},
        {"问题发现区域": "钳子冷锻", "发现人": "谢佳", "问题收集": "这个要看 刷漆", "问题分类": "5S"},
        {"问题发现区域": "钳子冷锻", "发现人": "谢佳", "问题收集": "钳子门口  不要放在这个地方", "问题分类": "5S"},
        {"问题发现区域": "活扳", "发现人": "谢佳", "问题收集": "刷完漆搬回去", "问题分类": "5S"},
        {"问题发现区域": "机加工(含刀具)", "发现人": "谢佳", "问题收集": "漏雨点", "问题分类": "5S"},
        {"问题发现区域": "机加工(含刀具)", "发现人": "谢佳", "问题收集": "补漆", "问题分类": "5S"},
        {"问题发现区域": "公共区域", "发现人": "谢佳", "问题收集": "二期门口雨伞架钥匙生锈", "问题分类": "5S"},
        {"问题发现区域": "电镀", "发现人": "谢佳", "问题收集": "漏雨，电镀门口", "问题分类": "5S"},
        {"问题发现区域": "电镀", "发现人": "谢佳", "问题收集": "这里需要包", "问题分类": "5S"},
        {"问题发现区域": "电镀", "发现人": "谢佳", "问题收集": "自动加药区域进展中，下周再来看", "问题分类": "5S"},
        {"问题发现区域": "电镀", "发现人": "谢佳", "问题收集": "AGV会看该区域", "问题分类": "5S"},
        {"问题发现区域": "电镀", "发现人": "谢佳", "问题收集": "电镀漏雨点", "问题分类": "5S"},
        {"问题发现区域": "电镀", "发现人": "谢佳", "问题收集": "下雨，水帘洞", "问题分类": "5S"},
        {"问题发现区域": "电镀", "发现人": "谢佳", "问题收集": "需要换", "问题分类": "5S"},
        {"问题发现区域": "锻造（含下料）", "发现人": "谢佳", "问题收集": "锻造漏雨点", "问题分类": "5S"},
        {"问题发现区域": "锻造（含下料）", "发现人": "谢佳", "问题收集": "重新包一下", "问题分类": "5S"},
        {"问题发现区域": "锻造（含下料）", "发现人": "谢佳", "问题收集": "锻造看板上没有问题显示", "问题分类": "Others"},
        {"问题发现区域": "锻造（含下料）", "发现人": "谢佳", "问题收集": "锻造看板需要更新", "问题分类": "5S"},
        {"问题发现区域": "公共区域", "发现人": "谢佳", "问题收集": "餐厅 框子清理，宣传栏擦，垃圾桶擦", "问题分类": "5S"},
        {"问题发现区域": "成品库、空柄库", "发现人": "谢佳", "问题收集": "成品库板台不能放太多", "问题分类": "5S"}
    ]

def get_category_mapping():
    """获取分类映射"""
    return {
        "Safety": "A",
        "Quality": "B",
        "Efficiency": "C",
        "5S": "D",
        "Cost": "E",
        "Delivery": "F",
        "Others": "G"
    }

def find_matching_image(problem_description, images_path):
    """增强的图片匹配函数"""
    if not problem_description or not images_path.exists():
        return None
    
    # 方法1: 精确匹配
    for img_file in images_path.glob("*.jpeg"):
        if problem_description in img_file.stem:
            return img_file
    
    # 方法2: 反向匹配
    for img_file in images_path.glob("*.jpeg"):
        img_name = img_file.stem
        if img_name in problem_description:
            return img_file
    
    # 方法3: 清理特殊字符后匹配
    problem_clean = problem_description.replace(" ", "").replace("，", "").replace("。", "")
    for img_file in images_path.glob("*.jpeg"):
        img_clean = img_file.stem.replace("_", "").replace("-", "").replace("--", "")
        if problem_clean in img_clean or img_clean in problem_clean:
            return img_file
    
    # 方法4: 关键词匹配
    keywords = problem_description.split()
    for img_file in images_path.glob("*.jpeg"):
        img_name = img_file.stem
        for keyword in keywords:
            if len(keyword) > 1 and keyword in img_name:
                return img_file
    
    return None

def handle_circle_markers(slide, target_category):
    """处理圆形标记 A-G 系统"""
    category_mapping = get_category_mapping()
    target_letter = category_mapping.get(target_category)
    
    if not target_letter:
        print(f"    警告: 未知分类: {target_category}")
        return
    
    print(f"    处理圆形标记: {target_category} -> {target_letter}")
    
    # 查找所有圆形和文本形状
    circles_to_remove = []
    target_circle = None
    
    for shape in slide.shapes:
        if hasattr(shape, "text_frame"):
            text = shape.text_frame.text.strip()
            
            # 检查是否是分类字母标记
            if text in ["A", "B", "C", "D", "E", "F", "G"]:
                if text == target_letter:
                    # 这是目标圆圈，添加勾选标记
                    shape.text_frame.text = "V"
                    target_circle = shape
                    print(f"      [V] 在圆圈 {text} 中添加勾选")
                else:
                    # 这是其他圆圈，标记为删除
                    circles_to_remove.append(shape)
                    print(f"      [X] 标记删除圆圈 {text}")
    
    # 删除未标记的圆圈
    for shape in circles_to_remove:
        try:
            # 删除形状的方法
            sp = shape._element
            sp.getparent().remove(sp)
            print(f"      已删除未使用的圆圈")
        except Exception as e:
            print(f"      删除圆圈失败: {e}")
    
    if target_circle:
        print(f"    [V] 圆圈标记处理完成: {target_category}")
    else:
        print(f"    [!] 未找到目标圆圈: {target_letter}")

def extract_zip_and_find_files(zip_path):
    """解压ZIP文件并查找Excel和图片"""
    import zipfile
    import tempfile
    
    try:
        zip_path = Path(zip_path)
        temp_dir = Path(tempfile.mkdtemp())
        
        # 解压ZIP文件
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        print(f"ZIP文件已解压到: {temp_dir}")
        
        # 查找Excel文件
        excel_files = list(temp_dir.rglob("*.xlsx"))
        if not excel_files:
            raise FileNotFoundError("未在ZIP文件中找到Excel文件")
        
        excel_path = excel_files[0]
        print(f"找到Excel文件: {excel_path}")
        
        # 查找图片文件夹
        image_folders = []
        for folder in temp_dir.rglob("*"):
            if folder.is_dir() and ("图片" in folder.name or "照片" in folder.name):
                image_folders.append(folder)
        
        if not image_folders:
            # 如果没有找到专门的图片文件夹，查找包含jpeg文件的文件夹
            for folder in temp_dir.rglob("*"):
                if folder.is_dir() and list(folder.glob("*.jpeg")):
                    image_folders.append(folder)
        
        if not image_folders:
            raise FileNotFoundError("未在ZIP文件中找到图片文件夹")
        
        images_path = image_folders[0]
        print(f"找到图片文件夹: {images_path}")
        
        return excel_path, images_path, temp_dir
        
    except Exception as e:
        print(f"解压ZIP文件失败: {e}")
        return None, None, None

def generate_ppt_with_user_files(ppt_file, zip_file, output_folder):
    """使用用户选择的文件生成PPT"""
    print("开始生成PPT...")
    
    try:
        # 解压ZIP文件并查找相关文件
        excel_path, images_path, temp_dir = extract_zip_and_find_files(zip_file)
        
        if not excel_path or not images_path:
            print("无法找到Excel文件或图片文件夹")
            return None
        
        # 显示找到的图片数量
        images = list(images_path.glob("*.jpeg"))
        print(f"发现 {len(images)} 张图片")
        
        # 加载PPT模板
        prs = Presentation(ppt_file)
        print(f"加载PPT模板成功，原有 {len(prs.slides)} 张幻灯片")
        
        # 更新第一页日期
        first_slide = prs.slides[0]
        current_date = datetime.now().strftime("%Y%m%d")
        
        for shape in first_slide.shapes:
            if hasattr(shape, "text_frame"):
                text = shape.text_frame.text
                if re.search(r'\d{8}', text):
                    new_text = re.sub(r'\d{8}', current_date, text)
                    shape.text_frame.text = new_text
                    print(f"日期已更新: {text} -> {new_text}")
                    break
        
        # 获取Excel数据（动态行数）
        data = read_excel_data(excel_path)
        print(f"从Excel读取到 {len(data)} 行数据，准备处理")
        
        # 获取模板幻灯片
        if len(prs.slides) < 2:
            print("错误: PPT模板需要至少2张幻灯片")
            return None
            
        template_slide = prs.slides[1]
        
        # 为每行数据创建幻灯片
        created_count = 0
        images_found = 0
        
        for i, row in enumerate(data, 1):
            print(f"\n创建第 {i+1} 页: {row['问题收集'][:30]}...")
            
            try:
                # 添加新幻灯片
                slide_layout = template_slide.slide_layout
                new_slide = prs.slides.add_slide(slide_layout)
                
                # 复制模板内容
                for shape in template_slide.shapes:
                    if hasattr(shape, "text_frame"):
                        # 创建新文本框
                        left = shape.left
                        top = shape.top
                        width = shape.width
                        height = shape.height
                        
                        new_textbox = new_slide.shapes.add_textbox(left, top, width, height)
                        original_text = shape.text_frame.text
                        
                        # 替换占位符内容
                        if original_text == "模具":
                            new_textbox.text_frame.text = row["问题发现区域"]
                            print(f"    占位符1: 模具 -> {row['问题发现区域']}")
                        elif original_text == "-":
                            new_textbox.text_frame.text = row["发现人"]
                            print(f"    占位符2: - -> {row['发现人']}")
                        elif "看板信息更新" in original_text or original_text == "看板信息更新":
                            new_textbox.text_frame.text = row["问题收集"]
                            print(f"    占位符4: {original_text} -> {row['问题收集']}")
                        else:
                            new_textbox.text_frame.text = original_text
                
                # 处理圆形标记系统
                handle_circle_markers(new_slide, row["问题分类"])
                
                # 添加图片到左边 - 完美位置
                image_path = find_matching_image(row["问题收集"], images_path)
                if image_path:
                    try:
                        left = Inches(0.5)  # 左边位置
                        top = Inches(2.1)   # 完美高度
                        width = Inches(3.5)
                        height = Inches(2.8)  # 完美高度
                        new_slide.shapes.add_picture(str(image_path), left, top, width, height)
                        print(f"    [V] 图片已添加: {image_path.name}")
                        images_found += 1
                    except Exception as e:
                        print(f"    [X] 图片添加失败: {e}")
                else:
                    print(f"    [X] 未找到匹配图片")
                
                created_count += 1
                
            except Exception as e:
                print(f"  创建第{i+1}页失败: {e}")
        
        # 删除原始的第二页模板幻灯片
        if len(prs.slides) > 1:
            try:
                # 删除第二张幻灯片（模板页）
                slide_to_remove = prs.slides[1]
                slide_id = slide_to_remove.slide_id
                
                # 从幻灯片列表中移除
                for slide_rel in list(prs.slides._sldIdLst):
                    if slide_rel.id == slide_id:
                        prs.slides._sldIdLst.remove(slide_rel)
                        print("[V] 已删除原始第二页模板幻灯片")
                        break
                        
            except Exception as e:
                print(f"删除模板幻灯片时发生错误: {e}")
        
        # 保存PPT - 使用简化的文件名
        current_date_str = datetime.now().strftime("%Y%m%d")
        output_name = f"Gemba巡厂报告{current_date_str}.pptx"
        output_path = Path(output_folder) / output_name
        
        prs.save(str(output_path))
        
        # 清理临时文件
        import shutil
        if temp_dir.exists():
            shutil.rmtree(temp_dir)
        
        print(f"\n[成功] 动态PPT生成成功!")
        print(f"文件: {output_name}")
        print(f"保存位置: {output_folder}")
        print(f"总页数: {len(prs.slides)} 页")
        print(f"Excel数据行数: {len(data)} 行")
        print(f"[V] 成功创建数据页: {created_count}/{len(data)} 页")
        print(f"成功添加图片: {images_found}/{len(data)} 页")
        print(f"原始模板页: 已删除")
        
        # 显示完成信息（Web兼容版本）
        print("\n=== PPT生成完成 ===")
        print(f"文件名: {output_name}")
        print(f"保存位置: {output_folder}")
        print(f"总页数: {len(prs.slides)} 页")
        print(f"Excel数据: {len(data)} 行")
        print(f"数据页: {created_count} 页")
        print(f"图片: {images_found} 张")
        print("==================")
        
        return str(output_path)
        
    except Exception as e:
        print(f"生成PPT时发生错误: {e}")
        import traceback
        traceback.print_exc()
        
        # 显示错误信息（Web兼容版本）
        print(f"\n=== PPT生成失败 ===")
        print(f"错误信息: {e}")
        print("==================")
        
        return None

def main():
    """主函数 - 带图形界面的PPT生成器"""
    try:
        print("===== Gemba巡厂PPT生成器 (图形界面版) =====")
        print("即将打开文件选择对话框...")
        
        # 选择文件
        ppt_file, zip_file, output_folder = select_files()
        
        if not all([ppt_file, zip_file, output_folder]):
            print("用户取消了文件选择")
            return 1
        
        print(f"\n用户选择:")
        print(f"PPT模板: {ppt_file}")
        print(f"ZIP文件: {zip_file}")
        print(f"输出位置: {output_folder}")
        
        # 生成PPT
        output_file = generate_ppt_with_user_files(ppt_file, zip_file, output_folder)
        
        if output_file:
            print("\n[成功] 程序执行成功!")
            return 0
        else:
            print("\n[失败] 程序执行失败")
            return 1
            
    except Exception as e:
        print(f"\n程序执行失败: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)