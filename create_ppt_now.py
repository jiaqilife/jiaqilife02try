#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
立即生成PPT - 使用内置库的简化版本
"""

import os
import sys
import shutil
from datetime import datetime
from pathlib import Path

def create_sample_ppt():
    """创建示例PPT文件"""
    print("开始生成PPT文件...")
    
    # 设置路径
    base_path = Path(r"C:\Users\86151\Desktop\巡厂自动PPT")
    template_path = base_path / "参观路线Gemba20250829.pptx"
    
    # 生成输出文件名
    output_filename = f"Gemba巡厂报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
    output_path = base_path / output_filename
    
    # 检查模板文件
    if not template_path.exists():
        print(f"错误: PPT模板文件不存在: {template_path}")
        return None
    
    try:
        # 复制模板文件作为输出文件
        shutil.copy2(str(template_path), str(output_path))
        print(f"成功复制模板文件到: {output_path}")
        
        # 模拟数据处理
        sample_data = [
            {"问题发现区域": "包装", "发现人": "谢佳", "问题收集": "码垛机器人旁边漏雨", "问题分类": "5S"},
            {"问题发现区域": "成品库、空柄库", "发现人": "谢佳", "问题收集": "成品库虚线还要有", "问题分类": "5S"},
            {"问题发现区域": "电镀", "发现人": "谢佳", "问题收集": "AGV会看该区域", "问题分类": "5S"}
        ]
        
        print(f"\n处理数据摘要:")
        print(f"- 数据行数: {len(sample_data)}")
        print(f"- 输出文件: {output_filename}")
        
        # 检查图片匹配
        images_path = base_path / "Gemba巡厂_V2_20250920170854" / "Files" / "待整改--现场图片"
        if images_path.exists():
            matched_count = 0
            print(f"\n图片匹配结果:")
            for row in sample_data:
                problem = row["问题收集"]
                found = False
                for img in images_path.glob("*.jpeg"):
                    if problem in img.stem:
                        print(f"✓ {problem} -> {img.name}")
                        matched_count += 1
                        found = True
                        break
                if not found:
                    print(f"✗ {problem} -> 未找到匹配图片")
            
            print(f"\n匹配统计: {matched_count}/{len(sample_data)} 张图片成功匹配")
        
        return str(output_path)
        
    except Exception as e:
        print(f"生成PPT时发生错误: {e}")
        return None

def create_detailed_instructions():
    """创建详细使用说明"""
    instructions = """
=== Gemba巡厂PPT生成器使用说明 ===

🎯 当前状态:
已成功生成基础PPT文件（基于模板复制）

📋 要获得完整功能，请按以下步骤操作：

1. 安装Python依赖库：
   pip install pandas python-pptx openpyxl

2. 运行完整版程序：
   python gemba_ppt_generator.py

🔧 程序功能：
✓ 自动更新PPT第一页日期
✓ 读取Excel数据（31行巡检记录）  
✓ 为每行数据创建新幻灯片
✓ 填充占位符（问题发现区域、发现人、问题收集）
✓ 根据问题分类打勾（5S、Safety、Quality等）
✓ 自动匹配并插入现场图片（31张可用）

📊 数据文件：
- Excel: Gemba巡厂_V2_20250920170854.xlsx
- 图片: 31张现场照片，支持智能匹配
- 模板: 参观路线Gemba20250829.pptx

💡 提示：
当前生成的文件是模板复制版本。
安装依赖库后可获得完整的自动化功能。
"""
    
    print(instructions)

def main():
    """主函数"""
    try:
        # 生成PPT文件
        output_file = create_sample_ppt()
        
        if output_file:
            print(f"\n✅ PPT文件已生成!")
            print(f"📁 文件位置: {output_file}")
            
            # 显示详细说明
            create_detailed_instructions()
            
            return 0
        else:
            print("\n❌ PPT生成失败")
            return 1
            
    except Exception as e:
        print(f"\n❌ 程序执行失败: {e}")
        return 1

if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)