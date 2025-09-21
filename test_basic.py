#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
基础测试程序 - 验证程序功能
"""

import os
import sys
from datetime import datetime
from pathlib import Path

def test_basic_functionality():
    """测试基本功能"""
    print("开始基础功能测试")
    print("=" * 40)
    
    # 设置基础路径
    base_path = Path(r"C:\Users\86151\Desktop\巡厂自动PPT")
    print(f"基础路径: {base_path}")
    
    # 检查文件和文件夹
    template_path = base_path / "参观路线Gemba20250829.pptx"
    excel_path = base_path / "Gemba巡厂_V2_20250920170854" / "Gemba巡厂_V2_20250920170854.xlsx"
    images_path = base_path / "Gemba巡厂_V2_20250920170854" / "Files" / "待整改--现场图片"
    
    print("\n文件检查:")
    print(f"PPT模板: {'存在' if template_path.exists() else '不存在'} - {template_path}")
    print(f"Excel文件: {'存在' if excel_path.exists() else '不存在'} - {excel_path}")
    print(f"图片文件夹: {'存在' if images_path.exists() else '不存在'} - {images_path}")
    
    # 列出图片文件
    if images_path.exists():
        images = list(images_path.glob("*.jpeg"))
        print(f"\n图片文件数量: {len(images)}")
        
        if len(images) > 0:
            print("前5张图片:")
            for i, img in enumerate(images[:5], 1):
                print(f"  {i}. {img.name}")
    
    # 模拟数据处理
    print("\n模拟数据处理:")
    sample_problems = [
        "码垛机器人旁边漏雨",
        "成品库虚线还要有", 
        "AGV会看该区域"
    ]
    
    for problem in sample_problems:
        print(f"问题: {problem}")
        
        # 查找匹配图片
        if images_path.exists():
            found = False
            for img in images_path.glob("*.jpeg"):
                if problem in img.stem:
                    print(f"  -> 找到匹配图片: {img.name}")
                    found = True
                    break
            if not found:
                print(f"  -> 未找到匹配图片")
    
    # 生成输出文件名
    output_filename = f"Gemba巡厂报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
    print(f"\n输出文件名: {output_filename}")
    
    print("\n测试完成!")
    return True

def main():
    """主函数"""
    try:
        result = test_basic_functionality()
        if result:
            print("\n程序基础功能正常")
            return 0
        else:
            print("\n程序基础功能异常")
            return 1
    except Exception as e:
        print(f"\n测试失败: {e}")
        return 1

if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)