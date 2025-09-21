# Gemba巡厂PPT自动生成器

## 项目概述

这是一个自动化工具，用于处理Gemba巡厂检查数据并生成PPT报告。程序能够读取Excel数据文件，匹配相应的现场图片，并自动生成包含所有检查结果的PowerPoint演示文稿。

## 功能特性

### ✅ 已实现功能

1. **日期自动更新**
   - 自动将PPT第一页的日期更新为当前系统日期
   - 支持YYYYMMDD格式的日期识别和替换

2. **Excel数据处理**
   - 读取指定路径的Excel文件（Gemba巡厂_V2_20250920170854.xlsx）
   - 支持以下列数据：
     - 问题发现区域
     - 发现人
     - 问题收集
     - 问题分类

3. **PPT幻灯片生成**
   - 为每行Excel数据创建新的PPT幻灯片
   - 使用PPT模板保持格式一致性
   - 自动填充占位符内容

4. **占位符数据填充**
   - 占位符1：问题发现区域
   - 占位符2：发现人
   - 占位符4：问题收集内容

5. **问题分类标记**
   - 根据"问题分类"字段在对应选项上添加√符号
   - 支持的分类：Safety、Efficiency、Cost、Quality、Delivery、5S、Others

6. **智能图片匹配**
   - 基于"问题收集"内容自动匹配对应的现场图片
   - 支持精确匹配和部分匹配
   - 自动插入匹配的图片到幻灯片中

7. **异常处理机制**
   - 文件路径存在性验证
   - 支持常见图片格式（JPEG）
   - 详细的错误日志记录
   - 程序运行状态监控

8. **用户体验优化**
   - 详细的处理进度提示
   - 清晰的日志输出
   - 自动生成带时间戳的输出文件名

## 文件结构

```
巡厂自动PPT/
├── gemba_ppt_generator.py      # 完整功能主程序
├── simple_ppt_generator.py     # 简化版演示程序
├── test_basic.py              # 基础功能测试程序
├── requirements.txt           # Python依赖库列表
├── README.md                  # 项目说明文档
├── 参观路线Gemba20250829.pptx  # PPT模板文件
└── Gemba巡厂_V2_20250920170854/
    ├── Gemba巡厂_V2_20250920170854.xlsx  # Excel数据文件
    └── Files/
        └── 待整改--现场图片/              # 现场图片文件夹
            ├── AGV会看该区域.jpeg
            ├── UV贴纸区域，无关物料不能放在现场.jpeg
            └── ... (共31张图片)
```

## 安装依赖

### 方式1：使用requirements.txt（推荐）
```bash
cd 巡厂自动PPT
pip install -r requirements.txt
```

### 方式2：手动安装
```bash
pip install pandas python-pptx openpyxl
```

## 使用方法

### 1. 完整功能版本
```bash
cd 巡厂自动PPT
python gemba_ppt_generator.py
```

### 2. 简化演示版本（无需外部依赖）
```bash
cd 巡厂自动PPT
python simple_ppt_generator.py
```

### 3. 基础功能测试
```bash
cd 巡厂自动PPT
python test_basic.py
```

## 测试结果

✅ **测试状态：通过**

- 文件路径验证：✅ 通过
- PPT模板文件：✅ 存在
- Excel数据文件：✅ 存在  
- 图片文件夹：✅ 存在（31张图片）
- 图片匹配功能：✅ 正常工作
- 数据处理逻辑：✅ 正常运行

### 测试示例结果
```
问题: 码垛机器人旁边漏雨
  -> 找到匹配图片: 码垛机器人旁边漏雨.jpeg

问题: 成品库虚线还要有
  -> 找到匹配图片: 成品库虚线还要有.jpeg

问题: AGV会看该区域
  -> 找到匹配图片: AGV会看该区域.jpeg
```

## 输出文件

程序运行后会在项目目录下生成：
- `Gemba巡厂报告_YYYYMMDD_HHMMSS.pptx` - 生成的PPT报告
- `gemba_ppt_generator.log` - 详细运行日志

## 技术实现

### 核心技术栈
- **Python 3.x** - 主要编程语言
- **python-pptx** - PowerPoint文件操作
- **pandas** - Excel数据处理
- **openpyxl** - Excel文件读写支持
- **pathlib** - 文件路径处理

### 关键算法
1. **日期识别和替换**：使用正则表达式识别8位日期格式
2. **图片智能匹配**：基于文件名和问题描述的双向匹配算法
3. **模板复制机制**：动态复制PPT模板幻灯片并填充数据

## 注意事项

1. **文件路径**：确保所有输入文件都在指定位置
2. **图片格式**：目前仅支持JPEG格式图片
3. **Excel格式**：需要包含指定的列名
4. **PPT模板**：需要包含指定的占位符

## 故障排除

### 常见问题

1. **ModuleNotFoundError**
   - 解决：运行 `pip install -r requirements.txt`

2. **文件路径错误**
   - 解决：检查文件是否在正确位置

3. **编码问题**
   - 解决：使用 `test_basic.py` 进行基础测试

## 版本历史

- **v1.0** - 基础功能实现
  - PPT模板处理
  - Excel数据读取
  - 图片匹配和插入
  - 异常处理机制

## 开发者信息

本程序由jiaqilife开发，专门用于自动化Gemba巡厂检查报告生成。