# Streamlit 部署修复技术规范

## 🚨 问题诊断

### 根本原因
1. **缺少 Streamlit 依赖**: `requirements.txt` 未包含 `streamlit` 包
2. **架构不兼容**: 使用 Tkinter 桌面 GUI，无法在 Web 环境运行
3. **部署配置错误**: Streamlit Community Cloud 找不到入口点

### 错误分析
```
sudo: /home/adminuser/venv/bin/streamlit: command not found
ERROR (not running) streamlit: ERROR (spawn error)
```

## 🔧 修复方案

### 1. 更新 requirements.txt
**当前内容:**
```
python-pptx==1.0.2
pandas==2.3.2
openpyxl==3.1.5
pathlib2==2.3.7
```

**修复后内容:**
```
streamlit>=1.28.0
python-pptx==1.0.2
pandas==2.3.2
openpyxl==3.1.5
pathlib2==2.3.7
Pillow>=8.0.0
```

### 2. 架构转换计划

#### Tkinter → Streamlit 组件映射

| Tkinter 组件 | Streamlit 替代 | 实现方式 |
|-------------|---------------|----------|
| `filedialog.askopenfilename()` | `st.file_uploader()` | 文件上传组件 |
| `filedialog.askdirectory()` | `st.text_input()` | 输出文件名输入 |
| `messagebox.showinfo()` | `st.success()` | 成功消息显示 |
| `messagebox.showerror()` | `st.error()` | 错误消息显示 |
| `tk.Tk().withdraw()` | 移除 | Web 应用无需主窗口 |

#### 需要转换的函数

1. **`select_files()` → `streamlit_file_interface()`**
   - 替换文件对话框为上传组件
   - 使用 session state 管理文件状态

2. **配置管理 → Session State**
   - `load_config()` → `st.session_state`
   - `save_config()` → 临时存储机制

3. **用户反馈系统**
   - 进度条: `st.progress()`
   - 状态信息: `st.info()`, `st.warning()`

### 3. 新建 Streamlit 应用文件

**文件名**: `app.py` (Streamlit Community Cloud 标准入口)

**核心功能保留**:
- `read_excel_data()` - Excel 读取逻辑
- `generate_ppt_with_user_files()` - PPT 生成核心
- `extract_zip_and_find_files()` - ZIP 处理
- `find_matching_image()` - 图片匹配
- `handle_circle_markers()` - PPT 标记处理

### 4. Streamlit 应用结构

```python
import streamlit as st
import tempfile
from pathlib import Path

# 页面配置
st.set_page_config(
    page_title="Gemba巡厂PPT生成器",
    page_icon="📊",
    layout="wide"
)

# 主界面
def main():
    st.title("🏭 Gemba巡厂PPT生成器")
    
    # 文件上传区域
    col1, col2 = st.columns(2)
    
    with col1:
        ppt_file = st.file_uploader(
            "上传PPT模板文件", 
            type=['pptx'],
            help="选择PPT模板文件"
        )
    
    with col2:
        zip_file = st.file_uploader(
            "上传数据压缩包", 
            type=['zip'],
            help="包含Excel数据和图片的ZIP文件"
        )
    
    # 输出文件名
    output_filename = st.text_input(
        "输出文件名", 
        value=f"Gemba巡厂报告{datetime.now().strftime('%Y%m%d')}.pptx"
    )
    
    # 生成按钮
    if st.button("🚀 生成PPT", type="primary"):
        if ppt_file and zip_file:
            generate_ppt_streamlit(ppt_file, zip_file, output_filename)
        else:
            st.error("请上传所有必需文件")

def generate_ppt_streamlit(ppt_file, zip_file, output_filename):
    """Streamlit 版本的 PPT 生成函数"""
    with st.spinner("正在生成PPT..."):
        # 使用临时文件处理上传内容
        with tempfile.TemporaryDirectory() as temp_dir:
            # 保存上传文件
            ppt_path = Path(temp_dir) / "template.pptx"
            zip_path = Path(temp_dir) / "data.zip"
            
            with open(ppt_path, "wb") as f:
                f.write(ppt_file.getvalue())
            
            with open(zip_path, "wb") as f:
                f.write(zip_file.getvalue())
            
            # 调用原有生成逻辑
            result = generate_ppt_with_user_files(
                str(ppt_path), 
                str(zip_path), 
                temp_dir
            )
            
            if result:
                # 提供下载
                with open(result, "rb") as f:
                    st.download_button(
                        label="📥 下载生成的PPT",
                        data=f.read(),
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                st.success("✅ PPT生成成功！")
            else:
                st.error("❌ PPT生成失败")

if __name__ == "__main__":
    main()
```

### 5. 部署配置

#### Streamlit Community Cloud 要求
1. **主文件名**: `app.py` (必须)
2. **Python 版本**: ≥ 3.8
3. **依赖文件**: `requirements.txt`
4. **仓库结构**:
   ```
   repository/
   ├── app.py                 # 主应用文件
   ├── requirements.txt       # 依赖列表
   └── [其他支持文件]
   ```

#### 环境变量配置
- 无需特殊环境变量
- 使用 Streamlit 内置 session state

### 6. 测试验证步骤

1. **本地测试**:
   ```bash
   pip install streamlit
   streamlit run app.py
   ```

2. **功能验证**:
   - 文件上传功能
   - PPT 生成逻辑
   - 下载功能
   - 错误处理

3. **部署验证**:
   - 推送到 GitHub
   - 连接 Streamlit Community Cloud
   - 验证在线运行

### 7. 潜在问题和解决方案

#### 内存限制
- **问题**: Streamlit Cloud 内存限制
- **解决**: 优化临时文件处理，及时清理

#### 文件大小限制
- **问题**: 上传文件大小限制
- **解决**: 添加文件大小检查和压缩

#### 处理时间
- **问题**: 长时间处理可能超时
- **解决**: 添加进度显示和分步处理

## 🚀 实施步骤

1. ✅ **架构分析完成**
2. ⏳ **创建 app.py 文件**
3. ⏳ **更新 requirements.txt**
4. ⏳ **测试本地运行**
5. ⏳ **部署到 Streamlit Cloud**
6. ⏳ **验证在线功能**

## 📝 部署检查清单

- [ ] `app.py` 文件存在且可运行
- [ ] `requirements.txt` 包含所有依赖
- [ ] 无 Tkinter 相关代码
- [ ] 文件上传功能正常
- [ ] PPT 生成逻辑无误
- [ ] 错误处理完善
- [ ] 本地测试通过
- [ ] GitHub 仓库已推送
- [ ] Streamlit Cloud 部署成功