#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简化测试版本 - 确保基本 Streamlit 功能工作
"""

# 🚨 Critical: Disable ALL GUI backends before any imports
import os
os.environ['MPLBACKEND'] = 'Agg'
os.environ['DISPLAY'] = ''
os.environ['QT_QPA_PLATFORM'] = 'offscreen'
os.environ['SDL_VIDEODRIVER'] = 'dummy'

import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='.*')

import streamlit as st

# 页面配置 - 必须在最开始
st.set_page_config(
    page_title="测试应用",
    page_icon="🧪",
    layout="wide"
)

# 测试内容
st.title("🧪 Streamlit 测试应用")
st.write("如果您看到这个消息，说明 Streamlit 基本功能正常！")

st.header("📝 基本功能测试")
st.write("这是一个简化的测试版本，用来确保 Streamlit 可以正常显示内容。")

# 交互测试
if st.button("点击测试"):
    st.success("✅ 按钮点击功能正常！")
    st.balloons()

# 侧边栏测试
with st.sidebar:
    st.header("侧边栏测试")
    st.write("如果您看到这个侧边栏，说明布局功能正常。")

st.info("💡 如果这个测试页面正常显示，我们就可以确定基础配置正确。")