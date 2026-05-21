#!/bin/bash
# 双击本文件即可启动「报告智能体」网页工具
cd "$(dirname "$0")"
source venv/bin/activate
echo "正在启动报告工具，浏览器将自动打开……（关闭本窗口即可停止）"
streamlit run app.py
