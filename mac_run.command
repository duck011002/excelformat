#!/bin/bash
# 解决双击时当前工作目录变成家目录的问题
cd "$(dirname "$0")"

echo "=========================================="
echo "    成绩 Excel 智能清洗工具 - Mac 启动器"
echo "=========================================="

# 1. 检查 Python3 环境
if ! command -v python3 &> /dev/null; then
    echo "❌ 错误: 未检测到 Python 3 环境！"
    echo "💡 提示: 请先访问 https://www.python.org/ 安装适用于 Mac 的 Python 3。"
    echo "安装完成后，重新双击此脚本即可。"
    echo ""
    echo "按任意键退出..."
    read -n 1
    exit 1
fi

# 2. 检查并安装依赖
echo "⏳ 正在检查并自动安装/升级依赖库，请稍候..."
python3 -m pip install -U pip
python3 -m pip install -r requirements.txt

# 3. 运行程序
echo "🚀 正在启动服务..."
python3 grade_excel_cleaner_v1/windows_launcher.py
