#!/bin/bash

echo "安装依赖..."
pip install -r requirements.txt

echo "开始打包Mac可执行程序..."
pyinstaller --noconfirm excel-app.spec

echo "打包完成，输出文件位于 dist 目录" 