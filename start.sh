#!/bin/bash
# 工作回報系統啟動腳本
cd "$(dirname "$0")"

echo "安裝所需套件..."
pip3 install -r requirements.txt --break-system-packages -q

echo ""
echo "啟動伺服器..."
python3 server.py
