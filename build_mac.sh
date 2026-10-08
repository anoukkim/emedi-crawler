#!/bin/bash
# emedi 수집기 - macOS 앱 만들기:  터미널에서  bash build_mac.sh
set -e
cd "$(dirname "$0")"
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt pyinstaller
pyinstaller --noconfirm --clean --windowed --name "emedi_crawler" \
  --collect-data selenium \
  --exclude-module tkinter --exclude-module matplotlib \
  app.py
echo
echo "완료: dist/emedi_crawler.app  →  응용 프로그램(Applications) 폴더로 드래그하세요."
echo "처음 열 때 '확인되지 않은 개발자' 경고가 나오면: 앱을 우클릭 → 열기 → 열기"
