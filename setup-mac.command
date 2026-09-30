#!/bin/bash
# 정렬 리모컨 — Mac PowerPoint 설치 스크립트
# 사용법: 이 파일을 우클릭 → 열기 (처음 한 번은 보안 경고가 떠요)

cd "$(dirname "$0")"

echo ""
echo "=== 정렬 리모컨 설치 ==="
echo "GitHub Pages 주소를 붙여넣고 Enter를 누르세요."
echo "예) https://hyejun.github.io/ppt-align-remote"
read -r -p "주소: " URL
URL="${URL%/}"

if [[ ! "$URL" =~ ^https:// ]]; then
  echo "https:// 로 시작하는 주소여야 해요. 다시 실행해 주세요."
  read -r -p "Enter를 누르면 닫힙니다." _
  exit 1
fi

WEF="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
mkdir -p "$WEF"
sed "s#https://YOUR-SITE#${URL}#g" manifest.xml > "$WEF/align-remote-manifest.xml"

echo ""
echo "설치 완료: $WEF/align-remote-manifest.xml"
echo "PowerPoint를 완전히 종료(⌘Q)한 뒤 다시 열고,"
echo "홈 탭 → 추가 기능(Add-ins) → 내 추가 기능에서 '정렬 리모컨'을 선택하세요."
read -r -p "Enter를 누르면 닫힙니다." _
