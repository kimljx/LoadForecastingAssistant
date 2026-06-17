#!/usr/bin/env bash
set -e

# 使用前请确保已经配置 GitHub 凭据，或将 origin URL 改为包含临时 Token 的地址。
REPO_URL="https://github.com/kimljx/LoadForecastingAssistant.git"

git init
git add .
git commit -m "初始化负荷预测与容载比反推助手" || true
git branch -M main
if git remote | grep -q origin; then
  git remote set-url origin "$REPO_URL"
else
  git remote add origin "$REPO_URL"
fi
git push -u origin main
