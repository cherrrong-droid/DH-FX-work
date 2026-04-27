# GitHub 동기화 자동화 스크립트
# 이 파일을 실행하면 현재 폴더의 모든 파일을 GitHub(cherrrong-droid/DH-FX-work)로 업로드합니다.

$repoUrl = "https://github.com/cherrrong-droid/DH-FX-work.git"

Write-Host ">>> GitHub 동기화를 시작합니다..." -ForegroundColor Cyan

# Git 초기화 및 리모트 설정 (이미 되어 있으면 스킵됨)
if (!(Test-Path .git)) {
    git init
    Write-Host "Git 저장소를 초기화했습니다."
}

if (!(git remote show origin)) {
    git remote add origin $repoUrl
    Write-Host "원격 저장소를 연결했습니다: $repoUrl"
} else {
    git remote set-url origin $repoUrl
}

# 파일 추가 및 커밋
git add .
git commit -m "DH FX work project 자동 업데이트 ($(Get-Date -Format 'yyyy-MM-dd HH:mm'))"

# 메인 브랜치로 푸시
git branch -M main
Write-Host "GitHub로 코드를 올리는 중입니다..." -ForegroundColor Yellow
git push -u origin main

Write-Host ">>> 완료되었습니다! 이제 GitHub에서 코드를 확인하실 수 있습니다." -ForegroundColor Green
Read-Host "종료하려면 엔터를 누르세요..."
