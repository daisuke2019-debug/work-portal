@echo off
chcp 65001 >nul
rem 店舗マスタミスポータルの毎日の更新（10月マスタミス フォルダのExcelを集計して公開する）
cd /d "%~dp0"
echo [1/3] Excelを集計しています...
python scripts\master-miss\update_all.py || goto :error
echo [2/3] ポータルのページを作っています...
python scripts\master-miss\build_portal.py || goto :error
echo [3/3] 公開しています...
git add site/master-miss/index.html
git diff --cached --quiet && (echo 変更はありませんでした。 & goto :done)
git commit -m "Update master-miss data" || goto :error
git push || goto :error
echo 完了しました。1分ほどで https://daisuke2019-debug.github.io/work-portal/master-miss/ に反映されます。
goto :done
:error
echo エラーが発生しました。画面の内容をタクトに見せてください。
:done
pause
