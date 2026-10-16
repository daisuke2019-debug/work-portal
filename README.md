# work-portal

仕事用ツールの公開リポジトリ。GitHub Pages で公開されるのは `site/` フォルダだけ。

| フォルダ | 内容 | URL |
|---|---|---|
| `site/srp/` | DCM SRP施策ポータル | https://daisuke2019-debug.github.io/work-portal/srp/ |
| `site/master-miss/` | 店舗マスタミスポータル | https://daisuke2019-debug.github.io/work-portal/master-miss/ |
| `site/annotator/` | 画像注釈・手順書スタジオ | https://daisuke2019-debug.github.io/work-portal/annotator/ |

## ルール
- 公開するファイルは必ず `site/` の中に置く。
- Excel・CSV・集計JSONは `data/` に置く（Gitには入らない）。
- マスタミスポータルの更新: `python scripts/master-miss/update_all.py` → `python scripts/master-miss/build_portal.py`（どのフォルダから実行しても `site/master-miss/index.html` に書き出す）。
