# Excelシートにデータを書き込むためのpythonプログラム

## 出張報告精算書作成

### createtravelreport.py

初期バージョン

inputタブに経費.xlsxファイルから該当する出張期間の交通費をコピペして、出張報告精算書を生成

### createtravelreport2.py

経費.xlsxから直接データを読み込んで、日付でソートし、金額が入っていない項目は、データベースから値を取得


### createtravelreport3.py

行数が10行を超える場合、出張報告精算書が10行しかないため、書き込めなかった問題を複数行を1行にする処理を追加

修正は、Claude Codeに全て書かせた

Claude Codeへの指示、および応答は、ClaudeCodeInstructionsResponses.mdを参照

## 経費データを日付順に並べ替える

### sortbymd.py

## 立て替え費用計算

### tatekae.py

個人アカウントで支払った費用および出張の日当を計算する

## expenseitemテーブルの内容をインポート・エクスポート

### expenseitem.py

```
python3 expenseitem.py import file名
python3 expenseitem.py export file名
```
