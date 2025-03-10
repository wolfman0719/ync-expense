# ワイエヌコーポレーション経費レポート作成を支援するツール

## 使用方法

### 経費精算を日付でソートするおよび過去データの入力を補完する

```
irispython sortbymd.py /Users/hsatoctr/work/expense/202410経費.xlsx
```

### 経費精算の内容から出張報告書を作成する

```
irispython createtravelreport.py /Users/hsatoctr/work/expense/1001出張報告精算書.xlsx
```

### 立て替え費用を計算する

```
irispython tatekae.py .
```

### 出張報告書作成用スクリプト

```
expense.zsh 2025 02
```

### 旅費精算書作成用スクリプト

```
travelreport.zsh 2025 03 12
```

### 立て替え費用作成用スクリプト

```
tatekae.zsh 2025 03
```

## 注意事項

openpyxlで生成したファイルは一度エクセルで保存し直さないと、値を取得できない場合がある
