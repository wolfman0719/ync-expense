# ワイエヌコーポレーション経費レポート作成を支援するツール

## 使用方法

### 経費精算を日付でソートするおよび過去データの入力を補完する

```
sudo python3 sortbymd.py /Users/hsatoctr/work/expense/202410経費.xlsx
```

### 経費精算の内容から出張報告書を作成する

```
sudo python3 createtravelreport.py /Users/hsatoctr/work/expense/1001出張報告精算書.xlsx
```

### 立て替え費用を計算する

```
python3 tatekae.py .
```

### 経費精算日付ソートスクリプト

```
expense.zsh 2025 02
```

### 出張報告書作成用スクリプト

```
travelreport.zsh 2025 03 12
```

### 立て替え費用作成用スクリプト

```
tatekae.zsh 2025 03
```

## 注意事項

### openpyxlの制限事項

openpyxlで生成したファイルは一度エクセルで保存し直さないと、値を取得できない場合がある

### 環境変数

```
export IRISINSTALLDIR=/opt/iris
export LD_LIBRARY_PATH=$IRISINSTALLDIR/bin:$LD_LIBRARY_PATH
# for MacOS
export DYLD_LIBRARY_PATH=$IRISINSTALLDIR/bin:$DYLD_LIBRARY_PATH
# for IRIS username
export IRISUSERNAME=SuperUser
export IRISPASSWORD=SYS
export IRISNAMESPACE=USER
```

### 実行ユーザー

以下のようなエラーが発生する原因は、実行ユーザーがIRISのオーナーと異なるから

MacOSの場合は、sudoでrootになる必要がある(IRISのオーナーがrootなので)


```
  File "/opt/iris/lib/python/iris.py", line 34, in <module>
    from pythonint import *
ImportError: IrisSecureStart failed: IRIS_ATTACH (-21)
```
