# 出張報告精算書作成
# ワークディレクトリにコピー
# expense.zsh 年 月 日
# 例 1202出張報告精算書.xlsxを編集する 
# irispython expense.zsh 2024 12 02
echo "エクセルファイルコピー" 
cp /Users/hsatoctr/Dropbox/yncorp/"$1"/"$2"/"$2""$3"出張報告精算書.xlsx /Users/hsatoctr/work/expense
echo "経費ファイル日付でソート中 $1$2経費.xlsx" 
irispython sortbymd.py /Users/hsatoctr/work/expense/"$1""$2"経費.xlsx
echo "出張報告精算書作成 $2$3出張報告精算書.xlsx" 
irispython createtravelreport.py /Users/hsatoctr/work/expense/"$2""$3"出張報告精算書.xlsx
echo "出張報告精算書を元のディレクトリにコピーする"
cp /Users/hsatoctr/work/expense/*.xlsx /Users/hsatoctr/Dropbox/yncorp/"$1"/"$2"/
echo "エクセルファイルを削除"
# rm /Users/hsatoctr/work/expense/*.xlsx
