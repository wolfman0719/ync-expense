# 経費ファイルを日付でソート
# expense.zsh 年 月 日
# 例 1202出張報告精算書.xlsxを編集する 
# sudo python3 expense.zsh 2024 12
echo "エクセルファイルコピー" 
cp /Users/hsatoctr/Dropbox/yncorp/"$1"/"$2"/"$1""$2"経費.xlsx /Users/hsatoctr/work/expense
echo "経費ファイル日付でソート中 $1$2経費.xlsx" 
irispython sortbymd.py /Users/hsatoctr/work/expense/"$1""$2"経費.xlsx
echo "経費ファイルを元のディレクトリにコピーする"
cp /Users/hsatoctr/work/expense/"$1""$2"経費.xlsx /Users/hsatoctr/Dropbox/yncorp/"$1"/"$2"/
echo "エクセルファイルを削除"
rm /Users/hsatoctr/work/expense/"$1""$2"経費.xlsx
