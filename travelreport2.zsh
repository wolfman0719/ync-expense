# 出張報告精算書作成
# ワークディレクトリにコピー
cp /Users/hsatoctr/Dropbox/yncorp/"$1"/"$2"/"$1""$2"経費.xlsx /Users/hsatoctr/work/expense
cp /Users/hsatoctr/Dropbox/yncorp/"$1"/"$2"/"$2""$3"出張報告精算書.xlsx /Users/hsatoctr/work/expense/
# travelreport2.zsh 年 月 日
# 例 1202出張報告精算書.xlsxを編集する 
# python3 travelreport2.zsh 2024 12 02
echo "出張報告精算書作成 $2$3出張報告精算書.xlsx"
irispython createtravelreport2.py /Users/hsatoctr/work/expense/"$1""$2"経費.xlsx /Users/hsatoctr/work/expense/"$2""$3"出張報告精算書.xlsx "$2" "$3"
echo "出張報告精算書を元のディレクトリにコピーする"
cp /Users/hsatoctr/work/expense/"$2""$3"出張報告精算書.xlsx /Users/hsatoctr/Dropbox/yncorp/"$1"/"$2"/
echo "エクセルファイルを削除"
rm /Users/hsatoctr/work/expense/"$2""$3"出張報告精算書.xlsx
rm /Users/hsatoctr/work/expense/"$1""$2"経費.xlsx