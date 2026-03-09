# 出張報告精算書作成
# ワークディレクトリにコピー
export IRISINSTALLDIR=/opt/iris
export LD_LIBRARY_PATH=$IRISINSTALLDIR/bin:$LD_LIBRARY_PATH
export DYLD_LIBRARY_PATH=$IRISINSTALLDIR/bin:$DYLD_LIBRARY_PATH
export IRISUSERNAME=_system
export IRISPASSWORD=SYS
export IRISNAMESPACE=USER
cp /Users/hsatoctr/Dropbox/yncorp/"$1"/"$2"/"$1""$2"経費.xlsx /Users/hsatoctr/work/expense
cp /Users/hsatoctr/Dropbox/yncorp/"$1"/"$2"/"$2""$3"出張報告精算書.xlsx /Users/hsatoctr/work/expense/
# travelreport3.zsh 年 月 日
# 例 1202出張報告精算書.xlsxを編集する 
# python3 travelreport3.zsh 2024 12 02
echo "出張報告精算書作成 $2$3出張報告精算書.xlsx"
# irispython createtravelreport3.py /Users/hsatoctr/work/expense/"$1""$2"経費.xlsx /Users/hsatoctr/work/expense/"$2""$3"出張報告精算書.xlsx "$2" "$3"
python3 createtravelreport3.py /Users/hsatoctr/work/expense/"$1""$2"経費.xlsx /Users/hsatoctr/work/expense/"$2""$3"出張報告精算書.xlsx "$2" "$3"
echo "出張報告精算書を元のディレクトリにコピーする"
cp /Users/hsatoctr/work/expense/"$2""$3"出張報告精算書.xlsx /Users/hsatoctr/Dropbox/yncorp/"$1"/"$2"/
echo "エクセルファイルを削除"
rm /Users/hsatoctr/work/expense/"$2""$3"出張報告精算書.xlsx
rm /Users/hsatoctr/work/expense/"$1""$2"経費.xlsx