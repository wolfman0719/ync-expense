# 立て替え費用計算
# tatekae.zsh 年 月
echo 
echo "エクセルファイルコピー" 
cp /Users/hsatoctr/Dropbox/yncorp/"$1"/"$2"/*.xlsx /Users/hsatoctr/work/expense
echo "立て替え費用計算" 
python3 tatekae.py .
echo "立て替え費用をを元のディレクトリにコピーする"
cp /Users/hsatoctr/work/expense/立て替え.xlsx /Users/hsatoctr/Dropbox/yncorp/"$1"/"$2"/
echo "エクセルファイルを削除"
rm /Users/hsatoctr/work/expense/*.xlsx
