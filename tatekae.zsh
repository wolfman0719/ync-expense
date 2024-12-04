# 立て替え費用計算
# tatekae.zsh 年 月
# 例 2024年12月分を作成する 
# irispython tatekae.zsh 2024 12
echo "エクセルファイルコピー" 
cp /Users/hsatoctr/Dropbox/yncorp/"$1"/"$2"/*.xlsx /Users/hsatoctr/work/expense
echo "立て替え費用計算" 
irispython tatekae.py .
echo "立て替え費用をを元のディレクトリにコピーする"
cp /Users/hsatoctr/work/expense/立て替え.xlsx /Users/hsatoctr/Dropbox/yncorp/"$1"/"$2"/
echo "エクセルファイルを削除"
# rm /Users/hsatoctr/work/expense/*.xlsx
