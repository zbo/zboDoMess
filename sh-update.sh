date
echo 'start gen category'
python3 gen_category.py
echo 'start real sina batch'
python3 real_sina_batch.py
echo 'start real update'
python3 real_update.py
open /Users/zhubo/code/zboDoMess/content/out.xlsx