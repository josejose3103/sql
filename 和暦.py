from jeraconv import jeraconv

# J2W クラスのインスタンス生成
# このタイミングで変換の要となる JSON データが load される
j2w = jeraconv.J2W()

# ex.1 : 一般的な使用方法
print(j2w.convert('文治6年'))
# result (int)1190

# ex.2 : 全角数字も使用可能
print(j2w.convert('平成30年'))
# result (int)2019

# ex.3 : "1年" は "元年" と書くことも可能
print(j2w.convert('令和元年'))
# result (int)2019

# ex.4 : 存在しない年号を指定すると ValueError を返す
#print(j2w.convert('牌孫4年'))
# result ValueError
