import os

font_path = "C:/Windows/Fonts/ipaexg.ttf"
if os.path.exists(font_path):
    print("フォントファイルが見つかりました:", font_path)
else:
    print("フォントファイルが見つかりません。")
