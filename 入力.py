# -*- coding: utf8 -*-

import sys
import tkinter
import tkinter.ttk as ttk
import sqlite3
import datetime

root = tkinter.Tk()
root.title("牛追加")
root.geometry("260x100")

#新規ウィンドウ登録.
def ViewNewWindow(self):
    newWindow = tkinter.Toplevel(root)
    newWindow.title("結果表示")
    newWindow.geometry("260x350")
    Make_ResultMap(self,newWindow)

#DBからの読み出し結果をTreeviewで表示
def Make_ResultMap(self,newWindow):
    #ウィンドウにTreeviewを登録
    tree = ttk.Treeview(newWindow)

    #列数を定義
    tree["columns"] = (1,2,3)
    tree["show"] = "headings"

    #各列ごとの幅(大きさ)設定
    tree.column(1,width=75)
    tree.column(2,width=100)
    tree.column(3,width=75)

    #表示する列名定義
    tree.heading(1,text="生年月日")
    tree.heading(2,text="農家名")
    tree.heading(3,text="個体ＩＤ")
    
    #DBから登録データ読み出し.
    Getlist = Get_sql()
    
    #treeviewに追加.
    for buf in Getlist:
        tree.insert("","end",values=buf)

    tree.pack()

#テキストボックスからの入力データ読み出し
def GetValues(self):
    value_item = Input_item.get()
    value_money = Input_money.get()

    #DBへのデータ登録.
    Insert_sql(value_item,value_money)

#DBからデータ取得.
def Get_sql():
    try:
        conn = sqlite3.connect('カルテ.mdb')
        c = conn.cursor()
        list_buf = []
        for buf in c.execute("select * from テーブル４"):
            list_buf.append(buf)
        conn.close()
    except sqlite3.Error as e:
        print(e)
    
    print(list_buf)
    return list_buf

#DBへ入力データ挿入.
def Insert_sql(item,money):
    now = datetime.datetime.now()
    
    try:
        conn = sqlite3.connect('database.db')
        c = conn.cursor()
        c.execute("insert into kakeibo values (?,?,?)",(now.strftime('%Y/%m/%d'),str(item),money))
        conn.commit()
        conn.close()
    except sqlite3.Error as e:
        print(e)

#main画面生成.
item = tkinter.Label(text=u'物品')
Input_item = tkinter.Entry(width=30)
Input_item.insert(tkinter.END,"文字列を入力してください")

money = tkinter.Label(text=u'金額')
Input_money = tkinter.Entry(width=30)
Input_money .insert(tkinter.END,"金額を入力してください")

#テキストボックス生成.
item.grid(row=0)
Input_item.grid(row=0,column=1)
money.grid(row=1)
Input_money.grid(row=1,column=1)

#ボタン生成
Button = tkinter.Button(root,text=u'ボタン')
Button.bind("<Button-1>",GetValues) #ボタンイベント登録.
Button.grid(row=2)

#ボタン生成
Button = tkinter.Button(root,text=u'結果表示')
Button.bind("<Button-1>",ViewNewWindow) #ボタンイベント登録.
Button.grid(row=3)

root.mainloop()
