# LinkedList

双方向連結リストです。ListItem クラスと LinkedList クラスで構成されています。

# 特徴
- メモリの許す限り要素を追加できる(一般的なリストと同様)
- EcmaScript の配列と似たメソッド名を持つ。
- 最後にアクセスした要素を内部のカーソルに保存するので、付近の要素へのアクセスが高速になる。
    - n 番目の要素にアクセスした後、n + 1 番目の要素へのアクセスが高速に行える。
    - ある要素へのアクセスする際に、最短のアクセス順序(先頭から or 末尾から or カーソルから)を選択する。
- nextItem メソッドでイテレータ的に走査可能。

# 無いもの
- ラムダ式が前提のメソッド(sort, map, filter etc.)。

# API

```

' 要素の追加と取り出し（カーソルは動かない）

dim list
set list = new LinkedList

call list.push(1) ' 末尾への追加 [1]
call list.unshift(true) ' 先頭への追加 [True, 1]
call list.insert(1, "a") ' 指定位置への追加 [True, "a", 1]

msgbox(list.toString(", ")) ' => "True, a, 1"

msgbox list.remove(1) ' => "a" 指定位置からの取り出し
msgbox list.shift() ' => True 先頭からの取り出し
msgbox list.pop() ' => 1 末尾からの取り出し

' ---

call list.push(1)
call list.push(2)
call list.push(3) ' [1, 2, 3]

' ---




' 要素の参照と書き換え（カーソルが動く）

msgbox list.item(1) ' => 2 カーソルの位置が 1 へ移動する
call list.setElement(1, 10) ' カーソルからアクセスするので高速に動作する
msgbox list.item(1) ' => 10

msgbox list.toString(", ") ' => "1, 10, 3"



' イテレータ

dim list2
set list2 = new LinkedList

call list.rewind() ' 最初の要素を取得してカーソルを初期位置に戻す。上記で要素の参照と書き換えを行っているため。
do while list.hasNextItem()
    call list2.unshift(list.nextItem())
loop

msgbox list2.toString(", ") ' => "3, 10, 1"


' toArray, clone, length, indexOf, currentItem, concat etc.

```
