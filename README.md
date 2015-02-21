# LinkedList

双方向連結リストです。ListItem クラスと LinkedList クラスで構成されています。

# 特徴
- メモリの許す限り要素を追加できる(一般的なリストと同様)
- EcmaScript の配列と似たメソッド名を持つ。
- 最後にアクセスした要素を内部のポインタに保存するので、付近の要素へのアクセスが高速になる。
    - n 番目の要素にアクセスした後、n + 1 番目の要素へのアクセスが高速に行える。
    - ある要素へのアクセスする際に、最短のアクセス順序(先頭から or 末尾から or ポインタから)を選択する。
- nextItem メソッドでイテレータ的に走査可能。

# 無いもの
- ラムダ式が前提のメソッド(sort, map, filter etc.)。

# API

```

' 要素の追加と取り出し（ポインタは動かない）

dim list
set list = new LinkedList

call list.push(1) ' 末尾への追加
call list.unshift(true) ' 先頭への追加
call list.insert(1, "a") ' 指定位置への追加

msgbox(list.toString(", ")) ' =>　"True, a, 1"

msgbox list.pop() ' => 1 末尾からの取り出し
msgbox list.shift() ' => True 先頭からの取り出し
msgbox list.remove(0) ' => "a" 指定位置からの取り出し

' ---

call list.push(1)
call list.push(2)
call list.push(3)

' ---




' 要素の参照と書き換え（ポインタが動く）

msgbox list.item(1) ' => 2 ポインタの位置が 1 へ移動する
call list.setElement(1, 10) ' ポインタからアクセスするので高速に動作する
msgbox list.item(1) ' => 10





' イテレータ

dim list2
set list2 = new LinkedList

dim item
item = list.rewind() ' ポインタの位置を最初に戻す。
do until isEmpty(item)
    call list2.unshift(item)
    item = list.nextItem()
loop

msgbox list2.toString(", ") ' => "3, 10, 1"

' toArray, clone, length, indexOf, currentItem, concat etc.

```
