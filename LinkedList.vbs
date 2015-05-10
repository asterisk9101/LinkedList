class ListItem
    ''' 双方向連結リストの要素を表すクラス
    private value
    private next_
    private prev_
    ''' ListItem は3つの内部パラメータを持つ。
    ''' value は、要素が保持している値を格納する。
    ''' next_ は、次の要素への参照を格納する。
    ''' prev_ は、前の要素への参照を格納する。
    
    private sub Class_Initialize
        ''' ListItem クラスのコンストラクタ
        value = Empty
        set next_ = nothing
        set prev_ = nothing
    end sub
    
    private sub Class_Terminate
        ''' ListItem クラスのデストラクタ
        value = Empty
        set next_ = nothing
        set prev_ = nothing
    end sub
    
    public function init(byval val, byval nx, byval pr)
        ''' ListItem の内部パラメータを一括して設定する。
        ''' setValue, setNext, setPrev メソッドで代替できる。
        ''' 第一引数 val として、ListItem に格納する値 (variant) を受け取る。
        ''' 第二引数 nx として、次の要素への参照 (ListItem or Nothing) を受け取る。
        ''' 第三引数 pr として、前の要素への参照 (ListItem or Nothing) を受け取る。
        ''' 戻り値として、自分自身への参照 (ListItem) を返す。
        call setValue(val)
        call setNext(nx)
        call setPrev(pr)
        set init = me
    end function
    
    public function setValue(byval val)
        ''' ListItem の value パラメータに値をセットする。
        ''' 第一引数 val として、ListItem に格納する値 (variant) を受け取る。
        ''' 戻り値として、自分自身への参照 (ListItem) を返す。
        call bind(value, val)
        set setValue = me
    end function
    
    public function setNext(byval item)
        ''' ListItem の next_ パラメータに、次の要素への参照をセットする。
        ''' 第一引数 item として、次の要素への参照 (ListItem or Nothing) を受け取る。
        ''' 戻り値として、自分自身への参照 (ListItem) を返す。
        if not item is nothing and TypeName(me) <> TypeName(item) then Err.raise(13) ' Type Miss Match
        set next_ = item
        set setNext = me
    end function
    
    public function setPrev(byval item)
        ''' ListItem の prev_ パラメータに、前の要素への参照をセットする。
        ''' 第一引数 item として、前の要素への参照 (ListItem or Nothing) を受け取る。
        ''' 戻り値として、自分自身への参照 (ListItem) を返す。
        if not item is nothing and TypeName(me) <> TypeName(item) then Err.raise() ' Type Miss Match
        set prev_ = item
        set setPrev = me
    end function
    
    public function getValue()
        ''' ListItem の value パラメータの値を返す。
        ''' 戻り値として、value パラメータの値 (variant) を返す。
        call bind(getValue, value)
    end function
    
    public function getNext()
        ''' ListItem の次の要素への参照を返す。
        ''' 戻り値として、次の要素への参照 (ListItem) を返す。
        set getNext = next_
    end function
    
    public function getPrev()
        ''' ListItem の前の要素への参照を返す。
        ''' 戻り値として、前の要素への参照 (ListItem) を返す。
        set getPrev = prev_
    end function
    
    public function clone()
        ''' ListItem を複製する。
        ''' 戻り値として、複製した要素への参照 (ListItem) を返す。
        set clone = (new ListItem).init(value, next_, prev_)
    end function
    
    private function bind(byref out, byval val)
        ''' 第二引数がオブジェクトかその他の値かを判断して第一引数を代入する。
        ''' 第一引数は参照として受け取るため、このユーティリティ関数で代入された結果は呼び出し元の変数に反映される。
        ''' 第一引数 out として、任意の変数への参照を受け取る。
        ''' 第二引数 val として、任意の値を受け取る。
        ''' 戻り値は返さない。
        if isObject(val) then
            set out = val
        else
            out = val
        end if
    end function
end class

class LinkedList
    ''' 双方向連結リストを表すクラス。
    private head
    private tail
    private cursor
    private position
    private count
    ''' クラス変数 head は、LinkedList の先頭の要素への参照 (ListItem) を格納する。
    ''' クラス変数 tail は、LinkedList の末尾の要素への参照 (ListItem) を格納する。
    ''' クラス変数 cursor は、最後にアクセスした要素への参照 (ListItem) を格納する。
    ''' クラス変数 position は、cursor の指す要素の位置 (Number) を格納する。
    ''' クラス変数 count は、LinkedList の要素の数 (Number) を格納する。
    
    private sub Class_Initialize
        ''' LinkedList のコンストラクタ
        set head = nothing
        set tail = nothing
        set cursor = nothing
        position = -1
        count = 0
    end sub
    
    public function init(byval ary)
        call Class_Initialize()
        if not isArray(ary) then ary = array(ary)
        
        dim iter
        for each iter in ary
            call me.push(iter)
        next
        
        set init = me
    end function
    
    public function push(byval val)
        ''' LinkedList の末尾に要素を追加する。
        ''' 戻り値として、リストに含まれる要素の数 (Number) を返す。
        call insert(count, val)
        push = count
    end function
    
    public function unshift(byval val)
        ''' LinkedList の先頭に要素を追加する。
        ''' 戻り値として、リストに含まれる要素の数 (Number) を返す。
        call insert(0, val)
        unshift = count
    end function
    
    public function insert(byval index, byval val)
        ''' 指定した位置に要素を追加する。
        ''' カーソル位置に要素が追加された場合、カーソル位置は後ろへ移動する。
        ''' カーソル位置よりも前に要素が追加された場合、カーソル位置は後ろへ移動する。
        ''' 不正な位置が指定された場合はエラーを発生させる。
        ''' 第一引数 index として、追加する要素の位置 (Number) を受け取る。
        ''' 第二引数 val として、追加する要素 (Variant) を受け取る。
        ''' 戻り値として、リストに含まれる要素の数 (Number) を返す。
        if index > count then err.raise(9)
        if index < 0 then err.raise(9)
        
        dim item, target
        set item = (new ListItem).init(val, nothing, nothing)
        if count = 0 then
            ' 空のリストに要素を追加する場合
            set head = item
            set tail = item
        elseif index = count then
            ' リストの最後に要素を追加する場合
            call item.setPrev(tail)
            call tail.setNext(item)
            set tail = item
        else
            ' リストの間に要素を追加する場合
            set target = getItemPointer(index)
            if target is head then
                call item.setNext(target)
                call target.setPrev(item)
                set head = item
            else
                call item.setPrev(target.getPrev())
                call item.setNext(target)
                call item.getPrev().setNext(item)
                call item.getNext().setPrev(item)
            end if
            if position >= index then
                position = position + 1
            end if
        end if
        
        count = count + 1
        insert = count
    end function
    
    public function pop()
        ''' LinkedList の末尾の要素を取り出す（削除する）。カーソル位置が移動する場合がある。
        ''' 戻り値として、LinkedList の末尾の要素 (Variant) を返す。
        call bind(pop, remove(count - 1))
    end function
    
    public function shift()
        ''' LinkedList の先頭の要素を取り出す（削除する）。カーソル位置が移動する場合がある。
        ''' 戻り値として、LinkedList の先頭の要素 (Variant) を返す。
        call bind(shift, remove(0))
    end function
    
    public function remove(byval index)
        ''' 指定した位置の要素を取り出す（削除する）。
        ''' カーソル位置の要素が取り出された場合、カーソル位置は前へずれる。
        ''' カーソル位置よりも前の要素が取り出された場合、カーソル位置は前へずれる。
        ''' 取り出す位置として不正な値が指定された場合はエラーを発生させる。
        ''' 第一引数 index として、取り出す要素の位置 (Number) を受け取る。
        ''' 戻り値として、取り出した要素 (Variant) を返す。
        if index > count - 1 then err.raise(9)
        if index < 0 then err.raise(9)
        
        dim item
        set item = getItemPointer(index)
        if count = 1 then
            set head = nothing
            set tail = nothing
            set cursor = nothing
            position = -1
        else
            if item is head then
                set head = item.getNext()
                call head.setPrev(nothing)
                position = position - 1
                if item is cursor then
                    set cursor = head
                end if
            elseif item is tail then
                set tail = item.getPrev()
                call tail.setNext(nothing)
                if item is cursor then
                    set cursor = nothing
                    position = -1
                end if
            else
                call item.getPrev().setNext(item.getNext())
                call item.getNext().setPrev(item.getPrev())
                if item is cursor then
                    set cursor = item.getNext()
                elseif index < position then
                    position = position - 1
                end if
            end if
        end if
        
        count = count - 1
        call bind(remove, item.getValue())
    end function
    
    public function item(byval index)
        ''' getItem メソッドのショートカット。
        call bind(item, getItem(index))
    end function
    
    public function getItem(byval index)
        ''' LinkedList の前から index 番目の要素を返す。カーソル (最後にアクセスした要素) の位置が index に移動する。
        ''' 第一引数 index として、要素の位置 (Number) を受け取る。
        ''' 戻り値として、LinkedList の前から index 番目の要素 (Variant) を返す。
        dim item
        set item = getItemPointer(index)
        call bind(getItem, item.getValue())
        set cursor = item
        position = index
    end function
    
    public function setItem(byval index, byval val)
        ''' LinkedList の前から index 番目の要素を書き換える。カーソル (最後にアクセスした要素) の位置が index に移動する。
        ''' 第一引数 index として、要素の位置 (Number) を受け取る。
        ''' 第二引数 val として、書き換える値 (Variant) を受け取る。
        ''' 戻り値を返さない。
        dim item
        set item = getItemPointer(index)
        call item.setValue(val)
        set cursor = item
        position = index
    end function
    
    private function getItemPointer(byval index)
        ''' LinkedList の前から index 番目の要素を返す。カーソルは移動しない。
        ''' index として不正な位置を受け取った場合、エラーを発生させる。
        ''' 第一引数 index として、要素の位置 (Number) を受け取る。
        ''' 戻り値として、LinkedList の前から index 番目の要素 (Variant) を返す。
        if index > count - 1 then err.raise(9)
        if index < 0 then err.raise(9)
        
        dim fromHead, fromTail, fromCursor, distance
        fromHead = index
        fromTail = count - index - 1
        fromCursor = abs(position - index)
        distance = min(array(fromHead, fromTail, fromCursor)) ' 最短距離を計算する
        
        ' 最短距離となる位置から検索開始
        dim item
        select case distance
        case fromHead
            set item = forward(head, fromHead)
        case fromTail
            set item = backward(tail, fromTail)
        case fromCursor
            if fromCursor = 0 then
                set item = cursor
            elseif position - index > 0 then
                set item = backward(cursor, fromCursor)
            else
                set item = forward(cursor, fromCursor)
            end if
        end select
        
        set getItemPointer = item
    end function
    
    private function min(byval ary)
        dim ret, iter
        ret = ary(0) ' 中身なしの場合はエラー
        for each iter in ary
            if ret > iter then ret = iter
        next
        min = ret
    end function
    
    private function forward(byval from, byval distance)
        ''' ある要素から指定された数だけ後ろの要素への参照を返す。
        ''' 第一引数 from として、現在のカーソル (ListItem) を受け取る。
        ''' 第二引数 distance として、移動する距離 (Number) を受け取る。
        ''' 戻り値として、ある要素から指定された数だけ後ろの要素への参照 (ListItem) を返す。
        do while distance <> 0
            set from = from.getNext()
            distance = distance - 1
        loop
        set forward = from
    end function
    
    private function backward(byval from, byval distance)
        ''' ある要素から指定された数だけ前の要素への参照を返す。
        ''' 第一引数 from として、現在のカーソル (ListItem) を受け取る。
        ''' 第二引数 distance として、移動する距離 (Number) を受け取る。
        ''' 戻り値として、ある要素から指定された数だけ前の要素への参照 (ListItem) を返す。
        do while distance <> 0
            set from = from.getPrev()
            distance = distance - 1
        loop
        set backward = from
    end function
    
    public function nextItem()
        if head is nothing then
            set cursor = nothing
            position = -1
            nextItem = empty
        elseif cursor is nothing then
            set cursor = head
            position = 0
            call bind(nextItem, head.getValue())
        elseif cursor.getNext() is nothing then
            set cursor = nothing
            position = -1
            nextItem = empty
        else
            set cursor = cursor.getNext()
            position = position + 1
            call bind(nextItem, cursor.getValue())
        end if
    end function
    
    public function hasNextItem()
        ''' 次の要素がある場合は True を返す。
        ''' 戻り値として、次の要素の有無 (True or False) を返す。
        if head is nothing then
            hasNextItem = false
        elseif cursor is nothing then
            hasNextItem = true
        elseif cursor.getNext() is nothing then
            hasNextItem = false
        else
            hasNextItem = true
        end if
    end function
    
    public function currentItem()
        ''' カーソル (最後にアクセスした要素) の位置にある要素を返す。
        ''' どの要素にもアクセスしたことがない場合はエラーを発生させる。
        ''' 戻り値として、カーソル位置の要素を返す。
        if cursor is nothing then 
            currentItem = empty
        else
            call bind(currentItem, cursor.getValue())
        end if
    end function
    
    public function index()
        ''' カーソル (最後にアクセスした要素) の位置を返す。
        ''' どの要素にもアクセスしたことが無い場合は -1 を返す。
        ''' 戻り値として、カーソルの位置 (Number) を返す。
        index = position
    end function
    
    public function concat(byval list)
        ''' この LinkedList の後ろに別の LinkedList を連結する。破壊的に動作する。
        ''' 第一引数として、別の LinkedList への参照 (LinkedList) を受け取る。
        ''' 戻り値として、自分自身への参照 (LinkedList) を返す。
        if TypeName(list) <> TypeName(me) then Err.raise(13)
        if me is list then set list = me.clone()
        
        dim i
        i = 0
        do while list.hasNextItem()
            me.push(list.getItem(i))
            i = i + 1
        loop
        
        set concat = me
    end function
    
    public function rewind()
        ''' カーソル (最後にアクセスした要素への参照) 位置を初期位置に戻す。
        ''' 戻り値を返さない。
        set cursor = nothing
        position = -1
    end function
    
    public function indexOf(byval key, byval start)
        ''' 指定された位置から後方へ向かって検索し、対象の要素を発見した位置を返す。発見できなかった場合、-1 を返す。
        ''' 検索する値としてオブジェクトを受け取った場合、そのオブジェクトの equals メソッドを使用し最初に true 返す要素を発見した位置を返す。
        ''' 検索を開始する位置として不正な値を受け取った場合、エラーを発生させる。
        ''' 第一引数 key として、検索する値 (variant) を受け取る。
        ''' 第二引数 start として、検索を開始する位置 (Number) を受け取る。
        ''' 戻り値として、LinkedList を後方へ検索して対象の要素を発見した位置 (Number) を返す。
        
        if start > count then Err.raise(9)
        if start < 0 then Err.raise(9)
        
        dim i, result, item
        i = start
        result = -1
        set item = getItemPointer(start)
        if isObject(key) then
            do until item is nothing
                if key.equals(item.getValue()) then
                    result = i
                    exit do
                end if
                set item = item.getNext()
                i = i + 1
            loop
        else
            do until item is nothing
                if key = item.getValue() then
                    result = i
                    exit do
                end if
                set item = item.getNext()
                i = i + 1
            loop
        end if
        indexOf = result
    end function
    
    public function lastIndexOf(byval key, byval start)
        ''' 指定された位置から前方へ向かって検索し、対象の要素を発見した位置を返す。発見できなかった場合、-1 を返す。
        ''' 検索する値としてオブジェクトを受け取った場合、そのオブジェクトの equals メソッドを使用し最初に true 返す要素を発見した位置を返す。
        ''' 検索を開始する位置として不正な値を受け取った場合、エラーを発生させる。
        ''' 第一引数 key として、検索する値 (variant) を受け取る。
        ''' 第二引数 start として、検索を開始する位置 (Number) を受け取る。
        ''' 戻り値として、LinkedList を前方へ検索して対象の要素を発見した位置 (Number) を返す。
        
        if start > count then Err.raise(9)
        if start < 0 then Err.raise(9)
        
        dim i, result, compare, item
        i = start
        result = -1
        
        set item = getItemPointer(start)
        
        if isObject(key) then
            do until item is nothing
                if key.equals(item.getValue()) then
                    result = i
                    exit do
                end if
                set item = item.getPrev()
                i = i - 1
            loop
        else
            do until item is nothing
                if key = item.getValue() then
                    result = i
                    exit do
                end if
                set item = item.getPrev()
                i = i - 1
            loop
        end if
        lastIndexOf = result
    end function
    
    public function clone()
        ''' LinkedList を複製する。カーソルは復元しない。
        ''' 戻り値として、LinkedList の複製 (LinkedList) を返す。
        dim list, i, item
        set list = new LinkedList
        i = 0
        set item = head
        do until item is nothing
            call list.push(item.getValue())
            set item = item.getNext()
            i = i + 1
        loop
        set clone = list
    end function
    
    public function toArray()
        ''' LinkedList の全ての要素を標準の配列に格納して返す。LinkedList が空の場合は null を返す。
        ''' 戻り値として、LinkedList の全ての要素を格納した配列 (Array or null) を返す。
        toArray = null
        if count = 0 then exit function
        
        redim ary(count - 1)
        
        dim item, i
        set item = head
        i = 0
        do until item is nothing
            call bind(ary(i), item.getValue())
            set item = item.getNext()
            i = i + 1
        loop
        
        toArray = ary
    end function
    
    public function toString()
        toString = joinToString(", ")
    end function
    
    public function joinToString(byval sep)
        ''' LinkedList の全ての要素を表す文字列を返す。
        ''' 第一引数 sep として、要素を分割する文字列 (String) を受け取る。
        ''' 戻り値として、LinkedList の全ての要素を表す文字列 (String) を返す。
        jointoString = ""
        dim ary
        ary = toArray()
        if isNull(ary) then exit function
        
        dim i, length
        i = 0
        length = ubound(ary)
        do while i < length
            select case VarType(ary(i))
            case 0 ary(i) = cstr(ary(i)) ' empty
            case 1 ary(i) = cstr(ary(i)) ' null
            case 2 ary(i) = cstr(ary(i)) ' integer
            case 3 ary(i) = cstr(ary(i)) ' long
            case 4 ary(i) = cstr(ary(i)) ' single
            case 5 ary(i) = cstr(ary(i)) ' double
            case 6 ary(i) = cstr(ary(i)) ' currency
            case 7 ary(i) = cstr(ary(i)) ' date
            case 8 ' String
            case 9 ary(i) = "Object(" & TypeName(ary(i)) & ")"
            case 10 ary(i) = cstr(ary(i))
            case 11 ary(i) = cstr(ary(i))
            case 12 ary(i) = cstr(ary(i))
            case 13 ary(i) = "Object(" & TypeName(ary(i)) & ")"
            case 17 ary(i) = cstr(ary(i))
            case else ary(i) = "Array(" & ubound(ary(i))& ")"
            end select
            i = i + 1
        loop
        joinToString = join(ary, sep)
    end function
    
    public function length()
        ''' LinkedList の長さを返す。
        ''' 戻り値として、LinkedList の長さ (Number) を返す。
        length = count
    end function
    
    private function bind(byref out, byval val)
        ''' 第二引数がオブジェクトかその他の値かを判断して第一引数を代入する。
        ''' 第一引数を参照として受け取るため、このユーティリティ関数で代入された結果は呼び出し元の変数に反映される。
        ''' 第一引数 out として、任意の変数への参照を受け取る。
        ''' 第二引数 val として、任意の値を受け取る。
        ''' 戻り値は返さない。
        if isObject(val) then
            set out = val
        else
            out = val
        end if
    end function
end class
