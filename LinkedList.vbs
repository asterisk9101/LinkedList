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
    private p
    private count
    ''' LinkedList クラスは 4 つの内部パラメータを持つ。
    ''' head は、LinkedList の先頭の要素への参照 (ListItem) を格納する。
    ''' tail は、LinkedList の末尾の要素への参照 (ListItem) を格納する。
    ''' p は、最後にアクセスされた位置の情報を格納した要素への参照 (ListItem) を格納する。
    '''     p.getValue() は、最後にアクセスした要素の位置 (Number) を返す。
    '''     p.getPrev() は、最後にアクセスした要素への参照 (ListItem) を返す。
    '''     p.getNext() は、最後にアクセスした要素の次の要素への参照 (ListItem) を返す。
    ''' count は、LinkedList の要素の数 (Number) を格納する。
    
    private sub Class_Initialize
        ''' LinkedList のコンストラクタ
        set head = nothing
        set tail = nothing
        set p = new ListItem
        count = 0
    end sub
    
    private sub Class_Terminate
        ''' LinkedList のデストラクタ
        set head = nothing
        set tail = nothing
        set p = nothing
        count = 0
    end sub
    
    public function push(byval val)
        ''' LinkedList の末尾に要素を追加する。ポインタ位置が移動する場合がある。
        ''' 戻り値は返さない。
        call insert(count, val)
    end function
    
    public function unshift(byval val)
        ''' LinkedList の先頭に要素を追加する。ポインタ位置が移動する場合がある。
        ''' 戻り値は返さない。
        call insert(0, val)
    end function
    
    public function insert(byval index, byval val)
        ''' 指定した位置に要素を追加する。
        ''' ポインタ位置に要素が追加された場合、ポインタ位置は後ろへ移動する。
        ''' ポインタ位置よりも前に要素が追加された場合、ポインタ位置は後ろへ移動する。
        ''' 不正な位置が指定された場合はエラーを発生させる。
        ''' 第一引数 index として、追加する要素の位置 (Number) を受け取る。
        ''' 第二引数 val として、追加する要素 (Variant) を受け取る。
        ''' 戻り値は返さない。
        if index > count then err.raise(9)
        if index < 0 then err.raise(9)
        
        dim item
        set item = new ListItem
        call item.setValue(val)
        
        if head is nothing then
            call setToHeadIsNothing(item)
        elseif head is tail then
            call setToHeadIsTail(index, item)
        else
            call setToListedItems(index, item)
        end if
        
        count = count + 1
    end function
    
    private function setToHeadIsNothing(byval item)
        ''' 空の LinkedList に要素を追加する。
        ''' 第一引数 val として、追加する要素 (Variant) を受け取る。
        ''' 戻り値は返さない。
        set head = item
        set tail = item
        call p.init(-1, item, nothing)
    end function
    
    private function setToHeadIsTail(byval index, byval item)
        ''' 要素数 1 の LinkedList に要素を追加する。
        ''' ポインタ位置に要素が追加された場合、ポインタ位置は後ろへずれる。
        ''' ポインタ位置よりも前に要素が追加された場合、ポインタ位置は後ろへずれる。
        ''' 第一引数 index として、追加する要素の位置 (0 or 1) を受け取る。
        ''' 第二引数 val として、追加する要素 (Variant) を受け取る。
        ''' 戻り値は返さない。
        if index = 0 then
            ' list の先頭に要素をインサートする
            
            call head.setPrev(item)
            call item.setNext(head)
            set head = item
            
            if p.getPrev() is nothing then
                ' ポインタが進んでいない場合
                call p.init(-1, item, nothing)
            else
                ' ポインタが進んでいる場合
                call p.init(2, nothing, tail)
            end if
        else
            ' list の末尾に要素をインサートする
            call head.setNext(item)
            call item.setPrev(item)
            set tail = item
            
            if p.getPrev() is nothing then
                ' ポインタが進んでいない場合
                call p.init(-1, head, nothing)
            else
                ' ポインタが進んでいる場合
                call p.setValue(0, item, tail)
            end if
        end if
    end function
    
    private function setToListedItems(byval index, byval item)
        ''' 要素数 1 以上の LinkedList の指定した位置に要素を追加する。
        ''' ポインタ位置に要素が追加された場合、ポインタ位置は後ろへずれる。
        ''' ポインタ位置よりも前に要素が追加された場合、ポインタ位置は後ろへずれる。
        ''' 第一引数 index として、追加する要素の位置 (Number) を受け取る。
        ''' 第二引数 val として、追加する要素 (Variant) を受け取る。
        ''' 戻り値は返さない。
        
        if index = 0 then
            ' list の先頭に要素を追加する場合
            call head.setPrev(item)
            call item.setNext(head)
            set head = item
            
            if p.getPrev() is nothing then
                ' ポインタが進んでいない場合
                call p.init(-1, head, nothing)
            else
                ' ポインタが進んでいる場合
                call p.setValue(p.getValue() + 1)
            end if
            exit function
        elseif index = count then
            ' list の末尾に要素を追加する場合
            call tail.setNext(item)
            call item.setPrev(tail)
            set tail = item
            
            ' ポインタが末尾を指している場合
            if p.getNext() is nothing then call p.setNext(item)
            exit function
        end if
        
        ' list の先頭と末尾以外に要素を追加する場合
        dim nx, pr
        set nx = getItemPointer(index)
        set pr = nx.getPrev()
        
        call nx.setPrev(item)
        call pr.setNext(item)
        call item.setNext(nx)
        call item.setPrev(pr)
        
        if index <= p.getValue() then
            ' ポインタの位置より前に要素が追加された場合
            call p.setValue(p.getValue() + 1)
        elseif index = p.getValue() + 1 then
            ' ポインタの次の位置に要素が追加された場合
            call p.setNext(item)
        end if
    end function
    
    public function pop()
        ''' LinkedList の末尾の要素を取り出す（削除する）。ポインタ位置が移動する場合がある。
        ''' 戻り値として、LinkedList の末尾の要素 (Variant) を返す。
        call bind(pop, remove(count - 1))
    end function
    
    public function shift()
        ''' LinkedList の先頭の要素を取り出す（削除する）。ポインタ位置が移動する場合がある。
        ''' 戻り値として、LinkedList の先頭の要素 (Variant) を返す。
        call bind(shift, remove(0))
    end function
    
    public function remove(byval index)
        ''' 指定した位置の要素を取り出す（削除する）。
        ''' ポインタ位置の要素が取り出された場合、ポインタ位置は前へずれる。
        ''' ポインタ位置よりも前の要素が取り出された場合、ポインタ位置は前へずれる。
        ''' 取り出す位置として不正な値が指定された場合はエラーを発生させる。
        ''' 第一引数 index として、取り出す要素の位置 (Number) を受け取る。
        ''' 戻り値として、取り出した要素 (Variant) を返す。
        if index > count - 1 then err.raise(9)
        if index < 0 then err.raise(9)
        
        dim item
        if head is tail then
            set item = getFromHeadIsTail()
        else
            set item = getFromListedItems(index)
        end if
        
        count = count - 1
        call bind(remove, item.getValue())
    end function
    
    private function getFromHeadIsTail()
        ''' 要素数 1 の LinkedList から要素を取り出す（削除する）。LinkedList は初期化される。
        ''' 戻り値として、取り出した要素への参照 (ListItem) を返す。
        dim item
        set item = head
        set head = nothing
        set tail = nothing
        call p.init(-1, nothing, nothing)
        set getFromHeadIsTail = item
    end function
    
    private function getFromListedItems(byval index)
        ''' 要素数 1 以上の LinkedList の指定した位置から要素を取り出す（削除する）。
        ''' ポインタ位置の要素が取り出された場合、ポインタ位置は前へずれる。
        ''' ポインタ位置よりも前の要素が取り出された場合、ポインタ位置は前へずれる。
        ''' 第一引数 index として、取り出す要素の位置 (Number) を受け取る。
        ''' 戻り値として、取り出した要素への参照 (ListItem) を返す。
        dim item
        if index = 0 then
            ' list の先頭の要素を削除する場合
            set item = head
            set head = head.getNext()
            call head.setPrev(nothing)
            
            ' 要素の削除によってポインタがずれる
            if p.getNext() is item then
                call p.init(-1, head, nothing)
            elseif p.getPrev() is item then
                call p.init(0, head.getNext(), head)
            else
                call p.setValue(p.getValue() -1)
            end if
            
            set getFromListedItems = item
            exit function
        elseif index = count - 1 then
            ' list の末尾の要素を削除する場合
            set item = tail
            set tail = tail.getPrev()
            call tail.setNext(nothing)
            
            ' 要素の削除によってポインタがずれる
            if p.getPrev() is item then
                call p.init(p.getValue() - 1, nothing, item.getPrev())
            elseif p.getNext() is item then
                call p.setNext(nothing)
            else
                ' 何もしない
            end if
            
            set getFromListedItems = item
            exit function
        end if
        
        dim target
        set target = getItemPointer(index)
        
        call target.getNext().setPrev(target.getPrev())
        call target.getPrev().setNext(target.getNext())
        set item = target
        
        ' 要素の削除によってポインタがずれる
        if p.getValue() >= index then
            call p.init(p.getValue() - 1, item.getNext(), item.getPrev())
        elseif p.getNext() is item then
            call p.init(p.getValue(), item.getNext(), item.getPrev())
        end if
        
        set getFromListedItems = item
    end function
    
    public function item(byval index)
        ''' getItem メソッドのショートカット。
        call bind(item, getItem(index))
    end function
    
    public function getItem(byval index)
        ''' LinkedList の前から index 番目の要素を返す。ポインタ (最後にアクセスした要素) の位置が index に移動する。
        ''' 第一引数 index として、要素の位置 (Number) を受け取る。
        ''' 戻り値として、LinkedList の前から index 番目の要素 (Variant) を返す。
        dim item
        set item = getItemPointer(index)
        call bind(getItem, getItem(index).getValue())
        call p.init(index, item.getNext(), item)
    end function
    
    public function setItem(byval index, byval val)
        ''' LinkedList の前から index 番目の要素を書き換える。ポインタ (最後にアクセスした要素) の位置が index に移動する。
        ''' 第一引数 index として、要素の位置 (Number) を受け取る。
        ''' 第二引数 val として、書き換える値 (Variant) を受け取る。
        ''' 戻り値を返さない。
        dim item
        set item = getItemPointer(index)
        call item.setValue(val)
        call p.init(index, item.getNext(), item)
    end function
    
    private function getItemPointer(byval index)
        ''' LinkedList の前から index 番目の要素を返す。ポインタは移動しない。
        ''' index として不正な位置を受け取った場合、エラーを発生させる。
        ''' 第一引数 index として、要素の位置 (Number) を受け取る。
        ''' 戻り値として、LinkedList の前から index 番目の要素 (Variant) を返す。
        if index > count - 1 then err.raise(9)
        if index < 0 then err.raise(9)
        
        ' ポインタ位置を考慮して、目的の位置へ最短の位置から検索を開始する。
        dim escape, pivot, item
        set escape = p.clone()
        
        pivot = p.getValue() ' ポインタの位置を基準とする
        
        if index = pivot + 1 then
            ' index がポインタの次の要素だった場合
            set item = p.getNext()
        elseif index = pivot then
            ' index がポインタ位置の要素だった場合
            set item = p.getPrev()
        elseif pivot = -1 then
            ' ポインタが初期位置にある場合（走査は head または tail から開始される）
            if count \ 2 >= index then
                set item = forward(head, index)
            else
                set item = backward(tail, count - index - 1)
            end if
        elseif pivot > index then
            ' index が head と pivot の間にある場合
            if pivot \ 2 < index then
                set item = backward(p, pivot - index - 1)
            else
                set item = forward(head, index)
            end if
        else
            ' index が pivot と tail の間にある場合
            if (count - pivot) \ 2 < index then
                set item = backward(tail, count - index - 1)
            else
                set item = forward(p, index - pivot)
            end if
        end if
        
        set p = escape
        
        set getItemPointer = item
    end function
    
    private function forward(byval from, byval distance)
        ''' ある要素から指定された数だけ後ろの要素への参照を返す。
        ''' 第一引数 from として、現在のポインタ (ListItem) を受け取る。
        ''' 第二引数 distance として、移動する距離 (Number) を受け取る。
        ''' 戻り値として、ある要素から指定された数だけ後ろの要素への参照 (ListItem) を返す。
        dim item
        set item = from
        do while distance <> 0
            set item = item.getNext()
            distance = distance - 1
        loop
        set forward = item
    end function
    
    private function backward(byval from, byval distance)
        ''' ある要素から指定された数だけ前の要素への参照を返す。
        ''' 第一引数 from として、現在のポインタ (ListItem) を受け取る。
        ''' 第二引数 distance として、移動する距離 (Number) を受け取る。
        ''' 戻り値として、ある要素から指定された数だけ前の要素への参照 (ListItem) を返す。
        dim item
        set item = from
        do while distance <> 0
            set item = item.getPrev()
            distance = distance - 1
        loop
        set backward = item
    end function
    
    public function nextItem()
        ''' ポインタ (最後にアクセスした要素) の位置の次の要素を返す。ポインタの位置は次の要素へ移る。
        ''' どの要素にもアクセスしたことがない場合は最初の要素を返す。
        ''' ポインタが末尾の要素にある場合は未定義 (Empty) を返す。
        ''' 戻り値として、ポインタ位置の次の要素 (Variant or Empty) を返す。
        dim ret
        if p.getNext() is nothing then
            ret = Empty
        else
            call bind(ret, p.getNext().getValue())
            call p.init(p.getValue() + 1, p.getNext().getNext(), p.getNext())
        end if
        call bind(nextItem, ret)
    end function
    
    public function hasNextItem()
        ''' 次の要素がある場合は True を返す。
        ''' 戻り値として、次の要素の有無 (True or False) を返す。
        if p.getNext() is nothing then
            hasNextItem = false
        else
            hasNextItem = true
        end if
    end function
    
    public function currentItem()
        ''' ポインタ (最後にアクセスした要素) の位置にある要素を返す。
        ''' どの要素にもアクセスしたことがない場合はエラーを発生させる。
        ''' 戻り値として、ポインタ位置の要素を返す。
        if p.getPrev() is nothing then Err.raise(9)
        call bind(currentItem, p.getPrev().getValue())
    end function
    
    public function index()
        ''' ポインタ (最後にアクセスした要素) の位置を返す。
        ''' どの要素にもアクセスしたことが無い場合は -1 を返す。
        ''' 戻り値として、ポインタの位置 (Number) を返す。
        index = p.getValue()
    end function
    
    public function concat(byval list)
        ''' この LinkedList の後ろに別の LinkedList を連結する。破壊的に動作する。
        ''' 第一引数として、別の LinkedList への参照 (LinkedList) を受け取る。
        ''' 戻り値として、自分自身への参照 (LinkedList) を返す。
        if TypeName(list) <> TypeName(me) then Err.raise(13)
        if me is list then set list = me.clone()
        
        dim escape, item
        escape = list.index()
        list.setIndex(-1)
        call bind(item, list.nextItem())
        do while not isEmpty(item)
            me.push(item)
            call bind(item, list.nextItem())
        loop
        call list.setIndex(escape)
        
        set concat = me
    end function
    
    public function rewind()
        ''' ポインタ (最後にアクセスした要素への参照) 位置を初期位置に戻す。
        ''' 戻り値を返さない。
        call p.init(-1, head, nothing)
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
        
        dim escape
        set escape = p.clone()
        
        dim length, i, result
        length = me.length()
        i = start
        result = -1
        if isObject(key) then
            do while i < length
                if key.equals(me.item(i)) then
                    result = i
                    exit do
                end if
                i = i + 1
            loop
        else
            do while i < length
                if key = me.item(i) then
                    result = i
                    exit do
                end if
                i = i + 1
            loop
        end if
        
        set p = escape
        
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
        
        dim escape
        set escape = p.clone()
        
        dim i, result
        i = start
        result = -1
        if isObject(key) then
            do while i >= 0
                if key.equals(me.item(i)) then
                    result = i
                    exit do
                end if
                i = i - 1
            loop
        else
            do while i >= 0
                if key = me.item(i) then
                    result = i
                    exit do
                end if
                i = i - 1
            loop
        end if
        
        set p = escape
        
        lastIndexOf = result
    end function
    
    public function clone()
        ''' LinkedList を複製する。
        ''' 戻り値として、LinkedList の複製 (LinkedList) を返す。
        dim list, index, length, i
        set list = new LinkedList
        index = me.index()
        length = me.length()
        i = 0
        do while i < length
            call list.push(me.item(i))
            i = i + 1
        loop
        set clone = list
    end function
    
    public function toArray()
        ''' LinkedList の全ての要素を標準の配列に格納して返す。LinkedList が空の場合は null を返す。
        ''' 戻り値として、LinkedList の全ての要素を格納した配列 (Array or null) を返す。
        if count = 0 then
            toArray = null
            exit function
        end if
        
        redim ary(count - 1)
        
        dim item, i
        set item = head
        i = 0
        do while not item is nothing
            call bind(ary(i), item.getValue())
            set item = item.getNext()
            i = i + 1
        loop
        
        toArray = ary
    end function
    
    public function toString(byval sep)
        ''' LinkedList の全ての要素を表す文字列を返す。
        ''' 第一引数 sep として、要素を分割する文字列 (String) を受け取る。
        ''' 戻り値として、LinkedList の全ての要素を表す文字列 (String) を返す。
        dim ary
        ary = toArray()
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
        toString = join(ary, sep)
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
