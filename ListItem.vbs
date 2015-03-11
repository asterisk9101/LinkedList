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
