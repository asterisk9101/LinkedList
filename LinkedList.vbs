class ListItem
    ''' �o�����A�����X�g�̗v�f��\���N���X
    private value
    private next_
    private prev_
    ''' ListItem ��3�̓����p�����[�^�����B
    ''' value �́A�v�f���ێ����Ă���l���i�[����B
    ''' next_ �́A���̗v�f�ւ̎Q�Ƃ��i�[����B
    ''' prev_ �́A�O�̗v�f�ւ̎Q�Ƃ��i�[����B
    
    private sub Class_Initialize
        ''' ListItem �N���X�̃R���X�g���N�^
        value = Empty
        set next_ = nothing
        set prev_ = nothing
    end sub
    
    private sub Class_Terminate
        ''' ListItem �N���X�̃f�X�g���N�^
        value = Empty
        set next_ = nothing
        set prev_ = nothing
    end sub
    
    public function init(byval val, byval nx, byval pr)
        ''' ListItem �̓����p�����[�^���ꊇ���Đݒ肷��B
        ''' setValue, setNext, setPrev ���\�b�h�ő�ւł���B
        ''' ������ val �Ƃ��āAListItem �Ɋi�[����l (variant) ���󂯎��B
        ''' ������ nx �Ƃ��āA���̗v�f�ւ̎Q�� (ListItem or Nothing) ���󂯎��B
        ''' ��O���� pr �Ƃ��āA�O�̗v�f�ւ̎Q�� (ListItem or Nothing) ���󂯎��B
        ''' �߂�l�Ƃ��āA�������g�ւ̎Q�� (ListItem) ��Ԃ��B
        call setValue(val)
        call setNext(nx)
        call setPrev(pr)
        set init = me
    end function
    
    public function setValue(byval val)
        ''' ListItem �� value �p�����[�^�ɒl���Z�b�g����B
        ''' ������ val �Ƃ��āAListItem �Ɋi�[����l (variant) ���󂯎��B
        ''' �߂�l�Ƃ��āA�������g�ւ̎Q�� (ListItem) ��Ԃ��B
        call bind(value, val)
        set setValue = me
    end function
    
    public function setNext(byval item)
        ''' ListItem �� next_ �p�����[�^�ɁA���̗v�f�ւ̎Q�Ƃ��Z�b�g����B
        ''' ������ item �Ƃ��āA���̗v�f�ւ̎Q�� (ListItem or Nothing) ���󂯎��B
        ''' �߂�l�Ƃ��āA�������g�ւ̎Q�� (ListItem) ��Ԃ��B
        if not item is nothing and TypeName(me) <> TypeName(item) then Err.raise(13) ' Type Miss Match
        set next_ = item
        set setNext = me
    end function
    
    public function setPrev(byval item)
        ''' ListItem �� prev_ �p�����[�^�ɁA�O�̗v�f�ւ̎Q�Ƃ��Z�b�g����B
        ''' ������ item �Ƃ��āA�O�̗v�f�ւ̎Q�� (ListItem or Nothing) ���󂯎��B
        ''' �߂�l�Ƃ��āA�������g�ւ̎Q�� (ListItem) ��Ԃ��B
        if not item is nothing and TypeName(me) <> TypeName(item) then Err.raise() ' Type Miss Match
        set prev_ = item
        set setPrev = me
    end function
    
    public function getValue()
        ''' ListItem �� value �p�����[�^�̒l��Ԃ��B
        ''' �߂�l�Ƃ��āAvalue �p�����[�^�̒l (variant) ��Ԃ��B
        call bind(getValue, value)
    end function
    
    public function getNext()
        ''' ListItem �̎��̗v�f�ւ̎Q�Ƃ�Ԃ��B
        ''' �߂�l�Ƃ��āA���̗v�f�ւ̎Q�� (ListItem) ��Ԃ��B
        set getNext = next_
    end function
    
    public function getPrev()
        ''' ListItem �̑O�̗v�f�ւ̎Q�Ƃ�Ԃ��B
        ''' �߂�l�Ƃ��āA�O�̗v�f�ւ̎Q�� (ListItem) ��Ԃ��B
        set getPrev = prev_
    end function
    
    public function clone()
        ''' ListItem �𕡐�����B
        ''' �߂�l�Ƃ��āA���������v�f�ւ̎Q�� (ListItem) ��Ԃ��B
        set clone = (new ListItem).init(value, next_, prev_)
    end function
    
    private function bind(byref out, byval val)
        ''' ���������I�u�W�F�N�g�����̑��̒l���𔻒f���đ�������������B
        ''' �������͎Q�ƂƂ��Ď󂯎�邽�߁A���̃��[�e�B���e�B�֐��ő�����ꂽ���ʂ͌Ăяo�����̕ϐ��ɔ��f�����B
        ''' ������ out �Ƃ��āA�C�ӂ̕ϐ��ւ̎Q�Ƃ��󂯎��B
        ''' ������ val �Ƃ��āA�C�ӂ̒l���󂯎��B
        ''' �߂�l�͕Ԃ��Ȃ��B
        if isObject(val) then
            set out = val
        else
            out = val
        end if
    end function
end class

class LinkedList
    ''' �o�����A�����X�g��\���N���X�B
    private head
    private tail
    private cursor
    private position
    private count
    ''' �N���X�ϐ� head �́ALinkedList �̐擪�̗v�f�ւ̎Q�� (ListItem) ���i�[����B
    ''' �N���X�ϐ� tail �́ALinkedList �̖����̗v�f�ւ̎Q�� (ListItem) ���i�[����B
    ''' �N���X�ϐ� cursor �́A�Ō�ɃA�N�Z�X�����v�f�ւ̎Q�� (ListItem) ���i�[����B
    ''' �N���X�ϐ� position �́Acursor �̎w���v�f�̈ʒu (Number) ���i�[����B
    ''' �N���X�ϐ� count �́ALinkedList �̗v�f�̐� (Number) ���i�[����B
    
    private sub Class_Initialize
        ''' LinkedList �̃R���X�g���N�^
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
        ''' LinkedList �̖����ɗv�f��ǉ�����B
        ''' �߂�l�Ƃ��āA���X�g�Ɋ܂܂��v�f�̐� (Number) ��Ԃ��B
        call insert(count, val)
        push = count
    end function
    
    public function unshift(byval val)
        ''' LinkedList �̐擪�ɗv�f��ǉ�����B
        ''' �߂�l�Ƃ��āA���X�g�Ɋ܂܂��v�f�̐� (Number) ��Ԃ��B
        call insert(0, val)
        unshift = count
    end function
    
    public function insert(byval index, byval val)
        ''' �w�肵���ʒu�ɗv�f��ǉ�����B
        ''' �J�[�\���ʒu�ɗv�f���ǉ����ꂽ�ꍇ�A�J�[�\���ʒu�͌��ֈړ�����B
        ''' �J�[�\���ʒu�����O�ɗv�f���ǉ����ꂽ�ꍇ�A�J�[�\���ʒu�͌��ֈړ�����B
        ''' �s���Ȉʒu���w�肳�ꂽ�ꍇ�̓G���[�𔭐�������B
        ''' ������ index �Ƃ��āA�ǉ�����v�f�̈ʒu (Number) ���󂯎��B
        ''' ������ val �Ƃ��āA�ǉ�����v�f (Variant) ���󂯎��B
        ''' �߂�l�Ƃ��āA���X�g�Ɋ܂܂��v�f�̐� (Number) ��Ԃ��B
        if index > count then err.raise(9)
        if index < 0 then err.raise(9)
        
        dim item, target
        set item = (new ListItem).init(val, nothing, nothing)
        if count = 0 then
            ' ��̃��X�g�ɗv�f��ǉ�����ꍇ
            set head = item
            set tail = item
        elseif index = count then
            ' ���X�g�̍Ō�ɗv�f��ǉ�����ꍇ
            call item.setPrev(tail)
            call tail.setNext(item)
            set tail = item
        else
            ' ���X�g�̊Ԃɗv�f��ǉ�����ꍇ
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
        ''' LinkedList �̖����̗v�f�����o���i�폜����j�B�J�[�\���ʒu���ړ�����ꍇ������B
        ''' �߂�l�Ƃ��āALinkedList �̖����̗v�f (Variant) ��Ԃ��B
        call bind(pop, remove(count - 1))
    end function
    
    public function shift()
        ''' LinkedList �̐擪�̗v�f�����o���i�폜����j�B�J�[�\���ʒu���ړ�����ꍇ������B
        ''' �߂�l�Ƃ��āALinkedList �̐擪�̗v�f (Variant) ��Ԃ��B
        call bind(shift, remove(0))
    end function
    
    public function remove(byval index)
        ''' �w�肵���ʒu�̗v�f�����o���i�폜����j�B
        ''' �J�[�\���ʒu�̗v�f�����o���ꂽ�ꍇ�A�J�[�\���ʒu�͑O�ւ����B
        ''' �J�[�\���ʒu�����O�̗v�f�����o���ꂽ�ꍇ�A�J�[�\���ʒu�͑O�ւ����B
        ''' ���o���ʒu�Ƃ��ĕs���Ȓl���w�肳�ꂽ�ꍇ�̓G���[�𔭐�������B
        ''' ������ index �Ƃ��āA���o���v�f�̈ʒu (Number) ���󂯎��B
        ''' �߂�l�Ƃ��āA���o�����v�f (Variant) ��Ԃ��B
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
        ''' getItem ���\�b�h�̃V���[�g�J�b�g�B
        call bind(item, getItem(index))
    end function
    
    public function getItem(byval index)
        ''' LinkedList �̑O���� index �Ԗڂ̗v�f��Ԃ��B�J�[�\�� (�Ō�ɃA�N�Z�X�����v�f) �̈ʒu�� index �Ɉړ�����B
        ''' ������ index �Ƃ��āA�v�f�̈ʒu (Number) ���󂯎��B
        ''' �߂�l�Ƃ��āALinkedList �̑O���� index �Ԗڂ̗v�f (Variant) ��Ԃ��B
        dim item
        set item = getItemPointer(index)
        call bind(getItem, item.getValue())
        set cursor = item
        position = index
    end function
    
    public function setItem(byval index, byval val)
        ''' LinkedList �̑O���� index �Ԗڂ̗v�f������������B�J�[�\�� (�Ō�ɃA�N�Z�X�����v�f) �̈ʒu�� index �Ɉړ�����B
        ''' ������ index �Ƃ��āA�v�f�̈ʒu (Number) ���󂯎��B
        ''' ������ val �Ƃ��āA����������l (Variant) ���󂯎��B
        ''' �߂�l��Ԃ��Ȃ��B
        dim item
        set item = getItemPointer(index)
        call item.setValue(val)
        set cursor = item
        position = index
    end function
    
    private function getItemPointer(byval index)
        ''' LinkedList �̑O���� index �Ԗڂ̗v�f��Ԃ��B�J�[�\���͈ړ����Ȃ��B
        ''' index �Ƃ��ĕs���Ȉʒu���󂯎�����ꍇ�A�G���[�𔭐�������B
        ''' ������ index �Ƃ��āA�v�f�̈ʒu (Number) ���󂯎��B
        ''' �߂�l�Ƃ��āALinkedList �̑O���� index �Ԗڂ̗v�f (Variant) ��Ԃ��B
        if index > count - 1 then err.raise(9)
        if index < 0 then err.raise(9)
        
        dim fromHead, fromTail, fromCursor, distance
        fromHead = index
        fromTail = count - index - 1
        fromCursor = abs(position - index)
        distance = min(array(fromHead, fromTail, fromCursor)) ' �ŒZ�������v�Z����
        
        ' �ŒZ�����ƂȂ�ʒu���猟���J�n
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
        ret = ary(0) ' ���g�Ȃ��̏ꍇ�̓G���[
        for each iter in ary
            if ret > iter then ret = iter
        next
        min = ret
    end function
    
    private function forward(byval from, byval distance)
        ''' ����v�f����w�肳�ꂽ���������̗v�f�ւ̎Q�Ƃ�Ԃ��B
        ''' ������ from �Ƃ��āA���݂̃J�[�\�� (ListItem) ���󂯎��B
        ''' ������ distance �Ƃ��āA�ړ����鋗�� (Number) ���󂯎��B
        ''' �߂�l�Ƃ��āA����v�f����w�肳�ꂽ���������̗v�f�ւ̎Q�� (ListItem) ��Ԃ��B
        do while distance <> 0
            set from = from.getNext()
            distance = distance - 1
        loop
        set forward = from
    end function
    
    private function backward(byval from, byval distance)
        ''' ����v�f����w�肳�ꂽ�������O�̗v�f�ւ̎Q�Ƃ�Ԃ��B
        ''' ������ from �Ƃ��āA���݂̃J�[�\�� (ListItem) ���󂯎��B
        ''' ������ distance �Ƃ��āA�ړ����鋗�� (Number) ���󂯎��B
        ''' �߂�l�Ƃ��āA����v�f����w�肳�ꂽ�������O�̗v�f�ւ̎Q�� (ListItem) ��Ԃ��B
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
        ''' ���̗v�f������ꍇ�� True ��Ԃ��B
        ''' �߂�l�Ƃ��āA���̗v�f�̗L�� (True or False) ��Ԃ��B
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
        ''' �J�[�\�� (�Ō�ɃA�N�Z�X�����v�f) �̈ʒu�ɂ���v�f��Ԃ��B
        ''' �ǂ̗v�f�ɂ��A�N�Z�X�������Ƃ��Ȃ��ꍇ�̓G���[�𔭐�������B
        ''' �߂�l�Ƃ��āA�J�[�\���ʒu�̗v�f��Ԃ��B
        if cursor is nothing then 
            currentItem = empty
        else
            call bind(currentItem, cursor.getValue())
        end if
    end function
    
    public function index()
        ''' �J�[�\�� (�Ō�ɃA�N�Z�X�����v�f) �̈ʒu��Ԃ��B
        ''' �ǂ̗v�f�ɂ��A�N�Z�X�������Ƃ������ꍇ�� -1 ��Ԃ��B
        ''' �߂�l�Ƃ��āA�J�[�\���̈ʒu (Number) ��Ԃ��B
        index = position
    end function
    
    public function concat(byval list)
        ''' ���� LinkedList �̌��ɕʂ� LinkedList ��A������B�j��I�ɓ��삷��B
        ''' �������Ƃ��āA�ʂ� LinkedList �ւ̎Q�� (LinkedList) ���󂯎��B
        ''' �߂�l�Ƃ��āA�������g�ւ̎Q�� (LinkedList) ��Ԃ��B
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
        ''' �J�[�\�� (�Ō�ɃA�N�Z�X�����v�f�ւ̎Q��) �ʒu�������ʒu�ɖ߂��B
        ''' �߂�l��Ԃ��Ȃ��B
        set cursor = nothing
        position = -1
    end function
    
    public function indexOf(byval key, byval start)
        ''' �w�肳�ꂽ�ʒu�������֌������Č������A�Ώۂ̗v�f�𔭌������ʒu��Ԃ��B�����ł��Ȃ������ꍇ�A-1 ��Ԃ��B
        ''' ��������l�Ƃ��ăI�u�W�F�N�g���󂯎�����ꍇ�A���̃I�u�W�F�N�g�� equals ���\�b�h���g�p���ŏ��� true �Ԃ��v�f�𔭌������ʒu��Ԃ��B
        ''' �������J�n����ʒu�Ƃ��ĕs���Ȓl���󂯎�����ꍇ�A�G���[�𔭐�������B
        ''' ������ key �Ƃ��āA��������l (variant) ���󂯎��B
        ''' ������ start �Ƃ��āA�������J�n����ʒu (Number) ���󂯎��B
        ''' �߂�l�Ƃ��āALinkedList ������֌������đΏۂ̗v�f�𔭌������ʒu (Number) ��Ԃ��B
        
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
        ''' �w�肳�ꂽ�ʒu����O���֌������Č������A�Ώۂ̗v�f�𔭌������ʒu��Ԃ��B�����ł��Ȃ������ꍇ�A-1 ��Ԃ��B
        ''' ��������l�Ƃ��ăI�u�W�F�N�g���󂯎�����ꍇ�A���̃I�u�W�F�N�g�� equals ���\�b�h���g�p���ŏ��� true �Ԃ��v�f�𔭌������ʒu��Ԃ��B
        ''' �������J�n����ʒu�Ƃ��ĕs���Ȓl���󂯎�����ꍇ�A�G���[�𔭐�������B
        ''' ������ key �Ƃ��āA��������l (variant) ���󂯎��B
        ''' ������ start �Ƃ��āA�������J�n����ʒu (Number) ���󂯎��B
        ''' �߂�l�Ƃ��āALinkedList ��O���֌������đΏۂ̗v�f�𔭌������ʒu (Number) ��Ԃ��B
        
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
        ''' LinkedList �𕡐�����B�J�[�\���͕������Ȃ��B
        ''' �߂�l�Ƃ��āALinkedList �̕��� (LinkedList) ��Ԃ��B
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
        ''' LinkedList �̑S�Ă̗v�f��W���̔z��Ɋi�[���ĕԂ��BLinkedList ����̏ꍇ�� null ��Ԃ��B
        ''' �߂�l�Ƃ��āALinkedList �̑S�Ă̗v�f���i�[�����z�� (Array or null) ��Ԃ��B
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
        ''' LinkedList �̑S�Ă̗v�f��\���������Ԃ��B
        ''' ������ sep �Ƃ��āA�v�f�𕪊����镶���� (String) ���󂯎��B
        ''' �߂�l�Ƃ��āALinkedList �̑S�Ă̗v�f��\�������� (String) ��Ԃ��B
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
        ''' LinkedList �̒�����Ԃ��B
        ''' �߂�l�Ƃ��āALinkedList �̒��� (Number) ��Ԃ��B
        length = count
    end function
    
    private function bind(byref out, byval val)
        ''' ���������I�u�W�F�N�g�����̑��̒l���𔻒f���đ�������������B
        ''' ���������Q�ƂƂ��Ď󂯎�邽�߁A���̃��[�e�B���e�B�֐��ő�����ꂽ���ʂ͌Ăяo�����̕ϐ��ɔ��f�����B
        ''' ������ out �Ƃ��āA�C�ӂ̕ϐ��ւ̎Q�Ƃ��󂯎��B
        ''' ������ val �Ƃ��āA�C�ӂ̒l���󂯎��B
        ''' �߂�l�͕Ԃ��Ȃ��B
        if isObject(val) then
            set out = val
        else
            out = val
        end if
    end function
end class
