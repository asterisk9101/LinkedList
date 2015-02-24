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
    private p
    private count
    ''' LinkedList �N���X�� 4 �̓����p�����[�^�����B
    ''' head �́ALinkedList �̐擪�̗v�f�ւ̎Q�� (ListItem) ���i�[����B
    ''' tail �́ALinkedList �̖����̗v�f�ւ̎Q�� (ListItem) ���i�[����B
    ''' p �́A�Ō�ɃA�N�Z�X���ꂽ�ʒu�̏����i�[�����v�f�ւ̎Q�� (ListItem) ���i�[����B
    '''     p.getValue() �́A�Ō�ɃA�N�Z�X�����v�f�̈ʒu (Number) ��Ԃ��B
    '''     p.getPrev() �́A�Ō�ɃA�N�Z�X�����v�f�ւ̎Q�� (ListItem) ��Ԃ��B
    '''     p.getNext() �́A�Ō�ɃA�N�Z�X�����v�f�̎��̗v�f�ւ̎Q�� (ListItem) ��Ԃ��B
    ''' count �́ALinkedList �̗v�f�̐� (Number) ���i�[����B
    
    private sub Class_Initialize
        ''' LinkedList �̃R���X�g���N�^
        set head = nothing
        set tail = nothing
        set p = new ListItem
        count = 0
    end sub
    
    private sub Class_Terminate
        ''' LinkedList �̃f�X�g���N�^
        set head = nothing
        set tail = nothing
        set p = nothing
        count = 0
    end sub
    
    public function push(byval val)
        ''' LinkedList �̖����ɗv�f��ǉ�����B�|�C���^�ʒu���ړ�����ꍇ������B
        ''' �߂�l�͕Ԃ��Ȃ��B
        call insert(count, val)
    end function
    
    public function unshift(byval val)
        ''' LinkedList �̐擪�ɗv�f��ǉ�����B�|�C���^�ʒu���ړ�����ꍇ������B
        ''' �߂�l�͕Ԃ��Ȃ��B
        call insert(0, val)
    end function
    
    public function insert(byval index, byval val)
        ''' �w�肵���ʒu�ɗv�f��ǉ�����B
        ''' �|�C���^�ʒu�ɗv�f���ǉ����ꂽ�ꍇ�A�|�C���^�ʒu�͌��ֈړ�����B
        ''' �|�C���^�ʒu�����O�ɗv�f���ǉ����ꂽ�ꍇ�A�|�C���^�ʒu�͌��ֈړ�����B
        ''' �s���Ȉʒu���w�肳�ꂽ�ꍇ�̓G���[�𔭐�������B
        ''' ������ index �Ƃ��āA�ǉ�����v�f�̈ʒu (Number) ���󂯎��B
        ''' ������ val �Ƃ��āA�ǉ�����v�f (Variant) ���󂯎��B
        ''' �߂�l�͕Ԃ��Ȃ��B
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
        ''' ��� LinkedList �ɗv�f��ǉ�����B
        ''' ������ val �Ƃ��āA�ǉ�����v�f (Variant) ���󂯎��B
        ''' �߂�l�͕Ԃ��Ȃ��B
        set head = item
        set tail = item
        call p.init(-1, item, nothing)
    end function
    
    private function setToHeadIsTail(byval index, byval item)
        ''' �v�f�� 1 �� LinkedList �ɗv�f��ǉ�����B
        ''' �|�C���^�ʒu�ɗv�f���ǉ����ꂽ�ꍇ�A�|�C���^�ʒu�͌��ւ����B
        ''' �|�C���^�ʒu�����O�ɗv�f���ǉ����ꂽ�ꍇ�A�|�C���^�ʒu�͌��ւ����B
        ''' ������ index �Ƃ��āA�ǉ�����v�f�̈ʒu (0 or 1) ���󂯎��B
        ''' ������ val �Ƃ��āA�ǉ�����v�f (Variant) ���󂯎��B
        ''' �߂�l�͕Ԃ��Ȃ��B
        if index = 0 then
            ' list �̐擪�ɗv�f���C���T�[�g����
            
            call head.setPrev(item)
            call item.setNext(head)
            set head = item
            
            if p.getPrev() is nothing then
                ' �|�C���^���i��ł��Ȃ��ꍇ
                call p.init(-1, item, nothing)
            else
                ' �|�C���^���i��ł���ꍇ
                call p.init(2, nothing, tail)
            end if
        else
            ' list �̖����ɗv�f���C���T�[�g����
            call head.setNext(item)
            call item.setPrev(item)
            set tail = item
            
            if p.getPrev() is nothing then
                ' �|�C���^���i��ł��Ȃ��ꍇ
                call p.init(-1, head, nothing)
            else
                ' �|�C���^���i��ł���ꍇ
                call p.setValue(0, item, tail)
            end if
        end if
    end function
    
    private function setToListedItems(byval index, byval item)
        ''' �v�f�� 1 �ȏ�� LinkedList �̎w�肵���ʒu�ɗv�f��ǉ�����B
        ''' �|�C���^�ʒu�ɗv�f���ǉ����ꂽ�ꍇ�A�|�C���^�ʒu�͌��ւ����B
        ''' �|�C���^�ʒu�����O�ɗv�f���ǉ����ꂽ�ꍇ�A�|�C���^�ʒu�͌��ւ����B
        ''' ������ index �Ƃ��āA�ǉ�����v�f�̈ʒu (Number) ���󂯎��B
        ''' ������ val �Ƃ��āA�ǉ�����v�f (Variant) ���󂯎��B
        ''' �߂�l�͕Ԃ��Ȃ��B
        
        if index = 0 then
            ' list �̐擪�ɗv�f��ǉ�����ꍇ
            call head.setPrev(item)
            call item.setNext(head)
            set head = item
            
            if p.getPrev() is nothing then
                ' �|�C���^���i��ł��Ȃ��ꍇ
                call p.init(-1, head, nothing)
            else
                ' �|�C���^���i��ł���ꍇ
                call p.setValue(p.getValue() + 1)
            end if
            exit function
        elseif index = count then
            ' list �̖����ɗv�f��ǉ�����ꍇ
            call tail.setNext(item)
            call item.setPrev(tail)
            set tail = item
            
            ' �|�C���^���������w���Ă���ꍇ
            if p.getNext() is nothing then call p.setNext(item)
            exit function
        end if
        
        ' list �̐擪�Ɩ����ȊO�ɗv�f��ǉ�����ꍇ
        dim nx, pr
        set nx = getItemPointer(index)
        set pr = nx.getPrev()
        
        call nx.setPrev(item)
        call pr.setNext(item)
        call item.setNext(nx)
        call item.setPrev(pr)
        
        if index <= p.getValue() then
            ' �|�C���^�̈ʒu���O�ɗv�f���ǉ����ꂽ�ꍇ
            call p.setValue(p.getValue() + 1)
        elseif index = p.getValue() + 1 then
            ' �|�C���^�̎��̈ʒu�ɗv�f���ǉ����ꂽ�ꍇ
            call p.setNext(item)
        end if
    end function
    
    public function pop()
        ''' LinkedList �̖����̗v�f�����o���i�폜����j�B�|�C���^�ʒu���ړ�����ꍇ������B
        ''' �߂�l�Ƃ��āALinkedList �̖����̗v�f (Variant) ��Ԃ��B
        call bind(pop, remove(count - 1))
    end function
    
    public function shift()
        ''' LinkedList �̐擪�̗v�f�����o���i�폜����j�B�|�C���^�ʒu���ړ�����ꍇ������B
        ''' �߂�l�Ƃ��āALinkedList �̐擪�̗v�f (Variant) ��Ԃ��B
        call bind(shift, remove(0))
    end function
    
    public function remove(byval index)
        ''' �w�肵���ʒu�̗v�f�����o���i�폜����j�B
        ''' �|�C���^�ʒu�̗v�f�����o���ꂽ�ꍇ�A�|�C���^�ʒu�͑O�ւ����B
        ''' �|�C���^�ʒu�����O�̗v�f�����o���ꂽ�ꍇ�A�|�C���^�ʒu�͑O�ւ����B
        ''' ���o���ʒu�Ƃ��ĕs���Ȓl���w�肳�ꂽ�ꍇ�̓G���[�𔭐�������B
        ''' ������ index �Ƃ��āA���o���v�f�̈ʒu (Number) ���󂯎��B
        ''' �߂�l�Ƃ��āA���o�����v�f (Variant) ��Ԃ��B
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
        ''' �v�f�� 1 �� LinkedList ����v�f�����o���i�폜����j�BLinkedList �͏����������B
        ''' �߂�l�Ƃ��āA���o�����v�f�ւ̎Q�� (ListItem) ��Ԃ��B
        dim item
        set item = head
        set head = nothing
        set tail = nothing
        call p.init(-1, nothing, nothing)
        set getFromHeadIsTail = item
    end function
    
    private function getFromListedItems(byval index)
        ''' �v�f�� 1 �ȏ�� LinkedList �̎w�肵���ʒu����v�f�����o���i�폜����j�B
        ''' �|�C���^�ʒu�̗v�f�����o���ꂽ�ꍇ�A�|�C���^�ʒu�͑O�ւ����B
        ''' �|�C���^�ʒu�����O�̗v�f�����o���ꂽ�ꍇ�A�|�C���^�ʒu�͑O�ւ����B
        ''' ������ index �Ƃ��āA���o���v�f�̈ʒu (Number) ���󂯎��B
        ''' �߂�l�Ƃ��āA���o�����v�f�ւ̎Q�� (ListItem) ��Ԃ��B
        dim item
        if index = 0 then
            ' list �̐擪�̗v�f���폜����ꍇ
            set item = head
            set head = head.getNext()
            call head.setPrev(nothing)
            
            ' �v�f�̍폜�ɂ���ă|�C���^�������
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
            ' list �̖����̗v�f���폜����ꍇ
            set item = tail
            set tail = tail.getPrev()
            call tail.setNext(nothing)
            
            ' �v�f�̍폜�ɂ���ă|�C���^�������
            if p.getPrev() is item then
                call p.init(p.getValue() - 1, nothing, item.getPrev())
            elseif p.getNext() is item then
                call p.setNext(nothing)
            else
                ' �������Ȃ�
            end if
            
            set getFromListedItems = item
            exit function
        end if
        
        dim target
        set target = getItemPointer(index)
        
        call target.getNext().setPrev(target.getPrev())
        call target.getPrev().setNext(target.getNext())
        set item = target
        
        ' �v�f�̍폜�ɂ���ă|�C���^�������
        if p.getValue() >= index then
            call p.init(p.getValue() - 1, item.getNext(), item.getPrev())
        elseif p.getNext() is item then
            call p.init(p.getValue(), item.getNext(), item.getPrev())
        end if
        
        set getFromListedItems = item
    end function
    
    public function item(byval index)
        ''' getItem ���\�b�h�̃V���[�g�J�b�g�B
        call bind(item, getItem(index))
    end function
    
    public function getItem(byval index)
        ''' LinkedList �̑O���� index �Ԗڂ̗v�f��Ԃ��B�|�C���^ (�Ō�ɃA�N�Z�X�����v�f) �̈ʒu�� index �Ɉړ�����B
        ''' ������ index �Ƃ��āA�v�f�̈ʒu (Number) ���󂯎��B
        ''' �߂�l�Ƃ��āALinkedList �̑O���� index �Ԗڂ̗v�f (Variant) ��Ԃ��B
        dim item
        set item = getItemPointer(index)
        call bind(getItem, getItem(index).getValue())
        call p.init(index, item.getNext(), item)
    end function
    
    public function setItem(byval index, byval val)
        ''' LinkedList �̑O���� index �Ԗڂ̗v�f������������B�|�C���^ (�Ō�ɃA�N�Z�X�����v�f) �̈ʒu�� index �Ɉړ�����B
        ''' ������ index �Ƃ��āA�v�f�̈ʒu (Number) ���󂯎��B
        ''' ������ val �Ƃ��āA����������l (Variant) ���󂯎��B
        ''' �߂�l��Ԃ��Ȃ��B
        dim item
        set item = getItemPointer(index)
        call item.setValue(val)
        call p.init(index, item.getNext(), item)
    end function
    
    private function getItemPointer(byval index)
        ''' LinkedList �̑O���� index �Ԗڂ̗v�f��Ԃ��B�|�C���^�͈ړ����Ȃ��B
        ''' index �Ƃ��ĕs���Ȉʒu���󂯎�����ꍇ�A�G���[�𔭐�������B
        ''' ������ index �Ƃ��āA�v�f�̈ʒu (Number) ���󂯎��B
        ''' �߂�l�Ƃ��āALinkedList �̑O���� index �Ԗڂ̗v�f (Variant) ��Ԃ��B
        if index > count - 1 then err.raise(9)
        if index < 0 then err.raise(9)
        
        ' �|�C���^�ʒu���l�����āA�ړI�̈ʒu�֍ŒZ�̈ʒu���猟�����J�n����B
        dim escape, pivot, item
        set escape = p.clone()
        
        pivot = p.getValue() ' �|�C���^�̈ʒu����Ƃ���
        
        if index = pivot + 1 then
            ' index ���|�C���^�̎��̗v�f�������ꍇ
            set item = p.getNext()
        elseif index = pivot then
            ' index ���|�C���^�ʒu�̗v�f�������ꍇ
            set item = p.getPrev()
        elseif pivot = -1 then
            ' �|�C���^�������ʒu�ɂ���ꍇ�i������ head �܂��� tail ����J�n�����j
            if count \ 2 >= index then
                set item = forward(head, index)
            else
                set item = backward(tail, count - index - 1)
            end if
        elseif pivot > index then
            ' index �� head �� pivot �̊Ԃɂ���ꍇ
            if pivot \ 2 < index then
                set item = backward(p, pivot - index - 1)
            else
                set item = forward(head, index)
            end if
        else
            ' index �� pivot �� tail �̊Ԃɂ���ꍇ
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
        ''' ����v�f����w�肳�ꂽ���������̗v�f�ւ̎Q�Ƃ�Ԃ��B
        ''' ������ from �Ƃ��āA���݂̃|�C���^ (ListItem) ���󂯎��B
        ''' ������ distance �Ƃ��āA�ړ����鋗�� (Number) ���󂯎��B
        ''' �߂�l�Ƃ��āA����v�f����w�肳�ꂽ���������̗v�f�ւ̎Q�� (ListItem) ��Ԃ��B
        dim item
        set item = from
        do while distance <> 0
            set item = item.getNext()
            distance = distance - 1
        loop
        set forward = item
    end function
    
    private function backward(byval from, byval distance)
        ''' ����v�f����w�肳�ꂽ�������O�̗v�f�ւ̎Q�Ƃ�Ԃ��B
        ''' ������ from �Ƃ��āA���݂̃|�C���^ (ListItem) ���󂯎��B
        ''' ������ distance �Ƃ��āA�ړ����鋗�� (Number) ���󂯎��B
        ''' �߂�l�Ƃ��āA����v�f����w�肳�ꂽ�������O�̗v�f�ւ̎Q�� (ListItem) ��Ԃ��B
        dim item
        set item = from
        do while distance <> 0
            set item = item.getPrev()
            distance = distance - 1
        loop
        set backward = item
    end function
    
    public function nextItem()
        ''' �|�C���^ (�Ō�ɃA�N�Z�X�����v�f) �̈ʒu�̎��̗v�f��Ԃ��B�|�C���^�̈ʒu�͎��̗v�f�ֈڂ�B
        ''' �ǂ̗v�f�ɂ��A�N�Z�X�������Ƃ��Ȃ��ꍇ�͍ŏ��̗v�f��Ԃ��B
        ''' �|�C���^�������̗v�f�ɂ���ꍇ�͖���` (Empty) ��Ԃ��B
        ''' �߂�l�Ƃ��āA�|�C���^�ʒu�̎��̗v�f (Variant or Empty) ��Ԃ��B
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
        ''' ���̗v�f������ꍇ�� True ��Ԃ��B
        ''' �߂�l�Ƃ��āA���̗v�f�̗L�� (True or False) ��Ԃ��B
        if p.getNext() is nothing then
            hasNextItem = false
        else
            hasNextItem = true
        end if
    end function
    
    public function currentItem()
        ''' �|�C���^ (�Ō�ɃA�N�Z�X�����v�f) �̈ʒu�ɂ���v�f��Ԃ��B
        ''' �ǂ̗v�f�ɂ��A�N�Z�X�������Ƃ��Ȃ��ꍇ�̓G���[�𔭐�������B
        ''' �߂�l�Ƃ��āA�|�C���^�ʒu�̗v�f��Ԃ��B
        if p.getPrev() is nothing then Err.raise(9)
        call bind(currentItem, p.getPrev().getValue())
    end function
    
    public function index()
        ''' �|�C���^ (�Ō�ɃA�N�Z�X�����v�f) �̈ʒu��Ԃ��B
        ''' �ǂ̗v�f�ɂ��A�N�Z�X�������Ƃ������ꍇ�� -1 ��Ԃ��B
        ''' �߂�l�Ƃ��āA�|�C���^�̈ʒu (Number) ��Ԃ��B
        index = p.getValue()
    end function
    
    public function concat(byval list)
        ''' ���� LinkedList �̌��ɕʂ� LinkedList ��A������B�j��I�ɓ��삷��B
        ''' �������Ƃ��āA�ʂ� LinkedList �ւ̎Q�� (LinkedList) ���󂯎��B
        ''' �߂�l�Ƃ��āA�������g�ւ̎Q�� (LinkedList) ��Ԃ��B
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
        ''' �|�C���^ (�Ō�ɃA�N�Z�X�����v�f�ւ̎Q��) �ʒu�������ʒu�ɖ߂��B
        ''' �߂�l��Ԃ��Ȃ��B
        call p.init(-1, head, nothing)
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
        ''' �w�肳�ꂽ�ʒu����O���֌������Č������A�Ώۂ̗v�f�𔭌������ʒu��Ԃ��B�����ł��Ȃ������ꍇ�A-1 ��Ԃ��B
        ''' ��������l�Ƃ��ăI�u�W�F�N�g���󂯎�����ꍇ�A���̃I�u�W�F�N�g�� equals ���\�b�h���g�p���ŏ��� true �Ԃ��v�f�𔭌������ʒu��Ԃ��B
        ''' �������J�n����ʒu�Ƃ��ĕs���Ȓl���󂯎�����ꍇ�A�G���[�𔭐�������B
        ''' ������ key �Ƃ��āA��������l (variant) ���󂯎��B
        ''' ������ start �Ƃ��āA�������J�n����ʒu (Number) ���󂯎��B
        ''' �߂�l�Ƃ��āALinkedList ��O���֌������đΏۂ̗v�f�𔭌������ʒu (Number) ��Ԃ��B
        
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
        ''' LinkedList �𕡐�����B
        ''' �߂�l�Ƃ��āALinkedList �̕��� (LinkedList) ��Ԃ��B
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
        ''' LinkedList �̑S�Ă̗v�f��W���̔z��Ɋi�[���ĕԂ��BLinkedList ����̏ꍇ�� null ��Ԃ��B
        ''' �߂�l�Ƃ��āALinkedList �̑S�Ă̗v�f���i�[�����z�� (Array or null) ��Ԃ��B
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
        ''' LinkedList �̑S�Ă̗v�f��\���������Ԃ��B
        ''' ������ sep �Ƃ��āA�v�f�𕪊����镶���� (String) ���󂯎��B
        ''' �߂�l�Ƃ��āALinkedList �̑S�Ă̗v�f��\�������� (String) ��Ԃ��B
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
