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
