VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@ģ�� clsStringBuilder-2019/8/9
'@��д lshuo
'@����
'    �ַ����������ࡣ��vbAccelerator cStringBuilder���·���CorLib�Ȱ汾�������
'@����
'
'@��ע
'
Option Explicit

'---------------------------------------------------------------------------
'                0��API�ͳ�������
'---------------------------------------------------------------------------
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
'˵�������ڴ���һ��λ���ƶ�����һ��λ��
'Destination:ָ���ƶ�Ŀ�ĵ���ʼ��ַ��ָ�롣
'Source:ָ��Ҫ�ƶ����ڴ����ʼ��ַ��ָ�롣
'Length:�ڴ��Ĵ�С���ֽ�Ϊ��λ�ƶ���
'ע����������������ΪRtlMoveMemory����������ʵ���������ġ��йظ�����Ϣ����μ�WinBase��h��Winnt.h��Դ��Ŀ�����ܻ��ص���
'           ��һ��������Ŀ�ĵأ������㹻�������ɳ����ֽڵ�Դ;���򣬿��ܻ���ֻ��������������ܵ��¾ܾ����񹥻�������з���Υ�����������������£����������������Ľ���ע���ִ�д��롣���Ŀ�ĵ���һ�����ڶ�ջ�Ļ���������������ˡ�Ҫע�⣬���һ�����������ȣ��ǽ��ֽڸ��Ƶ�Ŀ�ĵص�������������Ŀ�ĵصĴ�С��
Private Declare Function StrCSpn Lib "shlwapi.dll" Alias "StrCSpnW" (ByVal lpStr As Long, ByVal lpCharacters As Long) As Long
'@����
'    �����ַ������κ�һ���ַ��ĵ�һ�γ��֡������������ִ�Сд������������ģʽƥ���а�����ֹNULL�ַ���
'@ԭ��
'    int StrCSpnW(
'      PCWSTR pszStr,
'      PCWSTR pszSet
'    );
'@����
'pszStr
'    ָ��Ҫ��������null��β���ַ�����ָ��?
'pszSet
'    ָ����null��β���ַ�����ָ�룬���ַ�������Ҫ�������ַ���
'@����ֵ
'    ����pszSet���κ��ַ���pszStr�е�һ�γ��ֵ����������û���ҵ�ƥ�䣬�򷵻�pszStr�ĳ��ȡ�
'@��ע
'    ��������ķ���ֵ����pszStr�в�����pszSet���κ��ַ��ĳ�ʼ���ַ����ĳ��ȡ�
'@Requirements
'    Minimum supported client    Windows 2000 Professional, Windows XP [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Target Platform Windows
'    Header Shlwapi.h
'    Library Shlwapi.lib
'    DLL Shlwapi.dll (version 4.71 or later)
Private Declare Function StrCSpnI Lib "shlwapi.dll" Alias "StrCSpnIW" (ByVal lpStr As Long, ByVal lpCharacters As Long) As Long
'@����
'    �����ַ������κ�һ���ַ��ĵ�һ�γ��֡��������������ִ�Сд������������ģʽƥ���а�����ֹNULL�ַ���
'@ԭ��
'    int StrCSpnIW(
'      PCWSTR pszStr,
'      PCWSTR pszSet
'    );
'@����
'pszStr
'    ָ��Ҫ��������null��β���ַ�����ָ��?
'pszSet
'    ָ����null��β���ַ�����ָ�룬���ַ�������Ҫ�������ַ���
'@����ֵ
'    ����pszSet���κ��ַ���pszStr�е�һ�γ��ֵ����������û���ҵ�ƥ�䣬�򷵻�pszStr�ĳ��ȡ�
'@��ע
'    ��������ķ���ֵ����pszStr�в�����pszSet���κ��ַ��ĳ�ʼ���ַ����ĳ���?
'@Requirements
'    Minimum supported client    Windows 2000 Professional, Windows XP [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Target Platform Windows
'    Header Shlwapi.h
'    Library Shlwapi.lib
'    DLL Shlwapi.dll (version 4.71 or later)
Private Declare Function StrRStr Lib "shell32.dll" Alias "StrRStrW" (ByVal lpStart As Long, ByVal lpEnd As Long, ByVal lpSrch As Long) As Long

Private Declare Function StrRStrI Lib "shell32.dll" Alias "StrRStrIW" (ByVal lpStart As Long, ByVal lpEnd As Long, ByVal lpSrch As Long) As Long
'@����
'    �����ַ�����ָ�����ַ��������һ�γ��֡����ֱȽϲ����ִ�Сд��
'@ԭ��
'    PCWSTR StrRStrIW(
'      PCWSTR pszSource,
'      PCWSTR pszLast,
'      PCWSTR pszSrch
'    );
'@����
'pszSource
'    ָ����null��β��Դ�ַ�����ָ��?
'pszLast
'    ָ����������Χ��Դ�ַ�����ָ�롣��pszLast����Ϊָ��Դ�ַ����е�һ���ַ���������ʹ��ǰ����ַ�ֹͣ����pszLast����ΪNULL������������Դ�ַ�����
'pszSrch
'    ָ��Ҫ���������ַ�����ָ��?
'@����ֵ
'    ����ɹ����������ַ������һ�γ��ֵĵ�ַ�����򷵻�NULL��
'@Requirements
'    Minimum supported client    Windows 2000 Professional, Windows XP [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Target Platform Windows
'    Header Shlwapi.h
'    Library Shlwapi.lib
'    DLL Shlwapi.dll (version 4.71 or later)
Private Const M_LONG_MIN_BLOCK_SIZE     As Long = 16
Private Const M_LONG_INIT_BLOCK_SIZE    As Long = 1024

Private Const M_INT_ZERO                As Integer = 48             '0����
Private Const M_STR_ZERO                As String = "0"             '0����
Private Const M_INT_NINE                As Integer = 57             '9����
Private Const M_STR_NINE                As String = "9"             '9����

Private Const M_INT_UPPERCASE_A         As Integer = 65             '��дA
Private Const M_STR_UPPERCASE_A         As String = "A"             '��дA
Private Const M_INT_LOWERCASE_A         As Integer = 97             'СдA
Private Const M_STR_LOWERCASE_A         As String = "a"             'СдA
Private Const M_INT_UPPERCASE_F         As Integer = 70             '��дF
Private Const M_STR_UPPERCASE_F         As String = "F"             '��дF
Private Const M_INT_LOWERCASE_F         As Integer = 102            'СдF
Private Const M_STR_LOWERCASE_F         As String = "f"             'СдF
'---------------------------------------------------------------------------
'                1���������
'---------------------------------------------------------------------------
Private mbytBuffer()            As Byte             '����ʵ�ʴ洢����
Private mlnglpBuffer            As Long             '������ڴ��ָ��
'---------------------------------------------------------------------------
'                2�����Ա����붨��
'---------------------------------------------------------------------------
Private mlngAllocSize           As Long             'ʵ�ʷ�����ڴ��С
Private mlngLength              As Long             '��ǰ�Ѿ�ʹ�õ��ڴ��С
Private mlngBlockSize           As Long             '���С����������������
'Length(R):��ǰ������ַ�������
Public Property Get Length() As Long
   Length = mlngLength \ 2
End Property

'Capacity(R/W)������ȡָ����С�Ļ��������ַ���λ����������ʱ�����Զ���չ��С��Ԥ�ȷ��䣬��ֹ��η��䡣
Property Let Capacity(ByVal lngNewSize As Long)
    If lngNewSize < 1 Then Exit Property
    '����vb���ַ�����BSTR�ַ�ռ2���ֽ����Է���2�����ڴ�
    lngNewSize = lngNewSize + lngNewSize
    If lngNewSize <= mlngAllocSize And mlnglpBuffer <> 0 Then Exit Property  '�ڴ�黹�������Ѿ�����
    If lngNewSize Mod mlngBlockSize = 0 Then
        lngNewSize = lngNewSize + mlngBlockSize
    Else
        lngNewSize = ((lngNewSize + mlngBlockSize - 1) \ mlngBlockSize) * mlngBlockSize
    End If
    '����һ��������
    If mlnglpBuffer = 0 Then    'û�����ݻ�û�з�����ڴ�
        ReDim mbytBuffer(0 To lngNewSize)
    Else
        ReDim Preserve mbytBuffer(0 To lngNewSize)
    End If
    mlnglpBuffer = VarPtr(mbytBuffer(0))                 '������ָ��

    mlngAllocSize = lngNewSize
End Property

Public Property Get Capacity() As Long
    Capacity = mlngAllocSize \ 2
End Property

'BlockSize(R/W):����������ʱ���Զ��������������ַ���λ��ȱʡֵ1024
Public Property Get BlockSize() As Long
    BlockSize = mlngBlockSize \ 2
End Property

Public Property Let BlockSize(ByVal lngBlockSize As Long)
    If lngBlockSize < 0 Then Exit Property
    If lngBlockSize < M_LONG_MIN_BLOCK_SIZE Then
        lngBlockSize = M_LONG_MIN_BLOCK_SIZE
    End If
    mlngBlockSize = lngBlockSize * 2
End Property
'Value(R/W):���õ�ǰ����������
Public Property Let Value(ByRef strThis As String)
    Dim lngLen As Long
   
    lngLen = Len(strThis)
    If lngLen = 0 Then
        mlngLength = 0
    Else
        '�ڴ����
        Me.Capacity = lngLen
        mlngLength = lngLen + lngLen
        RtlMoveMemory ByVal mlnglpBuffer, ByVal StrPtr(strThis), mlngLength
    End If
End Property

Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
    Value = toString()
End Property

'---------------------------------------------------------------------------
'                3����������
'---------------------------------------------------------------------------
'@����    SetStringBuilder
'   ��ʼ��StringBuilder�������ó�ʼ�����Լ������Ŀ��С��
'@����ֵ clsStringBuilder
'   ������ʽ����
'@����:
'lngInitCapacity Long In
'   �������ĳ�ʼ������С�����ַ��㡣������lngBockSize�������������Զ������lngBockSize��������������ǰ����������lngInitCapacity��������ò�����Ч
'lngBockSize Long In Optional -1
'   Ϊ-1ʱ��ʹ�õ�ǰֵ�����޸ġ�<16ʱ��lngBockSizeΪ16
'@��ע
'
Public Function SetStringBuilder(ByVal lngInitCapacity As Long, Optional ByVal lngBockSize As Long) As clsStringBuilder
    Me.BlockSize = lngBockSize
    Me.Capacity = lngInitCapacity
    Set SetStringBuilder = Me
End Function

'@����    SetValue
'   ���ԭ�����ݣ���������ֵ
'@����ֵ  clsStringBuilder
'   ������ʽ����
'@����:
'strValue String In
'   ��ǰֵ
'@��ע
'
Public Function SetValue(ByRef strValue As String) As clsStringBuilder
    Me.Value = strValue
    Set SetValue = Me
End Function

'@����    Append
'   ����һ���ַ��������ص�ǰ���󣬷�����ʽ����
'@����ֵ  clsStringBuilder
'
'@����:
'strAppend String In
'   Ҫ���ӵ��ַ���
'lngTimes In Long Optional 1
'   ���Ӵ���,Ϊ0�����ӣ�������
'@��ע
'
Public Function Append(ByRef strAppend As String, Optional ByVal lngTimes As Long = 1) As clsStringBuilder
    Dim lngLen  As Long, i      As Long
    lngLen = Len(strAppend)
    If lngTimes > 0 Then
        If lngLen > 0 Then
            If mlngLength + lngLen * 2 * lngTimes > mlngAllocSize Then
                Me.Capacity = Me.Length + lngLen * lngTimes
            End If
            For i = 0 To lngTimes - 1
                RtlMoveMemory ByVal UnsignedLongAdd(mlnglpBuffer, mlngLength + i * 2 * lngLen), ByVal StrPtr(strAppend), lngLen + lngLen
            Next
            mlngLength = mlngLength + lngLen * 2 * lngTimes
        End If
    End If
    Set Append = Me
End Function

'@����    AppendOneUnioacdeChar4
'   ����һ��Unicoade�ַ���ת��Ϊ\uXXXX
'@����ֵ  clsStringBuilder
'
'@����:
'intChar In Integer
'   �ַ���ASC��ֵ������ChrW��ȡ
'@��ע
'
Public Function AppendOneCharUnicode4(ByRef intChar As Integer, Optional ByVal lngTimes As Long = 1) As clsStringBuilder
    Dim lngLen  As Long, i      As Long
    Dim strTmp  As String

    lngLen = 6
    If lngTimes > 0 Then
        If mlngLength + lngLen * 2 * lngTimes > mlngAllocSize Then
            Me.Capacity = Me.Length + lngLen * lngTimes
        End If
        strTmp = "\u" & Right$("000" & Hex$(intChar), 4)
        For i = 0 To lngTimes - 1
            RtlMoveMemory ByVal UnsignedLongAdd(mlnglpBuffer, mlngLength + i * 2 * lngLen), ByVal StrPtr(strTmp), lngLen + lngLen
        Next
        mlngLength = mlngLength + lngLen * 2 * lngTimes
    End If
    Set AppendOneCharUnicode4 = Me
End Function

'@����    AppendOneChar
'   ����һ���ַ������ص�ǰ���󣬷�����ʽ����
'@����ֵ  clsStringBuilder
'
'@����:
'intChar In Integer
'   �ַ���ASC��ֵ������ChrW��ȡ
'lngTimes In Long Optional 1
'   ���Ӵ���,Ϊ0�����ӣ�������
'@��ע
'
Public Function AppendOneChar(ByRef intChar As Integer, Optional ByVal lngTimes As Long = 1) As clsStringBuilder
    Dim i       As Long
    If lngTimes > 0 Then
        If intChar <> 0 Then
            If mlngLength + 2 * lngTimes > mlngAllocSize Then
                Me.Capacity = Me.Length + lngTimes
            End If
            For i = 0 To lngTimes - 1
                RtlMoveMemory ByVal UnsignedLongAdd(mlnglpBuffer, mlngLength + i * 2), intChar, 2
            Next
            mlngLength = mlngLength + 2 * lngTimes
        End If
    End If
    Set AppendOneChar = Me
End Function

'@����    AppendOneCharBy4Hex
'   ����һ����4λ16�����ַ�����ʾ���ַ������ص�ǰ���󣬷�����ʽ����
'@����ֵ  clsStringBuilder
'
'@����:
'intHex1 In Integer
'   ��һλ16�����ַ���
'intHex2 In Integer
'   �ڶ�λ16�����ַ���
'intHex3 In Integer
'   ����λ16�����ַ���
'intHex4 In Integer
'   ����λ16�����ַ���
'@��ע
'
Public Function AppendOneCharBy4Hex(ByRef intHex1 As Integer, ByRef intHex2 As Integer, ByRef intHex3 As Integer, ByRef intHex4 As Integer) As clsStringBuilder
    Dim intChar     As Integer
    
    On Error GoTo ErrH
    intChar = VBMakeWord(VBMakeByte(HexCharToByte(intHex1), HexCharToByte(intHex2)), VBMakeByte(HexCharToByte(intHex3), HexCharToByte(intHex4)))
    If intChar <> 0 Then
        If mlngLength + 2 > mlngAllocSize Then
            Me.Capacity = Me.Length + 1
        End If
        RtlMoveMemory ByVal UnsignedLongAdd(mlnglpBuffer, mlngLength), intChar, 2
        mlngLength = mlngLength + 2
    End If
    Set AppendOneCharBy4Hex = Me
    Exit Function
ErrH:
    Err.Clear
End Function

'@����    AppendOneCharBy2Hex
'   ����һ������λ16�����ַ�����ʾ���ַ������ص�ǰ���󣬷�����ʽ����
'@����ֵ  clsStringBuilder
'
'@����:
'Param ����
'
'@��ע
'
Public Function AppendOneCharBy2Hex(ByRef intHex1 As Integer, ByRef intHex2 As Integer) As clsStringBuilder
    Dim intChar     As Integer
    On Error GoTo ErrH
    intChar = VBMakeByte(HexCharToByte(intHex1), HexCharToByte(intHex2))
    If intChar <> 0 Then
        If mlngLength + 2 > mlngAllocSize Then
            Me.Capacity = Me.Length + 1
        End If
        RtlMoveMemory ByVal UnsignedLongAdd(mlnglpBuffer, mlngLength), intChar, 2
        mlngLength = mlngLength + 2
    End If
    Set AppendOneCharBy2Hex = Me
    Exit Function
ErrH:
    Err.Clear
End Function
'@����    AppendChars
'   ����һ���ַ����飬���ص�ǰ���󣬷�����ʽ����
'@����ֵ  clsStringBuilder
'
'@����:
'intChar In Integer()
'   �ַ���ASC��ֵ������ChrW��ȡ
'lngLength In Long Optional -1
'   ��ȡ�ĳ��ȡ�-1��ʾ�Զ���ȡ����ʼλ�ÿ�ʼ�������ַ�
'lngStartIndex In Long
'   ��ʼ��ȡ��λ�á�-1��ʾ�Ӵӿ�ʼ��ȡ
'@��ע
'
Public Function AppendChars(ByRef intArrChars() As Integer, Optional ByVal lngLength As Long = -1, Optional ByVal lngStartIndex As Long = -1) As clsStringBuilder
    If lngStartIndex = -1 Then
        lngStartIndex = 0
    End If
    If lngLength = -1 Then
        lngLength = UBound(intArrChars) - lngStartIndex + 1
    End If
    If lngLength > 0 Then
        Me.Capacity = Me.Length + lngLength
        RtlMoveMemory ByVal UnsignedLongAdd(mlnglpBuffer, mlngLength), ByVal VarPtr(intArrChars(lngStartIndex)), lngLength + lngLength
        mlngLength = mlngLength + lngLength + lngLength
    End If
    Set AppendChars = Me
End Function
'@����    AppendNewLine
'   ����һ�����з�
'@����ֵ  clsStringBuilder
'   ������ʽ����
'@����:
'@��ע
'
Public Function AppendNewLine() As clsStringBuilder
    Set AppendNewLine = Append(vbNewLine)
End Function

'@����    Insert
'   �ַ����м����һ���ַ�
'@����ֵ  clsStringBuilder
'   ������ʽ����
'@����:
'lngIndex Long In
'   ����lngIndex�ĺ���λ�á�������λ��>����ʱ���Զ��ں������ӡ�������λ��<=0ʱ�Զ���ǰ������
'strInsertValue String In
'   ������ַ���
'@��ע
'
Public Function Insert(ByVal lngIndex As Long, ByRef strInsertValue As String) As clsStringBuilder
    Dim lngLen      As Long
    Dim lngPos      As Long
    Dim lSize As Long
   
    If (lngIndex * 2 >= mlngLength) Then
       Set Insert = Append(strInsertValue)
    Else
        lngLen = Len(strInsertValue)
        If lngIndex <= 0 Then
            lngIndex = 0
        End If
        Me.Capacity = Me.Length + lngLen
        lngPos = UnsignedLongAdd(mlnglpBuffer, lngIndex + lngIndex)
        '��ԭʼ���ݺ���lngLen+lngLen,�ƶ�����mlngLength - lngIndex + lngIndex
        RtlMoveMemory ByVal UnsignedLongAdd(lngPos, lngLen + lngLen), ByVal lngPos, mlngLength - lngIndex + lngIndex
        RtlMoveMemory ByVal lngPos, ByVal StrPtr(strInsertValue), lngLen + lngLen
        mlngLength = mlngLength + lngLen + lngLen
        Set Insert = Me
   End If
End Function

'@����    Remove
'   �Ƴ�һ�������ַ���
'@����ֵ  clsStringBuilder
'   ������ʽ����
'@����:
'lngIndex Long In
'   �Ƴ�lngIndex�ĺ���λ�á����Ƴ�λ��>����ʱ�����Ƴ���������λ��<=0ʱ���ӿ�ʼλ���Ƴ�
'lngLen Long In
'   �Ƴ��ĳ���
'@��ע
'
Public Function Remove(ByVal lngIndex As Long, ByVal lngLen As Long) As clsStringBuilder
    If lngIndex <= 0 Then
        lngIndex = 0
    ElseIf lngIndex + lngIndex >= mlngLength Then
        Set Remove = Me
        Exit Function
    End If
    If lngLen + lngLen >= mlngLength Then
        lngLen = mlngLength
    End If
    If (lngIndex + lngLen) * 2 > mlngLength Then
        lngLen = (mlngLength - lngIndex - lngIndex) / 2
    End If
    If lngLen + lngLen = mlngLength Then
        Me.Clear
    ElseIf lngLen > 0 Then
        RtlMoveMemory ByVal UnsignedLongAdd(mlnglpBuffer, lngIndex * 2), ByVal UnsignedLongAdd(mlnglpBuffer, (lngIndex + lngLen) * 2), (mlngLength - (lngIndex + lngLen) * 2)
        mlngLength = mlngLength - lngLen * 2
    End If
    Set Remove = Me
End Function

'@����    IndexOf
'   Ѱ�� str �ڻ������е�һ�γ��ֵ�λ��(�ַ�λ�ô� 1 ����)
'@����ֵ  Long
'
'@����:
'strFind String In
'   Ҫ���ҵ��ַ���,����ÿ���ַ����ң����ҵ���ֹͣ
'blnIsIgnoreCase Boolean In Optional True
'   �Ƿ���Ӵ�Сд(Ĭ�Ϻ��Ӵ�Сд)
'@��ע
'
Function CharIndexOf(ByVal strFind As String, Optional ByVal blnIsIgnoreCase As Boolean = True) As Long
    If mlngLength < 1 Then
        CharIndexOf = -1
        Exit Function
    End If

    If blnIsIgnoreCase Then
        CharIndexOf = StrCSpnI(mlnglpBuffer, StrPtr(strFind))
    Else
        CharIndexOf = StrCSpn(mlnglpBuffer, StrPtr(strFind))
    End If
    If CharIndexOf + CharIndexOf = mlngLength Then
        CharIndexOf = -1
    Else
        If CharIndexOf >= 0 Then CharIndexOf = CharIndexOf + 1
    End If
End Function

'@����    CharLastIndexOf
'   Ѱ�� str �ڻ����������һ�γ��ֵ�λ��(�ַ�λ�ô� 1 ����)
'@����ֵ  Long
'
'@����:
'strFind String In
'   Ҫ���ҵ��ַ���,����ÿ���ַ����ң����ҵ���ֹͣ
'blnIsIgnoreCase Boolean In Optional True
'   �Ƿ���Ӵ�Сд(Ĭ�Ϻ��Ӵ�Сд)
'@��ע
'
Function CharLastIndexOf(ByVal strFind As String, Optional ByVal blnIsIgnoreCase As Boolean = True) As Long
    If mlngLength < 1 Then
        CharLastIndexOf = -1
        Exit Function
    End If

    If blnIsIgnoreCase Then
        CharLastIndexOf = StrRStrI(mlnglpBuffer, mlnglpBuffer + mlngLength, StrPtr(strFind))
    Else
        CharLastIndexOf = StrRStr(mlnglpBuffer, mlnglpBuffer + mlngLength, StrPtr(strFind))
    End If

    If CharLastIndexOf >= mlnglpBuffer Then
        CharLastIndexOf = (CharLastIndexOf - mlnglpBuffer) / 2 + 1
    Else
        CharLastIndexOf = -1
    End If
End Function

'@����    ToString
'   ��ȡ�ַ����ݣ�����ָ����ʼ������
'@����ֵ  String
'   ���ػ��������ַ���
'@����:
'lngStartIndex Long In Optional -1
'   �ַ�������ʼλ�ã�<=1,��ӿ�ͷ��ʼ
'lngLength      Long In Optional -1
'   ת���ĳ���  <0�򷵻�����ʣ�೤��
'@��ע
'
Public Function toString(Optional ByVal lngStartIndex As Long = -1, Optional ByVal lngLength As Long = -1) As String
    If mlngLength > 0 Then
        If lngStartIndex < 1 Then
            If lngLength < 0 Then
                toString = Left$(mbytBuffer, mlngLength \ 2)
            ElseIf lngLength <> 0 Then
                If lngLength < mlngLength - lngLength Then
                    toString = Left$(mbytBuffer, lngLength)
                Else
                    toString = Left$(mbytBuffer, mlngLength \ 2)
                End If
            End If
        Else
            If lngLength < 0 Then
                toString = Mid$(mbytBuffer, lngStartIndex, mlngLength \ 2 - lngStartIndex + 1)
            ElseIf lngLength <> 0 Then
                If lngLength < mlngLength - lngLength Then
                    toString = Left$(mbytBuffer, lngLength)
                Else
                    toString = Left$(mbytBuffer, mlngLength \ 2)
                End If
            End If
        End If
    End If
End Function

'@����    Reallocate
'   ����������ʱ���Զ����ݵ�ǰֵ������С�ռ䣬�ռ�Ϊ���ڵ�ǰֵ����С��BLOCKSIZE��������
'@����ֵ clsStringBuilder
'   ������ʽ����
'@����:
'@��ע
'
Public Function Reallocate() As clsStringBuilder
    Dim lngNewSize      As Long
    If mlngAllocSize - mlngLength > mlngBlockSize Then
        lngNewSize = mlngLength
        If lngNewSize Mod mlngBlockSize = 0 Then
            lngNewSize = lngNewSize + mlngBlockSize
        Else
            lngNewSize = ((lngNewSize + mlngBlockSize - 1) \ mlngBlockSize) * mlngBlockSize
        End If
        ReDim Preserve mbytBuffer(0 To lngNewSize)
        mlngAllocSize = lngNewSize
    End If
    Set Reallocate = Me
End Function


'@����    Clear
'   �����������
'@����ֵ clsStringBuilder
'   ������ʽ����
'@����:
'Param ����
'
'@��ע
'
Public Function Clear() As clsStringBuilder
    mlngLength = 0
    Set Clear = Me
End Function

'@����    TearDwon
'   ���ٻ��棬�Լ���������
'@����ֵ clsStringBuilder
'   ������ʽ����
'@����:
'@��ע
'
Public Function TearDwon() As clsStringBuilder
    Erase mbytBuffer
    mlnglpBuffer = 0
    mlngAllocSize = 0
    mlngLength = 0
    mlngBlockSize = M_LONG_INIT_BLOCK_SIZE * 2
    Set TearDwon = Me
End Function
'---------------------------------------------------------------------------
'                4��˽�з���
'---------------------------------------------------------------------------

'@����    UnsignedAdd
'   �޷��ų����ͼӷ�
'@����ֵ  Long
'
'@����:
'lX Long In ��ֵ
'   -1��ʾ4294967295,-2147483648��ʾ2147483648����ȡ�޷��ŵ�ֵ������ʹ��4294967296#*iif(Value<0,1,0)+Value����ȡ
'lY Long In ��ֵ
'   -1��ʾ4294967295,-2147483648��ʾ2147483648����ȡ�޷��ŵ�ֵ������ʹ��4294967296#*iif(Value<0,1,0)+Value����ȡ
'@��ע
'   �÷����������Long�����Ҳ�����Unsigned�����
'   (Start Xor &H80000000) + Incr Xor &H80000000�����ֻ�����Long���������������Unsigned�������2147483647+-2147483648
Public Function UnsignedLongAdd(ByVal lX As Long, ByVal lY As Long) As Long
    Dim lX4         As Long, lY4        As Long
    Dim lX8         As Long, lY8        As Long
    Dim lResult     As Long
    '��ȡ���λ'
    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    '����θ�λ'
    lX4 = lX And &H40000000
    lY4 = lY And &H40000000
    'ʣ��λֱ�����'
    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)

    If lX4 And lY4 Then '�θ�λͬʱΪ 1����Ϊ &H80000000
        'lResult = lResult + �θ�λ�ĺͣ���&H80000000�� + X�����λ + Y�����λ
        '���������λ &H100000000 �Ľ�λֱ�Ӻ�����
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf lX4 Or lY4 Then '�θ�λֻ��һ�� 1'
        'lResult = lResult + �θ�λ�ĺͣ���&H40000000�� + X�����λ + Y�����λ'
        If lResult And &H40000000 Then 'ʣ��λ�ĺ���θ�λ�н�λ
            '  lResult + �θ�λ�ĺͣ���&H40000000��
            '= lResult + &H80000000 - &H40000000
            '= lResult Xor &HC0000000
            lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
        Else
            lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
        End If
    Else '�θ�λȫ0����Ϊ0'
        'lResult = lResult + �θ�λ�ͣ���&H00000000�� + X�����λ + Y�����λ'
        lResult = lResult Xor lX8 Xor lY8
    End If

    UnsignedLongAdd = lResult
End Function
'@����    VBMakeWord
'   ������Byte����һ��Integer
'@����ֵ  Integer
'
'@����:
'bytLow In  Byte
'   ���ֽ�
'bytHigh In Byte
'   ���ֽ�
'@��ע
'
Private Function VBMakeWord(ByVal bytLow As Byte, ByVal bytHigh As Byte) As Integer
    VBMakeWord = bytHigh + bytLow * 256
End Function

'@����    VBMakeByte
'   ������λ�͸���λ��ֵ����һ���ֽ�
'@����ֵ  Byte
'
'@����:
'bytLow In Byte
'   ��4λ
'bytHigh In Byte
'   ����λ
'@��ע
'
Private Function VBMakeByte(ByVal bytLow As Byte, ByVal bytHigh As Byte) As Byte
    VBMakeByte = bytHigh + bytLow * 16
End Function
'@����    HexToByte
'   ��ʮ�������ַ�ת�����ֽ�
'@����ֵ  Byte
'
'@����:
'Param ����
'
'@��ע
'
Private Function HexCharToByte(ByVal intHexChar As Integer) As Byte
    Select Case intHexChar
        Case M_INT_ZERO To M_INT_NINE
            HexCharToByte = intHexChar - M_INT_ZERO
        Case M_INT_LOWERCASE_A To M_INT_LOWERCASE_F
            HexCharToByte = intHexChar - M_INT_LOWERCASE_A + 10
        Case M_INT_UPPERCASE_A To M_INT_UPPERCASE_F
            HexCharToByte = intHexChar - M_INT_UPPERCASE_A + 10
        Case Else
            'ֱ�ӷ��أ��ᵼ�³�����������߻�ȡ
            HexCharToByte = intHexChar
    End Select
End Function
'---------------------------------------------------------------------------
'                5�����󷽷����¼�
'---------------------------------------------------------------------------
Private Sub Class_Initialize()
    mlngAllocSize = 0
    mlngLength = 0
    mlngBlockSize = M_LONG_INIT_BLOCK_SIZE * 2
End Sub

Private Sub Class_Terminate()
    Call TearDwon
End Sub
