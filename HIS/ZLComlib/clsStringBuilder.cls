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
'@模块 clsStringBuilder-2019/8/9
'@编写 lshuo
'@功能
'    字符串缓冲区类。由vbAccelerator cStringBuilder、孤帆、CorLib等版本改造而来
'@引用
'
'@备注
'
Option Explicit

'---------------------------------------------------------------------------
'                0、API和常量声明
'---------------------------------------------------------------------------
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
'说明：将内存块从一个位置移动到另一个位置
'Destination:指向移动目的地起始地址的指针。
'Source:指向要移动的内存块起始地址的指针。
'Length:内存块的大小以字节为单位移动。
'注意事项：这个函数定义为RtlMoveMemory函数。它的实现是内联的。有关更多信息，请参见WinBase。h和Winnt.h。源和目标块可能会重叠。
'           第一个参数，目的地，必须足够大，以容纳长度字节的源;否则，可能会出现缓冲区溢出。这可能导致拒绝服务攻击，如果有访问违反，或者在最坏的情况下，允许攻击者向您的进程注入可执行代码。如果目的地是一个基于堆栈的缓冲区，则尤其如此。要注意，最后一个参数，长度，是将字节复制到目的地的数量，而不是目的地的大小。
Private Declare Function StrCSpn Lib "shlwapi.dll" Alias "StrCSpnW" (ByVal lpStr As Long, ByVal lpCharacters As Long) As Long
'@功能
'    搜索字符串中任何一组字符的第一次出现。搜索方法区分大小写，并且在搜索模式匹配中包含终止NULL字符。
'@原型
'    int StrCSpnW(
'      PCWSTR pszStr,
'      PCWSTR pszSet
'    );
'@参数
'pszStr
'    指向要搜索的以null结尾的字符串的指针?
'pszSet
'    指向以null结尾的字符串的指针，该字符串包含要搜索的字符。
'@返回值
'    返回pszSet中任何字符在pszStr中第一次出现的索引，如果没有找到匹配，则返回pszStr的长度。
'@备注
'    这个函数的返回值等于pszStr中不包含pszSet中任何字符的初始子字符串的长度。
'@Requirements
'    Minimum supported client    Windows 2000 Professional, Windows XP [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Target Platform Windows
'    Header Shlwapi.h
'    Library Shlwapi.lib
'    DLL Shlwapi.dll (version 4.71 or later)
Private Declare Function StrCSpnI Lib "shlwapi.dll" Alias "StrCSpnIW" (ByVal lpStr As Long, ByVal lpCharacters As Long) As Long
'@功能
'    搜索字符串中任何一组字符的第一次出现。搜索方法不区分大小写，并且在搜索模式匹配中包含终止NULL字符。
'@原型
'    int StrCSpnIW(
'      PCWSTR pszStr,
'      PCWSTR pszSet
'    );
'@参数
'pszStr
'    指向要搜索的以null结尾的字符串的指针?
'pszSet
'    指向以null结尾的字符串的指针，该字符串包含要搜索的字符。
'@返回值
'    返回pszSet中任何字符在pszStr中第一次出现的索引，如果没有找到匹配，则返回pszStr的长度。
'@备注
'    这个函数的返回值等于pszStr中不包含pszSet中任何字符的初始子字符串的长度?
'@Requirements
'    Minimum supported client    Windows 2000 Professional, Windows XP [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Target Platform Windows
'    Header Shlwapi.h
'    Library Shlwapi.lib
'    DLL Shlwapi.dll (version 4.71 or later)
Private Declare Function StrRStr Lib "shell32.dll" Alias "StrRStrW" (ByVal lpStart As Long, ByVal lpEnd As Long, ByVal lpSrch As Long) As Long

Private Declare Function StrRStrI Lib "shell32.dll" Alias "StrRStrIW" (ByVal lpStart As Long, ByVal lpEnd As Long, ByVal lpSrch As Long) As Long
'@功能
'    搜索字符串中指定子字符串的最后一次出现。这种比较不区分大小写。
'@原型
'    PCWSTR StrRStrIW(
'      PCWSTR pszSource,
'      PCWSTR pszLast,
'      PCWSTR pszSrch
'    );
'@参数
'pszSource
'    指向以null结尾的源字符串的指针?
'pszLast
'    指向定义搜索范围的源字符串的指针。将pszLast设置为指向源字符串中的一个字符，搜索将使用前面的字符停止。将pszLast设置为NULL，以搜索整个源字符串。
'pszSrch
'    指向要搜索的子字符串的指针?
'@返回值
'    如果成功，返回子字符串最后一次出现的地址，否则返回NULL。
'@Requirements
'    Minimum supported client    Windows 2000 Professional, Windows XP [desktop apps only]
'    Minimum supported server    Windows 2000 Server [desktop apps only]
'    Target Platform Windows
'    Header Shlwapi.h
'    Library Shlwapi.lib
'    DLL Shlwapi.dll (version 4.71 or later)
Private Const M_LONG_MIN_BLOCK_SIZE     As Long = 16
Private Const M_LONG_INIT_BLOCK_SIZE    As Long = 1024

Private Const M_INT_ZERO                As Integer = 48             '0符号
Private Const M_STR_ZERO                As String = "0"             '0符号
Private Const M_INT_NINE                As Integer = 57             '9符号
Private Const M_STR_NINE                As String = "9"             '9符号

Private Const M_INT_UPPERCASE_A         As Integer = 65             '大写A
Private Const M_STR_UPPERCASE_A         As String = "A"             '大写A
Private Const M_INT_LOWERCASE_A         As Integer = 97             '小写A
Private Const M_STR_LOWERCASE_A         As String = "a"             '小写A
Private Const M_INT_UPPERCASE_F         As Integer = 70             '大写F
Private Const M_STR_UPPERCASE_F         As String = "F"             '大写F
Private Const M_INT_LOWERCASE_F         As Integer = 102            '小写F
Private Const M_STR_LOWERCASE_F         As String = "f"             '小写F
'---------------------------------------------------------------------------
'                1、常规变量
'---------------------------------------------------------------------------
Private mbytBuffer()            As Byte             '数据实际存储区域
Private mlnglpBuffer            As Long             '分配的内存块指针
'---------------------------------------------------------------------------
'                2、属性变量与定义
'---------------------------------------------------------------------------
Private mlngAllocSize           As Long             '实际分配的内存大小
Private mlngLength              As Long             '当前已经使用的内存大小
Private mlngBlockSize           As Long             '块大小，按块增长缓冲区
'Length(R):当前缓存的字符数量。
Public Property Get Length() As Long
   Length = mlngLength \ 2
End Property

'Capacity(R/W)分配或获取指定大小的缓冲区，字符单位。当不够用时可以自动扩展大小，预先分配，防止多次分配。
Property Let Capacity(ByVal lngNewSize As Long)
    If lngNewSize < 1 Then Exit Property
    '由于vb的字符串是BSTR字符占2个字节所以分配2倍的内存
    lngNewSize = lngNewSize + lngNewSize
    If lngNewSize <= mlngAllocSize And mlnglpBuffer <> 0 Then Exit Property  '内存块还够用且已经分配
    If lngNewSize Mod mlngBlockSize = 0 Then
        lngNewSize = lngNewSize + mlngBlockSize
    Else
        lngNewSize = ((lngNewSize + mlngBlockSize - 1) \ mlngBlockSize) * mlngBlockSize
    End If
    '分配一个缓冲区
    If mlnglpBuffer = 0 Then    '没有数据或还没有分配过内存
        ReDim mbytBuffer(0 To lngNewSize)
    Else
        ReDim Preserve mbytBuffer(0 To lngNewSize)
    End If
    mlnglpBuffer = VarPtr(mbytBuffer(0))                 '缓冲区指针

    mlngAllocSize = lngNewSize
End Property

Public Property Get Capacity() As Long
    Capacity = mlngAllocSize \ 2
End Property

'BlockSize(R/W):缓冲区不够时，自动增长的数量，字符单位。缺省值1024
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
'Value(R/W):设置当前缓冲区内容
Public Property Let Value(ByRef strThis As String)
    Dim lngLen As Long
   
    lngLen = Len(strThis)
    If lngLen = 0 Then
        mlngLength = 0
    Else
        '内存分配
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
'                3、公共方法
'---------------------------------------------------------------------------
'@方法    SetStringBuilder
'   初始化StringBuilder对象，设置初始容量以及增长的块大小。
'@返回值 clsStringBuilder
'   方便链式处理
'@参数:
'lngInitCapacity Long In
'   缓冲区的初始容量大小，按字符算。若不是lngBockSize的整数倍，则自动换算成lngBockSize的整数倍。若当前的容量大于lngInitCapacity，则该设置不会生效
'lngBockSize Long In Optional -1
'   为-1时，使用当前值，不修改。<16时，lngBockSize为16
'@备注
'
Public Function SetStringBuilder(ByVal lngInitCapacity As Long, Optional ByVal lngBockSize As Long) As clsStringBuilder
    Me.BlockSize = lngBockSize
    Me.Capacity = lngInitCapacity
    Set SetStringBuilder = Me
End Function

'@方法    SetValue
'   清空原有内容，重新设置值
'@返回值  clsStringBuilder
'   方便链式处理
'@参数:
'strValue String In
'   当前值
'@备注
'
Public Function SetValue(ByRef strValue As String) As clsStringBuilder
    Me.Value = strValue
    Set SetValue = Me
End Function

'@方法    Append
'   添加一个字符串，返回当前对象，方便链式处理
'@返回值  clsStringBuilder
'
'@参数:
'strAppend String In
'   要添加的字符换
'lngTimes In Long Optional 1
'   添加次数,为0则不添加，方便编程
'@备注
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

'@方法    AppendOneUnioacdeChar4
'   添加一个Unicoade字符，转换为\uXXXX
'@返回值  clsStringBuilder
'
'@参数:
'intChar In Integer
'   字符的ASC码值，可用ChrW获取
'@备注
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

'@方法    AppendOneChar
'   添加一个字符，返回当前对象，方便链式处理
'@返回值  clsStringBuilder
'
'@参数:
'intChar In Integer
'   字符的ASC码值，可用ChrW获取
'lngTimes In Long Optional 1
'   添加次数,为0则不添加，方便编程
'@备注
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

'@方法    AppendOneCharBy4Hex
'   添加一个由4位16进制字符串表示的字符，返回当前对象，方便链式处理
'@返回值  clsStringBuilder
'
'@参数:
'intHex1 In Integer
'   第一位16进制字符的
'intHex2 In Integer
'   第二位16进制字符的
'intHex3 In Integer
'   第三位16进制字符的
'intHex4 In Integer
'   第四位16进制字符的
'@备注
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

'@方法    AppendOneCharBy2Hex
'   添加一个由两位16进制字符串表示的字符，返回当前对象，方便链式处理
'@返回值  clsStringBuilder
'
'@参数:
'Param 方向
'
'@备注
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
'@方法    AppendChars
'   添加一个字符数组，返回当前对象，方便链式处理
'@返回值  clsStringBuilder
'
'@参数:
'intChar In Integer()
'   字符的ASC码值，可用ChrW获取
'lngLength In Long Optional -1
'   截取的长度。-1表示自动截取从起始位置开始的所有字符
'lngStartIndex In Long
'   开始截取的位置。-1表示从从开始截取
'@备注
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
'@方法    AppendNewLine
'   添加一个换行符
'@返回值  clsStringBuilder
'   方便链式处理
'@参数:
'@备注
'
Public Function AppendNewLine() As clsStringBuilder
    Set AppendNewLine = Append(vbNewLine)
End Function

'@方法    Insert
'   字符串中间插入一个字符
'@返回值  clsStringBuilder
'   方便链式处理
'@参数:
'lngIndex Long In
'   插入lngIndex的后面位置。当插入位置>长度时，自动在后面添加。当插入位置<=0时自动在前面增加
'strInsertValue String In
'   插入的字符串
'@备注
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
        '将原始数据后移lngLen+lngLen,移动长度mlngLength - lngIndex + lngIndex
        RtlMoveMemory ByVal UnsignedLongAdd(lngPos, lngLen + lngLen), ByVal lngPos, mlngLength - lngIndex + lngIndex
        RtlMoveMemory ByVal lngPos, ByVal StrPtr(strInsertValue), lngLen + lngLen
        mlngLength = mlngLength + lngLen + lngLen
        Set Insert = Me
   End If
End Function

'@方法    Remove
'   移除一定量的字符串
'@返回值  clsStringBuilder
'   方便链式处理
'@参数:
'lngIndex Long In
'   移除lngIndex的后面位置。当移除位置>长度时，不移除。当插入位置<=0时，从开始位置移除
'lngLen Long In
'   移除的长度
'@备注
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

'@方法    IndexOf
'   寻找 str 在缓冲区中第一次出现的位置(字符位置从 1 数起)
'@返回值  Long
'
'@参数:
'strFind String In
'   要查找的字符串,按串每个字符查找，查找到则停止
'blnIsIgnoreCase Boolean In Optional True
'   是否忽视大小写(默认忽视大小写)
'@备注
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

'@方法    CharLastIndexOf
'   寻找 str 在缓冲区中最后一次出现的位置(字符位置从 1 数起)
'@返回值  Long
'
'@参数:
'strFind String In
'   要查找的字符串,按串每个字符查找，查找到则停止
'blnIsIgnoreCase Boolean In Optional True
'   是否忽视大小写(默认忽视大小写)
'@备注
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

'@方法    ToString
'   获取字符内容，可以指定起始、长度
'@返回值  String
'   返回缓冲区的字符串
'@参数:
'lngStartIndex Long In Optional -1
'   字符串的起始位置，<=1,则从开头开始
'lngLength      Long In Optional -1
'   转换的长度  <0则返回所有剩余长度
'@备注
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

'@方法    Reallocate
'   缓冲区过大时，自动根据当前值重新缩小空间，空间为大于当前值的最小的BLOCKSIZE的整数倍
'@返回值 clsStringBuilder
'   方便链式处理
'@参数:
'@备注
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


'@方法    Clear
'   清除缓存内容
'@返回值 clsStringBuilder
'   方便链式处理
'@参数:
'Param 方向
'
'@备注
'
Public Function Clear() As clsStringBuilder
    mlngLength = 0
    Set Clear = Me
End Function

'@方法    TearDwon
'   销毁缓存，以及整个分配
'@返回值 clsStringBuilder
'   方便链式处理
'@参数:
'@备注
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
'                4、私有方法
'---------------------------------------------------------------------------

'@方法    UnsignedAdd
'   无符号长整型加法
'@返回值  Long
'
'@参数:
'lX Long In 左值
'   -1表示4294967295,-2147483648表示2147483648，获取无符号的值，可以使用4294967296#*iif(Value<0,1,0)+Value来获取
'lY Long In 右值
'   -1表示4294967295,-2147483648表示2147483648，获取无符号的值，可以使用4294967296#*iif(Value<0,1,0)+Value来获取
'@备注
'   该方法即解决了Long溢出，也解决了Unsigned溢出。
'   (Start Xor &H80000000) + Incr Xor &H80000000，这个只解决了Long溢出，但是引入了Unsigned溢出。如2147483647+-2147483648
Public Function UnsignedLongAdd(ByVal lX As Long, ByVal lY As Long) As Long
    Dim lX4         As Long, lY4        As Long
    Dim lX8         As Long, lY8        As Long
    Dim lResult     As Long
    '提取最高位'
    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    '提起次高位'
    lX4 = lX And &H40000000
    lY4 = lY And &H40000000
    '剩余位直接相加'
    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)

    If lX4 And lY4 Then '次高位同时为 1，和为 &H80000000
        'lResult = lResult + 次高位的和（即&H80000000） + X的最高位 + Y的最高位
        '其中向更高位 &H100000000 的进位直接忽略了
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf lX4 Or lY4 Then '次高位只有一个 1'
        'lResult = lResult + 次高位的和（即&H40000000） + X的最高位 + Y的最高位'
        If lResult And &H40000000 Then '剩余位的和向次高位有进位
            '  lResult + 次高位的和（即&H40000000）
            '= lResult + &H80000000 - &H40000000
            '= lResult Xor &HC0000000
            lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
        Else
            lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
        End If
    Else '次高位全0，和为0'
        'lResult = lResult + 次高位和（即&H00000000） + X的最高位 + Y的最高位'
        lResult = lResult Xor lX8 Xor lY8
    End If

    UnsignedLongAdd = lResult
End Function
'@方法    VBMakeWord
'   将两个Byte生成一个Integer
'@返回值  Integer
'
'@参数:
'bytLow In  Byte
'   低字节
'bytHigh In Byte
'   高字节
'@备注
'
Private Function VBMakeWord(ByVal bytLow As Byte, ByVal bytHigh As Byte) As Integer
    VBMakeWord = bytHigh + bytLow * 256
End Function

'@方法    VBMakeByte
'   将低四位和高四位的值生成一个字节
'@返回值  Byte
'
'@参数:
'bytLow In Byte
'   低4位
'bytHigh In Byte
'   高四位
'@备注
'
Private Function VBMakeByte(ByVal bytLow As Byte, ByVal bytHigh As Byte) As Byte
    VBMakeByte = bytHigh + bytLow * 16
End Function
'@方法    HexToByte
'   将十六进制字符转换成字节
'@返回值  Byte
'
'@参数:
'Param 方向
'
'@备注
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
            '直接返回，会导致出错，错误外边获取
            HexCharToByte = intHexChar
    End Select
End Function
'---------------------------------------------------------------------------
'                5、对象方法与事件
'---------------------------------------------------------------------------
Private Sub Class_Initialize()
    mlngAllocSize = 0
    mlngLength = 0
    mlngBlockSize = M_LONG_INIT_BLOCK_SIZE * 2
End Sub

Private Sub Class_Terminate()
    Call TearDwon
End Sub

