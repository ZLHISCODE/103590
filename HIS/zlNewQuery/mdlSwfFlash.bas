Attribute VB_Name = "mdlSwfFlash"
Public Type FLASHHEADER
    intIsFlashMovie As Integer      '是否是SWF文件,或者发生错误
    lMHeight As Long                '电影的高      Pix
    lMWidth As Long                 '电影的宽      Pix
    bColorR As Byte                 '背景颜色的R值 Number
    bColorG As Byte                 '背景颜色的G值 Number
    bColorB As Byte                 '背景颜色的B值 Number
    intMTotalFrames As Integer      '电影的总帧数 Frames
    lMSize As Long                  '电影的大小  ByteNumber
    intMRate As Integer             '电影的速度  FPS
    bMVersion As Byte               '制作电影的Flash版本
End Type

'=========================================
'功能：取Flash文件的头部结构
'用法：
'  Dim FH As FlashHeader
'  FH = GetFlashHeader(strFlashFileName)
'返回:
'成功：FlashHeader结构        FH.intIsFlashMovie=1
'错误：文件找不到             FH.intIsFlashMovie=-1
'      不是FlashMovie文件     FH.intIsFlashMovie=0
'      未知错误:              FH.intIsFlashMovie=2
'=========================================
Public Function GetFlashHeader(strFileName As String) As FLASHHEADER
    Dim lFileNumber As Long                      '文件号
    Dim b(20) As Byte
    Dim strSWFSignature As String * 3            'SWF的签名
    Dim intTagSize As Integer                    '标签块的大小
    Dim lMWidth As Long                          '电影的宽
    Dim lMHeight As Long                         '电影的高
    Dim bMVersion As Byte
    Dim bColorR As Byte                          '背景颜色的R值 Number
    Dim bColorG As Byte                          '背景颜色的G值 Number
    Dim bColorB As Byte                          '背景颜色的B值 Number
    Dim intMTotalFrames As Integer               '电影的总帧数 Frames
    Dim lMSize As Long                           '电影的大小  ByteNumber
    Dim intMRate(1)  As Byte                     '电影的速度  FPS 帧每秒
    Dim nBites As Integer                        '一个Tag的大小,表示一个Tag占有的Bit位数
    
    Dim i As Integer
    Dim Tmpstring As String
 
    On Error GoTo errHand:
    
    '如果文件不存在，返回-1
    If Dir(strFileName) = "" Then
        GetFlashHeader.intIsFlashMovie = -1
        Exit Function
    End If
    
     '打开文件
    lFileNumber = FreeFile
    Open strFileName For Binary As #lFileNumber
         '读取签名
         Get #lFileNumber, , strSWFSignature
         '如果不是SWF文件，返回
         If strSWFSignature <> "FWS" Then
            GetFlashHeader.intIsFlashMovie = 0
            Close #lFileNumber
            Exit Function
         End If
         
         Get #lFileNumber, , bMVersion      '版本
         Get #lFileNumber, , lMSize         '电影大小
         Get #lFileNumber, , b()
         
        '第九位的开始(二进制码)的前5比特为这个标签的nBites
        '结构如下
        'Field      Type                Comment
        'Nbits      nBits = UB[5]       Bits in each rect value field
        'Xmin       SB[nBits]           X minimum position for rect
        'Xmax       SB[nBits]           X maximum position for rect
        'Ymin       SB[nBits]           Y minimum position for rect
        'Ymax       SB[nBits]           Y maximum position for rect
        '我不知道怎样去前5个Bits的内容（通过And 和 Or可以的），所以我选把那那些内容读出来，
        '转为二进制字符串,取前5个,再转为数字,
         nBites = Bin2Dec(Left(Dec2Bin(b(0)), 5))
         intTagSize = (nBites * 4 + 5) \ 8 + 1
         '动画的大小在关闭文件之后再计算
         
         Get #lFileNumber, 9 + intTagSize, intMRate         '速率
         Get #lFileNumber, , intMTotalFrames                '总帧数
         Get #lFileNumber, 9 + intTagSize + 6, bColorR      '背景颜色R
         Get #lFileNumber, , bColorG                        '背景颜色G
         Get #lFileNumber, , bColorB                        '背景颜色B
    Close #lFileNumber
    
    
    '取电影的原始高度
    '转为二进制字符串
    Tmpstring = ""
    For i = 0 To intTagSize - 1
        Tmpstring = Tmpstring & Dec2Bin(b(i))
    Next
    '宽：(第六个比特+nBites)开始,nBites长)\20
    GetFlashHeader.lMWidth = Bin2Dec(Mid(Tmpstring, 6 + nBites, nBites)) \ 20
    '高：(第六个比特+nBites*3)开始,nBites长)\20
    GetFlashHeader.lMHeight = Bin2Dec(Mid(Tmpstring, 6 + nBites * 3, nBites)) \ 20
    
    
    GetFlashHeader.intIsFlashMovie = 1
    GetFlashHeader.bMVersion = bMVersion
    GetFlashHeader.lMSize = lMSize
    GetFlashHeader.intMRate = intMRate(0) * 255 + intMRate(1)
    GetFlashHeader.intMTotalFrames = intMTotalFrames
    GetFlashHeader.bColorR = bColorR
    GetFlashHeader.bColorG = bColorG
    GetFlashHeader.bColorB = bColorB
    Exit Function
errHand:
   GetFlashHeader.intIsFlashMovie = 2
End Function

Public Function Bin2Dec(strBin As String) As Long
'功能：二进制转为十进制函数
'用法：Long  bin2dec(strBin as String)
'返回：  二进制的十进制 长整数（Long）
'错误：  返回-1
    
    Dim lDec As Long
    Dim lCount As Long
    Dim i As Long
    
    On Error GoTo errHand
    lDec = 0
    If strBin = "" Then strBin = "0"

    lCount = Len(strBin)
    For i = 1 To lCount
        lDec = lDec + CInt(Left(strBin, 1)) * 2 ^ (Len(strBin) - 1)
        strBin = Right(strBin, Len(strBin) - 1)
        DoEvents
    Next
    Bin2Dec = lDec
    Exit Function
errHand:
    Bin2Dec = -1
End Function

Public Function Dec2Bin(bDec As Byte) As String
'功能：十进制转为二进制函数
'用法：String  Dec2Bin(Bdec as Byte)
'返回：  十进制的二进制 字符串(String)
'错误：  返回"0"
    
    Dim strBin As String
    
    On Error GoTo Err
    If bDec > 255 Then
        Dec2Bin = "-1"
        Exit Function
    End If

    strBin = ""

    '转为字符串
    While bDec > 0
        strBin = bDec Mod 2 & strBin
        bDec = Fix(bDec / 2)
'        DoEvents
    Wend
    '补零足8位
    If Len(strBin) < 9 Then
        While Len(strBin) < 8
            strBin = "0" & strBin
        Wend
    End If
    Dec2Bin = strBin
    
    Exit Function
Err:
   Dec2Bin = "0"
End Function

