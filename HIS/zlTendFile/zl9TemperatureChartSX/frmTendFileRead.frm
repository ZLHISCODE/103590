VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTendFileRead 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "页眉页脚打印"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6840
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTendFileRead.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3465
      ScaleWidth      =   6585
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   6615
      Begin RichTextLib.RichTextBox rtbHead 
         Height          =   1380
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   6810
         _ExtentX        =   12012
         _ExtentY        =   2434
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   2
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"frmTendFileRead.frx":000C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtbFoot 
         Height          =   1380
         Left            =   0
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1950
         Width           =   6810
         _ExtentX        =   12012
         _ExtentY        =   2434
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   2
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"frmTendFileRead.frx":00A9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmTendFileRead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'##############################################################################################
'页眉页脚打印相关
Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type
'矩形
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'包含用于格式化指定设备的相关信息
Private Type FORMATRANGE
    hDC As Long             '渲染设备
    hdcTarget As Long       '目标设备
    rc As RECT              '渲染区域，单位：缇。
    rcPage As RECT          '渲染设备的整体区域，单位：缇。
    chrg As CHARRANGE       '用于格式化的文本范围。
End Type

Private Type PageInfo
    PageNumber As Long      '页码
    Start As Long           '字符起始位置
    End As Long             '字符终止位置
    ActualHeight As Long    '本页实际打印高度
End Type
Private AllPages() As PageInfo   '页信息
Private Const WM_PASTE = &H302&              '粘贴
Private Const WM_USER = &H400                '通常用 WM_USER + X 来自定义消息
Private Const EM_FORMATRANGE = (WM_USER + 57)    '为某一设备格式化指定范围的文本。
Private Const EM_SETTARGETDEVICE = (WM_USER + 72) '设置用于所见即所得的目标设备和行宽。
Private Const EM_HIDESELECTION = (WM_USER + 63)  '显示/隐藏文本。
Private Const PHYSICALOFFSETX = 112  '对于打印设备而言，表示从物理页的左边缘到可打印区域的左边缘的距离，采用设备单位。
Private Const PHYSICALOFFSETY = 113  '对于打印设备而言，表示从物理页的上边缘到可打印区域的上边缘的距离，采用设备单位。
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long '获取中英文混合字符串长度

Public Function InitRechBox(ByVal lngID As Long) As Boolean
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    '将将页眉页脚定义（病历页面格式）读取并装载到RichTextBox中
    strSQL = "Select  /*+ RULE */ 格式, 页脚, 种类||'-'||编号 AS KEY From 病历页面格式" & _
        "   Where 种类 = 3 And 编号 = (select 页面 from (select 页面 from 病历文件列表 where 种类=3  and 保留<>-1 order by 编号) where rownum<2)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取病历页面格式", lngID)
    If Not rsTmp.EOF Then
        '考虑到医院内护理文件页眉页脚格式统一，此处只读取一次
        Call ReadPageHead(rtbHead, rsTmp!Key)
        Call ReadPageFoot(rtbFoot, rsTmp!Key)
    End If
    
    InitRechBox = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function ReadPageHead(objHead As RichTextBox, ByVal strKey As String) As Boolean
'################################################################################################################
'## 功能：  读取页面图片
'## 参数：  病历种类-页面编号
'## 返回：  返回获得的图片变量。
'################################################################################################################
    Dim strFile As String, strZip As String
    strZip = zlBlobRead(12, strKey, App.Path & "\Head_L.zip")
    If gobjFSO.FileExists(strZip) Then
        strFile = UnzipTendPage(strZip, "Head_S.RTF")
        objHead.LoadFile strFile, rtfRTF           '读取文件
        gobjFSO.DeleteFile strFile, True      '删除临时文件
        ReadPageHead = True
    Else
        objHead.Text = ""
    End If
End Function

Public Function ReadPageFoot(objFoot As RichTextBox, ByVal strKey As String) As Boolean
'################################################################################################################
'## 功能：  读取页面图片  RichTextBox
'## 参数：  病历种类-页面编号
'## 返回：  返回获得的图片变量。
'################################################################################################################
    Dim strFile As String, strZip As String
    strZip = zlBlobRead(13, strKey, App.Path & "\Foot_L.zip")
    If gobjFSO.FileExists(strZip) Then
        strFile = UnzipTendPage(strZip, "Foot_S.RTF")
        objFoot.LoadFile strFile, rtfRTF           '读取文件
        gobjFSO.DeleteFile strFile, True      '删除临时文件
        ReadPageFoot = True
    Else
        objFoot.Text = ""
    End If
End Function

'################################################################################################################
'## 功能：  将指定的LOB字段复制为临时文件
'##
'## 参数：  Action      :操作类型（用以区别是操作哪个表）
'##         KeyWord     :确定数据记录的关键字，复合关键字以逗号分隔(仅5-电子病历格式为复合)
'##         strFile     :用户指定存放的文件名；不指定时，取当前路径产生文件名
'##
'## 返回：  存放内容的文件名，失败则返回零长度""
'##
'## 说明：  Action取值说明：
'##         0-病历标记图形；1-病历文件格式；2-病历文件图形；3-病历范文格式；4-病历范文图形；5-电子病历格式；6-电子病历图形；
'################################################################################################################
Public Function zlBlobRead(ByVal Action As Long, ByVal KeyWord As String, Optional ByRef strFile As String, Optional ByVal blnMoved As Boolean) As String
    
    Const conChunkSize As Integer = 10240
    Dim lngFileNum As Long, lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, strText As String
    Dim rsLob As New ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHand
    
    lngFileNum = FreeFile
    If strFile = "" Then
        lngCount = 0
        Do While True
            strFile = App.Path & "\zlBlobFile" & CStr(lngCount) & ".tmp"
            If Len(Dir(strFile)) = 0 Then Exit Do
            lngCount = lngCount + 1
        Loop
    End If
    Open strFile For Binary As lngFileNum
    
    gstrSQL = "Select Zl_Lob_Read([1],[2],[3],[4]) as 片段 From Dual"
    lngCount = 0
    Do
        Set rsLob = zlDatabase.OpenSQLRecord(gstrSQL, "zlBlobRead", Action, KeyWord, lngCount, IIf(blnMoved, 1, 0))
        If rsLob.EOF Then Exit Do
        If IsNull(rsLob.Fields(0).Value) Then Exit Do
        strText = rsLob.Fields(0).Value
        
        ReDim aryChunk(Len(strText) / 2 - 1) As Byte
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryChunk(lngBound) = CByte("&H" & Mid(strText, lngBound * 2 + 1, 2))
        Next
        
        Put lngFileNum, , aryChunk()
        lngCount = lngCount + 1
    Loop
    Close lngFileNum
    If lngCount = 0 Then Kill strFile: strFile = ""
    zlBlobRead = strFile
    Exit Function

ErrHand:
    Close lngFileNum
    Kill strFile: zlBlobRead = ""
    If ErrCenter = 1 Then
        Resume
    End If
End Function

'################################################################################################################
'## 功能：  在压缩文件相同目录释放产生解压文件
'## 参数：  strZipFile     :压缩文件
'## 返回：  解压文件名，失败则返回零长度""
'################################################################################################################
Public Function UnzipTendPage(ByVal strZipFile As String, ByVal strTarFile As String) As String
    Dim strZipPathTmp As String
    Dim strZipPath As String
    Dim strZipFileTmp As String
    Dim strZipFileName As String
    Dim mclsUnzip As New cUnzip
    
    On Error GoTo ErrHand
    
    If Not gobjFSO.FileExists(strZipFile) Then UnzipTendPage = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    
    strZipPath = gobjFSO.GetSpecialFolder(2)
    strZipPathTmp = strZipPath & Format(Now, "yyMMddHHmmss") & CStr(100 * Timer)
    Call gobjFSO.CreateFolder(strZipPathTmp)
    
    strZipFileTmp = strZipPathTmp ' & "\TMP.RTF"
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPathTmp
        .Unzip
    End With
    If gobjFSO.FolderExists(strZipFileTmp) Then
        
        strZipFileName = gobjFSO.GetFile(strZipFileTmp & "\" & strTarFile)
        Call gobjFSO.CopyFile(strZipFileName, "C:\" & strTarFile)
        
        On Error Resume Next
        gobjFSO.DeleteFolder strZipPathTmp, True
        gobjFSO.DeleteFile strZipFile, True
        
        UnzipTendPage = "C:\" & strTarFile
    Else
        UnzipTendPage = ""
    End If
ErrHand:
    Exit Function
End Function

Public Function GetTmpPath() As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strFileTemp As String
    Dim lngTemp As Long
    
    strFileTemp = Space(256)
    lngTemp = GetTempPath(256, strFileTemp)
    
    GetTmpPath = Mid(strFileTemp, 1, InStr(strFileTemp, Chr(0)) - 1)
End Function

Public Function PrintRTBData(ByVal objOutTo As Object, ByVal blnHead As Boolean, ByVal lngCurY As Long, Optional ByVal lngPage As Long = 0) As Boolean
    Dim fr As FORMATRANGE           '格式化的文本范围
    Dim rcDrawTo As RECT            '目标文字区域
    Dim rcPage As RECT              '目标页面区域
    Dim gTargetDC As Long
    Dim lngFoot As Long
    Dim lngOffsetLeft As Long
    Dim lngOffsetTop As Long
    Dim lngNextPos As Long, lngLen As Long, lngTmp As Long, lngPageCount As Long
    Dim rsTemp As New ADODB.Recordset
    Dim objRTB As RichTextBox
    
    If blnHead Then
        Set objRTB = rtbHead
    Else
        Set objRTB = rtbFoot
    End If
    
    lngLen = lstrlen(objRTB.Text)
    lngOffsetLeft = objOutTo.ScaleX(GetDeviceCaps(objOutTo.hDC, PHYSICALOFFSETX), vbPixels, vbTwips)
    lngOffsetTop = objOutTo.ScaleY(GetDeviceCaps(objOutTo.hDC, PHYSICALOFFSETY), vbPixels, vbTwips)
    
    'objOutTo.CurrentY = objOutTo.Height - objOutTo.ScaleX(Val(lngCurY * T_TwipsPerPixel.Y / 56.7), vbMillimeters, vbTwips)
    
    If blnHead Then
        objOutTo.Print ""
    Else
        'lngFoot = 180
        'objOutTo.CurrentX = (objOutTo.Width - 90 * LenB(StrConv("第--" & 5 & "--页", vbFromUnicode))) / 2
        'objOutTo.Print "第--" & 5 & "--页"
    End If
    
    gTargetDC = hDC
    With rcPage
        .Left = 0
        .Top = 0
        .Right = objOutTo.Width
        .Bottom = objOutTo.Height
    End With
    With rcDrawTo
        If blnHead Then
            .Left = lngOffsetLeft
            .Top = lngOffsetTop
            .Right = objOutTo.Width - lngOffsetLeft
            .Bottom = objOutTo.ScaleX(Val(lngCurY * T_TwipsPerPixel.Y), vbMillimeters, vbTwips) - 30
        Else
            .Left = lngOffsetLeft
            .Top = objOutTo.Height - objOutTo.ScaleX(lngCurY * T_TwipsPerPixel.Y / conRatemmToTwip, vbMillimeters, vbTwips)
            .Right = objOutTo.Width - lngOffsetLeft
            .Bottom = objOutTo.Height
        End If
    End With
    
    With fr
        .hDC = objOutTo.hDC
        .hdcTarget = gTargetDC
        .rc = rcDrawTo
        .rcPage = rcPage
        .chrg.cpMin = 0
        .chrg.cpMax = -1
    End With
    
    Do
        lngNextPos = SendMessage(objRTB.hWnd, EM_FORMATRANGE, 0, fr)
        
        lngPageCount = lngPageCount + 1             ' 页数＋1
        '记录分页信息
        ReDim Preserve AllPages(1 To lngPageCount) As PageInfo
        AllPages(lngPageCount).PageNumber = lngPageCount
        AllPages(lngPageCount).ActualHeight = fr.rc.Bottom - fr.rc.Top        '实际打印高度
        AllPages(lngPageCount).Start = lngTmp
        AllPages(lngPageCount).End = lngNextPos
        
        fr.chrg.cpMin = lngNextPos
        If lngNextPos <= lngTmp Or lngNextPos >= lngLen Then Exit Do      ' 完成所有页面的分页
        lngTmp = lngNextPos
    Loop
    
    Call SendMessage(objRTB.hWnd, EM_FORMATRANGE, 0, ByVal CLng(0))
    
    For lngLen = 1 To lngPageCount
        If lngLen > 1 Then Exit For
        With fr
            .hDC = objOutTo.hDC
            .hdcTarget = gTargetDC
            .rc = rcDrawTo
            .rcPage = rcPage
            .chrg.cpMin = AllPages(lngLen).Start
            .chrg.cpMax = AllPages(lngLen).End
        End With
        Call SendMessage(objRTB.hWnd, EM_FORMATRANGE, 1, fr)
        Call SendMessage(objRTB.hWnd, EM_FORMATRANGE, 0, ByVal CLng(0))
    Next
    
End Function

Private Sub Form_Load()
    Call InitRechBox(241)
End Sub
