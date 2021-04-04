VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmDiagReference 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "诊断参考"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13695
   Icon            =   "frmDiagReference.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   13695
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timer 
      Interval        =   50
      Left            =   5760
      Top             =   120
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   3960
      ScaleHeight     =   8175
      ScaleWidth      =   9375
      TabIndex        =   5
      Top             =   600
      Width           =   9375
      Begin SHDocVwCtl.WebBrowser webRpt 
         Height          =   4335
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   8055
         ExtentX         =   14208
         ExtentY         =   7646
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8F2EC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9135
      Left            =   0
      ScaleHeight     =   9135
      ScaleWidth      =   3615
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VSFlex8Ctl.VSFlexGrid vsDiagList 
         Height          =   2535
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   3375
         _cx             =   5953
         _cy             =   4471
         Appearance      =   3
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         MouseIcon       =   "frmDiagReference.frx":6852
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16369772
         ForeColorSel    =   -2147483641
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483637
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   500
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmDiagReference.frx":712C
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   0   'False
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsDiagList 
         Height          =   2175
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   3840
         Width           =   3255
         _cx             =   5741
         _cy             =   3836
         Appearance      =   3
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         MouseIcon       =   "frmDiagReference.frx":717B
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16369772
         ForeColorSel    =   -2147483641
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483643
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   500
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmDiagReference.frx":7A55
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   0   'False
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "鉴别诊断列表"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   3480
         Width           =   3240
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参考诊断列表"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   3360
      End
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参考内容"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   2
      Left            =   4080
      TabIndex        =   7
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmDiagReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjJson As Object '格式:[{"DIA_NAME":"异位妊娠","DIA_ID":"22e87bb5-8691-4cf1-b367-b3e39e47616c","DIA_NAME_QUERY":"妊娠","MATCH_RATE":0.5},{"DIA_NAME":"卵巢妊娠","DIA_ID":"e9c9a22e-2ae2-4509-aee2-bba45a361601","DIA_NAME_QUERY":"妊娠","MATCH_RATE":0.5}]
Private mstrAuthentication As String
Private mstrTempDel As String
Private mstrDiaID  As String   '记录参考内容避免重复刷新

Private Enum E_COL
    COL_名称
    COL_匹配度
End Enum

Public Function ShowMe(objfrmMain As Object, ByVal bytStyle As Byte, ByVal objJSON As Object) As Boolean
    Set mobjJson = objJSON
    mstrDiaID = ""
    Me.Show bytStyle, objfrmMain
End Function

Private Sub Form_Load()
    Dim strHead As String
    
    mstrAuthentication = "Basic " & zlStr.Base64Encode("xxx:xxx")
    strHead = "诊断名称,3000,1;匹配度,1000,4"
    Call Grid.Init(vsDiagList(0), strHead)
    vsDiagList(0).FixedAlignment(0) = flexAlignLeftCenter
    
    strHead = "诊断名称,3500,1"
    Call Grid.Init(vsDiagList(1), strHead)
    vsDiagList(1).FixedAlignment(0) = flexAlignLeftCenter
    Call LoadDiag(0, mobjJson)
    Call vsDiagList_Click(0)
    vsDiagList(0).BackColorSel = vsDiagList(0).BackColorBkg
    vsDiagList(1).BackColorSel = vsDiagList(1).BackColorBkg
    Me.BackColor = &H80000005
End Sub

Private Sub Form_Resize()
    Dim lngLeft As Long
    Dim lngSplit As Long
    
    On Error Resume Next
    lngLeft = 4000
    lngSplit = 120
    
    picLeft.Move 0, 0, lngLeft, Me.ScaleHeight
    lblInfo(2).Move lngLeft + lngSplit, lngSplit
    picMain.Move lngLeft, lblInfo(2).Top + lblInfo(2).Height + lngSplit, Me.ScaleWidth - lngLeft, Me.ScaleHeight - (lblInfo(2).Top + lblInfo(2).Height + lngSplit)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    webRpt.Navigate "xxx"
'    Call DeleteTempFile
    mstrDiaID = ""
End Sub

Private Sub picLeft_Resize()
    On Error Resume Next
    lblInfo(0).Move 120, 120, picLeft.ScaleWidth - 240
    lblInfo(1).Move 120, picLeft.Height / 2, picLeft.ScaleWidth - 240
    vsDiagList(0).Move 120, lblInfo(0).Top + lblInfo(0).Height + 120, picLeft.ScaleWidth - 240, picLeft.Height / 2 - (lblInfo(0).Top + lblInfo(0).Height + 240)
    vsDiagList(1).Move 120, lblInfo(1).Top + lblInfo(1).Height + 120, picLeft.ScaleWidth - 240, picLeft.Height - (lblInfo(1).Top + lblInfo(1).Height + 240)
End Sub

Private Sub LoadDiag(ByVal Index As Integer, ByVal objJSON As Object)
'---------------------------------------------------------------------------------------
' Procedure : LoadDiag
' Author    : YWJ
' Date      : 2018/11/2
' Purpose   :根据诊断信息加载
'---------------------------------------------------------------------------------------
    Dim i As Long
    
    With vsDiagList(Index)
        If Index = 0 Then
            .rows = objJSON.Count + 1
             
            For i = .FixedRows To objJSON.Count
                .TextMatrix(i, COL_名称) = objJSON.Item(i).Item("DIA_NAME")
                .RowData(i) = objJSON.Item(i).Item("DIA_ID") & ""
                .TextMatrix(i, COL_匹配度) = Format(objJSON.Item(i).Item("MATCH_RATE"), "0.00%")
            Next
        Else
            .rows = objJSON.Count + 1
            For i = .FixedRows To objJSON.Count
                .TextMatrix(i, COL_名称) = objJSON.Item(i).Item("DIFF_DIA_NAME")
                .RowData(i) = objJSON.Item(i).Item("DIFF_DIA_ID") & ""
            Next
        End If
        If i >= .FixedRows Then .Row = .FixedRows
    End With
End Sub

Private Sub picMain_Resize()
    On Error Resume Next
    webRpt.Move 0, 0, picMain.Width, picMain.Height
End Sub

Private Function ReadDiag(ByVal strDiaID As String) As Boolean
    Dim objJSON As Object
    Dim strInput As String
    Dim strRet As String
    Dim strMsg As String
    
    With vsDiagList(1)
        .rows = 1 '清空数据保留固定行
        .rows = 2 '设置默认行
        '[{"name": "input_in","value": "{\"DIA_ID\": \"6440b8e9-7e73-4779-b6f4-4db74539e64f\"}"}]
        Set objJSON = mdlJSON.parse("{}")
        Call objJSON.Add("DIA_ID", strDiaID)
        strInput = mdlJSON.toString(objJSON)
        Set objJSON = mdlJSON.parse("{}")
        Call objJSON.Add("name", "input_in")
        Call objJSON.Add("value", strInput)
        strInput = "[" & mdlJSON.toString(objJSON) & "]"
        WriteLog Me.Name, "ReadDiag", "【鉴别诊断查询】入参 URL:" & gstrAntidiastoleURL & "  传入值:" & strInput
        strRet = HttpPost(gstrAntidiastoleURL, strInput, responseText, , mstrAuthentication)
        'strRet :{"output_out":null}
        WriteLog Me.Name, "ReadDiag", "【鉴别诊断查询】返回值:" & strRet
        If strRet = "" Then
            strMsg = "未找到相应的鉴别诊断。"
            GoTo errMsg:
        End If
        Set objJSON = Nothing
        Set objJSON = mdlJSON.parse(strRet)
        If objJSON Is Nothing Then
           strMsg = "【鉴别诊断查询】返回值解析失败！" & mdlJSON.GetParserErrors()
           GoTo errMsg:
        End If
        strRet = NVL(objJSON.Item("output_out"))
        If strRet = "" Then strMsg = "未找到相应的鉴别诊断。": GoTo errMsg:
        Set objJSON = parse(strRet)
        If objJSON Is Nothing Then strMsg = "【鉴别诊断查询】返回值解析失败！" & mdlJSON.GetParserErrors(): GoTo errMsg:
    End With
    Call LoadDiag(1, objJSON)
    ReadDiag = True
    Exit Function
errMsg:
    vsDiagList(1).TextMatrix(1, COL_名称) = strMsg
    vsDiagList(1).Cell(flexcpForeColor, 1, COL_名称) = vbRed
    Exit Function
End Function


Private Function ReadContent(ByVal strDiaID As String) As Boolean
    Dim objJSON As Object
    Dim strInput As String
    Dim strRet As String
    Dim strFile As String
    
    On Error GoTo errH
    
    Me.MousePointer = 11
    '[{"name": "input_in","value": "{\"DIA_ID\": \"b7c49925-4b39-406d-82c6-5f630b400371\"}"}]
    Set objJSON = mdlJSON.parse("{}")
    Call objJSON.Add("DIA_ID", strDiaID)
    strInput = mdlJSON.toString(objJSON)
    Set objJSON = mdlJSON.parse("{}")
    Call objJSON.Add("name", "input_in")
    Call objJSON.Add("value", strInput)
    strInput = "[" & mdlJSON.toString(objJSON) & "]"
    WriteLog Me.Name, "ReadContent", "【诊断文档查询】入参 URL:" & gstrDiagContentURL & "  传入值:" & strInput
    strRet = HttpPost(gstrDiagContentURL, strInput, responseText, , mstrAuthentication)  'BASE64
    'strRet={"output_out":"ADFDASf"}
    WriteLog Me.Name, "ReadContent", "【诊断文档查询】返回值:" & strRet
    If strRet <> "" Then
        Set objJSON = parse(strRet)
        strRet = NVL(objJSON.Item("output_out"))
    End If
    If strRet = "" Then
        webRpt.Navigate "about:blank"
        webRpt.Document.Write "未找到相应的诊断文档。"
        webRpt.Refresh
        mstrTempDel = webRpt.Tag
        webRpt.Tag = ""             '记录待删除文件
    Else
        strFile = zlStr.DecodeBase64_File(strRet)
        webRpt.Navigate strFile
        mstrTempDel = webRpt.Tag
        webRpt.Tag = strFile       '记录待删除文件
    End If
    ReadContent = True
    Me.MousePointer = 0
    Exit Function
errH:
    Me.MousePointer = 0
    WriteLog Me.Name, "ReadContent", "【诊断文档查询】错误号:" & Err.Number & "错误描述：" & Err.Description
End Function

Private Function DeleteTempFile() As Boolean
    Dim objFile As New FileSystemObject
    Dim i As Long
    
    If mstrTempDel = "" Then Exit Function
    If objFile.FileExists(mstrTempDel) Then
        Do While i < 1000
            On Error Resume Next
            objFile.DeleteFile mstrTempDel, True
            If Err.Number = 0 Then
                mstrTempDel = ""
                Exit Do
            End If
            Err.Clear: On Error GoTo 0
        Loop
    End If
End Function

Private Sub Timer_Timer()
    If mstrTempDel <> "" Then
        Call DeleteTempFile
    End If
End Sub

Private Sub vsDiagList_Click(Index As Integer)
    Dim strDiaID As String
    With vsDiagList(Index)
        If .Row < .FixedRows Then Exit Sub
        strDiaID = CStr(.RowData(.Row) & "")
        If mstrDiaID = strDiaID And strDiaID <> "" Then Exit Sub
        lblInfo(2).Caption = IIf(Index = 0, "参考诊断：", "鉴别诊断：") & "【" & .TextMatrix(.Row, COL_名称) & "】"
        lblInfo(2).ForeColor = vbBlack
        If Index = 0 Then
            Call ReadDiag(strDiaID)
            Call ReadContent(strDiaID)
        Else
            Call ReadContent(strDiaID)
        End If
        mstrDiaID = strDiaID
    End With
End Sub

Private Sub vsDiagList_GotFocus(Index As Integer)
    vsDiagList(Index).BackColorSel = &HF9C86C
End Sub

Private Sub vsDiagList_LostFocus(Index As Integer)
    vsDiagList(Index).BackColorSel = vsDiagList(Index).BackColorBkg
End Sub
