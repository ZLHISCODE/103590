VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMsgCallSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "语音设置"
   ClientHeight    =   4245
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   10155
   Icon            =   "frmMsgCallSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   10155
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3075
      Left            =   105
      ScaleHeight     =   3075
      ScaleWidth      =   9405
      TabIndex        =   4
      Top             =   315
      Width           =   9405
      Begin VSFlex8Ctl.VSFlexGrid vsMsgList 
         Height          =   2670
         Left            =   180
         TabIndex        =   5
         Top             =   75
         Width           =   8865
         _cx             =   15637
         _cy             =   4710
         Appearance      =   2
         BorderStyle     =   0
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
         MousePointer    =   99
         MouseIcon       =   "frmMsgCallSetup.frx":06EA
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483643
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
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
         OwnerDraw       =   1
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   -2147483643
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.Frame fraBell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   225
      Begin VB.Image imgBell 
         Height          =   240
         Left            =   0
         Picture         =   "frmMsgCallSetup.frx":0FC4
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   300
      Left            =   8430
      TabIndex        =   1
      Top             =   3540
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   300
      Left            =   6945
      TabIndex        =   0
      Top             =   3540
      Width           =   1000
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   5010
      Top             =   -270
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgFile 
      Height          =   240
      Left            =   1500
      Picture         =   "frmMsgCallSetup.frx":7816
      Top             =   105
      Width           =   240
   End
   Begin VB.Label lblW 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1710
      TabIndex        =   6
      Top             =   60
      Width           =   90
   End
   Begin VB.Label lblDetail 
      Caption         =   "文本格式说明：可用[床号]或[住院号]字段，使用与VBScript兼容的表达式对消息内容进行编辑；字段项请使用方括符""[]""括起表示。"
      Height          =   390
      Left            =   -45
      TabIndex        =   3
      Top             =   3690
      Width           =   6585
   End
End
Attribute VB_Name = "frmMsgCallSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum Col
    COL_声音类型 = 1
    COL_试听
    COL_状态
    COL_提示方式
    COL_播报内容
    COL_播报次数
End Enum

Private mobjVBA As Object
Private mobjScript As clsScript
Private mobjVoice As Object
Private mstr消息行 As String ' = "新开消息,新停消息,新废消息,安排消息,危机值消息,输液拒绝消息,销帐申请消息"
Private mint场合 As Integer '0-门诊医生工作站，1－住院医生工作站，2－住院护士工作站，3－老版医技工作站
Private mrsPars As ADODB.Recordset

Public Function ShowMe(objFrm As Object, ByVal intType As Integer) As Boolean
    mint场合 = intType
    Me.Show 1, objFrm
End Function

Private Sub InitMsgListTable()
'功能：初始化表格内容，用在窗体个性化设置恢复之前
    Dim strHead As String, i As Integer
    Dim arrHead As Variant, arrCol As Variant
 
    strHead = "声音类型,1200,1;试听,540,4;状态,540,4;提示方式,800,4;播报内容,5000,1;播报次数,900,4"
    arrHead = Split(strHead, ";")
    With vsMsgList
        .Clear
        .FixedRows = 1: .FixedCols = 1
        .Rows = 2: .Cols = .FixedCols + UBound(arrHead) + 1
        For i = 0 To UBound(arrHead)
            .FixedAlignment(.FixedCols + i) = 4
            arrCol = Split(arrHead(i), ",")
            .TextMatrix(0, .FixedCols + i) = arrCol(0)
             .Cell(flexcpText, 0, .FixedCols + i) = arrCol(0)
            
            If UBound(arrCol) > 0 Then
                .ColWidth(.FixedCols + i) = Val(arrCol(1))
                .ColAlignment(.FixedCols + i) = Val(arrCol(2))
                .ColHidden(.FixedCols + i) = False
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .ColWidth(0) = 0
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    
    On Error GoTo errH
    
    With vsMsgList
        .Redraw = flexRDNone
        For i = .FixedRows To .Rows - 1
            If Not ChangePars(i) Then
                .Redraw = flexRDDirect
                Exit Sub
            End If
        Next
    End With
    
    mrsPars.Filter = "修改=1"
    If Not mrsPars.EOF Then
        For i = 1 To mrsPars.RecordCount
            Call zlDatabase.SetPara(mrsPars!参数名 & "", mrsPars!现参数值 & "", glngSys, Val(mrsPars!模块 & ""))
            mrsPars.MoveNext
        Next
    End If
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim varTmp As Variant
    Dim strTmp As String
    Dim i As Long, lng模块 As Long
    Dim objMsg As New clsCISMsg
    
    Call InitMsgListTable
    
    Call objMsg.InitRsMsgPar(mrsPars)
    
    If mint场合 = 0 Then
        lng模块 = p门诊医生站
    ElseIf mint场合 = 1 Then
        lng模块 = p住院医生站
    ElseIf mint场合 = 2 Then
        lng模块 = p住院护士站
    ElseIf mint场合 = 3 Then
        lng模块 = p医技工作站
    End If
    
    strTmp = objMsg.Get消息类别(mint场合)
    varTmp = Split(strTmp, ",")
    For i = 0 To UBound(varTmp)
        Call objMsg.AddDataToRsMsgPar(mrsPars, lng模块, i + 1, varTmp(i) & "语音配置", varTmp(i))
    Next
    
    mrsPars.Filter = 0
    mrsPars.Sort = "序号"
    mrsPars.MoveFirst
    With vsMsgList
        .Rows = UBound(varTmp) + 2
        For i = 1 To mrsPars.RecordCount
            .TextMatrix(i, COL_声音类型) = mrsPars!声音类型
            .TextMatrix(i, COL_状态) = IIf(1 = Val(mrsPars!状态 & ""), "开启", "关闭")
            .TextMatrix(i, COL_提示方式) = IIf(1 = Val(mrsPars!提示方式 & ""), "提示", "朗读")
            .TextMatrix(i, COL_播报内容) = mrsPars!内容 & ""
            .TextMatrix(i, COL_播报次数) = Val(mrsPars!次数 & "")
            Set .Cell(flexcpPicture, i, COL_试听) = imgBell.Picture
            .Cell(flexcpPictureAlignment, i, COL_试听) = 4
            mrsPars.MoveNext
        Next
    End With
    
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    Me.Height = 3700
    picMain.BackColor = &H8000000A
    picMain.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 700
    lblW.Visible = False
    vsMsgList.Move 0, 0, picMain.ScaleWidth - 30, picMain.ScaleHeight - 30
    
    cmdCancel.Top = Me.ScaleHeight - 500
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 100
    
    cmdOK.Top = cmdCancel.Top
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 100
    
    lblDetail.Top = cmdCancel.Top
    lblDetail.Left = picMain.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjVBA = Nothing
    Set mobjScript = Nothing
    Set mobjVoice = Nothing
    Set mrsPars = Nothing
End Sub

Private Sub imgBell_Click()
'功能：播放
    Dim str内容 As String
    Dim strTmp As String
    Dim lngRow As Long
    Dim objMsg As New clsCISMsg

    With vsMsgList
        If Not .Col = COL_试听 Then Exit Sub
        lngRow = Val(fraBell.Tag)
        str内容 = .TextMatrix(lngRow, COL_播报内容)
        If .TextMatrix(lngRow, COL_提示方式) = "提示" Then
            If str内容 = "" Then
                MsgBox "请选择一个音频文件(*.wav)。", vbInformation, gstrSysName
                Exit Sub
            Else
                If UCase(Right(str内容, 4)) <> ".WAV" Then
                    MsgBox "请选择一个正确格式的音频文件(*.wav)。", vbInformation, gstrSysName
                    Exit Sub
                ElseIf Not Dir(str内容) <> "" Then
                    MsgBox "未找文件[" & str内容 & "]，请检查。", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            Call sndPlaySound(str内容, 1)
        Else
            If str内容 = "" Then
                MsgBox "请定义一段朗读文本。", vbInformation, gstrSysName
            End If
            
            If mobjVoice Is Nothing Then
                Set mobjVoice = CreateObject("SAPI.SpVoice")
                Call objMsg.CreateScript(mobjVBA, mobjScript)
            End If
            
            strTmp = Check文本(str内容)
            If strTmp <> "" Then
                MsgBox .TextMatrix(lngRow, COL_声音类型) & "文本格式检查未通过：【" & strTmp & "】", vbInformation, gstrSysName
                Exit Sub
            End If
                
            str内容 = Get播放文本(str内容)
            mobjVoice.Speak str内容, 1
        End If
    End With
End Sub

Private Sub vsMsgList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = -1 Then Exit Sub
    With vsMsgList
        If NewCol = COL_播报内容 Then
            .Editable = flexEDKbdMouse
            If .TextMatrix(NewRow, COL_提示方式) = "提示" Then
                .ComboList = "..."
                Set .CellButtonPicture = imgFile.Picture
            Else
                .ComboList = ""
            End If
            .FocusRect = flexFocusLight
        ElseIf NewCol = COL_播报次数 Then
            .Editable = flexEDKbdMouse
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .Editable = flexEDNone
            .FocusRect = flexFocusNone
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vsMsgList_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
'功能：编辑表格
    Dim strFileDir As String
    On Error GoTo errH
    If Col = COL_播报内容 And vsMsgList.TextMatrix(Row, COL_提示方式) = "提示" Then
        vsMsgList.Redraw = flexRDNone
        With dlgFile
            .CancelError = False
            .Flags = cdlOFNHideReadOnly
            .Filter = "(*.wav)|*.wav"
            .FilterIndex = 2
            .ShowOpen
            strFileDir = .FileName
            If strFileDir = "" Then Exit Sub
        End With
        vsMsgList.TextMatrix(Row, COL_播报内容) = strFileDir
        vsMsgList.Redraw = flexRDDirect
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
 
Private Sub vsMsgList_Click()
    vsMsgList.Redraw = flexRDNone
    Call imgBell_Click
    vsMsgList.Redraw = flexRDDirect
End Sub

Private Sub vsMsgList_DblClick()
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strTmp As String
    
    With vsMsgList
        lngRow = .Row
        lngCol = .Col
        If lngRow >= .FixedRows Then
            If lngCol = COL_提示方式 Then
                If .TextMatrix(lngRow, lngCol) = "提示" Then
                    .TextMatrix(lngRow, lngCol) = "朗读"
                Else
                    .TextMatrix(lngRow, lngCol) = "提示"
                End If
                .TextMatrix(lngRow, COL_播报内容) = ""
            ElseIf lngCol = COL_状态 Then
                If .TextMatrix(lngRow, lngCol) = "开启" Then
                    .TextMatrix(lngRow, lngCol) = "关闭"
                Else
                    .TextMatrix(lngRow, lngCol) = "开启"
                End If
            End If
        End If
    End With
End Sub

Private Sub vsMsgList_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'说明：1.OwnerDraw要设置为Over(画出单元所有内容)
'      2.Cell的GridLine从上下左右向内都是从第1根线开始
'      3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim lngColor As Long, lngclrg, k As Long, n As Long
    Dim r1 As Integer, g1 As Integer, b1 As Integer
    Dim r2 As Integer, g2 As Integer, b2 As Integer
    Dim rg As Integer, gg As Integer, bg As Integer
    
    Dim lngFontW As Long
    Dim lng组ID As Long
    Dim strContent As String
    
    Dim vRect As RECT, vRect1 As RECT, vRect2 As RECT
    With vsMsgList
        '分类行背景色处理
        If .FixedRows - 1 = Row Then
            '获取矩形框
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Right - 1
            vRect.Bottom = Bottom - 1
            'draw frame
            lngColor = SetBkColor(hDC, 0)
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, k
    
            ' get colors
            r1 = 250: g1 = 250: b1 = 250   '渐变起始
            r2 = 229: g2 = 229: b2 = 229   '渐变终止
            ' show color
            vRect2 = vRect
            vRect2.Bottom = vRect.Bottom - (vRect.Bottom - vRect.Top) / 2
            vRect1 = vRect2
    
            For k = vRect2.Top To vRect2.Bottom
                rg = r1 + (k - vRect2.Top) * (r2 - r1) / (vRect2.Bottom - vRect2.Top)
                gg = g1 + (k - vRect2.Top) * (g2 - g1) / (vRect2.Bottom - vRect2.Top)
                bg = b1 + (k - vRect2.Top) * (b2 - b1) / (vRect2.Bottom - vRect2.Top)
                lngclrg = RGB(rg, gg, bg)
                SetBkColor hDC, lngclrg
                vRect1.Top = k
                ExtTextOut hDC, vRect1.Left, vRect1.Top, ETO_OPAQUE, vRect1, " ", 1, k
            Next
            ' get colors
            r1 = 229: g1 = 229: b1 = 229   '渐变起始
            r2 = 250: g2 = 250: b2 = 250   '渐变终止
            ' show color
            vRect2 = vRect
            vRect2.Top = vRect.Top + (vRect.Bottom - vRect.Top) / 2
            vRect1 = vRect2
            For k = vRect2.Top To vRect2.Bottom
                rg = r1 + (k - vRect2.Top) * (r2 - r1) / (vRect2.Bottom - vRect2.Top)
                gg = g1 + (k - vRect2.Top) * (g2 - g1) / (vRect2.Bottom - vRect2.Top)
                bg = b1 + (k - vRect2.Top) * (b2 - b1) / (vRect2.Bottom - vRect2.Top)
                lngclrg = RGB(rg, gg, bg)
                SetBkColor hDC, lngclrg
                vRect1.Top = k
                ExtTextOut hDC, vRect1.Left, vRect1.Top, ETO_OPAQUE, vRect1, " ", 1, k
            Next
    
            SetBkColor hDC, lngColor
            '将单元格字体绘到矩形区域
            strContent = .Cell(flexcpText, Row, Col)
            lblW.Caption = strContent: lblW.AutoSize = True
            vRect1.Top = vRect.Top + (vRect.Bottom - vRect.Top) / 2 - (lblW.Height / 2) / Screen.TwipsPerPixelY

            vRect1.Left = vRect.Left + (vRect.Right - vRect.Left) / 2 - (lblW.Width / 2) / Screen.TwipsPerPixelX
    
            TextOut hDC, vRect1.Left, vRect1.Top, strContent, LenB(StrConv(strContent, vbFromUnicode))
        End If
    End With
End Sub

Private Sub vsMsgList_KeyPress(KeyAscii As Integer)
    With vsMsgList
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call EnterNextCell(.Row, .Col)
        Else
             If .Col = COL_播报内容 Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsMsgList_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" '使按钮状态进入输入状态
                End If
            End If
        End If
    End With
End Sub

Private Sub vsMsgList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call EnterNextCell(Row, Col)
        Exit Sub
    End If
    If Col = COL_播报次数 Then
        If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0: Exit Sub
        End If
    ElseIf Col = COL_播报内容 And vsMsgList.TextMatrix(Row, COL_提示方式) = "提示" Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub vsMsgList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long
    Dim strTag As String
    Dim lngCol As Long
 
    lngRow = vsMsgList.MouseRow
    If Button = 0 And lngRow > 0 Then
        If vsMsgList.MouseCol = COL_试听 Then
            If Val(fraBell.Tag) <> lngRow Then
                With vsMsgList
                    fraBell.Visible = False
                    fraBell.Tag = lngRow
                    If lngRow = .Row Then
                        fraBell.BackColor = .BackColorSel
                    Else
                        fraBell.BackColor = .BackColor
                    End If
                    fraBell.Height = .RowHeight(lngRow) - 20
                    If fraBell.Height > 250 Then fraBell.Height = 250
                    
                    fraBell.Top = .Top + .RowPos(lngRow) + .RowHeight(lngRow) - fraBell.Height + 10
                    If fraBell.Top + fraBell.Height > .Top + .Height Then Exit Sub
                    
                    fraBell.Left = .Left + .ColPos(COL_试听) + (.ColWidth(COL_试听) - fraBell.Width) / 2
                    
                    fraBell.Visible = True
                End With
            End If
        Else
            If fraBell.Visible Then
                fraBell.Tag = ""
                fraBell.Visible = False
            End If
        End If
    End If
    
    With vsMsgList
        lngRow = .MouseRow
        lngCol = .MouseCol
        strTag = lngRow & "<T>" & lngCol
        If lngRow > 0 And .Tag <> strTag Then
            If lngCol = COL_提示方式 Or lngCol = COL_状态 Then
                Call ClearListSel
                .Cell(flexcpForeColor, lngRow, lngCol) = .ForeColorSel
                .Cell(flexcpBackColor, lngRow, lngCol) = vbWhite
                .CellBorderRange lngRow, lngCol, lngRow, lngCol, 1, 1, 1, 1, 1, 1, 1
                .Tag = strTag
                .ToolTipText = .Cell(flexcpText, lngRow, lngCol)
            Else
                Call ClearListSel
            End If
        End If
        If lngRow > 0 And lngCol > 0 Then .ToolTipText = .Cell(flexcpText, lngRow, lngCol)
        
        If lngCol = COL_试听 Then
            .MousePointer = 99
        Else
            .MousePointer = 0
        End If
    End With
End Sub

Private Sub ClearListSel()
    Dim lngCol As Long
    Dim lngRow As Long
    
    On Error Resume Next
    With vsMsgList
        If .Tag <> "" Then
            lngRow = Split(.Tag, "<T>")(0)
            lngCol = Split(.Tag, "<T>")(1)
            .Row = .FixedRows - 1
            .Cell(flexcpForeColor, lngRow, lngCol) = .ForeColor
            .Cell(flexcpBackColor, lngRow, lngCol) = .BackColor
            .CellBorderRange lngRow, lngCol, lngRow, lngCol, 0, 0, 0, 0, 0, 0, 0
            .ToolTipText = ""
            .Tag = ""
        End If
    End With
End Sub

Private Sub EnterNextCell(ByVal lngRow As Long, ByVal lngCol As Long)
'功能：定位到下一个可以输入的单元格
    Dim i As Long, j As Long
    Dim blnDo As Boolean
    
    With vsMsgList
        '从下一单元开始循环搜索
        For i = lngRow To .Rows - 1
            For j = IIf(i = lngRow, lngCol + 1, COL_提示方式) To .Cols - 1
                If j >= COL_播报内容 Then
                    Exit For
                End If
            Next
            If j <= .Cols - 1 Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
            blnDo = True
        End If
        .ShowCell .Row, .Col
        If Not blnDo Then Call zlCommFun.PressKey(vbKeyTab)
    End With
End Sub

Private Function Check文本(ByVal strText As String) As String
'功能：检查医嘱内容是否正确
'返回：错误信息
'      strPreview=预览医嘱内容效果
    Dim intLeft As Integer, intRight As Integer
    Dim strTmp As String, strPar As String
    Dim strMsg As String, i As Long
    Dim objVBA As Object, strEval As String
    Dim objScript As New clsScript
    
    If Trim(strText) = "" Then Exit Function
    If zlCommFun.ActualLen(strText) > 100 Then
        strMsg = "文本定义内容太长，只允许100个字符或50个汉字。"
        GoTo EndLine
    End If
        
    '检查配对情况
    For i = 1 To Len(strText)
        If Mid(strText, i, 1) = "[" Then
            intLeft = intLeft + 1
        ElseIf Mid(strText, i, 1) = "]" Then
            intRight = intRight + 1
            If intLeft <> intRight Then
                strMsg = """[""与""]""括号不配对。"
                GoTo EndLine
            End If
        End If
    Next
    If intLeft = 0 And intRight = 0 Then Exit Function
    If intLeft <> intRight Then
        strMsg = """[""与""]""括号不配对。"
        GoTo EndLine
    End If
    
    '检查字段名称
    strTmp = strText
    Do While InStr(strTmp, "[") > 0
        strTmp = Mid(strTmp, InStr(strTmp, "[") + 1)
        strPar = Trim(Left(strTmp, InStr(strTmp, "]") - 1))
                        
        If strPar = "" Then
            strMsg = """[]""括号之中没有书写字段名。"
            GoTo EndLine
        End If
        If "[床号]" <> "[" & strPar & "]" And "[住院号]" = "[" & strPar & "]" Then
            strMsg = "使用了不存在的""[" & strPar & "]""字段。"
            GoTo EndLine
        End If
    Loop
    
    '执行测试
    On Error Resume Next
    Set objVBA = CreateObject("ScriptControl")
    If objVBA Is Nothing Then
        strMsg = "Microsoft Script Control 未正确安装(msscript.ocx)，不能执行检查。请重新安装客户端程序。"
        GoTo EndLine
    End If
    err.Clear: On Error GoTo 0
    objVBA.Language = "VBScript"
    objVBA.AddObject "clsScript", objScript, True
    strEval = Replace(strText, "[", """")
    strEval = Replace(strEval, "]", """")
    On Error Resume Next
    Call objVBA.Eval(strEval)
    If objVBA.Error.Number <> 0 Then
        strMsg = objVBA.Error.Description
        objVBA.Error.Clear
    End If
EndLine:
    Check文本 = strMsg
End Function

Private Function CheckFile(ByVal strFile As String) As String
'功能：音频文件检查
    Dim strMsg As String
    
    If UCase(Right(strFile, 4)) <> ".WAV" Then
        strMsg = "请选择一个正确格式的音频文件(*.wav)。"
        MsgBox "请选择一个正确格式的音频文件(*.wav)。", vbInformation, gstrSysName
    ElseIf Not Dir(strFile) <> "" Then
        strMsg = "未找文件[" & strFile & "]，请检查。"
    End If
End Function

Private Function Get播放文本(ByVal strText As String) As String
'功能：获取播入的文本
    Dim str床号 As String, str住院号 As String
    Dim strVal As String
    
    str床号 = "1床"
    str住院号 = "201608010008号"
    strVal = strText
    strVal = Replace(strVal, "[床号]", """" & str床号 & """")
    strVal = Replace(strVal, "[住院号]", """" & str住院号 & """")
    
    On Error Resume Next
    strVal = mobjVBA.Eval(strVal)
    If mobjVBA.Error.Number <> 0 Then
        err.Clear
        strVal = ""
    End If
    Get播放文本 = strVal
End Function

Private Function ChangePars(ByVal lngRow As Long) As Boolean
'功能：参数变换
    Dim strTmp As String
    
    With vsMsgList
        If Not IsNumeric(.TextMatrix(lngRow, COL_播报次数)) Then
            MsgBox .TextMatrix(lngRow, COL_声音类型) & "播报次数必须为数字。", vbInformation, gstrSysName
            Exit Function
        End If
        
        If Val(.TextMatrix(lngRow, COL_播报次数)) > 6 Then
            MsgBox .TextMatrix(lngRow, COL_声音类型) & "播报次数最多只能设为6次。", vbInformation, gstrSysName
            Exit Function
        End If
        
        If .TextMatrix(lngRow, COL_提示方式) = "提示" Then
            If .TextMatrix(lngRow, COL_播报内容) = "" Then
                MsgBox .TextMatrix(lngRow, COL_声音类型) & "未设置提示音频文件。", vbInformation, gstrSysName
                Exit Function
            Else
                strTmp = CheckFile(.TextMatrix(lngRow, COL_播报内容))
                If strTmp <> "" Then
                    MsgBox .TextMatrix(lngRow, COL_声音类型) & "文件设置：【" & strTmp & "】", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Else
            If .TextMatrix(lngRow, COL_播报内容) = "" Then
                MsgBox .TextMatrix(lngRow, COL_声音类型) & "未设置提文本格式。", vbInformation, gstrSysName
                Exit Function
            Else
                strTmp = Check文本(.TextMatrix(lngRow, COL_播报内容))
                If strTmp <> "" Then
                    MsgBox .TextMatrix(lngRow, COL_声音类型) & "文本格式检查未通过：【" & strTmp & "】", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        
        mrsPars.Filter = "声音类型='" & .TextMatrix(lngRow, COL_声音类型) & "'"
        mrsPars!状态 = IIf("开启" = .TextMatrix(lngRow, COL_状态), 1, 0)
        mrsPars!提示方式 = IIf("提示" = .TextMatrix(lngRow, COL_提示方式), 1, 0)
        mrsPars!内容 = .TextMatrix(lngRow, COL_播报内容)
        mrsPars!次数 = Val(.TextMatrix(lngRow, COL_播报次数))
        mrsPars!现参数值 = mrsPars!状态 & "<sTab>" & mrsPars!提示方式 & "<sTab>" & mrsPars!内容 & "<sTab>" & mrsPars!次数
        
        If mrsPars!现参数值 <> mrsPars!原参数值 Then
            mrsPars!修改 = 1
        End If
        
        mrsPars.Update
    End With
    ChangePars = True
End Function
