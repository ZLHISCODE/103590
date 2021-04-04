VERSION 5.00
Begin VB.Form frmCaseTendBodySetLine 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "体温图线数据编辑"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   Icon            =   "frmCaseTendBodySetLine.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "说明"
      Height          =   1485
      Left            =   165
      TabIndex        =   17
      Top             =   2400
      Width           =   5115
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1155
         TabIndex        =   4
         Top             =   645
         Width           =   3540
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1155
         TabIndex        =   2
         Top             =   255
         Width           =   3540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "说明（下标）"
         Height          =   180
         Left            =   105
         TabIndex        =   3
         Top             =   690
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "说明（上标）"
         Height          =   180
         Left            =   105
         TabIndex        =   1
         Top             =   315
         Width           =   1080
      End
   End
   Begin zl9BodyEditorYDEY.VsfGrid vsf 
      Height          =   1950
      Left            =   165
      TabIndex        =   0
      Top             =   405
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   3440
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   5355
      TabIndex        =   7
      Top             =   3630
      Width           =   1100
   End
   Begin VB.CheckBox chkContinue 
      Caption         =   "时间递增连续输入(&N)"
      Height          =   210
      Left            =   4680
      TabIndex        =   8
      Top             =   150
      Width           =   2010
   End
   Begin VB.CommandButton cmdCanc 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5355
      TabIndex        =   6
      Top             =   855
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5355
      TabIndex        =   5
      Top             =   420
      Width           =   1100
   End
   Begin VB.Frame fraTop 
      Height          =   2415
      Left            =   6780
      TabIndex        =   10
      Top             =   1590
      Width           =   3750
      Begin VB.TextBox txtItem 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   1305
         MaxLength       =   10
         TabIndex        =   11
         Top             =   195
         Width           =   1320
      End
      Begin VB.ComboBox cboComment 
         Height          =   300
         Left            =   675
         TabIndex        =   14
         Top             =   1980
         Width           =   2610
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "收缩压"
         Height          =   180
         Index           =   0
         Left            =   255
         TabIndex        =   13
         Top             =   255
         Width           =   540
      End
      Begin VB.Label lblComment 
         Caption         =   "说明"
         Height          =   180
         Left            =   255
         TabIndex        =   12
         Top             =   2040
         Width           =   540
      End
   End
   Begin VB.Label lblPrompt 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   735
      TabIndex        =   16
      Top             =   4170
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "提示："
      Height          =   180
      Left            =   165
      TabIndex        =   15
      Top             =   4155
      Width           =   540
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "时间：2001年11月16日 4～8时"
      Height          =   180
      Left            =   195
      TabIndex        =   9
      Top             =   165
      Width           =   2430
   End
End
Attribute VB_Name = "frmCaseTendBodySetLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintNowCol As Integer
Private mintMinCol As Integer
Private mintMaxCol As Integer
Private mblnChanged As Boolean
Private mfrmParent As Object
Private mlng护理等级 As Long
Private mint心率应用 As Integer
Private mblnStart As Boolean
Private Enum mCol
    项目 = 1
    项目id
    最小值
    最大值
    单位值
    最高行
    数据
    标记
    部位
    未记说明
End Enum

Private Enum GraphDataRow
    更改标志 = 0
    曲线数据 = 1
    上标说明 = 2
    手术标志 = 3
    部位标志 = 4
    入院标志 = 5
    转科标志 = 6
    换床标志 = 7
    出院标志 = 8
    入科标志 = 9
    复试标志 = 10
    下标说明 = 11
    断开标志 = 12
    出生标志 = 13
    曲线时间 = 14
    未记说明 = 15
End Enum

Public Function ShowEdit(ByVal frmParent As Object, ByVal intNowCol As Integer, ByVal intMinCol As Integer, ByVal intMaxCol As Long, ByVal lng护理等级 As Long, ByVal int心率应用 As Integer) As Boolean
    
    mblnChanged = False
    mblnStart = True
    
    mint心率应用 = int心率应用
    
    If intNowCol = -1 Then Exit Function
    
    mintNowCol = intNowCol
    mintMinCol = intMinCol
    mintMaxCol = intMaxCol
    
    mlng护理等级 = lng护理等级
    
    Set mfrmParent = frmParent
    
    Call InitData
    Call LoadNowData
    
'    vsf.SetFocus
    
    Me.Show 1
    
    ShowEdit = mblnChanged
    
End Function

'Private Sub chk_Click(Index As Integer)
'
'    If chk(0).Value = 1 Then
'        vsf.EditMode(mCol.数据) = 0
'        vsf.EditMode(mCol.部位) = 0
'        vsf.EditMode(mCol.标记) = 0
'        vsf.Cell(flexcpText, 1, mCol.数据, vsf.Rows - 1, mCol.部位) = ""
'    Else
'        vsf.EditMode(mCol.数据) = 1
'    End If
'End Sub

'Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        zlCommFun.PressKey vbKeyTab
'    End If
'End Sub

Private Sub chkContinue_Click()
    vsf.SetFocus
    vsf.ShowCell vsf.Row, vsf.Col
End Sub

Private Sub cmdCanc_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim intCol As Integer
    Dim intCount As Integer
    Dim dbValue As Double
    Dim str方式 As String
    Dim aryValue() As String
    Dim aryPart() As String     '部位
    Dim intRewrite As Integer
    
    On Error GoTo ErrHead
    
    With mfrmParent.GetmshScale
        '保存注释说明
        aryValue = Split(.TextMatrix(GraphDataRow.更改标志, mintNowCol + .FixedCols), ";")
        intRewrite = Val(aryValue(0))
        If Trim(txt(0).Text) <> "" Or Trim(txt(1).Text) <> "" Then
            '存在内容，相当于增加或修改操作
            Select Case intRewrite
            Case 0
                aryValue(0) = 2
            Case 1
                aryValue(0) = 3
            Case 2
                aryValue(0) = 2
            Case 3
                aryValue(0) = 3
            Case 4
                aryValue(0) = 3
            End Select
        Else
            '没有内容，相当于删除操作
            Select Case intRewrite
            Case 0
                aryValue(0) = 0
            Case 1
                aryValue(0) = 4
            Case 2
                aryValue(0) = 0
            Case 3
                aryValue(0) = 4
            Case 4
                aryValue(0) = 4
            End Select
        End If
        .TextMatrix(GraphDataRow.更改标志, mintNowCol + .FixedCols) = Join(aryValue, ";")
        .TextMatrix(GraphDataRow.上标说明, mintNowCol + .FixedCols) = Trim(txt(0).Text)
        .TextMatrix(GraphDataRow.下标说明, mintNowCol + .FixedCols) = Trim(txt(1).Text)
'        .TextMatrix(GraphDataRow.断开标志, mintNowCol + .FixedCols) = chk(0).Value

        '保存线条数据
        For intCount = 1 To vsf.Rows - 1
        
            aryValue = Split(.TextMatrix(GraphDataRow.更改标志, mintNowCol + .FixedCols), ";")
            intRewrite = Val(aryValue(intCount))
            
            If Trim(vsf.TextMatrix(intCount, mCol.数据)) <> "" Or Trim(vsf.TextMatrix(intCount, mCol.标记)) <> "" Or Trim(vsf.TextMatrix(intCount, mCol.未记说明)) <> "" Then
                '存在内容，相当于增加或修改操作
                Select Case intRewrite
                Case 0
                    aryValue(intCount) = 2
                Case 1
                    aryValue(intCount) = 3
                Case 2
                    aryValue(intCount) = 2
                Case 3
                    aryValue(intCount) = 3
                Case 4
                    aryValue(intCount) = 3
                End Select
            Else
                '没有内容，相当于删除操作
                Select Case intRewrite
                Case 0
                    aryValue(intCount) = 0
                Case 1
                    aryValue(intCount) = 4
                Case 2
                    aryValue(intCount) = 0
                Case 3
                    aryValue(intCount) = 4
                Case 4
                    aryValue(intCount) = 4
                End Select
            End If
            
            .TextMatrix(GraphDataRow.更改标志, mintNowCol + .FixedCols) = Join(aryValue, ";")
            
            aryValue = Split(.TextMatrix(GraphDataRow.曲线数据, mintNowCol + .FixedCols), ";")
            If Trim(vsf.TextMatrix(intCount, mCol.数据)) <> "" Then
            
                dbValue = ((Val(vsf.TextMatrix(intCount, mCol.最大值)) - Val(vsf.TextMatrix(intCount, mCol.数据))) / Val(vsf.TextMatrix(intCount, mCol.单位值)) + Val(vsf.TextMatrix(intCount, mCol.最高行)) - 1) * .ROWHEIGHT(1)
                aryValue(intCount) = dbValue
                
                If Trim(vsf.TextMatrix(intCount, mCol.标记)) <> "" Then
                    dbValue = ((Val(vsf.TextMatrix(intCount, mCol.最大值)) - Val(vsf.TextMatrix(intCount, mCol.标记))) / Val(vsf.TextMatrix(intCount, mCol.单位值)) + Val(vsf.TextMatrix(intCount, mCol.最高行)) - 1) * .ROWHEIGHT(1)
                    aryValue(intCount) = aryValue(intCount) & "," & dbValue
                End If
            Else
                aryValue(intCount) = ""
            End If
            .TextMatrix(GraphDataRow.曲线数据, mintNowCol + .FixedCols) = Join(aryValue, ";")
            
            '处理体温的"不升"
            If Trim(vsf.TextMatrix(intCount, mCol.数据)) = "不升" And vsf.RowData(intCount) = 1 Then
                aryValue = Split(.TextMatrix(GraphDataRow.未记说明, mintNowCol + .FixedCols), ";")
                aryValue(intCount) = "不升"
                .TextMatrix(GraphDataRow.未记说明, mintNowCol + .FixedCols) = Join(aryValue, ";")
            Else
                aryValue = Split(.TextMatrix(GraphDataRow.未记说明, mintNowCol + .FixedCols), ";")
                aryValue(intCount) = vsf.TextMatrix(intCount, mCol.未记说明)
                .TextMatrix(GraphDataRow.未记说明, mintNowCol + .FixedCols) = Join(aryValue, ";")
            End If
            
            Select Case Val(vsf.RowData(intCount))
            Case 1, 2, 3            '.TextMatrix(GraphDataRow.部位标志, mintNowCol + .FixedCols)
                str方式 = ""
                If Trim(vsf.TextMatrix(intCount, mCol.部位)) <> "" Then
                    str方式 = Trim(vsf.TextMatrix(intCount, mCol.部位))
                Else
                    If Val(vsf.RowData(intCount)) = 1 Then str方式 = "腋温"
                End If
                
                '组织部位
                aryPart = Split(.TextMatrix(GraphDataRow.部位标志, mintNowCol + .FixedCols), ";")
                aryPart(intCount) = str方式
                .TextMatrix(GraphDataRow.部位标志, mintNowCol + .FixedCols) = Join(aryPart, ";")
                
            End Select
        Next

    End With
    
    '调用上级窗体进行图形处理
    Call mfrmParent.DrawPaper
    Call mfrmParent.DrawGraph
    
    mblnChanged = True
    
    If chkContinue.Value = 0 Then
        Unload Me
        Exit Sub
    End If
    
    If mintNowCol < mintMaxCol Then
        mintNowCol = mintNowCol + 1
    Else
        'If MsgBox("已经达到本体温表最大时间，是否重新输入？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Unload Me
        Exit Sub
'        Else
'            intNowCol = intMinCol
'        End If
    End If
    
    Call LoadNowData
    
    vsf.SetFocus
    vsf.ShowCell 1, 2
    Exit Sub
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function InitData() As Boolean
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    
    On Error GoTo ErrHead

    With vsf
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "项目", 1080, 1
        .NewColumn "项目id", 0, 1
        .NewColumn "最小值", 0, 1
        .NewColumn "最大值", 0, 1
        .NewColumn "单位值", 0, 1
        .NewColumn "最高行", 0, 1
        .NewColumn "数据", 900, 1, , 1
        .NewColumn "标记", 900, 1
        .NewColumn "部位", 750, 1
        .NewColumn "未记说明", 1080, 1, "...", 1
        
        .FixedCols = 2
                
        .Body.ColHidden(mCol.最小值) = True
        .Body.ColHidden(mCol.最大值) = True
        .Body.ColHidden(mCol.单位值) = True
        .Body.ColHidden(mCol.最高行) = True
        .Body.ColHidden(mCol.项目id) = True
        
        .Body.WordWrap = True
    End With

    InitData = True
    
    Exit Function
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'Private Sub txtComment_GotFocus()
'    zlControl.TxtSelAll cboComment
'    zlCommFun.OpenIme True
'End Sub
'
'Private Sub txtComment_KeyPress(KeyAscii As Integer)
'    If KeyAscii = Asc(";") Or KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
'    If KeyAscii = 13 Then
'        KeyAscii = 0
'        cmdOK.SetFocus
'    End If
'End Sub
'
'Private Sub txtComment_LostFocus()
'    cboComment.Text = Replace(Me.cboComment.Text, "'", "")
'    zlCommFun.OpenIme
'End Sub

Private Sub Form_Activate()
    
    If mblnStart = False Then Exit Sub
    mblnStart = False
    
    vsf.Col = mCol.数据
    vsf.SetFocus
    
End Sub

'Private Sub txt_Change(Index As Integer)
'
'    Select Case txt(Index).Text
'    Case "拒测", "未测", "请假", "外出"
'
'        chk(0).Value = 1
'
'    Case Else
'
'        chk(0).Value = 0
'
'    End Select
'
'End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    Call zlControl.TxtSelAll(txt(Index))
    
    Select Case Index
    Case 0, 1
        zlCommFun.OpenIme True
    End Select
    
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 0, 1
        zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)

    
End Sub



Private Sub txtItem_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtItem(Index)
End Sub

Private Sub txtItem_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If KeyAscii = Asc("'") Or KeyAscii = Asc(";") Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtItem_LostFocus(Index As Integer)
    txtItem(Index).Text = Replace(Me.txtItem(Index).Text, "'", "")
End Sub

Private Sub txtItem_Validate(Index As Integer, Cancel As Boolean)
    Dim aryPara() As String
    
    On Error GoTo ErrHead
    
    '提取指定项目定义：最大值；最小值；单位值；最高行
    If Trim(txtItem(Index).Text) = "" Then Exit Sub
    
    If Not MouseInRect(cmdCanc.hWnd) Then
        If IsNumeric(Trim(Me.txtItem(Index).Text)) = False Then
            MsgBox "项目【" & Me.lblItem(Index).Tag & "】的值必须为数字！", vbExclamation, gstrSysName
            Cancel = True: Exit Sub
        End If
        aryPara = Split(Me.txtItem(Index).Tag, ";")
        If Format(Me.txtItem(Index).Text) > Val(aryPara(0)) Or Format(Me.txtItem(Index).Text) < Val(aryPara(1)) Then
            MsgBox "项目【" & Me.lblItem(Index).Tag & "】的值超过允许范围：" & aryPara(1) & "～" & aryPara(0), vbExclamation, gstrSysName
            Cancel = True: Exit Sub
        End If
    End If
    Exit Sub
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadNowData()
    
    '装入指定日期数据
    Dim aryValue() As String
    Dim aryNote() As String
    Dim aryPara() As String
    Dim dtNow As Date
    Dim dtNowTmp As Date
    Dim lngHourBegin As Long
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim intCount As Integer
    Dim intCol As Long
    Dim dblValue As Double
    Dim strTmp As String
    
    On Error GoTo ErrHead
    
    lngHourBegin = Val(zlDatabase.GetPara(67, glngSys, , 4))

    aryValue = Split(mfrmParent.GetPicScale.Tag, ";")
    
    strTmp = GetCurveDateTime(mintNowCol + 1, CDate(aryValue(0)), lngHourBegin)
    If strTmp <> "" Then
        strTmp = "时间：" & Format(Split(strTmp, ",")(0), "yyyy-MM-dd") & " " & Format(Split(strTmp, ",")(0), "HH时mm分") & "～" & Format(Split(strTmp, ",")(1), "HH时mm分")
    End If
    lblTime.Caption = strTmp
    
    vsf.Rows = 2
    vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
    vsf.RowData(1) = 0
    With mfrmParent.GetmshScale
        
        aryValue = Split(.TextMatrix(GraphDataRow.曲线数据, mintNowCol + .FixedCols), ";")
        aryNote = Split(.TextMatrix(GraphDataRow.未记说明, mintNowCol + .FixedCols), ";")
        
        For intCount = 0 To .FixedCols - 1
            
            If Val(vsf.RowData(vsf.Rows - 1)) <> 0 Then vsf.Rows = vsf.Rows + 1
            
            vsf.RowData(vsf.Rows - 1) = .ColData(intCount)
            vsf.TextMatrix(vsf.Rows - 1, mCol.项目) = .TextMatrix(0, intCount)
            
            aryPara = Split(mfrmParent.GetpicLine(intCount).Tag, ";")
            'RS!最大值 & ";" & RS!最小值 & ";" & RS!单位值 & ";" & RS!最高行
            
            vsf.TextMatrix(vsf.Rows - 1, mCol.最小值) = aryPara(1)
            vsf.TextMatrix(vsf.Rows - 1, mCol.最大值) = aryPara(0)
            vsf.TextMatrix(vsf.Rows - 1, mCol.单位值) = aryPara(2)
            vsf.TextMatrix(vsf.Rows - 1, mCol.最高行) = aryPara(3)
            
            If Trim(aryValue(intCount + 1)) <> "" Then
                For intCol = 0 To UBound(Split(aryValue(intCount + 1), ","))
                    
                    dblValue = aryPara(0) - (Val(Split(aryValue(intCount + 1), ",")(intCol)) / .ROWHEIGHT(1) - aryPara(3) + 1) * aryPara(2)
                    If InStr(.TextMatrix(0, intCount), "体温") > 0 Then
                        dblValue = Format(dblValue, "0.00")
                    Else
                        dblValue = Format(dblValue, "0")
                    End If
                    
                    If intCol = 0 Then
                        If aryNote(intCount + 1) = "不升" And CStr(Val(dblValue)) = "0" Then
                            vsf.TextMatrix(vsf.Rows - 1, mCol.数据) = "不升"
                        Else
                            vsf.TextMatrix(vsf.Rows - 1, mCol.数据) = dblValue
                        End If
                    Else
                        vsf.TextMatrix(vsf.Rows - 1, mCol.标记) = dblValue
                    End If
                Next
                
            End If
            
            If aryNote(intCount + 1) = "不升" And CStr(Val(dblValue)) = "0" Then
                vsf.TextMatrix(vsf.Rows - 1, mCol.未记说明) = ""
            Else
                vsf.TextMatrix(vsf.Rows - 1, mCol.未记说明) = aryNote(intCount + 1)
            End If
            
            Select Case Val(vsf.RowData(vsf.Rows - 1))
            Case 1, 2, 3
                vsf.TextMatrix(vsf.Rows - 1, mCol.部位) = Split(.TextMatrix(GraphDataRow.部位标志, mintNowCol + .FixedCols), ";")(intCount + 1)
            End Select
        Next
        
        txt(0).Text = .TextMatrix(GraphDataRow.上标说明, mintNowCol + .FixedCols)
        txt(1).Text = .TextMatrix(GraphDataRow.下标说明, mintNowCol + .FixedCols)
'        chk(0).Value = Val(.TextMatrix(GraphDataRow.断开标志, mintNowCol + .FixedCols))
    End With
    
   
    Exit Sub
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Select Case Col
    Case mCol.数据
    
        Call vsf.Body.AutoSize(mCol.数据, mCol.数据)
        
    End Select
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    On Error Resume Next

    vsf.ComboList(mCol.部位) = ""
    vsf.EditMode(mCol.部位) = 0
    vsf.ComboList(mCol.标记) = ""
    vsf.EditMode(mCol.标记) = 0
    vsf.ComboList(mCol.部位) = ""
    
    Select Case Val(vsf.RowData(NewRow))
    Case -1
        
    Case 0
        
        vsf.ComboList(mCol.部位) = "|外出|检查|手术|拒查|未查"
    
    Case 1
        vsf.ComboList(mCol.部位) = "口温|腋温|肛温"
        vsf.EditMode(mCol.部位) = 1
        
        vsf.ComboList(mCol.标记) = ""
        vsf.EditMode(mCol.标记) = 1
    Case 2
        vsf.ComboList(mCol.部位) = " |起搏器"
        vsf.EditMode(mCol.部位) = 1
        
        If mint心率应用 = 2 Then
            vsf.ComboList(mCol.标记) = ""
            vsf.EditMode(mCol.标记) = 1
        End If
    Case 3
        vsf.ComboList(mCol.部位) = "自主呼吸|呼吸机"
        vsf.EditMode(mCol.部位) = 1
    
    End Select
    
'    If chk(0).Value = 1 Then
'        vsf.EditMode(mCol.数据) = 0
'        vsf.EditMode(mCol.部位) = 0
'        vsf.EditMode(mCol.标记) = 0
'    End If
    
    Dim strTmp As String
    
    If vsf.TextMatrix(NewRow, mCol.最小值) <> "" Or vsf.TextMatrix(NewRow, mCol.最小值) <> "" Then
        strTmp = "范围：" & vsf.TextMatrix(NewRow, mCol.最小值) & "～" & vsf.TextMatrix(NewRow, mCol.最大值) & " "
    End If
    
    Select Case Val(vsf.RowData(NewRow))
    Case 1
        strTmp = strTmp & "标记表示物理降温的温度，部位为测体温的部位。"
    Case 2
        strTmp = strTmp & "标记表示心率的值（与脉搏不同时才记录）。"
    End Select
    lblPrompt.Caption = strTmp
    
    If Val(vsf.RowData(NewRow)) = 0 Then
        zlCommFun.OpenIme True
    Else
        zlCommFun.OpenIme False
    End If
End Sub

Private Sub vsf_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case mCol.标记, mCol.数据, mCol.未记说明
        vsf.TextMatrix(Row, Col) = ""
    End Select
    Cancel = True
End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim intCount As Integer '未记说明有内容的个数,只有一项时自动将其它空白项更新
    Dim intRow As Integer, intRows As Integer
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    Select Case Col
    Case mCol.未记说明
        
        strSQL = "Select 编码,名称,RowNum As ID,1 As 末级 From 常用体温说明"
        If ShowGrdSelectDialog(Me, vsf, "名称,3000,0,0", Me.Name & "\常用体温说明", "请从下面选择一个未记录说明。", strSQL, rs, 4500, 4500, False, 2) Then
            vsf.EditText = zlCommFun.NVL(rs("名称").Value)
            vsf.Cell(flexcpData, Row, Col) = zlCommFun.NVL(rs("名称").Value)
            vsf.TextMatrix(Row, Col) = zlCommFun.NVL(rs("名称").Value)
            vsf.TextMatrix(vsf.Row, mCol.数据) = ""
            vsf.TextMatrix(vsf.Row, mCol.标记) = ""
            vsf.TextMatrix(vsf.Row, mCol.部位) = ""
            
            '如果其它曲线的未记数据为空,直接更新
            intRows = vsf.Rows - 1
            For intRow = 1 To intRows
                If vsf.TextMatrix(intRow, mCol.未记说明) = "" And vsf.TextMatrix(intRow, mCol.数据) = "" Then
                    intCount = intCount + 1
                End If
            Next
            '剩下的项目的数据与标记都为空则更新
            If intCount = intRows - 1 Then
                For intRow = 1 To intRows
                    If vsf.TextMatrix(intRow, mCol.未记说明) = "" Then
                        vsf.TextMatrix(intRow, mCol.未记说明) = zlCommFun.NVL(rs("名称").Value)
                    End If
                Next
            End If
        End If
    End Select
    
End Sub

Private Sub vsf_ChangeEdit()
    Select Case vsf.Col
    Case mCol.数据
        If Val(vsf.RowData(vsf.Row)) <> 0 Then
            vsf.TextMatrix(vsf.Row, mCol.数据) = vsf.EditText
            Call vsf.Body.AutoSize(mCol.数据, mCol.数据)
            
            If vsf.TextMatrix(vsf.Row, mCol.数据) <> "" Then vsf.TextMatrix(vsf.Row, mCol.未记说明) = ""
        End If
    Case mCol.标记
        
        vsf.TextMatrix(vsf.Row, mCol.未记说明) = ""
        
    Case mCol.未记说明
        vsf.TextMatrix(vsf.Row, mCol.未记说明) = vsf.EditText
        If vsf.TextMatrix(vsf.Row, mCol.未记说明) <> "" Then
            vsf.TextMatrix(vsf.Row, mCol.数据) = ""
            vsf.TextMatrix(vsf.Row, mCol.标记) = ""
            vsf.TextMatrix(vsf.Row, mCol.部位) = ""
        End If
        
    End Select
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    If KeyCode = vbKeyReturn Then
        
        If Col = mCol.未记说明 Then
            
            vsf.Cell(flexcpData, Row, Col) = vsf.EditText
            vsf.TextMatrix(Row, Col) = vsf.EditText
            
        End If
        
        If Row = vsf.Rows - 1 Then cmdOK.SetFocus
        
    End If
End Sub


Private Sub vsf_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    If KeyAscii = vbKeyReturn And Row = vsf.Rows - 1 Then
        cmdOK.SetFocus
    End If
    
    On Error Resume Next
    
    If KeyAscii <> vbKeyReturn Then
        If Val(vsf.RowData(Row)) <> 0 Then
            If Col <> mCol.未记说明 Then
                If FilterKeyAscii(KeyAscii, 99, "0123456789.") = 0 Then KeyAscii = 0
            Else
                If FilterKeyAscii(KeyAscii, 99, "'") > 0 Then KeyAscii = 0
            End If
        End If
    End If
    
End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    
    On Error Resume Next
    
    If KeyAscii <> vbKeyReturn Then
        If Val(vsf.RowData(Row)) <> 0 Then
'            If FilterKeyAscii(KeyAscii, 99, "0123456789.") = 0 Then KeyAscii = 0
        End If
    End If
    
End Sub

Private Sub vsf_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    Select Case Col
    Case mCol.数据
        GoTo CheckPoint
    Case mCol.标记
        
        Select Case Val(vsf.RowData(Row))
        Case 1
            GoTo CheckPoint
        Case 2
            GoTo CheckPoint
        End Select
        
    End Select
    
    Exit Sub
    
CheckPoint:
    If Trim(vsf.EditText) <> "" Then
        '体温不检查
        If vsf.RowData(Row) <> 1 And (vsf.TextMatrix(Row, mCol.最小值) <> "" Or vsf.TextMatrix(Row, mCol.最大值) <> "") Then
            

            If Val(vsf.EditText) < Val(vsf.TextMatrix(Row, mCol.最小值)) Or Val(vsf.EditText) > Val(vsf.TextMatrix(Row, mCol.最大值)) Then
'                Cancel = True
                ShowSimpleMsg "“" & vsf.TextMatrix(Row, mCol.项目) & " ”的范围应在（" & Val(vsf.TextMatrix(Row, mCol.最小值)) & "～" & Val(vsf.TextMatrix(Row, mCol.最大值)) & "）之间！"
            End If

            
        End If

    End If
End Sub



