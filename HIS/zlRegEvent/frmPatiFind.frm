VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPatiFind 
   AutoRedraw      =   -1  'True
   Caption         =   "查找病人"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7230
   Icon            =   "frmPatiFind.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraInfo 
      Caption         =   " 病人信息 "
      Height          =   2655
      Left            =   120
      TabIndex        =   17
      Top             =   30
      Width           =   5565
      Begin VB.TextBox txtIC卡 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3975
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1530
         Width           =   1455
      End
      Begin VB.TextBox txt医保号 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3975
         MaxLength       =   20
         TabIndex        =   10
         Top             =   1900
         Width           =   1455
      End
      Begin VB.TextBox txtValue 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1065
         MaxLength       =   20
         TabIndex        =   9
         Top             =   1900
         Width           =   1830
      End
      Begin VB.ComboBox cbo年龄单位 
         Height          =   300
         Left            =   2310
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1095
         Width           =   580
      End
      Begin MSComCtl2.DTPicker dtp就诊E 
         Height          =   300
         Left            =   3960
         TabIndex        =   12
         Top             =   2265
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   185925635
         CurrentDate     =   37401
      End
      Begin MSComCtl2.DTPicker dtp就诊B 
         Height          =   300
         Left            =   1065
         TabIndex        =   11
         Top             =   2265
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   185925635
         CurrentDate     =   37401
      End
      Begin VB.TextBox txt身份证 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1065
         MaxLength       =   18
         TabIndex        =   7
         Top             =   1500
         Width           =   1830
      End
      Begin MSComCtl2.DTPicker dtp生日 
         Height          =   300
         Left            =   3975
         TabIndex        =   6
         Top             =   1095
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   185925635
         CurrentDate     =   37401
      End
      Begin VB.TextBox txtOld 
         Height          =   300
         Left            =   1065
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1095
         Width           =   1215
      End
      Begin VB.ComboBox cbo性别 
         Height          =   300
         Left            =   3975
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   690
         Width           =   1455
      End
      Begin VB.TextBox txt姓名 
         Height          =   300
         Left            =   1065
         MaxLength       =   100
         TabIndex        =   2
         Top             =   690
         Width           =   1830
      End
      Begin VB.TextBox txt门诊号 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3975
         MaxLength       =   18
         TabIndex        =   1
         Top             =   285
         Width           =   1455
      End
      Begin VB.TextBox txt病人ID 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1065
         TabIndex        =   0
         Top             =   285
         Width           =   1830
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IC卡号"
         Height          =   180
         Left            =   3360
         TabIndex        =   31
         Top             =   1590
         Width           =   540
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医保号"
         Height          =   180
         Left            =   3375
         TabIndex        =   30
         Top             =   1965
         Width           =   540
      End
      Begin VB.Label lblKind 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "就诊卡↓"
         Height          =   180
         Left            =   270
         TabIndex        =   29
         Top             =   1965
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   3285
         TabIndex        =   26
         Top             =   2325
         Width           =   180
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上次就诊"
         Height          =   180
         Left            =   270
         TabIndex        =   25
         Top             =   2325
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号"
         Height          =   180
         Left            =   270
         TabIndex        =   24
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生日期"
         Height          =   180
         Left            =   3195
         TabIndex        =   23
         Top             =   1155
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Height          =   180
         Left            =   630
         TabIndex        =   22
         Top             =   1155
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   180
         Left            =   3555
         TabIndex        =   21
         Top             =   750
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   630
         TabIndex        =   20
         Top             =   750
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊号"
         Height          =   180
         Left            =   3375
         TabIndex        =   19
         Top             =   345
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人ID"
         Height          =   180
         Left            =   450
         TabIndex        =   18
         Top             =   345
         Width           =   540
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshPati 
      Height          =   2820
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "双击或回车查看更详细的信息"
      Top             =   3045
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   4974
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmPatiFind.frx":058A
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.PictureBox picCmd 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1980
      Left            =   5880
      ScaleHeight     =   1980
      ScaleWidth      =   1275
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   120
      Width           =   1275
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   75
         TabIndex        =   16
         Top             =   1020
         Width           =   1100
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "选择(&S)"
         Height          =   350
         Left            =   75
         TabIndex        =   15
         Top             =   600
         Width           =   1100
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "查找(&F)"
         Height          =   350
         Left            =   75
         TabIndex        =   13
         Top             =   15
         Width           =   1100
      End
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00707070&
      Caption         =   " 病人查找结果"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   27
      Top             =   2835
      Width           =   6990
   End
   Begin VB.Menu mnuPop 
      Caption         =   "医疗卡选择"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuItems 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmPatiFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mlng病人ID As Long '出口参数:病人ID
'-----------------------------------------------------
'结算卡相关
Private mcllBrushCard As Collection
Private Type Tp_CardSquare
    bln缺省卡号密文 As Boolean
    lng缺省卡类别ID As Long
    int缺省卡号长度 As Integer
End Type
Private mTyCard As Tp_CardSquare
'-----------------------------------------------------
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
'功能：根据当前设置条件查找病人
    Dim strSQL As String, i As Integer, rsTmp As ADODB.Recordset
    Dim DateB As Date, DateE As Date, strMCAccount As String, str非在院 As String
    Dim lng病人ID As Long, lng卡类别ID As Long, strErrMsg As String, strPassWord As String
    Dim strKind As String
    
    If Trim(txt病人ID.Text) <> "" Then
        strSQL = strSQL & " And 病人ID=[1]"
        lng病人ID = Val(Trim(txt病人ID.Text))
    End If
    If Trim(txt门诊号.Text) <> "" Then
        strSQL = strSQL & " And 门诊号=[2]"
    End If
    If Trim(txt姓名.Text) <> "" Then
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(txt姓名.Text, 1))) > 0 Then
            strSQL = strSQL & " And Upper(姓名) Like [3]"
        Else
            strSQL = strSQL & " And 姓名 Like [4]"
        End If
    End If
    If cbo性别.Text <> "" Then
        strSQL = strSQL & " And 性别=[5]"
    End If
    If Trim(txtOld.Text) <> "" Then
        strSQL = strSQL & " And 年龄=[6]"
    End If
    If Not IsNull(dtp生日.Value) Then
        strSQL = strSQL & " And 出生日期=[7]"
    End If
    If Trim(txt身份证.Text) <> "" Then
        strSQL = strSQL & " And 身份证号=[8]"
    End If
    If Trim(txtValue.Text) <> "" Then
        strKind = mnuPopuItems(Val(lblKind.Tag)).Tag
        Select Case strKind
        Case "姓名"
        Case Else
            '其他类别的,获取相关的病人ID
            '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
            '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
            '第7位后,就只能用索引,不然取不到数
            lng卡类别ID = Val(mcllBrushCard(Val(lblKind.Tag) + 1)(3))
            If lng卡类别ID <> 0 Then
                If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, Trim(txtValue.Text), False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
            Else
                If gobjSquare.objSquareCard.zlGetPatiID(strKind, Trim(txtValue.Text), False, lng病人ID, _
                    strPassWord, strErrMsg) = False Then lng病人ID = 0
            End If
            strSQL = strSQL & " And 病人ID=[1]"
        End Select
    End If
    If Not IsNull(dtp就诊B.Value) And Not IsNull(dtp就诊E.Value) Then
        If dtp就诊E.Value <= dtp就诊B.Value Then
            MsgBox "上次就诊的结束时间必须大于开始时间！", vbInformation, gstrSysName
            dtp就诊E.SetFocus: Exit Sub
        End If
        strSQL = strSQL & " And 就诊时间 Between [9] And [10]"
    ElseIf Not IsNull(dtp就诊B.Value) Then
        DateB = CDate(Format(dtp就诊B.Value, "yyyy-MM-dd 00:00:00"))
        DateE = CDate(Format(dtp就诊B.Value, "yyyy-MM-dd 23:59:59"))
        strSQL = strSQL & " And 就诊时间 Between [11] And [12]"
    End If
    
    If Trim(txtIC卡.Text) <> "" Then
        strSQL = strSQL & " And IC卡号=[13]"
    End If
    
    If Trim(txt医保号.Text) <> "" Then
        strSQL = strSQL & " And 医保号=[15]"
    End If
    
    If strSQL = "" Then
        MsgBox "请至少设置一个查找条件！", vbInformation, gstrSysName
        txt姓名.SetFocus: Exit Sub
    End If
    
    strMCAccount = Trim(txt医保号.Text)
    
    On Error GoTo errH
    Screen.MousePointer = 11
    str非在院 = " And Not Exists(Select 1 From 病案主页 Where 病人ID=A.病人ID And 主页ID<>0 And 主页ID=A.主页ID And Nvl(病人性质,0)=0 And 出院日期 is Null)"
    strSQL = _
        " Select " & _
        " 病人ID,门诊号,费别,医疗付款方式,姓名,性别,年龄,To_Char(出生日期,'YYYY-MM-DD') as 出生日期," & _
        " 身份证号,出生地点,家庭地址,工作单位,身份,职业,学历,To_Char(就诊时间,'YYYY-MM-DD HH24:MI') as 就诊时间" & _
        " From 病人信息 A" & _
        " Where 停用时间 is NULL " & str非在院 & strSQL & _
        " Order by 病人ID"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, Trim(txt门诊号.Text), UCase(txt姓名.Text) & "%", _
                txt姓名.Text & "%", cbo性别.Text, IIf(IsNumeric(txtOld.Text), txtOld.Text & cbo年龄单位.Text, txtOld.Text), IIf(IsNull(dtp生日.Value), "", dtp生日.Value), txt身份证.Text, _
                IIf(IsNull(dtp就诊B.Value), "", dtp就诊B.Value), IIf(IsNull(dtp就诊E.Value), "", dtp就诊E.Value), _
                DateB, DateE, Trim(txtIC卡.Text), "", strMCAccount)
    
    If Not rsTmp.EOF Then
        lblInfo.Caption = " 病人查找结果:共 " & rsTmp.RecordCount & " 个满足条件的病人"
        Set mshPati.DataSource = rsTmp
        For i = 0 To mshPati.Cols - 1
            mshPati.ColAlignmentFixed(i) = 4
        Next
        mshPati.TextMatrix(0, 0) = "病人ID": mshPati.ColWidth(0) = 750: mshPati.ColAlignment(0) = 1
        mshPati.TextMatrix(0, 1) = "门诊号": mshPati.ColWidth(1) = 750: mshPati.ColAlignment(1) = 1
        mshPati.TextMatrix(0, 2) = "费别": mshPati.ColWidth(2) = 850: mshPati.ColAlignment(2) = 1
        mshPati.TextMatrix(0, 3) = "付款方式": mshPati.ColWidth(3) = 850: mshPati.ColAlignment(3) = 1
        mshPati.TextMatrix(0, 4) = "姓名": mshPati.ColWidth(4) = 700: mshPati.ColAlignment(4) = 1
        mshPati.TextMatrix(0, 5) = "性别": mshPati.ColWidth(5) = 500: mshPati.ColAlignment(5) = 4
        mshPati.TextMatrix(0, 6) = "年龄": mshPati.ColWidth(6) = 500: mshPati.ColAlignment(6) = 1
        mshPati.TextMatrix(0, 7) = "出生日期": mshPati.ColWidth(7) = 1000: mshPati.ColAlignment(7) = 4
        mshPati.TextMatrix(0, 8) = "身份证号": mshPati.ColWidth(8) = 1600: mshPati.ColAlignment(8) = 1
        mshPati.TextMatrix(0, 9) = "出生地点": mshPati.ColWidth(9) = 2000: mshPati.ColAlignment(9) = 1
        mshPati.TextMatrix(0, 10) = "家庭地址": mshPati.ColWidth(10) = 2000: mshPati.ColAlignment(10) = 1
        mshPati.TextMatrix(0, 11) = "工作单位": mshPati.ColWidth(11) = 2000: mshPati.ColAlignment(11) = 1
        mshPati.TextMatrix(0, 12) = "身份": mshPati.ColWidth(12) = 1000: mshPati.ColAlignment(12) = 1
        mshPati.TextMatrix(0, 13) = "职业": mshPati.ColWidth(13) = 1000: mshPati.ColAlignment(13) = 1
        mshPati.TextMatrix(0, 14) = "学历": mshPati.ColWidth(14) = 500: mshPati.ColAlignment(14) = 1
        mshPati.TextMatrix(0, 15) = "上次就诊时间": mshPati.ColWidth(15) = 1600: mshPati.ColAlignment(15) = 4
    Else
        lblInfo.Caption = " 病人查找结果"
        mshPati.Clear
        mshPati.ClearStructure
        mshPati.Cols = 2: mshPati.Rows = 2
        mshPati.FixedCols = 0: mshPati.FixedRows = 1
    End If
    mshPati.Row = 1: mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
    mshPati.TopRow = 1
    Call mshPati_EnterCell
    Screen.MousePointer = 0
    mshPati.SetFocus
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub cmdSel_Click()
    If Val(mshPati.TextMatrix(mshPati.Row, 0)) = 0 Then
        MsgBox "没有病人信息可以选择！", vbInformation, gstrSysName
        Exit Sub
    End If
    mlng病人ID = Val(mshPati.TextMatrix(mshPati.Row, 0))
    Unload Me
End Sub

Private Sub dtp就诊B_Change()
    If IsNull(dtp就诊B.Value) Then dtp就诊E.Value = Null
End Sub

Private Sub dtp就诊B_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtp就诊E_Change()
    If IsNull(dtp就诊B.Value) And Not IsNull(dtp就诊E.Value) Then dtp就诊E.Value = Null
End Sub

Private Sub dtp就诊E_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtp生日_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("'&[]", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = 13 And TypeName(ActiveControl) <> "DTPicker" And Not ActiveControl Is mshPati Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim Datsys As Date
    
    txt姓名.MaxLength = zlGetPatiInforMaxLen.intPatiName
    
    Call InitMenus
    Call RestoreWinState(Me, App.ProductName)
    mlng病人ID = 0
    Datsys = zldatabase.Currentdate
    
    dtp就诊E.MaxDate = Datsys
    dtp就诊B.MaxDate = dtp就诊E.MaxDate
    dtp就诊B.Value = DateAdd("m", -1, Datsys)
    dtp就诊E.Value = Datsys
    dtp就诊B.Value = Null
    dtp就诊E.Value = Null
    
    dtp生日.MaxDate = Datsys
    dtp生日.Value = DateAdd("yyyy", -25, Datsys)
    dtp生日.Value = Null
    
    Call mshPati_EnterCell
    
    On Error GoTo errH
    strSQL = "Select 名称 From 性别"
    Call zldatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    cbo性别.AddItem ""
    Do While Not rsTmp.EOF
        cbo性别.AddItem rsTmp!名称
        rsTmp.MoveNext
    Loop
    cbo性别.ListIndex = 0
    
    '年龄单位
    cbo年龄单位.AddItem "岁"
    cbo年龄单位.AddItem "月"
    cbo年龄单位.AddItem "天"
    cbo年龄单位.ListIndex = 0
    
    txt医保号.MaxLength = 20
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.ScaleWidth - picCmd.Width > fraInfo.Left + fraInfo.Width Then
        picCmd.Left = Me.ScaleWidth - picCmd.Width
    Else
        picCmd.Left = fraInfo.Left + fraInfo.Width
    End If
    lblInfo.Width = Me.ScaleWidth - lblInfo.Left * 2
    mshPati.Width = Me.ScaleWidth - mshPati.Left * 2
    
    If Me.ScaleHeight - mshPati.Top - mshPati.Left > 1000 Then
        mshPati.Height = Me.ScaleHeight - mshPati.Top - mshPati.Left
    Else
        mshPati.Height = 1000
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lblKind_Click()
    PopupMenu mnuPop, 2
End Sub

Private Sub mshPati_DblClick()
    If mshPati.MouseRow > 0 Then Call mshPati_KeyPress(13)
End Sub

Private Sub mshPati_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Shift = 2 Then cmdSel_Click
End Sub

Private Sub mshPati_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Val(mshPati.TextMatrix(mshPati.Row, 0)) <> 0 Then
        frmDegreeCard.mlng病人ID = Val(mshPati.TextMatrix(mshPati.Row, 0))
        frmDegreeCard.Show 1, Me
    End If
End Sub

Private Sub txtIC卡_GotFocus()
    zlControl.TxtSelAll txtIC卡
End Sub
Private Sub txtValue_GotFocus()
    zlControl.TxtSelAll txtValue
End Sub
Private Sub txtValue_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean
    Dim strKind As String, intKind As Integer, int卡号长度 As Long
    Dim bln密文 As Boolean
    strKind = mnuPopuItems(Val(lblKind.Tag)).Tag
    intKind = Val(lblKind.Tag) + 1
    bln密文 = mcllBrushCard(intKind)(7) <> ""
    txtValue.PasswordChar = IIf(bln密文, "*", "")
    '取缺省的刷卡方式
            '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
            '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
            '第7位后,就只能用索引,不然取不到数
    Select Case strKind
    Case "姓名"
           blnCard = zlCommFun.InputIsCard(txtValue, KeyAscii, mTyCard.bln缺省卡号密文)
           int卡号长度 = mTyCard.int缺省卡号长度 - 1
    Case "门诊号"
            If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            int卡号长度 = 0
    Case "医保号"
            int卡号长度 = 0
    Case Else
        blnCard = zlCommFun.InputIsCard(txtValue, KeyAscii, bln密文)
        int卡号长度 = mcllBrushCard(intKind)(4)
    End Select
    If int卡号长度 > 0 Then
         '刷卡完毕或输入号码后回车
         If blnCard And Len(txtValue.Text) = int卡号长度 - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtValue.Text) <> "" Then
             If KeyAscii <> 13 Then
                 txtValue.Text = txtValue.Text & Chr(KeyAscii)
                 txtValue.SelStart = Len(txtValue.Text)
             End If
             KeyAscii = 0
             Call cmdFind_Click
             If mshPati.Rows > 1 Then
                If mshPati.TextMatrix(1, 0) = "" Then
                   txtValue.SetFocus
                   zlControl.TxtSelAll txtValue
                End If
            End If
        End If
    End If
End Sub

Private Sub txt医保号_GotFocus()
    zlControl.TxtSelAll txt医保号
End Sub

Private Sub txtIC卡_Validate(Cancel As Boolean)
    txtIC卡.Text = UCase(Trim(txtIC卡.Text))
End Sub

Private Sub txtValue_Validate(Cancel As Boolean)
    txtValue.Text = UCase(Trim(txtValue.Text))
End Sub

Private Sub txt医保号_Validate(Cancel As Boolean)
    txt医保号.Text = UCase(Trim(txt医保号.Text))
End Sub


Private Sub txt病人ID_GotFocus()
    zlControl.TxtSelAll txt病人ID
End Sub

Private Sub txt病人ID_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt门诊号_GotFocus()
    zlControl.TxtSelAll txt门诊号
End Sub

Private Sub txt门诊号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtOld_GotFocus()
    zlControl.TxtSelAll txtOld
End Sub

Private Sub txtOld_KeyDown(KeyCode As Integer, Shift As Integer)
    '窗体的keydown中已presskey,这里再处理一次,跳过年龄单位
    If KeyCode = vbKeyReturn And Not IsNumeric(txtOld.Text) Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtOld_Validate(Cancel As Boolean)
    Dim strTmp As String
    
    strTmp = cbo年龄单位.Text
    Select Case strTmp
        Case "岁"
            If Val(txtOld.Text) > 200 Then Cancel = True: Exit Sub
        Case "月"
            If Val(txtOld.Text) > 2400 Then Cancel = True: Exit Sub
        Case "天"
            If Val(txtOld.Text) > 73000 Then Cancel = True: Exit Sub
        Case Else
            Exit Sub
    End Select
End Sub

Private Sub txt身份证_GotFocus()
    zlControl.TxtSelAll txt身份证
    
End Sub

Private Sub txt身份证_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt姓名_GotFocus()
    zlControl.TxtSelAll txt姓名
End Sub

Private Sub mshPati_EnterCell()
    Dim i As Integer, blnPre As Boolean
    Dim intRow As Integer, intCol As Integer
    
    blnPre = mshPati.Redraw
    intRow = mshPati.Row: intCol = mshPati.Col
    mshPati.Redraw = False
    
    For i = 0 To mshPati.Cols - 1
        mshPati.Col = i
        mshPati.CellBackColor = mshPati.BackColorSel
        mshPati.CellForeColor = mshPati.ForeColorSel
    Next
    
    mshPati.Row = intRow:  mshPati.Col = intCol
    mshPati.Redraw = blnPre
End Sub

Private Sub mshPati_LeaveCell()
    Dim i As Integer, blnPre As Boolean
    
    blnPre = mshPati.Redraw
    mshPati.Redraw = False
    
    For i = 0 To mshPati.Cols - 1
        mshPati.Col = i
        mshPati.CellBackColor = mshPati.BackColor
        mshPati.CellForeColor = mshPati.ForeColor
    Next
    mshPati.Redraw = blnPre
End Sub

Private Sub InitMenus()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:动态加载相关的医疗卡类别菜单
    '编制:刘兴洪
    '日期:2011-10-21 15:29:07
    '问题:42315
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant, strKind As String
    Dim i As Integer, ObjItem As Menu, intDefaultKind As Integer
    Set mcllBrushCard = New Collection
    strKind = "就|就诊卡|0|0|18|0|0||"
    If Not gobjSquare.objSquareCard Is Nothing Then
        strKind = gobjSquare.objSquareCard.zlGetIDKindStr(strKind)
    End If
    intDefaultKind = 0
    varData = Split(strKind, ";")
    For i = 0 To UBound(varData)
        Set ObjItem = Me.mnuPopuItems(mnuPopuItems.UBound)
        If Not (ObjItem.Caption = "-" Or Trim(ObjItem.Caption) = "" Or Not ObjItem.Visible) Then
            Load mnuPopuItems(mnuPopuItems.UBound + 1)
            Set ObjItem = mnuPopuItems(mnuPopuItems.UBound)
        End If
        varTemp = Split(varData(i), "|")
        '取缺省的刷卡方式
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
        '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
        '第7位后,就只能用索引,不然取不到数
        mcllBrushCard.Add varTemp, varTemp(1)
        If Val(varTemp(5)) = 1 Then
            intDefaultKind = i
            mTyCard.bln缺省卡号密文 = Trim(varTemp(7)) <> ""
            mTyCard.lng缺省卡类别ID = Val(varTemp(3))
            mTyCard.int缺省卡号长度 = Val(varTemp(4))
        End If
        If i > 9 Then
            ObjItem.Caption = varTemp(1) & IIf(i - 9 > 24, "", "(&" & Chr(64 + i) & ")")
        Else
            ObjItem.Caption = varTemp(1) & "(&" & i & ")"
        End If
        ObjItem.Tag = CStr(varTemp(1))
    Next
    '设置缺省查找对象
    mnuPopuItems_Click (intDefaultKind)
End Sub
Private Sub mnuPopuItems_Click(Index As Integer)
    Dim i As Long
    For i = 0 To mnuPopuItems.UBound
        mnuPopuItems(i).Checked = i = Index
    Next
    lblKind.Caption = mnuPopuItems(Index).Tag & "↓"
    lblKind.Tag = Index
    lblKind.ToolTipText = mnuPopuItems(Index).Tag
    txtValue.ToolTipText = mnuPopuItems(Index).Tag
End Sub
