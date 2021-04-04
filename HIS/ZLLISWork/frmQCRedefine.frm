VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmQCRedefine 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "重新定值"
   ClientHeight    =   5325
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7215
   Icon            =   "frmQCRedefine.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraEdit 
      Height          =   1095
      Left            =   180
      TabIndex        =   18
      Top             =   4140
      Width           =   6915
      Begin VB.TextBox txt期间 
         Height          =   300
         Left            =   3435
         TabIndex        =   21
         Top             =   705
         Width           =   1500
      End
      Begin VB.CommandButton cmdCusum 
         Cancel          =   -1  'True
         Caption         =   "获取累积值(&S)"
         Height          =   350
         Left            =   5190
         TabIndex        =   12
         Top             =   0
         Width           =   1710
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "∧ 添加到控制值列表中(&A)"
         Height          =   350
         Index           =   0
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   2600
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "∨ 删除最后控制值(&D)"
         Height          =   350
         Index           =   1
         Left            =   2595
         TabIndex        =   14
         Top             =   0
         Width           =   2600
      End
      Begin VB.TextBox txt均值 
         Height          =   300
         Left            =   5040
         TabIndex        =   9
         Top             =   705
         Width           =   810
      End
      Begin VB.TextBox txtSD 
         Height          =   300
         Left            =   5865
         TabIndex        =   11
         Top             =   705
         Width           =   810
      End
      Begin MSComCtl2.DTPicker dtp日期 
         Height          =   300
         Left            =   1755
         TabIndex        =   6
         Top             =   705
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   39714819
         CurrentDate     =   39110
      End
      Begin MSComCtl2.DTPicker dtp开始日期 
         Height          =   300
         Left            =   90
         TabIndex        =   7
         Top             =   705
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   39714819
         CurrentDate     =   39110
      End
      Begin VB.Label lbl期间 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "期间"
         Height          =   180
         Left            =   3480
         TabIndex        =   20
         Top             =   480
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "结束日期"
         Height          =   180
         Left            =   1770
         TabIndex        =   19
         Top             =   480
         Width           =   720
      End
      Begin VB.Label lbl开始日期 
         AutoSize        =   -1  'True
         Caption         =   "开始日期"
         Height          =   180
         Left            =   90
         TabIndex        =   5
         Top             =   495
         Width           =   720
      End
      Begin VB.Label lbl均值 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "均值"
         Height          =   180
         Left            =   5055
         TabIndex        =   8
         Top             =   480
         Width           =   360
      End
      Begin VB.Label lblSD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "标准差"
         Height          =   180
         Left            =   5880
         TabIndex        =   10
         Top             =   480
         Width           =   540
      End
   End
   Begin VB.ComboBox cbo质控品 
      Height          =   300
      Left            =   780
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1050
      Width           =   6240
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -45
      TabIndex        =   17
      Top             =   345
      Width           =   7215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭(&X)"
      Height          =   350
      Left            =   5970
      TabIndex        =   15
      Top             =   420
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgValue 
      Height          =   2655
      Left            =   180
      TabIndex        =   4
      Top             =   1425
      Width           =   6885
      _cx             =   12144
      _cy             =   4683
      Appearance      =   2
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   15790320
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16635590
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   6
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
      OwnerDraw       =   0
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
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Image imgNote 
      Height          =   240
      Left            =   150
      Picture         =   "frmQCRedefine.frx":058A
      Top             =   75
      Width           =   240
   End
   Begin VB.Label lbl质控品 
      AutoSize        =   -1  'True
      Caption         =   "质控品"
      Height          =   180
      Left            =   180
      TabIndex        =   2
      Top             =   1110
      Width           =   540
   End
   Begin VB.Label lbl项目 
      AutoSize        =   -1  'True
      Caption         =   "项目: ####"
      Height          =   180
      Left            =   180
      TabIndex        =   1
      Top             =   787
      Width           =   900
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "设置或调整指定批号质控品的均值和标准差。"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   16
      Top             =   90
      Width           =   3600
   End
   Begin VB.Label lbl仪器 
      AutoSize        =   -1  'True
      Caption         =   "仪器: ####"
      Height          =   180
      Left            =   180
      TabIndex        =   0
      Top             =   465
      Width           =   900
   End
End
Attribute VB_Name = "frmQCRedefine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Enum mCol
    期间 = 0: 日期: 均值: SD: 备注
End Enum

Private mlngDevID As Long       '仪器id
Private mlngItemId As Long      '项目id
Private mdtSysdate As Date      '当前时间
Private mblnModify As Boolean   '是否执行

Dim lngCount As Long

'--------------------------------------------
'以下为窗体公共方法
'--------------------------------------------
Public Function ShowMe(frmParent As Form, lngDevId As Long, lngItemID As Long, Optional dtDefault As Date, Optional lngResId As Long) As Boolean
    '功能：根据指定仪器、项目、质控品，并显示重新定值窗口
    Dim rsTemp As New ADODB.Recordset
    
    mlngDevID = lngDevId
    mlngItemId = lngItemID
    If Not IsNull(dtDefault) Then Me.dtp日期.Value = dtDefault
    
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select 编码 || ', ' || 名称 As 仪器名 From 检验仪器 Where ID = [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngDevID)
    With rsTemp
        If .RecordCount <= 0 Then MsgBox "指定仪器不存在！", vbInformation, gstrSysName: Exit Function
        Me.lbl仪器.Caption = "仪器: " & !仪器名
    End With
    
    gstrSql = "Select 编码 || ', ' || 中文名 || ', ' || 英文名 As 项目名 From 诊治所见项目 Where ID = [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngItemId)
    With rsTemp
        If .RecordCount <= 0 Then MsgBox "指定项目不存在！", vbInformation, gstrSysName: Exit Function
        Me.lbl项目.Caption = "项目: " & !项目名
    End With
    
    gstrSql = "Select Sysdate As 日期, M.ID," & vbNewLine & _
        "       M.批号 || '-' || M.名称 || ' 水平:' || M.水平 ||" & vbNewLine & _
        "        LPad(To_Char(M.开始日期, 'yyyy-MM-dd') || ' ' || To_Char(M.结束日期, 'yyyy-MM-dd'), 200, ' ') As 质控品" & vbNewLine & _
        "From 检验质控品 M, 检验质控品项目 I" & vbNewLine & _
        "Where M.ID = I.质控品id And M.仪器id = [1] And I.项目id = [2] And Trunc(Sysdate) Between M.开始日期 And" & vbNewLine & _
        "      Trim(M.结束日期)" & vbNewLine & _
        "Order By M.开始日期, M.批号"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngDevID, mlngItemId)
    With rsTemp
        Me.cbo质控品.Clear
        Do While Not .EOF
            mdtSysdate = !日期
            Me.cbo质控品.AddItem "" & !质控品
            Me.cbo质控品.ItemData(Me.cbo质控品.NewIndex) = !ID
            .MoveNext
        Loop
        If Me.cbo质控品.ListCount = 0 Then MsgBox "该仪器项目无当前有效的质控品！", vbInformation, gstrSysName: Exit Function
        For lngCount = 0 To Me.cbo质控品.ListCount - 1
            If Me.cbo质控品.ItemData(Me.cbo质控品.NewIndex) = lngResId Then Me.cbo质控品.ListIndex = Me.cbo质控品.NewIndex: Exit For
        Next
        If Me.cbo质控品.ListIndex = -1 Then Me.cbo质控品.ListIndex = 0
    End With
    
    mblnModify = False
    Me.Show vbModal, frmParent
    ShowMe = mblnModify
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cbo质控品_Click()
    Dim rsTemp As New ADODB.Recordset
    
    Dim strPeriod As String
    
    If Me.cbo质控品.ListIndex = -1 Then Exit Sub
    
    gstrSql = "Select 开始日期 From 检验质控品 Where Id= [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(Me.cbo质控品.ItemData(Me.cbo质控品.ListIndex)))
    Do Until rsTemp.EOF
        Me.dtp开始日期.MinDate = rsTemp!开始日期
        rsTemp.MoveNext
    Loop
    
    gstrSql = "Select 期间, To_Char(开始日期, 'yyyy-MM-dd') As 开始日期, 均值, Sd," & vbNewLine & _
            "       设置人 || '在'|| To_Char(设置日期, 'yyyy-MM-dd') || '设置' As 备注" & vbNewLine & _
            "From 检验质控均值" & vbNewLine & _
            "Where 质控品id = [1] And 项目id = [2]" & vbNewLine & _
            "Order By 开始日期"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(Me.cbo质控品.ItemData(Me.cbo质控品.ListIndex)), mlngItemId)
    With Me.vfgValue
        .Clear
        Set .DataSource = rsTemp
        .ColWidth(mCol.日期) = 1100
        For lngCount = .FixedCols To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
        Next
        For lngCount = .FixedRows To .Rows - 1
            If Left(.TextMatrix(lngCount, mCol.均值), 1) = "." Then .TextMatrix(lngCount, mCol.均值) = "0" & .TextMatrix(lngCount, mCol.均值)
            If Left(.TextMatrix(lngCount, mCol.SD), 1) = "." Then .TextMatrix(lngCount, mCol.SD) = "0" & .TextMatrix(lngCount, mCol.SD)
        Next
        If .Rows > .FixedRows Then .Row = .Rows - 1
    End With
    
    strPeriod = Right(Me.cbo质控品.Text, 21)
    Me.dtp日期.MinDate = CDate(Left(strPeriod, 10))
    
    If CDate(Right(strPeriod, 10)) < mdtSysdate Then
        Me.dtp日期.MaxDate = CDate(Right(strPeriod, 10))
    Else
        Me.dtp日期.MaxDate = mdtSysdate
    End If
    Me.dtp开始日期.MaxDate = Me.dtp日期.MaxDate
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCusum_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim strAsk As String
    
    gstrSql = "Select Round(Avg(结果), 2) As 均值, Round(Stddev(结果), 2) As Sd, Count(*) As 次数" & vbNewLine & _
            "From (Select Trunc(Q.检验时间) As 日期," & vbNewLine & _
            "              Avg(zl_Lis_tonumber(Q.质控品ID,R.检验项目id,R.检验结果,R.ID)) As 结果" & vbNewLine & _
            "       From 检验质控记录 Q, 检验普通结果 R,检验质控报告 T" & vbNewLine & _
            "       Where Q.标本id = R.检验标本id And Q.质控品id = [1] And R.检验项目id + 0 = [2] And" & vbNewLine & _
            "             Nvl(R.弃用结果,0)=0 And R.ID=T.结果ID(+) And Q.检验时间 + 0 between To_Date([3], 'yyyy-MM-dd') And  To_Date([4], 'yyyy-MM-dd')+ 1 And Nvl(T.标记, 0) <> 2" & vbNewLine & _
            "       Group By Trunc(Q.检验时间))"

'    gstrSql = "Select Round(Avg(结果), 2) As 均值, Round(Stddev(结果), 2) As Sd, Count(*) As 次数" & vbNewLine & _
'            "From (Select Trunc(Q.检验时间) As 日期," & vbNewLine & _
'            "              Avg(zl_Lis_tonumber(Q.质控品ID,R.检验项目id,R.检验结果)) As 结果" & vbNewLine & _
'            "       From 检验质控记录 Q, 检验普通结果 R" & vbNewLine & _
'            "       Where Q.标本id = R.检验标本id And Q.质控品id = [1] And R.检验项目id + 0 = [2] And" & vbNewLine & _
'            "             Q.检验时间 + 0 <To_Date([4], 'yyyy-MM-dd')+ 1 And Nvl(Q.弃用记录, 0) = 0" & vbNewLine & _
'            "       Group By Trunc(Q.检验时间))"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, _
                    CLng(Me.cbo质控品.ItemData(Me.cbo质控品.ListIndex)), _
                    mlngItemId, Format(Me.dtp开始日期.Value, "yyyy-MM-dd"), Format(Me.dtp日期.Value, "yyyy-MM-dd"))
    If rsTemp.RecordCount <= 0 Then MsgBox "无法获取累计均值和标准差", vbInformation, gstrSysName: Exit Sub
    
    strAsk = "该质控品在" & Format(Me.dtp开始日期.Value, "yyyy年mm月dd日") & "至" & Format(Me.dtp日期.Value, "yyyy年mm月dd日") & "的累计值如下："
    strAsk = strAsk & vbCrLf & "   均值: " & IIf(Left(rsTemp!均值, 1) = ".", "0", "") & rsTemp!均值
    strAsk = strAsk & vbCrLf & "   SD值: " & IIf(Left(rsTemp!SD, 1) = ".", "0", "") & rsTemp!SD
    If rsTemp!次数 < 20 Then
        strAsk = strAsk & vbCrLf & "由于有效质控记录不到20次，不适宜直接将累计值作为控制值。"
    End If
    strAsk = strAsk & vbCrLf & vbCrLf & "要使用该累计值吗？"
    If MsgBox(strAsk, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        Me.txt均值.Text = IIf(Left(rsTemp!均值, 1) = ".", "0", "") & rsTemp!均值
        Me.txtSD.Text = IIf(Left(rsTemp!SD, 1) = ".", "0", "") & rsTemp!SD
        Me.txt期间.Text = Format(Me.dtp开始日期.Value, "yyyyMM")
    End If
    
    Me.txt均值.SetFocus
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdEdit_Click(Index As Integer)
    If Index = 0 Then
        With Me.vfgValue
            If .Rows > .FixedRows Then
                If .TextMatrix(.Rows - 1, mCol.日期) >= Format(Me.dtp日期.Value, "yyyy-MM-dd") Then
                    MsgBox "新控制值的结束日期必须大于上次日期！", vbInformation, gstrSysName
                    Me.dtp日期.SetFocus: Exit Sub
                End If
            End If
            If Val(Trim(Me.txt均值.Text)) = 0 And Val(Trim(Me.txtSD.Text)) = 0 Then
                MsgBox "必须同时设置均值(x)和标准差(SD)！", vbInformation, gstrSysName
                Me.txt均值.SetFocus: Exit Sub
            End If
        End With
        gstrSql = "Zl_检验质控均值_Edit(1," & Me.cbo质控品.ItemData(Me.cbo质控品.ListIndex) & "," & mlngItemId
        gstrSql = gstrSql & ",To_Date('" & Format(Me.dtp日期.Value, "yyyy-MM-dd") & "','yyyy-MM-dd')"
        gstrSql = gstrSql & "," & Val(Trim(Me.txt均值.Text)) & "," & Val(Trim(Me.txtSD.Text)) & ",'" & Replace(Trim(txt期间.Text), "'", "") & "')"
    Else
        With Me.vfgValue
            If .Rows <= .FixedRows Then
                MsgBox "已经没有控制值！", vbInformation, gstrSysName: Me.dtp日期.SetFocus: Exit Sub
            ElseIf .Rows = .FixedRows + 1 Then
                If DateAdd("m", 1, CDate(.TextMatrix(.Rows - 1, mCol.日期))) <= Me.dtp日期.MaxDate And DateAdd("m", 1, CDate(.TextMatrix(.Rows - 1, mCol.日期))) >= Me.dtp日期.MinDate Then
                    Me.dtp日期.Value = DateAdd("m", 1, CDate(.TextMatrix(.Rows - 1, mCol.日期)))
                End If
            Else
                If CDate(.TextMatrix(.Rows - 1, mCol.日期)) <= Me.dtp日期.MaxDate And CDate(.TextMatrix(.Rows - 1, mCol.日期)) >= Me.dtp日期.MinDate Then
                    Me.dtp日期.Value = CDate(.TextMatrix(.Rows - 1, mCol.日期))
                End If
            End If
            Me.txt均值.Text = Val(.TextMatrix(.Rows - 1, mCol.均值))
            Me.txtSD.Text = Val(.TextMatrix(.Rows - 1, mCol.SD))
        End With
        gstrSql = "Zl_检验质控均值_Edit(2," & Me.cbo质控品.ItemData(Me.cbo质控品.ListIndex) & "," & mlngItemId & ")"
    End If
    
    Err = 0: On Error GoTo ErrHand
    zldatabase.ExecuteProcedure gstrSql, Me.Caption
    mblnModify = True
    
    Call cbo质控品_Click
    If Index = 0 Then
        Me.vfgValue.SetFocus
    Else
        Me.dtp日期.SetFocus
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

