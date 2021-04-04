VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTendPrintAsk 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "打印选项"
   ClientHeight    =   3150
   ClientLeft      =   2550
   ClientTop       =   2625
   ClientWidth     =   4890
   HelpContextID   =   10322
   Icon            =   "frmTendPrintAsk.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab SSTab1 
      Height          =   2355
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   4154
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "连续打印"
      TabPicture(0)   =   "frmTendPrintAsk.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblNote"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl护理文件"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cbo护理文件"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "重新打印"
      TabPicture(1)   =   "frmTendPrintAsk.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraPrint(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "连续重打"
      TabPicture(2)   =   "frmTendPrintAsk.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraPrint(2)"
      Tab(2).ControlCount=   1
      Begin VB.ComboBox cbo护理文件 
         Height          =   300
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   3015
      End
      Begin VB.Frame fraPrint 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1785
         Index           =   2
         Left            =   -74940
         TabIndex        =   14
         Tag             =   "清除重打"
         Top             =   360
         Width           =   4635
         Begin VB.TextBox txtBegin 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1770
            MaxLength       =   3
            TabIndex        =   8
            Top             =   300
            Width           =   1035
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "没有需要重打的页"
            ForeColor       =   &H000000C0&
            Height          =   375
            Index           =   2
            Left            =   210
            TabIndex        =   17
            Top             =   870
            Width           =   4065
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "从指定的起始页开始进行连续重打，当已打印页的数据修改后引起行错位后使用该功能。"
            ForeColor       =   &H00C00000&
            Height          =   600
            Index           =   1
            Left            =   180
            TabIndex        =   9
            Top             =   1320
            Width           =   4320
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "起始页"
            Height          =   180
            Left            =   1170
            TabIndex        =   7
            Top             =   360
            Width           =   540
         End
      End
      Begin VB.Frame fraPrint 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1935
         Index           =   1
         Left            =   -74970
         TabIndex        =   13
         Tag             =   "重新打印"
         Top             =   360
         Width           =   4635
         Begin VB.TextBox txtPage 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   2025
            MaxLength       =   3
            TabIndex        =   5
            Top             =   285
            Width           =   570
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "打印指定页"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   1050
            TabIndex        =   16
            Top             =   345
            Width           =   900
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "没有需要重打的页"
            ForeColor       =   &H000000C0&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   15
            Top             =   870
            Width           =   4065
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "重新打印指定页号的护理数据。请在数据修改、纸张丢失、破损，或因打印机故障导致打印不成功的情况下使用该功能。"
            ForeColor       =   &H00C00000&
            Height          =   780
            Index           =   0
            Left            =   210
            TabIndex        =   6
            Top             =   1320
            Width           =   4320
         End
      End
      Begin VB.Label lbl护理文件 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "护理文件"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   1
         Top             =   540
         Width           =   720
      End
      Begin VB.Label lblNote 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "当前护理文件一直未打印过"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   1110
         TabIndex        =   3
         Top             =   900
         Width           =   2970
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "当不存在重打数据时才允许使用连续打印功能。"
         ForeColor       =   &H00C00000&
         Height          =   360
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   4320
      End
   End
   Begin VB.CommandButton cmd输出EXCEL 
      Caption         =   "输出到&Excel"
      Height          =   350
      Left            =   180
      TabIndex        =   12
      Top             =   2610
      Width           =   1245
   End
   Begin VB.CommandButton cmd预览 
      Caption         =   "预览(&V)"
      Height          =   350
      Left            =   2370
      TabIndex        =   11
      Top             =   2610
      Width           =   1100
   End
   Begin VB.CommandButton cmd打印 
      Caption         =   "打印(&P)"
      Height          =   350
      Left            =   3570
      TabIndex        =   10
      Top             =   2610
      Width           =   1100
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmTendPrintAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public byRunMode As Byte        '执行方式
Public intPage As Integer       '等待打印页码，续打时不需要处理，只有重打需要记录
Public intPageRows As Integer

Dim strSQL As String
Dim blnRePrint As Boolean       '只需重打修改过的数据
Dim blnRePrintAll As Boolean    '修改过的数据及后面打印的数据都需要重打，只能使用连续重打功能后再使用
Dim strRePrint As String
Dim intBeginPage As Integer
Dim intEndPage As Integer
Dim rsFile As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset


Private Sub SetCommandState()
    '存在重打数据则必须先重打
    cmd预览.Enabled = (intPageRows > 0)
    cmd打印.Enabled = (intPageRows > 0)
    cmd输出EXCEL.Enabled = (intPageRows > 0)
    Select Case SSTab1.Tab
    Case 续打   '只要存在重打数据则不允许使用续打功能
        cmd预览.Enabled = Not blnRePrint And (intPageRows > 0)
        cmd打印.Enabled = Not blnRePrint And (intPageRows > 0)
        cmd输出EXCEL.Enabled = Not blnRePrint And (intPageRows > 0)
    Case 重打   '如果行号变化引起大面积重打，需要使用连续重打功能
        cmd预览.Enabled = Not blnRePrintAll And (intPageRows > 0)
        cmd打印.Enabled = Not blnRePrintAll And (intPageRows > 0)
        cmd输出EXCEL.Enabled = Not blnRePrintAll And (intPageRows > 0)
    Case 连续重打   '连续重打功能用户自己使用，不控制
        cmd预览.Enabled = blnRePrintAll And (intPageRows > 0)
        cmd打印.Enabled = blnRePrintAll And (intPageRows > 0)
        cmd输出EXCEL.Enabled = blnRePrintAll And (intPageRows > 0)
    End Select
End Sub

Private Sub cbo护理文件_Click()
    Dim blnPrint As Boolean
    On Error GoTo errHand
    
    blnRePrint = False
    blnRePrintAll = False
    Call SetCommandState
    
    strSQL = " Select  d.内容文本" & vbNewLine & _
             " From 病历文件结构 d, 病历文件结构 p,病人护理文件 c" & vbNewLine & _
             " Where p.Id = d.父id And p.文件id = c.格式ID and C.ID=[1] And p.对象类型 = 1 And p.内容文本 = '表格样式' and d.要素名称='有效数据行'"
    Set rsTemp = OpenSQLRecord(strSQL, "读取最大数据行", cbo护理文件.ItemData(cbo护理文件.ListIndex))
    If rsTemp.RecordCount <> 0 Then
        intPageRows = NVL(rsTemp!内容文本, 0)
    End If
    
    strSQL = " Select 1 From 病人护理打印 Where 文件ID=[1] And 打印人 is Not NULL And Rownum<2"
    Set rsTemp = OpenSQLRecord(strSQL, "提取护理打印数据", cbo护理文件.ItemData(cbo护理文件.ListIndex))
    blnPrint = rsTemp.RecordCount
    If blnPrint Then
        '已经打印的文件，显示其打印页码
        strSQL = " Select Min(打印页号) AS 开始页号,Max(打印页号) AS 结束页号 From 病人护理打印 Where 文件ID=[1] "
        Set rsTemp = OpenSQLRecord(strSQL, "提取打印页范围", cbo护理文件.ItemData(cbo护理文件.ListIndex))
        intBeginPage = rsTemp!开始页号
        intEndPage = rsTemp!结束页号
        If rsTemp!开始页号 <> rsTemp!结束页号 Then
            lblNote.Caption = "已打印页范围：" & rsTemp!开始页号 & "-" & rsTemp!结束页号
        Else
            lblNote.Caption = "已打印页范围：" & rsTemp!开始页号
        End If
    Else
        lblNote.Caption = "该文件从未打印过！"
    End If
    
    '全院统一编号,用户想怎么打就怎么打,打错了他自己重打
    '检查是否存在需要重打的内容，如存在则不允许续打（提取未打印的最小发生时间的数据，然后检查该时间之后是否存在打印过的数据，如存在则说明存在需要重打的数据）
    strSQL = " Select 行差,打印页号 From 病人护理打印 Where 文件ID=[1] And 打印人 Is NULL And 打印页号 Is Not NULL Order by 打印页号"
    Set rsTemp = OpenSQLRecord(strSQL, "检查是否存在重打内容", cbo护理文件.ItemData(cbo护理文件.ListIndex))
    strRePrint = ""
    If rsTemp.RecordCount <> 0 Then
        blnRePrint = True
        lblTip(重打).Caption = ""
        Do While Not rsTemp.EOF
            If InStr(1, "," & strRePrint & ",", "," & rsTemp!打印页号 & ",") = 0 Then
                strRePrint = strRePrint & "," & rsTemp!打印页号
                lblTip(重打).Caption = lblTip(重打).Caption & "," & rsTemp!打印页号
            End If
            rsTemp.MoveNext
        Loop
        strRePrint = Mid(strRePrint, 2)
        lblTip(重打).Caption = "以下页码需要重打：" & Mid(lblTip(重打).Caption, 2)
        
        rsTemp.Filter = "行差<>0"
        blnRePrintAll = (rsTemp.RecordCount <> 0)
        If blnRePrintAll Then
            txtBegin.Text = rsTemp!打印页号
            lblTip(连续重打).Caption = "由于第" & rsTemp!打印页号 & "页修改后的数据行数发生了变化，从" & rsTemp!打印页号 & "页开始的页面全部需要重打！"
        End If
        rsTemp.Filter = 0
    End If
    
    If Not blnRePrint Then lblTip(重打).Caption = "没有需要重打的页"
    If Not blnRePrintAll Then lblTip(连续重打).Caption = "没有需要连续重打的页"
    
    Call SetCommandState
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function PrePrint() As Boolean
    On Error GoTo errHand
    gintPrintState = SSTab1.Tab + 1
    If SSTab1.Tab = 连续重打 Then
        '完成打印数据的连续重打功能
        If txtBegin.Text = "" Then
            MsgBox "请输入待连续重打的页码！", vbInformation, gstrSysName
            txtBegin.SetFocus
            Exit Function
        End If
        If Not IsNumeric(txtBegin.Text) Then
            MsgBox "输入的页码含有非法字符！", vbInformation, gstrSysName
            txtBegin.SetFocus
            Exit Function
        End If
        intPage = txtBegin.Text
    ElseIf SSTab1.Tab = 重打 Then
        If txtPage(0).Text = "" Then
            MsgBox "请输入重打的页码！", vbInformation, gstrSysName
            txtPage(0).SetFocus
            Exit Function
        End If
        If Not IsNumeric(txtPage(0).Text) Then
            MsgBox "输入的页码含有非法字符！", vbInformation, gstrSysName
            txtPage(0).SetFocus
            Exit Function
        End If
        If txtPage(0).Text < intBeginPage Then
            MsgBox "重打页码不能小于开始页码！", vbInformation, gstrSysName
            txtPage(0).SetFocus
            Exit Function
        End If
        If txtPage(0).Text > intEndPage Then
            MsgBox "重打页码不能大于结束页码！", vbInformation, gstrSysName
            txtPage(0).SetFocus
            Exit Function
        End If
        
        intPage = txtPage(0).Text
    Else
        intPage = 0
    End If
    
    PrePrint = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cmd打印_Click()
    If Not PrePrint Then Exit Sub
    byRunMode = 1
    Me.Hide
End Sub

Private Sub cmd输出EXCEL_Click()
    If Not PrePrint Then Exit Sub
    byRunMode = 3
    Me.Hide
End Sub

Private Sub cmd预览_Click()
    If Not PrePrint Then Exit Sub
    byRunMode = 2
    Me.Hide
End Sub

Private Sub Form_Load()
    '读取所有护理文件
    strSQL = " Select /*+RULE */ A.文件名称 " & vbNewLine & _
             " From 病人护理文件 A" & vbNewLine & _
             " Where A.ID=[1]"
    Set rsFile = OpenSQLRecord(strSQL, "读取所有护理文件", glng文件ID)
    Me.cbo护理文件.AddItem rsFile!文件名称
    Me.cbo护理文件.ItemData(Me.cbo护理文件.NewIndex) = glng文件ID
    cbo护理文件.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    byRunMode = 0
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Call cbo护理文件_Click
End Sub
