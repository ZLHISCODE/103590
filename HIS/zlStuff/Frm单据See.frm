VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form Frm单据See 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "单据查询"
   ClientHeight    =   6690
   ClientLeft      =   120
   ClientTop       =   2190
   ClientWidth     =   11730
   Icon            =   "Frm单据See.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   11730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Cmd取消 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   10515
      TabIndex        =   5
      Top             =   5880
      Width           =   1100
   End
   Begin VB.PictureBox Pic单据 
      BackColor       =   &H80000004&
      Height          =   5715
      Left            =   0
      ScaleHeight     =   5655
      ScaleWidth      =   11655
      TabIndex        =   9
      Top             =   0
      Width           =   11715
      Begin ZL9BillEdit.BillEdit msf 
         Height          =   2895
         Left            =   240
         TabIndex        =   46
         Top             =   1320
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   5106
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin VB.TextBox Txt开单人 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   765
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   4725
         Width           =   1005
      End
      Begin VB.TextBox Txt发药人 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   10155
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   4725
         Width           =   1005
      End
      Begin VB.TextBox Txt记费人 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   5745
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   4725
         Width           =   1005
      End
      Begin VB.ComboBox Cbo科室 
         Height          =   300
         Left            =   7140
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   765
         Width           =   1515
      End
      Begin VB.ComboBox Cbo性别 
         Height          =   300
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   795
         Width           =   915
      End
      Begin VB.TextBox Txt年龄 
         Height          =   285
         Left            =   5490
         MaxLength       =   3
         TabIndex        =   31
         Text            =   "20"
         Top             =   795
         Width           =   435
      End
      Begin VB.TextBox Txt住院号 
         Height          =   270
         Left            =   975
         TabIndex        =   29
         Top             =   435
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.TextBox Txt姓名 
         Height          =   270
         Left            =   975
         TabIndex        =   28
         Top             =   825
         Width           =   1365
      End
      Begin VB.ComboBox Cbo部门 
         Height          =   300
         Left            =   9000
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   780
         Width           =   2430
      End
      Begin VB.TextBox Txt发票号 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   5565
         MaxLength       =   8
         TabIndex        =   2
         Top             =   780
         Width           =   1035
      End
      Begin VB.TextBox TxtMoney 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   210
         TabIndex        =   24
         Top             =   4350
         Width           =   11235
      End
      Begin VB.ComboBox Cbo入出类别 
         Enabled         =   0   'False
         Height          =   300
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   420
         Width           =   1635
      End
      Begin VB.TextBox TxtNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10020
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   330
         Width           =   1425
      End
      Begin VB.TextBox Txt填制日期 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   5160
         Width           =   1395
      End
      Begin VB.TextBox Txt审核日期 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   10050
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   5160
         Width           =   1395
      End
      Begin VB.TextBox Txt审核人 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   7950
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   5160
         Width           =   1005
      End
      Begin VB.TextBox Txt填制人 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   5160
         Width           =   1005
      End
      Begin VB.TextBox Txt摘要 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   765
         TabIndex        =   4
         Top             =   4740
         Width           =   10665
      End
      Begin VB.CommandButton Cmd供药单位 
         Caption         =   "…"
         Height          =   285
         Left            =   11190
         TabIndex        =   20
         Top             =   780
         Width           =   255
      End
      Begin VB.TextBox Txt供药单位 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   9000
         MaxLength       =   50
         TabIndex        =   3
         Top             =   780
         Width           =   2205
      End
      Begin VB.ComboBox Cbo库房 
         Enabled         =   0   'False
         Height          =   300
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   810
         Width           =   1650
      End
      Begin MSComCtl2.DTPicker Dtp日期 
         Height          =   285
         Left            =   9750
         TabIndex        =   30
         Top             =   795
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   129761283
         CurrentDate     =   36471
      End
      Begin VB.Label Lbl开单人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开单人"
         Height          =   180
         Left            =   165
         TabIndex        =   45
         Top             =   4800
         Width           =   540
      End
      Begin VB.Label Lbl记费人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "记费人"
         Height          =   180
         Left            =   5145
         TabIndex        =   44
         Top             =   4800
         Width           =   540
      End
      Begin VB.Label Lbl发药人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发药人"
         Height          =   180
         Left            =   9570
         TabIndex        =   43
         Top             =   4800
         Width           =   540
      End
      Begin VB.Label Lbl科室 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "科室"
         Height          =   180
         Left            =   6705
         TabIndex        =   39
         Top             =   825
         Width           =   360
      End
      Begin VB.Label Lbl病人姓名 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "姓    名"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   225
         TabIndex        =   38
         Top             =   870
         Width           =   720
      End
      Begin VB.Label Lbl性别 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2790
         TabIndex        =   37
         Top             =   855
         Width           =   360
      End
      Begin VB.Label Lbl年龄 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5010
         TabIndex        =   36
         Top             =   855
         Width           =   360
      End
      Begin VB.Label Lbl日期 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "日期"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   9300
         TabIndex        =   35
         Top             =   855
         Width           =   360
      End
      Begin VB.Label Lbl住院号 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "住 院 号"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   210
         TabIndex        =   34
         Top             =   495
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Lbl部门 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "部门"
         Height          =   180
         Left            =   8565
         TabIndex        =   27
         Top             =   855
         Width           =   360
      End
      Begin VB.Label Lbl发票号 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发票号"
         Height          =   180
         Left            =   4920
         TabIndex        =   25
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Lbl入出类别 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入出类别"
         Height          =   180
         Left            =   225
         TabIndex        =   23
         Top             =   480
         Width           =   720
      End
      Begin VB.Label LblNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9675
         TabIndex        =   21
         Top             =   330
         Width           =   240
      End
      Begin VB.Label Lbl审核日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核日期"
         Height          =   180
         Left            =   9240
         TabIndex        =   17
         Top             =   5220
         Width           =   720
      End
      Begin VB.Label Lbl审核人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核人"
         Height          =   180
         Left            =   7365
         TabIndex        =   16
         Top             =   5220
         Width           =   540
      End
      Begin VB.Label Lbl填制日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "填制日期"
         Height          =   180
         Left            =   2040
         TabIndex        =   15
         Top             =   5220
         Width           =   720
      End
      Begin VB.Label Lbl填制人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "填制人"
         Height          =   180
         Left            =   180
         TabIndex        =   14
         Top             =   5220
         Width           =   540
      End
      Begin VB.Label Lbl摘要 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "摘  要"
         Height          =   180
         Left            =   180
         TabIndex        =   13
         Top             =   4815
         Width           =   540
      End
      Begin VB.Label Lbl供药单位 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "供药单位"
         Height          =   180
         Left            =   8205
         TabIndex        =   12
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Lbl库房 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "库    房"
         Height          =   180
         Left            =   210
         TabIndex        =   11
         Top             =   885
         Width           =   720
      End
      Begin VB.Label Lbl标题 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "单据查询"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   315
         TabIndex        =   10
         Top             =   150
         Width           =   11535
      End
   End
   Begin MSComctlLib.StatusBar Sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   6324
      Width           =   11736
      _ExtentX        =   20690
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "Frm单据See.frx":030A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15372
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1429
            MinWidth        =   1411
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
            Object.ToolTipText     =   "当前数字键状态"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
            Object.ToolTipText     =   "当前大写键状态"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Frm单据See"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strNo As String               '输入参数:单据号
Public int记录状态 As Integer        '输入参数:记录状态
Public byt单据 As Byte               '输入参数:单据标志:24-收费处方；25-记帐单处方；26-记帐表处方
Private blnFirst As Boolean
Private UnitLevel As Integer '单位级数

'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------


Private Sub RefreshHead()
    '-----------------------------------------------------
    '--功能:刷新表头结构
    '--参数:byt单据
    '          24-收费处方
    '          25-记帐单处方
    '          26-记帐表处方
    '--返回:
    '-----------------------------------------------------
    With msf
        Select Case byt单据
            Case 24, 25             '处方
                
                Me.Caption = "卫材处方单"
                Me.Lbl标题 = "卫材处方单"
               
                .Cols = 10
                .TextMatrix(0, 0) = "名称与编码"
                .TextMatrix(0, 1) = "规格"
                .TextMatrix(0, 2) = "产地"
                .TextMatrix(0, 3) = "单位"
                .TextMatrix(0, 4) = "数量"
                .TextMatrix(0, 5) = "付数"
                .TextMatrix(0, 6) = "售价"
                .TextMatrix(0, 7) = "金额"
                .TextMatrix(0, 8) = "成本金额"
                .TextMatrix(0, 9) = "差价"
                
                .ColWidth(0) = 2500
                .ColWidth(1) = 800
                .ColWidth(2) = 1000
                .ColWidth(3) = 500
                .ColWidth(4) = 1100
                .ColWidth(5) = 800
                .ColWidth(6) = 1000
                .ColWidth(7) = 800
                .ColWidth(8) = 800
                .ColWidth(9) = 800
                
                .ColData(0) = 1
                .ColData(1) = 5
                .ColData(2) = 5
                .ColData(3) = 5
                .ColData(4) = 4
                .ColData(5) = 5
                .ColData(6) = 5
                .ColData(7) = 5
                .ColData(8) = 5
            Case 26                     '摆药单
            
                Me.Caption = "记帐处方单"
                Me.Lbl标题 = "记帐处方单"
               
                .Cols = 11
                .TextMatrix(0, 0) = "科室"
                .TextMatrix(0, 1) = "病人姓名"
                .TextMatrix(0, 2) = "名称与编码"
                .TextMatrix(0, 3) = "规格"
                .TextMatrix(0, 4) = "产地"
                .TextMatrix(0, 5) = "单位"
                .TextMatrix(0, 6) = "数量"
                .TextMatrix(0, 7) = "售价"
                .TextMatrix(0, 8) = "金额"
                .TextMatrix(0, 9) = "成本金额"
                .TextMatrix(0, 10) = "差价"
                
                .ColWidth(0) = 1000
                .ColWidth(1) = 1000
                .ColWidth(2) = 2500
                .ColWidth(3) = 800
                .ColWidth(4) = 1000
                .ColWidth(5) = 500
                .ColWidth(6) = 1100
                .ColWidth(7) = 1000
                .ColWidth(8) = 800
                .ColWidth(9) = 800
                .ColWidth(10) = 800
                
                .ColData(0) = 4
                .ColData(1) = 4
                .ColData(2) = 1
                .ColData(3) = 5
                .ColData(4) = 5
                .ColData(5) = 5
                .ColData(6) = 4
                .ColData(7) = 5
                .ColData(8) = 5
                .ColData(9) = 5
                .ColData(10) = 0
                .PrimaryCol = 1
                .LocateCol = 9
        
        End Select
    End With
End Sub

Private Sub RefreshHeadData()
    '-----------------------------------------------------
    '--功能:刷新表上项或表下项数据
    '--参数:
    '     byt单据:
    '          24-收费处方
    '          25-记帐单处方
    '          26-记帐表处方
    '    strNO
    '--返回:
    '-----------------------------------------------------
    Dim RsHead As New ADODB.Recordset
    Dim strSql门诊 As String
    
    On Error GoTo errHandle
    Select Case byt单据
        Case 24, 25          '处方
            If byt单据 = 24 Then
                gstrSQL = " Select A.NO,B.名称 as 开单科室,G.姓名,G.性别,G.年龄,0 As 住院号,'' 床号,A.摘要,A.填制人," & _
                " To_Char(A.填制日期,'yyyy-MM-dd') as 填制日期,A.审核人,To_Char(A.审核日期,'yyyy-MM-dd') as 审核日期," & _
                " G.病人ID,0 主页ID,A.对方部门ID as 开单部门ID,G.操作员姓名 as 计费人,A.单据,A.ID,G.开单人 " & _
                          " From 药品收发记录 A,部门表 B,门诊费用记录 G " & _
                          " Where A.对方部门id=B.id(+) And A.单据=[2] And A.NO=G.NO(+) And A.费用ID=G.ID And A.NO=[1] And Rownum < 2 "
            Else
                gstrSQL = " Select A.NO,B.名称 as 开单科室,G.姓名,G.性别,G.年龄,C.住院号,G.床号,A.摘要,A.填制人," & _
                " To_Char(A.填制日期,'yyyy-MM-dd') as 填制日期,A.审核人,To_Char(A.审核日期,'yyyy-MM-dd') as 审核日期," & _
                " G.病人ID,G.主页ID,A.对方部门ID as 开单部门ID,G.操作员姓名 as 计费人,A.单据,A.ID,G.开单人 " & _
                          " From 药品收发记录 A,部门表 B,门诊费用记录 G,病人信息 C " & _
                          " Where A.对方部门id=B.id(+) And A.单据=[2] And C.病人id=G.病人id And A.费用ID=G.ID And A.NO=G.NO(+) And A.NO=[1] And Rownum < 2 "
                strSql门诊 = Replace(gstrSQL, "G.床号", "'' 床号")
                strSql门诊 = Replace(strSql门诊, "G.主页ID", "0 主页ID")
                gstrSQL = strSql门诊 & " Union All " & Replace(gstrSQL, "门诊费用记录", "住院费用记录")
            End If
        Case 26
            gstrSQL = " Select A.NO,A.摘要,A.填制人,To_Char(A.填制日期,'yyyy-MM-dd') as 填制日期,A.审核人,To_Char(A.审核日期,'yyyy-MM-dd') as 审核日期,G.执行部门ID,A.对方部门ID as 开单部门ID,G.操作员姓名 as 计费人,A.单据,A.ID" & _
                    " From 药品收发记录 A,部门表 B,住院费用记录 G " & _
                    " Where A.对方部门id=B.id(+) And A.单据=[2] And A.NO=G.NO(+) And A.费用ID=G.ID And A.NO=[1] And Rownum < 2 "
    End Select
    
    Set RsHead = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, byt单据)
    
    With RsHead
        If .RecordCount = 0 Then Exit Sub
        '填制表头数据
        Select Case byt单据
            Case 24, 25      '处方
                Me.TxtNo = !NO
                Me.Txt开单人 = !开单人
                Me.Txt记费人 = IIf(IsNull(!计费人), " ", !计费人)
                Me.Dtp日期 = !填制日期
                Lbl日期.Caption = "填制日期"
                
                Me.Txt姓名 = IIf(IsNull(!姓名), " ", !姓名)
                Me.Cbo科室.AddItem IIf(IsNull(!开单科室), "", !开单科室)
                Me.Cbo科室.ListIndex = Me.Cbo科室.NewIndex
                
                Me.Cbo性别.AddItem IIf(IsNull(!性别), "", !性别)
                Cbo性别.ListIndex = Cbo性别.NewIndex
                
                Me.Txt年龄 = IIf(IsNull(!年龄), 20, !年龄)
                Me.Txt住院号 = IIf(IsNull(!住院号), " ", !住院号)
                Me.Txt发药人 = IIf(IsNull(!审核人), " ", !审核人)
           
            Case 26                     '摆药单
                Me.TxtNo = !NO
                Me.Txt开单人 = !填制人
                Me.Txt记费人 = IIf(IsNull(!计费人), " ", !计费人)
                Me.Dtp日期 = !填制日期
        End Select
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefreshBodyData()
    '-----------------------------------------------------
    '--功能:刷新表体内容数据
    '--参数:
    '     byt单据:
    '          24-收费处方
    '          25-记帐单处方
    '          26-记帐表处方
    '    strNO
    '--返回:
    '-----------------------------------------------------
    Dim RsBody As New ADODB.Recordset
    Dim intRow As Long
    Dim ii As Long
    On Error GoTo errHandle
    Select Case byt单据
        Case 24        '收费处方
            gstrSQL = "SELECT DISTINCT '【'||F.编码||'】'||NVL(E.名称,F.名称) AS 材料信息,F.规格,F.产地,F.计算单位 AS 单位," & _
                " B.换算系数,A.实际数量 AS 数量,A.零售价 AS 单价," & _
                " A.零售金额,A.付数,A.成本金额,A.差价,A.药品ID,A.库房ID,A.入出系数,A.序号,A.ID,B.最大效期,B.指导零售价,B.指导差价率" & _
                " FROM 药品收发记录 A,材料特性 B,收费项目别名 E,收费项目目录 F " & _
                " WHERE A.药品ID=B.材料ID AND B.材料ID=F.ID " & _
                " AND B.材料ID=E.收费细目ID(+) AND E.性质(+)=3 AND E.码类(+)=1" & _
                " AND A.记录状态=[3] AND A.单据=[2] AND A.NO =[1]" & _
                " ORDER BY A.序号"
        Case 25      '记帐单处方
            gstrSQL = "" & _
                "   SELECT DISTINCT '【'||F.编码||'】'||NVL(E.名称,F.名称) AS 材料信息,F.规格,F.产地,F.计算单位 AS 单位," & _
                "       b.换算系数,A.实际数量 AS 数量,A.零售价 AS 单价," & _
                "       A.零售金额,A.付数,A.成本金额,A.差价,A.药品ID,A.库房ID,A.入出系数,A.序号,A.ID,B.最大效期,B.指导零售价,B.指导差价率" & _
                " FROM 药品收发记录 A,材料特性 B,收费项目别名 E,收费项目目录 F " & _
                " WHERE A.药品ID=B.材料ID AND B.材料ID=F.ID " & _
                " AND B.材料ID=E.收费细目ID(+) AND E.性质(+)=3 AND E.码类(+)=1" & _
                " AND A.单据=[2] AND A.记录状态=[3] AND A.NO =[1]" & _
                " ORDER BY A.序号"
        Case 26        '记帐表处方
            gstrSQL = "" & _
                "   SELECT DISTINCT P.名称 科室,G.姓名,G.性别,G.年龄,G.病人ID,G.主页ID,D.住院号,G.床号," & _
                "       '【'||F.编码||'】'||NVL(C.名称,F.名称) AS 材料信息,F.规格,F.产地," & _
                "       B.换算系数,A.实际数量 AS 数量,F.计算单位 AS 单位,A.零售价  AS 单价," & _
                "       A.零售金额,A.付数,A.成本金额,A.差价,A.药品ID,A.库房ID,A.入出系数,A.序号,A.ID,B.最大效期,B.指导零售价,B.指导差价率" & _
                " FROM 药品收发记录 A,材料特性 B,住院费用记录 G,病人信息 D,收费项目别名 C,收费项目目录 F,部门表 P " & _
                " WHERE A.NO=G.NO(+) AND A.费用ID=G.ID AND A.对方部门ID=P.ID AND A.药品ID=B.材料ID AND B.材料ID=F.ID AND G.病人ID=D.病人ID " & _
                " AND B.材料ID=C.收费细目ID(+) AND C.性质(+)=3 AND C.码类(+)=1 " & _
                " AND A.单据=[2] AND A.记录状态=[3] AND A.NO =[1]" & _
                " ORDER BY A.序号"
    End Select
    Set RsBody = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, byt单据, int记录状态)
    
    With RsBody
       If .RecordCount <> 0 Then msf.Rows = .RecordCount + 1
        
        '填制表头数据
        Select Case byt单据
            Case 24, 25          '处方
                    For intRow = 1 To .RecordCount
                        msf.TextMatrix(intRow, 0) = !材料信息
                        msf.TextMatrix(intRow, 1) = IIf(IsNull(!规格), "", !规格)
                        msf.TextMatrix(intRow, 2) = IIf(IsNull(!产地), "", !产地)
                        msf.TextMatrix(intRow, 3) = !单位
                        msf.TextMatrix(intRow, 4) = Format(!数量, mFMT.FM_数量)
                        msf.TextMatrix(intRow, 5) = Format(!付数, mFMT.FM_数量)
                        msf.TextMatrix(intRow, 6) = Format(!单价, mFMT.FM_零售价)
                        msf.TextMatrix(intRow, 7) = Format(!零售金额, mFMT.FM_金额)
                        msf.TextMatrix(intRow, 8) = Format(!成本金额, mFMT.FM_金额)
                        msf.TextMatrix(intRow, 9) = Format(!差价, mFMT.FM_金额)
                        msf.RowData(intRow) = !序号
                        .MoveNext
                    Next
            Case 26         '摆药
                    For intRow = 1 To .RecordCount
                        msf.TextMatrix(intRow, 0) = !科室
                        msf.TextMatrix(intRow, 1) = !姓名
                        msf.TextMatrix(intRow, 2) = !材料信息
                        msf.TextMatrix(intRow, 3) = IIf(IsNull(!规格), "", !规格)
                        msf.TextMatrix(intRow, 4) = IIf(IsNull(!产地), "", !产地)
                        msf.TextMatrix(intRow, 5) = !单位
                        msf.TextMatrix(intRow, 6) = Format(!数量, mFMT.FM_数量)
                        msf.TextMatrix(intRow, 7) = Format(!单价, mFMT.FM_零售价)
                        msf.TextMatrix(intRow, 8) = Format(!零售金额, mFMT.FM_金额)
                        msf.TextMatrix(intRow, 9) = Format(!成本金额, mFMT.FM_金额)
                        msf.TextMatrix(intRow, 10) = Format(!差价, mFMT.FM_金额)
                        msf.RowData(intRow) = !序号
                        .MoveNext
                    Next
        End Select
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Cmd取消_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If Not blnFirst Then Exit Sub
    SetEV
    blnFirst = False
    Me.TxtNo = strNo
    '刷新表头
    RefreshHead
    '刷新表头数据
    RefreshHeadData
    '刷新表体内容
    
    RefreshBodyData
    '显示合计金额
    SumDataMSf
    LockCons

End Sub

Private Sub Form_Load()
    blnFirst = True
  
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(0, g_成本价)
        .FM_金额 = GetFmtString(0, g_金额)
        .FM_零售价 = GetFmtString(0, g_售价)
        .FM_数量 = GetFmtString(0, g_数量)
    End With
    
    
    SetEV
    RestoreWinState Me, App.ProductName, Me.Caption
End Sub

Private Function LockCons()
    Me.Cbo入出类别.Enabled = False
    Me.Cbo库房.Enabled = False
    Me.Txt发票号.Enabled = False
    Me.Txt供药单位.Enabled = False
    Me.Cmd供药单位.Enabled = False
    Me.msf.Active = False
    Me.Cbo部门.Enabled = False
    Me.Txt摘要.Enabled = False
    Me.Txt填制人.Enabled = False
    Me.Txt审核人.Enabled = False
    Me.Txt填制日期.Enabled = False
    Me.Txt审核日期.Enabled = False
    Me.Txt姓名.Enabled = False
    Me.Txt年龄.Enabled = False
    Me.Cbo科室.Enabled = False
    Me.Dtp日期.Enabled = False
    Me.Txt住院号.Enabled = False
    Me.Cbo性别.Enabled = False
    Me.Txt发药人.Enabled = False
    Me.Txt记费人.Enabled = False
    Me.Txt开单人.Enabled = False
End Function

Private Sub SumDataMSf()
    '-------------------------------------------------------------
    '--功能:对各单据进行金额计算
    '--参数:
    '       byt单据:
    '          24-收费处方
    '          25-记帐单处方
    '          26-记帐表处方
    '-- 返回:
    '------------------------------------------------------------
    Dim curTotal As Double, Cur记帐金额 As Double, Cur记帐差价 As Double
    Dim intLop As Long
    curTotal = 0: Cur记帐金额 = 0: Cur记帐差价 = 0: TxtMoney = ""
    
    Select Case byt单据
    Case 24, 25                      ' 直接输入处方
        For intLop = 1 To msf.Rows - 1
            curTotal = curTotal + Val(msf.TextMatrix(intLop, 7))
            Cur记帐金额 = Cur记帐金额 + Val(msf.TextMatrix(intLop, 8))
            Cur记帐差价 = Cur记帐差价 + Val(msf.TextMatrix(intLop, 9))
        Next
        TxtMoney = "金额合计：" & Format(curTotal, mFMT.FM_金额) & Space(10) & "记帐金额合计：" & Format(Cur记帐金额, mFMT.FM_金额) & Space(10) & "记帐差价合计：" & Format(Cur记帐差价, mFMT.FM_金额)
    Case 26                          ' 直接输入摆药单
        For intLop = 1 To msf.Rows - 1
            curTotal = curTotal + Val(msf.TextMatrix(intLop, 7))
            Cur记帐金额 = Cur记帐金额 + Val(msf.TextMatrix(intLop, 8))
            Cur记帐差价 = Cur记帐差价 + Val(msf.TextMatrix(intLop, 9))
        Next
        TxtMoney = "金额合计：" & Format(curTotal, mFMT.FM_金额) & Space(10) & "记帐金额合计：" & Format(Cur记帐金额, mFMT.FM_金额) & Space(10) & "记帐差价合计：" & Format(Cur记帐差价, mFMT.FM_金额)
    End Select
End Sub

Private Function SetEV()
    '-------------------------------------------------------------
    '--功能:设置控件的Visible属性
    '--参数:
    '       byt单据:
    '          24-收费处方
    '          25-记帐单处方
    '          26-记帐表处方
    '-- 返回:
    '------------------------------------------------------------
    Me.Lbl入出类别.Visible = True
    Me.Cbo入出类别.Visible = True
    Me.Lbl库房.Caption = "库    房"
    
    Me.Lbl库房.Visible = True
    Me.Cbo库房.Visible = True
    Me.Lbl发票号.Visible = True
    Me.Txt发票号.Visible = True
    Me.Lbl供药单位.Visible = True
    Me.Txt供药单位.Visible = True
    Me.Cmd供药单位.Visible = True
    Me.Lbl部门.Visible = True
    Me.Cbo部门.Visible = True
    Me.Lbl摘要.Visible = True
    Me.Txt摘要.Visible = True
    Me.Lbl填制人.Visible = True
    Me.Txt填制人.Visible = True
    Me.Lbl审核人.Visible = True
    Me.Txt审核人.Visible = True
    Me.Lbl填制日期.Visible = True
    Me.Txt填制日期.Visible = True
    Me.Lbl审核人.Visible = True
    Me.Lbl审核日期.Visible = True
    Me.Txt审核日期.Visible = True
    Me.Lbl病人姓名.Visible = True
    Me.Txt姓名.Visible = True
    Me.Lbl年龄.Visible = True
    Me.Txt年龄.Visible = True
    Me.Lbl科室.Visible = True
    Me.Cbo科室.Visible = True
    Me.Lbl日期.Visible = True
    Me.Dtp日期.Visible = True
    Me.Lbl住院号.Visible = True
    Me.Txt住院号.Visible = True
    Me.Lbl性别.Visible = True
    Me.Cbo性别.Visible = True
    Me.Lbl发药人.Visible = True
    Me.Txt发药人.Visible = True
    Me.Lbl记费人.Visible = True
    Me.Txt记费人.Visible = True
    Me.Lbl开单人.Visible = True
    Me.Txt开单人.Visible = True
        
    Select Case byt单据
    
    Case 24, 25                                     '直接输入处方
        Me.Lbl入出类别.Visible = False
        Me.Cbo入出类别.Visible = False
        Me.Lbl库房.Visible = False
        Me.Cbo库房.Visible = False
        Me.Lbl发票号.Visible = False
        Me.Txt发票号.Visible = False
        Me.Lbl供药单位.Visible = False
        Me.Txt供药单位.Visible = False
        Me.Cmd供药单位.Visible = False
        Me.Lbl部门.Visible = False
        Me.Cbo部门.Visible = False
        Me.Lbl摘要.Visible = False
        Me.Txt摘要.Visible = False
        Me.Lbl填制人.Visible = False
        Me.Txt填制人.Visible = False
        Me.Lbl审核人.Visible = False
        Me.Txt审核人.Visible = False
        Me.Lbl填制日期.Visible = False
        Me.Txt填制日期.Visible = False
        Me.Lbl审核日期.Visible = False
        Me.Lbl审核人.Visible = False
        Me.Txt审核日期.Visible = False
    Case 26                                        '直接输入摆药单
        Me.Lbl库房.Caption = "科    室"
        Me.Lbl库房.Visible = False
        Me.Cbo库房.Visible = False
        Me.Lbl入出类别.Visible = False
        Me.Cbo入出类别.Visible = False
        Me.Lbl发票号.Visible = False
        Me.Txt发票号.Visible = False
        Me.Lbl供药单位.Visible = False
        Me.Txt供药单位.Visible = False
        Me.Cmd供药单位.Visible = False
        Me.Lbl部门.Visible = False
        Me.Cbo部门.Visible = False
        Me.Lbl摘要.Visible = False
        Me.Txt摘要.Visible = False
        Me.Lbl填制人.Visible = False
        Me.Txt填制人.Visible = False
        Me.Lbl审核人.Visible = False
        Me.Txt审核人.Visible = False
        Me.Lbl填制日期.Visible = False
        Me.Txt填制日期.Visible = False
        Me.Lbl审核人.Visible = False
        Me.Lbl审核日期.Visible = False
        Me.Txt审核日期.Visible = False
        Me.Lbl病人姓名.Visible = False
        Me.Txt姓名.Visible = False
        Me.Lbl年龄.Visible = False
        Me.Txt年龄.Visible = False
        Me.Lbl科室.Visible = False
        Me.Cbo科室.Visible = False
        Me.Lbl住院号.Visible = False
        Me.Txt住院号.Visible = False
        Me.Lbl性别.Visible = False
        Me.Cbo性别.Visible = False
    
    End Select
End Function

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, Me.Caption
End Sub
