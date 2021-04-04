VERSION 5.00
Begin VB.Form frmApparatusInfo 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   Caption         =   "仪器基本信息"
   ClientHeight    =   4290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8430
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox chk发送 
      Alignment       =   1  'Right Justify
      Caption         =   "发送时指定杯号"
      Height          =   195
      Left            =   4140
      TabIndex        =   50
      ToolTipText     =   "勾上表示在使用技师工作站的[发往仪器]功能时，要指定盘号和杯号；不勾则不指定。"
      Top             =   135
      Width           =   1575
   End
   Begin VB.CheckBox chkLog 
      Caption         =   "对数质控图"
      Height          =   210
      Left            =   7125
      TabIndex        =   49
      ToolTipText     =   "一般PCR仪器才采用对数质控图"
      Top             =   3885
      Width           =   1200
   End
   Begin VB.Frame fra酶标 
      Caption         =   "酶标仪设置"
      Height          =   1110
      Left            =   120
      TabIndex        =   38
      Top             =   2130
      Width           =   8160
      Begin VB.TextBox txt酶标 
         Height          =   300
         Index           =   4
         Left            =   5385
         MaxLength       =   40
         TabIndex        =   47
         Top             =   690
         Width           =   2610
      End
      Begin VB.TextBox txt酶标 
         Height          =   300
         Index           =   3
         Left            =   1335
         MaxLength       =   40
         TabIndex        =   45
         Top             =   690
         Width           =   2610
      End
      Begin VB.TextBox txt酶标 
         Height          =   300
         Index           =   2
         Left            =   7395
         MaxLength       =   40
         TabIndex        =   43
         Top             =   300
         Width           =   600
      End
      Begin VB.TextBox txt酶标 
         Height          =   300
         Index           =   1
         Left            =   4470
         MaxLength       =   40
         TabIndex        =   41
         Top             =   300
         Width           =   1800
      End
      Begin VB.TextBox txt酶标 
         Height          =   300
         Index           =   0
         Left            =   810
         MaxLength       =   40
         TabIndex        =   39
         Top             =   300
         Width           =   2500
      End
      Begin VB.Label lbl酶标 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "空白形式(&X)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   4
         Left            =   4275
         TabIndex        =   48
         Top             =   735
         Width           =   990
      End
      Begin VB.Label lbl酶标 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "进板方式(&F)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   46
         Top             =   735
         Width           =   990
      End
      Begin VB.Label lbl酶标 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "振板时间(&J)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   6345
         TabIndex        =   44
         Top             =   345
         Width           =   990
      End
      Begin VB.Label lbl酶标 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "振板频率(&L)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   3390
         TabIndex        =   42
         Top             =   345
         Width           =   990
      End
      Begin VB.Label lbl酶标 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "波长(&B)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   40
         Top             =   345
         Width           =   630
      End
   End
   Begin VB.ComboBox cbo仪器类别 
      Height          =   300
      Left            =   3870
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   900
      Width           =   1860
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -45
      TabIndex        =   36
      Top             =   3315
      Width           =   8535
   End
   Begin VB.ComboBox cbo校准物来源 
      Height          =   300
      ItemData        =   "frmApparatusInfo.frx":0000
      Left            =   4770
      List            =   "frmApparatusInfo.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   3855
      Width           =   2325
   End
   Begin VB.ComboBox cbo试剂来源 
      Height          =   300
      ItemData        =   "frmApparatusInfo.frx":0004
      Left            =   1380
      List            =   "frmApparatusInfo.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   3855
      Width           =   1950
   End
   Begin VB.TextBox txt仪器QC码 
      Height          =   300
      Left            =   6930
      MaxLength       =   8
      TabIndex        =   31
      Top             =   3420
      Width           =   1290
   End
   Begin VB.TextBox txt质控水平 
      Height          =   300
      Left            =   4785
      MaxLength       =   1
      TabIndex        =   29
      Top             =   3435
      Width           =   405
   End
   Begin VB.ComboBox cbo周期单位 
      Height          =   300
      ItemData        =   "frmApparatusInfo.frx":0008
      Left            =   2100
      List            =   "frmApparatusInfo.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   3435
      Width           =   645
   End
   Begin VB.TextBox txt质控周期 
      Height          =   300
      Left            =   1665
      MaxLength       =   5
      TabIndex        =   26
      Top             =   3435
      Width           =   420
   End
   Begin VB.ComboBox cbo校验位 
      Height          =   300
      ItemData        =   "frmApparatusInfo.frx":001E
      Left            =   6990
      List            =   "frmApparatusInfo.frx":0020
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   1710
      Width           =   1290
   End
   Begin VB.ComboBox cbo停止位 
      Height          =   300
      ItemData        =   "frmApparatusInfo.frx":0022
      Left            =   6990
      List            =   "frmApparatusInfo.frx":0024
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   1290
      Width           =   1290
   End
   Begin VB.ComboBox cbo数据位 
      Height          =   300
      ItemData        =   "frmApparatusInfo.frx":0026
      Left            =   6990
      List            =   "frmApparatusInfo.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   885
      Width           =   1290
   End
   Begin VB.ComboBox cbo波特率 
      Height          =   300
      ItemData        =   "frmApparatusInfo.frx":002A
      Left            =   6990
      List            =   "frmApparatusInfo.frx":002C
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   480
      Width           =   1290
   End
   Begin VB.ComboBox cbo通讯口 
      Height          =   300
      ItemData        =   "frmApparatusInfo.frx":002E
      Left            =   6990
      List            =   "frmApparatusInfo.frx":0030
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   75
      Width           =   1290
   End
   Begin VB.TextBox txt通讯程序名 
      Height          =   300
      Left            =   1380
      MaxLength       =   40
      TabIndex        =   14
      Top             =   1755
      Width           =   4335
   End
   Begin VB.TextBox txt连接计算机 
      Height          =   300
      Left            =   1380
      MaxLength       =   40
      TabIndex        =   10
      Top             =   1335
      Width           =   1695
   End
   Begin VB.ComboBox cbo使用小组 
      Height          =   300
      ItemData        =   "frmApparatusInfo.frx":0032
      Left            =   3870
      List            =   "frmApparatusInfo.frx":0034
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1335
      Width           =   1860
   End
   Begin VB.ComboBox cbo仪器类型 
      Height          =   300
      ItemData        =   "frmApparatusInfo.frx":0036
      Left            =   855
      List            =   "frmApparatusInfo.frx":0038
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   930
      Width           =   2235
   End
   Begin VB.TextBox txt编码 
      Height          =   300
      Left            =   855
      MaxLength       =   10
      TabIndex        =   1
      Top             =   120
      Width           =   1320
   End
   Begin VB.TextBox txt名称 
      Height          =   300
      Left            =   855
      MaxLength       =   20
      TabIndex        =   3
      Top             =   525
      Width           =   2610
   End
   Begin VB.TextBox txt简码 
      Height          =   300
      Left            =   4350
      MaxLength       =   10
      TabIndex        =   5
      Top             =   525
      Width           =   1365
   End
   Begin VB.Label lbl仪器类别 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "类别(&G)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3225
      TabIndex        =   37
      Top             =   990
      Width           =   630
   End
   Begin VB.Label lbl校准物来源 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "默认校准物来源"
      Height          =   180
      Left            =   3435
      TabIndex        =   34
      Top             =   3915
      Width           =   1260
   End
   Begin VB.Label lbl试剂来源 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "默认试剂来源"
      Height          =   180
      Left            =   180
      TabIndex        =   32
      Top             =   3915
      Width           =   1080
   End
   Begin VB.Label lbl仪器QC码 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "仪器QC码"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6075
      TabIndex        =   30
      Top             =   3480
      Width           =   720
   End
   Begin VB.Label lbl质控水平 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "每批检测     个水平"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4020
      TabIndex        =   28
      Top             =   3495
      Width           =   1710
   End
   Begin VB.Label lbl质控周期 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "质控要求: 至少每             进行一次质控,"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   25
      Top             =   3495
      Width           =   3780
   End
   Begin VB.Label lbl校验位 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "校验位(&5)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6135
      TabIndex        =   23
      Top             =   1770
      Width           =   810
   End
   Begin VB.Label lbl停止位 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "停止位(&4)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6135
      TabIndex        =   21
      Top             =   1350
      Width           =   810
   End
   Begin VB.Label lbl数据位 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "数据位(&3)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6135
      TabIndex        =   19
      Top             =   945
      Width           =   810
   End
   Begin VB.Label lbl波特率 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "波特率(&2)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6135
      TabIndex        =   17
      Top             =   540
      Width           =   810
   End
   Begin VB.Label lbl通讯口 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "通讯口(&1)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6135
      TabIndex        =   15
      Top             =   135
      Width           =   810
   End
   Begin VB.Label lbl通讯程序名 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "通讯程序名(&P)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   13
      Top             =   1800
      Width           =   1170
   End
   Begin VB.Label lbl连接计算机 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "连接计算机(&T)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   9
      Top             =   1395
      Width           =   1170
   End
   Begin VB.Label lbl使用小组 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "使用(&U)"
      Height          =   180
      Left            =   3225
      TabIndex        =   11
      Top             =   1395
      Width           =   630
   End
   Begin VB.Label lbl仪器类型 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "类型(&K)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   6
      Top             =   990
      Width           =   630
   End
   Begin VB.Label lbl编码 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "编码(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   630
   End
   Begin VB.Label lbl项目名称 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "名称(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   2
      Top             =   585
      Width           =   630
   End
   Begin VB.Label lbl简码 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "简码(&S)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3660
      TabIndex        =   4
      Top             =   585
      Width           =   630
   End
End
Attribute VB_Name = "frmApparatusInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLngAptId As Long          '当前显示的仪器id

Dim objNode As Node
Dim strTemp As String, aryTemp() As String
Dim lngCount As Long

'--------------------------------------------
'以下为窗体公共方法
'--------------------------------------------
Public Function zlRefresh(lngAptId As Long) As Boolean
    '功能：根据项目id刷新当前显示内容
    Dim rsTemp As New ADODB.Recordset, intIndex As Integer
    mLngAptId = lngAptId
    
    '清除此前项目的显示
    Me.txt编码.Text = "": Me.txt名称.Text = "": Me.txt简码.Text = ""
    Me.cbo仪器类型.ListIndex = -1: Me.cbo仪器类别.ListIndex = -1 'Me.chk微生物.Value = 0
    Me.txt连接计算机.Text = "": Me.cbo使用小组.ListIndex = -1: Me.txt通讯程序名.Text = ""
    Me.cbo通讯口.ListIndex = -1: Me.cbo波特率.ListIndex = -1
    Me.cbo数据位.ListIndex = -1: Me.cbo停止位.ListIndex = -1: Me.cbo校验位.ListIndex = -1
    Me.txt质控周期.Text = "": Me.cbo周期单位.ListIndex = -1:  Me.txt质控水平.Text = 0
    Me.txt仪器QC码.Text = "": Me.cbo试剂来源.ListIndex = -1: Me.cbo校准物来源.ListIndex = -1
    For intIndex = 0 To 4
        Me.txt酶标(intIndex).Text = ""
    Next
    If lngAptId = 0 Then zlRefresh = True: Exit Function
    
    '获取指定项目的信息
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select A.编码, A.名称, A.简码, A.仪器类型, A.微生物, A.连接计算机, A.通讯程序名, A.通讯端口, A.波特率, A.数据, A.停止位," & vbNewLine & _
            "       A.校验位, A.使用小组id, D.名称 As 使用小组, A.质控周期, A.周期单位, A.质控水平数, A.Qc码, A.试剂来源," & vbNewLine & _
            "       A.校准物来源,A.波长,A.振板频率,A.振板时间,A.进板方式,A.空白形式,A.对数质控图,A.发送时指定杯号 " & vbNewLine & _
            "From 检验仪器 A, 部门表 D" & vbNewLine & _
            "Where A.使用小组id = D.ID(+) And A.ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngAptId)
    With rsTemp
        If .RecordCount > 0 Then
            Me.txt编码.Text = "" & !编码: Me.txt名称.Text = "" & !名称: Me.txt简码.Text = "" & !简码
            For lngCount = 0 To Me.cbo仪器类型.ListCount - 1
                If Mid(Me.cbo仪器类型.List(lngCount), InStr(1, Me.cbo仪器类型.List(lngCount), "-") + 1) = "" & !仪器类型 Then
                    Me.cbo仪器类型.ListIndex = lngCount: Exit For
                End If
            Next
            'Me.chk微生物.Value = IIf(Val("" & !微生物) = 1, 1, 0)
            For intIndex = 0 To cbo仪器类别.ListCount - 1
                If Val(cbo仪器类别.List(intIndex)) = Val("" & !微生物) Then
                    cbo仪器类别.ListIndex = intIndex
                    Exit For
                End If
            Next
            
            Me.txt酶标(0).Text = "" & !波长
            Me.txt酶标(1).Text = "" & !振板频率
            Me.txt酶标(2).Text = "" & !振板时间
            Me.txt酶标(3).Text = "" & !进板方式
            Me.txt酶标(4).Text = "" & !空白形式
            Me.chkLog.Value = Val("" & !对数质控图)
            Me.chk发送.Value = Val("" & !发送时指定杯号)
            
            Me.txt连接计算机.Text = "" & !连接计算机
            For lngCount = 0 To Me.cbo使用小组.ListCount - 1
                If Me.cbo使用小组.ItemData(lngCount) = Val("" & !使用小组id) Then
                    Me.cbo使用小组.ListIndex = lngCount: Exit For
                End If
            Next
            Me.txt通讯程序名.Text = "" & !通讯程序名
                        
            For lngCount = 0 To Me.cbo通讯口.ListCount - 1
                If Me.cbo通讯口.List(lngCount) = "" & !通讯端口 Then Me.cbo通讯口.ListIndex = lngCount: Exit For
            Next
            For lngCount = 0 To Me.cbo波特率.ListCount - 1
                If Me.cbo波特率.List(lngCount) = "" & !波特率 Then Me.cbo波特率.ListIndex = lngCount: Exit For
            Next
            For lngCount = 0 To Me.cbo数据位.ListCount - 1
                If Me.cbo数据位.List(lngCount) = "" & !数据 Then Me.cbo数据位.ListIndex = lngCount: Exit For
            Next
            For lngCount = 0 To Me.cbo停止位.ListCount - 1
                If Me.cbo停止位.List(lngCount) = "" & !停止位 Then Me.cbo停止位.ListIndex = lngCount: Exit For
            Next
            For lngCount = 0 To Me.cbo校验位.ListCount - 1
                If InStr(1, Me.cbo校验位.List(lngCount), "-") = 0 Then
                    If Me.cbo校验位.List(lngCount) = "" & !校验位 Then Me.cbo校验位.ListIndex = lngCount: Exit For
                Else
                    If Left(Me.cbo校验位.List(lngCount), 1) = "" & !校验位 Then Me.cbo校验位.ListIndex = lngCount: Exit For
                End If
            Next
            
            Me.txt质控周期.Text = Val("" & !质控周期)
            If "" & !周期单位 <> "月" Then
                Me.cbo周期单位.ListIndex = 0
            Else
                Me.cbo周期单位.ListIndex = 1
            End If
            Me.txt质控水平.Text = Val("" & !质控水平数)
            Me.txt仪器QC码.Text = "" & !Qc码
            For lngCount = 0 To Me.cbo试剂来源.ListCount - 1
                If Me.cbo试剂来源.List(lngCount) = "" & !试剂来源 Then Me.cbo试剂来源.ListIndex = lngCount: Exit For
            Next
            For lngCount = 0 To Me.cbo校准物来源.ListCount - 1
                If Me.cbo校准物来源.List(lngCount) = "" & !校准物来源 Then Me.cbo校准物来源.ListIndex = lngCount: Exit For
            Next
        End If
    End With
        
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart(blnAdd As Boolean, lngAptId As Long) As Boolean
    '功能：开始项目编辑
    '参数： blnAdd-是否增加，否则为修改
    '       lngAptId-增加的参照项目，或者指定编辑的项目
    Dim rsTemp As New ADODB.Recordset, i As Integer
    If Me.cbo仪器类型.ListCount = 0 Then
        MsgBox "请先在字典中初始化“检验类型”！", vbInformation, gstrSysName
        zlEditStart = False: Exit Function
    End If
    
    If blnAdd Then
        Err = 0: On Error GoTo ErrHand
        gstrSql = "Select Nvl(Max(编码),'000') As 编码 From 检验仪器"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "zlEditStart")
        With rsTemp
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
            
'            Call SQLTest
            'Me.txt编码.Text = Right(String(10, "0") & Val(!编码) + 1, Len(!编码))
            Me.txt编码.Text = zlCommFun.IncStr(IIf("" & !编码 = "", "000", "" & !编码))
        End With
        
        '清除并设置默认值
        Me.txt名称.Text = "": Me.txt简码.Text = ""
        Me.cbo仪器类型.ListIndex = 0: Me.cbo仪器类别.ListIndex = 0 'Me.chk微生物.Value = 0
        Me.txt连接计算机.Text = "": Me.cbo使用小组.ListIndex = -1: Me.txt通讯程序名.Text = ""
        Me.cbo通讯口.ListIndex = 0: Me.cbo波特率.ListIndex = 5
        Me.cbo数据位.ListIndex = 4: Me.cbo停止位.ListIndex = 0: Me.cbo校验位.ListIndex = 3
        Me.txt质控周期.Text = 1: Me.cbo周期单位.ListIndex = 0: Me.txt质控水平.Text = 1
        Me.txt仪器QC码.Text = "": Me.cbo试剂来源.ListIndex = -1: Me.cbo校准物来源.ListIndex = -1
        For i = 0 To 4
            Me.txt酶标(i).Text = "": Me.txt酶标(i).Tag = ""
        Next
    Else
        If Val(Me.txt质控周期.Text) = 0 Then Me.txt质控周期.Text = 1
        If Me.cbo周期单位.ListIndex = -1 Then Me.cbo周期单位.ListIndex = 0
        If Val(Me.txt质控水平.Text) = 0 Then Me.txt质控水平.Text = 1
        For i = 0 To 4
            Me.txt酶标(i).Tag = Me.txt酶标(i).Text
        Next
    End If

    mLngAptId = lngAptId
    Me.Enabled = True: Me.Tag = IIf(blnAdd, "增加", "修改")
    Me.BackColor = RGB(250, 250, 250): Me.fra酶标.BackColor = Me.BackColor
    Me.txt编码.SetFocus
    zlEditStart = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditStart = False: Exit Function
End Function

Public Sub zlEditCancel()
    '功能：放弃正在进行的编辑
    Me.Enabled = False: Me.Tag = ""
    Me.BackColor = Me.fraLine.BackColor: Me.fra酶标.BackColor = Me.BackColor
    Call Me.zlRefresh(mLngAptId)
End Sub

Public Function zlEditSave() As Long
    '功能：保存正在进行的编辑,并返回正在编辑项目id,保存失败返回0
    Dim lngNewId As Long
    
    '一般特性检查
    If Trim(Me.txt编码.Text) = "" Then
        MsgBox "请输入编码！", vbInformation, gstrSysName
        Me.txt编码.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt编码.Text), vbFromUnicode)) > 3 Then
        MsgBox "编码的长度超长（最多3个字符）！", vbInformation, gstrSysName
        Me.txt编码.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Trim(Me.txt名称.Text) = "" Then
        MsgBox "请输入名称！", vbInformation, gstrSysName
        Me.txt名称.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Me.cbo仪器类型.ListIndex = -1 Then
        MsgBox "请选择仪器类型！", vbInformation, gstrSysName
        Me.cbo仪器类型.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt名称.Text), vbFromUnicode)) > Me.txt名称.MaxLength Then
        MsgBox "名称超长（最多" & Me.txt名称.MaxLength & "个字符或等长汉字）！", vbInformation, gstrSysName
        Me.txt名称.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt连接计算机.Text), vbFromUnicode)) > Me.txt连接计算机.MaxLength Then
        MsgBox "连接计算机超长（最多" & Me.txt连接计算机.MaxLength & "个字符或等长汉字）！", vbInformation, gstrSysName
        Me.txt连接计算机.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt通讯程序名.Text), vbFromUnicode)) > Me.txt通讯程序名.MaxLength Then
        MsgBox "通讯程序名超长（最多" & Me.txt通讯程序名.MaxLength & "个字符或等长汉字）！", vbInformation, gstrSysName
        Me.txt通讯程序名.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Val(Me.txt质控周期.Text) <= 0 Or Val(Me.txt质控周期.Text) > 365 Then
        MsgBox "请设置合理质控周期！", vbInformation, gstrSysName
        Me.txt质控周期.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Me.cbo周期单位.ListIndex = -1 Then
        MsgBox "请设置质控周期单位！", vbInformation, gstrSysName
        Me.cbo周期单位.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Val(Me.txt质控水平.Text) <= 0 Or Val(Me.txt质控水平.Text) > 9 Then
        MsgBox "请设置合理质控水平数（1～9）！", vbInformation, gstrSysName
        Me.txt质控水平.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt仪器QC码.Text), vbFromUnicode)) > Me.txt仪器QC码.MaxLength Then
        MsgBox "仪器QC码超长（最多" & Me.txt仪器QC码.MaxLength & "个字符）！", vbInformation, gstrSysName
        Me.txt仪器QC码.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    '数据保存语句组织
    If Me.Tag = "增加" Then
        lngNewId = zlDatabase.GetNextId("检验仪器")
    Else
        lngNewId = mLngAptId
    End If

    gstrSql = lngNewId & ",'" & Replace(Trim(Me.txt编码.Text), "'", "") & "','" & Replace(Trim(Me.txt名称.Text), "'", "") & "','" & Replace(Trim(Me.txt简码.Text), "'", "") & "'"
    gstrSql = gstrSql & ",'" & Replace(Trim(Me.txt连接计算机.Text), "'", "") & "','" & Replace(Trim(Me.txt通讯程序名.Text), "'", "") & "','" & Me.cbo通讯口.Text & "'"
    gstrSql = gstrSql & "," & Val(Me.cbo波特率.Text) & "," & Val(Me.cbo数据位.Text) & "," & Val(Me.cbo停止位.Text)
    If InStr(1, Me.cbo校验位.Text, "-") = 0 Then
        gstrSql = gstrSql & ",'" & Me.cbo校验位.Text & "'"
    Else
        gstrSql = gstrSql & ",'" & Left(Me.cbo校验位.Text, 1) & "'"
    End If
    gstrSql = gstrSql & ",'" & Mid(Me.cbo仪器类型.Text, InStr(1, Me.cbo仪器类型.Text, "-") + 1) & "'"
    gstrSql = gstrSql & "," & Val(Me.cbo仪器类别.Text)
    If Me.cbo使用小组.ListIndex = -1 Then
        gstrSql = gstrSql & ",Null"
    Else
        gstrSql = gstrSql & "," & Me.cbo使用小组.ItemData(Me.cbo使用小组.ListIndex)
    End If
    gstrSql = gstrSql & "," & Val(Me.txt质控周期.Text) & ",'" & Trim(Me.cbo周期单位.Text) & "'," & Val(Me.txt质控水平.Text)
    gstrSql = gstrSql & ",'" & Replace(Trim(Me.txt仪器QC码.Text), "'", "") & "','" & Trim(Me.cbo试剂来源.Text) & "','" & Trim(Me.cbo校准物来源.Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txt酶标(0).Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txt酶标(1).Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txt酶标(2).Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txt酶标(3).Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txt酶标(4).Text) & "'"
    gstrSql = gstrSql & "," & Me.chkLog.Value
    gstrSql = gstrSql & "," & Me.chk发送.Value
    If Me.Tag = "增加" Then
        gstrSql = "Zl_检验仪器_Insert(" & gstrSql & ")"
    Else
        gstrSql = "Zl_检验仪器_Update(" & gstrSql & ")"
    End If
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    If Me.Tag = "增加" Then mLngAptId = lngNewId
    Me.Enabled = False: Me.Tag = ""
    Me.BackColor = Me.fraLine.BackColor: Me.fra酶标.BackColor = Me.BackColor
    zlEditSave = mLngAptId: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

'--------------------------------------------
'以下为窗体控件响应事件
'--------------------------------------------

Private Sub cbo波特率_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo波特率_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo使用小组_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo使用小组_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo试剂来源_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo试剂来源_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo数据位_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo数据位_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo停止位_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo停止位_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo通讯口_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo通讯口_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo校验位_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo校验位_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo校准物来源_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo校准物来源_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo仪器类别_Click()
    Dim i As Integer
    If cbo仪器类别.ListIndex = 2 Then
        For i = 0 To 4
            Me.txt酶标(i).Enabled = True: Me.txt酶标(i).Text = Me.txt酶标(i).Tag
        Next
    Else
        For i = 0 To 4
            Me.txt酶标(i).Enabled = False:  If Me.txt酶标(i).Text <> "" Then Me.txt酶标(i).Tag = Me.txt酶标(i).Text: Me.txt酶标(i).Text = ""
        Next
    End If
End Sub

Private Sub cbo仪器类别_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo仪器类型_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo仪器类型_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo周期单位_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo周期单位_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

'Private Sub chk微生物_GotFocus()
'    Call zlCommFun.OpenIme(False)
'End Sub
'
'Private Sub chk微生物_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
'End Sub
    
Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    Dim aryTemp() As String
    Err = 0: On Error GoTo ErrHand
    '字段长度限制
    gstrSql = "Select A.编码, A.名称, A.简码, A.仪器类型, A.微生物, A.连接计算机, A.通讯程序名, A.通讯端口, A.波特率, A.数据, A.停止位," & vbNewLine & _
            "       A.校验位, A.使用小组id, D.名称 As 使用小组, A.QC码" & vbNewLine & _
            "From 检验仪器 A, 部门表 D" & vbNewLine & _
            "Where A.使用小组id = D.ID(+) And A.ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mLngAptId)
    With rsTemp
'        Me.txt编码.MaxLength = .Fields("编码").DefinedSize
        Me.txt名称.MaxLength = .Fields("名称").DefinedSize
        Me.txt简码.MaxLength = .Fields("简码").DefinedSize
        Me.txt连接计算机.MaxLength = .Fields("连接计算机").DefinedSize
        Me.txt通讯程序名.MaxLength = .Fields("通讯程序名").DefinedSize
        Me.txt仪器QC码.MaxLength = .Fields("QC码").DefinedSize
    End With
    
    '诊疗检验类型
    gstrSql = "Select 编码,名称 From 诊疗检验类型"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mLngAptId)
    With rsTemp
        Me.cbo仪器类型.Clear
        Do While Not .EOF
            Me.cbo仪器类型.AddItem Trim(!编码) & "-" & Trim(!名称)
            .MoveNext
        Loop
        If Me.cbo仪器类型.ListCount > 0 Then Me.cbo仪器类型.ListIndex = 0
    End With
    '检验仪器类别
    aryTemp = Split("0-普通仪器;1-微生物仪;2-酶标仪", ";")
    For lngCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo仪器类别.AddItem aryTemp(lngCount)
    Next
    Me.cbo仪器类别.ListIndex = 0
    
    '检验执行部门
    gstrSql = "Select Id, 编码, 名称" & vbNewLine & _
            "From 部门表 d, 部门性质说明 p" & vbNewLine & _
            "Where d.Id = p.部门id And p.工作性质 = '检验' And Instr(',1,2,3,', ','||p.服务对象||',') > 0 And" & vbNewLine & _
            "           (To_Char(d.撤档时间, 'YYYY-MM-DD') = '3000-01-01' Or d.撤档时间 Is Null)"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Do While Not .EOF
            Me.cbo使用小组.AddItem !编码 & "-" & !名称
            Me.cbo使用小组.ItemData(Me.cbo使用小组.NewIndex) = !ID
            .MoveNext
        Loop
    End With
    
    '试剂来源与校准物来源
    gstrSql = "Select 名称 From 质控试剂来源 Order By 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Do While Not .EOF
            Me.cbo试剂来源.AddItem "" & !名称
            Me.cbo校准物来源.AddItem "" & !名称
            .MoveNext
        Loop
    End With
    
    '其他固定内容装入
    For lngCount = 1 To 50: Me.cbo通讯口.AddItem "COM" & lngCount: Next
    Me.cbo通讯口.ListIndex = 0

    aryTemp = Split("110|300|600|1200|2400|4800|9600|14400|19200|28800|38400|56000|128000|256000", "|")
    For lngCount = LBound(aryTemp) To UBound(aryTemp): Me.cbo波特率.AddItem aryTemp(lngCount): Next
    Me.cbo波特率.ListIndex = 0

    aryTemp = Split("4|5|6|7|8", "|")
    For lngCount = LBound(aryTemp) To UBound(aryTemp): Me.cbo数据位.AddItem aryTemp(lngCount): Next
    Me.cbo数据位.ListIndex = 0

    aryTemp = Split("1|1.5|2", "|")
    For lngCount = LBound(aryTemp) To UBound(aryTemp): Me.cbo停止位.AddItem aryTemp(lngCount): Next
    Me.cbo停止位.ListIndex = 0

    aryTemp = Split("E-偶数|M-标记|N-缺省|None|O-奇数|S-空格", "|")
    For lngCount = LBound(aryTemp) To UBound(aryTemp): Me.cbo校验位.AddItem aryTemp(lngCount): Next
    Me.cbo校验位.ListIndex = 0
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt编码_GotFocus()
    Me.txt编码.SelStart = 0: Me.txt编码.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt编码_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt简码_GotFocus()
    Me.txt简码.SelStart = 0: Me.txt简码.SelLength = 1000
End Sub

Private Sub txt简码_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt连接计算机_GotFocus()
    Me.txt连接计算机.SelStart = 0: Me.txt连接计算机.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt连接计算机_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt名称_GotFocus()
    Me.txt名称.SelStart = 0: Me.txt名称.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt名称_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Me.txt名称.Text = MoveSpecialChar(Me.txt名称.Text)
        Me.txt简码.Text = zlStr.GetCodeByORCL(Me.txt名称.Text, False, Me.txt简码.MaxLength)
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End If
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt名称_LostFocus()
    Me.txt简码.Text = zlStr.GetCodeByORCL(Me.txt名称.Text, False, Me.txt简码.MaxLength)
End Sub

Private Sub txt通讯程序名_GotFocus()
    Me.txt通讯程序名.SelStart = 0: Me.txt通讯程序名.SelLength = 1000
End Sub

Private Sub txt通讯程序名_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt仪器QC码_GotFocus()
    Me.txt仪器QC码.SelStart = 0: Me.txt仪器QC码.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt仪器QC码_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt质控水平_GotFocus()
    Me.txt质控水平.SelStart = 0: Me.txt质控水平.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt质控水平_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii > Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt质控周期_GotFocus()
    Me.txt质控周期.SelStart = 0: Me.txt质控周期.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt质控周期_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub
