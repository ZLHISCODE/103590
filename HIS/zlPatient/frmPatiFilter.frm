VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.10#0"; "zlIDKind.ocx"
Begin VB.Form frmPatiFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤设置"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdDef 
      Caption         =   "缺省(&D)"
      Height          =   350
      Left            =   6480
      TabIndex        =   5
      Top             =   2475
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6480
      TabIndex        =   4
      Top             =   735
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6480
      TabIndex        =   3
      Top             =   300
      Width           =   1100
   End
   Begin VB.Frame fraBdr 
      Height          =   3420
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   6225
      Begin VB.ComboBox cboPayPlan 
         Height          =   300
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1800
         Width           =   2085
      End
      Begin VB.CheckBox chk登记 
         Caption         =   "登记时间"
         Height          =   195
         Left            =   180
         TabIndex        =   26
         Top             =   338
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.CheckBox chk出生 
         Caption         =   "出生日期"
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   713
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.CheckBox chk入院 
         Caption         =   "入院时间"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   2610
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.CommandButton cmd区域 
         Caption         =   "…"
         Height          =   255
         Left            =   5730
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "热键：F3"
         Top             =   1450
         Width           =   285
      End
      Begin VB.CheckBox chk出院 
         Caption         =   "出院时间"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   2985
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.ComboBox cbo性别 
         Height          =   300
         Left            =   3945
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1050
         Width           =   2085
      End
      Begin VB.TextBox txt区域 
         Height          =   300
         Left            =   3945
         MaxLength       =   30
         TabIndex        =   14
         Top             =   1425
         Width           =   2085
      End
      Begin MSComCtl2.DTPicker dtp出院E 
         Height          =   300
         Left            =   3945
         TabIndex        =   15
         Top             =   2925
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   149422083
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtp出院B 
         Height          =   300
         Left            =   1230
         TabIndex        =   16
         Top             =   2925
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   149422083
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtp入院E 
         Height          =   300
         Left            =   3945
         TabIndex        =   20
         Top             =   2550
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   149422083
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtp入院B 
         Height          =   300
         Left            =   1230
         TabIndex        =   21
         Top             =   2550
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   149422083
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtp出生E 
         Height          =   300
         Left            =   3945
         TabIndex        =   24
         Top             =   660
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   149422083
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtp登记E 
         Height          =   300
         Left            =   3945
         TabIndex        =   27
         Top             =   285
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   149422083
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtp出生B 
         Height          =   300
         Left            =   1230
         TabIndex        =   31
         Top             =   660
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   149422083
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtp登记B 
         Height          =   300
         Left            =   1230
         TabIndex        =   32
         Top             =   285
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   149422083
         CurrentDate     =   40544
      End
      Begin VB.TextBox txt编号 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1230
         TabIndex        =   30
         Top             =   1425
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.ComboBox cbo费别 
         Height          =   300
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1425
         Width           =   2085
      End
      Begin VB.TextBox txt住院号 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   1230
         MaxLength       =   18
         TabIndex        =   29
         Top             =   1050
         Width           =   2085
      End
      Begin zlIDKind.PatiIdentify PatiIdentify 
         Height          =   270
         Left            =   1230
         TabIndex        =   19
         Top             =   2175
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmPatiFilter.frx":0000
         BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoSize        =   -1  'True
         IDKindAppearance=   0
         ShowPropertySet =   -1  'True
         DefaultCardType =   "0"
         IDKindWidth     =   555
         FindPatiShowName=   0   'False
         HiddenMoseRightKey=   0   'False
         BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lbl医疗付款方式 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医疗付款方式"
         Height          =   180
         Left            =   120
         TabIndex        =   34
         Top             =   1860
         Width           =   1080
      End
      Begin VB.Label lbl登记 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   3540
         TabIndex        =   28
         Top             =   345
         Width           =   180
      End
      Begin VB.Label lbl出生 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   3540
         TabIndex        =   25
         Top             =   720
         Width           =   180
      End
      Begin VB.Label lbl入院 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   3540
         TabIndex        =   22
         Top             =   2610
         Width           =   180
      End
      Begin VB.Label lbl出院 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   3540
         TabIndex        =   17
         Top             =   2985
         Width           =   180
      End
      Begin VB.Label lblIDKind 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人"
         Height          =   180
         Left            =   840
         TabIndex        =   12
         Top             =   2220
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         Height          =   180
         Left            =   630
         TabIndex        =   11
         Top             =   1110
         Width           =   540
      End
      Begin VB.Label lbl区域 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "区域"
         Height          =   180
         Left            =   3450
         TabIndex        =   9
         Top             =   1485
         Width           =   360
      End
      Begin VB.Label lbl费别 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费别"
         Height          =   180
         Left            =   810
         TabIndex        =   8
         Top             =   1485
         Width           =   360
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   180
         Left            =   3450
         TabIndex        =   7
         Top             =   1110
         Width           =   360
      End
      Begin VB.Label lbl编号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "特殊编号"
         Height          =   180
         Left            =   480
         TabIndex        =   10
         Top             =   1485
         Visible         =   0   'False
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmPatiFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Public mbytType As Byte '入:病人清单类型0-所有,1-在院,2-出院,3-门诊,4-留观
Public mstrFilter As String '出:条件
Public mstrFilterInfo As String '病人信息 专用过滤条件
Public mbytInFun As Byte '0-普通调用,1-特殊病人过滤调用
Public mlngPatiId As Long   '病人ID

Private mstrFindType As String        '存储当前查找类型名称

Private Sub cmd区域_Click()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = GetArea(Me, txt区域, True)
    If Not rsTmp Is Nothing Then
        txt区域.Text = rsTmp!名称
        txt区域.SelStart = Len(txt区域.Text)
        txt区域.SetFocus
    Else
        zlControl.TxtSelAll txt区域
        txt区域.SetFocus
    End If
End Sub

Private Sub chk登记_Click()
    If chk登记.Tag <> "" Then chk登记.Value = 0: Exit Sub
    dtp登记B.Enabled = (chk登记.Value = 1)
    dtp登记E.Enabled = dtp登记B.Enabled
    If dtp登记B.Enabled Then dtp登记B.SetFocus
End Sub

Private Sub chk出生_Click()
    If chk出生.Tag <> "" Then chk出生.Value = 0: Exit Sub
    dtp出生B.Enabled = (chk出生.Value = 1)
    dtp出生E.Enabled = dtp出生B.Enabled
    If dtp出生B.Enabled Then dtp出生B.SetFocus
End Sub

Private Sub chk出院_Click()
    If chk出院.Tag <> "" Then chk出院.Value = 0: Exit Sub
    dtp出院B.Enabled = (chk出院.Value = 1)
    dtp出院E.Enabled = dtp出院B.Enabled
    If dtp出院B.Enabled Then dtp出院B.SetFocus
End Sub

Private Sub chk入院_Click()
    If chk入院.Tag <> "" Then chk入院.Value = 0: Exit Sub
    dtp入院B.Enabled = (chk入院.Value = 1)
    dtp入院E.Enabled = dtp入院B.Enabled
    If dtp入院B.Enabled Then dtp入院B.SetFocus
End Sub

Private Sub cmdCancel_Click()
    gblnOK = False
    Hide
End Sub

Private Sub cmdDef_Click()
    Form_Load
End Sub

Private Sub cmdOK_Click()
    txt住院号.Text = Trim(txt住院号.Text)
     
    If txt住院号.Text = "" And Trim(PatiIdentify.Text) = "" Then
        If chk登记.Value = 0 And chk入院.Value = 0 And chk出院.Value = 0 And mbytType <> 1 Then
            MsgBox "请至少选择一个登记时间范围.", vbInformation, gstrSysName
            chk登记.Value = 1
            Exit Sub
        End If
        
        If mbytType = 0 Then
            If chk登记.Value = 0 Then
                MsgBox "请至少选择一个登记时间范围.", vbInformation, gstrSysName
                chk登记.Value = 1
                Exit Sub
            End If
        End If
    End If
        
    Call MakeFilter
    gblnOK = True
    Hide
End Sub

Private Sub dtp出生E_Change()
    dtp出生B.MaxDate = dtp出生E.Value
End Sub

Private Sub dtp出院E_Change()
    dtp出院B.MaxDate = dtp出院E.Value
End Sub

Private Sub dtp登记E_Change()
    dtp登记B.MaxDate = dtp登记E.Value
End Sub

Private Sub dtp入院E_Change()
    dtp入院B.MaxDate = dtp入院E.Value
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Select Case mbytType
        Case 0
            dtp登记B.SetFocus
        Case 1
            chk入院.SetFocus
        Case 2
            dtp出院B.SetFocus
        Case 3, 4
            dtp登记B.SetFocus
    End Select
    mlngPatiId = 0
    PatiIdentify.Text = ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim Curdate As Date, datTmp As Date, i As Integer
    Dim strKindType As String
    
    If Not gobjSquare Is Nothing Then
        strKindType = "姓|姓名|0|0|0|0|0|0|0|0|0;就|就诊卡|0|0|0|0|0|0|0|0|0;门|门诊号|0|0|0|0|0|0|0|0|0;医|医保号|0|0|0|0|0|0|0|0|0;身|二代身份证|0|0|0|0|0|0|0|0|0;IC|IC卡|0|0|0|0|0|0|0|0|0;手|手机号|0|0|0|0|0|0|0|0|0"
        Call PatiIdentify.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, strKindType, "ZL9Patient")
        PatiIdentify.objIDKind.AllowAutoICCard = True
        PatiIdentify.objIDKind.AllowAutoIDCard = True
    End If
    PatiIdentify.Text = ""
    lbl费别.Visible = mbytInFun = 0
    cbo费别.Visible = mbytInFun = 0
    lbl编号.Visible = mbytInFun = 1
    txt编号.Visible = mbytInFun = 1
    
    If mbytInFun = 0 Then
        '费别
        If glngSys Like "8??" Then
            lbl费别.Caption = "会员等级"
        Else
            If mbytType = 0 Or mbytType = 3 Or mbytType = 4 Then
                lbl费别.Caption = "门诊费别"
            Else
                lbl费别.Caption = "住院费别"
            End If
        End If
        
        Set rsTmp = Nothing
        Set rsTmp = GetDictData("费别")
        cbo费别.Clear
        cbo费别.AddItem "所有费别"
        cbo费别.ListIndex = 0
        If Not rsTmp Is Nothing Then
            For i = 1 To rsTmp.RecordCount
                cbo费别.AddItem rsTmp!编码 & "-" & rsTmp!名称
                rsTmp.MoveNext
            Next
        End If
    ElseIf mbytInFun = 1 Then
        chk登记.Caption = "加入时间"
    End If
    
    '性别
    Set rsTmp = Nothing
    Set rsTmp = GetDictData("性别")
    cbo性别.Clear
    cbo性别.AddItem "所有性别"
    cbo性别.ListIndex = 0
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cbo性别.AddItem rsTmp!编码 & "-" & rsTmp!名称
            rsTmp.MoveNext
        Next
    End If
    
    '医疗付款方式
    Set rsTmp = Nothing
    Set rsTmp = GetDictData("医疗付款方式")
    cboPayPlan.Clear
    cboPayPlan.AddItem "所有方式"
    cboPayPlan.ListIndex = 0
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cboPayPlan.AddItem rsTmp!编码 & "-" & rsTmp!名称
            rsTmp.MoveNext
        Next
    End If


    '设置初始条件
    On Error Resume Next    '避免注册表存储无效时间时出错
    Curdate = zlDatabase.Currentdate
    dtp登记B.MaxDate = Format(DateAdd("d", 1, Curdate), dtp登记E.CustomFormat)
    dtp出生B.MaxDate = Format(Curdate, dtp出生E.CustomFormat)
    dtp入院B.MaxDate = Format(DateAdd("d", 1, Curdate), dtp入院E.CustomFormat)
    dtp出院B.MaxDate = Format(DateAdd("d", 1, Curdate), dtp出院E.CustomFormat)
        
    datTmp = Format(Curdate, "yyyy-MM-dd 00:00:00")
    dtp登记B.Value = datTmp
    datTmp = Format(Curdate, "yyyy-MM-dd 23:59:59")
    dtp登记E.Value = datTmp
    
    datTmp = CDate(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "出生开始时间", Format(DateAdd("yyyy", -100, Curdate), "yyyy-MM-dd")))
    dtp出生B.Value = datTmp
    datTmp = CDate(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "出生结束时间", Format(dtp出生B.MaxDate, dtp出生E.CustomFormat)))
    dtp出生E.Value = datTmp
    
    datTmp = CDate(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "入院开始时间", Format(Curdate, "YYYY-MM-DD")))
    dtp入院B.Value = datTmp
    datTmp = CDate(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "入院结束时间", Format(dtp入院B.MaxDate, dtp入院E.CustomFormat)))
    dtp入院E.Value = datTmp
    
    datTmp = CDate(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "出院开始时间", Format(Curdate, "YYYY-MM-DD")))
    dtp出院B.Value = datTmp
    datTmp = CDate(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "出院结束时间", Format(dtp出院B.MaxDate, dtp出院E.CustomFormat)))
    dtp出院E.Value = datTmp
    
    On Error GoTo 0
    
    
    Select Case mbytType
        Case 0 '所有病人
            chk登记.Value = 1
            chk出生.Value = 0
            chk入院.Value = 0
            chk出院.Value = 0
        Case 1 '在院病人
            chk登记.Value = 0
            chk出生.Value = 0
            chk入院.Value = 0
            chk出院.Value = 0: chk出院.Tag = 1
        Case 2 '出院病人
            chk登记.Value = 0
            chk出生.Value = 0
            chk入院.Value = 0
            chk出院.Value = 1
        Case 3, 4 '门诊病人
            chk登记.Value = 1
            chk出生.Value = 0
            chk入院.Value = 0: chk入院.Tag = 1
            chk出院.Value = 0: chk出院.Tag = 1
    End Select
    
    If glngSys Like "8??" And Not Visible Then
        chk入院.Visible = False
        dtp入院B.Visible = False
        dtp入院E.Visible = False
        lbl入院.Visible = False
        chk出院.Visible = False
        dtp出院B.Visible = False
        dtp出院E.Visible = False
        lbl出院.Visible = False
        fraBdr.Height = fraBdr.Height - 900
        Me.Height = Me.Height - 900
        cmdOK.Top = cmdOK.Top - 100
        cmdCancel.Top = cmdCancel.Top - 100
        cmdDef.Top = cmdDef.Top - 800
    End If
End Sub

Public Sub MakeFilter()
    
    mstrFilter = ""
    mstrFilterInfo = "" '只连接有索引列的条件
    If chk登记.Value = 1 Then
        mstrFilter = mstrFilter & " And A.登记时间 Between [3] And [4]"
        mstrFilterInfo = mstrFilterInfo & " And A.登记时间 Between [3] And [4]"
    End If
    If chk出生.Value = 1 Then mstrFilter = mstrFilter & " And A.出生日期 Between [5] And [6]"
    If chk入院.Value = 1 Then mstrFilter = mstrFilter & " And P.入院日期 Between [7] And [8]"
    If chk出院.Value = 1 Then mstrFilter = mstrFilter & " And P.出院日期 Between [9] And [10]"
    
    If txt住院号.Text <> "" Then
        mstrFilter = mstrFilter & " And A.病人ID = (Select Nvl(Max(病人ID),0) as 病人ID From 病案主页 Where 住院号=[11])"
        mstrFilterInfo = mstrFilterInfo & " And A.病人ID = (Select Nvl(Max(病人ID),0) as 病人ID From 病案主页 Where 住院号=[11])"
    End If
    If cbo性别.ListIndex <> 0 Then mstrFilter = mstrFilter & " And A.性别=[12]"
    If Trim(txt区域.Text) <> "" Then mstrFilter = mstrFilter & " And A.区域=[13]"
    
    '该条件仅用于特殊病人过滤
    If txt编号.Visible Then
        If txt编号.Text <> "" Then mstrFilter = mstrFilter & " And C.编号=[14]"
    Else
        '不同的查看范围时条件不同
        If cbo费别.ListIndex <> 0 Then
            If mbytType = 0 Or mbytType = 3 Or mbytType = 4 Then
                mstrFilter = mstrFilter & " And A.费别=[14]"
            Else
                mstrFilter = mstrFilter & " And P.费别=[14]"
            End If
        End If
    End If

    If Trim(PatiIdentify.Text) <> "" Then
        If mlngPatiId = 0 Then
            Select Case PatiIdentify.GetCurCard.名称  '"1-姓名;2-就诊卡;3-门诊号;4-医保号;5-身份证号;6-IC卡号;7-手机号"
                Case "姓名"
                    If chk登记.Value = 1 Or chk入院.Value = 1 Or chk出院.Value = 1 Then
                        mstrFilter = Replace(mstrFilter, "登记时间", "登记时间+0") & " And A.姓名 like [15]"
                        mstrFilterInfo = Replace(mstrFilterInfo, "登记时间", "登记时间+0") & " And A.姓名 like [15]"
                    Else
                        mstrFilter = Replace(mstrFilter, "登记时间", "登记时间+0") & " And A.姓名=[15]"
                        mstrFilterInfo = Replace(mstrFilterInfo, "登记时间", "登记时间+0") & " And A.姓名=[15]"
                    End If
                Case "就诊卡"
                    mstrFilter = Replace(mstrFilter, "登记时间", "登记时间+0") & " And A.就诊卡号=[15]"
                    mstrFilterInfo = Replace(mstrFilterInfo, "登记时间", "登记时间+0") & " And A.就诊卡号=[15]"
                Case "门诊号"
                    mstrFilter = Replace(mstrFilter, "登记时间", "登记时间+0") & " And A.门诊号=[15]"
                    mstrFilterInfo = Replace(mstrFilterInfo, "登记时间", "登记时间+0") & " And A.门诊号=[15]"
                Case "医保号"
                    mstrFilter = Replace(mstrFilter, "登记时间", "登记时间+0") & " And A.医保号=[15]"
                    mstrFilterInfo = Replace(mstrFilterInfo, "登记时间", "登记时间+0") & " And A.医保号=[15]"
                Case "二代身份证"
                    mstrFilter = Replace(mstrFilter, "登记时间", "登记时间+0") & " And A.身份证号=[15]"
                    mstrFilterInfo = Replace(mstrFilterInfo, "登记时间", "登记时间+0") & " And A.身份证号=[15]"
                Case "IC卡"
                    mstrFilter = Replace(mstrFilter, "登记时间", "登记时间+0") & " And A.IC卡号=[15]"
                    mstrFilterInfo = Replace(mstrFilterInfo, "登记时间", "登记时间+0") & " And A.IC卡号=[15]"
                Case "手机号"
                    mstrFilter = Replace(mstrFilter, "登记时间", "登记时间+0") & " And A.手机号=[15]"
                    mstrFilterInfo = Replace(mstrFilterInfo, "登记时间", "登记时间+0") & " And A.手机号=[15]"
            End Select
        Else
            mstrFilter = Replace(mstrFilter, "登记时间", "登记时间+0") & " And A.病人ID=[15]"
            mstrFilterInfo = Replace(mstrFilterInfo, "登记时间", "登记时间+0") & " And A.病人ID=[15]"
        End If
    End If
    If cboPayPlan.ListIndex <> 0 Then
        If mbytType = 0 Or mbytType = 3 Or mbytType = 4 Then
            mstrFilter = mstrFilter & " And A.医疗付款方式=[17]"
        Else
            mstrFilter = mstrFilter & " And P.医疗付款方式=[17]"
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbytInFun = 0

    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "出生开始时间", Format(Me.dtp出生B.Value, "YYYY-MM-DD")
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "出生结束时间", Format(Me.dtp出生E.Value, "yyyy-MM-dd 23:59:59")
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "入院开始时间", Format(Me.dtp入院B.Value, "YYYY-MM-DD")
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "入院结束时间", Format(Me.dtp入院E.Value, "yyyy-MM-dd 23:59:59")
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "出院开始时间", Format(Me.dtp出院B.Value, "YYYY-MM-DD")
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "出院结束时间", Format(Me.dtp出院E.Value, "yyyy-MM-dd 23:59:59")
End Sub

Private Sub PatiIdentify_Change()
    mlngPatiId = 0
End Sub

Private Sub PatiIdentify_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    mlngPatiId = 0
    If objHisPati Is Nothing Then Exit Sub
    mlngPatiId = objHisPati.病人ID
End Sub

Private Sub PatiIdentify_FindPatiBefore(ByVal objCard As zlIDKind.Card, blnCard As Boolean, strShowText As String, objCardData As zlIDKind.PatiInfor, blnFindPatied As Boolean, blnCancel As Boolean)
    If mstrFindType = objCard.名称 And InStr(";就诊卡;门诊号;姓名;二代身份证;IC卡;医保号;手机号;", ";" & mstrFindType & ";") > 0 Then
        '选择类型不同
        mlngPatiId = 0
        blnCancel = True
    End If
End Sub

Private Sub PatiIdentify_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    mstrFindType = objCard.名称
    PatiIdentify.Text = ""
    mlngPatiId = 0
End Sub

Private Sub txt编号_GotFocus()
    Call zlControl.TxtSelAll(txt编号)
End Sub

Private Sub txt编号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt区域_GotFocus()
    zlControl.TxtSelAll txt区域
    Call OpenIme(gstrIme)
End Sub

Private Sub txt区域_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt区域.Text <> "" Then
            Set rsTmp = GetArea(Me, txt区域)
            If Not rsTmp Is Nothing Then
                txt区域.Text = rsTmp!名称
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                zlControl.TxtSelAll txt区域
                txt区域.SetFocus
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt区域, KeyAscii
    End If
End Sub

Private Sub txt区域_LostFocus()
    If gstrIme <> "不自动开启" Then Call OpenIme
End Sub

Private Sub txt住院号_GotFocus()
    Call zlControl.TxtSelAll(txt住院号)
End Sub

Private Sub txt住院号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

