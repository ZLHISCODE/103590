VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmNurse 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "更改护理等级"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   Icon            =   "frmNurse.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   375
      TabIndex        =   11
      Top             =   2415
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   2265
      Left            =   90
      TabIndex        =   12
      Top             =   15
      Width           =   5445
      Begin VB.TextBox txt科室 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3390
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   630
         Width           =   1845
      End
      Begin VB.TextBox txt姓名 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   1515
      End
      Begin VB.TextBox txt性别 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3210
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   675
      End
      Begin VB.TextBox txt年龄 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   4545
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   690
      End
      Begin VB.TextBox txt住院号 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   630
         Width           =   1515
      End
      Begin VB.TextBox txt床号 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1005
         Width           =   1515
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   300
         Left            =   3390
         TabIndex        =   6
         Top             =   1005
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   19
         Format          =   "yyyy-MM-dd HH:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboNew 
         Height          =   300
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1785
         Width           =   4275
      End
      Begin VB.TextBox txtPre 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1395
         Width           =   4260
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前科室"
         Height          =   180
         Left            =   2610
         TabIndex        =   21
         Top             =   690
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "住院号"
         Height          =   180
         Left            =   375
         TabIndex        =   20
         Top             =   690
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   540
         TabIndex        =   19
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   180
         Left            =   2760
         TabIndex        =   18
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Height          =   180
         Left            =   4110
         TabIndex        =   17
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人床位"
         Height          =   180
         Left            =   195
         TabIndex        =   16
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "生效时间"
         Height          =   180
         Left            =   2610
         TabIndex        =   15
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "新护理"
         Height          =   180
         Left            =   375
         TabIndex        =   14
         Top             =   1845
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "原护理"
         Height          =   180
         Left            =   375
         TabIndex        =   13
         Top             =   1455
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4095
      TabIndex        =   10
      Top             =   2415
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2895
      TabIndex        =   9
      Top             =   2415
      Width           =   1100
   End
End
Attribute VB_Name = "frmNurse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mblnBed As Boolean
Private mlng病人ID As Long, mlng主页ID As Long

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, strLevel, lng护理ID As Long
    
    On Error GoTo errH
    
    gblnOK = False
    
    txtDate.Text = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    
    With frmManageCourse
        If mblnBed Then
            txt姓名.Text = .mrsBeds!姓名
            txt性别.Text = IIf(IsNull(.mrsBeds!性别), "", .mrsBeds!性别)
            txt年龄.Text = IIf(IsNull(.mrsBeds!年龄), "", .mrsBeds!年龄)
            txt住院号.Text = IIf(IsNull(.mrsBeds!住院号), "", .mrsBeds!住院号)
            txt科室.Text = .mrsBeds!当前科室
            txtPre.Text = Nvl(.mrsBeds!护理等级)
            
            '可能包房
            .mrsCBeds.Filter = "病人ID=" & .mrsBeds!病人ID
            Do While Not .mrsCBeds.EOF
                txt床号.Text = txt床号.Text & "," & .mrsCBeds!床号
                .mrsCBeds.MoveNext
            Loop
            txt床号.Text = Mid(txt床号.Text, 2)
            
            mlng病人ID = .mrsBeds!病人ID
            mlng主页ID = .mrsBeds!主页ID
            lng护理ID = Nvl(.mrsBeds!护理等级ID, 0)
        Else
            txt姓名.Text = .mrsFamily!姓名
            txt性别.Text = IIf(IsNull(.mrsFamily!性别), "", .mrsFamily!性别)
            txt年龄.Text = IIf(IsNull(.mrsFamily!年龄), "", .mrsFamily!年龄)
            txt住院号.Text = IIf(IsNull(.mrsFamily!住院号), "", .mrsFamily!住院号)
            txt科室.Text = .mrsFamily!当前科室
            txt床号.Text = "家庭病床"
            txtPre.Text = Nvl(.mrsFamily!护理等级)
            
            mlng病人ID = .mrsFamily!病人ID
            mlng主页ID = .mrsFamily!主页ID
            lng护理ID = Nvl(.mrsFamily!护理等级ID, 0)
        End If
        gstrSQL = "Select ID as 序号,编码,名称 From 收费项目目录" & _
            " Where 类别='H' And 项目特性>=1 And (撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or 撤档时间 is NULL) And ID<>[1]  Order by 编码"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng护理ID)
    End With
    
'    Set rsTmp = New ADODB.Recordset
'    rsTmp.CursorLocation = adUseClient
'    Call SQLTest(App.ProductName, Me.Caption, gstrSQL)
'    rsTmp.Open gstrSQL, gcnOracle, adOpenKeyset
'    Call SQLTest
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboNew.AddItem rsTmp!编码 & "-" & rsTmp!名称
            cboNew.ItemData(i - 1) = rsTmp!序号
            rsTmp.MoveNext
        Next
        cboNew.ListIndex = 0
    Else
        MsgBox "不能读取护理等级数据,请先到护理等级管理中设置！", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_LostFocus()
    If Not IsDate(txtDate.Text) Then txtDate.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim dMax As Date, strSql As String, Curdate As Date
    
    If cboNew.ListIndex = -1 Then
        MsgBox "请选择新的护理等级！", vbInformation, gstrSysName
        cboNew.SetFocus: Exit Sub
    End If
    If Not IsDate(txtDate.Text) Then
        MsgBox "请输入合法的生效时间！", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    
    dMax = GetMaxDate(mlng病人ID, mlng主页ID)
    If CDate(txtDate.Text) <= dMax Then
        MsgBox "生效时间必须大于该病人上次变动时间 " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & " ！", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
        
    '时间不能超过当前时间太长(一个月)
    Curdate = zlDatabase.Currentdate
    If CDate(txtDate.Text) > Curdate Then
        If CDate(txtDate.Text) - Curdate > 30 Then
            MsgBox "生效时间比当前时间大得过多,请检查！", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        If MsgBox("生效时间大于了当前系统时间,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txtDate.SetFocus: Exit Sub
        End If
    End If
        
    strSql = "zl_病人变动记录_Nurse(" & mlng病人ID & "," & mlng主页ID & "," & _
        cboNew.ItemData(cboNew.ListIndex) & ",To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS')," & _
        "'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
    
    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    gblnOK = True
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
