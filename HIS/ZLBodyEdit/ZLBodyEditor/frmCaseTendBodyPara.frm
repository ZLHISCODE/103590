VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCaseTendBodyPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "护理选项"
   ClientHeight    =   5970
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5550
   Icon            =   "frmCaseTendBodyPara.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox chk 
      Caption         =   "一个护理级别只显示一份护理文件"
      Height          =   315
      Index           =   6
      Left            =   165
      TabIndex        =   31
      Top             =   4560
      Width           =   4500
   End
   Begin VB.CheckBox chk 
      Caption         =   "婴儿体温单显示出院信息"
      Height          =   180
      Index           =   5
      Left            =   165
      TabIndex        =   33
      Top             =   5160
      Width           =   5190
   End
   Begin VB.CheckBox chk 
      Caption         =   "呼吸项表格数据打印输出时上下显示（无数据继承)"
      Height          =   315
      Index           =   4
      Left            =   165
      TabIndex        =   32
      Top             =   4830
      Width           =   5190
   End
   Begin VB.CheckBox chk 
      Caption         =   "按护理级别变化分开显示护理文件（体温除外）"
      Height          =   315
      Index           =   1
      Left            =   165
      TabIndex        =   30
      Top             =   4275
      Width           =   4500
   End
   Begin VB.Frame fra 
      Caption         =   "体温自动标志"
      Height          =   1770
      Index           =   15
      Left            =   150
      TabIndex        =   37
      Top             =   60
      Width           =   5310
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   7
         ItemData        =   "frmCaseTendBodyPara.frx":000C
         Left            =   3075
         List            =   "frmCaseTendBodyPara.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1395
         Width           =   2100
      End
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   6
         ItemData        =   "frmCaseTendBodyPara.frx":0010
         Left            =   525
         List            =   "frmCaseTendBodyPara.frx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1410
         Width           =   2100
      End
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   0
         ItemData        =   "frmCaseTendBodyPara.frx":0014
         Left            =   525
         List            =   "frmCaseTendBodyPara.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   2100
      End
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   1
         ItemData        =   "frmCaseTendBodyPara.frx":0018
         Left            =   3075
         List            =   "frmCaseTendBodyPara.frx":001A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   270
         Width           =   2100
      End
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   2
         ItemData        =   "frmCaseTendBodyPara.frx":001C
         Left            =   525
         List            =   "frmCaseTendBodyPara.frx":001E
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   660
         Width           =   2100
      End
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   3
         ItemData        =   "frmCaseTendBodyPara.frx":0020
         Left            =   3075
         List            =   "frmCaseTendBodyPara.frx":0022
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   660
         Width           =   2100
      End
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   4
         ItemData        =   "frmCaseTendBodyPara.frx":0024
         Left            =   525
         List            =   "frmCaseTendBodyPara.frx":0026
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1050
         Width           =   2100
      End
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   5
         ItemData        =   "frmCaseTendBodyPara.frx":0028
         Left            =   3075
         List            =   "frmCaseTendBodyPara.frx":002A
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1050
         Width           =   2100
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生"
         Height          =   180
         Index           =   2
         Left            =   2700
         TabIndex        =   14
         Top             =   1455
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "分娩"
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   12
         Top             =   1470
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院"
         Height          =   180
         Index           =   44
         Left            =   135
         TabIndex        =   0
         Top             =   315
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入科"
         Height          =   180
         Index           =   45
         Left            =   2700
         TabIndex        =   2
         Top             =   315
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "转科"
         Height          =   180
         Index           =   46
         Left            =   135
         TabIndex        =   4
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "换床"
         Height          =   180
         Index           =   48
         Left            =   2700
         TabIndex        =   6
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "手术"
         Height          =   180
         Index           =   49
         Left            =   135
         TabIndex        =   8
         Top             =   1110
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院"
         Height          =   180
         Index           =   50
         Left            =   2700
         TabIndex        =   10
         Top             =   1110
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3090
      TabIndex        =   34
      Top             =   5505
      Width           =   1100
   End
   Begin VB.CommandButton cmdCanc 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4335
      TabIndex        =   35
      Top             =   5505
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   2475
      Index           =   0
      Left            =   150
      TabIndex        =   36
      Top             =   1755
      Width           =   5310
      Begin VB.CheckBox chk 
         Caption         =   "体温单上显示病人的诊断信息"
         Height          =   315
         Index           =   3
         Left            =   210
         TabIndex        =   29
         Top             =   2115
         Width           =   4950
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1665
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   26
         Text            =   "1"
         Top             =   1515
         Width           =   420
      End
      Begin VB.CheckBox chk 
         Caption         =   "未记说明显示在体温单的下面（不勾此项时显示在上面）"
         Height          =   315
         Index           =   2
         Left            =   210
         TabIndex        =   28
         Top             =   1815
         Width           =   4950
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2580
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   23
         Text            =   "0"
         Top             =   1185
         Width           =   375
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   20
         Text            =   "0"
         Top             =   885
         Width           =   420
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   0
         Left            =   1155
         TabIndex        =   17
         Text            =   "14"
         Top             =   270
         Width           =   255
      End
      Begin VB.CheckBox chk 
         Caption         =   "手术后标注天数内再次手术时,停止前一次手术标注"
         Height          =   375
         Index           =   0
         Left            =   195
         TabIndex        =   18
         Top             =   510
         Width           =   4500
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   6
         Left            =   2085
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   885
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   2
         BuddyControl    =   "txt(6)"
         BuddyDispid     =   196615
         BuddyIndex      =   6
         OrigLeft        =   2190
         OrigTop         =   870
         OrigRight       =   2430
         OrigBottom      =   1170
         Max             =   4
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   1
         Left            =   2970
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1185
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   2
         BuddyControl    =   "txt(1)"
         BuddyDispid     =   196615
         BuddyIndex      =   1
         OrigLeft        =   2190
         OrigTop         =   870
         OrigRight       =   2430
         OrigBottom      =   1170
         Max             =   30
         Min             =   2
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   0
         Left            =   2100
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1515
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   2
         BuddyControl    =   "txt(2)"
         BuddyDispid     =   196615
         BuddyIndex      =   2
         OrigLeft        =   2190
         OrigTop         =   870
         OrigRight       =   2430
         OrigBottom      =   1170
         Max             =   30
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "允许录入超过当前        天的护理记录数据"
         Height          =   180
         Index           =   4
         Left            =   210
         TabIndex        =   25
         Top             =   1560
         Width           =   3600
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体温表输出时，表格数据固定        行"
         Height          =   180
         Index           =   3
         Left            =   210
         TabIndex        =   22
         Top             =   1230
         Width           =   3240
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体温开始记录时间"
         Height          =   180
         Index           =   31
         Left            =   210
         TabIndex        =   19
         Top             =   945
         Width           =   1440
      End
      Begin VB.Line Line1 
         X1              =   1125
         X2              =   1410
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "手术后标注    天"
         Height          =   180
         Index           =   0
         Left            =   195
         TabIndex        =   16
         Top             =   270
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmCaseTendBodyPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mfrmMain As Object
Private mblnOK As Boolean
Private mstrPrivs As String

Public Function ShowPara(ByVal frmMain As Object, ByVal strPrivs As String) As Boolean
    Dim intLoop As Integer
    Dim strTmp As String
    Dim strSQL As String, strPar As String
    Dim curDate As Date, intDay As Integer
    
    mblnOK = False
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
    
    '初始体温单标记
    '------------------------------------------------------------------------------------------------------------------
    cboBody(0).AddItem "0-不显示"
    cboBody(0).AddItem "1-显示说明"
    cboBody(0).AddItem "2-显示说明和时间"
    
    cboBody(1).AddItem "0-不显示"
    cboBody(1).AddItem "1-显示说明"
    cboBody(1).AddItem "2-显示说明和时间"
    
    cboBody(2).AddItem "0-不显示"
    cboBody(2).AddItem "1-显示说明"
    cboBody(2).AddItem "2-显示说明和时间"
    cboBody(2).AddItem "3-显示说明和科室"
    cboBody(2).AddItem "4-显示说明,科室,时间"
    
    cboBody(3).AddItem "0-不显示"
    cboBody(3).AddItem "1-显示说明"
    cboBody(3).AddItem "2-显示说明和时间"
    
    cboBody(4).AddItem "0-不显示"
    cboBody(4).AddItem "1-显示说明"
    cboBody(4).AddItem "2-显示说明和时间"
    
    cboBody(5).AddItem "0-不显示"
    cboBody(5).AddItem "1-显示说明"
    cboBody(5).AddItem "2-显示说明和时间"
    
    cboBody(6).AddItem "0-不显示"
    cboBody(6).AddItem "1-显示说明"
    cboBody(6).AddItem "2-显示说明和时间"
    
    cboBody(7).AddItem "0-不显示"
    cboBody(7).AddItem "1-显示说明"
    cboBody(7).AddItem "2-显示说明和时间"
    
    txt(6).Text = zlDatabase.GetPara("体温开始时间", glngSys, 1255, 4, Array(txt(6), ud(6), lbl(31)), InStr(mstrPrivs, "护理选项设置") > 0)
    txt(1).Text = zlDatabase.GetPara("体温表格行数", glngSys, 1255, 8, Array(txt(1), ud(1), lbl(3)), InStr(mstrPrivs, "护理选项设置") > 0)
    strTmp = zlDatabase.GetPara("体温单标记", glngSys, 1255, "1;1;1;1;1;1;1:1", Array(cboBody(0), cboBody(1), cboBody(2), cboBody(3), cboBody(4), cboBody(5), cboBody(6), cboBody(7)), InStr(mstrPrivs, "护理选项设置") > 0)
    
    For intLoop = 0 To 7
        If UBound(Split(strTmp, ";")) >= intLoop Then
            cboBody(intLoop).ListIndex = Val(Split(strTmp, ";")(intLoop))
        Else
            cboBody(intLoop).ListIndex = 0
        End If
    Next
    
    txt(0).Text = Val(zlDatabase.GetPara("手术后标注天数", glngSys, 1255, "10", Array(txt(0), lbl(0)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(0).Value = Val(zlDatabase.GetPara("再次手术停止前次标注", glngSys, 1255, "0", Array(chk(0)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(1).Value = Val(zlDatabase.GetPara("按护理级别分组", glngSys, 1255, "0", Array(chk(1)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(2).Value = Val(zlDatabase.GetPara("未记说明显示位置", glngSys, 1255, "0", Array(chk(2)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(3).Value = Val(zlDatabase.GetPara("体温单显示诊断", glngSys, 1255, "1", Array(chk(3)), InStr(mstrPrivs, "护理选项设置") > 0))
    txt(2).Text = Val(zlDatabase.GetPara("超期录入护理数据天数", glngSys, 1255, "1", Array(txt(2), lbl(4)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(4).Value = Val(zlDatabase.GetPara("呼吸表格输出", glngSys, 1255, "0", Array(chk(4)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(5).Value = Val(zlDatabase.GetPara("婴儿体温单显示出院信息", glngSys, 1255, "1", Array(chk(5)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(6).Value = Val(zlDatabase.GetPara("显示一份护理文件", glngSys, 1255, "1", Array(chk(6)), InStr(mstrPrivs, "护理选项设置") > 0))
    
    Me.Show 1, mfrmMain
    ShowPara = mblnOK
    
End Function

Private Sub cboBody_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCanc_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    Dim strTmp As String
        
    strTmp = cboBody(0).ListIndex & ";" & cboBody(1).ListIndex & ";" & cboBody(2).ListIndex & ";" & cboBody(3).ListIndex & ";" & cboBody(4).ListIndex & ";" & cboBody(5).ListIndex & ";" & cboBody(6).ListIndex & ";" & cboBody(7).ListIndex
    Call zlDatabase.SetPara("体温开始时间", Val(txt(6).Text), glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zlDatabase.SetPara("体温表格行数", Val(txt(1).Text), glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zlDatabase.SetPara("体温单标记", strTmp, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zlDatabase.SetPara("手术后标注天数", Val(txt(0).Text), glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zlDatabase.SetPara("超期录入护理数据天数", Val(txt(2).Text), glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zlDatabase.SetPara("再次手术停止前次标注", chk(0).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zlDatabase.SetPara("未记说明显示位置", chk(2).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zlDatabase.SetPara("体温单显示诊断", chk(3).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zlDatabase.SetPara("按护理级别分组", chk(1).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zlDatabase.SetPara("呼吸表格输出", chk(4).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zlDatabase.SetPara("婴儿体温单显示出院信息", chk(5).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zlDatabase.SetPara("显示一份护理文件", chk(6).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    
    mblnOK = True
    
    Unload Me
End Sub

Private Sub txt_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txt(Index))
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

