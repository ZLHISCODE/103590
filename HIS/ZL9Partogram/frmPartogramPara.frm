VERSION 5.00
Begin VB.Form frmPartogramPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "产程选项"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5355
   Icon            =   "frmPartogramPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fra1 
      Caption         =   "生产曲线标志(异常产)"
      Height          =   1005
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1215
      Width           =   3735
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   8
         ItemData        =   "frmPartogramPara.frx":000C
         Left            =   1110
         List            =   "frmPartogramPara.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   600
         Width           =   2100
      End
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   7
         ItemData        =   "frmPartogramPara.frx":0043
         Left            =   1110
         List            =   "frmPartogramPara.frx":0050
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   2100
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "先露下降"
         Height          =   180
         Index           =   7
         Left            =   360
         TabIndex        =   8
         Top             =   675
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "宫口扩大"
         Height          =   180
         Index           =   6
         Left            =   360
         TabIndex        =   6
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.ComboBox cboBody 
      Height          =   300
      Index           =   6
      ItemData        =   "frmPartogramPara.frx":007A
      Left            =   2205
      List            =   "frmPartogramPara.frx":0084
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   4695
      Width           =   1650
   End
   Begin VB.CheckBox chk 
      Caption         =   "产程图上显示产程时间"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   3405
      Width           =   2490
   End
   Begin VB.CheckBox chk 
      Caption         =   "产程图模式为交叉式(不勾为伴行式)"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   3645
      Width           =   3330
   End
   Begin VB.CheckBox chk 
      Caption         =   "先露高低显示在左侧(不勾为右侧)"
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   3885
      Width           =   3330
   End
   Begin VB.CheckBox chk 
      Caption         =   "产程图上显示警戒线"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   18
      Top             =   4125
      Width           =   3330
   End
   Begin VB.ComboBox cboBody 
      Height          =   300
      Index           =   4
      ItemData        =   "frmPartogramPara.frx":00A0
      Left            =   1200
      List            =   "frmPartogramPara.frx":00AA
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   4365
      Width           =   735
   End
   Begin VB.ComboBox cboBody 
      Height          =   300
      Index           =   5
      ItemData        =   "frmPartogramPara.frx":00BA
      Left            =   3120
      List            =   "frmPartogramPara.frx":00C4
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   4365
      Width           =   735
   End
   Begin VB.Frame fra1 
      Caption         =   "生产曲线标志(顺产)"
      Height          =   1005
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   0
         ItemData        =   "frmPartogramPara.frx":00D4
         Left            =   1110
         List            =   "frmPartogramPara.frx":00E1
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2100
      End
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   1
         ItemData        =   "frmPartogramPara.frx":010B
         Left            =   1110
         List            =   "frmPartogramPara.frx":0118
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   2100
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "宫口扩大"
         Height          =   180
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "先露下降"
         Height          =   180
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   660
         Width           =   720
      End
   End
   Begin VB.Frame fra3 
      Height          =   5040
      Left            =   3960
      TabIndex        =   23
      Top             =   15
      Width           =   15
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4080
      TabIndex        =   24
      Top             =   360
      Width           =   1100
   End
   Begin VB.CommandButton cmdCanc 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4080
      TabIndex        =   25
      Top             =   840
      Width           =   1100
   End
   Begin VB.Frame fra2 
      Caption         =   "生产措施标志"
      Height          =   1005
      Left            =   120
      TabIndex        =   10
      Top             =   2295
      Width           =   3735
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   3
         ItemData        =   "frmPartogramPara.frx":0142
         Left            =   1110
         List            =   "frmPartogramPara.frx":014C
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   600
         Width           =   2100
      End
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   2
         ItemData        =   "frmPartogramPara.frx":0168
         Left            =   1110
         List            =   "frmPartogramPara.frx":0175
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   240
         Width           =   2100
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "标志位置"
         Height          =   180
         Index           =   3
         Left            =   360
         TabIndex        =   13
         Top             =   660
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "标志内容"
         Height          =   180
         Index           =   2
         Left            =   360
         TabIndex        =   11
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "产程图0点与第一个曲线点"
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   26
      Top             =   4755
      Width           =   2070
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "警戒线显示为         异常线显示为"
      Height          =   180
      Index           =   5
      Left            =   120
      TabIndex        =   19
      Top             =   4395
      Width           =   2970
   End
End
Attribute VB_Name = "frmPartogramPara"
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
    Dim curDate As Date, intDay As Integer, lngValue As Long
    Dim intStart As Integer
    
    mblnOK = False
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
    
    '初始体温单标记
    '------------------------------------------------------------------------------------------------------------------
    '顺产
    cboBody(0).Clear
    cboBody(0).AddItem "0-不显示"
    cboBody(0).AddItem "1-显示虚线箭头"
    cboBody(0).AddItem "2-显示实线箭头"
    
    cboBody(1).Clear
    cboBody(1).AddItem "0-不显示"
    cboBody(1).AddItem "1-显示虚线箭头"
    cboBody(1).AddItem "2-显示实线箭头"
    
    '73309:刘鹏飞,2014-06-24
    '异常产
    cboBody(7).Clear
    cboBody(7).AddItem "0-不显示"
    cboBody(7).AddItem "1-显示虚线箭头"
    cboBody(7).AddItem "2-显示实线箭头"
    
    cboBody(8).Clear
    cboBody(8).AddItem "0-不显示"
    cboBody(8).AddItem "1-显示虚线箭头"
    cboBody(8).AddItem "2-显示实线箭头"
    cboBody(8).AddItem "3-显示直角虚线"
    
    cboBody(2).Clear
    cboBody(2).AddItem "0-不显示"
    cboBody(2).AddItem "1-显示生产"
    cboBody(2).AddItem "2-显示处理内容"
    
    cboBody(3).Clear
    cboBody(3).AddItem "0-宫口扩大"
    cboBody(3).AddItem "1-先露下降"
    
    cboBody(4).Clear
    cboBody(4).AddItem "0-虚线"
    cboBody(4).AddItem "1-实线"
    
    cboBody(5).Clear
    cboBody(5).AddItem "0-虚线"
    cboBody(5).AddItem "1-实线"
    
    '73309:刘鹏飞,2014-06-24
    cboBody(6).Clear
    cboBody(6).AddItem "0-不连线"
    cboBody(6).AddItem "1-以虚线连接"
    cboBody(6).AddItem "2-以实线连接"
    '产程生产曲线标志
    strTmp = zlDatabase.GetPara("产程生产曲线标志", glngSys, 1255, "1;1", Array(lbl(0), cboBody(0), lbl(1), cboBody(1)), InStr(mstrPrivs, "护理选项设置") > 0)
    For intLoop = 0 To 1
        If UBound(Split(strTmp, ";")) >= intLoop Then
            lngValue = Val(Split(strTmp, ";")(intLoop))
            If lngValue < 0 Or lngValue > cboBody(intLoop).ListCount - 1 Then lngValue = 0
            cboBody(intLoop).ListIndex = lngValue
        Else
            cboBody(intLoop).ListIndex = 0
        End If
    Next
    strTmp = cboBody(0).ListIndex & ";" & cboBody(1).ListIndex
    strTmp = zlDatabase.GetPara("产程生产曲线标志(异)", glngSys, 1255, strTmp, Array(lbl(6), cboBody(7), lbl(7), cboBody(8)), InStr(mstrPrivs, "护理选项设置") > 0)
    For intLoop = 0 To 1
        If UBound(Split(strTmp, ";")) >= intLoop Then
            lngValue = Val(Split(strTmp, ";")(intLoop))
            If lngValue < 0 Or lngValue > cboBody(intLoop + 7).ListCount - 1 Then lngValue = 0
            cboBody(intLoop + 7).ListIndex = lngValue
        Else
            cboBody(intLoop + 7).ListIndex = 0
        End If
    Next
    
    '产程生产措施标志
    strTmp = zlDatabase.GetPara("产程生产措施标志", glngSys, 1255, "1;1", Array(lbl(2), cboBody(2), lbl(3), cboBody(3)), InStr(mstrPrivs, "护理选项设置") > 0)
    For intLoop = 0 To 1
        If UBound(Split(strTmp, ";")) >= intLoop Then
            lngValue = Val(Split(strTmp, ";")(intLoop))
            If lngValue < 0 Or lngValue > cboBody(intLoop + 2).ListCount - 1 Then lngValue = 0
            cboBody(intLoop + 2).ListIndex = lngValue
        Else
            cboBody(intLoop + 2).ListIndex = 0
        End If
    Next
    '产程警戒线标志
    strTmp = zlDatabase.GetPara("产程警戒异常线标志", glngSys, 1255, "1;1", Array(lbl(5), cboBody(4), cboBody(5)), InStr(mstrPrivs, "护理选项设置") > 0)
    For intLoop = 0 To 1
        If UBound(Split(strTmp, ";")) >= intLoop Then
            lngValue = Val(Split(strTmp, ";")(intLoop))
            If lngValue < 0 Or lngValue > cboBody(intLoop + 4).ListCount - 1 Then lngValue = 0
            cboBody(intLoop + 4).ListIndex = lngValue
        Else
            cboBody(intLoop + 4).ListIndex = 0
        End If
    Next
    
    chk(0).Value = Val(zlDatabase.GetPara("产程图显示产程时间", glngSys, 1255, "1", Array(chk(0)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(1).Value = Val(zlDatabase.GetPara("产程图模式", glngSys, 1255, "0", Array(chk(1)), True))
    chk(2).Value = Val(zlDatabase.GetPara("先露高低显示位置", glngSys, 1255, "0", Array(chk(2)), InStr(mstrPrivs, "护理选项设置") > 0))
    chk(3).Value = Val(zlDatabase.GetPara("产程图显示警戒线", glngSys, 1255, "1", Array(chk(3)), InStr(mstrPrivs, "护理选项设置") > 0))
    
    strTmp = zlDatabase.GetPara("产程曲线点与0点连线", glngSys, 1255, "0", Array(lbl(4), cboBody(6)), InStr(mstrPrivs, "护理选项设置") > 0)
    If Val(strTmp) < 0 Or Val(strTmp) > 2 Then strTmp = "0"
    cboBody(6).ListIndex = Val(strTmp)
    
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
    
    
    strTmp = cboBody(0).ListIndex & ";" & cboBody(1).ListIndex
    Call zlDatabase.SetPara("产程生产曲线标志", strTmp, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    strTmp = cboBody(7).ListIndex & ";" & cboBody(8).ListIndex
    Call zlDatabase.SetPara("产程生产曲线标志(异)", strTmp, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    strTmp = cboBody(2).ListIndex & ";" & cboBody(3).ListIndex
    Call zlDatabase.SetPara("产程生产措施标志", strTmp, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    strTmp = cboBody(4).ListIndex & ";" & cboBody(5).ListIndex
    Call zlDatabase.SetPara("产程警戒异常线标志", strTmp, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
 
    Call zlDatabase.SetPara("产程图显示产程时间", chk(0).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zlDatabase.SetPara("产程图模式", chk(1).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zlDatabase.SetPara("先露高低显示位置", chk(2).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zlDatabase.SetPara("产程图显示警戒线", chk(3).Value, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    Call zlDatabase.SetPara("产程曲线点与0点连线", cboBody(6).ListIndex, glngSys, 1255, InStr(mstrPrivs, "护理选项设置") > 0)
    
    mblnOK = True
    Unload Me
End Sub


