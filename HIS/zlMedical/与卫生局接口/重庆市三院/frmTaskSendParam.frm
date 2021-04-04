VERSION 5.00
Begin VB.Form frmTaskSendFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "查找"
   ClientHeight    =   1785
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5190
   Icon            =   "frmTaskSendParam.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   1515
      Left            =   60
      TabIndex        =   7
      Top             =   90
      Width           =   3720
      Begin VB.CheckBox chk 
         Caption         =   "包括已发送"
         Height          =   270
         Left            =   1215
         TabIndex        =   4
         Top             =   1125
         Width           =   1260
      End
      Begin VB.ComboBox cbo 
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   0
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   660
         Width           =   2400
      End
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   255
         Width           =   2400
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检部门(&D)"
         Height          =   180
         Index           =   0
         Left            =   135
         TabIndex        =   0
         Top             =   300
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体检时间(&U)"
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   735
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3900
      TabIndex        =   5
      Top             =   180
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3900
      TabIndex        =   6
      Top             =   600
      Width           =   1100
   End
End
Attribute VB_Name = "frmTaskSendFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnStartUp As Boolean
Private mblnOK As Boolean

Public Function ShowFilter(ByVal frmMain As Object) As Boolean
    
    mblnOK = False
    
    If InitActivate = False Then Exit Function
    If LoadData = False Then Exit Function
        
    Me.Show 1, frmMain
    
    ShowFilter = mblnOK
    
End Function

Private Function InitActivate() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  初始化数据，发生在窗体的Activate事件
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    gstrSQL = "SELECT A.编码||'-'||A.名称,ID FROM 部门表 A,部门性质说明 B WHERE (A.撤档时间 IS NULL OR A.撤档时间 =TO_DATE('3000-01-01','YYYY-MM-DD')) AND A.ID=B.部门ID AND B.工作性质='体检' ORDER BY A.编码||'-'||A.名称"
    
    Call OpenRecordset(rs, Me.Caption)
    If rs.BOF Then
        ShowSimpleMsg "没有体检性质的部门，请在部门管理中设置！"
        Exit Function
    End If
    
    '绑定数据到控件中
    Call AddComboData(cboDept, rs)
    
    '初始选择数据处理
    CboLocate cboDept, UserInfo.部门ID, True
    If cboDept.ListCount > 0 And cboDept.ListIndex = -1 Then cboDept.ListIndex = 0
    
    cbo(0).AddItem "今  天"
    cbo(0).AddItem "昨  天"
    cbo(0).AddItem "本  周"
    cbo(0).AddItem "本  月"
    cbo(0).AddItem "本  季"
    cbo(0).AddItem "本半年"
    cbo(0).AddItem "本  年"
    cbo(0).AddItem "前三天"
    cbo(0).AddItem "前一周"
    cbo(0).AddItem "前半月"
    cbo(0).AddItem "前一月"
    cbo(0).AddItem "前二月"
    cbo(0).AddItem "前三月"
    cbo(0).AddItem "前半年"
    cbo(0).AddItem "前一年"
            
    
    InitActivate = True
    
    Exit Function
    
errHand:
'    If ErrCenter = 1 Then
        'Resume
    'End If
End Function

Private Function LoadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  装载数据
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHand
    
    CboLocate cboDept, Val(GetSetting("ZLSOFT", "公共全局\干保接口", "体检部门", "0")), True
    
    chk.Value = Val(GetSetting("ZLSOFT", "公共全局\干保接口", "包括已发送", "0"))
    
    On Error Resume Next
    cbo(0).Text = GetSetting("ZLSOFT", "公共全局\干保接口", "体检时间", "今  天")
    If cbo(0).ListIndex = -1 Then cbo(0).ListIndex = 0
        
    LoadData = True
    Exit Function
    
errHand:
'    If ErrCenter = 1 Then
        'Resume
    'End If
End Function

Private Function SaveData() As Boolean

    On Error GoTo errHand
    
    
    If cboDept.ListIndex = -1 Then
        Call SaveSetting("ZLSOFT", "公共全局\干保接口", "体检部门", "0")
    Else
        Call SaveSetting("ZLSOFT", "公共全局\干保接口", "体检部门", cboDept.ItemData(cboDept.ListIndex))
    End If
    
    Call SaveSetting("ZLSOFT", "公共全局\干保接口", "包括已发送", chk.Value)
    Call SaveSetting("ZLSOFT", "公共全局\干保接口", "体检时间", cbo(0).Text)
    
    SaveData = True
    
    Exit Function
    
errHand:
'    If ErrCenter = 1 Then
        'Resume
    'End If
    
End Function


Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        PressKey vbKeyTab
    End If
End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        PressKey vbKeyTab
    End If
End Sub

Private Sub chk_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If SaveData Then
        mblnOK = True
        Unload Me
    End If
End Sub




