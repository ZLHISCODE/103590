VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTechnoFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤设置"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   ControlBox      =   0   'False
   Icon            =   "frmTechnoFilter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdDef 
      Caption         =   "缺省(&D)"
      Height          =   350
      Left            =   4980
      TabIndex        =   11
      Top             =   1470
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   2445
      Left            =   120
      TabIndex        =   12
      Top             =   15
      Width           =   4680
      Begin VB.CheckBox chk记帐 
         Caption         =   "记帐单据"
         Height          =   210
         Left            =   3210
         TabIndex        =   2
         Top             =   360
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.TextBox txt姓名 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   3000
         MaxLength       =   100
         TabIndex        =   7
         Top             =   1500
         Width           =   1470
      End
      Begin VB.TextBox txt住院号 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   975
         MaxLength       =   18
         TabIndex        =   6
         Top             =   1500
         Width           =   1470
      End
      Begin VB.ComboBox cbo操作员 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1920
         Width           =   1470
      End
      Begin VB.CheckBox chk销帐 
         Caption         =   "销帐单据"
         Height          =   210
         Left            =   3210
         TabIndex        =   3
         Top             =   660
         Width           =   1020
      End
      Begin VB.TextBox txtNOBegin 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   975
         MaxLength       =   8
         TabIndex        =   4
         Top             =   1098
         Width           =   1470
      End
      Begin VB.TextBox txtNoEnd 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   5
         Top             =   1098
         Width           =   1470
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   975
         TabIndex        =   1
         Top             =   684
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   113115139
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   975
         TabIndex        =   0
         Top             =   270
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   113115139
         CurrentDate     =   36588
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   2595
         TabIndex        =   19
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         Height          =   180
         Left            =   360
         TabIndex        =   18
         Top             =   1560
         Width           =   540
      End
      Begin VB.Label lbl操作员 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "操作员"
         Height          =   180
         Left            =   360
         TabIndex        =   17
         Top             =   1980
         Width           =   540
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开始时间"
         Height          =   180
         Left            =   180
         TabIndex        =   16
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结束时间"
         Height          =   180
         Left            =   180
         TabIndex        =   15
         Top             =   744
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   2655
         TabIndex        =   14
         Top             =   1158
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据号"
         Height          =   180
         Left            =   360
         TabIndex        =   13
         Top             =   1158
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4980
      TabIndex        =   10
      Top             =   765
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4980
      TabIndex        =   9
      Top             =   345
      Width           =   1100
   End
End
Attribute VB_Name = "frmTechnoFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mstrFilter As String
Public mlngDept As Long
Public mblnDateMoved As Boolean '当前所选条件的数据是否在后备数据表中

Private Sub cbo操作员_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii >= 32 Then
        lngIdx = zlControl.CboMatchIndex(cbo操作员.hWnd, KeyAscii)
        If lngIdx = -1 And cbo操作员.ListCount > 0 Then lngIdx = 0
        cbo操作员.ListIndex = lngIdx
    End If
End Sub

Private Sub chk记帐_Click()
    If chk记帐.Value = 0 And chk销帐.Value = 0 Then
        chk记帐.Value = 1
    End If
End Sub

Private Sub chk销帐_Click()
    If chk记帐.Value = 0 And chk销帐.Value = 0 Then
        chk销帐.Value = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    gblnOK = False
    Hide
End Sub

Private Sub cmdDef_Click()
    Form_Load
End Sub



Private Sub cmdOK_Click()
    If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
        If txtNoEnd.Text < txtNOBegin.Text Then
            MsgBox "结束单据号不能小于开始单据号！", vbInformation, gstrSysName
            txtNoEnd.SetFocus: Exit Sub
        End If
    End If
    
    Call MakeFilter
    
    gblnOK = True
    Hide
End Sub

Private Sub dtpEnd_Change()
    dtpBegin.MaxDate = dtpEnd.Value
End Sub

Private Sub Form_Activate()
    dtpBegin.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim Curdate As Date, i As Long
    
    gblnOK = False
    
    txtNOBegin.Text = ""
    txtNoEnd.Text = ""
    txt住院号.Text = ""
    txt姓名.Text = ""
    
    chk记帐.Value = 1
    chk销帐.Value = 0
    
    '设置初始值
    
    Curdate = zlDatabase.Currentdate
    dtpBegin.MaxDate = Format(Curdate, "yyyy-MM-dd 23:59:59")
    dtpBegin.Value = Format(Curdate, "yyyy-MM-dd 00:00:00")
    dtpEnd.Value = dtpBegin.MaxDate
    
    Call LoadOper
End Sub

Public Function LoadOper() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    
    On Error GoTo errH
    
    cbo操作员.Clear
    cbo操作员.AddItem "所有操作员"
    cbo操作员.ListIndex = 0
    
    If mlngDept = 0 Then
        strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & _
            " From 人员表 A Where (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
            " Order by A.简码"
    Else
        strSQL = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & _
            " From 人员表 A,部门人员 C" & _
            " Where A.ID=C.人员ID And C.部门ID=[1] And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
            " Order by A.简码"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDept)
    
    For i = 1 To rsTmp.RecordCount
        cbo操作员.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        cbo操作员.ItemData(cbo操作员.NewIndex) = rsTmp!ID
        If rsTmp!ID = UserInfo.ID Then cbo操作员.ListIndex = cbo操作员.NewIndex
        rsTmp.MoveNext
    Next
    
    LoadOper = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    mlngDept = 0
End Sub

Private Sub txtNOBegin_Change()
    txtNoEnd.Enabled = Not (Trim(txtNOBegin.Text) = "")
    If Trim(txtNOBegin.Text = "") Then txtNoEnd.Text = ""
End Sub

Private Sub txtNOBegin_GotFocus()
    zlControl.TxtSelAll txtNOBegin
End Sub

Private Sub txtNOBegin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '46516
    zlControl.TxtCheckKeyPress txtNOBegin, KeyAscii, m文本式
End Sub

Private Sub txtNOBegin_LostFocus()
     If txtNOBegin.Text <> "" Then txtNOBegin.Text = GetFullNO(txtNOBegin.Text, 14)
End Sub

Private Sub txtNOEnd_LostFocus()
    If txtNoEnd.Text <> "" Then txtNoEnd.Text = GetFullNO(txtNoEnd.Text, 14)
End Sub

Private Sub txtNoEnd_GotFocus()
    zlControl.TxtSelAll txtNoEnd
End Sub

Private Sub txtNoEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '46516
    zlControl.TxtCheckKeyPress txtNoEnd, KeyAscii, m文本式
End Sub

Private Sub MakeFilter()
    mstrFilter = " And 登记时间 Between [1] And [2]"
    
    mblnDateMoved = zlDatabase.DateMoved(Format(IIf(dtpBegin.Value < dtpEnd.Value, dtpBegin.Value, dtpEnd.Value), dtpBegin.CustomFormat), , , Me.Caption)
    
    If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
        mstrFilter = mstrFilter & " And NO Between [3] And [4]"
    ElseIf txtNOBegin.Text <> "" Then
        mstrFilter = mstrFilter & " And NO=[3]"
    End If
    
    If chk记帐.Value = 1 And chk销帐.Value = 1 Then
        mstrFilter = mstrFilter & " And 记录状态 IN(1,2,3)"
    ElseIf chk记帐.Value = 1 Then
        mstrFilter = mstrFilter & " And 记录状态 IN(1,3)"
    Else
        mstrFilter = mstrFilter & " And 记录状态=2"
    End If
    
    If IsNumeric(txt住院号.Text) Then
        mstrFilter = mstrFilter & " And 病人ID = (Select Distinct 病人ID From 病案主页 Where 住院号 = [5])"
    End If
    
    If txt姓名.Text <> "" Then
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(txt姓名.Text, 1))) > 0 Then
            mstrFilter = mstrFilter & " And Upper(姓名) Like [6]"
        Else
            mstrFilter = mstrFilter & " And 姓名 Like [6]"
        End If
    End If
    
    If cbo操作员.ListIndex <> -1 Then
        If cbo操作员.ItemData(cbo操作员.ListIndex) <> 0 Then
            mstrFilter = mstrFilter & " And 操作员姓名||''=[7]"
        End If
    End If
    
    
End Sub

Private Sub txt住院号_GotFocus()
    zlControl.TxtSelAll txt住院号
End Sub

Private Sub txt住院号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt姓名_GotFocus()
    zlControl.TxtSelAll txt姓名
End Sub
