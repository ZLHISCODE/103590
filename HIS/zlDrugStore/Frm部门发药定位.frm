VERSION 5.00
Begin VB.Form Frm部门发药定位 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "定位"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "Frm部门发药定位.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txt住院号 
      Height          =   300
      Left            =   1140
      TabIndex        =   16
      Top             =   2940
      Width           =   2235
   End
   Begin VB.TextBox Txt开单医生 
      Height          =   300
      Left            =   1140
      MaxLength       =   8
      TabIndex        =   5
      Top             =   990
      Width           =   1215
   End
   Begin VB.ComboBox Cob类型 
      Height          =   300
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   210
      Width           =   2205
   End
   Begin VB.TextBox TxtNO 
      Height          =   300
      Left            =   1140
      MaxLength       =   8
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox Cob科室 
      Height          =   300
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1380
      Width           =   2205
   End
   Begin VB.TextBox Txt药品 
      Height          =   300
      Left            =   1140
      TabIndex        =   9
      Top             =   1770
      Width           =   1935
   End
   Begin VB.TextBox Txt床号 
      Height          =   300
      Left            =   1140
      TabIndex        =   12
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Txt姓名 
      Height          =   300
      Left            =   1140
      TabIndex        =   14
      Top             =   2550
      Width           =   1215
   End
   Begin VB.CommandButton cmd药品 
      Caption         =   "…"
      Height          =   300
      Left            =   3060
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1770
      Width           =   285
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3780
      TabIndex        =   17
      Top             =   210
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3780
      TabIndex        =   18
      Top             =   690
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "条件"
      Height          =   3855
      Left            =   3540
      TabIndex        =   19
      Top             =   -120
      Width           =   45
   End
   Begin VB.Label lbl住院号 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "住院号(&S)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   270
      TabIndex        =   15
      Top             =   3000
      Width           =   810
   End
   Begin VB.Label Lbl开单医生 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "医生(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   4
      Top             =   1050
      Width           =   630
   End
   Begin VB.Label Lbl类型 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "类型(&T)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   0
      Top             =   270
      Width           =   630
   End
   Begin VB.Label LblNO 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&NO"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   900
      TabIndex        =   2
      Top             =   660
      Width           =   180
   End
   Begin VB.Label Lbl科室 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "科室(&K)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   6
      Top             =   1440
      Width           =   630
   End
   Begin VB.Label Lbl药品 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "药品(&P)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   8
      Top             =   1830
      Width           =   630
   End
   Begin VB.Label Lbl床号 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "床号(&B)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   11
      Top             =   2220
      Width           =   630
   End
   Begin VB.Label Lbl姓名 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "姓名(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   13
      Top             =   2610
      Width           =   630
   End
End
Attribute VB_Name = "Frm部门发药定位"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strReturn As String
Private lng药房ID As Long
Private mstrPrivs As String
Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub cmdOK_Click()
    '组织查找串
    strReturn = ""
    If Cob类型.ListIndex <> 0 Then
        strReturn = strReturn & IIf(strReturn = "", "", " And ") & _
        "类型='" & Mid(Cob类型.Text, InStr(1, Cob类型.Text, "-") + 1) & "'"
    End If
    If Cob科室.ListIndex <> 0 Then
        strReturn = strReturn & IIf(strReturn = "", "", " And ") & _
        "科室='" & Mid(Cob科室.Text, InStr(1, Cob科室.Text, "-") + 1) & "'"
    End If
    If Trim(TxtNo) <> "" Then strReturn = strReturn & IIf(strReturn = "", "", " And ") & "NO='" & TxtNo.Text & "'"
    If Trim(Txt开单医生) <> "" Then strReturn = strReturn & IIf(strReturn = "", "", " And ") & "开单医生 Like '" & Txt开单医生.Text & "%'"
    If Val(txt药品.Tag) <> 0 Then strReturn = strReturn & IIf(strReturn = "", "", " And ") & "药品ID=" & txt药品.Tag
    If Trim(Txt床号) <> "" Then strReturn = strReturn & IIf(strReturn = "", "", " And ") & "床号='" & Txt床号.Text & "'"
    If Trim(Txt姓名) <> "" Then strReturn = strReturn & IIf(strReturn = "", "", " And ") & "姓名='" & Txt姓名.Text & "'"
    If Trim(txt住院号) <> "" Then strReturn = strReturn & IIf(strReturn = "", "", " And ") & "住院号=" & txt住院号.Text
    
    If strReturn = "" Then
        MsgBox "请输入需要查找的内容！", vbInformation, gstrSysName
        Cob类型.SetFocus
        Exit Sub
    End If
    
    Unload Me
End Sub

Private Sub cmd药品_Click()
    Dim RecReturn As New ADODB.Recordset
    
'    With Frm药品选择器
'        Set RecReturn = .ShowME(Me, 1, lng药房ID, , , False)
'    End With
    
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(2, "药品部门发药", lng药房ID, lng药房ID)
    End If
    Set RecReturn = frmSelector.ShowME(Me, 0, 2, , , , lng药房ID, , , False, , , , , False)
        
    With RecReturn
        If .EOF Then Exit Sub
        txt药品.Tag = !药品ID
        txt药品 = "[" & !药品编码 & "]" & IIf(IsNull(!通用名), "", !通用名)
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsPart As New ADODB.Recordset
    strReturn = ""
    
    On Error GoTo errHandle
    Me.Txt开单医生.Enabled = IsHavePrivs(mstrPrivs, "医生查询")
    
    With Cob类型
        .Clear
        .AddItem "所有单据"
        .AddItem "门诊记帐单"
        .AddItem "住院记帐单"
        .AddItem "记帐表"
        .AddItem "医嘱-长嘱"
        .AddItem "医嘱-临嘱"
        .ListIndex = 0
    End With
    
    '检测科室设置否(临床、手术)
    Cob科室.Clear
    Cob科室.AddItem "所有科室"
    
    gstrSQL = " Select 编码||'-'||名称 科室,ID From 部门表 " & _
             " Where ID in (Select 部门ID From 部门性质说明 Where 工作性质 In ('临床','手术') And 服务对象 IN(2,3))" & _
             " And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) " & _
             " Order By 编码||'-'||名称 "

    Set rsPart = zldatabase.OpenSQLRecord(gstrSQL, "Form_Load")
    With rsPart
        If .EOF Then
            MsgBox "请初始化临床科室及手术科室！（部门管理）", vbInformation, gstrSysName
            Exit Sub
        End If
        Do While Not .EOF
            Cob科室.AddItem !科室
            .MoveNext
        Loop
        Cob科室.ListIndex = 0
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function ShowME(ByVal 药房ID_IN As Long, ByVal frmParent As Object, ByVal In_权限 As String) As String
    lng药房ID = 药房ID_IN
    mstrPrivs = In_权限
    
    Me.Show 1, frmParent
    ShowME = strReturn
End Function

Private Sub Form_Unload(Cancel As Integer)
        Call ReleaseSelectorRS
End Sub

Private Sub txtNO_GotFocus()
    Call SelAll(TxtNo)
End Sub

Private Sub TxtNO_Validate(Cancel As Boolean)
    Dim intYear As Integer, strYear As String
    If Trim(TxtNo) = "" Then Exit Sub
    '--如果不满八位,则按规则产生--
    Me.TxtNo = UCase(LTrim(Me.TxtNo))
    If Len(TxtNo) < 8 Then
        intYear = Format(zldatabase.Currentdate, "YYYY") - 1990
        strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
        Me.TxtNo = strYear & String(7 - Len(TxtNo), "0") & Me.TxtNo
    End If
End Sub

Private Sub txt床号_GotFocus()
    Call SelAll(Txt床号)
End Sub

Private Sub Txt姓名_GotFocus()
    Call SelAll(Txt姓名)
End Sub

Private Sub Txt药品_GotFocus()
    Call SelAll(txt药品)
End Sub

Private Sub Txt药品_Validate(Cancel As Boolean)
    txt药品 = Trim(txt药品)
    If txt药品 = "" Then
        txt药品.Tag = 0
        Exit Sub
    End If
    
    Dim RecReturn As New ADODB.Recordset
    Dim sngLeft As Single, sngTop As Single
    
    If InStr(1, txt药品, "[") <> 0 And InStr(1, txt药品, "]") <> 0 Then txt药品.Text = Mid(txt药品.Text, 2, InStr(1, txt药品, "]") - 2)
    sngLeft = Me.Left + txt药品.Left + 50
    sngTop = Me.Top + (Me.Height - Me.ScaleHeight) + txt药品.Top + txt药品.Height - 100
    If DblFrmHeight + sngTop > Screen.Height Then sngTop = sngTop - DblFrmHeight - txt药品.Height + 50
    
'    With Frm药品多选选择器
'        Set RecReturn = .ShowME(Me, 1, lng药房ID, , , Txt药品.Text, sngLeft, sngTop, False)
'        If RecReturn.EOF Then Cancel = True: Exit Sub
'    End With
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(2, "药品部门发药", lng药房ID, lng药房ID)
    End If
    Set RecReturn = frmSelector.ShowME(Me, 1, 2, UCase(txt药品.Text), sngLeft, sngTop, lng药房ID, , , , False, , , , False)
    
    If RecReturn.EOF Then Cancel = True: Exit Sub
    txt药品.Tag = RecReturn!药品ID
    txt药品 = "[" & RecReturn!药品编码 & "]" & IIf(IsNull(RecReturn!通用名), "", RecReturn!通用名)
End Sub
