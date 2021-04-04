VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPatiFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "已诊病人过滤"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   Icon            =   "frmPatiFilter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame3 
      Height          =   120
      Left            =   -75
      TabIndex        =   20
      Top             =   570
      Width           =   5085
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   -45
      TabIndex        =   19
      Top             =   1530
      Width           =   5085
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -75
      TabIndex        =   18
      Top             =   2550
      Width           =   5130
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2940
      TabIndex        =   9
      Top             =   2805
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1845
      TabIndex        =   8
      Top             =   2805
      Width           =   1100
   End
   Begin VB.TextBox txt姓名 
      Height          =   300
      Left            =   2955
      TabIndex        =   7
      Top             =   2190
      Width           =   1260
   End
   Begin VB.TextBox txt就诊卡 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   960
      TabIndex        =   6
      Top             =   2190
      Width           =   1260
   End
   Begin VB.TextBox txt门诊号 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2955
      TabIndex        =   5
      Top             =   1785
      Width           =   1260
   End
   Begin VB.TextBox txt挂号单 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   960
      TabIndex        =   4
      Top             =   1785
      Width           =   1260
   End
   Begin VB.ComboBox cbo医生 
      Height          =   300
      Left            =   960
      TabIndex        =   3
      Top             =   1215
      Width           =   3255
   End
   Begin VB.ComboBox cbo科室 
      Height          =   300
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   3255
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   300
      Left            =   960
      TabIndex        =   0
      Top             =   210
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   102039555
      CurrentDate     =   38004
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   300
      Left            =   2955
      TabIndex        =   1
      Top             =   210
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   102039555
      CurrentDate     =   38004
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "姓名"
      Height          =   180
      Left            =   2535
      TabIndex        =   17
      Top             =   2250
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "就诊卡"
      Height          =   180
      Left            =   360
      TabIndex        =   16
      Top             =   2250
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "门诊号"
      Height          =   180
      Left            =   2355
      TabIndex        =   15
      Top             =   1845
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "挂号单"
      Height          =   180
      Left            =   360
      TabIndex        =   14
      Top             =   1845
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "就诊医生"
      Height          =   180
      Left            =   180
      TabIndex        =   13
      Top             =   1275
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "就诊科室"
      Height          =   180
      Left            =   180
      TabIndex        =   12
      Top             =   900
      Width           =   720
   End
   Begin VB.Label lbl开始时间 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "就诊时间"
      Height          =   180
      Left            =   180
      TabIndex        =   11
      Top             =   255
      Width           =   720
   End
   Begin VB.Label lbl结束时间 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "至"
      Height          =   180
      Left            =   2490
      TabIndex        =   10
      Top             =   270
      Width           =   180
   End
End
Attribute VB_Name = "frmPatiFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mstrLike As String
Private mblnOK As Boolean
Private mlngPreDept As Long
Private mdBegin As Date, mdEnd As Date
Private mlng科室ID As Long, mstr医生 As String
Private mstr挂号单 As String, mstr门诊号 As String, mstr就诊卡 As String, mstr姓名 As String

Public Function ShowMe(frmParent As Object, dBegin As Date, dEnd As Date, _
     lng科室ID As Long, str医生 As String, str挂号单 As String, _
     str门诊号 As String, str就诊卡 As String, str姓名 As String, strPrivs As String) As Boolean
    
    mdBegin = dBegin
    mdEnd = dEnd
    mlng科室ID = lng科室ID
    mstr医生 = str医生
    mstr挂号单 = str挂号单
    mstr门诊号 = str门诊号
    mstr就诊卡 = str就诊卡
    mstr姓名 = str姓名
    mstrPrivs = strPrivs
    
    Me.Show 1, frmParent
    
    If mblnOK Then
        dBegin = mdBegin
        dEnd = mdEnd
        lng科室ID = mlng科室ID
        str医生 = mstr医生
        str挂号单 = mstr挂号单
        str门诊号 = mstr门诊号
        str就诊卡 = mstr就诊卡
        str姓名 = mstr姓名
    End If
    ShowMe = mblnOK
End Function

Private Sub cbo科室_Click()
    If cbo科室.ListIndex <> -1 Then
        If mlngPreDept <> cbo科室.ItemData(cbo科室.ListIndex) Then
            mlngPreDept = cbo科室.ItemData(cbo科室.ListIndex)
            Call ReadDoctor(mlngPreDept)
        End If
    ElseIf mlngPreDept <> 0 Then
        mlngPreDept = 0
        Call ReadDoctor
    End If
End Sub

Private Sub cbo科室_GotFocus()
    Call zlControl.TxtSelAll(cbo科室)
End Sub

Private Sub cbo科室_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call ZLCommFun.PressKey(vbKeyTab) '非SetFocus方式会激活Validate事件,除非下一个是vsFlexGrid控件。
    Else
        If InStr(mstrPrivs, "所有操作员") = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub cbo科室_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnDo As Boolean
    
    On Error GoTo errH
        
    If cbo科室.Text <> "" Then
        blnDo = True
        If cbo科室.ListIndex <> -1 Then
            If cbo科室.List(cbo科室.ListIndex) = cbo科室.Text Then blnDo = False
        End If
        If blnDo Then
            strSql = "Select B.ID,B.编码,B.名称" & _
                " From 部门表 B,部门性质说明 C" & _
                " Where B.ID=C.部门ID And C.服务对象 In(1,3) And C.工作性质='临床'" & _
                " And (B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or B.撤档时间 is Null)" & _
                " And (B.编码 Like [1] Or Upper(B.简码) Like [2] Or Upper(B.名称) Like [2])" & _
                " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & _
                " Order by B.编码"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UCase(cbo科室.Text) & "%", mstrLike & UCase(cbo科室.Text) & "%")
            If Not rsTmp.EOF Then
                Call Cbo.SeekIndex(cbo科室, rsTmp!ID)
            ElseIf mlngPreDept <> 0 Then
                Call Cbo.SeekIndex(cbo科室, mlngPreDept)
            Else
                cbo科室.Text = ""
                Call cbo科室_Click
            End If
        End If
    Else
        Call cbo科室_Click
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo医生_GotFocus()
    Call zlControl.TxtSelAll(cbo医生)
End Sub

Private Sub cbo医生_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnDo As Boolean
    Dim lng科室ID As Long
    
    On Error GoTo errH
        
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        If cbo医生.Text <> "" Then
            blnDo = True
            If cbo医生.ListIndex <> -1 Then
                If cbo医生.List(cbo医生.ListIndex) = cbo医生.Text Then blnDo = False
            End If
            
            If blnDo Then
                If cbo科室.ListIndex <> -1 Then
                    lng科室ID = cbo科室.ItemData(cbo科室.ListIndex)
                End If
                If lng科室ID <> 0 Then
                    strSql = "Select Distinct A.简码,A.姓名" & _
                        " From 人员表 A,人员性质说明 B,部门人员 C" & _
                        " Where A.ID=B.人员ID And A.ID=C.人员ID" & _
                        " And B.人员性质='医生' And C.部门ID=[1]" & _
                        " And (A.编号 Like [2] Or Upper(A.简码) Like [3] Or Upper(A.姓名) Like [3])" & _
                        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
                        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                        " Order by A.简码"
                Else
                    strSql = "Select Distinct A.简码,A.姓名" & _
                        " From 人员表 A,人员性质说明 B" & _
                        " Where A.ID=B.人员ID And B.人员性质='医生'" & _
                        " And (A.编号 Like [2] Or Upper(A.简码) Like [3] Or Upper(A.姓名) Like [3])" & _
                        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
                        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                        " Order by A.简码"
                End If
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng科室ID, UCase(cbo医生.Text) & "%", mstrLike & UCase(cbo医生.Text) & "%")
                If Not rsTmp.EOF Then
                    Call Cbo.SeekIndex(cbo医生, rsTmp!姓名)
                Else
                    cbo医生.Text = ""
                End If
            End If
        End If
        Call ZLCommFun.PressKey(vbKeyTab)
    Else
        If InStr(mstrPrivs, "所有操作员") = 0 Then KeyAscii = 0
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim curDate As Date
    
    If cbo科室.ListIndex <> -1 Then
        mlng科室ID = cbo科室.ItemData(cbo科室.ListIndex)
    Else
        mlng科室ID = 0
    End If
    If cbo医生.ListIndex <> -1 Then
        mstr医生 = ZLCommFun.GetNeedName(cbo医生.Text)
    Else
        mstr医生 = ""
    End If
    If txt挂号单.Text <> "" Then
        mstr挂号单 = txt挂号单.Text
    Else
        mstr挂号单 = ""
    End If
    If txt门诊号.Text <> "" Then
        mstr门诊号 = txt门诊号.Text
    Else
        mstr门诊号 = ""
    End If
    If txt就诊卡.Text <> "" Then
        mstr就诊卡 = txt就诊卡.Text
    Else
        mstr就诊卡 = ""
    End If
    If txt姓名.Text <> "" Then
        mstr姓名 = txt姓名.Text
    Else
        mstr姓名 = ""
    End If
    
    mdBegin = dtpBegin.Value
    mdEnd = dtpEnd.Value
    
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 And Not (Me.ActiveControl Is cbo科室 Or Me.ActiveControl Is cbo医生) Then
        KeyAscii = 0
        Call ZLCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    mblnOK = False
    mlngPreDept = -1
    mstrLike = IIf(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "") '输入匹配方式
    
    Call Cbo.SeekIndex(cbo科室, mlng科室ID)
    Call Cbo.SeekIndex(cbo医生, mstr医生)
    dtpBegin.Value = mdBegin
    dtpEnd.Value = mdEnd
    txt挂号单.Text = mstr挂号单
    txt门诊号.Text = mstr门诊号
    txt就诊卡.Text = mstr就诊卡
    txt就诊卡.PasswordChar = IIf(gblnCardHide, "*", "")
    txt姓名.Text = mstr姓名
    
    On Error GoTo errH
    
    '读取门诊科室:缺省为无科室
    If InStr(mstrPrivs, "所有操作员") > 0 Then
        strSql = "Select Distinct B.ID,B.编码,B.名称" & _
            " From 部门表 B,部门性质说明 C" & _
            " Where B.ID=C.部门ID And C.服务对象 In(1,3) And C.工作性质='临床'" & _
            " And (B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or B.撤档时间 is Null)" & _
            " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & _
            " Order by B.编码"
    Else
        strSql = "Select Distinct B.ID,B.编码,B.名称" & _
            " From 部门表 B,部门性质说明 C,部门人员 D" & _
            " Where B.ID=C.部门ID And B.ID=D.部门ID And D.人员ID=[1]" & _
            " And C.服务对象 In(1,3) And C.工作性质='临床'" & _
            " And (B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or B.撤档时间 is Null)" & _
            " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & _
            " Order by B.编码"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    Do While Not rsTmp.EOF
        cbo科室.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cbo科室.ItemData(cbo科室.NewIndex) = rsTmp!ID
        If rsTmp!ID = mlng科室ID Then
            Call Cbo.SetIndex(cbo科室.hwnd, cbo科室.NewIndex)
        End If
        rsTmp.MoveNext
    Loop
        
    '读取门诊医生
    Call cbo科室_Click
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ReadDoctor(Optional ByVal lng科室ID As Long)
'功能：读取指定门诊科室的医生
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    cbo医生.Clear
    
    On Error GoTo errH
    
    If InStr(mstrPrivs, "所有操作员") = 0 Then
        strSql = "Select 简码,姓名 From 人员表 Where ID=[2]"
        cbo医生.Enabled = False
    ElseIf lng科室ID <> 0 Then
        strSql = "Select Distinct A.简码,A.姓名" & _
            " From 人员表 A,人员性质说明 B,部门人员 C" & _
            " Where A.ID=B.人员ID And A.ID=C.人员ID" & _
            " And B.人员性质='医生' And C.部门ID=[1]" & _
            " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order by A.简码"
    Else
        strSql = "Select Distinct A.简码,A.姓名" & _
            " From 人员表 A,人员性质说明 B" & _
            " Where A.ID=B.人员ID And B.人员性质='医生'" & _
            " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order by A.简码"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng科室ID, UserInfo.ID)
    Do While Not rsTmp.EOF
        cbo医生.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        If rsTmp!姓名 = mstr医生 Then
            cbo医生.ListIndex = cbo医生.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt挂号单_Change()
    If txt挂号单.Text <> "" Then
        txt就诊卡.Text = ""
        txt姓名.Text = ""
        txt门诊号.Text = ""
    End If
End Sub

Private Sub txt挂号单_GotFocus()
    Call zlControl.TxtSelAll(txt挂号单)
End Sub

Private Sub txt挂号单_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt挂号单_Validate(Cancel As Boolean)
    If IsNumeric(txt挂号单.Text) Then
        txt挂号单.Text = GetFullNO(txt挂号单.Text, 12)
    End If
End Sub

Private Sub txt就诊卡_Change()
    If txt就诊卡.Text <> "" Then
        txt门诊号.Text = ""
        txt姓名.Text = ""
        txt挂号单.Text = ""
    End If
End Sub

Private Sub txt就诊卡_GotFocus()
    Call zlControl.TxtSelAll(txt就诊卡)
End Sub

Private Sub txt就诊卡_KeyPress(KeyAscii As Integer)
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt门诊号_Change()
    If txt门诊号.Text <> "" Then
        txt就诊卡.Text = ""
        txt姓名.Text = ""
        txt挂号单.Text = ""
    End If
End Sub

Private Sub txt门诊号_GotFocus()
    Call zlControl.TxtSelAll(txt门诊号)
End Sub

Private Sub txt门诊号_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt姓名_Change()
    If txt姓名.Text <> "" Then
        txt门诊号.Text = ""
        txt就诊卡.Text = ""
        txt挂号单.Text = ""
    End If
End Sub

Private Sub txt姓名_GotFocus()
    Call zlControl.TxtSelAll(txt姓名)
End Sub
