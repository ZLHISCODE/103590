VERSION 5.00
Begin VB.Form frmClinicSetNewPlanDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "确定出诊表时间"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4275
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClinicSetNewPlanDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   360
      Left            =   2520
      TabIndex        =   8
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   360
      Left            =   1470
      TabIndex        =   7
      Top             =   1440
      Width           =   975
   End
   Begin VB.ComboBox cboWeek 
      Height          =   315
      Left            =   2940
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   660
      Width           =   765
   End
   Begin VB.ComboBox cboMonth 
      Height          =   315
      Left            =   1740
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   660
      Width           =   855
   End
   Begin VB.ComboBox cboYear 
      Height          =   315
      Left            =   450
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   660
      Width           =   945
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请确定出诊表的时间："
      Height          =   195
      Left            =   270
      TabIndex        =   0
      Top             =   180
      Width           =   1800
   End
   Begin VB.Label lblWeek 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "周"
      Height          =   195
      Left            =   3720
      TabIndex        =   6
      Top             =   720
      Width           =   180
   End
   Begin VB.Label lblMonth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "月"
      Height          =   195
      Left            =   2610
      TabIndex        =   4
      Top             =   720
      Width           =   180
   End
   Begin VB.Label lblYear 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "年"
      Height          =   195
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   180
   End
End
Attribute VB_Name = "frmClinicSetNewPlanDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytPlanType As Byte '1-月安排,2-周安排
Private mdtCur As Date
Private mblnOK As Boolean
Private mblnNOtClick As Boolean
Private mintYear As Integer, mintMonth As Integer, mintWeek As Integer

Public Function ShowMe(frmParent As Object, ByVal bytPlanType As Byte, ByVal dtCur As Date, _
    ByRef intYear As Integer, ByRef intMonth As Integer, Optional ByRef intWeek As Integer) As Boolean
    '程序入口，确定出诊表的时间
    mbytPlanType = bytPlanType: mdtCur = dtCur
    
    On Error Resume Next
    mblnOK = False
    Me.Show 1, frmParent
    ShowMe = mblnOK
    If mblnOK Then
        intYear = mintYear
        intMonth = mintMonth
        If bytPlanType = 2 Then intWeek = mintWeek
    End If
End Function

Private Sub cboMonth_Click()
    Dim intYear As Integer, intMonth As Integer
    Dim i As Integer, intWeekCount As Integer
    
    If mblnNOtClick Then Exit Sub
    Err = 0: On Error GoTo errHandler
    intYear = Val(cboYear.Text): intMonth = Val(cboMonth.Text)
    intWeekCount = GetWeekCount(intYear, intMonth)
    cboWeek.Clear
    For i = 1 To intWeekCount
        cboWeek.AddItem CStr(i)
    Next
    If cboWeek.ListCount > 0 Then cboWeek.ListIndex = 0
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboMonth_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboWeek_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboYear_Click()
    Dim intYear As Integer, intMonth As Integer
    Dim i As Integer, intWeekCount As Integer
    
    If mblnNOtClick Then Exit Sub
    Err = 0: On Error GoTo errHandler
    intYear = Val(cboYear.Text): intMonth = Val(cboMonth.Text)
    intWeekCount = GetWeekCount(intYear, intMonth)
    cboWeek.Clear
    For i = 1 To intWeekCount
        cboWeek.AddItem CStr(i)
    Next
    If cboWeek.ListCount > 0 Then cboWeek.ListIndex = 0
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboYear_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim intYear As Integer, intMonth As Integer, intWeek As Integer
    Dim i As Integer, blnIsOk As Boolean
    
    Err = 0: On Error GoTo errHandler
    mintYear = Val(cboYear.Text)
    mintMonth = Val(cboMonth.Text)
    mintWeek = Val(cboWeek.Text)
    '检查
    intYear = Year(mdtCur): intMonth = Month(mdtCur): intWeek = GetDateWeek(mdtCur)
    blnIsOk = True
    If mintYear < intYear Then
        blnIsOk = False
    ElseIf mintYear = intYear Then
        If mintMonth < intMonth Then
            blnIsOk = False
        ElseIf mintMonth = intMonth Then
            If mintWeek < intWeek And mbytPlanType = 2 Then
                blnIsOk = False
            End If
        End If
    End If
    If blnIsOk = False Then
        MsgBox "出诊表时间不能小于当前时间(" & intYear & "年" & intMonth & "月" & _
            IIf(mbytPlanType = 2, "第" & intWeek & "周", "") & ")！", vbInformation, gstrSysName
        If mbytPlanType = 1 Then
            If cboMonth.Visible And cboMonth.Enabled Then cboMonth.SetFocus
        Else
            If cboWeek.Visible And cboWeek.Enabled Then cboWeek.SetFocus
        End If
        Exit Sub
    End If
    
    mblnOK = True
    Unload Me
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim intYear As Integer, intMonth As Integer, intWeek As Integer
    Dim i As Integer, intWeekCount As Integer
    
    Err = 0: On Error GoTo errHandler
    intYear = Year(mdtCur): intMonth = Month(mdtCur): intWeek = GetDateWeek(mdtCur)
    '1.年，缺省加载五年
    cboYear.Clear
    For i = 0 To 4
        cboYear.AddItem CStr(intYear + i)
    Next
    mblnNOtClick = True
    If cboYear.ListCount > 0 Then cboYear.ListIndex = 0
    mblnNOtClick = False
    
    '2.月
    cboMonth.Clear
    For i = 1 To 12
        cboMonth.AddItem CStr(i)
        mblnNOtClick = True
        If i = intMonth Then cboMonth.ListIndex = cboMonth.NewIndex
        mblnNOtClick = False
    Next
    
    '3,周
    intWeekCount = GetWeekCount(intYear, intMonth)
    cboWeek.Clear
    For i = 1 To intWeekCount
        cboWeek.AddItem CStr(i)
        If i = intWeek Then cboWeek.ListIndex = cboWeek.NewIndex
    Next
    
    cboWeek.Visible = mbytPlanType = 2
    lblWeek.Visible = mbytPlanType = 2
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

