VERSION 5.00
Begin VB.Form FrmStuffFilter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "过滤"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5745
   Icon            =   "frmStuffFilter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fraLine 
      Height          =   60
      Left            =   75
      TabIndex        =   7
      Top             =   690
      Width           =   5760
   End
   Begin VB.CheckBox chk规格 
      Alignment       =   1  'Right Justify
      Caption         =   "过滤卫材规格"
      Height          =   210
      Left            =   480
      TabIndex        =   6
      Top             =   2025
      Width           =   1545
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   0
      TabIndex        =   5
      Top             =   2280
      Width           =   5760
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4500
      TabIndex        =   4
      Top             =   2430
      Width           =   1100
   End
   Begin VB.CommandButton cmdFilter 
      Cancel          =   -1  'True
      Caption         =   "过滤(&F)"
      Height          =   350
      Left            =   3240
      TabIndex        =   3
      Top             =   2430
      Width           =   1100
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1845
      TabIndex        =   1
      Top             =   1245
      Width           =   3525
   End
   Begin VB.TextBox txtSim 
      Height          =   300
      Left            =   1845
      TabIndex        =   2
      Top             =   1635
      Width           =   3525
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   1845
      TabIndex        =   0
      Top             =   855
      Width           =   3525
   End
   Begin VB.Label lblComment 
      Caption         =   "    输入希望过滤的卫材编码、名称、别名或者其简码。如存在多条，则返回多条过滤结果。"
      Height          =   435
      Left            =   1020
      TabIndex        =   11
      Top             =   165
      Width           =   4500
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   285
      Picture         =   "frmStuffFilter.frx":000C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "输入卫材名称"
      Height          =   180
      Left            =   495
      TabIndex        =   10
      Top             =   1305
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "输入卫材简码"
      Height          =   180
      Left            =   495
      TabIndex        =   9
      Top             =   1695
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "输入卫材编码"
      Height          =   180
      Left            =   495
      TabIndex        =   8
      Top             =   915
      Width           =   1080
   End
End
Attribute VB_Name = "FrmStuffFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbln显示停用 As Boolean

Public Sub ShowMe(ByVal FrmParent As Object, ByVal bln停用 As Boolean)
    mbln显示停用 = bln停用
    Me.Show , FrmParent
End Sub



Private Sub chk规格_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then OS.PressKey (vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdFilter_Click()
    Dim rs As New ADODB.Recordset
    Dim strResult As String
    Dim n As Long
    Dim strCondition As String
    
    If Len(Trim(Me.txtCode.Text)) > 0 Then
        If Me.chk规格.Value = 0 Then
            strCondition = " AND I.编码 LIKE [1] "
        Else
            strCondition = " AND D.编码 LIKE [1] "
        End If
    End If
    If Len(Trim(Me.txtName.Text)) > 0 Then
        strCondition = " AND N.名称 LIKE [2] "
    End If
    If Len(Trim(Me.txtSim.Text)) > 0 Then
        strCondition = " AND N.简码 LIKE [3] "
    End If
    
    If Len(strCondition) = 0 Then
        MsgBox "请输入卫材信息", vbExclamation, gstrSysName
        On Error Resume Next
        Me.Show
        Me.txtCode.SetFocus: Exit Sub
    End If

    
    If Me.chk规格.Value = 0 Then
        gstrSQL = "SELECT DISTINCT I.ID AS 诊疗ID" & _
                " FROM 诊疗项目目录 I,诊疗项目别名 N" & _
                " WHERE I.ID=N.诊疗项目ID AND I.类别 = '4' " & strCondition

    Else
        gstrSQL = "SELECT DISTINCT I.ID AS 诊疗ID " & _
                 " FROM 诊疗项目目录 I,材料特性 T,收费项目目录 D,收费项目别名 N " & _
                 " WHERE I.ID=T.诊疗ID And T.材料ID=D.ID AND T.材料ID=N.收费细目ID AND I.类别 = '4' " & strCondition
                 
    End If
    If Not mbln显示停用 Then
        gstrSQL = gstrSQL & " And (I.撤档时间 Is NULL Or to_Char(I.撤档时间,'yyyy-MM-dd')='3000-01-01')"
        If Me.chk规格.Value = 1 Then gstrSQL = gstrSQL & " And (D.撤档时间 Is NULL Or to_Char(D.撤档时间,'yyyy-MM-dd')='3000-01-01')"
    End If
    
    gstrSQL = gstrSQL & "order by 诊疗id"
    
    err = 0: On Error GoTo ErrHand
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Trim(Me.txtCode.Text) & "%", IIf(gstrMatchMethod = "0", "%", "") & Trim(Me.txtName.Text) & "%", IIf(gstrMatchMethod = "0", "%", "") & Trim(Me.txtSim.Text) & "%")
    
    With rs
        If .EOF Then
            MsgBox "没有找到卫材信息！", vbExclamation, gstrSysName
            On Error Resume Next
            Me.Show
            Me.txtCode.SetFocus
            Exit Sub
        Else
            For n = 1 To .RecordCount
                If n = 1 Then
                    strResult = Val(!诊疗id)
                Else
                    strResult = strResult & "," & Val(!诊疗id)
                End If
                .MoveNext
            Next
        End If
    End With
    
    Me.Hide
    Call frmStuffMgr.zlGetFilter(strResult)
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub txtAlia_Change()
    
End Sub

Private Sub txtAlia_GotFocus()
    Me.txtCode.SelStart = 0
    Me.txtCode.SelLength = 100
End Sub

Private Sub txtAlia_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtcode_GotFocus()
    Me.txtCode.SelStart = 0
    Me.txtCode.SelLength = 100
End Sub


Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtName_GotFocus()
    Me.txtCode.SelStart = 0
    Me.txtCode.SelLength = 100
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtSim_GotFocus()
    Me.txtCode.SelStart = 0
    Me.txtCode.SelLength = 100
End Sub

Private Sub txtSim_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab): Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub




