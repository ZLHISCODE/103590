VERSION 5.00
Begin VB.Form frmChargeSortRate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "统一实收比率"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   Icon            =   "frmChargeSortRate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Caption         =   "比率"
      Height          =   1725
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   3195
      Begin VB.TextBox txtPercentage 
         Height          =   300
         Left            =   1440
         TabIndex        =   2
         Top             =   1110
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "实收比率(&P)"
         Height          =   180
         Left            =   420
         TabIndex        =   7
         Top             =   1170
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   3
         Top             =   1110
         Width           =   150
      End
      Begin VB.Label lbl比率 
         Caption         =   "    所有的收入项目采用统一的实收比率。但前提是要求操作前的各收入项目均未按金额分段。"
         Height          =   585
         Left            =   450
         TabIndex        =   1
         Top             =   330
         Width           =   2595
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   3570
      TabIndex        =   4
      Top             =   270
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   3570
      TabIndex        =   6
      Top             =   1440
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3570
      TabIndex        =   5
      Top             =   690
      Width           =   1100
   End
End
Attribute VB_Name = "frmChargeSortRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnOk As Boolean       '是否成功
Dim mstr费别 As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
     ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ErrHandle
    Dim str比率  As String
    
    str比率 = Trim(txtPercentage.Text)
    If str比率 = "" Then
        MsgBox "实收比率不能为空。", vbExclamation, gstrSysName
        txtPercentage.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(str比率) Then
        MsgBox "实收比率应该是一个数值。", vbExclamation, gstrSysName
        txtPercentage.SetFocus
        zlControl.TxtSelAll txtPercentage
        Exit Sub
    End If
    If Val(str比率) < 0 Or Val(str比率) > 500 Then
        MsgBox "实收比率只能在 0～500之间。", vbExclamation, gstrSysName
        txtPercentage.SetFocus
        zlControl.TxtSelAll txtPercentage
        Exit Sub
    End If
    
    gstrSQL = "zl_费别_Unify('" & mstr费别 & "'," & txtPercentage.Text & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)

    mblnOk = True
    Unload Me
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function UnifyPercentage(ByVal str费别 As String, ByVal lng比率 As Long) As Boolean
'功能:用来与调用的收入项目管理窗口进行通讯的程序
'参数:str费别     要设置的费别
'     lng比率     百分比
'返回值:编辑成功返回True,否则为False
    
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    
    On Error GoTo ErrHandle
    rsTemp.CursorLocation = adUseClient
    gstrSQL = "select B.名称 from 费别明细 A,收入项目 B " & _
           " where a.收入项目ID=B.ID and A.费别=[1] " & _
           " group by B.ID,B.名称  having count(B.ID)>1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str费别)
        
    If rsTemp.RecordCount > 0 Then
        Do Until rsTemp.EOF
            strTemp = strTemp & rsTemp("名称") & ","
            rsTemp.MoveNext
        Loop
        strTemp = "    " & Mid(strTemp, 1, Len(strTemp) - 1)
        MsgBox "以下收入项目已经分段：" & vbCrLf & strTemp & vbCrLf & "本操作不能继续。", vbExclamation, gstrSysName
        Exit Function
    End If
    
    mstr费别 = str费别
    Frame1.Caption = str费别
    txtPercentage.Text = Format(lng比率, "###0.00;-##0.00;0.00;0.00")
    
    mblnOk = False
    frmChargeSortRate.Show vbModal, frmChargeSortGrade
    UnifyPercentage = mblnOk
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txtPercentage_GotFocus()
    zlControl.TxtSelAll txtPercentage
    OS.OpenIme False
End Sub
