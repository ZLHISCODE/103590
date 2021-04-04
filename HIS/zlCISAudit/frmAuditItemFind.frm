VERSION 5.00
Begin VB.Form frmAuditItemFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病案审查项目过滤"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "frmAuditItemFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdSelect 
      Height          =   300
      Left            =   2377
      Picture         =   "frmAuditItemFind.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   180
      Width           =   300
   End
   Begin VB.TextBox txtTypeID 
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   877
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   180
      Width           =   1440
   End
   Begin VB.TextBox txtName 
      DataField       =   "a.名称"
      Height          =   300
      Left            =   877
      MaxLength       =   100
      TabIndex        =   8
      Top             =   990
      Width           =   1440
   End
   Begin VB.TextBox txtCode 
      DataField       =   "a.编码"
      Height          =   300
      Left            =   877
      MaxLength       =   10
      TabIndex        =   4
      Top             =   570
      Width           =   1440
   End
   Begin VB.TextBox txtMnemonicCode 
      DataField       =   "a.简码"
      Height          =   300
      Left            =   3547
      MaxLength       =   100
      TabIndex        =   6
      Top             =   585
      Width           =   1440
   End
   Begin VB.ComboBox cboUsed 
      DataField       =   "a.适用对象"
      Height          =   300
      Left            =   3547
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   960
      Width           =   1440
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2422
      TabIndex        =   11
      Top             =   1695
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3577
      TabIndex        =   12
      Top             =   1695
      Width           =   1100
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   5145
      Y1              =   1485
      Y2              =   1485
   End
   Begin VB.Line Line1 
      X1              =   -15
      X2              =   5145
      Y1              =   1470
      Y2              =   1470
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "简码(&S)"
      Height          =   180
      Index           =   3
      Left            =   2790
      TabIndex        =   5
      Top             =   660
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "编码(&C)"
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   3
      Top             =   630
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "名称(&N)"
      Height          =   195
      Index           =   2
      Left            =   150
      TabIndex        =   7
      Top             =   1065
      Width           =   570
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "分类(&T)"
      Height          =   180
      Index           =   4
      Left            =   150
      TabIndex        =   0
      Top             =   240
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "适用对象(&D)"
      Height          =   180
      Index           =   5
      Left            =   2790
      TabIndex        =   9
      Top             =   1020
      Width           =   990
   End
End
Attribute VB_Name = "frmAuditItemFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnCancel              As Boolean
Public mstrWhere                As String
Private zlCheck                 As New clsCheck

Public Property Get blnCancel() As Boolean
    blnCancel = mblnCancel
End Property

Public Property Let blnCancel(ByVal vNewValue As Boolean)
    mblnCancel = vNewValue
End Property

Public Property Get strWhere() As String
    strWhere = mstrWhere
End Property

Public Property Let strWhere(ByVal vNewValue As String)
    mstrWhere = vNewValue
End Property

'========================================================================================
'=点击 取消
'========================================================================================
Private Sub CmdCancel_Click()
On Error GoTo ErrH
    blnCancel = True
    Unload Me
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'========================================================================================
'=点击 确定
'========================================================================================
Private Sub CmdOK_Click()
On Error GoTo ErrH
    blnCancel = False
    mstrWhere = zlCheck.Frm_GetFilter(Me)
    If txtTypeID.Tag <> "" Then
        If mstrWhere = "" Then
            mstrWhere = "分类ID='" & txtTypeID.Tag & "'"
        Else
            mstrWhere = mstrWhere & vbCrLf & "And 分类ID='" & txtTypeID.Tag & "'"
        End If
    End If
    If mstrWhere = "" Then
        MsgBox "至少得选择一个查询条件！", vbInformation, "中联提示"
        Exit Sub
    End If
    Unload Me
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 选择分类
'==============================================================================
Private Sub cmdSelect_Click()
    Dim intTypeID   As Integer
    Dim intLenght   As Integer
    Dim rsTemp      As ADODB.Recordset
    On Error GoTo ErrH
    
    With frmAuditItemTypeSelect
        .lngLeft = Me.Left + txtTypeID.Left + 10
        .lngTop = Me.Top + txtTypeID.Top + txtTypeID.Height * 2 + 10
        .Show vbModal
        If .blnCancel Then Set frmAuditItemTypeSelect = Nothing: Exit Sub
        intTypeID = .intTypeID
    End With
    gstrSQL = "select /*+ rule */id,上级ID,编码,名称 from 病案审查分类 a Where a.id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(intTypeID))
    If Not zlCheck.Connection_ChkRsState(rsTemp) Then
        txtTypeID.Tag = CStr(intTypeID)
        txtTypeID.Text = "[" + rsTemp!编码 + "]" & rsTemp!名称
    Else
        txtTypeID.Tag = "-1"
        txtTypeID.Text = "[全部]分类"
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
 
'==============================================================================
'=功能： 页面初始化
'==============================================================================
Private Sub Form_Load()
    Dim rsUsed      As New ADODB.Recordset
    On Error GoTo ErrH
    gstrSQL = "select 1 as ID ,'住院医嘱' as Name from dual union all" & vbCrLf & _
                "select 2 as ID ,'住院病历' as Name from dual union all" & vbCrLf & _
                "select 3 as ID ,'护理病历' as Name from dual union all" & vbCrLf & _
                "select 4 as ID ,'护理记录' as Name from dual union all" & vbCrLf & _
                "select 5 as ID ,'首页记录' as Name from dual union all" & vbCrLf & _
                "select 6 as ID ,'医嘱报告' as Name from dual union all" & vbCrLf & _
                "select 7 as ID ,'疾病证明' as Name from dual union all" & vbCrLf & _
                "select 8 as ID ,'知情文件' as Name from dual"
    Set rsUsed = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

    zlCheck.Cmb_List cboUsed, rsUsed, 2
    cboUsed.ListIndex = 0
    zlCheck.Sys_System Me
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
