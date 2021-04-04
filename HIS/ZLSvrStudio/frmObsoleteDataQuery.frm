VERSION 5.00
Object = "{CA73588D-282F-4592-9369-A61CC244FADA}#15.3#0"; "Codejock.SyntaxEdit.v15.3.1.ocx"
Begin VB.Form frmObsoleteDataQuery 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "数据查询SQL"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin XtremeSyntaxEdit.SyntaxEdit steSQL 
      Height          =   4845
      Left            =   60
      TabIndex        =   4
      Top             =   120
      Width           =   8745
      _Version        =   983043
      _ExtentX        =   15425
      _ExtentY        =   8546
      _StockProps     =   84
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   2
      EnableSyntaxColorization=   -1  'True
      ShowLineNumbers =   -1  'True
      ShowSelectionMargin=   -1  'True
      ShowScrollBarVert=   -1  'True
      ShowScrollBarHorz=   -1  'True
      EnableVirtualSpace=   0   'False
      EnableAutoIndent=   -1  'True
      ShowWhiteSpace  =   0   'False
      ShowCollapsibleNodes=   -1  'True
      AutoCompleteWndWidth=   160
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "刷新(&R)"
      Height          =   350
      Left            =   6585
      TabIndex        =   3
      Top             =   5130
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "退出(&C)"
      Height          =   350
      Left            =   7695
      TabIndex        =   1
      Top             =   5130
      Width           =   1100
   End
   Begin VB.Line linNumber 
      X1              =   915
      X2              =   1415
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label lblNumber 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   930
      TabIndex        =   2
      Top             =   5190
      Width           =   495
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "共查询到     条数据。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   75
      TabIndex        =   0
      Top             =   5205
      Width           =   2205
   End
End
Attribute VB_Name = "frmObsoleteDataQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrBusinessName As String

Public Sub ShowMe(ByVal strBusinessName As String)
    mstrBusinessName = strBusinessName
    Me.Show vbModal, frmMDIMain
End Sub

Private Sub cmdRefresh_Click()
    Dim rsTemp As ADODB.Recordset

    On Error GoTo errH
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, steSQL.Text, "执行查询到的SQL")
    lblNumber.Caption = rsTemp.Fields(0).value
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call FillSQL
End Sub

Private Sub FillSQL()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset

    On Error GoTo errH
    strSQL = "Select 数据查询sql From Zltools.Zlobsoletedatadeal Where 名称 = [1]"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "获取数据查询SQL", mstrBusinessName)
    If rsTemp.RecordCount > 0 Then
        steSQL.Text = rsTemp!数据查询sql
        Call cmdRefresh_Click
    End If
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub
