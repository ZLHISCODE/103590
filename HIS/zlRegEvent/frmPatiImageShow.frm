VERSION 5.00
Begin VB.Form frmPatiImageShow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label lblDefault 
      BackStyle       =   0  'Transparent
      Caption         =   "无照片"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   765
      Left            =   1410
      TabIndex        =   0
      Top             =   1020
      Width           =   2145
   End
   Begin VB.Image imgPatient 
      Height          =   2985
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4515
   End
End
Attribute VB_Name = "frmPatiImageShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng病人ID As Long
Private mblnOk As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    '调整显示位置
    Me.Left = Screen.Width - Me.Width - 500
    Me.Top = 500
    
    Call ReadPatPricture(mlng病人ID)
End Sub

Public Function ShowMe(ByVal lng病人ID As Long) As Boolean
    '功能：入口，显示病人照片
    '参数：
    '   lng病人ID - 显示照片的病人ID
    '返回：
    '   True - 正常显示，False - 显示错误或该病人无照片
    '编制：冉俊明
    '时间：2014-7-7
    mlng病人ID = lng病人ID: mblnOk = False
    Me.Show 1
    ShowMe = mblnOk
End Function

Private Function ReadPatPricture(lng病人ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取病人照片
    '入参:lng病人ID - 病人ID
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strTmp As String
    Dim rsData As Recordset
    
    strSQL = "Select 病人id,照片 From 病人照片 Where 病人id=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    If rsData.BOF = False Then
        strTmp = zlDatabase.ReadPicture(rsData, "照片", strTmp)
        imgPatient.Picture = LoadPicture(strTmp)
        If strTmp <> "" Then Kill strTmp
        lblDefault.Visible = False
        mblnOk = True
    End If
End Function

Private Sub Form_Resize()
    imgPatient.Width = Me.ScaleWidth
    imgPatient.Height = Me.ScaleHeight
End Sub
