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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Label lblDefault 
      BackStyle       =   0  'Transparent
      Caption         =   "����Ƭ"
      BeginProperty Font 
         Name            =   "��Բ"
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
Private mlng����ID As Long
Private mblnOk As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    '������ʾλ��
    Me.Left = Screen.Width - Me.Width - 500
    Me.Top = 500
    
    Call ReadPatPricture(mlng����ID)
End Sub

Public Function ShowMe(ByVal lng����ID As Long) As Boolean
    '���ܣ���ڣ���ʾ������Ƭ
    '������
    '   lng����ID - ��ʾ��Ƭ�Ĳ���ID
    '���أ�
    '   True - ������ʾ��False - ��ʾ�����ò�������Ƭ
    '���ƣ�Ƚ����
    'ʱ�䣺2014-7-7
    mlng����ID = lng����ID: mblnOk = False
    Me.Show 1
    ShowMe = mblnOk
End Function

Private Function ReadPatPricture(lng����ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ƭ
    '���:lng����ID - ����ID
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strTmp As String
    Dim rsData As Recordset
    
    strSQL = "Select ����id,��Ƭ From ������Ƭ Where ����id=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    If rsData.BOF = False Then
        strTmp = zlDatabase.ReadPicture(rsData, "��Ƭ", strTmp)
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
