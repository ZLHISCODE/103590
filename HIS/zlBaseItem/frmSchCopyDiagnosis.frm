VERSION 5.00
Begin VB.Form frmSchCopyDiagnosis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����ԤԼ�豸����Ŀ"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4095
   ControlBox      =   0   'False
   Icon            =   "frmSchCopyDiagnosis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4095
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   350
      Left            =   2280
      TabIndex        =   3
      Top             =   960
      Width           =   1100
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "ȷ��"
      Height          =   350
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   1100
   End
   Begin VB.ComboBox cboItem 
      Height          =   300
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   420
      Width           =   1335
   End
   Begin VB.Label lblCopy 
      AutoSize        =   -1  'True
      Caption         =   "����                 ��ȫ��������Ŀ"
      Height          =   180
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   3150
   End
End
Attribute VB_Name = "frmSchCopyDiagnosis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngDevice As Long
Private mstrResult As String


Public Function ShowMe(lngDevice As Long, ower As Object) As String
    mlngDevice = lngDevice
    mstrResult = ""
    
    Me.Show 1, ower
    
    ShowMe = mstrResult
    
End Function

Private Sub cmdCancel_Click()
    On Error GoTo errHandle
    
    mstrResult = ""
    Unload Me
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, "��ʾ"
    Err.Clear
End Sub

Private Sub cmdSure_Click()
    On Error GoTo errHandle
    
    If Len(cboItem.Text) = 0 Then
        MsgBox "��ѡ��Ҫ���Ƶ��豸��", vbInformation, "��ʾ"
        Exit Sub
    End If
    mstrResult = cboItem.Text
    Unload Me
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, "��ʾ"
    Err.Clear
End Sub

Private Sub Form_Load()
    Dim strSql As String
    Dim rsTemp As Recordset
    
    On Error GoTo errHandle
    
    strSql = "Select �豸���� From Ӱ��ԤԼ�豸 Where Ӱ����� In (Select Ӱ����� From Ӱ��ԤԼ�豸 Where Id = [1]) AND ID <> [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ѯԤԼ�豸", mlngDevice)
    
    If rsTemp.RecordCount <= 0 Then Exit Sub
    
    Do While Not rsTemp.EOF
        cboItem.AddItem rsTemp!�豸����
        rsTemp.MoveNext
    Loop
    
    If cboItem.ListCount > 0 Then cboItem.ListIndex = 0
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, "��ʾ"
    Err.Clear
End Sub
