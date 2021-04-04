VERSION 5.00
Begin VB.Form frmTableComment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ݱ����-��ע�༭"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5235
   Icon            =   "frmTableComment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5235
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdSaveContinue 
      Caption         =   "���沢����"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdSaveExit 
      Caption         =   "������˳�"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtComment 
      Height          =   1575
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   3855
   End
   Begin VB.TextBox txtTable 
      BackColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      Caption         =   "��ע"
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   360
   End
   Begin VB.Label lblTable 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   360
   End
End
Attribute VB_Name = "frmTableComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrAllTables As String '������ʽ  ",a,b,c,d,"
Private mblnFirst As Boolean
Private blnChange As Boolean

Public Function ShowMe(ByVal strAllTables As String, strTable As String) As Boolean
        
    mstrAllTables = strAllTables
    txtTable.Text = strTable
    txtComment.Text = GetAllComment(strTable)
    If mblnFirst = False Then
        blnChange = False
        mblnFirst = True
        Me.Show 1
    End If
    ShowMe = blnChange
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSaveContinue_Click()
    Dim strTable As String, strTemp As String
    Dim lngBegin As Long

    If SaveComment = False Then Exit Sub
    strTable = txtTable.Text
    lngBegin = InStr(mstrAllTables, "," & strTable & ",") + Len(strTable) + 2
    strTemp = Mid(mstrAllTables, lngBegin)
    If strTemp = "" Then
        MsgBox "�ѵ������һ�����޷����������˳������棡", vbExclamation, Me.Caption
        Unload Me
        Exit Sub
    Else
        strTemp = Mid(strTemp, 1, InStr(strTemp, ",") - 1)
        Call ShowMe(mstrAllTables, strTemp)
    End If
End Sub

Private Sub cmdSaveExit_Click()

    If SaveComment = False Then Exit Sub
    Unload Me
End Sub

Private Function SaveComment() As Boolean
'���浱ǰ��ı�ע
    Dim strComment As String
    
    On Error GoTo errH
    strComment = txtComment.Text
    If InStr(strComment, "'") > 0 Then
        MsgBox "��ע��Ϣ�в��ܰ��������ţ���ɾ�������ź������±��棡", vbExclamation, Me.Caption
        Exit Function
    End If
    gstrSQL = "Comment On Table " & txtTable.Text & " Is '" & strComment & "'"
    gcnOracle.Execute gstrSQL
    blnChange = True
    SaveComment = True
    Exit Function
errH:
    MsgBox err.Description, vbCritical, Me.Caption
    err.Clear
End Function

Private Sub Form_Activate()
    txtComment.SetFocus
    txtComment.SelStart = 0
    txtComment.SelLength = Len(txtComment.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnFirst = False
End Sub
