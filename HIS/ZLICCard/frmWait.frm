VERSION 5.00
Begin VB.Form frmWait 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ȴ�ǰ�û���Ӧ"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4740
   Icon            =   "frmWait.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2850
      Top             =   660
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3330
      TabIndex        =   1
      Top             =   720
      Width           =   1100
   End
   Begin VB.Label lblnote 
      AutoSize        =   -1  'True
      Caption         =   "������ǰ�û��ύ�������Ժ�......"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   420
      TabIndex        =   0
      Top             =   300
      Width           =   3360
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintState As Integer    '0-δ����;1-���ڴ���;2-�û�����;3-ʧ��;9-�ɹ�
Private mstr���� As String
Private mlng���к� As Long
Private mstr������Ϣ As String

'str������Ϣ������ʱ��¼������Ϣ�������ɹ�ʱ��¼���յ�������
Public Function SendRequest(ByVal str���� As String, ByVal lng���к� As Long, ByRef str������Ϣ As String) As Boolean
    mintState = 0
    mstr���� = str����
    mlng���к� = lng���к�
    mstr������Ϣ = ""
    Me.Show 1
    str������Ϣ = mstr������Ϣ
    SendRequest = (mintState = 9)
End Function

Private Sub cmdȡ��_Click()
    'ȡ����ֱ���˳������ˣ����������ݱ��е�״̬
    mintState = 3
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    '��������Ƿ��ѷ���
    
    strSQL = " Select Nvl(��־,0) AS ��־,����" & _
              " From ��Ϣ����" & _
              " Where ����='" & mstr���� & "' And ���к�=" & mlng���к�
    Call OpenRecordset(rsTemp, "��������Ƿ��ѷ���", strSQL)
    mintState = rsTemp!��־
    mstr������Ϣ = IIf(IsNull(rsTemp!����), "", rsTemp!����)
    
    Select Case mintState
    Case 0
        Exit Sub
    Case 1
        Me.lblnote.Caption = "����������ת���������Ժ�......"
    Case 2
        Me.lblnote.Caption = "�û�������"
        Timer1.Enabled = False
    Case 3
        Me.lblnote.Caption = "�������������˳���"
        Timer1.Enabled = False
    Case 9
        Timer1.Enabled = False
        Call OrganizeData
    End Select
    
    If Timer1.Enabled = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub OpenRecordset(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSQL As String = "")
'���ܣ��򿪼�¼��
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.CursorLocation = adUseClient
    
    rsTemp.Open strSQL, gcnConnect, adOpenStatic, adLockReadOnly
    Set rsTemp.ActiveConnection = Nothing
End Sub

Private Sub OrganizeData()
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    strSQL = " Select �������� From ��Ϣ���� Where ����='" & mstr���� & "' And ���к�=" & mlng���к� & " Order by �к�"
    Call OpenRecordset(rsTemp, "��ȡ��������", strSQL)
    If rsTemp.RecordCount = 0 Then Exit Sub
    
    With rsTemp
        mstr������Ϣ = ""
        Do While Not .EOF
            mstr������Ϣ = mstr������Ϣ & IIf(IsNull(!��������), "", !��������)
            .MoveNext
        Loop
    End With
End Sub
