VERSION 5.00
Begin VB.Form frmQrSrv 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame fraQueryRetrieve 
      Height          =   1935
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   11310
      Begin VB.Frame frmPatientID 
         Caption         =   "����IDƥ��"
         Height          =   975
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   11055
         Begin VB.OptionButton optMatch 
            Caption         =   "ҽ��ID"
            Height          =   195
            Index           =   2
            Left            =   8160
            TabIndex        =   5
            ToolTipText     =   "��ҽ��ID�����˺ͽ��յ�Ӱ�����ƥ��"
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton optMatch 
            Caption         =   "���˱�ʶ�ţ�����/סԺ�ţ�"
            Height          =   195
            Index           =   1
            Left            =   3720
            TabIndex        =   4
            ToolTipText     =   "�����˱�ʶ�Ž����˺ͽ��յ�Ӱ�����ƥ��"
            Top             =   480
            Width           =   2610
         End
         Begin VB.OptionButton optMatch 
            Caption         =   "����"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   3
            ToolTipText     =   "�����Ž����˺ͽ��յ�Ӱ�����ƥ��"
            Top             =   480
            Width           =   1065
         End
      End
      Begin VB.CheckBox chkAcceptCGET 
         Caption         =   "֧��C-GET"
         Height          =   255
         Left            =   180
         TabIndex        =   1
         Top             =   330
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmQrSrv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngSrvID As Long
Public Sub ShowRefresh(ByVal SrvID As Long)
    mlngSrvID = SrvID
    If mlngSrvID = 0 Then
        fraQueryRetrieve.Caption = "�Ϸ��б�����ѡ������δ���棬���ܽ������ã�"
        fraQueryRetrieve.Enabled = False
    Else
        fraQueryRetrieve.Caption = ""
        fraQueryRetrieve.Enabled = True
    End If
    RefreshPara
End Sub

Public Sub SavePara()
    Dim strData As String
    Dim i As Integer
    
    On Error GoTo ErrHandle
    gstrSQL = "Zl_Ӱ��DICOM�������_SAVE(" & mlngSrvID & ",'֧��C-GET','" & chkAcceptCGET.value & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����֧��C-GET")
    
    strData = 0
    For i = 0 To optMatch.UBound
        If optMatch(i).value = True Then
            strData = i
            Exit For
        End If
    Next
    If strData = "" Then strData = "0"
    gstrSQL = "Zl_Ӱ��DICOM�������_SAVE(" & mlngSrvID & ",'����IDƥ��','" & strData & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���没��IDƥ��")
   Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub RefreshPara()
Dim rsTemp As New ADODB.Recordset, i As Integer
        gstrSQL = "select ����ID,�������� ,����ֵ from Ӱ��DICOM������� where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����", mlngSrvID)
    chkAcceptCGET.value = False
    Do Until rsTemp.EOF
        Select Case rsTemp!��������
            Case "֧��C-GET"
                chkAcceptCGET.value = Nvl(rsTemp!����ֵ)
            Case "����IDƥ��"
                optMatch(Nvl(rsTemp!����ֵ, 0)) = True
        End Select
        rsTemp.MoveNext
    Loop
End Sub

