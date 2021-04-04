VERSION 5.00
Begin VB.Form frmLabRefuse 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ձ걾"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4470
   Icon            =   "frmLabRefuse.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboRefuse 
      Height          =   300
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   780
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   30
      TabIndex        =   4
      Top             =   2190
      Width           =   4485
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3240
      TabIndex        =   3
      Top             =   2400
      Width           =   1100
   End
   Begin VB.CommandButton cmdRefuse 
      Caption         =   "����(&F)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1950
      TabIndex        =   2
      Top             =   2400
      Width           =   1100
   End
   Begin VB.TextBox TxtRefuse 
      Height          =   975
      Left            =   960
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1110
      Width           =   3375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "����д�������ɺ󵥻����հ�ť"
      Height          =   180
      Left            =   960
      TabIndex        =   5
      Top             =   510
      Width           =   2520
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   240
      Picture         =   "frmLabRefuse.frx":000C
      Top             =   210
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�����ȷ��Ҫ���ոñ걾��?"
      Height          =   180
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   2250
   End
End
Attribute VB_Name = "frmLabRefuse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mlngKey As Long                                         '����걾ID
Dim blnComm As Boolean                                      '�Ƿ�����˫��ͨ��
Dim mfrmMain As Form
Dim objLISComm As Object
Dim mblnOK As Boolean
Dim mWinsockC As Winsock

Private Sub cboRefuse_Click()
    Me.TxtRefuse.Text = Mid(Me.cboRefuse.Text, InStr(Me.cboRefuse.Text, "-") + 1)
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRefuse_Click()
    Dim blnTran As Boolean
    Dim strSQL As String
    Dim rs As New ADODB.Recordset, strQrySQL As String
    Dim strDevices As String, aDevice() As String, strAdviceIDs As String, i As Integer
    Dim intType As Integer                  '�걾���:0=��ͨ��1=����
    Dim lngAdviceID As Long                 'ҽ��ID
    Dim intEmerge As Integer                '�Ƿ����ҽ��

    intEmerge = Val(zlDatabase.GetPara("����걾", 100, 1208, 0))
    
    If mlngKey = 0 Then Exit Sub
    If Trim(Me.TxtRefuse.Text) = "" Then
        MsgBox "����д��������!лл!", vbInformation, gstrSysName
        Me.TxtRefuse.SetFocus
        Exit Sub
    End If
    
    On Error GoTo ErrHand

    Me.MousePointer = vbHourglass
    strAdviceIDs = "": strDevices = ""

    strSQL = "select distinct nvl(b.�걾���,0) as �걾���,a.id as ҽ��Id " & _
             " from ����ҽ����¼ a,����걾��¼ b " & _
             " where a.id = b.ҽ��ID and a.id = [1] "
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngKey)

    If rs.BOF = False Then
        intType = rs("�걾���")
        lngAdviceID = rs("ҽ��ID")
    End If

    '����˫��ͨ��
    If blnComm Then
        strAdviceIDs = strAdviceIDs & "," & lngAdviceID

        strQrySQL = "Select Distinct ����ID From ����걾��¼ A,������Ŀ�ֲ� B" & _
            " Where B.ҽ��ID=[1] And B.�걾ID+0=A.ID"
        Set rs = zlDatabase.OpenSQLRecord(strQrySQL, Me.Caption, lngAdviceID)
        Do While Not rs.EOF
            If InStr(strDevices, "," & zlCommFun.Nvl(rs(0), 0)) = 0 Then
                strDevices = strDevices & "," & zlCommFun.Nvl(rs(0), 0)
            End If

            rs.MoveNext
        Loop
    End If

    '����˫��ͨ��
    If blnComm Then
        If Len(strDevices) > 0 Then strDevices = Mid(strDevices, 2)
        If Len(strAdviceIDs) > 0 Then strAdviceIDs = Mid(strAdviceIDs, 2)
        aDevice = Split(strDevices, ",")
        For i = 0 To UBound(aDevice)
            SendSample mWinsockC, mWinsockC.LocalIP, CLng(Val(aDevice(i))), "", 0, strAdviceIDs, True, IIf(intEmerge = 1 And intType = 1, 1, 0)
        Next
    End If
    Me.MousePointer = vbDefault
    
    blnTran = True
    gcnOracle.BeginTrans
        strSQL = "ZL_����걾��¼_ȡ������(" & lngAdviceID & ")"
        zlDatabase.ExecuteProcedure strSQL, gstrSysName
        strSQL = "Zl_����걾��¼_�걾����(" & mlngKey & ",'" & Me.TxtRefuse.Text & "','" & UserInfo.���� & "')"
        zlDatabase.ExecuteProcedure strSQL, gstrSysName
    gcnOracle.CommitTrans
    blnTran = False

'    SaveData = True
    Unload Me
    
    Exit Sub
    
ErrHand:
    
    Me.MousePointer = vbDefault
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
        

End Sub

Public Function ShowEdit(ByVal frmMain As Object, ByRef lngKey As Long, WinsockC As Winsock) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʾ�༭���壬������ô���Ľӿں���
    '����:  frmMain         ���ô������
    '       lngKey          �걾id
    '����:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    mlngKey = lngKey
    blnComm = Val(zlDatabase.GetPara("��������˫��", 100, 1208, 0))

    Set mfrmMain = frmMain

    
    If mlngKey = 0 Then Exit Function
    
    Set mWinsockC = WinsockC
    Me.Show 1, frmMain

    ShowEdit = mblnOK

End Function


Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    gstrSql = "select ����,���� from �����������"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Do While Not rsTmp.EOF
        With Me.cboRefuse
            .AddItem Nvl(rsTmp("����")) & "-" & Nvl(rsTmp("����"))
        End With
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
