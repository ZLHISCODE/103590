VERSION 5.00
Begin VB.Form frmLabSampleCheckFind 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "����"
   ClientHeight    =   2160
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboFind 
      Height          =   300
      ItemData        =   "frmLabSampleCheckFind.frx":0000
      Left            =   1530
      List            =   "frmLabSampleCheckFind.frx":0002
      TabIndex        =   3
      Top             =   705
      Width           =   3800
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4275
      TabIndex        =   2
      Top             =   1650
      Width           =   1100
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "������һ��(&N)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2700
      TabIndex        =   1
      Top             =   1650
      Width           =   1530
   End
   Begin VB.Frame fraLine 
      Height          =   60
      Left            =   -30
      TabIndex        =   0
      Top             =   1485
      Width           =   5925
   End
   Begin VB.Label Label1 
      Caption         =   "��������"
      Height          =   180
      Left            =   675
      TabIndex        =   6
      Top             =   750
      Width           =   915
   End
   Begin VB.Label lblComment 
      Caption         =   "    ����ϣ��������Ŀ�ı��롢���ơ�����ڶ�����������""������һ��""��ֱ���ҵ���ϣ�����ҵ���Ŀ��"
      Height          =   420
      Left            =   1035
      TabIndex        =   5
      Top             =   90
      Width           =   4500
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "(�����ҵ�10������ǰΪ��1��)"
      Height          =   180
      Left            =   615
      TabIndex        =   4
      Top             =   1215
      Width           =   2430
   End
   Begin VB.Image Image1 
      Height          =   765
      Left            =   105
      Picture         =   "frmLabSampleCheckFind.frx":0004
      Top             =   30
      Width           =   840
   End
End
Attribute VB_Name = "frmLabSampleCheckFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsFind As New ADODB.Recordset
Private mrsFindRecord As New ADODB.Recordset
Private strCurSql As String
Private intCount As Integer
Private mstrSql As String                 '���ҵ�SQL
Private mstrFind As String                '�ϴβ�������
Private mway As String                    '������Ҵ���ķ�ʽ
Public Event Finded(ByVal blnFind As Boolean, ByVal strVale As String)

Public Function ShowFind(ByVal strSQL As String)
    '���ܣ� ͨ�ò��Ҵ��壬��������ҵ������Ķ�λ��������λ�������������Finded�¼��д���
    
    If strSQL = "" Then Exit Function
    mstrSql = strSQL
    mway = "SQL"
    Me.Show vbModal
    
End Function

Public Function ShowFindRecordset(ByVal rsTmp As Recordset)
    '���ܣ� ͨ�ò��Ҵ��壬��������ҵ������Ķ�λ��������λ�������������Finded�¼��д���
    Set mrsFindRecord = rsTmp
    mway = "��¼��"
    Me.Show vbModal
    
End Function


Private Sub cboFind_Click()
    If Trim(Me.cboFind.Text) <> "" Then
        Me.cmdFind.Enabled = True
    Else
        Me.cmdFind.Enabled = False
    End If
End Sub

Private Sub cboFind_GotFocus()
    Me.cboFind.SelStart = 0: Me.cboFind.SelLength = 100
End Sub

Private Sub cboFind_KeyPress(KeyAscii As Integer)
'    If InStr(gSysParameter.InvaidWord, Chr(KeyAscii)) > 0 Then KeyAscii = 0
'    If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cboFind_KeyUp(KeyCode As Integer, Shift As Integer)
    If Trim(Me.cboFind.Text) <> "" Then
        Me.cmdFind.Enabled = True
    Else
        Me.cmdFind.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdFind_Click()
    Dim i As Integer
    Dim strFind As String
    Dim strFilter As String
    Dim strReturn As String

    If Trim(Me.cboFind.Text) = "" Then
        MsgBox "��������ҵ�����", vbExclamation
        Me.cboFind.SetFocus: Exit Sub
    End If
    strFind = ""
    For intCount = 0 To Me.cboFind.ListCount
        strFind = strFind & ";" & Me.cboFind.List(intCount)
    Next
    If InStr(1, strFind, ";" & Trim(Me.cboFind.Text)) = 0 Then
        Me.cboFind.AddItem Trim(Me.cboFind.Text), 0
    End If
    
    
    strFind = Trim(Me.cboFind.Text)
   
        
        
    Err = 0: On Error GoTo ErrHand
    Select Case mway
    Case "SQL"
        If mstrFind <> strFind Then
            strCurSql = ""
            mstrFind = strFind
        End If
        strReturn = ""
        With rsFind
            
            If strCurSql <> mstrSql Or .State <> adStateOpen Then
                If .State = adStateOpen Then .Close
                Set rsFind = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, CStr("%" & strFind & "%"))
                If rsFind.EOF Then
                    MsgBox "�����ڲ��ҵ����ݣ�", vbExclamation
                    rsFind.Close: Me.cmdFind.Enabled = False: Me.lblNote.Caption = ""
                    Me.cboFind.SetFocus: Exit Sub
                End If
                strCurSql = mstrSql
            Else
                rsFind.MoveNext
                If rsFind.EOF Then
                    MsgBox "�Ѳ��ҵ����һ����Ŀ��", vbExclamation
                    rsFind.Close: Me.cboFind.Text = "": Me.cmdFind.Enabled = False: Me.lblNote.Caption = ""
                    Me.cboFind.SetFocus: Exit Sub
                End If
            End If
            Me.lblNote.Caption = "(�����ҵ�" & rsFind.RecordCount & "������ǰΪ��" & rsFind.AbsolutePosition & "��)"
            For i = 0 To rsFind.Fields.Count - 1
                strReturn = strReturn & "," & rsFind.Fields(i).Value
            Next
        End With
        If strReturn <> "" Then strReturn = Mid(strReturn, 2)
        RaiseEvent Finded(True, strReturn)
    Case "��¼��"
        If mstrFind <> strFind Then
            mstrFind = strFind
            strReturn = ""
            strFilter = "�������� like '" & strFind & "*' or �������� like '" & strFind & "*'"
            mrsFindRecord.filter = strFilter
        End If
        If mrsFindRecord.RecordCount = 0 Then
            MsgBox "�����ڲ��ҵ����ݣ�", vbExclamation
            Me.cmdFind.Enabled = False: Me.lblNote.Caption = ""
            Me.cboFind.SetFocus: Exit Sub
        ElseIf mrsFindRecord.RecordCount = 1 Then
            If mrsFindRecord.EOF Then
                MsgBox "�Ѳ��ҵ����һ����Ŀ��", vbExclamation
                Me.cboFind.Text = "": Me.cmdFind.Enabled = False: Me.lblNote.Caption = ""
                mstrFind = "": Me.cboFind.SetFocus: Exit Sub
            Else
                For i = 0 To mrsFindRecord.Fields.Count - 1
                    strReturn = strReturn & "," & mrsFindRecord.Fields(i).Value
                Next
                Me.lblNote.Caption = "(�����ҵ�" & mrsFindRecord.RecordCount & "������ǰΪ��" & mrsFindRecord.AbsolutePosition & "��)"
            End If
            mrsFindRecord.MoveNext
        Else
            If mrsFindRecord.EOF Then
                MsgBox "�Ѳ��ҵ����һ����Ŀ��", vbExclamation
                Me.cboFind.Text = "": Me.cmdFind.Enabled = False: Me.lblNote.Caption = ""
                mstrFind = "": Me.cboFind.SetFocus: Exit Sub
            End If
            For i = 0 To mrsFindRecord.Fields.Count - 1
                strReturn = strReturn & "," & mrsFindRecord.Fields(i).Value
            Next
            Me.lblNote.Caption = "(�����ҵ�" & mrsFindRecord.RecordCount & "������ǰΪ��" & mrsFindRecord.AbsolutePosition & "��)"
            mrsFindRecord.MoveNext
        End If
        If strReturn <> "" Then strReturn = Mid(strReturn, 2)
        RaiseEvent Finded(True, strReturn)
    End Select
    Exit Sub
ErrHand:
'    If ComErrCenter() = 1 Then
'    Resume
'    End If
End Sub

Private Sub Form_Activate()
    Me.cboFind.SetFocus
End Sub

Private Sub Form_Load()
    strCurSql = ""
    Me.lblNote.Caption = ""
    mstrFind = ""
End Sub


Private Sub rsRecordset()

End Sub



