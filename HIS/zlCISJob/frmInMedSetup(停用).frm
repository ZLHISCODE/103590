VERSION 5.00
Begin VB.Form frmInMedSetup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "������Ŀ����"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4200
      TabIndex        =   10
      Top             =   4320
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2880
      TabIndex        =   6
      Top             =   4320
      Width           =   1100
   End
   Begin VB.TextBox txtEdit 
      Height          =   1335
      Index           =   2
      Left            =   960
      MaxLength       =   1000
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2880
      Width           =   4455
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   1
      Left            =   960
      MaxLength       =   20
      TabIndex        =   3
      Top             =   2460
      Width           =   4455
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   0
      Left            =   960
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1980
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   0
      TabIndex        =   9
      Top             =   1800
      Width           =   5535
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   120
      Picture         =   "frmInMedSetup.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����(&R)"
      Height          =   180
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����(&N)"
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����(&B)"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   630
   End
   Begin VB.Label lbl���� 
      Caption         =   "���ݰ�����"
      Height          =   1260
      Left            =   720
      TabIndex        =   8
      Top             =   480
      Width           =   4740
   End
   Begin VB.Label lbl��Ϣ 
      Caption         =   "�û��Զ���Ĳ�����ҳ��Ŀ.������Ŀ�ı��롢���ơ�����."
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmInMedSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum TextIndex
    Num���� = 0: Num����: Num����
End Enum
Private mblnChange As Boolean
Private mstr��ʽ As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    Dim strSQL As String
    
    If isVaild = False Then Exit Sub
On Error GoTo errHandle
    strSQL = "zl_������Ŀ_edit('" & txtEdit(TextIndex.Num����).Text & "','" & txtEdit(TextIndex.Num����).Text & "','" & txtEdit(TextIndex.Num����).Text & "','" & lblEdit(TextIndex.Num����).Tag & "'," & IIf(mstr��ʽ = "����", 0, 1) & ")"
    Call zldatabase.ExecuteProcedure(strSQL, Me.Caption)
    With frmInStationSetup.vsfMain
        If mstr��ʽ = "����" Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = txtEdit(TextIndex.Num����).Text
            .TextMatrix(.Rows - 1, 1) = txtEdit(TextIndex.Num����).Text
            .TextMatrix(.Rows - 1, 2) = txtEdit(TextIndex.Num����).Text
            For i = 0 To txtEdit.Count - 1
                txtEdit(i).Text = ""
            Next i
            txtEdit(TextIndex.Num����).Text = get����
            txtEdit(TextIndex.Num����).SetFocus
            mblnChange = False
        Else
            i = .FindRow(lblEdit(TextIndex.Num����).Tag, , 0)
            .Cell(flexcpText, i, 0, i, .Cols - 1) = ""
            .TextMatrix(i, 0) = txtEdit(TextIndex.Num����).Text
            .TextMatrix(i, 1) = txtEdit(TextIndex.Num����).Text
            .TextMatrix(i, 2) = txtEdit(TextIndex.Num����).Text
            mblnChange = False
            Unload Me
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    txtEdit(TextIndex.Num����).SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    lbl��Ϣ.Caption = "�û��Զ���Ĳ�����ҳ��Ŀ.������Ŀ�ı��롢���ơ�����."
    lbl����.Caption = "����:" & Chr(13) & Chr(10)
    lbl����.Caption = lbl����.Caption & "1.�գ���ʾΪ�ı�¼����Ŀ��������¼������" & Chr(13) & Chr(10)
    lbl����.Caption = lbl����.Caption & "2.ֵ���� AAA,BBB,CCC,DDD����ʾ����������ѡ��ָ������" & Chr(13) & Chr(10)
    lbl����.Caption = lbl����.Caption & "3.�߼�ֵ���� ""�Ƿ�""����ʾ��ѡ��ʽ" & Chr(13) & Chr(10)
    lbl����.Caption = lbl����.Caption & "4.���ַ�Χ -100...100��0.1-0.9:��ʾ����ָ����Χ������" & Chr(13) & Chr(10)
End Sub

Private Function isVaild() As Boolean
    Dim i As Integer
    Dim sngNum1, sngNum2 As Single
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim LngSize As Long
    
    For i = 0 To txtEdit.Count - 1
    '�����Ƿ���������ַ�
        If zlCommFun.StrIsValid(txtEdit(i).Text) = False Then
            txtEdit(i).SetFocus
            isVaild = False
            Exit Function
        End If
        If txtEdit(i).MaxLength <> 0 Then
            If zlCommFun.ActualLen(txtEdit(i).Text) > txtEdit(i).MaxLength Then
                MsgBox "���볤�Ȳ��ܴ���" & "[" & txtEdit(i).MaxLength & "]��", vbInformation, gstrSysName
                isVaild = False
                txtEdit(i).SetFocus
                Exit Function
            End If
        End If
    Next i
    On Error GoTo errH
    'Ϊ�ձ�ʾ��¼����Ŀ�����ٽ������ݼ��
    If txtEdit(2).Text = "" Then isVaild = True: Exit Function
        
    If InStr(txtEdit(TextIndex.Num����).Text, ",") <> 0 Then
        strSQL = "SELECT ��Ϣֵ from ������ҳ�ӱ� where ����id=0 and ��ҳid=0"
'        On Error Resume Next
        Set rsTemp = zldatabase.OpenSQLRecord(strSQL, Me.Caption)
        LngSize = rsTemp.Fields("��Ϣֵ").DefinedSize
        For i = 0 To UBound(Split(txtEdit(TextIndex.Num����).Text, ",")) - 1
            If zlCommFun.ActualLen(Split(txtEdit(TextIndex.Num����).Text, ",")(i)) > LngSize Then
                MsgBox "[" & Split(txtEdit(TextIndex.Num����).Text, ",")(i) & "]ѡ��ĳ��Ȳ��ܳ���" & LngSize & "���޸�!", vbInformation, gstrSysName
                isVaild = False
                txtEdit(TextIndex.Num����).SetFocus
                Exit Function
            End If
        Next i
        isVaild = True
    End If
    
    '�����������ַ�Χ�����Ǽ���Ƿ�Ϸ�
    If InStr(txtEdit(TextIndex.Num����).Text, "...") > 0 Then
        sngNum1 = Mid(txtEdit(2).Text, 1, InStr(txtEdit(2).Text, "...") - 1)
        sngNum2 = Mid(txtEdit(2).Text, InStr(txtEdit(2).Text, "...") + 3)
        If Not IsNumeric(sngNum1) Or Not IsNumeric(sngNum2) Then
            MsgBox "��������ȷ�����ַ�Χ!", vbInformation, gstrSysName
            txtEdit(2).SetFocus
            isVaild = False
            Exit Function
        End If
    ElseIf InStr(txtEdit(TextIndex.Num����).Text, "-") > 0 Then
    
        If InStr(txtEdit(2).Text, "-") = 1 Then
            If (InStr(2, txtEdit(2).Text, "-") - 1) > 0 Then
                sngNum1 = Mid(txtEdit(2).Text, 2, InStr(2, txtEdit(2).Text, "-") - 1)
                sngNum2 = Mid(txtEdit(2).Text, InStr(2, txtEdit(2).Text, "-") + 1)
            End If
        Else
            If Not IsNumeric(Mid(txtEdit(2).Text, 1, InStr(txtEdit(2).Text, "-") - 1)) Or _
                Not IsNumeric(Mid(txtEdit(2).Text, InStr(txtEdit(2).Text, "-") + 1)) Then
                
                MsgBox "��������ȷ�����ַ�Χ!", vbInformation, gstrSysName
                txtEdit(2).SetFocus
                isVaild = False
                Exit Function
            End If
            sngNum1 = Mid(txtEdit(2).Text, 1, InStr(txtEdit(2).Text, "-") - 1)
            sngNum2 = Mid(txtEdit(2).Text, InStr(txtEdit(2).Text, "-") + 1)
        End If
        If Not IsNumeric(sngNum1) Or Not IsNumeric(sngNum2) Then
            MsgBox "��������ȷ�����ַ�Χ!", vbInformation, gstrSysName
            txtEdit(2).SetFocus
            isVaild = False
            Exit Function
        End If
    End If
    
    If sngNum2 < sngNum1 Then
        MsgBox "��������ַ�Χ����,Ӧ��С����.", vbInformation, gstrSysName
        txtEdit(2).SetFocus
        isVaild = False
        Exit Function
    End If
    '�����Ƿ������ݵĶ��巶����
    If InStr(txtEdit(2).Text, ",") = 0 And InStr(txtEdit(2).Text, "�Ƿ�") = 0 And InStr(txtEdit(2).Text, "...") = 0 And InStr(txtEdit(2).Text, "-") = 0 Then
        MsgBox "��������ȷ������Ŀ����!", vbInformation, gstrSysName
        isVaild = False
        txtEdit(2).SetFocus
        Exit Function
    End If
    isVaild = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = True Then
        If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    mblnChange = False
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = TextIndex.Num���� Then
        If InStr("0123456789" & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
    If Index = TextIndex.Num���� Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If zlCommFun.ActualLen(txtEdit(TextIndex.Num����).Text) = 20 Then KeyAscii = 0
        End If
    End If
    If Index = TextIndex.Num���� Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If zlCommFun.ActualLen(txtEdit(TextIndex.Num����).Text) = 1000 Then KeyAscii = 0
        End If
    End If

End Sub

Public Sub ShowMe(str���� As String, str���� As String, str���� As String, str��ʽ As String, frmMain As Object)
    lblEdit(TextIndex.Num����).Tag = str����
    txtEdit(TextIndex.Num����).Text = str����
    
    txtEdit(TextIndex.Num����).Text = str����
    txtEdit(TextIndex.Num����).Text = str����
    mstr��ʽ = str��ʽ
    If str��ʽ = "����" Then
        txtEdit(TextIndex.Num����) = get����
    End If
    mblnChange = False
    Me.Show 1, frmMain
End Sub

Private Function get����() As String
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "select max(to_number(����)) as Maxcode from ������Ŀ"
On Error GoTo errHandle
    Call zldatabase.OpenRecordset(rsTemp, strSQL, Me.Caption)
    get���� = Right("000" & IIf(IsNull(rsTemp!maxcode), 1, rsTemp!maxcode + 1), 3)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


