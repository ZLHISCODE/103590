VERSION 5.00
Begin VB.Form frmDrugProducer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҩƷ������"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7545
   Icon            =   "frmDrugProducer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   375
      Left            =   6360
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame frmLine 
      Height          =   1575
      Left            =   6000
      TabIndex        =   6
      Top             =   -120
      Width           =   30
   End
   Begin VB.TextBox txtProducer 
      Enabled         =   0   'False
      Height          =   270
      Index           =   2
      Left            =   1080
      TabIndex        =   5
      Top             =   915
      Width           =   4695
   End
   Begin VB.TextBox txtProducer 
      Height          =   270
      Index           =   1
      Left            =   1080
      MaxLength       =   60
      TabIndex        =   4
      Top             =   555
      Width           =   4695
   End
   Begin VB.TextBox txtProducer 
      Enabled         =   0   'False
      Height          =   270
      Index           =   0
      Left            =   1080
      TabIndex        =   3
      Top             =   195
      Width           =   1215
   End
   Begin VB.Label lblProducer 
      AutoSize        =   -1  'True
      Caption         =   "����(&3)"
      Height          =   180
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   630
   End
   Begin VB.Label lblProducer 
      AutoSize        =   -1  'True
      Caption         =   "����(&2)"
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   630
   End
   Begin VB.Label lblProducer 
      AutoSize        =   -1  'True
      Caption         =   "����(&1)"
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   630
   End
End
Attribute VB_Name = "frmDrugProducer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsMaxs As New Recordset
    Dim ints���� As Integer, strCodes As String, strSpecifys As String
    
    On Error GoTo errHandle
    
    If Trim(txtProducer(1).Text) = "" Then
        MsgBox "���Ʋ���Ϊ�գ����飡", vbExclamation, gstrSysName
        txtProducer(1).SetFocus
        Exit Sub
    End If
    
    If LenB(StrConv(txtProducer(1).Text, vbFromUnicode)) > txtProducer(1).MaxLength Then
        MsgBox "���������ݲ��ܳ���" & Int(txtProducer(1).MaxLength / 2) & "�����ֻ�" & txtProducer(1).MaxLength & "���ַ�!", vbExclamation + vbOKOnly, gstrSysName
        txtProducer(1).SetFocus
        Exit Sub
    End If
    
    '����
    gstrSQL = "ZL_ҩƷ������_INSERT('" & txtProducer(0).Text & "','" & txtProducer(1).Text & "',substr('" & txtProducer(2).Text & "',0,10))"
    Call zldatabase.ExecuteProcedure(gstrSQL, "")
    
    'ˢ�½��棬�����ٴ�����
    gstrSQL = "SELECT Nvl(MAX(LENGTH(����)),2) As Length FROM ҩƷ������"
    Set rsMaxs = zldatabase.OpenSQLRecord(gstrSQL, "" & "-ҩƷ�����̱��볤��")
    ints���� = rsMaxs!length
    
    gstrSQL = "SELECT Nvl(MAX(LPAD(����," & ints���� & ",'0')),'00') As Code FROM ҩƷ������"
    Set rsMaxs = zldatabase.OpenSQLRecord(gstrSQL, "" & "-ҩƷ�����̱���")
    strCodes = rsMaxs!Code
    
    ints���� = Len(strCodes)
    strCodes = strCodes + 1
    If ints���� >= Len(strCodes) Then
        strCodes = String(ints���� - Len(strCodes), "0") & strCodes
    End If
    
    txtProducer(0).Text = strCodes
    txtProducer(1).Text = ""
    txtProducer(2).Text = ""
    
    txtProducer(1).SetFocus
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim rsMaxs As New Recordset
    Dim ints���� As Integer, strCodes As String, strSpecifys As String

    On Error GoTo errHandle
    
    Call GetDefineSize
    
    gstrSQL = "SELECT Nvl(MAX(LENGTH(����)),2) As Length FROM ҩƷ������"
    Set rsMaxs = zldatabase.OpenSQLRecord(gstrSQL, "" & "-ҩƷ�����̱��볤��")
    ints���� = rsMaxs!length
    
    gstrSQL = "SELECT Nvl(MAX(LPAD(����," & ints���� & ",'0')),'00') As Code FROM ҩƷ������"
    Set rsMaxs = zldatabase.OpenSQLRecord(gstrSQL, "" & "-ҩƷ�����̱���")
    strCodes = rsMaxs!Code
    
    ints���� = Len(strCodes)
    strCodes = strCodes + 1
    If ints���� >= Len(strCodes) Then
        strCodes = String(ints���� - Len(strCodes), "0") & strCodes
    End If
    
    txtProducer(0).Text = strCodes
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtProducer_Change(Index As Integer)
    txtProducer(2).Text = zlStr.GetCodeByVB(txtProducer(1).Text)
End Sub

Private Sub txtProducer_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 3 Then
        If InStr("0123456789", Chr(KeyAscii)) > 0 Or KeyAscii = 8 Then
            KeyAscii = KeyAscii
        Else
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub GetDefineSize()
    '���ܣ��õ����ݿ�ı��ֶεĳ���
    On Error GoTo errHandle
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
     
    strSQL = "Select t.�ϴβ��� as ������ From ҩƷ��� T Where Rownum < 1"
    Call zldatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    txtProducer(1).MaxLength = rsTmp.Fields("������").DefinedSize
   
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
