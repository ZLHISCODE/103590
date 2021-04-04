VERSION 5.00
Begin VB.Form frmRAItemsEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�Զ��������Ŀ"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7110
   Icon            =   "frmRAItemsEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   360
      Left            =   4200
      TabIndex        =   8
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   360
      Left            =   5520
      TabIndex        =   9
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CheckBox chkAdd 
      Caption         =   "���������Զ��������Ŀ(&T)"
      Height          =   180
      Left            =   240
      TabIndex        =   7
      Top             =   2950
      Width           =   2775
   End
   Begin VB.Frame fraSplit 
      Height          =   30
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   6855
   End
   Begin VB.TextBox txtContent 
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1200
      Width           =   4815
   End
   Begin VB.TextBox txtSimName 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1800
      TabIndex        =   3
      Top             =   720
      Width           =   4815
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblContent 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������(&M)"
      Height          =   180
      Left            =   600
      TabIndex        =   4
      Top             =   1230
      Width           =   990
   End
   Begin VB.Label lblSimName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���(&I)"
      Height          =   180
      Left            =   960
      TabIndex        =   2
      Top             =   750
      Width           =   630
   End
   Begin VB.Label lblCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����(&C)"
      Height          =   180
      Left            =   960
      TabIndex        =   0
      Top             =   270
      Width           =   630
   End
End
Attribute VB_Name = "frmRAItemsEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytMode As Byte                '����ģʽ��1-������2-�༭
Private mlngID As Long
'Private mblnOutPati As Boolean          '�������ã��༭ģʽ���õ�
'Private mblnInPati As Boolean           'סԺ���ã��༭ģʽ���õ�
Private mfrmOwner As Form

Public Sub ShowMe(ByVal bytMode As Byte, ByVal lngID As Long, ByVal frmOwner As Form)
'���ܣ��ϲ������ʾ������Ľӿ�
'������
'  bytMode������ģʽ��1-������2-�༭
'  lngID���Զ�����Ŀ��IDֵ�����Բ����룬��ʾ����ģʽ
'  frmOwner�������������

    If bytMode < 1 Or bytMode > 2 Then
        MsgBox "����ģʽ��������ȷ��", vbInformation, gstrSysName
        Exit Sub
    End If

    mbytMode = bytMode
    mlngID = lngID
    Set mfrmOwner = frmOwner
    
    InitCard
    
    Show vbModal, frmOwner

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    '�߼����
    If Validate() = False Then Exit Sub
    '����
    If Save() = False Then Exit Sub
    
    If mbytMode = 1 And chkAdd.Value Then
        '��������
        txtCode.Text = ""
        txtSimName.Text = ""
        txtContent.Text = ""
        txtCode.SetFocus
    Else
        Unload Me
    End If
    
End Sub

Private Function Validate() As Boolean
'���ܣ�����ǰ���߼���֤
'���أ�True�ɹ���Falseʧ��

    If Trim(txtCode.Text) = "" Then
        MsgBox "�����롱δ��д��", vbInformation, gstrSysName
        txtCode.SetFocus
        Exit Function
    End If
    If Trim(txtSimName.Text) = "" Then
        MsgBox "����ơ�δ��д��", vbInformation, gstrSysName
        txtSimName.SetFocus
        Exit Function
    End If
    If Trim(txtContent.Text) = "" Then
        MsgBox "������������δ��д��", vbInformation, gstrSysName
        txtContent.SetFocus
        Exit Function
    End If

    If Len(txtCode.Text) > txtCode.MaxLength Then
        MsgBox FormatEx("�����롱���������������[1]�����ֻ�[2]���ַ���", txtCode.MaxLength \ 2, txtCode.MaxLength), vbInformation, gstrSysName
        txtCode.SetFocus
        Exit Function
    End If
    
    If Len(txtSimName.Text) > txtSimName.MaxLength Then
        MsgBox FormatEx("����ơ����������������[1]�����ֻ�[2]���ַ���", txtSimName.MaxLength \ 2, txtSimName.MaxLength), vbInformation, gstrSysName
        txtSimName.SetFocus
        Exit Function
    End If
    
    If Len(txtContent.Text) > txtContent.MaxLength Then
        MsgBox FormatEx("���������������������������[1]�����ֻ�[2]���ַ���", txtContent.MaxLength \ 2, txtContent.MaxLength), vbInformation, gstrSysName
        txtContent.SetFocus
        Exit Function
    End If

    Validate = True
    
End Function

Private Function Save() As Boolean
'���ܣ���������
'���أ�True�ɹ���Falseʧ��

    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Long

    On Error GoTo errHandle
    If mbytMode = 1 Then
        '����ʱ����ȡIDֵ
        gstrSQL = "Select ���������Ŀ_ID.Nextval as ID From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���������Ŀ��ID")
        If rsTmp.EOF = False Then
            mlngID = rsTmp!ID
        End If
        rsTmp.Close
    End If
    
    With mfrmOwner.vsfItems
        .Redraw = False
        
        If mbytMode = 1 Then
            '����
            .Rows = .Rows + 1
            lngRow = .Rows - 1
            .TextMatrix(lngRow, .ColIndex("����")) = "1"
        Else
            '�༭
            lngRow = .Row
        End If
        .TextMatrix(lngRow, .ColIndex("ID")) = CStr(mlngID)
        .TextMatrix(lngRow, .ColIndex("���")) = "4-�Զ���"
        .TextMatrix(lngRow, .ColIndex("����")) = txtCode.Text
        .TextMatrix(lngRow, .ColIndex("���")) = txtSimName.Text
        .TextMatrix(lngRow, .ColIndex("��������")) = txtContent.Text
        .TextMatrix(lngRow, .ColIndex("�������")) = "2"
        
        .Redraw = True
    End With
    
    Save = True
    Exit Function
    
errHandle:
    If zl9ComLib.ErrCenter = 1 Then Resume
End Function

Private Sub InitCard()
'���ܣ���ʼ���༭��Ƭ

    Dim rsTmp As ADODB.Recordset
    
    '����TextBox��MaxLength
    SetTextMaxLen txtCode, "���������Ŀ.����"
    SetTextMaxLen txtSimName, "���������Ŀ.���"
    SetTextMaxLen txtContent, "���������Ŀ.����"
    
    If mbytMode = 1 Then Exit Sub          '��������ʼ��
    
    chkAdd.Visible = False
    
    On Error GoTo errHandle
    gstrSQL = "Select ����, ���, ����, �Ƿ���������, �Ƿ�סԺ���� From ���������Ŀ Where ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���������Ŀ", mlngID)
    If rsTmp.EOF = False Then
        '�༭���ݿ����Ŀ
        With rsTmp
            '��������
            txtCode.Text = !����
            txtSimName.Text = !���
            txtContent.Text = !����
'            mblnOutPati = Val(zlCommFun.NVL(!�Ƿ���������)) = 1
'            mblnInPati = Val(zlCommFun.NVL(!�Ƿ�סԺ����)) = 1
        End With
    Else
        '�༭����δ���浽���ݿ����Ŀ
        With mfrmOwner.vsfItems
            '��������
            txtCode.Text = .TextMatrix(.Row, .ColIndex("����"))
            txtSimName.Text = .TextMatrix(.Row, .ColIndex("���"))
            txtContent.Text = .TextMatrix(.Row, .ColIndex("��������"))
'            mblnOutPati = .TextMatrix(.Row, .ColIndex("�������"))
'            mblnInPati = .TextMatrix(.Row, .ColIndex("���סԺ"))
        End With
    End If
    rsTmp.Close

    Exit Sub
    
errHandle:
    If zl9ComLib.ErrCenter = 1 Then Resume
End Sub

Private Sub txtCode_GotFocus()
    Call zlControl.TxtSelAll(txtCode)
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If InStr("~`!@#$%^&*()+={[}]|\:;""'<,>.?/", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtContent_GotFocus()
    Call zlControl.TxtSelAll(txtContent)
End Sub

Private Sub txtContent_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txtContent_KeyPress(KeyAscii As Integer)
    If InStr("""'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtSimName_GotFocus()
    Call zlControl.TxtSelAll(txtSimName)
End Sub

Private Sub txtSimName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txtSimName_KeyPress(KeyAscii As Integer)
    If InStr("~`!@#$%^&*()+={[}]|\:;""'<,>.?/", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
