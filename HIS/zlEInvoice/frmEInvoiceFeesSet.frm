VERSION 5.00
Begin VB.Form frmEInvoiceFeeseSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�վݷ�Ŀ����"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5100
   Icon            =   "frmEInvoiceFeesSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.OptionButton Option���� 
      Caption         =   "סԺ"
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   13
      Top             =   2040
      Width           =   700
   End
   Begin VB.CommandButton cmd��Ŀ 
      Caption         =   "��"
      Height          =   250
      Left            =   3120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   255
   End
   Begin VB.OptionButton Option���� 
      Caption         =   "����"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   5
      Top             =   2040
      Width           =   700
   End
   Begin VB.OptionButton Option���� 
      Caption         =   "������"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   4
      Top             =   2040
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   2
      Left            =   960
      TabIndex        =   3
      Top             =   1462
      Width           =   2475
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   1
      Left            =   960
      MaxLength       =   20
      TabIndex        =   2
      Top             =   920
      Width           =   2475
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   335
      Width           =   2115
   End
   Begin VB.Frame fra 
      Height          =   3400
      Left            =   3600
      TabIndex        =   8
      Top             =   -120
      Width           =   10
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3840
      TabIndex        =   7
      Top             =   840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3840
      TabIndex        =   6
      Top             =   360
      Width           =   1100
   End
   Begin VB.Label lbl 
      Caption         =   "���ó���"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   730
   End
   Begin VB.Label lbl 
      Caption         =   "��    ��"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   1485
      Width           =   730
   End
   Begin VB.Label lbl 
      Caption         =   "��    ��"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   943
      Width           =   730
   End
   Begin VB.Label lbl 
      Caption         =   "�վݷ�Ŀ"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   358
      Width           =   730
   End
End
Attribute VB_Name = "frmEInvoiceFeeseSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TXT_Idex
    Idex_��Ŀ = 0
    Idex_���� = 1
    Idex_���� = 2
End Enum
Private mlngID As Long      '�վݷ�Ŀ����.ID���޸��Ǵ��룬������Ϊ0
Private mblnOK As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If IsValid = False Then Exit Sub
    If Save�վݷ�Ŀ���� = False Then Exit Sub
    mblnOK = True
    Unload Me
End Sub

Private Function Save�վݷ�Ŀ����() As Boolean
    Dim strSQL As String
    Dim lngNewID As Long, int���� As Integer
    
    On Error GoTo errHandle
    int���� = IIf(Option����(0).Value = True, 0, IIf(Option����(1).Value = True, 1, 2))
    If mlngID = 0 Then
        '�����վݷ�Ŀ����
        lngNewID = zlDatabase.GetNextId("�վݷ�Ŀ����")
        strSQL = "Zl_�վݷ�Ŀ����_Update("
'        ��������_In In Number,
        strSQL = strSQL & 0 & ","
'        Id_In       In �վݷ�Ŀ����.Id%Type,
        strSQL = strSQL & lngNewID & ","
'        �վݷ�Ŀ_In In �վݷ�Ŀ����.�վݷ�Ŀ%Type,
        strSQL = strSQL & "'" & txtEdit(Idex_��Ŀ).Text & "',"
'        ����_In     In �վݷ�Ŀ����.����%Type,
        strSQL = strSQL & "'" & txtEdit(Idex_����).Text & "',"
'        ����_In     In �վݷ�Ŀ����.����%Type,
        strSQL = strSQL & "'" & txtEdit(Idex_����).Text & "',"
'        ���ó���_In In �վݷ�Ŀ����.���ó���%Type
        strSQL = strSQL & int���� & ")"
    Else
        '�޸��վݷ�Ŀ����
        strSQL = "Zl_�վݷ�Ŀ����_Update("
'        ��������_In In Number,
        strSQL = strSQL & 1 & ","
'        Id_In       In �վݷ�Ŀ����.Id%Type,
        strSQL = strSQL & mlngID & ","
'        �վݷ�Ŀ_In In �վݷ�Ŀ����.�վݷ�Ŀ%Type,
        strSQL = strSQL & "'" & txtEdit(Idex_��Ŀ).Text & "',"
'        ����_In     In �վݷ�Ŀ����.����%Type,
        strSQL = strSQL & "'" & txtEdit(Idex_����).Text & "',"
'        ����_In     In �վݷ�Ŀ����.����%Type,
        strSQL = strSQL & "'" & txtEdit(Idex_����).Text & "',"
'        ���ó���_In In �վݷ�Ŀ����.���ó���%Type
        strSQL = strSQL & int���� & ")"
    End If
    Call zlDatabase.ExecuteProcedure(strSQL, "�վݷ�Ŀ����")
    
    Save�վݷ�Ŀ���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmd��Ŀ_Click()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim vRect As RECT
    
    strSQL = "Select Rownum As id, ����, Upper(����) as ����,Upper(����) as ���� From �վݷ�Ŀ Order  By ���� "
    vRect = zlControl.GetControlRect(txtEdit(Idex_��Ŀ).hWnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��ȡ�վݷ�Ŀ", False, "", "", False, False, _
            True, vRect.Left, vRect.Top, txtEdit(Idex_��Ŀ).Height, True, False, False)
     If Not rsTemp Is Nothing Then
        txtEdit(Idex_��Ŀ).Text = rsTemp("����")
    End If
    zlControl.ControlSetFocus txtEdit(Idex_��Ŀ)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    End If
    
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    Call Load�վݷ�ĿFromID(mlngID)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlngID = 0
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim vRect As RECT
    
    If Index = Idex_���� Then
        If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
    ElseIf Index = Idex_���� Then
        If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    ElseIf Index = Idex_��Ŀ Then
        If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        If KeyAscii <> vbKeyReturn Then Exit Sub
        strSQL = "Select Rownum As id, ����, Upper(����) as ����,Upper(����) as ���� From �վݷ�Ŀ " & _
                  "Where ���� Like Upper([1]) Or Upper(����) Like Upper([1]) Or Upper(����)  Like Upper([1]) " & _
                  "   Or Upper(zlPinYinCode(����)) Like Upper([1]) Order By ���� "
                  
        vRect = zlControl.GetControlRect(txtEdit(Idex_��Ŀ).hWnd)
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��ȡ�ͻ���", False, "", "", False, False, _
                True, vRect.Left, vRect.Top, txtEdit(Idex_��Ŀ).Height, True, False, False, "%" & txtEdit(Idex_��Ŀ).Text & "%")
         If Not rsTemp Is Nothing Then
            txtEdit(Idex_��Ŀ).Text = rsTemp("����")
            zlControl.ControlSetFocus txtEdit(Idex_��Ŀ)
        Else
            MsgBox "�����������Ϣδ�ҵ���Ч���վݷ�Ŀ�������ԣ�", vbInformation, gstrSysName
            txtEdit(Idex_��Ŀ).Text = ""
            zlControl.ControlSetFocus txtEdit(Idex_��Ŀ)
        End If
    End If
End Sub

Public Sub ShowMe(ByVal frmMain As Object, Optional ByVal lngID As Long, Optional blnRefresh As Boolean)
    mlngID = lngID
    mblnOK = False
    Me.Show 1, frmMain
    blnRefresh = mblnOK
End Sub

Private Function IsValid() As Boolean
    On Error GoTo errHandle

    If Len(Trim(txtEdit(Idex_��Ŀ).Text)) = 0 Then
        MsgBox "�վݷ�Ŀ����Ϊ�ա�", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(Idex_��Ŀ)
        Exit Function
    End If
    
    If Len(txtEdit(Idex_����).Text) = 0 Then
        MsgBox "���벻��Ϊ�ա�", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(Idex_����)
        Exit Function
    End If

    If Not IsNumeric(txtEdit(Idex_����).Text) Or InStr(txtEdit(Idex_����).Text, ",") > 0 Or InStr(txtEdit(Idex_����).Text, ".") > 0 Then
        MsgBox "����Ӧ��������ɡ�", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(Idex_����)
        Exit Function
    End If
    
    If Len(Trim(txtEdit(Idex_����).Text)) = 0 Then
        MsgBox "���Ʋ���Ϊ�ա�", vbExclamation, gstrSysName
        txtEdit(Idex_����).Text = ""
        zlControl.ControlSetFocus txtEdit(Idex_����)
        Exit Function
    End If
    
    If LenB(StrConv(txtEdit(Idex_����).Text, vbFromUnicode)) > 20 Then
        MsgBox "���Ƴ��Ȳ��ܳ���10�����ֻ���20���ַ���������¼�룡", vbInformation, gstrSysName
        zlControl.ControlSetFocus txtEdit(Idex_����)
        Exit Function
    End If

    IsValid = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Load�վݷ�ĿFromID(ByVal lngID As Long)
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim int���� As Integer
    
    If lngID = 0 Then Exit Sub
    strSQL = "Select �վݷ�Ŀ,����,����,���ó��� From �վݷ�Ŀ���� where ID =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
    If rsTemp.EOF Then Exit Sub
    With rsTemp
        txtEdit(Idex_��Ŀ).Text = NVL(!�վݷ�Ŀ)
        txtEdit(Idex_����).Text = NVL(!����)
        txtEdit(Idex_����).Text = NVL(!����)
        int���� = Val(!���ó���)
        Option����(int����).Value = True
    End With
End Sub


