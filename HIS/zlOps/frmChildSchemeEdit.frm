VERSION 5.00
Begin VB.Form frmChildSchemeEdit 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame fra 
      Height          =   3405
      Index           =   0
      Left            =   30
      TabIndex        =   8
      Top             =   0
      Width           =   8070
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1185
         TabIndex        =   3
         Top             =   570
         Width           =   1260
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   1185
         TabIndex        =   1
         Top             =   195
         Width           =   6450
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   3
         Left            =   3555
         TabIndex        =   5
         Top             =   570
         Width           =   4080
      End
      Begin VB.TextBox txt 
         Height          =   1710
         Index           =   0
         Left            =   1185
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   945
         Width           =   6120
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&B)"
         Height          =   180
         Index           =   0
         Left            =   135
         TabIndex        =   2
         Top             =   645
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&N)"
         Height          =   180
         Index           =   2
         Left            =   135
         TabIndex        =   0
         Top             =   270
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&S)"
         Height          =   180
         Index           =   3
         Left            =   2505
         TabIndex        =   4
         Top             =   645
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����˵��(&Z)"
         Height          =   180
         Index           =   9
         Left            =   135
         TabIndex        =   6
         Top             =   990
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmChildSchemeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################
Private mlngKey As Long
Private mlngReferKey As Long
Private mfrmMain As Object
Private mbytMode As Byte
Private mblnAllowModify As Boolean
Private mblnDataChanged As Boolean
Private mblnReading As Boolean
Public Event AfterDataChanged()

Public Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
    
    If mblnReading = False Then
        RaiseEvent AfterDataChanged
    End If
End Property

Public Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property


Public Property Let AllowModify(ByVal blnData As Boolean)
    mblnAllowModify = blnData
    Call ExecuteCommand("�ؼ�״̬")
End Property

Public Property Get AllowModify() As Boolean
    AllowModify = mblnAllowModify
End Property

Public Function InitData(ByVal frmMain As Object, Optional ByVal blnAllowModify As Boolean = True) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mblnAllowModify = blnAllowModify
    Set mfrmMain = frmMain
    
    Call ExecuteCommand("�ؼ�״̬")
    
    If ExecuteCommand("��ʼ����") = False Then Exit Function
    
    DataChanged = False
    
    InitData = True
    
End Function

Public Function RefreshData(ByVal lngKey As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************

    mlngKey = lngKey
    mbytMode = 2
    
    Call ExecuteCommand("�������")
    Call ExecuteCommand("�ؼ�״̬")
    
    If mlngKey > 0 Then
        If ExecuteCommand("��ȡ����") = False Then Exit Function
    End If

    DataChanged = False
    
    RefreshData = True
    
End Function

Public Function NewData(ByVal lngKey As Long, Optional ByVal lngReferKey As Long = 0) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************

    mlngKey = lngKey
    mlngReferKey = lngReferKey
    
    If mlngReferKey > 0 Then
        mlngKey = mlngReferKey
        Call ExecuteCommand("��ȡ����")
        mlngKey = lngKey
    End If
    
    mbytMode = 1
    
    Call ExecuteCommand("�������")
    Call ExecuteCommand("�ؼ�״̬")
    Call ExecuteCommand("ȱʡ����")
    
    DataChanged = True
    
    Call LocationObj(txt(2))
        
    NewData = True
End Function

Public Function ValidData() As Boolean
    '******************************************************************************************************************
    '���ܣ�У��༭���ݵ���Ч��
    '������
    '���أ�
    '******************************************************************************************************************
    
    If txt(3).Text = "" Then txt(3).Text = zlGetSymbol(txt(2).Text)
    
    If StrIsValid(txt(0).Text, txt(0).MaxLength) = False Then
        Call LocationObj(txt(0))
        Exit Function
    End If
    
    If StrIsValid(txt(1).Text, txt(1).MaxLength) = False Then
        Call LocationObj(txt(1))
        Exit Function
    End If
    
    If StrIsValid(txt(2).Text, txt(2).MaxLength) = False Then
        Call LocationObj(txt(2))
        Exit Function
    End If
    
    If StrIsValid(txt(3).Text, txt(3).MaxLength) = False Then
        Call LocationObj(txt(3))
        Exit Function
    End If
    
    If txt(1).Text = "" Then
        ShowSimpleMsg "���벻��Ϊ��ֵ���������룡"
        Call LocationObj(txt(1))
        Exit Function
    End If
        
    If CheckAllNumber(txt(1).Text) = False Then
        ShowSimpleMsg "�������Ϊ��������ɣ�"
        Call LocationObj(txt(1))
        Exit Function
    End If
    
    If txt(2).Text = "" Then
        ShowSimpleMsg "���Ʋ���Ϊ��ֵ���������룡"
        
        Call LocationObj(txt(2))
        Exit Function
    End If
    
    
    ValidData = True
    
End Function

Public Function SaveData(ByRef rsSQL As ADODB.Recordset, ByRef lngKey As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strSQL As String

    On Error GoTo errHand

    If mlngKey = 0 Then
        '����
        lngKey = zlDatabase.GetNextId("���������ο�")
        
        strSQL = "ZL_���������ο�_INSERT(" & lngKey & ",'" & txt(1).Text & "','" & txt(2).Text & "','" & txt(3).Text & "','" & txt(0).Text & "',1)"
        
    Else
        '�޸�
        lngKey = mlngKey
        
        strSQL = "ZL_���������ο�_UPDATE(" & lngKey & ",'" & txt(1).Text & "','" & txt(2).Text & "','" & txt(3).Text & "','" & txt(0).Text & "',1)"
        
    End If
    Call SQLRecordAdd(rsSQL, strSQL)
            
    SaveData = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim blnAllowModify As Boolean
    Dim strSQL As String
    
    On Error GoTo errHand
    
    mblnReading = True
    
    Select Case strCommand
    '--------------------------------------------------------------------------------------------------------------
    Case "��ʼ�ؼ�"
        
    '--------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"
                 
        '����������볤��
        txt(0).MaxLength = GetMaxLength("���������ο�", "˵��")
        txt(1).MaxLength = GetMaxLength("�����ĿĿ¼", "����")
        txt(2).MaxLength = GetMaxLength("�����ĿĿ¼", "����")
        txt(3).MaxLength = GetMaxLength("�����ĿĿ¼", "����")
            
    '--------------------------------------------------------------------------------------------------------------
    Case "�ؼ�״̬"
    
        blnAllowModify = mblnAllowModify
        If mlngKey = 0 And mbytMode = 2 Then blnAllowModify = False
        
        txt(0).Locked = Not blnAllowModify
        txt(1).Locked = Not blnAllowModify
        txt(2).Locked = Not blnAllowModify
        txt(3).Locked = Not blnAllowModify
            
    Case "�������"
        
        txt(1).Text = ""
        txt(2).Text = ""
        txt(3).Text = ""
        
        If mlngReferKey = 0 Then txt(0).Text = ""
        
    Case "��ȡ����"
        
        strSQL = "SELECT ����,����,����,˵�� FROM ���������ο� WHERE ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, mfrmMain.Caption, mlngKey)
        If rs.BOF = False Then
            txt(1).Text = zlCommFun.NVL(rs("����").Value)
            txt(2).Text = zlCommFun.NVL(rs("����").Value)
            txt(3).Text = zlCommFun.NVL(rs("����").Value)
            txt(0).Text = zlCommFun.NVL(rs("˵��").Value)
        End If
        
    '--------------------------------------------------------------------------------------------------------------
    Case "ȱʡ����"
                
        txt(1).Text = GetNextCode("���������ο�")
        
    End Select
    
    mblnReading = False
    
    ExecuteCommand = True

    Exit Function
    
errHand:

    mblnReading = False
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Sub Form_Load()
    fra(0).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    fra(0).Move 0, -75, Me.ScaleWidth, Me.ScaleHeight + 75
    
    txt(1).Move txt(1).Left, txt(1).Top
    txt(2).Move txt(2).Left, txt(2).Top, fra(0).Width - txt(2).Left - 75
    txt(3).Move txt(3).Left, txt(3).Top, fra(0).Width - txt(3).Left - 75
    txt(0).Move txt(0).Left, txt(0).Top, fra(0).Width - txt(0).Left - 75, fra(0).Height - txt(0).Top - 75

End Sub

Private Sub txt_Change(Index As Integer)
    DataChanged = True
End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txt(Index)
    
    Select Case Index
    Case 0, 2
        zlCommFun.OpenIme True
    End Select
        
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0

        Select Case Index
        Case 1
            If FilterKeyAscii(KeyAscii, 99, "0123456789") = 0 Then KeyAscii = 0
        Case 3
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End Select
    End If
    
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 0, 2
        zlCommFun.OpenIme False
        If Index = 2 Then
            If InStr(txt(Index).Text, "'") = 0 Then txt(3).Text = zlGetSymbol(txt(Index).Text)
        End If
    End Select
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    If Cancel Then Exit Sub
    
End Sub

