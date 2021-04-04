VERSION 5.00
Begin VB.Form frmAuditItemTypeEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����������༭"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "frmAuditItemTypeEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   210
      IMEMode         =   3  'DISABLE
      Left            =   990
      TabIndex        =   1
      Tag             =   "MAX"
      Top             =   195
      Width           =   3060
   End
   Begin VB.TextBox txtPriv 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   990
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   990
      Width           =   2745
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   990
      MaxLength       =   30
      TabIndex        =   3
      Top             =   600
      Width           =   3060
   End
   Begin VB.CommandButton cmdSelect 
      Height          =   300
      Left            =   3750
      Picture         =   "frmAuditItemTypeEdit.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   990
      Width           =   300
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3135
      TabIndex        =   9
      Top             =   2055
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2010
      TabIndex        =   8
      Top             =   2055
      Width           =   1100
   End
   Begin VB.CheckBox chkCodeUpdate 
      Caption         =   "������ı��볤�ȣ������˵�����ͬ������(&L)"
      Height          =   285
      Left            =   180
      TabIndex        =   7
      Top             =   1440
      Width           =   4095
   End
   Begin VB.TextBox txtParentCode 
      Enabled         =   0   'False
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   990
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   150
      Width           =   3060
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����(&B)"
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   285
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����(&N)"
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   675
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�ϼ�(&S)"
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   4
      Top             =   1065
      Width           =   630
   End
   Begin VB.Line Line1 
      X1              =   -15
      X2              =   5865
      Y1              =   1830
      Y2              =   1830
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   5865
      Y1              =   1845
      Y2              =   1845
   End
End
Attribute VB_Name = "frmAuditItemTypeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private mblnCancel          As Boolean  'ȷ�� or ȡ��
Private mstrID              As String   '��ǰID
Private mlngProjectID       As Long     '����ID
Private mstrProjectName     As String   '��������
Private zlCheck             As New clsCheck
Dim rsTemp                  As New ADODB.Recordset '��ȡ���ݼ�
Private mEditMode           As �༭ģʽ

Private mlngTypeID          As Long
Private mlngTypePrivID      As Long
Private mstrTypeCode        As String
Private mstrTypeName        As String
Private mintCodeChange      As Integer      '�Ƿ���޸ĳ���
Private mintCodeMaxLength   As Integer      '��¼������޸ĵ���󳤶�
Private mintCodeLenght      As Integer      '��¼�༭ʱ����չ����󳤶�

Public Property Get blnCancel() As Boolean
    blnCancel = mblnCancel
End Property

Public Property Let blnCancel(ByVal vNewValue As Boolean)
    mblnCancel = vNewValue
End Property

Public Property Get lngProjectID() As Long
    lngProjectID = mlngProjectID
End Property

Public Property Let lngProjectID(ByVal vNewValue As Long)
    mlngProjectID = vNewValue
End Property

Public Property Get strProjectName() As String
    strProjectName = mstrProjectName
End Property

Public Property Let strProjectName(ByVal vNewValue As String)
    mstrProjectName = vNewValue
End Property

Public Property Get strID() As String
    strID = mstrID
End Property

Public Property Let strID(ByVal vNewValue As String)
    mstrID = vNewValue
End Property

Public Property Get EditMode() As �༭ģʽ
    EditMode = mEditMode
End Property

Public Property Let EditMode(ByVal vNewValue As �༭ģʽ)
    mEditMode = vNewValue
End Property
 
Private Sub chkCodeUpdate_Click()
On Error GoTo ErrH
    If chkCodeUpdate.Value = 1 Then
        txtCode.MaxLength = mintCodeMaxLength
    Else
        txtCode.MaxLength = mintCodeLenght
    End If
    txtCode.Text = Left(txtCode.Text, txtCode.MaxLength)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub CmdCancel_Click()
    On Error GoTo ErrH
    
    mblnCancel = True
    Unload Me
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub CmdOK_Click()
    Dim strMsg      As String
    Dim varPos      As Variant
    Dim ctrSetF     As Control
    Dim strSQL      As String
    Dim rsTemp      As ADODB.Recordset
    On Error GoTo ErrH
    
    mlngTypeID = Me.strID
    mlngTypePrivID = Val(txtPriv.Tag)
    mstrTypeCode = txtParentCode.Text & txtCode.Text
    mstrTypeName = txtName.Text
    mintCodeChange = chkCodeUpdate.Value
    
    strMsg = ""
    strMsg = zlCheck.Chk_CheckTxtNull("����", txtCode, ctrSetF, strMsg)
    '�������ظ�
    strSQL = "select count(*) from ���������� where ���� = [1] and ID != [2]"
    If zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrTypeCode, mlngTypeID).Fields(0) <> "0" Then
        If ctrSetF Is Nothing Then Set ctrSetF = txtCode
        strMsg = strMsg & "���롾" & mstrTypeCode & "���Ѵ��ڣ�������¼����޸ģ�" & vbCrLf
    End If
    strMsg = zlCheck.Chk_CheckTxtNull("����", txtName, ctrSetF, strMsg)
    '�������ظ�
    If mlngTypePrivID <> "0" Then
        If mEditMode = ���� Then
            strSQL = "select count(*) from ���������� where ���� = [1] and �ϼ�ID = [2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrTypeName, mlngTypePrivID)
        ElseIf mEditMode = �޸� Then
            strSQL = "select count(*) from ���������� where ���� = [1] and �ϼ�ID = [2] and ID != [3]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrTypeName, mlngTypePrivID, mlngTypeID)
        End If
    Else
        strSQL = "select count(*) from ���������� where ���� = [1] and �ϼ�ID is null and ID != [3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrTypeName, mlngTypePrivID, mlngTypeID)
    End If
    
    
    If rsTemp.Fields(0) <> "0" Then
        
        If ctrSetF Is Nothing Then Set ctrSetF = txtName
        strMsg = strMsg & "���ơ�" & mstrTypeName & "���ڷ��ࡾ" & txtPriv.Text & "���Ѵ��ڣ�������¼����޸ģ�" & vbCrLf
    End If
    
    If zlCheck.Chk_CheckMsg(strMsg, ctrSetF) Then Exit Sub
    
     
    If Len(txtParentCode.Text & txtCode.Text) > 10 Then
        MsgBox "�����޷���������һ����������ѡ����������ϼ���"
        Exit Sub
    End If
        
    If mEditMode = �޸� Then
        '��Ᵽ��ID���ϼ�ID�Ƿ����ڵ�ǰID��ǰID���µ���ID
        '�ɵ��ϼ�ID
        Dim intOldTypePrivID        As Integer
        strSQL = "select �ϼ�ID from ���������� where ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngTypePrivID)
        If Not zlCheck.Connection_ChkRsState(rsTemp) Then
            intOldTypePrivID = 0 & rsTemp.Fields!�ϼ�ID
            '�µ��ϼ�ID ���ܴ����� �ɵ��ϼ�ID������ID
            strSQL = "Select * from(SELECT * FROM ���������� START WITH id = [1] CONNECT BY PRIOR ID = �ϼ�ID) where ID =[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngTypeID, mlngTypePrivID)
            
            If Not zlCheck.Connection_ChkRsState(rsTemp) Then
                If rsTemp.Fields(0) > 0 Or mlngTypeID = mlngTypePrivID Then
                    zlCheck.Msg_OK "�µķ��� ���ܴ����� �ɵķ��༰���ӷ�������"
                    Exit Sub
                End If
            End If
        Else
            intOldTypePrivID = -1
            '�µ��ϼ�ID ���ܴ����� �ɵ��ϼ�ID������ID --��Ŀ¼�����ж�
        End If
        '�޸�
        Call AuditItemTypeUpdate
    ElseIf mEditMode = ���� Then
        '����
        Call AuditItemTypeInsert
    End If
    strID = CStr(mlngTypeID)
    mblnCancel = False
    
    Unload Me
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSelect_Click()

    Dim intLenght   As Integer
    Dim strCode     As String
    
    
    On Error GoTo ErrH
    
    With frmAuditItemTypeSelect
        .lngLeft = Me.Left + txtPriv.Left + 10
        .lngTop = Me.Top + txtPriv.Top + txtPriv.Height * 2 + 10
        .intTypeID = Val(txtPriv.Tag)
        .Show vbModal
        If .blnCancel Then
            Set frmAuditItemTypeSelect = Nothing
            Exit Sub
        End If
        lngProjectID = .lngProjectID
        strProjectName = .strProjectName
        txtPriv.Tag = CStr(.intTypeID)
    End With
    
    
    intLenght = 0
    gstrSQL = "select /*+ rule */id,�ϼ�ID,����,���� from ���������� a Where a.id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txtPriv.Tag)
    If Not zlCheck.Connection_ChkRsState(rsTemp) Then
        intLenght = Len("" & rsTemp!����)
        txtParentCode.Text = "" & rsTemp!����
        txtPriv.Tag = "" & rsTemp!ID
        txtPriv.Text = "[" + rsTemp!���� + "]" & rsTemp!����
    Else
        txtParentCode.Text = ""
    End If
    mintCodeMaxLength = 10 - intLenght
    mintCodeLenght = intLenght
    
    txtCode.Move txtParentCode.Left + 20 + intLenght * 100, txtCode.Top, txtParentCode.Width - 50 - intLenght * 100
    If txtPriv.Tag = "-1" Then
        gstrSQL = "select max(����) as ���� from ���������� a Where a.�ϼ�ID is null"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strID)
        txtPriv.Text = strProjectName
        
    Else
        gstrSQL = "select max(����) as ���� from ���������� a Where a.�ϼ�ID = [1]"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strID)
    End If
    strCode = ""
    If Not zlCheck.Connection_ChkRsState(rsTemp) Then
        strCode = Mid("" & rsTemp!����, Len(txtParentCode.Text) + 1)
        txtCode.Text = IncStr(strCode)
        mintCodeMaxLength = mintCodeMaxLength - Len(strCode)
        mintCodeLenght = Len(strCode)
    End If
    If strCode = "" Then txtCode.Text = "01": mintCodeLenght = 2
    txtCode.MaxLength = mintCodeLenght
    chkCodeUpdate.Value = 1
    chkCodeUpdate.Enabled = False
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ��ʼ���޸� InitUpdate
'==============================================================================
Public Sub InitUpdate()
    Dim strCode             As String
    Dim strFormat           As String
    Dim i                   As Integer
    Dim intLenght           As Integer
    Dim intPrivLenght       As Integer
    
    On Error GoTo ErrH:
    
    intLenght = 0
    gstrSQL = "select /*+ rule */id,�ϼ�ID,����,���� from ���������� a Where a.id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strID)
    If Not zlCheck.Connection_ChkRsState(rsTemp) Then
        strCode = "" & rsTemp!����
        txtName.Text = "" & rsTemp!����
        intLenght = Len("" & rsTemp!����)
        gstrSQL = "select /*+ rule */id,�ϼ�ID,����,���� from ���������� a Where a.ID=[1]"
        txtPriv.Tag = "" & rsTemp!�ϼ�ID
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "" & rsTemp!�ϼ�ID)
        If Not zlCheck.Connection_ChkRsState(rsTemp) Then
            intPrivLenght = Len("" & rsTemp!����)
            If Left(strCode, intPrivLenght) = "" & rsTemp!���� Then
                txtParentCode.Text = Left("" & rsTemp!����, Len("" & rsTemp!����))
                txtPriv.Text = "[" + txtParentCode.Text + "]" & rsTemp!����
            Else
                intPrivLenght = 0
            End If
        Else
            txtPriv.Text = strProjectName
            txtPriv.Tag = "0"
        End If
    End If
    mintCodeMaxLength = 10 - intPrivLenght
    mintCodeLenght = Len(strCode) - intPrivLenght
    
    txtCode.MaxLength = mintCodeLenght
    txtCode.Move txtParentCode.Left + 20 + intPrivLenght * 100, txtCode.Top, txtParentCode.Width - 50 - intPrivLenght * 100
    txtCode.Text = Mid(strCode, intPrivLenght + 1)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ��ʼ������ InitInsert
'==============================================================================
Private Sub InitInsert()
    Dim strCode     As String
    Dim strFormat   As String
    Dim i           As Integer
    Dim intLenght   As Integer
    
    On Error GoTo ErrH:
    
    intLenght = 0
    gstrSQL = "select /*+ rule */id,�ϼ�ID,����,���� from ���������� a Where a.id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strID)
    If Not zlCheck.Connection_ChkRsState(rsTemp) Then
        intLenght = Len("" & rsTemp!����)
        txtParentCode.Text = "" & rsTemp!����
        txtPriv.Tag = "" & rsTemp!ID
        txtPriv.Text = "[" + rsTemp!���� + "]" & rsTemp!����
    Else
        txtParentCode.Text = ""
    End If
    mintCodeMaxLength = 10 - intLenght
    mintCodeLenght = intLenght
    
    txtCode.Move txtParentCode.Left + 20 + intLenght * 100, txtCode.Top, txtParentCode.Width - 50 - intLenght * 100
    If strID = "-1" Then
        gstrSQL = "select max(����) as ���� from ���������� a Where a.�ϼ�ID is null"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strID)
        txtPriv.Text = strProjectName
        
    Else
        gstrSQL = "select max(����) as ���� from ���������� a Where a.�ϼ�ID = [1]"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strID)
    End If
    strCode = ""
    If Not zlCheck.Connection_ChkRsState(rsTemp) Then
        strCode = Mid("" & rsTemp!����, Len(txtParentCode.Text) + 1)
        txtCode.Text = IncStr(strCode)
        mintCodeMaxLength = mintCodeMaxLength - Len(strCode)
        mintCodeLenght = Len(strCode)
    End If
    If strCode = "" Then txtCode.Text = "01": mintCodeLenght = 2
    txtCode.MaxLength = mintCodeLenght
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim rsUsed      As ADODB.Recordset
    On Error GoTo ErrH
    '�ֶο��
    Set rsUsed = zlCheck.GetRsFieldWidth("����������")
    rsUsed.Filter = "����='" & txtName.DataField & "'"
    If Not zlCheck.Connection_ChkRsState(rsUsed) Then txtName.MaxLength = "" & rsUsed.Fields("����")
    rsUsed.Filter = "����='" & txtCode.DataField & "'"
    If Not zlCheck.Connection_ChkRsState(rsUsed) Then txtCode.MaxLength = "" & rsUsed.Fields("����")
    
    zlCheck.Sys_System Me
    If EditMode = �޸� Then
        Call InitUpdate
    ElseIf EditMode = ���� Then
        Call InitInsert
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'========================================================================================
'=����
'========================================================================================
Public Sub AuditItemTypeInsert()
    On Error GoTo ErrH
    
    mlngTypeID = zlDatabase.GetNextId("����������")
    gstrSQL = "Zl_����������_Insert (" + CStr(mlngTypeID) + "," + IIf(mlngTypePrivID = 0, "NULL", CStr(mlngTypePrivID)) + "," + "'" + mstrTypeCode + "'" + "," + "'" + mstrTypeName + "'," & CStr(mintCodeChange) & "," & lngProjectID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'========================================================================================
'=�޸�
'========================================================================================
Public Sub AuditItemTypeUpdate()
    On Error GoTo ErrH
    
    gstrSQL = "Zl_����������_Update (" + CStr(mlngTypeID) + "," + IIf(mlngTypePrivID <= 0, "NULL", CStr(mlngTypePrivID)) + "," + "'" + mstrTypeCode + "'" + "," + "'" + mstrTypeName + "'," & CStr(mintCodeChange) & "," & lngProjectID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'========================================================================================
'=ɾ��
'========================================================================================
Public Sub AuditItemTypeDelete()
    Dim strSQL      As String
    On Error GoTo ErrH
    DoEvents
    strSQL = "Zl_����������_Delete (" & strID & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption

    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtCode_Change()
    Dim lngStart        As Long
    Dim lngLength       As Long
    Dim i               As Long
On Error GoTo ErrH
    
    lngLength = Len(txtCode.Text)
    lngStart = txtCode.SelStart
    For i = 1 To Len(txtCode.Text)
        If InStr(1, "0123456789", Mid(txtCode.Text, i, 1)) = 0 Then
            txtCode.Text = Replace(txtCode.Text, Mid(txtCode.Text, i, 1), "")
        End If
    Next
    
    If lngStart - (lngLength - Len(txtCode.Text)) > 0 Then txtCode.SelStart = lngStart - (lngLength - Len(txtCode.Text))
 
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrH
    If KeyAscii = 8 Then Exit Sub
    If InStr(1, "0123456789", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtCode_Validate(Cancel As Boolean)
    Dim strSQL      As String
    Dim i           As Integer
    Dim strFormat   As String
    On Error GoTo ErrH
    If chkCodeUpdate.Value <> 1 Then
        For i = 1 To txtCode.MaxLength
            strFormat = strFormat & "0"
        Next
        If txtCode.MaxLength > 0 Then txtCode.Text = Right(strFormat & txtCode.Text, Len(strFormat))
    End If
    DoEvents
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

