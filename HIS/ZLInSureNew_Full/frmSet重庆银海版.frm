VERSION 5.00
Begin VB.Form frmSet���������� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���в�������"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   Icon            =   "frmSet����������.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraҽԺ�ȼ� 
      Caption         =   "ҽԺ�ȼ�"
      Height          =   1365
      Left            =   150
      TabIndex        =   8
      Top             =   1980
      Width           =   4155
      Begin VB.CommandButton cmdҽԺ�ȼ� 
         Caption         =   "��"
         Height          =   300
         Left            =   3585
         TabIndex        =   12
         Top             =   900
         Width           =   285
      End
      Begin VB.TextBox txtҽԺ�ȼ� 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1350
         MaxLength       =   40
         TabIndex        =   11
         Top             =   900
         Width           =   2235
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ҽԺ�ȼ�(&L)"
         Height          =   180
         Index           =   3
         Left            =   300
         TabIndex        =   10
         Top             =   960
         Width           =   990
      End
      Begin VB.Label lbl˵�� 
         Caption         =   "    �õȼ����ڼ��㲿�ְ�ҽԺ�ȼ������޼۵�������Ŀ��ʵ�ʼ۸�"
         Height          =   480
         Left            =   390
         TabIndex        =   9
         Top             =   330
         Width           =   3450
      End
   End
   Begin VB.Frame fraҽ�������� 
      Caption         =   "ҽԺǰ��ҽ��������"
      Height          =   1605
      Left            =   150
      TabIndex        =   0
      Top             =   210
      Width           =   4155
      Begin VB.CommandButton cmdTest 
         Caption         =   "����(&T)"
         Height          =   1095
         Left            =   3000
         TabIndex        =   7
         Top             =   330
         Width           =   1005
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   6
         Top             =   1110
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1260
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   720
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   2
         Top             =   330
         Width           =   1635
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "������(&S)"
         Height          =   180
         Index           =   2
         Left            =   390
         TabIndex        =   5
         Top             =   1170
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&P)"
         Height          =   180
         Index           =   1
         Left            =   570
         TabIndex        =   3
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�û���(&U)"
         Height          =   180
         Index           =   0
         Left            =   390
         TabIndex        =   1
         Top             =   390
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4560
      TabIndex        =   14
      Top             =   780
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4560
      TabIndex        =   13
      Top             =   300
      Width           =   1100
   End
End
Attribute VB_Name = "frmSet����������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum enum�ı�
    textҽ���û� = 0
    Textҽ������ = 1
    Textҽ�������� = 2
End Enum

Private mblnOK As Boolean
Private mblnChange As Boolean
Private Const madLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Private Const madDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Private Const madDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����

Dim mcnTest As New ADODB.Connection

Private Sub cmdTest_Click()
    Dim rsTemp As New ADODB.Recordset
    If mcnTest.State = adStateOpen Then mcnTest.Close
    
    If OraDataOpen(mcnTest, txtEdit(Textҽ��������).Text, txtEdit(textҽ���û�).Text, txtEdit(Textҽ������).Tag) = False Then
        Exit Sub
    End If
    
    MsgBox "���ӳɹ���", vbInformation, gstrSysName
End Sub

Private Sub cmdҽԺ�ȼ�_Click()
    Dim strFields As String
    Dim rsHos_Info As New ADODB.Recordset
    On Error GoTo errHand
    If Not ҽ����ʼ��_����������(True) Then Exit Sub
    
    '��ʼ���ڲ���¼��
    strFields = "ҽ�ƻ������," & adVarChar & "," & 20 & "|ҽ�ƻ�������," & adVarChar & "," & 50 & _
                "|ҽ�ƻ����ȼ�," & adVarChar & "," & 5 & "|����𸶱�׼," & adVarChar & "," & 20
    Call Record_Init(rsHos_Info, strFields)
    
    '���ýӿڻ�ȡҽ�ƻ�����Ϣ
    Call ���ýӿ�_׼��_����������("05", "C:\CQYB_YH\Hos_info.txt")
    If Not ���ýӿ�_���������� Then Exit Sub
    If Not AnalyFile_HosInfo(rsHos_Info) Then Exit Sub
    
    '�ò���Աѡ��ҽԺ�ȼ�
    If frmListSel.ShowSelect(TYPE_����������, rsHos_Info, "ҽ�ƻ������", "��ѡ��ҽԺ�ȼ���", "�����������Ͽɵ�ҽ�ƻ�����Ϣ:") = True Then
        '01-һ��;05-����;08-����
        txtҽԺ�ȼ�.Tag = rsHos_Info!ҽ�ƻ������
        txtҽԺ�ȼ�.Text = IIf(rsHos_Info!ҽ�ƻ����ȼ� = "01", "����", IIf(rsHos_Info!ҽ�ƻ����ȼ� = "05", "����", "һ��"))
        MsgBox "�ɹ���ȡ��Ժ��ҽԺ�ȼ���ҽԺ���룡", vbInformation, gstrSysName
    End If
    
    rsHos_Info.Filter = 0
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    rsHos_Info.Filter = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If IsValid = False Then Exit Sub
    If SaveData = False Then Exit Sub
    
    mblnOK = True
    mblnChange = False
    Unload Me
End Sub

Private Function IsValid() As Boolean
    Dim lngCount As Long
    Dim strTitle As String
    Dim rsTemp As New ADODB.Recordset
    
    
    For lngCount = txtEdit.LBound To txtEdit.UBound
        If zlCommFun.StrIsValid(txtEdit(lngCount).Text, txtEdit(lngCount).MaxLength) = False Then
            zlControl.TxtSelAll txtEdit(lngCount)
            txtEdit(lngCount).SetFocus
            Exit Function
        End If
    Next
    
    If mcnTest.State = adStateClosed Then
        If OraDataOpen(mcnTest, txtEdit(Textҽ��������).Text, txtEdit(textҽ���û�).Text, txtEdit(Textҽ������).Tag, False) = False Then
            If MsgBox("ҽ�������������������ӣ��Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
        
    IsValid = True
End Function

Private Function SaveData() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lst As ListItem
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    'ɾ���Ѿ�����
    gstrSQL = "zl_���ղ���_Delete(" & TYPE_���������� & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '������������
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_���������� & ",null,'ҽ���û���','" & txtEdit(textҽ���û�).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_���������� & ",null,'ҽ���û�����','" & txtEdit(Textҽ������).Tag & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_���������� & ",null,'ҽ��������','" & txtEdit(Textҽ��������).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_���������� & ",null,'ҽԺ�ȼ�','" & txtҽԺ�ȼ�.Text & "',4)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '����ҽԺ���
    gstrSQL = "Select ����,˵��,�Ƿ��ֹ From ������� Where ���= [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_����������)
    '��������ҽ�������� 204-04-07
    gstrSQL = "zl_�������_Update(" & TYPE_���������� & ",'" & rsTemp!���� & "','" & IIf(IsNull(rsTemp!˵��), "", rsTemp!˵��) & "','" & Me.txtҽԺ�ȼ�.Tag & "'," & IIf(IsNull(rsTemp!�Ƿ��ֹ), 0, rsTemp!�Ƿ��ֹ) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gcnOracle.CommitTrans
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Function

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = Textҽ������ Then
        txtEdit(Index).Tag = txtEdit(Index).Text
    End If
    
    If Index = Textҽ�������� Or Index = Textҽ������ Or Index = textҽ���û� Then
        '�رն�ҽ�������������ӣ���Ϊ�ڲ����������ʱ��Ҫ���´�
        If mcnTest.State = adStateOpen Then mcnTest.Close
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Public Function ��������() As Boolean
'���ܣ������붫�󰢶��ɵ�ҽ���ӿ�
    Dim rsTemp As New ADODB.Recordset
    Dim str����ֵ As String
    
    mblnOK = False
    
    On Error GoTo errHandle
    
    'ȡ���ղ���
    gstrSQL = "select ������,����ֵ from ���ղ��� " & _
              " where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_����������)
    Do Until rsTemp.EOF
        str����ֵ = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        Select Case rsTemp("������")
            Case "ҽ���û���"
                txtEdit(textҽ���û�) = str����ֵ
            Case "ҽ��������"
                txtEdit(Textҽ��������) = str����ֵ
            Case "ҽ���û�����"
                txtEdit(Textҽ������).Text = "        "    '������
                txtEdit(Textҽ������).Tag = str����ֵ
            Case "ҽԺ�ȼ�"
                txtҽԺ�ȼ�.Text = str����ֵ
        End Select
        
        rsTemp.MoveNext
    Loop
    'ȡҽԺ���,��Ϊ����ʱҪ����
    gstrSQL = "Select ҽԺ���� From ������� Where ���= [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_����������)
    txtҽԺ�ȼ�.Tag = Nvl(rsTemp!ҽԺ����)
    
    mblnChange = False
    frmSet����������.Show vbModal, frmҽ�����
    
    �������� = mblnOK
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '��Ӽ�¼
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Private Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '��ʼ��ӳ���¼��
    'strFields:�ֶ���,����,����|�ֶ���,����,����    �������Ϊ��,��ȡĬ�ϳ���
    '�ַ���:adLongVarChar;������:adDouble;������:adDBDate
    
    '���ӣ�
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|��ĿID," & adDouble & ",18|ժҪ, " & adLongVarChar & ",50|" & _
    '"ɾ��," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '��ȡ�ֶ�ȱʡ����
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Function AnalyFile_HosInfo(ByVal rsHos_Info As ADODB.Recordset) As Boolean
    '�����ӿڷ��صĴ����ļ��������浽�м�⣨Ԥ���㷵�صĽ��80%��׼ȷ����˽��鲻���棩
    Dim lngCol As Long, lngCols As Long
    Dim strData As String, strHosinfo As String, strBuffer As String, strFields As String
    Dim arrCol
    Dim objStream As TextStream, objFileSystem As New FileSystemObject
    
    On Error GoTo errHand
    
    If Not objFileSystem.FileExists("C:\CQYB_YH\Hos_info.txt") Then Exit Function
    Set objStream = objFileSystem.OpenTextFile("C:\CQYB_YH\Hos_info.txt", ForReading, False, TristateMixed)
    
    strFields = ""
    For lngCol = 0 To rsHos_Info.Fields.Count - 1
        strFields = strFields & "|" & rsHos_Info.Fields(lngCol).Name
    Next
    strFields = Mid(strFields, 2)
    
    Do While Not objStream.AtEndOfStream
        strBuffer = objStream.ReadLine
        strHosinfo = ""
        arrCol = Split(strBuffer, vbTab)
        lngCols = UBound(arrCol)
        For lngCol = 0 To lngCols
            strHosinfo = strHosinfo & "|" & arrCol(lngCol)
        Next
        strHosinfo = Mid(strHosinfo, 2)
        Call Record_Add(rsHos_Info, strFields, strHosinfo)
    Loop
    objStream.Close
    
    AnalyFile_HosInfo = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
