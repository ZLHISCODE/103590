VERSION 5.00
Begin VB.Form frm�������㷽ʽ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������㷽ʽ"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "frm�������㷽ʽ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txt�����ֵ����� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1980
      TabIndex        =   12
      Top             =   2880
      Width           =   3885
   End
   Begin VB.TextBox txt���������׼ 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1980
      TabIndex        =   10
      Top             =   2490
      Width           =   3885
   End
   Begin VB.TextBox txt����ͳ��ֵ����� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1980
      TabIndex        =   8
      Top             =   2100
      Width           =   3885
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   30
      TabIndex        =   16
      Top             =   3390
      Width           =   6165
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "����(&W)"
      Height          =   350
      Left            =   180
      TabIndex        =   15
      Top             =   3570
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3390
      TabIndex        =   13
      Top             =   3570
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4650
      TabIndex        =   14
      Top             =   3570
      Width           =   1100
   End
   Begin VB.TextBox txt����ͳ�������׼ 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1980
      TabIndex        =   6
      Top             =   1710
      Width           =   3885
   End
   Begin VB.TextBox txt������Ϣ 
      Height          =   300
      Left            =   1980
      TabIndex        =   3
      Top             =   1320
      Width           =   3615
   End
   Begin VB.CommandButton cmd������Ϣ 
      Caption         =   "��"
      Height          =   300
      Left            =   5580
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1320
      Width           =   285
   End
   Begin VB.Label lbl�����ֵ����� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�����ֵ�����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   11
      Top             =   2940
      Width           =   1440
   End
   Begin VB.Label lbl���������׼ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���������׼"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   9
      Top             =   2550
      Width           =   1440
   End
   Begin VB.Label lbl����ͳ��ֵ����� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����ͳ��ֵ�����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   7
      Top             =   2160
      Width           =   1440
   End
   Begin VB.Label lblNote 
      Caption         =   "    ��ѡ��һ�������֣�����סԺ�����ò��ֶ�Ӧ�����㷽ʽ�Է��ý��н���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   405
      Index           =   1
      Left            =   1260
      TabIndex        =   1
      Top             =   750
      Width           =   4845
   End
   Begin VB.Label lbl����ͳ�������׼ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����ͳ�������׼"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   5
      Top             =   1770
      Width           =   1440
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "frm�������㷽ʽ.frx":000C
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    ����ǵ�һ��ʹ�û�ҽԺ�ĵ��������ݷ����仯����ʹ�����ع��ܣ��������������������ص����ء�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   555
      Index           =   0
      Left            =   1260
      TabIndex        =   0
      Top             =   150
      Width           =   4845
   End
   Begin VB.Label lbl������Ϣ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������(&J)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1080
      TabIndex        =   2
      Top             =   1380
      Width           =   810
   End
End
Attribute VB_Name = "frm�������㷽ʽ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const madLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Private Const madDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Private Const madDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����

Private mblnOK As Boolean
Private mint���� As Integer
Private mint������� As Integer
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mstr���� As String
Private mstrҽ���� As String
Private mstr�����ı�� As String
Private mstr���� As String
Private mbln���� As Boolean
Private mrs���� As New ADODB.Recordset

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

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDown_Click()
    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", mstrҽ�����ı���_����)
    If Not CommServer("GETHOSPSINGLEILLNESS") Then Exit Sub
    MsgBox "���سɹ���", vbInformation, gstrSysName
End Sub

Private Sub cmdOK_Click()
    Dim rsTemp As New ADODB.Recordset
    If txt������Ϣ.Tag = "" And txt������Ϣ <> "()����������" Then
        MsgBox "��ѡ��һ�������֣�", vbInformation, gstrSysName
        txt������Ϣ.SetFocus
        Exit Sub
    End If
    
    On Error Resume Next
    If mlng��ҳID <> 0 Then
        gstrSQL = " Select NVL(������㷽ʽ,'00') AS ������㷽ʽ From ҽ������סԺ��Ϣ Where ����ID=[1] And ��ҳID=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ������޸����㷽ʽ", mlng����ID, mlng��ҳID)
        If rsTemp.RecordCount <> 0 Then
            If Err = 0 Then
                If Mid(rsTemp!������㷽ʽ, 2, 1) <> "0" Then
                    MsgBox "ҽ���������ƣ��������޸ĸò��˵����㷽ʽ��", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
    End If
    
    '��ѡ������㷽ʽ�ϴ���ҽ������
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", mstrҽ����)
    Call InsertChild(mdomInput.documentElement, "RECKONINGTYPE", txt����ͳ�������׼.Tag)
    Call InsertChild(mdomInput.documentElement, "SINGLEILLNESSCODE", txt������Ϣ.Tag)
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName) ' ����Ա
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")) ' ��������
    If CommServer("SETRECKONINGTYPE") = False Then Exit Sub
    
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & mint���� & ",'������','''" & txt������Ϣ.Tag & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���浥���ֱ���")
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & mint���� & ",'���㷽ʽ','''" & txt����ͳ�������׼.Tag & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���浥���ֱ���")
    
    If mlng��ҳID <> 0 Then
        '�������õ����㷽ʽ
        gstrSQL = "ZL_ҽ������סԺ��Ϣ_INSERT(" & _
                  mlng����ID & "," & mlng��ҳID & ",'" & gstrUserName & "',2," & mint������� & ",'" & Split(txt������Ϣ.Text, ")")(1) & "',NULL,NULL," & _
                  "NULL,NULL,NULL,NULL,NULL,'" & txt������Ϣ.Tag & "','" & txt����ͳ�������׼.Text & "','" & txt����ͳ��ֵ�����.Text & "','" & _
                  txt���������׼.Text & "','" & txt�����ֵ�����.Text & "','" & txt����ͳ�������׼.Tag & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���浥���ֱ���")
        '�̳ظ� 2011-05-09 ��¼�������㷽ʽ����־������
        gstrSQL = "zl_������Ϣ��־_INSERT(" & mlng����ID & "," & mlng��ҳID & ",'" & txt������Ϣ.Text & "','" & txt����ͳ�������׼.Tag & "','" & UserInfo.���� & "',sysdate)"
        gcnGYYB.Execute gstrSQL
    End If
    MsgBox "���㷽ʽ���óɹ���", vbInformation, gstrSysName
    
    mblnOK = True
    Unload Me
End Sub

Private Sub cmd������Ϣ_Click()
    Dim blnReturn As Boolean
    blnReturn = frmListSel.ShowSelect(mint����, mrs����, "ID", "������ѡ��", "��ѡ�񵥲��֣�")
    If Not blnReturn Then mrs����.Filter = 0: Exit Sub
    
    txt������Ϣ.Text = "(" & mrs����!���� & ")" & mrs����!����
    txt������Ϣ.Tag = mrs����!����
    txt����ͳ�������׼.Tag = mrs����!���㷽ʽ
    txt����ͳ�������׼.Text = Nvl(mrs����!����ͳ�������׼)
    txt����ͳ��ֵ�����.Text = Nvl(mrs����!����ͳ��ֵ�����)
    txt�����ֵ�����.Text = Nvl(mrs����!�����ֵ�����)
    txt���������׼.Text = Nvl(mrs����!���������׼)
    mrs����.Filter = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    '��ȡ�ò��˵�ҽ����Ϣ
    gstrSQL = "Select �������,������ From �����ʻ� Where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ò��˵�ҽ����Ϣ", mlng����ID, mint����)
    txt������Ϣ.Text = Nvl(rsTemp!������)
    mint������� = rsTemp!�������
    mbln���� = (rsTemp!������� = "6")
    
    Call Get��֤_����(1, mstr����, mstrҽ����, mstr�����ı��, mstr����, mlng����ID)
    
    Call ��ȡ������
    Call ��ʾ������Ϣ
End Sub

Public Function ShowSelect(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal int���� As Integer, ByVal frmParent As Object) As Boolean
    mblnOK = False
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mint���� = int����
    Me.Show 1, frmParent
    ShowSelect = mblnOK
End Function

Private Function ��ȡ������() As Boolean
    Dim strFields As String, strValues As String
    Dim str���� As String, str���� As String, str���� As String, str���㷽ʽ As String
    Dim str����ͳ���嵥��׼ As String, str����ͳ��ֵ����� As String
    Dim str���������׼ As String, str�����ֵ����� As String
    
    Dim str��ǰ���� As String, str��ʼ���� As String, str�������� As String
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    
    Set mrs���� = New ADODB.Recordset
    strFields = "ID," & adVarChar & ",30|" & _
                "����," & adLongVarChar & ",30|" & _
                "����," & adLongVarChar & ",200|" & _
                "����," & adLongVarChar & ",30|" & _
                "���㷽ʽ," & adLongVarChar & ",10|" & _
                "����ͳ�������׼," & adLongVarChar & ",500|" & _
                "����ͳ��ֵ�����," & adLongVarChar & ",500|" & _
                "���������׼," & adLongVarChar & ",500|" & _
                "�����ֵ�����," & adLongVarChar & ",500"
    Call Record_Init(mrs����, strFields)
    
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "MITYPE", IIf(mbln����, "2", "1"))
    If CommServer("QUERYHOSPSINGLEILLNESS") = False Then Exit Function
    
    Set nodRowset = mdomOutput.documentElement.selectSingleNode("ROWSET")
    If nodRowset Is Nothing Then Exit Function
    '���ݱ���õ���������
    str��ǰ���� = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
    strFields = "ID|����|����|����|���㷽ʽ|����ͳ�������׼|����ͳ��ֵ�����|���������׼|�����ֵ�����"
    
    '�̶����ӿ����߽���
    str���� = ""
    str���� = "����������"
    str���㷽ʽ = 1
    str����ͳ���嵥��׼ = 0
    str����ͳ��ֵ����� = 0
    str���������׼ = 0
    str�����ֵ����� = 0
    str��ʼ���� = ""
    str�������� = ""
    str���� = zlCommFun.SpellCode(str����)
    strValues = str���� & "|" & str���� & "|" & str���� & "|" & str���� & "|" & str���㷽ʽ & "|" & str����ͳ���嵥��׼ & "|" & str����ͳ��ֵ����� & "|" & str���������׼ & "|" & str�����ֵ�����
    Call Record_Add(mrs����, strFields, strValues)

    For Each nodRow In nodRowset.childNodes
        str���� = GetAttributeValue(nodRow, "SINGLEILLNESSCODE")
        str���� = GetAttributeValue(nodRow, "SINGLEILLNESSNAME")
        str���㷽ʽ = GetAttributeValue(nodRow, "RECKONINGTYPE")
        str����ͳ���嵥��׼ = GetAttributeValue(nodRow, "PAYSTD")
        str����ͳ��ֵ����� = GetAttributeValue(nodRow, "PAYRATE")
        str���������׼ = GetAttributeValue(nodRow, "PAY2STD")
        str�����ֵ����� = GetAttributeValue(nodRow, "PAY2RATE")
        str��ʼ���� = Mid(GetAttributeValue(nodRow, "STARTDATE"), 1, 10)
        str�������� = Mid(GetAttributeValue(nodRow, "ENDDATE"), 1, 10)
        str���� = zlCommFun.SpellCode(str����)
        If str���� <> "" And str��ǰ���� >= str��ʼ���� And str��ǰ���� <= str�������� Then
            strValues = str���� & "|" & str���� & "|" & str���� & "|" & str���� & "|" & str���㷽ʽ & "|" & _
                str����ͳ���嵥��׼ & "|" & str����ͳ��ֵ����� & "|" & str���������׼ & "|" & str�����ֵ�����
            Call Record_Add(mrs����, strFields, strValues)
        End If
    Next
    ��ȡ������ = True
End Function

Private Function ��ʾ������Ϣ(Optional ByVal bln����ƥ�� As Boolean = False) As Boolean
    Dim blnReturn As Boolean
    Dim StrInput As String, strFilter As String
    
    If Trim(txt������Ϣ.Text) = "" Then Exit Function
    If InStr(1, txt������Ϣ.Text, "(") <> 0 Then
        If InStr(1, txt������Ϣ.Text, ")") <> 0 Then
            StrInput = Mid(txt������Ϣ.Text, 2, InStr(1, txt������Ϣ.Text, ")") - 2)
        Else
            StrInput = Mid(txt������Ϣ.Text, 2, Len(txt������Ϣ.Text) - 1)
        End If
    Else
        StrInput = txt������Ϣ.Text
    End If
    'bln����ƥ��:�����������ƥ�䣬�����Ǵ����ݿ�����ϴ���ѡ��Ĳ��֣���˲�ȡ����ƥ�䣬���б���������Ƶģ�������ͨ���������鲡��ʱ��Ҫ����ƥ��
    If bln����ƥ�� Then
        StrInput = UCase("'" & StrInput & "*'")
        strFilter = "���� Like " & StrInput & " Or ���� Like " & StrInput & " Or ���� Like " & StrInput
    Else
        StrInput = UCase("'" & StrInput & "'")
        strFilter = "����=" & StrInput
    End If
    
    With mrs����
        .Filter = strFilter
        If .RecordCount = 0 Then
            If bln����ƥ�� Then
                MsgBox "û���ҵ�ָ���ĵ����֣�[���ֱ���Ϊ:" & UCase(txt������Ϣ.Text) & "]", vbInformation, gstrSysName
            End If
            Call zlControl.TxtSelAll(txt������Ϣ)
            txt������Ϣ.Text = ""
            txt������Ϣ.Tag = ""
            txt����ͳ�������׼.Text = ""
            txt����ͳ�������׼.Tag = 1
            txt����ͳ��ֵ�����.Text = ""
            txt�����ֵ�����.Text = ""
            txt���������׼.Text = ""
            .Filter = 0
            Exit Function
        Else
            If mrs����.RecordCount > 1 Then
                blnReturn = frmListSel.ShowSelect(mint����, mrs����, "ID", "������ѡ��", "��ѡ�񵥲��֣�")
            Else
                blnReturn = True
            End If
            If blnReturn = False Then
                txt������Ϣ.Text = ""
                txt������Ϣ.Tag = ""
                txt����ͳ�������׼.Text = ""
                txt����ͳ��ֵ�����.Text = ""
                txt�����ֵ�����.Text = ""
                txt���������׼.Text = ""
                txt����ͳ�������׼.Tag = 1
                Call zlControl.TxtSelAll(txt������Ϣ)
            Else
                txt������Ϣ.Text = "(" & mrs����!���� & ")" & mrs����!����
                txt������Ϣ.Tag = mrs����!����
                txt����ͳ�������׼.Tag = mrs����!���㷽ʽ
                txt����ͳ�������׼.Text = Nvl(mrs����!����ͳ�������׼)
                txt����ͳ��ֵ�����.Text = Nvl(mrs����!����ͳ��ֵ�����)
                txt�����ֵ�����.Text = Nvl(mrs����!�����ֵ�����)
                txt���������׼.Text = Nvl(mrs����!���������׼)
                ��ʾ������Ϣ = True
            End If
        End If
        .Filter = 0
    End With
End Function

Private Sub txt������Ϣ_GotFocus()
    Call zlControl.TxtSelAll(txt������Ϣ)
End Sub

Private Sub txt������Ϣ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txt������Ϣ.Text) = "" Then Exit Sub
    
    If Not ��ʾ������Ϣ(True) Then Exit Sub
    Call zlCommFun.PressKey(vbKeyTab)
End Sub
