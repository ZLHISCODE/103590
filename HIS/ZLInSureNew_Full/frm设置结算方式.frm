VERSION 5.00
Begin VB.Form frm���ý��㷽ʽ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ý��㷽ʽ"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   Icon            =   "frm���ý��㷽ʽ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txt�����޶� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1890
      TabIndex        =   7
      Top             =   2340
      Width           =   3885
   End
   Begin VB.TextBox txtͳ����ɱ�׼ 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1890
      TabIndex        =   9
      Top             =   2730
      Width           =   3885
   End
   Begin VB.TextBox txt���˸�����׼ 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1890
      TabIndex        =   11
      Top             =   3120
      Width           =   3885
   End
   Begin VB.ComboBox cbo���㷽ʽ 
      Height          =   300
      Left            =   1890
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1560
      Width           =   3915
   End
   Begin VB.CommandButton cmd������Ϣ 
      Caption         =   "��"
      Enabled         =   0   'False
      Height          =   300
      Left            =   5520
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1950
      Width           =   285
   End
   Begin VB.TextBox txt������Ϣ 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1890
      TabIndex        =   4
      Top             =   1950
      Width           =   3615
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4590
      TabIndex        =   15
      Top             =   3780
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3330
      TabIndex        =   14
      Top             =   3780
      Width           =   1100
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "����(&W)"
      Height          =   350
      Left            =   240
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3780
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -30
      TabIndex        =   12
      Top             =   3600
      Width           =   6165
   End
   Begin VB.Label lbl�����޶� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�����޶�"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1080
      TabIndex        =   6
      Top             =   2400
      Width           =   720
   End
   Begin VB.Label lblͳ����ɱ�׼ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ͳ����ɱ�׼"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   720
      TabIndex        =   8
      Top             =   2790
      Width           =   1080
   End
   Begin VB.Label lbl���˸�����׼ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���˸�����׼"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   720
      TabIndex        =   10
      Top             =   3180
      Width           =   1080
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    ����ǵ�һ��ʹ�û�ҽԺ�ĵ����ֽ���Ŀ¼�����仯����ʹ�����ع��ܣ��������ֽ���Ŀ¼���ص����ء�"
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
      Height          =   375
      Index           =   1
      Left            =   1170
      TabIndex        =   16
      Top             =   150
      Width           =   4845
   End
   Begin VB.Label lbl���㷽ʽ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���㷽ʽ(&B)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   810
      TabIndex        =   1
      Top             =   1620
      Width           =   990
   End
   Begin VB.Label lbl������Ϣ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������(&J)"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   990
      TabIndex        =   3
      Top             =   2010
      Width           =   810
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frm���ý��㷽ʽ.frx":000C
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
      Height          =   765
      Index           =   0
      Left            =   1170
      TabIndex        =   0
      Top             =   630
      Width           =   4845
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   300
      Picture         =   "frm���ý��㷽ʽ.frx":00B2
      Top             =   300
      Width           =   480
   End
End
Attribute VB_Name = "frm���ý��㷽ʽ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const madLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Private Const madDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Private Const madDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����

Private mstrReturn As String
Private mint������� As Integer
Private mint���� As Integer
Private mlng����ID As Long
Private mlng��ҳID As Long                 '���ﲻ���ú������£�Ҳ�����±����ʻ��е�����
Private mstr���� As String
Private mstrҽ���� As String
Private mstr�����ı�� As String
Private mstr���� As String
Private mbln����  As Boolean
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

Private Sub cbo���㷽ʽ_Click()
    Dim blnEnable As Boolean
    blnEnable = (Me.cbo���㷽ʽ.ListIndex = 1)
    lbl������Ϣ.Enabled = blnEnable
    txt������Ϣ.Enabled = blnEnable
    cmd������Ϣ.Enabled = blnEnable
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDown_Click()
    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", mstrҽ�����ı���_����)
    If Not CommServer("GETHOSPSINGLEILLNESS_BG") Then Exit Sub
    MsgBox "���سɹ���", vbInformation, gstrSysName
End Sub

Private Sub cmdOK_Click()
    Dim rsTemp As New ADODB.Recordset
    If Me.cbo���㷽ʽ.ListIndex = 1 Then
        If txt������Ϣ.Tag = "" Then
            MsgBox "��ѡ��һ�������֣�", vbInformation, gstrSysName
            txt������Ϣ.SetFocus
            Exit Sub
        End If
    End If
    
    On Error Resume Next
    If mlng��ҳID <> 0 Then
        gstrSQL = " Select NVL(������㷽ʽ,'00') AS ������㷽ʽ From ҽ������סԺ��Ϣ Where ����ID=[1] And ��ҳID=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ������޸����㷽ʽ", mlng����ID, mlng��ҳID)
        If rsTemp.RecordCount <> 0 Then
            If Err = 0 Then
                If Mid(rsTemp!������㷽ʽ, 2, 1) <> "0" Then
                    MsgBox "ҽ���������ƣ��������޸ĸò��˵Ľ��㷽ʽ��", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
    End If
    
    '��ѡ������㷽ʽ�ϴ���ҽ������
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", mstrҽ����)
    Call InsertChild(mdomInput.documentElement, "CALTYPE", Me.cbo���㷽ʽ.ListIndex)
    Call InsertChild(mdomInput.documentElement, "SINGLEILLNESSCODE", txt������Ϣ.Tag)
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName) ' ����Ա
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")) ' ��������
    If mlng��ҳID <> 0 Then If CommServer("SETCALTYPE") = False Then Exit Sub
    
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & mint���� & ",'�����ֱ���_����','''" & txt������Ϣ.Tag & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "���浥������Ŀ")
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & mint���� & ",'���㷽ʽ','''" & Me.cbo���㷽ʽ.ListIndex & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������㷽ʽ")
    If mlng��ҳID <> 0 Then
        gstrSQL = "ZL_ҽ������סԺ��Ϣ_INSERT(" & _
            mlng����ID & "," & mlng��ҳID & ",'" & gstrUserName & "',1," & mint������� & ",'" & Split(txt������Ϣ.Text, ")")(1) & "',NULL,NULL,'" & _
            txt������Ϣ.Tag & "'," & txt�����޶�.Text & "," & txtͳ����ɱ�׼.Text & "," & txt���˸�����׼.Text & ",'" & Me.cbo���㷽ʽ.Text & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "������㷽ʽ")
        MsgBox "���㷽ʽ���óɹ���", vbInformation, gstrSysName
    End If
    
    mstrReturn = Me.cbo���㷽ʽ.ListIndex & ";" & Me.txt������Ϣ.Tag
    Unload Me
End Sub

Private Sub cmd������Ϣ_Click()
    Dim blnReturn As Boolean
    blnReturn = frmListSel.ShowSelect(mint����, mrs����, "ID", "������ѡ��", "��ѡ�񵥲��֣�")
    If Not blnReturn Then mrs����.Filter = 0: Exit Sub
    
    txt������Ϣ.Text = "(" & mrs����!���� & ")" & mrs����!����
    txt������Ϣ.Tag = mrs����!����
    txt�����޶�.Text = Nvl(mrs����!�����޶�)
    txtͳ����ɱ�׼.Text = Nvl(mrs����!ͳ����ɱ�׼)
    txt���˸�����׼.Text = Nvl(mrs����!���˸�����׼)
    mrs����.Filter = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    mstrReturn = ""
    
    Me.cbo���㷽ʽ.Clear
    Me.cbo���㷽ʽ.AddItem "����Ŀ����"
    Me.cbo���㷽ʽ.AddItem "�����ְ��ɽ���"
    Me.cbo���㷽ʽ.ListIndex = 0
    
    '��ȡ�ò��˵�ҽ����Ϣ
    gstrSQL = "Select �������,�����ֱ���_���� From �����ʻ� Where ����ID=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ò��˵�ҽ����Ϣ", mlng����ID, mint����)
    If mlng��ҳID <> 0 Then txt������Ϣ.Text = Nvl(rsTemp!�����ֱ���_����)
    mint������� = rsTemp!�������
    mbln���� = (rsTemp!������� = "6")
    
    Call Get��֤_����(1, mstr����, mstrҽ����, mstr�����ı��, mstr����, mlng����ID)
    
    Call ��ȡ������
    Call ��ʾ������Ϣ
    If txt������Ϣ.Text <> "" Then Me.cbo���㷽ʽ.ListIndex = 1
End Sub

Public Function ShowSelect(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal int���� As Integer, ByVal frmParent As Object) As String
    mlng����ID = lng����ID
    mint���� = int����
    mlng��ҳID = lng��ҳID
    If frmParent Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmParent
    End If
    ShowSelect = mstrReturn
End Function

Private Function ��ȡ������() As Boolean
    Dim strFields As String, strValues As String
    Dim str���� As String, str���� As String
    Dim dbl�����޶� As Double, dblͳ����ɱ�׼ As Double, dbl���˸�����׼ As Double
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    
    Dim str��ǰ���� As String, str��ʼ���� As String, str�������� As String
    Set mrs���� = New ADODB.Recordset
    strFields = "ID," & adVarChar & ",30|" & _
                "����," & adLongVarChar & ",30|" & _
                "����," & adLongVarChar & ",200|" & _
                "�����޶�," & adDouble & ",30|" & _
                "ͳ����ɱ�׼," & adDouble & ",30|" & _
                "���˸�����׼," & adDouble & ",30"
    Call Record_Init(mrs����, strFields)
    
    '��XML DomDocument������г�ʼ��
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "MITYPE", IIf(mbln����, "2", "1"))
    If CommServer("QUERYHOSPSINGLEILLNESS_BG") = False Then Exit Function
    
    Set nodRowset = mdomOutput.documentElement.selectSingleNode("ROWSET")
    If nodRowset Is Nothing Then Exit Function
    str��ǰ���� = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
    
    '���ݱ���õ���������
    strFields = "ID|����|����|�����޶�|ͳ����ɱ�׼|���˸�����׼"
    For Each nodRow In nodRowset.childNodes
        str���� = GetAttributeValue(nodRow, "SINGLEILLNESSCODE")
        str���� = GetAttributeValue(nodRow, "SINGLEILLNESSNAME")
        dbl�����޶� = Val(GetAttributeValue(nodRow, "PAYLMT"))
        dblͳ����ɱ�׼ = Val(GetAttributeValue(nodRow, "FUNDSTD"))
        dbl���˸�����׼ = Val(GetAttributeValue(nodRow, "PSNSTD"))
        str��ʼ���� = Mid(GetAttributeValue(nodRow, "STARTDATE"), 1, 10)
        str�������� = Mid(GetAttributeValue(nodRow, "ENDDATE"), 1, 10)
        If str���� <> "" And str��ǰ���� >= str��ʼ���� And str��ǰ���� <= str�������� Then
            strValues = str���� & "|" & str���� & "|" & str���� & "|" & dbl�����޶� & "|" & dblͳ����ɱ�׼ & "|" & dbl���˸�����׼
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
        strFilter = "���� Like " & StrInput & " Or ���� Like " & StrInput
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
            txt�����޶�.Text = ""
            txtͳ����ɱ�׼.Text = ""
            txt���˸�����׼.Text = ""
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
                txt�����޶�.Text = ""
                txtͳ����ɱ�׼.Text = ""
                txt���˸�����׼.Text = ""
                Call zlControl.TxtSelAll(txt������Ϣ)
            Else
                txt������Ϣ.Text = "(" & mrs����!���� & ")" & mrs����!����
                txt������Ϣ.Tag = mrs����!����
                txt�����޶�.Text = Nvl(mrs����!�����޶�)
                txtͳ����ɱ�׼.Text = Nvl(mrs����!ͳ����ɱ�׼)
                txt���˸�����׼.Text = Nvl(mrs����!���˸�����׼)
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
End Sub
