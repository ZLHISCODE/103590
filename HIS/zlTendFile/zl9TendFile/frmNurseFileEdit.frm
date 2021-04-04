VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmNurseFileEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ļ��༭"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4545
   Icon            =   "frmNurseFileEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCanCel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2640
      TabIndex        =   8
      Top             =   1920
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1410
      TabIndex        =   7
      Top             =   1920
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -120
      TabIndex        =   6
      Top             =   1710
      Width           =   4545
   End
   Begin MSMask.MaskEdBox mskEdit 
      Height          =   285
      Left            =   1230
      TabIndex        =   5
      Top             =   1110
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   19
      Format          =   "yyyy-MM-dd HH:mm:ss"
      Mask            =   "####-##-## ##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txt�ļ����� 
      Height          =   285
      Left            =   1230
      MaxLength       =   50
      TabIndex        =   3
      Top             =   720
      Width           =   2895
   End
   Begin VB.ComboBox cbo��ʽ��Դ 
      Height          =   300
      Left            =   1230
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   330
      Width           =   2895
   End
   Begin VB.Label lbl��ʼʱ�� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ʼʱ��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   4
      Top             =   1155
      Width           =   720
   End
   Begin VB.Label lbl�ļ����� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�ļ�����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   2
      Top             =   765
      Width           =   720
   End
   Begin VB.Label lbl��ʽ��Դ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ʽ��Դ"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   0
      Top             =   390
      Width           =   720
   End
End
Attribute VB_Name = "frmNurseFileEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintShowTime As Integer '���µ��ļ���ȱʡ��ʼʱ��:1-���ʱ��;0-��Ժʱ��
Private mstr��Ժʱ�� As String
Private mstr���ʱ�� As String
Private mstr��Ժʱ�� As String

Private mlngFile As Long        '�ļ�ID,����0��ʾ����,�����ʾ�޸�(�޸�ʱ�������޸��ļ���Դ)
Private mlngFormat As Long
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mintӤ�� As Long
Private mlng����ID As Long
Private mstrDept As String      '��ǰ����
Private mblnOK As Boolean       '�Ƿ񱣴�ɹ�
Private mblnExist���µ� As Boolean
Private mblnExist��¼�� As Boolean
Private mblnExist����ͼ As Boolean
Private mIntPartogramID As Boolean
Private mstrCurForamt As String '�����������µ���ʽID���(���ļ���ʼʱ������),��ʽ:30,40
Private mblnOnly As Boolean     'סԺ����ͬһʱ��ֻ��¼һ�ݻ����ļ�

Public Function ShowEditor(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intӤ�� As Integer, ByVal lng����ID As Long, _
    ByVal str���� As String, Optional lngFile As Long = 0, Optional lngFormat As Long = 0) As Boolean
    mblnOK = False
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mintӤ�� = intӤ��
    mlng����ID = lng����ID
    mlngFile = lngFile
    mlngFormat = lngFormat
    mstrDept = str����
    mIntPartogramID = -1
    Me.Show 1
    lngFile = mlngFile
    lngFormat = mlngFormat            '���ظ�ʽID,���ڶ�λ����ͬ��ʽ���ļ�
    ShowEditor = mblnOK
End Function

Private Sub cbo��ʽ��Դ_Click()
    Dim bln���µ� As Boolean
    Dim bln����ͼ As Boolean
    Dim strDate As String
    
    txt�ļ�����.Text = Split(cbo��ʽ��Դ.Text, "-")(1)
    If InStr(1, "," & cbo��ʽ��Դ.Tag & ",", "," & cbo��ʽ��Դ.ListIndex & ",") = 0 Or cbo��ʽ��Դ.Tag = "" Then
        If mIntPartogramID <> Me.cbo��ʽ��Դ.ItemData(Me.cbo��ʽ��Դ.ListIndex) Then
            txt�ļ�����.Text = "[" & mstrDept & "]" & txt�ļ�����.Text '��¼��ʱ���´���
        Else
            '����ͼ
            bln����ͼ = True
        End If
    Else
        '���µ�:Ŀǰ������Ӷ�����µ�
        txt�ļ�����.Text = "[" & mstrDept & "]" & txt�ļ�����.Text
    End If
    
    '����:������滤���ļ���ȱʡʱ��Ϊ��Ժʱ��,����Ϊ��ǰʱ��
    '�޸�:�����ļ��Ŀ�ʼʱ�䲻��С����Ժʱ��,���ܴ������ݷ���ʱ��,�����������ܴ��ڵ�ǰʱ��
    bln���µ� = (cbo��ʽ��Դ.Tag <> "" And InStr(1, "," & cbo��ʽ��Դ.Tag & ",", "," & cbo��ʽ��Դ.ListIndex & ",") > 0)
    If mlngFile = 0 Then
        If (Not mblnExist��¼�� And Not bln���µ�) Or (Not mblnExist���µ� And bln���µ�) Or (Not mblnExist����ͼ And bln����ͼ) Then
            '������������ʾ���ʱ��,������ʾ��Ժʱ��
            mskEdit.Text = Format(IIf(mstr���ʱ�� = "", mstr��Ժʱ��, mstr���ʱ��), "YYYY-MM-DD HH:mm:ss")
        Else
            If bln���µ� Then
                strDate = GetCreateWaveDate
            Else
                strDate = Format(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm") & ":00", "YYYY-MM-DD HH:mm:ss")
            End If
            If mstr��Ժʱ�� <> "" And Format(strDate, "YYYY-MM-DD HH:mm:ss") > Format(mstr��Ժʱ��, "YYYY-MM-DD HH:mm:ss") Then
                strDate = Format(mstr��Ժʱ��, "YYYY-MM-DD") & " 00:00:00"
            End If
            mskEdit.Text = strDate
        End If
        mskEdit.Tag = mskEdit.Text
    End If
    '���ѡ�е��ǵ�һ�����µ��Ļ�
    If IsFirstCurve Then
        mskEdit.Text = Format(IIf(mintShowTime = 1 And mstr���ʱ�� <> "", mstr���ʱ��, mstr��Ժʱ��), "YYYY-MM-DD HH:mm:ss")
        '56627:�ſ�Ӥ�����µ������޸ĵ����ƣ�ֻҪʱ�䲻С��Ӥ������ʱ�伴�ɡ�
        'mskEdit.Enabled = mintӤ�� = 0
    Else
        mskEdit.Enabled = True
    End If
End Sub

Private Function GetCreateWaveDate() As String
'-----------------------------------------------------------
'���ܣ��½����µ�ʱ��ȡ�ļ���ʼʱ��
'���ת�����ʱ��Ҫ���ڲ���֮ǰ���µ��ļ���������ݷ���ʱ���ʼʱ�䣬
'����ת�ƺ��½����µ�ʱ��Ϊת�����ʱ��,����Ϊ��ǰϵͳʱ��
'���أ��������µ���ʱ��
'-----------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strDate As String
    
    On Error GoTo ErrHand
    gstrSQL = _
        " SELECT E.��ʼʱ��" & vbNewLine & _
        " FROM ���˱䶯��¼ e" & vbNewLine & _
        " WHERE e.����id = [1] AND e.��ҳid = [2] AND Nvl(e.���Ӵ�λ, 0) = 0 AND e.��ʼʱ�� IS NOT NULL AND e.��ʼԭ�� IN (3, 15) AND" & vbNewLine & _
        "      e.��ֹʱ�� IS NULL AND e.��ʼʱ�� > (SELECT Nvl(MAX(b.����ʱ��), MAX(a.��ʼʱ��))" & vbNewLine & _
        "                                       FROM �����ļ��б� c, ���˻����ļ� a, ���˻������� b" & vbNewLine & _
        "                                       WHERE a.����id = [1] AND a.��ҳid = [2] AND Nvl(Ӥ��, 0) = [3] AND a.Id = b.�ļ�id(+) AND" & vbNewLine & _
        "                                             c.Id = a.��ʽid AND c.���� = 3 AND c.���� = -1)"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ļ�ʱ��", mlng����ID, mlng��ҳID, mintӤ��)
    If rsTemp.RecordCount > 0 Then
        strDate = Format(rsTemp!��ʼʱ��, "YYYY-MM-DD HH:mm:ss")
    Else
        strDate = Format(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm") & ":00", "YYYY-MM-DD HH:mm:ss")
    End If
    GetCreateWaveDate = strDate
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function IsFirstCurve() As Boolean
'����:���Ŀǰѡ�е����µ��Ƿ��ǵ�һ�����µ�
    Dim arrCode() As String
    Dim blnIsSelectCurve As Boolean
    
    blnIsSelectCurve = (cbo��ʽ��Դ.Tag <> "" And InStr(1, "," & cbo��ʽ��Դ.Tag & ",", "," & cbo��ʽ��Դ.ListIndex & ",") > 0)
    If blnIsSelectCurve = False Then IsFirstCurve = False: Exit Function
    If mblnExist���µ� = False Then IsFirstCurve = True: Exit Function
    arrCode = Split(mstrCurForamt, ",")
    If Val(arrCode(0)) = mlngFile And mlngFile <> 0 Then
        IsFirstCurve = True
    Else
        IsFirstCurve = False
    End If
End Function

Private Sub cmdCanCel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngID As Long
    Dim strDate As String, strTime As String
    Dim lngUpFileID As Long
    Dim strCurDate As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    If txt�ļ�����.Text = "" Then
        MsgBox "�������ļ����ƣ�", vbInformation, gstrSysName
        If txt�ļ�����.Enabled And txt�ļ�����.Visible Then txt�ļ�����.SetFocus
        Exit Sub
    End If
    If LenB(StrConv(txt�ļ�����.Text, vbFromUnicode)) > 50 Then
        MsgBox "�ļ����Ƴ����������50���ַ���25�����֣�", vbInformation, gstrSysName
        If txt�ļ�����.Enabled And txt�ļ�����.Visible Then txt�ļ�����.SetFocus
        Exit Sub
    End If
    
    strCurDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
    strDate = Format(mskEdit.Text, "YYYY-MM-DD HH:mm:ss")
    If Not IsDate(strDate) Then
        MsgBox "�ļ���ʼʱ���ʽ���ԣ��磺2011-4-13 23:59:00�������������룡", vbInformation, gstrSysName
        If IsDate(Format(mskEdit.Tag, "YYYY-MM-DD HH:mm:ss")) Then mskEdit.Text = mskEdit.Tag
        If mskEdit.Enabled And mskEdit.Visible Then mskEdit.SetFocus
        Exit Sub
    End If
    '56627
    '�����ļ����ڲ��ܴ��ڵ�ǰ���ڻ��Ժʱ��
    If mstr��Ժʱ�� = "" Then
        If strDate > strCurDate Then
            MsgBox "�ļ���ʼʱ�䲻�ܴ��ڵ�ǰʱ��[" & strCurDate & "]�����������룡", vbInformation, gstrSysName
            If IsDate(Format(mskEdit.Tag, "YYYY-MM-DD HH:mm:ss")) Then mskEdit.Text = mskEdit.Tag
            If mskEdit.Enabled And mskEdit.Visible Then mskEdit.SetFocus
            Exit Sub
        End If
    Else
        If strDate > Format(mstr��Ժʱ��, "YYYY-MM-DD HH:mm:ss") Then
            MsgBox "�ļ���ʼʱ�䲻�ܴ��ڳ�Ժʱ��[" & mstr��Ժʱ�� & "]�����������룡", vbInformation, gstrSysName
            If IsDate(Format(mskEdit.Tag, "YYYY-MM-DD HH:mm:ss")) Then mskEdit.Text = mskEdit.Tag
            If mskEdit.Enabled And mskEdit.Visible Then mskEdit.SetFocus
            Exit Sub
        End If
    End If
    '56627:����Ӥ����һ�����µ������µ���ʼʱ�䲻��С��Ӥ������ʱ��
    If IsFirstCurve And mintӤ�� <> 0 Then
        If Format(strDate, "YYYY-MM-DD HH:mm:ss") < Format(mstr���ʱ��, "YYYY-MM-DD HH:mm:ss") Then
            MsgBox "Ӥ���ļ��Ŀ�ʼʱ�䲻��С�ڳ���ʱ��[" & Format(mstr���ʱ��, "YYYY-MM-DD HH:mm:ss") & "],���������룡", vbInformation, gstrSysName
            If IsDate(Format(mskEdit.Tag, "YYYY-MM-DD HH:mm:ss")) Then mskEdit.Text = mskEdit.Tag
            If mskEdit.Enabled And mskEdit.Visible Then mskEdit.SetFocus
            Exit Sub
        End If
    End If
    '�������µ�
    If cbo��ʽ��Դ.Tag <> "" And InStr(1, "," & cbo��ʽ��Դ.Tag & ",", "," & cbo��ʽ��Դ.ListIndex & ",") > 0 Then
        gstrSQL = _
            " SELECT A.ID,MAX(A.�ļ�����) �ļ�����,MAX(A.��ʼʱ��) ��ʼʱ��,NVL(MAX(C.����ʱ��),MAX(A.��ʼʱ��)) ����ʱ��" & vbNewLine & _
            " FROM ���˻����ļ� A, �����ļ��б� B, ���˻������� C" & vbNewLine & _
            " WHERE A.��ʽID = B.ID AND A.ID = C.�ļ�ID(+) AND A.����ID = [1] AND A.��ҳID = [2] AND A.Ӥ�� = [3] AND B.���� = -1" & vbNewLine & _
            " GROUP BY A.ID"
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ��Ѷ������µ�", mlng����ID, mlng��ҳID, mintӤ��)
        '������½��ļ�(�ļ��Ŀ�ʼʱ��һ��Ҫ������һ�ļ���ʼʱ������ݷ���ʱ��)
        '������޸��ļ�,�޸��ļ��Ŀ�ʼʱ��Ҫ������һ�ļ��ļ���ʼʱ������ݷ���ʱ��,ҪС����һ�ļ��Ŀ�ʼʱ��
        If mlngFile = 0 Or (mlngFile <> 0 And Not IsDate(Format(mskEdit.Tag, "YYYY-MM-DD HH:mm:ss"))) Then
            rsTemp.Filter = ""
            strTime = strDate
        Else
            strTime = Format(mskEdit.Tag, "YYYY-MM-DD HH:mm:ss")
            rsTemp.Filter = "��ʼʱ��< '" & strTime & "'"
        End If
        rsTemp.Sort = "��ʼʱ�� DESC"
        If rsTemp.RecordCount > 0 Then
            lngUpFileID = rsTemp!ID
            If CDate(strDate) <= CDate(rsTemp!����ʱ��) Then
                MsgBox "�ļ���ʼʱ��Ҫ������һ�ļ���" & NVL(rsTemp!�ļ�����) & "���Ŀ�ʼ�����ݷ���ʱ�䡾" & Format(rsTemp!����ʱ��, "YYYY-MM-DD HH:mm:ss") & "�������������룡", vbInformation, gstrSysName
                If IsDate(Format(mskEdit.Tag, "YYYY-MM-DD HH:mm:ss")) Then mskEdit.Text = mskEdit.Tag
                If mskEdit.Enabled And mskEdit.Visible Then mskEdit.SetFocus
                Exit Sub
            End If
        End If
        strTime = Format(mskEdit.Tag, "YYYY-MM-DD HH:mm:ss")
        If mlngFile <> 0 And IsDate(strTime) Then
            rsTemp.Filter = "��ʼʱ��> '" & strTime & "'"
            rsTemp.Sort = "��ʼʱ�� ASC"
            If rsTemp.RecordCount > 0 Then
                If CDate(strDate) >= CDate(rsTemp!��ʼʱ��) Then
                    MsgBox "�ļ���ʼʱ��ҪС����һ�ļ���" & NVL(rsTemp!�ļ�����) & "���Ŀ�ʼʱ�䡾" & Format(rsTemp!��ʼʱ��, "YYYY-MM-DD HH:mm:ss") & "�������������룡", vbInformation, gstrSysName
                    If IsDate(Format(mskEdit.Tag, "YYYY-MM-DD HH:mm:ss")) Then mskEdit.Text = mskEdit.Tag
                    If mskEdit.Enabled And mskEdit.Visible Then mskEdit.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If
    
    If mlngFile = 0 Then
        lngID = zlDatabase.GetNextId("���˻����ļ�")
    Else
        lngID = mlngFile
    End If
    
    gstrSQL = "ZL_���˻����ļ�_UPDATE(" & lngID & "," & mlng����ID & "," & mlng����ID & "," & mlng��ҳID & "," & mintӤ�� & "," & _
              cbo��ʽ��Դ.ItemData(cbo��ʽ��Դ.ListIndex) & ",'" & txt�ļ�����.Text & "',to_date('" & strDate & "','yyyy-MM-dd hh24:mi:ss')," & IIf(mlngFile = 0 And mblnOnly, 1, 0) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������")
    mlngFile = lngID
    mlngFormat = mlngFile 'cbo��ʽ��Դ.ItemData(cbo��ʽ��Դ.ListIndex)
    
    '�����ǰ�������Ǽ�¼���Ҳ��������µ����Զ�����һ�����µ�
    If Not mblnExist���µ� And Not (cbo��ʽ��Դ.Tag <> "" And InStr(1, "," & cbo��ʽ��Դ.Tag & ",", "," & cbo��ʽ��Դ.ListIndex & ",") > 0) Then
        lngID = zlDatabase.GetNextId("���˻����ļ�")
        gstrSQL = "ZL_���˻����ļ�_UPDATE(" & lngID & "," & mlng����ID & "," & mlng����ID & "," & mlng��ҳID & "," & mintӤ�� & "," & _
                  cbo��ʽ��Դ.ItemData(Val(cbo��ʽ��Դ.Tag)) & ",'�������±�',to_date('" & IIf(mintShowTime = 1 And mstr���ʱ�� <> "", mstr���ʱ��, mstr��Ժʱ��) & "','yyyy-MM-dd hh24:mi:ss'),0)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "��������")
    ElseIf mblnExist���µ� And (cbo��ʽ��Դ.Tag <> "" And InStr(1, "," & cbo��ʽ��Դ.Tag & ",", "," & cbo��ʽ��Դ.ListIndex & ",") > 0) Then
        '����һ���µ��ļ��Ľ���ʱ�����Ϊ���ļ��Ŀ�ʼʱ��-1S
        If lngUpFileID > 0 Then
            strDate = CDate(strDate) - (1 / 24 / 60 / 60)
            gstrSQL = "ZL_���˻����ļ�_STATE(" & lngUpFileID & ",1,To_Date('" & strDate & "','yyyy-MM-dd hh24:mi:ss'))"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "����ļ�����")
        End If
    End If
    
    mblnOK = True
    Unload Me
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim blnSeek As Boolean
    Dim lng��ʽ As Long, lng���� As Long, str�ļ����� As String, str��ʼʱ�� As String, str��ʽ As String
    Dim rsTemp As New ADODB.Recordset
    Dim intIndex As Integer

    On Error GoTo ErrHand
    
    mskEdit.Enabled = True
    mintShowTime = zlDatabase.GetPara("���µ��ļ���ʼʱ��", glngSys, 1255, 1)
    
    '��ȡ��ǰ������Ժʱ��
    gstrSQL = " Select A.��Ժ����,A.��Ժ����,B.���� From ������ҳ A,���ű� B" & _
        " Where A.����ID=[1] And A.��ҳID=[2] And A.��Ժ����ID=B.ID(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ǰ������Ժʱ��", mlng����ID, mlng��ҳID)
    mstr��Ժʱ�� = Format(rsTemp!��Ժ����, "yyyy-MM-dd HH:mm:ss")
    mstr��Ժʱ�� = Format(NVL(rsTemp!��Ժ����), "yyyy-MM-dd HH:mm:ss")
    If mstrDept = "" Then mstrDept = NVL(rsTemp!����)
    
    mstr���ʱ�� = ""
    gstrSQL = " Select ��ʼʱ�� From ���˱䶯��¼ Where ����ID=[1] And ��ҳID=[2] And ��ʼԭ��=2 Order by ��ʼʱ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ǰ�������ʱ��", mlng����ID, mlng��ҳID)
    If rsTemp.RecordCount <> 0 Then
        mstr���ʱ�� = Format(rsTemp!��ʼʱ��, "yyyy-MM-dd HH:mm:ss")
    End If
    
    '�����Ӥ������ȡӤ���Ǽ�ʱ��Ϊ�ļ���ʼʱ��
    If mintӤ�� <> 0 Then
        gstrSQL = "select ����ʱ�� from ������������¼ where ����ID=[1] And ��ҳID=[2] And ���=[3] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡӤ���Ǽ�ʱ��", mlng����ID, mlng��ҳID, mintӤ��)
        If rsTemp.RecordCount <> 0 Then
            mstr���ʱ�� = Format(NVL(rsTemp!����ʱ��, mstr���ʱ��), "yyyy-MM-dd HH:mm:ss")
            mstr��Ժʱ�� = mstr���ʱ��
        End If
        '��ȡӤ����Ժ����
        gstrSQL = "Select b.����id, b.��ҳid, b.Ӥ��, ��ʼִ��ʱ��" & vbNewLine & _
            " From ����ҽ����¼ b, ������ĿĿ¼ c" & vbNewLine & _
            " Where b.������Ŀid + 0 = c.Id And b.ҽ��״̬ = 8 And Nvl(b.Ӥ��, 0) <> 0 And c.��� = 'Z' And c.�������� In ('3', '5', '11') And" & vbNewLine & _
            "      b.����id = [1] And b.��ҳid = [2] And b.Ӥ�� = [3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡӤ���Ǽ�ʱ��", mlng����ID, mlng��ҳID, mintӤ��)
        If rsTemp.RecordCount <> 0 Then
            mstr��Ժʱ�� = Format(NVL(rsTemp!��ʼִ��ʱ��, mstr��Ժʱ��), "yyyy-MM-dd HH:mm:ss")
        End If
    End If
    '����Ƿ����趨���µ�,���Ѵ����������ٴ�������µ�
    gstrSQL = " Select B.����,���,A.ID,A.��ʼʱ��" & _
              " From ���˻����ļ� A,�����ļ��б� B" & _
              " Where A.��ʽID=B.ID And A.����ID=[1] And A.��ҳID=[2] And A.Ӥ��=[3] Order by B.����,A.��ʼʱ��"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ��Ѷ������µ�", mlng����ID, mlng��ҳID, mintӤ��)
    rsTemp.Filter = "����=-1"
    rsTemp.Sort = "����,��ʼʱ��"
    mstrCurForamt = ""
    mblnExist���µ� = rsTemp.RecordCount
    With rsTemp
        Do While Not .EOF
            mstrCurForamt = IIf(mstrCurForamt = "", "", mstrCurForamt & ",") & Val(!ID)
            .MoveNext
        Loop
    End With
    
    rsTemp.Filter = "����=1"
    mblnExist����ͼ = rsTemp.RecordCount
    rsTemp.Filter = "����<>-1"
    mblnExist��¼�� = rsTemp.RecordCount
    rsTemp.Filter = 0
    
    If mintӤ�� <> 0 Then mblnExist����ͼ = True
    
    '��ȡ�ļ�����
    gstrSQL = "SELECT A.����ID, B.���� AS ����, A.��ʽID, A.�ļ�����, A.��ʼʱ��, C.����,C.��� || '-' || C.���� As ��ʽ" & vbNewLine & _
            "  FROM ���˻����ļ� A, �����ļ��б� C, ���ű� B" & vbNewLine & _
            "  WHERE A.��ʽID = C.ID AND A.����ID = B.ID AND A.ID = [1]"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ļ�����", mlngFile)
    If rsTemp.RecordCount <> 0 Then
        mlng����ID = rsTemp!����ID
        mstrDept = rsTemp!����
        lng��ʽ = rsTemp!��ʽID
        str�ļ����� = rsTemp!�ļ�����
        str��ʼʱ�� = Format(rsTemp!��ʼʱ��, "yyyy-MM-dd HH:mm:ss")
        lng���� = Val(NVL(rsTemp!����, 0))
        str��ʽ = NVL(rsTemp!��ʽ, "-")
    End If
    
    '��ȡ������صĲ����ļ�������
'    gstrSQL = " Select ID,����,���,���||'-'||���� AS ��ʽ From �����ļ��б� " & _
'              " Where ����=3 And (ͨ�� =1 OR (ͨ��=2 And ID IN " & _
'              "     (Select �ļ�ID FROM ����Ӧ�ÿ��� Where ����ID = [1]))) " & _
'              " Order by ����,���"
    gstrSQL = "Select ID, ����, ���, ��ʽ" & vbNewLine & _
        "From (Select ID, ����, ���, ��� || '-' || ���� As ��ʽ" & vbNewLine & _
        "       From �����ļ��б�" & vbNewLine & _
        "       Where ���� = 3 And ���� <> 1 And (ͨ�� = 1 Or (ͨ�� = 2 And ID In (Select �ļ�id From ����Ӧ�ÿ��� Where ����id = [1])))" & vbNewLine & _
        "       Union All" & vbNewLine & _
        "       Select ID, ����, ���, ��� || '-' || ���� As ��ʽ" & vbNewLine & _
        "       From �����ļ��б�" & vbNewLine & _
        "       Where ���� = 3 And ���� = 1 And" & vbNewLine & _
        "             1 = (Select 1" & vbNewLine & _
        "                  From ��������˵�� A, ���ű� B" & vbNewLine & _
        "                  Where a.�������� = '����' And a.����id = b.Id And b.Id = [1] And Rownum < 2))" & vbNewLine & _
        "Order By ����, ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������صĲ����ļ�", glng����ID)
    With rsTemp
        Me.cbo��ʽ��Դ.Clear
        Me.cbo��ʽ��Դ.Tag = ""
        intIndex = 0
        Do While Not .EOF
            'If (((!���� = -1 And NOT mblnExist���µ�) Or (!���� = 1 And Not mblnExist����ͼ) Or InStr(1, ",-1,1,", "," & !���� & ",") = 0) And mlngFile = 0) Or mlngFile <> 0 Then
            If (((!���� = 1 And Not mblnExist����ͼ) Or !���� <> 1) And mlngFile = 0) Or mlngFile <> 0 Then
                Me.cbo��ʽ��Դ.AddItem !��ʽ
                Me.cbo��ʽ��Դ.ItemData(Me.cbo��ʽ��Դ.NewIndex) = !ID
                If !���� = -1 Then Me.cbo��ʽ��Դ.Tag = IIf(Me.cbo��ʽ��Դ.Tag = "", "", Me.cbo��ʽ��Դ.Tag & ",") & intIndex
                If !���� = 1 Then mIntPartogramID = !ID
                If !ID = lng��ʽ Then
                    Me.cbo��ʽ��Դ.ListIndex = intIndex
                    blnSeek = True
                End If
                intIndex = intIndex + 1
            End If
            .MoveNext
        Loop
        If .RecordCount <> 0 And Not blnSeek Then Me.cbo��ʽ��Դ.ListIndex = 0
    End With
    
    If mlngFile <> 0 Then
        Me.txt�ļ�����.Text = str�ļ�����
        '���˻�����(A->B)�޸��ļ�ʱ��������ļ�������A����������������޷���ȡ�����ļ�����Ϣ��Ϊ�˱�֤�ļ���ȷ�ԣ���Ҫ�������⴦��
        If Not (Me.cbo��ʽ��Դ.ItemData(Me.cbo��ʽ��Դ.ListIndex) = lng��ʽ) Then
            Me.cbo��ʽ��Դ.AddItem str��ʽ
            Me.cbo��ʽ��Դ.ItemData(Me.cbo��ʽ��Դ.NewIndex) = lng��ʽ
            If lng���� = -1 Then Me.cbo��ʽ��Դ.Tag = IIf(Me.cbo��ʽ��Դ.Tag = "", "", Me.cbo��ʽ��Դ.Tag & ",") & Me.cbo��ʽ��Դ.NewIndex
            If lng���� = 1 Then mIntPartogramID = lng��ʽ
            Me.cbo��ʽ��Դ.ListIndex = Me.cbo��ʽ��Դ.NewIndex
        End If
        If str��ʼʱ�� <> "" Then Me.mskEdit.Text = Format(str��ʼʱ��, "YYYY-MM-DD HH:mm:ss"): mskEdit.Tag = mskEdit.Text
    Else
        mblnOnly = (Val(zlDatabase.GetPara("��Ӧ��ݻ����ļ�", glngSys, 1255, 0)) = 0)
    End If
    Me.cbo��ʽ��Դ.Enabled = (mlngFile = 0)
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mskEdit_GotFocus()
    Call zlControl.TxtSelAll(mskEdit)
End Sub

Private Sub txt�ļ�����_GotFocus()
    Call zlControl.TxtSelAll(txt�ļ�����)
End Sub
