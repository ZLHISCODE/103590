VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmNurseFileEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ļ��༭"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4155
   Icon            =   "frmNurseFileEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4155
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
      Left            =   1350
      TabIndex        =   5
      Top             =   1110
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   19
      Format          =   "yyyy-MM-dd HH:mm:ss"
      Mask            =   "####-##-## ##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txt�ļ����� 
      Height          =   285
      Left            =   1350
      MaxLength       =   50
      TabIndex        =   3
      Top             =   720
      Width           =   2415
   End
   Begin VB.ComboBox cbo��ʽ��Դ 
      Height          =   300
      Left            =   1350
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   330
      Width           =   2415
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
      Left            =   570
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
      Left            =   570
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
      Left            =   570
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
Private mstr��Ժʱ�� As String
Private mlngFile As Long        '�ļ�ID,����0��ʾ����,�����ʾ�޸�(�޸�ʱ�������޸��ļ���Դ)
Private mlng����id As Long
Private mlng��ҳid As Long
Private mintӤ�� As Long
Private mlng����ID As Long
Private mstrDept As String      '��ǰ����
Private mblnOK As Boolean       '�Ƿ񱣴�ɹ�
Private mblnExist���µ� As Boolean
Private mblnExist��¼�� As Boolean
Private mblnOnly As Boolean     'סԺ����ͬһʱ��ֻ��¼һ�ݻ����ļ�

Public Function ShowEditor(ByVal lng����id As Long, ByVal lng��ҳid As Long, ByVal intӤ�� As Integer, ByVal lng����ID As Long, _
    ByVal str���� As String, Optional ByVal lngFile As Long) As Boolean
    mblnOK = False
    mlng����id = lng����id
    mlng��ҳid = lng��ҳid
    mintӤ�� = intӤ��
    mlng����ID = lng����ID
    mlngFile = lngFile
    mstrDept = str����
    Me.Show 1
    ShowEditor = mblnOK
End Function

Private Sub cbo��ʽ��Դ_Click()
    Dim bln���µ� As Boolean
    txt�ļ�����.Text = Split(cbo��ʽ��Դ.Text, "-")(1)
    If cbo��ʽ��Դ.ListIndex <> Val(cbo��ʽ��Դ.Tag) Or cbo��ʽ��Դ.Tag = "" Then '��¼��ʱ���´���
        txt�ļ�����.Text = "[" & mstrDept & "]" & txt�ļ�����.Text
    End If
    
    '����:������滤���ļ���ȱʡʱ��Ϊ��Ժʱ��,����Ϊ��ǰʱ��
    '�޸�:�����ļ��Ŀ�ʼʱ�䲻��С����Ժʱ��,���ܴ������ݷ���ʱ��,�����������ܴ��ڵ�ǰʱ��
    bln���µ� = (cbo��ʽ��Դ.Tag <> "" And cbo��ʽ��Դ.ListIndex = Val(cbo��ʽ��Դ.Tag))
    If mlngFile = 0 Then
        If (Not mblnExist��¼�� And Not bln���µ�) Or (Not mblnExist���µ� And bln���µ�) Then
            mskEdit.Text = mstr��Ժʱ��
        Else
            mskEdit.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngID As Long
    On Error GoTo errHand
    If txt�ļ�����.Text = "" Then
        MsgBox "�������ļ����ƣ�", vbInformation, gstrSysName
        txt�ļ�����.SetFocus
        Exit Sub
    End If
    If LenB(StrConv(txt�ļ�����.Text, vbFromUnicode)) > 50 Then
        MsgBox "�ļ����Ƴ����������50���ַ���25�����֣�", vbInformation, gstrSysName
        txt�ļ�����.SetFocus
        Exit Sub
    End If
    
    If mlngFile = 0 Then
        lngID = zlDatabase.GetNextId("���˻����ļ�")
    Else
        lngID = mlngFile
    End If
    
    gstrSQL = "ZL_���˻����ļ�_UPDATE(" & lngID & "," & mlng����ID & "," & mlng����id & "," & mlng��ҳid & "," & mintӤ�� & "," & _
              cbo��ʽ��Դ.ItemData(cbo��ʽ��Դ.ListIndex) & ",'" & txt�ļ�����.Text & "',to_date('" & mskEdit.Text & "','yyyy-MM-dd hh24:mi:ss')," & IIf(mlngFile = 0 And mblnOnly, 1, 0) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "��������")
    
    mblnOK = True
    Unload Me
    Exit Sub
errHand:
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
    Dim lng��ʽ As Long, str�ļ����� As String, str��ʼʱ�� As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '��ȡ��ǰ������Ժʱ��
    gstrSQL = " Select ��Ժ���� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ǰ������Ժʱ��", mlng����id, mlng��ҳid)
    mstr��Ժʱ�� = Format(rsTemp!��Ժ����, "yyyy-MM-dd HH:mm:ss")
    
    '����Ƿ����趨���µ�,���Ѵ����������ٴ�������µ�
    gstrSQL = " Select B.����" & _
              " From ���˻����ļ� A,�����ļ��б� B" & _
              " Where A.��ʽID=B.ID And A.����ID=[1] And A.��ҳID=[2] And A.Ӥ��=[3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ��Ѷ������µ�", mlng����id, mlng��ҳid, mintӤ��)
    rsTemp.Filter = "����=-1"
    mblnExist���µ� = rsTemp.RecordCount
    rsTemp.Filter = "����<>-1"
    mblnExist��¼�� = rsTemp.RecordCount
    rsTemp.Filter = 0
    
    '��ȡ�ļ�����
    gstrSQL = " Select A.����ID,B.���� AS ����,A.��ʽID,A.�ļ�����,A.��ʼʱ�� " & _
              " From ���˻����ļ� A,���ű� B" & _
              " Where A.����ID=B.ID And A.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ļ�����", mlngFile)
    If rsTemp.RecordCount <> 0 Then
        mlng����ID = rsTemp!����ID
        mstrDept = rsTemp!����
        lng��ʽ = rsTemp!��ʽID
        str�ļ����� = rsTemp!�ļ�����
        str��ʼʱ�� = Format(rsTemp!��ʼʱ��, "yyyy-MM-dd HH:mm:ss")
    End If
    
    '��ȡ������صĲ����ļ�������
    gstrSQL = "Select ID,����,���||'-'||���� AS ��ʽ From �����ļ��б� Where ����=3 And ͨ�� > 0  Order by ����,���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������صĲ����ļ�")
    With rsTemp
        Me.cbo��ʽ��Դ.Clear
        Do While Not .EOF
            If (((!���� = -1 And Not mblnExist���µ�) Or !���� <> -1) And mlngFile = 0) Or mlngFile <> 0 Then
                Me.cbo��ʽ��Դ.AddItem !��ʽ
                Me.cbo��ʽ��Դ.ItemData(Me.cbo��ʽ��Դ.NewIndex) = !ID
                If !���� = -1 Then Me.cbo��ʽ��Դ.Tag = .AbsolutePosition - 1
                If !ID = lng��ʽ Then
                    Me.cbo��ʽ��Դ.ListIndex = .AbsolutePosition - 1
                    blnSeek = True
                End If
            End If
            .MoveNext
        Loop
        If .RecordCount <> 0 And Not blnSeek Then Me.cbo��ʽ��Դ.ListIndex = 0
    End With
    
    If mlngFile <> 0 Then
        Me.txt�ļ�����.Text = str�ļ�����
        If str��ʼʱ�� <> "" Then Me.mskEdit.Text = str��ʼʱ��
    Else
        mblnOnly = (Val(zlDatabase.GetPara("��Ӧ��ݻ����ļ�", glngSys, 1255, 0)) = 0)
    End If
    Me.cbo��ʽ��Դ.Enabled = (mlngFile = 0)
    Exit Sub
errHand:
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
