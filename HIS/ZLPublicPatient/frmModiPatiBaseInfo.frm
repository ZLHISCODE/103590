VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.9#0"; "zlIDKind.ocx"
Begin VB.Form frmModiPatiBaseInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���˻�����Ϣ����"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4965
   Icon            =   "frmModiPatiBaseInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtInfo 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2115
      MaxLength       =   100
      TabIndex        =   16
      Top             =   3000
      Width           =   2070
   End
   Begin VB.OptionButton optType 
      Caption         =   "סԺ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3390
      TabIndex        =   12
      Top             =   2085
      Width           =   870
   End
   Begin VB.OptionButton optType 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2115
      TabIndex        =   11
      Top             =   2085
      Width           =   855
   End
   Begin VB.ComboBox cmbNum 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmModiPatiBaseInfo.frx":030A
      Left            =   2115
      List            =   "frmModiPatiBaseInfo.frx":030C
      TabIndex        =   14
      Text            =   "cmbNum"
      Top             =   2475
      Width           =   2070
   End
   Begin VB.ComboBox cboAge 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1590
      Width           =   705
   End
   Begin VB.TextBox txtAge 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   2  'OFF
      Left            =   2115
      TabIndex        =   8
      Top             =   1590
      Width           =   1350
   End
   Begin VB.ComboBox cboSex 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmModiPatiBaseInfo.frx":030E
      Left            =   2115
      List            =   "frmModiPatiBaseInfo.frx":0310
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   675
      Width           =   2070
   End
   Begin VB.TextBox txtPatient 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2115
      MaxLength       =   100
      TabIndex        =   1
      Top             =   210
      Width           =   2070
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2010
      TabIndex        =   17
      Top             =   3690
      Width           =   1300
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3345
      TabIndex        =   18
      Top             =   3690
      Width           =   1300
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   0
      TabIndex        =   19
      Top             =   3480
      Width           =   5310
   End
   Begin MSMask.MaskEdBox medBirthdayTime 
      Height          =   360
      Left            =   3480
      TabIndex        =   6
      Top             =   1140
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   635
      _Version        =   393216
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "hh:mm"
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medBirthdayDate 
      Bindings        =   "frmModiPatiBaseInfo.frx":0312
      Height          =   360
      Left            =   2115
      TabIndex        =   5
      Top             =   1140
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   635
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "YYYY-MM-DD"
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin zlIDKind.IDKindNew IDKind 
      Height          =   360
      Left            =   1410
      TabIndex        =   20
      ToolTipText     =   "��ݼ�F4"
      Top             =   210
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   635
      Appearance      =   2
      IDKindStr       =   $"frmModiPatiBaseInfo.frx":031D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   12
      FontName        =   "����"
      IDKind          =   -1
      BackColor       =   -2147483633
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�޸�ԭ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1095
      TabIndex        =   15
      Top             =   3060
      Width           =   960
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1095
      TabIndex        =   10
      Top             =   2085
      Width           =   960
   End
   Begin VB.Label lblNum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Һŵ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1095
      TabIndex        =   13
      Top             =   2535
      Width           =   960
   End
   Begin VB.Label lbl�������� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1095
      TabIndex        =   4
      Top             =   1200
      Width           =   960
   End
   Begin VB.Image imgFlag 
      Height          =   480
      Left            =   285
      Picture         =   "frmModiPatiBaseInfo.frx":03A4
      Top             =   375
      Width           =   480
   End
   Begin VB.Label lblAge 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1560
      TabIndex        =   7
      Top             =   1650
      Width           =   480
   End
   Begin VB.Label lblSex 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Ա�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1545
      TabIndex        =   2
      Top             =   750
      Width           =   480
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   885
      TabIndex        =   0
      Top             =   270
      Width           =   480
   End
End
Attribute VB_Name = "frmModiPatiBaseInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mblnOK As Boolean
Private mlng����ID As Long
Private mlng����ID As Long
Private mstrģ�� As String
Private mint���� As Integer
Private mstrAgeAndBirth As String     '��¼�޸�ǰ��������ͳ������� ��ʽ��"����_��������"
Private mblnChange As Boolean
Private mblnDrop As Boolean
Private mrsTmp As New ADODB.Recordset
Private mblnNotClick As Boolean
Private mblnBatch As Boolean
Private mstrName As String '��¼�������˻�����Ϣǰǰ���˵�����

Public Function ShowMe(ByVal frmParent As Object, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal int���� As Integer, ByVal strģ�� As String, Optional ByVal blnBatch As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '���:lng����ID-����ID
    '     lng����ID=��0:�Һ�ID����ҳID(�����Զ���λ��Ҫ�޸ĵ�ĳһ��סԺ�����)������0��ʾ��Ҫ�û��ֹ�ѡ�������ﻹ��סԺ
    '     int���� 1-����;2-סԺ
    '     strģ��=���øù��ܵ�ģ����������"����Һ�"��"��鱨��"��
    '����:strInfo:��Ϣ�������µı仯��Ϣ
    '����:
    '����:������
    '����:2013-10-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlng����ID = lng����ID
    mlng����ID = lng����ID
    mint���� = int����
    mstrģ�� = strģ��
    If blnBatch = False Then
        If lng����ID = 0 Or lng����ID = 0 Then
            MsgBox "�������Ե���������Ϣʱ,�贫�����Ĳ���ID�;���ID!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    mblnBatch = blnBatch
    mblnChange = False
    mblnOK = False
    mblnNotClick = False
    
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Sub InitDicts()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHand
    
    txtPatient.Text = ""
    txtPatient.MaxLength = gobjComlib.Sys.FieldsLength("������Ϣ", "����")
    txtAge.Text = ""
    cboAge.Clear
    cboAge.AddItem "��"
    cboAge.AddItem "��"
    cboAge.AddItem "��"
    cboAge.ListIndex = 0
    txtAge.MaxLength = gobjComlib.Sys.FieldsLength("������Ϣ", "����")
    
    cboSex.Clear
    
    strSQL = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From �Ա� Order by ����"
    Call gobjDatabase.OpenRecordset(rsTmp, strSQL, "�Ա�")
    Do While Not rsTmp.EOF
        cboSex.AddItem rsTmp!���� & "-" & rsTmp!����
        If rsTmp!ȱʡ = 1 Then
            cboSex.ListIndex = cboSex.NewIndex
            cboSex.ItemData(cboSex.NewIndex) = 1
        End If
    rsTmp.MoveNext
    Loop
    
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Function LoadPatiBaseInfo() As Boolean
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim lngIndex As Long
    
    On Error GoTo ErrHand
    
    Call ClearInfo
    
    If mlng����ID <> 0 Then
        If mint���� = 1 Then '���ﲡ��
            strSQL = "Select  Nvl(a.����, b.����) ����, Nvl(a.�Ա�, b.�Ա�) �Ա�,nvl(a.����,b.����) ����,b.��������,B.��������,B.����" & vbNewLine & _
                " From ���˹Һż�¼ A,������Ϣ b" & vbNewLine & _
                " Where a.����id = [1] And A.id=[2] and a.����ID=B.����ID And b.ͣ��ʱ�� is NULL"
        Else 'סԺ����
            strSQL = " Select Nvl(a.����, b.����) ����, Nvl(a.�Ա�, b.�Ա�) �Ա�,nvl(a.����,b.����) ����,B.��������,B.��������,B.����,A.��Ժ���� " & vbNewLine & _
                    " From ������ҳ a, ������Ϣ b" & vbNewLine & _
                    " Where a.����id = b.����id And a.����id = [1] And a.��ҳid = [2] And b.ͣ��ʱ�� is NULL"
        End If
    Else
        strSQL = "Select ����,�Ա�,����,��������,��������,���� From ������Ϣ Where ����ID=[1] And ͣ��ʱ�� is NULL"
    End If
    
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ���˻�����Ϣ", mlng����ID, mlng����ID)
    
    mblnChange = False
    
    If Not rsTmp.EOF Then
        txtPatient.Text = gobjCommFun.Nvl(rsTmp!����)
        mstrName = gobjCommFun.Nvl(rsTmp!����)
        txtPatient.ForeColor = GetPatiColor(Nvl(rsTmp!��������), IIf(IsNull(rsTmp!����) = True, &H80000008, vbRed))
        lblName.Tag = txtPatient.ForeColor
        cboSex.ListIndex = gobjComlib.cbo.FindIndex(cboSex, Nvl(rsTmp!�Ա�), True)
        If cboSex.ListIndex = -1 And Not IsNull(rsTmp!�Ա�) Then
            cboSex.AddItem rsTmp!�Ա�, 0
            cboSex.ListIndex = cboSex.NewIndex
        End If
        Call gobjComlib.zlControl.LoadOldData("" & rsTmp!����, txtAge, cboAge)
        mblnChange = False
        medBirthdayDate.Text = Format(IIf(IsNull(rsTmp!��������), "____-__-__", rsTmp!��������), "YYYY-MM-DD")
        If Nvl(rsTmp!����) Like "Լ*" Or Trim(Nvl(rsTmp!����)) = "����" Then
            If "" & rsTmp!�������� = "____-__-__" Then
                medBirthdayDate.Enabled = False
                medBirthdayTime.Enabled = False
            End If
        Else
            medBirthdayDate.Enabled = True
            medBirthdayTime.Enabled = True
        End If
        mblnChange = True
        If mlng����ID <> 0 And mint���� = 2 Then medBirthdayDate.Tag = rsTmp!��Ժ���� & ""
        If Not IsNull(rsTmp!��������) Then
            If CDate(medBirthdayDate.Text) - CDate(rsTmp!��������) <> 0 Then
                mblnChange = False
                medBirthdayTime.Text = Format(rsTmp!��������, "HH:MM")
                mblnChange = True
            End If
        Else
            medBirthdayTime.Text = "__:__"
            mblnChange = False
            Call RecalcBirthDay
            mblnChange = True
        End If
    Else
        MsgBox "��ȡ���˻�����Ϣʧ��,��������ȷ��Ҫ������Ϣ�����Ĳ��ˣ�", vbInformation, gstrSysName
        mlng����ID = 0: mlng����ID = 0
        If mblnBatch = True Then
            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        Else
            On Error Resume Next
            Unload Me
            Err.Clear
        End If
        Exit Function
    End If
    mstrAgeAndBirth = txtAge.Text & cboAge.Text & "_" & medBirthdayDate.Text & medBirthdayTime.Text
    Call LoadPatiData
    
    mblnChange = True
    
    LoadPatiBaseInfo = True
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Private Sub LoadPatiData()
'-----------------------------------------------
'����:��ȡ���˾����¼��Ϣ(סԺ����������¼)
'
'-----------------------------------------------
    Dim strSQL As String
    Dim bln���� As Boolean, blnסԺ As Boolean
    
    On Error GoTo ErrHand
    strSQL = "Select * From(" & _
        " Select 1 ����,ID Id, No,0 ��������, to_char(�Ǽ�ʱ��,'YYYY-MM-DD hh24:mi:ss') �Ǽ�ʱ��,NULL ��������,NULL ����,NULL as ��Ժ���� " & vbNewLine & _
        " From ���˹Һż�¼" & vbNewLine & _
        " Where ����id = [1] And Mod(��¼״̬, 2) <> 0" & vbNewLine & _
        " Union All" & vbNewLine & _
        " Select 2 ����,��ҳId Id, '' || ��ҳid No,��������, to_char(�Ǽ�ʱ��,'YYYY-MM-DD hh24:mi:ss') �Ǽ�ʱ��,��������,����,��Ժ���� " & vbNewLine & _
        " From ������ҳ" & vbNewLine & _
        " Where ����id = [1] And Nvl(��ҳid, 0) <> 0) Order By No Desc"
    Set mrsTmp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ�����¼", mlng����ID)
    
    optType(0).Enabled = True
    optType(1).Enabled = True
    cmbNum.Enabled = True
    cmbNum.Clear
    If mrsTmp.RecordCount > 0 Then
        mrsTmp.Filter = "����=1"
        bln���� = mrsTmp.RecordCount > 0
        mrsTmp.Filter = "����=2"
        blnסԺ = mrsTmp.RecordCount > 0
        
        mblnChange = True
        If bln���� = True And blnסԺ = True Then
            If mlng����ID <> 0 Then
                If mint���� = 1 Then
                    optType(0).Value = True
                Else
                    optType(1).Value = True
                End If
            Else
                optType(0).Value = True
            End If
        Else
            If bln���� = True Then
                optType(0).Value = True
                optType(1).Enabled = False
            Else
                optType(1).Value = True
                optType(0).Enabled = False
            End If
        End If
        Call optType_Click(IIf(optType(0).Value = True, 0, 1))
    Else
        mblnChange = False
        '���˴�δ�ҺŻ�סԺ
        optType(0).Value = True
        optType(0).Enabled = False
        optType(1).Enabled = False
        cmbNum.Enabled = False
        mblnChange = True
    End If
    
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cboAge_LostFocus()
    If Trim(txtAge.Text) = "" Then Exit Sub
    If Not CheckOldData(txtAge, cboAge) Then Exit Sub
    
    If Not IsDate(medBirthdayDate.Text) Then
        mblnChange = False
        Call RecalcBirthDay
        mblnChange = True
    End If
End Sub

Private Sub cboSex_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cboSex.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call gobjCommFun.PressKey(vbKeyF4)
    lngIdx = gobjComlib.cbo.MatchIndex(cboSex.hWnd, KeyAscii, 0.5)
    If lngIdx <> -2 Then cboSex.ListIndex = lngIdx
End Sub

Private Sub cmbNum_Click()
    Dim lngColor As Long
    lngColor = Val(lblName.Tag)
    medBirthdayDate.Tag = ""
    If optType(0).Value = True Then
        txtPatient.ForeColor = lngColor
    Else
        If mrsTmp Is Nothing Then Exit Sub
        If mrsTmp.State = adStateClosed Then Exit Sub
        If optType(1).Value = True And cmbNum.ListIndex <> -1 Then
            mrsTmp.Filter = "����=2 And ID=" & Val(cmbNum.ItemData(cmbNum.ListIndex))
            If mrsTmp.RecordCount > 0 Then
                lngColor = GetPatiColor(Nvl(mrsTmp!��������), IIf(IsNull(mrsTmp!����) = True, &H80000008, vbRed))
                medBirthdayDate.Tag = mrsTmp!��Ժ���� & ""
            End If
        End If
        txtPatient.ForeColor = lngColor
    End If
End Sub

Private Sub cmbNum_KeyDown(KeyCode As Integer, Shift As Integer)
    If cmbNum.Locked Then Exit Sub
    mblnDrop = False
    If KeyCode = 13 Then
        mblnDrop = SendMessage(cmbNum.hWnd, &H157, 0, 0) = 1
    End If
End Sub

Private Sub cmbNum_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer, iCount As Integer
    Dim strText As String, strResult As String, strFilter As String
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮
    Dim rsTemp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        If cmbNum.Locked Then
            Call gobjCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        strText = UCase(cmbNum.Text)
        If cmbNum.ListIndex <> -1 Then
            '�����б�ʱ,�����ı�������������
            If strText <> cmbNum.List(cmbNum.ListIndex) Then Call gobjControl.CboSetIndex(cmbNum.hWnd, -1)
        End If
        If strText = "" Then
            cmbNum.ListIndex = -1
        ElseIf cmbNum.ListIndex = -1 Then
            intIdx = -1
            strFilter = "����=" & IIf(optType(0).Value = True, 1, 2)
            '�ȸ��Ƽ�¼��
            Set rsTemp = gobjDatabase.zlCopyDataStructure(mrsTmp)
            
            strCompents = Replace(gstrLike, "%", "*") & strText & "*"
            
            If IsNumeric(strText) Then
                intInputType = 0
            ElseIf gobjCommFun.IsCharAlpha(strText) Then
                intInputType = 1
            Else
                intInputType = 2
            End If
            
            mrsTmp.Filter = strFilter: iCount = 0
            With mrsTmp
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not mrsTmp.EOF
                    Select Case intInputType
                    Case 0  '�������ȫ����
                        '������������,��Ҫ���:
                        '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������
                        '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                        
                        
                        '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�Ӷ�λ
                        If Nvl(!NO) = strText Then strResult = Nvl(!NO): iCount = 0: Exit Do
                        
                        '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                        If Val(Nvl(!NO)) = Val(strText) Then
                            If iCount = 0 Then strResult = Nvl(!NO)
                            iCount = iCount + 1
                        End If
                        
                        '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                         If Val(Nvl(!NO)) Like strCompents Then
                            If isCheckExists(Nvl(!NO)) Then Call gobjDatabase.zlInsertCurrRowData(mrsTmp, rsTemp)
                         End If
                    Case 1  '�������ȫ��ĸ
                        '����:
                        ' 1.����ļ������,��ֱ�Ӷ�λ
                        ' 2.���ݲ�����ƥ����ͬ����
                        
                        '1.����ļ������,��ֱ�Ӷ�λ
                        If Trim(Nvl(!NO)) = strText Then
                            If iCount = 0 Then strResult = Nvl(!NO)   '���ܴ��ڶ����ͬ�Ķ��
                            iCount = iCount + 1
                        End If
                        
                        '2.���ݲ�����ƥ����ͬ����
                        If Trim(Nvl(!NO)) Like strCompents Then
                            If isCheckExists(Nvl(!NO)) Then Call gobjDatabase.zlInsertCurrRowData(mrsTmp, rsTemp)
                        End If
                    Case Else  ' 2-����
                        '����:���ܴ��ں��ֵ����,����������N001���������ZYK01�������
                        '1.����\�������,ֱ�Ӷ�λ
                        '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                        
                        '1.����\�������,ֱ�Ӷ�λ
                        If Trim(!NO) = strText Then
                            If iCount = 0 Then strResult = Nvl(!NO)   '���ܴ��ڶ����ͬ�Ķ��
                            iCount = iCount + 1
                        End If
                        
                        '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                        If Trim(!NO) Like strCompents Then
                            If isCheckExists(Nvl(!NO)) Then Call gobjDatabase.zlInsertCurrRowData(mrsTmp, rsTemp)
                        End If
                    End Select
                    mrsTmp.MoveNext
                Loop
            End With
            If iCount > 1 Then strResult = ""
            If strResult = "" And rsTemp.RecordCount = 1 Then strResult = Nvl(rsTemp!NO)
            'ֱ�Ӷ�λ
            If strResult <> "" Then
                rsTemp.Close: Set rsTemp = Nothing
                If isCheckExists(strResult, True) Then gobjCommFun.PressKey vbKeyTab
                Exit Sub
            End If
            
            '��Ҫ����Ƿ��ж������������ļ�¼
            If rsTemp.RecordCount <> 0 Then
                '�Ȱ�ĳ�ַ�ʽ��������
                If optType(0).Value = True Then
                    rsTemp.Sort = "�Ǽ�ʱ�� DESC"
                Else
                    rsTemp.Sort = "ID DESC"
                End If
                '����ѡ����
                Dim rsReturn As ADODB.Recordset
                If gobjDatabase.zlShowListSelect(Me, glngSys, 1101, cmbNum, rsTemp, True, "", "����", rsReturn) Then
                    If Not rsReturn Is Nothing Then
                        If rsReturn.RecordCount <> 0 Then
                            '���ж�λ
                            If isCheckExists(Nvl(rsReturn!NO), True) Then
                                'zlCommFun.PressKey vbKeyTab
                            End If
                        End If
                    End If
                End If
            Else
                'δ�ҵ�
                rsTemp.Close: Set rsTemp = Nothing
                KeyAscii = 0: gobjControl.TxtSelAll cmbNum: Exit Sub
            End If
            rsTemp.Close: Set rsTemp = Nothing
             
        ElseIf Not mblnDrop Then
            '�س���꾭��
            Call gobjCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        If cmbNum.ListIndex = -1 Then
            cmbNum.Text = ""
            Exit Sub
        Else
            If intIdx <> -1 And mblnDrop Then
                '�����س�-ǿ�м���Click
            ElseIf intIdx <> cmbNum.ListIndex And intIdx <> -1 Then
                '������ѡ��-�Զ�����Click
                cmbNum.SetFocus
                Call gobjCommFun.PressKey(vbKeyF4)
                Exit Sub
            ElseIf intIdx <> -1 Then
                'һ��������-ǿ�м���Click
            End If
        End If
        Call gobjCommFun.PressKey(vbKeyTab)
    Else
        If optType(0).Value = True Then
            If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then KeyAscii = 0
        Else
            If InStr("0123456789" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub cmbNum_Validate(Cancel As Boolean)
    If cmbNum.Text <> "" Then
        If gobjComlib.cbo.FindIndex(cmbNum, gobjComlib.ZLStr.NeedName(cmbNum.Text), True) = -1 Then cmbNum.ListIndex = -1: cmbNum.Text = ""
    End If
    If cmbNum.Text = "" And cmbNum.Enabled = True And cmbNum.ListCount > 0 Then '˵��¼�����Ϣ���������б���
        MsgBox "��ѡ��" & IIf(optType(0).Value = True, "�Һŵ���", "סԺ����"), vbInformation, gstrSysName
        Cancel = True
    End If
End Sub

Private Function isCheckExists(ByVal strNO As String, Optional blnLocateItem As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ƿ��ڿ����������б���.
    '���:str����-����
    '     blnLocateItem:�Ƿ�ֱ�Ӷ�λ
    '����:
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To cmbNum.ListCount - 1
        If IIf(optType(0).Value = True, gobjComlib.ZLStr.NeedName(cmbNum.List(i)), Val(cmbNum.List(i))) = strNO Then
            If blnLocateItem Then cmbNum.ListIndex = i
            isCheckExists = True
            Exit Function
        End If
    Next
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
'���ܣ��������У��ͱ���
    Dim strInfo As String
    Dim str���� As String, str�������� As String, str�Ա� As String
    Dim lngTmp As Long
    Dim blnTrue As Boolean
    Dim blnEMPI As Boolean
    
    '��һ�������ݺϷ���У��
    If mlng����ID = 0 Then
        MsgBox "������ȷ��Ҫ�����Ĳ���!", vbInformation, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Sub
    End If
    
    If Trim(txtPatient.Text) = "" Then
        MsgBox "�������벡�˵�������", vbInformation, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus: Exit Sub
    End If
    If cboSex.ListIndex = -1 Then
        MsgBox "����ȷ�����˵��Ա�", vbInformation, gstrSysName
        If cboSex.Enabled And cboSex.Visible Then cboSex.SetFocus: Exit Sub
    End If
    
    If medBirthdayDate.Enabled Then
        If Not IsDate(medBirthdayDate.Text) Then
            MsgBox "������ȷ���벡�˵ĳ������ڣ�", vbInformation, gstrSysName
            If medBirthdayDate.Enabled And medBirthdayDate.Visible Then medBirthdayDate.SetFocus: Exit Sub
        End If
    End If
    If Trim(txtAge.Text) = "" Then
        MsgBox "�������벡�˵����䣡", vbInformation, gstrSysName
        If txtAge.Enabled And txtAge.Visible Then txtAge.SetFocus: Exit Sub
    End If
    '103905 �޸�ԭ�������󳤶�Ϊ100���ַ�
    If Not gobjControl.TxtCheckInput(txtInfo, "�޸�ԭ��") Then Exit Sub
    '76409,������,2014-08-06,����Ϸ��Լ��
    str���� = txtAge.Text
    If IsNumeric(str����) Then str���� = str���� & cboAge.Text
    If IsDate(medBirthdayDate.Text) Then
        If medBirthdayTime.Text = "__:__" Then
            str�������� = Format(medBirthdayDate.Text, "YYYY-MM-DD")
        Else
            str�������� = Format(medBirthdayDate.Text & " " & medBirthdayTime.Text, "YYYY-MM-DD HH:MM:SS")
        End If
        If mstrAgeAndBirth = txtAge.Text & cboAge.Text & "_" & medBirthdayDate.Text & medBirthdayTime.Text Then
            '97836 ֻ�޸�����ʱ����ǿ���޸�����
            blnTrue = CheckAge(str����)
        Else
            If mint���� = 2 And IsDate(medBirthdayDate.Tag) Then
                blnTrue = CheckAge(str����, str��������, , medBirthdayDate.Tag)
            Else
                blnTrue = CheckAge(str����, str��������)
            End If
        End If
    Else
        blnTrue = CheckAge(str����)
    End If
    If blnTrue = False Then
        If txtAge.Enabled And txtAge.Visible Then txtAge.SetFocus: Exit Sub
    End If
    
    If Not gobjComlib.zlControl.TxtCheckInput(txtPatient, "����") Then Exit Sub
    If Not gobjComlib.zlControl.TxtCheckInput(txtAge, "����") Then Exit Sub
    If Not CheckOldData(txtAge, cboAge) Then Exit Sub
    
    If cmbNum.Enabled And cmbNum.ListIndex = -1 Then
        MsgBox "����ѡ��" & IIf(optType(0).Value = True, "�Һŵ���", "סԺ����") & "��", vbInformation, gstrSysName
        If cmbNum.Enabled And cmbNum.Visible Then cmbNum.SetFocus: Exit Sub
    End If
    
    If medBirthdayDate.Enabled Then
        If medBirthdayTime = "__:__" Then
            str�������� = Format(medBirthdayDate.Text, "YYYY-MM-DD")
        Else
            str�������� = Format(medBirthdayDate.Text & " " & medBirthdayTime.Text, "YYYY-MM-DD HH:mm")
        End If
    End If
    
    If InStr(1, cboSex.Text, "-") <> 0 Then
        str�Ա� = Split(cboSex.Text, "-")(1)
    Else
        str�Ա� = cboSex.Text
    End If
    
    str���� = Trim(txtAge.Text)
    If IsNumeric(str����) Then str���� = str���� & cboAge.Text
    If cmbNum.ListIndex >= 0 Then
        mint���� = IIf(optType(1).Value = True, 2, 1)
        mlng����ID = Val(cmbNum.ItemData(cmbNum.ListIndex))
    Else
        mint���� = 1
        mlng����ID = 0
    End If
    strInfo = Trim(txtInfo.Text)
    'EMPI���
    blnEMPI = EMPI_LoadPati(Trim(txtPatient.Text), str�Ա�, str��������)
    '�ڶ��������ݱ���
    On Error GoTo ErrHand
    If Trim(txtPatient.Text) <> Trim(mstrName) Then
        If MsgBox("���Ƿ񽫲���������" & mstrName & "������Ϊ��" & txtPatient.Text & "��,�Ƿ���ĵ�����", vbYesNo, gstrSysName) = vbYes Then
            If SaveBaseInfo(mlng����ID, mlng����ID, Trim(txtPatient.Text), str�Ա�, str����, str��������, mstrģ��, mint����, strInfo, True, blnEMPI) = False Then
                If strInfo <> "" Then
                    MsgBox strInfo, vbInformation, gstrSysName
                End If
                Exit Sub
            End If
        Else
            txtPatient.SetFocus
            txtPatient.SelStart = 0
            txtPatient.SelLength = Len(txtPatient.Text)
            Exit Sub
        End If
    Else
        If SaveBaseInfo(mlng����ID, mlng����ID, Trim(txtPatient.Text), str�Ա�, str����, str��������, mstrģ��, mint����, strInfo, True, blnEMPI) = False Then
            If strInfo <> "" Then
                MsgBox strInfo, vbInformation, gstrSysName
            End If
            Exit Sub
        End If
    End If
    If strInfo <> "" Then
        MsgBox strInfo, vbInformation, gstrSysName
    End If
    mblnOK = True
    If mblnBatch = False Then Unload Me: Exit Sub
    mlng����ID = 0: mlng����ID = 0
    Call ClearInfo
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intIndex As Integer
    If KeyCode = vbKeyReturn Then
       If ActiveControl.Name <> txtPatient.Name And ActiveControl.Name <> txtAge.Name And ActiveControl.Name <> cmbNum.Name Then
           Call gobjCommFun.PressKey(vbKeyTab)
       End If
    ElseIf KeyCode = vbKeyF4 And mblnBatch = True Then
        If Shift = vbCtrlMask And IDKind.Enabled Then
            intIndex = IDKind.GetKindIndex("IC����")
            If intIndex < 0 Then Exit Sub
            IDKind.IDKind = intIndex: Call IDKind_Click(IDKind.GetCurCard)
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    '������Ϣ��ʼ��
    Call InitDicts
    
    If mblnBatch = True Then
        Call CreateMobjCard
        Call CreateSquareCardObject(Me, 1101)
         '��ʼ��
        Call IDKind.zlInit(Me, 100, 1101, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
        
        If Not gobjSquare.objSquareCard Is Nothing Then
            IDKind.IDKindStr = gobjSquare.objSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
        End If
    Else
        IDKind.Visible = False
        lblName.Left = lblSex.Left
    End If
    
    If mlng����ID <> 0 Then
        Call LoadPatiBaseInfo
    Else
        Call ClearInfo
        txtPatient.Text = ""
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand As String
    Dim strOutPatiInforXml As String
    If mblnBatch = False Then Exit Sub
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = New clsICCard
            Call mobjICCard.SetParent(Me.hWnd)
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txtPatient.Text = mobjICCard.Read_Card()
            If txtPatient.Text <> "" Then
                Call txtPatient_KeyPress(vbKeyReturn)
            End If
        End If
        Exit Sub
    End If
    
    lng�����ID = objCard.�ӿ����
    If lng�����ID <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '����:�����ӿ�
    '    '���:frmMain-���õĸ�����
    '    '       lngModule-���õ�ģ���
    '    '       strExpand-��չ����,������
    '    '       blnOlnyCardNO-������ȡ����
    '    '����:strOutCardNO-���صĿ���
    '    '       strOutPatiInforXML-(������Ϣ����.XML��)
    '    '����:��������    True:���óɹ�,False:����ʧ��\
    If gobjSquare.objSquareCard.zlReadCard(Me, 1101, lng�����ID, False, strExpand, strOutCardNO, strOutPatiInforXml) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If mblnBatch = False Then Exit Sub
    Set gobjSquare.objCurCard = objCard
    txtPatient.PasswordChar = "": txtPatient.IMEMode = 0
    '��Ҫ�����Ϣ,����ˢ����,���л�,���������ʾʧȥ����
    If txtPatient.Text <> "" And mblnNotClick = False Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    gobjComlib.zlControl.TxtSelAll txtPatient
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If mblnBatch = False Then Exit Sub
    If txtPatient.Text <> "" Or txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.����
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
End Sub

Private Sub medBirthdayTime_Change()
    Dim strBirthday As String
    If IsDate(medBirthdayTime.Text) And IsDate(medBirthdayDate.Text) And mblnChange Then
        strBirthday = Format(medBirthdayDate.Text & " " & medBirthdayTime.Text, "YYYY-MM-DD HH:MM:SS")
        If mint���� = 2 And IsDate(medBirthdayDate.Tag) Then
            txtAge.Text = ReCalcOld(CDate(strBirthday), cboAge, , , CDate(medBirthdayDate.Tag))
        Else
            txtAge.Text = ReCalcOld(CDate(strBirthday), cboAge)
        End If
    End If
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    If mblnBatch = False Then Exit Sub
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("IC��", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strCardNo
    If txtPatient.Text <> "" Then Call FindPati(objCard, txtPatient.Text, False)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    If mblnBatch = False Then Exit Sub
    If txtPatient.Locked Or txtPatient.Text <> "" Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("���֤", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strID
    If txtPatient.Text <> "" Then Call FindPati(objCard, txtPatient.Text, False)
End Sub

Private Sub medBirthdayDate_Change()
    Dim strBirthday As String
    If IsDate(medBirthdayDate.Text) And mblnChange Then
        mblnChange = False
        medBirthdayDate.Text = Format(CDate(medBirthdayDate.Text), "yyyy-mm-dd") '0002-02-02�Զ�ת��Ϊ2002-02-02,����,��������2002,ʵ��ֵȴ��0002
        mblnChange = True
        If medBirthdayTime.Text = "__:__" Then
            strBirthday = Format(medBirthdayDate.Text, "YYYY-MM-DD")
        Else
            strBirthday = Format(medBirthdayDate.Text & " " & medBirthdayTime.Text, "YYYY-MM-DD HH:MM:SS")
        End If
        If mint���� = 2 And IsDate(medBirthdayDate.Tag) Then
            txtAge.Text = ReCalcOld(CDate(strBirthday), cboAge, , , CDate(medBirthdayDate.Tag))
        Else
            txtAge.Text = ReCalcOld(CDate(strBirthday), cboAge)
        End If
    End If
End Sub

Private Sub medBirthdayDate_GotFocus()
    Call gobjCommFun.OpenIme
    gobjComlib.zlControl.TxtSelAll medBirthdayDate
End Sub

Private Sub medBirthdayDate_LostFocus()
    If medBirthdayDate.Text <> "____-__-__" And Not IsDate(medBirthdayDate.Text) Then
        medBirthdayDate.SetFocus
    End If
End Sub

Private Sub medBirthdayTime_GotFocus()
    Call gobjCommFun.OpenIme
    gobjComlib.zlControl.TxtSelAll medBirthdayTime
End Sub

Private Sub medBirthdayTime_KeyPress(KeyAscii As Integer)
    If Not IsDate(medBirthdayDate) Then
        KeyAscii = 0
        medBirthdayTime.Text = "__:__"
    End If
End Sub

Private Sub medBirthdayTime_Validate(Cancel As Boolean)
    If medBirthdayTime.Text <> "__:__" And Not IsDate(medBirthdayTime.Text) Then
        medBirthdayTime.SetFocus
        Cancel = True
    End If
End Sub

Private Sub optType_Click(Index As Integer)
    If mblnChange = False Or mrsTmp Is Nothing Then Exit Sub
    If mrsTmp.State = adStateClosed Then Exit Sub
     
    If Index = 0 Then
        lblNum.Caption = "�Һŵ���"
        mrsTmp.Filter = "����=1"
    ElseIf Index = 1 Then
        lblNum.Caption = "סԺ����"
        mrsTmp.Filter = "����=2"
    End If
    If Index = 0 Or Index = 1 Then
        cmbNum.Clear
        Do While Not mrsTmp.EOF
            cmbNum.AddItem Nvl(mrsTmp!NO) & IIf(Val("" & mrsTmp!��������) = 1, "-��������", IIf(Val("" & mrsTmp!��������) = 2, "-סԺ����", ""))
            cmbNum.ItemData(cmbNum.NewIndex) = Val(mrsTmp!ID)
            If mlng����ID = Val(mrsTmp!ID) Then
                cmbNum.ListIndex = cmbNum.NewIndex
            End If
        mrsTmp.MoveNext
        Loop
        
        If cmbNum.ListIndex = -1 And cmbNum.ListCount > 0 Then cmbNum.ListIndex = 0
        cmbNum.Enabled = mblnBatch
    End If
End Sub

Private Sub txtAge_GotFocus()
    Call gobjCommFun.OpenIme
    gobjControl.TxtSelAll txtAge
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboAge.Visible = False And IsNumeric(txtAge.Text) Then
            Call txtAge_Validate(False)
            Call cboAge.SetFocus
        Else
            Call gobjCommFun.PressKey(vbKeyTab)
        End If
        If Not IsNumeric(txtAge.Text) Then Call gobjCommFun.PressKey(vbKeyTab)
    Else
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyAscii))) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtAge_Validate(Cancel As Boolean)
    If Not IsNumeric(txtAge.Text) And Trim(txtAge.Text) <> "" Then
        If Not Trim(txtAge.Text) Like "Լ*" And Trim(txtAge.Text) <> "����" Then
            cboAge.ListIndex = -1: cboAge.Visible = False
            medBirthdayDate.Enabled = True
            medBirthdayTime.Enabled = True
        ElseIf Trim(txtAge.Text) Like "Լ*" Or Trim(txtAge.Text) = "����" Then
            If Trim(medBirthdayDate.Text) = "____-__-__" Then
                medBirthdayDate.Enabled = False
                medBirthdayTime.Enabled = False
            End If
            cboAge.ListIndex = -1: cboAge.Visible = False
        End If
    ElseIf cboAge.Visible = False Or medBirthdayDate.Enabled = True Then
        cboAge.ListIndex = 0: cboAge.Visible = True
        medBirthdayDate.Enabled = True
        medBirthdayTime.Enabled = True
    Else
        medBirthdayDate.Enabled = True
        medBirthdayTime.Enabled = True
    End If
End Sub

Private Sub txtInfo_GotFocus()
    Call gobjCommFun.OpenIme
    gobjControl.TxtSelAll txtInfo
End Sub

Private Sub txtPatient_Change()
    If mblnBatch = False Then Exit Sub
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")
End Sub

Private Sub txtPatient_GotFocus()
    gobjComlib.zlControl.TxtSelAll txtPatient
    If mblnBatch = False Then Exit Sub
    If Not mobjIDCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then mobjIDCard.SetEnabled (True)
    If Not mobjICCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then mobjICCard.SetEnabled (True)
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If mblnBatch = False Then Exit Sub
    If txtPatient.Locked Or txtPatient.Enabled = False Then Exit Sub
    Call IDKind.ActiveFastKey
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    Dim blnCard As Boolean, blnName As Boolean
    
    If Trim(txtPatient.Text) = "" Then
        Exit Sub
    End If
    
    If mblnBatch = False Then
        If KeyAscii = 13 Then gobjCommFun.PressKey vbKeyTab
        Exit Sub
    End If
    
    If IDKind.GetCurCard.���� = "�����" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    End If
    
    If IDKind.GetCurCard.���� Like "����*" Then
        blnCard = gobjCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
    ElseIf IDKind.IDKind = IDKind.GetKindIndex("�����") Or IDKind.IDKind = IDKind.GetKindIndex("סԺ��") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        txtPatient.IMEMode = 0
    End If
    
    'ˢ����ϻ���������س�
    If blnCard And Len(Me.txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Me.txtPatient.Text <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        
        Call FindPati(IDKind.GetCurCard, txtPatient.Text, blnCard)
    End If
End Sub

Private Sub txtPatient_LostFocus()
    If mblnBatch = True Then
        If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
        If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (False)
    End If
    txtPatient.Text = Trim(txtPatient.Text)
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub CreateMobjCard()
    '����������
    Set mobjIDCard = New clsIDCard
    Call mobjIDCard.SetParent(Me.hWnd)
    Set mobjICCard = New clsICCard
    Call mobjICCard.SetParent(Me.hWnd)
    Set mobjICCard.gcnOracle = gcnOracle
End Sub

Private Function FindPati(ByVal objCard As Card, ByVal strInput As String, Optional blnCard As Boolean = False) As Boolean
    '��ȡ������Ϣ
    Dim blnName As Boolean
    If Not GetPatient(objCard, strInput, blnCard, blnName) Then
        If IsNumeric(txtPatient.Text) Then
            txtPatient.PasswordChar = "": txtPatient.IMEMode = 0: txtPatient.Text = ""
        End If
        Call gobjControl.TxtSelAll(txtPatient)
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        If blnName = True Then gobjCommFun.PressKey vbKeyTab
    Else
        txtPatient.PasswordChar = ""
        txtPatient.IMEMode = 0
        Call LoadPatiBaseInfo
    End If
    
    FindPati = True
End Function

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, Optional blnCard As Boolean = False, Optional blnName As Boolean = False) As Boolean
'���ܣ���ȡ������Ϣ
    Dim lng�����ID As Long, lng����ID As Long
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnDo As Boolean
    Dim blnHavePassWord As Boolean
    Dim strPassWord As String, strErrMsg As String
    Dim strCard As String, strPati As String
    Dim vRect As RECT
    Dim blnCancel As Boolean
    
    On Error GoTo errH
    
    strSQL = "Select A.����ID,A.����,A.�Ա�,A.����,A.��������" & _
        " From ������Ϣ A" & _
        " Where A.ͣ��ʱ�� is NULL"
        
    If blnCard = True And objCard.���� Like "����*" Then    'ˢ��
        If IDKind.Cards.��ȱʡ������ And Not IDKind.GetfaultCard Is Nothing Then
            lng�����ID = IDKind.GetfaultCard.�ӿ����
        Else
            lng�����ID = "-1"
        End If
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng����ID
        strSQL = strSQL & " And A.����ID=[1]"
        blnHavePassWord = True
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '����ID
        strSQL = strSQL & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��
        strSQL = strSQL & " And A.סԺ��=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����
        strSQL = strSQL & " And A.�����=[1]"
    Else
        Select Case objCard.����
            Case "����", "��������￨"
                blnName = (mlng����ID > 0)
                Exit Function
            Case "ҽ����"
                strInput = UCase(strInput)
                strSQL = strSQL & " And A.ҽ����=[2]"
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.�����=[2]"
            Case Else
                '��������,��ȡ��صĲ���ID
                If Val(objCard.�ӿ����) > 0 Then
                    lng�����ID = Val(objCard.�ӿ����)
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng����ID = 0 Then GoTo NotFoundPati:
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng����ID <= 0 Then GoTo NotFoundPati:
                strSQL = strSQL & " And A.����ID=[1]"
                strInput = "-" & lng����ID
                blnHavePassWord = True
        End Select
    End If
    
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    
    blnDo = Not rsTmp.EOF
    
    If blnDo Then
        mlng����ID = rsTmp!����ID
        mlng����ID = 0
        GetPatient = True
    Else
NotFoundPati:
        mlng����ID = 0
        mlng����ID = 0
        Call ClearInfo
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Sub ClearInfo()
    mblnChange = False
    mstrAgeAndBirth = ""
    Set mrsTmp = New ADODB.Recordset
    txtPatient.Tag = ""
    txtPatient.Text = ""
    txtPatient.ForeColor = &H80000008
    lblName.Tag = txtPatient.ForeColor
    medBirthdayDate.Text = "____-__-__"
    medBirthdayTime.Text = "__:__"
    medBirthdayDate.Tag = ""
    txtAge.Text = ""
    txtInfo.Text = ""
    cmbNum.Clear
    optType(0).Value = True
    optType(0).Enabled = False
    optType(1).Enabled = False
    cmbNum.Enabled = False
    mblnChange = True
    mstrName = ""
End Sub

Private Function EMPI_LoadPati(ByVal str���� As String, ByVal str�Ա� As String, ByVal str�������� As String) As Boolean
'����:��EMPI�������Ĳ�����Ϣ���µ�����
    Dim rsPatiIn As ADODB.Recordset
    Dim rsPatiOut As ADODB.Recordset
    Dim blnRet As Boolean
    
    If CreatePlugInOK(glngModule) Then
        '��֯���˻�����Ϣ
        Set rsPatiIn = New ADODB.Recordset
        With rsPatiIn.Fields
            .Append "����ID", adBigInt
            .Append "��ҳID", adBigInt
            .Append "�Һ�ID", adBigInt
            '-------------------------------
            .Append "�����", adVarChar, 18
            .Append "סԺ��", adVarChar, 18
            .Append "ҽ����", adVarChar, 30
            .Append "���֤��", adVarChar, 18
            .Append "����֤��", adVarChar, 20
            .Append "����", adVarChar, 100
            .Append "�Ա�", adVarChar, 4
            .Append "��������", adVarChar, 20 '���ڸ�ʽ��YYYY-MM-DD HH:MM:SS
            .Append "�����ص�", adVarChar, 100
            .Append "����", adVarChar, 30
            .Append "����", adVarChar, 20
            .Append "ѧ��", adVarChar, 10
            .Append "ְҵ", adVarChar, 80
            .Append "������λ", adVarChar, 100
            .Append "����", adVarChar, 30
            .Append "����״��", adVarChar, 4
            .Append "��ͥ�绰", adVarChar, 20
            .Append "��ϵ�˵绰", adVarChar, 20
            .Append "��λ�绰", adVarChar, 20
            .Append "��ͥ��ַ", adVarChar, 100
            .Append "��ͥ��ַ�ʱ�", adVarChar, 6
            .Append "���ڵ�ַ", adVarChar, 100
            .Append "���ڵ�ַ�ʱ�", adVarChar, 6
            .Append "��λ�ʱ�", adVarChar, 6
            .Append "��ϵ�˵�ַ", adVarChar, 100
            .Append "��ϵ�˹�ϵ", adVarChar, 30
            .Append "��ϵ������", adVarChar, 64
        End With
        rsPatiIn.CursorLocation = adUseClient
        rsPatiIn.LockType = adLockOptimistic
        rsPatiIn.CursorType = adOpenStatic
        rsPatiIn.Open
         '1-����;2-סԺ(lng����ID=0,��Ĭ��Ϊ1;lng����ID<>0,1-lng����IDΪ�Һ�ID,2-lng����IDΪ��ҳID)
        With rsPatiIn
            .AddNew
            !����ID = mlng����ID
            !��ҳID = IIf(mlng����ID <> 0, IIf(mint���� = 2, mlng����ID, 0), 0)
            !�Һ�ID = IIf(mlng����ID <> 0, IIf(mint���� = 1, mlng����ID, 0), 0)
            !���� = str����
            !�Ա� = str�Ա�
            !�������� = str��������
            .Update
            '-------------------------------------------------------
        End With
        
        '���ò�ѯ�ӿ�
        On Error Resume Next
        blnRet = gobjPlugIn.EMPI_QueryPatiInfo(glngSys, glngModule, rsPatiIn, rsPatiOut)
        If Err.Number = 438 Then blnRet = False
        Call zlPlugInErrH(Err, "EMPI_QueryPatiInfo")
        Err.Clear: On Error GoTo 0
        If Not blnRet Then Exit Function
        If rsPatiOut Is Nothing Then Exit Function
        If rsPatiOut.RecordCount = 0 Then Exit Function
        EMPI_LoadPati = True      '���ڱ���ҵ���������
    End If
End Function

Private Sub RecalcBirthDay()
'����:ͨ�����䷴�Ƴ�������
    Dim strBirth As String
    
    If RecalcBirth(Trim(txtAge.Text) & IIf(cboAge.Visible, Trim(cboAge.Text), ""), strBirth) Then
        If medBirthdayDate.Enabled Then medBirthdayDate.Text = Format(strBirth, "YYYY-MM-DD")
        If medBirthdayTime.Enabled Then medBirthdayTime.Text = IIf(Format(strBirth, "HH:MM") = "00:00", "__:__", Format(strBirth, "HH:MM"))
    End If
End Sub
