VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAutoOut 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ժ�Ǽ�"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   Icon            =   "frmAutoOut.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraInfo 
      Height          =   2625
      Left            =   105
      TabIndex        =   15
      Top             =   45
      Width           =   6195
      Begin VB.ComboBox cbo���� 
         Height          =   300
         ItemData        =   "frmAutoOut.frx":020A
         Left            =   4850
         List            =   "frmAutoOut.frx":021D
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2130
         Width           =   1215
      End
      Begin VB.ComboBox cbo��Ժ��� 
         Height          =   300
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   4230
      End
      Begin VB.CheckBox chk���� 
         Alignment       =   1  'Right Justify
         Caption         =   "����"
         Height          =   195
         Left            =   2835
         TabIndex        =   8
         Top             =   2190
         Width           =   660
      End
      Begin VB.TextBox txt���� 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4170
         MaxLength       =   3
         TabIndex        =   9
         Top             =   2130
         Width           =   405
      End
      Begin VB.CheckBox chkʬ�� 
         Alignment       =   1  'Right Justify
         Caption         =   "ʬ��"
         Enabled         =   0   'False
         Height          =   195
         Left            =   5430
         TabIndex        =   6
         Top             =   1800
         Width           =   660
      End
      Begin VB.TextBox txt��Ժ��� 
         Height          =   300
         Left            =   960
         MaxLength       =   100
         TabIndex        =   0
         Top             =   240
         Width           =   5100
      End
      Begin VB.CheckBox chk���� 
         Alignment       =   1  'Right Justify
         Caption         =   "ȷ��"
         Height          =   195
         Left            =   2835
         TabIndex        =   5
         Top             =   1800
         Width           =   660
      End
      Begin VB.TextBox txt��ҽ��� 
         Height          =   300
         Left            =   960
         MaxLength       =   100
         TabIndex        =   2
         Top             =   960
         Width           =   5100
      End
      Begin VB.ComboBox cbo��ҽ��Ժ��� 
         Height          =   300
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1320
         Width           =   4230
      End
      Begin VB.ComboBox cbo��Ժ��ʽ 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1740
         Width           =   1830
      End
      Begin MSComCtl2.UpDown UD���� 
         Height          =   300
         Left            =   4590
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2130
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txt����"
         BuddyDispid     =   196613
         OrigLeft        =   3945
         OrigTop         =   645
         OrigRight       =   4185
         OrigBottom      =   930
         Max             =   99999
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   300
         Left            =   960
         TabIndex        =   7
         Top             =   2130
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   19
         Format          =   "yyyy-MM-dd hh:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtOkDate 
         Height          =   300
         Left            =   3600
         TabIndex        =   23
         Top             =   1760
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   19
         Format          =   "yyyy-MM-dd hh:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ���"
         Height          =   180
         Left            =   1035
         TabIndex        =   22
         Top             =   660
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժʱ��"
         Height          =   180
         Left            =   180
         TabIndex        =   21
         Top             =   2190
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   3780
         TabIndex        =   20
         Top             =   2190
         Width           =   360
      End
      Begin VB.Label lbl��Ժ��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ���"
         Height          =   180
         Left            =   180
         TabIndex        =   19
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lbl��ҽ��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ҽ���"
         Height          =   180
         Left            =   180
         TabIndex        =   18
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ���"
         Height          =   180
         Left            =   1035
         TabIndex        =   17
         Top             =   1380
         Width           =   720
      End
      Begin VB.Label lbl��Ժ��ʽ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ��ʽ"
         Height          =   180
         Left            =   180
         TabIndex        =   16
         Top             =   1800
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   225
      TabIndex        =   14
      Top             =   2850
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3870
      TabIndex        =   12
      Top             =   2850
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5070
      TabIndex        =   13
      Top             =   2850
      Width           =   1100
   End
End
Attribute VB_Name = "frmAutoOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mlng����ID As Long
Public mlng��ҳID As Long
Public mstr�Ա� As String
Public mint���� As Integer
Public mlngDepID As Long '��Ժ����ID
Private mdteDeathDate As Date
Private mintDeath As Integer

Private Sub cbo��Ժ��ʽ_Click()
    If InStr(cbo��Ժ��ʽ.Text, "����") > 0 Then
        txt����.Text = ""
        chk����.Value = 0

        txt����.Enabled = (chk����.Value = 1)
        UD����.Enabled = txt����.Enabled

        chk����.Enabled = False
    
        chkʬ��.Enabled = True
    Else
        chk����.Enabled = True
        
        chkʬ��.Value = 0
        chkʬ��.Enabled = False
    End If
End Sub

Private Sub cbo��Ժ���_Click()
    Dim i As Integer
    If InStr(cbo��Ժ���.Text, "����") > 0 Then
        i = cbo.FindIndex(cbo��Ժ��ʽ, "����", True)
        If i <> -1 Then cbo��Ժ��ʽ.ListIndex = i
    End If
End Sub

Private Sub cbo����_Click()
    txt����.Enabled = (cbo����.ItemData(cbo����.ListIndex) <> 9)
    UD����.Enabled = txt����.Enabled
End Sub

Private Sub cbo��ҽ��Ժ���_Click()
    Dim i As Integer
    If InStr(cbo��ҽ��Ժ���.Text, "����") > 0 Then
        i = cbo.FindIndex(cbo��Ժ��ʽ, "����", True)
        If i <> -1 Then cbo��Ժ��ʽ.ListIndex = i
    End If
End Sub

Private Sub chk����_Click()
    txt����.Enabled = (chk����.Value = 1)
    UD����.Enabled = txt����.Enabled
    cbo����.Enabled = txt����.Enabled
    Call zlCommFun.PressKey(vbKeyTab)
End Sub
'����28982 by lesfeng 2010-06-09
Private Sub chk����_Click()
    If chk����.Value = 1 Then
        txtOkDate.Enabled = True
    Else
        txtOkDate.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
ShowHelp "zl9InPatient", Me.hWnd, "frmOut"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Not Me.ActiveControl Is txt��Ժ��� _
            And Not Me.ActiveControl Is txt��ҽ��� Then
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") And Not (Me.ActiveControl Is txt��Ժ��� Or Me.ActiveControl Is txt��ҽ���) Then KeyAscii = 0      '��������п�����'��
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset, rsDiagnosis As ADODB.Recordset
    Dim rsPatiInfo As ADODB.Recordset
    Dim i As Long, strSQL As String
    Dim dMax As Date, intԭ�� As Integer
    Dim lng����ID As Long, str��� As String, str��Ժ��� As String, str��ҽ��Ժ��� As String
        
    On Error GoTo errH
    '����28612 by lesfeng 2010-07-05
    
    mintDeath = 0
    mdteDeathDate = GetdeathTime(mlng����ID, mlng��ҳID)
    '����31652 by lesfeng �Ӳ�����ҳֱ����ȡȷ������
    Set rsPatiInfo = GetPatiInfo(mlng����ID, mlng��ҳID)
    
    txtDate.Text = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    dMax = GetMaxDate(mlng����ID, mlng��ҳID, intԭ��)
    If intԭ�� = 10 Then
        txtDate.Text = Format(dMax + 1 / 24 / 60 / 60, "yyyy-MM-dd HH:mm:ss")
    Else
        If dMax > CDate(txtDate.Text) Then
            txtDate.Text = Format(dMax + 1 / 24 / 60 / 60, "yyyy-MM-dd HH:mm:ss")
        End If
    End If
    '����28612 by lesfeng 2010-07-05
    If mintDeath = 1 Then
        txtDate.Text = Format(mdteDeathDate, "yyyy-MM-dd HH:mm:ss")
    End If
    
    If mlngDepID <> 0 Then
        txt��ҽ���.Enabled = (InStr(1, "," & GetDepCharacter(mlngDepID) & ",", ",��ҽ��,") > 0)
        txt��ҽ���.ToolTipText = "ֻ�е��������ڿ��ҵ�����Ϊ��ҽ��ʱ������������ҽ���!"
        cbo��ҽ��Ժ���.Enabled = txt��ҽ���.Enabled
    End If
    
     '��ʾ������ϼ�¼
    Set rsDiagnosis = GetDiagnosticInfo(mlng����ID, mlng��ҳID, "1,11,2,12,3,13", "2,3")
    If Not rsDiagnosis Is Nothing Then
        'a.��ҽ���
        rsDiagnosis.Filter = "�������=3 and ��¼��Դ=3"            '��ȡ��ҳ����ĳ�Ժ���
        If Not rsDiagnosis.EOF Then
            txt��Ժ���.Text = NVL(rsDiagnosis!�������): txt��Ժ���.Tag = NVL(rsDiagnosis!����ID, rsDiagnosis!���ID & ";"): lbl��Ժ���.Tag = txt��Ժ���.Text
            str��Ժ��� = "" & rsDiagnosis!��Ժ���
            '����28982 by lesfeng 2010-06-09
            chk����.Value = IIf(Val("" & rsDiagnosis!�Ƿ�����) = 1, 0, 1)
        Else
            rsDiagnosis.Filter = "�������=2 and ��¼��Դ=2"        '��ȡ��Ժ�Ǽǵ���Ժ���
            If Not rsDiagnosis.EOF Then
                txt��Ժ���.Text = NVL(rsDiagnosis!�������): txt��Ժ���.Tag = NVL(rsDiagnosis!����ID, rsDiagnosis!���ID & ";"): lbl��Ժ���.Tag = txt��Ժ���.Text
            Else
                rsDiagnosis.Filter = "�������=1 and ��¼��Դ=2"    '���ȡ��Ժ�Ǽǵ��������
                If Not rsDiagnosis.EOF Then
                    txt��Ժ���.Text = NVL(rsDiagnosis!�������): txt��Ժ���.Tag = NVL(rsDiagnosis!����ID, rsDiagnosis!���ID & ";"): lbl��Ժ���.Tag = txt��Ժ���.Text
                End If
            End If
        End If
        
        'b.��ҽ���
        If txt��ҽ���.Enabled Then
            rsDiagnosis.Filter = "�������=13 and ��¼��Դ=3"            '��ȡ��ҳ����ĳ�Ժ���
            If Not rsDiagnosis.EOF Then
                txt��ҽ���.Text = NVL(rsDiagnosis!�������): txt��ҽ���.Tag = NVL(rsDiagnosis!����ID, rsDiagnosis!���ID & ";"): lbl��ҽ���.Tag = txt��ҽ���.Text
                str��ҽ��Ժ��� = "" & rsDiagnosis!��Ժ���
            Else
                rsDiagnosis.Filter = "�������=12 and ��¼��Դ=2"        '��ȡ��Ժ�Ǽǵ���Ժ���
                If Not rsDiagnosis.EOF Then
                    txt��ҽ���.Text = NVL(rsDiagnosis!�������): txt��ҽ���.Tag = NVL(rsDiagnosis!����ID, rsDiagnosis!���ID & ";"): lbl��ҽ���.Tag = txt��ҽ���.Text
                Else
                    rsDiagnosis.Filter = "�������=11 and ��¼��Դ=2"    '���ȡ��Ժ�Ǽǵ��������
                    If Not rsDiagnosis.EOF Then
                        txt��ҽ���.Text = NVL(rsDiagnosis!�������): txt��ҽ���.Tag = NVL(rsDiagnosis!����ID, rsDiagnosis!���ID & ";"): lbl��ҽ���.Tag = txt��ҽ���.Text
                    End If
                End If
            End If
        End If
    End If
    '����28982 by lesfeng 2010-06-09
    '����31652 by lesfeng �Ӳ�����ҳֱ����ȡȷ������
    If Not IsNull(rsPatiInfo!ȷ������) Then
        txtOkDate.Text = Format(rsPatiInfo!ȷ������, "yyyy-MM-dd HH:mm:ss")
        chk����.Value = IIf(Val("" & rsPatiInfo!�Ƿ�ȷ��) = 1, 1, 0)
        If chk����.Value = 0 Then chk����.Value = 1
        chk����.Enabled = False
        txtOkDate.Enabled = False
    End If
        
    '��Ժ���
    cbo��Ժ���.AddItem "": cbo��Ժ���.ListIndex = cbo��Ժ���.NewIndex
    If cbo��ҽ��Ժ���.Enabled Then cbo��ҽ��Ժ���.AddItem "": cbo��ҽ��Ժ���.ListIndex = cbo��ҽ��Ժ���.NewIndex
    strSQL = "Select ����,����,Nvl(ȱʡ��־,0) as ȱʡ From ���ƽ�� Order by ����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo��Ժ���.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 Then
                If txt��Ժ���.Text <> "" Then cbo��Ժ���.ListIndex = cbo��Ժ���.NewIndex
                cbo��Ժ���.ItemData(cbo��Ժ���.NewIndex) = 1
            End If
        
            If cbo��ҽ��Ժ���.Enabled Then
                cbo��ҽ��Ժ���.AddItem rsTmp!���� & "-" & rsTmp!����
                If rsTmp!ȱʡ = 1 Then
                    If txt��ҽ���.Text <> "" Then cbo��ҽ��Ժ���.ListIndex = cbo��ҽ��Ժ���.NewIndex
                    cbo��ҽ��Ժ���.ItemData(cbo��ҽ��Ժ���.NewIndex) = 1
                End If
            End If
            
            rsTmp.MoveNext
        Next
    End If
    Call zlControl.CboLocate(cbo��Ժ���, str��Ժ���)
    Call zlControl.CboLocate(cbo��ҽ��Ժ���, str��ҽ��Ժ���)
    
    
    '��Ժ��ʽ
    strSQL = "Select ����,����,Nvl(ȱʡ��־,0) as ȱʡ From ��Ժ��ʽ Order by ����"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo��Ժ��ʽ.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 Then cbo��Ժ��ʽ.ListIndex = cbo��Ժ��ʽ.NewIndex
            rsTmp.MoveNext
        Next
    End If
        
    cbo����.ListIndex = 1
    Call chk����_Click
    '����28982 by lesfeng 2010-06-09
    If chk����.Enabled Then Call chk����_Click
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, dMax As Date, strSQL As String, Curdate As Date, blnTrans As Boolean
    Dim lng��ҽ����ID As Long, lng��ҽ����ID As Long
    Dim lng��ҽ���ID As Long, lng��ҽ���ID As Long
    Dim int���� As Integer
    Dim strȷ������  As String
    Dim str��Ժʱ�� As String
    Dim strInfo As String
    Dim rsPatiInfo As ADODB.Recordset
    
    On Error GoTo errH
    
    '��Ժ���
    If Not zlControl.TxtCheckInput(txt��Ժ���, "��Ժ���", txt��Ժ���.MaxLength) Then Exit Sub
    If Not zlControl.TxtCheckInput(txt��ҽ���, "��ҽ���", txt��ҽ���.MaxLength, True) Then Exit Sub
    If mint���� <> 0 Then
        If gclsInsure.GetCapability(support����¼��������, mlng����ID, mint����) Then
            If txt��Ժ���.Text = "" Then
                MsgBox "����д�ò��˵ĳ�Ժ��ϣ�", vbInformation, gstrSysName
                txt��Ժ���.SetFocus: Exit Sub
            End If
        End If
    End If
    If txt��Ժ���.Text <> "" And cbo��Ժ���.Text = "" Then
        MsgBox "��ѡ���Ժ��ϵĳ�Ժ�����", vbInformation, gstrSysName
        cbo��Ժ���.SetFocus: Exit Sub
    End If
    If txt��ҽ���.Text <> "" And cbo��ҽ��Ժ���.Text = "" And cbo��ҽ��Ժ���.Enabled Then
        MsgBox "��ѡ����ҽ��Ժ��ϵĳ�Ժ�����", vbInformation, gstrSysName
        cbo��ҽ��Ժ���.SetFocus: Exit Sub
    End If
    
    If Not IsDate(txtDate.Text) Then
        MsgBox "��������ȷ�Ĳ��˳�Ժʱ�䣡", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    
    'ʱ�䲻�ܳ�����ǰʱ��̫��(һ��)
    Curdate = zlDatabase.Currentdate
    If CDate(txtDate.Text) > Curdate Then
        If CDate(txtDate.Text) - Curdate > 7 Then
            MsgBox "��Ժʱ��ȵ�ǰʱ���ù���,���飡", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        If MsgBox("��Ժʱ������˵�ǰϵͳʱ��,ȷʵҪ��Ժ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txtDate.SetFocus: Exit Sub
        End If
    End If
    
    dMax = GetMaxDate(mlng����ID, mlng��ҳID)
    If Format(txtDate.Text, "yyyyMMddHHmmss") <= Format(dMax, "yyyyMMddHHmmss") Then
        MsgBox "���˳�Ժʱ�������ڸò����ϴα䶯ʱ�� " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & " ��", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    
    dMax = GetLastAdviceTime(mlng����ID, mlng��ҳID)
    If Format(txtDate.Text, "yyyyMMddHHmmss") < Format(dMax, "yyyyMMddHHmmss") Then
        If MsgBox("��Ժʱ��С�ڸò��������Чҽ����ʱ�� " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & ",ȷʵҪ��Ժ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txtDate.SetFocus: Exit Sub
        End If
    End If
    '����28612 by lesfeng 2010-07-05
    If InStr(cbo��Ժ��ʽ.Text, "����") = 0 And mintDeath = 1 Then
        If MsgBox("�ò��˴�����Ч�ٴ�����ҽ��,������ҽ����ʱ�� " & Format(mdteDeathDate, "yyyy-MM-dd HH:mm:ss") & ",����Ժ��ʽ��Ϊ����,ȷʵҪ��Ժ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            cbo��Ժ��ʽ.SetFocus: Exit Sub
        End If
    End If
    
    If InStr(cbo��Ժ��ʽ.Text, "����") > 0 And mintDeath = 1 Then
        If Format(txtDate.Text, "yyyyMMddHHmmss") <> Format(mdteDeathDate, "yyyyMMddHHmmss") Then
            If MsgBox("��Ժʱ�䲻���ڸò�����Ч�ٴ�����ҽ����ʱ�� " & Format(mdteDeathDate, "yyyy-MM-dd HH:mm:ss") & ",ȷʵҪ��Ժ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                txtDate.SetFocus: Exit Sub
            End If
        End If
    End If
    '����32764 by lesfeng 2010-09-13 ���ֲ���22��32 ����154��155
    '��鲡���Ƿ���δִ����ɵ�������Ŀ��δ��ҩƷ
    If gbyt���δִ�� <> 0 Then
        strInfo = ExistWaitExe(mlng����ID, mlng��ҳID)
        If strInfo <> "" Then
            If gbyt���δִ�� = 1 Then
                If MsgBox("���ֲ��˴�����δִ����ɵ����ݣ�" & _
                    vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "ȷ��Ҫ��Ժ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            Else
                MsgBox "���ֲ��˴�����δִ����ɵ����ݣ�" & vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "�������Ժ.", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If

    If gbyt���δ��ҩ <> 0 Then
        strInfo = ExistWaitDrug(mlng����ID, mlng��ҳID)
        If strInfo <> "" Then
            If gbyt���δ��ҩ = 1 Then
                If MsgBox("���ֲ���" & strInfo & vbCrLf & vbCrLf & "ȷ��Ҫ��Ժ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            Else
                MsgBox "���ֲ���" & strInfo & vbCrLf & vbCrLf & "�������Ժ��", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    
    '����28982 by lesfeng 2010-06-09
    strȷ������ = ""
    If chk����.Value = 1 Then
        Set rsPatiInfo = GetPatiInfo(mlng����ID, mlng��ҳID)
        str��Ժʱ�� = Format(rsPatiInfo!��Ժʱ��, "yyyy-MM-dd HH:mm:ss")
        If Not IsDate(txtOkDate.Text) Then
            MsgBox "��������ȷ�Ĳ���ȷ��ʱ�䣡", vbInformation, gstrSysName
            If txtOkDate.Enabled Then txtOkDate.SetFocus: Exit Sub
        End If
        If Format(txtOkDate.Text, "yyyyMMddHHmmss") >= Format(txtDate.Text, "yyyyMMddHHmmss") Then
            MsgBox "ȷ��ʱ�����С�ڲ��˳�Ժʱ�� " & Format(txtDate.Text, "yyyy-MM-dd HH:mm:ss") & " ��", vbInformation, gstrSysName
            If txtOkDate.Enabled Then txtOkDate.SetFocus: Exit Sub
        End If
        
        If Format(str��Ժʱ��, "yyyyMMddHHmmss") > Format(txtOkDate.Text, "yyyyMMddHHmmss") Then
            MsgBox "ȷ��ʱ�������ڵ��ڲ�����Ժʱ�� " & Format(str��Ժʱ��, "yyyy-MM-dd HH:mm:ss") & " ��", vbInformation, gstrSysName
            If txtOkDate.Enabled Then txtOkDate.SetFocus: Exit Sub
        End If
        
        strȷ������ = Format(txtOkDate.Text, "yyyy-MM-dd HH:mm:ss")
    End If
    
    If InStr(1, txt��Ժ���.Tag, ";") <= 0 Then
        lng��ҽ����ID = Val(txt��Ժ���.Tag)
    Else
        lng��ҽ���ID = Val(txt��Ժ���.Tag)
    End If
    If InStr(1, txt��ҽ���.Tag, ";") <= 0 Then
        lng��ҽ����ID = Val(txt��ҽ���.Tag)
    Else
        lng��ҽ���ID = Val(txt��ҽ���.Tag)
    End If
    
    If cbo����.ListIndex <> -1 Then int���� = cbo����.ItemData(cbo����.ListIndex)
    '����28982 by lesfeng 2010-06-09
    strSQL = "zl_���˱䶯��¼_Out(" & mlng����ID & "," & mlng��ҳID & "," & _
            ZVal(lng��ҽ����ID) & "," & ZVal(lng��ҽ���ID) & ",'" & Replace(txt��Ժ���.Text, "'", "''") & "','" & zlStr.NeedName(cbo��Ժ���.Text) & "'," & _
            ZVal(lng��ҽ����ID) & "," & ZVal(lng��ҽ���ID) & ",'" & Replace(txt��ҽ���.Text, "'", "''") & "','" & zlStr.NeedName(cbo��ҽ��Ժ���.Text) & "'," & _
            chk����.Value & ",'" & zlStr.NeedName(cbo��Ժ��ʽ.Text) & "',To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS')," & _
            IIf(chk����.Value = 1, int����, 0) & "," & IIf(chk����.Value = 1 And int���� <> 9, Val(txt����.Text), "Null") & "," & IIf(chkʬ��.Enabled, chkʬ��.Value, "NULL") & "," & _
            "'" & UserInfo.��� & "','" & UserInfo.���� & "'," & IIf(strȷ������ = "", "NULL", "To_Date('" & strȷ������ & "','YYYY-MM-DD HH24:MI:SS')") & ")"
    
    gcnOracle.BeginTrans: blnTrans = True
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        If mint���� <> 0 Then 'ҽ���Ķ�
            If Not gclsInsure.LeaveSwap(mlng����ID, mlng��ҳID, mint����) Then
                gcnOracle.RollbackTrans: Exit Sub
            End If
        End If
    gcnOracle.CommitTrans: blnTrans = False
    
    '���˺�:24662
    Dim strOutPut As String
    'If Not mobjICCard Is Nothing Then
        Call zlExcuteUploadSwap(mlng����ID, strOutPut)
    'End If
    
    
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlng����ID = 0
    mlng��ҳID = 0
    mint���� = 0
    mstr�Ա� = ""
End Sub

Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_LostFocus()
    If Not IsDate(txtDate.Text) Then txtDate.SetFocus
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt��Ժ���_GotFocus()
    zlControl.TxtSelAll txt��Ժ���
End Sub

Private Sub txt��ҽ���_GotFocus()
    zlControl.TxtSelAll txt��ҽ���
End Sub

Private Sub txt��Ժ���_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInput As String, strSex As String, lngTxtHeight As Long, vPoint As POINTAPI
    
If KeyAscii = 13 Then
        KeyAscii = 0
        If Not RequestCode Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt��Ժ���.Text = lbl��Ժ���.Tag And txt��Ժ���.Text <> "" Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt��Ժ���.Text = "" Then
            txt��Ժ���.Tag = "": lbl��Ժ���.Tag = ""
            Call zlCommFun.PressKey(vbKeyTab) '��������
        Else
            vPoint = zlControl.GetCoordPos(fraInfo.hWnd, txt��Ժ���.Left, txt��Ժ���.Top)
            strInput = UCase(txt��Ժ���.Text)
            strSex = mstr�Ա�
            lngTxtHeight = txt��Ժ���.Height
            Set rsTmp = GetDiseaseCode(Me, blnCancel, strInput, strSex, "D", vPoint.X, vPoint.Y, lngTxtHeight)
            If Not rsTmp Is Nothing Then
                txt��Ժ���.Tag = rsTmp!ID
                txt��Ժ���.Text = "(" & rsTmp!���� & ")" & rsTmp!����
                lbl��Ժ���.Tag = txt��Ժ���.Text '���ڻָ���ʾ
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                If Not blnCancel Then
                    MsgBox "û���ҵ�ƥ��ļ������롣", vbInformation, gstrSysName
                End If
                If lbl��Ժ���.Tag <> "" Then txt��Ժ���.Text = lbl��Ժ���.Tag
                Call txt��Ժ���_GotFocus
                txt��Ժ���.SetFocus
            End If
        End If
    Else
        CheckInputLen txt��Ժ���, KeyAscii
    End If
End Sub

Private Sub txt��ҽ���_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInput As String, strSex As String, lngTxtHeight As Long, vPoint As POINTAPI
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not RequestCode Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt��ҽ���.Text = lbl��ҽ���.Tag And txt��ҽ���.Text <> "" Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt��ҽ���.Text = "" Then
            txt��ҽ���.Tag = "": lbl��ҽ���.Tag = ""
            Call zlCommFun.PressKey(vbKeyTab) '��������
        Else
            vPoint = zlControl.GetCoordPos(fraInfo.hWnd, txt��ҽ���.Left, txt��ҽ���.Top)
            strInput = UCase(txt��ҽ���.Text)
            strSex = mstr�Ա�
            lngTxtHeight = txt��ҽ���.Height
            Set rsTmp = GetDiseaseCode(Me, blnCancel, strInput, strSex, "B", vPoint.X, vPoint.Y, lngTxtHeight)
            
            If Not rsTmp Is Nothing Then
                txt��ҽ���.Tag = rsTmp!ID
                txt��ҽ���.Text = "(" & rsTmp!���� & ")" & rsTmp!����
                lbl��ҽ���.Tag = txt��ҽ���.Text '���ڻָ���ʾ
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                If Not blnCancel Then
                    MsgBox "û���ҵ�ƥ��ļ������롣", vbInformation, gstrSysName
                End If
                If lbl��ҽ���.Tag <> "" Then txt��ҽ���.Text = lbl��ҽ���.Tag
                Call txt��ҽ���_GotFocus
                txt��ҽ���.SetFocus
                
            End If
        End If
    Else
        CheckInputLen txt��ҽ���, KeyAscii
    End If
End Sub

Private Sub txt��Ժ���_Validate(Cancel As Boolean)
    If Val(txt��Ժ���.Tag) > 0 And txt��Ժ���.Text <> lbl��Ժ���.Tag Then
        txt��Ժ���.Text = lbl��Ժ���.Tag
    ElseIf Val(txt��Ժ���.Tag) = 0 And RequestCode Then
        txt��Ժ���.Text = ""
    End If
    
    If txt��Ժ���.Text <> "" And cbo��Ժ���.Text = "" Then
        cbo��Ժ���.ListIndex = cbo.FindIndex(cbo��Ժ���, 1)
        If cbo��Ժ���.ListIndex = -1 Then cbo��Ժ���.ListIndex = 0
    ElseIf txt��Ժ���.Text = "" And cbo��Ժ���.Text <> "" Then
        cbo��Ժ���.ListIndex = 0
    End If
End Sub

Private Sub txt��ҽ���_Validate(Cancel As Boolean)
    If Val(txt��ҽ���.Tag) > 0 And txt��ҽ���.Text <> lbl��ҽ���.Tag Then
        txt��ҽ���.Text = lbl��ҽ���.Tag
    ElseIf Val(txt��ҽ���.Tag) = 0 And RequestCode Then
        txt��ҽ���.Text = ""
    End If
    
    If txt��ҽ���.Text <> "" And cbo��ҽ��Ժ���.Text = "" Then
        cbo��ҽ��Ժ���.ListIndex = cbo.FindIndex(cbo��ҽ��Ժ���, 1)
        If cbo��ҽ��Ժ���.ListIndex = -1 Then cbo��ҽ��Ժ���.ListIndex = 0
    ElseIf txt��ҽ���.Text = "" And cbo��ҽ��Ժ���.Text <> "" Then
        cbo��ҽ��Ժ���.ListIndex = 0
    End If
End Sub

Private Function RequestCode() As Boolean
    RequestCode = gint������� = 2 Or (gint������� = 3 And mint���� <> 0)
End Function

Public Function GetFindIllWhere(ByVal strInput As String, ByVal str���� As String) As String
'����:��ò��Ҽ�������Ŀ¼������
    Dim strWhere As String
    
    If zlCommFun.IsCharChinese(strInput) Then
        strWhere = str���� & ".���� like '" & gstrLike & strInput & "%'"
    Else
        strWhere = "(" & str���� & ".���� Like '" & strInput & "%'" & _
                " Or " & str���� & ".���� Like '" & gstrLike & strInput & "%'" & _
                " Or " & str���� & ".���� Like '" & gstrLike & strInput & "%')"
    End If
    GetFindIllWhere = strWhere
End Function
'����28982 by lesfeng 2010-06-09
Private Function GetPatiInfo(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
'���ܣ���ȡ���������Ϣ
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    '����31652 by lesfeng �Ӳ�����ҳֱ����ȡȷ�����ڣ������Ƿ�ȷ�Ｐȷ������
    strSQL = "" & _
        "   Select  nvl(B.����,A.����) as ����, nvl(B.�Ա�,Nvl(A.�Ա�,'δ֪')) as  �Ա�, B.����, B.����, B.��������, B.��ǰ����, B.����ȼ�id, B.סԺҽʦ, B.����ҽʦ, B.���λ�ʿ, B.��Ժ����id, B.��Ժ����id ��ס����id," & vbNewLine & _
        "           To_Char(B.��Ժ����, 'YYYY-MM-DD HH24:MI:SS') As ��Ժʱ��, B.��ǰ����id, B.סԺ��, D.���� as ��ǰ����, B.��Ժ���� as ��Ҫ����,B.�Ƿ�ȷ��,B.ȷ������" & vbNewLine & _
        "   From ������Ϣ A, ������ҳ B, ���ű� D" & vbNewLine & _
        "   Where A.����id = B.����id And B.����id = [1] And B.��ҳid = [2] And B.��Ժ����id = D.id"
   
    Set GetPatiInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
'����28612 by lesfeng 2010-07-05
Private Function GetdeathTime(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Date
'���ܣ���ȡָ�������Ƿ��������ҽ�������ڳ�Ժʱ��Ϊ����ʱ���1��
'˵�������ڻ�ȡ��������ʱ��Ϊ��Ժʱ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    GetdeathTime = CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"))
    
    On Error GoTo errH
    
    strSQL = "Select Max(Nvl(A.ִ����ֹʱ��, Nvl(A.�ϴ�ִ��ʱ��, A.��ʼִ��ʱ��)) + 1 / 24 / 60 ) As ʱ�� " & _
             "  From ����ҽ����¼ A, ������ĿĿ¼ B " & _
             " Where A.������� = 'Z' And A.������Ŀid = B.ID And B.�������� = 11 And B.��� = 'Z' And A.ҽ��״̬ In (3, 8, 9) And " & _
             "       A.����ID = [1] And A.��ҳID = [2] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
    
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!ʱ��) Then
            GetdeathTime = rsTmp!ʱ��
            mintDeath = 1
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


