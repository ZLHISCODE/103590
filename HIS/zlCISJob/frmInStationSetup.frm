VERSION 5.00
Begin VB.Form frmInStationSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "frmInStationSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraAdvice 
      Caption         =   "�������� "
      Height          =   2340
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4800
      Begin VB.CheckBox chkWarn 
         Caption         =   "Ѫ������"
         Height          =   195
         Index           =   12
         Left            =   465
         TabIndex        =   55
         Top             =   1710
         Width           =   1020
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "��Ѫ���"
         Height          =   195
         Index           =   11
         Left            =   3750
         TabIndex        =   53
         Top             =   1455
         Width           =   1020
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "�걾����"
         Height          =   195
         Index           =   10
         Left            =   2740
         TabIndex        =   52
         Top             =   1455
         Width           =   1040
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "ȡѪ֪ͨ"
         Height          =   195
         Index           =   9
         Left            =   1740
         TabIndex        =   50
         Top             =   1455
         Width           =   1040
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "RISԤԼ׼��"
         Height          =   195
         Index           =   8
         Left            =   465
         TabIndex        =   43
         Top             =   1455
         Width           =   1365
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "RISԤԼ"
         Height          =   195
         Index           =   7
         Left            =   3615
         TabIndex        =   42
         Top             =   1185
         Width           =   1035
      End
      Begin VB.CheckBox chkSoundHS 
         Caption         =   "����������ʾ"
         Height          =   195
         Left            =   300
         TabIndex        =   40
         Top             =   1980
         Width           =   1470
      End
      Begin VB.CommandButton cmdSoundHSSet 
         Caption         =   "��������(&S)"
         Height          =   350
         Left            =   1830
         TabIndex        =   39
         Top             =   1890
         Width           =   1410
      End
      Begin VB.TextBox txtNotifyAdvice 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   795
         MaxLength       =   3
         TabIndex        =   18
         Text            =   "10"
         Top             =   315
         Width           =   300
      End
      Begin VB.Frame fraNotifyAdvice 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   780
         TabIndex        =   17
         Top             =   495
         Width           =   300
      End
      Begin VB.Frame fraNotifyAdviceDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   780
         TabIndex        =   16
         Top             =   765
         Width           =   300
      End
      Begin VB.TextBox txtNotifyAdviceDay 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   795
         MaxLength       =   2
         TabIndex        =   15
         Text            =   "1"
         Top             =   585
         Width           =   300
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "�¿�"
         Height          =   195
         Index           =   0
         Left            =   1125
         TabIndex        =   14
         Top             =   915
         Width           =   675
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "��ͣ"
         Height          =   195
         Index           =   1
         Left            =   1875
         TabIndex        =   13
         Top             =   915
         Width           =   675
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "�·�"
         Height          =   195
         Index           =   2
         Left            =   2700
         TabIndex        =   12
         Top             =   915
         Width           =   660
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "����"
         Height          =   195
         Index           =   3
         Left            =   3495
         TabIndex        =   11
         Top             =   915
         Width           =   675
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "Σ��ֵ"
         Height          =   195
         Index           =   4
         Left            =   465
         TabIndex        =   10
         Top             =   1185
         Width           =   870
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "��Һ�ܾ�"
         Height          =   195
         Index           =   5
         Left            =   1395
         TabIndex        =   9
         Top             =   1185
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "��������"
         Height          =   195
         Index           =   6
         Left            =   2535
         TabIndex        =   8
         Top             =   1185
         Width           =   1035
      End
      Begin VB.CheckBox chkNotifyAdvice 
         Caption         =   "ÿ    �����Զ�ˢ��ҽ�����������е�����"
         Height          =   195
         Left            =   300
         TabIndex        =   19
         Top             =   330
         Width           =   3900
      End
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         Height          =   180
         Left            =   300
         TabIndex        =   21
         Top             =   915
         Width           =   810
      End
      Begin VB.Label lblNotifyAdviceDay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ���ڴ����ҽ��������ʾ����������"
         Height          =   180
         Left            =   570
         TabIndex        =   20
         Top             =   600
         Width           =   3420
      End
   End
   Begin VB.Frame fraEPR 
      Caption         =   "��������"
      Height          =   2340
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Width           =   4800
      Begin VB.CheckBox chkΣ��ֵ 
         Caption         =   "Σ��ֵ��������"
         Height          =   240
         Left            =   120
         TabIndex        =   56
         Top             =   1975
         Width           =   1590
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "��Ѫ��Ӧ"
         Height          =   195
         Index           =   26
         Left            =   3720
         TabIndex        =   54
         Top             =   1635
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "��Ѫ���"
         Height          =   195
         Index           =   25
         Left            =   2565
         TabIndex        =   51
         Top             =   1635
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "У������"
         Height          =   195
         Index           =   24
         Left            =   1440
         TabIndex        =   45
         Top             =   1635
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "��Ѫ���"
         Height          =   195
         Index           =   23
         Left            =   3720
         TabIndex        =   44
         Top             =   1380
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "�����ʿ�"
         Height          =   195
         Index           =   22
         Left            =   2565
         TabIndex        =   41
         Top             =   1380
         Width           =   1035
      End
      Begin VB.CheckBox chkSoundYS 
         Caption         =   "����������ʾ"
         Height          =   195
         Left            =   1800
         TabIndex        =   38
         Top             =   1998
         Width           =   1455
      End
      Begin VB.CommandButton cmdSoundYSSet 
         Caption         =   "��������(&S)"
         Height          =   350
         Left            =   3240
         TabIndex        =   37
         Top             =   1920
         Width           =   1410
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "��Ⱦ��"
         Height          =   195
         Index           =   21
         Left            =   1440
         TabIndex        =   36
         Top             =   1380
         Width           =   855
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "�������"
         Height          =   195
         Index           =   20
         Left            =   3720
         TabIndex        =   35
         Top             =   1125
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "ҽ�����"
         Height          =   195
         Index           =   19
         Left            =   2565
         TabIndex        =   34
         Top             =   1125
         Width           =   1035
      End
      Begin VB.TextBox txtNotifyEPR 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   840
         MaxLength       =   3
         TabIndex        =   30
         Text            =   "10"
         Top             =   270
         Width           =   300
      End
      Begin VB.Frame fraNotifyEPR 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   825
         TabIndex        =   29
         Top             =   510
         Width           =   300
      End
      Begin VB.Frame fraNotifyEPRDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   825
         TabIndex        =   28
         Top             =   780
         Width           =   300
      End
      Begin VB.TextBox txtNotifyEPRDay 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   840
         MaxLength       =   2
         TabIndex        =   27
         Text            =   "1"
         Top             =   600
         Width           =   300
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "��������"
         Height          =   195
         Index           =   15
         Left            =   1440
         TabIndex        =   26
         Top             =   885
         Width           =   1065
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "ҽ������"
         Height          =   195
         Index           =   16
         Left            =   2565
         TabIndex        =   25
         Top             =   885
         Width           =   1020
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "Σ��ֵ"
         Height          =   195
         Index           =   17
         Left            =   3720
         TabIndex        =   24
         Top             =   885
         Width           =   885
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "���泷��"
         Height          =   195
         Index           =   18
         Left            =   1440
         TabIndex        =   23
         Top             =   1125
         Width           =   1035
      End
      Begin VB.CheckBox chkNotifyEPR 
         Caption         =   "ÿ    �����Զ�ˢ�����������е�����"
         Height          =   195
         Left            =   360
         TabIndex        =   31
         Top             =   280
         Width           =   3900
      End
      Begin VB.Label Label�������� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         Height          =   180
         Left            =   600
         TabIndex        =   33
         Top             =   880
         Width           =   810
      End
      Begin VB.Label lblNotifyEPRDay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ������ɵ�������ʾ����������"
         Height          =   180
         Left            =   615
         TabIndex        =   32
         Top             =   615
         Width           =   3060
      End
   End
   Begin VB.Frame fraסԺ���Ƶ���ӡ 
      Caption         =   "ҽ�����ͺ�,���Ƶ���"
      Height          =   630
      Left            =   120
      TabIndex        =   46
      Top             =   2520
      Width           =   4815
      Begin VB.OptionButton optסԺ���Ƶ���ӡ 
         Caption         =   "����ӡ"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   49
         Top             =   300
         Width           =   840
      End
      Begin VB.OptionButton optסԺ���Ƶ���ӡ 
         Caption         =   "ѡ���Ƿ��ӡ"
         Height          =   180
         Index           =   1
         Left            =   1320
         TabIndex        =   48
         Top             =   300
         Width           =   1440
      End
      Begin VB.OptionButton optסԺ���Ƶ���ӡ 
         Caption         =   "�Զ���ӡ"
         Height          =   180
         Index           =   2
         Left            =   2880
         TabIndex        =   47
         Top             =   300
         Value           =   -1  'True
         Width           =   1080
      End
   End
   Begin VB.Frame fraBaby 
      Caption         =   "ҽ������ȱʡ��Χ(������)"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   4815
      Begin VB.OptionButton optBaby 
         Caption         =   "����ҽ��"
         Height          =   180
         Index           =   1
         Left            =   1440
         TabIndex        =   6
         Top             =   285
         Width           =   1200
      End
      Begin VB.OptionButton optBaby 
         Caption         =   "ȫ��ҽ��"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   285
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton optBaby 
         Caption         =   "Ӥ��ҽ��"
         Height          =   180
         Index           =   2
         Left            =   2880
         TabIndex        =   4
         Top             =   285
         Width           =   1440
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   5055
      TabIndex        =   0
      Top             =   3915
      Width           =   5055
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   3720
         TabIndex        =   2
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   2535
         TabIndex        =   1
         Top             =   120
         Width           =   1100
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   8040
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   8040
         Y1              =   15
         Y2              =   0
      End
   End
End
Attribute VB_Name = "frmInStationSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mbln��ʿվ As Boolean
Public mstrPrivs As String
Private mlngModual As Long

Private Enum Enum_chkWarn
    '��ʿվ���Ѳ���
    chkN�¿� = 0
    chkN��ͣ = 1
    chkN�·� = 2
    chkN���� = 3
    chkNΣ��ֵ = 4
    chkN��Һ�ܾ� = 5
    chkN�������� = 6
    chkNRISԤԼ = 7
    chkNRISԤԼ׼�� = 8
    chkȡѪ֪ͨ = 9
    chk�걾���� = 10
    chk��Ѫ��� = 11
    chkѪ������ = 12
    
    
    'ҽ��վ���Ѳ���
    chkD�������� = 15
    chkDҽ������ = 16
    chkDΣ��ֵ = 17
    chkD���泷�� = 18
    chkDҽ����� = 19
    chkD������� = 20
    chkD��Ⱦ�� = 21
    chkD�����ʿ� = 22
    chkD��Ѫ��� = 23
    chkDУ������ = 24
    chkD��Ѫ��� = 25
    chkD��Ѫ��Ӧ = 26
End Enum

Public Sub ShowMe()
    '���°�סԺ��ʿ����վ���ã���ʾ��ע��ť
    Me.Show vbModal
End Sub


Private Sub chkNotifyAdvice_Click()
    txtNotifyAdvice.Enabled = chkNotifyAdvice.Value = 1
    If Visible And txtNotifyAdvice.Enabled Then txtNotifyAdvice.SetFocus
End Sub

Private Sub chkNotifyEPR_Click()
    txtNotifyEPR.Enabled = chkNotifyEPR.Value = 1
    If Visible And txtNotifyEPR.Enabled Then txtNotifyEPR.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim curDate As Date
    Dim strTmp As String
    Dim i As Integer
    Dim blnSetup As Boolean
    
    If mbln��ʿվ Then
        If chkNotifyAdvice.Value = 1 And Val(txtNotifyAdvice.Text) = 0 Then
            If txtNotifyAdvice.Text = "" Then
                MsgBox "������ҽ�����ѵ��Զ�ˢ�¼����", vbInformation, gstrSysName
            Else
                MsgBox "ҽ�����ѵ��Զ�ˢ�¼������ӦΪ1���ӡ�", vbInformation, gstrSysName
            End If
            txtNotifyAdvice.SetFocus: Exit Sub
        End If
        If Val(txtNotifyAdviceDay.Text) = 0 Then
            If txtNotifyAdviceDay.Text = "" Then
                MsgBox "������Ҫ���ѵ�ҽ��������", vbInformation, gstrSysName
            Else
                MsgBox "Ҫ���ѵ�ҽ����������ӦΪ1�졣", vbInformation, gstrSysName
            End If
            txtNotifyAdviceDay.SetFocus: Exit Sub
        End If
    Else
        If chkNotifyEPR.Value = 1 And Val(txtNotifyEPR.Text) = 0 Then
            If txtNotifyEPR.Text = "" Then
                MsgBox "�����ò����������ѵ��Զ�ˢ�¼����", vbInformation, gstrSysName
            Else
                MsgBox "�����������ѵ��Զ�ˢ�¼������ӦΪ1���ӡ�", vbInformation, gstrSysName
            End If
            txtNotifyEPR.SetFocus: Exit Sub
        End If
        
        If Val(txtNotifyEPRDay.Text) = 0 Then
            If txtNotifyEPRDay.Text = "" Then
                MsgBox "������Ҫ�������ĵĲ������������", vbInformation, gstrSysName
            Else
                MsgBox "Ҫ�������ĵĲ��������������ӦΪ1�졣", vbInformation, gstrSysName
            End If
            txtNotifyEPRDay.SetFocus: Exit Sub
        End If
    End If
    

    blnSetup = InStr(";" & mstrPrivs & ";", ";��������;") > 0

    '�Զ�ˢ��ҽ������
    If mbln��ʿվ Then
        Call zlDatabase.SetPara("�Զ�ˢ��ҽ�����", IIf(chkNotifyAdvice.Value = 1, Val(txtNotifyAdvice.Text), ""), glngSys, pסԺ��ʿվ, blnSetup)
        Call zlDatabase.SetPara("�Զ�ˢ��ҽ������", Val(txtNotifyAdviceDay.Text), glngSys, pסԺ��ʿվ, blnSetup)
        strTmp = ""
        For i = chkN�¿� To chkѪ������
            strTmp = strTmp & chkWarn(i).Value
        Next
        Call zlDatabase.SetPara("�Զ�ˢ��ҽ������", strTmp, glngSys, pסԺ��ʿվ, blnSetup)
        Call zlDatabase.SetPara("ҽ������Χ", IIf(optBaby(0).Value, 0, IIf(optBaby(1).Value, 1, 2)), glngSys, pסԺҽ������, blnSetup)
        Call zlDatabase.SetPara("����������ʾ", chkSoundHS.Value, glngSys, pסԺ��ʿվ, blnSetup)
    Else
        Call zlDatabase.SetPara("�Զ�ˢ�²������ļ��", IIf(chkNotifyEPR.Value = 1, Val(txtNotifyEPR.Text), ""), glngSys, pסԺҽ��վ, blnSetup)
        Call zlDatabase.SetPara("�Զ�ˢ�²�����������", Val(txtNotifyEPRDay.Text), glngSys, pסԺҽ��վ, blnSetup)
        strTmp = ""
        For i = chkD�������� To chkD��Ѫ��Ӧ
            strTmp = strTmp & chkWarn(i).Value
        Next
        Call zlDatabase.SetPara("�Զ�ˢ������", strTmp, glngSys, pסԺҽ��վ, blnSetup)
        Call zlDatabase.SetPara("����������ʾ", chkSoundYS.Value, glngSys, pסԺҽ��վ, blnSetup)
    End If
    
    Call zlDatabase.SetPara("סԺΣ��ֵ��������", chkΣ��ֵ.Value, glngSys, pסԺҽ��վ, blnSetup)
    
    Call zlDatabase.SetPara("סԺ���͵��ݴ�ӡ", IIf(optסԺ���Ƶ���ӡ(0).Value, 0, IIf(optסԺ���Ƶ���ӡ(1).Value, 1, 2)), glngSys, pסԺҽ������, blnSetup)
    gblnOK = True
    Unload Me
End Sub

Private Sub cmdSoundYSSet_Click()
'ҽ��
    Call frmMsgCallSetup.ShowMe(Me, 1)
End Sub

Private Sub cmdSoundHSSet_Click()
'��ʿ
    Call frmMsgCallSetup.ShowMe(Me, 2)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strPar As String, i As Long
    Dim curDate As Date, intDay As Integer
    Dim intType As Integer
    Dim strNotify As String
    
    gblnOK = False
    mlngModual = IIf(mbln��ʿվ, pסԺ��ʿվ, pסԺҽ��վ)
    If mbln��ʿվ Then
        fraAdvice.Visible = True
        fraBaby.Visible = True
        fraEPR.Visible = False
    Else
        fraAdvice.Visible = False
        fraBaby.Visible = False
        fraEPR.Visible = True
        i = fraBaby.Height + 60
    End If
    Me.Height = Me.Height - i
            
    optסԺ���Ƶ���ӡ(Val(zlDatabase.GetPara("סԺ���͵��ݴ�ӡ", glngSys, pסԺҽ������, "0", Array(optסԺ���Ƶ���ӡ(0), optסԺ���Ƶ���ӡ(1), optסԺ���Ƶ���ӡ(2)), InStr(mstrPrivs, "��������") > 0))).Value = True
    chkWarn(chkȡѪ֪ͨ).Visible = gblnѪ��ϵͳ
    chkWarn(chk��Ѫ���).Visible = gblnѪ��ϵͳ
    chkWarn(chkѪ������).Visible = gblnѪ��ϵͳ
    
    'Σ��ֵ��������
    chkΣ��ֵ.Value = Val(zlDatabase.GetPara("סԺΣ��ֵ��������", glngSys, pסԺҽ��վ, "1", Array(chkΣ��ֵ), intType))

    '�Զ�ˢ��ҽ������
    If mbln��ʿվ Then
        strPar = zlDatabase.GetPara("�Զ�ˢ��ҽ�����", glngSys, mlngModual, , Array(chkNotifyAdvice), InStr(mstrPrivs, "��������") > 0, intType)
        If Val(strPar) > 0 Then
            chkNotifyAdvice.Value = 1: txtNotifyAdvice.Text = Val(strPar)
        End If
        'ǰ���¼��л��Զ����ã���˺���ǿ������
        If (intType = 3 Or intType = 15) And InStr(mstrPrivs, "��������") = 0 Then
            txtNotifyAdvice.Enabled = False
        End If
        
        strPar = zlDatabase.GetPara("�Զ�ˢ��ҽ������", glngSys, mlngModual, 1, Array(lblNotifyAdviceDay, txtNotifyAdviceDay), InStr(mstrPrivs, "��������") > 0)
        txtNotifyAdviceDay.Text = Val(strPar)
        
        strPar = zlDatabase.GetPara("�Զ�ˢ��ҽ������", glngSys, mlngModual, "000000000000", Array(lbl��������, chkWarn(0), chkWarn(1), chkWarn(2), chkWarn(3), chkWarn(4), chkWarn(5), chkWarn(6), chkWarn(7), chkWarn(8), chkWarn(9), chkWarn(10), chkWarn(11), chkWarn(12)), InStr(mstrPrivs, "��������") > 0)
        For i = 1 To Len(strPar)
            chkWarn(i - 1).Value = IIf(Val(Mid(strPar, i, 1)) = 1, 1, 0)
        Next
    
        optBaby(Val(zlDatabase.GetPara("ҽ������Χ", glngSys, pסԺҽ������, "0", Array(optBaby(0), optBaby(1), optBaby(2)), InStr(mstrPrivs, "��������") > 0))).Value = True
        
        chkSoundHS.Value = Val(zlDatabase.GetPara("����������ʾ", glngSys, mlngModual, "1", Array(chkSoundHS, cmdSoundHSSet), InStr(mstrPrivs, "��������") > 0))
        
    Else
        strPar = zlDatabase.GetPara("�Զ�ˢ�²������ļ��", glngSys, mlngModual, , Array(chkNotifyEPR), InStr(mstrPrivs, "��������") > 0, intType)
        If Val(strPar) > 0 Then
            chkNotifyEPR.Value = 1: txtNotifyEPR.Text = Val(strPar)
        End If
        'ǰ���¼��л��Զ����ã���˺���ǿ������
        If (intType = 3 Or intType = 15) And InStr(mstrPrivs, "��������") = 0 Then
            txtNotifyEPR.Enabled = False
        End If
        
        strPar = zlDatabase.GetPara("�Զ�ˢ�²�����������", glngSys, mlngModual, 1, Array(lblNotifyEPRDay, txtNotifyEPRDay), InStr(mstrPrivs, "��������") > 0)
        txtNotifyEPRDay.Text = Val(strPar)
       
        strNotify = zlDatabase.GetPara("�Զ�ˢ������", glngSys, pסԺҽ��վ, , Array(chkWarn(15), chkWarn(16), chkWarn(17), chkWarn(18), chkWarn(19), chkWarn(20), chkWarn(21), chkWarn(22), chkWarn(23), chkWarn(24), chkWarn(25), Label��������), InStr(mstrPrivs, "��������") > 0)
        chkWarn(chkD��������).Value = Val(Mid(strNotify, 1, 1))
        chkWarn(chkDҽ������).Value = Val(Mid(strNotify, 2, 1))
        chkWarn(chkDΣ��ֵ).Value = Val(Mid(strNotify, 3, 1))
        chkWarn(chkD���泷��).Value = Val(Mid(strNotify, 4, 1))
        chkWarn(chkDҽ�����).Value = Val(Mid(strNotify, 5, 1))
        chkWarn(chkD�������).Value = Val(Mid(strNotify, 6, 1))
        chkWarn(chkD��Ⱦ��).Value = Val(Mid(strNotify, 7, 1))
        chkWarn(chkD�����ʿ�).Value = Val(Mid(strNotify, 8, 1))
        chkWarn(chkD��Ѫ���).Value = Val(Mid(strNotify, 9, 1))
        chkWarn(chkD��Ѫ���).Visible = gblnѪ��ϵͳ
        chkWarn(chkDУ������).Value = Val(Mid(strNotify, 10, 1))
        chkWarn(chkD��Ѫ���).Value = Val(Mid(strNotify, 11, 1))
        chkWarn(chkD��Ѫ���).Visible = gblnѪ��ϵͳ
        chkWarn(chkD��Ѫ��Ӧ).Value = Val(Mid(strNotify, 12, 1))
        chkWarn(chkD��Ѫ��Ӧ).Visible = gblnѪ��ϵͳ
        
        If InStr(mstrPrivs, "��������") = 0 Then
            chkWarn(chkD��������).Enabled = False
            chkWarn(chkDҽ������).Enabled = False
            chkWarn(chkDΣ��ֵ).Enabled = False
            chkWarn(chkD���泷��).Enabled = False
            chkWarn(chkDҽ�����).Enabled = False
            chkWarn(chkD�������).Enabled = False
            chkWarn(chkD��Ⱦ��).Enabled = False
            chkWarn(chkD�����ʿ�).Enabled = False
            chkWarn(chkD��Ѫ���).Enabled = False
            chkWarn(chkDУ������).Enabled = False
            chkWarn(chkD��Ѫ���).Enabled = False
            chkWarn(chkD��Ѫ��Ӧ).Enabled = False
        End If
        
        chkSoundYS.Value = Val(zlDatabase.GetPara("����������ʾ", glngSys, mlngModual, "1", Array(chkSoundYS, cmdSoundYSSet), InStr(mstrPrivs, "��������") > 0))
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbln��ʿվ = False
End Sub

Private Sub txtNotifyAdvice_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyAdvice)
End Sub

Private Sub txtNotifyAdvice_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNotifyEPR_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyEPR)
End Sub

Private Sub txtNotifyEPR_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNotifyAdviceDay_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyAdviceDay)
End Sub

Private Sub txtNotifyAdviceDay_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNotifyEPRDay_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyEPRDay)
End Sub

Private Sub txtNotifyEPRDay_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

