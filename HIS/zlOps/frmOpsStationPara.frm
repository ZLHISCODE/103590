VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOpsStationPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   5760
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5925
   Icon            =   "frmOpsStationPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   150
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5265
      Width           =   1100
   End
   Begin TabDlg.SSTab tbs 
      Height          =   5085
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   8969
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&1.���� "
      TabPicture(0)   =   "frmOpsStationPara.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl(8)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txt"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "udn"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdDeviceSetup"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chk"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.CheckBox chk 
         Caption         =   "����ҽ������ʱ���Զ����ɷ���"
         Height          =   255
         Left            =   135
         TabIndex        =   31
         Top             =   4320
         Width           =   3045
      End
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "�豸����(&S)"
         Height          =   350
         Left            =   4065
         TabIndex        =   30
         Top             =   4650
         Width           =   1500
      End
      Begin MSComCtl2.UpDown udn 
         Height          =   300
         Left            =   1620
         TabIndex        =   24
         Top             =   4665
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txt"
         BuddyDispid     =   196613
         OrigLeft        =   2205
         OrigTop         =   4635
         OrigRight       =   2460
         OrigBottom      =   4965
         Max             =   60
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txt 
         Height          =   300
         Left            =   1125
         MaxLength       =   2
         TabIndex        =   23
         Top             =   4665
         Width           =   495
      End
      Begin VB.Frame Frame1 
         Caption         =   "ȱʡִ�п���"
         Height          =   2175
         Left            =   120
         TabIndex        =   10
         Top             =   2085
         Width           =   5475
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   9
            Left            =   3750
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   1770
            Width           =   1620
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   8
            Left            =   3750
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1410
            Width           =   1620
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   7
            Left            =   3750
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1020
            Width           =   1620
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   6
            Left            =   3750
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   645
            Width           =   1620
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   5
            Left            =   3750
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   270
            Width           =   1620
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   4
            Left            =   990
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1020
            Width           =   1620
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   3
            Left            =   990
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   645
            Width           =   1620
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   2
            Left            =   990
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   270
            Width           =   1620
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Index           =   10
            Left            =   3255
            TabIndex        =   29
            Top             =   1830
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Index           =   9
            Left            =   3255
            TabIndex        =   27
            Top             =   1470
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "סԺ��ҩ"
            Height          =   180
            Index           =   7
            Left            =   2880
            TabIndex        =   16
            Top             =   1095
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "סԺ��ҩ"
            Height          =   180
            Index           =   6
            Left            =   2880
            TabIndex        =   15
            Top             =   750
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "סԺ��ҩ"
            Height          =   180
            Index           =   5
            Left            =   2880
            TabIndex        =   14
            Top             =   375
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "������ҩ"
            Height          =   180
            Index           =   4
            Left            =   165
            TabIndex        =   13
            Top             =   1080
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�����ҩ"
            Height          =   180
            Index           =   2
            Left            =   180
            TabIndex        =   12
            Top             =   705
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "������ҩ"
            Height          =   180
            Index           =   0
            Left            =   165
            TabIndex        =   11
            Top             =   315
            Width           =   720
         End
      End
      Begin VB.Frame fra 
         Caption         =   "ȱʡʱ��"
         Height          =   1560
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   5475
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   1
            Left            =   2460
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1155
            Width           =   1920
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   0
            Left            =   2460
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   810
            Width           =   1920
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�����������Χ(&6)"
            Height          =   180
            Index           =   3
            Left            =   870
            TabIndex        =   2
            Top             =   1215
            Width           =   1530
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "������ʱ�䷶Χ(&5)"
            Height          =   180
            Index           =   1
            Left            =   870
            TabIndex        =   0
            Top             =   855
            Width           =   1530
         End
         Begin VB.Label lbl 
            Caption         =   "����������վ�еĴ������Լ������������ʱ�䷶Χ�ֱ��������ý���������"
            Height          =   405
            Index           =   11
            Left            =   780
            TabIndex        =   8
            Top             =   360
            Width           =   3840
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   165
            Picture         =   "frmOpsStationPara.frx":0028
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�Զ�ˢ��(&1)         ��"
         Height          =   180
         Index           =   8
         Left            =   105
         TabIndex        =   25
         Top             =   4725
         Width           =   1980
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3480
      TabIndex        =   4
      Top             =   5265
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4650
      TabIndex        =   5
      Top             =   5265
      Width           =   1100
   End
End
Attribute VB_Name = "frmOpsStationPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnOK As Boolean
Private mlngLoop As Long
Private mfrmMain As Object
Private mstrPrivs As String

Public Function ShowPara(ByVal frmMain As Object, ByVal strPrivs As String) As Boolean
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim strPar As String
    
    Dim objCbo As Object
    
    Dim intLoop As Integer
    
    mblnOK = False
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
    '��ʼ��
    '------------------------------------------------------------------------------------------------------------------
    For mlngLoop = 0 To 1
        cbo(mlngLoop).AddItem "��  ��"
        cbo(mlngLoop).AddItem "��  ��"
        cbo(mlngLoop).AddItem "��  ��"
        cbo(mlngLoop).AddItem "��  ��"
        cbo(mlngLoop).AddItem "��  ��"
        cbo(mlngLoop).AddItem "��  ��"
        cbo(mlngLoop).AddItem "��  ��"
        cbo(mlngLoop).AddItem "������"
        cbo(mlngLoop).AddItem "��  ��"
        cbo(mlngLoop).AddItem "ǰ����"
        cbo(mlngLoop).AddItem "ǰһ��"
        cbo(mlngLoop).AddItem "ǰ����"
        cbo(mlngLoop).AddItem "ǰһ��"
        cbo(mlngLoop).AddItem "ǰ����"
        cbo(mlngLoop).AddItem "ǰ����"
        cbo(mlngLoop).AddItem "ǰ����"
        cbo(mlngLoop).AddItem "ǰһ��"
        cbo(mlngLoop).AddItem "ǰ����"
    Next
    
    'ȱʡҩ��
    '------------------------------------------------------------------------------------------------------------------
    cbo(2).AddItem "�ֹ�ѡ��"
    cbo(3).AddItem "�ֹ�ѡ��"
    cbo(4).AddItem "�ֹ�ѡ��"
    cbo(5).AddItem "�ֹ�ѡ��"
    cbo(6).AddItem "�ֹ�ѡ��"
    cbo(7).AddItem "�ֹ�ѡ��"
    cbo(8).AddItem "�ֹ�ѡ��"
    cbo(9).AddItem "�ֹ�ѡ��"
    strSQL = _
        "Select Distinct A.ID,A.����,A.����,B.��������,B.�������" & _
        " From ���ű� A,��������˵�� B " & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And B.����ID=A.ID And B.������� IN(1,2,3)" & _
        " And B.�������� in('��ҩ��','��ҩ��','��ҩ��','���ϲ���')" & _
        " Order by A.����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    For intLoop = 1 To rsTmp.RecordCount
        If rsTmp!�������� = "��ҩ��" Then
            Set objCbo = IIf(rsTmp!������� = 1, cbo(2), IIf(rsTmp!������� = 2, cbo(5), Nothing))
        End If
        If rsTmp!�������� = "��ҩ��" Then
            Set objCbo = IIf(rsTmp!������� = 1, cbo(3), IIf(rsTmp!������� = 2, cbo(6), Nothing))
        End If
        If rsTmp!�������� = "��ҩ��" Then
            Set objCbo = IIf(rsTmp!������� = 1, cbo(4), IIf(rsTmp!������� = 2, cbo(7), Nothing))
        End If
        If objCbo Is Nothing Then
            If rsTmp!�������� = "��ҩ��" Then
                cbo(2).AddItem rsTmp!����
                cbo(2).ItemData(cbo(2).NewIndex) = rsTmp!ID
                cbo(5).AddItem rsTmp!����
                cbo(5).ItemData(cbo(5).NewIndex) = rsTmp!ID
            ElseIf rsTmp!�������� = "��ҩ��" Then
                cbo(3).AddItem rsTmp!����
                cbo(3).ItemData(cbo(3).NewIndex) = rsTmp!ID
                cbo(6).AddItem rsTmp!����
                cbo(6).ItemData(cbo(6).NewIndex) = rsTmp!ID
            ElseIf rsTmp!�������� = "��ҩ��" Then
                cbo(4).AddItem rsTmp!����
                cbo(4).ItemData(cbo(4).NewIndex) = rsTmp!ID
                cbo(7).AddItem rsTmp!����
                cbo(7).ItemData(cbo(7).NewIndex) = rsTmp!ID
            ElseIf rsTmp!�������� = "���ϲ���" Then
                cbo(8).AddItem rsTmp!����
                cbo(8).ItemData(cbo(8).NewIndex) = rsTmp!ID
            End If
        Else
            objCbo.AddItem rsTmp!����
            objCbo.ItemData(objCbo.NewIndex) = rsTmp!ID
        End If
        rsTmp.MoveNext
    Next
    
    strSQL = _
        "Select Distinct A.ID,A.����,A.����,B.��������,B.�������" & _
        " From ���ű� A,��������˵�� B " & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And B.����ID=A.ID And B.������� IN(1,2,3)" & _
        " And B.�������� in('����')" & _
        " Order by A.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTmp.BOF = False Then
        For intLoop = 1 To rsTmp.RecordCount
            cbo(9).AddItem rsTmp!����
            cbo(9).ItemData(cbo(9).NewIndex) = rsTmp!ID
            rsTmp.MoveNext
        Next
    End If
    
    
    strPar = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "����ȱʡ��ҩ��", "0")
    For intLoop = 0 To cbo(2).ListCount - 1
        If cbo(2).ItemData(intLoop) = Val(strPar) Then cbo(2).ListIndex = intLoop: Exit For
    Next
    
    strPar = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "����ȱʡ��ҩ��", "0")
    For intLoop = 0 To cbo(3).ListCount - 1
        If cbo(3).ItemData(intLoop) = Val(strPar) Then cbo(3).ListIndex = intLoop: Exit For
    Next
    
    strPar = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "����ȱʡ��ҩ��", "0")
    For intLoop = 0 To cbo(4).ListCount - 1
        If cbo(4).ItemData(intLoop) = Val(strPar) Then cbo(4).ListIndex = intLoop: Exit For
    Next
    
    strPar = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "סԺȱʡ��ҩ��", "0")
    For intLoop = 0 To cbo(5).ListCount - 1
        If cbo(5).ItemData(intLoop) = Val(strPar) Then cbo(5).ListIndex = intLoop: Exit For
    Next
    
    strPar = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "סԺȱʡ��ҩ��", "0")
    For intLoop = 0 To cbo(6).ListCount - 1
        If cbo(6).ItemData(intLoop) = Val(strPar) Then cbo(6).ListIndex = intLoop: Exit For
    Next
    
    strPar = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "סԺȱʡ��ҩ��", "0")
    For intLoop = 0 To cbo(7).ListCount - 1
        If cbo(7).ItemData(intLoop) = Val(strPar) Then cbo(7).ListIndex = intLoop: Exit For
    Next
        
    strPar = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "����ȱʡ����", "0")
    For intLoop = 0 To cbo(8).ListCount - 1
        If cbo(8).ItemData(intLoop) = Val(strPar) Then cbo(8).ListIndex = intLoop: Exit For
    Next
    
    strPar = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "����ȱʡ����", "0")
    For intLoop = 0 To cbo(9).ListCount - 1
        If cbo(9).ItemData(intLoop) = Val(strPar) Then cbo(9).ListIndex = intLoop: Exit For
    Next
    
    On Error Resume Next
    
    
    cbo(0).Text = GetPara("������ʱ�䷶Χ", mfrmMain.ģ���, , "��  ��")
    cbo(1).Text = GetPara("��������ʱ�䷶Χ", mfrmMain.ģ���, , "��  ��")
    
    txt.Text = GetPara("�Զ�ˢ�¼��", mfrmMain.ģ���, , "0")
    chk.Value = GetPara("����ʱ�������ɷ���", mfrmMain.ģ���, , "0")

    If cbo(0).ListIndex = -1 Then cbo(0).ListIndex = 0
    If cbo(1).ListIndex = -1 Then cbo(1).ListIndex = 0
    
    cbo(0).ForeColor = COLOR.����ģ��ɫ
    cbo(1).ForeColor = COLOR.����ģ��ɫ
    txt.ForeColor = COLOR.����ģ��ɫ
    lbl(11).ForeColor = COLOR.����ģ��ɫ
    lbl(1).ForeColor = COLOR.����ģ��ɫ
    lbl(3).ForeColor = COLOR.����ģ��ɫ
    lbl(8).ForeColor = COLOR.����ģ��ɫ
    fra.ForeColor = COLOR.����ģ��ɫ
    
    fra.Enabled = IsPrivs(mstrPrivs, "��������")
    cbo(0).Enabled = fra.Enabled
    cbo(1).Enabled = fra.Enabled
    udn.Enabled = fra.Enabled
    txt.Enabled = fra.Enabled
    
    Me.Show 1, frmMain
    
    ShowPara = mblnOK
    
End Function

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, ParamInfo.ϵͳ��, ParamInfo.ģ���)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((ParamInfo.ϵͳ��) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim strPar As String, i As Long
    
    Call SetPara("������ʱ�䷶Χ", cbo(0).Text, mfrmMain.ģ���)
    Call SetPara("��������ʱ�䷶Χ", cbo(1).Text, mfrmMain.ģ���)
    
    Call SetPara("�Զ�ˢ�¼��", Val(txt.Text), mfrmMain.ģ���)
    
    Call SetPara("����ʱ�������ɷ���", chk.Value, mfrmMain.ģ���)
        
    If cbo(2).ListIndex = -1 Then
        MsgBox "��ָ��ȱʡ��������ҩ����", vbInformation, gstrSysName
        cbo(2).SetFocus: Exit Sub
    End If
    If cbo(3).ListIndex = -1 Then
        MsgBox "��ָ��ȱʡ�������ҩ����", vbInformation, gstrSysName
        cbo(3).SetFocus: Exit Sub
    End If
    If cbo(4).ListIndex = -1 Then
        MsgBox "��ָ��ȱʡ��������ҩ����", vbInformation, gstrSysName
        cbo(4).SetFocus: Exit Sub
    End If
    If cbo(5).ListIndex = -1 Then
        MsgBox "��ָ��ȱʡ��סԺ��ҩ����", vbInformation, gstrSysName
        cbo(5).SetFocus: Exit Sub
    End If
    If cbo(6).ListIndex = -1 Then
        MsgBox "��ָ��ȱʡ��סԺ��ҩ����", vbInformation, gstrSysName
        cbo(6).SetFocus: Exit Sub
    End If
    If cbo(7).ListIndex = -1 Then
        MsgBox "��ָ��ȱʡ��סԺ��ҩ����", vbInformation, gstrSysName
        cbo(7).SetFocus: Exit Sub
    End If
    
    If cbo(8).ListIndex = -1 Then
        MsgBox "��ָ��ȱʡ�Ĳ���ִ�в��š�", vbInformation, gstrSysName
        cbo(8).SetFocus: Exit Sub
    End If
    If cbo(9).ListIndex = -1 Then
        MsgBox "��ָ��ȱʡ������ִ�в��š�", vbInformation, gstrSysName
        cbo(9).SetFocus: Exit Sub
    End If
    
    'ȱʡҩ��
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "����ȱʡ��ҩ��", cbo(2).ItemData(cbo(2).ListIndex)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "����ȱʡ��ҩ��", cbo(3).ItemData(cbo(3).ListIndex)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "����ȱʡ��ҩ��", cbo(4).ItemData(cbo(4).ListIndex)
    
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "סԺȱʡ��ҩ��", cbo(5).ItemData(cbo(5).ListIndex)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "סԺȱʡ��ҩ��", cbo(6).ItemData(cbo(6).ListIndex)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "סԺȱʡ��ҩ��", cbo(7).ItemData(cbo(7).ListIndex)
    
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "����ȱʡ����", cbo(8).ItemData(cbo(8).ListIndex)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "����ȱʡ����", cbo(9).ItemData(cbo(9).ListIndex)
    
    mblnOK = True

    
    Unload Me
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub tbs_Click(PreviousTab As Integer)
    tbs.ZOrder 0
End Sub

Private Sub txt_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txt.Text, txt.MaxLength)
    If Cancel Then Exit Sub

    If Val(txt.Text) < udn.Min Or Val(txt.Text) > udn.Max Then
        Cancel = True
    End If
End Sub
