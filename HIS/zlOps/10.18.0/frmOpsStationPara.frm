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
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txt"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "udn"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin MSComCtl2.UpDown udn 
         Height          =   300
         Left            =   1951
         TabIndex        =   24
         Top             =   4665
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txt"
         BuddyDispid     =   196621
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
         Left            =   1455
         MaxLength       =   2
         TabIndex        =   23
         Top             =   4665
         Width           =   495
      End
      Begin VB.Frame Frame1 
         Caption         =   "ҩ������"
         Height          =   2505
         Left            =   120
         TabIndex        =   10
         Top             =   2085
         Width           =   5475
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   7
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   2085
            Width           =   1920
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   6
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1725
            Width           =   1920
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   5
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1365
            Width           =   1920
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   4
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1005
            Width           =   1920
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   3
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   630
            Width           =   1920
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   2
            Left            =   1950
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   255
            Width           =   1920
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   270
            Picture         =   "frmOpsStationPara.frx":0028
            Top             =   225
            Width           =   480
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "סԺ��ҩ��"
            Height          =   180
            Index           =   7
            Left            =   990
            TabIndex        =   16
            Top             =   2130
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "סԺ��ҩ��"
            Height          =   180
            Index           =   6
            Left            =   990
            TabIndex        =   15
            Top             =   1785
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "סԺ��ҩ��"
            Height          =   180
            Index           =   5
            Left            =   990
            TabIndex        =   14
            Top             =   1410
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "������ҩ��"
            Height          =   180
            Index           =   4
            Left            =   990
            TabIndex        =   13
            Top             =   1065
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�����ҩ��"
            Height          =   180
            Index           =   2
            Left            =   990
            TabIndex        =   12
            Top             =   690
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "������ҩ��"
            Height          =   180
            Index           =   0
            Left            =   990
            TabIndex        =   11
            Top             =   300
            Width           =   900
         End
      End
      Begin VB.Frame Frame5 
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
         Begin VB.Label Label9 
            Caption         =   "����������վ�еĴ������Լ������������ʱ�䷶Χ�ֱ��������ý���������"
            Height          =   405
            Left            =   780
            TabIndex        =   8
            Top             =   360
            Width           =   3840
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   165
            Picture         =   "frmOpsStationPara.frx":08F2
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�Զ�ˢ��(&1)         ��"
         Height          =   180
         Index           =   8
         Left            =   435
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

Public Function ShowPara(ByVal frmMain As Object) As Boolean
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim strPar As String
    
    Dim objCbo As Object
    
    Dim intLoop As Integer
    
    mblnOK = False
    
    Set mfrmMain = frmMain
    '��ʼ��
    '------------------------------------------------------------------------------------------------------------------
    For mlngLoop = 0 To 1
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
    strSQL = _
        "Select Distinct A.ID,A.����,A.����,B.��������,B.�������" & _
        " From ���ű� A,��������˵�� B " & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And B.����ID=A.ID And B.������� IN(1,2,3)" & _
        " And B.�������� in('��ҩ��','��ҩ��','��ҩ��')" & _
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
            End If
        Else
            objCbo.AddItem rsTmp!����
            objCbo.ItemData(objCbo.NewIndex) = rsTmp!ID
        End If
        rsTmp.MoveNext
    Next
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
        
        
    On Error Resume Next
    
    
    cbo(0).Text = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.���ݿ��û� & "\" & App.ProductName & "\" & mfrmMain.Name, "������ʱ�䷶Χ", "��  ��")
    cbo(1).Text = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.���ݿ��û� & "\" & App.ProductName & "\" & mfrmMain.Name, "��������ʱ�䷶Χ", "��  ��")
    
    txt.Text = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.���ݿ��û� & "\" & App.ProductName & "\" & mfrmMain.Name, "�Զ�ˢ�¼��", "0")


    If cbo(0).ListIndex = -1 Then cbo(0).ListIndex = 0
    If cbo(1).ListIndex = -1 Then cbo(1).ListIndex = 0
    
    Me.Show 1, frmMain
    
    ShowPara = mblnOK
    
End Function

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((ParamInfo.ϵͳ��) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim strPar As String, i As Long
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.���ݿ��û� & "\" & App.ProductName & "\" & mfrmMain.Name, "������ʱ�䷶Χ", cbo(0).Text)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.���ݿ��û� & "\" & App.ProductName & "\" & mfrmMain.Name, "��������ʱ�䷶Χ", cbo(1).Text)
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.���ݿ��û� & "\" & App.ProductName & "\" & mfrmMain.Name, "�Զ�ˢ�¼��", Val(txt.Text))
        
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
    
    
    'ȱʡҩ��
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "����ȱʡ��ҩ��", cbo(2).ItemData(cbo(2).ListIndex)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "����ȱʡ��ҩ��", cbo(3).ItemData(cbo(3).ListIndex)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "����ȱʡ��ҩ��", cbo(4).ItemData(cbo(4).ListIndex)
    
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "סԺȱʡ��ҩ��", cbo(5).ItemData(cbo(5).ListIndex)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "סԺȱʡ��ҩ��", cbo(6).ItemData(cbo(6).ListIndex)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "סԺȱʡ��ҩ��", cbo(7).ItemData(cbo(7).ListIndex)
    
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