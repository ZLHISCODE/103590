VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMipComOption 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ѡ������"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6360
   Icon            =   "frmMipComOption.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   6360
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame2 
      Caption         =   "���Ѵ���͸����"
      Height          =   1230
      Left            =   60
      TabIndex        =   19
      Top             =   2895
      Width           =   6210
      Begin MSComctlLib.Slider sld 
         Height          =   465
         Left            =   1275
         TabIndex        =   22
         Top             =   645
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   820
         _Version        =   393216
         LargeChange     =   1
         Max             =   20
         SelStart        =   5
         Value           =   5
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "25%"
         Height          =   180
         Index           =   6
         Left            =   5715
         TabIndex        =   21
         Top             =   720
         Width           =   270
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "���������Ѵ�����������ƶ�ʱ��͸���̶�"
         Height          =   180
         Index           =   5
         Left            =   1395
         TabIndex        =   20
         Top             =   345
         Width           =   3420
      End
      Begin VB.Image img 
         Height          =   480
         Index           =   3
         Left            =   180
         Picture         =   "frmMipComOption.frx":6852
         Top             =   315
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "������־"
      Height          =   1275
      Left            =   60
      TabIndex        =   13
      Top             =   4200
      Width           =   6210
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   3285
         MaxLength       =   2
         TabIndex        =   18
         Text            =   "7"
         Top             =   810
         Width           =   510
      End
      Begin VB.CheckBox chk 
         Caption         =   "������־��¼"
         Height          =   285
         Left            =   1395
         TabIndex        =   14
         Top             =   810
         Width           =   1395
      End
      Begin MSComCtl2.UpDown upd 
         Height          =   300
         Index           =   1
         Left            =   3825
         TabIndex        =   17
         Top             =   795
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txt(1)"
         BuddyDispid     =   196614
         BuddyIndex      =   1
         OrigLeft        =   3270
         OrigTop         =   1005
         OrigRight       =   3525
         OrigBottom      =   1320
         Max             =   30
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����          ��"
         Height          =   180
         Index           =   4
         Left            =   2850
         TabIndex        =   16
         Top             =   855
         Width           =   1440
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�����Ƿ��¼������־�Լ���־���ݵı���ʱ��"
         Height          =   180
         Index           =   3
         Left            =   1365
         TabIndex        =   15
         Top             =   345
         Width           =   3780
      End
      Begin VB.Image img 
         Height          =   480
         Index           =   2
         Left            =   180
         Picture         =   "frmMipComOption.frx":81D4
         Top             =   345
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   345
      Left            =   3930
      TabIndex        =   11
      Top             =   5625
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   345
      Left            =   5145
      TabIndex        =   10
      Top             =   5625
      Width           =   1100
   End
   Begin VB.Frame fra 
      Caption         =   "ͣ��ʱ��"
      Height          =   1440
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   1380
      Width           =   6210
      Begin MSComCtl2.UpDown upd 
         Height          =   300
         Index           =   0
         Left            =   2985
         TabIndex        =   9
         Top             =   1020
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txt(0)"
         BuddyDispid     =   196614
         BuddyIndex      =   0
         OrigLeft        =   3270
         OrigTop         =   1005
         OrigRight       =   3525
         OrigBottom      =   1320
         Max             =   30
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   300
         Index           =   0
         Left            =   2445
         MaxLength       =   2
         TabIndex        =   8
         Top             =   1020
         Width           =   795
      End
      Begin VB.OptionButton opt 
         Caption         =   "�̶�ʱ��"
         Height          =   225
         Index           =   1
         Left            =   1365
         TabIndex        =   7
         Top             =   1065
         Value           =   -1  'True
         Width           =   1140
      End
      Begin VB.OptionButton opt 
         Caption         =   "һֱͣ��"
         Height          =   240
         Index           =   0
         Left            =   1365
         TabIndex        =   6
         Top             =   675
         Width           =   1155
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   0
         Left            =   3345
         TabIndex        =   12
         Top             =   1095
         Width           =   180
      End
      Begin VB.Image img 
         Height          =   480
         Index           =   1
         Left            =   210
         Picture         =   "frmMipComOption.frx":9B56
         Top             =   330
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "������������Ϣ�����˲���ʱͣ����ʱ��"
         Height          =   180
         Index           =   1
         Left            =   1350
         TabIndex        =   5
         Top             =   345
         Width           =   3240
      End
   End
   Begin VB.Frame fra 
      Caption         =   "��������"
      Height          =   1230
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   6210
      Begin VB.CommandButton cmdHear 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   675
         Width           =   1100
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         ItemData        =   "frmMipComOption.frx":B4D8
         Left            =   1455
         List            =   "frmMipComOption.frx":B4DA
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   705
         Width           =   3240
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����������Ϣʱ��������������"
         Height          =   180
         Index           =   2
         Left            =   1425
         TabIndex        =   4
         Top             =   345
         Width           =   2520
      End
      Begin VB.Image img 
         Height          =   480
         Index           =   0
         Left            =   180
         Picture         =   "frmMipComOption.frx":B4DC
         Top             =   375
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmMipComOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'��������
Private mblnDataChanged As Boolean
Private mstrTitle As String
Private mclsMipSystemData As clsMipSystemData

Public Event OptionChanged()

'######################################################################################################################
'�ӿڷ���

Public Function ShowDialog(ByVal frmParent As Object, ByVal strDataFile As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rsCondition As ADODB.Recordset
    Dim rs As zlDataSQLite.SQLiteRecordset
    Dim strPara As String
    Dim varPara As Variant
    Dim lngLoop As Long
        
    Call cbo(0).AddItem("��")
    For lngLoop = 101 To 111
        Call cbo(0).AddItem(GetWaveName(lngLoop))
        cbo(0).ItemData(cbo(0).NewIndex) = lngLoop
    Next
    cbo(0).ListIndex = 0
    
    Set mclsMipSystemData = New clsMipSystemData
    Call mclsMipSystemData.Initialize(strDataFile)
    
    txt(0).Text = "5"
        
    strPara = ""
    If mclsMipSystemData.OpenDataFile() = True Then
        
        '��Ϣ��������
        Set rsCondition = CreateCondition
        Call SetCondition(rsCondition, "�������", "1")
        rs = mclsMipSystemData.GetPara("Filter", rsCondition)
        If rs.DataSet.BOF = False Then
            strPara = NVL(rs.DataSet("Para_Value").Value)
            Call CboLocate(cbo(0), Val(strPara), True)
        End If

        '��Ϣͣ��ʱ��
        Set rsCondition = CreateCondition
        Call SetCondition(rsCondition, "�������", "2")
        rs = mclsMipSystemData.GetPara("Filter", rsCondition)
        If rs.DataSet.BOF = False Then
            strPara = NVL(rs.DataSet("Para_Value").Value)
            If Val(strPara) = 0 Then
                opt(0).Value = True
            Else
                opt(1).Value = True
                txt(0).Text = Val(strPara)
            End If
        End If
                
        '�Ƿ��¼��־
        Set rsCondition = CreateCondition
        Call SetCondition(rsCondition, "�������", "3")
        rs = mclsMipSystemData.GetPara("Filter", rsCondition)
        If rs.DataSet.BOF = False Then
            strPara = NVL(rs.DataSet("Para_Value").Value)
            chk.Value = Val(strPara)
        End If
        
                
        '��־����ʱ��
        Set rsCondition = CreateCondition
        Call SetCondition(rsCondition, "�������", "4")
        rs = mclsMipSystemData.GetPara("Filter", rsCondition)
        If rs.DataSet.BOF = False Then
            strPara = NVL(rs.DataSet("Para_Value").Value)
            txt(1).Text = Val(strPara)
        End If
        
        '͸����
        Set rsCondition = CreateCondition
        Call SetCondition(rsCondition, "�������", "5")
        rs = mclsMipSystemData.GetPara("Filter", rsCondition)
        If rs.DataSet.BOF = False Then
            strPara = NVL(rs.DataSet("Para_Value").Value)
            sld.Value = Val(strPara)
        End If
        
        txt(1).Enabled = (chk.Value = 1)
        upd(1).Enabled = (chk.Value = 1)
    End If
    
    mclsMipSystemData.CloseDataFile
        
    mblnDataChanged = False
    
    Me.Show , frmParent
        
    ShowDialog = mblnDataChanged
    
End Function

Private Function GetWaveName(ByVal lngNo As Long) As String
    
    Select Case lngNo
    Case 101
        GetWaveName = "����"
    Case 102
        GetWaveName = "����ռ�"
    Case 103
        GetWaveName = "�绰����1"
    Case 104
        GetWaveName = "�绰����2"
    Case 105
        GetWaveName = "�绰��"
    Case 106
        GetWaveName = "������"
    Case 107
        GetWaveName = "����"
    Case 108
        GetWaveName = "����"
    Case 109
        GetWaveName = "��ʾ"
    Case 110
        GetWaveName = "����Ϣ"
    Case 111
        GetWaveName = "����Ϣ(Ů��)"
    End Select
        
End Function


Private Function GetWaveCode(ByVal lngName As String) As Long
    
    Select Case lngName
    Case "����"
        GetWaveCode = 101
    Case "����ռ�"
        GetWaveCode = 102
    Case "�绰����1"
        GetWaveCode = 103
    Case "�绰����2"
        GetWaveCode = 104
    Case "�绰��"
        GetWaveCode = 105
    Case "������"
        GetWaveCode = 106
    Case "����"
        GetWaveCode = 107
    Case "����"
        GetWaveCode = 108
    Case "��ʾ"
        GetWaveCode = 109
    Case "����Ϣ"
        GetWaveCode = 110
    Case "����Ϣ(Ů��)"
        GetWaveCode = 111
    End Select
    
End Function

Private Sub chk_Click()
    txt(1).Enabled = (chk.Value = 1)
    upd(1).Enabled = (chk.Value = 1)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHear_Click()
    If cbo(0).Text = "" Then Exit Sub
    
    Call PlayWave(GetWaveCode(cbo(0).Text))
    
    cbo(0).SetFocus
End Sub

Private Sub cmdOK_Click()
    
    Dim blnRet As Boolean
        
    If mclsMipSystemData.OpenDataFile() = True Then
        
        blnRet = mclsMipSystemData.EditPara("1", GetWaveCode(cbo(0).Text))
        If blnRet Then
            If opt(0).Value = True Then
                blnRet = mclsMipSystemData.EditPara("2", 0)
            Else
                blnRet = mclsMipSystemData.EditPara("2", Val(txt(0).Text))
            End If
        End If
        If blnRet Then blnRet = mclsMipSystemData.EditPara("3", chk.Value)
        If blnRet Then blnRet = mclsMipSystemData.EditPara("4", Val(txt(1).Text))
        If blnRet Then blnRet = mclsMipSystemData.EditPara("5", sld.Value)
                        
        mclsMipSystemData.CloseDataFile
        
        If blnRet = True Then
            RaiseEvent OptionChanged
            mblnDataChanged = True
            Unload Me
            Exit Sub
        End If
    End If
    mclsMipSystemData.CloseDataFile
    
End Sub



Private Sub opt_Click(Index As Integer)
        
    txt(0).Visible = opt(1).Value
    upd(0).Visible = opt(1).Value
    lbl(0).Visible = opt(1).Value
            
End Sub

Private Sub sld_Change()
    lbl(6).Caption = sld.Value * 5 & "%"
End Sub

Private Sub sld_Click()
    lbl(6).Caption = sld.Value * 5 & "%"
End Sub
