VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNoticesEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ѱ༭"
   ClientHeight    =   6180
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9345
   Icon            =   "frmNoticesEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ImageList ils16 
      Left            =   1440
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNoticesEdit.frx":000C
            Key             =   "Station"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNoticesEdit.frx":238E
            Key             =   "Human"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNoticesEdit.frx":73F8
            Key             =   "Dept"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   105
      TabIndex        =   41
      Top             =   5730
      Width           =   1100
   End
   Begin VB.CommandButton cmdClear 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8160
      TabIndex        =   40
      Top             =   5730
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6960
      TabIndex        =   39
      Top             =   5730
      Width           =   1100
   End
   Begin TabDlg.SSTab tbs 
      Height          =   5580
      Left            =   15
      TabIndex        =   42
      Top             =   45
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   9843
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   635
      TabCaption(0)   =   "&1.����   "
      TabPicture(0)   =   "frmNoticesEdit.frx":BA72
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Image1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Image2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fra(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdOpen"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdHear"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdValid"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chk(3)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chk(2)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txt(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cbo(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txt(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdPlan"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "fra(1)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txt(1)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "&2.ϸ��   "
      TabPicture(1)   =   "frmNoticesEdit.frx":BA8E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chk(1)"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(3)=   "cmdAdd(0)"
      Tab(1).Control(4)=   "cmdRemove(0)"
      Tab(1).Control(5)=   "Frame2"
      Tab(1).Control(6)=   "Frame1"
      Tab(1).Control(7)=   "cmdAdd(1)"
      Tab(1).Control(8)=   "cmdRemove(1)"
      Tab(1).Control(9)=   "cmdAdd(2)"
      Tab(1).Control(10)=   "cmdRemove(2)"
      Tab(1).Control(11)=   "chk(4)"
      Tab(1).Control(12)=   "lvwHuman"
      Tab(1).Control(13)=   "lvwStation"
      Tab(1).Control(14)=   "lvwDept"
      Tab(1).Control(15)=   "lbl(5)"
      Tab(1).Control(16)=   "lbl(2)"
      Tab(1).Control(17)=   "lbl(12)"
      Tab(1).ControlCount=   18
      Begin VB.TextBox txt 
         Height          =   2235
         Index           =   1
         Left            =   1200
         MaxLength       =   4000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   51
         Top             =   2760
         Width           =   7920
      End
      Begin VB.Frame fra 
         Height          =   30
         Index           =   1
         Left            =   1410
         TabIndex        =   50
         Top             =   2580
         Width           =   7650
      End
      Begin VB.CommandButton cmdPlan 
         Caption         =   "ִ�мƻ�(&P)"
         Height          =   350
         Left            =   6720
         TabIndex        =   49
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CheckBox chk 
         Caption         =   "����ʱ��������(&T)"
         Height          =   225
         Index           =   1
         Left            =   -69780
         TabIndex        =   38
         Top             =   4875
         Value           =   1  'Checked
         Width           =   1980
      End
      Begin VB.Frame Frame4 
         Caption         =   "�������"
         Height          =   3900
         Left            =   -69780
         TabIndex        =   47
         Top             =   900
         Width           =   3615
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1560
            ScaleHeight     =   240
            ScaleWidth      =   1500
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   675
            Width           =   1500
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   3
            Left            =   1545
            MaxLength       =   5
            TabIndex        =   31
            Text            =   "1"
            Top             =   1455
            Width           =   615
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   1
            Left            =   2460
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   1440
            Width           =   900
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   4
            Left            =   1545
            MaxLength       =   5
            TabIndex        =   35
            Text            =   "2"
            Top             =   1845
            Width           =   615
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   2
            Left            =   1545
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   1050
            Width           =   1815
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   3
            Left            =   2445
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   1830
            Width           =   915
         End
         Begin VB.CheckBox chk 
            Caption         =   "��(&T)"
            Height          =   225
            Index           =   0
            Left            =   705
            TabIndex        =   25
            Top             =   690
            Width           =   780
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   1
            Left            =   1530
            TabIndex        =   27
            Top             =   645
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   104726531
            CurrentDate     =   38173
         End
         Begin MSComCtl2.UpDown udn 
            Height          =   285
            Index           =   0
            Left            =   2175
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   1455
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txt(3)"
            BuddyDispid     =   196614
            BuddyIndex      =   3
            OrigLeft        =   6330
            OrigTop         =   1830
            OrigRight       =   6570
            OrigBottom      =   2100
            Max             =   100
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udn 
            Height          =   285
            Index           =   1
            Left            =   2175
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   1845
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txt(4)"
            BuddyDispid     =   196614
            BuddyIndex      =   4
            OrigLeft        =   6330
            OrigTop         =   2205
            OrigRight       =   6570
            OrigBottom      =   2475
            Max             =   100
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   0
            Left            =   1545
            TabIndex        =   24
            Top             =   285
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   104726531
            CurrentDate     =   38173
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "��������(&B)"
            Height          =   180
            Index           =   3
            Left            =   150
            TabIndex        =   34
            Top             =   1905
            Width           =   990
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�������(&G)"
            Height          =   180
            Index           =   8
            Left            =   150
            TabIndex        =   30
            Top             =   1515
            Width           =   990
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "��ʼʱ�� ��(&F)"
            Height          =   180
            Index           =   4
            Left            =   150
            TabIndex        =   23
            Top             =   345
            Width           =   1260
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "��鷽ʽ(&F)"
            Height          =   180
            Index           =   9
            Left            =   150
            TabIndex        =   28
            Top             =   1095
            Width           =   990
         End
      End
      Begin VB.Frame Frame3 
         Height          =   30
         Left            =   -74130
         TabIndex        =   46
         Top             =   2355
         Width           =   3510
      End
      Begin VB.CommandButton cmdAdd 
         Height          =   350
         Index           =   0
         Left            =   -70950
         Picture         =   "frmNoticesEdit.frx":BAAA
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1140
         Width           =   390
      End
      Begin VB.CommandButton cmdRemove 
         Height          =   350
         Index           =   0
         Left            =   -70950
         Picture         =   "frmNoticesEdit.frx":10114
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1545
         Width           =   390
      End
      Begin VB.Frame Frame2 
         Height          =   75
         Left            =   -74025
         TabIndex        =   45
         Top             =   3810
         Width           =   3510
      End
      Begin VB.Frame Frame1 
         Height          =   30
         Left            =   -74130
         TabIndex        =   44
         Top             =   975
         Width           =   3510
      End
      Begin VB.CommandButton cmdAdd 
         Height          =   350
         Index           =   1
         Left            =   -70950
         Picture         =   "frmNoticesEdit.frx":1049E
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2505
         Width           =   390
      End
      Begin VB.CommandButton cmdRemove 
         Height          =   350
         Index           =   1
         Left            =   -70950
         Picture         =   "frmNoticesEdit.frx":154F8
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2910
         Width           =   390
      End
      Begin VB.CommandButton cmdAdd 
         Height          =   350
         Index           =   2
         Left            =   -70950
         Picture         =   "frmNoticesEdit.frx":15882
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3990
         Width           =   390
      End
      Begin VB.CommandButton cmdRemove 
         Height          =   350
         Index           =   2
         Left            =   -70950
         Picture         =   "frmNoticesEdit.frx":17BF4
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   4395
         Width           =   390
      End
      Begin VB.CheckBox chk 
         Caption         =   "������Ա(&A)"
         Height          =   225
         Index           =   4
         Left            =   -74820
         TabIndex        =   10
         Top             =   600
         Value           =   1  'Checked
         Width           =   2730
      End
      Begin VB.TextBox txt 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2550
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2100
         Width           =   6135
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         ItemData        =   "frmNoticesEdit.frx":17F7E
         Left            =   2550
         List            =   "frmNoticesEdit.frx":17F80
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1725
         Width           =   6135
      End
      Begin VB.TextBox txt 
         Height          =   720
         Index           =   0
         Left            =   1215
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   930
         Width           =   7800
      End
      Begin VB.CheckBox chk 
         Caption         =   "��������(&S)"
         Height          =   225
         Index           =   2
         Left            =   1215
         TabIndex        =   2
         Top             =   1740
         Width           =   1425
      End
      Begin VB.CheckBox chk 
         Caption         =   "���ѱ���(&R)"
         Height          =   225
         Index           =   3
         Left            =   1215
         TabIndex        =   5
         Top             =   2145
         Width           =   1545
      End
      Begin VB.CommandButton cmdValid 
         Caption         =   "У��SQL(&V)"
         Height          =   350
         Left            =   8040
         TabIndex        =   9
         Top             =   5040
         Width           =   1100
      End
      Begin VB.CommandButton cmdHear 
         Height          =   285
         Left            =   8760
         Picture         =   "frmNoticesEdit.frx":17F82
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1733
         Width           =   300
      End
      Begin VB.CommandButton cmdOpen 
         Height          =   285
         Left            =   8760
         Picture         =   "frmNoticesEdit.frx":1C964
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2160
         Width           =   300
      End
      Begin VB.Frame fra 
         Height          =   30
         Index           =   2
         Left            =   1200
         TabIndex        =   43
         Top             =   765
         Width           =   7740
      End
      Begin MSComctlLib.ListView lvwHuman 
         Height          =   1155
         Left            =   -74835
         TabIndex        =   16
         Top             =   2520
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         Icons           =   "ils16"
         SmallIcons      =   "ils16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "��Ա����"
            Object.Width           =   4233
         EndProperty
      End
      Begin MSComctlLib.ListView lvwStation 
         Height          =   1095
         Left            =   -74835
         TabIndex        =   20
         Top             =   3990
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1931
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         Icons           =   "ils16"
         SmallIcons      =   "ils16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����վ����"
            Object.Width           =   4233
         EndProperty
      End
      Begin MSComctlLib.ListView lvwDept 
         Height          =   1125
         Left            =   -74835
         TabIndex        =   12
         Top             =   1125
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1984
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         Icons           =   "ils16"
         SmallIcons      =   "ils16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "��������"
            Object.Width           =   4233
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ע:�����������п��԰���[����]������������������ѯ����ֵ��"
         Height          =   180
         Left            =   1185
         TabIndex        =   48
         Top             =   5130
         Width           =   5130
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��Ա(&M)"
         Height          =   180
         Index           =   5
         Left            =   -74850
         TabIndex        =   15
         Top             =   2295
         Width           =   630
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����(&D)"
         Height          =   180
         Index           =   2
         Left            =   -74835
         TabIndex        =   11
         Top             =   930
         Width           =   630
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����վ(&S)"
         Height          =   180
         Index           =   12
         Left            =   -74835
         TabIndex        =   19
         Top             =   3765
         Width           =   810
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   435
         Picture         =   "frmNoticesEdit.frx":1D7A6
         Top             =   2835
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   435
         Picture         =   "frmNoticesEdit.frx":21898
         Top             =   1050
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "��������(&L)"
         Height          =   180
         Index           =   1
         Left            =   330
         TabIndex        =   8
         Top             =   2520
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Index           =   0
         Left            =   330
         TabIndex        =   0
         Top             =   675
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmNoticesEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mlngKey As Long
Private mblnOk As Boolean
Private mlngSys As Long
Private mstr������ As String
Private mlngLoop As Long


Private Function CheckFullTable(ByVal strSQL As String) As String
    '--------------------------------------------------------------------------------------------
    '����:ȫ��ɨ����
    '����:������SQL���
    '����:
    '--------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strFull As String
    Dim strKey As String
    
    strKey = "�Զ����Ѽ��ȫ��ɨ��"
    
    On Error Resume Next
    gcnOracle.Execute "delete from PLAN_TABLE where STATEMENT_ID='" & strKey & "'"
    
    On Error GoTo errHand
    gcnOracle.Execute "explain plan set StaTement_ID='" & strKey & "' for " & strSQL
    
    If rs.State = adStateOpen Then rs.Close
    rs.Open "SELECT OBJECT_NAME FROM PLAN_TABLE WHERE upper(OBJECT_NAME)<>'DUAL' and STATEMENT_ID='" & strKey & "' AND upper(OPTIONS)='FULL'", gcnOracle
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            If InStr(strFull & ",", "," & rs("OBJECT_NAME").value & ",") = 0 Then
                strFull = strFull & "," & rs("OBJECT_NAME").value
            End If
            
            rs.MoveNext
        Loop
        If strFull <> "" Then strFull = Mid(strFull, 2)
    End If
    
    CheckFullTable = strFull
    
    Exit Function
errHand:
    
End Function

Public Function PlayWave(lngKey As Long) As String
    '����:����Դ�ļ��е�ָ����Դ���ɴ����ļ�
    '����:ID=��Դ��,strExt=Ҫ�����ļ�����չ��(��BMP)
    '����:�����ļ���
    
    Dim arrData() As Byte
    Dim intFile As Integer
    Dim strFile As String * 255
    Dim strR As String
    
    On Error Resume Next
    
    arrData = LoadResData(lngKey, "WAVE")
    intFile = FreeFile
    
    GetTempPath 255, strFile
    strR = Trim(Left(strFile, InStr(strFile, Chr(0)) - 1)) & CLng(Timer * 100) & ".pic"

    Open strR For Binary As intFile
    Put intFile, , arrData()
    Close intFile

    Call sndPlaySound(strR, SND_NODEFAULT Or SND_ASYNC)
    
    Kill strR
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
    
    End Select
    
End Function

Private Function CalcTimeUnit(ByVal lngData As Long, Optional ByVal strParam As String = "") As String
    
    Dim strNumber As String
    Dim strUnit As String
    
    If lngData / (24 * 60) >= 1 Then
        strNumber = lngData / (24 * 60)
        strUnit = "��"
    ElseIf (lngData / 60) >= 1 Then
        strNumber = (lngData / 60)
        strUnit = "Сʱ"
    Else
        strNumber = lngData
        strUnit = "����"
    End If
    
    Select Case strParam
    Case "����"
        CalcTimeUnit = strNumber
    Case "ʱ�䵥λ"
        CalcTimeUnit = strUnit
    Case ""
        CalcTimeUnit = strNumber & strUnit
    End Select
    
End Function

Private Function Nvl(ByVal varOld As Variant, Optional ByVal varNew As Variant = "") As Variant
    If IsNull(varOld) Then
        Nvl = varNew
    Else
        Nvl = varOld
    End If
    
End Function

Private Function CalcTimeToSecend(ByVal lngData As String, ByVal strUnit As String) As Long
    
    Select Case strUnit
    Case "����"
        CalcTimeToSecend = lngData
    Case "Сʱ"
        CalcTimeToSecend = lngData * 60
    Case "��"
        CalcTimeToSecend = lngData * 60 * 24
    End Select
    
    
End Function

Private Function ReadData() As Boolean
    '----------------------------------------------------------------------
    '���ܣ�
    '----------------------------------------------------------------------
    Dim objItem As ListItem
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    Set rs = OpenCursor(gcnOracle, "ZLTOOLS.B_Expert.Get_notices", mlngKey, 0)
    If rs.BOF = False Then
    
        txt(0).Text = Nvl(rs("��������").value)
        txt(0).Tag = txt(0).Text
        txt(1).Text = Nvl(rs("��������").value)
        txt(2).Tag = Nvl(rs("���ѱ���").value)
        txt(2).Text = Nvl(rs("��������").value)
        
        
        On Error Resume Next
        cbo(0).Text = GetWaveName(Nvl(rs("��������").value, 0))
        chk(2).value = IIf(cbo(0).ListIndex > 0, 1, 0)
        On Error GoTo errHand
        
        chk(1).value = Nvl(rs("���Ѵ���").value, 0)
        chk(3).value = IIf(txt(2).Text <> "", 1, 0)
                
        dtp(0).value = Format(rs("��ʼʱ��").value, dtp(0).CustomFormat)
        
        If IsNull(rs("��ֹʱ��").value) = False Then
            chk(0).value = 1
            dtp(1).value = Format(rs("��ֹʱ��").value, dtp(1).CustomFormat)
        End If
        
        cbo(2).ListIndex = IIf(Nvl(rs("�������").value, 0) = 0, 0, 1)
        If cbo(2).ListIndex = 1 Then
        
            On Error Resume Next
            
            txt(3).Text = CalcTimeUnit(Nvl(rs("�������").value, udn(0).Min), "����")
            cbo(1).Text = CalcTimeUnit(Nvl(rs("�������").value, udn(0).Min), "ʱ�䵥λ")
            
            txt(4).Text = CalcTimeUnit(Nvl(rs("��������").value, udn(0).Min), "����")
            cbo(3).Text = CalcTimeUnit(Nvl(rs("��������").value, udn(0).Min), "ʱ�䵥λ")
            
            On Error GoTo errHand
            
        End If
        
    End If
    
    '��ȡ���Ѷ�����������û�����ݣ��������������Ա
    Set rs = OpenCursor(gcnOracle, "ZLTOOLS.B_Expert.Get_noticeusr", 0, mlngKey)
    If rs.BOF = False Then
        chk(4).value = 0
        
        lvwDept.ListItems.Clear
        Set rs = OpenCursor(gcnOracle, "ZLTOOLS.B_Expert.Get_noticeusr", 2, mlngKey)
        If rs.BOF = False Then
            Do While Not rs.EOF
                Set objItem = lvwDept.ListItems.Add(, "K" & rs("��������").value, rs("��������").value, "Dept", "Dept")
                rs.MoveNext
            Loop
        End If
        
        lvwHuman.ListItems.Clear

        Set rs = OpenCursor(gcnOracle, "ZLTOOLS.B_Expert.Get_noticeusr", 1, mlngKey)
        If rs.BOF = False Then
            Do While Not rs.EOF
                Set objItem = lvwHuman.ListItems.Add(, "K" & rs("��������").value, rs("��������").value, "Human", "Human")
                rs.MoveNext
            Loop
        End If
                    
        lvwStation.ListItems.Clear
        Set rs = OpenCursor(gcnOracle, "ZLTOOLS.B_Expert.Get_noticeusr", 3, mlngKey)
        If rs.BOF = False Then
            Do While Not rs.EOF
                Set objItem = lvwStation.ListItems.Add(, "K" & rs("��������").value, rs("��������").value, "Station", "Station")
                rs.MoveNext
            Loop
        End If
        
    Else
        chk(4).value = 1
    End If
    ReadData = True
    
    Exit Function
    
errHand:
    
End Function

Public Function ReDimArray(ByRef strArray() As String) As Long
    '----------------------------------------------------------------------
    '���ܣ����¶�������
    '----------------------------------------------------------------------
    Dim lngCount As Long
    Dim strTmp As String
    
    On Error GoTo InitHand
    
    strTmp = strArray(1)
    
    lngCount = UBound(strArray) + 1
    
    GoTo OkHand
    
InitHand:
    
    lngCount = 1
    
OkHand:
    
    ReDim Preserve strArray(1 To lngCount)
            
    ReDimArray = lngCount
End Function

Public Function ReplaceAll(vTar As String, vFind As String, vRep As String) As String
    Dim intPos As Long
    Dim strTmp As String
    Dim intStart As Long
    
    ReplaceAll = vTar
    
    intStart = 1
    intPos = InStr(intStart, vTar, vFind)
    
    While intPos > 0
        
        strTmp = strTmp & Mid(vTar, intStart, intPos - intStart) & vRep
        
        intStart = intPos + Len(vFind)
        intPos = InStr(intStart, vTar, vFind)
    Wend
    
    strTmp = strTmp & Mid(vTar, intStart)
    
    ReplaceAll = strTmp
    
End Function


Private Function ValidSQL(ByVal strSQL As String, Optional ByRef strErr As String) As Boolean
    
    Dim rs As New ADODB.Recordset
    
    ValidSQL = True
    
    If Trim(strSQL) = "" Then Exit Function
    
    strSQL = ReplaceAll(UCase(strSQL), "[USER]", "'USER'")
    strSQL = "Select * From (" & strSQL & ") Where 1=2"
    
    On Error GoTo errHand
    
    rs.Open strSQL, gcnOracle
    
    Exit Function
    
errHand:
    ValidSQL = False
    strErr = err.Description
End Function

Private Function InitData() As Boolean
    
    chk(0).value = 0
    chk(1).value = 0
    chk(2).value = 0
    chk(3).value = 0
    chk(4).value = 1
    
    dtp(1).Enabled = False
        
    cbo(0).Enabled = False
    cmdHear.Enabled = False
    
    txt(2).Enabled = False
    cmdOpen.Enabled = False
    
    
    cbo(2).addItem "�������"
    cbo(2).addItem "���ڼ��"
    cbo(2).ListIndex = 0
    
    cbo(1).addItem "����"
    cbo(1).addItem "Сʱ"
    cbo(1).addItem "��"
    cbo(1).ListIndex = 0
    
    cbo(3).addItem "����"
    cbo(3).addItem "Сʱ"
    cbo(3).addItem "��"
    cbo(3).ListIndex = 0
    
    dtp(0).value = Format(Now, dtp(0).CustomFormat)
    dtp(1).value = Format(Now, dtp(0).CustomFormat)
    
    cbo(0).addItem "<��>"
    For mlngLoop = 101 To 110
        cbo(0).addItem GetWaveName(mlngLoop)
    Next
        
    cbo(0).ListIndex = 0
            
    tbs.Tab = 0
    
    InitData = True
    
End Function

Private Function SaveData(ByVal strFieldList As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim strSQL()  As String
    Dim lngKey As Long
    Dim strNote As String
    
    ReDim strSQL(1 To 1)
    
    If mlngKey = 0 Then
        
        rs.Open "SELECT Nvl(MAX(���),0)+1 AS ��� FROM ZLNOTICES", gcnOracle
        If rs.BOF Then Exit Function
        
        lngKey = rs("���").value
        
        strSQL(ReDimArray(strSQL)) = "ZL_ZLNOTICES_INSERT(" & lngKey & "," & _
                                                            IIf(mlngSys = 0, "NULL", mlngSys) & ",'" & _
                                                            ReplaceAll(txt(1).Text, "'", "''") & "','" & _
                                                            txt(0).Text & "'," & _
                                                            IIf(chk(3).value = 0, "NULL", "'" & txt(2).Tag & "'") & "," & _
                                                            IIf(chk(2).value = 0, "NULL", GetWaveCode(cbo(0).Text)) & "," & _
                                                            chk(1).value & "," & _
                                                            IIf(cbo(2).ListIndex = 0, "NULL", CalcTimeToSecend(Val(txt(3).Text), cbo(1).Text)) & "," & _
                                                            IIf(cbo(2).ListIndex = 0, "NULL", CalcTimeToSecend(Val(txt(4).Text), cbo(3).Text)) & "," & _
                                                            "TO_DATE('" & Format(dtp(0).value, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')," & _
                                                            IIf(chk(0).value = 0, "NULL", "TO_DATE('" & Format(dtp(1).value, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')") & ",'" & strFieldList & "')"
        
    Else
        lngKey = mlngKey
        strSQL(ReDimArray(strSQL)) = "ZL_ZLNOTICES_UPDATE(" & lngKey & "," & _
                                                            IIf(mlngSys = 0, "NULL", mlngSys) & ",'" & _
                                                            ReplaceAll(txt(1).Text, "'", "''") & "','" & _
                                                            txt(0).Text & "'," & _
                                                            IIf(chk(3).value = 0, "NULL", "'" & txt(2).Tag & "'") & "," & _
                                                            IIf(chk(2).value = 0, "NULL", GetWaveCode(cbo(0).Text)) & "," & _
                                                            chk(1).value & "," & _
                                                            IIf(cbo(2).ListIndex = 0, "NULL", CalcTimeToSecend(Val(txt(3).Text), cbo(1).Text)) & "," & _
                                                            IIf(cbo(2).ListIndex = 0, "NULL", CalcTimeToSecend(Val(txt(4).Text), cbo(3).Text)) & "," & _
                                                            "TO_DATE('" & Format(dtp(0).value, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')," & _
                                                            IIf(chk(0).value = 0, "NULL", "TO_DATE('" & Format(dtp(1).value, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')") & ",'" & strFieldList & "')"
                                                            
        strSQL(ReDimArray(strSQL)) = "ZL_ZLNOTICEUSR_DELETE(" & lngKey & ")"
    End If
    
    If chk(4).value = 0 Then
        For mlngLoop = 1 To lvwDept.ListItems.Count
            strSQL(ReDimArray(strSQL)) = "ZL_ZLNOTICEUSR_INSERT(" & lngKey & ",2,'" & lvwDept.ListItems(mlngLoop).Text & "')"
        Next
        
        For mlngLoop = 1 To lvwHuman.ListItems.Count
            strSQL(ReDimArray(strSQL)) = "ZL_ZLNOTICEUSR_INSERT(" & lngKey & ",1,'" & lvwHuman.ListItems(mlngLoop).Text & "')"
        Next
        
        For mlngLoop = 1 To lvwStation.ListItems.Count
            strSQL(ReDimArray(strSQL)) = "ZL_ZLNOTICEUSR_INSERT(" & lngKey & ",3,'" & lvwStation.ListItems(mlngLoop).Text & "')"
        Next
    End If
    
    On Error GoTo errHand
    gcnOracle.BeginTrans
    For mlngLoop = 1 To UBound(strSQL)
        If strSQL(mlngLoop) <> "" Then ExecuteProcedure strSQL(mlngLoop), Me.Caption
    Next
    gcnOracle.CommitTrans
    
    SaveData = True
    If mlngKey = 0 Then
        '������Ҫ������־
        Call SaveAuditLog(1, "����", "������ѳɹ�����������Ϊ��" & txt(0).Text & "��")
    Else
        If txt(0).Text <> txt(0).Tag Then
            '������Ҫ������־
            Call SaveAuditLog(2, "�޸�", "�޸����ѳɹ������������ɡ�" & txt(0).Tag & "���޸�Ϊ��" & txt(0).Text & "��")
        Else
            '������Ҫ������־
            Call SaveAuditLog(2, "�޸�", "���ѡ�" & txt(0).Text & "���޸ĳɹ�")
        End If
    End If
    Exit Function
    
errHand:
    gcnOracle.RollbackTrans
    MsgBox "����������Ϣʧ�ܣ�" & vbNewLine & err.Description, vbInformation, gstrSysName
End Function

Private Function CheckDataValid(ByRef strFieldList As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim intPos As Long
    Dim intPos1 As Long
    Dim intStart As Long
    Dim strFieldName As String
    Dim strSQL As String
    
    If Trim(txt(0).Text) = "" Then
        MsgBox "��������Ҫ���ѵ����ݣ�", vbInformation, gstrSysName
        tbs.Tab = 0
        txt(0).SetFocus
        Exit Function
    End If
    
    If StrIsValid(txt(0).Text, txt(0).MaxLength) = False Then
        tbs.Tab = 0
        txt(0).SetFocus
        Exit Function
    End If
    
    If Trim(txt(1).Text) <> "" Then
        If ValidSQL(txt(1).Text) = False Then
            MsgBox " ��������SQL�ǷǷ��ģ� ", vbInformation, gstrSysName
            tbs.Tab = 0
            txt(1).SetFocus
            Exit Function
        End If
    End If
    
    If LenB(StrConv(txt(1).Text, vbFromUnicode)) > txt(1).MaxLength Then
        MsgBox "���������ݲ��ܳ���" & Int(txt(1).MaxLength / 2) & "������" & "��" & txt(1).MaxLength & "����ĸ��", vbExclamation, gstrSysName
        tbs.Tab = 0
        txt(1).SetFocus
        Exit Function
    End If
    
    '������������е��ֶ��Ƿ��������������ֶ�֮��
    If txt(1).Text <> "" Then
        strSQL = txt(1).Text
        strSQL = Replace(UCase(strSQL), "[USER]", "'USER'")
        strSQL = "Select * From (" & strSQL & ") Where 1=2"
        
        rs.Open strSQL, gcnOracle
        
    End If

    intStart = 1
    intPos = InStr(intStart, txt(0).Text, "[")

    Do While intPos > 0

        intPos1 = InStr(intStart + 1, txt(0).Text, "]")
        If intPos1 > 0 Then

            strFieldName = Mid(txt(0).Text, intPos + 1, intPos1 - intPos - 1)
            If rs.State = adStateOpen Then
                For mlngLoop = 0 To rs.Fields.Count - 1
                    If rs.Fields(mlngLoop).Name = strFieldName Then
                        
                        strFieldList = strFieldList & "|[" & strFieldName & "];" & ConvertOracleType(rs.Fields(mlngLoop).type)
                        
                        Exit For
                    End If
                Next
                
'                If mlngLoop = rs.Fields.Count Then
'                    MsgBox "����������ָ�����ֶ������������в����ڣ�", vbInformation, gstrSysName
'                    tbs.Tab = 0
'                    txt(0).SetFocus
'                    Exit Function
'                End If
            Else
'                MsgBox "����������ָ�����ֶ������������в����ڣ�", vbInformation, gstrSysName
'                tbs.Tab = 0
'                txt(0).SetFocus
'                Exit Function
            End If
        Else
            Exit Do
        End If

        intStart = intPos1 + 1

        intPos = InStr(intStart, txt(0).Text, "[")
    Loop
    
    If strFieldList <> "" Then strFieldList = Mid(strFieldList, 2)
    
    CheckDataValid = True
    
End Function

Private Function ConvertOracleType(ByVal rsDataType As DataTypeEnum) As String
    
    '-------------------------------------------------------------------------------------
    '����:��vb��¼��������ת��Ϊoracle����������
    '-------------------------------------------------------------------------------------
    
    Select Case rsDataType
    
    Case adBigInt, adInteger, adCurrency, adTinyInt, adSmallInt
        ConvertOracleType = "NUMBER"
    Case adVarChar, adWChar, adVarWChar, adChar
        ConvertOracleType = "VARCHAR2"
    Case adDate
        ConvertOracleType = "DATE"
    Case Else
        ConvertOracleType = "NUMBER"
    End Select
    
End Function

Public Function ShowEdit(ByVal frmMain As Object, ByVal lngKey As Long, ByVal lngSys As Long, ByVal str������ As String, Optional ByVal blnBrowse As Boolean = False) As Boolean
    
    mblnOk = False
    
    mlngKey = lngKey
    mlngSys = lngSys
    mstr������ = str������
    
    If InitData = False Then Exit Function
            
    If mlngKey > 0 Then Call ReadData
                        
    If blnBrowse Then
        txt(0).Locked = True
        txt(1).Locked = True
        txt(2).Locked = True
        txt(3).Locked = True
        txt(4).Locked = True
        chk(0).Enabled = False
        chk(1).Enabled = False
        chk(2).Enabled = False
        chk(3).Enabled = False
        chk(4).Enabled = False
        
        cbo(0).Locked = True
        cbo(1).Locked = True
        cbo(2).Locked = True
        cbo(3).Locked = True
        
        dtp(0).Enabled = False
        dtp(1).Enabled = False
        
        cmdAdd(0).Enabled = False
        cmdAdd(1).Enabled = False
        cmdAdd(2).Enabled = False
        
        cmdRemove(0).Enabled = False
        cmdRemove(1).Enabled = False
        cmdRemove(2).Enabled = False
        
        cmdValid.Enabled = False
        
        cmdOpen.Enabled = False
        cmdHear.Enabled = False
        
        udn(0).Enabled = False
        udn(1).Enabled = False
    End If
    
    cmdAdd(0).Enabled = False
    cmdAdd(1).Enabled = False
    
    Call chk_Click(4)
                        
    cmdOK.Tag = ""
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOk
    
End Function

Private Sub cbo_Click(Index As Integer)
    
    Dim blnVisible As Boolean
    
    cmdOK.Tag = "Changed"
    
    Select Case Index
    Case 2
        blnVisible = (cbo(Index).ListIndex = 1)
        txt(3).Visible = blnVisible
        txt(4).Visible = blnVisible
        cbo(1).Visible = blnVisible
        cbo(3).Visible = blnVisible
        
        udn(0).Visible = blnVisible
        udn(1).Visible = blnVisible
        
        lbl(3).Visible = blnVisible
        lbl(8).Visible = blnVisible
        
    End Select
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
        If Index = 0 Then SendKeys "{TAB}"
    End If
End Sub

Private Sub chk_Click(Index As Integer)
    cmdOK.Tag = "Changed"
    
    Select Case Index
    Case 0
        dtp(1).Enabled = Not (chk(Index).value = 0)
        pic.Visible = Not dtp(1).Enabled
    Case 2
        cbo(0).Enabled = Not (chk(Index).value = 0)
        cmdHear.Enabled = cbo(0).Enabled
    Case 3
        txt(2).Enabled = Not (chk(Index).value = 0)
        cmdOpen.Enabled = txt(2).Enabled
    Case 4
        lvwDept.Enabled = (chk(Index).value = 0)
        lvwHuman.Enabled = lvwDept.Enabled
        lvwStation.Enabled = lvwDept.Enabled
        
        If mlngSys > 0 Then
            cmdAdd(0).Enabled = lvwDept.Enabled
            cmdAdd(1).Enabled = lvwDept.Enabled
        End If
        cmdAdd(2).Enabled = lvwDept.Enabled
        
        cmdRemove(0).Enabled = lvwDept.Enabled
        cmdRemove(1).Enabled = lvwDept.Enabled
        cmdRemove(2).Enabled = lvwDept.Enabled
        
    End Select
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cmdAdd_Click(Index As Integer)
    Dim objItem As ListItem
    Dim objPoint As POINTAPI
    Dim rs As New ADODB.Recordset
    Dim strInput As String
    
    Call ClientToScreen(cmdAdd(Index).hwnd, objPoint)
    
    Select Case Index
    Case 0
        gstrSQL = "SELECT -1 AS ID,0 AS �ϼ�id,'���в���' AS ����,'' AS ���� FROM DUAL " & _
                    "UNION ALL " & _
                    "SELECT ID,DECODE(�ϼ�id,NULL,-1,�ϼ�id) AS �ϼ�id,���� AS ����,���� FROM " & mstr������ & ".���ű� START WITH �ϼ�id IS NULL CONNECT BY PRIOR ID=�ϼ�id "
        rs.Open gstrSQL, gcnOracle
        If frmSelectTree.ShowSelect(Me, rs, objPoint.X * 15 - 30, objPoint.Y * 15 + cmdAdd(Index).Height - 30, 3000, 3900, cmdAdd(Index).Height, , Me.Name & "\����ѡ��", "������ϵ") Then
            
            If rs("ID").value > 0 Then
                On Error Resume Next
                
                Set objItem = lvwDept.ListItems.Add(, "K" & rs("����").value, rs("����").value, "Dept", "Dept")
                objItem.Selected = True
                objItem.EnsureVisible
            Else
                MsgBox "��ѡ������һ�����ţ�", vbInformation, gstrSysName
            End If
            
            cmdOK.Tag = "Changed"
        End If
        lvwDept.SetFocus
    Case 1
        
        gstrSQL = "SELECT ID,�ϼ�id,'' AS ���,'['||����||']'||���� AS ����,'' AS �Ա�,'' AS ����,0 AS ĩ�� FROM " & mstr������ & ".���ű� START WITH �ϼ�id IS NULL CONNECT BY PRIOR ID=�ϼ�id"
                
        gstrSQL = gstrSQL & " union all " & _
                  "select A.ID,B.����id AS �ϼ�id,A.���,A.���� AS ����,A.�Ա�,A.����,1 AS ĩ�� FROM " & mstr������ & ".��Ա�� A," & mstr������ & ".������Ա B WHERE (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) And B.ȱʡ=1 AND A.ID=B.��Աid"
                    
        rs.Open gstrSQL, gcnOracle
        
        If frmSelectExplorer.ShowSelect(Me, rs, objPoint.X * 15 - 30, objPoint.Y * 15 + cmdAdd(Index).Height - 30, 3900, 3900, cmdAdd(Index).Height, _
                                    "������Աѡ��", "���,1080,0,1;����,900,0,2", "������Ա") Then
            
            On Error Resume Next
            
            Set objItem = lvwHuman.ListItems.Add(, "K" & rs("����").value, rs("����").value, "Human", "Human")
            objItem.Selected = True
            objItem.EnsureVisible
            
            cmdOK.Tag = "Changed"
        End If
        
        lvwHuman.SetFocus
    Case 2
        If frmInputBox.ShowEdit(Me, strInput, "����վ������", "�����빤��վ�㣨��������������ơ�", "����վ��", 50) Then
        
            On Error Resume Next
            
            Set objItem = lvwStation.ListItems.Add(, "K" & strInput, strInput, "Station", "Station")
            objItem.Selected = True
            objItem.EnsureVisible
            
            cmdOK.Tag = "Changed"
        End If
    End Select
    
End Sub

Private Sub cmdClear_Click()
    Unload Me
End Sub

Private Sub cmdHear_Click()
    
    If cbo(0).Text = "" Then Exit Sub
    
    Call PlayWave(GetWaveCode(cbo(0).Text))
    
    cbo(0).SetFocus
    
End Sub

Private Sub cmdOK_Click()
    Dim strFieldList As String
    
    If Not txt(1).Locked Then   '���༭��ʱ��,�������ܼ��,���鿴����Ҫ����
        If gblnSystemUser Then
            If CheckSQLPlan(ReplaceAll(UCase(txt(1).Text), "[USER]", "'USER'")) = True Then
                If MsgBox("��ǰ����Դ�п��ܴ����������⣬�Ƿ�鿴ִ�мƻ���" & vbCrLf & "�����������档", vbQuestion + vbYesNo + vbDefaultButton2, "���ܼ��") = vbYes Then
                    frmSQLPlanEx.ShowMe Me, ReplaceAll(UCase(txt(1).Text), "[USER]", "'USER'")
                    Exit Sub
                End If
            End If
        End If
    End If
    
    If cmdOK.Tag <> "" Then
        
        If CheckDataValid(strFieldList) = False Then Exit Sub
        
        If SaveData(strFieldList) Then
            
            mblnOk = True
            
            If mlngKey = 0 Then
                
                txt(0).Text = ""
                txt(1).Text = ""
                txt(2).Text = ""
                txt(2).Tag = ""
                
                lvwDept.ListItems.Clear
                lvwHuman.ListItems.Clear
                lvwStation.ListItems.Clear
                
                cmdOK.Tag = ""
                tbs.Tab = 0
                txt(1).SetFocus
                
                Exit Sub
            Else
                cmdOK.Tag = ""
            End If
        End If
    End If
    
    Unload Me
End Sub

Private Sub cmdOpen_Click()
    
    Dim objPoint As POINTAPI
    Dim rs As New ADODB.Recordset
    Dim strInput As String
    
    Call ClientToScreen(txt(2).hwnd, objPoint)
    
    On Error GoTo errHand
    
    Set rs = OpenCursor(gcnOracle, "ZLTOOLS.B_Expert.Get_noticereport", mlngSys)
    If rs.BOF = False Then
        If frmSelectList.ShowSelect(Me, rs, "���,1500,0,1;����,1800,0,0;˵��,1500,0,0", objPoint.X * 15 - 30, objPoint.Y * 15 + cmdOpen.Height - 30, txt(2).Width + cmdOpen.Width, 3900, Me.Name & "\����ѡ��", "��ѡ����") Then
            
            txt(2).Text = rs("����").value
            txt(2).Tag = rs("���").value
            
            cmdOK.Tag = "Changed"
        End If
    Else
        MsgBox "���κα�����ã�", vbInformation, gstrSysName
    End If
    
    Exit Sub
    
errHand:
    
    MsgBox "��ȡ����ʱ����" & vbNewLine & err.Description, vbInformation, gstrSysName
    err.Clear
    
End Sub

Private Sub cmdPlan_Click()
    Dim strCaption As String
    
    If Not gblnSystemUser Then
        MsgBox "��ǰ�û�����ϵͳ�����ߣ��޷����ִ�мƻ�", , "��ʾ"
        Exit Sub
    End If
    If Not ValidSQL(txt(1).Text, strCaption) Then
        MsgBox " ��������SQL�ǷǷ��ģ� " & vbNewLine & strCaption, vbInformation, gstrSysName
        Exit Sub
    End If

    frmSQLPlanEx.ShowMe Me, ReplaceAll(UCase(txt(1).Text), "[USER]", "'USER'")

End Sub

Private Sub cmdRemove_Click(Index As Integer)
    Dim lngIndex As Long
    
    Select Case Index
    Case 0
        If lvwDept.SelectedItem Is Nothing Then Exit Sub
        
        lngIndex = lvwDept.SelectedItem.Index
        lvwDept.ListItems.Remove lvwDept.SelectedItem.Index
        Call NextLvwPos(lvwDept, lngIndex)
        
        cmdOK.Tag = "Changed"
    Case 1
        If lvwHuman.SelectedItem Is Nothing Then Exit Sub
        
        lngIndex = lvwHuman.SelectedItem.Index
        lvwHuman.ListItems.Remove lvwHuman.SelectedItem.Index
        Call NextLvwPos(lvwHuman, lngIndex)
        
        cmdOK.Tag = "Changed"
    Case 2
        If lvwStation.SelectedItem Is Nothing Then Exit Sub
        
        lngIndex = lvwStation.SelectedItem.Index
        lvwStation.ListItems.Remove lvwStation.SelectedItem.Index
        Call NextLvwPos(lvwStation, lngIndex)
        
        cmdOK.Tag = "Changed"
    End Select
End Sub

Private Sub cmdValid_Click()
    Dim strTmp As String
    Dim strErr As String
    
    If ValidSQL(txt(1).Text, strErr) Then
        MsgBox " ��������SQL�ǺϷ��ģ� ", vbInformation, gstrSysName
    Else
        MsgBox " ��������SQL�ǷǷ��ģ� " & vbNewLine & strErr, vbInformation, gstrSysName
    End If
    
    strTmp = CheckFullTable(txt(1).Text)
    If strTmp <> "" Then
        MsgBox "ע�⣬���±�Ϊȫ��ɨ�裺     " & vbCrLf & strTmp, vbInformation, gstrSysName
    End If
End Sub

Private Sub dtp_Change(Index As Integer)
    cmdOK.Tag = "Changed"
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdOK.Tag <> "" Then
        Cancel = (MsgBox("�޸ĺ���Զ����ѱ��뱣������Ч���Ƿ�������棿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo)
    End If
End Sub

Private Sub lvwDept_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
        SendKeys "{TAB}"
        SendKeys "{TAB}"
    End If
End Sub

Private Sub lvwHuman_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
        SendKeys "{TAB}"
        SendKeys "{TAB}"
    End If
End Sub

Private Sub lvwStation_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
        SendKeys "{TAB}"
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txt_Change(Index As Integer)
    cmdOK.Tag = "Changed"
End Sub

Private Sub txt_GotFocus(Index As Integer)
    txt(Index).SelStart = 0
    txt(Index).SelLength = Len(txt(Index).Text)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If Index = 1 Then
            tbs.Tab = 1
            chk(4).SetFocus
            Exit Sub
        End If
        SendKeys "{TAB}"
        If Index = 2 Then SendKeys "{TAB}"
        
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    If Index <> 1 Then
        Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    Else
        If LenB(StrConv(txt(Index).Text, vbFromUnicode)) > txt(Index).MaxLength And txt(Index).MaxLength > 0 Then
            MsgBox "���������ݲ��ܳ���" & Int(txt(Index).MaxLength / 2) & "������" & "��" & txt(Index).MaxLength & "����ĸ��", vbExclamation, gstrSysName
            Cancel = True
        End If
    End If
    
    If Cancel Then Exit Sub
    
    Select Case Index
    Case 3
        If Val(txt(Index).Text) < udn(0).Min Or Val(txt(Index).Text) > udn(0).Max Then
        
            Cancel = True
            MsgBox "������ڱ�����" & udn(0).Min & "��" & udn(0).Max & "֮�䣡", vbInformation, gstrSysName
            
        End If
    Case 4
        If Val(txt(Index).Text) < udn(1).Min Or Val(txt(Index).Text) > udn(1).Max Then
            Cancel = True
            MsgBox "�������ڱ�����" & udn(1).Min & "��" & udn(1).Max & "֮�䣡", vbInformation, gstrSysName
        End If
    End Select
    
End Sub
