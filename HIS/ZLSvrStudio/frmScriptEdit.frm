VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmScriptEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���  �������ļ���"
   ClientHeight    =   7464
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7788
   Icon            =   "frmScriptEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7464
   ScaleWidth      =   7788
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraBound 
      Height          =   30
      Index           =   1
      Left            =   -555
      TabIndex        =   38
      Top             =   1185
      Width           =   8520
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1200
      Left            =   -15
      ScaleHeight     =   1315.068
      ScaleMode       =   0  'User
      ScaleWidth      =   7800
      TabIndex        =   39
      Top             =   0
      Width           =   7800
      Begin VB.Label lblEXP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���Ӱ�װ·�����ɲ����ļ����ͻ��˶��λ�ã�[*]Ϊͨ�����|�ָ���·��"
         Height          =   180
         Index           =   3
         Left            =   1395
         TabIndex        =   43
         Top             =   915
         Width           =   6120
      End
      Begin VB.Label lblEXP 
         BackStyle       =   0  'Transparent
         Caption         =   "��Ҫע�᣺DLL��EXE�ļ�����Ҫע�ᣬ����ע�Ὣ���²�������ʹ��"
         Height          =   195
         Index           =   2
         Left            =   1395
         TabIndex        =   42
         Top             =   645
         Width           =   5850
      End
      Begin VB.Label lblEXP 
         BackStyle       =   0  'Transparent
         Caption         =   "��װ·���������ļ��������ͻ���λ�ã����ļ�λ���Զ����ɣ������޸�"
         Height          =   210
         Index           =   1
         Left            =   1395
         TabIndex        =   41
         Top             =   375
         Width           =   6210
      End
      Begin VB.Label lblEXP 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "�ļ�λ�ã��ļ�λ�þ�����װ·�������ļ�����λ�ã������Ʒ���"
         Height          =   180
         Index           =   0
         Left            =   1395
         TabIndex        =   40
         Top             =   105
         Width           =   5220
      End
      Begin VB.Image imgCaption 
         Height          =   576
         Left            =   480
         Picture         =   "frmScriptEdit.frx":6852
         Top             =   228
         Width           =   576
      End
   End
   Begin VB.Frame fraBound 
      Height          =   30
      Index           =   0
      Left            =   -255
      TabIndex        =   37
      Top             =   6870
      Width           =   8820
   End
   Begin VB.PictureBox picUniversalPathADD 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1365
      ScaleHeight     =   276
      ScaleWidth      =   5976
      TabIndex        =   35
      Top             =   2340
      Width           =   6000
      Begin VB.TextBox txtUniversalPathADD 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   45
         TabIndex        =   36
         ToolTipText     =   "[*]����ͨ�����|�ָ���·������[APPSOFT]\Dev_[*]|[APPSOFT]\MyTest\[*]\A"
         Top             =   45
         Width           =   5880
      End
   End
   Begin VB.PictureBox picExplanation 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1425
      Left            =   1365
      ScaleHeight     =   1404
      ScaleWidth      =   5976
      TabIndex        =   32
      Top             =   5160
      Width           =   6000
      Begin VB.TextBox txtExplanation 
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   1375
         Left            =   15
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Top             =   15
         Width           =   5950
      End
   End
   Begin VB.PictureBox picVision 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3960
      ScaleHeight     =   276
      ScaleWidth      =   1980
      TabIndex        =   29
      Top             =   2835
      Width           =   2000
      Begin VB.TextBox txtVision 
         BorderStyle     =   0  'None
         ForeColor       =   &H80000011&
         Height          =   285
         Left            =   45
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   45
         Width           =   1850
      End
   End
   Begin VB.PictureBox picCbo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1365
      ScaleHeight     =   252
      ScaleWidth      =   1632
      TabIndex        =   27
      Top             =   2850
      Width           =   1650
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   -30
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   -30
         Width           =   1700
      End
   End
   Begin VB.PictureBox picUniversalPath 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1365
      ScaleHeight     =   276
      ScaleWidth      =   5976
      TabIndex        =   25
      Top             =   1845
      Width           =   6000
      Begin VB.TextBox txtUniversalPath 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000011&
         Height          =   330
         Left            =   30
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   45
         Width           =   5900
      End
   End
   Begin VB.PictureBox picFilePath 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1365
      ScaleHeight     =   276
      ScaleWidth      =   5976
      TabIndex        =   23
      Top             =   1320
      Width           =   6000
      Begin VB.TextBox txtFilePath 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000011&
         Height          =   330
         Left            =   45
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   45
         Width           =   5900
      End
   End
   Begin VB.CheckBox chk 
      Caption         =   "ǿ���滻"
      Height          =   315
      Index           =   2
      Left            =   30
      TabIndex        =   22
      Top             =   5865
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Height          =   300
      Index           =   8
      Left            =   0
      Picture         =   "frmScriptEdit.frx":8394
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "���"
      Top             =   7005
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CommandButton cmd 
      Height          =   300
      Index           =   7
      Left            =   0
      Picture         =   "frmScriptEdit.frx":EBE6
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "ȫѡ"
      Top             =   7005
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CommandButton cmd 
      Height          =   300
      Index           =   6
      Left            =   -15
      Picture         =   "frmScriptEdit.frx":15438
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "��ѡ"
      Top             =   7020
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CommandButton cmd 
      Height          =   300
      Index           =   5
      Left            =   45
      Picture         =   "frmScriptEdit.frx":1BC8A
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "ȫ��"
      Top             =   7005
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CheckBox chk 
      Caption         =   "Ӧ��������ϵͳ"
      Height          =   315
      Index           =   1
      Left            =   30
      TabIndex        =   13
      Top             =   5520
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.CheckBox chk 
      Caption         =   "��Ҫע��"
      Height          =   210
      Index           =   0
      Left            =   6345
      TabIndex        =   12
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Height          =   300
      Index           =   4
      Left            =   7380
      Picture         =   "frmScriptEdit.frx":224DC
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "ȫ��"
      Top             =   3360
      Width           =   300
   End
   Begin VB.CommandButton cmd 
      Height          =   300
      Index           =   3
      Left            =   45
      Picture         =   "frmScriptEdit.frx":28D2E
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "��ѡ"
      Top             =   7020
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CommandButton cmd 
      Height          =   300
      Index           =   2
      Left            =   60
      Picture         =   "frmScriptEdit.frx":2F580
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "ȫѡ"
      Top             =   7005
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6285
      TabIndex        =   7
      Top             =   7020
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "����(&S)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4965
      TabIndex        =   6
      Top             =   7020
      Width           =   1100
   End
   Begin VB.CommandButton cmd 
      Height          =   300
      Index           =   1
      Left            =   75
      Picture         =   "frmScriptEdit.frx":35DD2
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "ѡ��λ��"
      Top             =   7005
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CommandButton cmd 
      Height          =   300
      Index           =   0
      Left            =   7380
      Picture         =   "frmScriptEdit.frx":3C624
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "ѡ���ļ�"
      Top             =   1320
      Width           =   300
   End
   Begin MSComDlg.CommonDialog Cdlg 
      Left            =   60
      Top             =   3780
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfCom 
      Height          =   1125
      Left            =   1380
      TabIndex        =   15
      Top             =   5250
      Visible         =   0   'False
      Width           =   1365
      _cx             =   2408
      _cy             =   1984
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   14737632
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   20
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmScriptEdit.frx":42E76
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   1
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   30
   End
   Begin MSComctlLib.ListView lvwSys 
      Height          =   1620
      Left            =   1365
      TabIndex        =   31
      Top             =   3360
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   2858
      View            =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "���"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   2
      Left            =   6330
      TabIndex        =   45
      Top             =   3120
      Width           =   1590
   End
   Begin VB.Label lblWarning 
      AutoSize        =   -1  'True
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   1
      Left            =   1395
      TabIndex        =   44
      Top             =   7125
      Width           =   90
   End
   Begin ComctlLib.ImageList imgList 
      Left            =   615
      Top             =   3735
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmScriptEdit.frx":42EB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmScriptEdit.frx":44A04
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "���Ӱ�װ·��"
      Height          =   225
      Index           =   6
      Left            =   165
      TabIndex        =   34
      Top             =   2415
      Width           =   1155
   End
   Begin VB.Label lblWarning 
      AutoSize        =   -1  'True
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   1365
      TabIndex        =   21
      Top             =   1635
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�ļ�˵��"
      Height          =   180
      Index           =   5
      Left            =   510
      TabIndex        =   20
      Top             =   5160
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ҵ�񲿼�"
      Height          =   180
      Index           =   4
      Left            =   525
      TabIndex        =   14
      Top             =   5160
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�汾��"
      Height          =   180
      Index           =   3
      Left            =   3375
      TabIndex        =   11
      Top             =   2910
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����ϵͳ"
      Height          =   180
      Index           =   2
      Left            =   510
      TabIndex        =   5
      Top             =   3375
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   1
      Left            =   510
      TabIndex        =   4
      Top             =   2910
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��װ·��"
      Height          =   180
      Index           =   0
      Left            =   525
      TabIndex        =   2
      Top             =   1920
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�ļ�λ��"
      Height          =   180
      Index           =   0
      Left            =   525
      TabIndex        =   0
      Top             =   1395
      Width           =   720
   End
End
Attribute VB_Name = "frmScriptEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_blnModed              As Boolean
Private m_str��ʽ               As String
Private m_strNum               As String
Private m_strPathJY             As String
Private m_strEditDate           As String
Private m_lngCurRow             As Long
Private m_strCurFileName        As String
Private mstr���                As String
Private m_strLocationName As String
Private mcllPath As Collection '��װ·��ת��ʵ��·������
Private mblnLoad As Boolean '��ʼ����־
Private mblnCheckRegist As Boolean  '�Զ�ע����ɹ���־
Private mblnRegist As Boolean '�Զ�ע���жϱ�־
Private mblnFileExist As Boolean '�����ļ��Ƿ����
Private mblnFileRepeat As Boolean '���ݿ��Ѿ����ڸ��ļ���־
Private mblnFileEsexpired As Boolean '�����ļ���־

Private Const FW_����1 = "���ļ����ļ��嵥���Ѿ����ڣ�"
Private Const FW_����2 = "���ļ��Ѿ����ã�"
Private Const FW_����3 = "���ļ������ļ�ȱʧ��"
Private Const FW_����4 = "���ļ��Զ�ע����ʧ�ܣ��������ж��Ƿ���Ҫע��"

Public Property Get Moded() As Boolean
   Moded = m_blnModed
End Property

Public Property Let Moded(ByVal blnModed As Boolean)
    m_blnModed = blnModed
End Property

Private Sub cbo_Click(Index As Integer)
    Dim i As Long
    On Error GoTo errH
    Select Case Index
    Case 0
        
    Case 1
        If m_str��ʽ = "����" Then
            If cbo(1).Text = "��������" Then
                Call cmd_Click(2)
            Else
'                Me.Caption = "��ӡ�" & cbo(1).Text & "��"
                Me.Caption = cbo(1).Text & " - ����"
                For i = 1 To lvwSys.ListItems.Count
                    If lvwSys.ListItems.Item(i).SubItems(1) = m_strNum Then
                        lvwSys.ListItems.Item(i).Checked = True
                    Else
                        lvwSys.ListItems.Item(i).Checked = False
                    End If
                Next
            End If
        End If
        
        If cbo(1).Text = "ϵͳ�ļ�" Then
            chk(2).Visible = True
        Else
            chk(2).Visible = False
        End If
    End Select
    Exit Sub
errH:

End Sub

Private Sub chk_Click(Index As Integer)
    Dim i As Integer
    Dim strErr As String
    
    Select Case Index
        Case 0 '����ע��
            If chk(0).value = 1 Then
                picUniversalPathADD.Enabled = False
                Label1(6).Enabled = False
                txtUniversalPathADD.Text = ""
'                txtUniversalPathADD.Tag = "0" '��ʾ��־
            Else
                picUniversalPathADD.Enabled = True
                Label1(6).Enabled = True
            End If
            If mblnLoad Then Exit Sub
'            If InStr(lbl(1).Caption, FW_����1) > 0 Then
'                If chk(0).value = 1 Then
'                    If MsgBox("���ļ��Ѿ����ڣ��ظ�ע�����ɻ����������ܹ�ѡע�ᣡ", vbInformation, gstrSysName) Then chk(0).value = 0: Exit Sub
'                End If
'            End If
            If Not mblnCheckRegist Then Exit Sub
            lblWarning(2).Caption = ""
            If mblnRegist And chk(0).value = 0 Then
                If MsgBox("���ļ�����Ҫע����ļ���ȡ������ܻ�����ļ�ע�᲻��ȷ���������Ƿ�ȷ��ȡ��ע�᣿", vbQuestion + vbOKCancel, gstrSysName) = vbCancel Then
                    chk(0).value = 1
                Else
                    lblWarning(2).Caption = "���ļ���Ҫע�ᣡ"
                End If
            End If
            If Not mblnRegist And chk(0).value = 1 Then
                If MsgBox("���ļ�������Ҫע����ļ���ע�����ܻ�����ļ�ע�᲻��ȷ���������Ƿ�ȷ����Ҫע�᣿", vbQuestion + vbOKCancel, gstrSysName) = vbCancel Then
                    chk(0).value = 0
                Else
                    lblWarning(2).Caption = "���ļ�����Ҫע�ᣡ"
                End If
            End If
        Case 1 'Ӧ��������ϵͳ
            If chk(1).value = 1 Then
                cbo(1).Enabled = False
                lvwSys.Enabled = False
                For i = 1 To lvwSys.ListItems.Count
                    lvwSys.ListItems.Item(i).Checked = True
                Next
            Else
                For i = 1 To lvwSys.ListItems.Count
                    If lvwSys.ListItems.Item(i).SubItems(1) = m_strNum Then
                        lvwSys.ListItems.Item(i).Checked = True
                    Else
                        lvwSys.ListItems.Item(i).Checked = False
                    End If
                Next
                cbo(1).Enabled = True
                lvwSys.Enabled = True
            End If
        Case 2 'ǿ�Ƹ���

    End Select
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim i As Long
    Dim strFilter   As String
    Dim strPath     As String
    Dim strSavePath As String
    
    Select Case Index
        Case 0 'ѡ���ļ�
'        Dim m_item As MSComctlLib.ListItem
            Dim clsPEFileCheck As New clsPEReader
            Dim objFile As New FileSystemObject
            Dim strErr As String
            strPath = OpenFile
          
            If Len(strPath) Then
                txtFilePath.Text = strPath
'                txtVision.Text = GetCommpentVersion(strPath)
                txtVision.Text = GetDealVersion(strPath)
                '����ļ�
                 CheckFile (strPath)
                '�����ļ�·������
                txtUniversalPath.Text = TransUniversalPath(strPath)
                
                If mblnRegist = True Then
                    chk(0).value = 1
                Else
                    chk(0).value = 0
                End If
             End If
    Case 1 'ѡ��λ��
        strSavePath = OpenFolder(Me)
        If strSavePath = "" Then
            Exit Sub
        Else
'            cbo(0).Text = strSavePath
             txtUniversalPath.Text = TransUniversalPath(strSavePath)
        End If
    Case 2 'ȫѡ
        For i = 1 To lvwSys.ListItems.Count
            lvwSys.ListItems.Item(i).Checked = True
        Next
    Case 3 '��ѡ
        For i = 1 To lvwSys.ListItems.Count
          If lvwSys.ListItems.Item(i).Checked Then
            lvwSys.ListItems.Item(i).Checked = False
          Else
            lvwSys.ListItems.Item(i).Checked = True
          End If
        Next
        
    Case 4 'ȫ��
        For i = 1 To lvwSys.ListItems.Count
            lvwSys.ListItems.Item(i).Checked = False
        Next
        txtExplanation.Text = ""
    Case 5 '��������ļ�
        Call StandardAllDel
    Case 6 'ɾ��Ӧ���ļ�
        Call StandardDel
    Case 7 '���Ӧ���ļ�
        Call AddFile
    Case 8 '���˵��
        txtExplanation.Text = ""
    End Select
End Sub

'==============================================================================
'=���ܣ�ȡ������
'==============================================================================
Private Sub cmdCancel_Click()
    If m_str��ʽ = "����" Then
        Moded = True
    Else
        Moded = False
    End If
    Unload Me
End Sub

'==============================================================================
'=���ܣ����湦��
'==============================================================================
Private Sub cmdOK_Click()
    Dim i As Long
    Dim blnSelect  As Boolean
    Dim lngTypeNum As Long
    Dim strPath    As String
    Dim strPathADD As String
    Dim strPathADDArr() As String
    Dim strTemp As String
    Dim ret        As Long
    On Error GoTo errH
    If txtFilePath = "" Then
        MsgBox "��ѡ���ļ�!", vbInformation, "��ʾ"
        txtFilePath.SetFocus
        Exit Sub
    End If
    
    For i = 1 To lvwSys.ListItems.Count
        If lvwSys.ListItems.Item(i).Checked Then
            blnSelect = True
            Exit For
        End If
    Next
    
    If Len(txtExplanation.Text) > 1900 Then
        MsgBox "�ļ�˵���벻Ҫ����2000���ַ�!", vbInformation, "��ʾ"
        txtExplanation.SetFocus
        Exit Sub
    End If
    
    If blnSelect = False Then
       MsgBox "��ѡ��ϵͳ���!", vbInformation, "��ʾ"
       lvwSys.SetFocus
       Exit Sub
    End If
    
    If txtUniversalPathADD.Text <> "" Then
        strPathADDArr() = Split(UCase(txtUniversalPathADD.Text), "|")
        For i = 0 To UBound(strPathADDArr)
            strTemp = strPathADDArr(i)
            If CheckUniversalADDPath(strTemp) = False Then
                MsgBox "���Ӱ�װ·�����Ϸ�������", vbInformation, gstrSysName
                txtUniversalPathADD.SetFocus
                Exit Sub
            End If
        Next
    End If
    
    strPathADD = UCase(txtUniversalPathADD.Text)
    strPath = txtUniversalPath.Text
    
    lngTypeNum = cbo(1).ItemData(cbo(1).ListIndex)

    If SaveDate(txtFilePath, lngTypeNum, strPath, strPathADD) Then
        If m_str��ʽ = "����" Then
            ret = MsgBox("����ɹ����Ƿ�������?", vbQuestion + vbYesNo, "��ʾ")
            If ret = vbYes Then
                txtFilePath.Text = ""
                txtFilePath.SetFocus
                lblWarning(0).Caption = "��ѡ���ļ�!"
                If chk(1).value = 0 Then
                    For i = 1 To lvwSys.ListItems.Count
                        If lvwSys.ListItems.Item(i).SubItems(1) = m_strNum Then
                            lvwSys.ListItems.Item(i).Checked = True
                        Else
                            lvwSys.ListItems.Item(i).Checked = False
                        End If
                    Next
                End If
            
                Exit Sub
            Else
                Call SaveSetting("zlSvrStudio", "parameter", "Type", cbo(1).Text)
                Moded = True
                Unload Me
            End If
        Else
            Call MsgBox("����ɹ���", vbInformation, gstrSysName)
            Call SaveSetting("zlSvrStudio", "parameter", "Type", cbo(1).Text)
            Moded = True
            Unload Me
        End If
    End If
    Exit Sub
errH:
 
End Sub

'==============================================================================
'=���ܣ�ҳ���ʼ��
'==============================================================================

'==============================================================================
'=���ܣ������ӿں��������ڴ����ʼ������:ID '��ʽΪ���룬��ID���ڣ�����IDֵǰ�ڵ���롣
'==============================================================================
Public Function ShowForm(��ʽ As String, ByVal �������� As String, ByVal �ļ����� As String, ByVal ����ϵͳ As String, ByVal ϵͳ�� As String, ByVal �汾�� As String, ByVal ��װ·�� As String, ByVal �޸����� As String, ByVal ����ϵͳNew As String, ByVal �ļ�˵�� As String, ByVal �����ļ� As String, ByVal �Զ�ע�� As Boolean, ByVal ǿ�Ƹ��� As Boolean, ByVal ��� As String, ByVal ���Ӱ�װ·�� As String) As String
    On Error GoTo errH
    Dim strPath As String
    Dim strType As String
    m_str��ʽ = ��ʽ
    m_strNum = ϵͳ��
    
    Set mcllPath = CheckAndAdjustFolder()
    mblnLoad = True
    
    If ��ʽ = "����" Then
        imgCaption.Picture = imgList.ListImages(1).Picture
        
        If ��� <> "0" Then
            mstr��� = ���
        Else
            mstr��� = "0"
        End If
        
        lbl(0).Caption = "�ļ�λ��"
        Call FillCboType
        Call ShowRowName
        Me.Caption = �������� & " - ����"
        
        '��ԭ�ϴ�ѡ���ֵ
        cmd(0).Enabled = True
        txtFilePath.Enabled = True
        strPath = GetSetting("zlSvrStudio", "parameter", "Path")
        strType = GetSetting("zlSvrStudio", "parameter", "Type")
        
        If strType <> "" Then
            cbo(1).Text = strType
        End If
        txtExplanation.Text = ""
        chk(0).value = 1
        chk(2).value = 0
        
        Call initvsfCom
    
    Else
        imgCaption.Picture = imgList.ListImages(2).Picture
        mstr��� = "0"
        lbl(0).Caption = "�ļ�����"
        
        Call FillCboType
        Call ShowRowName
        
        Me.Caption = �������� & " - �޸�"
        cmd(0).Enabled = False
        txtFilePath.Enabled = False
        
        txtUniversalPath.Text = IIf(��װ·�� = "0", "", ��װ·��)
        txtFilePath.Text = mcllPath("K_" & UCase(txtUniversalPath.Text)) & "\" & UCase(�ļ�����)
        '�ļ����
        CheckFile (txtFilePath.Text)
        
        txtUniversalPathADD.Text = ���Ӱ�װ·��
        m_strEditDate = �޸�����
        cbo(1).Text = ��������
        txtVision.Text = IIf(�汾�� = "0", "", �汾��)
        
        Dim i As Integer
        Dim j As Integer
        Dim strArr As Variant
        
        If ����ϵͳNew = "" Then
            For i = 1 To lvwSys.ListItems.Count
                lvwSys.ListItems.Item(i).Checked = True
            Next
        Else
            For i = 1 To lvwSys.ListItems.Count
                lvwSys.ListItems.Item(i).Checked = False
            Next
            
            strArr = Split(����ϵͳNew, ",")
            For i = 0 To UBound(strArr) - 1
                If strArr(i) <> "" Then
                    For j = 1 To lvwSys.ListItems.Count
                        If strArr(i) = lvwSys.ListItems.Item(j).SubItems(1) Then
                            lvwSys.ListItems.Item(j).Checked = True
                            Exit For
                        End If
                    Next
                End If
            Next
        End If
        
        If �ļ�˵�� = "0" Then
            txtExplanation.Text = ""
        Else
            txtExplanation.Text = �ļ�˵��
        End If
        
        If �Զ�ע�� Then
            chk(0).value = 1
        Else
            chk(0).value = 0
        End If
        
        If ǿ�Ƹ��� Then
            chk(2).value = 1
        Else
            chk(2).value = 0
        End If
        
        Call initvsfCom
        If Len(�����ļ�) > 0 Then
            Call refvsfCom(�����ļ�)
        End If
    End If
    mblnLoad = False
    Me.Show 1, frmMDIMain
    ShowForm = m_strLocationName
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Function

'��䰲װ·��Ĭ��ֵ
Private Sub FillCboPath()
    On Error GoTo errH
    With cbo(0)
        .Clear
        .AddItem "[AppSoft]"
        .ItemData(.NewIndex) = 0
        .AddItem "[System]"
        .ItemData(.NewIndex) = 1
        .AddItem "[Help]"
        .ItemData(.NewIndex) = 2
        .AddItem "[Public]"
        .ItemData(.NewIndex) = 3
        .ListIndex = 1
    End With
    Exit Sub
errH:

End Sub

'����ļ�����Ĭ��ֵ
Private Sub FillCboType()
    On Error GoTo errH
    With cbo(1)
        .Clear
        
        .AddItem "��������"
        .ItemData(.NewIndex) = 0
        .AddItem "Ӧ�ò���"
        .ItemData(.NewIndex) = 1
        .AddItem "�����ļ�"
        .ItemData(.NewIndex) = 2
        .AddItem "�����ļ�"
        .ItemData(.NewIndex) = 3
        .AddItem "��������"
        .ItemData(.NewIndex) = 4
        .AddItem "ϵͳ�ļ�"
        .ItemData(.NewIndex) = 5
        
        .ListIndex = 4
        
        cbo(1).Enabled = False
        
    End With
    Exit Sub
errH:

End Sub


''��ʾָ�������������
Private Sub ShowRowName()
  Dim strSQL As String, rs As New ADODB.Recordset
  Dim m_list As MSComctlLib.ListItem
  Dim i As Integer
  Dim str��� As String
  On Error GoTo errH
  lvwSys.ListItems.Clear
  strSQL = "select ����,��� from zlSystems order by ���"

  
   Call OpenRecordset(rs, strSQL, Me.Caption)
  
  If rs.RecordCount > 0 Then
    rs.MoveFirst
    Do Until rs.EOF
      str��� = Nvl(rs!���) \ 100
      Set m_list = lvwSys.ListItems.Add(, , "[" & str��� & "]" & Nvl(rs!����))
          m_list.SubItems(1) = str���
      rs.MoveNext
    Loop
  End If
  Exit Sub
errH:

End Sub

'��������
Private Function SaveDate(ByVal strFilePath As String, ByVal lngTypeNum As Long, ByVal strPath As String, ByVal strPathADD) As Boolean
    Dim rs          As New ADODB.Recordset
    Dim rsMaxID     As New ADODB.Recordset
    Dim rsShell     As New ADODB.Recordset
    Dim strSQL      As String
    Dim strName     As String '����
    Dim strVision   As String '�汾��
    Dim strEditDate As String '�޸�����
    Dim ret         As Long
    Dim strArr      As Variant
    Dim lng�汾��   As Double
    Dim i           As Long
    Dim strMax���  As String '������
    Dim strCurSelectSys As String
    Dim dateEdit    As Date  '�޸�����
    Dim lngSelectNum As Long 'ѡ����
    Dim bln����ϵͳ As Boolean
    
    
    Dim str����ϵͳ As String '����ϵͳ
    Dim str�ļ�˵�� As String '�ļ�˵��
    Dim str�����ļ� As String '�����ļ�
    Dim byt�Զ�ע�� As Byte
    Dim bytǿ�Ƹ��� As Byte
    
    On Error GoTo errH
    lngSelectNum = 0
    strName = GetFileName(strFilePath, , True)
    strSQL = "select �ļ���,����ϵͳ from zlFilesUpgrade where upper(�ļ���) = upper('" & strName & "') "
    Call OpenRecordset(rs, strSQL, Me.Caption)
    
    '���������
    If m_str��ʽ = "����" Then
        If mstr��� <> "0" Then
            strMax��� = CLng(Val(mstr���))
        Else
            strSQL = "select max(to_number(���)) as ��� from  zlFilesUpgrade"
             Call OpenRecordset(rsMaxID, strSQL, Me.Caption)
            If rsMaxID.RecordCount > 0 Then
                strMax��� = CLng(Nvl(rsMaxID!���, 0))
            Else
                strMax��� = 1
            End If
        End If
        
        '����޸�����
        dateEdit = Format(FileDateTime(strFilePath), "yyyy-MM-dd hh:mm:ss")
    Else
        dateEdit = Format(m_strEditDate, "yyyy-mm-dd hh:mm:ss")
    End If
    

    '�����ļ�
    '��ϴ洢�汾��
    strVision = txtVision.Text
    '��ǰѡ�������ϵͳ
    With lvwSys
        For i = 1 To .ListItems.Count
            If .ListItems.Item(i).Checked Then
                lngSelectNum = lngSelectNum + 1
                If strCurSelectSys = "" Then
                    strCurSelectSys = "," & .ListItems.Item(i).SubItems(1)
                Else
                    strCurSelectSys = strCurSelectSys & "," & .ListItems.Item(i).SubItems(1)
                End If
            End If
        Next
        If Len(strCurSelectSys) <> 0 Then
            strCurSelectSys = strCurSelectSys & ","
        End If
        If lngSelectNum = .ListItems.Count Then
            bln����ϵͳ = True
        Else
            bln����ϵͳ = False
        End If
    End With
    
    str�ļ�˵�� = txtExplanation.Text
    str�����ļ� = getFiles
    byt�Զ�ע�� = IIf(chk(0).value = 0, 0, 1)
    bytǿ�Ƹ��� = IIf(chk(2).value = 0, 0, 1)
    m_strLocationName = strName
    If rs.RecordCount > 0 Then
            '�޸�
            If bln����ϵͳ Then
                str����ϵͳ = ""
            Else
                If Nvl(rs!����ϵͳ) <> "" Then
                    str����ϵͳ = rs!����ϵͳ
                    str����ϵͳ = GetUnionSysNum(str����ϵͳ, strCurSelectSys)
                Else
                    str����ϵͳ = strCurSelectSys
                End If
            End If
            
            '����SQLִ��
            strSQL = "update Zlfilesupgrade set �ļ�����='" & lngTypeNum & "',�ļ��汾��='" & strVision & "',ҵ�񲿼�='" & str�����ļ� & "',����ϵͳ='" & str����ϵͳ & "',��װ·��='" & strPath & "',���Ӱ�װ·�� = '" & strPathADD & "'" & _
            ",�޸�����=" & IIf(IsDate(CStr(m_strEditDate)), "to_date('" & CStr(dateEdit) & "','yyyy-mm-dd hh24:mi:ss')", "''") & ",�ļ�˵��='" & str�ļ�˵�� & "',ǿ�Ƹ���=" & bytǿ�Ƹ��� & ",�Զ�ע��=" & byt�Զ�ע�� & " where upper(�ļ���)='" & UCase(strName) & "'"
            gcnOracle.Execute strSQL
            SaveDate = True
            '������Ҫ������־
            Call SaveAuditLog(1, "�ļ���������-�޸�", "�ɹ��޸�����Ϊ��" & strName & "���ĵ���������")
            Exit Function
    Else
        '����
        '����SQLִ��
        If bln����ϵͳ Then
            str����ϵͳ = ""
        Else
            str����ϵͳ = strCurSelectSys
        End If
        If mstr��� <> "0" Then
            strSQL = "update zlfilesupgrade set ���= ���+1 Where ���>" & Val(strMax���)
            gcnOracle.Execute strSQL
        End If
        strSQL = "insert into zlFilesUpgrade (���,�ļ�����,�ļ���,�ļ��汾��,�޸�����,ҵ�񲿼�,����ϵͳ,��װ·��,�ļ�˵��,ǿ�Ƹ���,�Զ�ע��,���Ӱ�װ·��) values (" & strMax��� + 1 & "," & lngTypeNum & "," & _
        "'" & strName & "','" & strVision & "',to_date('" & CStr(dateEdit) & "','yyyy-mm-dd hh24:mi:ss'),'" & str�����ļ� & "','" & str����ϵͳ & "','" & strPath & "','" & str�ļ�˵�� & "'," & bytǿ�Ƹ��� & " ," & byt�Զ�ע�� & ",'" & strPathADD & "')"
        gcnOracle.Execute strSQL
        SaveDate = True
        '������Ҫ������־
        Call SaveAuditLog(1, "�ļ���������-����", "�ɹ��������Ϊ��" & strName & "���ĵ���������")
        Exit Function
    End If
    
    Exit Function
errH:
End Function

Private Function GetUnionSysNum(ByVal str����ϵͳ As String, ByVal strCurSelectSys As String) As String
    On Error GoTo errH
    Dim strArr As Variant
    Dim i As Integer
    
    Dim strTemp As String
    strTemp = ""
    strArr = Split(strCurSelectSys, ",")
    For i = 0 To UBound(strArr) - 1
        If strArr(i) <> "" Then
            If InStrRev(strCurSelectSys, "," & strArr(i) & ",", 1) = 0 Then
                If strCurSelectSys <> "," & strArr(i) & "," Then
                    strTemp = strTemp & "," & strArr(i)
                End If
            End If
        End If
    Next
    
    If strTemp <> "" Then
        strTemp = strTemp & ","
        GetUnionSysNum = strTemp
        
    Else
        GetUnionSysNum = strCurSelectSys
    End If
    Exit Function
errH:
    
End Function

'==============================================================================
'=���ܣ� ����ļ��Ƿ�����ڲ��������ñ��Ƿ���Ҫ�Զ�ע�ᡢ�����ļ��Ƿ����
'=strFilePath �ļ�·��
'==============================================================================
Private Function CheckFile(ByVal strFilePath As String) As String
    On Error GoTo errH
    Dim rsTemp As New ADODB.Recordset
    Dim clsPEFileCheck As New clsPEReader
    Dim objFile As New FileSystemObject
    Dim strSQL As String
    Dim strFileName As String
    Dim strWarning As String
    
    lblWarning(1).Caption = ""
    strFileName = UCase(objFile.GetFileName(strFilePath))
'    UCase (GetFileName(strFilePath, , True))
    
    Select Case m_str��ʽ
        Case "����"
            '�ļ��Ƿ��Ѿ�����
            strSQL = "select ����ϵͳ,�ļ�����,ҵ�񲿼�,��װ·��,�ļ�˵��,1 from zlFilesUpgrade where upper(�ļ���) = upper('" & strFileName & "')"
            Call OpenRecordset(rsTemp, strSQL, Me.Caption)
            If rsTemp.EOF = False Then
                strWarning = FW_����1
                mblnFileRepeat = True
            End If
            '�����ļ��д���
            strSQL = "select 1 from zltools.zlfilesexpired where �ļ��� = '" & strFileName & "'"
            Call OpenRecordset(rsTemp, strSQL, Me.Caption)
            If rsTemp.EOF = False Then
                If strWarning <> "" Then
                    strWarning = strWarning & FW_����2
                Else
                    strWarning = FW_����2
                End If
                mblnFileEsexpired = True
            End If
            mblnFileExist = True
        Case "�޸�"
            '�����ļ��д���
            strSQL = "select 1 from zltools.zlfilesexpired where �ļ��� = '" & strFileName & "'"
            Call OpenRecordset(rsTemp, strSQL, Me.Caption)
            If rsTemp.EOF = False Then
                strWarning = FW_����2
                mblnFileEsexpired = True
            End If
            '�����ļ�ȱʧ���
            If objFile.FileExists(strFilePath) = False Then
                If strWarning <> "" Then
                    strWarning = strWarning & FW_����3
                Else
                    strWarning = FW_����3
                End If
                mblnFileExist = False
            Else
                mblnFileExist = True
            End If
    End Select
    
    '�ļ��Ƿ���Ҫ�Զ�ע����
    If clsPEFileCheck.LoadPEFile(strFilePath) Then
        mblnCheckRegist = True
        If clsPEFileCheck.IsDLL Or clsPEFileCheck.IsActivexEXE Then
            mblnRegist = True
        Else
            mblnRegist = False
        End If
    Else
        mblnCheckRegist = False
'        If strWarning <> "" Then
'            strWarning = strWarning & FW_����4
'        Else
'            strWarning = FW_����4
'        End If
    End If
    
    If InStr(strWarning, FW_����1) > 0 Or InStr(strWarning, FW_����2) > 0 Then
        lblWarning(1).Caption = "���ļ����ܱ��棬���飡"
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If
    lblWarning(0).Caption = strWarning
    CheckFile = strWarning
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Function

'==============================================================================
'=���ܣ� �����ļ�������
'==============================================================================
Private Sub AnalyzeFile(ByVal strFile As String)
    On Error GoTo errH
    Dim i As Integer
    Dim strWinSystemPath As String
    Dim StrWinPath       As String
    Dim strHelp          As String
    Dim strApp           As String
    
    strFile = UCase(strFile)
    strWinSystemPath = UCase(GetWinSystemPath())
    StrWinPath = UCase(GetWinPath())
'    strMainPan = UCase(Left(strWinPath, 1))
    strHelp = UCase(StrWinPath & "\HELP")


    If InStrRev(strFile, strWinSystemPath, -1, vbTextCompare) > 0 Then
        cbo(0).ListIndex = 1
    ElseIf InStrRev(strFile, strHelp, -1, vbTextCompare) > 0 Then
        cbo(0).ListIndex = 2
    ElseIf InStrRev(strFile, "\APPSOFT\", -1, vbTextCompare) > 0 Then
        strApp = GetAppSoft(strFile)
        If strApp = "" Then
            cbo(0).ListIndex = 0
        Else
            cbo(0).Text = "[APPSOFT]\" & strApp
        End If
    End If
    Exit Sub
errH:

End Sub

'==============================================================================
'=���ܣ����ļ�·��ת��Ϊͨ��·��
'==============================================================================
Private Function TransUniversalPath(ByVal strFile As String) As String
    Dim strWinSystemPath As String
    Dim blnIs64Bits As Boolean
    Dim arrTemp() As String
    Dim strPath As String
    Dim strSystemPath As String 'ϵͳsystem32λ��
    Dim strOS As String 'ϵͳ��
    Dim strDrive As String '�̷�
    Dim strCompare As String
    On Error GoTo errH
    blnIs64Bits = Is64bit
    
    strSystemPath = gobjFSO.GetSpecialFolder(SystemFolder)
    
    If blnIs64Bits Then '64ϵͳ��32λ����Ӧ�÷���C:\windows\SysWOW64
        strSystemPath = gobjFSO.GetParentFolderName(strSystemPath) & "\SysWOW64"
    End If
    If gobjFSO.FileExists(strFile) Then
        strFile = UCase(gobjFSO.GetParentFolderName(strFile))
    End If
    strSystemPath = UCase(strSystemPath)
    strOS = Split(strSystemPath, "\")(0)
    strDrive = Split(strFile, ":")(0)
    
    If InStr(strFile & "\", UCase(mcllPath("K_[SYSTEM]") & "\")) > 0 Then
        strPath = Replace(strFile, UCase(mcllPath("K_[SYSTEM]")), "[SYSTEM]")
    ElseIf InStr(strFile & "\", UCase(mcllPath("K_[APPSOFT]") & "\")) > 0 Then
        If InStr(strFile & "\", UCase(mcllPath("K_[PUBLIC]")) & "\") > 0 Then
            strPath = Replace(strFile, UCase(mcllPath("K_[PUBLIC]")), "[PUBLIC]")
        Else
            strPath = Replace(strFile, UCase(mcllPath("K_[APPSOFT]")), "[APPSOFT]")
        End If
    ElseIf InStr(strFile & "\", UCase(mcllPath("K_[OS:]")) & "\") > 0 Then
        strPath = Replace(strFile, strOS, "[OS:]")
    Else
        strPath = strFile
    End If
    
'    If intMode = 0 Then
'        arrTemp = Split(strFile, "\")
'        strPath = Replace(strPath, "\" & arrTemp(UBound(arrTemp)), "")
'    End If
    TransUniversalPath = strPath
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Function

'==============================================================================
'=���ܣ� ��ʼ��VSFCom
'==============================================================================
Private Sub initvsfCom()
    On Error GoTo errH
    With vsfCom
        .Tag = ""
'        .Redraw = flexRDNone
        .Rows = 6
        .Clear
        .Cols = 2
        .Cell(flexcpText, 0, 0) = "���"
        .Cell(flexcpAlignment, 0, 0) = flexAlignCenterCenter
        .ColWidth(0) = 500
        .Cell(flexcpText, 0, 1) = "�ļ���"
        .Cell(flexcpAlignment, 0, 1) = flexAlignCenterCenter
        .ColWidth(1) = 5000
'        .Cell(flexcpText, 0, 2) = "�汾��"
'        .Cell(flexcpAlignment, 0, 2) = flexAlignCenterCenter
'        .ColWidth(2) = 1000
'        .Cell(flexcpText, 0, 3) = "�޸�����"
'        .Cell(flexcpAlignment, 0, 2) = flexAlignCenterCenter
'        .ColWidth(3) = 1000
        '�Զ�����
        .WordWrap = True
        '�ϲ���Ԫ��
        .MergeCells = 0
        .MergeCol(.ColIndex("�ļ�����")) = True
        .MergeCol(.ColIndex("�ļ���")) = True
        '���ص�Ԫ��
        '�и�����
        .RowHeightMin = 300
        '���������
        .ColWidthMax = 4000
        '�Զ���Ӧ�иߡ��п�
        .AutoSizeMode = flexAutoSizeRowHeight
'        .AutoSize .ColIndex("�ļ���")
'        .SelectionMode = flexSelectionListBox
        .AllowBigSelection = False
        .Redraw = flexRDBuffered
        
    End With
    Exit Sub
errH:
 
End Sub

'ˢ��vsfCom
Private Sub refvsfCom(ByVal strFiles As String)
    On Error GoTo errH
    Dim i As Long
    Dim iNum As String
    Dim strTemp() As String
    Call initvsfCom
    If strFiles = "" Then Exit Sub
    strTemp = Split(strFiles, ",")
    
    With vsfCom
        .Rows = UBound(strTemp) + 2
        For i = 0 To UBound(strTemp)
            .Cell(flexcpText, i + 1, 0) = i + 1
            .Cell(flexcpAlignment, i + 1, 0) = flexAlignLeftCenter
            .Cell(flexcpText, i + 1, 1) = strTemp(i)
            .Cell(flexcpAlignment, i + 1, 1) = flexAlignLeftCenter
        Next
        
'        '�Զ�����
'        .WordWrap = True
'        '�ϲ���Ԫ��
'        .MergeCells = 2
'        .MergeCol(.ColIndex("�ļ�����")) = True
'        .MergeCol(.ColIndex("�ļ���")) = True
'        '���ص�Ԫ��
'        .ColWidth(.ColIndex("����ID")) = 0
'        '�и�����
'        .RowHeightMin = 300
'        '���������
'        .ColWidthMax = 7000
'        '�Զ���Ӧ�иߡ��п�
'        .AutoSizeMode = flexAutoSizeRowHeight
'        .AutoSize .ColIndex("ҵ�񲿼�")
'        .SelectionMode = flexSelectionListBox
'        .AllowBigSelection = False
'        .Redraw = flexRDBuffered
    End With
    
    Exit Sub
errH:
  
End Sub

Private Sub AddFile()
    Dim strFiles As String
    On Error GoTo errH
    
        strFiles = getFiles
        With frmEditFile
            .intItemFile = strFiles
            .intStrFile = txtFilePath.Text
            .strType = "0,1,2,3,4"
            .Show vbModal
            
            Call refvsfCom(.intItemFile)
         
        End With
        Set frmEditFile = Nothing
        Exit Sub
errH:
 
End Sub

Private Sub lvwSys_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        ReleaseCapture
    End If
End Sub

Private Sub txtExplanation_GotFocus()
    txtExplanation.BackColor = &HC0FFC0
End Sub

Private Sub txtExplanation_LostFocus()
    txtExplanation.BackColor = &H80000005
End Sub

Private Function getFiles() As String
    On Error GoTo errH
    Dim strTemp As String
    Dim i As Long
    strTemp = ""
    For i = 1 To vsfCom.Rows - 1
        If strTemp = "" Then
            If vsfCom.Cell(flexcpText, i, 1) <> "" Then
                strTemp = vsfCom.Cell(flexcpText, i, 1) & ","
            End If
        Else
            If vsfCom.Cell(flexcpText, i, 1) <> "" Then
                strTemp = strTemp & vsfCom.Cell(flexcpText, i, 1) & ","
            End If
        End If
    Next
    
    If Len(strTemp) <> 0 Then
        If Right(strTemp, 1) = "," Then
            strTemp = Left(strTemp, Len(strTemp) - 1)
        End If
        getFiles = strTemp
    End If
    Exit Function
errH:

End Function

'==============================================================================
'=ɾ������
'==============================================================================
Private Sub StandardDel()
    On Error GoTo errH
    Dim lngRow As Long
    Dim strSelectFile As String
    Dim strFiles As String

    If m_strCurFileName = "" Then Exit Sub
    strFiles = getFiles
    If strFiles <> "" Then
        lngRow = vsfCom.FindRow(CStr(m_strCurFileName), , 1)
        

        strFiles = strFiles & ","
        strFiles = Replace(strFiles, m_strCurFileName & ",", "")
        If Right(strFiles, 1) = "," Then
            strFiles = Left(strFiles, Len(strFiles) - 1)
        End If
        Call refvsfCom(strFiles)
        
        If lngRow <> -1 Then
            If lngRow >= 2 Then
              vsfCom.Select lngRow - 1, 1
              vsfCom.ShowCell lngRow - 1, 1
            End If
        End If
    End If
   
    Exit Sub
errH:
   
End Sub

'==============================================================================
'=ɾ����������
'==============================================================================
Private Sub StandardAllDel()
    On Error GoTo errH
    Call initvsfCom
    Exit Sub
errH:
 
End Sub

Private Sub txtUniversalPathADD_Change()
    txtUniversalPathADD.Text = UCase(txtUniversalPathADD.Text)
    txtUniversalPathADD.SelStart = Len(txtUniversalPathADD)
End Sub

'Private Sub txtUniversalPathADD_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If txtUniversalPathADD.Locked = True And txtUniversalPathADD.Tag = "0" Then
'        MsgBox "ȡ������Ҫע�ᡱ��ѡ������޸ĸ��Ӱ�װ·����", vbInformation, gstrSysName
'        txtUniversalPathADD.Tag = "1"
'    End If
'End Sub

'==============================================================================
'=���ܣ� �������б仯ʱ���»�����Ϣ
'==============================================================================
Private Sub vsfcom_RowColChange()
    On Error GoTo errH
    Call vsfcom_SelChange
    Exit Sub
errH:
  
End Sub

'==============================================================================
'=���ܣ� ����ѡ�����б仯ʱ���»�����Ϣ
'==============================================================================
Private Sub vsfcom_SelChange()
    Dim lngID       As Long
    On Error GoTo errH
    
'    fgMain.WallPaper = imgBG_fg(1).Picture
    m_lngCurRow = vsfCom.Row
    If m_lngCurRow = 0 Then Exit Sub
    m_strCurFileName = IIf(Len(vsfCom.Cell(flexcpText, m_lngCurRow, 1)) = 0, "", vsfCom.Cell(flexcpText, m_lngCurRow, 1))   '��ȡID
    
    Exit Sub
errH:
   
End Sub

Private Function GetAppSoft(ByVal strFile As String) As String
    On Error Resume Next
    Dim i As Integer
    Dim j As Integer
    Dim strTemp As String
    i = InStrRev(strFile, "\APPSOFT\", -1)
    strTemp = Right(strFile, Len(strFile) - (i + 8))
    i = InStrRev(strTemp, "\", -1)
    If i > 0 Then
        GetAppSoft = Left(strTemp, i)
    Else
        GetAppSoft = ""
    End If
End Function

Private Function CheckUniversalADDPath(strPath As String) As Boolean
    Dim strWinSystemPath As String
    Dim blnIs64Bits As Boolean
    Dim arrTemp() As String
    Dim strSystemPath As String 'ϵͳsystem32λ��
    Dim strOS As String 'ϵͳ��
    Dim strDrive As String '�̷�
    Dim strCompare As String
    Dim i As Integer
    
    On Error GoTo errH
    blnIs64Bits = Is64bit
    
    strSystemPath = gobjFSO.GetSpecialFolder(SystemFolder)
    
    If blnIs64Bits Then '64ϵͳ��32λ����Ӧ�÷���C:\windows\SysWOW64
        strSystemPath = gobjFSO.GetParentFolderName(strSystemPath) & "\SysWOW64"
    End If

    strSystemPath = UCase(strSystemPath)
    strOS = UCase(Split(strSystemPath, "\")(0))
    
    strPath = Replace(strPath, "[*]", "_")
    strPath = UCase(strPath)
    '�滻�ؼ���
    If InStr(strPath, "[SYSTEM]") > 0 Then
        strPath = Replace(strPath, "[SYSTEM]", UCase(mcllPath("K_[SYSTEM]")))
    ElseIf InStr(strPath, "[APPSOFT]") > 0 Then
        strPath = Replace(strPath, "[APPSOFT]", UCase(mcllPath("K_[APPSOFT]")))
    ElseIf InStr(strPath, "[PUBLIC]") > 0 Then
        strPath = Replace(strPath, "[PUBLIC]", UCase(mcllPath("K_[PUBLIC]")))
    ElseIf InStr(strPath, "[OS:]") > 0 Then
        strPath = Replace(strPath, "[OS:]", strOS)
    End If
    arrTemp = Split(strPath, "\")
    If UBound(arrTemp) < 1 Then CheckUniversalADDPath = False: Exit Function
    If Len(arrTemp(0)) <> 2 Then CheckUniversalADDPath = False: Exit Function
    strDrive = Split(arrTemp(0), ":")(0) '�̷�����Ƿ�Ϸ�
    If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", strDrive) < 0 Then CheckUniversalADDPath = False: Exit Function
    For i = 1 To UBound(arrTemp) '�ļ���������Ƿ�Ϸ�
        If InStr(arrTemp(i), "\") > 0 Or InStr(arrTemp(i), "/") > 0 Or InStr(arrTemp(i), ":") > 0 Or _
           InStr(arrTemp(i), "*") > 0 Or InStr(arrTemp(i), "?") > 0 Or InStr(arrTemp(i), """") > 0 Or _
           InStr(arrTemp(i), "<") > 0 Or InStr(arrTemp(i), ">") > 0 Or InStr(arrTemp(i), "|") > 0 Or arrTemp(i) = "" Then
            CheckUniversalADDPath = False
            Exit Function
        End If
    Next
    CheckUniversalADDPath = True
    Exit Function
errH:
'    MsgBox err.Description, vbInformation, gstrSysName
    CheckUniversalADDPath = False
    If False Then
        Resume
    End If
End Function

Public Function OpenFile() As String
    Dim strFilter As String
    
    strFilter = "�����ļ�" & "|" & "*.*" & "|" & "DLL�ļ�" & "|" & "*.DLL" & "|" & "OCX�ļ�" & "|" & "*.OCX"
    On Error GoTo err
    Cdlg.FileName = ""
    Cdlg.DialogTitle = "ѡ���ļ�"
    Cdlg.MaxFileSize = 8192
    Cdlg.CancelError = True
    Cdlg.InitDir = m_strPathJY
    Cdlg.FileName = ""
    Cdlg.Filter = strFilter
    Cdlg.Flags = cdlOFNExplorer
    Cdlg.ShowOpen
    If Cdlg.FileName <> "" Then
        OpenFile = Cdlg.FileName
    End If
    Exit Function
err:
    If err.Number = cdlCancel Then
        err.Clear
        OpenFile = ""
    End If
End Function
