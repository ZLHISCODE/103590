VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmClientUpgradeManage 
   BackColor       =   &H80000005&
   Caption         =   "�ͻ�����������"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11715
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmClientUpgradeManage.frx":0000
   ScaleHeight     =   6405
   ScaleWidth      =   11715
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicBar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1650
      Left            =   0
      Picture         =   "frmClientUpgradeManage.frx":803A
      ScaleHeight     =   1650
      ScaleWidth      =   37500
      TabIndex        =   3
      Top             =   570
      Width           =   37500
      Begin VB.PictureBox picNowTag 
         AutoSize        =   -1  'True
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   780
         Picture         =   "frmClientUpgradeManage.frx":D1724
         ScaleHeight     =   180
         ScaleWidth      =   315
         TabIndex        =   7
         Top             =   1470
         Width           =   315
      End
      Begin VB.Image imgArrow 
         Height          =   1125
         Index           =   2
         Left            =   8250
         Picture         =   "frmClientUpgradeManage.frx":D1A66
         Top             =   195
         Width           =   1125
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ͻ����������"
         Height          =   180
         Index           =   3
         Left            =   9555
         TabIndex        =   8
         Top             =   1155
         Width           =   1260
      End
      Begin VB.Image imgBtn 
         Height          =   960
         Index           =   3
         Left            =   9705
         Picture         =   "frmClientUpgradeManage.frx":D5D76
         Top             =   210
         Width           =   960
      End
      Begin VB.Image imgArrow 
         Height          =   1125
         Index           =   1
         Left            =   5085
         Picture         =   "frmClientUpgradeManage.frx":DAB21
         Top             =   195
         Width           =   1125
      End
      Begin VB.Image imgArrow 
         Height          =   1125
         Index           =   0
         Left            =   1965
         Picture         =   "frmClientUpgradeManage.frx":DEE31
         Top             =   195
         Width           =   1125
      End
      Begin VB.Image imgBtn 
         Height          =   825
         Index           =   0
         Left            =   600
         Picture         =   "frmClientUpgradeManage.frx":E3141
         ToolTipText     =   "���з�������������"
         Top             =   240
         Width           =   825
      End
      Begin VB.Image imgBtn 
         Height          =   825
         Index           =   1
         Left            =   3645
         Picture         =   "frmClientUpgradeManage.frx":E559D
         ToolTipText     =   "���������ϴ�����"
         Top             =   240
         Width           =   825
      End
      Begin VB.Image imgBtn 
         Height          =   825
         Index           =   2
         Left            =   6720
         Picture         =   "frmClientUpgradeManage.frx":E79F9
         ToolTipText     =   "�Կͻ���������������"
         Top             =   240
         Width           =   825
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ļ���������"
         Height          =   180
         Index           =   1
         Left            =   3528
         TabIndex        =   6
         Top             =   1152
         Width           =   1080
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ͻ�����������"
         Height          =   180
         Index           =   2
         Left            =   6495
         TabIndex        =   5
         Top             =   1150
         Width           =   1260
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ļ�����������"
         Height          =   180
         Index           =   0
         Left            =   372
         TabIndex        =   4
         Top             =   1152
         Width           =   1260
      End
   End
   Begin VB.Frame fraCaption 
      BackColor       =   &H00FFFFFF&
      Height          =   120
      Left            =   -135
      TabIndex        =   1
      Top             =   1050
      Width           =   10305
   End
   Begin XtremeSuiteControls.TabControl tbcThis 
      Height          =   1380
      Left            =   -15
      TabIndex        =   2
      Top             =   1860
      Width           =   1275
      _Version        =   589884
      _ExtentX        =   2249
      _ExtentY        =   2434
      _StockProps     =   64
   End
   Begin ComctlLib.ImageList imgList 
      Left            =   10710
      Top             =   195
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClientUpgradeManage.frx":E9E55
            Key             =   "����������-����"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClientUpgradeManage.frx":EC2BF
            Key             =   "����������-����"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClientUpgradeManage.frx":EE729
            Key             =   "����������-����"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClientUpgradeManage.frx":F0B93
            Key             =   "�����ļ�����-����"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClientUpgradeManage.frx":F2FFD
            Key             =   "�����ļ�����-����"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClientUpgradeManage.frx":F5467
            Key             =   "�����ļ�����-����"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClientUpgradeManage.frx":F78D1
            Key             =   "�ͻ�����������-����"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClientUpgradeManage.frx":F9D3B
            Key             =   "�ͻ�����������-����"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClientUpgradeManage.frx":FC1A5
            Key             =   "�ͻ�����������-����"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClientUpgradeManage.frx":FE60F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClientUpgradeManage.frx":101661
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmClientUpgradeManage.frx":1046B3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblcaption 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�ͻ�����������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   285
      TabIndex        =   0
      Top             =   150
      Width           =   1470
   End
End
Attribute VB_Name = "frmClientUpgradeManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobjPage(4) As Object
Private mblnMove As Boolean '����ͼ�갴ť��ʾ״̬
Private mpeSelect As PageEnum '��ǰѡ�й���ģ�� 0-���������� 1-�����ļ����� 2-�ͻ�����������
Private mblnLoad As Boolean '�����ж�ֵ ture - ���ڼ���  false - �������
Private mintPage As Integer
Private mstrFunc As String '��¼ģ�鹦��Ȩ���ַ���

'ҳ������
Private Enum PageEnum
    PE_�ļ����������� = 0
    PE_�ļ��������� = 1
    PE_�ͻ����������� = 2
    PE_�ͻ��������ſ� = 3
End Enum

'��ť״̬
Private Enum ImageState
    IS_���� = 1
    IS_���� = 2
    IS_���� = 3
End Enum

Private Enum PageBack
    PB_�ļ����������� = 10
    PB_�ļ��������� = 11
    PB_�ͻ����������� = 12
End Enum

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = mobjPage(mpeSelect).SupportPrint
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Call mobjPage(mpeSelect).SubPrint(bytMode)
End Sub

Private Sub InitTbcthis()
    On Error GoTo errH:
    mblnLoad = True
    With tbcThis
        .RemoveAll
        .InsertItem PE_�ļ�����������, "����������", mobjPage(PE_�ļ�����������).hwnd, PE_�ļ����������� * 3 + IS_����
        .InsertItem PE_�ļ���������, "�����ļ�����", mobjPage(PE_�ļ���������).hwnd, PE_�ļ��������� * 3 + IS_����
        .InsertItem PE_�ͻ�����������, "�ͻ�������", mobjPage(PE_�ͻ�����������).hwnd, PE_�ͻ����������� * 3 + IS_����
        .InsertItem PE_�ͻ��������ſ�, "�ͻ��������ſ�", mobjPage(PE_�ͻ��������ſ�).hwnd, PE_�ͻ��������ſ� * 3 + IS_����
    End With
    mblnLoad = False
    Exit Sub
errH:
End Sub

Private Sub Form_Load()
    On Error GoTo errH:
    Set mobjPage(PE_�ļ�����������) = New frmClientUpgradeSeverConfigure
    Set mobjPage(PE_�ļ���������) = New frmClientUpgradeFileManage
    Set mobjPage(PE_�ͻ�����������) = New frmClientUpgradeConfigure
    Set mobjPage(PE_�ͻ��������ſ�) = New frmClientUpgradeProfile
    
    '��ȡ��ǰ�û�ӵ�еĹ���Ȩ��
    mstrFunc = GetProgFuncs("0307")
    
    If Not CheckAndAdjustMustTable("ZLUPGRADESERVER", , False, , False) Then
        MsgBox "�뽫���ݿ�������10.35.40�Ժ�İ汾��ʹ�øù��ܣ�", vbInformation, gstrSysName
        Me.Tag = "HIDE"
        Exit Sub
    End If
    
    mintPage = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Caption, "ѡ��ҳǩ", "0"))
    Call InitTbcthis
    Call imgBtn_Click(mintPage) 'Ĭ����ʾ����������ҳ��
    Exit Sub
errH:
    Me.Tag = "HIDE"
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    If Me.Tag = "HIDE" Then Me.Hide
    tbcThis.Top = PicBar.Top + PicBar.Height - 350
    tbcThis.Width = Me.ScaleWidth + 50
    tbcThis.Height = Me.ScaleHeight - tbcThis.Top + 10
    imgBtn(0).Top = PicBar.Height / 2 - imgBtn(0).Height / 2 - 180
    imgBtn(1).Top = PicBar.Height / 2 - imgBtn(1).Height / 2 - 180
    imgBtn(2).Top = PicBar.Height / 2 - imgBtn(2).Height / 2 - 180
    lblPic(0).Top = imgBtn(0).Top + imgBtn(0).Height + 100
    lblPic(0).Left = imgBtn.Item(0).Left + (imgBtn.Item(0).Width / 2) - (lblPic(0).Width / 2)
    lblPic(1).Top = lblPic(0).Top
    lblPic(1).Left = imgBtn.Item(1).Left + (imgBtn.Item(1).Width / 2) - (lblPic(1).Width / 2)
    lblPic(2).Top = lblPic(0).Top
    lblPic(2).Left = imgBtn.Item(2).Left + (imgBtn.Item(2).Width / 2) - (lblPic(2).Width / 2)
    picNowTag.Top = PicBar.Height - picNowTag.Height
    fraCaption.Width = Me.ScaleWidth + 200
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mobjPage(PE_�ļ�����������) Is Nothing Then Unload mobjPage(PE_�ļ�����������)
    If Not mobjPage(PE_�ļ���������) Is Nothing Then Unload mobjPage(PE_�ļ���������)
    If Not mobjPage(PE_�ͻ�����������) Is Nothing Then Unload mobjPage(PE_�ͻ�����������)
    If Not mobjPage(PE_�ͻ��������ſ�) Is Nothing Then Unload mobjPage(PE_�ͻ��������ſ�)
    Set mobjPage(PE_�ļ�����������) = Nothing
    Set mobjPage(PE_�ļ���������) = Nothing
    Set mobjPage(PE_�ͻ�����������) = Nothing
    Set mobjPage(PE_�ͻ��������ſ�) = Nothing
End Sub

Private Sub imgBtn_Click(Index As Integer)
    imgBtn.Item(mpeSelect).Picture = imgList.ListImages.Item(mpeSelect * 3 + IS_����).Picture
    lblPic.Item(mpeSelect).Font.Bold = False
    Select Case Index
        Case 0, 1, 2, 3
            imgBtn.Item(Index).Picture = imgList.ListImages.Item(Index * 3 + IS_����).Picture   'ͼ�갴ť״̬�л�
            picNowTag.Left = imgBtn.Item(Index).Left + (imgBtn.Item(Index).Width / 2) - (picNowTag.Width / 2)
            lblPic.Item(Index).Font.Bold = True
            tbcThis.Item(Index).Selected = True
            mpeSelect = Index
    End Select
End Sub

Private Sub imgBtn_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If mpeSelect <> Index Then
        imgBtn.Item(Index).Picture = imgList.ListImages.Item(Index * 3 + IS_����).Picture
    End If
End Sub

Private Sub imgBtn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If mblnMove = False And mpeSelect <> Index Then
        imgBtn.Item(Index).Picture = imgList.ListImages.Item(Index * 3 + IS_����).Picture
        mblnMove = True
    End If
End Sub

Private Sub lblPic_Click(Index As Integer)
    Call imgBtn_Click(Index)
End Sub

Private Sub PicBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    
    For i = 0 To 3
        If mpeSelect <> i Then
            imgBtn.Item(i).Picture = imgList.ListImages.Item(i * 3 + IS_����).Picture
            lblPic.Item(i).Font.Bold = False
        End If
    Next
    mblnMove = False
End Sub

Private Sub tbcThis_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mblnLoad And mintPage <> 0 Then Exit Sub
    Me.Refresh
    Call mobjPage(Item.Index).RefreshData
    Call mobjPage(Item.Index).SetMenu
    '�������ַ���Ϊ�գ����ʾӵ��ȫ��Ȩ�ޣ���������Ȩ�޿���
    If mstrFunc <> "" Then
        Call mobjPage(Item.Index).SetControlEnable(mstrFunc)
    End If
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Caption, "ѡ��ҳǩ", Item.Index
End Sub

