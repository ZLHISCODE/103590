VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMidMain 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "���ݿ��Ż�����"
   ClientHeight    =   10605
   ClientLeft      =   165
   ClientTop       =   495
   ClientWidth     =   16080
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '����ȱʡ
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imgNormal 
      Left            =   12960
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":08CA
            Key             =   "�Ự����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":11A4
            Key             =   "�ռ����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1A7E
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2358
            Key             =   "���ݿ�����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2C32
            Key             =   "SQL����"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":350C
            Key             =   "�Ự����_hot"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3DE6
            Key             =   "�ռ����_hot"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":46C0
            Key             =   "�������_hot"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":4F9A
            Key             =   "���ݿ�����_hot"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":5874
            Key             =   "SQL����_hot"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgHot 
      Left            =   12360
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":614E
            Key             =   "�Ự����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":6A28
            Key             =   "�ռ����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":7302
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":7BDC
            Key             =   "���ݿ�����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":84B6
            Key             =   "SQL����"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tblMenu 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16080
      _ExtentX        =   28363
      _ExtentY        =   1508
      ButtonWidth     =   1640
      ButtonHeight    =   1455
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgNormal"
      HotImageList    =   "imgHot"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "���ݿ�����"
            Key             =   "_0601"
            Object.Tag             =   "����˵�������ݿ�����״���Ĳ鿴�����ٻ�ȡAWE��ASH��ADDM���档"
            ImageKey        =   "���ݿ�����"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SQL����"
            Key             =   "_0602"
            Object.Tag             =   "����˵����������SQL�Ŀ���ɸ�飬ִ�мƻ��������Զ��Ż�������ͳ����Ϣ�鿴��SQL��ص�ִ����Ϣ�͹��������ѯ���Ż�����ز����鿴��"
            ImageKey        =   "SQL����"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�������"
            Key             =   "_0605"
            Object.Tag             =   "����˵��������ֶζ�Ӧ������ȱʧ�����飬�������������ɾ����"
            ImageKey        =   "�������"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�Ự����"
            Key             =   "_0604"
            Object.Tag             =   "����˵����������������Ĳ�ѯ���Ự��ɱ��"
            ImageKey        =   "�Ự����"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�ռ����"
            Key             =   "_0606"
            Object.Tag             =   "����˵������������������ļ��еķֲ�����鿴��������������������������������ļ�����ʱ�ļ���UNDO�ļ���������"
            ImageKey        =   "�ռ����"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.PictureBox pctTip 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   6360
         ScaleHeight     =   855
         ScaleWidth      =   12000
         TabIndex        =   1
         Top             =   0
         Width           =   12000
         Begin VB.Label lblTip 
            AutoSize        =   -1  'True
            Caption         =   "����˵�������ݿ�����״���Ĳ鿴�����ٻ�ȡAWR��ASH��ADDM���档"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   0
            TabIndex        =   2
            Top             =   600
            Width           =   9600
         End
      End
   End
End
Attribute VB_Name = "frmMidMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub MDIForm_Load()
    tblMenu.Buttons(1).Image = "���ݿ�����_hot"
End Sub

Private Sub MDIForm_resize()
     frmParent.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub tblMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim btnTmp As MSComctlLib.Button
    
    frmParent.ShowForm Mid(Button.Key, 2)
    lblTip.Caption = Button.Tag
    
    For Each btnTmp In tblMenu.Buttons
        btnTmp.Image = btnTmp.Caption
    Next
    
    Button.Image = Button.Caption & "_hot"
    
End Sub

Public Sub SetToolBarEnable(ByVal blnEnable As Boolean)
    tblMenu.Enabled = blnEnable
End Sub


