VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmLogo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3330
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   6360
   Icon            =   "frmLogo.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmLogo.frx":5C12
   ScaleHeight     =   3330
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ImageList imgColor 
      Left            =   1260
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogo.frx":7FC7
            Key             =   "FILLCOLOR"
            Object.Tag             =   "562"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogo.frx":8133
            Key             =   "LINECOLOR"
            Object.Tag             =   "563"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogo.frx":828C
            Key             =   "FORECOLOR"
            Object.Tag             =   "564"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogo.frx":83D9
            Key             =   "FILLNONE"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogo.frx":8445
            Key             =   "FILLALL"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogo.frx":84AF
            Key             =   "FILLH"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogo.frx":8521
            Key             =   "FILLV"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogo.frx":8592
            Key             =   "FILLHV"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogo.frx":8605
            Key             =   "FILLL"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogo.frx":8680
            Key             =   "FILLR"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogo.frx":86F9
            Key             =   "FILLLR"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogo.frx":8772
            Key             =   "LINECONTINUE"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogo.frx":87CE
            Key             =   "LINEDOT"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogo.frx":882B
            Key             =   "LINEDASH"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogo.frx":8888
            Key             =   "LINEDASHDOT"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogo.frx":88E6
            Key             =   "LINEDASHDOT2"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogo.frx":8945
            Key             =   "LINEWIDTH1"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogo.frx":89A1
            Key             =   "LINEWIDTH2"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogo.frx":8A00
            Key             =   "LINEWIDTH3"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogo.frx":8A60
            Key             =   "LINEWIDTH4"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogo.frx":8AC2
            Key             =   "LINEWIDTH5"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.ImageManager ImageManager 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmLogo.frx":8B25
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "版权所有 Θ2005-2006 重庆中联信息产业有限公司.       保留所有权利."
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   1440
      TabIndex        =   0
      Top             =   2700
      Width           =   4830
   End
End
Attribute VB_Name = "frmLogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    '调入真彩色的程序图标
    Me.KeyPreview = True
End Sub
