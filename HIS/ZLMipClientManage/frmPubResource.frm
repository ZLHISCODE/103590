VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPubResource 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin MSComctlLib.ImageList ils16 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmPubResource.frx":0000
            Key             =   "file"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":015A
            Key             =   "constitute"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubResource.frx":69BC
            Key             =   "folder"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPubResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function GetImageCtl(Optional ByVal bytMode As Byte = 1) As ImageList
    Set GetImageCtl = ils16
End Function
