VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmToolBarWin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��Ӧ�̶�λ"
   ClientHeight    =   450
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   450
   ScaleWidth      =   2325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.ImageList ils2 
      Left            =   1560
      Top             =   570
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBarWin.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBarWin.frx":021A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBarWin.frx":0434
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBarWin.frx":064E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBarWin.frx":0868
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils1 
      Left            =   555
      Top             =   570
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBarWin.frx":0A82
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBarWin.frx":0C9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBarWin.frx":0EB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBarWin.frx":10D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolBarWin.frx":12EA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbTool 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   794
      ButtonWidth     =   820
      ButtonHeight    =   794
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ils1"
      HotImageList    =   "ils2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "First"
            Object.ToolTipText     =   "��һ��������������"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Previous"
            Object.ToolTipText     =   "��һ��������������"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Next"
            Object.ToolTipText     =   "��һ��������������"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Last"
            Object.ToolTipText     =   "���һ��������������"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "�˳���λ��ʽ"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmToolBarWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobjMainForm As Form

Public Sub ShowBar(winCaption As String, objParent As Form)
'--------------------------------------------------------------
'���ܣ����ö�λ���ߴ���
'������winCaption-------����Caption
'      objParent--------���ô���
'���أ�
'˵����
'--------------------------------------------------------------
    Me.Caption = winCaption
    Set mobjMainForm = objParent
    Me.Show
End Sub

Public Sub ����(Button As Integer, boolEnabled As Boolean)
'--------------------------------------------------------------
'���ܣ����ö�λ�������еİ�ť�Ƿ����
'������Button----------��ť���ͣ�0��ǰ�ư�ť��1�����ư�ť
'      boolEnabled-----��Ҫ���õ�Enabled����
'���أ�SQL���
'˵����
'--------------------------------------------------------------
    If Button = 0 Then
        tlbTool.Buttons("First").Enabled = boolEnabled
        tlbTool.Buttons("Previous").Enabled = boolEnabled
    End If
    If Button = 1 Then
        tlbTool.Buttons("Next").Enabled = boolEnabled
        tlbTool.Buttons("Last").Enabled = boolEnabled
    End If
End Sub

Private Sub Form_Load()
    Me.Height = 780
End Sub

Private Sub tlbTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo StopSub
    With mobjMainForm
        Select Case Button.Key
            Case "First"
                .subFirst
            Case "Previous"
                .subPrevious
            Case "Next"
                .subNext
            Case "Last"
                .subLast
            Case "Exit"
                Unload Me
        End Select
    End With
StopSub:
End Sub
