VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmCasePrint 
   AutoRedraw      =   -1  'True
   Caption         =   "����Ԥ��"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9030
   Icon            =   "frmCasePrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6090
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   1270
      BandCount       =   1
      _CBWidth        =   9030
      _CBHeight       =   720
      _Version        =   "6.7.8988"
      Child1          =   "tbr"
      MinHeight1      =   660
      Width1          =   810
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   660
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   8910
         _ExtentX        =   15716
         _ExtentY        =   1164
         ButtonWidth     =   820
         ButtonHeight    =   1164
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   13
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "ѡ��"
               Key             =   "Set"
               Description     =   "ѡ��"
               Object.ToolTipText     =   "������ӡѡ��"
               Object.Tag             =   "ѡ��"
               ImageKey        =   "Set"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "����"
               Key             =   "Scale"
               Description     =   "����"
               Object.ToolTipText     =   "��ʾ����"
               Object.Tag             =   "����"
               ImageKey        =   "Scale"
               Style           =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Key             =   "Par_"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ǰ"
               Key             =   "First"
               Description     =   "��ǰ"
               Object.ToolTipText     =   "��ǰҳ(Home)"
               Object.Tag             =   "��ǰ"
               ImageKey        =   "First"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ҳ"
               Key             =   "Previous"
               Description     =   "��ҳ"
               Object.ToolTipText     =   "��һҳ(PageUp)"
               Object.Tag             =   "��ҳ"
               ImageKey        =   "Previous"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ҳ"
               Key             =   "Next"
               Description     =   "��ҳ"
               Object.ToolTipText     =   "��һҳ(PageDown)"
               Object.Tag             =   "��ҳ"
               ImageKey        =   "Next"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���"
               Key             =   "Last"
               Description     =   "���"
               Object.ToolTipText     =   "���ҳ(End)"
               Object.Tag             =   "���"
               ImageKey        =   "Last"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         Begin VB.ComboBox cboPage 
            Height          =   300
            Left            =   5595
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   195
            Width           =   1335
         End
         Begin VB.TextBox txtPage 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   5010
            Locked          =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            Tag             =   "��ǰҳ              ��"
            Text            =   "��ǰҳ"
            Top             =   270
            Width           =   3660
         End
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   5730
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCasePrint.frx":014A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10874
            Object.ToolTipText     =   "��ӡ����Ϣ"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4755
      Left            =   0
      ScaleHeight     =   4755
      ScaleWidth      =   8760
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   720
      Width           =   8760
      Begin VB.PictureBox picPage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3390
         Index           =   0
         Left            =   270
         MouseIcon       =   "frmCasePrint.frx":09DE
         MousePointer    =   99  'Custom
         ScaleHeight     =   3390
         ScaleWidth      =   6990
         TabIndex        =   6
         Top             =   180
         Width           =   6990
      End
      Begin VB.PictureBox picShadow 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3390
         Left            =   330
         ScaleHeight     =   3390
         ScaleWidth      =   6990
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   255
         Width           =   6990
      End
   End
   Begin VB.HScrollBar scrHsc 
      DragIcon        =   "frmCasePrint.frx":0B30
      Height          =   250
      LargeChange     =   20
      Left            =   0
      Max             =   100
      MouseIcon       =   "frmCasePrint.frx":0E3A
      SmallChange     =   10
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5475
      Width           =   8760
   End
   Begin VB.VScrollBar scrVsc 
      DragIcon        =   "frmCasePrint.frx":0F8C
      Height          =   4755
      LargeChange     =   20
      Left            =   8775
      Max             =   100
      MouseIcon       =   "frmCasePrint.frx":1296
      SmallChange     =   10
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   735
      Width           =   250
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   630
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCasePrint.frx":13E8
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCasePrint.frx":1602
            Key             =   "Set"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCasePrint.frx":1CFC
            Key             =   "Scale"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCasePrint.frx":1F16
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCasePrint.frx":2130
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCasePrint.frx":234A
            Key             =   "First"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCasePrint.frx":2564
            Key             =   "Previous"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCasePrint.frx":277E
            Key             =   "Next"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCasePrint.frx":2998
            Key             =   "Last"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCasePrint.frx":2BB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCasePrint.frx":2DCC
            Key             =   "Set"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCasePrint.frx":34C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCasePrint.frx":36E0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCasePrint.frx":38FA
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCasePrint.frx":3B14
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCasePrint.frx":3D2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCasePrint.frx":3F48
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCasePrint.frx":4162
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileSet 
         Caption         =   "ѡ��(&O)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Quit 
         Caption         =   "�˳�(&X)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "��ͼ(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "��׼��ť(&B)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuView_Tlb_1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&L)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuView_Scale 
         Caption         =   "��ʾ����(&C)"
         Enabled         =   0   'False
         Visible         =   0   'False
         Begin VB.Menu mnuView_ScaleValue 
            Caption         =   "ԭʼ��С(&O)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuView_ScaleValue 
            Caption         =   "����ҳ��(&W)"
            Index           =   1
         End
         Begin VB.Menu mnuView_ScaleValue 
            Caption         =   "����ҳ��(&H)"
            Index           =   2
         End
         Begin VB.Menu mnuView_ScaleValue 
            Caption         =   "��ҳ��ʾ(&P)"
            Index           =   3
         End
         Begin VB.Menu mnuView_ScaleValue 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuView_ScaleValue 
            Caption         =   "200%"
            Index           =   5
         End
         Begin VB.Menu mnuView_ScaleValue 
            Caption         =   "150%"
            Index           =   6
         End
         Begin VB.Menu mnuView_ScaleValue 
            Caption         =   "75%"
            Index           =   7
         End
         Begin VB.Menu mnuView_ScaleValue 
            Caption         =   "50%"
            Index           =   8
         End
      End
      Begin VB.Menu mnuView_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView_Move 
         Caption         =   "��ǰҳ(&F)"
         Index           =   0
      End
      Begin VB.Menu mnuView_Move 
         Caption         =   "ǰһҳ(&P)"
         Index           =   1
      End
      Begin VB.Menu mnuView_Move 
         Caption         =   "��һҳ(&N)"
         Index           =   2
      End
      Begin VB.Menu mnuView_Move 
         Caption         =   "���ҳ(&L)"
         Index           =   3
      End
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView_reFlash 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "WEB�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmCasePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintCurPage As Integer
Private mlngPreX As Long, mlngPreY As Long
Private Const Shadow_W = 45 '��Ӱ���

Private mblnCurCase As Boolean
Private mlngCurCase As Long
Private mblnPatiInfo As Boolean
Private mlngBeginY As Long
Private mintBeginPage As Integer
Private mobjParent As Object
Private mlng����id As Long
Private mvar��ҳ�򵥾� As Variant
Private mlng�������� As Long
Private mPrintBegingPage As Long, mPrintEndPage As Long

Public Function Preview(objParent As Object, ByVal lng�������� As Long, ByVal blnCurCase As Boolean, ByVal lngCurCase As Long, _
    ByVal lng����ID As Long, ByVal var��ҳ�򵥾� As Variant, _
    ByVal blnPatiInfo As Boolean, ByVal lngBeginY As Long, ByVal intBeginPage As Integer, Optional ByVal lng��ʼҳ As Long = 0, Optional ByVal lng����ҳ As Long = 0) As Boolean
    '���ܣ���ָ���Ĳ���(��)���д�ӡԤ��
    '������blnCurCase=�Ƿ�ֻ��ӡ��ǰ����
    '      lngCurCase=��ǰ������˳���
    '      lng����ID=����ID
    '      mvar��ҳ�򵥾�=�������浱ǰ����ҳID���ǵ��ݺ�,��ҳID�Ķ���סԺ����,���ݺŵĶ������ﲡ��
    '      blnPatiInfo=�Ƿ���ʾ������Ϣ
    '      lngBeginY=��ʼ���λ��
    '      intBeginPage=��ʼҳ��,0��ʾ����ӡҳ��
    
    mblnCurCase = blnCurCase
    mlngCurCase = lngCurCase
    
    mlng����id = lng����ID
    mvar��ҳ�򵥾� = var��ҳ�򵥾�
    mlng�������� = lng��������
    
    mblnPatiInfo = blnPatiInfo
    mlngBeginY = lngBeginY
    mintBeginPage = intBeginPage
    mPrintBegingPage = lng��ʼҳ
    mPrintEndPage = lng����ҳ
    
    Set mobjParent = objParent
    
    Me.Show 1, mobjParent
    
    Preview = True
End Function

Public Sub Init()
    '���ܣ���ʼ����ҳ��
    Dim ObjPic As PictureBox
    
    For Each ObjPic In picPage
        If ObjPic.Index = 0 Then
            ObjPic.Cls
        Else
            Unload ObjPic
        End If
    Next
    
End Sub

Private Sub cboPage_Click()
    Dim i As Integer
    
    For i = 0 To picPage.UBound
        If i = cboPage.ListIndex Then
            mintCurPage = i
            picPage(i).Visible = True
            picPage(i).ZOrder
        Else
            picPage(i).Visible = False
        End If
    Next
    Call Form_Resize
End Sub

Private Sub cbr_HeightChanged(ByVal NewHeight As Single)
    txtPage.Top = (NewHeight - txtPage.Height) / 2
    cboPage.Top = (NewHeight - cboPage.Height) / 2
End Sub

Private Sub Form_Activate()
    Call SetPages
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If ActiveControl Is cboPage Then Exit Sub
    Select Case KeyCode
    Case vbKeyUp
        If scrVsc.Enabled And scrVsc.Value > scrVsc.Min Then
            If Shift = 2 Then
                scrVsc.Value = IIf(scrVsc.Value - scrVsc.LargeChange < scrVsc.Min, scrVsc.Min, scrVsc.Value - scrVsc.LargeChange)
            Else
                scrVsc.Value = IIf(scrVsc.Value - scrVsc.SmallChange < scrVsc.Min, scrVsc.Min, scrVsc.Value - scrVsc.SmallChange)
            End If
        End If
    Case vbKeyDown
        If scrVsc.Enabled And scrVsc.Value < scrVsc.Max Then
            If Shift = 2 Then
                scrVsc.Value = IIf(scrVsc.Value + scrVsc.LargeChange > scrVsc.Max, scrVsc.Max, scrVsc.Value + scrVsc.LargeChange)
            Else
                scrVsc.Value = IIf(scrVsc.Value + scrVsc.SmallChange > scrVsc.Max, scrVsc.Max, scrVsc.Value + scrVsc.SmallChange)
            End If
        End If
    Case vbKeyLeft
        If scrHsc.Enabled And scrHsc.Value > scrHsc.Min Then
            If Shift = 2 Then
                scrHsc.Value = IIf(scrHsc.Value - scrHsc.LargeChange < scrHsc.Min, scrHsc.Min, scrHsc.Value - scrHsc.LargeChange)
            Else
                scrHsc.Value = IIf(scrHsc.Value - scrHsc.SmallChange < scrHsc.Min, scrHsc.Min, scrHsc.Value - scrHsc.SmallChange)
            End If
        End If
    Case vbKeyRight
        If scrHsc.Enabled And scrHsc.Value < scrHsc.Max Then
            If Shift = 2 Then
                scrHsc.Value = IIf(scrHsc.Value + scrHsc.LargeChange > scrHsc.Max, scrHsc.Max, scrHsc.Value + scrHsc.LargeChange)
            Else
                scrHsc.Value = IIf(scrHsc.Value + scrHsc.SmallChange > scrHsc.Max, scrHsc.Max, scrHsc.Value + scrHsc.SmallChange)
            End If
        End If
    Case vbKeyHome
        mnuView_Move_Click 0
    Case vbKeyEnd
        mnuView_Move_Click 3
    Case vbKeyPageUp
        mnuView_Move_Click 1
    Case vbKeyPageDown
        mnuView_Move_Click 2
    End Select
End Sub

Private Sub mnuFile_Print_Click()
    '���ܣ���ӡ����
    Dim intPage As Integer
    
    If MsgBox("׼����ӡ��������ӡ��׼��������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    intPage = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "ֽ��", Printer.PaperSize)
    If IsWindowsNT And intPage = 256 Then DelCustomPaper
    
    If Not InitPrint(mobjParent) Then
        MsgBox "��ӡ����ʼ��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    sta.Panels(2).Text = "�������ӡ�� " & Printer.DeviceName & " ���..."
    Call PrintOutCase(mobjParent, Printer, mlng��������, mblnCurCase, mlngCurCase, mlng����id, mvar��ҳ�򵥾�, mblnPatiInfo, mlngBeginY, mintBeginPage, mPrintBegingPage, mPrintEndPage)
    'WinNT�Զ���ֽ�Ŵ���
    If IsWindowsNT And intPage = 256 Then DelCustomPaper
    
    Call InitPrint(mobjParent)
    
    Call ShowPrinterInfo
End Sub



Private Sub mnuView_Move_Click(Index As Integer)
    With cboPage
        Select Case Index
        Case 0
            .ListIndex = 0
        Case 1
            If .ListIndex - 1 >= 0 Then .ListIndex = .ListIndex - 1
        Case 2
            If .ListIndex + 1 <= .ListCount - 1 Then .ListIndex = .ListIndex + 1
        Case 3
            .ListIndex = .ListCount - 1
        End Select
    End With
End Sub

Private Sub mnuView_reFlash_Click()
    Me.Refresh
End Sub

Private Sub picback_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mlngPreX = X: mlngPreY = Y
    If Button = 2 Then PopupMenu mnuView, 2
End Sub

Private Sub picback_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If scrVsc.Enabled Then
            If (Y - mlngPreY) / 15 > 0 Then
                scrVsc.Value = IIf(scrVsc.Value - (Y - mlngPreY) / 15 < scrVsc.Min, scrVsc.Min, scrVsc.Value - (Y - mlngPreY) / 15)
            Else
                scrVsc.Value = IIf(scrVsc.Value - (Y - mlngPreY) / 15 > scrVsc.Max, scrVsc.Max, scrVsc.Value - (Y - mlngPreY) / 15)
            End If
        End If
        If scrHsc.Enabled Then
            If (X - mlngPreX) / 15 > 0 Then
                scrHsc.Value = IIf(scrHsc.Value - (X - mlngPreX) / 15 < scrHsc.Min, scrHsc.Min, scrHsc.Value - (X - mlngPreX) / 15)
            Else
                scrHsc.Value = IIf(scrHsc.Value - (X - mlngPreX) / 15 > scrHsc.Max, scrHsc.Max, scrHsc.Value - (X - mlngPreX) / 15)
            End If
        End If
    End If
End Sub

Private Sub picPage_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    mlngPreX = X: mlngPreY = Y
    If Button = 1 Then Set picPage(Index).MouseIcon = scrHsc.MouseIcon
    If Button = 2 Then PopupMenu mnuView, 2
End Sub

Private Sub picPage_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If scrVsc.Enabled Then
            If (Y - mlngPreY) / 15 > 0 Then
                scrVsc.Value = IIf(scrVsc.Value - (Y - mlngPreY) / 15 < scrVsc.Min, scrVsc.Min, scrVsc.Value - (Y - mlngPreY) / 15)
            Else
                scrVsc.Value = IIf(scrVsc.Value - (Y - mlngPreY) / 15 > scrVsc.Max, scrVsc.Max, scrVsc.Value - (Y - mlngPreY) / 15)
            End If
        End If
        If scrHsc.Enabled Then
            If (X - mlngPreX) / 15 > 0 Then
                scrHsc.Value = IIf(scrHsc.Value - (X - mlngPreX) / 15 < scrHsc.Min, scrHsc.Min, scrHsc.Value - (X - mlngPreX) / 15)
            Else
                scrHsc.Value = IIf(scrHsc.Value - (X - mlngPreX) / 15 > scrHsc.Max, scrHsc.Max, scrHsc.Value - (X - mlngPreX) / 15)
            End If
        End If
    End If
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long '������ռ�ø߶�
    Dim staH As Long '״̬��ռ�ø߶�
    
    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    '����ؼ���Ⱥ͸߶�
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(sta.Visible, sta.Height, 0)
    
    picBack.Top = ScaleTop + cbrH
    picBack.Left = ScaleLeft
    picBack.Width = ScaleWidth - scrVsc.Width
    picBack.Height = ScaleHeight - staH - cbrH - scrHsc.Height
    
    scrVsc.Top = picBack.Top
    scrVsc.Left = ScaleWidth - scrVsc.Width
    scrVsc.Height = picBack.Height
    
    scrHsc.Left = picBack.Left
    scrHsc.Top = picBack.Top + picBack.Height
    scrHsc.Width = picBack.Width
    
    picShadow.Width = picPage(mintCurPage).Width
    picShadow.Height = picPage(mintCurPage).Height
    
    '����Ԥ��ҳ
    If picBack.ScaleWidth >= picPage(mintCurPage).Width + Shadow_W * 4 Then
        picPage(mintCurPage).Left = (picBack.ScaleWidth - (picPage(mintCurPage).Width + Shadow_W * 4)) / 2 + Shadow_W * 2
        picShadow.Left = picPage(mintCurPage).Left + Shadow_W
        scrHsc.Enabled = False
    Else
        scrHsc.Max = (picPage(mintCurPage).Width + Shadow_W * 4 - picBack.ScaleWidth) / 15
        If scrHsc.Max / 3 < scrHsc.SmallChange Then
            scrHsc.LargeChange = scrHsc.SmallChange
        Else
            scrHsc.LargeChange = scrHsc.Max / 3
        End If
        scrHsc.Value = 0
        scrHsc.Enabled = True
        scrhsc_Change
    End If
    If picBack.ScaleHeight >= picPage(mintCurPage).Height + Shadow_W * 4 Then
        picPage(mintCurPage).Top = (picBack.ScaleHeight - (picPage(mintCurPage).Height + Shadow_W * 4)) / 2 + Shadow_W
        picShadow.Top = picPage(mintCurPage).Top + Shadow_W
        scrVsc.Enabled = False
    Else
        scrVsc.Max = (picPage(mintCurPage).Height + Shadow_W * 4 - picBack.ScaleHeight) / 15
        If scrVsc.Max / 3 < scrVsc.SmallChange Then
            scrVsc.LargeChange = scrVsc.SmallChange
        Else
            scrVsc.LargeChange = scrVsc.Max / 3
        End If
        scrVsc.Value = 0
        scrVsc.Enabled = True
        scrVsc_Change
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub mnuFile_Quit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Sub�鿴�˵�(ByVal mnuLable As String)
    Dim i As Integer
    Select Case mnuLable
    Case "��׼��ť(&B)"
        mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
        mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
        cbr.Visible = Not cbr.Visible
        Form_Resize
    Case "�ı���ǩ(&L)"
        mnuViewToolText.Checked = Not mnuViewToolText.Checked
        For i = 1 To tbr.Buttons.Count
            If mnuViewToolText.Checked Then
                tbr.Buttons(i).Caption = tbr.Buttons(i).Tag
            Else
                tbr.Buttons(i).Caption = ""
            End If
        Next
        cbr.Bands(1).MINHEIGHT = tbr.ButtonHeight
        Form_Resize
    Case "״̬��(&S)"
        mnuViewStatus.Checked = Not mnuViewStatus.Checked
        sta.Visible = Not sta.Visible
        Form_Resize
    End Select
End Sub

Private Sub mnuViewStatus_Click()
    Sub�鿴�˵� mnuViewStatus.Caption
End Sub

Private Sub mnuViewToolButton_Click()
    Sub�鿴�˵� mnuViewToolButton.Caption
End Sub

Private Sub mnuViewToolText_Click()
    Sub�鿴�˵� mnuViewToolText.Caption
End Sub

Private Sub picPage_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Set picPage(Index).MouseIcon = scrVsc.MouseIcon
End Sub

Private Sub scrVsc_Change()
    picPage((mintCurPage)).Top = -scrVsc.Value * 15# + Shadow_W * 2
    picShadow.Top = picPage(mintCurPage).Top + Shadow_W
    Me.Refresh
End Sub

Private Sub scrVsc_Scroll()
    Call scrVsc_Change
End Sub

Private Sub scrhsc_Change()
    picPage(mintCurPage).Left = -scrHsc.Value * 15# + Shadow_W * 2
    picShadow.Left = picPage(mintCurPage).Left + Shadow_W
    Me.Refresh
End Sub

Private Sub scrhsc_Scroll()
    Call scrhsc_Change
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Quit"
        mnuFile_Quit_Click
    Case "Scale"
        tbr_ButtonDropDown Button
    Case "First"
        mnuView_Move_Click 0
    Case "Previous"
        mnuView_Move_Click 1
    Case "Next"
        mnuView_Move_Click 2
    Case "Last"
        mnuView_Move_Click 3
    Case "Print"
        mnuFile_Print_Click
    End Select
End Sub

Private Sub tbr_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Scale"
        PopupButtonMenu tbr, Button, mnuView_Scale
    End Select
End Sub

Private Sub txtPage_GotFocus()
    cboPage.SetFocus
End Sub

Private Sub Form_Load()
    mintCurPage = 0
    RestoreWinState Me, App.ProductName
    
    Call ShowPrinterInfo
End Sub

Private Function GetScale() As Single
    '���ܣ����ص�ǰ��ʾ����
    Dim i As Integer
    For i = 0 To mnuView_ScaleValue.UBound
        If mnuView_ScaleValue(i).Checked Then
            Select Case mnuView_ScaleValue(i).Index
            Case 0 'ԭʼ��С
                GetScale = 1
            Case 1 '����ҳ��
                GetScale = (picBack.ScaleWidth - Shadow_W * 4) / Printer.Width
            Case 2 '����ҳ��
                GetScale = (picBack.ScaleHeight - Shadow_W * 4) / Printer.Height
            Case 3 '��ҳ��ʾ
                If picBack.ScaleWidth / Printer.Width < picBack.ScaleHeight / Printer.Height Then
                    GetScale = (picBack.ScaleWidth - Shadow_W * 4) / Printer.Width
                Else
                    GetScale = (picBack.ScaleHeight - Shadow_W * 4) / Printer.Height
                End If
            Case Else
                GetScale = CDbl(Val(mnuView_ScaleValue(i).Caption) / 100)
            End Select
            Exit For
        End If
    Next
End Function

Private Function SetScaleMenu(strCaption As String)
    Dim i As Integer
    For i = 0 To mnuView_ScaleValue.UBound
        If mnuView_ScaleValue(i).Caption = strCaption Then
            mnuView_ScaleValue(i).Checked = True
        Else
            mnuView_ScaleValue(i).Checked = False
        End If
    Next
End Function

Private Sub mnuHelpWebHome_Click()
    zlHomePage hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hwnd
End Sub

Private Sub SetPages()
    Dim i As Integer
    
    txtPage.Text = "��ǰҳ" & Space(17) & "�� " & picPage.Count & " ҳ"
    
    cboPage.Clear
    For i = 0 To picPage.UBound
        cboPage.AddItem "�� " & i + 1 & " ҳ"
        picPage(i).Visible = False
    Next
    If mintCurPage <= cboPage.ListCount Then
        cboPage.ListIndex = mintCurPage
    Else
        cboPage.ListIndex = 0
    End If
End Sub

Private Sub ShowPrinterInfo()
    sta.Panels(2).Text = "��ӡ��:" & Printer.DeviceName & "/ֽ��:" & _
    GetPaperName(Printer.PaperSize) & "/�ߴ�:" & _
        CLng(Printer.Width / 56.7) & "��" & CLng(Printer.Height / 56.7) & "/ֽ��:" & _
        IIf(Printer.Orientation = 1, "����", "����")
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

