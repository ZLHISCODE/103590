VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm���ձ������� 
   Caption         =   "��Ƚ������"
   ClientHeight    =   7215
   ClientLeft      =   825
   ClientTop       =   2505
   ClientWidth     =   7320
   Icon            =   "frm���ձ�������.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ils32 
      Left            =   1470
      Top             =   2730
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ձ�������.frx":08CA
            Key             =   "Fix"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ձ�������.frx":0BE4
            Key             =   "FixD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ձ�������.frx":0EFE
            Key             =   "Common"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ձ�������.frx":1218
            Key             =   "CommonD"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4905
      Left            =   2130
      ScaleHeight     =   4905
      ScaleWidth      =   4125
      TabIndex        =   8
      Top             =   1860
      Width           =   4125
      Begin VB.PictureBox picSplitH 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   45
         Left            =   30
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   4275
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1755
         Width           =   4275
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh֧���޶� 
         Height          =   2430
         Left            =   90
         TabIndex        =   9
         Top             =   375
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   4286
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483630
         BackColorBkg    =   -2147483643
         GridColor       =   4210752
         GridColorFixed  =   4210752
         FocusRect       =   2
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh֧������ 
         Height          =   1410
         Left            =   120
         TabIndex        =   10
         Top             =   3300
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   2487
         _Version        =   393216
         Rows            =   3
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483630
         BackColorBkg    =   -2147483643
         GridColor       =   4210752
         GridColorFixed  =   4210752
         FocusRect       =   2
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lbl֧���޶� 
         AutoSize        =   -1  'True
         Caption         =   "������ⶥ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   90
         Width           =   1260
      End
      Begin VB.Label lbl֧������ 
         AutoSize        =   -1  'True
         Caption         =   "ͳ��֧������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   90
         TabIndex        =   12
         Top             =   3000
         Width           =   1080
      End
   End
   Begin VB.ComboBox cmb���� 
      Height          =   300
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   900
      Width           =   2385
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   1270
      BandCount       =   1
      _CBWidth        =   7320
      _CBHeight       =   720
      _Version        =   "6.7.8988"
      Child1          =   "tbrThis"
      MinHeight1      =   660
      Width1          =   615
      FixedBackground1=   0   'False
      Key1            =   "Comm"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   660
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   7200
         _ExtentX        =   12700
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
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Object.ToolTipText     =   "��ӡԤ��"
               Object.Tag             =   "Ԥ��"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "New"
               Object.ToolTipText     =   "��������ȹ���"
               Object.Tag             =   "����"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Delete"
               Object.ToolTipText     =   "ɾ��ĩ��ȹ���"
               Object.Tag             =   "ɾ��"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Object.ToolTipText     =   "��������"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   5205
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ձ�������.frx":1532
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ձ�������.frx":174C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ձ�������.frx":1966
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ձ�������.frx":1B80
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ձ�������.frx":1D9A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ձ�������.frx":1FB4
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ձ�������.frx":21CE
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   4485
      Top             =   390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ձ�������.frx":23E8
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ձ�������.frx":2602
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ձ�������.frx":281C
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ձ�������.frx":2A36
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ձ�������.frx":2C50
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ձ�������.frx":2E6A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ձ�������.frx":3084
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwKind_S 
      Height          =   5325
      Left            =   105
      TabIndex        =   0
      Top             =   870
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   9393
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils32"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   6855
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   635
      SimpleText      =   $"frm���ձ�������.frx":329E
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm���ձ�������.frx":32E5
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7832
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
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
   Begin VB.PictureBox picSplitV 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6540
      Left            =   1890
      MousePointer    =   9  'Size W E
      ScaleHeight     =   6540
      ScaleWidth      =   45
      TabIndex        =   1
      Top             =   780
      Width           =   45
   End
   Begin MSComctlLib.TabStrip tab��� 
      Height          =   5400
      Left            =   2055
      TabIndex        =   5
      Top             =   1395
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   9525
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "2002��"
            ImageVarType    =   2
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "ҽ������(&N)"
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   2010
      TabIndex        =   7
      Top             =   930
      Width           =   990
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnusplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "��������ȹ���(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "ɾ��ĩ��ȹ���(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSplit0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditLimit 
         Caption         =   "�ⶥ������(&L)"
      End
      Begin VB.Menu mnuEditProportion 
         Caption         =   "ͳ��֧������(&P)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolspilt1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
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
         Caption         =   "Web�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelpSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frm���ձ�������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim msngStartX As Single, msngStartY As Single    '�ƶ�ǰ����λ��
Dim mstrKey As String
Dim mlng���� As Long
Dim mlng��ǰ�� As Long
Dim mblnLoad As Boolean

Private Sub Form_Activate()
    If mblnLoad = True Then
        '��ʾ��ǰ��
        lvwKind_S.SelectedItem.EnsureVisible
        lvwKind_S_ItemClick lvwKind_S.SelectedItem
    End If
    mblnLoad = False
End Sub

Private Sub Form_Load()
    RestoreWinState Me, App.ProductName
    
    mblnLoad = True
    Call Ȩ�޿���
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    
    sngTop = IIf(cbrThis.Visible, cbrThis.Top + cbrThis.Height, 0)
    sngBottom = ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    
    lvwKind_S.Top = sngTop
    lvwKind_S.Height = IIf(sngBottom - lvwKind_S.Top > 0, sngBottom - lvwKind_S.Top, 0)
    lvwKind_S.Left = ScaleLeft
    
    picSplitV.Top = sngTop
    picSplitV.Height = IIf(sngBottom - picSplitV.Top > 0, sngBottom - picSplitV.Top, 0)
    picSplitV.Left = lvwKind_S.Left + lvwKind_S.Width
    
    With cmb����
        '���ÿؼ�����߾�����
        lbl����.Left = picSplitV.Left + picSplitV.Width
        .Left = lbl����.Left + lbl����.Width + 30
        .Width = IIf(ScaleWidth - cmb����.Left > 0, ScaleWidth - cmb����.Left, 0)
    
        tab���.Left = lbl����.Left
        tab���.Width = IIf(ScaleWidth - tab���.Left > 0, ScaleWidth - tab���.Left, 0)
    End With
    
    With tab���
        If cmb����.Visible = True Then
            cmb����.Top = sngTop
            lbl����.Top = sngTop + 60
            .Top = cmb����.Top + cmb����.Height + 90
        Else
            .Top = sngTop
        End If
        
        .Height = IIf(sngBottom - .Top > 0, sngBottom - .Top, 0)
        picContainer.Left = .ClientLeft
        picContainer.Width = .ClientWidth
        picContainer.Top = .ClientTop
        picContainer.Height = .ClientHeight
    End With
    Me.Refresh
End Sub

Private Sub picContainer_Resize()
    
    With lbl֧���޶�
        .Top = 60
        .Left = 60
    
        msh֧���޶�.Top = .Top + .Height + 30
        msh֧���޶�.Left = .Left
       
    End With
    With msh֧���޶�
        .Width = IIf(picContainer.ScaleWidth - 120 > 0, picContainer.ScaleWidth - 120, 0)
    
        picSplitH.Top = msh֧���޶�.Top + msh֧���޶�.Height + 90
        picSplitH.Left = .Left
        picSplitH.Width = .Width
        
        lbl֧������.Top = picSplitH.Top + picSplitH.Height
        lbl֧������.Left = .Left
        
        msh֧������.Top = lbl֧������.Top + lbl֧������.Height + 30
        msh֧������.Height = IIf(picContainer.ScaleHeight - msh֧������.Top > 0, picContainer.ScaleHeight - msh֧������.Top, 0)
        msh֧������.Left = .Left
        msh֧������.Width = .Width
    End With
    
End Sub

Private Sub msh֧���޶�_DblClick()
    If mnuEditLimit.Visible = True And mnuEditLimit.Enabled = True Then
        Call mnuEditLimit_Click
    End If
End Sub

Private Sub msh֧������_DblClick()
    If mnuEditProportion.Visible = True And mnuEditProportion.Enabled = True Then
        Call mnuEditProportion_Click
    End If
End Sub

Private Sub lvwKind_S_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If mstrKey = Item.Key Then Exit Sub
    
    Dim rsTemp As New ADODB.Recordset
    
    cmb����.Clear
    cmb����.Visible = (Item.Tag = "1")
    lbl����.Visible = cmb����.Visible
    Call Form_Resize
    
    On Error GoTo errHandle
    If cmb����.Visible = False Then
        '��ҽ��ֻ����һ������
        cmb����.AddItem "1." & Item.Text
    Else
        gstrSQL = "select ���,����,���� from ��������Ŀ¼ where ����=" & Mid(Item.Key, 2) & " order by ���"
        Call OpenRecordset(rsTemp, Me.Caption)
        
        Do Until rsTemp.EOF
            cmb����.AddItem rsTemp("����") & "." & rsTemp("����")
            cmb����.ItemData(cmb����.NewIndex) = rsTemp("���")
            rsTemp.MoveNext
        Loop
    End If
    
    If cmb����.ListCount > 0 Then
        '�������Click�¼�
        zlControl.CboSetIndex cmb����.hwnd, 0
    End If
    Call Fill���
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmb����_Click()
    If cmb����.ItemData(cmb����.ListIndex) = mlng���� Then Exit Sub
    
    Call Fill���
End Sub

Private Sub Fill���()
'���ܣ����ݵ�ǰ���������ĵõ������Ϣ
    Dim lng���� As Long
    Dim rsTemp As New ADODB.Recordset
    
    If lvwKind_S.SelectedItem Is Nothing Then
        mstrKey = ""
        lng���� = 0
    Else
        mstrKey = lvwKind_S.SelectedItem.Key
        lng���� = Mid(mstrKey, 2)
    End If
    If cmb����.ListIndex < 0 Then
        mlng���� = -1
    Else
        mlng���� = cmb����.ItemData(cmb����.ListIndex)
    End If
    
    
    '������ҽ������ʹ�ù�ҽ����
    gstrSQL = "select distinct ��� from ���ձ������� where ����=" & lng���� & " and ����=" & mlng����
    Call OpenRecordset(rsTemp, Me.Caption)
    
    tab���.Tabs.Clear
    If rsTemp.RecordCount = 0 Then
        tab���.Tabs.Add , "K0", "�������Ϣ"
    Else
        Do Until rsTemp.EOF
            tab���.Tabs.Add , "K" & rsTemp("���"), rsTemp("���") & "��"
            If rsTemp("���") = mlng��ǰ�� Then
                tab���.Tabs("K" & mlng��ǰ��).Selected = True
            End If
            rsTemp.MoveNext
        Loop
    End If
    If tab���.SelectedItem Is Nothing Then
        tab���.Tabs(1).Selected = True
    End If
    
    Call tab���_Click
    
End Sub

Private Sub tab���_Click()
'���ܣ���ʾ��ǰ��ȵĽ���������û�У�����ʾ�ձ�
    Dim lng���� As Long
    Dim lng��� As Long
    
    lng��� = Mid(tab���.SelectedItem.Key, 2)
    If lng��� = 0 Then
        Call InitTable
    Else
        lng���� = Mid(lvwKind_S.SelectedItem.Key, 2)
        
        Call Fill֧���޶�(lng����, lng���)
        Call Fill֧������(lng����, lng���)
    End If
    
    Call SetMenu
End Sub

Private Sub Fill֧���޶�(ByVal lng���� As Long, ByVal lng��� As Long)
'���ܣ���ʾ��ǰ��ȵ�֧���޶�
'���������ݽ������ࡢ��ȣ�������ȫ�ֱ����õ�
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long, lngCol As Long
    Dim int��Ⱥ As Integer
    
    '�õ���Ա����
    gstrSQL = "select ���,���� from ������Ⱥ where ����=" & lng���� & " order by ���"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    '��ʼ�����
    With msh֧���޶�
        If rsTemp.RecordCount <> 0 Then
            .Rows = 3
            .FixedRows = 2
            .Cols = rsTemp.RecordCount * 2 + 1
            .MergeCells = flexMergeRestrictRows
            
            lngRow = 0
            lngCol = 1
            
            .TextMatrix(0, 0) = "סԺ����"
            .TextMatrix(1, 0) = "סԺ����"
            
            Do While Not rsTemp.EOF
                .TextMatrix(lngRow, lngCol) = rsTemp!����
                .ColData(lngCol) = rsTemp!���
                lngCol = lngCol + 1
                .TextMatrix(lngRow, lngCol) = rsTemp!����
                .ColData(lngCol) = rsTemp!���
                lngCol = lngCol + 1
                rsTemp.MoveNext
            Loop
            
            lngRow = lngRow + 1
            lngCol = 1
            rsTemp.MoveFirst
            Do While Not rsTemp.EOF
                .TextMatrix(lngRow, lngCol) = "��Ժ"
                lngCol = lngCol + 1
                .TextMatrix(lngRow, lngCol) = "תԺ"
                lngCol = lngCol + 1
                rsTemp.MoveNext
            Loop
        End If
    End With
    
    gstrSQL = " Select ��Ⱥ,��Ժ,���,��� From ���ձ������� " & _
              " Where ����=" & lng���� & " And ����=" & mlng���� & " And ���=" & lng��� & _
              " And ����=2 And ��Ժ=1 And ��Ⱥ=1 Order by ���"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    With msh֧���޶�
        '�����������
        For lngRow = 2 To .Rows - 1
            For lngCol = 1 To .Cols - 1
                .TextMatrix(lngRow, lngCol) = ""
            Next
        Next
        If rsTemp.RecordCount = 0 Then
            .Rows = 3
            .TextMatrix(.Rows - 1, 0) = 1
        End If
        
        Do Until rsTemp.EOF
            If Val(rsTemp!���) <> 0 Then
                lngRow = 1 + rsTemp!���
            Else
                lngRow = .Rows - 1
            End If
            lngCol = 0
            If rsTemp("���") = "A" Then
                .TextMatrix(lngRow, 0) = "�ⶥ��"
                .RowData(lngRow) = -1
            Else
                .TextMatrix(lngRow, 0) = "��" & rsTemp("���") & "��סԺ"
                .RowData(lngRow) = rsTemp!���
            End If
            .Rows = .Rows + 1
            rsTemp.MoveNext
        Loop
        If .TextMatrix(.Rows - 1, 0) = "" Then .Rows = .Rows - 1
        
        gstrSQL = " Select ��Ⱥ,��Ժ,���,��� From ���ձ������� " & _
                  " Where ����=" & lng���� & " And ����=" & mlng���� & " And ���=" & lng��� & _
                  " And ����=2 " & _
                  " Order by ��Ⱥ,��Ժ,���"
        Call OpenRecordset(rsTemp, Me.Caption)
        Do Until rsTemp.EOF
            If Val(rsTemp!���) <> 0 Then
                lngRow = 1 + rsTemp!���
            Else
                lngRow = .Rows - 1
            End If
            lngCol = rsTemp!��Ⱥ * 2 - IIf(rsTemp!��Ժ = 1, 1, 0)
            .TextMatrix(lngRow, lngCol) = Format(rsTemp("���"), "########0;-########0; ; ")
            rsTemp.MoveNext
        Loop
        If .TextMatrix(.Rows - 1, 0) = "" And .Rows - 1 > 3 Then .Rows = .Rows - 1
    End With
    
    '���ö���
    With msh֧���޶�
        For lngCol = 0 To .Cols - 1
            .ColAlignmentFixed(lngCol) = 4
            .ColWidth(lngCol) = 1200
        Next
        
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .MergeCol(0) = True
    End With
End Sub

Private Sub Fill֧������(ByVal lng���� As Long, ByVal lng��� As Long)
'���ܣ���ʾ��ǰ��ȵ�֧������
'���������ݽ������ࡢ��ȣ�������ȫ�ֱ����õ�
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long, lngCol As Long, lngRows As Long
    Dim col��ʼ�� As New Collection      'ÿ����Ա���ʵ���ʼ��
    
    '�õ���Ա����
    gstrSQL = "select ���,���� from ������Ⱥ where ����=" & lng���� & " order by ���"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    '��ʼ�����
    With msh֧������
        If rsTemp.RecordCount <> 0 Then
            .Rows = 3
            .FixedRows = 2
            .Cols = rsTemp.RecordCount * 2 + 1
            
            lngRow = 0
            lngCol = 1
            Do While Not rsTemp.EOF
                .TextMatrix(lngRow, lngCol) = rsTemp!����
                .ColData(lngCol) = rsTemp!���
                lngCol = lngCol + 1
                .TextMatrix(lngRow, lngCol) = rsTemp!����
                .ColData(lngCol) = rsTemp!���
                lngCol = lngCol + 1
                rsTemp.MoveNext
            Loop
            
            lngRow = lngRow + 1
            lngCol = 1
            rsTemp.MoveFirst
            Do While Not rsTemp.EOF
                .TextMatrix(lngRow, lngCol) = "��Ժ"
                lngCol = lngCol + 1
                .TextMatrix(lngRow, lngCol) = "תԺ"
                lngCol = lngCol + 1
                rsTemp.MoveNext
            Loop
        End If
    End With
    
    '�ٵõ����õ�
    gstrSQL = "select ����,���� from ���շ��õ� where ����=" & lng���� & " and ����=" & mlng���� & " order by ����"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    With msh֧������
        .ColWidth(0) = 1200
        .ColAlignment(0) = 7
        .TextMatrix(0, 0) = "���õ�"
        .TextMatrix(1, 0) = "���õ�"
        
        If rsTemp.RecordCount = 0 Then
            .RowData(2) = 0
        Else
            .Rows = .Rows + rsTemp.RecordCount - 1
            lngRow = 2
            Do Until rsTemp.EOF
                .TextMatrix(lngRow, 0) = rsTemp("����")
                .RowData(lngRow) = rsTemp("����")
                
                lngRow = lngRow + 1
                rsTemp.MoveNext
            Loop
        End If
        
        .Col = 0
        .Row = 0
        .ColSel = .Cols - 1
        .RowSel = 0
        .FillStyle = flexFillRepeat
        .CellAlignment = 4
        .FillStyle = flexFillSingle
        .AllowBigSelection = False
        .Row = 2: .Col = 1
    End With
    
    '���õ�֧������
    gstrSQL = " Select ��Ⱥ,��Ժ,����,���� From ���ձ������� " & _
              " Where ���=" & lng��� & " And ����=" & lng���� & " And ����=" & mlng���� & " And ����=1 " & _
              " Order by ��Ⱥ,��Ժ,����"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    With msh֧������
        '�����������
        For lngRow = 2 To .Rows - 1
            For lngCol = 1 To .Cols - 1
                .TextMatrix(lngRow, lngCol) = ""
            Next
        Next
        
        Do Until rsTemp.EOF
            lngRow = 1 + rsTemp!���� + 1
            lngCol = rsTemp!��Ⱥ * 2 - IIf(rsTemp!��Ժ = 1, 1, 0)
            
            .TextMatrix(lngRow, lngCol) = Format(rsTemp("����"), "0.00")
            rsTemp.MoveNext
        Loop
    End With
    
    '���ö���
    With msh֧������
        For lngCol = 0 To .Cols - 1
            .ColAlignmentFixed(lngCol) = 4
            .ColWidth(lngCol) = 1200
        Next
        
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .MergeCol(0) = True
    End With
End Sub

Private Sub mnuEditAdd_Click()
    Dim lng���� As Long
    Dim lng��� As Long, lngĩ�� As Long
    Dim blnReturn As VbMsgBoxResult
    
    lng���� = Mid(lvwKind_S.SelectedItem.Key, 2)
    lng��� = Val(Mid(tab���.SelectedItem.Key, 2))
    If lng��� = 0 Then
        '��ȫ����
        blnReturn = MsgBox("���Ƿ�Ҫ����" & mlng��ǰ�� & "��ȵĽ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName)
        gstrSQL = "ZL_���ձ�������_NEW(" & lng���� & "," & mlng���� & "," & mlng��ǰ�� & ")"
    Else
        lngĩ�� = Val(Mid(tab���.Tabs(tab���.Tabs.Count).Key, 2))
        blnReturn = MsgBox("���Ƿ�Ҫ��" & lngĩ�� & "��ȵĽ�������Ʋ�����Ϊ" & (lngĩ�� + 1) & "�ģ�", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName)
        gstrSQL = "ZL_���ձ�������_Copy(" & lng���� & "," & mlng���� & ")"
        
    End If
    
    If blnReturn = vbNo Then Exit Sub
    
    On Error GoTo ErrHand
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    'ˢ��
    Call Fill���
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditDelete_Click()
    Dim lng���� As Long
    Dim lng��� As Long
    
    
    lng���� = Mid(lvwKind_S.SelectedItem.Key, 2)
    lng��� = Val(Mid(tab���.SelectedItem.Key, 2))
    
    If MsgBox("�����Ҫɾ����" & lng��� & "��Ƚ��������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo ErrHand
    gcnOracle.BeginTrans
    gstrSQL = "ZL_���ձ�������_DELETE(1," & lng���� & "," & mlng���� & "," & lng��� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "ZL_���ձ�������_DELETE(2," & lng���� & "," & mlng���� & "," & lng��� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gcnOracle.CommitTrans
    
    Call Fill���
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub mnuEditLimit_Click()
    Dim lng���� As Long, lng��� As Long
    
    lng���� = Mid(lvwKind_S.SelectedItem.Key, 2)
    lng��� = Mid(tab���.SelectedItem.Key, 2)
    
    If frm����֧���޶�.�༭֧���޶�(lng����, mlng����, lng���) = True Then
        Call Fill֧���޶�(lng����, lng���)
    End If
End Sub

Private Sub mnuEditProportion_Click()
    Dim lng���� As Long, lng��� As Long
    
    lng���� = Mid(lvwKind_S.SelectedItem.Key, 2)
    lng��� = Mid(tab���.SelectedItem.Key, 2)
    
    If frm����֧������.�༭֧������(lng����, mlng����, lng���) = True Then
        Call Fill֧������(lng����, lng���)
    End If
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePreview_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnuFilePrintSet_Click()
    zlPrintSet
End Sub

Private Sub subPrint(ByVal bytMode As Byte)
    '����:���д�ӡ,Ԥ���������EXCEL
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    
    Dim objPrint As New zlPrintGrds
    Dim objRow As New zlTabAppRow
    
    Set objPrint.Grds = New Collection
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = "�����������"
        
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add lbl֧���޶�
    objPrint.UnderAppRows.Add objRow
    
'    Set objRow = New zlTabAppRow
'    objRow.Add lblTax
'    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
        
    objRow.Add "��ӡ��:" & UserInfo.����    '& "   ��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    objPrint.Grds.Add msh֧���޶�
    objPrint.Grds.Add msh֧������
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewGrds objPrint, 1
          Case 2
              zlPrintOrViewGrds objPrint, 2
          Case 3
              zlPrintOrViewGrds objPrint, 3
      End Select
    Else
        zlPrintOrViewGrds objPrint, bytMode
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
   ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuViewRefresh_Click()
'    Call RefList
End Sub

Private Sub mnuViewStatus_Click()
    Me.mnuViewStatus.Checked = Not Me.mnuViewStatus.Checked
    stbThis.Visible = Me.mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    Me.mnuViewToolButton.Checked = Not Me.mnuViewToolButton.Checked
    Me.mnuViewToolText.Enabled = Me.mnuViewToolButton.Checked
    Me.cbrThis.Visible = Me.mnuViewToolButton.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intCount As Integer, intRow As Integer, intCol As Integer
    
    Me.mnuViewToolText.Checked = Not Me.mnuViewToolText.Checked
    If Me.mnuViewToolText.Checked Then
        For intCount = 1 To Me.tbrThis.Buttons.Count
            Me.tbrThis.Buttons(intCount).Caption = Me.tbrThis.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To Me.tbrThis.Buttons.Count
            Me.tbrThis.Buttons(intCount).Caption = ""
        Next
    End If
    Me.cbrThis.Bands(1).MinHeight = Me.tbrThis.Height
    Me.cbrThis.Refresh
    Call Form_Resize
End Sub

Private Sub picSplitV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        msngStartX = x
    End If
End Sub

Private Sub picSplitV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sngTemp As Single
    
    If Button = 1 Then
        sngTemp = picSplitV.Left + x - msngStartX
        If sngTemp > 1000 And Me.ScaleWidth - (sngTemp + picSplitV.Width) > 1500 Then
            picSplitV.Left = sngTemp
            lvwKind_S.Width = picSplitV.Left - lvwKind_S.Left
            
            Call Form_Resize
        End If
        lvwKind_S.SetFocus
    End If
End Sub

Private Sub picSplitH_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        msngStartY = y
    End If
End Sub

Private Sub picSplitH_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sngTemp As Single
    
    If Button = 1 Then
        sngTemp = picSplitH.Top + y - msngStartY
        If sngTemp - msh֧���޶�.Top > 1500 And (msh֧������.Top + msh֧������.Height) - (sngTemp + picSplitV.Width) > 1500 Then
            picSplitH.Top = sngTemp
            msh֧���޶�.Height = picSplitH.Top - 90 - msh֧���޶�.Top
            
            Call picContainer_Resize
        End If
        msh֧���޶�.SetFocus
    End If
End Sub


Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Preview"
        mnuFilePreview_Click
    Case "Print"
        mnuFilePrint_Click
    Case "New"
        mnuEditAdd_Click
    Case "Delete"
        mnuEditDelete_Click
    Case "Help"
        mnuHelpTitle_Click
    Case "Exit"
        mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub Ȩ�޿���()
    If InStr(gstrPrivs, "��ɾ��") = 0 Then
        tbrThis.Buttons("New").Visible = False
        tbrThis.Buttons("Delete").Visible = False
        tbrThis.Buttons("Split1").Visible = False
        
        mnuEditAdd.Visible = False
        mnuEditDelete.Visible = False
        mnuEditSplit0.Visible = False
    End If
    
    If InStr(gstrPrivs, "�ⶥ������") = 0 Then
       mnuEditLimit.Visible = False
    End If
    
    If InStr(gstrPrivs, "ͳ��֧������") = 0 Then
        If mnuEditAdd.Visible = True Or mnuEditLimit.Visible = True Then
            mnuEditProportion.Visible = False
        Else
            mnuEditSplit0.Visible = True
            mnuEditProportion.Visible = False
            mnuEdit.Visible = False
        End If
    End If
End Sub

Private Sub SetMenu()
'���ܣ����ݵ�ǰ����ʾ�������ò˵�������
    Dim blnEnable As Boolean
    
    stbThis.Panels(2).Text = "��ǰѡ���ҽ������ǣ�" & lvwKind_S.SelectedItem.Text & " ���Ϊ��" & tab���.SelectedItem.Caption
    
    blnEnable = True 'Val(Mid(lvwKind_S.SelectedItem.Key, 2)) <> TYPE_�Թ��� And Val(Mid(lvwKind_S.SelectedItem.Key, 2)) <> TYPE_������
    
    mnuEditDelete.Enabled = (Mid(tab���.SelectedItem.Key, 2) > 0 And Mid(tab���.SelectedItem.Key, 2) <> mlng��ǰ��) And blnEnable
    If tab���.SelectedItem.Index < tab���.Tabs.Count Then
        'ֻ�ܴ����һ�꿪ʼɾ��
        mnuEditDelete.Enabled = False
    End If
    
    tbrThis.Buttons("Delete").Enabled = mnuEditDelete.Enabled
    
    mnuEditAdd.Enabled = cmb����.ListIndex >= 0 And blnEnable
    tbrThis.Buttons("New").Enabled = mnuEditAdd.Enabled
    'ֻҪ����һ������ȱʡ����֧���޶���
    mnuEditLimit.Enabled = (Mid(tab���.SelectedItem.Key, 2) > 0) And blnEnable
    '�����䵵����õ�
    mnuEditProportion.Enabled = (msh֧������.Rows - 1 >= 2) And blnEnable
End Sub

Private Sub InitTable()
    With msh֧���޶�
        .Cols = 2: .Rows = 2
        .Clear
        .ColAlignmentFixed(0) = 1
        .ColAlignment(1) = 7
        .ColWidth(0) = 1200
        .ColWidth(1) = 1200
        .TextMatrix(0, 0) = "סԺ����"
        .TextMatrix(0, 1) = "��Ⱥ"
        
        .Col = 0
        .Row = 0
        .ColSel = .Cols - 1
        .RowSel = 0
        .FillStyle = flexFillRepeat
        .CellAlignment = 4
        .FillStyle = flexFillSingle
        .AllowBigSelection = False
        .Row = 1: .Col = 1
    End With
    
    With msh֧������
        .Cols = 2: .Rows = 2
        .Clear
        .ColWidth(0) = 1200
        .ColWidth(1) = 1200
        .ColAlignmentFixed(0) = 1
        .ColAlignment(1) = 7
        
        .TextMatrix(0, 0) = "���õ�"
        .TextMatrix(0, 1) = "��Ⱥ"
        .RowData(1) = 0: .ColData(1) = 0
        
        .Col = 0
        .Row = 0
        .ColSel = .Cols - 1
        .RowSel = 0
        .FillStyle = flexFillRepeat
        .CellAlignment = 4
        .FillStyle = flexFillSingle
        .AllowBigSelection = False
        .Row = 1: .Col = 1
    End With
End Sub

Public Sub ShowForm(frmParent As Form)
'���ܣ�װ��ҽ�����
'˵����ʹ�ñ����ܵ���Ҫԭ�����ڳ����˳�ʱ���岻����
    Dim rsTemp As New ADODB.Recordset
    Dim strIcon As String, lst As ListItem
    
    gstrSQL = "select ���,����,�Ƿ�̶�,�������� from ������� where nvl(�Ƿ��ֹ,0)<>1  order by ���"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    If rsTemp.RecordCount = 0 Then
        '������ڴ����ʼ��ʱ���ã��Ͳ��ô�������������
        MsgBox "û�п��ñ�����𣬲���ʹ�ñ����ܡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If frm���ձ�������.Visible = True Then
        frm���ձ�������.Show
        Exit Sub
    End If
    
    '���ڲ��ܿ�ʼʹ�ÿؼ�
    Call InitTable
    mlng��ǰ�� = Format(zlDatabase.Currentdate, "yyyy")
    
    mstrKey = ""
    lvwKind_S.ListItems.Clear
    Do Until rsTemp.EOF
        strIcon = IIf(rsTemp("�Ƿ�̶�") = 1, "Fix", "Common")
        If rsTemp("���") = gintInsure Then strIcon = strIcon & "D"
        
        Set lst = lvwKind_S.ListItems.Add(, "K" & rsTemp("���"), rsTemp("����"), strIcon, strIcon)
        If rsTemp("���") = gintInsure Then
            lst.Selected = True
        End If
        
        lst.Tag = IIf(rsTemp("��������") = 1, 1, 0)
        
        rsTemp.MoveNext
    Loop
    If lvwKind_S.SelectedItem Is Nothing Then
        lvwKind_S.ListItems(1).Selected = True
    End If
    frm���ձ�������.Show , frmParent
End Sub


