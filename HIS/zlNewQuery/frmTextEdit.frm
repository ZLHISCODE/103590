VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTextEdit 
   Caption         =   "�ı��༭"
   ClientHeight    =   5895
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8880
   Icon            =   "frmTextEdit.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   Tag             =   "�ɱ仯��"
   Begin RichTextLib.RichTextBox rtb 
      Height          =   1350
      Left            =   3660
      TabIndex        =   3
      Top             =   2130
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   2381
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmTextEdit.frx":030A
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   855
      Top             =   3495
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ilsMenuHot 
      Left            =   6960
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextEdit.frx":03A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextEdit.frx":06C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextEdit.frx":09DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextEdit.frx":0BF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextEdit.frx":0E0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextEdit.frx":1029
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextEdit.frx":1243
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextEdit.frx":145D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextEdit.frx":1677
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextEdit.frx":1891
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextEdit.frx":1AB1
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   7560
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextEdit.frx":1CD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextEdit.frx":1FEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextEdit.frx":2305
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextEdit.frx":251F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextEdit.frx":2739
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextEdit.frx":2953
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextEdit.frx":2B6D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextEdit.frx":2D87
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextEdit.frx":2FA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextEdit.frx":31BB
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTextEdit.frx":33DB
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   8880
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   645
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   8760
         _ExtentX        =   15452
         _ExtentY        =   1138
         ButtonWidth     =   820
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsMenuHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ճ��"
               Key             =   "ճ��"
               Object.ToolTipText     =   "ճ��"
               Object.Tag             =   "ճ��"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "ɾ��"
               Object.ToolTipText     =   "ɾ��"
               Object.Tag             =   "ɾ��"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ɫ"
               Key             =   "��ɫ"
               Object.ToolTipText     =   "��ɫ"
               Object.Tag             =   "��ɫ"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   11
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   5535
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   635
      SimpleText      =   $"frmTextEdit.frx":35FB
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTextEdit.frx":3642
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10583
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
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileImport 
         Caption         =   "�����ı�(&I)"
      End
      Begin VB.Menu mnuFileOutPut 
         Caption         =   "�����ı�(&O)"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "����(&S)"
         Shortcut        =   ^S
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
      Begin VB.Menu mnuEditCopy 
         Caption         =   "����(&C)"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "ճ��(&P)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "ɾ��(&D)"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "����(&T)"
      End
      Begin VB.Menu mnuEditSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelAll 
         Caption         =   "ȫѡ(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFont 
         Caption         =   "����(&F)"
      End
      Begin VB.Menu mnuEditColor 
         Caption         =   "��ɫ(&L)"
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
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&Web�ϵ�����"
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
      Begin VB.Menu mnuHelpSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
   Begin VB.Menu mnuShort2 
      Caption         =   "��ݲ˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "����(&A)"
         Index           =   1
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "�޸�(&M)"
         Index           =   2
      End
      Begin VB.Menu mnuShortMenu2 
         Caption         =   "ɾ��(&D)"
         Index           =   3
      End
      Begin VB.Menu mnuShortsplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "��ͼ��(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "Сͼ��(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "�б�(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuShortIcon 
         Caption         =   "��ϸ����(&D)"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmTextEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'��ģ�������õ��ľֲ�����˵��
Private mblnFirst As Boolean                      '�Ƿ�Ϊ���ν��뱾ģ��(True:���ν���;False:���ǳ��ν���)
Private mOK As Boolean
Private mvarFile As New FileSystemObject
Private mvarText As TextStream

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    DoEvents
    
    '������ʾ������ݳ�ʼ������
    Call AdjustEanbled
    mnuFileSave.Enabled = False
    tbrThis.Buttons("����").Enabled = mnuFileSave.Enabled
End Sub

Private Sub Form_Load()
    '������ʾǰ�����ݳ�ʼ������
    mblnFirst = True
    RestoreWinState Me, App.ProductName
    
    mOK = False
    Call ModulePrivs
        
End Sub

Private Sub Form_Resize()
    '���ݴ���״̬,���������и��ؼ�����ʾλ��
    Dim sglCbrH As Single
    Dim sglStbH As Single
    
    On Error Resume Next
    sglCbrH = IIf(cbrThis.Visible, cbrThis.Height, 0)
    sglStbH = IIf(stbThis.Visible, stbThis.Height, 0)
            
    Call ResizeControl(rtb, 0, sglCbrH, Me.ScaleWidth, Me.ScaleHeight - sglCbrH - sglStbH)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rtb.Tag = "1" Then
        If MsgBox("�ı������Ѹ��ģ�ȷ�ϲ�������˳���", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    SaveWinState Me, App.ProductName
End Sub

Private Sub mnuEditColor_Click()
    On Error Resume Next
    dlg.CancelError = True
    dlg.flags = &H1
    rtb.SelStart = 0
    rtb.SelLength = Len(rtb.Text)
    dlg.Color = rtb.SelColor
    rtb.SelLength = 0
    dlg.ShowColor
    If Err.Number = 0 Then
        rtb.SelStart = 0
        rtb.SelLength = Len(rtb.Text)
        rtb.SelColor = dlg.Color
        rtb.SelLength = 0
        mnuFileSave.Enabled = True
        tbrThis.Buttons("����").Enabled = mnuFileSave.Enabled
    Else
        Err.Clear
    End If
End Sub

Private Sub mnuEditCopy_Click()
    If Len(rtb.SelText) <> 0 Then
        Clipboard.SetText rtb.SelText, vbCFRTF
        Call AdjustEanbled
    End If
End Sub

Private Sub mnuEditCut_Click()
    Call mnuEditCopy_Click
    rtb.SelText = ""
    Call AdjustEanbled
End Sub

Private Sub mnuEditDelete_Click()
    rtb.SelText = ""

End Sub

Private Sub mnuEditFont_Click()
    On Error Resume Next
    dlg.CancelError = True
    dlg.flags = &H3 Or &H100 Or &H400 Or &H200 Or &H10000
    
    dlg.FontName = rtb.Font.Name
    dlg.FontSize = rtb.Font.Size
    dlg.FontBold = rtb.Font.Bold
    dlg.FontItalic = rtb.Font.Italic
    rtb.SelStart = 0
    rtb.SelLength = Len(rtb.Text)
    dlg.Color = rtb.SelColor
    rtb.SelLength = 0
    
    dlg.ShowFont
    If Err.Number = 0 Then
        rtb.Font.Name = dlg.FontName
        rtb.Font.Size = dlg.FontSize
        rtb.Font.Bold = dlg.FontBold
        rtb.Font.Italic = dlg.FontItalic
        
        rtb.SelStart = 0
        rtb.SelLength = Len(rtb.Text)
        rtb.SelColor = dlg.Color
        rtb.SelLength = 0
        mnuFileSave.Enabled = True
        tbrThis.Buttons("����").Enabled = mnuFileSave.Enabled
    Else
        Err.Clear
    End If
End Sub

Private Sub mnuEditPaste_Click()
    rtb.SelText = Clipboard.GetText(vbCFRTF)
    Call AdjustEanbled
End Sub

Private Sub mnuEditSelAll_Click()
    rtb.SelStart = 0
    rtb.SelLength = Len(rtb.Text)
    Call AdjustEanbled
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileImport_Click()
    
    On Error Resume Next
    dlg.FileName = ""
    dlg.flags = &H4 Or &H200000
    dlg.Filter = "�ı��ļ�(*.txt)|*.TXT"
    dlg.DialogTitle = "��"
    dlg.FilterIndex = 0
    dlg.ShowOpen
    If dlg.FilterIndex > 0 Then
        Set mvarText = mvarFile.OpenTextFile(dlg.FileName)
        rtb.Text = mvarText.ReadAll
        mnuFileSave.Enabled = True
        tbrThis.Buttons("����").Enabled = mnuFileSave.Enabled
    End If
    
End Sub

Private Sub mnuFileOutPut_Click()
    
    On Error Resume Next
    dlg.FileName = ""
    dlg.flags = &H4 Or &H200000 Or &H2 Or &H800
    dlg.Filter = "�ı��ļ�(*.txt)|*.TXT"
    dlg.DefaultExt = ".txt"
    dlg.DialogTitle = "���Ϊ"
    dlg.FilterIndex = 0
    dlg.ShowSave
    If dlg.FilterIndex > 0 Then
        Set mvarText = mvarFile.CreateTextFile(dlg.FileName, True)
        mvarText.Write rtb.Text
        mvarText.Close
    End If
End Sub

Private Sub mnuFileSave_Click()
    Dim mTxt As TextBox
    
    Set mTxt = frmDefQueryItem.VisualTxt
    
    mTxt.Text = rtb.Text
    mTxt.FontName = rtb.Font.Name
    mTxt.FontSize = rtb.Font.Size
    mTxt.Font.Bold = rtb.Font.Bold
    mTxt.FontItalic = rtb.Font.Italic
    rtb.SelStart = 0
    rtb.SelLength = Len(rtb.Text)
    mTxt.ForeColor = IIf(IsNull(rtb.SelColor), 0, rtb.SelColor)
    rtb.SelLength = 0
    
    mnuFileSave.Enabled = False
    tbrThis.Buttons("����").Enabled = mnuFileSave.Enabled
    
    rtb.Tag = ""
    
    mOK = True
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTopic_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub


Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Long
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(i).Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call Form_Resize
    
End Sub

Private Sub rtb_Change()
    rtb.Tag = "1"
    mnuFileSave.Enabled = True
    tbrThis.Buttons("����").Enabled = mnuFileSave.Enabled
End Sub

Private Sub rtb_Click()
    Call AdjustEanbled
End Sub

Private Sub rtb_GotFocus()
    zlCommFun.OpenIme True
End Sub

Private Sub rtb_LostFocus()
    zlCommFun.OpenIme
End Sub

Private Sub rtb_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call AdjustEanbled
End Sub

Private Sub rtb_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call AdjustEanbled
    If Button = 2 Then Me.PopupMenu Me.mnuEdit, 2
End Sub

Private Sub rtb_SelChange()
    Call AdjustEanbled
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "����"
        Call mnuFileImport_Click
    Case "����"
        Call mnuFileOutPut_Click
    Case "����"
        Call mnuFileSave_Click
    Case "����"
        Call mnuEditCopy_Click
    Case "ճ��"
        Call mnuEditPaste_Click
    Case "ɾ��"
        Call mnuEditDelete_Click
    Case "����"
        Call mnuEditCut_Click
    Case "����"
        Call mnuEditFont_Click
    Case "��ɫ"
        Call mnuEditColor_Click
    Case "����"
        Call mnuHelpTopic_Click
    Case "�˳�"
        Call mnuFileExit_Click
    End Select
End Sub

'-----------------------------------------------------------------------------------------------------------------
'
'�������Զ��庯������̲���,������ģ����ʹ��
'
'-----------------------------------------------------------------------------------------------------------------
Private Sub ModulePrivs()
    '����ģ��Ȩ��,������������ػ���ʾ
    
    
End Sub

Public Function OpenTextEditDialog(frmMain As Form, objTxt As Object) As Boolean
    rtb.Text = objTxt.Text
    rtb.Font.Name = objTxt.FontName
    rtb.Font.Size = objTxt.FontSize
    rtb.Font.Bold = objTxt.FontBold
    rtb.Font.Italic = objTxt.FontItalic
    rtb.Tag = ""
    
    rtb.SelStart = 0
    rtb.SelLength = Len(rtb.Text)
    rtb.SelColor = objTxt.ForeColor
    rtb.SelLength = 0
    rtb.Tag = ""
    
    frmTextEdit.Show 1, frmMain
    OpenTextEditDialog = mOK
    
End Function

Private Sub AdjustEanbled()
    mnuEditPaste.Enabled = False
    mnuEditCopy.Enabled = False
    mnuEditCut.Enabled = False
    mnuEditDelete.Enabled = False
    
    If rtb.SelLength > 0 Then
        mnuEditCopy.Enabled = True
        mnuEditCut.Enabled = True
        mnuEditDelete.Enabled = True
    End If
    If Clipboard.GetText(vbCFRTF) <> "" Then
        mnuEditPaste.Enabled = True
    End If
    
    tbrThis.Buttons("����").Enabled = mnuEditCopy.Enabled
    tbrThis.Buttons("ճ��").Enabled = mnuEditPaste.Enabled
    tbrThis.Buttons("ɾ��").Enabled = mnuEditDelete.Enabled
    tbrThis.Buttons("����").Enabled = mnuEditCut.Enabled
    
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Me.PopupMenu Me.mnuViewTool, 2
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

