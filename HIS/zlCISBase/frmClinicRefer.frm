VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmCureREdit 
   BackColor       =   &H8000000C&
   Caption         =   "���Ʋο��༭"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8940
   Icon            =   "frmClinicRefer.frx":0000
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtInput 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   5970
      MultiLine       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5070
      Visible         =   0   'False
      Width           =   2445
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdRefer 
      Height          =   4935
      Left            =   390
      TabIndex        =   1
      Top             =   1125
      Width           =   7410
      _ExtentX        =   13070
      _ExtentY        =   8705
      _Version        =   393216
      BackColor       =   -2147483628
      Rows            =   10
      Cols            =   7
      FixedCols       =   3
      BackColorBkg    =   -2147483628
      GridColor       =   -2147483628
      GridColorFixed  =   16777215
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      MergeCells      =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
   End
   Begin ComCtl3.CoolBar clbThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   8940
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tlbThis"
      MinHeight1      =   720
      Width1          =   9705
      FixedBackground1=   0   'False
      Key1            =   "Comm"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tlbThis 
         Height          =   720
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   14
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Save"
               Description     =   "Ԥ��"
               Object.ToolTipText     =   "����ο�����"
               Object.Tag             =   "����"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�ָ�"
               Key             =   "Restore"
               Object.ToolTipText     =   "�ָ��ϴα���ʱ����"
               Object.Tag             =   "�ָ�"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Object.ToolTipText     =   "Ԥ���ο�����"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ�ο�����"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���"
               Key             =   "Insert"
               Object.ToolTipText     =   "�ڵ�ǰ���ݺ����һ��"
               Object.Tag             =   "���"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Delete"
               Object.ToolTipText     =   "ɾ��������"
               Object.Tag             =   "ɾ��"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��֢"
               Key             =   "Ban"
               Description     =   "��֢"
               Object.ToolTipText     =   "�޸ı��ζ�Ӧ����֢"
               Object.Tag             =   "��֢"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Find"
               Object.ToolTipText     =   "���ҵ���"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   10
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   7680
      Top             =   525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicRefer.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicRefer.frx":065C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicRefer.frx":0876
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicRefer.frx":0A90
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicRefer.frx":0CAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicRefer.frx":13A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicRefer.frx":1A9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicRefer.frx":2198
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicRefer.frx":23B2
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicRefer.frx":25D2
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   6915
      Top             =   525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicRefer.frx":27F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicRefer.frx":2A0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicRefer.frx":2C26
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicRefer.frx":2E40
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicRefer.frx":3060
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicRefer.frx":375A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicRefer.frx":3E54
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicRefer.frx":454E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicRefer.frx":4768
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicRefer.frx":4988
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   1530
      Top             =   6900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   8580
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmClinicRefer.frx":4BA8
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8123
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "���ߣ�ר������"
            TextSave        =   "���ߣ�ר������"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.Label lblScale 
      AutoSize        =   -1  'True
      Caption         =   "�����ߴ�"
      Height          =   180
      Left            =   7080
      TabIndex        =   4
      Top             =   6210
      Visible         =   0   'False
      Width           =   1185
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileSaveRefer 
         Caption         =   "����ο�(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileRestore 
         Caption         =   "�ָ�(&R)"
      End
      Begin VB.Menu mnuFileSaveTitle 
         Caption         =   "�������(&C)"
      End
      Begin VB.Menu mnuFileSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintset 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "Ԥ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileSpt2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditRowInsert 
         Caption         =   "��Ӷ���(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditRowDelete 
         Caption         =   "ɾ������(&D)"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuEditRowBan 
         Caption         =   "����֢(&B)..."
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuEditSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "����(&F)..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditReplace 
         Caption         =   "�滻(&R)..."
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuEditString 
         Caption         =   "�������(&S)..."
         Shortcut        =   ^T
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditSpt2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditTitleInsert 
         Caption         =   "��ӱ���(&I)..."
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu mnuEditTitleUpdate 
         Caption         =   "�޸ı���(&U)..."
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuEditTitleDelete 
         Caption         =   "ɾ������(&E)"
         Shortcut        =   +{DEL}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuToolBar 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolbarStand 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolbarText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStates 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFont 
         Caption         =   "����(&F)..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web�ϵ�����"
         WindowList      =   -1  'True
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)..."
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
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
Attribute VB_Name = "frmCureREdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mlngBarSize As Long

Dim rsTemp As New ADODB.Recordset
Dim rsMethod As New ADODB.Recordset
Dim strTemp As String
Dim intCount As Integer, lngRow As Integer, lngCol As Integer
Dim blnActive As Boolean

Const conRowHeight As Long = 255
Const conCol��Ŀ As Integer = 0
Const conCol���� As Integer = 1
Const conCol��� As Integer = 2
Const conCol���� As Integer = 3

Private Sub clbThis_Resize()
    Me.clbThis.Bands(1).MinHeight = Me.tlbThis.Height
    Me.clbThis.Refresh
    Call Form_Resize
End Sub

Private Sub Form_Activate()
    If blnActive Then Exit Sub
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select ID,����,����" & _
            " from ���Ʋο�Ŀ¼" & _
            " where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.hgdRefer.Tag))
    
    With rsTemp
        Me.Caption = !���� & "��Ӧ�òο�"
        Me.stbThis.Tag = IIf(IsNull(!����), "", !����)
        Me.stbThis.Panels(3).Text = "���ߣ�" & Me.stbThis.Tag
    End With
    If zlGetContent() = False Then MsgBox "������Ŀȱʡ�ο���Ŀ��ʧ(��ϵϵͳ����Ա)��", vbExclamation, gstrSysName: Unload Me: Exit Sub
    Call hgdRefer_RowColChange
    blnActive = True
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    blnActive = False
    Call RestoreWinState(Me, App.ProductName)
    With Me.hgdRefer
        .ColAlignmentFixed(conCol���) = 3
        .ColAlignment(conCol���� + 0) = 0
        .ColAlignment(conCol���� + 1) = 0
        .ColAlignment(conCol���� + 2) = 0
        .RowHeight(0) = 0
        .ColWidth(conCol��Ŀ) = 0
        .ColWidth(conCol����) = 0
        .ColWidth(conCol���) = 240
    End With
End Sub

Private Sub Form_Resize()
    Dim lngTools As Single, lngStatus As Single
    
    If blnActive Then Me.hgdRefer.SetFocus
    If WindowState = 1 Then Exit Sub
    lngTools = IIf(Me.clbThis.Visible, Me.clbThis.Height, 0)
    lngStatus = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    On Error Resume Next
    With Me.hgdRefer
        .Redraw = False
        .Left = Me.ScaleLeft
        .Top = lngTools
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - lngStatus - .Top
        .ColWidth(conCol���� + 0) = Me.TextWidth("������") + 90
        .ColWidth(conCol���� + 1) = Me.TextWidth("������") + 90
        .ColWidth(conCol���� + 2) = .Width - .ColWidth(conCol���) - .ColWidth(conCol����) - .ColWidth(conCol���� + 1) - mlngBarSize - 75
        Call zlGrdRowHeight
        .Redraw = True
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub hgdRefer_DblClick()
    Call hgdRefer_KeyPress(vbKeySpace)
End Sub

Private Sub hgdRefer_KeyPress(KeyAscii As Integer)
    With Me.hgdRefer
        Select Case KeyAscii
        Case vbKeyReturn, vbKeyTab
            .Row = .Row + IIf(.Row = .Rows - 1, 0, 1): Call hgdRefer_RowColChange: Exit Sub
        Case vbKeyDelete
            If .RowData(.Row) = 0 Then Exit Sub
            If .Col - conCol���� < .RowData(.Row) Then Exit Sub
            If .RowData(.Row) = 1 Then
                .TextMatrix(.Row, conCol���� + 1) = " "
                .TextMatrix(.Row, conCol���� + 2) = " "
            Else '.RowData(.Row) = 2 Then
                .TextMatrix(.Row, conCol���� + 2) = " "
            End If
        Case Else
            If .RowData(.Row) = 0 Then Exit Sub
            If .Col - conCol���� < .RowData(.Row) Then Exit Sub
            Me.txtInput.Top = .Top + .CellTop
            Me.txtInput.Height = .CellHeight
            If .RowData(.Row) = 1 Then
                Me.txtInput.Left = .Left + .ColWidth(conCol���) + .ColWidth(conCol����) + 45
                Me.txtInput.Width = .ColWidth(conCol���� + 1) + .ColWidth(conCol���� + 2) - 15
            Else '.RowData(.Row) = 2 Then
                Me.txtInput.Left = .Left + .ColWidth(conCol���) + .ColWidth(conCol����) + .ColWidth(conCol���� + 1) + 45
                Me.txtInput.Width = .ColWidth(conCol���� + 2) - 15
            End If
            If KeyAscii < 0 _
                Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9") _
                Or KeyAscii >= Asc("a") And KeyAscii <= Asc("z") _
                Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then
                Me.txtInput.Text = Chr(KeyAscii)
                Me.txtInput.SelStart = Len(Me.txtInput.Text)
            Else
                Me.txtInput.Text = .Text
                Me.txtInput.SelStart = 0
                Me.txtInput.SelLength = 30000
            End If
            Me.txtInput.Visible = True
            Me.mnuEditString.Visible = True
            Me.txtInput.SetFocus
        End Select
    End With
End Sub

Private Sub hgdRefer_KeyUp(KeyCode As Integer, Shift As Integer)
    With Me.hgdRefer
        Select Case KeyCode
        Case vbKeyDelete
            If .RowData(.Row) = 0 Then Exit Sub
            If .Col - conCol���� < .RowData(.Row) Then Exit Sub
            If .RowData(.Row) = 1 Then
                .TextMatrix(.Row, conCol���� + 1) = " "
                .TextMatrix(.Row, conCol���� + 2) = " "
            Else '.RowData(.Row) = 2 Then
                .TextMatrix(.Row, conCol���� + 2) = " "
            End If
        Case Else
        End Select
    End With
End Sub

Private Sub hgdRefer_RowColChange()
    With Me.hgdRefer
        '���ݲ�������
        If .RowData(.Row) = 0 Then
            '��Ŀ�в���ֱ��ɾ��
            Me.mnuEditRowDelete.Enabled = False
            Me.tlbThis.Buttons("Delete").Enabled = False
        Else
            Me.mnuEditRowDelete.Enabled = True
            Me.tlbThis.Buttons("Delete").Enabled = True
        End If
        If .TextMatrix(.Row, conCol���) <> "" Then
            Me.mnuEditRowBan.Enabled = True
            Me.tlbThis.Buttons("Ban").Enabled = True
        Else
            Me.mnuEditRowBan.Enabled = False
            Me.tlbThis.Buttons("Ban").Enabled = False
        End If
       
        If .RowIsVisible(.Row) = False Then .TopRow = .Row
        If .RowData(.Row) = 0 Then Exit Sub
        If .RowData(.Row) = 1 Then
            .Col = conCol���� + 1
        Else
            .Col = conCol���� + 2
        End If
    End With
End Sub

Private Sub mnuEditFind_Click()
    With frmDiagRefFind
        Set .frmParent = Me
        .Tag = "����"
        .Show , Me
    End With
End Sub

Private Sub mnuEditReplace_Click()
    With frmDiagRefFind
        Set .frmParent = Me
        .Tag = "�滻"
        .Show , Me
    End With
End Sub

Private Sub mnuEditRowBan_Click()
    Dim strBans As String
    With Me.hgdRefer
        If Trim(.TextMatrix(.Row, conCol���)) = "" Then Exit Sub
        strBans = .TextMatrix(.Row, conCol����)
    End With
    With frmCureRBans
        .strBans = strBans
        .Show 1, Me
        strBans = .strBans
        Unload frmCureRBans
    End With
    With Me.hgdRefer
        .TextMatrix(.Row, conCol����) = strBans
        If strBans = "" Then
            .TextMatrix(.Row, conCol���) = "��"
        Else
            .TextMatrix(.Row, conCol���) = "��"
        End If
    End With
End Sub

Private Sub mnuEditRowDelete_Click()
    Me.hgdRefer.SetFocus
    With Me.hgdRefer
        .Redraw = False
        For lngRow = .Row To .Rows - 2
            For lngCol = 0 To .Cols - 1
                .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow + 1, lngCol)
            Next
            .MergeRow(lngRow) = .MergeRow(lngRow + 1)
            .RowHeight(lngRow) = .RowHeight(lngRow + 1)
            .RowData(lngRow) = .RowData(lngRow + 1)
        Next
        .RowData(.Rows - 1) = 0
        .Rows = .Rows - 1
        .Redraw = True
    End With
    Call hgdRefer_RowColChange
End Sub

Private Sub mnuEditRowInsert_Click()
    Me.hgdRefer.SetFocus
    With Me.hgdRefer
        .Redraw = False
        .Rows = .Rows + 1
        For lngRow = .Rows - 1 To .Row + 1 Step -1
            For lngCol = 0 To .Cols - 1
                .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow - 1, lngCol)
            Next
            .MergeRow(lngRow) = .MergeRow(lngRow - 1)
            .RowHeight(lngRow) = .RowHeight(lngRow - 1)
            .RowData(lngRow) = .RowData(lngRow - 1)
        Next
        .TextMatrix(.Row + 1, conCol��Ŀ) = .TextMatrix(.Row, conCol��Ŀ)
        If .TextMatrix(.Row, conCol��Ŀ) <> "" Then
            If Split(.TextMatrix(.Row, conCol��Ŀ), ",")(3) = 1 Then
                .TextMatrix(.Row + 1, conCol���) = "��"
            Else
                .TextMatrix(.Row + 1, conCol���) = ""
            End If
            .TextMatrix(.Row + 1, conCol���� + 0) = ""
            If Split(.TextMatrix(.Row, conCol��Ŀ), ",")(2) = 1 Then
                .TextMatrix(.Row + 1, conCol���� + 1) = " "
                .TextMatrix(.Row + 1, conCol���� + 2) = " "
                .MergeRow(.Row + 1) = True
                .RowData(.Row + 1) = 1
            Else
                .TextMatrix(.Row + 1, conCol���� + 1) = ""
                .TextMatrix(.Row + 1, conCol���� + 2) = " "
                .MergeRow(.Row + 1) = False
                .RowData(.Row + 1) = 2
            End If
        Else
            .TextMatrix(.Row + 1, conCol���) = ""
            .TextMatrix(.Row + 1, conCol���� + 0) = ""
            .TextMatrix(.Row + 1, conCol���� + 1) = ""
            .TextMatrix(.Row + 1, conCol���� + 2) = " "
            .MergeRow(.Row + 1) = False
            .RowData(.Row + 1) = 2
        End If
        .RowHeight(.Row + 1) = conRowHeight
        .Redraw = True
    End With
    Call hgdRefer_RowColChange
End Sub

Private Sub mnuEditString_Click()
    If Me.txtInput.Visible = False Then Exit Sub
    strTemp = ShowSpecChar(Me)
    With Me.txtInput
        intCount = .SelStart
        .Text = Left(.Text, .SelStart) & strTemp & Mid(.Text, .SelStart + .SelLength + 1)
        .SelStart = intCount + Len(strTemp)
        DoEvents
        .Visible = True
        .SetFocus
        Me.mnuEditString.Visible = True
    End With
End Sub

Private Sub mnuEditTitleDelete_Click()
    Dim strTitle As String   '��ǰ��Ŀ����
    Dim lngCurRow As Long
    
    Me.hgdRefer.SetFocus
    With Me.hgdRefer
        strTitle = .TextMatrix(.Row, conCol��Ŀ)
        'ɾ�����
        If strTitle = Mid(.TextMatrix(0, conCol��Ŀ), 2) Then
            MsgBox "Ҫ��ο����ٱ���һ������Σ�����ɾ��", vbExclamation, gstrSysName
            Exit Sub
        End If
        If MsgBox("���ɾ����" & Split(strTitle, ",")(1) & "���������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        'ɾ������
        For lngCurRow = .Rows - 1 To .FixedRows Step -1
            If .TextMatrix(lngCurRow, conCol��Ŀ) = strTitle Then
                .Row = lngCurRow
                Call mnuEditRowDelete_Click
            End If
        Next
        .TextMatrix(0, conCol��Ŀ) = Split(.TextMatrix(0, conCol��Ŀ), ";" & strTitle)(0) & Split(.TextMatrix(0, conCol��Ŀ), ";" & strTitle)(1)
    End With
End Sub

Private Sub mnuEditTitleInsert_Click()
    Dim strLefts As String   '�Ѿ����ڵ�ǰ��ı���
    Dim strRights As String  '�Ѿ����ڵĺ���ı���
    Dim strTitle As String   '��ǰ��Ŀ����
    
    Me.hgdRefer.SetFocus
    With Me.hgdRefer
        strTitle = .TextMatrix(.Row, conCol��Ŀ)
        strLefts = Split(.TextMatrix(0, conCol��Ŀ), ";" & strTitle)(0) & ";" & strTitle
        strRights = Split(.TextMatrix(0, conCol��Ŀ), ";" & strTitle)(1)
    End With
    
    '---------------------------------------------
    '���ñ������ô��壬��ñ���
    With frmCureRTitle
        .optTier(0).Enabled = True
        .optTier(1).Enabled = True
        .Tag = "0,"  '����Ŀ�����
        If Split(strTitle, ",")(2) = 1 Then
            .optTier(0).Value = True
            .optTier(1).Value = False
        Else
            .optTier(0).Value = False
            .optTier(1).Value = True
        End If
        .chkBan.Value = Split(strTitle, ",")(3)
        .strLefts = strLefts
        .strRights = strRights
        .Show 1, Me
        strTitle = .strTitle
        Unload frmDiagTitle
    End With
    'ȡ�����ӣ��˳�����
    If strTitle = "" Then Exit Sub
    
    '---------------------------------------------
    '�������Ӵ���
    Dim strFromItem As String       '�������ӵ���Ŀ
    With Me.hgdRefer
        strFromItem = .TextMatrix(.Row, conCol��Ŀ)
        .TextMatrix(0, conCol��Ŀ) = strLefts & ";" & strTitle & strRights
        For lngRow = .Rows - 1 To .FixedRows Step -1
            If .TextMatrix(lngRow, conCol��Ŀ) = strFromItem Then Exit For
        Next
        .Row = lngRow
        Call mnuEditRowInsert_Click
        .Row = .Row + 1
        .TextMatrix(.Row, conCol��Ŀ) = strTitle
        .TextMatrix(.Row, conCol����) = ""
        .TextMatrix(.Row, conCol���) = ""
        If Split(strTitle, ",")(2) = 1 Then
            .TextMatrix(.Row, conCol���� + 0) = "��" & Split(strTitle, ",")(1) & "��"
            .TextMatrix(.Row, conCol���� + 1) = "��" & Split(strTitle, ",")(1) & "��"
            .TextMatrix(.Row, conCol���� + 2) = "��" & Split(strTitle, ",")(1) & "��"
        Else
            .TextMatrix(.Row, conCol���� + 0) = ""
            .TextMatrix(.Row, conCol���� + 1) = "��" & Split(strTitle, ",")(1) & "��"
            .TextMatrix(.Row, conCol���� + 2) = "��" & Split(strTitle, ",")(1) & "��"
        End If
        .MergeRow(.Row) = True
        .RowData(.Row) = 0
    End With
End Sub

Private Sub mnuEditTitleUpdate_Click()
    Dim strLefts As String   '�Ѿ����ڵ�ǰ��ı���
    Dim strRights As String  '�Ѿ����ڵĺ���ı���
    Dim strTitle As String   '��ǰ��Ŀ����
    Dim aryRows() As String
    
    Me.hgdRefer.SetFocus
    With Me.hgdRefer
        strTitle = .TextMatrix(.Row, conCol��Ŀ)
        strLefts = Split(.TextMatrix(0, conCol��Ŀ), ";" & strTitle)(0)
        strRights = Split(.TextMatrix(0, conCol��Ŀ), ";" & strTitle)(1)
    End With
    
    '---------------------------------------------
    '���ñ������ô��壬��ñ���
    With frmCureRTitle
        .Tag = Val(Split(strTitle, ",")(0)) & ","
        .txtName.Text = Split(strTitle, ",")(1)
        If Split(strTitle, ",")(2) = 1 Then
            .optTier(0).Value = True
            .optTier(1).Value = False
        Else
            .optTier(0).Value = False
            .optTier(1).Value = True
        End If
        .optTier(0).Enabled = False
        .optTier(1).Enabled = False
        .chkBan.Value = Split(strTitle, ",")(3)
        .strLefts = strLefts
        .strRights = strRights
        .Show 1, Me
        strTitle = .strTitle
        Unload frmDiagTitle
    End With
    'ȡ���޸ģ��˳�����
    If strTitle = "" Then Exit Sub
    
    '---------------------------------------------
    '�����޸Ĵ���
    Dim strFromItem As String       '���޸ĵ���Ŀ
    With Me.hgdRefer
        strFromItem = .TextMatrix(.Row, conCol��Ŀ)
        .TextMatrix(0, conCol��Ŀ) = strLefts & ";" & strTitle & strRights
        For lngRow = .FixedRows To .Rows - 1
            If .TextMatrix(lngRow, conCol��Ŀ) = strFromItem Then
                .TextMatrix(lngRow, conCol��Ŀ) = strTitle
                If Split(strTitle, ",")(3) = 1 And .RowData(lngRow) <> 0 Then
                    .TextMatrix(lngRow, conCol���) = "��"
                Else
                    .TextMatrix(lngRow, conCol���) = ""
                    .TextMatrix(lngRow, conCol����) = ""
                End If
                If .RowData(lngRow) = 0 Then
                    If Split(strTitle, ",")(2) = 1 Then
                        .TextMatrix(lngRow, conCol���� + 0) = "��" & Split(strTitle, ",")(1) & "��"
                        .TextMatrix(lngRow, conCol���� + 1) = "��" & Split(strTitle, ",")(1) & "��"
                        .TextMatrix(lngRow, conCol���� + 2) = "��" & Split(strTitle, ",")(1) & "��"
                    Else
                        .TextMatrix(lngRow, conCol���� + 0) = ""
                        .TextMatrix(lngRow, conCol���� + 1) = "��" & Split(strTitle, ",")(1) & "��"
                        .TextMatrix(lngRow, conCol���� + 2) = "��" & Split(strTitle, ",")(1) & "��"
                    End If
                End If
            End If
        Next
    End With
End Sub

Private Sub mnuFileExcel_Click()
    Call zlRptPrint(3)
End Sub

Private Sub mnuFilePreview_Click()
    Call zlRptPrint(0)
End Sub

Private Sub mnuFilePrint_Click()
    Call zlRptPrint(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub mnuFileRestore_Click()
    If MsgBox("����ָ��ϴα�������ݣ��ղŽ����޸Ľ���Ч" & vbCrLf & "��Ļָ���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Call zlGetContent
    Call hgdRefer_RowColChange
End Sub

Private Sub mnuFileSaveRefer_Click()
    Dim intOdd As Integer, intShowChars As Integer
    Dim strUpItem As String, strUpProof As String, strContent As String
    
    Me.hgdRefer.SetFocus
    
    '��Ŀ��������
    Call zlGrdSortItems
    
    Err = 0: On Error GoTo ErrHand
    gcnOracle.BeginTrans
    With Me.hgdRefer
        gstrSql = "zl_���Ʋο�����_Delete(" & .Tag & ")"
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
        
        intShowChars = Int(Me.stbThis.Panels(2).Width / Me.TextWidth("��"))
        intOdd = 0: strUpItem = "-"
        For lngRow = .FixedRows To .Rows - 1
            Me.stbThis.Panels(2).Text = String(intShowChars * lngRow / .Rows, "��")
            If .TextMatrix(lngRow, conCol��Ŀ) <> strUpItem Then
                intOdd = 0
            End If
            gstrSql = "zl_���Ʋο�����_Insert(" & .Tag & "," & _
                    Split(.TextMatrix(lngRow, conCol��Ŀ), ",")(0) & "," & _
                    "'" & Split(.TextMatrix(lngRow, conCol��Ŀ), ",")(1) & "'," & _
                    Split(.TextMatrix(lngRow, conCol��Ŀ), ",")(2) & "," & _
                    Split(.TextMatrix(lngRow, conCol��Ŀ), ",")(3) & ","
            gstrSql = gstrSql & intOdd & ","
            
            If .RowData(lngRow) = 0 Then
                gstrSql = gstrSql & "null,"
            Else
                strContent = Trim(.TextMatrix(lngRow, conCol���� + 2))
                strContent = Replace(strContent, vbCrLf, "")
                strContent = Replace(strContent, vbCr, "")
                strContent = Replace(strContent, vbLf, "")
                strContent = Replace(strContent, "'", StrConv("'", vbWide))
                strContent = Replace(strContent, "&", StrConv("&", vbWide))
                gstrSql = gstrSql & "'" & strContent & "',"
            End If
            If .TextMatrix(lngRow, conCol����) = "" Then
                gstrSql = gstrSql & "null,'" & Trim(Me.stbThis.Tag) & "')"
            Else
                gstrSql = gstrSql & "'" & .TextMatrix(lngRow, conCol����) & "','" & Trim(Me.stbThis.Tag) & "')"
            End If
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            
            intOdd = intOdd + 1
            strUpItem = .TextMatrix(lngRow, conCol��Ŀ)
        Next
    End With
    
    gcnOracle.CommitTrans
    Me.stbThis.Panels(2).Text = ""
    Exit Sub

ErrHand:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
    Me.stbThis.Panels(2).Text = ""
End Sub

Private Sub mnuFileSaveTitle_Click()
    If MsgBox("��ı��汾�ı�����Ϊ����ο��ο�ȱʡ������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
    
    Me.hgdRefer.SetFocus
    Call zlGrdSortItems     '��Ŀ��������
    
    gstrSql = "zl_���Ʋο���Ŀ_Save(" & Val(Me.Tag) & ",'" & Mid(Me.hgdRefer.TextMatrix(0, conCol��Ŀ), 2) & "')"
    Err = 0: On Error GoTo ErrHand
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuhelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuViewFont_Click()
    Me.hgdRefer.SetFocus
    With comDlg
        .FontName = Me.Font.Name
        .FontSize = Me.Font.Size
        .FontBold = Me.Font.Bold
        .FontItalic = Me.Font.Italic
        .Flags = cdlCFANSIOnly _
            + cdlCFApply _
            + cdlCFPrinterFonts
        .ShowFont
        Me.Font.Name = .FontName
        Me.Font.Size = .FontSize
        Me.Font.Bold = .FontBold
        Me.Font.Italic = .FontItalic
    End With
    Set Me.txtInput.Font = Me.Font
    Set Me.hgdRefer.Font = Me.Font
    Call Form_Resize
End Sub

Private Sub mnuViewStates_Click()
    Me.mnuViewStates.Checked = Not Me.mnuViewStates.Checked
    Me.stbThis.Visible = Me.mnuViewStates.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolbarStand_Click()
    Me.mnuViewToolbarStand.Checked = Not Me.mnuViewToolbarStand.Checked
    Me.clbThis.Visible = Me.mnuViewToolbarStand.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolBarText_Click()
    Dim i As Integer
    Me.mnuViewToolbarText.Checked = Not Me.mnuViewToolbarText.Checked
    If Me.mnuViewToolbarText.Checked Then
        For i = 1 To Me.tlbThis.Buttons.Count
            Me.tlbThis.Buttons(i).Caption = Me.tlbThis.Buttons(i).Tag
        Next
    Else
        For i = 1 To Me.tlbThis.Buttons.Count
            Me.tlbThis.Buttons(i).Caption = ""
        Next
    End If
    Me.clbThis.Bands(1).MinHeight = Me.tlbThis.Height
    Me.clbThis.Refresh
    Form_Resize
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Index = 3 Then
        strTemp = InputBox("���ο������ı�������" & vbCrLf & "  (ͨ��Ӧѡ��Ȩ��ר�ҵ�������Ϊ�ο�)", "����", Me.stbThis.Tag, Me.Left + Me.Width / 2 - 2500, Me.Top + Me.Height / 2)
        If Trim(strTemp) <> "" Then
            Me.stbThis.Tag = Left(strTemp, 10)
            Panel.Text = "���ߣ�" & Left(strTemp, 10)
        End If
    End If
End Sub

Private Sub tlbThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Save"
        Call mnuFileSaveRefer_Click
    Case "Restore"
        Call mnuFileRestore_Click
    Case "Preview"
        Call mnuFilePreview_Click
    Case "Print"
        Call mnuFilePrint_Click
    Case "Insert"
        Call mnuEditRowInsert_Click
    Case "Delete"
        Call mnuEditRowDelete_Click
    Case "Ban"
        Call mnuEditRowBan_Click
    Case "Find"
        Call mnuEditFind_Click
    Case "Help"
        Call mnuHelpHelp_Click
    Case "Exit"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub txtInput_Change()
    Dim lngColWidth As Long
    With Me.hgdRefer
        .Redraw = False
        If .RowData(.Row) = 2 Then
            lngColWidth = .ColWidth(conCol���� + 2)
            If Trim(Me.txtInput.Text) = "" Then
                .TextMatrix(.Row, conCol���� + 2) = " "
            Else
                .TextMatrix(.Row, conCol���� + 2) = Me.txtInput.Text
            End If
        ElseIf .RowData(.Row) = 1 Then
            lngColWidth = .ColWidth(conCol���� + 1) + .ColWidth(conCol���� + 2)
            If Trim(Me.txtInput.Text) = "" Then
                .TextMatrix(.Row, conCol���� + 1) = " "
                .TextMatrix(.Row, conCol���� + 2) = " "
            Else
                .TextMatrix(.Row, conCol���� + 1) = Me.txtInput.Text
                .TextMatrix(.Row, conCol���� + 2) = Me.txtInput.Text
            End If
        End If
        Me.lblScale.Width = lngColWidth - 90
        Me.lblScale.Caption = .TextMatrix(.Row, conCol���� + 2)
        .RowHeight(.Row) = Me.lblScale.Height + 75
        Me.txtInput.Height = .RowHeight(.Row)
        .Redraw = True
    End With
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Me.hgdRefer.SetFocus
        Call zlCommFun.PressKey(vbKeyReturn)
    End If
End Sub

Private Sub txtInput_LostFocus()
    Me.txtInput.Visible = False
    Me.mnuEditString.Visible = False
End Sub

Private Function zlGetContent() As Boolean
    '---------------------------------------------
    '��ȡ�ο�����
    '---------------------------------------------
    Err = 0: On Error GoTo ErrHand
    '--------------------------------------------------------
    Me.hgdRefer.Redraw = False
    Me.hgdRefer.Clear
    Me.hgdRefer.Rows = Me.hgdRefer.FixedRows + 1
    
    '����Ѿ������вο����ݣ�����ȡ��ʾ��
    gstrSql = "select ��Ŀ���,�ο���Ŀ,��Ŀ���,nvl(��������,0) as ��������," & _
            "       nvl(�����к�,0) as �����к�,nvl(�����ı�,'') as �����ı�" & _
            " from ���Ʋο�����" & _
            " where �ο�Ŀ¼id=[1] " & _
            " order by ��Ŀ���,�����к�"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.hgdRefer.Tag))
    
    With rsTemp
        lngRow = 0
        Do While Not .EOF
            lngRow = lngRow + 1
            If lngRow > Me.hgdRefer.Rows - Me.hgdRefer.FixedRows Then Me.hgdRefer.Rows = Me.hgdRefer.FixedRows + lngRow
            
            Me.hgdRefer.TextMatrix(lngRow, conCol��Ŀ) = !��Ŀ��� & "," & !�ο���Ŀ & "," & !��Ŀ��� & "," & !��������
            If InStr(1, Me.hgdRefer.TextMatrix(0, conCol��Ŀ), ";" & Me.hgdRefer.TextMatrix(lngRow, conCol��Ŀ)) = 0 Then
                Me.hgdRefer.TextMatrix(0, conCol��Ŀ) = Me.hgdRefer.TextMatrix(0, conCol��Ŀ) & ";" & Me.hgdRefer.TextMatrix(lngRow, conCol��Ŀ)
                If !��Ŀ��� = 1 Then
                    Me.hgdRefer.TextMatrix(lngRow, conCol���� + 0) = "��" & !�ο���Ŀ & "��"
                    Me.hgdRefer.TextMatrix(lngRow, conCol���� + 1) = "��" & !�ο���Ŀ & "��"
                    Me.hgdRefer.TextMatrix(lngRow, conCol���� + 2) = "��" & !�ο���Ŀ & "��"
                Else
                    Me.hgdRefer.TextMatrix(lngRow, conCol���� + 0) = ""
                    Me.hgdRefer.TextMatrix(lngRow, conCol���� + 1) = "��" & !�ο���Ŀ & "��"
                    Me.hgdRefer.TextMatrix(lngRow, conCol���� + 2) = "��" & !�ο���Ŀ & "��"
                End If
                Me.hgdRefer.MergeRow(lngRow) = True
                Me.hgdRefer.RowData(lngRow) = 0
                If Trim(!�����ı�) <> "" Then
                    lngRow = lngRow + 1
                    If lngRow > Me.hgdRefer.Rows - Me.hgdRefer.FixedRows Then Me.hgdRefer.Rows = Me.hgdRefer.FixedRows + lngRow
                    Me.hgdRefer.TextMatrix(lngRow, conCol��Ŀ) = !��Ŀ��� & "," & !�ο���Ŀ & "," & !��Ŀ��� & "," & !��������
                    If !��Ŀ��� = 1 Then
                        Me.hgdRefer.TextMatrix(lngRow, conCol���� + 1) = IIf(IsNull(!�����ı�), " ", !�����ı�)
                        Me.hgdRefer.TextMatrix(lngRow, conCol���� + 2) = IIf(IsNull(!�����ı�), " ", !�����ı�)
                        Me.hgdRefer.MergeRow(lngRow) = True
                        Me.hgdRefer.RowData(lngRow) = 1
                    Else
                        Me.hgdRefer.TextMatrix(lngRow, conCol���� + 1) = ""
                        Me.hgdRefer.TextMatrix(lngRow, conCol���� + 2) = IIf(IsNull(!�����ı�), " ", !�����ı�)
                        Me.hgdRefer.MergeRow(lngRow) = False
                        Me.hgdRefer.RowData(lngRow) = 2
                    End If
                End If
            Else
                Me.hgdRefer.TextMatrix(lngRow, conCol��Ŀ) = !��Ŀ��� & "," & !�ο���Ŀ & "," & !��Ŀ��� & "," & !��������
                If !��Ŀ��� = 1 Then
                    Me.hgdRefer.TextMatrix(lngRow, conCol���� + 1) = IIf(IsNull(!�����ı�), " ", !�����ı�)
                    Me.hgdRefer.TextMatrix(lngRow, conCol���� + 2) = IIf(IsNull(!�����ı�), " ", !�����ı�)
                    Me.hgdRefer.MergeRow(lngRow) = True
                    Me.hgdRefer.RowData(lngRow) = 1
                Else
                    Me.hgdRefer.TextMatrix(lngRow, conCol���� + 1) = ""
                    Me.hgdRefer.TextMatrix(lngRow, conCol���� + 2) = IIf(IsNull(!�����ı�), " ", !�����ı�)
                    Me.hgdRefer.MergeRow(lngRow) = False
                    Me.hgdRefer.RowData(lngRow) = 2
                End If
            End If
            
            If !�������� = 1 And Me.hgdRefer.RowData(lngRow) <> 0 Then
                Me.hgdRefer.TextMatrix(lngRow, conCol���) = "��"
                gstrSql = "select ����֢ID,��������" & _
                        " from ���Ʋο�����" & _
                        " where �ο�Ŀ¼id=[1] " & _
                        "       and �ο���Ŀ='" & !�ο���Ŀ & "'" & _
                        "       and nvl(�����к�,0)=" & !�����к�
                strTemp = ""
                Set rsMethod = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.hgdRefer.Tag))
                    
                With rsMethod
                    Do While Not .EOF
                        strTemp = strTemp & "|" & !����֢ID & "^" & IIf(IsNull(!��������), 1, !��������)
                        .MoveNext
                    Loop
                End With
                If strTemp <> "" Then
                    Me.hgdRefer.TextMatrix(lngRow, conCol����) = Mid(strTemp, 2)
                    Me.hgdRefer.TextMatrix(lngRow, conCol���) = "��"
                End If
            Else
                Me.hgdRefer.TextMatrix(lngRow, conCol���) = ""
            End If
            .MoveNext
        Loop
    End With
    If rsTemp.RecordCount > 0 Then
        Call zlGrdRowHeight
        Me.hgdRefer.Redraw = True
        zlGetContent = True: Exit Function
    End If
    
    '���û�б༭���ο�������ȱʡ����Ŀ��֯�ο���ʽ��
    gstrSql = "select ���,����,nvl(���,1) as ���,nvl(����,0) as ����" & _
            " from ���Ʋο���Ŀ" & _
            " where ����=[1] " & _
            " order by ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.Tag))
    
    With rsTemp
        lngRow = 0
        Do While Not .EOF
            lngRow = lngRow + 1
            If lngRow > Me.hgdRefer.Rows - Me.hgdRefer.FixedRows Then Me.hgdRefer.Rows = Me.hgdRefer.FixedRows + lngRow
            
            Me.hgdRefer.TextMatrix(lngRow, conCol��Ŀ) = !��� & "," & !���� & "," & !��� & "," & !����
            Me.hgdRefer.TextMatrix(0, conCol��Ŀ) = Me.hgdRefer.TextMatrix(0, conCol��Ŀ) & ";" & Me.hgdRefer.TextMatrix(lngRow, conCol��Ŀ)
            If !��� = 1 Then
                Me.hgdRefer.TextMatrix(lngRow, conCol���� + 0) = "��" & !���� & "��"
                Me.hgdRefer.TextMatrix(lngRow, conCol���� + 1) = "��" & !���� & "��"
                Me.hgdRefer.TextMatrix(lngRow, conCol���� + 2) = "��" & !���� & "��"
            Else
                Me.hgdRefer.TextMatrix(lngRow, conCol���� + 0) = ""
                Me.hgdRefer.TextMatrix(lngRow, conCol���� + 1) = "��" & !���� & "��"
                Me.hgdRefer.TextMatrix(lngRow, conCol���� + 2) = "��" & !���� & "��"
            End If
            Me.hgdRefer.MergeRow(lngRow) = True
            Me.hgdRefer.RowData(lngRow) = 0
            .MoveNext
        Loop
    End With
    Call zlGrdRowHeight
    Me.hgdRefer.Redraw = True
    
    If rsTemp.RecordCount = 0 Then
        zlGetContent = False
    Else
        zlGetContent = True
    End If
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub zlGrdRowHeight()
    '---------------------------------------------
    '���ݵ������ݵ�������������и߶ȣ��Ա�֤���ݵ�������ʾ
    '---------------------------------------------
    Dim lngColWidth As Long, lngTxtWidth As Long, intAskRows As Integer
    With Me.hgdRefer
        For lngRow = .FixedRows To .Rows - 1
            Select Case .RowData(lngRow)
            Case 0
                lngColWidth = .ColWidth(conCol���� + 2)
                If .TextMatrix(lngRow, conCol���� + 2) = .TextMatrix(lngRow, conCol���� + 1) Then
                    lngColWidth = lngColWidth + .ColWidth(conCol���� + 1)
                    If .TextMatrix(lngRow, conCol���� + 1) = .TextMatrix(lngRow, conCol����) Then
                        lngColWidth = lngColWidth + .ColWidth(conCol����)
                    End If
                End If
            Case 1
                lngColWidth = .ColWidth(conCol���� + 1) + .ColWidth(conCol���� + 2)
            Case 2
                lngColWidth = .ColWidth(conCol���� + 2)
            End Select
            Me.lblScale.Width = lngColWidth - 90
            Me.lblScale.Caption = .TextMatrix(lngRow, conCol���� + 2)
            .RowHeight(lngRow) = Me.lblScale.Height + 75
        Next
    End With
End Sub

Private Sub zlGrdSortItems()
    '---------------------------------------------
    '�����������õı�����Ŀ���Ա㱣��
    '---------------------------------------------
    Dim aryRows() As String, aryFlds() As String, strNewRows As String
    
    aryRows = Split(Mid(Me.hgdRefer.TextMatrix(0, conCol��Ŀ), 2), ";")
    For intCount = LBound(aryRows) To UBound(aryRows)
        aryFlds = Split(aryRows(intCount), ",")
        aryFlds(0) = intCount + 1
        strNewRows = Join(aryFlds, ",")
        
        '�����µ���Ŀ�޸ı�����Ŀ��Ԫ������
        With Me.hgdRefer
            For lngRow = .FixedRows To .Rows - 1
                If .TextMatrix(lngRow, conCol��Ŀ) = aryRows(intCount) Then
                    .TextMatrix(lngRow, conCol��Ŀ) = strNewRows
                End If
            Next
        End With
        aryRows(intCount) = strNewRows
    Next
    Me.hgdRefer.TextMatrix(0, conCol��Ŀ) = ";" & Join(aryRows, ";")
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '����:��¼���ӡ
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '-------------------------------------------------
    Dim objPrint As New zlPrint1Grd
    
    Set objPrint.Body = Me.hgdRefer
    With objPrint.Title
        .Text = Me.Caption
        .Font.Size = Me.Font.Size + 2
    End With
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode = 0 Then Exit Sub
    End If
    Call zlPrintOrView1Grd(objPrint, bytMode)
End Sub

Public Function zlWordSelect(lngCurRow As Long, strWord As String) As Long
    '-------------------------------------------------
    '����:��ָ����������ѡ��ָ�����ı�
    '���:  lngCurRow-ָ���У�strWord-ָ��ѡ�е��ı�
    '����:  δ���ҵ�������0�����ҵ��򷵻ظ��ı�����һ��λ��
    '-------------------------------------------------
    Me.txtInput.Visible = False
    Me.hgdRefer.Row = lngCurRow
    Me.hgdRefer.Col = conCol���� + 2
    With Me.hgdRefer
        If .RowData(.Row) = 0 Then zlWordSelect = 0: Exit Function
        Me.txtInput.Top = .Top + .CellTop
        Me.txtInput.Height = .CellHeight
        If .RowData(.Row) = 1 Then
            Me.txtInput.Left = .Left + .ColWidth(conCol���) + .ColWidth(conCol����) + 45
            Me.txtInput.Width = .ColWidth(conCol���� + 1) + .ColWidth(conCol���� + 2) - 15
        Else '.RowData(.Row) = 2 Then
            Me.txtInput.Left = .Left + .ColWidth(conCol���) + .ColWidth(conCol����) + .ColWidth(conCol���� + 1) + 45
            Me.txtInput.Width = .ColWidth(conCol���� + 2) - 15
        End If
        Me.txtInput.Text = .Text
        zlWordSelect = InStr(1, Me.txtInput.Text, strWord)
        If zlWordSelect <> 0 Then
            Me.txtInput.SelStart = zlWordSelect - 1
            Me.txtInput.SelLength = Len(strWord)
        End If
        Me.txtInput.Visible = True
        Me.mnuEditString.Visible = True
        Me.txtInput.SetFocus
    End With
    DoEvents

End Function

Public Sub zlWordReplace(lngCurRow As Long, strSource As String, strObject As String)
    '-------------------------------------------------
    '����:�滻ָ���е��ı�����
    '���:  lngCurRow-ָ���У�strSource-ָ�����滻���ı���strObject-�滻Ϊ�ı�
    '-------------------------------------------------
    Me.txtInput.Visible = False
    Me.hgdRefer.Row = lngCurRow
    Me.hgdRefer.Col = conCol���� + 2
    With Me.hgdRefer
        If .RowData(.Row) = 0 Then Exit Sub
        Me.txtInput.Text = .Text
        Me.txtInput.Text = Replace(Me.txtInput.Text, strSource, strObject)
        If .RowData(.Row) = 1 Then
            .TextMatrix(.Row, conCol���� + 1) = Me.txtInput.Text
            .TextMatrix(.Row, conCol���� + 2) = Me.txtInput.Text
        Else '.RowData(.Row) = 2 Then
            .TextMatrix(.Row, conCol���� + 2) = Me.txtInput.Text
        End If
    End With
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub
