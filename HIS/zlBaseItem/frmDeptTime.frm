VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDeptTime 
   AutoRedraw      =   -1  'True
   Caption         =   "����ʱ�䰲��"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6615
   Icon            =   "frmDeptTime.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4740
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh 
      Height          =   3615
      Left            =   2400
      TabIndex        =   4
      Top             =   750
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   6376
      _Version        =   393216
      Cols            =   3
      RowHeightMin    =   300
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      MergeCells      =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4065
      Left            =   2445
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4065
      ScaleWidth      =   45
      TabIndex        =   3
      Top             =   795
      Width           =   45
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   3630
      Left            =   0
      TabIndex        =   2
      Top             =   735
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   6403
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   6615
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   645
      Width1          =   810
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   645
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   1138
         ButtonWidth     =   820
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Description     =   "Ԥ��"
               Object.ToolTipText     =   "��ӡԤ��"
               Object.Tag             =   "Ԥ��"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Plan"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageKey        =   "Plan"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split2"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   75
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptTime.frx":0E42
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptTime.frx":105C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptTime.frx":1276
            Key             =   "Plan"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptTime.frx":1490
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptTime.frx":16AA
            Key             =   "Cancel"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptTime.frx":18C4
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptTime.frx":1ADE
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptTime.frx":1CF8
            Key             =   "Del"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   705
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptTime.frx":1F12
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptTime.frx":212C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptTime.frx":2346
            Key             =   "Plan"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptTime.frx":2560
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptTime.frx":277A
            Key             =   "Cancel"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptTime.frx":2994
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptTime.frx":2BAE
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptTime.frx":2DC8
            Key             =   "Del"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1125
      Top             =   1845
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptTime.frx":2FE2
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptTime.frx":3E34
            Key             =   "Unit"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshData 
      Height          =   1230
      Left            =   2985
      TabIndex        =   6
      Top             =   2025
      Visible         =   0   'False
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   2170
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      Redraw          =   0   'False
      HighLight       =   0
      ScrollBars      =   0
      MergeCells      =   1
      BorderStyle     =   0
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   4380
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   635
      SimpleText      =   $"frmDeptTime.frx":5B3E
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDeptTime.frx":5B85
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6588
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
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   3600
      TabIndex        =   5
      Top             =   2250
      Width           =   540
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFile_PrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFile_PreView 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile_Excel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnusplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_quit 
         Caption         =   "�˳�(&X)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEdit_Plan 
         Caption         =   "����(&P)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_All 
         Caption         =   "ȫ������(&A)"
      End
      Begin VB.Menu mnuEdit_Clear 
         Caption         =   "ȫ�����(&C)"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "����(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuReportItem 
         Caption         =   "-"
         Index           =   0
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
      Begin VB.Menu mnuviewsplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelp_Index 
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
      Begin VB.Menu mnuHelp_About 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmDeptTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intDay As Integer
Private mlngMode As Long
Private mstrPrivs As String                              'Ȩ�޴�
Private Sub Form_Load()
    mlngMode = glngModul
    mstrPrivs = gstrPrivs
    
    RestoreWinState Me, App.ProductName
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    If InStr(1, mstrPrivs, "����") = 0 Then
        mnuEdit_Plan.Enabled = False
        mnuEdit_Clear.Enabled = False
        mnuEdit_All.Enabled = False
        mnuEdit.Visible = False
        tbr.Buttons("Plan").Enabled = False
        tbr.Buttons("Plan").Visible = False
        tbr.Buttons("Split2").Visible = False
    End If
    
    lblInfo.Caption = "��ѡ�����Ĳ��ţ�" & vbCrLf & "��[����]���в��ţ�"
    '0-����,1-��һ
    intDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7
    InitTree
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long '������ռ�ø߶�s
    Dim staH As Long '״̬��ռ�ø߶�
    
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    
    '����ؼ���Ⱥ͸߶�
    cbrH = IIF(cbr.Visible, cbr.Height, 0)
    staH = IIF(stbThis.Visible, stbThis.Height, 0)
    
    tvw.Left = 0
    tvw.Top = cbrH
    tvw.Width = Me.ScaleWidth - pic.Width - msh.Width
    tvw.Height = Me.ScaleHeight - cbrH - staH
    
    pic.Left = tvw.Width
    pic.Top = tvw.Top
    pic.Height = tvw.Height
    
    msh.Top = pic.Top
    msh.Left = pic.Left + pic.Width
    msh.Height = pic.Height
    
    lblInfo.Top = msh.Top + (msh.Height - lblInfo.Height) / 2
    lblInfo.Left = msh.Left + (msh.Width - lblInfo.Width) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub mnuEdit_All_Click()
    Dim i As Byte
    If MsgBox("ȷʵҪ�������в���ȫ���ϰ���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        On Error GoTo errH
        gstrSQL = "zl_���Ű���_new"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        tvw_NodeClick tvw.SelectedItem
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEdit_Clear_Click()
    If MsgBox("ȷʵҪȡ�����в��ŵ��ϰల����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        On Error GoTo errH
        gstrSQL = "zl_���Ű���_delete"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        tvw_NodeClick tvw.SelectedItem
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEdit_Plan_Click()
    If tvw.SelectedItem.Key = "Root" Then Exit Sub
    frmDeptTimeEdit.Show 1, Me
    tvw_NodeClick tvw.SelectedItem
    
End Sub

Private Sub mnuFile_Excel_Click()
    OutputList 3
End Sub

Private Sub mnuFile_PreView_Click()
    OutputList 2
End Sub

Private Sub mnuFile_Print_Click()
    OutputList 1
End Sub

Private Sub mnuFile_PrintSet_Click()
    zlPrintSet
End Sub

Private Sub mnuHelp_Index_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    'Ĭ�ϲ���������=����ID
    Dim lng����ID As Long
    
    If Not Me.tvw.SelectedItem Is Nothing Then
        If Me.tvw.SelectedItem.Key <> "Root" Then
            lng����ID = Mid(Me.tvw.SelectedItem.Key, 2)
        End If
    End If
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "����=" & IIF(lng����ID = 0, "", lng����ID))
End Sub

Private Sub mnuViewRefresh_Click()
    tvw_NodeClick tvw.SelectedItem
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
    cbr.Visible = Not cbr.Visible
    Call Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intCount As Integer

    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    If mnuViewToolText.Checked Then
        For intCount = 1 To Me.tbr.Buttons.Count
            tbr.Buttons(intCount).Caption = tbr.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To tbr.Buttons.Count
            tbr.Buttons(intCount).Caption = ""
        Next
    End If
    Me.cbr.Bands(1).MinHeight = tbr.Height
    Me.cbr.Refresh
    Call Form_Resize
End Sub

Private Sub mnuFile_quit_Click()
    Unload Me
End Sub

Private Sub mnuHelp_About_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub msh_EnterCell()
    Dim i As Integer, j As Integer
    Dim intPreRow As Integer, intPreCol As Integer
    
    msh.redraw = False
    
    intPreRow = msh.Row
    intPreCol = msh.Col
    
    For i = 1 To msh.Rows - 1
        msh.Row = i
        For j = 1 To msh.Cols - 1
            msh.Col = j
            If msh.Row = intPreRow Then
                msh.CellBackColor = &HCFAB9E
            Else
                msh.CellBackColor = &H80000005  '&HCFAB9E
            End If
        Next
    Next
    
    msh.Row = intPreRow
    msh.Col = intPreCol
    msh.redraw = True
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If tvw.Width + X < 1500 Or msh.Width - X < 3000 Then Exit Sub
        pic.Left = pic.Left + X
        tvw.Width = tvw.Width + X
        msh.Left = msh.Left + X
        msh.Width = msh.Width - X
        lblInfo.Left = msh.Left + (msh.Width - lblInfo.Width) / 2
    End If
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_quit_Click
        Case "Plan"
            mnuEdit_Plan_Click
        Case "Print"
            mnuFile_Print_Click
        Case "Preview"
            mnuFile_PreView_Click
        Case "Help"
            mnuHelp_Index_Click
    End Select
End Sub

Private Sub InitTree()
    Dim rsTmp As New ADODB.Recordset
    Dim objNode As Node
    Dim i As Integer
    
    On Error GoTo errH
    
    tvw.Nodes.Clear
    Set objNode = tvw.Nodes.Add(, , "Root", "���в���", "Root")
    objNode.Tag = "���в���"
    objNode.Expanded = True
    objNode.Selected = True
    
    gstrSQL = "Select Level,A.* From ���ű� A Where ����ʱ��=TO_Date('3000-01-01','YYYY-MM-DD') Start With �ϼ�ID is NULL Connect by prior ID=�ϼ�ID Order by Level,����"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    For i = 1 To rsTmp.RecordCount
        If rsTmp!Level = 1 Then
            Set objNode = tvw.Nodes.Add("Root", 4, "_" & rsTmp!ID, "��" & rsTmp!���� & "��" & rsTmp!����, "Unit")
        Else
            Set objNode = tvw.Nodes.Add("_" & rsTmp!�ϼ�id, 4, "_" & rsTmp!ID, "��" & rsTmp!���� & "��" & rsTmp!����, "Unit")
        End If
        objNode.Tag = rsTmp!����
        rsTmp.MoveNext
    Next
    If tvw.Nodes.Count > 1 Then tvw.Nodes(1).Child.Selected = True
    tvw_NodeClick tvw.SelectedItem
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool, 2
    End If
End Sub

Private Sub tvw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And mnuEdit.Visible = True Then PopupMenu mnuEdit, 2
End Sub

Public Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    stbThis.Panels(2) = Node.Text & IIF(Node.Children = 0, "", ",�� " & Node.Children & " ���¼�����.")
    If Node.Key = "Root" Then
        msh.Visible = False
    Else
        msh.Visible = True
        ShowPlan CLng(Mid(Node.Key, 2))
    End If
End Sub

Private Sub ShowPlan(lngId As Long)
    Dim i As Integer, j As Integer
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    
    rsTmp.CursorLocation = adUseClient
    gstrSQL = "Select ����,��ʼʱ��,��ֹʱ�� From ���Ű��� Where ����ID=[1] Order by ��ʼʱ��"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngId)
        
    msh.redraw = False
    msh.Clear: msh.Rows = 2
    msh.TextMatrix(0, 0) = "����"
    msh.TextMatrix(0, 1) = "��ʼʱ��"
    msh.TextMatrix(0, 2) = "����ʱ��"
    For i = 0 To 6
        rsTmp.Filter = "����=" & i
        If rsTmp.RecordCount = 0 Then
            If i <> 0 Then msh.Rows = msh.Rows + 1
            msh.TextMatrix(msh.Rows - 1, 0) = Switch(i = 0, "����", i = 1, "��һ", i = 2, "�ܶ�", i = 3, "����", i = 4, "����", i = 5, "����", i = 6, "����")
            msh.RowData(msh.Rows - 1) = i
        Else
            rsTmp.MoveFirst: j = 0
            Do While Not rsTmp.EOF
                If Not (i = 0 And j = 0) Then msh.Rows = msh.Rows + 1
                msh.TextMatrix(msh.Rows - 1, 0) = Switch(i = 0, "����", i = 1, "��һ", i = 2, "�ܶ�", i = 3, "����", i = 4, "����", i = 5, "����", i = 6, "����")
                msh.TextMatrix(msh.Rows - 1, 1) = Format(IIF(IsNull(rsTmp!��ʼʱ��), "", rsTmp!��ʼʱ��), "hh:mm:ss")
                msh.TextMatrix(msh.Rows - 1, 2) = Format(IIF(IsNull(rsTmp!��ֹʱ��), "", rsTmp!��ֹʱ��), "hh:mm:ss")
                msh.RowData(msh.Rows - 1) = i
                j = j + 1
                rsTmp.MoveNext
            Loop
        End If
    Next
    
    '��ʽ����
    For i = 1 To msh.Rows - 1
        msh.Row = i
        For j = 0 To msh.Cols - 1
            msh.Col = j: msh.CellAlignment = 4
            If msh.RowData(i) = intDay And msh.Col = 0 Then
                msh.CellForeColor = &HFF0000
                msh.CellFontBold = True
            End If
        Next
    Next
    
    msh.MergeCol(0) = True
    msh.Row = 1: msh.Col = 1
    msh.ColWidth(0) = 600
    msh.ColWidth(1) = 1500
    msh.ColWidth(2) = 1500
    Call msh_EnterCell
    msh.redraw = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub OutputList(bytStyle As Byte)
'���ܣ�������б�
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    Dim objOut As New zlPrint1Grd
    Dim bytR As Byte
    
    On Error GoTo errH
    
    gstrSQL = _
        "SeLect A.����,B.����,B.��ʼʱ��,B.��ֹʱ�� From " & _
        "(Select ID,LPAD(' ',(Level-1)*2,' ')||'['||����||']'||���� as ���� " & _
        "From ���ű� " & _
        "Where ����ʱ��=To_DATE('3000-01-01','YYYY-MM-DD') " & _
        "Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID) A, " & _
        "(Select ����ID, " & _
        "'��'||Decode(����,0,'��',1,'һ',2,'��',3,'��',4,'��',5,'��',6,'��') as ����, " & _
        "' '||To_Char(��ʼʱ��,'HH24:MI:SS') as ��ʼʱ��, " & _
        "' '||To_Char(��ֹʱ��,'HH24:MI:SS') as ��ֹʱ�� From ���Ű���) B " & _
        "Where B.����ID = A.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

    If rsTmp.RecordCount = 0 Then Exit Sub

    Set mshData.DataSource = rsTmp
    mshData.ColWidth(0) = 3000
    mshData.ColWidth(1) = 800
    mshData.ColWidth(2) = 1500
    mshData.ColWidth(3) = 1500
    mshData.Row = 0
    For i = 0 To mshData.Cols - 1
        mshData.Col = i: mshData.CellAlignment = 4
        If i <> 0 Then mshData.ColAlignment(i) = 4
    Next
    mshData.MergeCol(0) = True
    mshData.MergeCol(1) = True
    
    objOut.Title.Text = "�����ϰల�ű�"
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    
    Set objOut.Body = mshData
    
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hwnd
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

