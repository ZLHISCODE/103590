VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTendPrintAsk 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��¼����ӡ"
   ClientHeight    =   4500
   ClientLeft      =   2550
   ClientTop       =   2625
   ClientWidth     =   6165
   HelpContextID   =   10322
   Icon            =   "frmTendPrintAsk.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin TabDlg.SSTab SSTPrint 
      Height          =   4380
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   7726
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "��ӡѡ��"
      TabPicture(0)   =   "frmTendPrintAsk.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picPrint"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "�����ӡ"
      TabPicture(1)   =   "frmTendPrintAsk.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraPrint(0)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "��ӡ����"
      TabPicture(2)   =   "frmTendPrintAsk.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkPrintSet(0)"
      Tab(2).Control(1)=   "chkPrintSet(1)"
      Tab(2).Control(2)=   "chkPrintSet(2)"
      Tab(2).ControlCount=   3
      Begin VB.CheckBox chkPrintSet 
         Caption         =   "Ԥ������ӡʱ����δ��ҳ���̶ֹ�������"
         Height          =   255
         Index           =   0
         Left            =   -74850
         TabIndex        =   13
         Top             =   1230
         Width           =   3795
      End
      Begin VB.CheckBox chkPrintSet 
         Caption         =   "Ԥ������ӡʱ������ҳ�Ž������(�ļ�δ������Ч)"
         Height          =   255
         Index           =   1
         Left            =   -74850
         TabIndex        =   12
         Top             =   810
         Width           =   4440
      End
      Begin VB.CheckBox chkPrintSet 
         Caption         =   "��ӡʱ����ҳ��ż���(������ҳ��˳�����)"
         Height          =   255
         Index           =   2
         Left            =   -74850
         TabIndex        =   11
         Top             =   1680
         Width           =   4080
      End
      Begin VB.Frame fraPrint 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1785
         Index           =   0
         Left            =   -74940
         TabIndex        =   6
         Tag             =   "����ش�"
         Top             =   930
         Width           =   4575
         Begin VB.TextBox txtClear 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1710
            MaxLength       =   5
            TabIndex        =   8
            Top             =   300
            Width           =   1035
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "���"
            Height          =   350
            Left            =   2760
            TabIndex        =   7
            Top             =   270
            Width           =   705
         End
         Begin VB.Label lblTag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Чҳ�뷶Χ:��1ҳ �� ��5ҳ"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   0
            Left            =   1095
            TabIndex        =   14
            Top             =   810
            Width           =   2430
         End
         Begin VB.Label lblTag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�������ʼҳ��ʼ�����д�ӡ���ݣ������´�ӡ"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   3
            Left            =   390
            TabIndex        =   10
            Top             =   1290
            Width           =   3780
         End
         Begin VB.Label lblPage 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ʼҳ"
            Height          =   180
            Left            =   1110
            TabIndex        =   9
            Top             =   360
            Width           =   540
         End
      End
      Begin VB.PictureBox picPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3870
         Left            =   30
         ScaleHeight     =   3870
         ScaleWidth      =   4605
         TabIndex        =   1
         Top             =   450
         Width           =   4605
         Begin VSFlex8Ctl.VSFlexGrid vfgPrint 
            Height          =   2655
            Left            =   90
            TabIndex        =   2
            Top             =   675
            Width           =   3405
            _cx             =   6006
            _cy             =   4683
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483643
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   255
            RowHeightMax    =   5000
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmTendPrintAsk.frx":0496
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
            AutoSizeMode    =   1
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   1
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   0   'False
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin XtremeCommandBars.CommandBars cbsMain 
            Left            =   90
            Top             =   60
            _Version        =   589884
            _ExtentX        =   635
            _ExtentY        =   635
            _StockProps     =   0
         End
      End
   End
   Begin VB.CommandButton cmdEXCEL 
      Caption         =   "�����&Excel"
      Height          =   350
      Left            =   4830
      TabIndex        =   5
      Top             =   4080
      Width           =   1245
   End
   Begin VB.CommandButton cmdPreView 
      Caption         =   "Ԥ��(&V)"
      Height          =   350
      Left            =   4830
      TabIndex        =   3
      Top             =   390
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��ӡ(&P)"
      Height          =   350
      Left            =   4830
      TabIndex        =   4
      Top             =   870
      Width           =   1245
   End
   Begin MSComDlg.CommonDialog comDlg 
      Bindings        =   "frmTendPrintAsk.frx":058D
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgData 
      Left            =   660
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendPrintAsk.frx":05A1
            Key             =   "�Ѵ�"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendPrintAsk.frx":0B3B
            Key             =   "δ��"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendPrintAsk.frx":10D5
            Key             =   "����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendPrintAsk.frx":166F
            Key             =   "�ش�"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   3810
      Top             =   1170
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendPrintAsk.frx":1C09
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTendPrintAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public mbytRunMode As Byte        'ִ�з�ʽ
Public mintPageRows As Integer
Public mstrPrintPages As String  '��ʽ��ҳ��;��ʶ(�����������ӡ),ҳ��;��ʶ......

Private mrsData As New ADODB.Recordset
Private mstrSQL As String
Private mbytFileState As Byte
Private Type Type_DataState
    bln���� As Boolean
    bln�ش� As Boolean
    bln���� As Boolean
End Type
Private mDataState As Type_DataState

Private Enum E_CommandBarId
    ID_���� = 1
    ID_�ش� = 2
    ID_���� = 3
End Enum

Public Property Get FileID() As Long
    FileID = glng�ļ�ID
End Property

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim intPrintTag As Integer
    Select Case Control.ID
        Case ID_����
            mbytFileState = 0
            Call LoadPrintData
        Case ID_�ش�
            mbytFileState = 1
            Call LoadPrintData
        Case ID_����
            mbytFileState = 2
            Call LoadPrintData
        Case conMenu_View_Refresh
            Call zlRefresh(glng�ļ�ID)
        Case conMenu_Edit_SelAll * 100# + 1, conMenu_Edit_SelAll * 100# + 2, conMenu_Edit_SelAll * 100# + 3, conMenu_Edit_SelAll * 100# + 4
            If Control.ID = conMenu_Edit_SelAll * 100# + 1 Then
                intPrintTag = 0
            ElseIf Control.ID = conMenu_Edit_SelAll * 100# + 2 Then
                intPrintTag = -1
            ElseIf Control.ID = conMenu_Edit_SelAll * 100# + 3 Then
                intPrintTag = 1
            Else
                intPrintTag = 2
            End If
            Call RevfgPrint(1, intPrintTag)
        Case conMenu_Edit_SelAll * 100# + 5
            Call RevfgPrint(3, intPrintTag)
        Case conMenu_Edit_SelAll * 100# + 6
            Call RevfgPrint(4, intPrintTag)
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngTop As Long, lngLeft As Long, lngRight As Long, lngBottom As Long
    On Error Resume Next
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    With vfgPrint
        .Left = lngLeft
        .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = lngBottom - lngTop
    End With
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim intPrintTag As Integer
    Select Case Control.ID
        Case ID_����
            Control.Enabled = mDataState.bln����
            Control.IconId = IIf(mbytFileState = 0, 90004, 90003)
        Case ID_�ش�
            Control.Enabled = mDataState.bln�ش�
            Control.IconId = IIf(mbytFileState = 1, 90004, 90003)
        Case ID_����
            Control.Enabled = mDataState.bln����
            Control.IconId = IIf(mbytFileState = 2, 90004, 90003)
        Case conMenu_View_Refresh
            Control.Visible = False
        Case conMenu_Edit_SelAll * 100# + 1, conMenu_Edit_SelAll * 100# + 2, conMenu_Edit_SelAll * 100# + 3, conMenu_Edit_SelAll * 100# + 4
            If Control.ID = conMenu_Edit_SelAll * 100# + 1 Then
                intPrintTag = 0
            ElseIf Control.ID = conMenu_Edit_SelAll * 100# + 2 Then
                intPrintTag = -1
            ElseIf Control.ID = conMenu_Edit_SelAll * 100# + 3 Then
                intPrintTag = 1
            Else
                intPrintTag = 2
            End If
            Control.Enabled = RevfgPrint(2, intPrintTag)
        Case conMenu_Edit_SelAll * 100# + 5, conMenu_Edit_SelAll * 100# + 6
            Control.Enabled = mDataState.bln����
    End Select
End Sub

Private Sub cmdClear_Click()
    Dim arrPage() As String
    On Error GoTo ErrHand
    
    If Not vfgPrint.Rows > vfgPrint.FixedRows Then Exit Sub
    If Trim(txtClear.Text) = "" Then Exit Sub
    If Not IsNumeric(txtClear.Text) Then
        MsgBox "��ʼҳ�ź��зǷ��ַ�,���飡", vbInformation, gstrSysName
        If txtClear.Enabled And txtClear.Visible Then txtClear.SetFocus
        Exit Sub
    End If
    If txtClear.Tag = "" Then txtClear.Tag = "0-0"
    arrPage = Split(txtClear.Tag, "-")
    If Not (Val(txtClear.Text) >= Val(arrPage(0)) And Val(txtClear.Text) <= Val(arrPage(1))) Then
        MsgBox "�����ҳ�Ų�����Чҳ�뷶Χ��,���飡", vbInformation, gstrSysName
        If txtClear.Enabled And txtClear.Visible Then txtClear.SetFocus
        Exit Sub
    End If
    
    Call zlDatabase.ExecuteProcedure("ZL_���˻����ӡ_CLEAR(0,0,0," & glng�ļ�ID & "," & Val(txtClear.Text) & ")", "�����ӡ����")
    MsgBox "����ɹ���", vbInformation, gstrSysName
    Call zlRefresh(glng�ļ�ID)
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdPrint_Click()
    If Not PrePrint Then Exit Sub
    mbytRunMode = 1
    Me.Hide
End Sub

Private Sub cmdEXCEL_Click()
    If Not PrePrint Then Exit Sub
    mbytRunMode = 3
    Me.Hide
End Sub

Private Sub cmdPreView_Click()
    If Not PrePrint Then Exit Sub
    mbytRunMode = 2
    Me.Hide
End Sub

Private Function PrePrint() As Boolean
    Dim lngRow As Long, lngPreRow As Long
    Dim strTmp As String, strTag As String
    Dim blnTrue As Boolean
    
    mstrPrintPages = ""
    With vfgPrint
        blnTrue = True
        lngPreRow = -1
        For lngRow = .FixedRows To .Rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("ѡ��"))) <> 0 Then
                mstrPrintPages = mstrPrintPages & "," & Val(.TextMatrix(lngRow, .ColIndex("ҳ��"))) & ";" & IIf(.RowData(lngRow) = -1, IIf(.TextMatrix(lngRow, .ColIndex("״̬")) = "������", 1, 2), 2)
                If .RowData(lngRow) = -1 Then
                    If .TextMatrix(lngRow, .ColIndex("״̬")) = "������" Then
                        strTag = "����"
                    Else
                        strTag = "�ش�"
                    End If
                    strTmp = strTmp & vbCrLf & "��[" & Val(.TextMatrix(lngRow, .ColIndex("ҳ��"))) & "]ҳ��" & strTag & "��"
                End If
                If lngPreRow > -1 And lngPreRow + 1 <> lngRow And blnTrue = True Then blnTrue = False
                lngPreRow = lngRow
            End If
        Next lngRow
    End With
    'ѡ����ż��ӡҳ���������
    If chkPrintSet(2).Value <> 0 And blnTrue = False Then
        If MsgBox("���δ�ӡ��ѡ�˲�������ż��ӡ��������ѡ���ҳ�벻������������ʹ����ż��ӡ���������Ƿ������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
        
    If strTmp <> "" Then
        If MsgBox("���ڱ��δ�ӡ������֮ǰδ��ҳ���Ѵ�ӡ��ҳ,�����Դ�ӡ״̬���к˶ԣ�" & strTmp & vbCrLf & _
            "�������Ƿ������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    mstrPrintPages = Mid(mstrPrintPages, 2)
    
    If mstrPrintPages = "" Then
        MsgBox "����ѡ��ҳ�룡", vbInformation, gstrSysName
        If vfgPrint.Enabled And vfgPrint.Visible Then vfgPrint.SetFocus
        Exit Function
    End If
    Call SaveParam ' �����������
    PrePrint = True
End Function

Private Sub Form_Load()
    Dim objTool As CommandBar
    Dim objControl As CommandBarControl
    Dim objChildControl As CommandBarControl
    On Error GoTo ErrHand
    '��ʼ���˵�
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize False, 24, 24
        .SetIconSize True, 16, 16
    End With
    cbsMain.VisualTheme = xtpThemeOffice2003
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    cbsMain.ActiveMenuBar.Visible = False
    
    '����������
    '-----------------------------------------------------
    Set objTool = cbsMain.Add("������", xtpBarTop)      '����
    objTool.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objTool.ModifyStyle XTP_CBRS_GRIPPER, 0
    'objTool.Closeable = False
    With objTool.Controls
        .Add xtpControlLabel, conMenu_View_Show, "��ʾ��"
        Set objControl = .Add(xtpControlButton, ID_����, "����"): objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, ID_�ش�, "�ش�"):   objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, ID_����, "����"):   objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��"):   objControl.Style = xtpButtonIconAndCaption
        objControl.Flags = xtpFlagRightAlign
        Set objControl = .Add(xtpControlButtonPopup, conMenu_Edit_SelAll, "����ѡ���")
        objControl.Flags = xtpFlagRightAlign: objControl.BeginGroup = True: objControl.Style = xtpButtonIconAndCaption
        With objControl.CommandBar.Controls
            Set objChildControl = .Add(xtpControlButton, conMenu_Edit_SelAll * 100# + 1, "����ӡ(&1)"): objChildControl.ToolTipText = "ѡ������δ���ҳ": objChildControl.Style = xtpButtonCaption
            Set objChildControl = .Add(xtpControlButton, conMenu_Edit_SelAll * 100# + 2, "������(&2)"): objChildControl.ToolTipText = "ѡ�����������ҳ": objChildControl.Style = xtpButtonCaption
            Set objChildControl = .Add(xtpControlButton, conMenu_Edit_SelAll * 100# + 3, "�Ѵ�ӡ(&3)"): objChildControl.ToolTipText = "ѡ�������Ѵ��ҳ": objChildControl.Style = xtpButtonCaption
            Set objChildControl = .Add(xtpControlButton, conMenu_Edit_SelAll * 100# + 4, "���ش�(&4)"): objChildControl.ToolTipText = "ѡ�������ش��ҳ": objChildControl.Style = xtpButtonCaption
            Set objChildControl = .Add(xtpControlButton, conMenu_Edit_SelAll * 100# + 5, "ȫѡ(&A)"): objChildControl.ToolTipText = "ѡ������ҳ": objChildControl.Style = xtpButtonCaption: objChildControl.BeginGroup = True
            Set objChildControl = .Add(xtpControlButton, conMenu_Edit_SelAll * 100# + 6, "ȫ��(&C)"): objChildControl.ToolTipText = "���������ҳ��ѡ��": objChildControl.Style = xtpButtonCaption
        End With
    End With
    cbsMain.KeyBindings.Add 0, VK_F5, conMenu_View_Refresh
    Call zlRefresh(glng�ļ�ID)
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function zlRefresh(ByVal lngFileID As Long) As Boolean
    Dim rsData As New ADODB.Recordset
    Dim lng��ʼҳ�� As Long, lngPati As Long, lngPage As Long, lngBaby As Long
    Dim strSQLNew As String
    On Error GoTo ErrHand
    mintPageRows = 0
    '��ȡ�ļ���Ϣ
    mstrSQL = " Select ����ID,��ҳID,NVL(Ӥ��,0) Ӥ��,�ļ����� From ���˻����ļ� Where ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ���л����ļ�", lngFileID)
    Me.Caption = NVL(rsData!�ļ�����)
    lngPati = rsData!����ID
    lngPage = rsData!��ҳID
    lngBaby = rsData!Ӥ��
    '��ȡ���ļ��Ŀ�ʼҳ��
    mstrSQL = "Select Min(��ʼҳ��) ��ʼҳ�� From ���˻����ӡ Where �ļ�ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ���ļ��Ŀ�ʼҳ��", lngFileID)
    lng��ʼҳ�� = Val(NVL(rsData!��ʼҳ��, 0))
    
    '��ȡ�ļ�������
    mstrSQL = " Select  d.�����ı�" & vbNewLine & _
             " From �����ļ��ṹ d, �����ļ��ṹ p,���˻����ļ� c" & vbNewLine & _
             " Where p.Id = d.��id And p.�ļ�id = c.��ʽID and C.ID=[1] And p.�������� = 1 And p.�����ı� = '�����ʽ' and d.Ҫ������='��Ч������'"
    Set rsData = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ���������", lngFileID)
    If rsData.RecordCount <> 0 Then
        mintPageRows = NVL(rsData!�����ı�, 0)
    End If
    mbytFileState = 0
    mDataState.bln���� = False: mDataState.bln�ش� = False: mDataState.bln���� = False
    '��ȡ��ӡ����б�
    If lng��ʼҳ�� = 0 Then
        strSQLNew = ""
    Else
        strSQLNew = _
        "       Union" & vbNewLine & _
        "       Select ��ʼҳ��, ��ӡҳ��, ������, ��ӡ��ʶ" & vbNewLine & _
        "       From (With ���˻����ļ�_F1 As (Select Id, ����id From ���˻����ļ� Where ����id = [2] And ��ҳid = [3] And Nvl(Ӥ��, 0) = [4])" & vbNewLine & _
        "               Select ����ҳ�� ��ʼҳ��, Decode(��ӡ����ҳ��, ����ҳ��, ����ҳ��, Null) ��ӡҳ��, �����к� ������," & vbNewLine & _
        "                   Decode(��ӡ����ҳ��, ����ҳ��, Decode(��ӡ��, Null, 2, 1), 0) ��ӡ��ʶ" & vbNewLine & _
        "               From ���˻����ӡ a, (Select Id From ���˻����ļ�_F1 Start With ����id = [1] Connect By Prior Id = ����id) b" & vbNewLine & _
        "               Where a.�ļ�id = b.Id And a.����ҳ�� = [5])"
    End If
    mstrSQL = "Select ��ʼҳ��, Max(��ӡҳ��) ��ӡҳ��,Max(������) ������, Decode(Max(��ӡ��ʶ), 1, Decode(Min(��ӡ��ʶ), 0, -1, 1), Max(��ӡ��ʶ)) ��ӡ��ʶ" & vbNewLine & _
        " From (Select ��ʼҳ��, ��ӡҳ��,��ʼ�к�+����-1 ������," & vbNewLine & _
        "              Decode(��ӡҳ��," & vbNewLine & _
        "                      Null," & vbNewLine & _
        "                      0," & vbNewLine & _
        "                      Decode(��ӡҳ��, ��ʼҳ��, Decode(��ӡ�к�, ��ʼ�к�, Decode(��ӡ��, Null, 2, 1), 2), 2)) ��ӡ��ʶ" & vbNewLine & _
        "       From ���˻����ӡ" & vbNewLine & _
        "       Where �ļ�id = [1]" & vbNewLine & _
        "       Union" & vbNewLine & _
        "       Select ����ҳ�� ��ʼҳ��, Decode(��ӡ����ҳ��, ����ҳ��, ����ҳ��, Null) ��ӡҳ��,�����к� ������," & vbNewLine & _
        "              Decode(��ӡ����ҳ��, ����ҳ��, Decode(��ӡ��, Null, 2, 1), 0) ��ӡ��ʶ" & vbNewLine & _
        "       From ���˻����ӡ" & vbNewLine & _
        "       Where �ļ�id = [1] And ����ҳ�� > ��ʼҳ��" & vbNewLine & _
        "       " & strSQLNew & ")" & vbNewLine & _
        " Group By ��ʼҳ��" & vbNewLine & _
        " Order By ��ʼҳ��"
    Set mrsData = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ��ӡ��Ϣ", lngFileID, lngPati, lngPage, lngBaby, lng��ʼҳ��)
    mrsData.Filter = ""
    mDataState.bln���� = mrsData.RecordCount > 0
    mrsData.Filter = "��ӡ��ʶ=-1 OR ��ӡ��ʶ=0"
    mDataState.bln���� = mrsData.RecordCount > 0
    mrsData.Filter = "��ӡ��ʶ=1 OR ��ӡ��ʶ=2"
    mDataState.bln�ش� = mrsData.RecordCount > 0
    
    If mDataState.bln���� = True Then
        mbytFileState = 0
    ElseIf mDataState.bln�ش� = True Then
        mbytFileState = 1
    Else
        mbytFileState = 2
    End If
    cmdPreView.Enabled = mDataState.bln����
    cmdPrint.Enabled = cmdPreView.Enabled
    cmdEXCEL.Enabled = cmdPreView.Enabled
    
    txtClear.Tag = "0-0"
    If mDataState.bln���� = False Then
        lblTag(0).Caption = "���ļ���δ¼������"
        cmdClear.Enabled = False
    Else
        mrsData.Filter = ""
        txtClear.Tag = Val(NVL(mrsData!��ʼҳ��))
        lblTag(0).Caption = "��Чҳ�뷶Χ:��" & mrsData!��ʼҳ�� & "ҳ"
        mrsData.MoveLast
        txtClear.Tag = txtClear.Tag & "-" & Val(NVL(mrsData!��ʼҳ��))
        lblTag(0).Caption = lblTag(0).Caption & " �� ��" & mrsData!��ʼҳ�� & "ҳ"
        cmdClear.Enabled = True
    End If
    lblTag(0).Left = (fraPrint(0).Width - lblTag(0).Width) \ 2
    If lblTag(0).Left < 0 Then lblTag(0).Left = 0
    
    Call LoadPrintData '��ʾ��ӡ�б�����
    Call LoadParam '���ز���
    
    zlRefresh = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function LoadPrintData() As Boolean
'����:�����ļ���Ϣ�Լ���ӡ�����Ϣ
    Dim lngRow As Long
    Dim strTag As String
    Dim stdPic As StdPicture
    
    On Error GoTo ErrHand
    
    Select Case mbytFileState
        Case 0
            mrsData.Filter = "��ӡ��ʶ=-1 OR ��ӡ��ʶ=0"
        Case 1
            mrsData.Filter = "��ӡ��ʶ=1 OR ��ӡ��ʶ=2"
        Case 2
            mrsData.Filter = ""
    End Select
    
    With vfgPrint
        .FixedRows = 1
        .FixedCols = 1
        .Rows = .FixedRows
        .Cols = 7
        .Editable = flexEDKbdMouse
        .MergeCells = flexMergeFixedOnly
        .MergeCellsFixed = flexMergeRestrictColumns
        .MergeRow(0) = True
        .MergeCol(.ColIndex("ͼƬ")) = True
        .MergeCol(.ColIndex("ѡ��")) = True
        .ColHidden(.ColIndex("�Ƿ���ҳ")) = True
         Set .Cell(flexcpPicture, 0, 0, .Rows - 1, .Cols - 1) = Nothing
        Do While Not mrsData.EOF
            Select Case Val(mrsData!��ӡ��ʶ)
                Case -1 '����ҳ
                    strTag = "������"
                    Set stdPic = imgData.ListImages("����").Picture
                Case 0  'δ��ҳ
                    strTag = "����ӡ"
                    Set stdPic = imgData.ListImages("δ��").Picture
                Case 1 '�Ѵ�ҳ
                    strTag = "�Ѵ�ӡ"
                    Set stdPic = imgData.ListImages("�Ѵ�").Picture
                Case 2 '�ش�ҳ
                    strTag = "���ش�"
                    Set stdPic = imgData.ListImages("�ش�").Picture
            End Select
            If mrsData.AbsolutePosition + .FixedRows > .Rows Then .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("ѡ��")) = 0
            .TextMatrix(.Rows - 1, .ColIndex("ҳ��")) = CStr(mrsData!��ʼҳ��)
            .TextMatrix(.Rows - 1, .ColIndex("״̬")) = strTag
            .TextMatrix(.Rows - 1, .ColIndex("�Ƿ���ҳ")) = IIf(Val(NVL(mrsData!������, 0)) >= mintPageRows, "��", "��")
            .TextMatrix(.Rows - 1, .ColIndex("��ӡҳ��")) = CStr(NVL(mrsData!��ӡҳ��))
            .RowData(.Rows - 1) = Val(mrsData!��ӡ��ʶ)
            Set .Cell(flexcpPicture, .Rows - 1, 1) = stdPic
        mrsData.MoveNext
        Loop
        
        If .FixedRows < .Rows Then
            .Cell(flexcpBackColor, .FixedRows, .ColIndex("ѡ��"), .Rows - 1, .Cols - 1) = &H80000005
            .RowSel = .FixedRows
            .ColSel = .ColIndex("ѡ��")
        End If
        If .Enabled = True And .Visible = True Then .SetFocus
    End With
    
    LoadPrintData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub AdjustRowFlag(ByRef objVsf As Object, ByVal intRow As Integer)
    '-----------------------------------------------------------------------------------------
    '����:
    '����:
    '-----------------------------------------------------------------------------------------
    If objVsf.FixedCols = 0 Then Exit Sub
    If objVsf.Rows <= objVsf.FixedRows Then Exit Sub
    If Not (objVsf.Cell(flexcpPicture, intRow, 0) Is Nothing) Then Exit Sub
    Set objVsf.Cell(flexcpPicture, 0, 0, objVsf.Rows - 1, 0) = Nothing
    Set objVsf.Cell(flexcpPicture, intRow, 0) = ils16.ListImages(1).Picture
End Sub

Private Function RevfgPrint(ByVal byt���� As Byte, Optional ByVal intPrintTag As Integer = 0) As Boolean
'byt����:
'       1:����intPrintTag����ĳ���ʶ��ѡ��
'       2:����intPrintTag�ж϶�Ӧ�ı�ʶ�Ƿ����,����TRUE OR False
'       3:ȫѡ
'       4:ȫ��
'intPrintTag:��ӡ��ʶ -1 -����0 -δ��;1 -�Ѵ�;2 -�ش�   byt����=3,4����
    Dim lngRow As Long
    Dim blnOK As Boolean
    
    blnOK = False
    With vfgPrint
        For lngRow = .FixedRows To .Rows - 1
            Select Case byt����
                Case 1
                    If Val(.RowData(lngRow)) = intPrintTag Then
                        .TextMatrix(lngRow, .ColIndex("ѡ��")) = 1
                        .Cell(flexcpBackColor, lngRow, .ColIndex("ѡ��"), lngRow, .Cols - 1) = RGB(135, 206, 235)
                    Else
                        .TextMatrix(lngRow, .ColIndex("ѡ��")) = 0
                        .Cell(flexcpBackColor, lngRow, .ColIndex("ѡ��"), lngRow, .Cols - 1) = &H80000005
                    End If
                Case 2
                    If Val(.RowData(lngRow)) = intPrintTag And .TextMatrix(lngRow, .ColIndex("ҳ��")) <> "" Then
                        blnOK = True
                        Exit For
                    End If
                Case 3
                    .TextMatrix(lngRow, .ColIndex("ѡ��")) = 1
                    .Cell(flexcpBackColor, lngRow, .ColIndex("ѡ��"), lngRow, .Cols - 1) = RGB(135, 206, 235)
                Case 4
                    .TextMatrix(lngRow, .ColIndex("ѡ��")) = 0
                    .Cell(flexcpBackColor, lngRow, .ColIndex("ѡ��"), lngRow, .Cols - 1) = &H80000005
            End Select
            
        Next lngRow
        If byt���� <> 2 Then
            If .RowSel >= .FixedRows And .Rows > .FixedRows Then
                .BackColorSel = .Cell(flexcpBackColor, .RowSel, .ColIndex("ѡ��"), .RowSel, .Cols - 1)
            End If
        End If
    End With
    If byt���� = 2 Then
        RevfgPrint = blnOK
    Else
        RevfgPrint = True
    End If
End Function
 
            
Private Sub Form_Unload(Cancel As Integer)
    Call SaveParam
    mbytRunMode = 0
End Sub


Private Sub SaveParam()
'�����ӡ��ز���
    '--56134:������,2012-12-19,��¼����ӡʱ,����δ��ҳ�հײ���������
    Call zlDatabase.SetPara("��¼��δ��ҳ��ӡ���", chkPrintSet(0).Value, glngSys, 1255)
    '--46506:������,2012-12-19,��¼����ӡʱ��������ҳ�Ž������(�ļ�Ϊ����ʱ��Ч)
    Call zlDatabase.SetPara("��¼����ҳ��ӡ", chkPrintSet(1).Value, glngSys, 1255)
    '--49753:������,2012-12-19,��¼����ӡʱ������ҳ��ż���
    Call zlDatabase.SetPara("��¼����ż��ӡ", chkPrintSet(2).Value, glngSys, 1255)
End Sub

Private Sub LoadParam()
'���ش�ӡ��ز���
    '--56134:������,2012-12-19,��¼����ӡʱ,����δ��ҳ�հײ���������
    chkPrintSet(0).Value = Val(zlDatabase.GetPara("��¼��δ��ҳ��ӡ���", glngSys, 1255, "0", Array(chkPrintSet(0)), True))
    '--46506:������,2012-12-19,��¼����ӡʱ��������ҳ�Ž������(�ļ�Ϊ����ʱ��Ч)
    chkPrintSet(1).Value = Val(zlDatabase.GetPara("��¼����ҳ��ӡ", glngSys, 1255, "0", Array(chkPrintSet(1)), True))
    '--49753:������,2012-12-19,��¼����ӡʱ������ҳ��ż���
    chkPrintSet(2).Value = Val(zlDatabase.GetPara("��¼����ż��ӡ", glngSys, 1255, "0", Array(chkPrintSet(2)), True))
End Sub

Private Sub Label1_Click(Index As Integer)

End Sub

Private Sub txtClear_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txtClear, KeyAscii, m����ʽ)
End Sub

Private Sub vfgPrint_AfterEdit(ByVal ROW As Long, ByVal COL As Long)
    Dim intValue As Integer
    Dim lngRow As Long, lngStartRow As Long, lngEndRow As Long
    With vfgPrint
        If ROW < .FixedRows Then Exit Sub
        If COL = .ColIndex("ѡ��") Then
            intValue = Val(.TextMatrix(ROW, COL))
            .Cell(flexcpBackColor, ROW, COL, ROW, .Cols - 1) = IIf(intValue = 0, &H80000005, RGB(135, 206, 235))
            .BackColorSel = IIf(intValue = 0, &H80000005, RGB(135, 206, 235))
            '�ж�Shift���Ƿ�������������������ѡ(����windows�ļ�ѡ��)
            If (GetAsyncKeyState(vbKeyShift) And &H8000) = &H8000 And intValue <> 0 Then
                lngStartRow = -1
                lngEndRow = -1
                For lngRow = ROW - 1 To .FixedRows Step -1
                    If Val(.TextMatrix(lngRow, COL)) <> 0 Then
                        lngStartRow = lngRow
                        lngEndRow = ROW
                        Exit For
                    End If
                Next lngRow
                If lngStartRow = -1 Then
                    For lngRow = ROW + 1 To .Rows - 1
                        If Val(.TextMatrix(lngRow, COL)) <> 0 Then
                            lngStartRow = ROW
                            lngEndRow = lngRow
                            Exit For
                        End If
                    Next lngRow
                End If
                If lngStartRow < lngEndRow Then
                    For lngRow = lngStartRow To lngEndRow
                        .TextMatrix(lngRow, COL) = 1
                        .Cell(flexcpBackColor, lngRow, COL, lngRow, .Cols - 1) = RGB(135, 206, 235)
                    Next
                End If
            End If
        ElseIf COL = .ColIndex("״̬") Then
            If .TextMatrix(ROW, COL) = "" Then .TextMatrix(ROW, COL) = "������"
        End If
    End With
End Sub

Private Sub vfgPrint_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim blnCancle As Boolean
    vfgPrint.ColComboList(NewCol) = ""
    
    Call AdjustRowFlag(vfgPrint, NewRow)
    Call vfgPrint_StartEdit(NewRow, NewCol, blnCancle)
    If blnCancle = False And vfgPrint.ColIndex("״̬") = NewCol Then
        vfgPrint.ColComboList(NewCol) = "������|���ش�"
    End If
End Sub

Private Sub vfgPrint_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    vfgPrint.BackColorSel = vfgPrint.Cell(flexcpBackColor, NewRowSel, vfgPrint.ColIndex("ѡ��"), NewRowSel, vfgPrint.Cols - 1)
End Sub

Private Sub vfgPrint_DblClick()
    Call vfgPrint_KeyPress(vbKeySpace)
End Sub

Private Sub vfgPrint_KeyPress(KeyAscii As Integer)
    Dim intValue As Integer
    If KeyAscii = vbKeySpace And vfgPrint.ROW >= vfgPrint.FixedRows And vfgPrint.ROW < vfgPrint.Rows And vfgPrint.COL <> vfgPrint.ColIndex("ѡ��") Then
        intValue = Val(vfgPrint.TextMatrix(vfgPrint.ROW, vfgPrint.ColIndex("ѡ��")))
        vfgPrint.TextMatrix(vfgPrint.ROW, vfgPrint.ColIndex("ѡ��")) = IIf(intValue = 0, 1, 0)
        intValue = Val(vfgPrint.TextMatrix(vfgPrint.ROW, vfgPrint.ColIndex("ѡ��")))
        vfgPrint.Cell(flexcpBackColor, vfgPrint.ROW, vfgPrint.ColIndex("ѡ��"), vfgPrint.ROW, vfgPrint.Cols - 1) = IIf(intValue = 0, &H80000005, RGB(135, 206, 235))
        vfgPrint.BackColorSel = IIf(intValue = 0, &H80000005, RGB(135, 206, 235))
    End If
End Sub

Private Sub vfgPrint_StartEdit(ByVal ROW As Long, ByVal COL As Long, Cancel As Boolean)
    With vfgPrint
        If ROW >= .FixedRows Then
            If COL = .ColIndex("ѡ��") Then
                Cancel = False
            ElseIf COL = .ColIndex("״̬") And .RowData(ROW) = -1 Then
                Cancel = False
            Else
                Cancel = True
            End If
        Else
            Cancel = True
        End If
    End With
End Sub
