VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmRunLimitManage 
   BackColor       =   &H80000005&
   Caption         =   "������ʱ����"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12360
   ControlBox      =   0   'False
   Icon            =   "frmRunLimitManage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "form3"
   MDIChild        =   -1  'True
   Picture         =   "frmRunLimitManage.frx":6852
   ScaleHeight     =   8655
   ScaleWidth      =   12360
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imgPlanDetail 
      Left            =   11730
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   97
      ImageHeight     =   30
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunLimitManage.frx":6D4B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picLeft 
      BorderStyle     =   0  'None
      Height          =   6645
      Left            =   0
      ScaleHeight     =   6645
      ScaleWidth      =   3405
      TabIndex        =   5
      Top             =   1290
      Width           =   3405
      Begin VB.PictureBox picPlanList 
         Height          =   3480
         Left            =   45
         ScaleHeight     =   3420
         ScaleWidth      =   3225
         TabIndex        =   9
         Top             =   3060
         Width           =   3285
         Begin VB.CommandButton cmdCancel 
            Caption         =   "ȡ��(&C)"
            Height          =   350
            Left            =   2055
            TabIndex        =   20
            Top             =   3030
            Width           =   1100
         End
         Begin VB.OptionButton optRemind 
            Caption         =   "��ֹ"
            Height          =   375
            Index           =   0
            Left            =   1005
            TabIndex        =   12
            Top             =   1020
            Width           =   1005
         End
         Begin VB.OptionButton optRemind 
            Caption         =   "����"
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   11
            Top             =   1095
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "����(&S)"
            Height          =   350
            Left            =   855
            TabIndex        =   14
            Top             =   3030
            Width           =   1100
         End
         Begin VB.TextBox txtLimitReason 
            Height          =   1245
            Left            =   120
            MaxLength       =   125
            MultiLine       =   -1  'True
            TabIndex        =   13
            Top             =   1665
            Width           =   3060
         End
         Begin VB.CommandButton cmdPlanSet 
            Caption         =   "��������(&S)"
            Height          =   350
            Left            =   1980
            TabIndex        =   16
            Top             =   390
            Width           =   1200
         End
         Begin VB.ComboBox cboPlanList 
            Height          =   300
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   420
            Width           =   1755
         End
         Begin VB.Label lblOption 
            AutoSize        =   -1  'True
            Caption         =   "����ѡ��"
            Height          =   180
            Left            =   135
            TabIndex        =   18
            Top             =   840
            Width           =   720
         End
         Begin VB.Label lblLimitReason 
            AutoSize        =   -1  'True
            Caption         =   "��ʱԭ��"
            Height          =   180
            Left            =   135
            TabIndex        =   17
            Top             =   1425
            Width           =   720
         End
         Begin VB.Label lblSelectPlan 
            AutoSize        =   -1  'True
            Caption         =   "��ǰ����"
            Height          =   180
            Left            =   135
            TabIndex        =   15
            Top             =   180
            Width           =   720
         End
      End
      Begin MSComctlLib.TreeView tvwModuleTree 
         Height          =   2565
         Left            =   45
         TabIndex        =   6
         Top             =   0
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   4524
         _Version        =   393217
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "img16"
         Appearance      =   1
      End
      Begin VB.PictureBox picTopOrButtom 
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   45
         MousePointer    =   7  'Size N S
         ScaleHeight     =   210
         ScaleWidth      =   3000
         TabIndex        =   22
         Top             =   2655
         Width           =   3000
      End
   End
   Begin VB.PictureBox picRight 
      Height          =   6630
      Left            =   3510
      ScaleHeight     =   6570
      ScaleWidth      =   8415
      TabIndex        =   3
      Top             =   1290
      Width           =   8475
      Begin VB.PictureBox picPlanDetail 
         BorderStyle     =   0  'None
         Height          =   5220
         Left            =   15
         ScaleHeight     =   5220
         ScaleWidth      =   8130
         TabIndex        =   7
         Top             =   1035
         Width           =   8130
         Begin VSFlex8Ctl.VSFlexGrid vsfPlanDetail 
            Height          =   4650
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   7905
            _cx             =   13944
            _cy             =   8202
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   16774866
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   16774866
            GridColor       =   -2147483633
            GridColorFixed  =   15984570
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483633
            FocusRect       =   1
            HighLight       =   0
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   0
            GridLineWidth   =   1
            Rows            =   8
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmRunLimitManage.frx":B021
            ScrollTrack     =   0   'False
            ScrollBars      =   1
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
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
            OwnerDraw       =   0
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
            AutoSizeMouse   =   -1  'True
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
      End
      Begin XtremeSuiteControls.TabControl tbcPage 
         Height          =   960
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1755
         _Version        =   589884
         _ExtentX        =   3096
         _ExtentY        =   1693
         _StockProps     =   64
      End
   End
   Begin VB.Frame fraMidButtomSplit 
      Height          =   45
      Left            =   0
      TabIndex        =   2
      Top             =   8055
      Width           =   12360
   End
   Begin VB.Frame frmTopMidSplit 
      Height          =   45
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   12360
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   11010
      Top             =   90
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
            Picture         =   "frmRunLimitManage.frx":B0B5
            Key             =   "system"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunLimitManage.frx":11917
            Key             =   "function"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunLimitManage.frx":18179
            Key             =   "program"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picLeftOrRight 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   2820
      Left            =   3125
      MousePointer    =   9  'Size W E
      ScaleHeight     =   2820
      ScaleWidth      =   30
      TabIndex        =   21
      Top             =   1290
      Width           =   30
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "ϵͳ��Σ���ܵ�ִ��ʱ�����"
      Height          =   180
      Left            =   1050
      TabIndex        =   19
      Top             =   780
      Width           =   2340
   End
   Begin VB.Image imgDescription 
      Height          =   720
      Left            =   150
      Picture         =   "frmRunLimitManage.frx":1E9DB
      Top             =   495
      Width           =   720
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������ʱ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   0
      Top             =   135
      Width           =   1440
   End
End
Attribute VB_Name = "frmRunLimitManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrLastPlan As String '��¼��һ��ѡ��ķ�������
Private mcolPlanDescription As Collection '��¼��������
Private mrsPlanList As ADODB.Recordset
Private Const vsfTitleBackColor = &HF0E5BD  '�������ݱ����ⱳ����ɫ
Private Const vsfContentBackColor = &HFFFAE4 '�������ݱ�����ݲ���ǳɫ����ɫ
Private Const vsfTitleHeight = 500
Private Const vsfRowHeight = 1000
Private Enum PlanList
    PL_��� = 0
    PL_���� = 1
    PL_���� = 2
    PL_����ѡ�� = 3
    PL_���� = 4
End Enum
Private Enum PlanDetailTitle
    PDT_���� = 0
    PDT_ʱ���1 = 1
    PDT_ʱ�����չ = 2
End Enum
Private Enum PlanDetail
    PD_���� = 0
    PD_������ = 1
    PD_����һ = 2
    PD_���ڶ� = 3
    PD_������ = 4
    PD_������ = 5
    PD_������ = 6
    PD_������ = 7
End Enum

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Private Sub cboPlanList_Click()
'ѡ��ĳ�������������Ҳ���潫�÷�������ϸʱ�䰲��չʾ����
    If cboPlanList.Text = "[�޷�������]" Then
        tbcPage.Item(0).Caption = "�޷���"
        optRemind(lblOption.Tag).value = True
        txtLimitReason.Text = lblLimitReason.Tag
        optRemind(0).Enabled = False
        optRemind(1).Enabled = False
        txtLimitReason.Enabled = False
        txtLimitReason.ForeColor = &H80000012
        txtLimitReason.BackColor = &H8000000F
    Else
        tbcPage.Item(0).Caption = cboPlanList.Text
        optRemind(0).Enabled = True
        optRemind(1).Enabled = True
        txtLimitReason.Enabled = True
        txtLimitReason.BackColor = &H80000005
    End If

    tbcPage.Item(0).Caption = IIf(cboPlanList.Text = "[�޷�������]", "�޷���", cboPlanList.Text)
    If Val(cboPlanList.Tag) = 0 Then
        Call SetEnabled(True)
        Call FillPlanDetail(cboPlanList.ItemData(cboPlanList.ListIndex))
    End If
End Sub

Private Sub cmdCancel_Click()
    '�����ȡ���󣬻ָ���ʼ״̬
    cboPlanList.ListIndex = lblSelectPlan.Tag
    optRemind(lblOption.Tag).value = True
    txtLimitReason.Text = lblLimitReason.Tag
    Call SetEnabled(False)
End Sub

Private Sub cmdPlanSet_Click()
    Dim lngPlanNo As Long

    If picPlanList.Visible Then
        lngPlanNo = cboPlanList.ItemData(cboPlanList.ListIndex)
        Call frmRunLimitPlanManage.ShowMe(lngPlanNo)
        Call FillPlanList(Split(tvwModuleTree.SelectedItem.Key, "_")(3), lngPlanNo)
    Else
        Call frmRunLimitPlanManage.ShowMe
    End If
End Sub

Private Sub cmdSave_Click()
'����ģ��ķ���ѡ�񣬲���ѡ���Լ���ʱԭ�����Ϣ
    On Error GoTo errh
    If cboPlanList.Text = "[�޷�������]" Then
        Call ExecuteProcedure("Zl_ZlRunLimitSet_Update(" & Split(tvwModuleTree.SelectedItem.Key, "_")(3) & _
                                ",Null)", "����ģ�鷽����Ϣ")
    Else
        If InStr(txtLimitReason.Text, "'") > 0 Then
            MsgBox "����ʱԭ���к��е����ţ���������д��", vbInformation, gstrSysName
            txtLimitReason.SetFocus
            Exit Sub
        ElseIf StrIsValid(txtLimitReason.Text, 250) = False Then
            txtLimitReason.SetFocus
            Exit Sub
        Else
            Call ExecuteProcedure("Zl_ZlRunLimitSet_Update(" & Split(tvwModuleTree.SelectedItem.Key, "_")(3) & _
                                "," & cboPlanList.ItemData(cboPlanList.ListIndex) & "," & IIf(optRemind(1).value, 1, 0) & _
                                ",'" & txtLimitReason.Text & "')", "����ģ�鷽����Ϣ")
        End If
    End If
    MsgBox "����ɹ���", vbInformation, gstrSysName
    Call SetEnabled(False)
    lblSelectPlan.Tag = cboPlanList.ListIndex
    lblOption.Tag = IIf(optRemind(1).value, 1, 0)
    lblLimitReason.Tag = txtLimitReason.Text
    Exit Sub
errh:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '��ֹ���뵥����
    If InStr("'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    '��tabControl�ؼ����г�ʼ��
    Call InitTabControl
    
    Call FormatVsfPlan

    '�������
    Call FillProgFunc
End Sub

'==============================================================================
'=���ܣ� ��ʼTab�ؼ�
'==============================================================================
Private Function InitTabControl() As Boolean
    Dim objTabItem As TabControlItem
    
    On Error GoTo errh
    With tbcPage
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .ShowIcons = True
            .OneNoteColors = True
            .DisableLunaColors = True
        End With
        '��һҳ
        Set objTabItem = .InsertItem(0, "Ԥ�跽��", picPlanDetail.hwnd, 0)
    End With

    InitTabControl = True

    Exit Function
errh:
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Function

Private Sub FormatVsfPlan()
    '�����Ҳ෽��չʾ����ʽ
    With vsfPlanDetail
        .Cell(flexcpPicture, 0, 0) = imgPlanDetail.ListImages(1).Picture
        .GridLines = flexGridNone
        .rowHeight(PD_����) = vsfTitleHeight
        .rowHeight(PD_������) = vsfRowHeight
        .rowHeight(PD_����һ) = vsfRowHeight
        .rowHeight(PD_���ڶ�) = vsfRowHeight
        .rowHeight(PD_������) = vsfRowHeight
        .rowHeight(PD_������) = vsfRowHeight
        .rowHeight(PD_������) = vsfRowHeight
        .rowHeight(PD_������) = vsfRowHeight
    End With
End Sub

Private Sub FillProgFunc()
    '������Ϸ�ģ�鹦�����νṹ
    Dim strSql As String
    Dim objNode As Node
    Dim rsTemp As ADODB.Recordset
    Dim lngSystemNo As Long
    Dim strModuleNo As String
    
    On Error GoTo errh
    strSql = "Select 0 ϵͳ, a.ģ��, b.���� ģ������, a.���, a.����, a.����ѡ��, a.�������, '������������' ϵͳ����" & vbNewLine & _
            "From ZlRunLimitSet A, zlSvrTools B" & vbNewLine & _
            "Where a.ģ�� = b.��� And a.ϵͳ Is Null" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select a.ϵͳ, a.ģ��, c.���� ģ������, a.���, a.����, a.����ѡ��, a.�������, b.���� ϵͳ����" & vbNewLine & _
            "From ZlRunLimitSet A, zlSystems B, zlPrograms C" & vbNewLine & _
            "Where a.ϵͳ = b.��� And a.ģ�� = c.���"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSql, Me.Caption)
    '������νṹ
    rsTemp.Sort = "ϵͳ,ģ��,���"
    lngSystemNo = -1
    With rsTemp
        Do While Not .EOF
            If !ϵͳ <> lngSystemNo Then
                Set objNode = tvwModuleTree.Nodes.Add(, , "S_" & !ϵͳ, !ϵͳ����, "system")
                objNode.Expanded = True
                lngSystemNo = !ϵͳ
            End If
            If !ϵͳ & "_" & !ģ�� <> strModuleNo Then
                Set objNode = tvwModuleTree.Nodes.Add("S_" & !ϵͳ, tvwChild, "M_" & !ϵͳ & "_" & !ģ��, !ģ������, "program")
                objNode.Expanded = True
                strModuleNo = !ϵͳ & "_" & !ģ��
            End If
            Call tvwModuleTree.Nodes.Add("M_" & !ϵͳ & "_" & !ģ��, tvwChild, "F_" & !ϵͳ & "_" & !ģ�� & "_" & !���, !����, "function")
            .MoveNext
        Loop
        If .RecordCount > 0 Then
            tvwModuleTree.Nodes(1).Child.Child.Selected = True
            Call tvwModuleTree_NodeClick(tvwModuleTree.SelectedItem)
        End If
    End With
    Exit Sub
errh:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    frmTopMidSplit.Width = Me.ScaleWidth
    fraMidButtomSplit.Width = Me.ScaleWidth
    fraMidButtomSplit.Top = Me.ScaleHeight
    picLeft.Height = fraMidButtomSplit.Top - picLeft.Top - 50
    picRight.Left = picLeft.Left + picLeft.Width + 50
    picRight.Height = picLeft.Height
    picLeftOrRight.Height = picLeft.Height
    picLeftOrRight.Left = picRight.Left - picLeftOrRight.Width
    picRight.Width = Me.ScaleWidth - picRight.Left - 30
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrLastPlan = ""
    Set mcolPlanDescription = Nothing
End Sub

Private Sub optRemind_Click(Index As Integer)
    If Val(cboPlanList.Tag) = 0 Then
        Call SetEnabled(True)
    End If
End Sub

Private Sub picLeftOrRight_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    On Error Resume Next
    If Button = 1 Then
        If picLeftOrRight.Left <= 3000 And X < 0 Then Exit Sub
        If picLeftOrRight.Left >= 9000 And X > 0 Then Exit Sub
        picLeftOrRight.Left = picLeftOrRight.Left + X
        picLeft.Width = picLeft.Width + X
        picRight.Left = picRight.Left + X
        picRight.Width = picRight.Width - X
    End If
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub picTopOrButtom_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 1 Then
        If picTopOrButtom.Top >= 6300 And y > 0 Then Exit Sub
        If picTopOrButtom.Top <= 3000 And y < 0 Then Exit Sub
        picTopOrButtom.Top = picTopOrButtom.Top + y
        tvwModuleTree.Height = tvwModuleTree.Height + y
        picPlanList.Height = picPlanList.Height - y
        picPlanList.Top = picPlanList.Top + y
    End If
End Sub

Private Sub picLeft_Resize()
    On Error Resume Next
    tvwModuleTree.Width = picLeft.Width - tvwModuleTree.Left
    If picPlanList.Visible Then
        tvwModuleTree.Height = picLeft.Height - picPlanList.Height - 45
    Else
        tvwModuleTree.Height = picLeft.Height
    End If
    picPlanList.Top = tvwModuleTree.Top + tvwModuleTree.Height + 45
    picPlanList.Width = tvwModuleTree.Width
    picTopOrButtom.Width = tvwModuleTree.Width
    picTopOrButtom.Top = picPlanList.Top - 80
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub picPlanDetail_Resize()
    On Error Resume Next
    vsfPlanDetail.Width = picPlanDetail.Width
    vsfPlanDetail.Height = picPlanDetail.Height
    Call AdjustFormDisplay
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub AdjustFormDisplay()
    With vsfPlanDetail
        .Select 0, 0, .Rows - 1, .Cols - 1
        .CellBorder &HE9D2A5, 1, 0, 1, 2, 2, 2
        .Cell(flexcpBackColor, PD_����, PDT_����, 0, .Cols - 1) = vsfTitleBackColor
        .Cell(flexcpBackColor, PD_����, PDT_����, .Rows - 1, 0) = vsfTitleBackColor
        .Cell(flexcpBackColor, PD_������, PDT_ʱ���1, PD_������, .Cols - 1) = vsfContentBackColor
        .Cell(flexcpBackColor, PD_���ڶ�, PDT_ʱ���1, PD_���ڶ�, .Cols - 1) = vsfContentBackColor
        .Cell(flexcpBackColor, PD_������, PDT_ʱ���1, PD_������, .Cols - 1) = vsfContentBackColor
        .Cell(flexcpBackColor, PD_������, PDT_ʱ���1, PD_������, .Cols - 1) = vsfContentBackColor
    End With
End Sub

Private Sub picRight_Resize()
    tbcPage.Width = picRight.Width
    tbcPage.Height = picRight.Height
End Sub

'���ģ�鹦�ܣ����·������б��Զ�ѡ�������õķ�������û�������õķ�������ѡ���һ��
Private Sub tvwModuleTree_NodeClick(ByVal Node As MSComctlLib.Node)
    If tvwModuleTree.Tag <> "" Then
        tvwModuleTree.Nodes(tvwModuleTree.Tag).BackColor = &H80000005
        tvwModuleTree.Nodes(tvwModuleTree.Tag).ForeColor = &H80000012
    End If
    Node.BackColor = &H8000000D
    Node.ForeColor = &H80000005
    tvwModuleTree.Tag = Node.Key
    If Mid(Node.Key, 1, 1) = "F" Then
        picRight.Visible = True
        tvwModuleTree.Height = picLeft.Height - picPlanList.Height - 45
        picPlanList.Top = tvwModuleTree.Top + tvwModuleTree.Height + 45
        picPlanList.Visible = True
        Call FillPlanList(Split(Node.Key, "_")(3))
        Call SetEnabled(False)
    Else
        Call ClearPlanDetail
        picRight.Visible = False
        picPlanList.Visible = False
        tvwModuleTree.Height = picLeft.Height
    End If
End Sub

Private Sub FillPlanList(ByVal lngFuncNo As Long, Optional ByVal lngPlanNo As Long)
'������·������б�
'lngFuncNo:���ܱ��
'lngPlanNo:������ţ�����ñ�Ŵ��ڣ����Զ�����λ���÷�����
    Dim strSql As String
    Dim i As Long
    
    On Error GoTo errh
    strSql = "Select a.���, a.����, a.����, a.�Ƿ�����, b.����ѡ��, b.��ʱԭ��, Decode(a.���, b.�������, 1, 0) ����" & vbNewLine & _
            "From ZlRunLimit A, ZlRunLimitSet B" & vbNewLine & _
            "Where b.��� = [1]" & vbNewLine & _
            "Order By a.���"
    Set mrsPlanList = gclsBase.OpenSQLRecord(gcnOracle, strSql, Me.Caption, lngFuncNo)
    
    Set mcolPlanDescription = New Collection
    cboPlanList.Clear
    cboPlanList.addItem "[�޷�������]"
    With mrsPlanList
        'cboPlanList.Tag���������жϸ�ģ�鹦���Ƿ������õķ�����
        '�Լ�����cboPlanList_Click��optRemind_Click��optProhibited_Click��txtLimitReason_Change����ʽ����
        cboPlanList.Tag = 1
        optRemind(Nvl(!����ѡ��, 1)).value = True
        lblOption.Tag = Nvl(!����ѡ��, 1)  '��¼��ʼ״̬�����ڵ��ȡ����ť��ָ�����
        txtLimitReason.Text = !��ʱԭ�� & ""
        lblLimitReason.Tag = txtLimitReason.Text   '��¼��ʼ״̬
        .Filter = "�Ƿ����� = 1"
        cboPlanList.Tag = 0
        Do While Not .EOF
            cboPlanList.addItem !����
            cboPlanList.ItemData(cboPlanList.NewIndex) = Val(!���)
            mcolPlanDescription.Add !���� & "", "K_" & !���
            If !���� = 1 And lngPlanNo = 0 Or lngPlanNo = !��� Then
                cboPlanList.Tag = 1
                cboPlanList.ListIndex = cboPlanList.NewIndex
            End If
            .MoveNext
        Loop
        If Val(cboPlanList.Tag) = 0 Then
            cboPlanList.Tag = 1
            cboPlanList.ListIndex = 0
        End If
        lblSelectPlan.Tag = cboPlanList.ListIndex   '��¼��ʼ״̬
        Call FillPlanDetail(cboPlanList.ItemData(cboPlanList.ListIndex))
        cboPlanList.Tag = 0
    End With
    Exit Sub
errh:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub FillPlanDetail(ByVal lngPlanNo As Long)
'�����ϸ������Ϣ
'lngPlanNo = �������
    Dim j As Long  '��ʾʱ���
    Dim lngLastWeekNo As Long
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errh
        '���ϵķ�����Ϣ���
        Call ClearPlanDetail
        
        '����·���
        strSql = "Select ����, To_Char(��ʼʱ��, 'HH24:MI:SS') ��ʼʱ��, To_Char(����ʱ��, 'HH24:MI:SS') ����ʱ��" & vbNewLine & _
                "From ZlRunLimitTime" & vbNewLine & _
                "Where ���� = [1]" & vbNewLine & _
                "Order By ����, ��ʼʱ��"
        Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSql, Me.Caption, lngPlanNo)
        With rsTemp
            Do While Not .EOF
                If !���� = lngLastWeekNo Then
                    j = j + 1
                    If j + 2 > vsfPlanDetail.Cols Then
                        vsfPlanDetail.Cols = j + 2
                        vsfPlanDetail.ColWidth(j) = vsfPlanDetail.ColWidth(PDT_ʱ���1)
                        vsfPlanDetail.TextMatrix(0, j) = "ʱ���" & j
                        vsfPlanDetail.ColAlignment(j) = flexAlignCenterCenter
                        Call AdjustFormDisplay
                    End If
                Else
                    j = 1
                End If
                vsfPlanDetail.TextMatrix(!���� + 1, j) = "�� " & !��ʼʱ�� & vbNewLine & vbNewLine & "ֹ " & !����ʱ��
                lngLastWeekNo = !����
                .MoveNext
            Loop
        End With
        If lngPlanNo > 0 Then
            vsfPlanDetail.ToolTipText = mcolPlanDescription("K_" & lngPlanNo)
        Else
            vsfPlanDetail.ToolTipText = ""
        End If
    Exit Sub
errh:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

'���ϵķ�����Ϣ���
Private Sub ClearPlanDetail()
    Dim i As Long
    
    With vsfPlanDetail
        .Cols = 3
        .TextMatrix(0, PDT_ʱ�����չ) = ""
        For i = PD_������ To PD_������
            .TextMatrix(i, PDT_ʱ���1) = ""
            .TextMatrix(i, PDT_ʱ�����չ) = ""
        Next
        Call AdjustFormDisplay
    End With
End Sub

Private Sub txtLimitReason_Change()
    If Val(cboPlanList.Tag) = 0 Then
        Call SetEnabled(True)
    End If
End Sub

Private Sub vsfPlanDetail_DblClick()
    Call cmdPlanSet_Click
End Sub

Private Sub SetEnabled(ByVal blnEnabled As Boolean)
    cmdSave.Enabled = blnEnabled
    cmdCancel.Enabled = blnEnabled
End Sub
