VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmRunLimitManage 
   BackColor       =   &H80000005&
   Caption         =   "������ʱ����"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12315
   ControlBox      =   0   'False
   Icon            =   "frmRunLimitManage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmRunLimitManage.frx":6852
   ScaleHeight     =   8655
   ScaleWidth      =   12315
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTop 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   0
      ScaleHeight     =   3015
      ScaleWidth      =   12255
      TabIndex        =   4
      Top             =   1290
      Width           =   12255
      Begin VSFlex8Ctl.VSFlexGrid vsfModuleList 
         Height          =   2175
         Left            =   45
         TabIndex        =   9
         Top             =   0
         Width           =   11955
         _cx             =   21087
         _cy             =   3836
         Appearance      =   1
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   260
         RowHeightMax    =   260
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmRunLimitManage.frx":6D4B
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   3
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   1
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
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
   Begin VB.CommandButton cmdPlanSet 
      Caption         =   "��������(&S)"
      Height          =   350
      Left            =   10995
      TabIndex        =   8
      Top             =   690
      Width           =   1200
   End
   Begin MSComctlLib.ImageList imgPlanDetail 
      Left            =   11625
      Top             =   0
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
            Picture         =   "frmRunLimitManage.frx":6E38
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBottom 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   3885
      Left            =   0
      ScaleHeight     =   3885
      ScaleWidth      =   12255
      TabIndex        =   2
      Top             =   4590
      Width           =   12255
      Begin VB.PictureBox picPlanDetail 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   0
         ScaleHeight     =   2415
         ScaleWidth      =   8130
         TabIndex        =   5
         Top             =   1035
         Width           =   8130
         Begin VSFlex8Ctl.VSFlexGrid vsfPlanDetail 
            Height          =   2115
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   7905
            _cx             =   13944
            _cy             =   3731
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
            Rows            =   3
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmRunLimitManage.frx":B10E
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
         Left            =   45
         TabIndex        =   3
         Top             =   0
         Width           =   1755
         _Version        =   589884
         _ExtentX        =   3096
         _ExtentY        =   1693
         _StockProps     =   64
      End
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
            Picture         =   "frmRunLimitManage.frx":B200
            Key             =   "system"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunLimitManage.frx":11A62
            Key             =   "function"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunLimitManage.frx":182C4
            Key             =   "program"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTopBottom 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   45
      MousePointer    =   7  'Size N S
      ScaleHeight     =   135
      ScaleWidth      =   4515
      TabIndex        =   10
      Top             =   4365
      Width           =   4515
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   $"frmRunLimitManage.frx":1EB26
      Height          =   360
      Left            =   1050
      TabIndex        =   7
      Top             =   675
      Width           =   7740
   End
   Begin VB.Image imgDescription 
      Height          =   720
      Left            =   150
      Picture         =   "frmRunLimitManage.frx":1EBD8
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
Private mrsPlanList As ADODB.Recordset
Private Const vsfTitleBackColor = &HF0E5BD  '�������ݱ����ⱳ����ɫ
Private Const vsfContentBackColor = &HFFFAE4 '�������ݱ�����ݲ���ǳɫ����ɫ
Private Const vsfTitleHeight = 500
Private Const vsfRowHeight = 1000
Private Enum ModuleList
    ML_��� = 0
    ML_ϵͳ = 1
    ML_ģ�� = 2
    ML_���� = 3
    ML_��ʱ���� = 4
    ML_����˵�� = 5
    ML_����ѡ�� = 6
    ML_��ʱԭ�� = 7
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

Private Sub cmdPlanSet_Click()
    Call frmRunLimitPlanManage.ShowMe(vsfPlanDetail.Tag)
    Call FillModuleData(vsfModuleList.Row)
End Sub

Private Function SaveData(ByVal lngRow As Long) As Boolean
'����ģ��ķ���ѡ�񣬲���ѡ���Լ���ʱԭ�����Ϣ
    On Error GoTo errH
    If InStr(vsfModuleList.TextMatrix(lngRow, ML_��ʱԭ��), "'") > 0 Then
        MsgBox "����ʱԭ���к��е����ţ���������д��", vbInformation, gstrSysName
        Exit Function
    ElseIf LenB(StrConv(vsfModuleList.TextMatrix(lngRow, ML_��ʱԭ��), vbFromUnicode)) > 250 Then
        MsgBox "����ʱԭ�����ݲ��ܲ���125�����ֻ�250���ַ�����������д��"
        Exit Function
    Else
        Call ExecuteProcedure("Zl_ZlRunLimitSet_Update(" & vsfModuleList.TextMatrix(lngRow, ML_���) & _
                            "," & Val(vsfPlanDetail.Tag) & "," & IIf(vsfModuleList.TextMatrix(lngRow, ML_����ѡ��) = "��ֹ", 0, 1) & _
                            ",'" & vsfModuleList.TextMatrix(lngRow, ML_��ʱԭ��) & "')", "����ģ�鷽����Ϣ")
    End If
    vsfModuleList.Tag = vsfModuleList.TextMatrix(vsfModuleList.RowSel, ML_��ʱ����) & "_" & _
                        vsfModuleList.TextMatrix(vsfModuleList.RowSel, ML_����ѡ��) & "_" & _
                        vsfModuleList.TextMatrix(vsfModuleList.RowSel, ML_��ʱԭ��)
    SaveData = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    '��ֹ���뵥����
    If InStr("'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    picTopBottom.Top = Val(GetSetting("ZLSOFT", "����ģ��\������������\������ʱ����", "picTopBottom_Top", "5000"))
    '��tabControl�ؼ����г�ʼ��
    Call InitTabControl
    '��ʼ������ʽ
    Call FormatPlanDetail
    '�������
    Call FillModuleData
End Sub

'==============================================================================
'=���ܣ� ��ʼTab�ؼ�
'==============================================================================
Private Function InitTabControl() As Boolean
    Dim objTabItem As TabControlItem
    
    On Error GoTo errH
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
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Function

Private Sub FormatPlanDetail()
    '�����·�����չʾ����ʽ
    With vsfPlanDetail
        .Cell(flexcpPicture, 0, 0) = imgPlanDetail.ListImages(1).Picture
        .GridLines = flexGridNone
        .rowHeight(PD_����) = vsfTitleHeight
        .rowHeight(PDT_ʱ���1) = vsfRowHeight
    End With
End Sub

'����Ϸ�ģ�鹦���б��еķ���������ѡ��������
Private Sub FormatPlanList()
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strComboList As String
    
    On Error GoTo errH
    strSql = "Select ����, ���� From Zlrunlimit Where �Ƿ����� = 1 Order by ���"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSql, Me.Caption)
    With rsTemp
        Do While Not .EOF
            strComboList = strComboList & "|" & !����
            .MoveNext
        Loop
    End With
    vsfModuleList.ColComboList(ML_��ʱ����) = "[�޷�������]|" & strComboList
    vsfModuleList.ColComboList(ML_����ѡ��) = "����|��ֹ"
    Exit Sub
errH:
    MsgBox err.Description, vbInformation
End Sub

'����Ϸ�ģ�鹦�ܼ�����ʱ��Ϣ
Private Sub FillModuleData(Optional ByVal lngRow As Long)
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo errH
    Call FormatPlanList
    strSql = "Select 0 ϵͳ, a.ģ��, b.���� ģ������, a.���, a.����, a.����ѡ��, c.���� ����, '������������' ϵͳ����, a.��ʱԭ��" & vbNewLine & _
            "From Zlrunlimitset A, zlSvrTools B, Zlrunlimit C" & vbNewLine & _
            "Where a.ģ�� = b.��� And a.ϵͳ Is Null And a.������� = c.���(+)" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select a.ϵͳ, a.ģ��, c.���� ģ������, a.���, a.����, a.����ѡ��, d.���� ����, b.���� ϵͳ����, a.��ʱԭ��" & vbNewLine & _
            "From Zlrunlimitset A, zlSystems B, zlPrograms C, Zlrunlimit D" & vbNewLine & _
            "Where a.ϵͳ = b.��� And a.ģ�� = c.��� And a.������� = d.���(+)"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSql, Me.Caption)
    '���ģ�鹦���б�
    rsTemp.Sort = "ϵͳ,ģ��,���"
    With rsTemp
        vsfModuleList.Rows = .RecordCount + 1
        For i = 1 To .RecordCount
            vsfModuleList.TextMatrix(i, ML_���) = !���
            vsfModuleList.TextMatrix(i, ML_ϵͳ) = !ϵͳ����
            vsfModuleList.TextMatrix(i, ML_ģ��) = !ģ������
            vsfModuleList.TextMatrix(i, ML_����) = !����
            vsfModuleList.TextMatrix(i, ML_��ʱ����) = Nvl(!����, "[�޷�������]")
            vsfModuleList.TextMatrix(i, ML_����ѡ��) = IIf(!����ѡ�� = 0, "��ֹ", "����")
            vsfModuleList.TextMatrix(i, ML_��ʱԭ��) = !��ʱԭ��
            .MoveNext
        Next
        If .RecordCount > 0 Then
            vsfModuleList.MergeCol(ML_ϵͳ) = True
            vsfModuleList.MergeCol(ML_ģ��) = True
            If lngRow = 0 Then
                vsfModuleList.Row = 1
            Else
                vsfModuleList.Row = lngRow
            End If
            Call vsfModuleList_Click
        End If
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    cmdPlanSet.Left = Me.ScaleWidth - cmdPlanSet.Width - 200
    frmTopMidSplit.Width = Me.ScaleWidth
    picTop.Height = picTopBottom.Top - picTop.Top
    picTop.Width = Me.ScaleWidth - 60
    picTopBottom.Width = picTop.Width - 45
    picBottom.Top = picTop.Top + picTop.Height + 60
    picBottom.Width = picTop.Width
    picBottom.Height = Me.ScaleHeight - picBottom.Top
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "ZLSOFT", "����ģ��\������������\������ʱ����", "picTopBottom_Top", picTopBottom.Top
    mstrLastPlan = ""
End Sub

Private Sub picTopBottom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If picTopBottom.Top >= 9000 And Y > 0 Then Exit Sub
        If picTopBottom.Top <= 5000 And Y < 0 Then Exit Sub
        picTopBottom.Top = picTopBottom.Top + Y
        picTop.Height = picTopBottom.Top - picTop.Top
        picBottom.Top = picTop.Top + picTop.Height + 60
        picBottom.Height = Me.ScaleHeight - picBottom.Top
    End If
End Sub

Private Sub picTop_Resize()
    On Error Resume Next
    vsfModuleList.Width = picTop.Width
    vsfModuleList.Height = picTop.Height
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
        .CellBorder &HE9D2A5, 1, 2, 1, 0, 2, 2
        .Cell(flexcpBackColor, PDT_����, PD_����, .Rows - 1, 0) = vsfTitleBackColor
        .Cell(flexcpBackColor, PDT_����, PD_����, 0, .Cols - 1) = vsfTitleBackColor
        .Cell(flexcpBackColor, PDT_ʱ���1, PD_������, .Rows - 1, PD_������) = vsfContentBackColor
        .Cell(flexcpBackColor, PDT_ʱ���1, PD_���ڶ�, .Rows - 1, PD_���ڶ�) = vsfContentBackColor
        .Cell(flexcpBackColor, PDT_ʱ���1, PD_������, .Rows - 1, PD_������) = vsfContentBackColor
        .Cell(flexcpBackColor, PDT_ʱ���1, PD_������, .Rows - 1, PD_������) = vsfContentBackColor
        .rowHeight(.Rows - 1) = picBottom.Height - (.Rows - 1) * .rowHeight(PDT_ʱ���1) + 200
    End With
End Sub

Private Sub picBottom_Resize()
    tbcPage.Width = picBottom.Width
    tbcPage.Height = picBottom.Height
End Sub

Private Sub FillPlanDetail(ByVal strPlanName As String)
'�����ϸ������Ϣ
'strPlanName = ��������
    Dim j As Long  '��ʾʱ���
    Dim lngLastWeekNo As Long
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errH
        '���ϵķ�����Ϣ���
        Call ClearPlanDetail
        If strPlanName = "[�޷�������]" Then
            tbcPage.Item(0).Caption = "�޷���"
        Else
            tbcPage.Item(0).Caption = strPlanName
        End If
        mstrLastPlan = strPlanName
        '����·���
        strSql = "Select b.���, a.����, To_Char(a.��ʼʱ��, 'HH24:MI:SS') ��ʼʱ��, To_Char(a.����ʱ��, 'HH24:MI:SS') ����ʱ��, b.����" & vbNewLine & _
                "From Zlrunlimittime A, Zlrunlimit B" & vbNewLine & _
                "Where a.���� = b.��� And b.���� = [1]" & vbNewLine & _
                "Order By a.����, a.��ʼʱ��"
        Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSql, Me.Caption, strPlanName)
        With rsTemp
            If .RecordCount > 0 Then
                vsfModuleList.TextMatrix(vsfModuleList.RowSel, ML_����˵��) = !���� & ""
                vsfPlanDetail.Tag = !���
            Else
                vsfModuleList.TextMatrix(vsfModuleList.RowSel, ML_����˵��) = ""
                vsfPlanDetail.Tag = 1
            End If
            Do While Not .EOF
                If !���� = lngLastWeekNo Then
                    j = j + 1
                    If j + 2 > vsfPlanDetail.Rows Then
                        vsfPlanDetail.Rows = j + 2
                        vsfPlanDetail.rowHeight(j) = vsfPlanDetail.rowHeight(PDT_ʱ���1)
                        vsfPlanDetail.TextMatrix(j, 0) = "ʱ���" & j
                        vsfPlanDetail.ColAlignment(j) = flexAlignCenterCenter
                    End If
                Else
                    j = 1
                End If
                vsfPlanDetail.TextMatrix(j, !���� + 1) = "�� " & !��ʼʱ�� & vbNewLine & vbNewLine & "ֹ " & !����ʱ��
                lngLastWeekNo = !����
                .MoveNext
            Loop
            Call AdjustFormDisplay
        End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Sub

'���ϵķ�����Ϣ���
Private Sub ClearPlanDetail()
    Dim i As Long
    
    With vsfPlanDetail
        .Rows = 3
        .TextMatrix(PDT_ʱ�����չ, 0) = ""
        For i = PD_������ To PD_������
            .TextMatrix(PDT_ʱ���1, i) = ""
            .TextMatrix(PDT_ʱ�����չ, i) = ""
        Next
        Call AdjustFormDisplay
    End With
End Sub


Private Sub vsfModuleList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfModuleList
        If .Tag <> .TextMatrix(Row, ML_��ʱ����) & "_" & .TextMatrix(Row, ML_����ѡ��) & "_" & .TextMatrix(Row, ML_��ʱԭ��) Then
            Call SaveData(Row)
        End If
    End With
End Sub

Private Sub vsfModuleList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = ML_ϵͳ Or Col = ML_ģ�� Or Col = ML_���� Then Cancel = True
End Sub

'�����ĳһ��ģ�鹦��ʱ�������·�չʾ����ģ�鹦��ѡ�����ʱ��������ϸ��Ϣ
Private Sub vsfModuleList_Click()
    With vsfModuleList
        .Tag = .TextMatrix(.RowSel, ML_��ʱ����) & "_" & .TextMatrix(.RowSel, ML_����ѡ��) & "_" & .TextMatrix(.RowSel, ML_��ʱԭ��)
        If mstrLastPlan = .TextMatrix(.RowSel, ML_��ʱ����) And .MouseRow = .Row Then Exit Sub
        Call FillPlanDetail(.TextMatrix(.RowSel, ML_��ʱ����))
    End With
End Sub

Private Sub vsfModuleList_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
    If Col = ML_��ʱ���� Then
        If mstrLastPlan = vsfModuleList.EditText Then Exit Sub
        Call FillPlanDetail(vsfModuleList.EditText)
    End If
End Sub

Private Sub vsfModuleList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sinRight As Single
    Dim sinLeftPlan As Single, sinLeftReason As Single
    Dim strTip As String
    Dim lngRow As Long
    
    lngRow = Int(Y / 260)
    With vsfModuleList
        If lngRow > .Rows - 1 Or lngRow = 0 Then
            Call ShowTipInfo(.hwnd, "")
            Exit Sub
        End If
        sinLeftPlan = .ColWidth(ML_ϵͳ) + .ColWidth(ML_ģ��) + .ColWidth(ML_����)
        sinRight = .ColWidth(ML_ϵͳ) + .ColWidth(ML_ģ��) + .ColWidth(ML_����) + .ColWidth(ML_��ʱ����)
        sinLeftReason = .ColWidth(ML_ϵͳ) + .ColWidth(ML_ģ��) + .ColWidth(ML_����) + .ColWidth(ML_��ʱ����) + .ColWidth(ML_����ѡ��)
        If X >= sinLeftPlan And X <= sinRight Then
            strTip = .TextMatrix(lngRow, ML_����˵��)
        ElseIf X > sinLeftReason Then
            strTip = .TextMatrix(lngRow, ML_��ʱԭ��)
        Else
            strTip = ""
        End If
        Call ShowTipInfo(.hwnd, strTip, True)
    End With
End Sub

Private Sub vsfPlanDetail_DblClick()
    Call cmdPlanSet_Click
End Sub
