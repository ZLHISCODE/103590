VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmQueryUnVerify 
   Caption         =   "δ��˵��ݲ�ѯ"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11415
   Icon            =   "frmQueryUnVerify.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   11415
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picData 
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   120
      ScaleHeight     =   5895
      ScaleWidth      =   11055
      TabIndex        =   9
      Top             =   1320
      Width           =   11055
      Begin VB.Frame fraLineH1 
         Height          =   50
         Left            =   0
         TabIndex        =   12
         Top             =   4320
         Width           =   3405
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   2500
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   9375
         _cx             =   16536
         _cy             =   4410
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
         BackColorSel    =   16769992
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmQueryUnVerify.frx":076A
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
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
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   0   'False
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
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   1245
         Left            =   0
         TabIndex        =   11
         Top             =   4560
         Width           =   11055
         _cx             =   19500
         _cy             =   2196
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
         BackColorSel    =   16769992
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmQueryUnVerify.frx":082A
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
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
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   0   'False
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
   Begin VB.PictureBox picCondition 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      ScaleHeight     =   615
      ScaleWidth      =   11175
      TabIndex        =   0
      Top             =   240
      Width           =   11175
      Begin VB.TextBox TxtҩƷ 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   300
         Left            =   6240
         MaxLength       =   50
         ScrollBars      =   3  'Both
         TabIndex        =   6
         Top             =   120
         Width           =   3255
      End
      Begin VB.CommandButton cmd��ѯ 
         Caption         =   "��ѯ(&S)"
         Height          =   350
         Left            =   9960
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton CmdҩƷ 
         Caption         =   "��"
         Enabled         =   0   'False
         Height          =   300
         Left            =   9480
         TabIndex        =   7
         Top             =   120
         Width           =   255
      End
      Begin VB.CheckBox chkDrug 
         BackColor       =   &H80000003&
         Caption         =   "ҩƷ"
         Height          =   255
         Left            =   5520
         TabIndex        =   5
         Top             =   143
         Width           =   735
      End
      Begin VB.ComboBox cboTime 
         Height          =   300
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   120
         Width           =   2040
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   600
         TabIndex        =   1
         Text            =   "cboStock"
         Top             =   120
         Width           =   1800
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "ʱ�䷶Χ"
         Height          =   180
         Left            =   2520
         TabIndex        =   4
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblStock 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "�ⷿ"
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   180
         Width           =   360
      End
   End
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      Caption         =   "ע�⣺��ɫ��ʾָ��ҩƷ�������⣡"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   240
      TabIndex        =   13
      Top             =   960
      Width           =   2880
   End
   Begin XtremeCommandBars.ImageManager imgPublic 
      Left            =   600
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmQueryUnVerify.frx":08ED
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmQueryUnVerify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnLoadDate As Boolean '�����Ƿ������
Private mintChoose���� As Byte          '0-�ۼ۵�λ;1-���ﵥλ;2-ҩ�ⵥλ;3-סԺ��λ
Private mintNumberDigit As Integer

Private Sub GetData()
    
End Sub

Public Sub ShowCard(FrmMain As Form, ByVal cboStcokMain As ComboBox, ByVal intChoose���� As Byte, ByVal intNumberDigit As Integer)
    Dim i As Integer
    
    cboStock.Clear
    For i = 1 To cboStcokMain.ListCount - 1 '�ų���һ�����пⷿ
        cboStock.AddItem cboStcokMain.List(i)
        cboStock.ItemData(cboStock.NewIndex) = cboStcokMain.ItemData(i)
    Next
    
    If cboStock.ListCount > 0 Then
        cboStock.ListIndex = IIf(cboStcokMain.ListIndex - 1 >= 0, cboStcokMain.ListIndex - 1, 0)
    End If
    
    mintChoose���� = intChoose����
    mintNumberDigit = intNumberDigit
    
    Me.Show vbModal, FrmMain
End Sub

Private Sub InitComandBars()
    '��ʼ���˵�������ȫ���˵����������������˵���
    Dim cbrControlMain As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim rsData As ADODB.Recordset
    Dim i As Integer
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsMain.VisualTheme = xtpThemeOffice2003

    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    
    Me.cbsMain.EnableCustomization False
    Me.cbsMain.Icons = Me.imgPublic.Icons
    
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Preview, "Ԥ��")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Print, "��ӡ")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Exit, "�˳�")
        cbrControlMain.BeginGroup = True
    End With
    
    For Each cbrControlMain In cbrToolBar.Controls
        cbrControlMain.Style = xtpButtonIconAndCaption
    Next
    cbsMain.Item(1).Visible = False
End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.id
        Case mconMenu_File_Preview
            subPrint 2
        Case mconMenu_File_Print
            subPrint 1
        Case mconMenu_File_Exit
            Unload Me
    End Select
    
End Sub

Private Sub cbsFilePreView()
    '��ӡԤ��
    vsfList.Redraw = flexRDNone
    subPrint 2
    vsfList.Redraw = flexRDDirect
    vsfList.Col = 0
    vsfList.ColSel = vsfList.Cols - 1
End Sub

Private Sub cbsFilePrint()
    '��ӡ
    vsfList.Redraw = flexRDNone
    subPrint 1
    vsfList.Redraw = flexRDDirect
    vsfList.Col = 0
    vsfList.ColSel = vsfList.Cols - 1
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    Me.picCondition.Move lngLeft, lngTop, lngRight - lngLeft
    
    Me.lblMsg.Move 0, Me.ScaleHeight - lblMsg.Height - 50, lblMsg.Width, lblMsg.Height
    
    Me.picData.Move lngLeft, picCondition.Top + picCondition.Height + 50, lngRight - lngLeft, _
        Me.ScaleHeight - Me.picCondition.Top - Me.picCondition.Height - lblMsg.Height - 150
End Sub


Private Sub chkDrug_Click()
    TxtҩƷ.Enabled = IIf(chkDrug.Value = 1, True, False)
    CmdҩƷ.Enabled = IIf(chkDrug.Value = 1, True, False)
    
    TxtҩƷ.BackColor = IIf(TxtҩƷ.Enabled, &H80000005, &H80000004)
End Sub

Private Sub cmd��ѯ_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim intDay As Integer
    Dim lngStockid As Long
    Dim strSql As String
    
    On Error GoTo errHandle
    
    blnLoadDate = False
    
    lngStockid = Val(cboStock.ItemData(cboStock.ListIndex))
    
    Select Case cboTime.ListIndex
        Case 0 '��ʾ7����
            intDay = 7
        Case 1 '��ʾ1������
            intDay = 30
        Case 2 '��ʾ3������
            intDay = 90
        Case 3 '��ʾ������
            intDay = 183
        Case 4 '��ʾ1����
            intDay = 365
    End Select
    
    If chkDrug.Value = 1 Then '��ѡ��ҩƷ
        If Val(TxtҩƷ.Tag) <= 0 Then
            MsgBox "��ѡ��Ҫ��ѯ��ҩƷ��", vbInformation + vbOKOnly, gstrSysName
            TxtҩƷ.SetFocus
            Exit Sub
        End If
        
        '���ݴ������������ʾ��λ
'        Select Case mintChoose����
'            Case 1
'                strSql = ", Decode(Sign(a.���ϵ�� * Nvl(a.����, 1) * Nvl(a.ʵ������, 0)), -1, a.���ϵ�� * Nvl(a.����, 1) * Nvl(a.ʵ������, 0), 0) As ���� ,C.���㵥λ as ��λ"
'            Case 2
'                strSql = ", Decode(Sign(a.���ϵ�� * Nvl(a.����, 1) * Nvl(a.ʵ������, 0)), -1, a.���ϵ�� * Nvl(a.����, 1) * Nvl(a.ʵ������, 0), 0)/D.�����װ As ����  ,D.���ﵥλ as ��λ"
'            Case 3
'                strSql = ", Decode(Sign(a.���ϵ�� * Nvl(a.����, 1) * Nvl(a.ʵ������, 0)), -1, a.���ϵ�� * Nvl(a.����, 1) * Nvl(a.ʵ������, 0), 0)/D.ҩ���װ As ����  ,D.ҩ�ⵥλ as ��λ"
'            Case 4
'                strSql = ", Decode(Sign(a.���ϵ�� * Nvl(a.����, 1) * Nvl(a.ʵ������, 0)), -1, a.���ϵ�� * Nvl(a.����, 1) * Nvl(a.ʵ������, 0), 0)/D.סԺ��װ As ���� ,D.סԺ��λ as ��λ"
'        End Select
        Select Case mintChoose����
            Case 1
                strSql = ", a.���ϵ�� * Nvl(a.����, 1) * Nvl(a.ʵ������, 0) As ���� ,C.���㵥λ as ��λ"
            Case 2
                strSql = ", a.���ϵ�� * Nvl(a.����, 1) * Nvl(a.ʵ������, 0)/D.�����װ As ����  ,D.���ﵥλ as ��λ"
            Case 3
                strSql = ", a.���ϵ�� * Nvl(a.����, 1) * Nvl(a.ʵ������, 0)/D.ҩ���װ As ����  ,D.ҩ�ⵥλ as ��λ"
            Case 4
                strSql = ", a.���ϵ�� * Nvl(a.����, 1) * Nvl(a.ʵ������, 0)/D.סԺ��װ As ���� ,D.סԺ��λ as ��λ"
        End Select
        
        '��ѯָ��ҩƷ��δ��˵��ݼ�������������ռ�ÿ�������
        gstrSQL = "Select a.id, a.������, Count(Distinct NO) As ��������, Sum(a.����) As ʵ������,Max(a.��λ) as ��λ " & vbNewLine & _
                "From (Select e.id ,e.���� ������, a.No" & strSql & vbNewLine & _
                "       From ҩƷ�շ���¼ A, δ��ҩƷ��¼ B ,�շ���ĿĿ¼ C,ҩƷ��� D, ҩƷ������ E" & vbNewLine & _
                "       Where a.Id = b.�շ�id And A.ҩƷid = C.id And C.id = d.ҩƷid And a.������id = e.Id " & IIf(lngStockid = 0, "", " And b.�ⷿid = [1] ") & " And b.ҩƷid = [2] " & IIf(cboTime.ListIndex = 5, "", " And a.�������� > sysdate - [3]") & " And Not Exists" & vbNewLine & _
                "        (Select 1 From ҩƷ�շ���¼ C Where b.�շ�id = c.Id And Nvl(c.��ҩ��ʽ, 0) = -1 And c.���� In (8, 9, 10))) A" & vbNewLine & _
                "Group By a.������,a.id"
    Else
        '��ѯδ��˵��ݼ���������
        gstrSQL = "Select e.id ,e.���� ������, Count(Distinct a.No) As ��������" & vbNewLine & _
                "From ҩƷ�շ���¼ A, δ��ҩƷ��¼ B, ҩƷ������ E" & vbNewLine & _
                "Where a.Id = b.�շ�id And a.������id = e.Id " & IIf(lngStockid = 0, "", " And b.�ⷿid = [1] ") & "" & IIf(cboTime.ListIndex = 5, "", " And a.�������� > sysdate - [3] ") & vbNewLine & _
                "Group By e.id ,e.����"
    End If
    
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "", lngStockid, Val(TxtҩƷ.Tag), intDay)
    
    With vsfList
        .rows = 1
        .rows = .rows + 1
        .Row = .rows - 1
        Do While Not rsTemp.EOF
            .TextMatrix(.Row, .ColIndex("������id")) = rsTemp!id
            .TextMatrix(.Row, .ColIndex("������")) = rsTemp!������
            .TextMatrix(.Row, .ColIndex("��������")) = rsTemp!��������
            If chkDrug.Value = 1 Then
                '��ɫ�������
                If rsTemp!ʵ������ < 0 Then
                    .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &HFF '�����ɫ
                Else
                    .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &H80000012
                End If
                
                .TextMatrix(.Row, .ColIndex("ʵ������")) = zlStr.FormatEx(Abs(rsTemp!ʵ������), mintNumberDigit, False, True) & rsTemp!��λ
                .Cell(flexcpFontBold, 1, .ColIndex("ʵ������"), .rows - 1, .ColIndex("ʵ������")) = True
            End If
            
            .rows = .rows + 1
            .Row = .Row + 1
            
            rsTemp.MoveNext
        Loop
        
        If Trim(.TextMatrix(.rows - 1, 0)) = "" Then .RemoveItem (.rows - 1) 'ɾ�����Ŀ���
    End With
    
    colHidden
    
    blnLoadDate = True
    
    vsfList_EnterCell '������ϸ
    
    vsfList.SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub CmdҩƷ_Click()
    Dim RecReturn As Recordset
    
    Call SetSelectorRS(1, "", cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , , , True)
    
'    Set RecReturn = FrmҩƷѡ����.ShowME(Me, 1, 0, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex))
    Set RecReturn = frmSelector.ShowME(Me, 0, 1, , , , cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , , , , 2, False)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gintҩƷ������ʾ = 1 Then
        TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & IIf(IsNull(RecReturn!��Ʒ��), RecReturn!ͨ����, RecReturn!��Ʒ��)
    Else
        TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & RecReturn!ͨ����
    End If
    TxtҩƷ.Tag = RecReturn!ҩƷid
    
    cmd��ѯ.SetFocus
End Sub

Private Sub Form_Load()
    
    With cboTime
        .Clear
        
        .AddItem "0-��ʾ7����"
        .AddItem "1-��ʾ1������"
        .AddItem "2-��ʾ3������"
        .AddItem "3-��ʾ������"
        .AddItem "4-��ʾ1����"
        .AddItem "5-��ʾ����"
        
        .ListIndex = 0
    End With
    
    Call InitComandBars
    
    colHidden
End Sub


Private Sub colHidden()
    '����ѡ������������
    With vsfList
        .colHidden(.ColIndex("ʵ������")) = chkDrug.Value = 0 'δѡ��ҩƷ�����ء�ռ�ÿ�����������
        
        .ColWidth(.ColIndex("ʵ������")) = IIf(.colHidden(.ColIndex("ʵ������")), 0, 900)

    End With
    With vsfDetail
        .colHidden(.ColIndex("����")) = chkDrug.Value = 0 'δѡ��ҩƷ�����ء���������
        
        .ColWidth(.ColIndex("����")) = IIf(.colHidden(.ColIndex("����")), 0, 1545)
    End With
End Sub

Private Sub fraLineH1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    With fraLineH1
        If .Top + y < 2000 Then Exit Sub
        If .Top + y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + y
    End With
    With vsfList
        .Height = fraLineH1.Top - .Top
    End With
    
    With vsfDetail
        .Top = fraLineH1.Top + fraLineH1.Height + 100
        .Height = ScaleHeight - .Top
    End With
    Me.Refresh
End Sub


Private Sub picData_Resize()
    On Error Resume Next
    
    With vsfList
        .Move 0, 0, picData.Width, 2500
    End With
    
    With fraLineH1
        .Move 0, vsfList.Top + vsfList.Height, picData.Width, fraLineH1.Height
    End With
    
    With vsfDetail
        .Move 0, fraLineH1.Top + fraLineH1.Height, picData.Width, picData.Height - fraLineH1.Top - fraLineH1.Height
    End With
End Sub


Private Sub TxtҩƷ_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim RecReturn As Recordset
    Dim strkey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(TxtҩƷ.Text) = "" Then Exit Sub
    sngLeft = Me.Left + picCondition.Left + TxtҩƷ.Left
    sngTop = Me.Top + picCondition.Top + TxtҩƷ.Top + TxtҩƷ.Height + Me.Height - Me.ScaleHeight '  50
    If sngTop + 3630 > Screen.Height Then
        sngTop = sngTop - TxtҩƷ.Height - 3630
    End If
    
    strkey = Trim(TxtҩƷ.Text)
    If Mid(strkey, 1, 1) = "[" Then
        If InStr(2, strkey, "]") <> 0 Then
            strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
        Else
            strkey = Mid(strkey, 2)
        End If
    End If
    
    Call SetSelectorRS(1, "", cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , , True)
    
'    Set RecReturn = FrmҩƷ��ѡѡ����.ShowME(Me, 1, , mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), strkey, sngLeft, sngTop)
    Set RecReturn = frmSelector.ShowME(Me, 1, 1, strkey, sngLeft, sngTop, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , , , , 2, False)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gintҩƷ������ʾ = 1 Then
        TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & IIf(IsNull(RecReturn!��Ʒ��), RecReturn!ͨ����, RecReturn!��Ʒ��)
    Else
        TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & RecReturn!ͨ����
    End If
    TxtҩƷ.Tag = RecReturn!ҩƷid
    
    cmd��ѯ.SetFocus
    
End Sub

Private Sub vsfList_EnterCell()
    Dim rsTemp As New ADODB.Recordset
    Dim intDay As Integer
    Dim lngStockid As Long
    Dim int������id As Integer
    Dim strSql As String
    
    On Error GoTo errHandle
    
    If Not blnLoadDate Then Exit Sub
    
    Select Case cboTime.ListIndex
        Case 0 '��ʾ7����
            intDay = 7
        Case 1 '��ʾ1������
            intDay = 30
        Case 2 '��ʾ3������
            intDay = 90
        Case 3 '��ʾ������
            intDay = 183
        Case 4 '��ʾ1����
            intDay = 365
    End Select
    
    lngStockid = Val(cboStock.ItemData(cboStock.ListIndex))
    int������id = Val(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("������id")))
    
    If chkDrug.Value = 1 Then '��ѡ��ҩƷ
        '���ݴ������������ʾ��λ
        Select Case mintChoose����
            Case 1
                strSql = ", a.���ϵ�� * Nvl(a.����, 1) * Nvl(a.ʵ������, 0) As ���� ,C.���㵥λ as ��λ"
            Case 2
                strSql = ", a.���ϵ�� * Nvl(a.����, 1) * Nvl(a.ʵ������, 0)/D.�����װ As ����  ,D.���ﵥλ as ��λ"
            Case 3
                strSql = ", a.���ϵ�� * Nvl(a.����, 1) * Nvl(a.ʵ������, 0)/D.ҩ���װ As ����  ,D.ҩ�ⵥλ as ��λ"
            Case 4
                strSql = ", a.���ϵ�� * Nvl(a.����, 1) * Nvl(a.ʵ������, 0)/D.סԺ��װ As ���� ,D.סԺ��λ as ��λ"
        End Select
        
        gstrSQL = "Select a.No, a.������, a.��������, a.ժҪ, Sum(a.����) ����,Max(a.��λ) as ��λ" & vbNewLine & _
                "From (Select a.No" & strSql & vbNewLine & _
                "       , a.������, a.��������, a.ժҪ" & vbNewLine & _
                "       From ҩƷ�շ���¼ A, δ��ҩƷ��¼ B,�շ���ĿĿ¼ C,ҩƷ��� D" & vbNewLine & _
                "       Where a.Id = b.�շ�id And A.ҩƷid = C.id And C.id = d.ҩƷid And b.�ⷿid = [1] And a.������id = [2] " & IIf(cboTime.ListIndex = 5, "", " And a.�������� > sysdate - [3]") & " And a.ҩƷid = [4] And Not Exists" & vbNewLine & _
                "        (Select 1 From ҩƷ�շ���¼ C Where b.�շ�id = c.Id And Nvl(c.��ҩ��ʽ, 0) = -1 And c.���� In (8, 9, 10))) A" & vbNewLine & _
                "Group By a.No, a.������, a.��������, a.ժҪ, a.����"


    Else
        gstrSQL = "Select Distinct a.No, a.������, a.��������, a.ժҪ" & vbNewLine & _
                "From ҩƷ�շ���¼ A, δ��ҩƷ��¼ B" & vbNewLine & _
                "Where a.Id = b.�շ�id And b.�ⷿid = [1] And a.������id = [2]" & IIf(cboTime.ListIndex = 5, "", " And a.�������� > sysdate - [3]") & vbNewLine & _
                "Order By NO"
    End If
    
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "", lngStockid, int������id, intDay, Val(TxtҩƷ.Tag))
    
    With vsfDetail
        .rows = 1
        .rows = .rows + 1
        .Row = .rows - 1
        Do While Not rsTemp.EOF
            .TextMatrix(.Row, .ColIndex("No")) = rsTemp!NO
            .TextMatrix(.Row, .ColIndex("������")) = rsTemp!������
            .TextMatrix(.Row, .ColIndex("��������")) = rsTemp!��������
            .TextMatrix(.Row, .ColIndex("ժҪ")) = "" & rsTemp!ժҪ
            If chkDrug.Value = 1 Then
                '��ɫ�������
                If rsTemp!���� < 0 Then
                    .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &HFF '�����ɫ
                Else
                    .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &H80000012
                End If
                .TextMatrix(.Row, .ColIndex("����")) = zlStr.FormatEx(Abs(rsTemp!����), mintNumberDigit, False, True) & rsTemp!��λ: .Cell(flexcpFontBold, 1, .ColIndex("����"), .rows - 1, .ColIndex("����")) = True
            End If
            
            .rows = .rows + 1
            .Row = .Row + 1
            
            rsTemp.MoveNext
        Loop
        
        If Trim(.TextMatrix(.rows - 1, 0)) = "" Then .RemoveItem (.rows - 1) 'ɾ�����Ŀ���
    End With
    
    colHidden
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub subPrint(bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow

    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = "δ��˵���"
        
    objRow.Add "���ţ�" & cboStock.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
        
    objRow.Add "��ӡ��:" & UserInfo.�û�����
    objRow.Add "��ӡ����:" & Format(Sys.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    If Me.ActiveControl Is vsfDetail Then
        Set objPrint.Body = vsfDetail
    Else
        Set objPrint.Body = vsfList
    End If
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

