VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmClinicHolidayManage 
   BorderStyle     =   0  'None
   Caption         =   "�ڼ��չ���"
   ClientHeight    =   7845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox pic������� 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3195
      Left            =   6780
      ScaleHeight     =   3195
      ScaleWidth      =   3195
      TabIndex        =   2
      Top             =   3810
      Width           =   3195
      Begin VSFlex8Ctl.VSFlexGrid vsf������� 
         Height          =   1155
         Left            =   60
         TabIndex        =   3
         Top             =   360
         Width           =   3015
         _cx             =   5318
         _cy             =   2037
         Appearance      =   2
         BorderStyle     =   0
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483638
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
      Begin VSFlex8Ctl.VSFlexGrid vsfWorkInfo 
         Height          =   1215
         Left            =   60
         TabIndex        =   7
         Top             =   2010
         Width           =   3015
         _cx             =   5318
         _cy             =   2143
         Appearance      =   2
         BorderStyle     =   0
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483638
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
      Begin XtremeSuiteControls.ShortcutCaption sccWorkInfo 
         Height          =   315
         Left            =   0
         TabIndex        =   8
         Top             =   1650
         Width           =   3105
         _Version        =   589884
         _ExtentX        =   5477
         _ExtentY        =   564
         _StockProps     =   6
         Caption         =   "�ϰ���Ϣ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         GradientColorLight=   0
         GradientColorDark=   0
      End
      Begin XtremeSuiteControls.ShortcutCaption scc������� 
         Height          =   320
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   3105
         _Version        =   589884
         _ExtentX        =   5477
         _ExtentY        =   564
         _StockProps     =   6
         Caption         =   "������Ϣ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         GradientColorLight=   0
         GradientColorDark=   0
      End
   End
   Begin zl9RegEvent.UserSelectPopup uspSelectYear 
      Height          =   315
      Left            =   390
      TabIndex        =   0
      Top             =   510
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   556
      PopupWidth      =   1000
   End
   Begin zl9RegEvent.UserDatePicker dtpDay 
      Height          =   3045
      Left            =   180
      TabIndex        =   1
      Top             =   3900
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   5371
      HolidayStart    =   42379.5026851852
      TitleBackColor  =   -2147483626
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfGrid 
      Height          =   2055
      Left            =   870
      TabIndex        =   5
      Top             =   1080
      Width           =   7905
      _cx             =   13944
      _cy             =   3625
      Appearance      =   2
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483638
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   2
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   2
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
      BackColorFrozen =   -2147483643
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Line LineX 
      BorderColor     =   &H8000000C&
      X1              =   6750
      X2              =   6750
      Y1              =   7530
      Y2              =   3720
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H8000000C&
      Height          =   735
      Left            =   180
      Top             =   1140
      Width           =   405
   End
   Begin XtremeSuiteControls.ShortcutCaption sccTitle 
      Height          =   360
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   8625
      _Version        =   589884
      _ExtentX        =   15214
      _ExtentY        =   635
      _StockProps     =   6
      Caption         =   "��������>�ڼ��չ���"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmClinicHolidayManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mcbsMain As Object          'CommandBar�ؼ�
Private mlngModule As Long
Private mstrPrivs As String

Private Enum mGridHeadCol
    COl_���� = 0
    COL_��ʼʱ�� = 1
    COL_����ʱ�� = 2
    COL_��ע = 3
    COL_����ԤԼ = 4
    COL_����Һ� = 5
    
    COL_��� = 0
    COL_ԭ�ϰ�ʱ�� = 1
    Col_����ʱ�� = 2
    
    COL_���� = 0
    COL_�Һ� = 1
    COL_ԤԼ = 2
End Enum
Private mdatStart As Date, mdatEnd As Date, mvarWorks As Variant
Private mlngYear As Long

Public Sub InitCommVariable(frmParent As Form, cbsThis As Object, _
    ByVal strPrivs As String, ByVal lngModule As Long)
    '��ʼ������
    Set mfrmMain = frmParent
    Set mcbsMain = cbsThis
    
    mstrPrivs = strPrivs
    mlngModule = lngModule
End Sub

Public Sub zlDefCommandBars(Optional ByVal blnInsideTools As Boolean)
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    
    Err = 0: On Error GoTo ErrHandler
    
    '�ļ��˵�
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    With cbrMenuBar.CommandBar.Controls
        '���������Excel֮��
        Set cbrControl = .Find(, conMenu_File_Excel)
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXML, "����ΪXML�ļ�(&L)��", cbrControl.Index + 1)
    End With

    '�༭�˵�:���ڹ���˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If
    
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", cbrMenuBar.Index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "���ӽڼ���(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸Ľڼ���(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ���ڼ���(&D)")
    End With

    '�鿴�˵�
    '-----------------------------------------------------
    Set cbrMenuBar = mcbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Find(, conMenu_View_Refresh) 'ˢ����ǰ(���ʱע�ⷴ��)
'        Set cbrControl = .Add(xtpControlButton, conMenu_View_Notify, "ˢ������(&B)", cbrControl.Index)
        cbrControl.BeginGroup = True
    End With
    
    '����������
    '-----------------------------------------------------
    Set cbrToolBar = mcbsMain(2)
    For Each cbrControl In cbrToolBar.Controls '�����ǰ������һ��Control
        If Val(Left(cbrControl.ID, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.ID, 1)) <> conMenu_ManagePopup Then
            Set cbrControl = cbrToolBar.Controls(cbrControl.Index - 1): Exit For
        End If
    Next
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "���ӽڼ���", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸Ľڼ���", cbrControl.Index + 1)
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ���ڼ���", cbrControl.Index + 1)
        .Item(cbrControl.Index + 1).BeginGroup = True
    End With
    
    '����Ŀ����
    '-----------------------------------------------------
    With mcbsMain.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("B"), conMenu_Edit_Modify
        .Add 0, VK_DELETE, conMenu_Edit_Delete
    End With
    
    '���ò���������
    '-----------------------------------------------------
    With mcbsMain.Options
'        .AddHiddenCommand conMenu_Edit_Archive
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Dim blnVisible As Boolean, blnEnabled As Boolean
    Dim dtTemp As Date
    
    If Not Me.Visible Then Exit Sub
    On Error Resume Next
    blnVisible = zlStr.IsHavePrivs(mstrPrivs, "�ڼ�������")
    If vsfGrid.Row > 0 Then
        dtTemp = Val(uspSelectYear.SelectedKey) & "/" & vsfGrid.TextMatrix(vsfGrid.Row, COL_��ʼʱ��)
        blnEnabled = DateDiff("d", Now, dtTemp) > 0  'С�ڵ�ǰ���ڵĲ���ɾ�����޸�
    End If

    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = vsfGrid.Rows > vsfGrid.FixedRows
    Case conMenu_EditPopup
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible
    Case conMenu_Edit_NewItem
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Modify
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And blnEnabled
    Case conMenu_Edit_Delete
        Control.Visible = blnVisible
        Control.Enabled = Control.Visible And blnEnabled
    End Select
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim strHolidayName As String
    Dim frmEdit As frmClinicHolidayEdit
    
    Err = 0: On Error GoTo ErrHandler
    Select Case Control.ID
    'bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Case conMenu_File_Preview: Call zlDataPrint(2)
    Case conMenu_File_Print: Call zlDataPrint(1)
    Case conMenu_File_Excel: Call zlDataPrint(3)
    Case conMenu_Edit_NewItem
        Set frmEdit = New frmClinicHolidayEdit
        If frmEdit.ShowMe(Me, Fun_Add) Then
            RefrashData mlngYear 'ˢ������
        End If
    Case conMenu_Edit_Modify
        If vsfGrid.Row < 1 Then Exit Sub
        
        strHolidayName = vsfGrid.TextMatrix(vsfGrid.Row, COl_����)
        Set frmEdit = New frmClinicHolidayEdit
        If frmEdit.ShowMe(Me, Fun_Update, mlngYear, strHolidayName) Then
            RefrashData mlngYear 'ˢ������
        End If
    Case conMenu_Edit_Delete
        If ExcuteDelete() Then
            RefrashData mlngYear 'ˢ������
        End If
    Case conMenu_View_Refresh
        RefrashData mlngYear 'ˢ������
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function ExcuteDelete() As Boolean
    '����:ִ��ɾ������
    Dim strSQL  As String, rsTemp As ADODB.Recordset
    Dim strHolidayName As String
    
    On Error GoTo ErrHandler
    If vsfGrid.Row <= 0 Then Exit Function
    
    strHolidayName = vsfGrid.TextMatrix(vsfGrid.Row, COl_����)
    
    If MsgBox("��ȷ��Ҫɾ��" & mlngYear & "�� " & strHolidayName & " ��", _
        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    'ɾ����Ч�Լ��
    strSQL = "Select 1" & vbNewLine & _
        "    From �ٴ������¼ A" & vbNewLine & _
        "    Where a.�������� >= (Select ��ʼ���� From �������ձ� Where ��� = [1] And �������� = [2] And ���� = 0 And Rownum<2)" & vbNewLine & _
        "          And a.�ϰ�ʱ�� Is Not Null And Nvl(a.�Ƿ񷢲�, 0) = 1 And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngYear, strHolidayName)
    If Not rsTemp.EOF Then
        MsgBox "��ǰ�ڼ��տ�ʼʱ��֮��������Ч�ĳ��ﰲ�ţ�����ɾ����", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Zl_�������ձ�_Delete(
    strSQL = "Zl_�������ձ�_Delete("
    '���_In     �������ձ�.���%Type,
    strSQL = strSQL & "" & mlngYear & ","
    '��������_In �������ձ�.��������%Type
    strSQL = strSQL & "'" & strHolidayName & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    ExcuteDelete = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitGridHead()
    Dim strHead As String
    Dim i As Long, varData As Variant
    
    Err = 0: On Error GoTo ErrHandler
    strHead = "����,4,800|��ʼʱ��,4,1300|����ʱ��,4,1300|��ע,1,4500|����ԤԼ,1,0|����Һ�,1,0"
    With vsfGrid
        .Redraw = False
        .Rows = 1
        .FixedCols = 1: .FixedRows = 1
        .HighLight = flexHighlightAlways
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .BackColorAlternate = G_AlternateColor '�н���ɫ
        
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        Call RestoreFlexState(vsfGrid, App.ProductName & "\" & Me.Name)
        .Redraw = True
    End With

    strHead = "���,4,500|ԭ�ϰ�����,4,1300|��������,4,1300"
    With vsf�������
        .Redraw = False
        .FixedCols = 1: .FixedRows = 1
        .HighLight = flexHighlightWithFocus
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .BackColorAlternate = G_AlternateColor '�н���ɫ
        .RowHeightMin = 300
        
        .Rows = 1
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .Redraw = True
    End With
    
    strHead = "����,4,1300|����Һ�,4,1000|����ԤԼ,4,1000"
    With vsfWorkInfo
        .Redraw = flexRDNone
        .FixedCols = 0: .FixedRows = 1
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .BackColorAlternate = G_AlternateColor
        .RowHeightMin = 300
        .Editable = flexEDNone
        
        .Rows = 1
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = flexAlignCenterCenter
            If i > 0 Then
                .ColDataType(i) = flexDTBoolean
            End If
        Next
        .Redraw = flexRDBuffered
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub dtpDay_DayMetrics(Day As Date, Metrics As UserDatePickerDayMetrics)
    Dim dtTemp As Date, i As Integer
    
    Err = 0: On Error GoTo ErrHandler
    If CStr(Day) = "00:00:00" Then Exit Sub
    
    If Weekday(Day) = vbSunday Or Weekday(Day) = vbSaturday Then
        Metrics.ForeColor = vbRed
    End If
    
    '����ݼ���
    If DateDiff("d", Day, mdatStart) <= 0 And DateDiff("d", Day, mdatEnd) >= 0 Then
'        If HolidayIsWork(Day) = False Then
            Metrics.BackColor = &HC0E0FF
            Metrics.IsHoliday = True
'        End If
    End If
    
    '��ǵ�����
    For i = 1 To vsf�������.Rows - 1
        dtTemp = vsf�������.TextMatrix(i, Col_����ʱ��)
        If DateDiff("d", Day, dtTemp) = 0 Then
            Metrics.BackColor = &HFFFFC0
            Metrics.IsWorkFromHoliday = True
        End If
    Next
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function HolidayIsWork(ByVal Day As Date) As Boolean
    '���ڼ����Ƿ��ϰ�
    Dim i As Integer, j As Integer
    Dim var����ԤԼ As Variant, var����Һ� As Variant
    
    Err = 0: On Error GoTo ErrHandler
    If vsfGrid.Row < 1 Then Exit Function
    
    Err = 0: On Error GoTo ErrHandler
    var����ԤԼ = Split(vsfGrid.TextMatrix(vsfGrid.Row, COL_����ԤԼ), ";")
    var����Һ� = Split(vsfGrid.TextMatrix(vsfGrid.Row, COL_����Һ�), ";")
    
    For j = 0 To UBound(var����ԤԼ)
        If DateDiff("d", Day, var����ԤԼ(j)) = 0 Then
            HolidayIsWork = True
            Exit Function
        End If
    Next
    
    For j = 0 To UBound(var����Һ�)
        If DateDiff("d", Day, var����Һ�(j)) = 0 Then
            HolidayIsWork = True
            Exit Function
        End If
    Next
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Form_Activate()
    On Error Resume Next
    Call mfrmMain.ActiveFormChange(Me)
End Sub

Private Sub Form_Load()
    Dim varRow As Variant, varCol As Variant
    Dim i As Long, j As Long
    Err = 0: On Error GoTo ErrHandler
    Call InitGridHead
    scc�������.GradientColorDark = dtpDay.TitleBackColor
    scc�������.GradientColorLight = dtpDay.TitleBackColor
    
    mlngYear = Year(zlDatabase.Currentdate)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function RefrashData(Optional ByVal lngYear As Long) As Boolean
    Dim i As Long
    Dim lngMaxYear As Long, lngMinYear As Long
    Dim strSQL As String, rs��� As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHandler
    lngMinYear = 3000: lngMaxYear = 1900
    
    strSQL = "Select Max(���) As �� From �������ձ� Group By ���"
    Set rs��� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Do While Not rs���.EOF
        If lngMinYear > Val(Nvl(rs���!��)) Then lngMinYear = Val(Nvl(rs���!��))
        If lngMaxYear < Val(Nvl(rs���!��)) Then lngMaxYear = Val(Nvl(rs���!��))
        rs���.MoveNext
    Loop
    
    If lngMinYear = 3000 Then lngMinYear = lngYear
    If lngMaxYear = 1900 Then lngMaxYear = lngYear
    If lngYear < lngMinYear Or lngYear > lngMaxYear Then
        lngYear = Year(zlDatabase.Currentdate)
    End If
    If lngYear < lngMinYear Then lngMinYear = lngYear
    If lngYear > lngMaxYear Then lngMaxYear = lngYear
    'Ϊ���ѡ������������
    uspSelectYear.Clear
    For i = lngMinYear To lngMaxYear
        uspSelectYear.AddItem i, i & "��"
    Next
    uspSelectYear.SelectedKey = lngYear '�ᴥ��uspSelectYear_ValueChanged�¼�
'    Call LoadData(lngYear)
    RefrashData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadData(Optional ByVal lngYear As Long) As Boolean
    Dim i As Long, j As Long, lngRow As Long
    Dim strSQL As String, rs�ڼ��� As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    mdatStart = Empty: mdatEnd = Empty: mvarWorks = Empty
    vsf�������.Clear 1: vsf�������.Rows = 1
    vsfWorkInfo.Clear 1: vsfWorkInfo.Rows = 1
    dtpDay.RedrawControl
    
    strSQL = "Select ���,��������,��ʼ����,��ֹ����,��ע,����ԤԼ����,����Һ����� From �������ձ�" & vbNewLine & _
            " Where Nvl(����,0)=0 And ���=[1]" & vbNewLine & _
            " Order By ���,��ʼ����"
    Set rs�ڼ��� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngYear)
    
    With rs�ڼ���
        lngRow = vsfGrid.Row
        vsfGrid.Rows = .RecordCount + 1
        i = 1
        Do While Not .EOF
            uspSelectYear.SelectedKey = Nvl(!���)
            vsfGrid.TextMatrix(i, COl_����) = Nvl(!��������)
            vsfGrid.TextMatrix(i, COL_��ʼʱ��) = Format(Nvl(!��ʼ����), "mm-dd hh:mm")
            vsfGrid.Cell(flexcpData, i, COL_��ʼʱ��) = Nvl(!��ʼ����)
            vsfGrid.TextMatrix(i, COL_����ʱ��) = Format(Nvl(!��ֹ����), "mm-dd hh:mm")
            vsfGrid.Cell(flexcpData, i, COL_����ʱ��) = Nvl(!��ֹ����)
            vsfGrid.TextMatrix(i, COL_��ע) = Nvl(!��ע)
            vsfGrid.TextMatrix(i, COL_����ԤԼ) = Nvl(!����ԤԼ����)
            vsfGrid.TextMatrix(i, COL_����Һ�) = Nvl(!����Һ�����)
            i = i + 1
            .MoveNext
        Loop
    End With
    If vsfGrid.Rows > 1 Then
        vsfGrid.Row = -1 '��֤��ѡ���в���������Ҳ����RowColChange�¼�
        If lngRow = 0 Then
            vsfGrid.Row = 1
        ElseIf lngRow > vsfGrid.Rows - 1 Then
            vsfGrid.Row = vsfGrid.Rows - 1
        Else
            vsfGrid.Row = lngRow
        End If
    End If
    
    Screen.MousePointer = vbDefault
    LoadData = True
    Exit Function
ErrHandler:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitDatePickerData(ByVal datStart As Date, ByVal datEnd As Date, Optional ByVal strWorks As String)
    '���ܣ����ݽڼ��տ�ʼʱ��ͽ���ʱ�䣬�Լ�����ʱ����ʾ����
    '������
    '   datStart - ��ʼʱ��
    '   datEnd - ����ʱ��
    '   varWorks - ����(�ϰ�)ʱ�䣬�����"��"�ָ�
    Err = 0: On Error GoTo ErrHandler
    If datStart > datEnd Then 'ȷ��ʱ���С
        Dim datTemp As Date
        datTemp = datStart: datStart = datEnd: datEnd = datTemp
    End If
    mvarWorks = Empty
    If strWorks <> "" Then mvarWorks = Split(strWorks, "��")
    mdatStart = datStart: mdatEnd = datEnd
    
    dtpDay.HolidayStart = mdatStart '�ᴥ��RedrawControl()����
'    dtpDay.RedrawControl
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    shpBorder.Move 0, 0, Me.ScaleWidth - 6, Me.ScaleHeight - 6
    
    sccTitle.Move 10, 0, Me.ScaleWidth
    With uspSelectYear
        .Left = 10: .Top = sccTitle.Top + sccTitle.Height
        .Width = Me.ScaleWidth - .Left + 10
    End With
    
    With vsfGrid
        .Left = 0: .Top = uspSelectYear.Top + uspSelectYear.Height
        .Width = Me.ScaleWidth - .Left + 10
        .Height = Me.ScaleHeight * 7 / 12 - .Top - 10
    End With
    
    With dtpDay
        .Left = 10: .Top = vsfGrid.Top + vsfGrid.Height
        .Width = Me.ScaleWidth - 4000 - .Left
        .Height = Me.ScaleHeight - .Top - 10
    End With
    LineX.X1 = dtpDay.Left + dtpDay.Width + 10: LineX.Y1 = dtpDay.Top
    LineX.X2 = dtpDay.Left + dtpDay.Width + 10: LineX.Y2 = dtpDay.Top + dtpDay.Height
    With pic�������
        .Left = dtpDay.Left + dtpDay.Width + 20: .Top = dtpDay.Top
        .Width = Me.ScaleWidth - .Left: .Height = dtpDay.Height
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveFlexState(vsfGrid, App.ProductName & "\" & Me.Name)
End Sub

Private Sub pic�������_Resize()
    Err = 0: On Error Resume Next
    scc�������.Move 0, 0, pic�������.ScaleWidth
    With vsf�������
        .Left = 0: .Top = scc�������.Top + scc�������.Height
        .Width = pic�������.ScaleWidth + 20
    End With
    
    sccWorkInfo.Move 0, vsf�������.Top + vsf�������.Height, pic�������.ScaleWidth
    With vsfWorkInfo
        .Left = 0: .Top = sccWorkInfo.Top + sccWorkInfo.Height
        .Width = pic�������.ScaleWidth + 20
        .Height = pic�������.ScaleHeight - .Top + 10
    End With
End Sub

Private Sub sccTitle_GotFocus()
    On Error Resume Next
    If vsfGrid.Visible And vsfGrid.Enabled Then vsfGrid.SetFocus
End Sub

Private Sub uspSelectYear_ValueChanged(ByVal strKey As String, ByVal strValue As String)
    dtpDay.HolidayStart = strKey & "/02/01"
    mlngYear = Val(strKey)
    Call LoadData(mlngYear)
End Sub

Private Sub vsfGrid_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Integer
    Dim strSQL As String, rs�ڼ��� As ADODB.Recordset
    Dim strHolidayName As String
    
    If NewRow < 1 Or vsfGrid.Rows - 1 < NewRow Then Exit Sub
    strHolidayName = vsfGrid.TextMatrix(NewRow, COl_����)
    With vsf�������
        strSQL = "Select ���,��������,��ʼ����,��ֹ����,��ע From �������ձ�" & _
                " Where Nvl(����,0)=1 And ���=[1] And ��������=[2]"
        Set rs�ڼ��� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngYear, strHolidayName)
        .Rows = rs�ڼ���.RecordCount + 1
        i = 1
        Do While Not rs�ڼ���.EOF
            .TextMatrix(i, COL_���) = i
            .TextMatrix(i, COL_ԭ�ϰ�ʱ��) = Format(Nvl(rs�ڼ���!��ֹ����), "yyyy-mm-dd")
            .TextMatrix(i, Col_����ʱ��) = Format(Nvl(rs�ڼ���!��ʼ����), "yyyy-mm-dd")
            .RowHeight(i) = .RowHeight(0)
            i = i + 1
            rs�ڼ���.MoveNext
        Loop
    End With
    
    Call ShowDateRangeToGrid(vsfGrid.Cell(flexcpData, NewRow, COL_��ʼʱ��), vsfGrid.Cell(flexcpData, NewRow, COL_����ʱ��))
    Call LoadDateRegist(vsfGrid.TextMatrix(NewRow, COL_����Һ�), vsfGrid.TextMatrix(NewRow, COL_����ԤԼ))
    
    If vsfGrid.TextMatrix(NewRow, COL_��ʼʱ��) <> "" Then
        InitDatePickerData vsfGrid.Cell(flexcpData, NewRow, COL_��ʼʱ��), _
              vsfGrid.Cell(flexcpData, NewRow, COL_����ʱ��)
    End If
End Sub

Private Sub vsfGrid_DblClick()
    Dim strHolidayName As String
    Dim dtTemp As Date
    Dim blnEnabled As Boolean
    Dim frmEdit As frmClinicHolidayEdit
    
    Err = 0: On Error GoTo ErrHandler
    If vsfGrid.Row <= 0 Then Exit Sub
    strHolidayName = vsfGrid.TextMatrix(vsfGrid.Row, COl_����)
    If vsfGrid.Row > 0 Then
        dtTemp = mlngYear & "/" & vsfGrid.TextMatrix(vsfGrid.Row, COL_��ʼʱ��)
        blnEnabled = DateDiff("d", Now, dtTemp) > 0  'С�ڵ�ǰ���ڵĲ���ɾ��
    End If
    
    Set frmEdit = New frmClinicHolidayEdit
    If zlStr.IsHavePrivs(mstrPrivs, "�ڼ�������") And blnEnabled Then
        If frmEdit.ShowMe(Me, Fun_Update, mlngYear, strHolidayName) Then Call RefrashData(mlngYear)   'ˢ������
    Else
        frmEdit.ShowMe Me, Fun_View, mlngYear, strHolidayName
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    
    Err = 0: On Error GoTo ErrHandler
    If Not (Button = vbRightButton) Then Exit Sub
    If Not (Me.Visible And Me.Enabled) Then Exit Sub
    Me.SetFocus: Call mfrmMain.ActiveFormChange(Me)
    
    Set objPopup = mcbsMain.FindControl(xtpControlPopup, conMenu_EditPopup, , True)
    If objPopup Is Nothing Then Exit Sub
    If objPopup.Visible = False Then Exit Sub
    objPopup.CommandBar.ShowPopup
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub zlDataPrint(BytMode As Byte)
    '����:���д�ӡ,Ԥ���������EXCEL
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    If UserInfo.���� = "" Then Call GetUserInfo
    Dim objOut As New zlPrint1Grd, objRow As New zlTabAppRow
    Dim bytR As Byte
    
    Err = 0: On Error GoTo ErrHandler
    objOut.Title.Text = mlngYear & "��ڼ����嵥"
    Set objOut.Body = vsfGrid
    
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True

    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    If BytMode = 1 Then
        bytR = zlPrintAsk(objOut)
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, BytMode
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ShowDateRangeToGrid(ByVal dtStart As Date, dtEnd As Date)
    '��ʾ���ڵ������
    Dim lngRow As Long, i As Integer
    Dim intCount As Integer
    
    Err = 0: On Error GoTo ErrHandler
    intCount = DateDiff("d", dtStart, dtEnd) '������
    With vsfWorkInfo
        .Clear 1
        .Rows = 1
        For i = 0 To intCount
            .Rows = .Rows + 1
            lngRow = .Rows - 1
            .TextMatrix(lngRow, COL_����) = Format(DateAdd("d", i, dtStart), "yyyy-mm-dd")
            .Cell(flexcpChecked, lngRow, COL_�Һ�, lngRow, COL_ԤԼ) = 2
        Next
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub LoadDateRegist(ByVal str����Һ� As String, ByVal str����ԤԼ As String)
    '����ԤԼ�Һ����
    Dim i As Integer, j As Integer
    Dim var����ԤԼ As Variant, var����Һ� As Variant
    
    Err = 0: On Error GoTo ErrHandler
    var����Һ� = Split(str����Һ�, ";")
    var����ԤԼ = Split(str����ԤԼ, ";")
    With vsfWorkInfo
        For i = 1 To .Rows - 1
            For j = 0 To UBound(var����Һ�)
                If DateDiff("d", .TextMatrix(i, COL_����), var����Һ�(j)) = 0 Then
                    .TextMatrix(i, COL_�Һ�) = 1
                End If
            Next
            For j = 0 To UBound(var����ԤԼ)
                If DateDiff("d", .TextMatrix(i, COL_����), var����ԤԼ(j)) = 0 Then
                    .TextMatrix(i, COL_ԤԼ) = 1
                End If
            Next
        Next
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

