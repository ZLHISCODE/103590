VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmSetClinicCharge 
   Caption         =   "��Ʋ���-���ö���"
   ClientHeight    =   7680
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   12840
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSetClinicCharge.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   12840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTwo 
      BorderStyle     =   0  'None
      Height          =   3540
      Left            =   7320
      ScaleHeight     =   3540
      ScaleWidth      =   4920
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2985
      Width           =   4920
      Begin VSFlex8Ctl.VSFlexGrid vsCharge1 
         Height          =   3030
         Left            =   210
         TabIndex        =   9
         Top             =   150
         Width           =   12240
         _cx             =   21590
         _cy             =   5345
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
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
         BackColorSel    =   -2147483637
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
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
   End
   Begin VB.PictureBox picOne 
      BorderStyle     =   0  'None
      Height          =   3555
      Left            =   285
      ScaleHeight     =   3555
      ScaleWidth      =   7680
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3240
      Width           =   7680
      Begin VSFlex8Ctl.VSFlexGrid vsCharge 
         Height          =   3030
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   12840
         _cx             =   22648
         _cy             =   5345
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
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
         BackColorSel    =   -2147483637
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
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
         Begin VB.TextBox txtEdit 
            Height          =   375
            Left            =   6390
            TabIndex        =   7
            Top             =   315
            Visible         =   0   'False
            Width           =   1125
         End
      End
   End
   Begin XtremeSuiteControls.TabControl TbcCharge 
      Height          =   3375
      Left            =   1800
      TabIndex        =   4
      Top             =   3930
      Width           =   7530
      _Version        =   589884
      _ExtentX        =   13282
      _ExtentY        =   5953
      _StockProps     =   64
   End
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   8730
      TabIndex        =   3
      Top             =   105
      Width           =   1590
   End
   Begin VB.Frame FraNs 
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   -510
      MousePointer    =   7  'Size N S
      TabIndex        =   1
      Top             =   5235
      Width           =   8640
   End
   Begin VSFlex8Ctl.VSFlexGrid vsClinic 
      Height          =   2985
      Left            =   15
      TabIndex        =   0
      Top             =   360
      Width           =   12840
      _cx             =   22648
      _cy             =   5265
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
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
      BackColorSel    =   -2147483637
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   7305
      Width           =   12840
      _ExtentX        =   22648
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmSetClinicCharge.frx":29F2
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17568
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   7575
      Top             =   45
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmSetClinicCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum COL_VsClinic
    ID = 0: �����: ���: ����: ����: �걾��λ: ���㵥λ: ��������: ִ��Ƶ��: �Ƽ�����: ���㷽ʽ: �������: �������: վ��: ����
End Enum
Private Enum COL_VsCharge
    ID = 0: ���: ��Ŀ��: ���: ��λ: �۸�: ����: �̶�: ����: �շѷ�ʽ: ͣ��
End Enum
Private mintFindType As Integer '����ģʽ 1-���� 2-���Ʊ���
Private mblnEditMode As Boolean '�Ƿ�༭ģʽ

Private mlngDeptID As Long      '����ID
Private mlngClinicID As Long    '��ǰ������ĿID
Private mlng�Ƽ����� As Long    '��ǰ������Ŀ�ļƼ�����
Private mblnCopy As Boolean     '�Ƿ����ʹ�ø��Ƶ����������� ����
Private mstrTitle As String     '����
Private mint���� As Integer     '���ó���,1-���2-סԺ
Private mblnModify As Boolean   '�Ƿ��޸�
Private mblnModifyPrivs As Boolean
Private mlngModul As Long       '���õ�ģ���

Public Sub ShowMe(ByVal lngDeptID As Long, ByVal lngMode As Long, frmMain As Form, int���� As Integer, ByVal blnModify As Boolean)
    Dim strSql As String, rsTmp As ADODB.Recordset
    On Error GoTo errH
    If lngDeptID <= 0 Then Exit Sub
    mlngDeptID = lngDeptID
    mlngClinicID = 0
    mint���� = int����
    mblnModifyPrivs = blnModify
    
'    If InStr(frmMain.Caption, "����ҽ������վ -") > 0 Then
'        mlngModul = 1260
'    ElseIf InStr(frmMain.Caption, "סԺ��ʿ����վ -") > 0 Then
'        mlngModul = 1262
'    ElseIf InStr(frmMain.Caption, "ҽ������վ -") > 0 Then
'        mlngModul = 1263
'    End If
    mlngModul = glngModul
    
    If mint���� = 2 Then
        strSql = "Select Distinct ID, ����, ����" & vbNewLine & _
        "From ���ű� A, ������Ա B, �ϻ���Ա�� C, ��������˵�� D" & vbNewLine & _
        "Where a.Id = b.����id And b.��Աid = c.��Աid And a.Id = d.����id And (d.������� = 2 or d.�������=3) " & _
        "    And d.�������� in ('����','���','����','����','����','����','Ӫ��') And A.ID<>[1] And c.�û��� = User"
    Else
        strSql = "Select Distinct ID, ����, ����" & vbNewLine & _
        "From ���ű� A, ������Ա B, �ϻ���Ա�� C, ��������˵�� D" & vbNewLine & _
        "Where a.Id = b.����id And b.��Աid = c.��Աid And a.Id = d.����id And (d.������� = 1 Or d.������� = 3) " & vbNewLine & _
        "    And Instr('�ٴ�,����,���,����,����,����,Ӫ��', d.��������) > 0 And a.Id <> [1] And c.�û��� = User"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngDeptID)
    mblnCopy = Not rsTmp.EOF
    
    strSql = "Select Distinct ID, ����, ���� From ���ű� A Where A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngDeptID)
    mstrTitle = "���ö���"
    If Not rsTmp.EOF Then mstrTitle = "" & rsTmp!���� & "(" & rsTmp!���� & ")" & "-" & IIf(mint���� = 2, "סԺ���ö���", "������ö���")
    Me.Show lngMode, frmMain

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer, objControl As CommandBarControl
    
    On Error GoTo ErrHandle
    Select Case Control.ID
 
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    
    Case conMenu_View_ToolBar_Button '������
        For i = 2 To cbsThis.Count
            Me.cbsThis(i).Visible = Not Me.cbsThis(i).Visible
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text '��ť����
        For i = 2 To cbsThis.Count
            For Each objControl In Me.cbsThis(i).Controls
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size '��ͼ��
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar '״̬��
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_Help_Web_Home 'Web�ϵ�����
        Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Forum '������̳
        Call zlWebForum(Me.hWnd)
    Case conMenu_Help_Web_Mail '���ͷ���
        Call zlMailTo(Me.hWnd)
    Case conMenu_Help_Help '����
        Call ShowHelp(gstrLisHelp, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_About '����
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_File_Exit '�˳�
        Unload Me
    Case conMenu_Edit_Modify '����༭
        mblnModify = False
        Call CheckIsNoCharge '����Ƿ��δ���չ����ã����������п��ҵĶ��ա�
        mblnEditMode = True
        vsClinic.Enabled = False
        txtFind.Enabled = vsClinic.Enabled
        With vsCharge
            .Editable = flexEDKbdMouse
            .SelectionMode = flexSelectionFree
            If TbcCharge.Selected.Index <> 0 Then TbcCharge.Item(0).Selected = True
            .SetFocus
            .Col = COL_VsCharge.��Ŀ��
        End With
        Select Case Trim("" & vsClinic.TextMatrix(vsClinic.Row, COL_VsClinic.�Ƽ�����))
            Case "�����Ƽ�": mlng�Ƽ����� = 0
            Case "���Ƽ�": mlng�Ƽ����� = 1
            Case "�ֹ��Ƽ�": mlng�Ƽ����� = 2
        End Select
    Case conMenu_Edit_Save   '����
        
        If SaveData(mlngDeptID) Then
            mblnEditMode = False
            vsClinic.Enabled = True
            txtFind.Enabled = vsClinic.Enabled
            vsCharge.Editable = flexEDNone
            vsCharge.SelectionMode = flexSelectionByRow
            Call vsChargeRef(mlngClinicID)
        End If
    Case conMenu_Edit_Untread 'ȡ��
        If mblnModify Then
            If MsgBox("�Ƿ������ǰ�ѵ��������ݣ�", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        mblnEditMode = False
        vsClinic.Enabled = True
        txtFind.Enabled = vsClinic.Enabled
        vsCharge.Editable = flexEDNone
        vsCharge.SelectionMode = flexSelectionByRow
        Call vsChargeRef(mlngClinicID)
    
    Case conMenu_Edit_NewItem '����
        mblnModify = True
        With vsCharge
            If Val(.TextMatrix(.Rows - 1, COL_VsCharge.ID)) <> 0 Then
                .Rows = .Rows + 1
            ElseIf .Rows = .FixedRows Then
                .Rows = .Rows + 1
            End If
            .Row = .Rows - 1
        End With
    Case conMenu_Edit_Delete 'ɾ��
        '
        mblnModify = True
        Call DeleteCharge
        
    Case conMenu_View_FindType  '���ҷ�ʽ
        
        cbsThis.RecalcLayout
        txtFind.Text = ""
        txtFind.SetFocus
        
    Case conMenu_View_Find '����
        If Me.ActiveControl Is txtFind Then
            txtFind.SetFocus '��ʱ��Ҫ��λһ��
            If txtFind.Text <> "" Then
                Call ExecuteFind
            End If
        Else
            txtFind.SetFocus
        End If
    Case conMenu_View_FindNext '������һ��
        If txtFind.Text = "" Then
            txtFind.SetFocus
        Else
            Call ExecuteFind(True)
        End If
    Case conMenu_Edit_Copy 'Ӧ������������
        Call DeptCopy
    End Select
    Exit Sub

ErrHandle:
    Call ErrCenter
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    If CommandBar.Parent Is Nothing Then Exit Sub
    '�Ҽ������˵�
'    Select Case CommandBar.Parent.ID
'    Case conMenu_View_FindType '���ҷ�ʽ
'        With CommandBar.Controls
'            If .Count = 0 Then
'                .Add xtpControlButton, conMenu_View_FindType * 100# + 1, "����(&1)"
'                .Add xtpControlButton, conMenu_View_FindType * 100# + 2, "���������(&2)"
'            End If
'        End With
'    End Select
End Sub

Private Sub cbsThis_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    FraNs.Left = Me.ScaleLeft
    FraNs.Width = Me.ScaleWidth
    
    With vsClinic
        .Left = lngLeft
        .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = FraNs.Top - .Top
    End With
    With TbcCharge
        .Left = lngLeft
        .Top = FraNs.Top + FraNs.Height
        .Width = lngRight - lngLeft
        If Me.stbThis.Visible Then
            .Height = Me.ScaleHeight - .Top - Me.stbThis.Height
        Else
            .Height = Me.ScaleHeight - .Top
        End If

    End With
    
'    With vsCharge
'        .Left = lngLeft
'        .Width = lngRight - lngLeft
'        .Top = FraNs.Top + FraNs.Height
'
'        If Me.stbThis.Visible Then
'            .Height = Me.ScaleHeight - .Top - Me.stbThis.Height
'        Else
'            .Height = Me.ScaleHeight - .Top
'        End If
'    End With


End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '������
        If cbsThis.Count >= 2 Then
            Control.Checked = Me.cbsThis(2).Visible
        End If
    Case conMenu_View_ToolBar_Text 'ͼ������
        If cbsThis.Count >= 2 Then
            Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size '��ͼ��
        Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar '״̬��
        Control.Checked = Me.stbThis.Visible
        
    '-------------------------------
    Case conMenu_Edit_Modify, conMenu_Edit_Copy
        Control.Enabled = Not mblnEditMode And mblnModifyPrivs
        If Control.ID = conMenu_Edit_Copy And Control.Enabled Then
            Control.Enabled = mblnCopy
        End If
    Case conMenu_Edit_Save, conMenu_Edit_Untread, conMenu_Edit_NewItem, conMenu_Edit_Delete
        Control.Enabled = mblnEditMode
'    Case conMenu_View_FindType '���ҷ�ʽ
'        If Control.Parent Is cbsThis.ActiveMenuBar Then
'            Control.Caption = "����" & IIf(mintFindType = 0, "����", "���������") & "����"
'        End If
    End Select
End Sub

Private Sub Form_Activate()
    mlngClinicID = -1
    Me.Caption = mstrTitle
    Call vsClinic_RowColChange
End Sub

Private Sub Form_Load()
    
    '��ʼ������
    
    Call initMenu
    Call initTbcCharge
    Call initVsClinic
    Call initVsCharge(vsCharge)
    Call initVsCharge(vsCharge1)
    'װ������
    
    Call zlRefRecords
    Call RestoreWinState(Me, App.ProductName)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnEditMode Then Call cbsThis_Execute(Me.cbsThis.FindControl(, conMenu_Edit_Untread))
    If mblnEditMode Then
        Cancel = True
        Exit Sub
    End If
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub FraNs_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        On Error Resume Next
        If vsClinic.Height + y < 1000 Or TbcCharge.Height - y < 1000 Then Exit Sub
        FraNs.Top = FraNs.Top + y
        vsClinic.Height = vsClinic.Height + y
        TbcCharge.Top = TbcCharge.Top + y
        TbcCharge.Height = TbcCharge.Height - y
    End If
End Sub

Private Sub initTbcCharge()
    With TbcCharge
        If .ItemCount <= 0 Then
            With .PaintManager
                .Appearance = xtpTabAppearanceExcel
                .ClientFrame = xtpTabFrameSingleLine
                
                .Position = xtpTabPositionTop  'ѡ��ڶ���
                .BoldSelected = True
                .OneNoteColors = True
                .ShowIcons = True
            End With
             
            .InsertItem(0, "��ǰ" & IIf(mint���� = 2, "����", "����"), picOne.hWnd, 0).Tag = "��ǰ" & IIf(mint���� = 2, "����", "����")
            .InsertItem(1, "ȫԺͨ��", picTwo.hWnd, 0).Tag = "ȫԺͨ��"
            .Item(0).Selected = True
        End If
    End With
End Sub

Private Sub initVsClinic()
    With vsClinic
        .Clear
        '��ɫ
        .BackColorBkg = vbWindowBackground          '���ڱ���*
        .BackColorSel = &HFFEBD7    'vbInactiveBorder            '�ǻ�߿�*
        .GridColor = vbActiveBorder                 '��߿� *
        .SheetBorder = vbWindowBackground           '���ڱ���*
        .ForeColorSel = vbWindowText                '�����ı�*
        '
        .FocusRect = flexFocusNone                  '�޵�Ԫ�񽹵��
        
        '��ʼ�У���
        .Cols = 15
        .Rows = 2
        
        .TextMatrix(0, COL_VsClinic.ID) = "ID": .ColWidth(COL_VsClinic.ID) = 0: .ColHidden(COL_VsClinic.ID) = True
        .TextMatrix(0, COL_VsClinic.�����) = "�����": .ColWidth(COL_VsClinic.�����) = 0: .ColHidden(COL_VsClinic.�����) = True
        .TextMatrix(0, COL_VsClinic.վ��) = "վ��": .ColWidth(COL_VsClinic.վ��) = 0: .ColHidden(COL_VsClinic.վ��) = True
        .TextMatrix(0, COL_VsClinic.վ��) = "����": .ColWidth(COL_VsClinic.����) = 0: .ColHidden(COL_VsClinic.����) = True
        .TextMatrix(0, COL_VsClinic.�걾��λ) = "�걾��λ": .ColWidth(COL_VsClinic.�걾��λ) = 0: .ColHidden(COL_VsClinic.�걾��λ) = True
        
        .TextMatrix(0, COL_VsClinic.���) = "���": .ColWidth(COL_VsClinic.���) = 600
        .TextMatrix(0, COL_VsClinic.����) = "����": .ColWidth(COL_VsClinic.����) = 1200
        .TextMatrix(0, COL_VsClinic.����) = "������Ŀ����": .ColWidth(COL_VsClinic.����) = 3200
        
        .TextMatrix(0, COL_VsClinic.���㵥λ) = "���㵥λ": .ColWidth(COL_VsClinic.���㵥λ) = 900
        .TextMatrix(0, COL_VsClinic.��������) = "��������": .ColWidth(COL_VsClinic.��������) = 1200
        .TextMatrix(0, COL_VsClinic.ִ��Ƶ��) = "ִ��Ƶ��": .ColWidth(COL_VsClinic.ִ��Ƶ��) = 1200
        .TextMatrix(0, COL_VsClinic.�Ƽ�����) = "�Ƽ�����": .ColWidth(COL_VsClinic.�Ƽ�����) = 1200
        .TextMatrix(0, COL_VsClinic.���㷽ʽ) = "���㷽ʽ": .ColWidth(COL_VsClinic.���㷽ʽ) = 1200
        .TextMatrix(0, COL_VsClinic.�������) = "�������": .ColWidth(COL_VsClinic.�������) = 1200
        .TextMatrix(0, COL_VsClinic.�������) = "�������": .ColWidth(COL_VsClinic.�������) = 1200
        
        .Cell(flexcpAlignment, 0, .FixedCols, 0, .Cols - 1) = flexAlignCenterCenter
        
        .AllowSelection = False
        .AllowUserResizing = flexResizeColumns
    End With

End Sub

Private Sub initVsCharge(ByRef objVsf As VSFlexGrid)
    With objVsf
        .Clear
        '��ɫ
        .BackColorBkg = vbWindowBackground          '���ڱ���*
        .BackColorSel = &HFFEBD7 'vbInactiveBorder            '�ǻ�߿�*
        .GridColor = vbActiveBorder                 '��߿� *
        .SheetBorder = vbWindowBackground           '���ڱ���*
        .ForeColorSel = vbWindowText                '�����ı�*
        '
        .FocusRect = flexFocusNone                  '�޵�Ԫ�񽹵��
        
        '��ʼ�У���
        .Cols = 11
        .Rows = 2
        .TextMatrix(0, COL_VsCharge.ID) = "ID": .ColWidth(COL_VsCharge.ID) = 0: .ColHidden(COL_VsCharge.ID) = True
        
        .TextMatrix(0, COL_VsCharge.���) = "���": .ColWidth(COL_VsCharge.���) = 500
        .TextMatrix(0, COL_VsCharge.��Ŀ��) = "�շ���Ŀ����": .ColWidth(COL_VsCharge.��Ŀ��) = 3600
        .TextMatrix(0, COL_VsCharge.���) = "���": .ColWidth(COL_VsCharge.���) = 2600
        .TextMatrix(0, COL_VsCharge.��λ) = "��λ": .ColWidth(COL_VsCharge.��λ) = 900
        .TextMatrix(0, COL_VsCharge.�۸�) = "�۸�": .ColWidth(COL_VsCharge.�۸�) = 1000
        .TextMatrix(0, COL_VsCharge.����) = "����": .ColWidth(COL_VsCharge.����) = 800
        .TextMatrix(0, COL_VsCharge.�̶�) = "�̶�": .ColWidth(COL_VsCharge.�̶�) = 500
        .TextMatrix(0, COL_VsCharge.����) = "����": .ColWidth(COL_VsCharge.����) = 500
        .TextMatrix(0, COL_VsCharge.�շѷ�ʽ) = "�շѷ�ʽ": .ColWidth(COL_VsCharge.�շѷ�ʽ) = 1800
        .TextMatrix(0, COL_VsCharge.ͣ��) = "ͣ��": .ColWidth(COL_VsCharge.ͣ��) = 500
        
        .Cell(flexcpAlignment, 0, .FixedCols, 0, .Cols - 1) = flexAlignCenterCenter
        .ColAlignment(COL_VsCharge.�̶�) = flexAlignCenterCenter
        .ColAlignment(COL_VsCharge.����) = flexAlignCenterCenter
        .AllowUserResizing = flexResizeColumns
        
        .ColComboList(COL_VsCharge.�շѷ�ʽ) = "0-������ȡ|1-�����Թܷ���|2-һ�η���ֻ��ȡһ��|3-����ֻ��ȡһ��|4-����δִ����ȡһ��|5-����ֻ��ȡһ�Σ��ų�������Ŀ|6-����δִ����ȡһ�Σ��ų�������Ŀ"
    End With
End Sub

Private Sub initMenu()
    Dim cbrControl As CommandBarControl
    Dim objPopup As CommandBarPopup

    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim objCustom As CommandBarControlCustom
 
    'Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    '-----------------------------------------------------

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsThis.VisualTheme = xtpThemeOffice2003
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsThis.EnableCustomization False
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    
    '-----------------------------------------------------
    '�˵�����
    Me.cbsThis.ActiveMenuBar.Title = "�˵�"
    'Call Me.cbsThis.ActiveMenuBar.EnableDocking(xtpFlagAlignTop) '�����ˣ��Ͳ��ܿ��Ʋ��ҿ��λ����
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "����(&M)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��(&U)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Copy, "Ӧ��������" & IIf(mint���� = 2, "����", "����") & "(&C)"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        
        'Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_FindType, "����(&Y)"): objPopup.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_FindNext, "������һ��")

        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set objPopup = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        objPopup.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        objPopup.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With

    '���������⴦��
    '-----------------------------------------------------
    
    '���˵��Ҳ�Ĳ��� �����￨�Ų��ң�֧��ˢ��
    With Me.cbsThis.ActiveMenuBar.Controls
        Set cbrControl = .Add(xtpControlLabel, conMenu_View_FindType, "����")
        cbrControl.ID = conMenu_View_FindType
        cbrControl.Flags = xtpFlagRightAlign
        
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Find, "")
        objCustom.Handle = txtFind.hWnd
        objCustom.Flags = xtpFlagRightAlign
'
    End With
    
    '�����
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("Z"), conMenu_Edit_Untread
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F3, conMenu_View_FindNext
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    

    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        'Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        'Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "����")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Copy, "Ӧ��������" & IIf(mint���� = 2, "����", "����")): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '���ò����ò˵�
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
End Sub

Private Sub picOne_Resize()
    With vsCharge
        .Top = picOne.ScaleTop
        .Width = picOne.ScaleWidth
        .Height = picOne.ScaleHeight
        .Left = picOne.ScaleLeft
    End With
End Sub

Private Sub picTwo_Resize()
    With vsCharge1
        .Top = picTwo.ScaleTop
        .Width = picTwo.ScaleWidth
        .Height = picTwo.ScaleHeight
        .Left = picTwo.ScaleLeft
    End With
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    '���س�
    If KeyAscii = vbKeyReturn Then
        ExecuteFind
    End If
End Sub

Private Sub vsCharge_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (Col = COL_VsCharge.��Ŀ�� Or _
        Col = COL_VsCharge.���� Or Col = COL_VsCharge.�շѷ�ʽ) Then
        Cancel = True
    End If
End Sub

Private Sub vsCharge_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If mblnEditMode Then
        If NewCol = COL_VsCharge.��Ŀ�� Then
            vsCharge.ComboList = "..."
        Else
            vsCharge.ComboList = ""
        End If
    Else
        vsCharge.ComboList = ""
    End If
End Sub

Private Sub vsCharge_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    '
    Call SelectCharge(vsCharge.Row, vsCharge.Col)
End Sub

Private Sub vsCharge_DblClick()
    With vsCharge
        If (.Col = COL_VsCharge.�̶� Or .Col = COL_VsCharge.����) And mblnEditMode = True Then
            mblnModify = True
            If .TextMatrix(.Row, .Col) = "" Then
                .TextMatrix(.Row, .Col) = "��"
            Else
                .TextMatrix(.Row, .Col) = ""
            End If
        End If
    End With
End Sub

Private Sub vsCharge_EnterCell()
    With vsCharge
    
        If mblnEditMode Then
            If .Col = COL_VsCharge.��Ŀ�� Then
                .FocusRect = flexFocusHeavy
                If txtEdit.Tag = "False" Then
                    txtEdit.Left = .CellLeft
                    txtEdit.Top = .CellTop
                    txtEdit.Height = .CellHeight - 12
                    txtEdit.Width = .CellWidth - 12
                    txtEdit.Tag = "True"
                End If
            ElseIf .Col = COL_VsCharge.���� _
                 Or .Col = COL_VsCharge.ͣ�� _
                 Or .Col = COL_VsCharge.�̶� _
                 Or .Col = COL_VsCharge.���� _
                 Or .Col = COL_VsCharge.�շѷ�ʽ _
                 Then
                .FocusRect = flexFocusHeavy
                txtEdit.Tag = "False"
            Else
                .FocusRect = flexFocusLight
                txtEdit.Tag = "False"
            End If
        Else
            .FocusRect = flexFocusNone
        End If
    End With
End Sub

Private Sub vsCharge_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Col = COL_VsCharge.��Ŀ�� Then
        vsCharge.ComboList = "..."
    Else
        vsCharge.ComboList = ""
    End If
End Sub

Private Sub vsCharge_KeyPress(KeyAscii As Integer)
    If Not mblnEditMode Then Exit Sub
    mblnModify = True
    With vsCharge
       If KeyAscii = vbKeyReturn Then
           KeyAscii = 0
           If .Col = COL_VsCharge.ͣ�� And .Row = .Rows - 1 And Val(.TextMatrix(.Row, COL_VsCharge.ID)) > 0 Then
               .Rows = .Rows + 1
               .Select .Rows - 1, COL_VsCharge.��Ŀ��
           ElseIf .Col = COL_VsCharge.��Ŀ�� Then
               .Select .Row, COL_VsCharge.����
           Else
               Call zlCommFun.PressKey(vbKeyRight)
           End If
       ElseIf .Col = COL_VsCharge.��Ŀ�� And .ComboList = "..." Then
           If KeyAscii = Asc("*") Then
               KeyAscii = 0
               txtEdit.Text = .EditText
               Call SelectCharge(.Row, .Col)
               txtEdit.Tag = False
               txtEdit.Visible = False
           Else
               .ComboList = "" 'ʹ��ť״̬��������״̬
           End If
    
       ElseIf (.Col = COL_VsCharge.�̶� _
               Or .Col = COL_VsCharge.����) _
           And Val(.TextMatrix(.Row, COL_VsCharge.ID)) > 0 _
           And KeyAscii = vbKeySpace Then
           
           If .TextMatrix(.Row, .Col) = "" Then
               .TextMatrix(.Row, .Col) = "��"
           Else
               .TextMatrix(.Row, .Col) = ""
           End If
       ElseIf KeyAscii = vbKeyDelete Then
           Call DeleteCharge
       End If
    End With
End Sub

Private Sub vsCharge_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With vsCharge
            If Col = COL_VsCharge.��Ŀ�� Then
                txtEdit.Text = .EditText
                .EditText = ""
                Call SelectCharge(Row, Col)
                txtEdit.Tag = False
                txtEdit.Visible = False
                .TextMatrix(.Row, COL_VsCharge.����) = 1
                .TextMatrix(.Row, COL_VsCharge.�շѷ�ʽ) = "0-������ȡ"
                .Select .Row, COL_VsCharge.���� - 1
            ElseIf Col = COL_VsCharge.���� Or Col = COL_VsCharge.�շѷ�ʽ Then
                Call zlCommFun.PressKey(vbKeyRight)
            End If
        End With
    End If
End Sub

Private Sub vsCharge_LeaveCell()
    With vsCharge
        On Error Resume Next
        If mblnEditMode Then
            .FocusRect = flexFocusLight
            Call .CellBorder(.GridColor, 0, 0, 0, 0, 0, 0)
        Else
            .FocusRect = flexFocusNone
            Call .CellBorder(.GridColor, 0, 0, 0, 0, 0, 0)
        End If
    End With
End Sub

Private Sub vsCharge_RowColChange()
    On Error GoTo ErrHandle
    
    With vsCharge
        If mblnEditMode Then
            If .Col = COL_VsCharge.��Ŀ�� Or .Col = COL_VsCharge.���� Or .Col = COL_VsCharge.�̶� Or .Col = COL_VsCharge.���� Or .Col = COL_VsCharge.�շѷ�ʽ Then
                '.FocusRect = flexFocusHeavy
                .FocusRect = flexFocusNone
                Call .CellBorder(vbBlue, 1, 1, 1.5, 1.5, 0, 0)
            Else
                .FocusRect = flexFocusLight
                Call .CellBorder(.GridColor, 0, 0, 0, 0, 0, 0)
            End If
            If txtEdit.Tag = "True" Then
                txtEdit.Left = .CellLeft
                txtEdit.Top = .CellTop
                txtEdit.Height = .CellHeight - 12
                txtEdit.Width = .CellWidth - 12
            End If
        Else
            .FocusRect = flexFocusNone
            Call .CellBorder(.GridColor, 0, 0, 0, 0, 0, 0)
        End If
    End With
    Call SetVsfCharge
    
    Exit Sub
ErrHandle:
    If Err.Number = 381 Then Exit Sub
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub DeleteCharge()
    'ɾ���շ���Ŀ
    Dim strName As String
    With vsCharge
        If .Row >= .FixedRows Then
            strName = .TextMatrix(.Row, COL_VsCharge.��Ŀ��)
            If MsgBox("�Ƿ�ɾ����" & strName & "����", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call .RemoveItem(.Row)
                stbThis.Panels(2).Text = "��" & strName & "����ɾ����"
            End If
        End If
    End With
End Sub

Private Sub vsCharge_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsCharge
        If Col = COL_VsCharge.���� Then
            If Not IsNumeric(.EditText) Then
                Cancel = True
                MsgBox "���������֣�"
                Exit Sub
            ElseIf Not Val(.EditText) > 0 And Val(.EditText) <= 9999 Then
                Cancel = True
                MsgBox "������0-9999֮������֣�"
                Exit Sub
            End If
        ElseIf Col = COL_VsCharge.���� Then
        ElseIf Col = COL_VsCharge.�̶� Then
        ElseIf Col = COL_VsCharge.�շѷ�ʽ Then
        ElseIf Col = COL_VsCharge.��Ŀ�� Then
        End If
        
    
    End With

End Sub

Private Sub vsClinic_RowColChange()
    With vsClinic
        If mlngClinicID = Val("" & .TextMatrix(.Row, COL_VsClinic.ID)) Then Exit Sub
        mlngClinicID = Val("" & .TextMatrix(.Row, COL_VsClinic.ID))
        Call vsChargeRef(mlngClinicID)
    End With
End Sub

Private Sub zlRefRecords(Optional lngItem As Long)
    Dim iSubItemIndex As Integer, intCount As Integer, strTemp As String
    Dim rsTemp As ADODB.Recordset, intSelectRow As Integer, blnCharge As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim eTime As Single, sTime As Single
    Dim strDepts As String
    '---------------------------------------------
    '��ȡ������Ŀ
    'ֻ��ȡ ������ĿĿ¼.ִ�п���=1-�������ڿ��� 2-�������ڲ��� ��4-ָ������ �У�ָ������Ϊ�������ID ����Ŀ
    '                 ���� �������Ϊ 1-���3-����סԺ(����Ϊ1-����ʱ) �� 2-סԺ,3-����סԺ(����Ϊ2-סԺʱ)
    '       ������Ŀ����ΪE-���� H-���� I-��ʳ
    '---------------------------------------------
    Err = 0: On Error GoTo ErrHand

    If mlngModul = 1263 Then
        'ҽ������վ
        gstrSql = "Select /*+Rule */ Distinct A.*,b.������Ŀid as �շ� From (" & vbNewLine
        strDepts = mlngDeptID
    Else
        '����ҽ������վ/סԺ��ʿ����վ
        strDepts = mlngDeptID
        If mlngModul = 1262 Then
            '��ʿվ�������������Ӧ����ִ�е���Ŀ
            gstrSql = "Select ����ID From �������Ҷ�Ӧ Where ����ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngDeptID)
            Do While Not rsTemp.EOF
                strDepts = strDepts & "," & rsTemp!����ID
                rsTemp.MoveNext
            Loop
        End If
        gstrSql = "Select /*+Rule */ Distinct A.*,b.������Ŀid as �շ� From (Select i.Id, i.����, i.����, i.�걾��λ, i.���㵥λ, i.��� As �����, k.���� As ���, i.��������, i.ִ��Ƶ��, i.���㷽ʽ, i.�������," & vbNewLine & _
                "       Decode(i.�������, 1, '����', 2, 'סԺ', 3, '�����סԺ', 4, '���', '��ֱ��Ӧ���ڲ���') As �������," & vbNewLine & _
                "       Nvl(i.����ʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) As ����ʱ��, i.վ��, i.�Ƽ�����, zlSpellCode(i.����) As ����" & vbNewLine & _
                "From ������ĿĿ¼ I, ������Ŀ��� K" & vbNewLine & _
                "Where i.��� = k.���� And (Instr(',E,H,I,', i.���) > 0 Or i.���='D' and i.��������='����') And (i.������� = [2] Or i.������� = 3) And" & vbNewLine & _
                "      instr([3],i.ִ�п���)>0 And (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) " & vbNewLine & _
                "Union All" & vbNewLine
        
    End If
    gstrSql = gstrSql & _
            "Select i.Id, i.����, i.����, i.�걾��λ, i.���㵥λ, i.��� As �����, k.���� As ���, i.��������, i.ִ��Ƶ��, i.���㷽ʽ, i.�������," & vbNewLine & _
            "       Decode(i.�������, 1, '����', 2, 'סԺ', 3, '�����סԺ', 4, '���', '��ֱ��Ӧ���ڲ���') As �������," & vbNewLine & _
            "       Nvl(i.����ʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) As ����ʱ��, i.վ��, i.�Ƽ�����, zlSpellCode(i.����) As ����" & vbNewLine & _
            "From ������ĿĿ¼ I, ������Ŀ��� K, ����ִ�п��� J" & vbNewLine & _
            "Where i.��� = k.���� And (Instr(',E,H,I,', i.���) > 0 Or i.���='D' and i.��������='����') And (i.������� = [2] Or i.������� = 3) And" & vbNewLine & _
            "      i.ִ�п��� = 4 And Nvl(j.������Դ,0) <> [2] And i.Id = j.������Ŀid And j.ִ�п���id  In (Select Column_Value From Table(f_Num2list([1]))) And" & vbNewLine & _
            "      (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))) A,(Select ������Ŀid From �����շѹ�ϵ Where ������Դ = [2] And ���ÿ���id  In (Select Column_Value From Table(f_Num2list([1]))) ) B" & vbNewLine & _
            "Where A.id=B.������ĿID(+) Order By ���, ��������, ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strDepts, mint����, IIf(mlngModul = 1262, "12", mint����))
    
    With vsClinic
        .Clear
        Call initVsClinic
        Do While Not rsTemp.EOF
            
            If Val(.TextMatrix(.Rows - 1, COL_VsClinic.ID)) <> 0 Then .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, COL_VsClinic.ID) = "" & rsTemp!ID
            .TextMatrix(.Rows - 1, COL_VsClinic.���) = "" & rsTemp!���
            .TextMatrix(.Rows - 1, COL_VsClinic.����) = "" & rsTemp!����
            .TextMatrix(.Rows - 1, COL_VsClinic.����) = "" & rsTemp!����
            .TextMatrix(.Rows - 1, COL_VsClinic.�걾��λ) = "" & rsTemp!�걾��λ
            .TextMatrix(.Rows - 1, COL_VsClinic.���㵥λ) = "" & rsTemp!���㵥λ
            
            .TextMatrix(.Rows - 1, COL_VsClinic.ִ��Ƶ��) = "" & rsTemp!ִ��Ƶ��

            Select Case Val("" & rsTemp!ִ��Ƶ��)
            Case 0
                .TextMatrix(.Rows - 1, COL_VsClinic.ִ��Ƶ��) = "��ѡƵ��"
            Case 1
                .TextMatrix(.Rows - 1, COL_VsClinic.ִ��Ƶ��) = "һ����"
            Case 2
                .TextMatrix(.Rows - 1, COL_VsClinic.ִ��Ƶ��) = "������"
            End Select
            Select Case Val("" & rsTemp!���㷽ʽ)
            Case 0
                .TextMatrix(.Rows - 1, COL_VsClinic.���㷽ʽ) = "��ȷ��"
            Case 1
                .TextMatrix(.Rows - 1, COL_VsClinic.���㷽ʽ) = "����"
            Case 2
                .TextMatrix(.Rows - 1, COL_VsClinic.���㷽ʽ) = "��ʱ"
            Case 3
                .TextMatrix(.Rows - 1, COL_VsClinic.���㷽ʽ) = "�ƴ�"
            End Select
            Select Case Val("" & rsTemp!�������)
            Case 0
                .TextMatrix(.Rows - 1, COL_VsClinic.�������) = "��������"
            Case 1
                .TextMatrix(.Rows - 1, COL_VsClinic.�������) = "ȡ������"
            End Select
            .TextMatrix(.Rows - 1, COL_VsClinic.�������) = "" & rsTemp!�������
            .TextMatrix(.Rows - 1, COL_VsClinic.վ��) = "" & rsTemp!վ��
            
            Select Case rsTemp!�����
            Case "E"
                intCount = Val("" & rsTemp!��������)
                strTemp = Switch(intCount = 0, "��ͨ", _
                                intCount = 1, "��������", _
                                intCount = 2, "��ҩ����(��ҩ)", _
                                intCount = 3, "��ҩ�巨", _
                                intCount = 4, "��ҩ��(��)��", _
                                intCount = 5, "��������", _
                                intCount = 6, "�ɼ�����", _
                                intCount = 7, "��Ѫ����", _
                                intCount = 8, "��Ѫ;��", _
                                intCount = 9, "��Ѫ�ɼ�")
                .TextMatrix(.Rows - 1, COL_VsClinic.��������) = strTemp
            Case "H"
                If IIf(IsNull(rsTemp!��������), "0", rsTemp!��������) = "1" Then
                    .TextMatrix(.Rows - 1, COL_VsClinic.��������) = "����ȼ�"
                Else
                    .TextMatrix(.Rows - 1, COL_VsClinic.��������) = "������"
                End If
            Case "Z"
                intCount = Val("" & rsTemp!��������)
                strTemp = Switch(intCount = 0, "��ͨ", _
                                intCount = 1, "����", _
                                intCount = 2, "סԺ", _
                                intCount = 3, "ת��", _
                                intCount = 4, "����", _
                                intCount = 5, "��Ժ", _
                                intCount = 6, "תԺ", _
                                intCount = 7, "����", _
                                intCount = 8, "����", _
                                intCount = 9, "����", _
                                intCount = 10, "��Σ", _
                                intCount = 11, "����", _
                                intCount = 12, "��¼�����")
                .TextMatrix(.Rows - 1, COL_VsClinic.��������) = strTemp
            Case Else
                .TextMatrix(.Rows - 1, COL_VsClinic.��������) = "" & rsTemp!��������
            End Select
            .TextMatrix(.Rows - 1, COL_VsClinic.�����) = rsTemp!�����
            
            Select Case Val("" & rsTemp!�Ƽ�����)
            Case 0
                .TextMatrix(.Rows - 1, COL_VsClinic.�Ƽ�����) = "�����Ƽ�"
            Case 1
                .TextMatrix(.Rows - 1, COL_VsClinic.�Ƽ�����) = "���Ƽ�"
            Case 2
                .TextMatrix(.Rows - 1, COL_VsClinic.�Ƽ�����) = "�ֹ��Ƽ�"
            End Select
            
            .TextMatrix(.Rows - 1, COL_VsClinic.����) = "" & rsTemp!����
            .Cell(flexcpForeColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = .ForeColor
            
            '��ͣ����Ŀ��ʾΪ��ɫ
            If Format(rsTemp!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                .Cell(flexcpForeColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = &HFF&
            End If
            
            '�ж��գ���ʾΪ��ɫ
            'gstrSql = "Select ������ĿID From �����շѹ�ϵ Where ������Դ = 2 And ������ĿID=[1] And ���ÿ���ID=[2]"
            'Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val("" & rsTemp!ID), mlngDeptID)
            If Val("" & rsTemp!�շ�) <> 0 Then
                .Cell(flexcpForeColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = vbBlue
            End If

            If mlngClinicID <> 0 And mlngClinicID = rsTemp!ID Then
                intSelectRow = .Rows - 1
            End If
            
            rsTemp.MoveNext
        Loop
        
        If intSelectRow <> 0 Then
            .Select intSelectRow, COL_VsClinic.����
        Else
            .Select .FixedRows, COL_VsClinic.����
        End If
    End With
    
    Err = 0: On Error Resume Next
    If Val(vsClinic.TextMatrix(vsClinic.Rows - 1, COL_VsClinic.ID)) <> 0 Then
        Me.stbThis.Panels(2).Text = "�÷��๲��" & Me.vsClinic.Rows - 1 & "����Ŀ"
    Else
        Call initVsCharge(vsCharge)
        Call initVsCharge(vsCharge1)
        Me.stbThis.Panels(2).Text = ""
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsChargeRef(ByVal lngClinicID As Long)
    Dim strSql As String, rsTemp As ADODB.Recordset
    Dim lngForeColor As Long, curTotal As Currency
    Dim strInfo As String
    On Error GoTo ErrHand
    
    Call initVsCharge(vsCharge)
    Call initVsCharge(vsCharge1)
    
    strInfo = ""
    Me.stbThis.Panels(2).Text = ""
    
    If lngClinicID <= 0 Then Exit Sub
    
    '---�������ķ��ö���
    gstrSql = "select I.ID,R.��鲿λ,R.��鷽��,R.��������,'['||I.����||']'||I.���� as ����,I.���,I.���㵥λ,decode(I.�Ƿ���,1,'���',to_char(P.�۸�)) As �۸�," & _
            "       nvl(R.�շ�����,0) as ����,nvl(R.���ж���,0) as �̶�,nvl(R.������Ŀ,0) as ����," & _
            "Nvl(I.����ʱ��,to_Date('3000-01-01','YYYY-MM-DD')) As ����ʱ��,Nvl(R.�շѷ�ʽ,0) As �շѷ�ʽ " & _
            " from �����շѹ�ϵ R,�շ���ĿĿ¼ I," & _
            "      (Select P.�շ�ϸĿid,sum(P.�ּ�) As �۸�" & _
            "      From �շѼ�Ŀ P " & _
            "      Where P.ִ������<=Sysdate And (P.��ֹ���� Is Null Or P.��ֹ����>=Sysdate)" & _
            IIf(gstrPriceClass = "", " And P.�۸�ȼ� Is Null ", " And P.�۸�ȼ� = [4] ") & _
            "      Group by P.�շ�ϸĿid) P" & _
            " where R.�շ���ĿID=I.ID and I.ID=P.�շ�ϸĿid(+) And R.������Դ = [3] and R.������ĿID=[1] And R.���ÿ���ID=[2]" & _
            " order by nvl(R.������Ŀ,0) ,R.ROWID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngClinicID, mlngDeptID, mint����, gstrPriceClass)
        
    With vsCharge
        curTotal = 0
        Do While Not rsTemp.EOF
            If Val(.TextMatrix(.Rows - 1, COL_VsCharge.ID)) <> 0 Then .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, COL_VsCharge.���) = .Rows - 1
            .TextMatrix(.Rows - 1, COL_VsCharge.ID) = "" & rsTemp!ID
            .TextMatrix(.Rows - 1, COL_VsCharge.��Ŀ��) = "" & rsTemp!����
            .TextMatrix(.Rows - 1, COL_VsCharge.���) = "" & rsTemp!���
            .TextMatrix(.Rows - 1, COL_VsCharge.��λ) = "" & rsTemp!���㵥λ
            .TextMatrix(.Rows - 1, COL_VsCharge.�۸�) = FormatEx(Format("" & rsTemp!�۸�, "0.00"), 2)
            .TextMatrix(.Rows - 1, COL_VsCharge.����) = FormatEx(Format("" & rsTemp!����, "0.00000"), 5)
            .TextMatrix(.Rows - 1, COL_VsCharge.�̶�) = IIf(Val("" & rsTemp!�̶�) = 0, "", "��")
            .TextMatrix(.Rows - 1, COL_VsCharge.����) = IIf(Val("" & rsTemp!����) = 0, "", "��")
            .TextMatrix(.Rows - 1, COL_VsCharge.ͣ��) = IIf(Format(rsTemp!����ʱ��, "YYYY-MM-DD") <> "3000-01-01", "ͣ��", "")
            
            Select Case rsTemp!�շѷ�ʽ
            Case 0
                .TextMatrix(.Rows - 1, COL_VsCharge.�շѷ�ʽ) = "0-������ȡ"
            Case 1
                .TextMatrix(.Rows - 1, COL_VsCharge.�շѷ�ʽ) = "1-�����Թܷ���"
            Case 2
                .TextMatrix(.Rows - 1, COL_VsCharge.�շѷ�ʽ) = "2-һ�η���ֻ��ȡһ��"
            Case 3
                .TextMatrix(.Rows - 1, COL_VsCharge.�շѷ�ʽ) = "3-����ֻ��ȡһ��"
            Case 4
                .TextMatrix(.Rows - 1, COL_VsCharge.�շѷ�ʽ) = "4-����δִ����ȡһ��"
            Case 5
                .TextMatrix(.Rows - 1, COL_VsCharge.�շѷ�ʽ) = "5-����ֻ��ȡһ�Σ��ų�������Ŀ"
            Case 6
                .TextMatrix(.Rows - 1, COL_VsCharge.�շѷ�ʽ) = "6-����δִ����ȡһ�Σ��ų�������Ŀ"
            Case Else
                .TextMatrix(.Rows - 1, COL_VsCharge.�շѷ�ʽ) = "0-������ȡ"
            End Select

            If Format(rsTemp!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                lngForeColor = &HFF&
            Else
                lngForeColor = &H0&
            End If
            .Cell(flexcpForeColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = lngForeColor
            
            curTotal = curTotal + Val("" & rsTemp!�۸�) * Val("" & rsTemp!����)
            rsTemp.MoveNext
        Loop
        If curTotal <> 0 Then strInfo = " " & IIf(mint���� = 2, "����", "����") & "���պϼƣ�" & FormatEx(Format(curTotal, "0.0000"), 4)

    End With
    '---- ���п��ҵ��շѶ���
    gstrSql = "select I.ID,R.��鲿λ,R.��鷽��,R.��������,'['||I.����||']'||I.���� as ����,I.���,I.���㵥λ,decode(I.�Ƿ���,1,'���',to_char(P.�۸�)) As �۸�," & _
            "       nvl(R.�շ�����,0) as ����,nvl(R.���ж���,0) as �̶�,nvl(R.������Ŀ,0) as ����," & _
            "Nvl(I.����ʱ��,to_Date('3000-01-01','YYYY-MM-DD')) As ����ʱ��,Nvl(R.�շѷ�ʽ,0) As �շѷ�ʽ " & _
            " from �����շѹ�ϵ R,�շ���ĿĿ¼ I," & _
            "      (Select P.�շ�ϸĿid,sum(P.�ּ�) As �۸�" & _
            "      From �շѼ�Ŀ P " & _
            "      Where P.ִ������<=Sysdate And (P.��ֹ���� Is Null Or P.��ֹ����>=Sysdate)" & _
            IIf(gstrPriceClass = "", " And P.�۸�ȼ� Is Null ", " And P.�۸�ȼ� = [2] ") & _
            "      Group by P.�շ�ϸĿid) P" & _
            " where R.�շ���ĿID=I.ID and I.ID=P.�շ�ϸĿid(+) And R.������Դ = 0 and R.������ĿID=[1] And R.���ÿ���ID Is Null" & _
            " order by nvl(R.������Ŀ,0) ,R.ROWID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngClinicID, gstrPriceClass)
    With vsCharge1
        Do While Not rsTemp.EOF
            If Val(.TextMatrix(.Rows - 1, COL_VsCharge.ID)) <> 0 Then .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, COL_VsCharge.���) = .Rows - 1
            .TextMatrix(.Rows - 1, COL_VsCharge.ID) = "" & rsTemp!ID
            .TextMatrix(.Rows - 1, COL_VsCharge.��Ŀ��) = "" & rsTemp!����
            .TextMatrix(.Rows - 1, COL_VsCharge.���) = "" & rsTemp!���
            .TextMatrix(.Rows - 1, COL_VsCharge.��λ) = "" & rsTemp!���㵥λ
            .TextMatrix(.Rows - 1, COL_VsCharge.�۸�) = FormatEx(Format("" & rsTemp!�۸�, "0.00"), 2)
            .TextMatrix(.Rows - 1, COL_VsCharge.����) = FormatEx(Format("" & rsTemp!����, "0.00000"), 5)
            .TextMatrix(.Rows - 1, COL_VsCharge.�̶�) = IIf(Val("" & rsTemp!�̶�) = 0, "", "��")
            .TextMatrix(.Rows - 1, COL_VsCharge.����) = IIf(Val("" & rsTemp!����) = 0, "", "��")
            .TextMatrix(.Rows - 1, COL_VsCharge.ͣ��) = IIf(Format(rsTemp!����ʱ��, "YYYY-MM-DD") <> "3000-01-01", "ͣ��", "")
            
            Select Case rsTemp!�շѷ�ʽ
            Case 0
                .TextMatrix(.Rows - 1, COL_VsCharge.�շѷ�ʽ) = "0-������ȡ"
            Case 1
                .TextMatrix(.Rows - 1, COL_VsCharge.�շѷ�ʽ) = "1-�����Թܷ���"
            Case 2
                .TextMatrix(.Rows - 1, COL_VsCharge.�շѷ�ʽ) = "2-һ�η���ֻ��ȡһ��"
            Case 3
                .TextMatrix(.Rows - 1, COL_VsCharge.�շѷ�ʽ) = "3-����ֻ��ȡһ��"
            Case 4
                .TextMatrix(.Rows - 1, COL_VsCharge.�շѷ�ʽ) = "4-����δִ����ȡһ��"
            Case 5
                .TextMatrix(.Rows - 1, COL_VsCharge.�շѷ�ʽ) = "5-����ֻ��ȡһ�Σ��ų�������Ŀ"
            Case 6
                .TextMatrix(.Rows - 1, COL_VsCharge.�շѷ�ʽ) = "6-����δִ����ȡһ�Σ��ų�������Ŀ"
            Case Else
                .TextMatrix(.Rows - 1, COL_VsCharge.�շѷ�ʽ) = "0-������ȡ"
            End Select

            If Format(rsTemp!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                lngForeColor = &HFF&
            Else
                lngForeColor = &H0&
            End If
            .Cell(flexcpForeColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = lngForeColor
            
            curTotal = curTotal + Val("" & rsTemp!�۸�) * Val("" & rsTemp!����)
            rsTemp.MoveNext
        Loop
        If curTotal <> 0 Then strInfo = strInfo & " ���п��Ҷ��պϼƣ�" & FormatEx(Format(curTotal, "0.0000"), 4)
    
    End With
    
    Call SetVsfCharge
    
    '-- �ϼ�
    
    Me.stbThis.Panels(2).Text = "�÷��๲��" & Me.vsClinic.Rows - 1 & "����Ŀ" & strInfo

    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SaveData(ByVal lngDeptID As Long, Optional blnShowInfo As Boolean = True) As Boolean
    '����ҵ������
    'lngDeptID : ����ID
    'blnShowInfo:�Ƿ���ʾ�����ʾ
    If mlngClinicID = 0 Then
        MsgBox "δ��ȷָ��������Ŀ��"
        Exit Function
    End If
    
    'У�����������ȫ��Ϊ����(�൱�ڲ����ײ�)����������ڴ����ֻ�����ұ�����һ������Ҹ��������Ϊ�̶���Ŀ(����ɾ��)��
    Dim bln���ڴ��� As Boolean
    Dim int������ As Integer
    Dim int���������� As Integer
    Dim intRows As Integer
    Dim rs As New ADODB.Recordset
    Dim intCount As Integer
    
    Err = 0: On Error GoTo ErrHand
    With vsCharge
        For intCount = .FixedRows To .Rows - 1
            If .TextMatrix(intCount, COL_VsCharge.����) = "��" Then
                bln���ڴ��� = True
                Exit For
            End If
        Next
        If bln���ڴ��� Then
            For intCount = .FixedRows To .Rows - 1
                If .TextMatrix(intCount, COL_VsCharge.����) <> "��" Then
                    int���������� = intCount
                    int������ = int������ + 1
                    If int������ > 1 Then
                        If blnShowInfo Then MsgBox "��ʾ��ֻ������һ�����"
                        Exit Function
                    End If
                End If
            Next
            
            If int������ = 1 Then
                If .TextMatrix(int����������, COL_VsCharge.�̶�) <> "��" Then
                    If blnShowInfo Then MsgBox "��ʾ����" & int���������� & "�����������Ϊ�̶���Ŀ��"
                    Exit Function
                End If
            End If
            If int������ = 0 Then
                If blnShowInfo Then MsgBox "��ʾ������Ҫ��һ�����"
                Exit Function
            End If
        End If
    End With
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim lngChargeID As Long
    '�������ļ۸��Ƿ���ڶ��������Ŀ�����������ʾ�����ܱ���
    If bln���ڴ��� Then
        lngChargeID = Val(Me.vsCharge.TextMatrix(int����������, COL_VsCharge.ID))
        gstrSql = "Select Id From �շѼ�Ŀ Where �շ�ϸĿid=[1] And ִ������ <= SYSDATE AND (��ֹ���� > SYSDATE OR ��ֹ���� IS NULL) " & _
                IIf(gstrPriceClass = "", " And �۸�ȼ� Is Null ", " And �۸�ȼ� = [2] ")
        
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngChargeID, gstrPriceClass)
        If rs.RecordCount > 1 Then
            If blnShowInfo Then MsgBox "��ʾ������ļ۸���ڶ��������Ŀ�����ܱ��档"
            Exit Function
        End If
        rs.Close
    End If
    
    Dim strCharges As String   '�����շ�ϸĿID,���ڼ���ظ���Ŀ
    Dim strContent() As String '���� zl_�����շ�_UPDATE Ҫ�õ� �շ����� ����
    strCharges = "": ReDim strContent(0) As String
    
    With Me.vsCharge
        For intCount = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(intCount, COL_VsCharge.��Ŀ��)) <> "" And Val(.TextMatrix(intCount, COL_VsCharge.ID)) <> 0 Then
                If Not IsNumeric(Nvl(.TextMatrix(intCount, COL_VsCharge.����), "X")) Then
                    If blnShowInfo Then MsgBox intCount & IIf(.TextMatrix(intCount, COL_VsCharge.����) = "", "�е���������Ϊ��", "������д���֡�")
                    Exit Function
                End If
                
                '������0.000��
                If Int(.TextMatrix(intCount, COL_VsCharge.����)) = 0 And .TextMatrix(intCount, COL_VsCharge.�̶�) = "��" Then
                    .TextMatrix(intCount, COL_VsCharge.�̶�) = ""
                    MsgBox intCount & "�е�����Ϊ0,������Ϊ�̶���,���Զ�����."
                    Exit Function
                End If
            
                If InStr(1, strCharges & ";", ";" & Val(.TextMatrix(intCount, COL_VsCharge.ID)) & ";") > 0 Then
                    If blnShowInfo Then MsgBox intCount & "���շ���Ŀ��ǰ����շ���Ŀ���ظ���"
                    Exit Function
                End If
                strCharges = strCharges & ";" & Val(.TextMatrix(intCount, COL_VsCharge.ID))
                If strContent(UBound(strContent)) <> "" Then ReDim Preserve strContent(UBound(strContent) + 1)
                '��"|"�ָ��������շ�����,ÿ����¼��"������ĿID^����^�̶�^����^����^��λ^��鷽��^�շѷ�ʽ"��֯
                strContent(UBound(strContent)) = Val(.TextMatrix(intCount, COL_VsCharge.ID)) & "^" & Val(.TextMatrix(intCount, COL_VsCharge.����)) & "^" & IIf(Trim(.TextMatrix(intCount, COL_VsCharge.�̶�)) = "", 0, 1) & "^" & IIf(Trim(.TextMatrix(intCount, COL_VsCharge.����)) = "", 0, 1) & "^0^^ " & Val(Mid(.TextMatrix(intCount, COL_VsCharge.�շѷ�ʽ), 1, 1))
            End If
        Next
    End With

    'If gstrSql <> "" Then gstrSql = Mid(gstrSql, 2)
    
    Dim lngCount As Long ' �ܸ���
    Dim lngLoop As Long, lngEndloop As Long
    Dim strItem As String, blnBeginTrans As Boolean, i As Integer
    
    lngCount = UBound(strContent)
    lngEndloop = 0
    gcnOracle.BeginTrans
    blnBeginTrans = True
    
    For lngLoop = 0 To lngCount
        
        strItem = strItem & "|" & strContent(lngLoop)
        If i >= 40 Then
            strItem = Mid(strItem, 2)
            
            gstrSql = "zl_�����շ�_UPDATE(" & mlngClinicID & "," & mlng�Ƽ����� & ",'" & strItem & "'," & IIf(lngEndloop = 0, 1, 0) & "," & IIf(lngDeptID = 0, "Null", lngDeptID) & "," & mint���� & ")"
            Err = 0: On Error GoTo ErrHand
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            i = 0: strItem = ""
            lngEndloop = lngEndloop + 1
        End If
        i = i + 1
    Next
    
    If Left(strItem, 1) = "|" Then
        strItem = Mid(strItem, 2)
        
        gstrSql = "zl_�����շ�_UPDATE(" & mlngClinicID & "," & mlng�Ƽ����� & ",'" & strItem & "'," & IIf(lngEndloop = 0, 1, 0) & "," & IIf(lngDeptID = 0, "Null", lngDeptID) & "," & mint���� & ")"
        Err = 0: On Error GoTo ErrHand
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    End If
    
    If lngLoop = 0 Then '11303 ����ȫ��ɾ�����յ��շ���Ŀ
        gstrSql = "zl_�����շ�_UPDATE(" & mlngClinicID & "," & mlng�Ƽ����� & ",'',1," & IIf(lngDeptID = 0, "Null", lngDeptID) & "," & mint���� & ")"
        Err = 0: On Error GoTo ErrHand
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    End If
    
    gcnOracle.CommitTrans
    blnBeginTrans = False
    SaveData = True
    If blnShowInfo Then MsgBox "�շѶ��ձ���ɹ���"
    
    Exit Function

ErrHand:
    If blnBeginTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SelectCharge(ByVal Row As Long, ByVal Col As Long)
    '��ȡ�շ���Ŀ
    
    'ֻ��ȡ �շ���ĿĿ¼.�������Ϊ 2��3 �����õ���Ŀ��
    '
    Dim rsTmp As New ADODB.Recordset
    Dim strSql   As String, strInput As String
    Dim vRect As RECT, blnCanel As Boolean
    Dim i As Integer
    On Error GoTo ErrHandle
    
    If Col = COL_VsCharge.��Ŀ�� Then
        '��ȡ��Ŀ
        '--------------------------------------------------------------------------------------
            strInput = DelInvalidChar(UCase(Trim(txtEdit)))
            If InStr(strInput, " ") > 0 Then
                strInput = Trim(Split(strInput, " ")(0))
            End If
            strSql = "Select distinct i.* " & vbNewLine & _
                    "From (Select Distinct i.Id, i.����, i.����, i.���, i.����, i.���㵥λ," & vbNewLine & _
                    "                       Decode(Nvl(i.�Ƿ���, 0), 0, LTrim(RTrim(To_Char(Nvl(d.�ּ�, 0), '9999999990.0000')))," & vbNewLine & _
                    "                               Decode(Instr('4567', ���), 0, LTrim(RTrim(To_Char(Nvl(d.ȱʡ�۸�, 0), '9999999990.0000'))), 'ʱ��')) As �ۼ�" & vbNewLine & _
                    "       From �շ���ĿĿ¼ I, �շѼ�Ŀ D" & vbNewLine & _
                    "       Where i.Id = d.�շ�ϸĿid(+) And" & vbNewLine & _
                    "             (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And d.ִ������ <= Sysdate And" & vbNewLine & _
                    "             (d.��ֹ���� > Sysdate Or d.��ֹ���� Is Null) And (i.������� = [3] Or i.������� = 3) " & vbNewLine & _
                    IIf(gstrPriceClass = "", " And D.�۸�ȼ� Is Null ", " And D.�۸�ȼ� = [4] ") & vbNewLine & _
                    "      ) I, �շ���Ŀ���� N" & vbNewLine & _
                    "Where i.Id = n.�շ�ϸĿid And Rownum<2000 "

            If strInput <> "" Then
                strSql = strSql & " and (I.���� like [1] " & _
                        "           or N.���� like [2] " & _
                        "           or N.���� like [2])"

            
            End If
            With vsCharge
                txtEdit.Left = .CellLeft
                txtEdit.Top = .CellTop
                txtEdit.Height = .CellHeight - 12
                txtEdit.Width = .CellWidth - 12
            End With

            vRect = zlControl.GetControlRect(txtEdit.hWnd)
            Set rsTmp = New ADODB.Recordset
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "������Ŀ", False, "", "ѡ�������Ŀ", False, False, True, _
                                                 vRect.Left, vRect.Top, txtEdit.Height, blnCanel, True, True, _
                                                 strInput & "%", gstrMatch & strInput & "%", mint����, gstrPriceClass)
            If Not blnCanel And rsTmp.State <> 0 Then
                If Not rsTmp.EOF Then
                    With vsCharge
                        .EditText = "[" & Trim("" & rsTmp.Fields("����") & "]" & rsTmp.Fields("����"))
                        .TextMatrix(.Row, COL_VsCharge.��Ŀ��) = "[" & Trim("" & rsTmp.Fields("����") & "]" & rsTmp.Fields("����"))
                        .TextMatrix(.Row, COL_VsCharge.ID) = "" & rsTmp.Fields("ID")
                        .TextMatrix(.Row, COL_VsCharge.���) = "" & rsTmp.Fields("���")
                        .TextMatrix(.Row, COL_VsCharge.�۸�) = "" & rsTmp.Fields("�ۼ�")
                        .TextMatrix(.Row, COL_VsCharge.��λ) = "" & rsTmp.Fields("���㵥λ")
                    End With
                End If
                Set rsTmp = Nothing
            End If
            txtEdit = ""
    End If
    Call zlCommFun.PressKey(vbKeyRight)
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub

Private Sub ExecuteFind(Optional ByVal blnNext As Boolean)
'���ܣ�����(��һ��)����
'������blnNext=�Ƿ������һ��
    Static blnReStart As Boolean
    Dim lngRow As Long, lngStart As Long
    Dim strFind As String, blnHave As Boolean
    strFind = IIf(gstrMatch = "%", "*", "") & DelInvalidChar(Trim(txtFind.Text))
            
    '��ʼ������
    With vsClinic
        If blnNext Then
            If .Row + 1 <= .Rows - 1 Then
                lngStart = .Row + 1
            Else
                MsgBox IIf(blnNext, "������", "") & "�Ҳ���������������Ŀ��", vbInformation, gstrSysName
                Exit Sub
            End If
        Else
            lngStart = .FixedRows
        End If
        For lngRow = lngStart To .Rows - 1

            If UCase(.TextMatrix(lngRow, COL_VsClinic.����)) Like UCase(strFind) & "*" _
               Or UCase(.TextMatrix(lngRow, COL_VsClinic.����)) Like UCase(strFind) & "*" _
               Or UCase(.TextMatrix(lngRow, COL_VsClinic.����)) Like UCase(strFind) & "*" _
            Then Exit For
        Next
        
        If lngRow <= .Rows - 1 Then
            '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
            .Select lngRow, COL_VsClinic.����
            .ShowCell lngRow, COL_VsClinic.����
        Else
            MsgBox IIf(blnNext, "������", "") & "�Ҳ���������������Ŀ��", vbInformation, gstrSysName
        End If
    End With


End Sub

Private Sub DeptCopy()
    '���Ʊ����ҵĸ���Ŀ���յ���������
    Dim strSql As String
    Dim rsDept As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInfo As String
    Dim varDept As Variant, strReturn  As String, strLine As String, i As Integer
    
    On Error GoTo ErrHandle
    If mint���� = 2 Then
        strSql = "Select Distinct a.����, a.����, ID " & vbNewLine & _
            "From ���ű� A, ������Ա B, �ϻ���Ա�� C, ��������˵�� D" & vbNewLine & _
            "Where a.Id = b.����id And b.��Աid = c.��Աid And a.Id = d.����id And (d.������� = 2 Or d.�������=3) And d.�������� = '����' And A.ID<>[1] And c.�û��� = User"
    Else
        strSql = "Select Distinct a.����, a.����, ID " & vbNewLine & _
            "From ���ű� A, ������Ա B, �ϻ���Ա�� C, ��������˵�� D" & vbNewLine & _
            "Where a.Id = b.����id And b.��Աid = c.��Աid And a.Id = d.����id And (d.������� = 1 Or d.�������=3) And Instr('����,���,����,����,Ӫ��', d.��������) > 0 And A.ID<>[1] And c.�û��� = User"
    
    End If
    Set rsDept = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngDeptID)
    strReturn = frmSelCur.ShowCurrSel(Me, rsDept, "����,1200,0,2;����,1800,0,2;ID,0,1,2", "ѡ��" & IIf(mint���� = 2, "����", "����"), True, , , 5000, True)
    If strReturn = "" Then Exit Sub
    varDept = Split(strReturn, "|")
    
    strInfo = ""
    For i = LBound(varDept) To UBound(varDept)
        '�����Ƿ������˶��գ�û�в��ܸ���
        strLine = varDept(i)
        If UBound(Split(strLine, ",")) = 2 Then
            strSql = "Select �շ���ĿID From �����շѹ�ϵ Where ������Դ=[2] And ���ÿ���ID=[1] and ������ĿID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CLng(Split(strLine, ",")(2)), mlngClinicID, mint����)
            If rsTmp.EOF Then
               Call SaveData(CLng(Split(strLine, ",")(2)), False)
            Else
               strInfo = IIf(strInfo = "", "", vbNewLine) & "" & Split(strLine, ",")(0) & " " & Split(strLine, ",")(1) & " ����Ŀ�Ѿ��趨�˷��ã�"
            End If
        End If
    Next
    If strInfo <> "" Then
        MsgBox strInfo
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CheckIsNoCharge()
    '�����δ���չ����ã����������п��ҵĶ��ա�
    Dim blnNoCharge As Boolean, intRow As Integer, intCol As Integer
    blnNoCharge = True
    With vsCharge
        For intRow = .FixedRows To .Rows - 1
             If Val(.TextMatrix(intRow, COL_VsCharge.ID)) <> 0 Then
                blnNoCharge = False
                Exit For
             End If
        Next
    End With
    If blnNoCharge Then
        If Me.vsCharge1.Rows < 2 Then
            Exit Sub
        ElseIf Me.vsCharge1.Rows = 2 Then
            If Me.vsCharge1.TextMatrix(1, COL_VsCharge.��Ŀ��) = "" Then Exit Sub
        End If
        If MsgBox("��ǰ�������շ���Ŀ���Ƿ��Զ�����ȫԺ���շ���Ŀ��", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        mblnModify = True
        With vsCharge1
            For intRow = .FixedRows To .Rows - 1
                If Val(.TextMatrix(intRow, COL_VsCharge.ID)) <> 0 Then
                   If Val(vsCharge.TextMatrix(vsCharge.Rows - 1, COL_VsCharge.ID)) <> 0 Then vsCharge.Rows = vsCharge.Rows + 1
                   For intCol = .FixedCols To .Cols - 1
                    vsCharge.TextMatrix(vsCharge.Rows - 1, intCol) = .TextMatrix(intRow, intCol)
                   Next
                End If
            Next
        End With
    End If
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
    If Me.vsClinic.Rows <= 1 Then Exit Sub
    
    '-------------------------------------------------
    '�������ݱ��
    'If zlReportToVSFlexGrid(Me.vsfPrint, Me.vsClinic) = False Then Exit Sub
    
    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vsClinic
    objPrint.Title.Text = "Ŀ¼"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub SetVsfCharge()
    With vsCharge
        .Cell(flexcpBackColor, 0, COL_VsCharge.���, .Rows - 1, COL_VsCharge.���) = &H8000000F
        .Cell(flexcpBackColor, 0, COL_VsCharge.���, .Rows - 1, COL_VsCharge.�۸�) = &H8000000F
        .Cell(flexcpBackColor, 0, COL_VsCharge.ͣ��, .Rows - 1, COL_VsCharge.ͣ��) = &H8000000F
    End With
End Sub
