VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmMicrobeAntiRef 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ϸ�������زο�"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11865
   Icon            =   "frmMicrobeAntiRef.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   11865
   StartUpPosition =   2  '��Ļ����
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   5685
      Left            =   165
      TabIndex        =   11
      Top             =   135
      Width           =   3030
      _Version        =   589884
      _ExtentX        =   5345
      _ExtentY        =   10028
      _StockProps     =   0
      BorderStyle     =   1
   End
   Begin VB.PictureBox picVfg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3270
      Left            =   3255
      ScaleHeight     =   3240
      ScaleWidth      =   8085
      TabIndex        =   14
      Top             =   405
      Width           =   8115
      Begin VSFlex8Ctl.VSFlexGrid vfgList 
         Height          =   3795
         Left            =   45
         TabIndex        =   15
         Top             =   45
         Width           =   8460
         _cx             =   14922
         _cy             =   6694
         Appearance      =   0
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
         BackColorFixed  =   15790320
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
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
   Begin VB.PictureBox picEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   3330
      ScaleHeight     =   1905
      ScaleWidth      =   8430
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3735
      Width           =   8460
      Begin VB.ComboBox cbo��� 
         Height          =   300
         Index           =   2
         Left            =   6270
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   900
         Width           =   1215
      End
      Begin VB.ComboBox cbo��� 
         Height          =   300
         Index           =   1
         Left            =   3600
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   900
         Width           =   1215
      End
      Begin VB.ComboBox cbo��� 
         Height          =   300
         Index           =   0
         Left            =   930
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   900
         Width           =   1215
      End
      Begin VB.CommandButton cmd������ 
         Caption         =   "��"
         Height          =   300
         Left            =   7845
         TabIndex        =   13
         Top             =   180
         Width           =   300
      End
      Begin VB.TextBox txt������ 
         Height          =   300
         Left            =   930
         TabIndex        =   1
         Top             =   195
         Width           =   6900
      End
      Begin VB.ComboBox cbo�жϷ�ʽ 
         Height          =   300
         Left            =   6270
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   540
         Width           =   1890
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         ItemData        =   "frmMicrobeAntiRef.frx":000C
         Left            =   930
         List            =   "frmMicrobeAntiRef.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   540
         Width           =   1215
      End
      Begin VB.TextBox txt�ο�ֵ 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   2760
         MaxLength       =   13
         TabIndex        =   3
         Top             =   540
         Width           =   900
      End
      Begin VB.TextBox txt�ο�ֵ 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   3900
         MaxLength       =   13
         TabIndex        =   4
         Top             =   540
         Width           =   900
      End
      Begin VB.TextBox txt��ע 
         Height          =   300
         Left            =   930
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1245
         Width           =   7230
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���ڲο�"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   5070
         TabIndex        =   21
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ڲο���Χ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   2325
         TabIndex        =   19
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���ڲο�"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   105
         TabIndex        =   17
         Top             =   960
         Width           =   720
      End
      Begin XtremeCommandBars.CommandBars cbsThis 
         Left            =   180
         Top             =   1410
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
      Begin VB.Label lbl������ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   105
         TabIndex        =   12
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lbl��ʽ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ο��жϷ�ʽ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5070
         TabIndex        =   10
         Top             =   600
         Width           =   1080
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ҩ������"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   105
         TabIndex        =   9
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lbl�ο� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ο�           ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2340
         TabIndex        =   8
         Top             =   600
         Width           =   1530
      End
      Begin VB.Label lbl��ע 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ע"
         Height          =   180
         Left            =   105
         TabIndex        =   7
         Top             =   1305
         Width           =   360
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmMicrobeAntiRef.frx":0010
      Left            =   2745
      Top             =   45
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmMicrobeAntiRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngϸ��id As Long
Private mlngItemID As Long '�ϴ�ѡ��ķ���

Private Enum mCol
    ͼ�� = 0: ����Id: ���: ����: Ӣ��
    ID = 0: ����: ������: Ӣ����: ҩ������: �ο���ֵ: �ο���ֵ: �ο�: �жϷ�ʽ: ��ע: �ؼ���: ��ֵ���: �м���: ��ֵ���
End Enum

Private Const Dkp_ID_Rpt As Integer = 1
Private Const Dkp_ID_vfg As Integer = 2
Private Const Dkp_ID_Edit As Integer = 3
Private cbrControl As CommandBarControl
Private mblnEdit As Boolean

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cbo�жϷ�ʽ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngCurRow As Long, lngRow As Long, lngCol As Long
    Dim str������ As String, rsTmp As ADODB.Recordset, strSQL As String
    Dim str�ؼ��� As String
    
    On Error GoTo errHandle
    With Me.vfgList
        Select Case Control.ID
        Case conMenu_Edit_NewItem
            .Rows = .Rows + 1: .Row = .Rows - 1
        Case conMenu_Edit_Delete
            str�ؼ��� = .TextMatrix(.Row, mCol.�ؼ���)
            If str�ؼ��� <> "" Then
                strSQL = "Zl_����ϸ�������زο�_Edit(2," & mlngϸ��id & "," & mlngItemID & "," & Split(str�ؼ���, ",")(0) & "," & Split(str�ؼ���, ",")(1) & ")"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            End If
            If .Row = .Rows - 1 Then
                .Rows = .Rows - 1: .Row = .Rows - 1
            Else
                lngCurRow = .Row
                For lngRow = lngCurRow To .Rows - 2
                    For lngCol = 0 To .Cols - 1
                        .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow + 1, lngCol)
                    Next
                Next
                .Rows = .Rows - 1
            End If
        Case conMenu_Edit_Adjust
            If Val(txt������.Tag) <> 0 Then
                For lngRow = .FixedRows To .Rows - 1
                    If lngRow <> .Row Then
                        If .TextMatrix(lngRow, mCol.ID) = Val(txt������.Tag) And .TextMatrix(lngRow, mCol.ҩ������) = Me.cbo����.Text Then
                            MsgBox "������ͬ��¼����,���ܸ���!", vbQuestion, Me.Caption
                            Exit Sub
                        End If
                    End If
                Next
            End If
            
            If Me.cbo����.Text = "" Then
                MsgBox "��������һ��ҩ�������󣬲��ܸ��²�����", vbInformation, Me.Caption
                Me.cbo����.SetFocus
            End If
            
            .TextMatrix(.Row, mCol.ID) = Val(txt������.Tag)
            .TextMatrix(.Row, mCol.����) = ""
            .TextMatrix(.Row, mCol.������) = ""
            .TextMatrix(.Row, mCol.Ӣ����) = ""
            
            If Val(txt������.Tag) <> 0 Then
                strSQL = "Select B.����, B.������, B.Ӣ���� From �����ÿ����� B Where ID=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(txt������.Tag))
                Do Until rsTmp.EOF
                    .TextMatrix(.Row, mCol.����) = "" & rsTmp!����
                    .TextMatrix(.Row, mCol.������) = "" & rsTmp!������
                    .TextMatrix(.Row, mCol.Ӣ����) = "" & rsTmp!Ӣ����
                    rsTmp.MoveNext
                Loop
            End If
            
            If Me.cbo����.ListIndex = -1 Then
                .TextMatrix(.Row, mCol.ҩ������) = ""
            Else
                .TextMatrix(.Row, mCol.ҩ������) = Me.cbo����.Text
                .TextMatrix(.Row, mCol.�ؼ���) = .TextMatrix(.Row, mCol.ID) & "," & Me.cbo����.ListIndex + 1
            End If
            
            .TextMatrix(.Row, mCol.�ο���ֵ) = IIf(IsNumeric(txt�ο�ֵ(0)), Me.txt�ο�ֵ(0), "")
            .TextMatrix(.Row, mCol.�ο���ֵ) = IIf(IsNumeric(txt�ο�ֵ(1)), Me.txt�ο�ֵ(1), "")
            
            If .TextMatrix(.Row, mCol.�ο���ֵ) = "" Or "" & .TextMatrix(.Row, mCol.�ο���ֵ) = "" Then
                .TextMatrix(.Row, mCol.�ο�) = FormatDecimal(.TextMatrix(.Row, mCol.�ο���ֵ)) & FormatDecimal(.TextMatrix(.Row, mCol.�ο���ֵ))
            Else
                .TextMatrix(.Row, mCol.�ο�) = FormatDecimal(.TextMatrix(.Row, mCol.�ο���ֵ)) & "��" & FormatDecimal(.TextMatrix(.Row, mCol.�ο���ֵ))
            End If

            If Me.cbo�жϷ�ʽ.ListIndex = -1 Then
                .TextMatrix(.Row, mCol.�жϷ�ʽ) = ""
            Else
                .TextMatrix(.Row, mCol.�жϷ�ʽ) = Me.cbo�жϷ�ʽ.Text
            End If
            .TextMatrix(.Row, mCol.��ע) = DelInvalidChar(Trim(Me.txt��ע.Text), "'")
            .TextMatrix(.Row, mCol.��ֵ���) = cbo���(0).Text
            .TextMatrix(.Row, mCol.�м���) = cbo���(1).Text
            .TextMatrix(.Row, mCol.��ֵ���) = cbo���(2).Text
            mblnEdit = True
        Case conMenu_Edit_Save
            Call zlSaveData
            Call initVfg(mlngItemID)
        End Select
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_Edit_Delete, conMenu_Edit_Adjust: Control.Enabled = Me.vfgList.Row >= Me.vfgList.FixedRows
    Case conMenu_Edit_Save: Control.Enabled = mblnEdit
    End Select

End Sub

Private Sub zlSaveData()
    Dim lngRow As Long, strSQL As String, strDelSQL As String, str�ؼ��� As String
    Dim str�� As String, str�� As String, str�� As String
    Dim strValue As String
    
    On Error GoTo errHandle
    With vfgList
        For lngRow = .FixedCols To .Rows - 1
            If Val(.TextMatrix(lngRow, mCol.ID)) <> 0 Then
                strSQL = "Zl_����ϸ�������زο�_Edit(1," & mlngϸ��id & "," & mlngItemID & "," & Val(.TextMatrix(lngRow, mCol.ID)) & "," & _
                         Getҩ������(.TextMatrix(lngRow, mCol.ҩ������)) & ","
                If IsNumeric(.TextMatrix(lngRow, mCol.�ο���ֵ)) = True Then
                    
                    strValue = .TextMatrix(lngRow, mCol.�ο���ֵ)
                    If Val(strValue) > 999999999 Or Val(Val(strValue) * 10000) - Int(Val(Val(strValue) * 10000)) > 0 Then
                        MsgBox "��" & lngRow & "�ο�ֵ̫��򾫶�̫�ߣ�", vbInformation, gstrSysName
                        Me.txt������.SetFocus: Exit Sub
                    End If
                    strSQL = strSQL & strValue & ","
                Else
                    strSQL = strSQL & "Null,"
                End If
                
                If IsNumeric(.TextMatrix(lngRow, mCol.�ο���ֵ)) = True Then
                    
                    strValue = .TextMatrix(lngRow, mCol.�ο���ֵ)
                    If Val(strValue) > 999999999 Or Val(Val(strValue) * 10000) - Int(Val(Val(strValue) * 10000)) > 0 Then
                        MsgBox "��" & lngRow & "�ο�ֵ̫��򾫶�̫�ߣ�", vbInformation, gstrSysName
                        Me.txt������.SetFocus:  Exit Sub
                    End If
                    strSQL = strSQL & strValue & ","
                Else
                    strSQL = strSQL & "Null,"
                End If
                strSQL = strSQL & Get�жϷ�ʽ(.TextMatrix(lngRow, mCol.�жϷ�ʽ)) & ",'" & .TextMatrix(lngRow, mCol.��ע) & "'"
                str�ؼ��� = .TextMatrix(lngRow, mCol.�ؼ���)
                str�� = .TextMatrix(lngRow, mCol.��ֵ���)
                str�� = .TextMatrix(lngRow, mCol.�м���)
                str�� = .TextMatrix(lngRow, mCol.��ֵ���)
                strSQL = strSQL & ",'" & str�� & "','" & str�� & "','" & str�� & "')"
                
                If Right(Trim(str�ؼ���), 1) <> "," Then
                    If str�ؼ��� <> "" Then
                        strDelSQL = "Zl_����ϸ�������زο�_Edit(2," & mlngϸ��id & "," & mlngItemID & "," & Split(str�ؼ���, ",")(0) & "," & Split(str�ؼ���, ",")(1) & ")"
                        Call zlDatabase.ExecuteProcedure(strDelSQL, Me.Caption)
                    End If
                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                End If
            End If
        Next
    End With
    mblnEdit = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub cmd������_Click()
    Dim rsTemp As ADODB.Recordset
    Dim blnReturn As Boolean
    On Error GoTo errHandle
    gstrSql = "Select B.ID, B.����, B.������, B.Ӣ����, ҩ������" & vbNewLine & _
            "From �����ÿ����� B, ���鿹������ҩ A" & vbNewLine & _
            "Where A.������id = B.Id  And A.�����ط���id = [1]"
    'Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngItemID)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "������ѡ��", True, "", "��ѡ������", False, False, False, 0, 0, 0, blnReturn, True, False, mlngItemID)
    If blnReturn = False Then
        If rsTemp.RecordCount > 0 Then
            txt������.Tag = rsTemp!ID
            txt������.Text = "(" & rsTemp!���� & ")" & rsTemp!������
            lbl������.Tag = txt������.Text '���ڻָ���ʾ
        Else
            txt������.Text = lbl������.Tag
            zlControl.TxtSelAll txt������
            Exit Sub
        End If
    End If
    txt������.SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = Dkp_ID_vfg Then
        Item.Handle = picVfg.hWnd
    ElseIf Item.ID = Dkp_ID_Rpt Then
        Item.Handle = rptList.hWnd
    ElseIf Item.ID = Dkp_ID_Edit Then
        Item.Handle = picEdit.hWnd
    End If
End Sub

Private Sub Form_Load()
    Call initDockPane
    Call initEdit
    Call initRpt
    
    '�ڲ��˵�����������
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set cbsThis.Icons = zlcommfun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
    End With
    Me.cbsThis.EnableCustomization False
    
    Me.cbsThis.ActiveMenuBar.Title = "�˵�"
    Me.cbsThis.ActiveMenuBar.Position = xtpBarBottom
    Me.cbsThis.ActiveMenuBar.EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    With Me.cbsThis.ActiveMenuBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "��������"): cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ������"): cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����"): cbrControl.Style = xtpButtonIconAndCaption
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "���µ��ο�ֵ�б���"): cbrControl.Flags = xtpFlagRightAlign: cbrControl.Style = xtpButtonIconAndCaption
    End With
    
    If Me.rptList.Tag = "Unload" Then
        MsgBox "�������á�ҩ�����鿹�����顱����ʹ�ô˹��ܣ�", vbInformation, Me.Caption
        Unload Me
    End If
    
End Sub



Private Sub picVfg_Resize()
    With vfgList
        .Left = picVfg.ScaleLeft
        .Width = picVfg.ScaleWidth
        .Height = picVfg.ScaleHeight
        .Top = picVfg.ScaleTop
    End With
End Sub

Private Sub rptList_SelectionChanged()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lng����ID As Long
    
      'չ��ѡ����
    With rptList
        If .FocusedRow Is Nothing And .Rows.count > 0 Then
            If .Rows(0).GroupRow Then
                Set .FocusedRow = .Rows(0).Childs(0)
            Else
                Set .FocusedRow = .Rows(0)
            End If
        End If
        If .FocusedRow Is Nothing Then Exit Sub
    End With
    
    If Not rptList.FocusedRow.GroupRow Then
        lng����ID = Val(rptList.FocusedRow.Record(mCol.����Id).Value)
        mlngItemID = lng����ID
        Call initVfg(lng����ID)
     End If

End Sub

Private Sub txt�ο�ֵ_GotFocus(Index As Integer)
    Me.txt�ο�ֵ(Index).SelStart = 0: Me.txt�ο�ֵ(Index).SelLength = 1000
    Call zlcommfun.OpenIme(False)
End Sub

Private Sub txt�ο�ֵ_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or InStr(".", Chr(KeyAscii)) > 0 Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub initEdit()
    Dim i As Integer
    cbo����.Clear
    cbo����.AddItem "MIC", 0
    cbo����.AddItem "DISK", 1
    cbo����.AddItem "K-B", 2
    cbo����.ListIndex = 0
    
    cbo�жϷ�ʽ.Clear
    cbo�жϷ�ʽ.AddItem "�ο�ֵ����", 0
    cbo�жϷ�ʽ.AddItem "�����ο�ֵ", 1
    cbo�жϷ�ʽ.ListIndex = 1
    
    For i = 0 To 2
        cbo���(i).Clear
        cbo���(i).AddItem "R-��ҩ"
        cbo���(i).AddItem "I-�н�"
        cbo���(i).AddItem "S-����"
        cbo���(i).ListIndex = 1
    Next
End Sub

Private Sub initDockPane()
    Dim paneRpt As Pane, paneVfg As Pane, paneEdit As Pane
    
    Me.dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
    
    Me.dkpMain.Options.HideClient = True
    
    Set paneRpt = Me.dkpMain.CreatePane(Dkp_ID_Rpt, 90, 190, DockLeftOf, Nothing)
    paneRpt.Title = "�����ط���"
    paneRpt.Handle = Me.rptList.hWnd
    paneRpt.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set paneVfg = Me.dkpMain.CreatePane(Dkp_ID_vfg, 230, 300, DockRightOf, paneRpt)
    paneVfg.Title = "�ο�ֵ�б�"
    paneVfg.Handle = Me.picVfg.hWnd
    paneVfg.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set paneEdit = Me.dkpMain.CreatePane(Dkp_ID_Edit, 100, 180, DockBottomOf, paneVfg)
    paneEdit.Title = ""
    paneEdit.Handle = Me.picEdit.hWnd
    paneEdit.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
End Sub

Private Sub initRpt()
    Dim rptCol As ReportColumn
    Dim rptRcd As ReportRecord
    Dim rptItem As ReportRecordItem
    Dim rptRow As ReportRow
    
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Long
    
    On Error GoTo errHandle
    With rptList
        .Columns.DeleteAll
        Set rptCol = .Columns.Add(mCol.ͼ��, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.����Id, "����id", 0, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.���, "���", 60, True): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.����, "����", 80, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.Ӣ��, "Ӣ��", 80, True): rptCol.Editable = False: rptCol.Groupable = False
    
        .Records.DeleteAll '���ԭ�б�
        strSQL = "Select B.id,B.����, B.����, B.Ӣ�� From ���鿹������ B, ����ϸ�������� A Where A.�����ط���id = B.ID And A.ϸ��id = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngϸ��id)
        If rsTmp.EOF Then
            Me.rptList.Tag = "Unload"
        Else
            Me.rptList.Tag = ""
        End If
        
        Do Until rsTmp.EOF
        
            Set rptRcd = Me.rptList.Records.Add()
            
            Set rptItem = rptRcd.AddItem(""): rptItem.Focusable = False
            Set rptItem = rptRcd.AddItem(CStr("" & rsTmp!ID)): rptItem.Focusable = False
            Set rptItem = rptRcd.AddItem(CStr("" & rsTmp!����)): rptItem.Focusable = False
            Set rptItem = rptRcd.AddItem(CStr("" & rsTmp!����)): rptItem.Focusable = False
            Set rptItem = rptRcd.AddItem(CStr("" & rsTmp!Ӣ��)): rptItem.Focusable = False
            rsTmp.MoveNext
        Loop
        .Populate
        Call rptList_SelectionChanged '����ѡ���¼�
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub initVfg(ByVal lng����ID As Long)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lngCount As Long
    Dim BlnFind As Boolean
    On Error GoTo errHandle
    With vfgList
        .Rows = 2: .Cols = 14: .FixedRows = 1: .FixedCols = 0
        
        .TextMatrix(0, mCol.ID) = "ID": .TextMatrix(0, mCol.����) = "����"
        .TextMatrix(0, mCol.������) = "������": .TextMatrix(0, mCol.Ӣ����) = "Ӣ����"
        .TextMatrix(0, mCol.ҩ������) = "ҩ������": .TextMatrix(0, mCol.�ο���ֵ) = "�ο���ֵ"
        .TextMatrix(0, mCol.�ο���ֵ) = "�ο���ֵ": .TextMatrix(0, mCol.�ο�) = "�ο�"
        .TextMatrix(0, mCol.�жϷ�ʽ) = "�жϷ�ʽ": .TextMatrix(0, mCol.��ע) = "��ע"
        .TextMatrix(0, mCol.�ؼ���) = "�ؼ���": .TextMatrix(0, mCol.��ֵ���) = "��ֵ���"
        .TextMatrix(0, mCol.�м���) = "�м���": .TextMatrix(0, mCol.��ֵ���) = "��ֵ���"
        
        .ColWidth(mCol.ID) = 0: .ColWidth(mCol.����) = 1000
        .ColWidth(mCol.������) = 2600: .ColWidth(mCol.Ӣ����) = 1000: .ColWidth(mCol.ҩ������) = 800
        .ColWidth(mCol.�ο���ֵ) = 0: .ColWidth(mCol.�ο���ֵ) = 0: .ColWidth(mCol.�ο�) = 1000
        .ColWidth(mCol.�жϷ�ʽ) = 1000: .ColWidth(mCol.��ע) = 1000: .ColWidth(mCol.�ؼ���) = 0
        .ColWidth(mCol.��ֵ���) = 0: .ColWidth(mCol.�м���) = 0: .ColWidth(mCol.��ֵ���) = 0
        
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            If .ColWidth(lngCount) = 0 Then .ColHidden(lngCount) = True
        Next
        
        strSQL = "Select B.ID, B.����, B.������, B.Ӣ����, A.ҩ������, A.�ο���ֵ, A.�ο���ֵ, A.�жϷ�ʽ, A.��ע, A.��ֵ���, A.�м���, A.��ֵ���" & vbNewLine & _
                "From �����ÿ����� B, ����ϸ�������زο� A" & vbNewLine & _
                "Where A.������id = B.ID And A.ϸ��id = [1] And A.�����ط���id = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngϸ��id, lng����ID)
        
        Do Until rsTmp.EOF
            .TextMatrix(.Rows - 1, mCol.ID) = Val("" & rsTmp!ID)
            .TextMatrix(.Rows - 1, mCol.����) = "" & rsTmp!����
            .TextMatrix(.Rows - 1, mCol.������) = "" & rsTmp!������
            .TextMatrix(.Rows - 1, mCol.Ӣ����) = "" & rsTmp!Ӣ����
            .TextMatrix(.Rows - 1, mCol.ҩ������) = Getҩ������("" & rsTmp!ҩ������)             '1-MIC;2-DISK;3-K-B
            .TextMatrix(.Rows - 1, mCol.�ο���ֵ) = FormatDecimal("" & rsTmp!�ο���ֵ)
            .TextMatrix(.Rows - 1, mCol.�ο���ֵ) = FormatDecimal("" & rsTmp!�ο���ֵ)
            If "" & rsTmp!�ο���ֵ = "" Or "" & rsTmp!�ο���ֵ = "" Then
                .TextMatrix(.Rows - 1, mCol.�ο�) = FormatDecimal("" & rsTmp!�ο���ֵ) & FormatDecimal("" & rsTmp!�ο���ֵ)
            Else
                .TextMatrix(.Rows - 1, mCol.�ο�) = FormatDecimal("" & rsTmp!�ο���ֵ) & "��" & FormatDecimal("" & rsTmp!�ο���ֵ)
            End If
            
            .TextMatrix(.Rows - 1, mCol.�жϷ�ʽ) = Get�жϷ�ʽ("" & rsTmp!�жϷ�ʽ)
            .TextMatrix(.Rows - 1, mCol.��ע) = "" & rsTmp!��ע
            .TextMatrix(.Rows - 1, mCol.�ؼ���) = Val("" & rsTmp!ID) & "," & rsTmp!ҩ������
            .TextMatrix(.Rows - 1, mCol.��ֵ���) = IIf(Trim("" & rsTmp!��ֵ���) = "", "S-����", "" & rsTmp!��ֵ���)
            .TextMatrix(.Rows - 1, mCol.�м���) = IIf(Trim("" & rsTmp!�м���) = "", "I-�н�", "" & rsTmp!�м���)
            .TextMatrix(.Rows - 1, mCol.��ֵ���) = IIf(Trim("" & rsTmp!��ֵ���) = "", "R-��ҩ", "" & rsTmp!��ֵ���)
            .Rows = .Rows + 1
            rsTmp.MoveNext
        Loop
        If .Rows > 2 Then .Rows = .Rows - 1
        
        '---- ����δ����ļ�¼,�����û�����
        strSQL = "Select A.ID, A.����, A.������, A.Ӣ����, A.ҩ������" & vbNewLine & _
                "From ���鿹������ҩ B, �����ÿ����� A" & vbNewLine & _
                "Where A.ID = B.������id And B.�����ط���id = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
        Do Until rsTmp.EOF
            BlnFind = False
            For lngCount = .FixedRows To .Rows - 1
                If .TextMatrix(lngCount, mCol.�ؼ���) = Val("" & rsTmp!ID) & "," & rsTmp!ҩ������ Then
                    BlnFind = True
                End If
            Next
            If BlnFind = False Then
                If Val(.TextMatrix(.Rows - 1, mCol.ID)) <> 0 Then
                    .Rows = .Rows + 1
                End If
                .TextMatrix(.Rows - 1, mCol.ID) = Val("" & rsTmp!ID)
                .TextMatrix(.Rows - 1, mCol.����) = "" & rsTmp!����
                .TextMatrix(.Rows - 1, mCol.������) = "" & rsTmp!������
                .TextMatrix(.Rows - 1, mCol.Ӣ����) = "" & rsTmp!Ӣ����
                .TextMatrix(.Rows - 1, mCol.ҩ������) = Getҩ������("" & rsTmp!ҩ������)
                .TextMatrix(.Rows - 1, mCol.�жϷ�ʽ) = Get�жϷ�ʽ("1")
                .TextMatrix(.Rows - 1, mCol.�ؼ���) = Val("" & rsTmp!ID) & "," & rsTmp!ҩ������
                .TextMatrix(.Rows - 1, mCol.��ֵ���) = "S-����"
                .TextMatrix(.Rows - 1, mCol.�м���) = "I-�н�"
                .TextMatrix(.Rows - 1, mCol.��ֵ���) = "R-��ҩ"

            End If
            rsTmp.MoveNext
        Loop
        
       ' If .Rows > 2 Then .Rows = .Rows - 1
        
        .Select .FixedRows, mCol.���
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function Get�жϷ�ʽ(ByVal strIn) As String
    '0-�ο�ֵ���� 1-�����ο�ֵ
    If strIn = "0" Then
        Get�жϷ�ʽ = "�ο�ֵ����"
    ElseIf strIn = "1" Then
        Get�жϷ�ʽ = "�����ο�ֵ"
    ElseIf strIn = "�ο�ֵ����" Then
        Get�жϷ�ʽ = "0"
    ElseIf strIn = "�����ο�ֵ" Then
        Get�жϷ�ʽ = "1"
    End If
End Function

Private Function Getҩ������(ByVal strIn) As String
    '1-MIC;2-DISK;3-K-B
    If strIn = "1" Then
        Getҩ������ = "MIC"
    ElseIf strIn = "2" Then
        Getҩ������ = "DISK"
    ElseIf strIn = "3" Then
        Getҩ������ = "K-B"
    ElseIf strIn = "MIC" Then
        Getҩ������ = "1"
    ElseIf strIn = "DISK" Then
        Getҩ������ = "2"
    ElseIf strIn = "K-B" Then
        Getҩ������ = "3"
    End If
    
End Function

Public Sub ShowMe(ByVal lngϸ��ID As Long, ByVal frmMain As Form)

    If lngϸ��ID <= 0 Then Exit Sub
    mlngϸ��id = lngϸ��ID
    On Error Resume Next
    Me.Show vbModal, frmMain
    
End Sub

Private Sub txt�ο�ֵ_Validate(Index As Integer, Cancel As Boolean)
    txt�ο�ֵ(Index).Text = FormatDecimal(txt�ο�ֵ(Index).Text)
End Sub

Private Function FormatDecimal(ByVal strIn As String) As String
    '��.5�����ı���ʽΪ0.5
    Dim strTmp As String
    If InStr(strIn, ".") > 0 Then
        strTmp = Mid(strIn, InStr(strIn, ".") + 1)
        FormatDecimal = Format(strIn, "0." & String(Len(strTmp), "0"))
    Else
        FormatDecimal = strIn
    End If
End Function

Private Sub txt������_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt������.Text = "" And txt������.Tag <> "" Then Exit Sub
    
    On Error GoTo errHandle
    strText = txt������.Text
    
    If InStr(1, strText, "(") <> 0 Then
        If InStr(1, strText, ")") <> 0 Then
            strText = Mid(strText, 2, InStr(1, strText, ")") - 2)
        End If
    End If
        
    gstrSql = "Select B.ID, B.����, B.������, B.Ӣ����, ҩ������" & vbNewLine & _
            "From �����ÿ����� B, ���鿹������ҩ A" & vbNewLine & _
            "Where A.������id = B.Id  And A.�����ط���id = [1] And (" & _
            zlcommfun.GetLike("B", "����", strText) & " or " & zlcommfun.GetLike("B", "������", strText) & " or " & zlcommfun.GetLike("B", "����", strText) & ")"

    'Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngItemID)
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "������ѡ��", True, "", "��ѡ������", False, False, False, 0, 0, 0, blnReturn, True, False, mlngItemID)
    If blnReturn = False Then
        If rsTemp.EOF = True Then
            '��¼����û�п�ѡ�������
            txt������.Text = lbl������.Tag
            zlControl.TxtSelAll txt������
            Exit Sub
        Else
            '�϶����м�¼����
            txt������.Tag = rsTemp!ID
            txt������.Text = "(" & rsTemp!���� & ")" & rsTemp!������
            lbl������.Tag = txt������.Text '���ڻָ���ʾ
        End If
    End If
    cbo����.SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vfgList_RowColChange()
    
    Dim str����  As String
    Dim str�ж� As String
    Dim str��� As String
    With vfgList
        If .Row >= .FixedRows Then
            If Val(.TextMatrix(.Row, mCol.ID)) = 0 Then Exit Sub
            txt������.Tag = Val(.TextMatrix(.Row, mCol.ID))
            txt������.Text = "(" & .TextMatrix(.Row, mCol.����) & ")" & .TextMatrix(.Row, mCol.������)
            lbl������.Tag = txt������.Text
            str���� = .TextMatrix(.Row, mCol.ҩ������)
            str�ж� = .TextMatrix(.Row, mCol.�жϷ�ʽ)
            
            cbo����.ListIndex = Val(Getҩ������(str����)) - 1
            cbo�жϷ�ʽ.ListIndex = Val(Get�жϷ�ʽ(str�ж�))
            txt�ο�ֵ(0) = .TextMatrix(.Row, mCol.�ο���ֵ)
            txt�ο�ֵ(1) = .TextMatrix(.Row, mCol.�ο���ֵ)
            txt��ע = .TextMatrix(.Row, mCol.��ע)
            
            str��� = .TextMatrix(.Row, mCol.��ֵ���)
            cbo���(0).Text = str���
            str��� = .TextMatrix(.Row, mCol.�м���)
            cbo���(1).Text = str���
            str��� = .TextMatrix(.Row, mCol.��ֵ���)
            cbo���(2).Text = str���
            
        End If
    End With
End Sub
