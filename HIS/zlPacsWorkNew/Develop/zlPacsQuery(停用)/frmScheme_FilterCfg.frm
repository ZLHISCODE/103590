VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmScheme_FilterCfg 
   BorderStyle     =   0  'None
   Caption         =   "���ҹ�������"
   ClientHeight    =   6735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   15900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdFilterReset 
      Caption         =   "�� ��"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8400
      TabIndex        =   15
      Top             =   5760
      Width           =   1100
   End
   Begin VB.CommandButton cmdDeleteFilter 
      Caption         =   "ɾ ��"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8400
      TabIndex        =   6
      Top             =   4320
      Width           =   1100
   End
   Begin VB.CommandButton cmdNewFilter 
      Caption         =   "�������ٹ�����"
      Enabled         =   0   'False
      Height          =   465
      Left            =   8400
      TabIndex        =   5
      Top             =   3720
      Width           =   1100
   End
   Begin VB.CommandButton cmdDeleteCondition 
      Caption         =   "ɾ ��"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8400
      TabIndex        =   1
      Top             =   1320
      Width           =   1100
   End
   Begin VB.CommandButton cmdNewCondition 
      Caption         =   "�����Զ�������"
      Enabled         =   0   'False
      Height          =   465
      Left            =   8400
      TabIndex        =   0
      Top             =   600
      Width           =   1100
   End
   Begin VB.Frame fraFilterSet 
      Height          =   30
      Left            =   1920
      TabIndex        =   10
      Top             =   3480
      Width           =   7215
   End
   Begin VB.Frame fraInputSet 
      Height          =   30
      Left            =   1680
      TabIndex        =   9
      Top             =   240
      Width           =   6255
   End
   Begin VB.CommandButton cmdLastCondition 
      Caption         =   "�� ��"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8400
      TabIndex        =   2
      Top             =   1920
      Width           =   1100
   End
   Begin VB.CommandButton cmdNextCondition 
      Caption         =   "�� ��"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8400
      TabIndex        =   3
      Top             =   2520
      Width           =   1100
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "�� ��"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8400
      TabIndex        =   4
      Top             =   3120
      Width           =   1100
   End
   Begin VB.CommandButton cmdLastFilter 
      Caption         =   "�� ��"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8400
      TabIndex        =   7
      Top             =   4800
      Width           =   1100
   End
   Begin VB.CommandButton cmdNextFilter 
      Caption         =   "����"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8400
      TabIndex        =   8
      Top             =   5280
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfConditonCfg 
      Height          =   2655
      Left            =   720
      TabIndex        =   13
      Top             =   480
      Width           =   7335
      _cx             =   12938
      _cy             =   4683
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
      BackColorBkg    =   -2147483636
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   0
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   350
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
   Begin VSFlex8Ctl.VSFlexGrid vsfFilter 
      Height          =   2655
      Left            =   720
      TabIndex        =   14
      Top             =   3720
      Width           =   7335
      _cx             =   12938
      _cy             =   4683
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
      BackColorBkg    =   -2147483636
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   0
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   350
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
   Begin VB.Label lblInput 
      AutoSize        =   -1  'True
      Caption         =   "����¼������"
      Height          =   180
      Left            =   600
      TabIndex        =   12
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label lblFilter 
      AutoSize        =   -1  'True
      Caption         =   "���ٹ�������"
      Height          =   180
      Left            =   600
      TabIndex        =   11
      Top             =   3360
      Width           =   1080
   End
End
Attribute VB_Name = "frmScheme_FilterCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mobjFilterCfg As clsScSerachCfg
Public mblnIsEdit As Boolean    '�Ƿ��ѱ༭

Private mblnState As Boolean    '�Ƿ����ڱ༭״̬
Private mblnNewState As Boolean
Private mobjCustomQueryForm As New frmSetDataFrom
Private mstrFilterItem As String
Private mstrQuerySql As String
Private mobjSqlScheme As New clsSqlScheme
Private Const M_STR_INPUTCOL = "¼����Ŀ|¼�뷽ʽ|�ؼ�����|Ĭ��ģ��ƥ�䷽ʽ|Ĭ��ֵ|������Դ|"
Private Const M_STR_FILTERCOL = "������Ŀ|ѡ��ʽ|������Դ|�Զ�����˽ű�|"
Private Const M_STR_INSTYLE = "0-����¼��|1-���¼��|2-���+����"
Private Const M_STR_CONSTYLE = "0-�ı���|1-���ڿ�|2-ʱ���|3-�����ڿ�|4-������|5-��ѡ��|6-�����|7-�����"
Private Const M_STR_LIKESTYLE = "0-����|1-��ƥ��|2-��ƥ��|3-ȫƥ��"
Private Const M_STR_CHKSTYLE = "��ѡ|��ѡ"
Private Enum ConColTitlte
    it¼����Ŀ = 0
    it¼�뷽ʽ = 1
    it�ؼ����� = 2
    itĬ��ֵģ��ƥ�䷽ʽ = 3
    itĬ��ֵ = 4
    it������Դ = 5
    itIsNew = 6
End Enum

Private Enum FilColTitlte
    ft������Ŀ = 0
    ftѡ��ʽ = 1
    ft������Դ = 2
    ft�Զ�����˽ű� = 3
    ftIsNew = 4
End Enum

Private Sub cmdDeleteCondition_Click()
    On Error GoTo errHandle
    
    If vsfConditonCfg.Rows < 2 Or IsSelectionRow(vsfConditonCfg) = False Then Exit Sub
    
    If Val(vsfConditonCfg.TextMatrix(vsfConditonCfg.Row, ConColTitlte.itIsNew)) = 1 Then
        vsfConditonCfg.RemoveItem vsfConditonCfg.Row
        mblnIsEdit = True
        If vsfConditonCfg.Rows < 2 Then
            cmdDeleteCondition.Enabled = False
        End If
    Else
        MsgBox "��ѯ�������������������ɾ��", vbInformation, Me.Caption
    End If
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cmdDeleteFilter_Click()
    On Error GoTo errHandle
    
    If vsfFilter.Rows < 2 Or IsSelectionRow(vsfFilter) = False Then Exit Sub
     
    vsfFilter.RemoveItem vsfFilter.Row
    mblnIsEdit = True
    If vsfFilter.Rows < 2 Then
        cmdDeleteFilter.Enabled = False
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cmdFilterReset_Click()
    On Error GoTo errHandle
    
    Call ShowFilterSet(mobjSqlScheme, 2)
    If vsfFilter.Rows > 1 Then
        cmdDeleteFilter.Enabled = True
    Else
        cmdDeleteFilter.Enabled = False
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cmdLastCondition_Click()
    On Error GoTo errHandle
    
    Call MoveUp(vsfConditonCfg)
    mblnIsEdit = True
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cmdLastFilter_Click()
    On Error GoTo errHandle
    
    Call MoveUp(vsfFilter)
    mblnIsEdit = True
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cmdNewCondition_Click()
    On Error GoTo errHandle
    
    If vsfConditonCfg.Rows = 1 Then
        cmdDeleteCondition.Enabled = True
    End If
    
    mblnNewState = True
    Call NewRow(vsfConditonCfg)
    mblnIsEdit = True
    vsfConditonCfg.TextMatrix(vsfConditonCfg.Row, ConColTitlte.itIsNew) = 1
    Call ConCfgDataDefalut(vsfConditonCfg.Row)
    vsfConditonCfg.EditCell
    mblnNewState = False
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cmdNewFilter_Click()
    On Error GoTo errHandle
    
    If vsfFilter.Rows = 1 Then
        cmdDeleteFilter.Enabled = True
    End If
    Call NewRow(vsfFilter)
    mblnIsEdit = True
    vsfFilter.TextMatrix(vsfFilter.Row, FilColTitlte.ftIsNew) = 1
    Call FiltetDataDefalut(vsfFilter.Row)
    vsfFilter.EditCell
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cmdNextCondition_Click()
    On Error GoTo errHandle
    
    Call MoveDown(vsfConditonCfg)
    mblnIsEdit = True
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cmdNextFilter_Click()
    On Error GoTo errHandle
    
    Call MoveDown(vsfFilter)
    mblnIsEdit = True
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cmdReset_Click()
'����
    On Error GoTo errHandle
    
    Call ShowFilterSet(mobjSqlScheme, 1)
    Call RefreshFilterSet(mstrQuerySql, mobjSqlScheme, True)
    If vsfConditonCfg.Rows > 1 Then
        cmdDeleteCondition.Enabled = True
    Else
        cmdDeleteCondition.Enabled = False
    End If
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub Form_Load()
    On Error GoTo errHandle
    
    mblnNewState = False
    Call GridInit(M_STR_INPUTCOL, vsfConditonCfg)
    Call GridInit(M_STR_FILTERCOL, vsfFilter)
    Call GridShow
    Call RefreshWindowState(False)
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    '¼�����ò���
    lblInput.Move Me.ScaleLeft + 100, Me.ScaleTop + 100
    fraInputSet.Move lblInput.Left + lblInput.Width, lblInput.Top + lblInput.Height / 2, Me.ScaleWidth - lblInput.Left
    vsfConditonCfg.Move Me.ScaleLeft + 100, fraInputSet.Top + 200, Me.ScaleWidth - 300 - cmdNewCondition.Width, (Me.ScaleHeight - vsfConditonCfg.Top * 2 - 300) / 2
    cmdNewCondition.Move vsfConditonCfg.Left + vsfConditonCfg.Width + 100, vsfConditonCfg.Top
    cmdDeleteCondition.Move cmdNewCondition.Left, cmdNewCondition.Top + cmdNewCondition.Height + 100
    cmdLastCondition.Move cmdNewCondition.Left, cmdDeleteCondition.Top + cmdDeleteCondition.Height + 100
    cmdNextCondition.Move cmdNewCondition.Left, cmdLastCondition.Top + cmdLastCondition.Height + 100
    cmdReset.Move cmdNewCondition.Left, cmdNextCondition.Top + cmdNextCondition.Height + 100
    
    '�������ò���
    lblFilter.Move lblInput.Left, vsfConditonCfg.Top + vsfConditonCfg.Height + 100
    fraFilterSet.Move lblFilter.Left + lblFilter.Width, lblFilter.Top + lblFilter.Height / 2, Me.ScaleWidth - fraFilterSet.Left
    vsfFilter.Move vsfConditonCfg.Left, fraFilterSet.Top + 200, vsfConditonCfg.Width, vsfConditonCfg.Height
    cmdNewFilter.Move cmdNewCondition.Left, vsfFilter.Top
    cmdDeleteFilter.Move cmdNewCondition.Left, cmdNewFilter.Top + cmdNewFilter.Height + 100
    cmdLastFilter.Move cmdNewCondition.Left, cmdDeleteFilter.Top + cmdDeleteFilter.Height + 100
    cmdNextFilter.Move cmdNewCondition.Left, cmdLastFilter.Top + cmdLastFilter.Height + 100
    cmdFilterReset.Move cmdNewCondition.Left, cmdNextFilter.Top + cmdNextFilter.Height + 100
End Sub


Private Sub GridShow()
    With vsfConditonCfg
        .ColHidden(ConColTitlte.itIsNew) = True
        .ColComboList(ConColTitlte.it¼�뷽ʽ) = M_STR_INSTYLE
        .ColComboList(ConColTitlte.it�ؼ�����) = M_STR_CONSTYLE
        .ColComboList(ConColTitlte.itĬ��ֵģ��ƥ�䷽ʽ) = M_STR_LIKESTYLE
        .ColComboList(ConColTitlte.itĬ��ֵ) = "..."
        .ColComboList(ConColTitlte.it������Դ) = "..."
        .ColWidth(ConColTitlte.itĬ��ֵģ��ƥ�䷽ʽ) = 1700
        .ColWidth(ConColTitlte.it�ؼ�����) = 1200
        .ColWidth(ConColTitlte.it¼����Ŀ) = 1200
        .ColWidth(ConColTitlte.itĬ��ֵ) = 2000
        .ColWidth(ConColTitlte.it¼�뷽ʽ) = 1200
    End With
    With vsfFilter
        .ColHidden(FilColTitlte.ftIsNew) = True
        .ColComboList(FilColTitlte.ftѡ��ʽ) = M_STR_CHKSTYLE
        .ColComboList(FilColTitlte.ft������Դ) = "..."
        .ColComboList(FilColTitlte.ft�Զ�����˽ű�) = "..."
        .ColWidth(FilColTitlte.ft������Ŀ) = 1200
        .ColWidth(FilColTitlte.ft������Դ) = 4000
    End With
End Sub

Private Sub vsfConditonCfg_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strValue As String
    Dim strPara As String
    
     If Col = ConColTitlte.itĬ��ֵ Or Col = ConColTitlte.it������Դ Then
        For i = 1 To Row - 1
            If vsfConditonCfg.RowHidden(i) = False And vsfConditonCfg.TextMatrix(i, ConColTitlte.it�ؼ�����) <> "8-��ѡ��" And Len(Trim(vsfConditonCfg.TextMatrix(i, ConColTitlte.it¼����Ŀ))) > 0 Then
                strPara = strPara & "|" & vsfConditonCfg.TextMatrix(i, ConColTitlte.it¼����Ŀ)
            End If
        Next
        strPara = Mid(strPara, 2)
        strValue = vsfConditonCfg.TextMatrix(Row, Col)
        strValue = mobjCustomQueryForm.ShowSqlFromWindow(strValue, strPara, IIf(Col = ConColTitlte.itĬ��ֵ, 1, 2), mblnState, Me)
        vsfConditonCfg.TextMatrix(Row, Col) = strValue
    End If
End Sub

Private Sub vsfConditonCfg_CellChanged(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo errHandle
    
    If mblnState Then
        mblnIsEdit = True
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub vsfFilter_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strValue As String
    Dim strPara As String
    
    If Col = FilColTitlte.ft������Դ Then
        For i = 1 To Row - 1
            If vsfFilter.RowHidden(i) = False And Len(Trim(vsfFilter.TextMatrix(i, FilColTitlte.ft������Ŀ))) > 0 Then
                strPara = strPara & "|" & vsfFilter.TextMatrix(i, FilColTitlte.ft������Ŀ)
            End If
        Next
        strPara = Mid(strPara, 2)
        strValue = vsfFilter.TextMatrix(Row, Col)
        strValue = mobjCustomQueryForm.ShowSqlFromWindow(strValue, strPara, 2, mblnState, Me)
        vsfFilter.TextMatrix(Row, Col) = strValue
    End If
    
    If Col = FilColTitlte.ft�Զ�����˽ű� Then
        strValue = vsfFilter.TextMatrix(Row, Col)
        strValue = mobjCustomQueryForm.ShowSqlFromWindow(strValue, "", 4, mblnState, Me)
        vsfFilter.TextMatrix(Row, Col) = strValue
    End If
End Sub

Private Sub vsfFilter_CellChanged(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo errHandle
    
    If mblnState Then
        mblnIsEdit = True
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub vsfFilter_RowColChange()
    Dim strFilterItem As String
    
    On Error GoTo errHandle
    
    If vsfFilter.Row < 1 Then Exit Sub
    vsfFilter.Editable = flexEDKbdMouse
    If mblnState Then
        strFilterItem = mstrFilterItem
        If vsfFilter.Col = 0 And vsfFilter.Row > 0 Then
            strFilterItem = InitFilterItem(mstrQuerySql)
            For i = 1 To vsfFilter.Row - 1
                If Val(vsfFilter.TextMatrix(i, FilColTitlte.ftIsNew)) = 1 And Len(Trim(vsfFilter.TextMatrix(i, FilColTitlte.ft������Ŀ))) > 0 Then
                    strFilterItem = strFilterItem & "|" & vsfFilter.TextMatrix(i, FilColTitlte.ft������Ŀ)
                End If
            Next
            vsfFilter.ColComboList(FilColTitlte.ft������Ŀ) = strFilterItem
        End If
    Else
        If Not (vsfFilter.Col = FilColTitlte.ft������Դ Or vsfFilter.Col = FilColTitlte.ft�Զ�����˽ű�) Then
            vsfFilter.Editable = flexEDNone
        End If
    End If

    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub


Private Sub vsfConditonCfg_RowColChange()
    On Error GoTo errHandle
    
    If vsfConditonCfg.Row < 1 Then Exit Sub
    vsfConditonCfg.Editable = flexEDKbdMouse
    If mblnState Then
        If vsfConditonCfg.Col = ConColTitlte.it¼����Ŀ And Not (Val(vsfConditonCfg.TextMatrix(vsfConditonCfg.Row, ConColTitlte.itIsNew)) = 1 Or mblnNewState) Or vsfConditonCfg.TextMatrix(vsfConditonCfg.Row, ConColTitlte.it�ؼ�����) = "8-��ѡ��" Then
            vsfConditonCfg.Editable = flexEDNone
        End If
    Else
        If Not (vsfConditonCfg.Col = ConColTitlte.itĬ��ֵ Or vsfConditonCfg.Col = ConColTitlte.it������Դ) Then
            vsfConditonCfg.Editable = flexEDNone
        End If
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub



Private Sub ConCfgDataDefalut(lngRow As Long)
'vsfConditonCfg����Ĭ��ֵ
    With vsfConditonCfg
        .TextMatrix(lngRow, ConColTitlte.it¼�뷽ʽ) = "0-����¼��"
        .TextMatrix(lngRow, ConColTitlte.it�ؼ�����) = "0-�ı���"
        .TextMatrix(lngRow, ConColTitlte.itĬ��ֵģ��ƥ�䷽ʽ) = "0-����"
        .TextMatrix(lngRow, ConColTitlte.itĬ��ֵ) = ""
        .TextMatrix(lngRow, ConColTitlte.it������Դ) = ""
    End With
End Sub

Private Sub FiltetDataDefalut(lngRow As Long)
'vsfFilter����Ĭ��ֵ
    With vsfFilter
        .TextMatrix(lngRow, FilColTitlte.ftѡ��ʽ) = "��ѡ"
        .TextMatrix(lngRow, FilColTitlte.ft������Դ) = ""
        .TextMatrix(lngRow, FilColTitlte.ft�Զ�����˽ű�) = ""
    End With
End Sub

Public Sub SetConditionCfg(objSqlScheme As clsSqlScheme)
    'д��¼������
    Dim objScSearchCfg As clsScSerachCfg
    Dim objScFilterCfg As clsScFilterCfg
    Dim i As Long
    
    If vsfConditonCfg.Rows < 2 Then Exit Sub
    For i = 1 To vsfConditonCfg.Rows - 1
        Set objScSearchCfg = New clsScSerachCfg
        With vsfConditonCfg
            If Len(.TextMatrix(i, ConColTitlte.it¼����Ŀ)) > 0 And .RowHidden(i) = False Then
                objScSearchCfg.Name = .TextMatrix(i, ConColTitlte.it¼����Ŀ)
                objScSearchCfg.InputType = SetConDataChange(i, ConColTitlte.it¼�뷽ʽ)
                objScSearchCfg.ControlType = SetConDataChange(i, ConColTitlte.it�ؼ�����)
                objScSearchCfg.Default = .TextMatrix(i, ConColTitlte.itĬ��ֵ)
                objScSearchCfg.LikeWay = SetConDataChange(i, ConColTitlte.itĬ��ֵģ��ƥ�䷽ʽ)
                objScSearchCfg.DataFrom = .TextMatrix(i, ConColTitlte.it������Դ)
                objSqlScheme.AddSerachCfg objScSearchCfg
            End If
        End With
    Next
    
    '���ٹ�������
    For i = 1 To vsfFilter.Rows - 1
        Set objScFilterCfg = New clsScFilterCfg
        With vsfFilter
            If Len(.TextMatrix(i, FilColTitlte.ft������Ŀ)) > 0 And .RowHidden(i) = False Then
                objScFilterCfg.Name = .TextMatrix(i, FilColTitlte.ft������Ŀ)
                objScFilterCfg.SelectWay = IIf(.TextMatrix(i, FilColTitlte.ftѡ��ʽ) = "��ѡ", 1, 0)
                objScFilterCfg.DataFrom = .TextMatrix(i, FilColTitlte.ft������Դ)
                objScFilterCfg.CustomScript = .TextMatrix(i, FilColTitlte.ft�Զ�����˽ű�)
                objSqlScheme.AddFilterCfg objScFilterCfg
            End If
        End With
    Next
End Sub


Private Function SetConDataChange(lngRow As Long, lngCol As Long) As Long
'vsfConditonCfgд������ת��
    Dim strValue As String
    Dim arrData() As String
    strValue = vsfConditonCfg.TextMatrix(lngRow, lngCol)
    
    If Len(strValue) = 0 Then
        SetConDataChange = 0
        Exit Function
    End If
    
    arrData = Split(strValue, "-")
    SetConDataChange = Val(arrData(0))
End Function

Private Function GetConDataChange(strItem As String, lngNo As Long) As String
'vsfConditonCfg��ȡ����ת��
    Dim arrContent() As String
    Dim arrText() As String
    Dim i As Long
    
    Select Case strItem
        Case "ConColTitlte"
            arrContent = Split(M_STR_INSTYLE, "|")
        Case "ControlType"
            arrContent = Split(M_STR_CONSTYLE, "|")
        Case "LikeWay"
            arrContent = Split(M_STR_LIKESTYLE, "|")
    End Select
    
    For i = 0 To UBound(arrContent)
        arrText = Split(arrContent(i), "-")
        If lngNo = arrText(0) Then
            GetConDataChange = arrText(0) & "-" & arrText(1)
            Exit Function
        ElseIf lngNo = 8 And strItem = "ControlType" Then
            GetConDataChange = "8-��ѡ��"
        End If
    Next
End Function

Public Sub RefreshWindowState(blnState As Boolean)
    mblnState = blnState
    cmdDeleteCondition.Enabled = False
    cmdDeleteFilter.Enabled = False
    cmdLastCondition.Enabled = blnState
    cmdLastFilter.Enabled = blnState
    cmdNewCondition.Enabled = blnState
    cmdNewFilter.Enabled = blnState
    cmdNextCondition.Enabled = blnState
    cmdNextFilter.Enabled = blnState
    cmdReset.Enabled = blnState
    cmdFilterReset.Enabled = blnState
    
    If blnState Then
        vsfConditonCfg.BackColor = &H80000005
        vsfFilter.BackColor = &H80000005
        If vsfConditonCfg.Rows > 1 Then
            cmdDeleteCondition.Enabled = blnState
        End If
        
        If vsfFilter.Rows > 1 Then
            cmdDeleteFilter.Enabled = blnState
        End If
    Else
        vsfConditonCfg.BackColor = &H8000000F
        vsfFilter.BackColor = &H8000000F
    End If
   
End Sub

Private Function InitFilterItem(strSchemeSql As String) As String
'���ÿ�ѡ������Ŀ
    Dim objSqlParse As New clsSqlParse
    Dim objQuery As New clsPacsQuery
    Dim rsRecord As ADODB.Recordset
    Dim strItem As String
    Dim i As Long

    objSqlParse.init strSchemeSql

    Set rsRecord = objQuery.GetQueryField(objSqlParse)
    mstrFilterItem = ""
    If rsRecord Is Nothing Then Exit Function
    For i = 0 To rsRecord.Fields.Count - 1
        InitFilterItem = InitFilterItem & "|" & rsRecord.Fields(i).Name
    Next
End Function

'Public Sub ClearScheme()
'    vsfConditonCfg.Rows = 1
'    vsfFilter.Rows = 1
'End Sub


Private Function GetQueryItem(strSchemeSql As String) As String
    Dim objSqlParse As New clsSqlParse
    Dim objQuery As New clsPacsQuery
    Dim rsRecord As ADODB.Recordset
    
    mstrDefinedItem = ""
    objSqlParse.init strSchemeSql
    Set rsRecord = objQuery.GetQueryField(objSqlParse)
    If rsRecord Is Nothing Then Exit Function
    For i = 0 To rsRecord.Fields.Count - 1
        GetQueryItem = GetQueryItem & "|" & rsRecord.Fields(i).Name
    Next
    
    GetQueryItem = GetQueryItem & "|"
End Function

Public Sub ShowFilterSet(objSqlScheme As clsSqlScheme, Optional lngReset As Long)
'����������ʾ
    Dim objScSearchCfg As New clsScSerachCfg
    Dim objScFilterCfg As New clsScFilterCfg
    Dim objSqlParse As New clsSqlParse
    Dim objQuery As New clsPacsQuery
    Dim rsRecord As ADODB.Recordset
    Dim strSelecItem As String
    Dim lngRow As Long
    Dim arrQueryPara() As String
    Dim strQueryItem As String
    Dim i As Long
    Dim j As Long
     
    Set mobjSqlScheme = objSqlScheme
    mstrQuerySql = objSqlScheme.Query
    
    objSqlParse.init objSqlScheme.Query
    Set rsRecord = objQuery.GetQueryField(objSqlParse)
    
    '��ʾ¼������
    If lngReset <> 2 Then
        vsfConditonCfg.Rows = 1
        For i = 1 To objSqlScheme.SerachCfgCount
            Set objScSearchCfg = objSqlScheme.SerachCfg(i)
            With vsfConditonCfg
                If InStr(1, UCase(gstrPara), "[" & UCase(objScSearchCfg.Name) & "]") = 0 Then
                    .Rows = .Rows + 1
                    lngRow = .Rows - 1
                    .TextMatrix(lngRow, ConColTitlte.it¼����Ŀ) = objScSearchCfg.Name
                    .TextMatrix(lngRow, ConColTitlte.it¼�뷽ʽ) = GetConDataChange("ConColTitlte", objScSearchCfg.InputType)
                    .TextMatrix(lngRow, ConColTitlte.it�ؼ�����) = GetConDataChange("ControlType", objScSearchCfg.ControlType)
                    If .TextMatrix(lngRow, ConColTitlte.it�ؼ�����) = "8-��ѡ��" Then
                        vsfConditonCfg.Cell(flexcpBackColor, lngRow, 0, lngRow, vsfConditonCfg.Cols - 1) = &HC0FFFF
                    End If
                    .TextMatrix(lngRow, ConColTitlte.itĬ��ֵģ��ƥ�䷽ʽ) = GetConDataChange("LikeWay", objScSearchCfg.LikeWay)
                    .TextMatrix(lngRow, ConColTitlte.itĬ��ֵ) = objScSearchCfg.Default
                    .TextMatrix(lngRow, ConColTitlte.it������Դ) = objScSearchCfg.DataFrom
                    .TextMatrix(lngRow, ConColTitlte.itIsNew) = IIf(objSqlParse.SqlStruct.HasParName(objScSearchCfg.Name), 0, 1)
                End If
            End With
        Next
    End If
    '���ٹ�������
    If lngReset <> 1 Then
        vsfFilter.Rows = 1
        For i = 1 To objSqlScheme.FilterCfgCount
            Set objScFilterCfg = objSqlScheme.FilterCfg(i)
            With vsfFilter
                .Rows = .Rows + 1
                .TextMatrix(i, FilColTitlte.ft������Ŀ) = objScFilterCfg.Name
                .TextMatrix(i, FilColTitlte.ftѡ��ʽ) = IIf(objScFilterCfg.SelectWay = swMulti, "��ѡ", "��ѡ")
                .TextMatrix(i, FilColTitlte.ft������Դ) = objScFilterCfg.DataFrom
                .TextMatrix(i, FilColTitlte.ft�Զ�����˽ű�) = objScFilterCfg.CustomScript
                .TextMatrix(i, FilColTitlte.ftIsNew) = IIf(HasSelectItem(objScFilterCfg.Name, rsRecord), 0, 1)
    
            End With
        Next
    End If
End Sub

Public Sub RefreshFilterSet(strQuerySql As String, objSqlScheme As clsSqlScheme, Optional lngReset As Long)
    Dim objSqlParse As New clsSqlParse
    Dim objQuery As New clsPacsQuery
    Dim rsRecord As ADODB.Recordset
    Dim strQueryPara As String
    Dim strCurPara As String
    Dim lngRow As Long
    Dim i As Long
    Dim j As Long
    Dim blnIsCusPara As Boolean
    Dim blnIsHave As Boolean
    
    mstrQuerySql = strQuerySql
    Set mobjSqlScheme = objSqlScheme
    objSqlParse.init strQuerySql
    Set rsRecord = objQuery.GetQueryField(objSqlParse)
    
    'ˢ��¼����Ŀ����
    If lngReset <> 2 Then
        For i = 1 To vsfConditonCfg.Rows - 1
            If Val(vsfConditonCfg.TextMatrix(i, ConColTitlte.itIsNew)) <> 1 And (Not vsfConditonCfg.RowHidden(i)) Then
                strQueryPara = strQueryPara & "," & "[" & vsfConditonCfg.TextMatrix(i, ConColTitlte.it¼����Ŀ) & "]"
                If Not objSqlParse.SqlStruct.HasParName(vsfConditonCfg.TextMatrix(i, ConColTitlte.it¼����Ŀ)) And InStr(1, UCase(gstrPara), "[" & UCase(vsfConditonCfg.TextMatrix(i, ConColTitlte.it¼����Ŀ)) & "]") = 0 Then
                    vsfConditonCfg.RowHidden(i) = True
                End If
            End If
        Next
        
        strQueryPara = Mid(strQueryPara, 2)
        For i = 1 To objSqlParse.SqlStruct.ParCount
            blnIsCusPara = False
            strCurPara = objSqlParse.SqlStruct.AllParameter(i)
            If InStr(strCurPara, "[@") > 0 Then
                strCurPara = Mid$(strCurPara, 3, InStr(strCurPara, ",") - 3)
                blnIsCusPara = True
            Else
                strCurPara = Mid(strCurPara, 2, Len(strCurPara) - 2)
            End If
            
            If InStr(1, UCase(strQueryPara), "[" & UCase(strCurPara) & "]") = 0 And InStr(1, UCase(gstrPara), "[" & UCase(strCurPara) & "]") = 0 Then
                '�Ƿ����Զ����ظ�
                blnIsHave = False
                For j = 1 To vsfConditonCfg.Rows - 1
                    If UCase(Trim(strCurPara)) = UCase(Trim(vsfConditonCfg.TextMatrix(j, ConColTitlte.it¼����Ŀ))) And (Not vsfConditonCfg.RowHidden(j)) Then
                        blnIsHave = True
                    End If
                Next
                If Not blnIsHave Then
                    vsfConditonCfg.AddItem strCurPara, vsfConditonCfg.Rows
                    Call ConCfgDataDefalut(vsfConditonCfg.Rows - 1)
                    If blnIsCusPara Then
                        vsfConditonCfg.TextMatrix(vsfConditonCfg.Rows - 1, ConColTitlte.it�ؼ�����) = "8-��ѡ��"
                        vsfConditonCfg.Cell(flexcpBackColor, vsfConditonCfg.Rows - 1, 0, vsfConditonCfg.Rows - 1, vsfConditonCfg.Cols - 1) = &HC0FFFF
                    End If
                End If
            End If
        Next
    End If
    If lngReset <> 1 Then
        'ˢ�¿��ٹ�������
        For i = 1 To vsfFilter.Rows - 1
            If Val(vsfFilter.TextMatrix(i, FilColTitlte.ftIsNew)) <> 1 And (Not vsfFilter.RowHidden(i)) Then
                If Not HasSelectItem(vsfFilter.TextMatrix(i, FilColTitlte.ft������Ŀ), rsRecord) Then
                    vsfFilter.RowHidden(i) = True
                End If
            End If
        Next
    End If
    
    Set objSqlParse = Nothing
    Set objQuery = Nothing
End Sub

Private Function HasSelectItem(strItem As String, rsRecord As Recordset) As Boolean
    Dim i As Long
    
    HasSelectItem = False
    For i = 0 To rsRecord.Fields.Count - 1
        If UCase(strItem) = UCase(rsRecord.Fields(i).Name) Then
            HasSelectItem = True
            Exit Function
        End If
    Next
End Function

Public Function IsEnabledSave() As Boolean
    Dim blnResult As Boolean
    
    blnResult = CheckRepet(vsfConditonCfg, ConColTitlte.it¼����Ŀ)
    If blnResult Then
        MsgBox "����¼��������¼����Ŀ���ظ�������", vbInformation, Me.Caption
        IsEnabledSave = False
        Exit Function
    End If
    
    blnResult = CheckRepet(vsfFilter, FilColTitlte.ft������Ŀ)
    If blnResult Then
        MsgBox "���ٹ��������й�����Ŀ���ظ�������", vbInformation, Me.Caption
        IsEnabledSave = False
        Exit Function
    End If
    
    IsEnabledSave = True
End Function

Public Sub UnloadMe()
    Set mobjFilterCfg = Nothing
    Set mobjCustomQueryForm = Nothing
    Set mobjSqlScheme = Nothing
End Sub

