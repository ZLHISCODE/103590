VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmTendCollect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����������"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5805
   Icon            =   "frmTendCollect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboSelect 
      BackColor       =   &H80000018&
      Height          =   300
      ItemData        =   "frmTendCollect.frx":000C
      Left            =   690
      List            =   "frmTendCollect.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1305
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1305
      TabIndex        =   1
      Top             =   2100
      Visible         =   0   'False
      Width           =   825
   End
   Begin VSFlex8Ctl.VSFlexGrid VsfData 
      Height          =   3495
      Left            =   60
      TabIndex        =   0
      Top             =   1140
      Width           =   5115
      _cx             =   9022
      _cy             =   6165
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
      BackColorSel    =   16764057
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   5000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmTendCollect.frx":0022
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   30
      Picture         =   "frmTendCollect.frx":0084
      Top             =   390
      Width           =   480
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   $"frmTendCollect.frx":094E
      Height          =   540
      Left            =   570
      TabIndex        =   3
      Top             =   435
      Width           =   5175
      WordWrap        =   -1  'True
   End
   Begin XtremeCommandBars.ImageManager imgPublic 
      Left            =   480
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmTendCollect.frx":09F8
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
Attribute VB_Name = "frmTendCollect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnInit As Boolean
Private mblnEdit As Boolean

Private Const conMenu_ɾ�� = 1
Private Const conMenu_���� = 2
Private Const conMenu_�ָ� = 3
Private Const conMenu_���� = 4
Private Const conMenu_�˳� = 5
Private mrsTemp As New ADODB.Recordset
Private mrsCollect As New ADODB.Recordset

Private Enum colMenu
    ��Ŀ���
    ��Ŀ����
    �󶨷���
    ѡ��
End Enum

Private Sub cboSelect_Click()
    Dim arrResult() As String
    Dim strValue As String
    Dim lngLoop As Long
    If Not mblnEdit Then mblnEdit = (VsfData.TextMatrix(VsfData.Row, 0) <> cboSelect.Text)
    With mrsCollect
    mrsCollect.Filter = "��Ŀ���� = '" & cboSelect.Text & "'"
    Do While Not .EOF
        If NVL(!��Ŀֵ��) <> "" Then
            arrResult = Split(NVL(!��Ŀֵ��), ";")
            strValue = Trim(VsfData.TextMatrix(VsfData.Row, �󶨷���))
            
            For lngLoop = 0 To UBound(arrResult)
                If strValue = "" Then
                    strValue = arrResult(lngLoop)
                Else
                    If Not InStr(1, "|" & strValue & "|", "|" & arrResult(lngLoop) & "|") > 0 Then
                        strValue = strValue & "|" & arrResult(lngLoop)
                    End If
                End If
            Next
            
            VsfData.TextMatrix(VsfData.Row, �󶨷���) = strValue
            
        Else
            strValue = Trim(VsfData.TextMatrix(VsfData.Row, �󶨷���))
            If strValue = "" Then
                VsfData.TextMatrix(VsfData.Row, �󶨷���) = NVL(!��Ŀ����)
            Else
                If Not InStr(1, "|" & strValue & "|", "|" & NVL(!��Ŀ����) & "|") > 0 Then
                    VsfData.TextMatrix(VsfData.Row, �󶨷���) = strValue & "|" & NVL(!��Ŀ����)
                End If
            End If
        End If
        .MoveNext
    Loop
    End With
    VsfData.Col = �󶨷���
    Call InitCons
End Sub

Private Sub cboSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cboSelect.ListIndex < 0 Then Exit Sub
        Call cboSelect_Click
        VsfData.Col = 1
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_����
        If Not CheckData Then Exit Sub
        If Not SaveData Then Exit Sub
    Case conMenu_�ָ�
        Call LoadData
    Case conMenu_����
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_�˳�
        Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Call Form_Resize
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_����
        Control.Enabled = mblnEdit
    Case conMenu_�ָ�
        Control.Enabled = mblnEdit
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr(" &[]{}+'""|", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call MainDefCommandBar
    Call LoadData
End Sub

Private Function CheckData() As Boolean
    
    CheckData = True
End Function

Private Function SaveData() As Boolean
    Dim lngOrder As Long
    Dim lngLoop As Long
    Dim strGCollect As String
    Dim strSQL() As String
    Dim blnTran As Boolean
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
         '����
    '    ��Ŀ���_IN IN  �����¼��Ŀ.��Ŀ���%TYPE,
    '    ��Ŀ����_IN IN  �����¼��Ŀ.��Ŀ����%TYPE,
    '    ��Ŀ����_IN IN  �����¼��Ŀ.��Ŀ����%TYPE,
    '    ��Ŀ����_IN IN  �����¼��Ŀ.��Ŀ����%TYPE,
    '    ��ĿС��_IN IN  �����¼��Ŀ.��ĿС��%TYPE,
    '    ��Ŀ��λ_IN IN  �����¼��Ŀ.��Ŀ��λ%TYPE,
    '    ��Ŀ��ʾ_IN IN  �����¼��Ŀ.��Ŀ��ʾ%TYPE,
    '    ��Ŀֵ��_IN IN  �����¼��Ŀ.��Ŀֵ��%TYPE,
    '    ����ȼ�_IN   IN  �����¼��Ŀ.����ȼ�%TYPE,
    '    ������_IN   IN  �����¼��Ŀ.������%TYPE,
    '    ��ĿID_IN   IN  �����¼��Ŀ.��ĿID%TYPE
    
    If VsfData.TextMatrix(lngLoop, �󶨷���) <> "" Then
        For lngLoop = VsfData.FixedRows To VsfData.Rows - 1
            lngOrder = Val(VsfData.TextMatrix(lngLoop, ��Ŀ���))
            With mrsTemp
                mrsTemp.Filter = "��Ŀ���=" & lngOrder
                If mrsTemp.RecordCount > 0 Then
                    strSQL(ReDimArray(strSQL)) = "ZL_�����¼��Ŀ_UPDATE(" & lngOrder & ",'" & _
                    NVL(!��Ŀ����) & "'," & NVL(!��Ŀ����) & "," & NVL(!��Ŀ����) & "," & NVL(!��ĿС��) & ",'" & _
                    NVL(!��Ŀ��λ) & "'," & NVL(!��Ŀ��ʾ) & ",'" & NVL(!��Ŀֵ��) & "'," & _
                    NVL(!����ȼ�) & ",'" & NVL(!������) & "','" & NVL(!��ĿID) & "'," & NVL(!Ӧ�÷�ʽ) & "," & _
                    NVL(!���ò���) & "," & NVL(!��Ŀ����) & "," & NVL(!Ӧ�ó���) & ",'" & NVL(!˵��) & "','" & _
                    NVL(!ȱʡֵ) & "','" & VsfData.TextMatrix(lngLoop, �󶨷���) & "')"
                
                End If
            End With
        Next
    End If
    
    blnTran = True
    gcnOracle.BeginTrans
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    gcnOracle.CommitTrans
    blnTran = False
    
    SaveData = True
    mblnEdit = False
    
    Exit Function
    
errHand:
    '������
    
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub LoadData()
    
    On Error GoTo errHand
    
    mblnEdit = False
    mblnInit = True
    Call InitCons
    With VsfData
        .Clear
        .Rows = 2
        .Cols = 4
        .TextMatrix(0, ��Ŀ���) = "��Ŀ���"
        .TextMatrix(0, ��Ŀ����) = "��Ŀ����"
        .TextMatrix(0, �󶨷���) = "��������"
        .TextMatrix(0, ѡ��) = "ѡ��"
        .ColWidth(��Ŀ����) = 1000
        .ColWidth(�󶨷���) = 2500
        .ColWidth(ѡ��) = 1500
        .ColHidden(��Ŀ���) = True
        .ColAlignment(��Ŀ����) = flexAlignLeftCenter
        .ColAlignment(�󶨷���) = flexAlignLeftCenter
        .ColAlignment(ѡ��) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, ��Ŀ���, 0, ѡ��) = flexAlignCenterCenter
    End With
    
    '�������
    gstrSQL = "" & _
            " select A.��Ŀ���,A.��Ŀ����,A.��Ŀ����,A.��Ŀ����,A.��ĿС��,A.��Ŀ��λ,A.��Ŀ��ʾ,A.��Ŀֵ��,A.����ȼ�," & _
            " A.������,A.��Ŀid,A.Ӧ�÷�ʽ,A.���ò���,A.��Ŀ����,A.Ӧ�ó���,A.˵��,A.ȱʡֵ,A.�������" & _
            " from �����¼��Ŀ A" & _
            " Where A.��Ŀ��ʾ=4 " & _
            " Order By A.��Ŀ���"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���л�����Ŀ")
    With mrsTemp
        Do While Not .EOF
            If VsfData.TextMatrix(.AbsolutePosition, ��Ŀ����) = "" Then VsfData.Rows = VsfData.Rows + 1
            VsfData.TextMatrix(.AbsolutePosition, ��Ŀ���) = CStr(!��Ŀ���)
            VsfData.TextMatrix(.AbsolutePosition, ��Ŀ����) = CStr(!��Ŀ����)
            VsfData.TextMatrix(.AbsolutePosition, �󶨷���) = NVL(!�������)
            .MoveNext
        Loop
    End With
    VsfData.Rows = VsfData.Rows - 1
    
    gstrSQL = "" & _
        " select A.��Ŀ���,A.��Ŀ����,A.��Ŀֵ�� " & _
        " from �����¼��Ŀ A" & _
        " Where A.��Ŀ����= 1 and A.��Ŀ��ʾ in (0,2) " & _
        " Order By A.��Ŀ���"
    'Ϊ��������ӻ��Ŀ
    Set mrsCollect = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���л�����Ŀ")
    With mrsCollect
        Me.cboSelect.Clear
        Do While Not .EOF
            cboSelect.AddItem !��Ŀ����
            cboSelect.ItemData(cboSelect.NewIndex) = !��Ŀ���
            .MoveNext
        Loop
    End With
    mblnInit = False
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub MainDefCommandBar()
'���ܣ������ڲ˵����岿��
'˵����
'1.���й��еĲ˵��Ͱ�ť�����У���Ϊ�Ӵ��崦��˵��Ļ�׼
'2.�����������������ҵ��Ĳ�ͬ�����ܲ�ͬ
    Dim objMenu As CommandBarPopup, objFile As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom
    Dim objControl As CommandBarControl
    Dim lngHandel As Long

    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    
    'cbsMain
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
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
    cbsMain.EnableCustomization False
    cbsMain.Icons = imgPublic.Icons
    
    '����������
    '-----------------------------------------------------
    cbsMain.DeleteAll
    Set objBar = cbsMain.Add("������", xtpBarTop)      '����
    objBar.EnableDocking xtpFlagStretched
    objBar.Closeable = False
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_����, "����"): objControl.STYLE = xtpButtonIconAndCaption: objControl.ToolTipText = "��������": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_�ָ�, "�ָ�"): objControl.STYLE = xtpButtonIconAndCaption: objControl.ToolTipText = "ȡ������"
        Set objControl = .Add(xtpControlButton, conMenu_����, "����"): objControl.STYLE = xtpButtonIconAndCaption: objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_�˳�, "�˳�"): objControl.STYLE = xtpButtonIconAndCaption
    End With
    
    '����һЩ�������ȼ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyS, conMenu_����             '����
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnEdit Then
        If MsgBox("�޸ĵ����ݻ�δ���棬��ȷ��Ҫ�˳���" & vbCrLf & "�㡰�ǡ�������޸Ĳ��˳����㡰�񡱼����޸ģ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Resize()
    Dim lngLeft As Long, lngTop As Long, lngHeight As Long, lngWidth As Long
    
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngWidth, lngHeight)
    imgNote.Move lngLeft + 20, lngTop - 20
    lblNote.Move lngLeft + imgNote.Width, lngTop
    With VsfData
        .Left = lngLeft
        .Top = lngTop + lblNote.Height + 30
        .Height = lngHeight - lngTop
        .Width = lngWidth
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mrsTemp Is Nothing Then Set mrsTemp = Nothing
    If Not mrsCollect Is Nothing Then Set mrsCollect = Nothing
End Sub

Private Sub txtInput_GotFocus()
    txtInput.SelStart = 0
    txtInput.SelLength = 100
    txtInput.Text = VsfData.TextMatrix(VsfData.Row, �󶨷���)
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtInput.Text) = "" And Trim(VsfData.TextMatrix(VsfData.Row, �󶨷���)) = "" Then Exit Sub
    If Not mblnEdit Then mblnEdit = (VsfData.TextMatrix(VsfData.Row, �󶨷���) <> Trim(txtInput.Text))
    VsfData.TextMatrix(VsfData.Row, �󶨷���) = Trim(txtInput.Text)
    If Trim(txtInput.Text) = "" Then VsfData.TextMatrix(VsfData.Row, �󶨷���) = " "
    VsfData.Col = �󶨷���
    If VsfData.Row + 1 <= VsfData.Rows - 1 Then VsfData.Row = VsfData.Row + 1
End Sub

Private Sub VsfData_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call InitCons
End Sub

Private Sub VsfData_DblClick()
    Call VsfData_EnterCell
    
'    Call VsfData_KeyDown(vbKeySpace, 0)
End Sub

Private Sub VsfData_EnterCell()
    Dim objCon As Object
    Dim lngLeft As Long, lngTop As Long, lngHeight As Long, lngWidth As Long
    
    Call InitCons
    If mblnInit Then Exit Sub
    If VsfData.Col < �󶨷��� Then Exit Sub
    
    If Not VsfData.RowIsVisible(VsfData.Row) Then VsfData.TopRow = VsfData.Row
    lngLeft = VsfData.Left + VsfData.CellLeft + 10
    lngTop = VsfData.Top + VsfData.CellTop + 10
    lngHeight = VsfData.CellHeight - 10
    lngWidth = VsfData.CellWidth - 10
    
    Select Case VsfData.Col
    Case ѡ��
        Set objCon = Me.cboSelect
    Case �󶨷���
        Set objCon = Me.txtInput
    End Select
    
    With objCon
        .Left = lngLeft
        .Top = lngTop
        If VsfData.Col <> ѡ�� Then .Height = lngHeight
        .Width = lngWidth
        
        On Error Resume Next
        Err = 0
        .Text = VsfData.TextMatrix(VsfData.Row, VsfData.Col)
        
        .Visible = True
        .SetFocus
    End With
End Sub

Private Sub InitCons()
    cboSelect.Visible = False
    txtInput.Visible = False
End Sub


