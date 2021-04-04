VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmItemWaveMan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������Ŀ����"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5895
   Icon            =   "frmItemWaveMan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VSFlex8Ctl.VSFlexGrid VsfData 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   480
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
      FormatString    =   $"frmItemWaveMan.frx":020A
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgPublic 
      Bindings        =   "frmItemWaveMan.frx":026C
      Left            =   480
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmItemWaveMan.frx":0280
   End
End
Attribute VB_Name = "frmItemWaveMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnEdit As Boolean

Private Const conMenu_���� = 2
Private Const conMenu_�ָ� = 3
Private Const conMenu_���� = 4
Private Const conMenu_�˳� = 5

Private Enum EnumCOl
    ��Ŀ���� = 0
    �Ƿ񲨶� = 1
End Enum


Private Function SaveData() As Boolean
'����������Ϣ
    Dim rs As New ADODB.Recordset
    Dim strData As String
    Dim intRow As Integer
    Dim lngOrder As Long
    Dim lngID As String
    Dim arrSQL() As String
    Dim i As Integer
    Dim blnTrans As Boolean
    
    On Error GoTo errHand
    
    For intRow = 1 To VsfData.Rows - 1
        lngOrder = VsfData.RowData(intRow)
        If lngOrder <> 0 And VsfData.TextMatrix(intRow, �Ƿ񲨶�) = "��" Then
            lngID = lngID & "," & lngOrder
            strData = strData & "|" & lngOrder & ";" & VsfData.TextMatrix(intRow, ��Ŀ����)
        End If
    Next intRow
    
    If Left(strData, 1) = "|" Then strData = Mid(strData, 2)
    If Left(lngID, 1) = "," Then lngID = Mid(lngID, 2)
    
    ReDim Preserve arrSQL(1 To 1)
    '�޸Ĳ�����Ŀ��¼Ƶ��(>2�����)
    If Val(lngID) <> 0 Then
        gstrSQL = "Select  /*+ RULE*/ ��Ŀ���,�������,��¼��,��¼��,��¼��,��¼ɫ,���ֵ,��Сֵ,��λֵ,��λ,�����,��¼Ƶ��,�̶ȼ��,��ʾ�� " & _
            "   From ���¼�¼��Ŀ A,Table(Cast(f_num2list([1]) As zlTools.t_Numlist)) B" & _
            "  Where A.��Ŀ���=B.Column_Value And nvl(A.��¼Ƶ��,0)>2"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "�����¼��Ŀ", CStr(lngID))

        With rs
            Do While Not .EOF
                gstrSQL = "ZL_���¼�¼��Ŀ_INSERT(" & NVL(!��Ŀ���, 0) & "," & _
                                                                Val(NVL(!�������)) & ",'" & _
                                                                Trim(NVL(!��¼��)) & "'," & _
                                                                Val(NVL(!��¼��)) & ",'" & _
                                                                NVL(!��¼��) & "'," & _
                                                                Val(NVL(!��¼ɫ)) & "," & _
                                                                IIf(Trim(NVL(!��Сֵ)) <> "", Val(NVL(!��Сֵ)), "NULL") & "," & _
                                                                IIf(Trim(NVL(!���ֵ)) <> "", Val(NVL(!���ֵ)), "NULL") & "," & _
                                                                IIf(Trim(NVL(!��λֵ)) <> "", Val(NVL(!��λֵ)), "NULL") & ",'" & _
                                                                Trim(NVL(!��λ)) & "'," & _
                                                                "NULL" & "," & _
                                                                2 & "," & IIf(Trim(NVL(!�̶ȼ��)) = "", "NULL", Val(NVL(!�̶ȼ��))) & "," & _
                                                                IIf(Trim(NVL(!��ʾ��)) = "", "NULL", Val(NVL(!��ʾ��))) & ")"
                arrSQL(ReDimArray(arrSQL)) = gstrSQL
            .MoveNext
            Loop
        End With
        
        '��������
        blnTrans = (rs.RecordCount > 0)
    End If
    gstrSQL = "zl_��������Ŀ_Upate('" & strData & "')"
    arrSQL(ReDimArray(arrSQL)) = gstrSQL
    
    If blnTrans Then gcnOracle.BeginTrans
    For i = 1 To UBound(arrSQL)
        If CStr(arrSQL(i)) <> "" Then Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "��������Ŀ")
    Next i
    
    If blnTrans Then gcnOracle.CommitTrans
    
    SaveData = True
    Exit Function
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function

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
    With VsfData
        .Left = lngLeft
        .Top = lngTop
        .Height = lngHeight - lngTop
        .Width = lngWidth
    End With
End Sub

Private Sub VsfData_DblClick()
    Call VsfData_KeyDown(vbKeySpace, 0)
End Sub

Private Sub VsfData_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngOrder As Long
    Dim intRow As Integer, intCount As Integer
    If VsfData.Col = �Ƿ񲨶� And KeyCode = vbKeySpace Then
        
        '���ò�����Ŀ
        lngOrder = VsfData.RowData(VsfData.Row)
        If lngOrder = 0 Then Exit Sub
        VsfData.TextMatrix(VsfData.Row, �Ƿ񲨶�) = IIf(VsfData.TextMatrix(VsfData.Row, �Ƿ񲨶�) = "��", "", "��")
        intCount = VsfData.Rows - 1
        If lngOrder = 4 Or lngOrder = 5 Then
            For intRow = 1 To intCount
                If VsfData.RowData(intRow) = IIf(lngOrder = 4, 5, 4) Then
                    VsfData.TextMatrix(intRow, �Ƿ񲨶�) = VsfData.TextMatrix(VsfData.Row, �Ƿ񲨶�)
                    Exit For
                End If
            Next intRow
        End If
        
        mblnEdit = True
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_����
        If Not SaveData Then Exit Sub
        mblnEdit = False
    Case conMenu_�ָ�
        Call LoadData
    Case conMenu_����
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_�˳�
        Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    With VsfData
        .Left = lngLeft
        .Top = lngTop
        .Width = lngRight
        .Height = lngBottom
    End With
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
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call MainDefCommandBar
    Call LoadData
End Sub

Private Sub LoadData()
'--��ʼ���������
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHand
    
    mblnEdit = False
    With VsfData
        .Clear
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 0
        .Cols = 2
        .TextMatrix(0, ��Ŀ����) = "��Ŀ����"
        .TextMatrix(0, �Ƿ񲨶�) = "�Ƿ񲨶�"
        .ColWidth(��Ŀ����) = 1400
        .ColWidth(�Ƿ񲨶�) = 900
        
        .ColAlignment(��Ŀ����) = flexAlignLeftCenter
        .ColAlignment(�Ƿ񲨶�) = flexAlignCenterCenter
    End With
    
    '�������
    gstrSQL = " SELECT A.��Ŀ���,A.��Ŀ����,DECODE(NVL(C.��Ŀ���,0),0,0,1) ������Ŀ" & vbNewLine & _
            "   FROM �����¼��Ŀ A,���¼�¼��Ŀ B,��������Ŀ C" & vbNewLine & _
            "   WHERE A.��Ŀ���=B.��Ŀ��� AND A.��Ŀ���=C.��Ŀ���(+) AND A.��Ŀ����=0" & vbNewLine & _
            "   AND A.��Ŀ��ʾ=0 AND B.��¼��=2 AND B.��Ŀ���<>3" & vbNewLine & _
            "   ORDER BY A.��Ŀ���"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���±����Ŀ")
    With rsTemp
        Do While Not .EOF
            If .AbsolutePosition >= VsfData.Rows Then VsfData.Rows = VsfData.Rows + 1
            VsfData.TextMatrix(.AbsolutePosition, ��Ŀ����) = CStr(!��Ŀ����)
            VsfData.TextMatrix(.AbsolutePosition, �Ƿ񲨶�) = IIf(NVL(!������Ŀ, 0) = 1, "��", "")
            VsfData.RowData(.AbsolutePosition) = CLng(!��Ŀ���)
            .MoveNext
        Loop
    End With
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

