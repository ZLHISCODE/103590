VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmStTableContent 
   Caption         =   "��׼·�����༭"
   ClientHeight    =   7425
   ClientLeft      =   3720
   ClientTop       =   2490
   ClientWidth     =   11760
   Icon            =   "frmStTableContent.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   11760
   WindowState     =   2  'Maximized
   Begin VSFlex8Ctl.VSFlexGrid vsPathTable 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   11775
      _cx             =   20770
      _cy             =   7011
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
      BackColorFixed  =   12632256
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   16444122
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   100
      RowHeightMax    =   3000
      ColWidthMin     =   100
      ColWidthMax     =   12000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmStTableContent.frx":076A
      ScrollTrack     =   0   'False
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
      WordWrap        =   -1  'True
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
   Begin XtremeCommandBars.ImageManager imgMain 
      Left            =   2520
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmStTableContent.frx":082A
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   240
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmStTableContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngStPathID As Long
Private mlng��� As Long
Private mblnOK As Boolean
Private mblnChange As Boolean

Public Function ShowMe(ByRef frmParent As Object, ByVal lngStPathID As Long, Optional ByVal lng��� As Long = 1) As Boolean
'���ܣ����ر�׼·����

    mlngStPathID = lngStPathID
    mlng��� = lng���
    mblnChange = False
    mblnOK = False
    Me.Show 1, frmParent
    ShowMe = mblnOK
    
End Function

Private Sub ExcExit()
'���ܣ��˳�����
    
    mblnOK = False
    Unload Me
    
End Sub

Private Sub ExcSaveTable()
    Dim strSql As String
    Dim arrSql As Variant
    Dim i As Long, j As Long
    Dim blnTrans As Boolean
    
    Call AdjustTable
    On Error GoTo errH
    If mblnChange Then
        arrSql = Array()
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        strSql = "Zl_��׼·����_ContentClear(" & mlngStPathID & "," & mlng��� & ")"
        arrSql(UBound(arrSql)) = strSql
        With vsPathTable
            For i = .FixedRows + 1 To .Rows - 1
                For j = .FixedCols + 1 To .Cols - 1
                    If Not (Trim(.TextMatrix(i, .FixedCols)) = "" And i = .Rows - 1 Or Trim(.TextMatrix(.FixedRows, j)) = "" And j = .Cols - 1) Then
                        strSql = "Zl_��׼·����_ContentInsert(" & mlngStPathID & "," & mlng��� & "," & i + 1 - .FixedRows & ",'" & _
                               Trim(.TextMatrix(i, .FixedCols)) & "'," & j + 1 - .FixedCols & ",'" & Trim(.TextMatrix(.FixedRows, j)) & "','" & Trim(.TextMatrix(i, j)) & "')"
                        ReDim Preserve arrSql(UBound(arrSql) + 1)
                        arrSql(UBound(arrSql)) = strSql
                    End If
                Next
            Next
        End With
        
        'ִ��SQL
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSql)
            zlDatabase.ExecuteProcedure CStr(arrSql(i)), Me.Caption
        Next
        gcnOracle.CommitTrans: blnTrans = False
        If blnTrans = False Then
            mblnOK = True
        End If
    Else
        mblnOK = False
    End If
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    With vsPathTable
        Select Case Control.ID
            Case conMenu_NewRow
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .Col = .FixedCols
                Call .ShowCell(.Rows - 1, .Col)
                .Cell(flexcpBackColor, .Rows - 1, 0) = &HE1FFE1
            Case conMenu_NewCol
                .Cols = .Cols + 1
                .Col = .Cols - 1
                .Row = .FixedRows
                Call .ShowCell(.Row, .Cols - 1)
                .Cell(flexcpBackColor, .FixedRows, .Cols - 1) = &HE1FFE1
            Case conMenu_DelCol
                Call ExecFunc(1)
                .Col = .Cols - 1
                .Row = .FixedRows
            Case conMenu_DelRow
                Call ExecFunc(0)
                .Col = .FixedCols
                .Row = .Rows - 1
            Case conMenu_ClearItem
                Call ExecFunc(2)
            Case conMenu_Edit
                Call vsPathTable_DblClick
            Case conMenu_Save
                Call ExcSaveTable
            Case conMenu_Exit
                Call ExcExit
        End Select
    End With
End Sub

Private Sub cbsMain_Resize()
    Dim lngTop As Long, lngLeft As Long, lngRight As Long, lngBottom As Long
    
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    vsPathTable.Left = lngLeft
    vsPathTable.Top = lngTop
    vsPathTable.Width = lngRight - lngLeft
    vsPathTable.Height = lngBottom - lngTop
    
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    With vsPathTable
        Select Case Control.ID
            Case conMenu_NewRow '����������8�У����һ�еķ�������Ϊ�գ��������������У���������д���һ�еķ������ƣ����������������
                '��ֹ��ͣ��������У�������д������ݣ�ʵ�ʷ�������ж���8�еģ�
                Control.Enabled = Not ((Trim(.TextMatrix(.Rows - 1, .FixedCols)) = "") And (.Rows > 8))
            Case conMenu_NewCol '����������12�У����һ�еĽ׶�����Ϊ�գ��������������У���������д���һ�еĽ׶����ƣ����������������
                '��ֹ��ͣ��������У�������д������ݣ�ʵ�ʽ׶κ����ж���12�еģ�
                Control.Enabled = Not ((Trim(.TextMatrix(.FixedRows, .Cols - 1)) = "") And (.Cols > 12))
            Case conMenu_DelCol
                Control.Enabled = (.Col = .Cols - 1 And .Col >= .FixedCols)
                If Control.Enabled Then
                    Control.Enabled = RowOrColCanDel(True)
                End If
            Case conMenu_DelRow
                Control.Enabled = (.Row = .Rows - 1 And .Row >= .FixedRows)
                If Control.Enabled Then
                    Control.Enabled = RowOrColCanDel(False)
                End If
            Case conMenu_ClearItem, conMenu_Edit
                Control.Enabled = Not (.Row <= .FixedRows And .Col <= .FixedCols)
        End Select
    End With
End Sub

Private Sub Form_Load()
'���ܣ����ر�׼·����
    Call InitCommandBar
    Call InitPathTable
    Call SetVsStyle
    
End Sub

Private Sub InitPathTable()
'���ܣ���ʼ����׼·����
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, j As Long
    Dim lngCols As Long, lngRows As Long
    
    On Error GoTo errH
        '��ȡ������
        strSql = "Select Max(�׶����) ��, Max(�������) �� From ��׼·���� A Where ��׼·��id = [1] And ����� = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngStPathID, mlng���)
        If rsTmp.RecordCount > 0 Then
            lngCols = rsTmp!��
            lngRows = rsTmp!��
        End If
        '��ȡ������
        strSql = "Select a.�������, a.��������, a.�׶����, a.�׶�����, a.·������ From ��׼·���� A Where ��׼·��id = [1] And ����� = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngStPathID, mlng���)
    
        '�����
        With vsPathTable
            .Redraw = False
            .Cols = 0
            .Rows = 0
            .Cols = IIf(lngCols = 1, lngCols + 1, lngCols)
            .Rows = IIf(lngRows = 1, lngRows + 2, lngRows + 1)
            .FixedCols = 0
            .FixedRows = 1
            
            If lngCols = 1 And lngRows = 1 Then  'û������
                .TextMatrix(.FixedRows, .FixedCols) = "ʱ��"
            Else
                For i = .FixedRows To .Rows - 1
                    For j = .FixedCols To .Cols - 1
                        If j > .FixedCols And i > .FixedCols Then
                            rsTmp.Filter = "�׶����=" & (j + 1 - .FixedCols) & " And �������=" & (i + 1 - .FixedRows)
                            If rsTmp.RecordCount > 0 Then
                                .TextMatrix(i, j) = rsTmp!·������ & "" '�����ǰ�棬��Ϊ��һ����Ԫ��ʹ��ͷ�У�Ҳ�����ݣ���������������
                                .Cell(flexcpData, i, j) = rsTmp!·������ & ""
                                .TextMatrix(i, .FixedCols) = nvl(rsTmp!��������, "")
                                .Cell(flexcpData, i, .FixedCols) = nvl(rsTmp!��������, "")
                                .TextMatrix(.FixedRows, j) = nvl(rsTmp!�׶�����, "")
                                .Cell(flexcpData, .FixedRows, j) = nvl(rsTmp!�׶�����, "")
                            End If
                        ElseIf j = .FixedCols And i = .FixedRows Then
                            .TextMatrix(i, j) = "ʱ��"
                        End If
                    Next
                Next
            End If
            .Redraw = True
            For i = 1 To .Cols - 1
                .ColWidth(i) = 4000
            Next
            .ColWidth(0) = 1500
            Call SetVsStyle
            
        End With
    Exit Sub
    
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    '��֤��С�ߴ���ڱ༭��
    If Me.Height < 4500 Then Me.Height = 4500
    If Me.Width < 6500 Then Me.Width = 6500
    
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
'���ܣ��Ƿ�δ������˳�����ʾ
    If Not mblnOK And mblnChange Then
        If MsgBox("�������Ѿ��ı䣬��δ���棬�˳����ᶪʧ��ĸı䣬�Ƿ��˳�", vbInformation + vbYesNo, gstrSysName) = vbNo Then
            Cancel = True
        End If
    End If
End Sub

Private Sub vsPathTable_DblClick()
'���ܣ��༭��Ŀ��Ԫ��
    Dim strReturn As String
    Dim vPoint As POINTAPI
    
    With vsPathTable
        '��һ����Ԫ������༭
        If .Row <= .FixedRows And .Col <= .FixedCols Then Exit Sub
        .TopRow = .Row
        strReturn = Trim(.TextMatrix(.Row, .Col))
        '��ȡ��ǰλ��
        vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
        If frmStTableItemEdit.ShowMe(Me, strReturn, vPoint.x, vPoint.y) = True Then
            .TextMatrix(.Row, .Col) = strReturn
            mblnChange = True
            Call SetVsStyle
        End If
    End With
End Sub

Private Sub vsPathTable_KeyDown(KeyCode As Integer, Shift As Integer)
'���ܣ�delete������ʵ��
    With vsPathTable
        If KeyCode = vbKeyDelete Then
            If Not (.Row <= .FixedRows And .Col <= .FixedCols) Then
                Call ExecFunc(2)
            End If
        End If
    End With
End Sub

Private Sub vsPathTable_KeyPress(KeyAscii As Integer)
'���ܣ�����س���λ����
    With vsPathTable
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub

Private Sub vsPathTable_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'���ܣ����Ƶ�Ԫ���Ƿ���Ա��༭
    '��һ����Ԫ������༭
    If Row <= vsPathTable.FixedRows And Col <= vsPathTable.FixedCols Then Cancel = True
End Sub

Private Sub DelCol(ByVal lngCol As Long)
'���ܣ�ɾ���ض���
'����   lngCol ��ǰҪɾ������
    Dim i As Long, j As Long

    With vsPathTable
        If lngCol = .Cols - 1 Then
            .Cols = .Cols - 1 '���һ��ֱ��ɾ��
        Else
            '����һ��ǰ�ƣ���ɾ�����һ��
            .Redraw = False
            For i = lngCol + 1 To .Cols - 1
                For j = .FixedRows To .Rows - 1
                    .TextMatrix(j, i - 1) = .TextMatrix(j, i)
                Next
            Next
            .Cols = .Cols - 1
            .Redraw = True
        End If
    End With
End Sub

Private Sub SetVsStyle()
'���ܣ������������ñ����ĵ�Ԫ��ĸ߶�����,�Լ�������ɫ�ȣ��Լ���Ԫ��ĺϲ���

    Dim i As Long, j As Long
    Dim lngmaxHeight As Long
    
    With vsPathTable
        
        '�޸ķ������ƣ��׶Σ�����Ӵ־���
        .Cell(flexcpFontBold, .FixedRows, .FixedCols, .Rows - 1, .FixedCols) = True
        .Cell(flexcpFontBold, .FixedRows, .FixedCols, .FixedRows, .Cols - 1) = True
        .Cell(flexcpAlignment, .FixedRows, .FixedCols, .Rows - 1, .FixedCols) = 4   '����
        .Cell(flexcpAlignment, .FixedRows, .FixedCols, .FixedRows, .Cols - 1) = 4 '����
        .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .FixedCols) = &HE1FFE1
        .Cell(flexcpBackColor, .FixedRows, .FixedCols, .FixedRows, .Cols - 1) = &HE1FFE1

        '��ȡͬһ����ߵĵ�Ԫ��߶ȸ�ֵ���и�
        For i = .FixedRows To .Rows - 1
            For j = .FixedCols To .Cols - 1
                If j = .FixedCols Then
                    lngmaxHeight = ComputerLines(.TextMatrix(i, j))
                Else
                    lngmaxHeight = IIf(lngmaxHeight > ComputerLines(.TextMatrix(i, j)), lngmaxHeight, ComputerLines(.TextMatrix(i, j)))
                End If
            Next
            .RowHeight(i) = IIf(lngmaxHeight < 5 And i <> .FixedRows, 5, lngmaxHeight) * Me.TextHeight("��") * 1.5
        Next
        .RowHeight(0) = Me.TextHeight("��") * 0.5
        
    End With
    
End Sub

Private Function ComputerLines(ByVal strInput As String) As Long
'���ܣ����������ı��лس����ĸ���
'������  strInput   Ҫ����س������ַ���
'���أ�   �س����ĸ���

    Dim strTmp As String
    Dim Count  As Long, lngPos As Long, lngLen As Long
    
    lngPos = InStr(strInput, Chr(13))
    lngLen = Len(strInput)
    strTmp = strInput
    
    Do While lngPos <> 0
        If Trim(strTmp) = "" Then Exit Do
        If lngPos + 1 <= lngLen Then
            strTmp = Mid(strTmp, lngPos + 1)
            Count = Count + 1
            lngPos = InStr(strTmp, Chr(13))
            lngLen = Len(strTmp)
        End If
    Loop
    
    ComputerLines = Count + 2
    
End Function

Private Sub AdjustTable()
    Dim i As Long, j As Long
    With vsPathTable
        .Redraw = False
        
        'ɾ����������Ϊ�յ���
        For i = .FixedRows + 1 To .Rows - 1 Step 1
            If Trim(.TextMatrix(i, .FixedCols)) = "" Then
                .RemoveItem i
                i = i - 1
                Exit For
            End If
        Next
        
        'ɾ���׶�����Ϊ�յ���
        For j = .FixedCols + 1 To .Cols - 1 Step 1
            If Trim(.TextMatrix(.FixedRows, j)) = "" Then
                Call DelCol(j)
                j = j - 1
                Exit For
            End If
        Next
        .Redraw = True
    End With

End Sub

Private Sub InitCommandBar()
'���ܣ���ʼ��������
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
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
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = imgMain.Icons
    
    '���ɹ�����
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        
        Set objControl = .Add(xtpControlButton, conMenu_NewRow, "������")
            objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, conMenu_DelRow, "ɾ����")
            objControl.Style = xtpButtonIconAndCaption
        
        Set objControl = .Add(xtpControlButton, conMenu_NewCol, "������")
            objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, conMenu_DelCol, "ɾ����")
            objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, conMenu_ClearItem, "ɾ������"): objControl.BeginGroup = True
            objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, conMenu_Edit, "�༭����")
            objControl.Style = xtpButtonIconAndCaption
            
        Set objControl = .Add(xtpControlButton, conMenu_Save, "����")
            objControl.Style = xtpButtonIconAndCaption
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Exit, "�˳�")
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    '�ȼ���:ע�ⲻ�ܺ�ϵͳ���ı��༭�ȼ���ͻ���Լ�Form_keydown�еĳ�ͻ
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyS, conMenu_Save
        .Add FCONTROL, vbKeyR, conMenu_NewRow
        .Add FCONTROL, vbKeyC, conMenu_NewCol
        .Add FCONTROL, vbKeyI, conMenu_ClearItem
        .Add FCONTROL, vbKeyQ, conMenu_DelRow
        .Add FCONTROL, vbKeyW, conMenu_DelCol
        .Add FCONTROL, vbKeyE, conMenu_Edit
        .Add FALT, vbKeyX, conMenu_Exit
    End With
End Sub

Private Sub ExecFunc(ByVal intMode As Integer)
    With vsPathTable
        .Redraw = False
        Select Case intMode
            Case 0
                'ɾ������
                If .Row = .Rows - 1 And .Row > .FixedRows Then
                    If MsgBox("��ȷ��Ҫɾ����ǰ����", vbInformation + vbYesNo, gstrSysName) = vbYes Then
                        .RemoveItem .Row
                        mblnChange = True
                    End If
                End If
            Case 1
                'ɾ���׶�
                If .Col = .Cols - 1 And .Col > .FixedCols Then
                    If MsgBox("��ȷ��Ҫɾ��������", vbInformation + vbYesNo, gstrSysName) = vbYes Then
                        Call DelCol(.Col)
                        mblnChange = True
                    End If
                End If
            Case 2
                'ɾ����ǰһ�׶ε�ǰ��������ݣ���������ƣ��׶�����
                If Not (.Row <= .FixedRows And .Col <= .FixedCols) Then
                    If MsgBox("��ȷ��Ҫ��յ�ǰ��Ԫ��������", vbInformation + vbYesNo, gstrSysName) = vbYes Then
                        .TextMatrix(.Row, .Col) = ""
                        mblnChange = True
                    End If
                End If
        End Select
        .Redraw = True
    End With
End Sub

Private Function RowOrColCanDel(ByVal blnIsCol As Boolean) As Boolean
'���ܣ�ɾ���л�ɾ�����Ƿ����
'������
'     blnIsCol �Ƿ�����е���Ч�Լ�飬true-���н��м��,false -���н��м��
    Dim i As Long, lngCount As Long
    
    With vsPathTable
        If blnIsCol Then
            For i = .FixedCols + 1 To .Cols - 2
                If Trim(.TextMatrix(.FixedRows, i)) <> "" Then
                    RowOrColCanDel = True
                    Exit Function
                End If
            Next
        Else
            For i = .FixedRows + 1 To .Rows - 2
                If Trim(.TextMatrix(i, .FixedCols)) <> "" Then
                    RowOrColCanDel = True
                    Exit Function
                End If
            Next
        End If
    End With
End Function

