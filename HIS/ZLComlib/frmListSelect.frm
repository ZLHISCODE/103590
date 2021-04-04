VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmListSel 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6735
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8865
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   8865
   Begin VSFlex8Ctl.VSFlexGrid vsSelect 
      Height          =   6000
      Left            =   225
      TabIndex        =   0
      Top             =   480
      Width           =   8205
      _cx             =   14473
      _cy             =   10583
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
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   8
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
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
      ExplorerBar     =   7
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
Attribute VB_Name = "frmListSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngSys As Long, mlngModule As Long
Private mfrmMain As Form, mobjControl As Object
Private mrsBindings As ADODB.Recordset, mrsOutSel As ADODB.Recordset
Private mblnShowHead As Boolean, mstr������ As String, mstrHideCols As String '��1,��2,...
Private mblnOK As Boolean
'-------------------------------------------------------------------------------------------------------------------
'�ؼ���λ
Private Type ty_ctlObject_Locale
    '�ؼ���λ��
    Left As Single
    Top As Single
    Width As Single
    Height As Single
    '�����б����С�߶ȺͿ��
    minWidth As Single
    minHeight As Single
    
    '�½��б��ʵ��λ��
    DownLeft As Single
    DownTop As Single
    DownWidth As Single
    DownHeight As Single
 
    
    '��ģ���
    ScreenWidth As Single
    ScreenHeight As Single
    
End Type
Private mTyCtl_Locale As ty_ctlObject_Locale
'-------------------------------------------------------------------------------------------------------------------
'--API����
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Function ShowSelect(ByVal frmMain As Form, ByVal lngSys As Long, ByVal lngModule As Long, ByVal objControl As Object, ByVal rsBindings As ADODB.Recordset, _
     Optional ByVal blnShowHead As Boolean = False, _
     Optional ByVal str������ As String = "", _
     Optional ByVal strHideCols As String = "", _
     Optional ByRef rsOutSel As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ�������
    '���:frmMain-���õ�������
    '     lngSys-ϵͳ��
    '     lngModule-ģ���
    '     objControl-�ؼ�����(Ŀǰֻ֧:textBox,Combox)
    '     rsBindings-�󶨵ļ�¼��(����Ϊ��,��Ҫ�ֶ�,ID,......)
    '     str����-���Ի�����Ĳ�����.
    '     blnShowHead-�Ƿ���ʾ����ͷ
    '
    '����:rsOutSel-ѡ���ļ�¼��
    '����:ѡ�з���True, ���򷵻�False(���԰�Esc���з���)
    '����:���˺�
    '����:2009-01-01 15:35:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set rsOutSel = Nothing
    
    'û������,ֱ�ӷ���
    If rsBindings.RecordCount = 0 Then Exit Function
    'ֻ��һ������,��ֱ�ӷ���
    If rsBindings.RecordCount = 1 Then Set rsOutSel = rsBindings: ShowSelect = True: Exit Function
    Set mfrmMain = frmMain: mlngSys = lngSys: mlngModule = lngModule
    Set mrsBindings = rsBindings: mblnOK = False: Set mobjControl = objControl
    mblnShowHead = blnShowHead: mstr������ = str������: mstrHideCols = strHideCols
    '������
    Call zlBindingData
    
    '��ʼ���ؼ�λ��
    Call InitCtrlLocal
    '��������λ��
    Call ReSetWindowsFormLocal
    Me.Show 1, frmMain
    Set rsOutSel = mrsOutSel
    ShowSelect = mblnOK
End Function
Private Function SelectedItem() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ��ָ��������
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-07-20 12:21:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngId As Long
    With vsSelect
        lngId = Val(.TextMatrix(.Row, .ColIndex("ID")))
    End With
    If lngId = 0 Then Exit Function
    
    Set mrsOutSel = mrsBindings
    mrsOutSel.Filter = "ID=" & lngId
    If mrsOutSel.RecordCount = 0 Then Exit Function
    mblnOK = True
    Unload Me
    SelectedItem = True
End Function


Private Function zlBindingData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-07-20 11:37:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim I As Long
    '��ʼ������
    With vsSelect
        .Redraw = flexRDNone
        Err = 0: On Error Resume Next
        Set vsSelect.Font = mobjControl.Font
        Set Me.Font = mobjControl.Font
        Err = 0: On Error GoTo 0
        Set .DataSource = mrsBindings
        If mrsBindings.EOF Then .Rows = 2: .Clear 1
        For I = 0 To .Cols - 1
            .ColKey(I) = Trim(.TextMatrix(0, I))
            If UCase(.ColKey(I)) Like "*ID" Then .ColHidden(I) = True
            If InStr(1, "," & mstrHideCols & ",", "," & UCase(.ColKey(I)) & ",") > 0 Then .ColHidden(I) = True
            .AutoSizeMode = flexAutoSizeColWidth
            .AutoSize 0, .Cols - 1
        Next
        .RowHidden(0) = False
        If mblnShowHead = False Then .RowHidden(0) = True
        '�ָ���˳��
        .Redraw = flexRDBuffered
        '�ָ�����ؼ�˳��
        If mstr������ <> "" Then Call zl_vsGrid_Para_Restore(mlngSys, mlngModule, vsSelect, mstr������)
    End With
End Function
Private Sub InitCtrlLocal()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ���ؼ���ʼ���ؼ�λ��
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-07-20 10:04:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngH As Long, lngW As Long, vRect As RECT, sngX As Single, sngY As Single
   
   'ͨ��Api������ؼ������������Ϣ
    Select Case UCase(TypeName(mobjControl))
    Case UCase("VSFlexGrid")
        Call CalcPosition(sngX, sngY, mobjControl)
        lngH = mobjControl.CellHeight
        lngW = mobjControl.CellWidth
        sngY = sngY - lngH
    Case UCase("BILLEDIT")
        Call CalcPosition(sngX, sngY, mobjControl.msfObj)
        lngH = mobjControl.msfObj.CellHeight
        lngW = mobjControl.msfObj.CellWidth
    Case Else
        vRect = GetControlRect(mobjControl.hwnd)
        sngX = vRect.Left - 15
        sngY = vRect.Top
        lngH = mobjControl.Height
        lngW = mobjControl.Width
    End Select

    With mTyCtl_Locale
        .Top = sngY
        .Left = sngX
        .Width = lngW
        .Height = lngH
        .minHeight = vsSelect.RowHeight(0) * 5 'һ������
        .minWidth = .Width
        .ScreenHeight = GetSystemMetrics(SM_CYFULLSCREEN) * 15   '��Ļ���ø߶�
        .ScreenWidth = Screen.Width  ' GetSystemMetrics(SM_CXVSCROLL) * 15 + 75  '��Ļ���ÿ��
    End With
End Sub

Public Sub ReSetWindowsFormLocal()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ô��ڵĴ�С��λ��
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-07-20 10:30:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblColsWidth As Double, dblRowsHeight As Double
    Dim dblTemp As Double
    Dim I As Long
    
    '��λ
    With mTyCtl_Locale
        .DownTop = .Top + .Height
        .DownLeft = .Left
        .DownWidth = .Width
    End With
    
    '�����������Ŀ��
    dblColsWidth = 0
    For I = 0 To vsSelect.Cols - 1
        If Not vsSelect.ColHidden(I) Then
            dblColsWidth = dblColsWidth + vsSelect.ColWidth(I) + Screen.TwipsPerPixelX
        End If
    Next
    dblColsWidth = dblColsWidth + 300
    
    '�����������ĸ߶�
    dblRowsHeight = vsSelect.Cell(flexcpHeight, 0, 0, 0, 0)
    
    dblRowsHeight = (dblRowsHeight) * (vsSelect.Rows) + 100
    If dblRowsHeight < mTyCtl_Locale.minHeight Then dblRowsHeight = mTyCtl_Locale.minHeight
    
    dblColsWidth = IIf(dblColsWidth < mTyCtl_Locale.minWidth, mTyCtl_Locale.minWidth, dblColsWidth)
        
    
    With mTyCtl_Locale
        '���㴰���Y������½Ӹ߶�
        If .ScreenHeight - (.Top + .Height + dblRowsHeight) < 0 Then
            '֤���ؼ�������Ҫ�߶ȱȿؼ����µ�λ��Ҫ��
            If dblRowsHeight < .Top Then
                '֤���ϲ�����װ������,��˿ؼ�����,�����ϲ���
               .DownHeight = dblRowsHeight
               .DownTop = .Top - dblRowsHeight
            ElseIf .Top > .ScreenHeight - (.Top + .Height) Then
                '�����������������,�˷�֧��ʾ������
                .DownTop = 0
                .DownHeight = .Top
            Else '�˷�֧��ʾ������
                .DownHeight = .ScreenHeight - (.Top + .Height)
            End If
        Else
            '֤�������б�����ڿؼ��·���ʾ
            .DownHeight = dblRowsHeight
        End If
        
        '���㴰���Y������������
        If .ScreenWidth - .Left >= dblColsWidth Then
            '������װ�������п�
            .DownWidth = dblColsWidth
            .DownLeft = .Left
        Else
           If .Left + .Width >= dblColsWidth Then
                '������װ�������п�
                .DownLeft = .Left + .Width - dblColsWidth
                .DownWidth = dblColsWidth
           ElseIf .Left + .Width > .ScreenWidth - .Left Then
                '֤����ߴ����ұ�
                .DownWidth = .Left + .Width
                .DownLeft = 0
           Else
                '֤���ұߴ������
                .DownWidth = .ScreenWidth - .Left
                .DownLeft = .Left
           End If
        End If
        
        '���Խ��ж�λ��
        Me.Left = .DownLeft
        Me.Top = .DownTop
        Me.Width = .DownWidth
        Me.Height = .DownHeight
        If vsSelect.Width > dblColsWidth - 300 Then
            vsSelect.ExtendLastCol = True
        Else
            vsSelect.ExtendLastCol = False
        End If
    End With
End Sub

Public Sub CalcPosition(ByRef X As Single, ByRef Y As Single, ByVal objBill As Object, Optional blnNoBill As Boolean = False)
    '----------------------------------------------------------------------
    '���ܣ� ����X,Y��ʵ�����꣬��������Ļ���������
    '������ X---���غ��������
    '       Y---�������������
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    Call ClientToScreen(objBill.hwnd, objPoint)
    If blnNoBill Then
        X = objPoint.X * 15
        Y = objPoint.Y * 15 + objBill.Height
    Else
        X = objPoint.X * 15 + objBill.CellLeft
        Y = objPoint.Y * 15 + objBill.CellTop + objBill.CellHeight
    End If
End Sub

Public Function zl_vsGrid_Para_Save(ByVal lngSys As Long, ByVal lngModule As Long, ByVal vsGrid As VSFlexGrid, ByVal strKey As String, _
    Optional blnǿ�Ʊ��� As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '����:����vsFlex�Ŀ�ȵ�ע���
    '����:vsGrid-��Ӧ������ؼ�
    '     strCaption-������
    '     strKey-����
    '����:����ɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2008/03/03
    '------------------------------------------------------------------------------
    Dim intCol As Integer, strCol As String, strColCaption As String, intRow As Integer
    Dim objDataBase As New clsDatabase
    
    If blnǿ�Ʊ��� = False Then
        zl_vsGrid_Para_Save = True
        If Val(objDataBase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Function
    End If
    
    zl_vsGrid_Para_Save = False
    With vsGrid
        strCol = ""
        For intCol = 0 To .Cols - 1
            strCol = strCol & "|" & .ColKey(intCol) & "," & .ColWidth(intCol) & "," & IIf(.ColHidden(intCol), 1, 0)
        Next
    End With
    If strCol <> "" Then strCol = Mid(strCol, 2)
    '�����ʽ:������,�п�,������|������,�п�,������|...
    objDataBase.SetPara strKey, strCol, lngSys, lngModule
    zl_vsGrid_Para_Save = True
End Function

Public Function zl_vsGrid_Para_Restore(ByVal lngSys As Long, ByVal lngModule As Long, ByVal vsGrid As VSFlexGrid, ByVal strKey As String, _
    Optional blnǿ�ƻָ����� As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '����:�����ݿ��лָ�����Ŀ�ȵ���Ϣ
    '����:vsGrid-��Ӧ������ؼ�
    '     strCaption-������
    '     strKey-����
    '     blnǿ�ƻָ�����-�����Ƿ񽫱���Ĳ���ֵ,����ǿ�ƻָ�
    '����:�ָ��ɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2008/03/03
    '------------------------------------------------------------------------------
    Dim strParaValue As String, intCols As Integer, arrReg As Variant, ArrTemp As Variant, intCol As Integer, intRow As Integer
    Dim intTemp As Integer, strColName As String
    Dim objDataBase As New clsDatabase
    
    If blnǿ�ƻָ����� = False Then
        zl_vsGrid_Para_Restore = True
        If Val(objDataBase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Function
    End If
    strParaValue = objDataBase.GetPara(strKey, lngSys, lngModule)
    
    
    zl_vsGrid_Para_Restore = False
    If strParaValue = "" Then Exit Function
    'strParaValue:�����ʽ:������,�п�,������|������,�п�,������|...
    Err = 0: On Error GoTo errHand:
    arrReg = Split(strParaValue, "|")
    intCols = UBound(arrReg) + 1
    With vsGrid
        For intCol = 0 To intCols - 1
            ArrTemp = Split(arrReg(intCol) & ",,", ",")
            strColName = ArrTemp(0)
            intTemp = .ColIndex(strColName)
            If intTemp <> -1 Then
                .ColWidth(intTemp) = Val(ArrTemp(1))
                If Val(ArrTemp(2)) = 1 Then
                    .ColHidden(intTemp) = True
                Else
                    .ColHidden(intTemp) = False
                End If
                If .ColWidth(intTemp) = 0 Then .ColHidden(intTemp) = True
                .ColPosition(.ColIndex(strColName)) = intCol
            End If
        Next
    End With
    zl_vsGrid_Para_Restore = True
    Exit Function
errHand:
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
        Call SelectedItem
    Case vbKeyEscape
        Unload Me: Exit Sub
    End Select
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With vsSelect
        .Left = ScaleLeft
        .Width = ScaleWidth
        .Top = ScaleTop
        .Height = ScaleHeight
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mstr������ <> "" Then zl_vsGrid_Para_Save mlngSys, mlngModule, vsSelect, mstr������
End Sub

Private Sub vsSelect_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call ReSetWindowsFormLocal
End Sub

Private Sub vsSelect_DblClick()
    Call SelectedItem
End Sub

Private Sub vsSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call SelectedItem
End Sub
