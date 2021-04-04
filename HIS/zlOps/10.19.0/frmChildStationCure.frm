VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmChildStationCure 
   BorderStyle     =   0  'None
   ClientHeight    =   3405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6255
   Icon            =   "frmChildStationCure.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   2145
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   450
      Width           =   3990
      _cx             =   7038
      _cy             =   3784
      Appearance      =   3
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483626
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   270
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmChildStationCure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################
'���������弶��������

Private mblnDataChanged As Boolean
Private mblnReading As Boolean
Private mfrmMain As Object
Private mblnAllowModify As Boolean
Private mlngKey As Long
Private mint������Դ As Integer
Private mlng���˿���id As Long
Private mstr������Դ As String
Private mlngҽ��id As Long
Private mbytMode As Byte
Private mstrPrivs As String
Private mobjStateInfo As CommandBarControl
Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1

Public Event AfterDataChanged()
Public Event AfterMakeCharge()

'######################################################################################################################
'�������Զ�����̻���

Public Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
    
    If mblnReading = False Then
        RaiseEvent AfterDataChanged
    End If
End Property

Public Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Public Function InitData(ByVal frmMain As Object, Optional ByVal blnAllowModify As Boolean = True) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mblnAllowModify = blnAllowModify
    Set mfrmMain = frmMain
    
    
    If ExecuteCommand("��ʼ�ؼ�") = False Then Exit Function
    If ExecuteCommand("��ʼ����") = False Then Exit Function
    
    Call ExecuteCommand("�ؼ�״̬")
    
    DataChanged = False
    
    InitData = True
    
End Function

Public Function RefreshData(ByVal lngKey As Long, _
                            Optional ByVal blnAllowModify As Boolean = True, _
                            Optional ByVal bytMode As Byte = 1, _
                            Optional ByVal str������Դ As String, _
                            Optional ByVal lngҽ��id As Long, _
                            Optional ByVal strPrivs As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������bytMode:1-׼��;2-�Ǽ�
    '���أ�
    '******************************************************************************************************************
    mblnAllowModify = blnAllowModify
    mlngKey = lngKey
    mbytMode = bytMode
    mstr������Դ = str������Դ
    mint������Դ = IIf(mstr������Դ = "סԺ", 2, 1)
    mlngҽ��id = lngҽ��id
    
    mstrPrivs = strPrivs
    
    Call ExecuteCommand("�������")
    Call ExecuteCommand("�ؼ�״̬")
    
    If mlngKey > 0 Then
        If ExecuteCommand("��ȡ����") = False Then Exit Function
    End If

    DataChanged = False
    
    RefreshData = True
    
End Function

Public Function ValidData() As Boolean
    '******************************************************************************************************************
    '���ܣ�У��༭���ݵ���Ч��
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngLoop As Long
    
    With vsf(0)
        For lngLoop = 1 To .Rows - 1
            If lngLoop <> .Rows - 1 Then
                If .RowData(lngLoop) = 0 Then
                    ShowSimpleMsg "�� " & lngLoop & " ���������벻����������������Ч����Ŀ��"
                    Call LocationGrid(vsf(0), lngLoop, .ColIndex("����"))
                    Exit Function
                End If
            End If
            
            If .RowData(lngLoop) > 0 Then
                If IsNumeric(.TextMatrix(lngLoop, .ColIndex("����"))) = False And .TextMatrix(lngLoop, .ColIndex("����")) <> "" Then
                    ShowSimpleMsg "�� " & lngLoop & " ���������벻��������������Ϊ��ֵ�ͣ�"
                    Call LocationGrid(vsf(0), lngLoop, .ColIndex("����"))
                    Exit Function
                End If
                
                
                If Val(.TextMatrix(lngLoop, .ColIndex("����"))) > 99999999 Then
                    ShowSimpleMsg "�� " & lngLoop & " ������̫�󣬱�������[0-99999999]�ڵ���ֵ��"
                    Call LocationGrid(vsf(0), lngLoop, .ColIndex("������"))
                    Exit Function
                End If
                                
                If Val(.TextMatrix(lngLoop, .ColIndex("ִ�п���id"))) <= 0 Then
                    ShowSimpleMsg "�� " & lngLoop & " ��û��ָ��ִ�п��ң�"
                    Call LocationGrid(vsf(0), lngLoop, .ColIndex("ִ�п���"))
                    Exit Function
                End If
            End If
        Next
    End With
    
    ValidData = True
    
End Function

Public Function ClearData() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Call ExecuteCommand("�������")
    
    ClearData = True
    
End Function

Public Function SaveData(ByRef rsSQL As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strSQL As String
    Dim lngLoop As Long
    
    On Error GoTo errHand

    '
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "ZL_���������Ƽ�_DELETE(" & mlngKey & ")"
    Call SQLRecordAdd(rsSQL, strSQL)
    
    
    With vsf(0)
        For lngLoop = 1 To .Rows - 1
            If .RowData(lngLoop) > 0 Then
                strSQL = "ZL_���������Ƽ�_INSERT(" & mlngKey & "," & lngLoop & "," & Val(.RowData(lngLoop)) & "," & Val(.TextMatrix(lngLoop, .ColIndex("����"))) & "," & Val(.TextMatrix(lngLoop, .ColIndex("ִ�п���id"))) & ")"
                Call SQLRecordAdd(rsSQL, strSQL)
            End If
        Next
    End With
            
    SaveData = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    Call CommandBarInit(cbsMain)
    
    cbsMain.Options.LargeIcons = False
    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ

    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    
    '------------------------------------------------------------------------------------------------------------------
    '����������:������������

    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched

    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "����", , , xtpButtonIconAndCaption)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "ɾ��", , , xtpButtonIconAndCaption)
        
    Set objPopup = NewToolBar(objBar, xtpControlPopup, conMenu_Edit_MakeCharge, "����", True, , xtpButtonIconAndCaption)
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Preferences, "����", True, , xtpButtonIconAndCaption)
    
    Set mobjStateInfo = NewToolBar(objBar, xtpControlLabel, 0, "")
    mobjStateInfo.Flags = xtpFlagRightAlign
End Function

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim objArray As Variant
    Dim blnAllowModify As Boolean
    
    On Error GoTo errHand
    
    Call SQLRecord(rsSQL)
    
    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ�ؼ�"

        Set mclsVsf = New clsVsf
        With mclsVsf
            Call .Initialize(Me.Controls, vsf(0), True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn

            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
            Call .AppendColumn("����", 2100, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("���", 900, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("����", 900, flexAlignLeftCenter, flexDTString, "0.00", , True)
            Call .AppendColumn("��λ", 600, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("ִ�п���", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("ִ�п���id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("����", 450, flexAlignLeftCenter, flexDTBoolean, "", , True)
            Call .AppendColumn("", 15, flexAlignLeftCenter, flexDTString, "", , True)
            
            .AppendRows = True
        End With
        
        Call InitCommandBar
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"
        
    '------------------------------------------------------------------------------------------------------------------
    Case "�������"
    
        mclsVsf.ClearGrid
        mobjStateInfo.Caption = ""
        cbsMain.RecalcLayout
        
    '--------------------------------------------------------------------------------------------------------------
    Case "�ؼ�״̬"
    
        blnAllowModify = mblnAllowModify
        If mlngKey = 0 Then blnAllowModify = False
        
        With mclsVsf
            
            If blnAllowModify Then
                Call .ModifyColumn(.ColIndex("ͼ��"), "", 255, flexAlignCenterCenter, flexDTString, "", "[ָʾ��]", False)
                
                Call .InitializeEdit(True, True, True)
                
                Call .InitializeEditColumn(.ColIndex("����"), True, vbVsfEditCommand)
                Call .InitializeEditColumn(.ColIndex("ִ�п���"), True, vbVsfEditCombox)
                Call .InitializeEditColumn(.ColIndex("����"), True, vbVsfEditText)
                
                .IndicatorCol = 0
                Set .IndicatorIcon = frmPubResource.GetImageList(16).ListImages("��ǰ").Picture
            Else
                Call .ModifyColumn(.ColIndex("ͼ��"), "", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
                Call .InitializeEdit(False, False, False)
            End If
        End With
    
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ����"

        mclsVsf.ClearGrid
        mobjStateInfo.Caption = ""
        
        gstrSQL = "SELECT B.���㵥λ As ��λ," & _
                    "A.�շ�ϸĿID As ID," & _
                    "C.���� AS ִ�п���," & _
                    "B.����,B.���," & _
                    "A.ִ�п���id," & _
                    "A.����," & _
                    "A.No,a.��¼����,Decode(A.No,Null,0,1) As ���� " & _
                    "FROM ���������Ƽ� A,�շ���ĿĿ¼ B,���ű� C " & _
                    "WHERE A.�շ�ϸĿID=B.ID " & _
                        "AND C.ID=A.ִ�п���id " & _
                        "AND A.��¼id=[1] ORDER BY A.���"
                            
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
        If rs.BOF = False Then
            If IsNull(rs("No").Value) = False Then
                Select Case rs("��¼����").Value
                Case 1
                    mobjStateInfo.Caption = "�������շѵ������ݺţ�" & rs("No").Value
                Case 2
                    mobjStateInfo.Caption = "�����ɼ��ʵ������ݺţ�" & rs("No").Value
                End Select
            Else
                mobjStateInfo.Caption = ""
            End If
            Call mclsVsf.LoadGrid(rs)
        End If
        cbsMain.RecalcLayout
    '------------------------------------------------------------------------------------------------------------------
    Case "�շ�ִ�п���"
        
        With vsf(0)
        
            gstrSQL = GetPublicSQL(SQL.�շ�ִ�п���)
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, GetDefaultDept(0, mint������Դ), Val(.RowData(.Row)), mlng���˿���id, UserInfo.����ID)
            
            If rs.BOF = False Then
                .TextMatrix(.Row, .ColIndex("ִ�п���")) = zlCommFun.NVL(rs("����").Value)
                .TextMatrix(.Row, .ColIndex("ִ�п���id")) = zlCommFun.NVL(rs("ID").Value)
                
                .ColComboList(.ColIndex("ִ�п���")) = .BuildComboList(rs, "����", "ID")
                
            Else
                .TextMatrix(.Row, .ColIndex("ִ�п���")) = UserInfo.��������
                .TextMatrix(.Row, .ColIndex("ִ�п���id")) = UserInfo.����ID
                .ColComboList(.ColIndex("ִ�п���")) = " |"
            End If
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "���Ʋο�����"
        
        With vsf(0)
            gstrSQL = GetPublicSQL(SQL.��������ѡ��)

            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngKey)
            
            If ShowPubSelect(Me, Nothing, 3, "����,900,0,;����,2400,0,;���,900,0,;����,1200,2,;��λ,810,0,", mfrmMain.Name & "\��������ѡ��", "�����������б���ѡ���������Ʋο�", rsData, rs, 8790, 4500, False, , , True) = 1 Then
                
                gstrSQL = GetPublicSQL(SQL.�������Ʋο�)
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, Val(rs("ID").Value))
                If rs.BOF = False Then
                    mclsVsf.ClearGrid
                    
                    Do While Not rs.EOF
                        If Val(.RowData(.Rows - 1)) > 0 Then .Rows = .Rows + 1
                        .Row = .Rows - 1
                        Call mclsVsf.LoadGridRow(.Row, rs)

                        Call ExecuteCommand("�շ�ִ�п���")
                        
                        rs.MoveNext
                    Loop
                    
                    DataChanged = True
                End If
            End If
        
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "���������շѵ�"
        If DataChanged Then
            If SaveData(rsSQL) Then
                If SQLRecordExecute(rsSQL, mfrmMain.Caption) = False Then Exit Function
                DataChanged = False
            End If
        End If
        
        If MsgBox("��������ɣ�����Զ�ɾ�������ϣ�ȷ��Ҫ���������������շѵ����µ�����", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Function
        
        strTmp = MakeChargeBill(mlngҽ��id, 1, "����", False, mstrPrivs)
        If strTmp <> "" Then
            ShowSimpleMsg "�����շѵ��Ѿ����ɣ����ݺţ�" & strTmp
            
            mobjStateInfo.Caption = "�������շѵ������ݺţ�" & strTmp
            cbsMain.RecalcLayout
            
            RaiseEvent AfterMakeCharge
        End If
        
        Exit Function
        
    '------------------------------------------------------------------------------------------------------------------
    Case "�������Ƽ��ʵ�"
        If DataChanged Then
            If SaveData(rsSQL) Then
                If SQLRecordExecute(rsSQL, mfrmMain.Caption) = False Then Exit Function
                DataChanged = False
            End If
        End If
            
            
        If MsgBox("��������ɣ�����Զ�ɾ�������ϣ�ȷ��Ҫ�������������ɼ��ʵ����µ�����", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Function
        
        strTmp = MakeChargeBill(mlngҽ��id, 2, "����", False, mstrPrivs)
        If strTmp <> "" Then
            ShowSimpleMsg "���Ƽ��ʵ��Ѿ����ɣ����ݺţ�" & strTmp
            mobjStateInfo.Caption = "�����ɼ��ʵ������ݺţ�" & strTmp
            cbsMain.RecalcLayout
            
            RaiseEvent AfterMakeCharge
        End If
        
        Exit Function
    '------------------------------------------------------------------------------------------------------------------
    Case "����������ѵ�"
    
        If DataChanged Then
            If SaveData(rsSQL) Then
                If SQLRecordExecute(rsSQL, mfrmMain.Caption) = False Then Exit Function
                DataChanged = False
            End If
        End If
            
        If MsgBox("��������ɣ�����Զ�ɾ�������ϣ�ȷ��Ҫ����������������ѵ����µ�����", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Function
        
        strTmp = MakeChargeBill(mlngҽ��id, 2, "����", True, mstrPrivs)
        If strTmp <> "" Then
            ShowSimpleMsg "������ѵ��Ѿ����ɣ����ݺţ�" & strTmp
            mobjStateInfo.Caption = "�����ɼ��ʵ������ݺţ�" & strTmp
            cbsMain.RecalcLayout
            
            RaiseEvent AfterMakeCharge
        End If
        
        Exit Function
    End Select
    
    ExecuteCommand = True
    
    Exit Function
    
    '������
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

Public Property Get Body(ByVal lngIndex As Long) As Object
    Set Body = vsf
End Property

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim intRow As Integer
    
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Preferences                       '����
    
        Call ExecuteCommand("���Ʋο�����")
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem                           '����
        
        Call mclsVsf.AppendRow
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete                            'ɾ��
        
        Call mclsVsf.DeleteRow(vsf(0).Row)

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MakeCharge * 2# + 1

        Call ExecuteCommand("���������շѵ�")

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MakeCharge * 2# + 2

        Call ExecuteCommand("�������Ƽ��ʵ�")
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MakeCharge * 2# + 3
                
        Call ExecuteCommand("����������ѵ�")
        
    End Select
End Sub

Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim objControl As CommandBarControl

    If CommandBar.Parent Is Nothing Then Exit Sub
    
    Select Case CommandBar.Parent.ID
    Case conMenu_Edit_MakeCharge

        With CommandBar.Controls

            .DeleteAll

            Set objControl = .Add(xtpControlButton, conMenu_Edit_MakeCharge * 2 + 1, "�շѵ���(&1)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_MakeCharge * 2 + 2, "���ʵ���(&2)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_MakeCharge * 2 + 3, "��ķ���(&3)")
            With cbsMain.KeyBindings
                .Add FCONTROL, vbKeyN, conMenu_Edit_MakeCharge * 2 + 1
                .Add FCONTROL, vbKeyB, conMenu_Edit_MakeCharge * 2 + 2
            End With
            
        End With
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long
    Dim lngTop  As Long
    Dim lngRight  As Long
    Dim lngBottom  As Long

    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    
    '���������ؼ�Resize����
    vsf(0).Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
    mclsVsf.AppendRows = True
    
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnAllowModify As Boolean
    
    On Error GoTo errHand
    
    blnAllowModify = mblnAllowModify
    If mlngKey = 0 Then blnAllowModify = False
    
    With vsf(0)
        Select Case Control.ID
        Case conMenu_Edit_Preferences
        
            Control.Enabled = blnAllowModify And (mbytMode = 2 Or mbytMode = 1)
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_NewItem
            
            Control.Enabled = blnAllowModify And Val(.RowData(.Rows - 1)) > 0
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Delete
            
            Control.Enabled = blnAllowModify And Val(.RowData(.Row)) > 0
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_MakeCharge
            
            Control.Visible = IsPrivs(mstrPrivs, "���ɸ���")
            Control.Enabled = blnAllowModify And Control.Visible And mbytMode = 2 And (Val(.RowData(1)) > 0 Or .Rows > 2)
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_MakeCharge * 2 + 1
            
            Control.Visible = (mstr������Դ = "����" And IsPrivs(mstrPrivs, "���ɸ���"))
            Control.Enabled = blnAllowModify And Control.Visible And mbytMode = 2 And (Val(.RowData(1)) > 0 Or .Rows > 2)
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_MakeCharge * 2 + 2, conMenu_Edit_MakeCharge * 2 + 3
            
            Control.Visible = IsPrivs(mstrPrivs, "���ɸ���")
            Control.Enabled = blnAllowModify And Control.Visible And mbytMode = 2 And (Val(.RowData(1)) > 0 Or .Rows > 2)
            
        End Select
    End With
errHand:
End Sub

'######################################################################################################################
'���������弰��ؼ����¼�����

Private Sub Form_Resize()
    On Error Resume Next
    
    vsf(0).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    mclsVsf.AppendRows = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjStateInfo = Nothing
    Set mclsVsf = Nothing
End Sub

Private Sub mclsVsf_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    DataChanged = True
End Sub

Private Sub mclsVsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    With vsf(0)
        Cancel = (Val(.RowData(Row)) <= 0)
    End With
End Sub

Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)

    '�༭����
    Call mclsVsf.AfterEdit(Row, Col)
    DataChanged = True
    
    With vsf(Index)
        Select Case Col
        Case .ColIndex("ִ�п���")
            .TextMatrix(Row, .ColIndex("ִ�п���id")) = .ComboData
            .TextMatrix(Row, .ColIndex("ִ�п���")) = .Cell(flexcpTextDisplay, Row, .ColIndex("ִ�п���"))
        End Select
    End With
    
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim bytRet As Byte
    
    With vsf(0)
        If Col = .ColIndex("����") Then
            
            gstrSQL = GetPublicSQL(SQL.������Ŀѡ��)
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)
            bytRet = ShowPubSelect(Me, vsf(0), 3, "����,1200,0,0;����,3000,0,0;���,900,0,0;��λ,900,0,0", Me.Name & "\������Ŀѡ��", "����±���ѡ��һ��������Ŀ", rsData, rs, 8790, 4500, , Val(.RowData(Row)))

            If bytRet = 1 Then
            
                If mclsVsf.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                    ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                    Exit Sub
                End If
    
                .EditText = zlCommFun.NVL(rs("����").Value)
                .TextMatrix(Row, mclsVsf.ColIndex("����")) = zlCommFun.NVL(rs("����").Value)
                .TextMatrix(Row, mclsVsf.ColIndex("���")) = zlCommFun.NVL(rs("���").Value)
                .TextMatrix(Row, mclsVsf.ColIndex("��λ")) = zlCommFun.NVL(rs("��λ").Value)
                .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
    
                Call ExecuteCommand("�շ�ִ�п���")
                
                DataChanged = True
                Call mclsVsf.LocationNextCell
            End If
            
            Call mclsVsf.SetFocus(, , True)
        End If
    End With
End Sub

Private Sub vsf_ChangeEdit(Index As Integer)
    With vsf(Index)
        Select Case .Col
        Case .ColIndex("����")
            .TextMatrix(.Row, .Col) = .EditText
        End Select

    End With
End Sub

Private Sub vsf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call mclsVsf.KeyDown(KeyCode, Shift)
End Sub

Private Sub vsf_KeyDownEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim strTmp As String
    Dim strText As String
    Dim bytMode As Byte
    
    With vsf(0)
        If KeyCode = vbKeyReturn Then
            If Col = .ColIndex("����") Then
                
                If InStr(vsf(0).EditText, "'") > 0 Then
                    KeyCode = 0
                    vsf(0).EditText = ""
                    Exit Sub
                End If

                strText = UCase(vsf(0).EditText)
                bytMode = GetApplyMode(strText)
                
                gstrSQL = GetPublicSQL(SQL.������Ŀ����, bytMode)

                strText = strText & "%"
                If ParamInfo.��Ŀ����ƥ�䷽ʽ = 1 Then
                    strTmp = strText
                Else
                    strTmp = "%" & strText
                End If
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp)

                If ShowPubSelect(Me, vsf(0), 2, "����,1200,0,0;����,3000,0,0;���,900,0,0;��λ,900,0,0", Me.Name & "\������Ŀ����", "����±���ѡ��һ��������Ŀ", rsData, rs, 8790, 4500, , Val(.RowData(Row))) = 1 Then

                    If mclsVsf.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                        ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                        Exit Sub
                    End If

                    .EditText = zlCommFun.NVL(rs("����").Value)
                    .TextMatrix(Row, .ColIndex("����")) = zlCommFun.NVL(rs("����").Value)
                    .TextMatrix(Row, .ColIndex("��λ")) = zlCommFun.NVL(rs("��λ").Value)
                    .TextMatrix(Row, .ColIndex("���")) = zlCommFun.NVL(rs("���").Value)
                    
                    .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)

                    Call ExecuteCommand("�շ�ִ�п���")
                
                    DataChanged = True
                    Call mclsVsf.LocationNextCell
                Else
                    KeyCode = 0

                    .Cell(flexcpData, Row, Col) = .Cell(flexcpData, Row, Col)
                    .EditText = .Cell(flexcpData, Row, Col)
                    .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                End If
                
                Call mclsVsf.SetFocus(, , True)
            
            Else
                Call mclsVsf.LocationNextCell
            End If
        Else
            DataChanged = True
        End If
    End With
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    '�༭����
    Call mclsVsf.KeyPress(KeyAscii)
End Sub

Private Sub vsf_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    '�༭����
    Call mclsVsf.KeyPressEdit(KeyAscii)
End Sub

Private Sub vsf_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Call mclsVsf.AutoAddRow(vsf(Index).MouseRow, vsf(Index).MouseCol)
    End Select
End Sub

Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    
    Select Case Button
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '�����˵�����
        Call SendLMouseButton(vsf(Index).hWnd, X, Y)
        Set cbrPopupBar = CopyMenu(mfrmMain.cbsMain, 3)
        cbrPopupBar.ShowPopup
    End Select
End Sub

Private Sub vsf_SetupEditWindow(Index As Integer, ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    '�༭����
    Call mclsVsf.EditSelAll
End Sub

Private Sub vsf_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '�༭����
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '�༭����
    Call mclsVsf.ValidateEdit(Col, Cancel)
End Sub






