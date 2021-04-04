VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmChildStationDrug 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   2145
      Index           =   0
      Left            =   285
      TabIndex        =   0
      Top             =   675
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
      ForeColorSel    =   16777215
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
      Rows            =   2
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
Attribute VB_Name = "frmChildStationDrug"
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
Private mbytMode As Byte
Private mstrPrivs As String
Private mstr������Դ As String
Private mlng���˿���id As Long
Private mlngҽ��id As Long
Private mobjStateInfo As CommandBarControl
Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1
Private mint������Դ As Integer
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
    '������
    '���أ�
    '******************************************************************************************************************
    mbytMode = bytMode
    mlngҽ��id = lngҽ��id
    mstr������Դ = str������Դ
    mint������Դ = IIf(str������Դ = "סԺ", 2, 1)
    
    mstrPrivs = strPrivs
    mblnAllowModify = blnAllowModify
    mlngKey = lngKey
    
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
                If Val(.RowData(lngLoop)) = 0 Then
                    ShowSimpleMsg "�� " & lngLoop & " ���������벻����������������Ч��ҩƷ��Ŀ��"
                    Call LocationGrid(vsf(0), lngLoop, .ColIndex("��ҩ����"))
                    Exit Function
                End If
            End If
            
            If Val(.RowData(lngLoop)) > 0 Then
                If IsNumeric(.TextMatrix(lngLoop, .ColIndex("׼������"))) = False And mbytMode = 1 And .TextMatrix(lngLoop, .ColIndex("׼������")) <> "" Then
                    ShowSimpleMsg "�� " & lngLoop & " ���������벻������ҩƷ��׼����������Ϊ��ֵ�ͣ�"
                    Call LocationGrid(vsf(0), lngLoop, .ColIndex("׼������"))
                    Exit Function
                End If
                
                If Val(.TextMatrix(lngLoop, .ColIndex("׼������"))) < 0 And mbytMode = 1 Then
                    ShowSimpleMsg "�� " & lngLoop & " ���������벻��������������ҩƷ��׼������[0-99999999]��"
                    Call LocationGrid(vsf(0), lngLoop, .ColIndex("׼������"))
                    Exit Function
                End If
                
                If Val(.TextMatrix(lngLoop, .ColIndex("׼������"))) > 999999999 And mbytMode = 1 Then
                    ShowSimpleMsg "�� " & lngLoop & " ������̫������[0-999999999]��Χ�ڵ���ֵ��"
                    Call LocationGrid(vsf(0), lngLoop, .ColIndex("׼������"))
                    Exit Function
                End If
                
                If IsNumeric(.TextMatrix(lngLoop, .ColIndex("ʵ������"))) = False And mbytMode = 2 And .TextMatrix(lngLoop, .ColIndex("ʵ������")) <> "" Then
                    ShowSimpleMsg "�� " & lngLoop & " ���������벻������ҩƷ��ʵ����������Ϊ��ֵ�ͣ�"
                    Call LocationGrid(vsf(0), lngLoop, .ColIndex("ʵ������"))
                    Exit Function
                End If
                
                If Val(.TextMatrix(lngLoop, .ColIndex("ʵ������"))) < 0 And mbytMode = 2 Then
                    ShowSimpleMsg "�� " & lngLoop & " ���������벻��������������ҩƷ��ʵ������[0-99999999]��"
                    Call LocationGrid(vsf(0), lngLoop, .ColIndex("ʵ������"))
                    Exit Function
                End If
                
                If Val(.TextMatrix(lngLoop, .ColIndex("ʵ������"))) > 99999999 And mbytMode = 2 Then
                    ShowSimpleMsg "�� " & lngLoop & " ������̫������[0-999999999]��Χ�ڵ���ֵ��"
                    Call LocationGrid(vsf(0), lngLoop, .ColIndex("ʵ������"))
                    Exit Function
                End If
                
                If Val(.TextMatrix(lngLoop, .ColIndex("ִ�п���id"))) <= 0 Then
                    ShowSimpleMsg "�� " & lngLoop & " ��û��ָ��ִ�п��ң�"
                    Call LocationGrid(vsf(0), lngLoop, .ColIndex("ִ�п���id"))
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
    strSQL = "ZL_����������ҩ_DELETE(" & mlngKey & ")"
    Call SQLRecordAdd(rsSQL, strSQL)
    
    
    With vsf(0)
        For lngLoop = 1 To .Rows - 1
            If .RowData(lngLoop) > 0 Then
                strSQL = "ZL_����������ҩ_INSERT(" & mlngKey & ",'" & zlCommFun.GetNeedName(.TextMatrix(lngLoop, .ColIndex("��ҩ����"))) & "'," & lngLoop & "," & Val(.RowData(lngLoop)) & ",'" & .TextMatrix(lngLoop, .ColIndex("ҩƷ����")) & "'," & Val(.TextMatrix(lngLoop, .ColIndex("׼������"))) & "," & Val(.TextMatrix(lngLoop, .ColIndex("ʵ������"))) & "," & Val(.TextMatrix(lngLoop, .ColIndex("ִ�п���id"))) & ")"
                Call SQLRecordAdd(rsSQL, strSQL)
            End If
        Next
    End With

    SaveData = True
    
    Exit Function
    
    '
    '------------------------------------------------------------------------------------------------------------------
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
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Default, "ȱʡ", , conMenu_Edit_Modify, xtpButtonIconAndCaption)
    Set objPopup = NewToolBar(objBar, xtpControlPopup, conMenu_Edit_MakeCharge, "���ɷ���", True, , xtpButtonIconAndCaption)
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Preferences, "ѡ�񷽰�", True, , xtpButtonIconAndCaption)

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
            Call .AppendColumn("��ҩ����", 1080, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("ҩƷ����", 2100, flexAlignLeftCenter, flexDTString, "", , True)
            
            Call .AppendColumn("���", 1500, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("׼������", 900, flexAlignRightCenter, flexDTString, "0.00", , True)
            Call .AppendColumn("ʵ������", 900, flexAlignRightCenter, flexDTString, "0.00", , True)
            Call .AppendColumn("��λ", 600, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("ִ�п���", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("ִ�п���id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("��������", 0, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("��������id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("���", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("���ÿ��", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("", 15, flexAlignLeftCenter, flexDTString, "", , True)
                        
            .AppendRows = True
        End With
        
        Call InitCommandBar
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"
    
    '------------------------------------------------------------------------------------------------------------------
    Case "�������"
    
        mclsVsf.ClearGrid
        mobjStateInfo.Caption = " "
        cbsMain.RecalcLayout
        
    '--------------------------------------------------------------------------------------------------------------
    Case "�ؼ�״̬"
    
        blnAllowModify = mblnAllowModify And (IsPrivs(mstrPrivs, "��ҩ׼��") Or IsPrivs(mstrPrivs, "��ҩ�Ǽ�"))
        If mlngKey = 0 Then blnAllowModify = False
        
        With mclsVsf

            If blnAllowModify Then
                Call .ModifyColumn(.ColIndex("ͼ��"), "", 255, flexAlignCenterCenter, flexDTString, "", "[ָʾ��]", False)
                
                Call .InitializeEdit(True, True, True)
                Call .InitializeEditColumn(.ColIndex("��ҩ����"), True, vbVsfEditCombox, " |")
                Call .InitializeEditColumn(.ColIndex("ҩƷ����"), True, vbVsfEditCommand)
                Call .InitializeEditColumn(.ColIndex("ִ�п���"), True, vbVsfEditCombox)
                Call .InitializeEditColumn(.ColIndex("��������"), True, vbVsfEditCombox)
                
                If mbytMode = 1 Then
                    Call .InitializeEditColumn(.ColIndex("׼������"), True, vbVsfEditText)
                    Call .InitializeEditColumn(.ColIndex("ʵ������"), False, vbVsfEditText)
                Else
                    Call .InitializeEditColumn(.ColIndex("׼������"), False, vbVsfEditText)
                    Call .InitializeEditColumn(.ColIndex("ʵ������"), True, vbVsfEditText)
                End If
                
                .IndicatorCol = 0
                Set .IndicatorIcon = frmPubResource.GetImageList(16).ListImages("��ǰ").Picture
                
                gstrSQL = "Select ����||'-'||���� As ���� From ������ҩ���� Order By ����"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)
                If rs.BOF = False Then
                    .DropColData(.ColIndex("��ҩ����")) = vsf(0).BuildComboList(rs, "����", "����", RGB(255, 255, 255))
                End If
                
            Else
                Call .InitializeEdit(False, False, False)
                Call .ModifyColumn(.ColIndex("ͼ��"), "", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
            End If
        End With
    
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ����"
            
        mclsVsf.ClearGrid
        mobjStateInfo.Caption = " "
        
        gstrSQL = "SELECT a.���," & _
                        "f.���㵥λ As ��λ," & _
                        "a.���� As ��ҩ����," & _
                        "a.ҩƷID As ID," & _
                        "a.ҩƷ����," & _
                        "a.׼������ AS ׼������," & _
                        "a.ʹ������ AS ʵ������," & _
                        "D.���� AS ִ�п���," & _
                        "a.ִ�п���id,F.���," & _
                        "f.��� " & _
                "FROM ����������ҩ A,ҩƷ��� B,����������¼ C,���ű� D,�շ���ĿĿ¼ F " & _
                "WHERE A.ҩƷid=B.ҩƷid " & _
                    "AND B.ҩƷid=F.ID " & _
                    "AND A.ִ�п���id=D.ID And C.ID=a.��¼id " & _
                    "AND A.��¼id=[1] " & _
                "ORDER BY A.���"
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
        If rs.BOF = False Then Call mclsVsf.LoadGrid(rs)
                
        mobjStateInfo.Caption = ""
        gstrSQL = "Select No,��¼���� From ������������ Where ��¼id=[1] And ��������=[2]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey, 1)
        If rs.BOF = False Then
            If IsNull(rs("No").Value) = False Then
                Select Case rs("��¼����").Value
                Case 1
                    mobjStateInfo.Caption = "�������շѵ������ݺţ�" & rs("No").Value
                Case 2
                    mobjStateInfo.Caption = "�����ɼ��ʵ������ݺţ�" & rs("No").Value
                End Select
            End If
        End If
        cbsMain.RecalcLayout
    '------------------------------------------------------------------------------------------------------------------
    Case "�շ�ִ�п���"
        
        With vsf(0)
        
            gstrSQL = GetPublicSQL(SQL.�շ�ִ�п���, CStr(varParam(0)))
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, GetDefaultDept(CStr(varParam(0)), mint������Դ), Val(.RowData(.Row)), mlng���˿���id, UserInfo.����ID)
            
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
    Case "ҩƷ�ο�����"
        
        With vsf(0)
            gstrSQL = GetPublicSQL(SQL.������ҩѡ��)

            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngKey)
            
            If ShowPubSelect(Me, Nothing, 3, "����,990,0,;����,900,0,;����,2400,0,;���,900,0,;����,1200,2,;��λ,810,0,", mfrmMain.Name & "\������ҩѡ��", "�����������б���ѡ��������ҩ�ο�", rsData, rs, 8790, 4500, False, , , True) = 1 Then
                
                gstrSQL = GetPublicSQL(SQL.������ҩ�ο�)
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, Val(rs("ID").Value))
                If rs.BOF = False Then
                    mclsVsf.ClearGrid
                    
                    Do While Not rs.EOF
                        If Val(.RowData(.Rows - 1)) > 0 Then .Rows = .Rows + 1
                        .Row = .Rows - 1
                        Call mclsVsf.LoadGridRow(.Row, rs)

                        Call ExecuteCommand("�շ�ִ�п���", zlCommFun.NVL(rs("���").Value))
                        Call CheckStorage(.Row)
                        
                        rs.MoveNext
                    Loop
                    DataChanged = True
                End If
            End If
        
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "����ҩƷ�շѵ�"
    
        If DataChanged Then
            If SaveData(rsSQL) Then
                If SQLRecordExecute(rsSQL, mfrmMain.Caption) = False Then Exit Function
                DataChanged = False
            End If
        End If
        
        If MsgBox("��������ɣ�����Զ�ɾ�������ϣ�ȷ��Ҫ��������ҩ�����շѵ����µ�����", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Function
        
        strTmp = MakeChargeBill(mlngҽ��id, 1, "��ҩ", False, mstrPrivs)
        If strTmp <> "" Then
            ShowSimpleMsg "��ҩ�շѵ��Ѿ����ɣ����ݺţ�" & strTmp
            mobjStateInfo.Caption = "�������շѵ������ݺţ�" & strTmp
            cbsMain.RecalcLayout
            RaiseEvent AfterMakeCharge
        End If
        
        Exit Function
        
    '------------------------------------------------------------------------------------------------------------------
    Case "����ҩƷ���ʵ�"
        If DataChanged Then
            If SaveData(rsSQL) Then
                If SQLRecordExecute(rsSQL, mfrmMain.Caption) = False Then Exit Function
                DataChanged = False
            End If
        End If
        
        If MsgBox("��������ɣ�����Զ�ɾ�������ϣ�ȷ��Ҫ��������ҩ���ɼ��ʵ����µ�����", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Function
        
        strTmp = MakeChargeBill(mlngҽ��id, 2, "��ҩ", False, mstrPrivs)
        If strTmp <> "" Then
            ShowSimpleMsg "��ҩ���ʵ��Ѿ����ɣ����ݺţ�" & strTmp
            mobjStateInfo.Caption = "�����ɼ��ʵ������ݺţ�" & strTmp
            cbsMain.RecalcLayout
            RaiseEvent AfterMakeCharge
        End If
        
        Exit Function
    '------------------------------------------------------------------------------------------------------------------
    Case "����ҩƷ��ѵ�"
        If DataChanged Then
            If SaveData(rsSQL) Then
                If SQLRecordExecute(rsSQL, mfrmMain.Caption) = False Then Exit Function
                DataChanged = False
            End If
        End If
        
        If MsgBox("��������ɣ�����Զ�ɾ�������ϣ�ȷ��Ҫ��������ҩ������ѵ����µ�����", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Function
        
        strTmp = MakeChargeBill(mlngҽ��id, 2, "��ҩ", True, mstrPrivs)
        If strTmp <> "" Then
            ShowSimpleMsg "��ҩ��ѵ��Ѿ����ɣ����ݺţ�" & strTmp
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

Private Function CheckStorage(ByVal intRow As Integer) As Boolean
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim byt��鷽ʽ As Byte
    
    With vsf(0)
            
        gstrSQL = GetPublicSQL(SQL.�����鷽ʽ, .TextMatrix(intRow, .ColIndex("���")))
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, .TextMatrix(intRow, .ColIndex("���")))
        If rs.BOF = False Then byt��鷽ʽ = zlCommFun.NVL(rs("��鷽ʽ").Value, 0)
        If byt��鷽ʽ <> 0 Then
        
            byt��鷽ʽ = zlCommFun.NVL(rs("��鷽ʽ").Value, 0)
            
            gstrSQL = GetPublicSQL(SQL.����������)
            
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, Val(.RowData(intRow)), Val(.TextMatrix(intRow, .ColIndex("ִ�п���id"))))
            If rs.BOF = False Then .TextMatrix(intRow, .ColIndex("���ÿ��")) = zlCommFun.NVL(rs("���").Value, 0)
            
            Call PromptStorageWarn(Val(.TextMatrix(intRow, .ColIndex("׼������"))), Val(.TextMatrix(intRow, .ColIndex("���ÿ��"))), .TextMatrix(intRow, .ColIndex("ҩƷ����")), .TextMatrix(intRow, .ColIndex("ִ�п���")), .TextMatrix(intRow, .ColIndex("��λ")), byt��鷽ʽ)
            
        End If
    End With
    
    CheckStorage = True
End Function

Public Property Get Body(ByVal lngIndex As Long) As Object
    Set Body = vsf
End Property

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim intRow As Integer
    
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Preferences                       '����
    
        Call ExecuteCommand("ҩƷ�ο�����")
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem                           '����
        
        Call mclsVsf.AppendRow
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete                            'ɾ��
        
        Call mclsVsf.DeleteRow(vsf(0).Row)
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Default                           'ȱʡ

        With vsf(0)
            For intRow = 1 To .Rows - 1
                .Cell(flexcpText, intRow, .ColIndex("ʵ������")) = .Cell(flexcpText, intRow, .ColIndex("׼������"))
            Next
            DataChanged = True
        End With
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MakeCharge * 2# + 1

        Call ExecuteCommand("����ҩƷ�շѵ�")

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MakeCharge * 2# + 2

        Call ExecuteCommand("����ҩƷ���ʵ�")
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MakeCharge * 2# + 3
                
        Call ExecuteCommand("����ҩƷ��ѵ�")

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
        
    Select Case Control.ID
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem, conMenu_Edit_Preferences
        
        Control.Visible = IsPrivs(mstrPrivs, "��ҩ׼��") Or IsPrivs(mstrPrivs, "��ҩ�Ǽ�")
        Control.Enabled = blnAllowModify And Control.Visible
        
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete
        
        Control.Visible = IsPrivs(mstrPrivs, "��ҩ׼��") Or IsPrivs(mstrPrivs, "��ҩ�Ǽ�")
        Control.Enabled = blnAllowModify And Control.Visible And Val(vsf(0).RowData(vsf(0).Row)) > 0
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Default
        
        Control.Visible = IsPrivs(mstrPrivs, "��ҩ�Ǽ�")
        Control.Enabled = blnAllowModify And mbytMode = 2 And Control.Visible
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MakeCharge
        
        Control.Visible = IsPrivs(mstrPrivs, "���ɸ���")
        Control.Enabled = blnAllowModify And mbytMode = 2 And Control.Visible
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MakeCharge * 2 + 1
        
        Control.Visible = (mstr������Դ = "����" And IsPrivs(mstrPrivs, "���ɸ���"))
        Control.Enabled = blnAllowModify And mbytMode = 2 And Control.Visible
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MakeCharge * 2 + 2, conMenu_Edit_MakeCharge * 2 + 3
        
        Control.Visible = IsPrivs(mstrPrivs, "���ɸ���")
        Control.Enabled = blnAllowModify And mbytMode = 2 And Control.Visible
        
    End Select

errHand:

End Sub

'######################################################################################################################
'���������弰��ؼ����¼�����

Private Sub Form_Resize()

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

        Case .ColIndex("��ҩ����")
                    
            If .ComboIndex > -1 Then
                .TextMatrix(Row, Col) = zlCommFun.GetNeedName(.ComboItem(.ComboIndex))
            End If
                
    
        Case .ColIndex("׼������"), .ColIndex("ʵ������")
            
            Call CheckStorage(Row)
            
        Case .ColIndex("ִ�п���")
        
            .TextMatrix(Row, .ColIndex("ִ�п���id")) = .ComboData
            .TextMatrix(Row, .ColIndex("ִ�п���")) = .Cell(flexcpTextDisplay, Row, .ColIndex("ִ�п���"))
            
            Call CheckStorage(Row)
                        
        Case .ColIndex("��������")
        
            .TextMatrix(Row, .ColIndex("��������id")) = .ComboData
            .TextMatrix(Row, .ColIndex("��������")) = .Cell(flexcpTextDisplay, Row, .ColIndex("��������"))
            
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
    Dim strCode As String
    Dim bln����� As Boolean
    
    With vsf(0)
        If Col = .ColIndex("ҩƷ����") Then
            
            strCode = zlCommFun.GetNeedName(.TextMatrix(Row, .ColIndex("��ҩ����")))
            
            gstrSQL = "Select �Ƿ������ From ������ҩ���� Where ����=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strCode)
            If rs.BOF = False Then bln����� = (zlCommFun.NVL(rs("�Ƿ������").Value, 0) = 1)
                
            If bln����� Then
                gstrSQL = GetPublicSQL(SQL.����ҩƷѡ��)
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)
                bytRet = ShowPubSelect(Me, vsf(0), 3, "����,1200,0,;����,2700,0,;���,900,0,;��λ,900,0,", Me.Name & "\����ҩƷѡ��", "����±���ѡ��һ��������Ŀ", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
            Else
                gstrSQL = GetPublicSQL(SQL.ҩƷ��Ŀѡ��)
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)
                bytRet = ShowPubSelect(Me, vsf(0), 3, "����,1200,0,;����,2700,0,;���,900,0,;��λ,900,0,", Me.Name & "\ҩƷ��Ŀѡ��", "����±���ѡ��һ��ҩƷ��Ŀ", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
            End If
            
            If bytRet = 1 Then
                If mclsVsf.CheckHave(zlCommFun.NVL(rs("ҩƷID").Value)) Then
                    ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                    Exit Sub
                End If
    
                .EditText = zlCommFun.NVL(rs("����").Value)
                .TextMatrix(Row, mclsVsf.ColIndex("ҩƷ����")) = zlCommFun.NVL(rs("����").Value)
                .TextMatrix(Row, mclsVsf.ColIndex("���")) = zlCommFun.NVL(rs("���").Value)
                .TextMatrix(Row, mclsVsf.ColIndex("��λ")) = zlCommFun.NVL(rs("��λ").Value)
                .TextMatrix(Row, mclsVsf.ColIndex("���")) = zlCommFun.NVL(rs("���").Value)
                .RowData(Row) = zlCommFun.NVL(rs("ҩƷID").Value, 0)
    
                Call ExecuteCommand("�շ�ִ�п���", zlCommFun.NVL(rs("���").Value))
                Call CheckStorage(Row)
                
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
        Case .ColIndex("׼������"), .ColIndex("ʵ������")
            .TextMatrix(.Row, .Col) = .EditText
        Case .ColIndex("��ҩ����")
            If .EditText <> .TextMatrix(.Row, .Col) Then
                .Cell(flexcpText, .Row, 0, .Row, .Cols - 1) = ""
                .RowData(.Row) = 0
            End If
        End Select

    End With
End Sub

Private Sub vsf_ComboDropDown(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim intLoop As Integer
    
    With vsf(Index)
        Select Case Col
        Case .ColIndex("��ҩ����")
            If .TextMatrix(Row, Col) <> "" Then
                For intLoop = 0 To .ComboCount - 1
                    If zlCommFun.GetNeedName(.ComboItem(intLoop)) = .TextMatrix(Row, Col) Then
                        .ComboIndex = intLoop
                        Exit For
                    End If
                Next
            End If
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
    Dim strCode As String
    Dim bln����� As Boolean
    
    With vsf(0)
        If KeyCode = vbKeyReturn Then
            
            If Col = .ColIndex("ҩƷ����") Then
                
                If InStr(vsf(0).EditText, "'") > 0 Then
                    KeyCode = 0
                    vsf(0).EditText = ""
                    Exit Sub
                End If

                strCode = zlCommFun.GetNeedName(.TextMatrix(Row, .ColIndex("��ҩ����")))
                
                gstrSQL = "Select �Ƿ������ From ������ҩ���� Where ����=[1]"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strCode)
                If rs.BOF = False Then bln����� = (zlCommFun.NVL(rs("�Ƿ������").Value, 0) = 1)
            
                strText = UCase(vsf(0).EditText)
                bytMode = GetApplyMode(strText)
                strTmp = IIf(ParamInfo.��Ŀ����ƥ�䷽ʽ = 1, "", "%") & strText & "%"
                
                If bln����� Then
                    gstrSQL = GetPublicSQL(SQL.����ҩƷ����, bytMode)
                Else
                    gstrSQL = GetPublicSQL(SQL.ҩƷ��Ŀ����, bytMode)
                End If

                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp)

                If ShowPubSelect(Me, vsf(0), 2, "����,1200,0,;����,2700,0,;���,900,0,;��λ,900,0,", Me.Name & "\ҩƷ��Ŀ����", "����±���ѡ��һ��ҩƷ��Ŀ", rsData, rs, 8790, 4500, , Val(.RowData(Row))) = 1 Then

                    If mclsVsf.CheckHave(zlCommFun.NVL(rs("ҩƷID").Value)) Then
                        ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                        Exit Sub
                    End If

                    .EditText = zlCommFun.NVL(rs("����").Value)
                    .TextMatrix(Row, .ColIndex("ҩƷ����")) = zlCommFun.NVL(rs("����").Value)
                    .TextMatrix(Row, .ColIndex("��λ")) = zlCommFun.NVL(rs("��λ").Value)
                    .TextMatrix(Row, .ColIndex("���")) = zlCommFun.NVL(rs("���").Value)
                    .TextMatrix(Row, .ColIndex("���")) = zlCommFun.NVL(rs("���").Value)
                    
                    .RowData(Row) = zlCommFun.NVL(rs("ҩƷID").Value, 0)

                    Call ExecuteCommand("�շ�ִ�п���", zlCommFun.NVL(rs("���").Value))
                    Call CheckStorage(Row)
                
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




