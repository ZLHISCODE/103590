VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{099B2A6C-9CCE-43CF-AEF0-C526C98F4B7F}#2.1#0"; "zlRichEditor.ocx"
Begin VB.Form frmChildStationInEPR 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   10305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2670
      Index           =   1
      Left            =   615
      ScaleHeight     =   2670
      ScaleWidth      =   6090
      TabIndex        =   2
      Top             =   3585
      Width           =   6090
      Begin zlRichEditor.Editor edt 
         Height          =   1245
         Left            =   3675
         TabIndex        =   3
         Top             =   765
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   2196
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   1785
      Index           =   0
      Left            =   945
      ScaleHeight     =   1785
      ScaleWidth      =   6090
      TabIndex        =   0
      Top             =   -75
      Width           =   6090
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1215
         Index           =   0
         Left            =   555
         TabIndex        =   1
         Top             =   270
         Width           =   1860
         _cx             =   3281
         _cy             =   2143
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483626
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmChildStationInEPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################
Private mlngKey As Long
Private mlngReferKey As Long
Private mfrmMain As Object
Private mbytMode As Byte
Private mblnAllowModify As Boolean
Private mblnDataChanged As Boolean
Private mblnReading As Boolean
Private mlng����id As Long
Private mlng��ҳid As Long
Private mlngҽ��id As Long
Private mlng����ID As Long
    
Private WithEvents mobjDoc As zlRichEPR.cEPRDocument
Attribute mobjDoc.VB_VarHelpID = -1
Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1

Public Event AfterDataChanged()

'######################################################################################################################

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

    If ExecuteCommand("��ʼ����") = False Then Exit Function
    Call ExecuteCommand("�ؼ�״̬")
    
    DataChanged = False
    
    InitData = True
    
End Function

Public Sub zlDefCommandBars(ByVal cbsMain As CommandBars)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    
    'ҽ���˵�:���ڹ���˵�(���������û��)���ļ��˵�����
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "����(&E)", objMenu.Index + 1, False)
    objMenu.ID = conMenu_EditPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Edit_NewItem, "����(&N)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")

    '����������:���ļ�������˵������ť֮��ʼ����
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain(2)
    
    For Each objControl In objBar.Controls  '�����ǰ������һ��Control
        If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
            Set objControl = objBar.Controls(objControl.Index - 1): Exit For
        End If
    Next
        
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Save, "����", True, , , , objControl.Index + 1)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ��", , , , , objControl.Index + 1)
    
    '����Ŀ����:���������������Ѵ���
    '------------------------------------------------------------------------------------------------------------------
'    With cbsMain.KeyBindings
'        .Add 0, vbKeyF2, conMenu_Edit_NewItem              '����
'    End With

End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim rs As New ADODB.Recordset
    
    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem * 2
        
        If Val(Split(Control.Parameter, ";")(1)) = -1 Then
        
            If Control.Caption = "�����¼" Then
            
                '�ر�Ĵ���ʽ
'                Call mobjDoc.ShowCaseNarcosis(mlng����id, mlng��ҳid, 0, mlng����ID, 1, mfrmMain, True)
                
            End If
            
        ElseIf Val(Split(Control.Parameter, ";")(0)) > 0 Then
        
            Call mobjDoc.InitEPRDoc(cprEM_����, cprET_�������༭, Val(Split(Control.Parameter, ";")(0)), cprPF_סԺ, mlng����id, mlng��ҳid, 0, mlng����ID, mlngҽ��id)
            mobjDoc.ShowEPREditor mfrmMain
            
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Modify
    
        With vsf(0)
            
            If Val(.TextMatrix(.Row, .ColIndex("����"))) = 0 Then Exit Sub
            
            gstrSQL = "Select a.����,a.���� From �����ļ��б� a,���Ӳ�����¼ b Where b.ID=[1] And b.�ļ�id=a.ID"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, Val(.RowData(.Row)))
            If rs("����").Value = "�����¼" And zlCommFun.NVL(rs("����").Value, 0) = -1 Then

                '�ر�Ĵ���ʽ
'                Call mobjDoc.ShowCaseNarcosis(mlng����id, mlng��ҳid, Val(.RowData(.Row)), mlng����ID, 2, mfrmMain, True)
                
            Else
            
                If Val(.RowData(.Row)) > 0 Then
                    Call mobjDoc.InitEPRDoc(cprEM_�޸�, cprET_�������༭, Val(.RowData(.Row)), cprPF_סԺ, mlng����id, mlng��ҳid, 0, mlng����ID, mlngҽ��id)
                    mobjDoc.ShowEPREditor mfrmMain
                End If
            End If
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete
        
        If ExecuteCommand("ɾ������") Then
            Call ExecuteCommand("��ȡ����")
        End If
        
    End Select
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_NewItem
        Control.Enabled = mblnAllowModify And mlngKey > 0
    Case conMenu_Edit_Modify, conMenu_Edit_Delete
        With vsf(0)
            Control.Enabled = mblnAllowModify And mlngKey > 0 And Val(.RowData(.Row)) > 0 And Val(.TextMatrix(.Row, .ColIndex("����"))) = 1
        End With
    End Select
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As CommandBar)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim objControl As CommandBarControl

    If CommandBar.Parent Is Nothing Then Exit Sub
    
    Select Case CommandBar.Parent.ID
    Case conMenu_Edit_NewItem
        With CommandBar.Controls
            .DeleteAll
            
            strSQL = "Select a.ID,a.����,a.���� From �����ļ��б� a,����ʱ��Ҫ�� b Where a.Id=b.�ļ�id And b.�¼�='����' And ����=[1]"
            Set rs = zlDatabase.OpenSQLRecord(strSQL, mfrmMain.Caption, 2)
            If rs.BOF = False Then
                Do While Not rs.EOF
                    Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem * 2, rs("����").Value)
                    objControl.Parameter = rs("ID").Value & ";" & zlCommFun.NVL(rs("����").Value, 0)
                    rs.MoveNext
                Loop
            End If
        End With
    End Select
        
End Sub

Public Function RefreshData(ByVal lngKey As Long, ByVal lng����id As Long, ByVal lng��ҳid As Long, ByVal lng����id As Long, ByVal lngҽ��id As Long, Optional ByVal blnAllowModify As Boolean = True) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mlngKey = lngKey
    mlng����id = lng����id
    mlng��ҳid = lng��ҳid
    mlngҽ��id = lngҽ��id
    mlng����ID = lng����id
    
    mblnAllowModify = blnAllowModify
    
    Call ExecuteCommand("�������")
    Call ExecuteCommand("�ؼ�״̬")
    
    If mlng����id > 0 Then
        If ExecuteCommand("��ȡ����") = False Then Exit Function
        
    End If
    Call ExecuteCommand("��ʾ����")
    
    DataChanged = False
    
    RefreshData = True
    
End Function

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
'    Dim objPane As Pane
'
'    Set objPane = dkpMain.CreatePane(10, 100, 100, DockTopOf, Nothing)
'    objPane.Title = "�����б�"
'    objPane.Options = PaneNoCaption
'
'    Set objPane = dkpMain.CreatePane(11, 100, 100, DockBottomOf, objPane)
'    objPane.Title = "��������"
'    objPane.Options = PaneNoCaption
'
''    dkpMain.SetCommandBars cbsMain
'
'    dkpMain.Options.ThemedFloatingFrames = True
'    dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
'    dkpMain.Options.AlphaDockingContext = True
'    dkpMain.Options.HideClient = True
End Sub

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    Dim blnAllowModify As Boolean
    
    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"
        
        Set mclsVsf = New clsVsf
        With mclsVsf
            Call .Initialize(Me.Controls, vsf(0), True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
            Call .AppendColumn("��������", 2100, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("������", 810, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("����ʱ��", 1670, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm", , True)
            Call .AppendColumn("������", 810, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("��ǰ�汾", 900, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("������", 900, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("����", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("��ǰ���", 900, flexAlignLeftCenter, flexDTString, "", , True)
            .AppendRows = True
        End With
        
        Call InitDockPannel
        
        Set mobjDoc = New zlRichEPR.cEPRDocument
        
    '------------------------------------------------------------------------------------------------------------------
    Case "�ؼ�״̬"
    
        blnAllowModify = mblnAllowModify
        If mlngKey = 0 And mbytMode = 2 Then blnAllowModify = False

    '------------------------------------------------------------------------------------------------------------------
    Case "�������"
    
        mclsVsf.ClearGrid
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ����"
        
        mclsVsf.ClearGrid
        
        gstrSQL = "Select r.Id, r.��������, r.������ As ������,Decode(s.��¼id,Null,0,[3],1,0) As ����," & _
                "        r.����ʱ��, r.������," & _
                "        r.���ʱ��, r.���汾 As ��ǰ�汾," & _
                "        Decode(r.���汾, 1, '��д��', '�޶���') || r.������ || '��' || To_Char(r.����ʱ��, 'yyyy-mm-dd hh24:mi') ||" & _
                "         Decode(Nvl(r.ǩ������, 0), 0, '����(δ���)', 1, '���', '��ǩ') As ��ǰ���, r.�鵵��, r.�鵵����," & _
                "        d.���� As ������, f.����, r.����״̬, p.����״̬" & _
                " From ���Ӳ�����¼ r, ���ű� d,������������ s," & _
                "      (Select Decode(��Ժ����, Null, Decode(״̬, 3, 'Ԥ��Ժ', '��Ժ'), '��Ժ') As ����״̬" & vbNewLine & _
                "        From ������ҳ" & vbNewLine & _
                "        Where ����id = [1] And ��ҳid = [2]) p," & _
                "      (Select d.Id As �ļ�id, f.����, f.���, f.���� As ҳ��, d.����" & _
                "        From �����ļ��б� d, ����ҳ���ʽ f" & _
                "        Where d.���� In (2, 5, 6) And d.���� = f.���� And d.ҳ�� = f.���) f" & _
                " Where r.�ļ�id = f.�ļ�id(+) And r.������Դ = 2 And r.�������� In (2, 5, 6) And r.����id = d.Id And r.����id = [1] And r.��ҳid = [2] And s.�ļ�id(+)=r.ID" & _
                " Order By r.��������, f.���, r.���, r.Id"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlng����id, mlng��ҳid, mlngKey)
        If rs.BOF = False Then Call mclsVsf.LoadGrid(rs)

    '------------------------------------------------------------------------------------------------------------------
    Case "��ʾ����"
        
        With vsf(0)
            Call ShowDocument(edt, Val(.RowData(.Row)))
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "ɾ������"
    
        With vsf(0)
            
            If Val(.TextMatrix(.Row, .ColIndex("����"))) = 0 Then Exit Function
            
            If MsgBox("�Ƿ����Ҫɾ����" & .TextMatrix(.Row, .ColIndex("��������")) & "���������ݣ�", vbYesNo + vbDefaultButton2 + vbQuestion, ParamInfo.ϵͳ����) = vbNo Then Exit Function
            
                        
            gstrSQL = "zl_������������_Delete(" & mlngKey & "," & Val(.RowData(.Row)) & ")"
            Call SQLRecordAdd(rsSQL, gstrSQL)
            
            ExecuteCommand = SQLRecordExecute(rsSQL, mfrmMain.Caption)
            
            Exit Function
        End With
    End Select

    ExecuteCommand = True

    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

'Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
'    Select Case Item.ID
'    Case 10
'        Item.Handle = picPane(0).hWnd
'    Case 11
'        Item.Handle = picPane(1).hWnd
'    End Select
'End Sub

Private Sub Form_Resize()
    On Error Resume Next
'
'    Call SetPaneRange(dkpMain, 10, 100, 100, Me.ScaleWidth, 250)
'
'    dkpMain.RecalcLayout
    
    picPane(0).Move 0, 0, Me.ScaleWidth
    picPane(1).Move 0, picPane(0).Top + picPane(0).Height + 30, Me.ScaleWidth, Me.ScaleHeight - (picPane(0).Top + picPane(0).Height + 30)
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    Set mclsVsf = Nothing
    Set mobjDoc = Nothing

End Sub

Private Sub mobjDoc_AfterSaved(lngRecordId As Long)

    'ˢ�½���
    Call ExecuteCommand("��ȡ����")
    Call ExecuteCommand("��ʾ����")
End Sub

Private Sub mobjDoc_BeforeSaved(lngRecordId As Long, Cancel As Boolean)
    Dim rsSQL As New ADODB.Recordset
    Dim strSQL As String
    
    If mlngKey > 0 And lngRecordId > 0 Then
        
        Call SQLRecord(rsSQL)
        
        strSQL = "zl_������������_Update(" & mlngKey & "," & lngRecordId & ")"
        Call SQLRecordAdd(rsSQL, strSQL)
        
        Cancel = Not SQLRecordExecute(rsSQL, mfrmMain.Caption, False)
        
    End If
    
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next

    Select Case Index
    Case 0
        vsf(0).Move 0, 0, picPane(Index).Width, picPane(Index).Height
        mclsVsf.AppendRows = True
    Case 1
        edt.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
    End Select
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow <> OldRow Then Call ExecuteCommand("��ʾ����")
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
