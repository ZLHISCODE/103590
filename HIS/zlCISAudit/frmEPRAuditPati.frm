VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Begin VB.Form frmEPRAuditPati 
   BorderStyle     =   0  'None
   Caption         =   "���˲�����д���"
   ClientHeight    =   6930
   ClientLeft      =   -60
   ClientTop       =   15
   ClientWidth     =   10455
   Icon            =   "frmEPRAuditPati.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3135
      Index           =   0
      Left            =   2745
      ScaleHeight     =   3135
      ScaleWidth      =   5580
      TabIndex        =   4
      Top             =   225
      Width           =   5580
      Begin VSFlex8Ctl.VSFlexGrid vfgPati 
         Height          =   2640
         Left            =   300
         TabIndex        =   5
         Top             =   270
         Width           =   4320
         _cx             =   7620
         _cy             =   4657
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
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
         WallPaperAlignment=   8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox picPane 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   3
      Left            =   15
      ScaleHeight     =   240
      ScaleWidth      =   1875
      TabIndex        =   2
      Top             =   5175
      Width           =   1905
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   -45
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   -30
         Width           =   1320
      End
   End
   Begin VB.PictureBox picPane 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   2
      Left            =   285
      ScaleHeight     =   240
      ScaleWidth      =   1320
      TabIndex        =   0
      Top             =   1830
      Width           =   1350
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   75
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   -30
         Width           =   930
      End
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmEPRAuditPati.frx":6852
      Left            =   525
      Top             =   75
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmEPRAuditPati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

'����
'----------------------------------------------------------------------------------------------------------------------
Private Enum mCol
    ��־ = 0: �¼�Ե��: Ӧд����: ����: ����ʱ��: Ҫ��ʱ��: ���ʱ��: ��ɼ�¼id: ��ǰʱ��: ��ע˵��
End Enum

Private Const conPane_Pati = 1
Private Const conPane_Audit = 2
Private Const conPane_Word = 3

'����
'----------------------------------------------------------------------------------------------------------------------
Private mlngDeptId As Long      '����id
Private mintKind As Integer     '��������
Private mstrDateFrom As String  '��ʼ����
Private mstrDateTo As String    '��������
Private mstrEvent As String     '�����¼���Χ
Private WithEvents mclsDockAduit As zlRichEPR.clsDockAduits
Attribute mclsDockAduit.VB_VarHelpID = -1
Private mfrmMain As Object
Private mblnReading As Boolean
Private mclsPati As clsVsf

'######################################################################################################################

Public Function zlInitData(ByVal frmMain As Object) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Set mfrmMain = frmMain
    
    If ExecuteCommand("��ʼ�ؼ�") = False Or ExecuteCommand("��ʼ����") = False Then Exit Function
    
End Function

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
        
    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview

        Call RptPrint(2)
    
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Print

        Call RptPrint(1)
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Excel

        Call RptPrint(3)
        
    End Select
    
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    With vfgPati
        Select Case Control.ID
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel               'Ԥ��,��ӡ,�����Excel
            If ActiveControl Is Nothing Then Exit Sub
            If Me.ActiveControl.Name = .Name Then
                Control.Enabled = (.Rows > .FixedRows)
            Else
                Control.Enabled = (.Rows > .FixedRows)
            End If
        
        End Select
        
    End With
    
End Sub

Public Function zlRefreshData(ByVal intKind As Integer, ByVal strDateFrom As String, ByVal strDateTo As String, Optional blnShow As Boolean) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim strOldValue As String
    On Error GoTo errH
    
    mstrDateFrom = strDateFrom
    mstrDateTo = strDateTo
    mintKind = intKind
    
    Call ExecuteCommand("��ʼ����")

    Select Case mintKind
    '------------------------------------------------------------------------------------------------------------------
    Case 1  '���ﲡ��
        
        strSQL = "Select D.ID, D.����, D.���� From ���ű� D, ��������˵�� M Where D.ID = M.����id And M.�������� = '�ٴ�' And M.������� In (1, 3) And ( TO_CHAR (D.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' or D.����ʱ�� is null) Order By D.����"
        
    '------------------------------------------------------------------------------------------------------------------
    Case 2  'סԺ����
        
        strSQL = "Select D.ID, D.����, D.���� From ���ű� D, ��������˵�� M Where D.ID = M.����id And M.�������� = '�ٴ�' And M.������� In (2, 3) And ( TO_CHAR (D.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' or D.����ʱ�� is null) Order By D.����"
        
    '------------------------------------------------------------------------------------------------------------------
    Case 4  '������
    
        strSQL = "Select D.ID, D.����, D.���� From ���ű� D, ��������˵�� M Where D.ID = M.����id And M.�������� = '�ٴ�' And M.������� In (2, 3) And ( TO_CHAR (D.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' or D.����ʱ�� is null) Order By D.����"

    End Select
    
    strOldValue = cboDept.Text
    cboDept.Clear
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            Call cboDept.AddItem(rs("����").Value)
            cboDept.ItemData(cboDept.NewIndex) = rs("ID").Value
            rs.MoveNext
        Loop
    End If
    
    mblnReading = True
    If Len(strOldValue) > 0 Then
        cboDept.Text = strOldValue
        mlngDeptId = cboDept.ItemData(cboDept.ListIndex)
    Else
        cboDept.ListIndex = 0
        mlngDeptId = cboDept.ItemData(cboDept.ListIndex)
    End If
    mblnReading = False
    
    If blnShow = False Then
        zlRefreshData = RefreshData(mintKind, mstrEvent, mlngDeptId, mstrDateFrom, mstrDateTo)
    End If
    Exit Function
errH:
    If Err.Number = 383 Then
        Err.Clear
        cboDept.ListIndex = 0
        Resume Next
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Function

'######################################################################################################################

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
    Dim objExtendedBar As CommandBar

    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    Call CommandBarInit(cbsThis)
    Set cbsThis.Icons = frmPubResource.imgApp.Icons
    cbsThis.Options.LargeIcons = False
    
    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ

    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsThis.ActiveMenuBar.Visible = False
    '------------------------------------------------------------------------------------------------------------------
            
            
    '���Ź�����
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsThis.Add("��׼", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    Set objControl = NewToolBar(objBar, xtpControlLabel, 7, "���ң�", , , xtpButtonIconAndCaption)
    Set cbrCustom = NewToolBar(objBar, xtpControlCustom, conMenu_Edit_NewItem, "")
    cbrCustom.Handle = picPane(3).hWnd
    
    Set objControl = NewToolBar(objBar, xtpControlLabel, 0, "���ͣ�", , , xtpButtonIconAndCaption)
    Set cbrCustom = NewToolBar(objBar, xtpControlCustom, conMenu_Edit_NewItem, "")
    cbrCustom.Handle = picPane(2).hWnd
    
    Set objControl = NewToolBar(objBar, xtpControlButton, 9, "���Ĳ���...", True, , xtpButtonIconAndCaption)
    
End Function


Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim intRow As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    Dim strNow As String
    Dim strNote As String
    
    On Error GoTo errHand
    
    mblnReading = True
    
    Call SQLRecord(rsSQL)

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ�ؼ�"
                
        Call InitGrid
        Call InitCommandBar
        
        '����ͣ������
        '--------------------------------------------------------------------------------------------------------------
        Dim objPane As Pane
        Set objPane = dkpMan.CreatePane(conPane_Pati, 300, 400, DockLeftOf, Nothing): objPane.Title = "�����б�": objPane.Options = PaneNoCaption
        Set objPane = dkpMan.CreatePane(conPane_Audit, 700, 100, DockRightOf, objPane): objPane.Title = "ʱ�����": objPane.Options = PaneNoCaption
        Set objPane = dkpMan.CreatePane(conPane_Word, 700, 300, DockBottomOf, objPane): objPane.Title = "���ݼ��": objPane.Options = PaneNoCaption
        
        dkpMan.SetCommandBars cbsThis
        Call DockPannelInit(dkpMan)
        
        
                            
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"
                        
        With cboType
            .Clear
            .AddItem "����"
            
            If mintKind = 1 Then
                .AddItem "����"
                .AddItem "����"
            Else
                .AddItem "��Ժ"
                .AddItem "ת��"
                .AddItem "��Ժ"
                .AddItem "����"
                .AddItem "ת��"
                .AddItem "����"
            End If
            
            .ListIndex = 0
            mstrEvent = "����"
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "ʱ�޼��"
        
        strSQL = "Zl_����ʱ�޼��_Neaten(" & Val(varParam(0)) & "," & Val(varParam(1)) & "," & mintKind & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, mfrmMain.Caption)
        
    End Select

    ExecuteCommand = True

    GoTo endHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
endHand:
    mblnReading = False
End Function

Private Function InitGrid(Optional ByVal strType As String = "2") As Boolean
    Set mclsPati = New clsVsf
    With mclsPati
        Call .Initialize(Me.Controls, vfgPati, False, False, frmPubResource.GetImageList(16))
        Call .ClearColumn
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[·��]", False)
        Call .AppendColumn("����id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
        If InStr(strType, "1") > 0 Then
            Call .AppendColumn("ID", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
        Else
            Call .AppendColumn("��ҳid", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
        End If
        If strType = "1" Then
            Call .AppendColumn("�����", 900, flexAlignLeftCenter, flexDTString, "", , True)
        Else
            Call .AppendColumn("סԺ��", 900, flexAlignLeftCenter, flexDTString, "", , True)
        End If
        Call .AppendColumn("����", 810, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("�Ա�", 990, flexAlignLeftCenter, flexDTString, "", , True)
        Select Case strType
            Case "2"
                Call .AppendColumn("��Ժʱ��", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "��Ժʱ��", True)
            Case "1"
                Call .AppendColumn("����ʱ��", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "����ʱ��", True)
            Case "21"
                Call .AppendColumn("ת��ʱ��", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "ת��ʱ��", True)
            Case "22"
                Call .AppendColumn("��Ժ����", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "��Ժ����", True)
            Case "23"
                Call .AppendColumn("��������", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "��������", True)
            Case "24"
                Call .AppendColumn("ת��ʱ��", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "ת��ʱ��", True)
            Case "25"
                Call .AppendColumn("����ʱ��", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "����ʱ��", True)
            Case Else
                Call .AppendColumn("��Ժʱ��", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "��Ժʱ��", True)
        End Select
        If strType = "1" Then
            Call .AppendColumn("ҽ��", 900, flexAlignLeftCenter, flexDTString, "", , True)
        End If
    End With
End Function
Private Function RefreshData(ByVal intKind As Integer, ByVal strEvent As String, ByVal lngDeptKey As Long, ByVal strDateFrom As String, ByVal strDateTo As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    Select Case intKind
    Case 1
        Call InitGrid("1")
        Select Case strEvent
        Case "����"
            strSQL = "Select Null as ·��, ����id, ID, �����, ����, �Ա�, To_Char(ִ��ʱ��, 'yyyy-mm-dd hh24:mi') As ����ʱ��, ִ���� As ҽ��" & vbNewLine & _
                    "From ���˹Һż�¼" & vbNewLine & _
                    "Where ִ�в���id + 0 = [1] And Nvl(ִ��״̬, 0) <> 0 And Nvl(����, 0) <> 1 And ��¼����=1 And ��¼״̬=1 And" & vbNewLine & _
                    "      �Ǽ�ʱ�� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By ִ��ʱ��"
        Case "����"
            strSQL = "Select Null as ·��, ����id, ID, �����, ����, �Ա�, To_Char(ִ��ʱ��, 'yyyy-mm-dd hh24:mi') As ����ʱ��, ִ���� As ҽ��" & vbNewLine & _
                    "From ���˹Һż�¼" & vbNewLine & _
                    "Where ִ�в���id + 0 = [1] And Nvl(ִ��״̬, 0) <> 0 And Nvl(����, 0) = 1 And ��¼����=1 And ��¼״̬=1 And" & vbNewLine & _
                    "      �Ǽ�ʱ�� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By ִ��ʱ��"
        Case Else
            strSQL = "Select Null as ·��,����id, ID, �����, ����, �Ա�, To_Char(ִ��ʱ��, 'yyyy-mm-dd hh24:mi') As ����ʱ��, ִ���� As ҽ��" & vbNewLine & _
                    "From ���˹Һż�¼" & vbNewLine & _
                    "Where ִ�в���id + 0 = [1] And Nvl(ִ��״̬, 0) <> 0 And ��¼����=1 And ��¼״̬=1 And" & vbNewLine & _
                    "      �Ǽ�ʱ�� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By ִ��ʱ��"
        End Select
    Case 2
        Select Case strEvent
        Case "��Ժ"
            Call InitGrid("2")
            strSQL = "Select Decode(P.·��״̬,Null,'',0,'','lujin') as ·��,P.����id, P.��ҳid, P.סԺ��, I.����, I.�Ա�, L.��Ժʱ��" & vbNewLine & _
                    "From ������Ϣ I, ������ҳ P," & vbNewLine & _
                    "     (Select ����id, ��ҳid, To_Char(Max(��ʼʱ��), 'yyyy-mm-dd hh24:mi') As ��Ժʱ��" & vbNewLine & _
                    "       From ���˱䶯��¼" & vbNewLine & _
                    "       Where ����id + 0 = [1] And ��ʼԭ�� In (1, 2, 9) And ��ʼʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "       Group By ����id, ��ҳid) L" & vbNewLine & _
                    "Where I.����id = P.����id And P.����id = L.����id And P.��ҳid = L.��ҳid" & vbNewLine & _
                    "Order By L.��Ժʱ��"
        Case "ת��"
            Call InitGrid("21")
            strSQL = "Select Decode(P.·��״̬,Null,'',0,'','lujin') as ·��,P.����id, P.��ҳid, P.סԺ��, I.����, I.�Ա�, L.ת��ʱ��" & vbNewLine & _
                    "From ������Ϣ I, ������ҳ P," & vbNewLine & _
                    "     (Select Distinct ����id, ��ҳid, To_Char(��ʼʱ��, 'yyyy-mm-dd hh24:mi') As ת��ʱ��" & vbNewLine & _
                    "       From ���˱䶯��¼" & vbNewLine & _
                    "       Where ����id + 0 = [1] And ��ʼԭ�� = 3 And ��ʼʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) L" & vbNewLine & _
                    "Where I.����id = P.����id And P.����id = L.����id And P.��ҳid = L.��ҳid" & vbNewLine & _
                    "Order By L.ת��ʱ��"
        Case "��Ժ"
            Call InitGrid("22")
            strSQL = "Select Decode(P.·��״̬,Null,'',0,'','lujin') as ·��,P.����id, P.��ҳid, P.סԺ��, I.����, I.�Ա�, To_Char(P.��Ժ����, 'yyyy-mm-dd hh24:mi') As ��Ժ����" & vbNewLine & _
                    "From ������Ϣ I, ������ҳ P" & vbNewLine & _
                    "Where I.����id = P.����id And P.��Ժ����id + 0 = [1] And P.��Ժ��ʽ <> '����' And" & vbNewLine & _
                    "      P.��Ժ���� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By P.��Ժ����"
        Case "����"
            Call InitGrid("23")
            strSQL = "Select Decode(P.·��״̬,Null,'',0,'','lujin') as ·��,P.����id, P.��ҳid, P.סԺ��, I.����, I.�Ա�, To_Char(P.��Ժ����, 'yyyy-mm-dd hh24:mi') As ��������" & vbNewLine & _
                    "From ������Ϣ I, ������ҳ P" & vbNewLine & _
                    "Where I.����id = P.����id And P.��Ժ����id + 0 = [1] And P.��Ժ��ʽ = '����' And" & vbNewLine & _
                    "      P.��Ժ���� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By P.��Ժ����"
        Case "ת��"
            Call InitGrid("24")
            strSQL = "Select Decode(P.·��״̬,Null,'',0,'','lujin') as ·��,P.����id, P.��ҳid, P.סԺ��, I.����, I.�Ա�, L.ת��ʱ��" & vbNewLine & _
                    "From ������Ϣ I, ������ҳ P," & vbNewLine & _
                    "     (Select Distinct ����id, ��ҳid, To_Char(��ֹʱ��, 'yyyy-mm-dd hh24:mi') As ת��ʱ��" & vbNewLine & _
                    "       From ���˱䶯��¼" & vbNewLine & _
                    "       Where ����id + 0 = [1] And ��ֹԭ�� = 3 And ��ֹʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) L" & vbNewLine & _
                    "Where I.����id = P.����id And P.����id = L.����id And P.��ҳid = L.��ҳid" & vbNewLine & _
                    "Order By L.ת��ʱ��"
        Case "����"
            Call InitGrid("25")
            strSQL = "Select Decode(P.·��״̬,Null,'',0,'','lujin') as ·��,P.����id, P.��ҳid, P.סԺ��, I.����, I.�Ա�, ����ʱ��" & vbNewLine & _
                    "From ������Ϣ I, ������ҳ P," & vbNewLine & _
                    "     (Select R.����id, R.��ҳid, To_Char(S.�״�ʱ��, 'yyyy-mm-dd hh24:mi') As ����ʱ��" & vbNewLine & _
                    "       From ����ҽ����¼ R, ����ҽ������ S" & vbNewLine & _
                    "       Where R.ID = S.ҽ��id And R.������� = 'F' And R.���id Is Null And R.ҽ����Ч = 1 And" & vbNewLine & _
                    "             (R.ҽ��״̬ = 8 Or R.ҽ��״̬ = 9) And R.���˿���id + 0 = [1] And" & vbNewLine & _
                    "             S.�״�ʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) L" & vbNewLine & _
                    "Where I.����id = P.����id And P.����id = L.����id And P.��ҳid = L.��ҳid" & vbNewLine & _
                    "Order By L.����ʱ��"
        Case Else
            Call InitGrid
            strSQL = "Select Decode(P.·��״̬,Null,'',0,'','lujin') as ·��,P.����id, P.��ҳid, P.סԺ��, I.����, I.�Ա�, P.��Ժ����" & vbNewLine & _
                    "From ������Ϣ I, ������ҳ P," & vbNewLine & _
                    "     (Select Distinct ����id, ��ҳid" & vbNewLine & _
                    "       From ���˱䶯��¼" & vbNewLine & _
                    "       Where ����id = [1] And" & vbNewLine & _
                    "             (��ʼԭ�� In (1, 2, 3, 9) And ��ʼʱ�� between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400 Or" & vbNewLine & _
                    "             ��ֹԭ�� In (1, 3, 10) And (��ֹʱ�� between To_Date([2], 'yyyy-mm-dd') and To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400 Or ��ֹʱ�� Is Null))) L" & vbNewLine & _
                    "Where I.����id = P.����id And P.����id = L.����id And P.��ҳid = L.��ҳid" & vbNewLine & _
                    "Order By P.��Ժ����"
                    
        End Select
    Case 4
        Select Case strEvent
        Case "��Ժ"
            Call InitGrid("2")
            strSQL = "Select Decode(P.·��״̬,Null,'',0,'','lujin') as ·��,P.����id, P.��ҳid, P.סԺ��, I.����, I.�Ա�, L.��Ժʱ��" & vbNewLine & _
                    "From ������Ϣ I, ������ҳ P," & vbNewLine & _
                    "     (Select ����id, ��ҳid, To_Char(Max(��ʼʱ��), 'yyyy-mm-dd hh24:mi') As ��Ժʱ��" & vbNewLine & _
                    "       From ���˱䶯��¼" & vbNewLine & _
                    "       Where ����id + 0 = [1] And ��ʼԭ�� In (1, 2, 9) And ��ʼʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "       Group By ����id, ��ҳid) L" & vbNewLine & _
                    "Where I.����id = P.����id And P.����id = L.����id And P.��ҳid = L.��ҳid" & vbNewLine & _
                    "Order By L.��Ժʱ��"
        Case "ת��"
            Call InitGrid("21")
            strSQL = "Select Decode(P.·��״̬,Null,'',0,'','lujin') as ·��,P.����id, P.��ҳid, P.סԺ��, I.����, I.�Ա�, L.ת��ʱ��" & vbNewLine & _
                    "From ������Ϣ I, ������ҳ P," & vbNewLine & _
                    "     (Select Distinct ����id, ��ҳid, To_Char(��ʼʱ��, 'yyyy-mm-dd hh24:mi') As ת��ʱ��" & vbNewLine & _
                    "       From ���˱䶯��¼" & vbNewLine & _
                    "       Where ����id + 0 = [1] And ��ʼԭ�� = 3 And ��ʼʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) L" & vbNewLine & _
                    "Where I.����id = P.����id And P.����id = L.����id And P.��ҳid = L.��ҳid" & vbNewLine & _
                    "Order By L.ת��ʱ��"
        Case "��Ժ"
            Call InitGrid("22")
            strSQL = "Select Decode(P.·��״̬,Null,'',0,'','lujin') as ·��,P.����id, P.��ҳid, P.סԺ��, I.����, I.�Ա�, To_Char(P.��Ժ����, 'yyyy-mm-dd hh24:mi') As ��Ժ����" & vbNewLine & _
                    "From ������Ϣ I, ������ҳ P" & vbNewLine & _
                    "Where I.����id = P.����id And P.��ǰ����id + 0 = [1] And P.��Ժ��ʽ <> '����' And" & vbNewLine & _
                    "      P.��Ժ���� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By P.��Ժ����"
        Case "����"
            Call InitGrid("23")
            strSQL = "Select Decode(P.·��״̬,Null,'',0,'','lujin') as ·��,P.����id, P.��ҳid, P.סԺ��, I.����, I.�Ա�, To_Char(P.��Ժ����, 'yyyy-mm-dd hh24:mi') As ��������" & vbNewLine & _
                    "From ������Ϣ I, ������ҳ P" & vbNewLine & _
                    "Where I.����id = P.����id And P.��ǰ����id + 0 = [1] And P.��Ժ��ʽ = '����' And" & vbNewLine & _
                    "      P.��Ժ���� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By P.��Ժ����"
        Case "ת��"
            Call InitGrid("24")
            strSQL = "Select Decode(P.·��״̬,Null,'',0,'','lujin') as ·��,P.����id, P.��ҳid, P.סԺ��, I.����, I.�Ա�, L.ת��ʱ��" & vbNewLine & _
                    "From ������Ϣ I, ������ҳ P," & vbNewLine & _
                    "     (Select Distinct ����id, ��ҳid, To_Char(��ֹʱ��, 'yyyy-mm-dd hh24:mi') As ת��ʱ��" & vbNewLine & _
                    "       From ���˱䶯��¼" & vbNewLine & _
                    "       Where ����id + 0 = [1] And ��ֹԭ�� = 3 And ��ֹʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) L" & vbNewLine & _
                    "Where I.����id = P.����id And P.����id = L.����id And P.��ҳid = L.��ҳid" & vbNewLine & _
                    "Order By L.ת��ʱ��"
        Case Else
            Call InitGrid
            strSQL = "Select Decode(P.·��״̬,Null,'',0,'','lujin') as ·��,P.����id, P.��ҳid, P.סԺ��, I.����, I.�Ա�, P.��Ժ����" & vbNewLine & _
                    "From ������Ϣ I, ������ҳ P," & vbNewLine & _
                    "     (Select Distinct ����id, ��ҳid" & vbNewLine & _
                    "       From ���˱䶯��¼" & vbNewLine & _
                    "       Where ����id = [1] And" & vbNewLine & _
                    "             (��ʼԭ�� In (1, 2, 3, 9) And ��ʼʱ�� between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400 Or" & vbNewLine & _
                    "             ��ֹԭ�� In (1, 3, 10) And (��ֹʱ�� between To_Date([2], 'yyyy-mm-dd') and To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400 Or ��ֹʱ�� Is Null))) L" & vbNewLine & _
                    "Where I.����id = P.����id And P.����id = L.����id And P.��ҳid = L.��ҳid" & vbNewLine & _
                    "Order By P.��Ժ����"
        End Select
    End Select
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngDeptKey, strDateFrom, strDateTo)
    
    Dim lngCount As Long
    mclsPati.ClearGrid
    Call mclsDockAduit.zlClearTime
    With vfgPati
        If rs.RecordCount > 0 Then
            Call mclsPati.LoadDataSource(rs)
            For lngCount = 0 To .Cols - 1
                .FixedAlignment(lngCount) = flexAlignCenterCenter
            Next
            .ColWidth(1) = 0
            .ColHidden(1) = True
            .Row = 0
        End If
        
    End With
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub RptPrint(ByVal bytMode As Byte)
    '******************************************************************************************************************
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '******************************************************************************************************************
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    Dim strDeptName As String
    
    If cboDept.ListIndex < 0 Then Exit Sub
    
    strDeptName = cboDept.List(cboDept.ListIndex)
    
    If Me.ActiveControl Is Nothing Then
        Set objPrint.Body = vfgPati
        objPrint.Title.Text = strDeptName & mstrEvent & "�����嵥"
    Else
        If Me.ActiveControl.Name = vfgPati.Name Then
            Set objPrint.Body = vfgPati
            objPrint.Title.Text = strDeptName & mstrEvent & "�����嵥"
        Else
                        
            Set objPrint.Body = mclsDockAduit.zlGetFormAuditTimeLimit.vfgAudit
            objPrint.Title.Text = "���˲���ʱ�ޱ���"
            Set objAppRow = New zlTabAppRow
            Call objAppRow.Add(Me.vfgPati.TextMatrix(Me.vfgPati.FixedRows - 1, 2) & ":" & Me.vfgPati.TextMatrix(Me.vfgPati.Row, 2))
            Call objAppRow.Add("����:" & Me.vfgPati.TextMatrix(Me.vfgPati.Row, 3))
            Call objPrint.UnderAppRows.Add(objAppRow)
        End If
    End If
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    Me.Tag = "Printing"
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Me.Tag = ""
End Sub


Private Sub cboDept_Click()
    If mblnReading Then Exit Sub
    If cboDept.ListIndex < 0 Then Exit Sub
    
    If mlngDeptId <> cboDept.ItemData(cboDept.ListIndex) Then
        mlngDeptId = cboDept.ItemData(cboDept.ListIndex)
        Call RefreshData(mintKind, mstrEvent, mlngDeptId, mstrDateFrom, mstrDateTo)
    End If
End Sub

Private Sub cboType_Click()
    If mblnReading Then Exit Sub
    If cboType.ListIndex < 0 Then Exit Sub
    
    If mstrEvent <> cboType.List(cboType.ListIndex) Then
        mstrEvent = cboType.List(cboType.ListIndex)
        Call RefreshData(mintKind, mstrEvent, mlngDeptId, mstrDateFrom, mstrDateTo)
    End If
    
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRecordId As Long
    
    Select Case Control.ID
    Case 9
        
        lngRecordId = mclsDockAduit.zlGetFormAuditTimeLimit.GetCurrentEPRKey

        If lngRecordId > 0 Then
            If mclsDockAduit Is Nothing Then Set mclsDockAduit = New zlRichEPR.clsDockAduits
            Call mclsDockAduit.zlOpenEPRDocument(lngRecordId, mfrmMain)
        End If
        
    End Select
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.ID
    Case 9
        Control.Enabled = (mclsDockAduit.zlGetFormAuditTimeLimit.GetCurrentEPRKey > 0)
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_Pati
        Item.Handle = picPane(0).hWnd
    Case conPane_Audit
        
        If mclsDockAduit Is Nothing Then Set mclsDockAduit = New zlRichEPR.clsDockAduits
        Call mclsDockAduit.zlInitTime(Me)
        
        Item.Handle = mclsDockAduit.zlGetFormAuditTimeLimit.hWnd
    Case conPane_Word
        If mclsDockAduit Is Nothing Then Set mclsDockAduit = New zlRichEPR.clsDockAduits
        Call mclsDockAduit.zlInitMonitor(Me)
    
        Item.Handle = mclsDockAduit.zlGetFormAuditMonitor.hWnd

    End Select
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call SetPaneRange(dkpMan, 1, 100, 15, 350, Me.ScaleHeight)
    
    dkpMan.RecalcLayout
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If Not (mclsDockAduit Is Nothing) Then Set mclsDockAduit = Nothing
        
End Sub

Private Sub mclsDockAduit_AfterDocumentChanged(ByVal lngEPRKey As Long)
    Call mclsDockAduit.zlRefreshMonitor(lngEPRKey)
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        vfgPati.Move 0, 0, picPane(Index).Width, picPane(Index).Height

        cboDept.Move -30, -30, picPane(3).Width + 45
        cboType.Move -30, -30, picPane(2).Width + 45
    End Select
End Sub

Private Sub vfgPati_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    With vfgPati
        If OldRow <> NewRow And NewRow > 0 Then
                                    
            Call mclsDockAduit.zlRefreshTime(Val(.TextMatrix(NewRow, 1)), Val(.TextMatrix(NewRow, 2)), mintKind)
                       
        End If
    End With
    
End Sub