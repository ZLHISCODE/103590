VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmEPRAuditFile 
   BorderStyle     =   0  'None
   Caption         =   "�����ļ��������"
   ClientHeight    =   6210
   ClientLeft      =   -60
   ClientTop       =   0
   ClientWidth     =   9630
   Icon            =   "frmEPRAuditFile.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2280
      Index           =   0
      Left            =   735
      ScaleHeight     =   2280
      ScaleWidth      =   4440
      TabIndex        =   3
      Top             =   960
      Width           =   4440
      Begin VSFlex8Ctl.VSFlexGrid vfgFile 
         Height          =   1200
         Left            =   735
         TabIndex        =   4
         Top             =   360
         Width           =   1845
         _cx             =   3254
         _cy             =   2117
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
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
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
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3135
      Index           =   1
      Left            =   945
      ScaleHeight     =   3135
      ScaleWidth      =   5625
      TabIndex        =   2
      Top             =   3435
      Width           =   5625
      Begin VSFlex8Ctl.VSFlexGrid vfgEPRs 
         Height          =   1200
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1845
         _cx             =   3254
         _cy             =   2117
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
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
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
         WallPaperAlignment=   9
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
      Left            =   5655
      ScaleHeight     =   240
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   540
      Width           =   1905
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   -45
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   -30
         Width           =   1320
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
      Bindings        =   "frmEPRAuditFile.frx":058A
      Left            =   525
      Top             =   150
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmEPRAuditFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

'����
'----------------------------------------------------------------------------------------------------------------------
Private Const conPane_File = 1
Private Const conPane_EPRs = 2
Private Const conPane_Word = 3


'����
'----------------------------------------------------------------------------------------------------------------------
Private mlngDeptId As Long      '����id
Private mstrDeptName As String  '������
Private mintKind As Integer     '��������
Private mstrDateFrom As String  '��ʼ����
Private mstrDateTo As String    '��������
Private mclsDockAduit As zlRichEPR.clsDockAduits
Private mfrmMain As Object
Private mblnReading As Boolean
Private mclsEPRs As clsVsf
Private mclsFile As clsVsf

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
    With vfgFile
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
    '������blnShow ����ֻˢ��cboDept
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim strOldValue As String
    On Error GoTo ErrH
    
    mstrDateFrom = strDateFrom
    mstrDateTo = strDateTo
    mintKind = intKind
    
    Select Case mintKind
    '------------------------------------------------------------------------------------------------------------------
    Case 1  '���ﲡ��
        
        strSQL = "Select distinct D.ID, D.����, D.���� From ���ű� D, ��������˵�� M Where D.ID = M.����id And M.�������� in ('�ٴ�','����','����') And M.������� In (1, 3) And ( TO_CHAR (D.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' or D.����ʱ�� is null) Order By D.����"
        
    '------------------------------------------------------------------------------------------------------------------
    Case 2  'סԺ����
        
        strSQL = "Select distinct D.ID, D.����, D.���� From ���ű� D, ��������˵�� M Where D.ID = M.����id And M.�������� in ('�ٴ�','����','����') And M.������� In (2, 3) And ( TO_CHAR (D.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' or D.����ʱ�� is null) Order By D.����"
        
    '------------------------------------------------------------------------------------------------------------------
    Case 4  '������
    
        strSQL = "Select distinct D.ID, D.����, D.���� From ���ű� D, ��������˵�� M Where D.ID = M.����id And M.�������� in ('�ٴ�','����','����') And M.������� In (2, 3) And ( TO_CHAR (D.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' or D.����ʱ�� is null) Order By D.����"

    End Select
    
    '------------------------------------------------------------------------------------------------------------------
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
    
    '------------------------------------------------------------------------------------------------------------------
    
    If blnShow = False Then
        zlRefreshData = RefreshData(mintKind, mlngDeptId, mstrDateFrom, mstrDateTo)
    End If
    Exit Function
ErrH:
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
    
    
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ
    '------------------------------------------------------------------------------------------------------------------
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsThis.ActiveMenuBar.Visible = False
                
            
    '���Ź�����
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsThis.Add("��׼", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    Set objControl = NewToolBar(objBar, xtpControlLabel, 7, "���ң�", , , xtpButtonIconAndCaption)
    Set cbrCustom = NewToolBar(objBar, xtpControlCustom, conMenu_Edit_NewItem, "")
    cbrCustom.Handle = picPane(3).hWnd
    
    Set objControl = NewToolBar(objBar, xtpControlButton, 9, "���Ĳ���...", True, , xtpButtonIconAndCaption)
    
End Function

Private Function InitGrid() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************

    '------------------------------------------------------------------------------------------------------------------
    Set mclsFile = New clsVsf
    With mclsFile
        
        Call .Initialize(Me.Controls, vfgFile, True, False, frmPubResource.GetImageList(16))
        Call .ClearColumn
        
        Call .AppendColumn("ID", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
        Call .AppendColumn("���", 750, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("����", 1500, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("Ӧд��", 810, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("�����", 810, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("����д", 810, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("Ҫ��", 1200, flexAlignLeftCenter, flexDTString, "", , True)
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    Set mclsEPRs = New clsVsf
    With mclsEPRs
        
        Call .Initialize(Me.Controls, vfgEPRs, False, False, frmPubResource.GetImageList(16))
        Call .ClearColumn
        
        Call .AppendColumn("ID", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
        Call .AppendColumn("����id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                
        Call .AppendColumn("�����", 1500, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("����", 1200, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("�Ա�", 750, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("��������", 1670, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "", True)
        Call .AppendColumn("������", 750, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("��д��", 750, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("���ʱ��", 1670, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "", True)
        
    End With
                    
End Function

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
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
        
        '����ͣ������
        '--------------------------------------------------------------------------------------------------------------
        Dim objPane As Pane
        Set objPane = dkpMan.CreatePane(conPane_File, 100, 100, DockLeftOf, Nothing): objPane.Title = "�ļ��嵥": objPane.Options = PaneNoCaption
        Set objPane = dkpMan.CreatePane(conPane_EPRs, 400, 300, DockRightOf, objPane): objPane.Title = "��д��¼": objPane.Options = PaneNoCaption
        Set objPane = dkpMan.CreatePane(conPane_Word, 600, 400, DockBottomOf, objPane): objPane.Title = "���ݼ��": objPane.Options = PaneNoCaption
        
        dkpMan.SetCommandBars cbsThis
        Call DockPannelInit(dkpMan)
                                
        Call InitCommandBar
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"
                        
        
                
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

Private Function RefreshData(ByVal intKind As Integer, ByVal lngDeptKey As Long, ByVal strDateFrom As String, ByVal strDateTo As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim intOut24h As Byte    '�Ƿ�����24Сʱ��Ժ��������0-������,1-���֣������Ƿ���24Сʱ�¼���Ӧ����ȷ��
    
    On Error GoTo errHand
    
    Select Case intKind
    '------------------------------------------------------------------------------------------------------------------
    Case 1
        strSQL = "Select F.ID, F.���, F.����, F.�¼� || 'ʱ��д' As Ҫ��, P.�˴� As Ӧд��, W.�����, W.����д" & vbNewLine & _
                "From (Select F.ID, F.���, F.����, F.�¼�, F.Ψһ, F.��дʱ��" & vbNewLine & _
                "       From (Select F.ID, F.���, F.����, F.ͨ��, A.����id, Q.�¼�, Q.Ψһ, Q.��дʱ��" & vbNewLine & _
                "              From �����ļ��б� F, ����Ӧ�ÿ��� A, ����ʱ��Ҫ�� Q" & vbNewLine & _
                "              Where F.ID = A.�ļ�id(+) And F.ID = Q.�ļ�id And F.���� = 1) F" & vbNewLine & _
                "       Where F.ͨ�� = 1 Or F.ͨ�� = 2 And F.����id = [1]) F," & vbNewLine & _
                "     (Select E.�¼�, Decode(E.�¼�, '����', ����, '����', ����, '����', ����, '����', ����) As �˴�" & vbNewLine & _
                "       From (Select Sum(Decode(����, 1, 0, 1)) As ����, Sum(Decode(����, 1, 0, Decode(����, 1, 0, 1))) As ����," & vbNewLine & _
                "                     Sum(Decode(����, 1, 0, Decode(����, 1, 1, 0))) As ����, Sum(Decode(����, 1, 1, 0)) As ����" & vbNewLine & _
                "              From ���˹Һż�¼" & vbNewLine & _
                "              Where ִ�в���id = [1] And Nvl(ִ��״̬, 0) <> 0 And ��¼����=1 And ��¼״̬=1 And �Ǽ�ʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "                    To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) P," & vbNewLine & _
                "            (Select Decode(Rownum, 1, '����', 2, '����', 3, '����', 4, '����') As �¼� From ������д�¼� Where Rownum < 5) E) P," & vbNewLine & _
                "     (Select �ļ�id, Sum(Decode(���ʱ��, Null, 0, 1)) As �����, Sum(Decode(���ʱ��, Null, 1, 0)) As ����д" & vbNewLine & _
                "       From ���Ӳ�����¼" & vbNewLine & _
                "       Where �������� = 1 And ����id + 0 = [1] And ����ʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By �ļ�id) W" & vbNewLine & _
                "Where F.�¼� = P.�¼� And P.�˴� > 0 And F.ID = W.�ļ�id(+)" & vbNewLine & _
                "Order By F.���"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, mfrmMain.Caption, lngDeptKey, strDateFrom, strDateTo)
    '------------------------------------------------------------------------------------------------------------------
    Case 2
        strSQL = "Select Sign(Nvl(Count(*), 0))" & vbNewLine & _
                "From (Select F.ID, F.ͨ��, A.����id" & vbNewLine & _
                "       From �����ļ��б� F, ����Ӧ�ÿ��� A, ����ʱ��Ҫ�� Q" & vbNewLine & _
                "       Where F.ID = A.�ļ�id(+) And F.ID = Q.�ļ�id And Q.�¼� In ('24Сʱ��Ժ', '24Сʱ����') And F.���� = 2) F" & vbNewLine & _
                "Where F.ͨ�� = 1 Or F.ͨ�� = 2 And F.����id = [1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, mfrmMain.Caption, mlngDeptId)
        If rs.RecordCount <= 0 Then
            intOut24h = 0
        Else
            intOut24h = rs.Fields(0).Value
        End If
        strSQL = "Select  F.ID, F.���, F.����, F.�¼� || Decode(Sign(F.��дʱ��), -1, 'ǰ', '��') || '��д' As Ҫ��," & vbNewLine & _
                "       Decode(F.Ψһ, 1, To_Char(P.�˴�), '<ѭ��>') As Ӧд��, W.�����, W.����д" & vbNewLine & _
                "From (Select F.ID, F.���, F.����, F.�¼�, F.Ψһ, F.��дʱ��" & vbNewLine & _
                "       From (Select F.ID, F.���, F.����, F.ͨ��, A.����id, Q.�¼�, Q.Ψһ, Q.��дʱ��" & vbNewLine & _
                "              From �����ļ��б� F, ����Ӧ�ÿ��� A, ����ʱ��Ҫ�� Q" & vbNewLine & _
                "              Where F.ID = A.�ļ�id(+) And F.ID = Q.�ļ�id And F.���� = 2) F" & vbNewLine & _
                "       Where F.ͨ�� = 1 Or F.ͨ�� = 2 And F.����id = [1]) F," & vbNewLine
        If intOut24h = 1 Then
            strSQL = strSQL & "     (Select E.�¼�, '��' As ʱ��, Decode(E.�¼�, '��Ժ', ��Ժ, '�״���Ժ', �״���Ժ, '�ٴ���Ժ', �ٴ���Ժ) As �˴�" & vbNewLine & _
                    "       From (Select Count(*) As ��Ժ, Sum(Decode(����Ժ, 1, 0, 1)) As �״���Ժ," & vbNewLine & _
                    "                     Sum(Decode(����Ժ, 1, 1, 0)) As �ٴ���Ժ" & vbNewLine & _
                    "              From ������ҳ" & vbNewLine & _
                    "              Where ��Ժ����id + 0 = [1] And Nvl(��Ժ����, Sysdate + 1) - ��Ժ���� > 1 And" & vbNewLine & _
                    "                    ��Ժ���� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) P," & vbNewLine & _
                    "            (Select Decode(Rownum, 1, '��Ժ', 2, '�״���Ժ', 3, '�ٴ���Ժ') As �¼� From ������д�¼� Where Rownum < 4) E" & vbNewLine & _
                    "       Union All" & vbNewLine & _
                    "       Select Decode(Sign(��Ժ���� - ��Ժ���� - 1), -1, Decode(��Ժ��ʽ, '����', '24Сʱ����', '24Сʱ��Ժ')," & vbNewLine & _
                    "                      Decode(��Ժ��ʽ, '����', '����', '��Ժ')) As �¼�, '��' As ʱ��, Count(*) As �˴�" & vbNewLine & _
                    "       From ������ҳ" & vbNewLine & _
                    "       Where ��Ժ����id + 0 = [1] And ��Ժ���� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "       Group By Decode(Sign(��Ժ���� - ��Ժ���� - 1), -1, Decode(��Ժ��ʽ, '����', '24Сʱ����', '24Сʱ��Ժ')," & vbNewLine & _
                    "                        Decode(��Ժ��ʽ, '����', '����', '��Ժ'))" & vbNewLine
        Else
            strSQL = strSQL & "     (Select E.�¼�, '��' As ʱ��, Decode(E.�¼�, '��Ժ', ��Ժ, '�״���Ժ', �״���Ժ, '�ٴ���Ժ', �ٴ���Ժ) As �˴�" & vbNewLine & _
                    "       From (Select Count(*) As ��Ժ, Sum(Decode(����Ժ, 1, 0, 1)) As �״���Ժ," & vbNewLine & _
                    "                     Sum(Decode(����Ժ, 1, 1, 0)) As �ٴ���Ժ" & vbNewLine & _
                    "              From ������ҳ" & vbNewLine & _
                    "              Where ��Ժ����id + 0 = [1] And" & vbNewLine & _
                    "                    ��Ժ���� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) P," & vbNewLine & _
                    "            (Select Decode(Rownum, 1, '��Ժ', 2, '�״���Ժ', 3, '�ٴ���Ժ') As �¼� From ������д�¼� Where Rownum < 4) E" & vbNewLine & _
                    "       Union All" & vbNewLine & _
                    "       Select Decode(��Ժ��ʽ, '����', '����', '��Ժ') As �¼�, '��' As ʱ��, Count(*) As �˴�" & vbNewLine & _
                    "       From ������ҳ" & vbNewLine & _
                    "       Where ��Ժ����id + 0 = [1] And ��Ժ���� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "       Group By Decode(��Ժ��ʽ, '����', '����', '��Ժ')" & vbNewLine
        End If
        strSQL = strSQL & "       Union All" & vbNewLine & _
                "       Select Decode(��ʼԭ��, 3, 'ת��', 7, '����') As �¼�, '��' As ʱ��, Count(*) As �˴�" & vbNewLine & _
                "       From ���˱䶯��¼" & vbNewLine & _
                "       Where ����id + 0 = [1] And ��ʼԭ�� In (3, 7) And Nvl(���Ӵ�λ, 0) = 0 And" & vbNewLine & _
                "             ��ʼʱ�� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By Decode(��ʼԭ��, 3, 'ת��', 7, '����')" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select Decode(��ֹԭ��, 3, 'ת��', 7, '����') As �¼�, 'ǰ' As ʱ��, Count(*) As �˴�" & vbNewLine & _
                "       From ���˱䶯��¼" & vbNewLine & _
                "       Where ����id + 0 = [1] And ��ֹԭ�� In (3, 7) And Nvl(���Ӵ�λ, 0) = 0 And" & vbNewLine & _
                "             ��ֹʱ�� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By Decode(��ֹԭ��, 3, 'ת��', 7, '����')" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select R.�¼�, E.ʱ��, Decode(E.ʱ��, 'ǰ', R.ǰ�˴�, '��', R.���˴�) As �˴�" & vbNewLine & _
                "       From (Select Decode(R.�������,'F', '����', Decode(I.��������, '7', '����', '����')) As �¼�," & vbNewLine & _
                "                     Sum(Decode(R.���˿���id, [1], 1, 0)) As ǰ�˴�, Sum(Decode(R.ִ�п���id, [1], 1, 0)) As ���˴�" & vbNewLine & _
                "              From ����ҽ����¼ R, ������ĿĿ¼ I, ����ҽ������ S" & vbNewLine & _
                "              Where R.ID = S.ҽ��id And R.������Ŀid = I.ID And" & vbNewLine & _
                "                    (R.������� = 'F' Or R.������� = 'Z' And I.�������� In ('7', '8')) And R.���id Is Null And" & vbNewLine & _
                "                    R.ҽ����Ч = 1 And (R.ҽ��״̬ = 8 Or R.ҽ��״̬ = 9) And" & vbNewLine & _
                "                    (R.���˿���id + 0 = [1] Or R.ִ�п���id + 0 = [1]) And S.����ʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "                    To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "              Group By Decode(R.�������,'F', '����', Decode(I.��������, '7', '����', '����'))) R," & vbNewLine & _
                "            (Select Decode(Rownum, 1, 'ǰ', 2, '��') As ʱ�� From ������д�¼� Where Rownum < 3) E) P," & vbNewLine
        
        strSQL = strSQL & "     (Select �ļ�id, Sum(Decode(���ʱ��, Null, 0, 1)) As �����, Sum(Decode(���ʱ��, Null, 1, 0)) As ����д" & vbNewLine & _
                "       From ���Ӳ�����¼" & vbNewLine & _
                "       Where �������� = 2 And ����id + 0 = [1] And ����ʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By �ļ�id) W" & vbNewLine & _
                "Where  F.�¼� = P.�¼� and Decode(Sign(F.��дʱ��), -1, 'ǰ', '��') = P.ʱ�� And P.�˴� > 0 And F.ID = W.�ļ�id(+)" & vbNewLine & _
                "Order By F.���"
                
        Set rs = zlDatabase.OpenSQLRecord(strSQL, mfrmMain.Caption, lngDeptKey, strDateFrom, strDateTo)
        
    '------------------------------------------------------------------------------------------------------------------
    Case 4
        strSQL = "Select F.ID, F.���, F.����, F.�¼� || Decode(Sign(F.��дʱ��), -1, 'ǰ', '��') || '��д' As Ҫ��," & vbNewLine & _
                "       Decode(F.Ψһ, 1, To_Char(P.�˴�), '<ѭ��>') As Ӧд��, W.�����, W.����д" & vbNewLine & _
                "From (Select F.ID, F.���, F.����, F.�¼�, F.Ψһ, F.��дʱ��" & vbNewLine & _
                "       From (Select F.ID, F.���, F.����, F.ͨ��, A.����id, Q.�¼�, Q.Ψһ, Q.��дʱ��" & vbNewLine & _
                "              From �����ļ��б� F, ����Ӧ�ÿ��� A, ����ʱ��Ҫ�� Q" & vbNewLine & _
                "              Where F.ID = A.�ļ�id(+) And F.ID = Q.�ļ�id And F.���� = 4) F" & vbNewLine & _
                "       Where F.ͨ�� = 1 Or F.ͨ�� = 2 And F.����id = [1]) F," & vbNewLine
        strSQL = strSQL & "     (Select '��Ժ' As �¼�, '��' As ʱ��, Count(*) As �˴�" & vbNewLine & _
                "       From ������ҳ" & vbNewLine & _
                "       Where ��Ժ����id + 0 = [1] And ��Ժ���� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select Decode(��Ժ��ʽ, '����', '����', '��Ժ') As �¼�, '��' As ʱ��, Count(*) As �˴�" & vbNewLine & _
                "       From ������ҳ" & vbNewLine & _
                "       Where ��ǰ����id + 0 = [1] And ��Ժ���� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By Decode(��Ժ��ʽ, '����', '����', '��Ժ')" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select Decode(��ʼԭ��, 3, 'ת��', 8, '����') As �¼�, '��' As ʱ��, Count(*) As �˴�" & vbNewLine & _
                "       From ���˱䶯��¼" & vbNewLine & _
                "       Where ����id + 0 = [1] And ��ʼԭ�� In (3, 8) And Nvl(���Ӵ�λ, 0) = 0 And" & vbNewLine & _
                "             ��ʼʱ�� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By Decode(��ʼԭ��, 3, 'ת��', 8, '����')" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select Decode(��ֹԭ��, 3, 'ת��', 8, '����') As �¼�, 'ǰ' As ʱ��, Count(*) As �˴�" & vbNewLine & _
                "       From ���˱䶯��¼" & vbNewLine & _
                "       Where ����id + 0 = [1] And ��ֹԭ�� In (3, 8) And Nvl(���Ӵ�λ, 0) = 0 And" & vbNewLine & _
                "             ��ֹʱ�� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By Decode(��ֹԭ��, 3, 'ת��', 8, '����')) P," & vbNewLine
        strSQL = strSQL & "     (Select �ļ�id, Sum(Decode(���ʱ��, Null, 0, 1)) As �����, Sum(Decode(���ʱ��, Null, 1, 0)) As ����д" & vbNewLine & _
                "       From ���Ӳ�����¼" & vbNewLine & _
                "       Where �������� = 4 And ����id + 0 = [1] And ����ʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By �ļ�id) W" & vbNewLine & _
                "Where F.�¼� = P.�¼� And Decode(Sign(F.��дʱ��), -1, 'ǰ', '��') = P.ʱ�� And P.�˴� > 0 And F.ID = W.�ļ�id(+)" & vbNewLine & _
                "Order By F.���"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, mfrmMain.Caption, lngDeptKey, strDateFrom, strDateTo)
    End Select
    
    Dim lngCount As Long
    
    mclsFile.ClearGrid
    mclsEPRs.ClearGrid
'    If Not (mfrmEPRAuditMonitor Is Nothing) Then mfrmEPRAuditMonitor.zlClearData
    If Not (mclsDockAduit Is Nothing) Then mclsDockAduit.zlClearMonitor
    
    With vfgFile
        If rs.RecordCount > 0 Then
            .Clear
            Set .DataSource = rs
            .ColWidth(0) = 0: .ColHidden(0) = True
            .ColAlignment(4) = flexAlignRightCenter
            For lngCount = 1 To .Cols - 1
                .FixedAlignment(lngCount) = flexAlignCenterCenter
            Next
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
    
    If Me.ActiveControl Is Nothing Then
        Set objPrint.Body = vfgFile
        objPrint.Title.Text = mstrDeptName & "�����ļ��б�"
    Else
        If Me.ActiveControl.Name = vfgFile.Name Then
            Set objPrint.Body = vfgFile
            objPrint.Title.Text = mstrDeptName & "�����ļ��б�"
        Else
            Set objPrint.Body = vfgEPRs
            objPrint.Title.Text = mstrDeptName & vfgFile.TextMatrix(vfgFile.Row, 2) & "�嵥"
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
        Call RefreshData(mintKind, mlngDeptId, mstrDateFrom, mstrDateTo)
    End If
    
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRecordId As Long
    
    Select Case Control.ID
    Case 9
        
        With Me.vfgEPRs
            lngRecordId = Val(.TextMatrix(.Row, .ColIndex("ID")))
        End With
        
        If lngRecordId > 0 Then
            If mclsDockAduit Is Nothing Then Set mclsDockAduit = New zlRichEPR.clsDockAduits
            Call mclsDockAduit.zlOpenEPRDocument(lngRecordId, mfrmMain)
        End If
        
    End Select
    
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case 9
        With vfgEPRs
            Control.Enabled = (Val(.TextMatrix(.Row, 0)) > 0)
        End With
    End Select
    
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_File
        Item.Handle = picPane(0).hWnd
    Case conPane_EPRs
        Item.Handle = picPane(1).hWnd
    Case conPane_Word

        If mclsDockAduit Is Nothing Then Set mclsDockAduit = New zlRichEPR.clsDockAduits
        Call mclsDockAduit.zlInitMonitor(Me)
        Item.Handle = mclsDockAduit.zlGetFormAuditMonitor.hWnd

    End Select
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    Call SetPaneRange(dkpMan, 1, 100, 15, 360, Me.ScaleHeight)
    
    dkpMan.RecalcLayout
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If Not (mclsDockAduit Is Nothing) Then Set mclsDockAduit = Nothing
    
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        vfgFile.Move 0, 0, picPane(Index).Width, picPane(Index).Height
        cboDept.Move -30, -30, picPane(3).Width + 45
    Case 1
        vfgEPRs.Move 0, 0, picPane(Index).Width, picPane(Index).Height
    End Select
End Sub

Private Sub vfgEPRs_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vfgEPRs
        If OldRow <> NewRow And NewRow > 0 Then
            
            If Not (mclsDockAduit Is Nothing) Then
                
                Call mclsDockAduit.zlRefreshMonitor(Val(vfgFile.TextMatrix(vfgFile.Row, 0)))
                
            End If
                        
        End If
    End With
End Sub

Private Function RefreshPatient(ByVal intKind As Integer, ByVal lngFileID As Long, ByVal lngDeptKey As Long, ByVal strDateFrom As String, ByVal strDateTo As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    Select Case intKind
    '------------------------------------------------------------------------------------------------------------------
    Case 1
        strSQL = "Select W.ID, P.����id, P.�����, P.����, P.�Ա�, To_Char(P.ִ��ʱ��, 'mm-dd hh24:mi') As ��������, W.������ As ��д��," & vbNewLine & _
                "       To_Char(W.���ʱ��, 'mm-dd hh24:mi') As ���ʱ��" & vbNewLine & _
                "From ���Ӳ�����¼ W, ���˹Һż�¼ P" & vbNewLine & _
                "Where W.��ҳid = P.ID And W.�������� = 1 And W.����id + 0 = [1] And W.�ļ�id + 0 = [4] And" & vbNewLine & _
                "      W.����ʱ�� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "Order By P.ִ��ʱ��"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, mfrmMain.Caption, lngDeptKey, strDateFrom, strDateTo, lngFileID)
    '------------------------------------------------------------------------------------------------------------------
    Case 2
        strSQL = "Select W.ID, I.����id, P.סԺ��, I.����, I.�Ա�, To_Char(P.��Ժ����, 'mm-dd hh24:mi') As ��Ժ����, W.������ As ��д��," & vbNewLine & _
                "       To_Char(W.���ʱ��, 'mm-dd hh24:mi') As ���ʱ��" & vbNewLine & _
                "From ���Ӳ�����¼ W, ������ҳ P, ������Ϣ I" & vbNewLine & _
                "Where I.����id = P.����id And P.����id = W.����id And P.��ҳid = W.��ҳid And W.�������� = 2 And W.����id + 0 = [1] And" & vbNewLine & _
                "      W.�ļ�id + 0 = [4] And W.����ʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "      To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "Order By ��Ժ����"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptId, mstrDateFrom, mstrDateTo, lngFileID)
    '------------------------------------------------------------------------------------------------------------------
    Case 4
        strSQL = "Select W.ID, I.����id, P.סԺ��, I.����, I.�Ա�, To_Char(P.��Ժ����, 'mm-dd hh24:mi') As ��Ժ����, W.������ As ��д��," & vbNewLine & _
                "       To_Char(W.���ʱ��, 'mm-dd hh24:mi') As ���ʱ��" & vbNewLine & _
                "From ���Ӳ�����¼ W, ������ҳ P, ������Ϣ I" & vbNewLine & _
                "Where I.����id = P.����id And P.����id = W.����id And P.��ҳid = W.��ҳid And W.�������� = 4 And W.����id + 0 = [1] And" & vbNewLine & _
                "      W.�ļ�id + 0 = [4] And W.����ʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "      To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "Order By ��Ժ����"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, mfrmMain.Caption, lngDeptKey, strDateFrom, strDateTo, lngFileID)
    End Select
    
    Dim lngCount As Long
    
    mclsEPRs.ClearGrid
    
    With Me.vfgEPRs
        
        If rs.RecordCount > 0 Then
            .Clear
            Set .DataSource = rs
            .ColWidth(0) = 0: .ColHidden(0) = True
            For lngCount = 1 To .Cols - 1
                .FixedAlignment(lngCount) = flexAlignCenterCenter
            Next
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

Private Sub vfgFile_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vfgFile
        If OldRow <> NewRow And NewRow > 0 Then
            
            Call RefreshPatient(mintKind, Val(.TextMatrix(NewRow, 0)), mlngDeptId, mstrDateFrom, mstrDateTo)
            
        End If
    End With
End Sub

