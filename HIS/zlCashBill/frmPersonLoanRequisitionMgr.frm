VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.Form frmPersonLoanRequisitionMgr 
   BorderStyle     =   0  'None
   Caption         =   "����б�"
   ClientHeight    =   8175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   4710
      Left            =   75
      TabIndex        =   0
      Top             =   1620
      Width           =   7980
      _Version        =   589884
      _ExtentX        =   14076
      _ExtentY        =   8308
      _StockProps     =   0
      BorderStyle     =   2
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPrint 
      Height          =   1365
      Left            =   8970
      TabIndex        =   1
      Top             =   2550
      Visible         =   0   'False
      Width           =   540
      _cx             =   952
      _cy             =   2408
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
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   1320
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonLoanRequisitionMgr.frx":0000
            Key             =   "�ȴ����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonLoanRequisitionMgr.frx":059A
            Key             =   "�ܾ�����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonLoanRequisitionMgr.frx":0B34
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonLoanRequisitionMgr.frx":10CE
            Key             =   "ȡ�����"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPersonLoanRequisitionMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar
Private mstrPrivs As String, mlngModule As Long, mArrFilter As Variant  '��������
Private mcbsThis As Object
Private Type rptColIndexType    '������
    ColID  As Integer
    Colͼ�� As Integer
    Col״̬ As Integer
    Col������ As Integer
    Col����ʱ�� As Integer
    Col��ע As Integer
    Col�����  As Integer
    Col����� As Integer
    Col���ʱ�� As Integer
    Colȡ��ʱ�� As Integer
    Colȡ��ԭ�� As Integer
End Type
Private mRptCol As rptColIndexType

Public Function zlReLoadData(ByVal mcllFilter As Variant) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¼�������
    '����:���سɹ�,����true,���򷵻�Flase
    '����:���˺�
    '����:2009-09-07 14:43:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    Set mArrFilter = mcllFilter
    Call LoadDataToRpt
    zlReLoadData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub InitReportColumn()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ȡ����
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-09-07 11:14:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCol As ReportColumn, i As Long
    
    With mRptCol
        .ColID = 0: i = i + 1
        .Colͼ�� = i: i = i + 1
        .Col״̬ = i: i = i + 1
        .Col������ = i: i = i + 1
        .Col����ʱ�� = i: i = i + 1
        .Col��ע = i: i = i + 1
        
        .Col����� = i: i = i + 1
        .Col����� = i: i = i + 1
        .Col���ʱ�� = i: i = i + 1
        .Colȡ��ʱ�� = i: i = i + 1
        .Colȡ��ԭ�� = i: i = i + 1
    End With
    With rptList
        '��ǰ˳��:ID,Colͼ��,Col״̬,������,����ʱ��,�����,�����,���ʱ��, ȡ��ʱ��,ȡ��ԭ��
        
       ' ����˳�������(�������Ϊ����)�ı��,Ҫ��Find(�к�)������,���Կ���Record(�к�)����������
        Set objCol = .Columns.Add(mRptCol.ColID, "ID", 0, False)
            objCol.Sortable = False: objCol.Visible = False
            
        Set objCol = .Columns.Add(mRptCol.Colͼ��, "", 25, False)
        
        Set objCol = .Columns.Add(mRptCol.Col״̬, "״̬", 55, False): objCol.Visible = False
        Set objCol = .Columns.Add(mRptCol.Col������, "������", 55, True): objCol.Visible = False
            'objCol.TreeColumn = True: 'objCol.Visible = False
            'objCol.Sortable = False: objCol.AllowDrag = False
        Set objCol = .Columns.Add(mRptCol.Col����ʱ��, "����ʱ��", 136, True)
            objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(mRptCol.Col��ע, "��ע", 300, True)
        Set objCol = .Columns.Add(mRptCol.Col�����, "�����", 65, True)
        objCol.Alignment = xtpAlignmentRight
        Set objCol = .Columns.Add(mRptCol.Col�����, "�����", 65, True)
        Set objCol = .Columns.Add(mRptCol.Col���ʱ��, "���ʱ��", 136, True)
        Set objCol = .Columns.Add(mRptCol.Colȡ��ʱ��, "ȡ��ʱ��", 136, True)
        Set objCol = .Columns.Add(mRptCol.Colȡ��ԭ��, "ȡ��ԭ��", 200, True)
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = objCol.Index = mRptCol.Col״̬
            objCol.Groupable = objCol.Index = mRptCol.Col�����
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '�з�����ʱ�������߱��ϻ�����һ������
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�����������Ա..."
        End With
        .PreviewMode = True: .AllowColumnRemove = False: .MultipleSelection = False '������SelectionChanged�¼�
        .ShowItemsInGroups = False: .SetImageList Me.imgList
        .GroupsOrder.Add .Columns(mRptCol.Col״̬): .GroupsOrder(0).SortAscending = True       '����֮��,��������в���ʾ,�����е������ǲ����
         
        '����֮�����ʧȥ��¼���е�˳��,���ǿ�м���������
        '.SortOrder.Add .Columns(mRptCol.Col�����): .SortOrder(0).SortAscending = True
    End With
End Sub

Private Sub Form_Load()
    '��ʼ��Ȩ�޴�
    mlngModule = glngModul: mstrPrivs = gstrPrivs
    '��ʼ����
    Call InitReportColumn
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With rptList
        .Left = ScaleLeft: .Top = ScaleTop
        .Width = ScaleWidth: .Height = ScaleHeight
    End With
End Sub
Private Sub LoadDataToRpt()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ݸ�����
    '����:���˺�
    '����:2009-09-07 11:53:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFilter As String, rsTemp As New ADODB.Recordset, str����� As String, j As Long, i As Long
    Dim objParent As ReportRecord, objRecord As ReportRecord, objItem As ReportRecordItem
    Dim strTemp As String
    Err = 0: On Error GoTo ErrHand:
    
    If CStr(mArrFilter("����ʱ��")(0)) <> "1901-01-01" And CStr(mArrFilter("���ʱ��")(0)) <> "1901-01-01" And CStr(mArrFilter("ȡ��ʱ��")(0)) <> "1901-01-01" Then
        strFilter = "   (����ʱ�� between [1] and [2] or ���ʱ�� between [3] and [4] or ȡ��ʱ�� between [5] and [6])  "
    ElseIf CStr(mArrFilter("����ʱ��")(0)) <> "1901-01-01" And CStr(mArrFilter("���ʱ��")(0)) <> "1901-01-01" And CStr(mArrFilter("ȡ��ʱ��")(0)) = "1901-01-01" Then
        strFilter = "   (����ʱ�� between [1] and [2] or ���ʱ�� between [3] and [4]   )   "
    ElseIf CStr(mArrFilter("����ʱ��")(0)) <> "1901-01-01" And CStr(mArrFilter("���ʱ��")(0)) = "1901-01-01" And CStr(mArrFilter("ȡ��ʱ��")(0)) <> "1901-01-01" Then
        strFilter = "   (����ʱ�� between [1] and [2] or ȡ��ʱ�� between [5] and [6])   "
    ElseIf CStr(mArrFilter("����ʱ��")(0)) = "1901-01-01" And CStr(mArrFilter("���ʱ��")(0)) <> "1901-01-01" And CStr(mArrFilter("ȡ��ʱ��")(0)) <> "1901-01-01" Then
        strFilter = "   ( ���ʱ�� between [3] and [4] or ȡ��ʱ�� between [5] and [6])  "
    ElseIf CStr(mArrFilter("����ʱ��")(0)) <> "1901-01-01" Then
        strFilter = "   (����ʱ�� between [1] and [2]   ) and ���ʱ�� is  Null "
    ElseIf CStr(mArrFilter("���ʱ��")(0)) <> "1901-01-01" Then
        strFilter = "   (���ʱ�� between [3] and [4])"
    Else
        strFilter = "   (ȡ��ʱ�� between [5] and [6] )"
    End If
 
    strFilter = strFilter & " and ����� = [7]"
    If CStr(mArrFilter("�����")) <> "" Then strFilter = strFilter & " and ����� like [8]"
    
    zlCommFun.ShowFlash "����װ�ؽ������,���Ժ�..."

    gstrSQL = " " & _
    "    Select Id, �����, ��ע, �����, to_char(����ʱ��,'yyyy-mm-dd hh24:mi:ss') as ����ʱ�� ,  " & _
    "           �����, to_char(���ʱ��,'yyyy-mm-dd hh24:mi:ss') as ���ʱ��, " & _
    "           to_char(ȡ��ʱ��,'yyyy-mm-dd hh24:mi:ss') as ȡ��ʱ��, ȡ��ԭ��, " & _
    "           decode(���ʱ��,NULL,'�ȴ����',decode(ȡ��ʱ��,NULL,'�Ѿ����','���ȡ��')) as ״̬, " & _
    "           decode(���ʱ��,NULL,1,decode(ȡ��ʱ��,NULL,2,3)) as ״̬��־ " & _
    "    From ��Ա����¼ " & _
    "    Where " & strFilter & _
    "    Order by ״̬,�����,����ʱ��"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
        CDate(mArrFilter("����ʱ��")(0)), CDate(mArrFilter("����ʱ��")(1)), _
        CDate(mArrFilter("���ʱ��")(0)), CDate(mArrFilter("���ʱ��")(1)), _
        CDate(mArrFilter("ȡ��ʱ��")(0)), CDate(mArrFilter("ȡ��ʱ��")(1)), _
        UserInfo.����, GetMatchingSting(CStr(mArrFilter("�����")), False))
    
    rptList.Records.DeleteAll
    rptList.Columns(mRptCol.ColID).Visible = False
    With rsTemp
        Do While Not .EOF
            Set objRecord = Me.rptList.Records.Add()
            objRecord.Tag = CStr(Nvl(!ID))  '���ڶ�λ
            
            'ID,Colͼ��,Col״̬,������,����ʱ��,�����,�����,���ʱ��, ȡ��ʱ��,ȡ��ԭ��
            Set objItem = objRecord.AddItem(Val(Nvl(!ID)))  '������Value��������
            objItem.Caption = CStr(Nvl(!ID))
        
            'ͼ��:ע�����������Ǵ�0��ʼ��š�
            '     ͼ��Value���ڴ���Ƿ����ύ��飬����Ŷ�ȡ
            Set objItem = objRecord.AddItem(-1)
            objItem.Caption = " "
            objItem.Icon = Decode(Nvl(!״̬), "�Ѿ����", 2, "���ȡ��", 3, 1)
        
        
            Set objItem = objRecord.AddItem(Val(Nvl(!״̬��־)))
            objItem.Caption = Nvl(!״̬)
'            If Nvl(!״̬) = "���ȡ��" Then
'                objRecord.PreviewText = "  ȡ��ԭ��:" & Nvl(!ȡ��ԭ��)
'            End If
            
            Set objItem = objRecord.AddItem(CStr(Nvl(!�����)))
            objItem.Caption = CStr(Nvl(!�����))
            
            Set objItem = objRecord.AddItem(CStr(Nvl(!����ʱ��)))
            objItem.Caption = CStr(Nvl(!����ʱ��))
            
            Set objItem = objRecord.AddItem(CStr(Nvl(!��ע)))
            
            Set objItem = objRecord.AddItem(CStr(Nvl(!�����)))
            objItem.Caption = Format(Val(Nvl(!�����)), "###0.00;-###0.00")
            
            Set objItem = objRecord.AddItem(CStr(Nvl(!�����)))
            objItem.Caption = CStr(Nvl(!�����))
            Set objItem = objRecord.AddItem(CStr(Nvl(!���ʱ��)))
            objItem.Caption = CStr(Nvl(!���ʱ��))
        
            Set objItem = objRecord.AddItem(CStr(Nvl(!ȡ��ʱ��)))
            objItem.Caption = CStr(Nvl(!ȡ��ʱ��))
            Set objItem = objRecord.AddItem(CStr(Nvl(!ȡ��ԭ��)))
            objItem.Caption = CStr(Nvl(!ȡ��ԭ��))
       
            '��ʾ��ɫ
            For j = 0 To rptList.Columns.Count - 1
                If j = mRptCol.Colȡ��ʱ�� Then
                    objRecord.Item(j).ForeColor = vbRed
                End If
            Next
           .MoveNext
        Loop
    End With
    rptList.Populate
    zlCommFun.StopFlash
   Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    zlCommFun.StopFlash
End Sub
Private Function GetCurrRecordFun(Optional ByRef lngID As Long = 0) As Byte
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��������
    '����:���ص�ǰѡ���ID
    '����:0-��ǰѡ��ķ���,�������κδ���;1-�ȴ����,2-�Ѿ����,��δȡ�����;3-�Ѿ���ȡ�����
    '����:���˺�
    '����:2009-09-09 09:26:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    With rptList.SelectedRows(0)
        If .GroupRow Then Exit Function
        lngID = Val(.Record(mRptCol.ColID).Value)
        GetCurrRecordFun = Val(.Record(mRptCol.Col״̬).Value)
        
    End With
    If lngID = 0 Then Exit Function
End Function
Private Function DeleteLoanRequisition() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ɾ��������
    '����:ɾ���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-09-09 09:42:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, str����� As String
    With rptList.SelectedRows(0)
        If .GroupRow Then Exit Function
        lngID = Val(.Record(mRptCol.ColID).Value)
        str����� = .Record(mRptCol.Col�����).Value
        If MsgBox("�����Ҫɾ����" & str����� & "���Ľ��������", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Function
    End With
    If lngID = 0 Then Exit Function
    
    'Zl_��Ա����¼_Delete(Id_In In ��Ա����¼.ID%Type) Is
    Err = 0: On Error GoTo ErrHand:
    gstrSQL = "Zl_��Ա����¼_Delete(" & lngID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    DeleteLoanRequisition = True
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function zlDefCommandBars(ByVal cbsThis As Object) As Boolean
    '----------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008/1/9
    '----------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup
      
    Err = 0: On Error GoTo ErrHand:
    '-----------------------------------------------------
    Set mcbsThis = cbsThis
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    
    cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    mcbrMenuBar.ID = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_BillPrintSet, "����ӡ����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): mcbrControl.BeginGroup = True
    End With

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    mcbrMenuBar.ID = conMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "������(&A)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸Ľ��(&M)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ�����(&D)")
    End With

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    mcbrMenuBar.ID = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        
        mcbrControl.BeginGroup = True
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_Expend, "չ��/�۵���(&X)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "�۵�������(&L)", -1, False)
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_Expend_AllExpend, "չ��������(&X)", -1, False)
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_Expend_CurCollapse, "�۵���ǰ��(&C)", -1, False): mcbrControl.BeginGroup = True
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_Expend_CurExpend, "չ����ǰ��(&E)", -1, False)
        End With
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    mcbrMenuBar.ID = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): mcbrControl.BeginGroup = True
    End With
    
    '�����
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FSHIFT, VK_DELETE, conMenu_Edit_Delete
        .Add FCONTROL, vbKeyAdd, conMenu_View_Expend_AllExpend 'չ��������
        .Add FCONTROL, vbKeySubtract, conMenu_View_Expend_AllCollapse '�۵�������
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '���ò����ò˵�
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    '-----------------------------------------------------
    '����������
    Set mcbrToolBar = cbsThis.Add("������", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "������"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
     zlDefCommandBars = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------
Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRow As Long, objRow As ReportRow
    Dim lngID  As Long
    '------------------------------------
    Select Case Control.ID
    'bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Case conMenu_File_Preview: Call zlRptPrint(2)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_View_Expend_CurCollapse '�۵���ǰ��
        If rptList.SelectedRows.Count > 0 Then
            If rptList.SelectedRows(0).GroupRow Then
                rptList.SelectedRows(0).Expanded = False
            ElseIf Not rptList.SelectedRows(0).ParentRow Is Nothing Then
                If rptList.SelectedRows(0).ParentRow.GroupRow Then
                    rptList.SelectedRows(0).ParentRow.Expanded = False
                End If
            End If
        End If
        '���۵���λ��������,�����Զ�������¼�
        'Call rptList_SelectionChanged
    Case conMenu_View_Expend_CurExpend 'չ����ǰ��
        If rptList.SelectedRows.Count > 0 Then
            rptList.SelectedRows(0).Expanded = True
        End If
    Case conMenu_View_Expend_AllCollapse '�۵�������
        For Each objRow In rptList.Rows
            If objRow.GroupRow Then objRow.Expanded = False
        Next
        '���۵���λ��������,�����Զ�������¼�
        'Call rptList_SelectionChanged
    Case conMenu_View_Expend_AllExpend 'չ��������
        For Each objRow In rptList.Rows
            If objRow.GroupRow Then objRow.Expanded = True
        Next
        
    Case conMenu_Edit_NewItem   '����
        If frmPersonLoanRequisitionEdit.ShowEdit(Me, FN_����, mstrPrivs, mlngModule) = False Then Exit Sub
        '����ˢ������
        Call LoadDataToRpt
    Case conMenu_Edit_Modify    '�޸�
        With rptList.SelectedRows(0)
            If .GroupRow Then Exit Sub
            lngID = Val(.Record(mRptCol.ColID).Value)
        End With
        If lngID = 0 Then Exit Sub
            
        If frmPersonLoanRequisitionEdit.ShowEdit(Me, FN_�޸�, mstrPrivs, mlngModule, lngID) = False Then Exit Sub
        '����ˢ������
        Call LoadDataToRpt
    Case conMenu_Edit_Delete 'ɾ������
        If DeleteLoanRequisition = False Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_View_Refresh   'ˢ��
        '����ˢ������
        Call LoadDataToRpt
    Case Else
        If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            'ִ�з�������ǰģ��ı���
            With rptList.SelectedRows(0)
                If .GroupRow = False Then
                    lngID = Val(.Record(mRptCol.ColID).Value)
                End If
            End With
            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "ID=" & lngID)
        End If
    End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean, lngID As Long, blnEnabled As Boolean
    If Me.Visible = False Then Exit Sub
    
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (Me.rptList.Records.Count >= 1)
    Case conMenu_Edit_NewItem '����
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "������")
            Control.Enabled = Control.Visible
    Case conMenu_Edit_Modify
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "�޸Ľ��")
        Control.Enabled = Control.Visible And GetCurrRecordFun(lngID) = 1 '0-��ǰѡ��ķ���,�������κδ���;1-�ȴ����,2-�Ѿ����,��δȡ�����;3-�Ѿ���ȡ�����
        Control.Enabled = Control.Enabled And lngID <> 0
        
    Case conMenu_Edit_Delete
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "ɾ�����")
        Control.Enabled = Control.Visible And GetCurrRecordFun(lngID) = 1 '0-��ǰѡ��ķ���,�������κδ���;1-�ȴ����,2-�Ѿ����,��δȡ�����;3-�Ѿ���ȡ�����
        Control.Enabled = Control.Enabled And lngID <> 0
    Case conMenu_View_Expend_CurExpend 'չ����ǰ��
        blnEnabled = False
        If rptList.SelectedRows.Count > 0 Then
            If rptList.SelectedRows(0).GroupRow Then
                blnEnabled = Not rptList.SelectedRows(0).Expanded
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_View_Expend_CurCollapse '�۵���ǰ��
        blnEnabled = False
        If rptList.SelectedRows.Count > 0 Then
            If rptList.SelectedRows(0).GroupRow Then
                blnEnabled = rptList.SelectedRows(0).Expanded
            ElseIf Not rptList.SelectedRows(0).ParentRow Is Nothing Then
                If rptList.SelectedRows(0).ParentRow.GroupRow Then
                    blnEnabled = rptList.SelectedRows(0).ParentRow.Expanded
                End If
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_View_Expend '�۵�/չ����
        Control.Enabled = rptList.GroupsOrder.Count > 0 And rptList.Rows.Count > 0
    Case conMenu_View_Refresh
        
    End Select
End Sub
Private Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���д�ӡ,Ԥ���������EXCEL
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '����:���˺�
    '����:2009-09-09 11:24:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, r As Long, rptRow As ReportRow, lngRow As Long
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
     
    With rptList
        vsPrint.Redraw = flexRDNone
        vsPrint.Cols = .Columns.Count + 1
        For i = 0 To .Columns.Count - 1
            vsPrint.TextMatrix(0, i) = .Columns(i).Caption
            vsPrint.ColWidth(i) = .Columns(i).Width * Screen.TwipsPerPixelX
        Next
        vsPrint.Clear 1
        vsPrint.Rows = 2: lngRow = 1
        For r = 0 To .Rows.Count - 1
            Set rptRow = .Rows(r)
            If rptRow.GroupRow = False Then
                For i = 0 To .Columns.Count - 1
                    vsPrint.TextMatrix(lngRow, i) = rptRow.Record(i).Caption
                Next
                lngRow = lngRow + 1
                vsPrint.Rows = vsPrint.Rows + 1
            End If
        Next
        vsPrint.Redraw = flexRDBuffered
    End With
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstr��λ���� & "����嵥"
    
    If CStr(mArrFilter("����ʱ��")(0)) <> "1901-01-01" Then
        objRow.Add "����ʱ�䣺" & CStr(mArrFilter("����ʱ��")(0)) & "��" & CStr(mArrFilter("����ʱ��")(1))
    End If
    If CStr(mArrFilter("���ʱ��")(0)) <> "1901-01-01" Then
        objRow.Add "���ʱ�䣺" & CStr(mArrFilter("���ʱ��")(0)) & "��" & CStr(mArrFilter("���ʱ��")(1))
    End If
    If CStr(mArrFilter("ȡ��ʱ��")(0)) <> "1901-01-01" Then
        objRow.Add "ȡ��ʱ�䣺" & CStr(mArrFilter("ȡ��ʱ��")(0)) & "��" & CStr(mArrFilter("ȡ��ʱ��")(1))
    End If
    
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "����ˣ�" & UserInfo.����
    If CStr(mArrFilter("�����")) <> "" Then objRow.Add "����ˣ�" & mArrFilter("�����")
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    Err = 0: On Error GoTo ErrHand:
    Set objPrint.Body = vsPrint
    If bytFunc = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytPrn
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub
Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim objHitTest As ReportHitTestInfo
    Dim objPopup As CommandBarPopup
        
    If Button = 2 Then
        Set objHitTest = rptList.HitTest(x, y)
        If objHitTest.ht = xtpHitTestReportArea And Not objHitTest.Row Is Nothing Then
            If objHitTest.Row.GroupRow Then
                Set objPopup = mcbsThis.FindControl(, conMenu_View_Expend, , True)
            ElseIf objHitTest.Row.Childs.Count = 0 Then
                Set objPopup = mcbsThis.ActiveMenuBar.Controls(2)
            End If
        Else
            Set objPopup = mcbsThis.ActiveMenuBar.Controls(2)
        End If
        rptList.SetFocus
        If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
    End If
End Sub
