VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTendItemMan 
   Caption         =   "�����¼��Ŀ����"
   ClientHeight    =   6750
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10305
   Icon            =   "frmTendItem.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   10305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   2040
      Left            =   435
      TabIndex        =   0
      Top             =   2295
      Width           =   1995
      _Version        =   589884
      _ExtentX        =   3519
      _ExtentY        =   3598
      _StockProps     =   0
      BorderStyle     =   2
      ShowGroupBox    =   -1  'True
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VB.PictureBox picColorItem 
      BackColor       =   &H00E4E8EA&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   6225
      ScaleHeight     =   255
      ScaleWidth      =   2520
      TabIndex        =   4
      Top             =   4470
      Width           =   2520
      Begin VB.Label lblColor 
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   735
         TabIndex        =   6
         Top             =   45
         Width           =   765
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��¼ɫ��"
         Height          =   180
         Left            =   30
         TabIndex        =   5
         Top             =   45
         Width           =   720
      End
   End
   Begin XtremeSuiteControls.TaskPanel tkp 
      Height          =   3030
      Left            =   6630
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1290
      Width           =   3045
      _Version        =   589884
      _ExtentX        =   5371
      _ExtentY        =   5345
      _StockProps     =   64
      VisualTheme     =   5
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   6375
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTendItem.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15266
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsList 
      Left            =   8715
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendItem.frx":0E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendItem.frx":767E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendItem.frx":77D8
            Key             =   "User"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfPrint 
      Height          =   540
      Left            =   6885
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5610
      Visible         =   0   'False
      Width           =   1095
      _cx             =   1931
      _cy             =   952
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
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
      GridColor       =   -2147483632
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
      DesignerControls=   "frmTendItem.frx":7932
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   690
      Top             =   60
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmTendItemMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'���ڼ���������########################################################################################################

Private Enum mCol
    ͼ�� = 0
    ��Ŀ���
    ��������
    ��Ŀ����
    ��Ŀ����
    ��Ŀ����
    ��ĿС��
    ��Ŀ��λ
    ��Ŀ��ʾ
    ��Ŀֵ��
    ��ͻ���
    ������Ŀ
    ������Ŀ
    ���ò���
    Ӧ�÷�ʽ
    ��Ŀ����
    Ӧ�ó���
    ��Ŀid
    ˵��
End Enum

Private mstrPrivs As String      '��ǰʹ����Ȩ�޴�
Private mblnStartUp As Boolean
Private mstrSQL As String
Private mblnOK As Boolean
'Private mblnShowStop As Boolean

'�Զ������/��������###################################################################################################

Public Sub zlRptPrint(ByVal bytMode As Byte)
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
    If rptList.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '�������ݱ��
    If zlReportToVSFlexGrid(vsfPrint, rptList) = False Then Exit Sub
    
    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    
    Set objPrint.Body = vsfPrint
    
    objPrint.Title.Text = "�����¼��Ŀ�嵥"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub RefreshStateInfo()
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�ˢ��״̬����ʾ��Ϣ
    '------------------------------------------------------------------------------------------------------------------
    
    stbThis.Panels(2).Text = "���� " & rptList.Records.Count & " �������¼��Ŀ��"
    
End Sub

Private Function InitGrid() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ���ʼ������ؼ�
    '------------------------------------------------------------------------------------------------------------------
    Dim rptCol As ReportColumn
    
    With rptList
        
        Set rptCol = .Columns.Add(mCol.ͼ��, "", 20, False)
        rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        
        Set rptCol = .Columns.Add(mCol.��Ŀ���, "���", 49, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.��������, "��������", 100, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        
        Set rptCol = .Columns.Add(mCol.��Ŀ����, "����", 100, True): rptCol.Editable = False: rptCol.Groupable = False
        
        Set rptCol = .Columns.Add(mCol.��Ŀ����, "����", 49, False): rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(mCol.��Ŀ����, "����", 49, False): rptCol.Editable = False: rptCol.Groupable = False
        rptCol.Alignment = xtpAlignmentRight
        
        Set rptCol = .Columns.Add(mCol.��ĿС��, "С��", 49, False): rptCol.Editable = False: rptCol.Groupable = False
        rptCol.Alignment = xtpAlignmentRight
        
        Set rptCol = .Columns.Add(mCol.��Ŀ��λ, "��λ", 49, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.��Ŀ��ʾ, "��ʾ", 49, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.��Ŀֵ��, "ֵ��", 160, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.��ͻ���, "��ͻ���", 80, False): rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(mCol.������Ŀ, "����", 49, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.������Ŀ, "����", 49, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.���ò���, "���ò���", 75, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.Ӧ�÷�ʽ, "Ӧ�÷�ʽ", 75, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.��Ŀ����, "��Ŀ����", 75, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.Ӧ�ó���, "Ӧ�ó���", 75, False): rptCol.Editable = False: rptCol.Groupable = False
        
        rptCol.Alignment = xtpAlignmentCenter
        
        Set rptCol = .Columns.Add(mCol.��Ŀid, "��Ŀid", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        
        Set rptCol = .Columns.Add(mCol.˵��, "��Ŀ˵��", 200, True): rptCol.Editable = False: rptCol.Groupable = False
        
        .SetImageList ilsList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GridLineColor = RGB(225, 225, 225)
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
        End With
        .PreviewMode = True
        
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns.Find(mCol.��������)
        .GroupsOrder(0).SortAscending = True
        .SortOrder.Add .Columns.Find(mCol.��Ŀ���)
    End With
    
    InitGrid = True
    
End Function

Private Function CreateToolBox() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    Dim objGrp As TaskPanelGroup
    Dim objItem As TaskPanelGroupItem
    Dim objIlsItem As Object
    
    Call tkp.SetImageList(ilsList)

    Set objGrp = tkp.Groups.Add(0, "Ҫ������")
    objGrp.Expandable = False
    
    Set objItem = objGrp.Items.Add(0, "��  �룺", xtpTaskItemTypeText)
    Call objGrp.Items.Add(0, "��������", xtpTaskItemTypeText)
    Call objGrp.Items.Add(0, "Ӣ������", xtpTaskItemTypeText)
    Call objGrp.Items.Add(0, "��  �ͣ�", xtpTaskItemTypeText)
             
    Set objGrp = tkp.Groups.Add(1, "�ٴ�����")
    objGrp.Expandable = False
    Call objGrp.Items.Add(1, "", xtpTaskItemTypeText)
    
    Call tkp.SetImageList(ilsList)
    Set objGrp = tkp.Groups.Add(2, "��������")
    objGrp.Expandable = False
    Call objGrp.Items.Add(2, "���кţ�", xtpTaskItemTypeText)
    Call objGrp.Items.Add(2, "��¼����", xtpTaskItemTypeText)
    Call objGrp.Items.Add(2, "��¼����", xtpTaskItemTypeText)
    Call objGrp.Items.Add(2, "��¼����", xtpTaskItemTypeText)
    
    'Set objItem = objGrp1.Items.Add(0, rs("��¼��").Value & "(" & rs("��¼��").Value & ")", xtpTaskItemTypeLink, ils16.ListImages("K" & NVL(rs("��¼ɫ"))).Index)

    Call objGrp.Items.Add(2, "��¼ɫ��", xtpTaskItemTypeControl)
    Call objGrp.Items.Add(2, "��Сֵ��", xtpTaskItemTypeText)
    Call objGrp.Items.Add(2, "���ֵ��", xtpTaskItemTypeText)
    Call objGrp.Items.Add(2, "��λֵ��", xtpTaskItemTypeText)
    Call objGrp.Items.Add(2, "����У�", xtpTaskItemTypeText)
    Call objGrp.Items.Add(2, "��¼Ƶ�Σ�", xtpTaskItemTypeText)
      
    Set tkp.Groups(3).Items(5).Control = picColorItem
    
    Set objGrp = tkp.Groups.Add(3, "���ÿ���")
    objGrp.Expandable = False
    Call objGrp.Items.Add(3, "�ڿƣ���ƣ�����ƣ���ƣ�ƨƨ��", xtpTaskItemTypeText)
    
    tkp.Animation = xtpTaskPanelAnimationNo
    tkp.Behaviour = xtpTaskPanelBehaviourExplorer
    tkp.HotTrackStyle = xtpTaskPanelHighlightItem
    
    tkp.SetGroupInnerMargins 0, 1, 1, 1
    
    tkp.AllowDrag = False
    tkp.SelectItemOnFocus = False

    tkp.Groups(1).Expanded = True
    
    
    CreateToolBox = True
    
End Function


Private Function zlMenuClick(ByVal strMenuItem As String, Optional ByVal strParam As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ����ܴ���
    '------------------------------------------------------------------------------------------------------------------
    Dim lngKey As Long
    Dim rptRcd As ReportRecord
    Dim rptItem As ReportRecordItem
    Dim objItem As Object
    Dim strTmp As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand

    
    Select Case strMenuItem
    Case "��ȡ����"
                
        mstrSQL = " SELECT A.��Ŀ���," & _
                          "A.������ As ��������," & _
                          "A.��Ŀ����," & _
                          "Decode(A.��Ŀ����,3,'�߼�',2,'����',1,'����','��ֵ') As ��Ŀ����," & _
                          "A.��Ŀ����," & _
                          "A.��ĿС��," & _
                          "A.��Ŀ��λ,Decode(A.��Ŀ����,1,'�̶���Ŀ','���Ŀ') As ��Ŀ����," & _
                          "Decode(A.��Ŀ��ʾ,1,'����',2,'��ѡ',3,'��ѡ',4,'����',5,'ѡ��','�ı�') As ��Ŀ��ʾ," & _
                          "A.��Ŀֵ��," & _
                          "Decode(A.����ȼ�,1,'һ������',2,'��������',3,'��������','�ؼ�����') As ��ͻ���," & _
                          "Decode(C.��Ŀ���,Null,'','��') As ������Ŀ," & _
                          "A.������Ŀ,Decode(A.���ò���,0,'����',1,'����',2,'Ӥ��') As ���ò���,Decode(A.Ӧ�÷�ʽ,0,'����ʹ��',1,'����ʹ��',2,'����������','') As Ӧ�÷�ʽ," & _
                          "Decode(A.Ӧ�ó���,1,'���µ�',2,'��¼��','ͨ��') As Ӧ�ó���," & _
                          "A.��Ŀid,A.˵�� " & _
                     "FROM �����¼��Ŀ A,���¼�¼��Ŀ C WHERE C.��Ŀ���(+)=A.��Ŀ��� Order By A.������,A.��Ŀ���"
        
        Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption)
        If rs.BOF = False Then
            rptList.Records.DeleteAll
            
            Do While Not rs.EOF
                
                Set rptRcd = rptList.Records.Add()

                Set rptItem = rptRcd.AddItem("")
                rptItem.Icon = IIf(Val(NVL(rs("��Ŀid"))) > 0, 1, 0)
                
                rptRcd.AddItem Zero(NVL(rs("��Ŀ���")))
                rptRcd.AddItem NVL(rs("��������"))
                rptRcd.AddItem NVL(rs("��Ŀ����"))
                rptRcd.AddItem NVL(rs("��Ŀ����"))
                
                rptRcd.AddItem Zero(NVL(rs("��Ŀ����")))
                rptRcd.AddItem Zero(NVL(rs("��ĿС��")))
                rptRcd.AddItem NVL(rs("��Ŀ��λ"))
                rptRcd.AddItem NVL(rs("��Ŀ��ʾ"))
                rptRcd.AddItem NVL(rs("��Ŀֵ��"))
                rptRcd.AddItem NVL(rs("��ͻ���"))
                rptRcd.AddItem NVL(rs("������Ŀ"))
                rptRcd.AddItem IIf(NVL(rs("������Ŀ")) = 1, "��", "")
                rptRcd.AddItem NVL(rs("���ò���"))
                rptRcd.AddItem NVL(rs("Ӧ�÷�ʽ"))
                rptRcd.AddItem NVL(rs("��Ŀ����"))
                rptRcd.AddItem NVL(rs("Ӧ�ó���"))
                rptRcd.AddItem Zero(NVL(rs("��Ŀid")))
                rptRcd.AddItem NVL(rs("˵��"))
                rs.MoveNext
            Loop
            
            rptList.Populate
            
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡҪ��"
               
        tkp.Groups(1).Items(1).Caption = "��  �룺"
        tkp.Groups(1).Items(2).Caption = "��������"
        tkp.Groups(1).Items(3).Caption = "Ӣ������"
        tkp.Groups(1).Items(4).Caption = "��  �ͣ�"
        
        tkp.Groups(2).Items(1).Caption = ""
        
        If Not (rptList.FocusedRow Is Nothing) Then
            If Not (rptList.FocusedRow.Record Is Nothing) Then lngKey = Val(rptList.FocusedRow.Record.Item(mCol.��Ŀid).Value)
        End If
                
        mstrSQL = "Select ����,������,Ӣ����,Decode(����,1,'����',2,'����',3,'�߼�','��ֵ') As ����,�ٴ����� From ����������Ŀ Where ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, lngKey)
        If rs.BOF = False Then
            tkp.Groups(1).Items(1).Caption = "��  �룺" & zlCommFun.NVL(rs("����"))
            tkp.Groups(1).Items(2).Caption = "��������" & zlCommFun.NVL(rs("������"))
            tkp.Groups(1).Items(3).Caption = "Ӣ������" & zlCommFun.NVL(rs("Ӣ����"))
            tkp.Groups(1).Items(4).Caption = "��  �ͣ�" & zlCommFun.NVL(rs("����"))
            tkp.Groups(2).Items(1).Caption = zlCommFun.NVL(rs("�ٴ�����"))
                        
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ���ÿ���"
        
        If Not (rptList.FocusedRow Is Nothing) Then
            If Not (rptList.FocusedRow.Record Is Nothing) Then lngKey = Val(rptList.FocusedRow.Record.Item(mCol.��Ŀ���).Value)
        End If
        
        tkp.Groups(4).Items(1).Caption = " "
        mstrSQL = "Select ���ÿ��� From �����¼��Ŀ Where ��Ŀ���=[1]"
        Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, lngKey)
        If rs.BOF = False Then
            Select Case zlCommFun.NVL(rs("���ÿ���"), 0)
            Case 0
                tkp.Groups(4).Items(1).Caption = "����Ŀ��ʱ��ʹ��"
            Case 1
                tkp.Groups(4).Items(1).Caption = "����ĿȫԺͨ��"
            Case 2
                mstrSQL = "Select b.���� From �������ÿ��� a,���ű� b Where a.��Ŀ���=[1] And a.����id=b.ID"
                Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, lngKey)
                If rs.BOF = False Then
                    strTmp = ""
                    Do While Not rs.EOF
                        strTmp = strTmp & "��" & zlCommFun.NVL(rs("����"))
                        rs.MoveNext
                    Loop
                    If strTmp <> "" Then strTmp = Mid(strTmp, 2)
                    tkp.Groups(4).Items(1).Caption = strTmp
                End If
                
            End Select
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ����"
    
        tkp.Groups(3).Items(1).Caption = "���кţ�"
        tkp.Groups(3).Items(2).Caption = "��¼����"
        tkp.Groups(3).Items(3).Caption = "��¼����"
        tkp.Groups(3).Items(4).Caption = "��¼����"
        tkp.Groups(3).Items(5).Caption = "��¼ɫ��"
        tkp.Groups(3).Items(6).Caption = "��Сֵ��"
        tkp.Groups(3).Items(7).Caption = "���ֵ��"
        tkp.Groups(3).Items(8).Caption = "��λֵ��"
        tkp.Groups(3).Items(9).Caption = "����У�"
        tkp.Groups(3).Items(10).Caption = "��¼Ƶ�Σ�"
        lblColor.BackStyle = 0
        
        If Not (rptList.FocusedRow Is Nothing) Then
            If Not (rptList.FocusedRow.Record Is Nothing) Then lngKey = Val(rptList.FocusedRow.Record.Item(mCol.��Ŀ���).Value)
        End If
        
        mstrSQL = "Select �������,��¼��,Decode(��¼��,1,'����',2,'���') As ��¼��,��¼��,��¼ɫ,��Сֵ,���ֵ,��λֵ,�����,��¼Ƶ�� From ���¼�¼��Ŀ Where ��Ŀ���=[1]"
        Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, lngKey)
        If rs.BOF = False Then
        
            tkp.Groups(3).Items(1).Caption = "���кţ�" & zlCommFun.NVL(rs("�������"))
            tkp.Groups(3).Items(2).Caption = "��¼����" & zlCommFun.NVL(rs("��¼��"))
            tkp.Groups(3).Items(3).Caption = "��¼����" & zlCommFun.NVL(rs("��¼��"))
            
            If lngKey = 1 Then
                tkp.Groups(3).Items(4).Caption = "��¼����" & zlCommFun.NVL(rs("��¼��").Value, "��,��,��")
            Else
                tkp.Groups(3).Items(4).Caption = "��¼����" & zlCommFun.NVL(rs("��¼��").Value)
            End If
            
            '������ɫ
            On Error Resume Next
            Set objItem = Nothing
            Set objItem = ilsList.ListImages("K" & NVL(rs("��¼ɫ"), 0))
            If objItem Is Nothing Then Call SetColorIcon(Me, "K" & NVL(rs("��¼ɫ"), 0), NVL(rs("��¼ɫ"), 0), ilsList)
            On Error GoTo 0
            
            
            tkp.Groups(3).Items(5).Caption = "��¼ɫ��" & zlCommFun.NVL(rs("��¼ɫ"))
'            If zlCommFun.NVL(rs("��¼ɫ"), -1) = -1 Then
'                lblColor.BackStyle = 0
'            Else
                lblColor.BackStyle = 1
                lblColor.BackColor = zlCommFun.NVL(rs("��¼ɫ"), 0)
'            End If

            tkp.Groups(3).Items(6).Caption = "��Сֵ��" & zlCommFun.NVL(rs("��Сֵ"))
                        
            tkp.Groups(3).Items(7).Caption = "���ֵ��" & zlCommFun.NVL(rs("���ֵ"))
            tkp.Groups(3).Items(8).Caption = "��λֵ��" & Format(zlCommFun.NVL(rs("��λֵ")), "0.0")
            tkp.Groups(3).Items(9).Caption = "����У�" & zlCommFun.NVL(rs("�����"))
            tkp.Groups(3).Items(10).Caption = "��¼Ƶ�Σ�" & zlCommFun.NVL(rs("��¼Ƶ��"))
            
        End If
    End Select
    
    cbsThis.RecalcLayout
    Call RefreshStateInfo
    
    zlMenuClick = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Public Function EditRefresh(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ���������/�޸ĺ��������Դ���
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long

    rptList.Records.DeleteAll
    
    Call zlMenuClick("��ȡ����")
    
    '�ָ�
    rptList.Populate
    
    For lngLoop = 0 To rptList.Rows.Count - 1
        If Not (rptList.Rows(lngLoop).Record Is Nothing) Then
            If Val(rptList.Rows(lngLoop).Record.Item(mCol.��Ŀ���).Value) = lngKey Then
                Set rptList.FocusedRow = rptList.Rows(lngLoop)
                Call rptList_SelectionChanged
                Exit For
            End If
        End If
    Next

End Function

Private Function InitMenuBar() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ���ʼ���˵���������
    '------------------------------------------------------------------------------------------------------------------
    Dim cbrMenuBar As Object
    Dim obj As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵���"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)")
        cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyTo, "���ÿ���(&T)"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "��������(&J)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Compend, "�������(&Q)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "�����ص�(&K)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Request, "����ģ��(&T)")
        
        '�°滤ʿ����վ��������
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CollectMan, "������Ŀ(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_AnimalPart, "���²�λ(&P)"): cbrControl.IconId = 2612
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "��¼Ƶ��(&L)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_WavyMan, "������Ŀ(&B)")
        '47964:������,2013-01-21,�����������ͬ�����ù���
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_WaveSynchro, "����ͬ��(&S)")
    End With

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        
'        Set cbrControl = .Add(xtpControlButton, conMenu_View_Show, "��ʾͣ��(&A)"): cbrControl.BeginGroup = True: cbrControl.IconId = 1
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False  '����
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)..."): cbrControl.BeginGroup = True
    End With
    
     '�����
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    

    '����������
    Set cbrToolBar = cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����")
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyTo, "���ÿ���"): cbrControl.BeginGroup = True
                
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����")
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.STYLE = xtpButtonIconAndCaption
    Next
    
    '��ȡ��������ģ��ı���:��Ϊ��һ���Զ�ȡ,ȫ�ֱ�������
    '---------------------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
End Function

'�ؼ��¼�##############################################################################################################

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strKey As String
    Dim lngLoop As Long
    Dim strSQL() As String
    Dim blnTran As Boolean
    Dim lngIndex As Long
    Dim cbrControl As Object
    Dim lngKey As Long
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
        
    Select Case Control.ID
        Case conMenu_File_PrintSet
            
            Call zlPrintSet
                    
        Case conMenu_File_Preview
            
            Call zlRptPrint(2)
        
        Case conMenu_File_Print
        
            Call zlRptPrint(1)
        
        Case conMenu_File_Excel
        
            Call zlRptPrint(3)
    
        Case conMenu_View_ToolBar_Button
        
            cbsThis(2).Visible = Not cbsThis(2).Visible
            cbsThis.RecalcLayout
        
        Case conMenu_View_ToolBar_Text
        
            For Each cbrControl In cbsThis(2).Controls
                cbrControl.STYLE = IIf(cbrControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            
            cbsThis.RecalcLayout
            
        Case conMenu_View_StatusBar
        
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
            
        Case conMenu_Edit_NewItem
            '������Ŀ
            
            lngKey = 0
            
            
            If Not (rptList.FocusedRow Is Nothing) Then
                If Not (rptList.FocusedRow.Record Is Nothing) Then lngKey = Val(rptList.FocusedRow.Record.Item(mCol.��Ŀ���).Value)
            End If
                
            If frmTendEdit.ShowEdit(Me, 0, lngKey) Then
                mblnOK = True
                rptList.SetFocus
            End If
    
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Modify
            
            If rptList.FocusedRow Is Nothing Then Exit Sub
            If rptList.FocusedRow.Record Is Nothing Then Exit Sub
            
            '�޸���Ŀ
            If frmTendEdit.ShowEdit(Me, Val(rptList.FocusedRow.Record.Item(mCol.��Ŀ���).Value)) Then
                mblnOK = True
                rptList.SetFocus
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Delete
            'ɾ����Ŀ
            If rptList.FocusedRow Is Nothing Then Exit Sub
            If rptList.FocusedRow.Record Is Nothing Then Exit Sub
            If CheckItemExistData(3, Val(rptList.FocusedRow.Record(mCol.��Ŀ���).Value), rptList.FocusedRow.Record(mCol.��Ŀ����).Value) = True Then Exit Sub
            If MsgBox("�����Ҫɾ����" & rptList.FocusedRow.Record(mCol.��Ŀ����).Value & "����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            strSQL(ReDimArray(strSQL)) = "ZL_�����¼��Ŀ_DELETE(" & Val(rptList.FocusedRow.Record(mCol.��Ŀ���).Value) & ")"
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_CollectMan    '������Ŀ����
            On Error Resume Next
            frmCollectMan.Show 1, Me
        
        Case conMenu_Edit_AnimalPart    '���²�λ����
            frmAnimalPartMan.Show 1, Me
            
        Case conMenu_Edit_Reuse         '��¼Ƶ�ι���
            frmItemRecordMan.Show 1, Me
        Case conMenu_Edit_WavyMan  '������Ŀ����
            frmItemWaveMan.Show 1, Me
        '47964:������,2013-01-21,�����������ͬ�����ù���
        Case conMenu_Edit_WaveSynchro '����ͬ������
            FrmTendWaveDataSet.Show 1, Me
        Case conMenu_Edit_ApplyTo
            If rptList.FocusedRow Is Nothing Then Exit Sub
            If rptList.FocusedRow.Record Is Nothing Then Exit Sub
            
            If frmTendItemDept.ShowMe(Me, Val(rptList.FocusedRow.Record.Item(mCol.��Ŀ���).Value)) Then
                Call rptList_SelectionChanged
            End If
            
'        Case conMenu_Edit_Stop
'            'ͣ����Ŀ
'            If rptList.FocusedRow Is Nothing Then Exit Sub
'            If rptList.FocusedRow.Record Is Nothing Then Exit Sub
'
'            If MsgBox("�����Ҫͣ�á�" & rptList.FocusedRow.Record(mCol.��Ŀ����).Value & "����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
'            strSQL(ReDimArray(strSQL)) = "ZL_�����¼��Ŀ_Stop(" & Val(rptList.FocusedRow.Record(mCol.��Ŀ���).Value) & ")"
'
'        Case conMenu_Edit_Reuse
'            '������Ŀ
'            If rptList.FocusedRow Is Nothing Then Exit Sub
'            If rptList.FocusedRow.Record Is Nothing Then Exit Sub
'
'            If MsgBox("�����Ҫ���á�" & rptList.FocusedRow.Record(mCol.��Ŀ����).Value & "����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
'            strSQL(ReDimArray(strSQL)) = "ZL_�����¼��Ŀ_Reuse(" & Val(rptList.FocusedRow.Record(mCol.��Ŀ���).Value) & ")"
        
        Case conMenu_Edit_Adjust
            
            If frmTendBodyArrage.ShowEdit(Me) Then
                Call rptList_SelectionChanged
            End If
        
        Case conMenu_Edit_Compend
            
            '����ҩƷ���������ϵ
            If frmTendDrink.ShowEdit(Me) Then
                Call rptList_SelectionChanged
            End If
            
        Case conMenu_Edit_MarkMap

            Call frmTendBlanket.ShowEdit(Me, mstrPrivs)
        
        Case conMenu_Edit_Request   '������Ŀģ��
            Call frmTendItemTemplate.ShowMe(Me, mstrPrivs)
        
        Case conMenu_View_Refresh
                            
            '����
            If Not (rptList.FocusedRow Is Nothing) Then
                If Not (rptList.FocusedRow.Record Is Nothing) Then strKey = Val(rptList.FocusedRow.Record(mCol.��Ŀ���).Value)
            End If

            rptList.Records.DeleteAll
            
            Call zlMenuClick("��ȡ����")

            '�ָ�
            For lngLoop = 0 To rptList.Rows.Count - 1
                If Not (rptList.Rows(lngLoop).Record Is Nothing) Then
                    If Val(rptList.Rows(lngLoop).Record.Item(mCol.��Ŀ���).Value) = Val(strKey) Then
                        Set rptList.FocusedRow = rptList.Rows(lngLoop)
                        Call rptList_SelectionChanged
                        Exit For
                    End If
                End If
            Next
            
        Case conMenu_Help_Help
        
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        
        Case conMenu_Help_About
            
            Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
            
        Case conMenu_Help_Web_Home
            
            Call zlHomePage(Me.hWnd)
            
        Case conMenu_Help_Web_Forum '������̳
            Call zlWebForum(Me.hWnd)

        Case conMenu_Help_Web_Mail
            
            Call zlMailTo(Me.hWnd)
            
        Case conMenu_File_Exit
            Unload Me
            Exit Sub
        Case Else
            'ִ�з�������ǰģ��ı���
'            Dim lng��Ŀ��� As Long, str��Ŀ���� As String
'            If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
'                If rptList.SelectedRows.Count > 0 Then
'                    If Not rptList.SelectedRows(0).GroupRow Then
'                        lng��Ŀ��� = Val(rptList.SelectedRows(0).Record(mCol.��Ŀ���).Value)
'                        str��Ŀ���� = rptList.SelectedRows(0).Record(mCol.��Ŀ����).Value
'                    End If
'                End If
'                If str��Ŀ���� <> "" Then
'                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "��Ŀ���=" & lng��Ŀ���, "��Ŀ����=" & str��Ŀ����)
'                Else
'                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
'                End If
'            End If
            Exit Sub
    End Select
    
    blnTran = True
    gcnOracle.BeginTrans
    
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    
    gcnOracle.CommitTrans
    blnTran = False
    
    Select Case Control.ID
        Case conMenu_Edit_Delete
            'ɾ����
            
            lngIndex = rptList.FocusedRow.Index
            rptList.Records.RemoveAt (rptList.FocusedRow.Record.Index)
            rptList.Populate
            
            If rptList.Records.Count > 0 Then
                lngIndex = IIf(rptList.Records.Count - 1 > lngIndex, lngIndex, rptList.Records.Count - 1)
                rptList.Rows(lngIndex).Selected = True
                Set rptList.FocusedRow = rptList.Rows(lngIndex)
            End If
            rptList.SetFocus
            Call rptList_SelectionChanged
            mblnOK = True
            
'        Case conMenu_Edit_Stop
'            '��д����ʱ���ɾ������
'
'            If mblnShowStop Then
'                rptList.FocusedRow.Record(mCol.����ʱ��).Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
'            Else
'                lngIndex = rptList.FocusedRow.Index
'                rptList.Records.RemoveAt (rptList.FocusedRow.Record.Index)
'                rptList.Populate
'
'                If rptList.Records.Count > 0 Then
'                    lngIndex = IIf(rptList.Records.Count - 1 > lngIndex, lngIndex, rptList.Records.Count - 1)
'                    rptList.Rows(lngIndex).Selected = True
'                    Set rptList.FocusedRow = rptList.Rows(lngIndex)
'                End If
'                rptList.SetFocus
'                Call rptList_SelectionChanged
'
'            End If
'
'        Case conMenu_Edit_Reuse
'            '���ĳ���ʱ��Ϊ��
'            If rptList.FocusedRow Is Nothing Then Exit Sub
'            If rptList.FocusedRow.Record Is Nothing Then Exit Sub
'
'            rptList.FocusedRow.Record(mCol.����ʱ��).Value = ""
    End Select
    
    cbsThis.RecalcLayout
    Call RefreshStateInfo
    
    Exit Sub
    
errHand:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)

    If stbThis.Visible Then Bottom = stbThis.Height
    
End Sub

Private Sub cbsThis_Resize()
    
    On Error Resume Next
    
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long  '�ͻ�����Ĵ�С

    Call cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    
    With rptList
        .Left = lngLeft
        .Width = lngRight - lngLeft - tkp.Width - 45
        .Top = lngTop
        .Height = lngBottom - lngTop
    End With
    
    With tkp
        .Left = rptList.Left + rptList.Width + 45
        .Top = rptList.Top
        .Height = rptList.Height
    End With
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup
        End Select
    End If
    
    Err = 0: On Error Resume Next
    
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (rptList.Records.Count > 0)
    Case conMenu_Edit_NewItem, conMenu_Edit_Adjust
        Control.Enabled = (InStr(1, mstrPrivs, "��ɾ��") > 0)
    Case conMenu_Edit_Request
        Control.Enabled = (InStr(1, mstrPrivs, "����ģ��") > 0)
    Case conMenu_Edit_CollectMan
        Control.Enabled = (InStr(1, mstrPrivs, "������Ŀ") > 0)
    Case conMenu_Edit_AnimalPart
        Control.Enabled = (InStr(1, mstrPrivs, "���²�λ") > 0)
    Case conMenu_Edit_WavyMan
        Control.Enabled = (InStr(1, mstrPrivs, "��������Ŀ") > 0)
    '47964:������,2013-01-21,�����������ͬ�����ù���
    Case conMenu_Edit_WaveSynchro '����ͬ������
        Control.Enabled = (InStr(1, mstrPrivs, "����ͬ����Ŀ") > 0)
    Case conMenu_Edit_Reuse
        Control.Enabled = (InStr(1, mstrPrivs, "��ɾ��") > 0)
    Case conMenu_Edit_Modify
        If Not (rptList.FocusedRow Is Nothing) Then
            If Not (rptList.FocusedRow.Record Is Nothing) Then
                Control.Enabled = (InStr(1, mstrPrivs, "��ɾ��") > 0)
            Else
                Control.Enabled = False
            End If
        Else
            Control.Enabled = False
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete
        If Not (rptList.FocusedRow Is Nothing) Then
            If Not (rptList.FocusedRow.Record Is Nothing) Then
                Control.Enabled = (InStr(1, mstrPrivs, "��ɾ��") > 0)
            Else
                Control.Enabled = False
            End If
        Else
            Control.Enabled = False
        End If
        If Control.Enabled Then Control.Enabled = (rptList.FocusedRow.Record.Item(mCol.������Ŀ).Value <> "��")
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_ApplyTo

        Control.Visible = (InStr(1, mstrPrivs, "��ɾ��") > 0)

        If Not (rptList.FocusedRow Is Nothing) Then
            If Not (rptList.FocusedRow.Record Is Nothing) Then
                Control.Enabled = (Control.Visible And rptList.FocusedRow.Record.Item(mCol.��Ŀ���).Value > 2)
            Else
                Control.Enabled = False
            End If
        Else
            Control.Enabled = False
        End If
    Case conMenu_View_ToolBar_Button
        Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text
        Control.Checked = Not (Me.cbsThis(2).Controls(1).STYLE = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size
        Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar
        Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub Form_Load()
    mstrPrivs = gstrPrivs
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    
    mblnStartUp = True
'    mblnShowStop = False
    
    Call InitCommonControls
        
    Call InitMenuBar
    
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitGrid
    Call CreateToolBox
    
    mblnStartUp = False
    
    Call zlMenuClick("��ȡ����")
    
    On Error Resume Next
    
    If rptList.Records.Count > 0 Then Set rptList.FocusedRow = rptList.Rows(0)
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call SaveWinState(Me, App.ProductName)
    
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Not (rptList.FocusedRow Is Nothing) Then
            Call rptList_RowDblClick(rptList.FocusedRow, rptList.FocusedRow.Record.Item(mCol.��Ŀ����))
        End If
    End If
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As Object
    
    If Button <> 2 Then Exit Sub
    
    If cbsThis.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = cbsThis.Add("�����˵�", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
    Next
    cbrPopupBar.ShowPopup
End Sub

Private Sub rptList_RowDblClick(ByVal ROW As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    
    If Not (rptList.FocusedRow Is Nothing) Then
          Call cbsThis_Execute(cbsThis.FindControl(, conMenu_Edit_Modify))
    End If
End Sub

Private Sub rptList_SelectionChanged()

    If mblnStartUp Then Exit Sub
        
    Call zlMenuClick("��ȡҪ��")
    Call zlMenuClick("��ȡ����")
    Call zlMenuClick("��ȡ���ÿ���")
        
End Sub


Public Function CheckItemExistData(ByVal bytType As Byte, ParamArray arrInput() As Variant) As Boolean
'����:����Ӧ�Ļ�����Ŀ�Ƿ������µ����¼���Ѿ���������
'bytType:1��ֻ������Ŀ�Ƿ��Ѿ������˻���ҵ�����ݡ�2��ֻ������Ŀ�Ƿ��Ѿ��󶨻����¼��:��������1��2
    Dim rsTemp As New ADODB.Recordset
    Dim strInfo As String
    Dim strSQL1 As String, strSQL2 As String
    On Error GoTo errHand
    CheckItemExistData = True
    strSQL1 = "Select Id" & vbNewLine & _
        " From (Select Id" & vbNewLine & _
        "       From ���˻�����ϸ" & vbNewLine & _
        "       Where ��Ŀ��� = [1] And Rownum < 2" & vbNewLine & _
        "       Union All" & vbNewLine & _
        "       Select Id" & vbNewLine & _
        "       From ���˻�������" & vbNewLine & _
        "       Where ��Ŀ��� = [1] And Rownum < 2)"
    strSQL2 = " Select a.���� " & vbNewLine & _
        " From �����ļ��ṹ d, �����ļ��ṹ p, �����ļ��б� a" & vbNewLine & _
        " Where p.Id = d.��id And p.�������� = 1 And p.�����ı� = '���м���' And d.Ҫ������ = [1] And p.�ļ�id = a.Id And a.���� = 3 And ���� <> -1 And Rownum <2"
    
    If bytType = 1 Then
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL1, "�����Ŀ�Ƿ��Ѿ����ڻ�������", Val(arrInput(0)))
        If rsTemp.RecordCount > 0 Then
            MsgBox "����Ŀ�Ѿ������˻������ݣ�������ɾ�����޸ģ�", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf bytType = 2 Then
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL2, "�����Ŀ�Ƿ��Ѱ󶨻����¼��", CStr(arrInput(0)))
        If rsTemp.RecordCount > 0 Then
            MsgBox "����Ŀ���뻤���ļ��󶨣����������ɾ�����޸����ƣ�", vbInformation, gstrSysName
            Exit Function
        End If
    Else
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL1, "�����Ŀ�Ƿ��Ѿ����ڻ�������", Val(arrInput(0)))
        If rsTemp.RecordCount > 0 Then
            MsgBox "����Ŀ�Ѿ������˻������ݣ�������ɾ����", vbInformation, gstrSysName
            Exit Function
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL2, "�����Ŀ�Ƿ��Ѱ󶨻����¼��", CStr(arrInput(1)))
        If rsTemp.RecordCount > 0 Then
            MsgBox "����Ŀ���뻤���ļ��󶨣����������ɾ�����޸����ƣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CheckItemExistData = False
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


