VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmTendItemTemplate 
   Caption         =   "������Ŀģ��"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8685
   Icon            =   "frmTendItemTemplate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   8685
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ImageList imgrpt 
      Left            =   4050
      Top             =   2580
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendItemTemplate.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendItemTemplate.frx":D0B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picDetail 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3195
      Left            =   5820
      ScaleHeight     =   3195
      ScaleWidth      =   2505
      TabIndex        =   2
      Top             =   930
      Width           =   2505
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   3105
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   2400
         _cx             =   4233
         _cy             =   5477
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
         BackColorFixed  =   15790320
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
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
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
   End
   Begin VB.PictureBox picHead 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3945
      Left            =   480
      ScaleHeight     =   3945
      ScaleWidth      =   4845
      TabIndex        =   1
      Top             =   750
      Width           =   4845
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   2040
         Left            =   0
         TabIndex        =   4
         Top             =   0
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
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5340
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTendItemTemplate.frx":13916
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10239
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
      DesignerControls=   "frmTendItemTemplate.frx":141A8
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
Attribute VB_Name = "frmTendItemTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���ڼ���������########################################################################################################
Private Enum mHeadCol
    ͼ��
    ģ������
    ���û���ȼ�
    ����ȼ�
    ����
    ����ID
End Enum
Private Enum mDetailCol
    ���
    ��Ŀ����
End Enum

Private mstrSel As String           '��¼��ǰѡ����Ŀ����Ϣ,��������,�޸ĺ�λ,�Ҳ�����λ�������Ŀ��;���Ϊ��,��ʾ��λ�ĵ�һ����Ŀ��
Private mlng����ID As Long          '��ǰ����Ա������ȱʡ����ID
Private mstr����ID As String        '��ǰ����Ա��������ID
Private mstrPrivs As String         '��ǰʹ����Ȩ�޴�
Private mblnStartUp As Boolean
Private mstrSQL As String
Private mstrDeptID As String        '��ǰ����Ա��������ID

'�Զ������/��������###################################################################################################

Public Sub ShowME(ByVal objParent As Object, ByVal strPrivs As String)
    mstrPrivs = strPrivs
    Me.Show 1, objParent
End Sub

Public Sub zlRptPrint(ByVal bytMode As Byte)
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
    If rptList.FocusedRow Is Nothing Then Exit Sub
    If rptList.FocusedRow.Record Is Nothing Then Exit Sub

    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow

    Set objPrint.Body = vsfDetail

    objPrint.Title.Text = rptList.FocusedRow.Record.Item(ģ������).Value
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("����:" & rptList.FocusedRow.Record.Item(����).Value)
    Call objPrint.UnderAppRows.Add(objAppRow)
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("���û���ȼ�:" & rptList.FocusedRow.Record.Item(���û���ȼ�).Value)
    Call objPrint.UnderAppRows.Add(objAppRow)
    
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
    stbThis.Panels(2).Text = "���� " & vsfDetail.Rows - 1 & " �������¼��Ŀ��"
End Sub

Private Function zlMenuClick(ByVal strMenuItem As String, Optional ByVal strParam As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ����ܴ���
    '------------------------------------------------------------------------------------------------------------------
    Dim arrData
    Dim blnSel As Boolean
    Dim str����ȼ� As String
    Dim intRow As Integer, intRows As Integer
    Dim rptRecord As ReportRecord, rptRecordItem As ReportRecordItem
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand

    Select Case strMenuItem
    Case "��ʼ��"
        '��ȡ��ǰ����Ա�����Ŀ���ID
        gstrSQL = " Select B.ID,B.����,B.���� " & _
                  " From ��������˵�� A,���ű� B,������Ա C" & _
                  " Where A.��������='�ٴ�' And A.������� IN (2,3) And A.����ID=B.ID" & _
                  " And B.ID=C.����ID And C.��ԱID=[1]" & _
                  " UNION " & _
                  " Select B.ID,B.����,B.���� " & _
                  " From ��������˵�� A,���ű� B,�������Ҷ�Ӧ C" & _
                  " Where A.��������='�ٴ�' And A.������� IN (2,3) And A.����ID=B.ID And B.ID=C.����ID And C.����ID=[2]"
        gstrSQL = " Select Distinct ID,����,���� From (" & gstrSQL & ") Order by ����"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngUserId, glngDeptId)
        With rsTemp
            mstr����ID = ""
            Do While Not .EOF
                mstr����ID = mstr����ID & "," & !ID
                If mlng����ID = 0 Then mlng����ID = !ID
                .MoveNext
            Loop
        End With
        
    Case "��ȡ����"
        Call InitGird
        rptList.Records.DeleteAll
        
        '��ȡ����ģ��
        mstrSQL = " Select Distinct B.ID,NVL(B.����,'') AS ����,ģ������,����ȼ� " & _
                  " From ������Ŀģ�� A,���ű� B" & _
                  " Where A.����ID=B.ID(+)" & _
                  " Order by NVL(B.����,''),Decode(����ȼ�,-1,5,����ȼ�) "
        Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption)
        If rsTemp.BOF = False Then
            Do While Not rsTemp.EOF
                Set rptRecord = rptList.Records.Add()
                Set rptRecordItem = rptRecord.AddItem("")
                rptRecordItem.Icon = IIf(Val(NVL(rsTemp!ID)) > 0, 1, 0)
                rptRecord.AddItem rsTemp.Fields("ģ������").Value
                
                Select Case rsTemp!����ȼ�
                Case -1
                    str����ȼ� = "����¼��ģ��"
                Case 0
                    str����ȼ� = "�ؼ�����¼��ģ��"
                Case 1
                    str����ȼ� = "һ������¼��ģ��"
                Case 2
                    str����ȼ� = "��������¼��ģ��"
                Case 3
                    str����ȼ� = "��������¼��ģ��"
                End Select
                rptRecord.AddItem str����ȼ�
                rptRecord.AddItem rsTemp.Fields("����ȼ�").Value
                rptRecord.AddItem rsTemp.Fields("����").Value
                rptRecord.AddItem NVL(rsTemp.Fields("ID").Value, 0)
                
                rsTemp.MoveNext
            Loop
        End If
        rptList.Populate
        
        '��λ��Ŀ
        On Error Resume Next
        If mstrSel = "" Then
            If rptList.Rows.Count > 0 Then Set rptList.FocusedRow = rptList.Rows(1)
        Else
            arrData = Split(mstrSel, "|")
            intRows = rptList.Rows.Count
            For intRow = 1 To intRows
                If Not rptList.Rows(intRow - 1).Record Is Nothing Then
                    If Val(rptList.Rows(intRow - 1).Record.Item(����ȼ�).Value) = arrData(0) And Val(rptList.Rows(intRow - 1).Record.Item(����ID).Value) = arrData(1) Then
                        blnSel = True
                        Set rptList.FocusedRow = rptList.Rows(intRow - 1)
                        Exit For
                    End If
                End If
            Next
            If blnSel = False Then
                'û�ҵ�,����ɾ������,ֱ�Ӷ�λ�����һ����¼
                If Val(arrData(2)) <= rptList.Rows.Count Then
                    On Error Resume Next
                    rptList.Rows(arrData(2)).Selected = True
                Else
                    '˵��ɾ���������һ��,��λ�����һ����
                    If rptList.Rows.Count > 0 Then Set rptList.FocusedRow = rptList.Rows(rptList.Rows.Count).Selected
                End If
            End If
        End If
        
    Case "��ȡģ������"
        Call InitGird
        
        If rptList.FocusedRow Is Nothing Then Exit Function
        If rptList.FocusedRow.Record Is Nothing Then Exit Function
        gstrSQL = " Select B.��Ŀ���,B.��Ŀ���� From ������Ŀģ�� A,�����¼��Ŀ B " & _
                  " Where A.��Ŀ���=B.��Ŀ��� And A.����ID =[1] And A.����ȼ�=[2]" & _
                  " Order by A.�������"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(rptList.FocusedRow.Record.Item(����ID).Value), CInt(rptList.FocusedRow.Record.Item(����ȼ�).Value))
        
        With rsTemp
            Do While Not .EOF
                If .AbsolutePosition > vsfDetail.Rows - 1 Then vsfDetail.Rows = vsfDetail.Rows + 1
                vsfDetail.TextMatrix(.AbsolutePosition, ���) = .Fields("��Ŀ���").Value
                vsfDetail.TextMatrix(.AbsolutePosition, ��Ŀ����) = .Fields("��Ŀ����").Value
                rsTemp.MoveNext
            Loop
        End With
        Call RefreshStateInfo
    End Select
    '------------------------------------------------------------------------------------------------------------------

    cbsThis.RecalcLayout
    Call RefreshStateInfo

    zlMenuClick = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitRpt(Optional ByVal intState As Integer = 0)
    '0��ʾ�������ˢ��;1-ֻˢ����ϸ��
    Dim rptCol As ReportColumn
    
    With rptList
        Set rptCol = .Columns.Add(ͼ��, "", 20, False)
        rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        
        Set rptCol = .Columns.Add(ģ������, "ģ������", 250, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(���û���ȼ�, "���û���ȼ�", 111, True): rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(����ȼ�, "����ȼ�", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(����, "����", 60, True): rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(����ID, "����ID", 0, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        
        .SetImageList imgrpt
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
        .GroupsOrder.Add .Columns.Find(����)
        .GroupsOrder(0).SortAscending = True
        .SortOrder.Add .Columns.Find(����ȼ�)
    End With
    
End Sub

Private Sub InitGird()
    With vsfDetail
        .Clear
        .Rows = 2: .Cols = 2
        .TextMatrix(0, ���) = "���"
        .TextMatrix(0, ��Ŀ����) = "��Ŀ����"
        .ColWidth(���) = 800
        .ColWidth(��Ŀ����) = 2000
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
    End With
End Sub

Private Function InitMenuBar() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ���ʼ���˵���������
    '------------------------------------------------------------------------------------------------------------------
    Dim cbrMenuBar As Object
    Dim obj As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrToolBar As CommandBar
    Dim objExtendedBar As CommandBar
    
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
    End With

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        
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
               
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����")
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '��ȡ��������ģ��ı���:��Ϊ��һ���Զ�ȡ,ȫ�ֱ�������
    '---------------------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
End Function

Private Sub SetDockRight(BarToDock As CommandBar, BarOnLeft As CommandBar)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    cbsThis.RecalcLayout
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    
    cbsThis.DockToolBar BarToDock, Right, (Bottom + Top) / 2, BarOnLeft.Position

End Sub

'�ؼ��¼�##############################################################################################################

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strKey As String
    Dim lngLoop As Long
    Dim lngIndex As Long
    Dim cbrControl As Object

    On Error GoTo errHand

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
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next

            cbsThis.RecalcLayout

        Case conMenu_View_StatusBar

            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout

        Case conMenu_Edit_NewItem
            '������Ŀ
            frmTendItemTemplateEdit.mstrPrivs = mstrPrivs
            strKey = frmTendItemTemplateEdit.ShowEditor(Me, mlng����ID, "", 9)
            If strKey = "" Then Exit Sub
            mstrSel = strKey
            Call zlMenuClick("��ȡ����")
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Modify

            If rptList.FocusedRow Is Nothing Then Exit Sub
            If rptList.FocusedRow.Record Is Nothing Then Exit Sub

            '�޸���Ŀ
            frmTendItemTemplateEdit.mstrPrivs = mstrPrivs
            If frmTendItemTemplateEdit.ShowEditor(Me, rptList.FocusedRow.Record.Item(����ID).Value, rptList.FocusedRow.Record.Item(ģ������).Value, CInt(rptList.FocusedRow.Record.Item(����ȼ�).Value)) <> "" Then Call zlMenuClick("��ȡ����")
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Delete
            'ɾ����Ŀ
            If rptList.FocusedRow Is Nothing Then Exit Sub
            If rptList.FocusedRow.Record Is Nothing Then Exit Sub

            If MsgBox("�����Ҫɾ����" & rptList.FocusedRow.Record.Item(ģ������).Value & "����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Call zlDatabase.ExecuteProcedure("zl_������Ŀģ��_Delete(" & rptList.FocusedRow.Record.Item(����ID).Value & "," & CInt(rptList.FocusedRow.Record.Item(����ȼ�).Value) & ")", "ɾ��ģ��")
            Call zlMenuClick("��ȡ����")
        '--------------------------------------------------------------------------------------------------------------

        Case conMenu_View_Refresh
            Call zlMenuClick("��ȡ����")

        Case conMenu_Help_Help

            Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))

        Case conMenu_Help_About

            Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)

        Case conMenu_Help_Web_Home

            Call zlHomePage(Me.hwnd)

        Case conMenu_Help_Web_Forum '������̳
            Call zlWebForum(Me.hwnd)

        Case conMenu_Help_Web_Mail

            Call zlMailTo(Me.hwnd)

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

    cbsThis.RecalcLayout
    Call RefreshStateInfo

    Exit Sub

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)

    If stbThis.Visible Then Bottom = stbThis.Height
    
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Err = 0: On Error Resume Next
    
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (vsfDetail.Rows - 1 > 0)
    Case conMenu_Edit_NewItem
        Control.Enabled = (InStr(1, mstrPrivs, "����ģ��") > 0)
    Case conMenu_Edit_Modify, conMenu_Edit_Delete
        If rptList.FocusedRow Is Nothing Then
            Control.Enabled = False
        Else
            If rptList.FocusedRow.Record Is Nothing Then
                Control.Enabled = False
            Else
                '����Ҫ�л���ģ��ı༭Ȩ��
                Control.Enabled = (InStr(1, mstrPrivs, "����ģ��") > 0)
                If Control.Enabled And Val(rptList.FocusedRow.Record.Item(����ID).Value) <> 0 Then
                    '����Ա����б༭��������ģ���Ȩ��,�������޸Ļ�ɾ��,��������
                    Control.Enabled = (InStr(1, ";" & mstrPrivs & ";", ";�༭��������ģ��;") > 0) Or (InStr(1, mstr����ID & ",", "," & rptList.FocusedRow.Record.Item(����ID).Value & ",") <> 0)
                End If
            End If
        End If
    Case conMenu_View_ToolBar_Button
        Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text
        Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size
        Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar
        Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picHead.hwnd
    Case 2
        Item.Handle = picDetail.hwnd
    End Select
End Sub

Private Sub Form_Load()
    Dim objPane As Pane
    
    mstrSel = ""
    mblnStartUp = True
    
    Call InitCommonControls
    Call InitMenuBar
    Call InitRpt
    
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.UseSplitterTracker = False 'ʵʱ�϶�
    dkpMan.Options.AlphaDockingContext = True
    dkpMan.Options.CloseGroupOnButtonClick = True
    dkpMan.Options.HideClient = True
    dkpMan.SetCommandBars cbsThis
    Set objPane = dkpMan.CreatePane(1, 5400, 0, DockLeftOf, Nothing): objPane.Title = "����": objPane.Options = PaneNoCaption
    Set objPane = dkpMan.CreatePane(2, vsfDetail.Width, vsfDetail.Height, DockRightOf, Nothing): objPane.Title = "�ӵ�": objPane.Options = PaneNoCaption
    
    Call RestoreWinState(Me, App.ProductName)
    
    mblnStartUp = False
    Call zlMenuClick("��ʼ��")
    Call zlMenuClick("��ȡ����")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub picDetail_Resize()
    With picDetail
        vsfDetail.Left = 0
        vsfDetail.Top = 0
        vsfDetail.Width = picDetail.Width
        vsfDetail.Height = picDetail.Height
    End With
End Sub

Private Sub picHead_Resize()
    With picHead
        rptList.Left = 0
        rptList.Top = 0
        rptList.Width = picHead.Width
        rptList.Height = picHead.Height
    End With
End Sub

Private Sub rptList_SelectionChanged()
    If rptList.FocusedRow Is Nothing Then Exit Sub
    If rptList.FocusedRow.Record Is Nothing Then Exit Sub
    
    mstrSel = Val(rptList.FocusedRow.Record.Item(����ȼ�).Value) & "|" & Val(rptList.FocusedRow.Record.Item(����ID).Value) & "|" & rptList.FocusedRow.Index
    Call zlMenuClick("��ȡģ������")
End Sub
