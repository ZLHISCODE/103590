VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDockInTendData 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4770
      Index           =   0
      Left            =   195
      ScaleHeight     =   4770
      ScaleWidth      =   9855
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   825
      Width           =   9855
      Begin XtremeReportControl.ReportControl rptData 
         Height          =   2190
         Left            =   435
         TabIndex        =   2
         Top             =   795
         Width           =   3930
         _Version        =   589884
         _ExtentX        =   6932
         _ExtentY        =   3863
         _StockProps     =   0
         ShowGroupBox    =   -1  'True
      End
      Begin VB.Frame fra 
         Height          =   540
         Left            =   15
         TabIndex        =   3
         Top             =   -45
         Width           =   9375
         Begin VB.ComboBox cbo 
            Height          =   300
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   150
            Width           =   3690
         End
         Begin VB.ComboBox cboBaby 
            Height          =   300
            Left            =   9135
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   165
            Width           =   1350
         End
         Begin VB.Label lblData 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��¼��Χ:"
            Height          =   180
            Left            =   60
            TabIndex        =   6
            Top             =   180
            Width           =   810
         End
      End
      Begin MSComctlLib.ImageList imgData 
         Left            =   3810
         Top             =   4080
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
               Picture         =   "frmDockInTendData.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDockInTendData.frx":6862
               Key             =   "����"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDockInTendData.frx":6DFC
               Key             =   "��ͨ"
            EndProperty
         EndProperty
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgPrint 
      Height          =   1395
      Left            =   10185
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1755
      Visible         =   0   'False
      Width           =   1335
      _cx             =   2355
      _cy             =   2461
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
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   0
      Top             =   15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmDockInTendData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
''######################################################################################################################
'
Private Enum mCol
    r��־ = 0: rID: r����ʱ��: r��¼��Ŀ: r��¼����: r����: r��ʿ: r�Ǽ�ʱ��: r����ID: r��¼���: r������:: rǩ����: rǩ��ʱ��: r��Ŀ���: r��ʼ�汾: rδ��˵��: r�鵵��: r�鵵ʱ��
    f��־ = 0: fID: f���: f�ļ�: f���ڷ�Χ: f����id: f������: f������: f����
    w��־ = 0: wID: wҳ����: wҳ������: w��������: w������: w����ʱ��: w������: w���ʱ��: w��ǰ�汾: wǩ������: w��ǰ���: w�鵵��: w�鵵����: w����ID: w������: w����״̬
End Enum

Private Enum mColWidth
    c��־ = 20: cID = 0: c����ʱ�� = 110: c��¼��Ŀ = 100: c��¼���� = 240: c���� = 100: c��ʿ = 60: c�Ǽ�ʱ�� = 110: c����id = 0: c��¼��� = 0: c������ = 100: cǩ���� = 60: cǩ��ʱ�� = 100: c��Ŀ��� = 0: c��ʼ�汾 = 0: cδ��˵�� = 60: c�鵵�� = 60: c�鵵ʱ�� = 110
End Enum

Private mstrPrivs As String                             '��ǰʹ���߶Ա�����(1255)��Ȩ�޴�
Private mlngPatiId As Long                              '����id
Private mlngPageId As Long                              '��ҳid
Private mlngDeptId As Long                              '��ǰ��������id���粡�˿��Һ͵�ǰ���Ҳ�һ�£����ܲ����鵵��Ĺ���
Private mblnEdit As Boolean                             '�Ƿ����������ͨ�����ϼ�������ݵ�ǰ���������Ƿ�ǰ���˲���������
Private mblnDoctorStation As Boolean
Private mintBaby As Integer
Private mblnArchived As Boolean
Private mfrmMain As Object
Private mbytFontSize As Byte
Private WithEvents mfrmCaseTendEdit As frmCaseTendEdit
Attribute mfrmCaseTendEdit.VB_VarHelpID = -1
Private WithEvents mfrmCaseTendEditForBatch As frmCaseTendEditForBatch
Attribute mfrmCaseTendEditForBatch.VB_VarHelpID = -1

Public Event Activate()
Public Event AfterDataChanged()
Public Event AfterArchiveChanged(ByVal blnArchived As Boolean)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)

''######################################################################################################################

Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-18 15:16
    '����:51746
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize
End Sub

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-18 15:16
    '����:51746
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim bytSize As Byte
    Dim lngCol As Long
    Dim PATI_COLWIDTH As Variant
    bytSize = IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))
    
    Me.FontSize = mbytFontSize
    Me.FontName = "����"
    
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
            Case UCase("Label")
                objCtrl.FontSize = mbytFontSize
                objCtrl.Height = TextHeight("��") + 20
            Case UCase("ComboBox")
                objCtrl.FontSize = mbytFontSize
            Case UCase("ReportControl")
                Set CtlFont = objCtrl.PaintManager.CaptionFont
                CtlFont.Size = mbytFontSize
                Set objCtrl.PaintManager.CaptionFont = CtlFont
                
                Set CtlFont = objCtrl.PaintManager.TextFont
                CtlFont.Size = mbytFontSize
                Set objCtrl.PaintManager.TextFont = CtlFont
                PATI_COLWIDTH = Array(c��־, cID, c����ʱ��, c��¼��Ŀ, c��¼����, c����, c��ʿ, c�Ǽ�ʱ��, c����id, c��¼���, c������, cǩ����, rǩ��ʱ��, r��Ŀ���, r��ʼ�汾, rδ��˵��, r�鵵��, r�鵵ʱ��)
                For lngCol = cID To rptData.Columns.Count - 1
                    rptData.Columns.Column(lngCol).Width = BlowUp(CDbl(PATI_COLWIDTH(lngCol)))
                Next lngCol
                '����п������
                objCtrl.Redraw
        End Select
    Next
    
    '����λ�õ���
    cbo.Top = 150
    cbo.Left = lblData.Left + lblData.Width
    cbo.Width = BlowUp(3690)
    lblData.Top = cbo.Top + (cbo.Height - lblData.Height) \ 2
    cboBaby.Top = cbo.Top
    cboBaby.Width = BlowUp(1350)
    fra.Height = cbo.Top + cbo.Height + 75
    Call picPane_Resize(0)
End Sub

Private Function BlowUp(ByRef dblChange As Double) As Double
    '�Ŵ����壬��Ԫ����
    BlowUp = dblChange + (dblChange * IIf(mbytFontSize = 12, 1, 0) / 3)
End Function

Private Function ShowOpenedForm() As Boolean
    
    Dim frmTemp As Form
    
    For Each frmTemp In Forms
        
        If frmTemp.Name = "frmCaseTendEdit" Then
            
            ShowSimpleMsg "�����¼�༭�����Ѵ򿪣������ظ��򿪣��Զ��ָ��Ѵ�״̬��"
            mfrmCaseTendEdit.Show
            
            If mfrmCaseTendEdit.WindowState = 1 Then mfrmCaseTendEdit.WindowState = 0
            mfrmCaseTendEdit.ZOrder 0
            ShowOpenedForm = True
            
            Exit Function
        End If
    Next
End Function

Public Function InitData(ByVal frmMain As Object, ByVal strPrivs As String) As Boolean
'******************************************************************************************************************
'���ܣ�
'������
'���أ�
'******************************************************************************************************************
Dim rptCol As ReportColumn
    mstrPrivs = strPrivs
    Set mfrmMain = frmMain
    
    '------------------------------------------
    '��¼���ݱ�����
    With rptData

        .SetImageList Me.imgData

        .AutoColumnSizing = False
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
        End With


        Set rptCol = .Columns.Add(mCol.r��־, "", mColWidth.c��־, False)
        rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter

        Set rptCol = .Columns.Add(mCol.rID, "ID", mColWidth.cID, False): rptCol.Editable = False: rptCol.Groupable = False

        Set rptCol = .Columns.Add(mCol.r����ʱ��, "����ʱ��", mColWidth.c����ʱ��, False): rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(mCol.r��¼��Ŀ, "��¼��Ŀ", mColWidth.c��¼��Ŀ, True): rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(mCol.r��¼����, "��¼����", mColWidth.c��¼����, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.AutoSize = True
        Set rptCol = .Columns.Add(mCol.r����, "����", mColWidth.c����, True): rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(mCol.r��ʿ, "��ʿ", mColWidth.c��ʿ, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.r�Ǽ�ʱ��, "�Ǽ�ʱ��", mColWidth.c�Ǽ�ʱ��, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.r����ID, "����ID", mColWidth.c����id, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.r��¼���, "��¼���", mColWidth.c��¼���, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.r������, "����", mColWidth.c������, True): rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(mCol.rǩ����, "ǩ����", mColWidth.cǩ����, True): rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(mCol.rǩ��ʱ��, "ǩ��ʱ��", mColWidth.cǩ��ʱ��, True): rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(mCol.r��Ŀ���, "��Ŀ���", mColWidth.c��Ŀ���, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.r��ʼ�汾, "��ʼ�汾", mColWidth.c��ʼ�汾, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.rδ��˵��, "δ��˵��", mColWidth.cδ��˵��, True):   rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(mCol.r�鵵��, "�鵵��", mColWidth.c�鵵��, True):   rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.r�鵵ʱ��, "�鵵ʱ��", mColWidth.c�鵵ʱ��, True):   rptCol.Editable = False: rptCol.Groupable = False

    End With

    With cboBaby
        .AddItem "���˱���"
        .ListIndex = 0
    End With

'    If ExecuteCommand("��ʼ�ؼ�") = False Or ExecuteCommand("��ʼ����") = False Then Exit Function
'    Call ExecuteCommand("��ע���")
'    Call ExecuteCommand("�ؼ�״̬")
    
End Function

Public Function RefreshData(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lngDeptId As Long, ByVal blnDoctorStation As Boolean, ByVal blnEdit As Boolean) As Boolean
    '******************************************************************************************************************
    '���ܣ�ˢ������
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    
    mlngPatiId = lng����ID
    mlngPageId = lng��ҳID
    mblnDoctorStation = blnDoctorStation
    mblnEdit = blnEdit And Not mblnMoved_HL
    mlngDeptId = lngDeptId
    
    cboBaby.Clear
    cboBaby.AddItem "���˱���"
    
    gstrSQL = "Select a.���,Decode(a.Ӥ������,Null,NVL(c.����,b.����) ||'֮��'||Trim(To_Char(a.���,'9')),a.Ӥ������) As Ӥ������" & _
        " From ������Ϣ b,������ҳ c,������������¼ a Where b.����id=c.����id And a.����id=c.����id And a.��ҳid=c.��ҳid And c.����id=[1] And c.��ҳid=[2]  Order By a.���"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngPatiId, mlngPageId)
    If rs.BOF = False Then
        Do While Not rs.EOF
            cboBaby.AddItem rs("Ӥ������").Value
            rs.MoveNext
        Loop
    End If

    cboBaby.ListIndex = 0
    cboBaby.Visible = (cboBaby.ListCount > 1)

    Call zlRefDate(mlngPatiId, mlngPageId)
    
End Function

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strInfo As String
        
    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview
        Call zlRptPrint(0)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Print
        Call zlRptPrint(1)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Excel
        Call zlRptPrint(3)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_RowPrint
        Call zlRptPrint(1)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem

        '����Ǽǣ�����ID;��ҳID;����ID;��Ժ;�༭;Ӥ��
        
        If ShowOpenedForm = False Then
            If mfrmCaseTendEdit Is Nothing Then Set mfrmCaseTendEdit = New frmCaseTendEdit
            If mfrmCaseTendEdit.ShowEdit(Me, mlngPatiId & ";" & mlngPageId & ";" & mlngDeptId & ";" & mintBaby & ";2;", 1, mstrPrivs) Then
    '            RaiseEvent AfterDataChanged
    '            cbo.Tag = ""
    '            Call zlRefRec
            End If
        End If

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Modify                    '�޸Ļ����¼����
        
        If ExecuteCommand("�޸Ļ�������") Then
'            cbo.Tag = ""
'            RaiseEvent AfterDataChanged
'            Call zlRefRec
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete

        'ɾ������Ǽ�
        If rptData.FocusedRow Is Nothing Then Exit Sub
        If rptData.FocusedRow.Record Is Nothing Then Exit Sub

        If MsgBox("ȷ��Ҫɾ����ǰ�Ļ����¼��", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then Exit Sub

        Dim strStart As String
        Dim strEnd As String
        Dim strDate As String

        strDate = rptData.FocusedRow.Record(mCol.r����ʱ��).Value
        strStart = strDate & ":00"
        strEnd = Format(DateAdd("n", 1, CDate(strDate)), "yyyy-MM-dd HH:mm") & ":00"

        gstrSQL = "ZL_���ӻ����¼_UPDATE("
        gstrSQL = gstrSQL & mlngPatiId & ","
        gstrSQL = gstrSQL & mlngPageId & ","
        gstrSQL = gstrSQL & mintBaby & ","
        gstrSQL = gstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
        gstrSQL = gstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
        gstrSQL = gstrSQL & "1,"
        gstrSQL = gstrSQL & Val(rptData.FocusedRow.Record(mCol.r��Ŀ���).Value) & ","
        gstrSQL = gstrSQL & Val(rptData.FocusedRow.Record(mCol.r��¼���).Value) & ","
        gstrSQL = gstrSQL & "NULL"
        gstrSQL = gstrSQL & ")"

        'ִ��
        Err = 0: On Error GoTo errHand
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Err = 0: On Error GoTo 0
        RaiseEvent AfterDataChanged
        cbo.Tag = ""
        Call zlRefRec

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Search

        '�����¼
        If rptData.FocusedRow Is Nothing Then Exit Sub
        If rptData.FocusedRow.Record Is Nothing Then Exit Sub
        
        If ShowOpenedForm = False Then
            If mfrmCaseTendEdit Is Nothing Then Set mfrmCaseTendEdit = New frmCaseTendEdit
            Call mfrmCaseTendEdit.ShowEdit(Me, mlngPatiId & ";" & mlngPageId & ";" & mlngDeptId & ";" & mintBaby & ";2;" & CStr(rptData.FocusedRow.Record.Item(mCol.r����ʱ��).Value), 5)
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Sign


        '�����¼
        If rptData.FocusedRow Is Nothing Then Exit Sub
        If rptData.FocusedRow.Record Is Nothing Then Exit Sub
        
        If ShowOpenedForm = False Then
            
            If mfrmCaseTendEdit Is Nothing Then Set mfrmCaseTendEdit = New frmCaseTendEdit
            If mfrmCaseTendEdit.ShowEdit(Me, mlngPatiId & ";" & mlngPageId & ";" & mlngDeptId & ";" & mintBaby & ";2;" & CStr(rptData.FocusedRow.Record.Item(mCol.r����ʱ��).Value), 3) Then
    '            cbo.Tag = ""
    '            RaiseEvent AfterDataChanged
    '            Call zlRefRec
            End If
        End If

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_SignEarse

        '�����¼
        If rptData.FocusedRow Is Nothing Then Exit Sub
        If rptData.FocusedRow.Record Is Nothing Then Exit Sub
        
        If ShowOpenedForm = False Then
            If mfrmCaseTendEdit Is Nothing Then Set mfrmCaseTendEdit = New frmCaseTendEdit
            If mfrmCaseTendEdit.ShowEdit(Me, mlngPatiId & ";" & mlngPageId & ";" & mlngDeptId & ";" & mintBaby & ";2;" & CStr(rptData.FocusedRow.Record.Item(mCol.r����ʱ��).Value), 4) Then
    '            cbo.Tag = ""
    '            RaiseEvent AfterDataChanged
    '            Call zlRefRec
            End If
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Archive * 10

        If MsgBox("��Ҫ���ò��˱���סԺ���л����¼�鵵��", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbYes Then

            Dim strNow As String

            strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            gstrSQL = "Zl_���ӻ����¼_Archive(" & mlngPatiId & "," & mlngPageId & "," & mintBaby & ",'" & gstrUserName & "',To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'))"
            Err = 0: On Error GoTo errHand
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            mblnArchived = True
                        
            cbo.Tag = ""
            Call zlRefRec
            
            Err = 0: On Error GoTo 0

        End If

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_UnArchive

        If mblnArchived Then
            If MsgBox("��Ҫ�����ò��˱���סԺ�����ѹ鵵�����¼��", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbYes Then

                gstrSQL = "Zl_���ӻ����¼_UnArchive(" & mlngPatiId & "," & mlngPageId & "," & mintBaby & ")"
                Err = 0: On Error GoTo errHand
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                
                mblnArchived = False
                cbo.Tag = ""
                Call zlRefRec
                
                Err = 0: On Error GoTo 0

            End If
        End If

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MarkMap
        '������ͼ������ID;��ҳID;����ID;��Ժ;�༭;Ӥ��
        If Not CreateBodyEditor Then Exit Sub
        If gobjBodyEditor.GetTendBody.ShowEdit(Me, mlngPatiId & ";" & mlngPageId & ";" & mlngDeptId & ";0;1;" & mintBaby, 2, mstrPrivs) Then

            Call zlRefDate(mlngPatiId, mlngPageId)

            RaiseEvent AfterDataChanged
            cbo.Tag = ""
            Call zlRefRec
        End If

    Case conMenu_File_PrintDayDetail        '����¼��
        If mfrmCaseTendEditForBatch Is Nothing Then Set mfrmCaseTendEditForBatch = New frmCaseTendEditForBatch
        Call mfrmCaseTendEditForBatch.ShowMe(Me, mlngDeptId, mstrPrivs)
        
    End Select
    
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
LL:
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
'    Dim lngCount As Long, blnFinished As Boolean, lngMaxVersion As Long, eSignLevel As EPRSignLevelEnum

    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_ExportToXML
        Control.Enabled = (rptData.Records.Count > 0)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Excel
        Control.Enabled = (rptData.Records.Count > 0)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem
        
        Control.Visible = (InStr(1, mstrPrivs, "�����¼�Ǽ�") > 0 And mblnDoctorStation = False)
        Control.Enabled = (Control.Visible And mblnEdit And mlngPatiId > 0 And mblnArchived = False And Not mblnMoved_HL)
                
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Modify
    
        Control.Enabled = (mblnEdit And mlngPatiId > 0 And mblnArchived = False And Not mblnMoved_HL)

        Control.Visible = (InStr(1, mstrPrivs, "�����¼�Ǽ�") > 0 And mblnDoctorStation = False)
        If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "�����¼�Ǽ�") > 0)
        If Control.Enabled Then Control.Enabled = Not (Me.rptData.FocusedRow Is Nothing)
        If Control.Enabled Then Control.Enabled = Not Me.rptData.FocusedRow.GroupRow
        If Control.Enabled Then Control.Enabled = (Me.rptData.FocusedRow.Record(mCol.rID).Value <> 0)
        If Control.Enabled Then Control.Enabled = (Me.rptData.FocusedRow.Record(mCol.rID).Value >= 0)
        If (InStr(1, mstrPrivs, "���˻����¼") = 0) Then
            If Control.Enabled Then Control.Enabled = (Me.rptData.FocusedRow.Record(mCol.r��ʿ).Value = gstrUserName)
        End If

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete
        Control.Enabled = (mblnEdit And mlngPatiId > 0 And mblnArchived = False And Not mblnMoved_HL)

        Control.Visible = (InStr(1, mstrPrivs, "�����¼�Ǽ�") > 0 And mblnDoctorStation = False)

        If Control.Enabled Then Control.Enabled = Control.Visible
        If Control.Enabled Then Control.Enabled = Not (Me.rptData.FocusedRow Is Nothing)
        If Control.Enabled Then Control.Enabled = Not Me.rptData.FocusedRow.GroupRow
        If Control.Enabled Then Control.Enabled = (Me.rptData.FocusedRow.Record(mCol.rID).Value <> 0)
        If Control.Enabled Then Control.Enabled = (Me.rptData.FocusedRow.Record(mCol.rID).Value >= 0)
        If (InStr(1, mstrPrivs, "���˻����¼") = 0) Then
            If Control.Enabled Then Control.Enabled = (Me.rptData.FocusedRow.Record(mCol.r��ʿ).Value = gstrUserName)
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Search

        Control.Visible = (mblnDoctorStation = False)
        Control.Enabled = (mlngPatiId > 0 And Control.Visible)
        If Control.Enabled Then Control.Enabled = Not (Me.rptData.FocusedRow Is Nothing)
        If Control.Enabled Then Control.Enabled = Not Me.rptData.FocusedRow.GroupRow
        If Control.Enabled Then Control.Enabled = (Me.rptData.FocusedRow.Record(mCol.rID).Value <> 0)
        If Control.Enabled Then Control.Enabled = (Me.rptData.FocusedRow.Record(mCol.rID).Value >= 0)
        If Control.Enabled Then Control.Enabled = (rptData.FocusedRow.Record(mCol.r��ʼ�汾).Value > 0)

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Sign

        Control.Visible = (InStr(1, mstrPrivs, "�����¼ǩ��") > 0 And mblnDoctorStation = False)
        Control.Enabled = (mblnEdit And mlngPatiId > 0 And Control.Visible And mblnArchived = False And Not mblnMoved_HL)
        If Control.Enabled Then Control.Enabled = Not (Me.rptData.FocusedRow Is Nothing)
        If Control.Enabled Then Control.Enabled = Not Me.rptData.FocusedRow.GroupRow
        If Control.Enabled Then Control.Enabled = (Me.rptData.FocusedRow.Record(mCol.rID).Value <> 0)
        If Control.Enabled Then Control.Enabled = (Me.rptData.FocusedRow.Record(mCol.rID).Value >= 0)
        If Control.Enabled Then Control.Enabled = (rptData.FocusedRow.Record(mCol.rǩ����).Value = "")
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_SignEarse
        Control.Visible = (InStr(1, mstrPrivs, "ȡ����¼ǩ��") > 0 And mblnDoctorStation = False)
        Control.Enabled = (mblnEdit And mlngPatiId > 0 And Control.Visible And mblnArchived = False And Not mblnMoved_HL)
        If Control.Enabled Then Control.Enabled = Not (Me.rptData.FocusedRow Is Nothing)
        If Control.Enabled Then Control.Enabled = Not Me.rptData.FocusedRow.GroupRow
        If Control.Enabled Then Control.Enabled = (Me.rptData.FocusedRow.Record(mCol.rID).Value <> 0)
        If Control.Enabled Then Control.Enabled = (Me.rptData.FocusedRow.Record(mCol.rID).Value >= 0)
        If Control.Enabled Then Control.Enabled = (rptData.FocusedRow.Record(mCol.r��ʼ�汾).Value > 0)
    
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Archive * 10

        Control.Visible = (InStr(1, mstrPrivs, "�����¼�鵵") > 0 And mblnDoctorStation = False And mblnArchived = False)
        Control.Enabled = Control.Visible And mblnEdit And Not mblnMoved_HL

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_UnArchive

        Control.Visible = (InStr(1, mstrPrivs, "ȡ����¼�鵵") > 0 And mblnDoctorStation = False And mblnArchived)
        Control.Enabled = Control.Visible And mblnEdit And Not mblnMoved_HL

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_EditPopup

        Control.Visible = (mblnDoctorStation = False And (InStr(1, mstrPrivs, "���µ���ͼ") > 0 Or InStr(1, mstrPrivs, "�����¼�Ǽ�") > 0))
        Control.Enabled = Control.Visible
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MarkMap

        Control.Visible = (InStr(1, mstrPrivs, "���µ���ͼ") > 0 And mblnDoctorStation = False)
        Control.Enabled = (mblnEdit And mlngPatiId > 0 And Control.Visible And mblnArchived = False And Not mblnMoved_HL)
    
    Case conMenu_File_PrintDayDetail
        Control.Enabled = (mblnEdit And mlngPatiId > 0 And mblnArchived = False And Not mblnMoved_HL)

        Control.Visible = (InStr(1, mstrPrivs, "�����¼�Ǽ�") > 0 And mblnDoctorStation = False)
        If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "�����¼�Ǽ�") > 0) And Not mblnDoctorStation
    End Select
End Sub

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

    Call SQLRecord(rsSQL)

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ�ؼ�"
        
        
               
            
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"

        
    '------------------------------------------------------------------------------------------------------------------
    Case "�ؼ�״̬"

        
    '------------------------------------------------------------------------------------------------------------------
    Case "ˢ��״̬"
        

        
    '------------------------------------------------------------------------------------------------------------------
    Case "ˢ������"
                
        
    '------------------------------------------------------------------------------------------------------------------
    Case "�޸Ļ�������"
    
        '����Ǽ�
        If rptData.FocusedRow Is Nothing Then Exit Function
        If rptData.FocusedRow.Record Is Nothing Then Exit Function
        
        If ShowOpenedForm = False Then
            If mfrmCaseTendEdit Is Nothing Then Set mfrmCaseTendEdit = New frmCaseTendEdit
            ExecuteCommand = mfrmCaseTendEdit.ShowEdit(Me, mlngPatiId & ";" & mlngPageId & ";" & mlngDeptId & ";" & mintBaby & ";2;" & CStr(rptData.FocusedRow.Record.Item(mCol.r����ʱ��).Value) & ";" & CStr(rptData.FocusedRow.Record.Item(mCol.rID).Value), 2, mstrPrivs)
        End If
        
        Exit Function
        
    End Select

    ExecuteCommand = True

    GoTo EndHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
EndHand:

End Function

Private Sub cbo_Click()

    Call zlRefRec

End Sub

Private Sub cboBaby_Click()

    If mintBaby = cboBaby.ListIndex Then Exit Sub
    mintBaby = cboBaby.ListIndex

    Call zlRefDate(mlngPatiId, mlngPageId)

End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    picPane(0).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mfrmMain = Nothing
    If Not mfrmCaseTendEdit Is Nothing Then Unload mfrmCaseTendEdit
    If Not mfrmCaseTendEditForBatch Is Nothing Then Unload mfrmCaseTendEditForBatch
    Set mfrmCaseTendEdit = Nothing
    Set mfrmCaseTendEditForBatch = Nothing
    
End Sub

Private Sub mfrmCaseTendEdit_AfterDataChanged()
    RaiseEvent AfterDataChanged
    cbo.Tag = ""
    Call zlRefDate(mlngPatiId, mlngPageId)
End Sub

Private Sub mfrmCaseTendEditForBatch_AfterDataChanged()
    RaiseEvent AfterDataChanged
    cbo.Tag = ""
    
    Call zlRefDate(mlngPatiId, mlngPageId)
End Sub

Private Sub rptData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Not (rptData.FocusedRow Is Nothing) Then
            Call rptData_RowDblClick(rptData.FocusedRow, rptData.FocusedRow.Record.Item(mCol.r����ʱ��))
        End If
    End If
End Sub

Private Sub rptData_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    
    RaiseEvent MouseUp(Button, Shift, X, Y)
    
End Sub

Private Sub rptData_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)

    If Not (rptData.FocusedRow Is Nothing) Then
        
        RaiseEvent RowDblClick(Row, Item)

    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next

    RaiseEvent Activate
End Sub

Private Function zlRefDate(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim cbrItem As CommandBarControl
    Dim intCount As Integer
    Dim strDateFrom As String
    Dim strDateTo As String
    Dim strEnterDate As String
    Dim intCol As Integer
    Dim strCaption As String
    Dim strParameter As String
    Dim strSvrCaption As String
    Dim strNow As String
    Dim strCut As String
    Dim lngLoop As Long
    Dim strTmp As String
    Dim lnglast����id As Long
    Dim intSvrDate As Integer
    Dim blnData As Boolean '�Ƿ�����ϰ�����
    
    strCut = "123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    
    If cbo.ListIndex >= 0 Then intSvrDate = cbo.ItemData(cbo.ListIndex)
    
    cbo.Clear
    cbo.Tag = ""
    cbo.AddItem "���м�¼"
    cbo.ItemData(cbo.NewIndex) = 0

    '------------------------------------------------------------------------------------------------------------------
                
    strSQL = "Select ��Ժʱ��, ��Ժʱ��, 1 + Nvl(Round((b.��Ժʱ�� - b.��Ժʱ��) / 7),-1) As ҳ��" & vbNewLine & _
                "  from (Select Min(����ʱ��) as ��Ժʱ��," & vbNewLine & _
                "               Max(����ʱ��) as ��Ժʱ��" & vbNewLine & _
                "          From ���˻����¼" & vbNewLine & _
                "         Where ����ID = [1] And ��ҳID = [2]) b"
    If mblnMoved_HL Then
        strSQL = Replace(strSQL, "���˻����¼", "H���˻����¼")
        strSQL = Replace(strSQL, "���˻�������", "H���˻�������")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "usrBodyEditor", lng����ID, lng��ҳID)
    If rsTmp.BOF Then Exit Function
    
    If NVL(rsTmp!��Ժʱ��) <> "" Then blnData = True
    
    '
    '------------------------------------------------------------------------------------------------------------------
                
'    strSQL = "Select 1 As ��ʼҳ��,1 + Round((a.��ֹʱ�� - a.��ʼʱ��) / 7) As ����ҳ��," & vbNewLine & _
'                "       ����id,c.����," & vbNewLine & _
'                "       ��ʼʱ��," & vbNewLine & _
'                "       ��ֹʱ��" & vbNewLine & _
'                "  from (Select ����id," & vbNewLine & _
'                "               Min(����ʱ��) as ��ʼʱ��," & vbNewLine & _
'                "               Max(����ʱ��) as ��ֹʱ��" & vbNewLine & _
'                "          From ���˻����¼" & vbNewLine & _
'                "         Where ����ID = [1] And ��ҳID = [2]" & vbNewLine & _
'                "         Group by ����id) a," & vbNewLine & _
'                "       (Select Min(��ʼʱ��) as ��Ժʱ��" & vbNewLine & _
'                "          From ���˱䶯��¼" & vbNewLine & _
'                "         Where ��ʼʱ�� is Not Null And ����ID = [1] And ��ҳID = [2]) b,���ű� c Where c.ID=a.����id " & vbNewLine & _
'                " order by a.��ʼʱ��"

    strSQL = "Select 1 As ��ʼҳ��, 1 + Round((a.��ֹʱ�� - a.��ʼʱ��) / 7) As ����ҳ��, ��ʼʱ��, ��ֹʱ��" & vbNewLine & _
             "   From (Select Min(����ʱ��) As ��ʼʱ��, Max(����ʱ��) As ��ֹʱ��" & vbNewLine & _
             "          From ���˻����¼" & vbNewLine & _
             "          Where ����id = [1] And ��ҳid = [2]) A," & vbNewLine & _
             "        (Select Min(��ʼʱ��) As ��Ժʱ��" & vbNewLine & _
             "          From ���˱䶯��¼" & vbNewLine & _
             "          Where ��ʼʱ�� Is Not Null And ����id = [1] And ��ҳid = [2]) B" & vbNewLine & _
             "   Order By a.��ʼʱ��"
    If mblnMoved_HL Then
        strSQL = Replace(strSQL, "���˻����¼", "H���˻����¼")
        strSQL = Replace(strSQL, "���˻�������", "H���˻�������")
    End If
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "usrBodyEditor", lng����ID, lng��ҳID)

    strEnterDate = Format(rsTmp!��Ժʱ��, "yyyy-MM-dd HH:mm:ss")
    For lngLoop = 0 To rsTmp("ҳ��").Value - 1

        strDateFrom = Format(rsTmp("��Ժʱ��").Value + 7 * lngLoop, "yyyy-MM-dd") & " 00:00:00"
        strDateTo = Format(rsTmp("��Ժʱ��").Value + 7 * (lngLoop + 1) - 1, "yyyy-MM-dd") & " 23:59:59"
        If strDateFrom < Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then
            strDateFrom = Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss")
        End If

        If strDateFrom < Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then

            If strDateFrom < Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then strDateFrom = Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss")
            If strDateTo > Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then strDateTo = Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss")

            rs.Filter = ""
            rs.Filter = "��ʼҳ��<=" & lngLoop + 1 & " And ����ҳ��>=" & lngLoop + 1
            rs.Sort = "��ʼʱ��"
            If rs.RecordCount > 0 Then rs.MoveFirst
            For intCol = 1 To rs.RecordCount

                If strDateFrom < Format(rs("��ʼʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then
                    strTmp = Format(rs("��ʼʱ��").Value, "yyyy-MM-dd HH:mm:ss")
                Else
                    strTmp = strDateFrom
                End If

                If strDateTo > Format(rs("��ֹʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then
                    strCaption = Format(rs("��ֹʱ��").Value, "yyyy-MM-dd HH:mm:ss")
                Else
                    strCaption = strDateTo
                End If

                strCaption = Format(strTmp, "yyyy��MM��dd��") & " �� " & Format(strCaption, "yyyy��MM��dd��")

                cbo.AddItem strCaption
                cbo.ItemData(cbo.NewIndex) = intCol

                rs.MoveNext

            Next
        End If

    Next
    
    If intSvrDate > 0 Then
        Call zlControl.CboLocate(cbo, intSvrDate)
        If cbo.ListIndex = -1 Then cbo.ListIndex = cbo.ListCount - 1
    Else
        cbo.ListIndex = cbo.ListCount - 1
    End If
    
    If mblnEdit = True Then
        '41778,������,2012-09-06
        '��������ϰ���°����ݶ��Ѿ����ڣ������κ����ơ����ֻ���°����ݣ�û���ϰ档���ϰ岻������ļ���
        'Ӥ��Ӧ�ú�ĸ��ʹ��ͬһ��ϵͳ��
        strSQL = "Select 1 From ���˻����ļ� A Where a.����id = [1] And a.��ҳid = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
        If rsTmp.RecordCount > 0 And blnData = False Then
            mblnEdit = False
        End If
    End If
    
    zlRefDate = True
End Function

Private Function zlRefRec() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
Dim rsTemp As New ADODB.Recordset
Dim strTmp As String
Dim strStart As String
Dim strEnd As String
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem

    On Error GoTo errHand

    If cbo.Tag = cbo.Text Then
        zlRefRec = True
        Exit Function
    End If

    cbo.Tag = cbo.Text

    mblnArchived = False

    gstrSQL = "Select �鵵��,�鵵ʱ�� From ���˻����¼ Where ����id=[1] And ��ҳid=[2] And Nvl(Ӥ��,0)=[3] And RowNum<2 And �鵵�� Is Not Null"
    If mblnMoved_HL Then
        gstrSQL = Replace(gstrSQL, "���˻����¼", "H���˻����¼")
        gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId, mintBaby)
    If rsTemp.BOF = False Then
        mblnArchived = True
    End If
    
    RaiseEvent AfterArchiveChanged(mblnArchived)
    '

    '------------------------------------------------------------------------------------------------------------------
    '��������ˢ��
    If cbo.ItemData(cbo.ListIndex) = 0 Then
        gstrSQL = "Select Decode(f.��¼��,Null,0,1) As ��ʼ�汾,c.Id, e.��¼�� As ǩ����,e.��Ŀ���� As ǩ��ʱ��,Nvl(c.δ��˵��,c.��¼����) As ����,c.��¼���,r.����ʱ��, c.��Ŀ����, Decode(c.��Ŀ���,1,Decode(c.��¼���,1,Null,Decode(c.���²�λ,Null,'Ҹ��',c.���²�λ)||':'),Null)||c.��¼���� || c.��Ŀ��λ || Decode(c.��¼���, 1, Decode(c.��Ŀ���,1,'(������)',Null), Null) As ��¼����," & _
                "        c.��Ŀ����, c.��Ŀ���, c.��¼��, nvl(c.�޸�ʱ��,r.����ʱ��) AS ����ʱ��, r.����id As ����id, d.���� As ������,c.δ��˵��,r.�鵵��,r.�鵵ʱ�� " & _
                " From ���˻����¼ r, ���˻������� c, ���ű� d,���˻������� e,���˻������� f " & _
                " Where r.Id = c.��¼id And r.����id = d.Id And r.����id = [1] And r.��ҳid = [2] And c.��¼���� = 1 And  Nvl(r.Ӥ��,0)=[3] And c.��ֹ�汾 Is Null And r.ID=e.��¼id(+) And e.��¼����(+)=5 And Nvl(r.���汾,1)=Nvl(e.��ʼ�汾(+),1)  And f.��¼id(+)=c.��¼id And f.��¼����(+)=5 And Nvl(f.��ʼ�汾(+),1)=1 " & _
                " Order By r.����ʱ�� Desc"
        If mblnMoved_HL Then
            gstrSQL = Replace(gstrSQL, "���˻����¼", "H���˻����¼")
            gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId, mintBaby)
    Else

        strTmp = Trim(Mid(cbo.Text, InStr(cbo.Text, ")") + 1))
        strStart = Format(Trim(Mid(strTmp, 1, InStr(strTmp, "��") - 1)), "yyyy-MM-dd")
        strEnd = Format(Trim(Mid(strTmp, InStr(strTmp, "��") + 1)), "yyyy-MM-dd") & " 23:59:59"

        gstrSQL = "Select Decode(f.��¼��,Null,0,1) As ��ʼ�汾,c.Id, e.��¼�� As ǩ����,e.��Ŀ���� As ǩ��ʱ��,c.��¼���� As ����,c.��¼���,r.����ʱ��, c.��Ŀ����, Decode(c.��Ŀ���,1,Decode(c.��¼���,1,Null,Decode(c.���²�λ,Null,'Ҹ��',c.���²�λ)||':'),Null)||c.��¼���� || c.��Ŀ��λ || Decode(c.��¼���, 1, Decode(c.��Ŀ���,1,'(������)',Null), Null) As ��¼����," & _
                "        c.��Ŀ����, c.��Ŀ���, c.��¼��, nvl(c.�޸�ʱ��,r.����ʱ��) AS ����ʱ��, r.����id As ����id, d.���� As ������,c.δ��˵��,r.�鵵��,r.�鵵ʱ�� " & _
                " From ���˻����¼ r, ���˻������� c, ���ű� d,���˻������� e,���˻������� f " & _
                " Where r.Id = c.��¼id And r.����id = d.Id And r.����id = [1] And r.��ҳid = [2] And c.��¼���� = 1 And  Nvl(r.Ӥ��,0)=[5]  And ����ʱ�� Between [3] And [4] And c.��ֹ�汾 Is Null And r.ID=e.��¼id(+) And e.��¼����(+)=5 And Nvl(r.���汾,1)=Nvl(e.��ʼ�汾(+),1) And f.��¼id(+)=c.��¼id And f.��¼����(+)=5 And Nvl(f.��ʼ�汾(+),1)=1 " & _
                " Order By r.����ʱ�� Desc"
        If mblnMoved_HL Then
            gstrSQL = Replace(gstrSQL, "���˻����¼", "H���˻����¼")
            gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
        End If

        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId, CDate(strStart), CDate(strEnd), mintBaby)
    End If

    rptData.Records.DeleteAll
    With rsTemp
        Do While Not rsTemp.EOF
            Set rptRcd = rptData.Records.Add()
            Set rptItem = rptRcd.AddItem(""): rptItem.Icon = 0
            rptRcd.AddItem CStr("" & !ID)
            rptRcd.AddItem Format(!����ʱ��, "yyyy-MM-dd hh:mm")
            rptRcd.AddItem CStr("" & !��Ŀ����)

            strTmp = CStr("" & !��¼����)
            Select Case rsTemp("��Ŀ���").Value
            Case 9
                If Right(rsTemp("����").Value, 1) = "C" Then
                    strTmp = CStr("" & !����)
                End If
            Case 10
                If zlCommFun.NVL(rsTemp("����").Value) <> "" Then
                    If Right(rsTemp("����").Value, 2) = "/E" Then
                        strTmp = CStr("" & !����)
                    ElseIf Right(rsTemp("����").Value, 1) = "E" Then
                        strTmp = CStr("" & !����)
                    ElseIf Right(rsTemp("����").Value, 1) = "*" Then
                        strTmp = CStr("" & !����)
                    End If
                End If
            End Select
            
            If zlCommFun.NVL(rsTemp("δ��˵��").Value) <> "" Then
                rptRcd.AddItem CStr(rsTemp("δ��˵��").Value)
            Else
                        
                If zlCommFun.NVL(rsTemp("����").Value) = "" Then
                    rptRcd.AddItem ""
                Else
                    rptRcd.AddItem strTmp
                End If
            End If

            rptRcd.AddItem CStr("" & !��Ŀ����)
            rptRcd.AddItem CStr("" & !��¼��)
            rptRcd.AddItem Format(!����ʱ��, "yyyy-MM-dd hh:mm")
            rptRcd.AddItem CStr("" & !����ID)
            rptRcd.AddItem CStr("" & !��¼���)
            rptRcd.AddItem CStr("" & !������)
            rptRcd.AddItem CStr("" & !ǩ����)
            rptRcd.AddItem Format(!ǩ��ʱ��, "yyyy-MM-dd hh:mm")
            rptRcd.AddItem CStr("" & !��Ŀ���)
            rptRcd.AddItem Val(zlCommFun.NVL(!��ʼ�汾))
            rptRcd.AddItem CStr("" & !δ��˵��)
            rptRcd.AddItem CStr("" & !�鵵��)
            If IsNull(!�鵵ʱ��) Then
                rptRcd.AddItem ""
            Else
                rptRcd.AddItem Format(!�鵵ʱ��, "yyyy-MM-dd hh:mm")
            End If
                        
            .MoveNext
        Loop
    End With
    Me.rptData.Populate
    If Me.rptData.Records.Count > 0 Then Set Me.rptData.FocusedRow = rptData.Rows(0)

    zlRefRec = True

    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '       strSubhead����ӡ�ĸ�����
    '-------------------------------------------------
Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
Dim rsTemp As New ADODB.Recordset
    
    '��û�����Ϣ
    Dim strSubhead As String
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select b.סԺ��, NVL(b.����,a.����) ���� From ������Ϣ a,������ҳ b Where a.����id=b.����id And b.����id = [1] And b.��ҳid=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId)
    If Not rsTemp.EOF Then
        strSubhead = "סԺ��:" & rsTemp!סԺ�� & "  ����:" & rsTemp!����
    Else
        strSubhead = ""
    End If
    Err = 0: On Error GoTo 0

    If Me.rptData.Records.Count = 0 Then Exit Sub
    If zlReportToVSFlexGrid(Me.vfgPrint, Me.rptData) = False Then Exit Sub

    Call vfgPrint.AutoSize(0, vfgPrint.Cols - 1)

    Set objPrint.Body = Me.vfgPrint
    objPrint.Title.Text = "�����¼�����嵥"


    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add(strSubhead)
    Call objAppRow.Add("��" & mlngPageId & "��סԺ")
    Call objPrint.UnderAppRows.Add(objAppRow)

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
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
    
        fra.Move 0, -90, picPane(Index).Width
        
        cboBaby.Move fra.Width - cboBaby.Width, cboBaby.Top
        
        rptData.Move 15, fra.Top + fra.Height + 15, picPane(Index).Width - 30, picPane(Index).Height - (fra.Top + fra.Height + 15) - 15
    End Select
End Sub
