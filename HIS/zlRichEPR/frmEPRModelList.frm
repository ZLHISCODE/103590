VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.Form frmEPRModelList 
   BorderStyle     =   0  'None
   Caption         =   "��������"
   ClientHeight    =   5190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5310
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   3090
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   4905
      _Version        =   589884
      _ExtentX        =   8652
      _ExtentY        =   5450
      _StockProps     =   0
      BorderStyle     =   2
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   165
      Top             =   3045
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
            Picture         =   "frmEPRModelList.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelList.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelList.frx":0B34
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vgdList 
      Height          =   900
      Left            =   135
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3765
      Visible         =   0   'False
      Width           =   1080
      _cx             =   1905
      _cy             =   1587
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
End
Attribute VB_Name = "frmEPRModelList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const con_UnDefine = -999

Const conColumn_ͼ�� = 0
Const conColumn_ID = 1
Const conColumn_��� = 2
Const conColumn_���� = 3
Const conColumn_˵�� = 4
Const conColumn_���� = 5
Const conColumn_��Ա = 6

'---------------------------------
'�����¼�
'---------------------------------
Public Event RightMouseUp(X As Long, Y As Long)                         '�Ҽ�����¼�
Public Event RowDblClick(ByVal Row As XtremeReportControl.IReportRow)   '˫��һ�л������ϰ��س�
Public Event SelRowChanged(ByVal Row As XtremeReportControl.IReportRow) 'ѡ���иı�

'---------------------------------
'�������
'---------------------------------
Private mlngFileID As Long           '��ǰָ�����ļ�id
Private mintPower As Integer        '�ʾ����Ȩ��Χ
'    mintPower=con_UnDefine��δ����;
'    mintPower=-1�����߱��ʾ����Ȩ;
'    mintPower=0��ȫԺ����ʱ��ʾ���е�ʾ����Ҳ���Ը���;
'    mintPower=1�����ң���ʱ��ʾȫԺͨ��ʾ��(����id is null)�����ڿ��ҹ��л�������Ա˽�е�ʾ���������ܸ���ȫԺͨ��ʾ��;
'    mintPower=2�����ˣ���ʱ��ʾȫԺͨ��ʾ��(����id is null)�����ڿ���ͨ��ʾ��(��Աid is null)�͸���ʾ����������ʾ���ɸ���

'---------------------------------
'��ʱ����
Dim rsTemp As New ADODB.Recordset
Dim lngCount As Long

Dim rptCol As ReportColumn
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow


Private Sub Form_Activate()
    Err = 0: On Error Resume Next
    If Me.rptList.Rows.Count > 0 Then
        If Me.rptList.FocusedRow Is Nothing Then Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
    End If
    Me.rptList.SetFocus
End Sub

Private Sub Form_Load()
    mintPower = con_UnDefine
    Call zlGetPower
    '-----------------------------------------------------
    With Me.rptList
        Set rptCol = .Columns.Add(conColumn_ͼ��, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(conColumn_ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(conColumn_���, "���", 40, False): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(conColumn_����, "����", 110, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(conColumn_˵��, "˵��", 200, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(conColumn_����, "����", 70, True): rptCol.Editable = False: rptCol.Groupable = True
        Set rptCol = .Columns.Add(conColumn_��Ա, "������", 50, False): rptCol.Editable = False: rptCol.Groupable = True
        
        .SetImageList Me.imgList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With
End Sub

Private Sub Form_Resize()
    With Me.rptList
        .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
        .Top = Me.ScaleTop: .Height = Me.ScaleHeight
    End With
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.rptList.Visible = False Then Exit Sub
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    If Me.rptList.FocusedRow.GroupRow Then Exit Sub
    Call rptList_RowDblClick(Me.rptList.FocusedRow, Me.rptList.FocusedRow.Record.Item(conColumn_���))
End Sub

Private Sub rptList_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Call Form_Activate
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If Button <> vbRightButton Then Exit Sub
    RaiseEvent RightMouseUp(X, Y)
End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    RaiseEvent RowDblClick(Row)
End Sub

Private Sub rptList_SelectionChanged()
    If Me.rptList.Visible = False Then Exit Sub
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    If Me.Tag = "" And (Val(Me.rptList.Tag) <> Me.rptList.FocusedRow.Index Or Me.rptList.Tag = "") Then
        RaiseEvent SelRowChanged(Me.rptList.FocusedRow)
        Me.rptList.Tag = Me.rptList.FocusedRow.Index
    End If
End Sub

'-----------------------------------------------------
'���幫������
'-----------------------------------------------------

Public Function zlRefresh(ByVal lngFileID As Long, Optional ByVal lngRecId As Long) As Long
    '���ܣ�ˢ��װ��ָ���ļ��ķ���Ŀ¼
    '������ lngFileId���ļ�ID
    '       lngRecId����Ҫ��λ���ķ���
    '���أ�ˢ��װ��ķ�����Ŀ
    Me.Tag = "zlRefresh"
    Me.rptList.Tag = ""
    mlngFileID = lngFileID
    Err = 0: On Error GoTo ErrHand
    Select Case mintPower
    Case 0
        gstrSQL = "Select l.Id, l.���, l.����, l.˵��, l.ͨ�ü�, d.���� As ����, p.���� As ��Ա " _
                & "From ��������Ŀ¼ l, ���ű� d, ��Ա�� p " _
                & "Where l.����id = d.Id And l.��Աid = p.Id And l.�ļ�id =[1] " _
                & "Order By l.���"
    Case 1
        gstrSQL = "Select l.Id, l.���, l.����, l.˵��, l.ͨ�ü�, d.���� As ����, p.���� As ��Ա " _
                & "From ��������Ŀ¼ l, ���ű� d, ��Ա�� p " _
                & "Where l.����id = d.Id(+) And l.��Աid = p.Id(+) And l.�ļ�id =[1] And " _
                & "      (Nvl(l.ͨ�ü�, 0) = 0 Or " _
                & "      l.ͨ�ü� in (1,2) And l.����id In (Select r.����id From ������Ա r, �ϻ���Ա�� u Where r.��Աid = u.��Աid And u.�û��� = User)) " _
                & "Order By l.���"
    Case Else
        gstrSQL = "Select l.Id, l.���, l.����, l.˵��, l.ͨ�ü�, d.���� As ����, p.���� As ��Ա " _
                & "From ��������Ŀ¼ l, ���ű� d, ��Ա�� p " _
                & "Where l.����id = d.Id(+) And l.��Աid = p.Id(+) And l.�ļ�id =[1]  And " _
                & "      (Nvl(l.ͨ�ü�, 0) = 0 Or " _
                & "      l.ͨ�ü� =1 And l.����id In (Select r.����id From ������Ա r, �ϻ���Ա�� u Where r.��Աid = u.��Աid And u.�û��� = User) Or " _
                & "      l.ͨ�ü� =2 And l.��Աid In (Select u.��Աid From �ϻ���Ա�� u Where u.�û��� = User)) " _
                & "Order By l.���"
    End Select
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
    
    Me.rptList.Records.DeleteAll
    Do While Not rsTemp.EOF
        Set rptRcd = Me.rptList.Records.Add()
        Set rptItem = rptRcd.AddItem(CStr(IIf(IsNull(rsTemp!ͨ�ü�), 0, rsTemp!ͨ�ü�)))
        rptItem.Icon = rptItem.Value: rptItem.GroupPriority = rptItem.Value: rptItem.SortPriority = rptItem.Value
        Select Case rptItem.Value
        Case 0: rptItem.GroupCaption = "1-ȫԺ"
        Case 1: rptItem.GroupCaption = "2-����"
        Case Else: rptItem.GroupCaption = "3-����"
        End Select
        rptRcd.AddItem CStr(rsTemp!ID)
        rptRcd.AddItem Val(CStr(rsTemp!���))
        rptRcd.AddItem CStr(rsTemp!����)
        rptRcd.AddItem CStr("" & rsTemp!˵��)
        rptRcd.AddItem CStr("" & rsTemp!����)
        rptRcd.AddItem CStr("" & rsTemp!��Ա)
        rsTemp.MoveNext
    Loop
    Me.rptList.Populate
    
    If Me.rptList.Rows.Count > 0 Then
        If lngRecId <> 0 Then
            For Each rptRow In Me.rptList.Rows
                Set Me.rptList.FocusedRow = rptRow
            Next
        End If
        If Me.rptList.FocusedRow Is Nothing Then Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
    End If
    Me.Tag = ""
    Call rptList_SelectionChanged
    zlRefresh = Me.rptList.Records.Count
    
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = Me.rptList.Records.Count
End Function

Public Function zlGetPower() As Integer
    '���ܣ���õ�ǰ�û��ķ��Ĺ����Ȩ��
    '���أ��ʾ����Ȩ����ֵ
    Dim strPrivs As String
    If mintPower = con_UnDefine Then
        strPrivs = GetPrivFunc(glngSys, 1070)
        If InStr(1, strPrivs, "ȫԺ��������") <> 0 Then
            mintPower = 0
        ElseIf InStr(1, strPrivs, "���Ҳ�������") <> 0 Then
            mintPower = 1
        ElseIf InStr(1, strPrivs, "���˲�������") <> 0 Then
            mintPower = 2
        Else
            mintPower = -1
        End If
    End If
    zlGetPower = mintPower
End Function

Public Function zlGetFocusedRow() As XtremeReportControl.IReportRow
    '���ܣ���ȡ��ǰѡ����
    If Me.rptList.FocusedRow Is Nothing Then
        Set zlGetFocusedRow = Nothing
    Else
        Set zlGetFocusedRow = Me.rptList.FocusedRow
    End If
End Function

Public Function zlRecordCount() As Long
    '����:���ص�ǰ�ļ�¼����
    zlRecordCount = Me.rptList.Records.Count
End Function

Public Function zlRecordNew(ByVal frmParent As Object) As Long
    '���ܣ������µķ���
    Dim lngItemId As Long
    If frmParent Is Nothing Then Set frmParent = Me
    lngItemId = frmEPRModelEdit.ShowMe(frmParent, True, CByte(mintPower), mlngFileID)
    If lngItemId <> 0 Then Call Me.zlRefresh(mlngFileID, lngItemId)
    zlRecordNew = lngItemId
    DoEvents: Call Form_Activate
End Function

Public Function zlRecordModify(ByVal frmParent As Object) As Long
    '���ܣ���ǰ�����޸�
    Dim lngItemId As Long
    If Me.rptList.FocusedRow Is Nothing Then zlRecordModify = 0: Exit Function
    If Me.rptList.FocusedRow.GroupRow = True Then zlRecordModify = 0: Exit Function
    lngItemId = Me.rptList.FocusedRow.Record.Item(conColumn_ID).Value
    If frmParent Is Nothing Then Set frmParent = Me
    
    lngItemId = frmEPRModelEdit.ShowMe(frmParent, False, CByte(mintPower), mlngFileID, lngItemId)
    If lngItemId <> 0 Then Call Me.zlRefresh(mlngFileID, lngItemId)
    zlRecordModify = lngItemId
    DoEvents: Call Form_Activate
End Function

Public Function zlRecordDelete() As Boolean
    '���ܣ���ǰ����ɾ��
    Dim lngIndex As Long, lngItemId As Long
    
    With Me.rptList
        If .FocusedRow Is Nothing Then Exit Function
        If .FocusedRow.GroupRow Then Exit Function
        
        If MsgBox("���ɾ���÷�����" & vbCrLf & "����" & .FocusedRow.Record(conColumn_����).Value, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        gstrSQL = "zl_��������Ŀ¼_delete('" & .FocusedRow.Record(conColumn_ID).Value & "')"
        Err = 0: On Error GoTo ErrHand
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        Err = 0: On Error GoTo 0
        lngIndex = .FocusedRow.Record.Index
        Call .Records.RemoveAt(.FocusedRow.Record.Index)
        .Populate
        If .Records.Count <> 0 Then
            If lngIndex >= .Records.Count Then lngIndex = 0
            lngItemId = .Records(lngIndex).Item(conColumn_ID).Value
        Else
            lngItemId = 0
        End If
        Call Me.zlRefresh(mlngFileID, lngItemId)
    End With
    zlRecordDelete = True
    DoEvents: Call Form_Activate
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub zlRptPrint(ByVal bytMode As Byte)
    '����:��¼���ӡ
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL

    If Me.rptList.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '�������ݱ��
    If zlReportToVSFlexGrid(Me.vgdList, Me.rptList) = False Then Exit Sub
    
    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    Dim strSubhead As String

    Err = 0: On Error GoTo ErrHand
    gstrSQL = "Select ���� From �����ļ��б� Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    If rsTemp.RecordCount > 0 Then strSubhead = rsTemp!����
    
    Err = 0: On Error Resume Next
    Set objPrint.Body = Me.vgdList
    objPrint.Title.Text = strSubhead & "����Ŀ¼"
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
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
