VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatiCureCardConsumeMgr 
   BorderStyle     =   0  'None
   Caption         =   "Ԥ�����"
   ClientHeight    =   5055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "ˢ��(&R)"
      Height          =   350
      Left            =   9225
      TabIndex        =   6
      Top             =   150
      Width           =   1100
   End
   Begin VB.PictureBox picFilter 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   210
      ScaleHeight     =   465
      ScaleWidth      =   9360
      TabIndex        =   9
      Top             =   180
      Width           =   9360
      Begin VB.Frame fraType 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5475
         TabIndex        =   11
         Top             =   90
         Width           =   3600
         Begin VB.OptionButton optType 
            Caption         =   "����Ԥ��"
            Height          =   180
            Index           =   0
            Left            =   -15
            TabIndex        =   3
            Top             =   0
            Width           =   1035
         End
         Begin VB.OptionButton optType 
            Caption         =   "סԺԤ��"
            Height          =   180
            Index           =   1
            Left            =   1155
            TabIndex        =   4
            Top             =   0
            Width           =   1065
         End
         Begin VB.OptionButton optType 
            Caption         =   "�����סԺ"
            Height          =   180
            Index           =   2
            Left            =   2340
            TabIndex        =   5
            Top             =   0
            Value           =   -1  'True
            Width           =   1380
         End
      End
      Begin MSComCtl2.DTPicker dtp��ʼ���� 
         Height          =   315
         Left            =   1020
         TabIndex        =   1
         Top             =   30
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   103874563
         CurrentDate     =   40722
      End
      Begin MSComCtl2.DTPicker dtp�������� 
         Height          =   315
         Left            =   3315
         TabIndex        =   2
         Top             =   30
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   103874563
         CurrentDate     =   40722
      End
      Begin VB.Label lblEndDate 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   0
         Left            =   3195
         TabIndex        =   10
         Top             =   75
         Width           =   180
      End
      Begin VB.Label lblStartDate 
         AutoSize        =   -1  'True
         Caption         =   "ʱ�䷶Χ(&E)"
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   0
         Top             =   90
         Width           =   990
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   2250
      Left            =   195
      TabIndex        =   7
      Top             =   1320
      Width           =   5745
      _cx             =   10134
      _cy             =   3969
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   9
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPatiCureCardConsumeMgr.frx":0000
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
      ExplorerBar     =   7
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
      Begin VB.PictureBox picImg 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   60
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   8
         Top             =   45
         Width           =   210
         Begin VB.Image imgCol 
            Height          =   195
            Left            =   0
            Picture         =   "frmPatiCureCardConsumeMgr.frx":0107
            ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
   End
End
Attribute VB_Name = "frmPatiCureCardConsumeMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long, mblnHaveData As Boolean
Private mstrCardNo As String, mlngCardTypeID As Long, mlng����ID As Long
Public Event zlPopupMenus(ByVal vsGrid As VSFlexGrid) '�����˵�����
Public Event AfterRowChange(ByVal vsGrid As VSFlexGrid) '�����˵�����

Public Function zlReLoadData(ByVal lng����ID As Long, ByVal lngCardTypeID As Long, ByVal strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¼�������
    '����:���سɹ�,����true,���򷵻�Flase
    '����:���˺�
    '����:2011-06-28 15:30:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mstrCardNo = strCardNo: mlngCardTypeID = lngCardTypeID: mlng����ID = lng����ID
    Err = 0: On Error GoTo ErrHand:
    Call LoadDataToRpt
    zlReLoadData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub InitVsGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2011-06-28 15:31:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsGrid
        'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
        .ColData(.ColIndex("ID")) = "-1|1"
        .ColData(.ColIndex("����")) = "1|1"
        .ColAlignment(.ColIndex("��������")) = flexAlignRightCenter
    End With
End Sub

Private Sub LoadDataToRpt()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ݸ�����
    '����:���˺�
    '����:2009-09-07 11:53:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, i As Long, lngRow As Long
    Dim str��� As String, dbl��� As Double, strSQL As String
    Dim strTable As String, strHTable As String, strIF As String, blnDataMove As String
    
    str��� = ""
    If optType(2).value Then
        str��� = str��� & "1,2"
        strIF = "  And Nvl(a.Ԥ�����, 2) in (1,2)"
    ElseIf optType(1).value Then
        str��� = str��� & "2"
        strIF = "  And Nvl(a.Ԥ�����, 2)=2"
    Else
        str��� = str��� & "1"
        strIF = "  And Nvl(a.Ԥ�����, 2)=1"
    End If
    
    mblnHaveData = False
    Err = 0: On Error GoTo ErrHand:
    '98616:���ϴ�,2016/7/26,ͳ��Ԥ��ʹ�������Ҫ�ų���Ԥ��Ϊ0 �ļ�¼����ֱ����Ԥ����¼ͳ������Ϊ��ʷ����û������
    '84389:���ϴ�,2015/5/6,�൥��һ�ν���󣬲�����no���ַ��ó�Ԥ�������ý���id
    '����:50472
    strTable = "" & _
    " Select ����id, �տ�ʱ��, 0 As ����,����id, Nvl(���, 0) As ���, 0 As ��Ԥ�� " & _
    "  From ����Ԥ����¼ A " & _
    "  Where �տ�ʱ�� >= [2] And ��¼���� =1  And ����id =[1] " & strIF & _
    "  Union All " & _
    "  Select a.����id, b.�շ�ʱ�� As �տ�ʱ��,2 As ����, b.id as ����id, 0 As ���, Nvl(��Ԥ��, 0) As ��Ԥ�� " & _
    "  From ����Ԥ����¼ A, ���˽��ʼ�¼ B " & _
    "  Where b.�շ�ʱ�� >= [2] And Mod(a.��¼����, 10) = 1  And a.����id = b.Id And a.����id =[1] " & strIF & _
    "  Union All " & _
    "  Select ����id, �շ�ʱ��, 1 As  ����, ����id, 0 As ���, Nvl(Sum(��Ԥ��), 0) As ��Ԥ�� " & _
    "  From (Select a.����id, b.�Ǽ�ʱ�� As �շ�ʱ��, a.No As ��ֵ���ݺ�, b.����id, 0 As ���, Max(Nvl(a.��Ԥ��, 0)) As ��Ԥ�� " & _
    "         From ����Ԥ����¼ A, ������ü�¼ B " & _
    "         Where b.�Ǽ�ʱ�� >= [2] And Mod(a.��¼����, 10) =1 And Nvl(b.���ʷ���, 0) = 0 And a.����id = b.����id And b.����id =[1] And " & _
    "           b.��¼���� In (1, 4) And Nvl(a.��Ԥ��, 0) <> 0 " & strIF & _
    "         Group By a.����id, b.�Ǽ�ʱ��, a.No, b.����id) " & _
    "  Group By ����id, �շ�ʱ��, ����id " & _
    "  Union All " & _
    "  Select ����id, �շ�ʱ��, 1 as ����, ����id, 0 As ���, Nvl(Sum(��Ԥ��), 0) As ��Ԥ�� " & _
    "  From (Select a.����id, b.�Ǽ�ʱ�� As �շ�ʱ��, a.No As ��ֵ���ݺ�, b.����id, 0 As ���, Max(Nvl(a.��Ԥ��, 0)) As ��Ԥ�� " & _
    "         From ����Ԥ����¼ A, סԺ���ü�¼ B " & _
    "         Where b.�Ǽ�ʱ�� >= [2] And Mod(a.��¼����, 10) =1 And a.����id = b.����id And b.����id =[1] And b.��¼���� = 5 And " & _
    "           Nvl(b.���ʷ���, 0) = 0 And Nvl(a.��Ԥ��, 0) <> 0 " & _
    "         Group By a.����id, b.�Ǽ�ʱ��, a.No, b.����id) " & _
    "  Group By ����id, �շ�ʱ��, ����id"
    blnDataMove = zlDatabase.DateMoved(Format(dtp��ʼ����.value, "yyyy-mm-dd"), , , Me.Caption)
    
    If blnDataMove Then
        strHTable = Replace(strTable, "����Ԥ����¼", "H����Ԥ����¼")
        strHTable = Replace(strHTable, "סԺ���ü�¼", "HסԺ���ü�¼")
        strHTable = Replace(strHTable, "������ü�¼", "H������ü�¼")
        strTable = strTable & " UNION ALL " & strHTable
    End If
    strSQL = " " & _
        "   Select /*+ RULE */ ���,�տ�ʱ��, ҵ������, Sum(�ڳ����) As �ڳ����, Sum(���ڳ�ֵ) As ���ڳ�ֵ, Sum(��������) As �������� " & _
        "   From (With Ԥ�� As ( " & strTable & ")" & _
        "          Select  0 as  ���,'' As �տ�ʱ��, '�ڳ�' As ҵ������, Sum(Nvl(Ԥ�����, 0)) As �ڳ����, 0 As ���ڳ�ֵ, 0 As �������� " & _
        "          From �������  A" & _
        "          Where ����id = [1] And ���� = 1 " & Replace(strIF, "Ԥ�����", "����") & _
        "          Union All " & _
        "          Select 0 as ���,'' As �տ�ʱ��, '�ڳ�' As ҵ������, -1 * Sum(Nvl(���, 0)) + Sum(Nvl(��Ԥ��, 0)) As �ڳ����, 0,  0 As �������� " & _
        "          From Ԥ�� " & _
        "          Where  �տ�ʱ�� >= [2] " & _
        "          Group By To_Char(�տ�ʱ��, 'yyyy-mm-dd') " & _
        "          Union All " & _
        "          Select 1 as ���,To_Char(�տ�ʱ��, 'yyyy-mm-dd') As �տ�ʱ��, '��ֵ' As ҵ������, 0 As �ڳ����, Sum(Nvl(���, 0)) As ��ֵ, 0 As �������� " & _
        "          From Ԥ�� " & _
        "          Where  �տ�ʱ�� Between [2] And [3] " & _
        "          Having Sum(Nvl(���, 0))<>0 " & _
        "          Group By To_Char(�տ�ʱ��, 'yyyy-mm-dd')" & _
        "         Union All " & _
        "          Select 1 as ���, To_Char(�տ�ʱ��, 'yyyy-mm-dd') As �տ�ʱ��, decode(����,1,'�շ�',2,'����','����') As ҵ������, 0 As �ڳ����, 0 As ��ֵ, " & _
        "                 Sum(Nvl(��Ԥ��, 0)) As ���� " & _
        "          From Ԥ�� " & _
        "          Where  �տ�ʱ�� Between [2] And [3] " & _
        "           Having Sum(Nvl(��Ԥ��, 0))<>0  Group By To_Char(�տ�ʱ��, 'yyyy-mm-dd') ,Decode(����, 1, '�շ�', 2, '����', '����')) " & _
        "          Group By  ���,�տ�ʱ��, ҵ������" & _
        "          Order By ���,�տ�ʱ�� "
        
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, dtp��ʼ����.value, dtp��������.value)
    With Me.vsGrid
        .Redraw = flexRDNone
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        .Clear 1
        .Row = 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("�տ�ʱ��")) = Nvl(rsTemp!�տ�ʱ��)
            .TextMatrix(lngRow, .ColIndex("ҵ������")) = Nvl(rsTemp!ҵ������)
            .TextMatrix(lngRow, .ColIndex("�ڳ����")) = Format(Val(Nvl(rsTemp!�ڳ����)), "####0.00;-###0.00; ;")
            .TextMatrix(lngRow, .ColIndex("���ڳ�ֵ")) = Format(Val(Nvl(rsTemp!���ڳ�ֵ)), "####0.00;-###0.00; ;")
            .TextMatrix(lngRow, .ColIndex("��������")) = Format(Val(Nvl(rsTemp!��������)), "####0.00;-###0.00; ;")
            dbl��� = dbl��� + Val(Nvl(rsTemp!�ڳ����)) + Val(Nvl(rsTemp!���ڳ�ֵ)) - Val(Nvl(rsTemp!��������))
            .TextMatrix(lngRow, .ColIndex("��δ���")) = Format(dbl���, "####0.00;-###0.00;;")
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        Call InitVsGrid
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        '�ָ�������
        zl_vsGrid_Para_Restore mlngModule, vsGrid, Me.Caption, "�ʻ�����б�", True
        .ColWidth(.ColIndex("��־")) = 285
        .ColAlignment(.ColIndex("��־")) = flexAlignCenterCenter
        .Redraw = flexRDBuffered
    End With
    mblnHaveData = rsTemp.RecordCount > 0
   Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
     Me.vsGrid.Redraw = flexRDBuffered
End Sub

Private Sub cmdRefresh_Click()
    Call LoadDataToRpt
End Sub

Private Sub Form_Load()
    mlngModule = glngModul: mstrPrivs = gstrPrivs
    dtp��������.MaxDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59")
    dtp��������.value = Format(dtp��������.MaxDate, "yyyy-mm-dd 23:59:59")
    dtp��ʼ����.MaxDate = dtp��������.MaxDate
    dtp��ʼ����.value = Format(DateAdd("m", -1, dtp��ʼ����.MaxDate), "yyyy-mm-dd 00:00:00")
    Call InitVsGrid
    Call vsGrid_GotFocus
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single
    Err = 0: On Error Resume Next
    If Me.ScaleWidth < 10455 Then
        fraType.Top = dtp��������.Top + dtp��������.Height + 120
        picFilter.Height = 445 + dtp��������.Height
        fraType.Left = dtp��ʼ����.Left
        cmdRefresh.Top = picFilter.Top + picFilter.Height - cmdRefresh.Height - 50
    Else
        fraType.Top = dtp��������.Top + (dtp��������.Height - fraType.Height) \ 2
        fraType.Left = dtp��������.Left + dtp��������.Width + 100
        picFilter.Height = 465
        cmdRefresh.Top = picFilter.Top + dtp��������.Top
    End If
    cmdRefresh.Left = Me.ScaleWidth - cmdRefresh.Width - 100
    With vsGrid
        .Left = ScaleLeft: .Top = picFilter.Top + picFilter.Height
        .Width = ScaleWidth: .Height = ScaleHeight - .Top
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "�ʻ�����б�", True
End Sub
Private Sub imgCol_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlcontrol.GetControlRect(picImg.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImg.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsGrid, lngLeft, lngTop, imgCol.Height)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "�ʻ�����б�", True
End Sub

 
Private Sub optType_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Private Sub picImg_Click()
    Call imgCol_Click
End Sub
 
Public Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���д�ӡ,Ԥ���������EXCEL
    '���:bytFunc=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '����:���˺�
    '����:2011-06-28 15:59:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim rsTemp As New ADODB.Recordset, vsGrid As VSFlexGrid
    Err = 0: On Error GoTo errH:
    gstrSQL = "Select   A.����,A.�Ա�,A.���� From ������Ϣ A where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID)
    If rsTemp.EOF = True Then Exit Sub '�޿���Ϣ���˳�
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
        
    objPrint.Title.Text = gstrUnitName & "�ʻ�������"
    
    objRow.Add "������" & Nvl(rsTemp!����)
    objRow.Add "���䣺" & Nvl(rsTemp!����)
    objRow.Add "�Ա�" & Nvl(rsTemp!�Ա�)
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    Err = 0: On Error GoTo ErrHand:
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .Cell(flexcpData, 0, intCol) = .ColWidth(intCol)
            If .ColHidden(intCol) Or intCol = .ColIndex("��־") Then .ColWidth(intCol) = 0
        Next
    End With
    Set objPrint.Body = vsGrid
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
    
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .ColWidth(intCol) = Val(.Cell(flexcpData, 0, intCol))
        Next
        .Redraw = flexRDBuffered
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .ColWidth(intCol) = Val(.Cell(flexcpData, 0, intCol))
        Next
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub vsGrid_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "�ʻ�����б�", True
End Sub

Private Sub vsGrid_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsGrid, OldRow, NewRow, OldCol, NewCol, gSysColor.lngGridColorSel
    If OldRow <> NewRow Then
        RaiseEvent AfterRowChange(vsGrid)
    End If
End Sub
Private Sub vsGrid_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "�ʻ�����б�", True
End Sub

 Private Sub vsGrid_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsGrid
        If Col = .ColIndex("��־") Then Cancel = True
    End With
End Sub
Private Sub vsGrid_GotFocus()
    zl_VsGridGotFocus vsGrid, gSysColor.lngGridColorSel
End Sub
Private Sub vsGrid_LostFocus()
    zl_VsGridLOSTFOCUS vsGrid, gSysColor.lngGridColorLost
End Sub

Private Sub vsGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button <> vbRightButton Then Exit Sub
    RaiseEvent zlPopupMenus(vsGrid)
End Sub
Private Sub dtp��������_Change()
     If dtp��������.value > dtp��ʼ����.MaxDate Then dtp��������.value = dtp��ʼ����.MaxDate
    If dtp��������.value < dtp��ʼ����.value Then
        dtp��ʼ����.value = dtp��������.value
    End If
End Sub
Private Sub dtp��������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub dtp��ʼ����_Change()
    If dtp��ʼ����.value > dtp��������.MaxDate Then dtp��ʼ����.value = dtp��������.MaxDate
    If dtp��������.value < dtp��ʼ����.value Then
        dtp��������.value = dtp��ʼ����.value
    End If
End Sub
Private Sub dtp��ʼ����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Public Function zlShowReport(lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҽ����֧��ᱨ��
    '���:lng����ID ����ID��
    '����:����
    '����:2012-06-12 15:59:59
    '����50122
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str��� As String
    Dim strDate As String
    
    strDate = vsGrid.TextMatrix(vsGrid.Row, vsGrid.ColIndex("�տ�ʱ��"))
    
    str��� = ""
    If optType(2).value Then
        str��� = "3"
    ElseIf optType(1).value Then
        str��� = "2"
    Else
        str��� = "1"
    End If
     If vsGrid.Row >= 2 Then Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1107_2", Me, "����ID=" & lng����ID, "����=" & CDate(strDate), "Ԥ�����=" & str���)
End Function
