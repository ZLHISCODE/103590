VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Begin VB.Form frmPersonOutPayEdit 
   BorderStyle     =   0  'None
   Caption         =   "�������"
   ClientHeight    =   7545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.ImageList ils16 
      Left            =   7170
      Top             =   1005
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
            Picture         =   "frmPersonOutPayEdit.frx":0000
            Key             =   "Cancel"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonOutPayEdit.frx":059A
            Key             =   "OutPay"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonOutPayEdit.frx":0B34
            Key             =   "Requisition"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonOutPayEdit.frx":10CE
            Key             =   "Out"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils24 
      Left            =   5805
      Top             =   735
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonOutPayEdit.frx":1668
            Key             =   "Requisition"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonOutPayEdit.frx":1D62
            Key             =   "OutPay"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonOutPayEdit.frx":245C
            Key             =   "Cancel"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picOutPay 
      BorderStyle     =   0  'None
      Height          =   3330
      Left            =   4005
      ScaleHeight     =   3330
      ScaleWidth      =   3555
      TabIndex        =   1
      Top             =   1620
      Width           =   3555
      Begin VSFlex8Ctl.VSFlexGrid vsOutPay 
         Height          =   2145
         Left            =   135
         TabIndex        =   2
         Top             =   585
         Width           =   2895
         _cx             =   5106
         _cy             =   3784
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   9
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPersonOutPayEdit.frx":2B56
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
         Begin VB.PictureBox picImgOutPay 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   30
            ScaleHeight     =   225
            ScaleWidth      =   210
            TabIndex        =   4
            Top             =   60
            Width           =   210
            Begin VB.Image imgColSel 
               Height          =   195
               Left            =   0
               Picture         =   "frmPersonOutPayEdit.frx":2CAC
               ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
      End
   End
   Begin VB.PictureBox picLoanRequisition 
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   75
      ScaleHeight     =   4695
      ScaleWidth      =   3450
      TabIndex        =   0
      Top             =   885
      Width           =   3450
      Begin VSFlex8Ctl.VSFlexGrid vsRequisition 
         Height          =   2145
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   2895
         _cx             =   5106
         _cy             =   3784
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   9
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPersonOutPayEdit.frx":31FA
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
         Begin VB.PictureBox picImgRequisition 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   45
            ScaleHeight     =   225
            ScaleWidth      =   210
            TabIndex        =   5
            Top             =   60
            Width           =   210
            Begin VB.Image imgColRequisition 
               Height          =   195
               Left            =   0
               Picture         =   "frmPersonOutPayEdit.frx":32E3
               ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   135
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmPersonOutPayEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar
Private mstrPrivs As String, mlngModule As Long
Private mArrFilter As Variant   '��������
Private mcbsThis As Object
Private Const conPane_Requisition = 0
Private Const conPane_OutPay = 1

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
    Call LoadRequisition
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
    '����:2009-09-09 15:45:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsRequisition
        'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
        .ColData(.ColIndex("ID")) = "-1|1"
        .ColHidden(.ColIndex("ID")) = True
        .ColData(.ColIndex("��־")) = "-1|1"
        .ColData(.ColIndex("�����")) = "1|0"
        .ColData(.ColIndex("�����")) = "1|0"
        .ColData(.ColIndex("����ʱ��")) = "1|0"
    End With
    With vsOutPay
        .ColHidden(.ColIndex("ID")) = True
        .ColData(.ColIndex("ID")) = "-1|1"
        .ColData(.ColIndex("��־")) = "-1|1"
        .ColData(.ColIndex("�����")) = "1|0"
        .ColData(.ColIndex("�����")) = "1|0"
        .ColData(.ColIndex("����ʱ��")) = "1|0"
    End With
    
End Sub
Private Sub LoadDataToRpt()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ݸ�����
    '����:���˺�
    '����:2009-09-07 11:53:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFilter As String, rsTemp As New ADODB.Recordset, lngRow As Long
    Dim blnHistory As Boolean, strStartDate As String
    
    
    Err = 0: On Error GoTo ErrHand:
    zlCommFun.ShowFlash "����װ�ؽ������,���Ժ�..."
    strStartDate = "3000-01-01 00:00:00"
    If strStartDate > CStr(mArrFilter("���-����ʱ��")(0)) And CStr(mArrFilter("���-����ʱ��")(0)) <> "1901-01-01" Then
        strStartDate = CStr(mArrFilter("���-����ʱ��")(0))
    End If
    If strStartDate > CStr(mArrFilter("���-���ʱ��")(0)) And CStr(mArrFilter("���-���ʱ��")(0)) <> "1901-01-01" Then
        strStartDate = CStr(mArrFilter("���-���ʱ��")(0))
    End If
    If strStartDate > CStr(mArrFilter("���-ȡ��ʱ��")(0)) And CStr(mArrFilter("���-ȡ��ʱ��")(0)) <> "1901-01-01" Then
        strStartDate = CStr(mArrFilter("���-ȡ��ʱ��")(0))
    End If
    
    If strStartDate <> "3000-01-01 00:00:00" Then blnHistory = zlDatabase.DateMoved(strStartDate, , , Me.Caption)


    If CStr(mArrFilter("���-����ʱ��")(0)) <> "1901-01-01" And CStr(mArrFilter("���-���ʱ��")(0)) <> "1901-01-01" And CStr(mArrFilter("���-ȡ��ʱ��")(0)) <> "1901-01-01" Then
        strFilter = "   (����ʱ�� between [1] and [2] or ���ʱ�� between [3] and [4] or ȡ��ʱ�� between [5] and [6]) and ���ʱ�� is not Null   "
         
    ElseIf CStr(mArrFilter("���-����ʱ��")(0)) <> "1901-01-01" And CStr(mArrFilter("���-���ʱ��")(0)) <> "1901-01-01" And CStr(mArrFilter("���-ȡ��ʱ��")(0)) = "1901-01-01" Then
        strFilter = "   (����ʱ�� between [1] and [2] or ���ʱ�� between [3] and [4] and ���ʱ�� is not Null )   "
    ElseIf CStr(mArrFilter("���-����ʱ��")(0)) <> "1901-01-01" And CStr(mArrFilter("���-���ʱ��")(0)) = "1901-01-01" And CStr(mArrFilter("���-ȡ��ʱ��")(0)) <> "1901-01-01" Then
        strFilter = "   (����ʱ�� between [1] and [2] or ȡ��ʱ�� between [5] and [6]) and ���ʱ�� is not Null   "
    ElseIf CStr(mArrFilter("���-����ʱ��")(0)) = "1901-01-01" And CStr(mArrFilter("���-���ʱ��")(0)) <> "1901-01-01" And CStr(mArrFilter("���-ȡ��ʱ��")(0)) <> "1901-01-01" Then
        strFilter = "   ( ���ʱ�� between [3] and [4] or ȡ��ʱ�� between [5] and [6]) and ���ʱ�� is not Null   "
    ElseIf CStr(mArrFilter("���-����ʱ��")(0)) <> "1901-01-01" Then
        strFilter = "   (����ʱ�� between [1] and [2]   ) and ���ʱ�� is not Null "
    ElseIf CStr(mArrFilter("���-���ʱ��")(0)) <> "1901-01-01" Then
        strFilter = "   (���ʱ�� between [3] and [4])"
    Else
        strFilter = "   (ȡ��ʱ�� between [5] and [6] )"
    End If
    
    If CStr(mArrFilter("�����")) <> "" Then strFilter = strFilter & " and ����� like [7]"
    strFilter = strFilter & " and ����� like [8]"
    
    gstrSQL = " " & _
    "    Select distinct  A.Id, A.�����, A.��ע, A.�����, to_char(A.����ʱ��,'yyyy-mm-dd hh24:mi:ss') as ����ʱ�� ,  " & _
    "           A.�����, to_char(A.���ʱ��,'yyyy-mm-dd hh24:mi:ss') as ���ʱ��, " & _
    "           to_char(A.ȡ��ʱ��,'yyyy-mm-dd hh24:mi:ss') as ȡ��ʱ��, A.ȡ��ԭ��,Decode(B.��¼ID,NULL,0,1) as �ѽɿ�" & _
    "    From ��Ա����¼ A,��Ա�սɶ��� B " & _
    "    Where A.ID=B.��¼ID(+) and B.����(+)=4 And " & strFilter
    If blnHistory Then
        gstrSQL = gstrSQL & vbCrLf & " Union ALL " & Replace(Replace(gstrSQL, "��Ա����¼", "H��Ա����¼"), "Decode(B.��¼ID,NULL,0,1)", " 2 ") & vbCrLf
    End If
    gstrSQL = gstrSQL & _
    "    Order by   ���ʱ��, ����� "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
        CDate(mArrFilter("���-����ʱ��")(0)), CDate(mArrFilter("���-����ʱ��")(1)), _
        CDate(mArrFilter("���-���ʱ��")(0)), CDate(mArrFilter("���-���ʱ��")(1)), _
        CDate(mArrFilter("���-ȡ��ʱ��")(0)), CDate(mArrFilter("���-ȡ��ʱ��")(1)), _
        CStr(mArrFilter("�����")), UserInfo.����)
    
    With Me.vsOutPay
        .Clear 1
        .Rows = 2: lngRow = 1
        .Redraw = flexRDNone
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ID")) = Nvl(rsTemp!ID)
            .TextMatrix(lngRow, .ColIndex("�����")) = Nvl(rsTemp!�����)
            .TextMatrix(lngRow, .ColIndex("�����")) = Format(Val(Nvl(rsTemp!�����)), "####0.00;-####0.00;;")
            .TextMatrix(lngRow, .ColIndex("����ʱ��")) = Nvl(rsTemp!����ʱ��)
            .TextMatrix(lngRow, .ColIndex("��ע")) = Nvl(rsTemp!��ע)
            .TextMatrix(lngRow, .ColIndex("�����")) = Nvl(rsTemp!�����)
            .TextMatrix(lngRow, .ColIndex("���ʱ��")) = Nvl(rsTemp!���ʱ��)
            .TextMatrix(lngRow, .ColIndex("ȡ��ʱ��")) = Nvl(rsTemp!ȡ��ʱ��)
            .TextMatrix(lngRow, .ColIndex("ȡ��ԭ��")) = Nvl(rsTemp!ȡ��ԭ��)
            If Nvl(rsTemp!ȡ��ʱ��) <> "" Then
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbRed
                .Cell(flexcpPicture, lngRow, .ColIndex("�����")) = ils16.ListImages("Cancel").Picture
            Else
                .Cell(flexcpPicture, lngRow, .ColIndex("�����")) = ils16.ListImages("OutPay").Picture
            End If
            If Val(Nvl(rsTemp!�ѽɿ�)) = 1 Then
                .Cell(flexcpPicture, lngRow, .ColIndex("��־")) = ils16.ListImages("Out").Picture
            ElseIf Val(Nvl(rsTemp!�ѽɿ�)) = 2 Then
                '�Ѿ�ת����ʷ����
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = &H8000000F
            End If
            .Cell(flexcpData, lngRow, .ColIndex("��־")) = Val(Nvl(rsTemp!�ѽɿ�))
            lngRow = lngRow + 1
           rsTemp.MoveNext
        Loop
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .Redraw = flexRDBuffered
        '�ָ�������
        zl_vsGrid_Para_Restore mlngModule, vsOutPay, "������", "����б�", True
        .ColWidth(.ColIndex("��־")) = 285
    End With
    zlCommFun.StopFlash
   Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    zlCommFun.StopFlash
     Me.vsOutPay.Redraw = flexRDBuffered
End Sub
Private Sub LoadRequisition()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������������
    '����:���˺�
    '����:2009-09-07 11:53:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, lngRow As Long
    
    Err = 0: On Error GoTo ErrHand:
    
    zlCommFun.ShowFlash "����װ�ؽ������,���Ժ�..."
    
    gstrSQL = " " & _
    "    Select Id, �����, ��ע, �����, to_char(����ʱ��,'yyyy-mm-dd hh24:mi:ss') as ����ʱ�� ,  " & _
    "           �����, to_char(���ʱ��,'yyyy-mm-dd hh24:mi:ss') as ���ʱ��, " & _
    "           to_char(ȡ��ʱ��,'yyyy-mm-dd hh24:mi:ss') as ȡ��ʱ��, ȡ��ԭ��" & _
    "    From ��Ա����¼ " & _
    "    Where ���ʱ�� is null and �����=[1] " & _
    "    Order by  ����ʱ��,����� "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UserInfo.����)
    
    With Me.vsRequisition
        .Clear 1
        .Rows = 2: .Redraw = flexRDNone
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ID")) = Nvl(rsTemp!ID)
            .TextMatrix(lngRow, .ColIndex("�����")) = Nvl(rsTemp!�����)
            .TextMatrix(lngRow, .ColIndex("�����")) = Format(Val(Nvl(rsTemp!�����)), "####0.00;-####0.00;;")
            .TextMatrix(lngRow, .ColIndex("����ʱ��")) = Nvl(rsTemp!����ʱ��)
            .TextMatrix(lngRow, .ColIndex("��ע")) = Nvl(rsTemp!��ע)
            .TextMatrix(lngRow, .ColIndex("�����")) = Nvl(rsTemp!�����)
            .Cell(flexcpPicture, lngRow, .ColIndex("�����")) = ils16.ListImages("Requisition").Picture
            lngRow = lngRow + 1
           rsTemp.MoveNext
        Loop
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        '�ָ�������
         zl_vsGrid_Para_Restore mlngModule, vsRequisition, "������", "�����б�", True
        .ColWidth(.ColIndex("��־")) = 285
        .Redraw = flexRDBuffered
    End With
    zlCommFun.StopFlash
   Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    zlCommFun.StopFlash
     Me.vsRequisition.Redraw = flexRDBuffered
End Sub
Private Sub InitPancel()
    Dim sngWidth As Single
    Dim strReg As String
    Dim panThis As Pane
    
    Set panThis = dkpMan.CreatePane(conPane_Requisition, 200, 580, DockLeftOf, Nothing)
    panThis.Title = "���������Ϣ"
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panThis.Tag = picLoanRequisition
    
    Set panThis = dkpMan.CreatePane(conPane_OutPay, 250, 580, DockRightOf, Nothing)
    panThis.Title = "�����Ϣ"
    panThis.Tag = conPane_OutPay
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.HideClient = True
    zlRestoreDockPanceToReg Me, dkpMan, "����"
End Sub
Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_Requisition
        Item.Handle = picLoanRequisition.hwnd
    Case conPane_OutPay
        Item.Handle = Me.picOutPay.hwnd
    End Select
End Sub
Private Sub Form_Load()
    mlngModule = glngModul: mstrPrivs = gstrPrivs
    
    Call InitPancel
    Call InitVsGrid
    Call vsOutPay_LostFocus: Call vsRequisition_LostFocus
    vsRequisition_GotFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zlSaveDockPanceToReg Me, dkpMan, "����"
    zl_vsGrid_Para_Save mlngModule, vsOutPay, "������", "����б�", True
    zl_vsGrid_Para_Save mlngModule, vsRequisition, "������", "�����б�", True
End Sub

Private Sub imgColRequisition_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlcontrol.GetControlRect(picImgRequisition.hwnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgRequisition.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsRequisition, lngLeft, lngTop, imgColSel.Height)
    zl_vsGrid_Para_Save mlngModule, vsRequisition, Me.Name, "�����б�", True
End Sub

Private Sub imgColSel_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlcontrol.GetControlRect(picImgOutPay.hwnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgOutPay.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsOutPay, lngLeft, lngTop, imgColSel.Height)
    zl_vsGrid_Para_Save mlngModule, vsOutPay, Me.Name, "����б�", True
End Sub

Private Sub picImgOutPay_Click()
    Call imgColSel_Click
End Sub

Private Sub picLoanRequisition_Resize()
    Err = 0: On Error Resume Next
    With picLoanRequisition
        vsRequisition.Left = .ScaleLeft
        vsRequisition.Width = .ScaleWidth
        vsRequisition.Top = .ScaleTop
        vsRequisition.Height = .ScaleHeight
    End With
End Sub
Private Sub picOutPay_Resize()
    Err = 0: On Error Resume Next
    With picOutPay
        vsOutPay.Left = .ScaleLeft
        vsOutPay.Width = .ScaleWidth
        vsOutPay.Top = .ScaleTop
        vsOutPay.Height = .ScaleHeight
    End With
End Sub


Public Function zlDefCommandBars(ByVal cbsThis As Object) As Boolean
    '----------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008/1/9
    '----------------------------------------------------------------------------------------
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
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "ȷ�Ͻ��(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ChargeOff, "ȡ�����(&M)")
    End With

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    mcbrMenuBar.ID = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
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
        .Add FCONTROL, Asc("A"), conMenu_Edit_Audit
        .Add FCONTROL, Asc("M"), conMenu_Edit_ChargeOff
        .Add FSHIFT, VK_DELETE, conMenu_Edit_Delete
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
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "ȷ�Ͻ��"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ChargeOff, "ȡ�����")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
     zlDefCommandBars = True
     zlcontrol.ControlSetFocus vsRequisition
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub ExcuteFunction(Optional ByVal blnOutPay As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ȷ�Ͻ��
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-09-09 12:04:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long
    
    If blnOutPay Then
        If zlStr.IsHavePrivs(mstrPrivs, "���ȷ��") = False Then Exit Sub
        With vsRequisition
            If .Row < 0 Or .Row > .Rows - 1 Then Exit Sub
            lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        End With
    Else
        If zlStr.IsHavePrivs(mstrPrivs, "ȡ�����") = False Then Exit Sub
        With vsOutPay
            If .Row < 0 Or .Row > .Rows - 1 Then Exit Sub
            If Trim(.TextMatrix(.Row, .ColIndex("ȡ��ʱ��"))) <> "" Then Exit Sub
            lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        End With
    End If
    If lngID = 0 Then Exit Sub
    If frmPersonLoanRequisitionEdit.ShowEdit(Me, IIf(blnOutPay, FN_���, FN_ȡ�����), mstrPrivs, mlngModule, lngID) = False Then Exit Sub
    '����ˢ������
    Call zlReLoadData(mArrFilter)
End Sub
'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------
Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRow As Long
    Dim lngID  As Long
    '------------------------------------
    Select Case Control.ID
    'bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Case conMenu_File_Preview: Call zlRptPrint(2)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_Edit_Audit   'ȷ�Ͻ��
        Call ExcuteFunction(True)
    Case conMenu_Edit_ChargeOff    'ȡ�����
        Call ExcuteFunction(False)
    Case conMenu_View_Refresh   'ˢ��
        '����ˢ������
        Call zlReLoadData(mArrFilter)
    Case Else
        If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            'ִ�з�������ǰģ��ı���
            If Me.ActiveControl Is vsRequisition Then
                With vsOutPay
                        lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
                End With
            Else
                With vsOutPay
                        lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
                End With
            End If
            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "ID=" & lngID)
        End If
    End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub

Private Function HaveData() As Boolean
    '����:�Ƿ�������
    If Me.ActiveControl Is vsRequisition Then
        With Me.vsRequisition
            HaveData = Val(.TextMatrix(.Row, .ColIndex("ID"))) <> 0
        End With
    Else
        With Me.vsOutPay
            HaveData = Val(.TextMatrix(.Row, .ColIndex("ID"))) <> 0
        End With
    End If
End Function

Private Function GetOutPayStaut() As Boolean
    '����:���״̬
    If Me.ActiveControl Is vsRequisition Then
        With Me.vsRequisition
            GetOutPayStaut = Val(.TextMatrix(.Row, .ColIndex("ID"))) <> 0
        End With
    Else
        With Me.vsOutPay
            GetOutPayStaut = Val(.TextMatrix(.Row, .ColIndex("ID"))) <> 0
        End With
    End If
End Function

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean, lngID As Long
    If Me.Visible = False Then Exit Sub
    
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = HaveData
    Case conMenu_Edit_Audit '���ȷ��
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "���ȷ��")
            Control.Enabled = Control.Visible And Val(vsRequisition.TextMatrix(vsRequisition.Row, vsRequisition.ColIndex("ID"))) <> 0 And Me.ActiveControl Is vsRequisition
    Case conMenu_Edit_ChargeOff
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "ȡ�����")
        With Me.vsOutPay
             Control.Enabled = Control.Visible And Val(.TextMatrix(.Row, .ColIndex("ID"))) <> 0 And Me.ActiveControl Is vsOutPay
             Control.Enabled = Control.Enabled And Trim(.TextMatrix(.Row, .ColIndex("ȡ��ʱ��"))) = "" And Val(.Cell(flexcpData, .Row, .ColIndex("��־"))) = 0
        End With
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
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim vsGrid As VSFlexGrid
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstr��λ���� & IIf(Not Me.ActiveControl Is vsRequisition, "����嵥", "��������嵥")
    If Not Me.ActiveControl Is vsRequisition Then
        If CStr(mArrFilter("���-����ʱ��")(0)) <> "1901-01-01" Then
            objRow.Add "����ʱ�䣺" & CStr(mArrFilter("���-����ʱ��")(0)) & "��" & CStr(mArrFilter("���-����ʱ��")(1))
        End If
        If CStr(mArrFilter("���-���ʱ��")(0)) <> "1901-01-01" Then
            objRow.Add "���ʱ�䣺" & CStr(mArrFilter("���-���ʱ��")(0)) & "��" & CStr(mArrFilter("���-���ʱ��")(1))
        End If
        If CStr(mArrFilter("���-ȡ��ʱ��")(0)) <> "1901-01-01" Then
            objRow.Add "ȡ��ʱ�䣺" & CStr(mArrFilter("���-ȡ��ʱ��")(0)) & "��" & CStr(mArrFilter("���-ȡ��ʱ��")(1))
        End If
        objPrint.UnderAppRows.Add objRow
        Set objRow = New zlTabAppRow
        objRow.Add "����ˣ�" & UserInfo.����
        If CStr(mArrFilter("�����")) <> "" Then objRow.Add "����ˣ�" & mArrFilter("���-�����")
        objPrint.UnderAppRows.Add objRow
        Set vsGrid = vsOutPay
    Else
        objRow.Add "����ˣ�" & UserInfo.����
        objPrint.UnderAppRows.Add objRow
        Set vsGrid = vsRequisition
    End If
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
    
    Set objPrint.Body = IIf(Not Me.ActiveControl Is vsRequisition, vsOutPay, vsRequisition)
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
End Sub

Private Sub vsOutPay_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsOutPay
        If Col = .ColIndex("��־") Then Cancel = True
    End With
End Sub

Private Sub vsOutPay_DblClick()
        ExcuteFunction False
End Sub

Private Sub vsOutPay_DragDrop(Source As Control, x As Single, y As Single)
    If Source Is vsRequisition Then
        '�϶�
        Call ExcuteFunction(True)   'ȷ��
    End If
End Sub

Private Sub vsOutPay_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    Static objIcon As IPictureDisp
    If Not Source Is vsOutPay Then
        If State = 0 Then
            Set objIcon = Source.DragIcon
        ElseIf State = 2 Then
            Set Source.DragIcon = ils16.ListImages("OutPay").Picture
        ElseIf State = 1 Then
            Set Source.DragIcon = objIcon
        End If
    End If
End Sub

Private Sub vsOutPay_GotFocus()
    vsOutPay.BackColorSel = &H8000000D
End Sub

Private Sub vsOutPay_LostFocus()
    vsOutPay.BackColorSel = &H8000000A
End Sub

Private Sub vsOutPay_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    If Button <> 2 Then Exit Sub
    zlcontrol.ControlSetFocus vsOutPay, True
    Set objPopup = mcbsThis.FindControl(xtpControlPopup, conMenu_EditPopup, , True)
    If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
End Sub

Private Sub vsOutPay_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If zlStr.IsHavePrivs(mstrPrivs, "ȡ�����") = False Then Exit Sub
        If Val(vsOutPay.TextMatrix(vsOutPay.Row, vsOutPay.ColIndex("ID"))) = 0 Then Exit Sub
        If Trim(vsOutPay.TextMatrix(vsOutPay.Row, vsOutPay.ColIndex("ȡ��ʱ��"))) <> "" Then Exit Sub
        
        Set vsOutPay.DragIcon = ils16.ListImages("OutPay").Picture
        vsOutPay.Drag 1
    End If
End Sub
 

Private Sub vsRequisition_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsRequisition
        If Col = .ColIndex("��־") Then Cancel = True
    End With
End Sub

Private Sub vsRequisition_DblClick()
    ExcuteFunction True
End Sub

Private Sub vsRequisition_DragDrop(Source As Control, x As Single, y As Single)
    If Source Is vsOutPay Then
        '�϶�
        Call ExcuteFunction(False)    'ȷ��
    End If
End Sub

Private Sub vsRequisition_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    Static objIcon As IPictureDisp
    If Not Source Is vsRequisition Then
        If State = 0 Then
            Set objIcon = Source.DragIcon
        ElseIf State = 2 Then
            Set Source.DragIcon = ils16.ListImages("Requisition").Picture
        ElseIf State = 1 Then
            Set Source.DragIcon = objIcon
        End If
    End If
End Sub

Private Sub vsRequisition_GotFocus()
    vsRequisition.BackColorSel = &H8000000D
End Sub

Private Sub vsRequisition_LostFocus()
    vsRequisition.BackColorSel = &H8000000A
End Sub

Private Sub vsRequisition_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    If Button <> 2 Then Exit Sub
    zlcontrol.ControlSetFocus vsRequisition, True
    Set objPopup = mcbsThis.FindControl(xtpControlPopup, conMenu_EditPopup, , True)
    If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
End Sub

Private Sub vsRequisition_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If zlStr.IsHavePrivs(mstrPrivs, "���ȷ��") = False Then Exit Sub
        If Val(vsRequisition.TextMatrix(vsRequisition.Row, vsRequisition.ColIndex("ID"))) = 0 Then Exit Sub
        Set vsRequisition.DragIcon = ils16.ListImages("Requisition").Picture
        vsRequisition.Drag 1
    End If
End Sub

