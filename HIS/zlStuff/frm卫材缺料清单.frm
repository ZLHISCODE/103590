VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm����ȱ���嵥 
   BorderStyle     =   0  'None
   Caption         =   "����ȱ���嵥"
   ClientHeight    =   4845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   4125
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7320
      _cx             =   12912
      _cy             =   7276
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483644
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   12632256
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
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
      Cols            =   18
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frm����ȱ���嵥.frx":0000
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
Attribute VB_Name = "frm����ȱ���嵥"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsNotPayStuff As ADODB.Recordset
Private mintUnit As Integer
Private mlngModule As Long
'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
Private Sub InitVsGrid()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ����ؼ�
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-12 10:27:06
    '-----------------------------------------------------------------------------------------------------------
    With vsGrid
        '0-��ѡ,1-��ѡ,-1-����
        .ColData(.ColIndex("״̬")) = 1
        .ColData(.ColIndex("��������")) = 1
        .ColData(.ColIndex("���ݺ�")) = 1
        .ColData(.ColIndex("��������")) = 1
        .ColData(.ColIndex("����")) = 1
    End With
End Sub

Private Sub Form_Resize()
    err = 0: On Error Resume Next
    With vsGrid
        .Top = ScaleTop
        .Width = ScaleWidth
        .Left = ScaleLeft
        .Height = ScaleHeight
    End With
End Sub
Public Function zlFullData(ByVal intUnit As Integer, ByVal rsNotPayStuff As ADODB.Recordset) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���������ݵ�Vss�ؼ���
    '���:rsNotPayStuff-δ�����嵥
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-23 17:11:13
    '-----------------------------------------------------------------------------------------------------------
    If mintUnit <> intUnit Then
        '��Ҫ��ʼ����ص����ָ�ʽ������
        Call Form_Load
    End If
    mintUnit = intUnit
    
    Set mrsNotPayStuff = rsNotPayStuff
    With vsGrid
        .Redraw = flexRDNone
        .Rows = .FixedRows + 1
        .Clear (1)
        '�������
        zlFullData = LoadDataToVssGrid
        .Redraw = flexRDBuffered
    End With
End Function
 
Private Sub Form_Load()
    zl_vsGrid_Para_Restore mlngModule, vsGrid, Me.Caption, "ȱ���嵥"
    Call InitVsGrid
    
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
    End With
    With mOraFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���, True)
        .FM_��� = GetFmtString(mintUnit, g_���, True)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�, True)
        .FM_���� = GetFmtString(mintUnit, g_����, True)
    End With
End Sub
Private Function LoadDataToVssGrid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����ص�������䵽ָ��������ؼ���
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-23 11:06:21
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    LoadDataToVssGrid = False
    
    err = 0: On Error GoTo ErrHand:

    '������ݵ��ؼ���
    mrsNotPayStuff.Filter = 0
    If mrsNotPayStuff.RecordCount <> 0 Then mrsNotPayStuff.MoveFirst
    
    With vsGrid
        If mrsNotPayStuff.EOF Then '
            LoadDataToVssGrid = True
            Exit Function
        End If
        lngRow = .FixedRows
        Do While Not mrsNotPayStuff.EOF
            If mrsNotPayStuff!ִ��״̬ = 0 Then
                .RowData(lngRow) = Val(mrsNotPayStuff!Id)
                .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsNotPayStuff!����)
                .Cell(flexcpData, lngRow, .ColIndex("���ݺ�")) = Val(NVL(mrsNotPayStuff!λ��))
                .TextMatrix(lngRow, .ColIndex("����ҽ��")) = NVL(mrsNotPayStuff!����ҽ��)
                .TextMatrix(lngRow, .ColIndex("״̬")) = NVL(mrsNotPayStuff!״̬)
                '24-�շѴ������ϣ�25-���ʵ��������ϣ�26-���ʱ������ϣ�
                .TextMatrix(lngRow, .ColIndex("��������")) = NVL(mrsNotPayStuff!����)
                .TextMatrix(lngRow, .ColIndex("���ݺ�")) = NVL(mrsNotPayStuff!NO)
                .TextMatrix(lngRow, .ColIndex("����Ա")) = NVL(mrsNotPayStuff!����Ա)
                .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsNotPayStuff!����)
                .TextMatrix(lngRow, .ColIndex("��������")) = NVL(mrsNotPayStuff!����)
                .TextMatrix(lngRow, .ColIndex("סԺ��")) = NVL(mrsNotPayStuff!סԺ��)
                .TextMatrix(lngRow, .ColIndex("��������")) = NVL(mrsNotPayStuff!��������)
                .TextMatrix(lngRow, .ColIndex("���")) = NVL(mrsNotPayStuff!���)
                .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsNotPayStuff!����)
                .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsNotPayStuff!����)
                '.TextMatrix(lngRow, .ColIndex("��")) = Format(Val(NVL(mrsNotPayStuff!��)), "###")
                .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsNotPayStuff!����)
                .TextMatrix(lngRow, .ColIndex("����")) = Format(Val(NVL(mrsNotPayStuff!����)) * mrsNotPayStuff!����ϵ��, mFMT.FM_���ۼ�)
                .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(NVL(mrsNotPayStuff!���)), mFMT.FM_���)
                .TextMatrix(lngRow, .ColIndex("˵��")) = NVL(mrsNotPayStuff!˵��)
                .TextMatrix(lngRow, .ColIndex("����ʱ��")) = NVL(mrsNotPayStuff!����ʱ��)
                .Rows = .Rows + 1
                lngRow = lngRow + 1
            End If
            mrsNotPayStuff.MoveNext
         Loop
    End With
    LoadDataToVssGrid = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Property Get zlHaveData() As Boolean
    Dim i As Integer
    With vsGrid
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("��������")) <> "" Then zlHaveData = True: Exit Function
        Next
    End With
    zlHaveData = False
End Property

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "ȱ���嵥"
End Sub

Public Sub zlSetFontSize(ByVal curFontSize As Currency)
    '-----------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-05-06 17:00:44
    '-----------------------------------------------------------------------------------------------------------
    With vsGrid
        .Font.Size = curFontSize
        Me.Font.Size = .Font.Size
        .Cell(flexcpFontSize, 0, 0, .Rows - 1, .Cols - 1) = .Font.Size
        
        .RowHeightMin = TextHeight("��") + 120
        .RowHeightMax = TextHeight("��") + 120
        .Refresh
    End With
End Sub


