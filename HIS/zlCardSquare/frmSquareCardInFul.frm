VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSquareCardInFul 
   BorderStyle     =   0  'None
   Caption         =   "��ֵ��¼"
   ClientHeight    =   5565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   4485
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8115
      _cx             =   14314
      _cy             =   7911
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSquareCardInFul.frx":0000
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
         Left            =   45
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   1
         Top             =   60
         Width           =   210
         Begin VB.Image imgCol 
            Height          =   195
            Left            =   0
            Picture         =   "frmSquareCardInFul.frx":018F
            ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
   End
End
Attribute VB_Name = "frmSquareCardInFul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long, mblnHaveData As Boolean
Private mlng���ѿ�ID As Long, mlng�ӿڱ�� As Long
Private mstr������ As String, mstr���� As String
'һЩ�����¼�
Public Event zlPopupMenus(ByVal vsGrid As VSFlexGrid) '�����˵�����
Public Event AfterRowChange(ByVal vsGrid As VSFlexGrid) '�����˵�����

'�����������
Public Property Get zlIsHaveData() As Boolean
    zlIsHaveData = mblnHaveData
End Property

Public Property Get zl�������() As Boolean
    With vsGrid
        If .Row < 1 Then zl������� = False: Exit Property
        zl������� = Val(.TextMatrix(.Row, .ColIndex("��¼״̬"))) = 1
    End With
End Property

Public Property Get zlGet��ֵID() As Long
    With vsGrid
        If .Row < 1 Then zlGet��ֵID = 0: Exit Property
        zlGet��ֵID = Val(.TextMatrix(.Row, .ColIndex("ID")))
    End With
End Property

Public Function zlReLoadData(ByVal lng�ӿڱ�� As Long, ByVal lng���ѿ�ID As Long, _
    ByVal str������ As String, ByVal str���� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¼�������
    '���:mcllFilter-��������(Ŀǰ��)
    '����:
    '����:���سɹ�,����true,���򷵻�Flase
    '����:���˺�
    '����:2009-11-20 16:00:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlng���ѿ�ID = lng���ѿ�ID: mlng�ӿڱ�� = lng�ӿڱ��
    mstr������ = str������: mstr���� = str����
    Err = 0: On Error GoTo Errhand:
    Call LoadDataToRpt
    zlReLoadData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitVsGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2009-11-20 16:05:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsGrid
        'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
        .ColData(.ColIndex("��¼״̬")) = "-1|1"
        .ColData(.ColIndex("ID")) = "-1|1"
        .ColHidden(.ColIndex("ID")) = True
        .ColHidden(.ColIndex("��¼״̬")) = True
        
        .ColData(.ColIndex("���")) = "1|1"
        .ColData(.ColIndex("��ֵ���")) = "1|0"
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
    Dim strSQL As String
    
    mblnHaveData = False
    Err = 0: On Error GoTo Errhand:
    strSQL = _
        "Select a.Id, a.���ѿ�id, a.���, a.��¼״̬, a.Ӧ�ս�� As ��ֵ���, a.���� As ��ֵ�ۿ�, a.ʵ�ս�� As �ɿ���," & vbNewLine & _
        "       To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��ֵʱ��, a. ����Ա����, a.�ɿ������� As �ɿ���, ��ע" & vbNewLine & _
        "From ���˿������¼ A" & vbNewLine & _
        "Where a.��¼���� = 2 And a.���ѿ�id = [1]" & vbNewLine & _
        "Order By ��� Desc, ��ֵʱ�� Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng���ѿ�ID)
    
    With Me.vsGrid
        .Redraw = flexRDNone
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ID")) = Nvl(rsTemp!id)
            .TextMatrix(lngRow, .ColIndex("���")) = Nvl(rsTemp!���)
            .Cell(flexcpData, lngRow, .ColIndex("���")) = Nvl(rsTemp!id)
            .TextMatrix(lngRow, .ColIndex("��¼״̬")) = Nvl(rsTemp!��¼״̬)
            .TextMatrix(lngRow, .ColIndex("��ֵ���")) = Format(Val(Nvl(rsTemp!��ֵ���)), "####0.00;-####0.00;;")
            .TextMatrix(lngRow, .ColIndex("��ֵ�ۿ�")) = Format(Val(Nvl(rsTemp!��ֵ�ۿ�)), "####0.00;-####0.00;;")
            .TextMatrix(lngRow, .ColIndex("�ɿ���")) = Format(Val(Nvl(rsTemp!�ɿ���)), "####0.00;-####0.00;;")
            .TextMatrix(lngRow, .ColIndex("��ֵʱ��")) = Nvl(rsTemp!��ֵʱ��)
            .TextMatrix(lngRow, .ColIndex("��ֵ˵��")) = Nvl(rsTemp!��ע)
            .TextMatrix(lngRow, .ColIndex("����Ա����")) = Nvl(rsTemp!����Ա����)
            .TextMatrix(lngRow, .ColIndex("�ɿ���")) = Nvl(rsTemp!�ɿ���)

            Select Case Val(Nvl(rsTemp!��¼״̬))
            Case 2
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbRed
            Case 3
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = &H8000000D
            Case Else
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = &H80000012
            End Select
            lngRow = lngRow + 1
           rsTemp.MoveNext
        Loop
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .Redraw = flexRDBuffered
        mblnHaveData = rsTemp.RecordCount > 0
        '�ָ�������
        zl_vsGrid_Para_Restore mlngModule, vsGrid, Me.Caption, "��ֵ�б�", True
        .ColWidth(.ColIndex("��־")) = 285
        If .Row <= 0 Then .Row = 1
    End With
    Call vsGrid_AfterRowColChange(-1, 0, 1, 1)
   Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    zlCommFun.StopFlash
End Sub

Private Sub Form_Load()
    mlngModule = glngModul: mstrPrivs = gstrPrivs

    Call InitVsGrid
    Call vsGrid_GotFocus
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With vsGrid
        .Left = ScaleLeft: .Top = ScaleTop
        .Width = ScaleWidth: .Height = ScaleHeight
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "��ֵ�б�", True
End Sub

Private Sub imgCol_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImg.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImg.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsGrid, lngLeft, lngTop, imgCol.Height)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "��ֵ�б�", True
End Sub

Private Sub picImg_Click()
    Call imgCol_Click
End Sub
 
Public Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���д�ӡ,Ԥ���������EXCEL
    '���:bytFunc=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '����:���˺�
    '����:2009-11-20 16:36:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim vsGrid As VSFlexGrid
    
    Err = 0: On Error GoTo errH:
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstrUnitName & "���ѿ���ֵ���"
    
    objRow.Add "�����ͣ�" & mstr������
    objRow.Add "���ţ�" & mstr����
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    Err = 0: On Error GoTo Errhand:
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
Errhand:
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

Private Sub vsGrid_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsGrid, OldRow, NewRow, OldCol, NewCol, gSysColor.lngGridColorSel
    If OldRow <> NewRow Then
        RaiseEvent AfterRowChange(vsGrid)
    End If
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
    zl_VsGridLostFocus vsGrid, gSysColor.lngGridColorLost
End Sub

Private Sub vsGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbRightButton Then Exit Sub
    RaiseEvent zlPopupMenus(vsGrid)
End Sub
