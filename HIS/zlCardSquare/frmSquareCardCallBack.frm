VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSquareCardCallBack 
   BorderStyle     =   0  'None
   Caption         =   "�����ռ�¼"
   ClientHeight    =   8805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   2250
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   3930
      _cx             =   6932
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
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSquareCardCallBack.frx":0000
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
         TabIndex        =   1
         Top             =   45
         Width           =   210
         Begin VB.Image imgCol 
            Height          =   195
            Left            =   0
            Picture         =   "frmSquareCardCallBack.frx":0066
            ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
   End
End
Attribute VB_Name = "frmSquareCardCallBack"
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

'�����������
Public Property Get zlIsHaveData() As Boolean
    zlIsHaveData = mblnHaveData
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
    Dim i As Long
    Dim strHead As String, varHead As Variant
    
    strHead = "��־,1,285|Id,1,0|������,1,600|����,4,1200|���,4,800|�ɷ��ֵ,4,800|��Ч��,4,1850|����ԭ��,1,2000|������,1,800|" & _
              "�쿨��,1,800|����ʱ��,4,1850|������,1,800|����ʱ��,4,1850|��ǰ״̬,4,1200|��ע,1,1800|������,7,1000|" & _
              "���۽��,7,1000|��ֵ�ۿ���,7,1000|���,7,1000|ͣ����,1,800|ͣ������,4,1850|�쿨����ID,1,0|�쿨����,1,1200|�������,1,1500"
    varHead = Split(strHead, "|")
    With vsGrid
        .Cols = UBound(varHead) + 1
        For i = 0 To UBound(varHead)
            .TextMatrix(0, i) = Split(varHead(i), ",")(0)
            .ColKey(i) = Split(varHead(i), ",")(0)
            If .TextMatrix(0, i) = "��־" Then .TextMatrix(0, i) = ""
            .ColAlignment(i) = Split(varHead(i), ",")(1)
            .ColWidth(i) = Split(varHead(i), ",")(2)
            If .ColWidth(i) = 0 Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        
        'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
        .ColData(.ColIndex("ID")) = "-1|1"
        .ColData(.ColIndex("����")) = "1|1"
        .ColData(.ColIndex("�쿨����ID")) = "-1|1"
    End With
End Sub

Private Sub LoadDataToRpt()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ݸ�����
    '����:���˺�
    '����:2009-09-07 11:53:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    
    mblnHaveData = False
    Err = 0: On Error GoTo Errhand:
    strSQL = "" & _
        "Select a.Id, a.������, a.����, a.���, Decode(a.�ɷ��ֵ, 1, '��', '') As �ɷ��ֵ, a.��Ч��," & vbNewLine & _
        "       a.����ԭ��, a.������, a.�쿨��, a.����ʱ��, a.������, a.����ʱ��," & vbNewLine & _
        "       Decode(Mod(a.��ǰ״̬, 10), 2, '����', 3, '�˿�', '����') As ��ǰ״̬, a.��ע, a.������, a.���۽��," & vbNewLine & _
        "       a.��ֵ�ۿ���, a.���, a.ͣ����, a.ͣ������, a.�쿨����id," & vbNewLine & _
        "       b.���� || '-' || b.���� As �쿨����, a.�������" & vbNewLine & _
        "From ���ѿ���Ϣ A, ���ű� B, (Select ����, Decode(��ǰ״̬, 2, 1, 3, 1, 0) + ��� As ��� From ���ѿ���Ϣ Where ID = [1]) C" & vbNewLine & _
        "Where a.���� = c.���� And a.��� < c.��� And a.�쿨����id = b.Id(+) And a.�ӿڱ�� = [2]" & vbNewLine & _
        "Order By ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng���ѿ�ID, mlng�ӿڱ��)
    
    With Me.vsGrid
        .Redraw = flexRDNone
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ID")) = Nvl(rsTemp!id)
            .TextMatrix(lngRow, .ColIndex("������")) = Nvl(rsTemp!������)
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("���")) = Nvl(rsTemp!���)
            .TextMatrix(lngRow, .ColIndex("�ɷ��ֵ")) = Nvl(rsTemp!�ɷ��ֵ)
            .TextMatrix(lngRow, .ColIndex("��Ч��")) = Format(Nvl(rsTemp!��Ч��), "yyyy-MM-dd HH:mm:ss")
            If Trim(.TextMatrix(lngRow, .ColIndex("��Ч��"))) >= "3000-01-01" Then .TextMatrix(lngRow, .ColIndex("��Ч��")) = ""
            
            .TextMatrix(lngRow, .ColIndex("����ԭ��")) = Nvl(rsTemp!����ԭ��)
            .TextMatrix(lngRow, .ColIndex("������")) = Nvl(rsTemp!������)
            .TextMatrix(lngRow, .ColIndex("�쿨��")) = Nvl(rsTemp!�쿨��)
            .TextMatrix(lngRow, .ColIndex("����ʱ��")) = Format(Nvl(rsTemp!����ʱ��), "yyyy-MM-dd HH:mm:ss")
            
            .TextMatrix(lngRow, .ColIndex("������")) = Nvl(rsTemp!������)
            .TextMatrix(lngRow, .ColIndex("����ʱ��")) = Format(Nvl(rsTemp!����ʱ��), "yyyy-MM-dd HH:mm:ss")
            If Trim(.TextMatrix(lngRow, .ColIndex("����ʱ��"))) >= "3000-01-01" Then .TextMatrix(lngRow, .ColIndex("����ʱ��")) = ""
            
            .TextMatrix(lngRow, .ColIndex("��ǰ״̬")) = Nvl(rsTemp!��ǰ״̬)
            .TextMatrix(lngRow, .ColIndex("��ע")) = Nvl(rsTemp!��ע)
            .TextMatrix(lngRow, .ColIndex("������")) = Format(Val(Nvl(rsTemp!������)), "####0.00;-####0.00;;")
            .TextMatrix(lngRow, .ColIndex("���۽��")) = Format(Val(Nvl(rsTemp!���۽��)), "####0.00;-####0.00;;")
            .TextMatrix(lngRow, .ColIndex("��ֵ�ۿ���")) = Val(Nvl(rsTemp!��ֵ�ۿ���))
            .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(Nvl(rsTemp!���)), "####0.00;-####0.00;;")
            
            .TextMatrix(lngRow, .ColIndex("ͣ����")) = Nvl(rsTemp!ͣ����)
            .TextMatrix(lngRow, .ColIndex("ͣ������")) = Format(Nvl(rsTemp!ͣ������), "yyyy-MM-dd HH:mm:ss")
            If Trim(.TextMatrix(lngRow, .ColIndex("ͣ������"))) >= "3000-01-01" Then .TextMatrix(lngRow, .ColIndex("ͣ������")) = ""
            
            .TextMatrix(lngRow, .ColIndex("�쿨����ID")) = Nvl(rsTemp!�쿨����id)
            .TextMatrix(lngRow, .ColIndex("�쿨����")) = Nvl(rsTemp!�쿨����)
            .TextMatrix(lngRow, .ColIndex("�������")) = Nvl(rsTemp!�������)

            lngRow = lngRow + 1
           rsTemp.MoveNext
        Loop
        
        '�ָ�������
        zl_vsGrid_Para_Restore mlngModule, vsGrid, Me.Caption, "�����б�", True
        If .Row <= 0 Then .Row = 1
        .Redraw = flexRDBuffered
    End With
    mblnHaveData = rsTemp.RecordCount > 0
   Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
     Me.vsGrid.Redraw = flexRDBuffered
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
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "�����б�", True
End Sub

Private Sub imgCol_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImg.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImg.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsGrid, lngLeft, lngTop, imgCol.Height)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "�����б�", True
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
    objPrint.Title.Text = gstrUnitName & "���ѿ��������"
    
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

Private Sub vsGrid_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "�����б�", True
End Sub

Private Sub vsGrid_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsGrid, OldRow, NewRow, OldCol, NewCol, gSysColor.lngGridColorSel
End Sub

Private Sub vsGrid_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "�����б�", True
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


