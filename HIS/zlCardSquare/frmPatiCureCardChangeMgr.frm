VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPatiCureCardChangeMgr 
   BorderStyle     =   0  'None
   Caption         =   "�䶯��Ϣ"
   ClientHeight    =   7680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   2250
      Left            =   0
      TabIndex        =   0
      Top             =   0
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
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPatiCureCardChangeMgr.frx":0000
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
      ExplorerBar     =   2
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
            Picture         =   "frmPatiCureCardChangeMgr.frx":0066
            ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
   End
End
Attribute VB_Name = "frmPatiCureCardChangeMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long, mblnHaveData As Boolean
Private mstrCardNo As String, mlngCardTypeID As Long
Public Event zlPopupMenus(ByVal vsGrid As VSFlexGrid) '�����˵�����
Public Event AfterRowChange(ByVal vsGrid As VSFlexGrid) '�����˵�����

Public Function zlReLoadData(ByVal lngCardTypeID As Long, ByVal strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¼�������
    '����:���سɹ�,����true,���򷵻�Flase
    '����:���˺�
    '����:2011-06-28 15:30:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mstrCardNo = strCardNo: mlngCardTypeID = lngCardTypeID
    Err = 0: On Error GoTo Errhand:
    Call LoadDataToRpt
    zlReLoadData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
 

Private Sub LoadDataToRpt()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ݸ�����
    '����:���˺�
    '����:2009-09-07 11:53:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, i As Long, strSQL As String
    mblnHaveData = False
    Err = 0: On Error GoTo Errhand:
    '�����:55447
    '69745:������,2014-02-13,��ȡ���ݿ�ʱ��ʱû�н��и�ʽת������ʱ����������
    strSQL = "" & _
    "   Select A.����, A.�䶯��� As �䶯���id, " & _
    "          Decode(A.�䶯���, 1, '����',11,'�󶨿�', 2, '����', 3, '����',13,'����ͣ��', 4, '�˿�',14,'ȡ����', 5, '�������', 6, '��ʧ', 16, 'ȡ����ʧ', '��') As �䶯���, " & _
    "          to_char(A.�䶯ʱ��,'yyyy-mm-dd hh24:mi:ss') as �䶯ʱ��, A.�䶯ԭ��, A.��ʧ��ʽ, A.����Ա����, to_char(A.�Ǽ�ʱ��,'yyyy-mm-dd hh24:mi:ss') as �Ǽ�ʱ��, A.�䶯id, B.����, B.�Ա�, B.���� " & _
    "   From ����ҽ�ƿ��䶯 A, ������Ϣ B " & _
    "   Where A.����id = B.����id And A.�����id = [1] And ���� = [2] " & _
    "   Order By �Ǽ�ʱ��  "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngCardTypeID, mstrCardNo)
    With Me.vsGrid
        .Redraw = flexRDNone
        Set .DataSource = rsTemp
        If .Rows <= 1 Then .Rows = 2
        For i = 1 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*ID" Then
                .ColAlignment(i) = flexAlignCenterCenter
                .ColHidden(i) = True
                .ColData(i) = "-1|1"
            ElseIf .ColKey(i) Like "�䶯���" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*ʱ��" Then
                .ColAlignment(i) = flexAlignCenterCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
            If .ColKey(i) = "����" Then
                .ColData(i) = "1|1"
            End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        '�ָ�������
        zl_vsGrid_Para_Restore mlngModule, vsGrid, Me.Caption, "ҽ�ƿ��䶯�б�", True
        .ColWidth(.ColIndex("��־")) = 285
        .ColData(.ColIndex("��־")) = "-1|1"
        
        .ColAlignment(.ColIndex("��־")) = flexAlignCenterCenter
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
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "ҽ�ƿ��䶯�б�", True
End Sub
Private Sub imgCol_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlcontrol.GetControlRect(picImg.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImg.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsGrid, lngLeft, lngTop, imgCol.Height)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "ҽ�ƿ��䶯�б�", True
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
    gstrSQL = "Select   A.����, to_char(A.����ʱ��,'yyyy-mm-dd hh24:mi:ss') From ����ҽ�ƿ���Ϣ A where �����ID=[1] and ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngCardTypeID, mstrCardNo)
    If rsTemp.EOF = True Then Exit Sub '�޿���Ϣ���˳�
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
        
    objPrint.Title.Text = gstrUnitName & "ҽ�ƿ��䶯���"
    
    objRow.Add "���ţ�" & Nvl(rsTemp!����)
    objRow.Add "����ʱ�䣺" & Nvl(rsTemp!����ʱ��)
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
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "ҽ�ƿ��䶯�б�", True
End Sub

Private Sub vsGrid_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow <= 0 Then Exit Sub
    
    zl_VsGridRowChange vsGrid, OldRow, NewRow, OldCol, NewCol, gSysColor.lngGridColorSel
    If OldRow <> NewRow Then
        RaiseEvent AfterRowChange(vsGrid)
    End If
End Sub

Private Sub vsGrid_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "ҽ�ƿ��䶯�б�", True
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
 






