VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm���ķ��ϻ��� 
   BorderStyle     =   0  'None
   Caption         =   "���Ļ��ܷ���"
   ClientHeight    =   4620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picHsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000010&
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   5535
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2160
      Width           =   5535
   End
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   2085
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7320
      _cx             =   12912
      _cy             =   3678
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
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frm���ķ��ϻ���.frx":0000
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
   Begin VSFlex8Ctl.VSFlexGrid vsfChargeOff 
      Height          =   2085
      Left            =   0
      TabIndex        =   1
      Top             =   2400
      Width           =   7320
      _cx             =   12912
      _cy             =   3678
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
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm���ķ��ϻ���.frx":0191
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
Attribute VB_Name = "frm���ķ��ϻ���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsNotPayStuff As ADODB.Recordset
Private mrsChargeOff As New ADODB.Recordset                   '������ʾ���������¼
Private mintUnit As Integer
'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString
Private mbln�ֿ��� As Boolean
Private mlngModule As Long
Private mbln����ʱ�������� As Boolean

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
        .ColData(.ColIndex("����")) = IIf(mbln�ֿ���, 1, -1)
        .ColData(.ColIndex("��������")) = 1
        .ColData(.ColIndex("ʵ������")) = 1
    End With
End Sub
Private Sub LoadDataToChargeOffList(ByVal lng����id As Long, ByVal lng����ID As Long)
    With vsfChargeOff
        .Rows = 1
        
        mrsChargeOff.Filter = "���ϲ���id=" & lng����id & " And ����id=" & lng����ID & " And ��������>0 "
        If mrsChargeOff.RecordCount = 0 Then Exit Sub
    
        .Redraw = flexRDNone
        
        Do While Not mrsChargeOff.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("����")) = mrsChargeOff!���ϲ���
            .TextMatrix(.Rows - 1, .ColIndex("NO")) = mrsChargeOff!NO
            .TextMatrix(.Rows - 1, .ColIndex("��������")) = mrsChargeOff!��������
            .TextMatrix(.Rows - 1, .ColIndex("����")) = IIf(NVL(mrsChargeOff!����) = "", "", mrsChargeOff!����)
            .TextMatrix(.Rows - 1, .ColIndex("����")) = IIf(NVL(mrsChargeOff!����) = "", "", mrsChargeOff!����)
            .TextMatrix(.Rows - 1, .ColIndex("׼������")) = Format(mrsChargeOff!׼������ / mrsChargeOff!��װ, mFMT.FM_����)
            .TextMatrix(.Rows - 1, .ColIndex("��������")) = Format(mrsChargeOff!�������� / mrsChargeOff!��װ, mFMT.FM_����)
            .TextMatrix(.Rows - 1, .ColIndex("��λ")) = mrsChargeOff!��λ
            
            .Cell(flexcpFontBold, .Rows - 1, .ColIndex("��������")) = True
            
            mrsChargeOff.MoveNext
        Loop
        
        .Redraw = flexRDDirect
    End With
End Sub


Public Function zlFullData(ByVal intUnit As Integer, ByVal rsNotPayStuff As ADODB.Recordset, ByVal rsChargeOff As ADODB.Recordset) As Boolean
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
    
    mbln����ʱ�������� = (Val(zlDataBase.GetPara("����ʱ�����������ʼ�¼", glngSys, 1723, , , True)) = 1)
    mbln�ֿ��� = mbln����ʱ��������
    
    Set mrsNotPayStuff = rsNotPayStuff
    Set mrsChargeOff = rsChargeOff
    With vsfChargeOff
        .Rows = 1
    End With
    
    With vsGrid
        .Redraw = flexRDNone
        .Rows = .FixedRows + 1
        .Clear (1)
        '�������
        zlFullData = LoadDataToVssGrid
        .Redraw = flexRDBuffered
    End With
    
    Call Form_Resize
End Function
 
Private Sub Form_Load()
    zl_vsGrid_Para_Restore mlngModule, vsGrid, Me.Caption, "�����嵥"
    
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

Private Sub Form_Resize()
    err = 0: On Error Resume Next
    With vsGrid
        .Top = ScaleTop
        .Width = ScaleWidth
        .Left = ScaleLeft
        .Height = IIf(mbln����ʱ�������� = False, ScaleHeight, ScaleHeight / 4 * 3)
    End With

    With picHsc
        .Visible = mbln����ʱ��������
        .Top = vsGrid.Top + vsGrid.Height
        .Width = ScaleWidth
    End With
    
    With vsfChargeOff
        .Visible = mbln����ʱ��������
        .Top = picHsc.Top + picHsc.Height
        .Width = ScaleWidth
        .Left = ScaleLeft
        .Height = ScaleHeight - picHsc.Top - picHsc.Height - 50
    End With
End Sub
Private Function LoadDataToVssGrid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��������䵽����ؼ���(��Ҫ�ǰ�Ʒ��ͳ��)
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-23 17:13:18
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, bln���� As Boolean, strKey As String, strTemp As String
    Dim strSort As String
    Dim dbl�������� As Double
    Dim lng����id As Long
    Dim lng����ID As Long
    
    LoadDataToVssGrid = False
    
    If mbln�ֿ��� = False Then
        vsGrid.ColHidden(vsGrid.ColIndex("����")) = True
        vsGrid.ColHidden(vsGrid.ColIndex("Ӧ������")) = True
        vsGrid.ColHidden(vsGrid.ColIndex("��������")) = True
    Else
        vsGrid.ColHidden(vsGrid.ColIndex("����")) = False
        vsGrid.ColHidden(vsGrid.ColIndex("Ӧ������")) = False
        vsGrid.ColHidden(vsGrid.ColIndex("��������")) = False
    End If
    
    If mrsNotPayStuff.RecordCount = 0 Then
         LoadDataToVssGrid = True
        Exit Function
    End If
        
    mrsNotPayStuff.Filter = 0
    mrsNotPayStuff.MoveFirst
    strSort = IIf(mbln�ֿ���, "����id Asc,", "") & "�������� Asc,��� Asc,���� Asc"
    
    mrsNotPayStuff.Sort = strSort

    With vsGrid
        .Subtotal flexSTClear
        .Rows = 2
        .Clear 1
        lngRow = .FixedRows - 1
        '�ֹ��������ʾ
        strKey = ""
        Do While Not mrsNotPayStuff.EOF
            strTemp = IIf(mbln�ֿ���, NVL(mrsNotPayStuff!����id) & "_", "") & NVL(mrsNotPayStuff!����ID) & IIf(bln����, "_" & NVL(mrsNotPayStuff!����), "")
            
            If strKey <> strTemp And mrsNotPayStuff!ִ��״̬ = 1 Then
                .Rows = .Rows + 1
                lngRow = lngRow + 1
                .TextMatrix(lngRow, .ColIndex("����")) = IIf(mbln�ֿ���, mrsNotPayStuff!����, "")
                .TextMatrix(lngRow, .ColIndex("��������")) = mrsNotPayStuff!��������
                .TextMatrix(lngRow, .ColIndex("����id")) = mrsNotPayStuff!����ID
                .TextMatrix(lngRow, .ColIndex("���")) = mrsNotPayStuff!���
                .TextMatrix(lngRow, .ColIndex("����")) = mrsNotPayStuff!����
                .TextMatrix(lngRow, .ColIndex("����")) = IIf(bln����, mrsNotPayStuff!����, "")
                .TextMatrix(lngRow, .ColIndex("����")) = Format(mrsNotPayStuff!���� * mrsNotPayStuff!����ϵ��, mFMT.FM_���ۼ�)
                .TextMatrix(lngRow, .ColIndex("��λ")) = NVL(mrsNotPayStuff!��λ)

                .Cell(flexcpData, lngRow, .ColIndex("����")) = IIf(bln����, NVL(mrsNotPayStuff!����), 0)
                .Cell(flexcpData, lngRow, .ColIndex("����")) = IIf(mbln�ֿ���, NVL(mrsNotPayStuff!����id), "")
                .Cell(flexcpData, lngRow, .ColIndex("��������")) = mrsNotPayStuff!����ID
                strKey = strTemp
            End If
            If mrsNotPayStuff!ִ��״̬ = 1 Then
                'ֻ�߱����ϵĲŻ���
                .Cell(flexcpData, lngRow, .ColIndex("Ӧ������")) = Val(.Cell(flexcpData, lngRow, .ColIndex("Ӧ������"))) + (mrsNotPayStuff!ʵ������ * mrsNotPayStuff!��)
                .TextMatrix(lngRow, .ColIndex("Ӧ������")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("Ӧ������"))) / mrsNotPayStuff!����ϵ��, mFMT.FM_����)

                .Cell(flexcpData, lngRow, .ColIndex("ʵ������")) = Val(.Cell(flexcpData, lngRow, .ColIndex("ʵ������"))) + (mrsNotPayStuff!ʵ������ * mrsNotPayStuff!��)
                .TextMatrix(lngRow, .ColIndex("ʵ������")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("ʵ������"))) / mrsNotPayStuff!����ϵ��, mFMT.FM_����)
                .Cell(flexcpFontBold, lngRow, .ColIndex("ʵ������")) = True
                
                .Cell(flexcpData, lngRow, .ColIndex("���")) = Val(.Cell(flexcpData, lngRow, .ColIndex("���"))) + (mrsNotPayStuff!���)
                .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("���"))), mFMT.FM_���)
            End If
            mrsNotPayStuff.MoveNext
        Loop
        
        '�ϲ���������
        mrsChargeOff.Filter = "ִ��״̬=1"
        If mbln�ֿ��� = True And mrsChargeOff.RecordCount > 0 Then
            For lngRow = 1 To .Rows - 1
                If Val(.Cell(flexcpData, lngRow, .ColIndex("����"))) > 0 Then
                    .TextMatrix(lngRow, .ColIndex("Ӧ������")) = .TextMatrix(lngRow, .ColIndex("ʵ������"))

                    lng����id = Val(.Cell(flexcpData, lngRow, .ColIndex("����")))
                    lng����ID = Val(.TextMatrix(lngRow, .ColIndex("����id")))

                    mrsChargeOff.Filter = " ִ��״̬=1 And ���ϲ���id=" & lng����id & " And ����ID=" & lng����ID
                    If mrsChargeOff.RecordCount > 0 Then
                        dbl�������� = 0
                        Do While Not mrsChargeOff.EOF
                            dbl�������� = dbl�������� + mrsChargeOff!��������
                            mrsChargeOff.MoveNext
                        Loop
                        
                        .TextMatrix(lngRow, .ColIndex("��������")) = Format(dbl��������, mFMT.FM_����)
                        .TextMatrix(lngRow, .ColIndex("ʵ������")) = Format(Val(.TextMatrix(lngRow, .ColIndex("Ӧ������"))) - Val(.TextMatrix(lngRow, .ColIndex("��������"))), mFMT.FM_����)
                    End If
                End If
            Next
        End If

        If mrsNotPayStuff.RecordCount <> 0 Then mrsNotPayStuff.MoveFirst
        If .Rows > 2 Then .Rows = .Rows - 1
        If .Rows = 2 And Val(.Cell(flexcpData, 1, .ColIndex("��������"))) = 0 Then
        Else
            Call SetTotalRowData(mbln�ֿ���)
        End If
        
        If Val(.Cell(flexcpData, .Row, .ColIndex("����"))) = 0 Then
            vsfChargeOff.Rows = 1
        Else
            Call LoadDataToChargeOffList(Val(.Cell(flexcpData, .Row, .ColIndex("����"))), Val(.TextMatrix(.Row, .ColIndex("����id"))))
        End If
    End With
    
    LoadDataToVssGrid = True
End Function
Private Function SetTotalRowData(ByVal bln���һ��� As Boolean) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:�����еĻ�������
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-22 10:22:21
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, lngCol As Long
    With vsGrid
        .Redraw = flexRDNone
        .OutlineBar = flexOutlineBarComplete
        .SubtotalPosition = flexSTBelow
        If bln���һ��� = True Then
            .Subtotal flexSTSum, .ColIndex("����"), .ColIndex("���"), mFMT.FM_���, , vbBlue, True, "С��"
        End If
        .Subtotal flexSTSum, -1, .ColIndex("���"), mFMT.FM_���, , vbBlue, True, "�ϼ�"
        If bln���һ��� = False Then .TextMatrix(.Rows - 1, .ColIndex("��������")) = "�ϼ�"
        .Redraw = flexRDBuffered
        
    End With
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
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "�����嵥"
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
    
    With vsfChargeOff
        .Font.Size = curFontSize
        Me.Font.Size = .Font.Size
        .Cell(flexcpFontSize, 0, 0, .Rows - 1, .Cols - 1) = .Font.Size
        
        .RowHeightMin = TextHeight("��") + 120
        .RowHeightMax = TextHeight("��") + 120
        .Refresh
    End With
End Sub

Private Sub vsGrid_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow <> OldRow Then
        With vsGrid
            If Val(.Cell(flexcpData, NewRow, .ColIndex("����"))) = 0 Then
                vsfChargeOff.Rows = 1
                Exit Sub
            End If
            
            Call LoadDataToChargeOffList(Val(.Cell(flexcpData, NewRow, .ColIndex("����"))), Val(.TextMatrix(NewRow, .ColIndex("����id"))))
        End With
    End If
End Sub

Private Sub vsGrid_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    With vsGrid
        If Position <= .ColIndex("��������") Then
            ShowMsgBox "���ܽ����ƶ�������������ǰ����!"
            Position = Col
        End If
    End With
End Sub
