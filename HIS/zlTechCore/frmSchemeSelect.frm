VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSchemeSelect 
   AutoRedraw      =   -1  'True
   Caption         =   "���׷���ѡ��"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9420
   Icon            =   "frmSchemeSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraCommand 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      TabIndex        =   7
      Top             =   5685
      Width           =   9390
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   7695
         TabIndex        =   2
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   6585
         TabIndex        =   1
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "ȫ��(&R)"
         Height          =   350
         Left            =   1755
         TabIndex        =   4
         ToolTipText     =   "Ctrl+R"
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "ȫѡ(&A)"
         Height          =   350
         Left            =   645
         TabIndex        =   3
         ToolTipText     =   "Ctrl+A"
         Top             =   135
         Width           =   1100
      End
   End
   Begin VB.PictureBox picTitle 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   9420
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   9420
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���׷�������"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   210
         TabIndex        =   6
         Top             =   75
         Width           =   1080
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsScheme 
      Align           =   1  'Align Top
      Height          =   5385
      Left            =   0
      TabIndex        =   0
      Top             =   300
      Width           =   9420
      _cx             =   16616
      _cy             =   9499
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
      BackColorSel    =   12632256
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   17
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSchemeSelect.frx":058A
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
      OwnerDraw       =   1
      Editable        =   2
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
      FrozenCols      =   1
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
Attribute VB_Name = "frmSchemeSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng����ID As Long 'IN
Private mint��Դ As Integer 'IN:1-����,2-סԺ
Private mstr��� As String 'Out
Private Enum COL���׷���
    colѡ�� = 0
    col��Ч = 1
    col���� = 2
    col���� = 3
    col������λ = 4
    col���� = 5
    col������λ = 6
    colƵ�� = 7
    col�÷� = 8
    col���� = 9
    colִ��ʱ�� = 10
    colִ�п��� = 11
    colִ������ = 12
    col��� = 13
    col��� = 14
    col��ĿID = 15
    col��� = 16
End Enum

Public Function ShowMe(frmParent As Object, ByVal lng����ID As Long, ByVal int��Դ As Integer) As String
'���أ�ѡ�����Ŀ���
'     "+���1,���2,...":��ʾ������Щ���
'     "-���1,���2,...":��ʾ�ſ���Щ���
'     "*"��ʾѡ������,""��ʾȡ���ٺ�
    mstr��� = ""
    mlng����ID = lng����ID
    mint��Դ = int��Դ
    Me.Show 1, frmParent
    ShowMe = mstr���
End Function

Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
    Dim i As Long, blnTmp As Boolean
    With vsScheme
        If .TextMatrix(lngRow, col���) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, col���)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, col���)) = Val(.TextMatrix(lngRow, col���)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, col���)) = Val(.TextMatrix(lngRow, col���)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col���)) = Val(.TextMatrix(lngRow, col���)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col���)) = Val(.TextMatrix(lngRow, col���)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInһ����ҩ = blnTmp
    End With
End Function

Private Sub cmdAll_Click()
    Dim i As Long
    With vsScheme
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, col���)) <> 0 And RowCanSelect(i) = 0 Then
                .TextMatrix(i, colѡ��) = -1
            End If
        Next
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Dim i As Long
    With vsScheme
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, colѡ��) = 0
        Next
    End With
End Sub

Private Sub cmdOK_Click()
    Dim strSel As String, strUnSel As String, i As Long
    
    With vsScheme
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, colѡ��)) <> 0 Then
                strSel = strSel & "," & Val(.TextMatrix(i, col���))
            Else
                strUnSel = strUnSel & "," & Val(.TextMatrix(i, col���))
            End If
        Next
        strSel = Mid(strSel, 2)
        strUnSel = Mid(strUnSel, 2)

        If strSel = "" Then
            MsgBox "��ӳ��׷�����ѡ����Ҫ����Ŀ���ݡ�", vbInformation, gstrSysName
            vsScheme.SetFocus: Exit Sub
        End If
        If strUnSel = "" Then
            mstr��� = "*"
        Else
            If UBound(Split(strSel, ",")) > UBound(Split(strUnSel, ",")) Then
                mstr��� = "-" & strUnSel
            Else
                mstr��� = "+" & strSel
            End If
        End If
    End With
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        Call cmdAll_Click
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call cmdClear_Click
    End If
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    
    lblTitle.Caption = Get��Ŀ����(mlng����ID)
    Call ShowScheme
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    vsScheme.Height = Me.ScaleHeight - picTitle.Height - fraCommand.Height
    fraCommand.Left = 0
    fraCommand.Top = vsScheme.Top + vsScheme.Height
    fraCommand.Width = Me.ScaleWidth
    
    If Me.ScaleWidth - cmdCancel.Width - cmdAll.Left - cmdOK.Width < cmdClear.Left + cmdClear.Width + 300 Then
        cmdOK.Left = cmdClear.Left + cmdClear.Width + 300
        cmdCancel.Left = cmdOK.Left + cmdOK.Width
    Else
        cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - cmdAll.Left
        cmdOK.Left = cmdCancel.Left - cmdOK.Width
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub vsScheme_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= vsScheme.FixedRows And NewCol >= vsScheme.FixedCols Then
        If NewRow <> OldRow Then
            vsScheme.ForeColorSel = vsScheme.CellForeColor
        End If
    End If
End Sub

Private Sub vsScheme_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Col = col���� Then
        vsScheme.AutoSize Col
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsScheme.TextMatrix(vsScheme.FixedRows - 1, Col) & "A")
        If vsScheme.ColWidth(Col) < lngW Then
            vsScheme.ColWidth(Col) = lngW
        ElseIf vsScheme.ColWidth(Col) > vsScheme.Width * 0.5 Then
            vsScheme.ColWidth(Col) = vsScheme.Width * 0.5
        End If
    End If
End Sub

Private Sub vsScheme_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = colѡ�� Then Cancel = True
End Sub

Private Sub vsScheme_DblClick()
    Call vsScheme_KeyPress(32)
End Sub

Private Sub vsScheme_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim i As Long
    
    If Col <> colѡ�� Then
        Cancel = True
    ElseIf Val(vsScheme.TextMatrix(vsScheme.Row, col���)) = 0 Then
        Cancel = True
    Else
        i = RowCanSelect(Row)
        If i > 0 Then
            Cancel = True
            MsgBox "��Ϊ""" & vsScheme.TextMatrix(i, col����) & """�ѳ�����������ƥ�䣬��ҽ�����ܱ�ѡ��", vbInformation, gstrSysName
        End If
    End If
End Sub

Private Sub vsScheme_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsScheme
        '����һ����ҩ������еı��߼�����
        lngLeft = col��Ч: lngRight = col��Ч
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = colƵ��: lngRight = col�÷�
            If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        End If
        
        If Not RowInһ����ҩ(Row, lngBegin, lngEnd) Then Exit Sub
        
        vRect.Left = Left '������߱����
        vRect.Right = Right - 1 '�����ұ߱����
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '���б�����������
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '���б����±���
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        If Between(Row, .Row, .RowSel) And Me.ActiveControl Is vsScheme Then
            SetBkColor hDC, SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsScheme_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = colѡ�� Then Call RowSelectSame(Row)
End Sub

Private Sub vsScheme_KeyPress(KeyAscii As Integer)
    Dim i As Long
    
    With vsScheme
        If KeyAscii = 32 Then
            If .Col <> colѡ�� Then
                KeyAscii = 0
                If Val(.TextMatrix(.Row, col���)) = 0 Then Exit Sub
                
                i = RowCanSelect(.Row)
                If i > 0 And Val(.TextMatrix(.Row, colѡ��)) = 0 Then
                    MsgBox "��Ϊ""" & .TextMatrix(i, col����) & """�ѳ�����������ƥ�䣬��ҽ�����ܱ�ѡ��", vbInformation, gstrSysName
                    Exit Sub
                End If
                
                .TextMatrix(.Row, colѡ��) = IIF(Val(.TextMatrix(.Row, colѡ��)) = 0, -1, 0)
                Call RowSelectSame(.Row)
            End If
        ElseIf KeyAscii = 13 Then
            KeyAscii = 0
            For i = .Row + 1 To .Rows - 1
                If Not .RowHidden(i) Then
                    .Row = i
                    Call .ShowCell(.Row, .Col)
                    Exit For
                End If
            Next
            If i > .Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub

Private Sub RowSelectSame(ByVal lngRow As Long)
'���ܣ�����ָ����(����Ϊ������)��ѡ��״̬,�����ҽ��һ��ѡ��
    Dim i As Long
    
    With vsScheme
        If Val(.TextMatrix(lngRow, col���)) <> 0 Then
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col���)) = Val(.TextMatrix(lngRow, col���)) _
                    Or Val(.TextMatrix(i, col���)) = Val(.TextMatrix(lngRow, col���)) Then
                    .TextMatrix(i, colѡ��) = .TextMatrix(lngRow, colѡ��)
                Else
                    Exit For
                End If
            Next
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col���)) = Val(.TextMatrix(lngRow, col���)) _
                    Or Val(.TextMatrix(i, col���)) = Val(.TextMatrix(lngRow, col���)) Then
                    .TextMatrix(i, colѡ��) = .TextMatrix(lngRow, colѡ��)
                Else
                    Exit For
                End If
            Next
        Else
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col���)) = Val(.TextMatrix(lngRow, col���)) Then
                    .TextMatrix(i, colѡ��) = .TextMatrix(lngRow, colѡ��)
                Else
                    Exit For
                End If
            Next
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col���)) = Val(.TextMatrix(lngRow, col���)) Then
                    .TextMatrix(i, colѡ��) = .TextMatrix(lngRow, colѡ��)
                Else
                    Exit For
                End If
            Next
        End If
    End With
End Sub

Private Function RowCanSelect(ByVal lngRow As Long) As Long
'���ܣ��ж�ָ���е�(���)ҽ���ɷ�ѡ��
'���أ��������ѡ�񣬷���0,���򷵻��к�
    Dim i As Long
    
    With vsScheme
        If .RowData(lngRow) = 1 Then RowCanSelect = lngRow: Exit Function
        
        If Val(.TextMatrix(lngRow, col���)) <> 0 Then
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col���)) = Val(.TextMatrix(lngRow, col���)) _
                    Or Val(.TextMatrix(i, col���)) = Val(.TextMatrix(lngRow, col���)) Then
                    If .RowData(i) = 1 Then RowCanSelect = i: Exit Function
                Else
                    Exit For
                End If
            Next
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col���)) = Val(.TextMatrix(lngRow, col���)) _
                    Or Val(.TextMatrix(i, col���)) = Val(.TextMatrix(lngRow, col���)) Then
                    If .RowData(i) = 1 Then RowCanSelect = i: Exit Function
                Else
                    Exit For
                End If
            Next
        Else
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col���)) = Val(.TextMatrix(lngRow, col���)) Then
                    If .RowData(i) = 1 Then RowCanSelect = i: Exit Function
                Else
                    Exit For
                End If
            Next
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col���)) = Val(.TextMatrix(lngRow, col���)) Then
                    If .RowData(i) = 1 Then RowCanSelect = i: Exit Function
                Else
                    Exit For
                End If
            Next
        End If
    End With
End Function

Private Function ShowScheme() As Boolean
'���ܣ���ȡ����ʾ���ݿ��еĳ��׷�������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    Dim str��ҩ As String, str�巨 As String
    Dim str�걾 As String, str���� As String
    Dim i As Long, j As Long, str��Դ As String
    
    str��Դ = IIF(mint��Դ = 1, "����", "סԺ")
    strSQL = "Select A.���,A.������,A.��Ч,A.������ĿID,A.�ܸ�����,A.��������," & _
        " A.ִ��Ƶ��,A.ҽ������,Nvl(C.����,Decode(Nvl(A.ִ������,0),0,'<����>',5,'<Ժ��ִ��>')) as ִ�п���," & _
        " A.ִ������,A.ʱ�䷽��,B.���,B.����,B.���㵥λ,A.�걾��λ," & _
        " B.�������,B.����ʱ��,E.������� as �շѷ���,E.����ʱ�� as �շѳ���," & _
        " D." & str��Դ & "��λ as ��װ��λ,D." & str��Դ & "��װ as ��װϵ��" & _
        " From ������Ŀ��� A,������ĿĿ¼ B,���ű� C,ҩƷ��� D,�շ���ĿĿ¼ E" & _
        " Where A.������ĿID=B.ID And A.ִ�п���ID=C.ID(+)" & _
        " And A.�շ�ϸĿID=D.ҩƷID(+) And A.�շ�ϸĿID=E.ID(+)" & _
        IIF(mint��Դ = 1, " And A.��Ч=1", "") & " And A.�������ID=[1]" & _
        " Order by A.���"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    With vsScheme
        .Redraw = flexRDNone
        .Rows = .FixedRows '����������
        If rsTmp.EOF Then
            .Rows = .FixedRows + 1
        Else
            .Rows = .FixedRows + rsTmp.RecordCount
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, colѡ��) = -1
                .TextMatrix(i, col��Ч) = IIF(Nvl(rsTmp!��Ч, 0) = 0, "����", "��ʱ")
                .TextMatrix(i, col����) = rsTmp!����
                .Cell(flexcpData, i, col����) = Nvl(rsTmp!�걾��λ) '����걾
                
                '����
                If InStr(",5,6,", rsTmp!���) > 0 Then
                    '��ҩ����������,�����۵�λ���,��װ��λ��ʾ
                    If Not IsNull(rsTmp!�ܸ�����) And Not IsNull(rsTmp!��װϵ��) Then
                        .TextMatrix(i, col����) = FormatEx(rsTmp!�ܸ����� / rsTmp!��װϵ��, 5)
                    End If
                    If Nvl(rsTmp!��Ч, 0) = 1 Then
                        .TextMatrix(i, col������λ) = Nvl(rsTmp!��װ��λ)
                    End If
                Else
                    '�����������ҩ����������
                    If Not IsNull(rsTmp!�ܸ�����) Then
                        .TextMatrix(i, col����) = rsTmp!�ܸ�����
                    End If
                    If rsTmp!��� = "E" And Nvl(rsTmp!������, 0) = 0 _
                        And Val(.TextMatrix(i - 1, col���)) = rsTmp!��� _
                        And InStr(",7,E,", .TextMatrix(i - 1, col���)) > 0 Then
                        .TextMatrix(i, col������λ) = "��" '��ҩ�䷽������λΪ"��"
                    ElseIf Nvl(rsTmp!��Ч, 0) = 1 Then
                        .TextMatrix(i, col������λ) = Nvl(rsTmp!���㵥λ)
                    End If
                End If
                                
                '����
                .TextMatrix(i, col����) = FormatEx(Nvl(rsTmp!��������), 4)
                If Not IsNull(rsTmp!��������) Then
                    .TextMatrix(i, col������λ) = Nvl(rsTmp!���㵥λ)
                End If
                
                .TextMatrix(i, colƵ��) = Nvl(rsTmp!ִ��Ƶ��)
                .TextMatrix(i, col����) = Nvl(rsTmp!ҽ������)
                .TextMatrix(i, colִ��ʱ��) = Nvl(rsTmp!ʱ�䷽��)
                .TextMatrix(i, colִ�п���) = Nvl(rsTmp!ִ�п���)
                .Cell(flexcpData, i, colִ������) = Nvl(rsTmp!ִ������, 0)
                .TextMatrix(i, col���) = rsTmp!���
                .TextMatrix(i, col���) = Nvl(rsTmp!������)
                .TextMatrix(i, col��ĿID) = rsTmp!������ĿID
                .TextMatrix(i, col���) = rsTmp!���
                
                '��ǰ������г����򲻷������Ŀ
                If Not (IsNull(rsTmp!����ʱ��) Or Format(Nvl(rsTmp!����ʱ��), "yyyy-MM-dd") = "3000-01-01") Then
                    .RowData(i) = 1
                ElseIf Not (Nvl(rsTmp!�������, 0) = 3 Or Nvl(rsTmp!�������, 0) = mint��Դ) Then
                    .RowData(i) = 1
                ElseIf Not IsNull(rsTmp!��װ��λ) Then
                    '��ҩƷ,ͬʱҪ�жϵ��շ���ĿĿ¼
                    If Not (IsNull(rsTmp!�շѳ���) Or Format(Nvl(rsTmp!�շѳ���), "yyyy-MM-dd") = "3000-01-01") Then
                        .RowData(i) = 1
                    ElseIf Not (Nvl(rsTmp!�շѷ���, 0) = 3 Or Nvl(rsTmp!�շѷ���, 0) = mint��Դ) Then
                        .RowData(i) = 1
                    End If
                End If
                
                rsTmp.MoveNext
            Next
            
            '�ٴ���һЩ�����е�����,��������ݵ���ʾ
            For i = 1 To .Rows - 1
                '��ҩ;��
                If .TextMatrix(i, col���) = "E" And Val(.TextMatrix(i, col���)) = 0 _
                    And Val(.TextMatrix(i - 1, col���)) = Val(.TextMatrix(i, col���)) _
                    And InStr(",5,6,", .TextMatrix(i - 1, col���)) > 0 Then
                    .RowHidden(i) = True
                    '��ʾ��ҩ;��
                    For j = i - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, col���)) = Val(.TextMatrix(i, col���)) Then
                            .TextMatrix(j, col�÷�) = .TextMatrix(i, col����)
                            
                            '��ʾ��ҩִ������
                            If Val(.Cell(flexcpData, j, colִ������)) = 5 And Val(.Cell(flexcpData, i, colִ������)) <> 5 Then
                                .TextMatrix(j, colִ������) = "�Ա�ҩ"
                            ElseIf Val(.Cell(flexcpData, j, colִ������)) <> 5 And Val(.Cell(flexcpData, i, colִ������)) = 5 Then
                                .TextMatrix(j, colִ������) = "��Ժ��ҩ"
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
                
                '��ҩ�䷽�ͼ������
                If .TextMatrix(i, col���) = "E" And Val(.TextMatrix(i, col���)) = 0 _
                    And Val(.TextMatrix(i - 1, col���)) = Val(.TextMatrix(i, col���)) _
                    And InStr(",7,E,C,", .TextMatrix(i - 1, col���)) > 0 Then
                    
                    str��ҩ = "": str�巨 = "": str�걾 = "": strTmp = ""
                    j = .FindRow(CStr(Val(.TextMatrix(i, col���))), , col���)
                    
                    '��ҩ�������ִ�п���
                    .TextMatrix(i, colִ�п���) = .TextMatrix(j, colִ�п���)
                    
                    '��ʾ��ҩ�䷽ִ������:��ҩƷΪ׼�ж�
                    If .TextMatrix(i - 1, col���) <> "C" Then
                        If Val(.Cell(flexcpData, j, colִ������)) = 5 And Val(.Cell(flexcpData, i, colִ������)) <> 5 Then
                            .TextMatrix(i, colִ������) = "�Ա�ҩ"
                        ElseIf Val(.Cell(flexcpData, j, colִ������)) <> 5 And Val(.Cell(flexcpData, i, colִ������)) = 5 Then
                            .TextMatrix(i, colִ������) = "��Ժ��ҩ"
                        End If
                    End If
                    
                    For j = j To i - 1
                        .RowHidden(j) = j <> i
                        If .TextMatrix(j, col���) = "7" Then
                            str��ҩ = str��ҩ & "," & RTrim(.TextMatrix(j, col����) & _
                                " " & .TextMatrix(j, col����) & .TextMatrix(j, col������λ) & _
                                " " & .TextMatrix(j, col����))
                        ElseIf .TextMatrix(j, col���) = "C" Then
                            strTmp = strTmp & "," & .TextMatrix(j, col����)
                            str�걾 = .Cell(flexcpData, j, col����) 'ȡ��һ��������Ŀ�ı걾
                        ElseIf .TextMatrix(j, col���) = "E" And Val(.TextMatrix(j, col���)) <> 0 Then
                            str�巨 = .TextMatrix(j, col����)
                        End If
                    Next
                    
                    .TextMatrix(i, col�÷�) = .TextMatrix(i, col����) '��ʾ��ҩ�÷������ɼ�����
                    
                    If .TextMatrix(i - 1, col���) = "C" Then
                        .TextMatrix(i, col����) = Mid(strTmp, 2) & IIF(str�걾 <> "", "(" & str�걾 & ")", "")
                    Else
                        .TextMatrix(i, col����) = "��ҩ�䷽," & .TextMatrix(i, colƵ��) & "," & _
                            str�巨 & "," & .TextMatrix(i, col����) & ":" & Mid(str��ҩ, 2)
                    End If
                End If
                
                '������
                If .TextMatrix(i, col���) = "D" And Val(.TextMatrix(i, col���)) = 0 Then
                    strTmp = ""
                    For j = i + 1 To .Rows - 1
                        If Val(.TextMatrix(j, col���)) = Val(.TextMatrix(i, col���)) Then
                            .RowHidden(j) = True
                            strTmp = strTmp & "," & .Cell(flexcpData, j, col����)
                        Else
                            Exit For
                        End If
                    Next
                    If strTmp <> "" Then
                        .TextMatrix(i, col����) = .TextMatrix(i, col����) & "(" & Mid(strTmp, 2) & ")"
                    End If
                End If
                
                '������Ŀ
                If .TextMatrix(i, col���) = "F" And Val(.TextMatrix(i, col���)) = 0 Then
                    strTmp = "": str���� = ""
                    For j = i + 1 To .Rows - 1
                        If Val(.TextMatrix(j, col���)) = Val(.TextMatrix(i, col���)) Then
                            .RowHidden(j) = True
                            If .TextMatrix(j, col���) = "F" Then
                                strTmp = strTmp & "," & .TextMatrix(j, col����)
                            ElseIf .TextMatrix(j, col���) = "G" Then
                                str���� = .TextMatrix(j, col����)
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    If strTmp <> "" Or str���� <> "" Then
                        If str���� <> "" Then
                            .TextMatrix(i, col����) = "�� " & str���� & " ���� " & .TextMatrix(i, col����)
                        Else
                            .TextMatrix(i, col����) = "�� " & .TextMatrix(i, col����)
                        End If
                        If strTmp <> "" Then
                            .TextMatrix(i, col����) = .TextMatrix(i, col����) & " �� " & Mid(strTmp, 2)
                        End If
                    End If
                End If
            Next
            
            '���˱�ǵ��е������һ�����,��ȡ��ѡ��
            For i = 1 To .Rows - 1
                If .RowData(i) = 1 Then
                    .TextMatrix(i, colѡ��) = 0
                    Call RowSelectSame(i)
                End If
            Next
        End If
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        .Row = .FixedRows: .Col = .FixedCols
        .AutoSize col����
        .Redraw = flexRDDirect
    End With
    ShowScheme = True
    Exit Function
errH:
    vsScheme.Redraw = flexRDDirect
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
