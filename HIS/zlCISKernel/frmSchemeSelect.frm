VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSchemeSelect 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
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
      Begin VB.Label lblRemark 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ע�⣺ǳ��ɫ����������û�п���ҩƷ�����ġ�"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   4995
         TabIndex        =   8
         Top             =   75
         Visible         =   0   'False
         Width           =   3960
      End
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
      Cols            =   28
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
Private mint��Դ As Integer 'IN:1-����,2-סԺ,3-�����סԺ
Private mstr��� As String 'Out
Private Enum COL���׷���
    colѡ�� = 0
    col��Ч = 1
    col���� = 2
    col���� = 3
    col������λ = 4
    col���� = 5
    col������λ = 6
    col���� = 7
    colƵ�� = 8
    col�÷� = 9
    col���� = 10
    colִ��ʱ�� = 11
    colִ�п��� = 12
    colִ������ = 13
    col��� = 14
    col��� = 15
    col��ĿID = 16
    col��� = 17
    col�걾��λ = 18
    col��鷽�� = 19
    col�Ƿ����� = 20
    col��ʾ = 21
    col������� = 22
    col��ֵ���� = 23
    colִ�б�� = 24
    col�Ա� = 25
    col����Ӧ�� = 26
    col�������� = 27
End Enum
Private mstr�Ա� As String
Private mlng���˿���id As Long
Private mbln������Ȩ�� As Boolean
Private mbln������Ȩ�� As Boolean
Private mbln������Ȩ�� As Boolean
Private mbln������Ȩ�� As Boolean
Private mlng�������� As Long '0-��ͨסԺ����,1-�������۲���,2-סԺ���۲���

Public Function ShowMe(frmParent As Object, ByVal lng����ID As Long, ByVal int��Դ As Integer, Optional ByVal lng���˿���ID As Long, _
        Optional ByVal str�Ա� As String, Optional ByVal lng�������� As Long) As String
'���أ�ѡ�����Ŀ���
'     "+���1,���2,...":��ʾ������Щ���
'     "-���1,���2,...":��ʾ�ſ���Щ���
'     "*"��ʾѡ������,""��ʾȡ���ٺ�
    mstr��� = ""
    mlng����ID = lng����ID
    mint��Դ = int��Դ
    mlng���˿���id = lng���˿���ID
    mstr�Ա� = str�Ա�
    mlng�������� = lng��������
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

Private Sub cmdALL_Click()
    Dim i As Long
    Dim lngEnd As Long
    
    With vsScheme
        For i = .FixedRows To .Rows - 1
            If CheckCanSelGroup(i, False) Then
                '��ǰ�ļ��ҽ����������Ϊ���׷���
                If .TextMatrix(i, col���) = "D" Then
                    If Val(.TextMatrix(i, col���)) = 0 Then
                        If Not CheckIsOldAdvice(i) Then
                            Call SelGroup(i, 1, lngEnd)
                        End If
                    Else
                        '�������Ѵ���
                    End If
                Else
                    Call SelGroup(i, 1, lngEnd)
                End If
            End If
            If i < lngEnd Then i = lngEnd
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
        Call cmdALL_Click
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call cmdClear_Click
    End If
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    
    lblTitle.Caption = Sys.RowValue("������ĿĿ¼", mlng����ID, "����")
    lblRemark.Visible = mlng���˿���id <> 0
    'ִ������
    If mint��Դ = 3 Then
        vsScheme.ColHidden(col����) = True
    Else
        vsScheme.ColHidden(col����) = Val(zlDatabase.GetPara("ҽ��ִ������", glngSys, IIF(mint��Դ = 1, p����ҽ���´�, pסԺҽ���´�))) = 0
    End If
    
    If mint��Դ = 1 Then
        mbln������Ȩ�� = InStr(GetTsPrivs(p����ҽ���´�), ";�´�����ҩ��;") = 0
        mbln������Ȩ�� = InStr(GetTsPrivs(p����ҽ���´�), ";�´ﶾ��ҩ��;") = 0
        mbln������Ȩ�� = InStr(GetTsPrivs(p����ҽ���´�), ";�´ﾫ��ҩ��;") = 0
        mbln������Ȩ�� = InStr(GetTsPrivs(p����ҽ���´�), ";�´����ҩ��;") = 0
    ElseIf mint��Դ = 2 Then
        mbln������Ȩ�� = InStr(GetTsPrivs(pסԺҽ���´�), ";�´�����ҩ��;") = 0
        mbln������Ȩ�� = InStr(GetTsPrivs(pסԺҽ���´�), ";�´ﶾ��ҩ��;") = 0
        mbln������Ȩ�� = InStr(GetTsPrivs(pסԺҽ���´�), ";�´ﾫ��ҩ��;") = 0
        mbln������Ȩ�� = InStr(GetTsPrivs(pסԺҽ���´�), ";�´����ҩ��;") = 0
    ElseIf mint��Դ = 3 Then
        mbln������Ȩ�� = InStr(GetTsPrivs(p����ҽ���´�), ";�´�����ҩ��;") = 0 And InStr(GetTsPrivs(pסԺҽ���´�), ";�´�����ҩ��;") = 0
        mbln������Ȩ�� = InStr(GetTsPrivs(p����ҽ���´�), ";�´ﶾ��ҩ��;") = 0 And InStr(GetTsPrivs(pסԺҽ���´�), ";�´ﶾ��ҩ��;") = 0
        mbln������Ȩ�� = InStr(GetTsPrivs(p����ҽ���´�), ";�´ﾫ��ҩ��;") = 0 And InStr(GetTsPrivs(pסԺҽ���´�), ";�´ﾫ��ҩ��;") = 0
        mbln������Ȩ�� = InStr(GetTsPrivs(p����ҽ���´�), ";�´����ҩ��;") = 0 And InStr(GetTsPrivs(pסԺҽ���´�), ";�´����ҩ��;") = 0
    End If
    
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
    With vsScheme
        If Col <> colѡ�� Then
            Cancel = True
        ElseIf Val(.TextMatrix(.Row, col���)) = 0 Then
            Cancel = True
        Else
            '��ǰ�ļ��ҽ��������ѡ��
            If CheckIsOldAdvice(Row) Then
                MsgBox "�ü��ҽ����ϵͳ������ǰ�´�ģ������з�ʽ�����ݣ����ܱ���Ϊ���׷�����", vbInformation, gstrSysName
                Cancel = True: Exit Sub
            End If
            If .TextMatrix(Row, colѡ��) <> 0 Then
                Call SelGroup(Row, 0)
            Else
                If CheckCanSelGroup(Row, True) Then
                    Call SelGroup(Row, -1)
                End If
            End If
            '�Ѿ������жϺ�ѡ�񣬲��败��AfterEdit�¼�
            Cancel = True
        End If
    End With
End Sub

Private Function CheckIsOldAdvice(ByVal lngRow As Long) As Boolean
'���ܣ����ָ���еļ��ҽ���Ƿ��Ϸ�ʽ
'������lngRow=���ҽ���ɼ���
    Dim lngIdx As Long

    With vsScheme
        If .TextMatrix(lngRow, col���) = "D" Then
            lngIdx = .FindRow(CStr(.TextMatrix(lngRow, col���)), lngRow + 1, col���)
            If lngIdx = -1 Then
                'CheckIsOldAdvice = True '��ǰ�ĵ���λ���
            ElseIf Val(.TextMatrix(lngIdx, col��ĿID)) <> Val(.TextMatrix(lngRow, col��ĿID)) Then
                CheckIsOldAdvice = True '��ǰ�Ķಿλ��Ŀ���
            End If
        End If
    End With
End Function

Private Sub vsScheme_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsScheme
        '����һ����ҩ������еı��߼�����
        lngLeft = col��Ч: lngRight = col��Ч
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = col����: lngRight = col�÷�
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
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, OS.SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsScheme_KeyPress(KeyAscii As Integer)
    Dim i As Long
    
    With vsScheme
        If KeyAscii = 32 Then
            If .Col <> colѡ�� Then
                KeyAscii = 0
                If .TextMatrix(.Row, colѡ��) = 0 Then
                    If CheckCanSelGroup(.Row, True) Then
                        Call SelGroup(.Row, -1)
                    End If
                Else
                    Call SelGroup(.Row, 0)
                End If
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


Private Function ShowScheme() As Boolean
'���ܣ���ȡ����ʾ���ݿ��еĳ��׷�������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    Dim str��ҩ As String, str�巨 As String
    Dim str�걾 As String, str���� As String
    Dim i As Long, j As Long, lngEnd As Long
    Dim strDepartments As String
    Dim lngSel As Long
    Dim str������� As String
    
    If mlng�������� = 1 Then
        str������� = ",1,2,"
    Else
        str������� = "," & mint��Դ & ","
    End If

    '���ﲻ֧������ҽ������
    strSQL = "Select (Select Count(1) From �������ÿ��� Where ��ĿID=b.ID) as ���ÿ�����,g.����id as ���ÿ���ID,A.���,A.������,A.��Ч,A.������ĿID,A.ҽ������,A.�ܸ�����,A.��������,A.����," & _
             " A.ִ��Ƶ��,A.ҽ������,Nvl(C.����,Decode(Nvl(A.ִ������,0),0,'<����>',5,'-')) as ִ�п���,Nvl(b.�����Ա�, 0) As �Ա�," & _
             " A.ִ������,A.ִ�б��,A.ʱ�䷽��,Nvl(B.���,'*') as ���,Nvl(E.����||Decode(E.���,NULL,NULL,' '||E.���),B.����) as ����,B.���㵥λ," & _
             " E.���㵥λ as ɢװ��λ,A.�걾��λ,A.��鷽��,B.�������,B.����Ӧ��,B.��������,B.����ʱ��,E.������� as �շѷ���,E.����ʱ�� as �շѳ���,E.ID as �շ���ĿID,Nvl(f.��������,0) as ��������," & _
             Decode(mint��Դ, 1, "D.�����װ as ��װϵ��,D.���ﵥλ as ��װ��λ", _
                    2, "D.סԺ��װ as ��װϵ��,D.סԺ��λ as ��װ��λ", 3, "1 as ��װϵ��,E.���㵥λ as ��װ��λ") & _
                    ",(Select f_List2str(Cast(Collect(j.����) As t_Strlist))" & vbNewLine & _
                    "         From ������Ŀ��� H, ������ĿĿ¼ J, �շ���ĿĿ¼ k" & vbNewLine & _
                    "         Where h.������Ŀid = j.Id And k.id(+)=h.�շ�ϸĿID And a.�������id = h.�������id And a.��� = h.������ And NVL(k.����ʱ��,j.����ʱ��) <> To_Date('3000/1/1', 'yyyy/mm/dd') And" & vbNewLine & _
                    "               (j.��� in ('C', '7') Or j.��� = 'E' And Nvl(j.ִ�з���,0) = 0 And j.�������� = '3')) As ��ʾ ,h.�������,h.��ֵ���� " & _
                    " From ������Ŀ��� A,������ĿĿ¼ B,���ű� C,ҩƷ��� D,�շ���ĿĿ¼ E,�������� F,�������ÿ��� G,ҩƷ���� H" & _
                    " Where A.������ĿID=B.ID" & IIF(mint��Դ = 1, "", "(+)") & " And A.ִ�п���ID=C.ID(+) And e.id=f.����id(+)  And h.ҩ��ID(+)=b.ID " & _
                    " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & _
                    " And A.�շ�ϸĿID=D.ҩƷID(+) And A.�շ�ϸĿID=E.ID(+) And b.id=g.��Ŀid(+) And g.����ID(+)=[2] " & _
                    IIF(mint��Դ = 1, " And A.��Ч=1", "") & " And A.�������ID=[1] Order by A.���"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng���˿���id)
    With vsScheme
        .Redraw = flexRDNone
        .Rows = .FixedRows    '����������
        If rsTmp.EOF Then
            .Rows = .FixedRows + 1
        Else
            .Rows = .FixedRows + rsTmp.RecordCount
            lngSel = IIF(Val(Sys.RowValue("������ĿĿ¼", mlng����ID, "ִ�з���") & "") = 1, -1, 0)
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, colѡ��) = lngSel
                .TextMatrix(i, col��Ч) = IIF(NVL(rsTmp!��Ч, 0) = 0, "����", "��ʱ")
                .TextMatrix(i, col����) = NVL(rsTmp!ҽ������, NVL(rsTmp!����))
                .Cell(flexcpData, i, col����) = .TextMatrix(i, col����)
                .TextMatrix(i, col�걾��λ) = NVL(rsTmp!�걾��λ)
                .Cell(flexcpData, i, col�걾��λ) = .TextMatrix(i, col�걾��λ)
                .TextMatrix(i, col��鷽��) = NVL(rsTmp!��鷽��)
                .Cell(flexcpData, i, col��鷽��) = .TextMatrix(i, col��鷽��)
                .RowData(i) = 0
                '����
                If InStr(",5,6,", rsTmp!���) > 0 Then
                    '��ҩ����������,�����۵�λ���,��װ��λ��ʾ
                    If Not IsNull(rsTmp!�ܸ�����) And Not IsNull(rsTmp!��װϵ��) Then
                        .TextMatrix(i, col����) = FormatEx(rsTmp!�ܸ����� / rsTmp!��װϵ��, 5)
                    End If
                    If NVL(rsTmp!��Ч, 0) = 1 Then
                        .TextMatrix(i, col������λ) = NVL(rsTmp!��װ��λ)
                    End If
                Else
                    '�����������ҩ����������
                    If Not IsNull(rsTmp!�ܸ�����) Then
                        .TextMatrix(i, col����) = rsTmp!�ܸ�����
                    End If
                    If rsTmp!��� = "E" And NVL(rsTmp!������, 0) = 0 _
                       And Val(.TextMatrix(i - 1, col���)) = rsTmp!��� _
                       And InStr(",7,E,", .TextMatrix(i - 1, col���)) > 0 Then
                        .TextMatrix(i, col������λ) = "��"    '��ҩ�䷽������λΪ"��"
                    ElseIf NVL(rsTmp!��Ч, 0) = 1 Then
                        If rsTmp!��� = "4" Then
                            .TextMatrix(i, col������λ) = NVL(rsTmp!ɢװ��λ)
                        Else
                            .TextMatrix(i, col������λ) = NVL(rsTmp!���㵥λ)
                        End If
                    End If
                End If

                '����
                .TextMatrix(i, col����) = FormatEx(NVL(rsTmp!��������), 4)
                If Not IsNull(rsTmp!��������) Then
                    If rsTmp!��� = "4" Then
                        .TextMatrix(i, col������λ) = NVL(rsTmp!ɢװ��λ)
                    Else
                        .TextMatrix(i, col������λ) = NVL(rsTmp!���㵥λ)
                    End If
                End If
                .TextMatrix(i, col����) = NVL(rsTmp!����)
                .TextMatrix(i, colƵ��) = NVL(rsTmp!ִ��Ƶ��)
                .TextMatrix(i, col����) = NVL(rsTmp!ҽ������)
                .TextMatrix(i, colִ��ʱ��) = NVL(rsTmp!ʱ�䷽��)
                .TextMatrix(i, colִ�п���) = NVL(rsTmp!ִ�п���)
                .Cell(flexcpData, i, colִ������) = NVL(rsTmp!ִ������, 0)
                .TextMatrix(i, col���) = rsTmp!���
                .TextMatrix(i, col���) = NVL(rsTmp!������)
                .TextMatrix(i, col��ĿID) = NVL(rsTmp!������ĿID)
                .TextMatrix(i, col���) = rsTmp!���
                .TextMatrix(i, col�������) = NVL(rsTmp!�������)
                .TextMatrix(i, col��ֵ����) = NVL(rsTmp!��ֵ����)
                .TextMatrix(i, colִ�б��) = rsTmp!ִ�б�� & ""
                .TextMatrix(i, col�Ա�) = Decode(Val(rsTmp!�Ա�), 0, "δ֪", 1, "��", 2, "Ů")
                .TextMatrix(i, col����Ӧ��) = rsTmp!����Ӧ�� & ""
                .TextMatrix(i, col��������) = rsTmp!�������� & ""
                
                
                '�жϷ�Ժ��ִ��ҩƷ�͸������������Ƿ��п��
                If mlng���˿���id <> 0 And InStr(",4,5,6,7,", rsTmp!��� & "") > 0 Then
                    strDepartments = ""
                    If Val(rsTmp!ִ������ & "") <> 5 And InStr(",5,6,7,", rsTmp!��� & "") > 0 And Val(rsTmp!�շ���ĿID & "") <> 0 Then
                        strDepartments = Get����ҩ��IDs(rsTmp!��� & "", NVL(rsTmp!������ĿID), Val(rsTmp!�շ���ĿID & ""), mlng���˿���id, mint��Դ)
                    ElseIf Val(rsTmp!��������) = 1 And rsTmp!��� & "" = "4" Then
                        strDepartments = Get���÷��ϲ���IDs(Val(rsTmp!�շ���ĿID & ""), mlng���˿���id, mint��Դ)
                    End If
                    '�жϿ���Ƿ��������
                    If strDepartments <> "" Then
                        If GetStock(Val(rsTmp!�շ���ĿID & ""), , mint��Դ, strDepartments, CDbl(Val(.TextMatrix(i, col����)))) = 0 Then
                            .TextMatrix(i, colѡ��) = 0
                            .Cell(flexcpBackColor, i, 0, i, col�Ƿ�����) = &H8000000F
                            If InStr(",5,6,7,", rsTmp!��� & "") > 0 Then
                                .Cell(flexcpData, i, col�Ƿ�����) = 1
                            End If
                        End If
                    Else
                        .TextMatrix(i, colѡ��) = 0
                        .Cell(flexcpBackColor, i, 0, i, col�Ƿ�����) = &H8000000F
                    End If
                End If
                '���������ָ�������ÿ��ң�������Ϊ���ɫ
                If Val(rsTmp!���ÿ����� & "") > 0 And rsTmp!���ÿ���ID & "" = "" Then
                    .TextMatrix(i, colѡ��) = 0
                    .TextMatrix(i, col�Ƿ�����) = "1"
                    '                    .Cell(flexcpBackColor, i, 0, i, col�Ƿ�����) = &HC0C0C0
                End If
                '�����ʾ��Ϊ�գ�������ҩ�䷽����ͣ����ҩ��巨
                .TextMatrix(i, col��ʾ) = rsTmp!��ʾ & ""
                If rsTmp!��ʾ & "" <> "" Then
                    .TextMatrix(i, colѡ��) = 0
                End If
                '���Ȩ��
                If mbln������Ȩ�� And .TextMatrix(i, col�������) = "����ҩ" Or _
                   mbln������Ȩ�� And .TextMatrix(i, col�������) = "����ҩ" Or _
                   mbln������Ȩ�� And (.TextMatrix(i, col�������) = "����I��") Or _
                   mbln������Ȩ�� And (.TextMatrix(i, col��ֵ����) = "����" Or .TextMatrix(i, col��ֵ����) = "����") Then
                    .TextMatrix(i, colѡ��) = 0
                End If

                '��Ѫҽ����飬�����м�������רҵ����ְ���ҽʦ�������´�
                If .TextMatrix(i, col���) = "K" And gbln��Ѫ�����м����� Then
                    If UserInfo.רҵ����ְ�� <> "����ҽʦ" And UserInfo.רҵ����ְ�� <> "����ҽʦ" And UserInfo.רҵ����ְ�� <> "������ҽʦ" Then
                        .TextMatrix(i, colѡ��) = 0
                    End If
                End If

                '��ǰ������г����򲻷������Ŀ
                If Not IsNull(rsTmp!������ĿID) Then
                    If Not (IsNull(rsTmp!����ʱ��) Or Format(NVL(rsTmp!����ʱ��), "yyyy-MM-dd") = "3000-01-01") Then
                        .RowData(i) = 1
                    ElseIf Not (NVL(rsTmp!�������, 0) = 3 Or InStr(str�������, "," & NVL(rsTmp!�������, 0) & ",") > 0) Or mint��Դ = 3 Then
                        .RowData(i) = 1
                    ElseIf Not IsNull(rsTmp!��װ��λ) Or rsTmp!��� & "" = "4" Then
                        '��ҩƷ,ͬʱҪ�жϵ��շ���ĿĿ¼
                        If Not (IsNull(rsTmp!�շѳ���) Or Format(NVL(rsTmp!�շѳ���), "yyyy-MM-dd") = "3000-01-01") Then
                            .RowData(i) = 1
                        ElseIf Not (NVL(rsTmp!�շѷ���, 0) = 3 Or NVL(rsTmp!�շѷ���, 0) = mint��Դ) Or mint��Դ = 3 Then
                            .RowData(i) = 1
                        End If
                    End If
                End If

                rsTmp.MoveNext
            Next

            '���һ��ҩƷ����һ��Ϊ��ѡ��(û�п��)�����ͬ�������ҩƷ�͸�ҩ;��Ҳ����Ϊ��ѡ��
            For i = 1 To .Rows - 1
                If .TextMatrix(i, colѡ��) = 0 Then
                    For j = 1 To .Rows - 1
                        'ҩƷ
                        If .TextMatrix(i, col���) <> "" And (.TextMatrix(j, col���) = .TextMatrix(i, col���) Or .TextMatrix(j, col���) = .TextMatrix(i, col���)) Or _
                           .TextMatrix(i, col���) = "" And .TextMatrix(i, col���) = .TextMatrix(j, col���) Then
                           If .TextMatrix(i, col���) = "7" And .TextMatrix(j, col���) = "" Then
                                .Cell(flexcpBackColor, j, 0, j, col�Ƿ�����) = &H8000000F
                           End If
                           .TextMatrix(j, colѡ��) = 0
                        End If
                    Next
                End If
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
                            .TextMatrix(j, col�÷�) = .TextMatrix(i, col����) & .TextMatrix(i, col����)

                            '��ʾ��ҩִ������
                            If Val(.Cell(flexcpData, j, colִ������)) = 5 And Val(.Cell(flexcpData, i, colִ������)) <> 5 Then
                                .TextMatrix(j, colִ������) = IIF(Val(.TextMatrix(j, colִ�б��)) = 2, "��ȡҩ", "�Ա�ҩ")
                            ElseIf Val(.Cell(flexcpData, j, colִ������)) <> 5 And Val(.Cell(flexcpData, i, colִ������)) = 5 Then
                                .TextMatrix(j, colִ������) = "��Ժ��ҩ"
                            Else
                                .TextMatrix(j, colִ������) = IIF(Val(.TextMatrix(j, colִ�б��)) = 1, "��ȡҩ", "����")
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If

                '��Ѫ;��
                If .TextMatrix(i, col���) = "E" And .TextMatrix(i - 1, col���) = "K" _
                   And Val(.TextMatrix(i, col���)) = Val(.TextMatrix(i - 1, col���)) Then
                    .RowHidden(i) = True
                    .TextMatrix(i - 1, col�÷�) = .TextMatrix(i, col����)
                    .TextMatrix(i - 1, col����) = .TextMatrix(i - 1, col����) & "(" & .TextMatrix(i, col����) & ")"
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
                            .TextMatrix(i, colִ������) = IIF(Val(.TextMatrix(j, colִ�б��)) = 2, "��ȡҩ", "�Ա�ҩ")
                        ElseIf Val(.Cell(flexcpData, j, colִ������)) <> 5 And Val(.Cell(flexcpData, i, colִ������)) = 5 Then
                            .TextMatrix(i, colִ������) = "��Ժ��ҩ"
                        Else
                            .TextMatrix(i, colִ������) = IIF(Val(.TextMatrix(j, colִ�б��)) = 1, "��ȡҩ", "����")
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
                            str�걾 = .TextMatrix(j, col�걾��λ)    'ȡ��һ��������Ŀ�ı걾
                        ElseIf .TextMatrix(j, col���) = "E" And Val(.TextMatrix(j, col���)) <> 0 Then
                            str�巨 = .TextMatrix(j, col����)
                        End If
                    Next

                    .TextMatrix(i, col�÷�) = .TextMatrix(i, col����)    '��ʾ��ҩ�÷������ɼ�����

                    If .TextMatrix(i - 1, col���) = "C" Then
                        .TextMatrix(i, col����) = Mid(strTmp, 2) & IIF(str�걾 <> "", "(" & str�걾 & ")", "")
                    Else
                        .TextMatrix(i, col����) = "��ҩ�䷽," & .TextMatrix(i, colƵ��) & "," & _
                                                str�巨 & "," & .TextMatrix(i, col����) & ":" & Mid(str��ҩ, 2)
                    End If
                End If

                '������
                If .TextMatrix(i, col���) = "D" And Val(.TextMatrix(i, col���)) = 0 Then
                    str�걾 = "": str�巨 = "": strTmp = ""
                    For j = i + 1 To .Rows - 1
                        If Val(.TextMatrix(j, col���)) = Val(.TextMatrix(i, col���)) Then
                            .RowHidden(j) = True
                            If .TextMatrix(j, col�걾��λ) <> "" _
                               And Val(.TextMatrix(j, col��ĿID)) = Val(.TextMatrix(i, col��ĿID)) Then    '��ͬ����ĿID�����·�ʽ
                                If .TextMatrix(j, col�걾��λ) <> strTmp And strTmp <> "" Then
                                    str�걾 = str�걾 & "," & strTmp & IIF(str�巨 <> "", "(" & Mid(str�巨, 2) & ")", "")
                                    str�巨 = ""
                                End If
                                If .TextMatrix(j, col��鷽��) <> "" Then
                                    str�巨 = str�巨 & "," & .TextMatrix(j, col��鷽��)
                                End If

                                strTmp = .TextMatrix(j, col�걾��λ)
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    If strTmp <> "" Then
                        str�걾 = str�걾 & "," & strTmp & IIF(str�巨 <> "", "(" & Mid(str�巨, 2) & ")", "")
                    End If
                    If str�걾 <> "" Then
                        .TextMatrix(i, col����) = .TextMatrix(i, col����) & ":" & Mid(str�걾, 2)
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
                    Call SelGroup(i, 0, lngEnd)
                End If
                If i < lngEnd Then i = lngEnd
            Next
        End If
        .ColHidden(col��Ч) = mint��Դ = 1
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

Private Function CheckCanSelRow(ByVal lngRow As Long) As String
'����:��ָ֤�����Ƿ����ѡ��
    Dim lngCol As Long
    Dim strContent As String
    
    With vsScheme
        If .TextMatrix(lngRow, col���) = "D" Then
            strContent = "[" & Trim(.Cell(flexcpData, lngRow, col�걾��λ)) & "]" & Trim(.Cell(flexcpData, lngRow, col��鷽��))
            If strContent <> "[]" Then
                strContent = Chr(34) & strContent & Chr(34)
            Else
                strContent = Chr(34) & .Cell(flexcpData, lngRow, col����) & Chr(34)
            End If
        Else
            strContent = Chr(34) & .Cell(flexcpData, lngRow, col����) & Chr(34)
        End If
        If .RowData(lngRow) = 1 Then
            CheckCanSelRow = strContent & "(�ѳ����򲻷���ǰ����)": Exit Function
        End If
        
        If .TextMatrix(lngRow, col�Ƿ�����) = "1" Then
            CheckCanSelRow = strContent & "(�������ڵ�ǰ����)": Exit Function
        End If
        
        If InStr("δ֪" & mstr�Ա�, .TextMatrix(lngRow, col�Ա�)) = 0 Then
            CheckCanSelRow = strContent & "(�������ڵ�ǰ�����Ա�)": Exit Function
        End If
        
        If mbln������Ȩ�� And .TextMatrix(lngRow, col�������) = "����ҩ" Then
            CheckCanSelRow = strContent & "(��������ҩƷȨ��)": Exit Function
        End If
        
        If mbln������Ȩ�� And .TextMatrix(lngRow, col�������) = "����ҩ" Then
            CheckCanSelRow = strContent & "(�޶���ҩƷȨ��)": Exit Function
        End If
        
        If mbln������Ȩ�� And (.TextMatrix(lngRow, col�������) = "����I��") Then
            CheckCanSelRow = strContent & "(�޾�����ҩƷȨ��)": Exit Function
        End If
        
        If mbln������Ȩ�� And (.TextMatrix(lngRow, col��ֵ����) = "����" Or .TextMatrix(lngRow, col��ֵ����) = "����") Then
            CheckCanSelRow = strContent & "(�޹�����ҩƷȨ��)": Exit Function
        End If
        
        '��Ѫҽ����飬�����м�������רҵ����ְ���ҽʦ�������´�
        If .TextMatrix(lngRow, col���) = "K" And gbln��Ѫ�����м����� Then
            If UserInfo.רҵ����ְ�� <> "����ҽʦ" And UserInfo.רҵ����ְ�� <> "����ҽʦ" And UserInfo.רҵ����ְ�� <> "������ҽʦ" Then
                CheckCanSelRow = Trim(.Cell(flexcpText, lngRow, col����) & "") & "(���м�������רҵ����ְ��)": Exit Function
            End If
        End If
    End With
End Function

Private Sub GetRowScope(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long, Optional lngRow��� As Long)
'����:��ȡһ��ҽ������ֹλ�ã�ͬʱ��ȡ��ҽ���к�
'����:
'   lngRow ��ǰ��
'����:
'   lngBegin ��ʼ��
'   lngEnd ��ֹ��
'   lngRow��� ��ҽ����

    Dim i As Long, lng��� As Long

    With vsScheme
        If .TextMatrix(lngRow, col���) = "" Then '����¼��
            lngRow��� = lngRow: lngBegin = lngRow: lngEnd = lngRow
            Exit Sub
        End If
        '��ȡ������
        If Val(.TextMatrix(lngRow, col���)) <> 0 Then
            lng��� = Val(.TextMatrix(lngRow, col���))
            lngRow��� = .FindRow(lng���, , col���, , True)
            If lngRow��� = -1 Then
                lngRow��� = lngRow
            End If
        Else
            lng��� = Val(.TextMatrix(lngRow, col���)): lngRow��� = lngRow
        End If
        
        lngBegin = lngRow���: lngEnd = lngRow���
        
        For i = lngRow��� - 1 To .FixedRows Step -1
            If Val(.TextMatrix(i, col���)) = lng��� Then
                lngBegin = i
            Else
                Exit For
            End If
        Next

        For i = lngRow��� + 1 To .Rows - 1
            If Val(.TextMatrix(i, col���)) = lng��� Then
                lngEnd = i
            Else
                Exit For
            End If
        Next
    End With
End Sub

Private Function CheckCanSelGroup(ByVal lngRow As Long, Optional ByVal blnAsk As Boolean = True) As Boolean
'���ܣ��жϱ���ҽ���Ƿ����ѡ��
'������
'   lngRow ��ǰ��
'   blnAsk �Ƿ������ʾ��ѯ�ʣ�ȫѡʱ��������ѯ��)
    Dim i As Long, strResult As String
    Dim lngBegin As Long, lngEnd As Long, lngRow��� As Long
    Dim bln�䷽ As Boolean, bln���� As Boolean, blnCanSel As Boolean
    Dim strMsg As String
    Dim blnMedicineAdvice As Boolean
    
    With vsScheme
        '��ȡ����ҽ����Ϣ
        Call GetRowScope(lngRow, lngBegin, lngEnd, lngRow���)
        
         '����Ƿ���ҽ����������Ŀδ��ѡ�����Ե���Ӧ�á���δ��ѡ�Ĳ������ơ�
        If lngBegin = lngEnd Then
            If Val(.TextMatrix(lngRow, col����Ӧ��)) = 0 And Val(.TextMatrix(lngRow, col��ĿID)) <> 0 Then
                If blnAsk Then
                    MsgBox "ҽ����" & .TextMatrix(lngRow, col����) & "����Ӧ��������Ŀ���ܵ���Ӧ�ã������Ա�ѡ���������ʣ�����ϵ����Ա��", vbInformation, gstrSysName
                End If
                Exit Function
            End If
        Else
            For i = lngBegin To lngEnd
                If InStr(",5,6,7,", .TextMatrix(i, col���)) > 0 Then
                    blnMedicineAdvice = True
                End If
            Next
            If Not blnMedicineAdvice Then
                For i = lngBegin To lngEnd
                    If Not (.TextMatrix(i, col���) = "G" Or (.TextMatrix(i, col���) = "E" And InStr(",2,3,4,6,7,8,", .TextMatrix(i, col��������)) > 0)) Then
                        If Val(.TextMatrix(i, col����Ӧ��)) = 0 And Val(.TextMatrix(i, col��ĿID)) <> 0 Then
                            If blnAsk Then
                                MsgBox "ҽ����" & .TextMatrix(i, col����) & "����Ӧ��������Ŀ���ܵ���Ӧ�ã������Ա�ѡ���������ʣ�����ϵ����Ա��", vbInformation, gstrSysName
                            End If
                            Exit Function
                        End If
                    End If
                Next
            End If
        End If
        
        '�����ò��� ָ��ҩ��ʱ���ƿ�� ������£��������´��治���ҽ��
        If gblnStock Then
            For i = lngBegin To lngEnd
                If Val(.Cell(flexcpData, i, col�Ƿ�����)) = 1 Then
                    If Val(.TextMatrix(lngBegin, col���)) = 7 Then
                        strMsg = strMsg & "," & .TextMatrix(i, col����)
                    Else
                        If blnAsk Then
                            MsgBox "��ҩƷ��治��,ϵͳ�����˲������´��治���ҩƷ�����ܱ�ѡ��", vbInformation, gstrSysName
                        End If
                        Exit Function
                    End If
                End If
            Next
        End If
        If strMsg <> "" Then
            MsgBox "���䷽�д��ڿ�治���ҩƷ(" & Mid(strMsg, 2) & ")��", vbInformation, gstrSysName
        End If
            
        strMsg = CheckCanSelRow(lngRow���)
        If strMsg <> "" Then '��ҽ������ҽ�����
            If blnAsk Then
                MsgBox "��ҽ����" & vbNewLine & strMsg & vbNewLine & "��Ч,���ܱ�ѡ��", vbInformation, gstrSysName
            End If
            Exit Function
        Else
            If lngBegin <> lngEnd Then
                If .TextMatrix(lngRow���, col���) = "E" Then
                    If lngRow��� - 2 >= lngBegin Then
                        If .TextMatrix(lngRow��� - 2, col���) = "7" And .TextMatrix(lngRow��� - 1, col���) = "E" Then '��ҩ�䷽�ļ������
                            strMsg = CheckCanSelRow(lngRow��� - 1)
                            If strMsg <> "" Then
                                If blnAsk Then
                                    MsgBox "����ҩ�䷽�м巨:" & vbNewLine & strMsg & vbNewLine & "��Ч,���ܱ�ѡ��", vbInformation, gstrSysName
                                End If
                                Exit Function
                            Else
                                bln�䷽ = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
        strMsg = ""
        '��ҽ��ȫ�����
        If lngBegin <> lngEnd Then
            For i = lngBegin To lngEnd
                If .TextMatrix(lngRow���, col���) = "F" Then blnCanSel = True '����ҽ����ҽ�����þͿ�ѡ
                If Not (i = lngRow��� Or bln�䷽ And i = lngRow��� - 1) Then
                    strResult = CheckCanSelRow(i)
                    If .TextMatrix(i, col���) = "C" Then bln���� = True
                    If strResult <> "" Then
                        strMsg = IIF(strMsg = "", "", strMsg & "��" & vbNewLine) & strResult
                    Else
                        If bln�䷽ Then  '��ҩ�䷽��һζ��ҩ���þͿ�ѡ���巨�Լ��÷�ǰ���Ѿ��жϣ�����������ֻҪһ����ҽ�����þͿ�ѡ
                            If .TextMatrix(i, col���) = "7" Then
                                blnCanSel = True
                            End If
                        Else
                            blnCanSel = True
                        End If
                    End If
                End If
            Next
        Else
            blnCanSel = True '����ҽ������ڸ�ҽ��ʱ�Ѿ����
        End If
        
        If Not blnCanSel Then strMsg = ""
        '��ҩ�䷽δ��ȡ��ҩƷ��Ϣ
        If .TextMatrix(lngRow���, col��ʾ) <> "" Then
            If (bln���� Or bln�䷽) And blnCanSel Then
                If bln�䷽ Then strMsg = strMsg & IIF(strMsg <> "", "��" & vbNewLine, "") & vsScheme.TextMatrix(lngRow���, col��ʾ) & "(��ͣ�û�û�п��ù��)"
            ElseIf bln�䷽ Then
                blnCanSel = False
                strMsg = "����ҩ�䷽��������ҩ�Ѿ���ͣ�û�û�п��ù��,���ܱ�ѡ��"
            ElseIf bln���� Then
                blnCanSel = False
                strMsg = "�ü�����������м�����Ŀ�Ѿ���ͣ��,���ܱ�ѡ��"
            End If
        End If
        If Not blnCanSel And strMsg = "" Then strMsg = "��ҽ���в�������Ч��Ŀ,���ܱ�ѡ��"

        If blnCanSel Then
            If strMsg <> "" Then
                If blnAsk Then
                    If MsgBox(IIF(InStr(1, strMsg, "��") > 0, "��ҽ����:" & vbNewLine & strMsg & vbNewLine & "��Ч,��Щ��Ŀ", "��ҽ����:" & vbNewLine & strMsg & vbNewLine & "��Ч,����Ŀ") & "���ᱻѡ��,�Ƿ�ѡ���ҽ����", _
                        vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        CheckCanSelGroup = True
                    End If
                End If
            Else
                CheckCanSelGroup = True
            End If
        Else
            If blnAsk Then
                MsgBox strMsg, vbInformation, gstrSysName
            End If
        End If
    End With
End Function

Private Sub SelGroup(ByVal lngRow As Long, ByVal intѡ�� As Integer, Optional ByRef lngEnd As Long)
'����:�������ѡ�����ҽ��
'������
'   lngRow ��ǰ��
'   lngEnd ����ҽ�����һ��
'   intѡ�� ѡ���� -1,���ѡ�񣨿�ѡ��ѡ�񣬲���ѡ��ѡ��),0��ѡ��,1��ȫѡ�����
    Dim lngBegin As Long
    Dim i As Long
    
    With vsScheme
    
        '��ȡ����ҽ����Ϣ
        Call GetRowScope(lngRow, lngBegin, lngEnd)
        'ѡ���ȡ��ѡ��
        If intѡ�� = -1 Then 'checkCanSelGroup(i,true)=true�����
            For i = lngBegin To lngEnd
                If CheckCanSelRow(i) = "" Then
                    .TextMatrix(i, colѡ��) = intѡ��
                Else
                    .TextMatrix(i, colѡ��) = 0
                End If
            Next
        Else 'checkCanSelGroup(i,false)=true ����û�ȡ��ѡ���ʹ��
            intѡ�� = intѡ�� * -1
            For i = lngBegin To lngEnd
                .TextMatrix(i, colѡ��) = intѡ��
            Next
        End If
    End With
End Sub

