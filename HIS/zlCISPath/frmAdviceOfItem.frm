VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAdviceOfItem 
   BorderStyle     =   0  'None
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   14070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   13995
      _cx             =   24686
      _cy             =   7011
      Appearance      =   2
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
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16444122
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   19
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAdviceOfItem.frx":0000
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
Attribute VB_Name = "frmAdviceOfItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Enum CONST_COL
    col��Ч = 0
    col���� = 1
    col���� = 2
    col���� = 3
    colƵ�� = 4
    col�÷� = 5
    col���� = 6
    colִ�п��� = 7
    colִ������ = 8
    colID = 9
    col���ID = 10
    col������ĿID = 11
    col������� = 12
    col�շ�ϸĿID = 13
    col�걾��λ = 14
    col��鷽�� = 15
    colִ��ʱ�� = 16
    col_��ʼִ��ʱ�� = 17
    col_��ֹʱ�� = 18
End Enum

Public Sub ShowAdvice(ByVal bytUseType As Byte, Optional ByVal strSQL As String, Optional ByVal lng·��ִ��ID As Long, Optional ByVal strҽ��IDs As String)
'���ܣ�·����Ŀ����ʱ����·������ѡ��һ��·����Ŀʱ����ʾ��Ӧ��ҽ���嵥
'������
'      bytUseType��     0-·����Ŀ����ʱ��ʾҽ��,1-����·��ִ����Ŀ��·������ʾҽ���嵥,2-��ӻ��޸�·������Ŀ��ʾҽ��
'      strSQL��         bytUseType=0ʱ���룬ҽ���嵥����Դ,�����ʱ����������
'      lng·��ִ��ID��  bytUseType=1ʱ���룬����·��ִ����Ŀ��ID
'      strҽ��IDs��     bytUseType=2ʱ���룬��ǰ��ӵ�ҽ��ID��
    Dim rsTmp As ADODB.Recordset
    Dim blnClear As Boolean
    
    If bytUseType = 0 Then
        If strSQL = "" Then blnClear = True
    ElseIf bytUseType = 1 Then
        If lng·��ִ��ID = 0 Then blnClear = True
    ElseIf bytUseType = 2 Then
        If strҽ��IDs = "" Then blnClear = True
    End If
    If blnClear Then
        vsAdvice.Rows = vsAdvice.FixedRows
        vsAdvice.Rows = vsAdvice.FixedRows + 1 '��һ�հ���
        Exit Sub
    End If
        
    If bytUseType <> 0 Then
        If bytUseType = 1 Then
            strSQL = "Select A.* From ����ҽ����¼ A,����·��ҽ�� B Where B.·��ִ��ID = [1] And A.ID = B.����ҽ��ID"
        Else
            strSQL = "Select * From ����ҽ����¼ a,Table(f_Num2list([2])) b Where a.ID = b.Column_value"
        End If
    End If
    
    '����SQL�����NULL�ֶ��ұ�(+)CBO�����������
    strSQL = "Select " & IIf(bytUseType = 2, "/*+ rule*/", "") & "A.ID,A.���ID,A.���," & IIf(bytUseType = 0, "A.��Ч", "A.ҽ����Ч") & " as ��Ч,A.������ĿID,A.ҽ������," & _
        " A.��������,A.ִ��Ƶ��,A.ҽ������,Nvl(C.����,Decode(Nvl(A.ִ������,0),0,'<����>',5,'<Ժ��ִ��>')) as ִ�п���," & _
        " A.ִ������, " & IIf(bytUseType = 0, "A.ʱ�䷽��", "A.ִ��ʱ�䷽��") & " as ʱ�䷽��,Nvl(B.���,'*') as �������,Nvl(D.����||Decode(D.���,NULL,NULL,' '||D.���),B.����) as ����," & _
        " B.���㵥λ,A.�걾��λ,A.��鷽��,A.�ܸ�����,D.���㵥λ as ������λ,D.ID as �շ�ϸĿID," & _
        " Nvl(B.����ʱ��,To_Date('3000-01-01','YYYY-MM-DD')) As ����ʱ��" & _
        IIf(bytUseType = 0, "", ",To_Char(A.��ʼִ��ʱ��,'YYYY-MM-DD HH24:MI') as ��ʼʱ��,To_Char(A.ִ����ֹʱ��,'YYYY-MM-DD HH24:MI') as ��ֹʱ��") & _
        IIf(bytUseType = 1, " ,a.ҽ��״̬", "") & _
        " From (" & strSQL & ") A,������ĿĿ¼ B,���ű� C,�շ���ĿĿ¼ D" & _
        " Where Nvl(A.������ĿID,-1)=B.ID(+) And Nvl(A.ִ�п���ID,-1)=C.ID(+) And Nvl(A.�շ�ϸĿID,-1)=D.ID(+)" & _
        " Order by A.���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ShowAdvice", lng·��ִ��ID, strҽ��IDs)
    Call LoadAdvice(rsTmp, bytUseType)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetAdviceTitle(Optional ByVal lngRows As Long = 5) As String
'���ܣ���ȡҽ������ҽ�����ݵ�����ַ���(���lngRows��)
    Dim strItem As String, i As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Not .RowHidden(i) Then
                If UBound(Split(strItem, "��")) + 1 > lngRows Then
                    strItem = strItem & "......"
                    Exit For
                Else
                    strItem = strItem & "��" & .TextMatrix(i, col����)
                End If
            End If
        Next
    End With
    GetAdviceTitle = Mid(strItem, 2)
End Function

Private Sub LoadAdvice(ByRef rsTmp As ADODB.Recordset, ByVal bytUseType As Byte)
'���ܣ���ʾ·����Ŀ��Ӧ��ҽ������
    Dim strTmp As String
    Dim str��ҩ As String, str�巨 As String
    Dim str���� As String, str�걾 As String
    Dim strFilter As String
    Dim i As Long, j As Long
    
    With vsAdvice
        .Redraw = flexRDNone
        .Rows = .FixedRows '����������
        .Rows = .FixedRows + rsTmp.RecordCount
        If bytUseType = 0 Then      '��Ŀҽ������
            .ColHidden(col_��ʼִ��ʱ��) = True
            .ColHidden(col_��ֹʱ��) = True
        ElseIf bytUseType = 2 Then  '���·������Ŀ
            .ColHidden(col_��ֹʱ��) = True
        End If
        
        For i = 1 To rsTmp.RecordCount
            .TextMatrix(i, col��Ч) = IIf(Nvl(rsTmp!��Ч, 0) = 0, "����", "��ʱ")
            .TextMatrix(i, col����) = Nvl(rsTmp!ҽ������, Nvl(rsTmp!����))
            If bytUseType = 1 Then
                If rsTmp!ҽ��״̬ = 4 Then
                    .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H808080 '��ɫ
                    .Cell(flexcpFontStrikethru, i, .FixedCols, i, .Cols - 1) = True
                End If
            End If
            
            .TextMatrix(i, col�걾��λ) = Nvl(rsTmp!�걾��λ) '����걾
            .TextMatrix(i, col��鷽��) = Nvl(rsTmp!��鷽��)
            .TextMatrix(i, col����) = FormatEx(Nvl(rsTmp!��������), 4)
            If Not IsNull(rsTmp!��������) Then
                If rsTmp!������� = "4" Then
                    .TextMatrix(i, col����) = .TextMatrix(i, col����) & Nvl(rsTmp!������λ)
                Else
                    .TextMatrix(i, col����) = .TextMatrix(i, col����) & Nvl(rsTmp!���㵥λ)
                End If
            End If
            If .TextMatrix(i, col��Ч) = "��ʱ" Then
                If Not IsNull(rsTmp!�ܸ�����) Then
                    .TextMatrix(i, col����) = FormatEx(Nvl(rsTmp!�ܸ�����), 4)
                    If Not IsNull(rsTmp!������λ) Then
                        .TextMatrix(i, col����) = .TextMatrix(i, col����) & Nvl(rsTmp!������λ)
                    ElseIf InStr(",4,5,6,7,", rsTmp!�������) = 0 Then
                        .TextMatrix(i, col����) = .TextMatrix(i, col����) & Nvl(rsTmp!���㵥λ)
                    End If
                End If
            End If
            .TextMatrix(i, colƵ��) = Nvl(rsTmp!ִ��Ƶ��)
            .TextMatrix(i, col����) = Nvl(rsTmp!ҽ������)
            .TextMatrix(i, colִ��ʱ��) = Nvl(rsTmp!ʱ�䷽��)
            .TextMatrix(i, colִ�п���) = Nvl(rsTmp!ִ�п���)
            .Cell(flexcpData, i, colִ������) = Nvl(rsTmp!ִ������, 0)
            .TextMatrix(i, colID) = rsTmp!ID
            .TextMatrix(i, col���ID) = "" & rsTmp!���ID
            .TextMatrix(i, col������ĿID) = "" & rsTmp!������ĿID
            .TextMatrix(i, col�շ�ϸĿID) = "" & rsTmp!�շ�ϸĿID
            .TextMatrix(i, col�������) = rsTmp!�������
            If Format(rsTmp!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &HFF&
            End If
            
            If bytUseType <> 0 Then
                .TextMatrix(i, col_��ʼִ��ʱ��) = "" & rsTmp!��ʼʱ��
                .TextMatrix(i, col_��ֹʱ��) = "" & rsTmp!��ֹʱ��
            End If
            rsTmp.MoveNext
        Next
        
        '�ٴ���һЩ�����е�����,��������ݵ���ʾ
        For i = 1 To .Rows - 1
            '��ҩ;��
            If .TextMatrix(i, col�������) = "E" And Val(.TextMatrix(i, col���ID)) = 0 _
                And Val(.TextMatrix(i - 1, col���ID)) = Val(.TextMatrix(i, colID)) _
                And InStr(",5,6,", .TextMatrix(i - 1, col�������)) > 0 Then
                .RowHidden(i) = True
                '��ʾ��ҩ;��
                For j = i - 1 To .FixedRows Step -1
                    If Val(.TextMatrix(j, col���ID)) = Val(.TextMatrix(i, colID)) Then
                        .TextMatrix(j, col�÷�) = .TextMatrix(i, col����)
                                                    
                        '��ʾ��ҩ��ִ������
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
            
            '��Ѫ;��
            If .TextMatrix(i, col�������) = "E" And .TextMatrix(i - 1, col�������) = "K" _
                And Val(.TextMatrix(i, col���ID)) = Val(.TextMatrix(i - 1, colID)) Then
                .RowHidden(i) = True
                .TextMatrix(i - 1, col�÷�) = .TextMatrix(i, col����)
                .TextMatrix(i - 1, col����) = .TextMatrix(i - 1, col����) & "(" & .TextMatrix(i, col����) & ")"
            End If
            
            '��ҩ�䷽�ͼ������
            If .TextMatrix(i, col�������) = "E" And Val(.TextMatrix(i, col���ID)) = 0 _
                And Val(.TextMatrix(i - 1, col���ID)) = Val(.TextMatrix(i, colID)) _
                And InStr(",7,E,C,", .TextMatrix(i - 1, col�������)) > 0 Then
                
                str��ҩ = "": str�巨 = "": str�걾 = "": strTmp = ""
                j = .FindRow(CStr(Val(.TextMatrix(i, colID))), , col���ID)
                
                '��ҩ�������ִ�п���
                .TextMatrix(i, colִ�п���) = .TextMatrix(j, colִ�п���)
                
                '��ʾ��ҩ�䷽ִ������:��ҩƷΪ׼�ж�
                If .TextMatrix(i - 1, col�������) <> "C" Then
                    If Val(.Cell(flexcpData, j, colִ������)) = 5 And Val(.Cell(flexcpData, i, colִ������)) <> 5 Then
                        .TextMatrix(i, colִ������) = "�Ա�ҩ"
                    ElseIf Val(.Cell(flexcpData, j, colִ������)) <> 5 And Val(.Cell(flexcpData, i, colִ������)) = 5 Then
                        .TextMatrix(i, colִ������) = "��Ժ��ҩ"
                    End If
                End If

                For j = j To i - 1
                    .RowHidden(j) = j <> i
                    If .TextMatrix(j, col�������) = "7" Then
                        str��ҩ = str��ҩ & "," & RTrim(.TextMatrix(j, col����) & _
                            " " & .TextMatrix(j, col����) & " " & .TextMatrix(j, col����))
                    ElseIf .TextMatrix(j, col�������) = "C" Then
                        strTmp = strTmp & "," & .TextMatrix(j, col����)
                        str�걾 = .TextMatrix(j, col�걾��λ) 'ȡ��һ��������Ŀ�ı걾
                    ElseIf .TextMatrix(j, col�������) = "E" And Val(.TextMatrix(j, col���ID)) <> 0 Then
                        str�巨 = .TextMatrix(j, col����)
                    End If
                Next
                
                .TextMatrix(i, col�÷�) = .TextMatrix(i, col����) '��ʾ��ҩ�÷������ɼ�����
                
                If .TextMatrix(i - 1, col�������) = "C" Then
                    .TextMatrix(i, col����) = Mid(strTmp, 2) & IIf(str�걾 <> "", "(" & str�걾 & ")", "")
                Else
                    .TextMatrix(i, col����) = "��ҩ�䷽," & .TextMatrix(i, colƵ��) & "," & _
                        str�巨 & "," & .TextMatrix(i, col����) & ":" & Mid(str��ҩ, 2)
                    .TextMatrix(i, col����) = .TextMatrix(i, col����) & "��"
                End If
            End If
            
            '������
            If .TextMatrix(i, col�������) = "D" And Val(.TextMatrix(i, col���ID)) = 0 Then
                str�걾 = "": str�巨 = "": strTmp = ""
                For j = i + 1 To .Rows - 1
                    If Val(.TextMatrix(j, col���ID)) = Val(.TextMatrix(i, colID)) Then
                        .RowHidden(j) = True
                        If .TextMatrix(j, col�걾��λ) <> "" _
                            And Val(.TextMatrix(j, col������ĿID)) = Val(.TextMatrix(i, col������ĿID)) Then '��ͬ����ĿID�����·�ʽ
                            If .TextMatrix(j, col�걾��λ) <> strTmp And strTmp <> "" Then
                                str�걾 = str�걾 & "," & strTmp & IIf(str�巨 <> "", "(" & Mid(str�巨, 2) & ")", "")
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
                    str�걾 = str�걾 & "," & strTmp & IIf(str�巨 <> "", "(" & Mid(str�巨, 2) & ")", "")
                End If
                If str�걾 <> "" Then '��ǰ�ļ�鷽ʽʱ����ʾ��ϸҽ������
                    .TextMatrix(i, col����) = .TextMatrix(i, col����) & ":" & Mid(str�걾, 2)
                End If
            End If
            
            '������Ŀ
            If .TextMatrix(i, col�������) = "F" And Val(.TextMatrix(i, col���ID)) = 0 Then
                strTmp = "": str���� = ""
                For j = i + 1 To .Rows - 1
                    If Val(.TextMatrix(j, col���ID)) = Val(.TextMatrix(i, colID)) Then
                        .RowHidden(j) = True
                        If .TextMatrix(j, col�������) = "F" Then
                            strTmp = strTmp & "," & .TextMatrix(j, col����)
                        ElseIf .TextMatrix(j, col�������) = "G" Then
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
        
        If .Rows > .FixedRows Then
            .Row = .FixedRows: .Col = .FixedCols
            .AutoSize col����
        Else
            .Rows = .FixedRows + 1
        End If
        .Redraw = flexRDDirect
    End With
End Sub


Private Sub Form_Resize()
    vsAdvice.Top = Me.ScaleTop + 60
    vsAdvice.Left = Me.ScaleLeft
    vsAdvice.Height = Me.ScaleHeight - 60
    vsAdvice.Width = Me.ScaleWidth
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= vsAdvice.FixedRows And NewCol >= vsAdvice.FixedCols Then
        If NewRow <> OldRow Then
            vsAdvice.ForeColorSel = vsAdvice.CellForeColor
        End If
    End If
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Col = col���� Then
        vsAdvice.AutoSize Col
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(vsAdvice.FixedRows - 1, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
    End If
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        '����һ����ҩ������еı��߼�����
        lngLeft = col��Ч: lngRight = col��Ч
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = colƵ��: lngRight = col�÷�
        End If
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = colִ��ʱ��: lngRight = col_��ֹʱ��
        End If
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        
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
        If Between(Row, .Row, .RowSel) And Me.ActiveControl Is vsAdvice Then
            SetBkColor hDC, SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
    Dim i As Long, blnTmp As Boolean
    
    With vsAdvice
        If .TextMatrix(lngRow, col�������) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, col�������)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInһ����ҩ = blnTmp
    End With
End Function

Private Sub vsAdvice_GotFocus()
    vsAdvice.BackColorSel = &HFFCC99
End Sub

Private Sub vsAdvice_LostFocus()
    vsAdvice.BackColorSel = &HFFEBD7
End Sub
