VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.UserControl UCAdviceList 
   ClientHeight    =   4305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14040
   ScaleHeight     =   4305
   ScaleWidth      =   14040
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13995
      _cx             =   24686
      _cy             =   7435
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
      Cols            =   23
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"UCAdviceList.ctx":0000
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
Attribute VB_Name = "UCAdviceList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Enum CONST_COL
    col��ѡ = 0
    colȱʡ = 1
    col��Ч = 2
    col���� = 3
    col���� = 4
    col���� = 5
    colƵ�� = 6
    col�÷� = 7
    col���� = 8
    colִ�п��� = 9
    colִ������ = 10
    COL_ִ�б�� = 11
    colID = 12
    COL���ID = 13
    col������ĿID = 14
    COL������� = 15
    col�շ�ϸĿID = 16
    col�걾��λ = 17
    col��鷽�� = 18
    colִ��ʱ�� = 19
    col_��ʼִ��ʱ�� = 20
    col_��ֹʱ�� = 21
    Col_�����ĿID = 22
End Enum

Private mbytUseType As Byte '0-·����Ŀ����ʱ��ʾҽ��,1-����·��ִ����Ŀ��·������ʾҽ���嵥,2-��ӻ��޸�·������Ŀ��ʾҽ��,3-��ѡҽ��ѡ��ʱ��ʾ��Ŀ�ı�ѡҽ��,4-�ٴ�·����Ŀ��������
Private mblnReadOnly As Boolean 'mbytUseType=0ʱ����
Public Event DataChange() '�û���ѡ���У��޸�������

Public Sub ShowAdvice(ByVal bytUseType As Byte, Optional ByVal strSql As String, Optional ByVal lng·��ִ��ID As Long, _
    Optional ByVal strҽ��IDs As String, Optional ByVal blnReadOnly As Boolean, Optional ByVal lng·����ĿID As Long, Optional ByVal strSelectedIDs As String, Optional ByVal int���� As Integer = 2)
'���ܣ�·����Ŀ����ʱ����·������ѡ��һ��·����Ŀʱ����ʾ��Ӧ��ҽ���嵥
'������
'      bytUseType��     0-·����Ŀ����ʱ��ʾҽ��,1-����·��ִ����Ŀ��·������ʾҽ���嵥,2-��ӻ��޸�·������Ŀ��ʾҽ��,3-��ѡҽ��ѡ��ʱ��ʾ��Ŀ�ı�ѡҽ��,4-�ٴ�·����Ŀ��������
'      strSQL��         bytUseType=0ʱ���룬ҽ���嵥����Դ,�����ʱ����������
'      lng·��ִ��ID��  bytUseType=1ʱ���룬����·��ִ����Ŀ��ID
'      strҽ��IDs��     bytUseType=2ʱ���룬��ǰ��ӵ�ҽ��ID��
'      blnReadOnly��    bytUseType=0ʱ���룬ֻ���鿴ʱ��������ı䡰ȱʡ�ͱ�ѡ���е�ֵ
'      lng·����ĿID:   bytUseType=3ʱ���룬ѡ��·����Ŀ�ı���ҽ����
'                       bytUseType=0ʱ����,1)δ��˰�����˰��ٴ�·��������ͬ�׶Σ����࣬��Ŀ�����£����ڲ����ҽ����·����Ŀ�ԱȲ鿴ʱ����ʾ��·����Ŀҽ���嵥��;
'                                          2)�鿴·���䶯��¼��ʾ
'      strSelectedIDs:  bytUseType=3ʱ���룬�������Ѿ�ѡ���ҽ������IDs��
'      int����=1-���2-סԺ
    Dim rsTmp As ADODB.Recordset
    Dim blnClear As Boolean
    Dim strӤ��SQL As String
    
    mbytUseType = bytUseType
    mblnReadOnly = blnReadOnly
    
    If bytUseType = 0 Or bytUseType = 4 Then
        If strSql = "" Then blnClear = True
    ElseIf bytUseType = 1 Then
        If lng·��ִ��ID = 0 Then blnClear = True
    ElseIf bytUseType = 2 Then
        If strҽ��IDs = "" Then blnClear = True
    End If
    If blnClear Then
        vsAdvice.Rows = vsAdvice.FixedRows
        vsAdvice.Rows = vsAdvice.FixedRows + 1 '��һ�հ���
        If bytUseType = 4 Then
            vsAdvice.ColHidden(col��ѡ) = True
            vsAdvice.ColHidden(colȱʡ) = True
        End If
        Exit Sub
    End If
        
    If bytUseType <> 0 And bytUseType <> 4 And bytUseType <> 3 Then
        If bytUseType = 1 Then
            strSql = "Select a.*" & vbNewLine & _
                    "From ����ҽ����¼ A, " & IIf(int���� = 1, "��������·��ҽ��", "����·��ҽ��") & " B, ������ĿĿ¼ C" & vbNewLine & _
                    "Where b.·��ִ��id = [1] And Not (c.��� = 'E' And c.�������� = '2') And a.Id = b.����ҽ��id And a.������Ŀid = c.Id" & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select Distinct c.*" & vbNewLine & _
                    "From ����ҽ����¼ A, " & IIf(int���� = 1, "��������·��ҽ��", "����·��ҽ��") & " B, ����ҽ����¼ C" & vbNewLine & _
                    "Where b.·��ִ��id = [1] And a.Id = b.����ҽ��id And c.Id = a.���id And a.������� In ('5', '6')"
        Else
            strSql = "Select /*+ Rule*/ a.*" & vbNewLine & _
                    "From ����ҽ����¼ A, Table(f_Num2list([2])) B, ������ĿĿ¼ C" & vbNewLine & _
                    "Where a.Id = b.Column_Value And Not (c.��� = 'E' And c.�������� = '2') And a.������Ŀid = c.Id" & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select c.*" & vbNewLine & _
                    "From ����ҽ����¼ A, Table(f_Num2list([2])) B, ����ҽ����¼ C" & vbNewLine & _
                    "Where a.Id = b.Column_Value And c.Id = a.���id And a.������� In ('5', '6')"

        End If
        vsAdvice.ColHidden(col��ѡ) = True
		strӤ��SQL = " A.Ӥ��,"
    Else
        If bytUseType = 3 Then
            vsAdvice.ColHidden(col��ѡ) = False
            vsAdvice.ColHidden(colȱʡ) = True
            vsAdvice.TextMatrix(0, col��ѡ) = "ѡ��"
        ElseIf bytUseType = 4 Then
            vsAdvice.ColHidden(col��ѡ) = True
            vsAdvice.ColHidden(colȱʡ) = True
        Else
            vsAdvice.TextMatrix(0, col��ѡ) = "��ѡ"
        End If
    End If
    
    '����SQL�����NULL�ֶ��ұ�(+)CBO�����������
    strSql = "Select " & IIf(bytUseType = 2, "/*+ rule*/", "") & "A.ID,A.���ID,A.���," & IIf(InStr(",0,3,4,", "," & bytUseType & ",") > 0, "A.��Ч", "A.ҽ����Ч") & " as ��Ч,A.������ĿID,A.ҽ������," & _
        " A.��������,A.ִ��Ƶ��,A.ҽ������,Nvl(C.����,Decode(Nvl(A.ִ������,0),0,'<����>',5,'-')) as ִ�п���," & _
        " A.ִ������, A.ִ�б��," & IIf(InStr(",0,3,4,", "," & bytUseType & ",") > 0, "A.ʱ�䷽��", "A.ִ��ʱ�䷽��") & " as ʱ�䷽��,Nvl(B.���,'*') as �������,Nvl(D.����||Decode(D.���,NULL,NULL,' '||D.���),B.����) as ����," & _
        " B.���㵥λ,A.�걾��λ,A.��鷽��,A.�ܸ�����,D.���㵥λ as ������λ,D.ID as �շ�ϸĿID," & _
        " Nvl(B.����ʱ��,To_Date('3000-01-01','YYYY-MM-DD')) As ����ʱ��" & _
        IIf(InStr(",0,3,4,", "," & bytUseType & ",") > 0, ",A.�Ƿ�ȱʡ,A.�Ƿ�ѡ,A.�����ĿID", ",To_Char(A.��ʼִ��ʱ��,'YYYY-MM-DD HH24:MI') as ��ʼʱ��,To_Char(A.ִ����ֹʱ��,'YYYY-MM-DD HH24:MI') as ��ֹʱ��") & _
        IIf(bytUseType = 1, " ,a.ҽ��״̬", "") & _
        " From (" & strSql & ") A,������ĿĿ¼ B,���ű� C,�շ���ĿĿ¼ D" & _
        " Where Nvl(A.������ĿID,-1)=B.ID(+) And Nvl(A.ִ�п���ID,-1)=C.ID(+) And Nvl(A.�շ�ϸĿID,-1)=D.ID(+)" & _
        " Order by " & strӤ��SQL & "A.���"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "ShowAdvice", lng·��ִ��ID, strҽ��IDs, lng·����ĿID, strSelectedIDs)
    Call LoadAdvice(rsTmp)
    
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

Public Function GetAdviceIDSelected(Optional ByVal bytType As Byte, Optional ByRef blnIsAllSelect As Boolean) As String
'���ܣ���ȡѡ���е�ҽ��ID��
'������bytType ��0-ȱʡ�У�1-��ѡ��
'     blnIsAllSelect-�Ƿ����ж�ѡ��
    Dim strItem As String, i As Long
    Dim lngCol As Long
    
    With vsAdvice
        If bytType = 1 Then
            lngCol = col��ѡ
        Else
            lngCol = colȱʡ
        End If
        blnIsAllSelect = True
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, lngCol)) = -1 Then
                strItem = strItem & "," & .TextMatrix(i, colID)
            Else
                If .RowHidden(i) = False Then
                    blnIsAllSelect = False
                End If
            End If
        Next
    End With
    GetAdviceIDSelected = Mid(strItem, 2)
End Function

Public Sub Setѡ���еĿɼ���(ByVal blnHide As Boolean)
'���ܣ���Ŀҽ������ʱ����ѡ���еĿɼ��ԣ�ȫ��ʹ��ʱ���ɼ�������ʹ��ʱ�ſɼ���
'      �鿴�䶯��¼����ѡ���в��ɼ�
    Dim strItem As String, i As Long
    
    With vsAdvice
        .Redraw = flexRDNone
        If blnHide Then
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, colȱʡ) = -1
                .TextMatrix(i, col��ѡ) = 0
            Next
        End If
        .ColHidden(colȱʡ) = blnHide
        .ColHidden(col��ѡ) = blnHide
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub LoadAdvice(ByRef rsTmp As ADODB.Recordset)
'���ܣ���ʾ·����Ŀ��Ӧ��ҽ������
    Dim strTmp As String
    Dim str�巨 As String
    Dim str���� As String, str�걾 As String
    Dim strFilter As String
    Dim i As Long, j As Long, k As Long

    With vsAdvice
        .Redraw = flexRDNone
        .Rows = .FixedRows    '����������
        .Rows = .FixedRows + rsTmp.RecordCount
        If mbytUseType = 0 Or mbytUseType = 3 Or mbytUseType = 4 Then      '��Ŀҽ������
            .ColHidden(col_��ʼִ��ʱ��) = True
            .ColHidden(col_��ֹʱ��) = True

            .Editable = flexEDKbdMouse    '����"ѡ��"��

        ElseIf mbytUseType = 2 Then  '���·������Ŀ
            .ColHidden(col_��ֹʱ��) = True
            .ColHidden(colȱʡ) = True
        Else
            .ColHidden(colȱʡ) = True
        End If

        For i = 1 To rsTmp.RecordCount
            If mbytUseType = 0 Then
                .TextMatrix(i, colȱʡ) = IIf(Val(rsTmp!�Ƿ�ȱʡ & "") = 1, -1, 0)
                .TextMatrix(i, col��ѡ) = IIf(Val(rsTmp!�Ƿ�ѡ & "") = 1, -1, 0)
            ElseIf mbytUseType = 3 Then
                .TextMatrix(i, col��ѡ) = IIf(Val(rsTmp!�Ƿ�ѡ & "") = 1, -1, 0)
            End If

            .TextMatrix(i, col��Ч) = IIf(Nvl(rsTmp!��Ч, 0) = 0, "����", "��ʱ")
            .TextMatrix(i, col����) = Nvl(rsTmp!ҽ������, Nvl(rsTmp!����))
            If mbytUseType = 1 Then
                If rsTmp!ҽ��״̬ = 4 Then
                    .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H808080    '��ɫ
                    .Cell(flexcpFontStrikethru, i, .FixedCols, i, .Cols - 1) = True
                End If
            End If

            .TextMatrix(i, col�걾��λ) = Nvl(rsTmp!�걾��λ)    '����걾
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
            .TextMatrix(i, COL���ID) = "" & rsTmp!���id
            .TextMatrix(i, col������ĿID) = "" & rsTmp!������ĿID
            .TextMatrix(i, col�շ�ϸĿID) = "" & rsTmp!�շ�ϸĿID
            .TextMatrix(i, COL�������) = rsTmp!�������
            If InStr(",E,", .TextMatrix(i, COL�������)) > 0 Then
                .TextMatrix(i, col�÷�) = Nvl(rsTmp!����)
            End If
            .TextMatrix(i, COL_ִ�б��) = Val("" & rsTmp!ִ�б��)

            If Format(rsTmp!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &HFF&
            End If

            If mbytUseType <> 0 And mbytUseType <> 3 And mbytUseType <> 4 Then
                .TextMatrix(i, col_��ʼִ��ʱ��) = "" & rsTmp!��ʼʱ��
                .TextMatrix(i, col_��ֹʱ��) = "" & rsTmp!��ֹʱ��
            End If
            
            If mbytUseType = 0 Or mbytUseType = 3 Or mbytUseType = 4 Then
                .TextMatrix(i, Col_�����ĿID) = "" & rsTmp!�����ĿID
            End If
            rsTmp.MoveNext
        Next

        '�ٴ���һЩ�����е�����,��������ݵ���ʾ
        For i = 1 To .Rows - 1
            '��ҩ;��
            If .TextMatrix(i, COL�������) = "E" And Val(.TextMatrix(i, COL���ID)) = 0 _
               And Val(.TextMatrix(i - 1, COL���ID)) = Val(.TextMatrix(i, colID)) _
               And InStr(",5,6,", .TextMatrix(i - 1, COL�������)) > 0 Then
                .RowHidden(i) = True
                '��ʾ��ҩ;��
                For j = i - 1 To .FixedRows Step -1
                    If Val(.TextMatrix(j, COL���ID)) = Val(.TextMatrix(i, colID)) Then
                        .TextMatrix(j, col�÷�) = .TextMatrix(i, col����) & .TextMatrix(i, col����) '�����д��� ����

                        '��ʾ��ҩ��ִ������
                        If Val(.Cell(flexcpData, j, colִ������)) = 5 And Val(.Cell(flexcpData, i, colִ������)) <> 5 Then
                            .TextMatrix(j, colִ������) = IIf(Val(.TextMatrix(j, COL_ִ�б��)) = 2, "��ȡҩ", "�Ա�ҩ")
                        ElseIf Val(.Cell(flexcpData, j, colִ������)) <> 5 And Val(.Cell(flexcpData, i, colִ������)) = 5 Then
                            .TextMatrix(j, colִ������) = "��Ժ��ҩ"
                        Else
                            .TextMatrix(j, colִ������) = IIf(Val(.TextMatrix(j, COL_ִ�б��)) = 0, "����", "��ȡҩ")
                        End If
                    Else
                        Exit For
                    End If
                Next
            End If

            '��Ѫ;��
            If .TextMatrix(i, COL�������) = "E" And .TextMatrix(i - 1, COL�������) = "K" _
               And Val(.TextMatrix(i, COL���ID)) = Val(.TextMatrix(i - 1, colID)) Then
                .RowHidden(i) = True
                .TextMatrix(i - 1, col�÷�) = .TextMatrix(i, col����)
                .TextMatrix(i - 1, col����) = .TextMatrix(i - 1, col����) & "(" & .TextMatrix(i, col����) & ")"
            End If

            '��ҩ�䷽�ͼ������
            If .TextMatrix(i, COL�������) = "E" And Val(.TextMatrix(i, COL���ID)) = 0 _
               And Val(.TextMatrix(i - 1, COL���ID)) = Val(.TextMatrix(i, colID)) _
               And InStr(",7,E,C,", .TextMatrix(i - 1, COL�������)) > 0 Then

                str�巨 = "": str�걾 = "": strTmp = ""
                j = .FindRow(CStr(Val(.TextMatrix(i, colID))), , COL���ID)

                '��ҩ�������ִ�п���
                .TextMatrix(i, colִ�п���) = .TextMatrix(j, colִ�п���)

                '��ʾ��ҩ�䷽ִ������:��ҩƷΪ׼�ж�
                If .TextMatrix(i - 1, COL�������) <> "C" Then
                    If Val(.Cell(flexcpData, j, colִ������)) = 5 And Val(.Cell(flexcpData, i, colִ������)) <> 5 Then
                        .TextMatrix(i, colִ������) = IIf(Val(.TextMatrix(j, COL_ִ�б��)) = 2, "��ȡҩ", "�Ա�ҩ")
                    ElseIf Val(.Cell(flexcpData, j, colִ������)) <> 5 And Val(.Cell(flexcpData, i, colִ������)) = 5 Then
                        .TextMatrix(i, colִ������) = "��Ժ��ҩ"
                    Else
                        .TextMatrix(i, colִ������) = IIf(Val(.TextMatrix(j, COL_ִ�б��)) = 0, "����", "��ȡҩ")
                    End If
                End If
                
                'j--��ϼ�����Ŀ���к�
                For k = j To i - 1
                    .RowHidden(k) = k <> i
                    '��ϼ�����Ŀ����ʾ����Ŀ
                    If .TextMatrix(k, COL�������) = "C" And Val(.TextMatrix(k, Col_�����ĿID)) = 0 Then
                        If mbytUseType = 0 Or mbytUseType = 3 Or mbytUseType = 4 Then
                            strTmp = strTmp & "," & .TextMatrix(k, col����)          'ȡ������Ŀ������
                            str�걾 = .TextMatrix(j, col�걾��λ)    'ȡ��һ��������Ŀ�ı걾
                        End If
                    ElseIf .TextMatrix(k, COL�������) = "E" And Val(.TextMatrix(k, COL���ID)) <> 0 Then
                        str�巨 = .TextMatrix(k, col����)
                    End If
                Next

                If .TextMatrix(i - 1, COL�������) = "C" Then
                    If mbytUseType = 0 Or mbytUseType = 3 Or mbytUseType = 4 Then
                        .TextMatrix(i, col����) = Mid(strTmp, 2) & IIf(str�걾 <> "", "(" & str�걾 & ")", "")
                    End If
                Else
                    .TextMatrix(i, col����) = "��ҩ�䷽," & .TextMatrix(i, colƵ��) & "," & _
                                            str�巨 & "," & .TextMatrix(i, col����)
                    .TextMatrix(i, col����) = .TextMatrix(i, col����) & "��"
                End If
            End If

            '������
            If .TextMatrix(i, COL�������) = "D" And Val(.TextMatrix(i, COL���ID)) = 0 Then
                str�걾 = "": str�巨 = "": strTmp = ""
                For j = i + 1 To .Rows - 1
                    If Val(.TextMatrix(j, COL���ID)) = Val(.TextMatrix(i, colID)) Then
                        .RowHidden(j) = True
                        If .TextMatrix(j, col�걾��λ) <> "" _
                           And Val(.TextMatrix(j, col������ĿID)) = Val(.TextMatrix(i, col������ĿID)) Then    '��ͬ����ĿID�����·�ʽ
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
                If str�걾 <> "" Then    '��ǰ�ļ�鷽ʽʱ����ʾ��ϸҽ������
                    .TextMatrix(i, col����) = .TextMatrix(i, col����) & ":" & Mid(str�걾, 2)
                End If
            End If

            '������Ŀ
            If .TextMatrix(i, COL�������) = "F" And Val(.TextMatrix(i, COL���ID)) = 0 Then
                strTmp = "": str���� = ""
                For j = i + 1 To .Rows - 1
                    If Val(.TextMatrix(j, COL���ID)) = Val(.TextMatrix(i, colID)) Then
                        .RowHidden(j) = True
                        If .TextMatrix(j, COL�������) = "F" Then
                            strTmp = strTmp & "," & .TextMatrix(j, col����)
                        ElseIf .TextMatrix(j, COL�������) = "G" Then
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

Private Sub UserControl_Resize()
    On Error Resume Next
    vsAdvice.Top = UserControl.ScaleTop + 60
    vsAdvice.Left = UserControl.ScaleLeft
    vsAdvice.Height = UserControl.ScaleHeight - 100
    vsAdvice.Width = UserControl.ScaleWidth
End Sub

Private Sub vsAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'����һ��ҽ���е�ѡ��״̬

    If Col = colȱʡ Or Col = col��ѡ Then
        Dim i As Long, lng��ID As Long, lngThis��ID As Long
        Dim lngBegin As Long, lngEnd As Long
        
        With vsAdvice
            'һ����ҩ��һ�����ã���������ҳ���ʼ��
            If Not RowInһ����ҩ(Row, lngBegin, lngEnd, True) Then
                
                If Val(.TextMatrix(Row, COL���ID)) = 0 Then
                    lng��ID = Val(.TextMatrix(Row, colID))
                Else
                    lng��ID = Val(.TextMatrix(Row, COL���ID))
                End If
                
                lngBegin = Row
                lngEnd = Row
                For i = Row - 1 To .FixedRows Step -1
                    If Val(.TextMatrix(i, COL���ID)) = 0 Then
                        lngThis��ID = Val(.TextMatrix(i, colID))
                    Else
                        lngThis��ID = Val(.TextMatrix(i, COL���ID))
                    End If
                    If lngThis��ID <> lng��ID Then
                        Exit For
                    Else
                        lngBegin = i
                    End If
                Next
                
                For i = Row + 1 To .Rows - 1
                    If Val(.TextMatrix(i, COL���ID)) = 0 Then
                        lngThis��ID = Val(.TextMatrix(i, colID))
                    Else
                        lngThis��ID = Val(.TextMatrix(i, COL���ID))
                    End If
                    If lngThis��ID <> lng��ID Then
                        Exit For
                    Else
                        lngEnd = i
                    End If
                Next
            End If
            
            For i = lngBegin To lngEnd
                If i <> Row Then
                    .TextMatrix(i, Col) = .TextMatrix(Row, Col)
                End If
                If Col = col��ѡ And .TextMatrix(Row, Col) = -1 And mbytUseType = 0 Then
                    .TextMatrix(i, colȱʡ) = 0
                End If
                If Col = colȱʡ And .TextMatrix(Row, Col) = -1 And mbytUseType = 0 Then
                    .TextMatrix(i, col��ѡ) = 0
                End If
            Next
            
            RaiseEvent DataChange
        End With
    End If
End Sub


Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long, Optional ByVal blnIsHide As Boolean) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
'������blnIsHide=��Χ�Ƿ�������ص���
    Dim i As Long, blnTmp As Boolean
    
    With vsAdvice
        If .TextMatrix(lngRow, COL�������) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, COL�������)) = 0 Then Exit Function
        
        If Val(.TextMatrix(lngRow - 1, COL���ID)) = Val(.TextMatrix(lngRow, COL���ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, COL���ID)) = Val(.TextMatrix(lngRow, COL���ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL���ID)) = Val(.TextMatrix(lngRow, COL���ID)) And Val(.TextMatrix(lngRow, COL���ID)) <> 0 Or ((Val(.TextMatrix(lngRow, COL���ID)) = Val(.TextMatrix(i, colID)) Or Val(.TextMatrix(i, COL���ID)) = Val(.TextMatrix(lngRow, colID))) And blnIsHide) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL���ID)) = Val(.TextMatrix(lngRow, COL���ID)) And Val(.TextMatrix(lngRow, COL���ID)) <> 0 Or ((Val(.TextMatrix(lngRow, COL���ID)) = Val(.TextMatrix(i, colID)) Or Val(.TextMatrix(i, COL���ID)) = Val(.TextMatrix(lngRow, colID))) And blnIsHide) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInһ����ҩ = blnTmp
    End With
End Function

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
        lngW = UserControl.TextWidth(vsAdvice.TextMatrix(vsAdvice.FixedRows - 1, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
    End If
End Sub

Private Sub vsAdvice_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = mblnReadOnly Or (Col <> colȱʡ And Col <> col��ѡ)
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
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, OS.SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsAdvice_GotFocus()
    vsAdvice.BackColorSel = &HFFCC99
End Sub

Private Sub vsAdvice_LostFocus()
    vsAdvice.BackColorSel = &HFFEBD7
End Sub

Public Sub SetVsAdviceFontSize(ByVal lngFontSize As Long)
'���ܣ�����ҽ���嵥�����壬�������иߺ��п�
    
    Call Grid.SetFontSize(vsAdvice, lngFontSize, col����)
End Sub


