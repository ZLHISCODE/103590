VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmFeeDetail 
   BorderStyle     =   0  'None
   Caption         =   "������ϸ�б�"
   ClientHeight    =   6780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
      Height          =   4395
      Left            =   1425
      TabIndex        =   0
      Top             =   990
      Width           =   5850
      _cx             =   10319
      _cy             =   7752
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
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
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
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
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFeeDetail.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   4
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
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
   End
End
Attribute VB_Name = "frmFeeDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngMode As Long, mstrPrivs As String, mblnOriginal As Boolean
Private mbytType As Byte, mstrBalanceID As String

Public Sub ShowMe(ByVal objFont As Object, _
      ByVal lngModule As Long, ByVal strPrivs As String, _
      ByVal bytType As Byte, ByVal strBalanceID As String)
    '-------------------------------------------------------------------------------------------------
    '����:�������,��ʾ���ݵ���ϸ����
    '���:objFont-����������
    '       lngModule-ģ���
    '       strPrivs-Ȩ�޴�
    '����   bytType:1-ȫ��ʾ;2-ֻ��ʾδ�˲���;
    '       strBalanceID -����ID
    '����:������
    '����:2014-06-13
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    Set vsfDetail.Font = objFont
    Call zlRefresh(bytType, strBalanceID)
End Sub

Public Sub zlRefresh(ByVal bytType As Byte, ByVal strBalanceID As String, Optional blnOriginal As Boolean = True)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ˢ������
    '����:������
    '����:2014-06-13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytType = bytType
    mstrBalanceID = strBalanceID
    mblnOriginal = blnOriginal
    Call ReadListData(bytType)
End Sub

Private Sub SetJZDetail()
    Dim strHead As String
    Dim i As Long
    Dim varData As Variant
    
    strHead = "���,4,800|����,1,2000|���,1,1200|����,7,800|��λ,4,800|����,7,1000|Ӧ�ս��,7,1000|ʵ�ս��,7,1000|ִ�п���,4,1000|����,4,1000|˵��,1,1800|��¼״̬,1,0"
    
    With vsfDetail
        .HighLight = flexHighlightWithFocus
        .Redraw = False
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            If Split(varData(i), ",")(0) = "ID" Then .ColHidden(i) = True
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .colAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        
'        .Row = 1: .Col = 0: .ColSel = .Cols - 1
        .Redraw = True

        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("��¼״̬"))) = 1 Then .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlack
            If Val(.TextMatrix(i, .ColIndex("��¼״̬"))) = 2 Then .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
            If Val(.TextMatrix(i, .ColIndex("��¼״̬"))) = 3 Then .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlue
        Next i
    End With
    
End Sub

Private Function CheckBalance(lngBalanceID As Long) As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    strSQL = "Select 1 From ����Ԥ����¼ Where �������= [1] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngBalanceID)
    CheckBalance = rsTemp.EOF
End Function

Private Function ReadListData(ByVal bytType As Byte) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ص���ϸ����
    '����:���ݻ�ȡ�ɹ�����true,���򷵻�False
    '����:������
    '����:2014-06-13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsMain As ADODB.Recordset, rsSub As ADODB.Recordset
    Dim strWithTable As String, strWhere As String, i As Long, strҩ����λ As String, strҩ����װ As String
    Dim strTable As String, lngMainRow As Long, blnDel As Boolean, blnҩ����λ As Boolean
    On Error GoTo errHandle
    blnDel = False
    If bytType = 1 Then
        '�շѵ�
        If CheckBalance(Val(mstrBalanceID)) = True Then
        '10.29��ǰ���ݵĻ�ȡ
            strSQL = _
                " Select NO As ���ݺ�, ���, ��������, ������, �ѱ�, ���, ����, ��Ʒ��, ���, " & _
                "       Sum(����) As ����, ����, Sum(Ӧ�ս��) As Ӧ�ս��, Sum(ʵ�ս��) As ʵ�ս��, ִ�п���, ����, ˵��, ��¼״̬" & vbNewLine & _
                " From (Select a.����ID,D1.���� as ��������,A.������,a.No,C.���� as ���,Nvl(E.����,B.����) as ����,E1.���� as ��Ʒ��,B.���," & _
                        IIf(blnҩ����λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X." & strҩ����λ & ")", "A.���㵥λ") & " as ��λ," & _
                "       To_Char(Avg(Nvl(A.����,1)*" & IIf(blnDel, "-1*", "") & "A.����)" & _
                        IIf(blnҩ����λ, "/Nvl(X." & strҩ����װ & ",1)", "") & ",'9999990.00000') as ����, " & _
                "       a.�ѱ�,To_Char(Sum(A.��׼����)" & _
                        IIf(blnҩ����λ, "*Nvl(X." & strҩ����װ & ",1)", "") & ",'999999" & gstrFeePrecisionFmt & "') as ����, " & _
                "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.Ӧ�ս��),'9999999" & gstrDec & "') as Ӧ�ս��, " & _
                "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.ʵ�ս��),'9999999" & gstrDec & "') as ʵ�ս��, " & _
                "       D.���� as ִ�п���,Nvl(A.��������,B.��������) as ����," & _
                "       Decode(Nvl(A.ִ��״̬,0),0,'δִ��',1,'��ȫִ��',2,'����ִ��',9,'�쳣�շѵ�','��'||ABS(A.ִ��״̬)||'���˷�') as ˵��," & _
                "       A.��¼״̬, Nvl(a.�۸񸸺�, a.���) As ���" & _
                " From  ������ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,���ű� D1,�շ���Ŀ���� E,�շ���Ŀ���� E1,ҩƷ��� X" & _
                " Where A.�շ�ϸĿID=B.ID and A.�շ����=C.���� And A.ִ�в���ID=D.ID(+) And A.�շ�ϸĿID=X.ҩƷID(+)" & _
                "       And A.��¼����=1 And A.����ID = [1] And A.��¼״̬" & IIf(blnDel, "=2", " IN(1,3)") & _
                "       And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And a.��������ID=D1.ID(+) And E.����(+)=1 And E.����(+)=" & IIf(1 = 1, 3, 1) & _
                "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3 " & _
                " Group by a.����id, D1.����, a.������, a.�ѱ�,a.No,Nvl(A.�۸񸸺�,A.���),C.����,Nvl(E.����,B.����),E1.���� , B.���,A.���㵥λ,D.����," & _
                "       Nvl(A.��������,B.��������),A.ִ��״̬,A.��¼״̬,X.ҩƷID )" & _
                " Group By NO, ���, ��������, ������, �ѱ�, ���, ����, ��Ʒ��, ���, ����, ִ�п���, ����, ˵��, ��¼״̬" & _
                " Order By ���ݺ�, ���"
        Else
            strSQL = _
                " Select NO As ���ݺ�, ���, ��������, ������, �ѱ�, ���, ����, ��Ʒ��, ���, " & _
                "       Sum(����) As ����, ����, Sum(Ӧ�ս��) As Ӧ�ս��, Sum(ʵ�ս��) As ʵ�ս��, ִ�п���, ����, ˵��, ��¼״̬" & vbNewLine & _
                " From (Select a.����ID,D1.���� as ��������,A.������,a.No,C.���� as ���,Nvl(E.����,B.����) as ����,E1.���� as ��Ʒ��,B.���," & _
                        IIf(blnҩ����λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X." & strҩ����λ & ")", "A.���㵥λ") & " as ��λ," & _
                "       To_Char(Avg(Nvl(A.����,1)*" & IIf(blnDel, "-1*", "") & "A.����)" & _
                        IIf(blnҩ����λ, "/Nvl(X." & strҩ����װ & ",1)", "") & ",'9999990.00000') as ����, " & _
                "       a.�ѱ�,To_Char(Sum(A.��׼����)" & _
                        IIf(blnҩ����λ, "*Nvl(X." & strҩ����װ & ",1)", "") & ",'999999" & gstrFeePrecisionFmt & "') as ����, " & _
                "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.Ӧ�ս��),'9999999" & gstrDec & "') as Ӧ�ս��, " & _
                "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.ʵ�ս��),'9999999" & gstrDec & "') as ʵ�ս��, " & _
                "       D.���� as ִ�п���,Nvl(A.��������,B.��������) as ����," & _
                "       Decode(Nvl(A.ִ��״̬,0),0,'δִ��',1,'��ȫִ��',2,'����ִ��',9,'�쳣�շѵ�','��'||ABS(A.ִ��״̬)||'���˷�') as ˵��," & _
                "       A.��¼״̬, Nvl(a.�۸񸸺�, a.���) As ���" & _
                " From  ������ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,���ű� D1,�շ���Ŀ���� E,�շ���Ŀ���� E1,ҩƷ��� X," & _
                "       (Select Distinct ����ID From ����Ԥ����¼ Where �������= [1]) F" & _
                " Where A.�շ�ϸĿID=B.ID and A.�շ����=C.���� And A.ִ�в���ID=D.ID(+) And A.�շ�ϸĿID=X.ҩƷID(+)" & _
                "       And A.��¼����=1 And A.����ID = F.����ID And A.��¼״̬" & IIf(blnDel, "=2", " IN(1,3)") & _
                "       And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(1 = 1, 3, 1) & _
                "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And A.��������ID=D1.ID(+) And E1.����(+)=1 And E1.����(+)=3" & _
                " Group by a.����id, D1.����, a.������, a.�ѱ�,a.No,Nvl(A.�۸񸸺�,A.���),C.����,Nvl(E.����,B.����),E1.���� , B.���,A.���㵥λ,D.����," & _
                "       Nvl(A.��������,B.��������),A.ִ��״̬,A.��¼״̬,X.ҩƷID )" & _
                " Group By NO, ���, ��������, ������, �ѱ�, ���, ����, ��Ʒ��, ���, ����, ִ�п���, ����, ˵��, ��¼״̬" & _
                " Order By ���ݺ�, ���"
        End If
        
        Set rsMain = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrBalanceID)
        Set vsfDetail.DataSource = rsMain
        Call SetDetail
    Else
        '���˵�
        strSQL = "" & _
            " Select C.���� As ���, Nvl(E.����, B.����) As ����, B.���, Avg(Nvl(A.����, 1) * A.����) As ����, A.���㵥λ As ��λ," & vbNewLine & _
            "        Sum(A.��׼����) As ����, LTrim(To_Char(Sum(A.Ӧ�ս��), '99999" & gstrDec & "')) As Ӧ�ս��," & vbNewLine & _
            "        LTrim(To_Char(Sum(A.ʵ�ս��), '99999" & gstrDec & "')) As ʵ�ս��, D.���� As ִ�п���,Nvl(A.��������,B.��������) As ����," & vbNewLine & _
            "        Decode(Nvl(A.ִ��״̬,0),0,'δִ��',1,'��ȫִ��',2,'����ִ��',9,'�쳣�շѵ�','��'||ABS(A.ִ��״̬)||'���˷�') as ˵��,A.��¼״̬" & _
            " From ������ü�¼ A, �շ���ĿĿ¼ B, �շ���Ŀ��� C, ���ű� D, �շ���Ŀ���� E" & vbNewLine & _
            " Where A.�շ�ϸĿid = B.ID And A.�շ���� = C.���� And A.ִ�в���id = D.ID(+) And A.NO = [1] And A.��¼���� = 2 And" & vbNewLine & _
            "       A.��¼״̬ In (1,3) And A.�շ�ϸĿid = E.�շ�ϸĿid(+) And E.����(+) = 1 And E.����(+) = 3" & vbNewLine & _
            " Group By Nvl(A.�۸񸸺�, A.���), A.��׼����, C.����, Nvl(E.����, B.����), B.���, A.���㵥λ, D.����, Nvl(A.��������,B.��������), A.ִ��״̬, A.��¼״̬" & vbNewLine & _
            " Order By Nvl(A.�۸񸸺�, A.���)"
        Set rsMain = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrBalanceID)
        Set vsfDetail.DataSource = rsMain
        Call SetJZDetail
    End If
    ReadListData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub DetailSplitGroup()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Է����б���Ϣ���з�����ʾ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim strTemp As String

    On Error GoTo errHandle
    With vsfDetail
        For i = 0 To .Cols - 1
            If i < .ColIndex("���") And i > .ColIndex("˵��") Then
                .ColHidden(i) = True
            End If
        Next
        
        .OutlineBar = flexOutlineBarComplete
        .Subtotal flexSTClear
        .MultiTotals = True
        '&H8000000F
        .Subtotal flexSTSum, .ColIndex("���ݺ�"), .ColIndex("ʵ�ս��"), , &H8000000F, , True, "%s", , True
        .Subtotal flexSTSum, .ColIndex("���ݺ�"), .ColIndex("Ӧ�ս��"), , &H8000000F, , True, "%s", , True
        .SubtotalPosition = flexSTAbove

        .Outline .ColIndex("���")
        .OutlineCol = .ColIndex("���")

        For i = 1 To .Rows - 1
            .MergeRow(i) = False
            If .IsSubtotal(i) Then
                .IsCollapsed(i) = flexOutlineExpanded
                strTemp = .Cell(flexcpTextDisplay, i, 0)
                .RowHeight(i) = 350

                .Cell(flexcpText, i, .ColIndex("���")) = strTemp

                 strTemp = .Cell(flexcpTextDisplay, i + 1, .ColIndex("���ݺ�"))
                 strTemp = strTemp & Space(2) & "�ѱ�:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("�ѱ�"))
                 strTemp = strTemp & Space(2) & "��������:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("��������"))
                 strTemp = strTemp & Space(2) & "������:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("������"))
                 .MergeRow(i) = True
                 .MergeCells = flexMergeRestrictRows
                 .Cell(flexcpAlignment, i, .ColIndex("���"), i, .ColIndex("���")) = 1
'                 If Val(.TextMatrix(i + 1, .ColIndex("��¼״̬"))) = 1 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlack
'                 If Val(.TextMatrix(i + 1, .ColIndex("��¼״̬"))) = 2 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbRed
'                 If Val(.TextMatrix(i + 1, .ColIndex("��¼״̬"))) = 3 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlue
                 
                 For j = 0 To .Cols - 1
                    If j < .ColIndex("Ӧ�ս��") Then
                        If j >= .ColIndex("���") Then
                            .Cell(flexcpText, i, j) = strTemp
                            .Cell(flexcpFontBold, i, j) = False
                        End If
                    ElseIf .ColIndex("ʵ�ս��") = j Then
                        .TextMatrix(i, j) = Format(Val(.TextMatrix(i, j)), gstrDec)
                        .Cell(flexcpFontBold, i, j) = False
                    ElseIf .ColIndex("Ӧ�ս��") = j Then
                        .TextMatrix(i, j) = " " & Format(Val(.TextMatrix(i, j)), gstrDec)
                        .Cell(flexcpFontBold, i, j) = False
                    End If
                 Next
            Else
                .TextMatrix(i, .ColIndex("����")) = Format(Val(.TextMatrix(i, .ColIndex("����"))), gstrDec)
                .TextMatrix(i, .ColIndex("Ӧ�ս��")) = Format(Val(.TextMatrix(i, .ColIndex("Ӧ�ս��"))), gstrDec)
                .TextMatrix(i, .ColIndex("ʵ�ս��")) = Format(Val(.TextMatrix(i, .ColIndex("ʵ�ս��"))), gstrDec)
            End If
        Next
        Call .AutoSize(.ColIndex("���"))
        Call .AutoSize(.ColIndex("����"))
        
        For j = 0 To .Cols - 1
            If j < .ColIndex("Ӧ�ս��") Then
                .MergeCol(j) = True
            Else
                .MergeCol(j) = False
            End If
        Next
        
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Long, intShowName As Integer
    Dim varData As Variant

    strHead = "���ݺ�,1,0|���,1,0|��������,1,0|������,1,0|�ѱ�,1,0|���,4,800|����,1,2000|��Ʒ��,1,2000|���,1,1200|����,7,800|����,7,1000|Ӧ�ս��,7,1000|ʵ�ս��,7,1000|ִ�п���,4,1000|����,4,1000|˵��,1,1800|��¼״̬,1,0"
    
    With vsfDetail
        .HighLight = flexHighlightWithFocus
        .Redraw = False
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            If Split(varData(i), ",")(0) = "ID" Then .ColHidden(i) = True
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .colAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        
        'Call RestoreFlexState(mshDetail, App.ProductName & "\" & Me.Name)
        
        '.Row = 0: .Col = 0: .ColSel = .Cols - 1
        .Redraw = True
        If .Rows > 1 Then
            If .TextMatrix(1, .ColIndex("���ݺ�")) <> "" Then Call DetailSplitGroup
        End If
        For i = 1 To .Rows - 1
            If .IsSubtotal(i) = False Then
                If Val(.TextMatrix(i, .ColIndex("��¼״̬"))) = 1 Then .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlack
                If Val(.TextMatrix(i, .ColIndex("��¼״̬"))) = 2 Then .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                If Val(.TextMatrix(i, .ColIndex("��¼״̬"))) = 3 Then .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlue
                .RowHeight(i) = 300
            End If
        Next i
        
        intShowName = Val(zlDatabase.GetPara("ҩƷ������ʾ"))
        If intShowName <> 2 Then
            .ColHidden(.ColIndex("��Ʒ��")) = True
        Else
            .ColHidden(.ColIndex("��Ʒ��")) = False
        End If
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With vsfDetail
        .Top = 0
        .Left = 0
        .Height = Me.Height
        .width = Me.width
    End With
End Sub
