VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmRegistPlanPlan 
   BorderStyle     =   0  'None
   Caption         =   "�ƻ����źű�"
   ClientHeight    =   3210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid vsPlan 
      Height          =   2145
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3405
      _cx             =   6006
      _cy             =   3784
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
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483641
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   26
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmRegistPlanPlan.frx":0000
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.PictureBox picImgPlan 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   30
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   1
         Top             =   195
         Width           =   210
         Begin VB.Image imgColPlan 
            Height          =   195
            Left            =   0
            Picture         =   "frmRegistPlanPlan.frx":032F
            ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
   End
End
Attribute VB_Name = "frmRegistPlanPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mstrPrivs As String
Private mArrFilter As Variant
Private Sub InitVsGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2009-09-09 15:45:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intType As Integer, i As Integer, objGrid As VSFlexGrid
   i = 0
    With vsPlan
        .Redraw = flexRDNone
        .Rows = 3: .FixedRows = 2
        .FixedCols = 1
        .Cols = 40:   .Clear
        .FrozenCols = 2
        .TextMatrix(0, i) = "  ": .ColWidth(i) = 285
        .TextMatrix(1, i) = "  ":  .ColKey(i) = "��־": i = i + 1
        
        .TextMatrix(0, i) = "ID": .ColHidden(i) = True: .ColWidth(i) = 0
        .TextMatrix(1, i) = "ID": .ColKey(i) = "ID": i = i + 1
         
        .TextMatrix(0, i) = "����": .ColWidth(i) = 720
        .TextMatrix(1, i) = "����": .ColKey(i) = "����": i = i + 1

        .TextMatrix(0, i) = "�ű�": .ColWidth(i) = 480
        .TextMatrix(1, i) = "�ű�": .ColKey(i) = "�ű�": i = i + 1

        .TextMatrix(0, i) = "����": .ColWidth(i) = 1000
        .TextMatrix(1, i) = "����": .ColKey(i) = "����": i = i + 1
        .TextMatrix(0, i) = "��Ŀ": .ColWidth(i) = 1000
        .TextMatrix(1, i) = "��Ŀ": .ColKey(i) = "��Ŀ": i = i + 1
        .TextMatrix(0, i) = "ҽ��":: .ColWidth(i) = 1000
        .TextMatrix(1, i) = "ҽ��": .ColKey(i) = "ҽ��": i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 495
        .TextMatrix(1, i) = "����": .ColKey(i) = "����": i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 450
        .TextMatrix(1, i) = "����": .ColKey(i) = "����-����": i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 450
        .TextMatrix(1, i) = "�޺�": .ColKey(i) = "����-�޺�": i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 450
        .TextMatrix(1, i) = "��Լ": .ColKey(i) = "����-��Լ": i = i + 1

        .TextMatrix(0, i) = "��һ": .ColWidth(i) = 450
        .TextMatrix(1, i) = "����": .ColKey(i) = "��һ-����": i = i + 1
        .TextMatrix(0, i) = "��һ": .ColWidth(i) = 450
        .TextMatrix(1, i) = "�޺�": .ColKey(i) = "��һ-�޺�": i = i + 1
        .TextMatrix(0, i) = "��һ": .ColWidth(i) = 450
        .TextMatrix(1, i) = "��Լ": .ColKey(i) = "��һ-��Լ": i = i + 1

        .TextMatrix(0, i) = "�ܶ�": .ColWidth(i) = 450
        .TextMatrix(1, i) = "����": .ColKey(i) = "�ܶ�-����": i = i + 1
        .TextMatrix(0, i) = "�ܶ�": .ColWidth(i) = 450
        .TextMatrix(1, i) = "�޺�": .ColKey(i) = "�ܶ�-�޺�": i = i + 1
        .TextMatrix(0, i) = "�ܶ�": .ColWidth(i) = 450
        .TextMatrix(1, i) = "��Լ": .ColKey(i) = "�ܶ�-��Լ": i = i + 1

        .TextMatrix(0, i) = "����": .ColWidth(i) = 450
        .TextMatrix(1, i) = "����": .ColKey(i) = "����-����": i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 450
        .TextMatrix(1, i) = "�޺�": .ColKey(i) = "����-�޺�": i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 450
        .TextMatrix(1, i) = "��Լ": .ColKey(i) = "����-��Լ": i = i + 1

        .TextMatrix(0, i) = "����": .ColWidth(i) = 450
        .TextMatrix(1, i) = "����": .ColKey(i) = "����-����": i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 450
        .TextMatrix(1, i) = "�޺�": .ColKey(i) = "����-�޺�": i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 450
        .TextMatrix(1, i) = "��Լ": .ColKey(i) = "����-��Լ": i = i + 1

        .TextMatrix(0, i) = "����": .ColWidth(i) = 450
        .TextMatrix(1, i) = "����": .ColKey(i) = "����-����": i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 450
        .TextMatrix(1, i) = "�޺�": .ColKey(i) = "����-�޺�": i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 450
        .TextMatrix(1, i) = "��Լ": .ColKey(i) = "����-��Լ": i = i + 1

        .TextMatrix(0, i) = "����": .ColWidth(i) = 450
        .TextMatrix(1, i) = "����": .ColKey(i) = "����-����": i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 450
        .TextMatrix(1, i) = "�޺�": .ColKey(i) = "����-�޺�": i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 450
        .TextMatrix(1, i) = "��Լ": .ColKey(i) = "����-��Լ": i = i + 1
        .TextMatrix(0, i) = "���﷽ʽ": .ColWidth(i) = 855
        .TextMatrix(1, i) = "���﷽ʽ": .ColKey(i) = "���﷽ʽ": i = i + 1
        .TextMatrix(0, i) = "IDS": .ColWidth(i) = 0: .ColHidden(i) = True
        .TextMatrix(1, i) = "IDS": .ColKey(i) = "IDS": i = i + 1
        .TextMatrix(0, i) = "��Чʱ��": .ColWidth(i) = 2000
        .TextMatrix(1, i) = "��Чʱ��": .ColKey(i) = "��Чʱ��": i = i + 1
        .TextMatrix(0, i) = "ʧЧʱ��": .ColWidth(i) = 1200
        .TextMatrix(1, i) = "ʧЧʱ��": .ColKey(i) = "ʧЧʱ��": i = i + 1
        .TextMatrix(0, i) = "���" & vbCrLf & "����": .ColWidth(i) = 765
        .TextMatrix(1, i) = "���" & vbCrLf & "����": .ColKey(i) = "��ſ���": i = i + 1
        
        .TextMatrix(0, i) = "������": .ColWidth(i) = 1000
        .TextMatrix(1, i) = "������": .ColKey(i) = "������": i = i + 1
        .TextMatrix(0, i) = "����ʱ��": .ColWidth(i) = 1200
        .TextMatrix(1, i) = "����ʱ��": .ColKey(i) = "����ʱ��": i = i + 1
        
        .TextMatrix(0, i) = "�����": .ColWidth(i) = 1000
        .TextMatrix(1, i) = "�����": .ColKey(i) = "�����": i = i + 1
        .TextMatrix(0, i) = "���ʱ��": .ColWidth(i) = 1200
        .TextMatrix(1, i) = "���ʱ��": .ColKey(i) = "���ʱ��": i = i + 1
        .TextMatrix(0, i) = "ʵ��ִ��ʱ��": .ColWidth(i) = 1500
        .TextMatrix(1, i) = "ʵ��ִ��ʱ��": .ColKey(i) = "ʵ��ִ��ʱ��": i = i + 1
        .TextMatrix(0, i) = "Ӧ������": .ColWidth(i) = 2000
        .TextMatrix(1, i) = "Ӧ������": .ColKey(i) = "Ӧ������": i = i + 1
        .Cell(flexcpText, 0, 0, .Rows - 1) = " "
        .MergeCells = flexMergeFixedOnly
        .MergeCellsFixed = flexMergeFree
        For i = 0 To .Cols - 1
            .MergeCol(i) = True:
            .FixedAlignment(i) = flexAlignCenterCenter
            'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
            Select Case .ColKey(i)
            Case "ID", "��־", "IDS"
                 .ColData(i) = "-1|1"
            Case "����", "�ű�", "��Чʱ��"
                .ColData(i) = "1|0"
            End Select
        Next
         .MergeRow(0) = True: .MergeRow(1) = True
        .Redraw = flexRDBuffered
    End With
End Sub
 
Public Sub zlRefreshData(ByVal ArrFilter As Variant)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ˢ������
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-09-15 11:19:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mArrFilter = ArrFilter
    Call LoadDataToList
End Sub
Private Sub LoadDataToList()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ݸ�����
    '����:���˺�
    '����:2009-09-07 11:53:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFilter As String, rsTemp As New ADODB.Recordset, lngRow As Long, strSQL As String
    Dim blnHistory As Boolean, strStartDate As String, lngPriID As Long
    Dim strTable As String
    Dim strWhere As String
    
    Err = 0: On Error GoTo Errhand:
    If CStr(mArrFilter("��Ч��")(0)) <> "1901-01-01" Then
        strFilter = "  And Nvl(A.��ʼʱ��,To_Date('3000-01-01','YYYY-MM-DD'))>=[4]   And Nvl(A.��ֹʱ��,To_Date('1900-01-01','YYYY-MM-DD'))<=[5]"
    End If
    If Val(mArrFilter("����ID")) > 0 Then strFilter = strFilter & " And A.����ID=[1]"
    If Val(mArrFilter("����ID")) = -1 Then strFilter = strFilter & " And A.����ID in (Select ����ID From ������Ա where ��Աid=" & UserInfo.ID & ") "
    
    '105869�����ݼƻ���ҽ������
    Select Case mArrFilter("ҽ��ID")(1)
    Case "ID"
         strFilter = strFilter & "  And C.ҽ��ID=[2]"
    Case "UPR"
         strFilter = strFilter & " And Upper(C.ҽ������)=[3]"
    Case "NONE"
         strFilter = strFilter & " And C.ҽ������=[3]"
    End Select

 
    If CStr(mArrFilter("��Чʱ��")(0)) <> "1901-01-01" And CStr(mArrFilter("����ʱ��")(0)) <> "1901-01-01" And CStr(mArrFilter("���ʱ��")(0)) <> "1901-01-01" Then
        strFilter = strFilter & "  And (C.��Чʱ�� between [6] and [7] or C.����ʱ�� between [8] and [9] or C.���ʱ�� between [10] and [11]) "
    ElseIf CStr(mArrFilter("��Чʱ��")(0)) <> "1901-01-01" And CStr(mArrFilter("����ʱ��")(0)) <> "1901-01-01" And CStr(mArrFilter("���ʱ��")(0)) = "1901-01-01" Then
        strFilter = strFilter & "  And (C.��Чʱ�� between [6] and [7] or C.����ʱ�� between [8] and [9] ) "
    ElseIf CStr(mArrFilter("��Чʱ��")(0)) <> "1901-01-01" And CStr(mArrFilter("����ʱ��")(0)) = "1901-01-01" And CStr(mArrFilter("���ʱ��")(0)) <> "1901-01-01" Then
        strFilter = strFilter & "  And (C.��Чʱ�� between [6] and [7]   or C.���ʱ�� between [10] and [11]) "
    ElseIf CStr(mArrFilter("��Чʱ��")(0)) = "1901-01-01" And CStr(mArrFilter("����ʱ��")(0)) <> "1901-01-01" And CStr(mArrFilter("���ʱ��")(0)) <> "1901-01-01" Then
        strFilter = strFilter & "  And (C.����ʱ�� between [8] and [9]  or C.���ʱ�� between [10] and [11]) "
    ElseIf CStr(mArrFilter("��Чʱ��")(0)) <> "1901-01-01" Then
        strFilter = strFilter & "  And (C.����ʱ�� between [6] and [7])  "
    ElseIf CStr(mArrFilter("����ʱ��")(0)) <> "1901-01-01" Then
        strFilter = strFilter & "  And (C.����ʱ�� between [8] and [9])  "
    ElseIf CStr(mArrFilter("���ʱ��")(0)) <> "1901-01-01" Then
        strFilter = strFilter & "  And (C.���ʱ�� between [10] and [11])  "
    End If
    
    If Val(mArrFilter("����δ��Ч�ƻ�")) = 1 Then strFilter = strFilter & " and C.��Чʱ��>nvl(A.��ʼʱ��,to_date('1901-01-01','yyyy-mm-dd'))"
    If Val(mArrFilter("����ʾδ��ƻ�")) = 1 Then strFilter = strFilter & " and  C.���ʱ�� IS NULL "


    strTable = "" & _
    "   Select C.ID, " & _
    "             Sum(Decode(B.������Ŀ,'����',B.�޺���,0)) as �����޺�, Sum(Decode(B.������Ŀ,'����',B.��Լ��,0))  as ������Լ," & _
    "             Sum(Decode(B.������Ŀ,'��һ',B.�޺���,0)) as ��һ�޺�, Sum(Decode(B.������Ŀ,'��һ',B.��Լ��,0))  as ��һ��Լ," & _
    "             Sum(Decode(B.������Ŀ,'�ܶ�',B.�޺���,0)) as �ܶ��޺�, Sum(Decode(B.������Ŀ,'�ܶ�',B.��Լ��,0))  as �ܶ���Լ," & _
    "             Sum(Decode(B.������Ŀ,'����',B.�޺���,0)) as �����޺�, Sum(Decode(B.������Ŀ,'����',B.��Լ��,0))  as ������Լ," & _
    "             Sum(Decode(B.������Ŀ,'����',B.�޺���,0)) as �����޺�, Sum(Decode(B.������Ŀ,'����',B.��Լ��,0))  as ������Լ," & _
    "             Sum(Decode(B.������Ŀ,'����',B.�޺���,0)) as �����޺�, Sum(Decode(B.������Ŀ,'����',B.��Լ��,0))  as ������Լ," & _
    "             Sum(Decode(B.������Ŀ,'����',B.�޺���,0)) as �����޺�, Sum(Decode(B.������Ŀ,'����',B.��Լ��,0))  as ������Լ" & _
    "   From �ҺŰ��żƻ� C,�Һżƻ����� B,�ҺŰ��� A  " & _
    "   Where C.ID=B.�ƻ�ID(+)   and C.����ID=A.ID  " & strFilter & _
    "   Group by C.ID"
    
    '105869��ȡ�ƻ���ҽ�����շ���Ŀ
    strSQL = " " & _
        "   Select P.*,B.���� As ��Ŀ,D.���� As ���� " & _
        "   From ( " & _
        "     Select  row_number()  over (Partition By �ƻ�id Order By �ƻ�id,���� Desc) As ���1,M.* " & _
        "     From ( " & _
        "       Select Level As ����, Sys_Connect_By_Path(��������, ';') �������Ҽ�, Q.*  " & _
        "       From (  Select  C.Id as �ƻ�ID,C.����ID ,A.����,  A.����,  A.����id,  C.��Ŀid, C.ҽ������,  C.ҽ��id,     " & _
        "                              C.����,C1.�����޺�,C1.������Լ,C.��һ,C1.��һ�޺�,C1.��һ��Լ,C.�ܶ�,C1.�ܶ��޺�,C1.�ܶ���Լ, " & _
        "                              C.����,C1.�����޺�,C1.������Լ,C.����,C1.�����޺�,C1.������Լ,C.����,C1.�����޺�,C1.������Լ, " & _
        "                              C.����,C1.�����޺�,C1.������Լ, " & _
        "                              A.��������,   Decode(Nvl(C.���﷽ʽ,0),0,'������',1,'ָ������',2,'��̬����',3,'ƽ������') as ���﷽ʽ ,  C.��ſ���," & _
        "                              to_char(A.��ʼʱ��,'yyyy-mm-dd hh24:mi:ss') ��ʼʱ��,  to_char(A.��ֹʱ��,'yyyy-mm-dd hh24:mi:ss') ��ֹʱ��," & _
        "                              to_char(C.��Чʱ��,'yyyy-mm-dd hh24:mi:ss') as ��Чʱ��,to_char(C.ʧЧʱ��,'yyyy-mm-dd hh24:mi:ss') as ʧЧʱ��," & _
        "                              to_char(C.ʵ����Ч,'yyyy-mm-dd hh24:mi:ss') as ʵ��ִ��ʱ��,            " & _
        "                              C.������,to_char(C.����ʱ��,'yyyy-mm-dd hh24:mi:ss') as ����ʱ��,            " & _
        "                              C.�����,to_char(C.���ʱ��,'yyyy-mm-dd hh24:mi:ss') as ���ʱ�� , " & _
        "                              b.��������,row_number() over (Partition By �ƻ�ID Order By �ƻ�id,��������) As ��� " & _
        "           From  (" & strTable & ") C1,�ҺŰ��żƻ� C,�ҺŰ��� A,�Һżƻ����� B " & _
        "           Where C.ID=C1.ID And C.����ID =A.Id And C.Id=B.�ƻ�ID(+)   " & _
        "           Order By �ƻ�ID,�������� ) Q " & _
        "        Connect By �ƻ�id= Prior �ƻ�id And ���-1 =Prior ��� " & _
        "        )  M ) P,�շ���ĿĿ¼ B,���ű� D " & _
        "    Where P.���1=1 And P.��Ŀid=b.Id And P.����id =d.Id(+) And (B.վ��='" & gstrNodeNo & "' Or b.վ�� is Null)   " & _
        "    Order By ����, ��Чʱ�� Desc, �ƻ�ID DESC"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
        Val(mArrFilter("����ID")), _
        Val(mArrFilter("ҽ��ID")(0)), _
        CStr(mArrFilter("ҽ��ID")(0)), _
        CDate(mArrFilter("��Ч��")(0)), CDate(mArrFilter("��Ч��")(1)), _
        CDate(mArrFilter("��Чʱ��")(0)), CDate(mArrFilter("��Чʱ��")(1)), _
        CDate(mArrFilter("����ʱ��")(0)), CDate(mArrFilter("����ʱ��")(1)), _
        CDate(mArrFilter("���ʱ��")(0)), CDate(mArrFilter("���ʱ��")(1)), _
        "")
      
    With Me.vsPlan
        If .Row > 0 And .Row <= .Rows - 1 Then lngPriID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        .Clear 1
        .Rows = 3: lngRow = 2
        .Redraw = flexRDNone
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 2
        If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ID")) = Nvl(rsTemp!�ƻ�Id)
            .Cell(flexcpData, lngRow, .ColIndex("ID")) = Nvl(rsTemp!����ID)
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("�ű�")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("��Ŀ")) = Nvl(rsTemp!��Ŀ)
            .TextMatrix(lngRow, .ColIndex("ҽ��")) = Nvl(rsTemp!ҽ������)
            .TextMatrix(lngRow, .ColIndex("����-����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("����-�޺�")) = Format(Val(Nvl(rsTemp!�����޺�)), "###;;")
            .TextMatrix(lngRow, .ColIndex("����-��Լ")) = Format(Val(Nvl(rsTemp!������Լ)), "###;;")
            .TextMatrix(lngRow, .ColIndex("��һ-����")) = Nvl(rsTemp!��һ)
            .TextMatrix(lngRow, .ColIndex("��һ-�޺�")) = Format(Val(Nvl(rsTemp!��һ�޺�)), "###;;")
            .TextMatrix(lngRow, .ColIndex("��һ-��Լ")) = Format(Val(Nvl(rsTemp!��һ��Լ)), "###;;")
            .TextMatrix(lngRow, .ColIndex("�ܶ�-����")) = Nvl(rsTemp!�ܶ�)
            .TextMatrix(lngRow, .ColIndex("�ܶ�-�޺�")) = Format(Val(Nvl(rsTemp!�ܶ��޺�)), "###;;")
            .TextMatrix(lngRow, .ColIndex("�ܶ�-��Լ")) = Format(Val(Nvl(rsTemp!�ܶ���Լ)), "###;;")
            .TextMatrix(lngRow, .ColIndex("����-����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("����-�޺�")) = Format(Val(Nvl(rsTemp!�����޺�)), "###;;")
            .TextMatrix(lngRow, .ColIndex("����-��Լ")) = Format(Val(Nvl(rsTemp!������Լ)), "###;;")
            .TextMatrix(lngRow, .ColIndex("����-����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("����-�޺�")) = Format(Val(Nvl(rsTemp!�����޺�)), "###;;")
            .TextMatrix(lngRow, .ColIndex("����-��Լ")) = Format(Val(Nvl(rsTemp!������Լ)), "###;;")
            .TextMatrix(lngRow, .ColIndex("����-����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("����-�޺�")) = Format(Val(Nvl(rsTemp!�����޺�)), "###;;")
            .TextMatrix(lngRow, .ColIndex("����-��Լ")) = Format(Val(Nvl(rsTemp!������Լ)), "###;;")
            .TextMatrix(lngRow, .ColIndex("����-����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("����-�޺�")) = Format(Val(Nvl(rsTemp!�����޺�)), "###;;")
            .TextMatrix(lngRow, .ColIndex("����-��Լ")) = Format(Val(Nvl(rsTemp!������Լ)), "###;;")
            .TextMatrix(lngRow, .ColIndex("����")) = IIf(Val(Nvl(rsTemp!��������)) = 0, "", "��")
            .TextMatrix(lngRow, .ColIndex("���﷽ʽ")) = Nvl(rsTemp!���﷽ʽ)
            .TextMatrix(lngRow, .ColIndex("IDS")) = Nvl(rsTemp!����ID) & "_" & Nvl(rsTemp!��ĿID) & "_" & Nvl(rsTemp!ҽ��ID)
            If Nvl(rsTemp!�������Ҽ�) <> "" Then
                .TextMatrix(lngRow, .ColIndex("Ӧ������")) = Mid(Nvl(rsTemp!�������Ҽ�), 2)  ' Read�ƻ�Ӧ������(lng����ID, Val(Nvl(rsTemp!�ƻ�ID)), False) ' Nvl(rsTemp!��������)
            End If
            
            If Not IsNull(rsTemp!��Чʱ��) Then
                .TextMatrix(lngRow, .ColIndex("��Чʱ��")) = Format(rsTemp!��Чʱ��, "yyyy-MM-dd HH:mm:ss")
                If Format(Nvl(rsTemp!��Чʱ��), "yyyy-MM-dd HH:mm:ss") <= Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") And Nvl(rsTemp!���ʱ��) <> "" Then
                    '�Ѿ���Ч,���ܸ���
                    .Cell(flexcpData, lngRow, .ColIndex("��Чʱ��")) = 1
                Else
                    'δ��Ч,�ܸ���
                    .Cell(flexcpData, lngRow, .ColIndex("��Чʱ��")) = 0
                End If
            End If
            .TextMatrix(lngRow, .ColIndex("ʧЧʱ��")) = Nvl(rsTemp!ʧЧʱ��)
            .TextMatrix(lngRow, .ColIndex("��ſ���")) = IIf(Val(Nvl(rsTemp!��ſ���)) = 0, "", "��")
            
            .TextMatrix(lngRow, .ColIndex("������")) = Nvl(rsTemp!������)
            .TextMatrix(lngRow, .ColIndex("����ʱ��")) = Nvl(rsTemp!����ʱ��)
            .TextMatrix(lngRow, .ColIndex("�����")) = Nvl(rsTemp!�����)
            .TextMatrix(lngRow, .ColIndex("���ʱ��")) = Nvl(rsTemp!���ʱ��)
            If Nvl(rsTemp!ʵ��ִ��ʱ��) < "3000-01-01" Then
                .TextMatrix(lngRow, .ColIndex("ʵ��ִ��ʱ��")) = Nvl(rsTemp!ʵ��ִ��ʱ��)
            End If
            If Val(.Cell(flexcpData, lngRow, .ColIndex("��Чʱ��"))) = 1 Then
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = &H80000010
            Else
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = .ForeColor
            End If
            lngRow = lngRow + 1
           rsTemp.MoveNext
        Loop
       ' .AutoSizeMode = flexAutoSizeColWidth
        '.AutoSize 0, .Cols - 1
        If lngPriID <> 0 Then
            lngRow = .FindRow(lngPriID, 0, .ColIndex("ID"), , True)
            If lngRow > 0 Then .Row = lngRow
        Else
            .Row = 1
        End If
        If .RowIsVisible(.Row) = False Then .TopRow = .Row
        '�ָ�������
        zl_vsGrid_Para_Restore mlngModule, vsPlan, Me.Caption, "�ƻ���Ϣ�б�", True
        .ColWidth(.ColIndex("��־")) = 285
        .Redraw = flexRDBuffered
    End With
   Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
     Me.vsPlan.Redraw = flexRDBuffered
End Sub

Private Sub Form_Load()
    mlngModule = glngModul: mstrPrivs = gstrPrivs
    Call InitVsGrid
    Call vsPlan_LostFocus
    vsPlan_GotFocus
End Sub
Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With vsPlan
        .Left = ScaleLeft
        .Top = ScaleTop
        .Width = ScaleWidth
        .Height = ScaleHeight
    End With
End Sub

Public Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���д�ӡ,Ԥ���������EXCEL
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '����:���˺�
    '����:2009-09-09 11:24:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim rsTemp As New ADODB.Recordset, strSQL As String
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstrUnitName & "�Һżƻ���"
    

    If CStr(mArrFilter("��Ч��")(0)) <> "1901-01-01" Then
        objRow.Add "Ч�ڷ�Χ��" & CStr(mArrFilter("��Ч��")(0)) & "��" & CStr(mArrFilter("��Ч��")(1))
    End If
    If Val(mArrFilter("����ID")) > 0 Then
        strSQL = "Select ���� From ���ű� where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mArrFilter("����ID")))
        If rsTemp.EOF Then
            objRow.Add "���ң����п���"
        Else
            objRow.Add "���ң�" & Nvl(rsTemp!����)
        End If
    ElseIf Val(mArrFilter("����ID")) = -1 Then
        objRow.Add "���ң�����Ա��������"
    Else
        objRow.Add "���ң����п���"
    End If
    Select Case mArrFilter("ҽ��ID")(1)
    Case "ID"
        strSQL = "Select ���� From ��Ա�� where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mArrFilter("ҽ��ID")(0)))
        If rsTemp.EOF Then
            objRow.Add "ҽ��������"
        Else
            objRow.Add "ҽ����" & Nvl(rsTemp!����)
        End If
    Case "UPR", "NONE"
            objRow.Add "ҽ����" & CStr(mArrFilter("ҽ��ID")(0))
    End Select
    objPrint.UnderAppRows.Add objRow
    
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    Err = 0: On Error GoTo Errhand:
    With vsPlan
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .Cell(flexcpData, 0, intCol) = .ColWidth(intCol)
            If .ColHidden(intCol) Or intCol = .ColIndex("��־") Then .ColWidth(intCol) = 0
        Next
    End With
    
    Set objPrint.Body = vsPlan
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
    
    With vsPlan
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .ColWidth(intCol) = Val(.Cell(flexcpData, 0, intCol))
        Next
        .Redraw = flexRDBuffered
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
    With vsPlan
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .ColWidth(intCol) = Val(.Cell(flexcpData, 0, intCol))
        Next
        .Redraw = flexRDBuffered
    End With
End Sub
Public Sub zlCallCustomReprot(ByVal frmMain As Form, ByVal lngSys As Long, strReprotName As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ص��Զ��屨��
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-09-15 11:10:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, str���� As String
    '����ID_��ĿID_ҽ��ID
    With vsPlan
        varData = Split(.TextMatrix(.Row, .ColIndex("IDS")) & "___", "_")
        str���� = Trim(.TextMatrix(.Row, .ColIndex("����")))
        If str���� <> "" Then
            Call ReportOpen(gcnOracle, lngSys, strReprotName, frmMain, _
                "����=" & str����, "�ű�=" & Trim(.TextMatrix(.Row, .ColIndex("�ű�"))), _
                "����=" & Val(varData(0)), _
                "��Ŀ=" & Val(varData(1)), _
                "ҽ��=" & Val(varData(2)))
        Else
            Call ReportOpen(gcnOracle, lngSys, strReprotName, frmMain)
        End If
    End With
End Sub
Public Property Get zlGet����ID(Optional blnPlanID As Boolean = True) As Long
    With vsPlan
        If blnPlanID Then
            zlGet����ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        Else
            zlGet����ID = Val(.Cell(flexcpData, .Row, .ColIndex("ID")))
        End If
    End With
End Property

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsPlan, Me.Caption, "�ƻ���Ϣ�б�", True, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub vsPlan_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsPlan, Me.Caption, "�ƻ���Ϣ�б�", True, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub vsPlan_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsPlan
        If Col = .ColIndex("��־") Then Cancel = True
    End With
End Sub
Private Sub vsPlan_GotFocus()
    vsPlan.BackColorSel = &H8000000D
End Sub

Private Sub vsPlan_LostFocus()
    vsPlan.BackColorSel = GRD_LOSTFOCUS_COLORSEL
End Sub
Private Sub vsPlan_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsPlan, Me.Caption, "�ƻ���Ϣ�б�", True, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub imgColPlan_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgPlan.Hwnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlan.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsPlan, lngLeft, lngTop, imgColPlan.Height)
    zl_vsGrid_Para_Save mlngModule, vsPlan, Me.Caption, "�ƻ���Ϣ�б�", True, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub picImgPlan_Click()
    Call imgColPlan_Click
End Sub
Public Property Get zlPlanStatus() As Long
    Dim lngID As Long
    '��ȡ�ƻ����ŵĵ�ǰ״̬
    '0-�����ڼƻ�����,1-δ���,2-�Ѿ����,3-�Ѿ���Ч
    With vsPlan
        lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        If lngID = 0 Then zlPlanStatus = 0: Exit Property
        If .TextMatrix(.Row, .ColIndex("���ʱ��")) <> "" Then
            zlPlanStatus = 2
            If Val(.Cell(flexcpData, .Row, .ColIndex("��Чʱ��"))) = 1 Then
                zlPlanStatus = 3
            End If
        Else
              zlPlanStatus = 1
        End If
    End With
End Property

Public Sub zlActtion()
    zlControl.ControlSetFocus vsPlan, True
End Sub

