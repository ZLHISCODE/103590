VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm���ľܷ����嵥 
   BorderStyle     =   0  'None
   Caption         =   "���ľܷ����嵥"
   ClientHeight    =   4965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   4125
      Left            =   450
      TabIndex        =   0
      Top             =   210
      Width           =   7320
      _cx             =   12912
      _cy             =   7276
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
      Cols            =   20
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frm���ľܷ����嵥.frx":0000
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
Attribute VB_Name = "frm���ľܷ����嵥"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsNotPayStuff As ADODB.Recordset
Private mintUnit As Integer
Private mstrPrivs As String
Private mlngModule As Long
Private mArrFilter As Variant   '��������
Private mfrmMain As Form        '������
Private mrs�ܷ� As ADODB.Recordset
Private mlngCount�ָ�  As Long
Private Const mstrAllType As String = "�ٴ�,����,���,����,����,����,Ӫ��"
Private mbln������ʱ����� As Boolean

'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString
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
        .ColData(.ColIndex("״̬")) = 1
        .ColData(.ColIndex("��������")) = 1
        .ColData(.ColIndex("���ݺ�")) = 1
        .ColData(.ColIndex("��������")) = 1
        .ColData(.ColIndex("����")) = 1
    End With
End Sub

Public Function zlRestorePayStuff() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:�ָ��Ѿ��ܷ�����������
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-24 11:40:50
    '-----------------------------------------------------------------------------------------------------------
    If ISValied = False Then Exit Function
    zlRestorePayStuff = SaveData()
End Function
Private Function ISValied() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���񷢵��������
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-24 11:41:49
    '-----------------------------------------------------------------------------------------------------------
    Dim blnHaveData As Boolean
    Dim lngRow As Long
    With vsGrid
        blnHaveData = False
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, .ColIndex("״̬")) = "�ָ�" Then
                blnHaveData = True: Exit For
            End If
        Next
        If blnHaveData = False Then
            ShowMsgBox "û��ѡ����Ҫ�ָ��ľܷ����ϣ��������ֹ!"
            Exit Function
        End If
    End With
    ISValied = True
End Function

Private Function SaveData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���������Ϸ��ŵľܷ����ֽ��лָ�����
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-24 11:44:19
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, cllProc As Collection
    Dim int���� As Integer
    
    Set cllProc = New Collection
    With vsGrid
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, .ColIndex("״̬")) = "�ָ�" And .RowData(lngRow) <> 0 Then
                If Val(.TextMatrix(lngRow, .ColIndex("��¼����"))) = 1 Or (Val(.TextMatrix(lngRow, .ColIndex("��¼����"))) = 2 And (Val(.TextMatrix(lngRow, .ColIndex("�����־")))) = 1 Or (Val(.TextMatrix(lngRow, .ColIndex("�����־")))) = 4) Then
                    int���� = 1
                Else
                    int���� = 2
                End If
            
                'Zl_�������Ϸ���_�ܷ��ָ�(Id_In In ҩƷ�շ���¼.ID%Type)
                gstrSQL = "Zl_�������Ϸ���_�ܷ��ָ�(" & .RowData(lngRow) & "," & int���� & ")"
                AddArray cllProc, gstrSQL
            End If
        Next
    End With
    err = 0: On Error GoTo ErrHand:
    ExecuteProcedureArrAy cllProc, Me.Caption
    mlngCount�ָ� = 0
    SaveData = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlRefreshData(ByVal frmMain As Form, ByVal strPrivs As String, ByVal lngModule As Long, ByVal intUnit As Integer, _
    ByVal arrFilter As Variant) As Boolean
     '-----------------------------------------------------------------------------------------------------------
    '����:����ˢ������
    '���:frmMain-������
    '     strPrivs-Ȩ�޴�
    '     lngModule-ģ���
    '     intUnit-��ʾ��λ(0-ɢװ��λ,1-��װ��λ)
    '     arrFilter-��������
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-22 14:25:18
    '-----------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain: mstrPrivs = strPrivs: mlngModule = lngModule:
    Set mArrFilter = arrFilter
    mintUnit = intUnit
    
    '��ʼ��ֵ
    Call Form_Load
    With vsGrid
        .Redraw = flexRDNone
        .Rows = .FixedRows + 1
        .Clear (1)
        '�������
        zlRefreshData = RefreshData
        .Redraw = flexRDBuffered
    End With
End Function
 
Private Sub Form_Resize()
    err = 0: On Error Resume Next
    With vsGrid
        .Top = ScaleTop
        .Width = ScaleWidth
        .Left = ScaleLeft
        .Height = ScaleHeight
    End With
End Sub
Public Function zlFullData(ByVal rsNotPayStuff As ADODB.Recordset) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���������ݵ�Vss�ؼ���
    '���:rsNotPayStuff-δ�����嵥
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-23 17:11:13
    '-----------------------------------------------------------------------------------------------------------
    Set mrsNotPayStuff = rsNotPayStuff
    With vsGrid
        .Redraw = flexRDNone
        .Rows = .FixedRows + 1
        .Clear (1)
        '�������
        zlFullData = LoadDataToVssGrid
        .Redraw = flexRDBuffered
    End With
    
    
End Function
 
Private Sub Form_Load()
    zl_vsGrid_Para_Restore mlngModule, vsGrid, Me.Caption, "�ܷ��嵥"
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
    
    mbln������ʱ����� = Val(zlDatabase.GetPara("����ҽ��������ʱ�����", glngSys, 1723, 0))
End Sub
Private Function LoadDataToVssGrid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����ص�������䵽ָ��������ؼ���
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-23 11:06:21
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    LoadDataToVssGrid = False
    
    err = 0: On Error GoTo ErrHand:
    mlngCount�ָ� = 0
    '������ݵ��ؼ���
    mrsNotPayStuff.Filter = 0
    If mrsNotPayStuff.RecordCount <> 0 Then mrsNotPayStuff.MoveFirst
    
    With vsGrid
        If mrs�ܷ�.RecordCount <> 0 Then mrs�ܷ�.MoveFirst
        lngRow = .FixedRows
        Do While Not mrs�ܷ�.EOF
            .RowData(lngRow) = Val(mrs�ܷ�!Id)
            .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrs�ܷ�!����)
            .TextMatrix(lngRow, .ColIndex("����ҽ��")) = NVL(mrs�ܷ�!����ҽ��)
            .TextMatrix(lngRow, .ColIndex("״̬")) = "������"
            '24-�շѴ������ϣ�25-���ʵ��������ϣ�26-���ʱ������ϣ�
            .TextMatrix(lngRow, .ColIndex("��������")) = Decode(NVL(mrs�ܷ�!����), 24, "�շѵ�", 25, "���ʵ�", 26, "���ʱ�", "��֪") & IIf(mrs�ܷ�!���շ� = 0, "(δ)", "")
            .Cell(flexcpData, lngRow, .ColIndex("��������")) = NVL(mrs�ܷ�!����)
            .TextMatrix(lngRow, .ColIndex("���ݺ�")) = NVL(mrs�ܷ�!NO)
            .Cell(flexcpData, lngRow, .ColIndex("���ݺ�")) = NVL(mrs�ܷ�!����ID)
            
            .TextMatrix(lngRow, .ColIndex("����Ա")) = NVL(mrs�ܷ�!�����)  '����Ա����
            .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrs�ܷ�!����)
            .TextMatrix(lngRow, .ColIndex("��������")) = NVL(mrs�ܷ�!����)
            .TextMatrix(lngRow, .ColIndex("סԺ��")) = NVL(mrs�ܷ�!סԺ��)
            .TextMatrix(lngRow, .ColIndex("��������")) = NVL(mrs�ܷ�!Ʒ��)
            .Cell(flexcpData, lngRow, .ColIndex("��������")) = NVL(mrs�ܷ�!ҩƷID)
            .TextMatrix(lngRow, .ColIndex("���")) = NVL(mrs�ܷ�!���)
            .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrs�ܷ�!����)
            .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrs�ܷ�!����)
            .Cell(flexcpData, lngRow, .ColIndex("����")) = NVL(mrs�ܷ�!����)
            
            '.TextMatrix(lngRow, .ColIndex("��")) = Format(Val(NVL(mrs�ܷ�!��)), "###")
            .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrs�ܷ�!����)
            .TextMatrix(lngRow, .ColIndex("����")) = Format(Val(NVL(mrs�ܷ�!����)) * mrs�ܷ�!����ϵ��, mFMT.FM_���ۼ�)
            .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(NVL(mrs�ܷ�!���)), mFMT.FM_���)
            .TextMatrix(lngRow, .ColIndex("˵��")) = NVL(mrs�ܷ�!˵��)
            .TextMatrix(lngRow, .ColIndex("����ʱ��")) = NVL(mrs�ܷ�!�Ǽ�ʱ��)
            .TextMatrix(lngRow, .ColIndex("��¼����")) = NVL(mrs�ܷ�!��¼����)
            .TextMatrix(lngRow, .ColIndex("�����־")) = NVL(mrs�ܷ�!�����־)
            lngRow = lngRow + 1: .Rows = .Rows + 1
            mrs�ܷ�.MoveNext
        Loop
        
        If mrsNotPayStuff.RecordCount <> 0 Then mrsNotPayStuff.MoveFirst
        Do While Not mrsNotPayStuff.EOF
            If mrsNotPayStuff!ִ��״̬ = 2 Then
                .RowData(lngRow) = 0
                .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsNotPayStuff!����)
                .TextMatrix(lngRow, .ColIndex("����ҽ��")) = NVL(mrsNotPayStuff!����ҽ��)
                .TextMatrix(lngRow, .ColIndex("״̬")) = ""
                '24-�շѴ������ϣ�25-���ʵ��������ϣ�26-���ʱ������ϣ�
                .TextMatrix(lngRow, .ColIndex("��������")) = NVL(mrsNotPayStuff!����)
                .TextMatrix(lngRow, .ColIndex("���ݺ�")) = NVL(mrsNotPayStuff!NO)
                .TextMatrix(lngRow, .ColIndex("����Ա")) = NVL(mrsNotPayStuff!����Ա)
                .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsNotPayStuff!����)
                .TextMatrix(lngRow, .ColIndex("��������")) = NVL(mrsNotPayStuff!����)
                .TextMatrix(lngRow, .ColIndex("סԺ��")) = NVL(mrsNotPayStuff!סԺ��)
                .TextMatrix(lngRow, .ColIndex("��������")) = NVL(mrsNotPayStuff!��������)
                .TextMatrix(lngRow, .ColIndex("���")) = NVL(mrsNotPayStuff!���)
                .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsNotPayStuff!����)
                .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsNotPayStuff!����)
                '.TextMatrix(lngRow, .ColIndex("��")) = Format(Val(NVL(mrsNotPayStuff!��)), "###")
                .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsNotPayStuff!����)
                .TextMatrix(lngRow, .ColIndex("����")) = Format(Val(NVL(mrsNotPayStuff!����)) * mrsNotPayStuff!����ϵ��, mFMT.FM_���ۼ�)
                .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(NVL(mrsNotPayStuff!���)), mFMT.FM_���)
                .TextMatrix(lngRow, .ColIndex("˵��")) = NVL(mrsNotPayStuff!˵��)
                .TextMatrix(lngRow, .ColIndex("����ʱ��")) = NVL(mrsNotPayStuff!����ʱ��)
                .TextMatrix(lngRow, .ColIndex("��¼����")) = NVL(mrsNotPayStuff!��¼����)
                .TextMatrix(lngRow, .ColIndex("�����־")) = NVL(mrsNotPayStuff!�����־)
            
                .Cell(flexcpData, lngRow, .ColIndex("���ݺ�")) = NVL(mrsNotPayStuff!����ID)
                .Cell(flexcpData, lngRow, .ColIndex("��������")) = NVL(mrsNotPayStuff!����)
                .Cell(flexcpData, lngRow, .ColIndex("��������")) = NVL(mrsNotPayStuff!����ID)
                .Rows = .Rows + 1
                lngRow = lngRow + 1
            End If
            mrsNotPayStuff.MoveNext
         Loop
         If .Rows > 2 Then .Rows = .Rows - 1
        If .Rows > 2 Then
            .Cell(flexcpBackColor, 1, .ColIndex("״̬"), .Rows - 1, .ColIndex("״̬")) = &HE7CFBA
        End If
         
    End With
    LoadDataToVssGrid = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function RefreshData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����ˢ�¾ܷ�����
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-24 10:47:45
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, strWhere As String, strWhere1 As String
    Dim lngRow As Long, strFields As String
    Dim str���� As String
    Dim str�������� As String
    Dim strסԺ As String
    Dim strSqlTmp As String
    
    On Error GoTo ErrHandle
    str�������� = zlDatabase.GetPara("�������Ϸ�ʽ", glngSys, mlngModule, "�ٴ�,����,���,����,����,����,Ӫ��")
    
    If mintUnit = 0 Then
        strFields = "X.���㵥λ ��λ, 1 ����ϵ��,"
    Else
        strFields = " D.��װ��λ ��λ, D.����ϵ��,"
    End If
    
    strWhere = ""
    If (Trim(mArrFilter("���ݺ�")(0)) <> "" And Trim(mArrFilter("���ݺ�")(1)) = "") Then
        strWhere = strWhere & "            AND s.NO =[6]  "
    ElseIf (Trim(mArrFilter("���ݺ�")(1)) <> "" And Trim(mArrFilter("���ݺ�")(0)) = "") Then
        strWhere = strWhere & "            AND s.NO =[7]  "
    ElseIf Trim(mArrFilter("���ݺ�")(0)) <> "" And Trim(mArrFilter("���ݺ�")(1)) <> "" Then
        strWhere = strWhere & "            AND ( s.NO between [6] and [7] )"
    End If
    
    strWhere1 = ""
    'If Val(mArrFilter("��������id")) <> 0 Then strWhere1 = "  AND c.��������ID=[5]  "
    If Trim(mArrFilter("��������ID")) <> "" Then
        Select Case Val(mArrFilter("��������"))
        Case 0  '�ٴ�
            strWhere1 = strWhere1 & " And Instr([5], ',' || c.��������id || ',') > 0 And c.���˿���id=c.��������id"
        Case 1 'ҽ��
            strWhere1 = strWhere1 & " And Instr([5], ',' || c.��������id || ',') > 0 And c.���˿���id<>c.��������id"
        Case Else
            '����
            If str�������� = "" Then
                strWhere1 = strWhere1 & " And Instr([5], ',' || c.���˲���ID || ',') > 0 And c.���˿���id=c.��������id"
            Else
                strWhere1 = strWhere1 & " And Instr([5], ',' || c.���˲���ID || ',') > 0 "
                If str�������� <> mstrAllType Then
                    strWhere1 = strWhere1 & " And C.��������id Not In (Select Distinct ����id From ��������˵�� " & _
                        " Where Instr([13],',' || �������� || ',') > 0) "
                End If
            End If
        End Select
    End If
        
    strWhere1 = strWhere1 & IIf(Val(mArrFilter("����ID")) = 0, "", "  AND c.����iD=[8]  ")
    strWhere1 = strWhere1 & IIf(Val(mArrFilter("סԺ��")) = 0, "", "  AND c.��ʶ��=[9] and c.�����־=2 ")
    strWhere1 = strWhere1 & IIf(Trim(mArrFilter("����")) = "", "", "  AND c.���� like [10] ")
    strWhere1 = strWhere1 & IIf(Val(mArrFilter("�����")) = 0, "", "  AND C.��ʶ��=[11] and c.�����־=1 ")
    strWhere1 = strWhere1 & IIf(Trim(mArrFilter("���￨��")) = "", "", "  AND c1.���￨��=[12]")
    
    
    If mbln������ʱ����� = False Then
        strSqlTmp = "Select S.ID, S.ҩƷid, S.��ҩ��, S.����, Nvl(S.����, 0) As ����, S.NO, S.����, S.ʵ������ As ����, S.����, s.����, S.����, " & _
            " S.���ۼ� As ����, S.���۽�� As ���, S.����, S.Ƶ��, S.�÷�, S.ժҪ ˵��, S.����id, S.�Է�����id " & _
            " From ҩƷ�շ���¼ S " & _
            " Where Mod(S.��¼״̬, 3) = 1 And Nvl(LTrim(RTrim(S.ժҪ)), 'Not�ܷ�') = '�ܷ�' And S.����� Is Null  " & _
            " And (S.�ⷿid + 0 = [1] Or S.�ⷿid Is Null)  " & strWhere & _
            " And (S.�������� Between [2] And [3] )  And S.���� In (Select * From Table(Cast(f_Num2List([4]) As Zltools.t_NumList)))"
    Else
        strSqlTmp = "" & _
            "Select S.ID, S.ҩƷid, S.��ҩ��, S.����, Nvl(S.����, 0) As ����, S.NO, S.����, S.ʵ������ As ����, S.����, s.����, S.����, " & _
            " S.���ۼ� As ����, S.���۽�� As ���, S.����, S.Ƶ��, S.�÷�, S.ժҪ ˵��, S.����id, S.�Է�����id " & _
            " From ҩƷ�շ���¼ S, ������ü�¼ A " & _
            " Where Mod(S.��¼״̬, 3) = 1 And Nvl(LTrim(RTrim(S.ժҪ)), 'Not�ܷ�') = '�ܷ�' And S.����� Is Null  " & _
            " And (S.�ⷿid + 0 = [1] Or S.�ⷿid Is Null)  " & strWhere & _
            " And (S.�������� Between [2] And [3] )  And S.���� In (Select * From Table(Cast(f_Num2List([4]) As Zltools.t_NumList))) " & _
            " And S.����id = A.Id And A.ҽ����� Is Null "
        strSqlTmp = strSqlTmp & " Union All " & _
            "Select S.ID, S.ҩƷid, S.��ҩ��, S.����, Nvl(S.����, 0) As ����, S.NO, S.����, S.ʵ������ As ����, S.����, s.����, S.����, " & _
            " S.���ۼ� As ����, S.���۽�� As ���, S.����, S.Ƶ��, S.�÷�, S.ժҪ ˵��, S.����id, S.�Է�����id " & _
            " From ҩƷ�շ���¼ S, ������ü�¼ A " & _
            " Where Mod(S.��¼״̬, 3) = 1 And Nvl(LTrim(RTrim(S.ժҪ)), 'Not�ܷ�') = '�ܷ�' And S.����� Is Null  " & _
            " And (S.�ⷿid + 0 = [1] Or S.�ⷿid Is Null)  " & strWhere & _
            " And (A.����ʱ�� Between [2] And [3] )  And S.���� In (Select * From Table(Cast(f_Num2List([4]) As Zltools.t_NumList))) " & _
            " And S.����id = A.Id And A.ҽ����� Is Not Null "
    End If
    
    gstrSQL = "" & _
     "Select Distinct S.ID, S.ҩƷid, P.���� ����, S.��ҩ��,C.����Ա���� As ����ҽ��, C.����Ա���� �����, S.����, S.����, S.NO, '' ����, C.����,C.��ʶ�� as סԺ��, " & _
     "                 C.��¼����,C.�����־,C.�Ǽ�ʱ��, '[' || X.���� || ']' || X.���� Ʒ��, S.���� ��, S.����, Nvl(D.���÷���, 0) ����, X.���, " & strFields & _
     "                Decode(S.����, Null, '', S.����) ����, Nvl(S.����, 0) ���� , S.����, " & _
     "                 S.���, S.����, S.Ƶ��, S.�÷�, S.˵��, C.ҽ�����,S.����ID,C.��¼״̬ As ���շ�, Nvl(s.����, Nvl(x.����, '')) ���� " & _
     " From (" & strSqlTmp & ") S, ������ü�¼ C,������Ϣ C1, ���ű� P, " & _
     "      �������� D, �շ���ĿĿ¼ X, �շ���Ŀ���� A  " & _
     " Where S.ҩƷid = D.����id And S.ҩƷid = X.ID  And S.ҩƷid = A.�շ�ϸĿid(+) And A.����(+) = 3  " & _
     "       And S.�Է�����id = P.ID  And S.����id = C.ID And Nvl(c.����״̬,0)<>1 and C.����ID=C1.����id(+)  " & vbCrLf & strWhere1
    
    '�ų���δ��ҩƷ�����ʼ�¼
    gstrSQL = gstrSQL & " And Not Exists (Select 1 From ���˷������� X " & _
        " Where X.������� = 0 And X.״̬+0 = 0 And X.�շ�ϸĿid+0 = S.ҩƷid And X.����id = S.����id) "
    
    '�շѴ�����ʾ��ʽ
    If Val(mArrFilter("�շѴ���")) = 1 Then
        gstrSQL = gstrSQL & " And C.��¼״̬=1 "
    ElseIf Val(mArrFilter("�շѴ���")) = 2 Then
        gstrSQL = gstrSQL & " And C.��¼״̬=0 "
    End If
    
    If Val(mArrFilter("��������")) = 0 Then
        '����
        str���� = Replace(gstrSQL, "c.���˲���ID", "c.��������id")
        strסԺ = Replace(gstrSQL, "'' ����", "c.����")
        strסԺ = Replace(strסԺ, "c.����", "nvl(R.����,c.����)")
        strסԺ = Replace(strסԺ, "C.����", "nvl(R.����,C.����) ����")
        strסԺ = Replace(strסԺ, "������ü�¼ C", "סԺ���ü�¼ C,������ҳ r")
        strסԺ = Replace(strסԺ, "And Nvl(c.����״̬,0)<>1", "and r.����id=c.����id and r.��ҳid=c.��ҳid " & IIf(Trim(mArrFilter("����")) = "", "", "   AND c.���� =[14] "))
        If Trim(mArrFilter("����")) <> "" Then str���� = str���� & " and 1=0"
        gstrSQL = str���� & " Union All " & strסԺ
    ElseIf Val(mArrFilter("��������")) = 1 Then
        gstrSQL = Replace(gstrSQL, "c.���˲���ID", "c.��������id")
    ElseIf Val(mArrFilter("��������")) = 2 Then
        'סԺ���ʵ�
        gstrSQL = Replace(gstrSQL, "'' ����", "c.����")
        gstrSQL = Replace(gstrSQL, "c.����", "nvl(R.����,c.����)")
        gstrSQL = Replace(gstrSQL, "C.����", "nvl(R.����,C.����) ����")
        gstrSQL = Replace(gstrSQL, "������ü�¼ C", "סԺ���ü�¼ C,������ҳ r")
        gstrSQL = Replace(gstrSQL, "And Nvl(c.����״̬,0)<>1", "and r.����id=c.����id and r.��ҳid=c.��ҳid " & IIf(Trim(mArrFilter("����")) = "", "", "   AND c.���� =[14] "))
    End If
     
    gstrSQL = gstrSQL & " Order By NO, ����"
    
    Set mrs�ܷ� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
        Val(mArrFilter("���ϲ���ID")), _
        CDate(mArrFilter("���ڷ�Χ")(0)), CDate(mArrFilter("���ڷ�Χ")(1)), _
        CStr("," & mArrFilter("����") & ","), _
        Val(mArrFilter("��������ID")), _
        CStr(mArrFilter("���ݺ�")(0)), CStr(mArrFilter("���ݺ�")(1)), _
        Val(mArrFilter("����ID")), Val(mArrFilter("סԺ��")), _
        CStr(mArrFilter("����")), Val(mArrFilter("�����")), CStr(mArrFilter("���￨��")), "," & str�������� & ",", Val(mArrFilter("����")))
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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
Public Property Get zlHaveSel�ָ�() As Boolean
    zlHaveSel�ָ� = mlngCount�ָ� > 0
End Property


Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Restore mlngModule, vsGrid, Me.Caption, "�ܷ��嵥"
End Sub

Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsGrid
        Select Case Col
        Case .ColIndex("״̬")
            If zlStr.IsHavePrivs(mstrPrivs, "�������ϻָ�") = False Then Cancel = True: Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub
Private Sub vsGrid_DblClick()
    Dim str״̬ As String
    If zlStr.IsHavePrivs(mstrPrivs, "�������ϻָ�") = False Then Exit Sub
    
    With vsGrid
        If .Row < 1 Then
            Exit Sub
        End If
        str״̬ = Trim(.TextMatrix(.Row, .ColIndex("״̬")))
        If str״̬ = "" Or str״̬ = "�ܷ�" Then Exit Sub
        .TextMatrix(.Row, .ColIndex("״̬")) = Decode(str״̬, "�ָ�", "������", "�ָ�")
        If .TextMatrix(.Row, .ColIndex("״̬")) = "�ָ�" Then
            mlngCount�ָ� = mlngCount�ָ� + 1
        Else
            mlngCount�ָ� = mlngCount�ָ� - 1
        End If
    End With
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
End Sub




