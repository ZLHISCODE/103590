VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm���������嵥 
   BorderStyle     =   0  'None
   Caption         =   "�����嵥"
   ClientHeight    =   4650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   3900
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   7485
      _cx             =   13203
      _cy             =   6879
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
      BackColorSel    =   16711680
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   22
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frm���������嵥.frx":0000
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
Attribute VB_Name = "frm���������嵥"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
Private mrsBackStuff As ADODB.Recordset
Private mintUnit As Integer
Private mstrPrivs As String
Private mlngModule As Long
Private mArrFilter As Variant   '��������
Private mfrmMain As Form        '������
Private mbln������ǩ�� As Boolean
Private mblnHave���� As Boolean '�Ƿ�ѡ���������ݵ�
Private Const mstrAllType As String = "�ٴ�,����,���,����,����,����,Ӫ��"

'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
Private mblnFilterChange As Boolean '���������˸ı�,��Ҫ����ˢ������
Private mbln��ʾ�������� As Boolean   '��ʾ�������̵ĵ���


Public Event zlRefreshDataRecordSet(ByVal rsNotStuffStuff As ADODB.Recordset)

Private mobjPlugIn As Object             '��ҽӿڶ���

Public Property Get In_PlugIn() As Object
    Set In_PlugIn = mobjPlugIn
End Property
Public Property Set In_PlugIn(ByVal objVal As Object)
    Set mobjPlugIn = objVal
End Property
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
        .ColData(.ColIndex("׼������")) = 1
        .ColData(.ColIndex("��������")) = 1
        
        .Cell(flexcpForeColor, 0, .ColIndex("״̬")) = vbBlue
        .Cell(flexcpForeColor, 0, .ColIndex("��������")) = vbBlue
        
        .Cell(flexcpFontBold, 0, .ColIndex("״̬")) = True
        .Cell(flexcpFontBold, 0, .ColIndex("��������")) = True
    End With
End Sub


Public Function zlBackPayStuff() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���Ѿ�������������
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-22 14:25:18
    '-----------------------------------------------------------------------------------------------------------
    If ISValied() = False Then Exit Function
    If SaveData() = False Then Exit Function
    zlBackPayStuff = True
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
    Set mfrmMain = frmMain: mstrPrivs = strPrivs: mlngModule = lngModule
    Set mArrFilter = arrFilter
    mintUnit = intUnit
    
    '��ʼ��ֵ
    Call initPara
    zlRefreshData = RefreshData

End Function

 
Private Sub initPara()
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
    With vsGrid
        .Editable = flexEDKbdMouse
    End With
    
End Sub
Private Sub Form_Load()
    mblnFilterChange = True
    mbln��ʾ�������� = False
    zl_vsGrid_Para_Restore mlngModule, vsGrid, Me.Caption, "�����嵥"
    Call InitVsGrid
    Call initPara
End Sub

Private Sub Form_Resize()
    err = 0: On Error Resume Next
    With vsGrid
        .Top = Me.ScaleTop
        .Width = ScaleWidth
        .Left = ScaleLeft
        .Height = ScaleHeight - .Top
    End With
End Sub
Private Sub initRecStruc()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ���ڲ���¼��
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-04-24 13:16:04
    '-----------------------------------------------------------------------------------------------------------
   '�ѷ�������¼��
    Set mrsBackStuff = New ADODB.Recordset
    With mrsBackStuff
        If .State = 1 Then .Close
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "ִ��״̬", adDouble, 1, adFldIsNullable
        .Fields.Append "��¼״̬", adDouble, 18, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "סԺ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "Ч��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adDouble, 2, adFldIsNullable
        .Fields.Append "����ϵ��", adDouble, 3, adFldIsNullable
        .Fields.Append "��", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "������", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "׼����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "������", adDouble, 18, adFldIsNullable
        .Fields.Append "�ɲ���", adDouble, 2, adFldIsNullable
        .Fields.Append "��λ", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "Ƶ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�÷�", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "˵��", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "����Ա", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����ҽ��", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "����ʱ��", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "λ��", adDouble, 18, adFldIsNullable
        .Fields.Append "ҽ��id", adDouble, 18, adFldIsNullable
        .Fields.Append "������", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "ʵ������", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "ʵ�ʼ۸�", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��¼����", adDouble, 18, adFldIsNullable
        .Fields.Append "�����־", adDouble, 18, adFldIsNullable
        .Fields.Append "����Ա", adLongVarChar, 50, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Function RefreshData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����ˢ������
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-24 13:37:03
    '-----------------------------------------------------------------------------------------------------------
    Dim strWhere As String, strWhere1 As String, strTemp As String, strFields As String
    Dim blnHistory As Boolean, strTable As String, strTable1 As String, rsTemp As New ADODB.Recordset
    Dim str���� As String
    Dim str�������� As String
    
    On Error GoTo ErrHandle
    str�������� = zlDatabase.GetPara("�������Ϸ�ʽ", glngSys, mlngModule, "�ٴ�,����,���,����,����,����,Ӫ��")
    
    If mblnFilterChange = False Then RefreshData = True: Exit Function
    
    '�ȼ���Ƿ������ʷ����
    blnHistory = zlDatabase.DateMoved(mArrFilter("���ڷ�Χ")(0), , , Me.Caption)
    
    Select Case mintUnit
    Case 0  'ɢװ��λ
         strFields = "X.���㵥λ ��λ,1 as ����ϵ��, "
    Case Else
         strFields = "D.��װ��λ ��λ,D.����ϵ��,"
    End Select
    

    strWhere1 = ""
    If (Trim(mArrFilter("���ݺ�")(0)) <> "" And Trim(mArrFilter("���ݺ�")(1)) = "") Then
        strWhere1 = strWhere1 & "            AND A.NO =[6]  "
    ElseIf (Trim(mArrFilter("���ݺ�")(1)) <> "" And Trim(mArrFilter("���ݺ�")(0)) = "") Then
        strWhere1 = strWhere1 & "            AND A.NO =[7]  "
    ElseIf Trim(mArrFilter("���ݺ�")(0)) <> "" And Trim(mArrFilter("���ݺ�")(1)) <> "" Then
        strWhere1 = strWhere & "            AND ( A.NO between [6] and [7] )"
    End If
    
    
    'If Val(mArrFilter("��������id")) <> 0 Then strWhere1 = "  AND H.��������ID=[5]  "
    If Trim(mArrFilter("��������ID")) <> "" Then
        Select Case Val(mArrFilter("��������"))
        Case 0  '�ٴ�
            strWhere1 = strWhere1 & " And Instr([5], ',' || H.��������id || ',') > 0 And H.���˿���id=H.��������id"
        Case 1 'ҽ��
            strWhere1 = strWhere1 & " And Instr([5], ',' || H.��������id || ',') > 0 And H.���˿���id<>H.��������id"
        Case Else
            '����
            If str�������� = "" Then
                strWhere1 = strWhere1 & " And Instr([5], ',' || H.���˲���ID || ',') > 0 And H.���˿���id=H.��������id"
            Else
                strWhere1 = strWhere1 & " And Instr([5], ',' || H.���˲���ID || ',') > 0 "
                If str�������� <> mstrAllType Then
                    strWhere1 = strWhere1 & " And H.��������id Not In (Select Distinct ����id From ��������˵�� " & _
                        " Where Instr([13],',' || �������� || ',') > 0) "
                End If
            End If
        End Select
    End If
    
    strWhere1 = strWhere1 & IIf(Val(mArrFilter("����ID")) = 0, "", "  AND H.����iD=[8]  ")
    strWhere1 = strWhere1 & IIf(Val(mArrFilter("סԺ��")) = 0, "", "  AND H.��ʶ��=[9] and H.�����־=2 ")
    strWhere1 = strWhere1 & IIf(Val(mArrFilter("�����")) = 0, "", "  AND H.��ʶ��=[11] and H.�����־=1 ")
    strWhere1 = strWhere1 & IIf(Trim(mArrFilter("���￨��")) = "", "", "  AND H1.���￨��=[12]  ")
 
    '��ȡ�ѷ��ϻ����ϵĽ��
    strTable = " " & _
    "   Select A.ID, A.NO, A.����, A.���, A.ҩƷid, A.����id, A.����, A.����, A.Ч��, Nvl(A.����, 0) ����, " & _
    "          Nvl(A.����, 1) ����, A.ʵ������ ʵ������, Nvl(A.����, 1) * A.ʵ������ - B.�ѷ����� ��������, B.�ѷ�����, " & _
    "          A.��¼״̬, A.���ۼ�, A.���۽��, A.����, A.Ƶ��, A.�÷�, A.ժҪ, A.�����, A.�������, A.�Է�����id, A.�ⷿid, " & _
    "          A.����, Decode(Nvl(A.������, ''), '', '', DECODE(mod(a.��¼״̬,3),2,'(��)','(��)') || A.������) ������, H.ҽ�����,H.����Ա����, " & _
    "          H.��� As �������,H.������ as ����ҽ��,H.����,H.����id,H.��¼����,H.�����־,H.��ʶ��,'' ����,1 �ɲ���" & _
    "   From ҩƷ�շ���¼ A, ������ü�¼ H,������Ϣ H1, " & _
    "(Select a.No, a.����, a.ҩƷid, a.���, Sum(Nvl(a.����, 1) * a.ʵ������) �ѷ�����" & vbNewLine & _
    "From ҩƷ�շ���¼ A," & vbNewLine & _
    "     (Select NO, ����, �ⷿid,���,ҩƷid From ҩƷ�շ���¼ Where �ⷿid + 0 = [1] And ������� Between [2] And [3] and (��¼״̬ = 1 Or Mod(��¼״̬, 3) = 0)) B" & vbNewLine & _
    "Where a.����� Is Not Null And a.�ⷿid = b.�ⷿid And a.No = b.No and A.���=B.��� and A.ҩƷid + 0 = B.ҩƷid And a.���� = b.����" & vbNewLine & _
    "Group By a.No, a.����, a.ҩƷid, a.���) B" & _
    "   Where A.NO = B.NO And A.���� = B.���� And A.ҩƷid + 0 = B.ҩƷid And A.��� = B.���  " & _
    "         And A.����� Is Not Null And (A.��¼״̬ = 1 Or Mod(A.��¼״̬, 3) = 0)  " & _
    "         And A.����id = H.ID And H.����ID=H1.����id(+) "
    
    If mbln��ʾ�������� = False Then
         strTable = strTable & " And B.�ѷ����� <> 0 " & vbCrLf & strWhere1
        
        If blnHistory Then
            strTable = AnalyseHistorySQL(strTable, "1 �ɲ���", "-99 �ɲ���")
        End If
       
        gstrSQL = " " & _
        "   Select /*+ cardinality(J,10)*/ Distinct S.ID, S.����, S.ҩƷid, S.NO, S.���, S.����, P.���� ����,S.��¼����,S.�����־, S.��ʶ��, S.����id, S.����,S.����Ա����, " & _
        "                   S.����, '[' || X.���� || ']' || X.���� Ʒ��, Nvl(D.���÷���, 0) ����, X.���," & strFields & _
        "                   S.���� ��, S.ʵ������ ����, S.��������, S.�ѷ����� ׼����, " & _
        "                   Decode(S.����, Null, '', S.����)  ����, " & _
        "                   Nvl(S.����, 0) ����, S.Ч��, S.���ۼ� ����, S.���۽�� ���, S.����, S.Ƶ��, S.�÷�, S.ժҪ ˵��, " & _
        "                   S.�����, To_Char(S.�������, 'YYYY-MM-DD HH24:MI:SS') ����ʱ��, 1 �ɲ���, S.ҽ�����, I.���㵥λ, " & _
        "                   Nvl(S.����, Nvl(X.����, '')) ����, Nvl(M.�����, -1) �����, Nvl(S.ҽ�����, -1) ҽ��id, S.������, " & _
        "                   '' �ⷿ��λ, M.���id, S.�������, Z.���� As ������,S.��¼״̬,S.����ҽ�� " & _
        "   From (" & vbCrLf & strTable & vbCrLf & ") S, ���ű� P,Table(Cast(f_Num2List([4]) As zlTools.t_NumList)) J, " & _
        "        �������� D, �շ���ĿĿ¼ X, �շ���Ŀ���� A, ������ĿĿ¼ I, ����ҽ����¼ M, ������Ŀ���� Z " & _
        "   Where S.ҩƷid = D.����id And S.�Է�����id + 0 = P.ID And D.����id = X.ID And D.����id = I.ID And S.ҽ����� = M.ID(+) And " & _
        "         D.����id = Z.������Ŀid(+) And Z.����(+) = 2 And D.����id = A.�շ�ϸĿid(+) And A.����(+) = 3 And S.���� =J.Column_Value  And " & _
        "         (S.��¼״̬ = 1 Or Mod(S. ��¼״̬, 3) = 0) And S.����� Is Not Null And S.�ⷿid + 0 = [1]  And " & _
        "         S.ʵ������ * S.���� > S.�������� "
    Else
        '�嵥��ʾÿ�ʲ�������
        strTable = strTable & strWhere1
        If blnHistory Then
            strTable = AnalyseHistorySQL(strTable, "1 �ɲ���", "-99 �ɲ���")
        End If
        
        strTable1 = " Union All " & _
        "     Select A.ID, A.NO, A.����, A.���, A.ҩƷid, A.����id, A.����, A.����, A.Ч��, Nvl(A.����, 0), Nvl(A.����, 1) ����, " & _
        "            A.ʵ������, 0 ������, 0 �ѷ�����, A.��¼״̬, A.���ۼ�, A.���۽��, A.����, A.Ƶ��, A.�÷�, A.ժҪ, A.�����, " & _
        "            A.�������, A.�Է�����id, A.�ⷿid, " & _
        "            A.����, " & _
        "            Decode(Nvl(A.������, ''), '', '',Decode(A.��¼״̬, 2,'(��)', '(��)' )|| A.������) ������,H.ҽ�����,H.����Ա����, " & _
        "          H.��� As �������,H.������ as ����ҽ��,H.����,H.����id,H.��¼����,H.�����־,H.��ʶ��,'' ����, Decode(A.��¼״̬, 1, 1,Mod(A.��¼״̬, 3) + 1) �ɲ��� " & _
        "     From ҩƷ�շ���¼ A, ������ü�¼ H ,������Ϣ H1" & _
        "     Where A.����id=H.id And H.����id=H1.����ID(+) and A.����� Is Not Null And Not (A.��¼״̬ = 1 Or Mod(A.��¼״̬, 3) = 0) And A.�ⷿid + 0 = [1] And " & _
        "           A.������� Between [2] And [3] " & strWhere1
        If blnHistory Then
            '��ʷ���ݣ����ܲ���
            strTable1 = AnalyseHistorySQL(strTable1, "Decode(A.��¼״̬, 1, 1,Mod(A.��¼״̬, 3) + 1) �ɲ���", "-99 �ɲ���")
        End If
        
        strTable = strTable & vbCrLf & strTable1
        gstrSQL = " " & _
        "     Select /*+ cardinality(J,10)*/ Distinct S.ID, S.����, S.ҩƷid, S.NO, S.���, S.����, P.���� ����, S.��¼����,S.�����־, S.��ʶ��, S.����id, S.����,S.����Ա����, " & _
        "                     S.����, '[' || X.���� || ']' || X.���� Ʒ��, Nvl(D.���÷���, 0) ����, X.���, " & strFields & _
        "                     S.���� ��, S.ʵ������ ����, S.��������, S.�ѷ����� ׼����, " & _
        "                     Decode(S.����, Null, '', S.����)  ����, " & _
        "                     Nvl(S.����, 0) ����, S.Ч��, S.���ۼ� ����, S.���۽�� ���, S.����, S.Ƶ��, S.�÷�, S.ժҪ ˵��, " & _
        "                     To_Char(S.�������, 'YYYY-MM-DD HH24:MI:SS') ����ʱ��, S.�����, S.�������, �ɲ���, S.ҽ�����, " & _
        "                     I.���㵥λ, Nvl(S.����, Nvl(X.����, '')) ����, Nvl(M.�����, -1) �����, " & _
        "                     Nvl(S.ҽ�����, -1) ҽ��id, S.������, '' �ⷿ��λ, Z.���� As ������,S.��¼״̬,s.����ҽ�� " & _
        "     From (" & strTable & ") S, ���ű� P, �������� D, �շ���ĿĿ¼ X,Table(Cast(f_Num2List([4]) As zlTools.t_NumList)) J, " & _
        "          �շ���Ŀ���� A, ������ĿĿ¼ I, ����ҽ����¼ M, ������Ŀ���� Z " & _
        "     Where S.ҩƷid = D.����id And D.����id = X.ID And S.�Է�����id + 0 = P.ID And D.����id = I.ID And " & _
        "           S.ҽ����� = M.ID(+) And D.����id = Z.������Ŀid(+) And Z.����(+) = 2 And D.����id = A.�շ�ϸĿid(+) And " & _
        "           A.����(+) = 3 And  S.���� =J.Column_Value And S.����� Is Not Null "
    End If
    
    If Val(mArrFilter("��������")) = 0 Then
        '����
        str���� = Replace(gstrSQL, "H.���˲���ID", "H.��������ID")
        str���� = str���� & IIf(Trim(mArrFilter("����")) = "", "", "  AND S.���� like [10] ")
        
        gstrSQL = Replace(gstrSQL, "'' ����", "H.����")
        gstrSQL = Replace(gstrSQL, "S.����", "R.����")
        gstrSQL = Replace(gstrSQL, "H.����", "H.��ҳid")
        gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
        gstrSQL = Replace(gstrSQL, "����ҽ����¼ M", "����ҽ����¼ M,������ҳ r")
        gstrSQL = gstrSQL & " and r.����id=S.����id and r.��ҳid=S.��ҳid " & IIf(Trim(mArrFilter("����")) = "", "", "   AND S.���� =[14] ")
        gstrSQL = gstrSQL & IIf(Trim(mArrFilter("����")) = "", "", "  AND R.���� like [10] ")
        If Trim(mArrFilter("����")) <> "" Then str���� = str���� & " and 1=0"
        gstrSQL = str���� & " Union All " & gstrSQL
    ElseIf Val(mArrFilter("��������")) = 2 Then
        'סԺ���ʵ�
        gstrSQL = Replace(gstrSQL, "'' ����", "H.����")
        gstrSQL = Replace(gstrSQL, "H.����", "H.��ҳid")
        gstrSQL = Replace(gstrSQL, "S.����", "R.����")
        gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
        gstrSQL = Replace(gstrSQL, "����ҽ����¼ M", "����ҽ����¼ M,������ҳ r")
        gstrSQL = gstrSQL & " and r.����id=S.����id and r.��ҳid=S.��ҳid " & IIf(Trim(mArrFilter("����")) = "", "", "   AND S.���� =[14] ")
        gstrSQL = gstrSQL & IIf(Trim(mArrFilter("����")) = "", "", "  AND R.���� like [10] ")
    End If
    
    If mbln��ʾ�������� = False Then
        gstrSQL = gstrSQL & " Order By NO, ����, �������"
    Else
        gstrSQL = gstrSQL & " Order By NO, ����, �������"
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
        Val(mArrFilter("���ϲ���ID")), _
        CDate(mArrFilter("���ڷ�Χ")(0)), CDate(mArrFilter("���ڷ�Χ")(1)), _
        CStr(mArrFilter("����")), _
        "," & Trim(mArrFilter("��������ID")) & ",", _
        CStr(mArrFilter("���ݺ�")(0)), CStr(mArrFilter("���ݺ�")(1)), _
        Val(mArrFilter("����ID")), Val(mArrFilter("סԺ��")), _
        CStr(mArrFilter("����")), Val(mArrFilter("�����")), CStr(mArrFilter("���￨��")), "," & str�������� & ",", Val(mArrFilter("����")))
    Call WhiteDataToRecord(rsTemp)
    With vsGrid
        .Redraw = flexRDNone
        .Rows = .FixedRows + 1
        .Clear (1)
        '�������
        Call FullDataToVsGrid
        .Redraw = flexRDBuffered
    End With
    mblnHave���� = False
    mblnFilterChange = False
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function FullDataToVsGrid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����ص�������䵽ָ��������ؼ���
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-23 11:06:21
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    FullDataToVsGrid = False
    
    err = 0: On Error GoTo ErrHand:

    '������ݵ��ؼ���
    If mrsBackStuff.RecordCount <> 0 Then mrsBackStuff.MoveFirst
    With vsGrid
        .Clear (1)
        If mrsBackStuff.EOF Then '
            .Rows = 2
            FullDataToVsGrid = True
            Exit Function
        End If
        
        .Rows = mrsBackStuff.RecordCount + .FixedRows
        lngRow = .FixedRows
        Do While Not mrsBackStuff.EOF
            .RowData(lngRow) = Val(mrsBackStuff!Id)
            .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsBackStuff!����)
            .TextMatrix(lngRow, .ColIndex("����ҽ��")) = NVL(mrsBackStuff!����ҽ��)
            .TextMatrix(lngRow, .ColIndex("״̬")) = IIf(Val(NVL(mrsBackStuff!ִ��״̬)) = 1, "������", "����")
            .TextMatrix(lngRow, .ColIndex("��������")) = NVL(mrsBackStuff!����)
            .TextMatrix(lngRow, .ColIndex("���ݺ�")) = NVL(mrsBackStuff!NO)
            .TextMatrix(lngRow, .ColIndex("����Ա")) = NVL(mrsBackStuff!����Ա)
            .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsBackStuff!����)
            .TextMatrix(lngRow, .ColIndex("��������")) = NVL(mrsBackStuff!����)
            .TextMatrix(lngRow, .ColIndex("סԺ��")) = NVL(mrsBackStuff!סԺ��)
            .TextMatrix(lngRow, .ColIndex("��������")) = NVL(mrsBackStuff!��������)
            .TextMatrix(lngRow, .ColIndex("���")) = NVL(mrsBackStuff!���)
            .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsBackStuff!����)
            .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsBackStuff!����) & IIf(Val(NVL(mrsBackStuff!����)) = 0, "", "(" & NVL(mrsBackStuff!����) & ")")
            .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsBackStuff!����)
            .TextMatrix(lngRow, .ColIndex("��������")) = NVL(mrsBackStuff!������)
            .TextMatrix(lngRow, .ColIndex("׼������")) = NVL(mrsBackStuff!׼����)
            .TextMatrix(lngRow, .ColIndex("��������")) = IIf(Val(NVL(mrsBackStuff!ִ��״̬)) = 1, Format("0", mFMT.FM_����), NVL(mrsBackStuff!׼����))
            .TextMatrix(lngRow, .ColIndex("����")) = Format(Val(NVL(mrsBackStuff!����)) * mrsBackStuff!����ϵ��, mFMT.FM_���ۼ�)
            .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(NVL(mrsBackStuff!���)), mFMT.FM_���)
            .TextMatrix(lngRow, .ColIndex("����Ա")) = NVL(mrsBackStuff!����Ա)
            .TextMatrix(lngRow, .ColIndex("����ʱ��")) = NVL(mrsBackStuff!����ʱ��)
            .TextMatrix(lngRow, .ColIndex("��/������")) = NVL(mrsBackStuff!������)
            
            .Cell(flexcpData, lngRow, .ColIndex("���ݺ�")) = Val(NVL(mrsBackStuff!λ��))
            .Cell(flexcpData, lngRow, .ColIndex("��������")) = Val(NVL(mrsBackStuff!�ɲ���))
            .Cell(flexcpData, lngRow, .ColIndex("״̬")) = Val(NVL(mrsBackStuff!��¼״̬))
            SetGRDCOLOR vsGrid, lngRow, IIf(NVL(mrsBackStuff!�ɲ���) = 1, 1, NVL(mrsBackStuff!��¼״̬, 0))
            lngRow = lngRow + 1
            mrsBackStuff.MoveNext
         Loop

         .Cell(flexcpFontBold, 1, .ColIndex("״̬"), .Rows - 1, .ColIndex("״̬")) = True
         .Cell(flexcpFontBold, 1, .ColIndex("��������"), .Rows - 1, .ColIndex("��������")) = True
    End With
    FullDataToVsGrid = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function WhiteDataToRecord(ByVal rsSource As ADODB.Recordset) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����ص�����д���ڲ���¼��(δ���ϲ���)
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-23 10:03:41
    '-----------------------------------------------------------------------------------------------------------
    Dim ArrayPhysic
    Dim IntArray As Integer
     
    err = 0: WhiteDataToRecord = False
    
    Call initRecStruc
    
    With rsSource
        Do While Not .EOF
            mrsBackStuff.AddNew
            mrsBackStuff!Id = !Id
            mrsBackStuff!����ID = !ҩƷID
            mrsBackStuff!λ�� = .AbsolutePosition
            mrsBackStuff!���� = !����
            mrsBackStuff!���� = Decode(NVL(!����), 24, "�շѵ�", 25, "���ʵ�", 26, "���ʱ�", "��֪")
            mrsBackStuff!ִ��״̬ = 1                        'ȱʡΪ������
            mrsBackStuff!NO = !NO
            mrsBackStuff!���� = !����
            mrsBackStuff!����ҽ�� = NVL(!����ҽ��)
            mrsBackStuff!��� = !���
            mrsBackStuff!����ID = Val(NVL(!����ID))
            mrsBackStuff!���� = !����
            mrsBackStuff!���� = IIf(IsNull(!����), "", !����)
            mrsBackStuff!סԺ�� = IIf(Val(NVL(!�����־)) = 2, NVL(!��ʶ��), "")
            mrsBackStuff!�������� = !Ʒ��
            mrsBackStuff!��� = IIf(IsNull(!���), "", !���)
            mrsBackStuff!���� = IIf(IsNull(!����), "", !����)
            mrsBackStuff!���� = IIf(IsNull(!����), 0, !����)
            mrsBackStuff!���� = IIf(IsNull(!����), 0, !����)
            mrsBackStuff!���� = IIf(IsNull(!����), "", !����)
            mrsBackStuff!Ч�� = IIf(IsNull(!Ч��), "", !Ч��)
            mrsBackStuff!����ϵ�� = !����ϵ��
            mrsBackStuff!�� = IIf(IsNull(!��), 1, !��)
            mrsBackStuff!���� = Format(Val(NVL(!����)) / !����ϵ��, mFMT.FM_����) & !��λ
            mrsBackStuff!������ = Format(Val(NVL(!��������)) / !����ϵ��, mFMT.FM_����)
            mrsBackStuff!׼���� = Format(Val(NVL(!׼����)) / !����ϵ��, mFMT.FM_����)
            mrsBackStuff!������ = Format(Val(NVL(!׼����)) / !����ϵ��, mFMT.FM_����)
            mrsBackStuff!��λ = !��λ
            mrsBackStuff!���� = Format(!���� * !����ϵ��, mFMT.FM_����)
            mrsBackStuff!��� = !���
            mrsBackStuff!���� = IIf(IsNull(!����), "", zlStr.FormatEx(!����, 5) & NVL(!���㵥λ))
           ' mrsBackStuff!������λ = NVL(!���㵥λ)
            mrsBackStuff!Ƶ�� = NVL(!Ƶ��)
            mrsBackStuff!�÷� = NVL(!�÷�)
            mrsBackStuff!˵�� = NVL(!˵��)
            mrsBackStuff!����Ա = NVL(!�����)
            mrsBackStuff!����ʱ�� = NVL(!����ʱ��)
            mrsBackStuff!�ɲ��� = IIf(Val(NVL(!׼����)) = 0, 0, IIf(IsNull(!�ɲ���), 0, !�ɲ���))
            mrsBackStuff!ҽ��id = !ҽ��id
            mrsBackStuff!������ = !������
            mrsBackStuff!ʵ������ = !׼����
            mrsBackStuff!ʵ�ʼ۸� = !����
            mrsBackStuff!��¼״̬ = Val(NVL(!��¼״̬))
            mrsBackStuff!��¼���� = Val(NVL(!��¼����))
            mrsBackStuff!�����־ = Val(NVL(!�����־))
            mrsBackStuff!����Ա = NVL(!����Ա����)
            .MoveNext
        Loop
    End With
    If err <> 0 Then
        MsgBox "�����ڲ���¼��ʱ����������Ԥ֪�Ĵ���", vbInformation, gstrSysName
        Call initRecStruc
        Exit Function
    End If
    RaiseEvent zlRefreshDataRecordSet(mrsBackStuff)
    WhiteDataToRecord = True
End Function


Private Function AnalyseHistorySQL(ByVal strSQL As String, Optional strԭ�� As String = "", Optional str�ִ� As String = "") As String
    '������ʷ���ݵ�SQL���
    Dim strTemp As String
    strTemp = Replace(strSQL, "ҩƷ�շ���¼", "HҩƷ�շ���¼")
    strTemp = Replace(strTemp, "������ü�¼", "H������ü�¼")
    strTemp = Replace(strTemp, "סԺ���ü�¼", "HסԺ���ü�¼")
    If strԭ�� <> "" Then
        strTemp = Replace(strTemp, strԭ��, str�ִ�)
    End If
    strTemp = strSQL & " Union ALL " & strTemp
    AnalyseHistorySQL = strTemp
End Function
Private Function ISValied() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��鷢���Ƿ�Ϸ�
    '���:
    '����:
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-23 14:25:36
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim lngRow As Integer, lng����ID As Long, rsCheck As ADODB.Recordset, dbl�ּ� As Double
    Dim str��� As String, blnHaveData As Boolean
    
    On Error GoTo ErrHandle
    ISValied = False
    
    '�ȳ�ʼ���
    Set rsCheck = CheckBillStruct
    lng����ID = 0
    '���ִ�пⷿ
    With mrsBackStuff
        If .RecordCount <> 0 Then .MoveFirst
        If .RecordCount = 0 Then Exit Function
        If .EOF Then Exit Function
        .Sort = "����ID Asc"
        blnHaveData = False
        Do While Not .EOF
            If Val(NVL(!ִ��״̬)) = 3 And Val(NVL(!������)) <> 0 Then
                '��Ҫ���ṩ����ٶȣ��ȴ����ڲ����ݼ�
                rsCheck.Filter = "���ݱ�ʶ='" & NVL(!NO) & "|" & NVL(!����) & "'"
                If rsCheck.RecordCount <> 0 Then
                    rsCheck.Find "����ID=" & Val(NVL(!����ID))
                    If rsCheck.EOF Then rsCheck.AddNew
                Else
                    rsCheck.AddNew
                End If
                rsCheck!���ݱ�ʶ = NVL(!NO) & "|" & NVL(!����)
                rsCheck!����ID = Val(NVL(!����ID))
                rsCheck!��¼���� = Val(NVL(!��¼����))
                rsCheck!�����־ = Val(NVL(!�����־))
                str��� = NVL(rsCheck!���)
                If InStr(1, "," & str��� & ",", "," & Val(NVL(!���)) & ",") = 0 Then
                    If str��� = "" Then
                        str��� = Val(NVL(!���))
                    Else
                        str��� = str��� & "," & Val(NVL(!���))
                    End If
                    rsCheck!��� = str���
                End If
                rsCheck.Update
                rsCheck.Filter = 0
                 '���ԭ�������������ڷ���
                If Val(NVL(!����)) = 0 And Val(NVL(!����)) = 1 Then
                    ShowMsgBox "���ģ�" & NVL(!��������) & "ԭ��û�з���,�������Ƿ�����,������ֹ!"
                    Exit Function
                End If
                
                '��Ҫ���ԭ�����ּ��Ƿ�һ��
                If lng����ID <> Val(NVL(!����ID)) Then
                
                    gstrSQL = "" & _
                        "   Select  b.�ּ�, Nvl(C.�Ƿ���, 0) �Ƿ��� " & _
                        "   From  �շѼ�Ŀ b, �շ���ĿĿ¼ C " & _
                        "   where   b.�շ�ϸĿID=C.id and  (SYSDATE BETWEEN b.ִ������ AND b.��ֹ���� Or  SYSDATE >= b.ִ������ AND b.��ֹ���� IS Null)" & _
                        GetPriceClassString("B") & " And C.id=[1]"
                        
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "[ȡԭʼ�۸�����¼۸�]", Val(NVL(!����ID)))
                End If
                If rsTemp.EOF Then
                    dbl�ּ� = Val(NVL(!ʵ�ʼ۸�))
                ElseIf Val(NVL(rsTemp!�Ƿ���)) = 1 Then
                    dbl�ּ� = Val(NVL(!ʵ�ʼ۸�))
                Else
                    dbl�ּ� = Val(NVL(rsTemp!�ּ�))
                End If
                If dbl�ּ� <> Val(NVL(!ʵ�ʼ۸�)) Then
                    If MsgBox("����[" & !�������� & "(" & !��� & ")]" & "ԭ��Ϊ" & Val(NVL(!ʵ�ʼ۸�)) & ",�ּ�Ϊ" & dbl�ּ� & "��" & vbCrLf & Space(4) & "��ҩ����������������ϸ��¼���Ƿ��������? ", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then
                        !ִ��״̬ = 0
                        .Update
                    End If
                End If
                blnHaveData = True
            End If
            .MoveNext
        Loop
    End With
    If blnHaveData = False Then
        ShowMsgBox "δѡ����Ҫ���ϵ��������ϣ�����!"
        Exit Function
    End If
    Dim strNo As String, lng���� As Long, lng����id As Long
    '��鵥��,��Ҫ�Ǽ�鴦���Ƿ��Ѿ�����,�����Ƿ��Ѿ���Ժ�����Ȩ�޽�����صļ��
    With rsCheck
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strNo = !���ݱ�ʶ & "|"
            lng���� = Split(strNo, "|")(1)
            strNo = Split(strNo, "|")(0)
            lng����id = !����ID
            str��� = NVL(!���)
            
            '�����ʴ����Ƿ��ܷ���
            If Check���ʴ���(mstrPrivs, lng����, strNo, str���, Val(!��¼����), Val(!�����־)) = False Then Exit Function
            If Check��Ժ����(mstrPrivs, lng����, strNo, Val(!��¼����), Val(!�����־), lng����id) = False Then Exit Function
            .MoveNext
        Loop
    End With
    ISValied = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
 

Private Function CheckBillStruct() As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ��������
    '���:
    '����:
    '����:�ɹ�,���ؿռ�¼���ṹ
    '����:���˺�
    '����:2008-04-23 14:41:41
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    
    With rsTemp
        If .State = 1 Then .Close
        .Fields.Append "���ݱ�ʶ", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 500, adFldIsNullable
        .Fields.Append "��¼����", adDouble, 18, adFldIsNullable
        .Fields.Append "�����־", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    Set CheckBillStruct = rsTemp
End Function
Private Function SaveData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��ָ���ķ�����Ŀ�������ϴ���
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-23 11:48:06
    '-----------------------------------------------------------------------------------------------------------
    Dim strDate As String, str������ As String, lng����id As Long, strID���� As String, dbl������ As Double
    Dim int�Զ����� As Integer
    Dim cllPro As Collection
    Dim strReturnInfo As String
    Dim strReserve As String
    Dim rsTemp As New ADODB.Recordset
    Dim bln�������� As Boolean
    Dim int�Զ�����_ԭʼֵ As Integer
    
    int�Զ�����_ԭʼֵ = IIf(Val(zlDatabase.GetPara("�Զ�����", glngSys, mlngModule)) = 1, 1, 0)
     
    SaveData = False
    err = 0: On Error GoTo ErrHand:
    strDate = Format(Sys.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    
    Set cllPro = New Collection
    
    With mrsBackStuff
        If .RecordCount <> 0 Then .MoveFirst
        If .RecordCount = 0 Then Exit Function
        If .EOF Then Exit Function
        
        If MsgBox("������ȷ��Ҫ�������ϲ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
        '��ҩ��ǩ��
        str������ = ""
        If mbln������ǩ�� Then
            str������ = zlDatabase.UserIdentify(Me, "������ǩ��", glngSys, mlngModule, "����")
            If str������ = "" Then
                Exit Function
            End If
        End If
        
        '���밴����ID������ID����
        .Sort = "����ID Asc"
        Do While Not .EOF
            If !ִ��״̬ = 3 And Val(NVL(!������)) <> 0 Then
                dbl������ = Val(!������) * !����ϵ��
                If Val(!׼����) = Val(!������) Then
                    dbl������ = Val(!ʵ������)
                End If
                If dbl������ <> 0 Then
                    int�Զ����� = int�Զ�����_ԭʼֵ
                    
                    If int�Զ����� <> 1 Then
                        '�ж��Ƿ񱸻�����
                        gstrSQL = " Select 1 From ҩƷ�շ���¼ Where ���� = 21 And ������� Is Not Null And ����id = (select ����id from ҩƷ�շ���¼ where id=[1]) And Rownum < 2 "
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ񱸻�����", NVL(!Id))
                        bln�������� = Not rsTemp.EOF
                        
                        '����Ǹ�ֵ����Ҳ�����Զ�����
                        If bln�������� Then int�Զ����� = 1
                    End If
                    
                    'Zl_�����շ���¼_��������
                    gstrSQL = "Zl_�����շ���¼_��������("
                    '    �շ�id_In   In ҩƷ�շ���¼.ID%Type,
                    gstrSQL = gstrSQL & "" & NVL(!Id) & ","
                    '    �����_In   In ҩƷ�շ���¼.�����%Type,
                    gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                    '    �������_In In ҩƷ�շ���¼.�������%Type,
                    gstrSQL = gstrSQL & "to_date('" & strDate & "','yyyy-mm-dd HH24:mi:ss'),"
                    '    ����_In     In ҩƷ���.�ϴ�����%Type := Null,
                        gstrSQL = gstrSQL & "'" & NVL(!����) & "',"
                    '    Ч��_In     In ҩƷ���.Ч��%Type := Null,
                    gstrSQL = gstrSQL & "" & IIf(IsNull(!Ч��), "NULL", IIf(NVL(!Ч��) = "", "NULL", "To_Date('" & Format(!Ч��, "yyyy-MM-dd") & "','yyyy-MM-dd')")) & ","
                    '    ����_In     In ҩƷ���.�ϴβ���%Type := Null,
                    gstrSQL = gstrSQL & "'" & NVL(!����) & "',"
                    '    ��������_In In ҩƷ�շ���¼.ʵ������%Type := Null,
                    gstrSQL = gstrSQL & "" & dbl������ & ","
                    '    �Զ�����_In Integer := 0,
                    gstrSQL = gstrSQL & "" & int�Զ����� & ","
                    '    ������_In   In ҩƷ�շ���¼.������%Type := Null
                    gstrSQL = gstrSQL & "'" & str������ & "')"
                    AddArray cllPro, gstrSQL
                    
                    strReturnInfo = IIf(strReturnInfo = "", "", strReturnInfo & "|") & NVL(!Id) & "," & dbl������
                End If
                
            End If
            .MoveNext
        Loop
    End With
        
    On Error GoTo ErrExcute:
    Call ExecuteProcedureArrAy(cllPro, Me.Caption)
    SaveData = True
    
    err = 0: On Error GoTo ErrHand:
    If zlStr.IsHavePrivs(mstrPrivs, "����֪ͨ��") Then
          If MsgBox("����Ҫ��ӡ�����嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
              Call zlPrintBill(False, strDate)
          End If
    End If
    
    '������ҩ�����ҽӿ�
    If Not mobjPlugIn Is Nothing Then
        mobjPlugIn.DrugReturnByID Val(mArrFilter("���ϲ���id")), strReturnInfo, CDate(strDate), strReserve
    End If
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Function
ErrExcute:
      gcnOracle.RollbackTrans
      If ErrCenter = 1 Then Resume
      Call SaveErrLog
End Function

Public Sub zlPrintBill(ByVal bln���ϵ� As Boolean, Optional str�������� As String = "", Optional int��ʽ As Integer = 1, Optional strPrivs As String = "", Optional blnPrintAsk As Boolean = False)
    '-----------------------------------------------------------------------------------------------------------
    '����:��ӡ����
    '���:bln���ϵ�-�Ƿ��ӡ�Ѿ����ϵķ��ϵ�
    '     strDate-��������
    '     int��ʽ-���ݴ�ӡ��ʽ
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-06 10:39:42
    '-----------------------------------------------------------------------------------------------------------
    If mstrPrivs = "" Then mstrPrivs = strPrivs
    
    err = 0: On Error GoTo ErrHand:
    If str�������� = "" And blnPrintAsk = False Then
        With vsGrid
            str�������� = Trim(.TextMatrix(.Row, .ColIndex("����ʱ��")))
        End With
    End If
    If bln���ϵ� Then
        Call PrintPayBill(str��������, int��ʽ)
        Exit Sub
    End If
    
    If zlStr.IsHavePrivs(mstrPrivs, "����֪ͨ��") Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_2", Me, "����ʱ��=" & str��������, "��λ=" & mintUnit + 1, 2)
    Else
        ShowMsgBox "�㲻�߱���ӡ����֪ͨ�������Ȩ��,����ϵͳ����Ա��ϵ!"
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
 
Private Sub PrintPayBill(Optional str�������� As String = "", Optional int��ʽ As Integer)
    '-----------------------------------------------------------------------------------------------------------
    '����:���ݻ�����ӡ
    '���:
    '     intStyle:0-�����Ϸ�ʽ��ӡ,1-���ݴ�ӡ,2-���ϵ���
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-05 10:36:44
    '-----------------------------------------------------------------------------------------------------------
    Dim strReg As String
    Dim bln�ѷ����嵥 As Boolean
    Dim intMsg As Integer   '0-��ʾ��ӡ,1-�Զ���ӡ,2-����ӡ
    
    intMsg = Val(zlDatabase.GetPara("���ϴ�ӡ���ѷ�ʽ", glngSys, mlngModule, "0"))
 
    bln�ѷ����嵥 = zlStr.IsHavePrivs(mstrPrivs, "��ӡ�ѷ����嵥")
    If bln�ѷ����嵥 = False Then
        ShowMsgBox "�㲻�߱���ӡ����֪ͨ�������Ȩ��,����ϵͳ����Ա��ϵ!"
        Exit Sub
    End If
    If intMsg = 0 Then
        '��ʾ��ӡ
        If MsgBox("����Ҫ��ӡ��ص�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
    ElseIf intMsg = 1 Then
        '�Զ���ӡ
    Else
        Exit Sub
    End If
    '���ŷ���
    If str�������� = "" Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_1", Me, _
           "�ⷿ=" & Val(mArrFilter("���ϲ���ID")), _
           "���Ϸ�ʽ=���ŷ���|3", _
           "��������=" & Val(mArrFilter("��������")), _
           "���տ���=" & ��ȡ���ղ�������(str��������), _
           "��λ=" & IIf(mintUnit = 0, 0, 1), _
           "ReportFormat=" & int��ʽ, "PrintEmpty=0", 2)
    Else
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_1", Me, _
           "�ⷿ=" & Val(mArrFilter("���ϲ���ID")), _
           "���Ϸ�ʽ=���ŷ���|3", _
           "��������=" & Val(mArrFilter("��������")), _
           "���տ���=" & ��ȡ���ղ�������(str��������), _
           "��λ=" & IIf(mintUnit = 0, 0, 1), _
           "���Ϻ�=" & str��������, _
           "ReportFormat=" & int��ʽ, "PrintEmpty=0", 2)
    End If
End Sub
 
Public Function CheckPrice(ByVal lngBillId As Long, ByRef strMsg As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    '�ж��ۼ��Ƿ��ǵ�ǰ�����ۼ�
    
    On Error GoTo ErrHandle
    'ȡԭʼ�۸���ּ�
    gstrSQL = "select nvl(a.���ۼ�,0) ԭ��,b.�ּ�, Nvl(C.�Ƿ���, 0) �Ƿ��� " & _
        " from ҩƷ�շ���¼ a,�շѼ�Ŀ b, �շ���ĿĿ¼ C " & _
        " where a.ҩƷid=b.�շ�ϸĿid And A.ҩƷid = C.ID  And (SYSDATE BETWEEN b.ִ������ AND b.��ֹ���� Or  SYSDATE >= b.ִ������ AND b.��ֹ���� IS Null)" & _
        GetPriceClassString("B") & " And a.id=[1]"
        
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "[ȡԭʼ�۸�����¼۸�]", lngBillId)
    
    If rsTemp.RecordCount = 0 Then
        CheckPrice = True
        Exit Function
    End If
    
    'ʱ��ҩƷ������
    If rsTemp!�Ƿ��� = 1 Then
        CheckPrice = True
        Exit Function
    End If
    
    '�Ƚϼ۸�
    If rsTemp!ԭ�� <> rsTemp!�ּ� Then
        strMsg = "ԭ��Ϊ" & rsTemp!ԭ�� & ",�ּ�Ϊ" & rsTemp!�ּ� & "��" & vbCrLf & Space(4) & "��ҩ������������ҩ��ϸ��¼���Ƿ������ҩ? "
        CheckPrice = False
        Exit Function
    End If
    
    CheckPrice = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Property Get zlHaveData() As Boolean
    If mrsBackStuff Is Nothing Then zlHaveData = False: Exit Sub
    zlHaveData = mrsBackStuff.RecordCount <> 0
End Property
Public Property Get zlHaveSel����() As Boolean
    zlHaveSel���� = mblnHave����
End Property
Public Property Get zl��ʾ�����̵���() As Boolean
    zl��ʾ�����̵��� = mbln��ʾ��������
End Property
Public Property Let zl��ʾ�����̵���(ByVal vNewValue As Boolean)
    If vNewValue <> mbln��ʾ�������� Then
        mbln��ʾ�������� = vNewValue
        mblnFilterChange = True
         With vsGrid
            .Redraw = flexRDNone
            .Rows = .FixedRows + 1
            .Clear (1)
            '�������
            Call RefreshData
            Call FullDataToVsGrid
            .Redraw = flexRDBuffered
        End With
    End If
End Property
Public Property Get zlArrFilter() As Variant
    Set zlArrFilter = mArrFilter
End Property

Public Property Set zlArrFilter(ByVal vNewValue As Variant)
    Set mArrFilter = vNewValue
    mblnFilterChange = True
End Property

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "�����嵥"
End Sub

Private Sub vsGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsGrid
        Select Case Col
            Case .ColIndex("״̬")
            
                '�ı�ִ��״̬
                Call SetExecuteStaut(Row, 0)
            Case .ColIndex("��������")
                Call SetExecuteStaut(Row, 1)
            Case Else
        End Select
    End With
End Sub

 Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsGrid
        Select Case Col
        Case .ColIndex("״̬")
            If zlStr.IsHavePrivs(mstrPrivs, "������������") = False Then Cancel = True: Exit Sub
            '��ʷ�����ǲ��ܸ��ĵ�
            If Trim(.TextMatrix(Row, .ColIndex("��������"))) = -99 Then Cancel = True
            '��������Ҳ���ܳ���
            If Val(.Cell(flexcpData, Row, .ColIndex("��������"))) <> 1 Then Cancel = True
        Case .ColIndex("��������")
            If Val(.TextMatrix(Row, .ColIndex("׼������"))) = 0 Then Cancel = True
        Case Else
            Cancel = True
        End Select
    End With
End Sub




Private Sub vsGrid_DblClick()
    Dim str״̬ As String
    If zlStr.IsHavePrivs(mstrPrivs, "������������") = False Then Exit Sub
    
    With vsGrid
        If .Row < 1 Then
            Exit Sub
        End If
        
        If .Col = .ColIndex("��������") Then Exit Sub
        
        If Val(.Cell(flexcpData, .Row, .ColIndex("��������"))) <> 1 Then Exit Sub

        str״̬ = Trim(.TextMatrix(.Row, .ColIndex("״̬")))
        .TextMatrix(.Row, .ColIndex("״̬")) = Decode(str״̬, "����", "������", "����")
        Call SetExecuteStaut(.Row, 0)
    End With
End Sub

Private Sub SetExecuteStaut(ByVal lngRow As Long, ByVal intType As Integer)
    '-----------------------------------------------------------------------------------------------------------
    '����:����ִ��״̬
    '���:lngRow-ָ������;intType-0:����״̬;1-������������
    '����:
    '����:
    '����:���˺�
    '����:2008-04-23 11:31:04
    '-----------------------------------------------------------------------------------------------------------
    Dim str״̬ As String, int״̬ As Integer, lngλ�� As Long
    With vsGrid
        str״̬ = Trim(.TextMatrix(lngRow, .ColIndex("״̬")))
        int״̬ = Decode(str״̬, "����", 3, 0)
        lngλ�� = Val(.Cell(flexcpData, lngRow, .ColIndex("���ݺ�")))
    End With
    With mrsBackStuff
         .Filter = 0
        .MoveFirst
        .Find "λ��=" & lngλ��
        If .EOF = False Then
            If intType = 0 Then
                !ִ��״̬ = int״̬:
                If int״̬ = 3 Then
                    !������ = !׼����
                Else
                    !������ = 0
                End If
            Else
                '������С�ڵ���0���ߴ���׼���������������ϣ�״̬��־Ϊ��������
                If Val(vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("��������"))) <= 0 Or Val(vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("��������"))) > vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("׼������")) Then
                    vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("��������")) = "0"
                    int״̬ = 0
                Else
                    int״̬ = 3
                End If
                
                !ִ��״̬ = int״̬
                If int״̬ = 3 Then
                    !������ = Val(vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("��������")))
                Else
                    !������ = 0
                End If
            End If
            
            .Update
            
            vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("��������")) = Format(Val(!������), mFMT.FM_����)
        End If
        '���ܿ�������Ҫ����ǰ��״̬
        .MoveFirst
        .Find "λ��=" & lngλ��
        If .EOF = False Then
            vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("״̬")) = Decode(NVL(!ִ��״̬), 3, "����", "������")
            vsGrid.Cell(flexcpBackColor, lngRow, vsGrid.ColIndex("״̬")) = Decode(NVL(!ִ��״̬), 3, &HFFC0C0, &HFFFFFF)
            vsGrid.Cell(flexcpBackColor, lngRow, vsGrid.ColIndex("��������")) = Decode(NVL(!ִ��״̬), 3, &HFFC0C0, &HFFFFFF)
            vsGrid.Cell(flexcpForeColor, lngRow, vsGrid.ColIndex("״̬")) = Decode(NVL(!ִ��״̬), 3, vbBlue, &H80000008)
            vsGrid.Cell(flexcpForeColor, lngRow, vsGrid.ColIndex("��������")) = Decode(NVL(!ִ��״̬), 3, vbBlue, &H80000008)
        End If
        .MoveFirst
        .Find "ִ��״̬=3"
        mblnHave���� = (.EOF = False)
    End With
End Sub
Private Sub SetGRDCOLOR(ByVal objGrd As Object, ByVal lngRow As Long, ByVal int��¼״̬ As Integer)
    '-----------------------------------------------------------------------------------------------------------
    '����:���ÿؼ�����ʾ��ɫ
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-03 17:11:56
    '-----------------------------------------------------------------------------------------------------------

    Dim lngColor As Long
    Dim i As Long
    If int��¼״̬ = 1 Then
        lngColor = &H80000008
    ElseIf zlCommFun.ZyMod(int��¼״̬, 3) = 2 Then
         lngColor = vbRed
    Else
        lngColor = vbBlue
    End If
    With vsGrid
        .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = lngColor
    End With
End Sub

 Private Function ��ȡ���ղ�������(ByVal strDate As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ���ղ��ŵĴ�ӡ����
    '���:
    '����:
    '����:�ɹ�,���� ��ʾ|IN(����ID,..) ,���򷵻�""
    '����:���˺�
    '����:2008-05-05 13:31:28
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, str��ʾ As String, strIDIn As String
    If strDate = "" And mArrFilter("��������id") = "" Then Exit Function
    
    On Error GoTo ErrHandle
    If mArrFilter("��������id") = "" And strDate <> "" Then
        'û������,���Ը���ѡ�������ȡ��ʾ����
        gstrSQL = "Select distinct D.ID,D.����,D.���� as ���� " & _
                 " From ҩƷ�շ���¼ S,������ü�¼ C,���ű� d " & _
                 " Where S.����ID=C.ID And Mod(S.��¼״̬,3) In (0,1) And S.����� Is Not Null " & _
                 "      And C.ִ��״̬=1 And S.�ⷿID=[1] And S.��ҩ��ʽ=3 And S.�������=[2] " & _
                 "      And S.���� In (24,25,26) "
        Select Case Val(mArrFilter("��������"))
            Case 0  '
                gstrSQL = gstrSQL & " and C.���˿���id=d.id(+) "
            Case 1 'ҽ��
                gstrSQL = gstrSQL & "  and C.��������id =d.id(+)"
            Case Else '����
                gstrSQL = gstrSQL & "  and C.���˲���ID =d.id(+)"
        End Select
        
        If mArrFilter("����") = "24" Then
            If Val(mArrFilter("��������")) = 2 Then Exit Function
        ElseIf mArrFilter("����") = "26" Then
            gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
        ElseIf InStr(1, mArrFilter("����"), "25") > 0 Or InStr(1, mArrFilter("����"), "26") > 0 Then
            If InStr(1, mArrFilter("����"), "24") > 0 And Val(mArrFilter("��������")) = 2 Then
                gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
            Else
                gstrSQL = gstrSQL & " Union All " & Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
            End If
        End If
        
        gstrSQL = gstrSQL & "order by ����"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mArrFilter("���ϲ���id")), CDate(strDate))
        With rsTemp
            Do While Not .EOF
                If NVL(!����, "") <> "" Then
                    str��ʾ = str��ʾ & "," & !����
                    strIDIn = strIDIn & "," & !Id
                End If
                
                rsTemp.MoveNext
            Loop
        End With
        
        If str��ʾ = "" Then
            ��ȡ���ղ������� = "���п���|Is Not Null"
        Else
            strIDIn = "0" & strIDIn
            str��ʾ = str��ʾ & "|" & " IN (" & strIDIn & ")"
            ��ȡ���ղ������� = str��ʾ
        End If
        
        Exit Function
    End If
    gstrSQL = "Select ID, ���� From ���ű� A, Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) J Where ID = J.Column_Value order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(mArrFilter("��������id")))
    With rsTemp
        Do While Not .EOF
            str��ʾ = str��ʾ & "," & !����
            rsTemp.MoveNext
        Loop
    End With
    str��ʾ = str��ʾ & "|" & " IN (" & CStr(mArrFilter("��������id")) & ")"
    ��ȡ���ղ������� = str��ʾ
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

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


Private Sub vsGrid_EnterCell()
    With vsGrid
        If .Row = 0 Then Exit Sub

        .Editable = flexEDNone
        .FocusRect = flexFocusLight

        If .Row > 0 Then
            Select Case .Col
            Case .ColIndex("��������")
                .Editable = flexEDKbdMouse
                .FocusRect = flexFocusSolid
            Case .ColIndex("״̬")
                .FocusRect = flexFocusSolid
            End Select
        End If

    End With
End Sub


Private Sub vsGrid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer
    
    With vsGrid
        strKey = .EditText
        Select Case .Col
            Case .ColIndex("��������")
                intDigit = IIf(mintUnit = 1, g_С��λ��.obj_��װС��.����С��, g_С��λ��.obj_ɢװС��.����С��)
        End Select
        
        If Col = .ColIndex("��������") Then
            If InStr("1234567890" + Chr(46) + Chr(8) + Chr(13) + Chr(Asc("-")), Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
                Exit Sub
            ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                If .EditSelLength = Len(strKey) Then Exit Sub
                If InStr(strKey, ".") <> 0 And Chr(KeyAscii) = "." Then   'ֻ�ܴ���һ��С����
                    KeyAscii = 0
                    Exit Sub
                End If
                If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= intDigit And strKey Like "*.*" Then
                    KeyAscii = 0
                    Exit Sub
                Else
                    Exit Sub
                End If
            End If
        End If
    End With
End Sub


