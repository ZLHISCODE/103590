VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmBloodListInfo 
   BorderStyle     =   0  'None
   Caption         =   "ѪҺ�б���Ϣ"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame fraExecUD 
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   90
      MousePointer    =   7  'Size N S
      TabIndex        =   2
      Top             =   1185
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VSFlex8Ctl.VSFlexGrid vsList 
      Height          =   1290
      Left            =   75
      TabIndex        =   0
      Top             =   1335
      Width           =   7125
      _cx             =   12568
      _cy             =   2275
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16444122
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   2
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
   Begin RichTextLib.RichTextBox rtfOther 
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   30
      Width           =   4410
      _ExtentX        =   7779
      _ExtentY        =   1508
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmBloodListInfo.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmBloodListInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngҽ��ID As Long
Private mlngFontSize As Long
Private mblnMoved As Boolean
Private mbln��Ѫ As Boolean
Private mclsVsf As clsVsf
Private mblnFistRefresh As Boolean
Private mblnShowInfo As Boolean

Public Function zlRefresh(ByVal lngҽ��ID As Long, Optional ByVal lngFontSize As Long = 9, Optional ByVal blnMoved As Boolean = False) As Boolean
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim blnShowBlood As Boolean
    
    'SQL�����ر���
    Dim strWhere As String
    On Error GoTo ErrHand
    
    mlngҽ��ID = lngҽ��ID
    mlngFontSize = lngFontSize
    mblnMoved = blnMoved
    
    If ShowOtherAppend(blnShowBlood) = False Then Exit Function
    If blnShowBlood = True Then
        If mbln��Ѫ = True Then
            strWhere = " And a.id=f.�շ�ID and f.�䷢ID=b.id "
        Else
            strWhere = " And a.id=f.�շ�ID(+) And a.�䷢ID=b.id "
        End If
        mblnShowInfo = True
        strSQL = _
            " Select a.Id, a.ѪҺid, a.Abo,a.Rh, To_Char(a.Ч��, 'YYYY-MM-DD hh24:mi') ѪҺЧ��, a.��ɫ ѪҺ��ɫ, a.��� Ѫ�����, a.��Ѫ��," & vbNewLine & _
            "       To_Char(a.��Ѫ����, 'YYYY-MM-DD hh24:mi') ��Ѫʱ��, a.�˶��� �����, To_Char(a.�˶�����, 'YYYY-MM-DD hh24:mi') ���ʱ��," & vbNewLine & _
            "       Nvl(a.��Ѫ״̬, 0) ��Ѫ״̬,nvl(f.ִ��״̬,0) ִ��״̬," & vbNewLine & _
            "       Decode(Nvl(f.ִ��״̬, 0),0,Decode(Nvl(f.����״̬, 0), 0, Decode(Nvl(a.��Ѫ״̬, 0),0,'����',1,'�����',2,'����',9,'����',3,Decode(a.�����, Null, '��Ѫ', '�ܷ�'),''), 2, '�ܾ�����', '�ѽ���'),1,'����ִ��',2,'���ִ��',3,'ִֹͣ��') ѪҺ״̬," & vbNewLine & _
            "       c.���� �ⷿ, d.ǩ����, To_Char(d.ǩ��ʱ��, 'YYYY-MM-DD hh24:mi') ǩ��ʱ��, a.Ѫ�����, a.ʵ������ As ����, e.���� As ѪҺ����, e.��� ѪҺ���," & vbNewLine & _
            "       (Select f_List2str(Cast(Collect(g.����) As t_Strlist))" & vbNewLine & _
            "         From ������ĿĿ¼ g, ѪҺ��Ѫ���� f" & vbNewLine & _
            "         Where f.��Ѫ����id = g.Id(+) And f.�շ�id = a.Id) ��Ѫ����," & vbNewLine & _
            "       (Select Max(f.��Ѫ����) From ������ĿĿ¼ g, ѪҺ��Ѫ���� f Where f.��Ѫ����id = g.Id(+) And f.�շ�id = a.Id) ��Ѫ����,a.��Ѫ��,A.��Ѫ����,a.ȡѪ��,a.ժҪ ��ѪժҪ" & vbNewLine & _
            " From ���ű� c, Ѫ��ǩ�� d, �շ���ĿĿ¼ e, ѪҺƷ�� k, ѪҺ��� l, ѪҺ�շ���¼ a,ѪҺ���ͼ�¼ f, ѪҺ��Ѫ��¼ b" & vbNewLine & _
            " Where c.Id = a.�ⷿid And a.��Ѫǩ��id = d.Id(+) And d.����(+) = 3 And a.ѪҺid = e.Id And k.Ʒ��id = l.Ʒ��id And l.���id = a.ѪҺid And" & vbNewLine & _
            "      Nvl(a.��д����, 0) <> 0 And a.���� = 6 And Mod(a.��¼״̬, 3) = 1 " & strWhere & " And b.����id = [1]" & vbNewLine & _
            " Order By a.��Ѫ����, a.���"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "�ѷ�ѪҺ��Ϣ��ȡ", lngҽ��ID)
        Call mclsVsf.LoadGrid(rsTemp, "", True)
    Else
        mblnShowInfo = False
    End If
    Call SetFontSize(mlngFontSize)
    mblnFistRefresh = False
    Call Form_Resize
    zlRefresh = True
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ShowOtherAppend(blnShowBlood As Boolean) As Boolean
'���ܣ���ʾָ����ҽ���������Ϣ
'˵����ֻ������״̬ͨ����δͨ����ҽ��
'���أ��Ƿ���Ҫ��ȡ����ʾѪҺ�б�
    Dim strSQL As String
    Dim int���� As Integer
    Dim rsTmp As ADODB.Recordset
    Dim str����Ա As String, strʱ�� As String, str״̬ As String, str����˵�� As String
    Dim strδ��ԭ�� As String
    
    Dim intִ�б�� As Integer, int���״̬ As Integer
    Dim str��鷽�� As String, bln��Ѫ As Boolean
    Dim arrCode, arrItem
    Dim i As Integer, lngIdx As Long
    On Error GoTo errH
    
    mbln��Ѫ = False
    blnShowBlood = False
    rtfOther.Text = "": rtfOther.SelStart = 0
    '��ȡҽ���������
    strSQL = _
        " Select b.���״̬, b.��鷽��,b.ִ�б��, c.��������, c.ִ�з���" & vbNewLine & _
        " From ������ĿĿ¼ c, ����ҽ����¼ a, ����ҽ����¼ b" & vbNewLine & _
        " Where c.Id = a.������Ŀid And a.���id = b.Id And a.������� = 'E' And b.Id = [1] And b.������� = 'K'"
    If mblnMoved = True Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngҽ��ID)
    'ѪҺҽ���϶���õ����ݣ��鲻���������˳�
    If rsTmp.EOF Then
        Exit Function
    End If
    int���״̬ = Val("" & rsTmp!���״̬)
    str��鷽�� = "" & rsTmp!��鷽��
    If str��鷽�� = "" Then
        If Val("" & rsTmp!��������) = "8" And Val("" & rsTmp!ִ�з���) = 1 Then
            bln��Ѫ = True
        End If
    Else
        bln��Ѫ = Val(str��鷽��) = 1
    End If
    mbln��Ѫ = bln��Ѫ
    str����Ա = "����ˣ�": strʱ�� = "���ʱ�䣺": str״̬ = ""
    
    If intִ�б�� = -1 Then '��ȡ���δ�õ�ԭ��
        strSQL = "Select ������Ա,����ʱ��,����˵�� From ����ҽ��״̬ Where ҽ��id = [1] And �������� = [2]"
        If mblnMoved = True Then
            strSQL = Replace(strSQL, "����ҽ��״̬", "H����ҽ��״̬")
        End If
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngҽ��ID, 17)
        If Not rsTmp.EOF Then
            strδ��ԭ�� = "δ��ԭ��" & rsTmp!����˵��
            strδ��ԭ�� = strδ��ԭ�� & "(����Ա��" & rsTmp!������Ա & "  ����ʱ�䣺" & Format(rsTmp!����ʱ�� & "", "YYYY-MM-DD HH:MM:SS") & ")"
        End If
    End If
    
    '��ǰ����Ѫҽ�����޷���Ӧ��ѪҺ��¼��ʲô������ʾ
    If bln��Ѫ = True And str��鷽�� = "" Then
        Exit Function
    End If
    'Ŀǰ��Ѫ�����ҽ��״ֻ̬���⼸�֣��������������Ҫ�����ⲿ������
    Select Case int���״̬
        Case 2 '������
            If bln��Ѫ = False Then
                int���� = 15 'Ѫ�����ͨ��
                str״̬ = "�����Ѫ"
                str����Ա = "��Ѫ����ˣ�"
                strʱ�� = "��Ѫ���ʱ�䣺"
            Else
                int���� = 15 'Ѫ�����ͨ��
                str״̬ = "��ɷ�Ѫ"
                str����Ա = "��Ѫ�����ˣ�"
                strʱ�� = "��Ѫ����ʱ�䣺"
            End If
            blnShowBlood = True
        Case 3 '(������Ѫ�ּ��������δͨ�������պ����δͨ��)
            If bln��Ѫ = True Then
                '��Ѫҽ�����÷ּ���ˣ����״̬=3˵���Ǿܾ���Ѫ
                int���� = 16
                str״̬ = "�ܾ���Ѫ"
                str����Ա = "�ܾ���Ѫ�ˣ�"
                strʱ�� = "�ܾ���Ѫʱ�䣺"
                str����˵�� = "�ܾ���Ѫԭ��"
            Else
                '��Ѫҽ����Ҫ�������Ѫ��˾ܾ����Ǿܾ���Ѫ
                strSQL = "Select 1 From ѪҺ��Ѫ��¼ Where ����ID=[1] and ��¼״̬=[2]"
                Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngҽ��ID, 3)
                If rsTmp.RecordCount > 0 Then
                    int���� = 16 '�ܾ���Ѫ
                    str״̬ = "�ܾ���Ѫ"
                    str����Ա = "�ܾ���Ѫ�ˣ�"
                    strʱ�� = "�ܾ���Ѫʱ�䣺"
                    str����˵�� = "�ܾ���Ѫԭ��"
                Else
                    int���� = 12 '���δͨ��
                    str״̬ = "���δͨ��"
                    str����Ա = "����ˣ�"
                    strʱ�� = "���ʱ�䣺"
                    str����˵�� = "���δͨ��ԭ��"
                End If
            End If
        Case 4
            str״̬ = "�ȴ���Ѫ"
            int���� = 11 '����ҽ�������������Ѫ�ֻ�������������ͨ��
        Case 5
            int���� = 14
            str״̬ = "������Ѫ"
            str����Ա = "��Ѫ�����ˣ�"
            strʱ�� = "��Ѫ����ʱ�䣺"
            blnShowBlood = True
        Case 6 '��Ѫ�ƽ��պ�ֹͣ��Ѫ
            int���� = 17
            str״̬ = "ֹͣ��Ѫ"
            str����Ա = "ֹͣ��Ѫ�ˣ�"
            strʱ�� = "ֹͣ��Ѫʱ�䣺"
            str����˵�� = "ֹͣ��Ѫԭ��"
        Case 1
            If bln��Ѫ = True Then
                str״̬ = "��Ѫҽ�����˶�"
                blnShowBlood = True
            End If
        Case Else
            If bln��Ѫ = True Then
                str״̬ = "�ȴ���Ѫ"
            End If
    End Select
    rtfOther.Text = ""
    strSQL = "Select ������Ա,����ʱ��,����˵�� From ����ҽ��״̬ Where ҽ��id = [1] And �������� = [2] order by ����ʱ��"
    If mblnMoved = True Then
        strSQL = Replace(strSQL, "����ҽ��״̬", "H����ҽ��״̬")
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngҽ��ID, int����)
    
    strSQL = ""
    arrCode = Array()
    arrItem = Array()
    With rtfOther
        If Not rsTmp.EOF Then
            ReDim Preserve arrCode(UBound(arrCode) + 1)
            arrCode(UBound(arrCode)) = "״̬��" & str״̬
            ReDim Preserve arrItem(UBound(arrItem) + 1)
            arrItem(UBound(arrItem)) = "״̬��"
            
            Do While Not rsTmp.EOF
                ReDim Preserve arrCode(UBound(arrCode) + 1)
                arrCode(UBound(arrCode)) = str����Ա & rsTmp!������Ա
                ReDim Preserve arrCode(UBound(arrCode) + 1)
                arrCode(UBound(arrCode)) = strʱ�� & Format(rsTmp!����ʱ�� & "", "YYYY-MM-DD HH:MM:SS")
                If str����˵�� <> "" Then
                    ReDim Preserve arrCode(UBound(arrCode) + 1)
                    arrCode(UBound(arrCode)) = str����˵�� & rsTmp!����˵��
                End If
                ReDim Preserve arrItem(UBound(arrItem) + 1)
                arrItem(UBound(arrItem)) = str����Ա
                ReDim Preserve arrItem(UBound(arrItem) + 1)
                arrItem(UBound(arrItem)) = strʱ��
                If str����˵�� <> "" Then
                    ReDim Preserve arrItem(UBound(arrItem) + 1)
                    arrItem(UBound(arrItem)) = str����˵��
                End If
                rsTmp.MoveNext
            Loop
            If strδ��ԭ�� <> "" Then
                ReDim Preserve arrCode(UBound(arrCode) + 1)
                arrCode(UBound(arrCode)) = strδ��ԭ��
                ReDim Preserve arrItem(UBound(arrItem) + 1)
                arrItem(UBound(arrItem)) = "δ��ԭ��"
            End If
        ElseIf str״̬ <> "" Then
            strSQL = "״̬��" & str״̬
            ReDim Preserve arrCode(UBound(arrCode) + 1)
            arrCode(UBound(arrCode)) = strSQL
            ReDim Preserve arrItem(UBound(arrItem) + 1)
            arrItem(UBound(arrItem)) = "״̬��"
        End If
        .SelStart = 0
        For i = 0 To UBound(arrCode)
            .SelBold = False
            .SelText = IIf(.Text = "", "", vbCrLf) & CStr(arrCode(i))
            lngIdx = .Find(CStr(arrItem(i)), , , rtfNoHighlight Or rtfMatchCase)
            If lngIdx <> -1 Then
                .SelStart = lngIdx
                .SelLength = Len(CStr(arrItem(i)))
                .SelBold = True
                .SelIndent = 100
            End If
            .SelStart = Len(.Text)
        Next i
        If UBound(arrItem) >= 0 Then
            lngIdx = .Find(CStr(arrItem(0)), 0, , rtfNoHighlight Or rtfMatchCase)
            If lngIdx <> -1 Then .SelStart = lngIdx + Len(CStr(arrItem(0)))
        End If
    End With
    ShowOtherAppend = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Sub SetFontSize(ByVal lngFontSize As Long)
    With rtfOther
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelFontSize = lngFontSize
        .SelLength = 0
    End With
    Call gobjComlib.zlControl.VSFSetFontSize(vsList, lngFontSize)
    '�״�ˢ�»ָ�����������ˢ�»ָ�
    If mblnFistRefresh = True Then
        Call gobjComlib.RestoreWinState(Me, "zlPublicBlood")
    End If
End Sub

Private Sub Form_Load()
    mblnFistRefresh = True
    mblnShowInfo = False
    Set mclsVsf = New clsVsf
    Call InitTable
End Sub

Private Sub InitTable()
'����ʼ��
    With mclsVsf
        Call .Initialize(Me.Controls, vsList, True, False)
        Call .ClearColumn
        
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False, , , True)
        Call .AppendColumn("ID", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)  '�շ�ID
        Call .AppendColumn("״̬", 810, flexAlignLeftCenter, flexDTString, , "ѪҺ״̬") '����ִ��״̬
        Call .AppendColumn("ѪҺ����", 1800, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("���", 810, flexAlignLeftCenter, flexDTString, , "ѪҺ���")
        Call .AppendColumn("ABO", 810, flexAlignLeftCenter, flexDTString, , "ABO", True)
        Call .AppendColumn("Rh(D)", 600, flexAlignLeftCenter, flexDTString, , "RH", True)
        Call .AppendColumn("Ѫ�����", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("Ч��", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm", "ѪҺЧ��", True)
        Call .AppendColumn("����", 500, flexAlignRightCenter, flexDTDecimal, , , , , , False)
        
        
        'ѪҺ�䷢��Ϣ
        Call .AppendColumn("��Ѫ��", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("��Ѫʱ��", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm")
        Call .AppendColumn("�����", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("���ʱ��", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm")
        Call .AppendColumn("��Ѫ����", 1500, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("��Ѫ����", 1500, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("��ѪժҪ", 2000, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("��Ѫ��", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("ȡѪ��", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("��Ѫʱ��", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm", "��Ѫ����")
        
        
        '������
        Call .AppendColumn("ѪҺID", 0, flexAlignLeftCenter, flexDTString, , , , , , True)
        Call .AppendColumn("��Ѫ״̬", 0, flexAlignLeftCenter, flexDTString, , , , , , True)
        Call .AppendColumn("ִ��״̬", 0, flexAlignLeftCenter, flexDTString, , , , , , True)
            
        .AppendRows = False
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With rtfOther
        .Top = 0
        .Left = 0
        .Width = Me.Width
        .Height = IIf(mblnShowInfo = True, Me.Height - vsList.Height - fraExecUD.Height, Me.Height)
    End With
    With fraExecUD
        .Left = 0
        .Top = rtfOther.Top + rtfOther.Height
        .Visible = mblnShowInfo
    End With
    With vsList
        .Left = 0
        .Top = fraExecUD.Top + fraExecUD.Height
        .Width = Me.Width
        .Visible = mblnShowInfo
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call gobjComlib.SaveWinState(Me, "zlPublicBlood")
    Set mclsVsf = Nothing
End Sub

Private Sub fraExecUD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If rtfOther.Height + Y < 700 Or vsList.Height - Y < 700 Then Exit Sub
        fraExecUD.Top = fraExecUD.Top + Y
        rtfOther.Height = rtfOther.Height + Y
        vsList.Top = vsList.Top + Y
        vsList.Height = vsList.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub vsList_DblClick()
    '��Ѫִ�м�¼�鿴
    If Val(vsList.TextMatrix(vsList.Row, vsList.ColIndex("ID"))) < 0 Then Exit Sub
    If Not (Val(vsList.TextMatrix(vsList.Row, vsList.ColIndex("ִ��״̬"))) > 0) Then Exit Sub
    Call frmBloodExecEdit.ViewExecution(Me, Val(vsList.TextMatrix(vsList.Row, vsList.ColIndex("ID"))))
End Sub
