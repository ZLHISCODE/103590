VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmCompoundMedicine 
   BorderStyle     =   0  'None
   Caption         =   "��Һ��ҩ��¼"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   12825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picExec 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   11865
      ScaleHeight     =   3000
      ScaleWidth      =   960
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   960
      Begin XtremeCommandBars.CommandBars cbsExec 
         Left            =   120
         Top             =   30
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.Frame fraExecUD 
      BorderStyle     =   0  'None
      Height          =   2445
      Left            =   1680
      MousePointer    =   9  'Size W E
      TabIndex        =   2
      Top             =   120
      Width           =   45
   End
   Begin VSFlex8Ctl.VSFlexGrid vsExec 
      Bindings        =   "frmCompoundMedicine.frx":0000
      Height          =   2955
      Left            =   1800
      TabIndex        =   0
      Top             =   0
      Width           =   10035
      _cx             =   17701
      _cy             =   5212
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
      BackColorSel    =   16772055
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   6
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCompoundMedicine.frx":0028
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
   Begin VSFlex8Ctl.VSFlexGrid vsSend 
      Align           =   3  'Align Left
      Bindings        =   "frmCompoundMedicine.frx":0177
      Height          =   3000
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1695
      _cx             =   2990
      _cy             =   5292
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
      BackColorSel    =   16772055
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
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmCompoundMedicine.frx":018B
      ScrollTrack     =   -1  'True
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
Attribute VB_Name = "frmCompoundMedicine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event SetEditState(ByVal blnEditState As Boolean)      '���༭״̬ʱ���ý�ֹ��ת�ƽ������������
Public Event StatusTextUpdate(ByVal Text As String) 'Ҫ�����������״̬������

Private mlng����ID As Long      '��ǰ����ѡ��Ĳ���
Private mlngAdviceID As Long    '��ҩ;����ҽ��ID
Private mlng����ID  As Long
Private mlng��ҳID As Long
Private mlng��������  As Long
Private mlng����ID  As Long
Private mstr���� As String
Private mstrסԺ�� As String
Private mstr���� As String
Private mlngҽ����Ч As Long
Private mCol As Collection
Private Const Col����ʱ�� = 0
Private mrsCompoundGroup As ADODB.Recordset '��ҩ����
Private mblnEdit As Boolean '�Ƿ�������޸�ģʽ���ض����ݺ��Զ�������ģʽ
Private mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mbln�������� As Boolean '�Ǵ������Һ������Һ֮���Ƿ���Խ�����������
Private mbln����޸� As Boolean
Private mbln��ҩ���ܸ�״̬ As Boolean
Private mfrmParent As Object '���������

Private Const conMenu_Adjust = 100
Private Const conMenu_Save = 101
Private Const conMenu_Undo = 102
Private Const conMenu_AdjustCancle = 103


Public Sub RefreshData(ByVal lngAdviceID As Long, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng�������� As Long, ByVal lngҽ����Ч As Long, _
        Optional ByRef objMip As Object, Optional frmParent As Object)
'���ܣ�����ҽ����¼����ҩ;����ҽ��ID����ˢ������
    mlngAdviceID = lngAdviceID
    mlng����ID = lng����ID
    mlng����ID = lng����ID
    mlng�������� = lng��������
    mlng��ҳID = lng��ҳID
    mlngҽ����Ч = lngҽ����Ч
    Set mfrmParent = frmParent
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    Call LoadSendList
End Sub


Private Sub Form_Load()
    Dim i As Long
    
    If GetInsidePrivs(pסԺҽ������) = "" Then Exit Sub
    mbln�������� = Val(zlDatabase.GetPara("��Һ��Һ����ҩ��������������", glngSys, 1345, 0)) = 1
    mbln����޸� = Val(zlDatabase.GetPara("�������", glngSys, 1345, 0)) = 1
    mbln��ҩ���ܸ�״̬ = Val(zlDatabase.GetPara("��Һ����ҩ���ٴ�������ı���״̬", glngSys, 1345, 0)) = 1
    vsSend.Rows = vsSend.FixedRows
    vsExec.Rows = vsExec.FixedRows
    
    Set mCol = New Collection
    For i = 0 To vsExec.Cols - 1
        mCol.Add i, vsExec.TextMatrix(0, i)
    Next
    
    Set mrsCompoundGroup = GetCompoundGroup
    vsExec.ColDataType(mCol("���")) = flexDTBoolean
    vsExec.ColHidden(mCol("����ԭ��")) = Not gblnҽ����ֹԭ��
    Call InitExecBar
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If WindowState = 1 Then Exit Sub
    
    vsExec.Top = Me.ScaleTop
    vsExec.Left = Me.ScaleLeft + vsSend.Width + 60
    vsExec.Width = Me.ScaleWidth - vsSend.Width - 60 - picExec.Width
    vsExec.Height = Me.ScaleHeight
    
    fraExecUD.Top = vsExec.Top
    fraExecUD.Left = vsExec.Left - fraExecUD.Width
    fraExecUD.Height = vsExec.Height
      
End Sub

Private Sub fraExecUD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsSend.Width + X < 50 Or vsExec.Width - X < 100 Then Exit Sub
        fraExecUD.Left = fraExecUD.Left + X
                
        vsSend.Width = vsSend.Width + X
        vsExec.Width = vsExec.Width - X
        vsExec.Left = vsExec.Left + X
    End If
End Sub


Private Function SaveData() As Boolean
    Dim colSQL As New Collection, blnTrans As Boolean
    Dim strSQL As String, i As Long, strCurDate As String
    Dim bytTmp As Byte
    Dim lngTmp As Long
    Dim strIDs As String, rsTmp As Recordset
    Dim colMsg As New Collection
    
    strCurDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    
    With vsExec
        For i = .FixedRows To .Rows - 1
            If Val(.Cell(flexcpData, i, mCol("״̬"))) = 1 Then
                lngTmp = .Cell(flexcpChecked, i, mCol("���"))
                If lngTmp <> 1 Then lngTmp = 0
                strIDs = strIDs & "," & .RowData(i)
                
                strSQL = "Zl_��Һ��ҩ��¼_Update(" & .RowData(i) & "," & lngTmp & "," & _
                    IIF(.Cell(flexcpData, i, mCol("��ҩ����")) = "", "Null", Val(.Cell(flexcpData, i, mCol("��ҩ����")))) & ",'" & UserInfo.���� & "'," & strCurDate & ")"
                colSQL.Add strSQL, "C" & colSQL.Count + 1
                
                colMsg.Add .RowData(i) & "," & Val(.Cell(flexcpData, i, mCol("��ҩ����"))), "K" & i
            End If
        Next
    End With
    If colSQL.Count = 0 Then
        RaiseEvent StatusTextUpdate("û�е����κ����Σ�")
    Else
        On Error GoTo errH
        strSQL = "select Count(1) as �Ƿ����� from ��Һ��ҩ��¼ where �Ƿ�����=1 And ID in(Select Column_Value From Table(f_Num2list([1])))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strIDs, 2))
        If rsTmp!�Ƿ����� > 0 Then
            MsgBox "��ǰ��������ҩ��¼�Ѿ�����Һ��ҩ������������ʱ��������е�����", vbInformation, "��Һ��Һ��¼"
            Exit Function
        End If
        gcnOracle.BeginTrans: blnTrans = True
            For i = 1 To colSQL.Count
                Call zlDatabase.ExecuteProcedure(colSQL("C" & i), Me.Caption)
            Next
        gcnOracle.CommitTrans: blnTrans = False
        
        If Not (mclsMipModule Is Nothing) Then
            If mclsMipModule.IsConnect Then
                For i = 1 To colMsg.Count
                    Call ZLHIS_CIS_008(mclsMipModule, mlng����ID, mstr����, mstrסԺ��, , mlng��ҳID, mlng����ID, , mlng����ID, "", , mstr����, mlngAdviceID, mlngҽ����Ч, _
                    Split(colMsg(i), ",")(0), Split(colMsg(i), ",")(1))
                Next
            End If
        End If
        
        RaiseEvent StatusTextUpdate("���ݱ���ɹ���")
        For i = vsExec.FixedRows To vsExec.Rows - 1
            vsExec.Cell(flexcpData, i, mCol("״̬")) = 0
        Next
    End If
        
    SaveData = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub LoadExecList(ByVal lngSendNO As Long)
'���ܣ���ȡ����ʾ��ҩ���μ�¼
'������lngSendNO=ҽ�����ͺ�
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Long, blnDo As Boolean
    Dim rsState As Recordset  '��Һ��ҩ״̬
    Dim strIDs As String, strTmp As String
    Dim lng��Һ����ID As Long
 
    strSQL = "Select ID,����id as ��Һ����ID,To_Char(ִ��ʱ��, 'YYYY-MM-DD HH24:MI') ִ��ʱ��, Nvl(�Ƿ���,0) �Ƿ���, ��ҩ����,ƿǩ��," & vbNewLine & _
            "       Decode(����״̬,1, '����ҩ',2, '����ҩ', 3,'����ҩ', 4,'����ҩ',5,'�ѷ���',6,'��ǩ��',7,'�Ѿܾ�ǩ��',8,'��ȷ�Ͼ���',9,'����������',10,'���������','�ѷ���') As ״̬," & _
            "'' AS ����������,'' as ��������ʱ��,'' as �������ʱ��,����,סԺ��,����,���˿���id" & vbNewLine & _
            "From ��Һ��ҩ��¼" & vbNewLine & _
            "Where ҽ��id = [1] And ���ͺ� = [2] And ����״̬ <> 8" & vbNewLine & _
            "Order By ִ��ʱ��"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngAdviceID, lngSendNO)
    Do While Not rsTmp.EOF
        strIDs = strIDs & "," & rsTmp!ID
        rsTmp.MoveNext
    Loop
    strIDs = Mid(strIDs, 2)
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        mstr���� = rsTmp!���� & ""
        mstrסԺ�� = rsTmp!סԺ�� & ""
        mstr���� = rsTmp!���� & ""
        mlng����ID = Val(rsTmp!���˿���id & "")
        lng��Һ����ID = Val(rsTmp!��Һ����ID & "")
        mrsCompoundGroup.Filter = "��������id=" & lng��Һ����ID
        For i = 1 To mrsCompoundGroup.RecordCount
            strTmp = strTmp & "|" & "#" & mrsCompoundGroup!���� & ";��" & mrsCompoundGroup!���� & "��:" & mrsCompoundGroup!��ҩʱ��
            mrsCompoundGroup.MoveNext
        Next
        strTmp = Mid(strTmp, 2)
        vsExec.ColComboList(mCol("��ҩ����")) = strTmp
    End If
    If strIDs <> "" Then
        strSQL = "Select ��ҩID,��������,������Ա,To_Char(����ʱ��,'YYYY-MM-DD HH24:MI') as ����ʱ��,����˵��,����ʱ�� as ���� from ��Һ��ҩ״̬ Where ��ҩID in(select Column_Value From Table(Cast(f_num2list([1]) As ZLTOOLS.t_numlist)))"
        Set rsState = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIDs)
    End If
    
    With vsExec
        .Redraw = False
        .Rows = .FixedRows
        .Rows = .FixedRows + rsTmp.RecordCount
                        
        For i = .FixedRows To .Rows - 1
            .RowData(i) = Val(rsTmp!ID)
            .TextMatrix(i, mCol("ִ��ʱ��")) = rsTmp!ִ��ʱ��
            .TextMatrix(i, mCol("���")) = rsTmp!�Ƿ���
            .Cell(flexcpData, i, mCol("���")) = Val(rsTmp!�Ƿ���)
            
            If IsNull(rsTmp!��ҩ����) Then
                .TextMatrix(i, mCol("��ҩ����")) = ""
                .Cell(flexcpData, i, mCol("��ҩ����")) = ""
            Else
                .TextMatrix(i, mCol("��ҩ����")) = "��" & rsTmp!��ҩ���� & "��"
                .Cell(flexcpData, i, mCol("��ҩ����")) = Val(rsTmp!��ҩ����)
            End If
            
            .TextMatrix(i, mCol("ƿǩ��")) = "" & rsTmp!ƿǩ��
            .TextMatrix(i, mCol("״̬")) = rsTmp!״̬
            .Cell(flexcpData, i, mCol("״̬")) = 0
            If blnDo = False Then
                If rsTmp!״̬ = "����ҩ" Then blnDo = True
            End If
            If rsTmp!״̬ = "����������" Or rsTmp!״̬ = "���������" Then
                rsState.Filter = "��ҩID=" & rsTmp!ID & " And ��������=9"
                rsState.Sort = "���� desc"
                If rsState.RecordCount > 0 Then
                    rsState.MoveFirst
                    .TextMatrix(i, mCol("����������")) = "" & rsState!������Ա
                    .TextMatrix(i, mCol("��������ʱ��")) = "" & rsState!����ʱ��
                    .TextMatrix(i, mCol("����ԭ��")) = "" & rsState!����˵��
                End If
            End If
            If rsTmp!״̬ = "���������" Then
                rsState.Filter = "��ҩID=" & rsTmp!ID & " And ��������=10"
                If rsState.RecordCount > 0 Then
                    rsState.MoveFirst
                    .TextMatrix(i, mCol("�������ʱ��")) = "" & rsState!����ʱ��
                End If
            End If
            
            If IsNull(rsTmp!��ҩ����) Then
                .TextMatrix(i, mCol("��ҩ����ʱ��")) = ""
            Else
                mrsCompoundGroup.Filter = "����=" & rsTmp!��ҩ���� & " and ��������id=" & lng��Һ����ID
                If mrsCompoundGroup.RecordCount > 0 Then
                    .TextMatrix(i, mCol("��ҩ����ʱ��")) = mrsCompoundGroup!��ҩʱ��
                End If
            End If
            
            .Cell(flexcpBackColor, i, mCol("���")) = COLEditBackColor   'ǳ��
            .Cell(flexcpBackColor, i, mCol("��ҩ����")) = COLEditBackColor
            rsTmp.MoveNext
        Next
        
        .Redraw = True
        If .Rows > .FixedRows Then
            .Row = .Rows - 1
            .TopRow = .Row
        End If
    End With
    
    If blnDo = False Then
        vsExec.Tag = "false"
    Else
        vsExec.Tag = ""
    End If
    mblnEdit = False
    vsExec.Editable = flexEDNone
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadSendList()
'���ܣ���ȡ����ʾҽ�����ͼ�¼
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Long
 
    strSQL = "Select To_Char(����ʱ��, 'YYYY-MM-DD HH24:MI') ����ʱ��, ���ͺ� From ����ҽ������ Where ҽ��id = [1] Order by ����ʱ�� Desc"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngAdviceID)

    With vsSend
        .Redraw = False
        .Rows = .FixedRows
        .Rows = .FixedRows + rsTmp.RecordCount
        For i = 1 To rsTmp.RecordCount
            .TextMatrix(i, 0) = rsTmp!����ʱ��
            .Cell(flexcpData, i, Col����ʱ��) = Val(rsTmp!���ͺ�)
            rsTmp.MoveNext
        Next
        .Redraw = True
        If .Rows > 1 Then .Row = 1
    End With

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetCompoundGroup() As ADODB.Recordset
'���ܣ���ȡ��ҩ��������
    Dim strSQL As String
    
    strSQL = "Select ��������id,����, ��ҩʱ�� From ��ҩ�������� Order By ����"
    On Error GoTo errH
    Set GetCompoundGroup = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set mCol = Nothing
    Set mrsCompoundGroup = Nothing
    Set mclsMipModule = Nothing
End Sub

Private Sub vsExec_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = mCol("��ҩ����") Then
        With vsExec
            'δѡ��ʱ�뿪����
            If .ComboIndex = -1 Or .Cell(flexcpData, Row, Col) = .ComboData Then
                If Val(.Cell(flexcpData, Row, Col)) <> 0 Then .TextMatrix(Row, Col) = "��" & .Cell(flexcpData, Row, Col) & "��"
                Exit Sub
            End If
            
            
            .Cell(flexcpData, Row, Col) = CStr(.ComboData)
            .TextMatrix(Row, Col) = "��" & .ComboData & "��"
            .TextMatrix(Row, mCol("��ҩ����ʱ��")) = Mid(.ComboItem, InStr(.ComboItem, ":") + 1)
            
            .Cell(flexcpData, Row, mCol("״̬")) = 1 '��ʾ�޸Ĺ��ļ�¼
        End With
    ElseIf Col = mCol("���") Then
        With vsExec
            If Val(.Cell(flexcpData, Row, mCol("���"))) <> Val(.TextMatrix(Row, mCol("���"))) Then
                .Cell(flexcpData, Row, mCol("״̬")) = 1 '��ʾ�޸Ĺ��ļ�¼
            End If
        End With
    End If
End Sub

Private Sub vsExec_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
    Call vsExec_KeyPressEdit(Row, Col, 13)  '���Զ�����vsExec_AfterEdit
End Sub

Private Sub vsExec_DblClick()
    Dim objControl As CommandBarControl
    
    If vsExec.Editable = flexEDNone Then
        If (vsExec.TextMatrix(vsExec.Row, mCol("״̬")) = "����ҩ" Or vsExec.MouseCol = mCol("���") And vsExec.TextMatrix(vsExec.Row, mCol("״̬")) = "����ҩ") And vsExec.TextMatrix(vsExec.Row, mCol("����������")) = "" Then
            If MsgBox("��ȷ��Ҫ����" & IIF(mbln����޸�, "��ҩ���λ���", "��ҩ����") & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                
                Set objControl = cbsExec.FindControl(, conMenu_Adjust)
                If Not objControl Is Nothing Then Call cbsExec_Execute(objControl)
                
            End If
        End If
    End If
End Sub

Private Sub vsExec_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        vsExec.Col = Col + 1
    End If
End Sub

Private Sub vsExec_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'˵�����޸Ĵ��ֻ���� ����ҩ�����ҩ ����״̬���ܲ��� ������� ���ơ�
'                       ����ҩ ״̬�޸Ĵ��ʱ���ܲ��� ��Һ����ҩ���ٴ�������ı���״̬ ���ơ�
    If (Col = mCol("���") And mbln����޸� Or Col = mCol("��ҩ����")) Then
        Cancel = (vsExec.TextMatrix(Row, mCol("״̬")) <> "����ҩ")
        If Col = mCol("���") Then
            If vsExec.TextMatrix(Row, mCol("״̬")) = "����ҩ" Then
                If mbln��ҩ���ܸ�״̬ Then
                    Cancel = True
                    MsgBox "����������Ѿ���ҩ����Һ���Ĵ��״̬��", vbInformation, Me.Caption
                    Exit Sub
                Else
                    Cancel = False
                End If
            End If
        End If
        If Cancel = False Then
            If vsExec.TextMatrix(Row, mCol("����������")) <> "" Then
                Cancel = True
                MsgBox "�Ѿ��������ʵļ�¼�������޸ġ�", vbInformation, Me.Caption
            End If
            '�ж�Ȩ��
            If Not Cancel Then
                If Col = mCol("���") Then
                    If InStr(GetInsidePrivs(pסԺҽ������), ";�޸���Һ���״̬;") = 0 Then
                        Cancel = True
                        MsgBox "��û���޸Ĵ��״̬Ȩ�ޣ����ܽ�������", vbInformation, Me.Caption
                    End If
                End If
                
                If Col = mCol("��ҩ����") Then
                    If InStr(GetInsidePrivs(pסԺҽ������), ";�޸���Һ����;") = 0 Then
                        Cancel = True
                        MsgBox "��û���޸���ҩ����Ȩ�ޣ����ܽ�������", vbInformation, Me.Caption
                    End If
                End If
            End If
        Else
            If Col = mCol("���") Then
                MsgBox "ֻ�д���ҩ�����ҩ�ļ�¼�ɴ����", vbInformation, Me.Caption
            Else
                MsgBox "ֻ�д���ҩ�ļ�¼�ɵ������Ρ�", vbInformation, Me.Caption
            End If
        End If
    Else
        Cancel = True
    End If
End Sub

Private Sub vsSend_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow And Me.Visible = True Then
        If Val(vsSend.Cell(flexcpData, NewRow, Col����ʱ��)) <> 0 Then
            Call LoadExecList(Val(vsSend.Cell(flexcpData, NewRow, Col����ʱ��)))
        End If
    End If
End Sub


Private Sub InitExecBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim strPrivs As String

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsExec.VisualTheme = xtpThemeOfficeXP  'ʹ��2003���ʱ����ť��ͻ��Ч����ֻ��һ����ťʱ���ÿ�
    With Me.cbsExec.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .UseFadedIcons = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
    End With
    Set cbsExec.Icons = zlCommFun.GetPubIcons
    cbsExec.EnableCustomization False
    cbsExec.ActiveMenuBar.Visible = False
        
    Set objBar = cbsExec.Add("������", xtpBarTop)
    objBar.EnableDocking xtpFlagStretched   '��Ȳ���ʱ�Զ�����
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = True
    objBar.SetIconSize 24, 24
    
    strPrivs = GetInsidePrivs(pסԺ���ʲ���)
            
    With objBar.Controls
        If InStr(strPrivs, ";ҩƷ��������;") > 0 Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ChargeDelApply, "��������")
            objControl.IconId = 3821
        End If
        Set objControl = .Add(xtpControlButton, conMenu_AdjustCancle, "ȡ������")
        objControl.BeginGroup = True
        objControl.ToolTipText = "ȡ����������"
        objControl.IconId = conMenu_Edit_Untread
        Set objControl = .Add(xtpControlButton, conMenu_Adjust, "��������")
        objControl.BeginGroup = True
        objControl.IconId = 3564
        Set objControl = .Add(xtpControlButton, conMenu_Save, "����")
        objControl.Visible = False
        objControl.IconId = 3503
        Set objControl = .Add(xtpControlButton, conMenu_Undo, "����")
        objControl.Visible = False
        objControl.IconId = 3014
        
        objControl.BeginGroup = True
    End With
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    picExec.BackColor = cbsExec.GetSpecialColor(STDCOLOR_BTNFACE)
End Sub

Private Sub cbsExec_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
            
    Select Case Control.ID
        Case conMenu_Adjust    '��������
            mblnEdit = True
            
            vsExec.Editable = flexEDKbdMouse
            vsSend.Enabled = False
            vsExec.SetFocus
            RaiseEvent SetEditState(True)
            
        Case conMenu_Save '����
            
            If SaveData Then
                mblnEdit = False
                
                vsExec.Editable = flexEDNone
                vsSend.Enabled = True
                RaiseEvent SetEditState(False)
                Call LoadExecList(Val(vsSend.Cell(flexcpData, vsSend.Row, Col����ʱ��)))
            End If
            
        Case conMenu_Undo  '����
            mblnEdit = False
            
            vsExec.Editable = flexEDNone
            vsSend.Enabled = True
            RaiseEvent SetEditState(False)
        
            Call LoadExecList(Val(vsSend.Cell(flexcpData, vsSend.Row, Col����ʱ��)))
        Case conMenu_Edit_ChargeDelApply    '��������
            Call ExecChargeDelApply(Control.Caption = "����")
        Case conMenu_AdjustCancle    'ȡ����������
            Call ExecCancleChargeDelApply
    End Select
    
End Sub

Private Sub cbsExec_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim str״̬ As String
    Dim blnVisible As Boolean
     
    Select Case Control.ID
        Case conMenu_Adjust    '��������
            Control.Enabled = Not mblnEdit And vsExec.Tag = "" And vsExec.TextMatrix(vsExec.Row, mCol("����������")) = ""
            Control.Visible = Not mblnEdit
        Case conMenu_Save, conMenu_Undo '����
            Control.Visible = mblnEdit
        Case conMenu_Edit_ChargeDelApply
            '�����ó�false������������Ʋ����ó�true����Ϊcommandbar��bug�������������������ȥ�Ż��������
            Control.Visible = False
            If vsExec.Row >= vsExec.FixedRows Then
                str״̬ = vsExec.TextMatrix(vsExec.Row, mCol("״̬"))
                Control.Enabled = False
                If Not mblnEdit Then
                    blnVisible = True
                    If vsExec.TextMatrix(vsExec.Row, mCol("����������")) = "" Then
                        If str״̬ = "����ҩ" Or str״̬ = "����ҩ" Then
                            Control.Enabled = True
                        ElseIf Not (str״̬ = "����ҩ" Or str״̬ = "����ҩ") Then
                            If vsExec.Cell(flexcpChecked, vsExec.Row, mCol("���")) = 1 Then
                                Control.Enabled = True
                            Else
                                If mbln�������� Then Control.Enabled = True
                            End If
                        End If
                    ElseIf vsExec.TextMatrix(vsExec.Row, mCol("����������")) <> "" And vsExec.TextMatrix(vsExec.Row, mCol("�������ʱ��")) = "" Then
                        blnVisible = False
                    End If
                Else
                    blnVisible = False
                End If
                If str״̬ = "����ҩ" And Control.Enabled Then
                    Control.Caption = "����"
                    Control.ToolTipText = "����"
                Else
                    Control.Caption = "��������"
                    Control.ToolTipText = "��������"
                End If
            Else
                Control.Enabled = False
                Control.Caption = "��������"
                Control.ToolTipText = "��������"
            End If
            Control.Visible = blnVisible
        Case conMenu_AdjustCancle
            If vsExec.Row >= vsExec.FixedRows Then
                If Not mblnEdit And vsExec.TextMatrix(vsExec.Row, mCol("����������")) <> "" And vsExec.TextMatrix(vsExec.Row, mCol("�������ʱ��")) = "" Then
                    Control.Visible = True
                Else
                    Control.Visible = False
                End If
            Else
                Control.Visible = False
            End If
    End Select
End Sub


Private Sub ExecChargeDelApply(Optional ByVal blnAutoAduit As Boolean)
'���ܣ�ִ����������
'������blnAutoAduit=true�Զ������������
    Dim lng��ҩ��¼ As Long, strDate As String, strNote As String
    Dim rsTmp As ADODB.Recordset, strSQL As String, i As Long
    Dim colSQL As New Collection, blnTrans As Boolean
    Dim strProg As String
    Dim strTmp As String
    Dim lng��˲���ID As Long
    Dim str����ԭ�� As String
    Dim strTab As String
    
    If InStr(GetInsidePrivs(pסԺ���ʲ���), ";ҩƷ��������;") = 0 Then
        MsgBox "��û��סԺ���˲���ģ���е�ҩƷ��������Ȩ�ޣ����ܽ����������ʡ�"
        Exit Sub
        
    End If
    If blnAutoAduit Then
        strProg = "����"
        If InStr(GetInsidePrivs(pסԺ���ʲ���), ";�������;") = 0 Then
            MsgBox "��û��סԺ���˲���ģ���е�ҩƷ�������Ȩ�ޣ����ܽ������ʡ�"
            Exit Sub
        End If
    Else
        strProg = "��������"
    End If
    
    With vsExec
        If .TextMatrix(.Row, mCol("��ҩ����")) = "" Then
            MsgBox "��ǰѡ���˵�" & .Row & "��,��ҩ����Ϊ�գ����ܽ���" & strProg & "��", vbInformation, gstrSysName
            Exit Sub
        End If
        On Error GoTo errH
        strSQL = "select �Ƿ�����,����״̬,�Ƿ���,����ID from ��Һ��ҩ��¼ where ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.RowData(.Row)))
        lng��˲���ID = Val(rsTmp!����ID & "")
        If Val(rsTmp!�Ƿ����� & "") = 1 Then
            MsgBox "��ǰ���ʵ���ҩ��¼�Ѿ�����Һ��ҩ������������ʱ������������ʡ�", vbInformation, "��Һ��Һ��¼"
            Exit Sub
        End If
        If Val(rsTmp!����״̬ & "") = 9 Or Val(rsTmp!����״̬ & "") = 10 Then
            MsgBox "��ǰ��ҩ��¼�Ѿ���" & strProg & "�����顣", vbInformation, "��Һ��Һ��¼"
            Call LoadExecList(Val(vsSend.Cell(flexcpData, vsSend.Row, Col����ʱ��)))
            Exit Sub
        End If
        If Val(rsTmp!����״̬ & "") >= 4 Then
            If Val(rsTmp!�Ƿ��� & "") = 0 And mbln�������� = False Then
                MsgBox "��ǰδ����ļ�¼�Ѿ���ҩ�����������������롣", vbInformation, "��Һ��Һ��¼"
                Call LoadExecList(Val(vsSend.Cell(flexcpData, vsSend.Row, Col����ʱ��)))
                Exit Sub
            End If
        End If
	If mlngҽ����Ч = 0 And mlng�������� <> 1 Then
            strTab = "סԺ���ü�¼"
        Else
            If GetAdviceFeeKind(mlngAdviceID) = 2 Then    'סԺҽ��վ�������ɷ��͵�����
                strTab = "סԺ���ü�¼"
            Else
                strTab = "������ü�¼"
            End If
        End If
        '77686,���ϴ�,2014/9/18,�����������
        strSQL = "Select b.����id, b.ҩƷid As �շ�ϸĿid, Sum(a.����) As ����, c.סԺ��װ, c.סԺ��λ, d.����, b.No, e.���,E.��¼״̬" & vbNewLine & _
            "From ��Һ��ҩ���� A, ҩƷ�շ���¼ B, " & strTab & " E, ҩƷ��� C, �շ���ĿĿ¼ D" & vbNewLine & _
            "Where a.��¼id = [1] And a.�շ�id = b.Id And b.����id = e.Id And b.ҩƷid = c.ҩƷid And c.ҩƷid = d.Id" & vbNewLine & _
            IIF(Not blnAutoAduit, " And b.����� is Not null", "") & vbNewLine & _
            " And instr( ',8,9,10,21,24,25,26,',','||B.����||',')>0 " & _
            "Group By b.����id, b.ҩƷid, c.סԺ��װ, c.סԺ��λ, d.����, b.No, e.���,E.��¼״̬" & vbNewLine & _
            "Order By e.���"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.RowData(.Row)))
        If rsTmp.RecordCount = 0 Then
            MsgBox "��ǰѡ���˵�" & .Row & "�У�û���ҵ���Ӧ����ҩ���ݣ����ܽ���" & strProg & "��", vbInformation, gstrSysName
            Exit Sub
        Else
            '����ҩ�Ķ����ѷ�ҩ�ģ���������Զ���˰��ѷ�ҩ������������
            'һ����ҩ���Σ�һ����������ҩƷ
            For i = 1 To rsTmp.RecordCount
                If Val(rsTmp!��¼״̬ & "") = 0 Then
                    MsgBox "��ǰ��ҩ��¼�ǻ��۵�������������ʡ�", vbInformation, gstrSysName
                    Exit Sub
                End If
                strNote = strNote & vbCrLf & rsTmp!���� & "��" & FormatEx(rsTmp!���� / rsTmp!סԺ��װ, 5) & rsTmp!סԺ��λ
                rsTmp.MoveNext
            Next
            If gblnҽ����ֹԭ�� Then
                Call frmAdviceStopTime.ShowMe(mfrmParent, mlngAdviceID, mlng����ID, 2, , str����ԭ��)
                If str����ԭ�� = "" Then Exit Sub
            End If
            If MsgBox("��ǰѡ���˵�" & .Row & "����ҩ���Σ�" & strNote & vbCrLf & "��ȷ��Ҫ����ЩҩƷ" & strProg & "��", vbYesNo + vbQuestion, gstrSysName) = vbNo Then
                Exit Sub
            End If
            
            rsTmp.MoveFirst
            
            strTmp = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            strDate = "To_Date('" & strTmp & "','YYYY-MM-DD HH24:MI:SS')"
            For i = 1 To rsTmp.RecordCount
                strSQL = "Zl_���˷�������_Insert(" & rsTmp!����ID & "," & rsTmp!�շ�ϸĿID & "," & _
                      mlng����ID & "," & rsTmp!���� & ",'" & UserInfo.���� & "'," & strDate & "," & IIF(blnAutoAduit, "0", "1") & ",1," & Val(.RowData(.Row)) & ",'" & str����ԭ�� & "')"
                colSQL.Add strSQL, "C" & colSQL.Count + 1
                '����ҩ״̬���Զ����
                If blnAutoAduit Then
                    strSQL = "Zl_���˷�������_Audit(" & rsTmp!����ID & "," & strDate & ",'" & _
                          UserInfo.���� & "'," & strDate & ",1,1,0" & ")"
                    colSQL.Add strSQL, "C" & colSQL.Count + 1
                    If strTab = "������ü�¼" Then
                        strSQL = "Zl_������ʼ�¼_Delete('" & rsTmp!NO & "', '" & rsTmp!��� & ":" & rsTmp!���� & ":" & Val(.RowData(.Row)) & "', '" & UserInfo.��� & "', '" & UserInfo.���� & "', 0)"
                    Else
                        strSQL = "Zl_סԺ���ʼ�¼_Delete('" & rsTmp!NO & "', '" & rsTmp!��� & ":" & rsTmp!���� & ":" & Val(.RowData(.Row)) & "', '" & UserInfo.��� & "', '" & UserInfo.���� & "', 2)"
                    End If
                    colSQL.Add strSQL, "C" & colSQL.Count + 1
                End If
                rsTmp.MoveNext
            Next
            
            On Error GoTo errH
            gcnOracle.BeginTrans: blnTrans = True
                For i = 1 To colSQL.Count
                    Call zlDatabase.ExecuteProcedure(colSQL("C" & i), Me.Caption)
                Next
            gcnOracle.CommitTrans: blnTrans = False
            'ZLHIS_CIS_013-סԺ������Һ��������
            If Not (mclsMipModule Is Nothing) Then
                If mclsMipModule.IsConnect Then
                    Call ZLHIS_CIS_013(mclsMipModule, mlng����ID, mstr����, mstrסԺ��, mlng��ҳID, mlng����ID, , mlng����ID, , mlngAdviceID, Val(vsExec.RowData(vsExec.Row)), strTmp, UserInfo.����, mlng����ID, , lng��˲���ID)
                End If
            End If
            
            MsgBox strProg & "�����ɹ���", vbInformation, gstrSysName
            i = .Row
            Call vsSend_AfterRowColChange(0, 0, vsSend.Row, vsSend.Col)
            If .Rows > i Then
                .Row = i
            End If
        End If
    End With
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub ExecCancleChargeDelApply()
'���ܣ�ȡ����������
    Dim lng��ҩ��¼ As Long, strDate As String, strNote As String
    Dim rsTmp As ADODB.Recordset, strSQL As String, i As Long
    Dim colSQL As New Collection, blnTrans As Boolean
    Dim strProg As String
    Dim strTmp As String
    
    If InStr(GetInsidePrivs(pסԺ���ʲ���), ";ҩƷ��������;") = 0 Then
        MsgBox "��û��סԺ���˲���ģ���е�ҩƷ��������Ȩ�ޣ����ܽ���ȡ���������ʡ�"
        Exit Sub
    End If
    strProg = "ȡ����������"
    With vsExec
        If .TextMatrix(.Row, mCol("��ҩ����")) = "" Then
            MsgBox "��ǰѡ���˵�" & .Row & "��,��ҩ����Ϊ�գ����ܽ���" & strProg & "��", vbInformation, gstrSysName
            Exit Sub
        End If
        On Error GoTo errH
        strSQL = "select Count(1) as �Ƿ����� from ��Һ��ҩ��¼ where �Ƿ�����=1 And ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.RowData(.Row)))
        If rsTmp!�Ƿ����� > 0 Then
            MsgBox "��ǰ���ʵ���ҩ��¼�Ѿ�����Һ��ҩ������������ʱ���������ȡ���������롣", vbInformation, "��Һ��Һ��¼"
            Exit Sub
        End If
        '77686,���ϴ�,2014/9/18,�����������
        strSQL = "Select distinct c.����id" & vbNewLine & _
                "From ��Һ��ҩ���� A, ҩƷ�շ���¼ B, ���˷������� C" & vbNewLine & _
                "Where a.��¼id = [1] And a.�շ�id = b.Id And b.����id = c.����id And c.����� Is Null " & _
                "And instr( ',8,9,10,21,24,25,26,',','||B.����||',')>0"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.RowData(.Row)))
        If rsTmp.RecordCount = 0 Then
            MsgBox "��ǰѡ���˵�" & .Row & "�У�û���ҵ���Ӧ�����������¼�����ܽ���" & strProg & "��", vbInformation, gstrSysName
            Exit Sub
        Else

            For i = 1 To rsTmp.RecordCount
                strTmp = strTmp & "," & rsTmp!����ID
                rsTmp.MoveNext
            Next
            strTmp = Mid(strTmp, 2)
            
            strSql = "Zl_���˷�������_Delete('" & strTmp & "'," & Val(.RowData(.Row)) & ")"
            colSQL.Add strSQL, "C" & colSQL.Count + 1

            If MsgBox("��ǰѡ���˵�" & .Row & "����ҩ��¼," & vbCrLf & "��ȷ��Ҫ����ЩҩƷ" & strProg & "��", vbYesNo + vbQuestion, gstrSysName) = vbNo Then
                Exit Sub
            End If
            
            On Error GoTo errH
            gcnOracle.BeginTrans: blnTrans = True
                For i = 1 To colSQL.Count
                    Call zlDatabase.ExecuteProcedure(colSQL("C" & i), Me.Caption)
                Next
            gcnOracle.CommitTrans: blnTrans = False
            
            MsgBox strProg & "�����ɹ���", vbInformation, gstrSysName
            i = .Row
            Call vsSend_AfterRowColChange(0, 0, vsSend.Row, vsSend.Col)
            If .Rows > i Then
                .Row = i
            End If
        End If
    End With
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

