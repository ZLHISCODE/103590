VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmNumSortSel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�ű�ѡ��"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6300
      TabIndex        =   2
      Top             =   615
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6300
      TabIndex        =   1
      Top             =   135
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshPlan 
      Height          =   5715
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   10081
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      RowHeightMin    =   250
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   1
      MousePointer    =   1
      FormatString    =   "^  �ű�|^    ����|^      ��Ŀ|^  ҽ��|ʱ���|�޺�|�ѹ�"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmNumSort.frx":0000
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frmNumSortSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const STR_COMP = "|',~" '�ָ��ַ���
Private mrsPlan As New ADODB.Recordset
Private mlngSect As Long
Private mlngID As Long
Private strSQL As String
Private mstrReturn As String
Private mblnOK As Boolean
Private i As Long
Private mstr�ű� As String

Public Function ShowMe(ByVal lng�Һ�ID As String, strReturn As String, frmParent As Form) As Boolean
'��ʾ�����岢����ѡ����Ƿ���ȷ
On Error GoTo errHandle

    mblnOK = False
    '���ҵ�ִ�п��Һͺű�
    strSQL = "Select B.�ű�,B.ִ�в���ID,A.�շ�ϸĿID " & _
        " From ������ü�¼ A,���˹Һż�¼ B" & _
        " Where A.��¼����=4 and A.��¼״̬=1 And A.���=1 And b.��¼����=1 and b.��¼״̬=1 and A.NO=B.NO And B.ID=[1]"
    Set mrsPlan = zlDatabase.OpenSQLRecord(strSQL, "�ű�ѡ����", lng�Һ�ID)
    
    If mrsPlan.RecordCount > 0 Then
        mrsPlan.MoveFirst
        mlngSect = mrsPlan!ִ�в���id
        mlngID = mrsPlan!�շ�ϸĿID
        mstr�ű� = mrsPlan!�ű�
    Else
        Exit Function
    End If
    
    Me.Show 1, frmParent
    '�ű�ID,��ĿID,ҽ��ID,ҽ��,����ID,����,����,�ű�
    If Not mblnOK Then
        strReturn = ",,,,,,,"
    Else
        strReturn = mstrReturn
        ShowMe = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetPlanGrid()
    Dim i As Integer
    
    '��ʼ���ű�
    With mshPlan
        .Redraw = False
        .Clear: .Rows = 2: .Cols = 17
        .TextMatrix(0, 0) = "IDS" '�ű�ID_��ĿID_ҽ��ID
        .TextMatrix(0, 1) = "����"
        .TextMatrix(0, 2) = "�ű�"
        .TextMatrix(0, 3) = "����"
        .TextMatrix(0, 4) = "��Ŀ"
        .TextMatrix(0, 5) = "ҽ��"
        .TextMatrix(0, 6) = "�޺�"
        .TextMatrix(0, 7) = "�ѹ�"
        .TextMatrix(0, 8) = "��"
        .TextMatrix(0, 9) = "һ"
        .TextMatrix(0, 10) = "��"
        .TextMatrix(0, 11) = "��"
        .TextMatrix(0, 12) = "��"
        .TextMatrix(0, 13) = "��"
        .TextMatrix(0, 14) = "��"
        .TextMatrix(0, 15) = "����"
        .TextMatrix(0, 16) = "����"
        
        If Not Visible Then
            .ColWidth(0) = 0
            .ColWidth(1) = 500
            .ColWidth(2) = 550
            .ColWidth(3) = 1150
            .ColWidth(4) = 1250
            .ColWidth(5) = 700
            .ColWidth(6) = 500
            .ColWidth(7) = 500
            .ColWidth(8) = 280
            .ColWidth(9) = 280
            .ColWidth(10) = 280
            .ColWidth(11) = 280
            .ColWidth(12) = 280
            .ColWidth(13) = 280
            .ColWidth(14) = 280
            .ColWidth(15) = 500
            .ColWidth(16) = 500
        End If
        
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 1
        .ColAlignment(5) = 1
        .ColAlignment(6) = 1
        .ColAlignment(7) = 1
        .ColAlignment(8) = 4
        .ColAlignment(9) = 4
        .ColAlignment(10) = 4
        .ColAlignment(11) = 4
        .ColAlignment(12) = 4
        .ColAlignment(13) = 4
        .ColAlignment(14) = 4
        .ColAlignment(15) = 4
        .ColAlignment(16) = 4
        
        If Not Visible Then Call RestoreFlexState(mshPlan, App.ProductName & "\" & Me.Name)
        
        For i = 0 To .Cols - 1
            .ColAlignmentFixed(i) = flexAlignCenterCenter
        Next
        
        .RowHeight(0) = 300
        
        .Redraw = True
    End With
End Sub

Private Function ShowPlans(Optional strSort As String = "�ű�", Optional blnDesc As Boolean) As Boolean
'���ܣ���ȡ���հ�������
    Dim i As Integer
    Dim strTime As String, strState As String
    
    On Error GoTo errH
    '�ò�����䵱ʱ��ȡ���ְ��ŵĹҺ����
    strState = _
        "Select A.ID as ����ID,B.����,B.�ѹ���" & vbCrLf & _
        " From �ҺŰ��� A,���˹ҺŻ��� B" & vbCrLf & _
        " Where a.����ID = B.����ID And a.��ĿID = B.��ĿID" & vbCrLf & _
        " And Nvl(A.ҽ��ID,0)=Nvl(B.ҽ��ID,0)" & vbCrLf & _
        " And Nvl(A.ҽ������,'ҽ��')=Nvl(B.ҽ������,'ҽ��') And (a.����=b.���� or b.���� is null )" & vbCrLf & _
        " And B.����=Trunc(Sysdate)"
    '�ò������ȡ��ʱ����Ӧ��ʱ���
    strTime = _
        "Select ʱ��� From ʱ��� Where" & vbCrLf & _
        " ('3000-01-10 '||To_Char(SysDate,'HH24:MI:SS')" & vbCrLf & _
        " Between" & vbCrLf & _
        " Decode(Sign(��ʼʱ�� - ��ֹʱ��),1,'3000-01-09 '||To_Char(��ʼʱ��,'HH24:MI:SS'),'3000-01-10 '||To_Char(��ʼʱ��,'HH24:MI:SS'))" & vbCrLf & _
        " And" & vbCrLf & _
        " '3000-01-10 '||To_Char(��ֹʱ��,'HH24:MI:SS'))" & vbCrLf & _
        " Or" & vbCrLf & _
        " ('3000-01-10 '||To_Char(SysDate,'HH24:MI:SS')" & vbCrLf & _
        " Between" & vbCrLf & _
        " '3000-01-10 '||To_Char(��ʼʱ��,'HH24:MI:SS')" & vbCrLf & _
        " And" & vbCrLf & _
        " Decode(Sign(��ʼʱ�� - ��ֹʱ��),1,'3000-01-11 '||To_Char(��ֹʱ��,'HH24:MI:SS'),'3000-01-10 '||To_Char(��ֹʱ��,'HH24:MI:SS')))"
    '�ò������ȡʱ���ڵİ��ż�״̬
    strSQL = _
        "Select Distinct P.ID,A.����,P.���� as �ű�,P.����," & vbCrLf & _
        " Decode(To_Char(SysDate,'D'),'1',P.����,'2',P.��һ,'3',P.�ܶ�,'4',P.����,'5',P.����,'6',P.����,'7',P.����) as ʱ��," & vbCrLf & _
        " P.����ID,B.���� As ����,P.��ĿID,C.���� As ��Ŀ,P.ҽ��ID,P.ҽ������ as ҽ��,Nvl(A.�ѹ���,0) as �ѹ�," & vbCrLf & _
        " F.�޺��� as �޺�,F.��Լ�� as ��Լ,Nvl(P.��������,0) as ����,Nvl(C.��Ŀ����,0) as ����," & vbCrLf & _
        " P.���� as ��,P.��һ as һ,P.�ܶ� as ��,P.���� as ��,P.���� as ��,P.���� as ��,P.���� as ��," & vbCrLf & _
        " Decode(P.���﷽ʽ,1,'ָ��',2,'��̬',3,'ƽ��',NULL) as ����" & vbCrLf & _
        " From �ҺŰ��� P,(" & strState & ") A,���ű� B,�շ���ĿĿ¼ C,ʱ��� X,�ҺŰ������� F " & vbCrLf & _
        " Where P.ID=A.����ID(+) And P.ͣ������ Is Null And P.����ID=B.ID And P.��ĿID=C.ID AND P.����id <>[1] and P.��ĿID=[2]" & vbCrLf & _
        " And P.Id = f.����id(+) And Not Exists (Select 1 From �ҺŰ���ͣ��״̬ Where Sysdate Between ��ʼֹͣʱ�� And ����ֹͣʱ�� And ����ID=P.ID) And" & _
        " Decode(To_Char(sysdate, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null) =f.������Ŀ(+)" & vbNewLine & _
        " And SysDate Between C.����ʱ�� And Nvl(C.����ʱ��,To_Date('3000-01-01','YYYY-MM-DD')) " & vbCrLf & _
        " And Decode(To_Char(SysDate,'D'),'1',P.����,'2',P.��һ,'3',P.�ܶ�,'4',P.����,'5',P.����,'6',P.����,'7',P.����,NULL) IN" & vbCrLf & _
        " (" & strTime & ") Order by " & strSort & IIf(blnDesc, " Desc", "") & IIf(strSort <> "�ű�", ",�ű�", "")
    
    Set mrsPlan = zlDatabase.OpenSQLRecord(strSQL, "�ű�ѡ����", mlngSect, mlngID)
    If mrsPlan.RecordCount > 0 Then
        mrsPlan.MoveFirst
        mshPlan.Rows = mrsPlan.RecordCount + 1
        For i = 1 To mrsPlan.RecordCount
            mshPlan.RowData(i) = mrsPlan!����ID
            mshPlan.TextMatrix(i, 0) = mrsPlan!ID & "," & mrsPlan!��ĿID & "," & IIf(IsNull(mrsPlan!ҽ��ID), 0, mrsPlan!ҽ��ID)
            mshPlan.TextMatrix(i, 1) = IIf(IsNull(mrsPlan!����), "", mrsPlan!����)
            mshPlan.TextMatrix(i, 2) = mrsPlan!�ű�
            mshPlan.TextMatrix(i, 3) = mrsPlan!����
            mshPlan.TextMatrix(i, 4) = mrsPlan!��Ŀ
            mshPlan.TextMatrix(i, 5) = IIf(IsNull(mrsPlan!ҽ��), "", mrsPlan!ҽ��)
            mshPlan.TextMatrix(i, 6) = IIf(IsNull(mrsPlan!�޺�), "", mrsPlan!�޺�)
            mshPlan.TextMatrix(i, 7) = IIf(mrsPlan!�ѹ� = 0, "", mrsPlan!�ѹ�)
            mshPlan.TextMatrix(i, 8) = Left(IIf(IsNull(mrsPlan!��), "", mrsPlan!��), 1)
            mshPlan.TextMatrix(i, 9) = Left(IIf(IsNull(mrsPlan!һ), "", mrsPlan!һ), 1)
            mshPlan.TextMatrix(i, 10) = Left(IIf(IsNull(mrsPlan!��), "", mrsPlan!��), 1)
            mshPlan.TextMatrix(i, 11) = Left(IIf(IsNull(mrsPlan!��), "", mrsPlan!��), 1)
            mshPlan.TextMatrix(i, 12) = Left(IIf(IsNull(mrsPlan!��), "", mrsPlan!��), 1)
            mshPlan.TextMatrix(i, 13) = Left(IIf(IsNull(mrsPlan!��), "", mrsPlan!��), 1)
            mshPlan.TextMatrix(i, 14) = Left(IIf(IsNull(mrsPlan!��), "", mrsPlan!��), 1)
            mshPlan.TextMatrix(i, 15) = IIf(mrsPlan!���� = 1, "��", "")
            mshPlan.TextMatrix(i, 16) = IIf(IsNull(mrsPlan!����), "", mrsPlan!����)
            mrsPlan.MoveNext
        Next
    Else
        Set mrsPlan = Nothing
        Call SetPlanGrid
    End If
    
    mshPlan.Col = 0: mshPlan.ColSel = mshPlan.Cols - 1
    Call mshPlan_EnterCell
    
    ShowPlans = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set mrsPlan = Nothing
End Function

Private Sub cmdOK_Click()
    mshPlan_DblClick
End Sub

Private Sub mshPlan_DblClick()
    If mshPlan.Row > 0 Then
        If mshPlan.TextMatrix(mshPlan.Row, 0) = "" Then
            MsgBox "û���ʺϻ��ŵĺű�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '�ű�ID,��ĿID,ҽ��ID,ҽ��,����ID,����,����,�ű�
        mstrReturn = mshPlan.TextMatrix(mshPlan.Row, 0) & "," & mshPlan.TextMatrix(mshPlan.Row, 5) & "," & mshPlan.RowData(mshPlan.Row) & "," & mshPlan.TextMatrix(mshPlan.Row, 3) & "," & mshPlan.TextMatrix(mshPlan.Row, 1) & "," & mshPlan.TextMatrix(mshPlan.Row, 2)
        mblnOK = True
        Unload Me
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    SetPlanGrid
    ShowPlans
End Sub

Private Sub mshPlan_EnterCell()
    Dim i As Integer, blnPre As Boolean
    Dim intRow As Integer, intCol As Integer
    
    blnPre = mshPlan.Redraw
    intRow = mshPlan.Row: intCol = mshPlan.Col
    mshPlan.Redraw = False
    
    For i = 0 To mshPlan.Cols - 1
        mshPlan.Col = i
        mshPlan.CellBackColor = mshPlan.BackColorSel
        mshPlan.CellForeColor = mshPlan.ForeColorSel
    Next
    
    mshPlan.Row = intRow:  mshPlan.Col = intCol
    mshPlan.Redraw = blnPre
End Sub

Private Sub mshPlan_LeaveCell()
    Dim i As Integer, blnPre As Boolean
    
    blnPre = mshPlan.Redraw
    mshPlan.Redraw = False
    
    For i = 0 To mshPlan.Cols - 1
        mshPlan.Col = i
        mshPlan.CellBackColor = mshPlan.BackColor
        mshPlan.CellForeColor = mshPlan.ForeColor
    Next
    
    mshPlan.Redraw = blnPre
End Sub

Private Sub mshPlan_SelChange()
    If mshPlan.Rows = 2 Then Exit Sub
    mshPlan.RowSel = mshPlan.Row
End Sub
