VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBookingDeferNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ԤԼ����"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
   Icon            =   "frmBookingDeferNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   8085
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6840
      TabIndex        =   6
      Top             =   90
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6840
      TabIndex        =   5
      Top             =   570
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   3900
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   6879
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
      FormatString    =   "^ ����|^     ʱ��|^       NO|^    Ʊ�ݺ�|^     ����|^ �Ա�|^ ����|^     �����"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmBookingDeferNew.frx":038A
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSComCtl2.DTPicker dtpDefer 
      Height          =   300
      Left            =   3720
      TabIndex        =   1
      Top             =   2490
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      _Version        =   393216
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   96468995
      CurrentDate     =   36588
   End
   Begin MSComCtl2.DTPicker dtpBooking 
      Height          =   300
      Left            =   1200
      TabIndex        =   2
      Top             =   2490
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      _Version        =   393216
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   96468995
      CurrentDate     =   36588
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshPlan 
      Height          =   2235
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   3942
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
      FormatString    =   "^  ����|  �ű�|^       ����|^    ҽ��|��Լ|��ʼʱ��|��ֹʱ��|��ſ���"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmBookingDeferNew.frx":06A4
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lblNewDate 
      Caption         =   "������"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblOldDate 
      Caption         =   "ԤԼ����"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   735
   End
End
Attribute VB_Name = "frmBookingDeferNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsList As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strBegin As String, strEnd As String, strTmp As String
    Dim strSNS As String, strDay As String, str�ű� As String
    Dim i As Long, intDay As Integer, intCol As Integer
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim lng��¼ID As Long
    
    If mrsList Is Nothing Then Exit Sub
    If mrsList.RecordCount = 0 Then Exit Sub
    str�ű� = mshPlan.TextMatrix(mshPlan.Row, GetPlanCol("�ű�"))
    lng��¼ID = Val(mshPlan.TextMatrix(mshPlan.Row, GetPlanCol("��¼ID")))
    If str�ű� = "" Then Exit Sub
    If dtpDefer.Value <= dtpBooking.Value Then
        MsgBox "ָ��������ʱ�������ھɵ�ԤԼʱ��!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '���úű����Ƿ��а���
    mrsList.Sort = "ʱ�� ASC"
    mrsList.MoveLast
    strBegin = mrsList!ʱ��
    
    mrsList.Sort = "ʱ�� DESC"
    mrsList.MoveFirst
    strEnd = mrsList!ʱ��
    
    intDay = Weekday(dtpDefer.Value, vbSunday)
    strDay = Choose(intDay, "����", "��һ", "�ܶ�", "����", "����", "����", "����")
    
    On Error GoTo errH
    strSQL = "Select ��ʼʱ��,��ֹʱ��" & vbNewLine & _
            "From (Select B.��ʼʱ��, Decode(Sign(B.��ֹʱ�� - B.��ʼʱ��), 1, B.��ֹʱ��, B.��ֹʱ�� + 1) ��ֹʱ��" & vbNewLine & _
            "       From �ٴ������¼ A, ʱ��� B, �ٴ������Դ C" & vbNewLine & _
            "       Where C.���� = [1] And A.��ԴID = C.ID And A." & strDay & " = B.ʱ��� And ([2] Between A.��ʼʱ�� And A.��ֹʱ�� Or A.��ʼʱ�� IS Null))" & vbNewLine & _
            "Where To_Date(To_char(��ʼʱ��,'yyyy-mm-dd ')||'" & strBegin & "','yyyy-mm-dd hh24:mi:ss') Between ��ʼʱ�� And ��ֹʱ�� " & _
            " And To_Date(To_char(��ʼʱ��,'yyyy-mm-dd ')||'" & strEnd & "','yyyy-mm-dd hh24:mi:ss') Between ��ʼʱ�� And ��ֹʱ��"
            
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str�ű�, CDate(Format(dtpDefer.Value, "yyyy-MM-dd 00:00:00")))
    If rsTmp.RecordCount = 0 Then
        MsgBox "ָ������������û�и�ҽ����Ч�ĹҺŰ���!" & vbCrLf & _
            "�����������ں͵�ǰ�ű�ĹҺŰ���.", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�������Ƿ�����,����ʾ,���̴���ʱ����
    intCol = GetListCol("����")
    For i = 1 To mshList.Rows - 1
        strTmp = Trim(mshList.TextMatrix(i, intCol))
        If strTmp <> "" Then strSNS = strSNS & ",'" & strTmp & "'"
    Next
    If strSNS <> "" Then
        strSQL = "Select ��� From �ٴ�������ſ��� Where Trunc(��ʼʱ��) = [1] And Instr([2], ','''||���||'''') > 0 And Not (״̬=3 And ����Ա����=[3])"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Format(dtpDefer.Value, "yyyy-MM-dd 00:00:00")), strSNS, CStr(UserInfo.����))
        strTmp = ""
        For i = 1 To rsTmp.RecordCount
            strTmp = strTmp & "," & rsTmp!���
            rsTmp.MoveNext
        Next
        
        If strTmp <> "" Then
            MsgBox "ע��:����ʱ��" & Format(dtpDefer.Value, "yyyy-MM-dd") & "����������ѱ�ʹ��:" & vbCrLf & Mid(strTmp, 2) & vbCrLf & _
                "ʹ����Щ��ŵ�ԤԼ�Һŵ�������ִ������!", vbInformation, gstrSysName
        End If
    End If
    
    strSQL = "zl_����ԤԼ�Һ�_Defer('" & str�ű� & "',To_date('" & Format(dtpBooking.Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS'),To_date('" & _
            Format(dtpDefer.Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS'),'" & UserInfo.���� & "'," & lng��¼ID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    Call dtpBooking_Change
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub dtpBooking_Change()
     If dtpDefer.Value <= dtpBooking.Value Then
        dtpDefer.Value = DateAdd("d", 1, dtpBooking.Value)
        dtpDefer.MinDate = dtpDefer.Value
    Else
        dtpDefer.MinDate = DateAdd("d", 1, dtpBooking.Value)
    End If
    
    Call SetPlanGrid
    Call ShowPlan(dtpBooking.Value)
    Call mshPlan_EnterCell
End Sub

Private Sub Form_Load()
    Dim Datsys As Date
    
    Datsys = zlDatabase.Currentdate
    dtpBooking.Value = DateAdd("d", 1, Datsys)
    dtpBooking.MinDate = dtpBooking.Value
    dtpDefer.Value = DateAdd("d", 1, dtpBooking.Value)
    dtpDefer.MinDate = dtpDefer.Value
    
    Call SetPlanGrid
    Call ShowPlan(dtpBooking.Value)
End Sub

Private Sub ShowPlan(datBooking As Date)
    Dim strSQL As String, i As Long, j As Long
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSQL = "Select Distinct E.����,E.���� �ű�, C.���� ����, B.ҽ������ ҽ��, Nvl(B.��Լ��,0) ��Լ," & vbNewLine & _
            " To_Char(B.��ʼʱ��,'YYYY-MM-DD') ��ʼʱ��,To_Char(B.��ֹʱ��,'YYYY-MM-DD') ��ֹʱ��,Decode(Nvl(B.�Ƿ���ſ���,0),1,'��',' ') as ��ſ���,B.ID As ��¼ID" & vbNewLine & _
            "From ������ü�¼ A, ���˹Һż�¼ D, �ٴ������¼ B, �ٴ������Դ E, ���ű� C " & vbNewLine & _
            "Where A.����ʱ�� Between [1] And [2] And A.��¼���� = 4 And A.��¼״̬ = 0 And A.���=1" & vbNewLine & _
            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & vbNewLine & _
            "      And A.NO = D.NO And D.�����¼ID = B.ID And B.��ԴID=E.ID And E.����id = C.ID "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Format(dtpBooking.Value, "yyyy-MM-dd 00:00:00")), _
                CDate(Format(dtpBooking.Value, "yyyy-MM-dd 23:59:59")))
    With mshPlan
        .ToolTipText = "�� " & rsTmp.RecordCount & " ����¼."
        .Rows = IIf(rsTmp.RecordCount = 0, 1, rsTmp.RecordCount) + 1
        For i = 1 To rsTmp.RecordCount
            For j = 0 To rsTmp.Fields.Count - 1
                .TextMatrix(i, j) = "" & rsTmp.Fields(j).Value
            Next
            rsTmp.MoveNext
        Next
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetPlanGrid()
    Dim i As Integer, strHead As String
    
    strHead = "����,1,600|�ű�,1,600|����,1,1050|ҽ��,4,800|��Լ,4,500|��ʼʱ��,4,1000|��ֹʱ��,4,1000|��ſ���,4,850|��¼ID,1,0"
       
    With mshPlan
        .Redraw = False
        .Clear: .Rows = 2
        .FixedRows = 1
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To .Cols - 1
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = flexAlignCenterCenter
        Next
        
        If Not Visible Then Call RestoreFlexState(mshPlan, App.ProductName & "\" & Me.Name)
        .RowHeight(0) = 300
        
        .Redraw = True
    End With
End Sub


Private Sub SetListGrid()
    Dim i As Integer, strHead As String
    
    strHead = "����,1,500|ʱ��,4,1200|NO,4,1250|Ʊ�ݺ�,4,1250|����,4,1250|�Ա�,4,500|����,4,800|�����,1,1450"
       
    With mshList
        .Redraw = False
        .Clear: .Rows = 2
        .FixedRows = 1
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To .Cols - 1
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = flexAlignCenterCenter
        Next
        
        If Not Visible Then Call RestoreFlexState(mshPlan, App.ProductName & "\" & Me.Name)
        .RowHeight(0) = 300
        
        .Redraw = True
    End With
End Sub

Private Sub ShowList(datBooking As Date, str�ű� As String)
    Dim strSQL As String, i As Long, j As Long
    On Error GoTo errH
    
    strSQL = "Select A.��ҩ���� ����,To_Char(A.����ʱ��,'hh24:mi:ss') ʱ��,A.NO, A.ʵ��Ʊ�� Ʊ�ݺ�, A.���� ����, A.�Ա�, A.����, A.��ʶ�� As �����" & vbNewLine & _
        "From ������ü�¼ A" & vbNewLine & _
        "Where A.����ʱ�� Between [1] And [2] And A.���㵥λ = [3] And A.��¼���� = 4 And A.��¼״̬ = 0 And A.��� = 1 Order by to_number(����)"


    Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Format(dtpBooking.Value, "yyyy-MM-dd 00:00:00")), _
                CDate(Format(dtpBooking.Value, "yyyy-MM-dd 23:59:59")), str�ű�)
    With mshList
        .ToolTipText = "�� " & mrsList.RecordCount & " ����¼."
        .Rows = IIf(mrsList.RecordCount = 0, 1, mrsList.RecordCount) + 1
        For i = 1 To mrsList.RecordCount
            For j = 0 To mrsList.Fields.Count - 1
                .TextMatrix(i, j) = "" & mrsList.Fields(j).Value
            Next
            mrsList.MoveNext
        Next
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set mrsList = Nothing
    
    Call SaveFlexState(mshPlan, App.ProductName & "\" & Me.Name)
    Call SaveFlexState(mshList, App.ProductName & "\" & Me.Name)
End Sub

Private Sub mshList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshList.MouseRow = 0 Then
        mshList.MousePointer = 99
    Else
        mshList.MousePointer = 0
    End If
End Sub

Private Sub mshList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intCol As Integer, intRow As Integer
    
    intCol = mshList.MouseCol
    intRow = mshList.MouseRow
    If intRow = 0 Then
        mshList.ColData(intCol) = (mshList.ColData(intCol) + 1) Mod 2
        mshList.ColSel = mshList.Col
        mshList.Sort = Val(mshList.ColData(intCol)) + 1 '1-��,2-��
    End If
End Sub

Private Sub mshPlan_EnterCell()
    Dim i As Integer, blnPre As Boolean
    Dim intRow As Integer, intCol As Integer, str�ű� As String
    
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
    
    str�ű� = mshPlan.TextMatrix(mshPlan.Row, GetPlanCol("�ű�"))
    Call SetListGrid
    Call ShowList(dtpBooking.Value, str�ű�)

    cmdOK.Enabled = (str�ű� <> "")
End Sub

Private Function GetPlanCol(strName As String) As Integer
    Dim i As Integer
    For i = 0 To mshPlan.Cols - 1
        If mshPlan.TextMatrix(0, i) = strName Then
            GetPlanCol = i: Exit For
        End If
    Next
End Function

Private Function GetListCol(strName As String) As Integer
    Dim i As Integer
    For i = 0 To mshList.Cols - 1
        If mshList.TextMatrix(0, i) = strName Then
            GetListCol = i: Exit For
        End If
    Next
End Function

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

Private Sub mshPlan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshPlan.MouseRow = 0 Then
        mshPlan.MousePointer = 99
    Else
        mshPlan.MousePointer = 0
    End If
End Sub

Private Sub mshPlan_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intCol As Integer, intRow As Integer
    
    intCol = mshPlan.MouseCol
    intRow = mshPlan.MouseRow
    If intRow = 0 Then
        mshPlan.ColData(intCol) = (mshPlan.ColData(intCol) + 1) Mod 2
        mshPlan.ColSel = mshPlan.Col
        mshPlan.Sort = Val(mshPlan.ColData(intCol)) + 1 '1-��,2-��
    End If
End Sub

Private Sub mshPlan_SelChange()
    If mshPlan.Rows = 2 Then Exit Sub
    mshPlan.RowSel = mshPlan.Row
End Sub




Private Sub mshList_EnterCell()
    Dim i As Integer, blnPre As Boolean
    Dim intRow As Integer, intCol As Integer
    
    blnPre = mshList.Redraw
    intRow = mshList.Row: intCol = mshList.Col
    mshList.Redraw = False
    
    For i = 0 To mshList.Cols - 1
        mshList.Col = i
        mshList.CellBackColor = mshList.BackColorSel
        mshList.CellForeColor = mshList.ForeColorSel
    Next
    
    mshList.Row = intRow:  mshList.Col = intCol
    mshList.Redraw = blnPre
End Sub

Private Sub mshList_LeaveCell()
    Dim i As Integer, blnPre As Boolean
    
    blnPre = mshList.Redraw
    mshList.Redraw = False
    
    For i = 0 To mshList.Cols - 1
        mshList.Col = i
        mshList.CellBackColor = mshList.BackColor
        mshList.CellForeColor = mshList.ForeColor
    Next
    
    mshList.Redraw = blnPre
End Sub

Private Sub mshList_SelChange()
    If mshList.Rows = 2 Then Exit Sub
    mshList.RowSel = mshList.Row
End Sub
