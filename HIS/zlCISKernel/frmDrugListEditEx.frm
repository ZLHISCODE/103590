VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmDrugListEditEx 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ҩ�䷽�༭"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4860
   Icon            =   "frmDrugListEditEx.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   4860
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboData 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   615
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   6000
      Width           =   2000
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H80000004&
      Cancel          =   -1  'True
      Height          =   315
      Left            =   3960
      Picture         =   "frmDrugListEditEx.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "ȡ��(Esc)"
      Top             =   6000
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   315
      Left            =   2880
      Picture         =   "frmDrugListEditEx.frx":6DDC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "ȷ��(F2)"
      Top             =   6000
      Width           =   885
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   5955
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4845
      _cx             =   8546
      _cy             =   10504
      Appearance      =   0
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
      MouseIcon       =   "frmDrugListEditEx.frx":7366
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16444122
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   16777215
      GridColorFixed  =   16777215
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   2
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   10000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDrugListEditEx.frx":7C40
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.PictureBox pictmp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1920
         ScaleHeight     =   240
         ScaleWidth      =   480
         TabIndex        =   5
         Top             =   1680
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.Label lblData 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�巨"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   6060
      Width           =   390
   End
End
Attribute VB_Name = "frmDrugListEditEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum COLDurg
    COL_������ĿID = 1
    COL_�շ�ϸĿID = 2
    col_��ҩ = 3
    COL_���� = 4
    COL_��λ = 5
    COL_��ע = 6
End Enum

Private Const GRD_UNEDITCELL_COLOR = &H8000000B  'δ�༭�ĵ�Ԫ����ɫ������ɫ


Private mblnReturn As Boolean
Private mblnOK As Boolean
Private mstrLike As String
Private mint���� As Integer
Private mlng�巨ID As Long
Private mstrData As String  '��ʽ:[�䷽����]��ҩ����<Data>������ĿID<Data>�շ�ϸĿID<Data>����<Data>��ע<Data>��λ


Public Function ShowEdit(frmParent As Object, ByRef strData As String, ByRef lng�巨 As Long) As Boolean
'���ܣ���ҩ�嵥��ҩ�༭��
'������vsTmp ������ҩ�嵥����ı��ؼ�

    On Error Resume Next
    mlng�巨ID = 0
    mstrData = ""
    mblnOK = False

    mstrData = strData
    mlng�巨ID = lng�巨
    
    Me.Show 1, frmParent
    strData = mstrData
    lng�巨 = mlng�巨ID
    ShowEdit = mblnOK
    On Error GoTo 0
End Function

Private Sub LoadData()
    Dim arrTime As Variant, arrTmp As Variant
    Dim i As Long
    With vsAdvice
        If mstrData = "" Then
            .Rows = 1
            .Rows = vsAdvice.Rows + 1
        Else
             .Redraw = flexRDNone
             .Rows = .FixedRows
             arrTime = Split(mstrData, "[�䷽����]")
            For i = 1 To UBound(arrTime)
                .Rows = .Rows + 1
                arrTmp = Split(arrTime(i), "<Data>")
                .TextMatrix(.Rows - 1, col_��ҩ) = arrTmp(0)
                .TextMatrix(.Rows - 1, COL_������ĿID) = arrTmp(1)
                .TextMatrix(.Rows - 1, COL_�շ�ϸĿID) = arrTmp(2)
                .TextMatrix(.Rows - 1, COL_����) = arrTmp(3)
                .TextMatrix(.Rows - 1, COL_��ע) = arrTmp(4)
                .TextMatrix(.Rows - 1, COL_��λ) = arrTmp(5)
            Next
             .Redraw = flexRDDirect
        End If
        .Row = .Rows - 1: .Col = col_��ҩ
        .ShowCell .Rows - 1, col_��ҩ
        .Cell(flexcpBackColor, .FixedRows, COL_��λ, .Rows - 1, COL_��λ) = GRD_UNEDITCELL_COLOR      '����ɫ
    End With
End Sub


Private Sub InitAdviceTable()
'���ܣ���ʼ��ҽ���嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long

    strHead = "������ĿID;�շ�ϸĿID;��ҩ,2000,1;����,850,4;��λ,850,4;��ע,950,1"
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1: .FixedCols = 1
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        .SelectionMode = flexSelectionFree
        .FocusRect = flexFocusSolid
        .HighLight = flexHighlightWithFocus
        .BackColorSel = &H404040

        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .FixedAlignment(.FixedCols + i) = 4
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                lngWidth = Val(Split(arrHead(i), ",")(1))
                .ColWidth(.FixedCols + i) = lngWidth
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                'Ϊ��֧��zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0    'Ϊ��֧��zl9PrintMode
            End If
            .ColData(.FixedCols + i) = .ColWidth(.FixedCols + i)    '��¼ԭʼ�п�������ѡ����
        Next
        .Editable = flexEDKbdMouse
    End With
End Sub

Public Function AdviceCellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    With vsAdvice
        If .ColHidden(lngCol) Then Exit Function
        '������������ҩ����
        If lngCol = COL_��λ Then Exit Function
        If .TextMatrix(lngRow, col_��ҩ) = "" Then
            If lngCol > col_��ҩ Then Exit Function
        End If
    End With
    AdviceCellEditable = True
End Function


Private Sub EnterNextCellAdvice()
    Dim i As Long, j As Long

    With vsAdvice
        '����һ��Ԫ��ʼѭ������
        If .Row < .FixedRows Then
            .Rows = .Rows + 1
            .Cell(flexcpBackColor, .FixedRows, COL_��λ, .Rows - 1, COL_��λ) = GRD_UNEDITCELL_COLOR      '����ɫ
            .ShowCell .Rows - 1, col_��ҩ
        End If
        For i = .Row To .Rows - 1
            For j = IIF(i = .Row, .Col + 1, col_��ҩ) To COL_��ע
                If AdviceCellEditable(i, j) Then Exit For
            Next
            If j <= COL_��ע Then Exit For
        Next
        If i <= .Rows - 1 Then
            Call .Select(i, j)
            .ShowCell .Row, .Col
        ElseIf i = .Rows And j > COL_��ע And .TextMatrix(.Rows - 1, col_��ҩ) <> "" Then
            .Rows = .Rows + 1
            .Cell(flexcpBackColor, .FixedRows, COL_��λ, .Rows - 1, COL_��λ) = GRD_UNEDITCELL_COLOR      '����ɫ
            .ShowCell .Rows - 1, col_��ҩ
        End If
    End With
End Sub

Private Function checkDrug()
    Dim i As Long, j As Long
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If .TextMatrix(i, col_��ҩ) <> "" Then
                '���䷽��Ϊ��ʱ�����ҩ�巨
                If cboData.ListIndex = -1 Then
                    MsgBox "��ȷ����ҩ�䷽�ļ巨��", vbInformation, gstrSysName
                    cboData.SetFocus: Exit Function
                End If
                
                
                If Val(.TextMatrix(i, COL_����)) <= 0 Then
                    MsgBox "��ҩ�䷽�ĵ���Ϊ������,��¼�롣", vbInformation, gstrSysName
                    .SetFocus
                    .Row = i: .Col = COL_����: Call vsAdvice.ShowCell(.Row, .Col)
                    Exit Function
                End If
                
                If i <> .Rows - 1 Then  '����Ƿ������ͬ��ҩ�䷽
                    For j = .Rows - 1 To i + 1 Step -1
                        If .TextMatrix(j, col_��ҩ) <> "" Then
                            If .TextMatrix(j, col_��ҩ) & "|" & .TextMatrix(j, COL_������ĿID) & "|" & .TextMatrix(j, COL_�շ�ϸĿID) = .TextMatrix(i, col_��ҩ) & "|" & .TextMatrix(i, COL_������ĿID) & "|" & .TextMatrix(i, COL_�շ�ϸĿID) Then
                                .SetFocus
                                MsgBox "���������ظ�����ҩ�嵥,���顣", vbInformation, gstrSysName
                                .Row = j: .Col = col_��ҩ: Call vsAdvice.ShowCell(.Row, .Col)
                                Exit Function
                            End If
                        End If
                    Next
                End If
            End If
        Next
        checkDrug = True
    End With
End Function


Private Sub cmdOK_Click()
    Dim i As Long
    Dim strTmp As String
    With vsAdvice
        If .Rows <= 2 And .TextMatrix(.Rows - 1, col_��ҩ) = "" And mstrData <> "" Then
           If MsgBox("��ȷ���Ƿ������ҩ�䷽���ݣ�", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                Exit Sub
           End If
           mstrData = ""
           mlng�巨ID = 0
           mblnOK = True
           Unload Me
        Else
            If checkDrug Then
                mlng�巨ID = Val(cboData.ItemData(cboData.ListIndex))
                For i = 1 To vsAdvice.Rows - 1
                    If .TextMatrix(i, col_��ҩ) <> "" Then
                        strTmp = strTmp & "[�䷽����]" & .TextMatrix(i, col_��ҩ) & "<Data>" & Val(.TextMatrix(i, COL_������ĿID)) & "<Data>" & Val(.TextMatrix(i, COL_�շ�ϸĿID)) & "<Data>" & FormatEx(NVL(.TextMatrix(i, COL_����)), 5) & "<Data>" & .TextMatrix(i, COL_��ע) & "<Data>" & .TextMatrix(i, COL_��λ)
                    End If
                Next
                mstrData = strTmp
                mblnOK = True
                Unload Me
            End If
        End If
    End With
End Sub


Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    If NewRow = -1 Or NewCol = -1 Then Exit Sub
    If (Not AdviceCellEditable(NewRow, NewCol)) Then
        vsAdvice.FocusRect = flexFocusLight
    Else
        vsAdvice.FocusRect = flexFocusSolid
    End If
End Sub

Private Sub vsAdvice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    With vsAdvice
        .EditSelStart = 0
        .EditSelLength = zlCommFun.ActualLen(.EditText)
    End With
End Sub

Private Sub vsAdvice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not AdviceCellEditable(Row, Col) Then
        Cancel = True
    End If
End Sub


Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    With vsAdvice
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0
            mblnReturn = True
            Call EnterNextCellAdvice
        Else
            If .ComboList = "..." Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
                End If
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsAdvice
        If Not KeyAscii = vbKeyReturn Then
            If Col = COL_���� Then
                If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            End If
            mblnReturn = False
        Else
            mblnReturn = True
        End If
    End With
End Sub


Private Sub vsAdvice_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsAdvice
        If KeyCode = vbKeyF4 Then
            If .ComboList = "..." Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .Row < 1 Then Exit Sub
            If MsgBox("ȷʵҪɾ��������ҩ�䷽��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                .RemoveItem .Row
                If .Rows = 1 Then
                    .Rows = .Rows + 1
                    .Cell(flexcpBackColor, .FixedRows, COL_��λ, .Rows - 1, COL_��λ) = GRD_UNEDITCELL_COLOR      '����ɫ
                    .Row = .Rows - 1: .Col = col_��ҩ
                    .ShowCell .Rows - 1, col_��ҩ
                End If
            Else
                Exit Sub
            End If
                
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsAdvice_KeyPress(KeyCode)
        End If
    End With
End Sub



Private Sub vsAdvice_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'vsAdvice_ValidateEdit�¼�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As PointAPI
    Dim strLike As String
    Dim strInput As String
    Dim lngMax As Long

    On Error GoTo errH
   With vsAdvice
        strLike = mstrLike
        If Len(.EditText) < 2 Then strLike = "" '�Ż�
        Select Case Col
            Case col_��ҩ
                If .EditText = "" Then
                    .EditText = .Cell(flexcpData, Row, col_��ҩ)
                    If mblnReturn Then Call EnterNextCellAdvice
                ElseIf .EditText = .Cell(flexcpData, Row, col_��ҩ) Then
                    If mblnReturn Then Call EnterNextCellAdvice
                Else
                    strInput = " And (A.���� Like [1] And E.����=[3]" & _
                        " Or E.���� Like [2] And E.����=[3] Or E.���� Like [2] And E.���� IN([3],3))"
                
                    If IsNumeric(.EditText) Then
                        '1X.����ȫ������ʱֻƥ�����'����ҩƷ,��Ҫƥ�����(����Ϊ3��������)
                        If Mid(gstrMatchMode, 1, 1) = "1" Then strInput = " And (A.���� Like [1] And E.����=[3] Or E.���� Like [2] And E.����=3)"
                    ElseIf zlCommFun.IsCharAlpha(.EditText) Then
                        'X1.����ȫ����ĸʱֻƥ�����
                        If Mid(gstrMatchMode, 2, 1) = "1" Then strInput = " And E.���� Like [2] And E.����=[3]"
                    ElseIf zlCommFun.IsCharChinese(.EditText) Then
                        '��������,��ֻƥ������a
                        strInput = " And E.���� Like [2] And E.����=[3]"
                    End If
                    
                    strInput = IIF(.EditText = "*", "", strInput)
                    strSQL = "Select distinct a.Id, b.Id As �շ�ϸĿid, a.����, b.���, a.���㵥λ" & _
                    " From ������ĿĿ¼ A, �շ���ĿĿ¼ B, ҩƷ��� C, ҩƷ���� D,������Ŀ���� E " & _
                    " Where c.ҩƷid= b.Id(+) And a.Id =c.ҩ��id(+) And c.ҩ��id = d.ҩ��id(+) And A.ID=E.������ĿID(+) And a.��� ='7' and (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & strInput
                    
                    vPoint = zlControl.GetCoordPos(vsAdvice.hwnd, vsAdvice.CellLeft, vsAdvice.CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ҩƷĿ¼", False, "", "", False, False, True, _
                        vPoint.X, vPoint.Y, vsAdvice.RowHeight(Row), blnCancel, False, True, UCase(.EditText) & "%", strLike & UCase(.EditText) & "%", mint���� + 1)
                    
                    If rsTmp Is Nothing Then
                        If Not blnCancel Then
                            MsgBox "δ�ҵ����õ��в�ҩ�����ȵ�������Ŀ���������á�", vbInformation, gstrSysName
                            Cancel = True
                        Else
                            .EditText = .Cell(flexcpData, Row, col_��ҩ)
                        End If
                        Exit Sub
                    Else
                        .EditText = rsTmp!���� & ""
                        .TextMatrix(Row, col_��ҩ) = rsTmp!���� & ""
                        .Cell(flexcpData, Row, col_��ҩ) = .TextMatrix(Row, col_��ҩ)
                        .TextMatrix(Row, COL_������ĿID) = Val(rsTmp!ID & "")
                        .TextMatrix(Row, COL_�շ�ϸĿID) = Val(rsTmp!�շ�ϸĿID & "")
                        .TextMatrix(Row, COL_��λ) = IIF(rsTmp!���㵥λ & "" = "", "g", rsTmp!���㵥λ & "")
                    End If
                End If
            Case COL_����
                lngMax = 10
            Case COL_��ע
                lngMax = 100
        End Select
        
        If LenB(StrConv(.EditText, vbFromUnicode)) > lngMax And lngMax <> 0 Then
            MsgBox "���ܳ���" & lngMax & "���ַ��ĳ��ȡ�", vbInformation, gstrSysName
            Cancel = True
            Exit Sub
        End If
        
        mblnReturn = False
   End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call InitAdviceTable
    Call Get�巨
    Call LoadData
    '����ƥ��
    mstrLike = IIF(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
    '����ƥ�䷽ʽ��0-ƴ��,1-���
    mint���� = Val(zlDatabase.GetPara("���뷽ʽ"))
End Sub



Private Sub Get�巨()
     '��ҩ�巨
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    On Error GoTo errH
    strSQL = "Select A.ID,A.����,A.���� From ������ĿĿ¼ A" & _
        " Where A.���='E' And A.��������='3' And A.������� IN(1,2,3)" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " Order by A.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTmp.Filter <> 0 Then rsTmp.Filter = 0
    If rsTmp.EOF Then
        MsgBox "δ�ҵ���Ч����ҩ�巨�����ȵ�������Ŀ���������á�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    For i = 1 To rsTmp.RecordCount
        cboData.AddItem rsTmp!���� & "-" & rsTmp!����
        cboData.ItemData(cboData.NewIndex) = rsTmp!ID
        If rsTmp!ID = mlng�巨ID Then
            Call Cbo.SetIndex(cboData.hwnd, cboData.NewIndex)
        End If
        rsTmp.MoveNext
    Next
    
    If cboData.ListCount = 1 And cboData.ListIndex = -1 Then Call Cbo.SetIndex(cboData.hwnd, 0)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    vsAdvice.SetFocus
End Sub


Private Sub cboData_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cboData.ListIndex <> -1 Then
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        lngIdx = Cbo.MatchIndex(cboData.hwnd, KeyAscii)
        If lngIdx = -1 And cboData.ListCount > 0 Then lngIdx = 0
        cboData.ListIndex = lngIdx
    End If
End Sub
