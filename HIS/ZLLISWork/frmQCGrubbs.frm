VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmQCGrubbs 
   BorderStyle     =   0  'None
   Caption         =   "Grubbs�ʿؼ�¼��"
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   10035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.ComboBox cboQCitem 
      Height          =   300
      Left            =   4380
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   6750
      Width           =   2595
   End
   Begin VB.OptionButton opt�ʿ�Ʒ 
      Caption         =   "473843A��ֵ�ʿ�Ʒ"
      Enabled         =   0   'False
      Height          =   180
      Index           =   0
      Left            =   75
      TabIndex        =   1
      Top             =   6855
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgData 
      Height          =   5010
      Left            =   450
      TabIndex        =   0
      Top             =   1620
      Width           =   9105
      _cx             =   16060
      _cy             =   8837
      Appearance      =   2
      BorderStyle     =   0
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
      BackColorFixed  =   15790320
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16635590
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
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   2475
      TabIndex        =   2
      Top             =   435
      Visible         =   0   'False
      Width           =   90
   End
End
Attribute VB_Name = "frmQCGrubbs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    ���� = 1:  ����: �ⶨֵ: ��ֵ: SD: SI����: SI����: N: n3s: n2s: ���: ������
End Enum

Private mstrResList As String
Private mlngItemID As Long
Private mstrFromDate As String
Private mstrToDate As String
Private mstr�ʿ�Ʒ���� As String

Public Function zlRefresh(strResList As String, lngItemID As Long, strFromDate As String, strToDate As String, str�ʿ�Ʒ���� As String) As Boolean
    '���ܣ�ˢ�±������������ʾ����
    '������ strResList  ��ǰѡ����ʿ�Ʒid�����Զ��ŷָ�
    '       lngItemId   ��ǰ��Ŀid
    '       strFromDate ��ʼ����
    '       strToDate   ��������
    Dim rsTemp As New adodb.Recordset
    Dim intCounts As Integer
    Dim lngResId As Long
    Dim lngCount As Long
    
    mstrResList = strResList
    mlngItemID = lngItemID
    mstrFromDate = strFromDate
    mstrToDate = strToDate
    mstr�ʿ�Ʒ���� = str�ʿ�Ʒ����
    Me.Tag = "��ˢ��"
    
    lngResId = 0
    intCounts = Me.cboQCitem.ListCount
    For lngCount = intCounts - 1 To 1 Step -1
        If Me.cboQCitem.ListIndex = lngCount Then lngResId = Val(Me.cboQCitem.ItemData(lngCount))
'        Unload Me.opt�ʿ�Ʒ(Me.opt�ʿ�Ʒ.UBound)
    Next
    cboQCitem.Clear
    
    Me.opt�ʿ�Ʒ(0).Enabled = False
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "Select A.ID, A.���� || '-' || A.���� As �ʿ�Ʒ, B.�����ʿ�ͼ From �����ʿ�Ʒ A,�������� B Where A.����ID=B.ID(+) And Instr(',' || [1] || ',', ',' || A.ID || ',') > 0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strResList)
    With rsTemp
        Do While Not .EOF
            If .AbsolutePosition > Me.cboQCitem.ListCount Then cboQCitem.AddItem "" & !�ʿ�Ʒ
            cboQCitem.ItemData(cboQCitem.NewIndex) = !ID
'            If .AbsolutePosition > Me.opt�ʿ�Ʒ.Count Then Load Me.opt�ʿ�Ʒ(.AbsolutePosition - 1)
'            Me.opt�ʿ�Ʒ(.AbsolutePosition - 1).Caption = "" & !�ʿ�Ʒ
'            Me.opt�ʿ�Ʒ(.AbsolutePosition - 1).Tag = !ID
'            Me.opt�ʿ�Ʒ(.AbsolutePosition - 1).Width = Me.TextWidth(Me.opt�ʿ�Ʒ(.AbsolutePosition - 1).Caption) + 360
'            Me.opt�ʿ�Ʒ(.AbsolutePosition - 1).Value = (lngResId = !ID)
'            Me.opt�ʿ�Ʒ(.AbsolutePosition - 1).Visible = True
'            Me.opt�ʿ�Ʒ(.AbsolutePosition - 1).Enabled = True
            
            .MoveNext
        Loop
    End With
    If rsTemp.RecordCount > 0 Then Me.cboQCitem.ListIndex = 0
    Me.Tag = ""
    Call RefGrid
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog

End Function

Private Sub RefGrid()
'
    Dim rsTemp As New adodb.Recordset
    Dim lngResId As Long, strLable As String, strUnit As String
    Dim intFormatNum As Integer, curTotal As Currency, strData As String
    Dim lng���� As Long, strLast���� As String, cur��ֵ, curSD As Currency, curMax As Currency, curMin As Currency
    Dim curSI�� As Currency, curSI�� As Currency, curn3s As Currency, curn2s As Currency, curCV As Currency
    Dim lngCount As Long, lngRow As Long, iCol As Integer
    On Error GoTo ErrHandle
    lngResId = 0
    For lngCount = 0 To Me.cboQCitem.ListCount - 1
        If Me.cboQCitem.ListIndex = lngCount Then lngResId = Val(Me.cboQCitem.ItemData(lngCount))
    Next
    If lngResId = 0 Then
        Me.opt�ʿ�Ʒ(0).Enabled = False
        Me.opt�ʿ�Ʒ(0).Value = True
        lngResId = Val(Me.opt�ʿ�Ʒ(0).Tag)
        Me.opt�ʿ�Ʒ(0).Enabled = True
    End If
    
    '��ȡС��λ��
'    gstrSql = "Select nvl(С��λ��,2) as С��λ�� from ����������Ŀ where ��ĿID = [1] "
'    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, mlngItemId)
'    If rsTemp.EOF = False Then intFormatNum = Val("" & rsTemp("С��λ��"))
    intFormatNum = 3
    Call initVfgData
    
    '��û�����������Ϣ

    gstrSql = "Select Distinct RPad('��λ��' || '" & gstrUnitName & "', 56, ' ') || '���ڣ�' As ��0," & vbNewLine & _
            "                RPad('������' || D.����, 56, ' ') || '�Լ���Դ��' || M.�Լ� As ��1," & vbNewLine & _
            "                RPad('��Ŀ��' || I.��Ŀ, 56, ' ') || 'У׼����Դ��' || M.У׼�� As ��2" & vbNewLine & _
            "From �������� D, �����ʿ�Ʒ M, (Select ������ || ',' || Ӣ���� As ��Ŀ From ����������Ŀ Where ID = [2]) I" & vbNewLine & _
            "Where D.ID = M.����id And Instr(',' || [1] || ',', ',' || M.ID || ',') > 0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrResList, mlngItemID)
    Me.vfgData.Visible = True
    Me.lblInfo.Visible = False
    If rsTemp.RecordCount <= 0 Then
       Me.lblInfo.Caption = "���ʿ�Ʒ��Ϣ��ȫ�棡"
       Me.lblInfo.Visible = True
       Me.vfgData.Visible = False
       Exit Sub
    End If
    '��ͷ����
    With vfgData
        For iCol = .FixedCols To .Cols - 1
            
            .TextMatrix(1, iCol) = "  " & rsTemp!��0 & Format(mstrFromDate, "yyyy��MM��dd��") & "��" & Format(mstrToDate, "yyyy��MM��dd��") & vbCrLf & _
                                   "  " & rsTemp!��1 & vbCrLf & "  " & rsTemp!��2
        Next
        .Cell(flexcpAlignment, 1, 0, 1, .Cols - 1) = flexAlignLeftCenter
    End With
    
    Call Form_Resize
    
    gstrSql = "Select Q.����ʱ�� as ����,Q.���Դ���, Zl_Lis_Tonumber(Q.�ʿ�Ʒid, R.������Ŀid, R.������,R.ID) As ���,R.���ý��,Q.������ " & vbNewLine & _
            "From �����ʿؼ�¼ Q, ������ͨ��� R,�����ʿ�Ʒ M, �����ʿؾ�ֵ X " & vbNewLine & _
            "Where Q.�걾id = R.����걾id And Q.�걾id = R.����걾id And" & vbNewLine & _
            "      Nvl(R.���ý��,0)=0 And Q.�ʿ�Ʒid =[1] And R.������Ŀid + 0 = [2] And" & vbNewLine & _
            "      Q.����ʱ�� Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd') And " & vbNewLine & _
            "      (Q.����ʱ�� Between X.��ʼ���� And NVL(X.��������,M.��������)) And " & vbNewLine & _
            "       Q.�ʿ�Ʒid=M.id And M.id=X.�ʿ�Ʒid  And  X.��ĿID = [2] And " & vbNewLine & _
            "      Instr(';'||[5]||';',';' || X.�ʿ�Ʒid||'='||To_char(X.��ʼ����,'yyyy-MM-dd')||','||to_char(Nvl(X.��������, M.��������),'yyyy-mm-dd')||';' ) > 0" & vbNewLine & _
            "Order By Q.����ʱ��,Q.���Դ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResId, mlngItemID, mstrFromDate, mstrToDate, mstr�ʿ�Ʒ����)
    lng���� = 0
    With vfgData
        Do Until rsTemp.EOF
            If strLast���� <> Format("" & rsTemp!����, "yyyy-MM-dd") & "," & rsTemp!���Դ��� Then
                lng���� = lng���� + 1
                .TextMatrix(.Rows - 1, mCol.����) = Format("" & rsTemp!����, "yyyy-MM-dd")
                .TextMatrix(.Rows - 1, mCol.����) = lng����
                .TextMatrix(.Rows - 1, mCol.�ⶨֵ) = Format(Val("" & rsTemp!���), "0." & String(intFormatNum, "0"))
                .TextMatrix(.Rows - 1, mCol.������) = "" & rsTemp!������
                    
                .Rows = .Rows + 1
                If lng���� >= 20 Then Exit Do
            ElseIf strLast���� <> "" Then
                .TextMatrix(.Rows - 2, mCol.�ⶨֵ) = Format(Val("" & rsTemp!���), "0." & String(intFormatNum, "0"))
                .TextMatrix(.Rows - 2, mCol.������) = "" & rsTemp!������
            End If
            strLast���� = Format("" & rsTemp!����, "yyyy-MM-dd") & "," & rsTemp!���Դ���
            rsTemp.MoveNext
        Loop
        curTotal = 0
        
        If .Rows > 4 Then .Rows = .Rows - 1
        
        If .Rows > 3 Then
            gstrSql = "Select n,n3s,n2s From �ʿؼ��̷� "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
        
            For lngRow = 3 To .Rows - 1
                strData = strData & "," & Val(.TextMatrix(lngRow, mCol.�ⶨֵ))
                If Val(.TextMatrix(lngRow, mCol.����)) > 2 Then
                
                    curn3s = Val("" & rsTemp!n3s)
                    curn2s = Val("" & rsTemp!n2s)
                    
                    .TextMatrix(lngRow, mCol.N) = rsTemp!N
                    .TextMatrix(lngRow, mCol.n3s) = curn3s
                    .TextMatrix(lngRow, mCol.n2s) = curn2s
                    

                    cur��ֵ = s(strData): curSD = stdev(strData)
                    curMax = Max(strData): curMin = Min(strData)
                    If curSD <> 0 Then
                        curSI�� = curMax / curSD - cur��ֵ / curSD: curSI�� = cur��ֵ / curSD - curMin / curSD
                    End If
                    .TextMatrix(lngRow, mCol.��ֵ) = Format(cur��ֵ, "0." & String(intFormatNum, "0"))
                    .TextMatrix(lngRow, mCol.SD) = Format(curSD, "0." & String(intFormatNum, "0"))
                    .TextMatrix(lngRow, mCol.SI����) = Format(curSI��, "0.00")
                    .TextMatrix(lngRow, mCol.SI����) = Format(curSI��, "0.00")
                    
                    If curSI�� > curn3s Or curSI�� > curn3s Then              '090504 ��һ������3s ��ʧ��
                        .TextMatrix(lngRow, mCol.���) = "#" 'ʧ��
                        .Cell(flexcpForeColor, lngRow, mCol.���) = &H40C0&   '090504 ��С��2s ���ڿ�
                    ElseIf curSI�� < curn2s And curSI�� < curn2s Then
                        .TextMatrix(lngRow, mCol.���) = "*" '�ڿ�
                    Else
                        .TextMatrix(lngRow, mCol.���) = "��" '����
                        .Cell(flexcpForeColor, lngRow, mCol.���) = &H80C0FF
                    End If
                    rsTemp.MoveNext
                End If
            Next
        End If
        '���һ��
        
        .Rows = .Rows + 1
        lngRow = .Rows - 1
        
        .TextMatrix(.Rows - 1, mCol.����) = lng���� & "���ڿ����ݲⶨ��"
        .TextMatrix(.Rows - 1, mCol.����) = lng���� & "���ڿ����ݲⶨ��"
        
        If Val(.TextMatrix(.Rows - 2, mCol.�ⶨֵ)) <> 0 And _
           Val(.TextMatrix(.Rows - 2, mCol.����)) > 2 And _
           Val(.TextMatrix(.Rows - 2, mCol.����)) < 21 Then
            curCV = curSD / Val(.TextMatrix(.Rows - 2, mCol.�ⶨֵ)) * 100
        End If
        For iCol = mCol.�ⶨֵ To .Cols - 1
            .TextMatrix(.Rows - 1, iCol) = "��ֵ=" & Format(cur��ֵ, "0.000") & Space(10) & "SD=" & Format(curSD, "0.000") & Space(10) & "CV%=" & Format(curCV, "0.000")
        Next
        .MergeRow(.Rows - 1) = True
        
        .Select 2, .FixedCols, .Rows - 1, .Cols - 1
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        .Select 2, .FixedCols
        '��β ˵��
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, mCol.����) = "˵����"
        For iCol = mCol.���� To .Cols - 1
            .TextMatrix(.Rows - 1, iCol) = "1��SI����ֵ������ֵ  < n2s  ʱΪ�ڿأ�n2s �� n3s ֮��Ϊ�澯״̬��> n3s  ʱ"
        Next
        .MergeRow(.Rows - 1) = True
        
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, mCol.����) = ""
        For iCol = mCol.���� To .Cols - 1
            .TextMatrix(.Rows - 1, iCol) = " ������ԡ�*����ʾ�ڿأ���������ʾ�澯����#����ʾʧ�ء�"
        Next
        .MergeRow(.Rows - 1) = True
        
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, mCol.����) = ""
        For iCol = mCol.���� To .Cols - 1
            .TextMatrix(.Rows - 1, iCol) = "2��ʧ�����ݼ�ԭ��Ҫ��д�������ʿ�ʧ�ر��桱��"
        Next
        .MergeRow(.Rows - 1) = True
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, mCol.����) = ""
        For iCol = mCol.���� To .Cols - 1
            .TextMatrix(.Rows - 1, iCol) = "3�����澯��ʧ��״̬�������ز⡣�����澯��ʧ�ص���ֵ�������ⶨ��ֵ����ʹ�á�"
        Next
        .MergeRow(.Rows - 1) = True
        
        .Cell(flexcpAlignment, lngRow, .FixedCols, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
        .Cell(flexcpAlignment, lngRow + 1, .FixedCols) = flexAlignRightCenter
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Function s(ByVal strVal As String) As Currency
'   ��ֵ
    Dim varInData As Variant, curX As Currency, i As Integer
    If Left(strVal, 1) = "," Then
        varInData = Split(Mid(strVal, 2), ",")
    Else
        varInData = Split(strVal, ",")
    End If
    For i = LBound(varInData) To UBound(varInData)
        curX = curX + Val(varInData(i))
    Next
    If i > 0 Then
        s = curX / i
    End If
End Function
Private Function stdev(ByVal strVal As String) As Currency
    '��׼��
    Dim varInData As Variant, curX As Currency, i As Integer, cur��ֵ As Currency
    
    If Left(strVal, 1) = "," Then
        varInData = Split(Mid(strVal, 2), ",")
    Else
        varInData = Split(strVal, ",")
    End If
    cur��ֵ = s(strVal)
    For i = LBound(varInData) To UBound(varInData)
        curX = curX + (Val(varInData(i)) - cur��ֵ) ^ 2
    Next
    If i - 1 > 0 Then
        stdev = Sqr(curX / (i - 1))
    End If
    'Sqr (��(xn - x��) ^ 2 / (N - 1))
End Function

Private Function Max(ByVal strVal As String) As Currency
    Dim varInData As Variant, curX As Currency, i As Integer
    If Left(strVal, 1) = "," Then
        varInData = Split(Mid(strVal, 2), ",")
    Else
        varInData = Split(strVal, ",")
    End If
    For i = LBound(varInData) To UBound(varInData)
        If i = LBound(varInData) Then
            curX = Val(varInData(i))
        Else
            If curX < Val(varInData(i)) Then curX = Val(varInData(i))
        End If
    Next
    Max = curX
End Function

Private Function Min(ByVal strVal As String) As Currency
    Dim varInData As Variant, curX As Currency, i As Integer
    If Left(strVal, 1) = "," Then
        varInData = Split(Mid(strVal, 2), ",")
    Else
        varInData = Split(strVal, ",")
    End If
    For i = LBound(varInData) To UBound(varInData)
        If i = LBound(varInData) Then
            curX = Val(varInData(i))
        Else
            If curX > Val(varInData(i)) Then curX = Val(varInData(i))
        End If
    Next
    Min = curX
End Function

Private Sub initVfgData()
    Dim iCol As Integer
    With vfgData
        .Editable = flexEDNone
        .GridLines = flexGridNone
        .Rows = 4: .Cols = 13
        .FixedCols = 1: .FixedRows = 3
        .MergeCells = flexMergeRestrictRows
        .BackColorFixed = .BackColor
        .ForeColorFixed = .ForeColor
        .GridColorFixed = .GridColor
        .GridLinesFixed = flexGridNone
        
        '-- ��ͷ
        For iCol = 0 To 1
            .MergeRow(iCol) = True
        Next
        
        For iCol = .FixedCols To .Cols - 1
            .TextMatrix(0, iCol) = "���̷�-�����ʿؼ�¼��"
        Next
        .Cell(flexcpFontSize, 0, .FixedCols, 0, .Cols - 1) = 18
        .Cell(flexcpFontBold, 0, .FixedCols, 0, .Cols - 1) = True
        .RowHeight(0) = 500
        .RowHeight(1) = 700
        .TextMatrix(2, mCol.����) = "����": .ColWidth(mCol.����) = 1100: .ColAlignment(mCol.����) = flexAlignCenterCenter
        .TextMatrix(2, mCol.����) = "����": .ColWidth(mCol.����) = 600: .ColAlignment(mCol.����) = flexAlignCenterCenter
        .TextMatrix(2, mCol.�ⶨֵ) = "�ⶨֵ": .ColWidth(mCol.�ⶨֵ) = 900: .ColAlignment(mCol.�ⶨֵ) = flexAlignCenterCenter
        .TextMatrix(2, mCol.��ֵ) = "��ֵ": .ColWidth(mCol.��ֵ) = 900: .ColAlignment(mCol.��ֵ) = flexAlignCenterCenter
        .TextMatrix(2, mCol.SD) = "SD": .ColWidth(mCol.SD) = 800: .ColAlignment(mCol.SD) = flexAlignCenterCenter
        .TextMatrix(2, mCol.SI����) = "SI����": .ColWidth(mCol.SI����) = 800: .ColAlignment(mCol.SI����) = flexAlignCenterCenter
        .TextMatrix(2, mCol.SI����) = "SI����": .ColWidth(mCol.SI����) = 800: .ColAlignment(mCol.SI����) = flexAlignCenterCenter
        .TextMatrix(2, mCol.N) = "n": .ColWidth(mCol.N) = 600: .ColAlignment(mCol.N) = flexAlignCenterCenter
        .TextMatrix(2, mCol.n3s) = "n3s": .ColWidth(mCol.n3s) = 600: .ColAlignment(mCol.n3s) = flexAlignCenterCenter
        .TextMatrix(2, mCol.n2s) = "n2s": .ColWidth(mCol.n2s) = 600: .ColAlignment(mCol.n2s) = flexAlignCenterCenter
        .TextMatrix(2, mCol.���) = "���": .ColWidth(mCol.���) = 500: .ColAlignment(mCol.���) = flexAlignCenterCenter
        .TextMatrix(2, mCol.������) = "������": .ColWidth(mCol.������) = 1100: .ColAlignment(mCol.������) = flexAlignCenterCenter
        
        .ColWidth(0) = 2000
    End With
End Sub

Public Sub ReportPrint(ByVal bytMode As Byte)
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��
    
    'д���ݵ���ʱ��
    Dim lngRow As Long, strSQL As String, lngResId As Long, lngCount As Long
    
    With Me.vfgData
        If .Rows <= 6 Then Exit Sub
        strSQL = "ZL_���̷���ӡ_Clear"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        For lngRow = 3 To Me.vfgData.Rows - 1
            If InStr(.TextMatrix(lngRow, mCol.����), "-") > 0 And _
               Val(.TextMatrix(lngRow, mCol.����)) > 0 And _
               Val(.TextMatrix(lngRow, mCol.����)) < 21 Then
                strSQL = "ZL_���̷���ӡ_Insert('" & .TextMatrix(lngRow, mCol.����) & "','" & _
                                                .TextMatrix(lngRow, mCol.����) & "','" & _
                                                .TextMatrix(lngRow, mCol.�ⶨֵ) & "','" & _
                                                .TextMatrix(lngRow, mCol.��ֵ) & "','" & _
                                                .TextMatrix(lngRow, mCol.SD) & "','" & _
                                                .TextMatrix(lngRow, mCol.SI����) & "','" & _
                                                .TextMatrix(lngRow, mCol.SI����) & "','" & _
                                                .TextMatrix(lngRow, mCol.���) & "','" & _
                                                .TextMatrix(lngRow, mCol.������) & "')"
               zlDatabase.ExecuteProcedure strSQL, Me.Caption
           End If
        Next
    End With
    '-------------------------------------------------
    '���ô�ӡ��������
    lngResId = 0
    For lngCount = 0 To Me.cboQCitem.ListCount - 1
        If Me.cboQCitem.ListIndex = lngCount Then lngResId = Val(Me.cboQCitem.ItemData(lngCount)): Exit For
    Next
    
    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1209_7", Me, _
                    "�ʿ�ƷID=" & lngResId, _
                    "��ĿID=" & mlngItemID, _
                    "��ʼ����=" & mstrFromDate, _
                    "��������=" & mstrToDate, _
                    IIf(bytMode, 1, 2))

End Sub

Private Sub Form_Load()
    initVfgData
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim lngCount As Long

    With Me.vfgData
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - Me.cboQCitem.Height - Screen.TwipsPerPixelY * 4
    End With
    
    With Me.cboQCitem
        .Left = Me.ScaleLeft + Screen.TwipsPerPixelX * 2
        .Top = Me.ScaleHeight - .Height - Screen.TwipsPerPixelY * 2
    End With
    
    With Me.opt�ʿ�Ʒ(0)
        .Left = Me.ScaleLeft + Screen.TwipsPerPixelX * 2
        .Top = Me.ScaleHeight - .Height - Screen.TwipsPerPixelY * 2
    End With
    For lngCount = 1 To Me.opt�ʿ�Ʒ.Count
        With Me.opt�ʿ�Ʒ(lngCount)
            .Left = Me.opt�ʿ�Ʒ(lngCount - 1).Left + Me.opt�ʿ�Ʒ(lngCount - 1).Width + Screen.TwipsPerPixelX * 10
            .Top = Me.opt�ʿ�Ʒ(lngCount - 1).Top
        End With
    Next
    
End Sub

Private Sub opt�ʿ�Ʒ_Click(Index As Integer)
    If Me.Visible = False Then Exit Sub
    If Me.opt�ʿ�Ʒ(Index).Enabled = False Then Exit Sub
    If Me.Tag = "��ˢ��" Then Exit Sub
    Call RefGrid
End Sub

Private Sub cboQCitem_Click()
    If Me.Visible = False Then Exit Sub
    If Me.Tag = "��ˢ��" Then Exit Sub
    Call RefGrid
End Sub
