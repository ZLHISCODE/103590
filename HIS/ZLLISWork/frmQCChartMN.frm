VERSION 5.00
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmQCChartMN 
   BorderStyle     =   0  'None
   Caption         =   "Monicaͼ"
   ClientHeight    =   5340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.ComboBox cboQCitem 
      Height          =   300
      Left            =   2970
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4950
      Width           =   2595
   End
   Begin VB.OptionButton opt�ʿ�Ʒ 
      Caption         =   "473843A��ֵ�ʿ�Ʒ"
      Enabled         =   0   'False
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   4905
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   2475
   End
   Begin C1Chart2D8.Chart2D chtThis 
      Height          =   4410
      Left            =   90
      TabIndex        =   1
      Top             =   150
      Width           =   7005
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   12356
      _ExtentY        =   7779
      _StockProps     =   0
      ControlProperties=   "frmQCChartMN.frx":0000
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmQCChartMN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrResList As String
Private mlngItemID As Long
Private mstrFromDate As String
Private mstrToDate As String

Dim lngCount As Long
Private mstr�ʿ�Ʒ���� As String
'--------------------------------------------
'����Ϊ���幫������
'--------------------------------------------
Public Sub ChartPrint()
    With Me.chtThis
'        .PrintChart oc2dFormatBitmap, oc2dScaleToFit, 0, 0, 0, 0
        .Save App.path & "\QC_Tmp0"
    End With
End Sub

Public Sub ChartSaveAs()
    Dim strBatCode As String
    For lngCount = 0 To Me.cboQCitem.ListCount - 1
        If Me.cboQCitem.ListIndex = lngCount Then strBatCode = Me.cboQCitem.Text: Exit For
    Next
    With Me.comDlg
        .CancelError = True
        .DialogTitle = "���Ϊ"
        .filter = "(ͼ���ļ�)|*.jpg"
        .FileName = strBatCode & Me.Caption & Format(mstrToDate, "yyyyMMdd") & ".jpg"
        Err = 0: On Error Resume Next
        .ShowSave
        If Err <> 0 Then Exit Sub
        If .FileName = "" Then Exit Sub
        Me.chtThis.SaveImageAsJpeg .FileName, 100, False, False, False
    End With
End Sub

Public Sub ChartCopy()
    Me.chtThis.CopyToClipboard (oc2dFormatBitmap)
End Sub

Public Function zlRefresh(strResList As String, lngItemID As Long, strFromDate As String, strToDate As String, str�ʿ�Ʒ���� As String) As Boolean
    '���ܣ�ˢ�±������������ʾ����
    '������ strResList  ��ǰѡ����ʿ�Ʒid�����Զ��ŷָ�
    '       lngItemId   ��ǰ��Ŀid
    '       strFromDate ��ʼ����
    '       strToDate   ��������
    Dim rsTemp As New adodb.Recordset
    Dim intCounts As Integer
    Dim lngResId As Long
    
    mstrResList = strResList
    mlngItemID = lngItemID
    mstrFromDate = strFromDate
    mstrToDate = strToDate
    mstr�ʿ�Ʒ���� = str�ʿ�Ʒ����
    lngResId = 0
    Me.Tag = "��ˢ��"
    intCounts = Me.cboQCitem.ListCount
    For lngCount = intCounts - 1 To 1 Step -1
        If Me.cboQCitem.ListIndex = lngCount Then lngResId = Val(Me.cboQCitem.ItemData(lngCount))
'        Unload Me.opt�ʿ�Ʒ(Me.opt�ʿ�Ʒ.UBound)
    Next
    cboQCitem.Clear
    
    Me.opt�ʿ�Ʒ(0).Enabled = False
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select ID, ���� || '-' || ���� As �ʿ�Ʒ From �����ʿ�Ʒ Where Instr(',' || [1] || ',', ',' || ID || ',') > 0"
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
    Call Form_Resize
    Call RefChart
    
    zlRefresh = True
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub RefChart()
    '���ܣ�ˢ��ͼ����ʾ
    Dim rsTemp As New adodb.Recordset
    Dim lngResId As Long, strLable As String, strUnit As String
    Dim dblAvg As Double, dblSD As Double, dblMax As Double
    Dim aryX() As Variant, aryY() As Variant, ary2() As Variant, lngHoles As Long
    
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
    
    '����ͼ�εĻ�����̬
    Me.chtThis.Reset
    Me.chtThis.AllowUserChanges = False
    Me.chtThis.ChartGroups(1).Data.NumSeries = 0
    With Me.chtThis.ChartArea
        .Axes("Y").Min = 0: .Axes("Y").Max = 1
        .Axes("X").Min = 0: .Axes("X").Max = 1
    End With
    With Me.chtThis.ChartGroups(1)
        .ChartType = oc2dTypePlot
        With .Data
            .LayOut = oc2dDataArray
            .NumSeries = 8
            .NumPoints(1) = 0
        End With
        .Styles(1).Symbol.Shape = oc2dShapeNone: .Styles(1).Line.COLOR = RGB(0, 0, 0)
        .Styles(2).Symbol.Shape = oc2dShapeNone: .Styles(2).Line.COLOR = RGB(200, 200, 0)
        .Styles(3).Symbol.Shape = oc2dShapeNone: .Styles(3).Line.COLOR = RGB(200, 200, 0)
        .Styles(4).Symbol.Shape = oc2dShapeNone: .Styles(4).Line.COLOR = RGB(255, 0, 0)
        .Styles(5).Symbol.Shape = oc2dShapeNone: .Styles(5).Line.COLOR = RGB(255, 0, 0)
        .Styles(6).Symbol.Shape = oc2dShapeOpenDiamond: .Styles(6).Line.Pattern = oc2dLineNone: .Styles(6).Symbol.COLOR = RGB(0, 64, 64)
        .Styles(7).Symbol.Shape = oc2dShapeDiamond: .Styles(7).Line.Pattern = oc2dLineNone: .Styles(7).Symbol.COLOR = RGB(0, 64, 64)
        .Styles(8).Symbol.Shape = oc2dShapeNone: .Styles(8).Line.COLOR = RGB(0, 0, 160): .Styles(8).Symbol.COLOR = RGB(0, 0, 160)
    End With
    With Me.chtThis.ChartGroups(2)
        .ChartType = oc2dTypeHiLo
        With .Data
            .LayOut = oc2dDataArray
            .NumSeries = 2
            .NumPoints(1) = 0
        End With
        .Styles(1).Line.COLOR = RGB(0, 64, 64)
    End With
    
    '��û�����������Ϣ
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select RPad('��λ��' || '" & gstrUnitName & "', 46, ' ') || '���ڷ�Χ��' As ��0," & vbNewLine & _
            "       RPad('������' || D.����, 46, ' ') ||" & vbNewLine & _
            "        RPad('�ο���ֵ��' || Replace(Replace(' 0' || X.��ֵ, ' 0.', '0.'), ' 0', ''), 26, ' ') || '��ⷽ����' || L.���� As ��1," & vbNewLine & _
            "       RPad('��Ŀ��' || I.������ || ',' || I.Ӣ����, 46, ' ') ||" & vbNewLine & _
            "        RPad('�ο�SDֵ��' || Replace(Replace(' 0' || X.Sd, ' 0.', '0.'), ' 0', ''), 26, ' ') || '�Լ���Դ��' || M.�Լ� As ��2," & vbNewLine & _
            "       RPad('�ʿ�Ʒ��' || M.���� || ',' || M.����, 46, ' ') ||" & vbNewLine & _
            "        RPad('�ο�CCV%��' || Replace(Replace(' 0' || X.Cv, ' 0.', '0.'), ' 0', ''), 26, ' ') || 'У׼����Դ��' || M.У׼�� As ��3," & vbNewLine & _
            "       X.��ֵ, X.Sd, I.��λ" & vbNewLine & _
            "From �������� D, �����ʿ�Ʒ M, �����ʿ�Ʒ��Ŀ X, ����������Ŀ I,�����ʿ�Ʒ��Ŀ L " & vbNewLine & _
            "Where D.ID = M.����id And M.ID = X.�ʿ�Ʒid And X.��Ŀid = I.ID And M.ID = [1] And X.��Ŀid = [2] " & vbNewLine & _
            " And M.ID = L.�ʿ�ƷID And L.��ĿID = [2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResId, mlngItemID)
    If rsTemp.RecordCount <= 0 Then Me.chtThis.Header.Text = "���ʿ�Ʒ��Ϣ��ȫ�棡": Exit Sub
    strLable = rsTemp!��0 & Format(mstrFromDate, "yyyy��MM��dd��") & "��" & Format(mstrToDate, "yyyy��MM��dd��")
    strLable = strLable & vbCrLf & rsTemp!��1 & vbCrLf & rsTemp!��2 & vbCrLf & rsTemp!��3
    dblAvg = Val("" & rsTemp!��ֵ): dblSD = Val("" & rsTemp!SD): strUnit = "" & rsTemp!��λ
    If dblAvg = 0 Or dblSD = 0 Then
        
        'MsgBox "û�����òο���ֵ��CCV���޷�����" & Me.Caption & "��", vbInformation, gstrSysName: Exit Sub
        Me.chtThis.Header.Text = "û�����òο���ֵ��CCV���޷�����" & Me.Caption & "��": Exit Sub
    End If
    
    '�����XY������
    With Me.chtThis.Header
        .Text = strLable
        .Adjust = oc2dAdjustLeft
    End With
    With Me.chtThis.ChartArea.Axes("Y")
        .MajorGrid.Spacing.IsDefault = False
        .AnnotationMethod = oc2dAnnotateValues
        .Title.Text = "�ⶨֵ" & IIf(strUnit = "", "", "(" & strUnit & ")")
    End With
    With Me.chtThis.ChartArea.Axes("Y2")
        .AnnotationMethod = oc2dAnnotateValueLabels   '������2��ʾֵ��ʾ
        .Title.Text = "������"
        .Multiplier = 1
        With .ValueLabels
            .RemoveAll
            .Add Val(dblAvg), "T         =" & Format(Val(dblAvg), "0.00")
            .Add Val(dblAvg) + 0.8 * Val(dblSD), "T+0.8CCV*T=" & Format(Val(dblAvg) + 0.8 * Val(dblSD), "0.00")
            .Add Val(dblAvg) - 0.8 * Val(dblSD), "T-0.8CCV*T=" & Format(Val(dblAvg) - 0.8 * Val(dblSD), "0.00")
            .Add Val(dblAvg) + 1.5 * Val(dblSD), "T+1.5CCV*T=" & Format(Val(dblAvg) + 1.5 * Val(dblSD), "0.00")
            .Add Val(dblAvg) - 1.5 * Val(dblSD), "T-1.5CCV*T=" & Format(Val(dblAvg) - 1.5 * Val(dblSD), "0.00")
        End With
    End With
    With Me.chtThis.ChartArea.Axes("X")
        .MajorGrid.Spacing.IsDefault = False
        .AnnotationMethod = oc2dAnnotateValueLabels   '��������ʾֵ��ʾ
        .Title.Text = "����"
        .AnnotationRotationAngle = 30
    End With
    
    '������֯
    gstrSql = "Select ����ʱ��, Max(Decode(����, '1-', ���, 0)) As ���1, Max(Decode(����, '1-', 0, ���)) As ���2" & vbNewLine & _
            "From (Select Q.����ʱ��, Q.���Դ��� || '-' || Decode(Nvl(T.���,0),2,2,Null) As ����," & vbNewLine & _
            "              zl_Lis_toNumber(Q.�ʿ�Ʒid,R.������Ŀid, R.������,R.id) As ���" & vbNewLine & _
            "       From �����ʿؼ�¼ Q, ������ͨ��� R,�����ʿر��� T, �����ʿ�Ʒ M, �����ʿؾ�ֵ X " & vbNewLine & _
            "       Where Q.�걾id = R.����걾id And /*Nvl(R.�Ƿ����, 0) = 1 And*/ Q.�ʿ�Ʒid + 0 = [1] And R.������Ŀid + 0 = [2] And" & vbNewLine & _
            "             Nvl(R.���ý��,0)=0 And R.ID=T.���ID(+) And (Q.����ʱ�� Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd'))" & vbNewLine & _
            "             And (Q.����ʱ�� Between X.��ʼ���� And NVL(X.��������,M.��������)) And " & vbNewLine & _
            "              Q.�ʿ�Ʒid=M.id And M.id=X.�ʿ�Ʒid  And  X.��ĿID = [2] And " & vbNewLine & _
            "             Instr(';'||[5]||';',';' || X.�ʿ�Ʒid||'='||To_char(X.��ʼ����,'yyyy-MM-dd')||','||to_char(Nvl(X.��������, M.��������),'yyyy-mm-dd')||';' ) > 0" & vbNewLine & _
            "       )" & vbNewLine & _
            "Group By ����ʱ��" & vbNewLine & _
            "Order By ����ʱ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResId, mlngItemID, mstrFromDate, mstrToDate, mstr�ʿ�Ʒ����)
    
    Me.chtThis.ChartArea.Axes("X").ValueLabels.RemoveAll
    lngHoles = 0
    With rsTemp
        ReDim aryX(.RecordCount)
        ReDim aryY(.RecordCount, 7)
        ReDim ary2(.RecordCount, 1)
        aryY(0, 0) = Val(dblAvg)
        aryY(0, 1) = Val(dblAvg) + 0.8 * Val(dblSD)
        aryY(0, 2) = Val(dblAvg) - 0.8 * Val(dblSD)
        aryY(0, 3) = Val(dblAvg) + 1.5 * Val(dblSD)
        aryY(0, 4) = Val(dblAvg) - 1.5 * Val(dblSD)
        aryY(0, 5) = Me.chtThis.ChartGroups(1).Data.HoleValue
        aryY(0, 6) = Me.chtThis.ChartGroups(1).Data.HoleValue
        aryY(0, 7) = Me.chtThis.ChartGroups(1).Data.HoleValue
        ary2(0, 0) = Me.chtThis.ChartGroups(1).Data.HoleValue
        ary2(0, 1) = Me.chtThis.ChartGroups(1).Data.HoleValue
        dblMax = 3 * Val(dblSD)
        Do While Not .EOF
            Me.chtThis.ChartArea.Axes("X").ValueLabels.Add .AbsolutePosition, Format(!����ʱ��, "M��d��")
            aryX(.AbsolutePosition) = .AbsolutePosition
            aryY(.AbsolutePosition, 0) = Val(dblAvg)
            aryY(.AbsolutePosition, 1) = Val(dblAvg) + 0.8 * Val(dblSD)
            aryY(.AbsolutePosition, 2) = Val(dblAvg) - 0.8 * Val(dblSD)
            aryY(.AbsolutePosition, 3) = Val(dblAvg) + 1.5 * Val(dblSD)
            aryY(.AbsolutePosition, 4) = Val(dblAvg) - 1.5 * Val(dblSD)
            If Val("" & !���1) = 0 Then
                aryY(.AbsolutePosition, 5) = Me.chtThis.ChartGroups(1).Data.HoleValue
            Else
                aryY(.AbsolutePosition, 5) = Val("" & !���1)
                If dblMax < Abs(Val(aryY(.AbsolutePosition, 5)) - Val(dblAvg)) Then dblMax = Abs(Val(aryY(.AbsolutePosition, 5)) - Val(dblAvg))
            End If
            If Val("" & !���2) = 0 Then
                aryY(.AbsolutePosition, 6) = Me.chtThis.ChartGroups(1).Data.HoleValue
            Else
                aryY(.AbsolutePosition, 6) = Val("" & !���2)
                If dblMax < Abs(Val(aryY(.AbsolutePosition, 6)) - Val(dblAvg)) Then dblMax = Abs(Val(aryY(.AbsolutePosition, 6)) - Val(dblAvg))
            End If
            If Val("" & !���1) = 0 Or Val("" & !���2) = 0 Then
                aryY(.AbsolutePosition, 7) = Me.chtThis.ChartGroups(1).Data.HoleValue: lngHoles = lngHoles + 1
            Else
                aryY(.AbsolutePosition, 7) = (aryY(.AbsolutePosition, 5) + aryY(.AbsolutePosition, 6)) / 2
            End If
            ary2(.AbsolutePosition, 0) = aryY(.AbsolutePosition, 5)
            ary2(.AbsolutePosition, 1) = aryY(.AbsolutePosition, 6)
            .MoveNext
        Loop
    End With
    If lngHoles > 3 Then
        Me.chtThis.Footer.Text = "ע�����ڸ��ʿ�Ʒ��" & lngHoles & "��û��ͬʱ�������β��ԣ�Ӱ���˸ÿ���ͼ�ı��֡�"
        Me.chtThis.Footer.Adjust = oc2dAdjustLeft
    Else
        Me.chtThis.Footer.Text = ""
    End If

    '���ˢ���ڲ�����
    With Me.chtThis
        .IsBatched = True
        With .ChartGroups(1).Data
            .NumPoints(1) = UBound(aryX) + 1
            Call .CopyXVectorIn(1, aryX)
            Call .CopyYArrayIn(aryY)
        End With
        With .ChartArea.Axes("Y")
            .Min = Val(dblAvg) - Val(dblMax)
            .Max = Val(dblAvg) + Val(dblMax)
        End With
        With .ChartArea.Axes("X")
            .Min = 0
            .Max = aryX(UBound(aryX))
        End With
        With .ChartGroups(2).Data
            .NumPoints(1) = UBound(aryX) + 1
            Call .CopyXVectorIn(1, aryX)
            Call .CopyYArrayIn(ary2)
        End With
        .IsBatched = False
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ChtThis_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim px As Long
    Dim py As Long
    Dim Series As Long
    Dim Point As Long
    Dim Distance As Long
    Dim Region As Long
    
    On Error Resume Next
    
    px = x / Screen.TwipsPerPixelX
    py = Y / Screen.TwipsPerPixelY
    
    If (Button = 0) Then
        With chtThis
            Region = .ChartGroups(1).CoordToDataIndex(px, py, oc2dFocusXY, Series, Point, Distance)
            If (Series > 0 And Point > 0) And (Distance <= 5) Then
                If (Region = oc2dRegionInChartArea) Then
                    .ToolTipText = .ChartGroups(1).Data(Series, Point)
                End If
            Else
                .ToolTipText = ""
                .Footer.Text = ""
            End If
            .Refresh
        End With
    End If
End Sub

'--------------------------------------------
'����Ϊ�ؼ��¼�����
'--------------------------------------------
Private Sub opt�ʿ�Ʒ_Click(Index As Integer)
    If Me.Visible = False Then Exit Sub
    If Me.opt�ʿ�Ʒ(Index).Enabled = False Then Exit Sub
    If Me.Tag = "��ˢ��" Then Exit Sub
    Call RefChart
End Sub

Private Sub cboQCitem_Click()
    If Me.Visible = False Then Exit Sub
    If Me.Tag = "��ˢ��" Then Exit Sub
    Call RefChart
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With Me.chtThis
        .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
        .Top = Me.ScaleTop: .Height = Me.ScaleHeight - Me.cboQCitem.Height - Screen.TwipsPerPixelY * 4
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

Public Function ZLGetMN_QCID() As Long
    '����       �õ���ǰʹ�õ��ʿ�Ʒ��ID
    For lngCount = 0 To Me.cboQCitem.ListCount - 1
        If Me.cboQCitem.ListIndex = lngCount Then ZLGetMN_QCID = Val(Me.cboQCitem.ItemData(lngCount)): Exit For
    Next
End Function

