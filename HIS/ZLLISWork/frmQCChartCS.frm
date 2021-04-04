VERSION 5.00
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmQCChartCS 
   BorderStyle     =   0  'None
   Caption         =   "�ۻ���ͼ"
   ClientHeight    =   5355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.ComboBox cboQCitem 
      Height          =   300
      Left            =   2730
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4950
      Width           =   2595
   End
   Begin VB.CheckBox chkUnion 
      Caption         =   "����Levey_Jenningsͼ"
      Height          =   180
      Left            =   5475
      TabIndex        =   2
      Top             =   5040
      Width           =   2115
   End
   Begin VB.OptionButton opt�ʿ�Ʒ 
      Caption         =   "473843A��ֵ�ʿ�Ʒ"
      Enabled         =   0   'False
      Height          =   180
      Index           =   0
      Left            =   255
      TabIndex        =   0
      Top             =   4980
      Width           =   2475
   End
   Begin C1Chart2D8.Chart2D chtThis 
      Height          =   4020
      Left            =   180
      TabIndex        =   1
      Top             =   165
      Width           =   7020
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   12382
      _ExtentY        =   7091
      _StockProps     =   0
      ControlProperties=   "frmQCChartCS.frx":0000
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmQCChartCS"
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
Dim mstr�ʿ�Ʒ���� As String

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
    gstrSql = "Select ID, ���� || '-' || ���� As �ʿ�Ʒ From �����ʿ�Ʒ Where Instr(',' || [1] || ',', ',' || ID || ',') > 0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strResList)
    With rsTemp
        Do While Not .EOF
            If .AbsolutePosition > Me.cboQCitem.ListCount Then cboQCitem.AddItem "" & !�ʿ�Ʒ
'            If .AbsolutePosition <> 1 Then Load Me.chtThis(.AbsolutePosition - 1)
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
    Call Form_Resize
    Me.Tag = ""
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
    Dim lngAllTimes As Long, lngSelTimes As Long
    Dim dblAvg As Double, dblSD As Double, dblMax As Double, dblK As Double, dblH As Double '��صĿ���ֵ����
    Dim aryX() As Variant, aryY() As Variant, arySum() As Variant
    
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
    With Me.chtThis
        .IsBatched = True
        .Reset
        .AllowUserChanges = False
        .ChartArea.Axes("Y").Min = 0: .ChartArea.Axes("Y").Max = 1
        .ChartArea.Axes("X").Min = 0: .ChartArea.Axes("X").Max = 1
        With .ChartGroups(1)
            .ChartType = oc2dTypePlot
            With .Data
                .NumSeries = 0
                .LayOut = oc2dDataArray
                .NumSeries = 7
                .NumPoints(1) = 0
            End With
        .Styles(1).Symbol.Shape = oc2dShapeNone: .Styles(1).Line.COLOR = RGB(0, 0, 0)
        .Styles(2).Symbol.Shape = oc2dShapeNone: .Styles(2).Line.COLOR = RGB(0, 128, 0)
        .Styles(3).Symbol.Shape = oc2dShapeNone: .Styles(3).Line.COLOR = RGB(0, 128, 0)
        .Styles(4).Symbol.Shape = oc2dShapeNone: .Styles(4).Line.COLOR = RGB(255, 0, 0)
        .Styles(5).Symbol.Shape = oc2dShapeNone: .Styles(5).Line.COLOR = RGB(255, 0, 0)
        .Styles(6).Symbol.Shape = oc2dShapeDot: .Styles(6).Line.COLOR = RGB(0, 0, 160): .Styles(6).Symbol.COLOR = RGB(0, 0, 160)
        .Styles(7).Symbol.Shape = oc2dShapeDiamond: .Styles(7).Line.COLOR = RGB(0, 128, 255): .Styles(7).Symbol.COLOR = RGB(0, 128, 255)
        End With
        .IsBatched = False
    End With
    Call chkUnion_Click
    
    '��û�����������Ϣ
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select RPad('��λ��' || '" & gstrUnitName & "', 46, ' ') || '���ڣ�' As ��0," & vbNewLine & _
            "       RPad('������' || D.����, 46, ' ') ||" & vbNewLine & _
            "        RPad('��ֵ��' || Replace(Replace(' 0' || X.��ֵ, ' 0.', '0.'), ' 0', ''), 26, ' ') || '��ⷽ����' || L.���� As ��1," & vbNewLine & _
            "       RPad('��Ŀ��' || I.������ || ',' || I.Ӣ����, 46, ' ') ||" & vbNewLine & _
            "        RPad('SDֵ��' || Replace(Replace(' 0' || X.Sd, ' 0.', '0.'), ' 0', ''), 26, ' ') || '�Լ���Դ��' || M.�Լ� As ��2," & vbNewLine & _
            "       RPad('�ʿ�Ʒ��' || M.���� || ',' || M.����, 46, ' ') || RPad('����' || R.����, 26, ' ') || 'У׼����Դ��' ||" & vbNewLine & _
            "        M.У׼�� As ��3, X.��ֵ, X.Sd, I.��λ, R.K, R.H" & vbNewLine & _
            "From �������� D, �����ʿ�Ʒ M, �����ʿؾ�ֵ X, ����������Ŀ I, ������������ A, �����ʿع��� R,�����ʿ�Ʒ��Ŀ L " & vbNewLine & _
            "Where D.ID = M.����id And M.ID = X.�ʿ�Ʒid And X.��Ŀid = I.ID And D.ID = A.����id And A.����id = R.ID And R.���� = 3 And" & vbNewLine & _
            "      M.ID = [1] And X.��Ŀid = [2] And M.ID = L.�ʿ�ƷID and L.��ĿID = [2] And " & vbNewLine & _
            "      Instr(';' || [3] || ';',';' || X.�ʿ�Ʒid||'='||To_char(X.��ʼ����,'yyyy-MM-dd')||','||to_char(Nvl(X.��������, M.��������),'yyyy-mm-dd')||';' ) > 0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResId, mlngItemID, mstr�ʿ�Ʒ����)
    If rsTemp.RecordCount <= 0 Then Me.chtThis.Header.Text = "���ʿ�Ʒ��Ϣ��ȫ�棡": Exit Sub
    strLable = rsTemp!��0 & Format(mstrFromDate, "yyyy��MM��dd��") & "��" & Format(mstrToDate, "yyyy��MM��dd��")
    strLable = strLable & vbCrLf & rsTemp!��1 & vbCrLf & rsTemp!��2 & vbCrLf & rsTemp!��3
    dblAvg = Val("" & rsTemp!��ֵ): dblSD = Val("" & rsTemp!SD): strUnit = "" & rsTemp!��λ
    dblK = rsTemp!k: dblH = rsTemp!H
    If dblAvg = 0 Or dblSD = 0 Then
         Me.chtThis.Header.Text = "��δ��ֵ���޷�����" & Me.Caption & "��": Exit Sub
    End If
    
    '���⡢XY������
    With Me.chtThis.Header
        .Text = strLable
        .Adjust = oc2dAdjustLeft
    End With
    With Me.chtThis.ChartArea.Axes("Y")
        .MajorGrid.Spacing.IsDefault = False
        .AnnotationMethod = oc2dAnnotateValues
        .Title.Text = "�ۻ���" & IIf(strUnit = "", "", "(" & strUnit & ")")
    End With
    With Me.chtThis.ChartArea.Axes("Y2")
        .AnnotationMethod = oc2dAnnotateValueLabels
        .Title.Text = "������"
        .Multiplier = 1
        With .ValueLabels
            .RemoveAll
            .Add 0, "0"
            .Add 0 + Val(dblK) * Val(dblSD), "Ku= " & Format(Val(dblK) * Val(dblSD), "0.00")
            .Add 0 - Val(dblK) * Val(dblSD), "Kl=" & Format(-Val(dblK) * Val(dblSD), "0.00")
            .Add 0 + Val(dblH) * Val(dblSD), "Hu= " & Format(Val(dblH) * Val(dblSD), "0.00")
            .Add 0 - Val(dblH) * Val(dblSD), "Hl=" & Format(-Val(dblH) * Val(dblSD), "0.00")
        End With
    End With
    With Me.chtThis.ChartArea.Axes("X")
        .MajorGrid.Spacing.IsDefault = False
        .AnnotationMethod = oc2dAnnotateValueLabels   '��������ʾֵ��ʾ
        .AnnotationPlacement = oc2dAnnotateMinimum
        .Title.Text = "���Դ���"
    End With
    
    '������֯
    gstrSql = "Select ����ʱ��, ����, Nvl(���, 0) As ���" & vbNewLine & _
            "From (Select Q.����ʱ��, To_Char(Q.���Դ���, '000') || '-' || Decode(Nvl(T.���, 0), 2, Q.���Դ���, 999) As ����," & vbNewLine & _
            "              zl_Lis_ToNumber(Q.�ʿ�Ʒid,R.������Ŀid,R.������,R.id) As ���" & vbNewLine & _
            "       From �����ʿؼ�¼ Q, ������ͨ��� R,�����ʿر��� T,�����ʿ�Ʒ M,�����ʿؾ�ֵ X " & vbNewLine & _
            "       Where Q.�걾id = R.����걾id And Nvl(R.���ý��,0)=0 And /*Nvl(R.�Ƿ����, 0) = 1 And*/ Q.�ʿ�Ʒid = [1] And R.������Ŀid + 0 = [2] And" & vbNewLine & _
            "             R.ID=T.���ID(+) And Q.����ʱ�� + 0 <= To_Date([3], 'yyyy-MM-dd')" & vbNewLine & _
            "       And Instr(';' || [4] || ';',';' || X.�ʿ�Ʒid||'='||To_char(X.��ʼ����,'yyyy-MM-dd')||','||to_char(Nvl(X.��������, M.��������),'yyyy-mm-dd')||';' ) > 0" & vbNewLine & _
            "       And Q.�ʿ�Ʒid=M.ID And M.id=X.�ʿ�ƷID and  X.��Ŀid = [2] And Q.����ʱ�� between X.��ʼ���� and Nvl(X.��������, M.��������) " & vbNewLine & _
            "      )Order By ����ʱ��, ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResId, mlngItemID, mstrToDate, mstr�ʿ�Ʒ����)
    '���ȼ����ۻ���
    With rsTemp
        ReDim arySum(.RecordCount, 1)
        arySum(0, 0) = 0: arySum(0, 1) = 0: lngSelTimes = 0: lngAllTimes = .RecordCount
        Do While Not .EOF
            If Val("" & !���) = 0 Then
                arySum(.AbsolutePosition, 0) = 0
            ElseIf Abs(Val("" & !���) - Val(dblAvg)) <= Val(dblK) * Val(dblSD) Then
                arySum(.AbsolutePosition, 0) = 0
            ElseIf Sgn(Val("" & !���) - Val(dblAvg)) <> Sgn(arySum(.AbsolutePosition - 1, 0)) Then
                arySum(.AbsolutePosition, 0) = Sgn(Val("" & !���) - Val(dblAvg)) * (Abs(Val("" & !���) - Val(dblAvg)) - Val(dblK) * Val(dblSD))
            Else
                arySum(.AbsolutePosition, 0) = arySum(.AbsolutePosition - 1, 0) + Sgn(Val("" & !���) - Val(dblAvg)) * (Abs(Val("" & !���) - Val(dblAvg)) - Val(dblK) * Val(dblSD))
            End If
            If Format(!����ʱ��, "yyyy-MM-dd") >= mstrFromDate Then lngSelTimes = lngSelTimes + 1
            arySum(.AbsolutePosition, 1) = Val("" & !���) - Val(dblAvg)
            .MoveNext
        Loop
    End With
    
    '�����䷶Χ�����ݸ����ͼ����
    Me.chtThis.ChartArea.Axes("X").ValueLabels.RemoveAll
    ReDim aryX(lngSelTimes)
    ReDim aryY(lngSelTimes, 6)
    aryY(0, 0) = 0
    aryY(0, 1) = 0 + Val(dblK) * Val(dblSD)
    aryY(0, 2) = 0 - Val(dblK) * Val(dblSD)
    aryY(0, 3) = 0 + Val(dblH) * Val(dblSD)
    aryY(0, 4) = 0 - Val(dblH) * Val(dblSD)
    aryY(0, 5) = Me.chtThis.ChartGroups(1).Data.HoleValue
    aryY(0, 6) = Me.chtThis.ChartGroups(1).Data.HoleValue
    dblMax = Val(dblH) * Val(dblSD)
    For lngCount = 1 To lngSelTimes
        Me.chtThis.ChartArea.Axes("X").ValueLabels.Add lngCount, lngCount
        aryX(lngCount) = lngCount
        aryY(lngCount, 0) = 0
        aryY(lngCount, 1) = 0 + Val(dblK) * Val(dblSD)
        aryY(lngCount, 2) = 0 - Val(dblK) * Val(dblSD)
        aryY(lngCount, 3) = 0 + Val(dblH) * Val(dblSD)
        aryY(lngCount, 4) = 0 - Val(dblH) * Val(dblSD)
        aryY(lngCount, 5) = arySum(lngAllTimes - lngSelTimes + lngCount, 0)
        aryY(lngCount, 6) = arySum(lngAllTimes - lngSelTimes + lngCount, 1)
        If dblMax < Abs(arySum(lngAllTimes - lngSelTimes + lngCount, 0)) Then dblMax = Abs(arySum(lngAllTimes - lngSelTimes + lngCount, 0))
        If dblMax < Abs(arySum(lngAllTimes - lngSelTimes + lngCount, 1)) Then dblMax = Abs(arySum(lngAllTimes - lngSelTimes + lngCount, 1))
    Next

    '���ˢ���ڲ�����
    With Me.chtThis
        .IsBatched = True
        With .ChartGroups(1).Data
            .NumPoints(1) = UBound(aryX) + 1
            Call .CopyXVectorIn(1, aryX)
            Call .CopyYArrayIn(aryY)
        End With
        With .ChartArea.Axes("Y")
            .Min = 0 - Val(dblMax) - 0.01
            .Max = 0 + Val(dblMax) + 0.01
        End With
        With .ChartArea.Axes("X")
            .Min = 0: .Max = aryX(UBound(aryX))
        End With
        .IsBatched = False
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'--------------------------------------------
'����Ϊ�ؼ��¼�����
'--------------------------------------------
Private Sub chkUnion_Click()
    With Me.chtThis.ChartGroups(1)
        If Me.chkUnion.Value = vbChecked Then
            .Styles(7).Line.Pattern = oc2dLineDotted: .Styles(7).Symbol.Shape = oc2dShapeDiamond
        Else
            .Styles(7).Line.Pattern = oc2dLineNone: .Styles(7).Symbol.Shape = oc2dShapeNone
        End If
    End With
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
    
    
    With Me.chkUnion
        .Left = Me.ScaleWidth - Me.chkUnion.Width
        .Top = Me.opt�ʿ�Ʒ(0).Top
    End With
End Sub

Public Function ZLGetCS_QCID() As Long
    '����       �õ���ǰʹ�õ��ʿ�Ʒ��ID
    For lngCount = 0 To Me.cboQCitem.ListCount - 1
        If Me.cboQCitem.ListIndex = lngCount Then ZLGetCS_QCID = Val(Me.cboQCitem.ItemData(lngCount)): Exit For
    Next
End Function

