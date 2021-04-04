VERSION 5.00
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmQCChartLJ 
   BorderStyle     =   0  'None
   Caption         =   "Levey_Jenningsͼ"
   ClientHeight    =   5355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.ComboBox cboQCitem 
      Height          =   300
      Left            =   2700
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4980
      Width           =   2595
   End
   Begin VB.ComboBox cbo��ʾ 
      Height          =   300
      Left            =   5790
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4950
      Width           =   1785
   End
   Begin VB.OptionButton opt�ʿ�Ʒ 
      Caption         =   "473843A��ֵ�ʿ�Ʒ"
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   4920
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   2475
   End
   Begin C1Chart2D8.Chart2D chtThis 
      Height          =   4410
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   165
      Width           =   7365
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   12991
      _ExtentY        =   7779
      _StockProps     =   0
      ControlProperties=   "frmQCChartLJ.frx":0000
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   0
      Top             =   15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmQCChartLJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrResList As String
Private mlngItemID As Long
Private mstrFromDate As String
Private mstrToDate As String
Private mstr�ʿ�Ʒ���� As String

Dim lngCount As Long
Private mArr() As String
Private mbln�����ʿ�ͼ As Boolean
Private mint��λ��ʾ As Integer
Private mLastXY As String           '���������ʾ������ʱ�ظ�ˢ��

'--------------------------------------------
'����Ϊ���幫������
'--------------------------------------------
Public Function ChartPrint() As Integer
    '�����м���ͼƬ
    Dim intLoop As Integer
    Dim intIndex As Integer
    Dim intCount As Integer
    For intLoop = 0 To Me.cboQCitem.ListCount - 1
        If Me.cboQCitem.ListIndex = intLoop Then
            intIndex = intLoop
        End If
    Next
    For intLoop = 0 To chtThis.Count - 1
        With Me.chtThis(intLoop)
            If .Visible = True Then
    '        .PrintChart oc2dFormatBitmap, oc2dScaleToFit, 0, 0, 0, 0
                .Save App.path & "\QC_Tmp" & intCount
                intCount = intCount + 1
            End If
        End With
    Next
    ChartPrint = intCount
End Function


Public Sub ChartSaveAs()
    Dim strBatCode As String
    Dim intLoop As Integer
    Dim intIndex As Integer
    For intLoop = 0 To Me.cboQCitem.ListCount - 1
        If Me.cboQCitem.ListIndex = intLoop Then
            intIndex = intLoop
        End If
    Next
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
        Me.chtThis(intIndex).SaveImageAsJpeg .FileName, 100, False, False, False
    End With
End Sub

Public Sub ChartCopy()
    Dim intLoop As Integer
    Dim intIndex As Integer
    For intLoop = 0 To Me.cboQCitem.ListCount - 1
        If Me.cboQCitem.ListIndex = intLoop Then
            intIndex = intLoop
        End If
    Next
    Me.chtThis(intIndex).CopyToClipboard (oc2dFormatBitmap)
End Sub

Public Function zlRefresh(strResList As String, lngItemID As Long, strFromDate As String, strToDate As String, str�ʿ�Ʒ���� As String, Optional ByVal int��λ��ʾ = 1) As Boolean
    '���ܣ�ˢ�±������������ʾ����
    '������ strResList  ��ǰѡ����ʿ�Ʒid�����Զ��ŷָ�
    '       lngItemId   ��ǰ��Ŀid
    '       strFromDate ��ʼ����
    '       strToDate   ��������
    '       strDateSpace ��;�ָ����ʿ�Ʒ���ڼ�
    '
    Dim rsTemp As New adodb.Recordset
    Dim intCounts As Integer
    Dim lngResId As Long
    Dim intͼ��Index As Integer
    Dim int��ǰ�ʿ�Index As Integer
    
    mstrResList = strResList
    mlngItemID = lngItemID
    mstrFromDate = strFromDate
    mstrToDate = strToDate
    mstr�ʿ�Ʒ���� = str�ʿ�Ʒ����
    mint��λ��ʾ = int��λ��ʾ

    For lngCount = 0 To Me.cboQCitem.ListCount - 1
        If Me.cboQCitem.ListIndex = lngCount Then
            int��ǰ�ʿ�Index = lngCount
        End If
    Next
    
    intͼ��Index = Me.cbo��ʾ.ListIndex
    
    
    lngResId = 0
   
    intCounts = Me.cboQCitem.ListCount
    For lngCount = intCounts - 1 To 1 Step -1
'        If Me.opt�ʿ�Ʒ(lngCount).Value Then lngResId = Val(Me.opt�ʿ�Ʒ(Me.opt�ʿ�Ʒ.UBound).Tag)
'        Unload Me.opt�ʿ�Ʒ(Me.opt�ʿ�Ʒ.UBound)
        Unload chtThis(Me.chtThis.UBound)
    Next
    cboQCitem.Clear
    Me.opt�ʿ�Ʒ(0).Enabled = False
    Err = 0: On Error GoTo ErrHand
    mbln�����ʿ�ͼ = False
    
    gstrSql = "Select A.ID, A.���� || '-' || A.���� As �ʿ�Ʒ, B.�����ʿ�ͼ From �����ʿ�Ʒ A,�������� B Where A.����ID=B.ID(+) And Instr(',' || [1] || ',', ',' || A.ID || ',') > 0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strResList)
    With rsTemp
        Do While Not .EOF
            If .AbsolutePosition > Me.cboQCitem.ListCount Then cboQCitem.AddItem "" & !�ʿ�Ʒ
            If .AbsolutePosition <> 1 Then Load Me.chtThis(.AbsolutePosition - 1)
            cboQCitem.ItemData(cboQCitem.NewIndex) = !ID
'            If .AbsolutePosition > Me.opt�ʿ�Ʒ.Count Then Load Me.opt�ʿ�Ʒ(.AbsolutePosition - 1): Load Me.chtThis(.AbsolutePosition - 1)
'            Me.opt�ʿ�Ʒ(.AbsolutePosition - 1).Caption = "" & !�ʿ�Ʒ
'            Me.opt�ʿ�Ʒ(.AbsolutePosition - 1).Tag = !ID
'            Me.opt�ʿ�Ʒ(.AbsolutePosition - 1).Width = Me.TextWidth(Me.opt�ʿ�Ʒ(.AbsolutePosition - 1).Caption) + 360
'            Me.opt�ʿ�Ʒ(.AbsolutePosition - 1).Value = (lngResId = !ID)
'            Me.opt�ʿ�Ʒ(.AbsolutePosition - 1).Visible = True
'            Me.opt�ʿ�Ʒ(.AbsolutePosition - 1).Enabled = True
            mbln�����ʿ�ͼ = Val("" & !�����ʿ�ͼ) = 1
            .MoveNext
        Loop
    End With
    If rsTemp.RecordCount > 0 Then Me.cboQCitem.ListIndex = 0
    Call Form_Resize
    
    
    Me.cbo��ʾ.Clear
    Me.cbo��ʾ.Tag = "��ˢ��"
    For lngCount = 0 To Me.cboQCitem.ListCount - 1
        If lngCount < 3 Then
            Me.cbo��ʾ.AddItem lngCount + 1 & "��ͼ"
        Else
            Exit For
        End If
        If lngCount = int��ǰ�ʿ�Index Then
            Me.cboQCitem.ListIndex = lngCount
        End If
        
    Next
    If Me.cbo��ʾ.ListCount > 0 Then
        Me.cbo��ʾ.ListIndex = IIf(intͼ��Index = -1, 0, intͼ��Index)
    End If
    Me.cbo��ʾ.Tag = ""
    
    For lngCount = 0 To Me.cboQCitem.ListCount - 1
        Call RefChart(CInt(lngCount))
    Next
    
    Call Form_Resize
    DoEvents
    If Me.cboQCitem.ListIndex >= 0 Then
        Call chtThis_Resize(Me.cboQCitem.ListIndex, chtThis(Me.cboQCitem.ListIndex).Width, chtThis(Me.cboQCitem.ListIndex).Height)
    End If
    zlRefresh = True
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub RefChart(intIndex As Integer)
    '���ܣ�ˢ��ͼ����ʾ
    Dim rsTemp As New adodb.Recordset
    Dim lngResId As Long, strLable As String, strUnit As String
    Dim dblAvg As Double, dblSD As Double, dblMax As Double
    Dim aryX() As Variant, aryY() As Variant
    Dim strCalc As String           '������
    Dim strStartDate As String, strEndDate As String
    Dim str�������� As String '���泬�������޵����ݣ�������ʾ
    Dim intLoop As Integer, dateLoop As Date '���ڲ���30�������
    Dim lngX As Long '��¼X�����
    Dim bln�ϲ��� As Boolean, strС�� As String, lngTmp As Long, strTmp As String
    Dim strAllCount As String, strCurCount '���д���,��ǰ����
    lngResId = 0
'    For lngCount = 0 To Me.opt�ʿ�Ʒ.UBound
'        If Me.opt�ʿ�Ʒ(lngCount).Value Then lngResId = Val(Me.opt�ʿ�Ʒ(lngCount).Tag): Exit For
'    Next
    lngResId = Val(Me.cboQCitem.ItemData(intIndex))
    If lngResId = 0 Then
        Me.opt�ʿ�Ʒ(0).Enabled = False
        Me.opt�ʿ�Ʒ(0).Value = True
        lngResId = Val(Me.opt�ʿ�Ʒ(0).Tag)
        Me.opt�ʿ�Ʒ(0).Enabled = True
    End If
    
    '����ͼ�εĻ�����̬
    With Me.chtThis(intIndex)
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
                .NumSeries = 15
                .NumPoints(1) = 0
            End With
            .Styles(1).Symbol.Shape = oc2dShapeNone: .Styles(1).Line.COLOR = RGB(0, 0, 0)
            .Styles(2).Symbol.Shape = oc2dShapeNone: .Styles(2).Line.COLOR = RGB(0, 128, 0)
            .Styles(3).Symbol.Shape = oc2dShapeNone: .Styles(3).Line.COLOR = RGB(0, 128, 0)
            .Styles(4).Symbol.Shape = oc2dShapeNone: .Styles(4).Line.COLOR = RGB(200, 200, 0)
            .Styles(5).Symbol.Shape = oc2dShapeNone: .Styles(5).Line.COLOR = RGB(200, 200, 0)
            .Styles(6).Symbol.Shape = oc2dShapeNone: .Styles(6).Line.COLOR = RGB(255, 0, 0)
            .Styles(7).Symbol.Shape = oc2dShapeNone: .Styles(7).Line.COLOR = RGB(255, 0, 0)
            .Styles(8).Symbol.Shape = oc2dShapeNone: .Styles(8).Line.COLOR = RGB(0, 0, 0)
            .Styles(9).Symbol.Shape = oc2dShapeNone: .Styles(9).Line.COLOR = RGB(0, 0, 0)
            .Styles(10).Symbol.Shape = oc2dShapeDot: .Styles(10).Line.COLOR = RGB(0, 0, 160): .Styles(10).Symbol.COLOR = RGB(0, 0, 160)
            .Styles(11).Symbol.Shape = oc2dShapeDot: .Styles(11).Line.Pattern = oc2dLineNone: .Styles(11).Symbol.COLOR = RGB(255, 0, 0)
            .Styles(12).Symbol.Shape = oc2dShapeDot: .Styles(12).Line.Pattern = oc2dLineNone: .Styles(12).Symbol.COLOR = RGB(255, 0, 0)
            .Styles(13).Symbol.Shape = oc2dShapeDot: .Styles(13).Line.Pattern = oc2dLineNone: .Styles(13).Symbol.COLOR = RGB(255, 0, 0)
            .Styles(14).Symbol.Shape = oc2dShapeDot: .Styles(14).Line.Pattern = oc2dLineNone: .Styles(14).Symbol.COLOR = RGB(255, 0, 0)
            .Styles(15).Symbol.Shape = oc2dShapeDot: .Styles(15).Line.Pattern = oc2dLineNone: .Styles(15).Symbol.COLOR = RGB(255, 0, 0)
        End With
        .IsBatched = False
    End With
    
    '��û�����������Ϣ
    Err = 0: On Error GoTo ErrHand
'    gstrSql = "Select RPad('��λ��' || '" & gstrUnitName & "', 46, ' ') || '���ڣ�' As ��0," & vbNewLine & _
            "       RPad('������' || D.����, 46, ' ') ||" & vbNewLine & _
            "        RPad('��ֵ��' || Replace(Replace(' 0' || X.��ֵ, ' 0.', '0.'), ' 0', ''), 26, ' ') || '��ⷽ����' || L.���� As ��1," & vbNewLine & _
            "       RPad('��Ŀ��' || I.������ || ',' || I.Ӣ����, 46, ' ') ||" & vbNewLine & _
            "        RPad('SDֵ��' || Replace(Replace(' 0' || X.Sd, ' 0.', '0.'), ' 0', ''), 26, ' ') || '�Լ���Դ��' || D.�Լ���Դ As ��2," & vbNewLine & _
            "       RPad('�ʿ�Ʒ��' || M.���� || ',' || M.����, 46, ' ') ||" & vbNewLine & _
            "        RPad('CV% ��' || Replace(Replace(' 0' || X.Cv * 100, ' 0.', '0.'), ' 0', ''), 26, ' ') || 'У׼����Դ��' ||" & vbNewLine & _
            "        D.У׼����Դ As ��3, X.��ֵ, X.Sd, I.��λ" & vbNewLine & _
            "From �������� D, �����ʿ�Ʒ M, �����ʿؾ�ֵ X, ����������Ŀ I,�����ʿ�Ʒ��Ŀ L" & vbNewLine & _
            "Where D.ID = M.����id And M.ID = X.�ʿ�Ʒid And X.��Ŀid = I.ID And M.ID = [1] And X.��Ŀid = [2] And" & vbNewLine & _
            "      M.id = L.�ʿ�ƷID and L.��ĿID = [2] And " & vbNewLine & _
            "      (To_Date([3], 'yyyy-MM-dd') Between X.��ʼ���� And Nvl(X.��������, M.��������)) And" & vbNewLine & _
            "      (To_Date([4], 'yyyy-MM-dd') Between X.��ʼ���� And Nvl(X.��������, M.��������))"
                
'    gstrSql = " Select A.��ʼ����, Nvl(A.��������, B.��������) As ��������" & vbNewLine & _
'                " From �����ʿؾ�ֵ A, �����ʿ�Ʒ B" & vbNewLine & _
'                " Where A.�ʿ�Ʒid = B.ID And �ʿ�Ʒid = [1] And ��Ŀid = [2] And �ڼ� = [3] "
'    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResId, mlngItemId, mstrDateSpace)
'
'    If rsTemp.EOF = True Then
'        MsgBox "û���ҵ���Ӧ���ڼ�<" & mstrDateSpace & ">!", vbInformation, gstrSysName: Exit Sub
'    End If
    Dim varTmp As Variant, intCount As Integer
    
    If InStr(mstr�ʿ�Ʒ����, ";") > 0 Then
        varTmp = Split(mstr�ʿ�Ʒ����, ";")
        For intCount = LBound(varTmp) To UBound(varTmp)
            If lngResId = Val(Split(varTmp(intCount), "=")(0)) Then
                strStartDate = Split(Split(varTmp(intCount), "=")(1), ",")(0)
                strEndDate = Split(Split(varTmp(intCount), "=")(1), ",")(1)
                Exit For
            End If
        Next
    End If
    
        '������ǰ�ĵ��÷�ʽ
    If strStartDate = "" Then
        strStartDate = mstrFromDate: strEndDate = mstrToDate
    Else
        If CDate(strStartDate) < CDate(mstrFromDate) Then strStartDate = mstrFromDate
        If CDate(strEndDate) > CDate(mstrToDate) Then strEndDate = mstrToDate
    End If
    
'    If CDate(mstrFromDate) < CDate(rsTemp("��ʼ����")) Then
'        strStartDate = Nvl(rsTemp("��ʼ����"))
'    End If
'    If CDate(mstrToDate) > CDate(rsTemp("��������")) Then
'        strEndDate = Nvl(rsTemp("��������"))
'    End If

    strС�� = "0000"
'    gstrSql = "Select Nvl(С��λ��,2) As С�� From ����������Ŀ  A,�����ʿ�Ʒ M Where m.����id=A.����Id And m.Id=[1]  And a.��Ŀid=[2]"
'    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResId, mlngItemID)
'    Do Until rsTemp.EOF
'        strС�� = String(Val("" & rsTemp!С��), "0")
'        rsTemp.MoveNext
'    Loop
                
    gstrSql = "Select RPad('��λ��' || '" & gstrUnitName & "', 59, ' ') || ' ���ڣ�' As ��0," & vbNewLine & _
                "       RPad('��Ŀ��' || I.������ || '/' || I.Ӣ����, 30, ' ')|| RPad(' ������' || L.����, 29, ' ')  ||RPad(' ������' || D.����, 25, ' ')  as ��1 ," & vbNewLine & _
                "        rpad('�����ֵ��' || Replace(Replace(' 0' || trim(to_char(X.��ֵ,'999990." & strС�� & "')), ' 0.', '0.'), ' 0', '') || '(' || I.��λ || ')' || '   SD: ' ||" & vbNewLine & _
                "        Replace(Replace(' 0' || trim(to_char(X.Sd,'999990." & strС�� & "')), ' 0.', '0.'), ' 0', '') || '(' || I.��λ || ')' || '   CV: ' ||" & vbNewLine & _
                "        Replace(Replace(' 0' || trim(to_char(X.Cv * 100,'999990." & strС�� & "')), ' 0.', '0.'), ' 0', '') || '%',60,' ') || RPad('�ʿ�Ʒ��' || M.����, 20, ' ') ||RPad('���ţ�' || M.����, 20, ' ')   As ��2," & vbNewLine & _
                "        RPad('�Լ���' || M.�Լ�, 20, ' ') || RPad('У׼�' || M.У׼��, 20, ' ') as ��3, X.��ֵ, X.Sd, I.��λ" & vbNewLine & _
                "From �������� D, �����ʿ�Ʒ M, �����ʿؾ�ֵ X, ����������Ŀ I, �����ʿ�Ʒ��Ŀ L" & vbNewLine & _
                "Where D.ID = M.����id And M.ID = X.�ʿ�Ʒid And X.��Ŀid = I.ID And M.ID = [1] And X.��Ŀid = [2] And M.ID = L.�ʿ�Ʒid And" & vbNewLine & _
                "      L.��Ŀid = [2] And " & vbNewLine & _
                "      Instr(';' || [3] || ';',';' || X.�ʿ�Ʒid||'='||To_char(X.��ʼ����,'yyyy-MM-dd')||','||to_char(Nvl(X.��������, M.��������),'yyyy-mm-dd')||';' ) > 0 "
            
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResId, mlngItemID, mstr�ʿ�Ʒ����)
    If rsTemp.RecordCount <= 0 Then Me.chtThis(intIndex).Header.Text = "���ʿ�Ʒ��Ϣ��ȫ�棡": Exit Sub
    'If rsTemp.RecordCount <= 0 Then MsgBox "���ʿ�Ʒ��Ϣ��ȫ�棡", vbInformation, gstrSysName: Exit Sub
   
    '��֯��ͷ
    
'    strLable = rsTemp!��0 & Format(strStartDate, "yyyy��MM��dd��") & "��" & Format(strEndDate, "yyyy��MM��dd��")
'    strLable = strLable & vbCrLf & " " & rsTemp!��1 & vbCrLf & " " & rsTemp!��2 & vbCrLf & " " & rsTemp!��3
    
    
    
    dblAvg = Val("" & rsTemp!��ֵ): dblSD = Val("" & rsTemp!SD): strUnit = "" & rsTemp!��λ
    If dblAvg = 0 Or dblSD = 0 Then
'        MsgBox "��δ��ֵ��SDΪ0���޷�����" & Me.Caption & "��", vbInformation, gstrSysName: Exit Sub
        Me.chtThis(intIndex).Header.Text = "��δ��ֵ��SDΪ0���޷�����" & Me.Caption & "��": Exit Sub
    End If
    If Me.cbo��ʾ.ListIndex > 0 Then
        '���⡢XY������
        With Me.chtThis(intIndex).Header
            .Text = cboQCitem.Text
            .Adjust = oc2dAdjustCenter
            .Font.Bold = True
            .Font.Size = 8
        End With
    Else
        '���⡢XY������
        With Me.chtThis(intIndex).Header
            .Text = "�����Levey-Jennings" & IIf(mbln�����ʿ�ͼ, "����", "") & "��������ͼ" & vbCrLf & " " & vbCrLf & " "
            .Adjust = oc2dAdjustCenter
            .Font.Bold = True
            .Font.Size = 16
        End With
        With Me.chtThis(intIndex)
            strUnit = Nvl(rsTemp("��λ"))
            .ChartLabels.RemoveAll
            '��0
            .ChartLabels.Add
            .ChartLabels(1).AttachMethod = oc2dAttachCoord
            .ChartLabels(1).Anchor = oc2dAnchorNorth
            .ChartLabels(1).Text = rsTemp!��0 & Format(strStartDate, "yyyy��MM��dd��") & "��" & Format(strEndDate, "yyyy��MM��dd��")
            .ChartLabels(1).AttachCoord.x = (.ChartLabels(1).Location.Width / 2) + (.Width / Screen.TwipsPerPixelX / 2) - (.ChartLabels(1).Location.Width / 2) - 50
            .ChartLabels(1).AttachCoord.Y = .Header.Location.Top + .Header.Location.Height - 20
            '��1
            .ChartLabels.Add
            .ChartLabels(2).AttachMethod = oc2dAttachCoord
            .ChartLabels(2).Adjust = oc2dAdjustRight
            .ChartLabels(2).Text = rsTemp!��1
    '        .ChartLabels(2).AttachCoord.X = .Header.Location.Left + (.ChartLabels(2).Location.Width / 2) - 180
            .ChartLabels(2).AttachCoord.x = (.ChartLabels(2).Location.Width / 2) + (.Width / Screen.TwipsPerPixelX / 2) - (.ChartLabels(1).Location.Width / 2) - 50
            .ChartLabels(2).AttachCoord.Y = .ChartLabels(1).Location.Top + .ChartLabels(1).Location.Height + 10
            
            
            '��2
            .ChartLabels.Add
            .ChartLabels(3).AttachMethod = oc2dAttachCoord
            .ChartLabels(3).Adjust = oc2dAdjustRight
            .ChartLabels(3).Text = rsTemp!��2
    '        .ChartLabels(2).AttachCoord.X = .Header.Location.Left + (.ChartLabels(2).Location.Width / 2) - 180
            .ChartLabels(3).AttachCoord.x = (.ChartLabels(3).Location.Width / 2) + (.Width / Screen.TwipsPerPixelX / 2) - (.ChartLabels(1).Location.Width / 2) - 50
            .ChartLabels(3).AttachCoord.Y = .ChartLabels(2).Location.Top + .ChartLabels(1).Location.Height + 10
            
            strCalc = ""
            strLable = rsTemp!��3
            '��������ֵ��SD
            gstrSql = "Select Round(Avg(���), 2) As ��ֵ, Round(Stddev(���), 2) As Sd, Count(*) As ����" & vbNewLine & _
                "From (Select Trunc(Q.����ʱ��) As ����," & vbNewLine & _
                "              Avg(zl_Lis_toNumber(Q.�ʿ�ƷID,R.������ĿID,R.������,R.ID)) As ���" & vbNewLine & _
                "       From �����ʿؼ�¼ Q, ������ͨ��� R,�����ʿر��� T" & vbNewLine & _
                "       Where Q.�걾id = R.����걾id And Q.�ʿ�Ʒid = [1] And R.������Ŀid + 0 = [2] And" & vbNewLine & _
                "             Nvl(R.���ý��,0)=0 And R.ID=T.���ID(+) And Q.����ʱ�� Between   [3] and [4]  And Nvl(T.���, 0) <> 2" & vbNewLine & _
                "       Group By Trunc(Q.����ʱ��))"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResId, mlngItemID, CDate(strStartDate), CDate(strEndDate))
            
            If rsTemp.EOF = False Then
                If rsTemp("��ֵ") = 0 Then
                    strCalc = "�����ֵ��" & Format(rsTemp("��ֵ"), "0." & strС��) & "(" & strUnit & _
                                ")   SD: " & Format(rsTemp("SD"), "0." & strС��) & _
                                "(" & strUnit & ")   CV: " & Format(0, "0." & strС��) & "%"
                Else
                    strCalc = "�����ֵ��" & Format(rsTemp("��ֵ"), "0." & strС��) & "(" & strUnit & _
                                ")   SD: " & Format(rsTemp("SD"), "0." & strС��) & _
                                "(" & strUnit & ")   CV: " & Format(rsTemp("SD") / rsTemp("��ֵ") * 100, "0." & strС��) & "%"
                End If
            End If
            If LenB(StrConv(strCalc, vbFromUnicode)) < 60 Then
                strCalc = strCalc & Space(60 - LenB(StrConv(strCalc, vbFromUnicode))) & strLable
            Else
                strCalc = strCalc & strLable
            End If
            '��3
            .ChartLabels.Add
            .ChartLabels(4).AttachMethod = oc2dAttachCoord
            .ChartLabels(4).Adjust = oc2dAdjustRight
            .ChartLabels(4).Text = strCalc
    '        .ChartLabels(3).AttachCoord.X = .Header.Location.Left + (.ChartLabels(3).Location.Width / 2) - 180
            .ChartLabels(4).AttachCoord.x = (.ChartLabels(4).Location.Width / 2) + (.Width / Screen.TwipsPerPixelX / 2) - (.ChartLabels(1).Location.Width / 2) - 50
            .ChartLabels(4).AttachCoord.Y = .ChartLabels(3).Location.Top + .ChartLabels(2).Location.Height + 10
            
            
                
        End With
    End If
    
    With Me.chtThis(intIndex).ChartArea.Axes("Y")
        .MajorGrid.Spacing.IsDefault = False
        .AnnotationMethod = oc2dAnnotateValues
        .Title.Text = "�ⶨֵ" & IIf(strUnit = "", "", "(" & strUnit & ")")
    End With
    With Me.chtThis(intIndex).ChartArea.Axes("Y2")
        .AnnotationMethod = oc2dAnnotateValueLabels   '������2��ʾֵ��ʾ
        .Title.Text = "������"
        .Multiplier = 1
        With .ValueLabels
            .RemoveAll
'            .Add Val(dblAvg), "CL=    " & Format(Val(dblAvg), "0.00")
'            .Add Val(dblAvg) + 1 * Val(dblSD), "CL+1SD=" & Format(Val(dblAvg) + 1 * Val(dblSD), "0.00")
'            .Add Val(dblAvg) - 1 * Val(dblSD), "CL-1SD=" & Format(Val(dblAvg) - 1 * Val(dblSD), "0.00")
'            .Add Val(dblAvg) + 2 * Val(dblSD), "CL+2SD=" & Format(Val(dblAvg) + 2 * Val(dblSD), "0.00")
'            .Add Val(dblAvg) - 2 * Val(dblSD), "CL-2SD=" & Format(Val(dblAvg) - 2 * Val(dblSD), "0.00")
'            .Add Val(dblAvg) + 3 * Val(dblSD), "CL+3SD=" & Format(Val(dblAvg) + 3 * Val(dblSD), "0.00")
'            .Add Val(dblAvg) - 3 * Val(dblSD), "CL-3SD=" & Format(Val(dblAvg) - 3 * Val(dblSD), "0.00")
            .Add Val(dblAvg), Format(Val(dblAvg), "##0.00##") & " CL"
            .Add Val(dblAvg) + 1 * Val(dblSD), Format(Round(Val(dblAvg) + 1 * Val(dblSD), 4), "##0.00##") & " CL+1SD"
            .Add Val(dblAvg) - 1 * Val(dblSD), Format(Round(Val(dblAvg) - 1 * Val(dblSD), 4), "##0.00##") & " CL-1SD"
            .Add Val(dblAvg) + 2 * Val(dblSD), Format(Round(Val(dblAvg) + 2 * Val(dblSD), 4), "##0.00##") & " CL+2SD"
            .Add Val(dblAvg) - 2 * Val(dblSD), Format(Round(Val(dblAvg) - 2 * Val(dblSD), 4), "##0.00##") & " CL-2SD"
            .Add Val(dblAvg) + 3 * Val(dblSD), Format(Round(Val(dblAvg) + 3 * Val(dblSD), 4), "##0.00##") & " CL+3SD"
            .Add Val(dblAvg) - 3 * Val(dblSD), Format(Round(Val(dblAvg) - 3 * Val(dblSD), 4), "##0.00##") & " CL-3SD"
        End With
    End With
    With Me.chtThis(intIndex).ChartArea.Axes("X")
        .MajorGrid.Spacing.IsDefault = False
        .AnnotationMethod = oc2dAnnotateValueLabels   '��������ʾֵ��ʾ
        .Title.Text = "����"
    End With
    
    '������֯
'    gstrSql = "Select ����ʱ��, ���Դ���, Max(Decode(���, 2, 0, ���)) As �ڿ�, Max(Decode(���, 2, ���, 0)) As ʧ��" & vbNewLine & _
'            "From (Select Q.����ʱ��, Q.���Դ���, T.���," & vbNewLine & _
'            "              Decode(I.ֵ����, Null, Zl_To_Number(R.������)," & vbNewLine & _
'            "                      Length(Substr(I.ֵ����, 1, Instr(I.ֵ����, ';' || RTrim(R.������) || ';'))) -" & vbNewLine & _
'            "                       Nvl(Length(Replace(Substr(I.ֵ����, 1, Instr(I.ֵ����, ';' || RTrim(R.������) || ';')), ';')), 0)) As ���" & vbNewLine & _
'            "       From �����ʿؼ�¼ Q, ������ͨ��� R, �����ʿر��� T," & vbNewLine & _
'            "            (Select Decode(�������, 3, Decode(RTrim(ȡֵ����), '', '', ';' || RTrim(ȡֵ����) || ';'), '') As ֵ����" & vbNewLine & _
'            "              From ������Ŀ" & vbNewLine & _
'            "              Where ������Ŀid = [2]) I" & vbNewLine & _
'            "       Where Q.�걾id = R.����걾id And R.ID = T.���id(+) And /*Nvl(R.�Ƿ����, 0) = 1 And*/ Q.�ʿ�Ʒid + 0 = [1] And" & vbNewLine & _
'            "             R.������Ŀid + 0 = [2] And" & vbNewLine & _
'            "             (Q.����ʱ�� Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd')))" & vbNewLine & _
'            "Group By ����ʱ��, ���Դ���" & vbNewLine & _
'            "Order By ����ʱ��, ���Դ���"
                
'    gstrSql = "Select ����ʱ��,  Max(Decode(���, 2, 0, ���)) As �ڿ�, Max(Decode(���, 2, ���, 0)) As ʧ��" & vbNewLine & _

    
    Set rsTemp = GetQCChartData(lngResId, mlngItemID, strStartDate, strEndDate)
    
    
    Me.chtThis(intIndex).ChartArea.Axes("X").ValueLabels.RemoveAll
    str�������� = ""
    With rsTemp
        If .RecordCount < 30 And mint��λ��ʾ = 1 Then
            intLoop = .RecordCount
            ReDim Preserve aryX(31)
            ReDim Preserve aryY(31, 14)
        Else
            intLoop = 0
            ReDim aryX(.RecordCount)
            ReDim aryY(.RecordCount, 14)
        End If

        For lngTmp = LBound(aryY) To UBound(aryY)
            aryX(lngTmp) = lngTmp
            aryY(lngTmp, 0) = Val(dblAvg)
            aryY(lngTmp, 1) = Val(dblAvg) + 1 * Val(dblSD)
            aryY(lngTmp, 2) = Val(dblAvg) - 1 * Val(dblSD)
            aryY(lngTmp, 3) = Val(dblAvg) + 2 * Val(dblSD)
            aryY(lngTmp, 4) = Val(dblAvg) - 2 * Val(dblSD)
            aryY(lngTmp, 5) = Val(dblAvg) + 3 * Val(dblSD)
            aryY(lngTmp, 6) = Val(dblAvg) - 3 * Val(dblSD)
            aryY(lngTmp, 7) = Val(dblAvg) + 4 * Val(dblSD)
            aryY(lngTmp, 8) = Val(dblAvg) - 4 * Val(dblSD)
            aryY(lngTmp, 9) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(lngTmp, 10) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(lngTmp, 11) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(lngTmp, 12) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(lngTmp, 13) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(lngTmp, 14) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
        Next
        

        dblMax = 4 * Val(dblSD)
        
        Do While Not .EOF
'            Me.ChtThis.ChartArea.Axes("X").ValueLabels.Add .AbsolutePosition, .AbsolutePosition
            bln�ϲ��� = False
            If lngX > 0 Then
                If Not (aryY(lngX, 9) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue And dateLoop = Format(Nvl(!����ʱ��), "yyyy-MM-dd")) Then
                    lngX = lngX + 1
                    If Format(Nvl(!����ʱ��), "dd") <> "01" Then
                        Me.chtThis(intIndex).ChartArea.Axes("X").ValueLabels.Add lngX, Format(Nvl(!����ʱ��), "dd")
                    Else
                        Me.chtThis(intIndex).ChartArea.Axes("X").ValueLabels.Add lngX, Format(Nvl(!����ʱ��), "mm" & "��")
                    End If
                Else
                    bln�ϲ��� = True
                    intLoop = intLoop - 1
                End If
            Else
                lngX = lngX + 1
                If Format(Nvl(!����ʱ��), "dd") <> "01" Then
                    Me.chtThis(intIndex).ChartArea.Axes("X").ValueLabels.Add lngX, Format(Nvl(!����ʱ��), "dd")
                Else
                    Me.chtThis(intIndex).ChartArea.Axes("X").ValueLabels.Add lngX, Format(Nvl(!����ʱ��), "mm" & "��")
                End If
            End If

            dateLoop = Format(Nvl(!����ʱ��), "yyyy-MM-dd")
            strAllCount = Trim$("" & !���Դ���)
            aryX(lngX) = lngX
            aryY(lngX, 0) = Val(dblAvg)
            aryY(lngX, 1) = Val(dblAvg) + 1 * Val(dblSD)
            aryY(lngX, 2) = Val(dblAvg) - 1 * Val(dblSD)
            aryY(lngX, 3) = Val(dblAvg) + 2 * Val(dblSD)
            aryY(lngX, 4) = Val(dblAvg) - 2 * Val(dblSD)
            aryY(lngX, 5) = Val(dblAvg) + 3 * Val(dblSD)
            aryY(lngX, 6) = Val(dblAvg) - 3 * Val(dblSD)
            aryY(lngX, 7) = Val(dblAvg) + 4 * Val(dblSD)
            aryY(lngX, 8) = Val(dblAvg) - 4 * Val(dblSD)
            
            
            If "" & !�ڿ� <> "" Then
'                aryY(lngX, 9) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue 'Val(dblAvg)
'            Else
                If Abs(Val("" & !�ڿ�) - Val(dblAvg)) > dblMax Then
                    aryY(lngX, 9) = IIf((Val("" & !�ڿ�) - Val(dblAvg)) < dblMax, Val(dblAvg) - dblMax + 0.03 * dblSD, Val(dblAvg) + dblMax - 0.03 * dblSD)
                ElseIf InStr(",0.00,1.00,2.00,3.00,4.00,", "," & Format((Abs(Val("" & !�ڿ�)) - dblAvg) / dblSD, "0.00") & ",") > 0 Then
                    aryY(lngX, 9) = Val("" & !�ڿ�) + 0.03 * dblSD
                Else
                    aryY(lngX, 9) = Val("" & !�ڿ�)
                End If
                strTmp = Val("" & !�ڿ�)
                If Left(strTmp, 1) = "." Then strTmp = "0" & strTmp
                
                
                If InStr(strAllCount, ",") > 0 Then
                    strCurCount = Mid$(strAllCount, 1, InStr(strAllCount, ",") - 1)
                    strAllCount = Mid$(strAllCount, InStr(strAllCount, ",") + 1)
                Else
                    strCurCount = strAllCount
                End If
                strTmp = "������:" & strTmp & " ����:" & Format(Nvl(!����ʱ��), "yyyy-MM-dd") & " " & Trim("" & !ʱ��) & " ��" & strCurCount & "��"
                str�������� = str�������� & "|" & lngX & ",9," & strTmp
                
'                If dblMax < Abs(Val("" & !�ڿ�) - Val(dblAvg)) Then dblMax = Abs(Val("" & !�ڿ�) - Val(dblAvg))
            End If
            
            If Not bln�ϲ��� Then
                If "" & !ʧ��1 <> "" Then
'                    aryY(lngX, 10) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
'                Else
                    If Abs(Val("" & !ʧ��1) - Val(dblAvg)) > dblMax Then
                        aryY(lngX, 10) = IIf((Val("" & !ʧ��1) - Val(dblAvg)) < dblMax, Val(dblAvg) - dblMax + 0.03 * dblSD, Val(dblAvg) + dblMax - 0.03 * dblSD)
                    ElseIf InStr(",0.00,1.00,2.00,3.00,4.00,", "," & Format((Abs(Val("" & !ʧ��1)) - dblAvg) / dblSD, "0.00") & ",") > 0 Then
                        aryY(lngX, 10) = Val("" & !ʧ��1) + 0.03 * dblSD
                    Else
                        aryY(lngX, 10) = Val("" & !ʧ��1)
                    End If
                    strTmp = Val("" & !ʧ��1)
                    If Left(strTmp, 1) = "." Then strTmp = "0" & strTmp
                    If InStr(strAllCount, ",") > 0 Then
                        strCurCount = Mid$(strAllCount, 1, InStr(strAllCount, ","))
                        strAllCount = Mid$(strAllCount, InStr(strAllCount, ",") + 1)
                    Else
                        strCurCount = strAllCount
                    End If
                    strTmp = "������:" & strTmp & " ����:" & Format(Nvl(!����ʱ��), "yyyy-MM-dd") & " " & Trim("" & !ʱ��) & " ��" & strCurCount & "��"
                    str�������� = str�������� & "|" & lngX & ",10," & strTmp
    '                If dblMax < Abs(Val("" & !ʧ��) - Val(dblAvg)) Then dblMax = Abs(Val("" & !ʧ��) - Val(dblAvg))
                End If
                
                If "" & !ʧ��2 <> "" Then
'                    aryY(lngX, 11) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
'                Else
                    If Abs(Val("" & !ʧ��2) - Val(dblAvg)) > dblMax Then
                        aryY(lngX, 11) = IIf((Val("" & !ʧ��2) - Val(dblAvg)) < dblMax, Val(dblAvg) - dblMax + 0.03 * dblSD, Val(dblAvg) + dblMax - 0.03 * dblSD)
                    ElseIf InStr(",0.00,1.00,2.00,3.00,4.00,", "," & Format((Abs(Val("" & !ʧ��2)) - dblAvg) / dblSD, "0.00") & ",") > 0 Then
                        aryY(lngX, 11) = Val("" & !ʧ��2) + 0.03 * dblSD
                    Else
                        aryY(lngX, 11) = Val("" & !ʧ��2)
                    End If
                    strTmp = Val("" & !ʧ��2)
                    If Left(strTmp, 1) = "." Then strTmp = "0" & strTmp
                    If InStr(strAllCount, ",") > 0 Then
                        strCurCount = Mid$(strAllCount, 1, InStr(strAllCount, ",") - 1)
                        strAllCount = Mid$(strAllCount, InStr(strAllCount, ",") + 1)
                    Else
                        strCurCount = strAllCount
                    End If
                    strTmp = "������:" & strTmp & " ����:" & Format(Nvl(!����ʱ��), "yyyy-MM-dd") & " " & Trim("" & !ʱ��) & " ��" & strCurCount & "��"
                    str�������� = str�������� & "|" & lngX & ",11," & strTmp
                End If
                
                If "" & !ʧ��3 <> "" Then
'                    aryY(lngX, 12) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
'                Else
                    If Abs(Val("" & !ʧ��3) - Val(dblAvg)) > dblMax Then
                        aryY(lngX, 12) = IIf((Val("" & !ʧ��3) - Val(dblAvg)) < dblMax, Val(dblAvg) - dblMax + 0.03 * dblSD, Val(dblAvg) + dblMax - 0.03 * dblSD)
                    ElseIf InStr(",0.00,1.00,2.00,3.00,4.00,", "," & Format((Abs(Val("" & !ʧ��3)) - dblAvg) / dblSD, "0.00") & ",") > 0 Then
                        aryY(lngX, 12) = Val("" & !ʧ��3) + 0.03 * dblSD
                    Else
                        aryY(lngX, 12) = Val("" & !ʧ��3)
                    End If
                    strTmp = Val("" & !ʧ��3)
                    If Left(strTmp, 1) = "." Then strTmp = "0" & strTmp
                    If InStr(strAllCount, ",") > 0 Then
                        strCurCount = Mid$(strAllCount, 1, InStr(strAllCount, ",") - 1)
                        strAllCount = Mid$(strAllCount, InStr(strAllCount, ",") + 1)
                    Else
                        strCurCount = strAllCount
                    End If
                    strTmp = "������:" & strTmp & " ����:" & Format(Nvl(!����ʱ��), "yyyy-MM-dd") & " " & Trim("" & !ʱ��) & " ��" & strCurCount & "��"
                    str�������� = str�������� & "|" & lngX & ",12," & strTmp
    '                If dblMax < Abs(Val("" & !ʧ��) - Val(dblAvg)) Then dblMax = Abs(Val("" & !ʧ��) - Val(dblAvg))
                End If
                
                If "" & !ʧ��4 <> "" Then
'                    aryY(lngX, 13) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
'                Else
                    If Abs(Val("" & !ʧ��4) - Val(dblAvg)) > dblMax Then
                        aryY(lngX, 13) = IIf((Val("" & !ʧ��4) - Val(dblAvg)) < dblMax, Val(dblAvg) - dblMax + 0.03 * dblSD, Val(dblAvg) + dblMax - 0.03 * dblSD)
                    ElseIf InStr(",0.00,1.00,2.00,3.00,4.00,", "," & Format((Abs(Val("" & !ʧ��4)) - dblAvg) / dblSD, "0.00") & ",") > 0 Then
                        aryY(lngX, 13) = Val("" & !ʧ��4) + 0.03 * dblSD
                    Else
                        aryY(lngX, 13) = Val("" & !ʧ��4)
                    End If
                    strTmp = Val("" & !ʧ��4)
                    If Left(strTmp, 1) = "." Then strTmp = "0" & strTmp
                    If InStr(strAllCount, ",") > 0 Then
                        strCurCount = Mid$(strAllCount, 1, InStr(strAllCount, ",") - 1)
                        strAllCount = Mid$(strAllCount, InStr(strAllCount, ",") + 1)
                    Else
                        strCurCount = strAllCount
                    End If
                    strTmp = "������:" & strTmp & " ����:" & Format(Nvl(!����ʱ��), "yyyy-MM-dd") & " " & Trim("" & !ʱ��) & " ��" & strCurCount & "��"
                    str�������� = str�������� & "|" & lngX & ",13," & strTmp
    '                If dblMax < Abs(Val("" & !ʧ��) - Val(dblAvg)) Then dblMax = Abs(Val("" & !ʧ��) - Val(dblAvg))
                End If
                
                If "" & !ʧ��5 <> "" Then
'                    aryY(lngX, 14) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
'                Else
                    If Abs(Val("" & !ʧ��5) - Val(dblAvg)) > dblMax Then
                        aryY(lngX, 14) = IIf((Val("" & !ʧ��5) - Val(dblAvg)) < dblMax, Val(dblAvg) - dblMax + 0.03 * dblSD, Val(dblAvg) + dblMax - 0.03 * dblSD)
                    ElseIf InStr(",0.00,1.00,2.00,3.00,4.00,", "," & Format((Abs(Val("" & !ʧ��4)) - dblAvg) / dblSD, "0.00") & ",") > 0 Then
                        aryY(lngX, 14) = Val("" & !ʧ��5) + 0.03 * dblSD
                    Else
                        aryY(lngX, 14) = Val("" & !ʧ��5)
                    End If
                    strTmp = Val("" & !ʧ��5)
                    If Left(strTmp, 1) = "." Then strTmp = "0" & strTmp
                    If InStr(strAllCount, ",") > 0 Then
                        strCurCount = Mid$(strAllCount, 1, InStr(strAllCount, ",") - 1)
                        strAllCount = Mid$(strAllCount, InStr(strAllCount, ",") + 1)
                    Else
                        strCurCount = strAllCount
                    End If
                    strTmp = "������:" & strTmp & " ����:" & Format(Nvl(!����ʱ��), "yyyy-MM-dd") & " " & Trim("" & !ʱ��) & " ��" & strCurCount & "��"
                    str�������� = str�������� & "|" & lngX & ",14," & strTmp
                    
    '                If dblMax < Abs(Val("" & !ʧ��) - Val(dblAvg)) Then dblMax = Abs(Val("" & !ʧ��) - Val(dblAvg))
                End If
            End If
            .MoveNext
        Loop
        
        Do While lngX < UBound(aryX)
            '�м�����кϲ��ĵ㣬���ﲹ������
            lngX = lngX + 1
            aryX(lngX) = lngX
            aryY(lngX, 0) = Val(dblAvg)
            aryY(lngX, 1) = Val(dblAvg) + 1 * Val(dblSD)
            aryY(lngX, 2) = Val(dblAvg) - 1 * Val(dblSD)
            aryY(lngX, 3) = Val(dblAvg) + 2 * Val(dblSD)
            aryY(lngX, 4) = Val(dblAvg) - 2 * Val(dblSD)
            aryY(lngX, 5) = Val(dblAvg) + 3 * Val(dblSD)
            aryY(lngX, 6) = Val(dblAvg) - 3 * Val(dblSD)
            aryY(lngX, 7) = Val(dblAvg) + 4 * Val(dblSD)
            aryY(lngX, 8) = Val(dblAvg) - 4 * Val(dblSD)
            
            aryY(lngX, 9) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(lngX, 10) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(lngX, 11) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(lngX, 12) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(lngX, 13) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(lngX, 14) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
        Loop
    End With
    
    '�������30�������,����30�������
    If intLoop > 0 And mint��λ��ʾ = 1 Then
        For intLoop = intLoop + 1 To 31
            
            dateLoop = DateAdd("d", 1, dateLoop)
            If dateLoop <= CDate(strEndDate) Then
                If Format(Nvl(dateLoop), "dd") <> "01" Then
                    Me.chtThis(intIndex).ChartArea.Axes("X").ValueLabels.Add intLoop, Format(Nvl(dateLoop), "dd")
                Else
                    Me.chtThis(intIndex).ChartArea.Axes("X").ValueLabels.Add intLoop, Format(Nvl(dateLoop), "mm" & "��")
                End If
            End If
            aryX(intLoop) = intLoop
            aryY(intLoop, 0) = Val(dblAvg)
            aryY(intLoop, 1) = Val(dblAvg) + 1 * Val(dblSD)
            aryY(intLoop, 2) = Val(dblAvg) - 1 * Val(dblSD)
            aryY(intLoop, 3) = Val(dblAvg) + 2 * Val(dblSD)
            aryY(intLoop, 4) = Val(dblAvg) - 2 * Val(dblSD)
            aryY(intLoop, 5) = Val(dblAvg) + 3 * Val(dblSD)
            aryY(intLoop, 6) = Val(dblAvg) - 3 * Val(dblSD)
            aryY(intLoop, 7) = Val(dblAvg) + 4 * Val(dblSD)
            aryY(intLoop, 8) = Val(dblAvg) - 4 * Val(dblSD)
            
            aryY(intLoop, 9) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(intLoop, 10) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(intLoop, 11) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(intLoop, 12) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(intLoop, 13) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
            aryY(intLoop, 14) = Me.chtThis(intIndex).ChartGroups(1).Data.HoleValue
        Next
    End If

    If str�������� <> "" Then
        str�������� = Mid(str��������, 2)
        ReDim mArr(intIndex + 1)
        mArr(intIndex) = str��������
    End If
    '���ˢ���ڲ�����
    With Me.chtThis(intIndex)
        .IsBatched = True
        With .ChartGroups(1).Data
            .NumPoints(1) = UBound(aryX) + 1
            Call .CopyXVectorIn(1, aryX)
            Call .CopyYArrayIn(aryY)
        End With
        With .ChartArea.Axes("Y")
            .Min = Val(dblAvg) - Val(dblMax)
            .Origin = .Min
            .Max = Val(dblAvg) + Val(dblMax)
            .MajorGrid.Spacing.IsDefault = False
            .AnnotationMethod = oc2dAnnotateValues
        End With
        With .ChartArea.Axes("X")
            .Min = 0: .Max = aryX(UBound(aryX))
        End With
        .IsBatched = False
    End With
    Call chtThis_Resize(intIndex, chtThis(intIndex).Width, chtThis(intIndex).Height)
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetQCChartData(ByVal lngResId As Long, ByVal lngItemID As Long, ByVal strDateS As String, ByVal strDateE As String) As adodb.Recordset
    'ȡ��ͼ�õ����ݼ�
    Dim rsTemp As adodb.Recordset, rsQcData As adodb.Recordset
    Dim strLastDate As String, i As Integer
    On Error GoTo errH
    Set rsQcData = New adodb.Recordset
    rsQcData.Fields.Append "����ʱ��", adVarChar, 50
    rsQcData.Fields.Append "ʱ��", adVarChar, 8
    rsQcData.Fields.Append "���Դ���", adVarChar, 10
    rsQcData.Fields.Append "�ڿ�", adVarChar, 30
    rsQcData.Fields.Append "ʧ��1", adVarChar, 30
    rsQcData.Fields.Append "ʧ��2", adVarChar, 30
    rsQcData.Fields.Append "ʧ��3", adVarChar, 30
    rsQcData.Fields.Append "ʧ��4", adVarChar, 30
    rsQcData.Fields.Append "ʧ��5", adVarChar, 30


    rsQcData.CursorLocation = adUseClient
    rsQcData.LockType = adLockOptimistic
    rsQcData.CursorType = adOpenStatic
    rsQcData.Open

                
    gstrSql = "Select Q.����ʱ��,Q.ʱ��, Q.���Դ���, Nvl(T.���,0) as ���," & vbNewLine & _
                "                     Zl_Lis_Tonumber(Q.�ʿ�ƷID, R.������Ŀid, R.������,R.ID) As ���" & vbNewLine & _
                "              From �����ʿؼ�¼ Q, ������ͨ��� R, �����ʿر��� T" & vbNewLine & _
                "              Where Q.�걾id = R.����걾id And R.ID = T.���id(+) And Q.�ʿ�Ʒid = [1] And" & vbNewLine & _
                "                    Nvl(R.���ý��,0)=0 And Q.����ʱ�� Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd') And " & vbNewLine & _
                "                    R.������Ŀid + 0 = [2] order by  Q.����ʱ��, Nvl(T.���,0), Q.���Դ��� "
                
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResId, lngItemID, strDateS, strDateE)
    
    Do Until rsTemp.EOF
        
        If Val("" & rsTemp.Fields("���").Value) = 2 Then
            'ʧ�ص�
            If strLastDate <> Format(rsTemp.Fields("����ʱ��").Value, "yyyy-MM-dd") Then
                rsQcData.AddNew
                rsQcData("����ʱ��") = Format(rsTemp.Fields("����ʱ��").Value, "yyyy-MM-dd")
                rsQcData("ʱ��") = Trim("" & rsTemp.Fields("ʱ��").Value)
                rsQcData("���Դ���") = rsTemp.Fields("���Դ���").Value
                rsQcData("�ڿ�") = ""
                rsQcData("ʧ��1") = Trim("" & rsTemp.Fields("���").Value)
                rsQcData("ʧ��2") = ""
                rsQcData("ʧ��3") = ""
                rsQcData("ʧ��4") = ""
                rsQcData("ʧ��5") = ""
            Else
               For i = 1 To 5
                    If Trim("" & rsQcData.Fields(3 + i).Value) = "" Then
                         rsQcData.Fields(2).Value = rsQcData.Fields(2).Value & "," & Trim("" & rsTemp.Fields("���Դ���").Value)
                         rsQcData.Fields(3 + i).Value = Trim("" & rsTemp.Fields("���").Value)
                         Exit For
                    End If
               Next
            End If
        Else
            '�ڿ��뾯��
            rsQcData.AddNew
            rsQcData("����ʱ��") = Format(rsTemp.Fields("����ʱ��").Value, "yyyy-MM-dd")
            rsQcData("ʱ��") = Trim("" & rsTemp.Fields("ʱ��").Value)
            rsQcData("���Դ���") = rsTemp.Fields("���Դ���").Value
            rsQcData("�ڿ�") = Trim("" & rsTemp.Fields("���").Value)
            rsQcData("ʧ��1") = ""
            rsQcData("ʧ��2") = ""
            rsQcData("ʧ��3") = ""
            rsQcData("ʧ��4") = ""
            rsQcData("ʧ��5") = ""
        End If
        strLastDate = Format(rsTemp.Fields("����ʱ��").Value, "yyyy-MM-dd")
        
        rsTemp.MoveNext
    Loop
    If rsQcData.RecordCount > 0 Then rsQcData.MoveFirst
    
    Set GetQCChartData = rsQcData
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cbo��ʾ_Click()
    Dim intLoop As Integer
    ReDim mArr(Me.cboQCitem.ListCount)
    If Me.cbo��ʾ.Tag = "��ˢ��" Then Exit Sub
    For intLoop = 0 To Me.cbo��ʾ.ListCount - 1
        Call RefChart(intLoop)
    Next
    Call Form_Resize
End Sub

Private Sub ChtThis_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim px As Long
    Dim py As Long
    Dim Series As Long
    Dim Point As Long
    Dim Distance As Long
    Dim Region As Long
    Dim i As Integer, strTmp As String
    Dim varTmp As Variant
    
    On Error Resume Next
    
    px = x / Screen.TwipsPerPixelX
    py = Y / Screen.TwipsPerPixelY
    If mLastXY = px & "," & py Then Exit Sub
    mLastXY = px & "," & py
    
    If (Button = 0) Then
        With chtThis(Index)
            Region = .ChartGroups(1).CoordToDataIndex(px, py, oc2dFocusXY, Series, Point, Distance)
            If (Series > 9 And Point > 0) And (Distance <= 5) Then
                If (Region = oc2dRegionInChartArea) Then
                    .ToolTipText = .ChartGroups(1).Data(Series, Point)
                    
                        If mArr(Index) <> "" Then
                            varTmp = Split(mArr(Index), "|")
                            For i = LBound(varTmp) To UBound(varTmp)
                                strTmp = varTmp(i)
                                If strTmp <> "" Then
                                    If Split(strTmp, ",")(0) = Point - 1 And Split(strTmp, ",")(1) = Series - 1 Then
                                        .ToolTipText = Split(strTmp, ",")(2)
                                    End If
                                End If
                            Next
                        End If
                    
                    If Left(.ToolTipText, 1) = "." Then .ToolTipText = "0" & .ToolTipText
                End If
            Else
                .ToolTipText = ""
                .Footer.Text = ""
            End If
            .Refresh
        End With
    End If
End Sub

Private Sub chtThis_Resize(Index As Integer, ByVal Width As Long, ByVal Height As Long)
    On Error Resume Next
    With Me.chtThis(Index)
        '��1
        .ChartLabels(1).AttachCoord.x = .Header.Location.Left + (.ChartLabels(1).Location.Width / 2) - 80
        .ChartLabels(1).AttachCoord.Y = .Header.Location.Top + .Header.Location.Height - 30
        '��2
        .ChartLabels(2).AttachCoord.x = .Header.Location.Left + (.ChartLabels(2).Location.Width / 2) - 80
        .ChartLabels(2).AttachCoord.Y = .ChartLabels(1).Location.Top + .ChartLabels(1).Location.Height + 10
        '��3
        .ChartLabels(3).AttachCoord.x = .Header.Location.Left + (.ChartLabels(3).Location.Width / 2) - 80
        .ChartLabels(3).AttachCoord.Y = .ChartLabels(2).Location.Top + .ChartLabels(2).Location.Height + 10
        '��3
        .ChartLabels(4).AttachCoord.x = .Header.Location.Left + (.ChartLabels(4).Location.Width / 2) - 80
        .ChartLabels(4).AttachCoord.Y = .ChartLabels(3).Location.Top + .ChartLabels(3).Location.Height + 10
    End With
End Sub

Private Sub Form_Load()
        
'    ReDim mArr(ChtThis.Count)
End Sub

'--------------------------------------------
'����Ϊ�ؼ��¼�����
'--------------------------------------------
Private Sub opt�ʿ�Ʒ_Click(Index As Integer)
    Dim intLoop As Integer
    If Me.Visible = False Then Exit Sub
    If Me.opt�ʿ�Ʒ(Index).Enabled = False Then Exit Sub
    If Me.cbo��ʾ.Tag = "��ˢ��" Then Exit Sub
    
    Call Form_Resize
    For intLoop = 0 To Me.chtThis.Count - 1
'        If intLoop = Index Then
'            Me.ChtThis(intLoop).Visible = True
'        Else
'            Me.ChtThis(intLoop).Visible = False
'        End If
        Call RefChart(intLoop)
    Next
    
End Sub

Private Sub cboQCitem_Click()
    Dim intLoop As Integer
    If Me.Visible = False Then Exit Sub
'    If Me.opt�ʿ�Ʒ(Index).Enabled = False Then Exit Sub
    If Me.cbo��ʾ.Tag = "��ˢ��" Then Exit Sub
    
    Call Form_Resize
    DoEvents

'        If intLoop = Index Then
'            Me.ChtThis(intLoop).Visible = True
'        Else
'            Me.ChtThis(intLoop).Visible = False
'        End If
        Call RefChart(cboQCitem.ListIndex)
'         Me.chtThis(cboQCitem.ListIndex).Visible = False
End Sub

Private Sub Form_Resize()
    Dim intLoop As Integer
    Dim intIndex As Integer
    Err = 0: On Error Resume Next
    
    For intLoop = 0 To Me.cboQCitem.ListCount - 1
        If Me.cboQCitem.ListIndex = intLoop Then
            intIndex = intLoop
        End If
        Me.chtThis(intLoop).Visible = False
    Next
    Select Case Me.cbo��ʾ.ListIndex + 1
        Case 1
            With Me.chtThis(intIndex)
                .Visible = True
                .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
                .Top = Me.ScaleTop: .Height = Me.ScaleHeight - Me.opt�ʿ�Ʒ(0).Height - Screen.TwipsPerPixelY * 4
            End With
        Case 2
            With Me.chtThis(intIndex)
                .Visible = True
                .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
                .Top = Me.ScaleTop
                .Height = (Me.ScaleHeight - Me.opt�ʿ�Ʒ(0).Height - Screen.TwipsPerPixelY * 4) / 2
            End With
            With Me.chtThis(intIndex + 1)
                .Visible = True
                .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
                .Top = Me.chtThis(intIndex).Top + Me.chtThis(intIndex).Height
                .Height = Me.chtThis(intIndex).Height
            End With
        Case 3
            With Me.chtThis(intIndex)
                .Visible = True
                .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
                .Top = Me.ScaleTop
                .Height = (Me.ScaleHeight - Me.opt�ʿ�Ʒ(0).Height - Screen.TwipsPerPixelY * 4) / 3
            End With
            With Me.chtThis(intIndex + 1)
                .Visible = True
                .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
                .Top = Me.chtThis(intIndex).Top + Me.chtThis(intIndex).Height
                .Height = Me.chtThis(intIndex).Height
            End With
            With Me.chtThis(intIndex + 2)
                .Visible = True
                .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
                .Top = Me.chtThis(intIndex + 1).Top + Me.chtThis(intIndex + 1).Height
                .Height = Me.chtThis(intIndex).Height
            End With
        Case 4
        Case 5
        Case 6
    End Select
'    With Me.ChtThis(0)
'        .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
'        .Top = Me.ScaleTop: .Height = Me.ScaleHeight - Me.opt�ʿ�Ʒ(0).Height - Screen.TwipsPerPixelY * 4
'    End With
    
    With Me.opt�ʿ�Ʒ(0)
        .Left = Me.ScaleLeft + Screen.TwipsPerPixelX * 2
        .Top = Me.ScaleHeight - .Height - Screen.TwipsPerPixelY * 2
    End With
    
    
    With Me.cboQCitem
        .Left = Me.ScaleLeft + Screen.TwipsPerPixelX * 2
        .Top = Me.ScaleHeight - .Height - Screen.TwipsPerPixelY * 2
    End With
    
    
    For lngCount = 1 To Me.opt�ʿ�Ʒ.Count
        With Me.opt�ʿ�Ʒ(lngCount)
            .Left = Me.opt�ʿ�Ʒ(lngCount - 1).Left + Me.opt�ʿ�Ʒ(lngCount - 1).Width + Screen.TwipsPerPixelX * 10
            .Top = Me.opt�ʿ�Ʒ(lngCount - 1).Top
        End With
    Next
    
    With Me.cbo��ʾ
        .Top = Me.opt�ʿ�Ʒ(0).Top
        .Left = Me.ScaleWidth - .Width - Screen.TwipsPerPixelX * 2
    End With
End Sub


Public Function ZLGetLJ_QCID() As Long
    '����       �õ���ǰʹ�õ��ʿ�Ʒ��ID
    For lngCount = 0 To Me.cboQCitem.ListCount - 1
        If Me.cboQCitem.ListIndex = lngCount Then ZLGetLJ_QCID = Val(Me.cboQCitem.ItemData(lngCount)): Exit For
    Next
End Function

Public Function ZLGetLJ_QCIDStr() As String
    '����       �õ���ǰʹ�õ��ʿ�ƷID��
    For lngCount = 0 To Me.cboQCitem.ListCount - 1
'        If Me.opt�ʿ�Ʒ(lngCount).Enabled = True Then
            ZLGetLJ_QCIDStr = ZLGetLJ_QCIDStr & "," & Val(Me.cboQCitem.ItemData(lngCount))
'        End If
    Next
    ZLGetLJ_QCIDStr = Mid(ZLGetLJ_QCIDStr, 2)
End Function



