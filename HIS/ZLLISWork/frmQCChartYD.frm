VERSION 5.00
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmQCChartYD 
   BorderStyle     =   0  'None
   Caption         =   "Youdenͼ"
   ClientHeight    =   5355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.ComboBox cbo�ʿ�Ʒ 
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      Left            =   4530
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4935
      Width           =   3180
   End
   Begin VB.ComboBox cbo�ʿ�Ʒ 
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   570
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   4920
      Width           =   3180
   End
   Begin C1Chart2D8.Chart2D chtThis 
      Height          =   2955
      Left            =   510
      TabIndex        =   1
      Top             =   645
      Width           =   6570
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   11589
      _ExtentY        =   5212
      _StockProps     =   0
      ControlProperties=   "frmQCChartYD.frx":0000
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbl�ʿ�Ʒ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Index           =   1
      Left            =   4050
      TabIndex        =   4
      Top             =   5010
      Width           =   360
   End
   Begin VB.Label lbl�ʿ�Ʒ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   2
      Top             =   4995
      Width           =   360
   End
End
Attribute VB_Name = "frmQCChartYD"
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
    With Me.comDlg
        .CancelError = True
        .DialogTitle = "���Ϊ"
        .filter = "(ͼ���ļ�)|*.jpg"
        .FileName = Me.Caption & Format(mstrToDate, "yyyyMMdd") & ".jpg"
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
    
    mstrResList = strResList
    mlngItemID = lngItemID
    mstrFromDate = strFromDate
    mstrToDate = strToDate
    mstr�ʿ�Ʒ���� = str�ʿ�Ʒ����
    
    Me.Tag = "��ˢ��"
    
    Me.cbo�ʿ�Ʒ(0).Enabled = False: Me.cbo�ʿ�Ʒ(1).Enabled = False
    Me.cbo�ʿ�Ʒ(0).Clear: Me.cbo�ʿ�Ʒ(1).Clear
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select ID, ���� || '-' || ���� As �ʿ�Ʒ From �����ʿ�Ʒ Where Instr(',' || [1] || ',', ',' || ID || ',') > 0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strResList)
    With rsTemp
        Do While Not .EOF
            For lngCount = 0 To 1
                Me.cbo�ʿ�Ʒ(lngCount).AddItem "" & !�ʿ�Ʒ
                Me.cbo�ʿ�Ʒ(lngCount).ItemData(Me.cbo�ʿ�Ʒ(lngCount).NewIndex) = !ID
            Next
            .MoveNext
        Loop
    End With
    If Me.cbo�ʿ�Ʒ(0).ListCount < 2 Then Me.chtThis.Header.Text = "������Ҫ�����ʿ�Ʒ���ܻ���Youdenͼ��": Exit Function
    Me.cbo�ʿ�Ʒ(0).ListIndex = 0: Me.cbo�ʿ�Ʒ(1).ListIndex = 1
    Me.cbo�ʿ�Ʒ(0).Enabled = True: Me.cbo�ʿ�Ʒ(1).Enabled = True
    
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
    Dim strLable As String
    Dim lngResIdY As Long, lngResIdX As Long
    Dim aryX() As Variant, aryY() As Variant
    
    lngResIdY = Me.cbo�ʿ�Ʒ(0).ItemData(Me.cbo�ʿ�Ʒ(0).ListIndex)
    lngResIdX = Me.cbo�ʿ�Ʒ(1).ItemData(Me.cbo�ʿ�Ʒ(1).ListIndex)
    
    '��û�����������Ϣ
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select RPad('��λ��' || '" & gstrUnitName & "', 56, ' ') || '���ڣ�' As ��0," & vbNewLine & _
            "         RPad('������' || D.����, 56, ' ') || '�Լ���Դ��' || M.�Լ� As ��1," & vbNewLine & _
            "         RPad('��Ŀ��' || I.��Ŀ, 56, ' ') || 'У׼����Դ��' || M.У׼�� As ��2" & vbNewLine & _
            "From �������� D, �����ʿ�Ʒ M, (Select ������ || ',' || Ӣ���� As ��Ŀ From ����������Ŀ Where ID = [2]) I" & vbNewLine & _
            "Where D.ID = M.����id And M.ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResIdY, mlngItemID)
    If rsTemp.RecordCount <= 0 Then Me.chtThis.Header.Text = "���ʿ�Ʒ��Ϣ��ȫ�棡": Exit Sub
    strLable = rsTemp!��0 & Format(mstrFromDate, "yyyy��MM��dd��") & "��" & Format(mstrToDate, "yyyy��MM��dd��")
    strLable = strLable & vbCrLf & rsTemp!��1 & vbCrLf & rsTemp!��2
    
    '��������������Ϊ0�����ͼ����ʾ
    Me.chtThis.ChartGroups(1).Data.NumSeries = 0
    With Me.chtThis.Header
        .Text = "�����Youdenͼ" & vbCrLf & " " & vbCrLf & " "
        .Adjust = oc2dAdjustCenter
        .Font.Bold = True
        .Font.Size = 16
    End With
    
    With Me.chtThis
        .ChartLabels.RemoveAll
        '��1
        .ChartLabels.Add
        .ChartLabels(1).AttachMethod = oc2dAttachCoord
        .ChartLabels(1).Text = rsTemp!��0 & Format(mstrFromDate, "yyyy��MM��dd��") & "��" & Format(mstrToDate, "yyyy��MM��dd��")
        .ChartLabels(1).AttachCoord.x = .Header.Location.Left + (.ChartLabels(1).Location.Width / 2) - 150
        .ChartLabels(1).AttachCoord.Y = .Header.Location.Top + .Header.Location.Height - 30
        '��2
        .ChartLabels.Add
        .ChartLabels(2).AttachMethod = oc2dAttachCoord
        .ChartLabels(2).Adjust = oc2dAdjustRight
        .ChartLabels(2).Text = rsTemp!��1
        .ChartLabels(2).AttachCoord.x = .Header.Location.Left + (.ChartLabels(2).Location.Width / 2) - 150
        .ChartLabels(2).AttachCoord.Y = .ChartLabels(1).Location.Top + .ChartLabels(1).Location.Height + 10
        '��3
        .ChartLabels.Add
        .ChartLabels(3).AttachMethod = oc2dAttachCoord
        .ChartLabels(3).Adjust = oc2dAdjustRight
        .ChartLabels(3).Text = rsTemp!��2
        .ChartLabels(3).AttachCoord.x = .Header.Location.Left + (.ChartLabels(3).Location.Width / 2) - 150
        .ChartLabels(3).AttachCoord.Y = .ChartLabels(2).Location.Top + .ChartLabels(2).Location.Height + 10
                
    End With
    
    '����ͼ�εĻ�����̬
    With Me.chtThis.ChartGroups(1)
        .ChartType = oc2dTypePlot
        With .Data
            .LayOut = oc2dDataArray
            .NumSeries = 1
            .NumPoints(1) = 4
        End With
        .Styles(1).Symbol.Shape = oc2dShapeDot: .Styles(1).Symbol.COLOR = RGB(0, 0, 160)
        .Styles(1).Line.Pattern = oc2dLineNone
    End With
    With Me.chtThis.ChartArea
        With .Axes("Y")
            .MajorGrid.Spacing.IsDefault = True
            .MajorGrid.Style.Pattern = oc2dLineSolid
            .AnnotationMethod = oc2dAnnotateValueLabels
            .Title.Text = Me.cbo�ʿ�Ʒ(0).Text
            .TitleRotation = oc2dRotate90Degrees
        End With
        With .Axes("Y2")
            .AnnotationMethod = oc2dAnnotateValueLabels
            .Multiplier = 1
        End With
        With .Axes("X")
            .MajorGrid.Spacing.IsDefault = True
            .MajorGrid.Style.Pattern = oc2dLineSolid
            .AnnotationMethod = oc2dAnnotateValueLabels
            .Title.Text = Me.cbo�ʿ�Ʒ(1).Text
        End With
    End With
    
    '������
    Dim dblAvgY As Double, dblSdY As Double, dblMaxY As Double
    Dim dblAvgX As Double, dblSdX As Double, dblMaxX As Double
    gstrSql = "Select X.�ʿ�Ʒid, X.��ֵ, Decode(X.Sd, Null, 1, 0, 1, X.Sd) As Sd" & vbNewLine & _
            "From �����ʿ�Ʒ M, �����ʿؾ�ֵ X" & vbNewLine & _
            "Where M.ID = X.�ʿ�Ʒid And (M.ID = [1] Or M.ID = [2]) And X.��Ŀid = [3] And" & vbNewLine & _
            "   Instr(';' || [4] || ';',';' || X.�ʿ�Ʒid||'='||To_char(X.��ʼ����,'yyyy-MM-dd')||','||to_char(Nvl(X.��������, M.��������),'yyyy-mm-dd')||';' ) > 0 "
            
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResIdY, lngResIdX, mlngItemID, mstr�ʿ�Ʒ����)
    With rsTemp
        Do While Not .EOF
            If lngResIdY = !�ʿ�Ʒid Then
                dblAvgY = Val("" & !��ֵ): dblSdY = Val("" & !SD):
            ElseIf lngResIdX = !�ʿ�Ʒid Then
                dblAvgX = Val("" & !��ֵ): dblSdX = Val("" & !SD):
            End If
            .MoveNext
        Loop
    End With
    If dblAvgY = 0 Or dblAvgX = 0 Or dblSdY = 0 Or dblSdX = 0 Then
        Me.chtThis.Header.Text = "��δ��ֵ��SDΪ0���޷�����" & Me.Caption & "��": Exit Sub
    End If
    With Me.chtThis.ChartArea.Axes("Y").ValueLabels
        .RemoveAll
        .Add Val(dblAvgY), Format(Val(dblAvgY), "0.00")
        .Add Val(dblAvgY) + 1 * Val(dblSdY), Format(Val(dblAvgY) + 1 * Val(dblSdY), "0.00")
        .Add Val(dblAvgY) - 1 * Val(dblSdY), Format(Val(dblAvgY) - 1 * Val(dblSdY), "0.00")
        .Add Val(dblAvgY) + 2 * Val(dblSdY), Format(Val(dblAvgY) + 2 * Val(dblSdY), "0.00")
        .Add Val(dblAvgY) - 2 * Val(dblSdY), Format(Val(dblAvgY) - 2 * Val(dblSdY), "0.00")
        .Add Val(dblAvgY) + 3 * Val(dblSdY), " ": .Add Val(dblAvgY) - 3 * Val(dblSdY), " "
    End With
    With Me.chtThis.ChartArea.Axes("Y2").ValueLabels
        .RemoveAll
        .Add Val(dblAvgY), "CL"
        .Add Val(dblAvgY) + 1 * Val(dblSdY), "CL+1SD"
        .Add Val(dblAvgY) - 1 * Val(dblSdY), "CL-1SD"
        .Add Val(dblAvgY) + 2 * Val(dblSdY), "CL+2SD"
        .Add Val(dblAvgY) - 2 * Val(dblSdY), "CL-2SD"
        .Add Val(dblAvgY) + 3 * Val(dblSdY), " "
        .Add Val(dblAvgY) - 3 * Val(dblSdY), " "
    End With
    With Me.chtThis.ChartArea.Axes("X").ValueLabels
        .RemoveAll
        .Add Val(dblAvgX), "CL=" & Format(Val(dblAvgX), "0.00")
        .Add Val(dblAvgX) + 1 * Val(dblSdX), "CL+1SD=" & Format(Val(dblAvgX) + 1 * Val(dblSdX), "0.00")
        .Add Val(dblAvgX) - 1 * Val(dblSdX), "CL-1SD=" & Format(Val(dblAvgX) - 1 * Val(dblSdX), "0.00")
        .Add Val(dblAvgX) + 2 * Val(dblSdX), "CL+2SD=" & Format(Val(dblAvgX) + 2 * Val(dblSdX), "0.00")
        .Add Val(dblAvgX) - 2 * Val(dblSdX), "CL-2SD=" & Format(Val(dblAvgX) - 2 * Val(dblSdX), "0.00")
        .Add Val(dblAvgX) + 3 * Val(dblSdX), " ": .Add Val(dblAvgX) - 3 * Val(dblSdX), " "
    End With
    
    '������֯
    gstrSql = "Select ����ʱ��, ����, Nvl(Max(Decode(�ʿ�Ʒid, [1], ���)), 0) As Y, Nvl(Max(Decode(�ʿ�Ʒid, [2], ���)), 0) As X" & vbNewLine & _
            "From (Select Q.����ʱ��, To_Char(Q.���Դ���, '000') || '-' || Decode(Nvl(T.���, 0), 0, 999, Q.���Դ���) As ����," & vbNewLine & _
            "              Q.�ʿ�Ʒid," & vbNewLine & _
            "              zl_Lis_ToNumber(Q.�ʿ�ƷID,R.������Ŀid,R.������,R.id) As ���" & vbNewLine & _
            "       From �����ʿؼ�¼ Q, ������ͨ��� R,�����ʿر��� T,�����ʿ�Ʒ M, �����ʿؾ�ֵ X " & vbNewLine & _
            "       Where Q.�걾id = R.����걾id And /*Nvl(R.�Ƿ����, 0) = 1 And*/ R.������Ŀid + 0 = [3] And" & vbNewLine & _
            "             Nvl(R.���ý��,0)=0 And R.ID=T.���ID(+) And (Q.�ʿ�Ʒid = [1] Or Q.�ʿ�Ʒid = [2]) And" & vbNewLine & _
            "             (Q.����ʱ�� Between To_Date([4], 'yyyy-MM-dd') And To_Date([5], 'yyyy-MM-dd')) And " & vbNewLine & _
            "             (Q.����ʱ�� Between X.��ʼ���� And NVL(X.��������,M.��������)) And " & vbNewLine & _
            "              Q.�ʿ�Ʒid=M.id And M.id=X.�ʿ�Ʒid  And  X.��ĿID = [3] And " & vbNewLine & _
            "             Instr(';'||[6]||';',';' || X.�ʿ�Ʒid||'='||To_char(X.��ʼ����,'yyyy-MM-dd')||','||to_char(Nvl(X.��������, M.��������),'yyyy-mm-dd')||';' ) > 0" & vbNewLine & _
            "      )" & vbNewLine & _
            "Group By ����ʱ��, ����" & vbNewLine & _
            "Order By ����ʱ��, ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResIdY, lngResIdX, mlngItemID, mstrFromDate, mstrToDate, mstr�ʿ�Ʒ����)
    With rsTemp
        If .RecordCount > 0 Then
            ReDim aryX(.RecordCount + 1)
            ReDim aryY(.RecordCount + 1, 0)
        Else
            ReDim aryX(1)
            ReDim aryY(1, 0)
        End If
        aryX(0) = Me.chtThis.ChartGroups(1).Data.HoleValue
        aryY(0, 0) = Me.chtThis.ChartGroups(1).Data.HoleValue
        Do While Not .EOF
            If !x = 0 Then
                aryX(.AbsolutePosition) = Me.chtThis.ChartGroups(1).Data.HoleValue
            Else
                aryX(.AbsolutePosition) = !x
            End If
            If !Y = 0 Then
                aryY(.AbsolutePosition, 0) = Me.chtThis.ChartGroups(1).Data.HoleValue
            Else
                aryY(.AbsolutePosition, 0) = !Y
            End If
            .MoveNext
        Loop
    End With

    '���ˢ���ڲ�����
    With Me.chtThis
        .IsBatched = True
        '����Ϊ����ʽ
        If .Width > .Height Then
            .ChartArea.Location.Height = .Height / Screen.TwipsPerPixelY - .ChartArea.Location.Top
            .ChartArea.Location.Width = .ChartArea.Location.Height + 100
        Else
            .ChartArea.Location.Width = .Width / Screen.TwipsPerPixelX - .ChartArea.Location.Left
            .ChartArea.Location.Height = .ChartArea.Location.Width - 100
        End If
        .ChartArea.Location.Left = .Width / Screen.TwipsPerPixelX / 2 - .ChartArea.Location.Width / 2
        
        With .ChartGroups(1).Data
            .NumPoints(1) = UBound(aryX) + 1
            Call .CopyXVectorIn(1, aryX)
            Call .CopyYArrayIn(aryY)
        End With
        With .ChartArea.Axes("Y")
            .Min = dblAvgY - 3 * dblSdY
            .Max = dblAvgY + 3 * dblSdY
        End With
        With .ChartArea.Axes("X")
            .Min = dblAvgX - 3 * dblSdX
            .Max = dblAvgX + 3 * dblSdX
        End With
        .IsBatched = False
        .AllowUserChanges = False
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'--------------------------------------------
'����Ϊ�ؼ��¼�����
'--------------------------------------------
Private Sub cbo�ʿ�Ʒ_Click(Index As Integer)
    Dim intBrother As Integer
    
    If Me.Visible = False Then Exit Sub
    If Me.cbo�ʿ�Ʒ(Index).Enabled = False Then Exit Sub
    
    If Index = 0 Then
        intBrother = 1
    Else
        intBrother = 0
    End If
    If Me.cbo�ʿ�Ʒ(Index).ListIndex = Me.cbo�ʿ�Ʒ(intBrother).ListIndex Then
        Me.cbo�ʿ�Ʒ(intBrother).Enabled = False
        For lngCount = 0 To Me.cbo�ʿ�Ʒ(intBrother).ListCount - 1
            If Me.cbo�ʿ�Ʒ(Index).ListIndex <> lngCount Then
                Me.cbo�ʿ�Ʒ(intBrother).ListIndex = lngCount
                Exit For
            End If
        Next
        Me.cbo�ʿ�Ʒ(intBrother).Enabled = True
    End If
    If Me.Tag = "��ˢ��" Then Exit Sub
    Call RefChart
    Me.chtThis.SetFocus
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

Private Sub chtThis_Resize(ByVal Width As Long, ByVal Height As Long)
    On Error Resume Next
    With Me.chtThis
        '��1
        .ChartLabels(1).AttachCoord.x = .Header.Location.Left + (.ChartLabels(1).Location.Width / 2) - 150
        .ChartLabels(1).AttachCoord.Y = .Header.Location.Top + .Header.Location.Height - 30
        '��2
        .ChartLabels(2).AttachCoord.x = .Header.Location.Left + (.ChartLabels(2).Location.Width / 2) - 150
        .ChartLabels(2).AttachCoord.Y = .ChartLabels(1).Location.Top + .ChartLabels(1).Location.Height + 10
        '��3
        .ChartLabels(3).AttachCoord.x = .Header.Location.Left + (.ChartLabels(3).Location.Width / 2) - 150
        .ChartLabels(3).AttachCoord.Y = .ChartLabels(2).Location.Top + .ChartLabels(2).Location.Height + 10
        
        If .Width > .Height Then
            .ChartArea.Location.Height = .Height / Screen.TwipsPerPixelY - .ChartArea.Location.Top
            .ChartArea.Location.Width = .ChartArea.Location.Height + 100
        Else
            .ChartArea.Location.Width = .Width / Screen.TwipsPerPixelX - .ChartArea.Location.Left
            .ChartArea.Location.Height = .ChartArea.Location.Width - 100
        End If
        
        .ChartArea.Location.Left = .Width / Screen.TwipsPerPixelX / 2 - .ChartArea.Location.Width / 2
    End With
    
    
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With Me.cbo�ʿ�Ʒ(0)
        .Left = Me.lbl�ʿ�Ʒ(0).Width + Screen.TwipsPerPixelX * 4
        .Width = Me.ScaleWidth / 2 - .Left
        .Top = Me.ScaleHeight - .Height
    End With
    With Me.lbl�ʿ�Ʒ(0)
        .Left = Screen.TwipsPerPixelX * 2
        .Top = Me.cbo�ʿ�Ʒ(0).Top + (Me.cbo�ʿ�Ʒ(0).Height - .Height) / 2
    End With
    
    With Me.cbo�ʿ�Ʒ(1)
        .Left = Me.ScaleWidth / 2 + Me.lbl�ʿ�Ʒ(1).Width + Screen.TwipsPerPixelX * 4
        .Width = Me.ScaleWidth - .Left
        .Top = Me.ScaleHeight - .Height
    End With
    With Me.lbl�ʿ�Ʒ(1)
        .Left = Me.ScaleWidth / 2 + Screen.TwipsPerPixelX * 2
        .Top = Me.cbo�ʿ�Ʒ(1).Top + (Me.cbo�ʿ�Ʒ(1).Height - .Height) / 2
    End With
    
    With Me.chtThis
        .Left = 0: .Width = Me.ScaleWidth
        .Top = 0: .Height = Me.ScaleHeight - .Top - Me.cbo�ʿ�Ʒ(0).Height
    End With
End Sub
