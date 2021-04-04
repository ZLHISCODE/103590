VERSION 5.00
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmQCChartZS 
   BorderStyle     =   0  'None
   Caption         =   "Z-����ͼ"
   ClientHeight    =   5352
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7752
   LinkTopic       =   "Form1"
   ScaleHeight     =   5352
   ScaleWidth      =   7752
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CheckBox chk�ʿ�Ʒ 
      Caption         =   "473843A��ֵ�ʿ�Ʒ"
      Enabled         =   0   'False
      Height          =   240
      Index           =   0
      Left            =   105
      TabIndex        =   1
      Top             =   5055
      Value           =   1  'Checked
      Width           =   1830
   End
   Begin C1Chart2D8.Chart2D chtThis 
      Height          =   3690
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   6630
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   11695
      _ExtentY        =   6509
      _StockProps     =   0
      ControlProperties=   "frmQCChartZS.frx":0000
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmQCChartZS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrResList As String
Private mlngItemID As Long
Private mstrFromDate As String
Private mstrToDate As String
Private mdblAVG As Double                           '��ֵ
Private mdblSD As Double                            'SD
Private mintFormatNum As Integer                    '��ʽ��С��λ��
Private mstr�ʿ�Ʒ���� As String
Dim lngCount As Long

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
    '       str�ʿ�Ʒ���� ��ʽ��: �ʿ�Ʒid=��ʼ���ڣ��������ڣ� �� ;�ָ�����ʿ�Ʒ��
    Dim rsTemp As New ADODB.Recordset
    Dim intCounts As Integer
    Dim lngResId As Long
    
    mstrResList = strResList
    mlngItemID = lngItemID
    mstrFromDate = strFromDate
    mstrToDate = strToDate
    mstr�ʿ�Ʒ���� = str�ʿ�Ʒ����
    
    intCounts = Me.chk�ʿ�Ʒ.Count
    For lngCount = intCounts - 1 To 1 Step -1
        Unload Me.chk�ʿ�Ʒ(Me.chk�ʿ�Ʒ.UBound)
    Next
    Me.chk�ʿ�Ʒ(0).Enabled = False
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select ID, ���� || '-' || ���� As �ʿ�Ʒ From �����ʿ�Ʒ Where Instr(',' || [1] || ',', ',' || ID || ',') > 0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strResList)
    With rsTemp
        Do While Not .EOF
            If .AbsolutePosition > Me.chk�ʿ�Ʒ.Count Then Load Me.chk�ʿ�Ʒ(.AbsolutePosition - 1)
            Me.chk�ʿ�Ʒ(.AbsolutePosition - 1).Caption = !�ʿ�Ʒ & " (��" & AskLevelNote(.AbsolutePosition) & "��)"
            Me.chk�ʿ�Ʒ(.AbsolutePosition - 1).Tag = !ID
            Me.chk�ʿ�Ʒ(.AbsolutePosition - 1).Width = Me.TextWidth(Me.chk�ʿ�Ʒ(.AbsolutePosition - 1).Caption) + 360
            Me.chk�ʿ�Ʒ(.AbsolutePosition - 1).Value = vbChecked
            Me.chk�ʿ�Ʒ(.AbsolutePosition - 1).Visible = True
            Me.chk�ʿ�Ʒ(.AbsolutePosition - 1).Enabled = True
            .MoveNext
        Loop
    End With
    
    Call RefChart
    Call Form_Resize
    zlRefresh = True
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function AskLevelShape(lngLevel As Long) As Long
    '���ܣ�ȷ����ͬ����ʿ�Ʒ����������״
    Select Case lngLevel
    Case 1: AskLevelShape = oc2dShapeDot
    Case 2: AskLevelShape = oc2dShapeBox
    Case 3: AskLevelShape = oc2dShapeTriangle
    Case 4: AskLevelShape = oc2dShapeDiamond
    Case 5: AskLevelShape = oc2dShapeStar
    Case 6: AskLevelShape = oc2dShapeCircle
    Case 7: AskLevelShape = oc2dShapeSquare
    Case 8: AskLevelShape = oc2dShapeOpenTriangle
    Case 9: AskLevelShape = oc2dShapeOpenDiamond
    Case Else: AskLevelShape = oc2dShapeCross
    End Select
End Function

Private Function AskLevelNote(lngLevel As Long) As String
    '���ܣ�ȷ����ͬ����ʿ�Ʒ����������״˵��
    Select Case lngLevel
    Case 1: AskLevelNote = "��"
    Case 2: AskLevelNote = "��"
    Case 3: AskLevelNote = "��"
    Case 4: AskLevelNote = "��"
    Case 5: AskLevelNote = "��"
    Case 6: AskLevelNote = "��"
    Case 7: AskLevelNote = "��"
    Case 8: AskLevelNote = "��"
    Case 9: AskLevelNote = "��"
    Case Else: AskLevelNote = "��"
    End Select
End Function

Private Function AskLevelColor(lngLevel As Long) As Long
    '���ܣ�ȷ����ͬ����ʿ�Ʒ��������ɫ
    Select Case lngLevel
    Case 1: AskLevelColor = RGB(0, 0, 160)
    Case 2: AskLevelColor = RGB(0, 128, 255)
    Case 3: AskLevelColor = RGB(0, 128, 64)
    Case 4: AskLevelColor = RGB(0, 64, 128)
    Case 5: AskLevelColor = RGB(64, 128, 128)
    Case 6: AskLevelColor = RGB(128, 128, 192)
    Case 7: AskLevelColor = RGB(128, 128, 64)
    Case 8: AskLevelColor = RGB(128, 128, 128)
    Case 9: AskLevelColor = RGB(0, 255, 64)
    Case Else: AskLevelColor = RGB(0, 0, 0)
    End Select
    
End Function

Private Sub RefChart()
    '���ܣ�ˢ��ͼ����ʾ
    Dim rsTemp As New ADODB.Recordset
    Dim strLable As String, intRow As Integer, intCol As Integer
    Dim dblMax As Double
    Dim aryX() As Variant, aryY() As Variant
    Dim intLoop As Integer, dateLoop As Date
    Dim bln�ϲ��� As Boolean
    Dim strLastData As String, strLastBadData As String
    Dim dblAvg1 As Double, dblSD1 As Double, dblAvg2 As Double, dblSD2 As Double, lngAVGcount As Long, intFormatNum As Integer
    '��ȡС��λ��
    gstrSql = "Select С��λ�� from ����������Ŀ where ��ĿID = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, mlngItemID)
    If rsTemp.EOF = False Then mintFormatNum = Val(Nvl(rsTemp("С��λ��"), 2))
    intFormatNum = mintFormatNum
    
    '--- ȡ��ֵ��SD
    gstrSql = "Select x.�ʿ�Ʒid, x.��ֵ, Decode(x.Sd, Null, 1, 0, 1, x.Sd) As Sd, x.��ʼ����, Nvl(x.��������, m.��������) As ��������" & vbNewLine & _
            "From �����ʿ�Ʒ M, �����ʿؾ�ֵ X" & vbNewLine & _
            "Where m.Id = x.�ʿ�Ʒid And x.��Ŀid = [1] And" & vbNewLine & _
            "      Instr(';'|| [2] ||';',';' || x.�ʿ�Ʒid || '=' || To_Char(x.��ʼ����, 'yyyy-MM-dd') || ',' || To_Char(Nvl(x.��������, m.��������), 'yyyy-mm-dd') || ';') > 0" & _
            " order by  �ʿ�Ʒid"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngItemID, mstr�ʿ�Ʒ����)
    lngAVGcount = 0
    Do Until rsTemp.EOF
    
        For lngCount = 0 To Me.chk�ʿ�Ʒ.Count - 1
            If Me.chk�ʿ�Ʒ(lngCount).Value = 1 Then
                If Val("" & rsTemp!�ʿ�Ʒid) = Val("" & Me.chk�ʿ�Ʒ(lngCount).Tag) Then
                    If lngAVGcount = 0 Then
                        dblAvg1 = Val("" & rsTemp!��ֵ)
                        dblSD1 = Val("" & rsTemp!SD)
                        '��ֵ��SD��Ĭ��С�����ĸ����ȸ����ĸ�
                        If InStr("" & rsTemp!��ֵ, ".") > 0 Then
                            If intFormatNum < Len(Mid("" & rsTemp!��ֵ, InStr("" & rsTemp!��ֵ, ".") + 1)) Then
                                intFormatNum = Len(Mid("" & rsTemp!��ֵ, InStr("" & rsTemp!��ֵ, ".") + 1))
                            End If
                        Else
                            If intFormatNum < 0 Then intFormatNum = 0
                        End If
                        
                        If InStr("" & rsTemp!SD, ".") > 0 Then
                            If intFormatNum < Len(Mid("" & rsTemp!SD, InStr("" & rsTemp!SD, ".") + 1)) Then
                                intFormatNum = Len(Mid("" & rsTemp!SD, InStr("" & rsTemp!SD, ".") + 1))
                            End If
                        Else
                            If intFormatNum < 0 Then intFormatNum = 0
                        End If
                        
                        lngAVGcount = lngAVGcount + 1
                    Else
                        dblAvg2 = Val("" & rsTemp!��ֵ)
                        dblSD2 = Val("" & rsTemp!SD)
                        lngAVGcount = lngAVGcount + 1
                        Exit Do
                    End If
                End If
            End If
        Next
        rsTemp.MoveNext
    Loop
        
        
    '��û�����������Ϣ
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select Distinct RPad('��λ��' || '" & gstrUnitName & "', 56, ' ') || '���ڣ�' As ��0," & vbNewLine & _
            "                RPad('������' || D.����, 56, ' ') || '�Լ���Դ��' || M.�Լ� As ��1," & vbNewLine & _
            "                RPad('��Ŀ��' || I.��Ŀ, 56, ' ') || 'У׼����Դ��' || M.У׼�� As ��2" & vbNewLine & _
            "From �������� D, �����ʿ�Ʒ M, (Select ������ || ',' || Ӣ���� As ��Ŀ From ����������Ŀ Where ID = [2]) I" & vbNewLine & _
            "Where D.ID = M.����id And Instr(',' || [1] || ',', ',' || M.ID || ',') > 0"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrResList, mlngItemID)
    If rsTemp.RecordCount <= 0 Then Me.chtThis.Header.Text = "���ʿ�Ʒ��Ϣ��ȫ�棡": Exit Sub
    strLable = rsTemp!��0 & Format(mstrFromDate, "yyyy��MM��dd��") & "��" & Format(mstrToDate, "yyyy��MM��dd��")
    strLable = strLable & vbCrLf & rsTemp!��1 & vbCrLf & rsTemp!��2
    
    '��������������Ϊ0�����ͼ����ʾ
    Me.chtThis.ChartGroups(1).Data.NumSeries = 0
    With Me.chtThis.Header
        .Text = "�����Z-����ͼ" & vbCrLf & " " & vbCrLf & " "
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
        .ChartLabels(1).AttachCoord.y = .Header.Location.Top + .Header.Location.Height - 30
        '��2
        .ChartLabels.Add
        .ChartLabels(2).AttachMethod = oc2dAttachCoord
        .ChartLabels(2).Adjust = oc2dAdjustRight
        .ChartLabels(2).Text = rsTemp!��1
        .ChartLabels(2).AttachCoord.x = .Header.Location.Left + (.ChartLabels(2).Location.Width / 2) - 150
        .ChartLabels(2).AttachCoord.y = .ChartLabels(1).Location.Top + .ChartLabels(1).Location.Height + 10
        '��3
        .ChartLabels.Add
        .ChartLabels(3).AttachMethod = oc2dAttachCoord
        .ChartLabels(3).Adjust = oc2dAdjustRight
        .ChartLabels(3).Text = rsTemp!��2
        .ChartLabels(3).AttachCoord.x = .Header.Location.Left + (.ChartLabels(3).Location.Width / 2) - 150
        .ChartLabels(3).AttachCoord.y = .ChartLabels(2).Location.Top + .ChartLabels(2).Location.Height + 10
        
    End With
    
    With Me.chtThis.Footer
        .Text = ""
        For lngCount = 0 To Me.chk�ʿ�Ʒ.Count - 1
            If chk�ʿ�Ʒ(lngCount).Value = 1 Then
                .Text = .Text & Space(6) & Me.chk�ʿ�Ʒ(lngCount).Caption
            End If
        Next
        .Text = Trim(.Text)
    End With
    
    '����ͼ�εĻ�����̬
    With Me.chtThis.ChartGroups(1)
        .ChartType = oc2dTypePlot
        With .Data
            .LayOut = oc2dDataArray
            .NumSeries = 9 + Me.chk�ʿ�Ʒ.Count * 6
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
        For lngCount = 1 To Me.chk�ʿ�Ʒ.Count
            Me.chk�ʿ�Ʒ(lngCount - 1).ForeColor = AskLevelColor(lngCount)
            .Styles(9 + lngCount * 6 - 5).Symbol.Shape = AskLevelShape(lngCount)
            .Styles(9 + lngCount * 6 - 5).Line.COLOR = Me.chk�ʿ�Ʒ(lngCount - 1).ForeColor
            .Styles(9 + lngCount * 6 - 5).Symbol.COLOR = Me.chk�ʿ�Ʒ(lngCount - 1).ForeColor
            
            .Styles(9 + lngCount * 6 - 4).Symbol.Shape = .Styles(9 + lngCount * 6 - 5).Symbol.Shape
            .Styles(9 + lngCount * 6 - 4).Line.Pattern = oc2dLineNone
            .Styles(9 + lngCount * 6 - 4).Symbol.COLOR = RGB(255, 0, 0)
            
            .Styles(9 + lngCount * 6 - 3).Symbol.Shape = .Styles(9 + lngCount * 6 - 5).Symbol.Shape
            .Styles(9 + lngCount * 6 - 3).Line.Pattern = oc2dLineNone
            .Styles(9 + lngCount * 6 - 3).Symbol.COLOR = RGB(255, 0, 0)
            
            .Styles(9 + lngCount * 6 - 2).Symbol.Shape = .Styles(9 + lngCount * 6 - 5).Symbol.Shape
            .Styles(9 + lngCount * 6 - 2).Line.Pattern = oc2dLineNone
            .Styles(9 + lngCount * 6 - 2).Symbol.COLOR = RGB(255, 0, 0)
            
            .Styles(9 + lngCount * 6 - 1).Symbol.Shape = .Styles(9 + lngCount * 6 - 5).Symbol.Shape
            .Styles(9 + lngCount * 6 - 1).Line.Pattern = oc2dLineNone
            .Styles(9 + lngCount * 6 - 1).Symbol.COLOR = RGB(255, 0, 0)
            
            .Styles(9 + lngCount * 6).Symbol.Shape = .Styles(9 + lngCount * 6 - 5).Symbol.Shape
            .Styles(9 + lngCount * 6).Line.Pattern = oc2dLineNone
            .Styles(9 + lngCount * 6).Symbol.COLOR = RGB(255, 0, 0)
            Call chk�ʿ�Ʒ_Click(CInt(lngCount - 1))
        Next
    End With
    With Me.chtThis.ChartArea.Axes("Y")
        .MajorGrid.Spacing.IsDefault = False
        .AnnotationMethod = oc2dAnnotateValueLabels
        .Title.Text = "�ⶨƫ��(SD)"
        With .ValueLabels
            .RemoveAll
            If lngAVGcount = 1 Then
                .Add 4, Format(Val(dblAvg1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 + 1, Format(Val(dblAvg1) + 1 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 - 1, Format(Val(dblAvg1) - 1 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 + 2, Format(Val(dblAvg1) + 2 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 - 2, Format(Val(dblAvg1) - 2 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 + 3, Format(Val(dblAvg1) + 3 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 - 3, Format(Val(dblAvg1) - 3 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 + 4, Format(Val(dblAvg1) + 4 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 - 4, Format(Val(dblAvg1) - 4 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
            
            ElseIf lngAVGcount = 2 Then
                .Add 4, Format(Val(dblAvg1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0"))) & " " & Format(Val(dblAvg2), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 + 1, Format(Val(dblAvg1) + 1 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0"))) & " " & Format(Val(dblAvg2) + 1 * Val(dblSD2), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 - 1, Format(Val(dblAvg1) - 1 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0"))) & " " & Format(Val(dblAvg2) - 1 * Val(dblSD2), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 + 2, Format(Val(dblAvg1) + 2 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0"))) & " " & Format(Val(dblAvg2) + 2 * Val(dblSD2), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 - 2, Format(Val(dblAvg1) - 2 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0"))) & " " & Format(Val(dblAvg2) - 2 * Val(dblSD2), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 + 3, Format(Val(dblAvg1) + 3 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0"))) & " " & Format(Val(dblAvg2) + 3 * Val(dblSD2), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 - 3, Format(Val(dblAvg1) - 3 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0"))) & " " & Format(Val(dblAvg2) - 3 * Val(dblSD2), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 + 4, Format(Val(dblAvg1) + 4 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0"))) & " " & Format(Val(dblAvg2) + 4 * Val(dblSD2), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
                .Add 4 - 4, Format(Val(dblAvg1) - 4 * Val(dblSD1), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0"))) & " " & Format(Val(dblAvg2) - 4 * Val(dblSD2), IIf(intFormatNum = 0, "0", "0." & String(intFormatNum, "0")))
            Else
                .Add 4, 0
                .Add 4 + 1, 1
                .Add 4 - 1, -1
                .Add 4 + 2, 2
                .Add 4 - 2, -2
                .Add 4 + 3, 3
                .Add 4 - 3, -3
                .Add 4 + 4, 4
                .Add 4 - 4, -4
                
            End If
        End With
    End With
    With Me.chtThis.ChartArea.Axes("Y2")
        .AnnotationMethod = oc2dAnnotateValueLabels
        .Title.Text = "������"
        .Multiplier = 1
        With .ValueLabels
            .RemoveAll
            .Add 4, "CL"
            .Add 4 + 1, "CL+1SD": .Add 4 - 1, "CL-1SD"
            .Add 4 + 2, "CL+2SD": .Add 4 - 2, "CL-2SD"
            .Add 4 + 3, "CL+3SD": .Add 4 - 3, "CL-3SD"
        End With
    End With
    With Me.chtThis.ChartArea.Axes("X")
        .MajorGrid.Spacing.IsDefault = False
        .AnnotationMethod = oc2dAnnotateValueLabels
        .Title.Text = "����"
    End With
    
    '������֯
'    gstrSql = "Select Q.����ʱ��, Q.���Դ���, Q.�ʿ�Ʒid,x.��ֵ,x.SD, Round(Max(Decode(Q.���, 2, 0, (Q.��� - X.��ֵ) / X.Sd)), 4) As �ڿ�," & vbNewLine & _
            "       Round(Max(Decode(���, 2, (Q.��� - X.��ֵ) / X.Sd, 0)), 4) As ʧ��" & vbNewLine & _
            "From (Select Q.����ʱ��, Q.���Դ���, Q.�ʿ�Ʒid, T.���," & vbNewLine & _
            "              Decode(I.ֵ����, Null, Zl_To_Number(R.������)," & vbNewLine & _
            "                      Length(Substr(I.ֵ����, 1, Instr(I.ֵ����, ';' || RTrim(R.������) || ';'))) -" & vbNewLine & _
            "                       Nvl(Length(Replace(Substr(I.ֵ����, 1, Instr(I.ֵ����, ';' || RTrim(R.������) || ';')), ';')), 0)) As ���" & vbNewLine & _
            "       From �����ʿؼ�¼ Q, ������ͨ��� R, �����ʿر��� T," & vbNewLine & _
            "            (Select Decode(�������, 3, Decode(RTrim(ȡֵ����), '', '', ';' || RTrim(ȡֵ����) || ';'), '') As ֵ����" & vbNewLine & _
            "              From ������Ŀ" & vbNewLine & _
            "              Where ������Ŀid = [2]) I" & vbNewLine & _
            "       Where Q.�걾id = R.����걾id And R.ID = T.���id(+) And /*Nvl(R.�Ƿ����, 0) = 1 And*/ " & vbNewLine & _
            "             Instr(',' || [1] || ',', ',' || Q.�ʿ�Ʒid || ',') > 0 And R.������Ŀid + 0 = [2] And" & vbNewLine & _
            "             (Q.����ʱ�� Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd'))) Q," & vbNewLine & _
            "     (Select X.�ʿ�Ʒid, X.��ֵ, Decode(X.Sd, Null, 1, 0, 1, X.Sd) As Sd" & vbNewLine & _
            "       From �����ʿ�Ʒ M, �����ʿؾ�ֵ X" & vbNewLine & _
            "       Where M.ID = X.�ʿ�Ʒid And Instr(',' || [1] || ',', ',' || X.�ʿ�Ʒid || ',') > 0 And X.��Ŀid = [2] And" & vbNewLine & _
            "             (To_Date([3], 'yyyy-MM-dd') Between X.��ʼ���� And Nvl(X.��������, M.��������)) And" & vbNewLine & _
            "             (To_Date([4], 'yyyy-MM-dd') Between X.��ʼ���� And Nvl(X.��������, M.��������))) X" & vbNewLine & _
            "Where Q.�ʿ�Ʒid = X.�ʿ�Ʒid" & vbNewLine & _
            "Group By Q.����ʱ��, Q.���Դ���, Q.�ʿ�Ʒid,x.��ֵ,x.SD " & vbNewLine & _
            "Order By Q.����ʱ��, Q.���Դ���"
            
    gstrSql = "Select ����ʱ��, ���Դ���, �ʿ�Ʒid, ��ֵ, Sd," & vbNewLine & _
                "       Max(�ڿ�) As �ڿ�,max(ʧ��1) As ʧ��1,max(ʧ��2) As ʧ��2,max(ʧ��3) As ʧ��3,max(ʧ��4) As ʧ��4 ,max(ʧ��5) As ʧ��5" & vbNewLine & _
                "From (Select Q.����ʱ��, Q.���Դ���, Q.�ʿ�Ʒid, X.��ֵ, X.Sd, to_char(Round(Max(Decode(Q.���, 2, 0, decode(Q.���,null,'',(Q.��� - X.��ֵ) / X.Sd))), 4)) As �ڿ�, 0 As ʧ��1," & vbNewLine & _
                "              0 As ʧ��2, 0 As ʧ��3, 0 As ʧ��4, 0 As ʧ��5" & vbNewLine & _
                "       From (Select Q.����ʱ��, Q.���Դ���, Q.�ʿ�Ʒid, T.���," & vbNewLine & _
                "                     zl_Lis_ToNumber(Q.�ʿ�ƷID,R.������Ŀid,R.������,R.id ) As ���" & vbNewLine & _
                "              From �����ʿؼ�¼ Q, ������ͨ��� R, �����ʿر��� T" & vbNewLine & _
                "              Where Q.�걾id = R.����걾id And R.ID = T.���id(+) And Nvl(R.���ý��,0)=0 And /*Nvl(R.�Ƿ����, 0) = 1 And*/" & vbNewLine & _
                "                    Instr(',' || [1] || ',', ',' || Q.�ʿ�ƷID || ',') > 0 And R.������Ŀid + 0 = [2] And" & vbNewLine & _
                "                    (Q.����ʱ�� Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd')) " & vbNewLine & _
                "                    " & vbNewLine & _
                "      ) Q, "
    gstrSql = gstrSql & "" & vbNewLine & _
                "     (Select X.�ʿ�Ʒid, X.��ֵ, Decode(X.Sd, Null, 1, 0, 1, X.Sd) As Sd, X.��ʼ����, Nvl(X.��������, M.��������) As ��������" & vbNewLine & _
                "       From �����ʿ�Ʒ M, �����ʿؾ�ֵ X" & vbNewLine & _
                "       Where M.ID = X.�ʿ�Ʒid And X.��Ŀid = [2] And" & vbNewLine & _
                "         Instr(';' || [5] || ';',';' || X.�ʿ�Ʒid||'='||To_char(X.��ʼ����,'yyyy-MM-dd')||','||to_char(Nvl(X.��������, M.��������),'yyyy-mm-dd')||';' ) > 0  " & vbNewLine & _
                "      ) X" & vbNewLine & _
                "Where Nvl(Q.���, 0) <> 2 And Q.�ʿ�Ʒid = X.�ʿ�Ʒid And  Q.����ʱ�� Between X.��ʼ���� And X.��������" & vbNewLine & _
                "Group By Q.����ʱ��, Q.���Դ���, Q.�ʿ�Ʒid, X.��ֵ, X.Sd" & vbNewLine & _
                "" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select Q.����ʱ��, Q.���Դ���, Q.�ʿ�Ʒid, X.��ֵ, X.Sd, '' As �ڿ�, Round(Max(Decode(���, 2, (Q.���1 - X.��ֵ) / X.Sd, 0)), 4) As ʧ��1," & vbNewLine & _
                "       Round(Max(Decode(���, 2, (Q.���2 - X.��ֵ) / X.Sd, 0)), 4) As ʧ��2," & vbNewLine & _
                "       Round(Max(Decode(���, 2, (Q.���3 - X.��ֵ) / X.Sd, 0)), 4) As ʧ��3," & vbNewLine & _
                "       Round(Max(Decode(���, 2, (Q.���4 - X.��ֵ) / X.Sd, 0)), 4) As ʧ��4," & vbNewLine & _
                "       Round(Max(Decode(���, 2, (Q.���5 - X.��ֵ) / X.Sd, 0)), 4) As ʧ��5" & vbNewLine & _
                "From (Select ����ʱ��, ���Դ���, �ʿ�Ʒid, ���, Max(Decode(Mod(�к�, 5), 1, ���, '')) As ���1," & vbNewLine & _
                "              Max(Decode(Mod(�к�, 5), 2, ���, '')) As ���2, Max(Decode(Mod(�к�, 5), 3, ���, '')) As ���3," & vbNewLine & _
                "              Max(Decode(Mod(�к�, 5), 4, ���, '')) As ���4, Max(Decode(Mod(�к�, 5), 0, ���, '')) As ���5" & vbNewLine & _
                "       From (Select Q.����ʱ��, Q.���Դ���, Q.�ʿ�Ʒid, T.���," & vbNewLine & _
                "                     zl_Lis_ToNumber(Q.�ʿ�ƷID,R.������Ŀid,R.������,R.id) As ���," & vbNewLine & _
                "                     Rownum As �к�" & vbNewLine & _
                "              From �����ʿؼ�¼ Q, ������ͨ��� R, �����ʿر��� T "
    gstrSql = gstrSql & "" & vbNewLine & _
                "                     Where Q.�걾id = R.����걾id And R.ID = T.���id And Nvl(R.���ý��,0)=0 And /*Nvl(R.�Ƿ����, 0) = 1 And*/" & vbNewLine & _
                "                           Instr(',' || [1] || ',', ',' || Q.�ʿ�Ʒid || ',') > 0 And R.������Ŀid + 0 = [2] And" & vbNewLine & _
                "                           (Q.����ʱ�� Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd')) And" & vbNewLine & _
                "                           Nvl(T.���, 0) = 2  " & vbNewLine & _
                "                           )" & vbNewLine & _
                "              Group By ����ʱ��, ���Դ���, �ʿ�Ʒid, ���) Q," & vbNewLine & _
                "     (Select X.�ʿ�Ʒid, X.��ֵ, Decode(X.Sd, Null, 1, 0, 1, X.Sd) As Sd, X.��ʼ����, Nvl(X.��������, M.��������) As ��������" & vbNewLine & _
                "       From �����ʿ�Ʒ M, �����ʿؾ�ֵ X" & vbNewLine & _
                "       Where M.ID = X.�ʿ�Ʒid And X.��Ŀid = [2] And" & vbNewLine & _
                "         Instr(';' || [5] || ';',';' || X.�ʿ�Ʒid||'='||To_char(X.��ʼ����,'yyyy-MM-dd')||','||to_char(Nvl(X.��������, M.��������),'yyyy-mm-dd')||';' ) > 0  " & vbNewLine & _
                "      ) X" & vbNewLine & _
                "       Where Q.�ʿ�Ʒid = X.�ʿ�Ʒid And  Q.����ʱ�� Between X.��ʼ���� And X.�������� " & vbNewLine & _
                "       Group By Q.����ʱ��, Q.���Դ���, Q.�ʿ�Ʒid, X.��ֵ, X.Sd)" & vbNewLine & _
                "group by ����ʱ��, ���Դ���, �ʿ�Ʒid, ��ֵ, Sd order by ����ʱ��,���Դ���"



    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrResList, mlngItemID, mstrFromDate, mstrToDate, mstr�ʿ�Ʒ����)
    Me.chtThis.ChartArea.Axes("X").ValueLabels.RemoveAll
    With rsTemp
        strLable = "": intRow = 0: strLastData = ""
        Do While Not .EOF
            '-1.�����ڣ�Ӧ�ӱ�ǩ
            '-2.������ͬ�����Դ�����ͬ,�ϴκͱ��ξ��ڿأ����,ʧ�ز���
            
            If strLable <> Format(!����ʱ��, "yyyy-MM-dd") Then
                intRow = intRow + 1
            ElseIf strLable = Format(!����ʱ��, "yyyy-MM-dd") And strLastData <> Trim("" & !���Դ���) And _
                (Trim("" & !�ڿ�) <> "" And strLastBadData <> "") Then
                intRow = intRow + 1

            End If
            strLable = Format(!����ʱ��, "yyyy-MM-dd")
            strLastData = Trim("" & !���Դ���)
            strLastBadData = Trim("" & !�ڿ�)
            .MoveNext
        Loop
        If intRow < 30 Then
            intLoop = intRow
            ReDim aryX(31)
            ReDim aryY(31, 8 + Me.chk�ʿ�Ʒ.Count * 6)
        Else
            intLoop = 0
            ReDim aryX(intRow)
            ReDim aryY(intRow, 8 + Me.chk�ʿ�Ʒ.Count * 6)
        End If
        aryY(0, 0) = 4
        aryY(0, 1) = 4 + 1: aryY(0, 2) = 4 - 1
        aryY(0, 3) = 4 + 2: aryY(0, 4) = 4 - 2
        aryY(0, 5) = 4 + 3: aryY(0, 6) = 4 - 3
        aryY(0, 7) = 4 + 4: aryY(0, 8) = 4 - 4
        For lngCount = 1 To Me.chk�ʿ�Ʒ.Count
            aryY(0, 8 + lngCount * 6 - 5) = Me.chtThis.ChartGroups(1).Data.HoleValue
            aryY(0, 8 + lngCount * 6 - 4) = Me.chtThis.ChartGroups(1).Data.HoleValue
            aryY(0, 8 + lngCount * 6 - 3) = Me.chtThis.ChartGroups(1).Data.HoleValue
            aryY(0, 8 + lngCount * 6 - 2) = Me.chtThis.ChartGroups(1).Data.HoleValue
            aryY(0, 8 + lngCount * 6 - 1) = Me.chtThis.ChartGroups(1).Data.HoleValue
            aryY(0, 8 + lngCount * 6) = Me.chtThis.ChartGroups(1).Data.HoleValue
        Next
        dblMax = 4
        strLable = "": intRow = 0
        If .RecordCount > 0 Then .MoveFirst
        strLastData = ""
        Do While Not .EOF
            mdblAVG = Val(Nvl(!��ֵ))
            mdblSD = Val(Nvl(!SD))
            If mdblAVG = 0 Or mdblSD = 0 Then
                Me.chtThis.Header.Text = "��δ��ֵ��SDΪ0���޷�����" & Me.Caption & "��": Exit Sub
            End If
            If strLable <> Format(!����ʱ��, "yyyy-MM-dd") Then
                
                intRow = intRow + 1
'                Me.ChtThis.ChartArea.Axes("X").ValueLabels.Add intRow, intRow
                
                bln�ϲ��� = False
                
                If Format(Nvl(!����ʱ��), "dd") <> "01" Then
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add intRow, Format(Nvl(!����ʱ��), "dd")
                Else
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add intRow, Format(Nvl(!����ʱ��), "mm" & "��")
                End If
                dateLoop = Format(Nvl(!����ʱ��), "yyyy-MM-dd")
                aryX(intRow) = intRow
                aryY(intRow, 0) = 4
                aryY(intRow, 1) = 4 + 1: aryY(intRow, 2) = 4 - 1
                aryY(intRow, 3) = 4 + 2: aryY(intRow, 4) = 4 - 2
                aryY(intRow, 5) = 4 + 3: aryY(intRow, 6) = 4 - 3
                aryY(intRow, 7) = 4 + 4: aryY(intRow, 8) = 4 - 4
                For lngCount = 1 To Me.chk�ʿ�Ʒ.Count
                    aryY(intRow, 8 + lngCount * 6 - 5) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    aryY(intRow, 8 + lngCount * 6 - 4) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    aryY(intRow, 8 + lngCount * 6 - 3) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    aryY(intRow, 8 + lngCount * 6 - 2) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    aryY(intRow, 8 + lngCount * 6 - 1) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    aryY(intRow, 8 + lngCount * 6) = Me.chtThis.ChartGroups(1).Data.HoleValue
                Next
            ElseIf strLable = Format(!����ʱ��, "yyyy-MM-dd") And strLastData <> Trim("" & !���Դ���) And _
                (Trim("" & !�ڿ�) <> "" And strLastBadData <> "") Then
                bln�ϲ��� = False
                intRow = intRow + 1

                If Format(Nvl(!����ʱ��), "dd") <> "01" Then
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add intRow, Format(Nvl(!����ʱ��), "dd")
                Else
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add intRow, Format(Nvl(!����ʱ��), "mm" & "��")
                End If
                dateLoop = Format(Nvl(!����ʱ��), "yyyy-MM-dd")
                aryX(intRow) = intRow
                aryY(intRow, 0) = 4
                aryY(intRow, 1) = 4 + 1: aryY(intRow, 2) = 4 - 1
                aryY(intRow, 3) = 4 + 2: aryY(intRow, 4) = 4 - 2
                aryY(intRow, 5) = 4 + 3: aryY(intRow, 6) = 4 - 3
                aryY(intRow, 7) = 4 + 4: aryY(intRow, 8) = 4 - 4
                For lngCount = 1 To Me.chk�ʿ�Ʒ.Count
                    aryY(intRow, 8 + lngCount * 6 - 5) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    aryY(intRow, 8 + lngCount * 6 - 4) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    aryY(intRow, 8 + lngCount * 6 - 3) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    aryY(intRow, 8 + lngCount * 6 - 2) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    aryY(intRow, 8 + lngCount * 6 - 1) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    aryY(intRow, 8 + lngCount * 6) = Me.chtThis.ChartGroups(1).Data.HoleValue
                Next
            Else
                bln�ϲ��� = True
            End If
            
            strLable = Format(!����ʱ��, "yyyy-MM-dd")
            strLastData = Trim("" & !���Դ���)
            strLastBadData = Trim("" & !�ڿ�)
                
            For lngCount = 1 To Me.chk�ʿ�Ʒ.Count
                If Val(Me.chk�ʿ�Ʒ(lngCount - 1).Tag) = Val("" & !�ʿ�Ʒid) Then
                    
                    '��������������ֵʱ��Ϊ��������ֵ
                    If Abs(Val("" & !�ڿ�)) > 4 Then
                        aryY(intRow, 8 + lngCount * 6 - 5) = 4 + IIf(Val("" & !�ڿ�) < -4, -4, 4)
                    Else
                        If Trim("" & !�ڿ�) = "" And bln�ϲ��� = False Then
                            aryY(intRow, 8 + lngCount * 6 - 5) = Me.chtThis.ChartGroups(1).Data.HoleValue
                        ElseIf Trim("" & !�ڿ�) <> "" Then
                            aryY(intRow, 8 + lngCount * 6 - 5) = 4 + Val("" & !�ڿ�)
                        End If
                    End If
                    
                    
                    If Val("" & !ʧ��1) = 0 And bln�ϲ��� = False Then
                        aryY(intRow, 8 + lngCount * 6 - 4) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    Else
                        '��������������ֵʱ��Ϊ��������ֵ
                        If Abs(Val("" & !ʧ��1)) > 4 Then
                            aryY(intRow, 8 + lngCount * 6 - 4) = 4 + IIf(Val("" & !ʧ��1) < -4, -4, 4)
                        ElseIf Val("" & !ʧ��1) <> 0 Then
                            aryY(intRow, 8 + lngCount * 6 - 4) = 4 + Val("" & !ʧ��1)
                        End If
                    End If
                    
                    If Val("" & !ʧ��2) = 0 And bln�ϲ��� = False Then
                        aryY(intRow, 8 + lngCount * 6 - 3) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    Else
                        '��������������ֵʱ��Ϊ��������ֵ
                        If Abs(Val("" & !ʧ��2)) > 4 Then
                            aryY(intRow, 8 + lngCount * 6 - 3) = 4 + IIf(Val("" & !ʧ��2) < -4, -4, 4)
                        ElseIf Val("" & !ʧ��2) <> 0 Then
                            aryY(intRow, 8 + lngCount * 6 - 3) = 4 + Val("" & !ʧ��2)
                        End If
                    End If
                    
                    If Val("" & !ʧ��3) = 0 And bln�ϲ��� = False Then
                        aryY(intRow, 8 + lngCount * 6 - 2) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    Else
                        '��������������ֵʱ��Ϊ��������ֵ
                        If Abs(Val("" & !ʧ��3)) > 4 Then
                            aryY(intRow, 8 + lngCount * 6 - 2) = 4 + IIf(Val("" & !ʧ��3) < -4, -4, 4)
                        ElseIf Val("" & !ʧ��3) <> 0 Then
                            aryY(intRow, 8 + lngCount * 6 - 2) = 4 + Val("" & !ʧ��3)
                        End If
                    End If
                    
                    If Val("" & !ʧ��4) = 0 And bln�ϲ��� = False Then
                        aryY(intRow, 8 + lngCount * 6 - 1) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    Else
                        '��������������ֵʱ��Ϊ��������ֵ
                        If Abs(Val("" & !ʧ��4)) > 4 Then
                            aryY(intRow, 8 + lngCount * 6 - 1) = 4 + IIf(Val("" & !ʧ��4) < -4, -4, 4)
                        ElseIf Val("" & !ʧ��4) <> 0 Then
                            aryY(intRow, 8 + lngCount * 6 - 1) = 4 + Val("" & !ʧ��4)
                        End If
                    End If
                    
                    If Val("" & !ʧ��5) = 0 And bln�ϲ��� = False Then
                        aryY(intRow, 8 + lngCount * 6) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    Else
                        '��������������ֵʱ��Ϊ��������ֵ
                        If Abs(Val("" & !ʧ��5)) > 4 Then
                            aryY(intRow, 8 + lngCount * 6) = 4 + IIf(Val("" & !ʧ��5) < -4, -4, 4)
                        ElseIf Val("" & !ʧ��5) <> 0 Then
                            aryY(intRow, 8 + lngCount * 6) = 4 + Val("" & !ʧ��5)
                        End If
                    End If

                    Exit For
                End If
            Next
            .MoveNext
        Loop
    End With
    '�������30�������,����30�������
    'intLoop = 11
    If intLoop <> 0 Then
        For intLoop = intLoop + 1 To 31
            dateLoop = DateAdd("d", 1, dateLoop)
            If dateLoop <= CDate(mstrToDate) Then
                If Format(Nvl(dateLoop), "dd") <> "01" Then
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add intLoop, Format(Nvl(dateLoop), "dd")
                Else
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add intLoop, Format(Nvl(dateLoop), "mm" & "��")
                End If
            End If
            aryX(intLoop) = intLoop
            aryY(intLoop, 0) = 4
            aryY(intLoop, 1) = 4 + 1: aryY(intLoop, 2) = 4 - 1
            aryY(intLoop, 3) = 4 + 2: aryY(intLoop, 4) = 4 - 2
            aryY(intLoop, 5) = 4 + 3: aryY(intLoop, 6) = 4 - 3
            aryY(intLoop, 7) = 4 + 4: aryY(intLoop, 8) = 4 - 4
            
            For lngCount = 1 To Me.chk�ʿ�Ʒ.Count
                aryY(intLoop, 8 + lngCount * 6 - 5) = Me.chtThis.ChartGroups(1).Data.HoleValue
                aryY(intLoop, 8 + lngCount * 6 - 4) = Me.chtThis.ChartGroups(1).Data.HoleValue
                aryY(intLoop, 8 + lngCount * 6 - 3) = Me.chtThis.ChartGroups(1).Data.HoleValue
                aryY(intLoop, 8 + lngCount * 6 - 2) = Me.chtThis.ChartGroups(1).Data.HoleValue
                aryY(intLoop, 8 + lngCount * 6 - 1) = Me.chtThis.ChartGroups(1).Data.HoleValue
                aryY(intLoop, 8 + lngCount * 6) = Me.chtThis.ChartGroups(1).Data.HoleValue
            Next
        Next
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
            .Min = 4 - Val(dblMax)
            .Max = 4 + Val(dblMax)
        End With
        With .ChartArea.Axes("X")
            .Max = aryX(UBound(aryX))
        End With
        .IsBatched = False
        .AllowUserChanges = False
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


'--------------------------------------------
'����Ϊ�ؼ��¼�����
'--------------------------------------------
Private Sub chk�ʿ�Ʒ_Click(Index As Integer)
    If Me.Visible = False Then Exit Sub
    If Me.chk�ʿ�Ʒ(Index).Enabled = False Then Exit Sub
    With Me.chtThis.ChartGroups(1)
        If .Data.NumSeries < 9 + (Index + 1) * 2 - 1 Then Exit Sub
        If Me.chk�ʿ�Ʒ(Index).Value = vbChecked Then
            .Styles(9 + (Index + 1) * 6 - 5).Line.Pattern = oc2dLineSolid
            .Styles(9 + (Index + 1) * 6 - 5).Symbol.Size = 7
        Else
            .Styles(9 + (Index + 1) * 6 - 5).Line.Pattern = oc2dLineNone
            .Styles(9 + (Index + 1) * 6 - 5).Symbol.Size = 0
        End If
    End With
End Sub

Private Sub ChtThis_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim px As Long
    Dim py As Long
    Dim Series As Long
    Dim Point As Long
    Dim Distance As Long
    Dim Region As Long
    
    On Error Resume Next
    
    px = x / Screen.TwipsPerPixelX
    py = y / Screen.TwipsPerPixelY
    
    If (Button = 0) Then
        With chtThis
            Region = .ChartGroups(1).CoordToDataIndex(px, py, oc2dFocusXY, Series, Point, Distance)
            If (Series > 0 And Point > 0) And (Distance <= 5) Then
                If (Region = oc2dRegionInChartArea) Then
                    .ToolTipText = (Val(.ChartGroups(1).Data(Series, Point)) - 4)  '* mdblSD + mdblAVG
                    If mintFormatNum > 0 Then
                        .ToolTipText = Format(.ToolTipText, "###0." & Replace(Space(mintFormatNum), " ", "#"))
                    End If
                End If
            Else
'                .ToolTipText = ""
'                .Footer.Text = ""
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
        .ChartLabels(1).AttachCoord.y = .Header.Location.Top + .Header.Location.Height - 30
        '��2
        .ChartLabels(2).AttachCoord.x = .Header.Location.Left + (.ChartLabels(2).Location.Width / 2) - 150
        .ChartLabels(2).AttachCoord.y = .ChartLabels(1).Location.Top + .ChartLabels(1).Location.Height + 10
        '��3
        .ChartLabels(3).AttachCoord.x = .Header.Location.Left + (.ChartLabels(3).Location.Width / 2) - 150
        .ChartLabels(3).AttachCoord.y = .ChartLabels(2).Location.Top + .ChartLabels(2).Location.Height + 10
        
    End With
    
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With Me.chtThis
        .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
        .Top = Me.ScaleTop: .Height = Me.ScaleHeight - Me.chk�ʿ�Ʒ(0).Height - Screen.TwipsPerPixelY * 4
    End With
    
    With Me.chk�ʿ�Ʒ(0)
        .Left = Me.ScaleLeft + Screen.TwipsPerPixelX * 2
        .Top = Me.ScaleHeight - .Height - Screen.TwipsPerPixelY * 2
    End With
    For lngCount = 1 To Me.chk�ʿ�Ʒ.Count
        With Me.chk�ʿ�Ʒ(lngCount)
            .Left = Me.chk�ʿ�Ʒ(lngCount - 1).Left + Me.chk�ʿ�Ʒ(lngCount - 1).Width + Screen.TwipsPerPixelX * 10
            .Top = Me.chk�ʿ�Ʒ(lngCount - 1).Top
        End With
    Next
End Sub






