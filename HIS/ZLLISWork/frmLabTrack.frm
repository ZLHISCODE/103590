VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~3.OCX"
Begin VB.Form frmLabTrack 
   BorderStyle     =   0  'None
   Caption         =   "��ʷ����"
   ClientHeight    =   5535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picChart 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   2565
      Left            =   165
      ScaleHeight     =   2565
      ScaleWidth      =   8550
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2880
      Width           =   8550
      Begin VB.OptionButton opt���� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "������(&3)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   3660
         TabIndex        =   14
         Top             =   45
         Width           =   1260
      End
      Begin VB.OptionButton opt���� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "������(&1)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   915
         TabIndex        =   12
         Top             =   45
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton opt���� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "���ֵ(&2)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   2250
         TabIndex        =   11
         Top             =   45
         Width           =   1260
      End
      Begin C1Chart2D8.Chart2D chtThis 
         Height          =   2085
         Left            =   30
         TabIndex        =   10
         Top             =   285
         Width           =   8520
         _Version        =   524288
         _Revision       =   7
         _ExtentX        =   15028
         _ExtentY        =   3678
         _StockProps     =   0
         ControlProperties=   "frmLabTrack.frx":0000
      End
      Begin VB.Label lbl��Ŀ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ŀ:RBC"
         Height          =   180
         Left            =   7665
         TabIndex        =   13
         Top             =   45
         Width           =   720
      End
      Begin VB.Label lblͼ������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ͼ������:"
         Height          =   180
         Left            =   90
         TabIndex        =   9
         Top             =   45
         Width           =   810
      End
   End
   Begin VB.PictureBox picData 
      BorderStyle     =   0  'None
      Height          =   2445
      Left            =   120
      ScaleHeight     =   2445
      ScaleWidth      =   8550
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   300
      Width           =   8550
      Begin VSFlex8Ctl.VSFlexGrid vfgData 
         Height          =   2085
         Left            =   0
         TabIndex        =   7
         Top             =   315
         Width           =   8565
         _cx             =   15108
         _cy             =   3678
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
         BackColorFixed  =   15790320
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
         Rows            =   3
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
      Begin VB.TextBox txt���� 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   5085
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "10"
         Top             =   60
         Width           =   525
      End
      Begin VB.TextBox txt���� 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   6945
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "3"
         Top             =   60
         Width           =   330
      End
      Begin VB.CheckBox chkHide 
         Appearance      =   0  'Flat
         Caption         =   "����������"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   15
         TabIndex        =   1
         Top             =   75
         Width           =   1275
      End
      Begin VB.CommandButton cmdRefersh 
         Caption         =   "ˢ��"
         Height          =   350
         Left            =   7500
         TabIndex        =   6
         Top             =   0
         Width           =   1320
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����������:"
         Height          =   180
         Left            =   3945
         TabIndex        =   5
         Top             =   75
         Width           =   1170
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ٴ���:"
         Height          =   180
         Left            =   5790
         TabIndex        =   4
         Top             =   75
         Width           =   1170
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmLabTrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    ������� = 0: ������: Ӣ����: ������: ��λ
End Enum

Private mlngRcdId As Long           '��ǰ��ʾ��������¼��id
Private mstrEndTime As String       '���μ���ʱ��
Private mintIdentMode As Integer    '��ʷ�Ƚϲ���ʶ��ʽ

Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
Dim lngCount As Long, lngRow As Long, lngCol As Long

Private Function zlGetCV(ParamArray dbInput() As Variant) As Double
    '���ܣ����ض����ֵ��CVֵ��ͳ�ƺ���(����ϵ��)
    '����������Ϊ��ֵ����
    Dim lngSubs As Long
    Dim dblSumAll As Double, dblSquSum As Double, dblSumSqu As Double
    Dim dblAV As Double, dblSD As Double
    
    If UBound(dbInput) < 1 Then zlGetCV = 0: Exit Function
    
    Err = 0: On Error GoTo 0
    dblSumAll = 0: dblSquSum = 0
    For lngSubs = LBound(dbInput) To UBound(dbInput)
        dblSumAll = dblSumAll + dbInput(lngSubs)
        dblSquSum = dblSquSum + dbInput(lngSubs) ^ 2
    Next
    If dblSumAll = 0 Then zlGetCV = 0: Exit Function
    dblSumSqu = dblSumAll ^ 2
    dblAV = dblSumAll / lngSubs
    dblSD = Sqr((dblSquSum - (dblSumSqu / lngSubs)) / (lngSubs - 1))
    zlGetCV = dblSD / dblAV * 100
    
End Function

Private Sub RefChart(Optional blnMust As Boolean)
    '���ܣ����ݵ�ǰ�Աȱ���ʾָ�����ݵı仯����
    '�������Ƿ�ǿ�����»�ȡ���ݽ���ˢ�£�������δ�仯ʱ��������ˢ�´���
    
    Dim aryX() As Variant, aryY() As Variant
    Dim intLoop As Integer, dblAvg As Double
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    Dim j As Integer
    If Val(Me.chtThis.Tag) <> Me.vfgData.Row Or blnMust Then
        Me.chtThis.Tag = Me.vfgData.Row
    Else
        Exit Sub
    End If
    
    '��������������Ϊ0�����ͼ����ʾ
    Me.chtThis.ChartGroups(1).Data.NumSeries = 0
    
    If Me.vfgData.Row < Me.vfgData.FixedRows Then Me.lbl��Ŀ.Caption = "": Exit Sub
    
    '���ԺͶ�����Ŀ����ͼ
    If Me.vfgData.TextMatrix(Me.vfgData.Row, mCol.�������) = "2" Or _
       Me.vfgData.TextMatrix(Me.vfgData.Row, mCol.�������) = "3" Then
       Me.chtThis.IsBatched = False
       Exit Sub
    End If
    
    
    '����ͼ�εĻ�����̬
    With Me.chtThis.ChartGroups(1)
        .ChartType = oc2dTypePlot  '����
        .Styles(oc2dTypePlot).Symbol.Shape = oc2dShapeBox
        With .Data
            .LayOut = oc2dDataArray
            .NumSeries = 1
            .NumPoints(1) = 4
        End With
    End With
    With Me.chtThis.ChartArea
        .Axes("X").MajorGrid.Spacing.IsDefault = True
        .Axes("Y").MajorGrid.Spacing.IsDefault = True
        .Axes("X").AnnotationMethod = oc2dAnnotateValueLabels   '��������ʾֵ��ʾ
'        .Axes("X").AnnotationRotationAngle = 10
    End With
    
    If Me.opt����(0).Value = True Then
        Me.chtThis.ChartArea.Axes("Y").Title.Text = "������"
    ElseIf Me.opt����(1).Value = True Then
        Me.chtThis.ChartArea.Axes("Y").Title.Text = "���ֵ"
    Else
        Me.chtThis.ChartArea.Axes("Y").Title.Text = "������"
    End If
    
    '������֯
    Dim strMaxValue As String, strMinValue As String
    
    strMaxValue = 0
    If Me.opt����(0).Value = True Or Me.opt����(1).Value = True Then
        For intLoop = 0 To (Me.vfgData.Cols - mCol.��λ - 1) / 2 - 1
            If Val(vfgData.TextMatrix(vfgData.Row, mCol.��λ + 1 + intLoop * 2)) <> 0 Then
                j = j + 1
            End If
        Next
        If j = 0 Then j = 1
        ReDim aryX(j - 1)
        ReDim aryY(j - 1, 0)
    Else
        gstrSql = "Select ����, ������, Ӣ����, ������Ŀid, ������, decode(����,null,������,����) as ���� " & vbNewLine & _
                    "From (Select Decode(E.�������, Null, D.����, E.�������) As ����, D.������, D.Ӣ����, B.������Ŀid, B.������, H.����" & vbNewLine & _
                    "       From ����걾��¼ A, ������ͨ��� B, ����������Ŀ C, ����������Ŀ D, ������Ŀ E, ���鱨����Ŀ F, ������ĿĿ¼ G," & vbNewLine & _
                    "            (Select ������Ŀid, ���� As ���� From ������Ŀ���� Where ���� = 9 And ���� = 1) H" & vbNewLine & _
                    "       Where A.ID = B.����걾id And B.������Ŀid = C.��Ŀid And B.������Ŀid = D.ID And Nvl(C.��������Ŀ, 0) = -1 And A.ID = [1] And" & vbNewLine & _
                    "             B.������Ŀid = E.������Ŀid And B.������Ŀid = F.������Ŀid And F.������Ŀid = G.ID And Nvl(G.�����Ŀ, 0) = 0 And" & vbNewLine & _
                    "             G.ID = H.������Ŀid(+)" & vbNewLine & _
                    "       Union All" & vbNewLine & _
                    "       Select Decode(E.�������, Null, D.����, E.�������) As ����, D.������, D.Ӣ����, B.������Ŀid, B.������, H.����" & vbNewLine & _
                    "       From ����걾��¼ A, ������ͨ��� B, ����������Ŀ C, ����������Ŀ D, ������Ŀ E, ���鱨����Ŀ F, ������ĿĿ¼ G," & vbNewLine & _
                    "            (Select ������Ŀid, ���� As ���� From ������Ŀ���� Where ���� = 9 And ���� = 1) H" & vbNewLine & _
                    "       Where A.ID = B.����걾id And B.������Ŀid = C.��Ŀid And B.������Ŀid = D.ID And Nvl(C.��������Ŀ, 0) = -1 And A.�ϲ�id = [1] And" & vbNewLine & _
                    "             B.������Ŀid = E.������Ŀid And B.������Ŀid = F.������Ŀid And F.������Ŀid = G.ID And Nvl(G.�����Ŀ, 0) = 0 And" & vbNewLine & _
                    "             G.ID = H.������Ŀid(+))" & vbNewLine & _
                    "Order By ����"

        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngRcdId)
        If rsTmp.RecordCount = 0 Then Exit Sub
        ReDim aryX(rsTmp.RecordCount - 1)
        ReDim aryY(rsTmp.RecordCount - 1, 0)
    End If
    Me.chtThis.ChartArea.Axes("X").ValueLabels.RemoveAll
    With Me.vfgData
        Me.lbl��Ŀ.Caption = "��Ŀ:" & .TextMatrix(.Row, mCol.������) & " (" & .TextMatrix(.Row, mCol.Ӣ����) & ")"
        For lngCount = 0 To (Me.vfgData.Cols - mCol.��λ - 1) / 2 - 1
            If Val(.TextMatrix(.Row, mCol.��λ + 1 + lngCount * 2)) <> 0 Then
                aryX(i) = i

                If Me.opt����(0).Value = True Then
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add lngCount, .TextMatrix(0, mCol.��λ + 1 + lngCount * 2)
                    If Val(.TextMatrix(.Row, mCol.��λ + 2 + lngCount * 2)) = 0 And Val(.TextMatrix(.Row, mCol.��λ + 1 + lngCount * 2)) = 0 Then
    '                    aryY(lngCount, 0) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    Else
                        aryY(i, 0) = Val(.TextMatrix(.Row, mCol.��λ + 2 + lngCount * 2))
                    End If
                ElseIf Me.opt����(1).Value = True Then
                    Me.chtThis.ChartArea.Axes("X").ValueLabels.Add lngCount, .TextMatrix(0, mCol.��λ + 1 + lngCount * 2)
                    If Val(.TextMatrix(.Row, mCol.��λ + 1 + lngCount * 2)) = 0 Then
    '                    aryY(lngCount, 0) = Me.chtThis.ChartGroups(1).Data.HoleValue
                    Else
                        aryY(i, 0) = Val(.TextMatrix(.Row, mCol.��λ + 1 + lngCount * 2))
                    End If
                End If
                If Val(strMaxValue) < Abs(Val(aryY(i, 0))) Then
                    strMaxValue = Abs(Val(aryY(i, 0)))
                End If
                If Val(strMinValue) > Abs(Val(aryY(i, 0))) Then
                    strMinValue = Abs(Val(aryY(i, 0)))
                End If
                i = i + 1
            End If
        Next
    End With
    
    With Me.vfgData
        If Me.opt����(2).Value = True Then
            For lngCount = LBound(aryX) To UBound(aryX)
                aryX(lngCount) = lngCount
                Me.chtThis.ChartArea.Axes("X").ValueLabels.Add lngCount, Nvl(rsTmp("����"))
                aryY(lngCount, 0) = Val(Nvl(rsTmp("������")))
                rsTmp.MoveNext
                If Val(strMaxValue) < Abs(Val(aryY(lngCount, 0))) Then
                    strMaxValue = Abs(Val(aryY(lngCount, 0)))
                End If
                If Val(strMinValue) > Abs(Val(aryY(lngCount, 0))) Then
                    strMinValue = Abs(Val(aryY(lngCount, 0)))
                End If
            Next
        End If
    End With
    
    '���ˢ���ڲ�����
    Me.chtThis.IsBatched = True
    Me.chtThis.ChartGroups(1).Data.NumPoints(1) = UBound(aryX) + 1
    Call Me.chtThis.ChartGroups(1).Data.CopyXVectorIn(1, aryX)
    Call Me.chtThis.ChartGroups(1).Data.CopyYArrayIn(aryY)
    
    If opt����(0).Value = True Then
        Me.chtThis.ChartArea.Axes("Y").Origin = 0
        Me.chtThis.ChartArea.Axes("Y").Min = -1 * Val(strMaxValue)
        Me.chtThis.ChartArea.Axes("Y").Max = Val(strMaxValue)
    ElseIf opt����(1).Value = True Then
        On Error Resume Next
        For intLoop = 0 To UBound(aryY, 1) - 1
            dblAvg = dblAvg + Val(aryY(intLoop, 0))
        Next
        If dblAvg <> 0 Then
            dblAvg = dblAvg / UBound(aryY, 1)
            Me.chtThis.ChartArea.Axes("Y").Origin = dblAvg
            If (dblAvg - Val(strMinValue)) < (Val(strMaxValue) - dblAvg) Then
                Me.chtThis.ChartArea.Axes("Y").Min = Val(dblAvg - (Val(strMaxValue) - dblAvg))
                Me.chtThis.ChartArea.Axes("Y").Max = Val(dblAvg + (Val(strMaxValue) - dblAvg))
            Else
                Me.chtThis.ChartArea.Axes("Y").Min = Val(dblAvg - (dblAvg - Val(strMinValue)))
                Me.chtThis.ChartArea.Axes("Y").Max = Val(dblAvg + (dblAvg - Val(strMinValue)))
            End If
        End If
    Else
        Me.chtThis.ChartArea.Axes("Y").Origin = 0
        Me.chtThis.ChartArea.Axes("Y").Min = 0
        Me.chtThis.ChartArea.Axes("Y").Max = Val(strMaxValue)
    End If
    Me.chtThis.IsBatched = False

End Sub

Private Sub setListFormat(Optional blnKeepData As Boolean)
    '���ܣ���ʼ�����òο�ֵ�б�
    '������ blnKeepData-�Ƿ������ݣ���ֻ���������ø�ʽ
    With Me.vfgData
        .Redraw = flexRDNone
        If blnKeepData = False Then
            .Clear
            .Rows = 2: .FixedRows = 2: .Cols = mCol.��λ + 1: .FixedCols = .Cols
            For lngCol = 0 To mCol.��λ: .TextMatrix(0, lngCol) = "��Ŀ": Next
            .TextMatrix(1, mCol.�������) = "�������"
            .TextMatrix(1, mCol.������) = "������"
            .TextMatrix(1, mCol.Ӣ����) = "Ӣ����"
            .TextMatrix(1, mCol.������) = "������"
            .TextMatrix(1, mCol.��λ) = "��λ"
            .MergeCells = flexMergeFixedOnly
            .MergeRow(0) = True
            .ColWidth(mCol.�������) = 0
            .ColWidth(mCol.������) = 1500
            .ColWidth(mCol.Ӣ����) = 900
            .ColWidth(mCol.������) = 0
            .ColWidth(mCol.��λ) = 500
        End If
        If .Cols > mCol.��λ + 1 Then
            .TextMatrix(0, mCol.��λ + 1) = "���ν��"
            .TextMatrix(1, mCol.��λ + 1) = "���ν��"
            .MergeCol(mCol.��λ + 1) = True
            .ColWidth(mCol.��λ + 2) = 0
        End If
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            If .ColWidth(lngCount) = 0 Then .ColHidden(lngCount) = True
        Next
        .AutoSize mCol.������, .Cols - 1
        If .Cols > .FixedCols Then .Col = .FixedCols
        If .Rows > .FixedRows Then .Row = .FixedRows
        Call RefChart(True)
        .Redraw = flexRDDirect
    End With
End Sub

Public Function zlRefresh(lngRcdId As Long) As Boolean
    '���ܣ���������idˢ�µ�ǰ��ʾ����
    '��������ǰ��Ŀid
    Dim lngDates As Long, lngTimes As Long
    Dim strRows As String, aryRows() As String
    Dim strCols As String, aryCols() As String
    Dim dblCurCV As Double     '�����CV
    Dim strPatientName As String                    '��������
    Dim strPatinetSex As String                     '�����ձ�
    Dim lngPatientID As Long
    
    If lngRcdId = 0 Then Call setListFormat: zlRefresh = True: Exit Function
    Err = 0: On Error GoTo ErrHand
    
    '��õ�ǰ�����ʱ�䡢��ĿҪ��ĸ���������ȡ��Ŀ�����ģ�
    If mlngRcdId <> lngRcdId Then
        mlngRcdId = lngRcdId
        gstrSql = "Select Nvl(L.����ʱ��, Sysdate) As ����ʱ��, Nvl(Max(��������), 0) As ����" & vbNewLine & _
                "From ������Ŀѡ�� O, ���鱨����Ŀ X, ������ͨ��� R, ����걾��¼ L" & vbNewLine & _
                "Where O.������Ŀid(+) = X.������Ŀid And X.������Ŀid = R.������Ŀid And R.����걾id = L.ID And L.ID = [1]" & vbNewLine & _
                "Group By Nvl(L.����ʱ��, Sysdate)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngRcdId)
        If rsTemp.RecordCount > 0 Then
            Me.txt����.Text = rsTemp!����
            mstrEndTime = Format(rsTemp!����ʱ��, "yyyy-MM-dd hh:mm:ss")
        Else
            Me.txt����.Text = 30
            mstrEndTime = Format(Now(), "yyyy-MM-dd hh:mm:ss")
        End If
    End If
    If Val(Me.txt����.Text) <= 0 Then Me.txt����.Text = 30
    If Val(Me.txt����.Text) <= 0 Then Me.txt���� = 3
    
    lngDates = Val(Me.txt����.Text)
    lngTimes = Val(Me.txt����.Text)
    
'    If mintIdentMode <> 0 Then
        gstrSql = "select ����,����ID,�Ա� from ����걾��¼ where id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngRcdId)
        If rsTemp.EOF = False Then strPatientName = Nvl(rsTemp("����")): lngPatientID = Nvl(rsTemp("����ID"), 0): strPatinetSex = Nvl(rsTemp("�Ա�"))
'    End If
    
    '��ѯ��������װ�룺
    gstrSql = "Select /*+ RULE */ I.ID, I.���� As ������, V.��д As Ӣ����, I.���㵥λ As ��λ, L.����, L.����ʱ��, L.������, V.���챨����,V.������� " & vbNewLine & _
            "From (Select L.������Ŀid, L.����, L.����ʱ��, L.������ " & vbNewLine & _
            "       From (Select M.����id As ����id, M.����, M.�Ա�, L.ID As ����, L.����ʱ��, R.������Ŀid, R.������,L.�걾���� " & vbNewLine & _
            "              From ����걾��¼ L, ������ͨ��� R, ����ҽ����¼ M, " & _
            "                   (select ����id,����,�Ա� from ������Ϣ where " & IIf(mintIdentMode = 0, " ����ID = [4] ", " ���� = [5] and �Ա� = [6] ") & " ) N " & vbNewLine & _
            "              Where M.ID = L.ҽ��id And L.ID = R.����걾id And  " & vbNewLine & _
            "                    L.����ʱ�� Between [2]  And" & vbNewLine & _
            "                    [3] and L.����id = N.����id ) L," & vbNewLine & _
            "            (Select M.����id As ����id, M.����, M.�Ա�, L.����ʱ��, R.������Ŀid,L.�걾���� " & vbNewLine & _
            "              From ����ҽ����¼ M, ����걾��¼ L, ������ͨ��� R" & vbNewLine & _
            "              Where M.ID = L.ҽ��id And L.ID = R.����걾id And L.ID = [1]) C" & vbNewLine & _
            "        " & IIf(mintIdentMode = 0, "Where L.����id = C.����id   ", " Where  l.���� = c.���� And l.�Ա� = c.�Ա�  ") & _
            "        And L.������Ŀid+0 = C.������Ŀid And L.�걾���� = C.�걾����  ) L, ������Ŀ V, ���鱨����Ŀ R, ������ĿĿ¼ I" & vbNewLine & _
            "Where L.������Ŀid = V.������Ŀid And L.������Ŀid = R.������Ŀid And R.������Ŀid = I.ID And I.�����Ŀ <> 1" & vbNewLine & _
            "Order By I.����, L.����ʱ�� desc"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngRcdId, CDate(Format(mstrEndTime, "yyyy-MM-dd 00:00:00")) - lngDates, _
                                       CDate(Format(mstrEndTime, "yyyy-MM-dd HH:MM:SS")), lngPatientID, strPatientName, strPatinetSex)
    
    Err = 0: On Error GoTo 0
    strRows = "": strCols = ""
    With Me.vfgData
        .Redraw = flexRDNone
        .Rows = .FixedRows: .Cols = .FixedCols
        lngRow = 0: lngCol = 0
        Do While Not rsTemp.EOF
            If InStr(1, strRows & ",", "," & rsTemp!ID & ",") = 0 Then
                strRows = strRows & "," & rsTemp!ID
                .Rows = .Rows + 1: lngRow = .Rows - 1
                .RowData(lngRow) = CLng(rsTemp!ID)
            Else
                aryRows = Split(strRows, ",")
                For lngCount = LBound(aryRows) To UBound(aryRows)
                    If Val(aryRows(lngCount)) = rsTemp!ID Then lngRow = .FixedRows - 1 + lngCount: Exit For
                Next
            End If
            .TextMatrix(lngRow, mCol.�������) = "" & rsTemp!�������
            .TextMatrix(lngRow, mCol.������) = "" & rsTemp!������
            .TextMatrix(lngRow, mCol.Ӣ����) = "" & rsTemp!Ӣ����
            .TextMatrix(lngRow, mCol.������) = Val("" & rsTemp!���챨����)
            .TextMatrix(lngRow, mCol.��λ) = "" & rsTemp!��λ
            
            If InStr(1, strCols & ",", "," & rsTemp!���� & ",") = 0 Then
                If UBound(Split(strCols, ",")) < lngTimes + 1 Then
                    strCols = strCols & "," & rsTemp!����
                    .Cols = .Cols + 2: lngCol = .Cols - 1
                    .ColData(lngCol - 1) = CLng(rsTemp!����): .ColData(lngCol) = CLng(rsTemp!����)
                    .TextMatrix(0, lngCol - 1) = Format(rsTemp!����ʱ��, "yy-MM-dd HH:mm")
                    .TextMatrix(0, lngCol) = .TextMatrix(0, lngCol - 1)
                    .TextMatrix(1, lngCol - 1) = "���ֵ": .TextMatrix(1, lngCol) = "������"
                    .TextMatrix(lngRow, lngCol - 1) = "" & rsTemp!������
                End If
            Else
                aryCols = Split(strCols, ",")
                For lngCount = LBound(aryCols) To UBound(aryCols)
                    If Val(aryCols(lngCount)) = rsTemp!���� Then lngCol = .FixedCols - 1 + lngCount * 2: Exit For
                Next
                .TextMatrix(lngRow, lngCol - 1) = "" & rsTemp!������
            End If
        
            rsTemp.MoveNext
        Loop
        
        '�����ʼ�����д�ͱ���ɫ����
        For lngRow = .FixedRows To .Rows - 1
            .TextMatrix(lngRow, mCol.��λ + 1) = Replace(Replace(" " & Trim(.TextMatrix(lngRow, mCol.��λ + 1)), " .", "0."), " ", "")
            For lngCol = mCol.��λ + 4 To .Cols - 1 Step 2
                .TextMatrix(lngRow, lngCol - 1) = Replace(Replace(" " & Trim(.TextMatrix(lngRow, lngCol - 1)), " .", "0."), " ", "")
                If Val(.TextMatrix(lngRow, lngCol - 1)) = 0 Or Val(.TextMatrix(lngRow, mCol.��λ + 1)) = 0 Then
                    dblCurCV = 0
                Else
                    dblCurCV = (Val(.TextMatrix(lngRow, lngCol - 1)) - Val(.TextMatrix(lngRow, mCol.��λ + 1))) / Val(.TextMatrix(lngRow, mCol.��λ + 1)) * 100
                End If
                .TextMatrix(lngRow, lngCol) = Format(dblCurCV, "0.00;-0.00; ; ")
                If Val(.TextMatrix(lngRow, mCol.������)) <> 0 And Abs(dblCurCV) > Val(.TextMatrix(lngRow, mCol.������)) Then
                    .Cell(flexcpBackColor, lngRow, lngCol) = RGB(248, 194, 169)
                End If
            Next
        Next
        .Redraw = flexRDDirect
    End With
    Call setListFormat(True)
    
    zlRefresh = True: Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    zlRefresh = False
End Function

Private Sub chkHide_Click()
    Me.vfgData.ColHidden(mCol.������) = (Me.chkHide.Value = vbChecked)
    If Me.Visible Then Me.vfgData.SetFocus
End Sub

Private Sub chtThis_GotFocus()
    Me.dkpMan.RecalcLayout
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

Private Sub cmdRefersh_Click()
    Call Me.zlRefresh(mlngRcdId)
    Me.vfgData.SetFocus
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1: Item.Handle = Me.picData.hWnd
    Case 2: Item.Handle = Me.picChart.hWnd
    End Select
End Sub

Private Sub Form_Load()

    '��ñ��ز�������
    mintIdentMode = zlDatabase.GetPara("��ʷ����ʶ��", 100, 1208, 1)
    '����������
    If Val(zlDatabase.GetPara("����������", 100, 1208, 0)) = 0 Then
        Me.chkHide.Value = vbUnchecked
    Else
        Me.chkHide.Value = vbChecked
    End If
    Me.txt����.Text = 3

    '������ʽ����
    '------------------------------------------------------
    mlngRcdId = 0
    Me.chkHide.BackColor = Me.picData.BackColor
    Me.opt����(0).BackColor = Me.picChart.BackColor
    Me.opt����(1).BackColor = Me.picChart.BackColor
    Me.opt����(2).BackColor = Me.picChart.BackColor
    Call setListFormat

    '���񻮷�
    '-----------------------------------------------------
    Dim panThis As Pane
    Set panThis = dkpMan.CreatePane(1, 200, 400, DockTopOf, Nothing)
    panThis.Title = "��ʷ�Աȱ�"
    panThis.Options = PaneNoCaption
    Set panThis = dkpMan.CreatePane(2, 200, 300, DockBottomOf, Nothing)
    panThis.Title = "��ʷ�Ա�ͼ"
    panThis.Options = PaneNoCaption
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.chkHide.Value = vbUnchecked Then
        zlDatabase.SetPara "����������", 0, 100, 1208
    Else
        zlDatabase.SetPara "����������", 1, 100, 1208
    End If
    Me.dkpMan.DestroyAll
End Sub

Private Sub opt����_Click(Index As Integer)
    Call RefChart(True)
    Me.vfgData.SetFocus
End Sub

Private Sub picChart_Resize()
    Err = 0: On Error Resume Next
    Me.lbl��Ŀ.Left = Me.ScaleWidth - Me.lbl��Ŀ.Width - 90
    With Me.chtThis
        .Left = 0: .Width = Me.picChart.ScaleWidth
        .Height = Me.picChart.ScaleHeight - .Top
    End With
End Sub

Private Sub picData_Resize()
    Err = 0: On Error Resume Next
    With Me.cmdRefersh
        .Left = Me.picData.ScaleWidth - .Width + 15
    End With
    Me.txt����.Left = Me.cmdRefersh.Left - 900
    Me.lbl����.Left = Me.txt����.Left - Me.lbl����.Width
    Me.txt����.Left = Me.lbl����.Left - 900
    Me.lbl����.Left = Me.txt����.Left - Me.lbl����.Width
    Me.chkHide.Left = 45
    
    With Me.vfgData
        .Left = -15: .Width = Me.picData.ScaleWidth - .Left * 2
        .Height = Me.picData.ScaleHeight - .Top
    End With
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub vfgData_RowColChange()
    Call RefChart
End Sub
