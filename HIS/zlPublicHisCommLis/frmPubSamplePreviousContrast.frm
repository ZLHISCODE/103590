VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Begin VB.Form frmPubSamplePreviousContrast 
   BorderStyle     =   0  'None
   Caption         =   "���αȶ�"
   ClientHeight    =   6165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox PICContrast 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5565
      Left            =   0
      ScaleHeight     =   5565
      ScaleWidth      =   7005
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7005
      Begin VB.PictureBox PicContrast_Bottom 
         Appearance      =   0  'Flat
         BackColor       =   &H00FCDBD8&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1635
         Left            =   0
         ScaleHeight     =   1635
         ScaleWidth      =   5280
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1110
         Width           =   5280
         Begin VB.OptionButton optContrast 
            Appearance      =   0  'Flat
            BackColor       =   &H00FCDBD8&
            Caption         =   "���ֵ(&2)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   2745
            TabIndex        =   8
            Top             =   60
            Width           =   1500
         End
         Begin VB.OptionButton optContrast 
            Appearance      =   0  'Flat
            BackColor       =   &H00FCDBD8&
            Caption         =   "������(&1)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   1170
            TabIndex        =   7
            Top             =   38
            Value           =   -1  'True
            Width           =   1395
         End
         Begin C1Chart2D8.Chart2D chtContrast 
            Height          =   975
            Left            =   60
            TabIndex        =   9
            Top             =   480
            Width           =   1005
            _Version        =   524288
            _Revision       =   7
            _ExtentX        =   1773
            _ExtentY        =   1720
            _StockProps     =   0
            ControlProperties=   "frmPubSamplePreviousContrast.frx":0000
         End
         Begin VB.Label lblCht 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ͼ������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   30
            TabIndex        =   10
            Top             =   60
            Width           =   960
         End
      End
      Begin VB.PictureBox PicContrast_Top 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   30
         ScaleHeight     =   975
         ScaleWidth      =   5970
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Width           =   5970
         Begin VB.TextBox txtMaxDay 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1650
            TabIndex        =   2
            Text            =   "30"
            Top             =   60
            Width           =   705
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFContrast 
            Height          =   1335
            Left            =   60
            TabIndex        =   3
            Top             =   480
            Width           =   2265
            _cx             =   3995
            _cy             =   2355
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16777215
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483635
            FloodColor      =   192
            SheetBorder     =   16777215
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   2
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   350
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
            Editable        =   2
            ShowComboButton =   0
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
         Begin VB.Label lblContrast 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ˢ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00F56C58&
            Height          =   240
            Left            =   2790
            MouseIcon       =   "frmPubSamplePreviousContrast.frx":0595
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   90
            Width           =   480
         End
         Begin VB.Label lblMaxDay 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����������:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   90
            TabIndex        =   4
            Top             =   90
            Width           =   1560
         End
      End
   End
End
Attribute VB_Name = "frmPubSamplePreviousContrast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-07-22
'ģ�鹦��:�걾������αȶ�
'---------------------------------------------------------------------------------------

Option Explicit

Private mlngSampleID As Long        '�걾ID
Private mdteS As Date               '����ʱ��
Private mintVersion As Integer      '�汾


'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-07-22
'��    ��:  ��������
'��    ��:
'           cnOracle        ���Ӷ���
'           �걾ID
'           dteS            �걾��������
'           intVersion      �汾25=�°棬10=�ϰ�
'��    ��:
'��    ��:
'����Ӱ��:
'����ע��:
'---------------------------------------------------------------------------------------
Public Sub InitData(ByVal lngSampleID As Long, ByVal dteS As Date, Optional ByVal intVersion As Integer = 25)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lngPaitID As Long
    
    mlngSampleID = lngSampleID
    mdteS = dteS
    mintVersion = intVersion
    
    If Val(txtMaxDay) > 365 Then
        If MsgBox("¼�����������������һ�꣬�Ƿ�����鿴�������ݣ�����������ܻᵼ�¼������������", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
            Exit Sub
        End If
    End If
    
    Me.VSFContrast.Rows = 1
    Me.VSFContrast.Rows = 2
    
    If intVersion = 25 Then
        strSQL = "select ����ID from ���鱨���¼ where ID=[1]"
        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "����ID", lngSampleID)
    Else
        strSQL = "select ����ID from ����걾��¼ where ID=[1]"
        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "����ID", lngSampleID)
    End If
    
    If Not rsTmp.EOF Then lngPaitID = Val(rsTmp("����ID") & "")
    If lngPaitID = 0 Then Exit Sub
    Call LoadContrastDBWriteVSF(Me.VSFContrast, lngSampleID, lngPaitID, dteS, Val(txtMaxDay.Text), intVersion)
End Sub

Public Function LoadContrastDBWriteVSF(vsfList As VSFlexGrid, lngSampleID As Long, lngPatientID As Long, SampleReportDate As Date, _
                                       intMaxDay As Integer, ByVal intVersion As Integer) As Boolean
      '����                   �����ݿ��ж����ȶ�����д��VSF��
          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String
          Dim intCol As Integer
          Dim dblTmp As Double
          Dim strData As String
          Dim blnTre As Boolean       '�Ƿ�����������걾


1         On Error GoTo LoadContrastDBWriteVSF_Error

2         blnTre = gobjLiscomlib.IsTre(lngSampleID)

3         If intVersion = 25 Then
4             If blnTre Then
                  '��������
5                 strSQL = "Select b.id, b.������, b.Ӣ����, b.��λ, a.id ����, c.����ʱ��, a.������, e.����ʱ��, b.���챨����, b.�������, a.�����־" & vbCrLf & _
                         "   From ���鱨����ϸ A, ����ָ�� B, ���鱨���¼ C, ��������걾 D, ��������ʱ�䷽�� E" & vbCrLf & _
                         "   Where a.��ĿID = b.id And a.�걾id = c.id And a.id = d.������ϸid And d.���ܷ���id = e.id And a.�걾ID =[1] order by a.id desc"
6                 Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "����ȶ�����", lngSampleID)
7             Else

8                 strSQL = "Select " & vbNewLine & _
                         " B.Id, B.������, B.Ӣ����, B.��λ, A.����, A.����ʱ��, A.������, A.����ʱ��, B.���챨����, B.�������, A.�����־" & vbNewLine & _
                           "From (Select B.��Ŀid ������Ŀid, B.����, B.����ʱ��, B.������, B.����ʱ��, B.�����־" & vbNewLine & _
                         "       From (Select A.Id ����, A.����id, A.����, A.�Ա�, A.�걾����, A.����ʱ��, B.��Ŀid, A.����ʱ��, B.�����־, a.��ҳid" & vbNewLine & _
                         "              From ���鱨���¼ A, ���鱨����ϸ B" & vbNewLine & _
                         "              Where A.Id = B.�걾id And A.Id = [1] ) A," & vbNewLine & _
                         "            (Select A.Id ����, A.����id, A.����, A.�Ա�, A.�걾����, A.����ʱ��, B.��Ŀid, B.������, A.����ʱ��, B.�����־, a.��ҳid" & vbNewLine & _
                         "              From ���鱨���¼ A, ���鱨����ϸ B" & vbNewLine & _
                         "              Where A.Id = B.�걾id And A.����id = [2] And" & vbNewLine & _
                         "                    ����ʱ�� Between [3] And" & vbNewLine & _
                         "                    [4] And A.Id <= [1] ) B," & vbNewLine & _
                         "            (Select A.Id ����" & vbNewLine & _
                         "              From ���鱨���¼ A, ���鱨����ϸ B" & vbNewLine & _
                         "              Where A.Id = B.�걾id And A.����id = [2] And" & vbNewLine & _
                         "                    ����ʱ��+0 Between [3] And" & vbNewLine & _
                         "                    [4] And A.Id <= [1] " & vbNewLine & _
                         "              Group By A.Id" & vbNewLine & _
                         "              Having Count(A.Id) > 0) C" & vbNewLine & _
                         "       Where A.����id = B.����id And A.��Ŀid + 0 = B.��Ŀid And Nvl(A.�걾����, 0) = Nvl(B.�걾����, 0) And A.���� = B.���� And A.�Ա� = B.�Ա� And a.��ҳID = b.��ҳID And" & vbNewLine & _
                         "             B.���� = C.����) A, ����ָ�� B" & vbNewLine & _
                           "Where A.������Ŀid = B.Id" & vbNewLine & _
                           "Order By  A.���� Desc ,LPad(B.�������, 10, '0'), B.Id"

9                 Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "����ȶ�����", lngSampleID, lngPatientID, CDate(Format(SampleReportDate - intMaxDay, "yyyy-MM-dd") & " 00:00:00"), _
                                            CDate(Format(SampleReportDate, "yyyy-MM-dd") & " 23:59:59"))
10            End If
11        Else
12            strSQL = "    Select " & vbNewLine & _
                     "       i.Id, i.���� As ������, v.��д As Ӣ����, i.���㵥λ As ��λ, a.����, a.����ʱ��, a.������, v.���챨����, v.�������" & vbNewLine & _
                     "       From (Select b.������Ŀid, b.����, b.����ʱ��, b.������" & vbNewLine & _
                     "              From (Select a.Id ����, a.����id, a.�걾����, a.���ʱ�� ����ʱ��, b.������Ŀid, b.������" & vbNewLine & _
                     "                     From ����걾��¼ A, ������ͨ��� B" & vbNewLine & _
                     "                     Where a.Id = b.����걾id And a.Id = [1] And ����id = [2] And b.������ Is Not Null) A," & vbNewLine & _
                     "                   (Select a.Id ����, a.����id, a.�걾����, a.���ʱ�� ����ʱ��, b.������Ŀid, b.������" & vbNewLine & _
                     "                     From ����걾��¼ A, ������ͨ��� B" & vbNewLine & _
                     "                     Where a.Id = b.����걾id And a.Id < [1] And ����id = [2]  And  ���ʱ�� Between [3] And [4]  And b.������ Is Not Null) B" & vbNewLine & _
                     "              Where a.����id = b.����id And a.������Ŀid + 0 = b.������Ŀid) A, ������Ŀ V, ���鱨����Ŀ R, ������ĿĿ¼ I" & vbNewLine & _
                     "       Where A.������Ŀid = v.������Ŀid And A.������Ŀid = r.������Ŀid And r.������Ŀid = i.ID And i.�����Ŀ <> 1"
13            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "����ȶ�����", lngSampleID, lngPatientID, CDate(Format(SampleReportDate - intMaxDay, "yyyy-MM-dd") & " 00:00:00"), _
                                     CDate(Format(SampleReportDate, "yyyy-MM-dd") & " 23:59:59"))
14        End If

15        gobjLiscomlib.vfgSetting 0, vsfList
16        With vsfList

17            .Rows = 1
18            .Cols = 2
19            .FixedRows = 1
              '        .FixedCols = 1
20            .TextMatrix(0, 0) = "������Ŀ": .ColWidth(0) = 2500
21            .TextMatrix(0, 1) = "��Ŀid": .ColWidth(1) = 0: .ColHidden(1) = True
              Dim strTimes As String
              Dim i As Long
              Dim strBT As String
              Dim strBH As String
              Dim J As Long
22            Do Until rsTmp.EOF
                  '���մ��������������ν������ʼ��д�뱾�ν��
                  '�ʼд���ʱ�� ��û�д�����ͨ�����ν��д���ϴν��
23                If strTimes = "" Or strTimes = rsTmp("����") & "" Then
24                    If strBT <> "**" Then
25                        strBH = "**"
26                        .Rows = .Rows + 1
27                        intCol = 1
28                        If .Cols - 1 < intCol Then
29                            .Cols = .Cols + 1
30                            .ColWidth(intCol) = 1500
31                        End If

32                        If intCol = 1 Then
                              'д����Ŀ
33                            .TextMatrix(.Rows - 1, 0) = rsTmp("������") & "(" & rsTmp("Ӣ����") & ")"
34                            .TextMatrix(.Rows - 1, 1) = rsTmp("id")
35                        End If
36                        intCol = intCol + 1
37                        If .Cols - 1 < intCol Then
38                            .Cols = .Cols + 1
39                            .ColWidth(intCol) = 1200: .ColAlignment(intCol) = flexAlignLeftCenter
40                            If blnTre Then
41                                .TextMatrix(0, intCol) = rsTmp("����ʱ��") & ""
42                            Else
43                                .TextMatrix(0, intCol) = "����"
44                            End If
45                        End If
                          'д������
46                        .TextMatrix(.Rows - 1, intCol) = rsTmp("������") & "(����)"
47                        .Cell(flexcpBackColor, .Rows - 1, intCol) = GetValColour(Val(rsTmp("�����־") & ""))
48                    End If
49                End If
                  '��һ�εĽ��д����һ��
50                If strTimes <> "" And strBH <> "**" Then
51                    If strTimes <> rsTmp("����") & "" Then
52                        If strBT = "" Then strBT = "**"
53                        For i = 1 To .Rows - 1
54                            If .TextMatrix(i, 1) = rsTmp("id") Then
55                                .Cols = .Cols + 1
56                                If blnTre Then
57                                    strData = rsTmp("����ʱ��") & ""
58                                    .ColWidth(.Cols - 1) = 2000: .ColAlignment(.Cols - 1) = flexAlignLeftCenter: .TextMatrix(0, .Cols - 1) = strData
59                                Else
60                                    strData = Format(rsTmp("����ʱ��") & "", "yyyy-MM-dd HH:mm")

61                                    .ColWidth(.Cols - 1) = 2000: .ColAlignment(.Cols - 1) = flexAlignLeftCenter: .TextMatrix(0, .Cols - 1) = "��" & .Cols - 3 & "��" & "(" & strData & ")"
62                                End If
63                                dblTmp = Val(CalcVolatility(.TextMatrix(i, 1), .TextMatrix(i, .Cols - 1)))
64                                If dblTmp <> 0 And Val(rsTmp("���챨����") & "") <> 0 Then
65                                    If dblTmp > Val(rsTmp("���챨����") & "") Then
66                                        .Cell(flexcpBackColor, i, .Cols - 1) = RGB(248, 194, 169)
67                                    End If
68                                End If
                                  'д������
69                                .TextMatrix(i, .Cols - 1) = rsTmp("������") & "(����)"
70                                .Cell(flexcpBackColor, i, .Cols - 1) = GetValColour(Val(rsTmp("�����־") & ""))
71                            End If
72                        Next
73                    Else
74                        For i = 1 To .Rows - 1
75                            If .TextMatrix(i, 1) = rsTmp("id") Then
                                  'д������
76                                .TextMatrix(i, .Cols - 1) = rsTmp("������") & "(����)"
77                                .Cell(flexcpBackColor, i, .Cols - 1) = GetValColour(Val(rsTmp("�����־") & ""))
78                            End If
79                        Next
80                    End If
81                End If
82                strTimes = rsTmp("����") & ""
83                strBH = ""
84                rsTmp.MoveNext
85            Loop
86            For i = 2 To .Cols - 1
87                For J = 1 To .Rows - 1
88                    If .TextMatrix(J, i) <> "" Then
89                        .TextMatrix(J, i) = Replace(.TextMatrix(J, i), "(����)", "")
90                        If .TextMatrix(J, i) = "" Then
91                            .TextMatrix(J, i) = "�޽��"
92                        End If
93                    Else
94                        .TextMatrix(J, i) = "δ��"
95                    End If
96                Next
97            Next

98            If .Rows > 1 Then
99                .Row = 1
100               Call VSFContrast_SelChange
101           End If

102       End With



103       Exit Function
LoadContrastDBWriteVSF_Error:
104       Call WriteErrLog("zl9LisInsideComm", "frmPubSamplePreviousContrast", "ִ��(LoadContrastDBWriteVSF)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
105       Err.Clear

End Function

Private Sub Form_Resize()
    On Error Resume Next
    With PICContrast
        .Left = 0
        .Top = 0
        .Width = Me.Width
        .Height = Me.Height
    End With
End Sub

Private Sub lblContrast_Click()
    Call InitData(mlngSampleID, mdteS, mintVersion)
End Sub

Private Sub PicContrast_Bottom_Resize()
    On Error Resume Next
    With Me.chtContrast
        .Top = lblCht.Top + lblCht.Height + 75
        .Left = 0
        .Width = Me.PicContrast_Bottom.ScaleWidth
        .Height = Me.PicContrast_Bottom.ScaleHeight
    End With
End Sub

Private Sub PICContrast_Resize()
    On Error Resume Next
    With Me.PicContrast_Top
        .Top = 0
        .Left = 0
        .Width = Me.PICContrast.ScaleWidth
        .Height = Me.PICContrast.ScaleHeight / 2
    End With
    With Me.PicContrast_Bottom
        .Top = PicContrast_Top.Top + PicContrast_Top.Height + 25
        .Left = 0
        .Width = Me.PicContrast_Top.Width
        .Height = Me.PICContrast.ScaleHeight - .Top
    End With
End Sub

Private Sub PicContrast_Top_Resize()
    On Error Resume Next
    With Me.VSFContrast
        .Top = Me.lblMaxDay.Top + lblMaxDay.Height + 50
        .Left = 0
        .Width = PicContrast_Top.ScaleWidth
        .Height = PicContrast_Top.ScaleHeight - .Top
    End With
End Sub

Private Sub txtMaxDay_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) Then Exit Sub
End Sub

Private Sub VSFContrast_SelChange()
    Dim strErr As String
    Dim intType As Integer
    With Me.VSFContrast
        If .Cols < 1 Then Exit Sub
        If .Row > 0 Then
            If Me.optContrast(0).value = True Then
                intType = 1
            Else
                intType = 2
            End If
            If LoadVSFContrastToCht(Me.VSFContrast, Me.chtContrast, .Row, intType, strErr) = False Then
                MsgBox strErr, vbInformation, gSysInfo.AppName
            End If
        End If
    End With
End Sub

Public Function LoadVSFContrastToCht(vsfList As VSFlexGrid, chtObj As Chart2D, intRow As Integer, intType As Integer, strErr As String) As Boolean
          '����           ��VSF��������д��Cht�ؼ�
          Dim intCol As Integer
          Dim dblMax As Double
          Dim i As Integer

1         On Error GoTo LoadVSFContrastToCht_Error

2         chtObj.ChartGroups(1).Data.NumSeries = 0
3         With chtObj.ChartGroups(1)
4             .ChartType = oc2dTypePlot  '����
5             .Styles(oc2dTypePlot).Symbol.Shape = oc2dShapeBox
6             With .Data
7                 .Layout = oc2dDataArray
8                 .NumSeries = 1
9                 .NumPoints(1) = vsfList.Cols - 1
10            End With
11        End With

12        With chtObj.ChartArea
13            .Axes("X").MajorGrid.Spacing.IsDefault = True
14            .Axes("Y").MajorGrid.Spacing.IsDefault = True
15            .Axes("X").AnnotationMethod = oc2dAnnotateValueLabels   '��������ʾֵ��ʾ

16        End With
17        With chtObj.ChartGroups(1).Data
18            For intCol = 2 To vsfList.Cols - 1
19                If IsNumeric(vsfList.TextMatrix(vsfList.Row, intCol)) = True Then
20                    i = i + 1
21                    Select Case intType
                          Case 1
22                            If intCol = 2 Then
23                                If vsfList.TextMatrix(vsfList.Row, 2) <> "" Then
24                                    .Y(1, i) = 0
25                                End If
26                            Else
27                                If IsNumeric(vsfList.TextMatrix(vsfList.Row, 2)) = True And IsNumeric(vsfList.TextMatrix(vsfList.Row, intCol)) = True Then
28                                    If CalcVolatility(vsfList.TextMatrix(vsfList.Row, 2), vsfList.TextMatrix(vsfList.Row, intCol)) <> "" Then
29                                        .Y(1, i) = Val(CalcVolatility(vsfList.TextMatrix(vsfList.Row, 2), vsfList.TextMatrix(vsfList.Row, intCol)))
30                                    Else
31                                        .Y(1, i) = 1E+308
32                                    End If
33                                End If
34                            End If
35                        Case 2
36                            If IsNumeric(vsfList.TextMatrix(vsfList.Row, intCol)) = True Then
37                                .Y(1, i) = IIf(vsfList.TextMatrix(vsfList.Row, intCol) = "", 1E+308, vsfList.TextMatrix(vsfList.Row, intCol))
38                            End If
39                    End Select
40                    If Abs(.Y(1, i)) > Abs(dblMax) And .Y(1, i) <> 1E+308 Then
41                        dblMax = .Y(1, i)
42                    End If
43                End If
44            Next
45        End With

46        With chtObj.ChartArea
47            Select Case intType
                  Case 1              '������
48                    .Axes("Y").DataMax = Abs(dblMax)
49                    .Axes("Y").DataMin = Abs(dblMax) * -1
50                    .Axes("Y").Origin = 0
51                Case 2              '���ֵ
52                    .Axes("Y").DataMax = Abs(dblMax) + Abs(dblMax) / 100 * 10
53                    .Axes("Y").DataMin = 0
54                    .Axes("Y").Origin = 0
55            End Select
56        End With
57        LoadVSFContrastToCht = True



58        Exit Function
LoadVSFContrastToCht_Error:
59        Call WriteErrLog("zl9LisInsideComm", "frmPubSamplePreviousContrast", "ִ��(LoadVSFContrastToCht)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
60        Err.Clear

End Function



Public Function CalcVolatility(strCalcA As String, strCalcB As String) As String
    '���������

    On Error Resume Next

    If strCalcA = "" Or strCalcB = "" Then
        CalcVolatility = ""
        Exit Function
    End If
    If Val(strCalcA) = 0 Or Val(strCalcB) = 0 Then
        CalcVolatility = ""
    End If

    '����
    CalcVolatility = (Val(strCalcB) - Val(strCalcA)) / Val(strCalcA) * 100
End Function
