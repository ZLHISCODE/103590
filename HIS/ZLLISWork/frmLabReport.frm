VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLabReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������ϸ��ӡ"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   10950
   Icon            =   "frmLabReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtEnd 
      Height          =   300
      Left            =   7470
      TabIndex        =   9
      Text            =   "8000"
      Top             =   120
      Width           =   945
   End
   Begin VB.TextBox txtStart 
      Height          =   300
      Left            =   6225
      TabIndex        =   8
      Text            =   "0"
      Top             =   120
      Width           =   945
   End
   Begin VB.ComboBox cbo���� 
      Height          =   300
      Left            =   690
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   120
      Width           =   2730
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��ӡ(&P)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   9705
      TabIndex        =   3
      Top             =   75
      Width           =   1100
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "ˢ��(&R)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8520
      TabIndex        =   2
      Top             =   90
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtp��ʼ���� 
      Height          =   315
      Left            =   4035
      TabIndex        =   1
      Top             =   120
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   208470019
      CurrentDate     =   39414
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPrint 
      Height          =   5805
      Left            =   150
      TabIndex        =   0
      Top             =   600
      Width           =   10665
      _cx             =   18812
      _cy             =   10239
      Appearance      =   1
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   3
      FixedCols       =   0
      RowHeightMin    =   280
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
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
   Begin VB.Label lbl�걾�� 
      Caption         =   "���             ��"
      Height          =   210
      Left            =   5715
      TabIndex        =   7
      Top             =   195
      Width           =   1740
   End
   Begin VB.Label lbl���� 
      Caption         =   "����"
      Height          =   195
      Left            =   210
      TabIndex        =   6
      Top             =   195
      Width           =   435
   End
   Begin VB.Label lbl���� 
      Caption         =   "����"
      Height          =   195
      Left            =   3525
      TabIndex        =   4
      Top             =   195
      Width           =   435
   End
End
Attribute VB_Name = "frmLabReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mLast����ID As Long

Private Sub cbo����_Click()
    If cbo����.ItemData(cbo����.ListIndex) <> mLast����ID Then
        Call cmdRefresh_Click
    End If
End Sub

Private Sub cmdPrint_Click()
    PrintSelection vsPrint, 0, 0, vsPrint.Rows - 1, vsPrint.Cols - 1
End Sub

Private Sub cmdRefresh_Click()
    LoadDataToVsf cbo����.ItemData(cbo����.ListIndex), Val(txtStart), Val(txtEnd), dtp��ʼ����.Value
    mLast����ID = cbo����.ItemData(cbo����.ListIndex)
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim rsTmp As adodb.Recordset
    
    With vsPrint
        .Rows = 4: .Cols = 8
        .MergeCellsFixed = flexMergeRestrictColumns
        
        .Cell(flexcpText, 0, 0, 0, .Cols - 1) = "ÿ�չ�������"
        .TextMatrix(0, 0) = "ÿ�չ�������"
        .MergeRow(0) = True
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        
        .Cell(flexcpText, 1, 0, 1, 6) = "���ڣ�" & Format(dtp��ʼ����.Value, "yyyy-MM-dd")
        .Cell(flexcpText, 1, 7, 1, 7) = "������" & cbo����.List(cbo����.ListIndex)
        .MergeRow(1) = True
        .Cell(flexcpAlignment, 1, 0, 1, .Cols - 1) = flexAlignLeftCenter
        .Cell(flexcpAlignment, 1, .Cols - 1, 1, .Cols - 1) = flexAlignRightCenter
        
        .TextMatrix(2, 0) = "���": .TextMatrix(2, 1) = "��Ա��": .TextMatrix(2, 2) = "����"
        .TextMatrix(2, 3) = "�Ա�": .TextMatrix(2, 4) = "����": .TextMatrix(2, 5) = "����"
        .TextMatrix(2, 6) = "����": .TextMatrix(2, 7) = "���"
        .Cell(flexcpAlignment, 2, 0, 2, .Cols - 1) = flexAlignCenterCenter
        
        .Cell(flexcpBackColor, 0, 0, 2, .Cols - 1) = vbWhite
        .GridLines = flexGridNone
        .GridLinesFixed = flexGridNone
    End With
    
    dtp��ʼ����.Value = Format(Now, "yyyy-MM-dd")
    
    strSQL = "Select ID,����,���� From �������� order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Do Until rsTmp.EOF
        cbo����.AddItem rsTmp.Fields("����") & "-" & rsTmp.Fields("����")
        cbo����.ItemData(cbo����.NewIndex) = rsTmp.Fields("ID")
        rsTmp.MoveNext
    Loop
    cmdRefresh.Enabled = False
    cmdPrint.Enabled = False
    
    If cbo����.ListCount > 0 Then
        cbo����.ListIndex = 0
        cmdRefresh.Enabled = True
    End If
    
End Sub

Private Sub PrintSelection(fg As VSFlexGrid, Row1&, Col1&, Row2&, Col2&)

'��ӡVSFlexGrid������
    Dim hl%, tr&, lc&, rd%

    hl = fg.HighLight: tr = fg.TopRow: lc = fg.LeftCol: rd = fg.Redraw

    fg.HighLight = 0
    
    fg.Redraw = flexRDNone
    ' hide non-selected rows and columns
    Dim i&

    For i = fg.FixedRows To fg.Rows - 1
       If i < Row1 Or i > Row2 Then fg.RowHidden(i) = True
    Next

    For i = fg.FixedCols To fg.Cols - 1
      If i < Col1 Or i > Col2 Then fg.ColHidden(i) = True
    Next
    

    ' scroll to top left corner
    fg.TopRow = fg.FixedRows
    fg.LeftCol = fg.FixedCols

    ' print visible area
    fg.PrintGrid , True, , 50, 200

    ' restore control
    fg.RowHidden(-1) = False

    fg.ColHidden(-1) = False

    fg.TopRow = tr: fg.LeftCol = lc: fg.HighLight = hl

    fg.Redraw = rd

  End Sub


Private Sub LoadDataToVsf(ByVal lng����ID As Long, lngStar��� As Long, lngEnd��� As Long, dateStart As Date)
    'װ�����ӡ����
        '<EhHeader>
        On Error GoTo LoadDataToVsf_Err
        '</EhHeader>
          Dim strSQL As String
          Dim rsTmp As adodb.Recordset
          Dim dateEnd As Date
          Dim strLastRow As String
          Dim lngCount As Long
          Dim iCol As Integer
          Dim lngRow As Long
          Dim str��Ϣ�� As String
          Dim strС�� As String
          Dim strEmergency As String
      
    '      lng����ID = 162
100       dateStart = Format(dateStart, "yyyy-MM-dd")
102       dateEnd = Format(dateStart, "yyyy-MM-dd 23:59:59")
    '      lngStar��� = 1: lngEnd��� = 8000
104       cmdPrint.Enabled = False
                
106     strSQL = "Select A.�걾���, Decode(A.������Դ, 2, A.סԺ��, A.�����) As ��Ա��, A.����, A.�Ա�, A.����, D.���� As ����, A.����, B.�������," & vbNewLine & _
                "                            RPad(C.��д, 6, ' ') || ' ' ||" & vbNewLine & _
                "                             RPad(Decode(������, '-', '����', '+', '����', '*', '*.**'," & vbNewLine & _
                "                                         Replace(������, '(+)', LPad(' ', 7 - Length(������)) || '(+)')) ||" & vbNewLine & _
                "                                  Decode(B.�����־, 3, '��', 2, '��', 1, '', 4, '�쳣', 5, '����', 6, '����', ''), 10, ' ') As ��Ϣ��" & vbNewLine & _
                "                     From ����ҽ����¼ E, ���ű� D, ������Ŀ C, ������ͨ��� B, ����걾��¼ A" & vbNewLine & _
                "                     Where A.ҽ��id = E.ID And E.��������id = D.ID And B.������Ŀid = C.������Ŀid And A.ID = B.����걾id And" & vbNewLine & _
                "                           To_Number(A.�걾���) Between [1] And [2] And B.������ Is Not Null And A.������ = B.��¼���� And" & vbNewLine & _
                "                           Nvl(A.����id, 0) = [3] And A.����ʱ�� Between [4] And" & vbNewLine & _
                "                           [5] " & vbNewLine & _
                "                     Order By LPad(�걾���, 4, '0'), B.�������"
108     strSQL = "Select A.�걾���, Decode(A.������Դ, 2, A.סԺ��, A.�����) As ��Ա��, A.����, A.�Ա�, A.����, D.���� As ����, A.����, B.�������," & vbNewLine & _
                "       RPad(C.��д, 6, ' ') As ��д, ������, �������, F.С��λ��," & vbNewLine & _
                "       Decode(B.�����־, 3, '��', 2, '��', 1, '', 4, '�쳣', 5, '����', 6, '����', '') As �����־,nvl(a.�걾���,0) as ���� " & vbNewLine & _
                "From (Select ��Ŀid, С��λ�� From ����������Ŀ Where ����id = 162) F, ����ҽ����¼ E, ���ű� D, ������Ŀ C, ������ͨ��� B, ����걾��¼ A" & vbNewLine & _
                "Where B.������Ŀid = F.��Ŀid(+) And A.ҽ��id = E.ID And E.��������id = D.ID And B.������Ŀid = C.������Ŀid And A.ID = B.����걾id And" & vbNewLine & _
                "      To_Number(A.�걾���) Between [1] And [2] And B.������ Is Not Null And A.������ = B.��¼���� And Nvl(A.����id, 0) = [3] And" & vbNewLine & _
                "      A.����ʱ�� Between [4] And [5]" & vbNewLine & _
                "Order By LPad(�걾���, 4, '0'),  nvl(a.�걾���,0),B.�������"


110     Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngStar���, lngEnd���, lng����ID, CDate(dateStart), CDate(dateEnd))
   
112     strLastRow = ""
114     With vsPrint
116         .Clear
118         .Rows = 4: .Cols = 8
120         .MergeCellsFixed = flexMergeRestrictColumns
        
122         .Cell(flexcpText, 0, 0, 0, .Cols - 1) = "ÿ�չ�������"
124         .TextMatrix(0, 0) = "ÿ�չ�������"
126         .MergeRow(0) = True
128         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        
130         .Cell(flexcpText, 1, 0, 1, 6) = "���ڣ�" & Format(dateStart, "yyyy-MM-dd")
132         .Cell(flexcpText, 1, 7, 1, 7) = "������" & cbo����.List(cbo����.ListIndex)
134         .MergeRow(1) = True
136         .Cell(flexcpAlignment, 1, 0, 1, .Cols - 1) = flexAlignLeftCenter
138         .Cell(flexcpAlignment, 1, .Cols - 1, 1, .Cols - 1) = flexAlignRightCenter
        
140         .TextMatrix(2, 0) = "���": .TextMatrix(2, 1) = "��Ա��": .TextMatrix(2, 2) = "����"
142         .TextMatrix(2, 3) = "�Ա�": .TextMatrix(2, 4) = "����": .TextMatrix(2, 5) = "����"
144         .TextMatrix(2, 6) = "����": .TextMatrix(2, 7) = "���"
146         .Cell(flexcpAlignment, 2, 0, 2, .Cols - 1) = flexAlignCenterCenter
        
148         .Cell(flexcpBackColor, 0, 0, 2, .Cols - 1) = vbWhite
150         .GridLines = flexGridNone
152         .GridLinesFixed = flexGridNone
        End With
154     Do Until rsTmp.EOF
156         With vsPrint
    '            If .TextMatrix(.Rows - 1, 0) <> "" Then
    '                .Rows = .Rows + 1
    '            End If
            
158             If strLastRow <> rsTmp.Fields("�걾���") Then
160                 If strLastRow <> "" And iCol <> 0 Then
162                     .Rows = .Rows + 1
164                     .RowHeight(.Rows - 1) = 320
                    End If
166                 .TextMatrix(.Rows - 1, 0) = "" & rsTmp.Fields("�걾���")
168                 .TextMatrix(.Rows - 1, 1) = "" & rsTmp.Fields("��Ա��")
170                 .TextMatrix(.Rows - 1, 2) = "" & rsTmp.Fields("����")
172                 .TextMatrix(.Rows - 1, 3) = "" & rsTmp.Fields("�Ա�")
174                 .TextMatrix(.Rows - 1, 4) = "" & rsTmp.Fields("����")
176                 .TextMatrix(.Rows - 1, 5) = "" & rsTmp.Fields("����")
178                 .TextMatrix(.Rows - 1, 6) = "" & rsTmp.Fields("����")
180                 .Select .Rows - 1, 0, .Rows - 1, .Cols - 1
182                 .CellBorder vbBlack, 0, 1, 0, 0, 0, 0
184                 lngCount = lngCount + 1
186                 iCol = 0
188             ElseIf strEmergency <> rsTmp.Fields("����") Then
190                 If strLastRow <> "" And iCol <> 0 Then
192                     .Rows = .Rows + 1
194                     .RowHeight(.Rows - 1) = 320
                    End If
196                 .TextMatrix(.Rows - 1, 0) = "" & rsTmp.Fields("�걾���")
198                 .TextMatrix(.Rows - 1, 1) = "" & rsTmp.Fields("��Ա��")
200                 .TextMatrix(.Rows - 1, 2) = "" & rsTmp.Fields("����")
202                 .TextMatrix(.Rows - 1, 3) = "" & rsTmp.Fields("�Ա�")
204                 .TextMatrix(.Rows - 1, 4) = "" & rsTmp.Fields("����")
206                 .TextMatrix(.Rows - 1, 5) = "" & rsTmp.Fields("����")
208                 .TextMatrix(.Rows - 1, 6) = "" & rsTmp.Fields("����")
210                 .Select .Rows - 1, 0, .Rows - 1, .Cols - 1
212                 .CellBorder vbBlack, 0, 1, 0, 0, 0, 0
214                 lngCount = lngCount + 1
216                 iCol = 0
                End If
            
218             iCol = iCol + 1
220             If iCol <= 4 Then
222                 str��Ϣ�� = ""
224                 If rsTmp.Fields("������") = "-" Then
226                     str��Ϣ�� = "����"
228                 ElseIf rsTmp.Fields("������") = "+" Then
230                     str��Ϣ�� = "����"
232                 ElseIf rsTmp.Fields("������") = "*" Then
234                     str��Ϣ�� = "*.**"
236                 ElseIf rsTmp.Fields("�������") = "1" And InStr(rsTmp.Fields("������"), ".") > 0 Then
238                     strС�� = String(IIf(IsNull(rsTmp.Fields("С��λ��")), 2, Val("" & rsTmp.Fields("С��λ��"))), "0")
240                     str��Ϣ�� = Format("" & rsTmp.Fields("������"), "0." & strС��)
                    Else
242                     str��Ϣ�� = rsTmp.Fields("������")
                    End If
244                 str��Ϣ�� = rsTmp.Fields("��д") & " " & str��Ϣ�� & rsTmp.Fields("�����־")
246                 If LenB(StrConv(str��Ϣ��, vbFromUnicode)) < 16 Then
248                     str��Ϣ�� = str��Ϣ�� & Space(16 - LenB(StrConv(str��Ϣ��, vbFromUnicode)))
                    Else
250                     str��Ϣ�� = zlCommFun.ToVarchar(str��Ϣ��, 16, " ")
                    End If
252                 .TextMatrix(.Rows - 1, 7) = .TextMatrix(.Rows - 1, 7) & str��Ϣ�� & IIf(iCol >= 4, Space(2), "")
254                 If iCol >= 4 Then
256                     iCol = 0
258                     .Rows = .Rows + 1
260                     .RowHeight(.Rows - 1) = 320
                    End If
                End If
            
262             strLastRow = rsTmp.Fields("�걾���")
264             strEmergency = rsTmp.Fields("����")
            End With
266         rsTmp.MoveNext
        Loop
268     With vsPrint
270         .Rows = .Rows + 1
272         .TextMatrix(.Rows - 1, 7) = "��������:" & lngCount
274         .AutoSizeMode = flexAutoSizeColWidth
276         .AutoSize 0, .Cols - 1
        End With
278     If lngCount > 0 Then cmdPrint.Enabled = True
        '<EhFooter>
        Exit Sub

LoadDataToVsf_Err:
        WriteLog "frmLabReport", "LoadDataToVsf", CStr(Erl()) & "�У�" & Err.Description
288     If ErrCenter() = 1 Then
290         Resume
        End If
        '</EhFooter>
End Sub

