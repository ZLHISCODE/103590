VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmAppforBillDesignEditItem 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��Ŀ�༭"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5160
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6810
      TabIndex        =   1
      Top             =   3600
      Width           =   1335
   End
   Begin XtremeSuiteControls.TabControl TabcrlPage 
      Height          =   3495
      Left            =   30
      TabIndex        =   3
      Top             =   -30
      Width           =   8745
      _Version        =   589884
      _ExtentX        =   15425
      _ExtentY        =   6165
      _StockProps     =   64
      Enabled         =   -1  'True
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFItem 
      Height          =   225
      Index           =   0
      Left            =   2790
      TabIndex        =   4
      Top             =   3510
      Width           =   225
      _cx             =   397
      _cy             =   397
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
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
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
      Editable        =   0
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
   Begin VB.Label Label1 
      Caption         =   "��ѡ�����뵥�������Ŀ��δ��ʾ����Ŀ������û�ж������Ʊ��룩"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   30
      TabIndex        =   0
      Top             =   3510
      Width           =   3690
   End
End
Attribute VB_Name = "frmAppforBillDesignEditItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnShow As Boolean                         '�����Ƿ���ʾ
Private mlngkeyID As Long                           '���뵥ID
Private mblnAllSite As Boolean                      '�Ƿ��в鿴����վ��Ȩ��
Private mblnTre As Boolean                          '�Ƿ�����������

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If SaveData = True Then
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    If mblnShow = False Then
        Call LoadDate
        Call ReadVSFSel(mlngkeyID)
        mblnShow = True
    End If
End Sub

Private Sub Form_Load()
    With Me.TabcrlPage
        .Icons = frmPubIcons.imgPublic.Icons
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.OneNoteColors = True
        .PaintManager.BoldSelected = True
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnShow = False
    mblnAllSite = False
    mblnTre = False
End Sub

Private Sub LoadDate()
      '���ܶ�������
          Dim strSQL As String
          Dim rsType As ADODB.Recordset
          Dim rsItem As ADODB.Recordset
          Dim intloop As Integer
          Dim intCol As Integer
          Dim intCols As Integer
          Dim intRow As Integer

1         On Error GoTo LoadDate_Error

2         If mblnAllSite Or gUserInfo.NodeNo = "-" Then
3             strSQL = "select distinct ���� from ���������Ŀ where ͣ������ is null and nvl(�Ƿ�������Ŀ,0)=[2] and ���Ʊ��� is not null"
4         Else
5             strSQL = "select distinct ���� from ���������Ŀ where ͣ������ is null and (վ�� = [1] Or վ�� Is Null)  and nvl(�Ƿ�������Ŀ,0)=[2]  and ���Ʊ��� is not null"
6         End If
7         Set rsType = ComOpenSQL(Sel_Lis_DB, strSQL, "���������Ŀ", gUserInfo.NodeNo, IIf(mblnTre, 1, 0))

8         Do Until rsType.EOF
9             If intloop > 0 Then
10                Load vsfItem(intloop)
11            End If

12            Call TabcrlPage.InsertItem(intloop, rsType("����") & "", vsfItem(intloop).hWnd, 0)

13            With vsfItem(intloop)
14                .GridLines = flexGridNone
15                .Cols = 6: intCols = 6
16                .ColKey(0) = "id1": .ColHidden(0) = True
17                .ColKey(1) = "��Ŀ1": .ColWidth(1) = 1800
18                .ColKey(2) = "id2": .ColHidden(2) = True
19                .ColKey(3) = "��Ŀ2": .ColWidth(3) = 1800
20                .ColKey(4) = "id3": .ColHidden(4) = True
21                .ColKey(5) = "��Ŀ3": .ColWidth(5) = 1800

22                intCol = 0
23                intRow = 0
24                If mblnAllSite Or gUserInfo.NodeNo = "-" Then
25                    strSQL = "select id,����,����,���� from ���������Ŀ where ͣ������ is null and ���� = [1] and nvl(�Ƿ�������Ŀ,0)=[3]"
26                Else
27                    strSQL = "select id,����,����,���� from ���������Ŀ where ͣ������ is null and ���� = [1]  and (վ�� = [2] Or վ�� Is Null)  and nvl(�Ƿ�������Ŀ,0)=[3]"
28                End If

29                Set rsItem = ComOpenSQL(Sel_Lis_DB, strSQL, "���������Ŀ", rsType("����") & "", gUserInfo.NodeNo, IIf(mblnTre, 1, 0))
30                Do Until rsItem.EOF
31                    .TextMatrix(intRow, intCol) = rsItem("id")
32                    .TextMatrix(intRow, intCol + 1) = rsItem("����"): .Cell(flexcpChecked, intRow, intCol + 1, intRow, intCol + 1) = 2

33                    If intCol + 2 >= intCols Then
34                        intRow = intRow + 1
35                        .Rows = intRow + 1
36                        intCol = 0
37                    Else
38                        intCol = intCol + 2
39                    End If

40                    rsItem.MoveNext
41                Loop
42            End With
              
43            Call SetColWith(vsfItem(intloop))
44            intloop = intloop + 1
45            rsType.MoveNext
46        Loop


47        Exit Sub
LoadDate_Error:
48        Call WriteErrLog("zl9LisInsideComm", "frmAppforBillDesignEditItem", "ִ��(LoadDate)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
49        Err.Clear
End Sub

Private Sub VSFItem_Click(Index As Integer)
    '��ѡѡ��
    With Me.vsfItem(Index)
        If .MouseRow >= 0 And .MouseCol >= 0 Then
            If .TextMatrix(.MouseRow, .MouseCol) = "" Then Exit Sub
            If .Cell(flexcpChecked, .MouseRow, .MouseCol) = 1 Then
                .Cell(flexcpChecked, .MouseRow, .MouseCol) = 2
            Else
                .Cell(flexcpChecked, .MouseRow, .MouseCol) = 1
            End If
        End If
    End With
End Sub

Private Sub VSFItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    With vsfItem(Index)
        If .MouseCol >= 0 And .MouseRow >= 0 Then
            .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
        End If
    End With
End Sub

Private Sub SetColWith(ByVal objVSF As VSFlexGrid)
    Dim lngColWidth As Long
    With objVSF
        .Width = TabcrlPage.Width
        lngColWidth = .Width / 3 - 100
        .AutoSize 0, .Cols - 1
        .ColWidth(1) = lngColWidth
        .ColWidth(3) = lngColWidth
    End With
End Sub

Private Function GetVSFSel() As String
          Dim intType As Integer
          Dim intCol As Integer
          Dim intRow As Integer
          
1         On Error GoTo GetVSFSel_Error

2         For intType = vsfItem.LBound To vsfItem.UBound
3             With vsfItem(intType)
4                 For intRow = 0 To .Rows - 1
5                     For intCol = 0 To .Cols / 2 - 1
6                         If .Cell(flexcpChecked, intRow, intCol * 2 + 1, intRow, intCol * 2 + 1) = 1 Then
7                             GetVSFSel = GetVSFSel & "," & .TextMatrix(intRow, intCol * 2)
8                         End If
9                     Next
10                Next
11            End With
12        Next
13        If GetVSFSel <> "" Then
14            GetVSFSel = Mid(GetVSFSel, 2)
15        End If


16        Exit Function
GetVSFSel_Error:
17        Call WriteErrLog("zl9LisInsideComm", "frmAppforBillDesignEditItem", "ִ��(GetVSFSel)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
18        Err.Clear
End Function

Private Sub ReadVSFSel(lngkeyID As Long)
          '����       ����ѡ��
          Dim intType As Integer
          Dim intCol As Integer
          Dim intRow As Integer
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strItem As String

1         On Error GoTo ReadVSFSel_Error

2         strSQL = "select ���ID from �������뵥��ϸ where ���뵥ID = [1] "
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "��ȡ�ѱ�����������", lngkeyID)
4         Do Until rsTmp.EOF
5             strItem = strItem & "," & rsTmp("���ID")
6             rsTmp.MoveNext
7         Loop
8         If Trim(strItem) = "" Then Exit Sub
          
9         strItem = Mid(strItem, 2)
          
10        For intType = vsfItem.LBound To vsfItem.UBound
11            With vsfItem(intType)
12                For intRow = 0 To .Rows - 1
13                    For intCol = 0 To .Cols / 2 - 1
14                        If InStr("," & strItem & ",", "," & Val(.TextMatrix(intRow, intCol * 2)) & ",") > 0 Then
15                            .Cell(flexcpChecked, intRow, intCol * 2 + 1, intRow, intCol * 2 + 1) = 1
16                        End If
17                    Next
18                Next
19            End With
20        Next


21        Exit Sub
ReadVSFSel_Error:
22        Call WriteErrLog("zl9LisInsideComm", "frmAppforBillDesignEditItem", "ִ��(ReadVSFSel)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
23        Err.Clear
          
End Sub

Public Sub ShowMe(frmObj As Object, lngkeyID As Long, ByVal blnAllSite As Boolean, ByVal blnTre As Boolean)
    '����       �򿪴��岢�������
    'blnTre     �Ƿ�����������
    mlngkeyID = lngkeyID
    mblnAllSite = blnAllSite
    mblnTre = blnTre
    Me.Show vbModal, frmObj
End Sub

Private Function SaveData() As Boolean
      '����           �������뵥�����Ŀ
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strItem As String
          Dim strArr() As String
          Dim blnTrs As Boolean
          Dim i As Integer

1         On Error GoTo SaveData_Error

2         strItem = GetVSFSel()

          '��ȡ��ѡ��Ŀ֮ǰ�����뵥�еķ���
3         strSQL = "Select a.����ID, b.Column_Value ���ID" & vbCrLf & _
                 " From (Select ����ID, ���ID From �������뵥��ϸ Where ���뵥ID = [1]) a," & vbCrLf & _
                 "     Table(Cast(F_Num2list([2]) As Zltools.T_Numlist)) B" & vbCrLf & _
                 " Where a.���ID(+) = b.Column_Value"
4         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�������뵥��ϸ", mlngkeyID, strItem)
5         strItem = ""
6         Do While Not rsTmp.EOF
7             strItem = strItem & ";" & rsTmp("���ID") & "," & rsTmp("����ID")
8             rsTmp.MoveNext
9         Loop
10        If strItem <> "" Then strItem = Mid(strItem, 2)

          '��������
11        gcnLisOracle.BeginTrans
12        blnTrs = True
13        strArr = TruncatedExtraLongStr(strItem, ";")
14        For i = 0 To UBound(strArr)
15            strSQL = "Zl_���뵥��ϸ_Insert('" & mlngkeyID & "','" & strArr(i) & "'," & i + 1 & ")"
16            ComExecuteProc Sel_Lis_DB, strSQL, "�������뵥��ϸ"
17        Next
18        gcnLisOracle.CommitTrans
19        blnTrs = False

20        SaveDBLog 18, 6, 0, "�༭", "�༭���뵥�����Ŀ:" & strItem, 1012, "���뵥����"
21        SaveData = True


22        Exit Function
SaveData_Error:
23        If blnTrs Then gcnLisOracle.RollbackTrans
24        Call WriteErrLog("zl9LisInsideComm", "frmAppforBillDesignEditItem", "ִ��(SaveData)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
25        Err.Clear

End Function
