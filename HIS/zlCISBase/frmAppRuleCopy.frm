VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.Form frmAppRuleCopy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����������"
   ClientHeight    =   7515
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7455
   Icon            =   "frmAppRuleCopy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin XtremeReportControl.ReportControl rptRule 
      Height          =   2865
      Left            =   150
      TabIndex        =   2
      Top             =   4485
      Width           =   7095
      _Version        =   589884
      _ExtentX        =   12515
      _ExtentY        =   5054
      _StockProps     =   0
      BorderStyle     =   2
      AllowColumnSort =   0   'False
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "�˳�(&E)"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   3825
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "����(&C)"
      Height          =   375
      Left            =   4710
      TabIndex        =   0
      Top             =   3825
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   3615
      Left            =   150
      TabIndex        =   3
      Top             =   105
      Width           =   7095
      _cx             =   12515
      _cy             =   6376
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
      BackColorFixed  =   -2147483633
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
      Cols            =   3
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   3795
      Top             =   4155
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppRuleCopy.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppRuleCopy.frx":05A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppRuleCopy.frx":0B40
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppRuleCopy.frx":10DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "ע�⣺����ʱ��ɾ����ǰ���Ѿ������˵Ĺ���"
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   195
      TabIndex        =   5
      Top             =   3900
      Width           =   3780
   End
   Begin VB.Label lblԴ 
      Caption         =   "�����ƹ���"
      Height          =   210
      Left            =   180
      TabIndex        =   4
      Top             =   4245
      Width           =   2895
   End
End
Attribute VB_Name = "frmAppRuleCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mlng����ID As Long, mlng��Ŀid As Long
Private rptCol As ReportColumn, rptRcd As ReportRecord, rptRow As ReportRow
Private Enum mColR
    ���� = 0: ID: �ж�: ����: ����Χ: ��ˮƽ: ���ϴ���: ��������: �Ƿ�ʹ��
End Enum

Public Sub ShowMe(ByVal lng����ID As Long, ByVal lng��Ŀid As Long, frmMain As Form)

    If lng����ID = 0 Or lng��Ŀid = 0 Then Exit Sub
    mlng����ID = lng����ID
    mlng��Ŀid = lng��Ŀid
    
    Me.Show vbModal, frmMain
    
End Sub

Private Function zlRefRule() As Long
    '���ܣ�ˢ��װ�뵱ǰ�����Ĺ���
    Dim rsTemp As New ADODB.Recordset
    Dim blnCopy As Boolean
    gstrSql = "Select R.ID, R.�ϼ�id, R.����, Decode(R.����, 'Y', '����: ', 'N', '����: ', '') || R.�ж� As �ж�, B.���� As ����," & vbNewLine & _
            "       Decode(R.����Χ, 1, '��ǰ��', '��' || R.����Χ || '��') As ����Χ, Decode(R.��ˮƽ, 1, '��', '') As ��ˮƽ," & vbNewLine & _
            "       Decode(Y����, 0, '��һ��', '����') As ���ϴ���, Decode(N����, 0, '��һ��', '����') As ��������,Decode(�Ƿ�ʹ��,1,'��','') as �Ƿ�ʹ��" & vbNewLine & _
            "From (Select Level As ���, ID, Nvl(�ϼ�id, 0) As �ϼ�id, ����, �ж�, ����id, ����Χ, ��ˮƽ, Y����, N����, �Ƿ�ʹ��" & vbNewLine & _
            "       From ������������" & vbNewLine & _
            "       Where ����id = [1] And nvl(��ĿID,0)=[2] " & vbNewLine & _
            "       Start With ����id = [1] And �ϼ�id Is Null" & vbNewLine & _
            "       Connect By Prior ID = �ϼ�id) R, �����ʿع��� B" & vbNewLine & _
            "Where R.����id = B.ID" & vbNewLine & _
            "Order By R.���, Decode(R.����, '0', 0, '1', 1, 'Y', 2, 'N', 3, 1)"
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlng����ID, mlng��Ŀid)
    Err = 0: On Error GoTo 0
    Me.rptRule.Records.DeleteAll
    Me.rptRule.Populate
    With rsTemp
    
        Do While Not .EOF
            blnCopy = True
            If Val("" & !�ϼ�ID) = 0 Then
                Set rptRcd = Me.rptRule.Records.Add()
            Else
                Me.rptRule.Populate
                For Each rptRow In Me.rptRule.Rows
                    If Val(rptRow.Record(mColR.ID).Value) = Val("" & !�ϼ�ID) Then
                        Set rptRcd = rptRow.Record.Childs.Add()
                    End If
                Next
            End If
            If "" & !���� = "1" Then
                rptRcd.AddItem("1").Icon = 3
            Else
                rptRcd.AddItem(CStr("" & !����)).Icon = 2

            End If
            rptRcd.AddItem CStr(!ID)
            rptRcd.AddItem CStr(!�ж�)
            rptRcd.AddItem CStr("" & !����)
            rptRcd.AddItem CStr("" & !����Χ)
            rptRcd.AddItem CStr("" & !��ˮƽ)
            rptRcd.AddItem CStr("" & !���ϴ���)
            rptRcd.AddItem CStr("" & !��������)
            rptRcd.AddItem (CStr("" & !�Ƿ�ʹ��))
            rptRcd.Expanded = True
            .MoveNext
        Loop
    End With
    Me.rptRule.Populate
    OKButton.Enabled = blnCopy
    zlRefRule = Me.rptRule.Records.Count
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefRule = Me.rptRule.Records.Count
End Function

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
  
    
    '-----------------------------------------------------
    '�����б��ʼ��
    With Me.rptRule
        .AutoColumnSizing = True
        Set rptCol = .Columns.Add(mColR.����, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False
        rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mColR.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mColR.�ж�, "�ж�����", 150, True): rptCol.Editable = False: rptCol.Groupable = False
        rptCol.TreeColumn = True
        Set rptCol = .Columns.Add(mColR.����, "�жϹ���", 82, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mColR.����Χ, "����Χ", 45, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mColR.��ˮƽ, "��ˮƽ", 45, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mColR.���ϴ���, "���ϴ���", 55, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mColR.��������, "��������", 55, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mColR.�Ƿ�ʹ��, "�Ƿ�ʹ��", 55, False): rptCol.Editable = False: rptCol.Groupable = False
        .SetImageList Me.ImgList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,��������Ŀ..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With
    
    Call zlRefRule
    
    '-----------------------------------------------------
    '��ʼ���ɸ��Ƶ���Ŀ
    
    strSQL = "Select Distinct C.��ĿID,Null as ѡ��,I.����, I.���� As ������, L.��д As Ӣ����" & vbNewLine & _
        "From ����������Ŀ C, ������Ŀ L, ���鱨����Ŀ R, ������ĿĿ¼ I, �����ʿ�Ʒ��Ŀ A" & vbNewLine & _
        "Where A.��Ŀid = C.��Ŀid And C.��Ŀid = L.������Ŀid And L.������Ŀid = R.������Ŀid And R.������Ŀid = I.ID And" & vbNewLine & _
        "      I.�����Ŀ <> 1 And L.��Ŀ��� <> 2 And C.����id = [1] And C.��ĿID<>[2] " & vbNewLine & _
        "Order By I.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��Ŀid)
    Set Me.vfgList.DataSource = rsTmp
    Me.vfgList.ColWidth(0) = 0
    Me.vfgList.ColHidden(0) = True
    Me.vfgList.Cell(flexcpChecked, 1, 1, Me.vfgList.Rows - 1, 1) = flexUnchecked
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub OKButton_Click()
    Dim intRow As Integer, lngObj��Ŀid As Long
    Dim strSQL As String
    Dim blnCheck As Boolean
    On Error GoTo errHandle
    
    With Me.vfgList
        For intRow = 1 To .Rows - 1
            If .Cell(flexcpChecked, intRow, 1) = flexChecked Then
                lngObj��Ŀid = Val("" & .TextMatrix(intRow, 0))
                If lngObj��Ŀid <> 0 Then
                    strSQL = "Zl_������������_Copy(" & mlng����ID & "," & mlng��Ŀid & "," & lngObj��Ŀid & ")"
                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                End If
                blnCheck = True
            End If
        Next
    End With
    
    If Not blnCheck Then
        MsgBox "������ѡ��һ����Ŀ��Ȼ���ٵ㸴�ƣ�", vbInformation, Me.CancelButton
    Else
        Unload Me
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vfgList_Click()
    With Me.vfgList
        Debug.Print .MouseCol
        If .MouseCol = 1 Then
            .Cell(flexcpChecked, .Row, 1) = IIf(.Cell(flexcpChecked, .Row, 1) = flexUnchecked, flexChecked, flexUnchecked)
        End If
    End With
End Sub
