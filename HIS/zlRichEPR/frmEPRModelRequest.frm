VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmEPRModelRequest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ӧ������"
   ClientHeight    =   4920
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6615
   Icon            =   "frmEPRModelRequest.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdSave 
      Caption         =   "�ָ�(&R)"
      Enabled         =   0   'False
      Height          =   350
      Index           =   1
      Left            =   5415
      TabIndex        =   12
      Top             =   1770
      Width           =   1035
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(&S)"
      Enabled         =   0   'False
      Height          =   350
      Index           =   0
      Left            =   4380
      TabIndex        =   11
      Top             =   1770
      Width           =   1035
   End
   Begin VB.Frame fraLine 
      Height          =   15
      Left            =   -45
      TabIndex        =   9
      Top             =   345
      Width           =   6975
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "�ر�(&C)"
      Height          =   350
      Left            =   5350
      TabIndex        =   7
      Top             =   4470
      Width           =   1100
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "������Ӧ���ڵ�ǰ�ļ�������ʾ��(&T)��"
      Height          =   350
      Left            =   150
      TabIndex        =   6
      Top             =   4485
      Width           =   3555
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "���(&A)"
      Height          =   350
      Index           =   0
      Left            =   2055
      TabIndex        =   5
      Top             =   1770
      Width           =   1035
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "ɾ��(&D)"
      Height          =   350
      Index           =   1
      Left            =   3105
      TabIndex        =   4
      Top             =   1770
      Width           =   1035
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgItems 
      Height          =   3765
      Left            =   150
      TabIndex        =   0
      Top             =   645
      Width           =   1875
      _cx             =   3307
      _cy             =   6641
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16761024
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483643
      GridColorFixed  =   -2147483643
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
      Rows            =   5
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
   Begin VSFlex8Ctl.VSFlexGrid vfgVal 
      Height          =   1110
      Left            =   2055
      TabIndex        =   1
      Top             =   645
      Width           =   4380
      _cx             =   7726
      _cy             =   1958
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16761024
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483643
      GridColorFixed  =   -2147483643
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
      Rows            =   5
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
   Begin VSFlex8Ctl.VSFlexGrid vfgSel 
      Height          =   1980
      Left            =   2055
      TabIndex        =   2
      Top             =   2430
      Width           =   4395
      _cx             =   7752
      _cy             =   3492
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16761024
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483643
      GridColorFixed  =   -2147483643
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
      Rows            =   5
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
   Begin VB.Label lblSel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ѡ����ֵ: (˫���������Ҫ�ı�ѡֵΪ����ֵ)"
      Height          =   180
      Left            =   2070
      TabIndex        =   13
      Top             =   2220
      Width           =   3960
   End
   Begin VB.Label lblItems 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������Ŀ:"
      Height          =   180
      Left            =   165
      TabIndex        =   10
      Top             =   435
      Width           =   810
   End
   Begin VB.Image imgNote 
      Height          =   240
      Left            =   150
      Picture         =   "frmEPRModelRequest.frx":000C
      Top             =   60
      Width           =   240
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "���ݵ�ǰ�ļ���������࣬���Ի���������Ŀ�����ض�Ӧ��������"
      Height          =   180
      Left            =   525
      TabIndex        =   8
      Top             =   90
      Width           =   5220
   End
   Begin VB.Label lblVal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ѡ����ֵ: (����ĿΪ����ֵʱ��ʹ�ø�ʾ��)"
      Height          =   180
      Left            =   2055
      TabIndex        =   3
      Top             =   435
      Width           =   3780
   End
End
Attribute VB_Name = "frmEPRModelRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    ������ = 0: ����ֵ
End Enum

Private mlngDemoId As Long      '��ǰʾ��ID
Private mintPower As Integer    'ʾ������Ȩ��Χ
Private mblnOK As Boolean       '�Ƿ�ȷ��


'-----------------------------------------------------
'����Ϊ�ⲿ��������
'-----------------------------------------------------
Public Function ShowMe(ByVal frmParent As Form, ByVal lngDemoId As Long, ByVal intPower As Integer) As Boolean
    '���ܣ���ʾ���༭����
    '������ frmParent-������
    '       lngDemoId-�ʾ�ʾ��ID
    Dim rsTemp As New ADODB.Recordset
    mlngDemoId = lngDemoId: mintPower = intPower
    
    'װ���ѡ��������
    Err = 0: On Error GoTo ErrHand
    gstrSQL = "Select �ļ�id, ���� From ��������Ŀ¼ Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngDemoId)
    If rsTemp.RecordCount <= 0 Then MsgBox "��ǰʾ��������(���ܱ������û�ɾ��)��", vbInformation, gstrSysName: Exit Function
    Me.cmdApply.Tag = rsTemp!�ļ�ID
    If Val("" & rsTemp!����) = 0 Then
        Me.Caption = "����Ӧ������"
        Me.cmdApply.Caption = "������Ӧ���ڵ�ǰ�ļ������з���(&T)��"
    Else
        Me.Caption = "Ƭ��Ӧ������"
        Me.cmdApply.Caption = "������Ӧ���ڵ�ǰ�ļ�������Ƭ��(&T)��"
    End If
    
    If RefList = False Then MsgBox "û�к��ʵ�������Ŀ��", vbInformation, gstrSysName: Exit Function
    
    Me.Show vbModal, frmParent
    ShowMe = mblnOK
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function

'-----------------------------------------------------
'����Ϊ�ڲ����ó���
'-----------------------------------------------------
Private Function RefList(Optional strItem As String) As Boolean
    '���ܣ�ˢ��װ����Ŀ�б�����λ��ָ����Ŀ
Dim rsTemp As New ADODB.Recordset
Dim lngCount As Long

    Err = 0: On Error GoTo ErrHand
    gstrSQL = "Select ���� As ������, ���� As ����ֵ From Table(Cast(f_Segment_������([1]) As " & gstrDbOwner & ".t_Dic_Rowset))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngDemoId)
    With Me.vfgItems
        .Clear
        Set .DataSource = rsTemp
        .ColWidth(mCol.����ֵ) = 0: .ColHidden(mCol.����ֵ) = True
        For lngCount = .FixedRows To .Rows - 1
            .Cell(flexcpFontBold, lngCount, mCol.������) = (.TextMatrix(lngCount, mCol.����ֵ) <> "")
            If .TextMatrix(lngCount, mCol.������) = strItem Then .Row = lngCount
        Next
        If .Row < .FixedRows Then .Row = .FixedRows
        Call vfgItems_AfterRowColChange(.Row, .Col, .Row, .Col)
    End With
    RefList = True
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------
Private Sub cmdApply_Click()
    Err = 0: On Error GoTo ErrHand
    If MsgBox("��Ľ�������Ӧ���ڵ�ǰ�ļ�������ʾ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Me.vfgItems.SetFocus: Exit Sub
    End If
    gstrSQL = "Zl_������������_Apply(" & mlngDemoId & "," & mintPower & ")"
    zlDatabase.ExecuteProcedure gstrSQL, "��������"
    Me.vfgItems.SetFocus: Exit Sub
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click(Index As Integer)
    Dim strItem As String
    If Index = 0 Then
        If Me.vfgSel.Rows < 1 Then MsgBox "û�п���ӵ�ֵ��", vbInformation, gstrSysName: Me.vfgVal.SetFocus: Exit Sub
        If Me.vfgSel.Row < 0 Then MsgBox "û�п���ӵ�ֵ��", vbInformation, gstrSysName: Me.vfgVal.SetFocus: Exit Sub
        Select Case Me.vfgItems.TextMatrix(Me.vfgItems.Row, mCol.������)
        Case "�������", "�������"
            If Me.vfgVal.Rows >= 1 Then MsgBox "����Ŀֻ������һ������ֵ��", vbInformation, gstrSysName: Me.vfgVal.SetFocus: Exit Sub
        End Select
        Me.vfgVal.AddItem Me.vfgSel.TextMatrix(Me.vfgSel.Row, 0)
        Me.vfgSel.RemoveItem Me.vfgSel.Row
        Me.vfgVal.SetFocus
    Else
        If Me.vfgVal.Rows < 1 Then MsgBox "û�п���ȥ��ֵ��", vbInformation, gstrSysName: Me.vfgSel.SetFocus: Exit Sub
        If Me.vfgVal.Row < 0 Then MsgBox "û�п���ȥ��ֵ��", vbInformation, gstrSysName: Me.vfgSel.SetFocus: Exit Sub
        Me.vfgSel.AddItem Me.vfgVal.TextMatrix(Me.vfgVal.Row, 0)
        Me.vfgVal.RemoveItem Me.vfgVal.Row
        Me.vfgSel.SetFocus
    End If
    If Me.vfgVal.Rows > 0 And Me.vfgVal.Row < 0 Then Me.vfgVal.Row = 0
    If Me.vfgSel.Rows > 0 And Me.vfgSel.Row < 0 Then Me.vfgSel.Row = 0
    
    Me.cmdEdit(0).Enabled = (Me.vfgSel.Rows > 0): Me.cmdEdit(1).Enabled = (Me.vfgVal.Rows > 0)
    Me.cmdSave(0).Enabled = True: Me.cmdSave(1).Enabled = True
End Sub

Private Sub cmdSave_Click(Index As Integer)
Dim strItem As String, strTerm As String
Dim lngCount As Long
    Err = 0: On Error GoTo ErrHand
    strItem = Me.vfgItems.TextMatrix(Me.vfgItems.Row, mCol.������)
    If Index = 0 Then
        strTerm = ""
        With Me.vfgVal
            For lngCount = .FixedRows To .Rows - 1
                strTerm = strTerm & vbTab & .TextMatrix(lngCount, 0)
            Next
        End With
        If strTerm <> "" Then strTerm = Mid(strTerm, 2)
        gstrSQL = "Zl_������������_Edit(" & mlngDemoId & ",'" & strItem & "','" & strTerm & "')"
        zlDatabase.ExecuteProcedure gstrSQL, "��������"
        mblnOK = True
    Else
        If MsgBox("��ķ�����ǰ�������޸���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    Me.cmdSave(0).Enabled = False: Me.cmdSave(1).Enabled = False
    Call RefList(strItem)
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vfgItems_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
Dim rsTemp As New ADODB.Recordset
Dim strItem As String, aryVal() As String
Dim lngCount As Long
    
    Me.cmdEdit(0).Enabled = False: Me.cmdEdit(1).Enabled = False
    Me.cmdSave(0).Enabled = False: Me.cmdSave(1).Enabled = False
    Me.vfgVal.Clear: Me.vfgVal.Rows = 0
    Me.vfgSel.Clear: Me.vfgSel.Rows = 0
    If NewRow < 0 Then Exit Sub
    strItem = Me.vfgItems.TextMatrix(NewRow, mCol.������)
    aryVal = Split(Me.vfgItems.TextMatrix(NewRow, mCol.����ֵ), vbTab)
    With Me.vfgVal
        For lngCount = 0 To UBound(aryVal)
            .AddItem aryVal(lngCount)
        Next
        If .Rows > 0 Then .Row = 0
    End With
    
    Err = 0: On Error GoTo ErrHand
    gstrSQL = "Select ���� As ��ѡֵ From Table(Cast(f_Segment_��ѡֵ([1], [2]) As " & gstrDbOwner & ".t_Dic_Rowset))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngDemoId, strItem)
    Set Me.vfgSel.DataSource = rsTemp
    Me.cmdEdit(0).Enabled = (Me.vfgSel.Rows > 0)
    Me.cmdEdit(1).Enabled = (Me.vfgVal.Rows > 0)
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vfgItems_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Dim strItem As String
    If Me.cmdSave(0).Enabled = False Then Exit Sub
    strItem = Me.vfgItems.TextMatrix(OldRow, mCol.������)
    If MsgBox("�Ѿ�������'" & strItem & "'������ֵ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then Exit Sub
    Cancel = True
End Sub

Private Sub vfgSel_DblClick()
    If Me.vfgSel.Rows < 1 Then Exit Sub
    If Me.vfgSel.Row < 0 Then Exit Sub
    Call cmdEdit_Click(0)
End Sub

Private Sub vfgVal_DblClick()
    If Me.vfgVal.Rows < 1 Then Exit Sub
    If Me.vfgVal.Row < 0 Then Exit Sub
    Call cmdEdit_Click(1)
End Sub
