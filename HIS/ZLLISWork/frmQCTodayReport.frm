VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmQCTodayReport 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5085
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid vfgWord 
      Height          =   4935
      Left            =   4770
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   2175
      _cx             =   3836
      _cy             =   8705
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16761024
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   16777215
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   2
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   1
      FixedRows       =   1
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
      AutoSizeMode    =   1
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
      WordWrap        =   -1  'True
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
   Begin VSFlex8Ctl.VSFlexGrid vfgReport 
      Height          =   4935
      Left            =   75
      TabIndex        =   1
      Top             =   60
      Width           =   4605
      _cx             =   8123
      _cy             =   8705
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
      BackColorSel    =   16635590
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   2
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   1
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
      AutoSizeMode    =   1
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
      WordWrap        =   -1  'True
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
End
Attribute VB_Name = "frmQCTodayReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngID As Long

Private Enum mRow
    ��� = 0: ����: ��ʾ: ԭ��: ��ʩ: ����: ����: �鵵
End Enum

'��ʱ����
Dim lngCount As Long

'--------------------------------------------
'����Ϊ���幫������
'--------------------------------------------
Private Sub setListFormat()
    '���ܣ���ʼ�����òο�ֵ�б�
    '������ blnKeepData-�Ƿ������ݣ���ֻ���������ø�ʽ
    With Me.vfgReport
        .Redraw = flexRDNone
        .Clear
        .Rows = 8: .FixedRows = 0: .Cols = 2: .FixedCols = 1
        .TextMatrix(mRow.���, 0) = "���"
        .TextMatrix(mRow.����, 0) = "����"
        .TextMatrix(mRow.��ʾ, 0) = "��ʾ"
        .TextMatrix(mRow.ԭ��, 0) = "ԭ��"
        .TextMatrix(mRow.��ʩ, 0) = "��ʩ"
        .TextMatrix(mRow.����, 0) = "����"
        .TextMatrix(mRow.����, 0) = "����"
        .TextMatrix(mRow.�鵵, 0) = "�鵵"
        .ColWidth(0) = 500
        .Redraw = flexRDDirect
    End With
End Sub

Public Function zlRefresh(lngID As Long) As Boolean
    '���ܣ�������Ŀidˢ�µ�ǰ��ʾ����
    Dim rsTemp As New adodb.Recordset
    mlngID = lngID
    
    '�����ǰ����ʾ
    Call setListFormat
    If lngID = 0 Then zlRefresh = True: Exit Function
    
    '��ȡָ������Ϣ
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select ���, ����, ��ʾ, ԭ��, ��ʩ, ����, ������, ����ʱ��, �鵵��, �鵵ʱ�� From �����ʿر��� Where ���id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngID)
    If rsTemp.RecordCount > 0 Then
        With Me.vfgReport
            .Redraw = flexRDNone
            Select Case Val("" & rsTemp!���)
            Case 1: .TextMatrix(mRow.���, 1) = "���棡"
            Case 2: .TextMatrix(mRow.���, 1) = "ʧ�أ�"
            End Select
            .TextMatrix(mRow.����, 1) = "" & rsTemp!����
            .TextMatrix(mRow.��ʾ, 1) = "" & rsTemp!��ʾ
            .TextMatrix(mRow.ԭ��, 1) = "" & rsTemp!ԭ��
            .TextMatrix(mRow.��ʩ, 1) = "" & rsTemp!��ʩ
            .TextMatrix(mRow.����, 1) = "" & rsTemp!����
            .TextMatrix(mRow.����, 1) = rsTemp!������ & IIf(IsNull(rsTemp!������), "", ", ") & Format(rsTemp!����ʱ��, "yyyy��MM��dd�� hh:mm")
            .TextMatrix(mRow.�鵵, 1) = rsTemp!�鵵�� & IIf(IsNull(rsTemp!�鵵��), "", ", ") & Format(rsTemp!�鵵ʱ��, "yyyy��MM��dd�� hh:mm")
            .Redraw = flexRDDirect
            Call .AutoSize(1)
        End With
    End If
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function ZlEditStart(lngID As Long) As Boolean
    '���ܣ���ʼ��Ŀ�༭
    '������lngId-ָ���༭����Ŀ
    Dim rsTemp As New adodb.Recordset
    
    Me.Tag = "�༭": Call Form_Resize
    
    On Error Resume Next
    Me.vfgReport.SetFocus
    ZlEditStart = True
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ZlEditStart = False: Exit Function
End Function

Public Sub zlEditCancel()
    '���ܣ��������ڽ��еı༭
    Me.Tag = "": Call Form_Resize
    Call Me.zlRefresh(mlngID)
End Sub

Public Function zlEditSave() As Long
    '���ܣ��������ڽ��еı༭,���������ڱ༭��Ŀid,����ʧ�ܷ���0
    
    With Me.vfgReport
        If .EditWindow <> 0 Then .TextMatrix(.Row, 1) = .EditText
        .TextMatrix(mRow.ԭ��, 1) = Trim(Replace(Replace(Replace(Replace(.TextMatrix(mRow.ԭ��, 1), vbCrLf, ""), vbCr, ""), vbLf, ""), "'", ""))
        .TextMatrix(mRow.��ʩ, 1) = Trim(Replace(Replace(Replace(Replace(.TextMatrix(mRow.��ʩ, 1), vbCrLf, ""), vbCr, ""), vbLf, ""), "'", ""))
        .TextMatrix(mRow.����, 1) = Trim(Replace(Replace(Replace(Replace(.TextMatrix(mRow.����, 1), vbCrLf, ""), vbCr, ""), vbLf, ""), "'", ""))
        If .TextMatrix(mRow.ԭ��, 1) = "" And .TextMatrix(mRow.��ʩ, 1) = "" And .TextMatrix(mRow.����, 1) = "" Then
            If MsgBox("��û����д�κα������ݣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then zlEditSave = 0: Exit Function
        End If
        If LenB(StrConv(.TextMatrix(mRow.ԭ��, 1), vbFromUnicode)) > 500 Then
            MsgBox "ԭ�򳬳������500���ַ���ȳ����֣���", vbInformation, gstrSysName
            .Row = mRow.ԭ��: zlEditSave = 0: Exit Function
        End If
        If LenB(StrConv(.TextMatrix(mRow.��ʩ, 1), vbFromUnicode)) > 500 Then
            MsgBox "��ʩ���������500���ַ���ȳ����֣���", vbInformation, gstrSysName
            .Row = mRow.��ʩ: zlEditSave = 0: Exit Function
        End If
        If LenB(StrConv(.TextMatrix(mRow.����, 1), vbFromUnicode)) > 500 Then
            MsgBox "���۳��������500���ַ���ȳ����֣���", vbInformation, gstrSysName
            .Row = mRow.����: zlEditSave = 0: Exit Function
        End If
        gstrSql = mlngID & ",'" & .TextMatrix(mRow.ԭ��, 1) & "'"
        gstrSql = gstrSql & ",'" & .TextMatrix(mRow.��ʩ, 1) & "'"
        gstrSql = gstrSql & ",'" & .TextMatrix(mRow.����, 1) & "'"
    End With
    gstrSql = "Zl_�����ʿر���_Edit(" & gstrSql & ")"
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    Me.Tag = "": Call Form_Resize
    zlEditSave = mlngID: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

'--------------------------------------------
'����Ϊ����ؼ���Ӧ�¼�
'--------------------------------------------
Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    
    Me.vfgWord.Left = Me.ScaleWidth - Me.vfgWord.Width - Me.vfgReport.Left
    Me.vfgWord.Height = Me.ScaleHeight - Me.vfgWord.Top * 2
    With Me.vfgReport
        .Height = Me.ScaleHeight - .Top * 2
        If Me.Tag = "" Then
            Me.vfgWord.Visible = False
            .Width = Me.ScaleWidth - .Left * 2
            .Editable = flexEDNone
            .FocusRect = flexFocusNone
        Else
            Me.vfgWord.Visible = True
            .Width = Me.vfgWord.Left - 45 - .Left
            .Editable = flexEDKbdMouse
            .FocusRect = flexFocusHeavy
        End If
        Call Me.vfgReport.AutoSize(1)
    End With
End Sub

Private Sub vfgReport_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call Me.vfgReport.AutoSize(Col, Col)
End Sub

Private Sub vfgReport_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim rsTemp As New adodb.Recordset
    Dim strGroup As String
    
    Select Case NewRow
    Case mRow.ԭ��: strGroup = "ԭ��"
    Case mRow.��ʩ: strGroup = "��ʩ"
    Case mRow.����: strGroup = "����"
    Case Else: Me.vfgWord.Rows = Me.vfgWord.FixedRows: Exit Sub
    End Select
    
    If OldRow = NewRow Then Exit Sub
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select ���� As ""��ѡ�ʾ�:"" From �ʿر���ʾ� Where ���� Is Null Or ���� = [1] Order By ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strGroup)
    Set Me.vfgWord.DataSource = rsTemp
    Call Me.vfgWord.AutoSize(0)
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vfgReport_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Row
    Case mRow.ԭ��, mRow.��ʩ, mRow.����: Cancel = False
    Case Else: Cancel = True
    End Select
End Sub

Private Sub vfgReport_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub vfgWord_DblClick()
    With Me.vfgReport
        If Me.vfgWord.Row < Me.vfgWord.FixedRows Then Exit Sub
        If Trim(.TextMatrix(.Row, 1)) = "" Then
            .TextMatrix(.Row, 1) = Me.vfgWord.Text
        Else
            .TextMatrix(.Row, 1) = Trim(.TextMatrix(.Row, 1)) & "��" & Me.vfgWord.Text
        End If
        Call .AutoSize(1)
    End With
End Sub
