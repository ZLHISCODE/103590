VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAppRuleSaveTo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���Ϊ����"
   ClientHeight    =   3585
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6195
   Icon            =   "frmAppRuleSaveTo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4830
      TabIndex        =   3
      Top             =   3120
      Width           =   1245
   End
   Begin VB.CommandButton cmdSaveTo 
      Caption         =   "����(&S)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4830
      TabIndex        =   2
      Top             =   2715
      Width           =   1245
   End
   Begin VB.TextBox txt������ 
      Height          =   300
      Left            =   960
      MaxLength       =   30
      TabIndex        =   1
      Top             =   2745
      Width           =   3045
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   2505
      Left            =   105
      TabIndex        =   4
      Top             =   105
      Width           =   5970
      _cx             =   10530
      _cy             =   4419
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
   Begin VB.Label lblˮƽ�� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������ÿ���������ˮƽ��Ϊ**������"
      Height          =   180
      Left            =   960
      TabIndex        =   5
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Label lbl������ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������(&N)"
      Height          =   180
      Left            =   105
      TabIndex        =   0
      Top             =   2805
      Width           =   810
   End
End
Attribute VB_Name = "frmAppRuleSaveTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Enum mCol
    ������ = 0: ˮƽ��: ������
End Enum

Private mlngDevId As Long           '����id
Private mblnOK As Boolean
Private mlngGroupID As Long         '����ID
'-----------------------------------------------------
'��ʱ����
'-----------------------------------------------------
Dim rptCol As ReportColumn
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow

Dim lngCount As Long

'-----------------------------------------------------
'����Ϊ�ڲ���������
'-----------------------------------------------------
Public Function ShowMe(frmParent As Form, lngDevId As Long, lngGroupID As Long) As Boolean
    '���ܣ�ˢ��װ��ָ������
    Dim rsTemp As New ADODB.Recordset
    
    mlngDevId = lngDevId
    mlngGroupID = lngGroupID
    
    gstrSql = "Select Decode(A.�ʿ�ˮƽ��, Null, 1, 0, 1, A.�ʿ�ˮƽ��) As ˮƽ�� From �������� A Where A.ID = [1]"
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngDevId)
    
    Me.lblˮƽ��.Caption = "������ÿ���������ˮƽ��Ϊ" & rsTemp!ˮƽ�� & "������"
    Me.txt������.Text = "�·���" & lngDevId & "(N=" & rsTemp!ˮƽ�� & ")"
    
    gstrSql = "Select Distinct ������, ˮƽ��, '������ÿ���������ˮƽ��Ϊ' || ˮƽ�� || '������...' As ������" & vbNewLine & _
        "From �����ʿط���" & vbNewLine & _
        "Order By ������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Set Me.vfgList.DataSource = rsTemp
    Me.vfgList.ColWidth(mCol.ˮƽ��) = 0
    Me.vfgList.ColHidden(mCol.ˮƽ��) = True
    
    mblnOK = False
    Me.Show vbModal, frmParent
    ShowMe = mblnOK
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ShowMe = False
End Function

'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------
Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdSaveTo_Click()
    If Trim(Me.txt������.Text) = "" Then
        MsgBox "�����뷶������", vbInformation, gstrSysName
        Me.txt������.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt������.Text), vbFromUnicode)) > Me.txt������.MaxLength Then
        MsgBox "���������������" & Me.txt������.MaxLength & "���ַ���ȳ����֣���", vbInformation, gstrSysName
        Me.txt������.SetFocus: Exit Sub
    End If
    
    With Me.vfgList
        For lngCount = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(lngCount, mCol.������)) = Trim(Me.txt������.Text) Then
                If MsgBox("���Ҫ�滻������" & .TextMatrix(.Row, mCol.������) & "����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        Next
    End With
    
    gstrSql = "Zl_�����ʿط���_Edit(1,'" & Trim(Me.txt������.Text) & "'," & mlngDevId & "," & mlngGroupID & ")"
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    mblnOK = True
    Unload Me: Exit Sub
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt������_GotFocus()
    Me.txt������.SelStart = 0: Me.txt������.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vfgList_DblClick()
    With Me.vfgList
        If .Row < .FixedRows Then Exit Sub
        Me.txt������.Text = .TextMatrix(.Row, mCol.������)
    End With
End Sub

Private Sub vfgList_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyDelete Then Exit Sub
    With Me.vfgList
        If .Row < .FixedRows Then Exit Sub
        If MsgBox("���Ҫɾ��������" & .TextMatrix(.Row, mCol.������) & "����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        gstrSql = "Zl_�����ʿط���_Edit(2,'" & Trim(.TextMatrix(.Row, mCol.������)) & "')"
    End With
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    Me.vfgList.RemoveItem Me.vfgList.Row
    Exit Sub
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
