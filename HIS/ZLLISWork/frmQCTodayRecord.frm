VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmQCTodayRecord 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�Ǽ�"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6885
   Icon            =   "frmQCTodayRecord.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   6885
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancle 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3630
      TabIndex        =   18
      Top             =   3315
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1665
      TabIndex        =   17
      Top             =   3315
      Width           =   1100
   End
   Begin VB.TextBox txt���� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   3525
      TabIndex        =   16
      Top             =   90
      Width           =   3225
   End
   Begin VB.TextBox txt������ 
      Enabled         =   0   'False
      Height          =   300
      Left            =   690
      TabIndex        =   11
      Top             =   480
      Width           =   1155
   End
   Begin VB.TextBox txt�걾�� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   690
      TabIndex        =   10
      Top             =   90
      Width           =   1155
   End
   Begin VB.ComboBox cbo�ʿ�Ʒ 
      Height          =   300
      ItemData        =   "frmQCTodayRecord.frx":000C
      Left            =   3525
      List            =   "frmQCTodayRecord.frx":000E
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   480
      Width           =   3225
   End
   Begin VB.CheckBox chk�������Լ� 
      Caption         =   "ʹ�����������Լ�"
      Height          =   210
      Left            =   120
      TabIndex        =   8
      Top             =   2685
      Width           =   1935
   End
   Begin VB.CheckBox chk�°�װ�Լ� 
      Caption         =   "ʹ�����°�װ�Լ�"
      Height          =   210
      Left            =   120
      TabIndex        =   7
      Top             =   2970
      Width           =   1935
   End
   Begin VB.CheckBox chk������У׼�� 
      Caption         =   "ʹ����������У׼��"
      Height          =   210
      Left            =   2385
      TabIndex        =   6
      Top             =   2685
      Width           =   1935
   End
   Begin VB.CheckBox chk�°�װУ׼�� 
      Caption         =   "ʹ�����°�װУ׼��"
      Height          =   210
      Left            =   2385
      TabIndex        =   5
      Top             =   2970
      Width           =   1935
   End
   Begin VB.CheckBox chk�°�װ������ 
      Caption         =   "ʹ�����°�װ������"
      Height          =   210
      Left            =   4815
      TabIndex        =   4
      Top             =   2685
      Width           =   1950
   End
   Begin VB.CheckBox chk����ά������ 
      Caption         =   "�ս�������ά������"
      Height          =   210
      Left            =   4815
      TabIndex        =   3
      Top             =   2970
      Width           =   1950
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "ѡ��(&S)"
      Height          =   350
      Left            =   1845
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   840
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgSelect 
      Height          =   2040
      Left            =   6915
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   315
      Visible         =   0   'False
      Width           =   4890
      _cx             =   8625
      _cy             =   3598
      Appearance      =   0
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
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
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
   Begin VSFlex8Ctl.VSFlexGrid vfgRecord 
      Height          =   1770
      Left            =   120
      TabIndex        =   0
      Top             =   870
      Width           =   6630
      _cx             =   11695
      _cy             =   3122
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
      BackColorSel    =   16635590
      ForeColorSel    =   -2147483640
      BackColorBkg    =   14737632
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
      Rows            =   5
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
   Begin VB.Label lbl������ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      Height          =   180
      Left            =   120
      TabIndex        =   15
      Top             =   540
      Width           =   540
   End
   Begin VB.Label lbl�ʿ�Ʒ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ʿ�Ʒ"
      Height          =   180
      Left            =   2910
      TabIndex        =   14
      Top             =   540
      Width           =   540
   End
   Begin VB.Label lbl�걾�� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�걾��"
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   150
      Width           =   540
   End
   Begin VB.Label lbl�������� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Left            =   3090
      TabIndex        =   12
      Top             =   150
      Width           =   360
   End
End
Attribute VB_Name = "frmQCTodayRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngID As Long
Private mstrSelDate As String
Private mblnAllDev As Boolean

Private Enum mCol
    ������ = 0: Ӣ����: ��λ: ���ֵ
End Enum

'��ʱ����
Dim lngCount As Long

'--------------------------------------------
'����Ϊ���幫������
'--------------------------------------------
Private Sub setListFormat(Optional blnKeepData As Boolean)
    '���ܣ���ʼ�����òο�ֵ�б�
    '������ blnKeepData-�Ƿ������ݣ���ֻ���������ø�ʽ
    With Me.vfgRecord
        .Redraw = flexRDNone
        If blnKeepData = False Then
            .Clear
            .Rows = 2: .FixedRows = 1: .Cols = 4: .FixedCols = 0
            .TextMatrix(0, mCol.������) = "������"
            .TextMatrix(0, mCol.Ӣ����) = "Ӣ����"
            .TextMatrix(0, mCol.��λ) = "��λ"
            .TextMatrix(0, mCol.���ֵ) = "���ֵ"
        End If
        .ColWidth(mCol.������) = 3000
        .ColWidth(mCol.Ӣ����) = 1200
        .ColWidth(mCol.��λ) = 1000
        .ColWidth(mCol.���ֵ) = 900
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            If .ColWidth(lngCount) = 0 Then .ColHidden(lngCount) = True
        Next
        .Redraw = flexRDDirect
    End With
End Sub

Public Function zlRefresh(lngID As Long) As Boolean
    '���ܣ�����idˢ�µ�ǰ��ʾ����
    Dim rsTemp As New ADODB.Recordset
    mlngID = lngID
    
    '�����ǰ��Ŀ����ʾ
    
    Me.txt�걾��.Text = "": Me.txt�걾��.Tag = "": Me.txt����.Text = ""
    Me.txt������.Text = "": Me.cbo�ʿ�Ʒ.Clear
    Me.chk�������Լ�.Value = vbUnchecked: Me.chk�°�װ�Լ�.Value = vbUnchecked
    Me.chk������У׼��.Value = vbUnchecked: Me.chk�°�װУ׼��.Value = vbUnchecked
    Me.chk�°�װ������.Value = vbUnchecked: Me.chk����ά������.Value = vbUnchecked
    
    If lngID = 0 Then Call setListFormat: zlRefresh = True: Exit Function
    
    '��ȡָ����Ŀ����Ϣ
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select L.�걾���, A.���� As ����, L.�ʿ�Ʒid, M.���� || '-' || M.���� As �ʿ�Ʒ, L.������, L.�������Լ�," & vbNewLine & _
            "       L.�°�װ�Լ�, L.������У׼��, L.�°�װУ׼��, L.�°�װ������, L.����ά������" & vbNewLine & _
            "From �����ʿؼ�¼ L, �������� A, �����ʿ�Ʒ M" & vbNewLine & _
            "Where L.����id = A.ID And L.�ʿ�Ʒid = M.ID And L.�걾id = [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngID)
    With rsTemp
        If .RecordCount > 0 Then
            Me.txt�걾��.Text = "" & !�걾���: Me.txt�걾��.Tag = lngID: Me.txt����.Text = "" & !����
            Me.txt������.Text = "" & !������
            Me.cbo�ʿ�Ʒ.AddItem "" & !�ʿ�Ʒ
            Me.cbo�ʿ�Ʒ.ItemData(Me.cbo�ʿ�Ʒ.NewIndex) = Val("" & !�ʿ�Ʒid)
            Me.cbo�ʿ�Ʒ.ListIndex = Me.cbo�ʿ�Ʒ.NewIndex
            Me.chk�������Լ�.Value = IIf(Val("" & !�������Լ�) = 0, vbUnchecked, vbChecked)
            Me.chk�°�װ�Լ�.Value = IIf(Val("" & !�°�װ�Լ�) = 0, vbUnchecked, vbChecked)
            Me.chk������У׼��.Value = IIf(Val("" & !������У׼��) = 0, vbUnchecked, vbChecked)
            Me.chk�°�װУ׼��.Value = IIf(Val("" & !�°�װУ׼��) = 0, vbUnchecked, vbChecked)
            Me.chk�°�װ������.Value = IIf(Val("" & !�°�װ������) = 0, vbUnchecked, vbChecked)
            Me.chk����ά������.Value = IIf(Val("" & !����ά������) = 0, vbUnchecked, vbChecked)
        End If
    End With
        
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function ZlEditStart(blnAdd As Boolean, lngID As Long, Optional strSelDate As String, Optional blnAllDev As Boolean) As Long
    '���ܣ���ʼ��Ŀ�༭
    '������ blnAdd-�Ƿ����ӣ�����Ϊ�޸�
    '       lngID-��ǰ�༭�ı걾id�����ߵ�ǰѡ�еı걾id
    '       strDate-ָ������
    '       blnAllDev-�Ƿ��������豸��Ȩ�ޣ�����ֻ���Ǳ����ŵ�����
    Dim rsTemp As New ADODB.Recordset
    
    If blnAdd Then
        Err = 0: On Error Resume Next
        mstrSelDate = Format(strSelDate, "yyyy-MM-dd")
        If Err <> 0 Or mstrSelDate = "" Then ZlEditStart = False: Exit Function
        Err = 0: On Error GoTo 0
        mblnAllDev = blnAllDev
    End If
    
    mlngID = lngID
    Call zlRefresh(lngID)
    If blnAdd Then
        Me.txt�걾��.Text = "": Me.txt�걾��.Tag = "": Me.txt����.Text = ""
        Me.txt������.Text = "": Me.cbo�ʿ�Ʒ.Clear
        Me.chk�������Լ�.Value = vbUnchecked: Me.chk�°�װ�Լ�.Value = vbUnchecked
        Me.chk������У׼��.Value = vbUnchecked: Me.chk�°�װУ׼��.Value = vbUnchecked
        Me.chk�°�װ������.Value = vbUnchecked: Me.chk����ά������.Value = vbUnchecked
        Call setListFormat(False)
    Else
        If Me.cbo�ʿ�Ʒ.ListIndex = -1 Then
            Me.cbo�ʿ�Ʒ.Tag = 0
        Else
            Me.cbo�ʿ�Ʒ.Tag = Me.cbo�ʿ�Ʒ.ItemData(Me.cbo�ʿ�Ʒ.ListIndex)
        End If
        Me.cbo�ʿ�Ʒ.Clear
        gstrSql = "Select Distinct M.ID, M.���� || '-' || M.���� || LPad('^,' || M.�걾��, 200, ' ') As �ʿ�Ʒ" & vbNewLine & _
                "From ����걾��¼ L, ������ͨ��� R, �����ʿ�Ʒ M, �����ʿ�Ʒ��Ŀ I" & vbNewLine & _
                "Where L.ID = R.����걾id And Nvl(L.������, 0) = Nvl(R.��¼����, 0) And Nvl(R.���ý��,0)=0 And L.����id = M.����id And M.ID = I.�ʿ�Ʒid And" & vbNewLine & _
                "      I.��Ŀid = R.������Ŀid And (L.����ʱ�� + 0 Between M.��ʼ���� And M.�������� + 1 - 1 / 86400) And L.ID = [1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngID)
        With rsTemp
            Do While Not .EOF
                Me.cbo�ʿ�Ʒ.AddItem "" & !�ʿ�Ʒ
                Me.cbo�ʿ�Ʒ.ItemData(Me.cbo�ʿ�Ʒ.NewIndex) = Val("" & !ID)
                If Val(Me.cbo�ʿ�Ʒ.Tag) = Val("" & !ID) Then Me.cbo�ʿ�Ʒ.ListIndex = Me.cbo�ʿ�Ʒ.NewIndex
                .MoveNext
            Loop
            If Me.cbo�ʿ�Ʒ.ListCount > 0 And Me.cbo�ʿ�Ʒ.ListIndex = -1 Then Me.cbo�ʿ�Ʒ.ListIndex = 0
        End With
    End If
    
    Me.Tag = IIf(blnAdd, "����", "�޸�"): Call Form_Resize
'    If blnAdd Then
'        Me.cmdSelect.SetFocus
'    Else
'        Me.cbo�ʿ�Ʒ.SetFocus
'    End If
    Me.Show vbModal
     
    ZlEditStart = mlngID
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
    Dim strInfo As String
    
    If Me.cbo�ʿ�Ʒ.ListIndex = -1 Then MsgBox "δѡ���ʿ�Ʒ��", vbInformation, gstrSysName: zlEditSave = 0: Exit Function
    
    strInfo = Split(Me.cbo�ʿ�Ʒ.Text, "^,")(1)
    If Trim(strInfo) <> "" And InStr(1, "," & strInfo & ",", "," & Trim(Me.txt�걾��.Text) & ",") = 0 Then
        strInfo = "��ǰ�걾�����ʿ�Ʒ�涨�������Ų��������飺"
        strInfo = strInfo & vbCrLf & "   ѡ���ʿ�Ʒ��Ũ��ˮƽ�Ƿ���ϣ�"
        strInfo = strInfo & vbCrLf & vbCrLf & "���ȷ����ȷ��ѡ���ǡ�������"
        If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then zlEditSave = 0: Exit Function
    End If

    gstrSql = "Zl_�����ʿؼ�¼_Edit(" & IIf(Me.Tag = "����", 1, 2)
    gstrSql = gstrSql & "," & Val(Me.txt�걾��.Tag) & "," & Me.cbo�ʿ�Ʒ.ItemData(Me.cbo�ʿ�Ʒ.ListIndex)
    gstrSql = gstrSql & "," & IIf(Me.chk�������Լ�.Value = vbChecked, 1, 0)
    gstrSql = gstrSql & "," & IIf(Me.chk�°�װ�Լ�.Value = vbChecked, 1, 0)
    gstrSql = gstrSql & "," & IIf(Me.chk������У׼��.Value = vbChecked, 1, 0)
    gstrSql = gstrSql & "," & IIf(Me.chk�°�װУ׼��.Value = vbChecked, 1, 0)
    gstrSql = gstrSql & "," & IIf(Me.chk�°�װ������.Value = vbChecked, 1, 0)
    gstrSql = gstrSql & "," & IIf(Me.chk����ά������.Value = vbChecked, 1, 0) & ")"
    
    Err = 0: On Error GoTo ErrHand
    zldatabase.ExecuteProcedure gstrSql, Me.Caption
    
    Me.Tag = "": Call Form_Resize
    zlEditSave = Val(Me.txt�걾��.Tag): Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

'--------------------------------------------
'����Ϊ����ؼ���Ӧ�¼�
'--------------------------------------------

Private Sub cbo�ʿ�Ʒ_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim lngID As Long, lngResId As Long
    
    lngID = Val(Me.txt�걾��.Tag)
    If lngID = 0 Then Call setListFormat(False): Exit Sub
    
    If Me.cbo�ʿ�Ʒ.ListIndex = -1 Then Exit Sub
    lngResId = Me.cbo�ʿ�Ʒ.ItemData(Me.cbo�ʿ�Ʒ.ListIndex)
    
    Err = 0: On Error GoTo ErrHand
    If Trim(Me.Tag) = "" Then
        gstrSql = "Select I.������, I.Ӣ����, I.��λ, R.������ As ���ֵ" & vbNewLine & _
            "From ������ͨ��� R, ����������Ŀ I" & vbNewLine & _
            "Where R.������Ŀid = I.ID And R.�Ƿ���� = 1 And Nvl(R.���ý��,0)=0 And R.����걾id = [1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngID)
    Else
        gstrSql = "Select I.������, I.Ӣ����, I.��λ, R.������ As ���ֵ" & vbNewLine & _
                "From ������ͨ��� R, ����������Ŀ I, (Select ��Ŀid From �����ʿ�Ʒ��Ŀ Where �ʿ�Ʒid = [2]) T" & vbNewLine & _
                "Where R.������Ŀid = I.ID And Nvl(R.���ý��,0)=0 And R.������Ŀid = T.��Ŀid And R.����걾id = [1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngID, lngResId)
    End If
    Set Me.vfgRecord.DataSource = rsTemp: Call setListFormat(True)
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo�ʿ�Ʒ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk�°�װ������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk�°�װ�Լ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk�°�װУ׼��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk�������Լ�_Click()
    If Me.chk�������Լ�.Value = vbChecked Then
        Me.chk�°�װ�Լ�.Value = vbChecked: Me.chk�°�װ�Լ�.Enabled = False
    Else
        Me.chk�°�װ�Լ�.Enabled = True
    End If
End Sub

Private Sub chk�������Լ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk������У׼��_Click()
    If Me.chk������У׼��.Value = vbChecked Then
        Me.chk�°�װУ׼��.Value = vbChecked: Me.chk�°�װУ׼��.Enabled = False
    Else
        Me.chk�°�װУ׼��.Enabled = True
    End If
End Sub

Private Sub chk������У׼��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk����ά������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mlngID = zlEditSave
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHand
    If mblnAllDev Then
        gstrSql = "Select L.ID, L.�걾��� As �걾��, D.���� As ����, L.������" & vbNewLine & _
                "From �������� D, ����걾��¼ L" & vbNewLine & _
                "Where L.����id = D.ID And Nvl(D.΢����, 0) = 0 And Nvl(L.�Ƿ��ʿ�Ʒ, 0) = 0 And" & vbNewLine & _
                "      Ltrim(L.����) Is Null And Nvl(L.����ID, 0) = 0 And" & vbNewLine & _
                "      L.����ʱ�� Between To_Date([1], 'yyyy-mm-dd') And To_Date([1], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "Order By D.����, L.�걾���"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrSelDate)
    Else
        gstrSql = "Select L.ID, L.�걾��� As �걾��, D.���� As ����, L.������" & vbNewLine & _
                "From (Select ID, ����, ΢���� From �������� Where ʹ��С��id In (Select ����id From ������Ա Where ��Աid = [2])) D," & vbNewLine & _
                "     ����걾��¼ L" & vbNewLine & _
                "Where L.����id = D.ID And Nvl(D.΢����, 0) = 0 And Nvl(L.�Ƿ��ʿ�Ʒ, 0) = 0 And" & vbNewLine & _
                "      Ltrim(L.����) Is Null And Nvl(L.����ID, 0) = 0 And" & vbNewLine & _
                "      L.����ʱ�� Between To_Date([1], 'yyyy-mm-dd') And To_Date([1], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "Order By D.����, L.�걾���"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrSelDate, glngUserId)
    End If
    If rsTemp.RecordCount <= 0 Then MsgBox "����(" & mstrSelDate & ")û�п�ѡ�����걾��¼��", vbInformation, gstrSysName: Exit Sub
    With Me.vfgSelect
        Set .DataSource = rsTemp
        .ColWidth(0) = 0: .ColHidden(0) = False
        .Left = Me.cmdSelect.Left:
        .Top = Me.cmdSelect.Top + Me.cmdSelect.Height
        .Height = Me.ScaleHeight - .Top - 300
        .ZOrder 0: .Visible = True: .SetFocus
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Call setListFormat(False)
End Sub

Private Sub Form_Resize()
    Select Case Trim(Me.Tag)
    Case "����"
        Me.Enabled = True: Me.BackColor = RGB(250, 250, 250)
        Me.cmdSelect.Visible = True
    Case "�޸�"
        Me.Enabled = True: Me.BackColor = RGB(250, 250, 250)
        Me.cmdSelect.Visible = False
        Me.vfgSelect.Visible = False
    Case Else
        Me.Enabled = False: Me.BackColor = &H8000000F
        Me.cmdSelect.Visible = False
        Me.vfgSelect.Visible = False
    End Select
    Me.chk�������Լ�.BackColor = Me.BackColor: Me.chk�°�װ�Լ�.BackColor = Me.BackColor
    Me.chk������У׼��.BackColor = Me.BackColor: Me.chk�°�װУ׼��.BackColor = Me.BackColor
    Me.chk�°�װ������.BackColor = Me.BackColor: Me.chk����ά������.BackColor = Me.BackColor
    
'    Me.chk�°�װ�Լ�.Top = Me.ScaleHeight - Me.cmdOK.Height - 350
'    Me.chk�°�װУ׼��.Top = Me.ScaleHeight - Me.cmdOK.Height - 350
'    Me.chk����ά������.Top = Me.ScaleHeight - Me.cmdOK.Height - 350
'    Me.chk�������Լ�.Top = Me.ScaleHeight - Me.cmdOK.Height - 350 * 2
'    Me.chk������У׼��.Top = Me.ScaleHeight - Me.cmdOK.Height - 350 * 2
'    Me.chk�°�װ������.Top = Me.ScaleHeight - Me.cmdOK.Height - 350 * 2
'    Me.vfgRecord.Height = Me.chk�������Լ�.Top - Me.vfgRecord.Top - 75
    
End Sub

Private Sub vfgSelect_DblClick()
    Dim rsTemp As New ADODB.Recordset
    
    With Me.vfgSelect
        If .Row < .FixedRows Then Exit Sub
        Me.txt�걾��.Text = .TextMatrix(.Row, 1): Me.txt�걾��.Tag = .TextMatrix(.Row, 0)
        Me.txt����.Text = .TextMatrix(.Row, 2): Me.txt������.Text = .TextMatrix(.Row, 3)
        .Visible = False
    End With
        
    If Me.cbo�ʿ�Ʒ.ListIndex = -1 Then
        Me.cbo�ʿ�Ʒ.Tag = 0
    Else
        Me.cbo�ʿ�Ʒ.Tag = Me.cbo�ʿ�Ʒ.ItemData(Me.cbo�ʿ�Ʒ.ListIndex)
    End If
    Me.cbo�ʿ�Ʒ.Clear
    
    gstrSql = "Select Distinct M.ID, M.���� || '-' || M.���� || LPad('^,' || M.�걾��, 200, ' ') As �ʿ�Ʒ" & vbNewLine & _
            "From ����걾��¼ L, ������ͨ��� R, �����ʿ�Ʒ M, �����ʿ�Ʒ��Ŀ I" & vbNewLine & _
            "Where L.ID = R.����걾id And Nvl(L.������, 0) = Nvl(R.��¼����, 0) And L.����id = M.����id And M.ID = I.�ʿ�Ʒid And" & vbNewLine & _
            "      Nvl(R.���ý��,0)=0 And I.��Ŀid = R.������Ŀid And (L.����ʱ�� + 0 Between M.��ʼ���� And M.�������� + 1 - 1 / 86400) And L.ID = [1]"
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(Me.txt�걾��.Tag))
    With rsTemp
        Do While Not .EOF
            Me.cbo�ʿ�Ʒ.AddItem "" & !�ʿ�Ʒ
            Me.cbo�ʿ�Ʒ.ItemData(Me.cbo�ʿ�Ʒ.NewIndex) = Val("" & !ID)
            If Val(Me.cbo�ʿ�Ʒ.Tag) = Val("" & !ID) Then Me.cbo�ʿ�Ʒ.ListIndex = Me.cbo�ʿ�Ʒ.NewIndex
            .MoveNext
        Loop
        If Me.cbo�ʿ�Ʒ.ListCount > 0 And Me.cbo�ʿ�Ʒ.ListIndex = -1 Then Me.cbo�ʿ�Ʒ.ListIndex = 0
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vfgSelect_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Call vfgSelect_DblClick
    ElseIf KeyCode = vbKeyReturn Then
        Call vfgSelect_DblClick
    ElseIf KeyCode = vbKeyEscape Then
        Me.vfgSelect.Visible = False
        Me.cmdSelect.SetFocus
    End If
End Sub

Private Sub vfgSelect_LostFocus()
    Me.vfgSelect.Visible = False
End Sub
