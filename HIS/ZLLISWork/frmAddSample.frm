VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAddSample 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�걾����"
   ClientHeight    =   7044
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   7620
   Icon            =   "frmAddSample.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7044
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComCtl2.DTPicker DTPApply 
      Height          =   255
      Left            =   4725
      TabIndex        =   29
      Top             =   5310
      Width           =   2745
      _ExtentX        =   4847
      _ExtentY        =   445
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   253558787
      CurrentDate     =   39475
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   5250
      TabIndex        =   27
      Top             =   3615
      Width           =   2325
   End
   Begin VB.TextBox txtComment 
      Height          =   300
      Left            =   4725
      TabIndex        =   12
      Top             =   6000
      Width           =   2730
   End
   Begin VB.ComboBox cboҽ�� 
      Height          =   300
      Left            =   6315
      TabIndex        =   10
      Top             =   4935
      Width           =   1155
   End
   Begin VB.ComboBox cbo�������� 
      Height          =   300
      ItemData        =   "frmAddSample.frx":020A
      Left            =   4725
      List            =   "frmAddSample.frx":020C
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   4935
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar PgbSave 
      Height          =   195
      Left            =   0
      TabIndex        =   23
      Top             =   6840
      Visible         =   0   'False
      Width           =   7635
      _ExtentX        =   13462
      _ExtentY        =   339
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CheckBox chkTest 
      Caption         =   "����ʾ΢���������Ŀ(&M)"
      Height          =   255
      Left            =   4935
      TabIndex        =   22
      Top             =   98
      Width           =   2460
   End
   Begin VB.ComboBox cboDevice 
      Height          =   300
      Left            =   4725
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   5640
      Width           =   2730
   End
   Begin VB.CheckBox chkOverWrite 
      Caption         =   "���ǵ��������ӵ���ͬ�����ż�¼(&W)"
      Height          =   195
      Left            =   90
      TabIndex        =   15
      Top             =   6495
      Width           =   3285
   End
   Begin VB.TextBox txt 
      Height          =   300
      Index           =   1
      Left            =   6540
      MaxLength       =   4
      TabIndex        =   7
      Text            =   "0001"
      Top             =   3795
      Width           =   585
   End
   Begin VB.TextBox txt 
      Height          =   300
      Index           =   0
      Left            =   5700
      MaxLength       =   4
      TabIndex        =   6
      Text            =   "0001"
      Top             =   3795
      Width           =   465
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5175
      TabIndex        =   13
      Top             =   6420
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6375
      TabIndex        =   14
      Top             =   6420
      Width           =   1100
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   3120
      Top             =   4110
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddSample.frx":020E
            Key             =   "���"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddSample.frx":07A8
            Key             =   "ָ��"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid Vsf 
      Height          =   2835
      Left            =   3900
      TabIndex        =   1
      Top             =   405
      Width           =   3615
      _cx             =   6376
      _cy             =   5001
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
      ForeColorSel    =   -2147483643
      BackColorBkg    =   -2147483643
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
      AllowUserResizing=   1
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   270
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
   Begin MSComctlLib.ListView lvwLabItem 
      Height          =   2835
      Left            =   90
      TabIndex        =   0
      Top             =   405
      Width           =   3810
      _ExtentX        =   6710
      _ExtentY        =   4995
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "��Ŀ"
         Text            =   "��Ŀ"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "����"
         Text            =   "����"
         Object.Width           =   1587
      EndProperty
   End
   Begin VB.ComboBox cboType 
      Height          =   300
      Left            =   2430
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   75
      Width           =   2340
   End
   Begin VB.OptionButton optType 
      Caption         =   "�����걾��(&2)"
      Height          =   180
      Index           =   1
      Left            =   4950
      TabIndex        =   5
      Top             =   4605
      Width           =   1530
   End
   Begin VB.OptionButton optType 
      Caption         =   "�����걾��(&1)"
      Height          =   180
      Index           =   0
      Left            =   4950
      TabIndex        =   4
      Top             =   4230
      Value           =   -1  'True
      Width           =   1530
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "��ɾ��ѡ�еı���������Ŀ(&D)"
      Height          =   350
      Left            =   2550
      TabIndex        =   17
      Top             =   3315
      Width           =   2580
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "����ӵ�����������Ŀ��(&S)"
      Height          =   350
      Left            =   30
      TabIndex        =   16
      Top             =   3315
      Width           =   2490
   End
   Begin VSFlex8Ctl.VSFlexGrid VsfSelect 
      Height          =   2220
      Left            =   60
      TabIndex        =   2
      Top             =   4065
      Width           =   3780
      _cx             =   6667
      _cy             =   3916
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
      ForeColorSel    =   -2147483643
      BackColorBkg    =   -2147483643
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
      AllowUserResizing=   1
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   270
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
   Begin VB.TextBox txt 
      Height          =   300
      Index           =   2
      Left            =   6540
      MaxLength       =   4
      TabIndex        =   8
      Top             =   4170
      Width           =   585
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   300
      Index           =   3
      Left            =   6540
      MaxLength       =   4
      TabIndex        =   11
      Top             =   4545
      Width           =   585
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "����ʱ��"
      Height          =   180
      Left            =   3930
      TabIndex        =   28
      Top             =   5340
      Width           =   720
   End
   Begin VB.Label lblSelect 
      AutoSize        =   -1  'True
      Caption         =   "����������Ŀ:"
      Height          =   180
      Left            =   90
      TabIndex        =   26
      Top             =   3840
      Width           =   1170
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "��    ע"
      Height          =   180
      Left            =   3930
      TabIndex        =   25
      Top             =   6075
      Width           =   720
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��    ��"
      Height          =   180
      Left            =   3930
      TabIndex        =   24
      Top             =   4995
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "��ʼ�걾�ţ���        ��        ��"
      Height          =   180
      Left            =   4290
      TabIndex        =   21
      Top             =   3840
      Width           =   3060
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��    ��"
      Height          =   180
      Left            =   3930
      TabIndex        =   19
      Top             =   5700
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ѡ����Ҫ��ӵ���Ŀ:  ����"
      Height          =   180
      Left            =   90
      TabIndex        =   18
      Top             =   135
      Width           =   2250
   End
End
Attribute VB_Name = "frmAddSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const COLOR_SELECT_BACK As Long = &H8000000D
Private Const COLOR_SELECT_FORE As Long = &H8000000E

Private mblnOK As Boolean

Private mstrLabType As String, mlngDefaultDevice As Long
Private mlngExeDeptID As Long

Public Function ShowEdit(ByVal frmMain As Form, ByVal strLabType As String, ByVal lngDeptID As Long, _
    Optional ByVal lngDefaultDevice As Long = -1) As Boolean
    mblnOK = False: ShowEdit = False
    
    mstrLabType = strLabType
    mlngDefaultDevice = lngDefaultDevice
    mlngExeDeptID = lngDeptID
    
    If Not InitData Then Exit Function
    
    Me.Show vbModal, frmMain
    
    ShowEdit = mblnOK
End Function

Private Function InitData() As Boolean
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Long
    
    InitData = False
    On Error GoTo DataError
    With Me.Vsf
        .Cols = 0
        NewColumn Vsf, "", 600, 9
        NewColumn Vsf, "��д", 900, 1
        NewColumn Vsf, "��Ŀָ��", 1800, 1
        .FixedCols = 1
    End With
    
    With Me.VsfSelect
        .Cols = 0
        NewColumn VsfSelect, "", 300, 9
        NewColumn VsfSelect, "", 300, 9
'        Set .Cell(flexcpPicture, 0, 1) = ils16.ListImages("���").Picture
'        NewColumn VsfSelect, "", 300, 9
'        Set .Cell(flexcpPicture, 0, 2) = ils16.ListImages("ָ��").Picture
        NewColumn VsfSelect, "��д", 900, 1
        NewColumn VsfSelect, "��Ŀָ��", 2000, 1
        NewColumn VsfSelect, "���ID", 0, 1
        .FixedCols = 1
    End With
    
    strSQL = "Select Distinct �������� As ���� From ������ĿĿ¼ A,����ִ�п��� B " & _
        "WHERE A.ID=B.������ĿID And A.���='C' And A.����Ӧ��=1 And B.ִ�п���ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngExeDeptID)
    If rsTmp.EOF Then Exit Function
    
    cboType.Clear
    Do While Not rsTmp.EOF
        cboType.AddItem Nvl(rsTmp(0))
    
        rsTmp.MoveNext
    Loop
    On Error Resume Next
    cboType.Text = mstrLabType
    If cboType.ListIndex = -1 Then cboType.ListIndex = 0
    On Error GoTo DataError
    
    Me.cbo��������.Clear
    Me.cbo��������.AddItem "": Me.cbo��������.ItemData(0) = 0
    strSQL = _
        " Select Distinct A.ID,A.����,A.����" & _
        " From ���ű� A,��������˵�� B " & _
        " Where B.����ID = A.ID " & _
        " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
        " And (B.�������� IN('�ٴ�','���'))" & _
        " Order by A.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        cbo��������.AddItem rsTmp!����
        cbo��������.ItemData(cbo��������.NewIndex) = rsTmp!ID
        
        rsTmp.MoveNext
    Next
    Me.cbo��������.ListIndex = -1: Me.cboҽ��.ListIndex = -1
    Me.txtComment = ""
    
    Me.DTPApply.Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    InitData = True
    Exit Function
DataError:
    If ErrCenter = 1 Then Resume
End Function

Private Sub GetLabItems(ByVal strLabType As String)
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim objItem As ListItem
    
    On Error GoTo DataError
    
    'û��ָ�������Ҫ
    If Not chkTest.Value = 1 Then
        strSQL = "Select Distinct A.ID,A.����, A.���� From ������ĿĿ¼ A,���鱨����Ŀ B,������Ŀ C,����ִ�п��� D " & _
            "Where A.ID=B.������ĿID And B.������ĿID=C.������ĿID And A.ID=D.������ĿID " & _
            "And ���='C' And ����Ӧ��=1 And C.��Ŀ��� In (1,3,4) And ��������=[1] And D.ִ�п���ID=[2]"
    Else
        strSQL = "Select Distinct A.ID,A.����, A.���� From ������ĿĿ¼ A,���鱨����Ŀ B,������Ŀ C,����ִ�п��� D " & _
            "Where A.ID=B.������ĿID And B.������ĿID=C.������ĿID And A.ID=D.������ĿID " & _
            "And ���='C' And ����Ӧ��=1 And C.��Ŀ���=2 And ��������=[1] And D.ִ�п���ID=[2]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strLabType, mlngExeDeptID)
    
    With lvwLabItem
        .ListItems.Clear
        ResetVsf Vsf
        
        Do While Not rsTmp.EOF
            Set objItem = .ListItems.Add(, "_" & rsTmp("ID"), rsTmp("����"))
            objItem.SubItems(.ColumnHeaders("����").Index - 1) = "" & rsTmp!����
        
            rsTmp.MoveNext
        Loop
        .ColumnHeaders("����").Position = 1
        .SortKey = .ColumnHeaders("����").Index - 1
        .SortOrder = lvwAscending
        If .ListItems.Count > 0 Then Set .SelectedItem = .ListItems(1)
    End With
    
    Exit Sub
DataError:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub GetDetails(ByVal lngItem As Long)
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo DataError
    
    strSQL = "Select D.ID, D.���� As ��Ŀָ��, B.��д From ���鱨����Ŀ A, ������Ŀ B, ���鱨����Ŀ C, ������ĿĿ¼ D " & _
        "Where A.������Ŀid = B.������Ŀid And B.������Ŀid = C.������ĿID And C.������Ŀid = D.ID And D.�����Ŀ = 0 " & _
        "And A.������Ŀid = [1] ORDER BY A.�������"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngItem)
    
    Vsf.TextMatrix(0, 0) = "#"
    FillGrid_UQ Vsf, rsTmp
    Vsf.TextMatrix(0, 0) = ""
    
    Vsf.Cell(flexcpChecked, 1, 0, Vsf.Rows - 1, 0) = 1
    Vsf.Cell(flexcpBackColor, 1, 1, Vsf.Rows - 1, Vsf.Cols - 1) = COLOR_SELECT_BACK
    Vsf.Cell(flexcpForeColor, 1, 1, Vsf.Rows - 1, Vsf.Cols - 1) = COLOR_SELECT_FORE
    Vsf.BackColorSel = COLOR_SELECT_BACK: Vsf.ForeColorSel = COLOR_SELECT_FORE
    
    Exit Sub
DataError:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cboDevice_Click()
    With cboDevice
        If .ItemData(.ListIndex) = -1 Then
            If Val(txt(0)) = 0 Then txt(0) = "0001"
            txt(0).Enabled = True
        Else
            txt(0) = "0001"
            txt(0).Enabled = False
        End If
    End With
End Sub

Private Sub cboType_Click()
    Call GetLabItems(cboType.Text)
    If Not lvwLabItem.SelectedItem Is Nothing Then GetDetails Val(Mid(lvwLabItem.SelectedItem.Key, 2))
    
    Call InitDevice(cboType.Text)
End Sub

Private Sub cbo��������_Click()
    If cbo��������.ListIndex > -1 Then InitDoctors cbo��������.ItemData(cbo��������.ListIndex)
End Sub

Private Sub cboҽ��_GotFocus()
    Call zlControl.TxtSelAll(cboҽ��)
End Sub

Private Sub cboҽ��_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
        
    If cboҽ��.ListIndex <> -1 Then Exit Sub '��ѡ��
    If cboҽ��.Text = "" Then '������
        Exit Sub
    End If
    
    strInput = UCase(NeedName(cboҽ��.Text))
    'ȫԺҽ��
    strSQL = "Select Distinct ����ID From ��������˵�� Where ������� IN(1,2,3)"
    strSQL = "Select Distinct A.ID,A.���,A.����,A.����" & _
        " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
        " Where A.ID=B.��ԱID And A.ID=C.��ԱID And C.��Ա����='ҽ��'" & _
        " And B.����ID IN(" & strSQL & ")" & _
        " And (Upper(A.���) Like [1] Or Upper(A.����) Like [2] Or Upper(A.����) Like [2])" & _
        " And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & _
        " Order by A.����"
    
    On Error GoTo errH
    vRect = GetControlRect(cboҽ��.hWnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����ҽ��", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cboҽ��.Height, blnCancel, False, True, strInput & "%", strInput & "%")
    If Not rsTmp Is Nothing Then
        cboҽ��.Text = rsTmp!����
    Else
        If Not blnCancel Then
            MsgBox "δ�ҵ���Ӧ��ҽ����", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub chkTest_Click()
    ResetVsf VsfSelect

    Call GetLabItems(cboType.Text)
    If Not lvwLabItem.SelectedItem Is Nothing Then GetDetails Val(Mid(lvwLabItem.SelectedItem.Key, 2))
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim i As Long
    
    With VsfSelect
        If .Row = 0 Then Exit Sub
        If Val(.RowData(.Row)) = 0 Then Exit Sub
        
        .RemoveItem .Row
        For i = 1 To .Rows - 1
            .TextMatrix(i, 0) = i
        Next
    End With
    Me.VsfSelect.SetFocus
End Sub

Private Sub cmdOK_Click()
    If VsfSelect.Rows < 2 Then Exit Sub
    If Val(VsfSelect.RowData(1)) = 0 Then Exit Sub
    
    If Val(txt(0)) <= 0 Then
        MsgBox "���ű������0", vbInformation + vbOKOnly, gstrSysName
        txt(0).SetFocus
        Exit Sub
    End If
    If Me.optType(0).Value And Val(txt(2)) < Val(txt(1)) Then
        MsgBox "�����걾�Ų���С����ʼ�걾��", vbInformation + vbOKOnly, gstrSysName
        txt(2).SetFocus
        Exit Sub
    End If
    If Me.optType(1).Value And Val(txt(3)) <= 0 Then
        MsgBox "�걾�����������0", vbInformation + vbOKOnly, gstrSysName
        txt(3).SetFocus
        Exit Sub
    End If
    
    If SaveData Then
        mblnOK = True
        Unload Me
    End If
End Sub

Private Sub cmdSelect_Click()
    Dim blnAll As Boolean, i As Long, lngRow As Long
    
    If Me.lvwLabItem.SelectedItem Is Nothing Then Exit Sub
    '����������
    If FindGridLine(VsfSelect, Val(Mid(Me.lvwLabItem.SelectedItem.Key, 2))) > 0 Then Exit Sub
    If FindItem(VsfSelect, Mid(Me.lvwLabItem.SelectedItem.Key, 2), 4) > 0 Then
    '����ָ�����
        blnAll = False
    Else
        blnAll = True
        For i = 1 To Vsf.Rows - 1
            If Not Val(Vsf.Cell(flexcpChecked, i, 0)) = 1 And Val(Vsf.RowData(i)) > 0 Then
                blnAll = False
                Exit For
            End If
        Next
    End If
    
    If Val(VsfSelect.RowData(VsfSelect.Rows - 1)) > 0 Or VsfSelect.Rows = 1 Then VsfSelect.Rows = VsfSelect.Rows + 1
    If blnAll Then
        If FindGridLine(VsfSelect, Val(Mid(Me.lvwLabItem.SelectedItem.Key, 2))) > 0 Then Exit Sub
        With VsfSelect
            lngRow = .Rows - 1
            .RowData(lngRow) = Val(Mid(Me.lvwLabItem.SelectedItem.Key, 2))
            .TextMatrix(lngRow, 0) = lngRow
            .Cell(flexcpPicture, lngRow, 1) = ils16.ListImages("���").Picture
            
            .TextMatrix(lngRow, 3) = Me.lvwLabItem.SelectedItem.Text
            .TextMatrix(lngRow, 4) = Val(.RowData(lngRow))
        End With
    Else
        With VsfSelect
            For i = 1 To Vsf.Rows - 1
                If Val(Vsf.Cell(flexcpChecked, i, 0)) = 1 And Val(Vsf.RowData(i)) > 0 Then
                    If FindGridLine(VsfSelect, Vsf.RowData(i)) = -1 Then
                        lngRow = .Rows - 1
                        .RowData(lngRow) = Val(Vsf.RowData(i))
                        .TextMatrix(lngRow, 0) = lngRow
                        .Cell(flexcpPicture, lngRow, 1) = ils16.ListImages("ָ��").Picture
                        .TextMatrix(lngRow, 2) = Vsf.TextMatrix(i, 1)
                        .TextMatrix(lngRow, 3) = Vsf.TextMatrix(i, 2)
                        .TextMatrix(lngRow, 4) = Val(Mid(Me.lvwLabItem.SelectedItem.Key, 2))
                        
                        VsfSelect.Rows = VsfSelect.Rows + 1
                    End If
                End If
            Next
        End With
    End If
    Me.VsfSelect.SetFocus
End Sub

Private Function FindItem(objMsf As Object, ByVal strSeek As String, ByVal FindCol As Long) As Long
    '-------------------------------------------------------------------------------------------------------------
    '����:����
    '����:
    '����:�кŻ�-1
    '-------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    FindItem = -1
    For i = 1 To objMsf.Rows - 1
        If objMsf.TextMatrix(i, FindCol) = strSeek Then Exit For
    Next
    If i <= objMsf.Rows - 1 Then FindItem = i
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub lvwLabItem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If lvwLabItem.SortKey = ColumnHeader.Index - 1 Then
        lvwLabItem.SortOrder = IIf(lvwLabItem.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        lvwLabItem.SortKey = ColumnHeader.Index - 1
        lvwLabItem.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwLabItem_ItemClick(ByVal Item As MSComctlLib.ListItem)
    GetDetails Val(Mid(Item.Key, 2))
End Sub

Private Sub optType_Click(Index As Integer)
    If Index = 0 Then
        txt(2).Enabled = True: txt(3).Enabled = False: Me.txt(2).SetFocus
    Else
        txt(2).Enabled = False: txt(3).Enabled = True: Me.txt(3).SetFocus
    End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
    With txt(Index)
        .SelStart = 0
        .SelLength = Len(txt(Index))
    End With
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = FilterKeyAscii(KeyAscii, 1)
End Sub

Private Sub txtComment_GotFocus()
    With txtComment
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtComment_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtComment_Validate(Cancel As Boolean)
    Cancel = False
    
    txtComment.Text = GetComment(txtComment.Text, Me.cboType.Text)
End Sub

Private Sub Vsf_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If Val(Vsf.Cell(flexcpChecked, NewRow, 0)) = 1 Then
        Vsf.BackColorSel = COLOR_SELECT_BACK: Vsf.ForeColorSel = COLOR_SELECT_FORE
    Else
        Vsf.BackColorSel = Vsf.BackColor: Vsf.ForeColorSel = Vsf.ForeColor
    End If
End Sub

Private Sub Vsf_Click()
    ShiftSelect
End Sub

Private Sub Vsf_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
        ShiftSelect
    End If
End Sub

Private Sub ShiftSelect()
    If Val(Vsf.Cell(flexcpChecked, Vsf.Row, 0)) = 1 Then
        Vsf.Cell(flexcpChecked, Vsf.Row, 0) = 0
        Vsf.Cell(flexcpBackColor, Vsf.Row, 1, Vsf.Row, Vsf.Cols - 1) = Vsf.BackColor
        Vsf.Cell(flexcpForeColor, Vsf.Row, 1, Vsf.Row, Vsf.Cols - 1) = Vsf.ForeColor
        
        Vsf.BackColorSel = Vsf.BackColor: Vsf.ForeColorSel = Vsf.ForeColor
    Else
        Vsf.Cell(flexcpChecked, Vsf.Row, 0) = 1
        Vsf.Cell(flexcpBackColor, Vsf.Row, 1, Vsf.Row, Vsf.Cols - 1) = COLOR_SELECT_BACK
        Vsf.Cell(flexcpForeColor, Vsf.Row, 1, Vsf.Row, Vsf.Cols - 1) = COLOR_SELECT_FORE
    
        Vsf.BackColorSel = COLOR_SELECT_BACK: Vsf.ForeColorSel = COLOR_SELECT_FORE
    End If
End Sub

Private Sub VsfSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cmdDelete_Click
End Sub

Private Sub InitDevice(ByVal strLabType As String)
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strOldDevice As String
    
    On Error GoTo DataError
    strSQL = "Select Distinct D.ID, D.����" & vbNewLine & _
        " From ������ĿĿ¼ A, ���鱨����Ŀ B, ����������Ŀ C, �������� D, ����ִ�п��� E" & vbNewLine & _
        " Where A.ID = B.������Ŀid And B.������Ŀid = C.��Ŀid And C.����id = D.ID And A.ID = E.������Ŀid And A.��� = 'C' And A.�������� = [1] And " & vbNewLine & _
        "      E.ִ�п���id = [2] And " & vbNewLine & _
        "      D.ID In (Select Distinct D.ID" & vbNewLine & _
        "               From ����С���Ա A, ����С�� B, ����С������ C, �������� D" & vbNewLine & _
        "               Where A.С��id = B.ID And B.ID = C.С��id��and ��Աid = [3] And C.����id = D.ID And C.���� = 1)"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strLabType, mlngExeDeptID, UserInfo.ID)
    
    With Me.cboDevice
        strOldDevice = .Text
        
        .Clear
        .AddItem "�ֹ�"
        .ItemData(.NewIndex) = -1
        
        Do While Not rsTmp.EOF
            .AddItem rsTmp("����")
            .ItemData(.NewIndex) = rsTmp("ID")
            
            rsTmp.MoveNext
        Loop
        
        On Error Resume Next
        If strOldDevice = "" Then
            If mlngDefaultDevice > 0 Then
                .ListIndex = FindComboItem(cboDevice, mlngDefaultDevice)
            Else
                .ListIndex = 0
            End If
        Else
            .Text = strOldDevice
        End If
        If .ListIndex = -1 Then .ListIndex = 0
    End With
    
    Exit Sub
DataError:
    If ErrCenter = 1 Then Resume
End Sub

Private Function SaveData() As Boolean
    Dim i As Long, lngSampleNum As Long, strSampleNO As Long, lngSampleID As Long
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strDate As String, lngDeviceID As Long
    
    Dim strItemIDs As String, strItemResults As String
    Dim blnTrans As Boolean
    Dim lngOrderDept As Long
    
    On Error GoTo DataError
    
    SaveData = True
    blnTrans = False
    
    If optType(1).Value Then
        lngSampleNum = Val(txt(3))
    Else
        lngSampleNum = Val(txt(2)) - Val(txt(1)) + 1
    End If
    
    strDate = Me.DTPApply.Value
    lngDeviceID = cboDevice.ItemData(cboDevice.ListIndex)
    '������
    strItemIDs = ""
    For i = 1 To VsfSelect.Rows - 1
        If Val(VsfSelect.RowData(i)) > 0 Then
            strItemIDs = strItemIDs & "," & Val(VsfSelect.RowData(i))
        End If
    Next
    If strItemIDs = "" Then
        SaveData = False
        Exit Function
    Else
        strItemIDs = Mid(strItemIDs, 2)
    End If
    
    If Not Me.chkTest.Value = 1 Then
'        strsql = "SELECT D.������ĿID As ID,'' As ԭʼ���,Decode(C.�������,3,Nvl(C.Ĭ��ֵ,'-'),2,C.Ĭ��ֵ,'') As ������," & _
            "'' AS �����־," & _
            "Trim(REPLACE(REPLACE(' '||zlGetReference(D.������ĿID,A.�걾��λ,0,NULL,[1]),' .','0.'),'��.','��0.')) AS ����ο� " & _
            "FROM ������ĿĿ¼ A,���鱨����Ŀ D,������Ŀ C " & _
            "WHERE A.ID = D.������ĿID And D.������ĿID = C.������ĿID " & _
                        "AND D.ϸ��ID IS NULL AND C.��Ŀ���<>2 " & _
                        "AND A.ID In (" & strItemIDs & ") Order By A.ID,D.�������"
        strSQL = "Select Id, ԭʼ���, ������, �����־, Rownum As �������, ������Ŀid,����ο�" & vbNewLine & _
                    "From (Select d.������Ŀid As Id, '' As ԭʼ���, Decode(c.�������, 3, Nvl(c.Ĭ��ֵ, '-'), 2, c.Ĭ��ֵ, '') As ������," & vbNewLine & _
                    "                           '' As �����־, d.�������, a.Id As ������Ŀid," & vbNewLine & _
                    "                           Trim(Replace(Replace(' ' || Zlgetreference(d.������Ŀid, a.�걾��λ, 0, Null), ' .', '0.'), '��.', '��0.')) As ����ο�" & vbNewLine & _
                    "            From ������ĿĿ¼ a, ���鱨����Ŀ d, ������Ŀ c" & vbNewLine & _
                    "            Where a.Id = d.������Ŀid And d.������Ŀid = c.������Ŀid And d.ϸ��id Is Null And c.��Ŀ��� <> 2 And" & vbNewLine & _
                    "                        a.Id In (Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & vbNewLine & _
                    "            Order By a.����, d.�������)"
                        
    Else
        '΢����ҩ��
'        strsql = "SELECT D.ϸ��ID As ID,'' As ԭʼ���,'' As ������,'' As �����־,'' As ����ο� " & _
            "FROM ������ĿĿ¼ A,���鱨����Ŀ D,����ϸ�� C " & _
            "WHERE A.ID = D.������ĿID And D.ϸ��ID = C.ID AND A.ID In (" & strItemIDs & ") Order By A.����,D.�������"
        strSQL = "Select Id, ԭʼ���, ������, �����־, ����ο�, Rownum As �������, ������Ŀid" & vbNewLine & _
                    "From (Select d.ϸ��id As Id, '' As ԭʼ���, C.Ĭ�Ͻ�� As ������, '' As �����־, '' As ����ο�, d.�������," & vbNewLine & _
                    "                           a.Id As ������Ŀid" & vbNewLine & _
                    "            From ������ĿĿ¼ a, ���鱨����Ŀ d, ����ϸ�� c" & vbNewLine & _
                    "            Where a.Id = d.������Ŀid And d.ϸ��id = c.Id And a.Id In (Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & vbNewLine & _
                    "            Order By a.����, d.�������)"
            
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strItemIDs)
    strItemResults = ""
    Do While Not rsTmp.EOF
        strItemResults = strItemResults & "|" & "^" & Nvl(rsTmp("ID")) & "^" & Nvl(rsTmp("������")) & _
            "^" & Nvl(rsTmp("�����־")) & "^" & Nvl(rsTmp("����ο�")) & "^" & Nvl(rsTmp("������ĿID")) & _
            "^" & Nvl(rsTmp("�������"))
        rsTmp.MoveNext
    Loop
    If strItemResults = "" And Not Me.chkTest.Value = 1 Then
        SaveData = False
        MsgBox "��ѡ����Ŀû��ָ�꣬���ܲ����걾", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If lngDeviceID = -1 Then
        strSQL = "Select ID From ����걾��¼ a " & _
            " Where a.����ʱ�� Between [1] And [2]" & _
            " And a.����ID Is NULL And a.�걾��� Between [4] And [5] And Nvl(a.�걾���,0)=[6]"
    Else
        strSQL = "Select ID From ����걾��¼ a " & _
            " Where a.����ʱ�� Between [1] And [2]" & _
            " And a.����ID=[3] And a.�걾��� Between [4] And [5] And Nvl(a.�걾���,0)=[6]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Format(strDate, "yyyy-MM-dd") & " 00:00:00"), _
        CDate(Format(strDate, "yyyy-MM-dd") & " 23:59:59"), lngDeviceID, _
        TransSampleNO(Val(txt(0)) & "-" & (Val(txt(1)) + i - 1)), _
        CStr(TransSampleNO(Val(txt(0)) & "-" & (Val(txt(1)) + i - 1)) + lngSampleNum), 0)
    If Not rsTmp.EOF Then
        If MsgBox("���β����Ĳ��ֱ걾�Ѿ����ڣ��Ƿ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            SaveData = False
            Exit Function
        End If
    End If
    
    If Len(strItemResults) > 0 Then strItemResults = Mid(strItemResults, 2)
    
    With PgbSave
        .Max = lngSampleNum
        .Value = 0
        .Visible = True
    End With
    Me.MousePointer = vbHourglass
    For i = 1 To lngSampleNum
        'תΪ��������ʽ
        strSampleNO = TransSampleNO(Val(txt(0)) & "-" & (Val(txt(1)) + i - 1))
        
'        gcnOracle.BeginTrans
'        blnTrans = True
        
        '�ж��Ƿ������걾
        If lngDeviceID = -1 Then
            strSQL = "Select ID From ����걾��¼ a " & _
                " Where a.����ʱ�� Between [1] And [2]" & _
                " And a.����ID Is NULL And a.�걾���=[4] And Nvl(a.�걾���,0)=[5]"
        Else
            strSQL = "Select ID From ����걾��¼ a " & _
                " Where a.����ʱ�� Between [1] And [2]" & _
                " And a.����ID=[3] And a.�걾���=[4] And Nvl(a.�걾���,0)=[5]"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Format(strDate, "yyyy-MM-dd") & " 00:00:00"), _
            CDate(Format(strDate, "yyyy-MM-dd") & " 23:59:59"), lngDeviceID, strSampleNO, 0)
        If rsTmp.EOF Then
            If Me.cbo��������.ListIndex = -1 Then
                lngOrderDept = 0
            Else
                lngOrderDept = Me.cbo��������.ItemData(Me.cbo��������.ListIndex)
            End If
            '�����걾������ʱ�걾��¼
            lngSampleID = zlDatabase.GetNextId("����걾��¼")
            strSQL = "ZL_����걾��¼_INSERT(" & lngSampleID & ",NULL,'" & _
                strSampleNO & "',NULL,NULL," & IIf(lngDeviceID = -1, "NULL", lngDeviceID) & ",NULL," & _
                "To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),NULL," & _
                "To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),''," & _
                "Null,To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),'',''," & mlngExeDeptID & ",0," & _
                IIf(chkTest.Value, 1, "Null") & ",'" & strItemIDs & "','" & txtComment & "'," & IIf(lngOrderDept = 0, "NULL", lngOrderDept) & _
                ",'" & Me.cboҽ�� & "')"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        Else
            lngSampleID = rsTmp("ID")
        End If
        '�ύ���
        strSQL = "Zl_������ͨ���_Write(" & lngSampleID & "," & IIf(lngDeviceID = -1, "NULL", lngDeviceID) & _
            ",'" & strItemResults & "'," & IIf(chkOverWrite.Value, 1, 0) & "," & IIf(chkTest.Value, 1, 0) & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        strSQL = "Zl_���¼�����_Cale(" & lngSampleID & ")"
'        gcnOracle.CommitTrans
        blnTrans = False
        PgbSave.Value = i
    Next
    
    PgbSave.Visible = False
    Me.MousePointer = vbDefault
    Exit Function
DataError:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then gcnOracle.RollbackTrans
    
    PgbSave.Visible = False
    Me.MousePointer = vbDefault
    
    SaveData = False
End Function

Private Sub InitDoctors(ByVal lng����ID As Long)
'���ܣ���ȡ��ǰ���������а�����������Ա
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    Me.cboҽ��.Clear
    
    '����ҽ����ʿ
    strSQL = _
        "Select Distinct A.ID,B.����ID,A.���,A.����,Upper(A.����) as ����," & _
        " C.��Ա����,Nvl(A.Ƹ�μ���ְ��,0) as ְ��" & _
        " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
        " Where A.ID=B.��ԱID And A.ID=C.��ԱID" & _
        " And C.��Ա���� IN('ҽ��') And B.����ID=[1] " & _
        " and (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) "
    strSQL = strSQL & " Order by ����,��Ա���� Desc"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboҽ��.AddItem rsTmp!����
            cboҽ��.ItemData(cboҽ��.ListCount - 1) = rsTmp!����ID
            
            If rsTmp!ID = UserInfo.ID And cboҽ��.ListIndex = -1 Then cboҽ��.ListIndex = cboҽ��.NewIndex
            rsTmp.MoveNext
        Next
        
        If cboҽ��.ListCount = 1 And cboҽ��.ListIndex = -1 Then cboҽ��.ListIndex = 0
    End If
End Sub
'ͨ�������ȡ���鱸ע
Private Function GetComment(ByVal strCode As String, ByVal strTYPE As String)
    Dim rsTmp As ADODB.Recordset, mstrSQL As String
    Dim objPoint As POINTAPI
    Dim sglX As Single, sglY As Single
    
    mstrSQL = "SELECT Rownum As ID,A.����,A.����,A.����,A.˵�� As ���� FROM ���鱸ע���� A " & _
        "WHERE (Instr(A.����,[1])>0 Or Instr(A.����,[1])>0) And (A.���� Is Null Or A.����=[2])"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, UCase(strCode), strTYPE)
    If rsTmp.EOF Then
        GetComment = strCode
    Else
        If rsTmp.RecordCount = 1 Then
            GetComment = Nvl(rsTmp("����"))
        Else
            Call ClientToScreen(txtComment.hWnd, objPoint)
    
            sglX = objPoint.x * 15 - 30
            sglY = objPoint.Y * 15 - 2000
            If frmSelectList.ShowSelect(Me, rsTmp, "����,800,0,0;����,1500,0,0;����,2500,0,0;����,5500,0,0", sglX, sglY, Me.txtComment.Width, 2000, Me.Name & "\���鱸עѡ��", "��ѡ����鱸ע") Then
                GetComment = Nvl(rsTmp("����"))
            Else
                GetComment = strCode
            End If
        End If
    End If
End Function
