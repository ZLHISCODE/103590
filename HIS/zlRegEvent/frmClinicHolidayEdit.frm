VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClinicHolidayEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ڼ�������"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9540
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClinicHolidayEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6405
      Left            =   7920
      TabIndex        =   24
      Top             =   -150
      Width           =   15
   End
   Begin VB.TextBox txtComment 
      Height          =   1305
      Left            =   930
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   4770
      Width           =   6795
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfRegistInfo 
      Height          =   1845
      Left            =   930
      TabIndex        =   10
      Top             =   960
      Width           =   6795
      _cx             =   11986
      _cy             =   3254
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483638
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   8
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmClinicHolidayEdit.frx":000C
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
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ɾ��(&D)"
      Enabled         =   0   'False
      Height          =   320
      Left            =   6840
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2970
      Width           =   885
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "����(&A)"
      Height          =   320
      Left            =   5940
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2970
      Width           =   885
   End
   Begin VB.ComboBox cboYear 
      Height          =   330
      Left            =   930
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   150
      Width           =   1485
   End
   Begin VB.ComboBox cboHolidayName 
      Height          =   330
      Left            =   5070
      TabIndex        =   3
      Top             =   150
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   360
      Left            =   8190
      TabIndex        =   22
      Top             =   780
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   360
      Left            =   8190
      TabIndex        =   21
      Top             =   300
      Width           =   1095
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   360
      Left            =   8190
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5520
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dtpEndTime 
      Height          =   330
      Left            =   6510
      TabIndex        =   9
      Top             =   585
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "HH:mm:ss"
      Format          =   169738243
      UpDown          =   -1  'True
      CurrentDate     =   42320.9999884259
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   330
      Left            =   5070
      TabIndex        =   8
      Top             =   585
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   -2147483630
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   169738243
      CurrentDate     =   42320
   End
   Begin MSComCtl2.DTPicker dtpOldWorkDate 
      Height          =   330
      Left            =   2040
      TabIndex        =   13
      Top             =   2970
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   -2147483630
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   169738243
      CurrentDate     =   42320
   End
   Begin MSComCtl2.DTPicker dtpNewWorkDate 
      Height          =   330
      Left            =   4500
      TabIndex        =   15
      Top             =   2970
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   -2147483630
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   169738243
      CurrentDate     =   42320
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   330
      Left            =   930
      TabIndex        =   5
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   -2147483630
      CalendarTitleForeColor=   -2147483628
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   169738243
      CurrentDate     =   42320
   End
   Begin MSComCtl2.DTPicker dtpStartTime 
      Height          =   330
      Left            =   2370
      TabIndex        =   6
      Top             =   600
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "HH:mm:ss"
      Format          =   169738243
      UpDown          =   -1  'True
      CurrentDate     =   42320
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf������� 
      Height          =   1275
      Left            =   960
      TabIndex        =   18
      Top             =   3330
      Width           =   6765
      _cx             =   11933
      _cy             =   2249
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483638
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmClinicHolidayEdit.frx":00CA
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
   Begin VB.Label lblHolidyName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "�ڼ���"
      Height          =   210
      Left            =   4410
      TabIndex        =   2
      Top             =   210
      Width           =   630
   End
   Begin VB.Label lblStartTime 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "��ʼʱ��"
      Height          =   210
      Left            =   60
      TabIndex        =   4
      Top             =   645
      Width           =   840
   End
   Begin VB.Label lblEndTime 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "����ʱ��"
      Height          =   210
      Left            =   4200
      TabIndex        =   7
      Top             =   630
      Width           =   840
   End
   Begin VB.Label lblComment 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "����˵��"
      Height          =   240
      Left            =   60
      TabIndex        =   19
      Top             =   4770
      Width           =   840
   End
   Begin VB.Label lblNewWorkDate 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   210
      Left            =   3630
      TabIndex        =   14
      Top             =   3015
      Width           =   840
   End
   Begin VB.Label lblOldWorkTime 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ԭ�ϰ�����"
      Height          =   210
      Left            =   960
      TabIndex        =   12
      Top             =   3015
      Width           =   1050
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   915
      X2              =   915
      Y1              =   3015
      Y2              =   4600
   End
   Begin VB.Label lbl������Ϣ 
      AutoSize        =   -1  'True
      Caption         =   "������Ϣ"
      Height          =   210
      Left            =   60
      TabIndex        =   11
      Top             =   3015
      Width           =   840
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblYear 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "���"
      Height          =   210
      Left            =   480
      TabIndex        =   0
      Top             =   210
      Width           =   420
   End
End
Attribute VB_Name = "frmClinicHolidayEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytFun As G_Enum_Fun '0-�鿴,1-���,2-����
Private mlngYear As Long
Private mstrHolidayName As String

Private Enum mGridHeadCol
    COL_��� = 0
    COL_ԭ�ϰ�ʱ�� = 1
    Col_����ʱ�� = 2
    
    COL_���� = 0
    COL_����Һ� = 1
    COL_����ԤԼ = 2
End Enum
Private mblnOK As Boolean
Private mblnNotClick As Boolean
Private mrsDefautHoliday As ADODB.Recordset
Private mstr����ԤԼ As String '��ʽ��yyyy-mm-dd;yyyy-mm-dd;...
Private mstr����Һ� As String '��ʽ��yyyy-mm-dd;yyyy-mm-dd;...

Public Function ShowMe(frmParent As Form, ByVal bytFun As G_Enum_Fun, _
    Optional ByVal lngYear As Long, Optional ByVal strHolidayName As String) As Boolean
    '��Σ�
    '   frmParent - ������
    '   bytFun - ��������, 0-�鿴��1-������2-�޸�
    mbytFun = bytFun
    mlngYear = lngYear: mstrHolidayName = strHolidayName
    
    Err = 0: On Error Resume Next
    mblnOK = False
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Sub cboHolidayName_Click()
    Err = 0: On Error GoTo ErrHandler
    If mblnNotClick Then Exit Sub
    LoadData cboHolidayName.Text
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cboHolidayName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboHolidayName_Validate(Cancel As Boolean)
    If zlCommFun.ActualLen(cboHolidayName.Text) > 50 Then
        MsgBox "�ڼ�������ֻ��������50���ַ���25�����֣�", vbInformation, gstrSysName
        zlControl.TxtSelAll cboHolidayName
        Cancel = True
    End If
End Sub

Private Sub cboYear_Click()
    Dim lngYear As Long
    
    Err = 0: On Error GoTo ErrHandler
    If mbytFun = Fun_View Then Exit Sub
    lngYear = Val(cboYear.Text)
    dtpStartDate.MaxDate = "9999-12-31"
    dtpStartDate.MinDate = lngYear & "-01-01": dtpStartDate.MaxDate = lngYear & "-12-31"
    If dtpStartDate.MinDate < DateAdd("d", 1, Format(Now, "yyyy-mm-dd")) Then dtpStartDate.MinDate = DateAdd("d", 1, Format(Now, "yyyy-mm-dd"))
    dtpStartDate.Value = dtpStartDate.MinDate
    
    dtpEndDate.MaxDate = "9999-12-31"
    dtpEndDate.MinDate = dtpStartDate.MinDate: dtpEndDate.MaxDate = lngYear & "-12-31"
    dtpEndDate.Value = dtpStartDate.Value
    
    dtpOldWorkDate.MaxDate = "9999-12-31"
    dtpOldWorkDate.MinDate = dtpStartDate.MinDate: dtpOldWorkDate.MaxDate = lngYear & "-12-31"
    dtpOldWorkDate.Value = dtpStartDate.Value
    
    dtpNewWorkDate.MaxDate = "9999-12-31"
    dtpNewWorkDate.MinDate = dtpStartDate.MinDate: dtpNewWorkDate.MaxDate = lngYear + 1 & "-01-31"
    dtpNewWorkDate.Value = dtpStartDate.Value
    Call ShowDateRangeToGrid(dtpStartDate.Value, dtpEndDate.Value)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cboYear_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub dtpEndDate_Change()
    Call ShowDateRangeToGrid(dtpStartDate.Value, dtpEndDate.Value)
End Sub

Private Sub dtpEndDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpEndTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpNewWorkDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpOldWorkDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpStartDate_Change()
    dtpEndDate.Value = dtpStartDate.Value
    dtpOldWorkDate.Value = dtpStartDate.Value
    dtpNewWorkDate.Value = dtpStartDate.Value
    Call ShowDateRangeToGrid(dtpStartDate.Value, dtpEndDate.Value)
End Sub

Private Sub dtpStartDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpStartTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    Dim varRow As Variant, i As Long, lngYear As Long
    Dim rs�ڼ��� As ADODB.Recordset, strSQL As String
    Dim varArray As Variant
    
    Err = 0: On Error GoTo ErrHandler
    mstr����ԤԼ = "": mstr����Һ� = ""
    Call InitGridHead
    cboYear.Clear
    For i = Year(Now) To Year(Now) + 10 'ȱʡ����10�깩ѡ��
        cboYear.AddItem i & "��"
    Next
    cboYear.ListIndex = 0
    
    varArray = Array("Ԫ����", "����", "��Ů��", "������", "�Ͷ���", "�����", "�����", "�����")
    cboHolidayName.Clear
    For i = 0 To UBound(varArray)
        cboHolidayName.AddItem varArray(i)
    Next
    If mbytFun = Fun_Add Or mbytFun = Fun_Update Then
        Call InitDefautHoliday
    End If
    Me.Caption = Choose(mbytFun + 1, "�鿴", "����", "�޸�", "ɾ��") & "�ڼ���"
    Call ShowDateRangeToGrid(dtpStartDate.Value, dtpEndDate.Value)
    
    If mbytFun = Fun_Add Then Exit Sub
    Select Case mbytFun
    Case Fun_View
        cmdCancel.Visible = False
        cmdOk.Left = cmdCancel.Left
        Call SetEnabled(Me.Controls, False)
    Case Fun_Update
        cboYear.Enabled = False
        cboHolidayName.Enabled = False
    End Select
    If LoadData(mstrHolidayName, mlngYear) = False Then Unload Me: Exit Sub
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function LoadData(ByVal strHolidayName As String, Optional ByVal lngYear As Long) As Boolean
    '��������
    'lngYear=0��ʾ���ýڼ���ȱʡʱ��
    Dim varRow As Variant, lngRow As Long, blnDefaut As Boolean
    Dim rs�ڼ��� As ADODB.Recordset, strSQL As String
    
    Err = 0: On Error GoTo ErrHandler
    blnDefaut = lngYear = 0
    strSQL = "Select ���,��������,��ʼ����,��ֹ����,��ע,����ԤԼ����,����Һ����� From �������ձ�" & vbNewLine & _
            " Where Nvl(����,0)=0 And ��������=[1]" & IIf(lngYear = 0, "", " And ���=[2]") & vbNewLine & _
            " Order By ��� Desc"
    Set rs�ڼ��� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strHolidayName, lngYear)
    
    If rs�ڼ���.EOF And blnDefaut Then '���ýڼ���ȱʡʱ��
        Set rs�ڼ��� = mrsDefautHoliday.Clone
        rs�ڼ���.Filter = "��������='" & strHolidayName & "'"
    End If
    If rs�ڼ���.RecordCount = 0 Then Exit Function
    
    If blnDefaut = False Then
        '���
        zlControl.CboSetText cboYear, lngYear & "��"
        
        If cboYear.Text = "" Then
            cboYear.AddItem lngYear & "��"
            cboYear.ListIndex = cboYear.NewIndex
        End If
    End If
    
    mstr����ԤԼ = Nvl(rs�ڼ���!����ԤԼ����)
    mstr����Һ� = Nvl(rs�ڼ���!����Һ�����)
    If lngYear <> 0 Then
        If Nvl(rs�ڼ���!��ʼ����) >= dtpStartDate.MinDate And Nvl(rs�ڼ���!��ʼ����) <= dtpStartDate.MaxDate Then
            dtpStartDate.Value = Format(Nvl(rs�ڼ���!��ʼ����), "yyyy-mm-dd")
        End If
        If Nvl(rs�ڼ���!��ֹ����) >= dtpEndDate.MinDate And Nvl(rs�ڼ���!��ֹ����) <= dtpEndDate.MaxDate Then
            dtpEndDate.Value = Format(Nvl(rs�ڼ���!��ֹ����), "yyyy-mm-dd")
        End If
    Else
        If Val(cboYear.Text) & Format(Nvl(rs�ڼ���!��ʼ����), "-mm-dd") >= dtpStartDate.MinDate _
            And Val(cboYear.Text) & Format(Nvl(rs�ڼ���!��ʼ����), "-mm-dd") <= dtpStartDate.MaxDate Then
            dtpStartDate.Value = Val(cboYear.Text) & Format(Nvl(rs�ڼ���!��ʼ����), "-mm-dd")
        End If
        If Val(cboYear.Text) & Format(Nvl(rs�ڼ���!��ֹ����), "-mm-dd") >= dtpEndDate.MinDate _
            And Val(cboYear.Text) & Format(Nvl(rs�ڼ���!��ֹ����), "-mm-dd") <= dtpEndDate.MaxDate Then
            dtpEndDate.Value = Val(cboYear.Text) & Format(Nvl(rs�ڼ���!��ֹ����), "-mm-dd")
        End If
    End If
    Call ShowDateRangeToGrid(dtpStartDate.Value, dtpEndDate.Value)
    dtpStartTime.Value = Format(Nvl(rs�ڼ���!��ʼ����), "hh:mm:ss")
    dtpEndTime.Value = Format(Nvl(rs�ڼ���!��ֹ����), "hh:mm:ss")
    dtpOldWorkDate.Value = dtpStartDate.Value
    dtpNewWorkDate.Value = dtpStartDate.Value
    If blnDefaut Then LoadData = True: Exit Function
    
    '��������
    mblnNotClick = True
    zlControl.CboSetText cboHolidayName, strHolidayName
    mblnNotClick = False
    If cboHolidayName.Text = "" Then
        cboHolidayName.AddItem strHolidayName
        mblnNotClick = True
        cboHolidayName.ListIndex = cboHolidayName.NewIndex
        mblnNotClick = False
    End If
    txtComment.Text = Nvl(rs�ڼ���!��ע)
    
    '�������
    strSQL = "Select ���,��������,��ʼ����,��ֹ����,��ע From �������ձ�" & vbNewLine & _
            " Where Nvl(����,0)=1 And ��������=[1] And ���=[2]"
    Set rs�ڼ��� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strHolidayName, lngYear)
    vsf�������.Rows = rs�ڼ���.RecordCount + 1
    lngRow = 1
    Do While Not rs�ڼ���.EOF
        vsf�������.TextMatrix(lngRow, COL_���) = lngRow
        vsf�������.TextMatrix(lngRow, COL_ԭ�ϰ�ʱ��) = Format(Nvl(rs�ڼ���!��ֹ����), "yyyy-mm-dd")
        vsf�������.TextMatrix(lngRow, Col_����ʱ��) = Format(Nvl(rs�ڼ���!��ʼ����), "yyyy-mm-dd")
        lngRow = lngRow + 1
        rs�ڼ���.MoveNext
    Loop
    LoadData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdAdd_Click()
    Dim i As Long
    
    Err = 0: On Error GoTo ErrHandler
    If Format(dtpStartDate.Value, "yyyy-mm-dd") & Format(dtpStartTime.Value, "hh:mm:ss") _
        >= Format(dtpEndDate.Value, "yyyy-mm-dd") & Format(dtpEndTime.Value, "hh:mm:ss") Then
        MsgBox "�ڼ��յ���ֹʱ�������ڿ�ʼʱ�䣡", vbInformation, gstrSysName
        If dtpEndDate.Visible And dtpEndDate.Enabled Then dtpEndDate.SetFocus
        Exit Sub
    End If
    If Format(dtpOldWorkDate.Value, "yyyy-mm-dd") = Format(dtpNewWorkDate.Value, "yyyy-mm-dd") Then
        MsgBox "����ʱ���ԭ�ϰ�ʱ�䲻��Ϊͬһ�죡", vbInformation, gstrSysName
        If dtpNewWorkDate.Visible And dtpNewWorkDate.Enabled Then dtpNewWorkDate.SetFocus
        Exit Sub
    End If
    If Format(dtpOldWorkDate.Value, "yyyy-mm-dd") < Format(dtpStartDate.Value, "yyyy-mm-dd") Or _
        Format(dtpOldWorkDate.Value, "yyyy-mm-dd") > Format(dtpEndDate.Value, "yyyy-mm-dd") Then
        MsgBox "ԭ�ϰ�ʱ������ڽڼ���ʱ�䷶Χ�ڣ�", vbInformation, gstrSysName
        If dtpOldWorkDate.Visible And dtpOldWorkDate.Enabled Then dtpOldWorkDate.SetFocus
        Exit Sub
    End If
'    If Weekday(dtpOldWorkDate.Value) = vbSaturday Or Weekday(dtpOldWorkDate.Value) = vbSunday Then
'        MsgBox "ԭ�ϰ�ʱ�䲻��Ϊ��Ϣ��(����������)��", vbInformation, gstrSysName
'        If dtpOldWorkDate.Visible And dtpOldWorkDate.Enabled Then dtpOldWorkDate.SetFocus
'        Exit Sub
'    End If
    If Format(dtpNewWorkDate.Value, "yyyy-mm-dd") >= Format(dtpStartDate.Value, "yyyy-mm-dd") And _
        Format(dtpNewWorkDate.Value, "yyyy-mm-dd") <= Format(dtpEndDate.Value, "yyyy-mm-dd") Then
        MsgBox "����ʱ�䲻���ڽڼ���ʱ�䷶Χ�ڣ�", vbInformation, gstrSysName
        If dtpNewWorkDate.Visible And dtpNewWorkDate.Enabled Then dtpNewWorkDate.SetFocus
        Exit Sub
    End If
'    If Not (Weekday(dtpNewWorkDate.Value) = vbSaturday Or Weekday(dtpNewWorkDate.Value) = vbSunday) Then
'        MsgBox "����ʱ�����Ϊ��Ϣ��(����������)��", vbInformation, gstrSysName
'        If dtpNewWorkDate.Visible And dtpNewWorkDate.Enabled Then dtpNewWorkDate.SetFocus
'        Exit Sub
'    End If
    
    For i = 1 To vsf�������.Rows - 1
        If Format(dtpOldWorkDate.Value, "yyyy-mm-dd") = vsf�������.TextMatrix(i, COL_ԭ�ϰ�ʱ��) Then
            MsgBox "ԭ�ϰ�ʱ�������õ����� " & vsf�������.TextMatrix(i, Col_����ʱ��) & " ��", vbInformation, gstrSysName
            If dtpOldWorkDate.Visible And dtpOldWorkDate.Enabled Then dtpOldWorkDate.SetFocus
            Exit Sub
        End If
        If Format(dtpNewWorkDate.Value, "yyyy-mm-dd") = vsf�������.TextMatrix(i, Col_����ʱ��) Then
            MsgBox "����ʱ���ѱ�����Ϊԭ�ϰ�ʱ�� " & vsf�������.TextMatrix(i, COL_ԭ�ϰ�ʱ��) & " �ĵ����գ�", vbInformation, gstrSysName
            If dtpNewWorkDate.Visible And dtpNewWorkDate.Enabled Then dtpNewWorkDate.SetFocus
            Exit Sub
        End If
    Next
    AddGridRow dtpOldWorkDate.Value, dtpNewWorkDate.Value
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    
    Err = 0: On Error GoTo ErrHandler
    If mbytFun = Fun_View Then Unload Me: Exit Sub
    
    cmdOk.Enabled = False
    If IsValied() = False Then cmdOk.Enabled = True: Exit Sub
    If SaveData() = False Then cmdOk.Enabled = True: Exit Sub
    
    mblnOK = True
    If mbytFun = Fun_Add Then
        Call ClearFaceInfor
        cmdOk.Enabled = True
        Exit Sub
    End If
    Unload Me
    Exit Sub
ErrHandler:
    cmdOk.Enabled = True
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function SaveData() As Boolean
    Dim strSQL As String, i As Long
    Dim str������� As String
    
    Err = 0: On Error GoTo ErrHandler
    With vsf�������
        For i = 1 To .Rows - 1
            str������� = str������� & ";" & .TextMatrix(i, Col_����ʱ��) & "~" & .TextMatrix(i, COL_ԭ�ϰ�ʱ��)
        Next
        If str������� <> "" Then str������� = Mid(str�������, 2)
    End With
    Call GetDateRegist
    
    Select Case mbytFun
    Case Fun_Add
        'Zl_�������ձ�_Modify(
        strSQL = "Zl_�������ձ�_Modify("
        '��������_In Number,--0-������1-�޸�
        strSQL = strSQL & "" & 0 & ","
        '���_In     �������ձ�.���%Type,
        strSQL = strSQL & "" & Val(cboYear.Text) & ","
        '��������_In �������ձ�.��������%Type,
        strSQL = strSQL & "'" & Trim(cboHolidayName.Text) & "',"
        '��ʼ����_In �������ձ�.��ʼ����%Type,
        strSQL = strSQL & "To_Date('" & Format(dtpStartDate.Value, "yyyy-mm-dd") & " " & Format(dtpStartTime.Value, "hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        '��ֹ����_In �������ձ�.��ֹ����%Type,
        strSQL = strSQL & "To_Date('" & Format(dtpEndDate.Value, "yyyy-mm-dd") & " " & Format(dtpEndTime.Value, "hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        '��ע_In     �������ձ�.��ע%Type,
        strSQL = strSQL & "'" & Trim(txtComment.Text) & "',"
        '�������_In Varchar2:=Null--��ʽ������ʱ��1~ ԭ�ϰ�ʱ��1;����ʱ��2~ ԭ�ϰ�ʱ��2
        strSQL = strSQL & "'" & str������� & "',"
        '����ԤԼ����_in ����ԤԼ������,��ʽ��yyyy-mm-dd;yyyy-mm-dd;...
        strSQL = strSQL & "'" & mstr����ԤԼ & "',"
        '����Һ�����_in ����Һŵ�����,��ʽ��yyyy-mm-dd;yyyy-mm-dd;...
        strSQL = strSQL & "'" & mstr����Һ� & "')"
    Case Fun_Update
        'Zl_�������ձ�_Modify(
        strSQL = "Zl_�������ձ�_Modify("
        '��������_In Number,--0-������1-�޸�
        strSQL = strSQL & "" & 1 & ","
        '���_In     �������ձ�.���%Type,
        strSQL = strSQL & "" & Val(cboYear.Text) & ","
        '��������_In �������ձ�.��������%Type,
        strSQL = strSQL & "'" & Trim(cboHolidayName.Text) & "',"
        '��ʼ����_In �������ձ�.��ʼ����%Type,
        strSQL = strSQL & "To_Date('" & Format(dtpStartDate.Value, "yyyy-mm-dd") & " " & Format(dtpStartTime.Value, "hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        '��ֹ����_In �������ձ�.��ֹ����%Type,
        strSQL = strSQL & "To_Date('" & Format(dtpEndDate.Value, "yyyy-mm-dd") & " " & Format(dtpEndTime.Value, "hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        '��ע_In     �������ձ�.��ע%Type,
        strSQL = strSQL & "'" & Trim(txtComment.Text) & "',"
        '�������_In Varchar2:=Null--��ʽ������ʱ��1~ ԭ�ϰ�ʱ��1;����ʱ��2~ ԭ�ϰ�ʱ��2
        strSQL = strSQL & "'" & str������� & "',"
        '����ԤԼ����_in ����ԤԼ������,��ʽ��yyyy-mm-dd;yyyy-mm-dd;...
        strSQL = strSQL & "'" & mstr����ԤԼ & "',"
        '����Һ�����_in ����Һŵ�����,��ʽ��yyyy-mm-dd;yyyy-mm-dd;...
        strSQL = strSQL & "'" & mstr����Һ� & "')"
    End Select
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitGridHead()
    Dim strHead As String
    Dim i As Long, varData As Variant
    
    Err = 0: On Error GoTo ErrHandler
    strHead = "���,4,700|ԭ�ϰ�����,4,1300|��������,4,1300"
    With vsf�������
        .Redraw = flexRDNone
        .FixedCols = 0: .FixedRows = 1
        .HighLight = flexHighlightAlways
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .BackColorAlternate = G_AlternateColor
        .RowHeightMin = 300
        
        .Rows = 1
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .Redraw = flexRDBuffered
    End With
    
    strHead = "����,4,1300|����Һ�,4,1000|����ԤԼ,4,1000"
    With vsfRegistInfo
        .Redraw = flexRDNone
        .FixedCols = 0: .FixedRows = 1
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .BackColorAlternate = G_AlternateColor
        .RowHeightMin = 300
        .Editable = IIf(mbytFun = Fun_View, flexEDNone, flexEDKbdMouse)
        
        .Rows = 1
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = flexAlignCenterCenter
            If i > 0 Then
                .ColDataType(i) = flexDTBoolean
            End If
        Next
        .Redraw = flexRDBuffered
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ShowDateRangeToGrid(ByVal dtStart As Date, dtEnd As Date)
    '��ʾ���ڵ������
    Dim lngRow As Long, i As Integer
    Dim intCount As Integer
    
    Err = 0: On Error GoTo ErrHandler
    intCount = DateDiff("d", dtStart, dtEnd) '������
    With vsfRegistInfo
        .Clear 1
        .Rows = 1
        For i = 0 To intCount
            .Rows = .Rows + 1
            lngRow = .Rows - 1
            .TextMatrix(lngRow, COL_����) = Format(DateAdd("d", i, dtStart), "yyyy-mm-dd")
            .Cell(flexcpChecked, lngRow, COL_����Һ�, lngRow, COL_����ԤԼ) = 2
        Next
    End With
    Call LoadDateRegist
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub GetDateRegist()
    '��ȡԤԼ�Һ����
    Dim i As Integer
    
    Err = 0: On Error GoTo ErrHandler
    mstr����ԤԼ = "": mstr����Һ� = ""
    With vsfRegistInfo
        For i = 1 To .Rows - 1
            If Abs(Val(.TextMatrix(i, COL_����Һ�))) = 1 Then
                mstr����Һ� = mstr����Һ� & ";" & Format(.TextMatrix(i, COL_����), "yyyy-mm-dd")
            End If
            If Abs(Val(.TextMatrix(i, COL_����ԤԼ))) = 1 Then
                mstr����ԤԼ = mstr����ԤԼ & ";" & Format(.TextMatrix(i, COL_����), "yyyy-mm-dd")
            End If
        Next
    End With
    If mstr����Һ� <> "" Then mstr����Һ� = Mid(mstr����Һ�, 2)
    If mstr����ԤԼ <> "" Then mstr����ԤԼ = Mid(mstr����ԤԼ, 2)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub LoadDateRegist()
    '����ԤԼ�Һ����
    Dim i As Integer, j As Integer
    Dim var����ԤԼ As Variant, var����Һ� As Variant
    
    Err = 0: On Error GoTo ErrHandler
    var����Һ� = Split(mstr����Һ�, ";")
    var����ԤԼ = Split(mstr����ԤԼ, ";")
    With vsfRegistInfo
        For i = 1 To .Rows - 1
            For j = 0 To UBound(var����Һ�)
                If DateDiff("d", .TextMatrix(i, COL_����), var����Һ�(j)) = 0 Then
                    .TextMatrix(i, COL_����Һ�) = 1
                End If
            Next
            For j = 0 To UBound(var����ԤԼ)
                If DateDiff("d", .TextMatrix(i, COL_����), var����ԤԼ(j)) = 0 Then
                    .TextMatrix(i, COL_����ԤԼ) = 1
                End If
            Next
        Next
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub AddGridRow(ByVal strOldWorkDate As String, ByVal strNewWorkDate As String)
    '��������
    Dim lngRow As Long
    
    Err = 0: On Error GoTo ErrHandler
    With vsf�������
        .Rows = .Rows + 1
        lngRow = .Rows - 1
        .TextMatrix(lngRow, COL_���) = .Rows - 1
        .TextMatrix(lngRow, COL_ԭ�ϰ�ʱ��) = Format(strOldWorkDate, "yyyy-mm-dd")
        .TextMatrix(lngRow, Col_����ʱ��) = Format(strNewWorkDate, "yyyy-mm-dd")
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsDefautHoliday = Nothing
End Sub

Private Sub txtComment_GotFocus()
    zlControl.TxtSelAll txtComment
End Sub

Private Sub txtComment_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtComment_Validate(Cancel As Boolean)
    If zlCommFun.ActualLen(txtComment.Text) > 100 Then
        MsgBox "����˵��ֻ��������100���ַ���50�����֣�", vbInformation, gstrSysName
        zlControl.TxtSelAll txtComment
        Cancel = True
    End If
End Sub

Private Sub vsfRegistInfo_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = COL_����Һ� Then
        If Abs(Val(vsfRegistInfo.TextMatrix(Row, Col))) <> 1 Then
            vsfRegistInfo.TextMatrix(Row, COL_����ԤԼ) = 0
        End If
    ElseIf Col = COL_����ԤԼ Then
        If Abs(Val(vsfRegistInfo.TextMatrix(Row, Col))) = 1 Then
            vsfRegistInfo.TextMatrix(Row, COL_����Һ�) = 1
        End If
    End If
    Call GetDateRegist
End Sub

Private Sub vsfRegistInfo_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COL_���� Then Cancel = True: Exit Sub
End Sub

Private Sub vsfRegistInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub vsfRegistInfo_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub vsf�������_EnterCell()
    cmdDelete.Enabled = vsf�������.Row > 0 And (mbytFun = Fun_Add Or mbytFun = Fun_Update)
End Sub

Private Sub cmdDelete_Click()
    Dim i As Integer
    
    Err = 0: On Error GoTo ErrHandler
    If vsf�������.Row > 0 Then
        If MsgBox("��ȷ��Ҫɾ���� " & vsf�������.Row & " �У�", vbQuestion + vbOKCancel + vbDefaultButton2, gstrSysName) = vbOK Then
            vsf�������.RemoveItem vsf�������.Row
            For i = 1 To vsf�������.Rows - 1 '���±��
                vsf�������.TextMatrix(i, 0) = i
            Next
        End If
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsf�������_GotFocus()
    If vsf�������.Rows > 1 Then
        vsf�������.Row = 1
    Else
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Function IsValied() As Boolean
    Dim rs�ڼ��� As ADODB.Recordset, strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim dtStart As Date, dtEnd As Date
    
    Err = 0: On Error GoTo ErrHandler
    If zlControl.FormCheckInput(Me) = False Then Exit Function
    dtStart = CDate(Format(dtpStartDate.Value, "yyyy-mm-dd ") & Format(dtpStartTime.Value, "hh:mm:ss"))
    dtEnd = CDate(Format(dtpEndDate.Value, "yyyy-mm-dd ") & Format(dtpEndTime.Value, "hh:mm:ss"))
    If cboYear.Text = "" Then
        MsgBox "��ݲ���Ϊ�գ�", vbInformation, gstrSysName
        If cboYear.Visible And cboYear.Enabled Then cboYear.SetFocus
        Exit Function
    End If
    If cboHolidayName.Text = "" Then
        MsgBox "�ڼ��ղ���Ϊ�գ�", vbInformation, gstrSysName
        If cboHolidayName.Visible And cboHolidayName.Enabled Then cboHolidayName.SetFocus
        Exit Function
    End If
    If zlCommFun.ActualLen(cboHolidayName.Text) > 50 Then
        MsgBox "�ڼ�������ֻ��������50���ַ���25�����֣�", vbInformation, gstrSysName
        If cboHolidayName.Visible And cboHolidayName.Enabled Then cboHolidayName.SetFocus
        zlControl.TxtSelAll cboHolidayName
        Exit Function
    End If
    
    If dtStart >= dtEnd Then
        MsgBox "�ڼ��յ���ֹʱ�������ڿ�ʼʱ�䣡", vbInformation, gstrSysName
        If dtpEndDate.Visible And dtpEndDate.Enabled Then dtpEndDate.SetFocus
        Exit Function
    End If
    
    If mbytFun = Fun_Add Then
        strSQL = "Select 1 From �������ձ� Where Nvl(����,0)=0 And ���=[1] And ��������=[2] And Rownum < 2"
        Set rs�ڼ��� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(cboYear.Text), Trim(cboHolidayName.Text))
        If Not rs�ڼ���.EOF Then
            MsgBox cboYear.Text & "�Ѵ��ڡ�" & cboHolidayName.Text & "����", vbInformation, gstrSysName
            If cboHolidayName.Visible And cboHolidayName.Enabled Then cboHolidayName.SetFocus
            zlControl.TxtSelAll cboHolidayName
            Exit Function
        End If
        
        strSQL = "Select 1 From �������ձ� Where ���� = 0 And [1] < ��ֹ���� And [2] > ��ʼ���� And Rownum < 2"
        Set rs�ڼ��� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtStart, dtEnd)
        If Not rs�ڼ���.EOF Then
            MsgBox "��ǰ�ڼ��յ�ʱ�䷶Χ���Ѵ��������ڼ��գ�", vbInformation, gstrSysName
            If dtpStartDate.Visible And dtpStartDate.Enabled Then dtpStartDate.SetFocus
            Exit Function
        End If
    Else
        strSQL = "Select 1" & vbNewLine & _
            "    From �ٴ������¼ A" & vbNewLine & _
            "    Where a.�������� >= (Select ��ʼ���� From �������ձ� Where ��� = [1] And �������� = [2] And ���� = 0 And Rownum<2)" & vbNewLine & _
            "          And a.�ϰ�ʱ�� Is Not Null And Nvl(a.�Ƿ񷢲�, 0) = 1 And Rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(cboYear.Text), Trim(cboHolidayName.Text))
        If Not rsTemp.EOF Then
            MsgBox "��ǰ�ڼ��տ�ʼʱ��֮��������Ч�ĳ��ﰲ�ţ������޸ģ�", vbInformation, gstrSysName
            Exit Function
        End If
        
        strSQL = "Select 1 From �������ձ� Where ���� = 0 And [1] < ��ֹ���� And [2] > ��ʼ���� And Not (��� = [3] And �������� = [4]) And Rownum < 2"
        Set rs�ڼ��� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtStart, dtEnd, Val(cboYear.Text), Trim(cboHolidayName.Text))
        If Not rs�ڼ���.EOF Then
            MsgBox "��ǰ�ڼ��յ�ʱ�䷶Χ���Ѵ��������ڼ��գ�", vbInformation, gstrSysName
            If dtpStartDate.Visible And dtpStartDate.Enabled Then dtpStartDate.SetFocus
            Exit Function
        End If
    End If
    
    strSQL = "Select 1" & vbNewLine & _
        "    From �ٴ������¼ A" & vbNewLine & _
        "    Where a.�������� >=[1] And a.�ϰ�ʱ�� Is Not Null And Nvl(a.�Ƿ񷢲�, 0) = 1 And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtStart)
    If Not rsTemp.EOF Then
        MsgBox "��ʼʱ��֮��������Ч�ĳ��ﰲ�ţ�" & IIf(mbytFun = Fun_Update, "�����޸ģ�", "�������ã�"), vbInformation, gstrSysName
        Exit Function
    End If
    
    strSQL = "Select 1 From �ٴ������¼" & vbNewLine & _
            " Where �������� Between [1] And [2] And Nvl(�Ƿ񷢲�, 0) = 1" & vbNewLine & _
            "       And (Nvl(��Լ��, 0) <> 0 Or Nvl(�ѹ���, 0) <> 0) And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtStart, dtEnd)
    If Not rsTemp.EOF Then
        MsgBox "��ǰ�ڼ��յ�ʱ�䷶Χ������ԤԼ�ҺŲ��ˣ�" & IIf(mbytFun = Fun_Update, "�����޸ģ�", "�������ã�"), vbInformation, gstrSysName
        Exit Function
    End If
    
    For i = 1 To vsf�������.Rows - 1
        If CDate(vsf�������.TextMatrix(i, COL_ԭ�ϰ�ʱ��)) < dtStart Or CDate(vsf�������.TextMatrix(i, COL_ԭ�ϰ�ʱ��)) > dtEnd Then
            MsgBox "��" & i & "��ԭ�ϰ�ʱ�䲻�ڽڼ���ʱ�䷶Χ�ڣ�", vbInformation, gstrSysName
            vsf�������.Row = i
            Exit Function
        End If
        If CDate(vsf�������.TextMatrix(i, Col_����ʱ��)) >= dtStart And CDate(vsf�������.TextMatrix(i, Col_����ʱ��)) <= dtEnd Then
            MsgBox "��" & i & "�е���ʱ�䲻���ڽڼ���ʱ�䷶Χ�ڣ�", vbInformation, gstrSysName
            vsf�������.Row = i
            Exit Function
        End If
    Next
    
    If zlCommFun.ActualLen(txtComment.Text) > 100 Then
        MsgBox "����˵��ֻ��������100���ַ���50�����֣�", vbInformation, gstrSysName
        If txtComment.Visible And txtComment.Enabled Then txtComment.SetFocus
        zlControl.TxtSelAll txtComment
        Exit Function
    End If
    
    IsValied = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function InitDefautHoliday() As Boolean
    'ȱʡ�ڼ��ռ�¼��
    Dim strHoliday As String
    Dim varHoliday As Variant
    Dim i As Integer, varTemp As Variant
    
    Err = 0: On Error GoTo ErrHandler
    Set mrsDefautHoliday = New ADODB.Recordset
    With mrsDefautHoliday
        '���,��������,��ʼ����,��ֹ����,��ע,����ԤԼ,����Һ�
        .Fields.Append "���", adBigInt, 10
        .Fields.Append "��������", adLongVarChar, 100
        .Fields.Append "��ʼ����", adLongVarChar, 100
        .Fields.Append "��ֹ����", adLongVarChar, 100
        .Fields.Append "��ע", adLongVarChar, 1000
        .Fields.Append "����ԤԼ����", adLongVarChar, 500
        .Fields.Append "����Һ�����", adLongVarChar, 500
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
        
        strHoliday = "Ԫ����,2016-1-1 00:00:00,2016-1-3 23:59:59|" & _
                    "����,2016-2-7 00:00:00,2016-2-13 23:59:59|" & _
                    "��Ů��,2016-3-8 12:00:00,2016-3-8 23:59:59|" & _
                    "������,2016-4-2 00:00:00,2016-4-4 23:59:59|" & _
                    "�Ͷ���,2016-5-1 00:00:00,2016-5-3 23:59:59|" & _
                    "�����,2016-6-9 00:00:00,2016-6-11 23:59:59|" & _
                    "�����,2016-9-15 00:00:00,2016-9-17 23:59:59|" & _
                    "�����,2016-10-1 00:00:00,2016-10-7 23:59:59"
        varHoliday = Split(strHoliday, "|")
        For i = 0 To UBound(varHoliday)
            varTemp = Split(varHoliday(i), ",")
            .AddNew
            !�������� = varTemp(0)
            !��ʼ���� = varTemp(1)
            !��ֹ���� = varTemp(2)
            .Update
        Next
    End With
    InitDefautHoliday = True
    Exit Function
ErrHandler:
    
End Function

Private Sub vsf�������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub ClearFaceInfor()
    '����:���������Ϣ���Ա�������������
    On Error GoTo errHandle
    mstr����ԤԼ = "": mstr����Һ� = ""
    If cboYear.ListCount > 0 Then cboYear.ListIndex = 0
    cboHolidayName.Text = "": cboHolidayName.ListIndex = -1
    
    dtpStartDate.Value = Format(dtpStartDate.MinDate, "yyyy-mm-dd")
    dtpEndDate.Value = Format(dtpEndDate.MinDate, "yyyy-mm-dd")
    Call ShowDateRangeToGrid(dtpStartDate.Value, dtpEndDate.Value)
    dtpStartTime.Value = "00:00:00"
    dtpEndTime.Value = "00:00:00"
    dtpOldWorkDate.Value = Format(dtpOldWorkDate.MinDate, "yyyy-mm-dd")
    dtpNewWorkDate.Value = Format(dtpNewWorkDate.MinDate, "yyyy-mm-dd")
    txtComment.Text = ""
    
    vsf�������.Clear 1
    vsf�������.Rows = 1
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
