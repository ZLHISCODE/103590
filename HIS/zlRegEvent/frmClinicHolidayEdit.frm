VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClinicHolidayEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "节假日设置"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9540
   BeginProperty Font 
      Name            =   "宋体"
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
   StartUpPosition =   2  '屏幕中心
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
         Name            =   "宋体"
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
      Caption         =   "删除(&D)"
      Enabled         =   0   'False
      Height          =   320
      Left            =   6840
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2970
      Width           =   885
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "增加(&A)"
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
      Caption         =   "取消(&C)"
      Height          =   360
      Left            =   8190
      TabIndex        =   22
      Top             =   780
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   360
      Left            =   8190
      TabIndex        =   21
      Top             =   300
      Width           =   1095
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
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
         Name            =   "宋体"
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
         Name            =   "宋体"
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
         Name            =   "宋体"
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
         Name            =   "宋体"
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
         Name            =   "宋体"
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
         Name            =   "宋体"
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
   Begin VSFlex8Ctl.VSFlexGrid vsf调休情况 
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
         Name            =   "宋体"
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
      Caption         =   "节假日"
      Height          =   210
      Left            =   4410
      TabIndex        =   2
      Top             =   210
      Width           =   630
   End
   Begin VB.Label lblStartTime 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "开始时间"
      Height          =   210
      Left            =   60
      TabIndex        =   4
      Top             =   645
      Width           =   840
   End
   Begin VB.Label lblEndTime 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "结束时间"
      Height          =   210
      Left            =   4200
      TabIndex        =   7
      Top             =   630
      Width           =   840
   End
   Begin VB.Label lblComment 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "补充说明"
      Height          =   240
      Left            =   60
      TabIndex        =   19
      Top             =   4770
      Width           =   840
   End
   Begin VB.Label lblNewWorkDate 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "调休日期"
      Height          =   210
      Left            =   3630
      TabIndex        =   14
      Top             =   3015
      Width           =   840
   End
   Begin VB.Label lblOldWorkTime 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "原上班日期"
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
   Begin VB.Label lbl调休信息 
      AutoSize        =   -1  'True
      Caption         =   "调休信息"
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
      Caption         =   "年份"
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
Private mbytFun As G_Enum_Fun '0-查看,1-添加,2-调整
Private mlngYear As Long
Private mstrHolidayName As String

Private Enum mGridHeadCol
    COL_序号 = 0
    COL_原上班时间 = 1
    Col_调休时间 = 2
    
    COL_日期 = 0
    COL_允许挂号 = 1
    COL_允许预约 = 2
End Enum
Private mblnOK As Boolean
Private mblnNotClick As Boolean
Private mrsDefautHoliday As ADODB.Recordset
Private mstr允许预约 As String '格式：yyyy-mm-dd;yyyy-mm-dd;...
Private mstr允许挂号 As String '格式：yyyy-mm-dd;yyyy-mm-dd;...

Public Function ShowMe(frmParent As Form, ByVal bytFun As G_Enum_Fun, _
    Optional ByVal lngYear As Long, Optional ByVal strHolidayName As String) As Boolean
    '入参：
    '   frmParent - 父窗口
    '   bytFun - 操作类型, 0-查看，1-新增，2-修改
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
        MsgBox "节假日名称只允许输入50个字符或25个汉字！", vbInformation, gstrSysName
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
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    Dim varRow As Variant, i As Long, lngYear As Long
    Dim rs节假日 As ADODB.Recordset, strSQL As String
    Dim varArray As Variant
    
    Err = 0: On Error GoTo ErrHandler
    mstr允许预约 = "": mstr允许挂号 = ""
    Call InitGridHead
    cboYear.Clear
    For i = Year(Now) To Year(Now) + 10 '缺省加入10年供选择
        cboYear.AddItem i & "年"
    Next
    cboYear.ListIndex = 0
    
    varArray = Array("元旦节", "春节", "妇女节", "清明节", "劳动节", "端午节", "中秋节", "国庆节")
    cboHolidayName.Clear
    For i = 0 To UBound(varArray)
        cboHolidayName.AddItem varArray(i)
    Next
    If mbytFun = Fun_Add Or mbytFun = Fun_Update Then
        Call InitDefautHoliday
    End If
    Me.Caption = Choose(mbytFun + 1, "查看", "新增", "修改", "删除") & "节假日"
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
    '加载数据
    'lngYear=0表示设置节假日缺省时间
    Dim varRow As Variant, lngRow As Long, blnDefaut As Boolean
    Dim rs节假日 As ADODB.Recordset, strSQL As String
    
    Err = 0: On Error GoTo ErrHandler
    blnDefaut = lngYear = 0
    strSQL = "Select 年份,节日名称,开始日期,终止日期,备注,允许预约日期,允许挂号日期 From 法定假日表" & vbNewLine & _
            " Where Nvl(性质,0)=0 And 节日名称=[1]" & IIf(lngYear = 0, "", " And 年份=[2]") & vbNewLine & _
            " Order By 年份 Desc"
    Set rs节假日 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strHolidayName, lngYear)
    
    If rs节假日.EOF And blnDefaut Then '设置节假日缺省时间
        Set rs节假日 = mrsDefautHoliday.Clone
        rs节假日.Filter = "节日名称='" & strHolidayName & "'"
    End If
    If rs节假日.RecordCount = 0 Then Exit Function
    
    If blnDefaut = False Then
        '年份
        zlControl.CboSetText cboYear, lngYear & "年"
        
        If cboYear.Text = "" Then
            cboYear.AddItem lngYear & "年"
            cboYear.ListIndex = cboYear.NewIndex
        End If
    End If
    
    mstr允许预约 = Nvl(rs节假日!允许预约日期)
    mstr允许挂号 = Nvl(rs节假日!允许挂号日期)
    If lngYear <> 0 Then
        If Nvl(rs节假日!开始日期) >= dtpStartDate.MinDate And Nvl(rs节假日!开始日期) <= dtpStartDate.MaxDate Then
            dtpStartDate.Value = Format(Nvl(rs节假日!开始日期), "yyyy-mm-dd")
        End If
        If Nvl(rs节假日!终止日期) >= dtpEndDate.MinDate And Nvl(rs节假日!终止日期) <= dtpEndDate.MaxDate Then
            dtpEndDate.Value = Format(Nvl(rs节假日!终止日期), "yyyy-mm-dd")
        End If
    Else
        If Val(cboYear.Text) & Format(Nvl(rs节假日!开始日期), "-mm-dd") >= dtpStartDate.MinDate _
            And Val(cboYear.Text) & Format(Nvl(rs节假日!开始日期), "-mm-dd") <= dtpStartDate.MaxDate Then
            dtpStartDate.Value = Val(cboYear.Text) & Format(Nvl(rs节假日!开始日期), "-mm-dd")
        End If
        If Val(cboYear.Text) & Format(Nvl(rs节假日!终止日期), "-mm-dd") >= dtpEndDate.MinDate _
            And Val(cboYear.Text) & Format(Nvl(rs节假日!终止日期), "-mm-dd") <= dtpEndDate.MaxDate Then
            dtpEndDate.Value = Val(cboYear.Text) & Format(Nvl(rs节假日!终止日期), "-mm-dd")
        End If
    End If
    Call ShowDateRangeToGrid(dtpStartDate.Value, dtpEndDate.Value)
    dtpStartTime.Value = Format(Nvl(rs节假日!开始日期), "hh:mm:ss")
    dtpEndTime.Value = Format(Nvl(rs节假日!终止日期), "hh:mm:ss")
    dtpOldWorkDate.Value = dtpStartDate.Value
    dtpNewWorkDate.Value = dtpStartDate.Value
    If blnDefaut Then LoadData = True: Exit Function
    
    '节日名称
    mblnNotClick = True
    zlControl.CboSetText cboHolidayName, strHolidayName
    mblnNotClick = False
    If cboHolidayName.Text = "" Then
        cboHolidayName.AddItem strHolidayName
        mblnNotClick = True
        cboHolidayName.ListIndex = cboHolidayName.NewIndex
        mblnNotClick = False
    End If
    txtComment.Text = Nvl(rs节假日!备注)
    
    '换休情况
    strSQL = "Select 年份,节日名称,开始日期,终止日期,备注 From 法定假日表" & vbNewLine & _
            " Where Nvl(性质,0)=1 And 节日名称=[1] And 年份=[2]"
    Set rs节假日 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strHolidayName, lngYear)
    vsf调休情况.Rows = rs节假日.RecordCount + 1
    lngRow = 1
    Do While Not rs节假日.EOF
        vsf调休情况.TextMatrix(lngRow, COL_序号) = lngRow
        vsf调休情况.TextMatrix(lngRow, COL_原上班时间) = Format(Nvl(rs节假日!终止日期), "yyyy-mm-dd")
        vsf调休情况.TextMatrix(lngRow, Col_调休时间) = Format(Nvl(rs节假日!开始日期), "yyyy-mm-dd")
        lngRow = lngRow + 1
        rs节假日.MoveNext
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
        MsgBox "节假日的终止时间必须大于开始时间！", vbInformation, gstrSysName
        If dtpEndDate.Visible And dtpEndDate.Enabled Then dtpEndDate.SetFocus
        Exit Sub
    End If
    If Format(dtpOldWorkDate.Value, "yyyy-mm-dd") = Format(dtpNewWorkDate.Value, "yyyy-mm-dd") Then
        MsgBox "调休时间和原上班时间不能为同一天！", vbInformation, gstrSysName
        If dtpNewWorkDate.Visible And dtpNewWorkDate.Enabled Then dtpNewWorkDate.SetFocus
        Exit Sub
    End If
    If Format(dtpOldWorkDate.Value, "yyyy-mm-dd") < Format(dtpStartDate.Value, "yyyy-mm-dd") Or _
        Format(dtpOldWorkDate.Value, "yyyy-mm-dd") > Format(dtpEndDate.Value, "yyyy-mm-dd") Then
        MsgBox "原上班时间必须在节假日时间范围内！", vbInformation, gstrSysName
        If dtpOldWorkDate.Visible And dtpOldWorkDate.Enabled Then dtpOldWorkDate.SetFocus
        Exit Sub
    End If
'    If Weekday(dtpOldWorkDate.Value) = vbSaturday Or Weekday(dtpOldWorkDate.Value) = vbSunday Then
'        MsgBox "原上班时间不能为休息日(周六、周日)！", vbInformation, gstrSysName
'        If dtpOldWorkDate.Visible And dtpOldWorkDate.Enabled Then dtpOldWorkDate.SetFocus
'        Exit Sub
'    End If
    If Format(dtpNewWorkDate.Value, "yyyy-mm-dd") >= Format(dtpStartDate.Value, "yyyy-mm-dd") And _
        Format(dtpNewWorkDate.Value, "yyyy-mm-dd") <= Format(dtpEndDate.Value, "yyyy-mm-dd") Then
        MsgBox "调休时间不能在节假日时间范围内！", vbInformation, gstrSysName
        If dtpNewWorkDate.Visible And dtpNewWorkDate.Enabled Then dtpNewWorkDate.SetFocus
        Exit Sub
    End If
'    If Not (Weekday(dtpNewWorkDate.Value) = vbSaturday Or Weekday(dtpNewWorkDate.Value) = vbSunday) Then
'        MsgBox "调休时间必须为休息日(周六、周日)！", vbInformation, gstrSysName
'        If dtpNewWorkDate.Visible And dtpNewWorkDate.Enabled Then dtpNewWorkDate.SetFocus
'        Exit Sub
'    End If
    
    For i = 1 To vsf调休情况.Rows - 1
        If Format(dtpOldWorkDate.Value, "yyyy-mm-dd") = vsf调休情况.TextMatrix(i, COL_原上班时间) Then
            MsgBox "原上班时间已设置调休日 " & vsf调休情况.TextMatrix(i, Col_调休时间) & " ！", vbInformation, gstrSysName
            If dtpOldWorkDate.Visible And dtpOldWorkDate.Enabled Then dtpOldWorkDate.SetFocus
            Exit Sub
        End If
        If Format(dtpNewWorkDate.Value, "yyyy-mm-dd") = vsf调休情况.TextMatrix(i, Col_调休时间) Then
            MsgBox "调休时间已被设置为原上班时间 " & vsf调休情况.TextMatrix(i, COL_原上班时间) & " 的调休日！", vbInformation, gstrSysName
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
    Dim str换休情况 As String
    
    Err = 0: On Error GoTo ErrHandler
    With vsf调休情况
        For i = 1 To .Rows - 1
            str换休情况 = str换休情况 & ";" & .TextMatrix(i, Col_调休时间) & "~" & .TextMatrix(i, COL_原上班时间)
        Next
        If str换休情况 <> "" Then str换休情况 = Mid(str换休情况, 2)
    End With
    Call GetDateRegist
    
    Select Case mbytFun
    Case Fun_Add
        'Zl_法定假日表_Modify(
        strSQL = "Zl_法定假日表_Modify("
        '操作类型_In Number,--0-新增，1-修改
        strSQL = strSQL & "" & 0 & ","
        '年份_In     法定假日表.年份%Type,
        strSQL = strSQL & "" & Val(cboYear.Text) & ","
        '节日名称_In 法定假日表.节日名称%Type,
        strSQL = strSQL & "'" & Trim(cboHolidayName.Text) & "',"
        '开始日期_In 法定假日表.开始日期%Type,
        strSQL = strSQL & "To_Date('" & Format(dtpStartDate.Value, "yyyy-mm-dd") & " " & Format(dtpStartTime.Value, "hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        '终止日期_In 法定假日表.终止日期%Type,
        strSQL = strSQL & "To_Date('" & Format(dtpEndDate.Value, "yyyy-mm-dd") & " " & Format(dtpEndTime.Value, "hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        '备注_In     法定假日表.备注%Type,
        strSQL = strSQL & "'" & Trim(txtComment.Text) & "',"
        '换休情况_In Varchar2:=Null--格式：调休时间1~ 原上班时间1;调休时间2~ 原上班时间2
        strSQL = strSQL & "'" & str换休情况 & "',"
        '允许预约日期_in 允许预约的日期,格式：yyyy-mm-dd;yyyy-mm-dd;...
        strSQL = strSQL & "'" & mstr允许预约 & "',"
        '允许挂号日期_in 允许挂号的日期,格式：yyyy-mm-dd;yyyy-mm-dd;...
        strSQL = strSQL & "'" & mstr允许挂号 & "')"
    Case Fun_Update
        'Zl_法定假日表_Modify(
        strSQL = "Zl_法定假日表_Modify("
        '操作类型_In Number,--0-新增，1-修改
        strSQL = strSQL & "" & 1 & ","
        '年份_In     法定假日表.年份%Type,
        strSQL = strSQL & "" & Val(cboYear.Text) & ","
        '节日名称_In 法定假日表.节日名称%Type,
        strSQL = strSQL & "'" & Trim(cboHolidayName.Text) & "',"
        '开始日期_In 法定假日表.开始日期%Type,
        strSQL = strSQL & "To_Date('" & Format(dtpStartDate.Value, "yyyy-mm-dd") & " " & Format(dtpStartTime.Value, "hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        '终止日期_In 法定假日表.终止日期%Type,
        strSQL = strSQL & "To_Date('" & Format(dtpEndDate.Value, "yyyy-mm-dd") & " " & Format(dtpEndTime.Value, "hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
        '备注_In     法定假日表.备注%Type,
        strSQL = strSQL & "'" & Trim(txtComment.Text) & "',"
        '换休情况_In Varchar2:=Null--格式：调休时间1~ 原上班时间1;调休时间2~ 原上班时间2
        strSQL = strSQL & "'" & str换休情况 & "',"
        '允许预约日期_in 允许预约的日期,格式：yyyy-mm-dd;yyyy-mm-dd;...
        strSQL = strSQL & "'" & mstr允许预约 & "',"
        '允许挂号日期_in 允许挂号的日期,格式：yyyy-mm-dd;yyyy-mm-dd;...
        strSQL = strSQL & "'" & mstr允许挂号 & "')"
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
    strHead = "序号,4,700|原上班日期,4,1300|调休日期,4,1300"
    With vsf调休情况
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
    
    strHead = "日期,4,1300|允许挂号,4,1000|允许预约,4,1000"
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
    '显示日期到表格中
    Dim lngRow As Long, i As Integer
    Dim intCount As Integer
    
    Err = 0: On Error GoTo ErrHandler
    intCount = DateDiff("d", dtStart, dtEnd) '总天数
    With vsfRegistInfo
        .Clear 1
        .Rows = 1
        For i = 0 To intCount
            .Rows = .Rows + 1
            lngRow = .Rows - 1
            .TextMatrix(lngRow, COL_日期) = Format(DateAdd("d", i, dtStart), "yyyy-mm-dd")
            .Cell(flexcpChecked, lngRow, COL_允许挂号, lngRow, COL_允许预约) = 2
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
    '获取预约挂号情况
    Dim i As Integer
    
    Err = 0: On Error GoTo ErrHandler
    mstr允许预约 = "": mstr允许挂号 = ""
    With vsfRegistInfo
        For i = 1 To .Rows - 1
            If Abs(Val(.TextMatrix(i, COL_允许挂号))) = 1 Then
                mstr允许挂号 = mstr允许挂号 & ";" & Format(.TextMatrix(i, COL_日期), "yyyy-mm-dd")
            End If
            If Abs(Val(.TextMatrix(i, COL_允许预约))) = 1 Then
                mstr允许预约 = mstr允许预约 & ";" & Format(.TextMatrix(i, COL_日期), "yyyy-mm-dd")
            End If
        Next
    End With
    If mstr允许挂号 <> "" Then mstr允许挂号 = Mid(mstr允许挂号, 2)
    If mstr允许预约 <> "" Then mstr允许预约 = Mid(mstr允许预约, 2)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub LoadDateRegist()
    '加载预约挂号情况
    Dim i As Integer, j As Integer
    Dim var允许预约 As Variant, var允许挂号 As Variant
    
    Err = 0: On Error GoTo ErrHandler
    var允许挂号 = Split(mstr允许挂号, ";")
    var允许预约 = Split(mstr允许预约, ";")
    With vsfRegistInfo
        For i = 1 To .Rows - 1
            For j = 0 To UBound(var允许挂号)
                If DateDiff("d", .TextMatrix(i, COL_日期), var允许挂号(j)) = 0 Then
                    .TextMatrix(i, COL_允许挂号) = 1
                End If
            Next
            For j = 0 To UBound(var允许预约)
                If DateDiff("d", .TextMatrix(i, COL_日期), var允许预约(j)) = 0 Then
                    .TextMatrix(i, COL_允许预约) = 1
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
    '新增调休
    Dim lngRow As Long
    
    Err = 0: On Error GoTo ErrHandler
    With vsf调休情况
        .Rows = .Rows + 1
        lngRow = .Rows - 1
        .TextMatrix(lngRow, COL_序号) = .Rows - 1
        .TextMatrix(lngRow, COL_原上班时间) = Format(strOldWorkDate, "yyyy-mm-dd")
        .TextMatrix(lngRow, Col_调休时间) = Format(strNewWorkDate, "yyyy-mm-dd")
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
        MsgBox "补充说明只允许输入100个字符或50个汉字！", vbInformation, gstrSysName
        zlControl.TxtSelAll txtComment
        Cancel = True
    End If
End Sub

Private Sub vsfRegistInfo_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = COL_允许挂号 Then
        If Abs(Val(vsfRegistInfo.TextMatrix(Row, Col))) <> 1 Then
            vsfRegistInfo.TextMatrix(Row, COL_允许预约) = 0
        End If
    ElseIf Col = COL_允许预约 Then
        If Abs(Val(vsfRegistInfo.TextMatrix(Row, Col))) = 1 Then
            vsfRegistInfo.TextMatrix(Row, COL_允许挂号) = 1
        End If
    End If
    Call GetDateRegist
End Sub

Private Sub vsfRegistInfo_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COL_日期 Then Cancel = True: Exit Sub
End Sub

Private Sub vsfRegistInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub vsfRegistInfo_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub vsf调休情况_EnterCell()
    cmdDelete.Enabled = vsf调休情况.Row > 0 And (mbytFun = Fun_Add Or mbytFun = Fun_Update)
End Sub

Private Sub cmdDelete_Click()
    Dim i As Integer
    
    Err = 0: On Error GoTo ErrHandler
    If vsf调休情况.Row > 0 Then
        If MsgBox("您确定要删除第 " & vsf调休情况.Row & " 行？", vbQuestion + vbOKCancel + vbDefaultButton2, gstrSysName) = vbOK Then
            vsf调休情况.RemoveItem vsf调休情况.Row
            For i = 1 To vsf调休情况.Rows - 1 '重新编号
                vsf调休情况.TextMatrix(i, 0) = i
            Next
        End If
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsf调休情况_GotFocus()
    If vsf调休情况.Rows > 1 Then
        vsf调休情况.Row = 1
    Else
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Function IsValied() As Boolean
    Dim rs节假日 As ADODB.Recordset, strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim dtStart As Date, dtEnd As Date
    
    Err = 0: On Error GoTo ErrHandler
    If zlControl.FormCheckInput(Me) = False Then Exit Function
    dtStart = CDate(Format(dtpStartDate.Value, "yyyy-mm-dd ") & Format(dtpStartTime.Value, "hh:mm:ss"))
    dtEnd = CDate(Format(dtpEndDate.Value, "yyyy-mm-dd ") & Format(dtpEndTime.Value, "hh:mm:ss"))
    If cboYear.Text = "" Then
        MsgBox "年份不能为空！", vbInformation, gstrSysName
        If cboYear.Visible And cboYear.Enabled Then cboYear.SetFocus
        Exit Function
    End If
    If cboHolidayName.Text = "" Then
        MsgBox "节假日不能为空！", vbInformation, gstrSysName
        If cboHolidayName.Visible And cboHolidayName.Enabled Then cboHolidayName.SetFocus
        Exit Function
    End If
    If zlCommFun.ActualLen(cboHolidayName.Text) > 50 Then
        MsgBox "节假日名称只允许输入50个字符或25个汉字！", vbInformation, gstrSysName
        If cboHolidayName.Visible And cboHolidayName.Enabled Then cboHolidayName.SetFocus
        zlControl.TxtSelAll cboHolidayName
        Exit Function
    End If
    
    If dtStart >= dtEnd Then
        MsgBox "节假日的终止时间必须大于开始时间！", vbInformation, gstrSysName
        If dtpEndDate.Visible And dtpEndDate.Enabled Then dtpEndDate.SetFocus
        Exit Function
    End If
    
    If mbytFun = Fun_Add Then
        strSQL = "Select 1 From 法定假日表 Where Nvl(性质,0)=0 And 年份=[1] And 节日名称=[2] And Rownum < 2"
        Set rs节假日 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(cboYear.Text), Trim(cboHolidayName.Text))
        If Not rs节假日.EOF Then
            MsgBox cboYear.Text & "已存在“" & cboHolidayName.Text & "”！", vbInformation, gstrSysName
            If cboHolidayName.Visible And cboHolidayName.Enabled Then cboHolidayName.SetFocus
            zlControl.TxtSelAll cboHolidayName
            Exit Function
        End If
        
        strSQL = "Select 1 From 法定假日表 Where 性质 = 0 And [1] < 终止日期 And [2] > 开始日期 And Rownum < 2"
        Set rs节假日 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtStart, dtEnd)
        If Not rs节假日.EOF Then
            MsgBox "当前节假日的时间范围内已存在其它节假日！", vbInformation, gstrSysName
            If dtpStartDate.Visible And dtpStartDate.Enabled Then dtpStartDate.SetFocus
            Exit Function
        End If
    Else
        strSQL = "Select 1" & vbNewLine & _
            "    From 临床出诊记录 A" & vbNewLine & _
            "    Where a.出诊日期 >= (Select 开始日期 From 法定假日表 Where 年份 = [1] And 节日名称 = [2] And 性质 = 0 And Rownum<2)" & vbNewLine & _
            "          And a.上班时段 Is Not Null And Nvl(a.是否发布, 0) = 1 And Rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(cboYear.Text), Trim(cboHolidayName.Text))
        If Not rsTemp.EOF Then
            MsgBox "当前节假日开始时间之后已有有效的出诊安排，不能修改！", vbInformation, gstrSysName
            Exit Function
        End If
        
        strSQL = "Select 1 From 法定假日表 Where 性质 = 0 And [1] < 终止日期 And [2] > 开始日期 And Not (年份 = [3] And 节日名称 = [4]) And Rownum < 2"
        Set rs节假日 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtStart, dtEnd, Val(cboYear.Text), Trim(cboHolidayName.Text))
        If Not rs节假日.EOF Then
            MsgBox "当前节假日的时间范围内已存在其它节假日！", vbInformation, gstrSysName
            If dtpStartDate.Visible And dtpStartDate.Enabled Then dtpStartDate.SetFocus
            Exit Function
        End If
    End If
    
    strSQL = "Select 1" & vbNewLine & _
        "    From 临床出诊记录 A" & vbNewLine & _
        "    Where a.出诊日期 >=[1] And a.上班时段 Is Not Null And Nvl(a.是否发布, 0) = 1 And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtStart)
    If Not rsTemp.EOF Then
        MsgBox "开始时间之后已有有效的出诊安排，" & IIf(mbytFun = Fun_Update, "不能修改！", "不能设置！"), vbInformation, gstrSysName
        Exit Function
    End If
    
    strSQL = "Select 1 From 临床出诊记录" & vbNewLine & _
            " Where 出诊日期 Between [1] And [2] And Nvl(是否发布, 0) = 1" & vbNewLine & _
            "       And (Nvl(已约数, 0) <> 0 Or Nvl(已挂数, 0) <> 0) And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtStart, dtEnd)
    If Not rsTemp.EOF Then
        MsgBox "当前节假日的时间范围内已有预约挂号病人，" & IIf(mbytFun = Fun_Update, "不能修改！", "不能设置！"), vbInformation, gstrSysName
        Exit Function
    End If
    
    For i = 1 To vsf调休情况.Rows - 1
        If CDate(vsf调休情况.TextMatrix(i, COL_原上班时间)) < dtStart Or CDate(vsf调休情况.TextMatrix(i, COL_原上班时间)) > dtEnd Then
            MsgBox "第" & i & "行原上班时间不在节假日时间范围内！", vbInformation, gstrSysName
            vsf调休情况.Row = i
            Exit Function
        End If
        If CDate(vsf调休情况.TextMatrix(i, Col_调休时间)) >= dtStart And CDate(vsf调休情况.TextMatrix(i, Col_调休时间)) <= dtEnd Then
            MsgBox "第" & i & "行调休时间不能在节假日时间范围内！", vbInformation, gstrSysName
            vsf调休情况.Row = i
            Exit Function
        End If
    Next
    
    If zlCommFun.ActualLen(txtComment.Text) > 100 Then
        MsgBox "补充说明只允许输入100个字符或50个汉字！", vbInformation, gstrSysName
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
    '缺省节假日记录集
    Dim strHoliday As String
    Dim varHoliday As Variant
    Dim i As Integer, varTemp As Variant
    
    Err = 0: On Error GoTo ErrHandler
    Set mrsDefautHoliday = New ADODB.Recordset
    With mrsDefautHoliday
        '年份,节日名称,开始日期,终止日期,备注,允许预约,允许挂号
        .Fields.Append "年份", adBigInt, 10
        .Fields.Append "节日名称", adLongVarChar, 100
        .Fields.Append "开始日期", adLongVarChar, 100
        .Fields.Append "终止日期", adLongVarChar, 100
        .Fields.Append "备注", adLongVarChar, 1000
        .Fields.Append "允许预约日期", adLongVarChar, 500
        .Fields.Append "允许挂号日期", adLongVarChar, 500
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
        
        strHoliday = "元旦节,2016-1-1 00:00:00,2016-1-3 23:59:59|" & _
                    "春节,2016-2-7 00:00:00,2016-2-13 23:59:59|" & _
                    "妇女节,2016-3-8 12:00:00,2016-3-8 23:59:59|" & _
                    "清明节,2016-4-2 00:00:00,2016-4-4 23:59:59|" & _
                    "劳动节,2016-5-1 00:00:00,2016-5-3 23:59:59|" & _
                    "端午节,2016-6-9 00:00:00,2016-6-11 23:59:59|" & _
                    "中秋节,2016-9-15 00:00:00,2016-9-17 23:59:59|" & _
                    "国庆节,2016-10-1 00:00:00,2016-10-7 23:59:59"
        varHoliday = Split(strHoliday, "|")
        For i = 0 To UBound(varHoliday)
            varTemp = Split(varHoliday(i), ",")
            .AddNew
            !节日名称 = varTemp(0)
            !开始日期 = varTemp(1)
            !终止日期 = varTemp(2)
            .Update
        Next
    End With
    InitDefautHoliday = True
    Exit Function
ErrHandler:
    
End Function

Private Sub vsf调休情况_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub ClearFaceInfor()
    '功能:清除界面信息，以便重新输入数据
    On Error GoTo errHandle
    mstr允许预约 = "": mstr允许挂号 = ""
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
    
    vsf调休情况.Clear 1
    vsf调休情况.Rows = 1
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
