VERSION 5.00
Object = "{B0475000-7740-11D1-BDC3-0020AF9F8E6E}#6.1#0"; "TTF16.ocx"
Begin VB.UserControl ctrlVisForm 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   KeyPreview      =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.VScrollBar VScroll 
      Height          =   3495
      Left            =   4440
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3240
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3135
      ScaleWidth      =   4215
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.PictureBox PicForm 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2160
         Left            =   840
         ScaleHeight     =   2160
         ScaleWidth      =   2640
         TabIndex        =   2
         Top             =   360
         Width           =   2640
         Begin VB.PictureBox fraTable 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1335
            Index           =   0
            Left            =   1320
            ScaleHeight     =   1335
            ScaleWidth      =   1575
            TabIndex        =   6
            Top             =   720
            Visible         =   0   'False
            Width           =   1575
            Begin TTF160Ctl.F1Book F1Book1 
               Height          =   735
               Index           =   0
               Left            =   0
               TabIndex        =   7
               Top             =   0
               Visible         =   0   'False
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   1296
               _0              =   $"ctrlVisForm.ctx":0000
               _1              =   $"ctrlVisForm.ctx":0409
               _2              =   $"ctrlVisForm.ctx":0812
               _3              =   $"ctrlVisForm.ctx":0C1B
               _4              =   $"ctrlVisForm.ctx":1024
               _5              =   $"ctrlVisForm.ctx":142D
               _6              =   $"ctrlVisForm.ctx":1836
               _7              =   $"ctrlVisForm.ctx":1C3F
               _8              =   $"ctrlVisForm.ctx":2048
               _count          =   9
               _ver            =   2
            End
            Begin zl9CISCore.VisItem VisItem 
               Height          =   345
               Index           =   0
               Left            =   0
               TabIndex        =   8
               Top             =   0
               Visible         =   0   'False
               Width           =   1200
               _ExtentX        =   2434
               _ExtentY        =   397
               MousePointer    =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
            End
         End
         Begin VB.PictureBox Line1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            FillStyle       =   0  'Solid
            ForeColor       =   &H80000008&
            Height          =   1335
            Index           =   0
            Left            =   360
            ScaleHeight     =   1305
            ScaleWidth      =   0
            TabIndex        =   5
            Top             =   0
            Visible         =   0   'False
            Width           =   8
         End
         Begin VB.PictureBox shpDot 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            BeginProperty DataFormat 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   75
            Index           =   0
            Left            =   3360
            MousePointer    =   6  'Size NE SW
            ScaleHeight     =   45
            ScaleWidth      =   45
            TabIndex        =   4
            Top             =   3120
            Visible         =   0   'False
            Width           =   75
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   0
            Left            =   480
            Locked          =   -1  'True
            MousePointer    =   5  'Size
            TabIndex        =   3
            Text            =   "标签"
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin zl9CISCore.VisItem VisItem1 
            Height          =   345
            Index           =   0
            Left            =   720
            TabIndex        =   9
            Top             =   1440
            Visible         =   0   'False
            Width           =   1200
            _extentx        =   2434
            _extenty        =   397
            mousepointer    =   0
            font            =   "ctrlVisForm.ctx":221B
            enabled         =   0
         End
         Begin VB.Shape shpSelect 
            BorderStyle     =   3  'Dot
            Height          =   855
            Left            =   3975
            Top             =   2280
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VB.TextBox txtBox 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   3000
         MultiLine       =   -1  'True
         TabIndex        =   1
         Text            =   "ctrlVisForm.ctx":223F
         Top             =   2880
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblText 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2400
         TabIndex        =   10
         Top             =   2880
         Visible         =   0   'False
         Width           =   60
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "ctrlVisForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private VisFormID As String
Private PatientID As String '病人ID
Private CheckID As String '病案ID或挂号单ID
Private PatientType As Integer '0=门诊病人 1=住院病人
Private mblnMoved As Boolean '数据是否已转出

Private bEnabled As Boolean
Private CurrCtrlName As String, CurrCtrlIndex As Integer
Private bNotRunSelChange As Boolean

Public Event NextControl()

Private mobjParentObject As Object

Public Property Set ParentObject(vData As Object)
    Set mobjParentObject = vData
End Property

Public Property Get ParentObject() As Object
    Set ParentObject = mobjParentObject
End Property

Private Property Let Modified(vData As Boolean)
    
    On Error Resume Next
    
    If mobjParentObject Is Nothing Then Exit Property
    
    mobjParentObject.Modified = vData
    
End Property

Private Property Get Modified() As Boolean
    
    On Error Resume Next
    
    If mobjParentObject Is Nothing Then Exit Property
    
    Modified = mobjParentObject.Modified
    
End Property

Private Sub F1Book1_DblClick(Index As Integer, ByVal nRow As Long, ByVal nCol As Long)
    F1Book1(Index).StartEdit False, True, False
End Sub

Private Sub F1Book1_EndEdit(Index As Integer, EditString As String, Cancel As Integer)
    Dim iDecPos As Integer
    With F1Book1(Index)
        If IsNumeric(EditString) Then
            iDecPos = InStr(EditString, ".")
            If iDecPos > 0 And iDecPos < Len(EditString) Then
                .NumberFormat = "#." + String(Len(EditString) - iDecPos, "0")
            Else
                .NumberFormat = "General"
            End If
        Else
            .NumberFormat = "General"
        End If
        .TextRC(.Row, .Col) = EditString
        .SetRowHeightAuto .Row, 1, .Row, .MaxCol, True
    End With
End Sub

Private Sub F1Book1_GotFocus(Index As Integer)
    With F1Book1(Index)
        .Row = IIf(.Row <= .FixedRows, .FixedRows + 1, .Row)
        .Col = IIf(.Col <= .FixedCols, .FixedCols + 1, .Col)
        
        .ShowActiveCell
        bNotRunSelChange = False
        
        CurrCtrlName = "F1Book1": CurrCtrlIndex = Index
        
        ShowVisItem fraTable(F1Book1(Index).Container.Index)
    End With
End Sub

Private Sub F1Book1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim NextCtrl As Control
    On Error Resume Next
    If KeyCode = vbKeyTab Then
        Set NextCtrl = NextElement(F1Book1(Index))
        If Not NextCtrl Is Nothing Then
            NextCtrl.SetFocus
        Else
            RaiseEvent NextControl
        End If
    End If
End Sub

Private Sub F1Book1_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim objCellFormat As TTF160Ctl.F1CellFormat
    On Error Resume Next
    With F1Book1(Index)
        Set objCellFormat = .GetCellFormat
        If Len(objCellFormat.ValidationText) > 0 Then
            F1Book1_SelChange Index
            KeyAscii = 0
        End If
    End With
End Sub

Private Sub F1Book1_LostFocus(Index As Integer)
    bNotRunSelChange = True
End Sub

Private Sub F1Book1_SelChange(Index As Integer)
    Dim objCellFormat As TTF160Ctl.F1CellFormat
    Dim aVisItemInfo() As String
    
    On Error Resume Next
    If bNotRunSelChange Then Exit Sub
    If UserControl.ActiveControl.Name <> "F1Book1" Then Exit Sub
    With F1Book1(Index)
        Set objCellFormat = .GetCellFormat
        If Len(objCellFormat.ValidationText) > 0 Then
            aVisItemInfo = Split(objCellFormat.ValidationText, ",")
            Me.VisItem(aVisItemInfo(1)).SetFocus
        End If
    End With
End Sub

Private Sub F1Book1_StartEdit(Index As Integer, EditString As String, Cancel As Integer)
    Dim objCellFormat As TTF160Ctl.F1CellFormat
    On Error Resume Next
    With F1Book1(Index)
        Set objCellFormat = .GetCellFormat
        If Len(objCellFormat.ValidationText) > 0 Then
            Cancel = True
        End If
    End With
End Sub

Private Sub F1Book1_TopLeftChanged(Index As Integer)
    If bNotRunSelChange Then Exit Sub
    
    bNotRunSelChange = True
    Proc_Table_TopLeftChanged F1Book1(Index)
    bNotRunSelChange = False
End Sub

Private Sub HScroll_Change()
    On Error Resume Next
    PicForm.Left = -1 * HScroll.Value
End Sub

Private Sub UserControl_InitProperties()
    bEnabled = False
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngNewRow As Long, lngNewCol As Long
    On Error Resume Next
    If UCase(UserControl.ActiveControl.Name) <> "F1BOOK1" Then Exit Sub
    
    If KeyCode = vbKeyReturn Then
        With UserControl.ActiveControl
            If .Row = .MaxRow Then
                lngNewRow = .FixedRows + 1
                If .Col = .MaxCol Then
                    lngNewCol = .FixedCols + 1
                Else
                    lngNewCol = .Col + 1
                End If
            Else
                lngNewRow = .Row + 1: lngNewCol = .Col
            End If
            .SetActiveCell lngNewRow, lngNewCol
            .ShowActiveCell
        End With
        KeyCode = 0
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set UserControl.Font = PropBag.ReadProperty("Font", UserControl.Font)
    SetFont
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    With PicMain
        .Left = 0: .Top = 0
        .Width = UserControl.ScaleWidth
        .Height = UserControl.ScaleHeight
        
        If .Width < PicForm.Width Then .Height = .Height - HScroll.Height
        If .Height < PicForm.Height Then .Width = .Width - VScroll.Width
        If .Width < PicForm.Width Then .Height = UserControl.ScaleHeight - HScroll.Height
        
        If .Height < UserControl.ScaleHeight Then
            HScroll.Left = 0
            HScroll.Top = UserControl.ScaleHeight - HScroll.Height
            HScroll.Width = .Width
            
            SetHScroll
            HScroll.Visible = True
        Else
            PicForm.Left = 0
            HScroll.Visible = False
        End If
        If .Width < UserControl.ScaleWidth Then
            VScroll.Left = UserControl.ScaleWidth - VScroll.Width
            VScroll.Top = 0
            VScroll.Height = .Height
            
            SetVScroll
            VScroll.Visible = True
        Else
            PicForm.Top = 0
            VScroll.Visible = False
        End If
    End With
End Sub

Public Property Get FormID() As String
    FormID = VisFormID
End Property
'
'Public Property Let FormID(ByVal vNewValue As String)
'    Dim tmpCtrl As Control, ValidCtrl As Boolean
'    Dim FormWidth As Long, FormHeight As Long
'    On Error Resume Next
'
'    VisFormID = vNewValue
'
'    '清除控件
'    For Each tmpCtrl In UserControl.Controls
'        ValidCtrl = True
'        If tmpCtrl.Container.Name <> "PicForm" Or Not tmpCtrl.Visible Then ValidCtrl = False
'        If ValidCtrl Then
'        '注意线的处理
'            If UCase(tmpCtrl.Name) = "FRATABLE" Then
'                Unload UserControl.Controls("F1Book1")(tmpCtrl.Index)
'            End If
'            Unload tmpCtrl
'        End If
'    Next
'
'    ReadForm VisFormID, FormWidth, FormHeight
'    If FormWidth > 0 Then UserControl.Width = FormWidth
'    If FormHeight > 0 Then UserControl.Height = FormHeight
'End Property

Public Sub ReadForm(ByVal strVisFormID As String, Optional Template As Boolean = True, Optional ByVal sPatientID As String = "", _
    Optional ByVal sPageID As String = "", Optional ByVal iPatientType As Integer = 0, _
    Optional objProgBar As ProgressBar, Optional blnReplaced As Boolean = False, Optional ByVal blnMoved As Boolean = False)
    '功能：读取所见单
    '参数：blnReplaced 是否强制替换具有替换域属性的所见项

    Dim tmpCtrl As Control, ValidCtrl As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim Seq As Long
    Dim FormWidth As Long, FormHeight As Long
    Dim iItemLen As Integer, sItemFormat As String, sDefaultValue As String
    Dim strSQL As String
    On Error Resume Next
    
    VisFormID = strVisFormID
    PatientID = sPatientID
    CheckID = sPageID
    PatientType = iPatientType
    mblnMoved = blnMoved
    
    PicForm.Left = 0: PicForm.Top = 0
    
    CurrCtrlName = ""
    
    '清除控件
    For Each tmpCtrl In UserControl.Controls
        ValidCtrl = True
        If tmpCtrl.Container.Name <> "PicForm" Or Not tmpCtrl.Visible Then ValidCtrl = False
        If ValidCtrl Then
        '注意线的处理
            If UCase(tmpCtrl.Name) = "FRATABLE" Then
                Unload UserControl.Controls("F1Book1")(tmpCtrl.Index)
            End If
            Unload tmpCtrl
        End If
    Next

    On Error GoTo DBError
    If Len(VisFormID) = 0 Then Exit Sub
    
    FormWidth = 0: FormHeight = 0
    
    Seq = 0
    If Template Then
        strSQL = "Select a.*,b.表示法,b.类型,b.长度,b.小数,b.数值域,b.替换域,b.中文名,b.文字表述,b.空值文字,b.编码 As 所见项编码 From 病历所见单 a,诊治所见项目 b" & _
            " Where a.元素ID=[1] And a.控件号>=0 " & _
            " And a.所见项ID=b.ID(+) Order By a.控件号"
        Set rsTmp = OpenSQLRecord(strSQL, "查询所见单项目", VisFormID)
    Else
        strSQL = "Select a.*,b.表示法,b.类型,b.长度,b.小数,b.数值域,b.替换域,b.中文名,b.文字表述,b.空值文字,b.编码 As 所见项编码" & _
            " From 病人病历所见单 a,诊治所见项目 b Where a.病历ID=[1] And a.控件号>=0 " + _
            " And a.所见项ID=b.ID(+) Order By a.控件号"
        If mblnMoved Then
            strSQL = Replace(strSQL, "病人病历所见单", "H病人病历所见单")
        End If
        Set rsTmp = OpenSQLRecord(strSQL, "查询所见单项目", VisFormID)
    End If
    
    If Not objProgBar Is Nothing Then objProgBar.Min = 0: objProgBar.Value = 0: objProgBar.Max = IIf(rsTmp.EOF, 1, rsTmp.RecordCount)
    Do While Not rsTmp.EOF
        Select Case rsTmp("控件类")
            Case 1
                Load UserControl.Text1(UserControl.Text1.Count)
                With UserControl.Text1(UserControl.Text1.Count - 1)
                    .Text = rsTmp("标题")
                    .Top = rsTmp("行"): .Left = rsTmp("列"): .Width = rsTmp("宽"): .Height = rsTmp("高")
                    .Alignment = rsTmp("对齐")
                    .Visible = True
                End With
                Set tmpCtrl = UserControl.Text1(UserControl.Text1.Count - 1)
            Case 9
                Load UserControl.Line1(UserControl.Line1.Count)
                With UserControl.Line1(UserControl.Line1.Count - 1)
                    .Top = rsTmp("行"): .Left = rsTmp("列"): .Width = rsTmp("宽"): .Height = rsTmp("高")
                    .Visible = True
                End With
                Set tmpCtrl = UserControl.Line1(UserControl.Line1.Count - 1)
            Case 2
                If Not IsNull(rsTmp("表示法")) Then
                    If Len(CurrCtrlName) = 0 Then
                        CurrCtrlName = "VisItem1": CurrCtrlIndex = UserControl.VisItem1.Count
                    End If
                    
                    '处理特殊所见项的长度和表示格式
                    If Template Then
                        sDefaultValue = IIf(IsNull(rsTmp("缺省内容")), "", rsTmp("缺省内容"))
                    Else
                        sDefaultValue = IIf(IsNull(rsTmp("所见内容")), "", rsTmp("所见内容"))
                    End If
                    Select Case rsTmp("所见项编码")
                        Case "101001" '年(YYYY)
                            iItemLen = 4: sItemFormat = "YYYY"
                        Case "101002" '月(YYYY-MM)
                            iItemLen = 7: sItemFormat = "YYYY-MM"
                        Case "101003" '日(YYYY-MM-DD)
                            iItemLen = 10: sItemFormat = "YYYY-MM-DD"
                        Case "101004" '时间(YYYY-MM-DD HH24:MI:SS)
                            iItemLen = 19: sItemFormat = "YYYY-MM-DD HH:MM:SS"
                        Case "101005" '时间(HH24:MI:SS)
                            iItemLen = 8: sItemFormat = "HH:MM:SS"
                        Case "101006" '时间(HH24:MI)
                            iItemLen = 5: sItemFormat = "HH:MM"
                        Case "101008" '当前日期
                            iItemLen = 10: sItemFormat = "YYYY-MM-DD"
                            sDefaultValue = Format(Date, sItemFormat)
                        Case "101009" '当前时间
                            iItemLen = 19: sItemFormat = "YYYY-MM-DD HH:MM:SS"
                            sDefaultValue = Format(zlDatabase.Currentdate, sItemFormat)
                        Case Else
                            iItemLen = IIf(IsNull(rsTmp("长度")), 10, rsTmp("长度"))
                            sItemFormat = ""
                    End Select
                    
                    Load UserControl.VisItem1(UserControl.VisItem1.Count)
                    With UserControl.VisItem1(UserControl.VisItem1.Count - 1)
                        If Template Then
                            .Init IIf(IsNull(rsTmp("标题")), "", rsTmp("标题")), IIf(IsNull(rsTmp("计量单位")), "", rsTmp("计量单位")), rsTmp("表示法"), rsTmp("类型"), iItemLen, IIf(IsNull(rsTmp("小数")), 0, rsTmp("小数")), IIf(IsNull(rsTmp("数值域")), "", rsTmp("数值域")), sDefaultValue, rsTmp("所见项ID"), IIf(IsNull(rsTmp("替换域")), "", IIf(rsTmp("替换域") = 1, rsTmp("中文名"), "")), IIf(IsNull(rsTmp("文字表述")), 1, rsTmp("文字表述")), IIf(IsNull(rsTmp("空值文字")), "", rsTmp("空值文字")), sItemFormat
                        Else
                            .Init IIf(IsNull(rsTmp("标题")), "", rsTmp("标题")), IIf(IsNull(rsTmp("计量单位")), "", rsTmp("计量单位")), rsTmp("表示法"), rsTmp("类型"), iItemLen, IIf(IsNull(rsTmp("小数")), 0, rsTmp("小数")), IIf(IsNull(rsTmp("数值域")), "", rsTmp("数值域")), sDefaultValue, rsTmp("所见项ID"), IIf(IsNull(rsTmp("替换域")), "", IIf(rsTmp("替换域") = 1, rsTmp("中文名"), "")), IIf(IsNull(rsTmp("文字表述")), 1, rsTmp("文字表述")), IIf(IsNull(rsTmp("空值文字")), "", rsTmp("空值文字")), sItemFormat
                        End If
                        
                        'Begin by CFR,2005-06-10
                        Set .ParentObject = mobjParentObject
                        'End by CFR
                        
                        '处理替换域
                        If Not IsNull(rsTmp("替换域")) Then
                            If rsTmp("替换域") = 1 And (Len(.Value) = 0 Or blnReplaced) Then .Value = GetSpecValue(rsTmp("中文名"), PatientID, CheckID, PatientType)
                        End If
                        
                        .Left = rsTmp("列"): .Top = rsTmp("行")
                        .Enabled = IIf(rsTmp("不可写") = 0, True, False)
'                        .Enabled = IIf(Len(.ExchangeField) = 0, IIf(rsTmp("不可写") = 0, True, False), False)
                        .AllowMask = IIf(IsNull(rsTmp("可屏蔽")), False, IIf(rsTmp("可屏蔽") = 0, False, True))
                        If rsTmp("不可写") = 0 Then .TabIndex = Seq
                        .Width = rsTmp("宽"): .Height = rsTmp("高")
                        .Visible = True
                    End With
                    Set tmpCtrl = UserControl.VisItem1(UserControl.VisItem1.Count - 1)
                    If rsTmp("不可写") = 0 Then Seq = Seq + 1
                End If
            Case 3
                If Len(CurrCtrlName) = 0 Then
                    CurrCtrlName = "F1Book1": CurrCtrlIndex = UserControl.F1Book1.Count
                End If
                
                Load UserControl.F1Book1(UserControl.F1Book1.Count)
                InitTable UserControl.F1Book1(UserControl.F1Book1.Count - 1)
                
                Load UserControl.fraTable(UserControl.fraTable.Count)
                Set UserControl.F1Book1(UserControl.F1Book1.Count - 1).Container = UserControl.fraTable(UserControl.fraTable.Count - 1)
                With UserControl.fraTable(UserControl.fraTable.Count - 1)
                    .Top = rsTmp("行"): .Left = rsTmp("列"): .Width = rsTmp("宽"): .Height = rsTmp("高")
                    .Visible = True
                End With
                With UserControl.F1Book1(UserControl.F1Book1.Count - 1)
                    .Left = 0: .Top = 0
                    .Width = UserControl.fraTable(UserControl.fraTable.Count - 1).Width
                    .Height = UserControl.fraTable(UserControl.fraTable.Count - 1).Height
                    .TabIndex = Seq
                        
                    .SetSelection 1, 1, .MaxRow, .MaxCol
                    .WordWrap = True
                    .SetSelection 1, 1, 1, 1
                    
                    .EnableProtection = True
                    
                    .Visible = True
                End With
                If Template Then
                    ReadTable UserControl.F1Book1(UserControl.F1Book1.Count - 1), VisFormID, rsTmp("控件号")
                Else
                    ReadTable_Patient UserControl.F1Book1(UserControl.F1Book1.Count - 1), VisFormID, rsTmp("控件号")
                End If
                Proc_Table_TopLeftChanged UserControl.F1Book1(UserControl.F1Book1.Count - 1)
                
                Set tmpCtrl = UserControl.fraTable(UserControl.fraTable.Count - 1)
                Seq = Seq + 1
        End Select
        If Not tmpCtrl Is Nothing Then
            If FormWidth < tmpCtrl.Left + tmpCtrl.Width + 30 Then FormWidth = tmpCtrl.Left + tmpCtrl.Width + 30
            If FormHeight < tmpCtrl.Top + tmpCtrl.Height + 30 Then FormHeight = tmpCtrl.Top + tmpCtrl.Height + 30
        End If
                    
        If Not objProgBar Is Nothing Then objProgBar.Value = rsTmp.AbsolutePosition
            
        rsTmp.MoveNext
    Loop
    
    If FormWidth > 0 Then
        PicForm.Width = FormWidth
        UserControl.Width = FormWidth + 10 '+ VScroll.Width
    End If
    If FormHeight > 0 Then
        PicForm.Height = FormHeight
        UserControl.Height = FormHeight + 10 ' + HScroll.Height
    End If
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub

Public Property Get Enabled() As Boolean
    Enabled = bEnabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    Dim tmpCtrl As Control
    bEnabled = vNewValue
    
    On Error Resume Next
    For Each tmpCtrl In UserControl.Controls
        Select Case UCase(tmpCtrl.Name)
            Case "VISITEM1", "VISITEM"
                tmpCtrl.AllowEdit = bEnabled
            Case "F1BOOK1"
                tmpCtrl.Enabled = bEnabled
        End Select
    Next
End Property

Private Sub UserControl_Terminate()
    
    On Error Resume Next
    
    Set mobjParentObject = Nothing
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Font", UserControl.Font
End Sub

Private Sub VisItem_GotFocus(Index As Integer)
    Dim aCellInfo() As String

    On Error Resume Next
    aCellInfo = Split(VisItem(Index).Tag, ",")
    
    F1Book1(CInt(aCellInfo(2))).SetActiveCell aCellInfo(0), aCellInfo(1)
End Sub

Private Sub VisItem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim aCellInfo() As String
    
    On Error Resume Next
    If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then
        aCellInfo = Split(VisItem(Index).Tag, ",")
        F1Book1(CInt(aCellInfo(2))).SetFocus
        zlCommFun.PressKey CByte(KeyCode)
    End If
End Sub

Private Sub VisItem1_GotFocus(Index As Integer)
    On Error Resume Next
    CurrCtrlName = "VisItem1": CurrCtrlIndex = Index

    ShowVisItem VisItem1(Index)
End Sub

Private Sub VisItem1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim NextCtrl As Control
    On Error Resume Next
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
        Set NextCtrl = NextElement(VisItem1(Index))
        If Not NextCtrl Is Nothing Then
            NextCtrl.SetFocus
        Else
            RaiseEvent NextControl
        End If
    End If
End Sub

Private Function NextElement(ctrlElement As Control) As Control
    Dim tmpCtrl As Control, bCurrCtrl As Boolean
    Set NextElement = Nothing
    bCurrCtrl = False
    
    For Each tmpCtrl In UserControl.Controls
        If InStr(",VISITEM1,F1BOOK1,", "," + UCase(tmpCtrl.Name) + ",") > 0 Then
            If tmpCtrl.Index > 0 Then
                If bCurrCtrl And tmpCtrl.Enabled Then
                    Set NextElement = tmpCtrl
                    Exit For
                Else
                    If tmpCtrl.Name = ctrlElement.Name And tmpCtrl.Index = ctrlElement.Index Then bCurrCtrl = True
                End If
            End If
        End If
    Next
End Function

Public Sub SaveForm(ElementID As String, cnOracle As ADODB.Connection, ErrorNumber As Long, ErrorMsg As String, Optional objProgBar As ProgressBar)
    Dim tmpCtrl As Control, ValidCtrl As Boolean
    Dim TabIndex As Long, Seq As Long, aTmp() As String
    Dim i As Integer

    On Error GoTo SaveError
    ErrorNumber = 0
    ErrorMsg = ""
    
    If Not objProgBar Is Nothing Then objProgBar.Min = 0: objProgBar.Value = 0: objProgBar.Max = UserControl.Controls.Count
    i = 0
    For Each tmpCtrl In UserControl.Controls
        ValidCtrl = False
        
        If InStr("TEXT1,LINE1,VISITEM1,FRATABLE", UCase(tmpCtrl.Name)) > 0 Then
            If tmpCtrl.Index > 0 Then ValidCtrl = True
        End If
        i = i + 1
        
        If ValidCtrl Then
            Select Case UCase(tmpCtrl.Name)
                Case "TEXT1"
                    Seq = tmpCtrl.TabIndex + 1
                    cnOracle.Execute "ZL_病人病历所见单_SAVE(" & ElementID & "," & Seq & ",'1','" + Replace(tmpCtrl.Text, "'", "''") + "'," & _
                    tmpCtrl.Top & "," & tmpCtrl.Left & "," & tmpCtrl.Width & "," & tmpCtrl.Height & "," & tmpCtrl.Alignment & "," & _
                    0 & ",0,'','','','')", , adCmdStoredProc
                Case "LINE1"
                    Seq = tmpCtrl.TabIndex + 1
                    cnOracle.Execute "ZL_病人病历所见单_SAVE(" & ElementID & "," & Seq & ",'9',''," & _
                    tmpCtrl.Top & "," & tmpCtrl.Left & "," & tmpCtrl.Width & "," & tmpCtrl.Height & ",0," & _
                    0 & ",0,'','','','')", , adCmdStoredProc
                Case "VISITEM1" '项目ID
                    Seq = tmpCtrl.TabIndex + 1
                    cnOracle.Execute "ZL_病人病历所见单_SAVE(" & ElementID & "," & Seq & ",'2','" + Replace(tmpCtrl.Title, "'", "''") + "'," & _
                    tmpCtrl.Top & "," & tmpCtrl.Left & "," & tmpCtrl.Width & "," & tmpCtrl.Height & ",0," & _
                    IIf(tmpCtrl.Enabled, 0, 1) & "," & IIf(tmpCtrl.AllowMask, 1, 0) & ",'" & tmpCtrl.ID & "','" & tmpCtrl.ItemType & "','" + tmpCtrl.Unit + "','" + Replace(tmpCtrl.Value, "'", "''") + "')", , adCmdStoredProc
                Case "FRATABLE" '元素ID
                    Seq = F1Book1(tmpCtrl.Index).TabIndex + 1
                    gcnOracle.Execute "ZL_病人病历所见单_SAVE(" & ElementID & "," & Seq & ",'3',''," & _
                    tmpCtrl.Top & "," & tmpCtrl.Left & "," & tmpCtrl.Width & "," & tmpCtrl.Height & ",0," & _
                    0 & ",0,'','','','" & ElementID & "')", , adCmdStoredProc

                    SaveTable_Patient ElementID, F1Book1(tmpCtrl.Index), gcnOracle, Seq
            End Select
        End If
        If Not objProgBar Is Nothing Then objProgBar.Value = i
    Next
    Exit Sub
SaveError:
    ErrorNumber = Err.Number
    ErrorMsg = Err.Description
End Sub

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "显示字体"
Attribute Font.VB_ProcData.VB_Invoke_Property = ";外观"
    Set Font = UserControl.Font
    
    SetFont
End Property

Public Property Set Font(ByVal vNewValue As StdFont)
    Set UserControl.Font = vNewValue
    
    SetFont
End Property

Private Sub SetFont()
    Dim tmpCtrl As Control
    
    On Error Resume Next
    For Each tmpCtrl In UserControl.Controls
        If InStr("F1BOOK1,TEXT1,VISITEM1", UCase(tmpCtrl.Name)) > 0 Then
            Select Case UCase(tmpCtrl.Name)
                Case "F1BOOK1"
                    tmpCtrl.DefaultFontName = UserControl.Font.Name
                    tmpCtrl.DefaultFontSize = -1 * ((UserControl.Font.Size / 72) * 1440)
                Case Else
                    Set tmpCtrl.Font = UserControl.Font
            End Select
        End If
    Next
End Sub

Private Sub SetHScroll()
    On Error Resume Next
    With HScroll
        .Min = 0
        .Max = PicForm.Width - PicMain.Width
        .SmallChange = PicMain.Width / 10
        .LargeChange = PicMain.Width
    End With
End Sub

Private Sub SetVScroll()
    On Error Resume Next
    With VScroll
        .Min = 0
        .Max = PicForm.Height - PicMain.Height
        .SmallChange = PicMain.Height / 10
        .LargeChange = PicMain.Height
    End With
End Sub

Private Sub VScroll_Change()
    On Error Resume Next
    PicForm.Top = -1 * VScroll.Value
End Sub

Private Sub ShowVisItem(ItemCtrl As Control)
    With ItemCtrl
        If .Left + PicForm.Left < 0 Or .Left + .Width + PicForm.Left > PicMain.Width Then HScroll.Value = IIf(.Left > HScroll.Max, HScroll.Max, .Left)
        If .Top + PicForm.Top < 0 Or .Top + .Height + PicForm.Top > PicMain.Height Then VScroll.Value = IIf(.Top > VScroll.Max, VScroll.Max, .Top)
    End With
End Sub

Public Property Get Text() As String
    Dim tmpCtrl As Control, ValidCtrl As Boolean

    Text = ""
    For Each tmpCtrl In UserControl.Controls
        ValidCtrl = False
        
        If InStr("VISITEM1", UCase(tmpCtrl.Name)) > 0 Then
            If tmpCtrl.Index > 0 Then ValidCtrl = True
        End If
        
        If ValidCtrl Then
            Select Case UCase(tmpCtrl.Name)
                Case "TEXT1"
                Case "VISITEM1" '项目ID
                    If Len(Trim(tmpCtrl.Value)) = 0 Then
                        If Len(tmpCtrl.NullString) > 0 Then Text = Text + " " + tmpCtrl.NullString
                    Else
                        Select Case tmpCtrl.TextMethod
                            Case 2
                                Text = Text + " " + tmpCtrl.Value + tmpCtrl.Unit + tmpCtrl.Title
                            Case 3
                                Text = Text + " " + tmpCtrl.Value + tmpCtrl.Unit
                            Case Else
                                Text = Text + " " + tmpCtrl.Title + "：" + tmpCtrl.Value + tmpCtrl.Unit
                        End Select
                    End If
                Case "FRATABLE" '元素ID
            End Select
        End If
    Next
    
    If Len(Text) > 0 Then Text = Mid(Text, 2)
End Property
'返回附加表元素的所见项
Public Property Get VisItem() As Object
    Set VisItem = UserControl.VisItem
End Property

Public Property Get Controls() As Object
    Set Controls = UserControl.Controls
End Property

