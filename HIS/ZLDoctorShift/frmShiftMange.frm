VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmShiftMange 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "班次管理"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6735
   Icon            =   "frmShiftMange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   6735
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdTypeOK 
      Appearance      =   0  'Flat
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3960
      TabIndex        =   2
      Top             =   3840
      Width           =   1100
   End
   Begin VB.CommandButton cmdTypeCancel 
      Appearance      =   0  'Flat
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5355
      TabIndex        =   1
      Top             =   3840
      Width           =   1100
   End
   Begin VB.ComboBox cboDept 
      Height          =   300
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   5505
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfType 
      Height          =   2655
      Left            =   960
      TabIndex        =   3
      Top             =   600
      Width           =   5490
      _cx             =   9684
      _cy             =   4683
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmShiftMange.frx":6852
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
      Editable        =   2
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
      Left            =   240
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShiftMange.frx":6934
            Key             =   "unCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShiftMange.frx":6ECE
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShiftMange.frx":7468
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShiftMange.frx":DCCA
            Key             =   "add"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShiftMange.frx":1452C
            Key             =   "Person"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShiftMange.frx":14F3E
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShiftMange.frx":1B7A0
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShiftMange.frx":1C1B2
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "时间格式：如18:00          24:00请用00:00表示"
      Height          =   180
      Left            =   960
      TabIndex        =   6
      Top             =   3360
      Width           =   4050
   End
   Begin VB.Label lblDept 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "科    室"
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   300
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "班次信息"
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   720
   End
End
Attribute VB_Name = "frmShiftMange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrDeptId As String
Private mlngDeptID As Long
Private mblnOk As Boolean

Public Function ShowMe(ByVal strDeptID As String, ByVal lngDeptId As Long) As Boolean

    mstrDeptId = strDeptID
    mlngDeptID = lngDeptId
    Me.Show 1
    ShowMe = mblnOk
End Function

Private Sub cboDept_Click()
    mlngDeptID = cboDept.ItemData(cboDept.ListIndex)
    Call LoadData
End Sub

Private Sub cmdTypeCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdTypeOK_Click()
    Dim i As Long
    Dim arrSQL() As Variant, varTemp As Variant
    Dim strTemp As String
    Dim blnBegin As Boolean
    Dim lngDeptId As Long
    
    arrSQL = Array()
    If CheckTypeData = False Then Exit Sub
    lngDeptId = cboDept.ItemData(cboDept.ListIndex)
    With vsfType
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("班次名称")) <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                If .TextMatrix(i, .ColIndex("原班次名称")) = "" Then
                    arrSQL(UBound(arrSQL)) = "Zl_医生值班班次_Edit(0," & lngDeptId & ",'" & .TextMatrix(i, .ColIndex("班次名称")) & "',null," & _
                        "to_date('" & Format(.TextMatrix(i, .ColIndex("开始时间")), "hh:mm") & "','hh24:mi'),to_date('" & Format(.TextMatrix(i, .ColIndex("结束时间")), "hh:mm") & "','hh24:mi'))"
                Else
                    arrSQL(UBound(arrSQL)) = "Zl_医生值班班次_Edit(1," & lngDeptId & ",'" & .TextMatrix(i, .ColIndex("班次名称")) & "','" & _
                        .TextMatrix(i, .ColIndex("原班次名称")) & "',to_date('" & Format(.TextMatrix(i, .ColIndex("开始时间")), "hh:mm") & "','hh24:mi'),to_date('" & Format(.TextMatrix(i, .ColIndex("结束时间")), "hh:mm") & "','hh24:mi'))"
                End If
            End If
        Next
    End With
    gcnOracle.BeginTrans
    blnBegin = True
    On Error GoTo ErrHand
    If vsfType.Tag <> "" Then
        '如果界面删除已经有的班次，需先删除
        varTemp = Split(vsfType.Tag, "<分隔符>")
        For i = 0 To UBound(varTemp)
            strTemp = "Zl_医生值班班次_Edit(2," & lngDeptId & ",'" & varTemp(i) & "')"
            If strTemp <> "" Then
                Call zlDatabase.ExecuteProcedure(strTemp, Me.Caption)
            End If
        Next
    End If
    For i = LBound(arrSQL) To UBound(arrSQL)
        strTemp = arrSQL(i)
        If strTemp <> "" Then
            Call zlDatabase.ExecuteProcedure(strTemp, Me.Caption)
        End If
    Next
    gcnOracle.CommitTrans
    mblnOk = True
    Unload Me
    Exit Sub
ErrHand:
    If blnBegin Then gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    
    mblnOk = False
    Set rsTemp = GetDeptName(mstrDeptId)
    Call zlControl.CboAddData(cboDept, rsTemp)
    For i = 0 To cboDept.ListCount - 1
        If cboDept.ItemData(i) = mlngDeptID Then
            cboDept.ListIndex = i
            Exit For
        Else
            If i = cboDept.ListCount - 1 Then
                cboDept.ListIndex = 0
            End If
        End If
    Next
    Call LoadData
End Sub

Private Sub LoadData()
'加载表格的班次数据
    Dim rsTemp As ADODB.Recordset

    Set rsTemp = GetShiftType(1, mlngDeptID)
    vsfType.Redraw = flexRDNone
    vsfType.Rows = rsTemp.RecordCount + 1
    Do While Not rsTemp.EOF
        vsfType.TextMatrix(rsTemp.AbsolutePosition, vsfType.ColIndex("原班次名称")) = rsTemp!班次名称
        vsfType.TextMatrix(rsTemp.AbsolutePosition, vsfType.ColIndex("班次名称")) = rsTemp!班次名称
        vsfType.TextMatrix(rsTemp.AbsolutePosition, vsfType.ColIndex("开始时间")) = rsTemp!开始时间 & ""
        vsfType.TextMatrix(rsTemp.AbsolutePosition, vsfType.ColIndex("结束时间")) = rsTemp!结束时间 & ""
        rsTemp.MoveNext
    Loop
    If vsfType.Rows = 1 Then
        vsfType.Rows = 2
    End If
    If vsfType.Rows > 7 Then
        vsfType.ColWidth(vsfType.ColIndex("班次名称")) = 1575
    Else
        vsfType.ColWidth(vsfType.ColIndex("班次名称")) = 1830
    End If
    Call vsfType_AfterRowColChange(2, 1, 1, 1)
    vsfType.Redraw = flexRDDirect
End Sub

Private Sub vsfType_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Or NewRow < 1 Then Exit Sub
    With vsfType
        .Cell(flexcpPicture, NewRow, .ColIndex("删除行")) = imgList.ListImages("delete").Picture
        .Cell(flexcpPicture, NewRow, .ColIndex("新增行")) = imgList.ListImages("add").Picture
        If OldRow < .Rows Then
            .Cell(flexcpPicture, OldRow, .ColIndex("删除行")) = ""
            .Cell(flexcpPicture, OldRow, .ColIndex("新增行")) = ""
        End If
    End With
End Sub

Private Sub vsfType_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsfType.ColIndex("新增行") Or Col = vsfType.ColIndex("删除行") Then Cancel = True
End Sub

Private Sub vsfType_Click()
    Dim strName As String
    
    With vsfType
        strName = .TextMatrix(.Row, .ColIndex("班次名称"))
        If .Col = .ColIndex("新增行") Then
            If strName <> "" And .TextMatrix(.Row, .ColIndex("开始时间")) <> "" And .TextMatrix(.Row, .ColIndex("结束时间")) <> "" Then
                .AddItem "", .Row + 1
                .Row = .Row + 1
                .Col = .ColIndex("班次名称")
                Call .ShowCell(.Row, .Col)
            End If
        ElseIf .Col = .ColIndex("删除行") Then
            Call vsfType_KeyDown(vbKeyDelete, 0)
        End If
    End With
End Sub

Private Sub vsfType_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strName As String
    
    With vsfType
        strName = .TextMatrix(.Row, .ColIndex("班次名称"))
        If KeyCode = vbKeyDelete Then
            If strName = "" And .TextMatrix(.Row, .ColIndex("开始时间")) = "" And .TextMatrix(.Row, .ColIndex("结束时间")) = "" Then
            Else
                If MsgBox("您确定删除" & IIf(strName = "", "第" & .Row & "行", "名称为【" & strName & "】") & "的班次信息吗?", vbDefaultButton2 + vbYesNo, Me.Caption) = vbNo Then
                    Exit Sub
                End If
            End If
            If .TextMatrix(.Row, .ColIndex("原班次名称")) <> "" Then
                vsfType.Tag = IIf(vsfType.Tag = "", "", vsfType.Tag & "<分隔符>") & strName
            End If
            If .Rows <= 2 Then
                .TextMatrix(1, .ColIndex("班次名称")) = ""
                .TextMatrix(1, .ColIndex("开始时间")) = ""
                .TextMatrix(1, .ColIndex("结束时间")) = ""
                .ShowCell .Row, .ColIndex("班次名称")
            Else
                .RemoveItem .Row
            End If
            Call vsfType_AfterRowColChange(0, 0, vsfType.Row, 1)
        End If
    End With
End Sub

Private Function CheckTypeData() As Boolean
'检查班次输入数据的合理性
    Dim i As Long
    Dim strNames As String, strName As String, strTemp As String
    Dim strBegins As String, strEnds As String, strBegin As String, strEnd As String
    
    With vsfType
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("班次名称")) <> "" Then
                strBegins = strBegins & "," & Format(.TextMatrix(i, .ColIndex("开始时间")), "HH:MM")
                strEnds = strEnds & "," & Format(.TextMatrix(i, .ColIndex("结束时间")), "HH:MM")
            End If
        Next
        For i = 1 To .Rows - 1
            strTemp = .TextMatrix(i, .ColIndex("班次名称"))
            If zlstr.ActualLen(strTemp) > 10 Then
                MsgBox "值班名称不得超过5个汉字，请检查！", vbExclamation, Me.Caption
                .ShowCell i, .ColIndex("班次名称")
                Exit Function
            End If
            If strTemp <> "" Then
                strBegin = Format(.TextMatrix(i, .ColIndex("开始时间")), "HH:MM")
                strEnd = Format(.TextMatrix(i, .ColIndex("结束时间")), "HH:MM")
                If Not CheckTime(strBegin, i, .Col) Then Exit Function
                If Not CheckTime(strEnd, i, .Col) Then Exit Function
                If InStr(strEnds & ",", "," & strBegin & ",") = 0 Then
                    MsgBox "【" & strTemp & "】的开始时间没有对应的相同结束时间，请检查！", vbExclamation, Me.Caption
                    .ShowCell i, .ColIndex("开始时间")
                    Exit Function
                End If
                If InStr(strBegins & ",", "," & strEnd & ",") = 0 Then
                    MsgBox "【" & strTemp & "】的结束时间没有对应的相同开始时间，请检查！", vbExclamation, Me.Caption
                    .ShowCell i, .ColIndex("结束时间")
                    Exit Function
                End If
                If InStr(strNames & ",", "," & strTemp & ",") = 0 Then
                    strNames = strNames & "," & strTemp
                Else
                    strName = IIf(strName = "", "", strTemp & "、")
                End If
                If strName <> "" Then
                    MsgBox "不允许出现重复班次名称：" & strName, vbExclamation, Me.Caption
                    Exit Function
                End If
            End If
        Next
    End With
    CheckTypeData = True
End Function
Private Function CheckTime(ByVal strDate As String, ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
'检查时间格式是否正确
    Dim varTemp As Variant
    
    If strDate = "" Then
        MsgBox "【" & strDate & "】时间不能为空，请输入！", vbExclamation, Me.Caption
        vsfType.ShowCell lngRow, lngCol
        Exit Function
    End If
    If InStr(strDate, ":") = 0 Then
        MsgBox "【" & strDate & "】时间没有冒号，请重新输入！", vbExclamation, Me.Caption
        vsfType.ShowCell lngRow, lngCol
        Exit Function
    End If
    varTemp = Split(strDate, ":")
    If varTemp(0) > 23 Then
        MsgBox "【" & strDate & "】的小时不能超过23，请重新输入！", vbExclamation, Me.Caption
        vsfType.ShowCell lngRow, lngCol
        Exit Function
    End If
    If Len(varTemp(0)) > 2 Then
        MsgBox "【" & strDate & "】的小时长度不能大于2位，请重新输入！", vbExclamation, Me.Caption
        vsfType.ShowCell lngRow, lngCol
        Exit Function
    End If
    If varTemp(1) > 59 Then
        MsgBox "【" & strDate & "】的分钟不能超过59，请重新输入！", vbExclamation, Me.Caption
        vsfType.ShowCell lngRow, lngCol
        Exit Function
    End If
    If Len(varTemp(1)) > 2 Then
        MsgBox "【" & strDate & "】的分钟长度不能大于2位，请重新输入！", vbExclamation, Me.Caption
        vsfType.ShowCell lngRow, lngCol
        Exit Function
    End If
    CheckTime = True
End Function

Private Sub vsfType_KeyPress(KeyAscii As Integer)

    With vsfType
        If .Col < .ColIndex("结束时间") Then
            .Col = .Col + 1
            .ShowCell .Row, .Col
        Else
            zlCommFun.PressKey (vbKeyTab)
        End If
    End With
End Sub

Private Sub vsfType_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsfType
        If Col = .ColIndex("开始时间") Or Col = .ColIndex("结束时间") Then
            If KeyAscii = vbKeyBack Then Exit Sub
            If KeyAscii = vbKeyReturn Then Exit Sub
            If KeyAscii = Asc("：") Then KeyAscii = Asc(":")
            
            If KeyAscii = Asc(":") And InStr(1, .EditText, ":") > 0 Then KeyAscii = 0: Exit Sub
            If KeyAscii = Asc(":") Then Exit Sub
        
            If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
                KeyAscii = 0
                Exit Sub
            End If
        ElseIf Col = .ColIndex("班次名称") Then
            If KeyAscii = Asc("'") Then KeyAscii = 0
        End If
    End With
End Sub


