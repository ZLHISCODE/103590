VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmDocShiftBase 
   Caption         =   "医生交接班内容设置"
   ClientHeight    =   8655
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13620
   Icon            =   "frmDocShiftBase.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   13620
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vsfPatiTypeInfo 
      Height          =   6975
      Left            =   3360
      TabIndex        =   1
      Top             =   1200
      Width           =   9255
      _cx             =   16325
      _cy             =   12303
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   18
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDocShiftBase.frx":5C02
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
   Begin VSFlex8Ctl.VSFlexGrid vsfPatiType 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   3075
      _cx             =   5424
      _cy             =   12303
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
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   400
      RowHeightMax    =   400
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDocShiftBase.frx":5EA5
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   8295
      Width           =   13620
      _ExtentX        =   24024
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDocShiftBase.frx":5F96
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21114
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   5640
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftBase.frx":682A
            Key             =   "unCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftBase.frx":6DC4
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftBase.frx":735E
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftBase.frx":DBC0
            Key             =   "add"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftBase.frx":14422
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftBase.frx":14E34
            Key             =   "CUp"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftBase.frx":152CE
            Key             =   "CMid"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftBase.frx":15768
            Key             =   "CDown"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftBase.frx":15C02
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftBase.frx":16614
            Key             =   "Person"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftBase.frx":17026
            Key             =   "Dept"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   2640
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgPublic 
      Bindings        =   "frmDocShiftBase.frx":1D888
      Left            =   3600
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmDocShiftBase.frx":1D89C
   End
   Begin VB.Label lblPatiTypeInfo 
      AutoSize        =   -1  'True
      Caption         =   "输入项目"
      Height          =   180
      Left            =   3360
      TabIndex        =   3
      Top             =   960
      Width           =   720
   End
   Begin VB.Label lblPatiType 
      AutoSize        =   -1  'True
      Caption         =   "病人类型"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   720
   End
End
Attribute VB_Name = "frmDocShiftBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrPrivs As String
Private mobjView As Object

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strSName As String, strPatiPrj As String
    Dim objControl As CommandBarControl
    Dim i As Long
    
    With vsfPatiType
        If .Rows > 1 Then
            strSName = .TextMatrix(.Row, .ColIndex("简称"))
        End If
    End With
    With vsfPatiTypeInfo
        If .Rows > 1 Then
            strPatiPrj = .TextMatrix(.Row, .ColIndex("项目名称"))
        End If
    End With
    Select Case Control.ID
        Case conMenu_DocShift_File_Preview
            Call PreView
        Case conMenu_DocShift_Edit_New
            Call PatiType(1, strSName)
        Case conMenu_DocShift_Edit_Modify
            Call PatiType(2, strSName)
        Case conMenu_DocShift_Edit_Delete
            If MsgBox("您确认删除简称为【" & strSName & "】的病人类型吗", vbInformation + vbDefaultButton2 + vbYesNo) = vbNo Then Exit Sub
            On Error GoTo errH
            gstrSql = "Zl_医生交接班病人类型_Edit(3,'" & strSName & "')"
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            With vsfPatiType
                .RemoveItem .Row
                If .Row > 0 Then Call vsfPatiType_Click
            End With
        Case conMenu_DocShift_Edit_Reuse
            Call ReuseStop(4, strSName)
        Case conMenu_DocShift_Edit_Stop
            Call ReuseStop(5, strSName)
        Case conMenu_DocShift_Edit_NewProject
            Call PatiPrj(1, strSName, strPatiPrj)
        Case conMenu_DocShift_Edit_ModifyProject
            Call PatiPrj(2, strSName, strPatiPrj)
        Case conMenu_DocShift_Edit_DeleteProject
            If MsgBox("您确认删除项目名称为【" & strPatiPrj & "】的病人类型吗", vbInformation + vbDefaultButton2 + vbYesNo) = vbNo Then Exit Sub
            On Error GoTo errH
            gstrSql = "Zl_医生交接班病人项目_Edit(3,'" & strSName & "','" & strPatiPrj & "')"
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            vsfPatiTypeInfo.RemoveItem vsfPatiTypeInfo.Row
        Case conMenu_DocShift_Edit_RowProject
            '(取消)合并行实质就是改变序号
            Call AdjustNum
        Case conMenu_View_ToolBar_Button '工具栏
            For i = 2 To cbsMain.Count
                Control.Checked = Not Control.Checked
                Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
            Next
            Call Form_Resize
            Me.cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Text '按钮文字
            Control.Checked = Not Control.Checked
            For i = 2 To cbsMain.Count
                For Each objControl In Me.cbsMain(i).Controls
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                Next
            Next
            Me.cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Size '大图标
            Control.Checked = Not Control.Checked
            Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
            Me.cbsMain.RecalcLayout
        Case conMenu_View_StatusBar '状态栏
            Control.Checked = Not Control.Checked
            Me.stbThis.Visible = Not Me.stbThis.Visible
            Call Form_Resize
            Me.cbsMain.RecalcLayout
        Case conMenu_DocShift_Help_Web_Home
            Call zlHomePage(Me.hwnd)
        Case conMenu_DocShift_Help_Web_Mail
            Call zlMailTo(Me.hwnd)
        Case conMenu_DocShift_Help_About
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_DocShift_File_Exit
            Unload Me
    End Select
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, Me.Caption
End Sub

Private Sub PreView()
'预览界面效果
    If CreateObj(mobjView) Then
        Call mobjView.ShowViewShift(Me, vsfPatiType.TextMatrix(vsfPatiType.Row, vsfPatiType.ColIndex("简称")))
    End If
End Sub

Private Function CreateObj(ByRef objView As Object) As Boolean
'创建对象
        
    If objView Is Nothing Then
        On Error Resume Next
        Set objView = CreateObject("zl9DoctorShift.clsDoctorShift")
        err.Clear: On Error GoTo 0
        If objView Is Nothing Then
            MsgBox "zl9DoctorShift部件未创建成功！", vbInformation, gstrSysName
            Exit Function
        Else
            Call objView.InitDoctorShift(glngSys, gcnOracle)
        End If
    End If
    CreateObj = True
End Function

Private Sub ReuseStop(ByVal bytType As Byte, ByVal strName As String)
'bytType:4-启用;5-停用

    If MsgBox("您确认" & IIf(bytType = 4, "启用", "停用") & "简称为【" & strName & "】的病人类型吗", vbInformation + vbDefaultButton2 + vbYesNo) = vbNo Then Exit Sub
    On Error GoTo errH
    gstrSql = "Zl_医生交接班病人类型_Edit(" & bytType & ",'" & strName & "')"
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            
    With vsfPatiType
        If bytType = 4 Then
            .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = vbBlack
        Else
            .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = vbRed
        End If
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, Me.Caption
End Sub

Private Sub AdjustNum()
'合并行,只有相邻的才可以合并，合并行数不得超过3个
    Dim i As Long, lngRow As Long, lngNum As Long, lngFirstNum As Long, lngSelRow As Long, lngRows As Long
    Dim strName As String, strTemp As String
    Dim blnCancel As Boolean
    Dim arrSql As Variant
    
    Dim objControl As CommandBarControl
        
    Set objControl = cbsMain.FindControl(, conMenu_DocShift_Edit_RowProject)
    On Error GoTo errH
    With vsfPatiType
        strName = .TextMatrix(.Row, .ColIndex("简称"))
    End With
    arrSql = Array()
    With vsfPatiTypeInfo
        If objControl.Caption = "取消合并行" Then
            '选择行的序号
            lngNum = .TextMatrix(.Row, .ColIndex("序号"))
            For i = 1 To .Rows - 1
                blnCancel = False
                If .TextMatrix(i, .ColIndex("序号")) = .TextMatrix(.Row, .ColIndex("序号")) Then
                    blnCancel = True
                    If lngFirstNum = 0 Then lngFirstNum = i
                    .TextMatrix(i, .ColIndex("合并")) = ""
                End If
                '当前行是取消合并行
                If blnCancel Then
                    If i > lngFirstNum And lngFirstNum <> 0 Then
                        lngNum = lngNum + 1
                        ReDim Preserve arrSql(UBound(arrSql) + 1)
                        arrSql(UBound(arrSql)) = "Zl_医生交接班病人项目_Edit(4,'" & strName & "','" & _
                            .TextMatrix(i, .ColIndex("项目名称")) & "',''," & lngNum & ")"
                    End If
                Else
                    '选择行同序号第一行之后的序号都需调整，如果之后有合并行，序号需相同
                    If i > lngFirstNum And lngFirstNum <> 0 Then
                        If lngRow <> .TextMatrix(i, .ColIndex("序号")) Then
                            lngNum = lngNum + 1
                        End If
                        ReDim Preserve arrSql(UBound(arrSql) + 1)
                        arrSql(UBound(arrSql)) = "Zl_医生交接班病人项目_Edit(4,'" & strName & "','" & _
                            .TextMatrix(i, .ColIndex("项目名称")) & "',''," & lngNum & ")"
                    End If
                End If
                lngRow = .TextMatrix(i, .ColIndex("序号"))
            Next
        Else
            lngSelRow = .Row
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, .ColIndex("选择")) = flexChecked Then
                    lngNum = lngNum + 1
                    If lngNum > 2 Then
                        MsgBox "合并行不得超过2行，请重新选择！", vbInformation, Me.Caption
                        Exit Sub
                    End If
                    If lngRow = 0 Then
                        lngFirstNum = .TextMatrix(i, .ColIndex("序号"))
                    Else
                        If i - lngRow <> 1 Then
                            MsgBox "合并行必须是相邻行的数据，请检查！", vbInformation, Me.Caption
                            Exit Sub
                        End If
                        If lngRows <> .TextMatrix(i, .ColIndex("输入行数")) Then
                            MsgBox "合并行的行数必须相同，请检查！", vbInformation, Me.Caption
                            Exit Sub
                        End If
                    End If
                    lngRow = i
                    lngRows = .TextMatrix(i, .ColIndex("输入行数"))
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = "Zl_医生交接班病人项目_Edit(4,'" & strName & "','" & _
                        .TextMatrix(i, .ColIndex("项目名称")) & "',''," & lngFirstNum & ")"
                End If
            Next
            If lngNum < 2 Then
                MsgBox "合并行不得少于两行数据，请检查是否勾选！", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
    End With
    For i = 0 To UBound(arrSql)
        strTemp = arrSql(i)
        Call zlDatabase.ExecuteProcedure(strTemp, "调整序号")
    Next
    If objControl.Caption <> "取消合并行" Then
        Call vsfPatiType_Click
        vsfPatiTypeInfo.Row = lngSelRow
    End If
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, Me.Caption
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnStop As Boolean, blnPatiType As Boolean, blnPatiInfo As Boolean
    Dim blnPriv As Boolean
    
    With vsfPatiType
        If .Row > 0 Then
            blnPatiType = True
            blnStop = .Cell(flexcpForeColor, .Row, 1, .Row, .Cols - 1) = vbRed
        End If
    End With
    
    With vsfPatiTypeInfo
        If .Row > 0 Then
            blnPatiInfo = True
        End If
    End With
    blnPriv = CheckPrivs("增删改")
    Select Case Control.ID
        Case conMenu_DocShift_File_Preview
            Control.Enabled = blnPriv And blnPatiType And Not blnStop
        Case conMenu_DocShift_Edit_New, conMenu_DocShift_Edit_Delete
            Control.Enabled = blnPriv And blnPatiType
        Case conMenu_DocShift_Edit_Modify
            Control.Enabled = blnPriv And blnPatiType And Not blnStop
        Case conMenu_DocShift_Edit_Reuse
            Control.Enabled = blnPriv And blnStop
        Case conMenu_DocShift_Edit_Stop
            Control.Enabled = blnPriv And Not blnStop
        Case conMenu_DocShift_Edit_NewProject
            Control.Enabled = blnPriv And Not blnStop
        Case conMenu_DocShift_Edit_ModifyProject, conMenu_DocShift_Edit_DeleteProject
            Control.Enabled = blnPriv And blnPatiInfo And Not blnStop
        Case conMenu_DocShift_Edit_RowProject
            '如果选择行是合并行，则工具栏显示取消合并行，反之，则显示合并行
            '这样设置caption才能生效
            Control.Enabled = False
            If blnPatiInfo And Not blnStop Then
                If vsfPatiTypeInfo.TextMatrix(vsfPatiTypeInfo.Row, vsfPatiTypeInfo.ColIndex("合并")) = "" Then
                    Control.Caption = "合并行      "
                Else
                    Control.Caption = "取消合并行"
                End If
            End If
            Control.Enabled = blnPriv And blnPatiInfo And Not blnStop
    End Select
End Sub

Private Function CheckPrivs(ByVal strPrivs As String) As Boolean
'检查是否有指定的权限
    
    If InStr(";" & mstrPrivs & ";", ";" & strPrivs & ";") = 0 Then Exit Function
    CheckPrivs = True
End Function

Private Sub Form_Load()

    mstrPrivs = gstrPrivs
    Call InitCommandBar
    Call LoadData
    
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub PatiType(bytType As Byte, Optional ByVal strSName As String)
'bytType:1-新增；2-修改
    Dim rsTemp As ADODB.Recordset
    
    If frmDocShiftTypeEdit.ShowMe(bytType, strSName) Then
        Set rsTemp = rsPatiType(strSName)
        If rsTemp.RecordCount = 1 Then
            With vsfPatiType
                If bytType = 1 Then
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                End If
                .TextMatrix(.Row, .ColIndex("简称")) = rsTemp!简称
                .TextMatrix(.Row, .ColIndex("病人名称")) = rsTemp!名称
                .TextMatrix(.Row, .ColIndex("起始描述")) = rsTemp!起始描述 & ""
                .TextMatrix(.Row, .ColIndex("提取SQL")) = rsTemp!提取SQL & ""
            End With
            Call vsfPatiType_Click
        End If
    End If
End Sub

Private Sub PatiPrj(bytType As Byte, ByVal strSName As String, ByVal strPatiPrj As String)
'bytType:1-新增；2-修改
    Call frmDocShiftProEdit.ShowMe(bytType, strSName, strPatiPrj)
End Sub

Public Sub RefreshPrj(ByVal bytType As Byte)
'bytType:1-新增；2-修改
'由于项目可以保存后继续新增，故子界面操作成功主界面都需要刷新
    Dim lngRow As Long
    
    lngRow = vsfPatiTypeInfo.Row
    Call vsfPatiType_Click
    If bytType = 1 Then
        vsfPatiTypeInfo.Row = vsfPatiTypeInfo.Rows - 1
    Else
        vsfPatiTypeInfo.Row = lngRow
    End If
End Sub

Private Sub LoadData()
    Dim rsTemp As ADODB.Recordset
    
    Set rsTemp = GetPatiType
    With vsfPatiType
        .Redraw = flexRDNone
        .Rows = 1
        .Rows = rsTemp.RecordCount + 1
        Do While Not rsTemp.EOF
            .TextMatrix(rsTemp.AbsolutePosition, .ColIndex("顺序")) = Val(rsTemp!顺序)
            .TextMatrix(rsTemp.AbsolutePosition, .ColIndex("简称")) = rsTemp!简称
            .TextMatrix(rsTemp.AbsolutePosition, .ColIndex("病人名称")) = rsTemp!名称
            .TextMatrix(rsTemp.AbsolutePosition, .ColIndex("起始描述")) = rsTemp!起始描述 & ""
            .TextMatrix(rsTemp.AbsolutePosition, .ColIndex("提取SQL")) = rsTemp!提取SQL & ""
            If rsTemp!是否停用 = 1 Then
                .Cell(flexcpForeColor, rsTemp.AbsolutePosition, 0, rsTemp.AbsolutePosition, .Cols - 1) = vbRed
            End If
            rsTemp.MoveNext
        Loop
        If .Rows > 1 Then
            .Row = 1
            Call vsfPatiType_Click
        End If
        .Redraw = flexRDDirect
    End With
End Sub

Private Function GetPatiType() As ADODB.Recordset
'获取病人类型
    
    gstrSql = "Select 顺序, 简称, 名称, 起始描述, 提取sql, 是否停用 From 医生交接班病人类型 Order By 顺序"
    Set GetPatiType = zlDatabase.OpenSQLRecord(gstrSql, "获取病人类型")
End Function

Private Sub Form_Resize()
    Dim lngTop As Long
    Dim lngHeight As Long
    
    On Error Resume Next
    
    If Not cbsMain(2).Visible Then
        lngTop = 500
    End If
    If stbThis.Visible Then
        lngHeight = stbThis.Height
    End If
    
    lblPatiType.Move 120, 1000 - lngTop
    vsfPatiType.Move 120, lblPatiType.Top + lblPatiType.Height + 100, 3075, Me.ScaleHeight - 1000 - lngHeight + lngTop - lblPatiType.Height - 100
    lblPatiTypeInfo.Move vsfPatiType.Left + vsfPatiType.Width + 150, lblPatiType.Top
    vsfPatiTypeInfo.Move lblPatiTypeInfo.Left, vsfPatiType.Top, Me.ScaleWidth - vsfPatiType.Width - 500, vsfPatiType.Height
End Sub

Private Sub vsfPatiType_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Or NewRow < 1 Then Exit Sub
    With vsfPatiType
        If NewRow = 1 Then
            If .Rows = 2 Then
                .Cell(flexcpPicture, NewRow, .ColIndex("上移")) = ""
                .Cell(flexcpPicture, NewRow, .ColIndex("下移")) = ""
            Else
                .Cell(flexcpPicture, NewRow, .ColIndex("上移")) = ""
                .Cell(flexcpPicture, NewRow, .ColIndex("下移")) = imgList.ListImages("Down").Picture
            End If
        Else
            If NewRow = .Rows - 1 Then
                .Cell(flexcpPicture, NewRow, .ColIndex("下移")) = ""
                .Cell(flexcpPicture, NewRow, .ColIndex("上移")) = imgList.ListImages("Up").Picture
            Else
                .Cell(flexcpPicture, NewRow, .ColIndex("上移")) = imgList.ListImages("Up").Picture
                .Cell(flexcpPicture, NewRow, .ColIndex("下移")) = imgList.ListImages("Down").Picture
            End If
        End If
        If OldRow < .Rows Then
            .Cell(flexcpPicture, OldRow, .ColIndex("上移")) = ""
            .Cell(flexcpPicture, OldRow, .ColIndex("下移")) = ""
        End If
    End With
End Sub

Private Sub vsfPatiType_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    With vsfPatiType
        If Not (.Col = .ColIndex("上移") Or .Col = .ColIndex("下移") Or .Col = .ColIndex("选择")) Then
            Cancel = True
        End If
    End With
End Sub

Private Sub vsfPatiType_Click()
    Dim strTemp As String
    Dim rsTemp As ADODB.Recordset
    Dim lngNum As Long, lngRow As Long, i As Long
    Dim blnAdd As Boolean
    Dim objControl As CommandBarControl
    
    On Error GoTo errH
    With vsfPatiType
        If vsfPatiType.Row <= 0 Then Exit Sub
        Set objControl = cbsMain.FindControl(, conMenu_DocShift_Edit_Modify)
        If objControl.Enabled Then
            If .Col = .ColIndex("上移") Then
                If Not .Cell(flexcpPicture, .Row, .ColIndex("上移")) Is Nothing Then
                    lngRow = .Row - 1
                End If
            ElseIf .Col = .ColIndex("下移") Then
                If Not .Cell(flexcpPicture, .Row, .ColIndex("下移")) Is Nothing Then
                    lngRow = .Row + 1
                End If
            End If
        End If
        If lngRow <> 0 Then
            '界面顺序不变，所以从1开始
            For i = 1 To .ColIndex("病人名称")
                strTemp = .TextMatrix(.Row, i)
                .TextMatrix(.Row, i) = .TextMatrix(lngRow, i)
                .TextMatrix(lngRow, i) = strTemp
            Next
            gstrSql = "Zl_医生交接班病人类型_Edit(6,'" & .TextMatrix(.Row, .ColIndex("简称")) & "','','','',''," & _
                Val(.TextMatrix(.Row, .ColIndex("顺序"))) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSql, "调整顺序")
            gstrSql = "Zl_医生交接班病人类型_Edit(6,'" & .TextMatrix(lngRow, .ColIndex("简称")) & "','','','',''," & _
                Val(.TextMatrix(lngRow, .ColIndex("顺序"))) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSql, "调整顺序")
            .Row = lngRow
        End If
    End With
    
    Set rsTemp = GetPatiTypeInfo(vsfPatiType.TextMatrix(vsfPatiType.Row, vsfPatiType.ColIndex("简称")))
    With vsfPatiTypeInfo
        .Redraw = flexRDNone
        .Rows = 1
        .Rows = rsTemp.RecordCount + 1
        Do While Not rsTemp.EOF
            lngRow = rsTemp.AbsolutePosition
            If rsTemp!序号 = lngNum Then
                '序号相同的需特出处理，界面显示为大括号括起来
                If blnAdd Then
                    If lngRow = .Rows - 1 Then
                        .TextMatrix(lngRow, .ColIndex("合并")) = "┗"
                    End If
                Else
                    If lngRow = .Rows - 1 Then
                        .TextMatrix(lngRow - 1, .ColIndex("合并")) = "┏"
                        .TextMatrix(lngRow, .ColIndex("合并")) = "┗"
                    Else
                        blnAdd = True
                        .TextMatrix(lngRow - 1, .ColIndex("合并")) = "┏"
                        .TextMatrix(lngRow, .ColIndex("合并")) = "┃"
                    End If
                End If
            Else
                If blnAdd Then
                    blnAdd = False
                    .TextMatrix(lngRow - 1, .ColIndex("合并")) = "┗"
                End If
            End If
            .TextMatrix(lngRow, .ColIndex("项目名称")) = rsTemp!项目名称
            .TextMatrix(lngRow, .ColIndex("序号")) = Val(rsTemp!序号 & "")
            .TextMatrix(lngRow, .ColIndex("项目类别")) = rsTemp!项目类别 & ""
            strTemp = rsTemp!输入形式 & ""
            strTemp = Mid(strTemp, InStr(strTemp, "-") + 1)
            .TextMatrix(lngRow, .ColIndex("输入形式")) = strTemp
            strTemp = rsTemp!输入类型 & ""
            strTemp = Mid(strTemp, InStr(strTemp, "-") + 1)
            .TextMatrix(lngRow, .ColIndex("输入类型")) = strTemp
            .TextMatrix(lngRow, .ColIndex("输入格式")) = rsTemp!输入格式 & ""
            .TextMatrix(lngRow, .ColIndex("输入值域")) = rsTemp!输入值域 & ""
            .TextMatrix(lngRow, .ColIndex("输入行数")) = rsTemp!输入行数 & ""
            strTemp = rsTemp!提取来源 & ""
            strTemp = Mid(strTemp, InStr(strTemp, "-") + 1)
            .TextMatrix(lngRow, .ColIndex("提取来源")) = strTemp
            .TextMatrix(lngRow, .ColIndex("提取病历")) = rsTemp!提取病历 & ""
            .TextMatrix(lngRow, .ColIndex("提取SQL")) = rsTemp!提取SQL & ""
            .TextMatrix(lngRow, .ColIndex("描述文字")) = rsTemp!描述文字 & ""
            .TextMatrix(lngRow, .ColIndex("是否只读")) = rsTemp!是否只读 & ""
            .TextMatrix(lngRow, .ColIndex("死亡则隐藏")) = rsTemp!死亡则隐藏 & ""
            '如果序号相同，则界面需要展现出来
            lngNum = Val(rsTemp!序号 & "")
            rsTemp.MoveNext
        Loop
        If .Rows > 1 Then .Row = 1
        .Redraw = flexRDDirect
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, Me.Caption
End Sub

Private Sub InitCommandBar()
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False

    Set cbsMain.Icons = imgPublic.Icons

    '菜单定义:包括公共部份
    '    请对xtpControlPopup类型的命令ID重新赋值
    '-----------------------------------------------------
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_DocShift_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_DocShift_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_File_Preview, "预览(&V)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_File_Exit, "退出(&X)"): objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_DocShift_PatiTypePopup, "病人类型(&E)", -1, False)
    objMenu.ID = conMenu_DocShift_PatiTypePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_New, "新增(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_Modify, "修改(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_Delete, "删除(&D)")
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_Reuse, "启用(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_Stop, "停用(&S)")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_DocShift_PatiProjectPopup, "病人项目(&E)", -1, False)
    objMenu.ID = conMenu_DocShift_PatiProjectPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_NewProject, "新增(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_ModifyProject, "修改(&C)")
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_DeleteProject, "删除(&S)")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_DocShift_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_DocShift_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_DocShift_View_ToolBar, "工具栏(&T)")
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_DocShift_View_ToolBar_Button, "标准按钮(&S)", -1, False)
            objControl.Checked = True
            Set objControl = .Add(xtpControlButton, conMenu_DocShift_View_ToolBar_Text, "文本标签(&T)", -1, False)
            objControl.Checked = True
            Set objControl = .Add(xtpControlButton, conMenu_DocShift_View_ToolBar_Size, "大图标(&B)", -1, False)
            objControl.Checked = True
        End With
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_View_StatusBar, "状态栏(&S)")
        objControl.Checked = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_DocShift_HelpPopup, "帮助(&H)", -1, False)
    objMenu.ID = conMenu_DocShift_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Help_Help, "帮助主题(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_DocShift_Help_Web, "&WEB上的" & gstrProductName)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_DocShift_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
            .Add xtpControlButton, conMenu_DocShift_Help_Web_Mail, "发送反馈(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Help_About, "关于(&A)…"): objControl.BeginGroup = True
    End With

    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagHideWrap
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_File_Preview, "预览")
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_New, "新增类型"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_Modify, "修改类型")
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_Delete, "删除类型")
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_Reuse, "启用(&R)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_Stop, "停用(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_NewProject, "新增项目"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_ModifyProject, "修改项目")
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_DeleteProject, "删除项目")
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_Edit_RowProject, "合并行      ")
        Set objControl = .Add(xtpControlButton, conMenu_DocShift_File_Exit, "退出"): objControl.BeginGroup = True
    End With
    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next

    '命令的快键绑定:公共部份主界面已处理
    '-----------------------------------------------------
    With cbsMain.KeyBindings
'
'        .Add FCONTROL, vbKeyA, conMenu_DocShift_Edit_New '新增
'        .Add FCONTROL, vbKeyM, conMenu_DocShift_Edit_Modify '修改
'        .Add 0, vbKeyDelete, conMenu_DocShift_Edit_Delete '删除
    End With
End Sub

Private Sub vsfPatiType_DblClick()
    Dim objControl As CommandBarControl
    
    With vsfPatiType
        If .MouseRow < 1 Then Exit Sub
        Set objControl = cbsMain.FindControl(, conMenu_DocShift_Edit_Modify)
        If objControl.Enabled = False Then Exit Sub
        Call PatiType(2, .TextMatrix(.Row, .ColIndex("简称")))
    End With
End Sub

Private Sub vsfPatiTypeInfo_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim blnCheck As Boolean
    Dim lngNum As Long, i As Long
    
    With vsfPatiTypeInfo
        If Col = .ColIndex("选择") Then
            blnCheck = .Cell(flexcpChecked, Row, Col) = flexChecked
            lngNum = .TextMatrix(Row, .ColIndex("序号"))
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("序号")) = lngNum Then
                    .Cell(flexcpChecked, i, .ColIndex("选择")) = IIf(blnCheck, flexChecked, flexUnchecked)
                End If
            Next
        End If
    End With
End Sub

Private Sub vsfPatiTypeInfo_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Long, lngNum As Long
    Dim blnUp As Boolean, blnDown As Boolean
    Dim lngBegin As Long, lngEnd As Long
    
    If OldRow = NewRow Or NewRow < 1 Then Exit Sub
    With vsfPatiTypeInfo
        lngNum = Val(.TextMatrix(NewRow, .ColIndex("序号")))
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("序号"))) = lngNum Then
                '记录合并行的第一行的行号
                If lngBegin = 0 Then lngBegin = i
                '记录合并行的最后一行的行号
                lngEnd = i
            End If
        Next
        '如果是合并行的数据，以相同序号的第一行来判断能否上移，以相同序号的最后一行判断能否下移
        If lngBegin = 1 Then
            If .Rows > 2 Then
                blnUp = False
            End If
        Else
            If lngBegin = .Rows - 1 Then
                blnUp = True
            Else
                blnUp = True
            End If
        End If
        If lngEnd = 1 Then
            If .Rows > 2 Then
                blnDown = True
            End If
        Else
            If lngEnd = .Rows - 1 Then
                blnDown = False
            Else
                blnDown = True
            End If
        End If
        .Cell(flexcpPicture, NewRow, .ColIndex("上移")) = IIf(blnUp, imgList.ListImages("Up").Picture, "")
        .Cell(flexcpPicture, NewRow, .ColIndex("下移")) = IIf(blnDown, imgList.ListImages("Down").Picture, "")
        If OldRow < .Rows Then
            .Cell(flexcpPicture, OldRow, .ColIndex("上移")) = ""
            .Cell(flexcpPicture, OldRow, .ColIndex("下移")) = ""
        End If
    End With
End Sub

Private Sub vsfPatiTypeInfo_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfPatiTypeInfo
        If Not (.Col = .ColIndex("上移") Or .Col = .ColIndex("下移") Or .Col = .ColIndex("选择")) Then
            Cancel = True
        End If
    End With
End Sub

Private Sub vsfPatiTypeInfo_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < vsfPatiTypeInfo.ColIndex("项目名称") Then Cancel = True
End Sub

Private Sub vsfPatiTypeInfo_Click()
    Dim lngRow As Long, i As Long, j As Long, lngChangeRow As Long, lngNum As Long, lngChangeNum As Long, m As Long
    Dim strPrj As String, strTemp As String
    Dim strName As String
    Dim blnUp As Boolean, blnDown As Boolean, blnFirst As Boolean, blnAddRow As Boolean
    Dim arrSql As Variant
    Dim objControl As CommandBarControl
        
    If vsfPatiTypeInfo.Row < 1 Then Exit Sub
    strName = vsfPatiType.TextMatrix(vsfPatiType.Row, vsfPatiType.ColIndex("简称"))
    With vsfPatiTypeInfo
        If .Row < 1 Then Exit Sub
        Set objControl = cbsMain.FindControl(, conMenu_DocShift_Edit_ModifyProject)
        If objControl.Enabled Then
            If .Col = .ColIndex("上移") Then
                '找出要交换的行的行号，由于涉及合并行，故上移取合并行第一行的上一行，下移取合并行最后一行的下一行
                If Not .Cell(flexcpPicture, .Row, .ColIndex("上移")) Is Nothing Then
                    For i = 1 To .Rows - 1
                        If .TextMatrix(i, .ColIndex("序号")) = .TextMatrix(.Row, .ColIndex("序号")) Then
                            lngRow = i - 1
                            Exit For
                        End If
                    Next
                    blnUp = True
                End If
            ElseIf .Col = .ColIndex("下移") Then
                If Not .Cell(flexcpPicture, .Row, .ColIndex("下移")) Is Nothing Then
                    For i = 1 To .Rows - 1
                        If .TextMatrix(i, .ColIndex("序号")) = .TextMatrix(.Row, .ColIndex("序号")) Then
                            lngRow = i + 1
                        End If
                    Next
                    blnDown = True
                End If
            End If
        End If
        If blnUp = False And blnDown = False Then Exit Sub
        
        lngChangeNum = .TextMatrix(lngRow, .ColIndex("序号")) '要上移或者下移后的序号
        lngNum = .TextMatrix(.Row, .ColIndex("序号")) '当前选择行的序号
        strPrj = .TextMatrix(.Row, .ColIndex("项目名称")) '选择行的名称
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("序号")) = lngChangeNum Then
                If blnUp Then
                    '上移取上移行同序号的第一行
                    lngChangeRow = i - 1
                    Exit For
                Else
                    '下移取下移行同序号的最后一行
                    lngChangeRow = i
                End If
            End If
        Next
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("序号")) = lngChangeNum Then
                '将要移动后的行的序号调整为-1
                .TextMatrix(i, .ColIndex("序号")) = -1
            End If
        Next
        '由于插入行时总行数会增加，故先算出插入了多少行
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("序号")) = lngNum Then
                m = m + 1
            End If
        Next
        '用插入行的方法来实现上移和下移
        For i = 1 To .Rows - 1 + m
            If .TextMatrix(i, .ColIndex("序号")) = lngNum Then
                '根据上面的算法得出的行号
                lngChangeRow = lngChangeRow + 1
                .AddItem "", lngChangeRow
                If blnFirst = False Then
                    blnFirst = True
                    '当插入行在选择行的上面时，保证数据的正确性，选择行应往下移动一行
                    If lngChangeRow < .Row Then
                        .Row = .Row + 1
                        blnAddRow = True
                    End If
                End If
                For j = .ColIndex("项目名称") To .ColIndex("死亡则隐藏")
                    .TextMatrix(lngChangeRow, j) = .TextMatrix(IIf(blnAddRow, i + 1, i), j)
                Next
                .TextMatrix(lngChangeRow, .ColIndex("合并")) = .TextMatrix(IIf(blnAddRow, i + 1, i), .ColIndex("合并"))
                .TextMatrix(lngChangeRow, .ColIndex("序号")) = lngChangeNum
                .TextMatrix(IIf(blnAddRow, i + 1, i), .ColIndex("序号")) = 0
            End If
        Next
        '调整回正确的序号
        For i = .Rows - 1 To 1 Step -1
            If .TextMatrix(i, .ColIndex("序号")) = -1 Then
                .TextMatrix(i, .ColIndex("序号")) = lngNum
            ElseIf .TextMatrix(i, .ColIndex("序号")) = 0 Then
                .RemoveItem i
            End If
        Next
        arrSql = Array()
        For i = 1 To .Rows - 1
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = "Zl_医生交接班病人项目_Edit(4,'" & strName & "','" & .TextMatrix(i, .ColIndex("项目名称")) & "',''," & _
            .TextMatrix(i, .ColIndex("序号")) & ")"
            If .TextMatrix(i, .ColIndex("项目名称")) = strPrj Then
                .Row = i
            End If
        Next
        On Error GoTo errH
        For i = 0 To UBound(arrSql)
            strTemp = arrSql(i)
            Call zlDatabase.ExecuteProcedure(strTemp, "调整序号")
        Next
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, Me.Caption
End Sub

Private Sub vsfPatiTypeInfo_DblClick()
    Dim objControl As CommandBarControl
    
    With vsfPatiTypeInfo
        If .MouseRow < 1 Then Exit Sub
        Set objControl = cbsMain.FindControl(, conMenu_DocShift_Edit_ModifyProject)
        If objControl.Enabled = False Then Exit Sub
        If .Col < .ColIndex("项目名称") Then Exit Sub
        Call PatiPrj(2, vsfPatiType.TextMatrix(vsfPatiType.Row, vsfPatiType.ColIndex("简称")), _
            .TextMatrix(.Row, .ColIndex("项目名称")))
    End With
End Sub

