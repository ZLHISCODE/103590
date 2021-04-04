VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExportScheme 
   Caption         =   "导出方案"
   ClientHeight    =   6390
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9990
   Icon            =   "frmExportScheme.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ImageList imgIcon 
      Left            =   6360
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExportScheme.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExportScheme.frx":D0B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkPicAll 
      Caption         =   "全选"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   5988
      Width           =   735
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "导出"
      Height          =   350
      Left            =   7560
      TabIndex        =   4
      Top             =   6000
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   350
      Left            =   8760
      TabIndex        =   3
      Top             =   6000
      Width           =   1100
   End
   Begin VB.CheckBox chkIcon 
      Caption         =   "导出图标资源"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   5988
      Width           =   3855
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   4560
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfScheme 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   9735
      _cx             =   17171
      _cy             =   9763
      Appearance      =   1
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
      BackColorBkg    =   -2147483636
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
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
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
   Begin VB.Image imgCheck 
      Height          =   255
      Left            =   3360
      Picture         =   "frmExportScheme.frx":13916
      Stretch         =   -1  'True
      Tag             =   "1"
      Top             =   6120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgNoCheck 
      Height          =   255
      Left            =   3120
      Picture         =   "frmExportScheme.frx":13C88
      Stretch         =   -1  'True
      Tag             =   "0"
      Top             =   6120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblSelect 
      AutoSize        =   -1  'True
      Caption         =   "选择需要导出的方案："
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1800
   End
End
Attribute VB_Name = "frmExportScheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private marrID() As Long   '导入方案的ID
Private mstrFile As String   '导入时路径
Private mblnIsExport As Boolean
Private mlngModuleNo As Long
Private mblnCancel As Boolean
Private mblnIcon As Boolean
Private Const M_STR_COLNAME = "序号|ID|选择|方案名称|方案说明"
Private Enum ColTitle
    ct序号 = 0
    ctID = 1
    ct选择 = 2
    ct方案名称 = 3
    ct方案说明 = 4
End Enum

Private Sub chkPicAll_Click()
    On Error GoTo errHandle
    
    chkPicAll.Caption = IIf(chkPicAll.Value = 1, "全清", "全选")
    If vsfScheme.Rows < 2 Then Exit Sub
    If chkPicAll.Value = 1 Then
        vsfScheme.Cell(flexcpData, 1, ColTitle.ct选择, vsfScheme.Rows - 1, ColTitle.ct选择) = 1
        vsfScheme.Cell(flexcpPicture, 1, ColTitle.ct选择, vsfScheme.Rows - 1, ColTitle.ct选择) = imgCheck.Picture
    Else
        vsfScheme.Cell(flexcpData, 1, ColTitle.ct选择, vsfScheme.Rows - 1, ColTitle.ct选择) = 0
        vsfScheme.Cell(flexcpPicture, 1, ColTitle.ct选择, vsfScheme.Rows - 1, ColTitle.ct选择) = imgNoCheck.Picture
    End If
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandle
    
    mblnCancel = False
    Unload Me
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cmdExport_Click()
    Dim arrID() As Long
    Dim i As Long
    Dim lngCount As Long
    Dim lngSchemeNum As Long
    
    On Error GoTo errHandle

    ReDim arrID(0)
    lngCount = 0
    lngSchemeNum = 0
    For i = 1 To vsfScheme.Rows - 1
        If vsfScheme.Cell(flexcpData, i, ColTitle.ct选择) = 1 Then
            If lngCount <> 0 Then
                ReDim Preserve arrID(UBound(arrID) + 1)
            End If
            arrID(UBound(arrID)) = vsfScheme.TextMatrix(i, ColTitle.ctID)
            lngCount = lngCount + 1
            lngSchemeNum = lngSchemeNum + 1
        End If
    Next
    
    If lngSchemeNum = 0 Then
        MsgBox "请先选择方案！", vbInformation, Me.Caption
        Exit Sub
    End If
    If mblnIsExport Then
        Call ExportScheme(arrID)
    Else
        marrID = arrID
    End If
    mblnIcon = chkIcon.Value
    mblnCancel = True
    Unload Me
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub Form_Load()
    On Error GoTo errHandle
    
    Call GridInit(M_STR_COLNAME, vsfScheme)
    Call InitInterFace
    Call InitSchemeList
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub InitSchemeList()
    Dim rsScheme As Recordset
    Dim strSql As String
    Dim rsData As Recordset
    Dim i As Long
    
    On Error GoTo errHandle
    
    If mblnIsExport Then
        strSql = "select rownum 序号,ID,'' as 选择,方案名称,方案说明 from 影像查询方案 where 所属模块 = [1] Order By 方案序号  "
        Set rsScheme = ExecuteSql(strSql, "查询方案信息", mlngModuleNo)
        Set vsfScheme.DataSource = rsScheme
        
        If rsScheme.RecordCount < 1 Then
            Exit Sub
        End If
        Call SchemeNo
    Else
    
        Set rsData = New ADODB.Recordset
        Call rsData.Open(mstrFile)
    
        If rsData.RecordCount <= 0 Then
            MsgBox "没有可用于导入的数据，请检查文件是否正确。", vbInformation, Me.Caption
            Exit Sub
        End If
        
        While Not rsData.EOF
            i = i + 1
            vsfScheme.Rows = vsfScheme.Rows + 1
            vsfScheme.TextMatrix(i, ColTitle.ct序号) = i
            vsfScheme.TextMatrix(i, ColTitle.ctID) = NVL(rsData.Fields!Id)
            vsfScheme.TextMatrix(i, ColTitle.ct方案名称) = NVL(rsData.Fields!方案名称)
            vsfScheme.TextMatrix(i, ColTitle.ct方案说明) = NVL(rsData.Fields!方案说明)
            rsData.MoveNext
        Wend
    End If
    
    vsfScheme.ColHidden(ColTitle.ctID) = True
    vsfScheme.Cell(flexcpData, 1, ColTitle.ct选择, vsfScheme.Rows - 1, ColTitle.ct选择) = 0
    vsfScheme.Cell(flexcpPicture, 1, ColTitle.ct选择, vsfScheme.Rows - 1, ColTitle.ct选择) = imgNoCheck.Picture
    vsfScheme.Cell(flexcpPictureAlignment, 1, ColTitle.ct选择, vsfScheme.Rows - 1, ColTitle.ct选择) = flexPicAlignCenterCenter
    vsfScheme.ColWidth(ColTitle.ct序号) = 480
    vsfScheme.ColWidth(ColTitle.ct选择) = 480
    vsfScheme.ColWidth(ColTitle.ct方案名称) = 2000
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub SchemeNo()
'调整方案序号
    Dim i As Long

    If vsfScheme.Rows < 2 Then Exit Sub
    For i = 1 To vsfScheme.Rows - 1
        vsfScheme.TextMatrix(i, ColTitle.ct序号) = i
    Next
End Sub

Private Sub ExportScheme(arrID() As Long)
'导出方案
    Dim objSqlScheme As clsSqlScheme
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim rsRecord As ADODB.Recordset
    Dim strSchemeText As String
    Dim strWhere As String
    Dim strPara As String
    Dim strIcon As String
    Dim strExIcon As String
    Dim arrRoad() As String
    Dim arrIconName() As String
    Dim strFile As String
    Dim strSchemeXml As String
    Dim i As Long
    Dim j As Long
    Dim m As Long
    Dim lngCount As Long
    
    dlgFile.Filter = "(*.XML)|*.XML|(*.*)|*.*"
    dlgFile.DefaultExt = "*.XML"

    dlgFile.FileName = ""
    dlgFile.ShowSave

    If dlgFile.FileName = "" Then Exit Sub
    For i = 0 To UBound(arrID)
        strWhere = strWhere & " or id ='" & arrID(i) & "'"
    Next

    strWhere = Mid(strWhere, 4)
    
    strSql = "select id, 方案名称,方案说明,'' as 方案内容" & _
            " from 影像查询方案 where (" & strWhere & ") and 所属模块 = [1]  order by 方案序号"
    Set rsData = ExecuteSql(strSql, "导出方案", mlngModuleNo)
    
    If rsData.RecordCount <= 0 Then
        MsgBox Me, "没有可用于导出的数据，请检查方案设置。", vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    rsData.MoveFirst
    While Not rsData.EOF
        strSchemeXml = ReadSchemeXml(rsData.Fields!Id, "")
        rsData.Fields!方案内容 = strSchemeXml
        rsData.MoveNext
    Wend
    
    '导出图标
    If chkIcon.Value = 1 Then
        rsData.MoveFirst
        While Not rsData.EOF
            strSchemeText = ""
            strSchemeText = strSchemeText & rsData!方案内容
            Set objSqlScheme = New clsSqlScheme
            Call objSqlScheme.OpenScheme(strSchemeText)
            For j = 1 To objSqlScheme.ShowCfgCount
                For m = 1 To objSqlScheme.ShowCfg(j).RowRelationCount
                    If Len(Trim(objSqlScheme.ShowCfg(j).RowRelation(m).Icon)) > 0 Then
                        If InStr(UCase(strIcon), UCase("[" & Trim(objSqlScheme.ShowCfg(j).RowRelation(m).Icon)) & "]") = 0 Then
                            strIcon = strIcon & ",[" & objSqlScheme.ShowCfg(j).RowRelation(m).Icon & "]"
                        End If
                    End If
                Next
                If Len(Trim(objSqlScheme.ShowCfg(j).Icon)) > 0 Then
                    If InStr(UCase(strIcon), UCase("[" & Trim(objSqlScheme.ShowCfg(j).Icon)) & "]") = 0 Then
                        strIcon = strIcon & ",[" & objSqlScheme.ShowCfg(j).Icon & "]"
                    End If
                End If
            Next
            Call rsData.MoveNext
        Wend
        strIcon = Mid(strIcon, 3)
        strIcon = Mid(strIcon, 1, Len(strIcon) - 1)
        arrIconName = Split(strIcon, "],[")
        arrRoad = Split(dlgFile.FileName, "\")

        lngCount = 0
        strFile = Replace(dlgFile.FileName, ".XML", "\")
        If Len(Dir(strFile)) > 0 Then
            If MsgBox("图标文件已存在,是否删除?", vbYesNo, Me.Caption) = vbYes Then
                Kill strFile
            End If
        End If

        MkDir strFile
        For i = 0 To UBound(arrIconName)
            Call zlBlobRead(arrIconName(i), strFile & "\" & arrIconName(i) & ".ico")
        Next
    End If
    
    Call rsData.Save(dlgFile.FileName, adPersistXML)
    
    MsgBox "已成功导出" & rsData.RecordCount & "条数据。", vbInformation, Me.Caption
    
    Unload Me
End Sub

Public Function ShowMe(lngModuleNo As Long, blnIsExport As Boolean, ByRef arrID() As Long, strFile As String, ByRef blnIcon As Boolean, ower As Object) As Boolean
    mlngModuleNo = lngModuleNo
    mblnIsExport = blnIsExport
    mstrFile = strFile
    
    Me.Show 1, ower
     
    If mblnCancel Then
        arrID = marrID
        blnIcon = mblnIcon
    End If
    ShowMe = mblnCancel
End Function

Private Sub InitInterFace()
    If mblnIsExport Then
        Me.Caption = "导出方案"
        lblSelect.Caption = "选择需要导出的方案："
        chkIcon.Caption = "导出图标资源"
        cmdExport.Caption = "导出"
        Me.Icon = imgIcon.ListImages(2).Picture
    Else
        Me.Caption = "导入方案"
        lblSelect.Caption = "选择需要导入的方案："
        chkIcon.Caption = "导入图标资源(如果存在该图标则自动忽略)"
        cmdExport.Caption = "导入"
        Me.Icon = imgIcon.ListImages(1).Picture
    End If
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    vsfScheme.Height = Me.ScaleHeight - vsfScheme.Top - cmdExport.Height - 120
    vsfScheme.Width = Me.ScaleWidth - vsfScheme.Left * 2
    
    chkPicAll.Top = vsfScheme.Top + vsfScheme.Height + 60
    chkIcon.Top = chkPicAll.Top
    
    cmdCancel.Top = chkPicAll.Top
    cmdCancel.Left = vsfScheme.Width + vsfScheme.Left - cmdCancel.Width
    
    cmdExport.Top = chkPicAll.Top
    cmdExport.Left = cmdCancel.Left - 60 - cmdExport.Width
End Sub

Private Sub vsfScheme_Click()
    Dim lngRow As Long
    Dim lngCol As Long

    On Error GoTo errHandle

    lngRow = vsfScheme.Row
    lngCol = vsfScheme.Col
    If lngRow > 0 Then
        If vsfScheme.Cell(flexcpData, lngRow, ColTitle.ct选择) = 1 Then
            vsfScheme.Cell(flexcpData, lngRow, ColTitle.ct选择) = 0
            vsfScheme.Cell(flexcpPicture, lngRow, ColTitle.ct选择) = imgNoCheck.Picture
        Else
            vsfScheme.Cell(flexcpPicture, lngRow, ColTitle.ct选择) = imgCheck.Picture
            vsfScheme.Cell(flexcpData, lngRow, ColTitle.ct选择) = 1
        End If
    End If

    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub
