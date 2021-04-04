VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmQueryFilter 
   BackColor       =   &H00E0E0E0&
   Caption         =   "查询过滤"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5625
   Icon            =   "frmQueryFilter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   5625
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cbxRange 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      ItemData        =   "frmQueryFilter.frx":000C
      Left            =   720
      List            =   "frmQueryFilter.frx":0031
      Style           =   2  'Dropdown List
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.Slider sdrRange 
      Height          =   210
      Left            =   2040
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1360
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   370
      _Version        =   393216
      Max             =   180
      TickStyle       =   3
      TickFrequency   =   3
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00E0E0E0&
      Caption         =   "恢 复(&C)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "取 消(&Q)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4296
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1185
   End
   Begin VB.CommandButton cmdSure 
      BackColor       =   &H00E0E0E0&
      Caption         =   "确 定(&S)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2928
      TabIndex        =   13
      Top             =   4920
      Width           =   1185
   End
   Begin VB.Frame framButton 
      BackColor       =   &H00E0E0E0&
      Height          =   795
      Left            =   -120
      TabIndex        =   9
      Top             =   4680
      Width           =   5895
   End
   Begin VB.ListBox lstObj 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Index           =   0
      ItemData        =   "frmQueryFilter.frx":0089
      Left            =   2088
      List            =   "frmQueryFilter.frx":008B
      Style           =   1  'Checkbox
      TabIndex        =   8
      Top             =   2496
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ComboBox cbxObj 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   0
      Left            =   2088
      TabIndex        =   7
      Text            =   "cbxObj"
      Top             =   2100
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox txtObj 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   302
      Index           =   0
      Left            =   2040
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ComboBox cbxWhere 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      ItemData        =   "frmQueryFilter.frx":008D
      Left            =   720
      List            =   "frmQueryFilter.frx":008F
      Style           =   2  'Dropdown List
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chkObj 
      BackColor       =   &H00E0E0E0&
      Caption         =   "可选条件"
      Height          =   255
      Index           =   0
      Left            =   2088
      TabIndex        =   3
      Top             =   3840
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.ComboBox cbxAge 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      ItemData        =   "frmQueryFilter.frx":0091
      Left            =   4365
      List            =   "frmQueryFilter.frx":00A1
      Style           =   2  'Dropdown List
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4260
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAge 
      Height          =   312
      Index           =   0
      Left            =   2088
      MaxLength       =   3
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4260
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ComboBox cbxDateUnit 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Index           =   0
      ItemData        =   "frmQueryFilter.frx":00BD
      Left            =   4125
      List            =   "frmQueryFilter.frx":00DC
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1020
      Visible         =   0   'False
      Width           =   1015
   End
   Begin MSScriptControlCtl.ScriptControl sctExecute 
      Left            =   900
      Top             =   3390
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin MSComCtl2.DTPicker dtpObj 
      Height          =   324
      Index           =   0
      Left            =   2088
      TabIndex        =   5
      Top             =   1020
      Visible         =   0   'False
      Width           =   2052
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   21430273
      CurrentDate     =   41297
   End
   Begin VB.Label labObj 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "标题占位:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   660
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label labError 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "没有需要录入的项目"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1470
      TabIndex        =   11
      Top             =   2505
      Visible         =   0   'False
      Width           =   3435
   End
   Begin VB.Label labMemo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   636
      Left            =   792
      TabIndex        =   10
      Top             =   108
      Width           =   4656
   End
   Begin VB.Image imgQuery 
      Height          =   720
      Left            =   36
      Picture         =   "frmQueryFilter.frx":0126
      Stretch         =   -1  'True
      Top             =   72
      Width           =   720
   End
   Begin VB.Shape shpBack 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00CEFFFA&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   5670
   End
End
Attribute VB_Name = "frmQueryFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Sql查询语句中的可选参数个数为"[@参数名,字段='value']"

Private mobjFilterValue As clsSqlFilterValue
Private mobjSchemeItem As TSchemeItem

Private mobjCmdQuery As Object       '创建录入界面时，保存的上一次创建的录入组件

Private mobjLastControl As Object       '创建录入界面时，保存的上一次创建的录入组件
Private maryInputTag() As TInputTag  '保存录入组件值改变后，需关联触发的控件

Private mdblFontSize As Double
Private mdblZoomRate As Double
Private mblnIsLoading As Boolean
Private mblnIsOK As Boolean
Private mblnIsMoreInput As Boolean  '是否有更多录入项
Private mblnIsResetDateItem As Boolean    '记录cbxrange下拉框是否重新设置选择

Private mlngFocusHwnd As Long
Private mblnIsEmbed As Boolean      '是否嵌入式

Private maryItemData(50, 100) As String '保存itemdata的数据

Private WithEvents mobjSqlParse As clsSqlParse
Attribute mobjSqlParse.VB_VarHelpID = -1

Property Get IsMoreInput() As Boolean
    IsMoreInput = mblnIsMoreInput
End Property

Property Get IsEmbed() As Boolean
    IsEmbed = mblnIsEmbed
End Property

Property Let IsEmbed(Value As Boolean)
    mblnIsEmbed = Value
End Property

Public Function ShowFilter(ByRef objSchemeItem As TSchemeItem, _
    ByVal dblFontSize As Double, owner As Object) As Boolean
'显示过滤窗口
    ShowFilter = False
    Set mobjCmdQuery = owner
    
    If objSchemeItem.FilterValues Is Nothing Then
        Set objSchemeItem.FilterValues = New clsSqlFilterValue
    End If
    
    Set mobjFilterValue = objSchemeItem.FilterValues
    
    mobjSchemeItem = objSchemeItem
    
    If mobjSchemeItem.SqlScheme Is Nothing Then
        MsgBox "查询方案 [" & objSchemeItem.BaseInfo.Name & "] 无效，不能获取方案的解析对象。", vbOKOnly, Me.Caption
        Exit Function
    End If
    
    Call ApplyOwnerFontSize(owner)
    
    If dblFontSize > 0 Then mdblFontSize = dblFontSize
    
    If mblnIsEmbed Then
        mblnIsOK = True
    Else
        mblnIsOK = False
        Me.Show 1, owner
    End If
    
    objSchemeItem = mobjSchemeItem
    
    ShowFilter = mblnIsOK
    
End Function

Public Sub UpdateInputData(ByVal strFilterName As String, strValue As Variant)
'配置界面录入数据
    Dim inputTag As TInputTag
    Dim objInputControl As Object
    Dim i As Long
    Dim j As Long
    Dim strParName As String
    
    For i = 1 To UBound(maryInputTag)
        inputTag = maryInputTag(i)
        
        strParName = IIf(inputTag.ControlType = ctChk, "@", IIf(inputTag.ControlType = ctQueryWay, "*", "")) & inputTag.ParName
        
        If strParName = strFilterName Then
            Set objInputControl = inputTag.InputControl
            
            Select Case inputTag.ControlType
                Case ctText '文本框
                    
                    If Trim(inputTag.ExtProperty) <> "" Then
                        If GetExtPropertyValue(inputTag.ExtProperty, EXT_LIKEWAY) <> "" Then
                            objInputControl.Text = Replace(strValue, "%", "")
                        End If
                    Else
                        objInputControl.Text = strValue
                    End If
                    
                Case ctDate, ctDateTime, ctTime, ctFastDate '日期框
                    objInputControl.Value = strValue
                    
                    If sdrRange.Tag <> "" Then
                        If inputTag.Index = sdrRange.Tag Or inputTag.Index = sdrRange.Tag + 1 Then
                            sdrRange.Value = dtpObj(sdrRange.Tag + 1).Value - dtpObj(sdrRange.Tag).Value
                        End If
                    End If
                Case ctAgeCbx   '年龄框
                    If strValue = "" Then Exit Sub
                    
                    Select Case cbxAge(inputTag.Index)
                        Case "S-岁"
                            strValue = CInt(Val(strValue) / 365) + IIf(Val(strValue) Mod 365 > 0, 1, 0)
                        Case "Y-月"
                            strValue = CInt(Val(strValue) / 30) + IIf(Val(strValue) Mod 30 > 0, 1, 0)
                        Case "Z-周"
                            strValue = CInt(Val(strValue) / 7) + IIf(Val(strValue) Mod 7 > 0, 1, 0)
                    End Select
                    
                    objInputControl.Text = strValue
                    
                Case ctCombobox, ctQueryWay '下拉框
                    '下拉框索引如果为0,表示没有进行选择
                    If maryItemData(inputTag.Index, 1) <> "" Then
                        For j = 0 To objInputControl.ListCount - 1
                            If maryItemData(inputTag.Index, j) = strValue Then
                                inputTag.InputControl.ListIndex = j
                                Exit Sub
                            End If
                        Next j
                    Else
                        objInputControl.Text = strValue
                    End If
                    
                Case ctList '列表框
                    If maryItemData(inputTag.Index, 0) <> "" Then
                        For j = 0 To objInputControl.ListCount - 1
                            If InStr("," & strValue & ",", "," & maryItemData(inputTag.Index, j) & ",") > 0 Then
                                objInputControl.Selected(j) = True
                            End If
                        Next j
                    Else
                        For j = 0 To objInputControl.ListCount - 1
                            If InStr("," & strValue & ",", "," & objInputControl.List(j) & ",") > 0 Then
                                objInputControl.Selected(j) = True
                            End If
                        Next j
                    End If
                    
                Case ctChk  '可选框
                    If CBool(strValue) <> False Then objInputControl.Value = 1
                    
                Case ctMutxCbx
                    If Trim(strValue) <> "" Then
                        cbxWhere.Text = inputTag.ParName
                        txtObj(cbxWhere.Tag).Text = strValue
                    End If
            End Select
            
            '需要触发change事件
            Call ControlChange(inputTag, IIf(Trim(strValue) = "", True, False))
                    
            Exit Sub
        End If
    Next i
End Sub

Private Sub ApplyOwnerFontSize(owner As Object)
On Error GoTo errHandle
    Dim dblSize As Double
    
    dblSize = owner.FontSize
    
    mdblFontSize = dblSize
    
Exit Sub
errHandle:
End Sub

Private Sub SetFontSize(ByVal lngFontSize As Double)
'设置字体大小
On Error Resume Next
    Dim i As Long
    Dim objControl As control
    Dim dblRate As Double
    Dim lngHeight As Long
    
    If lngFontSize <= 0 Then Exit Sub
    
    dblRate = lngFontSize / mdblFontSize
    For Each objControl In Me.Controls
        If objControl.Name <> "labMemo" Then
        
            lngHeight = objControl.Height
            
            objControl.Font.Size = lngFontSize
            
            objControl.Top = objControl.Top * dblRate
            objControl.Height = lngHeight * dblRate
            
            If objControl.Name <> "labObj" Then
                objControl.Left = objControl.Left * dblRate
            End If
            
            If objControl.Name = "chkObj" Then
                objControl.Width = objControl.Width * dblRate
            End If
            
            If objControl.Width + objControl.Left > Me.ScaleWidth Then
                objControl.Width = Me.ScaleWidth - objControl.Left - 75
            End If
            
        End If
    Next
    
    mdblFontSize = lngFontSize
    
Err.Clear
End Sub
Public Sub RefreshFontSize(ByVal dblFontSize As Double)
    If dblFontSize < 9 Or dblFontSize = mdblFontSize Then
        Exit Sub
    End If
    
    Call SetFontSize(dblFontSize)
    
    Call AutoHide
End Sub

Private Sub ConfigTitleDisplay()
    Me.Caption = mobjSchemeItem.SqlScheme.SchemeName
    labMemo.Caption = "说明:" & mobjSchemeItem.SqlScheme.Descript
End Sub

Private Function IsSql(ByVal strFrom As String) As Boolean
'是否sql语句
    Dim lngSelectIndex As Long
    Dim lngFromIndex As Long
    Dim strUCase As String
    
    IsSql = False
    strUCase = UCase(strFrom)
    
    lngSelectIndex = InStr(strUCase, "SELECT")
    lngFromIndex = InStr(strUCase, "FROM")
    
    If lngSelectIndex < 0 Or lngFromIndex < 0 Then Exit Function
    
    If lngFromIndex <= lngSelectIndex Then Exit Function
    
    IsSql = True
End Function


Private Sub ConfigSysDateInput(ByRef lngStartInputIndex As Long)
'配置系统时间录入
    Dim inputTag As TInputTag
    
    '开始日期条件处理
    inputTag.ParName = "系统.开始日期"
    inputTag.DisplayName = "[开始日期]"
    inputTag.DataFrom = ""
    inputTag.FromType = dbftText
    inputTag.ControlType = 3
    inputTag.Index = lngStartInputIndex
    inputTag.Default = mobjSchemeItem.Startdate
    
    ReDim inputTag.ParList(0)
    ReDim inputTag.ReleationInputIndex(0)
    
    Set inputTag.InputControl = CreateInputControl(inputTag.DisplayName, inputTag.ControlType, _
        lngStartInputIndex, Format(mobjSchemeItem.Startdate, "yyyy-mm-dd hh:mm:ss"))

    
    ReDim Preserve maryInputTag(lngStartInputIndex)
    maryInputTag(lngStartInputIndex) = inputTag
    
    lngStartInputIndex = lngStartInputIndex + 1
    
    '结束日期条件处理
    inputTag.ParName = "系统.结束日期"
    inputTag.DisplayName = "[结束日期]"
    inputTag.DataFrom = ""
    inputTag.FromType = dbftText
    inputTag.ControlType = 3
    inputTag.Index = lngStartInputIndex
    inputTag.Default = mobjSchemeItem.EndDate
    
    ReDim inputTag.ParList(0)
    ReDim inputTag.ReleationInputIndex(0)
    
    Set inputTag.InputControl = CreateInputControl(inputTag.DisplayName, inputTag.ControlType, _
        lngStartInputIndex, Format(mobjSchemeItem.EndDate, "yyyy-mm-dd hh:mm:ss"))

    
    ReDim Preserve maryInputTag(lngStartInputIndex)
    maryInputTag(lngStartInputIndex) = inputTag
    
    lngStartInputIndex = lngStartInputIndex + 1
    
    sdrRange.Top = inputTag.InputControl.Top + inputTag.InputControl.Height + 45
    sdrRange.Left = inputTag.InputControl.Left
    sdrRange.Width = inputTag.InputControl.Width
    sdrRange.Tag = 1
    
    cbxRange.Top = sdrRange.Top - 15
    cbxRange.Width = cbxRange.Width * mdblZoomRate
    cbxRange.Left = sdrRange.Left - cbxRange.Width - 120
    Call SetListIndex(cbxRange, 0)
    cbxRange.Tag = 1
    
    sdrRange.Value = CDate(Format(mobjSchemeItem.Startdate, "yyyy-MM-dd")) - CDate(Format(mobjSchemeItem.EndDate, "yyyy-MM-dd"))
    sdrRange.Visible = True
    cbxRange.Visible = True
    
    Set mobjLastControl = sdrRange
End Sub


Private Sub ConfigInputControl()
'配置界面录入
    Dim i As Long
    Dim objSqlScheme As clsSqlScheme
    Dim strParName As String
    Dim lngLastOrder As Long
    Dim inputTag As TInputTag
    Dim objSerachCfg As clsScSerachCfg
    Dim objSqlParse As clsSqlParse
    Dim lngInputIndex As Long
    
    lngInputIndex = 1
    
    ReDim maryInputTag(0)
    
    Set objSqlScheme = mobjSchemeItem.SqlScheme
    Set objSqlParse = New clsSqlParse
    
    '判断是否有系统的开始日期和结束日期条件......
    If InStr(objSqlScheme.Query, "[系统.开始日期]") > 0 _
        And InStr(objSqlScheme.Query, "[系统.结束日期]") > 0 Then
        
        Call ConfigSysDateInput(lngInputIndex)
    End If
    '
    
    For i = 1 To objSqlScheme.SerachCfgCount
        Set objSerachCfg = objSqlScheme.SerachCfg(i)
        
        inputTag.ParName = objSerachCfg.Name
        inputTag.ExtProperty = objSerachCfg.ExtProperty
        inputTag.DataFrom = Trim$(objSerachCfg.DataFrom)
        inputTag.FromType = dbftText
        inputTag.ControlType = objSerachCfg.ControlType
        inputTag.Index = lngInputIndex
        inputTag.Default = objSerachCfg.Default
        
        '判断数据来源类型
        If inputTag.DataFrom <> "" Then
            If IsSql(inputTag.DataFrom) Then
                inputTag.FromType = dbftSql
            End If
        End If
        
        ReDim inputTag.ParList(0)
        ReDim inputTag.ReleationInputIndex(0)
        
        If inputTag.FromType = 1 Then
            objSqlParse.init inputTag.DataFrom
            If objSqlParse.SqlStruct.ParCount > 0 Then
                CopyStrArray objSqlParse.SqlStruct.AllParameterAry, inputTag.ParList
            End If
        End If
        
        Set inputTag.InputControl = CreateInputControl(inputTag.ParName, inputTag.ControlType, lngInputIndex)
   
        ReDim Preserve maryInputTag(lngInputIndex)
        maryInputTag(lngInputIndex) = inputTag
        
        lngInputIndex = lngInputIndex + 1
    Next i
    
End Sub

Private Sub ControlChange(ByRef inputTag As TInputTag, Optional ByVal blnIsNull As Boolean = False)
'当前控件内容改变后，同步其他控件中数据来源以当前控件作为参数的数据
    Dim i As Long
    Dim j As Long
    Dim releationInputTag As TInputTag
    Dim lngBound As Long
    
    '如果tag为空，则计算该项目关联的录入配置
    If inputTag.Tag = "" Then
        ReDim inputTag.ReleationInputIndex(0)
        For i = inputTag.Index + 1 To UBound(maryInputTag)
            releationInputTag = maryInputTag(i)
            For j = 1 To UBound(releationInputTag.ParList)
                If releationInputTag.ParList(j) = "[" & inputTag.ParName & "]" Then
                
                    lngBound = UBound(inputTag.ReleationInputIndex) + 1
                    ReDim Preserve inputTag.ReleationInputIndex(lngBound)
                    
                    inputTag.ReleationInputIndex(lngBound) = i
                    Exit For
                End If
            Next j
        Next i
        
        inputTag.Tag = "1"
    End If
    
'    If blnIsNull Then
'        For i = 1 To UBound(inputTag.ReleationInputIndex)
'            Call ClearControlValue(maryInputTag(inputTag.ReleationInputIndex(i)).InputControl, maryInputTag(inputTag.ReleationInputIndex(i)).ControlType)
'        Next i
'    Else
        For i = 1 To UBound(inputTag.ReleationInputIndex)
            Call ConfigControlValue(maryInputTag(inputTag.ReleationInputIndex(i)), False)
        Next i
'    End If
    
End Sub


Private Function CreateInputControl(ByVal strName As String, ByVal lngInputType As Long, _
    ByVal lngOrder As Long, Optional ByVal strDefault As String = "") As Object
'创建录入组件
    Dim lngChkObjCount As Long
    Dim lngStartLeft As Long
    Dim blnIsOption As Boolean
    Dim lngStartTop As Long
'    Dim blnReplaceAsterisk As Boolean
    
    lngStartLeft = 1750 '1950
    lngStartLeft = lngStartLeft * mdblZoomRate
    
    lngStartTop = IIf(mblnIsEmbed, 120, 1080)
    
    blnIsOption = False
'    blnReplaceAsterisk = False
    
    Set CreateInputControl = Nothing
    
    Select Case lngInputType
        Case ctText
            '创建文本框组件
            Load txtObj(lngOrder)
            
            txtObj(lngOrder).Tag = strName
            
            txtObj(lngOrder).Left = lngStartLeft
            
            If mobjLastControl Is Nothing Then
                txtObj(lngOrder).Top = lngStartTop '315
            Else
                txtObj(lngOrder).Top = mobjLastControl.Top + mobjLastControl.Height + 120
            End If
            
            Set mobjLastControl = txtObj(lngOrder)
            
        Case ctDate, ctTime, ctDateTime, ctFastDate
            '创建日期框组件
            Load dtpObj(lngOrder)
                        
            dtpObj(lngOrder).Height = 288 * mdblZoomRate
            dtpObj(lngOrder).Format = dtpCustom
            dtpObj(lngOrder).CustomFormat = IIf(lngInputType = ctDate Or lngInputType = ctFastDate, "yyyy-MM-dd", IIf(lngInputType = ctTime, "HH:mm", "yyyy-MM-dd HH:mm"))
            
            dtpObj(lngOrder).UpDown = IIf(lngInputType = ctTime, True, False)
            
            dtpObj(lngOrder).Value = CurServerDate
            If strDefault <> "" Then dtpObj(lngOrder).Value = CDate(strDefault)

            dtpObj(lngOrder).Tag = strName
            
            dtpObj(lngOrder).Left = lngStartLeft
            
            
            If mobjLastControl Is Nothing Then
                dtpObj(lngOrder).Top = lngStartTop '315
            Else
                dtpObj(lngOrder).Top = mobjLastControl.Top + mobjLastControl.Height + 120
            End If
                        
            If lngInputType = ctFastDate Then
                '增加快选控件
                Load cbxDateUnit(lngOrder)
                
                cbxDateUnit(lngOrder).Tag = strName
                
                Call cbxDateUnit(lngOrder).AddItem("今天")
                Call cbxDateUnit(lngOrder).AddItem("前一天")
                Call cbxDateUnit(lngOrder).AddItem("前两天")
                Call cbxDateUnit(lngOrder).AddItem("前三天")
                Call cbxDateUnit(lngOrder).AddItem("前一周")
                Call cbxDateUnit(lngOrder).AddItem("前二周")
                Call cbxDateUnit(lngOrder).AddItem("前一月")
                Call cbxDateUnit(lngOrder).AddItem("前三月")
                Call cbxDateUnit(lngOrder).AddItem("前半年")
            
                cbxDateUnit(lngOrder).ListIndex = 0
                
                cbxDateUnit(lngOrder).Left = dtpObj(lngOrder).Left + (dtpObj(lngOrder).Width * mdblZoomRate) + 60
                cbxDateUnit(lngOrder).Width = cbxDateUnit(lngOrder).Width * mdblZoomRate
                cbxDateUnit(lngOrder).Top = dtpObj(lngOrder).Top
                
                cbxDateUnit(lngOrder).Visible = True
            Else
                dtpObj(lngOrder).Width = 3135
                
                If lngInputType = ctFastDate Then dtpObj(lngOrder).CheckBox = True
            End If
            
            Set mobjLastControl = dtpObj(lngOrder)
            
        Case ctCombobox, ctQueryWay
            '创建下拉框
            Load cbxObj(lngOrder)
            
            cbxObj(lngOrder).Tag = strName
            
            cbxObj(lngOrder).Left = lngStartLeft
            
            cbxObj(lngOrder).Text = ""
            
            If mobjLastControl Is Nothing Then
                cbxObj(lngOrder).Top = lngStartTop '315
            Else
                cbxObj(lngOrder).Top = mobjLastControl.Top + mobjLastControl.Height + 120
            End If
            
            If lngInputType = ctQueryWay Then
                 cbxObj(lngOrder).BackColor = &H8000000F
'                 blnReplaceAsterisk = True
            End If
            
            Set mobjLastControl = cbxObj(lngOrder)
        Case ctList
            '创建可多选的列表框
            Load lstObj(lngOrder)
            
            lstObj(lngOrder).Height = 1400 * mdblZoomRate
            
            lstObj(lngOrder).Tag = strName
            
            lstObj(lngOrder).Left = lngStartLeft
            
            If mobjLastControl Is Nothing Then
                lstObj(lngOrder).Top = lngStartTop '315
            Else
                lstObj(lngOrder).Top = mobjLastControl.Top + mobjLastControl.Height + 120
            End If
            
            Set mobjLastControl = lstObj(lngOrder)


        Case ctAgeCbx
            '创建年龄框组件
            Load txtAge(lngOrder)
            Load cbxAge(lngOrder)
            
            txtAge(lngOrder).Tag = strName
            cbxAge(lngOrder).Tag = strName
            
            txtAge(lngOrder).Left = lngStartLeft
            cbxAge(lngOrder).Left = lngStartLeft + (txtAge(lngOrder).Width * mdblZoomRate)
            cbxAge(lngOrder).Width = cbxAge(lngOrder).Width * mdblZoomRate
            
            If mobjLastControl Is Nothing Then
                txtAge(lngOrder).Top = lngStartTop '315
                cbxAge(lngOrder).Top = lngStartTop
            Else
                txtAge(lngOrder).Top = mobjLastControl.Top + mobjLastControl.Height + 120
                cbxAge(lngOrder).Top = txtAge(lngOrder).Top
            End If
            
            Call cbxAge(lngOrder).AddItem("S-岁")
            Call cbxAge(lngOrder).AddItem("Y-月")
            Call cbxAge(lngOrder).AddItem("Z-周")
            Call cbxAge(lngOrder).AddItem("T-天")
            
            cbxAge(lngOrder).ListIndex = 0
            cbxAge(lngOrder).Visible = True
            
            Set mobjLastControl = txtAge(lngOrder)
            
        Case ctMutxCbx  '互斥条件框
            If Trim(cbxWhere.Tag) = "" Then
                Load txtObj(lngOrder)
                
                txtObj(lngOrder).Width = txtObj(lngOrder).Width * mdblZoomRate
'                txtObj(lngOrder).Tag = strName
                txtObj(lngOrder).Left = lngStartLeft
                
                If mobjLastControl Is Nothing Then
                    txtObj(lngOrder).Top = lngStartTop '315
                Else
                    txtObj(lngOrder).Top = mobjLastControl.Top + mobjLastControl.Height + 120
                End If
                                
                
                Set mobjLastControl = txtObj(lngOrder)
            Else
                
            End If
            
        Case ctChk '可选条件
            Load chkObj(lngOrder)
            chkObj(lngOrder).Tag = strName
            chkObj(lngOrder).Caption = strName
            
            chkObj(lngOrder).Width = TextWidth(strName) * 1.2 + 252
            
            If Val(strDefault) <> 0 Then
                chkObj(lngOrder).Value = 1
            End If
            
            chkObj(lngOrder).Left = lngStartLeft
            
            If mobjLastControl Is Nothing Then
                chkObj(lngOrder).Top = lngStartTop '315
            Else
                chkObj(lngOrder).Top = mobjLastControl.Top + mobjLastControl.Height + 120
            End If
            
'            lngChkObjCount = chkObj.Count
'            If (lngChkObjCount Mod 2) = 0 Then
'                chkObj(lngOrder).Left = lngStartLeft
'
'                If mobjLastControl Is Nothing Then
'                    chkObj(lngOrder).Top = 1080 '315
'                Else
'                    chkObj(lngOrder).Top = mobjLastControl.Top + mobjLastControl.Height + 120
'                End If
'            Else
'                If chkObj(chkObj.UBound - 1).Width > 1485 Then
'                    chkObj(lngOrder).Left = lngStartLeft
'                    chkObj(lngOrder).Top = mobjLastControl.Top + mobjLastControl.Height + 120
'                Else
'                    chkObj(lngOrder).Left = 3600
'                    chkObj(lngOrder).Top = chkObj(chkObj.UBound - 1).Top
'                End If
'            End If
            
            
            Set mobjLastControl = chkObj(lngOrder)
'            mobjLastControl.Visible = True
            
            blnIsOption = True
'            Exit Function

            
    End Select
    
    mobjLastControl.Visible = True
    Set CreateInputControl = mobjLastControl
    
    If blnIsOption Then
        Exit Function
    End If
    
    If lngInputType = ctMutxCbx Then
        If Trim(cbxWhere.Tag) = "" Then
            cbxWhere.Visible = True
            cbxWhere.Width = cbxWhere.Width * mdblZoomRate
            cbxWhere.Tag = lngOrder
            
            
            cbxWhere.Left = mobjLastControl.Left - cbxWhere.Width - 120
            cbxWhere.Top = mobjLastControl.Top + 30
        End If
        
        cbxWhere.AddItem strName
        cbxWhere.ListIndex = 0
        
        Set CreateInputControl = cbxWhere 'txtObj(cbxWhere.Tag)
    Else
        mobjLastControl.Width = mobjLastControl.Width * mdblZoomRate
        
        '创建Label数据
        Load labObj(lngOrder)
        
'        If blnReplaceAsterisk Then
'            labObj(lngOrder).Caption = Mid(strName, 2, 100)
'        Else
            labObj(lngOrder).Caption = strName
'        End If
        
        labObj(lngOrder).Left = mobjLastControl.Left - labObj(lngOrder).Width - 120
        labObj(lngOrder).Top = mobjLastControl.Top + 60
        labObj(lngOrder).Visible = True
    End If
End Function

Private Sub cbxDateUnit_Change(Index As Integer)
On Error GoTo errHandle
    'dtpObj(Index).Value =
    Select Case cbxDateUnit(Index).ListIndex
        Case 0  '今天
            dtpObj(Index).Value = CurServerDate
        Case 1  '前一天
            dtpObj(Index).Value = CurServerDate - 1
        Case 2  '前两天
            dtpObj(Index).Value = CurServerDate - 2
        Case 3  '前三天
            dtpObj(Index).Value = CurServerDate - 3
        Case 4  '前一周
            dtpObj(Index).Value = CurServerDate - 7
        Case 5  '前二周
            dtpObj(Index).Value = CurServerDate - 14
        Case 6  '前一月
            dtpObj(Index).Value = CurServerDate - 30
        Case 7  '前三月
            dtpObj(Index).Value = CurServerDate - 90
        Case 8  '前半年
            dtpObj(Index).Value = CurServerDate - 180
    End Select
        
    Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub cbxObj_Change(Index As Integer)
'下拉框数据值被用户改变后，需要处理的数据加载
On Error GoTo errHandle
    If mblnIsLoading Then Exit Sub
    
    Call ControlChange(maryInputTag(Index), IIf(cbxObj(Index).Text = "", True, False))
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub cbxObjValidate(objCbo As ComboBox)
On Error GoTo errHandle
    Dim iCount As Integer
    Dim i As Integer

    iCount = Len(objCbo.Text)

    If iCount = 0 Then
        objCbo.ListIndex = 0
        Exit Sub
    End If

    For i = 0 To objCbo.ListCount
        If InStr(objCbo.List(i), objCbo.Text) > 0 Then
            objCbo.ListIndex = i
            Exit Sub
        End If
    Next
    objCbo.ListIndex = 0
    Exit Sub
errHandle:
End Sub

Private Sub cbxObj_Click(Index As Integer)
'下拉框数据值被用户改变后，需要处理的数据加载
On Error GoTo errHandle
    If mblnIsLoading Then Exit Sub
    
    Call ControlChange(maryInputTag(Index), IIf(cbxObj(Index).Text = "", True, False))
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub cbxObj_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo errH
    If KeyAscii = 13 Then Call cbxObjValidate(cbxObj(Index))
    Exit Sub
errH:
End Sub

Private Sub cbxRange_Click()
On Error GoTo errHandle
    Dim lngDays As Long
    
'        默认
'        今天
'        近两天
'        近三天
'        近一周
'        近两周
'        近半月
'        近一月
'        近二月
'        近三月

    If mblnIsResetDateItem = True Then Exit Sub
    
    Select Case cbxRange.Text
        Case "默认"
            lngDays = mobjSchemeItem.SqlScheme.DefaultQueryDays
        Case "今天"
            lngDays = 0
        Case "近一天"
            lngDays = 1
        Case "近二天"
            lngDays = 2
        Case "近三天"
            lngDays = 3
        Case "近一周"
            lngDays = 7
        Case "近二周"
            lngDays = 14
        Case "近半月"
            lngDays = 15
        Case "近一月"
            lngDays = 30
        Case "近二月"
            lngDays = 60
        Case "近三月"
            lngDays = 90
    End Select
    
    dtpObj(sdrRange.Tag + 1).Value = Format(Now, "yyyy-mm-dd 23:59:59")
    sdrRange.Value = lngDays
    
    Call sdrRange_Scroll
    gblnTimeChanged = True
    Exit Sub
errHandle:
End Sub


Private Sub cmdCancel_Click()
On Error GoTo errHandle
    
    Me.Hide
    Exit Sub
errHandle:
End Sub

Private Sub ClearInput()
    Dim objFree As Object
    
    '清除录入数据
    For Each objFree In txtObj
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then objFree.Text = ""
        End If
    Next
    
    For Each objFree In txtAge
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then objFree.Text = ""
        End If
    Next
    
    For Each objFree In cbxAge
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then objFree.ListIndex = 0
        End If
    Next
    
    For Each objFree In lstObj
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then Call objFree.Clear
        End If
    Next
    
    For Each objFree In cbxObj
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then objFree.Text = ""
        End If
    Next
    
    For Each objFree In dtpObj
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then objFree.Value = CurServerDate
        End If
    Next
    
    For Each objFree In cbxDateUnit
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then objFree.ListIndex = 0
        End If
    Next
    
    
    For Each objFree In chkObj
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then objFree.Value = 0
        End If
    Next
End Sub

Private Sub cmdClear_Click()
mblnIsLoading = True

On Error GoTo errHandle
    
    Call Restore(Nothing)
    
'    '清除录入数据
'    Call ClearInput
'
'    mblnIsLoading = False
    
    '载入配置的录入数据
'    Call LoadInputData
        
Exit Sub
errHandle:
    mblnIsLoading = False
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Public Function UpdateFindCondition(Optional blOnlyChangeDate As Boolean = False) As TSchemeItem
'配置查找环境条件
' blOnlyChangeDate 专用于日期值变化后更新日期参数
On Error GoTo errHandle
    Dim i As Long
    Dim inputTag As TInputTag
    Dim strExtValue As String
    Dim lwLikeWay As TLikeWay
    
    If ValidDateRange = False Then Exit Function
    
    For i = 1 To UBound(maryInputTag)
        inputTag = maryInputTag(i)
        
        If inputTag.ParName = "系统.开始日期" Then
            mobjSchemeItem.Startdate = Format(inputTag.InputControl.Value, "yyyy-mm-dd HH:mm")
            mobjFilterValue.ParData("系统.开始日期") = mobjSchemeItem.Startdate
        ElseIf inputTag.ParName = "系统.结束日期" Then
            mobjSchemeItem.EndDate = Format(inputTag.InputControl.Value, "yyyy-mm-dd HH:mm") '"yyyy-mm-dd 23:59:59"
            mobjFilterValue.ParData("系统.结束日期") = mobjSchemeItem.EndDate
        Else
            If Not blOnlyChangeDate Then
                If inputTag.ControlType = ctChk Then
                    Call mobjFilterValue.UpdateParValue("@" & inputTag.ParName, _
                                                GetControlValue(inputTag.InputControl, inputTag))
                                                
                ElseIf inputTag.ControlType = ctQueryWay Then
                    Call mobjFilterValue.UpdateParValue("*" & inputTag.ParName, _
                                                GetControlValue(inputTag.InputControl, inputTag))
                Else
                    '判断匹配方式
                    lwLikeWay = lwNormal
                    strExtValue = GetExtPropertyValue(inputTag.ExtProperty, EXT_LIKEWAY)
                    If strExtValue <> "" Then
                                    
                        If strExtValue = EXT_PRO_VALUE_LEFTWAY Then
                            lwLikeWay = lwLeft
                        ElseIf strExtValue = EXT_PRO_VALUE_RIGHTWAY Then
                            lwLikeWay = lwRight
                        ElseIf strExtValue = EXT_PRO_VALUE_FULLWAY Then
                            lwLikeWay = lwAll
                        End If
                    End If
                    
                    
                    Call mobjFilterValue.UpdateParValue(inputTag.ParName, _
                                            GetControlValue(inputTag.InputControl, inputTag, lwLikeWay))
                End If
            End If
        End If
    Next i
    
    UpdateFindCondition = mobjSchemeItem
    
Exit Function
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Function

Public Sub Restore(objFilterValue As clsSqlFilterValue, Optional ByVal isValidTime As Boolean = False, Optional ByVal dtstartTime As Date = Empty, Optional ByVal dtendTime As Date = Empty)
'恢复初始的查找条件，并重新检索数据
    Call ClearInput
    
    Call LoadBaseInputData
    
    '恢复默认日期范围设置
    If mobjSchemeItem.SqlScheme Is Nothing Then Exit Sub
    
    If Val(sdrRange.Tag) <> 0 And mobjSchemeItem.SqlScheme.DefaultQueryDays > 0 Then
        Call SetListIndex(cbxRange, 0)
        sdrRange.Value = 0
        
        If gblnTimeChanged And isValidTime And dtstartTime <> Empty And dtendTime <> Empty Then
            dtpObj(sdrRange.Tag + 1).Value = Format(dtendTime, "yyyy-mm-dd 23:59")
            
            dtpObj(sdrRange.Tag).Value = Format(dtstartTime, "yyyy-mm-dd 23:59")
        Else
            dtpObj(sdrRange.Tag + 1).Value = Format(Now, "yyyy-mm-dd 23:59")
            
            dtpObj(sdrRange.Tag).Value = Format(dtpObj(sdrRange.Tag + 1).Value - mobjSchemeItem.SqlScheme.DefaultQueryDays, "yyyy-mm-dd 00:00")
        End If
    End If
    
    If cbxWhere.Visible Then
        If cbxWhere.ListCount > 0 Then cbxWhere.ListIndex = 0
    End If
    
    'Call ReadUserInputConfig(objFilterValue)
End Sub

Private Function ValidDateRange() As Boolean
    Dim lngRange As Long
    Dim lngDays As Long
    Dim lngIndex As Long
    
    ValidDateRange = True
    
    If sdrRange.Visible And mobjSchemeItem.SqlScheme.dateRange > 0 Then
        lngDays = mobjSchemeItem.SqlScheme.dateRange * 366
        lngRange = dtpObj(sdrRange.Tag + 1).Value - dtpObj(sdrRange.Tag).Value
        
        If lngRange > lngDays Then
            MsgBox "查询范围不能超过系统设定的 [" & mobjSchemeItem.SqlScheme.dateRange & "] 年。", vbOKOnly, Me.Caption
            
            lngIndex = sdrRange.Tag
            dtpObj(lngIndex).Value = dtpObj(lngIndex).Value + (lngRange - lngDays)
         
            dtpObj(lngIndex).SetFocus
            
            ValidDateRange = False
            Exit Function
        End If
    End If
End Function

Private Sub cmdSure_Click()
    If ValidDateRange = False Then Exit Sub
            
    Call UpdateFindCondition
    
    Me.Hide
    
    mblnIsOK = True
End Sub

Private Sub dtpObj_Change(Index As Integer)
'日期框数据值被用户改变后，需要处理的数据加载
On Error GoTo errHandle
'    Dim lngDays As Double
'    Dim lngRange As Long

    
    If mblnIsLoading Then Exit Sub
    
    Call ControlChange(maryInputTag(Index))
    Call UpdateFindCondition(True)
    
    If sdrRange.Visible Then
        If Index = sdrRange.Tag Or Index = sdrRange.Tag + 1 Then
'            lngDays = mobjSchemeItem.SqlScheme.DateRange * 366
            
            sdrRange.Value = dtpObj(sdrRange.Tag + 1).Value - dtpObj(sdrRange.Tag).Value
            SetRange
'            lngRange = dtpObj(sdrRange.tag + 1).Value - dtpObj(sdrRange.tag).Value
'            If lngRange > lngDays Then
'                MsgBox "查询范围不能超过系统设定的 [" & mobjSchemeItem.SqlScheme.DateRange & "]年。", vbOKOnly
'                If Index = sdrRange.tag + 1 Then
'                    dtpObj(Index).Value = dtpObj(Index).Value - (lngRange - lngDays)
'                Else
'                    dtpObj(Index).Value = dtpObj(Index).Value + (lngRange - lngDays)
'                End If
'                dtpObj(Index).SetFocus
'            End If
        End If
    End If
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub



Public Sub LoadFace()
On Error GoTo errHandle
      
    mblnIsLoading = True
    
    Set mobjSqlParse = New clsSqlParse
    
    Call ConfigBaseWindow
    Call ConfigTitleDisplay
    Call ConfigInputControl
    
    Call UpdateWindowSize
    
    'Call AutoHide
    
    Call sctExecute.AddObject("Me", Me, True)
    
    Call LoadBaseInputData
    
    '根据filtervalue配置界面条件
    Call ReadUserInputConfig(mobjFilterValue)
    
    mblnIsLoading = False
Exit Sub
errHandle:
    mblnIsLoading = False
    MsgBox "查询过滤窗口加载失败:" & Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub dtpObj_DropDown(Index As Integer)
    gblnTimeChanged = True
End Sub

Private Sub dtpObj_KeyPress(Index As Integer, KeyAscii As Integer)
    gblnTimeChanged = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errHandle
    Dim lngHwnd As Long
    
    If KeyCode = vbKeyReturn Then
        If mblnIsEmbed And Me.ActiveControl.hwnd = mlngFocusHwnd Then
            Call mobjCmdQuery.SetFocus
        Else
            SendKeys ("{TAB}")
        End If
    End If
errHandle:
End Sub

Private Sub Form_Load()
    If mdblFontSize = 0 Then mdblFontSize = 9
    If mblnIsEmbed Then Exit Sub
    
    Call LoadFace
End Sub


Private Sub AutoHide()
'控件自动隐藏处理
On Error GoTo errHandle
    Dim obj As Object
    Dim lngMaxIndex As Long
    Dim lngHwnd As Long
    
    mblnIsMoreInput = False
    
    If mblnIsEmbed = False Then Exit Sub
    
    mlngFocusHwnd = 0
    lngHwnd = 0
    
    For Each obj In Me.Controls
        Call ControlVisible(obj)

        If UCase(TypeName(obj)) = "TEXTBOX" Or UCase(TypeName(obj)) = "DTPICKER" Or UCase(TypeName(obj)) = "COMBOBOX" Or UCase(TypeName(obj)) = "CHECKBOX" Then
            If UCase(TypeName(obj)) = "COMBOBOX" Then
                If obj.Style = 0 Then
                    obj.TabStop = True
                End If
            ElseIf UCase(TypeName(obj)) = "COMBOBOX" Then
                obj.TabStop = True
            End If
            
            If obj.Visible Then
                If lngMaxIndex < obj.TabIndex Then
                    lngMaxIndex = obj.TabIndex
                    lngHwnd = obj.hwnd
                End If
            End If
        End If

    Next
    mlngFocusHwnd = lngHwnd

    Exit Sub
errHandle:
    Err.Clear
End Sub

Private Sub ControlVisible(obj As Object)
On Error Resume Next
    Dim blnVisible As Boolean
    
    If obj.Name <> "cbxWhere" And obj.Name <> "cbxRange" Then
        If Val(obj.Index) <= 0 Then Exit Sub
    End If
    
    If InStr(obj.Name, "[系统.") > 0 Or (InStr(obj.Tag, "[系统.") > 0) Then Exit Sub
    
    blnVisible = IIf(obj.Top + obj.Height > Me.ScaleHeight, False, True)
    
    If obj.Name = "cbxWhere" Then
        If blnVisible Then
            If obj.ListCount <= 0 Then
                blnVisible = False
            End If
        End If
    ElseIf obj.Name = "sdrRange" Then
        If obj.Tag = "" Then
            blnVisible = False
        End If
        
        cbxRange.Visible = blnVisible
    ElseIf obj.Name = "cbxRange" Then
        Exit Sub
    End If
    
    obj.Visible = blnVisible
    
    '如果有控件被隐藏，则表示还有更多录入控件没有被显示出来
    If obj.Visible = False And obj.Name <> "cbxWhere" Then mblnIsMoreInput = True
    
Err.Clear
End Sub

Public Sub ReadUserInputConfig(objFilterValue As clsSqlFilterValue)
On Error GoTo errHandle
    Dim i As Long
    
    If Not objFilterValue Is Nothing Then
        For i = 1 To objFilterValue.Count
            Call UpdateInputData(objFilterValue.Item(i).Name, objFilterValue.Item(i).Value)
        Next i
    End If
    SetRange
Exit Sub
errHandle:
    Debug.Print "ReadUserConfig Err:" & Err.Description
End Sub

Public Function GetFromData(ByVal strSql As String) As ADODB.Recordset
'获取来源数据
On Error GoTo errHandle
    Dim strQuerySql As String
    
    Set GetFromData = Nothing
    
    Call mobjSqlParse.init(strSql)
    
    strQuerySql = mobjSqlParse.GetQuerySql(False)
    
    Set GetFromData = ExecuteCore(strQuerySql, "获取条件数据", mobjSqlParse.ParValues)
Exit Function
errHandle:
    Err.Raise -1, "frmQueryFilter.GetFromData", "[GetFromData]处理错误>>" & vbCrLf & "  查询语句为：" & strSql & vbCrLf & Err.Description
    Resume
End Function

Private Sub LoadBaseInputData()
'加载可选录入数据
'录入项的条件不允许超过20
    Dim i As Long
    
    Dim inputTag As TInputTag
    Dim inputLen As Long

    
    
    inputLen = UBound(maryInputTag)
    For i = 1 To inputLen
        inputTag = maryInputTag(i)
        
        Call ConfigControlValue(inputTag, True)
    Next i
End Sub

Private Sub ConfigControlValue(ByRef inputTag As TInputTag, ByVal blnIsSetDefault As Boolean)
'根据数据来源配置控件录入值
    Dim i As Long
    Dim lngInputType As Long
    Dim strTextDataSource() As String
    Dim rsSqlDataSource As ADODB.Recordset
    Dim strDefaultValue As String
    Dim objInputControl As Object
    Dim strDataItem As String
    
    lngInputType = inputTag.ControlType
    
    If inputTag.FromType = dbftText Then
        strTextDataSource = Split(inputTag.DataFrom, ";")
    Else
        Set rsSqlDataSource = GetFromData(inputTag.DataFrom)
    End If
    
    strDefaultValue = RunScripting(sctExecute, inputTag.Default)
    
    Set objInputControl = inputTag.InputControl
    
    Select Case lngInputType

        Case 0
            '读取文本框显示的数据
            If inputTag.FromType = dbftText Then
                Call SetControlValue(objInputControl, inputTag.ControlType, inputTag.DataFrom)
            Else
                Call SetControlValue(objInputControl, inputTag.ControlType, rsSqlDataSource(0).Value)
            End If

            If strDefaultValue <> "" Then
                Call SetControlValue(objInputControl, inputTag.ControlType, strDefaultValue)
            End If
        Case 1, 2, 3, 10
            '读取日期框显示的数据


            If strDefaultValue <> "" And strDefaultValue <> CDate(0) Then
                Call SetControlValue(objInputControl, inputTag.ControlType, strDefaultValue)
            Else
                If inputTag.ParName = "系统.开始日期" Then
                    Call SetControlValue(objInputControl, inputTag.ControlType, Format(Now, "yyyy-mm-dd 00:00"))
                    Exit Sub
                End If
                
                If inputTag.ParName = "系统.结束日期" Then
                    Call SetControlValue(objInputControl, inputTag.ControlType, Format(Now, "yyyy-mm-dd 23:59"))
                    Exit Sub
                End If
                
                If inputTag.FromType = dbftText Then
                    Call SetControlValue(objInputControl, inputTag.ControlType, Format(Now, "yyyy-mm-dd"))
                Else
                    Call SetControlValue(objInputControl, inputTag.ControlType, rsSqlDataSource(0).Value)
                End If
            End If
        Case 4, 9
            '读取下拉框显示的数据
            objInputControl.Clear
            
            If lngInputType <> ctQueryWay Then
                objInputControl.AddItem ""
            End If
            
            If inputTag.FromType = dbftText Then
                
                For i = 0 To UBound(strTextDataSource)
                    If i >= 100 Then Exit For
                    
                    strDataItem = strTextDataSource(i)
                    
                    If Trim$(strDataItem) <> "" Then
                        objInputControl.AddItem ParseInputValue(strDataItem, False)
'                        objInputControl.ItemData(objInputControl.ListCount - 1) = Val(ParseInputValue(strDataItem, True))
                        maryItemData(inputTag.Index, objInputControl.ListCount - 1) = ParseInputValue(strDataItem, True)
                    End If
                Next i
            Else
                
                i = 0
                Do While Not rsSqlDataSource.EOF
                    If i >= 100 Then Exit Do
                    i = i + 1
                    
                    strDataItem = rsSqlDataSource(0).Value
                    
                    If Trim$(strDataItem) <> "" Then
                        objInputControl.AddItem ParseInputValue(strDataItem, False)
'                        objInputControl.ItemData(objInputControl.ListCount - 1) = Val(ParseInputValue(strDataItem, True))
                        maryItemData(inputTag.Index, objInputControl.ListCount - 1) = ParseInputValue(strDataItem, True)
                    End If
                    
                    rsSqlDataSource.MoveNext
                Loop
            End If

            If strDefaultValue <> "" Then
                Call SetControlValue(objInputControl, inputTag.ControlType, strDefaultValue)
            Else
                If objInputControl.ListCount > 0 Then
                    objInputControl.ListIndex = 0
                Else
                    objInputControl.Text = ""
                End If
            End If
        Case 5
            '读取可多选列表框显示的数据
            objInputControl.Clear
            
            If inputTag.FromType = dbftText Then
                For i = 0 To UBound(strTextDataSource)
                    If i >= 100 Then Exit For
                    
                    strDataItem = strTextDataSource(i)
                    
                    If Trim$(strDataItem) <> "" Then
                        objInputControl.AddItem ParseInputValue(strDataItem, False)
'                        objInputControl.ItemData(objInputControl.ListCount - 1) = Val(ParseInputValue(strDataItem, True))
                        maryItemData(inputTag.Index, objInputControl.ListCount - 1) = ParseInputValue(strDataItem, True)
                    End If
                    
                    If InStr(strDefaultValue, strDataItem) > 0 Then
                        objInputControl.Selected(objInputControl.ListCount - 1) = True
                    End If
                Next i
            Else
                i = 0
                Do While Not rsSqlDataSource.EOF
                    If i >= 100 Then Exit Do
                    i = i + 1
                    
                    strDataItem = rsSqlDataSource(0).Value
                    
                    If Trim$(strDataItem) <> "" Then
                        objInputControl.AddItem ParseInputValue(strDataItem, False)
'                        objInputControl.ItemData(objInputControl.ListCount - 1) = Val(ParseInputValue(strDataItem, True))
                        maryItemData(inputTag.Index, objInputControl.ListCount - 1) = ParseInputValue(strDataItem, True)
                    End If

                    If InStr(strDefaultValue, rsSqlDataSource(0).Value) > 0 Then
                        objInputControl.Selected(objInputControl.ListCount - 1) = True
                    End If

                    rsSqlDataSource.MoveNext
                Loop
            End If
        Case 8
            
    End Select
End Sub

Private Function ParseInputValue(ByVal strSourceValue As String, ByVal blnIsItemData As Boolean) As String
On Error GoTo errHandle
    Dim lngSplitIndex As Long
    
    ParseInputValue = strSourceValue
    
    If InStr(Trim$(strSourceValue), "@") = 1 Then
        If blnIsItemData = False Then
            ParseInputValue = Mid(strSourceValue, 2, 255)
        Else
            ParseInputValue = ""
        End If
        Exit Function
    End If
    
    lngSplitIndex = InStr(strSourceValue, "-")
    
    If blnIsItemData Then
        If lngSplitIndex <= 0 Then
            ParseInputValue = ""
        Else
            ParseInputValue = Mid(strSourceValue, 1, lngSplitIndex - 1)
        End If
    Else
        If lngSplitIndex <= 0 Then
            ParseInputValue = strSourceValue
        Else
            ParseInputValue = Mid(strSourceValue, lngSplitIndex + 1, 255)
        End If
    End If
Exit Function
errHandle:
    ParseInputValue = ""
End Function

Private Sub SetControlValue(objInputControl As Object, ByVal lngInputType As Long, ByVal strValue As Variant)
'对控件的文本或者value属性赋值
On Error Resume Next
    Dim i As Long
    
    Select Case lngInputType
        Case ctText '文本框
            objInputControl.Text = strValue
        Case ctDate, ctDateTime, ctTime, ctFastDate '日期框
            objInputControl.Value = strValue
        Case ctCombobox, ctQueryWay '下拉框
            objInputControl.Text = strValue
        Case ctList '列表框
            For i = 0 To objInputControl.ListCount - 1
                If objInputControl.List(i) = strValue Then
                    objInputControl.Selected(i) = True
                End If
            Next i
        Case ctChk  '可选框
            If CBool(strValue) <> False Then objInputControl.Value = 1
    End Select
End Sub

Private Sub ClearControlValue(objInputControl As Object, ByVal lngInputType As Long)
    Select Case lngInputType
        Case ctText '文本框
            objInputControl.Text = ""
        Case ctCombobox '下拉框
            Call objInputControl.Clear
        Case ctList '列表框
            Call objInputControl.Clear
    End Select
End Sub

Private Sub UpdateWindowSize()
    If Not mobjLastControl Is Nothing Then
        framButton.Top = mobjLastControl.Top + mobjLastControl.Height + 120 + 15
        Me.Height = framButton.Top + framButton.Height + 400 - 15
        
        cmdClear.Top = framButton.Top + 240 * mdblZoomRate
        cmdCancel.Top = framButton.Top + 240 * mdblZoomRate
        cmdSure.Top = framButton.Top + 240 * mdblZoomRate
        
        labError.Visible = False
    Else
        labError.Visible = True
    End If
End Sub

Private Sub ConfigBaseWindow()
    mdblZoomRate = 1
    If mdblFontSize > 9 Then
        Call SetFontSize(mdblFontSize)
        Me.FontSize = mdblFontSize
        
        mdblZoomRate = 1 + (mdblFontSize / 2 - 5) / 10
    End If
    
    Me.Width = 5724 * mdblZoomRate
    
    If mblnIsEmbed Then
        shpBack.Visible = False
        labMemo.Visible = False
        imgQuery.Visible = False
        
        framButton.Visible = False
        
        cmdClear.Visible = False
        cmdCancel.Visible = False
        cmdSure.Visible = False
    Else
        shpBack.Width = 5675 * mdblZoomRate
        labMemo.Width = 4656 * mdblZoomRate
            
        framButton.Left = -45
        framButton.Width = Me.ScaleWidth + 90
        framButton.Height = 795 * mdblZoomRate
        
        
        cmdClear.Width = 1300 * mdblZoomRate
        cmdClear.Height = 375 * mdblZoomRate
        
        
        cmdCancel.Width = 1300 * mdblZoomRate
        cmdCancel.Height = 375 * mdblZoomRate
        cmdCancel.Left = Me.Width - cmdCancel.Width - 240
        
        cmdSure.Width = 1300 * mdblZoomRate
        cmdSure.Height = 375 * mdblZoomRate
        cmdSure.Left = cmdCancel.Left - cmdSure.Width - 240
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 5 Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub Form_Resize()
On Error Resume Next

    Call AutoAdjustWidth
    Call AutoHide
    
End Sub


Private Sub AutoAdjustWidth()
On Error GoTo errHandle
    Dim obj As Object
    
    mblnIsMoreInput = False
    
    If mblnIsEmbed = False Then Exit Sub
    
    For Each obj In Me.Controls
        Call AdjustControlWidth(obj)
    Next
    
        
Exit Sub
errHandle:
    Err.Clear
End Sub

Private Sub AdjustControlWidth(objControl As Object)
On Error Resume Next
    If objControl.Visible Then
        If InStr("txtObj,cbxObj,lstObj,sdrRange,dtpObj,cbxDateUnit,txtAge,cbxAge,chkObj", objControl.Name) > 0 Then
            If (objControl.Left + objControl.Width + 75) > Me.Width Then
                objControl.Width = Me.Width - objControl.Left - 75
            End If
            
            If objControl.Width < 3135 And (objControl.Left + objControl.Width + 75) < Me.Width Then
                objControl.Width = Me.Width - objControl.Left - 75
            End If
        End If
    End If
Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set mobjCmdQuery = Nothing
    Set mobjSqlParse = Nothing
    Set mobjFilterValue = Nothing
    Set mobjLastControl = Nothing
End Sub

Private Sub lstObj_Click(Index As Integer)
'多选框数据值被用户改变后，需要处理的数据加载
On Error GoTo errHandle
    If mblnIsLoading Then Exit Sub
    
    Call ControlChange(maryInputTag(Index))
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub mobjSqlParse_OnGetParameterValue(ByVal strParName As String, Value As Variant)
    '读取参数
    Dim i As Long
    Dim inputTag As TInputTag
    
    For i = 1 To UBound(maryInputTag)
        inputTag = maryInputTag(i)
        If inputTag.ParName = strParName Then
            Value = GetControlValue(inputTag.InputControl, inputTag)
            Exit Sub
        End If
    Next i
    
    Call GetSysPar(strParName, Value)
End Sub


Private Sub GetSysPar(ByVal strParName As String, ByRef Value As Variant)
On Error GoTo errH
''[系统.系统号],[系统.模块号],[系统.科室ID],[系统.用户ID],[系统.用户账号]
'[系统.服务器日期],[系统.服务器时间],[系统.本地日期],[系统.本地时间]
'[系统.开始日期],[系统.结束日期]"
'
'[系统.病人ID],[系统.医嘱ID]
    Select Case strParName
        Case "系统.系统号"
            Value = glngSysNo
            
        Case "系统.模块号"
            Value = glngModuleNo
            
        Case "系统.科室ID"
            Value = gstrDeptId
        
        Case "系统.用户ID"
            Value = glngUserId
            
        Case "系统.用户账号"
            Value = gstrUserAccount
            
        Case "系统.用户名称"
            Value = gstrUserName
            
        Case "系统.服务器日期"
            Value = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
            
        Case "系统.服务器时间"
            Value = Format(zlDatabase.Currentdate, "HH:mm:ss")
            
        Case "系统.本地日期"
            Value = Date
            
        Case "系统.本地时间"
            Value = Time
            
        Case Else
    End Select
errH:
End Sub


Private Function GetControlValue(objInputControl As Object, ByRef inputTag As TInputTag, Optional ByVal lngLikeWay As TLikeWay = lwNormal) As Variant
    Dim i As Long
    Dim blnIsUpper As Boolean
    Dim blnIsNumber As Boolean
    
    If inputTag.ControlType = ctMutxCbx Then
        blnIsUpper = IIf(mobjSchemeItem.SqlScheme.GetSerachExtValue(inputTag.ParName, EXT_UPPERCONVERT) = "1", True, False)
        blnIsNumber = IIf(mobjSchemeItem.SqlScheme.GetSerachExtValue(inputTag.ParName, EXT_NUMBERCONVERT) = "1", True, False)
    Else
        blnIsUpper = IIf(GetExtPropertyValue(inputTag.ExtProperty, EXT_UPPERCONVERT) = "1", True, False)
        blnIsNumber = IIf(GetExtPropertyValue(inputTag.ExtProperty, EXT_NUMBERCONVERT) = "1", True, False)
    End If
    
    Select Case inputTag.ControlType
        Case ctText  'textbox文本
            If Len(objInputControl.Text) > 0 Then
                If blnIsUpper Then objInputControl.Text = UCase(objInputControl.Text)
                If blnIsNumber Then objInputControl.Text = Val(objInputControl.Text)
            End If
            
            GetControlValue = objInputControl.Text
        Case ctDate  'dtpicker日期
            GetControlValue = CDate(Format(objInputControl.Value, "yyyy-MM-dd"))
        Case ctTime  'dtpicker时间
            GetControlValue = CDate(Format(objInputControl.Value, "HH:mm"))
        Case ctDateTime  'dtpicker日期时间
            GetControlValue = CDate(Format(objInputControl.Value, "yyyy-MM-dd HH:mm"))
        Case ctCombobox, ctQueryWay  'combobox下拉
            If Len(objInputControl.Text) > 0 Then
                If blnIsUpper Then objInputControl.Text = UCase(objInputControl.Text)
                If blnIsNumber Then objInputControl.Text = Val(objInputControl.Text)
            End If
            
            GetControlValue = objInputControl.Text
            
            If Trim(objInputControl.Text) = "" Then Exit Function
            
            If objInputControl.ListIndex >= 0 Then
'                If objInputControl.ItemData(objInputControl.ListIndex) <> 0 Then
'                    GetControlValue = objInputControl.ItemData(objInputControl.ListIndex)
'                End If
                If maryItemData(inputTag.Index, objInputControl.ListIndex) <> "" Then
                    GetControlValue = maryItemData(inputTag.Index, objInputControl.ListIndex)
                End If
            Else
                For i = 0 To objInputControl.ListCount - 1
                    If objInputControl.List(i) = objInputControl.Text Then
                        If maryItemData(inputTag.Index, i) <> "" Then
                            GetControlValue = maryItemData(inputTag.Index, i)
                        End If
                        
                        Exit Function
                    End If
                Next i
            End If
        Case ctList  'list列表
            For i = 0 To objInputControl.ListCount - 1
                If objInputControl.Selected(i) Then
                    If GetControlValue <> "" Then GetControlValue = GetControlValue & ","
'                    If objInputControl.ItemData(i) <> 0 Then
'                        GetControlValue = GetControlValue & objInputControl.ItemData(i)
'                    Else
'                        GetControlValue = GetControlValue & objInputControl.List(i)
'                    End If
                    If maryItemData(inputTag.Index, i) <> "" Then
                        GetControlValue = GetControlValue & maryItemData(inputTag.Index, i)
                    Else
                        GetControlValue = GetControlValue & objInputControl.List(i)
                    End If
                End If
            Next i
        Case ctChk  'checkbox可选框
            GetControlValue = IIf(objInputControl.Value <> 0, True, False)
            
        Case ctAgeCbx  '年龄组合框
            If Trim(objInputControl.Text) = "" Then Exit Function
            
            If Len(objInputControl.Text) > 0 Then
                If blnIsUpper Then objInputControl.Text = UCase(objInputControl.Text)
                If blnIsNumber Then objInputControl.Text = Val(objInputControl.Text)
            End If
            
            GetControlValue = GetAgeDays(objInputControl.Text, cbxAge(objInputControl.Index).Text)
        Case ctMutxCbx  '互n斥条件框组合
            If objInputControl.Text = inputTag.ParName Then
            
                If Len(txtObj(objInputControl.Tag).Text) > 0 Then
                    If blnIsUpper Then txtObj(objInputControl.Tag).Text = UCase(txtObj(objInputControl.Tag).Text)
                    If blnIsNumber Then txtObj(objInputControl.Tag).Text = Val(txtObj(objInputControl.Tag).Text)
                End If
                
                GetControlValue = txtObj(objInputControl.Tag).Text
            End If
        Case ctFastDate  '日期快选组合
            GetControlValue = CDate(Format(objInputControl.Value, "yyyy-MM-dd"))
        
    End Select
    
    If lngLikeWay <> lwNormal Then
        If IsEmpty(GetControlValue) Or IsNull(GetControlValue) Or GetControlValue = "" Then Exit Function
        
        If InStr(GetControlValue, "%") <> 1 _
            And InStr(GetControlValue, "%") <> Len(GetControlValue) Then
            Select Case lngLikeWay
                Case lwLeft
                    GetControlValue = GetControlValue & "%"
                Case lwRight
                    GetControlValue = "%" & GetControlValue
                Case lwAll
                    GetControlValue = "%" & GetControlValue & "%"
            End Select
        End If
    End If
    
    
End Function

Private Function GetAgeDays(ByVal strAge As String, ByVal strUnit As String) As Long
'转换为年龄天数
    Select Case strUnit
        Case "S-岁"
            GetAgeDays = Val(strAge) * 365
        Case "Y-月"
            GetAgeDays = Val(strAge) * 30
        Case "Z-周"
            GetAgeDays = Val(strAge) * 7
        Case "T-天"
            GetAgeDays = Val(strAge) * 1
    End Select
End Function

Private Function SetRange()
'计算开始日期和结束日期的间隔天数，根据间隔天数选择对应下拉框选项
On Error GoTo errHandle
    Dim intDays, intDefultDays As Integer
    
    mblnIsResetDateItem = True
    
    intDays = DateDiff("d", dtpObj(Val(sdrRange.Tag)).Value, dtpObj(Val(sdrRange.Tag) + 1).Value)
    intDefultDays = mobjSchemeItem.SqlScheme.DefaultQueryDays
        
     Select Case intDays
            Case intDefultDays
            cbxRange.Text = cbxRange.List(0)  '"默认"
            Case 0
            cbxRange.Text = cbxRange.List(1)  '"今天"
            Case 2
            cbxRange.Text = cbxRange.List(2)  '"近二天"
            Case 3
            cbxRange.Text = cbxRange.List(3)  '"近三天"
            Case 7
            cbxRange.Text = cbxRange.List(4)  '"近一周"
            Case 14
            cbxRange.Text = cbxRange.List(5)  '"近二周"
            Case 15
            cbxRange.Text = cbxRange.List(6)  '"近半月"
            Case 30
            cbxRange.Text = cbxRange.List(7)  '"近一月"
            Case 60
            cbxRange.Text = cbxRange.List(8)  '"近二月"
            Case 90
            cbxRange.Text = cbxRange.List(9)  '"近三月"
            Case Else
            cbxRange.Text = cbxRange.List(10)  '"自定义"
        End Select
    mblnIsResetDateItem = False
    Exit Function
errHandle:
    Debug.Print "ERR>>SetRange:" & Err.Description
End Function

Private Sub sdrRange_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    gblnTimeChanged = True
End Sub

Private Sub sdrRange_Scroll()
On Error GoTo errHandle
    dtpObj(Val(sdrRange.Tag)).Value = Format(dtpObj(sdrRange.Tag + 1).Value - sdrRange.Value, "yyyy-MM-dd 00:00")
    dtpObj(Val(sdrRange.Tag) + 1).Value = Format(dtpObj(sdrRange.Tag + 1).Value, "yyyy-MM-dd 23:59")
    SetRange
Exit Sub
errHandle:
    Debug.Print "ERR>>sdrRange_Scroll:" & Err.Description
End Sub

Private Sub txtAge_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtObj_Change(Index As Integer)
'文本框数据值被用户改变后，需要处理的数据加载
On Error GoTo errHandle
    If mblnIsLoading Then Exit Sub
    
    Call ControlChange(maryInputTag(Index))
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

