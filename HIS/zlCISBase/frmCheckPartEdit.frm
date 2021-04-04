VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCheckPartEdit 
   BorderStyle     =   0  'None
   Caption         =   "检查部位编辑"
   ClientHeight    =   8325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6195
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox cbo适用性别 
      Height          =   300
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1920
      Width           =   2160
   End
   Begin VB.ComboBox cbo方法 
      Height          =   300
      Left            =   3855
      TabIndex        =   17
      Text            =   "Combo1"
      Top             =   2565
      Width           =   2160
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "添加为附加可选方法(&A)"
      Enabled         =   0   'False
      Height          =   350
      Index           =   2
      Left            =   3855
      Picture         =   "frmCheckPartEdit.frx":0000
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3630
      Width           =   2160
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "添加为共用可选方法(&P)"
      Enabled         =   0   'False
      Height          =   350
      Index           =   1
      Left            =   3855
      Picture         =   "frmCheckPartEdit.frx":014A
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3300
      Width           =   2160
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "添加为互斥基础方法(&B)"
      Enabled         =   0   'False
      Height          =   350
      Index           =   0
      Left            =   3855
      Picture         =   "frmCheckPartEdit.frx":0294
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2970
      Width           =   2160
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "删除当前方法(&D)      "
      Enabled         =   0   'False
      Height          =   350
      Index           =   3
      Left            =   3855
      Picture         =   "frmCheckPartEdit.frx":03DE
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4155
      Width           =   2160
   End
   Begin VB.Frame fraLine 
      Height          =   15
      Left            =   -390
      TabIndex        =   15
      Top             =   1020
      Width           =   7320
   End
   Begin VB.TextBox txt名称 
      Height          =   300
      Left            =   1755
      MaxLength       =   60
      TabIndex        =   3
      Top             =   135
      Width           =   1920
   End
   Begin VB.TextBox txt编码 
      Height          =   300
      Left            =   555
      MaxLength       =   4
      TabIndex        =   1
      Top             =   135
      Width           =   525
   End
   Begin VB.ComboBox cbo分组 
      Height          =   300
      Left            =   4410
      Sorted          =   -1  'True
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   135
      Width           =   1620
   End
   Begin VB.TextBox txt备注 
      Height          =   300
      Left            =   555
      MaxLength       =   60
      TabIndex        =   7
      Top             =   555
      Width           =   5460
   End
   Begin VSFlex8Ctl.VSFlexGrid vfg方法 
      Height          =   3090
      Left            =   135
      TabIndex        =   9
      Top             =   1410
      Width           =   3555
      _cx             =   6271
      _cy             =   5450
      Appearance      =   2
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
      BackColorSel    =   16635590
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483643
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
      GridLines       =   0
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   7
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
      Begin MSComctlLib.ImageList imgList 
         Left            =   2985
         Top             =   1860
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
               Picture         =   "frmCheckPartEdit.frx":0528
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCheckPartEdit.frx":0AC2
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label lbl适用性别 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "适用性别"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3840
      TabIndex        =   19
      Top             =   1680
      Width           =   720
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmCheckPartEdit.frx":105C
      ForeColor       =   &H00008000&
      Height          =   1980
      Left            =   345
      TabIndex        =   16
      Top             =   4740
      Width           =   5580
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgNote 
      Height          =   240
      Left            =   15
      Picture         =   "frmCheckPartEdit.frx":1202
      Top             =   4710
      Width           =   240
   End
   Begin VB.Label lbl命名 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "方法命名"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3855
      TabIndex        =   10
      Top             =   2325
      Width           =   720
   End
   Begin VB.Label lbl组织 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "检查方法及方法组织:"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   135
      TabIndex        =   8
      Top             =   1170
      Width           =   1710
   End
   Begin VB.Label lbl名称 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "名称"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1335
      TabIndex        =   2
      Top             =   195
      Width           =   360
   End
   Begin VB.Label lbl编码 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "编码"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   135
      TabIndex        =   0
      Top             =   195
      Width           =   360
   End
   Begin VB.Label lbl分组 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "分组"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3975
      TabIndex        =   4
      Top             =   195
      Width           =   360
   End
   Begin VB.Label lbl备注 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "备注"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   135
      TabIndex        =   6
      Top             =   615
      Width           =   360
   End
End
Attribute VB_Name = "frmCheckPartEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    方法 = 0
    造影 = 2
End Enum

Private mstrKind As String          '当前类型
Private mstrPart As String          '当前部位
Private mblnPACSInterface As Boolean        '启用影像信息系统接口
Dim lngCount As Long

'--------------------------------------------
'以下为窗体公共方法
'--------------------------------------------
Private Sub FormatList(Optional strMode As String)
    '功能：初始化设置参考值列表
    '参数：strMode-方法串
    Dim aryItem() As String, strItems As String, strTemp As String
    Dim aryChild() As String, lngChild As Long
    With Me.vfg方法
        .Redraw = flexRDNone
        .Clear
        .Rows = 1: .FixedRows = 1: .Cols = 3: .FixedCols = 0
        .TextMatrix(0, mCol.方法) = "检查方法": .ColWidth(mCol.方法) = 280: .FixedAlignment(mCol.方法) = flexAlignCenterCenter
        .TextMatrix(0, mCol.方法 + 1) = "检查方法": .ColWidth(mCol.方法 + 1) = 2500
        .TextMatrix(0, mCol.造影) = "造影"
        .MergeCells = flexMergeFree: .MergeRow(0) = True
        If strMode = "" Then .Redraw = flexRDDirect: Exit Sub
        
        strItems = ""
        strTemp = ""
        If InStr(1, strMode, vbTab) > 0 Then strMode = Mid(strMode, 1, InStr(1, strMode, vbTab) - 1) & ";" & Mid(strMode, InStr(1, strMode, vbTab))
        For lngCount = 1 To Len(strMode)
            If Mid(strMode, lngCount, 1) = vbTab And lngCount <> 2 Then
                 If Mid(strTemp, Len(strTemp), 1) <> ";" Then strTemp = strTemp & ";"
            End If
            strTemp = strTemp & Mid(strMode, lngCount, 1)
        Next
        strMode = strTemp
        
        aryItem() = IIf(Mid(strMode, 1, 1) = ";", Split(Mid(strMode, 2), ";"), Split(strMode, ";"))
        For lngCount = 0 To UBound(aryItem)
            strTemp = aryItem(lngCount)
            If InStr(1, aryItem(lngCount), ",") > 0 Then strTemp = Left(aryItem(lngCount), InStr(1, aryItem(lngCount), ",") - 1)
            .Rows = .Rows + 1: .MergeRow(.Rows - 1) = True
            If InStr(1, strTemp, vbTab) = 0 Then
                .RowData(.Rows - 1) = 1
            Else
                .RowData(.Rows - 1) = 2
                strTemp = Mid(strTemp, 2)
            End If
            Set .Cell(flexcpPicture, .Rows - 1, mCol.方法) = Me.imgList.ListImages(.RowData(.Rows - 1)).Picture
            .TextMatrix(.Rows - 1, mCol.方法) = Mid(strTemp, 2)
            .TextMatrix(.Rows - 1, mCol.方法 + 1) = .TextMatrix(.Rows - 1, mCol.方法)
            If Val(Left(strTemp, 1)) = 1 Then
                .Cell(flexcpChecked, .Rows - 1, mCol.造影) = flexChecked
            Else
                .Cell(flexcpChecked, .Rows - 1, mCol.造影) = flexUnchecked
            End If
            If InStr(1, aryItem(lngCount), ",") > 0 Then
                strTemp = Mid(aryItem(lngCount), InStr(1, aryItem(lngCount), ",") + 1)
                aryChild = Split(strTemp, ",")
                For lngChild = 0 To UBound(aryChild)
                    strTemp = aryChild(lngChild)
                    .Rows = .Rows + 1: .MergeRow(.Rows - 1) = True
                    .RowData(.Rows - 1) = 2
                    Set .Cell(flexcpPicture, .Rows - 1, mCol.方法 + 1) = Me.imgList.ListImages(.RowData(.Rows - 1)).Picture
                    .TextMatrix(.Rows - 1, mCol.方法 + 1) = Mid(strTemp, 2)
                    If Val(Left(strTemp, 1)) = 1 Then
                        .Cell(flexcpChecked, .Rows - 1, mCol.造影) = flexChecked
                    Else
                        .Cell(flexcpChecked, .Rows - 1, mCol.造影) = flexUnchecked
                    End If
                Next
            End If
        Next

        If .Rows > .FixedRows Then .Row = .FixedRows
        .Redraw = flexRDDirect
    End With
End Sub

Public Function zlRefresh(strKind As String, strPart As String) As Boolean
    '功能：根据项目id刷新当前显示内容
    Dim rsTemp As New ADODB.Recordset
    mstrKind = strKind: mstrPart = strPart
    
    '清除此前项目的显示
    Me.txt编码.Text = "": Me.txt名称.Text = "": Me.cbo分组.Text = "": Me.txt备注.Text = ""
    If mstrPart = "" Then FormatList: zlRefresh = True: Exit Function
    
    '获取指定项目的信息
    err = 0: On Error GoTo ErrHand
    gstrSql = "Select 编码, 名称, 分组, 备注, 方法,适用性别 From 诊疗检查部位 Where 类型 = [1] And 编码 = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrKind, mstrPart)
    With rsTemp
        Me.txt编码.MaxLength = .Fields("编码").DefinedSize
        Me.txt名称.MaxLength = .Fields("名称").DefinedSize
        Me.cbo分组.Tag = .Fields("分组").DefinedSize
        Me.cbo方法.Tag = .Fields("名称").DefinedSize
        Me.txt备注.MaxLength = .Fields("备注").DefinedSize
        If .RecordCount > 0 Then
            Me.txt编码.Text = "" & !编码
            Me.txt名称.Text = "" & !名称
            Me.cbo分组.Text = "" & !分组
            Me.txt备注.Text = "" & !备注
            Me.cbo适用性别.ListIndex = NVL(!适用性别, 0)
            Call FormatList("" & !方法)
        End If
    End With
    
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart(blnAdd As Boolean, strPart As String) As Boolean
    '功能：开始项目编辑
    '参数： blnAdd-是否增加，否则为修改
    '       strPart-增加的参照项目，或者指定编辑的项目
    Dim rsTemp As New ADODB.Recordset
    Dim strValue As String

    err = 0: On Error GoTo ErrHand
    gstrSql = "Select Distinct 分组 From 诊疗检查部位 Where 类型 = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrKind)
    With rsTemp
        strValue = Me.cbo分组.Text
        Me.cbo分组.Clear
        Do While Not .EOF
            Me.cbo分组.AddItem "" & !分组
            .MoveNext
        Loop
        Me.cbo分组.Text = strValue
    End With
    
    gstrSql = "Select 名称 From Table(Cast(f_Check_Motheds([1]) As " & gstrDBOwner & ".t_Dic_Rowset)) Order By 名称"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrKind)
    With rsTemp
        strValue = Me.cbo方法.Text
        Me.cbo方法.Clear
        Do While Not .EOF
            Me.cbo方法.AddItem "" & !名称
            .MoveNext
        Loop
        Me.cbo方法.Text = strValue
    End With
    
   
    
    If blnAdd Then
        gstrSql = "Select Nvl(Max(编码), 0) As 编码, Nvl(Max(Length(编码)), 0) As 长度 From 诊疗检查部位 Where 类型 = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrKind)
        With rsTemp
            If !长度 <> 0 And !长度 <= Me.txt编码.MaxLength Then
                Me.txt编码.Text = Format(Val(!编码) + 1, String(!长度, "0"))
            Else
                Me.txt编码.Text = Format(Val(!编码) + 1, String(Me.txt编码.MaxLength, "0"))
            End If
        End With
        
        '清除并设置备注值
        Me.txt名称.Text = "": Me.txt备注.Text = ""
        'Call FormatList
    End If
    
    mstrPart = strPart
    Me.Enabled = True: Me.Tag = IIf(blnAdd, "增加", "修改"): Call Form_Resize
    Me.txt编码.SetFocus
    zlEditStart = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditStart = False: Exit Function
End Function

Public Sub zlEditCancel()
    '功能：放弃正在进行的编辑
    Me.Enabled = False: Me.Tag = "": Call Form_Resize
    Call zlRefresh(mstrKind, mstrPart)
End Sub

Public Function zlEditSave() As String
    '功能：保存正在进行的编辑,并返回正在编辑项目编码,保存失败返回""
    Dim strOption As String, strCheck As String, lngCheck As Long
    Dim blnTrans As Boolean, blnRisTrans As Boolean
    Dim strUpOption As String, strUpCheck As String
        
    With Me.vfg方法
        strOption = ""
        strUpOption = ""
        For lngCount = .FixedRows To .Rows - 1
            If .RowData(lngCount) = 1 Then
                strOption = strOption & ";" & IIf(.Cell(flexcpChecked, lngCount, mCol.造影) = flexChecked, 1, 0)
                strOption = strOption & .TextMatrix(lngCount, mCol.方法)
                strUpOption = strUpOption & ";|"
                strUpOption = strUpOption & .TextMatrix(lngCount, mCol.方法)
                strCheck = ""
                strUpCheck = ""
                For lngCheck = lngCount + 1 To .Rows - 1
                    If .TextMatrix(lngCheck, mCol.方法) <> "" Then Exit For
                    strCheck = strCheck & "," & IIf(.Cell(flexcpChecked, lngCheck, mCol.造影) = flexChecked, 1, 0)
                    strCheck = strCheck & .TextMatrix(lngCheck, mCol.方法 + 1)
                    strUpCheck = strUpCheck & "," & .TextMatrix(lngCount, mCol.方法) & "|" & .TextMatrix(lngCheck, mCol.方法 + 1)
                Next
                strOption = strOption & strCheck
                strUpOption = strUpOption & strUpCheck
            End If
            
            If .RowData(lngCount) = 2 And .TextMatrix(lngCount, mCol.方法) <> "" Then
                strUpCheck = ""
                strCheck = ""
                strCheck = strCheck & vbTab & IIf(.Cell(flexcpChecked, lngCount, mCol.造影) = flexChecked, 1, 0)
                strCheck = strCheck & .TextMatrix(lngCount, mCol.方法)
                strUpCheck = strUpCheck & ";|" & .TextMatrix(lngCount, mCol.方法)
                
                If strCheck <> "" Then strOption = strOption & vbTab & Mid(strCheck, 2)
                If strUpCheck <> "" Then strUpOption = strUpOption & ";" & Mid(strUpCheck, 2)
                strCheck = ""
                strUpCheck = ""
                For lngCheck = lngCount + 1 To .Rows - 1
                    If .TextMatrix(lngCheck, mCol.方法) <> "" Then Exit For
                    strCheck = strCheck & "," & IIf(.Cell(flexcpChecked, lngCheck, mCol.造影) = flexChecked, 1, 0)
                    strCheck = strCheck & .TextMatrix(lngCheck, mCol.方法 + 1)
                    strUpCheck = strUpCheck & "," & .TextMatrix(lngCount, mCol.方法) & "|" & .TextMatrix(lngCheck, mCol.方法 + 1)
                Next
                strOption = strOption & strCheck
                strUpOption = strUpOption & strUpCheck
            End If
        Next
'        If strOption <> "" Then strOption = Mid(strOption, 2)
        
        If strOption = "" Then
            MsgBox "请至少设置一种检查方法！", vbInformation, gstrSysName
            .SetFocus: zlEditSave = "": Exit Function
        End If
    End With
    
    
    '一般特性检查
    If Trim(Me.txt编码.Text) = "" Then
        MsgBox "请输入编码！", vbInformation, gstrSysName
        Me.txt编码.SetFocus: zlEditSave = "": Exit Function
    End If
    If Val(Me.txt编码.Text) > Val(String(Me.txt编码.MaxLength, "9")) Then
        MsgBox "编码太大！", vbInformation, gstrSysName
        Me.txt编码.SetFocus: zlEditSave = "": Exit Function
    End If
    If Trim(Me.txt名称.Text) = "" Then
        MsgBox "请输入名称！", vbInformation, gstrSysName
        Me.txt名称.SetFocus: zlEditSave = "": Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt名称.Text), vbFromUnicode)) > Me.txt名称.MaxLength Then
        MsgBox "名称超长（最多" & Me.txt名称.MaxLength & "个字符或等长汉字）！", vbInformation, gstrSysName
        Me.txt名称.SetFocus: zlEditSave = "": Exit Function
    End If
    If Trim(Me.cbo分组.Text) = "" Then
        MsgBox "请输入分组！", vbInformation, gstrSysName
        Me.cbo分组.SetFocus: zlEditSave = "": Exit Function
    End If
    If LenB(StrConv(Trim(Me.cbo分组.Text), vbFromUnicode)) > Val(Me.cbo分组.Tag) Then
        MsgBox "分组超长（最多" & Val(Me.cbo分组.Tag) & "个字符）！", vbInformation, gstrSysName
        Me.cbo分组.SetFocus: zlEditSave = "": Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt备注.Text), vbFromUnicode)) > Me.txt备注.MaxLength Then
        MsgBox "备注超长（最多" & Me.txt备注.MaxLength & "个字符）！", vbInformation, gstrSysName
        Me.txt备注.SetFocus: zlEditSave = "": Exit Function
    End If
    
    '数据保存语句组织
    gstrSql = "'" & mstrKind & "','" & mstrPart & "','" & Trim(Me.txt编码.Text) & "','" & Trim(Me.txt名称.Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.cbo分组.Text) & "','" & Trim(Me.txt备注.Text) & "','" & strOption & "'"
    gstrSql = gstrSql & "," & Me.cbo适用性别.ListIndex & ",'" & strUpOption & "'"
    
    If Me.Tag = "增加" Then
        gstrSql = "Zl_诊疗检查部位_Edit(1," & gstrSql & ")"
    Else
        gstrSql = "Zl_诊疗检查部位_Edit(2," & gstrSql & ")"
    End If
    
    err = 0: On Error GoTo ErrHand
    
    If Me.Tag <> "增加" Then
        '新网RIS接口，检查部位修改时，先删除原部位对应的诊疗项目部位；启用参数，接口部件有效的前提下
        '放到HIS执行过程之前
        If mblnPACSInterface = True Then
            If Not gobjRIS Is Nothing Then
                Dim strSql As String
                Dim rsData As ADODB.Recordset
            
                strSql = "Select 名称 From 诊疗检查部位 Where 编码 = [2] And 类型 = [1]"
                Set rsData = zlDatabase.OpenSQLRecord(strSql, "取原部位名称", mstrKind, mstrPart)
                If rsData.RecordCount > 0 Then
                    '传入部位类型和原部位名称
                    If gobjRIS.HISBasicDictTable(RISBaseItemType.ClinicItemPart, RISBaseItemOper.Delete, mstrKind & "|" & rsData!名称) <> 1 Then
                        '出错时提示接口错误信息
                        If gobjRIS.LastErrorInfo <> "" Then
                            MsgBox gobjRIS.LastErrorInfo, vbInformation, gstrSysName
                        Else
                            MsgBox "调用RIS接口错误，不能继续当前操作！请与系统管理员联系", vbInformation, gstrSysName
                        End If
                        
                        Exit Function
                    End If
                        
                    blnRisTrans = True
                End If
            Else
                '接口部件无效时禁止并提示
                 MsgBox "RIS接口创建失败，不能继续当前操作！可能是接口文件安装或注册不正常，请与系统管理员联系。", vbInformation, gstrSysName
                 
                 Exit Function
            End If
        End If
    End If
    
    gcnOracle.BeginTrans
    blnTrans = True
    
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    
    If Me.Tag <> "增加" Then
        '新网RIS接口，检查部位修改时，在删除了原部位对应的诊疗项目部位的前提下再增加新部位对应的方法；启用参数，接口部件有效的前提下
        '放到HIS执行过程之后
        If mblnPACSInterface = True Then
            If Not gobjRIS Is Nothing Then
                '传入部位类型和新部位名称
                If gobjRIS.HISBasicDictTable(RISBaseItemType.ClinicItemPart, RISBaseItemOper.AddNew, mstrKind & "|" & Trim(Me.txt名称.Text)) <> 1 Then
                    gcnOracle.RollbackTrans
                    
                    '出错时提示接口错误信息
                    If gobjRIS.LastErrorInfo <> "" Then
                        MsgBox gobjRIS.LastErrorInfo, vbInformation, gstrSysName
                    Else
                        MsgBox "调用RIS接口错误，不能继续当前操作！请与系统管理员联系", vbInformation, gstrSysName
                    End If
                    
                    Exit Function
                End If
                    
                blnRisTrans = True
            Else
                gcnOracle.RollbackTrans
                
                '接口部件无效时禁止并提示
                 MsgBox "RIS接口创建失败，不能继续当前操作！可能是接口文件安装或注册不正常，请与系统管理员联系。", vbInformation, gstrSysName
                 
                 Exit Function
            End If
        End If
    End If
    
    gcnOracle.CommitTrans
    blnTrans = False
    blnRisTrans = False
    
    mstrPart = Trim(Me.txt编码.Text)
    Me.Enabled = False: Me.Tag = "": Call Form_Resize
    zlEditSave = mstrPart: Exit Function
    
ErrHand:
    If blnTrans = True Then gcnOracle.RollbackTrans
    
    'Ris接口和HIS不同步时，写错误日志
    If blnRisTrans = True And Not gobjRIS Is Nothing Then
        MsgBox "HIS新增或删除检查部位错误，RIS接口和HIS数据不同步，请与系统管理员联系。", vbInformation, gstrSysName
        
        On Error Resume Next
        Call gobjRIS.WriteCommLog("frmCheckPartList：cbsThis_Execute", "HIS新增或删除检查部位错误，RIS接口和HIS数据不同步", "类型=" & mstrKind & " " & "部位名称=" & Trim(Me.txt名称.Text), 0)
    End If
    
    
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = "": Exit Function
End Function

'--------------------------------------------
'以下为窗体控件响应事件
'--------------------------------------------
Private Sub cbo方法_GotFocus()
    Me.cbo方法.SelStart = 0: Me.cbo方法.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub cbo方法_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR & ",;", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub cbo分组_GotFocus()
    Me.cbo分组.SelStart = 0: Me.cbo分组.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub cbo分组_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdEdit_Click(Index As Integer)
    Dim lngRowAdd As Long, lngRowParent As Long
    Dim i As Long
    
    With Me.vfg方法
        lngRowAdd = 0
        If Index = 3 Then
            '删除处理
            If .TextMatrix(.Row, mCol.方法) = "" Then
                .RemoveItem .Row
            Else
                .RemoveItem .Row
                Do While .TextMatrix(.Row, mCol.方法) = "" And .Row <= .Rows - 1
                    .RemoveItem .Row
                Loop
            End If
            Me.cbo方法.SetFocus
            Exit Sub
        End If
        
        '添加处理
        Me.cbo方法.Text = Replace(Me.cbo方法.Text, ",", "")
        Me.cbo方法.Text = Replace(Me.cbo方法.Text, ";", "")
        If Trim(Me.cbo方法.Text) = "" Then MsgBox "请指定方法名称！", vbInformation, gstrSysName: Me.cbo方法.SetFocus: Exit Sub
        If LenB(StrConv(Trim(Me.cbo方法.Text), vbFromUnicode)) > Val(Me.cbo方法.Tag) Then
            MsgBox "检查方法超长（最多" & Val(Me.cbo方法.Tag) & "个字符）！", vbInformation, gstrSysName
            Me.cbo方法.SetFocus: Exit Sub
        End If
        Select Case Index
        Case 0
            For lngCount = .FixedRows To .Rows - 1
                If .TextMatrix(lngCount, mCol.方法 + 1) = Trim(Me.cbo方法.Text) Then
                    MsgBox "已经设置了该方法，不能重复！", vbInformation, gstrSysName: Me.cbo方法.SetFocus: Exit Sub
                End If
'                If .TextMatrix(lngCount, mCol.方法) <> "" And .RowData(lngCount) = 2 And lngRowAdd = 0 Then lngRowAdd = lngCount
            Next
            If lngRowAdd = 0 Then lngRowAdd = .Rows
            .AddItem Trim(Me.cbo方法.Text) & vbTab & Trim(Me.cbo方法.Text), lngRowAdd
            .Row = lngRowAdd: .RowData(.Row) = 1: .MergeRow(.Row) = True
            Set .Cell(flexcpPicture, .Row, mCol.方法) = Me.imgList.ListImages(.RowData(.Row)).Picture
        Case 1
            For lngCount = .FixedRows To .Rows - 1
                If .TextMatrix(lngCount, mCol.方法 + 1) = Trim(Me.cbo方法.Text) Then
                    MsgBox "已经设置了该方法，不能重复！", vbInformation, gstrSysName: Me.cbo方法.SetFocus: Exit Sub
                End If
            Next
            lngRowAdd = .Rows
            .AddItem Trim(Me.cbo方法.Text) & vbTab & Trim(Me.cbo方法.Text), .Rows
            .Row = .Rows - 1: .RowData(.Row) = 2: .MergeRow(.Row) = True
            Set .Cell(flexcpPicture, .Row, mCol.方法) = Me.imgList.ListImages(.RowData(.Row)).Picture
        Case 2
            For lngCount = .FixedRows To .Rows - 1
                If .TextMatrix(lngCount, mCol.方法) = Trim(Me.cbo方法.Text) Then
                    MsgBox "已经设置了该方法，不能重复！", vbInformation, gstrSysName: Me.cbo方法.SetFocus: Exit Sub
                End If
            Next
            lngRowParent = .Row
            If .TextMatrix(.Row, mCol.方法) = "" Then
                For lngCount = .Row - 1 To .FixedRows Step -1
                    If .TextMatrix(lngCount, mCol.方法) <> "" Then lngRowParent = lngCount: Exit For
                Next
            Else
                lngRowParent = .Row
            End If
            For lngCount = lngRowParent + 1 To .Rows - 1
                If .TextMatrix(lngCount, mCol.方法) <> "" Then lngRowAdd = lngCount: Exit For
                If .TextMatrix(lngCount, mCol.方法 + 1) = Trim(Me.cbo方法.Text) Then
                    MsgBox "已经设置了该方法，不能重复！", vbInformation, gstrSysName: Me.cbo方法.SetFocus: Exit Sub
                End If
            Next
            '共用方法下不允许有相同的附加方法
            For lngCount = .FixedRows To .Rows - 1
                If .TextMatrix(lngCount, mCol.方法) <> "" And .RowData(lngCount) <> .RowData(.Row) Then
                    For i = lngCount + 1 To .Rows - 1
                        If .TextMatrix(i, mCol.方法) <> "" Then Exit For
                        If .TextMatrix(i, 1) = Me.cbo方法.Text Then
                            MsgBox "共用可选方法下不允许添加相同的附加方法！", vbInformation, gstrSysName: Me.cbo方法.SetFocus: Exit Sub
                        End If
                    Next
                End If
            Next
            If lngRowAdd = 0 Then lngRowAdd = .Rows
            .AddItem "" & vbTab & Trim(Me.cbo方法.Text), lngRowAdd
            .Row = lngRowAdd: .RowData(.Row) = 2: .MergeRow(.Row) = True
            Set .Cell(flexcpPicture, .Row, mCol.方法 + 1) = Me.imgList.ListImages(.RowData(.Row)).Picture
        End Select
        If InStr(1, Trim(Me.cbo方法.Text), "增强") > 0 Or InStr(1, Trim(Me.cbo方法.Text), "造影") > 0 Then
            .Cell(flexcpChecked, .Row, mCol.造影) = flexChecked
        Else
            .Cell(flexcpChecked, .Row, mCol.造影) = flexUnchecked
        End If
        If .RowIsVisible(.Row) = False Then .TopRow = .Row
        Me.cbo方法.SetFocus
    End With
    Call vfg方法_RowColChange
End Sub

Private Sub Form_Load()
    
    mblnPACSInterface = (Val(zlDatabase.GetPara(255, glngSys, , "0")) = 1)
    
    Me.cbo适用性别.Clear
    With Me.cbo适用性别
        .AddItem "0-无性别区分"
        .AddItem "1-男性"
        .AddItem "1-女性"
    End With
    
    Me.cbo适用性别.ListIndex = 0
    
    mstrPart = ""
    Call FormatList
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    err = 0: On Error Resume Next
    If Me.Tag <> "" Then
        Me.BackColor = RGB(230, 230, 230)
        Me.vfg方法.FocusRect = flexFocusHeavy
        Me.cmdEdit(0).Enabled = True
        Me.cmdEdit(1).Enabled = True
        Me.cmdEdit(2).Enabled = True
        Me.cmdEdit(3).Enabled = True
    Else
        Me.BackColor = &H8000000F
        Me.vfg方法.FocusRect = flexFocusNone
        Me.cmdEdit(0).Enabled = False
        Me.cmdEdit(1).Enabled = False
        Me.cmdEdit(2).Enabled = False
        Me.cmdEdit(3).Enabled = False
    End If
End Sub

Private Sub txt备注_GotFocus()
    Me.txt备注.SelStart = 0: Me.txt备注.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt备注_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt编码_GotFocus()
    Me.txt编码.SelStart = 0: Me.txt编码.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt编码_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt名称_GotFocus()
    Me.txt名称.SelStart = 0: Me.txt名称.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt名称_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vfg方法_DblClick()
    If Me.vfg方法.MouseRow < Me.vfg方法.FixedRows Then Exit Sub
    If Me.Tag = "" Then Exit Sub
    With Me.vfg方法
        If .Row < .FixedRows And .Row > .Rows - 1 Then Exit Sub
        If .Cell(flexcpChecked, .Row, mCol.造影) = flexChecked Then
            .Cell(flexcpChecked, .Row, mCol.造影) = flexUnchecked
        Else
            .Cell(flexcpChecked, .Row, mCol.造影) = flexChecked
        End If
        For lngCount = .FixedRows To .Rows - 1
            If .TextMatrix(lngCount, mCol.方法 + 1) = .TextMatrix(.Row, mCol.方法 + 1) Then
                .Cell(flexcpChecked, lngCount, mCol.造影) = .Cell(flexcpChecked, .Row, mCol.造影)
            End If
        Next
    End With
End Sub

Private Sub vfg方法_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If KeyAscii <> vbKeySpace Then Exit Sub
    Call vfg方法_DblClick
End Sub

Private Sub vfg方法_RowColChange()
    With Me.vfg方法
        If .Row < .FixedRows Then
            Me.cbo方法.Text = ""
            Me.cmdEdit(0).Enabled = Me.Enabled
            Me.cmdEdit(1).Enabled = Me.Enabled
            Me.cmdEdit(2).Enabled = False
            Me.cmdEdit(3).Enabled = False
        Else
            Me.cbo方法.Text = .TextMatrix(.Row, mCol.方法 + 1)
            Me.cmdEdit(0).Enabled = Me.Enabled
            Me.cmdEdit(1).Enabled = Me.Enabled
            Me.cmdEdit(2).Enabled = Me.Enabled
            Me.cmdEdit(3).Enabled = Me.Enabled
        End If
    End With
End Sub

