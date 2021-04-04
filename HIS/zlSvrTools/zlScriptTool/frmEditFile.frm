VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmEditFile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "文件选择"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6900
   Icon            =   "frmEditFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox chk部件 
      Caption         =   "公共部件"
      Height          =   240
      Index           =   0
      Left            =   105
      TabIndex        =   7
      Top             =   5580
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CheckBox chk部件 
      Caption         =   "应用部件"
      Height          =   240
      Index           =   1
      Left            =   1515
      TabIndex        =   6
      Top             =   5580
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CheckBox chk部件 
      Caption         =   "帮助文件"
      Height          =   240
      Index           =   2
      Left            =   2925
      TabIndex        =   5
      Top             =   5580
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CheckBox chk部件 
      Caption         =   "其他文件"
      Height          =   240
      Index           =   3
      Left            =   4320
      TabIndex        =   4
      Top             =   5580
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CheckBox chk部件 
      Caption         =   "三方部件"
      Height          =   240
      Index           =   4
      Left            =   5715
      TabIndex        =   3
      Top             =   5580
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4410
      TabIndex        =   2
      Top             =   5895
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5625
      TabIndex        =   1
      Top             =   5895
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfFiles 
      Height          =   5370
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   6660
      _cx             =   11747
      _cy             =   9472
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
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
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmEditFile.frx":6852
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
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   1
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
   Begin VB.Line Line6 
      X1              =   -45
      X2              =   10950
      Y1              =   5670
      Y2              =   5670
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000005&
      X1              =   -30
      X2              =   10950
      Y1              =   5970
      Y2              =   5985
   End
End
Attribute VB_Name = "frmEditFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Const conField = "Select case when b.COLUMN_VALUE is null then '0' else '1' end 选择 ,a.序号 as 文件ID,a.文件类型 as 文件类型,a.文件名 as 文件名称,a.文件说明 as 文件说明" & vbNewLine & _
'                         "from zlfilesupgrade A, Table (Cast(f_Str2List([1])  As zlTools.t_StrList)) B" & vbNewLine & _
'                         "where upper(a.文件名) = upper(b.COLUMN_VALUE(+)) And a.文件类型 in ([2]) order by a.文件名"

                         
Private Const conField = "select 选择,文件ID,文件类型,文件名称,文件说明 from" & vbNewLine & _
                         "(" & vbNewLine & _
                         "Select case when b.COLUMN_VALUE is null then '0' else '1' end 选择 ,a.主版本 as 文件ID,a.次版本 as 文件类型,a.部件 as 文件名称,a.名称 as 文件说明" & vbNewLine & _
                         "from zlcomponent A, Table (Cast(f_Str2List([1])  As zlTools.t_StrList)) B" & vbNewLine & _
                         "where upper(a.部件) = upper(b.COLUMN_VALUE(+)) and upper(a.部件) <> upper([2])" & vbNewLine & _
                         "Union" & vbNewLine & _
                         "Select case when b.COLUMN_VALUE is null then '0' else '1' end 选择 ,10 as 文件ID,27 as 文件类型,'zlSvrStudio' as 文件名称,'服务器管理工具' as 文件说明" & vbNewLine & _
                          "from dual A left join Table (Cast(f_Str2List([1])  As zlTools.t_StrList)) B on upper('zlSvrStudio')= upper(b.COLUMN_VALUE(+))" & vbNewLine & _
                        " ) order by 文件名称"

                         
                         
Private mintItemFile          As String
Private mintStrFile             As String '源文件名
Private mstrType                As String
Private rsM                     As ADODB.Recordset

Public Property Get intItemFile() As String
    intItemFile = mintItemFile
End Property

Public Property Let intItemFile(ByVal vNewValue As String)
    mintItemFile = vNewValue
End Property

Public Property Get intStrFile() As String
    intStrFile = mintStrFile
End Property

Public Property Let intStrFile(ByVal vNewValue As String)
    mintStrFile = vNewValue
End Property

Public Property Get strType() As String
    strType = mstrType
End Property

Public Property Let strType(ByVal vNewValue As String)
    mstrType = vNewValue
End Property

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngLoop         As Long
    Dim strFiles        As String
    
    With vsfFiles
        If .Rows <= 1 Then
            MsgBox "没有可供选择的文件！", vbExclamation, "提示"
            Exit Sub
        End If
'        strFiles = ""
'        For lngLoop = 1 To .Rows - 1
'            If Abs(.TextMatrix(lngLoop, .ColIndex("选择"))) = "1" Then
'                strFiles = strFiles & IIf(Len(.Cell(flexcpText, lngLoop, 4)) = 0, 0, .Cell(flexcpText, lngLoop, 4)) & ","
'            End If
'        Next
        If Len(strFiles) <> 0 Then
            strFiles = Left(strFiles, Len(strFiles) - 1)
            If LenB(strFiles) > 2000 Then
                MsgBox "选择文件过多，请重新选择", vbCritical, "提示"
                Exit Sub
            End If
        End If
    End With
'    mintItemFile = strFiles
    Unload Me
End Sub

Private Sub Form_Load()
    gstrSql = conField
    
'    gstrSql = Replace(gstrSql, "[2]", mstrType)
    If mintStrFile = "" Then
        mintStrFile = "1"
    Else
        mintStrFile = GetFileName(mintStrFile)
    End If
    Set rsM = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mintItemFile, mintStrFile)
    Set vsfFiles.DataSource = rsM
End Sub

Private Sub vsfFiles_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strFilename As String
    If vsfFiles.ColKey(Col) <> "选择" Then Exit Sub
    strFilename = vsfFiles.Cell(flexcpText, Row, 4)
    If strFilename = "" Then Exit Sub
    If vsfFiles.Cell(flexcpText, Row, 1) = "-1" Then
        '选择
        If mintItemFile = "" Then
            mintItemFile = strFilename
        Else
            mintItemFile = mintItemFile & "," & strFilename
        End If
    Else
        '未选择
        If mintItemFile <> "" Then
            If Right(mintItemFile, 1) <> "," Then
                mintItemFile = mintItemFile & ","
            End If
            
            mintItemFile = Replace(mintItemFile, strFilename & ",", "")
            
            If Right(mintItemFile, 1) = "," Then
                mintItemFile = Left(mintItemFile, Len(mintItemFile) - 1)
            End If
        End If
    End If
End Sub

Private Sub vsfFiles_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsfFiles.ColKey(Col) <> "选择" Then
        Cancel = True
    End If
End Sub


Private Sub chk部件_Click(Index As Integer)
    Dim strTemp As String
    On Error GoTo errH
    If chk部件(0).Value Then
        strTemp = "0,"
    End If
    
    If chk部件(1).Value Then
        If Len(strTemp) = 0 Then
            strTemp = "1,"
        Else
            strTemp = strTemp & "1,"
        End If
    End If
    
    If chk部件(2).Value Then
        If Len(strTemp) = 0 Then
            strTemp = "2,"
        Else
            strTemp = strTemp & "2,"
        End If
    End If
    
    If chk部件(3).Value Then
        If Len(strTemp) = 0 Then
            strTemp = "3,"
        Else
            strTemp = strTemp & "3,"
        End If
    End If
    
    If chk部件(4).Value Then
        If Len(strTemp) = 0 Then
            strTemp = "4"
        Else
            strTemp = strTemp & "4"
        End If
    End If
    
    If Len(strTemp) > 0 Then
        Call ShowData(strTemp)
    Else
        Call ShowData("Clear")
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

''在数据库中过滤
Private Sub ShowData(ByVal strOption As String)
    On Error GoTo errH
    gstrSql = conField
    
    If Len(strOption) > 0 Then
        If strOption = "Clear" Then
            gstrSql = Replace(gstrSql, "[2]", "-1")
            Set rsM = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mintItemFile)
            Set vsfFiles.DataSource = rsM
        Else
            If Right(strOption, 1) = "," Then
                strOption = Left(strOption, Len(strOption) - 1)
            End If
            gstrSql = Replace(gstrSql, "[2]", strOption)
            Set rsM = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mintItemFile)
            Set vsfFiles.DataSource = rsM
        End If
    Else
        '搜索使用
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

''在控件中过滤
''Private Sub ShowData(ByVal strOption As String)
''    On Error GoTo errH
''    Dim strTemp() As String
''    Dim i As Integer
''    Dim strFilter As String
''    Dim strFilters As String
''    strFilters = ""
''    strFilter = "文件类型="
''    gstrSql = conField
''    If Len(strOption) > 0 Then
''        If strOption = "Clear" Then
''            rsM.Filter = "文件类型=-1"
''            Set vsfFiles.DataSource = rsM
''        Else
''            If Right(strOption, 1) = "," Then
''                strOption = Left(strOption, Len(strOption) - 1)
''            End If
''
''            strTemp = Split(strOption, ",")
''            For i = 0 To UBound(strTemp)
''                If i = UBound(strTemp) Then
''                    strFilters = strFilters & strFilter & strTemp(i)
''                Else
''                    strFilters = strFilters & strFilter & strTemp(i) & " or "
''                End If
''            Next
''            If strFilters <> "" Then
''                rsM.Filter = strFilters
''                Set vsfFiles.DataSource = rsM
''            End If
''
''        End If
''    Else
''        '搜索使用
''    End If
''    Exit Sub
''errH:
''    If ErrCenter() = 1 Then Resume
''    Call SaveErrLog
''End Sub
