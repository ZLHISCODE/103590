VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAppCheck 
   BackColor       =   &H00FFFFFF&
   Caption         =   "对象检查修复"
   ClientHeight    =   8835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13980
   ControlBox      =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmAppCheck.frx":0000
   ScaleHeight     =   8835
   ScaleWidth      =   13980
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkProcedure 
      BackColor       =   &H8000000E&
      Caption         =   "检查过程/函数的有效性"
      Height          =   495
      Left            =   9120
      TabIndex        =   16
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CheckBox chkParameters 
      BackColor       =   &H8000000E&
      Caption         =   "检查参数名称的一致性"
      Height          =   375
      Left            =   11520
      TabIndex        =   15
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "对象检查修复"
      Height          =   465
      Index           =   0
      Left            =   1080
      TabIndex        =   13
      Top             =   3600
      Width           =   1890
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "公共同义词修正"
      Height          =   465
      Index           =   1
      Left            =   1080
      TabIndex        =   12
      Top             =   4440
      Width           =   1890
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "管理工具权限修正"
      Height          =   465
      Index           =   2
      Left            =   1080
      TabIndex        =   11
      Top             =   5400
      Width           =   1890
   End
   Begin VB.CheckBox chkIndex 
      BackColor       =   &H8000000E&
      Caption         =   "检查索引表空间的一致性"
      Height          =   465
      Left            =   3240
      TabIndex        =   10
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CheckBox chkReport 
      BackColor       =   &H8000000E&
      Caption         =   "检查当前版本中报表是否存在"
      Height          =   465
      Left            =   6360
      TabIndex        =   9
      Top             =   3600
      Width           =   2775
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   13920
      TabIndex        =   1
      Top             =   8295
      Visible         =   0   'False
      Width           =   13980
      Begin MSComctlLib.ProgressBar pgbState 
         Height          =   180
         Left            =   135
         TabIndex        =   2
         Top             =   255
         Width           =   6180
         _ExtentX        =   10901
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar pgbProgress 
         Height          =   180
         Left            =   7080
         TabIndex        =   3
         Top             =   255
         Width           =   6180
         _ExtentX        =   10901
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "正在检查"
         Height          =   180
         Left            =   135
         TabIndex        =   5
         Top             =   60
         Width           =   810
      End
      Begin VB.Label lblProgress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "已完成："
         Height          =   180
         Left            =   6840
         TabIndex        =   4
         Top             =   0
         Width           =   720
      End
      Begin VB.Line Linepgb 
         BorderColor     =   &H80000006&
         X1              =   6600
         X2              =   6600
         Y1              =   0
         Y2              =   720
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfSelSys 
      Height          =   1695
      Left            =   1080
      TabIndex        =   6
      Top             =   900
      Width           =   11175
      _cx             =   19711
      _cy             =   2990
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
      Rows            =   1
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   3240
      TabIndex        =   14
      Top             =   4200
      Width           =   7695
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   360
      Picture         =   "frmAppCheck.frx":803A
      Top             =   600
      Width           =   480
   End
   Begin VB.Label lblMainPath 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "系统安装目录：C:\Appsoft"
      Height          =   180
      Left            =   1080
      TabIndex        =   8
      Tag             =   "C:\Appsoft"
      Top             =   660
      Width           =   2160
   End
   Begin VB.Label lblSel 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "更改…"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   3420
      TabIndex        =   7
      Top             =   660
      Width           =   540
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "对象检查修复"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   315
      TabIndex        =   0
      Top             =   120
      Width           =   1440
   End
End
Attribute VB_Name = "frmAppCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MSTR_COL = ",300,4;编号,500,1;名称,2000,1;当前版本,1400,1;脚本版本,1400,1;配置文件,4050,1;所有者,0,1;共享号,0,1;,400,4"
Private Enum SysSelCol
    Col_选择 = 0
    Col_系统编号 = 1
    Col_系统名称 = 2
    Col_当前版本 = 3
    Col_脚本版本 = 4
    Col_配置文件 = 5
    Col_所有者 = 6
    Col_共享号 = 7
    Col_空白 = 8
End Enum
Private mclsRunScript As New clsRunScript
Private mrsLocalFile As New ADODB.Recordset

Private mrsSequenceFromFile As ADODB.Recordset
Private mrsViewFromFile As ADODB.Recordset
Private mrsPackageFromFile As ADODB.Recordset
Private mrsFildFromFile As ADODB.Recordset
Private mrsConstraintFromFile As ADODB.Recordset
Private mrsIndexFromFile As ADODB.Recordset
Private mrsProcedureFromFile As ADODB.Recordset

Private mrsSequenceFromDB As ADODB.Recordset
Private mrsViewFromDB As ADODB.Recordset
Private mrsPackageFromDB As ADODB.Recordset
Private mrsFildFromDB As ADODB.Recordset
Private mrsConstraintFromDB As ADODB.Recordset
Private mrsIndexFromDB As ADODB.Recordset
Private mrsProcedureFromDB As ADODB.Recordset

Private mrsDataFromFile As ADODB.Recordset
Private mrsDataFromDB As ADODB.Recordset

Private mlngSysNum As Long
Private mlngShare As Long
Private mlngProgress As Long
Private mblnzlTables As Boolean

Private Sub cmdFunction_Click(Index As Integer)
    Dim rsProData As ADODB.Recordset
    Dim rsChooseSysInfo As ADODB.Recordset
    Dim lngConsuming As Long
    Dim strSQL As String
    Dim strTemp As String
    Dim cnTools As ADODB.Connection
    Dim lngProgress As Long
    Dim strOwner As String
    
    If MsgBox("""" & Split(cmdFunction(Index).Caption, "(")(0) & """操作将可能消耗较多的资源和花费较长的时间，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    Select Case Index
        Case 0
            Set rsChooseSysInfo = CopyNewRec(Nothing, True, , _
                                Array("系统编号", adDouble, 10, 0, "系统名称", adVarChar, 50, Empty, _
                                      "当前版本", adVarChar, 20, Empty, "脚本版本", adVarChar, 20, Empty, _
                                      "配置文件", adVarChar, 500, Empty, "所有者", adVarChar, 50, Empty, _
                                      "共享号", adDouble, 10, 0))
            
            If CheckChoose(rsChooseSysInfo) = False Then Exit Sub
            '所需安装和升级脚本路径记录集初始化
            Set mrsLocalFile = IniFilePathRecordset

            Call InirsFile

            '数据记录集初始化
            Set mrsDataFromFile = InitDataRecordset
            
            mlngProgress = 3 + rsChooseSysInfo.RecordCount * 8
            picStatus.Visible = True
            Enabled = False
            Call ShowFinalPro(1)
            
            If CollectObj(rsChooseSysInfo) = False Then
                picStatus.Visible = False
                Enabled = True
                Exit Sub
            End If
            
            Set rsProData = InitProDataRecordset
            lngProgress = 3
            rsChooseSysInfo.MoveFirst
            Do While Not rsChooseSysInfo.EOF
                If strOwner <> rsChooseSysInfo!所有者 Then
                    strOwner = rsChooseSysInfo!所有者
                    Call CollectObjFromDB(strOwner, rsChooseSysInfo!系统编号)
                    Call GainData(mrsSequenceFromFile, mrsViewFromFile, mrsPackageFromFile, mrsFildFromFile, mrsConstraintFromFile, mrsIndexFromFile, mrsProcedureFromFile, mrsDataFromFile, _
                        mrsSequenceFromDB, mrsViewFromDB, mrsPackageFromDB, mrsFildFromDB, mrsConstraintFromDB, mrsIndexFromDB, mrsProcedureFromDB, mrsDataFromDB, _
                        IIf(chkIndex.value = 1, True, False), IIf(chkReport.value = 1, True, False), mblnzlTables, IIf(chkProcedure.value = 1, True, False), _
                        IIf(chkParameters.value = 1, True, False))
                End If
                    Call CompareCheck(rsChooseSysInfo!系统编号, rsChooseSysInfo!系统名称, rsProData, lngProgress)
                DoEvents
                rsChooseSysInfo.MoveNext
            Loop
            
            picStatus.Visible = False
            Enabled = True
            rsProData.Filter = ""
            If rsProData.RecordCount > 0 Then
                Call frmAppChkRpt.ShowMe(lblMainPath.Tag, rsProData, mrsDataFromFile)
            Else
                MsgBox "未检查出需修复的对象！"
            End If
            Call Release
        Case 1
            '创建当前所有者的全部对象的公共同义词('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION')
            gcnOracle.Execute "Zl_Createpubsynonyms", , adCmdStoredProc
            
            MsgBox "修正公共同义词完成！", vbInformation, gstrSysName
        Case 2
            '对象权限修正
            Set cnTools = GetConnection("ZLTOOLS")
            If cnTools Is Nothing Then Exit Sub
            Call ReGrantForTools(cnTools, , True)
            MsgBox "管理工具权限修正完成！", vbInformation, gstrSysName
    End Select
End Sub

Private Sub InirsFile()
'功能：初始化脚本数据记录集

    Set mrsSequenceFromFile = CopyNewRec(Nothing, True, , Array("系统编号", adDouble, 10, 0, "SQL", adVarChar, 2000, Empty, "名称", adVarChar, 100, Empty))
    Set mrsViewFromFile = CopyNewRec(Nothing, True, , Array("系统编号", adDouble, 10, 0, "SQL", adVarChar, 2000, Empty, "名称", adVarChar, 100, Empty))
    Set mrsPackageFromFile = CopyNewRec(Nothing, True, , Array("系统编号", adDouble, 10, 0, "SQL", adVarChar, 2000, Empty, "名称", adVarChar, 100, Empty, "STATUS", adVarChar, 20, Empty))
    Set mrsFildFromFile = CopyNewRec(Nothing, True, , Array("系统编号", adDouble, 10, 0, "SQL", adVarChar, 2000, Empty, "表名", adVarChar, 100, Empty, _
                        "字段", adVarChar, 200, Empty, "字段类型", adVarChar, 20, Empty, "字段长度", adVarChar, 10, Empty))
    Set mrsConstraintFromFile = CopyNewRec(Nothing, True, , Array("系统编号", adDouble, 10, 0, "SQL", adVarChar, 2000, Empty, "表名", adVarChar, 100, Empty, _
                        "名称", adVarChar, 100, Empty, "字段", adVarChar, 200, Empty, "表空间", adVarChar, 20, Empty))
    Set mrsIndexFromFile = CopyNewRec(Nothing, True, , Array("系统编号", adDouble, 10, 0, "SQL", adVarChar, 2000, Empty, "表名", adVarChar, 100, Empty, _
                        "名称", adVarChar, 100, Empty, "字段", adVarChar, 200, Empty, "表空间", adVarChar, 20, Empty))
    Set mrsProcedureFromFile = CopyNewRec(Nothing, True, , Array("系统编号", adDouble, 10, 0, "SQL", adVarChar, 2000, Empty, "名称", adVarChar, 100, Empty, "字段", adVarChar, 1000, Empty))
    
End Sub

Private Function CheckChoose(ByRef rsFileInfor As ADODB.Recordset) As Boolean
'功能：检查对象检查修复前是否勾选系统;勾选的系统是否存在本地配置文件;当前用户是否能够检查所勾选的系统
'参数：rsFileInfor：保存所选系统的列表内容
    Dim i As Long
    Dim strFile As String
    Dim blnFile As Boolean
    Dim strOraVer As String
    Dim strLocalVer As String
    Dim varTemp As Variant
    Dim cnTools As ADODB.Connection
    
    blnFile = False
    With vsfSelSys
        For i = .FixedRows To .Rows - .FixedRows
            If .Cell(flexcpChecked, i, Col_选择) = flexChecked Then
                If .TextMatrix(i, Col_配置文件) = "" Then
                    strFile = IIf(strFile = "", .TextMatrix(i, Col_系统名称), strFile & "、" & .TextMatrix(i, Col_系统名称))
                Else
                    strOraVer = VerFull(.TextMatrix(i, Col_当前版本))
                    strLocalVer = VerFull(.TextMatrix(i, Col_脚本版本))
                    If Split(strOraVer, ".")(1) = Split(strLocalVer, ".")(1) Then
                        varTemp = Split(.TextMatrix(i, Col_当前版本), ".")
                        If strOraVer > strLocalVer Then
                            MsgBox .TextMatrix(i, Col_系统名称) & "当前版本大于脚本版本，无法进行对象检查修复，请检查！"
                            Exit Function
                        End If
                        If UBound(varTemp) > 2 Then
                            If strOraVer <> strLocalVer Then
                                MsgBox .TextMatrix(i, Col_系统名称) & "为特殊sp版本，脚本版本必须与当前版本一致！"
                                Exit Function
                            End If
                        End If
                        If .TextMatrix(i, Col_系统名称) = "服务器管理工具" Then .TextMatrix(i, Col_所有者) = "ZLTOOLS"
                        rsFileInfor.AddNew Array("系统编号", "系统名称", "当前版本", "脚本版本", "配置文件", "所有者", "共享号"), Array(IIf(.TextMatrix(i, Col_系统编号) = "", 0, .TextMatrix(i, Col_系统编号)), _
                                .TextMatrix(i, Col_系统名称), .TextMatrix(i, Col_当前版本), .TextMatrix(i, Col_脚本版本), _
                                .TextMatrix(i, Col_配置文件), .TextMatrix(i, Col_所有者), IIf(.TextMatrix(i, Col_共享号) = "", 0, .TextMatrix(i, Col_共享号)))
                        blnFile = True
                    Else
                        MsgBox "脚本版本与当前版本的大版本不一致，无法进行检查！"
                        Exit Function
                    End If
                End If
            End If
            If i = .Rows - .FixedRows Then
                If strFile <> "" Then
                    MsgBox strFile & "的本地配置文件不存在，无法进行对象检查修复，请检查！"
                    Exit Function
                End If
                If blnFile = False Then
                    MsgBox "没有勾选系统，无法进行对象检查修复，请选择！"
                    Exit Function
                End If
            End If
        Next
    End With
    
    If rsFileInfor.RecordCount > 0 Then rsFileInfor.MoveFirst
    Do While Not rsFileInfor.EOF
        If gstrUserName = "ZLTOOLS" Then
            If rsFileInfor!系统名称 <> "服务器管理工具" Then
                MsgBox "ZLTOOLS用户只能检查服务器管理工具，请重新勾选系统或切换用户！"
                Exit Function
            End If
        ElseIf gblnDBA Then
            
        Else
            If rsFileInfor!所有者 <> gstrUserName Then
                If rsFileInfor!系统名称 = "服务器管理工具" Then
                    If gcnTools Is Nothing Then
                        MsgBox gstrUserName & "不是DBA用户，也不是ZLTOOLS用户，需连接管理工具用户才能对管理工具进行检查！"
                        Set gcnTools = GetConnection("ZLTOOLS")
                        If gcnTools Is Nothing Then
                            MsgBox "管理工具用户连接失败，无法进行检查！"
                            Exit Function
                        End If
                    End If
                Else
                    MsgBox gstrUserName & "不是DBA用户，也不是" & rsFileInfor!系统名称 & "的所有者，不能进行该系统的检查，请重新勾选系统或切换用户！"
                    Exit Function
                End If
            End If
        End If
        rsFileInfor.MoveNext
    Loop
    CheckChoose = True
End Function

Private Function CollectObj(ByVal rsChoose As ADODB.Recordset) As Boolean
'功能：保存脚本解析和数据库获取的对象信息
'参数：rsChoose：所勾选的系统信息
    Dim varTemp As Variant
    Dim strBigVer As String
    Dim rsTemp As ADODB.Recordset
    Dim strOwner As String
    Dim strSQL As String
    Dim rsUpgrade As New ADODB.Recordset
    Dim strFilePath As String
    Dim strSPInfo As String

    strSQL = "Select Nvl(a.系统, 0) 系统, Nvl(b.名称, '服务器管理工具') 系统名称, 结果版本" & vbNewLine & _
            "From Zlupgrade a, Zlsystems b" & vbNewLine & _
            "Where Length(a.结果版本) > 10 And a.升迁结果 = 0 And a.系统 = b.编号(+)" & vbNewLine & _
            "Order By 系统"
    Set rsUpgrade = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "数据库升迁情况")
    Do While Not rsUpgrade.EOF
        strSPInfo = strSPInfo & rsUpgrade!系统名称 & vbTab & rsUpgrade!结果版本 & vbCrLf
        rsUpgrade.MoveNext
    Loop
    If rsUpgrade.RecordCount > 0 Then
        If MsgBox("当前版本在升级过程中含有特殊sp的升级，请确保以下系统的特殊sp文件存在再进行检查！" & vbCrLf & strSPInfo, vbDefaultButton2 + vbYesNo, "提示") = vbNo Then
            Exit Function
        End If
    End If
    
    With rsChoose
        .MoveFirst
        Do While Not .EOF
            If CheckSetFile(!配置文件, !系统编号) Then
                varTemp = Split(!当前版本, ".")
                If varTemp(2) <> 0 Then
                    strBigVer = varTemp(0) & "." & varTemp(1) & ".0"
                    '由于公共函数(GetUpgradeFiles)获取的是当前当前版本到当前脚本的文件，这里利用该函数获取当前数据库的大版本到本地脚本最大版本，然后删除版本大于当前版本的脚本文件路径
                    Set rsTemp = Nothing
                    Set rsTemp = GetUpgradeFiles(rsTemp, !系统编号, strBigVer, !配置文件, , , , , , , False)
                    Call AddUpFile(rsTemp, !当前版本, !共享号)
                End If
            Else
                picStatus.Visible = False
                Enabled = True
                CollectObj = False
                Exit Function
            End If
            pgbState.value = .AbsolutePosition / .RecordCount * 100
            DoEvents
            .MoveNext
        Loop
    End With
        
    Call ShowFinalPro(2)
    
    With mrsLocalFile
        .MoveFirst
        Do While Not .EOF
            lblStatus.Caption = "正在收集脚本对象信息：" & !FilePath
            mlngSysNum = !SystemNum
            mlngShare = IIf(IsNull(!共享号) = True, 0, !共享号)
            Call CollectObjFromFile
            Call DealSpeObj
            pgbState.value = .AbsolutePosition / .RecordCount * 100
            DoEvents
            .MoveNext
        Loop
    End With
    Call CollectDataFromDB
    Call ShowFinalPro(3)
    CollectObj = True
End Function

Private Sub DealSpeObj()
'功能：删除特殊情况的对象，如optional脚本等
    Dim strFilter As String
    Dim strSQL As String
    
    If mlngSysNum = 100 Then
        If mrsLocalFile!Filename = "ZL1_10.35.30.SQL" Then
            strFilter = "表名='影像收藏类别' and 字段='创建人' and 系统编号=" & mlngSysNum
            Call RecDelete(mrsFildFromFile, strFilter)
        ElseIf mrsLocalFile!Filename = "ZL1_10.35.60.SQL" Then
            strFilter = "名称='影像临时记录_IX_检查号' and 系统编号=" & mlngSysNum
            Call RecDelete(mrsIndexFromFile, strFilter)
        ElseIf mrsLocalFile!Filename = "ZL1_10.35.80.SQL" Then
            mrsFildFromFile.Filter = "表名='消费卡目录'"
            Do While Not mrsFildFromFile.EOF
                mrsFildFromFile!表名 = "消费卡信息"
                mrsFildFromFile!SQL = Replace(mrsFildFromFile!SQL, "消费卡目录", "消费卡信息")
                mrsFildFromFile.Update
                mrsFildFromFile.MoveNext
            Loop
            mrsConstraintFromFile.Filter = "名称 like '消费卡目录*'"
            Do While Not mrsConstraintFromFile.EOF
                mrsConstraintFromFile!表名 = "消费卡信息"
                mrsConstraintFromFile!名称 = Replace(mrsConstraintFromFile!名称, "消费卡目录", "消费卡信息")
                mrsConstraintFromFile!SQL = Replace(mrsConstraintFromFile!SQL, "消费卡目录", "消费卡信息")
                mrsConstraintFromFile.Update
                mrsConstraintFromFile.MoveNext
            Loop
            mrsIndexFromFile.Filter = "名称 like '消费卡目录*'"
            Do While Not mrsIndexFromFile.EOF
                mrsIndexFromFile!表名 = "消费卡信息"
                mrsIndexFromFile!名称 = Replace(mrsIndexFromFile!名称, "消费卡目录", "消费卡信息")
                mrsIndexFromFile!SQL = Replace(mrsIndexFromFile!SQL, "消费卡目录", "消费卡信息")
                mrsIndexFromFile.Update
                mrsIndexFromFile.MoveNext
            Loop
            mrsSequenceFromFile.Filter = "名称 like '消费卡目录*'"
            Do While Not mrsSequenceFromFile.EOF
                mrsSequenceFromFile!名称 = Replace(mrsSequenceFromFile!名称, "消费卡目录", "消费卡信息")
                mrsSequenceFromFile!SQL = Replace(mrsSequenceFromFile!SQL, "消费卡目录", "消费卡信息")
                mrsSequenceFromFile.Update
                mrsSequenceFromFile.MoveNext
            Loop
            
            mrsFildFromFile.Filter = "表名='卡消费接口目录'"
            Do While Not mrsFildFromFile.EOF
                mrsFildFromFile!表名 = "消费卡类别目录"
                mrsFildFromFile!SQL = Replace(mrsFildFromFile!SQL, "卡消费接口目录", "消费卡类别目录")
                mrsFildFromFile.Update
                mrsFildFromFile.MoveNext
            Loop
            mrsConstraintFromFile.Filter = "名称 like '卡消费接口目录*'"
            Do While Not mrsConstraintFromFile.EOF
                mrsConstraintFromFile!表名 = "消费卡类别目录"
                mrsConstraintFromFile!名称 = Replace(mrsConstraintFromFile!名称, "卡消费接口目录", "消费卡类别目录")
                mrsConstraintFromFile!SQL = Replace(mrsConstraintFromFile!SQL, "卡消费接口目录", "消费卡类别目录")
                mrsConstraintFromFile.Update
                mrsConstraintFromFile.MoveNext
            Loop
            mrsIndexFromFile.Filter = "名称 like '卡消费接口目录*'"
            Do While Not mrsIndexFromFile.EOF
                mrsIndexFromFile!表名 = "消费卡类别目录"
                mrsIndexFromFile!名称 = Replace(mrsIndexFromFile!名称, "卡消费接口目录", "消费卡类别目录")
                mrsIndexFromFile!SQL = Replace(mrsIndexFromFile!SQL, "卡消费接口目录", "消费卡类别目录")
                mrsIndexFromFile.Update
                mrsIndexFromFile.MoveNext
            Loop
            
            strFilter = "表名='消费卡充值记录'"
            Call RecDelete(mrsFildFromFile, strFilter)
            strFilter = "表名='消费卡充值记录' or 名称 like '消费卡充值记录*'"
            Call RecDelete(mrsConstraintFromFile, strFilter)
            strFilter = "表名='消费卡充值记录' or 名称 like '消费卡充值记录*'"
            Call RecDelete(mrsIndexFromFile, strFilter)
            strFilter = "名称 like '消费卡充值记录*'"
            Call RecDelete(mrsSequenceFromFile, strFilter)
            
            strFilter = "表名='病人卡结算对照'"
            Call RecDelete(mrsFildFromFile, strFilter)
            strFilter = "表名='病人卡结算对照' or 名称 like '病人卡结算对照*'"
            Call RecDelete(mrsConstraintFromFile, strFilter)
            strFilter = "表名='病人卡结算对照' or 名称 like '病人卡结算对照*'"
            Call RecDelete(mrsIndexFromFile, strFilter)
            
            '删除消费卡信息中的字段：结算方式、缴款组ID、单位开户行、单位帐号、结算号码
            strFilter = "(表名='消费卡信息' and 字段='结算方式') or (表名='消费卡信息' and 字段='缴款组ID') or (表名='消费卡信息' and 字段='单位开户行') or (表名='消费卡信息' and 字段='单位帐号') or (表名='消费卡信息' and 字段='结算号码')"
            Call RecDelete(mrsFildFromFile, strFilter)
            
            strFilter = "表名='消费卡信息' and 名称='消费卡信息_FK_缴款组ID'"
            Call RecDelete(mrsConstraintFromFile, strFilter)
            
            If mblnzlTables Then
                strFilter = "(对象='卡消费接口目录' and 类别='表目录') or (对象='病人卡结算对照' and 类别='表目录') or (对象='消费卡充值记录' and 类别='表目录')"
                Call RecDelete(mrsDataFromFile, strFilter)
            End If
        End If
    ElseIf mlngSysNum = 0 Then
        If mrsLocalFile!Filename = "ZLUPGRADE10.35.30.SQL" Then
            strFilter = "表名='ZLRPTRUNHISTORY' and 字段='执行人员ID' and 系统编号=" & mlngSysNum
            Call RecDelete(mrsFildFromFile, strFilter)
            strFilter = "表名='ZLREPORTS' and 字段='执行人员ID' and 系统编号=" & mlngSysNum
            Call RecDelete(mrsFildFromFile, strFilter)
        ElseIf mrsLocalFile!Filename = "ZLUPGRADE10.35.90.SQL" Then
            strFilter = "表名='ZLPERIODS' and 系统编号=" & mlngSysNum
            Call RecDelete(mrsFildFromFile, strFilter)
            strFilter = "对象='ZLPERIODS' and 系统编号=" & mlngSysNum
            Call RecDelete(mrsDataFromFile, strFilter)
            strFilter = "名称 like 'ZLPERIODS*' and 系统编号=" & mlngSysNum
            Call RecDelete(mrsConstraintFromFile, strFilter)
            strFilter = "名称 like 'ZLPERIODS*' and 系统编号=" & mlngSysNum
            Call RecDelete(mrsIndexFromFile, strFilter)
        End If
    ElseIf mlngSysNum = 2100 Then
        If mrsLocalFile!Filename = "ZL21_10.35.10.SQL" Then
            strFilter = "表名='体检任务人员' and 字段='指引单打印'"
            Call RecDelete(mrsFildFromFile, strFilter)
        End If
    ElseIf mlngSysNum = 2200 Then
        If mrsLocalFile!Filename = "ZL22_10.35.80.SQL" Then
            strFilter = "表名='血液配血单据'"
            Call RecDelete(mrsFildFromFile, strFilter)
            strFilter = "表名='血液配血单据' or 名称 like '血液配血单据*'"
            Call RecDelete(mrsConstraintFromFile, strFilter)
            strFilter = "表名='血液配血单据' or 名称 like '血液配血单据*'"
            Call RecDelete(mrsIndexFromFile, strFilter)
            
            strFilter = "对象='血液配血单据'"
            Call RecDelete(mrsDataFromFile, strFilter)
        End If
    ElseIf mlngSysNum = 2400 Then
        If mrsLocalFile!Filename = "ZL24_10.35.60.SQL" Then
            '这两条SQL在匿名块中
'            strFilter = "(名称='手术性质分类_PK' and 系统编号=" & mlngSysNum & ") or (名称='手术性质分类_UQ_名称' and 系统编号=" & mlngSysNum & ")"
'            Call RecDelete(mrsConstraintFromFile, strFilter)
            strSQL = "Alter Table 手术性质分类 Add Constraint 手术性质分类_PK Primary Key (编码) Using Index Pctfree 5 Tablespace zl9indexhis"
            mrsConstraintFromFile.AddNew Array("系统编号", "SQL", "表名", "名称", "字段", "表空间"), _
                            Array(mlngSysNum, strSQL, "手术性质分类", "手术性质分类_PK", "编码", "ZL9INDEXHIS")
            mrsIndexFromFile.AddNew Array("系统编号", "SQL", "表名", "名称", "字段", "表空间"), _
                            Array(mlngSysNum, strSQL, "手术性质分类", "手术性质分类_PK", "编码", "ZL9INDEXHIS")
            strSQL = "Alter Table 手术性质分类 Add Constraint 手术性质分类_UQ_名称 Unique (名称) Using Index Pctfree 5 Tablespace zl9indexhis"
            mrsConstraintFromFile.AddNew Array("系统编号", "SQL", "表名", "名称", "字段", "表空间"), _
                            Array(mlngSysNum, strSQL, "手术性质分类", "手术性质分类_UQ_名称", "名称", "ZL9INDEXHIS")
            mrsIndexFromFile.AddNew Array("系统编号", "SQL", "表名", "名称", "字段", "表空间"), _
                            Array(mlngSysNum, strSQL, "手术性质分类", "手术性质分类_UQ_名称", "名称", "ZL9INDEXHIS")
        End If
    ElseIf mlngSysNum = 2600 Then
        If mrsLocalFile!Filename = "ZL26_10.35.60.SQL" Then
            '这两条SQL在匿名块中
'            strFilter = "(名称='导诊属性选项_PK' and 系统编号=" & mlngSysNum & ") or (名称='导诊播报发布_PK' and 系统编号=" & mlngSysNum & ")"
'            Call RecDelete(mrsConstraintFromFile, strFilter)
            strSQL = "Alter Table 导诊属性选项 Add Constraint 导诊属性选项_PK Primary Key (属性目录id,编码) Using Index  Tablespace zl9IndexPss"
            mrsConstraintFromFile.AddNew Array("系统编号", "SQL", "表名", "名称", "字段", "表空间"), _
                            Array(mlngSysNum, strSQL, "导诊属性选项", "导诊属性选项_PK", "属性目录ID,编码", "ZL9INDEXPSS")
            mrsIndexFromFile.AddNew Array("系统编号", "SQL", "表名", "名称", "字段", "表空间"), _
                            Array(mlngSysNum, strSQL, "导诊属性选项", "导诊属性选项_PK", "属性目录ID,编码", "ZL9INDEXPSS")
            strSQL = "Alter Table 导诊播报发布 Add Constraint 导诊播报发布_PK Primary Key (播报目录ID,发布序号) Using Index  Tablespace zl9IndexPss"
            mrsConstraintFromFile.AddNew Array("系统编号", "SQL", "表名", "名称", "字段", "表空间"), _
                            Array(mlngSysNum, strSQL, "导诊播报发布", "导诊播报发布_PK", "播报目录ID,发布序号", "ZL9INDEXPSS")
            mrsIndexFromFile.AddNew Array("系统编号", "SQL", "表名", "名称", "字段", "表空间"), _
                            Array(mlngSysNum, strSQL, "导诊播报发布", "导诊播报发布_PK", "播报目录ID,发布序号", "ZL9INDEXPSS")
        End If
    ElseIf mlngSysNum = 300 Then
        If mlngShare <> 0 Then
            strFilter = "表名='挂号项目'"
            Call RecDelete(mrsFildFromFile, strFilter)
            strFilter = "名称 like '挂号项目*'"
            Call RecDelete(mrsConstraintFromFile, strFilter)
            strFilter = "名称 like '挂号项目*'"
            Call RecDelete(mrsIndexFromFile, strFilter)
            '删除共享时，对于的模块和功能数据
            strFilter = "(类别='模块' and 序号=1001 and 对象='部门管理' and 系统编号=300) or (类别='模块' and 序号=1002 and 对象='人员管理' and 系统编号=300) or (类别='模块' and 序号=1013 and 对象='疾病编码管理' and 系统编号=300)" & _
                " or (类别='功能' and 系统编号=300 and 序号=1001) or (类别='功能' and 系统编号=300 and 序号=1002) or (类别='功能' and 系统编号=300 and 序号=1013)"
            Call RecDelete(mrsDataFromFile, strFilter)
        End If
    ElseIf mlngSysNum = 400 Then
        strFilter = "(类别='模块' and 序号=1001 and 对象='部门管理' and 系统编号=400) or (类别='模块' and 序号=1002 and 对象='人员管理' and 系统编号=400) or (类别='模块' and 序号=1010 and 对象='期间划分调整' and 系统编号=400) or (类别='模块' and 序号=1025 and 对象='供货商管理' and 系统编号=400)" & _
            " or (类别='功能' and 系统编号=400 and 序号=1001) or (类别='功能' and 系统编号=400 and 序号=1002) or (类别='功能' and 系统编号=400 and 序号=1010) or (类别='功能' and 系统编号=400 and 序号=1025)"
        Call RecDelete(mrsDataFromFile, strFilter)
    ElseIf mlngSysNum = 600 Then
        strFilter = "(类别='模块' and 序号=1001 and 对象='部门管理' and 系统编号=600) or (类别='模块' and 序号=1002 and 对象='人员管理' and 系统编号=600) or (类别='模块' and 序号=1010 and 对象='期间划分调整' and 系统编号=600) or (类别='模块' and 序号=1025 and 对象='供货商管理' and 系统编号=600)" & _
            " or (类别='功能' and 系统编号=600 and 序号=1001) or (类别='功能' and 系统编号=600 and 序号=1002) or (类别='功能' and 系统编号=600 and 序号=1010) or (类别='功能' and 系统编号=600 and 序号=1025)"
        Call RecDelete(mrsDataFromFile, strFilter)
    End If
End Sub

Public Sub ShowProgress(ByRef strSysName As String, ByRef lngNum As Long, ByRef lngCurNum As Long, ByRef strObjType As String, Optional ByRef strName As String)
'功能：每检查一类对象显示进度条
        
    lblStatus.Caption = IIf(strName = "", "正在检查" & strSysName & "的" & strObjType & "...", "正在检查" & strSysName & "的" & strObjType & "：" & strName)
    pgbState.value = lngCurNum / lngNum * 100
    If pgbState.value = 100 Then pgbState.value = 0
End Sub

Public Sub ShowFinalPro(ByRef lngNum As Long)
'功能：显示总的进度条
    lblProgress.Caption = "已完成：" & Round(lngNum / mlngProgress * 100) & "%"
    pgbProgress.value = lngNum / mlngProgress * 100
    If pgbProgress.value = 100 Then
        pgbProgress.value = 0
        lblProgress.Caption = ""
    End If
End Sub

Private Sub AddUpFile(ByVal rsTemp As ADODB.Recordset, ByVal strCurrver As String, ByVal lngShare As Long)
'功能：获取到的升级脚本路径到模块脚本路径记录集中
'参数：rsTemp：根据公共函数获取的升级脚本;strCurrver：当前当前版本
    Dim i As Long
    Dim strVer As String
    Dim strFilter As String

    strVer = VerFull(strCurrver)
    strFilter = "FullSPVer>" & strVer & " or FileName like '*HISTORY*' or FileName like '*OPTIONAL*'"
    Call RecDelete(rsTemp, strFilter)
    
    rsTemp.Filter = "": rsTemp.Sort = "FullSPVer Asc"
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
        mrsLocalFile.AddNew Array("FilePath", "SystemNum", "FileName", "FileType", "FullVer", "共享号"), Array(rsTemp!FilePath, rsTemp!系统编号, UCase(rsTemp!Filename), "升级脚本", rsTemp!FullSPVer, lngShare)
        rsTemp.MoveNext
    Next
End Sub

Private Sub CollectObjFromDB(ByRef strOwner As String, ByRef lngNum As Long)
'功能：获取数据库的对象信息
    Dim strSQL As String
    Dim cnChoose As New ADODB.Connection
    
    If gblnDBA Then
        strSQL = "select '序列' 类别,a.SEQUENCE_NAME 名称 from Dba_SEQUENCES a where a.SEQUENCE_OWNER='" & strOwner & "'"
        Set mrsSequenceFromDB = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "序列")
        strSQL = "select '视图' 类别,a.Object_Name 名称 from Dba_Objects a where  a.OBJECT_TYPE Like 'VIEW' and a.owner ='" & strOwner & "'"
        Set mrsViewFromDB = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "视图")
        strSQL = "select '包' 类别,a.Object_Name 名称,a.STATUS from Dba_Objects a where a.OBJECT_TYPE in('PACKAGE') and OWNER='" & strOwner & "'"
        Set mrsPackageFromDB = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "包")
        strSQL = "select '字段' 类别, a.TABLE_NAME 表名,a.COLUMN_NAME 名称,a.COLUMN_NAME 字段,a.DATA_TYPE 字段类型,a.DATA_LENGTH 字段长度,a.DATA_PRECISION 字段实际长度,a.DATA_SCALE 字段小数长度 From DBA_TAB_COLUMNS a WHERE a.OWNER='" & strOwner & "'"
        Set mrsFildFromDB = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "字段")
        strSQL = "Select 类别, 表名, 名称, f_List2str(Cast(Collect(字段 Order By Position) As t_Strlist)) 字段, Status" & vbNewLine & _
                "From (Select '约束' 类别, a.Table_Name 表名, a.Constraint_Name 名称, b.Column_Name 字段, b.Position Position, a.Status" & vbNewLine & _
                "       From Dba_Constraints a, Dba_Cons_Columns b" & vbNewLine & _
                "       Where a.Owner = b.Owner And a.Constraint_Name = b.Constraint_Name And a.Constraint_Type In ('R', 'P', 'U') And" & vbNewLine & _
                "             a.Owner = '" & strOwner & "'" & vbNewLine & _
                "       Order By a.Constraint_Name, b.Position)" & vbNewLine & _
                "Group By 类别, 表名, 名称, Status"
        Set mrsConstraintFromDB = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "约束")
        strSQL = "Select 类别, 表名, 名称, f_List2str(Cast(Collect(字段 Order By Position) As t_Strlist)) 字段, 表空间, Uniqueness, Status" & vbNewLine & _
                "From (Select '索引' 类别, a.Table_Name 表名, a.Index_Name 名称, a.Column_Name 字段, b.Tablespace_Name 表空间, b.Uniqueness Uniqueness," & vbNewLine & _
                "              b.Status, a.Column_Position Position" & vbNewLine & _
                "       From All_Ind_Columns a, Dba_Indexes b" & vbNewLine & _
                "       Where a.Index_Name = b.Index_Name And a.Index_Owner = b.Owner And a.Index_Name Not Like '%$%' And" & vbNewLine & _
                "             b.Owner ='" & strOwner & "'" & vbNewLine & _
                "       Order By a.Index_Name, a.Column_Position)" & vbNewLine & _
                "Group By 类别, 表名, 名称, 表空间, Uniqueness, Status"
        Set mrsIndexFromDB = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "索引")
        strSQL = "Select 类别, 名称,f_List2str(Cast(Collect(字段 Order By Position) As t_Strlist)) 字段,Status" & vbNewLine & _
                "From(Select '过程/函数' 类别, b.Object_Name 名称, a.Argument_Name 字段,b.Status,a.Position Position" & vbNewLine & _
                "From Dba_Arguments a, Dba_Objects b" & vbNewLine & _
                "Where a.Package_Name Is Null And a.Object_Id(+) = b.Object_Id And" & vbNewLine & _
                "b.Object_Type In ('FUNCTION', 'PROCEDURE') And Not (a.Argument_Name Is Null And a.Data_Type Is Not Null) And" & vbNewLine & _
                "b.Owner ='" & strOwner & "'" & vbNewLine & _
                "Order By b.Object_Name,a.Position)" & vbNewLine & _
                "Group By  名称,类别,Status"
        Set mrsProcedureFromDB = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "过程/函数")
    Else
        If lngNum = 0 And gstrUserName <> "ZLTOOLS" Then
            Set cnChoose = gcnTools
        Else
            Set cnChoose = gcnOracle
        End If
        strSQL = "select '序列' 类别,a.SEQUENCE_NAME 名称 from User_SEQUENCES a"
        Set mrsSequenceFromDB = gclsBase.OpenSQLRecord(cnChoose, strSQL, "序列")
        strSQL = "select '视图' 类别,a.Object_Name 名称 from User_Objects a where  a.OBJECT_TYPE Like 'VIEW'"
        Set mrsViewFromDB = gclsBase.OpenSQLRecord(cnChoose, strSQL, "视图")
        strSQL = "select '包' 类别,a.Object_Name 名称,a.STATUS from User_Objects a where a.OBJECT_TYPE in('PACKAGE')"
        Set mrsPackageFromDB = gclsBase.OpenSQLRecord(cnChoose, strSQL, "包")
        strSQL = "select '字段' 类别, a.TABLE_NAME 表名,a.COLUMN_NAME 名称,a.COLUMN_NAME 字段,a.DATA_TYPE 字段类型,a.DATA_LENGTH 字段长度,a.DATA_PRECISION 字段实际长度,a.DATA_SCALE 字段小数长度 From User_TAB_COLUMNS a"
        Set mrsFildFromDB = gclsBase.OpenSQLRecord(cnChoose, strSQL, "字段")
        strSQL = "Select 类别, 表名, 名称, f_List2str(Cast(Collect(字段 Order By Position) As t_Strlist)) 字段, Status" & vbNewLine & _
                "From (Select '约束' 类别, a.Table_Name 表名, a.Constraint_Name 名称, b.Column_Name 字段, b.Position Position, a.Status" & vbNewLine & _
                "       From User_Constraints a, User_Cons_Columns b" & vbNewLine & _
                "       Where a.Constraint_Name = b.Constraint_Name And a.Constraint_Type In ('R', 'P', 'U')" & vbNewLine & _
                "       Order By a.Constraint_Name, b.Position)" & vbNewLine & _
                "Group By 类别, 表名, 名称, Status"
        Set mrsConstraintFromDB = gclsBase.OpenSQLRecord(cnChoose, strSQL, "约束")
        strSQL = "Select 类别, 表名, 名称, f_List2str(Cast(Collect(字段 Order By Position) As t_Strlist)) 字段, 表空间, Uniqueness, Status" & vbNewLine & _
                "From (Select '索引' 类别, a.Table_Name 表名, a.Index_Name 名称, a.Column_Name 字段, b.Tablespace_Name 表空间, b.Uniqueness Uniqueness," & vbNewLine & _
                "              b.Status, a.Column_Position Position" & vbNewLine & _
                "       From User_Ind_Columns a, User_Indexes b" & vbNewLine & _
                "       Where a.Index_Name = b.Index_Name And a.Index_Name Not Like '%$%' " & vbNewLine & _
                "       Order By a.Index_Name, a.Column_Position)" & vbNewLine & _
                "Group By 类别, 表名, 名称, 表空间, Uniqueness, Status"
        Set mrsIndexFromDB = gclsBase.OpenSQLRecord(cnChoose, strSQL, "索引")
        strSQL = "Select 类别, 名称,f_List2str(Cast(Collect(字段 Order By Position) As t_Strlist)) 字段,Status" & vbNewLine & _
                "From(Select '过程/函数' 类别, b.Object_Name 名称, a.Argument_Name 字段,b.Status,a.Position Position" & vbNewLine & _
                "From User_Arguments a, User_Objects b" & vbNewLine & _
                "Where a.Package_Name Is Null And a.Object_Id(+) = b.Object_Id And" & vbNewLine & _
                "b.Object_Type In ('FUNCTION', 'PROCEDURE') And Not (a.Argument_Name Is Null And a.Data_Type Is Not Null) " & vbNewLine & _
                "Order By b.Object_Name,a.Position)" & vbNewLine & _
                "Group By  名称,类别,Status"
        Set mrsProcedureFromDB = gclsBase.OpenSQLRecord(cnChoose, strSQL, "过程/函数")
    End If
End Sub

Private Sub CollectDataFromDB()
'功能：数据库基础数据保存
    Dim strSQL As String
    
    If mblnzlTables Then
        strSQL = "Select '模块' 类别, Nvl(系统, 0) 系统编号, 序号, 标题 对象, Null 参数号, Null 参数名 From Zlprograms Union All" & vbNewLine & _
                "Select '功能' 类别, Nvl(系统, 0) 系统编号, 序号, 功能 对象, Null 参数号, Null 参数名 From Zlprogfuncs Union All" & vbNewLine & _
                "Select '参数' 类别, Nvl(系统, 0) 系统编号, Null 序号, Nvl(模块 || '', 'NULL') 对象, 参数号, Upper(参数名) 参数名 From Zlparameters Union All" & vbNewLine & _
                "Select '报表' 类别, Nvl(系统, 0) 系统编号, Null 序号, 编号 对象, Null 参数号, Null 参数名 From Zlreports Union All" & vbNewLine & _
                "Select '表目录' 类别, 系统 系统编号, Null 序号, 表名 对象, Null 参数号, Null 参数名 From Zltables"
    Else
        strSQL = "Select '模块' 类别, Nvl(系统, 0) 系统编号, 序号, 标题 对象, Null 参数号, Null 参数名 From Zlprograms Union All" & vbNewLine & _
            "Select '功能' 类别, Nvl(系统, 0) 系统编号, 序号, 功能 对象, Null 参数号, Null 参数名 From Zlprogfuncs Union All" & vbNewLine & _
            "Select '参数' 类别, Nvl(系统, 0) 系统编号, Null 序号, Nvl(模块 || '', 'NULL') 对象, 参数号, Upper(参数名) 参数名 From Zlparameters Union All" & vbNewLine & _
            "Select '报表' 类别, Nvl(系统, 0) 系统编号, Null 序号, 编号 对象, Null 参数号, Null 参数名 From Zlreports"
    End If
    lblStatus.Caption = "正在收集数据库的基础数据信息..."
    Set mrsDataFromDB = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "基础数据")
    
End Sub

Private Sub CollectObjFromFile()
    '存储本地脚本的数据对象
    Dim i As Long
    Dim lngSys As Long
    Dim strSQL As String
    Dim strTemp As String
    Dim varTemp As Variant
    Dim varFild As Variant
    Dim strName As String
    Dim objText As TextStream
    Dim strTableName As String
    Dim strFild As String
    Dim strFildType As String
    Dim strReFild As String
    Dim strTableSpace As String
    Dim strFildLength As String
    Dim rsTemp As ADODB.Recordset
    
    If mrsLocalFile!Filename = "ZLSEQUENCE.SQL" Then
        Set objText = gobjFile.OpenTextFile(mrsLocalFile!FilePath, ForReading)
        Do While Not objText.AtEndOfStream
            strSQL = objText.ReadLine
            If strSQL <> "" And Mid(strSQL, 1, 2) <> "--" Then
                strSQL = iniSQL(strSQL)
                If strSQL Like "CREATE SEQUENCE*" Then
                    strName = Trim(Replace(strSQL, "CREATE SEQUENCE", ""))
                    strName = Mid(strName, 1, InStr(strName, " ") - 1)
                    mrsSequenceFromFile.AddNew Array("系统编号", "SQL", "名称"), Array(mlngSysNum, strSQL, strName)
                End If
            End If
        Loop
        objText.Close
    ElseIf mrsLocalFile!Filename = "ZLTABLE.SQL" Then
        If mclsRunScript.OpenFile(mrsLocalFile!FilePath) Then
            Do While Not mclsRunScript.EOF
                strSQL = iniSQL(mclsRunScript.SQLInfo.SQL)
                If strSQL Like "CREATE TABLE*" Then
                    strSQL = Replace(strSQL, "NUMERIC", "NUMBER")
                    strTemp = Trim(Replace(strSQL, "CREATE TABLE", ""))
                    strTableName = Trim(Replace(Mid(strTemp, 1, InStr(strTemp, "(") - 1), vbCrLf, ""))
                    strTemp = Replace(strSQL, vbCrLf, "||")
                    If InStr(strTemp, "))") > 0 Then
                        strFild = Mid(strTemp, InStr(strTemp, "(") + 1, InStr(strTemp, "))") - InStr(strTemp, "("))
                    ElseIf InStr(strTemp, "||)") > 0 Then
                        strFild = Mid(strTemp, InStr(strTemp, "(") + 1, InStr(strTemp, "||)") - InStr(strTemp, "("))
                    ElseIf InStr(strTemp, ")||") > 0 Then
                        strFild = Mid(strTemp, InStr(strTemp, "(") + 1, InStr(strTemp, ")||") - InStr(strTemp, "("))
                    Else
                        strFild = Mid(strTemp, InStr(strTemp, "(") + 1)
                    End If
                    varTemp = Split(strFild, "||")
                    For i = LBound(varTemp) To UBound(varTemp)
                        varTemp(i) = Trim(varTemp(i))
                        If InStr(varTemp(i), "TABLESPACE") > 0 Then Exit For
                        If varTemp(i) <> "" And varTemp(i) <> ")" And InStr(varTemp(i), "TABLESPACE") = 0 Then
                            strFildType = ""
                            strFildLength = ""
                            strFild = TrimEx(Mid(varTemp(i), 1, InStr(varTemp(i), " ")))
                            strTemp = Trim(Mid(varTemp(i), InStr(varTemp(i), " ") + 1))
                            If InStr(strTemp, "DATE") > 0 Then
                                strFildType = "DATE"
                            ElseIf InStr(strTemp, "LONG RAW") > 0 Then
                                strFildType = "LONG RAW"
                            Else
                                If InStr(strTemp, ")") > 0 Then
                                    strTemp = Trim(Mid(strTemp, 1, InStr(strTemp, ")") - 1))
                                ElseIf InStr(strTemp, " ") > 0 Then
                                    strTemp = Trim(Mid(strTemp, 1, InStr(strTemp, " ") - 1))
                                End If
                                If InStr(strTemp, "(") > 0 Then
                                    strFildType = Mid(strTemp, 1, InStr(strTemp, "(") - 1)
                                    strFildLength = Mid(strTemp, InStr(strTemp, "(") + 1)
                                ElseIf InStr(strTemp, ",") > 0 Then
                                    strFildType = Mid(strTemp, 1, Len(strTemp) - 1)
                                ElseIf InStr(strTemp, ")") > 0 Then
                                    strFildType = Mid(strTemp, 1, InStr(strTemp, ")") - 1)
                                Else
                                    strFildType = strTemp
                                End If
                                strFildType = Trim(Replace(strFildType, "|", ""))
                                '字段的名称和字段列赋的相同的值
                            End If
                            If strFild <> "" Then
                                mrsFildFromFile.AddNew Array("系统编号", "SQL", "表名", "字段", "字段类型", "字段长度"), _
                                                    Array(mlngSysNum, strSQL, strTableName, strFild, strFildType, strFildLength)
                            End If
                        End If
                    Next
                End If
                DoEvents
                Call mclsRunScript.ReadNextSQL
            Loop
        End If
    ElseIf mrsLocalFile!Filename = "ZLCONSTRAINT.SQL" Then
        Set objText = gobjFile.OpenTextFile(mrsLocalFile!FilePath, ForReading)
        Do While Not objText.AtEndOfStream
            strSQL = objText.ReadLine
            If strSQL <> "" And Mid(strSQL, 1, 2) <> "--" Then
            strSQL = iniSQL(strSQL)
                If strSQL Like "ALTER TABLE * ADD CONSTRAINT*" Or strSQL Like "ALTER TABLE * MODIFY * CONSTRAINT*" Then
                    varTemp = Split(strSQL, "CONSTRAINT")
                    strTemp = Trim(Replace(varTemp(1), "CONSTRAINT", ""))
                    '获取约束名称
                    strName = TrimEx(Mid(strTemp, 1, InStr(strTemp, " ") - 1))
                    strTemp = Trim(Replace(varTemp(0), "ALTER TABLE", ""))
                    '获取表名
                    strTableName = Trim(Mid(strTemp, 1, InStr(strTemp, " ")))
                    If InStr(strSQL, "ADD") > 0 Then
                        '获取约束字段
                        strFild = Replace(Trim(Mid(strSQL, InStr(strSQL, "(") + 1, InStr(strSQL, ")") - InStr(strSQL, "(") - 1)), " ", "")
                        strTableSpace = GetTableSpace(strSQL)
                        If InStr(strSQL, "NOVALIDATE") = 0 Then strSQL = strSQL & " NOVALIDATE"
                        mrsConstraintFromFile.AddNew Array("系统编号", "SQL", "表名", "名称", "字段", "表空间"), _
                                            Array(mlngSysNum, strSQL, strTableName, strName, strFild, strTableSpace)
                        If InStr(strSQL, "PRIMARY") > 0 Or InStr(strSQL, "UNIQUE") > 0 Then
                            If strTableSpace <> "" Then
                                strTemp = "Create Unique Index " & strName & " On " & strTableName & "(" & strFild & ") Tablespace " & strTableSpace & " Nologging"
                                strTemp = strTemp & "||" & strSQL
                            Else
                                strTemp = "Create Unique Index " & strName & " On " & strTableName & "(" & strFild & ") Nologging"
                                strTemp = strTemp & "||" & strSQL
                            End If
                            mrsIndexFromFile.AddNew Array("系统编号", "SQL", "表名", "名称", "字段", "表空间"), _
                                Array(mlngSysNum, strTemp, strTableName, strName, strFild, strTableSpace)
                        End If
                    End If
                End If
            End If
        Loop
        objText.Close
    ElseIf mrsLocalFile!Filename = "ZLINDEX.SQL" Then
        Set objText = gobjFile.OpenTextFile(mrsLocalFile!FilePath, ForReading)
        Do While Not objText.AtEndOfStream
            strSQL = objText.ReadLine
            If strSQL <> "" And Mid(strSQL, 1, 2) <> "--" Then
                strSQL = iniSQL(strSQL)
                If strSQL Like "CREATE INDEX*" Then
                    varTemp = Split(strSQL, "ON")
                    strName = Trim(Replace(varTemp(0), "CREATE INDEX", ""))
                    strTableName = Trim(Mid(varTemp(1), 1, InStr(varTemp(1), "(") - 1))
                    strFild = Replace(Mid(varTemp(1), InStr(varTemp(1), "(") + 1, InStrRev(varTemp(1), ")") - InStr(varTemp(1), "(") - 1), " ", "")
                    strTableSpace = GetTableSpace(strSQL)
                    If InStr(strSQL, "NOLOGGING") = 0 Then strSQL = strSQL & " NOLOGGING"
                    mrsIndexFromFile.AddNew Array("系统编号", "SQL", "表名", "名称", "字段", "表空间"), _
                        Array(mlngSysNum, strSQL, strTableName, strName, strFild, strTableSpace)
                End If
            End If
        Loop
        objText.Close
    ElseIf mrsLocalFile!Filename = "ZLVIEW.SQL" Then
        If mclsRunScript.OpenFile(mrsLocalFile!FilePath) Then
            Do While Not mclsRunScript.EOF
                strSQL = iniSQL(mclsRunScript.SQLInfo.SQL)
                If strSQL <> "" And Mid(strSQL, 1, 2) <> "--" Then
                    If strSQL Like "CREATE OR REPLACE VIEW*" Then
                        strName = Trim(Replace(strSQL, "CREATE OR REPLACE VIEW", ""))
                        strName = Mid(strName, 1, InStr(strName, " ") - 1)
                        mrsViewFromFile.AddNew Array("系统编号", "SQL", "名称"), Array(mlngSysNum, strSQL, strName)
                    End If
                End If
                DoEvents
                Call mclsRunScript.ReadNextSQL
            Loop
        End If
    ElseIf mrsLocalFile!Filename = "ZLPACKAGE.SQL" Then
        If mclsRunScript.OpenFile(mrsLocalFile!FilePath) Then
            Do While Not mclsRunScript.EOF
                strSQL = iniSQL(mclsRunScript.SQLInfo.SQL)
                If strSQL <> "" Then
                    strSQL = iniSQL(strSQL)
                        If strSQL Like "CREATE OR REPLACE PACKAGE*" And Not strSQL Like "CREATE OR REPLACE PACKAGE BODY*" Then
                            strName = Trim(Replace(strSQL, "CREATE OR REPLACE PACKAGE", ""))
                            strName = Mid(strName, 1, InStr(strName, " ") - 1)
                            mrsPackageFromFile.AddNew Array("系统编号", "SQL", "名称"), Array(mlngSysNum, strSQL, strName)
                        End If
                End If
                DoEvents
                Call mclsRunScript.ReadNextSQL
            Loop
        End If
    ElseIf mrsLocalFile!Filename = "ZLPROGRAM.SQL" Then
        If mclsRunScript.OpenFile(mrsLocalFile!FilePath) Then
            Do While Not mclsRunScript.EOF
                strSQL = iniSQL(mclsRunScript.SQLInfo.SQL)
                If strSQL <> "" Then
                    If mlngSysNum = 2700 Then
                        If strSQL Like "CREATE OR REPLACE PROCEDURE ZLHIS.ZL_体检人员项目_REJECT*" Then
                            strSQL = Replace(strSQL, "ZLHIS.", "")
                        End If
                    End If
                    If strSQL Like "CREATE OR REPLACE PROCEDURE*" Then
                        strName = Trim(Replace(strSQL, "CREATE OR REPLACE PROCEDURE", ""))
                    Else
                        strName = Trim(Replace(strSQL, "CREATE OR REPLACE FUNCTION", ""))
                    End If
                    If InStr(strName, vbCrLf) > 0 Then strName = Mid(strName, 1, InStr(strName, vbCrLf) - 1)
                    If InStr(strName, "(") > 0 Then strName = Trim(Mid(strName, 1, InStr(strName, "(") - 1))
                    If InStr(strName, " ") > 0 Then strName = Trim(Mid(strName, 1, InStr(strName, " ") - 1))
                    strName = Trim(strName)
                    If InStr(strSQL, "(") - InStr(strSQL, strName) - Len(strName) < 5 And Mid(InStr(strSQL, "(") - InStr(strSQL, strName) - Len(strName), 1, 1) <> "-" Then
                        strTemp = Mid(strSQL, InStr(strSQL, "(") + 1, InStr(strSQL, ")") - InStr(strSQL, "(") - 1)
                        If InStr(strTemp, "= ','") > 0 Then
                            varTemp = Split(strTemp, vbCrLf)
                        Else
                            varTemp = Split(strTemp, ",")
                        End If
                        strFild = ""
                        For i = 0 To UBound(varTemp)
                            varTemp(i) = Trim(Replace(varTemp(i), vbCrLf, ""))
                            If varTemp(i) <> "" Then
                                strFild = IIf(strFild = "", Trim(Mid(varTemp(i), 1, InStr(varTemp(i), " ") - 1)), strFild & "," & Trim(Mid(varTemp(i), 1, InStr(varTemp(i), " ") - 1)))
                            End If
                        Next
                        mrsProcedureFromFile.AddNew Array("系统编号", "SQL", "名称", "字段"), Array(mlngSysNum, strSQL, strName, strFild)
                    ElseIf strSQL Like "CREATE OR REPLACE VIEW *" Then
                        strName = Trim(Replace(strSQL, "CREATE OR REPLACE VIEW", ""))
                        strName = Mid(strName, 1, InStr(strName, " ") - 1)
                        mrsViewFromFile.AddNew Array("系统编号", "SQL", "名称"), Array(mlngSysNum, mclsRunScript.SQLInfo.SQL, strName)
                    End If
                End If
                DoEvents
                Call mclsRunScript.ReadNextSQL
            Loop
        End If
    ElseIf mrsLocalFile!Filename = "ZLMANDATA.SQL" Then
        If mclsRunScript.OpenFile(mrsLocalFile!FilePath) Then
            Do While Not mclsRunScript.EOF
                strSQL = iniSQL(mclsRunScript.SQLInfo.SQL)
                If strSQL <> "" Then
                    If strSQL Like "INSERT INTO*" Then
                        If strSQL Like "INSERT INTO ZLTABLES*" Then
                            strTemp = Replace(Mid(strSQL, InStr(strSQL, "(") + 1, InStr(strSQL, ")") - InStr(strSQL, "(") - 1), " ", "")
                            varFild = Split(strTemp, ",")
                            Set rsTemp = SetSelectRecordset(strSQL, strTemp, varFild, "ZLTABLES")
                            If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                            Do While Not rsTemp.EOF
                                mrsDataFromFile.AddNew Array("类别", "SQL", "对象", "系统编号"), Array("表目录", rsTemp!修正SQL, rsTemp!表名, mlngSysNum)
                                rsTemp.MoveNext
                            Loop
                        ElseIf strSQL Like "INSERT INTO ZLPROGRAMS*" Then
                            strTemp = Replace(Mid(strSQL, InStr(strSQL, "(") + 1, InStr(strSQL, ")") - InStr(strSQL, "(") - 1), " ", "")
                            varFild = Split(strTemp, ",")
                            Set rsTemp = SetSelectRecordset(strSQL, strTemp, varFild, "ZLPROGRAMS")
                            If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                            Do While Not rsTemp.EOF
                                mrsDataFromFile.AddNew Array("类别", "SQL", "系统编号", "序号", "对象"), Array("模块", rsTemp!修正SQL, mlngSysNum, rsTemp!序号, rsTemp!标题)
                                rsTemp.MoveNext
                            Loop
                        ElseIf strSQL Like "INSERT INTO ZLPROGFUNCS*" Then
                            strTemp = Replace(Mid(strSQL, InStr(strSQL, "(") + 1, InStr(strSQL, ")") - InStr(strSQL, "(") - 1), " ", "")
                            varFild = Split(strTemp, ",")
                            Set rsTemp = SetSelectRecordset(strSQL, strTemp, varFild, "ZLPROGFUNCS")
                            If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                            Do While Not rsTemp.EOF
                                mrsDataFromFile.AddNew Array("类别", "SQL", "系统编号", "序号", "对象"), Array("功能", rsTemp!修正SQL, mlngSysNum, rsTemp!序号, rsTemp!功能)
                                rsTemp.MoveNext
                            Loop
                        ElseIf strSQL Like "INSERT INTO ZLPARAMETERS*" Then
                            If strSQL = "INSERT INTO ZLPARAMETERS(ID,系统,模块,私有,本机,授权,固定,部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明)" & vbNewLine & _
                                        "SELECT ZLPARAMETERS_ID.NEXTVAL,2500,2500,-NULL,-NULL,-NULL,-NULL,-NULL,A.* FROM (" & vbNewLine & _
                                        "SELECT 私有,本机,授权,固定,部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明 FROM ZLPARAMETERS WHERE 1 = 0 UNION ALL" & vbNewLine & _
                                        "SELECT 0,0,1,1,0,0,27,'电子签名认证中心','0','0','在对标本进行核收、审核时进行签名。','用户选择签名方式其0为不使用电子签名外，0以上的值表示使用不同的认证中心进行签名。'||CHR(13)||'0=不使用电子签名'||CHR(13)||'1=辽宁省数字证书认证中心'||CHR(13)||'2=广西省数字证书证中心'||CHR(13)||'3=重庆市数字证书认证中心'||CHR(13)||'4=山东省数字证书认证中心'||CHR(13)||'5=吉林中心医院认证中心'||CHR(13)||'6=吉林省医院认证中心',NULL,NULL,NULL FROM DUAL UNION ALL" & vbNewLine & _
                                        "SELECT 私有,本机,授权,固定,部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明 FROM ZLPARAMETERS WHERE 1 = 0) A" Then
                                strSQL = "INSERT INTO ZLPARAMETERS(ID,系统,模块,私有,本机,授权,固定,部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明)" & vbNewLine & _
                                        "SELECT ZLPARAMETERS_ID.NEXTVAL,2500,2500,A.* FROM (" & vbNewLine & _
                                        "SELECT 私有,本机,授权,固定,部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明 FROM ZLPARAMETERS WHERE 1 = 0 UNION ALL" & vbNewLine & _
                                        "SELECT 0,0,1,1,0,0,27,'电子签名认证中心','0','0','在对标本进行核收、审核时进行签名。','用户选择签名方式其0为不使用电子签名外，0以上的值表示使用不同的认证中心进行签名。'||CHR(13)||'0=不使用电子签名'||CHR(13)||'1=辽宁省数字证书认证中心'||CHR(13)||'2=广西省数字证书证中心'||CHR(13)||'3=重庆市数字证书认证中心'||CHR(13)||'4=山东省数字证书认证中心'||CHR(13)||'5=吉林中心医院认证中心'||CHR(13)||'6=吉林省医院认证中心',NULL,NULL,NULL FROM DUAL UNION ALL" & vbNewLine & _
                                        "SELECT 私有,本机,授权,固定,部门,性质,参数号,参数名,参数值,缺省值,影响控制说明,参数值含义,关联说明,适用说明,警告说明 FROM ZLPARAMETERS WHERE 1 = 0) A"
                            End If
                            strTemp = Replace(Mid(strSQL, InStr(strSQL, "(") + 1, InStr(strSQL, ")") - InStr(strSQL, "(") - 1), " ", "")
                            varFild = Split(strTemp, ",")
                            Set rsTemp = SetSelectRecordset(strSQL, strTemp, varFild, "ZLPARAMETERS")
                            If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                            Do While Not rsTemp.EOF
                                If IsNull(rsTemp!系统) Or InStr(rsTemp!系统, "NULL") > 0 Or rsTemp!模块 = """" Then
                                    lngSys = 0
                                Else
                                    lngSys = mlngSysNum
                                End If
                                If InStr(strTemp, "模块") > 0 Then
                                    If InStr(UCase(rsTemp!模块), "NULL") > 0 Or rsTemp!模块 = """" Then
                                        mrsDataFromFile.AddNew Array("类别", "SQL", "系统编号", "对象", "参数号", "参数名"), Array("参数", rsTemp!修正SQL, lngSys, "NULL", rsTemp!参数号, rsTemp!参数名)
                                    Else
                                        mrsDataFromFile.AddNew Array("类别", "SQL", "系统编号", "对象", "参数号", "参数名"), Array("参数", rsTemp!修正SQL, lngSys, rsTemp!模块, rsTemp!参数号, rsTemp!参数名)
                                    End If
                                Else
                                    mrsDataFromFile.AddNew Array("类别", "SQL", "系统编号", "对象", "参数号", "参数名"), Array("参数", rsTemp!修正SQL, lngSys, "NULL", rsTemp!参数号, rsTemp!参数名)
                                End If
                                rsTemp.MoveNext
                            Loop
                        End If
                    End If
                End If
                DoEvents
                Call mclsRunScript.ReadNextSQL
            Loop
        End If
    ElseIf mrsLocalFile!Filename = "ZLREPORT.SQL" Then
        If mclsRunScript.OpenFile(mrsLocalFile!FilePath) Then
            Do While Not mclsRunScript.EOF
                strSQL = UCase(mclsRunScript.SQLInfo.SQL)
                If strSQL Like "INSERT INTO ZLREPORTS*" Then
                    strSQL = iniSQL(strSQL)
                    varFild = Split(Replace(Mid(strSQL, InStr(strSQL, "(") + 1, InStr(strSQL, ")") - InStr(strSQL, "(") - 1), " ", ""), ",")
                    Set rsTemp = SetSelectRecordset(strSQL, strTemp, varFild, "ZLREPORTS")
                    If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                    Do While Not rsTemp.EOF
                        mrsDataFromFile.AddNew Array("类别", "SQL", "系统编号", "对象", "名称"), Array("报表", rsTemp!修正SQL, mlngSysNum, rsTemp!编号, rsTemp!名称)
                        rsTemp.MoveNext
                    Loop
                End If
                DoEvents
                Call mclsRunScript.ReadNextSQL
            Loop
        End If
    Else
        Call GetAnyObject
    End If
End Sub

Private Sub GetAnyObject()
'保存升级脚本中所有可能的SQL对象
    Dim strSQL As String
    Dim strIniSQL As String
    Dim strName As String
    Dim strTableName As String
    Dim strTableSpace As String
    Dim strFild As String
    Dim strFildType As String
    Dim strReFild As String
    Dim strFildLength As String
    Dim varTemp As Variant
    Dim i As Long
    Dim lngSys As Long
    Dim strFilter As String
    Dim strTemp As String
    Dim rsTemp As ADODB.Recordset
    
    If mclsRunScript.OpenFile(mrsLocalFile!FilePath) Then
        Do While Not mclsRunScript.EOF
            strSQL = mclsRunScript.SQLInfo.SQL
            strIniSQL = iniSQL(strSQL)
            If strIniSQL <> "" Then
                If strIniSQL Like "CREATE*" Then
                    If strIniSQL Like "CREATE OR REPLACE*" Then
                        If strIniSQL Like "CREATE OR REPLACE PROCEDURE*" Or strIniSQL Like "CREATE OR REPLACE FUNCTION*" Then
                            If strIniSQL Like "CREATE OR REPLACE PROCEDURE*" Then
                                strName = Trim(Replace(strIniSQL, "CREATE OR REPLACE PROCEDURE", ""))
                            Else
                                strName = Trim(Replace(strIniSQL, "CREATE OR REPLACE FUNCTION", ""))
                            End If
                            If InStr(strName, vbCrLf) > 0 Then strName = Mid(strName, 1, InStr(strName, vbCrLf) - 1)
                            If InStr(strName, " ") > 0 Then strName = Trim(Mid(strName, 1, InStr(strName, " ") - 1))
                            If InStr(strName, "(") > 0 Then strName = Trim(Mid(strName, 1, InStr(strName, "(") - 1))
                            strFilter = "系统编号='" & mlngSysNum & "' and 名称='" & strName & "'"
                            Call RecDelete(mrsProcedureFromFile, strFilter)
                            If InStr(strIniSQL, "(") - InStr(strIniSQL, strName) - Len(strName) < 5 And Mid(InStr(strIniSQL, "(") - InStr(strIniSQL, strName) - Len(strName), 1, 1) <> "-" Then
                                strFild = Mid(strIniSQL, InStr(strIniSQL, "(") + 1, InStr(strIniSQL, ")") - InStr(strIniSQL, "(") - 1)
                                If InStr(strFild, "','") > 0 Then
                                    varTemp = Split(strFild, vbCrLf)
                                Else
                                    varTemp = Split(strFild, ",")
                                End If
                                strFild = ""
                                For i = 0 To UBound(varTemp)
                                    varTemp(i) = Trim(Replace(varTemp(i), vbCrLf, ""))
                                    If varTemp(i) <> "" Then
                                        strFild = IIf(strFild = "", Trim(Mid(varTemp(i), 1, InStr(varTemp(i), " ") - 1)), strFild & "," & Trim(Mid(varTemp(i), 1, InStr(varTemp(i), " ") - 1)))
                                    End If
                                Next
                                mrsProcedureFromFile.AddNew Array("系统编号", "SQL", "名称", "字段"), Array(mlngSysNum, strSQL, strName, strFild)
                            End If
                        ElseIf strIniSQL Like "CREATE OR REPLACE VIEW *" Then
                            strName = Trim(Replace(strIniSQL, "CREATE OR REPLACE VIEW", ""))
                            strName = Mid(strName, 1, InStr(strName, " ") - 1)
                            strFilter = "系统编号='" & mlngSysNum & "' and 名称='" & strName & "'"
                            Call RecDelete(mrsViewFromFile, strFilter)
                            mrsViewFromFile.AddNew Array("系统编号", "SQL", "名称"), Array(mlngSysNum, strSQL, strName)
                        ElseIf strIniSQL Like "CREATE OR REPLACE PACKAGE*" And Not strIniSQL Like "CREATE OR REPLACE PACKAGE BODY*" Then
                            If InStr(strIniSQL, vbCrLf) > 0 Then strName = Mid(strIniSQL, 1, InStr(strIniSQL, vbCrLf) - 1)
                             strName = Trim(Replace(strName, "CREATE OR REPLACE PACKAGE", ""))
                            strName = Mid(strName, 1, InStr(strName, " ") - 1)
                            strFilter = "系统编号='" & mlngSysNum & "' and 名称='" & strName & "'"
                            Call RecDelete(mrsPackageFromFile, strFilter)
                            mrsPackageFromFile.AddNew Array("系统编号", "SQL", "名称"), Array(mlngSysNum, strSQL, strName)
                        End If
                    ElseIf strIniSQL Like "CREATE INDEX *" Or strIniSQL Like "CREATE UNIQUE INDEX*" Then
                        varTemp = Split(strIniSQL, " ON ")
                        If strIniSQL Like "CREATE INDEX *" Then
                            strName = Trim(Replace(varTemp(0), "CREATE INDEX", ""))
                        Else
                            strName = Trim(Replace(varTemp(0), "CREATE UNIQUE INDEX", ""))
                        End If
                        strTableName = Trim(Mid(varTemp(1), 1, InStr(varTemp(1), "(") - 1))
                        strFild = Replace(Mid(varTemp(1), InStr(varTemp(1), "(") + 1, InStrRev(varTemp(1), ")") - InStr(varTemp(1), "(") - 1), " ", "")
                        strTableSpace = GetTableSpace(strIniSQL)
                        strFilter = "系统编号='" & mlngSysNum & "' and 名称='" & strName & "' and 表名='" & strTableName & "'"
                        Call RecDelete(mrsIndexFromFile, strFilter)
                        mrsIndexFromFile.AddNew Array("系统编号", "SQL", "表名", "名称", "字段", "表空间"), _
                                            Array(mlngSysNum, strIniSQL, strTableName, strName, strFild, strTableSpace)
                    ElseIf strIniSQL Like "CREATE SEQUENCE*" Then
                        strName = Trim(Replace(strIniSQL, "CREATE SEQUENCE", ""))
                        strName = Mid(strName, 1, InStr(strName, " ") - 1)
                        strFilter = "系统编号='" & mlngSysNum & "' and 名称='" & strName & "'"
                        Call RecDelete(mrsSequenceFromFile, strFilter)
                        mrsSequenceFromFile.AddNew Array("系统编号", "SQL", "名称"), Array(mlngSysNum, strSQL, strName)
                    ElseIf strIniSQL Like "CREATE TABLE *" Then
                        strIniSQL = Trim(Replace(strIniSQL, "CREATE TABLE", ""))
                        strTableName = Trim(Replace(Mid(strIniSQL, 1, InStr(strIniSQL, "(") - 1), vbCrLf, ""))
                        strIniSQL = Replace(strIniSQL, vbCrLf, "||")
                        If InStr(strIniSQL, "))") > 0 Then
                            strFild = Mid(strIniSQL, InStr(strIniSQL, "(") + 1, InStr(strIniSQL, "))") - InStr(strIniSQL, "("))
                        ElseIf InStr(strIniSQL, "||)") > 0 Then
                            strFild = Mid(strIniSQL, InStr(strIniSQL, "(") + 1, InStr(strIniSQL, "||)") - InStr(strIniSQL, "("))
                        ElseIf InStr(strIniSQL, ")||") > 0 Then
                            strFild = Mid(strIniSQL, InStr(strIniSQL, "(") + 1, InStr(strIniSQL, ")||") - InStr(strIniSQL, "("))
                        Else
                            strFild = Mid(strIniSQL, InStr(strIniSQL, "(") + 1)
                        End If
                        varTemp = Split(strFild, "||")
                        For i = LBound(varTemp) To UBound(varTemp)
                            varTemp(i) = Trim(varTemp(i))
                            If varTemp(i) <> "" And varTemp(i) <> ")" And InStr(varTemp(i), "TABLESPACE") = 0 Then
                                If InStr(varTemp(i), "TABLESPACE") > 0 Then Exit For
                                strFildType = ""
                                strFildLength = ""
                                strFild = TrimEx(Mid(varTemp(i), 1, InStr(varTemp(i), " ")))
                                strIniSQL = Trim(Mid(varTemp(i), InStr(varTemp(i), " ") + 1))
                                If InStr(strIniSQL, "DATE") > 0 Then
                                    strFildType = "DATE"
                                ElseIf InStr(strIniSQL, "LONG RAW") > 0 Then
                                    strFildType = "LONG RAW"
                                Else
                                    If InStr(strIniSQL, ")") > 0 Then
                                        strIniSQL = Trim(Mid(strIniSQL, 1, InStr(strIniSQL, ")") - 1))
                                    ElseIf InStr(strIniSQL, " ") > 0 Then
                                        strIniSQL = Trim(Mid(strIniSQL, 1, InStr(strIniSQL, " ") - 1))
                                    End If
                                    If InStr(strIniSQL, "(") > 0 Then
                                        strFildType = Mid(strIniSQL, 1, InStr(strIniSQL, "(") - 1)
                                        strFildLength = Mid(strIniSQL, InStr(strIniSQL, "(") + 1)
                                    ElseIf InStr(strIniSQL, ",") > 0 Then
                                        strFildType = Mid(strIniSQL, 1, Len(strIniSQL) - 1)
                                    ElseIf InStr(strIniSQL, ")") > 0 Then
                                        strFildType = Mid(strIniSQL, 1, InStr(strIniSQL, ")") - 1)
                                    Else
                                        strFildType = strIniSQL
                                    End If
                                    strFildType = Trim(Replace(strFildType, "|", ""))
                                End If
                                If strFild <> "" Then
                                    mrsFildFromFile.AddNew Array("系统编号", "SQL", "表名", "字段", "字段类型", "字段长度"), _
                                                        Array(mlngSysNum, strSQL, strTableName, strFild, strFildType, strFildLength)
                                End If
                            End If
                        Next
                    End If
                ElseIf strIniSQL Like "ALTER*" Then
                    If strIniSQL Like "ALTER TABLE*" Then
                        If InStr(strIniSQL, "CONSTRAINT") > 0 Then
                            If InStr(strIniSQL, "_CK_") = 0 Then
                                If strIniSQL Like "ALTER TABLE * ADD CONSTRAINT *" Then
                                    varTemp = Split(strIniSQL, "ADD CONSTRAINT")
                                    strName = Trim(Replace(varTemp(1), "CONSTRAINT", ""))
                                    '获取约束名称
                                    strName = TrimEx(Mid(strName, 1, InStr(strName, " ") - 1))
                                    '获取表名
                                    strTableName = Trim(Replace(varTemp(0), "ALTER TABLE", ""))
                                    strFild = Replace(Trim(Mid(strIniSQL, InStr(strIniSQL, "(") + 1, InStr(strIniSQL, ")") - InStr(strIniSQL, "(") - 1)), " ", "")
                                    strTableSpace = GetTableSpace(strIniSQL)
                                    strFilter = "表名='" & strTableName & "' and 名称='" & strName & "' and 系统编号=" & mlngSysNum
                                    Call RecDelete(mrsConstraintFromFile, strFilter)
                                    If InStr(strSQL, "NOVALIDATE") = 0 Then strSQL = strSQL & " NOVALIDATE"
                                    mrsConstraintFromFile.AddNew Array("系统编号", "SQL", "表名", "名称", "字段", "表空间"), _
                                                        Array(mlngSysNum, strIniSQL, strTableName, strName, strFild, strTableSpace)
                                    If InStr(strIniSQL, "PRIMARY") > 0 Or InStr(strIniSQL, "UNIQUE") > 0 Then
                                        strFilter = "表名='" & strTableName & "' and 名称='" & strName & "' and 系统编号=" & mlngSysNum
                                        Call RecDelete(mrsIndexFromFile, strFilter)
                                        If strTableSpace <> "" Then
                                            strSQL = "Create Unique Index " & strName & " On " & strTableName & "(" & strFild & ") Tablespace " & strTableSpace & " Nologging"
                                            strSQL = strSQL & "||" & strIniSQL
                                        Else
                                            strSQL = "Create Unique Index " & strName & " On " & strTableName & "(" & strFild & ") Nologging"
                                            strSQL = strSQL & "||" & strIniSQL
                                        End If
                                        mrsIndexFromFile.AddNew Array("系统编号", "SQL", "表名", "名称", "字段", "表空间"), _
                                                        Array(mlngSysNum, strSQL, strTableName, strName, strFild, strTableSpace)
                                    End If
                                ElseIf strIniSQL Like "ALTER TABLE*DROP CONSTRAINT*" Then
                                    varTemp = Split(strIniSQL, "DROP CONSTRAINT")
                                    strName = Trim(Replace(varTemp(1), "CONSTRAINT", ""))
                                    If InStr(strName, " ") > 0 Then strName = TrimEx(Mid(strName, 1, InStr(strName, " ") - 1))
                                    strTableName = Trim(Replace(varTemp(0), "ALTER TABLE", ""))
                                    strFilter = "表名='" & strTableName & "' and 名称='" & strName & "' and 系统编号=" & mlngSysNum
                                    Call RecDelete(mrsConstraintFromFile, strFilter)
                                    Call RecDelete(mrsIndexFromFile, strFilter)
                                'Alter Table 病人抗生素记录 rename Constraint 病人抗生素记录_主页ID to 病人抗生素记录_FK_主页ID
                                ElseIf strIniSQL Like "ALTER TABLE*RENAME CONSTRAINT*" Then
                                    strTemp = Mid(strIniSQL, InStr(strIniSQL, "CONSTRAINT") + 11)
                                    varTemp = Split(strIniSQL, "RENAME CONSTRAINT")
                                    strTableName = Trim(Replace(varTemp(0), "ALTER TABLE", ""))
                                    If strTableName = "卡消费接口目录" Then
                                        strTableName = "消费卡类别目录"
                                    ElseIf strTableName = "消费卡目录" Then
                                        strTableName = "消费卡信息"
                                    End If
                                    varTemp = Split(varTemp(1), "TO")
                                    varTemp(0) = Trim(varTemp(0))
                                    varTemp(1) = Trim(varTemp(1))
                                    mrsConstraintFromFile.Filter = "名称='" & varTemp(0) & "' and 系统编号=" & mlngSysNum
                                    If mrsConstraintFromFile.RecordCount > 0 Then
                                        mrsConstraintFromFile!名称 = varTemp(1)
                                        mrsConstraintFromFile!表名 = strTableName
                                        mrsConstraintFromFile!SQL = Replace(mrsConstraintFromFile!SQL, varTemp(0), varTemp(1))
                                        mrsConstraintFromFile!SQL = Replace(mrsConstraintFromFile!SQL, mrsConstraintFromFile!表名, strTableName)
                                        mrsConstraintFromFile.Update
                                    End If
                                    mrsIndexFromFile.Filter = "名称='" & varTemp(0) & "' and 系统编号=" & mlngSysNum
                                    If mrsIndexFromFile.RecordCount > 0 Then
                                        mrsIndexFromFile!名称 = varTemp(1)
                                        mrsIndexFromFile!表名 = strTableName
                                        mrsIndexFromFile!SQL = "Alter Index " & mrsIndexFromFile!名称 & " rebulid nologging"
                                        mrsIndexFromFile.Update
                                    End If
                                End If
                            End If
                        Else
                            If strIniSQL Like "ALTER TABLE*ADD*" Then
                                If strIniSQL = "ALTER TABLE 时间段 ADD (" & vbNewLine & _
                                            " 站点 VARCHAR2(1)," & vbNewLine & _
                                            " 号类 VARCHAR2(10)," & vbNewLine & _
                                            " 出诊预留时间 NUMBER(18)," & vbNewLine & _
                                            " 休息时段 VARCHAR2(200))" Then
                                    strSQL = "alter table 时间段 add 站点 VARCHAR2(1)"
                                    mrsFildFromFile.AddNew Array("系统编号", "SQL", "表名", "字段", "字段类型", "字段长度"), _
                                        Array(mlngSysNum, strSQL, "时间段", "站点", "VARCHAR2", 1)
                                    strSQL = "alter table 时间段 add 号类 VARCHAR2(10)"
                                    mrsFildFromFile.AddNew Array("系统编号", "SQL", "表名", "字段", "字段类型", "字段长度"), _
                                        Array(mlngSysNum, strSQL, "时间段", "号类", "VARCHAR2", 10)
                                    strSQL = "alter table 时间段 add 出诊预留时间 NUMBER(18)"
                                    mrsFildFromFile.AddNew Array("系统编号", "SQL", "表名", "字段", "字段类型", "字段长度"), _
                                        Array(mlngSysNum, strSQL, "时间段", "出诊预留时间", "NUMBER", 18)
                                    strSQL = "alter table 时间段 add 休息时段 VARCHAR2(200)"
                                    mrsFildFromFile.AddNew Array("系统编号", "SQL", "表名", "字段", "字段类型", "字段长度"), _
                                        Array(mlngSysNum, strSQL, "时间段", "休息时段", "VARCHAR2", strFildLength)
                                ElseIf strIniSQL = "ALTER TABLE 人员收缴记录 ADD(" & vbNewLine & _
                                                    " 是否挂号 NUMBER(1)," & vbNewLine & _
                                                    " 是否就诊卡 NUMBER(1)," & vbNewLine & _
                                                    " 是否消费卡 NUMBER(1)," & vbNewLine & _
                                                    " 是否收费 NUMBER(1)," & vbNewLine & _
                                                    " 预交类别 NUMBER(2)," & vbNewLine & _
                                                    " 是否结帐 NUMBER(1))" Then
                                    strSQL = "alter table 人员收缴记录 add 是否挂号 NUMBER(1)"
                                    mrsFildFromFile.AddNew Array("系统编号", "SQL", "表名", "字段", "字段类型", "字段长度"), _
                                        Array(mlngSysNum, strSQL, "人员收缴记录", "是否挂号", "NUMBER", 1)
                                    strSQL = "alter table 人员收缴记录 add 是否就诊卡 NUMBER(1)"
                                    mrsFildFromFile.AddNew Array("系统编号", "SQL", "表名", "字段", "字段类型", "字段长度"), _
                                        Array(mlngSysNum, strSQL, "人员收缴记录", "是否就诊卡", "NUMBER", 1)
                                    strSQL = "alter table 人员收缴记录 add 是否消费卡 NUMBER(1)"
                                    mrsFildFromFile.AddNew Array("系统编号", "SQL", "表名", "字段", "字段类型", "字段长度"), _
                                        Array(mlngSysNum, strSQL, "人员收缴记录", "是否消费卡", "NUMBER", 1)
                                    strSQL = "alter table 人员收缴记录 add 是否收费 NUMBER(1)"
                                    mrsFildFromFile.AddNew Array("系统编号", "SQL", "表名", "字段", "字段类型", "字段长度"), _
                                        Array(mlngSysNum, strSQL, "人员收缴记录", "是否收费", "NUMBER", 1)
                                    strSQL = "alter table 人员收缴记录 add 预交类别 NUMBER(1)"
                                    mrsFildFromFile.AddNew Array("系统编号", "SQL", "表名", "字段", "字段类型", "字段长度"), _
                                        Array(mlngSysNum, strSQL, "人员收缴记录", "预交类别", "NUMBER", 2)
                                    strSQL = "alter table 人员收缴记录 add 是否结帐 NUMBER(1)"
                                    mrsFildFromFile.AddNew Array("系统编号", "SQL", "表名", "字段", "字段类型", "字段长度"), _
                                        Array(mlngSysNum, strSQL, "人员收缴记录", "是否结帐", "NUMBER", 1)
                                ElseIf strIniSQL = "ALTER TABLE ZLREPORTS ADD (执行人员 VARCHAR2(20), 最后执行时间 DATE)" Then
                                    strSQL = "alter table ZLREPORTS add 执行人员 varchar2(20)"
                                    mrsFildFromFile.AddNew Array("系统编号", "SQL", "表名", "字段", "字段类型", "字段长度"), _
                                    Array(mlngSysNum, strSQL, "ZLREPORTS", "执行人员", "VARCHAR2", 20)
                                    strSQL = "alter table ZLREPORTS add 最后执行时间 varchar2(20)"
                                    mrsFildFromFile.AddNew Array("系统编号", "SQL", "表名", "字段", "字段类型", "字段长度"), _
                                    Array(mlngSysNum, strSQL, "ZLREPORTS", "最后执行时间", "DATE", "")
'                                ElseIf strIniSQL = "ALTER TABLE 血液输血常规 ADD (是否婴儿 NUMBER (1)" Then
'                                    strIniSQL = "ALTER TABLE 血液输血常规 ADD 是否婴儿 NUMBER (1))"
                                    
                                Else
                                    strIniSQL = Replace(strIniSQL, vbCrLf, " ")
                                    varTemp = Split(strIniSQL, "ADD")
                                    strTableName = Trim(Replace(varTemp(0), "ALTER TABLE", ""))
                                    strName = Trim(varTemp(1))
                                    If Mid(strName, 1, 1) = "(" Then
                                        strName = Mid(strName, 2, InStrRev(strName, ")") - 2)
                                    End If
                                    strFild = Mid(strName, 1, InStr(strName, " ") - 1)
                                    If InStr(strName, "(") > 0 Then
                                        strFildType = Trim(Replace(strName, strFild, ""))
                                        strFildType = Trim(Mid(strFildType, 1, InStr(strFildType, "(") - 1))
                                        strFildLength = Trim(Mid(strName, InStr(strName, "(") + 1, InStr(strName, ")") - InStr(strName, "(") - 1))
                                    Else
                                        strFildType = Trim(Replace(strName, strFild, ""))
                                        If InStr(strFildType, " ") > 0 Then strFildType = Mid(strFildType, 1, InStr(strFildType, " ") - 1)
                                        If InStr(strFildType, ")") > 0 Then strFildType = Trim(Mid(strFildType, 1, InStr(strFildType, ")") - 1))
                                    End If
                                    strFilter = "表名='" & strTableName & "' and 字段='" & strFild & "' and 系统编号=" & mlngSysNum
                                    Call RecDelete(mrsFildFromFile, strFilter)
                                    mrsFildFromFile.AddNew Array("系统编号", "SQL", "表名", "字段", "字段类型", "字段长度"), _
                                        Array(mlngSysNum, strSQL, strTableName, strFild, strFildType, strFildLength)
                                End If
                            ElseIf strIniSQL Like "ALTER TABLE*MODIFY*" Then
                                varTemp = Split(strIniSQL, "MODIFY")
                                varTemp(1) = Trim(varTemp(1))
                                If InStr(varTemp(1), "NULL") = 0 And InStr(varTemp(1), "DEFAULT") = 0 Then
                                    strTableName = Trim(Replace(varTemp(0), "ALTER TABLE", ""))
                                    varTemp(1) = Trim(varTemp(1))
                                    If Mid(varTemp(1), 1, 1) = "(" Then varTemp(1) = Mid(varTemp(1), 2, Len(varTemp(1)) - 2)
                                    strFild = Mid(varTemp(1), 1, InStr(varTemp(1), " ") - 1)
                                    strTemp = Trim(Replace(varTemp(1), strFild, ""))
                                    If InStr(strTemp, "(") > 0 Then
                                        strFildType = Mid(strTemp, 1, InStr(strTemp, "(") - 1)
                                        strFildLength = Mid(strTemp, InStr(strTemp, "(") + 1, InStr(strTemp, ")") - InStr(strTemp, "(") - 1)
                                    Else
                                        strFildType = strTemp
                                        strFildLength = ""
                                    End If
                                    mrsFildFromFile.Filter = "表名='" & strTableName & "' and 字段='" & strFild & "'"
                                    If mrsFildFromFile.RecordCount > 0 Then
                                        mrsFildFromFile!字段 = strFild
                                        mrsFildFromFile!字段类型 = strFildType
                                        mrsFildFromFile!字段长度 = strFildLength
                                        mrsFildFromFile!SQL = strIniSQL
                                        mrsFildFromFile.Update
                                    End If
                                End If
                            ElseIf strIniSQL Like "ALTER TABLE*DROP COLUMN*" Then
                                varTemp = Split(strIniSQL, "DROP COLUMN")
                                strTableName = Trim(Replace(varTemp(0), "ALTER TABLE", ""))
                                strFild = Trim(varTemp(1))
                                strFilter = "表名='" & strTableName & "' and 字段='" & strFild & "' and 系统编号=" & mlngSysNum
                                Call RecDelete(mrsFildFromFile, strFilter)
                            ElseIf strIniSQL Like "ALTER TABLE*RENAME COLUMN*" Then
                                varTemp = Split(strIniSQL, "RENAME COLUMN")
                                strTableName = Trim(Replace(varTemp(0), "ALTER TABLE", ""))
                                varTemp = Split(varTemp(1), "TO")
                                strFild = Trim(varTemp(0))
                                strTemp = Trim(varTemp(1))
                                If strTemp Like "*BAK" Then
                                    strFilter = "表名='" & strTableName & "' and 字段='" & strFild & "'"
                                    Call RecDelete(mrsFildFromFile, strFilter)
                                Else
                                    mrsFildFromFile.Filter = "表名='" & strTableName & "' and 字段='" & strFild & "'"
                                    If mrsFildFromFile.RecordCount > 0 Then
                                        mrsFildFromFile!字段 = strTemp
                                        mrsFildFromFile.Update
                                    End If
                                End If
                            ElseIf strIniSQL Like "ALTER TABLE*RENAME TO*" Then
                                varTemp = Split(strIniSQL, "RENAME TO")
                                strTableName = Trim(Replace(varTemp(0), "ALTER TABLE", ""))
                                strTemp = Trim(varTemp(1))
                                If strTemp Like "*BAK" Then
                                    strName = Trim(Replace(strIniSQL, "DROP TABLE", ""))
                                    strFilter = "表名='" & strTableName & "' and 系统编号=" & mlngSysNum
                                    Call RecDelete(mrsFildFromFile, strFilter)
                                    strFilter = "名称 like '" & strTableName & "*' and 系统编号=" & mlngSysNum
                                    Call RecDelete(mrsConstraintFromFile, strFilter)
                                    strFilter = "名称 like '" & strTableName & "*' and 系统编号=" & mlngSysNum
                                    Call RecDelete(mrsIndexFromFile, strFilter)
                                    strFilter = "名称 like '" & strTableName & "*' and 系统编号=" & mlngSysNum
                                    Call RecDelete(mrsSequenceFromFile, strFilter)
                                    strFilter = "类别='表目录' and 对象 = '" & strTableName & "' and 系统编号=" & mlngSysNum
                                    Call RecDelete(mrsDataFromFile, strFilter)
                                Else
                                    mrsFildFromFile.Filter = "表名='" & strTableName & "'"
                                    Do While Not mrsFildFromFile.EOF
                                        mrsFildFromFile!表名 = strTemp
                                        mrsFildFromFile.Update
                                        mrsFildFromFile.MoveNext
                                    Loop
                                    mrsConstraintFromFile.Filter = "名称 like '" & strTableName & "*'"
                                    Do While Not mrsConstraintFromFile.EOF
                                        mrsConstraintFromFile!表名 = strTemp
                                        mrsConstraintFromFile!名称 = Replace(mrsConstraintFromFile!名称, strTableName, strTemp)
                                        mrsConstraintFromFile.Update
                                        mrsConstraintFromFile.MoveNext
                                    Loop
                                    mrsIndexFromFile.Filter = "名称 like '" & strTableName & "*'"
                                    Do While Not mrsIndexFromFile.EOF
                                        mrsIndexFromFile!名称 = Replace(mrsIndexFromFile!名称, strTableName, strTemp)
                                        mrsIndexFromFile.Update
                                        mrsIndexFromFile.MoveNext
                                    Loop
                                    mrsSequenceFromFile.Filter = "名称 like '" & strTableName & "*'"
                                    Do While Not mrsSequenceFromFile.EOF
                                        mrsSequenceFromFile!名称 = Replace(mrsSequenceFromFile!名称, strTableName, strTemp)
                                        mrsSequenceFromFile.Update
                                        mrsSequenceFromFile.MoveNext
                                    Loop
                                End If
                            End If
                        End If
                    ElseIf strIniSQL Like "ALTER INDEX*" Then
                        If strIniSQL Like "ALTER INDEX*RENAME TO*" Then
                            strTemp = Replace(strIniSQL, "ALTER INDEX", "")
                            varTemp = Split(strTemp, "RENAME TO")
                            varTemp(0) = Trim(varTemp(0))
                            varTemp(1) = Trim(varTemp(1))
                            strTableName = Mid(varTemp(1), 1, InStr(varTemp(1), "_") - 1)
                            mrsIndexFromFile.Filter = "名称='" & varTemp(0) & "'"
                            If mrsIndexFromFile.RecordCount > 0 Then
                                mrsIndexFromFile!名称 = varTemp(1)
                                mrsIndexFromFile!表名 = strTableName
                                mrsIndexFromFile!SQL = "Alter Index " & mrsIndexFromFile!名称 & " rebulid nologging"
                                mrsIndexFromFile.Update
                            End If
                        ElseIf strIniSQL Like "ALTER INDEX*REBUILD TABLESPACE*" Then
                            varTemp = Split(strIniSQL, "REBUILD TABLESPACE")
                            strTemp = Trim(Replace(varTemp(0), "ALTER INDEX", ""))
                            varTemp(1) = Trim(varTemp(1))
                            mrsIndexFromFile.Filter = "名称='" & strTemp & "'"
                            If mrsIndexFromFile.RecordCount > 0 Then
                                mrsIndexFromFile!表空间 = varTemp(1)
                                mrsIndexFromFile.Update
                            End If
                        End If
                    End If
                ElseIf strIniSQL Like "INSERT INTO*" Then
                    If strIniSQL Like "INSERT INTO ZLTABLES*" Then
                        strTemp = Replace(Mid(strIniSQL, InStr(strIniSQL, "(") + 1, InStr(strIniSQL, ")") - InStr(strIniSQL, "(") - 1), " ", "")
                        varTemp = Split(strTemp, ",")
                        Set rsTemp = SetSelectRecordset(strIniSQL, strTemp, varTemp, "ZLTABLES")
                        If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                        Do While Not rsTemp.EOF
                            mrsDataFromFile.Filter = "类别='表目录' and 对象='" & rsTemp!表名 & "' and 系统编号=" & mlngSysNum
                            If mrsDataFromFile.RecordCount = 0 Then
                                mrsDataFromFile.AddNew Array("类别", "SQL", "对象", "系统编号"), Array("表目录", rsTemp!修正SQL, rsTemp!表名, mlngSysNum)
                            End If
                            rsTemp.MoveNext
                        Loop
                    ElseIf strIniSQL Like "INSERT INTO ZLPROGRAMS*" Then
                        If InStr(strIniSQL, "共享号 IS NULL") > 0 And mlngShare = 0 Then
                            If strIniSQL = "INSERT INTO ZLPROGRAMS" & vbNewLine & _
                                " (序号, 标题, 说明, 系统, 部件)" & vbNewLine & _
                                " SELECT 1082 序号, '医生授权管理' 标题, '拥有本模块权限的人员可对本科室或全院临床医师的所有权限进行集中管理。' AS 说明, 0 系统, 'ZL9CISBASE' 部件" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPROGRAMS WHERE 系统 = 0 AND 序号 = 1082 AND 标题 = '医生授权管理')" Then
                                strIniSQL = "INSERT INTO ZLPROGRAMS" & vbNewLine & _
                                " (序号, 标题, 说明, 系统, 部件)" & vbNewLine & _
                                " SELECT 1082, '医生授权管理', '拥有本模块权限的人员可对本科室或全院临床医师的所有权限进行集中管理。', 0, 'ZL9CISBASE'" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPROGRAMS WHERE 系统 = 0 AND 序号 = 1082 AND 标题 = '医生授权管理')"
                            ElseIf strIniSQL = "INSERT INTO ZLPROGRAMS" & vbNewLine & _
                                " (序号, 标题, 说明, 系统, 部件)" & vbNewLine & _
                                " SELECT 2228 序号, '范文审核' 标题, '用于对范文进行审核操作' AS 说明, 0 系统, 'ZL9EMRINTERFACE' 部件" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPROGRAMS WHERE 系统 = 0 AND 序号 = 2228 AND 标题='范文审核')" Then
                                strIniSQL = "INSERT INTO ZLPROGRAMS" & vbNewLine & _
                                " (序号, 标题, 说明, 系统, 部件)" & vbNewLine & _
                                " SELECT 2228,'范文审核','用于对范文进行审核操作',0,'ZL9EMRINTERFACE'" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPROGRAMS WHERE 系统 = 0 AND 序号 = 2228 AND 标题='范文审核')"
                            ElseIf strIniSQL = "INSERT INTO ZLPROGRAMS" & vbNewLine & _
                                " (序号, 标题, 说明, 系统, 部件)" & vbNewLine & _
                                " SELECT 2227 序号, '取消完成审批' 标题, '用于在病历完成后需要再次修改时进行审批操作' AS 说明, 0 系统, 'ZL9EMRINTERFACE' 部件" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPROGRAMS WHERE 系统 = 0 AND 序号 = 2227 AND 标题='取消完成审批')" Then
                                strIniSQL = "INSERT INTO ZLPROGRAMS" & vbNewLine & _
                                " (序号, 标题, 说明, 系统, 部件)" & vbNewLine & _
                                " SELECT 2227,'取消完成审批','用于在病历完成后需要再次修改时进行审批操作',0,'ZL9EMRINTERFACE'" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPROGRAMS WHERE 系统 = 0 AND 序号 = 2227 AND 标题='取消完成审批')"
                            ElseIf strIniSQL = "INSERT INTO ZLPROGRAMS" & vbNewLine & _
                                " (序号, 标题, 说明, 系统, 部件)" & vbNewLine & _
                                " SELECT 2226 序号, '终末质控接收' 标题, '用于终末质控前进行接收工作及工作量统计' AS 说明, 0 系统, 'ZL9EMRINTERFACE' 部件" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPROGRAMS WHERE 序号 = 2226 AND 标题='终末质控接收')" Then
                                strIniSQL = "INSERT INTO ZLPROGRAMS" & vbNewLine & _
                                " (序号, 标题, 说明, 系统, 部件)" & vbNewLine & _
                                " SELECT 2226, '终末质控接收', '用于终末质控前进行接收工作及工作量统计', 0, 'ZL9EMRINTERFACE'" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPROGRAMS WHERE 序号 = 2226 AND 标题='终末质控接收')"
                            ElseIf strIniSQL = "INSERT INTO ZLPROGRAMS" & vbNewLine & _
                                " (序号, 标题, 说明, 系统, 部件)" & vbNewLine & _
                                " SELECT 2228 序号, '范文审核' 标题, '用于对范文进行审核操作' AS 说明, 0 系统, 'ZL9EMRINTERFACE' 部件" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPROGRAMS WHERE 序号 = 2228 AND 标题='范文审核')" Then
                                strIniSQL = "INSERT INTO ZLPROGRAMS" & vbNewLine & _
                                " (序号, 标题, 说明, 系统, 部件)" & vbNewLine & _
                                " SELECT 2228, '范文审核', '用于对范文进行审核操作', 0, 'ZL9EMRINTERFACE'" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPROGRAMS WHERE 序号 = 2228 AND 标题='范文审核')"
                            ElseIf strIniSQL = "INSERT INTO ZLPROGRAMS" & vbNewLine & _
                                " (序号, 标题, 说明, 系统, 部件)" & vbNewLine & _
                                " SELECT 2227 序号, '取消完成审批' 标题, '用于在病历完成后需要再次修改时进行审批操作' AS 说明, 0 系统, 'ZL9EMRINTERFACE' 部件" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPROGRAMS WHERE 序号 = 2227 AND 标题='取消完成审批')" Then
                                strIniSQL = "INSERT INTO ZLPROGRAMS" & vbNewLine & _
                                " (序号, 标题, 说明, 系统, 部件)" & vbNewLine & _
                                " SELECT 2227, '取消完成审批', '用于在病历完成后需要再次修改时进行审批操作', 0, 'ZL9EMRINTERFACE'" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPROGRAMS WHERE 序号 = 2227 AND 标题='取消完成审批')"
                                
                            End If
                            strTemp = Replace(Mid(strIniSQL, InStr(strIniSQL, "(") + 1, InStr(strIniSQL, ")") - InStr(strIniSQL, "(") - 1), " ", "")
                            varTemp = Split(strTemp, ",")
                            Set rsTemp = SetSelectRecordset(strIniSQL, strTemp, varTemp, "ZLPROGRAMS")
                            If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                            Do While Not rsTemp.EOF
                                mrsDataFromFile.AddNew Array("类别", "SQL", "系统编号", "序号", "对象"), Array("模块", rsTemp!修正SQL, mlngSysNum, Trim(rsTemp!序号), Trim(rsTemp!标题))
                                rsTemp.MoveNext
                            Loop
                        ElseIf strIniSQL Like "INSERT INTO ZLPROGFUNCS*" Then
                            If InStr(strIniSQL, "共享号 IS NULL") > 0 And mlngShare = 0 Then
                                strTemp = Replace(Mid(strIniSQL, InStr(strIniSQL, "(") + 1, InStr(strIniSQL, ")") - InStr(strIniSQL, "(") - 1), " ", "")
                                varTemp = Split(strTemp, ",")
                                Set rsTemp = SetSelectRecordset(strIniSQL, strTemp, varTemp, "ZLPROGFUNCS")
                                If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                                Do While Not rsTemp.EOF
                                    mrsDataFromFile.AddNew Array("类别", "SQL", "系统编号", "序号", "对象"), Array("功能", rsTemp!修正SQL, mlngSysNum, rsTemp!序号, rsTemp!功能)
                                    rsTemp.MoveNext
                                Loop
                            End If
                        ElseIf strIniSQL Like "INSERT INTO ZLPARAMETERS*" Then
                            If strIniSQL = "INSERT INTO ZLPARAMETERS" & vbNewLine & _
                                " (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)" & vbNewLine & _
                                " SELECT ZLPARAMETERS_ID.NEXTVAL, 0, 1124, 0, 0, 0, 0, 16, '补结算有效天数'," & vbNewLine & _
                                " (SELECT DECODE(SUBSTR(NVL(参数值, 缺省值), 1, 1), '0', '3', SUBSTR(NVL(参数值, 缺省值), 1, 1)) AS VALIDDAY" & vbNewLine & _
                                " FROM ZLPARAMETERS" & vbNewLine & _
                                " WHERE 系统 = 0 AND 模块 IS NULL AND 参数名 = '挂号有效天数'), '3', '可进行医保补充结算的费用有效天数。'" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPARAMETERS WHERE 系统 = 0 AND 模块 = 1124 AND 参数名 = '补结算有效天数')" Then
                                strIniSQL = "INSERT INTO ZLPARAMETERS" & vbNewLine & _
                                " (ID, 系统, 模块, 私有, 本机, 授权, 固定, 参数号, 参数名, 参数值, 缺省值, 参数说明)" & vbNewLine & _
                                " SELECT ZLPARAMETERS_ID.NEXTVAL, 0, 1124, 0, 0, 0, 0, 16, '补结算有效天数','0', '3', '可进行医保补充结算的费用有效天数。'" & vbNewLine & _
                                " FROM DUAL" & vbNewLine & _
                                " WHERE NOT EXISTS (SELECT 1 FROM ZLPARAMETERS WHERE 系统 = 0 AND 模块 = 1124 AND 参数名 = '补结算有效天数')"
                            End If
                            strTemp = Replace(Mid(strIniSQL, InStr(strIniSQL, "(") + 1, InStr(strIniSQL, ")") - InStr(strIniSQL, "(") - 1), " ", "")
                            varTemp = Split(strTemp, ",")
                            Set rsTemp = SetSelectRecordset(strIniSQL, strTemp, varTemp, "ZLPARAMETERS")
                            If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                            Do While Not rsTemp.EOF
                                If IsNull(rsTemp!系统) Or InStr(rsTemp!系统, "NULL") > 0 Or rsTemp!模块 = """" Then
                                    lngSys = 0
                                Else
                                    lngSys = mlngSysNum
                                End If
                                If InStr(strTemp, "模块") > 0 Then
                                    If InStr(UCase(rsTemp!模块), "NULL") > 0 Or rsTemp!模块 = """" Then
                                        mrsDataFromFile.AddNew Array("类别", "SQL", "系统编号", "对象", "参数号", "参数名"), Array("参数", rsTemp!修正SQL, lngSys, "NULL", rsTemp!参数号, rsTemp!参数名)
                                    Else
                                        mrsDataFromFile.AddNew Array("类别", "SQL", "系统编号", "对象", "参数号", "参数名"), Array("参数", rsTemp!修正SQL, lngSys, rsTemp!模块, rsTemp!参数号, rsTemp!参数名)
                                    End If
                                Else
                                    mrsDataFromFile.AddNew Array("类别", "SQL", "系统编号", "对象", "参数号", "参数名"), Array("参数", rsTemp!修正SQL, lngSys, "NULL", rsTemp!参数号, rsTemp!参数名)
                                End If
                                rsTemp.MoveNext
                            Loop
                        ElseIf strIniSQL Like "INSERT INTO ZLREPORTS*" Then
                            strTemp = Replace(Mid(strIniSQL, InStr(strIniSQL, "(") + 1, InStr(strIniSQL, ")") - InStr(strIniSQL, "(") - 1), " ", "")
                            varTemp = Split(Replace(Mid(strIniSQL, InStr(strIniSQL, "(") + 1, InStr(strIniSQL, ")") - InStr(strIniSQL, "(") - 1), " ", ""), ",")
                            Set rsTemp = SetSelectRecordset(strIniSQL, strTemp, varTemp, "ZLREPORTS")
                            If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                            Do While Not rsTemp.EOF
                                mrsDataFromFile.AddNew Array("类别", "SQL", "系统编号", "对象", "名称"), Array("报表", rsTemp!修正SQL, mlngSysNum, rsTemp!编号, rsTemp!名称)
                                rsTemp.MoveNext
                            Loop
                        End If
                    End If
                ElseIf strIniSQL Like "UPDATE ZLPARAMETERS*" Then
                    If strIniSQL = "UPDATE ZLPARAMETERS" & vbNewLine & _
                            "SET 参数名 = '一卡通消费刷卡控制', 参数值 = DECODE(参数值, NULL, NULL, '1|' || 参数值), 缺省值 = '1|0'," & vbNewLine & _
                            " 影响控制说明 = '在以下操作前，是否需要病人刷卡进行密码验证：' || CHR(10) ||" & vbNewLine & _
                            " ' 1)门诊记帐，销帐，记帐划价审核' || CHR(10) ||" & vbNewLine & _
                            " ' 2)门诊结帐使用预存款或门诊结帐作废退预存款' || CHR(10) ||" & vbNewLine & _
                            " ' 3)门诊收费使用预存款或门诊退费退回预存款' || CHR(10) ||" & vbNewLine & _
                            " ' 4)门诊挂号使用预存款或退号退回预存款' || CHR(10) ||" & vbNewLine & _
                            " ' 5)门诊医嘱发送为记帐单，住院医嘱发送为门诊记帐单；针对门诊记帐费用：住院护士站执行完成，医技工作站执行完成或批量执行完成'," & vbNewLine & _
                            " 参数值含义 = '参数格式:消费刷卡控制|退费刷卡控制' || CHR(10) ||" & vbNewLine & _
                            " ' 1.消费刷卡控制:0-不进行刷卡控制，1-门诊消费时需要刷卡验证，2-门诊消费时设置密码的(只要存在一张卡有密码的，就代表设置了密码的)，则必须刷卡验证。' || CHR(10) ||" & vbNewLine & _
                            " ' 2.退费刷卡控制:0-不进行刷卡控制，1-门诊退费时需要刷卡验证，2-门诊退费时设置密码的(只要存在一张卡有密码的，就代表设置了密码的)，则必须刷卡验证。'," & vbNewLine & _
                            " 关联说明 = '参数""门诊发送为划价单诊疗类别""，如果是发送为划价单，并且不是执行后本科自动审核的情况，则不会弹出密码验证，因为还没有实际扣减病人的费用'," & vbNewLine & _
                            " 适用说明 = '如果病人刷卡消费无须密码验证，则可能存在卡被盗刷的安全风险。有的医院为了方便病人就诊，接受这种风险，在发卡时，医院与病人一般需要签定协议 '," & vbNewLine & _
                            " 警告说明 = '此参数建议不调整为""不进行刷卡控制""，这样可能会存在病人资金安全隐患，为了避免隐患，请要求每个病人都设置刷卡消费密码，以保证资金安全'" & vbNewLine & _
                            "WHERE 系统 = 0 AND 模块 IS NULL AND 参数号 = 28" Then
                        strIniSQL = "UPDATE ZLPARAMETERS" & vbNewLine & _
                            "SET 参数名 = '一卡通消费刷卡控制', 参数值 = """", 缺省值 = '1|0'," & vbNewLine & _
                            " 影响控制说明 = '在以下操作前，是否需要病人刷卡进行密码验证：' || CHR(10) ||" & vbNewLine & _
                            " ' 1)门诊记帐，销帐，记帐划价审核' || CHR(10) ||" & vbNewLine & _
                            " ' 2)门诊结帐使用预存款或门诊结帐作废退预存款' || CHR(10) ||" & vbNewLine & _
                            " ' 3)门诊收费使用预存款或门诊退费退回预存款' || CHR(10) ||" & vbNewLine & _
                            " ' 4)门诊挂号使用预存款或退号退回预存款' || CHR(10) ||" & vbNewLine & _
                            " ' 5)门诊医嘱发送为记帐单，住院医嘱发送为门诊记帐单；针对门诊记帐费用：住院护士站执行完成，医技工作站执行完成或批量执行完成'," & vbNewLine & _
                            " 参数值含义 = '参数格式:消费刷卡控制|退费刷卡控制' || CHR(10) ||" & vbNewLine & _
                            " ' 1.消费刷卡控制:0-不进行刷卡控制，1-门诊消费时需要刷卡验证，2-门诊消费时设置密码的(只要存在一张卡有密码的，就代表设置了密码的)，则必须刷卡验证。' || CHR(10) ||" & vbNewLine & _
                            " ' 2.退费刷卡控制:0-不进行刷卡控制，1-门诊退费时需要刷卡验证，2-门诊退费时设置密码的(只要存在一张卡有密码的，就代表设置了密码的)，则必须刷卡验证。'," & vbNewLine & _
                            " 关联说明 = '参数""门诊发送为划价单诊疗类别""，如果是发送为划价单，并且不是执行后本科自动审核的情况，则不会弹出密码验证，因为还没有实际扣减病人的费用'," & vbNewLine & _
                            " 适用说明 = '如果病人刷卡消费无须密码验证，则可能存在卡被盗刷的安全风险。有的医院为了方便病人就诊，接受这种风险，在发卡时，医院与病人一般需要签定协议 '," & vbNewLine & _
                            " 警告说明 = '此参数建议不调整为""不进行刷卡控制""，这样可能会存在病人资金安全隐患，为了避免隐患，请要求每个病人都设置刷卡消费密码，以保证资金安全'" & vbNewLine & _
                            "WHERE 系统 = 0 AND 模块 IS NULL AND 参数号 = 28"
                    End If
                    strTemp = Mid(strIniSQL, InStr(strIniSQL, "SET") + 3, InStr(strIniSQL, "WHERE") - InStr(strIniSQL, "SET") - 3)
                    If (InStr(strTemp, "参数名") > 0 Or InStr(strTemp, "参数号") > 0 Or InStr(strTemp, "模块") > 0) Then
                        If InStr(strIniSQL, "NOT EXISTS") > 0 Then
                            strTemp = Mid(strIniSQL, 1, InStr(strIniSQL, "NOT EXISTS") - 1)
                        Else
                            strTemp = strIniSQL
                        End If
                        strTemp = Mid(strTemp, InStr(strTemp, "WHERE") + 6)
                        varTemp = Split(strTemp, "AND")
                        strFild = ""
                        strReFild = ""
                        strName = ""
                        For i = 0 To UBound(varTemp)
                            varTemp(i) = Trim(varTemp(i))
                            If varTemp(i) <> "" Then
                                If InStr(varTemp(i), "IS NULL") = 0 Then
                                    strTemp = Trim(Mid(varTemp(i), 1, InStr(varTemp(i), "=") - 1))
                                    If InStr(strTemp, "参数名") > 0 Then
                                        strFild = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                    ElseIf InStr(strTemp, "参数号") > 0 Then
                                        strReFild = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                    ElseIf InStr(strTemp, "模块") > 0 Then
                                        strName = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                    End If
                                Else
                                    strTemp = Mid(varTemp(i), 1, InStr(varTemp(i), " ") - 1)
                                    If InStr(strTemp, "参数名") > 0 Then
                                        strFild = "NULL"
                                    ElseIf InStr(strTemp, "参数号") > 0 Then
                                        strReFild = "NULL"
                                    ElseIf InStr(strTemp, "模块") > 0 Then
                                        strName = "NULL"
                                    End If
                                End If
                            End If
                        Next
                        If strFild <> "" And strReFild = "" Then
                            mrsDataFromFile.Filter = "系统编号=" & mlngSysNum & " and 对象=" & strName & " and 参数名='" & strFild & "'"
                        ElseIf strFild = "" And strReFild <> "" Then
                            mrsDataFromFile.Filter = "系统编号=" & mlngSysNum & " and 对象=" & strName & " and 参数号='" & strReFild & "'"
                        ElseIf strFild <> "" And strReFild <> "" Then
                            mrsDataFromFile.Filter = "系统编号=" & mlngSysNum & " and 对象=" & strName & " and 参数号='" & strReFild & "' and 参数名='" & strFild & "'"
                        End If
                        If mrsDataFromFile.RecordCount > 0 Then
                            strFild = ""
                            strReFild = ""
                            strName = ""
                            strTemp = Replace(Mid(strIniSQL, InStr(strIniSQL, "SET") + 3, InStr(strIniSQL, "WHERE") - InStr(strIniSQL, "SET") - 3), vbCrLf, " ")
                            varTemp = Split(strTemp, ",")
                            For i = 0 To UBound(varTemp)
                                strTemp = Trim(Mid(varTemp(i), 1, InStr(varTemp(i), "=") - 1))
                                If strTemp = "参数名" Then
                                    strFild = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                ElseIf strTemp = "参数号" Then
                                    strReFild = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                ElseIf strTemp = "模块" Then
                                    strName = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                End If
                            Next
                            If strFild <> "" Then mrsDataFromFile!参数名 = strFild
                            If strReFild <> "" Then mrsDataFromFile!参数号 = strReFild
                            If strName <> "" Then mrsDataFromFile!对象 = strName
                            mrsDataFromFile.Update
                        End If
                    End If
                ElseIf strIniSQL Like "DELETE*" Then
                    If strIniSQL Like "DELETE ZLPARAMETERS*" Or strIniSQL Like "DELETE FROM ZLPARAMETERS*" Then
                        strTemp = Mid(strIniSQL, InStr(strIniSQL, "WHERE") + 6)
                        If InStr(strTemp, "IN") = 0 And InStr(strTemp, "OR") = 0 Then
                            strName = ""
                            strFild = ""
                            strReFild = ""
                            varTemp = Split(strTemp, "AND")
                            For i = 0 To UBound(varTemp)
                                If InStr(varTemp(i), "NULL") > 0 Then
                                    If InStr(varTemp(i), "模块") > 0 Then
                                        strName = "NULL"
                                    End If
                                Else
                                    strTemp = Trim(Mid(varTemp(i), 1, InStr(varTemp(i), "=") - 1))
                                    If InStr(strTemp, "NVL") > 0 Then
                                        strTemp = Trim(Mid(strTemp, InStr(strTemp, "(") + 1, InStr(strTemp, ",") - InStr(strTemp, "(") - 1))
                                    End If
                                    If strTemp = "模块" Then
                                        strName = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                    ElseIf strTemp = "参数名" Then
                                        strFild = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                    ElseIf strTemp = "参数号" Then
                                        strReFild = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                    End If
                                End If
                            Next
                            If strFild <> "" And strReFild <> "" Then
                                strFilter = "类别='参数' and 系统编号=" & mlngSysNum & " and 对象=" & strName & " and 参数名='" & strFild & "' and 参数号='" & strReFild & "'"
                            ElseIf strFild = "" Then
                                strFilter = "类别='参数' and 系统编号=" & mlngSysNum & " and 对象=" & strName & " and 参数号='" & strReFild & "'"
                            ElseIf strReFild = "" Then
                                strFilter = "类别='参数' and 系统编号=" & mlngSysNum & " and 对象=" & strName & " and 参数名='" & strFild & "'"
                            End If
                            If strName = "NULL" Then
                                strFilter = Replace(strFilter, "=NULL", "='NULL'")
                            End If
                            Call RecDelete(mrsDataFromFile, strFilter)
                        ElseIf InStr(strTemp, "IN") = 0 And InStr(strTemp, "OR") > 0 Then
                            varTemp = Split(strTemp, "AND")
                            For i = 0 To UBound(varTemp)
                                If InStr(varTemp(i), "(") = 0 Then
                                    strTemp = Trim(Mid(varTemp(i), 1, InStr(varTemp(i), "=") - 1))
                                    If strTemp = "参数名" Then
                                        strFild = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                    End If
                                Else
                                    strName = Mid(varTemp(i), InStr(varTemp(i), "(") + 1, InStr(varTemp(i), ")") - InStr(varTemp(i), "(") - 1)
                                End If
                            Next
                            If strName = "模块 = 1291 OR 模块 = 1294" Then
                                strFilter = "类别='参数' and 系统编号=" & mlngSysNum & " and 对象=1291 and 参数名='" & strFild & "'"
                                Call RecDelete(mrsDataFromFile, strFilter)
                                strFilter = "类别='参数' and 系统编号=" & mlngSysNum & " and 对象=1294 and 参数名='" & strFild & "'"
                                Call RecDelete(mrsDataFromFile, strFilter)
                            End If
                        ElseIf strIniSQL = "DELETE FROM ZLPARAMETERS WHERE 系统=&N_SYSTEM AND NVL(模块,0) IN (1252,1253) AND 参数名='自动处理皮试'" Then
                            strFilter = "类别='参数' and 系统编号=" & mlngSysNum & " and 对象=1252 and 参数名='自动处理皮试'"
                            Call RecDelete(mrsDataFromFile, strFilter)
                            strFilter = "类别='参数' and 系统编号=" & mlngSysNum & " and 对象=1253 and 参数名='自动处理皮试'"
                            Call RecDelete(mrsDataFromFile, strFilter)
                        ElseIf strIniSQL = "DELETE ZLPARAMETERS WHERE 系统 = &N_SYSTEM AND (模块 = 1252 AND 参数号 IN (22, 24) OR 模块 = 1253 AND 参数号 IN (17, 19, 45))" Then
                            strFilter = "类别='参数' and 系统编号=" & mlngSysNum & " and 对象=1252 and 参数号=22"
                            Call RecDelete(mrsDataFromFile, strFilter)
                            strFilter = "类别='参数' and 系统编号=" & mlngSysNum & " and 对象=1252 and 参数号=24"
                            Call RecDelete(mrsDataFromFile, strFilter)
                            strFilter = "类别='参数' and 系统编号=" & mlngSysNum & " and 对象=1253 and 参数号=17"
                            Call RecDelete(mrsDataFromFile, strFilter)
                            strFilter = "类别='参数' and 系统编号=" & mlngSysNum & " and 对象=1253 and 参数号=19"
                            Call RecDelete(mrsDataFromFile, strFilter)
                            strFilter = "类别='参数' and 系统编号=" & mlngSysNum & " and 对象=1253 and 参数号=45"
                            Call RecDelete(mrsDataFromFile, strFilter)
                        End If
                    ElseIf strIniSQL Like "DELETE ZLPROGFUNCS*" Or strIniSQL Like "DELETE FROM ZLPROGFUNCS*" Then
                        strTemp = Mid(strIniSQL, InStr(strIniSQL, "WHERE") + 6)
                        If InStr(strTemp, "OR") = 0 Then
                            varTemp = Split(strTemp, "AND")
                            For i = 0 To UBound(varTemp)
                                strTemp = Trim(Mid(varTemp(i), 1, InStr(varTemp(i), "=") - 1))
                                If strTemp = "序号" Then
                                    strFild = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                ElseIf strTemp = "功能" Then
                                    strReFild = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                End If
                            Next
                            strFilter = "类别='功能' and 序号=" & strFild & " and 对象='" & strReFild & "' and 系统编号=" & mlngSysNum
                            Call RecDelete(mrsDataFromFile, strFilter)
                        Else
                            varTemp = Split(strTemp, "OR")
                            For i = 0 To UBound(varTemp)
                                strFild = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                strFilter = "类别='功能' and 对象='" & strFild & "' and 系统编号=" & mlngSysNum
                                Call RecDelete(mrsDataFromFile, strFilter)
                            Next
                        End If
                    End If
                ElseIf strIniSQL Like "DROP*" Then
                    If strIniSQL Like "DROP SEQUENCE*" Then
                        strName = Trim(Replace(strIniSQL, "DROP SEQUENCE", ""))
                        strFilter = "名称='" & strName & "' and 系统编号=" & mlngSysNum
                        Call RecDelete(mrsSequenceFromFile, strFilter)
                    ElseIf strIniSQL Like "DROP INDEX*" Then
                        strName = Trim(Replace(strIniSQL, "DROP INDEX", ""))
                        strFilter = "名称='" & strName & "' and 系统编号=" & mlngSysNum
                        Call RecDelete(mrsIndexFromFile, strFilter)
                    ElseIf strIniSQL Like "DROP TABLE*" Then
                        strName = Trim(Replace(strIniSQL, "DROP TABLE", ""))
                        strFilter = "表名='" & strName & "' and 系统编号=" & mlngSysNum
                        Call RecDelete(mrsFildFromFile, strFilter)
                        strFilter = "名称 like '" & strName & "*' and 系统编号=" & mlngSysNum
                        Call RecDelete(mrsConstraintFromFile, strFilter)
                        strFilter = "名称 like '" & strName & "*' and 系统编号=" & mlngSysNum
                        Call RecDelete(mrsIndexFromFile, strFilter)
                        strFilter = "名称 like '" & strName & "*' and 系统编号=" & mlngSysNum
                        Call RecDelete(mrsSequenceFromFile, strFilter)
                        strFilter = "类别='表目录' and 对象 = '" & strName & "' and 系统编号=" & mlngSysNum
                        Call RecDelete(mrsDataFromFile, strFilter)
                    ElseIf strIniSQL Like "DROP PROCEDURE*" Then
                        strName = Trim(Replace(strIniSQL, "DROP PROCEDURE", ""))
                        strFilter = "名称='" & strName & "' and 系统编号=" & mlngSysNum
                        Call RecDelete(mrsProcedureFromFile, strFilter)
                    ElseIf strIniSQL Like "DROP FUNCTION*" Then
                        strName = Trim(Replace(strIniSQL, "DROP FUNCTION", ""))
                        strFilter = "名称='" & strName & "' and 系统编号=" & mlngSysNum
                        Call RecDelete(mrsProcedureFromFile, strFilter)
                    ElseIf strIniSQL Like "DROP VIEW*" Then
                        strName = Trim(Replace(strIniSQL, "DROP VIEW", ""))
                        strFilter = "名称='" & strName & "' and 系统编号=" & mlngSysNum
                        Call RecDelete(mrsViewFromFile, strFilter)
                    End If
                ElseIf strIniSQL Like "UPDATE ZLTABLES*" Then
                    strTemp = Mid(strIniSQL, InStr(strIniSQL, "SET") + 3, InStr(strIniSQL, "WHERE") - InStr(strIniSQL, "SET") - 3)
                    If InStr(strTemp, "表名") > 0 Then
                        strTemp = Mid(strIniSQL, InStr(strIniSQL, "WHERE") + 6)
                        varTemp = Split(strTemp, "AND")
                        strTemp = ""
                        strFild = ""
                        strReFild = ""
                        For i = 0 To UBound(varTemp)
                            strTemp = Trim(Mid(varTemp(i), 1, InStr(varTemp(i), "=") - 1))
                            If strTemp <> "" And strTemp = "表名" Then
                                strTableName = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                            End If
                        Next
                        strFilter = "系统编号=" & mlngSysNum
                        If strTableName <> "" Then strFilter = strFilter & " and 对象='" & strTableName & "'"
                        mrsDataFromFile.Filter = strFilter
                        If mrsDataFromFile.RecordCount > 0 Then
                            strTemp = Mid(strIniSQL, InStr(strIniSQL, "SET") + 3, InStr(strIniSQL, "WHERE") - InStr(strIniSQL, "SET") - 3)
                            varTemp = Split(strTemp, ",")
                            strTemp = ""
                            strFild = ""
                            strReFild = ""
                            For i = 0 To UBound(varTemp)
                                strTemp = Trim(Mid(varTemp(i), 1, InStr(varTemp(i), "=") - 1))
                                If strTemp = "表名" Then
                                    strTableName = Trim(Replace(Mid(varTemp(i), InStr(varTemp(i), "=") + 1), "'", ""))
                                End If
                            Next
                            If strTableName <> "" Then
                                mrsDataFromFile!对象 = strTableName
                                mrsDataFromFile.Update
                            End If
                        End If
                    End If
                End If
            End If
            DoEvents
            Call mclsRunScript.ReadNextSQL
        Loop
    End If
End Sub

Private Function iniSQL(ByVal strSQL As String) As String
    
    strSQL = Trim(UCase(strSQL))
    strSQL = Replace(strSQL, "ZLTOOLS.", "")
    strSQL = Replace(strSQL, Chr(0), " ")
    strSQL = Replace(strSQL, vbTab, " ")
    Do While InStr(strSQL, "  ") > 0
        strSQL = Replace(strSQL, "  ", " ")
    Loop
    If Mid(strSQL, 1, 11) = "INSERT INTO" Or Mid(strSQL, 1, 6) = "UPDATE" Then
        Call ReplaceMark(strSQL, strSQL)
    Else
        If strSQL Like "CREATE TABLE*" Or strSQL Like "CREATE OR REPLACE PROCEDURE*" Or strSQL Like "CREATE OR REPLACE FUNCTION*" Then
            strSQL = TrimAllComment(strSQL)
        End If
    End If
    
    If Right(strSQL, 1) = ";" Then
        strSQL = Mid(strSQL, 1, Len(strSQL) - 1)
    End If
    iniSQL = strSQL
End Function

Private Function ReplaceMark(ByRef strSQL As String, ByVal strCut As String) As String
    Dim strTemp As String
    Dim strCutSQL As String
    Dim strReplaceCutSQL As String
    Dim strIniSQL As String
    Dim lngBegin As Long
    
    If InStr(strCut, "'") > 0 Then
        lngBegin = InStr(strCut, "'") + 1
        strTemp = Mid(strCut, lngBegin)
        strCutSQL = "'" & Mid(strTemp, 1, InStr(strTemp, "'"))
        If InStr(strCutSQL, ",") > 0 Then
            If strCutSQL = "','" Then
                If InStr(Mid(strTemp, 1, 6), "||") > 0 Then
                    strCutSQL = "'" & Mid(strTemp, 1, 6)
                    strReplaceCutSQL = Replace(strCutSQL, ",", "，")
                    strSQL = Replace(strSQL, strCutSQL, strReplaceCutSQL)
                End If
            Else
                strReplaceCutSQL = Replace(strCutSQL, ",", "，")
                strSQL = Replace(strSQL, strCutSQL, strReplaceCutSQL)
            End If
        End If
        lngBegin = InStr(strTemp, "'") + 1
        strTemp = Mid(strTemp, lngBegin)
        Call ReplaceMark(strSQL, strTemp)
    End If
End Function

Public Function TrimAllComment(ByVal strSQL As String) As String
'功能：去掉写在单行strSQL语句后面的"--"或者"/"注释(只针对本模块，主要是用于去掉表或者过程/函数字段或参数后面的注释)
'说明：主要是RunSQLFile的子函数
    Dim strTemp As String
    Dim strModifySQL As String
    Dim varTemp As Variant
    Dim blnStr As Boolean
    Dim i As Long
    
    If Mid(strSQL, 1, 2) = "--" Or strSQL = "" Or Mid(strSQL, 1, 1) = "/" Then Exit Function
    varTemp = Split(strSQL, vbCrLf)
    For i = 0 To UBound(varTemp)
        If InStr(varTemp(i), "--") > 0 Then
            strTemp = Mid(varTemp(i), 1, InStr(varTemp(i), "--") - 1)
            strModifySQL = IIf(strModifySQL = "", strTemp, strModifySQL & vbCrLf & strTemp)
        ElseIf InStr(varTemp(i), "/") > 0 Then
            strTemp = Mid(varTemp(i), 1, InStr(varTemp(i), "/") - 1)
            strModifySQL = IIf(strModifySQL = "", strTemp, strModifySQL & vbCrLf & strTemp)
        Else
            strModifySQL = IIf(strModifySQL = "", varTemp(i), strModifySQL & vbCrLf & varTemp(i))
        End If
    Next
    TrimAllComment = strModifySQL
End Function

Public Function GetTableSpace(ByVal strSQL As String) As String
'功能：如果一条语句里有表空间，则返回表空间名；若无，则返回空
    Dim strTemp As String
    
    If InStr(strSQL, "TABLESPACE") > 0 Then
        strSQL = Replace(strSQL, vbCrLf, " ")
        strTemp = Trim(Right(strSQL, Len(strSQL) - InStrRev(strSQL, "TABLESPACE") - 10))
        If InStr(strTemp, " ") > 0 Then
            GetTableSpace = Mid(strTemp, 1, InStr(strTemp, " ") - 1)
        ElseIf InStr(strTemp, ";") > 0 Then
            GetTableSpace = Mid(strTemp, 1, InStr(strTemp, ";") - 1)
        Else
            GetTableSpace = strTemp
        End If
    Else
        GetTableSpace = ""
    End If
End Function

Private Function CheckSetFile(ByVal strPath As String, ByVal strSysNum As Long) As Boolean
'功能：检查本地安装脚本
'参数：strPath-主路径，strSysNum-系统编号
    Dim strMainPath As String
    Dim strProblem As String
    
    If strSysNum = 0 Then
        Call AddFilePath(UCase(strPath), "ZLSERVER.SQL", "管理工具文件", strSysNum, strProblem)
        CheckSetFile = True
        Exit Function
    End If
    
    strMainPath = UCase(Mid(strPath, 1, InStrRev(strPath, "\")))
    
    '检查安装脚本是否存在
    Call AddFilePath(strMainPath & "ZLSEQUENCE.SQL", "ZLSEQUENCE.SQL", "序列文件", strSysNum, strProblem)
    Call AddFilePath(strMainPath & "ZLTABLE.SQL", "ZLTABLE.SQL", "数据表文件", strSysNum, strProblem)
    Call AddFilePath(strMainPath & "ZLCONSTRAINT.SQL", "ZLCONSTRAINT.SQL", "约束文件", strSysNum, strProblem)
    Call AddFilePath(strMainPath & "ZLINDEX.SQL", "ZLINDEX.SQL", "索引文件", strSysNum, strProblem)
    Call AddFilePath(strMainPath & "ZLVIEW.SQL", "ZLVIEW.SQL", "视图文件", strSysNum, strProblem)
    Call AddFilePath(strMainPath & "ZLPROGRAM.SQL", "ZLPROGRAM.SQL", "函数过程文件", strSysNum, strProblem)
    Call AddFilePath(strMainPath & "ZLMANDATA.SQL", "ZLMANDATA.SQL", "管理数据文件", strSysNum, strProblem)
    Call AddFilePath(strMainPath & "ZLREPORT.SQL", "ZLREPORT.SQL", "报表数据文件", strSysNum, strProblem)
    
    If strProblem <> "" Then
        MsgBox "以下服务器安装的相关文件丢失，不能继续，包括：" & strProblem, vbExclamation, gstrSysName
        Exit Function
    End If
    
    '因血库，设备，病案，手麻安装脚本没有包脚本文件，所以这样处理
    If Dir(strMainPath & "ZLPACKAGE.SQL") <> "" Then
        mrsLocalFile.AddNew Array("FilePath", "SystemNum", "FileName", "FileType"), _
                            Array(strMainPath & "ZLPACKAGE.SQL", strSysNum, "ZLPACKAGE.SQL", "安装脚本")
    End If
    CheckSetFile = True
End Function

Private Sub AddFilePath(ByVal strPath As String, ByVal strFileName As String, ByVal strFileType As String, ByVal strSysNum As Long, ByRef strProblem As String)

    If Dir(strPath) = "" Then
        strProblem = strProblem & vbCr & strFileType & strPath
    Else
        mrsLocalFile.AddNew Array("FilePath", "SystemNum", "FileName", "FileType"), Array(strPath, strSysNum, strFileName, "安装脚本")
    End If
End Sub

Private Sub cmdFunction_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngTemp As Long
    Dim i As Long
    
    For i = 0 To cmdFunction.UBound
        If i = Index Then
            If cmdFunction(Index) Is ActiveControl And cmdFunction(Index).FontBold = True Then Exit Sub
            
            For lngTemp = 0 To cmdFunction.UBound
                cmdFunction(lngTemp).FontBold = False
            Next
            cmdFunction(i).FontBold = True
            cmdFunction(i).SetFocus
            Select Case i
            Case 0
                lblNote.Caption = "通过收集并解析本地安装和升级脚本（截止到当前版本）中的数据结构和基础管理数据，与数据库中正在使用的进行对比检查。" & vbNewLine & _
                            vbCrLf & "检查的数据结构包括：表、字段、约束、索引、序列、视图、包、存储过程。" & vbNewLine & _
                                    "基础管理数据包括：模块、功能、参数、报表、表目录。" & vbNewLine & _
                            vbCrLf & "检查过程比较耗时，请耐心等待。检查结果可输出报告文档，可以选择全部或部分问题进行修复。" & vbNewLine & _
                                    "修复操作可能涉及相关对象的独占操作，将会影响相关产品功能的正常运行，建议在业务空闲期间执行。" & vbNewLine & _
                                    "一般建议在升级完成后执行本操作，以检查的修正因升级脚本执行出错被忽略后可能导致的结构和数据不完整。"
            Case 1
                lblNote.Caption = "公共同义词指向应用系统所有者的实际对象（表、存储过程等），用于普通操作员执行SQL时访问，以避免在SQL的对象名称前添加所有者前缀。" & vbNewLine & _
                            vbCrLf & "如果缺失公共同义词，普通操作人员执行相关SQL时就可能出错。" & vbNewLine & _
                                    "升级完成后会自动进行修正，本功能一般用于未通过升级程序而临时执行脚本后的修正。"
            Case 2
                lblNote.Caption = "以管理工具用户ZLTOOLS执行权限检查和修正，包括补充缺少的公共同义词（ZLTOOLS所有对象的），" & vbNewLine & _
                                    "ZLTOOLS的所有对象授予Public的公共权限、授予应用系统所有者和历史空间所有者的全部权限。" & vbNewLine & _
                            vbCrLf & "升级完成后会自动进行修正，本功能一般用于未通过升级程序而临时执行脚本后的修正。"
            End Select
        End If
    Next
End Sub

Private Sub Form_Load()

    lblNote.Caption = "通过收集并解析本地安装和升级脚本（截止到当前版本）中的数据结构和基础管理数据，与数据库中正在使用的进行对比检查。" & vbNewLine & _
                vbCrLf & "检查的数据结构包括：表、字段、约束、索引、序列、视图、包、存储过程。" & vbNewLine & _
                        "基础管理数据包括：模块、功能、参数、报表、表目录。" & vbNewLine & _
                vbCrLf & "检查过程比较耗时，请耐心等待。检查结果可输出报告文档，可以选择全部或部分问题进行修复。" & vbNewLine & _
                        "修复操作可能涉及相关对象的独占操作，将会影响相关产品功能的正常运行，建议在业务空闲期间执行。" & vbNewLine & _
                        "一般建议在升级完成后执行本操作，以检查的修正因升级脚本执行出错被忽略后可能导致的结构和数据不完整。"
    Call IniVSF
    Call GetVersion
    On Error Resume Next
    gcnOracle.Execute "select 表名 from zltables"
    If err.Number = 0 Then mblnzlTables = True
    err.Clear: On Error GoTo 0
End Sub

Private Sub IniVSF()
'功能：初始化VSF
    Dim rsSys As ADODB.Recordset
    Dim strSQL As String
    Dim i As Long
    
    Set rsSys = GetSystemList
    Call InitTable(vsfSelSys, MSTR_COL)

    With vsfSelSys
        .TextMatrix(.Rows - 1, Col_系统名称) = "服务器管理工具"
        .TextMatrix(.Rows - 1, Col_当前版本) = GetToolsVersion
        .TextMatrix(.Rows - 1, Col_共享号) = 0
        Do While Not rsSys.EOF
            If Val(Split(rsSys!系统版本号, ".")(0)) > 9 And rsSys!系统编号 <> "2300" Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, Col_系统编号) = rsSys!系统编号 & ""
                .TextMatrix(.Rows - 1, Col_系统名称) = rsSys!系统名称 & ""
                .TextMatrix(.Rows - 1, Col_当前版本) = rsSys!系统版本号 & ""
                .TextMatrix(.Rows - 1, Col_所有者) = rsSys!系统所有者 & ""
                .TextMatrix(.Rows - 1, Col_共享号) = rsSys!共享号 & ""
            End If
            rsSys.MoveNext
        Loop
        .Cell(flexcpChecked, 0, 0, .Rows - .FixedRows) = flexChecked
        For i = 0 To .Rows - 1
            .rowHeight(i) = 300
        Next
    End With
End Sub

Private Sub GetVersion(Optional ByVal strMainPath As String)
'获取本地文件的系统版本号
    Dim i As Long
    Dim strTemp As String
    Dim strMaxVer As String
    Dim strPath As String
    Dim varTemp As Variant
    Dim rsTemp As ADODB.Recordset
    Dim rsSetFile As ADODB.Recordset
    Dim blnExist As Boolean
    
    Set rsSetFile = GetSystemSetupIni
    '管理工具的当前脚本版本获取
    If strMainPath <> "" Then
        strTemp = strMainPath
        strPath = strTemp & "\TOOLS\zlServer.sql"
        lblMainPath.Caption = "系统安装目录：" & strMainPath
    Else
        rsSetFile.Filter = "系统编号=100"
        If rsSetFile.RecordCount <> 0 Then
            strTemp = rsSetFile!文件名
            strPath = Mid(strTemp, 1, 1) & ":\APPSOFT\TOOLS\zlServer.sql"
            lblMainPath.Caption = "系统安装目录：" & Mid(strTemp, 1, 1) & ":\Appsoft"
        Else
            strPath = "C:\APPSOFT\TOOLS\ZLSERVER.SQL"
            lblMainPath.Caption = "系统安装目录：C:\Appsoft"
        End If
    End If
    blnExist = True
    With vsfSelSys
        '应用系统的当前脚本版本获取
        For i = .FixedRows To .Rows - .FixedRows
            If i = .FixedRows Then
                strPath = Replace(strPath, "\\", "\")
                If Dir(strPath) <> "" Then
                    .TextMatrix(1, Col_配置文件) = strPath
                Else
                    blnExist = False
                    .TextMatrix(1, Col_配置文件) = ""
                End If
            Else
                rsSetFile.Filter = "系统编号=" & .TextMatrix(i, Col_系统编号) & ""
                If rsSetFile.RecordCount > 0 Then
                    If strMainPath <> "" Then
                        varTemp = Split(rsSetFile!文件名, "APPSOFT")
                        strTemp = strMainPath & varTemp(1)
                    Else
                        strTemp = rsSetFile!文件名
                    End If
                    strTemp = Replace(strTemp, "\\", "\")
                    If Dir(strTemp) <> "" Then
                        .TextMatrix(i, Col_配置文件) = strTemp
                    Else
                        blnExist = False
                        .TextMatrix(i, Col_配置文件) = ""
                    End If
                End If
            End If
            If blnExist Then
                strMaxVer = ""
                varTemp = Split(.TextMatrix(i, Col_当前版本), ".")
                Set rsTemp = GetUpgradeFiles(rsTemp, Val(.TextMatrix(i, Col_系统编号)), "10.34.0", .TextMatrix(i, Col_配置文件), "", "", strMaxVer)
                .Cell(flexcpText, i, Col_脚本版本) = strMaxVer
            Else
                .Cell(flexcpText, i, Col_脚本版本) = ""
            End If
        Next
        If blnExist = False Then MsgBox "安装目录没有脚本文件，请重新选择！"
        .Row = 1
        Call .ShowCell(1, 1)
    End With
End Sub

Private Sub Form_Resize()
    Dim i As Long
    
    On Error Resume Next
    With imgMain
        .Top = 700
        .Left = ScaleLeft + 200
    End With
    
    With vsfSelSys
        .Top = lblMainPath.Top + lblMainPath.Height + 50
        .Width = ScaleWidth - .Left - imgMain.Left
        .ColWidth(Col_空白) = ScaleWidth - .Left - 5000 - 280
        .Height = .Rows * 300 + 50
    End With
    
    cmdFunction(0).Top = vsfSelSys.Top + vsfSelSys.Height + 400
    cmdFunction(0).Left = vsfSelSys.Left

    chkIndex.Top = cmdFunction(0).Top
    chkIndex.Left = cmdFunction(0).Left + cmdFunction(0).Width + 300
    chkReport.Top = chkIndex.Top
    chkReport.Left = chkIndex.Left + chkIndex.Width + 500
    chkProcedure.Top = chkReport.Top
    chkProcedure.Left = chkReport.Left + chkReport.Width + 500
    chkParameters.Top = chkProcedure.Top
    chkParameters.Left = chkProcedure.Left + chkProcedure.Width + 500
    
    lblNote.Left = chkIndex.Left
    lblNote.Top = cmdFunction(0).Top + cmdFunction(0).Height + 50
    lblNote.Width = ScaleWidth - lblNote.Left
    lblNote.Height = 1600
    
    cmdFunction(2).Top = lblNote.Top + lblNote.Height - cmdFunction(2).Height
    cmdFunction(2).Left = cmdFunction(0).Left
    
    cmdFunction(1).Top = cmdFunction(2).Top - cmdFunction(1).Height - 50
    cmdFunction(1).Left = cmdFunction(1).Left
    
End Sub

Private Sub picStatus_Resize()
    If picStatus.ScaleWidth < 1000 Then Exit Sub
    
    With pgbProgress
        .Left = 150
        .Width = (picStatus.ScaleWidth - 150 * 4) / 2
        .Top = 240
        .Height = 180
        lblProgress.Left = .Left
        lblProgress.Top = .Top - lblProgress.Height
    End With
    
    
    With pgbState
        .Left = pgbProgress.Left + pgbProgress.Width + 300
        .Width = pgbProgress.Width
        .Top = 240
        .Height = 180
        lblStatus.Left = .Left
        lblStatus.Top = lblProgress.Top
    End With
    
    With Linepgb
        .x1 = pgbState.Left - 150
        .X2 = pgbState.Left - 150
        .y1 = 0
        .Y2 = picStatus.Height
    End With
End Sub

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口

End Sub

Private Sub lblSel_Click()
    Dim strFolderName As String
    Dim strOldPath As String
    strFolderName = lblMainPath.Tag
    
    strFolderName = OpenFolder(Me, "选择系统安装目录")
    If strFolderName = "" Then Exit Sub
    lblMainPath.Tag = strFolderName
    Call GetVersion(strFolderName)
End Sub

Private Sub vsfSelSys_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    Dim strNum As String
    
    With vsfSelSys
        If Col = Col_选择 Then
            If Row = 0 Then
                If .Cell(flexcpChecked, 0, Col_选择) = flexChecked Then
                    .Cell(flexcpChecked, 0, Col_选择) = flexChecked
                    For i = .FixedRows To .Rows - .FixedRows
                        .Cell(flexcpChecked, i, Col_选择) = flexChecked
                    Next
                Else
                    .Cell(flexcpChecked, 0, Col_选择) = flexUnchecked
                    For i = .FixedRows To .Rows - .FixedRows
                        .Cell(flexcpChecked, i, Col_选择) = flexUnchecked
                    Next
                End If
            ElseIf Row <> 0 Then
                If .Cell(flexcpChecked, 0, Col_选择) = flexChecked Then
                    .Cell(flexcpChecked, 0, Col_选择) = flexUnchecked
                End If
                For i = .FixedRows To .Rows - .FixedRows
                    If .Cell(flexcpChecked, i, Col_选择) = flexUnchecked Then
                        Exit For
                    Else
                        If i = .Rows - .FixedRows Then
                            .Cell(flexcpChecked, 0, Col_选择) = flexChecked
                        End If
                    End If
                Next
            End If
        End If
    End With
End Sub

Private Function CheckHistorySpaceEx(ByVal lngSys As Long) As Boolean
    '功能:检查当前系统下是否有历史表空间的表信息

    Dim rsTmp As New ADODB.Recordset
    
    On Error Resume Next '可能当前用户没有表权限
    gstrSQL = "Select 名称,所有者 From Zltools.Zlbakspaces Where 系统 = " & lngSys & "  And 当前 = 1 And 只读 = 0"
    Call OpenRecordset(rsTmp, gstrSQL, "读取历史表空间所有者")
    CheckHistorySpaceEx = Not rsTmp.EOF
    On Error GoTo 0
End Function

Private Sub Release()
'修正完成后释放模块窗体

    Set mrsSequenceFromFile = Nothing
    Set mrsViewFromFile = Nothing
    Set mrsPackageFromFile = Nothing
    Set mrsFildFromFile = Nothing
    Set mrsConstraintFromFile = Nothing
    Set mrsIndexFromFile = Nothing
    Set mrsProcedureFromFile = Nothing
    Set mrsDataFromFile = Nothing
    
    Set mrsSequenceFromDB = Nothing
    Set mrsViewFromDB = Nothing
    Set mrsPackageFromDB = Nothing
    Set mrsFildFromDB = Nothing
    Set mrsConstraintFromDB = Nothing
    Set mrsIndexFromDB = Nothing
    Set mrsProcedureFromDB = Nothing
    Set mrsDataFromDB = Nothing
End Sub

Private Sub vsfSelSys_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    If Col <> 0 Then Cancel = True
End Sub


