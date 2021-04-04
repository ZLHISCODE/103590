VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{B26F6243-4C7D-11D1-910E-00600807163F}#2.78#0"; "Xcdzip35.ocx"
Begin VB.Form frmPriceImp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医价数据导入"
   ClientHeight    =   3075
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6210
   FillColor       =   &H80000012&
   Icon            =   "frmPriceImp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭(&X)"
      Height          =   350
      Left            =   4845
      TabIndex        =   9
      Top             =   2550
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   135
      Picture         =   "frmPriceImp.frx":058A
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2550
      Width           =   1100
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "导入(&I)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   3735
      TabIndex        =   7
      Top             =   2550
      Width           =   1100
   End
   Begin VB.Frame frmLine 
      Height          =   45
      Left            =   -30
      TabIndex        =   5
      Top             =   2370
      Width           =   6360
   End
   Begin VB.TextBox txtFile 
      Height          =   300
      Left            =   1110
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1245
      Width           =   4785
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "选择(&S)…"
      Height          =   350
      Left            =   4815
      TabIndex        =   2
      Top             =   885
      Width           =   1100
   End
   Begin MSComctlLib.ProgressBar pgbImp 
      Height          =   240
      Left            =   1110
      TabIndex        =   4
      Top             =   2115
      Visible         =   0   'False
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComDlg.CommonDialog cdgThis 
      Left            =   120
      Top             =   1620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblImp 
      AutoSize        =   -1  'True
      Caption         =   "正在导入医价数据"
      Height          =   180
      Left            =   1095
      TabIndex        =   6
      Top             =   1875
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      Caption         =   "标准医价文件"
      Height          =   180
      Left            =   1095
      TabIndex        =   1
      Top             =   975
      Width           =   1080
   End
   Begin XCEEDZIPLib.XceedZip zip 
      Left            =   135
      Top             =   825
      _Version        =   131150
      _ExtentX        =   794
      _ExtentY        =   794
      _StockProps     =   0
      Compression     =   6
      ClearDisks      =   0   'False
      ExtractDirectory=   ""
      FilesToProcess  =   ""
      IncludeDirectoryEntries=   0   'False
      IncludeHiddenFiles=   0   'False
      IncludeVolumeLabel=   0   'False
      ModifiedDate    =   "01011980"
      MoveFiles       =   0   'False
      MultidiskMode   =   0   'False
      Overwrite       =   0
      Password        =   ""
      Recurse         =   0   'False
      SelfExtracting  =   0   'False
      SfxBinary       =   ""
      SfxConfigFile   =   ""
      StoredExtensions=   ".ZIP;.LZH;.ARC;.ARJ;.ZOO"
      TempPath        =   ""
      UsePaths        =   -1  'True
      UseTempFile     =   -1  'True
      ZipFileName     =   ""
      InternalState   =   "7f6ba9d4"
      SfxExtractDirectory=   ""
      SfxRunExePath   =   ""
      SfxReadmePath   =   ""
      SfxDefaultPassword=   ""
      SfxOverwrite    =   0
      SfxPromptForDirectory=   -1  'True
      SfxShowProgress =   -1  'True
      SfxPromptForPassword=   -1  'True
      SfxPromptCreateDirectory=   -1  'True
      SfxProgramGroup =   ""
      SfxProgramGroupItems=   ""
      SfxRegisterExtensions=   ""
      SfxInstallMode  =   0   'False
   End
   Begin VB.Image imgICO 
      Height          =   480
      Left            =   135
      Picture         =   "frmPriceImp.frx":0E54
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblNote 
      Caption         =   "    选择正确的医价数据文件，将医价数据完整导入本系统，以便保证系统的收费项目和价格符合价格政策的规定。"
      Height          =   390
      Left            =   705
      TabIndex        =   0
      Top             =   165
      Width           =   5220
   End
End
Attribute VB_Name = "frmPriceImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private strTmpPath As String

Dim objFile As New FileSystemObject
Dim objText As TextStream

Private Function GetDateString(ByVal strDateString As String) As String
    '一个临时的把输入串转换为标准日期格式的字串的函数，主要功能是去掉毫秒
    '只针对一般的日期格式串，如"2005-8-15 13:51:32:953"，去掉毫秒就是"2005-8-15 13:51:32"
    Dim strInput As String      '保存输入字串
    Dim strOutput As String     '输出字串
    Dim strDatePart As String   '日期部分
    Dim strTimePart As String   '时间部分
    Dim intSpace As Integer     '空格所在位置
    Const cstDateTimeFormat = "yyyy-mm-dd hh:mm:ss"     '标准的日期时间格式
    Const cstDateFormat = "yyyy-mm-dd"                  '标准的日期格式
    Const cstTimeFormat = "hh:mm:ss"                    '标准的时间格式
    Dim strTime() As String     '临时保存时间分隔的数组
    
    '先去掉首尾空格
    strInput = Trim(strDateString)
    
    '如果为空就退出
    If strInput = "" Then
        strOutput = ""
        GetDateString = strOutput
        Exit Function
    End If
    
    '如果输入串可以转化为日期，就直接转换为标准的日期格式字串输出
    If IsDate(strInput) Then
        strOutput = Format(CDate(strInput), cstDateTimeFormat)
        GetDateString = strOutput
        Exit Function
    End If
    
    '判断输入串中间是否存在空格
    intSpace = InStr(strInput, " ")
    If intSpace > 0 Then    '存在空格就分隔为日期和时间部分
        strDatePart = Mid(strInput, 1, intSpace - 1)
        strTimePart = Mid(strInput, intSpace + 1)
    Else    '没有空格，显然不是正确的日期串，只能退出
        GetDateString = ""
        Exit Function
    End If
    
    '判断日期字串部分是否可以转换为日期
    If IsDate(strDatePart) Then
        strDatePart = Format(CDate(strDatePart), cstDateFormat)
    Else    '不能转换，只能退出
        GetDateString = ""
        Exit Function
    End If
    
    '把用:分隔的时间部分分解成数组
    strTime = Split(strTimePart, ":")
    
    If UBound(strTime) > 2 Then     '如果数组上限大于2，即分解成了4个部分，也就说明含有毫秒
        '只保留时分秒部分
        strTimePart = strTime(0) & ":" & strTime(1) & ":" & strTime(2)
    End If
    
    '判断时间字串部分是否可以转换为日期
    If IsDate(strTimePart) Then
        strTimePart = Format(CDate(strTimePart), cstTimeFormat)
    Else    '不能转换，只能退出
        GetDateString = ""
        Exit Function
    End If
    
    '重新组合字串
    strOutput = strDatePart & " " & strTimePart
    
    GetDateString = strOutput
    
End Function
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdExecute_Click()
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim lngCount As Long
    Dim strLine As String
    Dim aryField() As String
    
    If MsgBox("真的现在导入标准医价文件吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    Err = 0: On Error Resume Next
    Set objText = objFile.OpenTextFile(strTmpPath & "\item.txt")
    If Err <> 0 Then
        MsgBox "无法打开医价标准文件！", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    Do Until objText.AtEndOfStream
        objText.ReadLine
    Loop
    lngCount = objText.Line
    objText.Close
    
    '开始导入
    Err = 0: On Error GoTo ErrHand
    Set objText = objFile.OpenTextFile(strTmpPath & "\item.txt")
    
    Me.lblImp.Visible = True: Me.pgbImp.Visible = True
    DoEvents
    
    gcnOracle.BeginTrans
    gstrSQL = "Delete From 标准医价规范"
    gcnOracle.Execute gstrSQL
    
    Do While Not objText.AtEndOfStream
        strLine = Trim(objText.ReadLine)
        aryField = Split(strLine, vbTab)
        gstrSQL = "Insert Into 标准医价规范(项目编码, 项目名称, 拼音码, 项目别名, 计价单位, 项目内涵, 除外内容, 项目说明, 项目价格, 重复标志, 医院等级, 注销标志, 财务编码, 最高限价, 最低限价, 调价日期)"
        gstrSQL = gstrSQL & " Values('" & Trim(Replace(aryField(0), "'", "''")) & "'"
        gstrSQL = gstrSQL & ",'" & Trim(Replace(aryField(1), "'", "''")) & "'"
        gstrSQL = gstrSQL & ",'" & Trim(Replace(aryField(2), "'", "''")) & "'"
        gstrSQL = gstrSQL & ",'" & Trim(Replace(aryField(3), "'", "''")) & "'"
        gstrSQL = gstrSQL & ",'" & Trim(Replace(aryField(4), "'", "''")) & "'"
        gstrSQL = gstrSQL & ",'" & Trim(Replace(aryField(5), "'", "''")) & "'"
        gstrSQL = gstrSQL & ",'" & Trim(Replace(aryField(6), "'", "''")) & "'"
        gstrSQL = gstrSQL & ",'" & Trim(Replace(aryField(7), "'", "''")) & "'"
        gstrSQL = gstrSQL & "," & Format(IIf(Not IsNumeric(Replace(aryField(8), "'", "''")), 0, Replace(aryField(8), "'", "''")), "0.00")
        gstrSQL = gstrSQL & ",'" & Trim(Replace(aryField(9), "'", "''")) & "'"
        gstrSQL = gstrSQL & ",'" & Trim(Replace(aryField(10), "'", "''")) & "'"
        gstrSQL = gstrSQL & ",'" & Trim(Replace(aryField(11), "'", "''")) & "'"
        gstrSQL = gstrSQL & ",'" & Trim(Replace(aryField(12), "'", "''")) & "'"
        gstrSQL = gstrSQL & "," & Format(IIf(Not IsNumeric(Replace(aryField(13), "'", "''")), 0, Replace(aryField(13), "'", "''")), "0.00")
        gstrSQL = gstrSQL & "," & Format(IIf(Not IsNumeric(Replace(aryField(14), "'", "''")), 0, Replace(aryField(14), "'", "''")), "0.00")
        If InStr(1, aryField(15), ".") > 0 Then
            aryField(15) = Mid(aryField(15), 1, InStr(1, aryField(15), ".") - 1)
        End If
'        gstrSQL = gstrSQL & ",to_date('" & Format(aryField(15), "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS'))"
        If GetDateString(aryField(15)) = "" Then
            gstrSQL = gstrSQL & ",NULL)"
        Else
            gstrSQL = gstrSQL & ",to_date('" & GetDateString(aryField(15)) & "','YYYY-MM-DD HH24:MI:SS'))"
        End If
        gcnOracle.Execute gstrSQL
        Me.pgbImp.Value = Int(objText.Line / lngCount * 100)
    Loop
    gcnOracle.CommitTrans
    objText.Close
    
    MsgBox "标准医价导入成功完成！", vbExclamation, gstrSysName
    Me.lblImp.Visible = False: Me.pgbImp.Visible = False
    Exit Sub

ErrHand:
    gcnOracle.RollbackTrans
    objText.Close
    MsgBox "标准医价导入失败，请系统管理员检查医价文件！", vbExclamation, gstrSysName
    Me.lblImp.Visible = False: Me.pgbImp.Visible = False
End Sub

Private Sub cmdFile_Click()
    With Me.cdgThis
        .FileName = Me.txtFile.Text
        .DialogTitle = "选择标准医价文件"
        .Filter = "(标准医价文件)|*.zl"
        .ShowOpen
        If .FileName = "" Then
            Exit Sub
        Else
            Me.txtFile.Text = .FileName
        End If
        If Dir(Me.txtFile.Text) = "" Then
            MsgBox "医价文件不存在！", vbExclamation, gstrSysName
            Me.txtFile.Text = ""
            Me.cmdExecute.Enabled = False
            Exit Sub
        End If
    End With
    
    Err = 0: On Error Resume Next
    Kill strTmpPath & "\item.txt"
    Err = 0: On Error GoTo 0
    With Me.zip
        .FilesToProcess = "*"
        .Password = "zlhis"
        .UsePaths = False
        .ZipFileName = Me.txtFile.Text
        .ExtractDirectory = strTmpPath
        .Extract (0)
    End With
    
    If Dir(strTmpPath & "\item.txt") = "" Then
        MsgBox "该文件不是正确的医价文件！", vbExclamation, gstrSysName
        Me.txtFile.Text = ""
        Me.cmdExecute.Enabled = False
    Else
        Me.cmdExecute.Enabled = True
    End If
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Form_Load()
    Dim strInput As String * 255
    Call GetTempPath(255, strInput)
    strTmpPath = Left(strInput, InStr(strInput, Chr(0)) - 1)
End Sub
