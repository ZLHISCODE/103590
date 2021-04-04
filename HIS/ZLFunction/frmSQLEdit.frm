VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSQLEdit 
   AutoRedraw      =   -1  'True
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7425
   Icon            =   "frmSQLEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   7425
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ListView lvwFunc 
      Height          =   4695
      Left            =   2520
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   8281
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "函数名"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.PictureBox picItem 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   0
      ScaleHeight     =   1680
      ScaleWidth      =   7425
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   7425
      Begin VB.CommandButton cmdFunc 
         Caption         =   "选择(&S)"
         Height          =   350
         Left            =   4755
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "F3:选择系统函数"
         Top             =   120
         Width           =   1100
      End
      Begin VB.TextBox txt函数名 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   0
         Top             =   180
         Width           =   3090
      End
      Begin VB.TextBox txt说明 
         BackColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   840
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   840
         Width           =   5025
      End
      Begin VB.TextBox txt中文名 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   840
         MaxLength       =   100
         TabIndex        =   1
         Top             =   495
         Width           =   5025
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Left            =   840
         ScaleHeight     =   240
         ScaleWidth      =   3795
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   135
         Width           =   3855
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ZL_FUN_"
            ForeColor       =   &H00808080&
            Height          =   180
            Left            =   30
            TabIndex        =   16
            Top             =   15
            Width           =   630
         End
      End
      Begin MSComctlLib.ImageList img16 
         Left            =   6240
         Top             =   255
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSQLEdit.frx":014A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "函数体：如果要使用动态时间参数，请以 zlBeginTime 或 zlEndTime 作为参数名。"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   225
         TabIndex        =   20
         Top             =   1455
         Width           =   6660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "说明"
         Height          =   180
         Left            =   405
         TabIndex        =   19
         Top             =   900
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "函数名"
         Height          =   180
         Left            =   225
         TabIndex        =   18
         Top             =   195
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "中文名"
         Height          =   180
         Left            =   225
         TabIndex        =   17
         Top             =   555
         Width           =   540
      End
   End
   Begin VB.PictureBox picEdit 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      Height          =   3945
      Left            =   15
      ScaleHeight     =   3885
      ScaleWidth      =   7350
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1680
      Width           =   7410
      Begin VB.TextBox txtSQL 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3885
         Left            =   250
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   0
         Width           =   7095
      End
   End
   Begin VB.PictureBox picDown 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   7425
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5640
      Width           =   7425
      Begin VB.PictureBox picCmd 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   420
         ScaleHeight     =   345
         ScaleWidth      =   5400
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   135
         Width           =   5400
         Begin VB.CommandButton cmdCancel 
            Caption         =   "取消(&C)"
            Height          =   350
            Left            =   4305
            TabIndex        =   7
            Top             =   0
            Width           =   1100
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "确定(&O)"
            Height          =   350
            Left            =   2880
            TabIndex        =   6
            Top             =   0
            Width           =   1100
         End
         Begin VB.CommandButton cmdCompile 
            Caption         =   "编译(&M)"
            Height          =   350
            Left            =   15
            TabIndex        =   4
            Top             =   0
            Width           =   1100
         End
         Begin VB.CommandButton cmdPar 
            Caption         =   "参数(&P)"
            Height          =   350
            Left            =   1440
            TabIndex        =   5
            Top             =   0
            Width           =   1100
         End
      End
      Begin VB.Frame fra1 
         Height          =   30
         Left            =   -15
         TabIndex        =   12
         Top             =   0
         Width           =   7785
      End
   End
End
Attribute VB_Name = "frmSQLEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'入口参数：
Public mblnModi As Boolean
Public mlngSys As Long '函数所在系统
Public mstrOwner As String '函数的所有者

Private mintNum As Integer '函数号
Private mobjPars As FuncPars '函数参数集

Private WithEvents objTab As clsTabInput
Attribute objTab.VB_VarHelpID = -1
Private lngSelStart As Long
Private lngSelLen As Long

Private mblnSQLNoDo As Boolean
Private mblnTXTNoDo As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFunc_Click()
    lvwFunc.Visible = True
    lvwFunc.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String, strPar As String
    Dim i As Integer, arrPar() As String
    Dim StrName As String, j As Integer
    Dim strObject As String, objPar As FuncPar
    
    If Not CheckLen(txt函数名, txt函数名.MaxLength, "函数名", False) Then Exit Sub
    If Not CheckLen(txt中文名, txt中文名.MaxLength, "函数中文名", False) Then Exit Sub
    If Not CheckLen(txt说明, txt说明.MaxLength, "函数说明") Then Exit Sub
    
    strSQL = Trim(txtSQL.Text)
    StrName = GetFuncName(strSQL)
    If StrName = "" Then
        MsgBox "函数代码书写错误！请检查。", vbInformation, App.Title
        Exit Sub
    End If
    If UCase(StrName) <> UCase("ZL_FUN_" & txt函数名.Text) Then
        MsgBox "函数代码中书写的函数名与定义的不一致！请修正。", vbInformation, App.Title
        Exit Sub
    End If
    strSQL = FuncOwnerName(strSQL, StrName, mstrOwner)
    
    If Not CheckFunc(strSQL) Then Exit Sub
    strPar = GetFuncPars(strSQL)
    
    If strPar <> "" And mobjPars.Count = 0 Then
        If MsgBox("函数代码中定义了参数,但没有为这些参数设置取值方法,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
    Else
        arrPar = Split(strPar, ";")
        For i = 0 To UBound(arrPar)
            StrName = Split(arrPar(i), ",")(0)
            For j = 1 To mobjPars.Count
                If UCase(mobjPars(j).名称) = UCase(StrName) Then
                    Exit For
                End If
            Next
            If j > mobjPars.Count Then
                If MsgBox("函数代码中定义了参数""" & StrName & """,但没有它设置取值方法,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
                If MsgBox("要忽略所有未设置取值方法的参数吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then Exit For
            End If
        Next
    End If
    
    '删除无效参数设置
    For Each objPar In mobjPars
        If InStr(";" & strPar, ";" & objPar.名称 & ",") = 0 Then
            mobjPars.Remove "_" & objPar.名称
        Else
            If objPar.缺省值 <> "选择器定义…" Then
                objPar.对象 = ""
                objPar.分类SQL = ""
                objPar.分类字段 = ""
                objPar.明细SQL = ""
                objPar.明细字段 = ""
                If Not objPar.缺省值 Like "*…" Then objPar.值列表 = ""
            End If
        End If
    Next
    
    '检查对象权限
    For i = 1 To mobjPars.Count
        If mobjPars(i).分类SQL <> "" Then
            strObject = strObject & "," & SQLObject(mobjPars(i).分类SQL)
        End If
        If mobjPars(i).明细SQL <> "" Then
            strObject = strObject & "," & SQLObject(mobjPars(i).明细SQL)
        End If
    Next
    strObject = Mid(strObject, 2)
    strObject = CheckObjectPriv(strObject, mstrOwner)
    
    If strObject <> "" Then
        MsgBox "当前用户不具有下列对象或没有权限访问这些对象:" & vbCrLf & vbCrLf & strObject, vbInformation, App.Title
        Exit Sub
    End If
    
    '保存
    Screen.MousePointer = 11
    If Not SaveFunc Then
        Screen.MousePointer = 0: Exit Sub
    End If
    Screen.MousePointer = 0
    
    gblnOK = True
    Unload Me
End Sub

Private Sub cmdPar_Click()
    Dim strSQL As String, strPar As String
    Dim blnOK As Boolean, objPars As FuncPars
    
    strSQL = Trim(txtSQL.Text)
    strPar = GetFuncPars(strSQL)
    
    If strPar = "" Then
        Set mobjPars = New FuncPars
        MsgBox "在函数代码中没有定义参数！", vbInformation, App.Title
        Exit Sub
    End If
    
    blnOK = gblnOK
    
    frmParEdit.mstrPars = strPar
    frmParEdit.mlngSys = mlngSys
    frmParEdit.mstrOwner = mstrOwner
    Set frmParEdit.mobjPars = mobjPars
    frmParEdit.Show 1, Me
    If gblnOK Then
        Call CopyPars(frmParEdit.mobjPars, mobjPars)
        Unload frmParEdit
    End If
    
    gblnOK = blnOK
End Sub

Private Sub cmdCompile_Click()
    Dim strSQL As String, strPar As String
    Dim arrPar() As String, i As Integer
    Dim StrName As String, strText As String
    
    strSQL = Trim(txtSQL.Text)
    If strSQL = "" Then
        MsgBox "请先输入函数代码！", vbInformation, App.Title
        txtSQL.SetFocus: Exit Sub
    End If
    
    StrName = GetFuncName(strSQL)
    If StrName = "" Then
        MsgBox "函数代码书写错误！请检查。", vbInformation, App.Title
        Exit Sub
    End If
    If UCase(StrName) <> UCase("ZL_FUN_" & txt函数名.Text) Then
        MsgBox "函数代码中书写的函数名与定义的不一致！请修正。", vbInformation, App.Title
        Exit Sub
    End If
    strSQL = FuncOwnerName(strSQL, StrName, mstrOwner)
    
    If Not mblnModi Then
        grsObject.Filter = "OWNER='" & UCase(mstrOwner) & "' And OBJECT_TYPE='FUNCTION' And OBJECT_NAME='" & UCase(StrName) & "'"
        If Not grsObject.EOF Then
            If MsgBox("系统中已经存在同名的函数,要覆盖它吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then
                If MsgBox("要读取已存在的函数代码替换当前书写的代码吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbYes Then
                    strText = GetFunSource(mstrOwner, StrName)
                    strText = GetShowCode(strText, StrName)
                    txtSQL.Text = strText
                End If
                Exit Sub
            End If
        End If
    End If
    
    If Not CheckFunc(strSQL) Then Exit Sub
    
    txtSQL.Tag = txtSQL.Text
    cmdPar.Enabled = True
    cmdOK.Enabled = True
End Sub

Private Function CheckFunc(strSQL As String) As Boolean
'功能：检查函数代码的正确性
    Dim arrPar() As String, i As Integer, strPar As String
        
    Screen.MousePointer = 11
    
    On Error Resume Next
        
    'vbCr在数据库中会转换成空格,只保留vbLf
    gcnOracle.Execute Replace(strSQL, vbCr, "")
    
    '函数内部错误不会激活Error错误。
    If gcnOracle.Errors.Count > 0 Then
        Screen.MousePointer = 0
        MsgBox gcnOracle.Errors(0).Description, vbExclamation, App.Title
        Exit Function
    Else
        Screen.MousePointer = 0
        '检查参数合法性
        strPar = GetFuncPars(strSQL)
        arrPar = Split(strPar, ";")
        
        If InStr(UCase(strPar), "%TYPE") > 0 Then
            MsgBox "在函数中不允许参数类型使用表字段进行定义！", vbInformation, App.Title
            lngSelStart = InStr(UCase(txtSQL.Text), "%TYPE")
            lngSelLen = 0: txtSQL.SetFocus: Exit Function
        End If
        
        '检查动态时间参数
        If InStr(UCase(";" & strPar), ";ZLBEGINTIME,") > 0 And InStr(UCase(";" & strPar), ";ZLENDTIME,") = 0 Then
            MsgBox "动态时间参数 zlBeginTime 和 zlEndTime 必须配合使用！", vbInformation, App.Title
            lngSelStart = InStr(UCase(txtSQL.Text), "ZLBEGINTIME")
            lngSelLen = 0: txtSQL.SetFocus: Exit Function
        End If
        If InStr(UCase(";" & strPar), ";ZLBEGINTIME,") = 0 And InStr(UCase(";" & strPar), ";ZLENDTIME,") > 0 Then
            MsgBox "动态时间参数 zlBeginTime 和 zlEndTime 必须配合使用！", vbInformation, App.Title
            lngSelStart = InStr(UCase(txtSQL.Text), "ZLENDTIME")
            lngSelLen = 0: txtSQL.SetFocus: Exit Function
        End If
        For i = 0 To UBound(arrPar)
            If UCase(Split(arrPar(i), ",")(0)) = "ZLBEGINTIME" And Not UCase(Split(arrPar(i), ",")(1)) Like "*DATE*" Then
                MsgBox "动态时间参数zlBeginTime必须定义为日期类型！", vbInformation, App.Title
                lngSelStart = InStr(UCase(txtSQL.Text), "ZLBEGINTIME") + Len("ZLBEGINTIME")
                lngSelLen = 0: txtSQL.SetFocus: Exit Function
            ElseIf UCase(Split(arrPar(i), ",")(0)) = "ZLENDTIME" And Not UCase(Split(arrPar(i), ",")(1)) Like "*DATE*" Then
                MsgBox "动态时间参数zlEndTime必须定义为日期类型！", vbInformation, App.Title
                lngSelStart = InStr(UCase(txtSQL.Text), "ZLENDTIME") + Len("ZLENDTIME")
                lngSelLen = 0: txtSQL.SetFocus: Exit Function
            End If
        Next
        
        '检查参数名长度
        For i = 0 To UBound(arrPar)
            If TLen(CStr(Split(arrPar(i), ",")(0))) > 30 Then
                MsgBox "参数""" & Split(arrPar(i), ",")(0) & """的长度不能超过30个字符！", vbInformation, App.Title
                lngSelStart = InStr(UCase(txtSQL.Text), Split(arrPar(i), ",")(0))
                lngSelLen = 0: txtSQL.SetFocus: Exit Function
            End If
        Next
    End If
    CheckFunc = True
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        If cmdFunc.Enabled And cmdFunc.Visible Then Call cmdFunc_Click
    ElseIf KeyCode = vbKeyEscape Then
        If lvwFunc.Visible Then
            lvwFunc.Visible = False
            txt函数名.SetFocus
        Else
            cmdCancel_Click
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim strText As String
    
    RestoreWinState Me, App.ProductName
    
    If Not InDesign Then
        glngMinW = 400: glngMinH = 260
        glngMaxW = 1600: glngMaxH = 1200
        glngOldProc = GetWindowLong(hwnd, GWL_WNDPROC)
        Call SetWindowLong(hwnd, GWL_WNDPROC, AddressOf CustomMessage)
    End If
    
    gblnOK = False
    mblnSQLNoDo = False
    mblnTXTNoDo = False
    
    lngSelStart = -1
    
    If mblnModi Then
        Caption = "修改函数"
        
        cmdFunc.Visible = False
        txt函数名.Enabled = False
        
        With frmMain.lvw.SelectedItem
            mintNum = Val(Mid(.Key, 2))
            txt函数名.Text = IIf(UCase(.Text) Like "ZL_FUN_*", Mid(.Text, 8), .Text)
            txt中文名.Text = .SubItems(2)
            txt说明.Text = .SubItems(3)
            
            strText = GetFunSource(mstrOwner, .Text)
            Set mobjPars = ReadFuncPars(mlngSys, mintNum)
            strText = GetShowCode(strText, .Text)
        End With
    Else
        Caption = "新增函数"
        mintNum = NextFuncNo(mlngSys)
        txt函数名.Text = mintNum
        txt中文名.Text = ""
        txt说明.Text = ""
    
        strText = _
            "Create Or Replace Function ZL_FUN_" & txt函数名.Text & vbCrLf & _
            "Return Number" & vbCrLf & _
            "As" & vbCrLf & _
            "Begin" & vbCrLf & _
            "    Return(0);" & vbCrLf & _
            "End;"
        Set mobjPars = New FuncPars
        
        Call FillUserFunc
    End If
    txtSQL.Tag = strText '记录原代码
    txtSQL.Text = txtSQL.Tag
End Sub

Private Function GetShowCode(strCode As String, strFunc As String) As String
'功能：格式化函数代码
    Dim strTmp As String, i As Long
    
    strTmp = strCode
    
    strTmp = Trim(Mid(strTmp, InStr(UCase(strTmp), UCase(strFunc))))
    strTmp = "Create Or Replace Function " & strTmp
    
    '处理回车:仅vbCrlf在文本框中不能正确显示
    strTmp = Replace(strTmp, vbCrLf, Chr(188))
    strTmp = Replace(strTmp, vbCr, Chr(188))
    strTmp = Replace(strTmp, vbLf, Chr(188))
    strTmp = Replace(strTmp, Chr(188), vbCrLf)
    If Left(strTmp, 1) = vbCr Or Left(strTmp, 1) = vbLf Then strTmp = Mid(strTmp, 2)
    If Left(strTmp, 1) = vbCr Or Left(strTmp, 1) = vbLf Then strTmp = Mid(strTmp, 2)
    GetShowCode = strTmp
End Function

Private Sub Form_Resize()
    On Error Resume Next
    
    picEdit.Left = ScaleLeft + 30
    picEdit.Top = ScaleTop + picItem.Height
    picEdit.Width = ScaleWidth - 60
    picEdit.Height = ScaleHeight - picDown.Height - picItem.Height
    
    txtSQL.Left = 0
    txtSQL.Top = 0
    txtSQL.Width = picEdit.ScaleWidth
    txtSQL.Height = picEdit.ScaleHeight
    
    fra1.Width = Width + 60
    
    If ScaleWidth - picCmd.Width - cmdOK.Width / 2 > 300 Then
        picCmd.Left = ScaleWidth - picCmd.Width - cmdOK.Width / 2
    Else
        picCmd.Left = 300
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjPars = Nothing
    
    Set objTab = Nothing
    Set gobjTab = Nothing '这个可以引起对象被释放
    
    SaveWinState Me, App.ProductName
    
    If Not InDesign Then Call SetWindowLong(hwnd, GWL_WNDPROC, glngOldProc)
End Sub

Private Sub lvwFunc_DblClick()
    Dim strText As String, StrName As String
    
    If lvwFunc.SelectedItem Is Nothing Then Exit Sub
    
    strText = GetFunSource(mstrOwner, lvwFunc.SelectedItem.Text)
    strText = GetShowCode(strText, lvwFunc.SelectedItem.Text)
    txtSQL.Text = strText
    
    StrName = IIf(UCase(lvwFunc.SelectedItem.Text) Like "ZL_FUN_*", Mid(lvwFunc.SelectedItem.Text, 8), lvwFunc.SelectedItem.Text)
    If UCase(StrName) Like "ZL_FUN*" Then StrName = Mid(StrName, 7)
    If UCase(StrName) Like "ZL_*" Then StrName = Mid(StrName, 4)
    If UCase(StrName) Like "ZL*_*" Then StrName = Mid(StrName, InStr(StrName, "_") + 1)
    
    If txt函数名.Text = StrName Then
        txt函数名.Text = StrName
        Call txt函数名_Change
    Else
        txt函数名.Text = StrName
    End If
    
    lvwFunc.Visible = False
    txt函数名.SetFocus
End Sub

Private Sub lvwFunc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call lvwFunc_DblClick
End Sub

Private Sub lvwFunc_LostFocus()
    lvwFunc.Visible = False
End Sub

Private Sub lvwFunc_Validate(Cancel As Boolean)
    lvwFunc.Visible = False
End Sub

Private Sub objTab_sTabKeyDown()
    Dim strNew As String
    Dim lngStart As Long
    
    If Not ActiveControl Is txtSQL Then Exit Sub
    
    If Len(txtSQL.SelText) = 0 Then
        If txtSQL.SelStart > 1 Then
            If Mid(txtSQL.Text, txtSQL.SelStart - 1, 1) = vbTab Then
                txtSQL.SelStart = txtSQL.SelStart - 1
                txtSQL.SelLength = 1
                txtSQL.SelText = ""
            End If
        ElseIf txtSQL.SelStart = 1 And Left(txtSQL.Text, 1) = vbTab Then
            txtSQL.SelStart = 0
            txtSQL.SelLength = 1
            txtSQL.SelText = ""
        End If
    Else '成段处理
        lngStart = txtSQL.SelStart
        
        '选中第一行特殊处理(这时开头没有vbCr,也没有vbLf)
        strNew = Mid(Replace(vbCr & txtSQL.SelText, vbCr & vbTab, vbCr), 2)
        '其它行(开头是vbLf)
        strNew = Replace(strNew, vbLf & vbTab, vbLf)
        
        If txtSQL.SelText <> strNew Then
            txtSQL.SelText = strNew
            txtSQL.SelStart = lngStart
            txtSQL.SelLength = Len(strNew)
        End If
    End If
End Sub

Private Sub objTab_TabKeyDown()
    Dim strNew As String
    Dim lngStart As Long
        
    If Not ActiveControl Is txtSQL Then Exit Sub
    
    If Len(txtSQL.SelText) = 0 Then
        txtSQL.SelText = vbTab
    Else '成段处理
        lngStart = txtSQL.SelStart
        
        '选中第一行特殊处理(这时开头没有vbCr,也没有vbLf)
        strNew = Mid(Replace(vbCrLf & txtSQL.SelText, vbCrLf, vbCrLf & vbTab), 3)
        '最后行以vbCrLf结束
        If Right(strNew, 3) = vbCrLf & vbTab Then strNew = Left(strNew, Len(strNew) - 1)
        
        If txtSQL.SelText <> strNew Then
            txtSQL.SelText = strNew
            txtSQL.SelStart = lngStart
            txtSQL.SelLength = Len(strNew)
        End If
    End If
End Sub

Private Sub txtSQL_Change()
    Dim StrName As String
    
    If mblnSQLNoDo Then Exit Sub
    
    If txtSQL.Text <> txtSQL.Tag Then
        cmdOK.Enabled = False
        cmdPar.Enabled = False
    Else
        cmdOK.Enabled = True
        cmdPar.Enabled = True
    End If
    
    '根据函数代码自动更名
    If Not mblnModi And Visible Then
        StrName = GetFuncName(txtSQL.Text)
        If StrName <> "" Then
            mblnTXTNoDo = True
            txt函数名.Text = IIf(UCase(StrName) Like "ZL_FUN_*", Mid(StrName, 8), StrName)
            mblnTXTNoDo = False
        End If
    End If
End Sub

Private Sub txtSQL_GotFocus()
    If lngSelStart >= 0 Then
        txtSQL.SelStart = lngSelStart
        txtSQL.SelLength = lngSelLen
    End If
    Set objTab = New clsTabInput
End Sub

Private Sub txtSQL_KeyDown(KeyCode As Integer, Shift As Integer)
'Ctrl+A
    If KeyCode = vbKeyA And Shift = 2 Then SelAll txtSQL
End Sub

Private Sub txtSQL_KeyPress(KeyAscii As Integer)
    '
End Sub

Private Sub txtSQL_LostFocus()
    lngSelStart = txtSQL.SelStart
    lngSelLen = txtSQL.SelLength
    Set objTab = Nothing
    Set gobjTab = Nothing '这个可以引起对象被释放
End Sub

Private Sub txt函数名_Change()
    Dim StrName As String, strTmp As String
    Dim i As Integer
    
    If mblnTXTNoDo Then Exit Sub
    
    '更改名称时自动修改函数代码
    If Visible Then
        strTmp = txtSQL.Text
        StrName = GetFuncName(strTmp)
        If StrName = "" Then Exit Sub
        
        i = InStr(strTmp, StrName)
        Do While i > 0
            strTmp = Left(strTmp, i - 1) & "@#$" & Mid(strTmp, i + Len(StrName))
            i = InStr(strTmp, StrName)
        Loop
        lngSelStart = txtSQL.SelStart
        lngSelLen = 0
        mblnSQLNoDo = True
        txtSQL.Text = Replace(strTmp, "@#$", "ZL_FUN_" & txt函数名.Text)
        txtSQL.SelStart = lngSelStart
        mblnSQLNoDo = False
    End If
End Sub

Private Sub txt函数名_GotFocus()
    SelAll txt函数名
End Sub

Private Sub txt函数名_KeyPress(KeyAscii As Integer)
    If InStr("'&", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub txt说明_GotFocus()
    SelAll txt说明
End Sub

Private Sub txt说明_KeyPress(KeyAscii As Integer)
    If InStr("'&", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub txt中文名_GotFocus()
    SelAll txt中文名
End Sub

Private Sub txt中文名_KeyPress(KeyAscii As Integer)
    If InStr("'&", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Function NextFuncNo(lngSys As Long) As Integer
'功能：获取指定系统中的一个新函数号
'说明：自动取缺号
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = _
        "Select A.函数号 From zlFunctions A " & _
        "Where A.系统=" & lngSys & " And Not Exists(Select 函数号 From zlFunctions B Where B.函数号=A.函数号+1 And B.系统=A.系统) " & _
        "Order by A.函数号"
    
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "NextFuncNo")
    
    If Not rsTmp.EOF Then
        NextFuncNo = rsTmp!函数号 + 1
    Else
        NextFuncNo = 1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveFunc() As Boolean
    Dim i As Integer, arrSQL() As String, objItem As ListItem
    
    ReDim arrSQL(0)
    If mblnModi Then
        arrSQL(UBound(arrSQL)) = "Update zlFunctions Set" & _
            " 函数名='ZL_FUN_" & txt函数名.Text & "',中文名='" & txt中文名.Text & "',说明='" & txt说明.Text & "'" & _
            " Where 系统=" & mlngSys & " And 函数号=" & mintNum
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Delete From zlFuncPars Where 系统=" & mlngSys & " And 函数号=" & mintNum
    Else
        arrSQL(UBound(arrSQL)) = "Insert Into zlFunctions(系统,函数号,函数名,中文名,说明) Values(" & _
            mlngSys & "," & mintNum & ",'ZL_FUN_" & txt函数名.Text & "'," & _
            "'" & txt中文名.Text & "','" & txt说明.Text & "')"
    End If
    
    If mobjPars.Count > 0 Then
        For i = 1 To mobjPars.Count
            With mobjPars(i)
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Insert Into zlFuncPars(系统,函数号,参数号,参数名,中文名,类型,缺省值," & _
                    "格式,值列表,分类SQL,明细SQL,分类字段,明细字段,对象,组名) Values(" & mlngSys & "," & mintNum & "," & _
                    i & ",'" & .名称 & "','" & .中文名 & "'," & .类型 & ",'" & .缺省值 & "'," & .格式 & "," & _
                    "'" & .值列表 & "'," & AdjustStr(.分类SQL) & "," & AdjustStr(.明细SQL) & ",'" & .分类字段 & "','" & .明细字段 & "'," & _
                    "'" & .对象 & "','" & .组名 & "')"
            End With
        Next
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        gcnOracle.Execute arrSQL(i)
    Next
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    With frmMain.lvw
        If mblnModi Then
            Set objItem = .SelectedItem
            objItem.Text = "ZL_FUN_" & txt函数名.Text
            objItem.SubItems(2) = txt中文名.Text
            objItem.SubItems(3) = txt说明.Text
        Else
            Set objItem = .ListItems.Add(, "_" & mintNum, "ZL_FUN_" & txt函数名.Text, 1, 1)
            objItem.SubItems(1) = mintNum
            objItem.SubItems(2) = txt中文名.Text
            objItem.SubItems(3) = txt说明.Text
        End If
        objItem.Selected = True
        .SelectedItem.EnsureVisible
    End With
        
    SaveFunc = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Function

Private Function FillUserFunc() As Boolean
    Dim i As Integer, strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    
    lvwFunc.ListItems.Clear
    
    strSQL = "Select 编号 From zlSystems Where 所有者=(Select 所有者 From zlSystems Where 编号=" & mlngSys & ")"
    
    strSQL = "Select * From User_Objects" & _
        " Where Object_Type='FUNCTION' And Object_Name Not IN(" & _
        " Select Upper(函数名) From zlFunctions Where 系统 IN(" & strSQL & "))" & _
        " And Object_Name Like 'ZL_FUN_%'" & _
        " Order by Object_Name"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "FillUserFunc")
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            lvwFunc.ListItems.Add , , rsTmp!Object_Name, , 1
            rsTmp.MoveNext
        Next
    End If
    
    FillUserFunc = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
