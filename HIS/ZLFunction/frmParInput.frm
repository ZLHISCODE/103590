VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmParInput 
   AutoRedraw      =   -1  'True
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8925
   Icon            =   "frmParInput.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picHead 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   3630
      ScaleHeight     =   630
      ScaleWidth      =   5220
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   15
      Width           =   5220
      Begin VB.Frame Frame2 
         Height          =   45
         Left            =   -45
         TabIndex        =   8
         Top             =   570
         Width           =   7500
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "请在下面所列出的函数参数中输入或选择你所需要的参数值！"
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   960
         TabIndex        =   7
         Top             =   120
         UseMnemonic     =   0   'False
         Width           =   4020
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   300
         Picture         =   "frmParInput.frx":014A
         Top             =   45
         Width           =   480
      End
   End
   Begin VB.PictureBox picCmd 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   3660
      ScaleHeight     =   675
      ScaleWidth      =   5220
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4680
      Width           =   5220
      Begin VB.CommandButton cmdTest 
         Caption         =   "测试(&T)"
         Height          =   350
         Left            =   1390
         TabIndex        =   2
         Top             =   150
         Width           =   1100
      End
      Begin VB.CommandButton cmdDefault 
         Caption         =   "缺省(&D)"
         Height          =   350
         Left            =   210
         TabIndex        =   5
         Top             =   150
         Width           =   1100
      End
      Begin VB.Frame Frame1 
         Height          =   75
         Left            =   -150
         TabIndex        =   16
         Top             =   -45
         Width           =   7290
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   2570
         TabIndex        =   3
         Top             =   150
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   3750
         TabIndex        =   4
         Top             =   150
         Width           =   1100
      End
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   3630
      ScaleHeight     =   645
      ScaleWidth      =   5070
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3990
      Width           =   5070
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C00000&
         Height          =   585
         Left            =   45
         TabIndex        =   24
         Top             =   30
         Width           =   4980
      End
   End
   Begin VB.PictureBox picPar 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   3630
      ScaleHeight     =   705
      ScaleWidth      =   5220
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   660
      Width           =   5220
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   1800
         TabIndex        =   13
         Top             =   195
         Visible         =   0   'False
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   12946264
         CalendarTitleForeColor=   16777215
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
         Format          =   41418755
         CurrentDate     =   36731
      End
      Begin VB.CommandButton cmd 
         Caption         =   "…"
         Height          =   240
         Index           =   0
         Left            =   4140
         TabIndex        =   11
         ToolTipText     =   "按 F2 打开选择器"
         Top             =   225
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1800
         TabIndex        =   10
         Top             =   195
         Visible         =   0   'False
         Width           =   2640
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   1800
         TabIndex        =   12
         Top             =   195
         Visible         =   0   'False
         Width           =   2640
      End
      Begin VB.Frame fraGroup 
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   0
         Left            =   450
         TabIndex        =   20
         Top             =   -60
         Visible         =   0   'False
         Width           =   4020
      End
      Begin VB.CheckBox chk 
         Caption         =   "#"
         Height          =   195
         Index           =   0
         Left            =   1800
         TabIndex        =   17
         Top             =   255
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.Frame fra 
         ForeColor       =   &H00800000&
         Height          =   645
         Index           =   0
         Left            =   450
         TabIndex        =   18
         Top             =   60
         Visible         =   0   'False
         Width           =   4020
         Begin VB.OptionButton opt 
            Caption         =   "#"
            Height          =   180
            Index           =   0
            Left            =   105
            MaskColor       =   &H8000000F&
            TabIndex        =   19
            Top             =   270
            Visible         =   0   'False
            Width           =   1150
         End
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参数名称"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   690
         TabIndex        =   9
         Top             =   255
         Visible         =   0   'False
         Width           =   1005
      End
   End
   Begin VB.Frame fraFun 
      Caption         =   " 函数列表 "
      Height          =   5235
      Left            =   60
      TabIndex        =   21
      Top             =   30
      Width           =   3540
      Begin MSComctlLib.ImageCombo cboSys 
         Height          =   315
         Left            =   630
         TabIndex        =   0
         Top             =   270
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin MSComctlLib.ListView lvw 
         Height          =   4410
         Left            =   165
         TabIndex        =   1
         Top             =   660
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   7779
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDropMode     =   1
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "img32"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         OLEDropMode     =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "函数名"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "函数号"
            Object.Width           =   1376
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "中文名"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "说明"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "状态"
            Object.Width           =   970
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "系统"
         Height          =   180
         Left            =   165
         TabIndex        =   22
         Top             =   330
         Width           =   360
      End
   End
   Begin MSComctlLib.ImageList img32 
      Left            =   885
      Top             =   465
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParInput.frx":0454
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   315
      Top             =   465
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
            Picture         =   "frmParInput.frx":076E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmParInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'可选入口参数:如果使用,则对指定函数进行测试
Public mlngSys As Long '系统号
Public mstrOwner As String '所有者
Public mintNum As Integer '函数号
Public mstrFun As String '函数名

'入、出口参数：在函数向导中可作入
Public mstrExp As String '函数公式

Private mobjPars As FuncPars '函数参数集
Private mstrVals As String '函数参数值

Private mstrPars As String '当前函数参数描述串
Private mstrCode As String '当前函数代码

Private mblnMatch As Boolean
Private mstrKey As String
Private mstrPreFun As String

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim LngIdx As Long
    
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{Tab}": Exit Sub
    If InStr("~`!@#$^&"";|'", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If mobjPars("_" & lbl(Index).ToolTipText).类型 = 1 Then
        If InStr("-0.123456789" & Chr(8) & Chr(3) & Chr(22) & Chr(24) & Chr(26), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    End If
    
    If KeyAscii <> 8 Then
        If SendMessage(cbo(Index).hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then SendKeys "{F4}"
        LngIdx = MatchIndex(cbo(Index), KeyAscii)
        If LngIdx <> -2 Then cbo(Index).ListIndex = LngIdx
    End If
End Sub

Private Sub cboSys_Click()
    If cboSys.SelectedItem Is Nothing Then Exit Sub
    If cboSys.SelectedItem.Key = mstrKey Then Exit Sub
    mstrKey = cboSys.SelectedItem.Key
    
    mlngSys = Val(Mid(cboSys.SelectedItem.Key, 2))
    mstrOwner = cboSys.SelectedItem.Tag
    
    Call ReadFunc
    
    If Not lvw.SelectedItem Is Nothing Then
        mstrPreFun = ""
        Call lvw_ItemClick(lvw.SelectedItem)
        cmdTest.Enabled = True
        cmdOK.Enabled = True
    Else
        lblInfo.Caption = ""
        mintNum = 0
        mstrFun = ""
        mstrCode = ""
        mstrPars = ""
        Set mobjPars = New FuncPars
        Call ShowPars
        cmdTest.Enabled = False
        cmdOK.Enabled = False
    End If
End Sub

Private Sub ReadFunc()
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, objItem As ListItem
    Dim strSQL As String, j As Integer
    
    On Error GoTo errH
    
    lvw.ListItems.Clear
    If cboSys.SelectedItem Is Nothing Then Exit Sub
    
    strSQL = "Select A.*,B.STATUS From zlFunctions A,All_Objects B" & _
        " Where A.系统=" & mlngSys & " And B.Owner='" & mstrOwner & "'" & _
        " And B.Object_Type='FUNCTION' And Upper(A.函数名)=B.Object_Name" & _
        " Order by A.函数号"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "ReadFunc")
    For i = 1 To rsTmp.RecordCount
        grsObject.Filter = "OWNER='" & UCase(mstrOwner) & "' And OBJECT_TYPE='FUNCTION' And OBJECT_NAME='" & UCase(rsTmp!函数名) & "'"
        If Not grsObject.EOF Then
            Set objItem = lvw.ListItems.Add(, "_" & rsTmp!函数号, rsTmp!函数名, 1, 1)
            objItem.SubItems(1) = rsTmp!函数号
            objItem.SubItems(2) = IIf(IsNull(rsTmp!中文名), "", rsTmp!中文名)
            objItem.SubItems(3) = IIf(IsNull(rsTmp!说明), "", rsTmp!说明)
            If rsTmp!Status <> "VALID" Then
                objItem.SubItems(4) = "×"
                objItem.ForeColor = &H808080
                For j = 1 To objItem.ListSubItems.Count
                    objItem.ListSubItems(j).ForeColor = &H808080
                Next
            End If
        End If
        rsTmp.MoveNext
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboSys_GotFocus()
    SelAll cboSys
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{Tab}": Exit Sub
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim tmpPar As FuncPar, str明细对象 As String, str分类对象 As String
    
    For Each tmpPar In mobjPars
        If tmpPar.名称 = lbl(Index).ToolTipText Then
            If mblnMatch And txt(Index).Tag = "" Then frmSelect.mstrMatch = txt(Index).Text
            
            If InStr(tmpPar.对象, "|") > 0 Then
                str明细对象 = Split(tmpPar.对象, "|")(0)
                str分类对象 = Split(tmpPar.对象, "|")(1)
            End If
            frmSelect.mstrSQLList = SQLOwner(RemoveNote(tmpPar.明细SQL), str明细对象)
            frmSelect.mstrSQLTree = SQLOwner(RemoveNote(tmpPar.分类SQL), str分类对象)
            frmSelect.mstrFLDList = tmpPar.明细字段
            frmSelect.mstrFLDTree = tmpPar.分类字段
            frmSelect.mstrParName = tmpPar.中文名
            frmSelect.mbytDataType = tmpPar.类型
            
            frmSelect.mlngSeekHwnd = cmd(Index).hwnd
            
            On Error Resume Next
            Err.Clear
            
            frmSelect.Show 1, Me
            If gblnOK Then
                txt(Index).Text = frmSelect.mstrOutDisp
                txt(Index).Tag = frmSelect.mstrOutBand
                Unload frmSelect
                
                SendKeys "{Tab}"
                gblnOK = False '恢复进入时的状态
            ElseIf mblnMatch Then
                txt(Index).Text = ""
                txt(Index).Tag = ""
            End If
            
            mblnMatch = False
            Exit For
        End If
    Next
    txt(Index).SetFocus
End Sub

Private Sub cmdCancel_Click()
    mstrExp = ""
    Unload Me
End Sub

Private Sub cmdDefault_Click()
    Call ShowPars
    If cmdOK.Visible And cmdOK.Enabled Then
        cmdOK.SetFocus
    ElseIf cmdTest.Visible And cmdTest.Enabled Then
        cmdTest.SetFocus
    End If
End Sub

Private Function CheckInput() As Boolean
    Dim i As Integer, j As Integer
    Dim strParName As String, curDate As Date
    
    '先检查合法性
    For i = 1 To lbl.UBound
        strParName = lbl(i).ToolTipText
        
        If mobjPars("_" & strParName).缺省值 = "固定值列表…" Then
            Select Case mobjPars("_" & strParName).格式
                Case 0
'                    If Trim(cbo(i).Text) = "" Then
'                        MsgBox "请选择""" & strParName & """的条件值！", vbInformation, App.Title
'                        cbo(i).SetFocus: Exit Function
'                    End If
                    If Trim(cbo(i).Text) <> "" Then
                        If GetCboIndex(cbo(i), cbo(i).Text) = -1 Then '是否人为输入
                            '类型检查
                            Select Case mobjPars("_" & strParName).类型
                                Case 1
                                    If Not IsNumeric(cbo(i).Text) Then
                                        MsgBox "你输入的""" & strParName & """的条件值类型应该为数字型！", vbInformation, App.Title
                                        cbo(i).SetFocus: Exit Function
                                    End If
                                Case 2
                                    If Not IsDate(cbo(i).Text) Then
                                        MsgBox "你输入的""" & strParName & """的条件值类型应该为日期型！", vbInformation, App.Title
                                        cbo(i).SetFocus: Exit Function
                                    End If
                            End Select
                        End If
                    End If
            End Select
        ElseIf mobjPars("_" & strParName).缺省值 = "选择器定义…" Then
'            If Trim(txt(i).Text) = "" Then
'                MsgBox "请选择""" & strParName & """的条件值！", vbInformation, App.Title
'                txt(i).SetFocus: Exit Function
'            End If
            If Trim(txt(i).Text) <> "" Then
                If txt(i).Tag = "" Then '是否人为输入
                    If mobjPars("_" & strParName).值列表 Like "*|*" Then
                        If Split(mobjPars("_" & strParName).值列表, "|")(0) <> txt(i).Text Then
                            '类型检查
                            Select Case mobjPars("_" & strParName).类型
                                Case 1
                                    If Not IsNumeric(txt(i).Text) Then
                                        MsgBox "你输入的""" & strParName & """的条件值类型应该为数字型！", vbInformation, App.Title
                                        txt(i).SetFocus: Exit Function
                                    End If
                                Case 2
                                    If Not IsDate(txt(i).Text) Then
                                        MsgBox "你输入的""" & strParName & """的条件值类型应该为日期型！", vbInformation, App.Title
                                        txt(i).SetFocus: Exit Function
                                    End If
                            End Select
                        Else
                            '输入值与定义的缺省值相同,则还原为缺省值
                            txt(i).Tag = Split(mobjPars("_" & strParName).值列表, "|")(1)
                        End If
                    Else
                        '类型检查
                        Select Case mobjPars("_" & strParName).类型
                            Case 1
                                If Not IsNumeric(txt(i).Text) Then
                                    MsgBox "你输入的""" & strParName & """的条件值类型应该为数字型！", vbInformation, App.Title
                                    txt(i).SetFocus: Exit Function
                                End If
                            Case 2
                                If Not IsDate(txt(i).Text) Then
                                    MsgBox "你输入的""" & strParName & """的条件值类型应该为日期型！", vbInformation, App.Title
                                    txt(i).SetFocus: Exit Function
                                End If
                        End Select
                    End If
                End If
            End If
        Else
            Select Case mobjPars("_" & strParName).类型
                Case 0, 3
'                    If Trim(txt(i).Text) = "" Then
'                        MsgBox "请输入""" & strParName & """的条件值！", vbInformation, App.Title
'                        txt(i).SetFocus: Exit Function
'                    End If
                    If Trim(txt(i).Text) <> "" Then
                        If TLen(txt(i).Text) > 255 Then
                            MsgBox """" & strParName & """的条件值长度不能超过255个字符！", vbInformation, App.Title
                            txt(i).SetFocus: Exit Function
                        End If
                    End If
                Case 1
'                    If Trim(txt(i).Text) = "" Then
'                        MsgBox "请输入""" & strParName & """的条件值！", vbInformation, App.Title
'                        txt(i).SetFocus: Exit Function
'                    End If
                    If Trim(txt(i).Text) <> "" Then
                        If TLen(txt(i).Text) > 255 Then
                            MsgBox """" & strParName & """的条件值长度不能超过255个字符！", vbInformation, App.Title
                            txt(i).SetFocus: Exit Function
                        End If
                        If Not IsNumeric(txt(i).Text) Then
                            MsgBox """" & strParName & """的条件值类型应该为数字型！", vbInformation, App.Title
                            txt(i).SetFocus: Exit Function
                        End If
                    End If
                Case 2 '日期时间最大值检查
                    If Not IsNull(dtp(i).Value) Then
                        curDate = Currentdate
                        If Not (mobjPars("_" & strParName).缺省值 Like "&下一*" Or mobjPars("_" & strParName).缺省值 Like "&后一*" Or _
                            mobjPars("_" & strParName).缺省值 Like "&*结束*" Or mobjPars("_" & strParName).缺省值 Like "&*月末*" Or _
                            mobjPars("_" & strParName).缺省值 Like "&*年末*") Then
                            
                            If mobjPars("_" & strParName).缺省值 Like "*时间*" Then
                                If Format(dtp(i).Value, "yyyy-MM-dd HH:mm:ss") > Format(curDate, "yyyy-MM-dd HH:mm:ss") Then
                                    MsgBox """" & strParName & """ 的条件值不能超过当前时间！", vbInformation, App.Title
                                    dtp(i).SetFocus: Exit Function
                                End If
                            Else
                                If Format(dtp(i).Value, "yyyy-MM-dd") > Format(curDate, "yyyy-MM-dd") Then
                                    MsgBox """" & strParName & """ 的条件值不能超过当前日期！", vbInformation, App.Title
                                    dtp(i).SetFocus: Exit Function
                                End If
                            End If
                        End If
                    End If
            End Select
        End If
    Next
    
    CheckInput = True
End Function

Private Function GetInput() As FuncPars
    Dim i As Integer, j As Integer
    Dim strTmp As String, strDisp As String
    Dim strParName As String, objPars As FuncPars

    Call CopyPars(mobjPars, objPars)
    
    '再取值
    For i = 1 To lbl.UBound
        strParName = lbl(i).ToolTipText
        
        If objPars("_" & strParName).缺省值 = "固定值列表…" Then
            Select Case objPars("_" & strParName).格式
                Case 0
                    If GetCboIndex(cbo(i), cbo(i).Text) = -1 Then '是否人为输入
                        objPars("_" & strParName).缺省值 = cbo(i).Text
                    Else
                        '列表选择
                        strTmp = objPars("_" & strParName).值列表
                        For j = 0 To UBound(Split(strTmp, "|"))
                            strDisp = Split(Split(strTmp, "|")(j), ",")(0)
                            If Left(strDisp, 1) = "√" Then strDisp = Mid(strDisp, 2)
                            If strDisp = cbo(i).Text Then
                                objPars("_" & strParName).缺省值 = Split(Split(strTmp, "|")(j), ",")(1)
                                Exit For
                            End If
                        Next
                    End If
                Case 1
                    For j = 1 To opt.UBound
                        If opt(j).Container.Index = i Then
                            If opt(j).Value Then
                                objPars("_" & strParName).缺省值 = opt(j).Tag
                            End If
                        End If
                    Next
                Case 2
                    strTmp = objPars("_" & strParName).值列表
                    For j = 0 To 1
                        strDisp = Split(Split(strTmp, "|")(j), ",")(0)
                        If chk(i).Value = 0 Then
                            If Left(strDisp, 1) <> "√" Then
                                objPars("_" & strParName).缺省值 = Split(Split(strTmp, "|")(j), ",")(1)
                            End If
                        Else
                            If Left(strDisp, 1) = "√" Then
                                objPars("_" & strParName).缺省值 = Split(Split(strTmp, "|")(j), ",")(1)
                            End If
                        End If
                    Next
            End Select
        ElseIf objPars("_" & strParName).缺省值 = "选择器定义…" Then
            If txt(i).Tag = "" Then '是否人为输入
                If InStr(objPars("_" & strParName).值列表, "|") > 0 Then
                    If txt(i).Text = Split(objPars("_" & strParName).值列表, "|")(0) Then
                        objPars("_" & strParName).缺省值 = Split(objPars("_" & strParName).值列表, "|")(1)
                    Else
                        objPars("_" & strParName).缺省值 = txt(i).Text
                    End If
                Else
                    objPars("_" & strParName).缺省值 = txt(i).Text
                End If
            Else
                '列表选择
                objPars("_" & strParName).缺省值 = txt(i).Tag
            End If
        Else
            Select Case objPars("_" & strParName).类型
                Case 0, 1, 3
                    objPars("_" & strParName).缺省值 = txt(i).Text
                Case 2
                    If IsNull(dtp(i).Value) Then
                        objPars("_" & strParName).缺省值 = ""
                    Else
                        objPars("_" & strParName).缺省值 = Format(dtp(i).Value, dtp(i).CustomFormat)
                    End If
            End Select
        End If
    Next
    Set GetInput = objPars
End Function

Private Sub cmdOK_Click()
    Dim objPars As FuncPars
    
    If Not CheckInput Then Exit Sub
    Set objPars = GetInput
    
    mstrExp = GetFunctionExp(mstrOwner, mstrFun, mstrPars, objPars)
    
    gblnOK = True
    Unload Me
End Sub

Private Sub cmdTest_Click()
    Dim strReturn As String
    Dim objPars As FuncPars
    
    If Not CheckInput Then Exit Sub
    Set objPars = GetInput
    
    strReturn = ExeFunction(mstrOwner, mstrFun, mstrPars, objPars)
    If strReturn Like "ERROR*" Then
        If Not picInfo.Visible Then
            MsgBox "执行失败：" & vbCrLf & vbCrLf & Mid(strReturn, 6), vbInformation, App.Title
        Else
            lblInfo.ForeColor = &HC0
            lblInfo.Caption = "注意:" & vbCrLf & Mid(strReturn, 6)
        End If
    Else
        If Not picInfo.Visible Then
            MsgBox "执行成功：" & vbCrLf & vbCrLf & mstrFun & " = " & strReturn & "    ", vbInformation, App.Title
        Else
            lblInfo.ForeColor = &HC00000
            lblInfo.Caption = "提示:执行成功," & mstrFun & " = " & strReturn
        End If
        VBA.Beep
    End If
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub dtp_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    fraFun.Left = 60
    fraFun.Top = 90
    
    picHead.Top = 0
    picHead.Left = Me.ScaleWidth - picHead.Width
    
    picCmd.Top = Me.ScaleHeight - picCmd.Height
    picCmd.Left = picHead.Left
    
    picPar.Top = picHead.Top + picHead.Height
    picPar.Left = picHead.Left
    
    picInfo.Left = picHead.Left + 30
    picInfo.Top = picCmd.Top - picInfo.Height - 30
    picInfo.Width = picHead.Width - 45
    
    fraFun.Width = Me.ScaleWidth - picHead.Width - fraFun.Left
    fraFun.Height = Me.ScaleHeight - fraFun.Top - fraFun.Left
    
    lvw.Width = fraFun.Width - lvw.Left * 2
    cboSys.Width = lvw.Width - (cboSys.Left - lvw.Left)
    lvw.Height = fraFun.Height - lvw.Top - 150
    
    RaisEffect picInfo, -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjPars = Nothing
    mlngSys = 0
    mstrOwner = ""
    mintNum = 0
    mstrFun = ""
    mstrCode = ""
    mstrPars = ""
    mstrVals = ""
    
    If Not InDesign And glngOldProc <> 0 Then Call SetWindowLong(hwnd, GWL_WNDPROC, glngOldProc)
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim strObject As String, i As Integer
    
    If mstrPreFun = Item.Key Then Exit Sub
    mstrPreFun = Item.Key
    
    mintNum = Val(Mid(Item.Key, 2))
    mstrFun = lvw.SelectedItem.Text
    
    '先显示空的参数
    mstrCode = ""
    mstrPars = ""
    Set mobjPars = New FuncPars
    Call ShowPars
    
    mstrCode = GetFunSource(mstrOwner, mstrFun)
    mstrPars = GetFuncPars(mstrCode)
    Set mobjPars = ReadFuncPars(mlngSys, mintNum)
    Call ReplaceSysNum
    
    '各种检查
    Me.Refresh
    '当前所有者函数本身是否有Execute权限(所有者或DBA一定具有)
    grsObject.Filter = "OWNER='" & UCase(mstrOwner) & "' And OBJECT_TYPE='FUNCTION' And OBJECT_NAME='" & UCase(mstrFun) & "'"
    If grsObject.EOF Then
        lblInfo.ForeColor = &HC0&
        lblInfo.Caption = "注意:当前用户没有权限执行该函数！"
        'cmdTest.Enabled = False: cmdOK.Enabled = False
        Exit Sub
    End If
    '函数代码是否存在决定是否有权限执行
    If mstrCode = "" Then
        lblInfo.ForeColor = &HC0&
        lblInfo.Caption = "注意:不能读取函数代码,你可能没有权限执行该函数！"
        'cmdTest.Enabled = False: cmdOK.Enabled = False
        Exit Sub
    End If
    '是否定义了参数值
    If mstrPars <> "" And mobjPars.Count = 0 Then
        lblInfo.ForeColor = &HC0&
        lblInfo.Caption = "注意:函数代码中定义了参数,但没有定义这些参数的取值方法！"
        'cmdTest.Enabled = False: cmdOK.Enabled = False
        Exit Sub
    End If
    '函数参数选择器中的对象是否有权限
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
        lblInfo.ForeColor = &HC0&
        lblInfo.Caption = "注意:当前用户不具有下列对象或没有权限访问这些对象:" & vbCrLf & strObject
        'cmdTest.Enabled = False: cmdOK.Enabled = False
        Exit Sub
    End If
    
    lblInfo.ForeColor = &HC00000
    lblInfo.Caption = Item.SubItems(2) & "(" & mstrFun & "):" & Item.SubItems(3)
    
    Call ShowPars
    cmdTest.Enabled = True
    cmdOK.Enabled = True
End Sub

Private Sub opt_GotFocus(Index As Integer)
    If opt(Index).Value Then
        '这样做的目的是避免按TAB键时自动切换到下一个选项
        opt(Index).Value = False
        opt(Index).Value = True
    End If
End Sub

Private Sub opt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{Tab}": Exit Sub
End Sub

Private Sub txt_GotFocus(Index As Integer)
    SelAll txt(Index)
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 And txt(Index).ToolTipText <> "" Then
        If cmd(Index).Enabled And cmd(Index).Visible Then Call cmd_Click(Index)
    End If
    If txt(Index).Locked Then Exit Sub
    
    '人为输入时(不选择)，清除绑定值作为人为输入的标志
    '144=Num;112-123=F1-F12;229=开始输入汉字
    If KeyCode >= 48 And KeyCode <> 144 And KeyCode <> 229 _
        And Not (KeyCode >= 112 And KeyCode <= 123) Then
        txt(Index).Tag = ""
    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strTmp As String
    If KeyAscii = 13 Then
        If txt(Index).Tag = "" And txt(Index).ToolTipText <> "" Then
            strTmp = mobjPars("_" & lbl(Index).ToolTipText).值列表
            If InStr(strTmp, "|") > 0 Then
                If Split(strTmp, "|")(0) = txt(Index).Text Then
                    KeyAscii = 0: SendKeys "{Tab}": Exit Sub
                End If
            End If
        
            '想输入匹配
            KeyAscii = 0
            If txt(Index).Text <> "" Then
                If cmd(Index).Enabled And cmd(Index).Visible Then
                    mblnMatch = True
                    Call cmd_Click(Index)
                End If
            End If
            Exit Sub
        Else
            '想移动焦点
            KeyAscii = 0: SendKeys "{Tab}": Exit Sub
        End If
    End If
    
    If txt(Index).Locked Then Exit Sub
    
    If InStr("~`!@#$^&"";|'" & Chr(3) & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    If txt(Index).ToolTipText = "" And mobjPars("_" & lbl(Index).ToolTipText).类型 = 1 Then
        If InStr("-0.123456789" & Chr(8) & Chr(24) & Chr(26), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    '人为输入时(不选择)，清除绑定值作为人为输入的标志
    '这里只处理汉字,其它在KeyDown中处理
    If KeyAscii < 0 Then txt(Index).Tag = ""
End Sub

Private Sub ShowPars()
'功能：显示函数的参数(控件、值)
'说明：mstrVals<>""时，顺序包含了各个参数的值
    Dim i As Integer, j As Integer
    Dim tmpPar As FuncPar, strTmp As String
    Dim lngCurH As Long, objTmp As Object
    Dim intCurTab As Integer, blnCmd As Boolean
    Dim strGroup As String, objGroup As Object
    Dim strCur As String, strPre As String
    Dim objLoad As Object, blnExist As Boolean
    Dim str对象 As String, strVal As String
    
    Screen.MousePointer = 11
    If Visible Then LockWindowUpdate Me.hwnd
    
    For Each objLoad In lbl
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In txt
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In cmd
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In cbo
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In dtp
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In opt
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In chk
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In fra
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    For Each objLoad In fraGroup
        If objLoad.Index <> 0 Then Unload objLoad
    Next
    
    '产生参数输入框组
    i = 0: lngCurH = lbl(0).Top: intCurTab = 2
    For Each tmpPar In mobjPars
        '当前值
        If UBound(Split(mstrVals, "|")) >= i Then
            strVal = Split(mstrVals, "|")(i)
            If Left(strVal, 1) = "'" And Right(strVal, 1) = "'" Then
                strVal = Mid(strVal, 2, Len(strVal) - 2)
            End If
        Else
            strVal = ""
        End If
        
        i = i + 1
        
        Load lbl(i)
        lbl(i).Caption = GetLenStr(tmpPar.中文名, 1800, Me)
        lbl(i).ToolTipText = tmpPar.名称
        lbl(i).Left = txt(0).Left - lbl(i).Width - 30
        lbl(i).Top = lngCurH
        lbl(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
        lbl(i).Visible = True
        
        If tmpPar.缺省值 = "固定值列表…" Then
            If tmpPar.格式 = 0 Then '下拉框
                Load cbo(i): Set objTmp = cbo(i)
                cbo(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                cbo(i).Left = cbo(0).Left: cbo(i).Top = lbl(i).Top - (cbo(i).Height - lbl(i).Height) / 2
                For j = 0 To UBound(Split(tmpPar.值列表, "|"))
                    strTmp = Split(Split(tmpPar.值列表, "|")(j), ",")(0)
                    
                    If Left(strTmp, 1) = "√" Then
                        cbo(i).AddItem Mid(strTmp, 2)
                        If cbo(i).ListIndex = -1 Then cbo(i).ListIndex = cbo(i).NewIndex
                    Else
                        cbo(i).AddItem strTmp
                    End If
                                        
                    '绑定值相同则赋值
                    If strVal = Split(Split(tmpPar.值列表, "|")(j), ",")(1) And strVal <> "" Then
                        cbo(i).ListIndex = cbo(i).NewIndex
                    End If
                Next
                cbo(i).Visible = True
            ElseIf tmpPar.格式 = 1 Then '单选框
                Load fra(i): Set objTmp = fra(i)
                fra(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                fra(i).Left = fra(0).Left: fra(i).Top = lbl(i).Top - 50
                
                lbl(i).Visible = False
                fra(i).Caption = lbl(i).Caption
                                
                j = UBound(Split(tmpPar.值列表, "|")) + 1 '可选数
                j = CInt((j / 3) + 0.4) '行数
                
                fra(i).Height = fra(0).Height + (j - 1) * (opt(0).Height * 1.6) - opt(0).Height * 0.3
                
                blnExist = False
                For j = 0 To UBound(Split(tmpPar.值列表, "|"))
                    strTmp = Split(Split(tmpPar.值列表, "|")(j), ",")(0)
                    
                    Load opt(opt.UBound + 1)
                    Set opt(opt.UBound).Container = fra(i)
                    opt(opt.UBound).TabIndex = intCurTab: intCurTab = intCurTab + 1
                    opt(opt.UBound).Tag = Split(Split(tmpPar.值列表, "|")(j), ",")(1) '存放绑定值
                    
                    If InStr(",0,1,3,", "," & UBound(Split(tmpPar.值列表, "|")) & ",") > 0 Then
                        '只有1,2,4个的情况特殊处理
                        If j = 0 Or j = 1 Then 'Top
                            opt(opt.UBound).Top = opt(0).Top
                        Else
                            opt(opt.UBound).Top = opt(0).Top + opt(0).Height * 1.6
                        End If
                        If j = 0 Or j = 2 Then 'Left
                            opt(opt.UBound).Left = opt(0).Left + 150
                        Else
                            opt(opt.UBound).Left = opt(0).Left + (opt(0).Width * 1.4 + 60) + 150
                        End If
                        
                        If Left(strTmp, 1) = "√" Then
                            opt(opt.UBound).Caption = GetLenStr(Mid(strTmp, 2), opt(0).Width * 1.4 - 200, Me)
                            opt(opt.UBound).ToolTipText = Mid(strTmp, 2)
                            If Not blnExist Then opt(opt.UBound).Value = True
                        Else
                            opt(opt.UBound).Caption = GetLenStr(strTmp, opt(0).Width * 1.4 - 200, Me)
                            opt(opt.UBound).ToolTipText = strTmp
                        End If
                    Else
                        opt(opt.UBound).Top = opt(0).Top + (CInt(((j + 1) / 3) + 0.4) - 1) * (opt(0).Height * 1.6)
                        opt(opt.UBound).Left = opt(0).Left + (IIf(((j + 1) Mod 3) = 0, 3, ((j + 1) Mod 3)) - 1) * (opt(0).Width + 60)
                        
                        If Left(strTmp, 1) = "√" Then
                            opt(opt.UBound).Caption = GetLenStr(Mid(strTmp, 2), opt(0).Width - 200, Me)
                            opt(opt.UBound).ToolTipText = Mid(strTmp, 2)
                            If Not blnExist Then opt(opt.UBound).Value = True
                        Else
                            opt(opt.UBound).Caption = GetLenStr(strTmp, opt(0).Width - 200, Me)
                            opt(opt.UBound).ToolTipText = strTmp
                        End If
                    End If
                    
                    '绑定值相同则赋值
                    If strVal = opt(opt.UBound).Tag And strVal <> "" Then
                        opt(opt.UBound).Value = True
                        blnExist = True
                    End If
                    
                    opt(opt.UBound).Width = TextWidth(opt(opt.UBound).Caption) + 300
                    opt(opt.UBound).Visible = True
                Next
                
                fra(i).ZOrder 1 '放在最下面
                fra(i).Visible = True
            ElseIf tmpPar.格式 = 2 Then '单个复选框
                lbl(i).Visible = False
                
                Load chk(i): Set objTmp = chk(i)
                chk(i).Caption = lbl(i).Caption
                chk(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                chk(i).Left = chk(0).Left: chk(i).Top = lbl(i).Top - (chk(i).Height - lbl(i).Height) / 2
                chk(i).Width = TextWidth(chk(i).Caption) + 300
                
                If Left(Split(Split(tmpPar.值列表, "|")(0), ",")(0), 1) = "√" Then chk(i).Value = 1
                For j = 0 To 1
                    strTmp = Split(Split(tmpPar.值列表, "|")(j), ",")(0)
                    
                    '如果绑定值相同,则赋值
                    If strVal = Split(Split(tmpPar.值列表, "|")(j), ",")(1) And strVal <> "" Then
                        If Left(Split(Split(tmpPar.值列表, "|")(j), ",")(0), 1) = "√" Then
                            chk(i).Value = 1
                        Else
                            chk(i).Value = 0
                        End If
                    End If
                Next
                chk(i).Visible = True
            End If
        ElseIf tmpPar.缺省值 = "选择器定义…" Then
            Load txt(i): Set objTmp = txt(i)
            txt(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
            txt(i).Left = txt(0).Left: txt(i).Top = lbl(i).Top - (txt(i).Height - lbl(i).Height) / 2
            txt(i).ToolTipText = "按 F2 打开选择器"
            
            blnCmd = True
            If tmpPar.值列表 Like "*|*" Then
                '使用缺省定义的缺省值
                txt(i).Text = Split(tmpPar.值列表, "|")(0)
                txt(i).Tag = Split(tmpPar.值列表, "|")(1)
                
                '虽然有缺省,但如果没有其它可选则不可见
                strTmp = ""
                If InStr(tmpPar.对象, "|") > 0 Then strTmp = Split(tmpPar.对象, "|")(0)
                strTmp = SQLOwner(Replace(RemoveNote(tmpPar.明细SQL), "[*]", ""), strTmp)
                strTmp = GetDefaultValue(strTmp, tmpPar.明细字段)
                If strTmp <> "" Then
                    blnCmd = (CLng((Split(strTmp, "|")(2))) > 0)
                Else
                    blnCmd = False
                End If
            ElseIf tmpPar.明细SQL <> "" Then
                '取明细SQL结果中第一行值,如果只有一行,则不用选
                strTmp = ""
                If InStr(tmpPar.对象, "|") > 0 Then strTmp = Split(tmpPar.对象, "|")(0)
                strTmp = SQLOwner(Replace(RemoveNote(tmpPar.明细SQL), "[*]", ""), strTmp)
                strTmp = GetDefaultValue(strTmp, tmpPar.明细字段)
                If strTmp <> "" Then
                    txt(i).Text = Split(strTmp, "|")(0)
                    txt(i).Tag = Split(strTmp, "|")(1)
                    blnCmd = (CLng((Split(strTmp, "|")(2))) > 1)
                Else
                    blnCmd = False
                End If
            End If
                                    
            '根据绑定值赋值
            If strVal <> "" Then
                strTmp = ""
                If tmpPar.值列表 Like "*|*" Then
                    strTmp = Split(tmpPar.值列表, "|")(1)
                    If (strVal = strTmp) Or (UCase(strVal) = "NULL" And Trim(strTmp) = "") Then
                        txt(i).Text = Split(tmpPar.值列表, "|")(0)
                        txt(i).Tag = Split(tmpPar.值列表, "|")(1)
                        strTmp = "OK"
                    Else
                        strTmp = ""
                    End If
                End If
                
                If strTmp = "" Then
                    If InStr(tmpPar.对象, "|") > 0 Then strTmp = Split(tmpPar.对象, "|")(0)
                    strTmp = SQLOwner(Replace(RemoveNote(tmpPar.明细SQL), "[*]", ""), strTmp)
                    strTmp = GetBalndValue(strTmp, tmpPar.明细字段, strVal)
                    If strTmp <> "" Then
                        txt(i).Text = Split(strTmp, "|")(0)
                        txt(i).Tag = Split(strTmp, "|")(1)
                    End If
                End If
            End If
            
            Load cmd(i)
            cmd(i).Top = txt(i).Top + 30
            cmd(i).Left = txt(i).Left + txt(i).Width - cmd(i).Width - 30
            cmd(i).Height = txt(i).Height - 45
            cmd(i).TabStop = False
            cmd(i).ZOrder
            
            txt(i).Visible = True
            cmd(i).Visible = blnCmd
            
            '可否输入匹配
            txt(i).Locked = Not ((InStr(tmpPar.分类SQL, "[*]") > 0 Or InStr(tmpPar.明细SQL, "[*]") > 0) And blnCmd)
        Else
            If tmpPar.类型 = 2 Then
                Load dtp(i): Set objTmp = dtp(i)
                dtp(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                dtp(i).Left = dtp(0).Left: dtp(i).Top = lbl(i).Top - (dtp(i).Height - lbl(i).Height) / 2
                If InStr(tmpPar.缺省值, ":") > 0 Or InStr(tmpPar.缺省值, "时间") > 0 Then
                    dtp(i).CustomFormat = "yyyy年MM月dd日 HH:mm:ss"
                    dtp(i).Width = 2640
                Else
                    dtp(i).CustomFormat = "yyyy年MM月dd日"
                    dtp(i).Width = 1635
                End If
                If tmpPar.缺省值 <> "" Then
                    If Left(tmpPar.缺省值, 1) = "&" Then
                        dtp(i).Value = GetParVBMacro(tmpPar.缺省值)
                    Else
                        dtp(i).Value = Format(tmpPar.缺省值, dtp(i).CustomFormat)
                    End If
                Else
                    dtp(i).Value = Null
                End If
                
                '非宏日期才赋值
                If Left(tmpPar.缺省值, 1) <> "&" And strVal <> "" Then
                    If UCase(strVal) Like "TO_DATE('*','*')" Then
                        dtp(i).Value = GetDate(strVal)
                    Else
                        dtp(i).Value = Null
                    End If
                End If
                
                dtp(i).Visible = True
            Else
                Load txt(i): Set objTmp = txt(i)
                txt(i).Left = txt(0).Left: txt(i).Top = lbl(i).Top - (txt(i).Height - lbl(i).Height) / 2
                txt(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                txt(i).Text = tmpPar.缺省值
                
                '赋值
                If strVal <> "" Then txt(i).Text = strVal
                
                txt(i).Visible = True
            End If
        End If
        If objTmp.Name = "fra" Then
            lngCurH = lngCurH + objTmp.Height + 180
        Else
            lngCurH = lngCurH + txt(0).Height + 150
        End If
        
        lbl(i).Tag = tmpPar.组名 & "," & objTmp.Name
        If tmpPar.缺省值 = "选择器定义…" Then lbl(i).Tag = lbl(i).Tag & ",cmd"
    Next
    cmdOK.TabIndex = intCurTab: intCurTab = intCurTab + 1
    cmdCancel.TabIndex = cmdOK.TabIndex + 1
    
    picPar.Height = lngCurH
    picPar.Visible = Not (lbl.UBound = 0)
    
    '处理参数组
    For i = 1 To lbl.UBound
        strCur = ""
        If strGroup <> CStr(Split(lbl(i).Tag, ",")(0)) And CStr(Split(lbl(i).Tag, ",")(0)) <> "" Then
            Load fraGroup(fraGroup.UBound + 1)
            Set objGroup = fraGroup(fraGroup.UBound)
            objGroup.Caption = CStr(Split(lbl(i).Tag, ",")(0))
            objGroup.Top = lbl(i).Top - 150
            objGroup.ZOrder 1
            objGroup.Visible = True
            
            Select Case CStr(Split(lbl(i).Tag, ",")(1))
                Case "txt"
                    Set objTmp = txt(i)
                Case "cbo"
                    Set objTmp = cbo(i)
                Case "dtp"
                    Set objTmp = dtp(i)
                Case "chk"
                    Set objTmp = chk(i)
            End Select
            
            lngCurH = 195 '当前Top位置
            
            Set objTmp.Container = objGroup
            objTmp.Top = lngCurH + (300 - objTmp.Height) / 2
            objTmp.Left = 1250
            
            Set lbl(i).Container = objGroup
            lbl(i).Top = objTmp.Top + (objTmp.Height - lbl(i).Height) / 2
            lbl(i).Left = objTmp.Left - lbl(i).Width - 30
            lbl(i).Caption = GetLenStr(lbl(i).Caption, 1200, Me)
            
            If UBound(Split(lbl(i).Tag, ",")) = 2 Then
                Set cmd(i).Container = objGroup
                cmd(i).Top = objTmp.Top + 30
                cmd(i).Left = objTmp.Left + objTmp.Width - cmd(i).Width - 30
            End If

            lngCurH = lngCurH + txt(0).Height + 50 '当前Top位置
        ElseIf strGroup = CStr(Split(lbl(i).Tag, ",")(0)) And CStr(Split(lbl(i).Tag, ",")(0)) <> "" Then
            strCur = "Add"
            Select Case CStr(Split(lbl(i).Tag, ",")(1))
                Case "txt"
                    Set objTmp = txt(i)
                Case "cbo"
                    Set objTmp = cbo(i)
                Case "dtp"
                    Set objTmp = dtp(i)
                Case "chk"
                    Set objTmp = chk(i)
            End Select
            
            Set objTmp.Container = objGroup
            objTmp.Top = lngCurH + (300 - objTmp.Height) / 2
            objTmp.Left = 1250
            
            Set lbl(i).Container = objGroup
            lbl(i).Top = objTmp.Top + (objTmp.Height - lbl(i).Height) / 2
            lbl(i).Left = objTmp.Left - lbl(i).Width - 30
            lbl(i).Caption = GetLenStr(lbl(i).Caption, 1200, Me)
            
            If UBound(Split(lbl(i).Tag, ",")) = 2 Then
                Set cmd(i).Container = objGroup
                cmd(i).Top = objTmp.Top + 30
                cmd(i).Left = objTmp.Left + objTmp.Width - cmd(i).Width - 30
            End If
                        
            lngCurH = lngCurH + txt(0).Height + 50 '当前Top位置
            
            objGroup.Height = objTmp.Top + objTmp.Height + 90  '框高度
            
            '该框以下的条件输入全部下移
            For j = i + 1 To lbl.UBound
                If Split(lbl(j).Tag, ",")(0) <> "fra" Then
                    lbl(j).Top = lbl(j).Top + 60
                    Select Case CStr(Split(lbl(j).Tag, ",")(1))
                        Case "txt"
                            txt(j).Top = txt(j).Top + 60
                        Case "cbo"
                            cbo(j).Top = cbo(j).Top + 60
                        Case "dtp"
                            dtp(j).Top = dtp(j).Top + 60
                        Case "chk"
                            chk(j).Top = chk(j).Top + 60
                    End Select
                    If UBound(Split(lbl(j).Tag, ",")) = 2 Then
                        cmd(j).Top = cmd(j).Top + 60
                    End If
                End If
            Next
        End If
        If strPre = "Add" And strCur = "" Then
            picPar.Height = picPar.Height + 60
        End If
        strPre = strCur
        strGroup = CStr(Split(lbl(i).Tag, ",")(0))
    Next
    
    '没有参数组但有多项单选框时,向该框对齐
    If fraGroup.UBound = 0 And fra.UBound > 0 Then
        For Each objTmp In fra
            objTmp.Left = txt(0).Left - 615
        Next
    End If
        
    cmdTest.TabIndex = intCurTab: intCurTab = intCurTab + 1
    cmdOK.TabIndex = intCurTab: intCurTab = intCurTab + 1
    cmdCancel.TabIndex = intCurTab: intCurTab = intCurTab + 1
    cmdDefault.TabIndex = intCurTab: intCurTab = intCurTab + 1
    
    If Caption = "函数测试" Then
        Height = picHead.Height + IIf(lbl.UBound = 0, 0, picPar.Height) + picCmd.Height + 580
    End If
    
    If Visible Then LockWindowUpdate 0
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    Dim strOwner As String, strFunc As String, i As Integer
    Dim lngSys As Long
    
    gblnOK = False
    mblnMatch = False
    mstrKey = ""
    glngOldProc = 0
    mstrPreFun = ""
    
    dtp(0).Value = Currentdate
    
    If mintNum = 0 Then
        Caption = "函数选择"
        Call ReadSystem
        
        If mstrExp <> "" Then
            Call SplitFunc(mstrExp, strOwner, strFunc, mstrVals)
            
            '定位系统
            lngSys = GetFuncSys(strOwner, strFunc)
            For i = 1 To cboSys.ComboItems.Count
                If Val(Mid(cboSys.ComboItems(i).Key, 2)) = lngSys Then
                    cboSys.ComboItems(i).Selected = True: Exit For
                End If
            Next
            If i <= cboSys.ComboItems.Count Then
                Call cboSys_Click
                
                '定位函数
                For i = 1 To lvw.ListItems.Count
                    If UCase(lvw.ListItems(i).Text) = UCase(strFunc) Then
                        lvw.ListItems(i).Selected = True
                        lvw.ListItems(i).EnsureVisible
                        Exit For
                    End If
                Next
                If i <= lvw.ListItems.Count Then
                    mstrPreFun = ""
                    Call lvw_ItemClick(lvw.SelectedItem)
                End If
            End If
            
            mstrExp = "": mstrVals = ""
        End If
    Else
        Caption = "函数测试"
        
        cmdOK.Visible = False
        cmdTest.Left = cmdOK.Left
        fraFun.Visible = False
        picInfo.Visible = False
        Me.Width = picPar.Width + 120
        
        mstrCode = GetFunSource(mstrOwner, mstrFun)
        mstrPars = GetFuncPars(mstrCode)
        Set mobjPars = ReadFuncPars(mlngSys, mintNum)
        Call ReplaceSysNum
        
        If mstrPars <> "" And mobjPars.Count = 0 Then
            MsgBox "函数代码中定义了参数,但没有定义这些参数的取值方法！", vbInformation, App.Title
        End If
        
        Call ShowPars
                
        '限定窗体大小
        If Not InDesign Then
            glngMinW = Me.Width \ 15: glngMinH = Me.Height \ 15
            glngMaxW = Me.Width \ 15: glngMaxH = Me.Height \ 15
            glngOldProc = GetWindowLong(hwnd, GWL_WNDPROC)
            Call SetWindowLong(hwnd, GWL_WNDPROC, AddressOf CustomMessage)
        End If
    End If
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        lngTXTProc = GetWindowLong(txt(Index).hwnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txt(Index).hwnd, GWL_WNDPROC, lngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Dim strTmp As String
    
    If txt(Index).Tag = "" And txt(Index).ToolTipText <> "" Then
        strTmp = mobjPars("_" & lbl(Index).ToolTipText).值列表
        If InStr(strTmp, "|") > 0 Then
            If Split(strTmp, "|")(0) = txt(Index).Text Then Exit Sub
        End If
        '强行输入匹配
        If txt(Index).Text <> "" Then
            If cmd(Index).Enabled And cmd(Index).Visible Then
                mblnMatch = True
                Call cmd_Click(Index)
            End If
            Cancel = True
        End If
    End If
End Sub

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvw.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvw.SortOrder = lvwDescending
    Else
        lvw.SortOrder = lvwAscending
    End If
    lvw.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvw.SelectedItem Is Nothing Then lvw.SelectedItem.EnsureVisible
End Sub

Public Sub ReadSystem()
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, strSQL As String
    Dim objItem As ComboItem
    
    On Error GoTo errH
    
    cboSys.ComboItems.Clear
    strSQL = "Select * From zlSystems Order by 编号"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "ReadSystem")
    For i = 1 To rsTmp.RecordCount
        Set objItem = cboSys.ComboItems.Add(, "_" & rsTmp!编号, rsTmp!名称 & "-" & Right(Format(rsTmp!编号, "00000"), 2) & "(" & rsTmp!所有者 & ")")
        objItem.Tag = rsTmp!所有者
        If rsTmp!所有者 = gstrDBUser And cboSys.SelectedItem Is Nothing Then
            objItem.Selected = True
        End If
        rsTmp.MoveNext
    Next
    If Not cboSys.SelectedItem Is Nothing Then Call cboSys_Click
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ReplaceSysNum()
    Dim i As Integer
    For i = 1 To mobjPars.Count
        mobjPars(i).分类SQL = Replace(mobjPars(i).分类SQL, "[系统]", mlngSys)
        mobjPars(i).明细SQL = Replace(mobjPars(i).明细SQL, "[系统]", mlngSys)
    Next
End Sub
