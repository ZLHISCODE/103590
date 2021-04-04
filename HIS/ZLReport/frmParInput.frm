VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmParInput 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   Icon            =   "frmParInput.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   6435
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1365
      Width           =   6435
      Begin VB.CheckBox chkAutoSave 
         Caption         =   "保存结束时间"
         Height          =   270
         Left            =   195
         TabIndex        =   11
         Top             =   220
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.CommandButton cmdDefault 
         Caption         =   "条件(&D)"
         Height          =   350
         Left            =   1755
         TabIndex        =   12
         Top             =   180
         Width           =   1100
      End
      Begin VB.Frame Frame1 
         Height          =   75
         Left            =   -150
         TabIndex        =   14
         Top             =   -45
         Width           =   7290
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   4035
         TabIndex        =   9
         Top             =   180
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   5160
         TabIndex        =   10
         Top             =   180
         Width           =   1100
      End
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   6435
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6435
      Begin VB.Frame Frame2 
         Height          =   45
         Left            =   -45
         TabIndex        =   2
         Top             =   570
         Width           =   7000
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "请在各个报表条件中输入或选择你本次查询所需要的条件值！"
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   1080
         TabIndex        =   1
         Top             =   120
         UseMnemonic     =   0   'False
         Width           =   4020
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   390
         Picture         =   "frmParInput.frx":014A
         Top             =   75
         Width           =   480
      End
   End
   Begin VB.PictureBox picPar 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   6435
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   630
      Width           =   6435
      Begin VB.CommandButton cmdSelNone 
         Cancel          =   -1  'True
         Caption         =   "全清"
         Height          =   350
         Left            =   5685
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.CommandButton cmdSelAll 
         Caption         =   "全选"
         Height          =   350
         Left            =   5040
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Frame fraGroup 
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   0
         Left            =   1125
         TabIndex        =   17
         Top             =   -60
         Visible         =   0   'False
         Width           =   3825
      End
      Begin VB.Frame fra 
         ForeColor       =   &H00800000&
         Height          =   645
         Index           =   0
         Left            =   1125
         TabIndex        =   15
         Top             =   60
         Visible         =   0   'False
         Width           =   3825
         Begin VB.OptionButton opt 
            Caption         =   "#"
            Height          =   180
            Index           =   0
            Left            =   105
            MaskColor       =   &H8000000F&
            TabIndex        =   16
            Top             =   270
            Visible         =   0   'False
            Width           =   1150
         End
      End
      Begin VB.CommandButton cmd 
         Caption         =   "…"
         Height          =   240
         Index           =   0
         Left            =   4425
         TabIndex        =   5
         ToolTipText     =   "按 F2 打开选择器"
         Top             =   225
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   2250
         TabIndex        =   4
         Top             =   195
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   2250
         TabIndex        =   6
         Top             =   195
         Visible         =   0   'False
         Width           =   2460
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   2250
         TabIndex        =   7
         Top             =   195
         Visible         =   0   'False
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   12946264
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
         Format          =   416743427
         CurrentDate     =   36731
      End
      Begin VB.CheckBox chk 
         Caption         =   "#"
         Height          =   195
         Index           =   0
         Left            =   2250
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参数名称"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1470
         TabIndex        =   3
         Top             =   255
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.Menu PopMenu 
      Caption         =   "弹出菜单(&P)"
      Visible         =   0   'False
      Begin VB.Menu PopMenu_Cond 
         Caption         =   "条件1"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu split0 
         Caption         =   "-"
      End
      Begin VB.Menu PopMenu_Save 
         Caption         =   "保存(&S)"
      End
      Begin VB.Menu PopMenu_Saveas 
         Caption         =   "另存为(&A)..."
      End
      Begin VB.Menu PopMenu_Del 
         Caption         =   "删除(&C)"
         Enabled         =   0   'False
      End
      Begin VB.Menu split1 
         Caption         =   "-"
      End
      Begin VB.Menu PopMenu_Default 
         Caption         =   "缺省(&D)"
      End
   End
End
Attribute VB_Name = "frmParInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mblnReset As Boolean 'I:是否用户选择重置条件而进入
Public mstrTitle As String 'I:窗体标题
Public mobjPars As RPTPars  'IO:按名称唯一的参数对象
Public mobjDefPars As RPTPars '当前报表原始的参数内容,用于恢复缺省值
Public mlngReport As Long   '报表ID
Public mblnOK As Boolean
Public mobjRPTDatas As RPTDatas

Private mint条件号 As Integer '报表条件号
Private mintMenu As Integer   '当前选择的条件菜单的索引
Private mblnMatch As Boolean
Private mintBegin As Integer
Private mintEnd As Integer

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim LngIdx As Long
    
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{Tab}": Exit Sub
    If InStr("~`!@#$^&"";|'", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If mobjPars("_" & lbl(Index).ToolTipText).类型 = 1 Then
        If InStr("-0.123456789" & Chr(8) & Chr(3) & Chr(22) & Chr(24) & Chr(26), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    End If
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{Tab}": Exit Sub
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim tmpPar As RPTPar, str明细对象 As String, str分类对象 As String
    Dim frmNewSelect As New frmSelect
    Dim strSQL明细 As String, strSQL分类 As String
    Dim colValue As New Collection    '参数现有的值
    
    For Each tmpPar In mobjPars
        If tmpPar.名称 = lbl(Index).ToolTipText Then
            If mblnMatch And txt(Index).Tag = "" Then frmNewSelect.strMatch = txt(Index).Text
            
            If InStr(tmpPar.对象, "|") > 0 Then
                str明细对象 = Split(tmpPar.对象, "|")(0)
                str分类对象 = Split(tmpPar.对象, "|")(1)
            End If
            strSQL明细 = tmpPar.明细SQL
            strSQL分类 = tmpPar.分类SQL
            Set colValue = GetValues
            Call CheckParsRela(strSQL明细, Nothing, tmpPar.名称, True, colValue, mobjPars)
            Call CheckParsRela(strSQL分类, Nothing, tmpPar.名称, True, colValue, mobjPars)
            frmNewSelect.strSQLList = SQLOwner(RemoveNote(strSQL明细), str明细对象)
            frmNewSelect.strSQLTree = SQLOwner(RemoveNote(strSQL分类), str分类对象)
            frmNewSelect.strFLDList = tmpPar.明细字段
            frmNewSelect.strFLDTree = tmpPar.分类字段
            frmNewSelect.strParName = tmpPar.名称
            frmNewSelect.bytType = tmpPar.类型
            frmNewSelect.mblnMulti = tmpPar.格式 = 1
            frmNewSelect.lngSeekHwnd = cmd(Index).hwnd
            frmNewSelect.mintConnect = GetDBConnectNo(tmpPar, mobjRPTDatas)
            
            On Error Resume Next
            Err.Clear
            
            frmNewSelect.Show 1, Me
            If frmNewSelect.mblnOK Then
                txt(Index).Text = frmNewSelect.strOutDisp
                txt(Index).Tag = frmNewSelect.strOutBand
                Unload frmNewSelect
                
                SendKeys "{Tab}"
                mblnOK = False '恢复进入时的状态
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
    Unload Me
End Sub

Private Sub cmdDefault_Click()
    Call PopupMenu(PopMenu, , cmdDefault.Left, picCmd.Top + cmdDefault.Top + cmdDefault.Height)
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer, j As Integer
    Dim strTmp As String, strDisp As String
    Dim strParName As String, curDate As Date
    
    '先检查合法性
    For i = 1 To lbl.UBound
        strParName = lbl(i).ToolTipText
        
        If mobjPars("_" & strParName).缺省值 = "固定值列表…" Then
            Select Case mobjPars("_" & strParName).格式
                Case 0
                    If Trim(cbo(i).Text) = "" Then
                        MsgBox "请选择""" & strParName & """的条件值！", vbInformation, App.Title
                        If cbo(i).Enabled Then cbo(i).SetFocus
                        Exit Sub
                    End If
                    If GetCboIndex(cbo(i), cbo(i).Text) = -1 Then '是否人为输入
                        '类型检查
                        Select Case mobjPars("_" & strParName).类型
                            Case 1
                                If Not IsNumeric(cbo(i).Text) Then
                                    MsgBox "你输入的""" & strParName & """的条件值类型应该为数字型！", vbInformation, App.Title
                                    If cbo(i).Enabled Then cbo(i).SetFocus
                                    Exit Sub
                                End If
                            Case 2
                                If Not IsDate(cbo(i).Text) Then
                                    MsgBox "你输入的""" & strParName & """的条件值类型应该为日期型！", vbInformation, App.Title
                                    If cbo(i).Enabled Then cbo(i).SetFocus
                                    Exit Sub
                                End If
                        End Select
                    End If
            End Select
        ElseIf mobjPars("_" & strParName).缺省值 = "选择器定义…" Then
            If Trim(txt(i).Text) = "" Then
                MsgBox "请选择""" & strParName & """的条件值！", vbInformation, App.Title
                If txt(i).Enabled Then txt(i).SetFocus
                Exit Sub
            End If
            If txt(i).Tag = "" Then '是否人为输入
                If mobjPars("_" & strParName).值列表 Like "*|*" Then
                    If Split(mobjPars("_" & strParName).值列表, "|")(0) <> txt(i).Text Then
                        '类型检查
                        Select Case mobjPars("_" & strParName).类型
                            Case 1
                                If Not IsNumeric(txt(i).Text) Then
                                    MsgBox "你输入的""" & strParName & """的条件值类型应该为数字型！", vbInformation, App.Title
                                    If txt(i).Enabled Then txt(i).SetFocus
                                    Exit Sub
                                End If
                            Case 2
                                If Not IsDate(txt(i).Text) Then
                                    MsgBox "你输入的""" & strParName & """的条件值类型应该为日期型！", vbInformation, App.Title
                                    If txt(i).Enabled Then txt(i).SetFocus
                                    Exit Sub
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
                                If txt(i).Enabled Then txt(i).SetFocus
                                Exit Sub
                            End If
                        Case 2
                            If Not IsDate(txt(i).Text) Then
                                MsgBox "你输入的""" & strParName & """的条件值类型应该为日期型！", vbInformation, App.Title
                                If txt(i).Enabled Then txt(i).SetFocus
                                Exit Sub
                            End If
                    End Select
                End If
            End If
        Else
            Select Case mobjPars("_" & strParName).类型
                Case 0, 3
                    If Trim(txt(i).Text) = "" Then
                        MsgBox "请输入""" & strParName & """的条件值！", vbInformation, App.Title
                        If txt(i).Enabled Then txt(i).SetFocus
                        Exit Sub
                    End If
                    If TLen(txt(i).Text) > 4000 Then
                        MsgBox """" & strParName & """的条件值长度不能超过4000个字符！", vbInformation, App.Title
                        If txt(i).Enabled Then txt(i).SetFocus
                        Exit Sub
                    End If
                Case 1
                    If Trim(txt(i).Text) = "" Then
                        MsgBox "请输入""" & strParName & """的条件值！", vbInformation, App.Title
                        If txt(i).Enabled Then txt(i).SetFocus
                        Exit Sub
                    End If
                    If TLen(txt(i).Text) > 4000 Then
                        MsgBox """" & strParName & """的条件值长度不能超过4000个字符！", vbInformation, App.Title
                        If txt(i).Enabled Then txt(i).SetFocus
                        Exit Sub
                    End If
                    If Not IsNumeric(txt(i).Text) Then
                        MsgBox """" & strParName & """的条件值类型应该为数字型！", vbInformation, App.Title
                        If txt(i).Enabled Then txt(i).SetFocus
                        Exit Sub
                    End If
                Case 2 '日期时间最大值检查
                    curDate = Currentdate
                    If Not (mobjPars("_" & strParName).缺省值 Like "&下一*" Or mobjPars("_" & strParName).Reserve Like "&下一*" Or _
                        mobjPars("_" & strParName).缺省值 Like "&后一*" Or mobjPars("_" & strParName).Reserve Like "&后一*" Or _
                        mobjPars("_" & strParName).缺省值 Like "&*结束*" Or mobjPars("_" & strParName).Reserve Like "&*结束*" Or _
                        mobjPars("_" & strParName).缺省值 Like "&*月末*" Or mobjPars("_" & strParName).缺省值 Like "&*年末*" Or _
                        mobjPars("_" & strParName).Reserve Like "&*月末*" Or mobjPars("_" & strParName).Reserve Like "&*年末*") Then
                        
                        If mobjPars("_" & strParName).缺省值 Like "*时间*" Or mobjPars("_" & strParName).Reserve Like "*时间*" Then
                            If Format(dtp(i).Value, "yyyy-MM-dd HH:mm:ss") > Format(curDate, "yyyy-MM-dd HH:mm:ss") Then
                                MsgBox """" & strParName & """ 的条件值不能超过当前时间！", vbInformation, App.Title
                                If dtp(i).Enabled Then dtp(i).SetFocus
                                Exit Sub
                            End If
                        Else
                            If Format(dtp(i).Value, "yyyy-MM-dd") > Format(curDate, "yyyy-MM-dd") Then
                                MsgBox """" & strParName & """ 的条件值不能超过当前日期！", vbInformation, App.Title
                                If dtp(i).Enabled Then dtp(i).SetFocus
                                Exit Sub
                            End If
                        End If
                    End If
            End Select
        End If
    Next
        
    '再取值
    For i = 1 To lbl.UBound
        strParName = lbl(i).ToolTipText
        
        If mobjPars("_" & strParName).缺省值 = "固定值列表…" Then
            Select Case mobjPars("_" & strParName).格式
                Case 0
                    If GetCboIndex(cbo(i), cbo(i).Text) = -1 Then '是否人为输入
                        'Reserve字段保存本次条件的"宏条件值|显示值"
                        mobjPars("_" & strParName).Reserve = "固定值列表…|" & cbo(i).Text
                        mobjPars("_" & strParName).缺省值 = cbo(i).Text
                    Else
                        '列表选择
                        'Reserve字段保存本次条件的"宏条件值|显示值"
                        '不好的分隔符
                        mobjPars("_" & strParName).Reserve = "固定值列表…|" & cbo(i).Text
                        strTmp = mobjPars("_" & strParName).值列表
                        For j = 0 To UBound(Split(strTmp, "|"))
                            strDisp = Split(Split(strTmp, "|")(j), ",")(0)
                            If Left(strDisp, 1) = "√" Then strDisp = Mid(strDisp, 2)
                            If strDisp = cbo(i).Text Then
                                mobjPars("_" & strParName).缺省值 = Split(Split(strTmp, "|")(j), ",")(1)
                                Exit For
                            End If
                        Next
                    End If
                Case 1
                    For j = 1 To opt.UBound
                        If opt(j).Container.Index = i Then
                            If opt(j).Value Then
                                'Reserve字段保存本次条件的"宏条件值|显示值"
                                mobjPars("_" & strParName).Reserve = "固定值列表…|" & opt(j).ToolTipText
                                mobjPars("_" & strParName).缺省值 = opt(j).Tag
                            End If
                        End If
                    Next
                Case 2
                    'Reserve字段保存本次条件的"宏条件值|显示值"
                    '不好的分隔符
                    strTmp = mobjPars("_" & strParName).值列表
                    For j = 0 To 1
                        strDisp = Split(Split(strTmp, "|")(j), ",")(0)
                        If chk(i).Value = 0 Then
                            If Left(strDisp, 1) <> "√" Then
                                mobjPars("_" & strParName).Reserve = "固定值列表…|" & strDisp
                                mobjPars("_" & strParName).缺省值 = Split(Split(strTmp, "|")(j), ",")(1)
                            End If
                        Else
                            If Left(strDisp, 1) = "√" Then
                                mobjPars("_" & strParName).Reserve = "固定值列表…|" & Mid(strDisp, 2)
                                mobjPars("_" & strParName).缺省值 = Split(Split(strTmp, "|")(j), ",")(1)
                            End If
                        End If
                    Next
            End Select
        ElseIf mobjPars("_" & strParName).缺省值 = "选择器定义…" Then
            If txt(i).Tag = "" Then '是否人为输入
                'Reserve字段保存本次条件的"宏条件值|显示值"
                mobjPars("_" & strParName).Reserve = "选择器定义…|"
                mobjPars("_" & strParName).缺省值 = txt(i).Text
            Else
                '列表选择
                'Reserve字段保存本次条件的"宏条件值|显示值"
                mobjPars("_" & strParName).Reserve = "选择器定义…|" & txt(i).Text
                mobjPars("_" & strParName).缺省值 = txt(i).Tag
            End If
        Else
            Select Case mobjPars("_" & strParName).类型
                Case 0, 1, 3
                    mobjPars("_" & strParName).缺省值 = txt(i).Text
                Case 2
                    If mobjPars("_" & strParName).缺省值 Like "&*" Then
                        mobjPars("_" & strParName).Reserve = mobjPars("_" & strParName).缺省值
                    End If
                    mobjPars("_" & strParName).缺省值 = Format(dtp(i).Value, dtp(i).CustomFormat)
                    '保存到注册表
                    If dtp(i).CustomFormat Like "*HH:mm:ss" Then
                        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mstrTitle, lbl(i).ToolTipText & "时间", Format(dtp(i).Value, "HH:mm:ss")
                    End If
            End Select
        End If
    Next
    
    '保存相对开始结束时间(不管勾没有)
    If mintBegin <> -1 And mintEnd <> -1 Then
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mstrTitle, "AutoSave", chkAutoSave.Value
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mstrTitle, "BeginTime", Format(dtp(mintBegin).Value, dtp(mintBegin).CustomFormat)
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mstrTitle, "EndTime", Format(dtp(mintEnd).Value, dtp(mintEnd).CustomFormat)
    End If
    
    mblnOK = True
    Hide
End Sub

Private Function GetValues() As Collection
'功能：获取现有的界面上的参数值
    Dim i As Integer, j As Integer
    Dim strParName As String, strTmp As String
    Dim strDisp As String, colValue As New Collection
     
    For i = 1 To lbl.UBound
        strParName = lbl(i).ToolTipText
        
        If mobjPars("_" & strParName).缺省值 = "固定值列表…" Then
            Select Case mobjPars("_" & strParName).格式
                Case 0
                    If GetCboIndex(cbo(i), cbo(i).Text) = -1 Then '是否人为输入
                        'Reserve字段保存本次条件的"宏条件值|显示值"
                        colValue.Add cbo(i).Text, "_" & strParName
                    Else
                        '列表选择
                        'Reserve字段保存本次条件的"宏条件值|显示值"
                        '不好的分隔符
                        strTmp = mobjPars("_" & strParName).值列表
                        For j = 0 To UBound(Split(strTmp, "|"))
                            strDisp = Split(Split(strTmp, "|")(j), ",")(0)
                            If Left(strDisp, 1) = "√" Then strDisp = Mid(strDisp, 2)
                            If strDisp = cbo(i).Text Then
                                colValue.Add Split(Split(strTmp, "|")(j), ",")(1), "_" & strParName
                                Exit For
                            End If
                        Next
                    End If
                Case 1
                    For j = 1 To opt.UBound
                        If opt(j).Container.Index = i Then
                            If opt(j).Value Then
                                colValue.Add opt(j).Tag, "_" & strParName
                            End If
                        End If
                    Next
                Case 2
                    'Reserve字段保存本次条件的"宏条件值|显示值"
                    '不好的分隔符
                    strTmp = mobjPars("_" & strParName).值列表
                    For j = 0 To 1
                        strDisp = Split(Split(strTmp, "|")(j), ",")(0)
                        If chk(i).Value = 0 Then
                            If Left(strDisp, 1) <> "√" Then
                                colValue.Add Split(Split(strTmp, "|")(j), ",")(1), "_" & strParName
                            End If
                        Else
                            If Left(strDisp, 1) = "√" Then
                                colValue.Add Split(Split(strTmp, "|")(j), ",")(1), "_" & strParName
                            End If
                        End If
                    Next
            End Select
        ElseIf mobjPars("_" & strParName).缺省值 = "选择器定义…" Then
            If txt(i).Tag = "" Then '是否人为输入
                'Reserve字段保存本次条件的"宏条件值|显示值"
                colValue.Add txt(i).Text, "_" & strParName
            Else
                '列表选择
                'Reserve字段保存本次条件的"宏条件值|显示值"
                colValue.Add txt(i).Tag, "_" & strParName
            End If
        Else
            Select Case mobjPars("_" & strParName).类型
                Case 0, 1, 3
                    colValue.Add txt(i).Text, "_" & strParName
                Case 2
                    strTmp = dtp(i).CustomFormat
                    If strTmp Like "* *:*:*" Then
                        colValue.Add Format(dtp(i).Value, "YYYY-MM-DD hh:mm:ss"), "_" & strParName
                    Else
                        colValue.Add Format(dtp(i).Value, "YYYY-MM-DD"), "_" & strParName
                    End If
            End Select
        End If
    Next
    Set GetValues = colValue
End Function

Private Sub cmdSelAll_Click()
    Dim chkTmp As CheckBox
    
    For Each chkTmp In chk
        If chkTmp.Enabled Then
            chkTmp.Value = 1
        End If
    Next
End Sub

Private Sub cmdSelNone_Click()
    Dim chkTmp As CheckBox
    
    For Each chkTmp In chk
        If chkTmp.Enabled Then
            chkTmp.Value = 0
        End If
    Next
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub dtp_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnReset = False
    Set mobjPars = Nothing
    Me.Tag = ""
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

Private Sub PopMenu_Cond_Click(Index As Integer)
    Dim objCondPars As New RPTPars
    Dim i As Integer
    
    '执行LoadCond后Index会变为0，所以用i
    i = Index
    Set objCondPars = mdlPublic.RPTParsCondExec(mlngReport, Val(PopMenu_Cond(i).Tag), mobjDefPars)
    If Not objCondPars Is Nothing Then
        Call LoadCond(objCondPars)
        mint条件号 = Val(PopMenu_Cond(i).Tag)
        mintMenu = i
        Call UpdateMenuItemCheck(mintMenu)
    End If
End Sub

Private Sub UpdateMenuItemCheck(ByVal vCondNo As Integer)
    Dim i As Integer
    
    For i = 1 To PopMenu_Cond.count - 1
        PopMenu_Cond(i).Checked = vCondNo = i
    Next
    PopMenu_Default.Checked = vCondNo = 0
End Sub

Private Sub PopMenu_Default_Click()
    mint条件号 = 0
    mintMenu = 0
    Call LoadCond(mobjDefPars)
    PopMenu_Del.Enabled = False
    Call UpdateMenuItemCheck(mintMenu)
End Sub

Private Sub LoadCond(ByVal objPars As RPTPars)
    Me.Tag = "1"
    LockWindowUpdate Me.hwnd
    Call CopyPars(objPars, mobjPars)
    Call Form_Load
    cmdOK.SetFocus
    LockWindowUpdate 0
    
    PopMenu_Del.Enabled = True
End Sub

Private Sub PopMenu_Del_Click()
    If mdlPublic.RPTParsCondDel(mlngReport, mint条件号) Then
        mint条件号 = 0
        mintMenu = 0
        Call LoadCondsMenu
        Call PopMenu_Default_Click
    End If
End Sub

Private Sub PopMenu_Save_Click()
    If mdlPublic.RPTParsCondSave(mlngReport, mint条件号, mobjPars, mobjDefPars, Me) Then
        If mintMenu = 0 Then
            '从缺省状态下保存，更新为新增的条件
            Call PopMenu_Cond_Click(PopMenu_Cond.count - 1)
        Else
            '从条件状态下保存
            Call PopMenu_Cond_Click(mintMenu)
        End If
    End If
End Sub

Private Sub PopMenu_Saveas_Click()
    If mdlPublic.RPTParsCondSave(mlngReport, mint条件号, mobjPars, mobjDefPars, Me, True) Then
        If mintMenu = 0 Then
            '从缺省状态下保存，更新为新增的条件
            Call PopMenu_Cond_Click(PopMenu_Cond.count - 1)
        Else
            '从条件状态下保存
            Call PopMenu_Cond_Click(mintMenu)
        End If
    End If
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
    If KeyCode >= 48 And KeyCode <> 144 _
        And Not (KeyCode >= 112 And KeyCode <= 123) Then
        txt(Index).Tag = ""
    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txt(Index).Tag = "" And txt(Index).ToolTipText <> "" Then
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
    
    If InStr("~`!@#$^&"";|" & Chr(3) & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    If txt(Index).ToolTipText = "" And mobjPars("_" & lbl(Index).ToolTipText).类型 = 1 Then
        If InStr("-0.123456789" & Chr(8) & Chr(24) & Chr(26), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    '人为输入时(不选择)，清除绑定值作为人为输入的标志
    '这里只处理汉字,其它在KeyDown中处理
    If KeyAscii < 0 Then txt(Index).Tag = ""
End Sub

Private Sub Form_Load()
    Dim i As Long, j As Long, k As Long
    Dim tmpPar As RPTPar, strTmp As String
    Dim lngCurH As Long, objTmp As Object
    Dim intCurTab As Integer, blnCmd As Boolean
    Dim strGroup As String, objGroup As Object
    Dim strCur As String, strPre As String
    Dim objLoad As Object, blnExist As Boolean
    Dim strBegin As String, strEnd As String
    Dim blnFlag As Boolean
    
    mblnOK = False
    mblnMatch = False
    mintBegin = -1: mintEnd = -1
    Caption = "设置条件 - " & mstrTitle
    mint条件号 = 0: mintMenu = 0
    
    '卸载控件
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
        
    Call LoadCondsMenu
    If Me.Tag = "" Then
        Call UpdateMenuItemCheck(0)
    End If
    
    '产生参数输入框组
    i = 0: lngCurH = lbl(0).Top
    For Each tmpPar In mobjPars
        i = i + 1
        
        Load lbl(i)
        lbl(i).Caption = tmpPar.名称 & "(&" & i & ")"
        lbl(i).ToolTipText = tmpPar.名称
        lbl(i).Left = txt(0).Left - lbl(i).Width - 30
        lbl(i).Top = lngCurH
        lbl(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
        lbl(i).Visible = True
        
        If tmpPar.缺省值 = "固定值列表…" Then
            If tmpPar.格式 = 0 Then '下拉框
                Load cbo(i): Set objTmp = cbo(i)
                If tmpPar.是否锁定 Then objTmp.Enabled = False
                cbo(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                cbo(i).Left = cbo(0).Left: cbo(i).Top = lbl(i).Top - (cbo(i).Height - lbl(i).Height) / 2
                '不好的分隔符
                For j = 0 To UBound(Split(tmpPar.值列表, "|"))
                    strTmp = Split(Split(tmpPar.值列表, "|")(j), ",")(0)
                    
                    If Left(strTmp, 1) = "√" Then
                        cbo(i).AddItem Mid(strTmp, 2)
                        If cbo(i).ListIndex = -1 Then cbo(i).ListIndex = cbo(i).NewIndex
                    Else
                        cbo(i).AddItem strTmp
                    End If
                    
                    '重置条件时Reserve存放了"显示值|绑定值"
                    '根据上次显示值来定位缺省项
                    If tmpPar.Reserve Like "*|*" Then
                        If Split(tmpPar.Reserve, "|")(0) = "程序传入" Then
                            '处理程序只传入了绑定值,未传入显示值,自动寻找显示值的情况
                            If Split(tmpPar.Reserve, "|")(1) = Split(Split(tmpPar.值列表, "|")(j), ",")(1) Then
                                cbo(i).ListIndex = cbo(i).NewIndex
                            End If
                        Else
                            If Left(strTmp, 1) = "√" Then
                                If Split(tmpPar.Reserve, "|")(0) = Mid(strTmp, 2) Then cbo(i).ListIndex = cbo(i).NewIndex
                            Else
                                If Split(tmpPar.Reserve, "|")(0) = strTmp Then cbo(i).ListIndex = cbo(i).NewIndex
                            End If
                            
                            '上次人为输入的值与某个绑定值相同,则定位
                            '因为多个选择值中绑定值可能重复,所以此段可不要
                            If Split(tmpPar.Reserve, "|")(0) = Split(Split(tmpPar.值列表, "|")(j), ",")(1) Then
                                cbo(i).ListIndex = cbo(i).NewIndex
                            End If
                        End If
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
                '不好的分隔符
                For j = 0 To UBound(Split(tmpPar.值列表, "|"))
                    strTmp = Split(Split(tmpPar.值列表, "|")(j), ",")(0)
                    
                    Load opt(opt.UBound + 1)
                    Set opt(opt.UBound).Container = fra(i)
                    opt(opt.UBound).TabIndex = intCurTab: intCurTab = intCurTab + 1
                    opt(opt.UBound).Tag = Split(Split(tmpPar.值列表, "|")(j), ",")(1) '存放绑定值
                    If tmpPar.是否锁定 Then opt(opt.UBound).Enabled = False
                    
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
                        opt(opt.UBound).Left = opt(0).Left + (IIF(((j + 1) Mod 3) = 0, 3, ((j + 1) Mod 3)) - 1) * (opt(0).Width + 60)
                        
                        If Left(strTmp, 1) = "√" Then
                            opt(opt.UBound).Caption = GetLenStr(Mid(strTmp, 2), opt(0).Width - 200, Me)
                            opt(opt.UBound).ToolTipText = Mid(strTmp, 2)
                            If Not blnExist Then opt(opt.UBound).Value = True
                        Else
                            opt(opt.UBound).Caption = GetLenStr(strTmp, opt(0).Width - 200, Me)
                            opt(opt.UBound).ToolTipText = strTmp
                        End If
                    End If

                    opt(opt.UBound).Width = TextWidth(opt(opt.UBound).Caption) + 300
                    
                    '重置条件时Reserve存放了"显示值|绑定值"
                    '根据上次选择值来定位缺省项
                    If tmpPar.Reserve Like "*|*" Then
                        If Split(tmpPar.Reserve, "|")(0) = "程序传入" Then
                            '处理程序只传入了绑定值,未传入显示值,自动寻找显示值的情况
                            If Split(tmpPar.Reserve, "|")(1) = Split(Split(tmpPar.值列表, "|")(j), ",")(1) Then
                                opt(opt.UBound).Value = True: blnExist = True
                            End If
                        Else
                            If Left(strTmp, 1) = "√" Then
                                If Split(tmpPar.Reserve, "|")(0) = Mid(strTmp, 2) Then
                                    opt(opt.UBound).Value = True: blnExist = True
                                End If
                            Else
                                If Split(tmpPar.Reserve, "|")(0) = strTmp Then
                                    opt(opt.UBound).Value = True: blnExist = True
                                End If
                            End If
                        End If
                    End If
                    
                    opt(opt.UBound).Visible = True
                Next
                
                fra(i).ZOrder 1 '放在最下面
                fra(i).Visible = True
            ElseIf tmpPar.格式 = 2 Then '单个复选框
                
                lbl(i).Visible = False
                If cmdSelAll.Tag = "" Then cmdSelAll.Top = lbl(i).Top: cmdSelNone.Top = lbl(i).Top
                cmdSelAll.Visible = True: cmdSelNone.Visible = True: cmdSelAll.Tag = "1"
                Load chk(i): Set objTmp = chk(i)
                If tmpPar.是否锁定 Then objTmp.Enabled = False
                chk(i).Caption = lbl(i).Caption
                chk(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                chk(i).Left = chk(0).Left: chk(i).Top = lbl(i).Top - (chk(i).Height - lbl(i).Height) / 2
                chk(i).Width = TextWidth(chk(i).Caption) + 230
                If tmpPar.组名 <> "" Then
                    If k > 0 Then
                        If fra(0).Width + fra(0).Left - chk(k).Left - chk(k).Width > (fra(0).Width - 1550) And chk(i).Width < (fra(0).Width - 1550) Then
                            chk(i).Left = fra(0).Left + 1550
                            blnFlag = True
                        ElseIf fra(0).Width + fra(0).Left - chk(k).Left - chk(k).Width > (fra(0).Width - 2800) And chk(i).Width < (fra(0).Width - 2800) Then
                            chk(i).Left = fra(0).Left + 2800
                            blnFlag = True
                        Else
                            chk(i).Left = fra(0).Left + 300
                        End If
                    Else
                        chk(i).Left = fra(0).Left + 300
                    End If
                    k = i
                End If
                
                If Left(Split(Split(tmpPar.值列表, "|")(0), ",")(0), 1) = "√" Then chk(i).Value = 1
                '不好的分隔符
                For j = 0 To 1
                    strTmp = Split(Split(tmpPar.值列表, "|")(j), ",")(0)
                    '重置条件时Reserve存放上次了"显示值|绑定值"
                    '根据上次选择值来定位本次缺省项
                    If tmpPar.Reserve Like "*|*" Then
                        If Split(tmpPar.Reserve, "|")(0) = "程序传入" Then
                            '处理程序只传入了绑定值,未传入显示值,自动寻找显示值的情况
                            If Split(tmpPar.Reserve, "|")(1) = Split(Split(tmpPar.值列表, "|")(j), ",")(1) Then
                                chk(i).Value = IIF(Left(strTmp, 1) = "√", 1, 0)
                            End If
                        Else
                            If Left(strTmp, 1) = "√" Then
                                If Split(tmpPar.Reserve, "|")(0) = Mid(strTmp, 2) Then
                                    chk(i).Value = IIF(Left(strTmp, 1) = "√", 1, 0)
                                End If
                            Else
                                If Split(tmpPar.Reserve, "|")(0) = strTmp Then
                                    chk(i).Value = IIF(Left(strTmp, 1) = "√", 1, 0)
                                End If
                            End If
                        End If
                    End If
                Next
                chk(i).Visible = True
            End If
        ElseIf tmpPar.缺省值 = "选择器定义…" Then
            Load txt(i): Set objTmp = txt(i)
            If tmpPar.是否锁定 Then objTmp.Enabled = False
            txt(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
            txt(i).Left = txt(0).Left: txt(i).Top = lbl(i).Top - (txt(i).Height - lbl(i).Height) / 2
            txt(i).ToolTipText = "按 F2 打开选择器"
            
            blnCmd = True
            If tmpPar.Reserve Like "*|*" Then
                If Split(tmpPar.Reserve, "|")(0) <> "" Then
                    strTmp = ""
                    
                    '处理程序只传入了绑定值,未传入显示值,自动寻找显示值的情况
                    If Split(tmpPar.Reserve, "|")(0) = "程序传入" Then
                        If InStr(tmpPar.对象, "|") > 0 Then strTmp = Split(tmpPar.对象, "|")(0)
                        strTmp = SQLOwner(Replace(RemoveNote(tmpPar.明细SQL), "[*]", ""), strTmp)
                        Call CheckParsRela(strTmp, Nothing, tmpPar.名称, True, , mobjPars)
                        strTmp = GetDefaultValue(strTmp, tmpPar.明细字段, Split(tmpPar.Reserve, "|")(1) _
                            , GetDBConnectNo(tmpPar, mobjRPTDatas))
                        If strTmp <> "" Then
                            tmpPar.Reserve = Split(strTmp, "|")(0) & "|" & Split(tmpPar.Reserve, "|")(1)
                        ElseIf tmpPar.值列表 Like "*|*" Then
                            If Split(tmpPar.Reserve, "|")(1) = Split(tmpPar.值列表, "|")(1) Then
                                tmpPar.Reserve = tmpPar.值列表 '与定义的缺省绑定值相同
                            End If
                        End If
                    End If

                    '重置条件时Reserve存放了"显示值|绑定值"
                    If Split(tmpPar.Reserve, "|")(0) = "程序传入" Then
                        '没有找到且与缺省绑定值不同,则显示为绑定值
                        txt(i).Text = Split(tmpPar.Reserve, "|")(1)
                    Else
                        txt(i).Text = Split(tmpPar.Reserve, "|")(0)
                    End If
                    txt(i).Tag = Split(tmpPar.Reserve, "|")(1)
                    
                    '虽然有缺省,但如果没有其它可选则不可见
                    If strTmp = "" Then '利用前面可能求"程序传入"显示值时的结果
                        If InStr(tmpPar.对象, "|") > 0 Then strTmp = Split(tmpPar.对象, "|")(0)
                        strTmp = SQLOwner(Replace(RemoveNote(tmpPar.明细SQL), "[*]", ""), strTmp)
                        Call CheckParsRela(strTmp, Nothing, tmpPar.名称, True, , mobjPars)
                        strTmp = GetDefaultValue(strTmp, tmpPar.明细字段, , GetDBConnectNo(tmpPar, mobjRPTDatas))
                    End If
                    
                    If strTmp <> "" Then
                        blnCmd = (CLng((Split(strTmp, "|")(2))) > 0)
                    Else
                        blnCmd = False
                    End If
                Else
                    '使用缺省定义的缺省值
                    If tmpPar.值列表 Like "*|*" Then
                        txt(i).Text = Split(tmpPar.值列表, "|")(0)
                        txt(i).Tag = Split(tmpPar.值列表, "|")(1)
                    ElseIf tmpPar.明细SQL <> "" Then
                        '取明细SQL结果中第一行值,如果只有一行,则不用选
                        strTmp = ""
                        If InStr(tmpPar.对象, "|") > 0 Then strTmp = Split(tmpPar.对象, "|")(0)
                        strTmp = SQLOwner(Replace(RemoveNote(tmpPar.明细SQL), "[*]", ""), strTmp)
                        Call CheckParsRela(strTmp, Nothing, tmpPar.名称, True, , mobjPars)
                        strTmp = GetDefaultValue(strTmp, tmpPar.明细字段, , GetDBConnectNo(tmpPar, mobjRPTDatas))
                        If strTmp <> "" Then
                            txt(i).Text = Split(strTmp, "|")(0)
                            txt(i).Tag = Split(strTmp, "|")(1)
                            If tmpPar.格式 = 1 Then txt(i).Tag = " IN (" & txt(i).Tag & ") "
                            blnCmd = (CLng((Split(strTmp, "|")(2))) > 1)
                        Else
                            blnCmd = False
                        End If
                    End If
                End If
            Else
                If tmpPar.值列表 Like "*|*" Then
                    '使用缺省定义的缺省值
                    txt(i).Text = Split(tmpPar.值列表, "|")(0)
                    txt(i).Tag = Split(tmpPar.值列表, "|")(1)
                    
                    '虽然有缺省,但如果没有其它可选则不可见
                    strTmp = ""
                    If InStr(tmpPar.对象, "|") > 0 Then strTmp = Split(tmpPar.对象, "|")(0)
                    strTmp = SQLOwner(Replace(RemoveNote(tmpPar.明细SQL), "[*]", ""), strTmp)
                    Call CheckParsRela(strTmp, Nothing, tmpPar.名称, True, , mobjPars)
                    strTmp = GetDefaultValue(strTmp, tmpPar.明细字段, , GetDBConnectNo(tmpPar, mobjRPTDatas))
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
                    Call CheckParsRela(strTmp, Nothing, tmpPar.名称, True, , mobjPars)
                    strTmp = GetDefaultValue(strTmp, tmpPar.明细字段, , GetDBConnectNo(tmpPar, mobjRPTDatas))
                    If strTmp <> "" Then
                        txt(i).Text = Split(strTmp, "|")(0)
                        txt(i).Tag = Split(strTmp, "|")(1)
                        If tmpPar.格式 = 1 Then txt(i).Tag = " IN (" & txt(i).Tag & ") "
                        blnCmd = (CLng((Split(strTmp, "|")(2))) > 1)
                    Else
                        blnCmd = False
                    End If
                End If
            End If
                        
            Load cmd(i)
            If tmpPar.是否锁定 Then cmd(i).Enabled = False
            cmd(i).Top = txt(i).Top
            cmd(i).Left = txt(i).Left + txt(i).Width + 15
            cmd(i).Height = txt(i).Height
            cmd(i).TabStop = False
            cmd(i).ZOrder
            
            txt(i).Visible = True
            cmd(i).Visible = blnCmd
            
            '可否输入匹配
            txt(i).Locked = Not ((InStr(tmpPar.分类SQL, "[*]") > 0 Or InStr(tmpPar.明细SQL, "[*]") > 0) And blnCmd)
        Else
            If tmpPar.类型 = 2 Then
                Load dtp(i): Set objTmp = dtp(i)
                If tmpPar.是否锁定 Then objTmp.Enabled = False
                dtp(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                dtp(i).Left = dtp(0).Left: dtp(i).Top = lbl(i).Top - (dtp(i).Height - lbl(i).Height) / 2
                If InStr(tmpPar.缺省值, ":") > 0 Or InStr(tmpPar.缺省值, "时间") > 0 Then
                    dtp(i).CustomFormat = "yyyy年MM月dd日 HH:mm:ss"
                    dtp(i).Width = 2460
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
                    dtp(i).Value = Currentdate
                End If
                
'                '注册表保存值
'                If dtp(i).CustomFormat Like "*HH:mm:ss" And Left(tmpPar.缺省值, 1) <> "&" Then
'                    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mstrTitle, lbl(i).ToolTipText & "时间", Format(dtp(i).Value, "HH:mm:ss"))
'                    dtp(i).Value = CDate(Format(dtp(i).Value, Left(dtp(i).CustomFormat, InStr(dtp(i).CustomFormat, "HH:mm:ss") - 1)) & strTmp)
'                End If
                
                '是否开始结束时间(日报有时点)
                If dtp(i).CustomFormat Like "*HH:mm:ss" Then
                    If tmpPar.名称 Like "开始*" Or tmpPar.名称 Like "起始*" Then
                        mintBegin = i
                    ElseIf tmpPar.名称 Like "结束*" Or tmpPar.名称 Like "终止*" Then
                        mintEnd = i
                    End If
                End If
                
                dtp(i).Visible = True
            Else
                Load txt(i): Set objTmp = txt(i)
                If tmpPar.是否锁定 Then objTmp.Enabled = False
                txt(i).Left = txt(0).Left: txt(i).Top = lbl(i).Top - (txt(i).Height - lbl(i).Height) / 2
                txt(i).TabIndex = intCurTab: intCurTab = intCurTab + 1
                txt(i).Text = tmpPar.缺省值
                txt(i).Visible = True
            End If
        End If
        If objTmp.name = "fra" Then
            lngCurH = lngCurH + objTmp.Height + 180
        Else
            If blnFlag = False Then
                lngCurH = lngCurH + txt(0).Height + 150
            End If
            blnFlag = False
        End If
        
        lbl(i).Tag = tmpPar.组名 & "," & objTmp.name
        If tmpPar.缺省值 = "选择器定义…" Then lbl(i).Tag = lbl(i).Tag & ",cmd"
    Next
    cmdOK.TabIndex = intCurTab: intCurTab = intCurTab + 1
    cmdCancel.TabIndex = cmdOK.TabIndex + 1
    
    picPar.Height = lngCurH
    picPar.Visible = Not (lbl.UBound = 0)
    
    k = 0
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
                    k = i
            End Select
            
            lngCurH = 195 '当前Top位置
            
            Set objTmp.Container = objGroup
            objTmp.Top = lngCurH + (300 - objTmp.Height) / 2
            If CStr(Split(lbl(i).Tag, ",")(1)) = "chk" Then
                objTmp.Left = 300
            Else
                objTmp.Left = 1250
            End If
            
            Set lbl(i).Container = objGroup
            lbl(i).Top = objTmp.Top + (objTmp.Height - lbl(i).Height) / 2
            lbl(i).Left = objTmp.Left - lbl(i).Width - 30
            lbl(i).Caption = GetLenStr(lbl(i).ToolTipText, 900, Me) & Mid(lbl(i).Caption, InStr(lbl(i).Caption, "("))
            
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
            '如果为chk，则先判断是否一行能容纳控件
            If CStr(Split(lbl(i).Tag, ",")(1)) = "chk" Then
                If objGroup.Width - chk(k).Left - chk(k).Width >= (objGroup.Width - 1550) And chk(i).Width < (objGroup.Width - 1550) Then
                    chk(i).Left = 1550
                    blnFlag = True
                ElseIf objGroup.Width - chk(k).Left - chk(k).Width > (objGroup.Width - 2800) And chk(i).Width < (objGroup.Width - 2800) Then
                    chk(i).Left = 2800
                    blnFlag = True
                Else
                    chk(i).Left = 300
                End If
            Else
                objTmp.Left = 1250
            End If
            
            If CStr(Split(lbl(i).Tag, ",")(1)) = "chk" Then
                If objGroup.Width - chk(k).Left - chk(k).Width >= (objGroup.Width - 1550) And chk(i).Width < (objGroup.Width - 1550) Then
                    objTmp.Top = chk(k).Top
                    blnFlag = True
                ElseIf objGroup.Width - chk(k).Left - chk(k).Width > (objGroup.Width - 2800) And chk(i).Width < (objGroup.Width - 2800) Then
                    objTmp.Top = chk(k).Top
                    blnFlag = True
                Else
                    objTmp.Top = lngCurH + (300 - objTmp.Height) / 2
                End If
            Else
                objTmp.Top = lngCurH + (300 - objTmp.Height) / 2
            End If
            If CStr(Split(lbl(i).Tag, ",")(1)) = "chk" Then
                k = i
            End If
            Set lbl(i).Container = objGroup
            lbl(i).Top = objTmp.Top + (objTmp.Height - lbl(i).Height) / 2
            lbl(i).Left = objTmp.Left - lbl(i).Width - 30
            lbl(i).Caption = GetLenStr(lbl(i).ToolTipText, 900, Me) & Mid(lbl(i).Caption, InStr(lbl(i).Caption, "("))
            
            If UBound(Split(lbl(i).Tag, ",")) = 2 Then
                Set cmd(i).Container = objGroup
                cmd(i).Top = objTmp.Top + 30
                cmd(i).Left = objTmp.Left + objTmp.Width - cmd(i).Width - 30
            End If
            
            If blnFlag = False Then
                lngCurH = lngCurH + txt(0).Height + 50 '当前Top位置
            End If
            
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
        blnFlag = False
    Next
    
    '没有参数组但有多项单选框时,向该框对齐
    If fraGroup.UBound = 0 And fra.UBound > 0 Then
        For Each objTmp In fra
            objTmp.Left = txt(0).Left - 400
        Next
    End If
            
    Me.Height = picInfo.Height + IIF(lbl.UBound = 0, 0, picPar.Height) + picCmd.Height + 380
    
    '处理开始结束时间
    If mintBegin <> -1 And mintEnd <> -1 Then
        chkAutoSave.Visible = True
        chkAutoSave.Value = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mstrTitle, "AutoSave", 0))
        
        '用户重新重置条件或在恢复缺省值时不处理
        If Not (mblnReset Or Visible) Then
            strBegin = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mstrTitle, "BeginTime", "")
            strEnd = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.name & mstrTitle, "EndTime", "")
                    
            '如果上次选择了保存
            If chkAutoSave.Value = 1 And IsDate(strBegin) And IsDate(strEnd) Then
                '将上次的结束时间作为本次开始时间(+1s)
                dtp(mintBegin).Value = Format(DateAdd("s", 1, CDate(strEnd)), dtp(mintBegin).CustomFormat)
            End If
        End If
    Else
        chkAutoSave.Visible = False
        cmdDefault.Left = chkAutoSave.Left
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
    If txt(Index).Tag = "" And txt(Index).ToolTipText <> "" Then
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

Private Sub LoadCondsMenu()
    Dim strSQL As String
    Dim i As Integer
    Dim rsPara As New ADODB.Recordset
    Dim blnRetry As Boolean
    
    If mlngReport = 0 Then Exit Sub
    
    On Error GoTo hErr
    
    '删除条件菜单
    For i = 1 To PopMenu_Cond.count - 1
        Unload PopMenu_Cond(i)
    Next
    
    '先装入用户已设定的缺省条件
    blnRetry = True
    strSQL = "Select Distinct 条件号,条件名称 From zlRptConds Where 报表ID=[1] Order by 条件号"
    Set rsPara = OpenSQLRecord(strSQL, Me.Caption, mlngReport)
    blnRetry = False
    
    With rsPara
        If .RecordCount = 0 Then
            Me.split0.Visible = False
            If mlngReport = 0 Then
                PopMenu_Save.Enabled = False
                PopMenu_Saveas.Enabled = False
                Me.split1.Enabled = False
            End If
        Else
            Me.split0.Visible = True
            PopMenu_Save.Enabled = True
            PopMenu_Saveas.Enabled = True
            Me.split1.Enabled = True
            Do While Not .EOF
                i = .AbsolutePosition
                Load PopMenu_Cond(i)
                PopMenu_Cond(i).Caption = !条件名称 & "(&" & i & ")"
                PopMenu_Cond(i).Visible = True
                PopMenu_Cond(i).Tag = !条件号
                .MoveNext
            Loop
        End If
    End With
    
    Exit Sub
    
hErr:
    If blnRetry Then
        If ErrCenter = 1 Then Resume
    Else
        Call ErrCenter
    End If
End Sub
