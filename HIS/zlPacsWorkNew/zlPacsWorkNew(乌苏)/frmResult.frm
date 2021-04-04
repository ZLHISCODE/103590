VERSION 5.00
Begin VB.Form frmResult 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "报告结果"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10125
   ControlBox      =   0   'False
   Icon            =   "frmResult.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraCriticalValues 
      Caption         =   "危急值"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   1890
      Begin VB.OptionButton optCriticalValues 
         Caption         =   "普通"
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
         Index           =   1
         Left            =   480
         TabIndex        =   20
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton optCriticalValues 
         Caption         =   "危急"
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
         Index           =   2
         Left            =   480
         TabIndex        =   19
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.Frame fraReportLevel 
      Caption         =   "报告质量"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   6120
      TabIndex        =   11
      Top             =   120
      Width           =   1890
      Begin VB.OptionButton optReportLevel 
         Caption         =   "丁"
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
         Index           =   4
         Left            =   360
         TabIndex        =   15
         Top             =   1440
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optReportLevel 
         Caption         =   "丙"
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
         Index           =   3
         Left            =   360
         TabIndex        =   14
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optReportLevel 
         Caption         =   "甲"
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
         Index           =   1
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optReportLevel 
         Caption         =   "乙"
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
         Index           =   2
         Left            =   360
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame fraFuHeLevel 
      Caption         =   "符合情况"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   8115
      TabIndex        =   7
      Top             =   120
      Width           =   1890
      Begin VB.OptionButton optFuHeLevel 
         Caption         =   "不 符 合"
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
         Index           =   3
         Left            =   360
         TabIndex        =   10
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton optFuHeLevel 
         Caption         =   "符    合"
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
         Index           =   1
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optFuHeLevel 
         Caption         =   "基本符合"
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
         Index           =   2
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.Frame fraImageLevel 
      Caption         =   "影像质量"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   4140
      TabIndex        =   2
      Top             =   120
      Width           =   1890
      Begin VB.OptionButton optImageLevel 
         Caption         =   "丁"
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
         Index           =   4
         Left            =   360
         TabIndex        =   17
         Top             =   1500
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optImageLevel 
         Caption         =   "丙"
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
         Index           =   3
         Left            =   360
         TabIndex        =   16
         Top             =   1160
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optImageLevel 
         Caption         =   "乙"
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
         Index           =   2
         Left            =   360
         TabIndex        =   6
         Top             =   780
         Width           =   1455
      End
      Begin VB.OptionButton optImageLevel 
         Caption         =   "甲"
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
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   420
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame fraResult 
      Caption         =   "阴阳性"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   2145
      TabIndex        =   1
      Top             =   120
      Width           =   1890
      Begin VB.OptionButton optResult 
         Caption         =   "阴性"
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
         Index           =   2
         Left            =   480
         TabIndex        =   4
         Top             =   1320
         Width           =   855
      End
      Begin VB.OptionButton optResult 
         Caption         =   "阳性"
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
         Index           =   1
         Left            =   480
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   350
      Left            =   6720
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmResult.frx":000C
      TabIndex        =   0
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   1100
   End
End
Attribute VB_Name = "frmResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintResult As Integer    '检查结果
Private mintImageLevel As Integer '胶片质量
Private mintFuHeLevel As Integer  '符合情况
Private mintReportLevel As Integer '报告质量
Private mintCriticalValues As Integer '危急情况
Private mstrResult As String
Public mlngModul As Long      '模块号调用

Public Function zlGetResult(frmParent As Form, ByVal lngModul As Long, ByVal strQueryId As String, lngCur科室ID As Long, strResultInput As String) As String
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim blnShowResult As Boolean
    Dim blnShowCriticalValues As Boolean
    Dim blnShowImageLevel As Boolean
    Dim blnShowReportLevel As Boolean
    Dim blnShowFuHeLevel As Boolean
    Dim strImageLevel As String
    Dim strReportLevel As String
    Dim intTxtLen As Integer
    Dim i As Integer
    Dim lngFramCount As Long
    
    zlGetResult = ""
    mlngModul = lngModul
    
    If strQueryId Like String(Len(strQueryId), "#") Then
        strSql = "Select a.影像质量,a.符合情况,a.报告质量,b.结果阳性,a.危急状态 From 影像检查记录 a,病人医嘱发送 b  " _
                & " Where a.医嘱ID= b.医嘱ID And a.发送号 =b.发送号 And  a.医嘱ID=[1] "
    Else
        strSql = "Select B.危急状态, A.结果阳性, B.影像质量, A.报告质量, B.符合情况,B.医嘱ID " & _
                 "From 影像报告记录 A, 影像检查记录 B " & _
                 "Where A.ID=[1] and A.医嘱id = B.医嘱id"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取报告结果", strQueryId)
    
    '设置已经选择的阳性结果
    If Nvl(rsTemp!结果阳性) = 1 Then
        optResult(1).value = True   '阳性按钮
        mintResult = 1
    Else
        optResult(2).value = True   '阴性按钮
        mintResult = 2
    End If
    
    
    '如果记录为空则使用默认值
    If Nvl(rsTemp!结果阳性) = "" Then
         '如果勾选参数 诊断结果默认阳性 自动选择阳性按钮 反之选择阴性按钮
        If Val(GetDeptPara(lngCur科室ID, "诊断结果默认阳性", 0)) = 1 Then
            optResult(1).value = True
            mintResult = 1
        Else
            optResult(2).value = True
            mintResult = 2
        End If
    End If
    
    lngFramCount = 1
    
    blnShowResult = True
    blnShowCriticalValues = True
    blnShowImageLevel = True
    blnShowReportLevel = True
    blnShowFuHeLevel = True
    
    If InStr(strResultInput, "危急状态") > 0 Then
        blnShowCriticalValues = True
        
        If Nvl(rsTemp!危急状态) = 1 Then
            optCriticalValues(2).value = True   '危急按钮
            mintCriticalValues = 2
        Else
            optCriticalValues(1).value = True   '正常按钮
            mintCriticalValues = 1
        End If
        
        If mintCriticalValues = 2 Then
            '如果为危急状态为'危急',则标记为'阳性'，且不可更改
            optResult(1).value = True
            optResult(1).Enabled = False
            optResult(2).Enabled = False
        End If
        
        lngFramCount = lngFramCount + 1
    Else
        blnShowCriticalValues = False
        fraCriticalValues.Visible = False
    End If
    
    
    
    If InStr(strResultInput, "结果阳性") > 0 Then
        blnShowResult = True
        lngFramCount = lngFramCount + 1
    Else
        blnShowResult = False
        fraResult.Visible = False
    End If
    
    
    
    '影像质量区域
    strImageLevel = Nvl(GetDeptPara(lngCur科室ID, "影像质量等级", "甲,乙"))
    intTxtLen = Len(strImageLevel) - Len(Replace(strImageLevel, ",", "")) + 1
    
    If mlngModul = 1290 Then
        If InStr(strResultInput, "影像质量") <= 0 Then        '不显示
            fraImageLevel.Visible = False
            blnShowImageLevel = False
            
            If IsNull(rsTemp!影像质量) Then mintImageLevel = 0
        Else
            lngFramCount = lngFramCount + 1
            mintImageLevel = 1
        End If
    Else
        fraImageLevel.Visible = False
        blnShowImageLevel = False
        If IsNull(rsTemp!影像质量) Then mintImageLevel = 0
    End If
    
    '固定最多是4个影像等级  所以循环4次
    For i = 1 To 4
        If i <= intTxtLen Then
            optImageLevel(i).Visible = True
            
            If Trim(Split(strImageLevel, ",")(i - 1)) <> "" Then
                optImageLevel(i).Caption = Trim(Split(strImageLevel, ",")(i - 1))
            Else
                optImageLevel(i).Caption = "未设置"
            End If
            
            If Nvl(rsTemp!影像质量) = i Then
                optImageLevel(i).value = True
                mintImageLevel = i
            End If

        Else
            optImageLevel(i).Visible = False
        End If
    Next i
    
    '通过设置的等级个数来判断top 的值
    Select Case intTxtLen
        Case 2
            optImageLevel(1).Top = 600
            optImageLevel(2).Top = 1320
        Case 3
            optImageLevel(1).Top = 480
            optImageLevel(2).Top = 960
            optImageLevel(3).Top = 1440
        Case 4
            optImageLevel(1).Top = 420
            optImageLevel(2).Top = 780
            optImageLevel(3).Top = 1160
            optImageLevel(4).Top = 1500
    End Select
    
    
    
     '报告质量区域
    strReportLevel = Nvl(GetDeptPara(lngCur科室ID, "报告质量等级", "甲,乙"))
    intTxtLen = Len(strReportLevel) - Len(Replace(strReportLevel, ",", "")) + 1
    
    If InStr(strResultInput, "报告质量") <= 0 Then       '不显示
        fraReportLevel.Visible = False
        blnShowReportLevel = False
        
        If IsNull(rsTemp!报告质量) Then mintReportLevel = 0
    Else
        lngFramCount = lngFramCount + 1
        mintReportLevel = 1
    End If
    
    '固定最多是4个等级  所以循环4次
    For i = 1 To 4
        If i <= intTxtLen Then
            optReportLevel(i).Visible = True
            
            If Trim(Split(strReportLevel, ",")(i - 1)) <> "" Then
                optReportLevel(i).Caption = Trim(Split(strReportLevel, ",")(i - 1))
            Else
                optReportLevel(i).Caption = "未设置"
            End If
            
            If Nvl(rsTemp!报告质量) = i Then
                optReportLevel(i).value = True
                mintReportLevel = i
            End If
        Else
            optReportLevel(i).Visible = False
        End If
    Next i

    '通过设置的等级个数来判断top 的值
    Select Case intTxtLen
        Case 2
            optReportLevel(1).Top = 600
            optReportLevel(2).Top = 1320
        Case 3
            optReportLevel(1).Top = 480
            optReportLevel(2).Top = 960
            optReportLevel(3).Top = 1440
        Case 4
            optReportLevel(1).Top = 420
            optReportLevel(2).Top = 780
            optReportLevel(3).Top = 1160
            optReportLevel(4).Top = 1500
    End Select
    
    '显示符合情况
    If Nvl(rsTemp!符合情况) = "" Or Nvl(rsTemp!符合情况) = "符合" Then
        optFuHeLevel(1).value = True
        mintFuHeLevel = 1
    ElseIf Nvl(rsTemp!符合情况) = "基本符合" Then
        optFuHeLevel(2).value = True
        mintFuHeLevel = 2
    Else
        optImageLevel(3).value = True
        mintFuHeLevel = 3
    End If

    If mlngModul = 1294 Or InStr(strResultInput, "符合情况") <= 0 Then        '不显示
        fraFuHeLevel.Visible = False
        blnShowFuHeLevel = False
        mintFuHeLevel = 1
    Else
        lngFramCount = lngFramCount + 1
    End If
    
    Me.Width = IIf(blnShowResult, fraResult.Width, 0) + IIf(blnShowCriticalValues, fraCriticalValues.Width, 0) + IIf(blnShowImageLevel, fraImageLevel.Width, 0) + IIf(blnShowReportLevel, fraReportLevel.Width, 0) + IIf(blnShowFuHeLevel, fraFuHeLevel.Width, 0) + 120 + lngFramCount * 120
    
    cmdOK.Left = Me.Width - cmdOK.Width - 240
    
    If Not blnShowCriticalValues Then
        fraResult.Left = fraCriticalValues.Left
    Else
        fraResult.Left = fraCriticalValues.Left + fraCriticalValues.Width + 120
    End If
    
    If Not blnShowResult Then
        fraImageLevel.Left = fraResult.Left
    Else
        fraImageLevel.Left = fraResult.Left + fraResult.Width + 120
    End If

    If Not blnShowImageLevel Then
        fraReportLevel.Left = fraImageLevel.Left
    Else
        fraReportLevel.Left = fraImageLevel.Left + fraImageLevel.Width + 120
    End If

    If Not blnShowReportLevel Then
        fraFuHeLevel.Left = fraReportLevel.Left
    Else
        fraFuHeLevel.Left = fraReportLevel.Left + fraReportLevel.Width + 120
    End If
    
    If blnShowResult = False And blnShowCriticalValues = False And blnShowImageLevel = False And blnShowReportLevel = False And blnShowFuHeLevel = False Then
        Unload Me
        Exit Function
    End If

    Me.Show 1, frmParent
    zlGetResult = mstrResult
End Function

Private Sub cmdCancel_Click()
    mstrResult = ""
    Unload Me
End Sub

Private Sub CmdOK_Click()
    mstrResult = mintCriticalValues & "-" & mintResult & "-" & mintImageLevel & "-" & mintReportLevel & "-" & mintFuHeLevel
    Unload Me
End Sub

Private Sub Form_Load()
    '窗口置顶
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3 '将窗口置顶
End Sub

Private Sub optCriticalValues_Click(Index As Integer)
    mintCriticalValues = Index
    If mintCriticalValues = 1 Then
        optResult(1).Enabled = True
        optResult(2).Enabled = True
    Else
        '如果为危急状态为'危急',则标记为'阳性'，且不可更改
        optResult(1).value = True
        optResult(1).Enabled = False
        optResult(2).Enabled = False
    End If
End Sub

Private Sub optFuHeLevel_Click(Index As Integer)
    mintFuHeLevel = Index
End Sub

Private Sub optImageLevel_Click(Index As Integer)
    mintImageLevel = Index
End Sub

Private Sub optReportLevel_Click(Index As Integer)
    mintReportLevel = Index
End Sub

Private Sub optResult_Click(Index As Integer)
    mintResult = Index
End Sub


