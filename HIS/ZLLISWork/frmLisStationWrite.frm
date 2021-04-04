VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmLisStationWrite 
   BorderStyle     =   0  'None
   Caption         =   "普通报告填写"
   ClientHeight    =   5205
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7920
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmLisStationWrite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommDialog 
      Left            =   6000
      Top             =   4500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8595
      Begin VB.CheckBox chkYiQiTiShi 
         Appearance      =   0  'Flat
         Caption         =   "仪器审核提示"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   5580
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.CheckBox chkYiQiBiaoShi 
         Appearance      =   0  'Flat
         Caption         =   "仪器标识"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   6420
         TabIndex        =   23
         Top             =   240
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.PictureBox PicFilter 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   5160
         MouseIcon       =   "frmLisStationWrite.frx":0E42
         Picture         =   "frmLisStationWrite.frx":0F94
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   22
         Top             =   45
         Width           =   240
      End
      Begin VB.CheckBox chkOriginal 
         Appearance      =   0  'Flat
         Caption         =   "原始"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   780
         TabIndex        =   15
         Top             =   30
         Width           =   690
      End
      Begin VB.CheckBox chkLast 
         Appearance      =   0  'Flat
         Caption         =   "上次"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1470
         TabIndex        =   14
         Top             =   30
         Width           =   690
      End
      Begin VB.CheckBox chkSign 
         Appearance      =   0  'Flat
         Caption         =   "标志"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2190
         TabIndex        =   13
         Top             =   30
         Width           =   660
      End
      Begin VB.CheckBox chkUnit 
         Appearance      =   0  'Flat
         Caption         =   "单位"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2910
         TabIndex        =   12
         Top             =   30
         Width           =   660
      End
      Begin VB.CheckBox chkReferrence 
         Appearance      =   0  'Flat
         Caption         =   "参考"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3630
         TabIndex        =   11
         Top             =   30
         Width           =   660
      End
      Begin VB.CheckBox chkMB 
         Appearance      =   0  'Flat
         Caption         =   "酶标"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   4350
         TabIndex        =   10
         Top             =   30
         Width           =   660
      End
      Begin VB.CheckBox chkChina 
         Appearance      =   0  'Flat
         Caption         =   "中文"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   60
         TabIndex        =   9
         Top             =   30
         Width           =   690
      End
      Begin VB.Label lblLow 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         Height          =   210
         Left            =   5535
         TabIndex        =   21
         Top             =   45
         Width           =   285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "偏低"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   5895
         TabIndex        =   20
         Top             =   60
         Width           =   360
      End
      Begin VB.Label lblHigh 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Height          =   210
         Left            =   6315
         TabIndex        =   19
         Top             =   45
         Width           =   285
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "偏高"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   6645
         TabIndex        =   18
         Top             =   60
         Width           =   360
      End
      Begin VB.Label lblExigency 
         BackColor       =   &H000040C0&
         Height          =   210
         Left            =   7095
         TabIndex        =   17
         Top             =   45
         Width           =   285
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "警示"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   7425
         TabIndex        =   16
         Top             =   60
         Width           =   360
      End
   End
   Begin MSComctlLib.ListView lvwSelect 
      Height          =   2685
      Left            =   5490
      TabIndex        =   2
      Top             =   435
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   4736
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "结果选择"
         Object.Width           =   2999
      EndProperty
   End
   Begin zl9LisWork.VsfGrid vsf 
      Height          =   2850
      Left            =   0
      TabIndex        =   0
      Top             =   390
      Width           =   5430
      _ExtentX        =   9578
      _ExtentY        =   5027
   End
   Begin MSComctlLib.StatusBar sbrInfo 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   4845
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4586
            MinWidth        =   4586
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4586
            MinWidth        =   4586
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraComment 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   3420
      Width           =   7050
      Begin VB.TextBox txtDiagnose 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   3960
         Locked          =   -1  'True
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   60
         Width           =   3000
      End
      Begin VB.TextBox txtComment 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   450
         Locked          =   -1  'True
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   60
         Width           =   3000
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "诊断信息"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   3540
         TabIndex        =   7
         Top             =   90
         Width           =   405
      End
      Begin VB.Label lblComment 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "检验备注"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   30
         TabIndex        =   6
         Top             =   90
         Width           =   405
      End
   End
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   7020
      Top             =   4440
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmLisStationWrite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Private mlngKey As Long    '标本ID
Private mDeviceID As Long
Private mstrType As String '检验类型
Private mblnEdit As Boolean '是否允许编辑
Private mbytRedoNumber As Long '重做次数
Private mblnLoadHistory As Boolean '是否装入历史数据
Private mSelectRedo As Boolean '是否选择了重做
Private mblnChangeEdit As Boolean, mblnEvent As Boolean
Private mLngPatientID As Long                               '病人ID
Private mstrPatientName As String                           '病人姓名
Private lngReferenceLow As Long                             '参考低颜色
Private lngReferenceHigh As Long                            '参考高颜色
Private lngReferenceExigency As Long                        '参考警示颜色
Public mblnPatientFind As Boolean                           '是否按病人来查看
Const mintColCount As Integer = 29                          '最示列表中分列显示一列中有多少个COL

Private Enum mCol
    检验项目 = 1
    原始结果
    检验结果
    单位
    CV
    结果标志
    上次结果
    上次时间
    结果参考
    结果类型
    仪器id
    计算公式
    结果范围
    固定项目
    小数
    警戒上限
    警戒下限
    诊疗项目id
    排列序号
    标本ID
    od
    CUTOFF
    COV
    酶标板ID
    变异报警
    变异警示
    仪器提示
    仪器审核标识
End Enum

Public Event StartEdit(Cancel As Boolean)

Private Function CalcDefaultFlag(ByVal strValue As String, ByVal strReference As String, Optional ByVal bytMode As Byte = 1, _
    Optional ByVal strAlarmLow As String, Optional ByVal strAlarmHigh As String, Optional ByVal lngItemID As Long) As String
    
    '--------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '返回:
    '--------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim rsTmp As New ADODB.Recordset
    
    
    If Len(Trim(strValue)) = 0 Then CalcDefaultFlag = "": Exit Function
    
    CalcDefaultFlag = ""
    
    If InStr(strReference, vbCrLf) > 0 Then strReference = Mid(strReference, 1, InStr(strReference, vbCrLf) - 1)
    If Trim(strReference) = "" Then Exit Function
                
    If bytMode = 2 Or (bytMode = 3 And IsNumeric(strValue) = False) Then  '定性、半定量
        If bytMode = 2 Or InStr(strReference, "～") = 0 Or Trim(strValue) Like "*阳*" Or Trim(strValue) Like "*+*" Or _
            Trim(strValue) Like "*±*" Or Trim(strValue) Like "*阴*" Or Trim(strValue) Like "*-*" Then
            '定性或无范围参考的半定量
            If (Len(Trim(strReference)) > 0 And (Trim(strReference) Like (Trim(strValue) & "*") Or Trim(strReference) Like ("*" & Trim(strValue)))) Or _
                (Not (Trim(strValue) Like "*阳*" Or Trim(strValue) Like "*+*" Or Trim(strValue) Like "*±*")) Then
                CalcDefaultFlag = ""
            Else
                CalcDefaultFlag = "异常"
            End If
            Exit Function
        Else
            '获取半定量值
            For i = 1 To Len(Trim(strValue))
                If InStr("01234567890.", Mid(strValue, i, 1)) > 0 Then Exit For
            Next
            If i > Len(Trim(strValue)) Then Exit Function
            strValue = Val(Mid(strValue, i))
        End If
    End If
    
'    If InStr(strValue, ">") Then CalcDefaultFlag = "↑": Exit Function
'    If InStr(strValue, "<") Then CalcDefaultFlag = "↓": Exit Function
    strValue = Replace(strValue, "<", "")
    strValue = Replace(strValue, ">", "")
    
    '如果不是数字就不做计算直接退出
    If IsNumeric(strValue) = False Then Exit Function
    
    
    
    If InStr(strReference, "～") > 0 Then
        
        '如果小于参考低值
        If Val(strValue) < Val(Mid(strReference, 1, InStr(strReference, "～") - 1)) And _
            Len(Trim(Mid(strReference, 1, InStr(strReference, "～") - 1))) > 0 Then
            CalcDefaultFlag = "↓"
        End If
        
        '如果大于参考高值
        If Val(strValue) > Val(Mid(strReference, InStr(strReference, "～") + 1)) And _
            Len(Trim(Mid(strReference, InStr(strReference, "～") + 1))) > 0 Then
            CalcDefaultFlag = "↑"
        End If
        
        If CalcDefaultFlag <> "" Then
            '高低判断
            If Len(Trim(strAlarmLow)) > 0 And Val(strAlarmLow) <> 0 Then
                If Val(strValue) < Val(strAlarmLow) Then
                    CalcDefaultFlag = "↓↓"
                    Exit Function
                End If
            End If
            If Len(Trim(strAlarmHigh)) > 0 And Val(strAlarmHigh) <> 0 Then
                If Val(strValue) > Val(strAlarmHigh) Then
                    CalcDefaultFlag = "↑↑"
                    Exit Function
                End If
            End If
        End If
    Else
        '高低判断
        If Len(Trim(strAlarmLow)) > 0 And Val(strAlarmLow) <> 0 Then
            If Val(strValue) < Val(strAlarmLow) Then
                CalcDefaultFlag = "↓↓"
                Exit Function
            End If
        End If
        If Len(Trim(strAlarmHigh)) > 0 And Val(strAlarmHigh) <> 0 Then
            If Val(strValue) > Val(strAlarmHigh) Then
                CalcDefaultFlag = "↑↑"
                Exit Function
            End If
        End If
    End If
    
    gstrSql = "select nvl(多参考,0) as 多参考 from 检验项目 where 诊治项目id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    If rsTmp.EOF = False Then
        If rsTmp("多参考") = 1 Then
            CalcDefaultFlag = ""
        End If
    End If

End Function

Private Function CalcExpress(ByVal Vsf As Object, ByVal strExPress As String) As String
    
    '--------------------------------------------------------------------------------------------------------
    '功能:在表格中计算某一表达式的结果
    '参数:vsf           存放数据的表格
    '     strExpress    要计算的表达式
    '返回:计算结果值
    '--------------------------------------------------------------------------------------------------------
    
    Dim strTmpPress As String
    Dim rs As New ADODB.Recordset
    
    Dim lngTmpID As Long
    Dim lngLeftPos As Long
    Dim lngRightPos As Long
    Dim lngLoop As Long
    Dim sglValue As String
    Dim intCol As Integer, intCols As Integer
    
    On Error GoTo errH
    
    CalcExpress = 0
    
    strTmpPress = strExPress
    If strTmpPress <> "" Then
        
        intCols = GetColCount(Vsf.Cols)
        If intCols = 0 Then intCols = 1
        
        lngLeftPos = InStr(strTmpPress, "[")
        lngRightPos = InStr(strTmpPress, "]")
        
        Do While lngLeftPos > 0
        
            lngTmpID = Val(Mid(strTmpPress, lngLeftPos + 1, lngRightPos - lngLeftPos - 1))
            
            '判断lngTmpID是否也是计算项目
            For intCol = 0 To intCols - 1
                For lngLoop = 1 To Vsf.Rows - 1
'                    If Val(Vsf.RowData(lngLoop)) = lngTmpID Then
                    If Val(Me.Vsf.Cell(flexcpData, lngLoop, intCol * mintColCount, lngLoop, intCol * mintColCount)) = lngTmpID Then
                        If Trim(Vsf.TextMatrix(lngLoop, mCol.计算公式 + intCol * mintColCount)) <> "" Then
                            '是计算项目,先计算出此结果
                            sglValue = CalcExpress(Vsf, Trim(Vsf.TextMatrix(lngLoop, mCol.计算公式 + intCol * mintColCount)))
                        Else
                            '不是计算项目,直接取此结果
                            sglValue = Vsf.TextMatrix(lngLoop, mCol.检验结果 + intCol * mintColCount)
                            If sglValue = "" Then
                                CalcExpress = ""
                                Exit Function
                            Else
                                sglValue = Val(sglValue)
                            End If
                        End If
                        
                        Exit For
                        
                    End If
                Next
                If Val(sglValue) <> 0 Then Exit For
            Next
            
            '在当前表格中没有此检验项目,认为结果为零
            If lngLoop = Vsf.Rows Then sglValue = 0
                                        
            '以结果替代表达式中的计算因子
            strTmpPress = Mid(strTmpPress, 1, lngLeftPos - 1) & sglValue & Mid(strTmpPress, lngRightPos + 1)
            
            '查下一个计算因子的位置
            lngLeftPos = InStr(strTmpPress, "[")
            lngRightPos = InStr(strTmpPress, "]")
            sglValue = ""
        Loop
                
        '计算表达式的结果
        On Error Resume Next
        Set rs = zlDatabase.OpenSQLRecord("SELECT " & strTmpPress & " AS 结果 FROM DUAL", Me.Caption)
        If rs.BOF = False Then CalcExpress = zlCommFun.Nvl(rs("结果"), 0)
        On Error GoTo 0
        
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ShowValue(ByVal intType As Integer, Optional intAttr As Integer = 2, Optional lngItemID As Long = 0)
    'intType：1－列出结果值、2－列出备注值
    'intAttr：2－文字、3－半定量
    Dim rs As New ADODB.Recordset
    Dim strsql As String, strValue As String, i As Long, aValues() As String
    Dim intColCount As Integer
    
    On Error GoTo errH
    
    Select Case intType
        Case 1
            strsql = "SELECT ROWNUM AS ID,编码,名称 As 取值 FROM 检验结果描述 A " & _
                " WHERE 分类=[1]"
            intColCount = GetColCount(Vsf.Col)
            Set rs = zlDatabase.OpenSQLRecord(strsql, Me.Caption, Vsf.TextMatrix(Vsf.Row, mCol.结果范围 + intColCount * mintColCount))
            With lvwSelect
                .ListItems.Clear
                .Tag = 1
                
                Do While Not rs.EOF
                    .ListItems.Add , "_" & rs("ID"), Nvl(rs("取值"))
                
                    rs.MoveNext
                Loop
            End With
        
            If intAttr <> 1 Then '非定量，取值序列
                strsql = "SELECT 取值序列 FROM 检验项目 WHERE 诊治项目ID=[1]"
                Set rs = zlDatabase.OpenSQLRecord(strsql, Me.Caption, lngItemID)
                If rs.EOF Then
                    strValue = "-|±|+|++|+++|++++"
                Else
                    strValue = Nvl(rs("取值序列"), "-|±|+|++|+++|++++")
                    strValue = Replace(strValue, ";", "|")
                End If
                aValues = Split(strValue, "|")
                With lvwSelect
                    For i = 0 To UBound(aValues)
                        .ListItems.Add , "V" & i, aValues(i)
                    Next
                End With
            End If
        Case 2
            strsql = "SELECT Rownum As ID,A.编码,A.简码,A.名称,A.说明 As 取值 FROM 检验备注文字 A " & _
                "WHERE A.分类 Is Null Or A.分类=[1]"
            Set rs = zlDatabase.OpenSQLRecord(strsql, Me.Caption, mstrType)
            With lvwSelect
                .ListItems.Clear
                .Tag = 2
                
                Do While Not rs.EOF
                    .ListItems.Add , "_" & rs("ID"), Nvl(rs("取值"))
                
                    rs.MoveNext
                Loop
            End With
    End Select
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function ReadPatient() As Boolean
    '-----------------------------------------------------------------------------------------
    '功能:
    '-----------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset, mstrSQL As String
    Dim lngPatientID As Long
    Dim mbytMode As Integer '0＝无主标本
    Dim mlngLoop As Long
    Dim strTmp As String
    Dim lngAdvice As Long     '医嘱ID
    Dim intColCount As Integer, intCol As Integer
    Dim blnMoved As Boolean                                         '是否移出
    Dim strSQLbak As String
    Dim strStart As String
    Dim strEnd As String
    
    On Error GoTo ErrHand
'    If mblnLoadHistory Then ReadPatient = True: Exit Function
    mblnLoadHistory = True
    
    
    Vsf.Rows = 2
    Vsf.Cell(flexcpText, 1, 0, 1, Vsf.Cols - 1) = ""
    Vsf.Cell(flexcpForeColor, 1, 0, 1, 1) = vbBack
    
    strStart = GetDateTime(Split(zlDatabase.GetPara("标本范围", 100, 1208, "今  天") & ";", ";")(0), 1)
    strEnd = GetDateTime(Split(zlDatabase.GetPara("标本范围", 100, 1208, "今  天") & ";", ";")(0), 2)
    
    If strStart = "自定义" Then
        strStart = Split(zlDatabase.GetPara("标本范围", 100, 1208, "今  天") & ";" & Now & ";" & Now, ";")(1)
        strStart = Format(strStart, "yyyy-mm-dd 00:00:00")
        strEnd = Split(zlDatabase.GetPara("标本范围", 100, 1208, "今  天") & ";" & Now & ";" & Now, ";")(2)
        strEnd = Format(strEnd, "yyyy-mm-dd 23:59:59")
    Else
        If strStart = "" Then strStart = GetDateTime("今  天", 1)
        If strEnd = "" Then strEnd = GetDateTime("今  天", 2)
    End If

    mstrSQL = "Select /*+ rule */ Distinct A.标本ID ,a.诊疗项目id ,A.编码, a.排列序号, a.固定项目, a.Id, a.检验项目, a.原始结果, a.上次结果, a.上次时间, a.Cv," & vbNewLine & _
                "            Decode(a.本次结果, '-', '阴性（-）', '+', '阳性（+）', '*', '*.**', a.本次结果) As 本次结果, Rownum As 序号, a.计算公式," & vbNewLine & _
                "            a.结果类型, a.标志, a.仪器id, a.标本类别, a.核收时间, a.标本序号, a.标本号显示, a.检验备注, a.姓名, a.性别, a.年龄, a.门诊号, a.住院号," & vbNewLine & _
                "            a.当前床号, a.主页id, a.结果范围, Nvl(G.小数位数,2) as 小数, a.警戒上限, a.警戒下限, a.单位,a.结果参考 as 参考, " & vbNewLine & _
                "                           Trim(Replace(Replace(' ' || Zlgetreference(a.Id, a.标本类型, Decode(a.性别, '男', 1, '女', 2, 0), a.出生日期," & vbNewLine & _
                "                                                                                                                   a.仪器id, a.年龄,a.申请科室id), ' .', '0.'), '～.', '～0.')) As 参考1," & vbNewLine & _
                "            a.OD,a.CUTOFF,a.COV,a.酶标板ID,a.变异报警,a.变异警示,lpad(编码,10,'0') as 排序,a.检验人,a.标本类型 " & vbNewLine & _
                "From (Select A.id as 标本ID ,b.诊疗项目id, decode(d.排列序号,Null,nvl(h.编码,C.编码),d.排列序号) as 编码, Nvl(b.排列序号, 9999) As 排列序号, Decode(b.诊疗项目id, Null, 0, 1) As 固定项目," & vbNewLine & _
                "                           b.检验项目id As Id, " & vbNewLine & _
                "                           " & IIf(chkChina.Value = 1, " c.中文名 || Decode(d.缩写, Null, '', '(' || d.缩写 || ')') As 检验项目 ", "d.缩写 as 检验项目 ") & vbNewLine & _
                "                           , b.原始结果," & vbNewLine & _
                "                           '' As 上次结果, '' As 上次时间, '' As Cv, b.检验结果 As 本次结果, d.计算公式, d.结果类型," & vbNewLine & _
                "                           Decode(b.结果标志, 3, '↑', 2, '↓', 1, '', 4, '异常', 5, '↓↓', 6, '↑↑', '') As 标志," & vbNewLine & _
                "                           Nvl(a.仪器id, -1) As 仪器id, Nvl(a.标本类别, 0) As 标本类别, a.核收时间, a.标本序号," & vbNewLine & _
                "                           Decode(a.仪器id, Null," & vbNewLine & _
                "                                           To_Char(Trunc(a.标本序号 / 10000) + 1, '0000') || '-' || To_Char(Mod(a.标本序号, 10000), '0000')," & vbNewLine & _
                "                                           a.标本序号) As 标本号显示, a.检验备注, a.姓名, a.性别, a.年龄, a.标本类型,a.出生日期,a.门诊号, a.住院号," & vbNewLine & _
                "                           a.床号 As 当前床号, a.主页id, d.结果范围, d.警戒上限, d.警戒下限, d.单位,b.OD,B.CUTOFF,B.SCO as COV,b.酶标板ID, " & vbNewLine & _
                "                           d.变异报警率 as  变异报警,d.变异警示率 as 变异警示,b.结果参考,a.检验人,a.申请科室ID " & vbNewLine & _
                "            From 检验标本记录 a, 检验普通结果 b, 诊治所见项目 c, 检验项目 d, 诊疗项目目录 h" & vbNewLine & _
                "            Where a.Id = b.检验标本id And b.检验项目id = c.Id And c.Id = d.诊治项目id And" & vbNewLine & _
                "                        b.诊疗项目id = h.Id(+) And b.记录类型 = 0 And a.病人ID = [1] and a.核收时间 between [2] and [3] " & vbNewLine & _
                "            ) A ,检验仪器项目 G" & _
                "  Where A.仪器id = G.仪器id(+) And A.ID = G.项目id(+) "


    
    If blnMoved Then
        strSQLbak = mstrSQL
        strSQLbak = Replace(strSQLbak, "检验标本记录", "H检验标本记录")
        strSQLbak = Replace(strSQLbak, "检验普通结果", "H检验普通结果")
        strSQLbak = Replace(strSQLbak, "检验申请项目", "H检验申请项目")
        mstrSQL = mstrSQL & " Union ALL " & strSQLbak
    End If
    
    mstrSQL = mstrSQL & " Order by 排序,排列序号,核收时间 desc "
    
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mLngPatientID, CDate(strStart), CDate(strEnd))
    
    If rs.BOF = False Then
        '初始标本信息
        mDeviceID = rs("仪器ID")
        Me.txtComment.Text = Nvl(rs("检验备注"))
        
        Vsf.TextMatrix(0, 0) = "#"
'        Call FillGrid_UQ(Vsf, rs, Array("", "", "", ""))
        Call ReadVsf_Patient(rs, Array("", "", "", ""))
        Vsf.TextMatrix(0, 0) = ""
        Vsf.Cell(flexcpBackColor, 1, 0, Vsf.Rows - 1, 0) = &HFDD6C6
        rs.MoveFirst
        
        Call FormatVsfCell(Vsf, mCol.检验结果, "0.0######", IIf(Nvl(rs("结果类型"), 0) = 1, 0, 1), _
                IIf(mDeviceID > 0, mCol.小数, -1))
                
        Call FormatVsfCell(Vsf, mCol.原始结果, "0.0######", IIf(Nvl(rs("结果类型"), 0) = 1, 0, 1), _
                IIf(mDeviceID > 0, mCol.小数, -1))
        
'        If chkLast.Value Then LoadLastValue
        '--每次都读出历史结果
        LoadLastValue
    Else
        mDeviceID = -1
        Me.txtComment.Text = ""
        ResetVsf Vsf
    End If
    
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        For mlngLoop = 1 To Vsf.Rows - 1
            Call ApplyResultColor(Vsf, mlngLoop, mCol.检验结果 + intCol * mintColCount, _
                Decode(Vsf.TextMatrix(mlngLoop, mCol.结果标志 + intCol * mintColCount), "↑", 3, "↓", 2, "异常", 4, "↑↑", 6, "↓↓", 5, 1))
        Next
    Next
    
    '写入诊断信息
    Me.txtDiagnose.Text = ""
    gstrSql = "Select b.医嘱id, b.项目, b.排列, b.内容" & vbNewLine & _
                "From 检验标本记录 a, 病人医嘱附件 b" & vbNewLine & _
                "Where a.医嘱id = b.医嘱id and a.ID = [1] " & vbNewLine & _
                "Order By 医嘱id, 排列"
    Set rs = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey)
    
    Do Until rs.EOF
        strTmp = strTmp & Nvl(rs("项目")) & ":" & Replace(Nvl(rs("内容")), vbCrLf, vbCrLf & "    ") & vbCrLf
        rs.MoveNext
    Loop
    Me.txtDiagnose.Text = strTmp
    ReadPatient = True
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function ReadVsf_Patient(ByVal rsData As ADODB.Recordset, Optional ByVal MaskArray As Variant, Optional ByVal blnClear As Boolean = True) As Boolean
    Dim lngLoop As Long
    Dim strMask As String
    Dim lngRow As Long, lngCurrRow As Long
    Dim strOldValue As String, strNewValue As String
    Dim intColCount  As Integer
    Dim intCol As Integer, intRow As Integer
    Dim lngHeight As Long
    Dim blnShowType As Boolean
    Dim lngItem As Long                                     '诊疗项目ID
    Dim intItemCount As Integer                             '当前有几个项目
    Dim rsTmp As New ADODB.Recordset
    blnShowType = zlDatabase.GetPara("自适应显示结果", 100, 1208, False)
    If fraComment.Tag <> "" Then blnShowType = True
    
    If blnClear Then
        Vsf.Rows = 2
        Vsf.RowData(1) = 0
        For lngLoop = 0 To Vsf.Cols - 1
            Vsf.TextMatrix(1, lngLoop) = ""
            Vsf.Cell(flexcpData, 1, lngLoop, 1, lngLoop) = ""
        Next
        lngRow = 0
        Vsf.Cols = mintColCount
    Else
        '预先有一空行
        With Vsf
            intColCount = GetColCount(.Cols)
            If intColCount = 0 Then intColCount = 1
            For intCol = 0 To intColCount - 1
                For intRow = 1 To .Rows - 1
                    If Val(.Cell(flexcpData, intRow, intCol * mintColCount, intRow, intCol * mintColCount)) = 0 Then
                        lngRow = intRow - 1
                        intColCount = intCol
                        Exit For
                    End If
                Next
            Next
        End With
    End If
    
    
    With Vsf.Body
        If .ClientHeight < .CellHeight * 15 Then
            lngHeight = .CellHeight * 15
        Else
            lngHeight = .ClientHeight
        End If
    End With
    
    Do While Not rsData.EOF
        lngCurrRow = FindRepeatLine(Vsf, CStr(zlCommFun.Nvl(rsData("ID"))))
'        lngCurrRow = -1
        If lngCurrRow = -1 Then
            '--------------------------------------诊疗项目不一样时先增加诊疗项目------------------------------------------
            If lngItem <> rsData("诊疗项目ID") Then
                intItemCount = intItemCount + 1
                gstrSql = "select 名称 from 诊疗项目目录 where id = [1] "
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(rsData("诊疗项目id")))
                
                With Vsf.Body

                    If (.CellHeight + 15) * (lngRow + 2) > lngHeight And blnShowType = True Then
                        intColCount = intColCount + 1
                        lngRow = 1
                        With Vsf
                            .NewColumn "#", 300, 7
                            .NewColumn "检验项目", 2100, 1
                            .NewColumn "原始结果", 0, 1
                            .NewColumn "本次结果", 1200, 1, , 1
                            .NewColumn "单位", 1000, 1
                            .NewColumn "CV", 0, 1
                            .NewColumn "标志", 450, 1
                            .NewColumn "上次结果", 0, 1
                            .NewColumn "上次时间", 0, 1
                            .NewColumn "参考", 1300, 1
                            .NewColumn "结果类型", 0, 1
                            .NewColumn "仪器id", 0, 1
                            .NewColumn "计算公式", 0, 1
                            .NewColumn "结果范围", 0, 1
                            .NewColumn "固定项目", 0, 1
                            .NewColumn "小数", 0, 1
                            .NewColumn "警戒上限", 0, 1
                            .NewColumn "警戒下限", 0, 1
                            .NewColumn "诊疗项目ID", 0, 1
                            .NewColumn "排列序号", 0, 1
                            .NewColumn "标本ID", 0, 1
                            .NewColumn "OD", 700, 1, , 1
                            .NewColumn "CUTOFF", 700, 1
                            .NewColumn "COV", 700, 1
                            .NewColumn "酶标板ID", 0, 1
                            .NewColumn "变异报警", 0, 1
                            .NewColumn "变异警示", 0, 1
                            .NewColumn "仪器提示", 1000, 1
                            .NewColumn "仪器审核标识", 1200, 1
                        End With
                    Else
                        lngRow = lngRow + 1
                    End If
                End With
                
                lngCurrRow = lngRow
                
            
            
                If Vsf.Rows < lngRow + 1 Then Vsf.Rows = lngRow + 1
                
                On Error Resume Next
                
                On Error GoTo ErrHand
                
                For lngLoop = 1 To mintColCount - 1
                    '设置合并
                    Me.Vsf.Body.MergeCol(lngLoop) = True
                    Me.Vsf.Body.MergeRow(lngCurrRow) = True
                    Me.Vsf.Body.MergeCells = flexMergeFree
                    intCol = intColCount * mintColCount + lngLoop
                    Vsf.TextMatrix(lngCurrRow, intCol) = rsTmp("名称") & "(" & rsData("检验人") & " " & Format(rsData("核收时间"), "yyyy-mm-dd") & ")"
                Next
                '设置颜色
                Vsf.Cell(flexcpBackColor, lngCurrRow, intColCount * mintColCount, lngCurrRow, intColCount * mintColCount + mintColCount - 1) = &HFDD6C6
                Me.Vsf.Cell(flexcpFontBold, lngCurrRow, intColCount * mintColCount, lngCurrRow, intColCount * mintColCount + mintColCount - 1) = True
                            
            End If
            '-----------------------------------------------------------------------------------------------------
            With Vsf.Body

                If (.CellHeight + 15) * (lngRow + 2) > lngHeight And blnShowType = True Then
                    intColCount = intColCount + 1
                    lngRow = 1
                    With Vsf
                        .NewColumn "#", 300, 7
                        .NewColumn "检验项目", 2100, 1
                        .NewColumn "原始结果", 0, 1
                        .NewColumn "本次结果", 1200, 1, , 1
                        .NewColumn "单位", 1000, 1
                        .NewColumn "CV", 0, 1
                        .NewColumn "标志", 450, 1
                        .NewColumn "上次结果", 0, 1
                        .NewColumn "上次时间", 0, 1
                        .NewColumn "参考", 1300, 1
                        .NewColumn "结果类型", 0, 1
                        .NewColumn "仪器id", 0, 1
                        .NewColumn "计算公式", 0, 1
                        .NewColumn "结果范围", 0, 1
                        .NewColumn "固定项目", 0, 1
                        .NewColumn "小数", 0, 1
                        .NewColumn "警戒上限", 0, 1
                        .NewColumn "警戒下限", 0, 1
                        .NewColumn "诊疗项目ID", 0, 1
                        .NewColumn "排列序号", 0, 1
                        .NewColumn "标本ID", 0, 1
                        .NewColumn "OD", 700, 1, , 1
                        .NewColumn "CUTOFF", 700, 1
                        .NewColumn "COV", 700, 1
                        .NewColumn "酶标板ID", 0, 1
                        .NewColumn "变异报警", 0, 1
                        .NewColumn "变异警示", 0, 1
                        .NewColumn "仪器提示", 1000, 1
                        .NewColumn "仪器审核标识", 1200, 1
                    End With
                Else
                    lngRow = lngRow + 1
                End If
            End With
            
            lngCurrRow = lngRow
        
            If Vsf.Rows < lngRow + 1 Then Vsf.Rows = lngRow + 1
            
            On Error Resume Next
'            Vsf.RowData(lngCurrRow) = CStr(zlCommFun.Nvl(rsData("ID")))
            Vsf.Cell(flexcpData, lngCurrRow, intColCount * mintColCount, lngCurrRow, intColCount * mintColCount) = CStr(Nvl(rsData("ID")))
            
            On Error GoTo ErrHand
            
            For lngLoop = 0 To mintColCount - 1
                intCol = intColCount * mintColCount + lngLoop
                
                If Trim(Vsf.TextMatrix(0, intCol)) <> "" Then
                    If Vsf.TextMatrix(0, intCol) = "#" Then
                        Vsf.TextMatrix(lngCurrRow, intCol) = IIf(intColCount > 0, intColCount * (Vsf.Body.Rows - 1) + lngCurrRow, lngCurrRow) - intItemCount
                        Vsf.Cell(flexcpBackColor, lngCurrRow, intCol, lngCurrRow, intCol) = &HFDD6C6
                    Else
                        On Error Resume Next
                        strMask = ""
                        strMask = MaskArray(intCol)
                                                
                        On Error GoTo ErrHand

                         
                        If strMask <> "" Then
                            strNewValue = Format(zlCommFun.Nvl(rsData(Vsf.TextMatrix(0, intCol))), strMask)
                        Else
                            strNewValue = zlCommFun.Nvl(rsData(Vsf.TextMatrix(0, intCol)))
                        End If
                        Vsf.TextMatrix(lngCurrRow, intCol) = strNewValue
                    End If
                End If
                
            Next
        End If
        lngItem = Val(Nvl(rsData("诊疗项目ID"), 0))
        rsData.MoveNext
    Loop
'    Call chkOriginal_Click: Call chkLast_Click: Call chkSign_Click
'    Call chkUnit_Click: Call chkReferrence_Click: Call chkMB_Click
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        Vsf.Body.ColWidth(mCol.检验项目 + intCol * mintColCount) = IIf(chkChina.Value, 2100, 1000)
        Vsf.Body.ColWidth(mCol.原始结果 + intCol * mintColCount) = IIf(chkOriginal.Value, 900, 0)
        Vsf.Body.ColWidth(mCol.上次结果 + intCol * mintColCount) = IIf(chkLast.Value, 900, 0)
        Vsf.Body.ColWidth(mCol.上次时间 + intCol * mintColCount) = IIf(chkLast.Value, 1000, 0)
        Vsf.Body.ColWidth(mCol.结果标志 + intCol * mintColCount) = IIf(chkSign.Value, 450, 0)
        Vsf.Body.ColWidth(mCol.单位 + intCol * mintColCount) = IIf(chkUnit.Value, 1000, 0)
        Vsf.Body.ColWidth(mCol.结果参考 + intCol * mintColCount) = IIf(chkReferrence.Value, 1300, 0)
        Vsf.Body.ColWidth(mCol.od + intCol * mintColCount) = IIf(chkMB.Value, 700, 0)
        Vsf.Body.ColWidth(mCol.CUTOFF + intCol * mintColCount) = IIf(chkMB.Value, 700, 0)
        Vsf.Body.ColWidth(mCol.COV + intCol * mintColCount) = IIf(chkMB.Value, 700, 0)
        Vsf.Body.ColWidth(mCol.仪器提示 + intCol * mintColCount) = IIf(chkYiQiTiShi.Value, 1000, 0)
        Vsf.Body.ColWidth(mCol.仪器审核标识 + intCol * mintColCount) = IIf(chkYiQiBiaoShi.Value, 1200, 0)
    Next
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function ReadData() As Boolean
    '-----------------------------------------------------------------------------------------
    '功能:
    '-----------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset, mstrSQL As String
    Dim lngPatientID As Long
    Dim mbytMode As Integer '0＝无主标本
    Dim mlngLoop As Long
    Dim strTmp As String
    Dim lngAdvice As Long     '医嘱ID
    Dim intColCount As Integer, intCol As Integer
    Dim blnMoved As Boolean                                         '是否移出
    Dim strSQLbak As String
    
    On Error GoTo ErrHand
    If mblnLoadHistory Then ReadData = True: Exit Function
    mblnLoadHistory = True
    
    
    Vsf.Rows = 2
    Vsf.Cell(flexcpText, 1, 0, 1, Vsf.Cols - 1) = ""
    Vsf.Cell(flexcpForeColor, 1, 0, 1, 1) = vbBack

     '1-正常、2-偏低、3-偏高、4-阳性(异常)、5-警戒下限、6-警戒上限
    mstrSQL = "Select a.医嘱Id,a.报告结果, a.病人id, a.主页id, a.操作类型, a.检验人, a.检验时间, a.审核人, a.审核时间,病人ID,姓名,a.初审人,a.初审时间  " & vbNewLine & _
            "From 检验标本记录 a" & vbNewLine & _
            "Where a.Id = [1] "
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mlngKey)
    If Not rs.EOF Then
        lngPatientID = Nvl(rs("病人ID"), 0)
        mstrType = Nvl(rs("操作类型"))
        '是否不区分仪器来显示核收项目
        lngAdvice = zlDatabase.GetPara("不区分仪器显示核收项目", 100, 1208, 0)
        If lngAdvice = 0 Then
            lngAdvice = 0
        Else
            lngAdvice = Nvl(rs("医嘱ID"), 0)
        End If
        mbytRedoNumber = IIf(mSelectRedo, mbytRedoNumber - 1, Nvl(rs("报告结果"), 0))
        
        mSelectRedo = False
        
        mbytMode = IIf(IsNull(rs("病人ID")), 0, 1)
        
        With sbrInfo
            .Panels(1).Text = "报告人：" & Nvl(rs("检验人"))
            .Panels(2).Text = "报告时间：" & IIf(IsNull(rs("检验时间")), "", Format(rs("检验时间"), "yyyy-MM-dd hh:mm"))
            If Nvl(rs("审核人")) <> "" Then
                .Panels(3).Text = "审核人：" & Nvl(rs("审核人"))
                .Panels(4).Text = "审核时间：" & IIf(IsNull(rs("审核时间")), "", Format(rs("审核时间"), "yyyy-MM-dd hh:mm"))
            Else
                If Nvl(rs("初审人")) <> "" Then
                    .Panels(3).Text = "初审人：" & Nvl(rs("初审人"))
                    .Panels(4).Text = "初审时间：" & IIf(IsNull(rs("初审时间")), "", Format(rs("初审时间"), "yyyy-MM-dd hh:mm"))
                Else
                    .Panels(3).Text = "审核人：" & Nvl(rs("审核人"))
                    .Panels(4).Text = "审核时间：" & IIf(IsNull(rs("审核时间")), "", Format(rs("审核时间"), "yyyy-MM-dd hh:mm"))
                End If
            End If
        End With
        blnMoved = MovedByDate(CDate(Format(Nvl(rs("检验时间")), "yyyy-MM-dd hh:mm:ss")))
        mLngPatientID = Nvl(rs("病人ID"), 0)
        mstrPatientName = Nvl(rs("姓名"))
    Else
        mstrSQL = "Select a.医嘱Id,a.报告结果, a.病人id, a.主页id, a.操作类型, a.检验人, a.检验时间, a.审核人, a.审核时间,病人ID,姓名,a.初审人,a.初审时间 " & vbNewLine & _
            "From h检验标本记录 a" & vbNewLine & _
            "Where a.Id = [1] "
        Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mlngKey)
        If Not rs.EOF Then
            lngPatientID = Nvl(rs("病人ID"), 0)
            mstrType = Nvl(rs("操作类型"))
            '是否不区分仪器来显示核收项目
            lngAdvice = zlDatabase.GetPara("不区分仪器显示核收项目", 100, 1208, 0)
            If lngAdvice = 0 Then
                lngAdvice = 0
            Else
                lngAdvice = Nvl(rs("医嘱ID"), 0)
            End If
            mbytRedoNumber = IIf(mSelectRedo, mbytRedoNumber - 1, Nvl(rs("报告结果"), 0))
            
            mSelectRedo = False
            
            mbytMode = IIf(IsNull(rs("病人ID")), 0, 1)
            
            With sbrInfo
                .Panels(1).Text = "报告人：" & Nvl(rs("检验人"))
                .Panels(2).Text = "报告时间：" & IIf(IsNull(rs("检验时间")), "", Format(rs("检验时间"), "yyyy-MM-dd hh:mm"))
                If Nvl(rs("审核人")) <> "" Then
                    .Panels(3).Text = "审核人：" & Nvl(rs("审核人"))
                    .Panels(4).Text = "审核时间：" & IIf(IsNull(rs("审核时间")), "", Format(rs("审核时间"), "yyyy-MM-dd hh:mm"))
                Else
                    If Nvl(rs("初审人")) <> "" Then
                        .Panels(3).Text = "初审人：" & Nvl(rs("初审人"))
                        .Panels(4).Text = "初审时间：" & IIf(IsNull(rs("初审时间")), "", Format(rs("初审时间"), "yyyy-MM-dd hh:mm"))
                    Else
                        .Panels(3).Text = "审核人：" & Nvl(rs("审核人"))
                        .Panels(4).Text = "审核时间：" & IIf(IsNull(rs("审核时间")), "", Format(rs("审核时间"), "yyyy-MM-dd hh:mm"))
                    End If
                End If
            End With
            blnMoved = MovedByDate(CDate(Format(Nvl(rs("检验时间")), "yyyy-MM-dd hh:mm:ss")))
            mLngPatientID = Nvl(rs("病人ID"), 0)
            mstrPatientName = Nvl(rs("姓名"))
        Else
            lngPatientID = 0
            mstrType = ""
            mbytRedoNumber = 0
            
            mbytMode = 0
            
            With sbrInfo
                .Panels(1).Text = "报告人："
                .Panels(2).Text = "报告时间："
                .Panels(3).Text = "审核人："
                .Panels(4).Text = "审核时间："
            End With
        End If
    End If
    '1-正常、2-偏低、3-偏高、4-阳性(异常)、5-警戒下限、6-警戒上限
    If mbytMode = 1 Then

'        mstrSQL = "Select /*+ rule */ 标本ID, 诊疗项目id, 排列序号, 固定项目, Id, 检验项目, 原始结果, 上次结果, 上次时间, Cv," & vbNewLine & _
'                    "            Decode(本次结果, '-', '阴性（-）', '+', '阳性（+）', '*', '*.**', 本次结果) As 本次结果, Rownum As 序号, 计算公式," & vbNewLine & _
'                    "            结果类型, 标志, 仪器id, 标本类别, 核收时间, 标本序号, 标本号显示, 检验备注, 姓名, 性别, 年龄, 门诊号, 住院号," & vbNewLine & _
'                    "            当前床号, 主页id, 结果范围, 小数, 警戒上限, 警戒下限, 单位, 参考" & vbNewLine & _
'                    "From (Select a.ID as 标本ID,b.诊疗项目id, h.编码, Nvl(b.排列序号, 9999) As 排列序号, Decode(b.诊疗项目id, Null, 0, 1) As 固定项目," & vbNewLine & _
'                    "                           b.检验项目id As Id, c.中文名 || Decode(d.缩写, Null, '', '(' || d.缩写 || ')') As 检验项目, b.原始结果," & vbNewLine & _
'                    "                           '' As 上次结果, '' As 上次时间, '' As Cv, b.检验结果 As 本次结果, d.计算公式, d.结果类型," & vbNewLine & _
'                    "                           Decode(b.结果标志, 3, '↑', 2, '↓', 1, '', 4, '异常', 5, '↓↓', 6, '↑↑', '') As 标志," & vbNewLine & _
'                    "                           Nvl(a.仪器id, -1) As 仪器id, Nvl(a.标本类别, 0) As 标本类别, a.核收时间, a.标本序号," & vbNewLine & _
'                    "                           Decode(a.仪器id, Null," & vbNewLine & _
'                    "                                           To_Char(Trunc(a.标本序号 / 10000) + 1, '0000') || '-' || To_Char(Mod(a.标本序号, 10000), '0000')," & vbNewLine & _
'                    "                                           a.标本序号) As 标本号显示, a.检验备注, a.姓名, a.性别, a.年龄, a.门诊号, a.住院号," & vbNewLine & _
'                    "                           a.床号 As 当前床号, a.主页id, d.结果范围, Nvl(g.小数位数, 2) As 小数, d.警戒上限, d.警戒下限, d.单位," & vbNewLine & _
'                    "                           Trim(Replace(Replace(' ' || Zlgetreference(c.Id, a.标本类型, Decode(a.性别, '男', 1, '女', 2, 0), a.出生日期," & vbNewLine & _
'                    "                                                                                                                   a.仪器id, a.年龄), ' .', '0.'), '～.', '～0.')) As 参考" & vbNewLine & _
'                    "            From 检验标本记录 a, 检验普通结果 b, 诊治所见项目 c, 检验项目 d, 检验仪器项目 g, 诊疗项目目录 h" & vbNewLine & _
'                    "            Where a.Id = b.检验标本id And b.检验项目id = c.Id And c.Id = d.诊治项目id And" & vbNewLine & _
'                    "                        (g.仪器id = a.仪器id + 0 Or g.仪器id Is Null Or a.仪器id Is Null) And b.检验项目id = g.项目id(+) And" & vbNewLine & _
'                    "                        b.诊疗项目id = h.Id(+) And b.记录类型 = [1] And " & IIf(lngAdvice = 0, " a.Id = [2] ", " a.医嘱ID = [4] ")
'        mstrSQL = mstrSQL & " Union All" & vbNewLine & _
'                    "           Select a.ID as 标本ID,b.诊疗项目id, h.编码, Nvl(b.排列序号, 9999) As 排列序号, Decode(b.诊疗项目id, Null, 0, 1) As 固定项目," & vbNewLine & _
'                    "                           b.检验项目id As Id, c.中文名 || Decode(d.缩写, Null, '', '(' || d.缩写 || ')') As 检验项目, b.原始结果," & vbNewLine & _
'                    "                           '' As 上次结果, '' As 上次时间, '' As Cv, b.检验结果 As 本次结果, d.计算公式, d.结果类型," & vbNewLine & _
'                    "                           Decode(b.结果标志, 3, '↑', 2, '↓', 1, '', 4, '异常', 5, '↓↓', 6, '↑↑', '') As 标志," & vbNewLine & _
'                    "                           Nvl(a.仪器id, -1) As 仪器id, Nvl(a.标本类别, 0) As 标本类别, a.核收时间, a.标本序号," & vbNewLine & _
'                    "                           Decode(a.仪器id, Null," & vbNewLine & _
'                    "                                           To_Char(Trunc(a.标本序号 / 10000) + 1, '0000') || '-' || To_Char(Mod(a.标本序号, 10000), '0000')," & vbNewLine & _
'                    "                                           a.标本序号) As 标本号显示, a.检验备注, a.姓名, a.性别, a.年龄, a.门诊号, a.住院号," & vbNewLine & _
'                    "                           a.床号 As 当前床号, a.主页id, d.结果范围, Nvl(g.小数位数, 2) As 小数, d.警戒上限, d.警戒下限, d.单位," & vbNewLine & _
'                    "                           Trim(Replace(Replace(' ' || Zlgetreference(c.Id, a.标本类型, Decode(a.性别, '男', 1, '女', 2, 0), a.出生日期," & vbNewLine & _
'                    "                                                                                                                   a.仪器id, a.年龄), ' .', '0.'), '～.', '～0.')) As 参考" & vbNewLine & _
'                    "            From 检验标本记录 a, 检验普通结果 b, 诊治所见项目 c, 检验项目 d, 检验仪器项目 g, 诊疗项目目录 h" & vbNewLine & _
'                    "            Where a.Id = b.检验标本id And b.检验项目id = c.Id And c.Id = d.诊治项目id And" & vbNewLine & _
'                    "                        (g.仪器id = a.仪器id + 0 Or g.仪器id Is Null Or a.仪器id Is Null) And b.检验项目id = g.项目id(+) And" & vbNewLine & _
'                    "                        b.诊疗项目id = h.Id(+) And b.记录类型 = [1] And " & IIf(lngAdvice = 0, " a.Id = [2] ", " a.医嘱ID = [4] ") & vbNewLine & _
'                    "            Order By 编码, 排列序号)"
        '2008-02-13 修改排序 陈东
        mstrSQL = "Select /*+ rule */ Distinct A.标本ID ,a.诊疗项目id ,A.编码, a.排列序号, a.固定项目, a.Id, a.检验项目, a.原始结果, a.上次结果, a.上次时间, a.Cv," & vbNewLine & _
                    "            Decode(a.本次结果, '-', '阴性（-）', '+', '阳性（+）', '*', '*.**', a.本次结果) As 本次结果, Rownum As 序号, a.计算公式," & vbNewLine & _
                    "            a.结果类型, a.标志, a.仪器id, a.标本类别, a.核收时间, a.标本序号, a.标本号显示, a.检验备注, a.姓名, a.性别, a.年龄, a.门诊号, a.住院号," & vbNewLine & _
                    "            a.当前床号, a.主页id, a.结果范围, Nvl(G.小数位数,2) as 小数, " & vbNewLine & _
                    "            Trim(Replace(Replace(' ' || Zl_Get_Reference(4,a.Id, a.标本类型, Decode(a.性别, '男', 1, '女', 2, 0), a.出生日期," & vbNewLine & _
                    "                           a.仪器id, a.年龄,a.申请科室ID), ' .', '0.'), '～.', '～0.')) As 警戒上限," & vbNewLine & _
                    "            Trim(Replace(Replace(' ' || Zl_Get_Reference(3,a.Id, a.标本类型, Decode(a.性别, '男', 1, '女', 2, 0), a.出生日期," & vbNewLine & _
                    "                           a.仪器id, a.年龄,a.申请科室ID), ' .', '0.'), '～.', '～0.')) As 警戒下限," & vbNewLine & _
                    "            a.单位,a.结果参考 as 参考, " & vbNewLine & _
                    "                           Trim(Replace(Replace(' ' || Zlgetreference(a.Id, a.标本类型, Decode(a.性别, '男', 1, '女', 2, 0), a.出生日期," & vbNewLine & _
                    "                                                                                                                   a.仪器id, a.年龄,a.申请科室ID), ' .', '0.'), '～.', '～0.')) As 参考1," & vbNewLine & _
                    "            a.OD,a.CUTOFF,a.COV,a.酶标板ID,a.变异报警,a.变异警示,lpad(编码,4,'0') as 排序,a.标本类型,a.仪器提示,decode(a.仪器审核标识,1,'√',0,'×','') as 仪器审核标识  " & vbNewLine & _
                    "From (Select A.id as 标本ID ,b.诊疗项目id, decode(d.排列序号,Null,nvl(h.编码,C.编码),d.排列序号) as 编码, Nvl(b.排列序号, 9999) As 排列序号, Decode(b.诊疗项目id, Null, 0, 1) As 固定项目,"
                    
          mstrSQL = mstrSQL & " " & _
                    "                           b.检验项目id As Id, " & vbNewLine & _
                    "                           " & IIf(chkChina.Value = 1, " c.中文名 || Decode(d.缩写, Null, '', '(' || d.缩写 || ')') As 检验项目 ", "d.缩写 as 检验项目 ") & vbNewLine & _
                    "                           , b.原始结果," & vbNewLine & _
                    "                           '' As 上次结果, '' As 上次时间, '' As Cv, b.检验结果 As 本次结果, d.计算公式, d.结果类型," & vbNewLine & _
                    "                           Decode(b.结果标志, 3, '↑', 2, '↓', 1, '', 4, '异常', 5, '↓↓', 6, '↑↑', '') As 标志," & vbNewLine & _
                    "                           Nvl(a.仪器id, -1) As 仪器id, Nvl(a.标本类别, 0) As 标本类别, a.核收时间, a.标本序号," & vbNewLine & _
                    "                           Decode(a.仪器id, Null," & vbNewLine & _
                    "                                           To_Char(Trunc(a.标本序号 / 10000) + 1, '0000') || '-' || To_Char(Mod(a.标本序号, 10000), '0000')," & vbNewLine & _
                    "                                           a.标本序号) As 标本号显示, a.检验备注, a.姓名, a.性别, a.年龄, a.标本类型,a.出生日期,a.门诊号, a.住院号," & vbNewLine & _
                    "                           a.床号 As 当前床号, a.主页id, d.结果范围, d.警戒上限, d.警戒下限, d.单位,b.OD,B.CUTOFF,B.SCO as COV,b.酶标板ID, " & vbNewLine & _
                    "                           d.变异报警率 as  变异报警,d.变异警示率 as 变异警示,b.结果参考,a.申请科室ID ,e.仪器是否审核 as  仪器审核标识,e.审核内容  as 仪器提示 " & vbNewLine & _
                    ",Zl_To_Number(Zl_Get_Reference(1, b.检验项目id, a.标本类型, Decode(a.性别, '男', 1, '女', 2, 0), a.出生日期,a.仪器id, a.年龄,a.申请科室ID)) as 参考ID " & vbNewLine & _
                    "            From 检验标本记录 a, 检验普通结果 b, 诊治所见项目 c, 检验项目 d, 诊疗项目目录 h,检验流水线指标  e" & vbNewLine & _
                    "            Where a.Id = b.检验标本id And b.检验项目id = c.Id And c.Id = d.诊治项目id And  b.诊疗项目id = h.Id(+) and  b.检验标本id=e.标本id(+) and  b.检验项目id = e.项目id(+) And b.记录类型 = [1] And " & IIf(lngAdvice = 0, " a.Id = [2] ", " a.医嘱ID = [4] ")
        mstrSQL = mstrSQL & " Union All" & vbNewLine & _
                    "           Select a.id as 标本ID ,b.诊疗项目id, decode(d.排列序号,Null,nvl(h.编码,C.编码),d.排列序号) as 编码, Nvl(b.排列序号, 9999) As 排列序号, Decode(b.诊疗项目id, Null, 0, 1) As 固定项目," & vbNewLine & _
                    "                           b.检验项目id As Id, " & _
                    "                           " & IIf(chkChina.Value = 1, "c.中文名 || Decode(d.缩写, Null, '', '(' || d.缩写 || ')') As 检验项目", "d.缩写 as 检验项目 ") & vbNewLine & _
                    "                           , b.原始结果," & vbNewLine & _
                    "                           '' As 上次结果, '' As 上次时间, '' As Cv, b.检验结果 As 本次结果, d.计算公式, d.结果类型," & vbNewLine & _
                    "                           Decode(b.结果标志, 3, '↑', 2, '↓', 1, '', 4, '异常', 5, '↓↓', 6, '↑↑', '') As 标志," & vbNewLine & _
                    "                           Nvl(a.仪器id, -1) As 仪器id, Nvl(a.标本类别, 0) As 标本类别, a.核收时间, a.标本序号," & vbNewLine & _
                    "                           Decode(a.仪器id, Null," & vbNewLine & _
                    "                                           To_Char(Trunc(a.标本序号 / 10000) + 1, '0000') || '-' || To_Char(Mod(a.标本序号, 10000), '0000')," & vbNewLine & _
                    "                                           a.标本序号) As 标本号显示, a.检验备注, a.姓名, a.性别, a.年龄, a.标本类型,a.出生日期, a.门诊号, a.住院号," & vbNewLine & _
                    "                           a.床号 As 当前床号, a.主页id, d.结果范围, d.警戒上限, d.警戒下限, d.单位,b.OD,B.CUTOFF,B.SCO as COV,b.酶标板ID, " & vbNewLine & _
                    "                           d.变异报警率 as 变异报警,d.变异警示率 as 变异警示,b.结果参考,a.申请科室ID ,e.仪器是否审核 as  仪器审核标识 ,e.审核内容 as 仪器提示   " & vbNewLine & _
                    ",Zl_To_Number(Zl_Get_Reference(1, b.检验项目id, a.标本类型, Decode(a.性别, '男', 1, '女', 2, 0), a.出生日期,a.仪器id, a.年龄,a.申请科室ID)) as 参考ID " & vbNewLine & _
                    "            From 检验标本记录 a, 检验普通结果 b, 诊治所见项目 c, 检验项目 d,  诊疗项目目录 h,检验流水线指标  e" & vbNewLine & _
                    "            Where a.Id = b.检验标本id And b.检验项目id = c.Id And c.Id = d.诊治项目id And" & vbNewLine & _
                    "                        b.诊疗项目id = h.Id(+) And b.记录类型 = [1]  and  b.检验标本id=e.标本id(+) and  b.检验项目id = e.项目id(+)  And a.合并id = [2]" & vbNewLine & _
                    "            ) A ,检验仪器项目 G,检验项目参考 F" & _
                    "  Where A.仪器id = G.仪器id(+) And A.ID = G.项目id(+) and a.参考id=f.id(+)"

    Else
'        mstrSQL = "Select /*+ rule */ A.*,Rownum As 序号 From (SELECT a.检验标本ID as 标本ID,A.诊疗项目ID,A.排列序号,B.ID," & _
'                        "B.中文名||DECODE(C.缩写,NULL,'','('||C.缩写||')') AS 检验项目," & _
'                        "A.原始结果," & _
'                        "'' As 上次结果,'' as 上次时间 ,''As CV," & _
'                        "A.检验结果 As 本次结果," & _
'                        "C.计算公式," & _
'                        "C.结果类型," & _
'                        "DECODE(A.结果标志,3,'↑',2,'↓',1,'',4,'异常',5,'↓↓',6,'↑↑','') AS 标志," & _
'                        "Trim(REPLACE(REPLACE(' '||zlGetReference(B.ID,D.标本类型,0,NULL,D.仪器ID),' .','0.'),'～.','～0.')) AS 参考," & _
'                        "Nvl(D.仪器ID,-1) As 仪器ID,Nvl(D.标本类别,0) As 标本类别,D.核收时间,D.标本序号,D.检验备注,C.结果范围,0 As 固定项目,Nvl(X.小数位数,2) AS 小数," & _
'                        "C.警戒上限,C.警戒下限,C.单位 " & _
'                    "FROM 检验普通结果 A,诊治所见项目 B,检验项目 C,检验标本记录 D,检验仪器项目 X " & _
'                    "WHERE A.检验项目id = B.ID " & _
'                        "AND B.ID = C.诊治项目ID " & _
'                        "AND A.记录类型 = [1] " & _
'                        "AND D.ID=A.检验标本ID " & _
'                        "AND A.检验项目id=X.项目ID(+) AND (X.仪器ID=D.仪器ID+0 OR X.仪器ID IS NULL OR D.仪器ID IS NULL) " & _
'                        "AND D.ID= [2] Order By B.编码) A"
        '2008-02-13 修改排序 陈东
        mstrSQL = "Select /*+ rule */ Distinct A.标本ID,A.诊疗项目id, A.编码, a.排列序号, a.Id, a.检验项目, a.原始结果, a.上次结果, a.上次时间, a.Cv," & _
                  "a.本次结果,a.计算公式,a.结果类型,a.标志,a.仪器提示,decode(a.仪器审核标识,1,'√',0,'×','') as 仪器审核标识," & vbNewLine & _
                  "Trim(REPLACE(REPLACE(' '||zlGetReference(A.ID,A.标本类型,0,NULL,A.仪器ID),' .','0.'),'～.','～0.')) AS 参考1,a.标本类型,a.结果参考 as 参考, " & _
                  " a.仪器ID,a.标本类别,a.核收时间,a.标本序号,a.检验备注,a.结果范围,a.固定项目,Nvl(X.小数位数,2) AS 小数," & _
                  " Trim(REPLACE(REPLACE(' '||Zl_Get_Reference(4,A.ID,A.标本类型,0,NULL,A.仪器ID),' .','0.'),'～.','～0.')) AS 警戒上限, " & _
                  " Trim(REPLACE(REPLACE(' '||Zl_Get_Reference(3,A.ID,A.标本类型,0,NULL,A.仪器ID),' .','0.'),'～.','～0.')) AS 警戒下限, " & _
                  "a.单位" & _
                  ",Rownum As 序号,A.OD,A.CUTOFF,A.COV,a.酶标板Id,a.变异报警,a.变异警示,lpad(编码,4,'0') as 排序 From (SELECT D.id as 标本ID,A.诊疗项目ID,A.排列序号,B.ID," & _
                        IIf(chkChina.Value = 1, "B.中文名||DECODE(C.缩写,NULL,'','('||C.缩写||')') AS 检验项目", "C.缩写 as 检验项目 ") & vbNewLine & _
                        ",decode(c.排列序号,Null,nvl(h.编码,b.编码),c.排列序号) as 编码," & _
                        "A.原始结果," & _
                        "'' As 上次结果,'' as 上次时间 ,''As CV," & _
                        "A.检验结果 As 本次结果," & _
                        "C.计算公式," & _
                        "C.结果类型," & _
                        "DECODE(A.结果标志,3,'↑',2,'↓',1,'',4,'异常',5,'↓↓',6,'↑↑','') AS 标志," & _
                        "Nvl(D.仪器ID,-1) As 仪器ID,Nvl(D.标本类别,0) As 标本类别,D.核收时间,D.标本序号,D.检验备注,C.结果范围,0 As 固定项目," & _
                        "C.警戒上限,C.警戒下限,C.单位,D.标本类型 ,A.OD,A.CUTOFF,A.SCO as COV,a.酶标板ID,c.变异报警率 as 变异报警,c.变异警示率 as 变异警示,a.结果参考 ,e.仪器是否审核 as  仪器审核标识,e.审核内容  as 仪器提示 " & _
                        ",Zl_To_Number(Zl_Get_Reference(1, a.检验项目id, d.标本类型, Decode(d.性别, '男', 1, '女', 2, 0), d.出生日期,d.仪器id, d.年龄,d.申请科室ID)) as 参考ID " & vbNewLine & _
                    "FROM 检验普通结果 A,诊治所见项目 B,检验项目 C,检验标本记录 D,诊疗项目目录 H ,检验流水线指标  e " & _
                    "WHERE A.诊疗项目ID=H.ID(+) And A.检验项目id = B.ID " & _
                        "AND B.ID = C.诊治项目ID and  a.检验标本id=e.标本id(+) and  a.检验项目id = e.项目id(+)  " & _
                        "AND A.记录类型 = [1] " & _
                        "AND D.ID=A.检验标本ID " & _
                        "AND D.ID= [2] ) A,检验仪器项目 X,检验项目参考 F where a.仪器ID=X.仪器ID(+) and A.ID=X.项目ID(+) and A.参考id=f.id(+)"
    End If
    
    If blnMoved Then
        strSQLbak = mstrSQL
        strSQLbak = Replace(strSQLbak, "检验标本记录", "H检验标本记录")
        strSQLbak = Replace(strSQLbak, "检验普通结果", "H检验普通结果")
        strSQLbak = Replace(strSQLbak, "检验申请项目", "H检验申请项目")
        mstrSQL = mstrSQL & " Union ALL " & strSQLbak
    End If
    
    mstrSQL = mstrSQL & " Order by 排序,排列序号 "
    
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mbytRedoNumber, IIf(mlngKey = 0, -1, mlngKey), lngPatientID, lngAdvice)
    
    If rs.BOF = False Then
        '初始标本信息
        mDeviceID = rs("仪器ID")
        Me.txtComment.Text = Nvl(rs("检验备注"))
        
        Vsf.TextMatrix(0, 0) = "#"
'        Call FillGrid_UQ(Vsf, rs, Array("", "", "", ""))
        Call ReadVsf(rs, Array("", "", "", ""))
        Vsf.TextMatrix(0, 0) = ""
        Vsf.Cell(flexcpBackColor, 1, 0, Vsf.Rows - 1, 0) = &HFDD6C6
        rs.MoveFirst
        
        Call FormatVsfCell(Vsf, mCol.检验结果, "0.0######", IIf(Nvl(rs("结果类型"), 0) = 1, 0, 1), _
                IIf(mDeviceID > 0, mCol.小数, -1))
                
        Call FormatVsfCell(Vsf, mCol.原始结果, "0.0######", IIf(Nvl(rs("结果类型"), 0) = 1, 0, 1), _
                IIf(mDeviceID > 0, mCol.小数, -1))
        
'        If chkLast.Value Then LoadLastValue
        '--每次都读出历史结果
        LoadLastValue
    Else
        mDeviceID = -1
        Me.txtComment.Text = ""
        ResetVsf Vsf
    End If
    
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        For mlngLoop = 1 To Vsf.Rows - 1
            Call ApplyResultColor(Vsf, mlngLoop, mCol.检验结果 + intCol * mintColCount, _
                Decode(Vsf.TextMatrix(mlngLoop, mCol.结果标志 + intCol * mintColCount), "↑", 3, "↓", 2, "异常", 4, "↑↑", 6, "↓↓", 5, 1))
        Next
    Next
    
    '写入诊断信息
    Me.txtDiagnose.Text = ""
    gstrSql = "Select b.医嘱id, b.项目, b.排列, b.内容" & vbNewLine & _
                "From 检验标本记录 a, 病人医嘱附件 b" & vbNewLine & _
                "Where a.医嘱id = b.医嘱id and a.ID = [1] " & vbNewLine & _
                "Order By 医嘱id, 排列"
    Set rs = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey)
    
    Do Until rs.EOF
        strTmp = strTmp & Nvl(rs("项目")) & ":" & Replace(Nvl(rs("内容")), vbCrLf, vbCrLf & "    ") & vbCrLf
        rs.MoveNext
    Loop
    Me.txtDiagnose.Text = strTmp
    
    If mbytRedoNumber > 0 Then
        gstrSql = "select 检验备注 from 检验普通结果 where 检验标本id = [1] and 记录类型 = [2] "
        Set rs = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey, mbytRedoNumber)
        If rs.EOF = False Then
            txtComment.Text = rs("检验备注") & ""
        End If
    End If
    
    ReadData = True
    
    Exit Function
    
ErrHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub FormatVsfCell(objVsf As Object, ByVal lngCol As Long, ByVal strFormat As String, Optional ByVal iType As Integer = -1, Optional ByVal iTypeCol As Integer = -1)
    'iType：需格式化的数据类型
    '  0：数字、1：字符、2：日期、3：逻辑、-1：不限（缺省）
    'iTypeCol：小数位数的存储字段序号
    Dim lngLoop As Long
    Dim intColCount As Integer
    Dim intCol As Integer
    
    intColCount = GetColCount(objVsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        For lngLoop = 1 To objVsf.Rows - 1
            If iType = 0 And IsNumeric("-" & objVsf.TextMatrix(lngLoop, lngCol + intCol * mintColCount)) And iTypeCol <> -1 Then
                If InStr(UCase(objVsf.TextMatrix(lngLoop, lngCol + intCol * mintColCount)), "E") = 0 Then
                    objVsf.TextMatrix(lngLoop, lngCol + intCol * mintColCount) = Format(objVsf.TextMatrix(lngLoop, lngCol + intCol * mintColCount), _
                        IIf(Val(objVsf.TextMatrix(lngLoop, iTypeCol + intCol * mintColCount)) = 0, "#0", "0." & String(Val(objVsf.TextMatrix(lngLoop, iTypeCol + intCol * mintColCount)), "0")))
                End If
            Else
                '曾超修改：半定量和定性的不需要格式化(有科学计数法)
    '            If IsNumeric("-" & objVsf.TextMatrix(lngLoop, lngCol)) Then objVsf.TextMatrix(lngLoop, lngCol) = Format(objVsf.TextMatrix(lngLoop, lngCol), strFormat)
            End If
        Next
    Next
End Sub

Public Function zlRefresh(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：显示数据
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------
    mblnLoadHistory = False ' IIf(mlngKey <> lngKey, False, mblnLoadHistory)
    mlngKey = lngKey
    fraComment.Tag = ""
    Call Form_Resize
'    SetEditState False
    '初始仪器列表
    If ReadData = False Then Exit Function
    
    zlRefresh = True
End Function
Public Function zlRefreshPatient(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：显示数据（按病人)
    '参数：lngkey 病人ID
    '返回：
    '------------------------------------------------------------------------------------------------------
    mLngPatientID = lngKey
    fraComment.Tag = "不显示"
    zlRefreshPatient = True
    Call ReadPatient
End Function

Public Function ZlEditStart(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：编辑数据
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim intColCount As Integer
    
    SetEditState True
    
    If mlngKey <> lngKey Then
        mlngKey = lngKey
        If ReadData = False Then Exit Function
    End If
    
    mblnChangeEdit = False
    ZlEditStart = True
    intColCount = GetColCount(Vsf.Col)
    With Vsf
        If .Col = mCol.检验结果 + intColCount * mintColCount Or .Col = mCol.结果标志 + intColCount * mintColCount _
           Or .Col = mCol.od + intColCount * mintColCount Then
            Select Case .Col
                Case mCol.检验结果 + intColCount * mintColCount
                    .EditMode(.Col) = 1
                Case mCol.结果标志 + intColCount * mintColCount
                    If .TextMatrix(.Row, .Col) <> "↓↓" And .TextMatrix(.Row, .Col) <> "↑↑" Then .EditMode(.Col) = 1
                Case mCol.od + intColCount * mintColCount
                    .EditMode(.Col) = 1
            End Select
        Else
            .Col = mCol.检验结果 + intColCount * mintColCount
        End If
        '如果是仪器标本，则查找无结果的指标开始填写
'        If mDeviceID > 0 And Not mblnEvent Then
'            For i = 1 To .Rows - 1
'                If Len(Trim(.TextMatrix(i, mCol.检验结果))) = 0 Then Exit For
'            Next
'            If i <= .Rows - 1 And i >= 1 Then
'                .Row = i
'                .ShowCell .Row, mCol.检验结果
'            End If
'        End If
        mblnEvent = False
        
        .SetFocus
    End With
End Function

Public Function ZlSave() As Boolean
    If SaveData() = False Then Exit Function

    ZlSave = True
End Function

Public Function ZlCancel() As Boolean
    '提示是否保存
    SetEditState False
    '重新显示报告结果
    mblnLoadHistory = False
    
    Vsf.EditMode(mCol.检验结果) = 1: Vsf.EditMode(mCol.结果标志) = 0
    Vsf.Rows = 2
    Vsf.Cols = mintColCount
    Vsf.Cell(flexcpText, 1, 0, 1, Vsf.Cols - 1) = ""
    Vsf.Cell(flexcpData, 1, 0, Vsf.Rows - 1, Vsf.Cols - 1) = 0
    'Call ReadData
    
    ZlCancel = True
End Function

Public Function ZlClearForm() As Boolean
    '清空结果
    mblnLoadHistory = False
    mlngKey = 0
    With sbrInfo
        .Panels(1).Text = "报告人："
        .Panels(2).Text = "报告时间："
        .Panels(3).Text = "审核人："
        .Panels(4).Text = "审核时间："
    End With
    Me.txtComment = ""
    ResetVsf Vsf
End Function

Private Sub SetEditState(ByVal blnEdit As Boolean)
    Dim intColCount As Integer
    mblnEdit = blnEdit
'    vsf.Body.Editable = IIf(blnEdit, flexEDKbdMouse, flexEDNone)
    txtComment.Locked = Not blnEdit
    Me.lvwSelect.Visible = blnEdit
    intColCount = GetColCount(Me.Vsf.Col)
    If Me.lvwSelect.Visible Then
        ShowValue 1, Val(Vsf.TextMatrix(Vsf.Row, mCol.结果类型 + intColCount * mintColCount)), Vsf.Cell(flexcpData, Vsf.Row, intColCount * mintColCount, Vsf.Row, intColCount * mintColCount)
    End If
    Call Form_Resize
End Sub

Private Function SaveData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim strNow As String
    Dim bytResultFlag As Byte, mlngLoop As Long
    Dim intColCount As Integer
    Dim intCol As Integer
    Dim intLoop As Integer

    Dim strsql() As String

    Dim strTmp As String, rsTmp As ADODB.Recordset
    If Vsf.Rows > 1 Then
        Vsf.Row = Vsf.Row - 1
        Vsf.Row = Vsf.Row + 1
    Else
        Vsf.Row = Vsf.Row + 1
        Vsf.Row = Vsf.Row - 1
    End If
    If Not mblnChangeEdit Then SaveData = True: Exit Function

    On Error GoTo ErrHand
    ReDim strsql(1 To 1)
    '读取检验时间
    strNow = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    '读取检验标本记录信息
    strTmp = "Select 医嘱id ,采样时间 ,采样人 , 申请时间 ,nvl(紧急,0) as 紧急, " & _
        "样本条码 , 申请类型 , 执行科室id , 检验人 ,检验时间 ,标本序号,审核人 " & _
        "From 检验标本记录 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, Me.Caption, mlngKey)
    If rsTmp.EOF Then
        SaveData = False
        Exit Function
    Else
        If rsTmp!审核人 & "" <> "" Then
            MsgBox "该标本已被其他用户审核！", vbInformation, gstrSysName
            SaveData = False
            mblnChangeEdit = False
            Call frmLabMain.zlRefreshData
            Exit Function
        End If
    End If
    
    Vsf.SetFocus
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        For mlngLoop = 1 To Vsf.Rows - 1
            'If Val(vsf.RowData(mlngLoop)) > 0 Then
            If Val(Vsf.Cell(flexcpData, mlngLoop, intCol * mintColCount, mlngLoop, intCol * mintColCount)) > 0 Then
                bytResultFlag = 0
                If Trim(Vsf.TextMatrix(mlngLoop, mCol.结果标志 + intCol * mintColCount)) <> "" Then
                    bytResultFlag = Decode(Vsf.TextMatrix(mlngLoop, mCol.结果标志 + intCol * mintColCount), "↑", 3, "↓", 2, "异常", 4, "↓↓", 5, "↑↑", 6, 1)
                End If
                '处理只输入了结果马上点保存的情况
                If mlngLoop = Vsf.Row And mCol.检验结果 + intCol * mintColCount = Vsf.Col And Vsf.EditText <> "" Then
                    Vsf.TextMatrix(mlngLoop, mCol.检验结果 + intCol * mintColCount) = Vsf.EditText
                End If
                strsql(ReDimArray(strsql)) = "ZL_检验标本记录_报告填写(" & CLng(Vsf.TextMatrix(mlngLoop, mCol.标本ID + intCol * mintColCount)) & "," & _
                Val(Vsf.Cell(flexcpData, mlngLoop, intCol * mintColCount, mlngLoop, intCol * mintColCount)) & "," & _
                mbytRedoNumber & ",'" & Vsf.TextMatrix(mlngLoop, mCol.检验结果 + intCol * mintColCount) & "',TO_DATE('" & strNow & "','yyyy-mm-dd hh24:mi:ss')," & _
                IIf(bytResultFlag = 0, "NULL", bytResultFlag) & ",'" & Vsf.TextMatrix(mlngLoop, mCol.结果参考 + intCol * mintColCount) & "',1,NULL,0," & IIf(intCol = 0 And mlngLoop = 1, 1, 0) & _
                ",'" & Vsf.TextMatrix(mlngLoop, mCol.原始结果 + intCol * mintColCount) & "'," & Vsf.TextMatrix(mlngLoop, mCol.诊疗项目id + intCol * mintColCount) & _
                "," & IIf(Vsf.TextMatrix(mlngLoop, mCol.排列序号 + intCol * mintColCount) = "", Vsf.TextMatrix(mlngLoop, intCol * mintColCount), Vsf.TextMatrix(mlngLoop, mCol.排列序号 + intCol * mintColCount)) & _
                ",'" & Vsf.TextMatrix(mlngLoop, mCol.od + intCol * mintColCount) & _
                "','" & Vsf.TextMatrix(mlngLoop, mCol.CUTOFF + intCol * mintColCount) & "','" & Vsf.TextMatrix(mlngLoop, mCol.COV + intCol * mintColCount) & _
                "'," & IIf(Vsf.TextMatrix(mlngLoop, mCol.酶标板ID + intCol * mintColCount) = "", "Null", Vsf.TextMatrix(mlngLoop, mCol.酶标板ID + intCol * mintColCount)) & _
                ",'" & txtComment & "','" & UserInfo.姓名 & "',1)"
                intLoop = intLoop + 1
            End If
        Next
    Next
    
    If intLoop = 0 Then
        strsql(ReDimArray(strsql)) = "Zl_检验普通结果_Delete(" & mlngKey & ")"
    End If
    
    
    '重新计算计算项目
    strsql(ReDimArray(strsql)) = "Zl_重新计算结果_Cale(" & mlngKey & ")"
    
    blnTran = True

    gcnOracle.BeginTrans
    For mlngLoop = 1 To UBound(strsql)
        If strsql(mlngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(Replace(strsql(mlngLoop), ",,", ",Null,"), Me.Caption)
    Next
    gcnOracle.CommitTrans
    gstrSql = "ZL_检验标本记录_报告选择(" & mlngKey & "," & mbytRedoNumber & ",'" & txtComment & "')"
    zlDatabase.ExecuteProcedure gstrSql, gstrSysName
    
    If Signature(mlngKey, gstrDBUser, "报告") = False Then
        Exit Function
    End If

    

    SaveData = True
    mblnEdit = False
    Exit Function
ErrHand:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    
End Function

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim p As POINTAPI
    Dim intColCount As Integer
    Dim lngItemID As Long
    
    If Control.ID = 100 Then
        
        intColCount = GetColCount(Me.Vsf.Col)
        lngItemID = Val(Vsf.Cell(flexcpData, Me.Vsf.Row, intColCount * mintColCount, Me.Vsf.Row, intColCount * mintColCount))
        Call GetCursorPos(p)
        With frmLisStationWriteInfo
            .Top = p.Y * Screen.TwipsPerPixelY
            .Left = p.X * Screen.TwipsPerPixelX
            .ShowME Me, lngItemID
        End With
    Else
        mbytRedoNumber = Control.ID
        mSelectRedo = True
        mblnLoadHistory = False
        ReadData
    End If
End Sub

Private Sub cbrthis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control.ID = mbytRedoNumber + 1 Then Control.Checked = True
End Sub

Private Sub chkChina_Click()
    Dim intColCount As Integer, intCol As Integer
    intColCount = GetColCount(Vsf.Cols)

    mblnLoadHistory = False
    If fraComment.Tag = "" Then
        Call ReadData
    Else
        ReadPatient
    End If
End Sub

Private Sub chkLast_Click()
    Dim intColCount As Integer, intCol As Integer
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
        
    For intCol = 0 To intColCount - 1
        Vsf.Body.ColWidth(mCol.上次结果 + intCol * mintColCount) = IIf(chkLast.Value, 900, 0)
        Vsf.Body.ColWidth(mCol.上次时间 + intCol * mintColCount) = IIf(chkLast.Value, 1000, 0)
    Next
'    vsf.Body.ColWidth(mCol.CV) = IIf(chkLast.Value, 400, 0)
    
    If chkLast.Value Then LoadLastValue
End Sub

Private Sub chkLastDate_Click()

End Sub

Private Sub chkMB_Click()
    Dim intColCount As Integer, intCol As Integer
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        Vsf.Body.ColWidth(mCol.od + intCol * mintColCount) = IIf(chkMB.Value, 700, 0)
        Vsf.Body.ColWidth(mCol.CUTOFF + intCol * mintColCount) = IIf(chkMB.Value, 700, 0)
        Vsf.Body.ColWidth(mCol.COV + intCol * mintColCount) = IIf(chkMB.Value, 700, 0)
    Next
End Sub

Private Sub chkOriginal_Click()
    Dim intColCount As Integer, intCol As Integer
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        Vsf.Body.ColWidth(mCol.原始结果 + intCol * mintColCount) = IIf(chkOriginal.Value, 900, 0)
    Next
End Sub

Private Sub chkReferrence_Click()
    Dim intColCount As Integer, intCol As Integer
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        Vsf.Body.ColWidth(mCol.结果参考 + intCol * mintColCount) = IIf(chkReferrence.Value, 1300, 0)
    Next
End Sub

Private Sub chkSign_Click()
    Dim intColCount As Integer, intCol As Integer
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        Vsf.Body.ColWidth(mCol.结果标志 + intCol * mintColCount) = IIf(chkSign.Value, 450, 0)
    Next
End Sub

Private Sub chkUnit_Click()
    Dim intColCount As Integer, intCol As Integer
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        Vsf.Body.ColWidth(mCol.单位 + intCol * mintColCount) = IIf(chkUnit.Value, 1000, 0)
    Next
End Sub

Private Sub chkYiQiBiaoShi_Click()
    Dim intColCount As Integer, intCol As Integer
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        Vsf.Body.ColWidth(mCol.仪器审核标识 + intCol * mintColCount) = IIf(chkYiQiBiaoShi.Value, 1200, 0)
    Next
End Sub

Private Sub chkYiQiTiShi_Click()
    Dim intColCount As Integer, intCol As Integer
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        Vsf.Body.ColWidth(mCol.仪器提示 + intCol * mintColCount) = IIf(chkYiQiTiShi.Value, 1000, 0)
    Next
End Sub

Private Sub Form_Load()
    With Vsf
        .Body.BackColor = &H80000005
        .Body.Appearance = flex3DLight
        .Body.BorderStyle = flexBorderFlat
        .Body.BackColorFixed = &HFDD6C6
        .Body.GridLinesFixed = flexGridFlat
        .Body.RowHeightMin = 300
        .Body.Editable = flexEDKbdMouse
        
        .Cols = 0
        .NewColumn "", 300, 7
        .NewColumn "检验项目", 2100, 1
        .NewColumn "原始结果", 0, 1
        .NewColumn "本次结果", 1200, 1, , 1
        .NewColumn "单位", 1000, 1
        .NewColumn "CV", 0, 1
        .NewColumn "标志", 450, 1
        .NewColumn "上次结果", 0, 1
        .NewColumn "上次时间", 0, 1
        .NewColumn "参考", 1300, 1
        .NewColumn "结果类型", 0, 1
        .NewColumn "仪器id", 0, 1
        .NewColumn "计算公式", 0, 1
        .NewColumn "结果范围", 0, 1
        .NewColumn "固定项目", 0, 1
        .NewColumn "小数", 0, 1
        .NewColumn "警戒上限", 0, 1
        .NewColumn "警戒下限", 0, 1
        .NewColumn "诊疗项目ID", 0, 1
        .NewColumn "排列序号", 0, 1
        .NewColumn "标本ID", 0, 1
        .NewColumn "OD", 700, 1, , 1
        .NewColumn "CUTOFF", 700, 1
        .NewColumn "COV", 700, 1
        .NewColumn "酶标板ID", 0, 1
        .NewColumn "变异报警", 0, 1
        .NewColumn "变异警示", 0, 1
        .NewColumn "仪器提示", 1000, 1
        .NewColumn "仪器审核标识", 1200, 1
        .FixedCols = 0
    End With
    lvwSelect.Tag = 1 '默认选择指标结果
    mblnLoadHistory = False
    
    If mblnPatientFind = False Then
        '取报告选项
        chkOriginal.Value = Val(zlDatabase.GetPara("frmLisStationWrite_查看原始结果", 100, 1208, 0))
        chkLast.Value = Val(zlDatabase.GetPara("frmLisStationWrite_查看上次结果", 100, 1208, 0))
        chkSign.Value = Val(zlDatabase.GetPara("frmLisStationWrite_查看标志", 100, 1208, 1))
        chkUnit.Value = Val(zlDatabase.GetPara("frmLisStationWrite_查看单位", 100, 1208, 1))
        chkReferrence.Value = Val(zlDatabase.GetPara("frmLisStationWrite_查看参考", 100, 1208, 1))
        chkMB.Value = Val(zlDatabase.GetPara("frmLisStationWrite_查看酶标", 100, 1208, 1))
        chkChina.Value = Val(zlDatabase.GetPara("frmLisStationWrite_查看中文", 100, 1208, 1))
        chkYiQiTiShi.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_仪器提示", 100, 1208, 1), 1, 0)
        chkYiQiBiaoShi.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_仪器审核标识", 100, 1208, 1), 1, 0)
    Else
        chkChina.Value = 1
        chkSign.Value = 1
        
    End If
    Vsf.Body.ColWidth(mCol.上次结果) = IIf(chkLast.Value, 900, 0)
    Vsf.Body.ColWidth(mCol.原始结果) = IIf(chkOriginal.Value, 900, 0)
    Vsf.Body.ColWidth(mCol.od) = IIf(chkMB.Value, 700, 0)
    Vsf.Body.ColWidth(mCol.CUTOFF) = IIf(chkMB.Value, 700, 0)
    Vsf.Body.ColWidth(mCol.COV) = IIf(chkMB.Value, 700, 0)
    Vsf.Body.ColWidth(mCol.检验项目) = IIf(chkChina.Value, 1000, 2100)
    Vsf.Body.ColWidth(mCol.仪器提示) = IIf(chkYiQiTiShi.Value, 1000, 0)
    Vsf.Body.ColWidth(mCol.仪器审核标识) = IIf(chkYiQiBiaoShi.Value, 1200, 0)
    SetEditState False
    
    '读取颜色
    lngReferenceLow = Val(zlDatabase.GetPara("参考颜色_偏低", 100, 1208, 0))
    If lngReferenceLow = 0 Then lngReferenceLow = 8454143
    lblLow.BackColor = lngReferenceLow
    lngReferenceHigh = Val(zlDatabase.GetPara("参考颜色_偏高", 100, 1208, 0))
    If lngReferenceHigh = 0 Then lngReferenceHigh = 8438015
    lblHigh.BackColor = lngReferenceHigh
    lngReferenceExigency = Val(zlDatabase.GetPara("参考颜色_警示", 100, 1208, 0))
    If lngReferenceExigency = 0 Then lngReferenceExigency = 16576
    lblExigency.BackColor = lngReferenceExigency
    
    
    Call RestoreFlexState(Vsf, Me.Name)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With fraComment
        .Left = 0
        .Top = Me.ScaleHeight - Me.sbrInfo.Height - .Height - 30
        .Width = Me.ScaleWidth - .Left
    End With
    With txtComment
'        .Width = Me.fraComment.Width - .Left - txtDiagnose.Width - Me.Label2.Width
        .Width = Me.fraComment.Width / 2
        .Height = fraComment.Height - 20
    End With
    
    With Me.Label2
        .Left = Me.txtComment.Left + Me.txtComment.Width + 20
    End With
    
    With Me.txtDiagnose
        .Left = Me.Label2.Left + Me.Label2.Width + 20
        .Width = fraComment.Width - txtComment.Left - txtComment.Width - 400
        .Height = txtComment.Height
    End With
        
    With fraTitle
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
    End With
    
    With lvwSelect
        .Left = Me.ScaleWidth - .Width - 30
        .Top = fraTitle.Top + fraTitle.Height + 30
        .Height = fraComment.Top - .Top + 30
    End With
    If fraComment.Tag = "" Then
        fraComment.Visible = zlDatabase.GetPara("显示检验备注", 100, 1208, True)
    Else
        fraComment.Visible = False
    End If
    
    With Vsf
        .Left = -15
        .Top = fraTitle.Top + fraTitle.Height + 30
        .Width = IIf(Me.lvwSelect.Visible, Me.lvwSelect.Left, Me.ScaleWidth) - 30 - .Left
        If fraComment.Visible Then
            .Height = fraComment.Top - .Top + 30
        Else
            .Height = fraComment.Top + fraComment.Height - .Top + 30
        End If
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveFlexState(Vsf, Me.Name)
    
    zlDatabase.SetPara "frmLisStationWrite_查看原始结果", Me.chkOriginal.Value, 100, 1208
    zlDatabase.SetPara "frmLisStationWrite_查看上次结果", Me.chkLast.Value, 100, 1208
    zlDatabase.SetPara "frmLisStationWrite_查看标志", Me.chkSign.Value, 100, 1208
    zlDatabase.SetPara "frmLisStationWrite_查看单位", Me.chkUnit.Value, 100, 1208
    zlDatabase.SetPara "frmLisStationWrite_查看参考", Me.chkReferrence.Value, 100, 1208
    zlDatabase.SetPara "frmLisStationWrite_查看酶标", Me.chkMB.Value, 100, 1208
    zlDatabase.SetPara "frmLisStationWrite_查看中文", Me.chkChina.Value, 100, 1208
    fraComment.Tag = ""
    '退出窗体时还原变量
    mblnEdit = False
End Sub

Private Sub lblExigency_DblClick()
    CommDialog.ShowColor
    If CommDialog.COLOR <> 0 Then
        lblExigency.BackColor = CommDialog.COLOR
        Call zlDatabase.SetPara("参考颜色_警示", CommDialog.COLOR, 100, 1208)
    End If
End Sub

Private Sub lblHigh_DblClick()
    CommDialog.ShowColor
    If CommDialog.COLOR <> 0 Then
        lblHigh.BackColor = CommDialog.COLOR
        Call zlDatabase.SetPara("参考颜色_偏高", CommDialog.COLOR, 100, 1208)
    End If
End Sub

Private Sub lblLow_DblClick()
    CommDialog.ShowColor
    If CommDialog.COLOR <> 0 Then
        lblLow.BackColor = CommDialog.COLOR
        Call zlDatabase.SetPara("参考颜色_偏低", CommDialog.COLOR, 100, 1208)
    End If
End Sub

Private Sub lvwSelect_DblClick()
    Dim intColCount As Integer
    
    If lvwSelect.SelectedItem Is Nothing Then Exit Sub
    If Not mblnEdit Then Exit Sub
    
    On Error GoTo errH
    
    intColCount = GetColCount(Vsf.Col)
    
    Select Case Val(lvwSelect.Tag)
        Case 1 '选择结果
            If Val(Vsf.Cell(flexcpData, Vsf.Row, intColCount * mintColCount, Vsf.Row, intColCount * mintColCount)) > 0 Then
                Vsf.TextMatrix(Vsf.Row, mCol.检验结果 + intColCount * mintColCount) = lvwSelect.SelectedItem.Text
                '产生缺省的结果标志
                Vsf.TextMatrix(Vsf.Row, mCol.结果标志 + intColCount * mintColCount) = CalcDefaultFlag(Trim(Vsf.TextMatrix(Vsf.Row, mCol.检验结果 + intColCount * mintColCount)), _
                Trim(Vsf.TextMatrix(Vsf.Row, mCol.结果参考 + intColCount * mintColCount)), Val(Vsf.TextMatrix(Vsf.Row, mCol.结果类型 + intColCount * mintColCount)), _
                    Vsf.TextMatrix(Vsf.Row, mCol.警戒下限 + intColCount * mintColCount), Vsf.TextMatrix(Vsf.Row, mCol.警戒上限 + intColCount * mintColCount))
                
                '根据结果应用颜色标志
                Call ApplyResultColor(Vsf, Vsf.Row, mCol.检验结果 + intColCount * mintColCount, _
                    Decode(Vsf.TextMatrix(Vsf.Row, mCol.结果标志 + intColCount * mintColCount), "↑", 3, "↓", 2, "异常", 4, "↓↓", 5, "↑↑", 6, 1))
                
                Vsf.SetFocus
                
                mblnChangeEdit = True
            End If
        Case 2 '选择备注
            Me.txtComment.SelText = lvwSelect.SelectedItem.Text
            
            mblnChangeEdit = True
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub picFilter_Click()
    Dim p As POINTAPI
    Dim blnOriginal As Boolean
    Dim blnLast As Boolean
    Dim blnChina As Boolean
    Dim blnSign As Boolean
    If Me.picFilter.Tag = "" Then
        Me.picFilter.Tag = "True"
    Else
        Me.picFilter.Tag = ""
    End If
    Call GetCursorPos(p)
    With frmLabMainSizer
        .Top = p.Y * Screen.TwipsPerPixelY
        .Left = p.X * Screen.TwipsPerPixelX
        .ShowME Me, "frmLisStationWrite", IIf(Me.picFilter.Tag = "", True, False)
    End With

    If mblnPatientFind = False Then
'        '取报告选项
        chkOriginal.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_查看原始结果", 100, 1208, 0), 1, 0)
        chkLast.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_查看上次结果", 100, 1208, 0), 1, 0)
        chkSign.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_查看标志", 100, 1208, 1), 1, 0)
        chkUnit.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_查看单位", 100, 1208, 1), 1, 0)
        chkReferrence.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_查看参考", 100, 1208, 1), 1, 0)
        chkMB.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_查看酶标", 100, 1208, 1), 1, 0)
        chkChina.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_查看中文", 100, 1208, 1), 1, 0)
        chkYiQiTiShi.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_仪器提示", 100, 1208, 1), 1, 0)
        chkYiQiBiaoShi.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_仪器审核标识", 100, 1208, 1), 1, 0)
    Else
        chkChina.Value = 1
        chkSign.Value = 1
    End If
End Sub

Private Sub picFilter_LostFocus()
    frmLabMainSizer.ShowME Me, "frmLisStationWrite", True
    Me.picFilter.Tag = ""
    If mblnPatientFind = False Then
'        '取报告选项
        chkOriginal.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_查看原始结果", 100, 1208, 0), 1, 0)
        chkLast.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_查看上次结果", 100, 1208, 0), 1, 0)
        chkSign.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_查看标志", 100, 1208, 1), 1, 0)
        chkUnit.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_查看单位", 100, 1208, 1), 1, 0)
        chkReferrence.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_查看参考", 100, 1208, 1), 1, 0)
        chkMB.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_查看酶标", 100, 1208, 1), 1, 0)
        chkChina.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_查看中文", 100, 1208, 1), 1, 0)
        chkYiQiTiShi.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_仪器提示", 100, 1208, 1), 1, 0)
        chkYiQiBiaoShi.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_仪器审核标识", 100, 1208, 1), 1, 0)
    Else
        chkChina.Value = 1
        chkSign.Value = 1
    End If
End Sub

Private Sub picFilter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picFilter.BackColor = &HFFFFFF
    picFilter.BorderStyle = 1
End Sub

Private Sub picFilter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picFilter.BackColor = &H8000000F
    picFilter.BorderStyle = 0
End Sub

Private Sub txtComment_Change()
    mblnChangeEdit = True
End Sub

Private Sub txtComment_GotFocus()
    With txtComment
'        .SelStart = 0
'        .SelLength = Len(.Text)
    End With
    
    If mblnEdit Then ShowValue 2
End Sub

Private Sub txtComment_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    Else
        RaiseEvent StartEdit(False)
        mblnChangeEdit = True
        txtComment.SetFocus
'        txtComment.SelLength = 0
    End If
End Sub

Private Sub vsf_AfterDeleteCell(ByVal Row As Long, ByVal Col As Long)
    mblnChangeEdit = True
End Sub

Private Sub vsf_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    mblnChangeEdit = True
'    Call RenumVsf(vsf, 0)
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strReference As String, strItemIDs As String
    Dim lngCount As Long, mlngLoop As Long
    Dim intColCount As Integer
    Dim intCol As Integer, intCols As Integer
    Dim str复查 As String
    On Error GoTo errH
    
    intColCount = GetColCount(Col)
    Select Case Col
        Case mCol.结果标志 + intColCount * mintColCount
            Select Case Val(Left(Vsf.TextMatrix(Row, mCol.结果标志 + intColCount * mintColCount), 1))
                Case 3
                    Vsf.TextMatrix(Row, Col) = "↑"
                Case 2
                    Vsf.TextMatrix(Row, Col) = "↓"
                Case 4
                    Vsf.TextMatrix(Row, Col) = "异常"
                Case 5
                    Vsf.TextMatrix(Row, Col) = "↓↓"
                Case 6
                    Vsf.TextMatrix(Row, Col) = "↑↑"
                Case Else
                    Vsf.TextMatrix(Row, Col) = ""
            End Select
            Call ApplyResultColor(Vsf, Row, mCol.检验结果 + intColCount * mintColCount, _
                Decode(Vsf.TextMatrix(Row, mCol.结果标志 + intColCount * mintColCount), "↑", 3, "↓", 2, "异常", 4, "↓↓", 5, "↑↑", 6, 1))
        Case mCol.检验结果 + intColCount * mintColCount
            '先处理是否调用模板
            If Left(Vsf.TextMatrix(Row, mCol.检验结果 + intColCount * mintColCount), 1) = "/" Then
                If LoadModel(Mid(Vsf.TextMatrix(Row, mCol.检验结果 + intColCount * mintColCount), 2)) Then
                    mblnChangeEdit = True
                    Exit Sub
                End If
            End If
            '格式化结果
            If Val(Vsf.TextMatrix(Row, mCol.结果类型 + intColCount * mintColCount)) <> 2 And Val(Vsf.TextMatrix(Row, mCol.结果类型 + intColCount * mintColCount)) <> 3 And IsNumeric(Vsf.TextMatrix(Row, mCol.检验结果 + intColCount * mintColCount)) Then
                If InStr(Vsf.TextMatrix(Row, mCol.检验结果 + intColCount * mintColCount), "E") = 0 Then
                    Vsf.TextMatrix(Row, mCol.检验结果 + intColCount * mintColCount) = Format(Vsf.TextMatrix(Row, mCol.检验结果 + intColCount * mintColCount), _
                        "0" & IIf(Val(Vsf.TextMatrix(Row, mCol.小数 + intColCount * mintColCount)) = 0, "", "." & String(Val(Vsf.TextMatrix(Row, mCol.小数 + intColCount * mintColCount)), "0")))
                
                    str复查 = Get复查标记(mlngKey, Vsf.Cell(flexcpData, Row, intColCount * mintColCount, Row, intColCount * mintColCount), Vsf.TextMatrix(Row, mCol.检验结果 + intColCount * mintColCount))
                    If str复查 <> "" Then
                        Vsf.TextMatrix(Row, mCol.检验项目 + intColCount * mintColCount) = Trim(Replace(Vsf.TextMatrix(Row, mCol.检验项目 + intColCount * mintColCount), "需复查", "")) & " " & str复查
                        Vsf.Cell(flexcpForeColor, Row, mCol.检验项目 + intColCount * mintColCount, Row, mCol.检验项目 + intColCount * mintColCount) = COLOR.橙色
                    Else
                        Vsf.TextMatrix(Row, mCol.检验项目 + intColCount * mintColCount) = Trim(Replace(Vsf.TextMatrix(Row, mCol.检验项目 + intColCount * mintColCount), "需复查", ""))
                    End If
                End If
            End If
            
            '产生缺省的结果标志
            Vsf.TextMatrix(Row, mCol.结果标志 + intColCount * mintColCount) = CalcDefaultFlag(Trim(Vsf.TextMatrix(Row, Col)), Trim(Vsf.TextMatrix(Row, mCol.结果参考 + intColCount * mintColCount)), Val(Vsf.TextMatrix(Row, mCol.结果类型 + intColCount * mintColCount)), _
                Vsf.TextMatrix(Row, mCol.警戒下限 + intColCount * mintColCount), Vsf.TextMatrix(Row, mCol.警戒上限 + intColCount * mintColCount), _
                Vsf.Cell(flexcpData, Row, intColCount * mintColCount, Row, intColCount * mintColCount))
            
            '根据结果应用颜色标志
            Call ApplyResultColor(Vsf, Row, mCol.检验结果 + intColCount * mintColCount, _
                Decode(Vsf.TextMatrix(Row, mCol.结果标志 + intColCount * mintColCount), "↑", 3, "↓", 2, "异常", 4, "↓↓", 5, "↑↑", 6, 1))
            
            '自动计算计算项目结果
            intCols = GetColCount(Vsf.Cols)
            If intCols = 0 Then intCols = 1
            For intCol = 0 To intCols - 1
                For mlngLoop = 1 To Vsf.Rows - 1
                    If Trim(Vsf.TextMatrix(mlngLoop, mCol.计算公式 + intCol * mintColCount)) <> "" Then
                        
                        
                        Vsf.TextMatrix(mlngLoop, mCol.检验结果 + intCol * mintColCount) = Format(CalcExpress(Vsf, Trim(Vsf.TextMatrix(mlngLoop, mCol.计算公式 + intCol * mintColCount))), _
                            "0" & IIf(Val(Vsf.TextMatrix(mlngLoop, mCol.小数 + intCol * mintColCount)) = 0, "", "." & String(Val(Vsf.TextMatrix(mlngLoop, mCol.小数 + intCol * mintColCount)), "0")))
                        If CalcExpress(Vsf, Trim(Vsf.TextMatrix(mlngLoop, mCol.计算公式 + intCol * mintColCount))) = "" Then
                            Vsf.TextMatrix(mlngLoop, mCol.检验结果 + intCol * mintColCount) = ""
                        End If
    
                        '产生缺省的结果标志
                        Vsf.TextMatrix(mlngLoop, mCol.结果标志 + intCol * mintColCount) = CalcDefaultFlag(Trim(Vsf.TextMatrix(mlngLoop, mCol.检验结果 + intCol * mintColCount)), Trim(Vsf.TextMatrix(mlngLoop, mCol.结果参考 + intCol * mintColCount)), Val(Vsf.TextMatrix(mlngLoop, mCol.结果类型 + intCol * mintColCount)), _
                            Vsf.TextMatrix(mlngLoop, mCol.警戒下限 + intCol * mintColCount), Vsf.TextMatrix(mlngLoop, mCol.警戒上限 + intCol * mintColCount))
                
                        '根据结果应用颜色标志
                        Call ApplyResultColor(Vsf, mlngLoop, mCol.检验结果 + intCol * mintColCount, _
                            Decode(Vsf.TextMatrix(mlngLoop, mCol.结果标志 + intCol * mintColCount), "↑", 3, "↓", 2, "异常", 4, "↓↓", 5, "↑↑", 6, 1))
                    End If
                Next
            Next
        Case mCol.检验项目 + intColCount * mintColCount
            strItemIDs = GetLabItems(Vsf, mstrType, Vsf.TextMatrix(Row, Col)): gintSelectFocus = 3: ' lvwSelect.SetFocus
'            vsf.SetFocus
            If strItemIDs <> "" Then
                Call AddItems(strItemIDs): Vsf.EditMode(mCol.检验结果 + intColCount * mintColCount) = 1
                Vsf.Col = mCol.检验结果 + intColCount * mintColCount: Vsf.ShowCell Vsf.Row, Vsf.Col
                ShowValue 1, Val(Vsf.TextMatrix(Row, mCol.结果类型 + intColCount * mintColCount)), Vsf.Cell(flexcpData, Row, intColCount * mintColCount, Row, intColCount * mintColCount)
            Else
                Vsf.TextMatrix(Row, Col) = ""
            End If
    End Select

    mblnChangeEdit = True
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim intColCount As Integer
    Dim lngItemID As Long
    
    On Error GoTo errH
     
    
    intColCount = GetColCount(NewCol)
    lngItemID = Val(Vsf.Cell(flexcpData, NewRow, intColCount * mintColCount, NewRow, intColCount * mintColCount))
    frmLisStationWriteInfo.SelectItem lngItemID
    If lngItemID = 0 Then
        Vsf.Col = mCol.检验项目 + intColCount * mintColCount
        Exit Sub
    End If
    If OldRow = NewRow Then Exit Sub
    
    If mblnEdit Then
        ShowValue 1, Val(Vsf.TextMatrix(NewRow, mCol.结果类型 + intColCount * mintColCount)), Val(Vsf.Cell(flexcpData, NewRow, intColCount * mintColCount, NewRow, intColCount * mintColCount))
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
End Sub

Private Sub vsf_BeforeComboList(ByVal NewCol As Long, ComboList As String, Cancel As Boolean)
    '1-正常、2-偏低、3-偏高、4-阳性
    '1:数字,2:文字，3：阴阳型(+-)
    Dim intCol As String
    Dim intColCount As Integer
    
    On Error GoTo errH
    intColCount = GetColCount(NewCol)
    
    If NewCol = mCol.结果标志 + intColCount * mintColCount Then
        Select Case Val(Vsf.TextMatrix(Vsf.Row, mCol.结果类型 + intColCount * mintColCount))
            Case 1  '数字
                ComboList = "1-正常|2-偏低|3-偏高"
            Case 2  '定性
                ComboList = "1-正常|4-异常"
            Case 3  '半定量
                ComboList = "1-正常|2-偏低|3-偏高|4-异常"
        End Select
    ElseIf NewCol = mCol.检验结果 + intColCount * mintColCount Then
        ComboList = "" '"|-|+|--|++|+-"
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsf_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim intColCount As Integer
    Dim lng标本ID As Long
    Dim lng项目ID As Long
    Dim strsql As String
    Dim rsTmp As New ADODB.Recordset
    If Not mblnEdit Then Cancel = True: Exit Sub
    intColCount = GetColCount(Col)
    lng项目ID = Val(Vsf.Cell(flexcpData, Row, intColCount * mintColCount, Row, intColCount * mintColCount))
    lng标本ID = Val(Vsf.TextMatrix(Row, mCol.标本ID + intColCount * mintColCount))
    
    strsql = "Select Distinct C.报告项目id" & vbNewLine & _
            "From 检验标本记录 A, 检验项目分布 B, 检验报告项目 C, 病人医嘱记录 D" & vbNewLine & _
            "Where A.Id = B.标本id And B.医嘱id = D.相关id And D.诊疗项目id = C.诊疗项目id And A.Id = [1] And C.报告项目id = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, lng标本ID, lng项目ID)
    
    If rsTmp.EOF = False Then Cancel = True
    
'    If Val(vsf.TextMatrix(Row, mCol.固定项目 + intColCount * mintColCount)) = 1 Then Cancel = True
End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Dim intColCount As Integer
    If Not mblnEdit Then Cancel = True: Exit Sub
    intColCount = GetColCount(Col)
    If Val(Vsf.Cell(flexcpData, Row, intColCount * mintColCount, Row, intColCount * mintColCount)) = 0 Then Cancel = True
End Sub

Private Sub Vsf_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Dim intCol As Integer
    Dim intColCount As Integer
    If Not mblnEdit Then Exit Sub
    '如果是文字型的
    On Error Resume Next
    
    intColCount = GetColCount(NewCol)
    
    
    If Not mblnEdit Then Exit Sub
    
    If OldCol <> NewCol And mblnEdit Then
        Vsf.EditMode(OldCol) = 0
        Select Case NewCol
            Case mCol.检验结果 + intColCount * mintColCount
                Vsf.EditMode(NewCol) = 1
            Case mCol.od + intColCount * mintColCount
                Vsf.EditMode(NewCol) = 1
            Case mCol.结果标志 + intColCount * mintColCount
                If Vsf.TextMatrix(NewRow, NewCol) <> "↓↓" And Vsf.TextMatrix(NewRow, NewCol) <> "↑↑" Then
                    Vsf.EditMode(NewCol) = 1
                Else
                    Vsf.EditMode(mCol.检验结果 + intColCount * mintColCount) = 1
                End If
            Case mCol.od + intColCount * mintColCount, mCol.CUTOFF + intColCount * mintColCount, mCol.COV + intColCount * mintColCount
                '允许酶标的几个值修改
                Vsf.EditMode(NewCol) = 1
            Case Else
                Vsf.EditMode(mCol.检验结果 + intColCount * mintColCount) = 1
        End Select
    End If
    
    If NewCol = mCol.检验结果 + intColCount * mintColCount Then
        Select Case Val(Vsf.TextMatrix(NewRow, mCol.结果类型 + intColCount * mintColCount))
            Case 3
                Vsf.ComboList(mCol.检验结果 + intColCount * mintColCount) = " "
                Vsf.VsfComboList = "" '"|-|+|--|++|+-"
'                If Len(Trim(vsf.TextMatrix(NewRow, mCol.检验结果 + intColCount * mintColCount))) = 0 Then vsf.TextMatrix(NewRow, mCol.检验结果 + intColCount * mintColCount) = "-"
            Case Else
                Vsf.ComboList(mCol.检验结果 + intColCount * mintColCount) = ""
                Vsf.VsfComboList = ""
        End Select
    ElseIf NewCol = mCol.检验项目 + intColCount * mintColCount Then
        If Val(Vsf.Cell(flexcpData, NewRow, intColCount * mintColCount, NewRow, intColCount * mintColCount)) = 0 Then
            Vsf.EditMode(mCol.检验项目 + intColCount * mintColCount) = 1
            Vsf.ComboList(mCol.检验项目 + intColCount * mintColCount) = "..."
            Vsf.VsfComboList = "..."
        Else
            Vsf.EditMode(mCol.检验项目 + intColCount * mintColCount) = 0
            Vsf.ComboList(mCol.检验项目 + intColCount * mintColCount) = ""
            Vsf.VsfComboList = ""
        End If
    ElseIf NewCol = mCol.结果标志 + intColCount * mintColCount Then
        Vsf.ComboList(NewCol) = " "
        
        Select Case Val(Vsf.TextMatrix(NewRow, mCol.结果类型 + intColCount * mintColCount))
            Case 1  '数字
                Vsf.VsfComboList = "1-正常|2-偏低|3-偏高"
            Case 2  '定性
                Vsf.VsfComboList = "1-正常|4-异常"
            Case 3  '半定量
                Vsf.VsfComboList = "1-正常|2-偏低|3-偏高|4-异常"
        End Select
    End If
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strItemIDs As String
    Dim intColCount As Integer
    
    On Error GoTo errH
    intColCount = GetColCount(Col)
    Select Case Col
        Case mCol.检验项目 + intColCount * mintColCount
            strItemIDs = GetLabItems(Vsf, mstrType): gintSelectFocus = 3: lvwSelect.SetFocus
            Vsf.SetFocus
            If strItemIDs <> "" Then
                Call AddItems(strItemIDs): Vsf.EditMode(mCol.检验结果 + intColCount * mintColCount) = 1: Vsf.Col = mCol.检验结果 + intColCount * mintColCount
                ShowValue 1, Val(Vsf.TextMatrix(Row, mCol.结果类型 + intColCount * mintColCount)), Vsf.Cell(flexcpData, Row, intColCount * mintColCount, Row, intColCount * mintColCount)
                mblnChangeEdit = True
            End If
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsf_GotFocus()
    Dim intColCount As Integer
    intColCount = GetColCount(Me.Vsf.Col)
    If mblnEdit Then
        ShowValue 1, Val(Vsf.TextMatrix(Vsf.Row, mCol.结果类型 + intColCount * mintColCount)), Vsf.Cell(flexcpData, Vsf.Row, intColCount * mintColCount, Vsf.Row, intColCount * mintColCount)
    End If
End Sub

Private Sub Vsf_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnCancel  As Boolean
    Dim intColCount As Integer
    Dim intCol As Integer, intRow As Integer, intTemp As Integer
    Dim intItem As Integer
    Dim lngLoop As Long
    Dim intStart As Integer
    Dim lngColor As Long, lngForeColor As Long
    
    On Error GoTo ErrHand
    
    If KeyCode = vbKeyDelete Then
        If Shift = 0 And Vsf.Body.Editable <> flexEDNone Then
            '删除整行及内容
            KeyCode = 0
            blnCancel = False
            
            Call vsf_BeforeDeleteRow(Vsf.Row, Vsf.Col, blnCancel)
            
            If blnCancel Then Exit Sub
            
            intColCount = GetColCount(Vsf.Cols)
            
            If intColCount = 0 Then
                '单行处理
                If Vsf.Rows > 1 Then
                    If Vsf.Rows = 2 And Vsf.Row = 1 Then
                        For lngLoop = 0 To Vsf.Cols - 1
                            Vsf.TextMatrix(1, lngLoop) = ""
                        Next
                        Vsf.RowData(1) = ""
                    Else
                        Vsf.RemoveItem Vsf.Row
                    End If
                    Call vsf_AfterDeleteRow(Vsf.Row, Vsf.Col)
                End If
            Else
                '多行处理
                If intColCount = 0 Then intColCount = 1
                intTemp = GetColCount(Vsf.Col)
                intStart = Vsf.Row
                For intCol = intTemp To intColCount - 1
                    For intRow = intStart To Vsf.Rows - 1
                        With Vsf
                            lngColor = &H80000005
                            lngForeColor = COLOR.默认前景色
                            If intRow < Vsf.Rows - 1 Then
                                .Cell(flexcpData, intRow, intCol * mintColCount, intRow, intCol * mintColCount) = _
                                .Cell(flexcpData, intRow + 1, intCol * mintColCount, intRow + 1, intCol * mintColCount)
                                Vsf.Cell(flexcpBackColor, intRow, mCol.检验结果 + intCol * mintColCount, intRow, mCol.检验结果 + intCol * mintColCount) = lngColor
                                Vsf.Cell(flexcpForeColor, intRow, mCol.检验结果 + intCol * mintColCount, intRow, mCol.检验结果 + intCol * mintColCount) = lngForeColor
                                For intItem = 1 To 20
                                    .TextMatrix(intRow, intItem + intCol * mintColCount) = .TextMatrix(intRow + 1, intItem + intCol * mintColCount)
                                Next
                            Else
                                '分列显示了
                                If intCol + 1 <= intColCount - 1 Then
                                    .Cell(flexcpData, intRow, intCol * mintColCount, intRow, intCol * mintColCount) = _
                                    .Cell(flexcpData, 1, (intCol + 1) * mintColCount, 1, (intCol + 1) * mintColCount)
                                    Vsf.Cell(flexcpBackColor, intRow, mCol.检验结果 + intCol * mintColCount, intRow, mCol.检验结果 + intCol * mintColCount) = lngColor
                                    Vsf.Cell(flexcpForeColor, intRow, mCol.检验结果 + intCol * mintColCount, intRow, mCol.检验结果 + intCol * mintColCount) = lngForeColor
                                    For intItem = 1 To 20
                                        .TextMatrix(intRow, intItem + intCol * mintColCount) = .TextMatrix(1, intItem + (intCol + 1) * mintColCount)
                                    Next
                                    intStart = 1
                                End If
                            End If
                        End With
                    Next
                Next
                lngLoop = 0
                For intCol = 0 To intColCount - 1
                    For intRow = 1 To Vsf.Rows - 1
                        If Val(Vsf.Cell(flexcpData, intRow, intCol * mintColCount, intRow, intCol * mintColCount)) <> 0 Then
                            lngLoop = lngLoop + 1
                            Vsf.TextMatrix(intRow, intCol * mintColCount) = lngLoop
                            Vsf.Cell(flexcpBackColor, intRow, intCol * mintColCount, intRow, intCol * mintColCount) = &HFDD6C6
                        Else
                            Vsf.TextMatrix(intRow, intCol * mintColCount) = ""
                            Vsf.Cell(flexcpBackColor, intRow, intCol * mintColCount, intRow, intCol * mintColCount) = &H80000005
                        End If
                    Next
                Next
            End If
        End If
    End If
    
    Exit Sub
    
ErrHand:
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    
    Dim strSvrText As String, strItemIDs As String
    Dim intColCount As Integer
    
    On Error GoTo errH
    intColCount = GetColCount(Col)

    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        If InStr(Vsf.EditText, "'") > 0 Then
            Cancel = True
            Exit Sub
        End If

        Select Case Col
            Case mCol.检验项目 + intColCount * mintColCount
                Vsf.Cell(flexcpData, Row, Col, Row, Col) = Vsf.EditText
                Call vsf_AfterEdit(Row, Col)
            Case mCol.检验结果 + intColCount * mintColCount
                If Row = Vsf.Rows - 1 And Col + mintColCount <= Vsf.Cols Then
                    Vsf.Row = 1
                    Vsf.Col = mCol.检验结果 + (intColCount + 1) * mintColCount
                Else
                    If Row = Vsf.Rows - 1 Then
                        Vsf.Rows = Vsf.Rows + 1
                        Vsf.Row = Vsf.Row + 1
                        Vsf.Col = mCol.检验项目 + intColCount * mintColCount
                        Vsf.ShowCell Vsf.Row, Vsf.Col
                    Else
                        Vsf.Row = Vsf.Row + 1
                        Vsf.Col = mCol.检验结果 + intColCount * mintColCount
                    End If
                    
                End If
            Case mCol.od + intColCount * mintColCount
                Vsf.SetFocus
        End Select
        
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsf_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    Dim intColCount As Integer
    
    On Error GoTo errH
    intColCount = GetColCount(Col)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Row < Vsf.Rows - 1 Then
            Vsf.Row = Vsf.Row + 1
        ElseIf Col + mintColCount <= Vsf.Cols Then
            Vsf.Row = 1
            Vsf.Col = Vsf.Col + mintColCount
        Else
            Vsf.Rows = Vsf.Rows + 1
            Vsf.Row = Vsf.Row + 1
            Vsf.Col = mCol.检验项目 + intColCount * mintColCount
            Vsf.ShowCell Vsf.Row, Vsf.Col
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim intColCount As Integer
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Exit Sub
    If Chr(KeyAscii) = "'" Then KeyAscii = 0: Exit Sub
    intColCount = GetColCount(Col)
    If Col = mCol.检验结果 + intColCount * mintColCount Then
        Select Case Val(Vsf.TextMatrix(Vsf.Row, mCol.结果类型 + intColCount * mintColCount))
            Case 1
                KeyAscii = FilterKeyAscii(KeyAscii, 2)
        End Select
        mblnChangeEdit = True
    End If
End Sub
'通过简码获取检验备注
Private Function GetComment(ByVal strCode As String, ByVal strTYPE As String)
    Dim rsTmp As ADODB.Recordset
    Dim objPoint As POINTAPI, mstrSQL As String
    Dim sglX As Single, sglY As Single
    
    mstrSQL = "SELECT Rownum As ID,A.编码,A.简码,A.名称,A.说明 As 内容 FROM 检验备注文字 A " & _
        "WHERE (Instr(A.编码,[1])>0 Or Instr(A.简码,[1])>0) And (A.分类 Is Null Or A.分类=[2])"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, UCase(strCode), mstrType)
    If rsTmp.EOF Then
        GetComment = strCode
    Else
        If rsTmp.RecordCount = 1 Then
            GetComment = Nvl(rsTmp("内容"))
        Else
            Call ClientToScreen(txtComment.hWnd, objPoint)
    
            sglX = objPoint.X * 15 - 30
            sglY = objPoint.Y * 15 - 2000
            If frmSelectList.ShowSelect(Me, rsTmp, "编码,800,0,0;简码,1500,0,0;名称,2500,0,0;内容,5500,0,0", sglX, sglY, Me.txtComment.Width, 2000, Me.Name & "\检验备注选择", "请选择检验备注") Then
                GetComment = Nvl(rsTmp("内容"))
            Else
                GetComment = strCode
            End If
        End If
    End If
End Function

Private Sub AddItems(ByVal strItemIDs As String)
'添加检验项目(不含微生物项目)
'strItemIDs：诊疗项目ID串，以，分隔
    Dim strsql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strsql = "SELECT " & mlngKey & " as 标本ID, B.ID,B.中文名||DECODE(C.缩写,NULL,'','('||C.缩写||')') AS 检验项目,'' As 上次结果,''As CV," & _
        "'' As 原始结果,Decode(C.结果类型,3,Nvl(C.默认值,'-'),2,C.默认值,'') As 本次结果,C.计算公式,C.结果类型," & _
        "'' AS 标志,'' as OD,'' as CUTOFF,'' as COV, '' as 酶标板ID,c.变异报警率 as 变异报警,c.变异警示率 as 变异警示, " & _
        "Trim(REPLACE(REPLACE(' '||zlGetReference(B.ID,A.标本部位,0,NULL,[1]),' .','0.'),'～.','～0.')) AS 参考," & _
        "[1] As 仪器ID,C.结果范围,0 As 固定项目,Nvl(E.小数位数,2) As 小数,C.警戒上限,C.警戒下限,C.单位,'' as 上次时间,a.id as 诊疗项目ID,'' as 排列序号 " & _
        ",Zl_To_Number(Zl_Get_Reference(1, b.id, A.标本部位, 0, Null,[1])) as 参考ID " & vbNewLine & _
        "FROM 诊疗项目目录 A,检验报告项目 D,诊治所见项目 B,检验项目 C,检验仪器项目 E " & _
        "WHERE A.ID = D.诊疗项目ID And D.报告项目ID=B.ID " & _
                    "AND B.ID = C.诊治项目ID And D.报告项目ID=E.项目ID(+) And E.仪器ID(+)=[1] " & _
                    "AND D.细菌ID IS NULL AND C.项目类别<>2 " & _
                    "AND A.ID In (Select * From Table(Cast(f_Num2list([2]) As zlTools.t_Numlist)))  "
                    
    strsql = "Select a.标本id,a.id,a.检验项目,a.上次结果,a.cv,a.原始结果,a.本次结果,a.计算公式,a.结果类型,a.标志,a.od,a.cutoff,a.cov,a.酶标板id" & _
           ",a.变异报警,a.变异警示,a.参考,a.仪器id,a.结果范围,a.固定项目,a.小数,f.警示上限 as 警戒上限,f.警示下限 as 警戒下限,a.单位,a.上次时间,a.诊疗项目id,a.排列序号,null as 仪器提示,null as 仪器审核标识" & _
           " From (" & strsql & ") a,检验项目参考 F Where a.参考id=f.id(+) Order By A.ID,a.排列序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, mDeviceID, strItemIDs)
    
    If Not rsTmp.EOF Then
        Vsf.TextMatrix(0, 0) = "#"
'        Call FillGrid_UQ(vsf, rsTmp, Array("", "", "", ""), False)
        Call ReadVsf(rsTmp, Array("", "", "", ""), False)
        Vsf.TextMatrix(0, 0) = ""
        Vsf.Cell(flexcpBackColor, 1, 0, Vsf.Rows - 1, 0) = &HFDD6C6
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function LoadModel(ByVal strCode As String) As Boolean
'调入报告模板(不含微生物项目)
'strCode：模板编码或简码
    Dim strsql As String, rsTmp As ADODB.Recordset
    Dim lngCurrRow As Long
    Dim intColCount As Integer
    Dim intCol As Integer
    
    On Error GoTo errH
    
    LoadModel = False
    strsql = "SELECT B.ID,B.中文名||DECODE(C.缩写,NULL,'','('||C.缩写||')') AS 检验项目,'' As 上次结果,''As CV," & _
        "'' As 原始结果,A.检验结果 As 本次结果,C.计算公式,C.结果类型," & _
        "'' AS 标志," & _
        "Trim(REPLACE(REPLACE(' '||zlGetReference(B.ID,'',0,NULL,''),' .','0.'),'～.','～0.')) AS 参考," & _
        "[2] As 仪器ID,C.结果范围,0 As 固定项目,2 As 小数,C.警戒上限,C.警戒下限,C.单位 " & _
        ",zl_Get_Reference(1,B.ID,'',0,NULL,'') as 参考id " & _
        "FROM 检验模板内容 A,诊治所见项目 B,检验项目 C,检验模板目录 D " & _
        "WHERE A.项目ID=B.ID AND B.ID = C.诊治项目ID And D.ID=A.模板ID " & _
                    "AND A.细菌ID IS NULL AND (D.编码=[1] Or D.简码=[1])"
    strsql = "Select a.ID,a.检验项目,a.上次结果,a.原始结果,a.本次结果,a.计算公式,a.结果类型,a.标志,a.参考,a.仪器id" & _
            ",a.结果范围,a.固定项目,a.小数,f.警示上限 as 警戒上限,f.警示下限 as 警戒下限,a.单位" & _
            " From (" & strsql & ") a,检验项目参考 F where a.参考id=F.ID(+)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, strCode, mDeviceID)
    
    If Not rsTmp.EOF Then
        Do While Not rsTmp.EOF
            lngCurrRow = FindRepeatLine(Vsf, CStr(zlCommFun.Nvl(rsTmp("ID"))))
            If lngCurrRow > 0 Then
                intColCount = GetColCount(Vsf.Cols)
                If intColCount = 0 Then intColCount = 1
                For intCol = 0 To intColCount - 1
                    If Val(Vsf.Cell(flexcpData, lngCurrRow, intCol * mintColCount, lngCurrRow, intCol * mintColCount)) = Nvl(rsTmp("ID")) Then
                        Exit For
                    End If
                Next
                Vsf.TextMatrix(lngCurrRow, mCol.检验结果 + intCol * mintColCount) = Nvl(rsTmp("本次结果"))
                Vsf.TextMatrix(lngCurrRow, mCol.结果参考 + intCol * mintColCount) = Nvl(rsTmp("参考"))
                '产生缺省的结果标志
                Vsf.TextMatrix(lngCurrRow, mCol.结果标志 + intCol * mintColCount) = CalcDefaultFlag(Trim(Vsf.TextMatrix(lngCurrRow, mCol.检验结果 + intCol * mintColCount)), _
                    Trim(Vsf.TextMatrix(lngCurrRow, mCol.结果参考 + intCol * mintColCount)), Val(Vsf.TextMatrix(lngCurrRow, mCol.结果类型 + intCol * mintColCount)), _
                    Vsf.TextMatrix(lngCurrRow, mCol.警戒下限 + intCol * mintColCount), Vsf.TextMatrix(lngCurrRow, mCol.警戒上限 + intCol * mintColCount))
                
                '根据结果应用颜色标志
                Call ApplyResultColor(Vsf, lngCurrRow, mCol.检验结果 + intCol * mintColCount, _
                    Decode(Vsf.TextMatrix(lngCurrRow, mCol.结果标志 + intCol * mintColCount), "↑", 3, "↓", 2, "异常", 4, "↓↓", 5, "↑↑", 6, 1))
            End If
            
            rsTmp.MoveNext
        Loop
        
        LoadModel = True
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadLastValue()
    '功能：装入上次结果数据
    Dim lngDates As Long, lngTimes As Long
    Dim strRows As String, aryRows() As String
    Dim strCols As String, aryCols() As String
    Dim dblCurCV As Double     '计算的CV
    Dim lngDays As Long, mstrEndTime As String, rsTemp As ADODB.Recordset
    Dim lngRow As Long, lngCurrKey As Long
    Dim intFindMode As Integer          '病人查找模式
    Dim intColCount As Integer, intCol As Integer
    Dim dblCalc As Double
    Dim strTag As String                '显示标识
    Dim dbl变异报警 As Double, dbl变异警示 As Double
    Dim intSampleType As Integer
    
    If mlngKey = 0 Then Exit Sub
    
    Err = 0: On Error GoTo ErrHand
    
    intFindMode = zlDatabase.GetPara("历史病人识别", 100, 1208, 0)
    intSampleType = zlDatabase.GetPara("上次结果不参照标本类型", 100, 1208, 0)
    
    '获得当前检验的时间、项目要求的跟踪天数（取项目中最大的）
    gstrSql = "Select Nvl(L.检验时间, Sysdate) As 检验时间, Nvl(Max(跟踪天数), 0) As 天数" & vbNewLine & _
            "From 检验项目选项 O, 检验报告项目 X, 检验普通结果 R, 检验标本记录 L" & vbNewLine & _
            "Where O.诊疗项目id(+) = X.诊疗项目id And X.报告项目id = R.检验项目id And R.检验标本id = L.ID And L.ID = [1]" & vbNewLine & _
            "Group By Nvl(L.检验时间, Sysdate)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey)
    If rsTemp.RecordCount > 0 Then
        lngDays = rsTemp!天数
        mstrEndTime = Format(rsTemp!检验时间, "yyyy-MM-dd hh:mm:ss")
    Else
        lngDays = 30
        mstrEndTime = Format(Now(), "yyyy-MM-dd hh:mm:ss")
    End If
    lngDays = IIf(lngDays = 0, 30, lngDays)
    lngDates = lngDays
    
    '查询历次数据装入：
    gstrSql = "Select L.检验项目id As ID, V.缩写 As 英文名, L.次数, L.检验时间, L.检验结果, V.变异报警率" & vbNewLine & _
            "From (Select L.检验项目id, L.次数, L.检验时间, L.检验结果" & vbNewLine & _
            "       From (Select L.病人id,L.姓名,L.性别,L.年龄, L.ID As 次数, L.检验时间, R.检验项目id, R.检验结果,L.标本类型 " & vbNewLine & _
            "              From 检验标本记录 L, 检验普通结果 R" & vbNewLine & _
            "              Where L.ID = R.检验标本id AND L.报告结果=R.记录类型 AND " & vbNewLine & _
            "                    L.检验时间  Between To_Date([2], 'yyyy-mm-dd hh24:mi:ss') - [3] And" & vbNewLine & _
            "                    To_Date([2], 'yyyy-mm-dd hh24:mi:ss') And L.ID<>[1] And " & vbNewLine & _
            "                    " & IIf(intFindMode = 0, " L.病人id = [4] ", " L.病人ID in (select 病人id from 病人信息 where 姓名 = [5] )") & ") L," & vbNewLine & _
            "            (Select L.病人id,L.姓名,L.性别,L.年龄,R.检验项目id,L.标本类型 " & vbNewLine & _
            "              FROM 检验标本记录 L,检验普通结果 R" & vbNewLine & _
            "              WHERE L.ID = [1] AND L.ID = R.检验标本id AND L.报告结果=R.记录类型) C" & vbNewLine & _
            "       Where " & IIf(intFindMode = 0, " L.病人id = C.病人id ", " L.姓名 = C.姓名 ") & vbNewLine & _
            "             AND L.检验项目id+0 =C.检验项目id" & IIf(intSampleType = 0, " And nvl(L.标本类型,'') = nvl(C.标本类型,'')", "") & ") L, 检验项目 V " & vbNewLine & _
            "Where L.检验项目id = V.诊治项目id" & vbNewLine & _
            "Order By L.次数 Desc"
'    gstrSql = "Select /*+ rule */" & vbNewLine & _
                " L.检验项目id As ID, V.缩写 As 英文名, L.次数, L.检验时间, L.检验结果, V.变异报警率" & vbNewLine & _
                " From (Select L.病人id, L.姓名, L.性别, L.年龄, L.ID As 次数, L.检验时间, R.检验项目id, R.检验结果" & vbNewLine & _
                "       From 检验标本记录 L, 检验普通结果 R" & vbNewLine & _
                "       Where L.ID = R.检验标本id And L.报告结果 = R.记录类型 And" & vbNewLine & _
                "             L.检验时间 Between To_Date([2], 'yyyy-mm-dd hh24:mi:ss') - [3] And " & vbNewLine & _
                "             To_Date([2], 'yyyy-mm-dd hh24:mi:ss') And L.ID<>[1] " & vbNewLine & _
                "             And " & IIf(intFindMode = 0, " L.病人id = [4] ", " L.姓名=[5] ") & ") L, " & vbNewLine & _
                "       检验项目 V" & vbNewLine & _
                " Where L.检验项目id = V.诊治项目id" & vbNewLine & _
                " Order By L.次数 Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey, mstrEndTime, lngDates, mLngPatientID, mstrPatientName)
    
    
    Err = 0: On Error GoTo 0
    With Me.Vsf
        lngCurrKey = 0: lngTimes = 0
        Do While Not rsTemp.EOF
            If lngCurrKey <> rsTemp("次数") Then
                '只检索一次结果
                If lngTimes = 1 Then Exit Do
                lngTimes = lngTimes + 1
                lngCurrKey = rsTemp("次数")
            End If
            lngRow = FindRepeatLine(Vsf, rsTemp("ID"))
            If lngRow > 0 Then
                intColCount = GetColCount(Vsf.Cols)
                If intColCount = 0 Then intColCount = 1
                For intCol = 0 To intColCount - 1
                    If Val(.Cell(flexcpData, lngRow, intCol * mintColCount, lngRow, intCol * mintColCount)) = Nvl(rsTemp("ID")) Then
                        .TextMatrix(lngRow, mCol.上次结果 + intCol * mintColCount) = Nvl(rsTemp("检验结果"))
                        If Nvl(rsTemp("检验结果")) <> "" Then
                            .TextMatrix(lngRow, mCol.上次时间 + intCol * mintColCount) = Format(Nvl(rsTemp("检验时间")), "YYYY-MM-DD")
                            dblCalc = 0
                            If Val(.TextMatrix(lngRow, mCol.检验结果 + intCol * mintColCount)) <> 0 Then
                                dblCalc = (Val(Nvl(rsTemp("检验结果"))) - Val(.TextMatrix(lngRow, mCol.检验结果 + intCol * mintColCount))) / _
                                Val(.TextMatrix(lngRow, mCol.检验结果 + intCol * mintColCount)) * 100
                                dblCalc = Format(dblCalc, "00#.##")
                            End If
                            strTag = ""
                            
                            dbl变异报警 = Val(.TextMatrix(lngRow, mCol.变异报警 + intCol * mintColCount))
                            dbl变异警示 = Val(.TextMatrix(lngRow, mCol.变异警示 + intCol * mintColCount))
                            
                            If dblCalc > 0 Then
                                If dblCalc >= dbl变异警示 And dbl变异警示 <> 0 Then
                                    strTag = "↑↑"
                                ElseIf dblCalc >= 10 And dbl变异报警 <> 0 Then
                                    strTag = "↑"
                                End If
                            Else
                                If Abs(dblCalc) >= dbl变异警示 And dbl变异警示 <> 0 Then
                                    strTag = "↓↓"
                                ElseIf Abs(dblCalc) >= dbl变异报警 And dbl变异报警 <> 0 Then
                                    strTag = "↓"
                                End If
                            End If
                            .TextMatrix(lngRow, mCol.检验项目 + intCol * mintColCount) = Replace(.TextMatrix(lngRow, mCol.检验项目 + intCol * mintColCount), "↑", "")
                            .TextMatrix(lngRow, mCol.检验项目 + intCol * mintColCount) = Replace(.TextMatrix(lngRow, mCol.检验项目 + intCol * mintColCount), "↓", "")
                            .TextMatrix(lngRow, mCol.检验项目 + intCol * mintColCount) = .TextMatrix(lngRow, mCol.检验项目 + intCol * mintColCount) & strTag
                            Select Case strTag
'                                Case "↑"
'                                    .Cell(flexcpForeColor, lngRow, mCol.检验项目 + intCol * mintColCount, lngRow, mCol.检验项目 + intCol * mintColCount) = COLOR.超标背景色
'                                Case "↓"
'                                    .Cell(flexcpForeColor, lngRow, mCol.检验项目 + intCol * mintColCount, lngRow, mCol.检验项目 + intCol * mintColCount) = COLOR.低标背景色 + 300
                                Case "↓↓", "↑↑"
                                    .Cell(flexcpForeColor, lngRow, mCol.检验项目 + intCol * mintColCount, lngRow, mCol.检验项目 + intCol * mintColCount) = COLOR.橙色
                            End Select
                        End If
                        
                    End If
                Next
            End If
        
            rsTemp.MoveNext
        Loop
        
'        '变异率计算填写和报警色处理
'        For lngRow = .FixedRows To .Rows - 1
'            .TextMatrix(lngRow, mCol.单位 + 1) = Replace(Replace(" " & Trim(.TextMatrix(lngRow, mCol.单位 + 1)), " .", "0."), " ", "")
'            For lngCol = mCol.单位 + 4 To .Cols - 1 Step 2
'                .TextMatrix(lngRow, lngCol - 1) = Replace(Replace(" " & Trim(.TextMatrix(lngRow, lngCol - 1)), " .", "0."), " ", "")
'                If Val(.TextMatrix(lngRow, lngCol - 1)) = 0 Or Val(.TextMatrix(lngRow, mCol.单位 + 1)) = 0 Then
'                    dblCurCV = 0
'                Else
'                    dblCurCV = (Val(.TextMatrix(lngRow, lngCol - 1)) - Val(.TextMatrix(lngRow, mCol.单位 + 1))) / Val(.TextMatrix(lngRow, mCol.单位 + 1)) * 100
'                End If
'                .TextMatrix(lngRow, lngCol) = Format(dblCurCV, "0.00;-0.00; ; ")
'                If Val(.TextMatrix(lngRow, mCol.报警率)) <> 0 And Abs(dblCurCV) > Val(.TextMatrix(lngRow, mCol.报警率)) Then
'                    .Cell(flexcpBackColor, lngRow, lngCol) = RGB(248, 194, 169)
'                End If
'            Next
'        Next
'        .Redraw = flexRDDirect
    End With
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
    
    Call FormatVsfCell(Vsf, mCol.上次结果, "0.0######", 0, IIf(mDeviceID > 0, mCol.小数, -1))
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Vsf_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim rsTmp As New ADODB.Recordset
    
    If Button <> vbRightButton Then Exit Sub
    
    On Error GoTo errH
    
    gstrSql = "select distinct nvl(记录类型,0) as 记录类型 from 检验普通结果 where 检验标本id = [1] order by nvl(记录类型,0) "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, mlngKey)
    Set cbrPopupBar = Me.cbrthis.Add("弹出菜单", xtpBarPopup)
    Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, 100, "临床意义")
    If rsTmp.RecordCount > 1 Then
        Do Until rsTmp.EOF
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, rsTmp(0) + 1, "选择第" & rsTmp(0) + 1 & "次")
            rsTmp.MoveNext
        Loop
'        cbrPopupBar.ShowPopup
    End If
    cbrPopupBar.ShowPopup
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not mblnEdit Then
        RaiseEvent StartEdit(Cancel)
        If mblnPatientFind = True Then Cancel = True
        If Cancel = False Then mblnEvent = True
    End If
    
End Sub


Public Sub Resize()
    '供主窗体调用
    Call Form_Resize
End Sub

Private Function ReadVsf(ByVal rsData As ADODB.Recordset, Optional ByVal MaskArray As Variant, Optional ByVal blnClear As Boolean = True) As Boolean
    Dim lngLoop As Long
    Dim strMask As String
    Dim lngRow As Long, lngCurrRow As Long
    Dim strOldValue As String, strNewValue As String
    Dim intColCount  As Integer
    Dim intCol As Integer, intRow As Integer
    Dim lngHeight As Long
    Dim blnShowType As Boolean
    Dim str复查 As String
    blnShowType = zlDatabase.GetPara("自适应显示结果", 100, 1208, False)
    If fraComment.Tag <> "" Then blnShowType = True
    
    If blnClear Then
        Vsf.Rows = 2
        Vsf.RowData(1) = 0
        For lngLoop = 0 To Vsf.Cols - 1
            Vsf.TextMatrix(1, lngLoop) = ""
            Vsf.Cell(flexcpData, 1, lngLoop, 1, lngLoop) = ""
        Next
        lngRow = 0
        Vsf.Cols = mintColCount
    Else
        '预先有一空行
        With Vsf
            intColCount = GetColCount(.Cols)
            If intColCount = 0 Then intColCount = 1
            For intCol = 0 To intColCount - 1
                For intRow = 1 To .Rows - 1
                    If Val(.Cell(flexcpData, intRow, intCol * mintColCount, intRow, intCol * mintColCount)) = 0 Then
                        lngRow = intRow - 1
                        intColCount = intCol
                        Exit For
                    End If
                Next
            Next
        End With
    End If
    
    
    With Vsf.Body
        If .ClientHeight < .CellHeight * 15 Then
            lngHeight = .CellHeight * 15
        Else
            lngHeight = .ClientHeight
        End If
    End With
    Do While Not rsData.EOF
        lngCurrRow = FindRepeatLine(Vsf, CStr(zlCommFun.Nvl(rsData("ID"))))
'        lngCurrRow = -1
        If lngCurrRow = -1 Then
            With Vsf.Body

                If (.CellHeight + 15) * (lngRow + 2) > lngHeight And blnShowType = True Then
                    intColCount = intColCount + 1
                    lngRow = 1
                    With Vsf
                        .NewColumn "#", 300, 7
                        .NewColumn "检验项目", 2100, 1
                        .NewColumn "原始结果", 0, 1
                        .NewColumn "本次结果", 1200, 1, , 1
                        .NewColumn "单位", 1000, 1
                        .NewColumn "CV", 0, 1
                        .NewColumn "标志", 450, 1
                        .NewColumn "上次结果", 0, 1
                        .NewColumn "上次时间", 0, 1
                        .NewColumn "参考", 1300, 1
                        .NewColumn "结果类型", 0, 1
                        .NewColumn "仪器id", 0, 1
                        .NewColumn "计算公式", 0, 1
                        .NewColumn "结果范围", 0, 1
                        .NewColumn "固定项目", 0, 1
                        .NewColumn "小数", 0, 1
                        .NewColumn "警戒上限", 0, 1
                        .NewColumn "警戒下限", 0, 1
                        .NewColumn "诊疗项目ID", 0, 1
                        .NewColumn "排列序号", 0, 1
                        .NewColumn "标本ID", 0, 1
                        .NewColumn "OD", 700, 1, , 1
                        .NewColumn "CUTOFF", 700, 1
                        .NewColumn "COV", 700, 1
                        .NewColumn "酶标板ID", 0, 1
                        .NewColumn "变异报警", 0, 1
                        .NewColumn "变异警示", 0, 1
                        .NewColumn "仪器提示", 1000, 1
                        .NewColumn "仪器审核标识", 1200, 1
                    End With
                Else
                    lngRow = lngRow + 1
                End If
            End With
            
            lngCurrRow = lngRow
        
            If Vsf.Rows < lngRow + 1 Then Vsf.Rows = lngRow + 1
            
            On Error Resume Next
'            Vsf.RowData(lngCurrRow) = CStr(zlCommFun.Nvl(rsData("ID")))
            Vsf.Cell(flexcpData, lngCurrRow, intColCount * mintColCount, lngCurrRow, intColCount * mintColCount) = CStr(Nvl(rsData("ID")))
            
            On Error GoTo ErrHand
            
            str复查 = Get复查标记(Val("" & rsData("标本ID")), Val("" & rsData("ID")), "" & rsData("本次结果"))
            
            For lngLoop = 0 To mintColCount - 1
                intCol = intColCount * mintColCount + lngLoop
                
                If Trim(Vsf.TextMatrix(0, intCol)) <> "" Then
                    If Vsf.TextMatrix(0, intCol) = "#" Then
                        Vsf.TextMatrix(lngCurrRow, intCol) = IIf(intColCount > 0, intColCount * (Vsf.Body.Rows - 1) + lngCurrRow, lngCurrRow)
                        Vsf.Cell(flexcpBackColor, lngCurrRow, intCol, lngCurrRow, intCol) = &HFDD6C6
                    Else
                        On Error Resume Next
                        strMask = ""
                        strMask = MaskArray(intCol)
                                                
                        On Error GoTo ErrHand
                        
                        If strMask <> "" Then
                            strNewValue = Format(zlCommFun.Nvl(rsData(Vsf.TextMatrix(0, intCol))), strMask)
                        Else
                            strNewValue = zlCommFun.Nvl(rsData(Vsf.TextMatrix(0, intCol)))
                        End If
                        If str复查 <> "" Then
                            If rsData(Vsf.TextMatrix(0, intCol)).Name = "检验项目" Then
                                strNewValue = strNewValue & " " & str复查
                                Vsf.Cell(flexcpForeColor, lngCurrRow, intCol, lngCurrRow, intCol) = COLOR.橙色
                            End If
                        End If
                        Vsf.TextMatrix(lngCurrRow, intCol) = strNewValue
                    End If
                End If
                
            Next
        End If
        
        rsData.MoveNext
    Loop
'    Call chkOriginal_Click: Call chkLast_Click: Call chkSign_Click
'    Call chkUnit_Click: Call chkReferrence_Click: Call chkMB_Click
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        Vsf.Body.ColWidth(mCol.检验项目 + intCol * mintColCount) = IIf(chkChina.Value, 2100, 1000)
        Vsf.Body.ColWidth(mCol.原始结果 + intCol * mintColCount) = IIf(chkOriginal.Value, 900, 0)
        Vsf.Body.ColWidth(mCol.上次结果 + intCol * mintColCount) = IIf(chkLast.Value, 900, 0)
        Vsf.Body.ColWidth(mCol.上次时间 + intCol * mintColCount) = IIf(chkLast.Value, 1000, 0)
        Vsf.Body.ColWidth(mCol.结果标志 + intCol * mintColCount) = IIf(chkSign.Value, 450, 0)
        Vsf.Body.ColWidth(mCol.单位 + intCol * mintColCount) = IIf(chkUnit.Value, 1000, 0)
        Vsf.Body.ColWidth(mCol.结果参考 + intCol * mintColCount) = IIf(chkReferrence.Value, 1300, 0)
        Vsf.Body.ColWidth(mCol.od + intCol * mintColCount) = IIf(chkMB.Value, 700, 0)
        Vsf.Body.ColWidth(mCol.CUTOFF + intCol * mintColCount) = IIf(chkMB.Value, 700, 0)
        Vsf.Body.ColWidth(mCol.COV + intCol * mintColCount) = IIf(chkMB.Value, 700, 0)
        Vsf.Body.ColWidth(mCol.仪器提示 + intCol * mintColCount) = IIf(chkYiQiTiShi.Value, 1000, 0)
        Vsf.Body.ColWidth(mCol.仪器审核标识 + intCol * mintColCount) = IIf(chkYiQiBiaoShi.Value, 1200, 0)
    Next
    
    Exit Function
    
ErrHand:

    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function GetColCount(Col As Long) As Integer
    '功能               返回当前是第几个分列项目
    '参数               当前列数
    Dim dblTmp As Double
    If Col <= mintColCount Then
        GetColCount = 0
    Else
        dblTmp = Col / mintColCount
        If InStr(dblTmp, ".") > 0 Then
            GetColCount = Mid(dblTmp, 1, InStr(dblTmp, ".") - 1)
        Else
            GetColCount = dblTmp
        End If
    End If
End Function
Private Function FindRepeatLine(ByRef objMsf As Object, ByVal strSeekID As String) As Long
    '-------------------------------------------------------------------------------------------------------------
    '功能:查找RowData等于strSeekID的行
    '参数:
    '返回:行号或-1
    '-------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim intColCount As Integer, intCol As Integer
    FindRepeatLine = -1
    intColCount = GetColCount(objMsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        For i = 1 To objMsf.Rows - 1
            If Val(Me.Vsf.Cell(flexcpData, i, intCol * mintColCount, i, intCol * mintColCount)) = strSeekID Then
                FindRepeatLine = i
                Exit For
            End If
'            If objMsf.RowData(i) = strSeekID Then Exit For
        Next
    Next
    If i <= objMsf.Rows - 1 Then FindRepeatLine = i
End Function

Private Function Get复查标记(ByVal lng标本ID As Long, ByVal lng项目ID As Long, ByVal str检验结果 As String) As String
    Dim str复查 As String, strWhere As String
    Dim rsTmp As ADODB.Recordset
    Dim bln符合条件 As Boolean
    Dim lng参考ID As Long
    str复查 = ""
    
    If Not IsNumeric(str检验结果) Then
        Exit Function
    End If
    
    gstrSql = "Select Zl_Get_Reference(1, " & lng项目ID & ", a.标本类型, Decode(a.性别, '男', 1, '女', 2, 0), a.出生日期, a.仪器id, a.年龄,a.申请科室iD) As 参考id" & vbNewLine & _
                "From 检验标本记录 A" & vbNewLine & _
                "Where a.Id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng标本ID)
    If rsTmp.EOF = True Then Exit Function
    lng参考ID = Val("" & rsTmp!参考id)
    If lng参考ID <> 0 Then
        gstrSql = "Select 复查上限, 复查下限 From 检验项目参考 A Where a.Id = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng参考ID)
        Do Until rsTmp.EOF
            If "" & rsTmp!复查下限 <> "" And "" & rsTmp!复查上限 <> "" Then
                If Val(str检验结果) < Val("" & rsTmp!复查下限) Or Val(str检验结果) > Val("" & rsTmp!复查上限) Then
                    Get复查标记 = "需复查"
                    Exit Function
                End If
            ElseIf "" & rsTmp!复查下限 = "" And "" & rsTmp!复查上限 <> "" Then
                If Val(str检验结果) > Val("" & rsTmp!复查上限) Then
                    Get复查标记 = "需复查"
                    Exit Function
                End If
            ElseIf "" & rsTmp!复查下限 <> "" And "" & rsTmp!复查上限 = "" Then
                If Val(str检验结果) < Val("" & rsTmp!复查下限) Then
                    Get复查标记 = "需复查"
                    Exit Function
                End If
            End If
            rsTmp.MoveNext
        Loop
    End If
End Function
