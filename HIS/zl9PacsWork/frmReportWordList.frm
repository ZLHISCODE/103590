VERSION 5.00
Begin VB.Form frmReportWordList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "保存词句示范"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8400
   Icon            =   "frmReportWordList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "插入标记"
      Height          =   735
      Left            =   240
      TabIndex        =   11
      Top             =   1200
      Width           =   6375
      Begin VB.CommandButton cmdWordTag 
         Caption         =   "检查所见"
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   15
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdWordTag 
         Caption         =   "诊断意见"
         Height          =   375
         Index           =   2
         Left            =   3000
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdWordTag 
         Caption         =   "建议"
         Height          =   375
         Index           =   3
         Left            =   4920
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdCompound 
         Caption         =   "组合"
         Height          =   375
         Left            =   5280
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.TextBox txtWord 
      Height          =   3735
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2160
      Width           =   7935
   End
   Begin VB.TextBox txt分类 
      Height          =   300
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   8
      Top             =   510
      Width           =   5265
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6960
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txt名称 
      Height          =   300
      Left            =   1320
      TabIndex        =   4
      Top             =   120
      Width           =   5265
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6960
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.OptionButton opt范围 
      Caption         =   "全院通用(&1)"
      Height          =   180
      Index           =   0
      Left            =   1320
      TabIndex        =   2
      Top             =   900
      Width           =   1305
   End
   Begin VB.OptionButton opt范围 
      Caption         =   "科内通用(&2)"
      Height          =   180
      Index           =   1
      Left            =   3360
      TabIndex        =   1
      Top             =   900
      Width           =   1305
   End
   Begin VB.OptionButton opt范围 
      Caption         =   "个人使用(&3)"
      Height          =   180
      Index           =   2
      Left            =   5400
      TabIndex        =   0
      Top             =   900
      Value           =   -1  'True
      Width           =   1305
   End
   Begin VB.Label lbl分类名称 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "分类名称(&C)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   10
      Top             =   570
      Width           =   990
   End
   Begin VB.Label lbl名称 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "词句名称(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   7
      Top             =   180
      Width           =   990
   End
   Begin VB.Label lbl范围 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "使用范围(&U)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   6
      Top             =   900
      Width           =   990
   End
End
Attribute VB_Name = "frmReportWordList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngClassID As Long     '词句分类ID
Private mstrClassName As String '词句分类名称
Private mlngWordID As Long      '词句ID
Private mlngDeptID As Long      '科室ID
Private mstr编号 As String

Public Sub zlShowMe(frmParent As Object, txtWordString As String, intWordPower As Integer, _
        lngClassID As Long, strClassName As String, lngDeptID As Long, _
        Optional ByVal lngWordID As Long)
'参数： txtWordString ---修改或者添加的词句内容
'       intWordPower --- 修改词句的权限（0-全院；1-科室；2-个人）
'       lngClassID --- 词句分类ID
'       strClassName --- 词句分类的名称
'       lngDeptID --- 科室ID
'       lngWordID --- 词句的ID，修改词句时需要提供
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strWordName As String
    
    mlngClassID = lngClassID
    mstrClassName = strClassName
    mlngWordID = lngWordID
    mlngDeptID = lngDeptID
    
    Me.txt分类.Text = mstrClassName
    Me.txt分类.Tag = mlngClassID
    Me.txt名称.MaxLength = 80       '.Fields("名称").DefinedSize
    
    If lngWordID = 0 Then
        frmReportWordList.Caption = "新增示范词句"
        mstr编号 = zlDefaultWordCode(mlngClassID)
        Me.txtWord.Text = txtWordString
        Me.txt名称.Text = strWordName
    Else
        frmReportWordList.Caption = "修改示范词句"
        
        '从词句示范中读取词句内容
        strSQL = "Select a.名称,a.通用级,a.编号, b.排列次序,b.内容文本 " & _
                 " From 病历词句示范 a,病历词句组成 b Where a.Id=[1] And a.Id=b.词句ID  order by 排列次序 "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngWordID)
        Me.txtWord.Text = ""
        If Not rsTemp.EOF Then
            mstr编号 = Nvl(rsTemp!编号)
            Me.opt范围(Nvl(rsTemp!通用级, 0)).value = True
            Me.txt名称.Text = Nvl(rsTemp!名称)
        End If
        While rsTemp.EOF = False
            Me.txtWord.Text = Me.txtWord.Text & Nvl(rsTemp!内容文本)
            rsTemp.MoveNext
        Wend
    End If
    
    Select Case intWordPower
    Case 2: Me.opt范围(0).Enabled = False: Me.opt范围(1).Enabled = False
    Case 1: Me.opt范围(0).Enabled = False
    End Select
    
    frmReportWordList.Show 1, frmParent
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim ArraySQL() As String
    Dim lngCount As Long
    Dim i As Integer
    Dim strText As String
    Dim blnAdd As Boolean   'True-新增词句示范；False-修改词句示范
    
    '检测输入内容的合法性
    If Trim(Me.txt名称.Text) = "" Then
        MsgBoxD Me, "请输入名称！", vbInformation, gstrSysName
        Me.txt名称.SetFocus
        Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt名称.Text), vbFromUnicode)) > Me.txt名称.MaxLength Then
        MsgBoxD Me, "名称超长（最多" & Me.txt名称.MaxLength & "个字符或等长的汉字）！", vbInformation, gstrSysName
        Me.txt名称.SetFocus
        Exit Sub
    End If
    If Len(Trim(txtWord.Text)) = 0 Then
        MsgBoxD Me, "请输入词句示范内容后再保存。", vbOKOnly, gstrSysName
        Exit Sub
    End If
    
    If mlngWordID = 0 Then      '新增示范词句
        mlngWordID = zlDatabase.GetNextId("病历词句示范")
        blnAdd = True
    Else                        '修改示范词句
        blnAdd = False
    End If
    
    '保存词句示范内容
    strSQL = mlngWordID & "," & Val(Me.txt分类.Tag) & ",'" & mstr编号 & "','" & Trim(Me.txt名称.Text) & "'"
    If Me.opt范围(0).value Then
        strSQL = strSQL & ",0"
    ElseIf Me.opt范围(1).value Then
        strSQL = strSQL & ",1"
    Else
        strSQL = strSQL & ",2"
    End If
    strSQL = strSQL & "," & mlngDeptID & "," & UserInfo.ID
    strSQL = "Zl_病历词句示范_Edit(" & IIf(blnAdd = True, 1, 2) & "," & strSQL & ")"
    
    '新增词句组成，加上标记：#***#
    If InStr(txtWord.Text, "<<") > 0 Then
        '第一个<<前的内容忽略不计
        strText = Mid(txtWord.Text, InStr(txtWord.Text, "<<"))
        
        If InStr(strText, "<<所见>>") > 0 Then strText = Replace(strText, "<<所见>>", "#***#<<所见>>")
        If InStr(strText, "<<诊断>>") > 0 Then strText = Replace(strText, "<<诊断>>", "#***#<<诊断>>")
        If InStr(strText, "<<建议>>") > 0 Then strText = Replace(strText, "<<建议>>", "#***#<<建议>>")
        If Mid(strText, 1, 5) = "#***#" Then strText = Mid(strText, 6)
    Else
        strText = "#***#" & txtWord.Text
    End If
    
    '获取SQL语句数组
    ReDim ArraySQL(1 To 2) As String
    ArraySQL(1) = strSQL
    
    '前期处理
    ArraySQL(2) = "Zl_病历词句组成_Beforesave(" & mlngWordID & ")"
    
    '获取保存SQL数组
    Call GetSaveSQL(ArraySQL, strText)
    
    '后期处理
    lngCount = UBound(ArraySQL) + 1
    ReDim Preserve ArraySQL(1 To lngCount) As String
    strSQL = "Zl_病历词句组成_Aftersave(" & mlngWordID & ")"
    ArraySQL(lngCount) = strSQL
    
    '执行保存操作
    err = 0: On Error GoTo errHand
    gcnOracle.BeginTrans
    For i = 1 To UBound(ArraySQL)
        strSQL = ArraySQL(i)
        Call zlDatabase.ExecuteProcedure(strSQL, "frmReportWordList")
    Next
    gcnOracle.CommitTrans
        
    Unload Me
    Exit Sub
errHand:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub GetSaveSQL(ByRef ArrySQL() As String, strText As String)
'组织保存词句组成的SQL语句
'参数： ArrySQL --- SQL 语句数组
'       strText --- 要保存的词句示范内容
    
    Dim strLine As String       '一行文本，回车之间的文本
    Dim lng序号 As Long         '按照CRLF来分段
    Dim i As Integer
    On Error GoTo err
    
    lng序号 = 1
    
    For i = 0 To UBound(Split(strText, "#***#"))
        strLine = Split(strText, "#***#")(i)
        Call GetPlainTextSaveSQL(ArrySQL, strLine, lng序号)
    Next
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetPlainTextSaveSQL(ByRef ArraySQL() As String, ByVal strIn As String, ByRef lng序号 As Long) As Boolean
'对纯文本获取将其保存到词句组成的SQL语句，长度大于4000的字符串，分行存储，序号递增之！
'参数： ArraySQL --- SQL 语句数组
'       strIn --- 需要保存的文本
'       lng序号 --- 序号
    
    Dim lngLen As Long, strSub As String, i As Long, lngID As Long
    Dim lngCount As Long, lID As Long
    strIn = Replace(strIn, "'", "' || chr(39) || '")
    strIn = Replace(strIn, vbCrLf, "' || chr(13) || chr(10) || '")  '本来strIn是不允许有vbCrlf的。
    lngLen = Len(strIn)
    
    '按照4000为界分段存储。
    i = 0
    Do While (i * 2000 + 1 <= lngLen)
        lngCount = UBound(ArraySQL) + 1
        ReDim Preserve ArraySQL(1 To lngCount) As String

        strSub = Mid(strIn, i * 2000 + 1, 2000)

        gstrSQL = "Zl_病历词句组成_Insert(" & mlngWordID & "," & lng序号 & ",0,'" & strSub & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL)"
        
        ArraySQL(lngCount) = gstrSQL
       
        lng序号 = lng序号 + 1
        i = i + 1
    Loop
    GetPlainTextSaveSQL = True
End Function

Private Sub cmdWordTag_Click(Index As Integer)
'插入“检查所见”，“诊断意见”，“建议”
    Dim strTag As String
    Dim strTemp
    
    On Error GoTo err
    Select Case Index
    Case 1
        strTag = "<<所见>>"
    Case 2
        strTag = "<<诊断>>"
    Case 3
        strTag = "<<建议>>"
    End Select
    
    txtWord.Text = Left(txtWord.Text, txtWord.SelStart) & strTag & Mid(txtWord.Text, txtWord.SelStart + 1)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

