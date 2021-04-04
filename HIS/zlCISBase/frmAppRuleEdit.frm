VERSION 5.00
Begin VB.Form frmAppRuleEdit 
   BorderStyle     =   0  'None
   Caption         =   "仪器质控规则"
   ClientHeight    =   5055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8265
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox chkUse 
      Caption         =   "是否使用"
      Height          =   180
      Left            =   6960
      TabIndex        =   23
      ToolTipText     =   "在计算时，是否使用此规则"
      Top             =   75
      Width           =   1065
   End
   Begin VB.TextBox txt提示 
      Enabled         =   0   'False
      Height          =   780
      Index           =   1
      Left            =   480
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   4140
      Width           =   7575
   End
   Begin VB.CheckBox chk结束 
      Caption         =   "结束，并给予提示:"
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   21
      Top             =   3900
      Width           =   2640
   End
   Begin VB.TextBox txt标记规则 
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      Left            =   4395
      MaxLength       =   20
      TabIndex        =   20
      Top             =   3510
      Width           =   1860
   End
   Begin VB.ComboBox cbo标记级 
      Height          =   300
      Index           =   1
      Left            =   1245
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   3510
      Width           =   2025
   End
   Begin VB.CheckBox chk多水平 
      Caption         =   "多个水平"
      Enabled         =   0   'False
      Height          =   240
      Left            =   6960
      TabIndex        =   7
      Top             =   690
      Width           =   1065
   End
   Begin VB.TextBox txt提示 
      Enabled         =   0   'False
      Height          =   780
      Index           =   0
      Left            =   480
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   2340
      Width           =   7575
   End
   Begin VB.CheckBox chk结束 
      Caption         =   "结束，并给予提示:"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   14
      Top             =   2100
      Width           =   2640
   End
   Begin VB.TextBox txt标记规则 
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   4395
      MaxLength       =   20
      TabIndex        =   13
      Top             =   1695
      Width           =   1860
   End
   Begin VB.ComboBox cbo标记级 
      Height          =   300
      Index           =   0
      Left            =   1245
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1695
      Width           =   2025
   End
   Begin VB.ComboBox cbo批范围 
      Height          =   300
      Left            =   6960
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1005
      Width           =   1050
   End
   Begin VB.ComboBox cbo性质 
      Height          =   300
      Left            =   915
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1005
      Width           =   2370
   End
   Begin VB.TextBox txt判断 
      Height          =   555
      Left            =   480
      MaxLength       =   80
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   375
      Width           =   6225
   End
   Begin VB.ComboBox cbo规则 
      Height          =   300
      Left            =   4395
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1005
      Width           =   2340
   End
   Begin VB.Label lbl处理 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3.不符合(N)判断规则的处理:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   135
      TabIndex        =   16
      Top             =   3225
      Width           =   2580
   End
   Begin VB.Label lbl标记规则 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "标记规则                     (可不与当前规则一致)"
      Height          =   180
      Index           =   1
      Left            =   3615
      TabIndex        =   19
      Top             =   3570
      Width           =   4410
   End
   Begin VB.Label lbl标记级 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "标记等级"
      Height          =   180
      Index           =   1
      Left            =   480
      TabIndex        =   17
      Top             =   3570
      Width           =   720
   End
   Begin VB.Label lbl处理 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2.符合(Y)判断规则的处理:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   9
      Top             =   1425
      Width           =   2385
   End
   Begin VB.Label lbl基本信息 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.判断描述与对应规则:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   135
      TabIndex        =   0
      Top             =   105
      Width           =   2070
   End
   Begin VB.Label lbl标记规则 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "标记规则                     (可不与当前规则一致)"
      Height          =   180
      Index           =   0
      Left            =   3615
      TabIndex        =   12
      Top             =   1755
      Width           =   4410
   End
   Begin VB.Label lbl标记级 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "标记等级"
      Height          =   180
      Index           =   0
      Left            =   480
      TabIndex        =   10
      Top             =   1755
      Width           =   720
   End
   Begin VB.Label lbl批范围 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "标本范围:"
      Height          =   180
      Left            =   6960
      TabIndex        =   6
      Top             =   435
      Width           =   810
   End
   Begin VB.Label lbl性质 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "性质"
      Height          =   180
      Left            =   480
      TabIndex        =   2
      Top             =   1065
      Width           =   360
   End
   Begin VB.Label lbl规则 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "对应规则"
      Height          =   180
      Left            =   3615
      TabIndex        =   4
      Top             =   1065
      Width           =   720
   End
End
Attribute VB_Name = "frmAppRuleEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngRuleId As Long          '当前显示的规则id
Private mlngParent As Long          '当前规则的上级id
Private mlngDevId As Long           '当前显示的仪器id
Private mlngGroupID As Long         '当前显示的分组ID

Dim lngCount As Long

'--------------------------------------------
'以下为窗体公共方法
'--------------------------------------------
Public Function zlRefresh(lngRuleId As Long) As Boolean
    '功能：根据项目id刷新当前显示内容
    Dim rsTemp As New ADODB.Recordset
    mlngRuleId = lngRuleId: mlngParent = 0
    
    '清除此前项目的显示
    Me.txt判断.Text = "":
    Me.cbo性质.Clear: Me.cbo性质.ListIndex = -1
    Me.cbo规则.Clear: Me.cbo规则.ListIndex = -1
    Me.cbo批范围.ListIndex = 0: Me.chk多水平.Value = vbUnchecked
    Me.cbo标记级(0).ListIndex = 0: Me.txt标记规则(0).Text = ""
    Me.chk结束(0).Value = vbChecked: Me.txt提示(0).Text = ""
    Me.cbo标记级(1).ListIndex = 0: Me.txt标记规则(1).Text = ""
    Me.chk结束(1).Value = vbChecked: Me.txt提示(1).Text = ""
    Me.chkUse = vbUnchecked
    If lngRuleId = 0 Then zlRefresh = True: Exit Function
    
    '获取指定项目的信息
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select R.上级id, D.质控水平数, R.判断, R.规则id, R.性质, R.批范围, R.多水平, R.Y标记级, R.Y规则, R.Y结束, R.Y提示," & vbNewLine & _
            "       R.N标记级, R.N规则, R.N结束, R.N提示, R.是否使用 " & vbNewLine & _
            "From 检验仪器规则 R, 检验仪器 D" & vbNewLine & _
            "Where R.仪器id = D.ID And R.ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngRuleId)
    With rsTemp
        Me.chk多水平.Tag = 0
        If .RecordCount > 0 Then
            mlngParent = Val("" & !上级ID)
            Me.chk多水平.Tag = Val("" & !质控水平数)
            If Val("" & !是否使用) = 1 Then Me.chkUse = vbChecked
            
            Me.txt判断.Text = "" & !判断
            Select Case "" & !性质
            Case "0": Me.cbo性质.AddItem "0-开始规则": Me.cbo性质.ListIndex = 0
            Case "Y": Me.cbo性质.AddItem "Y-上一步符合时执行的规则": Me.cbo性质.ListIndex = 0
            Case "N": Me.cbo性质.AddItem "N-上一步不符时执行的规则": Me.cbo性质.ListIndex = 0
            Case "1": Me.cbo性质.AddItem "1-附加规则": Me.cbo性质.ListIndex = 0
            End Select
            Me.cbo规则.Tag = "" & !规则id
            For lngCount = 0 To Me.cbo批范围.ListCount - 1
                If lngCount = Val("" & !批范围) - 1 Then Me.cbo批范围.ListIndex = lngCount: Exit For
            Next
            If Val("" & !多水平) = 1 Then Me.chk多水平.Value = vbChecked
            
            For lngCount = 0 To Me.cbo标记级(0).ListCount - 1
                If lngCount = Val("" & !Y标记级) Then Me.cbo标记级(0).ListIndex = lngCount: Exit For
            Next
            Me.txt标记规则(0).Text = "" & !Y规则
            Me.chk结束(0).Value = IIf(Val("" & !Y结束) = 0, vbUnchecked, vbChecked)
            Me.txt提示(0).Text = "" & !Y提示
            
            For lngCount = 0 To Me.cbo标记级(1).ListCount - 1
                If lngCount = Val("" & !n标记级) Then Me.cbo标记级(1).ListIndex = lngCount: Exit For
            Next
            Me.txt标记规则(1).Text = "" & !N规则
            Me.chk结束(1).Value = IIf(Val("" & !N结束) = 0, vbUnchecked, vbChecked)
            Me.txt提示(1).Text = "" & !N提示
            
            If "" & !性质 = "1" Then
                Me.chk结束(0).Value = vbChecked: Me.chk结束(0).Enabled = False
                Me.chk结束(1).Value = vbChecked: Me.chk结束(1).Enabled = False
            Else
                Me.chk结束(0).Enabled = True
                Me.chk结束(1).Enabled = True
            End If
        End If
    End With
    
    '目前只允许常用规则组成多步骤规则，计算控制界限和累积和规则只能作为附加规则，且只能在每批水平数>1时选择计算控制界限规则
    If Left(Me.cbo性质.Text, 1) = "1" Then
        If Val(Me.chk多水平.Tag) > 1 Then
            gstrSql = "Select ID, RPad(名称, 200, ' ') || 多水平 || ',' || N As 名称 From 检验质控规则 Order By 种类, 编码"
        Else
            gstrSql = "Select ID, RPad(名称, 200, ' ') || 多水平 || ',' || N As 名称 From 检验质控规则 Where 种类 In (1, 3) Order By 种类, 编码"
        End If
    Else
        gstrSql = "Select ID, RPad(名称, 200, ' ') || 多水平 || ',' || N As 名称 From 检验质控规则 Where 种类 = 1 Order By 种类, 编码"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Me.cbo规则.Clear
        Do While Not .EOF
            Me.cbo规则.AddItem "" & !名称
            Me.cbo规则.ItemData(Me.cbo规则.NewIndex) = Val("" & !ID)
            If Val("" & !ID) = Val(Me.cbo规则.Tag) Then Me.cbo规则.ListIndex = Me.cbo规则.NewIndex
            .MoveNext
        Loop
    End With
    
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart(blnAdd As Boolean, lngRuleId As Long, lngDevId As Long, lngGroupID As Long, Optional blnSingle As Boolean) As Boolean
    '功能：开始项目编辑
    '参数： blnAdd-是否增加，否则为修改
    '       lngRuleId-增加时,为当前规则上级规则，没有上级时为0；修改时为当前规则的ID
    '       lngDevId-当前增加物品的所属设备id
    '       blnSingle-在增加时有效，指明是增加多规则项还是单规则
    Dim rsTemp As New ADODB.Recordset
    Dim strKind As String
    
    mlngDevId = lngDevId
    mlngGroupID = lngGroupID
    Err = 0: On Error GoTo ErrHand
    
    '设备水平数
    If blnAdd Then
        Me.chk多水平.Tag = 0
        gstrSql = "Select 质控水平数 From 检验仪器 Where ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngDevId)
        With rsTemp
            If .RecordCount > 0 Then Me.chk多水平.Tag = Val("" & !质控水平数)
        End With
    
        Me.cbo性质.Clear
        If blnSingle Then
            Me.cbo性质.AddItem "1-附加规则"
        Else
            If lngRuleId = 0 Then
                gstrSql = "Select Decode(Nvl(Count(*), 0), 0, 1, 0) As 许可 From 检验仪器规则 Where 仪器id = [1] And 项目id =[2] And 性质 = '0'"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngDevId, lngGroupID)
                With rsTemp
                    If .RecordCount > 0 Then
                        If !许可 Then Me.cbo性质.AddItem "0-开始规则"
                    End If
                End With
            Else
                gstrSql = "Select Decode(P.Y结束, 1, 0, Decode(Y存在, 0, 1, 0)) As Y可, Decode(P.N结束, 1, 0, Decode(N存在, 0, 1, 0)) As N可" & vbNewLine & _
                        "From (Select Nvl(Y结束, 0) As Y结束, Nvl(N结束, 0) As N结束 From 检验仪器规则 Where ID = [1]) P," & vbNewLine & _
                        "     (Select Nvl(Sum(Decode(性质, 'Y', 1, 0)), 0) As Y存在, Nvl(Sum(Decode(性质, 'N', 1, 0)), 0) As N存在" & vbNewLine & _
                        "       From 检验仪器规则" & vbNewLine & _
                        "       Where 上级id = [1]) C"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngRuleId)
                With rsTemp
                    If .RecordCount > 0 Then
                        If !Y可 Then Me.cbo性质.AddItem "Y-上一步符合时执行的规则"
                        If !N可 Then Me.cbo性质.AddItem "N-上一步不符时执行的规则"
                    End If
                End With
            End If
            If Me.cbo性质.ListCount = 0 Then
                MsgBox "已经存在相应的规则或上一步已经结束！", vbInformation, gstrSysName
                zlEditStart = False: Exit Function
            End If
        End If
        Me.cbo性质.ListIndex = 0
        
        '目前只允许常用规则组成多步骤规则，计算控制界限和累积和规则只能作为附加规则，且只能在每批水平数>1时选择计算控制界限规则
        If Left(Me.cbo性质.Text, 1) = "1" Then
'            If Val(Me.chk多水平.Tag) > 1 Then
'                gstrSql = "Select ID, RPad(名称, 200, ' ') || 多水平 || ',' || N As 名称 From 检验质控规则 Where 种类 In (2, 3) Order By 种类, 编码"
'            Else
'                gstrSql = "Select ID, RPad(名称, 200, ' ') || 多水平 || ',' || N As 名称 From 检验质控规则 Where 种类 = 3 Order By 种类, 编码"
'            End If

            If Val(Me.chk多水平.Tag) > 1 Then
                '每批水平>1时，可以选择所有规则作为附加规则
                gstrSql = "Select ID, RPad(名称, 200, ' ') || 多水平 || ',' || N As 名称 From 检验质控规则 Order By 种类, 编码"
            Else
                '每批水平=1时，可以选择常用质控规则和累积和规则作为附加规则
                gstrSql = "Select ID, RPad(名称, 200, ' ') || 多水平 || ',' || N As 名称 From 检验质控规则 Where 种类 In (1, 3) Order By 种类, 编码"
            End If
        Else
            gstrSql = "Select ID, RPad(名称, 200, ' ') || 多水平 || ',' || N As 名称 From 检验质控规则 Where 种类 = 1 Order By 种类, 编码"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
        With rsTemp
            Me.cbo规则.Clear
            Do While Not .EOF
                Me.cbo规则.AddItem "" & !名称
                Me.cbo规则.ItemData(Me.cbo规则.NewIndex) = Val("" & !ID)
                .MoveNext
            Loop
        End With
        If Me.cbo规则.ListCount = 0 Then
            MsgBox "请首先初始化检验质控规则！", vbInformation, gstrSysName
            zlEditStart = False: Exit Function
        Else
            Me.cbo规则.ListIndex = 0
        End If
        mlngParent = lngRuleId
    
        Me.txt判断.Text = ""
        Me.cbo批范围.ListIndex = 0
        Me.cbo标记级(0).ListIndex = 0: Me.txt标记规则(0).Text = ""
        Me.chk结束(0).Value = vbChecked: Me.txt提示(0).Text = ""
        Me.cbo标记级(1).ListIndex = 0: Me.txt标记规则(1).Text = ""
        Me.chk结束(1).Value = vbChecked: Me.txt提示(1).Text = ""
        If blnSingle Then
            Me.chk结束(0).Enabled = False
            Me.chk结束(1).Enabled = False
        Else
            Me.chk结束(0).Enabled = True
            Me.chk结束(1).Enabled = True
        End If
        Me.chkUse.Value = vbChecked
    Else
        strKind = Left(Me.cbo性质.Text, 1)
        Me.cbo性质.Clear
        Select Case strKind
        Case "0": Me.cbo性质.AddItem "0-开始规则": Me.cbo性质.ListIndex = 0
        Case "Y", "N"
            gstrSql = "Select Decode(P.Y结束, 1, 0, Decode(Y存在, 0, 1, 0)) As Y可, Decode(P.N结束, 1, 0, Decode(N存在, 0, 1, 0)) As N可" & vbNewLine & _
                    "From (Select Nvl(Y结束, 0) As Y结束, Nvl(N结束, 0) As N结束 From 检验仪器规则 Where ID = [1]) P," & vbNewLine & _
                    "     (Select Nvl(Sum(Decode(性质, 'Y', 1, 0)), 0) As Y存在, Nvl(Sum(Decode(性质, 'N', 1, 0)), 0) As N存在" & vbNewLine & _
                    "       From 检验仪器规则" & vbNewLine & _
                    "       Where 上级id = [1]) C"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngParent)
            With rsTemp
                If .RecordCount > 0 Then
                    '实际只可能有一个许可存在
                    If !Y可 Then Me.cbo性质.AddItem "Y-上一步符合时执行的规则"
                    If !N可 Then Me.cbo性质.AddItem "N-上一步不符时执行的规则"
                End If
            End With
            If strKind = "Y" Then
                Me.cbo性质.AddItem "Y-上一步符合时执行的规则": Me.cbo性质.ListIndex = Me.cbo性质.NewIndex
            Else
                Me.cbo性质.AddItem "N-上一步不符时执行的规则": Me.cbo性质.ListIndex = Me.cbo性质.NewIndex
            End If
        Case "1": Me.cbo性质.AddItem "1-附加规则": Me.cbo性质.ListIndex = 0
        End Select
    End If
    
    Me.Enabled = True: Me.Tag = IIf(blnAdd, "增加", "修改")
    Me.BackColor = RGB(250, 250, 250)
    Me.chk多水平.BackColor = Me.BackColor
    Me.chk结束(0).BackColor = Me.BackColor
    Me.chk结束(1).BackColor = Me.BackColor
    
    Me.txt判断.SetFocus
    zlEditStart = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditStart = False: Exit Function
End Function

Public Sub zlEditCancel()
    '功能：放弃正在进行的编辑
    Me.Enabled = False: Me.Tag = ""
    Me.BackColor = &H8000000F
    Me.chk多水平.BackColor = Me.BackColor
    Me.chk结束(0).BackColor = Me.BackColor
    Me.chk结束(1).BackColor = Me.BackColor
    
    Call Me.zlRefresh(mlngRuleId)
End Sub

Public Function zlEditSave() As Long
    '功能：保存正在进行的编辑,并返回正在编辑项目id,保存失败返回0
    Dim lngNewId As Long, blnMatch As Boolean

    '一般特性检查
    If Trim(Me.txt判断.Text) = "" Then
        MsgBox "请输入判断！", vbInformation, gstrSysName
        Me.txt判断.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt判断.Text), vbFromUnicode)) > Me.txt判断.MaxLength Then
        MsgBox "判断超长（最多" & Me.txt判断.MaxLength & "个字符或等长汉字）！", vbInformation, gstrSysName
        Me.txt判断.SetFocus: zlEditSave = 0: Exit Function
    End If
    Me.txt判断.Text = Replace(Me.txt判断.Text, vbCrLf, "")
    Me.txt判断.Text = Replace(Me.txt判断.Text, vbCr, "")
    Me.txt判断.Text = Replace(Me.txt判断.Text, vbLf, "")
    
    If Me.cbo性质.ListIndex = -1 Then
        MsgBox "请指明性质！", vbInformation, gstrSysName
        Me.txt判断.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Me.cbo规则.ListIndex = -1 Then
        MsgBox "请指明规则！", vbInformation, gstrSysName
        Me.cbo规则.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Left(Me.cbo性质.Text, 1) <> "1" Then
        blnMatch = False
        If Me.chk多水平.Value = vbChecked Then
            If Val(Me.cbo批范围.Tag) > Me.cbo批范围.ListIndex * Val(Me.chk多水平.Tag) And Val(Me.cbo批范围.Tag) <= (Me.cbo批范围.ListIndex + 1) * Val(Me.chk多水平.Tag) Then blnMatch = True
        Else
            If Val(Me.cbo批范围.Tag) = (Me.cbo批范围.ListIndex + 1) Then blnMatch = True
        End If
        If blnMatch = False Then
            MsgBox "批范围需要和规则要求的标本范围匹配！", vbInformation, gstrSysName
            Call chk多水平_Click
            Me.cbo批范围.SetFocus: zlEditSave = 0: Exit Function
        End If
    Else
        If Val(Me.cbo批范围.Tag) > (Me.cbo批范围.ListIndex + 1) * IIf(Me.chk多水平.Value = vbChecked, Val(Me.chk多水平.Tag), 1) Then
            MsgBox "批范围需要和规则要求的标本范围匹配！", vbInformation, gstrSysName
            Me.cbo批范围.SetFocus: zlEditSave = 0: Exit Function
        End If
    End If
    
    For lngCount = 0 To 1
        If Me.txt标记规则(lngCount).Enabled Then
            If Trim(Me.txt标记规则(lngCount).Text = "") Then
                MsgBox "当标记警告或失控时，需要指明标记规则！", vbInformation, gstrSysName
                Me.txt标记规则(lngCount).SetFocus: zlEditSave = 0: Exit Function
            End If
            If LenB(StrConv(Trim(Me.txt标记规则(lngCount).Text), vbFromUnicode)) > Me.txt标记规则(lngCount).MaxLength Then
                MsgBox "标记规则超长（最多" & Me.txt标记规则(lngCount).MaxLength & "个字符或等长汉字）！", vbInformation, gstrSysName
                Me.txt标记规则(lngCount).SetFocus: zlEditSave = 0: Exit Function
            End If
        End If
        If Me.txt提示(lngCount).Enabled Then
            If Trim(Me.txt提示(lngCount).Text = "") Then
                MsgBox "当结束时，需要填写提示内容！", vbInformation, gstrSysName
                Me.txt提示(lngCount).SetFocus: zlEditSave = 0: Exit Function
            End If
            If LenB(StrConv(Trim(Me.txt提示(lngCount).Text), vbFromUnicode)) > Me.txt提示(lngCount).MaxLength Then
                MsgBox "提示超长（最多" & Me.txt提示(lngCount).MaxLength & "个字符或等长汉字）！", vbInformation, gstrSysName
                Me.txt提示(lngCount).SetFocus: zlEditSave = 0: Exit Function
            End If
        End If
        Me.txt提示(lngCount).Text = Replace(Me.txt提示(lngCount).Text, vbCrLf, "")
        Me.txt提示(lngCount).Text = Replace(Me.txt提示(lngCount).Text, vbCr, "")
        Me.txt提示(lngCount).Text = Replace(Me.txt提示(lngCount).Text, vbLf, "")
    Next
    
    '数据保存语句组织
    gstrSql = "'" & Trim(Me.txt判断.Text) & "','" & Left(Me.cbo性质.Text, 1) & "'," & Me.cbo规则.ItemData(Me.cbo规则.ListIndex)
    gstrSql = gstrSql & "," & Me.cbo批范围.ListIndex + 1 & "," & IIf(Me.chk多水平.Value = vbChecked, 1, 0)
    For lngCount = 0 To 1
        gstrSql = gstrSql & "," & Me.cbo标记级(lngCount).ListIndex
        If Me.cbo标记级(lngCount).ListIndex > 0 Then
            gstrSql = gstrSql & ",'" & Trim(Me.txt标记规则(lngCount).Text) & "'"
        Else
            gstrSql = gstrSql & ",''"
        End If
        If Me.chk结束(lngCount).Value = vbChecked Then
            gstrSql = gstrSql & ",1,'" & Trim(Me.txt提示(lngCount).Text) & "'"
        Else
            gstrSql = gstrSql & ",0,''"
        End If
        
    Next
    
    gstrSql = gstrSql & "," & IIf(Me.chkUse.Value = vbChecked, 1, 0)
    
    If Me.Tag = "增加" Then
        lngNewId = zlDatabase.GetNextId("检验仪器规则")
        gstrSql = "Zl_检验仪器规则_Edit(1," & lngNewId & "," & mlngParent & "," & mlngDevId & "," & mlngGroupID & "," & gstrSql & ")"
    Else
        gstrSql = "Zl_检验仪器规则_Edit(2," & mlngRuleId & "," & mlngParent & "," & mlngDevId & "," & mlngGroupID & "," & gstrSql & ")"
    End If
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    If Me.Tag = "增加" Then mlngRuleId = lngNewId
    Me.Enabled = False: Me.Tag = ""
    Me.BackColor = &H8000000F
    Me.chk多水平.BackColor = Me.BackColor
    Me.chk结束(0).BackColor = Me.BackColor
    Me.chk结束(1).BackColor = Me.BackColor
    
    zlEditSave = mlngRuleId: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

'--------------------------------------------
'以下为窗体控件响应事件
'--------------------------------------------
Private Sub cbo标记级_Click(Index As Integer)
    If Me.cbo标记级(Index).ListIndex <= 0 Then
        Me.txt标记规则(Index).Enabled = False
    Else
        Me.txt标记规则(Index).Enabled = True
    End If
End Sub

Private Sub cbo标记级_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo标记级_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cbo规则_Click()
    Dim strRule As String
    
    '编辑状态，自动设置违背的规则
    If Me.Tag <> "" Then
        strRule = Trim(Left(Me.cbo规则.Text, 100))
        Me.txt标记规则(0).Text = strRule
        Me.txt标记规则(1).Text = strRule
    End If
    
    '记录当前规则要求的检测个数,判断多水平允许
    Me.cbo批范围.Tag = Split(Trim(Mid(Me.cbo规则.Text, 200)), ",")(1)
    If Val(Split(Trim(Mid(Me.cbo规则.Text, 200)), ",")(0)) = 0 Or Val(Me.chk多水平.Tag) <= 1 Then
        Me.chk多水平.Value = vbUnchecked: Me.chk多水平.Enabled = False
    Else
        Me.chk多水平.Enabled = True
    End If
    
    Call chk多水平_Click
End Sub

Private Sub cbo规则_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo规则_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cbo批范围_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo批范围_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cbo性质_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo性质_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub chk多水平_Click()
    Dim lngBatch As Long
    If Val(Me.cbo批范围.Tag) <> 0 Then
        If Me.chk多水平.Value = vbUnchecked Then
            lngBatch = Val(Me.cbo批范围.Tag) - 1
        Else
            lngBatch = Int(Val(Me.cbo批范围.Tag) / Val(Me.chk多水平.Tag) + 0.9) - 1
        End If
        If lngBatch < 0 Then lngBatch = 0
        Me.cbo批范围.ListIndex = lngBatch
    End If
End Sub

Private Sub chk多水平_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chk多水平_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub chk结束_Click(Index As Integer)
    If Me.chk结束(Index).Value = vbChecked Then
        Me.txt提示(Index).Enabled = True
    Else
        Me.txt提示(Index).Enabled = False
    End If
End Sub

Private Sub chk结束_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chk结束_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub Form_Load()
    mlngRuleId = 0: mlngDevId = 0: mlngGroupID = 0
        
    Me.cbo批范围.AddItem "当前批"
    For lngCount = 2 To 31
        Me.cbo批范围.AddItem "近" & lngCount & "批"
    Next
    
    Me.cbo标记级(0).AddItem "0-不标记"
    Me.cbo标记级(0).AddItem "1-标记为警告"
    Me.cbo标记级(0).AddItem "2-标记为失控"
    
    Me.cbo标记级(1).AddItem "0-不标记"
    Me.cbo标记级(1).AddItem "1-标记为警告"
    Me.cbo标记级(1).AddItem "2-标记为失控"
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    Me.cbo批范围.Left = Me.ScaleWidth - Me.cbo批范围.Width - 150
    Me.lbl批范围.Left = Me.cbo批范围.Left
    Me.chk多水平.Left = Me.cbo批范围.Left
    Me.txt判断.Width = Me.cbo批范围.Left - Me.txt判断.Left - 300
    Me.cbo规则.Width = Me.txt判断.Left + Me.txt判断.Width - Me.cbo规则.Left
    
    Me.txt提示(0).Width = Me.ScaleWidth - Me.txt提示(0).Left - 150
    Me.txt提示(1).Width = Me.ScaleWidth - Me.txt提示(1).Left - 150
    
    Me.chkUse.Left = Me.chk多水平.Left
End Sub

Private Sub txt标记规则_GotFocus(Index As Integer)
    Me.txt标记规则(Index).SelStart = 0: Me.txt标记规则(Index).SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt标记规则_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt判断_GotFocus()
    Me.txt判断.SelStart = 0: Me.txt判断.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt判断_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt提示_GotFocus(Index As Integer)
    Me.txt提示(Index).SelStart = 0: Me.txt提示(Index).SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt提示_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
