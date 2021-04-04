VERSION 5.00
Begin VB.Form frm保险项目编辑 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "保险项目编辑"
   ClientHeight    =   5835
   ClientLeft      =   2760
   ClientTop       =   3645
   ClientWidth     =   7500
   Icon            =   "frm保险项目编辑.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cbo类别 
      Height          =   300
      Left            =   4980
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   4590
      Width           =   2415
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   13
      Left            =   1335
      MaxLength       =   50
      TabIndex        =   27
      Tag             =   "剂型"
      Top             =   4590
      Width           =   2415
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   12
      Left            =   4995
      MaxLength       =   20
      TabIndex        =   3
      Tag             =   "医保编码"
      Top             =   1185
      Width           =   2385
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   11
      Left            =   1335
      MaxLength       =   20
      TabIndex        =   23
      Tag             =   "每日最大用量"
      Top             =   4215
      Width           =   1950
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   10
      Left            =   4980
      MaxLength       =   20
      TabIndex        =   21
      Top             =   3840
      Width           =   2430
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   9
      Left            =   1335
      MaxLength       =   20
      TabIndex        =   19
      Tag             =   "最小包装单位"
      Top             =   3840
      Width           =   1950
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   8
      Left            =   1335
      MaxLength       =   50
      TabIndex        =   17
      Top             =   3450
      Width           =   6045
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   7
      Left            =   1335
      MaxLength       =   100
      TabIndex        =   15
      Tag             =   "别名"
      Top             =   3090
      Width           =   6045
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   6
      Left            =   4980
      MaxLength       =   3
      TabIndex        =   25
      Tag             =   "目录分类"
      Top             =   4215
      Width           =   2415
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   5
      Left            =   4995
      MaxLength       =   20
      TabIndex        =   9
      Tag             =   "拼音码"
      Top             =   1965
      Width           =   2385
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   4
      Left            =   1335
      MaxLength       =   20
      TabIndex        =   7
      Tag             =   "五笔码"
      Top             =   1965
      Width           =   1980
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   3
      Left            =   1335
      MaxLength       =   100
      TabIndex        =   13
      Tag             =   "通用英文名"
      Top             =   2715
      Width           =   6045
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   2
      Left            =   1335
      MaxLength       =   100
      TabIndex        =   11
      Tag             =   "通用中文名"
      Top             =   2355
      Width           =   6045
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   34
      Top             =   930
      Width           =   7620
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   90
      TabIndex        =   32
      Top             =   5265
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   0
      Left            =   -120
      TabIndex        =   33
      Top             =   5115
      Width           =   7680
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   0
      Left            =   1335
      MaxLength       =   20
      TabIndex        =   1
      Tag             =   "项目编码"
      Top             =   1185
      Width           =   2010
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   1
      Left            =   1335
      MaxLength       =   100
      TabIndex        =   5
      Tag             =   "项目名称"
      Top             =   1575
      Width           =   6045
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6285
      TabIndex        =   31
      Top             =   5265
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5040
      TabIndex        =   30
      Top             =   5265
      Width           =   1100
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "医保类别(&T)"
      Height          =   180
      Left            =   3960
      TabIndex        =   28
      Top             =   4650
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "剂型(&J)"
      Height          =   180
      Index           =   13
      Left            =   675
      TabIndex        =   26
      Top             =   4650
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "医保编码(&Y)"
      Height          =   180
      Index           =   12
      Left            =   4020
      TabIndex        =   2
      Top             =   1245
      Width           =   990
   End
   Begin VB.Label lblInfor 
      Caption         =   "新增医疗单位申请的商品名,目前只能作全自费项目处理。"
      Height          =   240
      Left            =   1065
      TabIndex        =   35
      Top             =   510
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frm保险项目编辑.frx":000C
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "每日最大用量(&K)"
      Height          =   180
      Index           =   11
      Left            =   -15
      TabIndex        =   22
      Top             =   4275
      Width           =   1350
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "最小计量单位(&F)"
      Height          =   180
      Index           =   10
      Left            =   3600
      TabIndex        =   20
      Tag             =   "最小计量单位"
      Top             =   3900
      Width           =   1350
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "最小包装单位(&X)"
      Height          =   180
      Index           =   9
      Left            =   -15
      TabIndex        =   18
      Top             =   3900
      Width           =   1350
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "包装规格(&G)"
      Height          =   180
      Index           =   8
      Left            =   345
      TabIndex        =   16
      Top             =   3510
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "别名(&A)"
      Height          =   180
      Index           =   7
      Left            =   705
      TabIndex        =   14
      Top             =   3195
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "目录分类(&M)"
      Height          =   180
      Index           =   6
      Left            =   3975
      TabIndex        =   24
      Top             =   4275
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "拼音码(&P)"
      Height          =   180
      Index           =   5
      Left            =   4185
      TabIndex        =   8
      Top             =   2025
      Width           =   810
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "五笔码(&W)"
      Height          =   180
      Index           =   4
      Left            =   525
      TabIndex        =   6
      Top             =   2025
      Width           =   810
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "通用英文名(&E)"
      Height          =   180
      Index           =   3
      Left            =   165
      TabIndex        =   12
      Top             =   2775
      Width           =   1170
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "通用中文名(&Z)"
      Height          =   180
      Index           =   2
      Left            =   165
      TabIndex        =   10
      Top             =   2415
      Width           =   1170
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "项目编码(&U)"
      Height          =   180
      Index           =   0
      Left            =   345
      TabIndex        =   0
      Top             =   1245
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "项目名称(&N)"
      Height          =   180
      Index           =   1
      Left            =   345
      TabIndex        =   4
      Top             =   1590
      Width           =   990
   End
End
Attribute VB_Name = "frm保险项目编辑"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '是否改变了
Dim mintSuccess As Integer
Dim mstr大类编码 As String
Dim mstr商品代码 As String
'
Dim mblnFirst  As Boolean


Private Sub cbo类别_Change()
    mblnChange = True
    SetOk
End Sub

Private Sub cbo类别_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    Call LoadCboData
    mblnChange = False
End Sub
Private Sub LoadCboData()
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select * From 医保项目分类 "
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcnOracle_CQYB
        cbo类别.Clear
        Do While Not .EOF
            cbo类别.AddItem Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称)
            If Nvl(rsTemp!编码) = mstr大类编码 Then
                cbo类别.ListIndex = cbo类别.NewIndex
            End If
            .MoveNext
        Loop
    End With
End Sub
Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub
'
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
     Dim intIndex As Integer
    
    If IsValid() = False Then Exit Sub
    If SaveData() = False Then Exit Sub
    
    mintSuccess = mintSuccess + 1
    mblnChange = False
    For intIndex = 0 To 12
        txtEdit(intIndex).Text = ""
    Next
    If txtEdit(0).Enabled Then
        txtEdit(0).SetFocus
    End If
    SetOk
End Sub



Public Function EditCard(ByVal frmMain As Object, ByVal str商品代码 As String, ByVal str大类编码 As String) As Boolean
    '功能:用来与调用的医保项目管理窗口进行通讯的程序
    '参数:str序号           当前编辑的医保类别的的序号
    '返回值:编辑成功返回True,否则为False
    
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer

    mstr商品代码 = str商品代码
    mstr大类编码 = str大类编码
   
    mintSuccess = 0

    
    If str商品代码 <> "" Then
        gstrSQL = "Select 商品代码, 商品名, 药品通用中文名, 药品通用英文名, 五笔助记码1, 拼音助记码1, 目录分类, 别名, 包装规格, 最小包装单位, 最小计量单位, 每日最大用量, 医保编码,剂型 From 医保服务项目目录 where 商品代码='" & mstr商品代码 & "'"
        
        rsTemp.CursorLocation = adUseClient
        rsTemp.Open gstrSQL, gcnOracle_CQYB, adOpenStatic
        If rsTemp.RecordCount = 0 Then
            ShowMsgbox "该项目已经被他人删除，不能进行修改."
            Exit Function
        End If
        For i = 0 To 13
            txtEdit(i).Text = Nvl(rsTemp.Fields(i))
        Next
    End If
    mblnChange = False
    Me.Show 1, frmMain
    EditCard = mintSuccess > 0
End Function

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    SetOk
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    zlCommFun.OpenIme True
End Sub


Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
End Sub

Private Function IsValid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:验证数据的合法性
    '--入参数:
    '--出参数:
    '--返  回:验证合法,返回True,否则=false
    '-----------------------------------------------------------------------------------------------------------


    Dim intIndex As Integer
    
    Dim strTemp As String
    
      For intIndex = 0 To 13
        strTemp = Trim(txtEdit(intIndex).Text)
        If intIndex = 0 Or intIndex = 1 Then
            If strTemp = "" Then
                ShowMsgbox txtEdit(intIndex).Tag & "必需输入!"
                If txtEdit(intIndex).Enabled Then txtEdit(intIndex).SetFocus
                Exit Function
            End If
        End If
        
        If strTemp <> "" Then
            If LenB(StrConv(strTemp, vbFromUnicode)) > txtEdit(intIndex).MaxLength Then
                ShowMsgbox txtEdit(intIndex).Tag & "超长,最多能输入" & txtEdit(intIndex).MaxLength / 2 & "个汉字或" & txtEdit(intIndex).MaxLength & "个字符!"
                If txtEdit(intIndex).Enabled Then txtEdit(intIndex).SetFocus
                Exit Function
            End If
            If InStr(1, strTemp, "'") <> 0 Then
                ShowMsgbox txtEdit(intIndex).Tag & "不能输入单引号!"
                If txtEdit(intIndex).Enabled Then txtEdit(intIndex).SetFocus
                Exit Function
            End If
        End If
    Next
    IsValid = True
End Function
Private Function SaveData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:保存数据
    '--入参数:
    '--出参数:
    '--返  回:保存成功,返回True,否则=false
    '-----------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim intIndex As Integer
    Dim str自制商品代码 As String
    Dim rsTemp As New ADODB.Recordset
    Dim str目录分类 As String
    If cbo类别.Text = "" Then
    Else
        mstr大类编码 = Split(cbo类别.Text, "-")(0)
    End If
    gstrSQL = "Select 商品代码,目录分类 From 医保服务项目目录 Where 医院大类编码='" & mstr大类编码 & "' And  医保标识 like '__03' and rownum<=1"
    
    zlDataBase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    If rsTemp.EOF Then
        ShowMsgbox "无自费商品!"
        Exit Function
    End If
    str自制商品代码 = Nvl(rsTemp!商品代码)
    str目录分类 = Nvl(rsTemp!目录分类)
    SaveData = False
    
    On Error GoTo errHandle
     
    gstrSQL = "ZL_医保服务项目目录_Insert( "

    '过程参数如下:
    '医院大类编码_IN     IN 医保服务项目目录.医院大类编码%TYPE,
    '    医保编码_IN     IN 医保服务项目目录.医保编码%TYPE,
    '    药品通用中文名_IN       IN 医保服务项目目录.药品通用中文名%TYPE,
    '    药品通用英文名_IN       IN 医保服务项目目录.药品通用英文名%TYPE,
    '    商品代码_IN     IN 医保服务项目目录.商品代码%TYPE,
    '    商品名_IN       IN 医保服务项目目录.商品名%TYPE,
    '    商品曾用名_IN       IN 医保服务项目目录.商品曾用名%TYPE,
    '    别名_IN         IN 医保服务项目目录.别名%TYPE,
    '    包装规格_IN     IN 医保服务项目目录.包装规格%TYPE,
    '    最小包装单位_IN     IN 医保服务项目目录.最小包装单位%TYPE,
    '    最小计量单位_IN     IN 医保服务项目目录.最小计量单位%TYPE,
    '    每日最大用量_IN     IN 医保服务项目目录.每日最大用量%TYPE,
    '    五笔助记码1_IN      IN 医保服务项目目录.五笔助记码1%TYPE,
    '    拼音助记码1_IN      IN 医保服务项目目录.拼音助记码1%TYPE,
    '    剂型_iN
    '    目录分类_IN     IN 医保服务项目目录.目录分类%TYPE
    '    标准代码

    gstrSQL = gstrSQL & "'" & _
            mstr大类编码 & "'," & _
            IIf(Trim(txtEdit(12).Text) = "", "NULL", "'" & Trim(txtEdit(12).Text) & "'") & "," & _
            IIf(Trim(txtEdit(2).Text) = "", "NULL", "'" & Trim(txtEdit(2).Text) & "'") & "," & _
            IIf(Trim(txtEdit(3).Text) = "", "NULL", "'" & Trim(txtEdit(3).Text) & "'") & "," & _
            IIf(Trim(txtEdit(0).Text) = "", "NULL", "'" & Trim(txtEdit(0).Text) & "'") & "," & _
            IIf(Trim(txtEdit(1).Text) = "", "NULL", "'" & Trim(txtEdit(1).Text) & "'") & "," & _
            "NULL" & "," & _
            IIf(Trim(txtEdit(7).Text) = "", "NULL", "'" & Trim(txtEdit(7).Text) & "'") & "," & _
            IIf(Trim(txtEdit(8).Text) = "", "NULL", "'" & Trim(txtEdit(8).Text) & "'") & "," & _
            IIf(Trim(txtEdit(9).Text) = "", "NULL", "'" & Trim(txtEdit(9).Text) & "'") & "," & _
            IIf(Trim(txtEdit(10).Text) = "", "NULL", "'" & Trim(txtEdit(10).Text) & "'") & "," & _
            IIf(Trim(txtEdit(11).Text) = "", "NULL", "'" & Trim(txtEdit(11).Text) & "'") & "," & _
            IIf(Trim(txtEdit(4).Text) = "", "NULL", "'" & Trim(txtEdit(4).Text) & "'") & "," & _
            IIf(Trim(txtEdit(5).Text) = "", "NULL", "'" & Trim(txtEdit(5).Text) & "'") & "," & _
            IIf(Trim(txtEdit(13).Text) = "", "NULL", "'" & Trim(txtEdit(13).Text) & "'") & ",'" & _
            str目录分类 & "'," & _
            IIf(str自制商品代码 = "", "NULL", "'" & str自制商品代码 & "'") & "" & _
            ")"
    
    Call SQLTest(App.ProductName, "新增保险项目", gstrSQL)
    gcnOracle_CQYB.Execute gstrSQL, , adCmdStoredProc
    Call SQLTest
    
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub SetOk()
    cmdOK.Enabled = mblnChange
End Sub
