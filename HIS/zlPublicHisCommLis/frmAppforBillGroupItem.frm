VERSION 5.00
Begin VB.Form frmAppforBillGroupItem 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "分组设置"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   0
      TabIndex        =   5
      Top             =   1710
      Width           =   3705
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1230
      TabIndex        =   3
      Top             =   990
      Width           =   2055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   780
      TabIndex        =   2
      Top             =   2010
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2280
      TabIndex        =   1
      Top             =   2010
      Width           =   1335
   End
   Begin VB.TextBox txtNO 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1230
      TabIndex        =   0
      Top             =   360
      Width           =   2025
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "分类名称:"
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
      Left            =   90
      TabIndex        =   6
      Top             =   1050
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "编码:"
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
      Left            =   480
      TabIndex        =   4
      Top             =   420
      Width           =   600
   End
End
Attribute VB_Name = "frmAppforBillGroupItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnfrmShow As Boolean                     '窗体是否显示
Private mlngkeyID As Long                          '分组ID
Private mstrNO As String                           '编码
Private mstrName As String                         '名称
Private mlngMainID As Long                         '申请单id
Private mstrNametext As String

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If SaveDate = True Then
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    If mblnfrmShow = False Then
        If mlngkeyID = 0 Then
            Call GetMaxNO
            Me.TxtName.SetFocus
        Else
            Me.txtNO = mstrNO
            Me.TxtName = mstrName
        End If
        mblnfrmShow = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnfrmShow = False
End Sub
Private Sub txtName_Change()
    If LenB(StrConv(TxtName.Text, vbFromUnicode)) > 20 Then MsgBox "名称不能超过20个字节!", vbExclamation + vbOKOnly, "名称过长": TxtName.Text = mstrNametext
End Sub
Private Sub txtName_GotFocus()
    TxtName.SelStart = 0
    TxtName.SelLength = Len(TxtName)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdOK_Click
    End If
    mstrNametext = TxtName.Text
End Sub

Private Function SaveDate() As Boolean
          Dim strSQL As String
              
1         On Error GoTo SaveDate_Error

2         If Trim(Me.txtNO.Text) = "" Then
3             MsgBox "请输入编码后才能保存!", vbInformation, "增加申请单分组"
4             Me.txtNO.SetFocus
5             Exit Function
6         End If
          
7         If Trim(Me.TxtName.Text) = "" Then
8             MsgBox "请输入名称后才能保存!", vbInformation, "增加申请单分组"
9             Me.TxtName.SetFocus
10            Exit Function
11        End If
          
          '保存
12        strSQL = "Zl_检验申请单分组_EDIT(1," & mlngkeyID & ",'" & Me.txtNO & "','" & Me.TxtName & "'," & mlngMainID & ")"
13        ComExecuteProc Sel_Lis_DB, strSQL, "保存申请分类"
14        SaveDBLog 18, 6, 0, "新增", "新增项目分组:" & TxtName.Text, 1012, "申请单设置"
15        SaveDate = True


16        Exit Function
SaveDate_Error:
17        Call writeErrLog("zl9LisInsideComm", "frmAppforBillGroupItem", "执行(SaveDate)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
18        Err.Clear
          
End Function

Private Sub txtNO_GotFocus()
    txtNO.SelStart = 0
    txtNO.SelLength = Len(txtNO)
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtName.SetFocus
    End If
End Sub

Private Sub GetMaxNO()
          '功能：         初始化数据
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
              
1         On Error GoTo GetMaxNO_Error

2         strSQL = "select nvl(max(编码),0) 编码 from 检验申请单分组 "
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "检验申请单分组")
4         If rsTmp("编码") = 0 Then
5             Me.txtNO = "001"
6         Else
7             Me.txtNO = Format(Val(rsTmp("编码")) + 1, "000")
8         End If


9         Exit Sub
GetMaxNO_Error:
10        Call writeErrLog("zl9LisInsideComm", "frmAppforBillGroupItem", "执行(GetMaxNO)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
11        Err.Clear
          
End Sub

Public Sub showMe(objfrm As Object, lngMainID As Long, lngID As Long, strNO As String, strName As String)
    '功能           打开主窗体
    
    mlngMainID = lngMainID
    mlngkeyID = lngID
    mstrNO = strNO
    mstrName = strName
    Me.Show vbModal, objfrm
End Sub


