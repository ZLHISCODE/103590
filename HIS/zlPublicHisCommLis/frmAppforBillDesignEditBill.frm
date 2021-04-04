VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAppforBillDesignEditBill 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "增加申请单分类"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkTre 
      Caption         =   "耐受申请单"
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
      Left            =   3180
      TabIndex        =   10
      Top             =   300
      Width           =   1545
   End
   Begin MSComDlg.CommonDialog ComDialPublic 
      Left            =   4650
      Top             =   990
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picColour 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1650
      ScaleHeight     =   345
      ScaleWidth      =   1335
      TabIndex        =   8
      Top             =   1200
      Width           =   1365
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
      Left            =   1650
      TabIndex        =   6
      Top             =   240
      Width           =   1395
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
      Left            =   3570
      TabIndex        =   4
      Top             =   2010
      Width           =   1335
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
      Left            =   1920
      TabIndex        =   3
      Top             =   2010
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -120
      TabIndex        =   2
      Top             =   1800
      Width           =   5265
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
      Left            =   1650
      TabIndex        =   1
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label lblcolour 
      AutoSize        =   -1  'True
      Caption         =   "单击修改颜色"
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
      Left            =   3090
      TabIndex        =   9
      Top             =   1260
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "单据颜色:"
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
      Left            =   420
      TabIndex        =   7
      Top             =   1260
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
      Left            =   900
      TabIndex        =   5
      Top             =   300
      Width           =   600
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
      Left            =   420
      TabIndex        =   0
      Top             =   780
      Width           =   1080
   End
End
Attribute VB_Name = "frmAppforBillDesignEditBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnfrmShow As Boolean                      '窗体是否显示
Private mlngkeyID As Long                           '分类ID
Private mstrNO As String                            '编码
Private mstrName As String                          '名称
Private mlngDeptID As Long                          '科室ID
Private mlngColour As Long                          '颜色
Private mblnTrs As Boolean                          '是否耐受试验申请单

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
            Me.txtNO.SetFocus
        Else
            Me.txtNO = mstrNO
            Me.txtName = mstrName
            Me.picColour.BackColor = mlngColour
            Me.chkTre.value = IIf(mblnTrs, 1, 0)
        End If
        If Not VerCompare(gSysInfo.VersionHIS, "10.35.90") <> -1 Then   '低于10.35.90版本不支持耐受试验申请单设置
            Me.chkTre.value = 0
            Me.chkTre.Visible = False
        End If
        mblnfrmShow = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnfrmShow = False
    mblnTrs = False
End Sub

Private Sub Label4_Click()

End Sub

Private Sub lblcolour_Click()
    picColour.BackColor = GetSelColour(picColour.BackColor)
End Sub

Private Sub txtName_GotFocus()
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdOK_Click
    End If
End Sub

Private Function SaveDate() As Boolean
          Dim strSQL As String
          
1         On Error GoTo SaveDate_Error

2         If Trim(Me.txtNO.Text) = "" Then
3             MsgBox "请输入编码后才能保存!", vbInformation, "增加申请单"
4             Me.txtNO.SetFocus
5             Exit Function
6         End If
          
7         If Trim(Me.txtName.Text) = "" Then
8             MsgBox "请输入名称后才能保存!", vbInformation, "增加申请单"
9             Me.txtName.SetFocus
10            Exit Function
11        End If
          
          '保存
12        strSQL = "Zl_检验申请单_EDIT('" & IIf(mlngkeyID = 0, 1, 2) & "','" & mlngkeyID & "'," & IIf(mlngDeptID = 0, "NULL", "'" & mlngDeptID & "'") & _
                          ",'" & Me.txtNO & "','" & Me.txtName & "','" & picColour.BackColor & "'," & Me.chkTre.value & ")"
13        ComExecuteProc Sel_Lis_DB, strSQL, "保存申请分类"
          
14        If mlngkeyID = 0 Then
15            SaveDBLog 18, 6, 0, "新增", "新增申请单:" & txtName.Text, 1012, "申请单设置"
16        Else
17            SaveDBLog 18, 6, 0, "修改", "修改申请单:" & txtName.Text, 1012, "申请单设置"
18        End If
          
19        SaveDate = True


20        Exit Function
SaveDate_Error:
21        Call writeErrLog("zl9LisInsideComm", "frmAppforBillDesignEditBill", "执行(SaveDate)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
22        Err.Clear
          
End Function

Private Sub txtNO_GotFocus()
    txtNO.SelStart = 0
    txtNO.SelLength = Len(txtNO)
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtName.SetFocus
    End If
End Sub

Private Sub GetMaxNO()
          '功能：         初始化数据
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
              
1         On Error GoTo GetMaxNO_Error

2         strSQL = "select nvl(max(编码),0) 编码 from 检验申请单 "
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "检验申请单")
4         If rsTmp("编码") = 0 Then
5             Me.txtNO = "001"
6         Else
7             Me.txtNO = Format(Val(rsTmp("编码")) + 1, "000")
8         End If
          


9         Exit Sub
GetMaxNO_Error:
10        Call writeErrLog("zl9LisInsideComm", "frmAppforBillDesignEditBill", "执行(GetMaxNO)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
11        Err.Clear
          
End Sub

Public Sub showMe(objfrm As Object, lngID As Long, lngDeptID As Long, strNO As String, strName As String, lngColour As Long, ByVal blnTrs As Boolean)
    '功能           打开主窗体
    
    mlngkeyID = lngID
    mstrNO = strNO
    mstrName = strName
    mlngDeptID = lngDeptID
    mlngColour = lngColour
    mblnTrs = blnTrs
    Me.Show vbModal, objfrm
End Sub

Private Function GetSelColour(lngColour As Long) As Long
    '功能   打开颜色选择器选择颜色
    With ComDialPublic
        .Color = lngColour
        .ShowColor
        GetSelColour = .Color
    End With
    
End Function
