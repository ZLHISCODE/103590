VERSION 5.00
Begin VB.Form frmPreCompendEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病历预制提纲编辑"
   ClientHeight    =   5325
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5055
   Icon            =   "frmOutlineEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -135
      TabIndex        =   14
      Top             =   4725
      Width           =   5760
   End
   Begin VB.ListBox lstApply 
      Height          =   1740
      ItemData        =   "frmOutlineEdit.frx":058A
      Left            =   1710
      List            =   "frmOutlineEdit.frx":058C
      Style           =   1  'Checkbox
      TabIndex        =   9
      Top             =   2310
      Width           =   3075
   End
   Begin VB.CheckBox chkCopy 
      Caption         =   "复用提纲(&R)"
      Height          =   210
      Left            =   675
      TabIndex        =   10
      Top             =   4155
      Width           =   1305
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2550
      TabIndex        =   11
      Top             =   4860
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3700
      TabIndex        =   12
      Top             =   4860
      Width           =   1100
   End
   Begin VB.TextBox txtExplain 
      Height          =   660
      Left            =   1380
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1515
      Width           =   3420
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1380
      TabIndex        =   2
      Top             =   1125
      Width           =   3405
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   1380
      TabIndex        =   1
      Top             =   750
      Width           =   795
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   660
      TabIndex        =   0
      Top             =   600
      Width           =   5310
   End
   Begin VB.Label lblCopy 
      AutoSize        =   -1  'True
      Caption         =   "该提纲的内容是否可在后续病历书写时直接引入。"
      Height          =   180
      Left            =   945
      TabIndex        =   13
      Top             =   4410
      Width           =   3960
   End
   Begin VB.Label lblApply 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "应用范围(&A)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   675
      TabIndex        =   8
      Top             =   2310
      Width           =   990
   End
   Begin VB.Label lblExplain 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "说明(&M)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   675
      TabIndex        =   7
      Top             =   1575
      Width           =   630
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   120
      Picture         =   "frmOutlineEdit.frx":058E
      Top             =   165
      Width           =   480
   End
   Begin VB.Label lblNote 
      Caption         =   "在病历文件定义前设置管理常用的病历提纲条目，以便在多个病历文件重复应用。"
      Height          =   345
      Left            =   675
      TabIndex        =   5
      Top             =   135
      Width           =   4140
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblName 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "名称(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   675
      TabIndex        =   4
      Top             =   1185
      Width           =   630
   End
   Begin VB.Label lblCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "编号(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   675
      TabIndex        =   3
      Top             =   810
      Width           =   630
   End
End
Attribute VB_Name = "frmPreCompendEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'说明：
'   1、上级程序通过本窗体ShowMe函数，将父窗体、编辑单据ID,编辑状态等信息传递进入本程序
'   2、编辑状态：由Me.tag存放，分别为"新增"、"修改"，由上级程序通过ShowMe传入
'---------------------------------------------------
Private mlngItemID As Long       '被编辑的记录ID，修改、查阅时由上级程序通过ShowMe传递进入,新增时为0，
Private mblnOK As Boolean        '是否完成编辑退出

'临时变量
Dim rsTemp As New ADODB.Recordset
Dim lngCount As Long
Dim strApply As String

Public Function ShowMe(ByVal frmParent As Object, ByVal blnAdd As Boolean, Optional ByVal lngItemId As Long) As Long
    '---------------------------------------------------
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '返回：确定返回新增或修改的ID；取消返回0
    '---------------------------------------------------
    If blnAdd Then
        Me.Tag = "新增"
    Else
        Me.Tag = "修改"
    End If
    mlngItemID = lngItemId
    
    With Me.lstApply
        .Clear
        .AddItem "1-门诊病历"
        .AddItem "2-住院病历"
        .AddItem "3-护理记录"
        .AddItem "4-护理病历"
        .AddItem "5-疾病证明报告"
        .AddItem "6-知情文件"
        .AddItem "7-诊疗申请"
        .AddItem "8-诊疗报告"
        For lngCount = 1 To .ListCount
            .Selected(lngCount - 1) = True
        Next
        .Selected(2) = False
        .ListIndex = 0
    End With
    
    '提取信息
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select 对象序号, 内容文本, 对象属性, Nvl(复用提纲, 0) As 复用提纲, 使用时机 From 病历文件结构 Where 文件id Is Null And ID=[1]"
    Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption, mlngItemID)
    With rsTemp
        If .RecordCount > 0 Then
            Me.txtCode.Text = !对象序号: Me.txtName.Text = "" & !内容文本: Me.txtExplain.Text = "" & !对象属性
            Me.chkCopy.Value = !复用提纲
            strApply = Left("" & !使用时机, 8)
            For lngCount = 1 To Len(strApply)
                Me.lstApply.Selected(lngCount - 1) = IIf(Val(Mid(strApply, lngCount, 1)) = 0, False, True)
            Next
        End If
        Me.txtCode.MaxLength = 3
        Me.txtName.MaxLength = 30
        Me.txtExplain.MaxLength = 200
    End With
    If Me.Tag = "新增" Then
        gstrSQL = "Select nvl(max(对象序号),0) as 对象序号 From 病历文件结构 Where 文件id Is Null"
        Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption)
        Me.txtCode.Text = Val(Format(Val(rsTemp!对象序号) + 1, String(Me.txtCode.MaxLength, "0")))
    End If
    
    '显示窗体
    Me.Show vbModal, frmParent
    If mblnOK Then
        ShowMe = mlngItemID
    Else
        ShowMe = 0
    End If
    Unload Me
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ShowMe = 0
End Function

Private Sub chkCopy_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chkCopy_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False: Me.Hide
End Sub

Private Sub cmdOK_Click()
    If Trim(Me.txtCode.Text) = "" Then MsgBox "请输入编号！", vbInformation, gstrSysName: Me.txtCode.SetFocus: Exit Sub
    If Trim(Me.txtName.Text) = "" Then MsgBox "请输入名称！", vbInformation, gstrSysName: Me.txtName.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txtName.Text), vbFromUnicode)) > Me.txtName.MaxLength Then
        MsgBox "名称超长（最多" & Me.txtName.MaxLength & "个字符或等长的汉字）！", vbInformation, gstrSysName: Me.txtName.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txtExplain.Text), vbFromUnicode)) > Me.txtExplain.MaxLength Then
        MsgBox "说明超长（最多" & Me.txtExplain.MaxLength & "个字符或等长的汉字）！", vbInformation, gstrSysName: Me.txtExplain.SetFocus: Exit Sub
    End If
    
    Dim RS As New ADODB.Recordset, strS As String, lngSum As Long
    '数据保存
    With Me.lstApply
        strApply = ""
        For lngCount = 1 To .ListCount
            strApply = strApply & IIf(.Selected(lngCount - 1) = True, "1", "0")
        Next
    End With
    If Me.Tag = "新增" Then
        gstrSQL = Trim(Me.txtCode.Text) & ",'" & Trim(Me.txtName.Text) & "','" & Trim(Me.txtExplain.Text) & "'," & IIf(Me.chkCopy.Value = 1, 1, 0) & ",'" & strApply & "'"
        mlngItemID = zlDatabase.GetNextId("病历文件结构")
        gstrSQL = "Zl_病历预制提纲_Insert(" & mlngItemID & "," & gstrSQL & ")"
    Else
        gstrSQL = "Select count(A.ID) From 病历文件结构 A, 病历文件列表 B Where A.预制提纲id = [1] And A.文件id = B.ID"
        strS = ""
        With Me.lstApply
            For lngCount = 1 To .ListCount
                If .Selected(lngCount - 1) = False Then
                    If strS = "" Then
                        strS = " B.种类 = " & lngCount & " "
                    Else
                        strS = strS & " or B.种类 = " & lngCount & " "
                    End If
                End If
            Next
            If strS <> "" Then gstrSQL = gstrSQL & " And (" & strS & ")"
        End With
        Set RS = OpenSQLRecord(gstrSQL, Me.Caption, mlngItemID)
        If Not RS.EOF Then
            lngSum = RS(0)
            If lngSum > 0 Then
                If MsgBox("该预制提纲已经在其他种类的病历中使用，是否继续？", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then Exit Sub
            End If
        End If
        RS.Close
        gstrSQL = Trim(Me.txtCode.Text) & ",'" & Trim(Me.txtName.Text) & "','" & Trim(Me.txtExplain.Text) & "'," & IIf(Me.chkCopy.Value = 1, 1, 0) & ",'" & strApply & "'"
        gstrSQL = "Zl_病历预制提纲_Update(" & mlngItemID & "," & gstrSQL & ")"
    End If
    
    Err = 0: On Error GoTo errHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSQL): gcnOracle.Execute gstrSQL, , adCmdStoredProc: Call SQLTest
    mblnOK = True: Me.Hide
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lstApply_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub lstApply_ItemCheck(Item As Integer)
    If Item = 2 Then Me.lstApply.Selected(Item) = False
End Sub

Private Sub lstApply_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtCode_Change()
    txtCode = Val(txtCode)
End Sub

Private Sub txtCode_GotFocus()
    Me.txtCode.SelStart = 0: Me.txtCode.SelLength = 100
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtExplain_Change()
    ValidControlText txtExplain
End Sub

Private Sub txtName_Change()
    ValidControlText txtName
End Sub

Private Sub txtName_GotFocus()
    Me.txtName.SelStart = 0: Me.txtName.SelLength = 4000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtExplain_GotFocus()
    Me.txtExplain.SelStart = 0: Me.txtExplain.SelLength = 4000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtExplain_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
