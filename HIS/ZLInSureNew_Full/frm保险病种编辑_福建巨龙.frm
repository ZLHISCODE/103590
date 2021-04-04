VERSION 5.00
Begin VB.Form frm保险病种编辑_福建巨龙 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "保险病种编辑"
   ClientHeight    =   5295
   ClientLeft      =   2760
   ClientTop       =   3645
   ClientWidth     =   4620
   Icon            =   "frm保险病种编辑_福建巨龙.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra项目 
      Caption         =   "特准使用项目"
      Height          =   2325
      Left            =   120
      TabIndex        =   10
      Top             =   2340
      Width           =   4365
      Begin VB.CommandButton cmdClear 
         Caption         =   "清除(&L)"
         CausesValidation=   0   'False
         Height          =   350
         Left            =   3150
         TabIndex        =   14
         Top             =   1770
         Width           =   1100
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "删除(&D)"
         CausesValidation=   0   'False
         Height          =   350
         Left            =   3150
         TabIndex        =   13
         Top             =   1320
         Width           =   1100
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "增加(&A)"
         CausesValidation=   0   'False
         Height          =   350
         Left            =   3150
         TabIndex        =   12
         Top             =   270
         Width           =   1100
      End
      Begin VB.ListBox lst项目 
         Height          =   1860
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   270
         Width           =   2955
      End
   End
   Begin VB.Frame fra基本 
      Caption         =   "基本"
      Height          =   2085
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   4365
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   825
         MaxLength       =   20
         TabIndex        =   2
         Top             =   390
         Width           =   1995
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   825
         MaxLength       =   50
         TabIndex        =   4
         Top             =   780
         Width           =   3375
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   825
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1185
         Width           =   1095
      End
      Begin VB.OptionButton opt类别 
         Caption         =   "慢性病(&M)"
         Height          =   180
         Index           =   1
         Left            =   1440
         TabIndex        =   8
         Top             =   1740
         Width           =   1155
      End
      Begin VB.OptionButton opt类别 
         Caption         =   "普通病(&G)"
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   7
         Top             =   1740
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton opt类别 
         Caption         =   "特种病(&S)"
         Height          =   180
         Index           =   2
         Left            =   2745
         TabIndex        =   9
         Top             =   1740
         Width           =   1155
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "名称(&N)"
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   3
         Top             =   840
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "简码(&S)"
         Height          =   180
         Index           =   2
         Left            =   150
         TabIndex        =   5
         Top             =   1245
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "编码(&U)"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   450
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   195
      TabIndex        =   17
      Top             =   4800
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2160
      TabIndex        =   15
      Top             =   4800
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3360
      TabIndex        =   16
      Top             =   4800
      Width           =   1100
   End
End
Attribute VB_Name = "frm保险病种编辑_福建巨龙"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enum编辑
    text编码 = 0
    Text名称 = 1
    Text简码 = 2
End Enum

Dim mlng险类 As Long
Dim mstrID As String         '当前编辑的医保大类ID
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '是否改变了

Private Sub cmdADD_Click()
    Dim strID As String, str编码 As String, str名称 As String
    Dim blnReturn As Boolean
    Dim lngIndex As Long
    
    If frm收费细目选择.ShowTree(strID, str编码, str名称) = True Then
        For lngIndex = 0 To lst项目.ListCount - 1
            '已经加入
            If lst项目.ItemData(lngIndex) = Val(strID) Then Exit Sub
        Next
        
        lst项目.AddItem "【" & str编码 & "】" & str名称
        lst项目.ItemData(lst项目.NewIndex) = Val(strID)
        
        mblnChange = True
    End If
End Sub

Private Sub CmdClear_Click()
    lst项目.Clear
    mblnChange = True
End Sub

Private Sub cmdDelete_Click()
    If lst项目.ListIndex < 0 Then Exit Sub
    
    lst项目.RemoveItem lst项目.ListIndex
    mblnChange = True
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdOK_Click()
    Dim lngIndex As Long
    
    If IsValid() = False Then Exit Sub
    If Save项目() = False Then Exit Sub
    
    If mstrID = "" Then
        '连续新增
        txtEdit(text编码).Text = GetMaxCode 'zlDatabase.GetMax("保险病种", "编码", 6, " where 险类=" & mlng险类)
        For lngIndex = Text名称 To Text简码
            txtEdit(lngIndex).Text = ""
        Next
        lst项目.Clear
        
        mblnChange = False
        txtEdit(text编码).SetFocus
    Else
        mblnChange = False
        Unload Me
    End If
End Sub

Private Function GetMaxCode() As String
'功能：读取指定表的本级编码的最大值
'返回：成功返回 下级最大编码; 否者返回 0
    Dim rsTemp As New ADODB.Recordset
    Dim varTemp As Variant
    Dim lngLengh As Long
    
    On Error GoTo ErrHand
    With rsTemp
        gstrSQL = "SELECT max(length(substr(名称,1,instr(名称,'@@')-1))) as 最长值 FROM 保险病种 where 险类=" & mlng险类
        Call OpenRecordset(rsTemp, Me.Caption)
        If rsTemp.EOF = True Then
            GetMaxCode = "1"
            Exit Function
        Else
            lngLengh = Nvl(rsTemp("最长值"), "1")
        End If
        
        gstrSQL = "SELECT MAX(LPAD(substr(名称,1,instr(名称,'@@')-1)," & lngLengh & ",' ')) as 最大值 FROM 保险病种 where 险类=" & mlng险类
        Call OpenRecordset(rsTemp, Me.Caption)
        If rsTemp.EOF Then
            GetMaxCode = Format(1, String(lngLengh, "0"))
            Exit Function
        End If
        
        varTemp = Nvl(rsTemp("最大值"), "0")
        If IsNumeric(varTemp) Then
            GetMaxCode = CStr(Val(varTemp) + 1)
            GetMaxCode = Format(GetMaxCode, String(lngLengh, "0"))
        Else
            GetMaxCode = Mid(varTemp, 1, Len(varTemp) - 1) & Chr(asc(Right(varTemp, 1)) + 1)
            GetMaxCode = Trim(GetMaxCode)
        End If
        .Close
    End With
    Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
End Function

Private Function Save项目() As Boolean
    Dim lngID As Long, lng类别 As Long
    Dim lngIndex As Long, lst As ListItem
    Dim str特准项目 As String
    Dim strCode As String
    Dim rsTmp As New ADODB.Recordset
    
    For lngIndex = opt类别.LBound To opt类别.UBound
        If opt类别(lngIndex).Value = True Then
            lng类别 = lngIndex
            Exit For
        End If
    Next
    
    For lngIndex = 0 To lst项目.ListCount - 1
        str特准项目 = str特准项目 & lst项目.ItemData(lngIndex) & ":"
    Next
    
    On Error GoTo errHandle
    
    If mstrID = "" Then
        '新增
        If CheckCode(txtEdit(text编码)) = False Then Exit Function
        lngID = zlDatabase.GetNextId("保险病种")
        '获取保险编码
        strCode = zlDatabase.GetMax("保险病种", "编码", 6, " Where 险类=" & mlng险类)
        gstrSQL = "zl_保险病种_INSERT(" & lngID & "," & mlng险类 & ",'" & strCode & "','" & _
                Trim(txtEdit(text编码).Text) & "@@" & Trim(txtEdit(Text名称).Text) & "','" & Trim(txtEdit(Text简码).Text) & "'," & lng类别 & ",null,null,'" & str特准项目 & "')"
    Else
        '修改
        If CheckCode(txtEdit(text编码), False) = False Then Exit Function
        '获取保险编码
        gstrSQL = "Select 编码 From 保险病种 Where 险类=" & mlng险类 & " And ID=" & mstrID
        Call OpenRecordset(rsTmp, "获取当前保险病种的编码")
        strCode = rsTmp!编码
        
        gstrSQL = "zl_保险病种_Update(" & mstrID & ",'" & strCode & "','" & _
                Trim(txtEdit(text编码).Text) & "@@" & Trim(txtEdit(Text名称).Text) & "','" & Trim(txtEdit(Text简码).Text) & "'," & lng类别 & ",null,null,'" & str特准项目 & "')"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '更新主界面
    If mstrID = "" Then
        Set lst = frm保险病种.lvwItem.ListItems.Add(, "K" & lngID, txtEdit(text编码), "Disease", "Disease")
    Else
        Set lst = frm保险病种.lvwItem.SelectedItem
    End If
    lst.SubItems(1) = Trim(txtEdit(Text名称).Text)
    lst.SubItems(2) = Trim(txtEdit(Text简码).Text)
    lst.SubItems(3) = IIf(lng类别 = 0, "普通病", IIf(lng类别 = 1, "慢性病", "特种病"))
    
    Save项目 = True
    mblnOK = True
    Exit Function

errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckCode(ByVal strCode As String, Optional ByVal blnNew As Boolean = True) As Boolean
    Dim rsCode As New ADODB.Recordset
    '因为编码超长，只有将编码与名称保存在名称列，而编码列实际保存的是记录数，在用户修改编码时，需要判断编码是否重复
    
    CheckCode = False
    gstrSQL = "Select 1 From 保险病种 Where substr(名称,1,instr(名称,'@@')-1)='" & strCode & "'" & IIf(blnNew, "", " And ID<>" & mstrID)
    Call OpenRecordset(rsCode, "判断编码是否重复")
    
    If Not rsCode.EOF Then
        MsgBox "保险病种编码重复！", vbInformation, gstrSysName
        txtEdit(text编码).SetFocus
        Exit Function
    End If
    CheckCode = True
End Function

Private Function IsValid() As Boolean
'功能:分析输入有关医保类别的内容是否有效
'参数:
'返回值:有效返回True,否则为False
    Dim lngIndex As Integer
    For lngIndex = text编码 To Text简码
        If zlCommFun.StrIsValid(Trim(txtEdit(lngIndex).Text), txtEdit(lngIndex).MaxLength) = False Then
            txtEdit(lngIndex).SetFocus
            zlControl.TxtSelAll txtEdit(lngIndex)
            Exit Function
        End If
        
        If lngIndex = text编码 Or lngIndex = Text名称 Then
            If Len(Trim(txtEdit(lngIndex).Text)) = 0 Then
                txtEdit(lngIndex).Text = ""
                MsgBox "编码或名称都不能为空。", vbExclamation, gstrSysName
                txtEdit(lngIndex).SetFocus
                Exit Function
            End If
        End If
    Next
    
    If lst项目.ListCount > 50 Then
        MsgBox "该病种的特准医保项目太多，不能超过50个。", vbInformation, gstrSysName
        Exit Function
    End If
    IsValid = True
End Function

Private Sub opt类别_Click(Index As Integer)
    mblnChange = True
End Sub

Private Sub opt类别_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = Text名称 Then
        txtEdit(Text简码).Text = zlCommFun.SpellCode(txtEdit(Text名称).Text)
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    Select Case Index
        Case Text名称
          zlCommFun.OpenIme True
        Case Else
            zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0 '使之不响
        zlCommFun.PressKey (vbKeyTab)
    Else
        If Index = text编码 Then
            KeyAscii = asc(UCase(Chr(KeyAscii)))
            If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
End Sub

Public Function 编辑病种(ByVal lng险类 As Long, ByVal strID As String) As Boolean
'功能:用来与调用的医保类别管理窗口进行通讯的程序
'参数:str序号           当前编辑的医保类别的的序号
'返回值:编辑成功返回True,否则为False
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    
    mblnOK = False
    mlng险类 = lng险类
    mstrID = strID
    
    rsTemp.CursorLocation = adUseClient
    If mstrID <> "" Then
        '修改医保大类
        gstrSQL = "select substr(名称,1,instr(名称,'@@')-1) 编码,substr(名称,instr(名称,'@@')+2) 名称,简码,nvl(类别,'0') as 类别 from 保险病种 where ID=" & mstrID
        Call OpenRecordset(rsTemp, Me.Caption)
        
        txtEdit(text编码).Text = rsTemp("编码")
        txtEdit(Text名称).Text = rsTemp("名称")
        txtEdit(Text简码).Text = IIf(IsNull(rsTemp("简码")), "", rsTemp("简码"))
        opt类别(rsTemp("类别")).Value = True
        
        '修改医保大类
        gstrSQL = "select A.ID,A.编码,A.名称 from 收费细目 A,保险特准项目 B where A.ID=B.收费细目ID and B.病种ID=" & mstrID
        Call OpenRecordset(rsTemp, Me.Caption)
        
        Do Until rsTemp.EOF
            lst项目.AddItem "【" & rsTemp("编码") & "】" & rsTemp("名称")
            lst项目.ItemData(lst项目.NewIndex) = rsTemp("ID")
            rsTemp.MoveNext
        Loop
    Else
        '新增医保大类
        txtEdit(text编码).Text = GetMaxCode 'zlDatabase.GetMax("保险病种", "编码", 6, " where 险类=" & mlng险类)
    End If
    
    
    mblnChange = False
    frm保险病种编辑_福建巨龙.Show vbModal, frm保险病种
    编辑病种 = mblnOK
End Function

