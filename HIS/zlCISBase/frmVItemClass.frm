VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVItemClass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "诊治所见分类"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   Icon            =   "frmVItemClass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   2310
      Left            =   -3315
      TabIndex        =   16
      Tag             =   "1000"
      Top             =   2415
      Visible         =   0   'False
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   4075
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgList"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   300
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3030
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   1875
      MaxLength       =   6
      TabIndex        =   3
      Tag             =   "编码"
      Text            =   "000000"
      Top             =   1335
      Width           =   570
   End
   Begin VB.TextBox txtSymbol 
      Height          =   300
      Left            =   1740
      MaxLength       =   10
      TabIndex        =   7
      Tag             =   "名称"
      Top             =   2025
      Width           =   1425
   End
   Begin VB.CheckBox chkCodeLen 
      Caption         =   "允许更改编码长度，并按此调整各同级编码(&L)"
      Height          =   285
      Left            =   1020
      TabIndex        =   8
      Top             =   2565
      Width           =   4290
   End
   Begin VB.Frame fraBottom 
      Height          =   75
      Left            =   -15
      TabIndex        =   13
      Top             =   2865
      Width           =   5745
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2760
      TabIndex        =   9
      Top             =   3045
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3990
      TabIndex        =   10
      Top             =   3045
      Width           =   1100
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1740
      MaxLength       =   40
      TabIndex        =   5
      Tag             =   "名称"
      Top             =   1665
      Width           =   3195
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&P"
      Height          =   300
      Left            =   4650
      TabIndex        =   11
      Top             =   750
      Width           =   285
   End
   Begin VB.TextBox txtParent 
      Height          =   300
      Left            =   1740
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "(无)"
      Top             =   750
      Width           =   2895
   End
   Begin VB.TextBox txtUpCode 
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   1740
      MaxLength       =   8
      TabIndex        =   12
      TabStop         =   0   'False
      Tag             =   "编码"
      Text            =   "0000"
      Top             =   1290
      Width           =   975
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   60
      Top             =   1230
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemClass.frx":000C
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVItemClass.frx":05A6
            Key             =   "expend"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblHint 
      AutoSize        =   -1  'True
      Caption         =   "(提示：按Del清除上级，设置初级分类)"
      Height          =   180
      Left            =   1740
      TabIndex        =   18
      Top             =   1065
      Width           =   3150
   End
   Begin VB.Label lblSymbol 
      AutoSize        =   -1  'True
      Caption         =   "简码(&S)"
      Height          =   180
      Left            =   1020
      TabIndex        =   6
      Top             =   2085
      Width           =   630
   End
   Begin VB.Label lblKind 
      AutoSize        =   -1  'True
      Caption         =   "病史体征"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   90
      TabIndex        =   15
      Top             =   555
      Width           =   720
   End
   Begin VB.Label lblNote 
      Caption         =   "诊治所见项目可根据其检查记录方式或数据类型等各科特点进行统一分类设置。"
      Height          =   345
      Left            =   990
      TabIndex        =   14
      Top             =   210
      Width           =   4170
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   210
      Picture         =   "frmVItemClass.frx":0B40
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblCode 
      AutoSize        =   -1  'True
      Caption         =   "编码(&D)"
      Height          =   180
      Left            =   1020
      TabIndex        =   2
      Top             =   1350
      Width           =   630
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "名称(&N)"
      Height          =   180
      Left            =   1020
      TabIndex        =   4
      Top             =   1710
      Width           =   630
   End
   Begin VB.Label lblParent 
      AutoSize        =   -1  'True
      Caption         =   "上级(&U)"
      Height          =   180
      Left            =   1020
      TabIndex        =   0
      Top             =   810
      Width           =   630
   End
End
Attribute VB_Name = "frmVItemClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTemp As New ADODB.Recordset
Dim intMaxLen As Integer
Dim objNode As Node

Private Sub chkCodeLen_Click()
    If Me.chkCodeLen.Value = 1 Then
        Me.txtCode.MaxLength = intMaxLen - Len(Me.txtUpCode.Text)
    Else
        Me.txtCode.MaxLength = Me.txtCode.Tag
        Me.txtCode.Text = Mid(Me.txtCode.Text, 1, Me.txtCode.MaxLength)
    End If
    Me.txtCode.SetFocus
End Sub

Private Sub chkCodeLen_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim lngItemID As Long
    
    If Me.txtCode.MaxLength = 0 Then
        MsgBox "上级编码已经达到最大长度，不能设置下级！", vbExclamation, gstrSysName
        Me.cmdCancel.SetFocus
        Exit Sub
    End If
    If Trim(Me.txtCode.Text) = "" Then
        MsgBox "编码必须输入", vbExclamation, gstrSysName
        Me.txtCode.SetFocus
        Exit Sub
    End If
    If Me.chkCodeLen.Value = 0 And Len(Trim(Me.txtCode.Text)) <> Me.txtCode.MaxLength Then
        MsgBox "编码长度必须为" & Me.txtCode.MaxLength & "位，除非你选择更改长度选项", vbExclamation, gstrSysName
        Me.txtCode.SetFocus
        Exit Sub
    End If
    If Trim(Me.txtName.Text) = "" Then
        MsgBox "名称必须输入", vbExclamation, gstrSysName
        Me.txtName.SetFocus
        Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txtName.Text), vbFromUnicode)) > Me.txtName.MaxLength Then
        MsgBox "名称超过" & Me.txtName.MaxLength & "的长度限制", vbExclamation, gstrSysName
        Me.txtName.SetFocus
        Exit Sub
    End If
    
    Err = 0: On Error GoTo ErrHand
    If Me.Tag = "增加" Then
        lngItemID = zlDatabase.GetNextId("诊治所见分类")
        gstrSql = "zl_所见分类_Insert(" & _
            lngItemID & "," & _
            Me.txtParent.Tag & "," & _
            "'" & Me.txtUpCode.Text & Trim(Me.txtCode.Text) & "'," & _
            "'" & Trim(Me.txtName.Text) & "'," & _
            "'" & Trim(Me.txtSymbol.Text) & "'," & _
            Me.lblKind.Tag & "," & _
            Me.chkCodeLen.Value & ")"
    Else
        lngItemID = Me.Tag
        gstrSql = "zl_所见分类_Update(" & _
            lngItemID & "," & _
            Me.txtParent.Tag & "," & _
            "'" & Me.txtUpCode.Text & Trim(Me.txtCode.Text) & "'," & _
            "'" & Trim(Me.txtName.Text) & "'," & _
            "'" & Trim(Me.txtSymbol.Text) & "'," & _
            Me.chkCodeLen.Value & ")"
    End If
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    Unload Me
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSelect_Click()
    With Me.tvwClass
        .Left = Me.txtParent.Left
        .Top = Me.txtParent.Top + Me.txtParent.Height
        .Visible = True
        .SetFocus
    End With
End Sub

Private Sub Form_Activate()
    Err = 0: On Error GoTo ErrHand
    
    If Me.Tag = "增加" Then
        gstrSql = "select ID,上级ID,编码,名称,简码" & _
                " From 诊治所见分类" & _
                " Where 性质 = [1] " & _
                " start with 上级ID is null" & _
                " connect by prior ID=上级ID"
    Else
        gstrSql = "select ID,上级ID,编码,名称,简码" & _
                " From 诊治所见分类" & _
                " Where 性质 = [1] " & _
                "       and id not in (select id from 诊治所见分类 start with 上级id = [2] " & _
                "           connect by prior id = 上级id )" & _
                "       and id <> [2] " & _
                " start with 上级ID is null" & _
                " connect by prior ID=上级ID"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.lblKind.Tag, Me.Tag)
    
    With rsTemp
        intMaxLen = .Fields("编码").DefinedSize
        Me.txtName.MaxLength = .Fields("名称").DefinedSize
        Me.txtSymbol.MaxLength = .Fields("简码").DefinedSize
        
        Me.tvwClass.Visible = False
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!上级ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, "[" & !编码 & "]" & !名称, "close")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !上级ID, tvwChild, "_" & !ID, "[" & !编码 & "]" & !名称, "close")
            End If
            objNode.Sorted = True
            objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
    End With
    If Me.Tag = "增加" Then
        Call zlDefaultCode
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lblCode_Click()
    Me.txtCode.SetFocus
End Sub

Private Sub lblName_Click()
    Me.txtName.SetFocus
End Sub

Private Sub lblParent_Click()
    Me.txtParent.SetFocus
End Sub

Private Sub lblSymbol_Click()
    Me.txtSymbol.SetFocus
End Sub

Private Sub tvwClass_DblClick()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Me.txtParent.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
    Me.txtParent.SetFocus
    Call zlDefaultCode
End Sub

Private Sub tvwClass_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
        If Me.tvwClass.SelectedItem.Children > 0 Then Exit Sub
        Call tvwClass_DblClick
    Case vbKeySpace
        Call tvwClass_DblClick
    Case vbKeyEscape
        Call tvwClass_LostFocus
    End Select
End Sub

Private Sub tvwClass_LostFocus()
    If cmdSelect Is ActiveControl Then Exit Sub
    Me.tvwClass.Visible = False
End Sub

Private Sub txtcode_GotFocus()
    Me.txtCode.SelStart = 0: Me.txtCode.SelLength = 100
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtName_GotFocus()
    Call zlCommFun.OpenIme(True)
    Me.txtName.SelStart = 0: Me.txtName.SelLength = 100
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub txtName_KeyUp(KeyCode As Integer, Shift As Integer)
    Me.txtSymbol.Text = zlStr.GetCodeByORCL(Me.txtName.Text, True, 10)
End Sub

Private Sub txtName_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtParent_GotFocus()
    Me.txtParent.SelStart = 0: Me.txtParent.SelLength = 100
End Sub

Private Sub txtParent_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Me.txtParent.Tag = 0
        Call zlDefaultCode
    End If
    Me.txtParent.SelStart = 0: Me.txtParent.SelLength = 100
End Sub

Private Sub txtParent_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtSymbol_GotFocus()
    Me.txtSymbol.SelStart = 0: Me.txtSymbol.SelLength = 100
End Sub

Private Sub txtSymbol_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtUpCode_Change()
    Me.txtCode.Width = txtUpCode.Width - TextWidth(txtUpCode.Text) - 120
    Me.txtCode.Left = txtUpCode.Left + TextWidth(txtUpCode.Text) + 60
End Sub

Private Sub zlDefaultCode()
    '-----------------------------------------------------
    '功能：根据选择的上级ID(存放于Me.txtParent.Tag)，调整设置编码的缺省值
    '-----------------------------------------------------
    Err = 0: On Error GoTo ErrHand
    
    Me.chkCodeLen.Value = 0
    Me.chkCodeLen.Enabled = True
   
    If Me.txtParent.Tag = 0 Then
        Me.txtParent.Text = "(无)"
        Me.txtUpCode.Text = ""
        gstrSql = "select max(编码) as 编码 From 诊治所见分类 Where 性质=[1] and 上级ID is null "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.lblKind.Tag)
            
        With rsTemp
            If IIf(IsNull(!编码), "", !编码) = "" Then
                Me.txtCode.Text = "01"
                Me.txtCode.MaxLength = intMaxLen
                Me.txtCode.Tag = Me.txtCode.MaxLength
                Me.chkCodeLen.Value = 1
                Me.chkCodeLen.Enabled = False
            Else
                Me.txtCode.MaxLength = Len(Trim(!编码))
                Me.txtCode.Tag = Me.txtCode.MaxLength
                If !编码 = String(Me.txtCode.MaxLength, "9") Then
                    If Me.txtCode.MaxLength >= intMaxLen Then
                        MsgBox "最大编码和编码长度已经达到最大限制，无法递增编码", vbExclamation, gstrSysName
                        Me.txtCode.Text = Space(Me.txtCode.MaxLength)
                        Me.chkCodeLen.Value = 0
                        Me.chkCodeLen.Enabled = False
                    Else
                        MsgBox "最大编码已经达到本级限制，你可以扩充编码长度以满足需要", vbExclamation, gstrSysName
                        Me.txtCode.Text = "1" & String(Me.txtCode.MaxLength, "0")
                        Me.txtCode.MaxLength = Me.txtCode.MaxLength + 1
                        Me.txtCode.Tag = Me.txtCode.MaxLength
                        Me.chkCodeLen.Value = 1
                    End If
                Else
                    Me.txtCode.Text = Format(Mid(!编码, Len(Me.txtUpCode.Text) + 1) + 1, String(Me.txtCode.MaxLength, "0"))
                End If
            End If
        End With
    Else
        With Me.tvwClass
            .Nodes("_" & Me.txtParent.Tag).Selected = True
            Me.txtParent.Text = .SelectedItem.Text
            Me.txtUpCode.Text = Mid(Split(.SelectedItem.Text, "]")(0), 2)
            If .SelectedItem.Children = 0 Then
                Me.txtCode.MaxLength = intMaxLen - Len(Me.txtUpCode.Text)
                Me.txtCode.Tag = Me.txtCode.MaxLength
                If Me.txtCode.MaxLength > 1 Then
                    Me.txtCode.Text = "01"
                Else
                    Me.txtCode.Text = "1"
                End If
                Me.chkCodeLen.Value = 1
                Me.chkCodeLen.Enabled = False
            Else
                gstrSql = "select nvl(max(编码),'') as 编码  From 诊治所见分类 Where 性质=[1] and 上级ID=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.lblKind.Tag, Mid(.SelectedItem.Key, 2))
                
                With rsTemp
                    Me.txtCode.MaxLength = Len(!编码) - Len(Me.txtUpCode.Text)
                    Me.txtCode.Tag = Me.txtCode.MaxLength
                    If Mid(!编码, Len(Me.txtUpCode.Text) + 1) = String(Me.txtCode.MaxLength, "9") Then
                        If Len(Me.txtUpCode.Text) + Me.txtCode.MaxLength >= intMaxLen Then
                            MsgBox "该分类下级最大编码和编码长度已经达到最大限制，无法递增编码", vbExclamation, gstrSysName
                            Me.txtCode.Text = Space(Me.txtCode.MaxLength)
                            Me.chkCodeLen.Value = 0
                            Me.chkCodeLen.Enabled = False
                        Else
                            MsgBox "该分类下级最大编码已经达到本级限制，你可以扩充编码长度以满足需要", vbExclamation, gstrSysName
                            Me.txtCode.Text = "1" & String(Me.txtCode.MaxLength, "0")
                            Me.txtCode.MaxLength = Me.txtCode.MaxLength + 1
                            Me.txtCode.Tag = Me.txtCode.MaxLength
                            Me.chkCodeLen.Value = 1
                        End If
                    Else
                        Me.txtCode.Text = Format(Mid(!编码, Len(Me.txtUpCode.Text) + 1) + 1, String(Me.txtCode.MaxLength, "0"))
                    End If
                End With
            End If
        End With
    End If
    Me.txtParent.SetFocus
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
