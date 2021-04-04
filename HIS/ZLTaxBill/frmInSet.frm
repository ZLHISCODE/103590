VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInSet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "住院税控设置"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   Icon            =   "frmInSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkUse 
      Caption         =   "使用外挂式税控器开票"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   5250
      Width           =   2100
   End
   Begin VB.CheckBox chk密码器 
      Alignment       =   1  'Right Justify
      Caption         =   "使用密码器"
      Height          =   195
      Left            =   3255
      TabIndex        =   2
      Top             =   173
      Width           =   1200
   End
   Begin VB.TextBox txtTaxName 
      Height          =   300
      Left            =   3360
      MaxLength       =   12
      TabIndex        =   5
      Top             =   4860
      Width           =   1335
   End
   Begin VB.TextBox txtTaxNo 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3360
      MaxLength       =   3
      TabIndex        =   4
      Top             =   4260
      Width           =   1335
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "缺省(&D)"
      Height          =   350
      Left            =   3465
      TabIndex        =   8
      Top             =   1635
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&S)"
      Height          =   350
      Left            =   3465
      TabIndex        =   6
      Top             =   825
      Width           =   1100
   End
   Begin VB.TextBox txt前缀 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1485
      MaxLength       =   13
      TabIndex        =   1
      Top             =   120
      Width           =   1530
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   30
      TabIndex        =   12
      Top             =   495
      Width           =   4890
   End
   Begin MSComctlLib.ListView lvwItem 
      Height          =   4350
      Left            =   120
      TabIndex        =   3
      Top             =   810
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   7673
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img16"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   3465
      Picture         =   "frmInSet.frx":058A
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1980
      Width           =   1100
   End
   Begin VB.CommandButton cmdCanc 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3465
      TabIndex        =   7
      Top             =   1170
      Width           =   1100
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1890
      Top             =   2730
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInSet.frx":06D4
            Key             =   "Item"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTaxName 
      AutoSize        =   -1  'True
      Caption         =   "税票项目名称"
      Height          =   180
      Left            =   3135
      TabIndex        =   16
      Top             =   4650
      Width           =   1080
   End
   Begin VB.Label lblTaxNo 
      AutoSize        =   -1  'True
      Caption         =   "税票项目序号"
      Height          =   180
      Left            =   3135
      TabIndex        =   15
      Top             =   4050
      Width           =   1080
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      Caption         =   "对照税票实际允许项目和相关规定，设置当前税票项目，以便系统能正确传递。"
      Height          =   720
      Left            =   3150
      TabIndex        =   14
      Top             =   2970
      Width           =   1620
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblNotice 
      AutoSize        =   -1  'True
      Caption         =   "注意："
      Height          =   180
      Left            =   3135
      TabIndex        =   13
      Top             =   2715
      Width           =   540
   End
   Begin VB.Label lbl前缀 
      AutoSize        =   -1  'True
      Caption         =   "住院收据号前缀"
      Height          =   180
      Left            =   150
      TabIndex        =   10
      Top             =   180
      Width           =   1260
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      Caption         =   "收据费目对应税票许可项目序号:"
      Height          =   180
      Left            =   150
      TabIndex        =   11
      Top             =   585
      Width           =   2610
   End
End
Attribute VB_Name = "frmInSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
Dim strItems As String
Dim aryItem() As String
Dim intCount As Integer
Dim blnUpdate As Boolean

Private Sub cmdCanc_Click()
    Unload Me
End Sub

Private Sub cmdDefault_Click()
    For Each objItem In Me.lvwItem.ListItems
        objItem.SubItems(1) = Format(Mid(objItem.Key, 2), "000") & "-" & Mid(objItem.Text, InStr(1, objItem, "-") + 1)
    Next
    If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
    Call lvwItem_ItemClick(Me.lvwItem.SelectedItem)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, 100
End Sub

Private Sub cmdOK_Click()
    Dim strSave As String, strEmpty As String
    strSave = "": strEmpty = ""
    For Each objItem In Me.lvwItem.ListItems
        If objItem.SubItems(1) = "" Then
            strEmpty = strEmpty & vbCrLf & Space(8) & objItem.Text
        Else
            strSave = strSave & "|" & Mid(objItem.Text, InStr(1, objItem.Text, "-") + 1) & ";" & objItem.SubItems(1)
        End If
    Next
    If strEmpty <> "" Then
        strEmpty = "以下收据费目未设置对应的住院税票项目：" & strEmpty
        strEmpty = strEmpty & vbCrLf & "如果确认这些项目不会在住院发生，可以继续？"
        If MsgBox(strEmpty, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    
    If strSave <> "" Then strSave = Mid(strSave, 2)
    Call SaveSetting("ZLSOFT", "公共全局\税票打印", "住院税票项目", strSave)
    Call SaveSetting("ZLSOFT", "公共全局\税票打印", "住院税票前缀", Trim(Me.txt前缀.Text))
    
    Call SaveSetting("ZLSOFT", "公共全局\税票打印", "住院使用密码器", Me.chk密码器.Value)
    Call SaveSetting("ZLSOFT", "公共全局\税票打印", "住院使用税票打印", Me.chkUse.Value)
    
    Unload Me
End Sub

Private Sub Form_Activate()
    If Me.lvwItem.ListItems.Count = 0 Then
        MsgBox "没有设置收据费目，无法设置对应的住院税票对应项目！", vbExclamation, gstrSysName
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    blnUpdate = True
    Me.chkUse.Value = Val(GetSetting("ZLSOFT", "公共全局\税票打印", "住院使用税票打印", 0))
    Me.chk密码器.Value = Val(GetSetting("ZLSOFT", "公共全局\税票打印", "住院使用密码器", 0))
    Me.txt前缀.Text = GetSetting("ZLSOFT", "公共全局\税票打印", "住院税票前缀", "2030030301001")
    strItems = GetSetting("ZLSOFT", "公共全局\税票打印", "住院税票项目", "")
    aryItem = Split(strItems, "|")
    
    Me.lvwItem.ColumnHeaders.Clear
    With Me.lvwItem.ColumnHeaders
        .Clear
        .Add , "收据费目", "收据费目", 1400
        .Add , "税控项目", "税控项目", 1400
    End With
    
    gstrSql = "select * from 收据费目"
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly
        Call SQLTest
        
        Me.lvwItem.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItem.ListItems.Add(, "_" & !编码, !编码 & "-" & !名称, "Item", "Item")
            For intCount = LBound(aryItem) To UBound(aryItem)
                If Split(aryItem(intCount), ";")(0) = !名称 Then
                    objItem.SubItems(1) = Split(aryItem(intCount), ";")(1): Exit For
                End If
            Next
            .MoveNext
        Loop
    End With
    If Me.lvwItem.ListItems.Count > 0 Then
        Me.lvwItem.ListItems(1).Selected = True
        Call lvwItem_ItemClick(Me.lvwItem.SelectedItem)
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lvwItem_ItemClick(ByVal Item As MSComctlLib.ListItem)
    blnUpdate = False
    If InStr(1, Item.SubItems(1), "-") > 0 Then
        Me.txtTaxNo.Text = Mid(Item.SubItems(1), 1, InStr(1, Item.SubItems(1), "-") - 1)
        Me.txtTaxName.Text = Mid(Item.SubItems(1), InStr(1, Item.SubItems(1), "-") + 1)
    Else
        Me.txtTaxNo.Text = "": Me.txtTaxName.Text = ""
    End If
    blnUpdate = True
End Sub

Private Sub txtTaxName_Change()
    If blnUpdate = False Then Exit Sub
    If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
    If Trim(Me.txtTaxNo.Text) <> "" And Trim(Me.txtTaxName.Text) <> "" Then
        Me.lvwItem.SelectedItem.SubItems(1) = Format(Trim(Me.txtTaxNo.Text), "000") & "-" & Trim(Me.txtTaxName.Text)
    Else
        Me.lvwItem.SelectedItem.SubItems(1) = ""
    End If
End Sub

Private Sub txtTaxName_GotFocus()
    Me.txtTaxName.SelStart = 0: Me.txtTaxName.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtTaxName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@$%^&*_+|`;'"":/,?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtTaxName_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtTaxNo_Change()
    If blnUpdate = False Then Exit Sub
    If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
    If Trim(Me.txtTaxNo.Text) <> "" And Trim(Me.txtTaxName.Text) <> "" Then
        Me.lvwItem.SelectedItem.SubItems(1) = Format(Trim(Me.txtTaxNo.Text), "000") & "-" & Trim(Me.txtTaxName.Text)
    Else
        Me.lvwItem.SelectedItem.SubItems(1) = ""
    End If
End Sub

Private Sub txtTaxNo_GotFocus()
    Me.txtTaxNo.SelStart = 0: Me.txtTaxNo.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtTaxNo_KeyPress(KeyAscii As Integer)
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

Private Sub txt前缀_GotFocus()
    Me.txt前缀.SelStart = 0: Me.txt前缀.SelLength = 100
End Sub
