VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWritImp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人病历引入"
   ClientHeight    =   3150
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6075
   Icon            =   "frmWritImp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4740
      TabIndex        =   3
      Top             =   135
      Width           =   1200
   End
   Begin MSComctlLib.ListView lvwPati 
      Height          =   2565
      Left            =   825
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3090
      Visible         =   0   'False
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   4524
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwWrits 
      Height          =   2355
      Left            =   150
      TabIndex        =   2
      Top             =   720
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   4154
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtPati 
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   840
      MaxLength       =   11
      TabIndex        =   1
      ToolTipText     =   "请按""-病人ID""、""+住院号""、""*门诊号""形式输入或直接输入姓名查找"
      Top             =   120
      Width           =   900
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   4740
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1035
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4740
      TabIndex        =   4
      Top             =   495
      Width           =   1200
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   5100
      Picture         =   "frmWritImp.frx":08CA
      Top             =   2310
      Width           =   480
   End
   Begin VB.Label lblWrit 
      AutoSize        =   -1  'True
      Caption         =   "入院记录："
      Height          =   180
      Left            =   150
      TabIndex        =   8
      Top             =   525
      Width           =   900
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "姓名:        性别:    年龄:  "
      Height          =   180
      Left            =   1860
      TabIndex        =   6
      Top             =   165
      Width           =   2610
   End
   Begin VB.Label lblPati 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "病人(&P)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   150
      TabIndex        =   0
      Top             =   165
      Width           =   630
   End
End
Attribute VB_Name = "frmWritImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lngFileId As Long
Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
Dim intCount As Integer, intRow As Integer, intCol As Integer
Dim strTemp As String, aryTemp() As String

Private Sub cmdCancel_Click()
    lngFileId = 0
    Me.Hide
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    lngFileId = Mid(Me.lvwWrits.SelectedItem.Key, 2)
    Me.Hide
End Sub

Private Sub Form_Activate()
    gstrSql = "select 名称 from 病历文件目录 where ID=" & Me.lblWrit.Tag
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        If .State = adStateOpen Then .Close
        Call SQLTest(App.Title, Me.Caption, gstrSql): .Open gstrSql, gcnOracle: Call SQLTest
        Me.lblWrit.Caption = !名称 & ":"
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then cmdCancel_Click
End Sub

Private Sub Form_Load()
    
    Me.lvwWrits.ListItems.Clear
    With Me.lvwWrits.ColumnHeaders
        .Clear
        .Add , "序号", "序号", 600
        .Add , "书写人", "书写人", 900
        .Add , "书写日期", "书写日期", 1700
    End With
    With Me.lvwWrits
        .SortKey = .ColumnHeaders("序号").Index - 1: .SortOrder = lvwAscending
    End With
    
    With Me.lvwPati.ColumnHeaders
        .Clear
        .Add , "病人ID", "病人ID", 800
        .Add , "门诊号", "门诊号", 800
        .Add , "住院号", "住院号", 800
        .Add , "姓名", "姓名", 900
        .Add , "性别", "性别", 600
        .Add , "年龄", "年龄", 600
    End With
    With Me.lvwPati
        .SortKey = .ColumnHeaders("病人ID").Index - 1: .SortOrder = lvwAscending
    End With

End Sub

Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwPati.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwPati.SortOrder = IIf(Me.lvwPati.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwPati.SortKey = ColumnHeader.Index - 1
        Me.lvwPati.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwPati_DblClick()
    If Me.lvwPati.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwPati
        If Val(Me.txtPati.Tag) <> Val(.SelectedItem.Text) Then
            Me.txtPati.Tag = .SelectedItem.Text
            Me.txtPati.Text = Me.txtPati.Tag
            Me.lblInfo.Caption = "姓名:" & .SelectedItem.SubItems(.ColumnHeaders("姓名").Index - 1) & _
                    Space(2) & "性别:" & .SelectedItem.SubItems(.ColumnHeaders("性别").Index - 1) & _
                    Space(2) & "年龄:" & .SelectedItem.SubItems(.ColumnHeaders("年龄").Index - 1)
            Me.lblInfo.Tag = .SelectedItem.SubItems(.ColumnHeaders("姓名").Index - 1)
        End If
        Me.txtPati.SetFocus
        Call RefereshWrits
        Call zlCommFun.PressKey(vbKeyTab)
    End With
End Sub

Private Sub lvwPati_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwPati.SelectedItem Is Nothing Then Exit Sub
        Call lvwPati_DblClick
    End Select
End Sub

Private Sub lvwPati_LostFocus()
    Me.lvwPati.Visible = False
End Sub

Private Sub lvwWrits_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwWrits.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwWrits.SortOrder = IIf(Me.lvwWrits.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwWrits.SortKey = ColumnHeader.Index - 1
        Me.lvwWrits.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwWrits_DblClick()
    If Me.lvwWrits.SelectedItem Is Nothing Then Exit Sub
    Call cmdOK_Click
End Sub

Private Sub lvwWrits_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Me.lvwWrits.SelectedItem Is Nothing Then Exit Sub
    Call cmdOK_Click
End Sub

Private Sub txtPati_GotFocus()
    Me.txtPati.SelStart = 0: Me.txtPati.SelLength = 100
End Sub

Private Sub txtPati_KeyPress(KeyAscii As Integer)
    If InStr("~!@#$^&()|=`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii <> vbKeyReturn Then Exit Sub
    Me.txtPati.Text = Trim(Me.txtPati.Text)
    If Me.txtPati.Text = "" Then Me.txtPati.Text = Me.txtPati.Tag: Exit Sub
    
    Select Case Left(Me.txtPati.Text, 1)
    Case "-", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0" '病人ID
        gstrSql = "select 病人ID,门诊号,住院号,姓名,性别,年龄" & _
                " from 病人信息" & _
                " where 病人id=" & Abs(Val(Me.txtPati.Text))
    Case "+"        '住院号
        gstrSql = "select 病人ID,门诊号,住院号,姓名,性别,年龄" & _
                " from 病人信息" & _
                " where 住院号=" & Val(Me.txtPati.Text)
    Case "*"        '门诊号
        gstrSql = "select 病人ID,门诊号,住院号,姓名,性别,年龄" & _
                " from 病人信息" & _
                " where 门诊号=" & Val(Mid(Me.txtPati.Text, 2))
    Case Else       '病人姓名
        gstrSql = "select 病人ID,门诊号,住院号,姓名,性别,年龄" & _
                " from 病人信息" & _
                " where 姓名 like '" & Me.txtPati.Text & "%'"
    End Select
    
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        If .State = adStateOpen Then .Close
        Call SQLTest(App.Title, Me.Caption, gstrSql): .Open gstrSql, gcnOracle: Call SQLTest
        If .BOF Or .EOF = 1 Then
            MsgBox "未找到指定病人", vbExclamation, gstrSysName
            Me.txtPati.Text = "": Me.txtPati.SetFocus: Exit Sub
        End If
        If .RecordCount = 1 Then
            If Val(Me.txtPati.Tag) <> !病人ID Then
                Me.txtPati.Tag = !病人ID: Me.txtPati.Text = Me.txtPati.Tag
                Me.lblInfo.Caption = "姓名:" & Trim(!姓名) & _
                        Space(2) & "性别:" & IIf(IsNull(!性别), "", !性别) & _
                        Space(2) & "年龄:" & IIf(IsNull(!年龄), "", !年龄)
                Me.lblInfo.Tag = !姓名
            End If
            Call RefereshWrits
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwPati.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwPati.ListItems.Add(, "_" & !病人ID, !病人ID)
            objItem.SubItems(Me.lvwPati.ColumnHeaders("门诊号").Index - 1) = IIf(IsNull(!门诊号), "", !门诊号)
            objItem.SubItems(Me.lvwPati.ColumnHeaders("住院号").Index - 1) = IIf(IsNull(!住院号), "", !住院号)
            objItem.SubItems(Me.lvwPati.ColumnHeaders("姓名").Index - 1) = IIf(IsNull(!姓名), "", !姓名)
            objItem.SubItems(Me.lvwPati.ColumnHeaders("性别").Index - 1) = IIf(IsNull(!性别), "", !性别)
            objItem.SubItems(Me.lvwPati.ColumnHeaders("年龄").Index - 1) = IIf(IsNull(!年龄), "", !年龄)
            .MoveNext
        Loop
        Me.lvwPati.ListItems(1).Selected = True
    End With
    With Me.lvwPati
        .Left = Me.txtPati.Left
        .Top = Me.txtPati.Top + Me.txtPati.Height
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtPati_LostFocus()
    Me.txtPati.Text = Me.txtPati.Tag
End Sub

Private Sub RefereshWrits()
    gstrSql = "select ID,Rownum As 序号,书写人,书写日期 From 病人病历记录 where 病人ID=" & Me.txtPati.Tag & " and 文件id=" & Me.lblWrit.Tag
    Err = 0: On Error GoTo ErrHand
    Me.cmdOK.Enabled = False
    Me.lvwWrits.ListItems.Clear
    With rsTemp
        If .State = adStateOpen Then .Close
        Call SQLTest(App.Title, Me.Caption, gstrSql): .Open gstrSql, gcnOracle: Call SQLTest
        Do While Not .EOF
            Set objItem = Me.lvwWrits.ListItems.Add(, "_" & !ID, !序号)
            objItem.SubItems(Me.lvwWrits.ColumnHeaders("书写人").Index - 1) = IIf(IsNull(!书写人), "", !书写人)
            objItem.SubItems(Me.lvwWrits.ColumnHeaders("书写日期").Index - 1) = IIf(IsNull(!书写日期), "", Format(!书写日期, "YYYY-MM-DD HH:MM"))
            .MoveNext
        Loop
    End With
    If Me.lvwWrits.ListItems.Count > 0 Then Me.cmdOK.Enabled = True
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
