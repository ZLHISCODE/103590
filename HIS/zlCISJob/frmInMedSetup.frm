VERSION 5.00
Begin VB.Form frmInMedSetup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "病案项目增加"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4200
      TabIndex        =   10
      Top             =   4320
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2880
      TabIndex        =   6
      Top             =   4320
      Width           =   1100
   End
   Begin VB.TextBox txtEdit 
      Height          =   1335
      Index           =   2
      Left            =   960
      MaxLength       =   1000
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2880
      Width           =   4455
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   1
      Left            =   960
      MaxLength       =   20
      TabIndex        =   3
      Top             =   2460
      Width           =   4455
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   0
      Left            =   960
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1980
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   0
      TabIndex        =   9
      Top             =   1800
      Width           =   5535
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   120
      Picture         =   "frmInMedSetup.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "内容(&R)"
      Height          =   180
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "名称(&N)"
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "编码(&B)"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   630
   End
   Begin VB.Label lbl内容 
      Caption         =   "内容包括："
      Height          =   1260
      Left            =   720
      TabIndex        =   8
      Top             =   480
      Width           =   4740
   End
   Begin VB.Label lbl信息 
      Caption         =   "用户自定义的病案首页项目.包括项目的编码、名称、内容."
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmInMedSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum TextIndex
    Num编码 = 0: Num名称: Num内容
End Enum
Private mblnChange As Boolean
Private mstr方式 As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    Dim strSQL As String
    
    If isVaild = False Then Exit Sub
On Error GoTo errHandle
    strSQL = "zl_病案项目_edit('" & txtEdit(TextIndex.Num编码).Text & "','" & txtEdit(TextIndex.Num名称).Text & "','" & txtEdit(TextIndex.Num内容).Text & "','" & lblEdit(TextIndex.Num编码).Tag & "'," & IIf(mstr方式 = "新增", 0, 1) & ")"
    Call zldatabase.ExecuteProcedure(strSQL, Me.Caption)
    With frmInStationSetup.vsfMain
        If mstr方式 = "新增" Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = txtEdit(TextIndex.Num编码).Text
            .TextMatrix(.Rows - 1, 1) = txtEdit(TextIndex.Num名称).Text
            .TextMatrix(.Rows - 1, 2) = txtEdit(TextIndex.Num内容).Text
            For i = 0 To txtEdit.Count - 1
                txtEdit(i).Text = ""
            Next i
            txtEdit(TextIndex.Num编码).Text = get编码
            txtEdit(TextIndex.Num名称).SetFocus
            mblnChange = False
        Else
            i = .FindRow(lblEdit(TextIndex.Num编码).Tag, , 0)
            .Cell(flexcpText, i, 0, i, .Cols - 1) = ""
            .TextMatrix(i, 0) = txtEdit(TextIndex.Num编码).Text
            .TextMatrix(i, 1) = txtEdit(TextIndex.Num名称).Text
            .TextMatrix(i, 2) = txtEdit(TextIndex.Num内容).Text
            mblnChange = False
            Unload Me
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    txtEdit(TextIndex.Num名称).SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    lbl信息.Caption = "用户自定义的病案首页项目.包括项目的编码、名称、内容."
    lbl内容.Caption = "内容:" & Chr(13) & Chr(10)
    lbl内容.Caption = lbl内容.Caption & "1.空：表示为文本录入项目，可自由录入文字" & Chr(13) & Chr(10)
    lbl内容.Caption = lbl内容.Caption & "2.值序列 AAA,BBB,CCC,DDD：表示从下拉框中选择指定的项" & Chr(13) & Chr(10)
    lbl内容.Caption = lbl内容.Caption & "3.逻辑值序列 ""是否""：表示勾选方式" & Chr(13) & Chr(10)
    lbl内容.Caption = lbl内容.Caption & "4.数字范围 -100...100或0.1-0.9:表示输入指定范围的数字" & Chr(13) & Chr(10)
End Sub

Private Function isVaild() As Boolean
    Dim i As Integer
    Dim sngNum1, sngNum2 As Single
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim LngSize As Long
    
    For i = 0 To txtEdit.Count - 1
    '检验是否包含特殊字符
        If zlCommFun.StrIsValid(txtEdit(i).Text) = False Then
            txtEdit(i).SetFocus
            isVaild = False
            Exit Function
        End If
        If txtEdit(i).MaxLength <> 0 Then
            If zlCommFun.ActualLen(txtEdit(i).Text) > txtEdit(i).MaxLength Then
                MsgBox "输入长度不能大于" & "[" & txtEdit(i).MaxLength & "]！", vbInformation, gstrSysName
                isVaild = False
                txtEdit(i).SetFocus
                Exit Function
            End If
        End If
    Next i
    On Error GoTo errH
    '为空表示是录入项目不用再进行内容检测
    If txtEdit(2).Text = "" Then isVaild = True: Exit Function
        
    If InStr(txtEdit(TextIndex.Num内容).Text, ",") <> 0 Then
        strSQL = "SELECT 信息值 from 病案主页从表 where 病人id=0 and 主页id=0"
'        On Error Resume Next
        Set rsTemp = zldatabase.OpenSQLRecord(strSQL, Me.Caption)
        LngSize = rsTemp.Fields("信息值").DefinedSize
        For i = 0 To UBound(Split(txtEdit(TextIndex.Num内容).Text, ",")) - 1
            If zlCommFun.ActualLen(Split(txtEdit(TextIndex.Num内容).Text, ",")(i)) > LngSize Then
                MsgBox "[" & Split(txtEdit(TextIndex.Num内容).Text, ",")(i) & "]选项的长度不能超过" & LngSize & "请修改!", vbInformation, gstrSysName
                isVaild = False
                txtEdit(TextIndex.Num内容).SetFocus
                Exit Function
            End If
        Next i
        isVaild = True
    End If
    
    '当内容是数字范围类型是检测是否合法
    If InStr(txtEdit(TextIndex.Num内容).Text, "...") > 0 Then
        sngNum1 = Mid(txtEdit(2).Text, 1, InStr(txtEdit(2).Text, "...") - 1)
        sngNum2 = Mid(txtEdit(2).Text, InStr(txtEdit(2).Text, "...") + 3)
        If Not IsNumeric(sngNum1) Or Not IsNumeric(sngNum2) Then
            MsgBox "请输入正确的数字范围!", vbInformation, gstrSysName
            txtEdit(2).SetFocus
            isVaild = False
            Exit Function
        End If
    ElseIf InStr(txtEdit(TextIndex.Num内容).Text, "-") > 0 Then
    
        If InStr(txtEdit(2).Text, "-") = 1 Then
            If (InStr(2, txtEdit(2).Text, "-") - 1) > 0 Then
                sngNum1 = Mid(txtEdit(2).Text, 2, InStr(2, txtEdit(2).Text, "-") - 1)
                sngNum2 = Mid(txtEdit(2).Text, InStr(2, txtEdit(2).Text, "-") + 1)
            End If
        Else
            If Not IsNumeric(Mid(txtEdit(2).Text, 1, InStr(txtEdit(2).Text, "-") - 1)) Or _
                Not IsNumeric(Mid(txtEdit(2).Text, InStr(txtEdit(2).Text, "-") + 1)) Then
                
                MsgBox "请输入正确的数字范围!", vbInformation, gstrSysName
                txtEdit(2).SetFocus
                isVaild = False
                Exit Function
            End If
            sngNum1 = Mid(txtEdit(2).Text, 1, InStr(txtEdit(2).Text, "-") - 1)
            sngNum2 = Mid(txtEdit(2).Text, InStr(txtEdit(2).Text, "-") + 1)
        End If
        If Not IsNumeric(sngNum1) Or Not IsNumeric(sngNum2) Then
            MsgBox "请输入正确的数字范围!", vbInformation, gstrSysName
            txtEdit(2).SetFocus
            isVaild = False
            Exit Function
        End If
    End If
    
    If sngNum2 < sngNum1 Then
        MsgBox "输入的数字范围有误,应从小到大.", vbInformation, gstrSysName
        txtEdit(2).SetFocus
        isVaild = False
        Exit Function
    End If
    '内容是否在内容的定义范畴内
    If InStr(txtEdit(2).Text, ",") = 0 And InStr(txtEdit(2).Text, "是否") = 0 And InStr(txtEdit(2).Text, "...") = 0 And InStr(txtEdit(2).Text, "-") = 0 Then
        MsgBox "请输入正确病案项目内容!", vbInformation, gstrSysName
        isVaild = False
        txtEdit(2).SetFocus
        Exit Function
    End If
    isVaild = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = True Then
        If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    mblnChange = False
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = TextIndex.Num编码 Then
        If InStr("0123456789" & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
    If Index = TextIndex.Num名称 Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If zlCommFun.ActualLen(txtEdit(TextIndex.Num名称).Text) = 20 Then KeyAscii = 0
        End If
    End If
    If Index = TextIndex.Num内容 Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If zlCommFun.ActualLen(txtEdit(TextIndex.Num内容).Text) = 1000 Then KeyAscii = 0
        End If
    End If

End Sub

Public Sub ShowMe(str编码 As String, str名称 As String, str内容 As String, str方式 As String, frmMain As Object)
    lblEdit(TextIndex.Num编码).Tag = str编码
    txtEdit(TextIndex.Num编码).Text = str编码
    
    txtEdit(TextIndex.Num名称).Text = str名称
    txtEdit(TextIndex.Num内容).Text = str内容
    mstr方式 = str方式
    If str方式 = "新增" Then
        txtEdit(TextIndex.Num编码) = get编码
    End If
    mblnChange = False
    Me.Show 1, frmMain
End Sub

Private Function get编码() As String
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "select max(to_number(编码)) as Maxcode from 病案项目"
On Error GoTo errHandle
    Call zldatabase.OpenRecordset(rsTemp, strSQL, Me.Caption)
    get编码 = Right("000" & IIf(IsNull(rsTemp!maxcode), 1, rsTemp!maxcode + 1), 3)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


