VERSION 5.00
Begin VB.Form frmExamineFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "查找病人审批项目"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5880
   Icon            =   "frmExamineFind.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "定位条件"
      Height          =   1455
      Left            =   30
      TabIndex        =   12
      Top             =   30
      Width           =   2895
      Begin VB.ComboBox cbo类别 
         Height          =   300
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   600
         TabIndex        =   3
         Top             =   630
         Width           =   2175
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   600
         TabIndex        =   5
         Top             =   1020
         Width           =   2175
      End
      Begin VB.Label lbl种类 
         AutoSize        =   -1  'True
         Caption         =   "类别"
         Height          =   180
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "编码"
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   690
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "名称"
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   360
      End
   End
   Begin VB.Frame fra匹配 
      Caption         =   "匹配方式"
      Height          =   1455
      Left            =   3000
      TabIndex        =   8
      Top             =   30
      Width           =   1500
      Begin VB.OptionButton optMatch 
         Caption         =   "任意匹配"
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   10
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton optMatch 
         Caption         =   "从左匹配"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   9
         Top             =   450
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "退出(&E)"
      Height          =   350
      Left            =   4680
      TabIndex        =   7
      Top             =   600
      Width           =   1100
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "定位(&L)"
      Height          =   350
      Left            =   4680
      TabIndex        =   6
      Top             =   120
      Width           =   1100
   End
   Begin VB.Label lbl结果 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 请输入查找条件"
      ForeColor       =   &H00008000&
      Height          =   855
      Left            =   30
      TabIndex        =   11
      Top             =   1560
      Width           =   4515
   End
End
Attribute VB_Name = "frmExamineFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mrsfind As New ADODB.Recordset
Private mblnFind As Boolean

Private Sub cbo类别_Click()
    mblnFind = False
    cmdFind.Enabled = True
    cmdFind.Caption = "定位(&L)"
End Sub

Private Sub cmdCancel_Click()
    mblnFind = False
    Set mrsfind = Nothing
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error GoTo ErrHandle
    
    If mblnFind = True Then
        If Not mrsfind.EOF Then mrsfind.MoveNext
        Call LocateItem
        Exit Sub
    End If
    
    If IsValid = False Then Exit Sub
    gstrSQL = ""
    
    If cbo类别.ListIndex > 0 Then
        gstrSQL = "类别 ='" & cbo类别.Text & "' And "
    End If
    
    If txtEdit(0).Text <> "" Then
        gstrSQL = gstrSQL & "编码 like '" & IIf(optMatch(1).Value = True, "%", "") & UCase(txtEdit(0).Text) & "%' and "
    End If
        
    If txtEdit(1).Text <> "" Then
        gstrSQL = gstrSQL & "名称 like '" & IIf(optMatch(1).Value = True, "%", "") & UCase(txtEdit(1).Text) & "%' and "
    End If
    
    If frmManageExamine.tbsClass.Visible = True Then
        If cbo类别.ListIndex > 0 Then
            Set frmManageExamine.tbsClass.SelectedItem = frmManageExamine.tbsClass.Tabs.Item(cbo类别.Text)
        Else
            Set frmManageExamine.tbsClass.SelectedItem = frmManageExamine.tbsClass.Tabs.Item("全部")
        End If
    End If
    
    If gstrSQL <> "" Then
        gstrSQL = Mid(gstrSQL, 1, Len(gstrSQL) - 4)
    Else
        MsgBox "请输入查找条件。", vbExclamation, gstrSysName
        txtEdit(0).SetFocus
        Exit Sub
    End If
        

    mrsfind.Filter = gstrSQL
    If Not mrsfind.EOF Then
        Call LocateItem
        mblnFind = True
    Else
        lbl结果.Caption = " 已经定位完所有找到的信息，请重新输入条件"
        lbl结果.ForeColor = &HFF&
        Beep
        Exit Sub
    End If
    Exit Sub

ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function IsValid() As Boolean
'功能:分析输入的内容是否有效
'参数:
'返回值:有效返回True,否则为False
    Dim i As Integer
    Dim strTemp As String
    For i = 0 To 1
        strTemp = Trim(txtEdit(i).Text)
        If InStr(strTemp, "'") > 0 Then
            MsgBox "所输入内容含有非法字符。", vbExclamation, gstrSysName
            txtEdit(i).SetFocus
            zlControl.TxtSelAll txtEdit(i)
            Exit Function
        End If
    Next
    IsValid = True
End Function

Private Sub LocateItem()
    If mrsfind.RecordCount = 0 Then
        lbl结果.Caption = " 没有找到符合条件的信息!"
        lbl结果.ForeColor = &HFF&
        Beep
        Exit Sub
    End If
    If mrsfind.EOF = True Then
        lbl结果.Caption = " 已经定位完所有找到的信息，请重新输入条件"
        lbl结果.ForeColor = &HFF&
        Beep
        Exit Sub
    End If
    lbl结果.Caption = "  找到" & mrsfind.RecordCount & "条符合条件的信息。" & vbCrLf & "当前是第" & mrsfind.AbsolutePosition & _
                    "条，" & "名称：" & mrsfind("名称")
    lbl结果.ForeColor = &H8000000D
    
    If mrsfind.RecordCount > 0 Then
        If mrsfind.RecordCount <> mrsfind.AbsolutePosition Then
            cmdFind.Caption = "下一个(&L)"
        Else
            cmdFind.Caption = "定位(&L)"
            cmdFind.Enabled = False
            lbl结果.Caption = lbl结果.Caption & vbCrLf & "已经定位到最后一条信息，请重新输入条件"
        End If
    End If
    
    frmManageExamine.vsExist.Row = frmManageExamine.vsExist.FindRow(mrsfind("编码"), , 1)

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub optMatch_Click(Index As Integer)
    mblnFind = False
    cmdFind.Enabled = True
    cmdFind.Caption = "定位(&L)"
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnFind = False
    cmdFind.Enabled = True
    cmdFind.Caption = "定位(&L)"
End Sub
'问题29712 by lesfeng 2010-05-11
Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If InStr("[]:：;；?？'‘||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    End If
End Sub

