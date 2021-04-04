VERSION 5.00
Begin VB.Form frmExecuteSet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   Icon            =   "frmExecuteSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdRegPrint 
      Caption         =   "执行登记单打印设置"
      Height          =   350
      Left            =   2295
      TabIndex        =   21
      Top             =   2595
      Width           =   1860
   End
   Begin VB.ListBox lst类别 
      Columns         =   2
      ForeColor       =   &H80000012&
      Height          =   2580
      IMEMode         =   3  'DISABLE
      ItemData        =   "frmExecuteSet.frx":058A
      Left            =   165
      List            =   "frmExecuteSet.frx":058C
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   360
      Width           =   2040
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5235
      TabIndex        =   2
      Top             =   3270
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4080
      TabIndex        =   1
      Top             =   3270
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   120
      TabIndex        =   4
      Top             =   3030
      Width           =   6375
   End
   Begin VB.Frame Frame1 
      Caption         =   "单据来源"
      Height          =   2235
      Left            =   2295
      TabIndex        =   3
      Top             =   270
      Width           =   4125
      Begin VB.Frame fra来源 
         Height          =   1530
         Index           =   2
         Left            =   2760
         TabIndex        =   17
         Top             =   600
         Width           =   1260
         Begin VB.OptionButton opt2 
            Caption         =   "未审核"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1020
         End
         Begin VB.OptionButton opt2 
            Caption         =   "已审核"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   19
            Top             =   525
            Width           =   1020
         End
         Begin VB.OptionButton opt2 
            Caption         =   "所有单据"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   18
            Top             =   795
            Width           =   1020
         End
      End
      Begin VB.CheckBox chk来源 
         Caption         =   "体检"
         Height          =   195
         Index           =   2
         Left            =   2760
         TabIndex        =   16
         Top             =   360
         Value           =   1  'Checked
         Width           =   660
      End
      Begin VB.CheckBox chk来源 
         Caption         =   "住院"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   15
         Top             =   360
         Value           =   1  'Checked
         Width           =   660
      End
      Begin VB.CheckBox chk来源 
         Caption         =   "门诊"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   14
         Top             =   360
         Value           =   1  'Checked
         Width           =   660
      End
      Begin VB.Frame fra来源 
         Height          =   1530
         Index           =   1
         Left            =   1440
         TabIndex        =   10
         Top             =   600
         Width           =   1260
         Begin VB.OptionButton opt1 
            Caption         =   "所有单据"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   13
            Top             =   795
            Width           =   1020
         End
         Begin VB.OptionButton opt1 
            Caption         =   "已审核"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   525
            Width           =   1020
         End
         Begin VB.OptionButton opt1 
            Caption         =   "未审核"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1020
         End
      End
      Begin VB.Frame fra来源 
         Height          =   1530
         Index           =   0
         Left            =   165
         TabIndex        =   6
         Top             =   600
         Width           =   1260
         Begin VB.OptionButton opt0 
            Caption         =   "所有单据"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   9
            Top             =   795
            Width           =   1020
         End
         Begin VB.OptionButton opt0 
            Caption         =   "已收费"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   525
            Width           =   1020
         End
         Begin VB.OptionButton opt0 
            Caption         =   "未收费"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1020
         End
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "项目类别"
      Height          =   180
      Left            =   180
      TabIndex        =   5
      Top             =   90
      Width           =   720
   End
End
Attribute VB_Name = "frmExecuteSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mblnOk As Boolean
Public mstrPrivs As String
Public mlngModul As Long

Private Sub chk来源_Click(Index As Integer)
    If chk来源(0).Value = 0 And chk来源(1).Value = 0 And chk来源(2).Value = 0 Then
        chk来源((Index + 1) Mod 3).Value = 1
    End If
    fra来源(Index).Enabled = chk来源(Index).Value = 1
    Call SetOptionState

End Sub
Private Sub SetOptionState()
    Dim i As Integer
    
    For i = 0 To 2
        opt0(i).Enabled = fra来源(0).Enabled
        opt1(i).Enabled = fra来源(1).Enabled
        opt2(i).Enabled = fra来源(2).Enabled
    Next
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strTmp As String, i As Integer, j As Integer
    Dim blnHavePrivs As Boolean
    
    blnHavePrivs = InStr(1, mstrPrivs, ";参数设置;") > 0
    
    For i = 0 To lst类别.ListCount - 1
        If lst类别.Selected(i) Then
            strTmp = strTmp & ",'" & Chr(lst类别.ItemData(i)) & "'"
        End If
    Next
    
    strTmp = Mid(strTmp, 2)
    If strTmp = "" Then
        MsgBox "请至少选择一种类别。", vbInformation, gstrSysName
        lst类别.SetFocus: Exit Sub
    End If
    If UBound(Split(strTmp, ",")) + 1 = lst类别.ListCount Then strTmp = ""
    
    zlDatabase.SetPara "医技执行类别", strTmp, glngSys, mlngModul, blnHavePrivs
    
    strTmp = IIf(chk来源(0).Value = 1, "1", "0") & IIf(chk来源(1).Value = 1, "1", "0") & IIf(chk来源(2).Value = 1, "1", "0")
    zlDatabase.SetPara "医技病人来源", strTmp, glngSys, mlngModul, blnHavePrivs
    For j = 0 To 2
        If chk来源(j).Value = 1 Then
            strTmp = ""
            For i = 0 To 2
                If j = 0 Then
                    If opt0(i).Value = True Then strTmp = i: Exit For
                ElseIf j = 1 Then
                    If opt1(i).Value = True Then strTmp = i: Exit For
                Else
                    If opt2(i).Value = True Then strTmp = i: Exit For
                End If
            Next
            If strTmp = "" Then strTmp = "2"
            zlDatabase.SetPara Choose(j + 1, "医技门诊单据类型", "医技住院单据类型", "医技体检单据类型"), strTmp, glngSys, mlngModul, blnHavePrivs
        End If
    Next
    
    Call InitLocPar(mlngModul)
    
    mblnOk = True
    Unload Me
End Sub

Private Sub cmdRegPrint_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1142", Me)
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, str类别 As String, i As Long, blnParSet As Boolean
    
    mblnOk = False
    blnParSet = InStr(mstrPrivs, ";参数设置;") > 0
    
    str类别 = zlDatabase.GetPara("医技病人来源", glngSys, mlngModul, "111", Array(chk来源(0), chk来源(1), chk来源(2)), blnParSet)
    '处理旧数据
    If Len(str类别) = 1 Then
        If str类别 = "0" Then
            str类别 = "111"
        ElseIf str类别 = "1" Then
            str类别 = "101"
        Else
            str类别 = "010"
        End If
    End If
    
    chk来源(0).Value = Val(Mid(str类别, 1, 1))
    chk来源(1).Value = Val(Mid(str类别, 2, 1))
    chk来源(2).Value = Val(Mid(str类别, 3, 1))
    
    i = Val(zlDatabase.GetPara("医技门诊单据类型", glngSys, mlngModul, 2, Array(opt0(0), opt0(1), opt0(2)), blnParSet))
    opt0(i).Value = True
    i = Val(zlDatabase.GetPara("医技住院单据类型", glngSys, mlngModul, 2, Array(opt1(0), opt1(1), opt1(2)), blnParSet))
    opt1(i).Value = True
    i = Val(zlDatabase.GetPara("医技体检单据类型", glngSys, mlngModul, 2, Array(opt2(0), opt2(1), opt2(2)), blnParSet))
    opt2(i).Value = True
    
    
    fra来源(0).Enabled = chk来源(0).Value = 1
    fra来源(1).Enabled = chk来源(1).Value = 1
    fra来源(2).Enabled = chk来源(2).Value = 1
    
    lst类别.Clear
    str类别 = zlDatabase.GetPara("医技执行类别", glngSys, mlngModul, "", Array(lst类别), blnParSet)
    Err = 0: On Error GoTo errH:
    strSQL = "Select 编码,名称,简码,固定,序号 From 收费项目类别 Where 编码 Not IN('1','5','6','7','J') Order by 编码"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        lst类别.AddItem rsTmp!名称
        lst类别.ItemData(lst类别.NewIndex) = Asc(rsTmp!编码)
        If str类别 = "" Then
            lst类别.Selected(lst类别.NewIndex) = True
        Else
            If InStr(str类别, "'" & rsTmp!编码 & "'") > 0 Then
                lst类别.Selected(lst类别.NewIndex) = True
            End If
        End If
        rsTmp.MoveNext
    Next
    lst类别.ListIndex = 0
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
