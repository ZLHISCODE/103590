VERSION 5.00
Begin VB.Form frmEPRModelsInfo 
   BackColor       =   &H00E7CFBA&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5730
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   270
      Index           =   2
      Left            =   825
      MaxLength       =   5
      TabIndex        =   4
      Top             =   105
      Width           =   2025
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   270
      Index           =   3
      Left            =   840
      MaxLength       =   30
      TabIndex        =   1
      Top             =   540
      Width           =   2025
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   270
      Index           =   4
      Left            =   3645
      MaxLength       =   10
      TabIndex        =   2
      Top             =   540
      Width           =   1620
   End
   Begin VB.ComboBox cbolevel 
      Enabled         =   0   'False
      Height          =   300
      ItemData        =   "frmEPRModelsInfo.frx":0000
      Left            =   3645
      List            =   "frmEPRModelsInfo.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   1620
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   270
      Index           =   5
      Left            =   855
      MaxLength       =   100
      TabIndex        =   3
      Top             =   975
      Width           =   4410
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E7CFBA&
      Caption         =   "简码"
      Height          =   240
      Index           =   4
      Left            =   3000
      TabIndex        =   9
      Top             =   555
      Width           =   630
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E7CFBA&
      Caption         =   "通用级"
      Height          =   240
      Index           =   6
      Left            =   3000
      TabIndex        =   8
      Top             =   120
      Width           =   630
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E7CFBA&
      Caption         =   "编号"
      Height          =   240
      Index           =   2
      Left            =   315
      TabIndex        =   7
      Top             =   120
      Width           =   630
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E7CFBA&
      Caption         =   "名称"
      Height          =   240
      Index           =   3
      Left            =   315
      TabIndex        =   6
      Top             =   555
      Width           =   630
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E7CFBA&
      Caption         =   "说明"
      Height          =   240
      Index           =   5
      Left            =   315
      TabIndex        =   5
      Top             =   990
      Width           =   630
   End
End
Attribute VB_Name = "frmEPRModelsInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Function zlSaveData() As Boolean

    On Error GoTo ErrHandle
    If Trim(Text1(3).Text) = "" Then
        MsgBox "名称不能为空，请检查！", vbInformation, gstrSysName
        Exit Function
    End If
    If Not CheckLen(Text1(3), 30, "姓名") Then Exit Function
    If Not CheckLen(Text1(2), 5, "编码") Then Exit Function
    If Not CheckLen(Text1(4), 10, "简码") Then Exit Function
    If Not CheckLen(Text1(5), 100, "说明") Then Exit Function
    
    gstrSQL = "zl_病历范文包_Update(" & Val(Text1(2).Tag) & ",'" & Text1(2).Text & "','" & Text1(3).Text & "'" & _
                                    ",'" & Text1(4) & "','" & Text1(5).Text & "'," & NeedNo(cbolevel.Text) & "," & IIf(Label1(5).Tag = "", glngDeptId, Label1(5).Tag) & _
                                    "," & IIf(Text1(5).Tag = "", glngUserId, Text1(5).Tag) & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    zlSaveData = True
    zlEndEdit
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Sub zlEndEdit()
Dim i As Integer
    For i = Text1.LBound To Text1.UBound
        Text1(i).Enabled = False
    Next
    cbolevel.Enabled = False
End Sub
Public Sub zlEditStart()
Dim i As Integer
    For i = Text1.LBound To Text1.UBound
        Text1(i).Enabled = True
    Next
    cbolevel.Enabled = True
End Sub
Public Sub zlRefresh(ByVal strInfo As String, ByVal strPrivs As String)
'strInfo=""表示新增,否则表示修改
'0ID|1编号|2名称|3简码|4说明|5通用级|6科室ID|7人员ID
    With cbolevel
        .Clear
        If InStr(strPrivs, "个人病历范文") > 0 Then .AddItem "2-个人使用"
        If InStr(strPrivs, "科室病历范文") > 0 Then .AddItem "1-科室通用"
        If InStr(strPrivs, "全院病历范文") > 0 Then .AddItem "0-全院通用"
    End With
    If strInfo = "" Then
        Dim rsTemp As ADODB.Recordset
        gstrSQL = "select Max(编号) as 编号 from 病历范文包"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        Text1(2).Text = Format(GetNumber(NVL(rsTemp!编号, "0")) + 1, "00000")
        Text1(2).Tag = ""
        Text1(3).Text = ""
        Text1(4).Text = ""
        Text1(5).Text = ""
        Label1(5).Tag = ""
        Text1(5).Tag = ""
        cbolevel.ListIndex = 0
        zlEditStart
        If Text1(3).Enabled Then
            Text1(3).SetFocus
        End If
        
    Else
        Text1(2).Tag = Split(strInfo, "|")(0) 'ID
        Text1(2).Text = Split(strInfo, "|")(1) '编号
        Text1(3).Text = Split(strInfo, "|")(2) '名称
        Text1(4).Text = Split(strInfo, "|")(3) '简码
        Text1(5).Text = Split(strInfo, "|")(4) '说明
        Call zlControl.CboSetText(cbolevel, Split(strInfo, "|")(5))
        Label1(5).Tag = Split(strInfo, "|")(6) '科室ID
        Text1(5).Tag = Split(strInfo, "|")(7) '人员ID
        If Text1(5).Tag <> "" And Text1(5).Tag <> glngUserId And gstrDbOwner <> gstrDBUser Then cbolevel.Enabled = False '非自已制作的,且当前用户非最高权限不能更改等级
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        ZLCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 3
            If KeyAscii = vbKeyReturn Then
                Text1(4).Text = ZLCommFun.SpellCode(Text1(3).Text)
            End If
        Case 2
            If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        Case Else
            If InStr("',", Chr(KeyAscii)) > 0 Then
                KeyAscii = 0
            End If
    End Select
End Sub
