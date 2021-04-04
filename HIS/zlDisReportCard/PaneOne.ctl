VERSION 5.00
Begin VB.UserControl PaneOne 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0FFC0&
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9825
   ScaleHeight     =   975
   ScaleWidth      =   9825
   Begin VB.TextBox txtNumber 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   975
      TabIndex        =   0
      Tag             =   "137,119"
      Top             =   660
      Width           =   1455
   End
   Begin zlDisReportCard.uCheckNorm ucCheckType 
      Height          =   270
      Index           =   0
      Left            =   5145
      TabIndex        =   1
      Tag             =   "414,115"
      Top             =   615
      Width           =   1575
      _ExtentX        =   51038
      _ExtentY        =   476
      Checked         =   -1  'True
      BackColor       =   -2147483643
      Caption         =   "1、 初次报告"
      BoxVisible      =   0   'False
   End
   Begin zlDisReportCard.uCheckNorm ucCheckType 
      Height          =   270
      Index           =   1
      Left            =   6735
      TabIndex        =   2
      Tag             =   "509,115"
      Top             =   615
      Width           =   1575
      _ExtentX        =   51038
      _ExtentY        =   476
      BackColor       =   -2147483643
      Caption         =   "2、订正报告"
      BoxVisible      =   0   'False
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "卡片编号："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Tag             =   "78,119"
      Top             =   660
      Width           =   900
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "中华人民共和国传染病报告卡"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2550
      TabIndex        =   4
      Tag             =   "241,88"
      Top             =   0
      Width           =   4875
   End
   Begin VB.Line Line1 
      Index           =   0
      Tag             =   "137,130,238"
      X1              =   975
      X2              =   2400
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblReport 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "报卡类别："
      Height          =   180
      Index           =   1
      Left            =   4215
      TabIndex        =   3
      Tag             =   "349,119"
      Top             =   660
      Width           =   900
   End
End
Attribute VB_Name = "PaneOne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mcolLoadData As Collection  '保存控件显示信息

Public Sub SetMyFoucs()
    Call txtNumber.SetFocus
    txtNumber.SelStart = Len(txtNumber.Text)
End Sub

Public Function HaveChanged() As Boolean
'判断控件显示信息是否发生变化
    Dim objCtl As Control
    Dim i As Integer
    i = 0
    HaveChanged = False
    If mcolLoadData Is Nothing Then
        Set mcolLoadData = New Collection
    End If
    If mcolLoadData.Count <= 0 Then
        Exit Function
    End If
    For Each objCtl In UserControl.Controls
        Select Case TypeName(objCtl)
            Case "TextBox"
                If objCtl.Text <> mcolLoadData("K" & i) Then
                    HaveChanged = True
                    Exit Function
                End If
            Case "uCheckNorm"
                If IIf(objCtl.Checked = True, 1, 0) <> mcolLoadData("K" & i) Then
                    HaveChanged = True
                    Exit Function
                End If
        End Select
        i = i + 1
    Next
End Function

Private Sub SaveLoadData()
'功能：保存控件显示信息
    Dim objCtl As Control
    Dim i As Integer
    i = 0
    Set mcolLoadData = New Collection
    For Each objCtl In UserControl.Controls
        Select Case TypeName(objCtl)
            Case "TextBox"
                Call mcolLoadData.Add(objCtl.Text, "K" & i)
            Case "uCheckNorm"
                Call mcolLoadData.Add(IIf(objCtl.Checked = True, 1, 0), "K" & i)
        End Select
        i = i + 1
    Next
End Sub

Public Sub ClearMe()
    Dim objCtl As Control
    
    On Error GoTo errHand
    For Each objCtl In UserControl.Controls
        Call ClearInfo(objCtl)
    Next
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Public Sub PrintOne()
'功能:打印
    Dim objCtl As Control
    For Each objCtl In UserControl.Controls
        Call PrintInfo(objCtl)
    Next
End Sub

Public Function MakeSaveSql(arrSql() As Variant, colCls As Collection, strFileId As String) As Boolean
'功能:合成保存Sql语句
    Dim strObjNo As String
    Dim strContent As String
    Dim strReportInfo As String
    Dim strTmp As String
    
    On Error GoTo errHand
    strObjNo = "1$2"
    
    strTmp = IIf(ucCheckType(0).Checked = True, 1, 2)
    
    strContent = txtNumber.Text & "$"
    strTmp = IIf(ucCheckType(0).Checked = True, ucCheckType(0).Caption, ucCheckType(1).Caption)
    strContent = strContent & strTmp & "$"
    
    strReportInfo = strObjNo & "|" & strContent
    MakeSaveSql = GetSaveSql(arrSql, colCls, strFileId, strReportInfo)
    Call SaveLoadData
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Function

Public Sub LoadData(colData As Collection, bytType As Byte, ByVal strChkType As String)
    On Error GoTo errHand
    Dim objCtl As Control
    Dim strTmp As String
    If bytType = 1 Then
        txtNumber.Text = CStr(colData("K1"))       '卡片编号
        For Each objCtl In UserControl.Controls
            If TypeName(objCtl) = "uCheckNorm" Then
                strTmp = Trim(objCtl.Caption)
                strTmp = Replace(strTmp, "(", "")
                strTmp = Replace(strTmp, ")", "")
                strTmp = Replace(strTmp, "、", "")
                If InStr(strChkType, strTmp) > 0 And Trim(strTmp) <> "" Then
                    objCtl.Checked = True
                End If
            End If
        Next
    End If
    Call SaveLoadData
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub
Private Sub UserControl_Initialize()
    UserControl.BackColor = vbWindowBackground
End Sub
