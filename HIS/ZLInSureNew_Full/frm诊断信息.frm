VERSION 5.00
Begin VB.Form frm诊断信息 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "请录入诊断信息：标准的ICD-10编码"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   Icon            =   "frm诊断信息.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2490
      TabIndex        =   3
      Top             =   1500
      Width           =   1150
   End
   Begin VB.CommandButton cmdCanc 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3870
      TabIndex        =   4
      Top             =   1500
      Width           =   1150
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   60
      TabIndex        =   2
      Top             =   1320
      Width           =   5475
   End
   Begin VB.TextBox txt疾病信息 
      Height          =   300
      Left            =   1260
      TabIndex        =   0
      Top             =   810
      Width           =   3675
   End
   Begin VB.Label lblOld 
      Height          =   225
      Left            =   240
      TabIndex        =   6
      Top             =   450
      Width           =   4725
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   150
      Width           =   4425
   End
   Begin VB.Label lbl疾病信息 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "疾病信息"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   1
      Top             =   870
      Width           =   810
   End
End
Attribute VB_Name = "frm诊断信息"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type POINTAPI
     x As Long
     y As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private lngTXTProc As Long '保存默认的消息函数的地址
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private str性别 As String
Private mlng病人ID As Long
Private mbln必须录入 As Boolean
Private mstr诊断编码 As String
Private mstr诊断名称 As String

Private Sub cmdCanc_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If mbln必须录入 Then
        If Val(txt疾病信息.Tag) = 0 Then
            MsgBox "医保中心要求，必须按ICD-10标准录入诊断信息！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    If Trim(txt疾病信息.Text) <> "" Then
        mstr诊断编码 = Mid(txt疾病信息.Text, 2, InStr(1, txt疾病信息.Text, ")") - 2)
        mstr诊断名称 = Mid(txt疾病信息.Text, InStr(1, txt疾病信息.Text, ")") + 1)
    End If
    
    Unload Me
End Sub

Public Sub ShowME(ByVal lng病人ID As Long, ByRef str诊断编码 As String, ByRef str诊断名称 As String, Optional ByVal bln必须录入 As Boolean = True)
    mlng病人ID = lng病人ID
    mbln必须录入 = bln必须录入
    mstr诊断编码 = str诊断编码
    mstr诊断名称 = str诊断名称
    Me.Show 1
    str诊断编码 = mstr诊断编码
    str诊断名称 = mstr诊断名称
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    lblOld = "原诊断为：(" & mstr诊断编码 & ")" & mstr诊断名称
    mstr诊断编码 = ""
    mstr诊断名称 = ""
    
    '提取病人信息(一个病人不可能属于多个医保)
    gstrSQL = " Select A.姓名,A.性别,B.医保号,B.卡号 " & _
              " From 病人信息 A,保险帐户 B" & _
              " Where A.病人ID=B.病人ID And A.病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人信息", mlng病人ID)
    lblNote.Caption = "姓名:" & rsTemp!姓名 & "  性别:" & rsTemp!性别 & "  医保号:" & Nvl(rsTemp!医保号) & "  卡号:" & Nvl(rsTemp!卡号)
    str性别 = rsTemp!性别
End Sub

Private Sub txt疾病信息_GotFocus()
    zlControl.TxtSelAll txt疾病信息
End Sub

Private Sub txt疾病信息_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim strLike As String, str性别 As String
    Dim vPoint As POINTAPI, StrInput As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt疾病信息.Text = lbl疾病信息.Tag And txt疾病信息.Text <> "" Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt疾病信息.Text = "" Then
            txt疾病信息.Tag = "": lbl疾病信息.Tag = ""
            Call zlCommFun.PressKey(vbKeyTab) '允许不输入
        Else
            strLike = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
            StrInput = UCase(txt疾病信息.Text)
            If str性别 = "男" Then
                str性别 = " And (A.性别限制='男' Or A.性别限制 is NULL)"
            ElseIf str性别 = "女" Then
                str性别 = " And (A.性别限制='女' Or A.性别限制 is NULL)"
            Else
                str性别 = ""
            End If
            strSQL = "Select A.ID,A.编码,A.附码,A.名称,A.简码,A.说明,A.性别限制,B.类别" & _
                " From 疾病编码目录 A,疾病编码类别 B" & _
                " Where A.类别=B.编码 And A.类别 Not IN('B','Z')" & _
                " And (A.编码 Like '" & StrInput & "%'" & _
                " Or Upper(A.名称) Like '" & strLike & StrInput & "%'" & _
                " Or Upper(A.简码) Like '" & strLike & StrInput & "%'" & _
                " Or Upper(A.附码) Like '" & strLike & StrInput & "%')" & _
                " And Rownum<=100" & str性别 & _
                " Order by A.类别,A.编码"
            vPoint = GetCoordPos(Me.hwnd, txt疾病信息.Left, txt疾病信息.Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "疾病编码Input", , , , , , True, vPoint.x, vPoint.y, txt疾病信息.Height, blnCancel, , True)
            If Not rsTmp Is Nothing Then
                txt疾病信息.Tag = rsTmp!ID
                txt疾病信息.Text = "(" & rsTmp!编码 & ")" & rsTmp!名称
                lbl疾病信息.Tag = txt疾病信息.Text '用于恢复显示
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                If Not blnCancel Then
                    MsgBox "没有找到匹配的疾病编码。", vbInformation, gstrSysName
                End If
                If lbl疾病信息.Tag <> "" Then txt疾病信息.Text = lbl疾病信息.Tag
                Call txt疾病信息_GotFocus
                txt疾病信息.SetFocus
            End If
        End If
    End If
End Sub

Private Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
'功能：得控件中指定坐标在屏幕中的位置(Twip)
    Dim vPoint As POINTAPI
    vPoint.x = lngX / Screen.TwipsPerPixelX: vPoint.y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.x = vPoint.x * Screen.TwipsPerPixelX: vPoint.y = vPoint.y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

Private Function GetControlRect(ByVal lngHwnd As Long) As RECT
'功能：获取指定控件在屏幕中的位置(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function
