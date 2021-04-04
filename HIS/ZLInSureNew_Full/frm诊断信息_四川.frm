VERSION 5.00
Begin VB.Form frm诊断信息_四川 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "请录入诊断信息：标准的ICD-10编码"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   Icon            =   "frm诊断信息_四川.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Txt疾病信息3 
      Height          =   300
      Left            =   1260
      TabIndex        =   2
      Top             =   1620
      Width           =   3675
   End
   Begin VB.TextBox Txt疾病信息2 
      Height          =   300
      Left            =   1260
      TabIndex        =   1
      Top             =   1080
      Width           =   3675
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2490
      TabIndex        =   3
      Top             =   2310
      Width           =   1150
   End
   Begin VB.CommandButton cmdCanc 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3870
      TabIndex        =   6
      Top             =   2310
      Width           =   1150
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   60
      TabIndex        =   5
      Top             =   2130
      Width           =   5475
   End
   Begin VB.TextBox txt疾病信息 
      Height          =   300
      Left            =   1260
      TabIndex        =   0
      Top             =   540
      Width           =   3675
   End
   Begin VB.Label Lbl疾病信息3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "疾病信息3"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   9
      Top             =   1680
      Width           =   810
   End
   Begin VB.Label Lbl疾病信息2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "疾病信息2"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   8
      Top             =   1140
      Width           =   810
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   510
      TabIndex        =   7
      Top             =   150
      Width           =   4425
   End
   Begin VB.Label lbl疾病信息 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "疾病信息1"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Width           =   810
   End
End
Attribute VB_Name = "frm诊断信息_四川"
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
Private mlng主页ID As Long

Private mint诊断类型 As Integer
Private mint诊断数量 As Integer
Private mint险类 As Integer

Private mbln必须录入 As Boolean


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
    

    'HIS+

    If txt疾病信息.Text <> "" Then
        gstrSQL = "ZL_诊断情况补充信息_INSERT(" & mint诊断类型 & "," & mlng病人ID & "," & mlng主页ID & "," & txt疾病信息.Tag & "," & _
                  "'" & txt疾病信息.Text & "',1 )"
    End If
    ExecuteProcedure_南充阆中 "保存诊断记录到中间库"
    If Txt疾病信息2.Text <> "" Then
        gstrSQL = "ZL_诊断情况补充信息_INSERT(" & mint诊断类型 & "," & mlng病人ID & "," & mlng主页ID & "," & Txt疾病信息2.Tag & "," & _
                  "'" & Txt疾病信息2.Text & "',2 )"
    End If
    ExecuteProcedure_南充阆中 "保存诊断记录到中间库"
    If Txt疾病信息3.Text <> "" Then
        gstrSQL = "ZL_诊断情况补充信息_INSERT(" & mint诊断类型 & "," & mlng病人ID & "," & mlng主页ID & "," & Txt疾病信息3.Tag & "," & _
                  "'" & Txt疾病信息3.Text & "',3 )"
    End If
    ExecuteProcedure_南充阆中 "保存诊断记录到中间库"
    Unload Me
End Sub

Public Sub ShowME(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal int诊断类型 As Integer, ByVal int诊断数量 As Integer, ByVal int险类 As Integer, Optional ByVal bln必须录入 As Boolean = True)
'mint诊断类型:2表示入院诊断 3表示出院诊断(首页整理)
'mint诊断数量:允许输入的最大诊断数量－1(在出院界面可以录入一个诊断),本界面最多支持输入3个
    mlng病人ID = lng病人ID
    mint诊断类型 = int诊断类型
    mint诊断数量 = int诊断数量
    mint险类 = int险类
    mlng主页ID = lng主页ID
    
    mbln必须录入 = bln必须录入
    
    Me.Show 1

End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    If mint诊断类型 = 1 Then
        Me.Caption = "请补充录入入院诊断信息：标准的ICD-10编码"
    Else
        Me.Caption = "请补充录入出院诊断信息：标准的ICD-10编码"
    End If
    If mint诊断数量 = 1 Then
        lbl疾病信息.Enabled = True
        txt疾病信息.Enabled = True
        Lbl疾病信息2.Enabled = False
        Txt疾病信息2.Enabled = False
        Lbl疾病信息3.Enabled = False
        Txt疾病信息3.Enabled = False
    End If
    If mint诊断数量 = 2 Then
        lbl疾病信息.Enabled = True
        txt疾病信息.Enabled = True
        Lbl疾病信息2.Enabled = True
        Txt疾病信息2.Enabled = True
        Lbl疾病信息3.Enabled = False
        Txt疾病信息3.Enabled = False
    End If
    If mint诊断数量 = 3 Then
        lbl疾病信息.Enabled = True
        txt疾病信息.Enabled = True
        Lbl疾病信息2.Enabled = True
        Txt疾病信息2.Enabled = True
        Lbl疾病信息3.Enabled = True
        Txt疾病信息3.Enabled = True
    End If
    '提取病人信息
    gstrSQL = " Select A.姓名,A.性别,B.医保号,B.卡号 " & _
              " From 病人信息 A,保险帐户 B" & _
              " Where A.病人ID=B.病人ID And A.病人ID=" & mlng病人ID & " And B.险类=" & mint险类
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人信息")
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
                txt疾病信息.Text = rsTmp!名称
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

Private Sub txt疾病信息2_GotFocus()
    zlControl.TxtSelAll Txt疾病信息2
End Sub

Private Sub txt疾病信息3_GotFocus()
    zlControl.TxtSelAll Txt疾病信息3
End Sub
Private Sub txt疾病信息2_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim strLike As String, str性别 As String
    Dim vPoint As POINTAPI, StrInput As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Txt疾病信息2.Text = Lbl疾病信息2.Tag And Txt疾病信息2.Text <> "" Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf Txt疾病信息2.Text = "" Then
            Txt疾病信息2.Tag = "": lbl疾病信息.Tag = ""
            Call zlCommFun.PressKey(vbKeyTab) '允许不输入
        Else
            strLike = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
            StrInput = UCase(Txt疾病信息2.Text)
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
            vPoint = GetCoordPos(Me.hwnd, Txt疾病信息2.Left, Txt疾病信息2.Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "疾病编码Input", , , , , , True, vPoint.x, vPoint.y, Txt疾病信息2.Height, blnCancel, , True)
            If Not rsTmp Is Nothing Then
                Txt疾病信息2.Tag = rsTmp!ID
                Txt疾病信息2.Text = rsTmp!名称
                Lbl疾病信息2.Tag = Txt疾病信息2.Text '用于恢复显示
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                If Not blnCancel Then
                    MsgBox "没有找到匹配的疾病编码。", vbInformation, gstrSysName
                End If
                If Lbl疾病信息2.Tag <> "" Then Txt疾病信息2.Text = Lbl疾病信息2.Tag
                Call txt疾病信息2_GotFocus
                Txt疾病信息2.SetFocus
            End If
        End If
    End If
End Sub
Private Sub txt疾病信息3_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim strLike As String, str性别 As String
    Dim vPoint As POINTAPI, StrInput As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Txt疾病信息3.Text = Lbl疾病信息3.Tag And Txt疾病信息3.Text <> "" Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf Txt疾病信息3.Text = "" Then
            Txt疾病信息3.Tag = "": Lbl疾病信息3.Tag = ""
            Call zlCommFun.PressKey(vbKeyTab) '允许不输入
        Else
            strLike = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
            StrInput = UCase(Txt疾病信息3.Text)
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
            vPoint = GetCoordPos(Me.hwnd, Txt疾病信息3.Left, Txt疾病信息3.Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "疾病编码Input", , , , , , True, vPoint.x, vPoint.y, Txt疾病信息3.Height, blnCancel, , True)
            If Not rsTmp Is Nothing Then
                Txt疾病信息3.Tag = rsTmp!ID
                Txt疾病信息3.Text = rsTmp!名称
                Lbl疾病信息3.Tag = Txt疾病信息3.Text '用于恢复显示
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                If Not blnCancel Then
                    MsgBox "没有找到匹配的疾病编码。", vbInformation, gstrSysName
                End If
                If Lbl疾病信息3.Tag <> "" Then Txt疾病信息3.Text = Lbl疾病信息3.Tag
                Call txt疾病信息3_GotFocus
                Txt疾病信息3.SetFocus
            End If
        End If
    End If
End Sub

