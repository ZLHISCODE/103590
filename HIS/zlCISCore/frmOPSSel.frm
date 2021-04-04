VERSION 5.00
Begin VB.Form frmOPSSel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "手术选择器"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   Icon            =   "frmOPSSel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "描述内容"
      Height          =   945
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   5655
      Begin VB.TextBox txt1 
         Height          =   300
         Left            =   780
         TabIndex        =   2
         Tag             =   "100"
         Top             =   315
         Width           =   4560
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "手术"
         Height          =   180
         Left            =   330
         TabIndex        =   1
         Top             =   375
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6045
      TabIndex        =   5
      Top             =   195
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6045
      TabIndex        =   6
      Top             =   645
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   6045
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1410
      Width           =   1100
   End
   Begin VB.OptionButton Opt 
      Caption         =   "(&1)当输入诊断时从ICD-9-CM3手术编码里提取"
      Height          =   255
      Index           =   0
      Left            =   375
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1200
      Value           =   -1  'True
      Width           =   4380
   End
   Begin VB.OptionButton Opt 
      Caption         =   "(&2)当输入诊断时从诊疗项目目录的手术项目里提取"
      Height          =   255
      Index           =   1
      Left            =   375
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1515
      Width           =   4380
   End
End
Attribute VB_Name = "frmOPSSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const LAWLChar = "';`|,"""

Private mblnCancel As Boolean
Private mstrTxt1 As String  '疾病描述
Private mlngID1 As Long     '操作ID
Private mlngID2 As Long     '项目ID
Dim i As Long, j As Long

Private strSQL As String
Private rsTmp As New ADODB.Recordset

Public Function ShowSel(frmParent As Object, strReturn As String) As Boolean
    '显示窗体
    Dim strTmp As String
    Dim strTmp1 As String
    Dim i As Long
    
    mblnCancel = False
    '将传入的参数进行分解以便得到以前的设置并设置到本选择器中,不至于在使用了本次选择器后将以前的参数去掉了
    If Trim(strReturn) <> "" Then
        i = InStr(strReturn, ";")
        If i > 0 Then
            '找到描述
            mstrTxt1 = Left(strReturn, i - 1)
            strTmp = Mid(strReturn, i + 1)
            i = InStr(strTmp, ";")
            If i > 0 Then
                '找到操作ID
                strTmp1 = Left(strTmp, i - 1)
                strTmp = Mid(strTmp, i + 1)
                If IsNumeric(strTmp1) Then
                    mlngID1 = CLng(strTmp1)
                Else
                    mlngID1 = 0
                End If
                '项目ID
                If IsNumeric(strTmp) Then
                    mlngID2 = CLng(strTmp)
                Else
                    mlngID2 = 0
                End If
            Else
                mlngID1 = 0
                mlngID2 = 0
            End If
        Else
            mstrTxt1 = ""
            mlngID1 = 0
            mlngID2 = 0
        End If
    Else
        mstrTxt1 = ""
        mlngID1 = 0
        mlngID2 = 0
    End If
    
    
    Me.Show 1, frmParent
    If mblnCancel = False Then
        '返回格式:  疾病描述内容;操作ID;项目ID
        strReturn = Replace(mstrTxt1, ";", "；") & ";" & mlngID1 & ";" & mlngID2
        ShowSel = True
    End If
End Function

Private Function LocalCheck是否非法(txt As Control, ByVal strLawlChar As String) As Boolean
    '功能:检查是不是包含strLawlChar里的字符串,如果有就返回为真否则就返回否
    On Error GoTo ErrHandle
    Dim strSour As String
    
    If TypeOf txt Is TextBox Or TypeOf txt Is ComboBox Then
        If TypeOf txt Is ComboBox Then
            If txt.Style <> 0 Then
                '不管ComboBox为选择的情况，只管输入的情况
                LocalCheck是否非法 = True
                Exit Function
            End If
        End If
        strSour = txt.Text
        If Len(strSour) > 0 Then
            For i = 1 To Len(strLawlChar)
                If InStr(strSour, Mid(strLawlChar, i, 1)) > 0 Then
                    txt.SelStart = InStr(strSour, Mid(strLawlChar, i, 1))
                    txt.SelLength = 1
                    MsgBox "文本里包含有非法字符！", vbInformation, gstrSysName
                    LocalCheck是否非法 = True
                    Exit Function
                End If
            Next
            If VarType(txt.Tag) = vbLong Or VarType(txt.Tag) = vbInteger Then
                If zlCommFun.ActualLen(strSour) > txt.Tag And txt.Tag > 0 Then
                    MsgBox "您所输入的文本超长！", vbInformation, gstrSysName
                    LocalCheck是否非法 = True
                End If
            ElseIf VarType(txt.Tag) = vbString And IsNumeric(txt.Tag) Then
                If zlCommFun.ActualLen(strSour) > CLng(txt.Tag) And CLng(txt.Tag) > 0 Then
                    MsgBox "您所输入的文本超长！", vbInformation, gstrSysName
                    LocalCheck是否非法 = True
                End If
            End If
        End If
    End If
    Exit Function
ErrHandle:
    If gcnOracle Is Nothing Then Exit Function
    If gcnOracle.State <> adStateOpen Then Exit Function
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrHandle
    If KeyCode = 13 And Shift = 0 Then
        If Not TypeOf ActiveControl Is CommandButton Then
            zlCommFun.PressKey vbKeyTab
        End If
    End If
    Exit Sub
ErrHandle:
    If gcnOracle Is Nothing Then Exit Sub
    If gcnOracle.State <> adStateOpen Then Exit Sub
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    mblnCancel = True
    Unload Me
End Sub

Private Sub cmdCancel_GotFocus()
    zlCommFun.OpenIme
End Sub

Private Sub cmdHelp_GotFocus()
    zlCommFun.OpenIme
End Sub

Private Sub cmdOK_Click()
    zlCommFun.OpenIme
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Opt(0).Value = IIf(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\手术选择器", "ICD-9-CM3手术编码", "1") = "1", True, False)
    Me.Opt(1).Value = Not Me.Opt(0).Value
    Me.txt1.Text = mstrTxt1
    Me.txt1.SelStart = Len(Me.txt1.Text)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\手术选择器", "ICD-9-CM3手术编码", IIf(Opt(0).Value = True, "1", "0"))
End Sub

Private Sub txt1_KeyPress(KeyAscii As Integer)
    Dim strWidth As String
    Dim blnMatching As Boolean
    Dim CurPoint As POINTAPI
    
    If InStr("'~|;,.?", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If KeyAscii = vbKeyReturn Then
        If Trim(txt1.Text) <> "" Then
            If Asc(Left(txt1.Text, 1)) < 0 Then
                Exit Sub
            End If
        End If
        If gcnOracle Is Nothing Then Exit Sub
        If gcnOracle.State <> adStateOpen Then Exit Sub
        
        blnMatching = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", "0") = "0", True, False)
        If Opt(0).Value = True Then '从疾病编码
            strSQL = "( UPPER(编码) like '" & UCase(txt1.Text) & "%' or " & _
            "  UPPER(名称) like '" & IIf(blnMatching = True, "%", "") & UCase(txt1.Text) & "%' or " & _
                "  UPPER(简码) like '" & IIf(blnMatching = True, "%", "") & UCase(txt1.Text) & "%' or " & _
                "  UPPER(附码) like '" & IIf(blnMatching = True, "%", "") & UCase(txt1.Text) & "%' ) "
            
            strSQL = "select id,编码,名称,简码,附码 from 疾病编码目录 where 类别='S' AND    " & strSQL
        Else    '从诊疗项目目录
            strSQL = "( UPPER(a.编码) like '" & UCase(txt1.Text) & "%' or " & _
            "  UPPER(a.名称) like '" & IIf(blnMatching = True, "%", "") & UCase(txt1.Text) & "%')"
            
            strSQL = "SELECT a.id,a.编码,a.名称  FROM 诊疗项目目录 a,诊疗项目别名 b WHERE a.id=b.诊疗项目id AND a.类别='F' AND (A.撤档时间 = to_date('3000-01-01','yyyy-mm-dd') OR A.撤档时间 IS NULL) AND " & strSQL
        End If
        
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "手术概要记录单选择器")
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            '定位选择器
            CurPoint.x = (txt1.Left) / Screen.TwipsPerPixelX
            CurPoint.y = (txt1.Top + txt1.Height + Screen.TwipsPerPixelY) / Screen.TwipsPerPixelY
            ClientToScreen Frame1.hwnd, CurPoint
            If Opt(0).Value = True Then '从疾病编码
                '初始选择器
                strWidth = "0;1200;" & IIf(txt1.Width - 1200 - 1200 - 800 - Screen.TwipsPerPixelX * 26 < 1500, "1500", txt1.Width - 1200 - 1200 - 800 - Screen.TwipsPerPixelX * 26) & ";1200;800"
            Else
                '初始选择器
                strWidth = "0;1200;" & IIf(txt1.Width - 1200 - Screen.TwipsPerPixelX * 26 < 1500, "1500", txt1.Width - 1200 - Screen.TwipsPerPixelX * 26)
            End If
            strWidth = frmSelectChild.ShowSelectChild(Me, CurPoint.x * Screen.TwipsPerPixelX, CurPoint.y * Screen.TwipsPerPixelY, txt1.Width, Screen.TwipsPerPixelY * 300, rsTmp, strWidth)
            If Trim(strWidth) = "" Or Trim(strWidth) = ";;;;" Or Trim(strWidth) = ";;" Then
                Exit Sub
            End If
            '求出返回的参数
            txt1.Text = Split(strWidth, ";")(2)
            If IsNumeric(Split(strWidth, ";")(0)) Then
                If Opt(0).Value = True Then '从疾病编码
                    mlngID1 = CLng(Trim(Split(strWidth, ";")(0)))
                Else
                    mlngID2 = CLng(Trim(Split(strWidth, ";")(0)))
                End If
            End If
        ElseIf rsTmp.RecordCount = 1 Then
            txt1.Text = zlCommFun.Nvl(rsTmp!名称)
            If Opt(0).Value = True Then '从疾病编码
                mlngID1 = zlCommFun.Nvl(rsTmp!ID, 0)
            Else
                mlngID2 = zlCommFun.Nvl(rsTmp!ID, 0)
            End If
        End If
    Else
        If InStr(LAWLChar, Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
            Beep
            Beep
            Beep
        End If
    End If
End Sub

Private Sub txt1_LostFocus()
    Dim strTmp As String
    strTmp = txt1.Text
    For i = 1 To Len(LAWLChar)
        strTmp = Replace(strTmp, Mid(LAWLChar, i, 1), "")
    Next
    txt1.Text = strTmp
    zlCommFun.OpenIme
End Sub

Private Sub txt1_Change()
    mstrTxt1 = txt1.Text
End Sub
