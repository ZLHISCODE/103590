VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHandBackSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5055
   Icon            =   "frmHandBackSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2160
      TabIndex        =   19
      Top             =   3120
      Width           =   1100
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3480
      TabIndex        =   18
      Top             =   3120
      Width           =   1100
   End
   Begin VB.Frame fraCondition 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.TextBox txtInput 
         Height          =   300
         Index           =   2
         Left            =   840
         TabIndex        =   16
         ToolTipText     =   "输入生产商编码、简码或名称"
         Top             =   2280
         Width           =   3480
      End
      Begin VB.CommandButton CmdSelecter 
         Caption         =   "…"
         Height          =   300
         Index           =   2
         Left            =   4320
         TabIndex        =   15
         Top             =   2280
         Width           =   255
      End
      Begin VB.TextBox txtInput 
         Height          =   300
         Index           =   1
         Left            =   840
         TabIndex        =   13
         ToolTipText     =   "输入供应商编码、简码或名称"
         Top             =   1800
         Width           =   3480
      End
      Begin VB.CommandButton CmdSelecter 
         Caption         =   "…"
         Height          =   300
         Index           =   1
         Left            =   4320
         TabIndex        =   12
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton CmdSelecter 
         Caption         =   "…"
         Height          =   300
         Index           =   0
         Left            =   4320
         TabIndex        =   9
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox txt结束NO 
         Height          =   300
         Left            =   2970
         MaxLength       =   8
         TabIndex        =   2
         Top             =   360
         Width           =   1605
      End
      Begin VB.TextBox txt开始No 
         Height          =   300
         Left            =   840
         MaxLength       =   8
         TabIndex        =   1
         Top             =   360
         Width           =   1605
      End
      Begin MSComCtl2.DTPicker dtp开始时间 
         Height          =   315
         Left            =   840
         TabIndex        =   5
         Top             =   840
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   166658051
         CurrentDate     =   36263
      End
      Begin MSComCtl2.DTPicker dtp结束时间 
         Height          =   315
         Left            =   2970
         TabIndex        =   6
         Top             =   840
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   166658051
         CurrentDate     =   36263
      End
      Begin VB.TextBox txtInput 
         Height          =   300
         Index           =   0
         Left            =   840
         TabIndex        =   10
         ToolTipText     =   "输入药品编码、简码或名称"
         Top             =   1320
         Width           =   3480
      End
      Begin VB.Label lblInputTxt 
         AutoSize        =   -1  'True
         Caption         =   "生产商"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   2340
         Width           =   540
      End
      Begin VB.Label lblInputTxt 
         AutoSize        =   -1  'True
         Caption         =   "供应商"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   1860
         Width           =   540
      End
      Begin VB.Label lblInputTxt 
         AutoSize        =   -1  'True
         Caption         =   "药品"
         Height          =   180
         Index           =   0
         Left            =   300
         TabIndex        =   11
         Top             =   1380
         Width           =   360
      End
      Begin VB.Label lbl至 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "～"
         Height          =   180
         Index           =   0
         Left            =   2640
         TabIndex        =   8
         Top             =   900
         Width           =   180
      End
      Begin VB.Label lbl时间 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "日期"
         Height          =   180
         Left            =   300
         TabIndex        =   7
         Top             =   900
         Width           =   360
      End
      Begin VB.Label lbl至 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "～"
         Height          =   180
         Index           =   1
         Left            =   2640
         TabIndex        =   4
         Top             =   420
         Width           =   180
      End
      Begin VB.Label LblNO 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "No"
         Height          =   180
         Left            =   480
         TabIndex        =   3
         Top             =   420
         Width           =   180
      End
   End
End
Attribute VB_Name = "frmHandBackSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngMode As Long            '0-未审核;1-已审核
Private mfrmMain As Form            '父窗体
Private mblnChange As Boolean
Private mlng库房ID As Long

Private Enum InputType
    药品 = 0
    供应商 = 1
    生产商 = 2
End Enum

Private Type Type_SQLCondition
    strNO开始 As String
    strNO结束 As String
    str填制时间开始 As String
    str填制时间结束 As String
    str审核时间开始 As String
    str审核时间结束 As String
    lng药品ID As Long
    lng供应商ID As Long
    str生产商 As String
End Type

Private SQLCondition As Type_SQLCondition
Private Sub CmdSelecter_Click(Index As Integer)
    Dim RecReturn As ADODB.Recordset
    
    If Index = InputType.药品 Then
        
        Call SetSelectorRS(1, "药品外购入库管理", mlng库房ID, mlng库房ID, , , , True)
        
'        Set RecReturn = Frm药品选择器.ShowME(Me, 1, 0, mlng库房ID, mlng库房ID)
        Set RecReturn = frmSelector.showMe(Me, 0, 1, , , , mlng库房ID, mlng库房ID, mlng库房ID, , , , , 2, False)
        
        If RecReturn.RecordCount = 0 Then
            Call zlControl.TxtSelAll(txtInput(Index))
            Exit Sub
        End If
            
        If gint药品名称显示 = 1 Then
            txtInput(Index).Text = "[" & RecReturn!药品编码 & "]" & IIf(IsNull(RecReturn!商品名), RecReturn!通用名, RecReturn!商品名)
        Else
            txtInput(Index).Text = "[" & RecReturn!药品编码 & "]" & RecReturn!通用名
        End If
        txtInput(Index).Tag = RecReturn!药品id
    Else
        If GetTxtInputReturn(Index, txtInput(Index), "") = False Then
            Call zlControl.TxtSelAll(txtInput(Index))
        End If
    End If
End Sub


Private Sub Cmd取消_Click()
    Unload Me
End Sub

Private Sub Cmd确定_Click()
    If Len(txt开始No.Text) = 8 Then
        SQLCondition.strNO开始 = txt开始No.Text
    End If
    
    If Len(txt结束NO.Text) = 8 Then
        SQLCondition.strNO结束 = txt结束NO.Text
    End If

    If mlngMode = 0 Then
        SQLCondition.str填制时间开始 = Format(dtp开始时间.Value, "yyyy-mm-dd") & " 00:00:00"
        SQLCondition.str填制时间结束 = Format(dtp结束时间.Value, "yyyy-mm-dd") & " 23:59:59"
    Else
        SQLCondition.str审核时间开始 = Format(dtp开始时间.Value, "yyyy-mm-dd") & " 00:00:00"
        SQLCondition.str审核时间结束 = Format(dtp结束时间.Value, "yyyy-mm-dd") & " 23:59:59"
    End If
    
    SQLCondition.lng药品ID = Val(txtInput(InputType.药品).Tag)
    SQLCondition.lng供应商ID = Val(txtInput(InputType.供应商).Tag)
    SQLCondition.str生产商 = txtInput(InputType.生产商).Text
    
    mblnChange = True
    
    Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call ReleaseSelectorRS
End Sub

Private Sub txtInput_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtInput(Index)
End Sub

Public Function GetSearch(ByVal FrmMain As Form, ByVal lngMode As Long, ByVal lng库房ID As Long, _
        ByRef strNO开始 As String, _
        ByRef strNO结束 As String, _
        ByRef str填制时间开始 As String, _
        ByRef str填制时间结束 As String, _
        ByRef str审核时间开始 As String, _
        ByRef str审核时间结束 As String, _
        ByRef lng药品ID As Long, _
        ByRef lng供应商ID As Long, _
        ByRef str生产商 As String) As Boolean
    
    mblnChange = False
    mlngMode = lngMode
    mlng库房ID = lng库房ID
    Set mfrmMain = FrmMain
    
    If lngMode = 0 Then
        dtp开始时间.Value = CDate(str填制时间开始)
        dtp结束时间.Value = CDate(str填制时间结束)
    Else
        dtp开始时间.Value = CDate(str审核时间开始)
        dtp结束时间.Value = CDate(str审核时间结束)
    End If
    
    Me.Show vbModal, mfrmMain
    
    GetSearch = mblnChange
    
    strNO开始 = SQLCondition.strNO开始
    strNO结束 = SQLCondition.strNO结束
    
    If lngMode = 0 Then
        str填制时间开始 = SQLCondition.str填制时间开始
        str填制时间结束 = SQLCondition.str填制时间结束
    Else
        str审核时间开始 = SQLCondition.str审核时间开始
        str审核时间结束 = SQLCondition.str审核时间结束
    End If
    
    lng药品ID = SQLCondition.lng药品ID
    lng供应商ID = SQLCondition.lng供应商ID
    str生产商 = SQLCondition.str生产商
End Function
Private Sub txtInput_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtInput(Index).Text) = "" Then Exit Sub
    
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim RecReturn As Recordset
    Dim strkey As String
    
    If Index = InputType.药品 Then
        If KeyCode <> vbKeyReturn Then Exit Sub
        If Trim(txtInput(Index).Text) = "" Then Exit Sub
        sngLeft = Me.Left + fraCondition.Left + txtInput(Index).Left
        sngTop = Me.Top + fraCondition.Top + txtInput(Index).Top + txtInput(Index).Height + Me.Height - Me.ScaleHeight '  50
        If sngTop + 3630 > Screen.Height Then
            sngTop = sngTop - txtInput(Index).Height - 3630
        End If
        
        strkey = Trim(txtInput(Index).Text)
        If Mid(strkey, 1, 1) = "[" Then
            If InStr(2, strkey, "]") <> 0 Then
                strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
            Else
                strkey = Mid(strkey, 2)
            End If
        End If
        
        Call SetSelectorRS(1, "药品外购入库管理", mlng库房ID, mlng库房ID, , , , True)
        
'        Set RecReturn = Frm药品多选选择器.ShowME(Me, 1, , mlng库房ID, mlng库房ID, strkey, sngLeft, sngTop)
        Set RecReturn = frmSelector.showMe(Me, 1, 1, strkey, sngLeft, sngTop, mlng库房ID, mlng库房ID, mlng库房ID, , , , , 2, False)
        
        If RecReturn.RecordCount = 0 Then
            Call zlControl.TxtSelAll(txtInput(Index))
            Exit Sub
        End If
        
        If gint药品名称显示 = 1 Then
            txtInput(Index).Text = "[" & RecReturn!药品编码 & "]" & IIf(IsNull(RecReturn!商品名), RecReturn!通用名, RecReturn!商品名)
        Else
            txtInput(Index).Text = "[" & RecReturn!药品编码 & "]" & RecReturn!通用名
        End If
        txtInput(Index).Tag = RecReturn!药品id
    Else
        If GetTxtInputReturn(Index, txtInput(Index), Trim(txtInput(Index).Text)) = False Then
            Call zlControl.TxtSelAll(txtInput(Index))
        End If
    End If
End Sub

Private Function GetTxtInputReturn(ByVal intType As Integer, ByVal txtObj As TextBox, ByVal strkey As String) As Boolean
    Dim vRect As RECT
    Dim lngH As Long
    Dim strReturn As String
    
    vRect = zlControl.GetControlRect(txtObj.hWnd)
    lngH = txtObj.Height
    vRect.Left = vRect.Left - 15
    
    strReturn = SelectInput(intType, Trim(strkey), vRect.Left, vRect.Top, lngH)
    
    If strReturn = "" Then Exit Function
        
    txtObj.Tag = Val(Split(strReturn, ";")(0))
    txtObj.Text = Split(strReturn, ";")(1)
    
    GetTxtInputReturn = True
End Function

Private Function SelectInput(ByVal intType As Integer, ByVal strkey As String, ByVal sngX As Single, ByVal sngY As Single, ByVal sngH As Single) As String
    '选择器：支持对药品、供应商、生产商的选择
    'intType：0-药品;1-供应商;2-生产商
    'strKey：空-全部;非空-模糊匹配
    'SelectInput返回值：空-没找到匹配记录;
    '                 非空-药品（药品ID;药品名称;规格;单位;包装）
    '                     -供应商（供应商ID;供应商名称）
    '                     -生产商（生产商ID;生产商名称）
    
    Dim rsTemp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim strSubUnit As String
    Dim strFindString As String
    Dim strReturn As String
    Dim strSql药品 As String
    
    Err = 0: On Error GoTo ErrHand:
    
    strkey = UCase(Trim(strkey))
    
    Select Case intType
    Case InputType.药品
        If strkey <> "" Then
            strFindString = " And (B.编码 Like [1] OR B.名称 Like [2] OR C.简码 LIKE [2])"
            If IsNumeric(strkey) Then                         '10,11.输入全是数字时只匹配编码'对于药品,则要匹配简码(码类为3的数字码)
                If Mid(gtype_UserSysParms.P44_输入匹配, 1, 1) = "1" Then strFindString = " And (B.编码 Like [1] Or B.简码 Like [2] And C.码类=3)"
            ElseIf zlStr.IsCharAlpha(strkey) Then         '01,11.输入全是字母时只匹配简码
                If Mid(gtype_UserSysParms.P44_输入匹配, 2, 1) = "1" Then strFindString = " And C.简码 Like [2] "
            ElseIf zlStr.IsCharChinese(strkey) Then
                strFindString = " And B.名称 Like [2] "
            End If
        End If
        
        If strkey = "" Then
            If gint药品名称显示 = 0 Then
                strSql药品 = ",'['||编码||']'|| 通用名 As 药品名称"
            ElseIf gint药品名称显示 = 1 Then
                strSql药品 = ",'['||编码||']'|| Nvl(商品名,通用名) As 药品名称"
            ElseIf gint药品名称显示 = 2 Then
                strSql药品 = ",'['||编码||']'|| 通用名 As 药品名称,商品名"
            End If
        Else
            strSql药品 = ",'['||编码||']'|| 输入名称 As 药品名称"
        End If
        
        gstrSQL = "Select Rownum As ID, 药品id" & strSql药品 & ",规格,产地 as 生产商,商品名 " & _
            " From (Select Distinct A.药品id, B.编码, B.输入名称, B.名称 As 通用名,C.名称 As 商品名, B.规格,B.产地 " & _
            " From 药品规格 A, " & _
            " (Select B.ID, B.编码, B.名称, B.规格,B.产地, C.名称 As 输入名称 From 收费项目目录 B, 收费项目别名 C " & _
            " Where (B.站点 = [3] Or B.站点 is Null) And B.ID = C.收费细目id And B.类别 In ('5', '6', '7') " & strFindString & ") B, 收费项目别名 C " & _
            " Where A.药品id = B.ID And A.药品id = C.收费细目id(+) And C.性质(+) = 3 "

        gstrSQL = gstrSQL & " Order By B.编码)"
        
        Set rsTemp = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "药品选择器", False, "", "选择药品", False, False, True, sngX, sngY, sngH, blnCancel, False, False, _
                        strkey & "%", "%" & strkey & "%", _
                        gstrNodeNo)
        
        If blnCancel = True Then Exit Function
        
        If rsTemp Is Nothing Then
            strReturn = ""
        Else
            strReturn = rsTemp!药品id & ";" & rsTemp!药品名称
        End If
    Case InputType.供应商
        gstrSQL = "Select id,名称,编码,简码 From 供应商 " & _
                  "Where (站点 = [3] Or 站点 is Null) " & _
                  "  And (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null ) And 末级=1 " & _
                  "  And (substr(类型,1,1)=1 Or Nvl(末级,0)=0) And (编码 like [1] or 简码 like [2] or 名称 like [2])"
        Set rsTemp = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "供应商选择器", False, "", "选择供应商", False, False, True, sngX, sngY, sngH, blnCancel, False, False, _
                        strkey & "%", "%" & strkey & "%", _
                        gstrNodeNo)
        
        If blnCancel = True Then Exit Function
        
        If rsTemp Is Nothing Then
            MsgBox "输入值无效！", vbInformation, gstrSysName
            strReturn = ""
        Else
            strReturn = rsTemp!id & ";" & rsTemp!名称
        End If
    Case InputType.生产商
        gstrSQL = "Select Rownum As ID,名称,编码,简码 From 药品生产商 " & _
                  "Where (站点 = [3] Or 站点 is Null) And (编码 like [1] Or 简码 like [2] Or 名称 like [2]) " & _
                  "Order By 编码"
        Set rsTemp = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "生产商选择器", False, "", "选择生产商", False, False, True, sngX, sngY, sngH, blnCancel, False, False, _
                        strkey & "%", "%" & strkey & "%", _
                        gstrNodeNo)
        
        If blnCancel = True Then Exit Function
        
        If rsTemp Is Nothing Then
            MsgBox "输入值无效！", vbInformation, gstrSysName
            strReturn = ""
        Else
            strReturn = rsTemp!id & ";" & rsTemp!名称
        End If
    End Select
    
    SelectInput = strReturn
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function



Private Sub txt结束NO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Call txt结束NO_Validate(True)
End Sub
Private Sub txt结束NO_Validate(Cancel As Boolean)
    If IsNumeric(txt结束NO.Text) Then
        txt结束NO.Text = zlCommFun.GetFullNO(txt结束NO.Text, 92)
    End If
End Sub


Private Sub txt开始No_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Call txt开始No_Validate(True)
End Sub
Private Sub txt开始No_Validate(Cancel As Boolean)
    If IsNumeric(txt开始No.Text) Then
        txt开始No.Text = zlCommFun.GetFullNO(txt开始No.Text, 92)
    End If
End Sub


