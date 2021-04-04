VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmBedLevel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "更改床位等级"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   Icon            =   "frmBedLevel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   435
      TabIndex        =   11
      Top             =   2445
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   12
      Top             =   15
      Width           =   5460
      Begin VB.TextBox txt科室 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   630
         Width           =   1830
      End
      Begin VB.TextBox txt姓名 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   1515
      End
      Begin VB.TextBox txt性别 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3210
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   675
      End
      Begin VB.TextBox txt年龄 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   690
      End
      Begin VB.TextBox txt住院号 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   630
         Width           =   1515
      End
      Begin VB.TextBox txt床号 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1005
         Width           =   1515
      End
      Begin VB.ComboBox cboNew 
         Height          =   300
         Left            =   975
         TabIndex        =   8
         Text            =   "cboNew"
         Top             =   1770
         Width           =   4260
      End
      Begin VB.TextBox txtPre 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1395
         Width           =   4260
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   300
         Left            =   3420
         TabIndex        =   6
         Top             =   1005
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   19
         Format          =   "yyyy-MM-dd HH:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboLevel 
         Height          =   300
         Left            =   1095
         Style           =   2  'Dropdown List
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1770
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前科室"
         Height          =   180
         Left            =   2640
         TabIndex        =   21
         Top             =   690
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "住院号"
         Height          =   180
         Left            =   375
         TabIndex        =   20
         Top             =   690
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   540
         TabIndex        =   19
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   180
         Left            =   2760
         TabIndex        =   18
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Height          =   180
         Left            =   4125
         TabIndex        =   17
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "变动床位"
         Height          =   180
         Left            =   195
         TabIndex        =   16
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "新等级"
         Height          =   180
         Left            =   375
         TabIndex        =   14
         Top             =   1830
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "原等级"
         Height          =   180
         Left            =   375
         TabIndex        =   13
         Top             =   1455
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "生效时间"
         Height          =   180
         Left            =   2640
         TabIndex        =   15
         Top             =   1065
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4155
      TabIndex        =   10
      Top             =   2445
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2895
      TabIndex        =   9
      Top             =   2445
      Width           =   1100
   End
End
Attribute VB_Name = "frmBedLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mlng病人ID As Long
Public mlng主页ID As Long
Public mstr床号 As String

Private mrsPatiInfo As ADODB.Recordset
Private mrsBedLevel As ADODB.Recordset

Private Sub cboNew_GotFocus()
    zlControl.TxtSelAll cboNew
End Sub

Private Sub cboNew_KeyPress(KeyAscii As Integer)
    '69273:刘鹏飞,2014-01-03,快速定位床位等级
    Dim lngIdx As Long
    Dim i As Long, iCount As Integer
    Dim strText As String, strResult As String, strFilter As String
    Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
    Dim strCompents As String '匹配串
    Dim rsTemp As ADODB.Recordset
    
    If KeyAscii = 13 Then
        If cboNew.Locked Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        strText = UCase(cboNew.Text)
        If cboNew.ListIndex <> -1 Then
            '弹出列表时,又在文本框输入了内容
            If strText <> cboNew.List(cboNew.ListIndex) Then Call cbo.SetIndex(cboNew.hWnd, -1)
        End If
        If strText = "" Then
            cboNew.ListIndex = -1
        ElseIf cboNew.ListIndex = -1 Then
            strFilter = ""
            '先复制记录集
            Set rsTemp = zlDatabase.zlCopyDataStructure(mrsBedLevel)
            strCompents = Replace(gstrLike, "%", "*") & strText & "*"
            
            If IsNumeric(strText) Then
                intInputType = 0
            ElseIf zlCommFun.IsCharAlpha(strText) Then
                intInputType = 1
            Else
                intInputType = 2
            End If
            
            mrsBedLevel.Filter = strFilter: iCount = 0
            With mrsBedLevel
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not mrsBedLevel.EOF
                    Select Case intInputType
                    Case 0  '输入的是全数字
                        '如果输入的数字,需要检查:
                        '1.编号输入值相等,主要输入如:12 匹配000012这种情况
                        '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                        
                        '主要是检查输入的内容与编号完全相同,则直接就定位到该姓名
                        If Nvl(!编码) = strText Then strResult = Nvl(!名称): iCount = 0: Exit Do
                        
                        '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等.因此如果存在此种情况,需要弹出选择器供选择
                        If Val(Nvl(!编码)) = Val(strText) Then
                            If iCount = 0 Then strResult = Nvl(!名称)
                            iCount = iCount + 1
                        End If
                        
                        '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                         If Val(Nvl(!编码)) Like strText & "*" Then
                            If isCheckBedLevelExists(Nvl(!名称)) Then Call zlDatabase.zlInsertCurrRowData(mrsBedLevel, rsTemp)
                         End If
                    Case 1  '输入的是全字母
                        '规则:
                        ' 1.输入的简码相等,则直接定位
                        ' 2.根据参数来匹配相同数据
                        
                        '1.输入的简码相等,则直接定位
                        If Trim(Nvl(!简码)) = strText Then
                            If iCount = 0 Then strResult = Nvl(!名称)   '可能存在多个相同的多个
                            iCount = iCount + 1
                        End If
                        
                        '2.根据参数来匹配相同数据
                        If Trim(Nvl(!简码)) Like strCompents Then
                            If isCheckBedLevelExists(Nvl(!名称)) Then Call zlDatabase.zlInsertCurrRowData(mrsBedLevel, rsTemp)
                        End If
                    Case Else  ' 2-其他
                        '规则:可能存在汉字等情况,或编号类似于N001简码可能有ZYK01这种情况
                        '1.编码\简码相等,直接定位
                        '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                        
                        '1.编码\简码相等,直接定位
                        If Trim(!编码) = strText Or Trim(!简码) = strText Or Trim(!名称) = strText Then
                            If iCount = 0 Then strResult = Nvl(!名称)   '可能存在多个相同的多个
                            iCount = iCount + 1
                        End If
                        
                        '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                        If Trim(!编码) Like strText & "*" Or Trim(Nvl(!简码)) Like strCompents Or Trim(Nvl(!名称)) Like strCompents Then
                            If isCheckBedLevelExists(Nvl(!名称)) Then Call zlDatabase.zlInsertCurrRowData(mrsBedLevel, rsTemp)
                        End If
                    End Select
                    mrsBedLevel.MoveNext
                Loop
            End With
            If iCount > 1 Then strResult = ""
            If strResult = "" And rsTemp.RecordCount = 1 Then strResult = Nvl(rsTemp!名称)
            '直接定位
            If strResult <> "" Then
                rsTemp.Close: Set rsTemp = Nothing
                If isCheckBedLevelExists(strResult, True) Then cboNew.SetFocus:  zlCommFun.PressKey vbKeyTab
                Exit Sub
            End If
            
            '需要检查是否有多条满足条件的记录
            If rsTemp.RecordCount <> 0 Then
                '先按某种方式进行排序
                rsTemp.Sort = "简码,编码"
                '弹出选择器
                Dim rsReturn As ADODB.Recordset
                If zlDatabase.zlShowListSelect(Me, glngSys, 1130, cboNew, rsTemp, True, "", "", rsReturn) Then
                    If Not rsReturn Is Nothing Then
                        If rsReturn.RecordCount <> 0 Then
                            '进行定位
                            If isCheckBedLevelExists(Nvl(rsReturn!名称), True) Then
                                cboNew.SetFocus
                                zlCommFun.PressKey vbKeyTab
                                Exit Sub
                            End If
                        End If
                    End If
                Else
                    cboNew.SetFocus
                    Exit Sub
                End If
            Else
                '未找到
                rsTemp.Close: Set rsTemp = Nothing
                KeyAscii = 0: cboNew.ListIndex = -1: zlControl.TxtSelAll cboNew: Exit Sub
            End If
            rsTemp.Close: Set rsTemp = Nothing
        End If
        
        If cboNew.ListIndex = -1 Then
            cboNew.Text = ""
            Exit Sub
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Function isCheckBedLevelExists(ByVal str名称 As String, Optional blnLocateItem As Boolean = False, Optional ByVal blnLevel As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查名称是否在床位等级下拉列表中
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    If blnLevel = True Then
        For i = 0 To cboLevel.ListCount - 1
            If cboLevel.List(i) = str名称 Then
                If blnLocateItem Then cboNew.ListIndex = i
                isCheckBedLevelExists = True
                Exit Function
            End If
        Next
    Else
        For i = 0 To cboNew.ListCount - 1
            If cboNew.List(i) = str名称 Then
                If blnLocateItem Then cboNew.ListIndex = i
                isCheckBedLevelExists = True
                Exit Function
            End If
        Next
    End If
End Function

Private Sub cboNew_Validate(Cancel As Boolean)
    If isCheckBedLevelExists(cboNew.Text, True, False) = False Then
        cboNew.Text = ""
        cboNew.ListIndex = -1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, lngLevel As Long
    
    
    gblnOK = False
    Set mrsPatiInfo = GetPatiInfo(mlng病人ID, mlng主页ID)
    Set rsTmp = GetPatiBeds(mlng病人ID, mstr床号)
    
    With mrsPatiInfo
       txt姓名.Text = !姓名
       txt性别.Text = "" & !性别
       txt年龄.Text = "" & !年龄
       txt住院号.Text = "" & !住院号
       txt科室.Text = "" & !当前科室
       txt床号.Text = mstr床号
       txtPre.Text = rsTmp!床位等级
       lngLevel = Val("" & rsTmp!床位等级id)
    End With

    txtDate.Text = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    
    On Error GoTo errH
    '69273:刘鹏飞,2014-01-03,提供床位登记的快速查找
    gstrSQL = "Select ID,编码,名称,zlspellcode(名称,20) 简码 From 收费项目目录 Where 类别='J' And (撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or 撤档时间 is NULL) And ID<>[1] Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngLevel)
    cboNew.Clear
    cboLevel.Clear: cboLevel.Visible = False
    Set mrsBedLevel = rsTmp.Clone
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboNew.AddItem rsTmp!编码 & "-" & rsTmp!名称
            cboNew.ItemData(i - 1) = rsTmp!ID
            cboLevel.AddItem rsTmp!名称
            cboLevel.ItemData(i - 1) = rsTmp!ID
            rsTmp.MoveNext
        Next
        cboNew.ListIndex = 0
    Else
        MsgBox "不能读取床位等级数据,请先到床位等级管理中设置！", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If IsDate(txtDate.Text) Then Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtDate_LostFocus()
    If Not IsDate(txtDate.Text) Then txtDate.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim dMax As Date, strSQL As String
    Dim Curdate As Date
    
    If cboNew.ListIndex = -1 Then
        MsgBox "请选择新的床位等级！", vbInformation, gstrSysName
        cboNew.SetFocus: Exit Sub
    End If
    If Not IsDate(txtDate.Text) Then
        MsgBox "请输入合法的生效时间！", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    
    dMax = GetMaxDate(mlng病人ID, mlng主页ID)
    If CDate(txtDate.Text) <= dMax Then
        MsgBox "生效时间必须大于该病人上次变动时间 " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & " ！", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    
    '时间不能超过当前时间太长(一个月)
    Curdate = zlDatabase.Currentdate
    If CDate(txtDate.Text) > Curdate Then
        If CDate(txtDate.Text) - Curdate > 30 Then
            MsgBox "生效时间比当前时间大得过多,请检查！", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        If MsgBox("生效时间大于了当前系统时间,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txtDate.SetFocus: Exit Sub
        End If
    End If
        
    strSQL = "zl_病人变动记录_BedLevel(" & mlng病人ID & "," & mlng主页ID & ",'" & txt床号.Text & "'," & _
        cboNew.ItemData(cboNew.ListIndex) & ",To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS')," & _
        "'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
    
    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    gblnOK = True
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function ShowMe(frmParent As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str床号 As String) As Boolean
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mstr床号 = str床号
    
    Me.Show 1, frmParent
    ShowMe = gblnOK
End Function
