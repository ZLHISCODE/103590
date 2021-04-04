VERSION 5.00
Begin VB.Form frmUfgColsList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "数据列表配置"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUfgColsList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&S)"
      Height          =   405
      Left            =   1440
      TabIndex        =   4
      Top             =   3960
      Width           =   825
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&E)"
      Height          =   405
      Left            =   2400
      TabIndex        =   3
      Top             =   3960
      Width           =   820
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "恢复默认(&D)"
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   1185
   End
   Begin VB.ListBox lstUfgColsName 
      Height          =   3435
      ItemData        =   "frmUfgColsList.frx":6852
      Left            =   120
      List            =   "frmUfgColsList.frx":6854
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "勾选需要显示的列"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   1680
   End
End
Attribute VB_Name = "frmUfgColsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mUcData As ucFlexGrid
Private mStrDefaultColNames As String



Public Sub ShowUfgColsListWindow(ByRef UcData As ucFlexGrid, ByVal strDefaultColNames As String)
'打开列名列表窗体并加载默认显示的列
    
    '将UcData对象存入模块级变量
    Set mUcData = UcData
    mStrDefaultColNames = strDefaultColNames
    
    '调用加载默认显示的列
    Call LoadColsList(UcData)
    
    '加载窗体
    Call Show(1)
End Sub



Private Sub cmdOK_Click()
'确定保存属性操作
On Error GoTo ErrHandle

    Dim i As Integer
    Dim strColsName As String
    Dim strColName  As String
    Dim strProperty As String
    Dim objProperty As Scripting.Dictionary
    
    '按钮是否被点击
    cmdOK.Tag = True
    
    If cmdDefault.Tag = "True" Then mUcData.ColNames = mStrDefaultColNames
    
    For i = 0 To lstUfgColsName.ListCount - 1
        '判断是否选中某列
        If lstUfgColsName.Selected(i) Then
            strColsName = strColsName + lstUfgColsName.list(i) & ","
        End If
    Next

    '把hide属性写入flexcpdata
    For i = 1 To mUcData.DataGrid.Cols - 1
        strColName = mUcData.DataGrid.Cell(flexcpText, 0, i)
        
        Set objProperty = mUcData.DataGrid.Cell(flexcpData, 0, i)
        
        If Not objProperty Is Nothing Then
            strProperty = Mid(objProperty(TColPro.cpProperty), InStrRev(objProperty(TColPro.cpProperty), "@") + 1)
            
            '判断匹配选中列 匹配则删除hide属性  未匹配就添加hide属性
            If InStr(strColsName, strColName) = 0 Then
            
                '如果属性字符串中存在 uncfg 属性就跳过
                If InStr(strProperty, "uncfg") = 0 Then
                    '如果属性字符串中已有 hide 属性就跳过  没有则追加
                    If InStr(strProperty, "hide") Then
                        
                        objProperty(TColPro.cpProperty) = mUcData.GetFieldName(i) & "@" & strProperty
                        
                    Else
                        '添加隐藏属性后 保存列的属性字符串
                        objProperty(TColPro.cpProperty) = mUcData.GetFieldName(i) & "@" & strProperty & ",hide"
                    End If
                    
                    '隐藏列
                     mUcData.DataGrid.ColHidden(i) = True
                     mUcData.DataGrid.Cell(flexcpData, 0, i)(TColPro.cpIsHide) = True
                     '调用计算CheckBox位置方法
                     mUcData.RefreshCbxPostion
                End If
                
            Else
            
                 '如果属性字符串中存在 uncfg 属性就跳过
                If InStr(strProperty, "uncfg") = 0 Then
                    '如果属性字符串中已有 hide 属性就去掉  没有则忽略
                    If InStr(strProperty, "hide") Then
                        
                        '将hide属性删除 并保存列的属性字符串
                        strProperty = Replace(strProperty, ",hide", "")
                        objProperty(TColPro.cpProperty) = mUcData.GetFieldName(i) & "@" & strProperty
                    Else
                    
                        objProperty(TColPro.cpProperty) = mUcData.GetFieldName(i) & "@" & strProperty
                        
                    End If
                    
                     '显示列
                     mUcData.DataGrid.ColHidden(i) = False
                     mUcData.DataGrid.Cell(flexcpData, 0, i)(TColPro.cpIsHide) = False
                     '调用计算CheckBox位置方法
                     mUcData.RefreshCbxPostion
                End If
        
            End If
        End If
    Next
    
    Call Me.Hide
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub LoadColsList(UcData As ucFlexGrid)
'加载默认显示的列
    Dim i As Integer
    Dim j As Integer
    Dim strProperty As String
    Dim objProperty As Scripting.Dictionary
    
    '给控件下标赋初始值
    j = 0
    '清空List控件
    lstUfgColsName.Clear
    
    For i = 1 To UcData.DataGrid.Cols - 1
        Set objProperty = mUcData.DataGrid.Cell(flexcpData, 0, i)
        
        If Not objProperty Is Nothing Then
            strProperty = Mid(objProperty(TColPro.cpProperty), InStrRev(objProperty(TColPro.cpProperty), "@") + 1)
            
            If InStr(strProperty, "uncfg") = 0 Then
                '加载字符
                frmUfgColsList.lstUfgColsName.list(j) = UcData.DataGrid.Cell(flexcpText, 0, i)
                
                '判断是否默认为隐藏列
                If InStr(strProperty, "hide") = 0 Then
                    '加载默认显示列
                    frmUfgColsList.lstUfgColsName.Selected(j) = True
                End If
                
                j = j + 1
            End If
        End If
    Next

End Sub


Private Sub cmdDefault_Click()
'恢复默认勾选
On Error GoTo ErrHandle
    Dim i As Integer
    Dim j As Integer
    Dim strProperty As String
    Dim strTemp As String
    Dim strColNames() As String
    
    '按钮是否被点击
    cmdDefault.Tag = True
    
    '将列名配置串分离并存入数组
    strColNames() = Split(mStrDefaultColNames, "|")
    
     '清空List控件
    lstUfgColsName.Clear
    
    For i = 1 To UBound(strColNames()) - 1
        strProperty = strColNames(i)
        
        If InStr(strProperty, "uncfg") = 0 Then
            
            '判断字符串是否包含 “>” “,”符号，需要进行截取操作
             If InStr(strProperty, ">") > 0 Then
                strTemp = Mid(strProperty, 1, InStr(strProperty, ">") - 1)

                If InStr(strTemp, ",") > 0 Then
                    strTemp = Mid(strTemp, 1, InStr(strTemp, ",") - 1)
                Else
                    strTemp = strTemp
                End If
             Else
                If InStr(strProperty, ",") > 0 Then
                    strTemp = Mid(strProperty, 1, InStr(strProperty, ",") - 1)
                Else
                    strTemp = strProperty
                End If
             End If

            '加载字符
            frmUfgColsList.lstUfgColsName.list(j) = strTemp
            
            '判断是否默认为隐藏列
            If InStr(strProperty, "hide") = 0 Then
                '加载默认显示列
                frmUfgColsList.lstUfgColsName.Selected(j) = True

            End If
            
            j = j + 1
        End If
    Next

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdExit_Click()
'卸载窗体
On Error GoTo ErrHandle

    Unload Me
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
'将窗口置顶
    '如果默认列名串等于空，则禁用恢复默认按钮
    If Trim(mStrDefaultColNames) = "" Then cmdDefault.Enabled = False
    '将窗口置顶
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3

End Sub
