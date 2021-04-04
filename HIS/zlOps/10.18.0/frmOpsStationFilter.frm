VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOpsStationFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤条件"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5550
   Icon            =   "frmOpsStationFilter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4305
      TabIndex        =   2
      Top             =   555
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4305
      TabIndex        =   1
      Top             =   135
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   4305
      TabIndex        =   0
      Top             =   1305
      Width           =   1100
   End
   Begin VB.Frame Frame3 
      Caption         =   "时间范围"
      Height          =   2040
      Left            =   60
      TabIndex        =   3
      Top             =   30
      Width           =   4110
      Begin VB.CheckBox chk 
         Caption         =   "按完成时间查询(&1)"
         Height          =   195
         Index           =   0
         Left            =   1215
         TabIndex        =   24
         Top             =   1725
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   1200
         TabIndex        =   4
         Top             =   300
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   114425859
         CurrentDate     =   38083
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   1
         Left            =   1200
         TabIndex        =   5
         Top             =   645
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   114425859
         CurrentDate     =   38083
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   2
         Left            =   1200
         TabIndex        =   20
         Top             =   1005
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   114425859
         CurrentDate     =   38083
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   3
         Left            =   1200
         TabIndex        =   21
         Top             =   1350
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   114425859
         CurrentDate     =   38083
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "～"
         Height          =   180
         Index           =   7
         Left            =   1005
         TabIndex        =   23
         Top             =   1380
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "已完手术(&F)"
         Height          =   180
         Index           =   5
         Left            =   150
         TabIndex        =   22
         Top             =   1050
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "等待手术(&2)"
         Height          =   180
         Index           =   8
         Left            =   150
         TabIndex        =   7
         Top             =   360
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "～"
         Height          =   180
         Index           =   9
         Left            =   1005
         TabIndex        =   6
         Top             =   690
         Width           =   180
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2340
      Left            =   60
      TabIndex        =   8
      Top             =   1995
      Width           =   4110
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   1155
         TabIndex        =   26
         Top             =   1935
         Width           =   2700
      End
      Begin VB.CommandButton cmd 
         Height          =   300
         Index           =   0
         Left            =   3540
         Picture         =   "frmOpsStationFilter.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1560
         Width           =   300
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1155
         TabIndex        =   12
         Top             =   195
         Width           =   2700
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1155
         TabIndex        =   11
         Top             =   540
         Width           =   2700
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   1155
         TabIndex        =   10
         Top             =   1230
         Width           =   2700
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   4
         Left            =   1155
         TabIndex        =   9
         Top             =   885
         Width           =   2700
      End
      Begin VB.TextBox txt 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   3
         Left            =   1155
         TabIndex        =   13
         Top             =   1575
         Width           =   2370
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "手术间(&9)"
         Height          =   180
         Index           =   6
         Left            =   210
         TabIndex        =   25
         Top             =   1995
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "姓  名(&4)"
         Height          =   180
         Index           =   0
         Left            =   210
         TabIndex        =   18
         Top             =   270
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "住院号(&5)"
         Height          =   180
         Index           =   1
         Left            =   210
         TabIndex        =   17
         Top             =   600
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "门诊号(&7)"
         Height          =   180
         Index           =   2
         Left            =   210
         TabIndex        =   16
         Top             =   1275
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "床  号(&6)"
         Height          =   180
         Index           =   4
         Left            =   210
         TabIndex        =   15
         Top             =   945
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "手  术(&8)"
         Height          =   180
         Index           =   3
         Left            =   210
         TabIndex        =   14
         Top             =   1620
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmOpsStationFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'窗体级变量定义
'######################################################################################################################

Private Type Items
    手术名称 As String
End Type
Private mlngDeptKey As Long
Private mblnOK As Boolean
Private mstrCondition As String
Private usrSaveItem As Items

'自定义过程或函数
'######################################################################################################################

Public Function ShowSearch(ByVal frmMain As Form, ByRef strCondition As String, ByVal lngDeptKey As Long) As Boolean
    '******************************************************************************************************************
    '
    '
    '
    '******************************************************************************************************************
    
    Dim rs As New ADODB.Recordset
    
    mblnOK = False
    mlngDeptKey = lngDeptKey
    
    mstrCondition = strCondition
    'mstrCondition格式：开始时间;结束时间;姓名;住院号;床号;门诊号;手术名称;手术id;开始时间;结束时间;完成时间标志
    
    cbo(0).Clear
    gstrSQL = "SELECT RowNum As ID,执行间 As 名称 FROM 医技执行房间 WHERE 科室id=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngDeptKey)
    If rs.BOF = False Then Call AddComboData(cbo(0), rs)
    
    dtp(0).Value = Format(Split(mstrCondition, ";")(0), dtp(0).CustomFormat)
    dtp(1).Value = Format(Split(mstrCondition, ";")(1), dtp(1).CustomFormat)

    txt(0).Text = Split(mstrCondition, ";")(2)
    txt(1).Text = Split(mstrCondition, ";")(3)
    txt(4).Text = Split(mstrCondition, ";")(4)
    txt(2).Text = Split(mstrCondition, ";")(5)
    txt(3).Text = Split(mstrCondition, ";")(6)
    cmd(0).Tag = Val(Split(mstrCondition, ";")(7))

    dtp(2).Value = Format(Split(mstrCondition, ";")(8), dtp(2).CustomFormat)
    dtp(3).Value = Format(Split(mstrCondition, ";")(9), dtp(3).CustomFormat)
    chk(0).Value = Val(Split(mstrCondition, ";")(10))
    cbo(0).Text = Trim(Split(mstrCondition, ";")(11))
        
    txt(3).Tag = ""
    
    Me.Show 1, frmMain
    
    strCondition = mstrCondition
    ShowSearch = mblnOK
End Function

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub cbo_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(cbo(Index).Text, 0)
End Sub

'控件或窗体事件
'######################################################################################################################

Private Sub cmd_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset

    Select Case Index
    '------------------------------------------------------------------------------------------------------------------
    Case 0
        gstrSQL = GetPublicSQL(SQL.手术项目选择)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If ShowPubSelect(Me, txt(3), 2, "编码,1200,0,;名称,2700,0,", Me.Name & "\手术项目选择", "请从下表中选择一个手术项目", rsData, rs, 8790, 4500, , Val(cmd(0).Tag)) = 1 Then
            If Val(cmd(0).Tag) <> zlCommFun.NVL(rs("ID").Value) Then
                txt(3).Text = AppendCode(zlCommFun.NVL(rs("名称").Value), zlCommFun.NVL(rs("编码").Value))
                cmd(0).Tag = zlCommFun.NVL(rs("ID").Value)
                txt(3).Tag = ""
                usrSaveItem.手术名称 = txt(3).Text
            End If
        End If
    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((ParamInfo.系统号) / 100))
End Sub

Private Sub cmdOK_Click()
    
    mstrCondition = Format(dtp(0).Value, dtp(0).CustomFormat) & ";" & Format(dtp(1).Value, dtp(1).CustomFormat) & ";" & _
                    txt(0).Text & ";" & txt(1).Text & ";" & txt(4).Text & ";" & txt(2).Text & ";" & txt(3).Text & ";" & _
                    IIf(Val(cmd(0).Tag) = 0, "", cmd(0).Tag) & ";" & _
                    Format(dtp(2).Value, dtp(2).CustomFormat) & ";" & Format(dtp(3).Value, dtp(3).CustomFormat) & ";" & chk(0).Value & ";" & cbo(0).Text
    mblnOK = True
    
    Unload Me
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt_Change(Index As Integer)
    
    Select Case Index
    Case 3
        txt(Index).Tag = "Changed"
    End Select
    
End Sub

Private Sub txt_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txt(Index))
    
    Select Case Index
    Case 0, 3
        zlCommFun.OpenIme True
    End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strText As String
    Dim strTmp As String
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim bytMode As Byte
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Select Case Index
        Case 3
            If txt(Index).Tag <> "" Then
                txt(Index).Tag = ""
                
                strText = UCase(txt(Index).Text)
                bytMode = GetApplyMode(strText)

                strText = strText & "%"
                strTmp = IIf(ParamInfo.项目输入匹配方式 = 1, "", "%") & strText

                gstrSQL = GetPublicSQL(SQL.手术项目过滤, bytMode)
                
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strText, strTmp)
                If ShowPubSelect(Me, txt(Index), 2, "编码,1200,0,1;名称,2700,0,0", Me.Name & "\手术项目过滤", "请从下面选择一个手术项目", rsData, rs, , , , Val(cmd(0).Tag)) = 1 Then
                    If Val(cmd(0).Tag) <> zlCommFun.NVL(rs("ID")) Then
            
                        txt(Index).Text = AppendCode(zlCommFun.NVL(rs("名称")), zlCommFun.NVL(rs("编码")))
                        cmd(0).Tag = zlCommFun.NVL(rs("ID"))
                        txt(Index).Tag = ""
                        
                        usrSaveItem.手术名称 = txt(Index).Text
                        
                    End If
                Else
                    txt(Index).Text = usrSaveItem.手术名称
                    txt(Index).Tag = ""
                    Exit Sub
                End If

            End If
        Case Else
            zlCommFun.PressKey vbKeyTab
        End Select
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 0, 3
        zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    If Cancel Then Exit Sub

    Select Case Index
    Case 3
        If (txt(Index).Tag = "Changed") Then
            txt(Index).Text = usrSaveItem.手术名称
            txt(Index).Tag = ""
        End If
    End Select
End Sub
