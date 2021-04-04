VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmExeEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "执行登记"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   Icon            =   "frmExeEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdSample 
      Caption         =   "样句(&M)"
      Height          =   350
      Left            =   1440
      TabIndex        =   5
      Top             =   2370
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2910
      TabIndex        =   3
      Top             =   2370
      Width           =   1100
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4050
      TabIndex        =   4
      Top             =   2370
      Width           =   1100
   End
   Begin VB.CommandButton cmdSetup 
      Caption         =   "票据设置(&S)"
      Height          =   350
      Left            =   105
      TabIndex        =   6
      Top             =   2370
      Width           =   1305
   End
   Begin MSComCtl2.DTPicker dtp执行时间 
      Height          =   315
      Left            =   3255
      TabIndex        =   2
      Top             =   1830
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   107216899
      CurrentDate     =   37447
   End
   Begin VB.ComboBox cbo执行人 
      Height          =   300
      Left            =   735
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1830
      Width           =   1395
   End
   Begin VB.TextBox txt结论 
      Height          =   1440
      Left            =   135
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   300
      Width           =   5160
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -105
      TabIndex        =   10
      Top             =   2130
      Width           =   5760
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "执行时间"
      Height          =   180
      Left            =   2505
      TabIndex        =   9
      Top             =   1890
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "执行人"
      Height          =   180
      Left            =   150
      TabIndex        =   8
      Top             =   1890
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "执行情况:"
      Height          =   180
      Left            =   165
      TabIndex        =   7
      Top             =   75
      Width           =   810
   End
   Begin VB.Menu mnuSample 
      Caption         =   "结论(&S)"
      Visible         =   0   'False
      Begin VB.Menu mnuSampleAdd 
         Caption         =   "保存当前结论(&S)"
      End
      Begin VB.Menu mnuSampleDel 
         Caption         =   "删除已有结论(&D)"
         Begin VB.Menu mnuSampleItemDel 
            Caption         =   "<无结论样句>"
            Index           =   0
         End
      End
      Begin VB.Menu mnuSample_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSampleItem 
         Caption         =   "<无结论样句>"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmExeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
'入
Public mlngDeptID As Long '当前执行科室
Public mblnView As Boolean
Public mstrDate As String
'入/出
Public mstrOper As String '如果已执行,则为执行人
Public mstrLog As String '当前记录结论
'出
Public mvDate As Date '执行时间

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSql As String
        
    If Not mblnView Then
        If Not zlcontrol.TxtCheckInput(txt结论, "结论", , True) Then Exit Sub
        
        If cbo执行人.ListIndex = -1 Then
            MsgBox "请确定执行登记人！", vbInformation, gstrSysName
            cbo执行人.SetFocus: Exit Sub
        End If
        
        mstrLog = txt结论.Text
        mstrOper = zlStr.NeedName(cbo执行人.Text)
        mvDate = dtp执行时间.Value
        
        gblnOK = True
    End If
    
    Unload Me
End Sub

Private Sub cmdSample_Click()
    If LoadSample Then
        PopupMenu mnuSample, 2, cmdSample.Left, cmdSample.Top + cmdSample.Height, mnuSampleAdd
    End If
End Sub

Private Sub cmdSetup_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1142", Me)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyCode = vbKeyEscape Then
        If mblnView Then Unload Me
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    gblnOK = False
        
    If Not mblnView Then
        Call InitOper
        txt结论.Text = mstrLog
        '新登记或修改都以当前人员身份,当前时间执行登记
        dtp执行时间.Value = zlDatabase.Currentdate
        cbo执行人.Enabled = Not gbln本人执行
        cbo执行人.ListIndex = cbo.FindIndex(cbo执行人, UserInfo.ID)
        If cbo执行人.ListIndex = -1 And Not cbo执行人.Enabled Then
            MsgBox "你不属于当前执行科室，且你不能以其它人员身份登记执行情况。", vbInformation, gstrSysName
            Unload Me: Exit Sub
        ElseIf cbo执行人.ListIndex = -1 Then
            cbo执行人.ListIndex = cbo.FindIndex(cbo执行人, mstrOper, True)
        End If
    Else
        Caption = "查看登记"
        
        txt结论.Text = mstrLog
        dtp执行时间.Value = Format(mstrDate, "yyyy-MM-dd hh:mm:ss")
        cbo执行人.AddItem mstrOper
        cbo执行人.ListIndex = cbo执行人.NewIndex
        
        cmdSetup.Visible = False
        cmdSample.Visible = False
        cmdCancel.Visible = False
        dtp执行时间.Enabled = False
        cbo执行人.Enabled = False
        txt结论.Locked = True
        
        cmdOK.Left = cmdOK.Left + cmdOK.Width / 2
    End If
End Sub

Private Sub InitOper()
    Dim strSql As String, i As Long
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSql = "Select Distinct A.ID,A.编号,A.姓名,A.简码" & _
        " From 人员表 A,部门人员 B" & _
        " Where A.ID=B.人员ID And B.部门ID=[1] And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
        " Order by A.简码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngDeptID)
    
    For i = 1 To rsTmp.RecordCount
        cbo执行人.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        cbo执行人.ItemData(cbo执行人.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Next
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnView = False
End Sub

Private Sub mnuSampleAdd_Click()
    Dim i As Integer, intCount As Integer
    
    If Trim(txt结论.Text) = "" Then
        MsgBox "请输入具体的内容！", vbInformation, gstrSysName
        txt结论.SetFocus
        Exit Sub
    End If
    
    For i = 0 To mnuSampleItem.UBound
        If mnuSampleItem(i).Tag = txt结论.Text Then
            MsgBox "该结论已经保存为了样句！", vbInformation, gstrSysName
            txt结论.SetFocus
            Exit Sub
        End If
    Next
        
    intCount = 0
    If mnuSampleItem(0).Caption <> "<无结论样句>" Then
        intCount = mnuSampleItem.UBound + 1
    End If
    
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\结论样句", "Count", intCount + 1)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\结论样句", "Item" & intCount, txt结论.Text)
End Sub

Private Sub mnuSampleItem_Click(Index As Integer)
    If mnuSampleItem(Index).Caption = "<无结论样句>" Then Exit Sub
    
    txt结论.Text = mnuSampleItem(Index).Tag
    txt结论.SetFocus
End Sub

Private Sub mnuSampleItemDel_Click(Index As Integer)
    Dim i As Integer, intCount As Integer
    Dim intDel As Integer, strText As String
    
    If mnuSampleItem(Index).Caption = "<无结论样句>" Then Exit Sub
    
    intCount = 0: intDel = Index
    If mnuSampleItem(0).Caption <> "<无结论样句>" Then
        intCount = mnuSampleItem.UBound + 1
    End If
    
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\结论样句", "Count", intCount - 1)
    
    For i = intDel To intCount - 2
        strText = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\结论样句", "Item" & i + 1, "")
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\结论样句", "Item" & i, strText)
    Next
    
    Call DeleteSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\结论样句", "Item" & intCount - 1)
End Sub

Private Sub txt结论_GotFocus()
    zlcontrol.TxtSelAll txt结论
End Sub

Private Function LoadSample() As Boolean
    Dim intCount As Integer
    Dim i As Integer
    Dim objMenu As Object
    
    '清除现有内容
    For Each objMenu In mnuSampleItem
        objMenu.Tag = ""
        If objMenu.Index <> 0 Then
            Unload objMenu
        Else
            objMenu.Caption = "<无结论样句>"
        End If
    Next
    For Each objMenu In mnuSampleItemDel
        objMenu.Tag = ""
        If objMenu.Index <> 0 Then
            Unload objMenu
        Else
            objMenu.Caption = "<无结论样句>"
        End If
    Next
    
    intCount = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\结论样句", "Count", 0)
    For i = 0 To intCount - 1
        If i <> 0 Then
            Load mnuSampleItem(i)
            Load mnuSampleItemDel(i)
        End If
        mnuSampleItem(i).Visible = True
        mnuSampleItemDel(i).Visible = True
        mnuSampleItem(i).Tag = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\结论样句", "Item" & i, "")
        mnuSampleItemDel(i).Tag = mnuSampleItem(i).Tag
        
        If zlCommFun.ActualLen(mnuSampleItem(i).Tag) > 20 Then
            mnuSampleItem(i).Caption = Left(mnuSampleItem(i).Tag, 20) & " ..."
        Else
            mnuSampleItem(i).Caption = mnuSampleItem(i).Tag
        End If
        mnuSampleItemDel(i).Caption = mnuSampleItem(i).Caption
    Next
    LoadSample = True
End Function

Private Sub txt结论_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And mblnView Then
        lngTXTProc = GetWindowLong(txt结论.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt结论.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt结论_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And mblnView Then
        Call SetWindowLong(txt结论.hWnd, GWL_WNDPROC, lngTXTProc)
    End If
End Sub

