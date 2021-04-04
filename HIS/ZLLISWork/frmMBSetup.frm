VERSION 5.00
Begin VB.Form frmMBSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "酶标仪设置"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboMachine 
      Height          =   300
      Left            =   1470
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   180
      Width           =   2505
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3435
      TabIndex        =   9
      Top             =   1950
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2265
      TabIndex        =   8
      Top             =   1950
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   -45
      TabIndex        =   10
      Top             =   1680
      Width           =   4785
   End
   Begin VB.ComboBox cboPosi 
      Height          =   300
      Left            =   1470
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1215
      Width           =   1215
   End
   Begin VB.TextBox txtNO 
      Height          =   300
      Left            =   1470
      MaxLength       =   20
      TabIndex        =   5
      Top             =   870
      Width           =   1185
   End
   Begin VB.TextBox txtItem 
      Height          =   300
      Left            =   1470
      TabIndex        =   3
      Top             =   525
      Width           =   2505
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "检验仪器(&M)"
      Height          =   180
      Left            =   435
      TabIndex        =   0
      Top             =   255
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "起始位置(&S)"
      Height          =   180
      Left            =   435
      TabIndex        =   6
      Top             =   1260
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "起始标本号(&H)"
      Height          =   180
      Left            =   255
      TabIndex        =   4
      Top             =   930
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "检验项目(&I)"
      Height          =   180
      Left            =   435
      TabIndex        =   2
      Top             =   585
      Width           =   990
   End
End
Attribute VB_Name = "frmMBSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mstrItem As String
Public Function ShowMe(ByVal frmMain As Object) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, j As Integer
    Dim lngDeviceID As Long, strItem As String
    
    mblnOK = False
    mstrItem = ""
    
    On Error GoTo DBError
    
    '检验仪器
    gstrSql = "Select * From 检验仪器"
    OpenRecord rsTmp, gstrSql, Me.Caption
    If rsTmp.EOF Then
        MsgBox "没有初始检验仪器，无法设置！", vbCritical, Me.Caption
        Unload Me
        Exit Function
    End If
    
    With cboMachine
        .Clear
        Do While Not rsTmp.EOF
            .AddItem "(" & rsTmp("编码") & ")" & rsTmp("名称")
            .ItemData(.ListCount - 1) = rsTmp("ID")
            
            rsTmp.MoveNext
        Loop
    End With
    lngDeviceID = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "酶标仪器", -1))
    If lngDeviceID = -1 Then
        cboMachine.ListIndex = 0
    Else
        cboMachine.ListIndex = FindComboItem(cboMachine, lngDeviceID)
    End If
    
    On Error Resume Next
    '检验项目
    strItem = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "酶标仪项目", "")
    If Len(strItem) = 0 Then
        mstrItem = ""
        txtItem = "": txtItem.Tag = ""
    Else
        mstrItem = Split(strItem, "|")(0)
        txtItem = mstrItem: txtItem.Tag = Split(strItem, "|")(1)
    End If
    '标本号
    txtNO = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "酶标仪标本号", "")
    '起始位置
    With cboPosi
        .Clear
        For i = 1 To 8
            For j = 1 To 12
                .AddItem Chr(64 + i) & Format(j, "0#")
            Next j
        Next i
    End With
    cboPosi.Text = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "酶标仪起始位置", "A01")
    
    
    Me.Show vbModal, frmMain
    
    ShowMe = mblnOK
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cboMachine_Click()
    mstrItem = "": txtItem = ""
End Sub

Private Sub cboMachine_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cboPosi_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Len(Trim(txtItem)) = 0 Then
        MsgBox "请指定当前酶标仪的检验项目", , gstrSysName
        txtItem.SetFocus
        Exit Sub
    End If
    If Len(Trim(txtNO)) = 0 Then
        MsgBox "请输入初始的标本号", , gstrSysName
        txtNO.SetFocus
        Exit Sub
    End If
    
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName, "酶标仪器", cboMachine.ItemData(cboMachine.ListIndex))
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName, "酶标仪项目", txtItem & "|" & txtItem.Tag)
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName, "酶标仪标本号", txtNO)
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName, "酶标仪起始位置", cboPosi.Text)
    
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtItem_GotFocus()
    zlControl.TxtSelAll txtItem
End Sub

Private Sub txtItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtItem_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
    
    If mstrItem = txtItem Then Exit Sub
        
    strSQL = "SELECT Distinct 项目ID As ID,通道编码,中文名||'('||英文名||')' As 检验项目 " & _
        "FROM 检验仪器项目 A,诊治所见项目 B,诊疗项目目录 C,诊疗项目别名 D " & _
        "WHERE A.项目id=B.ID And A.仪器ID=[1] AND B.编码=C.编码 AND C.ID=D.诊疗项目ID " & _
        "AND (Upper(B.英文名) LIKE [2] OR Upper(D.简码) LIKE [2] OR Upper(B.中文名) LIKE [2])"
    
    On Error GoTo errH
    vRect = GetControlRect(txtItem.Hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "检验项目", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, txtItem.Height, blnCancel, False, True, cboMachine.ItemData(cboMachine.ListIndex), UCase(txtItem) & "%")
    If Not rsTmp Is Nothing Then
        txtItem.Text = rsTmp!检验项目
        txtItem.Tag = rsTmp!通道编码
        mstrItem = txtItem
    Else
        If Not blnCancel Then
            MsgBox "未找到对应的检验项目！", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtNO_GotFocus()
    zlControl.TxtSelAll txtNO
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub
