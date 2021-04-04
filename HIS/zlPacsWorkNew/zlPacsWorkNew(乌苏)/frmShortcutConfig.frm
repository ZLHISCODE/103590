VERSION 5.00
Begin VB.Form frmShortcutConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "快捷键配置"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9990
   Icon            =   "frmShortcutConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdDefault 
      Caption         =   "恢复默认(&D)"
      Height          =   400
      Left            =   4680
      TabIndex        =   5
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删 除(&D)"
      Height          =   400
      Left            =   6000
      TabIndex        =   3
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "确 定(&S)"
      Height          =   400
      Left            =   7320
      TabIndex        =   2
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取 消(&C)"
      Height          =   400
      Left            =   8640
      TabIndex        =   1
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   160
      Width           =   9735
      Begin zl9PACSWork.ucFlexGrid ufgShoftcut 
         Height          =   5895
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   10398
         DefaultCols     =   ""
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         HeadFontCharset =   134
         HeadFontWeight  =   400
         DataFontCharset =   134
         DataFontWeight  =   400
      End
   End
End
Attribute VB_Name = "frmShortcutConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#Const DebugState = False



Private Const C_STR_SHORTCUT_LIST_COLS As String = "|菜单分组,merge,read,w1800,txtcenter|ID,hide,key,uncfg|项目,hide,uncfg|模块号,hide,uncfg|菜单说明,read,w2100,uncfg|控制键,hide,uncfg|字符键,hide,uncfg|默认键,hide,uncfg|快捷键>组合名,read,w2100,uncfg|当前键>组合名,hide,uncfg|"
Private mstrProject As String
Private mlngMudule As Long

Public blnIsOk As Boolean

Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer


Private Enum ShiftKeys
    AltKey = &H1
    CtrlKey = &H2
    ShiftKey = &H4
End Enum



Public Sub ShowShortcutConfig(ByVal strProject As String, ByVal lngModule As Long, owner As Object)
    mstrProject = strProject
    mlngMudule = lngModule
    
    blnIsOk = False
    
    Call Me.Show(1, owner)
End Sub

Private Sub cmdCancel_Click()
    blnIsOk = False
    
    Call Me.Hide
End Sub

Private Sub cmdDefault_Click()
On Error GoTo ErrHandle
    Call LoadDefaultShortcut
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub LoadDefaultShortcut()
    Dim i As Long
    Dim strDefaultKey As String
    Dim intShift As Integer
    Dim intKeyCode As Integer
    
    For i = 1 To ufgShoftcut.GridRows - 1
        strDefaultKey = ufgShoftcut.Text(i, "默认键")
        
        If strDefaultKey <> "" Then
            intShift = Val(Mid(strDefaultKey, 1, InStr(strDefaultKey, "+") - 1))
            intKeyCode = Val(Mid(strDefaultKey, InStr(strDefaultKey, "+") + 1, 8))
        
            ufgShoftcut.Text(i, "控制键") = intShift
            ufgShoftcut.Text(i, "字符键") = intKeyCode
            ufgShoftcut.Text(i, "快捷键") = GetKyeAlias(intKeyCode, intShift)
        Else
            ufgShoftcut.Text(i, "控制键") = 0
            ufgShoftcut.Text(i, "字符键") = 0
            ufgShoftcut.Text(i, "快捷键") = ""
        End If
    Next i
End Sub






Private Sub cmdSure_Click()
On Error GoTo ErrHandle
    Call UpdateShortCut
    
    blnIsOk = True
    
    Call Me.Hide
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call RestoreWinState(Me, App.ProductName)
    
    blnIsOk = False
    
    #If DebugState = True Then
        mstrProject = "ZL9PACSWORK"
        mlngMudule = 1290
        
        Call InitDebugObject(1290, Me, "zlhis", "HIS")
    #End If
    
    Call InitShoftcutList
    
    Call LoadShortCutData
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub InitShoftcutList()
    '设置行数
    ufgShoftcut.GridRows = glngStandardRowCount
    '设置行高
    ufgShoftcut.RowHeightMin = glngStandardRowHeight
    
    '禁止右键弹出列表配置窗口
    ufgShoftcut.IsEjectConfig = False
    
    ufgShoftcut.IsKeepRows = False
    ufgShoftcut.ColNames = C_STR_SHORTCUT_LIST_COLS
    ufgShoftcut.DefaultColNames = C_STR_SHORTCUT_LIST_COLS
    ufgShoftcut.ColConvertFormat = ""
End Sub


Private Sub LoadShortCutData()
    Dim strSql As String
    
    strSql = "select a.id, a.项目, a.模块号, a.菜单ID, a.菜单分组, a.菜单说明, nvl(b.控制键, a.控制键) as 控制键, " & _
             "nvl(b.字符键, a.字符键) as 字符键, a.默认键, decode(nvl(b.快捷功能ID,''),'',a.组合名,b.组合名) as 组合名, a.分组序号 " & _
             "from 快捷功能信息 a, (select 快捷功能ID, 控制键, 字符键, 组合名 from 快捷功能关联 where 用户id=[1] )b " & _
             "where a.id=b.快捷功能ID(+) and a.项目=[2] and a.模块号=[3] order by a.分组序号,a.id"
        
    Set ufgShoftcut.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID, UCase(mstrProject), mlngMudule)
    Call ufgShoftcut.RefreshData
    
End Sub


Private Function GetKyeAlias(KeyCode As Integer, Shift As Integer) As String

    Dim strShift As String
    Dim strTemp As String
    
    
    strShift = IIf((Shift And vbCtrlMask) <> 0, "CTRL", "")
    
    strTemp = IIf((Shift And vbShiftMask) <> 0, "SHIFT", "")
    If strTemp <> "" Then
        If strShift <> "" Then strShift = strShift & "+"
        strShift = strShift & strTemp
    End If
    
    strTemp = IIf((Shift And vbAltMask) <> 0, "ALT", "")
    If strTemp <> "" Then
        If strShift <> "" Then strShift = strShift & "+"
        strShift = strShift & strTemp
    End If
    
     
    
             
    strTemp = ""
    If KeyCode >= 48 And KeyCode <= 90 Then
        strTemp = Chr(KeyCode)
        
        If strShift = "" Then strShift = "MENU"
    End If
    
    If KeyCode >= vbKeyF1 And KeyCode <= vbKeyF12 Then
        strTemp = "F" & (KeyCode - 111)
    End If
    
    Select Case KeyCode
        Case vbKeySpace
            strTemp = "SPACE"
    End Select
    
    
    If strTemp <> "" Then
        If strShift <> "" Then strShift = strShift & "+"
        strShift = strShift & strTemp
    End If
    
    GetKyeAlias = strShift
                
End Function


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    err.Clear
End Sub

Private Sub ufgShoftcut_OnKeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Dim strAlias As String
    
    If Not ufgShoftcut.IsSelectionRow Then Exit Sub
    
    '删除快捷键
    If KeyCode = vbKeyDelete Then
        ufgShoftcut.Text(ufgShoftcut.SelectionRow, "快捷键") = ""
        ufgShoftcut.Text(ufgShoftcut.SelectionRow, "控制键") = 0
        ufgShoftcut.Text(ufgShoftcut.SelectionRow, "字符键") = 0
        
        Exit Sub
    End If
    
    strAlias = GetKyeAlias(KeyCode, Shift)
    If strAlias = "" Then Exit Sub
    
    
    ufgShoftcut.Text(ufgShoftcut.SelectionRow, "快捷键") = strAlias
    
    ufgShoftcut.Text(ufgShoftcut.SelectionRow, "控制键") = Shift
    ufgShoftcut.Text(ufgShoftcut.SelectionRow, "字符键") = KeyCode
    
End Sub

Private Sub cmdDelete_Click()
'删除按钮 执行删除选择行的快捷键设置
On Error Resume Next
    
    If Not ufgShoftcut.IsSelectionRow Then Exit Sub
    
    ufgShoftcut.Text(ufgShoftcut.SelectionRow, "快捷键") = ""
    ufgShoftcut.Text(ufgShoftcut.SelectionRow, "控制键") = 0
    ufgShoftcut.Text(ufgShoftcut.SelectionRow, "字符键") = 0
    
End Sub


'更新快捷键配置
Private Sub UpdateShortCut()
    Dim i As Long
    Dim strSql As String
    
    For i = 1 To ufgShoftcut.GridRows - 1
        If ufgShoftcut.Text(i, "快捷键") <> ufgShoftcut.Text(i, "当前键") Then
            strSql = "ZL_快捷键_更新(" & ufgShoftcut.KeyValue(i) & "," & _
                                            Val(UserInfo.ID) & "," & _
                                            Val(ufgShoftcut.Text(i, "控制键")) & "," & _
                                            Val(ufgShoftcut.Text(i, "字符键")) & ",'" & _
                                            ufgShoftcut.Text(i, "快捷键") & "')"
    
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        End If
        '更新行状态
        ufgShoftcut.RowState(i) = TDataRowState.Normal
    Next i
    
End Sub



Private Sub ufgShoftcut_OnKeyUp(KeyCode As Integer, Shift As Integer)
'判断是否存在重复的快捷键
On Error GoTo ErrHandle
    Dim strKeyAlias As String
    Dim lngFindIndex As Long
    
    If Not ufgShoftcut.IsSelectionRow Then Exit Sub
    
    strKeyAlias = ufgShoftcut.Text(ufgShoftcut.SelectionRow, "快捷键")
    
    If InStr(strKeyAlias, "MENU+") >= 1 Then Exit Sub
    
    ufgShoftcut.Text(ufgShoftcut.SelectionRow, "快捷键") = ""
    
    lngFindIndex = ufgShoftcut.FindRowIndex(strKeyAlias, "快捷键", True)
    
    If lngFindIndex > 0 Then
        ufgShoftcut.Text(ufgShoftcut.SelectionRow, "控制键") = 0
        ufgShoftcut.Text(ufgShoftcut.SelectionRow, "字符键") = 0
        
        Call MsgBoxD(Me, "存在重复的快捷键，请重新设置。", vbOKOnly, Me.Caption)
    Else
        ufgShoftcut.Text(ufgShoftcut.SelectionRow, "快捷键") = strKeyAlias
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub
