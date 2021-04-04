VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmColChoose 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12885
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   12885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fraLine 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.TextBox txtPrint 
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   840
         TabIndex        =   2
         Top             =   3720
         Width           =   800
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   0
         TabIndex        =   1
         Top             =   3720
         Width           =   800
      End
      Begin MSComctlLib.ListView lvwSub 
         Height          =   3180
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   5609
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "名称"
            Object.Width           =   2117
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgChoose 
      Left            =   8160
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColChoose.frx":0000
            Key             =   "NoFilter_1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColChoose.frx":6862
            Key             =   "NoFilter"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColChoose.frx":7274
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColChoose.frx":7C86
            Key             =   "Filter_1"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmColChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type
Private mblnOK As Boolean
Private mvsfMain As VSFlexGrid '窗体卸载时不可清空
Private mcolData As New Collection '保存数据形式：a.名称 In('部门表')<JM>1<JM>'部门表'；中间值含义0-输入，1-勾选
Private mstrFild As String
Private mlngMouseCol As Long

Public Function ShowMe(ByRef objbind As VSFlexGrid, ByVal strFild As String, ByRef strCondition As String) As Boolean
'参数：strFild-该列所表示的字段表达式
    Dim sglX As Single
    Dim sglY As Single
    Dim strTemp As String
    Dim i As Long
    
    Set mvsfMain = objbind
    mlngMouseCol = objbind.MouseCol
    Call CalcSelectFormPostion(sglX, sglY)
        
    mstrFild = strFild
    mblnOK = False
    Me.Left = sglX
    Me.Top = sglY
    Call LoadData
    If lvwSub.ListItems.Count > 7 Then
        lvwSub.Height = 10 * 200
    Else
        lvwSub.Height = (lvwSub.ListItems.Count + 2) * 200
    End If

    Call GetHistory
    Me.Show 1
    ShowMe = mblnOK
    If mblnOK Then
        For i = 1 To mcolData.Count
            strTemp = IIf(strTemp = "", "", strTemp & " And ") & Split(mcolData.Item(i), "<JM>")(0)
        Next
        If strTemp <> "" Then strTemp = " And " & strTemp
        strCondition = strTemp
    Else
        strCondition = ""
    End If
    strTemp = GetColItem(mcolData, "K_" & mlngMouseCol)
    If strTemp = "" Then
        mvsfMain.Cell(flexcpPicture, 0, mlngMouseCol) = imgChoose.ListImages("NoFilter").Picture
    Else
        mvsfMain.Cell(flexcpPicture, 0, mlngMouseCol) = imgChoose.ListImages("Filter").Picture
    End If
End Function

Private Sub CalcSelectFormPostion(ByRef sglX As Single, ByRef sglY As Single)

    Dim lngTrayH As Long
    Dim sglObjHeight As Single
    Dim objPoint As POINTAPI
    Dim lngH0 As Long
    Dim lngH1 As Long
    
    '------------------------------------------------------------------------------------------------------------------
    Call ClientToScreen(mvsfMain.hwnd, objPoint)
    sglX = objPoint.x * Screen.TwipsPerPixelX + mvsfMain.Cell(flexcpLeft, 0, mlngMouseCol)
    If mvsfMain.CellHeight = 0 Then
        sglY = objPoint.y * Screen.TwipsPerPixelY + mvsfMain.rowHeight(0) + 50
    Else
        sglY = objPoint.y * Screen.TwipsPerPixelY + mvsfMain.CellHeight
    End If
    sglObjHeight = mvsfMain.CellHeight
    
    '检查是否超过屏幕高和宽度
    '------------------------------------------------------------------------------------------------------------------
    lngTrayH = GetTrayHeight
    
    If sglX > Screen.Width Then
        If Screen.Width >= 0 Then
            sglX = Screen.Width
        Else
            sglX = 0
        End If
    End If
    
    If sglY > (Screen.Height - lngTrayH) Then
        If (sglY - sglObjHeight) >= 0 Then
            '放在输入框的上面
            sglY = sglY - sglObjHeight - 2 * Screen.TwipsPerPixelY
        Else
            '分别计算放置上面和放置下面的高度,取最大高度
            lngH0 = sglY - sglObjHeight
            lngH1 = Screen.Height - lngTrayH - sglY
            
            If lngH0 > lngH1 Then
            
                '上面高
                sglY = 0
            End If
        End If
    End If
    
End Sub

Private Sub LoadData()
'加载列的所有可能值
    Dim i As Long
    Dim strData As String
    Dim varTemp As Variant
    
    strData = mvsfMain.ColData(mlngMouseCol)
    strData = Mid("全部" & strData, 1, Len("全部" & strData) - 1)
    varTemp = Split(strData, ",")
    For i = 0 To UBound(varTemp)
        lvwSub.ListItems.Add i + 1, , varTemp(i)
    Next
End Sub

Private Sub GetHistory()
    Dim strTemp As String
    Dim varTemp As Variant
    Dim i As Long
    
    With mvsfMain
        strTemp = GetColItem(mcolData, "K_" & mlngMouseCol)
        If strTemp = "" Then
            For i = 1 To lvwSub.ListItems.Count
                lvwSub.ListItems(i).Checked = True
            Next
            Exit Sub
        End If
        varTemp = Split(strTemp, "<JM>")
        If varTemp(1) = 0 Then
            txtPrint.Text = varTemp(2)
        Else
            strTemp = Replace(varTemp(2), "'", "")
            varTemp = Split(strTemp, ",")
            For i = 0 To UBound(varTemp)
                If varTemp(i) <> "" Then
                    lvwSub.FindItem(varTemp(i)).Checked = True
                End If
            Next
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    Dim strTxt As String
    Dim strTemp As String, strSQL As String, strChoose As String
    Dim blnNull As Boolean, blnNotNull As Boolean

    strTxt = txtPrint.Text
    If Trim(strTxt) <> "" Then
        If InStr(strTxt, "'") > 0 Then
            MsgBox "过滤条件不允许有单引号，请重新输入！", vbExclamation, Me.Caption
            txtPrint.SetFocus
            Exit Sub
        End If
        mblnOK = True
        Call RemoveCol
        strSQL = " " & mstrFild & " Like '%" & strTxt & "%' "
        mcolData.Add strSQL & "<JM>0<JM>" & strTxt, "K_" & mlngMouseCol
    Else
        With lvwSub
            If .ListItems(1).Checked = True Then
                '当选择全部时,移去原先已有的值,表示当前列的全部
                mblnOK = True
                Call RemoveCol
            Else
                For i = 2 To .ListItems.Count
                    If .ListItems(i).Checked = True Then
                        strTemp = IIf(strTemp = "", "", strTemp & ",") & "'" & .ListItems(i).Text & "'"
                    End If
                Next
                If strTemp <> "" Then
                    strSQL = " " & mstrFild & " In(" & strTemp & ")"
                Else
                    Unload Me
                    Exit Sub
                End If
                Call RemoveCol
                mcolData.Add strSQL & "<JM>1<JM>" & strTemp, "K_" & mlngMouseCol
                mblnOK = True
            End If
        End With
    End If
    Unload Me
End Sub

Public Sub ClearCol()
'在所有选择全部和需要重新刷新时，清空集合数据
    
    Do While mcolData.Count > 0
        mcolData.Remove 1
    Loop
End Sub

Private Sub RemoveCol()
    Dim strTemp As String

    strTemp = GetColItem(mcolData, "K_" & mlngMouseCol)
    If strTemp <> "" Then mcolData.Remove "K_" & mlngMouseCol
End Sub

Private Function GetColItem(ByRef colTemp As Collection, ByRef strKey As String) As String
'功能：按关键字获取集合项目
    On Error Resume Next
    GetColItem = colTemp(strKey)
    err.Clear
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Resize()

    If mvsfMain.ColWidth(mlngMouseCol) < 2000 Then
        Me.Width = 1900
    Else
        Me.Width = mvsfMain.ColWidth(mlngMouseCol)
    End If
    txtPrint.Top = 100
    txtPrint.Left = 15
    txtPrint.Width = Me.Width - 45
    lvwSub.Top = txtPrint.Height + txtPrint.Top
    lvwSub.Left = 15
    lvwSub.Width = Me.Width - 45
    lvwSub.ColumnHeaders.Item(1).Width = lvwSub.Width
    Me.Height = lvwSub.Height + txtPrint.Height + cmdOK.Height * 1.4
    cmdOK.Left = Me.Width / 2 - cmdOK.Width - 100
    cmdCancel.Left = cmdOK.Left + cmdOK.Width + 150
    cmdOK.Top = txtPrint.Height + lvwSub.Height + 80
    cmdCancel.Top = cmdOK.Top
    fraLine.Top = -75
    fraLine.Left = 0
    fraLine.Height = Me.Height
    fraLine.Width = Me.Width
End Sub

Private Sub lvwSub_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim i As Long
    Dim blnAllCheck As Boolean
    
    Item.Selected = True
    If lvwSub.SelectedItem.Index = 1 Then
        For i = 2 To lvwSub.ListItems.Count
            lvwSub.ListItems(i).Checked = IIf(lvwSub.ListItems(1).Checked = True, True, False)
        Next
    Else
        If lvwSub.ListItems(1).Checked = True Then
            lvwSub.ListItems(1).Checked = False
        Else
            For i = 2 To lvwSub.ListItems.Count
                If lvwSub.ListItems(i).Checked = True Then
                    If i = lvwSub.ListItems.Count Then
                        blnAllCheck = True
                    End If
                Else
                    Exit For
                End If
            Next
            If blnAllCheck = True Then lvwSub.ListItems(1).Checked = True
        End If
    End If
End Sub

Private Sub txtPrint_KeyPress(KeyAscii As Integer)
    Dim strTxt As String
    Dim strSQL As String
    
    If KeyAscii = 13 Then
        strTxt = txtPrint.Text
        If InStr(strTxt, "'") > 0 Then
            MsgBox "过滤条件不允许有单引号，请重新输入！", vbExclamation, Me.Caption
            txtPrint.SetFocus
            Exit Sub
        End If
        mblnOK = True
        Call RemoveCol
        strSQL = " " & mstrFild & " Like '%" & strTxt & "%' "
        mcolData.Add strSQL & "<JM>0<JM>" & strTxt, "K_" & mlngMouseCol
        Unload Me
    End If
End Sub

Private Function GetTrayHeight() As Long
    '******************************************************************************************************************
    '功能:获取任务栏的高度
    '******************************************************************************************************************
    Dim lngHwd As Long
    Dim objRect As RECT
    
    On Error Resume Next
    
    lngHwd = FindWindow("shell_traywnd", "")
    Call GetWindowRect(lngHwd, objRect)

    GetTrayHeight = Screen.TwipsPerPixelX * (objRect.Bottom - objRect.Top)
    
    If GetTrayHeight < 0 Then GetTrayHeight = 0
    
End Function
