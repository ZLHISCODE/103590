VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTendBlanketEdit 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   6195
      Index           =   0
      Left            =   105
      ScaleHeight     =   6195
      ScaleWidth      =   5040
      TabIndex        =   8
      Top             =   90
      Width           =   5040
      Begin VB.Frame fra 
         Height          =   6105
         Left            =   105
         TabIndex        =   9
         Top             =   -90
         Width           =   4815
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   1
            Left            =   1125
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1680
            Width           =   1995
         End
         Begin VB.PictureBox pic 
            BorderStyle     =   0  'None
            Height          =   2640
            Left            =   1095
            ScaleHeight     =   2640
            ScaleWidth      =   2460
            TabIndex        =   15
            Top             =   2880
            Width           =   2460
            Begin VB.CommandButton cmd 
               Caption         =   "自定义颜色(&M)…"
               Height          =   350
               Index           =   2
               Left            =   30
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   2220
               Width           =   2205
            End
            Begin zlRichEPR.ColorPicker usrColor 
               Height          =   2190
               Left            =   0
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   0
               Width           =   2190
               _ExtentX        =   3863
               _ExtentY        =   3863
            End
         End
         Begin VB.PictureBox picDemo 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   1125
            ScaleHeight     =   375
            ScaleWidth      =   2160
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   2445
            Width           =   2190
         End
         Begin VB.PictureBox picBack 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   250
            Left            =   1125
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   5655
            Width           =   250
            Begin VB.PictureBox picIcon 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   30
               ScaleHeight     =   180
               ScaleWidth      =   180
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   15
               Width           =   180
            End
         End
         Begin VB.CommandButton cmd 
            Height          =   300
            Index           =   1
            Left            =   1860
            Picture         =   "frmTendBlanketEdit.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   5625
            Width           =   300
         End
         Begin VB.ListBox lst 
            Height          =   1320
            Index           =   0
            Left            =   1125
            Style           =   1  'Checkbox
            TabIndex        =   1
            Top             =   270
            Width           =   1995
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   0
            Left            =   1125
            TabIndex        =   5
            Top             =   2070
            Width           =   1995
         End
         Begin VB.CommandButton cmd 
            Height          =   300
            Index           =   0
            Left            =   1530
            Picture         =   "frmTendBlanketEdit.frx":058A
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   5625
            Width           =   300
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "体温部位(&2)"
            Height          =   180
            Index           =   4
            Left            =   90
            TabIndex        =   2
            Top             =   1740
            Width           =   990
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "符号颜色(&4)"
            Height          =   180
            Index           =   3
            Left            =   90
            TabIndex        =   14
            Top             =   2550
            Width           =   990
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "重叠项目(&1)"
            Height          =   180
            Index           =   0
            Left            =   90
            TabIndex        =   0
            Top             =   300
            Width           =   990
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "标记符号(&3)"
            Height          =   180
            Index           =   1
            Left            =   90
            TabIndex        =   4
            Top             =   2115
            Width           =   990
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "标记图形(&5)"
            Height          =   180
            Index           =   2
            Left            =   90
            TabIndex        =   6
            Top             =   5685
            Width           =   990
         End
      End
   End
   Begin MSComctlLib.ImageList ils 
      Left            =   6885
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   6045
      Top             =   450
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmTendBlanketEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################
Private mlngKey As Long
Private mlngUpKey As Long
Private mlngReferKey As Long
Private mfrmMain As Object
Private mbytMode As Byte
Private mblnAllowModify As Boolean
Private mblnDataChanged As Boolean
Private mblnReading As Boolean

Public Event AfterDataChanged()

'######################################################################################################################
Public Property Let AllowModify(blnData As Boolean)
    mblnAllowModify = blnData
End Property

Public Property Get AllowModify() As Boolean
    AllowModify = mblnAllowModify
End Property

Public Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData

    If mblnReading = False Then
        RaiseEvent AfterDataChanged
    End If
End Property

Public Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Public Function InitData(ByVal frmMain As Object, Optional ByVal blnAllowModify As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    AllowModify = blnAllowModify
    Set mfrmMain = frmMain
    
    Call ExecuteCommand("控件状态")
    If ExecuteCommand("初始数据") = False Then Exit Function

    DataChanged = False

    InitData = True

End Function

Public Function RefreshData(ByVal lngKey As Long, Optional ByVal lngUpKey As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    mlngKey = lngKey
    mlngUpKey = lngUpKey
    mbytMode = 2

    Call ExecuteCommand("清空数据")
    Call ExecuteCommand("控件状态")
    Call ExecuteCommand("读取数据")

    DataChanged = False

    RefreshData = True

End Function

Public Function NewData(Optional ByVal lngReferKey As Long = 0) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    mlngKey = 0
    mlngReferKey = lngReferKey

    mbytMode = 1

    Call ExecuteCommand("清空数据")
    Call ExecuteCommand("控件状态")
    If mlngReferKey > 0 Then
        Call ExecuteCommand("读取数据", mlngReferKey)
    End If

    DataChanged = True

    Call LocationObj(lst(0))

    NewData = True
End Function

Public Function ValidData() As Boolean
    '******************************************************************************************************************
    '功能：校验编辑数据的有效性
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim intCount As Integer

    For intLoop = 0 To lst(0).ListCount - 1
        If lst(0).Selected(intLoop) Then
            intCount = intCount + 1
        End If
    Next
    If intCount < 2 Then
        ShowSimpleMsg "必须选择两个及以上体温曲线项目！"
        Call LocationObj(lst(0))
        Exit Function
    End If
    
    If picIcon.Tag = "" And Trim(Cbo(0).Text) = "" Then
    
        ShowSimpleMsg "必须设置一个代表字符或图形！"
        Call LocationObj(Cbo(0))
        Exit Function
        
    End If
    
    ValidData = True

End Function

Public Function SaveData(ByRef rsSQL As ADODB.Recordset, ByRef lngKey As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim str项目序号 As String
    Dim str重叠项目 As String
    Dim intLoop As Integer
    Dim strTmp As String

    On Error GoTo errHand

    If mlngKey = 0 Then
        '新增
        strSQL = "Select Nvl(Max(序号),0)+1 As 序号 From 体温重叠标记"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If rs.BOF Then Exit Function

        lngKey = rs("序号").Value
    Else
        '修改
        lngKey = mlngKey
    End If

    str项目序号 = ""
    str重叠项目 = ""
    For intLoop = 0 To lst(0).ListCount - 1
        If lst(0).Selected(intLoop) Then
            str项目序号 = str项目序号 & "," & lst(0).ItemData(intLoop)
            
            If lst(0).ItemData(intLoop) = 1 And Trim(Cbo(1).Text) <> "" Then
                str重叠项目 = str重叠项目 & "," & lst(0).List(intLoop) & "(" & Trim(Cbo(1).Text) & ")"
            Else
                str重叠项目 = str重叠项目 & "," & lst(0).List(intLoop)
            End If
            
        End If
    Next
    If str项目序号 <> "" Then str项目序号 = Mid(str项目序号, 2)
    If str重叠项目 <> "" Then str重叠项目 = Mid(str重叠项目, 2)

    If picBack.Tag <> "Delete" Then

        strSQL = "zl_体温重叠标记_Update(" & lngKey & ",'" & str项目序号 & "','" & str重叠项目 & "','" & Trim(Cbo(0).Text) & "'," & Val(picDemo.Tag) & ",1,'" & IIf(Cbo(1).Enabled, Trim(Cbo(1).Text), "") & "')"

    Else
        strSQL = "zl_体温重叠标记_Update(" & lngKey & ",'" & str项目序号 & "','" & str重叠项目 & "','" & Trim(Cbo(0).Text) & "'," & Val(picDemo.Tag) & ",0,'" & IIf(Cbo(1).Enabled, Trim(Cbo(1).Text), "") & "')"
    End If
    Call SQLRecordAdd(rsSQL, strSQL)

    If picBack.Tag <> "" Then
        If picBack.Tag <> "Delete" Then
            strTmp = picBack.Tag
            Call SQLRecordAdd(rsSQL, "", 0, 100, "9;" & lngKey & ";" & strTmp)
        End If
    End If

    SaveData = True

    Exit Function

    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
End Function

'（三）模块内的属性、方法及函数
'######################################################################################################################

Private Function DrawDemo(pic As PictureBox, lngColor As Long) As Boolean
    Dim lngStartX As Long
    Dim lngStartY As Long

    pic.Cls

    lngStartX = (pic.Width - pic.TextWidth("标记符号")) / 2
    lngStartY = (pic.Height - pic.TextHeight("标记符号") * 3) / 2

    Call DrawText(pic, pic.TextWidth("AA"), (pic.Height - pic.TextHeight("标记符号")) / 2, "标记符号", lngColor)

    usrColor.COLOR = lngColor
End Function

Private Function ExecuteCommand(ByVal strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能： 内部模块执行命令
    '参数： strCommand          命令
    '       varParam            传入参数,可变参数方式
    '返回： 执行成功返回True;否则返回False
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    Dim blnAllowModify As Boolean

    On Error GoTo errHand
    mblnReading = True
    Call SQLRecord(rsSQL)

    Select Case strCommand
    '--------------------------------------------------------------------------------------------------------------
    Case "初始数据"
        Cbo(0).Clear
        Cbo(0).AddItem "A"
        Cbo(0).AddItem "B"
        Cbo(0).AddItem "C"
        
        gstrSQL = "Select 部位 From 体温部位 Where 项目序号 = 1 Order By 部位 Desc"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "体温部位", mlngKey)
        Cbo(1).Clear
        Cbo(1).AddItem ""
        Do While Not rs.EOF
            Cbo(1).AddItem rs!部位
            rs.MoveNext
        Loop
        Cbo(1).ListIndex = 0
        
        lst(0).Clear
        gstrSQL = "Select 项目序号,记录名 From 体温记录项目 Where 记录法=1"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
        If rs.BOF = False Then
            Do While Not rs.EOF

                lst(0).AddItem zlCommFun.NVL(rs("记录名").Value)
                lst(0).ItemData(lst(0).NewIndex) = zlCommFun.NVL(rs("项目序号").Value, 0)

                rs.MoveNext
            Loop
        End If
        
        Call ExecuteCommand("控件状态")
        
    '--------------------------------------------------------------------------------------------------------------
    Case "清空数据"

        For intLoop = 0 To lst(0).ListCount - 1
            lst(0).Selected(intLoop) = False
        Next
        Cbo(0).Text = ""
        picBack.Tag = ""
        picIcon.Cls
        picIcon.Tag = ""

    '------------------------------------------------------------------------------------------------------------------
    Case "控件状态"
        blnAllowModify = mblnAllowModify
        If mlngKey = 0 And mbytMode = 2 Then blnAllowModify = False

        lst(0).Enabled = blnAllowModify
        Cbo(0).Enabled = blnAllowModify
        cmd(0).Enabled = blnAllowModify
        cmd(1).Enabled = blnAllowModify
        pic.Enabled = blnAllowModify
        
    '--------------------------------------------------------------------------------------------------------------
    Case "读取数据"

        gstrSQL = "Select 序号,项目序号,标记符号,标记颜色,体温部位 From 体温重叠标记 start with 序号=[1] Connect by prior 序号=上级序号"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)

        If rs.BOF = False Then

            Cbo(0).Text = zlCommFun.NVL(rs("标记符号").Value)
            
            picDemo.Tag = zlCommFun.NVL(rs("标记颜色"), 0)
            Call DrawDemo(picDemo, Val(picDemo.Tag))

            rs.MoveNext
            Do While Not rs.EOF
                For intLoop = 0 To lst(0).ListCount - 1
                    If lst(0).ItemData(intLoop) = zlCommFun.NVL(rs("项目序号").Value, 0) Then
                        lst(0).Selected(intLoop) = True
                        If zlCommFun.NVL(rs("项目序号").Value, 0) = 1 Then
                            zlControl.CboLocate Cbo(1), zlCommFun.NVL(rs("体温部位").Value)
                        End If
                        Exit For
                    End If
                Next
                rs.MoveNext
            Loop

            picIcon.Tag = ""
            '读取标记图形并显示
            strTmp = zlBlobRead(9, mlngKey)
            If Dir(strTmp) <> "" And strTmp <> "" Then
                Call DrawPicture(picIcon, strTmp, 0, 0, picIcon.Width, picIcon.Height)
                picIcon.Tag = "Have"
                Kill strTmp
            End If

        End If

    End Select

    ExecuteCommand = True

    GoTo EndHand

    '错误处理
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog

    '结束处理
    '------------------------------------------------------------------------------------------------------------------
EndHand:

    mblnReading = False
End Function

Private Sub cbo_Change(Index As Integer)
    DataChanged = True
End Sub

Private Sub cbo_Click(Index As Integer)
    DataChanged = True
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub


Private Sub cmd_Click(Index As Integer)

    Dim strTmp As String
'    Dim objStd As StdPicture
    Dim objFile As New FileSystemObject

    Select Case Index
    Case 0
        With dlg
            .DialogTitle = "请选择要添加的标记图标文件"
            .Filter = "标记图标(*.ico)|*.ico"

            On Error Resume Next

            .flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
            .Filename = ""
            .MaxFileSize = 32767
            .ShowOpen

            If Err.Number = 0 And .Filename <> "" Then

                strTmp = .Filename

                picBack.Tag = strTmp

                On Error GoTo errHand

                picIcon.Cls
                Call DrawPicture(picIcon, strTmp, 0, 0, picIcon.Width, picIcon.Height)
                picIcon.Tag = "Have"
                
                DataChanged = True
            Else
                Err.Clear
            End If
        End With
    Case 1
        picIcon.Cls
        picIcon.Tag = ""
        picBack.Tag = "Delete"
        DataChanged = True
    Case 2

        dlg.COLOR = Val(picDemo.Tag)
        dlg.ShowColor

        If dlg.COLOR <> Val(picDemo.Tag) Then

            picDemo.Tag = dlg.COLOR
            Call DrawDemo(picDemo, dlg.COLOR)

            DataChanged = True

        End If

    End Select

    Exit Sub

    '------------------------------------------------------------------------------------------------------------------
errHand:
    ShowSimpleMsg "不能打开文件(" & strTmp & "),该文件可能正在使用或文件不存在!"

End Sub

Private Sub Form_Resize()
    On Error Resume Next

    picPane(0).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight

End Sub

Private Sub lst_ItemCheck(Index As Integer, Item As Integer)
    DataChanged = True
    
    If lst(Index).List(Item) = "体温" Then
        Cbo(1).Enabled = lst(Index).Selected(Item) And Cbo(0).Enabled
    End If
End Sub

Private Sub lst_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub picPane_Resize(Index As Integer)

    On Error Resume Next

    fra.Move 0, -75, Me.ScaleWidth, Me.ScaleHeight + 75

    lst(0).Move lst(0).Left, lst(0).Top, fra.Width - lst(0).Left - 75
    Cbo(0).Move Cbo(0).Left, Cbo(0).Top, fra.Width - Cbo(0).Left - 75
    picDemo.Move picDemo.Left, picDemo.Top, fra.Width - picDemo.Left - 75

End Sub

Private Sub usrColor_pOK()
    If usrColor.COLOR < 0 Then usrColor.COLOR = 0   '控件缺省颜色为负数，而颜色的有效值大于零
    If Val(picDemo.Tag) = usrColor.COLOR Then Exit Sub
    picDemo.Tag = usrColor.COLOR
    Call DrawDemo(picDemo, usrColor.COLOR)
    DataChanged = True
End Sub

