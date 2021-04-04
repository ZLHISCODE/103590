VERSION 5.00
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.12#0"; "zlIDKind.ocx"
Begin VB.UserControl usrCardPeople 
   ClientHeight    =   6405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3825
   ScaleHeight     =   6405
   ScaleWidth      =   3825
   Begin VB.PictureBox picFont 
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   2535
      ScaleHeight     =   150
      ScaleWidth      =   210
      TabIndex        =   21
      Top             =   135
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox pic10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   120
      ScaleHeight     =   270
      ScaleWidth      =   3465
      TabIndex        =   15
      Top             =   5640
      Visible         =   0   'False
      Width           =   3495
      Begin VB.Frame frm10 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   1800
         TabIndex        =   20
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox TXT10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   1800
         TabIndex        =   19
         Text            =   "7"
         ToolTipText     =   "显示当前页数，可以输入指定页数，并按回车跳转到指定页"
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lbl10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "总的页数"
         Top             =   0
         Width           =   650
      End
      Begin VB.Label lbl12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "下一页"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2520
         TabIndex        =   17
         ToolTipText     =   "点击跳到下一页"
         Top             =   30
         Width           =   705
      End
      Begin VB.Label lbl11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "上一页"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   960
         TabIndex        =   16
         ToolTipText     =   "点击跳到上一页"
         Top             =   0
         Width           =   705
      End
   End
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4425
      Left            =   120
      ScaleHeight     =   4395
      ScaleWidth      =   3465
      TabIndex        =   2
      Top             =   1110
      Width           =   3495
      Begin VB.PictureBox Pic2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1000
         Index           =   0
         Left            =   240
         Picture         =   "usrCardPeople.ctx":0000
         ScaleHeight     =   1030.769
         ScaleMode       =   0  'User
         ScaleWidth      =   4095
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   4095
         Begin VB.PictureBox pic4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   255
            TabIndex        =   5
            Top             =   480
            Width           =   255
            Begin VB.CheckBox chk1 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   6
               Top             =   0
               Width           =   255
            End
         End
         Begin VB.Shape shpRight 
            BorderColor     =   &H8000000D&
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00FF0000&
            FillStyle       =   0  'Solid
            Height          =   495
            Index           =   0
            Left            =   3090
            Top             =   225
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Shape shpLeft 
            BorderColor     =   &H8000000D&
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00FF0000&
            FillStyle       =   0  'Solid
            Height          =   495
            Index           =   0
            Left            =   2625
            Top             =   210
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Shape shpBottom 
            BorderColor     =   &H8000000D&
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00FF0000&
            FillStyle       =   0  'Solid
            Height          =   44
            Index           =   0
            Left            =   2640
            Top             =   660
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Shape shpTop 
            BorderColor     =   &H00FF8080&
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00FF0000&
            FillStyle       =   0  'Solid
            Height          =   44
            Index           =   0
            Left            =   2625
            Top             =   210
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lbl1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "①"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   7.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   150
            Index           =   0
            Left            =   840
            TabIndex        =   14
            Top             =   405
            Width           =   150
         End
         Begin VB.Label lbl2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "②"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   255
            TabIndex        =   13
            Top             =   120
            Width           =   345
         End
         Begin VB.Label lbl3 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "③"
            Height          =   180
            Index           =   0
            Left            =   270
            TabIndex        =   12
            Top             =   600
            Width           =   195
         End
         Begin VB.Label lbl4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "④"
            Height          =   180
            Index           =   0
            Left            =   720
            TabIndex        =   11
            Top             =   120
            Width           =   180
         End
         Begin VB.Label lbl5 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "⑤"
            Height          =   180
            Index           =   0
            Left            =   1290
            TabIndex        =   10
            Top             =   120
            Width           =   195
         End
         Begin VB.Label lbl6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "⑥"
            Height          =   175
            Index           =   0
            Left            =   1995
            TabIndex        =   9
            Top             =   105
            Width           =   180
         End
         Begin VB.Label lbl7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "⑦"
            Height          =   180
            Index           =   0
            Left            =   720
            TabIndex        =   8
            Top             =   600
            Width           =   180
         End
         Begin VB.Label lbl8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "⑧"
            Height          =   180
            Index           =   0
            Left            =   1875
            TabIndex        =   7
            Top             =   600
            Width           =   180
         End
         Begin VB.Image ImgCard 
            Height          =   255
            Index           =   0
            Left            =   120
            Stretch         =   -1  'True
            Top             =   0
            Width           =   135
         End
      End
      Begin VB.VScrollBar VS1 
         Height          =   840
         Left            =   100
         TabIndex        =   3
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.PictureBox Pic3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      ScaleHeight     =   585
      ScaleWidth      =   3465
      TabIndex        =   0
      Top             =   360
      Width           =   3495
      Begin VB.CheckBox chkFilter 
         Height          =   375
         Left            =   3000
         Picture         =   "usrCardPeople.ctx":752A
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   375
      End
      Begin zlIDKind.PatiIdentify pi1 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"usrCardPeople.ctx":84A4
         BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindAppearance=   0
         ShowSortName    =   -1  'True
         DefaultCardType =   "就诊卡"
         IDKindWidth     =   555
         BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AllowAutoCommCard=   -1  'True
         NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
      End
   End
End
Attribute VB_Name = "usrCardPeople"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngModule As Long
Private mlngCount As Long '表示病人的个数，会在查询病人或过滤病人后改变
Private mblnFilter As Boolean '判断数据是否经过了过滤，如果经过了过滤就显示过滤的数据，如果没有过滤则显示最初的出局，这个标志不用置false，因为每次大过滤都会重置数据
Private mstrFilterName As String '过滤条件名
Private mArr规则 '存储用户自定义的各个控件应该存放数据的标题，也相当于一种规则限定，所以命名为规则
Private mRsBR As ADODB.Recordset '存放当前界面上要显示的数据
Private mRsAll As ADODB.Recordset '存放传入的所有的数据
Private mstrReturn As String '点击选项卡后的返回值，是由选项卡上1~8号label的值，通过"|"分割组成的字符串
Private mrsReturn As ADODB.Recordset
Private mlngSelTab As Long
Private m_CanCheck As Boolean
Private m_def_CanCheck As Boolean
Private mstrLocalID As String
Private mlngLocalIDNum As Long
Private mblnInit As Boolean
Private mstrPIText As String '存放上一次PI1.text中的内容
Private mblnFineseSearch As Boolean
Private mblnNewSearch As Boolean '表示重新开始查询一个查询
Private mstrCardNo As String
Private mlngPatiID As Long
Private mstrFindKey As String
Private mImgList As Object
Private mdblVS系数 As Double '存放(所有卡片控件总和的高度/10000）这个系数,当所有卡片控件总和的高度>10000的情况下使用。
'*\CardChanged事件，点击不同的选项卡时会响应该事件，多用于获取控件的返回值mstrReturn，无其他功能。
Public Event CardChanged() '每次变更选项卡时的事件，有助于获取选中选项卡中的数据
Public Event AfterPatiFind(ByVal strIDKindstr As String, ByVal strValue As String, ByVal blnNext As Boolean, blnfind As Boolean)  '查找的IDKindStr不存卡片上，则返回事件有调整程序处理
Public Event CodeFilter(ByVal strCode As String)
Private mbln初始化 As Boolean '卡片是否已经加载
Private mlng页数 As Long '通过传入数据的个数，来确定页数
Private mColRs As New Collection '如果记录集中的数量大于50，就要使用集合将记录集中的数据分开。
Private marrFilter '存放过滤后的数据
Private mstrOldPiText As String '存放旧的查询条件数据
Private mblnFilterClick As Boolean '是否点击过滤按钮
'Public Event GetChecked() '按钮点击事件
'API将颜色代码转化为颜色
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Sub ShowPeople(Optional ByVal rsBR As ADODB.Recordset, Optional blnSelFree As Boolean = False, Optional blnFilterRefresh As Boolean = False)
    '功能：调用该控件的方法，能够未控件提供初始的过滤条件等
    '参数：rsBR要显示的数据源（数据源中要存在ID，返回值中会返回ID号，返回id是为了方便用户查询）
    '       blnSelFree为false表示控件自动根据以前的数据定位，true则控件不会自动定位数据
    Dim lngi As Long
    Dim lngselnum As Long
    Dim rsbrcopy As ADODB.Recordset
    ReDim marrFilter(0 To 0)
    If mblnInit = False Then Exit Sub '未初始化则跳出
    
    mdblVS系数 = 1 '系数为1，表示不变,用于滚动条定位
    mstrReturn = ""
    '初始化页数和初始载入数据。
    If rsBR Is Nothing Then
        pic2(0).Visible = False
        pi1.Enabled = False
        Exit Sub
    End If
    mblnFilter = False
    Set rsbrcopy = rsBR
'    Set mRsAll = rsBR
    If Not blnFilterRefresh Then Call CopyRecord(rsBR, mRsAll)
    mlngCount = rsbrcopy.RecordCount '传入记录集数据的个数
    mlng页数 = Fix(mlngCount / 50) + IIf(mlngCount Mod 50 = 0, 0, 1) '可以划分的页数
    
    '根据情况确定要显示的部分内容
    If mlngCount <= 50 Then
        Set mRsBR = rsbrcopy
        pic10.Visible = False
    Else
        splitRsToCol rsbrcopy
        Set mRsBR = mColRs("1页")
        mlngCount = mRsBR.RecordCount
        pic10.Visible = True
    End If
    
    lbl10.Caption = "共" & mlng页数 & "页"
    TXT10.Text = 1
    
    UserControl_Resize '这个resize是为了根据页数确定是否显示下面的跳页信息。
    
    '初始化页面控件内容状态
    Call ExecuteCommand("清除控件") '初始化用户控件时将选项卡清空
    Call ExecuteCommand("初始控件")
    
    '********************************************************************
    '如果有两个页面需要来回切换的话，下面的操作可以记录切换前数据的状态。
    '但其实本控件只需要提供选中函数(setCardFocus)，供调用页面进行操作，本身不用
    '做下面的操作，不过为了兼容以前的程序，以下内容保留。
    '********************************************************************
    mlngLocalIDNum = -1 'id所在位置

    For lngi = 0 To UBound(mArr规则) 'id只能有一个
        If UCase(mArr规则(lngi, 0)) = "ID" Then
            mlngLocalIDNum = (lngi + 1) * 2 - 1
            Exit For
        End If
        If UCase(mArr规则(lngi, 1)) = "ID" Then
            mlngLocalIDNum = (lngi + 1) * 2
            Exit For
        End If
    Next
    
    lngselnum = -1
    If mstrLocalID <> "" And mRsBR.RecordCount > 0 And mlngLocalIDNum >= 0 Then '规则中有ID且mrsbr中有数据且有旧ID则取选项卡号
        For lngi = 0 To mRsBR.RecordCount - 1
            If mRsBR.Fields("ID").Value = mstrLocalID Then
                lngselnum = lngi
                Exit For
            End If
            mRsBR.MoveNext
        Next
        mRsBR.MoveFirst
    End If

    If mstrLocalID = "" Or lngselnum = -1 Or mlngLocalIDNum < 0 Or mlngCount = 0 Then
        RaiseEvent CardChanged
    Else
        If blnSelFree = False Then
            Call SelectPeopleCard(lngselnum)
        End If
    End If
End Sub

Private Sub splitRsToCol(rs As ADODB.Recordset)
    '功能：将记录集中的数据按照50条一组的原则分组，并放到集合中
    Dim lngi As Long
    Dim lngj As Long
    Dim lngk As Long
    Dim lngCount As Long
    Dim lngPage As Long
    Dim rsCopy As ADODB.Recordset
    Dim ArrRs()
    ReDim ArrRs(0 To 0)

    If rs Is Nothing Then Exit Sub
    Call CopyRecord(rs, rsCopy)

    lngCount = rsCopy.RecordCount
    Set mColRs = Nothing
    rsCopy.PageSize = 50 '50个数据一组划分记录集
    lngPage = rsCopy.PageCount '存储分组页数
    
    For lngi = 1 To lngPage '动态创建记录集数组
        ReDim Preserve ArrRs(0 To UBound(ArrRs) + 1)
        Set ArrRs(UBound(ArrRs)) = New ADODB.Recordset
        Call RsTitelCopy(rsCopy, ArrRs(UBound(ArrRs)))
    Next
    
    rsCopy.MoveFirst
    
    For lngi = 1 To lngPage '组合集合
        rsCopy.AbsolutePage = lngi
        For lngj = 1 To rsCopy.PageSize
            If rsCopy.EOF Then Exit For
            
            ArrRs(lngi).AddNew
                
            For lngk = 0 To rsCopy.Fields.Count - 1
                ArrRs(lngi).Fields(lngk).Value = rsCopy.Fields(lngk).Value
            Next
            
            ArrRs(lngi).Update
            
            rsCopy.MoveNext
        Next
        ArrRs(lngi).MoveFirst
        mColRs.Add ArrRs(lngi), lngi & "页"
    Next
    
End Sub

Private Function GetValue(lngnum As Long, Index As Integer) As String
    '功能：返回指定index选项卡上的指定控件上的内容
    Dim lngi As Long
    Select Case lngnum
        Case 1
            GetValue = lbl1(Index).Caption & ""
        Case 2
            GetValue = lbl1(Index).Tag & ""
        Case 3
            GetValue = lbl2(Index).Caption & ""
        Case 4
            GetValue = lbl2(Index).Tag & ""
        Case 5
            GetValue = lbl3(Index).Caption & ""
        Case 6
            GetValue = lbl3(Index).Tag & ""
        Case 7
            GetValue = lbl4(Index).Caption & ""
        Case 8
            GetValue = lbl4(Index).Tag & ""
        Case 9
            GetValue = lbl5(Index).Caption & ""
        Case 10
            GetValue = lbl5(Index).Tag & ""
        Case 11
            GetValue = lbl6(Index).Caption & ""
        Case 12
            GetValue = lbl6(Index).Tag & ""
        Case 13
            GetValue = lbl7(Index).Caption & ""
        Case 14
            GetValue = lbl7(Index).Tag & ""
        Case 15
            GetValue = lbl8(Index).Caption & ""
        Case 16
            GetValue = lbl8(Index).Tag & ""
    End Select
End Function

'鼠标滚动所需对象属性
Public Property Get objPicBack() As PictureBox
    Set objPicBack = pic1
End Property

Public Property Get FScrollBar() As VScrollBar
    Set FScrollBar = VS1
End Property

Public Property Let FindStart(newFindStart As Boolean)
    '功能：用户切换页面后，要重新初始化查询，由于如果使用一个本控件，所有查询都是通用的，
    '      如果在其他页面查询后调到另一个页面，mblnFineseSearch不会改变，也就是说程序会默认已查询完毕，
    '      这时将无法进行多人次的查询。注：由于查询部分变化，FindStart不再有效，只是为了兼容以前的程序，这里保留
    mblnNewSearch = newFindStart
    pi1.Text = ""
End Property

Public Property Let locked(blnlocked As Boolean)
    pic1.Enabled = Not blnlocked
    pic3.Enabled = Not blnlocked
    pic10.Enabled = Not blnlocked
End Property

Public Property Get strReturn() As String
    strReturn = mstrReturn
End Property

Public Property Get rsReturn() As ADODB.Recordset
    Set rsReturn = mrsReturn
End Property
Public Property Get CanCheck() As Boolean
    CanCheck = m_CanCheck
    Call pic1_Resize
    Call UserControl_Resize
    Call pic1_Resize
End Property
Public Property Let CanCheck(newCanCheck As Boolean)
    Dim lngi As Long
    m_CanCheck = newCanCheck
    For lngi = 0 To chk1.Count - 1
        chk1(lngi).Value = 0
    Next
    Call pic1_Resize
    Call UserControl_Resize
    Call pic1_Resize
End Property

Public Sub FilterRefreshByCode(rs As Recordset)
    Dim strFilter As String
    
    If rs.State = adStateClosed Then Exit Sub
    If rs.RecordCount = 0 Then Exit Sub
    rs.MoveFirst
    Do While Not rs.EOF
        strFilter = " or id = " & rs!id & strFilter
        rs.MoveNext
    Loop
    mRsAll.Filter = Mid(strFilter, 5)
End Sub

Private Sub FilterRefresh()
    Dim rs As New Recordset
    Dim strPatiID As String
'-------------------------------
    On Error GoTo errH
    
    If chkFilter.Value <> 1 Or pi1.Text = "" Then
        mRsAll.Filter = ""
    Else
        If mstrFilterName = "姓名" Then
            mRsAll.Filter = "姓名 like '" & pi1.Text & "%'"
        ElseIf mstrFilterName <> "姓名" And mstrFilterName <> "血袋编号" Then     'And mRsAll.Fields("ID") > 0
            If mlngPatiID <> 0 Then mRsAll.Filter = "病人ID = " & mlngPatiID
        ElseIf mstrFilterName = "血袋编号" Then
            RaiseEvent CodeFilter(pi1.Text)
        End If
    End If
    mstrOldPiText = mstrFilterName & "-" & pi1.Text
    mblnFilterClick = False
    Call ShowPeople(mRsAll, False, True)
    Exit Sub
errH:
    If Err.Number = 0 Then
        Resume Next
    End If
End Sub

Private Sub chkFilter_Click()
    mblnFilterClick = True
    If chkFilter.Value = 1 And mstrFilterName <> "姓名" And mstrFilterName <> "血袋编号" Then
        pi1.SetFocus
        Call gobjCommFun.PressKey(vbKeyReturn)
    Else
        Call FilterRefresh
    End If
End Sub

Private Sub lbl1_DblClick(Index As Integer)
    chk1(Index).Value = IIf(chk1(Index).Value = 1, 0, 1)
End Sub

Private Sub lbl1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    Set picFont.Font = lbl1(Index).Font
    If lbl1(Index).Width < picFont.TextWidth(lbl1(Index).Caption) Then
        strInfo = lbl1(Index).Caption
    End If
    Call gobjCommFun.ShowTipInfo(pic2(Index).hWnd, strInfo, True)
End Sub

Private Sub lbl11_Click()
    Dim lngTXT As Long

    lngTXT = Val(TXT10.Text) - 1
    If lngTXT > 0 Then
        TXT10.Text = lngTXT
        SetPage lngTXT
    End If
End Sub

Private Sub lbl12_Click()
    Dim lngTXT As Long

    lngTXT = Val(TXT10.Text) + 1
    If lngTXT <= mlng页数 Then
        TXT10.Text = lngTXT
        SetPage lngTXT
    End If
End Sub

Private Sub lbl2_DblClick(Index As Integer)
    chk1(Index).Value = IIf(chk1(Index).Value = 1, 0, 1)
End Sub

Private Sub lbl3_DblClick(Index As Integer)
    chk1(Index).Value = IIf(chk1(Index).Value = 1, 0, 1)
End Sub

Private Sub lbl3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    Set picFont.Font = lbl3(Index).Font
    If lbl3(Index).Width < picFont.TextWidth(lbl3(Index).Caption) Then
        strInfo = lbl3(Index).Caption
    End If
    Call gobjCommFun.ShowTipInfo(pic2(Index).hWnd, strInfo, True)
End Sub

Private Sub lbl4_DblClick(Index As Integer)
    chk1(Index).Value = IIf(chk1(Index).Value = 1, 0, 1)
End Sub

Private Sub lbl4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    Set picFont.Font = lbl4(Index).Font
    If lbl4(Index).Width < picFont.TextWidth(lbl4(Index).Caption) Then
        strInfo = lbl4(Index).Caption
    End If
    Call gobjCommFun.ShowTipInfo(pic2(Index).hWnd, strInfo, True)
End Sub

Private Sub lbl5_DblClick(Index As Integer)
    chk1(Index).Value = IIf(chk1(Index).Value = 1, 0, 1)
End Sub

Private Sub lbl5_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    Set picFont.Font = lbl5(Index).Font
    If lbl5(Index).Width < picFont.TextWidth(lbl5(Index).Caption) Then
        strInfo = lbl5(Index).Caption
    End If
    Call gobjCommFun.ShowTipInfo(pic2(Index).hWnd, strInfo, True)
End Sub

Private Sub lbl6_DblClick(Index As Integer)
    chk1(Index).Value = IIf(chk1(Index).Value = 1, 0, 1)
    pic4(Index).Tag = chk1(Index).Value
End Sub

Private Sub lbl6_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    Set picFont.Font = lbl6(Index).Font
    If lbl6(Index).Width < picFont.TextWidth(lbl6(Index).Caption) Then
        strInfo = lbl6(Index).Caption
    End If
    Call gobjCommFun.ShowTipInfo(pic2(Index).hWnd, strInfo, True)
End Sub

Private Sub lbl7_DblClick(Index As Integer)
    chk1(Index).Value = IIf(chk1(Index).Value = 1, 0, 1)
End Sub

Private Sub lbl7_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    Set picFont.Font = lbl7(Index).Font
    If lbl7(Index).Width < picFont.TextWidth(lbl7(Index).Caption) Then
        strInfo = lbl7(Index).Caption
    End If
    Call gobjCommFun.ShowTipInfo(pic2(Index).hWnd, strInfo, True)
End Sub

Private Sub lbl8_DblClick(Index As Integer)
    chk1(Index).Value = IIf(chk1(Index).Value = 1, 0, 1)
End Sub

Private Sub lbl8_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    Set picFont.Font = lbl8(Index).Font
    If lbl8(Index).Width < picFont.TextWidth(lbl8(Index).Caption) Then
        strInfo = lbl8(Index).Caption
    End If
    Call gobjCommFun.ShowTipInfo(pic2(Index).hWnd, strInfo, True)
End Sub

Private Sub SetPage(lngPage As Long)
    '根据传入参数加载第lngPage页
    Dim strcbo As String
    
    TXT10.Text = lngPage
    strcbo = lngPage & "页"
    Set mRsBR = mColRs(strcbo)
    mlngCount = mRsBR.RecordCount
    Call ExecuteCommand("清除控件")
    Call ExecuteCommand("初始控件")
End Sub

Private Sub pi1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0 '"'"单引号输入无效
End Sub

Private Sub pic10_Resize()
    On Error GoTo Errorhand
    lbl10.Move pic10.ScaleLeft, pic10.ScaleTop + 30
    lbl11.Move (pic10.Width - lbl10.Width - TXT10.Width - lbl12.Width - lbl11.Width) / 2 + lbl10.Left + lbl10.Width, lbl10.Top
    TXT10.Move lbl11.Left + lbl11.Width, pic10.ScaleTop
    frm10.Move TXT10.Left, TXT10.Top + TXT10.Height, TXT10.Width
    lbl12.Move TXT10.Left + TXT10.Width, lbl10.Top
Errorhand:
End Sub

Private Sub pic2_DblClick(Index As Integer)
    chk1(Index).Value = IIf(chk1(Index).Value = 1, 0, 1)
End Sub

Private Sub PI1_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    '功能：读取卡号、查找数据
    If objHisPati Is Nothing Then
        mlngPatiID = 0
    Else
        mlngPatiID = objHisPati.病人id
    End If
    mstrCardNo = ""
    If blnCard = True And Not objCardData Is Nothing Then mstrCardNo = objCardData.卡号 '如果有卡且卡中数据不为空则将卡号赋值给mstrCardNo
    If mblnFilterClick Or mstrOldPiText <> mstrFilterName & "-" & pi1.Text And chkFilter.Value = 1 Then    '上次与本次相同时不需要根据过滤条件刷新记录集
        Call FilterRefresh
    End If
    Call FindPatiCard(False, True)
End Sub

Public Sub FindPatiByVbKey(Optional ByVal blnNext As Boolean)
    'blnNext=false 表示定位查找：True表示开始查找或查找下一个
    If blnNext = True Then '查找下一个
        If pi1.Text = "" Then
            If pi1.Enabled And pi1.Visible Then pi1.SetFocus
        Else
            Call FindPatiCard(True)
        End If
    Else  '查找：定位到当前控件
        If UserControl.ActiveControl Is pi1 Then
            pi1.SetFocus '有时需要定位一下
            If pi1.Text <> "" Then
                Call FindPatiCard
            End If
        Else
            pi1.SetFocus
        End If
    End If
End Sub

Private Sub FindPatiCard(Optional ByVal blnNext As Boolean, Optional ByVal blnPati As Boolean = True)
    '查找病人,通过快捷键直接
    Dim blnfind As Boolean
    If pi1.Enabled = False Then Exit Sub
    If pi1.Text = "" Then mlngPatiID = 0: Exit Sub
    
    If mlngPatiID <> 0 Then blnfind = findIdPeoPle(CStr(mlngPatiID), blnPati, blnNext) '通过id查询病人
    If pi1.Text <> "" And blnfind = False Then blnfind = FindPati(pi1.Text, blnNext) '查找数据
    If blnfind = False Then
        RaiseEvent AfterPatiFind(mstrFilterName, pi1.Text, blnNext, blnfind)
        If blnfind = False Then
            MsgBox "没有找到符合条件的数据！", vbInformation, gstrSysName
        End If
    End If
End Sub

Private Function FindPati(ByVal strValue As String, Optional blnNext As Boolean = True) As Boolean
    '功能：查找符合条件的数据并定位
    Dim lngi As Long
    Dim lngj As Long
    Dim blnisend As Boolean
    Dim blnTitle As Boolean
    Dim lngFind As Long '从记录集中找到的数据的位置
    Dim lngselnum As Long 'lngfind处理过后的能够定位到卡片的位置
    Dim lng所在页 As Long '查询到的卡片所在的页数
    Dim rsData As ADODB.Recordset
    Dim bln病历号 As Boolean, strCardName As String
    
    If strValue = "" Then Exit Function
    strCardName = mstrFilterName
    '判断查询条件在mArr规则中是否存在
    For lngi = 0 To UBound(mArr规则)
        If mArr规则(lngi, 0) = strCardName Or mArr规则(lngi, 1) = strCardName Then
            blnTitle = True
            Exit For
        End If
        If mArr规则(lngi, 0) = "病历号" Or mArr规则(lngi, 1) = "病历号" Then
            bln病历号 = True
        End If
    Next
    If blnTitle = False Then
        If (strCardName = "住院号" Or strCardName = "门诊号") And bln病历号 = True Then
            strCardName = "病历号"
        Else
            Exit Function '如果查询条件在规则中不存在则直接退出。
        End If
    End If
    
    If Not mRsAll.EOF Then
        mRsAll.MoveFirst
    Else
        Exit Function
    End If
    
    CopyRecord mRsAll, rsData
    
    If rsData Is Nothing Then Exit Function '如果mrsbr本身没有数据时，直接跳出查找
    
   If blnNext = False Then
        lngselnum = -1

    Else
        If pic10.Visible = True Then
            lngselnum = (TXT10.Text - 1) * 50 + mlngSelTab + 1
        Else
            lngselnum = mlngSelTab + 1
        End If
    End If
    
     '如果查询条件没有改变，则不需要从新过滤数据，直接使用上次过滤得到的数组
    If mstrOldPiText = strCardName & "-" & strValue And UBound(marrFilter) > 0 Then
        setPosition marrFilter, lngselnum
        FindPati = True
        Exit Function
    End If

    ReDim marrFilter(0 To 0)
    '过滤
    If strCardName = "姓名" Then
        rsData.Filter = "姓名 like '" & strValue & "%'"
    Else
        rsData.Filter = strCardName & IIf(IsNumeric(strValue) = True, "=" & Val(strValue), "='" & strValue & "'")
    End If
    If rsData.RecordCount = 0 Then Exit Function '无法过滤到数据，则直接退出查找
    '定位到查找到的数据
    For lngi = 0 To rsData.RecordCount - 1
        ReDim Preserve marrFilter(UBound(marrFilter) + 1)
        marrFilter(UBound(marrFilter)) = rsData.Bookmark
        rsData.MoveNext
    Next
    setPosition marrFilter, lngselnum, blnNext
    '这里记录之前查找的信息，用于判断查询条件是否改变
    mstrOldPiText = strCardName & "-" & strValue
    FindPati = True
End Function

Private Sub setPosition(arr As Variant, lngnum As Long, Optional blnNext As Boolean = True)
    '查找符合条件的数据，并定位
    Dim lngi As Long
    Dim lngselnum As Long
    Dim blnisend As Boolean
    Dim lng所在页 As Long '查询到的卡片所在的页数
    If UBound(arr) = 0 Then Exit Sub
    
    blnisend = False
    For lngi = 1 To UBound(arr)
        If lngnum < arr(lngi) Then  '如果当前选中数据之后有匹配数据则定位到匹配数据
            lngselnum = arr(lngi)
            blnisend = True
            Exit For
        End If
    Next
    
    If blnisend = False Then
        If blnNext = True Then
            MsgBox "后面已找不到符合条件的病人。", vbInformation, gstrSysName
            Exit Sub
        Else
            lngselnum = arr(1) '如果当前选中数据之后无匹配数据则定位到第一个匹配数据的位置
        End If
    End If
    
    If pic10.Visible = True Then '如果有多页数据，则要将查询到的数据所在位置做一定的处理，还要改变当前页
        lng所在页 = Fix((lngselnum - 1) / 50) + IIf(mRsAll.RecordCount Mod 50 = 0, 0, 1)
        lngselnum = (lngselnum - 1) Mod 50
        If lng所在页 <> Val(TXT10.Text) Then
            SetPage lng所在页
        End If
    Else
        lngselnum = lngselnum - 1
    End If
    
    Call SelectPeopleCard(lngselnum)
End Sub

Public Function findIdPeoPle(strKey As String, Optional ByVal blnPatiID As Boolean = True, Optional blnNext As Boolean = True) As Boolean
'功能：查询操作的简化版，是在知道病人id的情况下，直接通过病人id查询数据,FindPati并不能通过病人id进行查询。
'lngID：病人ID（一般是控件通过查询控件内部查找使用）；key(主键，其他窗体调用用于定位病人)
    Dim rsData As ADODB.Recordset
    Dim arrFilter
    Dim lngselnum As Long
    Dim lngi As Long
    
    CopyRecord mRsAll, rsData
    
    If rsData Is Nothing Then Exit Function '如果mrsbr本身没有数据时，直接跳出查找
    
    If blnNext = False Then
        lngselnum = -1
    Else
        If pic10.Visible = True Then
            lngselnum = (TXT10.Text - 1) * 50 + mlngSelTab + 1
        Else
            lngselnum = mlngSelTab + 1
        End If
    End If
    
    If blnPatiID = True Then
        rsData.Filter = "病人ID=" & Val(strKey)
    Else
        If IsNumeric(strKey) Then
            rsData.Filter = "ID=" & Val(strKey)
        Else
            rsData.Filter = "ID='" & strKey & "'"
        End If
    End If

    If rsData.RecordCount = 0 Then findIdPeoPle = False: Exit Function '无法过滤到数据，则直接退出查找
    ReDim arrFilter(0 To 0)
    '定位到查找到的数据
    For lngi = 0 To rsData.RecordCount - 1
        ReDim Preserve arrFilter(UBound(arrFilter) + 1)
        arrFilter(UBound(arrFilter)) = rsData.Bookmark
        rsData.MoveNext
    Next
    
    setPosition arrFilter, lngselnum, blnNext
    
    findIdPeoPle = True
End Function

Private Function GetLblValue(Title As String, Index As Integer) As String
    '功能：获取Pic2中相应控件的值,在比较时都会使用ucase将字母转化为大写，以方便操作，避免失误
    '参数：title-根据Title值查询是哪个控件，index-根据title查询到的控件的index
    Dim lngi As Long
    
    If mblnInit = False Then Exit Function
    
    If UBound(mArr规则) <> 0 Then
        For lngi = 0 To UBound(mArr规则)
            If UCase(mArr规则(lngi, 0)) = UCase(Title) Then
                GetLblValue = GetValue((lngi + 1) * 2 - 1, Index)
                Exit For
            End If
            If UCase(mArr规则(lngi, 1)) = UCase(Title) Then
                GetLblValue = GetValue((lngi + 1) * 2, Index)
                Exit For
            End If
        Next
    End If
End Function

Public Sub SetPIFocus()
    '将焦点聚焦到PI1上
    If pi1.Enabled And pi1.Visible Then pi1.SetFocus
End Sub

Private Sub pic2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call gobjCommFun.ShowTipInfo(pic2(Index).hWnd, "")
End Sub

Private Sub TXT10_KeyPress(KeyAscii As Integer)
    Dim strKey As String
    Dim strTXT As String
    strKey = Chr(KeyAscii)
    If Not IsNumeric(strKey) And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then KeyAscii = 0: Exit Sub '非数字、回车、退格则退出
    If KeyAscii = vbKeyReturn Then
        strTXT = TXT10.Text
        If Val(strTXT) < 1 Or Val(strTXT) > mlng页数 Then Exit Sub  '非正确页数跳出
        SetPage Val(strTXT)
    End If
End Sub

Private Sub TXT10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    gobjCommFun.ShowTipInfo TXT10.hWnd, "显示当前页数，可以输入指定页数，并按回车跳转到指定页"
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    pi1.ActiveFastKey
End Sub

Private Sub UserControl_Terminate()
    Call ExecuteCommand("清除控件")
    If Not mColRs Is Nothing Then Set mColRs = Nothing
    If Not mRsBR Is Nothing Then Set mRsBR = Nothing
    If Not mRsAll Is Nothing Then Set mRsAll = Nothing
    If Not mrsReturn Is Nothing Then Set mrsReturn = Nothing
    Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\" & "病人列表_" & mlngModule, "检索定位", mstrFilterName)
    mstrLocalID = ""
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '功能：保存相关用户定义属性数据
    Call PropBag.WriteProperty("CanCheck", m_CanCheck, m_def_CanCheck)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_CanCheck = PropBag.ReadProperty("CanCheck", m_def_CanCheck)
End Sub

Public Sub UserInit(ByVal frmMain As Object, str规则 As String, Optional imgList As Object, Optional ByVal lngModule As Long = 0, Optional ByVal strIDKindstr As String = "")
    '功能：初始化每个控件的内容和属性颜色字体等
    '参数：str规则：由“a|b|c|d|e|f|g|h;a|b|c|d|e|f|g|h;.....”组成,一共9组，多余9组会忽略后面的几组，
    '               这9组数据分别代表pic2界面上的1~8个label控件组和一个image控件组
    '      a|b|c|d|e|f|g|h：显示数据|辅助数据|是否显示|字体|字号|字体颜色|背景颜色|图标;
    '      显示数据：字符，显示在界面上的数据
    '      辅助数据：字符，存放在控件的tag中的数据
    '      是否显示：数字，0表示显示（默认），1表示不显示,用于imgCard控件：0表示方式一取图，1表示方式二取图
    '      字体：字符或数字，字符的字体为空或者为0表示默认,字体中的内容为"宋体""隶书"等,对imgCard无效
    '      字号：数字，字体的大小,对imgCard无效
    '      字体颜色：数字，字体的颜色，为空或者0表示默认黑色，用RGB转换后的数字表示,对imgCard无效
    '      背景颜色：数字，控件的背景颜色，为空或者0表示默认控件颜色，用RGB转换后的数字表示,对imgCard无效
    '      图标：只对imgCard有效，保存图片在imgList中的编号
    '            注：方式一表示所有人员卡片都取图标中编号的图片，方式二表示每个人员卡片单独取传入记录集中图标字段中编号的图片
    '               可以理解为方式一所有人员的图片都一样，方式二可以根据不同需求改变图片
    '      imgList：如果要添加图标，需要提供图标来源的imglist，图标都是根据imglist来取用，目前只支持imglist
    Dim ArrS
    Dim ArrM
    Dim lngi As Long
    Dim lngj As Long
    Dim strCardName As String
    
    On Error GoTo ErrHand
    mlngModule = lngModule
    Set mImgList = imgList
    strCardName = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\" & "病人列表_" & lngModule, "检索定位", "姓名")
    '初始化Patidentify控件
    Call CreateSquareCardObject(frmMain, 2200, lngModule)
    '如果未传入IDKindStr将使用默认的
    If strIDKindstr = "" Then strIDKindstr = "姓|姓名|0|0|0|0|0|0;住|住院号|0|0|0|0|0|0;门|门诊号|0|0|0|0|0|0;就|就诊卡|0|0|8|0|0|0;身|二代身份证|0|0|0|0|0|0;IC|IC卡|1|0|0|0|0|0"
    If Not gobjCardSquare Is Nothing Then
        strIDKindstr = gobjCardSquare.zlGetIDKindStr(strIDKindstr)
    End If
    '这个对象传入Nothing,传入主窗体，主窗体关闭时会触发active事件（应该是多次刷多次调用该方法的问题）
    Call pi1.zlInit(Nothing, 2200, , gcnOracle, gstrDBUser, gobjCardSquare, strIDKindstr)
    pi1.FindPatiShowName = False
    pi1.IDKindIDX = pi1.GetKindIndex(strCardName)
    pi1.AutoSize = True
'    PI1.ShowPropertySet = True
    pi1.objIDKind.AllowAutoICCard = True
    pi1.objIDKind.AllowAutoIDCard = True
    
    ArrS = Split(str规则, ";")
    ReDim mArr规则(0 To 8, 0 To 7) 'mArr规则对应9个控件，每个控件的8个属性，这个是固定死的
    For lngi = 0 To UBound(ArrS)
        If lngi > 8 Then Exit For
        ArrM = Split(ArrS(lngi), "|")
        For lngj = 0 To UBound(ArrM)
            If lngj > 7 Then Exit For
            mArr规则(lngi, lngj) = ArrM(lngj)
        Next
    Next
    '在初始化规则后，首先就要改变控件的字体等属性
    SetLabelProper
    mblnInit = True
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function SetLabelProper()
    Dim objFont As StdFont
    Set objFont = picFont.Font
    '初始化pic2上的控件组
    lbl1(0).FontName = IIf(mArr规则(0, 3) & "" <> "", mArr规则(0, 3), "宋体") '字体
    lbl1(0).FontSize = IIf(Val(mArr规则(0, 4) & "") <> 0, Val(mArr规则(0, 4)), 9) '字号
    lbl1(0).ForeColor = IIf(Val(mArr规则(0, 5) & "") <> 0, GetRGBFromOLEColor(Val(mArr规则(0, 5))), &H0&)        '字体颜色
    lbl1(0).BackColor = IIf(Val(mArr规则(0, 6) & "") <> 0, GetRGBFromOLEColor(Val(mArr规则(0, 6))), &HFFFFFF) '控件背景颜色
    Set picFont.Font = lbl1(0).Font
    lbl1(0).Height = picFont.TextHeight("刘") '设置高度
    
    lbl2(0).FontName = IIf(mArr规则(1, 3) & "" <> "", mArr规则(1, 3), "宋体") '字体
    lbl2(0).FontSize = IIf(Val(mArr规则(1, 4) & "") <> 0, Val(mArr规则(1, 4)), 16) '字号
    lbl2(0).ForeColor = IIf(Val(mArr规则(1, 5) & "") <> 0, GetRGBFromOLEColor(Val(mArr规则(1, 5))), &H0&)        '字体颜色
    lbl2(0).BackColor = IIf(Val(mArr规则(1, 6) & "") <> 0, GetRGBFromOLEColor(Val(mArr规则(1, 6))), &HFFFFFF) '控件背景颜色

    lbl3(0).FontName = IIf(mArr规则(2, 3) & "" <> "", mArr规则(2, 3), "宋体") '字体
    lbl3(0).FontSize = IIf(Val(mArr规则(2, 4) & "") <> 0, Val(mArr规则(2, 4)), 9) '字号
    lbl3(0).ForeColor = IIf(Val(mArr规则(2, 5) & "") <> 0, GetRGBFromOLEColor(Val(mArr规则(2, 5))), &H0&)        '字体颜色
    lbl3(0).BackColor = IIf(Val(mArr规则(2, 6) & "") <> 0, GetRGBFromOLEColor(Val(mArr规则(2, 6))), &HFFFFFF) '控件背景颜色
    Set picFont.Font = lbl3(0).Font
    lbl3(0).Height = picFont.TextHeight("刘") '设置高度

    lbl4(0).FontName = IIf(mArr规则(3, 3) & "" <> "", mArr规则(3, 3), "宋体") '字体
    lbl4(0).FontSize = IIf(Val(mArr规则(3, 4) & "") <> 0, Val(mArr规则(3, 4)), 9) '字号
    lbl4(0).ForeColor = IIf(Val(mArr规则(3, 5) & "") <> 0, GetRGBFromOLEColor(Val(mArr规则(3, 5))), &H0&)        '字体颜色
    lbl4(0).BackColor = IIf(Val(mArr规则(3, 6) & "") <> 0, GetRGBFromOLEColor(Val(mArr规则(3, 6))), &HFFFFFF) '控件背景颜色
    Set picFont.Font = lbl4(0).Font
    lbl4(0).Height = picFont.TextHeight("刘") '设置高度
    
    lbl5(0).FontName = IIf(mArr规则(4, 3) & "" <> "", mArr规则(4, 3), "宋体") '字体
    lbl5(0).FontSize = IIf(Val(mArr规则(4, 4) & "") <> 0, Val(mArr规则(4, 4)), 9) '字号
    lbl5(0).ForeColor = IIf(Val(mArr规则(4, 5) & "") <> 0, GetRGBFromOLEColor(Val(mArr规则(4, 5))), &H0&)        '字体颜色
    lbl5(0).BackColor = IIf(Val(mArr规则(4, 6) & "") <> 0, GetRGBFromOLEColor(Val(mArr规则(4, 6))), &HFFFFFF) '控件背景颜色
    Set picFont.Font = lbl5(0).Font
    lbl5(0).Height = picFont.TextHeight("刘") '设置高度
    
    lbl6(0).FontName = IIf(mArr规则(5, 3) & "" <> "", mArr规则(5, 3), "宋体") '字体
    lbl6(0).FontSize = IIf(Val(mArr规则(5, 4) & "") <> 0, Val(mArr规则(5, 4)), 9) '字号
    lbl6(0).ForeColor = IIf(Val(mArr规则(5, 5) & "") <> 0, GetRGBFromOLEColor(Val(mArr规则(5, 5))), &H0&)        '字体颜色
    lbl6(0).BackColor = IIf(Val(mArr规则(5, 6) & "") <> 0, GetRGBFromOLEColor(Val(mArr规则(5, 6))), &HFFFFFF) '控件背景颜色
    Set picFont.Font = lbl6(0).Font
    lbl6(0).Height = picFont.TextHeight("刘") '设置高度

    lbl7(0).FontName = IIf(mArr规则(6, 3) & "" <> "", mArr规则(6, 3), "宋体") '字体
    lbl7(0).FontSize = IIf(Val(mArr规则(6, 4) & "") <> 0, Val(mArr规则(6, 4)), 9) '字号
    lbl7(0).ForeColor = IIf(Val(mArr规则(6, 5) & "") <> 0, GetRGBFromOLEColor(Val(mArr规则(6, 5))), &H0&)        '字体颜色
    lbl7(0).BackColor = IIf(Val(mArr规则(6, 6) & "") <> 0, GetRGBFromOLEColor(Val(mArr规则(6, 6))), &HFFFFFF) '控件背景颜色
    Set picFont.Font = lbl7(0).Font
    lbl7(0).Height = picFont.TextHeight("刘") '设置高度

    lbl8(0).FontName = IIf(mArr规则(7, 3) & "" <> "", mArr规则(7, 3), "宋体") '字体
    lbl8(0).FontSize = IIf(Val(mArr规则(7, 4) & "") <> 0, Val(mArr规则(7, 4)), 9) '字号
    lbl8(0).ForeColor = IIf(Val(mArr规则(7, 5) & "") <> 0, GetRGBFromOLEColor(Val(mArr规则(7, 5))), &H0&)        '字体颜色
    lbl8(0).BackColor = IIf(Val(mArr规则(7, 6) & "") <> 0, GetRGBFromOLEColor(Val(mArr规则(7, 6))), &HFFFFFF) '控件背景颜色
    Set picFont.Font = lbl8(0).Font
    lbl8(0).Height = picFont.TextHeight("刘") '设置高度
    
    Set picFont.Font = objFont
    Set ImgCard(0).Picture = Nothing
End Function

Public Function GetCheckedData() As ADODB.Recordset
    '功能：返回多个选中控件的数据，不会返回image中的数据！
    Dim lngi As Long
    Dim strName As String 'mArr规则有可能为空，多个空不能同时作为题目，所以当mArr规则为空时自定义一个题目
    strName = "自定义"
    If pic2(0).Visible = False Then Exit Function '调用界面没有传递参数时，无法提取
    If pic2.Count <= 0 Then Exit Function '如果界面无数据，那么点击按钮是无效得
    If m_CanCheck = False Then Exit Function '如果m_CanCheck = False，无法get数据
    Set mrsReturn = New ADODB.Recordset
    With mrsReturn '初始化rsReturn
        For lngi = 0 To UBound(mArr规则)
            .Fields.Append IIf(mArr规则(lngi, 0) = "", strName & lngi, mArr规则(lngi, 0)), adLongVarChar, 100, adFldIsNullable
            .Fields.Append IIf(mArr规则(lngi, 1) = "", strName & lngi, mArr规则(lngi, 1)), adLongVarChar, 100, adFldIsNullable
        Next
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        For lngi = 0 To pic2.Count - 1
            If Val(pic4(lngi).Tag) = 1 Then
                .AddNew
                .Fields(0).Value = lbl1(lngi).Caption
                .Fields(1).Value = lbl1(lngi).Tag
                .Fields(2).Value = lbl2(lngi).Caption
                .Fields(3).Value = lbl2(lngi).Tag
                .Fields(4).Value = lbl3(lngi).Caption
                .Fields(5).Value = lbl3(lngi).Tag
                .Fields(6).Value = lbl4(lngi).Caption
                .Fields(7).Value = lbl4(lngi).Tag
                .Fields(8).Value = lbl5(lngi).Caption
                .Fields(9).Value = lbl5(lngi).Tag
                .Fields(10).Value = lbl6(lngi).Caption
                .Fields(11).Value = lbl6(lngi).Tag
                .Fields(12).Value = lbl7(lngi).Caption
                .Fields(13).Value = lbl7(lngi).Tag
                .Fields(14).Value = lbl8(lngi).Caption
                .Fields(15).Value = lbl8(lngi).Tag
                .Update
            End If
        Next
        If .RecordCount > 0 Then
            .MoveFirst
        End If
    End With
    Set GetCheckedData = mrsReturn
'    RaiseEvent GetChecked
End Function

'下面这些全都是为了实现点击选择选项卡的目的
Private Sub lbl1_Click(Index As Integer)
    Call SelectPeopleCard(Index, True)
End Sub

Private Sub chk1_Click(Index As Integer)
    pic4(Index).Tag = chk1(Index).Value
End Sub
Private Sub Pic4_Click(Index As Integer)
    chk1(Index).Value = IIf(chk1(Index).Value = 1, 0, 1)
End Sub

Private Sub lbl2_Click(Index As Integer)
    Call SelectPeopleCard(Index, True)
End Sub

Private Sub lbl3_Click(Index As Integer)
    Call SelectPeopleCard(Index, True)
End Sub

Private Sub lbl4_Click(Index As Integer)
    Call SelectPeopleCard(Index, True)
End Sub

Private Sub lbl5_Click(Index As Integer)
    Call SelectPeopleCard(Index, True)
End Sub

Private Sub lbl6_Click(Index As Integer)
    Call SelectPeopleCard(Index, True)
End Sub

Private Sub lbl7_Click(Index As Integer)
    Call SelectPeopleCard(Index, True)
End Sub

Private Sub lbl8_Click(Index As Integer)
    Call SelectPeopleCard(Index, True)
End Sub

Private Sub pi1_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    pi1.ShowPropertySet = True
    If mstrFilterName <> objCard.名称 Then pi1.Text = ""
    mstrFilterName = objCard.名称
End Sub

Private Sub pic1_Resize()
    Dim lngi As Long
    Dim lngVSMAX As Long
    Dim lngHeight As Long
    Dim blnVshVisible As Boolean
    
    '控件加载完成且控件显示才进行位置调整
    If mbln初始化 = False Then Exit Sub

    '如果pic2有变化则更改各个控件的大小位置等
    If mlngCount = 0 Then Exit Sub '如果无数据则不需要对控件进行调整
    
    On Error Resume Next
    Call LockWindowUpdate(UserControl.hWnd)
    lngHeight = (mlngCount - 1) * 950 + 900 - pic1.ScaleHeight   'pic2(mlngCount - 1).Top - pic2(0).Top + pic2(mlngCount - 1).Height + 50 - pic1.ScaleHeight
    If lngHeight > 10000 Then
        mdblVS系数 = lngHeight / 10000
    Else
        mdblVS系数 = 1
    End If
    VS1.Left = pic1.Width - VS1.Width
    VS1.Top = pic1.ScaleTop
    VS1.Height = pic1.ScaleHeight
    If lngHeight >= 0 Then
        VS1.Min = 1
        VS1.Max = IIf(lngHeight > 10000, 10000, lngHeight)
        VS1.Value = 1
        If lngHeight > 950 Then
            VS1.SmallChange = 950 / mdblVS系数
            VS1.LargeChange = 950 / mdblVS系数
        Else
            VS1.SmallChange = lngHeight / mdblVS系数
            VS1.LargeChange = lngHeight / mdblVS系数
        End If
        blnVshVisible = True
        VS1.Visible = True
    Else
        blnVshVisible = False
        VS1.Visible = False
    End If

    For lngi = 0 To mlngCount - 1
        pic2(lngi).Move 100, 950 * lngi + 50, pi1.Width + chkFilter.Width - IIf(blnVshVisible = True, VS1.Width, 0) + 100, 850
        pic2(lngi).Tag = pic2(lngi).Top  '保留每个pic2的原始顶端位置，滚动条要用
        pic2(lngi).AutoRedraw = True
        pic2(lngi).PaintPicture pic2(lngi).Picture, 0, 0, pic2(lngi).ScaleWidth, pic2(lngi).ScaleHeight '加载背景图片
        
        shpLeft(lngi).Move 0, 0, 45, pic2(lngi).Height
        shpRight(lngi).Move pic2(lngi).Width - 45, 0, 45, pic2(lngi).Height
        shpTop(lngi).Move 0, 0, pic2(lngi).Width, 45
        shpBottom(lngi).Move 0, pic2(lngi).Height - 45, pic2(lngi).Width, 45
        
        If m_CanCheck = True Then '显示左边的选择栏
            pic4(lngi).Move 60, (pic2(lngi).Height - pic4(lngi).Height) \ 2
            lbl2(lngi).Move pic4(lngi).Left + pic4(lngi).Width + 30, (pic2(lngi).Height - lbl2(lngi).Height) / 2 - 60

            chk1(lngi).Visible = True
            pic4(lngi).Visible = True
        Else
            lbl2(lngi).Move 60, (pic2(lngi).Height - lbl2(lngi).Height) / 2 - 60
            chk1(lngi).Visible = False
            pic4(lngi).Visible = False
        End If
        If Val(mArr规则(0, 2)) = 1 Then lbl1(lngi).Height = 50  '特殊处理,当只显示颜色时，调整宽度
        lbl1(lngi).Move lbl2(lngi).Left + lbl2(lngi).Width + 45, (pic2(lngi).Height - lbl1(lngi).Height) / 2, pic2(lngi).Width - lbl2(lngi).Left - lbl2(lngi).Width - 45 - 120
        lbl4(lngi).Move lbl1(lngi).Left, 100, (lbl1(lngi).Width - 60) / 3
        lbl5(lngi).Move lbl4(lngi).Left + lbl4(lngi).Width + 30, lbl4(lngi).Top, lbl4(lngi).Width
        lbl6(lngi).Move lbl5(lngi).Left + lbl5(lngi).Width + 30, lbl5(lngi).Top, lbl1(lngi).Width - lbl4(lngi).Width - lbl5(lngi).Width - 60
        
        lbl7(lngi).Move lbl1(lngi).Left, pic2(lngi).Height - 300, (lbl1(lngi).Width - 30) / 2
        lbl8(lngi).Move lbl7(lngi).Left + lbl7(lngi).Width + 30, lbl7(lngi).Top, lbl1(lngi).Width - lbl7(lngi).Width - 30
        lbl3(lngi).Move lbl2(lngi).Left, lbl7(lngi).Top, lbl2(lngi).Width
        ImgCard(lngi).Move 60, 30, 200, 200
        ImgCard(lngi).ZOrder 0
    Next
    If mlngSelTab >= 0 And mlngSelTab < mlngCount Then Call SelectPeopleCard(mlngSelTab)
    Call LockWindowUpdate(0)
    If Err <> 0 Then Err.Clear
End Sub

Public Sub SetCardFocus(strTitle As String, strfind As String)
    '功能：根据提供的标题和内容参数查找相关选项卡并定位
    '参数：strTitle-要定位数据的类型比如病人id、主页id、姓名等，strFind-于strTitle对应的内容123、123、张三等
    Dim ArrTitle
    Dim ArrFind
    Dim lngi As Integer
    Dim lngj As Integer
    Dim lngCount As Long
    Dim strCopyTitle As String
    Dim rsData As ADODB.Recordset
    Dim strFilter As String
    Dim lng所在页 As Long
    
    If strTitle = "" Or strfind = "" Then Exit Sub
    
    On Error GoTo ErrHand
    If Not mRsAll.EOF Then
        mRsAll.MoveFirst
    Else
        Exit Sub
    End If
    
    CopyRecord mRsAll, rsData
    
    If rsData Is Nothing Then Exit Sub '如果mrsbr本身没有数据时，直接跳出查找
    
    lngCount = 0
    ArrTitle = Split(strTitle, "'")
    ArrFind = Split(strfind, "'")
    
    For lngi = 0 To UBound(ArrTitle)
        strFilter = strFilter & ArrTitle(lngi) & "=" & ArrFind(lngi) & " and "
    Next
    strFilter = Left(strFilter, Len(strFilter) - 4)
    
    rsData.Filter = strFilter
    
    If rsData.RecordCount > 0 Then
        lngCount = rsData.Bookmark
        If pic10.Visible = True Then '如果有多页数据，则要将查询到的数据所在位置做一定的处理，还要改变当前页
            lng所在页 = Fix((lngCount - 1) / 50) + IIf(mRsAll.RecordCount Mod 50 = 0, 0, 1)
            lngCount = (lngCount - 1) Mod 50
            If lng所在页 <> Val(TXT10.Text) Then
                SetPage lng所在页
            End If
        Else
            lngCount = lngCount - 1
        End If
    End If
    Call SelectPeopleCard(lngCount)
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function ExecuteCommand(ParamArray varCmd() As Variant) As Boolean
    Dim intLoop As Integer
    Dim rsSAD As New ADODB.Recordset
    Dim lngi As Long
    Dim lngj As Long
    Dim lngHeight As Long
    On Error GoTo Error
    
    Call SQLRecord(rsSAD)
    
    For intLoop = 0 To UBound(varCmd)
        Select Case varCmd(intLoop)
            Case "初始控件":
                '调整选项卡的位置和个数
                If mlngCount = 0 Then ExecuteCommand = False: Exit Function
                Call LockWindowUpdate(UserControl.hWnd)
                For lngi = 0 To mlngCount - 1
                    If lngi = 0 Then
                        pic2(lngi).Visible = True
                    Else
                        Load pic2(lngi)
                        Load pic4(lngi)
                        Load chk1(lngi)
                        Load lbl1(lngi)
                        Load lbl2(lngi)
                        Load lbl3(lngi)
                        Load lbl4(lngi)
                        Load lbl5(lngi)
                        Load lbl6(lngi)
                        Load lbl7(lngi)
                        Load lbl8(lngi)
                        Load ImgCard(lngi)
                        Load shpLeft(lngi): shpLeft(lngi).Visible = False
                        Load shpRight(lngi): shpRight(lngi).Visible = False
                        Load shpTop(lngi): shpTop(lngi).Visible = False
                        Load shpBottom(lngi): shpBottom(lngi).Visible = False
                        
                         '将标签放在容器里
                        Set pic4(lngi).Container = pic2(lngi)
                        Set chk1(lngi).Container = pic4(lngi)
                        Set lbl1(lngi).Container = pic2(lngi)
                        Set lbl2(lngi).Container = pic2(lngi)
                        Set lbl3(lngi).Container = pic2(lngi)
                        Set lbl4(lngi).Container = pic2(lngi)
                        Set lbl5(lngi).Container = pic2(lngi)
                        Set lbl6(lngi).Container = pic2(lngi)
                        Set lbl7(lngi).Container = pic2(lngi)
                        Set lbl8(lngi).Container = pic2(lngi)
                        Set ImgCard(lngi).Container = pic2(lngi)
                        Set ImgCard(lngi).Picture = Nothing
                        Set shpLeft(lngi).Container = pic2(lngi)
                        Set shpRight(lngi).Container = pic2(lngi)
                        Set shpTop(lngi).Container = pic2(lngi)
                        Set shpBottom(lngi).Container = pic2(lngi)
                        
                        pic2(lngi).Visible = True
                        pic4(lngi).Visible = True
                        chk1(lngi).Visible = True
                        lbl1(lngi).Visible = True
                        lbl2(lngi).Visible = True
                        lbl3(lngi).Visible = True
                        lbl4(lngi).Visible = True
                        lbl5(lngi).Visible = True
                        lbl6(lngi).Visible = True
                        lbl7(lngi).Visible = True
                        lbl8(lngi).Visible = True
                        ImgCard(lngi).Visible = True
                    End If
                Next
                Call LockWindowUpdate(0)
                UserControl.Refresh
                For lngi = 0 To mlngCount - 1
                    Call LoadData(lngi, mRsBR, mImgList)
                    UserControl.Refresh
                    mRsBR.MoveNext
                Next
                
                If mRsBR.RecordCount <> 0 Then
                    mRsBR.MoveFirst
                End If
                mbln初始化 = True
                Call pic1_Resize
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "清除控件":
                mbln初始化 = False
                VS1.Visible = False
                For lngi = 0 To pic2.Count - 1
                    If lngi = 0 Then
                        pic2(lngi).Visible = False
                        shpLeft(lngi).Visible = False
                        shpRight(lngi).Visible = False
                        shpTop(lngi).Visible = False
                        shpBottom(lngi).Visible = False
                    Else
                        Unload chk1(lngi)
                        Unload pic4(lngi)
                        Unload lbl1(lngi)
                        Unload lbl2(lngi)
                        Unload lbl3(lngi)
                        Unload lbl4(lngi)
                        Unload lbl5(lngi)
                        Unload lbl6(lngi)
                        Unload lbl7(lngi)
                        Unload lbl8(lngi)
                        Unload ImgCard(lngi)
                        Unload shpLeft(lngi)
                        Unload shpRight(lngi)
                        Unload shpTop(lngi)
                        Unload shpBottom(lngi)
                        Unload pic2(lngi)
                    End If
                Next
                UserControl.Refresh
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End Select
    Next
    ExecuteCommand = True
    Exit Function
Error:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    ExecuteCommand = False
End Function

Private Sub pic2_Click(Index As Integer)
    Call SelectPeopleCard(Index, True)
End Sub

Private Sub LoadData(Index As Long, rsData As ADODB.Recordset, Optional imgList As Object)
    '功能:将数据加载到页面上,
    '参数:rsdata-要加载的数据，imgList-imgList控件，存放图标
    '返回:
    Dim blnHaveImage As Boolean
    Dim lngj As Long
    Dim blnFunc As Boolean
    Dim lngImgList As Long
    Dim str颜色 As String, str背景颜色 As String '存放查询到的背景颜色转化为RGB后的数据
    Dim str规则主 As String '存放marr规则中的主数据
    Dim str规则辅 As String '存放marr规则中的辅助数据
    Dim lng颜色 As Long, lng背景颜色 As Long '和str颜色配套使用
    If Not imgList Is Nothing Then blnHaveImage = True

    If Val(mArr规则(8, 2) & "") = 0 And (mArr规则(8, 0) Like "图标" Or mArr规则(8, 1) Like "图标") _
        And blnHaveImage Then
        If Val(mArr规则(8, 7)) > 0 And Val(mArr规则(8, 7)) <= imgList.ListImages.Count Then
        '图标提取方式一,图片的索引在0到imagelist的最大索引之间
            ImgCard(Index).Picture = imgList.ListImages(Val(mArr规则(8, 7))).Picture
        End If
    End If
    
    For lngj = 0 To UBound(mArr规则, 1)
        str颜色 = getRsValue(mArr规则(lngj, 0) & "颜色", rsData) '记录集中有 主数据+颜色这种形式的数据时，将该数据取出，并添加到相应的控件
        lng颜色 = GetRGBFromOLEColor(Val(str颜色))
        str背景颜色 = getRsValue(mArr规则(lngj, 0) & "背景颜色", rsData)
        lng背景颜色 = GetRGBFromOLEColor(Val(str背景颜色))
        str规则主 = getRsValue(mArr规则(lngj, 0), rsData)
        str规则辅 = getRsValue(mArr规则(lngj, 1), rsData)
        Select Case lngj
            Case 0: 'lbl1
                lbl1(Index).Caption = IIf(Val(mArr规则(lngj, 2)) = 0, str规则主, "")
                lbl1(Index).Tag = str规则辅
                If str颜色 <> "" Then
                    lbl1(Index).ForeColor = lng颜色
                End If
                If str背景颜色 <> "" Then
                    lbl1(Index).BackColor = lng背景颜色
                End If
                
            Case 1: 'lbl2
                lbl2(Index).Caption = IIf(Val(mArr规则(lngj, 2)) = 0, str规则主, "")
                lbl2(Index).Tag = str规则辅
                If str颜色 <> "" Then
                    lbl2(Index).ForeColor = lng颜色
                End If
                If str背景颜色 <> "" Then
                    lbl2(Index).BackColor = lng背景颜色
                End If
                
            Case 2: 'lbl3
                lbl3(Index).Caption = IIf(Val(mArr规则(lngj, 2)) = 0, str规则主, "")
                lbl3(Index).Tag = str规则辅
                If str颜色 <> "" Then
                    lbl3(Index).ForeColor = lng颜色
                End If
                If str背景颜色 <> "" Then
                    lbl3(Index).BackColor = lng背景颜色
                End If
                
            Case 3: 'lbl4
                lbl4(Index).Caption = IIf(Val(mArr规则(lngj, 2)) = 0, str规则主, "")
                lbl4(Index).Tag = str规则辅
                If str颜色 <> "" Then
                    lbl4(Index).ForeColor = lng颜色
                End If
                If str背景颜色 <> "" Then
                    lbl4(Index).BackColor = lng背景颜色
                End If
                
            Case 4: 'lbl5
                lbl5(Index).Caption = IIf(Val(mArr规则(lngj, 2)) = 0, str规则主, "")
                lbl5(Index).Tag = str规则辅
                If str颜色 <> "" Then
                    lbl5(Index).ForeColor = lng颜色
                End If
                If str背景颜色 <> "" Then
                    lbl5(Index).BackColor = lng背景颜色
                End If
                
            Case 5: 'lbl6
                lbl6(Index).Caption = IIf(Val(mArr规则(lngj, 2)) = 0, str规则主, "")
                lbl6(Index).Tag = str规则辅
                If str颜色 <> "" Then
                    lbl6(Index).ForeColor = lng颜色
                End If
                If str背景颜色 <> "" Then
                    lbl6(Index).BackColor = lng背景颜色
                End If
                
            Case 6: 'lbl7
                lbl7(Index).Caption = IIf(Val(mArr规则(lngj, 2)) = 0, str规则主, "")
                lbl7(Index).Tag = str规则辅
                If str颜色 <> "" Then
                    lbl7(Index).ForeColor = lng颜色
                End If
                If str背景颜色 <> "" Then
                    lbl7(Index).BackColor = lng背景颜色
                End If
                
            Case 7: 'lbl8
                lbl8(Index).Caption = IIf(Val(mArr规则(lngj, 2)) = 0, str规则主, "")
                lbl8(Index).Tag = str规则辅
                If str颜色 <> "" Then
                    lbl8(Index).ForeColor = lng颜色
                End If
                If str背景颜色 <> "" Then
                    lbl8(Index).BackColor = lng背景颜色
                End If
                
            Case 8: 'imgCard，方式二，图片的索引在0到imagelist的最大索引之间
                'imgCard的处理方式和label的处理方式不同
                If mArr规则(lngj, 0) = "图标" And blnHaveImage Then
                    lngImgList = Val(str规则主)
                    If lngImgList > 0 And lngImgList <= imgList.ListImages.Count Then
                        ImgCard(Index).Picture = imgList.ListImages(lngImgList).Picture
                    End If
                End If
        End Select
    Next
End Sub

Private Function getRsValue(name, rs As ADODB.Recordset) As String
    '获取对应相应列的数据
    Dim str As String
    On Error Resume Next
    str = rs.Fields(name).Value & ""
    If Err.Description <> "" Then
        str = ""
        Err.Description = ""
    End If
    getRsValue = str
End Function

Private Sub SelectPeopleCard(ByVal Index As Integer, Optional ByVal blnClick As Boolean = False)
'功能：更换选项卡后变更状态，及UCE控件的显示
'参数：Index--卡片索引。blnClick：是否是点击选择卡片(用于判断多次点击同一卡片刷新的处理)
    Dim lngi As Long
    Dim strRetrun As String
    
    mlngSelTab = Index
    If pic2.Count > 0 Then
        For lngi = 0 To pic2.Count - 1 '为了显示边线
            shpLeft(lngi).Visible = IIf(lngi = Index, True, False)
            shpRight(lngi).Visible = IIf(lngi = Index, True, False)
            shpTop(lngi).Visible = IIf(lngi = Index, True, False)
            shpBottom(lngi).Visible = IIf(lngi = Index, True, False)
            If lngi = Index Then
               shpLeft(lngi).ZOrder 0
               shpRight(lngi).ZOrder 0
               shpTop(lngi).ZOrder 0
               shpBottom(lngi).ZOrder 0
               If pic2(lngi).Visible And pic2(lngi).Enabled Then pic2(lngi).SetFocus
            End If
        Next
    End If
    
     '当有滚动条时，且选中的数据在pic1未显示部分，那么要移动滚动条已方便用户查看
    If VS1.Visible = True And pic2(Index).Top + pic2(Index).Height > pic1.ScaleHeight Then '当控件位于显示界面以下时
        VS1.Value = (Abs(pic2(Index).Top - pic2(0).Top - pic1.ScaleHeight + pic2(Index).Height + 50)) / mdblVS系数
    ElseIf VS1.Visible = True And pic2(Index).Top < 0 Then  '当控件位于显示界面以上时
        VS1.Value = (Abs(pic2(Index).Top - pic2(0).Top + 50)) / mdblVS系数 + 1
    End If
    '返回一个病人的数据，不会返回image中的数据
    strRetrun = lbl1(Index).Caption & "'" & lbl1(Index).Tag & "'" & lbl2(Index).Caption & "'" & lbl2(Index).Tag & "'" & lbl3(Index).Caption & "'" & lbl3(Index).Tag & "'" & lbl4(Index).Caption & "'" & lbl4(Index).Tag & "'" & _
                 lbl5(Index).Caption & "'" & lbl5(Index).Tag & "'" & lbl6(Index).Caption & "'" & lbl6(Index).Tag & "'" & lbl7(Index).Caption & "'" & lbl7(Index).Tag & "'" & lbl8(Index).Caption & "'" & lbl8(Index).Tag

    If mlngLocalIDNum >= 0 Then '获取该选项卡的ID
        mstrLocalID = GetValue(mlngLocalIDNum, Index)
    End If
    If mstrReturn = strRetrun And blnClick = True Then Exit Sub
    mstrReturn = strRetrun
    RaiseEvent CardChanged
End Sub

Private Sub UserControl_Resize()
    '控件改变大小pic1也会改变大小
    On Error GoTo Errorhand
    pic3.Left = UserControl.ScaleLeft
    pic3.Top = UserControl.ScaleTop
    pic3.Width = UserControl.ScaleWidth
    pic3.Height = 575
    
    '调整pi1的位置
    pi1.Left = pic3.ScaleLeft + 100
    pi1.Top = pic3.ScaleTop + 75
    pi1.Width = pic3.ScaleWidth - 250 - chkFilter.Width - 50
    '调整过滤按钮位置
    chkFilter.Move pi1.Left + pi1.Width + 50, pi1.Top
    
    If mlng页数 = 1 Or mlng页数 = 0 Then
        pic1.Left = UserControl.ScaleLeft
        pic1.Top = UserControl.ScaleTop + pic3.Height
        pic1.Width = UserControl.ScaleWidth
        If UserControl.ScaleHeight > pic3.Height Then
            pic1.Height = UserControl.ScaleHeight - pic3.Height
        End If
        pic10.Visible = False
'        pic2(0).Enabled = True
    Else
        pic1.Left = UserControl.ScaleLeft
        pic1.Top = UserControl.ScaleTop + pic3.Height
        pic1.Width = UserControl.ScaleWidth
        If UserControl.ScaleHeight > pic3.Height + pic10.Height Then
            pic1.Height = UserControl.ScaleHeight - pic3.Height - pic10.Height
        End If
        
        pic10.Move UserControl.ScaleLeft, pic1.Top + pic1.Height, pic1.Width
        pic10.Visible = True
    End If
Errorhand:
End Sub

Private Sub VS1_Change()
    VS1_Scroll
End Sub

Private Sub VS1_Scroll()
    '滚动滑轮，picturebox移动
    Dim lngi As Long
    If pic2.Count > 0 Then
        Call LockWindowUpdate(UserControl.hWnd)
        For lngi = 0 To pic2.Count - 1
            pic2(lngi).Top = Val(pic2(lngi).Tag) - VS1.Value * mdblVS系数
        Next
        Call LockWindowUpdate(0)
    End If
End Sub

Private Function GetRGBFromOLEColor(ByVal dwOleColour As Long) As Long
    '将VB的颜色转换为RGB表示
    Dim clrref As Long
    Dim r As Long, g As Long, b As Long
    
    OleTranslateColor dwOleColour, 0, clrref
    
    b = (clrref \ 65536) And &HFF
    g = (clrref \ 256) And &HFF
    r = clrref And &HFF
    
    GetRGBFromOLEColor = RGB(r, g, b)
End Function


Private Sub RsTitelCopy(ByVal RsProm As ADODB.Recordset, ToRs)
    '功能：新建ToRs记录集，将RsProm的结构复制到ToRs上
    '参数：RsProm-原记录集，ToRs-新建的记录集，因为有程序需要传入动态创建的记录集，这些记录集放在数组中，所以tors不限制一定是记录集类型
    Dim lngi As Long
    Set ToRs = New ADODB.Recordset
    With ToRs '初始化rsReturn
        For lngi = 0 To RsProm.Fields.Count - 1
            .Fields.Append RsProm.Fields(lngi).name, adLongVarChar, 100, adFldIsNullable
        Next
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub CopyRecord(ByVal RsProm As ADODB.Recordset, ToRs As ADODB.Recordset)
    '功能：将记录集RsProm的结构还有数据都复制给ToRs
    '参数：RsProm-要赋值的记录集，ToRs-目标记录集
    Dim lngi As Long
    Dim lngj As Long
    Call RsTitelCopy(RsProm, ToRs)
    With ToRs
        If RsProm.RecordCount > 0 Then '以前没有对rsbr的数据做判断会报错
            For lngi = 0 To RsProm.RecordCount - 1
                .AddNew
                For lngj = 0 To RsProm.Fields.Count - 1
                    .Fields(lngj).Value = RsProm.Fields(lngj).Value
                Next
                .Update
                RsProm.MoveNext
            Next
            RsProm.MoveFirst
            If .RecordCount > 0 Then
                .MoveFirst
            End If
        End If
    End With
End Sub

