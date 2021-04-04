VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDateTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "日期时间"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   Icon            =   "frmDateTime.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5475
      TabIndex        =   13
      Top             =   4755
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4215
      TabIndex        =   12
      Top             =   4755
      Width           =   1100
   End
   Begin VB.CheckBox chkTime 
      Caption         =   "时间(&T)"
      Height          =   195
      Left            =   4335
      TabIndex        =   5
      Top             =   120
      Value           =   1  'Checked
      Width           =   960
   End
   Begin VB.Frame fraTime 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   4320
      TabIndex        =   6
      Top             =   270
      Width           =   2235
      Begin VB.ListBox lstTime 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   0
         TabIndex        =   10
         Top             =   2505
         Width           =   2235
      End
      Begin VB.PictureBox picTime 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1770
         Left            =   0
         ScaleHeight     =   1740
         ScaleWidth      =   2205
         TabIndex        =   7
         Top             =   90
         Width           =   2235
         Begin VB.Line linHand 
            X1              =   240
            X2              =   960
            Y1              =   735
            Y2              =   585
         End
         Begin VB.Shape shpCenter 
            Height          =   90
            Left            =   675
            Shape           =   3  'Circle
            Top             =   720
            Width           =   90
         End
         Begin VB.Shape shpDot 
            FillColor       =   &H00FFFFFF&
            Height          =   105
            Index           =   0
            Left            =   690
            Shape           =   3  'Circle
            Top             =   180
            Width           =   135
         End
      End
      Begin MSComCtl2.DTPicker dtTime 
         Height          =   300
         Left            =   360
         TabIndex        =   8
         Top             =   1920
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   16515075
         UpDown          =   -1  'True
         CurrentDate     =   38549.5423726852
      End
      Begin VB.Label lblTimeType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "格式类型"
         Height          =   180
         Left            =   0
         TabIndex        =   11
         Top             =   2310
         Width           =   720
      End
      Begin VB.Label lblAmOrPm 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上午"
         Height          =   180
         Left            =   1620
         TabIndex        =   9
         Top             =   1980
         Width           =   360
      End
   End
   Begin VB.CheckBox chkDate 
      Caption         =   "日期(&D)"
      Height          =   195
      Left            =   165
      TabIndex        =   1
      Top             =   120
      Value           =   1  'Checked
      Width           =   930
   End
   Begin VB.Frame fraDate 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   165
      TabIndex        =   0
      Top             =   270
      Width           =   4035
      Begin VB.ListBox lstDate 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   0
         TabIndex        =   3
         Top             =   2505
         Width           =   4035
      End
      Begin MSComCtl2.MonthView mvwDate 
         Height          =   2160
         Left            =   0
         TabIndex        =   2
         Top             =   90
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   3810
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         StartOfWeek     =   16515073
         CurrentDate     =   38549
         MaxDate         =   401769
         MinDate         =   367
      End
      Begin VB.Label lblDateType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "格式类型"
         Height          =   180
         Left            =   0
         TabIndex        =   4
         Top             =   2325
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmDateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const conPI As Double = 3.14159265358979
Dim blnHandMove As Boolean

Dim blnOK As Boolean

Private Sub chkDate_Click()
    Me.fraDate.Enabled = IIf(Me.chkDate.Value = vbChecked, True, False)
    If Me.chkDate.Value = vbChecked Or Me.chkTime.Value = vbChecked Then
        Me.cmdOK.Enabled = True
    Else
        Me.cmdOK.Enabled = False
    End If
End Sub

Private Sub chkTime_Click()
    Me.fraTime.Enabled = IIf(Me.chkTime.Value = vbChecked, True, False)
    If Me.chkDate.Value = vbChecked Or Me.chkTime.Value = vbChecked Then
        Me.cmdOK.Enabled = True
    Else
        Me.cmdOK.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    blnOK = False: Me.Hide
End Sub

Private Sub cmdOK_Click()
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Caption, "TimeCheck", chkTime.Value
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Caption, "DateCheck", chkDate.Value
    If lstTime.ListIndex >= 0 Then
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Caption, "TimeType", lstTime.ListIndex
    End If
    If lstDate.ListIndex >= 0 Then
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Caption, "DateType", lstDate.ListIndex
    End If
    
    blnOK = True: Me.Hide
End Sub

Private Sub dtTime_Change()
    Call SetTimer(Hour(Me.dtTime.Value) + Minute(Me.dtTime.Value) / 60 + Second(Me.dtTime.Value) / 60 / 60)
End Sub

Private Sub lblAmOrPm_Click()
    If Me.lblAmOrPm.Caption = "下午" Then
        Me.lblAmOrPm.Caption = "上午"
        Me.dtTime.Value = Me.dtTime.Value - 12 / 24
    Else
        Me.lblAmOrPm.Caption = "下午"
        Me.dtTime.Value = Me.dtTime.Value + 12 / 24
    End If
    Call SetTimer(Hour(Me.dtTime.Value) + Minute(Me.dtTime.Value) / 60 + Second(Me.dtTime.Value) / 60 / 60)
End Sub

Private Sub mvwDate_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)
    Dim intYear As Integer, intMonth As Integer, intDay As Integer
    Dim strYear As String, strMonth As String, strDay As String

    intYear = Year(Me.mvwDate.Value)
    intMonth = Month(Me.mvwDate.Value)
    intDay = Day(Me.mvwDate.Value)
    
    strYear = GetChineseNumber(Mid(intYear, 1, 1), True) & GetChineseNumber(Mid(intYear, 2, 1), True) & GetChineseNumber(Mid(intYear, 3, 1), True) & GetChineseNumber(Mid(intYear, 4, 1), True)
    strMonth = GetChineseNumber(intMonth)
    strDay = GetChineseNumber(intDay)
    
    With Me.lstDate
        .Clear
        .AddItem strYear & "年" & strMonth & "月" & strDay & "日"
        .AddItem strYear & "年" & strMonth & "月"
        .AddItem strMonth & "月" & strDay & "日"
        .AddItem intYear & "年" & intMonth & "月" & intDay & "日"
        .AddItem intYear & "年" & intMonth & "月"
        .AddItem intMonth & "月" & intDay & "日"
        .AddItem intYear & "-" & intMonth & "-" & intDay
        .AddItem intMonth & "-" & intDay
        .AddItem WeekdayName(Weekday(Me.mvwDate.Value))
        .AddItem GetSolarTerm(Me.mvwDate.Value)
        If Val(.Tag) >= 0 Then
            .ListIndex = Val(.Tag)
        End If
        If .ListIndex = -1 Then .ListIndex = 0
        .TopIndex = .ListIndex
    End With
End Sub

Private Sub picTime_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    blnHandMove = True
End Sub

Private Sub picTime_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim xLine As Long, yLine As Long
    Dim dblSin As Double, dblValue As Double
    
    If blnHandMove = False Then Exit Sub
    xLine = x - (Me.shpCenter.Left + Me.shpCenter.Width / 2)
    yLine = y - (Me.shpCenter.Top + Me.shpCenter.Height / 2)
    
    If xLine = 0 And yLine = 0 Then Exit Sub
    
    dblSin = yLine / Sqr(xLine ^ 2 + yLine ^ 2)
    If dblSin = 1 Then
        dblValue = 6
    ElseIf dblSin = -1 Then
        dblValue = 0
    Else
        If Sgn(xLine) >= 0 Then
            dblValue = Round(Atn(dblSin / Sqr(-dblSin * dblSin + 1)) / conPI * 180 / 30, 1) + 3
        Else
            dblValue = 9 - Round(Atn(dblSin / Sqr(-dblSin * dblSin + 1)) / conPI * 180 / 30, 1)
        End If
    End If
    If Me.lblAmOrPm.Caption = "下午" And dblValue < 12 Then dblValue = dblValue + 12
    Call SetTimer(dblValue)
End Sub

Private Sub picTime_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call picTime_MouseMove(Button, Shift, x, y)
    blnHandMove = False
    Call dtTime_Change
End Sub

Private Function GetSolarTerm(ByVal dtAsk As Date) As String
    '功能：获得指定日期的节气
    '参数：dtAsk，公历日期
    
    Const conYearMinutes As Double = 525948.76   '每年的分钟数，一年实际是365.242194444天，按分钟计算基本能准确
    Dim dtBaseDate As Date
    Dim aryTermName() As String
    Dim aryTermData() As String
    
    Dim dblMinutes As Double
    Dim dtTermDate As Date
    
    dtAsk = Int(dtAsk) + 2 / 24 + 5 / 24 / 60
    dtBaseDate = Format("1900-01-06 2:05:00", "YYYY-MM-DD hh:mm:ss")
    If dtAsk < dtBaseDate Then GetSolarTerm = "": Exit Function
    aryTermName = Split("小寒,大寒,立春,雨水,惊蛰,春分,清明,谷雨,立夏,小满,芒种,夏至,小暑,大暑,立秋,处暑,白露,秋分,寒露,霜降,立冬,小雪,大雪,冬至", ",")
    aryTermData = Split("0,21208,42467,63836,85337,107014,128867,150921,173149,195551,218072,240693,263343,285989,308563,331033,353350,375494,397447,419210,440795,462224,483532,504758", ",")
      
    Dim intCount As Integer
    For intCount = 0 To UBound(aryTermData)
        dblMinutes = conYearMinutes * (Year(dtAsk) - 1900) + CLng(aryTermData(intCount))
        dtTermDate = DateAdd("n", dblMinutes, dtBaseDate)
        
        If DateDiff("d", dtAsk, dtTermDate) >= 0 Then
            Select Case DateDiff("d", dtAsk, dtTermDate)
            Case 0
                GetSolarTerm = aryTermName(intCount)
            Case Is < 8
                GetSolarTerm = aryTermName(intCount) & "前" & DateDiff("d", dtAsk, dtTermDate) & "天"
            Case Else
                If intCount <> 0 Then
                    dblMinutes = conYearMinutes * (Year(dtAsk) - 1900) + CLng(aryTermData(intCount - 1))
                    dtTermDate = DateAdd("n", dblMinutes, dtBaseDate)
                    GetSolarTerm = aryTermName(intCount - 1) & "后" & Abs(DateDiff("d", dtAsk, dtTermDate)) & "天"
                Else
                    dblMinutes = conYearMinutes * (Year(dtAsk) - 1 - 1900) + CLng(aryTermData(UBound(aryTermData)))
                    dtTermDate = DateAdd("n", dblMinutes, dtBaseDate)
                    GetSolarTerm = aryTermName(UBound(aryTermData)) & "后" & Abs(DateDiff("d", dtAsk, dtTermDate)) & "天"
                End If
            End Select
            Exit Function
        ElseIf intCount = UBound(aryTermData) Then
            GetSolarTerm = aryTermName(intCount) & "后" & Abs(DateDiff("d", dtAsk, dtTermDate)) & "天"
            Exit Function
        End If
    Next
    GetSolarTerm = ""
End Function

Private Sub DrawWatch()
    Dim CenterX As Long, CenterY As Long, lngRadii As Long
    CenterX = Me.picTime.ScaleWidth / 2
    CenterY = Me.picTime.ScaleHeight / 2
    If CenterX < CenterY Then
        lngRadii = CenterX - 60
    Else
        lngRadii = CenterY - 60
    End If
    
    Dim intHour As Integer, x As Long, y As Long
    x = CenterX - Me.shpCenter.Width / 2
    y = CenterY - Me.shpCenter.Height / 2
    Me.shpCenter.Move x, y
    
    Me.linHand.X1 = x + Me.shpCenter.Width / 2
    Me.linHand.Y1 = y + Me.shpCenter.Height / 2
    
    For intHour = 0 To 11
        If intHour > Me.shpDot.Count - 1 Then
            Load Me.shpDot(intHour)
        End If
        If intHour Mod 3 = 0 Then
            Me.shpDot(intHour).Width = 60
            Me.shpDot(intHour).Height = 60
        Else
            Me.shpDot(intHour).Width = 45
            Me.shpDot(intHour).Height = 45
        End If
        x = CenterX + lngRadii * Sin(intHour * 30 / 180 * conPI) - Me.shpDot(intHour).Width / 2
        y = CenterY - lngRadii * Cos(intHour * 30 / 180 * conPI) - Me.shpDot(intHour).Height / 2
        Me.shpDot(intHour).Move x, y
        Me.shpDot(intHour).Visible = True
    Next
End Sub

Private Sub SetTimer(ByVal dblTime As Double)
    
    Dim CenterX As Long, CenterY As Long, lngRadii As Long
    Dim intCount As Integer
    
    CenterX = Me.picTime.ScaleWidth / 2
    CenterY = Me.picTime.ScaleHeight / 2
    If CenterX < CenterY Then
        lngRadii = CenterX - 60
    Else
        lngRadii = CenterY - 60
    End If
    
    If dblTime < 12 Then
        Me.lblAmOrPm.Caption = "上午"
    Else
        Me.lblAmOrPm.Caption = "下午"
    End If
    
    If blnHandMove = True Then
        Me.dtTime.Value = Int(Me.dtTime.Value) + dblTime / 24
    End If
    Me.linHand.X2 = CenterX + lngRadii * Sin(dblTime * 30 / 180 * conPI)
    Me.linHand.Y2 = CenterY - lngRadii * Cos(dblTime * 30 / 180 * conPI)
    
    For intCount = 0 To 11
        If intCount = IIf(dblTime < 12, dblTime, dblTime - 12) Then
            Me.shpDot(intCount).BorderColor = RGB(255, 0, 0)
            Beep
        Else
            Me.shpDot(intCount).BorderColor = RGB(0, 0, 0)
        End If
    Next
    
    If blnHandMove = True Then Exit Sub

    '设置格式
    Dim intHour As Integer, intMinute As Integer, intSecond As Integer
    
    intHour = Hour(Me.dtTime.Value)
    intMinute = Minute(Me.dtTime.Value)
    intSecond = Second(Me.dtTime.Value)
    
    With Me.lstTime
        .Clear
        .AddItem IIf(intHour < 12, "上午", "下午") & GetChineseNumber(IIf(intHour < 12, intHour, intHour - 12)) & "时" & GetChineseNumber(intMinute) & "分"
        .AddItem GetChineseNumber(intHour) & "时" & GetChineseNumber(intMinute) & "分"
        .AddItem IIf(intHour < 12, "上午", "下午") & IIf(intHour < 12, intHour, intHour - 12) & "时" & intMinute & "分" & intSecond & "秒"
        .AddItem IIf(intHour < 12, "上午", "下午") & IIf(intHour < 12, intHour, intHour - 12) & "时" & intMinute & "分"
        .AddItem intHour & "时" & intMinute & "分" & intSecond & "秒"
        .AddItem intHour & "时" & intMinute & "分"
        .AddItem IIf(intHour < 12, intHour, intHour - 12) & ":" & Format(intMinute, "00") & ":" & Format(intSecond, "00") & IIf(intHour < 12, " AM", " PM")
        .AddItem IIf(intHour < 12, intHour, intHour - 12) & ":" & Format(intMinute, "00") & IIf(intHour < 12, " AM", " PM")
        .AddItem intHour & ":" & Format(intMinute, "00") & ":" & Format(intSecond, "00")
        .AddItem intHour & ":" & Format(intMinute, "00")
        If Val(.Tag) >= 0 Then
            .ListIndex = Val(.Tag)
        End If
        If .ListIndex = -1 Then .ListIndex = 0
        .TopIndex = .ListIndex
    End With
End Sub

Private Function GetChineseNumber(ByVal bytNumber As Byte, Optional blnZeroCircle As Boolean) As String
    '功能：返回汉字数字
    '参数：
    '   bytNumber,要处理的数字，本函数要求不大于99;
    '   blnZeroCircle,是否以○代表0，否则表现为零
    
    Dim bytBit1 As Byte, bytBit2 As Byte
    Dim strBit1 As String, strBit2 As String
    
    If bytNumber > 99 Then GetChineseNumber = "": Exit Function
    
    bytBit1 = bytNumber \ 10: bytBit2 = bytNumber Mod 10
    
    If bytBit1 = 0 Then
        strBit1 = ""
        If blnZeroCircle = False Then
            strBit2 = Split("零,一,二,三,四,五,六,七,八,九", ",")(bytBit2)
        Else
            strBit2 = Split("○,一,二,三,四,五,六,七,八,九", ",")(bytBit2)
        End If
    Else
        strBit1 = Split(",,二,三,四,五,六,七,八,九", ",")(bytBit1) & "十"
        strBit2 = Split(",一,二,三,四,五,六,七,八,九", ",")(bytBit2)
    End If
    GetChineseNumber = strBit1 & strBit2
End Function

Public Function ShowMe(Optional MinDate As Date, Optional MaxDate As Date) As String
    '功能：显示本对话框
    '参数：
    '   MinDate,允许的最小日期
    '   MaxDate,允许的最大日期
    '返回：设置的日期时间串
    
    lstDate.Tag = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Caption, "DateType", 0)
    lstTime.Tag = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Caption, "TimeType", 0)
    chkTime.Value = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Caption, "TimeCheck", 0)
    chkDate.Value = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Caption, "DateCheck", 0)
    Dim dtInit As Date
    
    If MinDate = 0 Then
        Me.mvwDate.MinDate = Format("1901-01-01", "YYYY-MM-DD")
    Else
        Me.mvwDate.MinDate = MinDate
    End If
    If MaxDate = 0 Then
        Me.mvwDate.MaxDate = Format("3000-01-01", "YYYY-MM-DD")
    Else
        Me.mvwDate.MaxDate = MaxDate
    End If
    
    dtInit = Now()
    If dtInit < Me.mvwDate.MinDate Then dtInit = Me.mvwDate.MinDate
    If dtInit > Me.mvwDate.MaxDate Then dtInit = Me.mvwDate.MaxDate
    Me.mvwDate.Value = dtInit
    Call mvwDate_SelChange(Me.mvwDate.MinDate, Me.mvwDate.MaxDate, False)
    
    Call DrawWatch
    Me.dtTime.Value = Now()
    Call dtTime_Change
    
    blnOK = False
    Me.Show vbModal
    If blnOK = False Then Exit Function
        
    ShowMe = Trim(IIf(Me.chkDate.Value = vbChecked, Me.lstDate.Text, "") & " " & IIf(Me.chkTime.Value = vbChecked, Me.lstTime.Text, ""))
    Unload Me
End Function
