VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmDateSel 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2550
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4020
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4020
   StartUpPosition =   3  '窗口缺省
   Begin MSComCtl2.MonthView mtvSel 
      Height          =   2160
      Left            =   -15
      TabIndex        =   0
      Top             =   375
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   3810
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   0
      StartOfWeek     =   114491393
      TitleBackColor  =   -2147483635
      TitleForeColor  =   -2147483634
      CurrentDate     =   39759
   End
   Begin XtremeSuiteControls.ShortcutCaption stcTittle 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3990
      _Version        =   589884
      _ExtentX        =   7038
      _ExtentY        =   661
      _StockProps     =   6
      Caption         =   "日期选择器"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmDateSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private mstrCurDate As String
Private mstrMinDate As String
Private mstrMaxDate As String
Private msngX As Single, msngY As Single, mlngTxtH As Long

Public Function SelectDate(ByVal frmMain As Form, ByVal sngX As Single, ByVal sngY As Single, lngTxtH As Long, ByRef strDate As String, _
    Optional strMinDate As String = "", Optional strMaxDate As String = "") As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:日期选择器控件
    '入参:strDate-默认指向日期  :'yyyy-mm-dd的形式
    '     strMinDate-默认的最小日期
    '     strMaxDate-默认的最大日期
    '出参:
    '返回:选择,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-11-07 11:23:35
    '-----------------------------------------------------------------------------------------------------------
    mstrCurDate = strDate: mstrMinDate = strMinDate: mstrMaxDate = strMaxDate
    msngX = sngX: msngY = sngY: mlngTxtH = lngTxtH
    
    mblnOk = False
    Me.Show 1, frmMain
    If mblnOk = False Then mstrCurDate = ""
    strDate = mstrCurDate
    SelectDate = mblnOk
End Function
Private Sub InitData()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-11-07 11:30:47
    '-----------------------------------------------------------------------------------------------------------
    Dim strDabaseDate As String
    On Error GoTo ErrHandle
    strDabaseDate = Format(Sys.Currentdate, "yyyy-mm-dd")
    
    If mstrCurDate = "" Then mstrCurDate = strDabaseDate
    If mstrMaxDate = "" Then mstrMaxDate = strDabaseDate
    If mstrMinDate = "" Then mstrMinDate = "1901-01-01"
    mtvSel.MinDate = CDate(mstrMinDate)
    If CDate(mstrMaxDate) < CDate(mstrMinDate) Then
        mstrMaxDate = "9999-12-31"
    End If
    mtvSel.MaxDate = CDate(mstrMaxDate)
    If CDate(mstrCurDate) < CDate(mstrMinDate) Then
        mstrCurDate = mstrMinDate
    End If
    mtvSel.Value = CDate(mstrCurDate)
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub Form_Activate()
    Call zlControl.ControlSetFocus(mtvSel)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call OnSelect
        Exit Sub
    End If
    If KeyCode = vbKeyEscape Then '
        mblnOk = False: mstrCurDate = ""
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Call InitData
    With Me
        If msngX + .Width > Screen.Width Then
            .Left = Screen.Width - .Width
        Else
            .Left = msngX
        End If
        If msngY + .Height > Screen.Height Then
           .Top = msngY - mlngTxtH - .Height
        Else
            .Top = msngY
        End If
    End With
    
End Sub
Private Sub OnSelect()
    '-----------------------------------------------------------------------------------------------------------
    '功能:确认被选择的日期
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-11-07 11:36:07
    '-----------------------------------------------------------------------------------------------------------
    mstrCurDate = Format(mtvSel.Value, "yyyy-mm-dd")
    mblnOk = True
    Unload Me
End Sub

Private Sub mtvSel_DateDblClick(ByVal DateDblClicked As Date)
    Call OnSelect
End Sub

Private Sub mtvSel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call OnSelect
End Sub
