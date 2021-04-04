VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmDownDate 
   BorderStyle     =   0  'None
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2985
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdCancel 
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2655
      TabIndex        =   2
      Top             =   30
      Width           =   315
   End
   Begin MSComCtl2.MonthView mthDate 
      Height          =   2220
      Left            =   -30
      TabIndex        =   0
      Top             =   345
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483634
      BackColor       =   -2147483632
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   -2147483635
      StartOfWeek     =   132907009
      TitleBackColor  =   14737632
      TrailingForeColor=   65535
      CurrentDate     =   40357
   End
   Begin XtremeSuiteControls.ShortcutCaption stcCaption 
      Height          =   360
      Left            =   15
      TabIndex        =   1
      Top             =   15
      Width           =   2955
      _Version        =   589884
      _ExtentX        =   5212
      _ExtentY        =   635
      _StockProps     =   6
      Caption         =   "日期选择"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmDownDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private Type tyLocalWin
        Left  As Single
        Top As Single
        Width As Single
        Height As Single
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private mdtValue As Date
Private mvRect As RECT
Private msngHeight As Single
Public Event DateClick(ByVal DateClicked As Date)

Private Function GetObjectRect(ByVal lngHwnd As Long) As RECT
'功能：获取指定控件在屏幕中的位置(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetObjectRect = vRect
End Function

Public Property Get MaxDate() As Date
    MaxDate = mthDate.MaxDate
End Property
Public Property Let MaxDate(ByVal vNewValue As Date)
        mthDate.MaxDate = vNewValue
End Property
Public Property Get MinDate() As Date
    MinDate = mthDate.MinDate
End Property
Public Property Let MinDate(ByVal vNewValue As Date)
    mthDate.MinDate = vNewValue
End Property
Public Property Get Value() As Date
       Value = mthDate.Value
End Property
Public Property Let Value(ByVal vNewValue As Date)
    mthDate.Value = vNewValue
End Property
Public Function ShowDate(ByVal objCtl As Object, ByVal dtMaxDate As Date, dtMinDate As Date, dtValue As Date) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：下拉选择数据
    '编制：刘兴洪
    '日期：2010-06-28 15:44:42
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    mblnOk = False
    mvRect = GetObjectRect(objCtl.hWnd)
    msngHeight = objCtl.Height
    With mthDate
        .MaxDate = dtMaxDate: .MinDate = dtMinDate: .Value = dtValue
    End With
    mdtValue = dtValue
    Me.Show 1
    ShowDate = mblnOk
    dtValue = mdtValue
End Function
Private Sub cmdCancel_Click()
    mblnOk = False: Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmdCancel_Click
    End If
End Sub

Private Sub Form_Load()
        With Me
            .Left = mvRect.Left
            .Top = mvRect.Top + msngHeight
        End With
End Sub

Private Sub mthDate_DateClick(ByVal DateClicked As Date)
        RaiseEvent DateClick(DateClicked)
End Sub

Private Sub mthDate_DblClick()
        mdtValue = mthDate.Value
        mblnOk = True
        Unload Me
End Sub
