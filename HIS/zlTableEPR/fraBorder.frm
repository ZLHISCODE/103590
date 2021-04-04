VERSION 5.00
Begin VB.Form frmBorder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "边框样式"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   Icon            =   "fraBorder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3750
      TabIndex        =   18
      Top             =   4575
      Width           =   1120
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2445
      TabIndex        =   17
      Top             =   4575
      Width           =   1120
   End
   Begin VB.Frame fraCorlor 
      Caption         =   "线色"
      Height          =   2415
      Left            =   2565
      TabIndex        =   6
      Top             =   135
      Width           =   2370
      Begin zlTableEPR.ColorPicker ColorPicker1 
         Height          =   2190
         Left            =   75
         TabIndex        =   7
         Top             =   180
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   3863
      End
   End
   Begin VB.Frame fraShow 
      Caption         =   "边框样式"
      Height          =   1830
      Left            =   135
      TabIndex        =   5
      Top             =   2655
      Width           =   4785
      Begin VB.PictureBox picStyle 
         BackColor       =   &H00FFFFFF&
         Height          =   1380
         Left            =   270
         ScaleHeight     =   1320
         ScaleWidth      =   4185
         TabIndex        =   8
         Top             =   330
         Width           =   4245
         Begin VB.PictureBox picLine 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   960
            Index           =   3
            Left            =   3840
            ScaleHeight     =   960
            ScaleWidth      =   345
            TabIndex        =   16
            Top             =   180
            Width           =   345
            Begin VB.Line lineShow 
               Index           =   3
               X1              =   180
               X2              =   180
               Y1              =   -15
               Y2              =   930
            End
         End
         Begin VB.PictureBox picLine 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   960
            Index           =   2
            Left            =   30
            ScaleHeight     =   960
            ScaleWidth      =   345
            TabIndex        =   15
            Top             =   180
            Width           =   345
            Begin VB.Line lineShow 
               Index           =   2
               X1              =   195
               X2              =   195
               Y1              =   -30
               Y2              =   915
            End
         End
         Begin VB.PictureBox picLine 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   0
            Left            =   255
            ScaleHeight     =   345
            ScaleWidth      =   3765
            TabIndex        =   9
            Top             =   90
            Width           =   3765
            Begin VB.Line lineShow 
               Index           =   0
               X1              =   -30
               X2              =   3735
               Y1              =   60
               Y2              =   60
            End
         End
         Begin VB.PictureBox picLine 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   1
            Left            =   255
            ScaleHeight     =   345
            ScaleWidth      =   3765
            TabIndex        =   14
            Top             =   990
            Width           =   3765
            Begin VB.Line lineShow 
               Index           =   1
               X1              =   -60
               X2              =   3705
               Y1              =   165
               Y2              =   165
            End
         End
         Begin VB.Label lblborder 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "┏"
            Height          =   195
            Index           =   3
            Left            =   3960
            TabIndex        =   13
            Top             =   1080
            Width           =   165
         End
         Begin VB.Label lblborder 
            BackColor       =   &H00FFFFFF&
            Caption         =   "┓"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Top             =   1080
            Width           =   150
         End
         Begin VB.Label lblborder 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "┗"
            Height          =   120
            Index           =   1
            Left            =   3990
            TabIndex        =   11
            Top             =   60
            Width           =   135
         End
         Begin VB.Label lblborder 
            BackColor       =   &H00FFFFFF&
            Caption         =   "┛┓┏┗"
            Height          =   120
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   60
            Width           =   150
         End
      End
   End
   Begin VB.Frame frmLineType 
      Caption         =   "线型"
      Height          =   2415
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   2370
      Begin VB.OptionButton optType 
         Caption         =   "虚线边框"
         Height          =   180
         Index           =   4
         Left            =   180
         TabIndex        =   4
         Top             =   1470
         Width           =   1020
      End
      Begin VB.OptionButton optType 
         Caption         =   "有边框"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   975
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optType 
         Caption         =   "无边框"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   495
         Width           =   855
      End
      Begin VB.OptionButton optType 
         Caption         =   "粗线边框"
         Height          =   180
         Index           =   5
         Left            =   180
         TabIndex        =   1
         Top             =   1950
         Width           =   1020
      End
      Begin VB.Line lineType 
         BorderWidth     =   2
         Index           =   6
         X1              =   1380
         X2              =   1980
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line lineType 
         BorderStyle     =   3  'Dot
         Index           =   4
         X1              =   1380
         X2              =   1980
         Y1              =   1545
         Y2              =   1545
      End
      Begin VB.Line lineType 
         Index           =   1
         X1              =   1380
         X2              =   1980
         Y1              =   1065
         Y2              =   1065
      End
      Begin VB.Line lineType 
         BorderColor     =   &H80000003&
         Index           =   0
         X1              =   1380
         X2              =   1980
         Y1              =   585
         Y2              =   585
      End
   End
End
Attribute VB_Name = "frmBorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlOutBorder As Integer, mlLeftBorder As Integer, mlRightBorder As Integer, mlTopBorder As Integer, mlBottomBorder As Integer, mlShade As Integer
Private mlOutColor As Long, mlLeftColor As Long, mlRightColor As Long, mlTopColor As Long, mlBottomColor As Long
Private mblnOK As Boolean, mlineType As Long, mlineColor As Long
Public Function ShowMe(lOutBorder As Integer, lLeftBorder As Integer, lRightBorder As Integer, lTopBorder As Integer, lBottomBorder As Integer, lShade As Integer, _
    lOutColor As Long, lLeftColor As Long, lRightColor As Long, lTopColor As Long, lBottomColor As Long, frmPar As Object) As Boolean
    mblnOK = False: mlineColor = 0: mlineType = 1
    mlOutBorder = lOutBorder: mlLeftBorder = lLeftBorder: mlRightBorder = lRightBorder: mlTopBorder = lTopBorder: mlBottomBorder = lBottomBorder: mlShade = lShade
    mlOutColor = lOutColor: mlLeftColor = lLeftColor: mlRightColor = lRightColor: mlTopColor = lTopColor: mlBottomColor = lBottomColor
    
    '无线
    lineShow(0).Visible = (lTopBorder <> F1BorderNone): lineShow(1).Visible = (lBottomBorder <> F1BorderNone)
    lineShow(2).Visible = (lLeftBorder <> F1BorderNone): lineShow(3).Visible = (lRightBorder <> F1BorderNone)
    
    '有线
    If lTopBorder = F1BorderThin Then lineShow(0).BorderStyle = 1: lineShow(0).BorderWidth = 1
    If lBottomBorder = F1BorderThin Then lineShow(1).BorderStyle = 1: lineShow(1).BorderWidth = 1
    If lLeftBorder = F1BorderThin Then lineShow(2).BorderStyle = 1: lineShow(2).BorderWidth = 1
    If lRightBorder = F1BorderThin Then lineShow(3).BorderStyle = 1: lineShow(3).BorderWidth = 1
    
    '虚线
    If lTopBorder = F1BorderDotted Then lineShow(0).BorderStyle = 3: lineShow(0).BorderWidth = 1
    If lBottomBorder = F1BorderDotted Then lineShow(1).BorderStyle = 3: lineShow(1).BorderWidth = 1
    If lLeftBorder = F1BorderDotted Then lineShow(2).BorderStyle = 3: lineShow(2).BorderWidth = 1
    If lRightBorder = F1BorderDotted Then lineShow(3).BorderStyle = 3: lineShow(3).BorderWidth = 1
    
    '粗线
    If lTopBorder = F1BorderThick Then lineShow(0).BorderStyle = 1: lineShow(0).BorderWidth = 2
    If lBottomBorder = F1BorderThick Then lineShow(1).BorderStyle = 1: lineShow(1).BorderWidth = 2
    If lLeftBorder = F1BorderThick Then lineShow(2).BorderStyle = 1: lineShow(2).BorderWidth = 2
    If lRightBorder = F1BorderThick Then lineShow(3).BorderStyle = 1: lineShow(3).BorderWidth = 2
    
    '线色
    lineShow(0).BorderColor = lTopColor: lineShow(1).BorderColor = lBottomColor
    lineShow(2).BorderColor = lLeftColor: lineShow(3).BorderColor = lRightColor
    
    Me.Show 1, frmPar
    If Not mblnOK Then Exit Function
    ShowMe = True
    lOutBorder = mlOutBorder: lLeftBorder = mlLeftBorder: lRightBorder = mlRightBorder: lTopBorder = mlTopBorder: lBottomBorder = mlBottomBorder: lShade = mlShade
    lOutColor = mlOutColor: lLeftColor = mlLeftColor: lRightColor = mlRightColor: lTopColor = mlTopColor: lBottomColor = mlBottomColor
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mblnOK = True: Unload Me
End Sub

Private Sub ColorPicker1_pOK(ByVal ControlSelf As Boolean)
    mlineColor = ColorPicker1.Color
    If mlineColor = tomAutoColor Then mlineColor = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub optType_Click(Index As Integer)
    mlineType = Index
End Sub

Private Sub picLine_Click(Index As Integer)
    lineShow(Index).Visible = Not lineShow(Index).Visible
    Select Case Index
        Case 0
            If Not lineShow(Index).Visible Then
                mlTopBorder = F1BorderNone
            Else
                mlTopBorder = mlineType: mlTopColor = mlineColor: lineShow(Index).BorderColor = mlineColor
                lineShow(Index).BorderStyle = Decode(mlineType, F1BorderThin, 1, F1BorderDotted, 3, F1BorderThick, 1)
                lineShow(Index).BorderWidth = IIf(mlineType = F1BorderThick, 2, 1)
            End If
        Case 1
            If Not lineShow(Index).Visible Then
                mlBottomBorder = F1BorderNone
            Else
                mlBottomBorder = mlineType: mlBottomColor = mlineColor: lineShow(Index).BorderColor = mlineColor
                lineShow(Index).BorderStyle = Decode(mlineType, F1BorderThin, 1, F1BorderDotted, 3, F1BorderThick, 1)
                lineShow(Index).BorderWidth = IIf(mlineType = F1BorderThick, 2, 1)
            End If
        Case 2
            If Not lineShow(Index).Visible Then
                mlLeftBorder = F1BorderNone
            Else
                mlLeftBorder = mlineType: mlLeftColor = mlineColor: lineShow(Index).BorderColor = mlineColor
                lineShow(Index).BorderStyle = Decode(mlineType, F1BorderThin, 1, F1BorderDotted, 3, F1BorderThick, 1)
                lineShow(Index).BorderWidth = IIf(mlineType = F1BorderThick, 2, 1)
            End If
        Case 3
            If Not lineShow(Index).Visible Then
                mlRightBorder = F1BorderNone
            Else
                mlRightBorder = mlineType: mlRightColor = mlineColor: lineShow(Index).BorderColor = mlineColor
                lineShow(Index).BorderStyle = Decode(mlineType, F1BorderThin, 1, F1BorderDotted, 3, F1BorderThick, 1)
                lineShow(Index).BorderWidth = IIf(mlineType = F1BorderThick, 2, 1)
            End If
    End Select
End Sub
