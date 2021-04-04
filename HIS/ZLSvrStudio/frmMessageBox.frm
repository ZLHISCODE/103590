VERSION 5.00
Begin VB.Form frmMessageBox 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "中联软件"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5415
   Icon            =   "frmMessageBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5415
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picInformation 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2805
      Index           =   2
      Left            =   0
      ScaleHeight     =   2805
      ScaleWidth      =   5415
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   5420
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "不影响"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   23
         Left            =   2325
         TabIndex        =   34
         Top             =   255
         Width           =   765
      End
      Begin VB.Label lblWarning 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "可能会      文件清单损坏"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   22
         Left            =   1785
         TabIndex        =   33
         Top             =   1215
         Width           =   2160
      End
      Begin VB.Label lblWarningCustom 
         BackStyle       =   0  'Transparent
         Caption         =   "是否确认删除该文件？"
         ForeColor       =   &H000000FF&
         Height          =   750
         Index           =   2
         Left            =   1425
         TabIndex        =   17
         Top             =   2010
         Width           =   3915
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "3.删除可能会使当前文件清单损坏，请谨慎操作"
         Height          =   390
         Index           =   10
         Left            =   1245
         TabIndex        =   16
         Top             =   1215
         Width           =   4125
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "2.删除后可以重新添加"
         Height          =   390
         Index           =   9
         Left            =   1245
         TabIndex        =   15
         Top             =   735
         Width           =   3705
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "1.删除文件后不影响客户端已存在文件"
         Height          =   390
         Index           =   8
         Left            =   1245
         TabIndex        =   14
         Top             =   255
         Width           =   3765
      End
      Begin VB.Image imgMessage 
         Height          =   720
         Index           =   2
         Left            =   255
         Picture         =   "frmMessageBox.frx":0CCA
         Top             =   420
         Width           =   720
      End
   End
   Begin VB.PictureBox picInformation 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2805
      Index           =   1
      Left            =   0
      ScaleHeight     =   2805
      ScaleWidth      =   5415
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   5420
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "弃用的文件              删除后"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   21
         Left            =   1785
         TabIndex        =   32
         Top             =   1155
         Width           =   3240
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "加至当前文件清单"
         Height          =   330
         Index           =   20
         Left            =   1425
         TabIndex        =   31
         Top             =   1395
         Width           =   2685
      End
      Begin VB.Label lblWarning 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "不能还原"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   18
         Left            =   2160
         TabIndex        =   30
         Top             =   690
         Width           =   720
      End
      Begin VB.Label lblWarning 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "自动清除该文件"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   3
         Left            =   3765
         TabIndex        =   29
         Top             =   225
         Width           =   1260
      End
      Begin VB.Image imgMessage 
         Height          =   720
         Index           =   1
         Left            =   255
         Picture         =   "frmMessageBox.frx":280C
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "1.弃用文件后，客户端升级时会"
         Height          =   390
         Index           =   7
         Left            =   1245
         TabIndex        =   12
         Top             =   225
         Width           =   4125
      End
      Begin VB.Label lblWarning 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2.弃用文件        ，只能重新添加"
         Height          =   180
         Index           =   6
         Left            =   1245
         TabIndex        =   11
         Top             =   690
         Width           =   2880
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "3.已经          需要在弃用表中      ，才能添"
         Height          =   240
         Index           =   5
         Left            =   1245
         TabIndex        =   10
         Top             =   1155
         Width           =   4125
      End
      Begin VB.Label lblWarningCustom 
         BackStyle       =   0  'Transparent
         Caption         =   "是否确认弃用当前文件？"
         ForeColor       =   &H000000FF&
         Height          =   735
         Index           =   1
         Left            =   1410
         TabIndex        =   9
         Top             =   2115
         Width           =   3915
      End
   End
   Begin VB.PictureBox picInformation 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2805
      Index           =   0
      Left            =   0
      ScaleHeight     =   2805
      ScaleWidth      =   5415
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   5420
      Begin VB.Label lblWarning 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "修正后需要重新上传所有文件"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   19
         Left            =   1425
         TabIndex        =   28
         Top             =   1440
         Width           =   2340
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "将会不能正常进行"
         Height          =   225
         Index           =   14
         Left            =   1425
         TabIndex        =   22
         Top             =   1695
         Width           =   4155
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "添加的第三方部件会保留"
         Height          =   255
         Index           =   13
         Left            =   1440
         TabIndex        =   21
         Top             =   1110
         Width           =   3435
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "，            时，需要"
         Height          =   240
         Index           =   12
         Left            =   1620
         TabIndex        =   20
         Top             =   480
         Width           =   4485
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "常  自动注册出错        使用该功能修正"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   11
         Left            =   1425
         TabIndex        =   19
         Top             =   480
         Width           =   3945
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "修改文件清单后            升级后使用异"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   1785
         TabIndex        =   18
         Top             =   225
         Width           =   4410
      End
      Begin VB.Label lblWarningCustom 
         BackStyle       =   0  'Transparent
         Caption         =   "是否确认修正当前在用文件清单？"
         ForeColor       =   &H000000FF&
         Height          =   645
         Index           =   0
         Left            =   1440
         TabIndex        =   7
         Top             =   2190
         Width           =   3900
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "3.                          ，否则客户端升级   "
         Height          =   225
         Index           =   2
         Left            =   1245
         TabIndex        =   6
         Top             =   1440
         Width           =   4300
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "2.使用标准文件清单表修正当前文件清单后，用户"
         Height          =   255
         Index           =   1
         Left            =   1245
         TabIndex        =   5
         Top             =   825
         Width           =   4125
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "1.手动              ，造成客户端        "
         Height          =   255
         Index           =   0
         Left            =   1245
         TabIndex        =   4
         Top             =   225
         Width           =   4300
      End
      Begin VB.Image imgMessage 
         Height          =   720
         Index           =   0
         Left            =   255
         Picture         =   "frmMessageBox.frx":434E
         Top             =   420
         Width           =   720
      End
   End
   Begin VB.PictureBox picInformation 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2805
      Index           =   3
      Left            =   0
      ScaleHeight     =   2805
      ScaleWidth      =   5415
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   5420
      Begin VB.Image imgMessage 
         Height          =   720
         Index           =   3
         Left            =   330
         Picture         =   "frmMessageBox.frx":5E90
         Top             =   510
         Width           =   720
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "1.删除文件后不影响客户端已存在文件"
         Height          =   390
         Index           =   17
         Left            =   1470
         TabIndex        =   27
         Top             =   315
         Width           =   3765
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "2.删除后可以重新添加"
         Height          =   390
         Index           =   16
         Left            =   1470
         TabIndex        =   26
         Top             =   780
         Width           =   3705
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "3.删除可能会使当前文件清单损坏，请谨慎"
         Height          =   390
         Index           =   15
         Left            =   1470
         TabIndex        =   25
         Top             =   1245
         Width           =   3750
      End
      Begin VB.Label lblWarningCustom 
         BackStyle       =   0  'Transparent
         Caption         =   "是否确认删除该文件？"
         Height          =   750
         Index           =   3
         Left            =   1455
         TabIndex        =   24
         Top             =   2010
         Width           =   3915
      End
   End
   Begin VB.PictureBox picOpretion 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   5415
      TabIndex        =   1
      Top             =   2805
      Width           =   5420
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&Q)"
         Height          =   450
         Left            =   3915
         TabIndex        =   3
         Top             =   105
         Width           =   1350
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确认(&A)"
         Height          =   450
         Left            =   2415
         TabIndex        =   2
         Top             =   105
         Width           =   1350
      End
   End
End
Attribute VB_Name = "frmMessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private blnInformation As Boolean

Private Enum MessageMode
    MM_在用文件清单修复 = 0
    MM_弃用文件 = 1
    MM_删除文件 = 2
End Enum

Public Function ShowMe(intMode As Integer, Optional strCaption As String = "", Optional strMessage As String = "") As Boolean
    If strCaption <> "" Then Me.Caption = strCaption

    Select Case intMode
        Case MM_在用文件清单修复
            picInformation(MM_在用文件清单修复).Visible = True
            If strMessage <> "" Then lblWarningCustom(MM_在用文件清单修复).Caption = strMessage
        Case MM_弃用文件
            picInformation(MM_弃用文件).Visible = True
            If strMessage <> "" Then lblWarningCustom(MM_弃用文件).Caption = strMessage
        Case MM_删除文件
            picInformation(MM_删除文件).Visible = True
            If strMessage <> "" Then lblWarningCustom(MM_删除文件).Caption = strMessage
    End Select
    
    Call picInformation(intMode).Move(0, 0, Me.ScaleWidth, Me.ScaleHeight - picOpretion.Height)
    
    Me.Show 1, frmMDIMain
    ShowMe = blnInformation
End Function

Private Sub cmdCancel_Click()
    blnInformation = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    blnInformation = True
    Unload Me
End Sub

