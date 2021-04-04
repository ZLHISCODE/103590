VERSION 5.00
Begin VB.Form frmCharCount 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "字数统计"
   ClientHeight    =   3045
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5160
   Icon            =   "frmCharCount.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "关闭(&C)"
      Height          =   350
      Left            =   3585
      TabIndex        =   17
      Top             =   75
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   405
      TabIndex        =   16
      Top             =   2790
      Width           =   4305
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   405
      TabIndex        =   15
      Top             =   555
      Width           =   4305
   End
   Begin VB.Label lblValue 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "-"
      Height          =   180
      Index           =   6
      Left            =   4185
      TabIndex        =   14
      Top             =   2505
      Width           =   90
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      Caption         =   "双字节字符数"
      Height          =   180
      Index           =   6
      Left            =   630
      TabIndex        =   13
      Top             =   2505
      Width           =   1080
   End
   Begin VB.Label lblValue 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "-"
      Height          =   180
      Index           =   5
      Left            =   4185
      TabIndex        =   12
      Top             =   2205
      Width           =   90
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      Caption         =   "单字节字符数"
      Height          =   180
      Index           =   5
      Left            =   630
      TabIndex        =   11
      Top             =   2205
      Width           =   1080
   End
   Begin VB.Label lblValue 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "-"
      Height          =   180
      Index           =   4
      Left            =   4185
      TabIndex        =   10
      Top             =   1905
      Width           =   90
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      Caption         =   "字符数(不计空格)"
      Height          =   180
      Index           =   4
      Left            =   630
      TabIndex        =   9
      Top             =   1905
      Width           =   1440
   End
   Begin VB.Label lblValue 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "-"
      Height          =   180
      Index           =   3
      Left            =   4185
      TabIndex        =   8
      Top             =   1590
      Width           =   90
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      Caption         =   "字符数(计算空格)"
      Height          =   180
      Index           =   3
      Left            =   630
      TabIndex        =   7
      Top             =   1590
      Width           =   1440
   End
   Begin VB.Label lblValue 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "-"
      Height          =   180
      Index           =   2
      Left            =   4185
      TabIndex        =   6
      Top             =   1290
      Width           =   90
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      Caption         =   "段落数"
      Height          =   180
      Index           =   2
      Left            =   630
      TabIndex        =   5
      Top             =   1290
      Width           =   540
   End
   Begin VB.Label lblValue 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "-"
      Height          =   180
      Index           =   1
      Left            =   4185
      TabIndex        =   4
      Top             =   990
      Width           =   90
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      Caption         =   "行数"
      Height          =   180
      Index           =   1
      Left            =   630
      TabIndex        =   3
      Top             =   990
      Width           =   360
   End
   Begin VB.Label lblValue 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "-"
      Height          =   180
      Index           =   0
      Left            =   4185
      TabIndex        =   2
      Top             =   690
      Width           =   90
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      Caption         =   "页数"
      Height          =   180
      Index           =   0
      Left            =   630
      TabIndex        =   1
      Top             =   690
      Width           =   360
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "统计信息"
      Height          =   180
      Left            =   450
      TabIndex        =   0
      Top             =   315
      Width           =   720
   End
End
Attribute VB_Name = "frmCharCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Function ShowMe(Editor As Editor) As Boolean
    '功能：显示本对话框
    '参数：
    '   Editor,编辑器对象
    
    Dim strText As String, lngNum As Long
    Dim lngWidth1  As Long, lngWidth2 As Long
    
    '页数
    Me.lblValue(0).Caption = Editor.PageCount
    '行数
    Me.lblValue(1).Caption = Editor.LineCount
    
    strText = Editor.Text
    
    '段落数
    lngNum = UBound(Split(strText, vbCrLf))
    If lngNum = -1 Then
        Me.lblValue(2).Caption = 1
    Else
        Me.lblValue(2).Caption = lngNum + 1
    End If
    '字数(计算空格)
    strText = Replace(strText, vbCrLf, "")
    strText = Replace(strText, vbTab, "")
    Me.lblValue(3).Caption = Len(strText)
    
    '字数(不计空格)
    strText = Replace(strText, " ", "")
    Me.lblValue(4).Caption = Len(strText)

    '单字节字符数，在宋体5号字体下，TextWidth返回指只能大约2700个全角字符
    lngWidth1 = 0: lngWidth2 = 0
    For lngNum = 0 To Val(Me.lblValue(4).Caption) Step 2000
        lngWidth1 = lngWidth1 + Me.TextWidth(Mid(strText, lngNum + 1, 2000))
        lngWidth2 = lngWidth2 + Me.TextWidth(StrConv(Mid(strText, lngNum + 1, 2000), vbWide))
    Next
    Me.lblValue(5).Caption = (lngWidth2 - lngWidth1) / Me.TextWidth("A")
    
    '双字节字符数
    Me.lblValue(6).Caption = Val(Me.lblValue(4).Caption) - Val(Me.lblValue(5).Caption)
    
    Me.Show vbModal
    ShowMe = True
End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub
