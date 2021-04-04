VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmSetStatics 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "统计信息设置"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6165
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSetStatics.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin RichTextLib.RichTextBox txtSql 
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5106
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmSetStatics.frx":4D4A
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   5040
      TabIndex        =   2
      Top             =   3120
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3960
      TabIndex        =   1
      Top             =   3120
      Width           =   990
   End
   Begin VB.CheckBox chkLock 
      Caption         =   "锁定统计信息"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3168
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.Label lblHundredMin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "缩小100倍"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   2040
      MouseIcon       =   "frmSetStatics.frx":4DDB
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   3210
      Width           =   810
   End
   Begin VB.Label lblHundred 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "扩大100倍"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   3000
      MouseIcon       =   "frmSetStatics.frx":4F2D
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3210
      Width           =   810
   End
End
Attribute VB_Name = "frmSetStatics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrOwner As String
Private mstrTable As String
Private mblnResult As Boolean

Public Function ShowSet(ByVal strSql As String, ByVal strOwner As String, ByVal strTable As String) As Boolean
    
    txtSql.Text = strSql
    mstrOwner = strOwner
    mstrTable = strTable
    
    ChangeColor
    Me.Show 1
    
    ShowSet = mblnResult
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub ChangeColor()
    '功能改变文本框中注释的颜色
    
    With txtSql
        .SelStart = 0
        .SelLength = Len(txtSql.Text)
        .SelColor = vbBlack
        
        .SelStart = 0
        .SelLength = InStr(1, txtSql.Text, "sys.", vbTextCompare) - 1
        .SelColor = vbBlue
        
        
    End With
    
End Sub

Private Sub cmdOK_Click()
    
    On Error GoTo errH
    gcnOracle.Execute "Begin  " & txtSql.Text & "   End;"
    
    If chkLock.Value = vbChecked Then
        gcnOracle.Execute "Begin" & vbNewLine & _
                                    "  Sys.Dbms_Stats.Lock_Table_Stats(Ownname => '" & mstrOwner & "', Tabname => '" & mstrTable & "');" & vbNewLine & _
                                    "End;"
    End If
    
    mblnResult = True
    Unload Me
    Exit Sub
errH:
    ErrCenter
End Sub

Private Sub lblHundredMin_Click()
    HundredNum "Numrows =>", 2
    HundredNum "Numblks =>", 2
    HundredNum "Numlblks =>", 2
    HundredNum "Distcnt =>", 2
    HundredNum "Numdist =>", 2
    HundredNum "Nullcnt =>", 2
    ChangeColor
End Sub

Private Sub lblHundred_Click()
    HundredNum "Numrows =>", 1
    HundredNum "Numblks =>", 1
    HundredNum "Numlblks =>", 1
    HundredNum "Distcnt =>", 1
    HundredNum "Numdist =>", 1
    HundredNum "Nullcnt =>", 1
    ChangeColor
End Sub


Private Sub HundredNum(ByVal strFind As String, ByVal intType As Integer)
    '功能:将文本框中传入字符后的数据扩大或缩小100倍
    'intType =1 扩大,  intType=2 缩小
    Dim strTmp As String, lngNumrows As Long
    
    On Error Resume Next
    
    strTmp = txtSql.Text
    
    strTmp = Mid(strTmp, InStr(1, strTmp, strFind, vbTextCompare) + Len(strFind))
    strTmp = Mid(strTmp, 1, InStr(1, strTmp, ",", vbTextCompare) - 1)
    
    If intType = 1 Then
        lngNumrows = IIf(Val(strTmp) = 0, 100, Val(strTmp) * 100)
    Else
        lngNumrows = IIf(Val(strTmp) = 0, strTmp, Val(strTmp) \ 100)
    End If
     
    If Err.Number <> 0 Then '发生溢出错误
        lngNumrows = Val(strTmp)
    End If
    
    txtSql.Text = Mid(txtSql.Text, 1, InStr(1, txtSql.Text, strFind) - 1) & Replace(txtSql.Text, Trim(strTmp), lngNumrows, InStr(1, txtSql.Text, strFind), 1)
End Sub


