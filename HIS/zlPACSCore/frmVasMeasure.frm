VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVasMeasure 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "血管狭窄测量结果"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4335
   Icon            =   "frmVasMeasure.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command1 
      Caption         =   "关闭(&C)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2880
      TabIndex        =   15
      Top             =   4800
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Caption         =   "测量结果："
      Height          =   3135
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   4095
      Begin VB.TextBox txtResult 
         ForeColor       =   &H8000000C&
         Height          =   300
         Index           =   5
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   2670
         Width           =   2200
      End
      Begin VB.TextBox txtResult 
         ForeColor       =   &H8000000C&
         Height          =   300
         Index           =   4
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2280
         Width           =   2200
      End
      Begin VB.TextBox txtResult 
         ForeColor       =   &H8000000C&
         Height          =   300
         Index           =   3
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1500
         Width           =   2200
      End
      Begin VB.TextBox txtResult 
         ForeColor       =   &H8000000C&
         Height          =   300
         Index           =   2
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1125
         Width           =   2200
      End
      Begin VB.TextBox txtResult 
         ForeColor       =   &H8000000C&
         Height          =   300
         Index           =   1
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   750
         Width           =   2200
      End
      Begin VB.TextBox txtResult 
         ForeColor       =   &H8000000C&
         Height          =   300
         Index           =   0
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   360
         Width           =   2200
      End
      Begin VB.Label Label2 
         Caption         =   "面积狭窄度："
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   21
         Top             =   2700
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "直径狭窄度："
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   20
         Top             =   2310
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "狭窄面积："
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   19
         Top             =   1523
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "正常面积："
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   1148
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "狭窄直径："
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   773
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "正常直径："
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   383
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000F&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00808080&
         FillStyle       =   7  'Diagonal Cross
         Height          =   135
         Left            =   240
         Top             =   2040
         Width           =   3600
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "狭窄血管："
      Height          =   1215
      Index           =   1
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   1900
      Begin VB.TextBox txtThreshold 
         Height          =   300
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin MSComCtl2.UpDown updThreshold 
         Height          =   300
         Index           =   1
         Left            =   1200
         TabIndex        =   5
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "自动测量阈值"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "正常血管："
      Height          =   1215
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1900
      Begin MSComCtl2.UpDown updThreshold 
         Height          =   300
         Index           =   0
         Left            =   1200
         TabIndex        =   3
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtThreshold 
         Height          =   300
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "自动测量阈值"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmVasMeasure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lblText As DicomLabel
Public f As frmViewer
Dim lblStandard As DicomLabel
Dim lblNarrow As DicomLabel

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim lblTemp1 As DicomLabel
    Dim lblTemp2 As DicomLabel
    Set lblTemp1 = lblText
    
    If lblTemp1.Tag = "VAS1T" Or lblTemp1.Tag = "VAS2T" Then
        For i = 1 To 3
            If Not lblTemp1.TagObject Is Nothing Then
                Set lblTemp1 = lblTemp1.TagObject
            Else
                Exit Sub
            End If
        Next i
        Set lblTemp2 = lblTemp1
        For i = 1 To 4
            If Not lblTemp2.TagObject Is Nothing Then
                Set lblTemp2 = lblTemp2.TagObject
            Else
                Exit Sub
            End If
        Next i
    Else
        Exit Sub
    End If
    
    If lblText.Tag = "VAS1T" Then
        Set lblNarrow = lblTemp1
        Set lblStandard = lblTemp2
    Else
        Set lblStandard = lblTemp1
        Set lblNarrow = lblTemp2
    End If
    If funGetMeasureResult = True Then
        '填写阈值
        Me.txtThreshold(0).Text = Right(lblStandard.Text, Len(lblStandard.Text) - InStr(lblStandard.Text, ":")) 'intNarrowThreshold
        Me.txtThreshold(1).Text = Right(lblNarrow.Text, Len(lblNarrow.Text) - InStr(lblNarrow.Text, ":")) '  intStandardThreshold
        subModifyThreshold 1
    End If
End Sub

Private Function funGetMeasureResult() As Boolean
'------------------------------------------------
'功能：填写血管狭窄测量的结果
'参数：无
'返回：True-正常填写完成；False－填写过程出错。
'上级函数或过程：Form_Load
'下级函数或过程：无
'引用的外部参数：lblNarrow,lblStandard
'编制人：黄捷 2005-8-22
'------------------------------------------------
    Dim strResult As String
    Dim lngNarrowDiameter As Long
    Dim lngNarrowArea As Long
    Dim lngStandardDiameter As Long
    Dim lngStandardArea As Long
    If lblText Is Nothing Then Exit Function
    '填写狭窄血管直径和面积
    strResult = lblNarrow.TagObject.Text
    Me.txtResult(1).Text = Mid(strResult, InStr(strResult, "：") + 1, InStr(strResult, vbCrLf) - InStr(strResult, "：") - 1)
    lngNarrowDiameter = Val(left(Me.txtResult(1).Text, InStr(Me.txtResult(1).Text, "(") - 1))
    strResult = Right(strResult, Len(strResult) - InStr(strResult, vbCrLf))
    Me.txtResult(3).Text = Mid(strResult, InStr(strResult, "：") + 1, InStr(strResult, ")") - InStr(strResult, "："))
    lngNarrowArea = Val(left(Me.txtResult(3).Text, InStr(Me.txtResult(3).Text, "(") - 1))
    '填写正常血管直径和面积
    strResult = lblStandard.TagObject.Text
    Me.txtResult(0).Text = Mid(strResult, InStr(strResult, "：") + 1, InStr(strResult, vbCrLf) - InStr(strResult, "：") - 1)
    lngStandardDiameter = Val(left(Me.txtResult(0).Text, InStr(Me.txtResult(0).Text, "(") - 1))
    strResult = Right(strResult, Len(strResult) - InStr(strResult, vbCrLf))
    Me.txtResult(2).Text = Right(strResult, Len(strResult) - InStr(strResult, "："))
    lngStandardArea = Val(left(Me.txtResult(2).Text, InStr(Me.txtResult(2).Text, "(") - 1))
    '填写血管直径狭窄度和面积狭窄度
    Me.txtResult(4).Text = Format((lngNarrowDiameter / lngStandardDiameter * 100), "0.000") & "%"
    Me.txtResult(5).Text = Format((lngNarrowArea / lngStandardArea * 100), "0.000") & "%"
    funGetMeasureResult = True
End Function

Private Sub txtThreshold_Change(Index As Integer)
    If Val(Me.txtThreshold(Index).Text) > 500 Then
        Me.txtThreshold(Index).Text = 500
    ElseIf Val(Me.txtThreshold(Index).Text) < 1 Then
        Me.txtThreshold(Index).Text = 1
    End If
End Sub

Private Sub txtThreshold_Click(Index As Integer)
    Me.txtThreshold(Index).SelStart = 0
    Me.txtThreshold(Index).SelLength = Len(Me.txtThreshold(Index).Text)
End Sub

Private Sub txtThreshold_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then subModifyThreshold Index + 1
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub updThreshold_DownClick(Index As Integer)
    If Me.txtThreshold(Index).Text > 1 Then
        Me.txtThreshold(Index).Text = Me.txtThreshold(Index).Text - 1
    Else
        Me.txtThreshold(Index).Text = 1
    End If
    subModifyThreshold Index + 1
End Sub

Private Sub updThreshold_UpClick(Index As Integer)
    If Me.txtThreshold(Index).Text < 500 Then
        Me.txtThreshold(Index).Text = Me.txtThreshold(Index).Text + 1
    Else
        Me.txtThreshold(Index).Text = 500
    End If
    subModifyThreshold Index + 1
End Sub

Private Sub subModifyThreshold(intVas As Integer)
'intVas=1 正常血管；2－狭窄血管。
    Dim lblVas As DicomLabel
    Dim intThreshold As Integer
    Dim x1 As Long, x2 As Long, y1 As Long, y2 As Long
    Dim lngDiameter As Long
    Dim lngArea As Long
    Dim l As DicomLabel
    Dim lngNarrowDiameter As Long
    Dim lngStandardDiameter As Long
    
    intStandardThreshold = Me.txtThreshold(0).Text
    intNarrowThreshold = Me.txtThreshold(1).Text
    lngStandardDiameter = Val(left(lblStandard.Text, InStr(lblStandard.Text, ":") - 1))
    lngNarrowDiameter = Val(left(lblNarrow.Text, InStr(lblNarrow.Text, ":") - 1))
    
    If intVas = 1 Then  '正常血管
        Set lblVas = lblStandard
    Else                '狭窄血管
        Set lblVas = lblNarrow
    End If
    
    '计算血管壁
    If funDrawVas(lblVas, f.SelectedImage, intVas) = True Then
        funGetMeasureResult
        lblStandard.TagObject.Text = "正常血管直径：" & Me.txtResult(0).Text & vbCrLf _
                                & "血管面积：" & Me.txtResult(2).Text
        lblNarrow.TagObject.Text = "狭窄血管直径：" & Me.txtResult(1).Text & vbCrLf _
                                    & "血管面积：" & Me.txtResult(3).Text & vbCrLf _
                                    & "直径狭窄度：" & Me.txtResult(4).Text & vbCrLf _
                                    & "面积狭窄度：" & Me.txtResult(5).Text
    End If
    f.SelectedImage.Refresh False
End Sub
