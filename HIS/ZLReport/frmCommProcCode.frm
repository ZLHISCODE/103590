VERSION 5.00
Object = "{CA73588D-282F-4592-9369-A61CC244FADA}#15.3#0"; "Codejock.SyntaxEdit.v15.3.1.ocx"
Begin VB.Form frmCommProcCode 
   Caption         =   "代码详情"
   ClientHeight    =   9555
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14835
   Icon            =   "frmCommProcCode.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9555
   ScaleWidth      =   14835
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   3000
      Left            =   3480
      ScaleHeight     =   3000
      ScaleWidth      =   5145
      TabIndex        =   0
      Top             =   1440
      Width           =   5145
      Begin XtremeSyntaxEdit.SyntaxEdit txtCode 
         Height          =   1650
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   3090
         _Version        =   983043
         _ExtentX        =   5450
         _ExtentY        =   2910
         _StockProps     =   84
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         EnableSyntaxColorization=   -1  'True
         ShowLineNumbers =   -1  'True
         ShowSelectionMargin=   -1  'True
         ShowScrollBarVert=   -1  'True
         ShowScrollBarHorz=   -1  'True
         EnableVirtualSpace=   0   'False
         EnableAutoIndent=   -1  'True
         ShowWhiteSpace  =   0   'False
         ShowCollapsibleNodes=   -1  'True
         AutoCompleteWndWidth=   160
      End
   End
End
Attribute VB_Name = "frmCommProcCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsData As ADODB.Recordset
Private mblnStartUp As Boolean

Private Sub Form_Activate()
    Dim i As Integer
    Dim strFlag As String
    
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    
    If mrsData.BOF = False Then
        For i = 0 To mrsData.RecordCount - 1
            strFlag = strFlag & Nvl(mrsData("text").Value, "") & vbCrLf
            mrsData.MoveNext
        Next
        If strFlag <> "" Then
            txtCode.Text = strFlag
        End If
    End If
End Sub

Private Sub Form_Load()
    mblnStartUp = True
'    txtCode.SyntaxSet "[Schemes]" & vbCrLf & "SQL" & vbCrLf & "[Themes]" & vbCrLf & "Default" & vbCrLf & "Alternative" & vbCrLf
''    Debug.Print Now
'    txtCode.SyntaxScheme = gstrColor
'    Debug.Print Now
End Sub

Public Function ShowMe(ByVal objMain As Object, ByVal rsData As ADODB.Recordset)
    Set mrsData = rsData
    Me.Show 1, objMain
End Function

Private Sub Form_Resize()
    On Error Resume Next
    picPane.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub SyntaxEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        Unload Me
    End Select
End Sub

Private Sub picPane_Resize()
    txtCode.Move 15, 15, Me.ScaleWidth - 30, Me.ScaleHeight - 30
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        Unload Me
    End Select
End Sub
