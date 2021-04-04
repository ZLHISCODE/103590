VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmLabMicrobe3Report 
   BorderStyle     =   0  'None
   Caption         =   "微生物三级报告"
   ClientHeight    =   7830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   11100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtReport3 
      Height          =   1755
      Left            =   270
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   5520
      Width           =   8385
   End
   Begin VB.TextBox txtReport2 
      Height          =   1755
      Left            =   270
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   3150
      Width           =   8385
   End
   Begin VB.TextBox txtReport1 
      Height          =   1755
      Left            =   270
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   840
      Width           =   8385
   End
   Begin XtremeSuiteControls.ShortcutCaption srtReport3 
      Height          =   405
      Left            =   450
      TabIndex        =   4
      Top             =   5040
      Width           =   1455
      _Version        =   589884
      _ExtentX        =   2566
      _ExtentY        =   714
      _StockProps     =   6
      Caption         =   "三级报告"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeSuiteControls.ShortcutCaption srtReport2 
      Height          =   405
      Left            =   300
      TabIndex        =   2
      Top             =   2700
      Width           =   1455
      _Version        =   589884
      _ExtentX        =   2566
      _ExtentY        =   714
      _StockProps     =   6
      Caption         =   "二级报告"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeSuiteControls.ShortcutCaption srtReport1 
      Height          =   405
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
      _Version        =   589884
      _ExtentX        =   2566
      _ExtentY        =   714
      _StockProps     =   6
      Caption         =   "一级报告"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmLabMicrobe3Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Event StartEdit(Cancel As Boolean)

Private Sub Form_Resize()
    Dim lngHeight As Long
    
    On Error Resume Next
    
    lngHeight = Me.Height / 3
    
    
    With Me.srtReport1
        .Top = 0
        .Left = 0
        .Width = Me.Width
    End With
    
    With Me.txtReport1
        .Top = Me.srtReport1.Top + Me.srtReport1.Height
        .Left = 0
        .Width = Me.Width
        .Height = lngHeight - Me.srtReport1.Height
    End With
    
    With Me.srtReport2
        .Top = Me.txtReport1.Top + Me.txtReport1.Height
        .Left = 0
        .Width = Me.Width
    End With
    
    With Me.txtReport2
        .Top = Me.srtReport2.Top + Me.srtReport2.Height
        .Left = 0
        .Width = Me.Width
        .Height = Me.txtReport1.Height
    End With
    
    With Me.srtReport3
        .Top = Me.txtReport2.Top + Me.txtReport2.Height
        .Left = 0
        .Width = Me.Width
    End With
    
    With Me.txtReport3
        .Top = Me.srtReport3.Top + Me.srtReport3.Height
        .Left = 0
        .Width = Me.Width
        .Height = Me.txtReport1.Height
    End With
End Sub
Public Function ZlEditStart() As BOOL
    Me.txtReport1.Locked = False
    Me.txtReport2.Locked = False
    Me.txtReport3.Locked = False
    Me.txtReport1.SetFocus
End Function
Public Function zlRefresh(ByVal lngKey As Long) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "select 一级报告,二级报告,三级报告 from 检验标本记录 where id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)
    
    If rsTmp.EOF = False Then
        Me.txtReport1.Text = Nvl(rsTmp("一级报告"))
        Me.txtReport2.Text = Nvl(rsTmp("二级报告"))
        Me.txtReport3.Text = Nvl(rsTmp("三级报告"))
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function ZlCancel() As Boolean
    Me.txtReport1.Locked = True
    Me.txtReport2.Locked = True
    Me.txtReport3.Locked = True
End Function
Public Function ZlSave(lngKey As Long) As Boolean
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim blnRollBak As Boolean
    
    On Error GoTo errH
    
    strSQL = "select id,姓名,性别,年龄,一级报告,二级报告,三级报告 from 检验标本记录 where id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)
    
    If rsTmp.EOF = False Then


        
        strSQL = "Zl_检验标本记录_Update(" & rsTmp("ID") & ",'" & rsTmp("姓名") & "','" & rsTmp("性别") & "','" & rsTmp("年龄") & "'," & _
                "Null,Null,Null,Null,Null,Null,Null,Null,'" & _
                Me.txtReport1.Text & "','" & Me.txtReport2.Text & "','" & Me.txtReport3.Text & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        strSQL = "Zl_检验报告单_Update(" & rsTmp("ID") & ",0,'" & gstrUnitName & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
   End If
    
    Me.txtReport1.Locked = True
    Me.txtReport2.Locked = True
    Me.txtReport3.Locked = True
    ZlSave = True
    
    Exit Function
errH:


    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function

Private Sub txtReport1_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent StartEdit(False)
End Sub

Private Sub txtReport2_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent StartEdit(False)
End Sub

Private Sub txtReport3_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent StartEdit(False)
End Sub
