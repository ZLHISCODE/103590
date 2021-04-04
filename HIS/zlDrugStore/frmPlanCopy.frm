VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPlanCopy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "排班复制"
   ClientHeight    =   2124
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   3540
   Icon            =   "frmPlanCopy.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2124
   ScaleWidth      =   3540
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   1920
      TabIndex        =   1
      Top             =   1440
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker Dtp复制时间 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   180
      Width           =   1692
      _ExtentX        =   2985
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   111542275
      CurrentDate     =   39998
   End
   Begin MSComCtl2.DTPicker Dtp生成时间 
      Height          =   300
      Left            =   1320
      TabIndex        =   3
      Top             =   780
      Width           =   1692
      _ExtentX        =   2985
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   111542275
      CurrentDate     =   39998
   End
   Begin VB.Label lbl生成时间 
      AutoSize        =   -1  'True
      Caption         =   "生成时间"
      Height          =   180
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Width           =   720
   End
   Begin VB.Label lbl复制时间 
      AutoSize        =   -1  'True
      Caption         =   "复制时间"
      Height          =   180
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmPlanCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstr复制时间 As String
Private mlng部门id As Long

Public Sub ShowCard(FrmMain As Form, ByVal str复制时间 As String, ByVal lng部门id As Long)
    mstr复制时间 = str复制时间
    mlng部门id = lng部门id
    
    Me.Show vbModal, FrmMain
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    Dim strsql As String
    Dim i  As Integer
    Dim arrSql As Variant
    Dim j As Integer
    Dim bln允许复制 As Boolean
    
    If Dtp复制时间.Value > zldatabase.Currentdate Then
        If MsgBox("复制时间大于今天，是否继续排班？", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        End If
    End If
    
    With frmPlan.vsfPlan
        bln允许复制 = False
        For i = 1 To .rows - 1
            If .TextMatrix(i, .ColIndex("审核人")) <> "" Or .TextMatrix(i, .ColIndex("摆药人")) <> "" Or .TextMatrix(i, .ColIndex("核对人")) <> "" Or .TextMatrix(i, .ColIndex("配液人")) <> "" Or .TextMatrix(i, .ColIndex("复核人")) <> "" Then
                bln允许复制 = True
                Exit For
            End If
        Next
    End With
    
    If bln允许复制 = False Then
        MsgBox "复制内容不允许为空!"
        Unload Me
        Exit Sub
    End If
    
    arrSql = Array()
    With frmPlan.vsfPlan
        For i = 1 To .rows - 1
            
            If Val(.TextMatrix(i, .ColIndex("配液台id"))) = 0 Then
                Exit Sub
            End If
            strsql = "Zl_配液工作安排_设置("
            strsql = strsql & mlng部门id
            strsql = strsql & ",to_date('" & Format(Dtp生成时间.Value, "Short Date") & "' ,'yyyy-mm-dd')"
            strsql = strsql & "," & Val(.TextMatrix(i, .ColIndex("配液台id")))
            strsql = strsql & ",'" & .TextMatrix(i, .ColIndex("配药批次")) & "'"
            strsql = strsql & ",'" & .TextMatrix(i, .ColIndex("审核人")) & "'"
            strsql = strsql & ",'" & .TextMatrix(i, .ColIndex("摆药人")) & "'"
            strsql = strsql & ",'" & .TextMatrix(i, .ColIndex("核对人")) & "'"
            strsql = strsql & ",'" & .TextMatrix(i, .ColIndex("配液人")) & "'"
            strsql = strsql & ",'" & .TextMatrix(i, .ColIndex("复核人")) & "'"
            strsql = strsql & "," & i
            strsql = strsql & ")"

            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = strsql
        Next
    End With
    
    On Error GoTo errHandle
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSql)
        Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "CmdSave_Click")
    Next
    gcnOracle.CommitTrans
    
    MsgBox Format(Dtp生成时间.Value, "yyyy-MM-dd") & "  排班复制成功"

    Exit Sub
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Dtp生成时间_Validate(Cancel As Boolean)
    If Dtp生成时间.Value < CDate(Format(zldatabase.Currentdate, "yyyy-MM-dd") & " 00:00:00") Then
        MsgBox "生成时间不能小于当天！"
        Cancel = True
    End If
End Sub

Private Sub Form_Load()
    Dtp复制时间.Value = mstr复制时间
    Dtp生成时间.Value = zldatabase.Currentdate
    
End Sub
