VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPlanCopy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�Űิ��"
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
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   1920
      TabIndex        =   1
      Top             =   1440
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker Dtp����ʱ�� 
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
         Name            =   "����"
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
   Begin MSComCtl2.DTPicker Dtp����ʱ�� 
      Height          =   300
      Left            =   1320
      TabIndex        =   3
      Top             =   780
      Width           =   1692
      _ExtentX        =   2985
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
   Begin VB.Label lbl����ʱ�� 
      AutoSize        =   -1  'True
      Caption         =   "����ʱ��"
      Height          =   180
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Width           =   720
   End
   Begin VB.Label lbl����ʱ�� 
      AutoSize        =   -1  'True
      Caption         =   "����ʱ��"
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

Private mstr����ʱ�� As String
Private mlng����id As Long

Public Sub ShowCard(FrmMain As Form, ByVal str����ʱ�� As String, ByVal lng����id As Long)
    mstr����ʱ�� = str����ʱ��
    mlng����id = lng����id
    
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
    Dim bln������ As Boolean
    
    If Dtp����ʱ��.Value > zldatabase.Currentdate Then
        If MsgBox("����ʱ����ڽ��죬�Ƿ�����Űࣿ", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        End If
    End If
    
    With frmPlan.vsfPlan
        bln������ = False
        For i = 1 To .rows - 1
            If .TextMatrix(i, .ColIndex("�����")) <> "" Or .TextMatrix(i, .ColIndex("��ҩ��")) <> "" Or .TextMatrix(i, .ColIndex("�˶���")) <> "" Or .TextMatrix(i, .ColIndex("��Һ��")) <> "" Or .TextMatrix(i, .ColIndex("������")) <> "" Then
                bln������ = True
                Exit For
            End If
        Next
    End With
    
    If bln������ = False Then
        MsgBox "�������ݲ�����Ϊ��!"
        Unload Me
        Exit Sub
    End If
    
    arrSql = Array()
    With frmPlan.vsfPlan
        For i = 1 To .rows - 1
            
            If Val(.TextMatrix(i, .ColIndex("��Һ̨id"))) = 0 Then
                Exit Sub
            End If
            strsql = "Zl_��Һ��������_����("
            strsql = strsql & mlng����id
            strsql = strsql & ",to_date('" & Format(Dtp����ʱ��.Value, "Short Date") & "' ,'yyyy-mm-dd')"
            strsql = strsql & "," & Val(.TextMatrix(i, .ColIndex("��Һ̨id")))
            strsql = strsql & ",'" & .TextMatrix(i, .ColIndex("��ҩ����")) & "'"
            strsql = strsql & ",'" & .TextMatrix(i, .ColIndex("�����")) & "'"
            strsql = strsql & ",'" & .TextMatrix(i, .ColIndex("��ҩ��")) & "'"
            strsql = strsql & ",'" & .TextMatrix(i, .ColIndex("�˶���")) & "'"
            strsql = strsql & ",'" & .TextMatrix(i, .ColIndex("��Һ��")) & "'"
            strsql = strsql & ",'" & .TextMatrix(i, .ColIndex("������")) & "'"
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
    
    MsgBox Format(Dtp����ʱ��.Value, "yyyy-MM-dd") & "  �Űิ�Ƴɹ�"

    Exit Sub
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Dtp����ʱ��_Validate(Cancel As Boolean)
    If Dtp����ʱ��.Value < CDate(Format(zldatabase.Currentdate, "yyyy-MM-dd") & " 00:00:00") Then
        MsgBox "����ʱ�䲻��С�ڵ��죡"
        Cancel = True
    End If
End Sub

Private Sub Form_Load()
    Dtp����ʱ��.Value = mstr����ʱ��
    Dtp����ʱ��.Value = zldatabase.Currentdate
    
End Sub
