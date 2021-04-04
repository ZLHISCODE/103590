VERSION 5.00
Begin VB.Form frmMicrobeFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ϸ������..."
   ClientHeight    =   2655
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5595
   ControlBox      =   0   'False
   Icon            =   "frmMicrobeFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkUpper 
      Caption         =   "���ִ�Сд(&U)"
      Height          =   210
      Left            =   1920
      TabIndex        =   7
      Top             =   1320
      Width           =   2040
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4290
      TabIndex        =   3
      Top             =   2160
      Width           =   1100
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "������һ��(&N)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2715
      TabIndex        =   2
      Top             =   2160
      Width           =   1530
   End
   Begin VB.Frame fraLine 
      Height          =   60
      Left            =   0
      TabIndex        =   4
      Top             =   2010
      Width           =   5565
   End
   Begin VB.ComboBox cboSource 
      Height          =   300
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   930
      Width           =   3435
   End
   Begin VB.Label lblComment 
      Caption         =   "    ����ϣ�����ҵļ���ϸ���ı��롢��������Ӣ����������ڶ�����������""������һ��""��ֱ���ҵ���ϣ�����ҵ���Ŀ��"
      Height          =   525
      Left            =   885
      TabIndex        =   6
      Top             =   135
      Width           =   4500
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "(�����ҵ�10������ǰΪ��1��)"
      Height          =   180
      Left            =   870
      TabIndex        =   5
      Top             =   1680
      Width           =   2430
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   195
      Picture         =   "frmMicrobeFind.frx":058A
      Top             =   150
      Width           =   480
   End
   Begin VB.Label lblSource 
      AutoSize        =   -1  'True
      Caption         =   "��������(&F)"
      Height          =   180
      Left            =   885
      TabIndex        =   0
      Top             =   1005
      Width           =   990
   End
End
Attribute VB_Name = "frmMicrobeFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsFind As New ADODB.Recordset
Private strCurSql As String

Dim intCount As Integer

Private Sub cboSource_Click()
    If Trim(Me.cboSource.Text) <> "" Then
        Me.cmdFind.Enabled = True
    Else
        Me.cmdFind.Enabled = False
    End If
End Sub

Private Sub cboSource_GotFocus()
    Me.cboSource.SelStart = 0: Me.cboSource.SelLength = 100
End Sub

Private Sub cboSource_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cboSource_KeyUp(KeyCode As Integer, Shift As Integer)
    If Trim(Me.cboSource.Text) <> "" Then
        Me.cmdFind.Enabled = True
    Else
        Me.cmdFind.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdFind_Click()
    Dim lngItemID As Long
    Dim strFind As String
    
    If Trim(Me.cboSource.Text) = "" Then
        MsgBox "��������ҵ�����", vbExclamation, gstrSysName
        Me.cboSource.SetFocus: Exit Sub
    End If
    strFind = ""
    For intCount = 0 To Me.cboSource.ListCount
        strFind = strFind & ";" & Me.cboSource.List(intCount)
    Next
    If InStr(1, strFind, ";" & Trim(Me.cboSource.Text)) = 0 Then
        Me.cboSource.AddItem Trim(Me.cboSource.Text), 0
    End If
    
    If Me.chkUpper.Value = 0 Then
        strFind = Replace(UCase(Trim(Me.cboSource.Text)), "'", "")
        gstrSql = "Select M.ID, M.����, M.������, M.Ӣ����, M.����" & vbNewLine & _
                "From ����ϸ�� M" & vbNewLine & _
                "Where ���� Like '" & strFind & "%' Or Upper(������) Like '" & gstrMatch & strFind & _
                "%' Or Upper(Ӣ����) Like '" & gstrMatch & strFind & "%' Or Upper(����) Like '" & gstrMatch & strFind & "%'"
                
    Else
        strFind = Replace(Trim(Me.cboSource.Text), "'", "")
        gstrSql = "Select M.ID, M.����, M.������, M.Ӣ����, M.����" & vbNewLine & _
                "From ����ϸ�� M" & vbNewLine & _
                "Where ���� Like '" & strFind & "%' Or Upper(������) Like '" & gstrMatch & strFind & _
                "%' Or Upper(Ӣ����) Like '" & gstrMatch & strFind & "%' Or Upper(����) Like '" & gstrMatch & strFind & "%'"
        
    End If
'    If frmLabItems.mblnShowStop = False Then
'        gstrSql = gstrSql & "       And (����ʱ�� Is null Or ����ʱ��=To_date('3000-01-01','YYYY-MM-DD'))"
'    End If
    
    Err = 0: On Error GoTo ErrHand
    With rsFind
        If strCurSql <> gstrSql Or .State <> adStateOpen Then
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
            Set rsFind = zldatabase.OpenSQLRecord(gstrSql, "cmdFind_Click")
'            Call SQLTest
            If rsFind.EOF Then
                MsgBox "�����ڲ��ҵ����ݣ�", vbExclamation, gstrSysName
                Me.cmdFind.Enabled = False: Me.lblNote.Caption = ""
                Me.cboSource.SetFocus: Exit Sub
            End If
            strCurSql = gstrSql
        Else
            .MoveNext
            If .EOF Then
                MsgBox "�Ѳ��ҵ����һ����Ŀ��", vbExclamation, gstrSysName
                .Close: Me.cboSource.Text = "": Me.cmdFind.Enabled = False: Me.lblNote.Caption = ""
                Me.cboSource.SetFocus: Exit Sub
            End If
        End If
        Me.lblNote.Caption = "(�����ҵ�" & rsFind.RecordCount & "������ǰΪ��" & rsFind.AbsolutePosition & "��)"
        lngItemID = rsFind!ID
    End With
    
    Call FrmMicrobeList.zlRefList(lngItemID)
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    Me.cboSource.SetFocus
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    strCurSql = ""
    Me.lblNote.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub
