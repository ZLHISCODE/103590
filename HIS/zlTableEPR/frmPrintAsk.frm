VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrintAsk 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ӡ"
   ClientHeight    =   3795
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6600
   Icon            =   "frmPrintAsk.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5370
      TabIndex        =   18
      Top             =   3360
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4155
      TabIndex        =   17
      Top             =   3360
      Width           =   1100
   End
   Begin VB.Frame fraCopy 
      Caption         =   "����"
      Height          =   2145
      Left            =   3300
      TabIndex        =   12
      Top             =   1110
      Width           =   3225
      Begin VB.CheckBox chkCopyOrder 
         Caption         =   "��ݴ�ӡ(&T)"
         Height          =   195
         Left            =   1800
         TabIndex        =   16
         Top             =   1335
         Value           =   1  'Checked
         Width           =   1305
      End
      Begin VB.TextBox txtCopies 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1140
         MaxLength       =   6
         TabIndex        =   14
         Text            =   "1"
         Top             =   240
         Width           =   1635
      End
      Begin MSComCtl2.UpDown udCopies 
         Height          =   300
         Left            =   2775
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtCopies"
         BuddyDispid     =   196613
         OrigLeft        =   1830
         OrigTop         =   893
         OrigRight       =   2070
         OrigBottom      =   1178
         Max             =   32767
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Image imgCopyOrder 
         Height          =   540
         Index           =   1
         Left            =   135
         Picture         =   "frmPrintAsk.frx":000C
         Top             =   1185
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Image imgCopyOrder 
         Height          =   720
         Index           =   0
         Left            =   195
         Picture         =   "frmPrintAsk.frx":2A7E
         Top             =   1095
         Width           =   1380
      End
      Begin VB.Label lblCopies 
         AutoSize        =   -1  'True
         Caption         =   "����(&P)"
         Height          =   180
         Left            =   315
         TabIndex        =   13
         Top             =   300
         Width           =   630
      End
   End
   Begin VB.Frame fraPageScope 
      Caption         =   "ҳ�淶Χ"
      Height          =   2145
      Left            =   60
      TabIndex        =   4
      Top             =   1110
      Width           =   3165
      Begin VB.ComboBox cboPageOddEven 
         Height          =   300
         ItemData        =   "frmPrintAsk.frx":5E80
         Left            =   900
         List            =   "frmPrintAsk.frx":5E8D
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1710
         Width           =   1875
      End
      Begin VB.TextBox txtPageScope 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1560
         TabIndex        =   8
         Top             =   855
         Width           =   1215
      End
      Begin VB.OptionButton optPageScope 
         Caption         =   "ҳ�뷶Χ(&G)"
         Height          =   180
         Index           =   2
         Left            =   195
         TabIndex        =   7
         Top             =   915
         Width           =   1425
      End
      Begin VB.OptionButton optPageScope 
         Caption         =   "��ǰҳ(&T)"
         Height          =   180
         Index           =   1
         Left            =   195
         TabIndex        =   6
         Top             =   622
         Width           =   1425
      End
      Begin VB.OptionButton optPageScope 
         Caption         =   "ȫ��(&A)"
         Height          =   180
         Index           =   0
         Left            =   195
         TabIndex        =   5
         Top             =   330
         Value           =   -1  'True
         Width           =   1425
      End
      Begin VB.Label lblPageOddEven 
         AutoSize        =   -1  'True
         Caption         =   "��ӡ(&E)"
         Height          =   180
         Left            =   195
         TabIndex        =   10
         Top             =   1770
         Width           =   630
      End
      Begin VB.Label lblPageNote 
         AutoSize        =   -1  'True
         Caption         =   "�����ҳ���/���ö��ŷָ���ҳ�뷶Χ(���磺1,3,5-12)."
         Height          =   540
         Left            =   240
         TabIndex        =   9
         Top             =   1245
         Width           =   2790
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraPrinter 
      Caption         =   "��ӡ��"
      Height          =   975
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6450
      Begin VB.ComboBox cboPrinterName 
         Height          =   300
         Left            =   1665
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   4605
      End
      Begin VB.Image imgPrinter 
         Height          =   360
         Left            =   270
         Picture         =   "frmPrintAsk.frx":5EA3
         Top             =   270
         Width           =   360
      End
      Begin VB.Label lblPrinterInfo 
         AutoSize        =   -1  'True
         Caption         =   "λ��:���ӵ�LTP1:      Ĭ�ϴ�ӡ��:��"
         Height          =   180
         Left            =   1680
         TabIndex        =   3
         Top             =   645
         Width           =   3150
      End
      Begin VB.Label lblPrinterName 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Left            =   945
         TabIndex        =   1
         Top             =   300
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmPrintAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public blnOK As Boolean
Public mstrPageRange As String
Private intCount As Integer

Private Sub cboPrinterName_Click()
    With Me.lblPrinterInfo
        .Caption = "λ��:���ӵ�" & Printers(Me.cboPrinterName.ListIndex).Port
        .Caption = .Caption & Space(6) & "Ĭ�ϴ�ӡ��:" & IIf(Printers(Me.cboPrinterName.ListIndex).DeviceName = Printer.DeviceName, "��", "��")
    End With
End Sub

Private Sub chkCopyOrder_Click()
    If Me.chkCopyOrder.Value = vbChecked Then
        Me.imgCopyOrder(0).Visible = True
        Me.imgCopyOrder(1).Visible = False
    Else
        Me.imgCopyOrder(0).Visible = False
        Me.imgCopyOrder(1).Visible = True
    End If
End Sub

Private Sub cmdCancel_Click()
    blnOK = False: Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim t As Variant, aryPage() As String, blnError As Boolean
    Dim i As Long
    
    If Me.optPageScope(2).Value = True Then
        'ҳ�뷶Χ
        t = Split(Me.txtPageScope.Text, ",")
        For i = 0 To UBound(t)
            '��Ч�Լ��
            aryPage = Split(t(i), "-")
            blnError = False
            If UBound(aryPage) = 0 Then
                If Val(t(i)) < Split(Me.txtPageScope.Tag, "-")(0) Then blnError = True
                If Val(t(i)) > Split(Me.txtPageScope.Tag, "-")(1) Then blnError = True
            ElseIf UBound(aryPage) = 1 Then
                If Val(Split(t(i), "-")(0)) > Val(Split(t(i), "-")(1)) Then blnError = True
                If Val(Split(t(i), "-")(0)) < Split(Me.txtPageScope.Tag, "-")(0) Then blnError = True
                If Val(Split(t(i), "-")(0)) > Split(Me.txtPageScope.Tag, "-")(1) Then blnError = True
                If Val(Split(t(i), "-")(1)) < Split(Me.txtPageScope.Tag, "-")(0) Then blnError = True
                If Val(Split(t(i), "-")(1)) > Split(Me.txtPageScope.Tag, "-")(1) Then blnError = True
            Else
                blnError = True
            End If
            If blnError = True Then
                MsgBox "ҳ�벻��������Χ" & Me.txtPageScope.Tag & "��", vbExclamation, Me.Caption
                Me.txtPageScope.SetFocus
                Exit Sub
            End If
        Next
        '����ҳ�뷶Χ
        Me.txtPageScope.Tag = Me.txtPageScope.Text
    End If
    blnOK = True: Me.Hide
End Sub

Private Sub Form_Load()
    
    txtPageScope.Tag = mstrPageRange
    txtPageScope.Text = mstrPageRange
    With Me.cboPrinterName
        .Clear
        For intCount = 0 To Printers.Count - 1
            .AddItem Printers(intCount).DeviceName
            If Printers(intCount).DeviceName = Printer.DeviceName Then .ListIndex = intCount
        Next
    End With
    Me.cboPageOddEven.ListIndex = 0
End Sub

Private Sub optPageScope_Click(Index As Integer)
    If Me.optPageScope(2).Value = True Then
        Me.txtPageScope.Enabled = True
        Me.txtPageScope.SetFocus
    Else
        Me.txtPageScope.Enabled = False
    End If
End Sub

Private Sub txtCopies_Change()
    If Val(Me.txtCopies.Text) > Me.udCopies.Max Then Me.txtCopies.Text = Me.udCopies.Max
    If Val(Me.txtCopies.Text) < Me.udCopies.Min Then Me.txtCopies.Text = Me.udCopies.Min
End Sub

Private Sub txtCopies_GotFocus()
    With Me.txtCopies
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Call OpenIme(False)
End Sub

Private Sub txtCopies_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab): Exit Sub
    If InStr("1234567890" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtPageScope_GotFocus()
    With Me.txtPageScope
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Call OpenIme(False)
End Sub

Private Sub txtPageScope_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab): Exit Sub
    If InStr("1234567890-," & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
