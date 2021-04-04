VERSION 5.00
Begin VB.Form frmPartogramPrintSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ӡѡ��"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4605
   Icon            =   "frmPartogramPrintSet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��ӡ(&P)"
      Height          =   350
      Left            =   2160
      TabIndex        =   7
      Top             =   1680
      Width           =   1100
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Ԥ��(&V)"
      Height          =   350
      Left            =   2160
      TabIndex        =   6
      Top             =   1680
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3360
      TabIndex        =   8
      Top             =   1680
      Width           =   1100
   End
   Begin VB.Frame fra���� 
      Caption         =   "��ӡ��Χ"
      Height          =   1380
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4380
      Begin VB.TextBox txtPage 
         Height          =   285
         Left            =   2850
         MaxLength       =   1
         TabIndex        =   5
         Top             =   908
         Width           =   465
      End
      Begin VB.TextBox txtFile 
         Height          =   285
         Left            =   1350
         MaxLength       =   1
         TabIndex        =   4
         Top             =   908
         Width           =   495
      End
      Begin VB.OptionButton opt��ǰ 
         Caption         =   "ֻ��ӡ��ǰѡ��Ĳ���ͼ(&2)"
         Height          =   180
         Left            =   480
         TabIndex        =   2
         Top             =   600
         Value           =   -1  'True
         Width           =   2745
      End
      Begin VB.OptionButton optȫ�� 
         Caption         =   "��ӡȫ������ͼ(&1)"
         Height          =   180
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "��ӡ��       ���ļ��ĵ�      ҳ(&3)"
         Height          =   180
         Left            =   480
         TabIndex        =   3
         Top             =   960
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmPartogramPrintSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytOp As Byte
Private mlngFileCount As Long
Private mFrmParent As Object
Private mlngFileIndex As Long
Private mlngFilePage As Long
Private mbytMode As Byte

Public Function PrintSet(ByVal frmParent As Form, ByVal bytMode As Byte, ByVal lngFileCount As Long, lngFileIndex As Long, lngFilePage As Long) As Byte
'--------------------------------------
'���ܣ���ӡǰ����ѯ��,ȷ����ӡģʽ
'���أ�0-ȡ�� 1-Ԥ�� 2-��ӡ
'--------------------------------------
    mbytOp = 0
    mbytMode = bytMode
    mlngFileCount = lngFileCount
    mlngFileIndex = lngFileIndex
    mlngFilePage = lngFilePage
    
    Set mFrmParent = frmParent
    
    Me.Show 1, frmParent
    lngFileIndex = mlngFileIndex
    lngFilePage = mlngFilePage
    
    PrintSet = mbytOp
End Function

Private Sub cmdCancel_Click()
    mbytOp = 0
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If Not GetValue Then Exit Sub
    mbytOp = 2
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    If Not GetValue Then Exit Sub
    mbytOp = 1
    Unload Me
End Sub

Private Sub Form_Load()
    Dim lngIndex As Long
    
    txtFile.Text = 1
    txtPage.Text = 1
    If mbytMode = 1 Then
        cmdPreview.Visible = True
        cmdPreview.Enabled = True
        cmdPrint.Visible = False
        cmdPrint.Enabled = False
    Else
        cmdPreview.Visible = False
        cmdPreview.Enabled = False
        cmdPrint.Visible = True
        cmdPrint.Enabled = True
    End If
    
    lngIndex = Val(GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\����ͼ", "��ӡѡ��", 1))
    Select Case lngIndex
        Case 0
            optȫ��.Value = True
        Case 2
            opt����.Value = True
        Case Else
            opt��ǰ.Value = True
    End Select
End Sub

Private Sub opt����_Click()
    txtFile.Enabled = opt����.Value
    txtPage.Enabled = opt����.Value
End Sub

Private Sub txtFile_GotFocus()
    Call zlControl.TxtSelAll(txtFile)
End Sub

Private Sub txtFile_KeyPress(KeyAscii As Integer)
    If KeyAscii < Asc("1") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPage_GotFocus()
    Call zlControl.TxtSelAll(txtPage)
End Sub

Private Sub txtPage_KeyPress(KeyAscii As Integer)
    If KeyAscii < Asc("1") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
End Sub

Private Function GetValue() As Boolean
    Dim lngPage As Long
    Dim lngIndex As Long
    
    If opt����.Value = True Then
        If Val(txtFile.Text) < 0 Or Val(txtFile.Text) > mlngFileCount Then
            MsgBox "�ļ�����������1��" & mlngFileCount & "֮��!", vbInformation, gstrSysName
            Exit Function
        End If
        
        lngPage = GetFilePage(mFrmParent.FileID, mFrmParent.PatiID, mFrmParent.PageID, Val(txtFile.Text))
        If Val(txtPage.Text) < 0 Or Val(txtPage.Text) > lngPage Then
            MsgBox "�ļ�ҳ��������1��" & lngPage & "֮��!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If optȫ��.Value = True Then
        mlngFileIndex = -1
        mlngFilePage = -1
        lngIndex = 0
    ElseIf opt��ǰ.Value = True Then
        mlngFilePage = -1
        lngIndex = 1
    Else
        mlngFileIndex = Val(txtFile.Text)
        mlngFilePage = Val(txtPage.Text)
        lngIndex = 2
    End If
    
    '�����û�ѡ��
    Call SaveSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\����ͼ", "��ӡѡ��", lngIndex)

    GetValue = True
End Function
