VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPageSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҳ������"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5325
   Icon            =   "frmPageSetup.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   3
      Left            =   930
      TabIndex        =   42
      Top             =   3825
      Width           =   1590
   End
   Begin VB.CommandButton cmdDeskColor 
      Caption         =   "�ĵ�����ɫ(&A)..."
      Height          =   350
      Left            =   465
      TabIndex        =   41
      Top             =   4365
      Width           =   1860
   End
   Begin VB.CommandButton cmdPaperColor 
      Caption         =   "ҳ�汳��ɫ(&G)..."
      Height          =   350
      Left            =   465
      TabIndex        =   40
      Top             =   4020
      Width           =   1860
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2775
      TabIndex        =   29
      Top             =   5025
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3990
      TabIndex        =   30
      Top             =   5025
      Width           =   1100
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Ĭ��ҳ��(&D)..."
      Height          =   350
      Left            =   255
      TabIndex        =   38
      Top             =   5025
      Width           =   1500
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   4
      Left            =   195
      TabIndex        =   37
      Top             =   4890
      Width           =   4890
   End
   Begin VB.PictureBox picViewer 
      BackColor       =   &H00808080&
      Height          =   2160
      Left            =   2775
      ScaleHeight     =   2100
      ScaleWidth      =   2100
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   2550
      Width           =   2160
      Begin VB.PictureBox picPaper 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1830
         Left            =   405
         ScaleHeight     =   1800
         ScaleWidth      =   1335
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   60
         Width           =   1365
         Begin VB.Line linMarjin 
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   3  'Dot
            Index           =   1
            X1              =   0
            X2              =   1410
            Y1              =   1215
            Y2              =   1215
         End
         Begin VB.Line linMarjin 
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   3  'Dot
            Index           =   3
            X1              =   930
            X2              =   930
            Y1              =   0
            Y2              =   1530
         End
         Begin VB.Line linMarjin 
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   3  'Dot
            Index           =   0
            X1              =   0
            X2              =   1410
            Y1              =   105
            Y2              =   105
         End
         Begin VB.Line linMarjin 
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   3  'Dot
            Index           =   2
            X1              =   105
            X2              =   105
            Y1              =   0
            Y2              =   1530
         End
      End
   End
   Begin VB.OptionButton optOrient 
      Caption         =   "����(&S)"
      Height          =   270
      Index           =   1
      Left            =   960
      TabIndex        =   28
      Top             =   3285
      Width           =   1065
   End
   Begin VB.OptionButton optOrient 
      Caption         =   "����(&P)"
      Height          =   270
      Index           =   0
      Left            =   960
      TabIndex        =   27
      Top             =   2790
      Value           =   -1  'True
      Width           =   1065
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   2
      Left            =   945
      TabIndex        =   34
      Top             =   2535
      Width           =   1590
   End
   Begin VB.TextBox txtMarjin 
      Height          =   300
      Index           =   0
      Left            =   1140
      MaxLength       =   6
      TabIndex        =   12
      Text            =   "25.4"
      Top             =   1605
      Width           =   735
   End
   Begin VB.TextBox txtMarjin 
      Height          =   300
      Index           =   1
      Left            =   3495
      MaxLength       =   6
      TabIndex        =   16
      Text            =   "25.4"
      Top             =   1605
      Width           =   735
   End
   Begin VB.TextBox txtMarjin 
      Height          =   300
      Index           =   2
      Left            =   1140
      MaxLength       =   6
      TabIndex        =   20
      Text            =   "31.7"
      Top             =   2010
      Width           =   735
   End
   Begin VB.TextBox txtMarjin 
      Height          =   300
      Index           =   3
      Left            =   3495
      MaxLength       =   6
      TabIndex        =   24
      Text            =   "31.7"
      Top             =   2010
      Width           =   735
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   945
      TabIndex        =   32
      Top             =   1410
      Width           =   4170
   End
   Begin VB.TextBox txtWidth 
      Height          =   300
      Left            =   1140
      MaxLength       =   6
      TabIndex        =   4
      Text            =   "210.05"
      Top             =   885
      Width           =   735
   End
   Begin VB.TextBox txtHeight 
      Height          =   300
      Left            =   3480
      MaxLength       =   6
      TabIndex        =   8
      Text            =   "297.08"
      Top             =   885
      Width           =   705
   End
   Begin VB.ComboBox cboKind 
      Height          =   300
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   420
      Width           =   3285
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   945
      TabIndex        =   0
      Top             =   210
      Width           =   4170
   End
   Begin MSComCtl2.UpDown udHeight 
      Height          =   300
      Left            =   4185
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   885
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      BuddyControl    =   "txtHeight"
      BuddyDispid     =   196621
      OrigLeft        =   4170
      OrigTop         =   900
      OrigRight       =   4410
      OrigBottom      =   1185
      Max             =   765
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown udWidth 
      Height          =   300
      Left            =   1875
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   885
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      BuddyControl    =   "txtWidth"
      BuddyDispid     =   196620
      OrigLeft        =   1830
      OrigTop         =   893
      OrigRight       =   2070
      OrigBottom      =   1178
      Max             =   765
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown udMarjin 
      Height          =   300
      Index           =   0
      Left            =   1875
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1605
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      BuddyControl    =   "txtMarjin(0)"
      BuddyDispid     =   196619
      BuddyIndex      =   0
      OrigLeft        =   1830
      OrigTop         =   1605
      OrigRight       =   2070
      OrigBottom      =   1905
      Max             =   210
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown udMarjin 
      Height          =   300
      Index           =   1
      Left            =   4230
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1605
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      BuddyControl    =   "txtMarjin(1)"
      BuddyDispid     =   196619
      BuddyIndex      =   1
      OrigLeft        =   4185
      OrigTop         =   1605
      OrigRight       =   4425
      OrigBottom      =   1905
      Max             =   210
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown udMarjin 
      Height          =   300
      Index           =   2
      Left            =   1875
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2010
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      BuddyControl    =   "txtMarjin(2)"
      BuddyDispid     =   196619
      BuddyIndex      =   2
      OrigLeft        =   1830
      OrigTop         =   2010
      OrigRight       =   2070
      OrigBottom      =   2310
      Max             =   210
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown udMarjin 
      Height          =   300
      Index           =   3
      Left            =   4230
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2010
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      BuddyControl    =   "txtMarjin(3)"
      BuddyDispid     =   196619
      BuddyIndex      =   3
      OrigLeft        =   4185
      OrigTop         =   2010
      OrigRight       =   4425
      OrigBottom      =   2310
      Max             =   210
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   1770
      Top             =   5055
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label lblColor 
      AutoSize        =   -1  'True
      Caption         =   "ҳ����ɫ"
      Height          =   180
      Left            =   195
      TabIndex        =   43
      Top             =   3750
      Width           =   720
   End
   Begin VB.Label lblKind 
      AutoSize        =   -1  'True
      Caption         =   "�ߴ�(&K)"
      Height          =   180
      Left            =   435
      TabIndex        =   1
      Top             =   480
      Width           =   630
   End
   Begin VB.Image imgOrient 
      Height          =   480
      Index           =   0
      Left            =   435
      Picture         =   "frmPageSetup.frx":000C
      Top             =   2700
      Width           =   480
   End
   Begin VB.Image imgOrient 
      Height          =   480
      Index           =   1
      Left            =   450
      Picture         =   "frmPageSetup.frx":08D6
      Top             =   3165
      Width           =   480
   End
   Begin VB.Label lblOrient 
      AutoSize        =   -1  'True
      Caption         =   "ֽ�ŷ���"
      Height          =   180
      Left            =   195
      TabIndex        =   35
      Top             =   2460
      Width           =   720
   End
   Begin VB.Label lblUnit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Index           =   5
      Left            =   4515
      TabIndex        =   26
      Top             =   2070
      Width           =   360
   End
   Begin VB.Label lblUnit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Index           =   4
      Left            =   2130
      TabIndex        =   22
      Top             =   2070
      Width           =   360
   End
   Begin VB.Label lblUnit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Index           =   3
      Left            =   4515
      TabIndex        =   18
      Top             =   1665
      Width           =   360
   End
   Begin VB.Label lblUnit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Index           =   2
      Left            =   2130
      TabIndex        =   14
      Top             =   1665
      Width           =   360
   End
   Begin VB.Label lblMarjin 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��(&T)"
      Height          =   180
      Index           =   0
      Left            =   615
      TabIndex        =   11
      Top             =   1665
      Width           =   450
   End
   Begin VB.Label lblMarjin 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��(&B)"
      Height          =   180
      Index           =   1
      Left            =   2925
      TabIndex        =   15
      Top             =   1665
      Width           =   450
   End
   Begin VB.Label lblMarjin 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��(&L)"
      Height          =   180
      Index           =   2
      Left            =   615
      TabIndex        =   19
      Top             =   2070
      Width           =   450
   End
   Begin VB.Label lblMarjin 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��(&R)"
      Height          =   180
      Index           =   3
      Left            =   2925
      TabIndex        =   23
      Top             =   2070
      Width           =   450
   End
   Begin VB.Label lblRound 
      AutoSize        =   -1  'True
      Caption         =   "ҳ�߾�"
      Height          =   180
      Left            =   195
      TabIndex        =   33
      Top             =   1335
      Width           =   540
   End
   Begin VB.Label lblUnit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Index           =   1
      Left            =   4515
      TabIndex        =   10
      Top             =   945
      Width           =   360
   End
   Begin VB.Label lblWidth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���(&W)"
      Height          =   180
      Left            =   435
      TabIndex        =   3
      Top             =   945
      Width           =   630
   End
   Begin VB.Label lblHeight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�߶�(&H)"
      Height          =   180
      Left            =   2745
      TabIndex        =   7
      Top             =   945
      Width           =   630
   End
   Begin VB.Label lblUnit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Index           =   0
      Left            =   2130
      TabIndex        =   6
      Top             =   945
      Width           =   360
   End
   Begin VB.Label lblPaper 
      AutoSize        =   -1  'True
      Caption         =   "ֽ������"
      Height          =   180
      Left            =   195
      TabIndex        =   31
      Top             =   135
      Width           =   720
   End
End
Attribute VB_Name = "frmPageSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const conTwipmm As Double = 56.6857142857143     '������羵�ת������

Dim blnOK As Boolean
Dim blnInSelect As Boolean  '����ֽ��ѡ����
Dim intCount As Integer
Dim aryItems() As String

Private Sub cboKind_Click()
    aryItems = Split(Me.cboKind.Text, ",")
    blnInSelect = True
    If Me.cboKind.ListIndex <> Me.cboKind.ListCount - 1 Then
        Me.txtHeight.Text = Format(aryItems(1) / conTwipmm, "0.00")
        Me.txtWidth.Text = Format(aryItems(2) / conTwipmm, "0.00")
    Else
        If Val(Me.txtHeight.Text) > Int(aryItems(1) / conTwipmm * 100) / 100 Then Me.txtHeight.Text = Int(aryItems(1) / conTwipmm * 100) / 100
        If Val(Me.txtWidth.Text) > Int(aryItems(2) / conTwipmm * 100) / 100 Then Me.txtWidth.Text = Int(aryItems(2) / conTwipmm * 100) / 100
    End If
    
    Call RedrawSample
    
    blnInSelect = False
End Sub

Private Sub cboKind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    blnOK = False: Me.Hide
End Sub

Private Sub cmdDefault_Click()
    Dim strMsgInfo As String
    
    If Not ValidSet() Then Exit Sub
    
    strMsgInfo = "�Ƿ񽫵�ǰ���õ�ֽ�����ࡢҳ�߾ࡢֽ�ŷ�����ΪĬ��ҳ�����ñ��棿" & _
        vbCrLf & "�˸��Ľ�Ӱ���µ��ĵ���"
    If MsgBox(strMsgInfo, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Sub
    
    If Me.cboKind.ListIndex = Me.cboKind.ListCount - 1 Then
        SaveSetting UCase(App.ProductName), "PAGE", UCase("PaperKind"), cprPKCustom
    Else
        SaveSetting UCase(App.ProductName), "PAGE", UCase("PaperKind"), Me.cboKind.ListIndex + 1
    End If
    
    If Me.optOrient(0) Then
        SaveSetting UCase(App.ProductName), "PAGE", UCase("PaperOrient"), cprPOPortrait
        If Me.cboKind.ListIndex <> Me.cboKind.ListCount - 1 Then
            aryItems = Split(Me.cboKind.Text, ",")
            SaveSetting UCase(App.ProductName), "PAGE", UCase("PaperHeight"), aryItems(1)
            SaveSetting UCase(App.ProductName), "PAGE", UCase("PaperWidth"), aryItems(2)
        Else
            SaveSetting UCase(App.ProductName), "PAGE", UCase("PaperHeight"), Int(Val(Me.txtHeight.Text) * conTwipmm)
            SaveSetting UCase(App.ProductName), "PAGE", UCase("PaperWidth"), Int(Val(Me.txtWidth.Text) * conTwipmm)
        End If
        SaveSetting UCase(App.ProductName), "PAGE", UCase("MarginTop"), Int(Val(Me.txtMarjin(0).Text) * conTwipmm)
        SaveSetting UCase(App.ProductName), "PAGE", UCase("MarginBottom"), Int(Val(Me.txtMarjin(1).Text) * conTwipmm)
        SaveSetting UCase(App.ProductName), "PAGE", UCase("MarginLeft"), Int(Val(Me.txtMarjin(2).Text) * conTwipmm)
        SaveSetting UCase(App.ProductName), "PAGE", UCase("MarginRight"), Int(Val(Me.txtMarjin(3).Text) * conTwipmm)
    Else
        SaveSetting UCase(App.ProductName), "PAGE", UCase("PaperOrient"), cprPOLandscape
        If Me.cboKind.ListIndex <> Me.cboKind.ListCount - 1 Then
            aryItems = Split(Me.cboKind.Text, ",")
            SaveSetting UCase(App.ProductName), "PAGE", UCase("PaperHeight"), aryItems(2)
            SaveSetting UCase(App.ProductName), "PAGE", UCase("PaperWidth"), aryItems(1)
        Else
            SaveSetting UCase(App.ProductName), "PAGE", UCase("PaperHeight"), Int(Val(Me.txtWidth.Text) * conTwipmm)
            SaveSetting UCase(App.ProductName), "PAGE", UCase("PaperWidth"), Int(Val(Me.txtHeight.Text) * conTwipmm)
        End If
        SaveSetting UCase(App.ProductName), "PAGE", UCase("MarginTop"), Int(Val(Me.txtMarjin(2).Text) * conTwipmm)
        SaveSetting UCase(App.ProductName), "PAGE", UCase("MarginBottom"), Int(Val(Me.txtMarjin(3).Text) * conTwipmm)
        SaveSetting UCase(App.ProductName), "PAGE", UCase("MarginLeft"), Int(Val(Me.txtMarjin(0).Text) * conTwipmm)
        SaveSetting UCase(App.ProductName), "PAGE", UCase("MarginRight"), Int(Val(Me.txtMarjin(1).Text) * conTwipmm)
    End If
End Sub

Private Sub cmdDeskColor_Click()
    With Me.dlgThis
        If Me.picViewer.BackColor <> tomAutoColor Then .Color = Me.picViewer.BackColor
        .DialogTitle = "�ĵ�����ɫ"
        Err = 0: On Error Resume Next
        .ShowColor
        If Err.Number <> 0 Then Exit Sub
        Me.cmdDeskColor.Tag = ""
        Me.picViewer.BackColor = .Color
    End With
End Sub

Private Sub cmdOK_Click()
    If ValidSet() Then blnOK = True: Me.Hide
End Sub

Private Sub cmdPaperColor_Click()
    With Me.dlgThis
        If Me.picPaper.BackColor <> tomAutoColor Then .Color = Me.picPaper.BackColor
        .DialogTitle = "ҳ�汳��ɫ"
        Err = 0: On Error Resume Next
        .ShowColor
        If Err.Number <> 0 Then Exit Sub
        Me.cmdPaperColor.Tag = ""
        Me.picPaper.BackColor = .Color
    End With
End Sub

Private Sub Form_Activate()
    If Me.cmdPaperColor.Visible = False And Me.cmdDeskColor.Visible = False Then
        Me.lblColor.Visible = False: Me.fraLine(3).Visible = False
    End If
End Sub

Private Sub optOrient_Click(Index As Integer)
    Dim strCaption As String
    
    strCaption = Me.lblWidth.Caption
    Me.lblWidth.Caption = Me.lblHeight.Caption
    Me.lblHeight.Caption = strCaption
    
    strCaption = Me.lblMarjin(0).Caption
    Me.lblMarjin(0).Caption = Me.lblMarjin(2).Caption
    Me.lblMarjin(2).Caption = strCaption
    
    strCaption = Me.lblMarjin(1).Caption
    Me.lblMarjin(1).Caption = Me.lblMarjin(3).Caption
    Me.lblMarjin(3).Caption = strCaption
    
    Call RedrawSample

End Sub

Private Sub optOrient_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub txtHeight_Change()
    If blnInSelect Then Exit Sub
    Me.cboKind.ListIndex = Me.cboKind.ListCount - 1
    Call RedrawSample
End Sub

Private Sub txtHeight_GotFocus()
    With Me.txtHeight
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Call OpenIme(False)
End Sub

Private Sub txtHeight_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab): Exit Sub
    If InStr("1234567890." & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtMarjin_Change(Index As Integer)
    Call RedrawMarjin(Index)
End Sub

Private Sub txtMarjin_GotFocus(Index As Integer)
    With Me.txtMarjin(Index)
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Call OpenIme(False)
End Sub

Private Sub txtMarjin_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab): Exit Sub
    If InStr("1234567890." & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtWidth_Change()
    If blnInSelect Then Exit Sub
    Me.cboKind.ListIndex = Me.cboKind.ListCount - 1
    Call RedrawSample
End Sub

Private Sub txtWidth_GotFocus()
    With Me.txtWidth
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Call OpenIme(False)
End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab): Exit Sub
    If InStr("1234567890." & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Public Function ShowMe(Editor As Editor, Optional intFlags As Integer) As Boolean
    '���ܣ���ʾ��ҳ��Ի���
    '������
    '   Editor,��Ҫ����ҳ��ı༭������
    '   intFlags,�Ƿ��ֹ��صĸ���Ч��ѡ�
    '       intFlags and (2^0) <> 0,��ֹ����ҳ�汳��ɫ����
    '       intFlags and (2^1) <> 0,��ֹ�����ĵ�����ɫ����
    
    'װ��ҳ�������б�
    With Me.cboKind
        .Clear
        For intCount = LBound(PaperKindConst) To UBound(PaperKindConst)
            .AddItem PaperKindConst(intCount)
            .itemData(.NewIndex) = Split(PaperKindConst(intCount), ",")(7)
        Next
    End With
    cboKind.ListIndex = SeekCboIndex(cboKind, Editor.PaperKind)
    
    If Editor.PaperOrient = cprPOPortrait Then
        Me.optOrient(0).Value = True
        Me.txtMarjin(0).Text = Round(Editor.MarginTop / conTwipmm, 2)
        Me.txtMarjin(1).Text = Round(Editor.MarginBottom / conTwipmm, 2)
        Me.txtMarjin(2).Text = Round(Editor.MarginLeft / conTwipmm, 2)
        Me.txtMarjin(3).Text = Round(Editor.MarginRight / conTwipmm, 2)
    Else
        Me.optOrient(1).Value = True
        Me.txtMarjin(2).Text = Round(Editor.MarginTop / conTwipmm, 2)
        Me.txtMarjin(3).Text = Round(Editor.MarginBottom / conTwipmm, 2)
        Me.txtMarjin(0).Text = Round(Editor.MarginLeft / conTwipmm, 2)
        Me.txtMarjin(1).Text = Round(Editor.MarginRight / conTwipmm, 2)
    End If
    
    If Me.cboKind.ListIndex = Me.cboKind.ListCount - 1 Then
        Me.txtHeight.Text = Round(Editor.PaperHeight / conTwipmm * 100) / 100
        Me.txtWidth.Text = Round(Editor.PaperWidth / conTwipmm * 100) / 100
    End If
    
    If (intFlags And (2 ^ 0)) <> 0 Then '��ֹҳ����ɫ
        Me.cmdPaperColor.Visible = False
    Else
        If Editor.PaperColor = tomAutoColor Then
            Me.cmdPaperColor.Tag = CStr(tomAutoColor)
        Else
            Me.picPaper.BackColor = Editor.PaperColor
        End If
    End If
    If (intFlags And (2 ^ 1)) <> 0 Then '��ֹ�ĵ�������ɫ
        Me.cmdDeskColor.Visible = False
    Else
        If Editor.BackColor = tomAutoColor Then
            Me.cmdDeskColor.Tag = CStr(tomAutoColor)
        Else
            Me.picViewer.BackColor = Editor.BackColor
        End If
    End If
    
    blnOK = False
    Me.Show 1
    If blnOK = False Then ShowMe = False: Unload Me: Exit Function
    
    If cboKind.itemData(cboKind.ListIndex) = cprPKCustom Then
        Editor.PaperKind = cprPKCustom
    Else
        Editor.PaperKind = cboKind.itemData(cboKind.ListIndex)
    End If
    
    If Me.optOrient(0) Then
        Editor.PaperOrient = cprPOPortrait
        If Me.cboKind.ListIndex <> Me.cboKind.ListCount - 1 Then
            aryItems = Split(Me.cboKind.Text, ",")
            Editor.PaperHeight = aryItems(1)
            Editor.PaperWidth = aryItems(2)
        Else
            Editor.PaperHeight = Int(Val(Me.txtHeight.Text) * conTwipmm)
            Editor.PaperWidth = Int(Val(Me.txtWidth.Text) * conTwipmm)
        End If
        Editor.MarginTop = Int(Val(Me.txtMarjin(0).Text) * conTwipmm)
        Editor.MarginBottom = Int(Val(Me.txtMarjin(1).Text) * conTwipmm)
        Editor.MarginLeft = Int(Val(Me.txtMarjin(2).Text) * conTwipmm)
        Editor.MarginRight = Int(Val(Me.txtMarjin(3).Text) * conTwipmm)
    Else
        Editor.PaperOrient = cprPOLandscape
        If Me.cboKind.ListIndex <> Me.cboKind.ListCount - 1 Then
            aryItems = Split(Me.cboKind.Text, ",")
            Editor.PaperHeight = aryItems(2)
            Editor.PaperWidth = aryItems(1)
        Else
            Editor.PaperHeight = Int(Val(Me.txtWidth.Text) * conTwipmm)
            Editor.PaperWidth = Int(Val(Me.txtHeight.Text) * conTwipmm)
        End If
        Editor.MarginTop = Int(Val(Me.txtMarjin(2).Text) * conTwipmm)
        Editor.MarginBottom = Int(Val(Me.txtMarjin(3).Text) * conTwipmm)
        Editor.MarginLeft = Int(Val(Me.txtMarjin(0).Text) * conTwipmm)
        Editor.MarginRight = Int(Val(Me.txtMarjin(1).Text) * conTwipmm)
    End If
    
    If (intFlags And (2 ^ 0)) = 0 And Me.cmdPaperColor.Tag <> CStr(tomAutoColor) Then Editor.PaperColor = Me.picPaper.BackColor
    If (intFlags And (2 ^ 1)) = 0 And Me.cmdDeskColor.Tag <> CStr(tomAutoColor) Then Editor.BackColor = Me.picViewer.BackColor
    
    ShowMe = True: Unload Me
End Function

Private Function ValidSet() As Boolean
    '���ܣ�������õĺ����ԣ�����ʾ�����Զ�����
    
    Dim dblMarjin As Double
    aryItems = Split(Me.cboKind.Text, ",")
    
    '�Զ���ֽ�ţ���Ҫ����ȸ߶��Ƿ񳬹��߽�
    If Me.cboKind.ListIndex = Me.cboKind.ListCount - 1 Then
        If Val(txtHeight.Text) = 0 Or Val(txtWidth.Text) = 0 Then
            MsgBox "��ָ��ֽ�ſ�Ⱥ͸߶ȣ�", vbInformation, Me.Caption
            Exit Function
        End If
        If (Me.txtHeight.Text) > Int(aryItems(1) / conTwipmm * 100) / 100 Then
            ValidSet = False
            If MsgBox(IIf(Me.optOrient(0).Value = True, "�߶�", "���") & "�����Զ���ֽ�����ơ��Ƿ��Զ�������", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
                Me.txtHeight.Text = Int(aryItems(1) / conTwipmm * 100) / 100
            End If
            Me.udHeight.SetFocus
            Exit Function
        End If
        If Val(Me.txtWidth.Text) > Int(aryItems(2) / conTwipmm * 100) / 100 Then
            ValidSet = False
            If MsgBox(IIf(Me.optOrient(0).Value = True, "���", "�߶�") & "�����Զ���ֽ�����ơ��Ƿ��Զ�������", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
                Me.txtWidth.Text = Int(aryItems(2) / conTwipmm * 100) / 100
            End If
            Me.txtWidth.SetFocus
            Exit Function
        End If
    End If
    
    '�ϱ߾��ж�
    If Int(Val(Me.txtMarjin(0).Text) * conTwipmm) < aryItems(3) Then
        ValidSet = False
        If MsgBox(IIf(Me.optOrient(0).Value = True, "�ϱ߾�", "��߾�") & "���� " & Trim(aryItems(0)) & " ���ơ��Ƿ��Զ�������", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            dblMarjin = aryItems(3) / conTwipmm * 100
            If dblMarjin = Int(dblMarjin) Then
                Me.txtMarjin(0).Text = Int(dblMarjin) / 100
            Else
                Me.txtMarjin(0).Text = Int(dblMarjin) / 100 + 0.01
            End If
        End If
        Me.txtMarjin(0).SetFocus
        Exit Function
    End If
    If Int((Val(Me.txtMarjin(0).Text) + 10) * conTwipmm) > aryItems(1) / 2 Then
        ValidSet = False
        If MsgBox(IIf(Me.optOrient(0).Value = True, "�ϱ߾�", "��߾�") & "���� " & Trim(aryItems(0)) & " ���ơ��Ƿ��Զ�������", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            dblMarjin = (aryItems(1) / 2 / conTwipmm - 10) * 100
            If dblMarjin = Int(dblMarjin) Then
                Me.txtMarjin(0).Text = Int(dblMarjin) / 100
            Else
                Me.txtMarjin(0).Text = Int(dblMarjin) / 100 - 0.01
            End If
        End If
        Me.txtMarjin(0).SetFocus
        Exit Function
    End If
    
    '�±߾��ж�
    If Int(Val(Me.txtMarjin(1).Text) * conTwipmm) < aryItems(4) Then
        ValidSet = False
        If MsgBox(IIf(Me.optOrient(0).Value = True, "�±߾�", "�ұ߾�") & "���� " & Trim(aryItems(0)) & " ���ơ��Ƿ��Զ�������", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            dblMarjin = aryItems(4) / conTwipmm * 100
            If dblMarjin = Int(dblMarjin) Then
                Me.txtMarjin(1).Text = Int(dblMarjin) / 100
            Else
                Me.txtMarjin(1).Text = Int(dblMarjin) / 100 + 0.01
            End If
        End If
        Me.txtMarjin(1).SetFocus
        Exit Function
    End If
    If Int((Val(Me.txtMarjin(1).Text) + 10) * conTwipmm) > aryItems(1) / 2 Then
        ValidSet = False
        If MsgBox(IIf(Me.optOrient(0).Value = True, "�±߾�", "�ұ߾�") & "���� " & Trim(aryItems(0)) & " ���ơ��Ƿ��Զ�������", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            dblMarjin = (aryItems(1) / 2 / conTwipmm - 10) * 100
            If dblMarjin = Int(dblMarjin) Then
                Me.txtMarjin(1).Text = Int(dblMarjin) / 100
            Else
                Me.txtMarjin(1).Text = Int(dblMarjin) / 100 - 0.01
            End If
        End If
        Me.txtMarjin(1).SetFocus
        Exit Function
    End If
    
    '��߾��ж�
    If Int(Val(Me.txtMarjin(2).Text) * conTwipmm) < aryItems(5) Then
        ValidSet = False
        If MsgBox(IIf(Me.optOrient(0).Value = True, "��߾�", "�ϱ߾�") & "���� " & Trim(aryItems(0)) & " ���ơ��Ƿ��Զ�������", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            dblMarjin = aryItems(5) / conTwipmm * 100
            If dblMarjin = Int(dblMarjin) Then
                Me.txtMarjin(2).Text = Int(dblMarjin) / 100
            Else
                Me.txtMarjin(2).Text = Int(dblMarjin) / 100 + 0.01
            End If
        End If
        Me.txtMarjin(2).SetFocus
        Exit Function
    End If
    If Int((Val(Me.txtMarjin(2).Text) + 10) * conTwipmm) > aryItems(2) / 2 Then
        ValidSet = False
        If MsgBox(IIf(Me.optOrient(0).Value = True, "��߾�", "�ϱ߾�") & "���� " & Trim(aryItems(0)) & " ���ơ��Ƿ��Զ�������", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            dblMarjin = (aryItems(2) / 2 / conTwipmm - 10) * 100
            If dblMarjin = Int(dblMarjin) Then
                Me.txtMarjin(2).Text = Int(dblMarjin) / 100
            Else
                Me.txtMarjin(2).Text = Int(dblMarjin) / 100 - 0.01
            End If
        End If
        Me.txtMarjin(2).SetFocus
        Exit Function
    End If
    
    '�ұ߾��ж�
    If Int(Val(Me.txtMarjin(3).Text) * conTwipmm) < aryItems(6) Then
        ValidSet = False
        If MsgBox(IIf(Me.optOrient(0).Value = True, "�ұ߾�", "�±߾�") & "���� " & Trim(aryItems(0)) & " ���ơ��Ƿ��Զ�������", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            dblMarjin = aryItems(6) / conTwipmm * 100
            If dblMarjin = Int(dblMarjin) Then
                Me.txtMarjin(3).Text = Int(dblMarjin) / 100
            Else
                Me.txtMarjin(3).Text = Int(dblMarjin) / 100 + 0.01
            End If
        End If
        Me.txtMarjin(3).SetFocus
        Exit Function
    End If
    If Int((Val(Me.txtMarjin(3).Text) + 10) * conTwipmm) > aryItems(2) / 2 Then
        ValidSet = False
        If MsgBox(IIf(Me.optOrient(0).Value = True, "�ұ߾�", "�±߾�") & "���� " & Trim(aryItems(0)) & " ���ơ��Ƿ��Զ�������", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            dblMarjin = (aryItems(2) / 2 / conTwipmm - 10) * 100
            If dblMarjin = Int(dblMarjin) Then
                Me.txtMarjin(3).Text = Int(dblMarjin) / 100
            Else
                Me.txtMarjin(3).Text = Int(dblMarjin) / 100 - 0.01
            End If
        End If
        Me.txtMarjin(3).SetFocus
        Exit Function
    End If
    
    ValidSet = True
End Function

Private Sub RedrawSample()
    '���ܣ����»���ҳ��ʾ��
    Dim dblWidth As Double, dblHeight As Double
    
    If Val(Trim(txtWidth.Text)) = 0 Then Exit Sub
    If Val(Trim(txtHeight.Text)) = 0 Then Exit Sub
    If Val(Trim(txtMarjin(0).Text)) = 0 Then Exit Sub
    If Val(Trim(txtMarjin(1).Text)) = 0 Then Exit Sub
    If Val(Trim(txtMarjin(2).Text)) = 0 Then Exit Sub
    If Val(Trim(txtMarjin(3).Text)) = 0 Then Exit Sub
    
    If Me.optOrient(0).Value Then
        dblWidth = Val(Me.txtWidth.Text): dblHeight = Val(Me.txtHeight.Text)
    Else
        dblWidth = Val(Me.txtHeight.Text): dblHeight = Val(Me.txtWidth.Text)
    End If
    
    With Me.picPaper
        If dblWidth < dblHeight Then
            .Top = 45: .Height = Me.picViewer.ScaleHeight - 90
            .Width = .Height / dblHeight * dblWidth
            .Left = (Me.picViewer.ScaleWidth - .Width) / 2
        Else
            .Left = 45: .Width = Me.picViewer.ScaleWidth - 90
            .Height = .Width / dblWidth * dblHeight
            .Top = (Me.picViewer.ScaleHeight - .Height) / 2
        End If
    End With
    
    Call RedrawMarjin(0)
    Call RedrawMarjin(1)
    Call RedrawMarjin(2)
    Call RedrawMarjin(3)

End Sub

Private Sub RedrawMarjin(Index As Integer)
    '���ܣ����»���ָ���ı߾�ʾ����
    '������index��0��1��2��3�ֱ�Ϊ�������ұ߾����ã��ڷ���仯ʱ�ͱ߾��߶�Ӧ��ϵ�仯
    
    If Val(Trim(txtWidth.Text)) = 0 Then Exit Sub
    If Val(Trim(txtHeight.Text)) = 0 Then Exit Sub
    If Val(Trim(txtMarjin(0).Text)) = 0 Then Exit Sub
    If Val(Trim(txtMarjin(1).Text)) = 0 Then Exit Sub
    If Val(Trim(txtMarjin(2).Text)) = 0 Then Exit Sub
    If Val(Trim(txtMarjin(3).Text)) = 0 Then Exit Sub
    Select Case Index
    Case 0
        If Me.optOrient(0).Value Then
            With Me.linMarjin(0)
                .X1 = 0: .X2 = Me.picPaper.ScaleWidth - 15
                .Y1 = Val(Me.txtMarjin(0).Text) / Val(Me.txtHeight.Text) * (Me.picPaper.ScaleHeight - 15): .Y2 = .Y1
            End With
        Else
            With Me.linMarjin(2)
                .X1 = Val(Me.txtMarjin(0).Text) / Val(Me.txtHeight.Text) * (Me.picPaper.ScaleWidth - 15): .X2 = .X1
                .Y1 = 0: .Y2 = Me.picPaper.ScaleHeight - 15
            End With
        End If
    Case 1
        If Me.optOrient(0).Value Then
            With Me.linMarjin(1)
                .X1 = 0: .X2 = Me.picPaper.ScaleWidth - 15
                .Y1 = (1 - Val(Me.txtMarjin(1).Text) / Val(Me.txtHeight.Text)) * (Me.picPaper.ScaleHeight - 15): .Y2 = .Y1
            End With
        Else
            With Me.linMarjin(3)
                .X1 = (1 - Val(Me.txtMarjin(1).Text) / Val(Me.txtHeight.Text)) * (Me.picPaper.ScaleWidth - 15): .X2 = .X1
                .Y1 = 0: .Y2 = Me.picPaper.ScaleHeight - 15
            End With
        End If
    Case 2
        If Me.optOrient(0).Value Then
            With Me.linMarjin(2)
                .X1 = Val(Me.txtMarjin(2).Text) / Val(Me.txtWidth.Text) * (Me.picPaper.ScaleWidth - 15): .X2 = .X1
                .Y1 = 0: .Y2 = Me.picPaper.ScaleHeight - 15
            End With
        Else
            With Me.linMarjin(0)
                .X1 = 0: .X2 = Me.picPaper.ScaleWidth - 15
                .Y1 = Val(Me.txtMarjin(2).Text) / Val(Me.txtWidth.Text) * (Me.picPaper.ScaleHeight - 15): .Y2 = .Y1
            End With
        End If
    Case 3
        If Me.optOrient(0).Value Then
            With Me.linMarjin(3)
                .X1 = (1 - Val(Me.txtMarjin(3).Text) / Val(Me.txtWidth.Text)) * (Me.picPaper.ScaleWidth - 15): .X2 = .X1
                .Y1 = 0: .Y2 = Me.picPaper.ScaleHeight - 15
            End With
        Else
            With Me.linMarjin(1)
                .X1 = 0: .X2 = Me.picPaper.ScaleWidth - 15
                .Y1 = (1 - Val(Me.txtMarjin(3).Text) / Val(Me.txtWidth.Text)) * (Me.picPaper.ScaleHeight - 15): .Y2 = .Y1
            End With
        End If
    End Select
End Sub

