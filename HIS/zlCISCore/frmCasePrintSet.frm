VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCasePrintSet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������ӡѡ��"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   Icon            =   "frmCasePrintSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��ӡ(&P)"
      Height          =   350
      Left            =   5400
      TabIndex        =   9
      Top             =   1380
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5400
      TabIndex        =   10
      Top             =   1845
      Width           =   1100
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Ԥ��(&V)"
      Height          =   350
      Left            =   5400
      TabIndex        =   8
      Top             =   930
      Width           =   1100
   End
   Begin VB.Frame fra��ӡ 
      Caption         =   "��ӡ"
      Height          =   2280
      Left            =   120
      TabIndex        =   12
      Top             =   795
      Width           =   5100
      Begin VB.TextBox TxtEnd 
         Height          =   300
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "1"
         Top             =   1860
         Width           =   360
      End
      Begin VB.TextBox TxtBegin 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "1"
         Top             =   1860
         Width           =   360
      End
      Begin VB.CheckBox ChkPrintPage 
         Caption         =   "��ӡָ����Χ"
         Height          =   225
         Left            =   270
         TabIndex        =   20
         Top             =   1560
         Width           =   1635
      End
      Begin VB.PictureBox picBack 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2130
         Left            =   2910
         ScaleHeight     =   491.128
         ScaleMode       =   0  'User
         ScaleWidth      =   491.128
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   120
         Width           =   2130
         Begin VB.PictureBox picPaper 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            ForeColor       =   &H80000008&
            Height          =   1485
            Left            =   405
            ScaleHeight     =   1455
            ScaleMode       =   0  'User
            ScaleWidth      =   1140
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "�϶���ɫ�����ı���ʼλ��"
            Top             =   270
            Width           =   1170
            Begin VB.PictureBox pic��ʼ 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   15
               Left            =   0
               MousePointer    =   7  'Size N S
               ScaleHeight     =   15
               ScaleMode       =   0  'User
               ScaleWidth      =   1140
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   135
               Width           =   1140
            End
         End
         Begin VB.PictureBox picShadow 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1485
            Left            =   450
            ScaleHeight     =   1485
            ScaleWidth      =   1170
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   315
            Width           =   1170
         End
      End
      Begin VB.TextBox txt��ʼ 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1050
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "25"
         Top             =   555
         Width           =   600
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "����ʼλ�ô�ӡ��ϸ������Ϣ"
         Height          =   195
         Left            =   285
         TabIndex        =   2
         Top             =   300
         Value           =   1  'Checked
         Width           =   2640
      End
      Begin MSComCtl2.UpDown UDҳ�� 
         Height          =   300
         Left            =   1665
         TabIndex        =   7
         Top             =   1185
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtҳ��"
         BuddyDispid     =   196623
         OrigLeft        =   1590
         OrigTop         =   1365
         OrigRight       =   1830
         OrigBottom      =   1665
         Max             =   999
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtҳ�� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1050
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "1"
         Top             =   1185
         Width           =   570
      End
      Begin VB.CheckBox chkҳ�� 
         Alignment       =   1  'Right Justify
         Caption         =   "��ӡҳ��"
         Height          =   195
         Left            =   225
         TabIndex        =   5
         Top             =   915
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin MSComCtl2.UpDown UD��ʼ 
         Height          =   300
         Left            =   1665
         TabIndex        =   4
         Top             =   555
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   25
         BuddyControl    =   "txt��ʼ"
         BuddyDispid     =   196620
         OrigLeft        =   1590
         OrigTop         =   705
         OrigRight       =   1830
         OrigBottom      =   1005
         Max             =   460
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Left            =   1200
         TabIndex        =   23
         Top             =   1860
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "TxtBegin"
         BuddyDispid     =   196614
         OrigLeft        =   1590
         OrigTop         =   1365
         OrigRight       =   1830
         OrigBottom      =   1665
         Max             =   999
         Min             =   1
         Enabled         =   0   'False
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   300
         Left            =   2550
         TabIndex        =   26
         Top             =   1860
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "TxtEnd"
         BuddyDispid     =   196613
         OrigLeft        =   1590
         OrigTop         =   1365
         OrigRight       =   1830
         OrigBottom      =   1665
         Max             =   999
         Min             =   1
         Enabled         =   0   'False
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "����ҳ"
         Height          =   180
         Left            =   1590
         TabIndex        =   24
         Top             =   1920
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "��ʼҳ"
         Height          =   180
         Left            =   270
         TabIndex        =   21
         Top             =   1920
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         Height          =   180
         Left            =   1965
         TabIndex        =   19
         Top             =   585
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼλ��"
         Height          =   180
         Left            =   255
         TabIndex        =   14
         Top             =   615
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼҳ��"
         Height          =   180
         Left            =   255
         TabIndex        =   13
         Top             =   1260
         Width           =   720
      End
   End
   Begin VB.Frame fra���� 
      Caption         =   "����"
      Height          =   645
      Left            =   120
      TabIndex        =   11
      Top             =   75
      Width           =   5100
      Begin VB.OptionButton opt���� 
         Caption         =   "�ӵ�ǰ������ʼ������ӡ"
         Height          =   180
         Left            =   2400
         TabIndex        =   1
         Top             =   285
         Width           =   2280
      End
      Begin VB.OptionButton opt��ǰ 
         Caption         =   "ֻ��ӡ��ǰѡ��Ĳ���"
         Height          =   180
         Left            =   225
         TabIndex        =   0
         Top             =   285
         Value           =   -1  'True
         Width           =   2100
      End
   End
End
Attribute VB_Name = "frmCasePrintSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytOpt As Byte

Private mblnFirst As Boolean
Private mblnCurCase As Boolean
Private mblnPatiInfo As Boolean
Private mlngBeginY As Long
Private mintBeginPage As Integer
Private mlng������¼ID As Long
Private mlngPatientID As Long               '����ID
Private mlngPageID As Integer               '������ҳID

Private mlngPrintBeginPage As Long
Private mlngPrintEndPage As Long

Private mlngWidth As Long '�Զ���ֽ�ſ��,Twip
Private mlngHeight As Long '�Զ���ֽ�Ÿ߶�'Twip
Private mlngLeft As Long '��߾�'mm
Private mlngRight As Long '�ұ߾�'mm
Private mlngTop As Long '�ϱ߾�'mm
Private mlngBottom As Long '�±߾�'mm

Private Sub ChkPrintPage_Click()
    If ChkPrintPage.Value = 0 Then
        Me.UpDown1.Enabled = False
        Me.UpDown2.Enabled = False
        mlngPrintBeginPage = 0
        mlngPrintEndPage = 0
        Me.TxtBegin.Locked = True
        Me.TxtEnd.Locked = True
    Else
        Me.UpDown1.Enabled = True
        Me.UpDown2.Enabled = True
        mlngPrintBeginPage = Me.TxtBegin
        mlngPrintEndPage = Me.TxtEnd
        Me.TxtBegin.Locked = False
        Me.TxtEnd.Locked = False
    End If
End Sub

Private Sub chkҳ��_Click()
    txtҳ��.Enabled = chkҳ��.Value = 1
    UDҳ��.Enabled = chkҳ��.Value = 1
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    If Not GetValue Then Exit Sub
    mbytOpt = 1
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If Not GetValue Then Exit Sub
    mbytOpt = 2
    Unload Me
End Sub

Private Sub Form_Load()
Dim rsTmp As New ADODB.Recordset
Dim strSQL As String

    mbytOpt = 0
    
    '��ʾֽ�Ŵ�ӡλ�õ���ͼ
    mlngWidth = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "���", Printer.Width)
    mlngHeight = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�߶�", Printer.Height)
    mlngLeft = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "��߾�", OFFSET_LEFT)
    mlngRight = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�ұ߾�", OFFSET_RIGHT)
    mlngTop = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�ϱ߾�", OFFSET_TOP)
    mlngBottom = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "�±߾�", OFFSET_BOTTOM)
    
    If mlngWidth > mlngHeight Then
        picBack.ScaleWidth = mlngWidth / 56.7 * 1.1
        picBack.ScaleHeight = mlngWidth / 56.7 * 1.1
    Else
        picBack.ScaleWidth = mlngHeight / 56.7 * 1.1
        picBack.ScaleHeight = mlngHeight / 56.7 * 1.1
    End If
    picPaper.Width = mlngWidth / 56.7
    picPaper.Height = mlngHeight / 56.7
    picPaper.Left = (picBack.ScaleWidth - picPaper.Width) / 2
    picPaper.Top = (picBack.ScaleHeight - picPaper.Height) / 2
    picShadow.Width = picPaper.Width
    picShadow.Height = picPaper.Height
    picShadow.Left = picPaper.Left + 5
    picShadow.Top = picPaper.Top + 5
    
    picPaper.ScaleWidth = mlngWidth / 56.7
    picPaper.ScaleHeight = mlngHeight / 56.7
    
    '��ʹ����ӡλ��
    InitPrintPosition
    
    '�����ؼ�ֵ��ʼ
    If mblnCurCase Then
        opt��ǰ.Value = True
    Else
        opt����.Value = True
    End If
    chk����.Value = IIf(mblnPatiInfo, 1, 0)
    
    chkҳ��.Value = IIf(mintBeginPage <> 0, 1, 0)
    UDҳ��.Value = IIf(mintBeginPage = 0, 1, mintBeginPage)
    
    If Not mblnFirst Then
        opt��ǰ.Enabled = False
        opt����.Enabled = False
        
        cmdPrint.Visible = False
        cmdCancel.Top = cmdPrint.Top
        cmdPreview.Caption = "ȷ��(&O)"
        cmdPreview.Default = True
    End If
    mlngPrintBeginPage = 0
    mlngPrintEndPage = 0
End Sub

Public Function PrintSet(objParent As Object, ByVal blnFirst As Boolean, ByRef blnCurCase As Boolean, _
    ByRef blnPatiInfo As Boolean, ByRef lngBeginY As Long, ByRef intBeginPage As Integer, Optional ByVal lng������¼ID As Long = 0, _
    Optional ByRef lng��ʼҳ As Long = 0, Optional ByRef lng����ҳ As Long = 0, Optional ByRef lngPatientID As Long, Optional ByRef lngPageID As Long) As Byte
    '���ܣ����ô�ӡѡ��
    '������blnFirst=�Ƿ��һ�ε���,����ֻ��"ȷ��","ȡ��",�Ҳ������޸Ĳ�����ӡ����
    '      blnCurCase=T=ֻ��ӡ��ǰ����,F=�ӵ�ǰ������ʼ������ӡ����
    '      blnPatiInfo=����ǰ��ӡ������Ϣ
    '      lngBeginY=���β�����ʼ��ӡλ��'mm
    '      intBeginPage=��ʼҳ��,Ϊ0��ʾ����ӡҳ��
    '      lng������¼ID = ͨ�����ID���Դ����ݿ��ж����һ�ε���������Ĵ�ӡλ��,�Ա��ڴ�ӡʱ�����ϴδ�ӡ
    '      lngPatientID����ID
    '      lngPageID������ҳID
    '���أ�0-ȡ��,1-Ԥ��,2-��ӡ
    
    mblnFirst = blnFirst
    mblnCurCase = blnCurCase
    mblnPatiInfo = blnPatiInfo
    mlngBeginY = lngBeginY
    mintBeginPage = intBeginPage
    mlngPrintBeginPage = lng��ʼҳ
    mlngPrintEndPage = lng����ҳ
    mlng������¼ID = lng������¼ID
    mlngPatientID = lngPatientID
    mlngPageID = lngPageID
    Me.Show 1, objParent
    
    blnCurCase = mblnCurCase
    blnPatiInfo = mblnPatiInfo
    lngBeginY = mlngBeginY
    intBeginPage = mintBeginPage
    lng��ʼҳ = mlngPrintBeginPage
    lng����ҳ = mlngPrintEndPage
    PrintSet = mbytOpt
End Function

Private Sub opt��ǰ_Click()
    InitPrintPosition
End Sub

Private Sub opt����_Click()
    InitPrintPosition
End Sub

Private Sub pic��ʼ_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If pic��ʼ.Top + y > UD��ʼ.Max Or pic��ʼ.Top + y < UD��ʼ.Min Then Exit Sub
        pic��ʼ.Top = pic��ʼ.Top + y
        UD��ʼ.Value = pic��ʼ.Top
        Call DrawPage
        Me.Refresh
    End If
End Sub

Private Sub TxtBegin_GotFocus()
    Me.TxtBegin.SelStart = 0
    Me.TxtBegin.SelLength = Len(Me.TxtBegin)
End Sub

Private Sub TxtBegin_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub TxtEnd_GotFocus()
    Me.TxtEnd.SelStart = 0
    Me.TxtEnd.SelLength = Len(Me.TxtEnd.Text)
End Sub

Private Sub TxtEnd_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt��ʼ_Change()
    If Val(txt��ʼ.Text) >= UD��ʼ.Min And Val(txt��ʼ.Text) <= UD��ʼ.Max Then
        UD��ʼ.Value = Val(txt��ʼ.Text)
    End If
End Sub

Private Sub txt��ʼ_GotFocus()
    zlControl.TxtSelAll txt��ʼ
End Sub

Private Sub txt��ʼ_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtҳ��_GotFocus()
    zlControl.TxtSelAll txtҳ��
End Sub

Private Sub txtҳ��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Function GetValue() As Boolean
    If Not (Val(txt��ʼ.Text) >= UD��ʼ.Min And Val(txt��ʼ.Text) <= UD��ʼ.Max) Then
        MsgBox "��ʼλ��Ӧ���� " & UD��ʼ.Min & " �� " & UD��ʼ.Max & " ֮�䣡", vbInformation, gstrSysName
        txt��ʼ.SetFocus: Exit Function
    End If
    If Len(Trim(TxtBegin.Text)) < 1 Then
        TxtBegin.Text = 1
    End If
    If Len(Trim(TxtEnd.Text)) < 1 Then
        TxtEnd.Text = 1
    End If
    mlngPrintBeginPage = Me.TxtBegin
    mlngPrintEndPage = Me.TxtEnd
    If Val(TxtBegin.Text) > Val(TxtEnd.Text) Then
        MsgBox "��ʼҳӦ����1--" & Val(TxtEnd.Text) & "֮�䣡", vbInformation, gstrSysName
        TxtBegin.SetFocus: Exit Function
    End If
    mblnCurCase = opt��ǰ.Value
    mblnPatiInfo = chk����.Value = 1
    mlngBeginY = Val(txt��ʼ.Text)
    If chkҳ��.Value = 1 Then
        mintBeginPage = Val(txtҳ��.Text)
    Else
        mintBeginPage = 0
    End If
    
    GetValue = True
End Function

Private Sub UD��ʼ_Change()
    pic��ʼ.Top = UD��ʼ.Value
    Call DrawPage
End Sub

Private Sub DrawPage()
    picPaper.Cls
    picPaper.Line (0, mlngTop)-(picPaper.ScaleWidth, mlngTop), &H808080
    picPaper.Line (0, picPaper.ScaleHeight - mlngBottom)-(picPaper.ScaleWidth, picPaper.ScaleHeight - mlngBottom), &H808080
    picPaper.Line (mlngLeft, 0)-(mlngLeft, picPaper.ScaleHeight), &H808080
    picPaper.Line (picPaper.ScaleWidth - mlngRight, 0)-(picPaper.ScaleWidth - mlngRight, picPaper.ScaleHeight), &H808080
    
    picPaper.Line (mlngLeft, UD��ʼ.Value)-(picPaper.ScaleWidth - mlngRight, picPaper.ScaleHeight - mlngBottom), &H808080, B
End Sub

Private Sub UpDown1_DownClick()
    If Val(Me.TxtBegin) > 1 Then
        Me.TxtBegin = Val(Me.TxtBegin) - 1
    End If
    mlngPrintBeginPage = Me.TxtBegin
End Sub

Private Sub UpDown1_UpClick()
    Me.TxtBegin = Val(Me.TxtBegin) + 1
    mlngPrintBeginPage = Me.TxtBegin
End Sub
Private Sub UpDown2_DownClick()
    Me.TxtEnd = Val(Me.TxtEnd) - 1
    mlngPrintEndPage = Me.TxtEnd
End Sub

Private Sub UpDown2_UpClick()
    Me.TxtEnd = Val(Me.TxtEnd) + 1
    mlngPrintEndPage = Me.TxtEnd
End Sub
Sub InitPrintPosition()
    '����:          ��ʹ����ӡλ��
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    '���´������ж������һ�ε�ҳ����Y����λ��
    mlngBeginY = 0
    If mlng������¼ID > 0 Then
        strSQL = _
            "SELECT nvl(��ʼҳ��,1) ҳ��, nvl(��ʼλ��,0) Y" & vbCrLf & _
            "  FROM ������ӡ��¼" & vbCrLf & _
            " WHERE ������¼ID = " & mlng������¼ID
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        If rsTmp.RecordCount > 0 Then
            mintBeginPage = rsTmp!ҳ��
            mlngBeginY = Me.picPaper.ScaleY(rsTmp!y, vbTwips, vbMillimeters)
        Else
            If opt����.Value = True Then
                strSQL = "select  nvl(��ʼҳ��,1) ҳ��, nvl(����λ��,0) Y from ���˲�����¼ a , ������ӡ��¼ b where " & _
                " a.id = b.������¼ID " & " and a.����id = " & mlngPatientID & " and a.��ҳID = " & mlngPageID & " order by  ��ӡʱ�� desc,����ҳ�� desc "
                Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
                If rsTmp.EOF <> True Then
                    mintBeginPage = rsTmp!ҳ��
                    mlngBeginY = Me.picPaper.ScaleY(rsTmp!y, vbTwips, vbMillimeters)
                End If
            End If
        End If
    End If
    
    '�Գ�ʼλ��
    If Not (mlngBeginY >= mlngTop And mlngBeginY <= picPaper.ScaleHeight - mlngBottom * 2) Then
        mlngBeginY = mlngTop
    End If
    pic��ʼ.Left = 0
    pic��ʼ.Width = picPaper.ScaleWidth
    pic��ʼ.Top = mlngBeginY
    
    UD��ʼ.Min = mlngTop
    UD��ʼ.Max = picPaper.ScaleHeight - 2 * mlngBottom
    UD��ʼ.Value = mlngBeginY
     
    pic��ʼ.ScaleHeight = 1 '��Ȼ�����϶�
    txt��ʼ.Text = mlngBeginY
    Call DrawPage
End Sub
