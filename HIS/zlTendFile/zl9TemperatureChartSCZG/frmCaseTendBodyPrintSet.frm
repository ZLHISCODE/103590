VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCaseTendBodyPrintSet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ӡѡ��"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "frmCaseTendBodyPrintSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fra 
      Caption         =   "����"
      Height          =   1320
      Left            =   120
      TabIndex        =   12
      Top             =   2625
      Width           =   4380
      Begin VB.CheckBox chk 
         Caption         =   "����ӡ���µ��·�������˵����Ϣ(&9)"
         Height          =   195
         Index           =   1
         Left            =   915
         TabIndex        =   16
         Top             =   1035
         Value           =   1  'Checked
         Width           =   3360
      End
      Begin VB.CheckBox chk 
         Caption         =   "����ӡ���ʺ�����������ߺ���Ӱ(&8)"
         Height          =   195
         Index           =   0
         Left            =   915
         TabIndex        =   15
         Top             =   720
         Value           =   1  'Checked
         Width           =   3360
      End
      Begin VB.TextBox txt 
         Height          =   300
         Left            =   960
         TabIndex        =   14
         Top             =   255
         Width           =   3210
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "�ʿغ�(&7)"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   300
         Width           =   810
      End
   End
   Begin VB.Frame fra��ӡ 
      Caption         =   "��ӡҳ��"
      Height          =   1080
      Left            =   120
      TabIndex        =   6
      Top             =   1500
      Width           =   4380
      Begin VB.CheckBox chkOper 
         Caption         =   "��ӡ��ӡ��(&6)"
         Height          =   195
         Left            =   2625
         TabIndex        =   11
         Top             =   765
         Value           =   1  'Checked
         Width           =   1650
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "��ӡסԺ����(&5)"
         Height          =   195
         Left            =   525
         TabIndex        =   10
         Top             =   765
         Value           =   1  'Checked
         Width           =   1650
      End
      Begin VB.CheckBox chkҳ�� 
         Caption         =   "��ӡҳ�ţ���һҳҳ�ű�ʾΪ(&4)"
         Height          =   195
         Left            =   525
         TabIndex        =   7
         Top             =   405
         Value           =   1  'Checked
         Width           =   2910
      End
      Begin VB.TextBox txt��ʼ 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1050
         MaxLength       =   3
         TabIndex        =   24
         Text            =   "25"
         Top             =   1680
         Visible         =   0   'False
         Width           =   600
      End
      Begin MSComCtl2.UpDown UDҳ�� 
         Height          =   300
         Left            =   3795
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   345
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtҳ��"
         BuddyDispid     =   196618
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
      Begin MSComCtl2.UpDown UD��ʼ 
         Height          =   300
         Left            =   1665
         TabIndex        =   25
         Top             =   1680
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   25
         BuddyControl    =   "txt��ʼ"
         BuddyDispid     =   196617
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
      Begin VB.TextBox txtҳ�� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3435
         MaxLength       =   3
         TabIndex        =   8
         Text            =   "1"
         Top             =   360
         Width           =   360
      End
      Begin VB.PictureBox picBack 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2130
         Left            =   2850
         ScaleHeight     =   491.128
         ScaleMode       =   0  'User
         ScaleWidth      =   491.128
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2175
         Visible         =   0   'False
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
            TabIndex        =   21
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
               TabIndex        =   22
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
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   315
            Width           =   1170
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼλ��"
         Height          =   180
         Left            =   255
         TabIndex        =   27
         Top             =   1740
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         Height          =   180
         Left            =   1965
         TabIndex        =   26
         Top             =   1710
         Visible         =   0   'False
         Width           =   180
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��ӡ(&P)"
      Height          =   350
      Left            =   4620
      TabIndex        =   17
      Top             =   165
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4620
      TabIndex        =   18
      Top             =   570
      Width           =   1100
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Ԥ��(&V)"
      Height          =   350
      Left            =   4620
      TabIndex        =   19
      Top             =   165
      Width           =   1100
   End
   Begin VB.Frame fra���� 
      Caption         =   "��ӡ��Χ"
      Height          =   1380
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   4380
      Begin VB.TextBox txtEnd 
         Height          =   285
         Left            =   2380
         MaxLength       =   4
         TabIndex        =   5
         Top             =   908
         Width           =   465
      End
      Begin VB.TextBox txtBegin 
         Height          =   285
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   4
         Top             =   908
         Width           =   495
      End
      Begin VB.OptionButton optȫ�� 
         Caption         =   "��ӡȫ�����µ�(&3)"
         Height          =   180
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
      Begin VB.OptionButton opt��ǰ 
         Caption         =   "ֻ��ӡ��ǰѡ������±�(&1)"
         Height          =   180
         Left            =   480
         TabIndex        =   2
         Top             =   600
         Value           =   -1  'True
         Width           =   2745
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "�ɵ�"
         Height          =   180
         Left            =   480
         TabIndex        =   3
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ҳ����      ҳ(&2)"
         Height          =   195
         Left            =   1780
         TabIndex        =   28
         Top             =   960
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmCaseTendBodyPrintSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytOpt As Byte

Private mblnFirst As Boolean
Private mlng�ļ�ID As Long
Private mlngAllPage As Long
Private mintPrintRange As Integer
Private mstrPage As String 'ѡ��������ӡʱ��¼��ʼҲ�ͽ���ҳ��
Private mlngBeginY As Long
Private mintBeginPage As Integer
Private mlngWidth As Long '�Զ���ֽ�ſ��,Twip
Private mlngHeight As Long '�Զ���ֽ�Ÿ߶�'Twip
Private mlngLeft As Long '��߾�'mm
Private mlngRight As Long '�ұ߾�'mm
Private mlngTop As Long '�ϱ߾�'mm
Private mlngBottom As Long '�±߾�'mm
Private mblnInit As Boolean

Private mstrPrivs As String

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
    Call zlDatabase.SetPara("�ʿغ�", txt.Text, glngSys, 1255, InStr(mstrPrivs, ";����ѡ������;") > 0)
    If Not GetValue Then Exit Sub
    mbytOpt = 2
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 39 Then KeyCode = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    mbytOpt = 0
    
    '��ʾֽ�Ŵ�ӡλ�õ���ͼ
        
    mlngWidth = Val(zlDatabase.GetPara("���µ����", glngSys, 1255, Printer.Width))
    mlngHeight = Val(zlDatabase.GetPara("���µ��߶�", glngSys, 1255, Printer.Height))
    mlngLeft = Val(zlDatabase.GetPara("���µ���߾�", glngSys, 1255, OFFSET_LEFT))
    mlngRight = Val(zlDatabase.GetPara("���µ��ұ߾�", glngSys, 1255, OFFSET_RIGHT))
    mlngTop = Val(zlDatabase.GetPara("���µ��ϱ߾�", glngSys, 1255, OFFSET_TOP))
    mlngBottom = Val(zlDatabase.GetPara("���µ��±߾�", glngSys, 1255, OFFSET_BOTTOM))
    
    txt.Text = zlDatabase.GetPara("�ʿغ�", glngSys, 1255, "", Array(txt), InStr(mstrPrivs, "����ѡ������") > 0)
    
    If mlngWidth > mlngHeight Then
        picBack.ScaleWidth = mlngWidth / conRatemmToTwip * 1.1
        picBack.ScaleHeight = mlngWidth / conRatemmToTwip * 1.1
    Else
        picBack.ScaleWidth = mlngHeight / conRatemmToTwip * 1.1
        picBack.ScaleHeight = mlngHeight / conRatemmToTwip * 1.1
    End If
    picPaper.Width = mlngWidth / conRatemmToTwip
    picPaper.Height = mlngHeight / conRatemmToTwip
    picPaper.Left = (picBack.ScaleWidth - picPaper.Width) / 2
    picPaper.Top = (picBack.ScaleHeight - picPaper.Height) / 2
    picShadow.Width = picPaper.Width
    picShadow.Height = picPaper.Height
    picShadow.Left = picPaper.Left + 5
    picShadow.Top = picPaper.Top + 5
    
    picPaper.ScaleWidth = mlngWidth / conRatemmToTwip
    picPaper.ScaleHeight = mlngHeight / conRatemmToTwip
    
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
    
    Call DrawPage
    
    mintPrintRange = Val(zlDatabase.GetPara("������ӡ", glngSys, 1255, "1", Array(opt��ǰ, opt����, optȫ��)))
    Select Case mintPrintRange
    Case 0
        opt��ǰ.Value = True
    Case 1
        opt����.Value = True
    Case 2
        optȫ��.Value = True
    End Select
    
    chkҳ��.Value = Val(zlDatabase.GetPara("��ӡҳ��", glngSys, 1255, "1", Array(chkҳ��)))
    txtҳ��.Text = Val(zlDatabase.GetPara("��ʼҳ��", glngSys, 1255, "1", Array(txtҳ��, UDҳ��)))
    chk����.Value = Val(zlDatabase.GetPara("��ӡ����", glngSys, 1255, "0", Array(chk����)))
    '67405:������,2013-11-25
    chkOper.Value = Val(zlDatabase.GetPara("��ӡ��ӡ��", glngSys, 1255, "0", Array(chkOper)))
    chk(0).Value = Val(zlDatabase.GetPara("����ӡ�������ͼ��", glngSys, 1255, "0", Array(chk(0))))
    chk(1).Value = Val(zlDatabase.GetPara("���µ�����ӡ����˵��", glngSys, 1255, "0", Array(chk(1)), InStr(mstrPrivs, ";����ѡ������;") > 0))
    
    mintBeginPage = Val(txtҳ��.Text)
    
    UDҳ��.Value = IIf(mintBeginPage = 0, 1, mintBeginPage)

End Sub

Public Function PrintSet(objParent As Object, ByVal blnFirst As Boolean, ByVal strParam As String, ByRef intPrintRange As Integer, ByRef lngBeginY As Long, ByRef intBeginPage As Integer, strPage As String, ByVal strPrivs As String, ByVal bytMode As Byte) As Byte
'���ܣ����ô�ӡѡ��
'������blnFirst=�Ƿ��һ�ε���,����ֻ��"ȷ��","ȡ��",�Ҳ������޸Ĳ�����ӡ����
'      strParam �ɵ�ǰҳ������ӡ�� ��Ҫ��ȡ �ļ�ID;�������µ���ҳ��
'      blnCurCase=T=ֻ��ӡ��ǰ����,F=�ӵ�ǰ������ʼ������ӡ����
'      lngBeginY=���β�����ʼ��ӡλ��'mm
'      intBeginPage=��ʼҳ��,Ϊ0��ʾ����ӡҳ��
'      strPage
'���أ�0-ȡ��,1-Ԥ��,2-��ӡ
    
    mstrPrivs = strPrivs
    mblnFirst = blnFirst
    
    If strParam <> "" Then
        If InStr(1, strParam, ";") = 0 Then
            mlng�ļ�ID = Val(strParam)
        Else
            mlng�ļ�ID = Val(Split(strParam, ";")(0))
            mlngAllPage = Val(Split(strParam, ";")(1))
        End If
    End If
    mintPrintRange = intPrintRange
    mlngBeginY = lngBeginY
    mintBeginPage = intBeginPage
    mblnInit = True
    cmdPrint.Visible = (bytMode = 1)
    cmdPreview.Visible = (bytMode = 2)
    
    Call GetPageNum(mlng�ļ�ID)
    mblnInit = False
    If Not mblnFirst Then
        opt��ǰ.Enabled = False
        opt����.Enabled = False
        txtBegin.Enabled = False
        txtEnd.Enabled = False
        cmdPrint.Visible = False
        cmdCancel.Top = cmdPrint.Top
        cmdPreview.Caption = "ȷ��(&O)"
        cmdPreview.Default = True
    End If
    Me.Show 1, objParent
    
    intPrintRange = mintPrintRange
    lngBeginY = mlngBeginY
    intBeginPage = mintBeginPage
    strPage = mstrPage
    PrintSet = mbytOpt
End Function

Public Function GetPageNum(ByVal lng�ļ�ID As Long) As Boolean
'------------------------------------------------
'��ȡ��ӡҳ��
'------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo Errhand
    strSQL = "select ��ʼҳ��,����ҳ�� From ���˻����ӡ where �ļ�ID=[1] and Rownum<2 Order by ����ʱ�� DESC"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��ӡ����", lng�ļ�ID)
    If rsTemp.RecordCount > 0 Then
        txtBegin.Text = Val(zlCommFun.Nvl(rsTemp!����ҳ��, 1)) + 1
    Else
        txtBegin.Text = 1
        txtEnd.Text = mlngAllPage
    End If
    If Val(txtBegin.Text) < 0 Then txtBegin.Text = 1
    txtEnd.Text = mlngAllPage
    If Val(txtEnd.Text) < 0 Then txtEnd.Text = 1
    If Val(txtEnd.Text) < Val(txtBegin.Text) Then txtBegin.Text = txtEnd.Text
    
    GetPageNum = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub Form_Unload(Cancel As Integer)
    
    If opt��ǰ.Value Then
        Call zlDatabase.SetPara("������ӡ", "0", glngSys, 1255)
    ElseIf opt����.Value Then
        Call zlDatabase.SetPara("������ӡ", "1", glngSys, 1255)
    Else
        Call zlDatabase.SetPara("������ӡ", "2", glngSys, 1255)
    End If
    
    Call zlDatabase.SetPara("��ӡҳ��", chkҳ��.Value, glngSys, 1255)
    Call zlDatabase.SetPara("��ʼҳ��", Val(txtҳ��.Text), glngSys, 1255)
    Call zlDatabase.SetPara("��ӡ����", chk����.Value, glngSys, 1255)
    '67405:������,2013-11-25,���"��ӡ��ӡ��"
    Call zlDatabase.SetPara("��ӡ��ӡ��", chkOper.Value, glngSys, 1255)
    Call zlDatabase.SetPara("����ӡ�������ͼ��", chk(0).Value, glngSys, 1255)
    Call zlDatabase.SetPara("�ʿغ�", txt.Text, glngSys, 1255, InStr(mstrPrivs, ";����ѡ������;") > 0)
    Call zlDatabase.SetPara("���µ�����ӡ����˵��", chk(1).Value, glngSys, 1255, InStr(mstrPrivs, ";����ѡ������;") > 0)
    
End Sub

Private Sub pic��ʼ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If pic��ʼ.Top + Y > UD��ʼ.Max Or pic��ʼ.Top + Y < UD��ʼ.Min Then Exit Sub
        pic��ʼ.Top = pic��ʼ.Top + Y
        UD��ʼ.Value = pic��ʼ.Top
        Call DrawPage
        Me.Refresh
    End If
End Sub

Private Sub txtBegin_Change()
    If Not mblnInit Then opt����.Value = True
End Sub

Private Sub txtBegin_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtBegin_GotFocus()
    zlControl.TxtSelAll txtBegin
End Sub

Private Sub txtEnd_Change()
    If Not mblnInit Then opt����.Value = True
End Sub

Private Sub txtEnd_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtEnd_GotFocus()
    zlControl.TxtSelAll txtEnd
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
    Dim bln���� As Boolean
    bln���� = False
    If Not (Val(txt��ʼ.Text) >= UD��ʼ.Min And Val(txt��ʼ.Text) <= UD��ʼ.Max) Then
        MsgBox "��ʼλ��Ӧ���� " & UD��ʼ.Min & " �� " & UD��ʼ.Max & " ֮�䣡", vbInformation, gstrSysName
        txt��ʼ.SetFocus: Exit Function
    End If
    
    mstrPage = ""
    
    If opt��ǰ.Value Then
        mintPrintRange = 0
    ElseIf opt����.Value Then
        mintPrintRange = 1
        bln���� = True
    Else
        mintPrintRange = 2
    End If
    
    If bln���� = True Then
        If mlngAllPage > 0 Then
            If CInt(txtBegin.Text) < 1 Or CInt(txtBegin.Text) > mlngAllPage Then
                MsgBox "�����ҳ�볬����Χ�ˡ�", vbCritical, gstrSysName
                txtBegin.SelStart = 0
                txtBegin.SelLength = 5
                If txtBegin.Enabled = True Then txtBegin.SetFocus
                Exit Function
            End If
            If CInt(txtEnd.Text) < 1 Or CInt(txtEnd.Text) > mlngAllPage Then
                MsgBox "�����ҳ�볬����Χ�ˡ�", vbCritical, gstrSysName
                txtEnd.SelStart = 0
                txtEnd.SelLength = 5
                If txtEnd.Enabled = True Then txtEnd.SetFocus
                Exit Function
            End If
        End If
        If CInt(txtEnd.Text) < CInt(txtBegin.Text) Then
            MsgBox "��ʼҳ�볬���˽���ҳ�롣", vbCritical, gstrSysName
            txtBegin.SelStart = 0
            txtBegin.SelLength = 5
            If txtBegin.Enabled = True Then txtBegin.SetFocus
            Exit Function
        End If
        mstrPage = Val(txtBegin.Text) - 1 & ";" & Val(txtEnd.Text) - 1
    End If
    
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

