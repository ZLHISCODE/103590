VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCaseTendBodyPrintSet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���µ���ӡ"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   Icon            =   "frmCaseTendBodyPrintSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin TabDlg.SSTab SSTab1 
      Height          =   5130
      Left            =   30
      TabIndex        =   3
      Top             =   90
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   9049
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "�����ӡ"
      TabPicture(0)   =   "frmCaseTendBodyPrintSet.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblIn"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "picInfo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "picCHKH"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "��ӡѡ��"
      TabPicture(1)   =   "frmCaseTendBodyPrintSet.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fra��ӡ"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.PictureBox picCHKH 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4365
         Left            =   15
         ScaleHeight     =   4365
         ScaleWidth      =   6690
         TabIndex        =   30
         Top             =   705
         Visible         =   0   'False
         Width           =   6690
         Begin VB.PictureBox picVsh 
            Appearance      =   0  'Flat
            BackColor       =   &H80000003&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1080
            Left            =   360
            ScaleHeight     =   1080
            ScaleWidth      =   1635
            TabIndex        =   32
            Top             =   1080
            Width           =   1635
            Begin VB.PictureBox picPrint 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   555
               Index           =   0
               Left            =   0
               ScaleHeight     =   555
               ScaleWidth      =   480
               TabIndex        =   33
               Top             =   -15
               Visible         =   0   'False
               Width           =   480
               Begin VB.Label lblNum 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  Caption         =   "9999"
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9
                     Charset         =   134
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   180
                  Index           =   0
                  Left            =   30
                  TabIndex        =   34
                  Top             =   315
                  Visible         =   0   'False
                  Width           =   435
               End
               Begin VB.Image imgIco 
                  Height          =   240
                  Index           =   0
                  Left            =   120
                  Picture         =   "frmCaseTendBodyPrintSet.frx":0044
                  Top             =   45
                  Width           =   240
               End
            End
         End
         Begin VB.VScrollBar vsc 
            Height          =   1815
            Left            =   4725
            SmallChange     =   50
            TabIndex        =   31
            Top             =   60
            Visible         =   0   'False
            Width           =   200
         End
      End
      Begin VB.PictureBox picInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   75
         ScaleHeight     =   270
         ScaleWidth      =   6525
         TabIndex        =   24
         Top             =   375
         Width           =   6525
         Begin VB.OptionButton optSelect 
            Caption         =   "����"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   29
            Top             =   15
            Width           =   660
         End
         Begin VB.OptionButton optSelect 
            Caption         =   "�Ѵ�"
            Height          =   180
            Index           =   1
            Left            =   930
            TabIndex        =   28
            Top             =   15
            Width           =   660
         End
         Begin VB.OptionButton optSelect 
            Caption         =   "�ش�"
            Height          =   180
            Index           =   2
            Left            =   1860
            TabIndex        =   27
            Top             =   15
            Width           =   660
         End
         Begin VB.OptionButton optSelect 
            Caption         =   "����"
            Height          =   180
            Index           =   3
            Left            =   2805
            TabIndex        =   26
            Top             =   15
            Value           =   -1  'True
            Width           =   660
         End
         Begin VB.CheckBox chkSelect 
            Caption         =   "ȫѡ"
            Height          =   180
            Left            =   5910
            TabIndex        =   25
            Top             =   15
            Width           =   705
         End
         Begin VB.Image imgIcon 
            DragIcon        =   "frmCaseTendBodyPrintSet.frx":05CE
            Height          =   240
            Index           =   0
            Left            =   660
            MouseIcon       =   "frmCaseTendBodyPrintSet.frx":0CB8
            Picture         =   "frmCaseTendBodyPrintSet.frx":1242
            Top             =   -15
            Width           =   240
         End
         Begin VB.Image imgIcon 
            DragIcon        =   "frmCaseTendBodyPrintSet.frx":17CC
            Height          =   240
            Index           =   1
            Left            =   1575
            MouseIcon       =   "frmCaseTendBodyPrintSet.frx":1EB6
            Picture         =   "frmCaseTendBodyPrintSet.frx":2440
            Top             =   -15
            Width           =   240
         End
         Begin VB.Image imgIcon 
            DragIcon        =   "frmCaseTendBodyPrintSet.frx":29CA
            Height          =   240
            Index           =   2
            Left            =   2505
            MouseIcon       =   "frmCaseTendBodyPrintSet.frx":30B4
            Picture         =   "frmCaseTendBodyPrintSet.frx":363E
            Top             =   -15
            Width           =   240
         End
      End
      Begin VB.Frame fra��ӡ 
         Caption         =   "��ӡҳ��"
         Height          =   1080
         Left            =   -74760
         TabIndex        =   8
         Top             =   420
         Width           =   4380
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
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   2175
            Visible         =   0   'False
            Width           =   2130
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
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   315
               Width           =   1170
            End
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
               TabIndex        =   17
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
         End
         Begin VB.TextBox txtҳ�� 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3450
            MaxLength       =   3
            TabIndex        =   15
            Text            =   "1"
            Top             =   360
            Width           =   285
         End
         Begin VB.TextBox txt��ʼ 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1050
            MaxLength       =   3
            TabIndex        =   12
            Text            =   "25"
            Top             =   1680
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.CheckBox chkҳ�� 
            Caption         =   "��ӡҳ�ţ���һҳҳ�ű�ʾΪ(&4)"
            Height          =   195
            Left            =   525
            TabIndex        =   11
            Top             =   405
            Value           =   1  'Checked
            Width           =   2910
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
         Begin VB.CheckBox chkOper 
            Caption         =   "��ӡ��ӡ��(&6)"
            Height          =   195
            Left            =   2625
            TabIndex        =   9
            Top             =   765
            Value           =   1  'Checked
            Width           =   1650
         End
         Begin MSComCtl2.UpDown UDҳ�� 
            Height          =   300
            Left            =   3735
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   345
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtҳ��"
            BuddyDispid     =   196624
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
            TabIndex        =   14
            Top             =   1680
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Value           =   25
            BuddyControl    =   "txt��ʼ"
            BuddyDispid     =   196625
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            Height          =   180
            Left            =   1965
            TabIndex        =   21
            Top             =   1710
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ʼλ��"
            Height          =   180
            Left            =   255
            TabIndex        =   20
            Top             =   1740
            Visible         =   0   'False
            Width           =   720
         End
      End
      Begin VB.Frame fra 
         Caption         =   "����"
         Height          =   1440
         Left            =   -74775
         TabIndex        =   4
         Top             =   1560
         Width           =   4380
         Begin VB.CheckBox chk 
            Caption         =   "����ӡ���µ��·�������˵����Ϣ(&9)"
            Height          =   195
            Index           =   1
            Left            =   900
            TabIndex        =   22
            Top             =   1080
            Value           =   1  'Checked
            Width           =   3360
         End
         Begin VB.TextBox txt 
            Height          =   300
            Left            =   960
            TabIndex        =   6
            Top             =   255
            Width           =   3210
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ӡ���ʺ�����������ߺ���Ӱ(&8)"
            Height          =   195
            Index           =   0
            Left            =   915
            TabIndex        =   5
            Top             =   705
            Value           =   1  'Checked
            Width           =   3360
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "�ʿغ�(&7)"
            Height          =   180
            Left            =   120
            TabIndex        =   7
            Top             =   300
            Width           =   810
         End
      End
      Begin VB.Label lblIn 
         Caption         =   "��11ҳ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   2025
         TabIndex        =   36
         Top             =   60
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "˫��ѡ�л�ȡ��(��SHIFT�����Ƭ��Χѡ��)"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   3180
         TabIndex        =   35
         Top             =   30
         Width           =   3510
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��ӡ(&P)"
      Height          =   350
      Left            =   90
      TabIndex        =   0
      Top             =   5310
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5655
      TabIndex        =   1
      Top             =   5310
      Width           =   1100
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Ԥ��(&V)"
      Height          =   350
      Left            =   90
      TabIndex        =   2
      Top             =   5310
      Width           =   1100
   End
   Begin VB.Label lblSelect 
      Caption         =   "��ѡҳ��"
      Height          =   180
      Left            =   1290
      TabIndex        =   23
      Top             =   5400
      Width           =   4125
   End
End
Attribute VB_Name = "frmCaseTendBodyPrintSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytOpt As Byte

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
Private mblnShift As Boolean

Private mstrPrivs As String

Private Sub chkSelect_Click()
    Dim i As Long
    For i = 0 To picPrint.Count - 1
        If picPrint(i).Visible = True Then
            picPrint(i).BackColor = IIf(chkSelect.Value = 0, &HE0E0E0, vbRed)
            picPrint(i).Tag = IIf(chkSelect.Value = 0, "0", "1")
        End If
    Next
    Call SetSelectInfo
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
    Call zlDatabase.SetPara("�ʿغ�", txt.Text, glngSys, 1255, InStr(mstrPrivs, ";����ѡ������;") > 0)
    If Not GetValue Then Exit Sub
    mbytOpt = 2
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 39 Then KeyCode = 0
    If Shift = 1 And KeyCode = vbKeyShift Then
        mblnShift = True
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyShift Then
        mblnShift = False
    End If
End Sub

Private Sub Form_Load()
    mbytOpt = 0
    mblnShift = False
    
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
    
    mintPrintRange = Val(zlDatabase.GetPara("������ӡ", glngSys, 1255, "1"))
    
    chkҳ��.Value = Val(zlDatabase.GetPara("��ӡҳ��", glngSys, 1255, "1", Array(chkҳ��)))
    txtҳ��.Text = Val(zlDatabase.GetPara("��ʼҳ��", glngSys, 1255, "1", Array(txtҳ��, UDҳ��)))
    chk����.Value = Val(zlDatabase.GetPara("��ӡ����", glngSys, 1255, "0", Array(chk����)))
    '67405:������,2013-11-25
    chkOper.Value = Val(zlDatabase.GetPara("��ӡ��ӡ��", glngSys, 1255, "0", Array(chkOper)))
    chk(0).Value = Val(zlDatabase.GetPara("����ӡ�������ͼ��", glngSys, 1255, "0", Array(chk(0))))
    
    mintBeginPage = Val(txtҳ��.Text)
    
    UDҳ��.Value = IIf(mintBeginPage = 0, 1, mintBeginPage)

End Sub

Public Function PrintSet(objParent As Object, ByVal strParam As String, ByRef intPrintRange As Integer, ByRef lngBeginY As Long, ByRef intBeginPage As Integer, strPage As String, ByVal strPrivs As String, ByVal bytMode As Byte) As Byte
'���ܣ����ô�ӡѡ��
'������blnFirst=�Ƿ��һ�ε���,����ֻ��"ȷ��","ȡ��",�Ҳ������޸Ĳ�����ӡ����
'      strParam �ɵ�ǰҳ������ӡ�� ��Ҫ��ȡ �ļ�ID;�������µ���ҳ��
'      blnCurCase=T=ֻ��ӡ��ǰ����,F=�ӵ�ǰ������ʼ������ӡ����
'      lngBeginY=���β�����ʼ��ӡλ��'mm
'      intBeginPage=��ʼҳ��,Ϊ0��ʾ����ӡҳ��
'      strPage
'���أ�0-ȡ��,1-Ԥ��,2-��ӡ
    
    mstrPrivs = strPrivs
    
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
    lblIn.Caption = "��" & mlngAllPage & "ҳ"
    
    Call GetPageNum(mlng�ļ�ID)
    mblnInit = False
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
    Dim lngPage As Long
    
    On Error GoTo Errhand
    lngPage = 1
    Call LoadPages
    strSQL = "select ��ӡҳ��,��ӡ�� From ���µ���ӡ where �ļ�ID=[1] Order by ��ӡҳ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��ӡ����", lng�ļ�ID)
    
    Do While Not rsTemp.EOF
        lngPage = Val("" & rsTemp!��ӡҳ��)
        If lngPage > 0 And lngPage <= picPrint.Count Then
            If "" & rsTemp!��ӡ�� = "" Then
                imgIco(lngPage - 1).Picture = imgIcon(2).Picture
                imgIco(lngPage - 1).Tag = "�ش�"
            Else
                imgIco(lngPage - 1).Picture = imgIcon(1).Picture
                imgIco(lngPage - 1).Tag = "�Ѵ�"
            End If
        End If
        rsTemp.MoveNext
    Loop
    
    GetPageNum = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub LoadPages()
    Dim i As Long
    
    picCHKH.Visible = mlngAllPage > 0
   
    For i = 0 To mlngAllPage - 1
        If i > 0 Then
            Load picPrint(i)
            Load imgIco(i)
            imgIco(i).Left = imgIco(0).Left
            imgIco(i).Top = imgIco(0).Top
            Set imgIco(i).Container = picPrint(i)
            Load lblNum(i)
            lblNum(i).Left = lblNum(0).Left
            lblNum(i).Top = lblNum(0).Top
            Set lblNum(i).Container = picPrint(i)
        End If
        imgIco(i).Tag = "����"
        imgIco(i).Visible = True
        imgIco(i).ZOrder 0
        lblNum(i).Visible = True
        lblNum(i).ZOrder 0
        lblNum(i).Caption = i + 1
        picPrint(i).Tag = "0"
        picPrint(i).BackColor = &HE0E0E0
    Next
    Call ShowCard
End Sub

Private Sub ShowCard(Optional ByVal strTag As String = "ALL")
    '��ʾ�����п�Ƭλ��
    Dim i As Long, j As Long
    Dim lngLeft As Long, lngTop As Long
    Dim lngwNum As Long, lngHnum As Long
    Dim lngPageCount As Long, lngHeight As Long
    Dim lngCHeight As Long
    
    Call LockWindowUpdate(Me.hWnd)
    
    With picVsh
        .Left = 0
        .Top = 0
        .Width = picCHKH.Width
        .Height = picCHKH.Height
        .BackColor = picCHKH.BackColor
    End With
    lngHeight = picVsh.Height
    vsc.Value = 0
    vsc.Visible = False
    For i = 0 To picPrint.Count - 1
        If imgIco(i).Tag = strTag Or strTag = "ALL" Then
            lngPageCount = lngPageCount + 1
        End If
    Next
     '-----�����Ƿ���ʾ������(��Ƭ����̶�)
    '����ÿҳ����ʾ�Ŀ�Ƭ��Ŀ
    lngwNum = picVsh.Width \ (picPrint(0).Width + 60)
    '����߶��Ƿ񳬳�����
    lngHnum = (lngPageCount \ lngwNum)
    If lngwNum * lngHnum < lngPageCount Then lngHnum = lngHnum + 1
    If lngHnum * (picPrint(0).Height + 60) > picVsh.Height Then
        lngHeight = lngHnum * (picPrint(0).Height + 60)
    End If
    '˵����ʾ������
    If lngHeight > picCHKH.Height Then
        lngwNum = (picVsh.Width - vsc.Width) \ (picPrint(0).Width + 60)
        lngHnum = (lngPageCount \ lngwNum)
        If lngwNum * lngHnum < lngPageCount Then lngHnum = lngHnum + 1
        If lngHnum * (picPrint(0).Height + 60) > picVsh.Height Then
            lngHeight = lngHnum * (picPrint(0).Height + 60)
        End If
    End If
    picVsh.Height = lngHeight
    lngCHeight = picVsh.Height - picCHKH.Height
    If lngCHeight > 0 Then
        If lngCHeight < picPrint(0).Height Then
            vsc.Max = 1
            vsc.SmallChange = 1
        Else
            vsc.Max = lngCHeight
            vsc.SmallChange = picPrint(0).Height + 60
        End If
        vsc.Top = 0
        vsc.Left = picCHKH.Width - vsc.Width
        vsc.Height = picCHKH.Height
        vsc.Visible = True
        picVsh.Width = vsc.Left
    End If
    
    j = -1
    For i = 0 To picPrint.Count - 1
        If imgIco(i).Tag = strTag Or strTag = "ALL" Then
            If j > -1 Then
                lngLeft = picPrint(j).Left + picPrint(j).Width + 60
                If lngLeft + picPrint(i).Width > picVsh.Width Then
                    lngTop = picPrint(j).Top + picPrint(j).Height + 60
                    lngLeft = 60
                Else
                    lngTop = picPrint(j).Top
                End If
            Else
                lngLeft = 60
                lngTop = 60
            End If
            j = i
            picPrint(i).Left = lngLeft
            picPrint(i).Top = lngTop
            picPrint(i).Visible = True
        End If
    Next
    Call LockWindowUpdate(0)
End Sub

Private Sub SetSelectInfo()
    Dim i As Long
    Dim strInfo As String
    Dim lngStartPage As Long, lngEndPage As Long, lngPrePage As Long
    
    For i = 0 To picPrint.Count - 1
        If Val(picPrint(i).Tag) = "1" Then
            If lngPrePage = 0 Then
                lngStartPage = i + 1
                lngEndPage = lngStartPage
            ElseIf lngPrePage = i Then
                lngEndPage = i + 1
            Else
                strInfo = strInfo & "��" & lngStartPage & IIf(lngEndPage = lngStartPage, "", "-" & lngEndPage)
                lngStartPage = i + 1
                lngEndPage = lngStartPage
            End If
            lngPrePage = i + 1
        End If
    Next
    If lngStartPage > 0 Then
        strInfo = strInfo & "��" & lngStartPage & IIf(lngEndPage = lngStartPage, "", "-" & lngEndPage)
    End If
    If Left(strInfo, 1) = "��" Then strInfo = Mid(strInfo, 2)
    lblSelect.Caption = "��ѡҳ�룺" & strInfo
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call zlDatabase.SetPara("������ӡ", mintPrintRange, glngSys, 1255)
    Call zlDatabase.SetPara("��ӡҳ��", chkҳ��.Value, glngSys, 1255)
    Call zlDatabase.SetPara("��ʼҳ��", Val(txtҳ��.Text), glngSys, 1255)
    Call zlDatabase.SetPara("��ӡ����", chk����.Value, glngSys, 1255)
    '67405:������,2013-11-25,���"��ӡ��ӡ��"
    Call zlDatabase.SetPara("��ӡ��ӡ��", chkOper.Value, glngSys, 1255)
    Call zlDatabase.SetPara("����ӡ�������ͼ��", chk(0).Value, glngSys, 1255)
    Call zlDatabase.SetPara("�ʿغ�", txt.Text, glngSys, 1255, InStr(mstrPrivs, ";����ѡ������;") > 0)
End Sub

Private Sub imgIco_Click(Index As Integer)
    Call picPrint_Click(Index)
End Sub

Private Sub imgIco_DblClick(Index As Integer)
    Call picPrint_DblClick(Index)
End Sub

Private Sub lblNum_Click(Index As Integer)
    Call picPrint_Click(Index)
End Sub

Private Sub lblNum_DblClick(Index As Integer)
    Call picPrint_DblClick(Index)
End Sub

Private Sub optSelect_Click(Index As Integer)
    Dim i As Long
    Dim strTag As String
    
    If Me.Visible = False Then Exit Sub
    If optSelect(0).Value = True Then
        strTag = "����"
    ElseIf optSelect(1).Value = True Then
        strTag = "�Ѵ�"
    ElseIf optSelect(2).Value = True Then
        strTag = "�ش�"
    Else
        strTag = "ALL"
    End If
    For i = 0 To picPrint.Count - 1
        picPrint(i).Tag = "0"
        picPrint(i).Visible = False
        picPrint(i).BackColor = &HE0E0E0
    Next
    Call ShowCard(strTag)
    If chkSelect.Value = 1 Then Call chkSelect_Click
End Sub

Private Sub picPrint_Click(Index As Integer)
    '����Ƿ���shift����
    Dim blnShift As Boolean
    Dim lngIndex As Long, lngStartIndex As Long, lngEndIndex As Long
    Dim i As Long
    
    lngIndex = -1
    If mblnShift = False Then Exit Sub
    For i = Index - 1 To 0 Step -1
        If Val(picPrint(i).Tag) = 1 Then
            lngIndex = i
            Exit For
        End If
    Next
    If lngIndex = -1 Then
        For i = Index + 1 To picPrint.Count - 1
            If Val(picPrint(i).Tag) = 1 Then
                lngIndex = i
                Exit For
            End If
        Next
    End If
    If lngIndex <> -1 Then
        If lngIndex > Index Then
            lngStartIndex = Index
            lngEndIndex = lngIndex
        Else
            lngStartIndex = lngIndex
            lngEndIndex = Index
        End If
        For i = lngStartIndex To lngEndIndex
            picPrint(i).Tag = 1
            picPrint(i).BackColor = IIf(Val(picPrint(i).Tag) = 1, vbRed, &HE0E0E0)
        Next
        Call SetSelectInfo
    End If
End Sub

Private Sub picPrint_DblClick(Index As Integer)
    picPrint(Index).Tag = IIf(picPrint(Index).BackColor = &HE0E0E0, 1, 0)
    picPrint(Index).BackColor = IIf(Val(picPrint(Index).Tag) = 1, vbRed, &HE0E0E0)
    Call SetSelectInfo
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
    Dim i As Long
    Dim arrPage
    
    bln���� = False
    If Not (Val(txt��ʼ.Text) >= UD��ʼ.Min And Val(txt��ʼ.Text) <= UD��ʼ.Max) Then
        MsgBox "��ʼλ��Ӧ���� " & UD��ʼ.Min & " �� " & UD��ʼ.Max & " ֮�䣡", vbInformation, gstrSysName
        txt��ʼ.SetFocus: Exit Function
    End If
    
    arrPage = Array()
    For i = 0 To picPrint.Count - 1
        If picPrint(i).Visible = True And Val(picPrint(i).Tag) = 1 Then
            ReDim Preserve arrPage(UBound(arrPage) + 1)
            arrPage(UBound(arrPage)) = i
        End If
    Next i
    If UBound(arrPage) = -1 Then
        MsgBox "��ѡ��Ҫ��ӡ��ҳ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    mstrPage = ""
    If optSelect(3).Value = True And chkSelect.Value = 1 Then
        mintPrintRange = 2 'ȫ����ӡ
        mstrPage = 0 & ";" & UBound(arrPage)
    ElseIf UBound(arrPage) = 0 Then
        mintPrintRange = 0 '��ǰҳ
        mstrPage = arrPage(0)
    Else
        mintPrintRange = 1 '��ҳ��ӡ
        mstrPage = Join(arrPage, ";")
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

Private Sub vsc_Change()
    picVsh.Top = (-1) * vsc.Value
End Sub

