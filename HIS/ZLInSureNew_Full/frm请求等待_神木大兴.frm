VERSION 5.00
Begin VB.Form frm����ȴ�_��ľ���� 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ȴ�����"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   ControlBox      =   0   'False
   LinkTopic       =   "frmWait"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   110
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1935
      Top             =   -90
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3480
      TabIndex        =   3
      Top             =   1200
      Width           =   1100
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   0
      Picture         =   "frm����ȴ�_��ľ����.frx":0000
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   335
      TabIndex        =   0
      Top             =   945
      Width           =   5025
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3435
      Top             =   -150
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   255
      TabIndex        =   4
      Top             =   555
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   255
      Picture         =   "frm����ȴ�_��ľ����.frx":096C
      Stretch         =   -1  'True
      Top             =   105
      Width           =   480
   End
   Begin VB.Label lbl���� 
      BackStyle       =   0  'Transparent
      Caption         =   "�Ѿ��ύ�������ڵȴ�������Ӧ...."
      Height          =   180
      Left            =   1020
      TabIndex        =   1
      Top             =   450
      Width           =   4140
   End
   Begin VB.Label lblBack 
      BackColor       =   &H8000000A&
      Height          =   630
      Left            =   -30
      TabIndex        =   2
      Top             =   1035
      Width           =   5895
   End
End
Attribute VB_Name = "frm����ȴ�_��ľ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType  As Byte   '0-����,1-סԺ
Private mstr���� As String   'IC����
Private mblnOK As Boolean

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub Form_Load()
    lblInfor.Caption = Decode(mbytType, 0, "�� ��", "ס Ժ")
End Sub

Private Sub Timer1_Timer()
    Static i As Long
    i = i + 20
    If i >= Picture1.ScaleWidth Then i = 1
    Picture1.PaintPicture Picture1.Picture, i, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight, 0, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight
    Picture1.PaintPicture Picture1.Picture, 0, 0, i, Picture1.ScaleHeight, Picture1.ScaleWidth - i, 0, i, Picture1.ScaleHeight
End Sub
Public Function ShowWait(ByVal bytType As Byte, ByVal str���� As String) As Boolean
    '����:��ʾ�ȴ�����
    'bytType :0-����,1-סԺ
    mbytType = bytType
    mstr���� = str����
    Me.Show 1
    ShowWait = mblnOK
End Function

Private Sub Timer2_Timer()
    Timer2.Enabled = False
    If ISCHECKDATA = True Then
        mblnOK = True
        Unload Me
        Exit Sub
    End If
    
    Timer2.Enabled = True
End Sub
Private Function ISCHECKDATA() As Boolean
    '����:����������
    Dim rsTemp As New ADODB.Recordset
    DebugTool "��ʼ��������Ϣ"
    
    ISCHECKDATA = False
    Select Case mbytType
    Case 0  '����
        gstrSQL = "" & _
            "   Select  ybkh ҽ������, cfbh �������, jssj ����ʱ��, jsbz ҽ�������־, " & _
            "           fyhj �����ܷ���, kszf ����֧��, tczf ͳ��֧��, ybje Ӧ���ֽ��, xm ��������   " & _
            "   From MZ_JSLSB  " & _
            "   Where upper(JSBZ) ='T' and ybkh='" & mstr���� & "'"
    Case Else   'סԺ
        gstrSQL = "" & _
           "   Select  ybkh ҽ������, zybh סԺ���, rysj ��Ժʱ��, cysj ����ʱ��, jsbz ҽ�������־, tpbz ҽ����Ʊ��־, " & _
           "           yybz ҽԺ�����־, fyhj �����ܷ���, kszf ����֧��, tczf ͳ��֧��, gwycb ����Ա����," & _
           "           yj Ѻ���ܶ�, ybje Ӧ���ֽ��, gfcwf ���Ѵ�λ��, zfcwf �ԷѴ�λ��, gftwf ���ѵ��·�, zftwf �Էѵ��·� " & _
           "   from zy_jslsb   " & _
           "   where upper(JSBZ)='T'  and ybkh='" & mstr���� & "'"
    End Select
    OpenRecordset_��ľ���� rsTemp, "��ȡ������Ϣ", gstrSQL
    If rsTemp.EOF Then
        Exit Function
    End If
    DebugTool "��������Ϣ�ɹ�"
    ISCHECKDATA = True
End Function
