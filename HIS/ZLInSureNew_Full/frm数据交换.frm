VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm���ݽ��� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ڽ������ݽ���..."
   ClientHeight    =   2700
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4560
   Icon            =   "frm���ݽ���.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Timer Timsearch 
      Interval        =   100
      Left            =   1320
      Top             =   2160
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Enabled         =   0   'False
      Height          =   372
      Left            =   2280
      TabIndex        =   7
      Top             =   2160
      Width           =   972
   End
   Begin MSComCtl2.Animation Avi 
      Height          =   492
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   612
      _ExtentX        =   1085
      _ExtentY        =   873
      _Version        =   393216
      FullWidth       =   51
      FullHeight      =   41
   End
   Begin VB.Timer TimWrite 
      Interval        =   100
      Left            =   360
      Top             =   2160
   End
   Begin VB.CommandButton cmdCancle 
      Caption         =   "ȡ��(&C)"
      Height          =   372
      Left            =   3360
      TabIndex        =   8
      Top             =   2160
      Width           =   972
   End
   Begin VB.PictureBox PicWrite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   0
      Picture         =   "frm���ݽ���.frx":000C
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   355
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1560
      Width           =   5325
   End
   Begin VB.Label lbl��ʾβ 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1008
      TabIndex        =   4
      Top             =   1224
      Width           =   120
   End
   Begin VB.Label lbl��ʾͷ 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1008
      TabIndex        =   2
      Top             =   456
      Width           =   120
   End
   Begin VB.Label lblҽ������ 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1008
      TabIndex        =   3
      Top             =   840
      Width           =   96
   End
   Begin VB.Label lbl�ȴ���ʾ 
      AutoSize        =   -1  'True
      Caption         =   "���ڵȴ�ҽ�������ļ�..."
      Height          =   180
      Left            =   1200
      TabIndex        =   6
      Top             =   1800
      Width           =   2088
   End
   Begin VB.Label lbl�л����� 
      AutoSize        =   -1  'True
      Caption         =   "�����л���ҽ��������������Ӧ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   144
      TabIndex        =   0
      Top             =   120
      Width           =   4080
   End
End
Attribute VB_Name = "frm���ݽ���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFile As String, mstrStream As String
Private mreturn As Boolean
Private mbytType As Byte
Private mlng����ID As Long
Private strFile As String

Private Sub cmdCancle_Click()
    If MsgBox("��ע��:���ҽ������ѽ���,�벻Ҫʹ�ô˹���,���ܻ�������߽�ƽ��" & vbCrLf & _
        "ȡ�������ļ���ȡ��", vbYesNo + vbInformation + vbDefaultButton2, gstrSysName) = vbYes Then
        mstrStream = ""
        Me.Hide
    End If
End Sub
Private Sub cmdOK_Click()
    Dim strIdentify As String
    Dim strAddition As String
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    strIdentify = "": strAddition = ""
    If mbytType = 1 Then
        Set mdomInput = New MSXML2.DOMDocument
        mdomInput.Load ("c:\njyb\zydjxx.xml")
        Set nodRowset = mdomInput.documentElement.selectSingleNode("RECORD")
        strIdentify = nodRowset.selectSingleNode("TBR").Text & ";"                                     '0����
        strIdentify = strIdentify & nodRowset.selectSingleNode("TBR").Text & ";"                    '1ҽ���ţ����˱�ţ�
        strIdentify = strIdentify & ";"                                 '2����
        strIdentify = strIdentify & nodRowset.selectSingleNode("XM").Text & ";"                   '3����
        strIdentify = strIdentify & nodRowset.selectSingleNode("XB").Text & ";"                               '4�Ա�
        strIdentify = strIdentify & ";"                             '5��������
        strIdentify = strIdentify & nodRowset.selectSingleNode("SFZH").Text & ";"                                 '6���֤
        strIdentify = strIdentify & ";"                              '7.��λ����(����)
        strAddition = "0;"                                          '8.���Ĵ���
        strAddition = strAddition & nodRowset.selectSingleNode("XH").Text & ";"                             '9.˳���
        strAddition = strAddition & nodRowset.selectSingleNode("XZMC").Text & ";"                        '10��Ա���
        strAddition = strAddition & nodRowset.selectSingleNode("ZHYE").Text & ";"                             '11�ʻ����
        strAddition = strAddition & "0;"                           '12��ǰ״̬
        strAddition = strAddition & ";"                            '13����ID
        strAddition = strAddition & "1;"                           '14��ְ(1,2,3)
        strAddition = strAddition & ";"                             '15����֤��
        strAddition = strAddition & ";"                             '16�����
        strAddition = strAddition & ";"                             '17�Ҷȼ�
        strAddition = strAddition & ";"                             '18�ʻ������ۼ�
        strAddition = strAddition & "0;"                            '19�ʻ�֧���ۼ�
        strAddition = strAddition & "0;"                            '20���깤���ܶ�
        strAddition = strAddition & "0"                            '21סԺ�����ۼ�
    
        mlng����ID = BuildPatiInfo(1, strIdentify & strAddition, mlng����ID, TYPE_�Ͼ���)
        '���ظ�ʽ:�м���벡��ID
        If mlng����ID > 0 Then
            mstrStream = strIdentify & mlng����ID & ";" & strAddition
        End If
    End If
    
    TimWrite.Enabled = False
    mreturn = True
    Call DebugTool("�����ѽ��գ����أ���ǰmbytType=" & mbytType)
    Me.Hide
'    Call Kill("C:\NJYB\zydjxx.xml")
End Sub

Private Sub Form_Load()
    cmdOK.Enabled = False
    mstrFile = gstrAviPath & "\FINDFILE.AVI"
    Call aviMove
End Sub
Private Sub aviMove()
    On Error Resume Next
    With Avi
        .Open (mstrFile)
        .AutoPlay = True
        .Play
    End With
End Sub

Private Sub Timsearch_Timer()
    Dim strTemp As String
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    Dim xm As String
    Select Case mbytType
        Case 1
            strFile = "C:\NJYB\ZYDJXX.XML"
        Case 9
            strFile = "C:\NJYB\CYJSD.XML"
        Case Else
            strFile = "C:\NJYB\MZJSHZ.XML"
    End Select
    If Not FileExists(strFile) Then Exit Sub
    
    strTemp = Trim(mdl�Ͼ���.readTxtFile(strFile))
    If strTemp <> "" Then
'        Timsearch.Enabled = False
        mstrStream = strTemp
        cmdOK.Enabled = True
    Else
        lbl��ʾͷ.Caption = ""
        lblҽ������.Caption = ""
        lbl��ʾβ.Caption = ""
        cmdOK.Enabled = False
        Exit Sub
    End If
    
    If mbytType = 9 Then
        lbl��ʾͷ.Caption = "�ѷ��ֽ����ļ�"
        lblҽ������.Caption = ""
        lbl��ʾβ.Caption = ""
    Else
'        Set mdomInput = New MSXML2.DOMDocument
'        mdomInput.Load ("c:\njyb\mzjshz.xml")
'        Set nodRowset = mdomInput.documentElement.selectSingleNode("RECORD")
        lbl��ʾͷ.Caption = "����ҽ������:"
'        lblҽ������.Caption = nodRowset.selectSingleNode("TBR").Text
'        xm = nodRowset.selectSingleNode("XM").Text
'        lbl��ʾβ.Caption = IIf(mbytType = 1, "����:" & xm, "�Ƿ���ȷ��")
    End If
End Sub

Private Sub TimWrite_Timer()
    Static i As Long
    i = i + 20
    If i > PicWrite.ScaleWidth Then i = 1
    
    Call PicWrite.PaintPicture(PicWrite, i, 0, PicWrite.ScaleWidth - i, PicWrite.ScaleHeight, 0, 0, PicWrite.ScaleWidth - i, PicWrite.ScaleHeight)
    Call PicWrite.PaintPicture(PicWrite, 0, 0, i, PicWrite.ScaleHeight, PicWrite.ScaleWidth - i, 0, i, PicWrite.ScaleHeight)
End Sub

Public Function getFeeBalance(Optional bytType As Byte, Optional lng����ID As Long) As String
    mbytType = bytType
    mlng����ID = lng����ID
    Me.Show 1
     
    If mreturn Then
        lng����ID = mlng����ID
        getFeeBalance = mstrStream
    End If
End Function
