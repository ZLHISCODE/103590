VERSION 5.00
Begin VB.Form frmPartogramPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ѡ��"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5355
   Icon            =   "frmPartogramPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fra1 
      Caption         =   "�������߱�־(�쳣��)"
      Height          =   1005
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1215
      Width           =   3735
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   8
         ItemData        =   "frmPartogramPara.frx":000C
         Left            =   1110
         List            =   "frmPartogramPara.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   600
         Width           =   2100
      End
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   7
         ItemData        =   "frmPartogramPara.frx":0043
         Left            =   1110
         List            =   "frmPartogramPara.frx":0050
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   2100
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��¶�½�"
         Height          =   180
         Index           =   7
         Left            =   360
         TabIndex        =   8
         Top             =   675
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Index           =   6
         Left            =   360
         TabIndex        =   6
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.ComboBox cboBody 
      Height          =   300
      Index           =   6
      ItemData        =   "frmPartogramPara.frx":007A
      Left            =   2205
      List            =   "frmPartogramPara.frx":0084
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   4695
      Width           =   1650
   End
   Begin VB.CheckBox chk 
      Caption         =   "����ͼ����ʾ����ʱ��"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   3405
      Width           =   2490
   End
   Begin VB.CheckBox chk 
      Caption         =   "����ͼģʽΪ����ʽ(����Ϊ����ʽ)"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   3645
      Width           =   3330
   End
   Begin VB.CheckBox chk 
      Caption         =   "��¶�ߵ���ʾ�����(����Ϊ�Ҳ�)"
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   3885
      Width           =   3330
   End
   Begin VB.CheckBox chk 
      Caption         =   "����ͼ����ʾ������"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   18
      Top             =   4125
      Width           =   3330
   End
   Begin VB.ComboBox cboBody 
      Height          =   300
      Index           =   4
      ItemData        =   "frmPartogramPara.frx":00A0
      Left            =   1200
      List            =   "frmPartogramPara.frx":00AA
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   4365
      Width           =   735
   End
   Begin VB.ComboBox cboBody 
      Height          =   300
      Index           =   5
      ItemData        =   "frmPartogramPara.frx":00BA
      Left            =   3120
      List            =   "frmPartogramPara.frx":00C4
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   4365
      Width           =   735
   End
   Begin VB.Frame fra1 
      Caption         =   "�������߱�־(˳��)"
      Height          =   1005
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   0
         ItemData        =   "frmPartogramPara.frx":00D4
         Left            =   1110
         List            =   "frmPartogramPara.frx":00E1
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2100
      End
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   1
         ItemData        =   "frmPartogramPara.frx":010B
         Left            =   1110
         List            =   "frmPartogramPara.frx":0118
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   2100
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��¶�½�"
         Height          =   180
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   660
         Width           =   720
      End
   End
   Begin VB.Frame fra3 
      Height          =   5040
      Left            =   3960
      TabIndex        =   23
      Top             =   15
      Width           =   15
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4080
      TabIndex        =   24
      Top             =   360
      Width           =   1100
   End
   Begin VB.CommandButton cmdCanc 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4080
      TabIndex        =   25
      Top             =   840
      Width           =   1100
   End
   Begin VB.Frame fra2 
      Caption         =   "������ʩ��־"
      Height          =   1005
      Left            =   120
      TabIndex        =   10
      Top             =   2295
      Width           =   3735
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   3
         ItemData        =   "frmPartogramPara.frx":0142
         Left            =   1110
         List            =   "frmPartogramPara.frx":014C
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   600
         Width           =   2100
      End
      Begin VB.ComboBox cboBody 
         Height          =   300
         Index           =   2
         ItemData        =   "frmPartogramPara.frx":0168
         Left            =   1110
         List            =   "frmPartogramPara.frx":0175
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   240
         Width           =   2100
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��־λ��"
         Height          =   180
         Index           =   3
         Left            =   360
         TabIndex        =   13
         Top             =   660
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��־����"
         Height          =   180
         Index           =   2
         Left            =   360
         TabIndex        =   11
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����ͼ0�����һ�����ߵ�"
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   26
      Top             =   4755
      Width           =   2070
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��������ʾΪ         �쳣����ʾΪ"
      Height          =   180
      Index           =   5
      Left            =   120
      TabIndex        =   19
      Top             =   4395
      Width           =   2970
   End
End
Attribute VB_Name = "frmPartogramPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmMain As Object
Private mblnOK As Boolean
Private mstrPrivs As String

Public Function ShowPara(ByVal frmMain As Object, ByVal strPrivs As String) As Boolean
    Dim intLoop As Integer
    Dim strTmp As String
    Dim strSQL As String, strPar As String
    Dim curDate As Date, intDay As Integer, lngValue As Long
    Dim intStart As Integer
    
    mblnOK = False
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
    
    '��ʼ���µ����
    '------------------------------------------------------------------------------------------------------------------
    '˳��
    cboBody(0).Clear
    cboBody(0).AddItem "0-����ʾ"
    cboBody(0).AddItem "1-��ʾ���߼�ͷ"
    cboBody(0).AddItem "2-��ʾʵ�߼�ͷ"
    
    cboBody(1).Clear
    cboBody(1).AddItem "0-����ʾ"
    cboBody(1).AddItem "1-��ʾ���߼�ͷ"
    cboBody(1).AddItem "2-��ʾʵ�߼�ͷ"
    
    '73309:������,2014-06-24
    '�쳣��
    cboBody(7).Clear
    cboBody(7).AddItem "0-����ʾ"
    cboBody(7).AddItem "1-��ʾ���߼�ͷ"
    cboBody(7).AddItem "2-��ʾʵ�߼�ͷ"
    
    cboBody(8).Clear
    cboBody(8).AddItem "0-����ʾ"
    cboBody(8).AddItem "1-��ʾ���߼�ͷ"
    cboBody(8).AddItem "2-��ʾʵ�߼�ͷ"
    cboBody(8).AddItem "3-��ʾֱ������"
    
    cboBody(2).Clear
    cboBody(2).AddItem "0-����ʾ"
    cboBody(2).AddItem "1-��ʾ����"
    cboBody(2).AddItem "2-��ʾ��������"
    
    cboBody(3).Clear
    cboBody(3).AddItem "0-��������"
    cboBody(3).AddItem "1-��¶�½�"
    
    cboBody(4).Clear
    cboBody(4).AddItem "0-����"
    cboBody(4).AddItem "1-ʵ��"
    
    cboBody(5).Clear
    cboBody(5).AddItem "0-����"
    cboBody(5).AddItem "1-ʵ��"
    
    '73309:������,2014-06-24
    cboBody(6).Clear
    cboBody(6).AddItem "0-������"
    cboBody(6).AddItem "1-����������"
    cboBody(6).AddItem "2-��ʵ������"
    '�����������߱�־
    strTmp = zlDatabase.GetPara("�����������߱�־", glngSys, 1255, "1;1", Array(lbl(0), cboBody(0), lbl(1), cboBody(1)), InStr(mstrPrivs, "����ѡ������") > 0)
    For intLoop = 0 To 1
        If UBound(Split(strTmp, ";")) >= intLoop Then
            lngValue = Val(Split(strTmp, ";")(intLoop))
            If lngValue < 0 Or lngValue > cboBody(intLoop).ListCount - 1 Then lngValue = 0
            cboBody(intLoop).ListIndex = lngValue
        Else
            cboBody(intLoop).ListIndex = 0
        End If
    Next
    strTmp = cboBody(0).ListIndex & ";" & cboBody(1).ListIndex
    strTmp = zlDatabase.GetPara("�����������߱�־(��)", glngSys, 1255, strTmp, Array(lbl(6), cboBody(7), lbl(7), cboBody(8)), InStr(mstrPrivs, "����ѡ������") > 0)
    For intLoop = 0 To 1
        If UBound(Split(strTmp, ";")) >= intLoop Then
            lngValue = Val(Split(strTmp, ";")(intLoop))
            If lngValue < 0 Or lngValue > cboBody(intLoop + 7).ListCount - 1 Then lngValue = 0
            cboBody(intLoop + 7).ListIndex = lngValue
        Else
            cboBody(intLoop + 7).ListIndex = 0
        End If
    Next
    
    '����������ʩ��־
    strTmp = zlDatabase.GetPara("����������ʩ��־", glngSys, 1255, "1;1", Array(lbl(2), cboBody(2), lbl(3), cboBody(3)), InStr(mstrPrivs, "����ѡ������") > 0)
    For intLoop = 0 To 1
        If UBound(Split(strTmp, ";")) >= intLoop Then
            lngValue = Val(Split(strTmp, ";")(intLoop))
            If lngValue < 0 Or lngValue > cboBody(intLoop + 2).ListCount - 1 Then lngValue = 0
            cboBody(intLoop + 2).ListIndex = lngValue
        Else
            cboBody(intLoop + 2).ListIndex = 0
        End If
    Next
    '���̾����߱�־
    strTmp = zlDatabase.GetPara("���̾����쳣�߱�־", glngSys, 1255, "1;1", Array(lbl(5), cboBody(4), cboBody(5)), InStr(mstrPrivs, "����ѡ������") > 0)
    For intLoop = 0 To 1
        If UBound(Split(strTmp, ";")) >= intLoop Then
            lngValue = Val(Split(strTmp, ";")(intLoop))
            If lngValue < 0 Or lngValue > cboBody(intLoop + 4).ListCount - 1 Then lngValue = 0
            cboBody(intLoop + 4).ListIndex = lngValue
        Else
            cboBody(intLoop + 4).ListIndex = 0
        End If
    Next
    
    chk(0).Value = Val(zlDatabase.GetPara("����ͼ��ʾ����ʱ��", glngSys, 1255, "1", Array(chk(0)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(1).Value = Val(zlDatabase.GetPara("����ͼģʽ", glngSys, 1255, "0", Array(chk(1)), True))
    chk(2).Value = Val(zlDatabase.GetPara("��¶�ߵ���ʾλ��", glngSys, 1255, "0", Array(chk(2)), InStr(mstrPrivs, "����ѡ������") > 0))
    chk(3).Value = Val(zlDatabase.GetPara("����ͼ��ʾ������", glngSys, 1255, "1", Array(chk(3)), InStr(mstrPrivs, "����ѡ������") > 0))
    
    strTmp = zlDatabase.GetPara("�������ߵ���0������", glngSys, 1255, "0", Array(lbl(4), cboBody(6)), InStr(mstrPrivs, "����ѡ������") > 0)
    If Val(strTmp) < 0 Or Val(strTmp) > 2 Then strTmp = "0"
    cboBody(6).ListIndex = Val(strTmp)
    
    Me.Show 1, mfrmMain
    ShowPara = mblnOK
    
End Function

Private Sub cboBody_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCanc_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strTmp As String
    
    
    strTmp = cboBody(0).ListIndex & ";" & cboBody(1).ListIndex
    Call zlDatabase.SetPara("�����������߱�־", strTmp, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    strTmp = cboBody(7).ListIndex & ";" & cboBody(8).ListIndex
    Call zlDatabase.SetPara("�����������߱�־(��)", strTmp, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    strTmp = cboBody(2).ListIndex & ";" & cboBody(3).ListIndex
    Call zlDatabase.SetPara("����������ʩ��־", strTmp, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    strTmp = cboBody(4).ListIndex & ";" & cboBody(5).ListIndex
    Call zlDatabase.SetPara("���̾����쳣�߱�־", strTmp, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
 
    Call zlDatabase.SetPara("����ͼ��ʾ����ʱ��", chk(0).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("����ͼģʽ", chk(1).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("��¶�ߵ���ʾλ��", chk(2).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("����ͼ��ʾ������", chk(3).Value, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("�������ߵ���0������", cboBody(6).ListIndex, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    
    mblnOK = True
    Unload Me
End Sub


