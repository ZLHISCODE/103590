VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmCISAduitPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   Icon            =   "frmCISAduitPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picPane 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2430
      Index           =   0
      Left            =   270
      ScaleHeight     =   2430
      ScaleWidth      =   4470
      TabIndex        =   10
      Top             =   555
      Width           =   4470
      Begin VB.ComboBox cbo 
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   4
         Left            =   2445
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2055
         Width           =   1920
      End
      Begin VB.ComboBox cbo 
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   3
         Left            =   2445
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1665
         Width           =   1920
      End
      Begin VB.Frame fra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   2
         Left            =   930
         TabIndex        =   11
         Top             =   165
         Width           =   4815
      End
      Begin VB.ComboBox cbo 
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   0
         Left            =   2445
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   870
         Width           =   1920
      End
      Begin VB.ComboBox cbo 
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   1
         Left            =   2445
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1260
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.ComboBox cbo 
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   2
         Left            =   2445
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1260
         Width           =   1920
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&4.סԺҽ����ӡ"
         Height          =   180
         Index           =   9
         Left            =   1020
         TabIndex        =   16
         Top             =   2100
         Width           =   1260
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&3.����ʱ��"
         Height          =   180
         Index           =   7
         Left            =   1020
         TabIndex        =   6
         Top             =   1710
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&1.�ύʱ��"
         Height          =   180
         Index           =   1
         Left            =   1020
         TabIndex        =   0
         Top             =   930
         Width           =   900
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   360
         Picture         =   "frmCISAduitPara.frx":000C
         Top             =   390
         Width           =   480
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "��ѯ���Ӳ�������ȱʡʱ�䷶Χ��"
         Height          =   405
         Left            =   960
         TabIndex        =   13
         Top             =   570
         Width           =   4065
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ȱʡʱ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   14
         Left            =   195
         TabIndex        =   12
         Top             =   150
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&2.��Ժ����"
         Height          =   180
         Index           =   4
         Left            =   1020
         TabIndex        =   4
         Top             =   1305
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&2.�鵵����"
         Height          =   180
         Index           =   0
         Left            =   1020
         TabIndex        =   2
         Top             =   1305
         Visible         =   0   'False
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3705
      TabIndex        =   9
      Top             =   3135
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2490
      TabIndex        =   8
      Top             =   3135
      Width           =   1100
   End
   Begin XtremeSuiteControls.TabControl tbc 
      Height          =   2925
      Left            =   135
      TabIndex        =   14
      Top             =   90
      Width           =   4680
      _Version        =   589884
      _ExtentX        =   8255
      _ExtentY        =   5159
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmCISAduitPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################
Private mblnOK As Boolean
Private mfrmMain As Object
Private mstrPrivs As String

'######################################################################################################################

Public Function ShowEdit(ByVal frmMain As Object, ByVal strPrivs As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    mblnOK = False
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
    
    If ExecuteCommand("��ʼ����") = False Then Exit Function
    If ExecuteCommand("��ȡ����") = False Then Exit Function
        
    DataChanged = False
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function ExecuteCommand(ParamArray varCmd() As Variant) As Boolean
    Dim intLoop As Integer
    Dim intCount As Integer
    Dim intCol As Integer
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    Dim varTmp As Variant
    Dim varAry As Variant
    Dim blnAllowModify As Boolean

    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    For intLoop = 0 To UBound(varCmd)
        Select Case varCmd(intLoop)
        '--------------------------------------------------------------------------------------------------------------
        Case "��ʼ����"
            With tbc
                With .PaintManager
                    .Appearance = xtpTabAppearancePropertyPage2003
                    .BoldSelected = True
                    .COLOR = xtpTabColorDefault
                    .ColorSet.ButtonSelected = COLOR.��ɫ
                    .ShowIcons = True
                End With
                
                .InsertItem 0, "���� ", picPane(0).hWnd, 0
                .Item(0).Selected = True
            End With
            
            For intCount = 0 To 3
                With cbo(intCount)
                    .Clear
                    .AddItem "��  ��"
                    .AddItem "��  ��"
                    .AddItem "��  ��"
                    .AddItem "��  ��"
                    .AddItem "��  ��"
                    .AddItem "������"
                    .AddItem "��  ��"
                    .AddItem "ǰ����"
                    .AddItem "ǰһ��"
                    .AddItem "ǰ����"
                    .AddItem "ǰһ��"
                    .AddItem "ǰ����"
                    .AddItem "ǰ����"
                    .AddItem "ǰ����"
                    .AddItem "ǰһ��"
                    .AddItem "ǰ����"
                End With
            Next
            
            With cbo(4)
                .Clear
                .AddItem "����ҽ����"
                .AddItem "����ҽ����"
            End With
            
        Case "��ȡ����"
            
            On Error Resume Next
            cbo(0).Text = zlDatabase.GetPara("���ȱʡ��Χ", ParamInfo.ϵͳ��, mfrmMain.ģ���, "��  ��", Array(cbo(0)), IsPrivs(mstrPrivs, "��������"))
            cbo(1).Text = zlDatabase.GetPara("�鵵ȱʡ��Χ", ParamInfo.ϵͳ��, mfrmMain.ģ���, "��  ��", Array(cbo(1)), IsPrivs(mstrPrivs, "��������"))
            cbo(2).Text = zlDatabase.GetPara("��Ժȱʡ��Χ", ParamInfo.ϵͳ��, mfrmMain.ģ���, "��  ��", Array(cbo(2)), IsPrivs(mstrPrivs, "��������"))
            cbo(3).Text = zlDatabase.GetPara("ҽ��ȱʡ��Χ", ParamInfo.ϵͳ��, mfrmMain.ģ���, "��  ��", Array(cbo(3)), IsPrivs(mstrPrivs, "��������"))
            cbo(4).Text = zlDatabase.GetPara("סԺҽ����ӡ", ParamInfo.ϵͳ��, mfrmMain.ģ���, "����ҽ����", Array(cbo(4)), IsPrivs(mstrPrivs, "��������"))
            On Error GoTo errHand
            
            If cbo(0).ListCount > 0 And cbo(0).ListIndex = -1 Then cbo(0).ListIndex = 0
            If cbo(1).ListCount > 0 And cbo(1).ListIndex = -1 Then cbo(1).ListIndex = 0
            If cbo(2).ListCount > 0 And cbo(2).ListIndex = -1 Then cbo(2).ListIndex = 0
            If cbo(3).ListCount > 0 And cbo(3).ListIndex = -1 Then cbo(3).ListIndex = 0
            If cbo(4).ListCount > 0 And cbo(4).ListIndex = -1 Then cbo(4).ListIndex = 0
            
        Case "��������"
            
            Call SetPara("���ȱʡ��Χ", cbo(0).Text, mfrmMain.ģ���, IsPrivs(mstrPrivs, "��������"))
            Call SetPara("�鵵ȱʡ��Χ", cbo(1).Text, mfrmMain.ģ���, IsPrivs(mstrPrivs, "��������"))
            Call SetPara("��Ժȱʡ��Χ", cbo(2).Text, mfrmMain.ģ���, IsPrivs(mstrPrivs, "��������"))
            Call SetPara("ҽ��ȱʡ��Χ", cbo(3).Text, mfrmMain.ģ���, IsPrivs(mstrPrivs, "��������"))
            Call SetPara("סԺҽ����ӡ", cbo(4).Text, mfrmMain.ģ���, IsPrivs(mstrPrivs, "��������"))
        End Select
    Next

    ExecuteCommand = True

    Exit Function
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Property Let DataChanged(ByVal blnData As Boolean)
    cmdOK.Tag = IIf(blnData, "Changed", "")
End Property

Private Property Get DataChanged() As Boolean
    DataChanged = (cmdOK.Tag = "Changed")
End Property

'######################################################################################################################

Private Sub cbo_Click(Index As Integer)
    DataChanged = True
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOk_Click()
    
    If DataChanged Then
        If ExecuteCommand("��������") Then
            
            DataChanged = False
            
            mblnOK = True
        Else
            Exit Sub
        End If
    End If
    
    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If DataChanged Then
        Cancel = (MsgBox("�������޸ĵĲ������뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, ParamInfo.ϵͳ����) = vbNo)
    End If
End Sub
