VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmCISBorrowPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   Icon            =   "frmCISBorrowPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   90
      TabIndex        =   2
      Top             =   1965
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3720
      TabIndex        =   1
      Top             =   1965
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4935
      TabIndex        =   0
      Top             =   1965
      Width           =   1100
   End
   Begin VB.PictureBox picPane 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1425
      Index           =   0
      Left            =   255
      ScaleHeight     =   1425
      ScaleWidth      =   5880
      TabIndex        =   3
      Top             =   330
      Width           =   5880
      Begin VB.ComboBox cbo 
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   0
         Left            =   2460
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   870
         Width           =   1920
      End
      Begin VB.Frame fra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   105
         Index           =   2
         Left            =   930
         TabIndex        =   5
         Top             =   165
         Width           =   4815
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
         TabIndex        =   9
         Top             =   150
         Width           =   720
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "��ѯ���Ӳ������������ȱʡʱ�䷶Χ��"
         Height          =   405
         Left            =   1035
         TabIndex        =   8
         Top             =   555
         Width           =   4065
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   360
         Picture         =   "frmCISBorrowPara.frx":000C
         Top             =   390
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȱʡ��Χ(&1)"
         Height          =   180
         Index           =   1
         Left            =   1380
         TabIndex        =   7
         Top             =   930
         Width           =   990
      End
   End
   Begin XtremeSuiteControls.TabControl tbc 
      Height          =   1830
      Left            =   105
      TabIndex        =   4
      Top             =   30
      Width           =   5970
      _Version        =   589884
      _ExtentX        =   10530
      _ExtentY        =   3228
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmCISBorrowPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################
Private mblnOK As Boolean
Private mfrmMain As Object
Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1
Private mstrPrivs As String
Private mblnBorrowAccount As Boolean '��������¼�����ԭ��

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
    '--------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ�
    '--------------------------------------------------------------------------------------------------------------
    Dim intLoop As Integer
    Dim intCount As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
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
            

            cbo(0).Clear
            cbo(0).AddItem "��  ��"
            cbo(0).AddItem "��  ��"
            cbo(0).AddItem "��  ��"
            cbo(0).AddItem "��  ��"
            cbo(0).AddItem "��  ��"
            cbo(0).AddItem "������"
            cbo(0).AddItem "��  ��"
            cbo(0).AddItem "ǰ����"
            cbo(0).AddItem "ǰһ��"
            cbo(0).AddItem "ǰ����"
            cbo(0).AddItem "ǰһ��"
            cbo(0).AddItem "ǰ����"
            cbo(0).AddItem "ǰ����"
            cbo(0).AddItem "ǰ����"
            cbo(0).AddItem "ǰһ��"
            cbo(0).AddItem "ǰ����"
            
        Case "��ȡ����"
            
            On Error Resume Next
            cbo(0).Text = zlDatabase.GetPara("�Ǽ�ȱʡ��Χ", ParamInfo.ϵͳ��, mfrmMain.ģ���, "��  ��", Array(cbo(0)), IsPrivs(mstrPrivs, "��������"))
            On Error GoTo errHand
            
            If cbo(0).ListCount > 0 And cbo(0).ListIndex = -1 Then cbo(0).ListIndex = 0
            
        Case "��������"
            
            Call SetPara("�Ǽ�ȱʡ��Χ", cbo(0).Text, mfrmMain.ģ���, IsPrivs(mstrPrivs, "��������"))
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

Private Sub chkBorrowAccount_Click()
    DataChanged = True
End Sub

Private Sub chkBorrowReason_Click()
    DataChanged = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
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
    
    Set mclsVsf = Nothing
End Sub
