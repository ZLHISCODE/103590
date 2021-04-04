VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmOpsStationScheme 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ӵ�ǰ��������������������"
   ClientHeight    =   6270
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9285
   Icon            =   "frmOpsStationScheme.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   1170
      Index           =   0
      Left            =   510
      ScaleHeight     =   1170
      ScaleWidth      =   7515
      TabIndex        =   0
      Top             =   180
      Width           =   7515
      Begin VB.CommandButton cmd 
         Caption         =   "���ɷ���(&S)"
         Height          =   350
         Index           =   1
         Left            =   3900
         TabIndex        =   7
         Top             =   705
         Width           =   1470
      End
      Begin VB.CheckBox chk 
         Caption         =   "��������"
         Height          =   360
         Index           =   2
         Left            =   2400
         TabIndex        =   6
         Top             =   765
         Value           =   1  'Checked
         Width           =   1485
      End
      Begin VB.CheckBox chk 
         Caption         =   "��������"
         Height          =   285
         Index           =   1
         Left            =   1245
         TabIndex        =   5
         Top             =   765
         Value           =   1  'Checked
         Width           =   1485
      End
      Begin VB.CheckBox chk 
         Caption         =   "������ҩ"
         Height          =   285
         Index           =   0
         Left            =   60
         TabIndex        =   4
         Top             =   765
         Value           =   1  'Checked
         Width           =   1485
      End
      Begin VB.CommandButton cmd 
         Caption         =   "ѡ�񷽰�(&O)"
         Height          =   350
         Index           =   0
         Left            =   1590
         TabIndex        =   3
         Top             =   60
         Width           =   1470
      End
      Begin VB.OptionButton opt 
         Caption         =   "����Ϊ�·���"
         Height          =   330
         Index           =   1
         Left            =   0
         TabIndex        =   2
         Top             =   390
         Value           =   -1  'True
         Width           =   1515
      End
      Begin VB.OptionButton opt 
         Caption         =   "�滻ԭ�з���"
         Height          =   330
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   75
         Width           =   1590
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   60
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmOpsStationScheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���������弶��������
'######################################################################################################################

'��������


'��������
Private mfrmMain As Object
Private mstrPrivs As String
Private mblnStartUp As Boolean
Private mblnAllowClose As Boolean
Private mlngTmp As Long
Private mlngKey As Long
Private mstrFindKey As String
Private mblnDataChanged As Boolean
Private mblnNew As Boolean
Private WithEvents mfrmChildSchemeEdit As frmChildSchemeEdit
Attribute mfrmChildSchemeEdit.VB_VarHelpID = -1

Public Function ShowEdit(ByVal frmMain As Object) As Boolean
    
    Set mfrmMain = frmMain
    Me.Show , mfrmMain
    
End Function


Private Sub InitDockPannel()
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 200, 100, DockTopOf, Nothing)
    objPane.Title = "���������б�"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(2, 350, 300, DockBottomOf, Nothing)
    objPane.Title = "��ϸ����"
    objPane.Options = PaneNoCaption
        
    Call DockPannelInit(dkpMain)

End Sub

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim intRow As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String

    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ�ؼ�"

        Call InitDockPannel
        
    '------------------------------------------------------------------------------------------------------------------
    Case "У������"

        '1.У����ϸ����
        '--------------------------------------------------------------------------------------------------------------
        If opt(1).Value Then
            If mfrmChildSchemeEdit.ValidData = False Then Exit Function
        End If
        
        If chk(0).Value = 0 And chk(1).Value = 0 And chk(2).Value = 0 Then
            ShowSimpleMsg "����ѡ��������ҩ,�������ϻ���������!"
            Exit Function
        End If
        
        ExecuteCommand = True

        Exit Function
    '------------------------------------------------------------------------------------------------------------------
    Case "��������"

        '1.������ϸ����
        '--------------------------------------------------------------------------------------------------------------
        mlngTmp = Val(cmd(0).Tag)
        If opt(1).Value Then
            If mfrmChildSchemeEdit.SaveData(rsSQL, mlngTmp) = False Then Exit Function
        End If
                
        strSQL = "zl_���������ο�_Make(" & mfrmMain.GetRecordID & "," & mlngTmp & "," & chk(0).Value & "," & chk(1).Value & "," & chk(2).Value & ")"
        Call SQLRecordAdd(rsSQL, strSQL)

        ExecuteCommand = SQLRecordExecute(rsSQL, Me.Caption)

        Exit Function

    End Select

    ExecuteCommand = True

    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function


Private Sub cmd_Click(Index As Integer)
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset

    
    Select Case Index
    Case 0
        strSQL = "SELECT A.ID,A.����,A.����,A.����,A.˵�� FROM ���������ο� A "
        
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
        If ShowPubSelect(Me, cmd(0), 2, "����,900,0,;����,2400,0,;����,900,0,;˵��,1500,0,", Me.Name & "\��������ѡ��", "����±���ѡ��һ����������", rsData, rs, 8790, 4500, , Val(cmd(0).Tag)) = 1 Then
            If Val(cmd(0).Tag) <> zlCommFun.NVL(rs("ID")) Then
                
                cmd(0).Tag = zlCommFun.NVL(rs("ID"))
                Call mfrmChildSchemeEdit.RefreshData(zlCommFun.NVL(rs("ID").Value, 0))
    
            End If
        End If
     Case 1
        If ExecuteCommand("У������") Then
            If ExecuteCommand("��������") Then
                Unload Me
            End If
        End If
     End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(0).hWnd
    Case 2
        Set mfrmChildSchemeEdit = New frmChildSchemeEdit
        Item.Handle = mfrmChildSchemeEdit.hWnd
        Call mfrmChildSchemeEdit.InitData(Me, False)
    End Select
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    DoEvents

    If ExecuteCommand("��ʼ����") = False Then GoTo errHand

    Call ExecuteCommand("ˢ������")
    
    Call opt_Click(1)
    
    mblnAllowClose = True
    Exit Sub

errHand:
    mblnAllowClose = True
    Unload Me
End Sub

Private Sub Form_Load()
    
    mblnStartUp = True
    mblnAllowClose = False
    
    opt(0).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    opt(1).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    chk(0).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    chk(1).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    chk(2).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    picPane(0).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    
    Call ExecuteCommand("��ʼ�ؼ�")
         
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    Call SetPaneRange(dkpMain, 1, 15, 80, Me.ScaleWidth, 80)
    
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mfrmChildSchemeEdit = Nothing
End Sub

Private Sub opt_Click(Index As Integer)
    
    Select Case Index
    Case 0
        cmd(0).Enabled = True
        cmd(0).Tag = ""
        mfrmChildSchemeEdit.AllowModify = False
        Call mfrmChildSchemeEdit.RefreshData(0)
    Case 1
        cmd(0).Tag = ""
        cmd(0).Enabled = False
        mfrmChildSchemeEdit.AllowModify = True
        Call mfrmChildSchemeEdit.NewData(0)
    End Select
End Sub
