VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEPRFileTimeout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ʱ��Ҫ��"
   ClientHeight    =   4260
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "frmEPRFileTimeout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ListBox lstEquate 
      Enabled         =   0   'False
      Height          =   1530
      Left            =   3510
      Style           =   1  'Checkbox
      TabIndex        =   28
      Top             =   1725
      Width           =   2280
   End
   Begin VB.CheckBox chkMust 
      Caption         =   "����(&M)"
      Height          =   225
      Left            =   4935
      TabIndex        =   5
      Top             =   1080
      Value           =   1  'Checked
      Width           =   945
   End
   Begin VB.PictureBox picCycle 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1020
      Left            =   420
      ScaleHeight     =   1020
      ScaleWidth      =   2625
      TabIndex        =   32
      Top             =   1725
      Width           =   2625
      Begin VB.TextBox txtCycle 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   2
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   24
         Text            =   "0"
         Top             =   690
         Width           =   585
      End
      Begin VB.TextBox txtCycle 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   21
         Text            =   "0"
         Top             =   345
         Width           =   585
      End
      Begin VB.TextBox txtCycle 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   1725
         MaxLength       =   6
         TabIndex        =   18
         Text            =   "0"
         Top             =   0
         Width           =   585
      End
      Begin MSComCtl2.UpDown updCycle 
         Height          =   300
         Index           =   0
         Left            =   2310
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   0
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtCycle(0)"
         BuddyDispid     =   196612
         BuddyIndex      =   0
         OrigLeft        =   2340
         OrigTop         =   15
         OrigRight       =   2580
         OrigBottom      =   315
         Max             =   2400
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updCycle 
         Height          =   300
         Index           =   1
         Left            =   2310
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   345
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtCycle(1)"
         BuddyDispid     =   196612
         BuddyIndex      =   1
         OrigLeft        =   2340
         OrigTop         =   360
         OrigRight       =   2580
         OrigBottom      =   660
         Max             =   2400
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updCycle 
         Height          =   300
         Index           =   2
         Left            =   2310
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   690
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtCycle(2)"
         BuddyDispid     =   196612
         BuddyIndex      =   2
         OrigLeft        =   2340
         OrigTop         =   705
         OrigRight       =   2580
         OrigBottom      =   1005
         Max             =   2400
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblCycle 
         AutoSize        =   -1  'True
         Caption         =   "&3)��Σ������д����"
         Height          =   180
         Index           =   2
         Left            =   0
         TabIndex        =   23
         Top             =   750
         Width           =   1620
      End
      Begin VB.Label lblCycle 
         AutoSize        =   -1  'True
         Caption         =   "&2)���ز�����д����"
         Height          =   180
         Index           =   1
         Left            =   0
         TabIndex        =   20
         Top             =   405
         Width           =   1620
      End
      Begin VB.Label lblCycle 
         AutoSize        =   -1  'True
         Caption         =   "&1)һ�㲡����д����"
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   17
         Top             =   60
         Width           =   1620
      End
   End
   Begin VB.PictureBox picOnce 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1575
      Left            =   420
      ScaleHeight     =   1575
      ScaleWidth      =   2625
      TabIndex        =   31
      Top             =   1725
      Width           =   2625
      Begin VB.TextBox txtTime 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   2
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   15
         Text            =   "0"
         Top             =   915
         Width           =   600
      End
      Begin VB.TextBox txtTime 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   12
         Text            =   "0"
         Top             =   585
         Width           =   600
      End
      Begin VB.TextBox txtTime 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   9
         Text            =   "0"
         Top             =   255
         Width           =   600
      End
      Begin VB.OptionButton optTime 
         Caption         =   "�ڹ涨ʱ�������(&T)"
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Value           =   -1  'True
         Width           =   2025
      End
      Begin VB.OptionButton optTime 
         Caption         =   "���¼�֮ǰ���(&P)"
         Height          =   180
         Index           =   1
         Left            =   0
         TabIndex        =   7
         Top             =   1335
         Width           =   1875
      End
      Begin MSComCtl2.UpDown updTime 
         Height          =   300
         Index           =   0
         Left            =   2220
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtTime(0)"
         BuddyDispid     =   196615
         BuddyIndex      =   0
         OrigLeft        =   2220
         OrigTop         =   495
         OrigRight       =   2460
         OrigBottom      =   795
         Max             =   2400
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updTime 
         Height          =   300
         Index           =   1
         Left            =   2220
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   585
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtTime(1)"
         BuddyDispid     =   196615
         BuddyIndex      =   1
         OrigLeft        =   2220
         OrigTop         =   825
         OrigRight       =   2460
         OrigBottom      =   1125
         Max             =   2400
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updTime 
         Height          =   300
         Index           =   2
         Left            =   2220
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   915
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtTime(2)"
         BuddyDispid     =   196615
         BuddyIndex      =   2
         OrigLeft        =   2235
         OrigTop         =   1155
         OrigRight       =   2475
         OrigBottom      =   1455
         Max             =   2400
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         Caption         =   "&3)�������ʱ��"
         Height          =   180
         Index           =   2
         Left            =   285
         TabIndex        =   14
         Top             =   975
         Width           =   1260
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         Caption         =   "&2)�������ʱ��"
         Height          =   180
         Index           =   1
         Left            =   285
         TabIndex        =   11
         Top             =   645
         Width           =   1260
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         Caption         =   "&1)��д���ʱ��"
         Height          =   180
         Index           =   0
         Left            =   285
         TabIndex        =   8
         Top             =   315
         Width           =   1260
      End
   End
   Begin VB.ComboBox cboOnly 
      Height          =   300
      Left            =   3690
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1035
      Width           =   1185
   End
   Begin VB.ComboBox cboEvent 
      Height          =   300
      Left            =   1740
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1050
      Width           =   1350
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4695
      TabIndex        =   30
      Top             =   3810
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3555
      TabIndex        =   29
      Top             =   3810
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -120
      TabIndex        =   27
      Top             =   3660
      Width           =   6525
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   -60
      TabIndex        =   26
      Top             =   585
      Width           =   6525
   End
   Begin VB.Label lblEquate 
      AutoSize        =   -1  'True
      Caption         =   "�ɵ�ͬ�������ѡ����:"
      Height          =   180
      Left            =   3480
      TabIndex        =   35
      Top             =   1470
      Width           =   2070
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      Caption         =   "˵��:ʱ����СʱΪ��λ,0��ʾ����ȷ��Ҫ��"
      Height          =   180
      Left            =   240
      TabIndex        =   34
      Top             =   3375
      Width           =   3690
   End
   Begin VB.Label lblLimit 
      AutoSize        =   -1  'True
      Caption         =   "ʱ��Ҫ��:"
      Height          =   180
      Left            =   240
      TabIndex        =   33
      Top             =   1470
      Width           =   810
   End
   Begin VB.Label lblBasic 
      AutoSize        =   -1  'True
      Caption         =   "����Ҫ��: �ڲ���                �����"
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   1110
      Width           =   3420
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      Caption         =   "�ļ�����: 001-��Ժ��¼"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   750
      Width           =   1980
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   225
      Picture         =   "frmEPRFileTimeout.frx":058A
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "Ϊ��Ч��������������֤�������ʱ�䣬�ɶԲ����ļ��Ķ�Ӧ�¼������ʱ�޽������á�"
      Height          =   360
      Left            =   780
      TabIndex        =   0
      Top             =   120
      Width           =   5070
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmEPRFileTimeout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mintKind As Integer       '��������
Private mlngFileID As Long        '�����ļ�ID
Private mblnOK As Boolean


Public Function ShowMe(ByVal frmParent As Object, ByVal lngFileID As Long) As Boolean
    '---------------------------------------------------
    '���ܣ��ϼ�������ñ�����ģ����ݲ���������ʾ����
    '---------------------------------------------------
Dim rsTemp As New ADODB.Recordset
Dim lngCount As Long
    mlngFileID = lngFileID
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select ����, ���, ���� From �����ļ��б� Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    With rsTemp
        If .RecordCount = 0 Then MsgBox "�ļ���ʧ(���ܱ������û�ɾ��)��", vbInformation, gstrSysName: Unload Me: Exit Function
        mintKind = !����
        Me.lblFile.Caption = "�ļ�����:   " & !��� & "-" & !����
    End With
    
    '---------------------------------------------------
    Me.cboOnly.AddItem "ѭ����¼": Me.cboOnly.AddItem "��дһ��"
    
    Select Case mintKind
    Case 1  '���ﲡ��
        gstrSQL = "Select ����, ��ǰ����, ѭ������ From ������д�¼� Where ���� = [1] Order By ���"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, 1)
        With rsTemp
            Do While Not .EOF
                Me.cboEvent.AddItem "" & !����
                Me.cboEvent.ItemData(Me.cboEvent.NewIndex) = IIf(IsNull(!��ǰ����), "0", !��ǰ����) & IIf(IsNull(!ѭ������), "0", !ѭ������)
                .MoveNext
            Loop
        End With
        
        gstrSQL = "Select �¼�, ����" & _
                " From ����ʱ��Ҫ�� Where �ļ�id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
        With rsTemp
            If .RecordCount > 0 Then
                For lngCount = 0 To Me.cboEvent.ListCount
                    If !�¼� = Me.cboEvent.List(lngCount) Then Me.cboEvent.ListIndex = lngCount: Exit For
                Next
                Me.chkMust.Value = IIf(IsNull(!����), 0, !����)
            End If
            If Me.cboEvent.ListIndex = -1 Then Me.cboEvent.ListIndex = 0
        End With
        Me.cboOnly.ListIndex = 1: Me.cboOnly.Enabled = False
        
        Me.lblLimit.Caption = Me.lblLimit.Caption & " �ڽ������ξ���ǰ��ɡ�"
        Me.picCycle.Visible = False: Me.picOnce.Visible = False
        Me.lblEquate.Visible = False: Me.lstEquate.Visible = False
        
    
    Case 2, 4       'סԺ������������
        gstrSQL = "Select ����, ��ǰ����, ѭ������ From ������д�¼� Where ���� = [1] Order By ���"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mintKind)
        With rsTemp
            Do While Not .EOF
                Me.cboEvent.AddItem "" & !����
                Me.cboEvent.ItemData(Me.cboEvent.NewIndex) = IIf(IsNull(!��ǰ����), "0", !��ǰ����) & IIf(IsNull(!ѭ������), "0", !ѭ������)
                .MoveNext
            Loop
        End With
        
        gstrSQL = "Select �¼�, ����, Ψһ, ��дʱ��, ����ʱ��, ���ʱ��, һ������, ��������, ��Σ����" & _
                " From ����ʱ��Ҫ�� Where �ļ�id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
        With rsTemp
            If .RecordCount > 0 Then
                Me.chkMust.Value = IIf(IsNull(!����), 0, !����)
                For lngCount = 0 To Me.cboEvent.ListCount
                    If !�¼� = Me.cboEvent.List(lngCount) Then Me.cboEvent.ListIndex = lngCount: Exit For
                Next
                If IIf(IsNull(!Ψһ), 0, !Ψһ) <> 1 Then
                    Me.cboOnly.ListIndex = 0
                Else
                    Me.cboOnly.ListIndex = 1
                End If
                If !��дʱ�� >= 0 Then
                    Me.optTime(0).Value = True
                    Me.txtTime(0).Text = "" & !��дʱ��: Me.updTime(0).Value = Val(Me.txtTime(0).Text)
                    Me.txtTime(1).Text = "" & !����ʱ��: Me.updTime(1).Value = Val(Me.txtTime(1).Text)
                    Me.txtTime(2).Text = "" & !���ʱ��: Me.updTime(2).Value = Val(Me.txtTime(2).Text)
                Else
                    Me.optTime(1).Value = True
                End If
                Me.txtCycle(0).Text = "" & !һ������: Me.updCycle(0).Value = Val(Me.txtCycle(0).Text)
                Me.txtCycle(1).Text = "" & !��������: Me.updCycle(1).Value = Val(Me.txtCycle(1).Text)
                Me.txtCycle(2).Text = "" & !��Σ����: Me.updCycle(2).Value = Val(Me.txtCycle(2).Text)
            End If
            If Me.cboEvent.ListIndex = -1 Then Me.cboEvent.ListIndex = 0
            If Me.cboOnly.ListIndex = -1 Then Me.cboOnly.ListIndex = 1
        End With
        
        '����ļ�����
        Me.lstEquate.Enabled = True
        gstrSQL = "Select l.Id, l.���, l.����, Decode(e.���id, Null, 0, 1) As ���" & _
                " From �����ļ��б� l, ����ʱ��Ҫ�� r, (Select ���id From ���������ϵ Where �ļ�id = [1]) e" & _
                " Where l.Id = r.�ļ�id And l.Id = e.���id(+) And l.���� = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID, mintKind)
        With rsTemp
            Do While Not .EOF
                If lngFileID <> !ID Then
                    Me.lstEquate.AddItem !��� & "-" & !����
                    Me.lstEquate.ItemData(Me.lstEquate.NewIndex) = !ID
                    Me.lstEquate.Selected(Me.lstEquate.NewIndex) = CBool(!���)
                End If
                .MoveNext
            Loop
        End With
        
        '��������жϣ�������סԺ���Ԥ�����ʱ������ʾ�����������ʱ��Ҫ��
        gstrSQL = "Select m.Id " & _
                " From �����ļ��ṹ m, �����ļ��ṹ p " & _
                " Where m.Ԥ�����id = p.Id And p.������� = 3 And m.�ļ�id = [1] And p.�ļ�id Is Null"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
        Me.txtTime(2).Visible = (rsTemp.RecordCount > 0)
        Me.lblTime(2).Visible = (rsTemp.RecordCount > 0)
        Me.updTime(2).Visible = (rsTemp.RecordCount > 0)
    
    Case Else
        MsgBox "�ļ��������", vbInformation, gstrSysName: Unload Me: Exit Function
    End Select
    
    '---------------------------------------------------
    Me.Show vbModal, frmParent
    '---------------------------------------------------
    ShowMe = mblnOK: Unload Me
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cboEvent_Click()
    If mintKind = 1 Then Exit Sub
    
    With Me.cboEvent
        '�����¼�����Ҫ����ǰ����
        If Left(Format(.ItemData(.ListIndex), "00"), 1) = "0" Then
            Me.optTime(0).Value = True: Me.optTime(1).Value = False: Me.optTime(1).Enabled = False
        Else
            Me.optTime(1).Enabled = True
        End If
        '�����¼�����Ҫ��ѭ������
        If Mid(Format(.ItemData(.ListIndex), "00"), 2) = "0" Then
            Me.cboOnly.ListIndex = 1: Me.cboOnly.Enabled = False
        Else
            Me.cboOnly.Enabled = True
        End If
    End With
End Sub

Private Sub cboEvent_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboOnly_Click()
    If Me.cboOnly.ListIndex = 0 Then
        Me.picOnce.Visible = False: Me.picCycle.Visible = True
    Else
        Me.picOnce.Visible = True: Me.picCycle.Visible = False
    End If
    If mintKind = 1 Then
        Me.picOnce.Enabled = False: Me.picCycle.Enabled = False
    Else
        Me.picOnce.Enabled = (Me.cboOnly.ListIndex = 1): Me.picCycle.Enabled = (Me.cboOnly.ListIndex = 0)
    End If
End Sub

Private Sub cboOnly_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkMust_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False: Me.Hide
End Sub

Private Sub cmdOK_Click()
Dim strTemp As String
Dim lngCount As Long
    
    'ʱ����ȷ�Լ��
    If mintKind <> 1 Then
        If Me.cboOnly.ListIndex = 0 Then
            If Val(Me.txtCycle(0).Text) < Val(Me.txtCycle(1).Text) Then MsgBox "һ�㲡����д���ڲ���С�ڲ��ز�����д���ڣ�", vbInformation, gstrSysName: Exit Sub
            If Val(Me.txtCycle(1).Text) < Val(Me.txtCycle(2).Text) Then MsgBox "���ز�����д���ڲ���С�ڲ�Σ������д���ڣ�", vbInformation, gstrSysName: Exit Sub
        ElseIf Me.optTime(0).Value = True Then
            If Val(Me.txtTime(0).Text) > Val(Me.txtTime(1).Text) And Val(Me.txtTime(1).Text) <> 0 Then
                MsgBox "��ǩʱ�޲���С����дʱ�ޣ�", vbInformation, gstrSysName: Exit Sub
            End If
            If Me.txtTime(2).Visible Then
                If Val(Me.txtTime(1).Text) > Val(Me.txtTime(2).Text) And Val(Me.txtTime(2).Text) <> 0 Then
                    MsgBox "������ϲ���С����ǩʱ�ޣ�", vbInformation, gstrSysName: Exit Sub
                End If
            End If
        End If
    End If
    
    '����SQL��֯
    If mintKind = 1 Then
        gstrSQL = "Zl_�����ļ��б�_Timeout(" & mlngFileID & ",'" & Me.cboEvent.Text & "',1," & Me.chkMust.Value & ",0,0,0,0,0,0,null)"
    Else
        gstrSQL = mlngFileID & ",'" & Me.cboEvent.Text & "'," & Me.cboOnly.ListIndex & "," & Me.chkMust.Value
        If Me.cboOnly.ListIndex = 0 Then
            gstrSQL = gstrSQL & ",0,0,0," & Val(Me.txtCycle(0).Text) & "," & Val(Me.txtCycle(1).Text) & "," & Val(Me.txtCycle(2).Text)
        ElseIf Me.optTime(0).Value = False Then
            gstrSQL = gstrSQL & ",-1,0,0,0,0,0"
        Else
            gstrSQL = gstrSQL & "," & Val(Me.txtTime(0).Text) & "," & Val(Me.txtTime(1).Text)
            If Me.txtTime(2).Visible Then
                gstrSQL = gstrSQL & "," & Val(Me.txtTime(2).Text) & ",0,0,0"
            Else
                gstrSQL = gstrSQL & ",0,0,0,0"
            End If
        End If
        strTemp = ""
        For lngCount = 0 To Me.lstEquate.ListCount - 1
            If Me.lstEquate.Selected(lngCount) Then strTemp = strTemp & ";" & Me.lstEquate.ItemData(lngCount)
        Next
        If strTemp <> "" Then strTemp = Mid(strTemp, 2)
        gstrSQL = "Zl_�����ļ��б�_Timeout(" & gstrSQL & ",'" & strTemp & "')"
    End If
    
    Err = 0: On Error GoTo errHand
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    mblnOK = True: Me.Hide: Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lstEquate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub optTime_Click(Index As Integer)
    Me.txtTime(0).Enabled = Me.optTime(0).Value: Me.updTime(0).Enabled = Me.txtTime(0).Enabled
    Me.txtTime(1).Enabled = Me.optTime(0).Value: Me.updTime(1).Enabled = Me.txtTime(1).Enabled
    Me.txtTime(2).Enabled = Me.optTime(0).Value: Me.updTime(2).Enabled = Me.txtTime(2).Enabled
End Sub

Private Sub optTime_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtCycle_Change(Index As Integer)
    If Val(Me.txtCycle(Index).Text) > Me.updCycle(Index).Max Or Val(Me.txtCycle(Index).Text) < Me.updCycle(Index).Min Then
        Me.txtCycle(Index).Text = Me.updCycle(Index).Min
    End If
End Sub

Private Sub txtCycle_GotFocus(Index As Integer)
    Me.txtCycle(Index).SelStart = 0: Me.txtCycle(Index).SelLength = 100
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtCycle_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("1234567890" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtTime_Change(Index As Integer)
    If Val(Me.txtTime(Index).Text) > Me.updTime(Index).Max Or Val(Me.txtTime(Index).Text) < Me.updTime(Index).Min Then
        Me.txtTime(Index).Text = Me.updTime(Index).Min
    End If
End Sub

Private Sub txtTime_GotFocus(Index As Integer)
    Me.txtTime(Index).SelStart = 0: Me.txtTime(Index).SelLength = 100
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtTime_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("1234567890" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

