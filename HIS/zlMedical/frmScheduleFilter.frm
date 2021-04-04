VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmScheduleFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   2955
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6405
   Icon            =   "frmScheduleFilter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5205
      TabIndex        =   16
      Top             =   255
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5205
      TabIndex        =   17
      Top             =   735
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "��������"
      Height          =   2730
      Left            =   105
      TabIndex        =   18
      Top             =   105
      Width           =   4980
      Begin VB.OptionButton opt 
         Caption         =   "����(&7)"
         Height          =   240
         Index           =   2
         Left            =   3615
         TabIndex        =   14
         Top             =   1935
         Width           =   930
      End
      Begin VB.OptionButton opt 
         Caption         =   "����(&6)"
         Height          =   240
         Index           =   1
         Left            =   2475
         TabIndex        =   13
         Top             =   1935
         Width           =   1095
      End
      Begin VB.OptionButton opt 
         Caption         =   "����(&5)"
         Height          =   240
         Index           =   0
         Left            =   1380
         TabIndex        =   12
         Top             =   1935
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.PictureBox picCmd 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3915
         ScaleHeight     =   255
         ScaleWidth      =   675
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1515
         Width           =   675
         Begin VB.CommandButton cmdClear 
            Caption         =   "X"
            Height          =   240
            Index           =   0
            Left            =   285
            TabIndex        =   11
            Top             =   0
            Width           =   255
         End
         Begin VB.CommandButton cmd 
            Caption         =   "&P"
            Height          =   240
            Index           =   0
            Left            =   15
            TabIndex        =   10
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "������ȷ��(&5)"
         Height          =   195
         Index           =   0
         Left            =   3105
         TabIndex        =   15
         Top             =   2370
         Width           =   1530
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   1410
         TabIndex        =   7
         Top             =   1095
         Width           =   3075
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1410
         TabIndex        =   5
         Top             =   720
         Width           =   3075
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   1410
         TabIndex        =   1
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   74973187
         CurrentDate     =   38357
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   1
         Left            =   3150
         TabIndex        =   3
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   74973187
         CurrentDate     =   38357
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1410
         TabIndex        =   9
         Top             =   1485
         Width           =   3075
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "��������(&4)"
         Height          =   180
         Index           =   7
         Left            =   270
         TabIndex        =   8
         Top             =   1560
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   4
         Left            =   2835
         TabIndex        =   2
         Top             =   420
         Width           =   180
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "��쵥(&3)"
         Height          =   180
         Index           =   3
         Left            =   450
         TabIndex        =   6
         Top             =   1170
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "ԤԼ��(&2)"
         Height          =   180
         Index           =   1
         Left            =   450
         TabIndex        =   4
         Top             =   810
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "ԤԼʱ��(&1)"
         Height          =   180
         Index           =   0
         Left            =   270
         TabIndex        =   0
         Top             =   435
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmScheduleFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnOK As Boolean
Private mfrmMain As Object
Private mstrCondition As String

Private Type Items
    ���� As String
End Type

Private usrSaveGroup As Items

Public Function ShowFilter(ByVal frmMain As Object, ByRef strCondition As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '------------------------------------------------------------------------------------------------------------------
    mblnOK = False
            
    Set mfrmMain = frmMain
    mstrCondition = strCondition
            
    If ReadCondition = False Then Exit Function
            
    usrSaveGroup.���� = txt(1).Text
    txt(1).Tag = ""
    
    Me.Show 1, mfrmMain
    
    If mblnOK Then strCondition = mstrCondition
    
    ShowFilter = mblnOK
    
End Function

Private Function ReadCondition() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim varCondition As Variant
    
    On Error GoTo errHand
    
    '��'Ϊ�ָ���
    '�洢��ʽ:��ʼʱ��'����ʱ��'ԤԼ��'����'ԤԼ����'ԤԼ����id'����ȷ��
    varCondition = Split(mstrCondition, "'")
    
    dtp(0).Value = Format(varCondition(0), dtp(0).CustomFormat)
    dtp(1).Value = Format(varCondition(1), dtp(1).CustomFormat)
    txt(0).Text = varCondition(2)
    txt(2).Text = varCondition(3)
    txt(1).Text = varCondition(4)
    cmd(0).Tag = Val(varCondition(5))
    chk(0).Value = Val(varCondition(6))
    opt(Val(varCondition(7))).Value = True
    
    ReadCondition = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub chk_Click(Index As Integer)
    If Index = 1 Then
        txt(1).Enabled = (chk(Index).Value = 1)
        cmd(0).Enabled = txt(1).Enabled
    End If
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim rsData As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    gstrSQL = GetPublicSQL(SQL.�������ѡ��)
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    If ShowTxtSelect(Me, txt(1), "����,900,0,1;����,1500,0,1;����,900,0,1;��ַ,3000,0,1", Me.Name & "\�������ѡ��", "�����±���ѡ��һ������/��λ��", rsData, rs, 8790, 5100) Then
    
        txt(1).Text = zlCommFun.NVL(rs("����").Value)
        cmd(Index).Tag = zlCommFun.NVL(rs("ID").Value, 0)
        usrSaveGroup.���� = txt(1).Text
                
    End If

    txt(1).SetFocus
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click(Index As Integer)
    cmd(0).Tag = ""
    txt(1).Text = ""
    txt(1).Tag = ""
End Sub

Private Sub cmdOK_Click()
    Dim intYear As Integer
    Dim strYear As String
    
    '�Զ����뵥�ݺ�
    If (UCase(Left(txt(2).Text, 1)) < "A" Or UCase(Left(txt(2).Text, 1)) > "Z") And Trim(txt(2).Text) <> "" Then
        intYear = Format(zlDatabase.Currentdate, "YYYY") - 1990
        strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
        txt(2).Text = strYear & Right("0000000" & txt(2).Text, 7)
    End If
    
    mstrCondition = ""
    mstrCondition = Format(dtp(0).Value, dtp(0).CustomFormat) & "'" & Format(dtp(1).Value, dtp(1).CustomFormat)
    mstrCondition = mstrCondition & "'" & Trim(txt(0).Text)
    mstrCondition = mstrCondition & "'" & Trim(txt(2).Text)
    If Val(cmd(0).Tag) > 0 Then
        mstrCondition = mstrCondition & "'" & Trim(txt(1).Text)
        mstrCondition = mstrCondition & "'" & Val(cmd(0).Tag)
    Else
        mstrCondition = mstrCondition & "''"
    End If
    mstrCondition = mstrCondition & "'" & chk(0).Value
    If opt(0).Value Then
        mstrCondition = mstrCondition & "'0"
    ElseIf opt(1).Value Then
        mstrCondition = mstrCondition & "'1"
    Else
        mstrCondition = mstrCondition & "'2"
    End If
    
    mblnOK = True
    
    Unload Me
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub opt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt_Change(Index As Integer)
    If Index = 1 Then
        txt(Index).Tag = "Changed"
        cmd(0).Tag = ""
    End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
    
    Select Case Index
    Case 0, 1
        zlCommFun.OpenIme True
    End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        
        If txt(Index).Tag = "Changed" And Index = 1 Then
            gstrSQL = GetPublicSQL(SQL.�������ѡ��)
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "%" & UCase(txt(Index).Text) & "%")
            
            If ShowTxtFilter(Me, txt(Index), "����,1800,0,0;����,900,0,0;����,900,0,0;��ϵ��,900,0,0;�绰,1200,0,0", Me.Name & "\�������ѡ��", "�������ѡ��һ�����嵥λ", rsData, rs) Then
                
                txt(1).Text = zlCommFun.NVL(rs("����"))
                cmd(0).Tag = zlCommFun.NVL(rs("ID"))
                
                usrSaveGroup.���� = txt(1).Text
            Else
                txt(Index).Text = usrSaveGroup.����
                Exit Sub
            End If
        End If
        
        zlCommFun.PressKey vbKeyTab
        
        Select Case Index
        Case 1
            zlCommFun.PressKey vbKeyTab
            zlCommFun.PressKey vbKeyTab
        End Select
        
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        If Index = 2 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 0, 1
        zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    
    If Index = 1 Then
        If txt(Index).Tag = "Changed" Then
            txt(Index).Text = usrSaveGroup.����
            txt(Index).Tag = ""
        End If
    End If
End Sub
