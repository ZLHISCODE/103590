VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmStuffQueryFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���Ҳ���"
   ClientHeight    =   3165
   ClientLeft      =   3135
   ClientTop       =   4320
   ClientWidth     =   5985
   Icon            =   "FrmStuffQueryFind.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox Pic���� 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   -30
      ScaleHeight     =   3135
      ScaleWidth      =   6135
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   6135
      Begin VB.Frame fra 
         Height          =   75
         Index           =   1
         Left            =   0
         TabIndex        =   19
         Top             =   2565
         Width           =   6075
      End
      Begin VB.Frame fra 
         Height          =   45
         Index           =   0
         Left            =   75
         TabIndex        =   18
         Top             =   645
         Width           =   5925
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
         Height          =   1575
         Left            =   585
         TabIndex        =   16
         Top             =   3240
         Visible         =   0   'False
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   2778
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   32768
         AllowBigSelection=   0   'False
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton CmdHelp 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   135
         Picture         =   "FrmStuffQueryFind.frx":020A
         TabIndex        =   17
         Top             =   2745
         Width           =   1100
      End
      Begin VB.CommandButton CmdSelect 
         Caption         =   "��"
         Height          =   300
         Left            =   5475
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2130
         Width           =   255
      End
      Begin VB.TextBox TxtSelect���� 
         Height          =   300
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   11
         Top             =   2115
         Width           =   4520
      End
      Begin VB.CommandButton Cmd���� 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   3390
         Picture         =   "FrmStuffQueryFind.frx":0354
         TabIndex        =   13
         Top             =   2745
         Width           =   1100
      End
      Begin VB.CommandButton Cmd���� 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   4800
         Picture         =   "FrmStuffQueryFind.frx":049E
         TabIndex        =   14
         Top             =   2745
         Width           =   1100
      End
      Begin VB.TextBox Txt���� 
         Height          =   300
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   3
         Top             =   1290
         Width           =   1875
      End
      Begin VB.TextBox Txt���ϱ��� 
         Height          =   300
         Left            =   1200
         TabIndex        =   1
         Top             =   840
         Width           =   1875
      End
      Begin VB.TextBox Txt���� 
         Height          =   300
         Left            =   3840
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1290
         Width           =   1875
      End
      Begin VB.TextBox txt��� 
         Height          =   300
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1710
         Width           =   1875
      End
      Begin VB.TextBox Txt���� 
         Height          =   300
         Left            =   3840
         MaxLength       =   30
         TabIndex        =   9
         Top             =   1740
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label lbl 
         Caption         =   "��������������ָ�����ϵĿ��,���ͬʱ���ö���,������֮�����ҵĹ�ϵ."
         Height          =   345
         Left            =   915
         TabIndex        =   20
         Top             =   255
         Width           =   4935
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   225
         Picture         =   "FrmStuffQueryFind.frx":05E8
         Top             =   60
         Width           =   480
      End
      Begin VB.Label lblָ������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ָ������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   10
         Top             =   2175
         Width           =   720
      End
      Begin VB.Label Lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3390
         TabIndex        =   8
         Top             =   1800
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Lbl��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   720
         TabIndex        =   6
         Top             =   1770
         Width           =   360
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3390
         TabIndex        =   4
         Top             =   1350
         Width           =   360
      End
      Begin VB.Label Lbl���ϱ��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   720
         TabIndex        =   0
         Top             =   900
         Width           =   360
      End
      Begin VB.Label Lblͨ������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   720
         TabIndex        =   2
         Top             =   1350
         Width           =   360
      End
   End
End
Attribute VB_Name = "FrmStuffQueryFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrTemp As String
Public mstrBit As Byte '�ó�����ҵ�ƥ�䷽ʽ
Dim mrsTemp As ADODB.Recordset
Public mstrOthers As Variant   '0-����,1-����,2-����,3-���,4-����,5-ָ������

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub cmdSelect_Click()
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select ����,����,���� From ����������  where (վ��=[1] or վ�� is null) Order By ���� "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "-����������", gstrNodeNo)
        
    With rsTemp
        If .EOF Then
            MsgBox "���ʼ�����������̣��ֵ������", vbInformation, gstrSysName
             Me.TxtSelect����.SetFocus: Exit Sub
        End If
                
        If .RecordCount > 1 Then
            Set mshSelect.Recordset = rsTemp
            With mshSelect
                .Top = TxtSelect����.Top - .Height
                .Left = TxtSelect����.Left
                .Visible = True
                .SetFocus
                .ColWidth(0) = 800
                .ColWidth(1) = 800
                .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                .Row = 1
                .Col = 0
                .ColSel = .Cols - 1
                .ZOrder
                Exit Sub
                
            End With
        Else
            TxtSelect���� = IIf(IsNull(!����), "", !����)
            TxtSelect����.Tag = 1
            Cmd����.SetFocus
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Cmd����_Click()
    
    '0-����,1-����,2-����,3-���,4-����,5-ָ������
    
    mstrOthers(4) = Trim(Txt����.Text)
    
    '����:[1]�ⷿ,[2]-���� ,[3]-����,[4]-����,[5]-���,[6]-����,[7]-ָ������
    If LTrim(txt����) = "" And LTrim(Txt���ϱ���) = "" And LTrim(Txt����) = "" And LTrim(txt���) = "" And LTrim(TxtSelect����) = "" Then MsgBox "����������һ����Ϣ��", vbInformation, gstrSysName
    
    mstrTemp = ""
    If LTrim(txt����) <> "" Then
        mstrTemp = "Q.���� like [3]"
        mstrOthers(1) = IIf(mstrBit = "0", "%", "") & LTrim(txt����) & "%"
    End If
    
    If LTrim(Txt���ϱ���) <> "" Then
        If LTrim(mstrTemp) = "" Then
            mstrTemp = "Q.���� like [2] "
            mstrOthers(0) = IIf(mstrBit = "0", "%", "") & UCase(LTrim(Txt���ϱ���)) & "%"
        Else
            mstrTemp = mstrTemp & " And Q.���� like [2] "
            mstrOthers(0) = IIf(mstrBit = "0", "%", "") & UCase(LTrim(Txt���ϱ���)) & "%"
        End If
    End If
    
    If LTrim(Txt����) <> "" Then
        If LTrim(mstrTemp) = "" Then
            mstrTemp = " M.����id in (Select �շ�ϸĿID from �շ���Ŀ����  where ���� like [4] )"
            mstrOthers(2) = IIf(mstrBit = "0", "%", "") & UCase(LTrim(Txt����)) & "%"
               
        Else
            mstrTemp = mstrTemp & " And  M.����id in (Select �շ�ϸĿID from �շ���Ŀ����  where ���� like [4] )"
            mstrOthers(2) = IIf(mstrBit = "0", "%", "") & UCase(LTrim(Txt����)) & "%"
        End If
    End If
    
    If LTrim(txt���) <> "" Then
        mstrOthers(3) = IIf(mstrBit = "0", "%", "") & UCase(LTrim(txt���)) & "%"
        If LTrim(mstrTemp) = "" Then
            mstrTemp = " upper(Q.���) like [5] "
        Else
            mstrTemp = mstrTemp & " And upper(Q.���) like [5] "
        End If
    End If
    
    If LTrim(TxtSelect����) <> "" Then
        mstrOthers(5) = IIf(mstrBit = "0", "%", "") & UCase(LTrim(TxtSelect����)) & "%"
            
        If LTrim(mstrTemp) = "" Then
        
            mstrTemp = "Upper(Q.����) like [7] "
        Else
            mstrTemp = mstrTemp & " And upper(Q.����) like [7] "
        End If
    End If
    Me.Hide
End Sub

Private Sub Cmd����_Click()
    mstrTemp = ""
    Me.Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = vbKeyEscape Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strOthers(0 To 6) As String
    For i = 0 To 6
        strOthers(i) = ""
    Next
    mstrOthers = strOthers
    mstrBit = gstrMatchMethod
End Sub


Private Sub Pic����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        OS.PressKey (vbKeyTab)
    End If
End Sub

Private Sub TxtSelect����_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        
        If Trim(TxtSelect����) = "" Then Exit Sub
        TxtSelect���� = UCase(TxtSelect����)
    
        Dim rsTemp As New ADODB.Recordset
        
        On Error GoTo ErrHandle
        gstrSQL = "" & _
            "   Select ����,����,���� " & _
            "   From ���������� " & _
            "   Where (���� like [1] or ���� like upper([1]) or  ���� like upper([1])) And (վ��=[1] or վ�� is null) "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����������", IIf(gstrMatchMethod = "0", "%", "") & TxtSelect���� & "%", gstrNodeNo)
        
        With rsTemp
            
            If .EOF Then
                MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
                KeyCode = 0
                Exit Sub
            End If
            If .RecordCount > 1 Then
                
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = TxtSelect����.Top - .Height
                    .Left = TxtSelect����.Left
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 1000
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    .ZOrder
                    Exit Sub
                    
                End With
            Else
                TxtSelect���� = IIf(IsNull(!����), "", !����)
                TxtSelect����.Tag = 1
                Cmd����.SetFocus
            End If
        End With
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub TxtSelect����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    OS.PressKey (vbKeyTab)
End Sub

Private Sub txt���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    OS.PressKey (vbKeyTab)
End Sub

Private Sub Txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    OS.PressKey (vbKeyTab)

End Sub

Private Sub Txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    OS.PressKey (vbKeyTab)
End Sub

Private Sub Txt���ϱ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    OS.PressKey (vbKeyTab)
End Sub


Private Sub mshSelect_DblClick()
    mshSelect_KeyPress 13
End Sub

Private Sub mshSelect_KeyPress(KeyAscii As Integer)
    With mshSelect
        If KeyAscii = 13 Then
            TxtSelect����.Text = .TextMatrix(.Row, 1)
            TxtSelect����.Tag = 1
            Cmd����.SetFocus
            .Visible = False
            Exit Sub
        End If
    End With
End Sub

Private Sub mshSelect_LostFocus()
    mshSelect.Visible = False
End Sub

