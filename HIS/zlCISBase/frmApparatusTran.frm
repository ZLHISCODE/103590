VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmApparatusTran 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ת������"
   ClientHeight    =   3060
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5415
   Icon            =   "frmApparatusTran.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3855
      TabIndex        =   9
      Top             =   2550
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2625
      TabIndex        =   8
      Top             =   2550
      Width           =   1100
   End
   Begin VB.ComboBox cboת������ 
      Enabled         =   0   'False
      Height          =   300
      ItemData        =   "frmApparatusTran.frx":058A
      Left            =   2370
      List            =   "frmApparatusTran.frx":058C
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1530
      Width           =   2280
   End
   Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
      Height          =   300
      Left            =   2370
      TabIndex        =   4
      Top             =   1185
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "yyyy-MM-dd hh:mm"
      Format          =   69074947
      CurrentDate     =   39062
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -30
      TabIndex        =   3
      Top             =   2385
      Width           =   5700
   End
   Begin VB.OptionButton optTran 
      Caption         =   "ȡ������ת��(&2)"
      Height          =   180
      Index           =   1
      Left            =   1110
      TabIndex        =   2
      Top             =   2025
      Value           =   -1  'True
      Width           =   2670
   End
   Begin VB.OptionButton optTran 
      Caption         =   "ת����������(&1)"
      Height          =   180
      Index           =   0
      Left            =   1110
      TabIndex        =   1
      Top             =   900
      Width           =   2670
   End
   Begin VB.Label lblת������ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ת������(&M)"
      Height          =   180
      Left            =   1350
      TabIndex        =   7
      Top             =   1590
      Width           =   990
   End
   Begin VB.Label lbl��ʼʱ�� 
      AutoSize        =   -1  'True
      Caption         =   "��ʼʱ��(&T)"
      Height          =   180
      Left            =   1350
      TabIndex        =   5
      Top             =   1245
      Width           =   990
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   180
      Picture         =   "frmApparatusTran.frx":058E
      Top             =   165
      Width           =   480
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    ��������Ϊ���ɵ�ԭ����ʱ���ܽ��п�չ������������ת���������Ա������������ת���������ָ�����������ȡ��ת�����á�"
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   750
      TabIndex        =   0
      Top             =   135
      Width           =   4470
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmApparatusTran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLngAptId As Long   '��ǰ����id
Private mblnOK As Boolean

'��ʱ����
Dim rsTemp As New ADODB.Recordset
Dim lngCount As Long

Public Function ShowMe(ByVal frmParent As Form, lngAptId As Long) As Boolean
    mLngAptId = lngAptId
    
    Me.dtp��ʼʱ��.MinDate = Now - 365
    Me.dtp��ʼʱ��.Value = Now()
    
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select T.ID, T.����, T.����, N.ת������id, N.ת������" & vbNewLine & _
            "From �������� T, �������� N" & vbNewLine & _
            "Where T.�������� = N.�������� And T.ID <> [1] And N.ID = [1] And" & vbNewLine & _
            "      (T.ת������ Is Null Or T.ת������ Is Not Null And T.ת������id = [1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngAptId)
    With rsTemp
        Me.cboת������.Clear
        Do While Not .EOF
            Me.cboת������.AddItem !���� & "-" & !����
            Me.cboת������.ItemData(Me.cboת������.NewIndex) = !ID
            If Val("" & !ת������id) = !ID Then
                Me.dtp��ʼʱ��.Value = Format(!ת������, "yyyy-MM-dd hh:mm")
                Me.cboת������.ListIndex = Me.cboת������.NewIndex
            End If
            .MoveNext
        Loop
    End With
    If Me.cboת������.ListCount = 0 Then
        Me.optTran(0).Value = False: Me.optTran(1).Value = True
        Me.optTran(0).Enabled = False: Me.optTran(1).Enabled = False
    Else
        Me.optTran(0).Enabled = True: Me.optTran(1).Enabled = True
        If Me.cboת������.ListIndex = -1 Then
            Me.cboת������.ListIndex = 0
            Me.optTran(0).Value = False: Me.optTran(1).Value = True
        Else
            Me.optTran(0).Value = True: Me.optTran(1).Value = False
        End If
    End If
    
    Me.Show vbModal, frmParent
    ShowMe = mblnOK: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ShowMe = False: Exit Function
End Function

Private Sub cboת������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False: Unload Me
End Sub

Private Sub cmdOK_Click()

    If Me.optTran(0).Value = True Then
        If Me.cboת������.ListIndex = -1 Then MsgBox "��δ����ת����Ŀ��������", vbInformation, gstrSysName: Exit Sub
        gstrSql = mLngAptId & ",To_Date('" & Format(Me.dtp��ʼʱ��, "yyyy-MM-dd hh:mm") & "','yyyy-mm-dd hh24:mi')"
        gstrSql = gstrSql & "," & Me.cboת������.ItemData(Me.cboת������.ListIndex)
    Else
        gstrSql = mLngAptId & ",Null,Null"
    End If
    gstrSql = "Zl_��������ת��_Set(" & gstrSql & ")"
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest

    mblnOK = True: Unload Me
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub dtp��ʼʱ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Activate()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub optTran_Click(Index As Integer)
    Me.dtp��ʼʱ��.Enabled = Me.optTran(0).Value
    Me.cboת������.Enabled = Me.optTran(0).Value
End Sub

Private Sub optTran_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub
