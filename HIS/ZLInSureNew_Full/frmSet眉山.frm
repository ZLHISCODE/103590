VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BILLEDIT.OCX"
Begin VB.Form frmSetüɽ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���в�������"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   Icon            =   "frmSetüɽ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3450
      TabIndex        =   16
      Top             =   4920
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2160
      TabIndex        =   15
      Top             =   4920
      Width           =   1100
   End
   Begin TabDlg.SSTab TabShow 
      Height          =   1785
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   3149
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "������(&1)"
      TabPicture(0)   =   "frmSetüɽ.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra���ز�_������"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "��ҵ��˾(&2)"
      TabPicture(1)   =   "frmSetüɽ.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra���ز�_��ҵ��˾"
      Tab(1).ControlCount=   1
      Begin VB.Frame fra���ز�_��ҵ��˾ 
         Caption         =   "���ز�(&K)"
         Height          =   1125
         Left            =   -74760
         TabIndex        =   8
         Top             =   480
         Width           =   4215
         Begin VB.TextBox txt���ز���������_��ҵ��˾ 
            Height          =   300
            Left            =   1230
            MaxLength       =   16
            TabIndex        =   10
            Tag             =   "21"
            Top             =   270
            Width           =   2235
         End
         Begin VB.TextBox txt���ز������޶�_��ҵ��˾ 
            Height          =   300
            Left            =   1230
            MaxLength       =   16
            TabIndex        =   13
            Tag             =   "22"
            Top             =   660
            Width           =   2235
         End
         Begin VB.Label lbl���ز���������_��ҵ��˾ 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   360
            TabIndex        =   9
            Top             =   330
            Width           =   720
         End
         Begin VB.Label lbl���ز������޶�_��ҵ��˾ 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�����޶�"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   360
            TabIndex        =   12
            Top             =   720
            Width           =   720
         End
         Begin VB.Label lbl��λ 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   180
            Index           =   6
            Left            =   3540
            TabIndex        =   11
            Top             =   330
            Width           =   90
         End
         Begin VB.Label lbl��λ 
            AutoSize        =   -1  'True
            Caption         =   "Ԫ"
            Height          =   180
            Index           =   5
            Left            =   3540
            TabIndex        =   17
            Top             =   690
            Width           =   180
         End
      End
      Begin VB.Frame fra���ز�_������ 
         Caption         =   "���ز�(&M)"
         Height          =   1125
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   4215
         Begin VB.TextBox txt���ز������޶�_������ 
            Height          =   300
            Left            =   1230
            MaxLength       =   16
            TabIndex        =   6
            Tag             =   "12"
            Top             =   660
            Width           =   2235
         End
         Begin VB.TextBox txt���ز���������_������ 
            Height          =   300
            Left            =   1230
            MaxLength       =   16
            TabIndex        =   3
            Tag             =   "11"
            Top             =   270
            Width           =   2235
         End
         Begin VB.Label lbl��λ 
            AutoSize        =   -1  'True
            Caption         =   "Ԫ"
            Height          =   180
            Index           =   1
            Left            =   3540
            TabIndex        =   7
            Top             =   690
            Width           =   180
         End
         Begin VB.Label lbl��λ 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   180
            Index           =   0
            Left            =   3540
            TabIndex        =   4
            Top             =   330
            Width           =   90
         End
         Begin VB.Label lbl���ز������޶�_������ 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�����޶�"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   360
            TabIndex        =   5
            Top             =   720
            Width           =   720
         End
         Begin VB.Label lbl���ز���������_������ 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   360
            TabIndex        =   2
            Top             =   330
            Width           =   720
         End
      End
   End
   Begin ZL9BillEdit.BillEdit Bill 
      Height          =   2625
      Left            =   0
      TabIndex        =   14
      Top             =   2190
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4630
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.Label lblNOte 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Caption         =   "����ҽ��֧�������ʵ�ʱ���������"
      ForeColor       =   &H80000002&
      Height          =   180
      Left            =   90
      TabIndex        =   18
      Top             =   1890
      Width           =   2880
   End
End
Attribute VB_Name = "frmSetüɽ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng���� As Long
Private mblnReturn As Boolean
Private Const str�޶� As String = "txt���ز������޶�_������|txt���ز������޶�_��ҵ��˾"

Public Function ShowSet() As Boolean
    mblnReturn = False
    Me.Show 1
    ShowSet = mblnReturn
End Function

Private Sub Bill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub Bill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Bill
        If .TxtVisible = False Then Exit Sub
        If Not IsNumeric(.Text) Then
            MsgBox "�������������0-100��", vbInformation, gstrSysName
            Exit Sub
        End If
        If Not (Val(.Text) >= 0 And Val(.Text) <= 100) Then
            MsgBox "����ı�������С��������100��", vbInformation, gstrSysName
            Exit Sub
        End If
        .Text = Format(.Text, "#####0.00;-#####0.00; ;")
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngRow As Long
    Dim TextBox As Object
    On Error GoTo errHand
    
    gcnOracle.BeginTrans
    'ɾ���Ѿ�����
    gstrSQL = "zl_���ղ���_Delete(" & mlng���� & ",1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Delete(" & mlng���� & ",2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '������������
    For Each TextBox In Me.Controls
        If TypeName(TextBox) = "TextBox" Then
            gstrSQL = "zl_���ղ���_Insert(" & mlng���� & "," & Mid(TextBox.Tag, 1, 1) & ",'" & Mid(TextBox.Name, 4) & "','" & Val(TextBox.Text) & "'," & Mid(TextBox.Tag, 2) & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
    Next
    For lngRow = 1 To Bill.Rows - 1
        gstrSQL = "zl_���ղ���_Insert(" & mlng���� & ",1,'" & Bill.TextMatrix(lngRow, 0) & "','" & Val(Bill.TextMatrix(lngRow, 1)) & "'," & 10 + lngRow & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    
    gcnOracle.CommitTrans
    mblnReturn = True
    
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Me.ActiveControl.Name <> "Bill" Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim TextBox As Object
    Dim rsTemp As New ADODB.Recordset
    
    Call InitBill
    mlng���� = TYPE_�Ĵ�üɽ
    gstrSQL = "Select * From ���ղ��� Where ����=" & mlng���� & " Order by ����,���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����)
    
    If rsTemp.EOF Then Exit Sub
    For Each TextBox In Me.Controls
        If TypeName(TextBox) = "TextBox" Then
            rsTemp.MoveFirst
            rsTemp.Filter = "����=" & Mid(TextBox.Tag, 1, 1) & " And ���=" & Mid(TextBox.Tag, 2)
            If Not rsTemp.EOF Then
                TextBox.Text = Format(rsTemp!����ֵ, "#####0.00;-#####0.00; ;")
            End If
        End If
    Next
    
    rsTemp.Filter = 0
    gstrSQL = " Select A.����,B.����ֵ From ����֧������ A,(Select * From ���ղ��� Where ����=[1] And ����=1 And ���>10) B " & _
              " Where A.����=[2] And A.����=B.������(+) Order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����)
    With rsTemp
        If rsTemp.RecordCount <> 0 Then Bill.Rows = 1 + rsTemp.RecordCount
        Do While Not .EOF
            Bill.TextMatrix(.AbsolutePosition, 0) = !����
            Bill.TextMatrix(.AbsolutePosition, 1) = Format(Nvl(!����ֵ, 100), "#####0.00;-#####0.00; ;")
            .MoveNext
        Loop
    End With
    Bill.Active = (rsTemp.RecordCount <> 0)
    Bill.AllowAddRow = False
End Sub

Private Sub InitBill()
    With Bill
        .ClearBill
        .Active = True
        .Rows = 2
        .Cols = 2
        .msfObj.FixedCols = 1
        
        .TextMatrix(0, 0) = "����"
        .TextMatrix(0, 1) = "ʵ�ʱ�������"
        
        .ColWidth(0) = 1500
        .ColWidth(1) = 2000
        .msfObj.ColAlignmentFixed = 1
        .ColData(0) = 5
        .ColData(1) = 4
        
        .PrimaryCol = 1
        .LocateCol = 1
    End With
End Sub

Private Function Valid() As Boolean
    Dim TextBox As Object
    Dim sinMax As Single, sinMin As Single
    
    For Each TextBox In Me.Controls
        If TypeName(TextBox) = "TextBox" Then
            If InStr(1, str�޶�, TextBox.Name) <> 0 Then
                '���ܴ���90009000999.99�Ҳ���С��0
                sinMax = 1E+15
                sinMin = 0
            Else
                '���ܴ���100��С��0
                sinMax = 100
                sinMin = 0
            End If
            If Trim(TextBox.Text) <> "" Then
                If Not IsNumeric(TextBox.Text) Then
                    MsgBox Mid(TextBox.Name, 4) & "�к��зǷ��ַ������飡", vbInformation, gstrSysName
                    TextBox.SetFocus
                    Exit Function
                End If
                If Val(TextBox.Text) > sinMax Then
                    MsgBox Mid(TextBox.Name, 4) & "�������ֵ��", vbInformation, gstrSysName
                    TextBox.SetFocus
                    Exit Function
                End If
                If Val(TextBox.Text) < sinMin Then
                    MsgBox Mid(TextBox.Name, 4) & "����С���㣡", vbInformation, gstrSysName
                    TextBox.SetFocus
                    Exit Function
                End If
                
            End If
        End If
    Next
    
    Valid = True
End Function

Private Sub txt���ز���������_������_GotFocus()
    Call zlControl.TxtSelAll(txt���ز���������_������)
End Sub

Private Sub txt���ز���������_������_LostFocus()
    txt���ز���������_������.Text = Format(txt���ز���������_������.Text, "#####0.00;-#####0.00; ;")
End Sub

Private Sub txt���ز������޶�_������_GotFocus()
    Call zlControl.TxtSelAll(txt���ز������޶�_������)
End Sub

Private Sub txt���ز������޶�_������_LostFocus()
    txt���ز������޶�_������.Text = Format(txt���ز������޶�_������.Text, "#####0.00;-#####0.00; ;")
End Sub

Private Sub txt���ز���������_��ҵ��˾_GotFocus()
    tabShow.Tab = 1
    Call zlControl.TxtSelAll(txt���ز���������_��ҵ��˾)
End Sub

Private Sub txt���ز���������_��ҵ��˾_LostFocus()
    txt���ز���������_��ҵ��˾.Text = Format(txt���ز���������_��ҵ��˾.Text, "#####0.00;-#####0.00; ;")
End Sub

Private Sub txt���ز������޶�_��ҵ��˾_GotFocus()
    Call zlControl.TxtSelAll(txt���ز������޶�_��ҵ��˾)
End Sub

Private Sub txt���ز������޶�_��ҵ��˾_LostFocus()
    txt���ز������޶�_��ҵ��˾.Text = Format(txt���ز������޶�_��ҵ��˾.Text, "#####0.00;-#####0.00; ;")
End Sub
