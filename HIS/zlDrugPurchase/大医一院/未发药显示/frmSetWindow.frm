VERSION 5.00
Begin VB.Form frmSetWindow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ҩ��������"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3930
   Icon            =   "frmSetWindow.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   -120
      TabIndex        =   6
      Top             =   1600
      Width           =   5025
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   720
      TabIndex        =   4
      Top             =   1800
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2040
      TabIndex        =   5
      Top             =   1800
      Width           =   1100
   End
   Begin VB.ComboBox cboҩ�� 
      ForeColor       =   &H80000012&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1155
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2280
   End
   Begin VB.ListBox lst��ҩ���� 
      Columns         =   1
      ForeColor       =   &H80000012&
      Height          =   900
      IMEMode         =   3  'DISABLE
      ItemData        =   "frmSetWindow.frx":000C
      Left            =   1155
      List            =   "frmSetWindow.frx":000E
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   645
      Width           =   2280
   End
   Begin VB.Label Lblҩ�� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҩ��"
      Height          =   180
      Left            =   720
      TabIndex        =   0
      Top             =   300
      Width           =   360
   End
   Begin VB.Label Lbl��ҩ���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ҩ����"
      Height          =   180
      Left            =   360
      TabIndex        =   2
      Top             =   690
      Width           =   720
   End
End
Attribute VB_Name = "frmSetWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboҩ��_Click()
    Dim intDO As Integer
    Dim bln���� As Boolean, blnסԺ As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strTmp As String
    
    '�����ܣ����û������ҩ���������涼������
    If Me.cboҩ��.ListCount = 0 Then Exit Sub
    If Val(Me.cboҩ��.Tag) = Me.cboҩ��.ListIndex Then
        Exit Sub
    Else
        Me.cboҩ��.Tag = Me.cboҩ��.ListIndex
    End If
    
    '����ҩ����ʾ��λ
    strTmp = " Select ���� From ��ҩ���� Where ҩ��ID=" & Me.cboҩ��.ItemData(Me.cboҩ��.ListIndex)
    rsTmp.Open strTmp, gcnOracle
    
    With rsTmp
        Me.lst��ҩ����.Clear
        lst��ҩ����.Columns = 2
        Do While Not .EOF
            lst��ҩ����.AddItem !����
            .MoveNext
        Loop
        .Close
    End With

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub cmdOK_Click()
    Dim i As Integer
    Dim strFormNO As String
    
    If Me.cboҩ��.ListCount = 0 Then Exit Sub
    
    For i = 0 To Me.lst��ҩ����.ListCount - 1
        If Me.lst��ҩ����.Selected(i) Then
            strFormNO = Me.lst��ҩ����.List(i)
            Exit For
        End If
    Next
    
    SaveSetting "ZLSOFT", "δ��ҩ������ʾ", "ҩ��", cboҩ��.ItemData(cboҩ��.ListIndex)
    frmUnSendDrug.Entry cboҩ��.ItemData(cboҩ��.ListIndex), strFormNO
    Unload Me

End Sub

Private Sub Form_Load()
    Dim strTmp As String
    Dim lngStockID As Long, i As Long
    Dim rsTmp As New ADODB.Recordset
    strTmp = "Select Distinct p.Id, p.����" & vbNewLine & _
            "From ���ű� P" & vbNewLine & _
            "Where p.Id In (Select Distinct ����id From ��������˵�� Where �������� Like '%ҩ��') And" & vbNewLine & _
            "      (p.����ʱ�� Is Null Or p.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd'))" & vbNewLine & _
            "Order By p.����"
    With rsTmp
        Me.cboҩ��.Clear
        .Open strTmp, gcnOracle
        Do While Not .EOF
            Me.cboҩ��.AddItem !����
            Me.cboҩ��.ItemData(Me.cboҩ��.NewIndex) = !ID
            .MoveNext
        Loop
        .Close
        If Me.cboҩ��.ListCount > 0 Then
            lngStockID = Val(GetSetting(appName:="ZLSOFT", Section:="δ��ҩ������ʾ", Key:="ҩ��", Default:=""))
            If lngStockID > 0 Then
                For i = 0 To cboҩ��.ListCount - 1
                    If cboҩ��.ItemData(i) = lngStockID Then
                        cboҩ��.ListIndex = i
                        Exit For
                    End If
                Next
            Else
                cboҩ��.ListIndex = 0
            End If
        End If
    End With
    Call cboҩ��_Click
End Sub

Private Sub lst��ҩ����_ItemCheck(Item As Integer)
    Dim i As Integer
    On Error Resume Next
    For i = 0 To lst��ҩ����.ListCount - 1
        If i <> Item Then
            lst��ҩ����.Selected(i) = False
        End If
    Next
End Sub

