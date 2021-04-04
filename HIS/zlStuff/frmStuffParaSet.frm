VERSION 5.00
Begin VB.Form frmStuffParaSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4860
   Icon            =   "frmStuffParaSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   4860
   StartUpPosition =   2  '��Ļ����
   Begin VB.ComboBox cboSendDept 
      Height          =   300
      ItemData        =   "frmStuffParaSet.frx":6852
      Left            =   1080
      List            =   "frmStuffParaSet.frx":6854
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   112
      Width           =   3375
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3840
      TabIndex        =   5
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2640
      TabIndex        =   4
      Top             =   3600
      Width           =   975
   End
   Begin VB.Frame fraOther 
      Caption         =   "���ܿ��������豸����"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   4575
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "�豸����"
         Height          =   350
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame fraUnit 
      Caption         =   "ȱʡ��λ"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   4575
      Begin VB.OptionButton optUnit 
         Caption         =   "��װ��λ"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   13
         Top             =   300
         Width           =   1215
      End
      Begin VB.OptionButton optUnit 
         Caption         =   "ɢװ��λ"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame fraType 
      Caption         =   "ҵ������"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4575
      Begin VB.ComboBox cboNO 
         Height          =   300
         ItemData        =   "frmStuffParaSet.frx":6856
         Left            =   960
         List            =   "frmStuffParaSet.frx":6858
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   700
         Width           =   3375
      End
      Begin VB.CheckBox chkType 
         Caption         =   "���ʱ�"
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   9
         Top             =   300
         Width           =   855
      End
      Begin VB.CheckBox chkType 
         Caption         =   "���ʵ�"
         Height          =   255
         Index           =   1
         Left            =   2220
         TabIndex        =   8
         Top             =   300
         Width           =   855
      End
      Begin VB.CheckBox chkType 
         Caption         =   "�շѵ�"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   7
         Top             =   300
         Width           =   855
      End
      Begin VB.Label lblNo 
         Caption         =   "�շѵ���"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   723
         Width           =   735
      End
      Begin VB.Label lblType 
         Caption         =   "��������"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.Label lblSendDept 
      Caption         =   "���ϲ���"
      Height          =   225
      Left            =   240
      TabIndex        =   14
      Top             =   150
      Width           =   735
   End
End
Attribute VB_Name = "frmStuffParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long  'ģ���
Private mstrPrivs As String 'Ȩ�޴�
Private mblnOk As Boolean   '���������Ƿ�ɹ�

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call FS.DeviceSetup(Me, glngSys, mlngModule)
End Sub

Private Sub CmdSave_Click()
    Dim strҵ������ As String
    
    strҵ������ = IIf(chkType(0).Value = 1, "24", "0")
    strҵ������ = strҵ������ & IIf(chkType(1).Value = 1, ",25", ",0")
    strҵ������ = strҵ������ & IIf(chkType(2).Value = 1, ",26", ",0")
    
    On Error GoTo ErrHandle
    Call zlDatabase.SetPara("��ѯҵ������", strҵ������, glngSys, mlngModule)
    Call zlDatabase.SetPara("���ĵ�λ", IIf(optUnit(1).Value = True, 1, 0), glngSys, mlngModule)
    Call zlDatabase.SetPara("�շѴ�����ʾ��ʽ", cboNO.ListIndex, glngSys, mlngModule)
    
    Call zlDatabase.SetPara("���Ͽ���", cboSendDept.ItemData(cboSendDept.ListIndex), glngSys, mlngModule)
    Unload Me
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim strReg As String
    Dim arrStr As Variant
    Dim i As Integer
    Dim blnSetPara As Boolean
    Dim lng���ϲ���ID As Long

    With cboNO
        .Clear
        .AddItem "1-��ʾ���еĴ���"
        .AddItem "2-����ʾ���շѴ���"
        .AddItem "3-����ʾδ�շѴ���"
        .ListIndex = 0
    End With
    
    blnSetPara = (InStr(1, mstrPrivs, "��������") > 0)
    
    strReg = Val(zlDatabase.GetPara("�շѴ�����ʾ��ʽ", glngSys, mlngModule, 0, Array(LblNo, cboNO), blnSetPara))
    If Val(strReg) >= 0 And strReg <= 2 Then
        cboNO.ListIndex = Val(strReg)
    Else
        cboNO.ListIndex = 0
    End If
    
    strReg = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngModule, "0", Array(optUnit(0), optUnit(1), fraUnit), blnSetPara))
    optUnit(0).Value = False
    optUnit(1).Value = False
    If Val(strReg) = 0 Then
        optUnit(0).Value = True
    Else
        optUnit(1).Value = True
    End If
    
    strReg = Trim(zlDatabase.GetPara("��ѯҵ������", glngSys, mlngModule, "", Array(lblType, chkType(0), chkType(1), chkType(2), fraType), blnSetPara))
    If strReg = "" Then strReg = "24,25,26"
    arrStr = Split(strReg & "," & "," & ",", ",")
    For i = 0 To UBound(arrStr)
        If i > 2 Then Exit For
        chkType(i).Value = IIf(Val(arrStr(i)) > 0, 1, 0)
    Next
    
    lng���ϲ���ID = Val(zlDatabase.GetPara("���Ͽ���", glngSys, mlngModule, "0", Array(lblSendDept, cboSendDept), blnSetPara))
    Call LoadDept(lng���ϲ���ID)
    
    
End Sub

Public Function ShowSetPara(ByVal frmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------------------
    '����:���ò������
    '����:
    '����:���óɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
    '���ز�������
    Me.Show 1, frmMain
    ShowSetPara = mblnOk
End Function

Private Sub LoadDept(ByVal lng���ϲ���ID As Long)
    Dim rsTemp As Recordset
    
    Set rsTemp = Stuff_GetDept(mstrPrivs)
    
    'װ�뷢�ϲ�������
    With cboSendDept
        .Clear
        Do While Not rsTemp.EOF
            .AddItem rsTemp!����
            .ItemData(.NewIndex) = rsTemp!Id
            If rsTemp!Id = lng���ϲ���ID Then
                .ListIndex = .NewIndex
            End If
            rsTemp.MoveNext
        Loop
        If .ListIndex = -1 Then .ListIndex = 0
    End With
End Sub
