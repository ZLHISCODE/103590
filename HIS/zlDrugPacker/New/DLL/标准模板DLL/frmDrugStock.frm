VERSION 5.00
Begin VB.Form frmDrugStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ҩƷ����ϴ�"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5730
   Icon            =   "frmDrugStock.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   5730
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   360
      Left            =   4320
      TabIndex        =   5
      Top             =   4560
      Width           =   1110
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "�ϴ�(&U)"
      Height          =   360
      Left            =   3120
      TabIndex        =   4
      Top             =   4560
      Width           =   1110
   End
   Begin VB.ListBox lstMess 
      Height          =   3480
      ItemData        =   "frmDrugStock.frx":0A02
      Left            =   240
      List            =   "frmDrugStock.frx":0A04
      TabIndex        =   2
      Top             =   960
      Width           =   5175
   End
   Begin VB.ComboBox cboDevice 
      Height          =   300
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label lblMess 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ϴ���Ϣ"
      Height          =   180
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   720
   End
   Begin VB.Label lblLink 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�豸"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   360
   End
End
Attribute VB_Name = "frmDrugStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngDeptID As Long

Public Sub ShowMe(ByVal lngDeptID As Long, ByVal frmOwner As Form)
    mlngDeptID = lngDeptID
    
    Call InitDevice
    
    If cboDevice.ListCount <= 0 Then
        MsgBox "��δע��ҩ���Զ����豸��", vbInformation, GSTR_INTERFACE_NAME
        Unload Me
        Exit Sub
    End If

    Me.Show vbModal, frmOwner
End Sub

Private Sub InitDevice()
'���ܣ���������
    
    Dim rsTmp As ADODB.Recordset
        
    'ͬһ�����ӵ��豸��ֻȡһ���豸
    gstrSQL = "Select a.Id, b.����, b.���� " & vbCr & _
              "From (Select Max(ID) ID " & vbCr & _
              "       From ҩ����ҩ�豸 " & vbCr & _
              "       Where �������� Is Not Null And �Ƿ����� = 1 And ʹ�ò���ID = [1] " & vbCr & _
              "       Group By ��������, ��������) A, ҩ����ҩ�豸 B " & vbCr & _
              "Where a.Id = b.Id " & vbCr & _
              "Order By b.���� "

    On Error GoTo errHandle
    Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "��ȡҩ����ҩ�豸", mlngDeptID)
    Do While Not rsTmp.EOF
        cboDevice.AddItem "��" & rsTmp!���� & "��" & rsTmp!����
        cboDevice.ItemData(cboDevice.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    If cboDevice.ListCount > 0 Then cboDevice.ListIndex = 0
    
    Exit Sub
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub


Private Sub cboDevice_Change()
    cmdUpload.Enabled = cboDevice.ListIndex >= 0
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdUpload_Click()
    Dim objDevice As clsDevice
    Dim rsData As ADODB.Recordset
    Dim strTmp As String

    '�豸����
    Set objDevice = New clsDevice
    objDevice.ID = cboDevice.ItemData(cboDevice.ListIndex)
    
    If objDevice.Status = False Then
        MsgBox "��" & objDevice.Name & "���豸δ���ӳɹ���", vbInformation, GSTR_INTERFACE_NAME
        Set objDevice = Nothing
        Exit Sub
    End If
    
    cmdUpload.Enabled = False
    Screen.MousePointer = vbHourglass
    
    '�õ�Ҫ�ϴ�������
    Set rsData = mdlProcessData.ProcDrugStock(mlngDeptID, objDevice)
    
    '�ϴ����豸
    If Not rsData Is Nothing Then
        lstMess.Clear
        With rsData
            If .State <> adStateOpen Then .Open
            If .RecordCount > 0 Then .MoveFirst
            Do While .EOF = False
                If mdlDrugPacker.DrugStock(objDevice, !Content) = False Then
                    strTmp = "ҩƷ��" & !Drug & "���ϴ�ʧ�ܣ�"
                    lstMess.AddItem strTmp
                End If
                .MoveNext
            Loop
            .Close
        End With
    End If
    
    Set objDevice = Nothing
    Screen.MousePointer = vbDefault
    cmdUpload.Enabled = True
    
End Sub

