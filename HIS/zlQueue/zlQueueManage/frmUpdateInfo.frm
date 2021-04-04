VERSION 5.00
Begin VB.Form frmUpdateInfo 
   Caption         =   "�޸Ķ�����Ϣ"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4365
   Icon            =   "frmUpdateInfo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   4365
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4095
      Begin VB.ComboBox cboҽ�� 
         Height          =   300
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1800
         Width           =   2295
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txt�������� 
         Height          =   350
         Left            =   1560
         TabIndex        =   6
         Top             =   840
         Width           =   2295
      End
      Begin VB.ComboBox cboQueueName 
         Height          =   300
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "ҽ������"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "���� "
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "��������"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "��������"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   350
      Left            =   2280
      TabIndex        =   1
      Top             =   2640
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Height          =   350
      Left            =   480
      TabIndex        =   0
      Top             =   2640
      Width           =   1100
   End
End
Attribute VB_Name = "frmUpdateInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mrsRoomData As ADODB.Recordset
Public mrsDoctorData As ADODB.Recordset
Public mlngCurrentQueueId As Long
Public mblnIsAllowChange As Boolean
Public mblnIsAlreadyProcess As Boolean


Private Const C_STR_MSGINF As String = "�޸��Ŷ���Ϣ"

Private mstr�������� As String
Private mstr�������� As String
Private mstr���� As String
Private mstrҽ������ As String

Public Event OnQueueChange(ByVal lngQueueId As Long, ByVal strQueue As String, ByVal strPatient As String, ByVal strRoom As String, ByVal strDoctor As String, ByRef blnIsAllowChange As Boolean, ByRef blnIsAlreadyProcess As Boolean)


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub LoadQueueName(ByRef astr��������() As String)
    
End Sub

Public Function zlShowMe(frmParent As Form, ByRef astr��������() As String, ByRef str�������� As String, str�������� As String, _
            ByRef str���� As String, ByRef strҽ������ As String) As Boolean
    Dim i As Integer
    
    mstr�������� = str��������
    mstr�������� = str��������
    mstr���� = str����
    mstrҽ������ = strҽ������

    On Error GoTo err
    
    cboQueueName.Clear
    
    If SafeArrayGetDim(astr��������) <> 0 Then
        For i = 1 To UBound(astr��������)
            cboQueueName.AddItem astr��������(i)
            If astr��������(i) = str�������� Then cboQueueName.ListIndex = i - 1
        Next i
        
        If cboQueueName.ListIndex = -1 Then Exit Function
        
        txt�������� = mstr��������
        
        '��������cbo����
        cbo����.Clear
        If Not mrsRoomData Is Nothing Then
            If mrsRoomData.RecordCount < 1 Then
                cbo����.Enabled = False
                MsgBox "��ѡ����������", vbInformation, C_STR_MSGINF
            End If
            For i = 1 To mrsRoomData.RecordCount
                cbo����.AddItem Nvl(mrsRoomData!RoomCode) & "-" & Nvl(mrsRoomData!RoomName)
                cbo����.ItemData(i - 1) = Nvl(mrsRoomData!RoomID)
                
                If Nvl(mrsRoomData!RoomName) = mstr���� Then
                    cbo����.ListIndex = i - 1
                End If
                mrsRoomData.MoveNext
            Next
        Else
            cbo����.Enabled = False
            MsgBox "��ѡ����������", vbInformation, C_STR_MSGINF
        End If
        
        '����ҽ��cbo����
        cboҽ��.Clear
        If Not mrsDoctorData Is Nothing Then
        
            If mrsDoctorData.RecordCount < 1 Then
                cboҽ��.Enabled = False
                MsgBox "��ѡҽ��������", vbInformation, C_STR_MSGINF
            End If
            
            For i = 1 To mrsDoctorData.RecordCount
                cboҽ��.AddItem Nvl(mrsDoctorData!DoctorIdCode) & "-" & Nvl(mrsDoctorData!DoctorIdName)
                cboҽ��.ItemData(i - 1) = Nvl(mrsDoctorData!DoctorId)
                
                If Nvl(mrsDoctorData!DoctorIdName) = mstrҽ������ Then
                    cboҽ��.ListIndex = i - 1
                End If
                mrsDoctorData.MoveNext
            Next
        Else
            cboҽ��.Enabled = False
            MsgBox "��ѡҽ��������", vbInformation, C_STR_MSGINF
        End If

        
        Me.Show 1, frmParent

        If mstr�������� <> str�������� Or mstr�������� <> str�������� Or mstr���� <> str���� Or mstrҽ������ <> strҽ������ Then
            str�������� = mstr��������
            str�������� = mstr��������
            strҽ������ = mstrҽ������
            str���� = mstr����

            zlShowMe = True
            
        End If
    End If
       
    
      
    Exit Function
    
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function getNameByCbo(ByVal strText As String) As String
'���cboѡ�����ݵ�����
    On Error GoTo errh
    
    getNameByCbo = ""
    If InStr(strText, "-") < 1 Then Exit Function
   
    getNameByCbo = Mid(strText, InStr(strText, "-") + 1, Len(strText))
    Exit Function
    
errh:
    Resume
    getNameByCbo = ""
End Function

Private Function getCodeByCbo(ByVal strText As String) As Long
'���cboѡ�����ݵļ���
    On Error GoTo errh
    
    getCodeByCbo = 0
    If InStr(strText, "-") < 1 Then Exit Function
    
    getCodeByCbo = Val(Mid(strText, 1, InStr(strText, "-") - 1))
    Exit Function
      
errh:
    Resume
    getCodeByCbo = 0
End Function

Private Sub cmdOK_Click()
    
    mstr�������� = cboQueueName.Text
    mstr�������� = txt��������.Text
    
    If mstrҽ������ <> getNameByCbo(cboҽ��.Text) And cboҽ��.Enabled = True Then mstrҽ������ = getNameByCbo(cboҽ��.Text)

    If mstr���� <> getNameByCbo(cbo����.Text) And cbo����.Enabled = True Then mstr���� = getNameByCbo(cbo����.Text)

    RaiseEvent OnQueueChange(mlngCurrentQueueId, mstr��������, mstr��������, mstr����, mstrҽ������, mblnIsAllowChange, mblnIsAlreadyProcess)
    
    If mblnIsAllowChange = False Then Exit Sub
    
    Unload Me
End Sub

