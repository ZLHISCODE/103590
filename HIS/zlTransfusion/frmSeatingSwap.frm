VERSION 5.00
Begin VB.Form frmSeatingSwap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������λ"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   Icon            =   "frmSeatingSwap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   4845
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraObject 
      Height          =   1470
      Left            =   120
      TabIndex        =   14
      Top             =   1950
      Width           =   4545
      Begin VB.TextBox txtObject 
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   870
         TabIndex        =   17
         Top             =   960
         Width           =   3540
      End
      Begin VB.TextBox txtObject 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   870
         TabIndex        =   15
         Top             =   600
         Width           =   3540
      End
      Begin VB.ComboBox cboObject 
         Height          =   300
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3540
      End
      Begin VB.Label lblObject 
         Caption         =   "����λ"
         Height          =   180
         Index           =   0
         Left            =   285
         TabIndex        =   19
         Top             =   300
         Width           =   555
      End
      Begin VB.Label lblObject 
         Caption         =   "�շ���Ŀ"
         Height          =   180
         Index           =   2
         Left            =   105
         TabIndex        =   18
         Top             =   1005
         Width           =   735
      End
      Begin VB.Label lblObject 
         Caption         =   "�۸�"
         Height          =   180
         Index           =   1
         Left            =   495
         TabIndex        =   16
         Top             =   690
         Width           =   360
      End
   End
   Begin VB.Frame fraSource 
      Height          =   1860
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   4545
      Begin VB.TextBox txtSource 
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   900
         TabIndex        =   12
         Top             =   1395
         Width           =   3540
      End
      Begin VB.TextBox txtSource 
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   900
         TabIndex        =   10
         Top             =   1010
         Width           =   3540
      End
      Begin VB.TextBox txtSource 
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   900
         TabIndex        =   8
         Top             =   625
         Width           =   3540
      End
      Begin VB.TextBox txtSource 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   2715
         TabIndex        =   6
         Top             =   240
         Width           =   1710
      End
      Begin VB.TextBox txtSource 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   885
         TabIndex        =   4
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label lblSoucer 
         Alignment       =   1  'Right Justify
         Caption         =   "�շ���Ŀ"
         Height          =   180
         Index           =   4
         Left            =   75
         TabIndex        =   13
         Top             =   1455
         Width           =   780
      End
      Begin VB.Label lblSoucer 
         Caption         =   "�۸�"
         Height          =   180
         Index           =   3
         Left            =   465
         TabIndex        =   11
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label lblSoucer 
         Alignment       =   1  'Right Justify
         Caption         =   "��λ"
         Height          =   180
         Index           =   2
         Left            =   60
         TabIndex        =   9
         Top             =   675
         Width           =   780
      End
      Begin VB.Label lblSoucer 
         Caption         =   "����"
         Height          =   180
         Index           =   1
         Left            =   2310
         TabIndex        =   7
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lblSoucer 
         Alignment       =   1  'Right Justify
         Caption         =   "����"
         Height          =   180
         Index           =   0
         Left            =   60
         TabIndex        =   5
         Top             =   285
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdCancle 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3405
      TabIndex        =   2
      Top             =   3540
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2205
      TabIndex        =   1
      Top             =   3540
      Width           =   1100
   End
End
Attribute VB_Name = "frmSeatingSwap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mObjKey As String
Private mObjSeatings  As Seatings

Public Function ObjectKey(ByVal strSourceKey As String, ByVal objSeatings As Seatings, ByVal frmMain As Form, Optional strObjKey As String) As String
    
    Dim objSeating As Seating, intIndex As Integer
    
    If objSeatings Is Nothing Then Exit Function
    Set mObjSeatings = objSeatings
    
    Set objSeating = mObjSeatings(strSourceKey)
    If objSeating Is Nothing Then Exit Function
    
    txtSource(0) = objSeating.����
    txtSource(1) = objSeatings.��������
    txtSource(2) = IIf(objSeating.���� = "", "��ͨ��λ", objSeating.����) & " " & objSeating.���
    txtSource(3) = Format(objSeating.�ּ�, "0.00")
    txtSource(4) = objSeating.�շ���Ŀ
    mObjKey = strObjKey
    
    cboObject.Clear
    For Each objSeating In mObjSeatings
        If Val("" & objSeating.����ID) = 0 And objSeating.״̬ = 0 Then
            cboObject.AddItem IIf(objSeating.���� = "", "��ͨ��λ", objSeating.����) & " " & objSeating.���
            
            If strObjKey <> "" Then
                If cboObject.List(cboObject.ListCount - 1) = IIf(mObjSeatings(strObjKey).���� = "", "��ͨ��λ", mObjSeatings(strObjKey).����) & " " & mObjSeatings(strObjKey).��� Then
                     intIndex = cboObject.ListCount - 1
                End If
            End If
        End If
    Next
    
    If cboObject.ListCount > 0 Then
        cboObject.ListIndex = intIndex
        Call cboObject_Click
        Me.Show vbModal, frmMain
        ObjectKey = mObjKey
        mObjKey = ""
    Else
        MsgBox "�޿�����λ��", vbInformation, gstrSysName
    End If
End Function

Private Sub cboObject_Click()
    mObjKey = getKey(cboObject.List(cboObject.ListIndex))
    txtObject(1) = Format(mObjSeatings(mObjKey).�ּ�, "0.00")
    txtObject(2) = mObjSeatings(mObjKey).�շ���Ŀ
    
End Sub

Private Sub cmdCancle_Click()
    mObjKey = ""
    Unload Me
End Sub

Private Sub cmdOk_Click()
    mObjKey = getKey(cboObject.List(cboObject.ListIndex))
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mObjSeatings = Nothing
End Sub
 

Private Function getKey(ByVal strType As String) As String
    Dim objSeating As Seating
    For Each objSeating In mObjSeatings
        If strType = IIf(objSeating.���� = "", "��ͨ��λ", objSeating.����) & " " & objSeating.��� Then
            getKey = objSeating.Key
            Exit For
        End If
    Next
    
End Function
