VERSION 5.00
Begin VB.Form frm��ʼ���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���µ���ʼ�����趨"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4665
   Icon            =   "frm��ʼ����.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1920
      TabIndex        =   4
      Top             =   1710
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3180
      TabIndex        =   5
      Top             =   1710
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   -60
      TabIndex        =   3
      Top             =   1380
      Width           =   5175
   End
   Begin VB.ComboBox cbo��ʼ���� 
      Height          =   300
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   900
      Width           =   3105
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frm��ʼ����.frx":000C
      ForeColor       =   &H00C00000&
      Height          =   585
      Left            =   300
      TabIndex        =   2
      Top             =   150
      Width           =   4035
   End
   Begin VB.Label lbl��ʼ���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ʼ����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   315
      TabIndex        =   0
      Top             =   960
      Width           =   720
   End
End
Attribute VB_Name = "frm��ʼ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mblnReturn As Boolean

Public Function ShowME(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    If lng����ID = 0 Then Exit Function
    
    mblnReturn = False
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    Me.Show 1
    ShowME = mblnReturn
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHand
    
    gcnOracle.Execute "ZL_���µ���ʼ����_UPDATE(" & mlng����ID & "," & mlng��ҳID & ",'" & Mid(Me.cbo��ʼ����.Text, 2, 19) & "')", , adCmdStoredProc
    
    mblnReturn = True
    Unload Me
    Exit Sub
errHand:
    If errCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Dim str��ʼ���� As String
    Dim intStart As Integer, intEnd As Integer
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = " SELECT '��'||TO_CHAR(��ʼʱ��,'YYYY-MM-DD HH24:MI:SS')||DECODE(��ʼԭ��,1,'��Ժ',2,'���-'||B.����,'ת��-'||B.����) AS ����" & _
              " FROM ���˱䶯��¼ A,���ű� B" & _
              " WHERE A.����ID=B.ID AND A.��ʼԭ�� IN (1,2,3) AND A.����ID=[1] AND A.��ҳID=[2]" & _
              " ORDER BY A.����ID,A.��ҳID,A.��ʼԭ��,A.��ʼʱ��"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    With rsTemp
        Me.cbo��ʼ����.Clear
        Do While Not .EOF
            Me.cbo��ʼ����.AddItem !����
            .MoveNext
        Loop
        Me.cbo��ʼ����.ListIndex = 0
    End With
    
    '��ȡ������ҳ�ӱ��е����µ���ʼ����
    gstrSQL = " Select ��Ϣֵ From ������ҳ�ӱ� Where ����ID=[1] And ��ҳID=[2] And ��Ϣ��='���µ���ʼ����'"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If rsTemp.RecordCount <> 0 Then
        str��ʼ���� = Format(rsTemp!��Ϣֵ, "yyyy-MM-dd HH:mm:ss")
    End If
    
    '��λ
    If str��ʼ���� <> "" Then
        intEnd = Me.cbo��ʼ����.ListCount
        For intStart = 1 To intEnd
            If InStr(1, Me.cbo��ʼ����.List(intStart - 1), str��ʼ����) <> 0 Then
                Me.cbo��ʼ����.ListIndex = intStart - 1
                Exit For
            End If
        Next
    End If
    
    'ֻҪ�������������������޸����µ���ʼ����
    gstrSQL = " Select 1 From ���˻����¼ A,���˻������� B Where B.��¼ID=A.ID And B.��Ŀ��� IN (1,2,3) and A.����ID=[1] And A.��ҳID=[2] And Rownum<2"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    cmdOK.Enabled = (rsTemp.RecordCount = 0)
End Sub
