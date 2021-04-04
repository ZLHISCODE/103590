VERSION 5.00
Begin VB.Form frmDiseaseStationSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������淶Χ����"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   30
      TabIndex        =   4
      Top             =   525
      Width           =   5730
   End
   Begin VB.ListBox lstFiles 
      Height          =   1320
      Left            =   1080
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   960
      Width           =   3210
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   2385
      Width           =   5730
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2565
      TabIndex        =   1
      Top             =   2565
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3675
      TabIndex        =   0
      Top             =   2565
      Width           =   1100
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   195
      Picture         =   "frmDiseaseStationSet.frx":0000
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ñ�����վ����ļ��������ļ���"
      Height          =   180
      Left            =   720
      TabIndex        =   6
      Top             =   150
      Width           =   3975
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblFiles 
      AutoSize        =   -1  'True
      Caption         =   "������վ�ɹ����ļ�(&F):"
      Height          =   180
      Left            =   90
      TabIndex        =   5
      Top             =   690
      Width           =   1980
   End
End
Attribute VB_Name = "frmDiseaseStationSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOk As Boolean

Public Function ShowMe(ByVal frmParent As Object, ByVal blnFiles As Boolean, ByRef strFiles As String) As Boolean
'���ܣ���ʾ�����岢�ṩ�û�����
'������ blnFiles,   �Ƿ����������ļ�
'       strFiles,   Ŀǰ�ɹ�����ļ�id�б�
    Dim strSetFiles As String, strReturn As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngCount As Long
    Dim strSQL As String

    strSetFiles = Trim(gobjComlib.zlDatabase.GetPara("������վ�ɹ����ļ�", glngSys, 1278))
  On Error GoTo errHand
    strSQL = "Select Id, ���, ���� From �����ļ��б� Where ���� = 5  Order By ���"
    Set rsTemp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With rsTemp
        Me.lstFiles.Clear
        Do While Not .EOF
            'Ϊ֧���²����ո񿴲�����ͬʱ�������ָ���
            Me.lstFiles.AddItem !��� & "-" & !���� & "                                   " & !ID
            Me.lstFiles.ItemData(Me.lstFiles.NewIndex) = !ID
            If InStr(1, "," & strSetFiles & ",", "," & !ID & ",") > 0 Then
                Me.lstFiles.Selected(Me.lstFiles.NewIndex) = True
            End If
            .MoveNext
        Loop
    End With
    
    Me.lstFiles.Enabled = blnFiles
    
    '��ʾ����
    Me.Show vbModal, frmParent
    
    '���ش���
    If mblnOk Then
        If Me.lstFiles.Enabled Then
            strFiles = ""
            For lngCount = 0 To Me.lstFiles.ListCount - 1
                If Me.lstFiles.Selected(lngCount) Then
                    If IsNumeric(Split(lstFiles.List(lngCount), "                                   ")(1)) Then
                        strFiles = strFiles & "," & Me.lstFiles.ItemData(lngCount)
                    End If
                End If
            Next
            If strFiles <> "" Then strFiles = Mid(strFiles, 2)
            Call gobjComlib.zlDatabase.SetPara("������վ�ɹ����ļ�", strFiles, glngSys, 1278)
        End If
    End If
    ShowMe = mblnOk
    Unload Me
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Unload Me
End Function

Private Sub cmdCancel_Click()
    mblnOk = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim blnSelected As Boolean
    Dim lngCount As Long
   
    If Me.lstFiles.Enabled Then
        For lngCount = 0 To Me.lstFiles.ListCount - 1
            If Me.lstFiles.Selected(lngCount) Then
                blnSelected = True
                Exit For
            End If
        Next
        If Not blnSelected Then
            MsgBox "û�����ñ�����վ�ɹ���ļ��������ļ���", vbExclamation, gstrSysName: Me.lstFiles.SetFocus: Exit Sub
        End If
    End If
    mblnOk = True
    Me.Hide
End Sub

Private Sub lstFiles_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call gobjComlib.zlCommFun.PressKey(vbKeyTab)
End Sub
