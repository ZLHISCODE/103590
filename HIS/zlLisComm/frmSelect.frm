VERSION 5.00
Begin VB.Form frmSelect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ѡ������"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancle 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4260
      TabIndex        =   2
      Top             =   2820
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   780
      TabIndex        =   1
      Top             =   2775
      Width           =   1100
   End
   Begin VB.ListBox lst�ɼ������� 
      Height          =   2580
      Left            =   195
      TabIndex        =   0
      Top             =   105
      Width           =   5955
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean

Public Function Select����() As Boolean
    mblnOK = False
    Me.Show vbModal
    Select���� = mblnOK
End Function

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim lngID As Long
    Dim i As Integer
    Dim blnAdd As Boolean
    If lst�ɼ�������.ListIndex >= 0 Then
        lngID = lst�ɼ�������.ItemData(lst�ɼ�������.ListIndex)
        blnAdd = False
        For i = LBound(g����) To UBound(g����)
            If g����(i).ID <= 0 Then
                blnAdd = True
                Exit For
            End If
        Next
        If Not blnAdd Then
            ReDim Preserve g����(UBound(g����) + 1)
            i = UBound(g����)
            blnAdd = True
        End If
        If blnAdd Then
            g����(i).ID = lngID
            g����(i).COM�� = 0
            g����(i).���� = 0
            g����(i).������ = 9600
            g����(i).����λ = 8
            g����(i).ֹͣλ = 1
            g����(i).У��λ = "N"
            g����(i).���� = 0
            g����(i).�ַ�ģʽ = 0
            g����(i).IP = "127.0.0.1"
            g����(i).IP�˿� = "6666"
            g����(i).���� = 0
            g����(i).SaveAsID = 0
            g����(i).�Զ�Ӧ�� = "0"
            g����(i).�ɷ��Ѻ˱걾 = "1"
            g����(i).ͨѶĿ¼ = App.Path & "\Dev_" & lngID
            g����(i).�Զ������ = ""
            g����(i).�Զ������ʿ� = 0
            g����(i).���Ϊͨ���� = 0
            mblnOK = True
            Unload Me
        End If
    End If
    
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, lngCount As Long
    Dim blnAdd As Boolean
    
    Set rsTmp = GetDevices
    lst�ɼ�������.Clear
    
    If rsTmp Is Nothing Then Exit Sub
    lngCount = 0
    Do Until rsTmp.EOF
        lngCount = lngCount + 1
        blnAdd = True
        For i = LBound(g����) To UBound(g����)
            If g����(i).ID = rsTmp!ID Then
                blnAdd = False
                Exit For
            End If
        Next
        '������������
        If gstr�������� <> "" Then
            If lngCount > Val(gstr��������) Then
                blnAdd = False
            End If
        End If
        
        If blnAdd Then
            lst�ɼ�������.AddItem "(" & rsTmp!���� & ")" & rsTmp!����
            lst�ɼ�������.ItemData(lst�ɼ�������.NewIndex) = rsTmp!ID
        End If
        rsTmp.MoveNext
    Loop
End Sub
