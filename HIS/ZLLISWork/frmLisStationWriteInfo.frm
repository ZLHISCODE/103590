VERSION 5.00
Begin VB.Form frmLisStationWriteInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ٴ�����"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5820
   Icon            =   "frmLisStationWriteInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5820
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txt�ٴ����� 
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   150
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   150
      Width           =   5355
   End
End
Attribute VB_Name = "frmLisStationWriteInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mlngItemID As Long

Private Sub Form_Load()
    GetInfo mlngItemID
End Sub

Private Sub Form_Resize()
    With Me.txt�ٴ�����
        .Top = 0
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With
End Sub

Public Sub ShowMe(objfrm As Object, lngItemID As Long)
    mlngItemID = lngItemID
    Me.Show , objfrm
End Sub

Private Sub GetInfo(lngItemID As Long)
    Dim rsTmp As New ADODB.Recordset
    gstrSql = "Select B.����, B.������, B.Ӣ����,a.�ٴ����� From ������Ŀ A, ����������Ŀ B Where A.������Ŀid = B.ID And B.ID = [1]"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    If rsTmp.EOF = False Then
        Me.txt�ٴ�����.Text = Nvl(rsTmp("����")) & "-" & Nvl(rsTmp("������")) & IIf(Nvl(rsTmp("Ӣ����")) <> "", "(" & Nvl(rsTmp("Ӣ����")) & ")", "") & _
                            vbCrLf & "  " & Nvl(rsTmp("�ٴ�����"))
    Else
        Me.txt�ٴ�����.Text = ""
    End If
End Sub

Public Sub SelectItem(lngItemID As Long)
    If Me.Visible = True Then
        GetInfo lngItemID
    End If
End Sub

