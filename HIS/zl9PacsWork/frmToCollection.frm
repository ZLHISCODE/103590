VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmToCollection 
   Caption         =   "��ӵ��ղ�"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmToCollection.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox PicButton 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      ScaleHeight     =   615
      ScaleWidth      =   4575
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4920
      Width           =   4575
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&S)"
         Height          =   360
         Left            =   2400
         TabIndex        =   2
         Top             =   120
         Width           =   990
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   360
         Left            =   3600
         TabIndex        =   1
         Top             =   120
         Width           =   990
      End
   End
   Begin MSComctlLib.TreeView tvwCollectionType 
      Height          =   4095
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   7223
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imgList"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   4080
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToCollection.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToCollection.frx":6BEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToCollection.frx":6F86
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgTree 
      Height          =   480
      Index           =   1
      Left            =   240
      Picture         =   "frmToCollection.frx":7320
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblPrompt 
      AutoSize        =   -1  'True
      Caption         =   "��ѡ����Ҫ�ղص���Ŀ¼"
      Height          =   195
      Left            =   960
      TabIndex        =   4
      Top             =   240
      Width           =   1980
   End
End
Attribute VB_Name = "frmToCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strSQL As String
Private mAdviceID As Long
Private mSendID As Long

Public Sub ShowToCollectionWind(Optional owner As Form = Nothing, Optional AdviceId As Long, Optional SendID As Long)
'��ʾ�ղع�����
    mAdviceID = AdviceId
    mSendID = SendID
    
    '����TreeView����
    Call LoadTreeView
    If tvwCollectionType.Nodes.Item(1).Children = 0 Then
        MsgBox "���ȵ��ղع����������ղ�Ŀ¼", , gstrSysName
        Exit Sub
    End If
 
    Call Me.Show(1, owner)
End Sub

Private Sub LoadTreeView()
'����TreeView���ݷ���
    Dim i As Long
    Dim objNode As Node
    Dim strSQL As String
    Dim rsTvwData As ADODB.Recordset
    
On Error GoTo errHand

    strSql = "select ID,�ϼ�ID,�ղ����,�Ƿ��� from Ӱ���ղ���� where ������ID= " & UserInfo.ID & " or ������ID is null Start With �ϼ�id Is Null Connect By Prior ID = �ϼ�id"
    Set rsTvwData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    With rsTvwData
        Me.tvwCollectionType.Nodes.Clear
        
        Do While Not .EOF
            If IsNull(!�ϼ�ID) Then
                Set objNode = Me.tvwCollectionType.Nodes.Add(, , "_" & Nvl(!ID), Nvl(!�ղ����), IIf(!�Ƿ��� = 0, 1, 3), IIf(Nvl(!�Ƿ���) = 0, 2, 3))
            Else
                Set objNode = Me.tvwCollectionType.Nodes.Add("_" & Nvl(!�ϼ�ID), tvwChild, "_" & Nvl(!ID), Nvl(!�ղ����), IIf(Nvl(!�Ƿ���) = 0, 1, 3), IIf(Nvl(!�Ƿ���) = 0, 2, 3))
            End If
            objNode.Sorted = True
            objNode.Expanded = True
            .MoveNext
        Loop
    End With
    
   Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub tvwCollectionType_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo errHand
'���Ϊ�����ڵ㣬�����ȷ����ť
    
 If Trim(Node.Text) = "�ղ����" Then
    cmdOK.Enabled = False
 Else
    cmdOK.Enabled = True
 End If
    
 Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub CmdOK_Click()
'ִ������ղز���
On Error GoTo errHand
Dim rsTemp As ADODB.Recordset
Dim dtServicesTime As String
Dim strSQL As String


    
     '�ж���ͬ�ղ������� �ղ������Ƿ��ظ�
     strSql = "select b.ҽ��id from Ӱ���ղ���� a,Ӱ���ղ����� b where a.id = b.�ղ�id and a.������ID= " & UserInfo.ID & " and a.�ղ����='" & Trim(tvwCollectionType.SelectedItem.Text) & "'"
     
     Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
     
     Do While Not rsTemp.EOF
        If Nvl(rsTemp!ҽ��ID) = mAdviceID Then
            Call MsgBoxD(Me, "�ü���ѱ�[ " & Trim(tvwCollectionType.SelectedItem.Text) & " ]�ղء�", vbOKOnly, Me.Caption)
            Exit Sub
        End If
        rsTemp.MoveNext
     Loop
    
    '��ǰ������ʱ��
    dtServicesTime = zlDatabase.Currentdate
     
    strSQL = "Zl_Ӱ���ղ�����_����(" & Val(Mid(tvwCollectionType.SelectedItem.Key, 2, 5)) & "," & mAdviceID & "," & zlStr.To_Date(dtServicesTime) & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
     
     '��ӳɹ� �رմ���
     Unload Me

 Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
'�رմ���
On Error GoTo errHand

    Unload Me

 Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    
    '��ʼ��ʱ����ȷ����ť
    cmdOK.Enabled = False
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    tvwCollectionType.Top = 660
    tvwCollectionType.Left = 80
    tvwCollectionType.Height = Me.ScaleHeight - PicButton.Height - 120
    tvwCollectionType.Width = Me.ScaleWidth - 160

    PicButton.Top = tvwCollectionType.Height + 140
    PicButton.Left = 0
    PicButton.Width = Me.ScaleWidth

    cmdOK.Left = PicButton.Width - cmdOK.Width - 1300
    cmdCancel.Left = PicButton.Width - cmdCancel.Width - 100
    
End Sub




