VERSION 5.00
Begin VB.Form frm��¼����_�����山 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ѡ��"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310
   Icon            =   "frm��¼����_�����山.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fra1 
      Caption         =   "��Ժ������Ϣ"
      Height          =   1395
      Left            =   60
      TabIndex        =   7
      Top             =   1065
      Width           =   7980
      Begin VB.TextBox Txt���� 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1095
         TabIndex        =   2
         Top             =   240
         Width           =   6780
      End
      Begin VB.TextBox Txt���� 
         Height          =   300
         Index           =   1
         Left            =   1095
         TabIndex        =   4
         Top             =   615
         Width           =   6780
      End
      Begin VB.TextBox Txt���� 
         Height          =   300
         Index           =   2
         Left            =   1095
         TabIndex        =   6
         Top             =   975
         Width           =   6780
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����(&1)"
         Height          =   180
         Index           =   0
         Left            =   420
         TabIndex        =   1
         Top             =   300
         Width           =   630
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����һ(&2)"
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   690
         Width           =   810
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "���ֶ�(&3)"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   810
      End
   End
   Begin VB.Frame fra2 
      Caption         =   "��Ժ������Ϣ"
      Height          =   1395
      Left            =   60
      TabIndex        =   8
      Top             =   2595
      Width           =   7965
      Begin VB.TextBox Txt���� 
         Height          =   300
         Index           =   3
         Left            =   1080
         TabIndex        =   10
         Top             =   270
         Width           =   6780
      End
      Begin VB.TextBox Txt���� 
         Height          =   300
         Index           =   4
         Left            =   1080
         TabIndex        =   12
         Top             =   645
         Width           =   6780
      End
      Begin VB.TextBox Txt���� 
         Height          =   300
         Index           =   5
         Left            =   1080
         TabIndex        =   14
         Top             =   1005
         Width           =   6780
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����(&4)"
         Height          =   180
         Index           =   3
         Left            =   405
         TabIndex        =   9
         Top             =   330
         Width           =   630
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����һ(&5)"
         Height          =   180
         Index           =   4
         Left            =   225
         TabIndex        =   11
         Top             =   720
         Width           =   810
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "���ֶ�(&6)"
         Height          =   180
         Index           =   5
         Left            =   225
         TabIndex        =   13
         Top             =   1110
         Width           =   810
      End
   End
   Begin VB.Frame fra 
      Height          =   105
      Index           =   0
      Left            =   15
      TabIndex        =   18
      Top             =   720
      Width           =   8475
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5790
      TabIndex        =   16
      Top             =   4335
      Width           =   1100
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7050
      TabIndex        =   17
      Top             =   4335
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   105
      Left            =   -150
      TabIndex        =   15
      Top             =   4110
      Width           =   8715
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   165
      Picture         =   "frm��¼����_�����山.frx":000C
      Stretch         =   -1  'True
      Top             =   210
      Width           =   510
   End
   Begin VB.Label lblPatient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�������山ҽ��    ���ţ�01234567    "
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   210
      Left            =   960
      TabIndex        =   0
      Top             =   450
      Width           =   7275
   End
End
Attribute VB_Name = "frm��¼����_�����山"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mblnStart As Boolean
Private mintInsure As Integer
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mbln��¼ As Boolean
Private Sub Txt����_Change(Index As Integer)
    Txt����(Index).Tag = ""
End Sub
Private Sub Txt����_GotFocus(Index As Integer)
        'ҽ��Ҫ��������ϱ�������
        zlControl.TxtSelAll Txt����(Index)
End Sub

Private Sub Txt����_KeyPress(Index As Integer, KeyAscii As Integer)
  Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim strLike As String, str�Ա� As String
    Dim StrInput As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Txt����(Index).Text = "" Or Txt����(Index).Tag <> "" Then
            Call zlCommFun.PressKey(vbKeyTab) '��������
        Else
            StrInput = UCase(Txt����(Index).Text)
            gstrSQL = "" & _
            "   Select id, ����, ����, ֧�����, ������, ���ֽ���취, ���칹������ " & _
            "   From ҽ������Ŀ¼" & _
            "   Where ����=2 and (" & zlCommFun.GetLike("", "����", StrInput) & " Or " & _
                        zlCommFun.GetLike("", "����", StrInput) & " Or " & _
                        zlCommFun.GetLike("", "������", StrInput) & ") "
            
            Dim sngLeft As Single, sngTop As Single
            
            If Index >= 3 Then
                sngLeft = Txt����(Index).Left + Me.Left + fra2.Left
                sngTop = Txt����(Index).Top + Me.Top + fra2.Top
            Else
                sngLeft = Txt����(Index).Left + Me.Left + fra1.Left
                sngTop = Txt����(Index).Top + Me.Top + fra1.Top
            End If
            
            Set rsTmp = zlDatabase.ShowSelect(Me, gstrSQL, 0, "���ֱ���ѡ��", , , , , , True, _
                sngLeft, sngTop, Txt����(Index).Height, blnCancel, , True)
            If Not rsTmp Is Nothing Then
                Txt����(Index).Text = "(" & rsTmp!���� & ")" & rsTmp!����
                Txt����(Index).Tag = Nvl(rsTmp!ID)
                lbl����(Index).Tag = Nvl(rsTmp!����)
                If Index < 5 Then
                   If Txt����(Index + 1).Enabled Then Txt����(Index + 1).SetFocus
                Else
                   If cmdOK.Enabled Then cmdOK.SetFocus
                End If
            Else
                If Not blnCancel Then
                    MsgBox "û���ҵ�ƥ��Ĳ��ֱ��롣", vbInformation, gstrSysName
                End If
                Call Txt����_GotFocus(Index)
                Txt����(Index).SetFocus
            End If
        End If
    Else
        zlControl.TxtCheckKeyPress Txt����(Index), KeyAscii, m�ı�ʽ
    End If
End Sub
Public Sub Load��ʷ������Ϣ()
    '����:������ʷҽ�����˵Ŀ�����Ϣ
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0: On Error GoTo errHand:
    
    gstrSQL = "Select ����,���,����ID,�������,���� from ����������_91 where  ����id=" & mlng����ID & IIf(mlng��ҳID = 0, " and ��ҳid is null ", " and ��ҳid=" & mlng��ҳID) & " and ���� IN (1,2)"
    Call OpenRecordset_OtherBase(rsTemp, "��ȡ������", gstrSQL, gcnOracle_CQYB)
    
    With rsTemp
        Do While Not .EOF
            If Val(Nvl(!����)) = 1 Then
                Select Case Nvl(!���, 0)
                Case 1
                    Txt����(0).Text = IIf(IsNull(!�������), "", "(" & Nvl(!�������) & ")") & Nvl(!����)
                    Txt����(0).Tag = Val(Nvl(!����ID))
                    lbl����(0).Tag = Nvl(!�������)
                Case 2
                    Txt����(1).Text = IIf(IsNull(!�������), "", "(" & Nvl(!�������) & ")") & Nvl(!����)
                    Txt����(1).Tag = Val(Nvl(!����ID))
                    lbl����(1).Tag = Nvl(!�������)
                Case 3
                    Txt����(2).Text = IIf(IsNull(!�������), "", "(" & Nvl(!�������) & ")") & Nvl(!����)
                    Txt����(2).Tag = Val(Nvl(!����ID))
                    lbl����(2).Tag = Nvl(!�������)
                    
                End Select
            Else
                Select Case Nvl(!���, 0)
                Case 1
                    Txt����(3).Text = IIf(IsNull(!�������), "", "(" & Nvl(!�������) & ")") & Nvl(!����)
                    Txt����(3).Tag = Val(Nvl(!����ID))
                    lbl����(3).Tag = Nvl(!�������)
                    
                Case 2
                    Txt����(4).Text = IIf(IsNull(!�������), "", "(" & Nvl(!�������) & ")") & Nvl(!����)
                    Txt����(4).Tag = Val(Nvl(!����ID))
                    lbl����(4).Tag = Nvl(!�������)
                Case 3
                    Txt����(5).Text = IIf(IsNull(!�������), "", "(" & Nvl(!�������) & ")") & Nvl(!����)
                    Txt����(5).Tag = Val(Nvl(!����ID))
                    lbl����(5).Tag = Nvl(!�������)
                End Select
            End If
           .MoveNext
        Loop
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub






Private Sub cmdOK_Click()
    
     '����ѡ������Ϣ
    If Trim(Txt����(0).Text) = "" Then
        MsgBox "��Ϊ�òα�����ѡ����Ժ���飡", vbInformation, gstrSysName
        If Txt����(0).Enabled Then Txt����(0).SetFocus
        Exit Sub
    End If
    If Trim(Txt����(3).Text) = "" Then
        MsgBox "��Ϊ�òα�����ѡ���Ժ���飡", vbInformation, gstrSysName
        If Txt����(3).Enabled Then Txt����(3).SetFocus
        Exit Sub
    End If
    
    '���没��
    Err = 0: On Error GoTo errHand
    'gcnOracle.BeginTrans
    gcnOracle_CQYB.BeginTrans
    '������Ժ����
    Call Get������Ϣ(False)
    Call Save������Ϣ(mlng����ID, mlng��ҳID, 1)
    
    Call Get������Ϣ(True)
    Call Save������Ϣ(mlng����ID, mlng��ҳID, 2)
    
    '1.���浽�����ʻ�
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & mlng����ID & "," & mintInsure & ",'����ID','''" & g�������_�����山.����ID & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ID")
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & mlng����ID & "," & mintInsure & ",'����1ID','" & IIf(g�������_�����山.����1ID = 0, "NULL", g�������_�����山.����1ID) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ID")
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & mlng����ID & "," & mintInsure & ",'����2ID','" & IIf(g�������_�����山.����2ID = 0, "NULL", g�������_�����山.����2ID) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ID")
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & mlng����ID & "," & mintInsure & ",'����1','''" & g�������_�����山.�������� & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ID")
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & mlng����ID & "," & mintInsure & ",'����2','''" & g�������_�����山.��������1 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ID")
    gstrSQL = "ZL_�����ʻ�_������Ϣ(" & mlng����ID & "," & mintInsure & ",'����3','''" & g�������_�����山.��������2 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ID")
   ' gcnOracle.CommitTrans
    gcnOracle_CQYB.CommitTrans
    mblnOK = True
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
  '  gcnOracle.RollbackTrans
    gcnOracle_CQYB.CRollbackTrans
End Sub

Private Sub Get������Ϣ(ByVal bln��Ժ As Boolean)
    '����:��ȡ������Ϣ
    Dim i As Integer
    i = IIf(bln��Ժ, 3, 0)
    
    g�������_�����山.����ID = Val(Txt����(i).Tag)
    If g�������_�����山.����ID <> 0 Then
        g�������_�����山.������� = Trim(lbl����(i).Tag)
        g�������_�����山.�������� = Replace(Trim(Txt����(i).Text), "(" & Trim(lbl����(i).Tag) & ")", "", 1, 1)
    Else
        g�������_�����山.������� = ""
        g�������_�����山.�������� = Trim(Txt����(i).Text)
    End If
    i = i + 1
    g�������_�����山.����1ID = Val(Txt����(i).Tag)
    If g�������_�����山.����1ID <> 0 Then
        g�������_�����山.�������1 = Trim(lbl����(i).Tag)
        g�������_�����山.��������1 = Replace(Trim(Txt����(i).Text), "(" & Trim(lbl����(i).Tag) & ")", "", 1, 1)
    Else
        g�������_�����山.�������1 = ""
        g�������_�����山.��������1 = Trim(Txt����(i).Text)
    End If

    i = i + 1
    g�������_�����山.����2ID = Val(Txt����(i).Tag)
    If g�������_�����山.����2ID <> 0 Then
        g�������_�����山.�������2 = Trim(lbl����(i).Tag)
        g�������_�����山.��������2 = Replace(Trim(Txt����(i).Text), "(" & Trim(lbl����(i).Tag) & ")", "", 1, 1)
    Else
        g�������_�����山.�������2 = ""
        g�������_�����山.��������2 = Trim(Txt����(i).Text)
    End If
End Sub

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If Not mblnStart Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = " Select B.����,A.����,A.ҽ���� " & _
              " From �����ʻ� A,������Ϣ B " & _
              " Where A.����ID=B.����ID And A.����ID=[1] And A.����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ò��˵Ļ�����Ϣ", mlng����ID, mintInsure)
    If rsTemp.EOF Then
        mblnStart = False
        MsgBox "ҽ�����˲�����!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    lblPatient.Caption = "������" & Nvl(rsTemp!����) & Space(4) & "���ţ�" & Nvl(rsTemp!����) & Space(4) & "���˱�ţ�" & Nvl(rsTemp!ҽ����)
    
    Call Load��ʷ������Ϣ
    
    Err = 0: On Error Resume Next
    fra1.Enabled = mbln��¼
    Txt����(0).Enabled = mbln��¼
    Txt����(1).Enabled = mbln��¼
    Txt����(2).Enabled = mbln��¼
    
    mblnStart = True
End Sub

Public Function ShowSelect(ByVal intinsure As Integer, _
    ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional bln��¼ As Boolean = False) As Boolean
    'ѡ���˵���Ժ���鼰��Ժ���飬ͬʱ�����˱���סԺ�������Ϣ��ʾ����
    '���±����ʻ��Ĳ���ID����Ժ���飩����Ժ���飬������Ժ���鼰��Ժ������뷵�ظ�����ģ��
    
    mblnOK = False
    mintInsure = intinsure
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mbln��¼ = bln��¼
    Me.Show 1
    ShowSelect = mblnOK
End Function


