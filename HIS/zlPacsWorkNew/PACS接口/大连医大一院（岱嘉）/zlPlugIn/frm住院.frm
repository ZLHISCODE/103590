VERSION 5.00
Begin VB.Form frmסԺ 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   750
   ClientLeft      =   6300
   ClientTop       =   0
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   937.5
   ScaleMode       =   0  'User
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt��Ա��� 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   465
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   60
      Width           =   735
   End
   Begin VB.TextBox txt�ܷ��� 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   60
      Width           =   750
   End
   Begin VB.TextBox txt��� 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2925
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   60
      Width           =   750
   End
   Begin VB.TextBox txt��ϸ 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   60
      Width           =   1965
   End
   Begin VB.TextBox txt˵�� 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   435
      Width           =   1965
   End
   Begin VB.TextBox txt���� 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "ȫ�Źؽ��û�����˫�ࣩ"
      Top             =   435
      Width           =   2115
   End
   Begin VB.TextBox txt��� 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   930
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "������"
      Top             =   450
      Width           =   555
   End
   Begin VB.Label lab���� 
      AutoSize        =   -1  'True
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   210
      Left            =   255
      TabIndex        =   12
      Top             =   450
      Width           =   450
   End
   Begin VB.Label lab��Ա��� 
      AutoSize        =   -1  'True
      Caption         =   "��Ա"
      Height          =   180
      Left            =   90
      TabIndex        =   11
      Top             =   60
      Width           =   360
   End
   Begin VB.Line Line1 
      DrawMode        =   1  'Blackness
      X1              =   465
      X2              =   1240
      Y1              =   356.25
      Y2              =   356.25
   End
   Begin VB.Line Line2 
      DrawMode        =   1  'Blackness
      X1              =   1650
      X2              =   2460
      Y1              =   356.25
      Y2              =   356.25
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   1260
      TabIndex        =   10
      Top             =   60
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "���"
      Height          =   180
      Left            =   2505
      TabIndex        =   9
      Top             =   60
      Width           =   360
   End
   Begin VB.Line Line3 
      DrawMode        =   1  'Blackness
      X1              =   2925
      X2              =   3735
      Y1              =   356.25
      Y2              =   356.25
   End
   Begin VB.Line Line4 
      DrawMode        =   1  'Blackness
      X1              =   3975
      X2              =   5945
      Y1              =   356.25
      Y2              =   356.25
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   180
      Left            =   3780
      TabIndex        =   8
      Top             =   60
      Width           =   180
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      X1              =   -15
      X2              =   8530
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0080FFFF&
      X1              =   -15
      X2              =   8530
      Y1              =   468.75
      Y2              =   468.75
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   180
      Left            =   3780
      TabIndex        =   7
      Top             =   495
      Width           =   180
   End
   Begin VB.Line lin���� 
      BorderColor     =   &H00000000&
      X1              =   1650
      X2              =   3795
      Y1              =   843.75
      Y2              =   843.75
   End
   Begin VB.Line Line8 
      X1              =   3975
      X2              =   6120
      Y1              =   825
      Y2              =   825
   End
End
Attribute VB_Name = "frmסԺ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngPatiID          As Long
Private mvarRecId           As Variant
Private mvarKeyId           As Variant
Private mstrReserve         As String

Const col������ = &HFF&
Const col��ͨ�� = vbBlack
Const col���Բ� = &HFF0000
Const col���ֲ� = &HFF00FF

Private Type typ_������Ϣ
    str����                 As String
    str���                 As String
    str����                 As String
    str˵��                 As String
    color                   As Long
End Type
Private var����             As typ_������Ϣ

Const con���ݿɱ�����       As Double = 8000
Dim rsTmp                   As ADODB.Recordset

Public Property Let PatiID(ByVal vNewValue As Long)
    mlngPatiID = vNewValue
End Property

Public Property Let RecId(ByVal vNewValue As Variant)
    mvarRecId = vNewValue
End Property

Public Property Let KeyId(ByVal vNewValue As Variant)
    mvarKeyId = vNewValue
End Property

Public Property Let Reserve(ByVal vNewValue As String)
    mstrReserve = vNewValue
End Property

Public Sub RefreshData()
    Dim rtn                 As Long
    Dim rsSum               As ADODB.Recordset
    Dim dbl�ܷ���           As Double
    Dim dbl��������ܷ���   As Double
    Dim dbl�����ܶ�         As Double
    Dim lng����ID           As Long
    Dim intInsure           As Integer
    
    DoEvents
    Me.Show
    rtn = SetWindowPos(Me.hWnd, -1, CurrentX, CurrentY, 0, 0, 3)
    
    '����Ƿ��������
    gstrSql = "select 1 from ���������¼ A,����_�������� B where A.������ĿID = B.����ID And A.����ID=[1] AND A.��ҳID=[2]"
    lab����.Visible = Not ChkRsState(gDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngPatiID, Val(mvarRecId)))
    '��ȡ��Ա���
    gstrSql = "select A.����,A.��ְ,B.���� from �����ʻ� A ,������Ⱥ B where A.��ְ=B.��� AND A.����=B.���� And A.����ID=[1]"
    Set rsTmp = gDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngPatiID)
    If ChkRsState(rsTmp) Then
        txt��Ա���.Text = ""
        Me.Hide
    Else
        intInsure = rsTmp!����
        txt��Ա���.Text = rsTmp!����
        If ChkRsState(rsTmp) Then
            Me.Height = 370
        Else
            '��ȡ������Ϣ
            gstrSql = "select C.ID,B.����,DECODE(C.���,1,'���Բ�',2,'���ֲ�',3,'������','��ͨ��') AS ���,B.����,B.˵��" & vbCrLf & _
                  "from (Select ������� From ������ϼ�¼ where ID IN (select Max(ID) as ID from ������ϼ�¼ where �������=1 AND ����ID = [1] And ��ҳID = [2] group by ����ID,��ҳID )) A,��������Ŀ¼ B,���ղ��� C" & vbCrLf & _
                  "where zl_split(zl_split(A.�������,')',0),'(',1)=B.���� AND B.����=C.����(+)"
            Set rsTmp = gDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngPatiID, Val(mvarRecId))
            
            If ChkRsState(rsTmp) Then
                Me.Height = 370
                '��ʹ�÷���
                gstrSql = "select nvl(sum(ʵ�ս��),0) as ��� from סԺ���ü�¼ where ����ID = [1] And ��ҳID = [2]"
                dbl�ܷ��� = gDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngPatiID, Val(mvarRecId)).Fields(0)
                '��ȡ�����޶�
                gstrSql = "select  nvl(���ƽ��,0) from ����_�����޶� where ����ID in (select nvl(��Ժ����ID,��Ժ����ID) " & _
                          "from ������ҳ where ����ID=[1] And ��ҳID=[2])"
                Set rsTmp = gDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngPatiID, Val(mvarRecId))
                If ChkRsState(rsTmp) Then
                    dbl�����ܶ� = 0
                Else
                    dbl�����ܶ� = rsTmp.Fields(0)
                End If
                '�ܷ���
                txt�ܷ���.Text = Format(dbl�ܷ���, "0.00")
                txt��ϸ.Text = "  �ƣ�" & Format(dbl�����ܶ�, "0")
                txt���.Text = Format(dbl�����ܶ� + dbl��������ܷ��� - dbl�ܷ���, "0.00")
                txt���.ForeColor = IIf(Val(txt���.Text) < 0, col������, col���Բ�)
            Else
                Me.Height = 730
                var����.color = Decode(rsTmp!���, "���Բ�", col���Բ�, "���ֲ�", col���ֲ�, "������", col������, col��ͨ��)
                var����.str���� = "" & rsTmp!����
                var����.str��� = "" & rsTmp!���
                var����.str���� = "" & rsTmp!����
                var����.str˵�� = "" & rsTmp!˵��
                txt���.ForeColor = var����.color
                txt���.Text = var����.str���
                txt����.Text = var����.str����
                txt˵��.Text = var����.str˵��
                '�����޶�
                lng����ID = Val("" & rsTmp!ID)
                '��⵱ǰ�����Ƿ����޶�
                gstrSql = "select nvl(���ƽ��,0) as ��� from ����_�����޶� where ����=[1] And ����ID=[2]"
                Set rsTmp = gDatabase.OpenSQLRecord(gstrSql, Me.Caption, intInsure, lng����ID)
                '��ʹ�÷���
                gstrSql = "select nvl(sum(ʵ�ս��),0) as ��� from סԺ���ü�¼ where ����ID = [1] And ��ҳID = [2]"
                dbl�ܷ��� = gDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngPatiID, Val(mvarRecId)).Fields(0)
                If ChkRsState(rsTmp) Then
                    '��ȡ�����޶�
                    gstrSql = "select  nvl(���ƽ��,0) from ����_�����޶� where ����ID in (select nvl(��Ժ����ID,��Ժ����ID) " & _
                              "from ������ҳ where ����ID=[1] And ��ҳID=[2])"
                    Set rsTmp = gDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngPatiID, Val(mvarRecId))
                    If ChkRsState(rsTmp) Then
                        dbl�����ܶ� = 0
                    Else
                        dbl�����ܶ� = rsTmp.Fields(0)
                    End If
                    '�ܷ���
                    txt�ܷ���.Text = Format(dbl�ܷ���, "0.00")
                    txt��ϸ.Text = "  �ƣ�" & Format(dbl�����ܶ�, "0")
                    txt���.Text = Format(dbl�����ܶ� + dbl��������ܷ��� - dbl�ܷ���, "0.00")
                    txt���.ForeColor = IIf(Val(txt���.Text) < 0, col������, col���Բ�)
                Else
                    '��ȡ�����޶�
                    dbl�����ܶ� = rsTmp!���
                    gstrSql = "select nvl(sum(���ƽ��),0) as ��� from סԺ���ü�¼ A,����_���ֲ��� B " & _
                              "where A.�շ�ϸĿID = B.�շ�ID And  ����ID = [2] And ��ҳID = [3]  And ����=[1] And B.����ID=[4]"
                    dbl��������ܷ��� = gDatabase.OpenSQLRecord(gstrSql, Me.Caption, intInsure, mlngPatiID, Val(mvarRecId), lng����ID).Fields(0)
                    '�ܷ���
                    txt�ܷ���.Text = Format(dbl�ܷ���, "0.00")
                    txt��ϸ.Text = "  ����" & Format(dbl�����ܶ�, "0") & ";�أ�" & Format(dbl��������ܷ���, "0")
                    txt���.Text = Format(dbl�����ܶ� - dbl�ܷ���, "0.00")
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Resize()
    Me.Top = 0
    Me.Left = 6300
End Sub
 
