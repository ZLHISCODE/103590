VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmManualArchive 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�ֶ��鵵"
   ClientHeight    =   3630
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin TabDlg.SSTab sstabManualArchive 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   6376
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "��һ��"
      TabPicture(0)   =   "frmManualArchive.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdStep1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "�ڶ���"
      TabPicture(1)   =   "frmManualArchive.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdStep2"
      Tab(1).Control(1)=   "cobManualMoveDelete"
      Tab(1).Control(2)=   "Label4(2)"
      Tab(1).Control(3)=   "Label4(1)"
      Tab(1).Control(4)=   "Label4(0)"
      Tab(1).Control(5)=   "Label3"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "������"
      TabPicture(2)   =   "frmManualArchive.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lblDetail"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdStep3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.CommandButton cmdStep2 
         Caption         =   "��һ��"
         Height          =   350
         Left            =   -71520
         TabIndex        =   12
         Top             =   3100
         Width           =   1100
      End
      Begin VB.ComboBox cobManualMoveDelete 
         Height          =   300
         ItemData        =   "frmManualArchive.frx":0054
         Left            =   -73920
         List            =   "frmManualArchive.frx":0061
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2520
         Width           =   2415
      End
      Begin VB.CommandButton cmdStep3 
         Caption         =   "��ʼ�鵵"
         Height          =   350
         Left            =   3360
         TabIndex        =   7
         Top             =   3120
         Width           =   1100
      End
      Begin VB.CommandButton cmdStep1 
         Caption         =   "��һ��"
         Height          =   350
         Left            =   -71500
         TabIndex        =   6
         Top             =   3100
         Width           =   1100
      End
      Begin VB.Frame Frame1 
         Caption         =   "ѡ��洢�豸"
         Height          =   1815
         Left            =   -74760
         TabIndex        =   2
         Top             =   1200
         Width           =   4335
         Begin VB.ComboBox cobDevice 
            Height          =   315
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1200
            Width           =   2520
         End
         Begin VB.OptionButton optManualSelectDevice 
            Caption         =   "�ֹ�ָ���洢�豸"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   4
            Top             =   720
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.OptionButton optManualSelectDevice 
            Caption         =   "�Զ�ѡ��洢�豸"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   3
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.Label Label4 
         Caption         =   "3��ֻɾ�������ƶ����ݣ�����ɾ��Դ�豸�е�����"
         Height          =   495
         Index           =   2
         Left            =   -74760
         TabIndex        =   15
         Top             =   1800
         Width           =   4215
      End
      Begin VB.Label Label4 
         Caption         =   "2���鵵��ɾ���������ݴ�Դ�豸�ƶ���Ŀ���豸��Ȼ��ɾ��Դ�豸������"
         Height          =   495
         Index           =   1
         Left            =   -74760
         TabIndex        =   14
         Top             =   1320
         Width           =   4215
      End
      Begin VB.Label Label4 
         Caption         =   "1��ֻ�鵵�������ݴ�Դ�豸�ƶ���Ŀ���豸��ͬʱ����Դ�豸������"
         Height          =   495
         Index           =   0
         Left            =   -74760
         TabIndex        =   13
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label Label3 
         Caption         =   "�鵵��ʽѡ��"
         Height          =   375
         Left            =   -74760
         TabIndex        =   10
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label Label2 
         Caption         =   "ȷ�Ϲ鵵���ã�"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblDetail 
         Caption         =   "Դ�豸��"
         Height          =   1335
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   4455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ѡ��洢�豸����Ĭ������£�ϵͳ���Զ�ѡ��һ����ѹ鵵Ŀ���豸"
         Height          =   360
         Left            =   -74760
         TabIndex        =   1
         Top             =   600
         Width           =   4260
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmManualArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bArchive As Boolean          '��ʶ�鵵/���鵵 .1---�鵵;0---���鵵

Private Sub Command1_Click()
    
End Sub

Private Sub cmdStep1_Click()
    
    Me.sstabManualArchive.Tab = 1
End Sub

Private Sub Command3_Click()

End Sub

Private Sub cmdStep2_Click()
    Me.sstabManualArchive.Tab = 2
End Sub

Private Sub cmdStep3_Click()

    On Error GoTo errH
    '��ʼ�鵵
    If Me.cmdStep3.Caption = "���" Then
        Unload Me
    Else
        '�����ݿ�鵵��ҵ�����һ���鵵��ҵ��¼
        Dim strSQL As String
        Dim tmpset As ADODB.Recordset
        Dim lngJobNum As Long
        
        If Me.cobDevice.ListCount <= 0 Then
            MsgBox "û��Ŀ���豸�������豸���á�"
            Exit Sub
        End If
        Dim strDeviceNo As String
        strDeviceNo = Left(Me.cobDevice.List(Me.cobDevice.ListIndex), InStr(Me.cobDevice.List(Me.cobDevice.ListIndex), "-") - 1)
        
        strSQL = "select Ӱ��鵵��ҵ_ID.nextval as JobID from dual"
        Set tmpset = gcnOracle.Execute(strSQL)
        lngJobNum = tmpset!JobID
        strSQL = "Insert into Ӱ��鵵��ҵ (����,����,ִ��ʱ��,Դ�豸,Ŀ���豸,ָ���豸,�Ƿ�Ǩ��,�Ƿ�ɾ��,�Զ�����,ִ�й���) values (" & _
                 lngJobNum & ",'�ֶ��鵵" & lngJobNum & "',to_date('" & Date & " " & Time & "','yyyy-mm-dd hh24:mi:ss') " & _
                 IIf(bArchive, ",'1','2','", ",'2','1','") & IIf(Me.optManualSelectDevice(1).Value = True, "", strDeviceNo) & "' ," & _
                 IIf(Me.cobManualMoveDelete.ListIndex <> 2, 1, 0) & "," & _
                 IIf(Me.cobManualMoveDelete.ListIndex <> 0, 1, 0) & ",0,0)"
        gcnOracle.Execute (strSQL)
        
        zl9comlib.ZlCommFun.ShowFlash
        'ִ�й鵵��ҵ
        frmMain.funcdoArchiveJob lngJobNum
        zl9comlib.ZlCommFun.StopFlash
        Me.cmdStep3.Caption = "���"
        frmMain.ShowChkRecord
    End If
    Exit Sub
errH:
    zl9comlib.ZlCommFun.StopFlash
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optManualSelectDevice_Click(Index As Integer)
    If Index = 1 Then       '�Զ�ѡ�񱸷��豸
        Me.cobDevice.Enabled = False
    Else                    '�ֹ�ָ�������豸
        Me.cobDevice.Enabled = True
    End If
End Sub

Private Sub sstabManualArchive_Click(PreviousTab As Integer)
    If Me.sstabManualArchive.Tab = 2 Then              '���ҳ��
        Me.lblDetail.Caption = "Դ�豸��" & IIf(bArchive = True, "���洢�豸", "�����洢�豸") & vbCrLf & vbCrLf & _
                               "Ŀ���豸��" & IIf(Me.optManualSelectDevice(1).Value = True, "�Զ�ѡ��", Me.cobDevice.Text) & vbCrLf & vbCrLf & _
                               "�鵵��" & IIf(Me.cobManualMoveDelete.ListIndex <> 2, "��", "��") & vbCrLf & vbCrLf & _
                               "ɾ����" & IIf(Me.cobManualMoveDelete.ListIndex <> 0, "��", "��")
        
    End If
End Sub

 
