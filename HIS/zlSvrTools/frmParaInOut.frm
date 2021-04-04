VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmParaInOut 
   Caption         =   "�������뵼��ѡ��"
   ClientHeight    =   3630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6060
   Icon            =   "frmParaInOut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6060
   StartUpPosition =   1  '����������
   Begin VB.Frame fraSplit 
      BackColor       =   &H80000012&
      Height          =   30
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   3000
      Width           =   6700
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   6060
      TabIndex        =   0
      Top             =   3015
      Width           =   6060
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   4515
         TabIndex        =   2
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   3360
         TabIndex        =   1
         Top             =   120
         Width           =   1100
      End
   End
   Begin VB.PictureBox picSet 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2925
      Left            =   0
      ScaleHeight     =   2925
      ScaleWidth      =   6075
      TabIndex        =   3
      Top             =   0
      Width           =   6075
      Begin VB.CommandButton cmdFile 
         Caption         =   "��"
         Height          =   255
         Left            =   5355
         TabIndex        =   11
         Top             =   203
         Width           =   300
      End
      Begin VB.Frame fraHos 
         Caption         =   "ҽԺѡ��"
         Height          =   1215
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   5535
         Begin VB.OptionButton optCurHos 
            Caption         =   "��Ժ"
            Height          =   255
            Left            =   360
            TabIndex        =   9
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optOtherHos 
            Caption         =   "��Ժ"
            Height          =   180
            Left            =   1560
            TabIndex        =   8
            Top             =   390
            Width           =   855
         End
         Begin VB.Label lblInfo 
            Caption         =   "���������嵥������˽�в������á����Ų������á�"
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   360
            TabIndex        =   14
            Top             =   720
            Width           =   4935
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame fraSys 
         Caption         =   "ϵͳѡ��"
         Height          =   735
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   5535
         Begin VB.OptionButton optAllSys 
            Caption         =   "����ϵͳ"
            Height          =   180
            Left            =   1560
            TabIndex        =   6
            Top             =   390
            Width           =   1095
         End
         Begin VB.OptionButton optCurSys 
            Caption         =   "��ǰϵͳ"
            Height          =   255
            Left            =   360
            TabIndex        =   5
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.TextBox txtFile 
         Height          =   300
         Left            =   1080
         MaxLength       =   256
         TabIndex        =   12
         Top             =   180
         Width           =   4575
      End
      Begin VB.Label lblPath 
         AutoSize        =   -1  'True
         Caption         =   "����·��"
         Height          =   180
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   720
      End
   End
   Begin MSComDlg.CommonDialog cmmFile 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmParaInOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mpstType As ParaSetType '0-���롣1-������2-Excelת��XMl,3-XML�ϲ�
Public Enum ParaSetType
    PST_Imp = 0
    PST_Exp = 1
End Enum
Private mstrReturn As String
Private mlngSys As Long '�����浱ǰ��ϵͳ

Public Function ShowMe(ByVal pstType As ParaSetType, Optional ByVal lngSys As Long) As String
    mpstType = pstType
    mlngSys = lngSys
    mstrReturn = ""
    Me.Show vbModal
    ShowMe = mstrReturn
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFile_Click()
    cmmFile.FileName = txtFile.Text
    If mpstType = PST_Imp Then
        cmmFile.Filter = "�����ļ�(*.xml)|*.xml"
        cmmFile.ShowOpen
    Else
        cmmFile.Filter = "�����ļ�(*.xml)|*.xml"
        cmmFile.ShowSave
    End If
    If cmmFile.FileName <> "" Then
        If mpstType = PST_Imp Then
            If CheckImpFile(cmmFile.FileName, True) Then
                txtFile.Text = cmmFile.FileName
            End If
        Else
            txtFile.Text = cmmFile.FileName
        End If
    End If
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    mstrReturn = txtFile.Text & "|" & IIf(optCurSys.value, 0, 1) & "|" & IIf(optCurHos.value, 0, 1)
    If txtFile.Text <> "" Then
        Call SaveSetting("ZLSOFT", "�û�����", "�������뵼��", txtFile.Text)
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strPath As String
    picSet.Visible = False
    Me.Width = 6300: Me.Height = 4200
    Select Case mpstType
        Case PST_Exp
            Me.Caption = "��������": lblPath.Caption = "�����ļ�"
            picSet.Visible = True
            If gblnInIDE Then
                txtFile.Text = gobjFile.GetFile("C:\APPSOFT\zlSvrStudio.exe").ParentFolder & "\ZLParasInfo.xml"
            Else
                txtFile.Text = GetSetting("ZLSOFT", "�û�����", "�������뵼��", App.Path & "\ZLParasInfo.xml")
            End If
        Case PST_Imp
            Me.Caption = "��������": lblPath.Caption = "�����ļ�"
            picSet.Visible = True
            If gblnInIDE Then
                txtFile.Text = gobjFile.GetFile("C:\APPSOFT\zlSvrStudio.exe").ParentFolder & "\ZLParasInfo.xml"
            Else
                txtFile.Text = GetSetting("ZLSOFT", "�û�����", "�������뵼��", App.Path & "\ZLParasInfo.xml")
            End If
            If Not CheckImpFile(txtFile.Text) Then
                txtFile.Text = ""
            End If
    End Select
    Call optCurHos_Click
End Sub

Private Sub Form_Resize()
    Me.Width = 6300
    Me.Height = 4200
End Sub

Private Sub optCurHos_Click()
    If mpstType = PST_Exp Then
        lblInfo.Caption = "���������嵥������˽�в������á����Ų������á�"
    Else
        lblInfo.Caption = "��������嵥������˽�в������á����Ų������á�"
    End If
End Sub

Private Sub optOtherHos_Click()
    If mpstType = PST_Exp Then
        lblInfo.Caption = "���������嵥"
    Else
        lblInfo.Caption = "��������嵥"
    End If
End Sub

Private Sub picBottom_Resize()
    cmdCancel.Left = picBottom.ScaleWidth - 120 - cmdCancel.Width
    cmdOK.Left = cmdCancel.Left - 60 - cmdOK.Width
End Sub

Private Function CheckImpFile(ByVal strFile As String, Optional blnMsg As Boolean) As Boolean
'���ܣ���鵼���ļ�
'������strFile=�����ļ�
'          blnMsg=�Ƿ񵯳���Ϣ��ʾ
'���أ��Ƿ���ͨ��
    Dim rsParas As ADODB.Recordset, rsDBSys As ADODB.Recordset, rsComInfo As ADODB.Recordset
    Dim lngSys As Long, blnDetial As Boolean
    
    On Error GoTo errH
    If Dir(strFile) = "" Then Exit Function
    Set rsParas = New ADODB.Recordset
    '��ȡ��������
    rsParas.Open strFile, , adOpenStatic, adLockOptimistic, adCmdFile
    
    optOtherHos.Enabled = True: optCurHos.Enabled = True
    optCurSys.Enabled = True: optAllSys.Enabled = True
    rsParas.Filter = "���� = -99" '������Ϣ
    If rsParas.EOF Then
        If blnMsg Then MsgBox "�ò����ļ���δ������Ч��������Ϣ���޷����룡", vbInformation, gstrSysName
        Exit Function
    End If
    blnDetial = Val(rsParas!˽��) <> 0: lngSys = Val(rsParas!������)
    '�ж��Ƿ��пɵ����ϵͳ
    rsParas.Filter = "����=-9"
    Set rsDBSys = GetALLPars(-9)
    '���� ������, �汾�� ����ֵ, User ȱʡֵ,To_Char(Sysdate, 'yyyy-mm-dd HH24:mi:ss') Ӱ�����˵��
    Set rsComInfo = GetCompareRec(rsDBSys, rsParas, "ϵͳ", "����ֵ")
    rsComInfo.Filter = "State=0 OR State=2"
    If rsComInfo.RecordCount = 0 Then 'û�п��Ե����ϵͳ
        If blnMsg Then MsgBox "�ò����ļ���δ���ֿɵ����ϵͳ��", vbInformation, gstrSysName
        Exit Function
    End If
    If Not blnDetial Then '���ļ�δ�������Ų������飬����˽�в������飬ֻ��ѡ����Ժ����
        optOtherHos.value = True
        optOtherHos.Enabled = False: optCurHos.Enabled = False
    End If
    If lngSys <> -1 Then
        If lngSys <> mlngSys Then '������ϵͳ�Ĳ����ļ�
            optAllSys.value = True
        Else '�ǵ�ǰϵͳ�Ĳ����ļ�
            optCurSys.value = True
        End If
        optOtherHos.Enabled = False: optCurHos.Enabled = False
    Else '���ϵͳ�Ĳ����ļ�
        rsComInfo.Filter = "State<>-1 And MainKey='" & mlngSys & "'" '�鿴�����ļ����Ƿ��е�ǰϵͳ�Ĳ���
        If rsComInfo.EOF Then '�����ڵ�ǰϵͳ�Ĳ���������ѡ��ǰϵͳ
            optAllSys.value = True
            optOtherHos.Enabled = False: optCurHos.Enabled = False
        End If
    End If
    CheckImpFile = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Function
