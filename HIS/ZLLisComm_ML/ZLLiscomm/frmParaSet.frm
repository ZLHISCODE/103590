VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmParaSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8865
   Icon            =   "frmParaSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtMicrobe 
      Height          =   270
      Left            =   7020
      TabIndex        =   49
      Top             =   3930
      Width           =   510
   End
   Begin VB.CheckBox chkQCCalc 
      Caption         =   "�����ʿ����ݺ��Ƿ�����ʿؼ���"
      Height          =   195
      Left            =   2775
      TabIndex        =   47
      Top             =   3960
      Width           =   3135
   End
   Begin VB.ComboBox cboAutoCheck 
      Height          =   300
      Left            =   4815
      Style           =   2  'Dropdown List
      TabIndex        =   45
      Top             =   3585
      Width           =   1890
   End
   Begin MSComDlg.CommonDialog dlgDir 
      Left            =   2085
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdPath 
      Caption         =   "��"
      Height          =   285
      Left            =   8355
      TabIndex        =   43
      Top             =   345
      Width           =   270
   End
   Begin VB.TextBox txtPath 
      Height          =   300
      Left            =   4110
      TabIndex        =   42
      ToolTipText     =   "����ָ�����ݽ��ճ����Ŀ¼"
      Top             =   375
      Width           =   4230
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "�������Ѻ��ձ걾(��ȡ������Ϣʱ��������ȡ�Ѻ��յı걾)"
      Height          =   195
      Left            =   2775
      TabIndex        =   41
      Top             =   3315
      Width           =   5595
   End
   Begin VB.TextBox txt��� 
      Height          =   270
      Left            =   3075
      TabIndex        =   39
      Top             =   2940
      Width           =   510
   End
   Begin VB.Frame fraSaveAs 
      Height          =   1440
      Left            =   2790
      TabIndex        =   35
      Top             =   4200
      Width           =   5880
      Begin VB.CheckBox chkTonDao 
         Alignment       =   1  'Right Justify
         Caption         =   "������б�������ȡͨ����(������������ϸ˵��)"
         Height          =   210
         Left            =   105
         TabIndex        =   48
         ToolTipText     =   $"frmParaSet.frx":000C
         Top             =   1065
         Width           =   4485
      End
      Begin VB.ComboBox cboSaveAs 
         Height          =   300
         Left            =   1815
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   180
         Width           =   3780
      End
      Begin VB.Label Label9 
         Caption         =   "���ݱ��浽ָ������"
         Height          =   210
         Left            =   105
         TabIndex        =   38
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "        �벻Ҫ������ģ�������ý����ڽ����������������յ������ݱ��浽��ָ����������"
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   450
         TabIndex        =   37
         Top             =   555
         Width           =   5115
      End
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "�Ƴ�(&M)"
      Height          =   350
      Left            =   1260
      TabIndex        =   32
      Top             =   5685
      Width           =   1100
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "����(&A)"
      Height          =   350
      Left            =   135
      TabIndex        =   31
      Top             =   5685
      Width           =   1100
   End
   Begin VB.CheckBox chkClear 
      Caption         =   "��ս�����־"
      Height          =   225
      Left            =   2910
      TabIndex        =   29
      Top             =   5730
      Width           =   1440
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7515
      TabIndex        =   28
      Top             =   5685
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6225
      TabIndex        =   27
      Top             =   5685
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   4545
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5685
      Width           =   1100
   End
   Begin TabDlg.SSTab sstbSet 
      Height          =   2040
      Left            =   2790
      TabIndex        =   0
      Top             =   810
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   3598
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "COMͨ������(&M)"
      TabPicture(0)   =   "frmParaSet.frx":0101
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkCom"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "TCP/IPͨ������(&T)"
      TabPicture(1)   =   "frmParaSet.frx":011D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraIP"
      Tab(1).Control(1)=   "ChkIP"
      Tab(1).ControlCount=   2
      Begin VB.Frame fraIP 
         Caption         =   "����"
         Height          =   1035
         Left            =   -74790
         TabIndex        =   14
         Top             =   855
         Width           =   5505
         Begin VB.ComboBox cboInMode 
            Height          =   300
            Left            =   4305
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   615
            Width           =   1080
         End
         Begin VB.OptionButton OptHost 
            Caption         =   "��Ϊ����"
            Height          =   255
            Index           =   0
            Left            =   2805
            TabIndex        =   20
            Top             =   225
            Width           =   1095
         End
         Begin VB.OptionButton OptHost 
            Caption         =   "��Ϊ�ն�"
            Height          =   225
            Index           =   1
            Left            =   1230
            TabIndex        =   19
            Top             =   225
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.TextBox txtPort 
            Height          =   300
            Left            =   2760
            MaxLength       =   5
            TabIndex        =   16
            Text            =   "66666"
            Top             =   615
            Width           =   630
         End
         Begin VB.TextBox txtIP 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   780
            MaxLength       =   15
            TabIndex        =   15
            Text            =   "0.0.0.0"
            Top             =   615
            Width           =   1500
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "����ģʽ"
            Height          =   255
            Left            =   3495
            TabIndex        =   24
            Top             =   660
            Width           =   735
         End
         Begin VB.Label lblPort 
            Alignment       =   1  'Right Justify
            Caption         =   "�˿�"
            Height          =   180
            Left            =   2025
            TabIndex        =   18
            Top             =   660
            Width           =   705
         End
         Begin VB.Label lblIP 
            Alignment       =   1  'Right Justify
            Caption         =   "����IP"
            Height          =   180
            Left            =   30
            TabIndex        =   17
            Top             =   660
            Width           =   690
         End
      End
      Begin VB.CheckBox ChkIP 
         Caption         =   "����TCP/IPͨ��"
         Height          =   240
         Left            =   -71100
         TabIndex        =   13
         Top             =   585
         Width           =   1680
      End
      Begin VB.CheckBox chkCom 
         Caption         =   "����COMͨ��"
         Height          =   240
         Left            =   4260
         TabIndex        =   12
         Top             =   450
         Width           =   1440
      End
      Begin VB.Frame Frame1 
         Caption         =   "�˿�����"
         Height          =   1335
         Left            =   105
         TabIndex        =   1
         Top             =   615
         Width           =   5640
         Begin VB.TextBox txtCom 
            Height          =   270
            Left            =   480
            TabIndex        =   33
            Top             =   600
            Width           =   510
         End
         Begin VB.ComboBox cboAttr 
            Height          =   300
            Index           =   9
            ItemData        =   "frmParaSet.frx":0139
            Left            =   4155
            List            =   "frmParaSet.frx":013B
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   990
            Width           =   1200
         End
         Begin VB.ComboBox cboAttr 
            Height          =   300
            Index           =   1
            Left            =   2100
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   255
            Width           =   1230
         End
         Begin VB.ComboBox cboAttr 
            Height          =   300
            Index           =   4
            ItemData        =   "frmParaSet.frx":013D
            Left            =   4155
            List            =   "frmParaSet.frx":013F
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   630
            Width           =   1215
         End
         Begin VB.ComboBox cboAttr 
            Height          =   300
            Index           =   3
            Left            =   2100
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   630
            Width           =   1230
         End
         Begin VB.ComboBox cboAttr 
            Height          =   300
            Index           =   2
            Left            =   4155
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   255
            Width           =   1215
         End
         Begin VB.ComboBox cboAttr 
            Height          =   300
            Index           =   5
            ItemData        =   "frmParaSet.frx":0141
            Left            =   2100
            List            =   "frmParaSet.frx":0151
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   990
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "COM"
            Height          =   180
            Left            =   135
            TabIndex        =   34
            Top             =   645
            Width           =   315
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "����ģʽ"
            Height          =   255
            Left            =   3390
            TabIndex        =   22
            Top             =   1035
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "�����ٶ�"
            Height          =   255
            Left            =   1260
            TabIndex        =   11
            Top             =   315
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "ֹͣλ"
            Height          =   285
            Left            =   3390
            TabIndex        =   10
            Top             =   675
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "��żλ"
            Height          =   285
            Left            =   1425
            TabIndex        =   9
            Top             =   675
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "����λ"
            Height          =   285
            Left            =   3390
            TabIndex        =   8
            Top             =   315
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "����Э��"
            Height          =   255
            Left            =   1260
            TabIndex        =   7
            Top             =   1035
            Width           =   735
         End
      End
   End
   Begin VB.ListBox Lst���� 
      Height          =   5280
      Left            =   90
      TabIndex        =   25
      Top             =   360
      Width           =   2565
   End
   Begin VB.Label lblMicrobe 
      AutoSize        =   -1  'True
      Caption         =   "΢�����ѯ       ���ڵ�����"
      Height          =   180
      Left            =   6060
      TabIndex        =   50
      ToolTipText     =   "��Ҫ�ӿڳ���֧�ֲŻᷢ�����"
      Top             =   3975
      Width           =   2520
   End
   Begin VB.Label lblAutoCheck 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������Զ���ˣ������                      (Ϊ�ղ������Զ����)"
      Height          =   180
      Left            =   2775
      TabIndex        =   46
      Top             =   3660
      Width           =   5760
   End
   Begin VB.Label Label12 
      Caption         =   "ͨѶ����Ŀ¼"
      Height          =   210
      Left            =   2880
      TabIndex        =   44
      Top             =   420
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "ÿ        ���Զ�Ӧ��ȡֵΪ0-3600,��Ϊ0����ʾ��ʹ�ô˹���)"
      Height          =   195
      Left            =   2775
      TabIndex        =   40
      ToolTipText     =   "��Ҫ�ӿڳ���֧�ֲŻᷢ�����"
      Top             =   2985
      Width           =   5715
   End
   Begin VB.Label lbl 
      Caption         =   "������������"
      Height          =   195
      Left            =   135
      TabIndex        =   30
      Top             =   75
      Width           =   1260
   End
End
Attribute VB_Name = "frmParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ifOK As Boolean
Private mblnEdit As Boolean '�Ƿ���Ȩ�޽����޸�

Private iLastDev As Long

Public Function ShowMe(objParent As Object) As Boolean
    Me.chkClear.Value = IIf(gblnClearData, 1, 0)
    Me.Show vbModal, objParent
    ShowMe = ifOK
End Function

Private Sub cboAttr_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call gobjCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkCom_Click()
    If chkCom.Value = 0 Then
        ChkIP.Value = 1
        sstbSet.Tab = 1
    Else
        ChkIP.Value = 0
    End If
End Sub

Private Sub ChkIP_Click()
    If ChkIP.Value = 0 Then
        chkCom.Value = 1
        sstbSet.Tab = 0
    Else
        chkCom.Value = 0
    End If
End Sub

Private Sub cmdAdd_Click()
    If frmSelect.Select���� Then
        iLastDev = -1
        LoadPropertySettings
        If Lst����.ListCount > 0 Then Lst����.ListIndex = 0
    End If
End Sub

Private Sub cmdCancel_Click()

    Unload Me
End Sub

Private Sub cmdDel_Click()
    Dim lngID As Long, i As Integer
    Dim lastIndex As Long
    Dim fsoTmp As New FileSystemObject
    Dim tsmTmp As TextStream, intTime As Integer
    
    If Lst����.ListCount <= 0 Then Exit Sub
    lngID = Lst����.ItemData(Lst����.ListIndex)
    If lngID > 0 Then

        For i = LBound(g����) To UBound(g����)
            If lngID = g����(i).ID Then
                g����(i).ID = 0
                Exit For
            End If
        Next
        
        If g����(i).ͨѶĿ¼ <> "" Then
            If fsoTmp.FolderExists(g����(i).ͨѶĿ¼) Then
                If MsgBox("�Ƿ������������ͨѶ��־��", vbYesNo + vbDefaultButton2, "��ʾ") = vbYes Then
                    If fsoTmp.FileExists(g����(i).ͨѶĿ¼ & "\Lock.txt") Then
                        Set tsmTmp = fsoTmp.CreateTextFile(g����(i).ͨѶĿ¼ & "\Send\CloseExe.txt")
                        tsmTmp.WriteLine Format(Now, "yyyy-MM-dd HH:mm:ss")
                        tsmTmp.Close
                        Set tsmTmp = Nothing
                    End If
                    intTime = 0
                    Do While intTime < 3000
                        If fsoTmp.FileExists(g����(i).ͨѶĿ¼ & "\Lock.txt") = False Then
                            fsoTmp.DeleteFolder g����(i).ͨѶĿ¼
                            Exit Do
                        End If
                        intTime = intTime + 1
                    Loop
                End If
            End If
        End If
        
        lastIndex = Lst����.ListIndex
        Lst����.RemoveItem lastIndex
        
        
        iLastDev = -1
        If lastIndex - 1 >= 0 Then
            Lst����.ListIndex = lastIndex - 1
        Else
            If Lst����.ListCount > 0 Then Lst����.ListIndex = 0
        End If
    End If
    
End Sub

Private Sub cmdHelp_Click()
    gobjComLib.ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer, blnNoDev As Boolean, strMsg As String, lng����ID As Long, str����ֵ As String
    Dim strIDs As String
    '����ǰ���ñ��浽�ڴ���

    If mblnEdit Then
        If Lst����.ListCount > 0 Then
            iLastDev = Lst����.ListIndex: Lst����_Click
        End If
        blnNoDev = True
        str����ֵ = ""
        strMsg = ""
        
        For i = LBound(g����) To UBound(g����)
            If g����(i).ID > 0 Then
    
                blnNoDev = False
                '������
                
                If g����(i).���� = 1 Then
                    'TCP/IP
                    
                    If ValidateIP(g����(i).IP) Then strMsg = strMsg & vbNewLine & g����(i).�������� & " IP����"
                    
                    If ValidatePort(g����(i).IP�˿�) Then strMsg = strMsg & vbNewLine & g����(i).�������� & " IP�˿ڴ���"
                    
                    If Not ValidateIP(g����(i).IP) And Not ValidatePort(g����(i).IP�˿�) Then
                        If InStr(str����ֵ, "," & g����(i).IP & ":" & g����(i).IP�˿�) > 0 Then
                            strMsg = strMsg & vbNewLine & g����(i).�������� & " IP��ַ�Ͷ˿��ظ�����"
                        Else
                            str����ֵ = str����ֵ & "," & g����(i).IP & ":" & g����(i).IP�˿�
                        End If
                    End If
                Else
                    'COM
                    If g����(i).COM�� = 0 Then
                        strMsg = strMsg & vbNewLine & g����(i).�������� & " COM�����ô���"
                    Else
                        If InStr(str����ֵ, ",COM" & g����(i).COM��) > 0 Then
                            strMsg = strMsg & vbNewLine & g����(i).�������� & " COM���ظ�����"
                        Else
                            str����ֵ = str����ֵ & ",COM" & g����(i).COM��
                        End If
                    End If
                End If
                
                If Val(g����(i).�Զ�Ӧ��) < 0 Or Val(g����(i).�Զ�Ӧ��) > 3600 Then
                    strMsg = strMsg & vbNewLine & g����(i).�������� & " �Զ�Ӧ��ʱ����0 - 3600��֮��"
                End If
                If txtMicrobe <> "" Then
                    If Val(txtMicrobe) < 0 Or Val(txtMicrobe) > 365 Then
                        strMsg = strMsg & "΢����������ѯ���ֻ������365��"
                    End If
                End If
                If Trim(g����(i).ͨѶĿ¼) = "" Then
                    strMsg = strMsg & vbNewLine & g����(i).�������� & " ͨѶĿ¼���ò���ȷ"
                End If
            End If
        Next
        
        If strMsg <> "" Then
            MsgBox "���������������飺" & strMsg, vbQuestion
            Exit Sub
        End If
        
        If blnNoDev Then
            If MsgBox("û�������κ�������ϵͳ�����ܽ��ռ������ݣ��Ƿ������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                Lst����.SetFocus: Exit Sub
            End If
        Else
            If MsgBox("ϵͳ���������Ӽ������������ݽ��չ��̽���ͣ���Ƿ������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                Lst����.SetFocus: Exit Sub
            End If
        End If
        SavePortsSetting
    End If
    If txtMicrobe <> "" Then
        If Val(txtMicrobe) > 0 And Val(txtMicrobe) <= 365 Then
            Call gobjDatabase.SetPara("΢�����ѯʱ��", Val(txtMicrobe), glngSys, 1208)
        End If
    End If
    If gblnFromDB Then
        Call gobjDatabase.SetPara("��ս�����־", Me.chkClear.Value, glngSys, 1208)
    Else
        Call SaveSetting("ZLSOFT", "����ģ��\ZlLISSrv", "��ս�����־", CStr(Me.chkClear.Value))
    End If

    ifOK = True
    Unload Me
End Sub

Private Sub cmdPath_Click()
    Dim strResFolder As String
    strResFolder = BrowseForFolder(hwnd, "��ѡ��һ��Ŀ¼.")
    If strResFolder <> "" Then
        txtPath.Text = strResFolder
    End If
     
End Sub

Private Sub Form_Activate()
    Dim objControl As Object
    mblnEdit = InStr(";" & gstrPrivs & ";", ";ͨѶ��������;") > 0

    If Not mblnEdit Then
        For Each objControl In Me.Controls
            If InStr("chkClear,cmdHelp,cmdOK,cmdCancel,lvwComm,sstbSet", objControl.Name) > 0 Then
                objControl.Enabled = True
            Else
                If InStr("dlgDir", objControl.Name) <= 0 Then objControl.Enabled = False
            End If
        Next
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    Call cmdCancel_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    ifOK = False
    mblnEdit = False

    iLastDev = -1
    LoadPropertySettings
    If Lst����.ListCount > 0 Then Lst����.ListIndex = 0
    
End Sub

Private Sub LoadPropertySettings()
    Dim rsDev As adodb.Recordset
    Dim strSQL As String
    On Error GoTo hErr
    '���봮�������趨---������
    Dim i As Integer
    With cboAttr(1)
        .AddItem "110"
        .AddItem "300"
        .AddItem "600"
        .AddItem "1200"
        .AddItem "2400"
        .AddItem "4800"
        .AddItem "9600"
        .AddItem "14400"
        .AddItem "19200"
        .AddItem "28800"
        .AddItem "38400"
        .AddItem "56000"
        .AddItem "57600"
        .AddItem "115200"
        .AddItem "128000"
        .AddItem "256000"
    End With
    
    ' ��������λ����
    With cboAttr(2)
        .AddItem "4"
        .AddItem "5"
        .AddItem "6"
        .AddItem "7"
        .AddItem "8"
    End With
    
    ' ������ż��������
    With cboAttr(3)
        .AddItem "None"
        .AddItem "Odd"
        .AddItem "Even"
        .AddItem "Mark"
        .AddItem "Space"
    End With
    
    ' ����ֹͣλ����
    With cboAttr(4)
        .AddItem "1"
        .AddItem "1.5"
        .AddItem "2"
    End With
    '
    
    With cboAttr(9) '����ģʽ
        .Clear
        .AddItem "�ַ�"
        .AddItem "��ģʽ"
    End With
    
    With cboInMode
        .Clear
        .AddItem "�ַ�"
        .AddItem "��ģʽ"
    End With
    
    '��������
    Set rsDev = GetDevices
'    With cboAttr(0)
'        .Clear
'        .AddItem "δָ���豸"
'        .ItemData(0) = 0

        cboSaveAs.Clear
        cboSaveAs.AddItem "ȱʡ"
        cboSaveAs.ItemData(0) = 0

    If Not rsDev Is Nothing Then
        Do While Not rsDev.EOF
'                .AddItem "(" & rsDev("����") & ")" & rsDev("����")
'                .ItemData(.ListCount - 1) = rsDev("ID")

            cboSaveAs.AddItem "(" & rsDev("����") & ")" & rsDev("����")
            cboSaveAs.ItemData(cboSaveAs.ListCount - 1) = rsDev("ID")
    
            rsDev.MoveNext
        Loop
    End If
    Lst����.Clear
    For i = LBound(g����) To UBound(g����)
       If g����(i).ID > 0 Then
           rsDev.Filter = "ID=" & g����(i).ID
           If Not rsDev.EOF Then
               Lst����.AddItem "(" & rsDev("����") & ")" & rsDev("����")
               Lst����.ItemData(Lst����.ListCount - 1) = rsDev("ID")
           End If
       End If
    Next
    
    With cboAutoCheck
        .Clear
        .AddItem ""
        strSQL = "Select Distinct b.���� From ����С���Ա a, ��Ա�� b Where a.��Աid = b.Id Order By b.����"
        Set rsDev = gobjDatabase.OpenSqlRecord(strSQL, "ȡ�����Ա")
        Do Until rsDev.EOF
            .AddItem "" & rsDev!����
            rsDev.MoveNext
        Loop
        .ListIndex = 0
    End With
    
    txtMicrobe = gobjDatabase.GetPara("΢�����ѯʱ��", 100, 1208, 0)
    
'    End With
    Exit Sub
hErr:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Sub


Private Sub Lst����_Click()
    Dim lng����ID As Long
    Dim i As Integer, intTmp As Integer
    On Error GoTo errH
    
    If iLastDev > -1 Then
        lng����ID = Val(Lst����.ItemData(iLastDev))
        
         For i = LBound(g����) To UBound(g����)
            If Val(g����(i).ID) = lng����ID Then
                '�����޸�
                g����(i).IP = txtIP
                g����(i).IP�˿� = CLng(Val(txtPort))
                g����(i).SaveAsID = Val(cboSaveAs.ItemData(cboSaveAs.ListIndex))
                g����(i).������ = CLng(Val(cboAttr(1).Text))
                g����(i).����λ = cboAttr(2).Text
                g����(i).���� = ChkIP.Value
                g����(i).COM�� = CInt(Val(txtCom))
                g����(i).У��λ = Left(cboAttr(3).Text, 1)
                g����(i).ֹͣλ = cboAttr(4).Text
                g����(i).���� = cboAttr(5).ListIndex
                g����(i).���� = IIf(OptHost(0).Value, 1, 0)
                g����(i).�ַ�ģʽ = IIf(chkCom.Value = 1, cboAttr(9).ListIndex, cboInMode.ListIndex)
                If IsNumeric(Trim(Me.txt���.Text)) Then
                    g����(i).�Զ�Ӧ�� = Trim(txt���.Text)
                End If
                g����(i).�ɷ��Ѻ˱걾 = Val(chk����.Value)
                g����(i).ͨѶĿ¼ = Trim(txtPath.Text)
                g����(i).�Զ������ = Trim(cboAutoCheck.Text)
                g����(i).�Զ������ʿ� = Val(chkQCCalc.Value)
                g����(i).���Ϊͨ���� = Val(chkTonDao.Value)
                Exit For
            End If
        Next
    End If
    lng����ID = Val(Lst����.ItemData(Lst����.ListIndex))
    
    If lng����ID > 0 Then
        For i = LBound(g����) To UBound(g����)
            
            If Val(g����(i).ID) = lng����ID Then
                
                If g����(i).���� = 0 Then
                    txtCom = g����(i).COM��
                    ChkIP.Value = 0
                    chkCom.Value = 1
                    sstbSet.Tab = 0
                    Me.cboAttr(1).Text = g����(i).������
                    Me.cboAttr(2).Text = g����(i).����λ
                    Me.cboAttr(3).Text = Switch(UCase(g����(i).У��λ) = "N", "None", _
                        UCase(g����(i).У��λ) = "E", "Even", _
                        UCase(g����(i).У��λ) = "O", "Odd", _
                        UCase(g����(i).У��λ) = "M", "Mark", _
                        UCase(g����(i).У��λ) = "S", "Space")
                    Me.cboAttr(4).Text = g����(i).ֹͣλ
                    Me.cboAttr(5).ListIndex = Val(g����(i).����)

                Else
                    txtCom = g����(i).COM��
                    ChkIP.Value = 1
                    chkCom.Value = 0
                    sstbSet.Tab = 1
                                    
                    txtPort = g����(i).IP�˿�
                    txtIP = g����(i).IP
                    OptHost(0).Value = g����(i).���� = 1
                    
                    If OptHost(0).Value Then
                        Call OptHost_Click(1)
                    Else
                        Call OptHost_Click(0)
                    End If
                End If
                Me.cboAttr(9).ListIndex = Val(g����(i).�ַ�ģʽ)
                cboInMode.ListIndex = Val(g����(i).�ַ�ģʽ)
                Me.txt���.Text = CStr(g����(i).�Զ�Ӧ��)
                If Left(Me.txt���, 1) = "." Then Me.txt���.Text = "0" & Me.txt���.Text
                
                Me.cboSaveAs.ListIndex = GetComboxIndex(cboSaveAs, g����(i).SaveAsID)
                Me.chk����.Value = g����(i).�ɷ��Ѻ˱걾
                Me.txtPath = g����(i).ͨѶĿ¼
                cboAutoCheck.ListIndex = 0
                
                If Trim(g����(i).�Զ������) <> "" Then
                    For intTmp = 0 To cboAutoCheck.ListCount - 1
                        If cboAutoCheck.List(intTmp) = g����(i).�Զ������ Then
                            cboAutoCheck.ListIndex = intTmp
                            Exit For
                        End If
                    Next
                End If
                
                Me.chkQCCalc.Value = g����(i).�Զ������ʿ�
                Me.chkTonDao.Value = g����(i).���Ϊͨ����
            End If
        Next
        
    End If
    iLastDev = Lst����.ListIndex
    Exit Sub
errH:
    MsgBox Err.Description
End Sub

Private Sub OptHost_Click(Index As Integer)
    If Index = 0 Then
        lblIP.Caption = "����IP"
        lblPort.Caption = "�˿�"
    Else
        lblIP.Caption = "����IP"
        lblPort.Caption = "�˿�"
    End If
End Sub


Private Sub txtMicrobe_KeyPress(KeyAscii As Integer)
    Dim lngTag As Long
    Dim strTmp As String
    Dim lngDay As Long
    lngTag = FilterKeyAscii(KeyAscii, 1)
    KeyAscii = lngTag
    
    strTmp = Mid(txtMicrobe.Text, txtMicrobe.SelStart + 1, txtMicrobe.SelLength)
    lngDay = Val(Replace(txtMicrobe.Text, strTmp, "") & Chr(KeyAscii))
    
    If lngDay > 365 Then
        MsgBox "���������������365�죬����!", vbInformation, "����������ʾ"
        KeyAscii = 0
        Exit Sub
        
    End If
End Sub

Public Function FilterKeyAscii(ByVal KeyAscii As Long, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Long

    FilterKeyAscii = KeyAscii

    If Chr(KeyAscii) = "'" Then
        FilterKeyAscii = 0
        Exit Function
    End If

    If KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or KeyAscii = vbKeyBack Then
        Exit Function
    End If

    Select Case bytMode
    Case 1      '������
        If InStr("0123456789<>", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 2      '��С��
        If InStr("0123456789.-<>+Ee", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 99
        If InStr(KeyCustom, Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    End Select
End Function
