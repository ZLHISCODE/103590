VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTechnicSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   Icon            =   "frmTechnicSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox PicAction 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6055
      Left            =   120
      ScaleHeight     =   6030
      ScaleWidth      =   6585
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.ComboBox cbxMoneyExeModle 
         Height          =   300
         ItemData        =   "frmTechnicSetup.frx":000C
         Left            =   1410
         List            =   "frmTechnicSetup.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   4935
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Frame frmPatholParameter 
         Height          =   3855
         Left            =   120
         TabIndex        =   24
         Top             =   0
         Width           =   3495
         Begin VB.TextBox txtMoleculeReport 
            Height          =   270
            Left            =   1680
            TabIndex        =   38
            Top             =   2595
            Width           =   1335
         End
         Begin VB.TextBox txtSpecialStainReport 
            Height          =   270
            Left            =   1680
            TabIndex        =   37
            Top             =   2235
            Width           =   1335
         End
         Begin VB.TextBox txtImmuneReport 
            Height          =   270
            Left            =   1680
            TabIndex        =   36
            Top             =   1875
            Width           =   1335
         End
         Begin VB.TextBox txtNormalReport 
            Height          =   270
            Left            =   1680
            TabIndex        =   35
            Top             =   1515
            Width           =   1335
         End
         Begin VB.TextBox txtGrossDescribe 
            Height          =   270
            Left            =   1680
            TabIndex        =   34
            Top             =   1155
            Width           =   1335
         End
         Begin VB.CheckBox chkIsDirectPrint 
            Caption         =   "�Ƿ�ֱ�Ӵ�ӡ"
            Height          =   180
            Left            =   120
            TabIndex        =   33
            Top             =   840
            Width           =   1575
         End
         Begin VB.CheckBox chkKuaiPian 
            Caption         =   "��Ƭ"
            Height          =   375
            Left            =   2760
            TabIndex        =   32
            Top             =   3330
            Width           =   495
         End
         Begin VB.CheckBox chkShiJian 
            Caption         =   "ʬ��"
            Height          =   375
            Left            =   2280
            TabIndex        =   31
            Top             =   3330
            Width           =   495
         End
         Begin VB.CheckBox chkHuiZhen 
            Caption         =   "����"
            Height          =   375
            Left            =   1800
            TabIndex        =   30
            Top             =   3330
            Width           =   495
         End
         Begin VB.CheckBox chkXiBao 
            Caption         =   "ϸ��"
            Height          =   375
            Left            =   1320
            TabIndex        =   29
            Top             =   3330
            Width           =   495
         End
         Begin VB.CheckBox chkBingDong 
            Caption         =   "����"
            Height          =   375
            Left            =   840
            TabIndex        =   28
            Top             =   3330
            Width           =   495
         End
         Begin VB.CheckBox chkChangGui 
            Caption         =   "����"
            Height          =   375
            Left            =   360
            TabIndex        =   27
            Top             =   3330
            Width           =   495
         End
         Begin VB.TextBox txtDecalinHintTime 
            Height          =   270
            Left            =   1800
            TabIndex        =   26
            Text            =   "30"
            Top             =   500
            Width           =   495
         End
         Begin VB.CheckBox chkDecalin 
            Caption         =   "�Ѹ�������������"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            ToolTipText     =   "���Ѹ�ʱ�䵽�˻���������ʾ��"
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblSpecialStainReport 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ⱦ����ģ�壺"
            Height          =   180
            Left            =   360
            TabIndex        =   46
            Top             =   2280
            Width           =   1260
         End
         Begin VB.Label lblMoleculeReport 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ӱ���ģ�壺"
            Height          =   180
            Left            =   360
            TabIndex        =   45
            Top             =   2640
            Width           =   1260
         End
         Begin VB.Label lblImmuneReport 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���߱���ģ�壺"
            Height          =   180
            Left            =   360
            TabIndex        =   44
            Top             =   1920
            Width           =   1260
         End
         Begin VB.Label lblNormalReport 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���汨��ģ�壺"
            Height          =   180
            Left            =   360
            TabIndex        =   43
            Top             =   1560
            Width           =   1260
         End
         Begin VB.Label lblGrossDescribe 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�޼�����ģ�壺"
            Height          =   180
            Left            =   360
            TabIndex        =   42
            Top             =   1200
            Width           =   1260
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ӧ�ʾ����"
            Height          =   180
            Left            =   1800
            TabIndex        =   41
            Top             =   900
            Width           =   1080
         End
         Begin VB.Label labHint 
            Caption         =   "�����¼�����ʱ�Զ������������ڣ�"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   3060
            Width           =   3135
         End
         Begin VB.Label Label3 
            Caption         =   "���Ѽ��ʱ����      ��"
            Height          =   255
            Left            =   600
            TabIndex        =   39
            Top             =   550
            Width           =   2055
         End
      End
      Begin VB.ComboBox cboExecuteRooms 
         Height          =   300
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3560
         Width           =   1935
      End
      Begin VB.CommandButton CmdDevSet 
         Caption         =   "�豸����(&S)"
         Height          =   350
         Left            =   1440
         TabIndex        =   20
         Top             =   5550
         Width           =   1170
      End
      Begin VB.CheckBox ChkOpenReport 
         Caption         =   "��ʼ�����Զ��򿪱��洰��"
         Height          =   255
         Left            =   3840
         TabIndex        =   19
         ToolTipText     =   "�ڱ������Զ��򿪱��洰�ڡ�"
         Top             =   3065
         Width           =   2640
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   5355
         TabIndex        =   18
         Top             =   5550
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   4095
         TabIndex        =   17
         Top             =   5550
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   180
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   5550
         Width           =   1100
      End
      Begin VB.CheckBox chkPatTrack 
         Caption         =   "����״̬����"
         Height          =   255
         Left            =   3840
         TabIndex        =   15
         ToolTipText     =   "�ڶԲ���һϵ�еļ������У�ʼ�ձ��ֵ�ǰѡ�е���ͬһ�����ˡ�"
         Top             =   2500
         Width           =   2640
      End
      Begin VB.CheckBox chkBatchInput 
         Caption         =   "������������"
         Height          =   255
         Left            =   3840
         TabIndex        =   14
         ToolTipText     =   "���������Ǽǡ�"
         Top             =   805
         Width           =   2640
      End
      Begin VB.CheckBox chkView 
         Caption         =   "��д����ʱ�򿪹�Ƭվ"
         Height          =   255
         Left            =   3840
         TabIndex        =   13
         ToolTipText     =   "�򿪱�����д����ʱ�Զ��򿪹�Ƭվ��"
         Top             =   1935
         Width           =   2280
      End
      Begin VB.Frame Frame6 
         Height          =   30
         Index           =   0
         Left            =   0
         TabIndex        =   12
         Top             =   5400
         Width           =   6615
      End
      Begin VB.CheckBox chkCancelCheck 
         Caption         =   "����ʾ��ȡ���ĵǼ�"
         Height          =   255
         Left            =   3840
         TabIndex        =   11
         ToolTipText     =   "�ڲ��˼���б��в���ʾ�Ѿ���ȡ���ĵǼǼ�¼��"
         Top             =   1370
         Width           =   2640
      End
      Begin VB.CommandButton cmd3DSetup 
         Caption         =   "3D����"
         Height          =   350
         Left            =   2760
         TabIndex        =   10
         Top             =   5550
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.CheckBox chkAutoPrint 
         Caption         =   "�������Զ���ӡ���뵥"
         Height          =   255
         Left            =   3840
         TabIndex        =   9
         ToolTipText     =   "���˱������Զ���ӡ���뵥��"
         Top             =   240
         Width           =   2100
      End
      Begin VB.CheckBox chkExitAfterSign 
         Caption         =   "ǩ�����˳�"
         Height          =   255
         Left            =   3840
         TabIndex        =   8
         ToolTipText     =   "����ǩ�����Զ��˳�������д��"
         Top             =   3630
         Width           =   1320
      End
      Begin VB.Frame Frame2 
         Caption         =   "�Ǽ�ģʽ"
         Height          =   880
         Left            =   120
         TabIndex        =   3
         Top             =   3900
         Width           =   3495
         Begin VB.CheckBox chkInputOutInfo 
            Caption         =   "¼����Ժ��Ϣ"
            Height          =   255
            Left            =   1710
            TabIndex        =   49
            ToolTipText     =   "�ڵǼǴ���¼���ͼ쵥λ���ͼ�ҽ����"
            Top             =   865
            Width           =   1590
         End
         Begin VB.OptionButton optCheckInMode 
            Caption         =   "����ģʽ"
            Height          =   255
            Index           =   1
            Left            =   135
            TabIndex        =   7
            ToolTipText     =   "ֻ��ʾ��¼���Ҫ��Ŀ��"
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optCheckInMode 
            Caption         =   "����ģʽ"
            Height          =   180
            Index           =   2
            Left            =   135
            TabIndex        =   6
            ToolTipText     =   "��ʾ��¼��ȫ����Ŀ��"
            Top             =   570
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.CheckBox chkAddons 
            Caption         =   "����ʾ��������"
            Height          =   255
            Left            =   1710
            TabIndex        =   5
            ToolTipText     =   "�ڵǼǱ������ڲ���ʾ��������һ�"
            Top             =   540
            Width           =   1590
         End
         Begin VB.CheckBox chkReagent 
            Caption         =   "����ʾ��Ӱ��"
            Height          =   255
            Left            =   1710
            TabIndex        =   4
            ToolTipText     =   "�ڵǼǱ������ڲ���ʾ��Ӱ��һ�"
            Top             =   225
            Width           =   1680
         End
      End
      Begin VB.ComboBox cbxMainPage 
         Height          =   300
         ItemData        =   "frmTechnicSetup.frx":0041
         Left            =   3840
         List            =   "frmTechnicSetup.frx":0043
         TabIndex        =   2
         Top             =   4970
         Width           =   2655
      End
      Begin VB.CheckBox chkStartVideoCapture 
         Caption         =   "����ı�ɼ������С"
         Height          =   255
         Left            =   3840
         TabIndex        =   1
         ToolTipText     =   "�����û��ı���Ƶͼ��ɼ�����ĳߴ��С��"
         Top             =   4200
         Value           =   1  'Checked
         Width           =   2400
      End
      Begin MSComDlg.CommonDialog dlgFont 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label labMoneyExemodel 
         Caption         =   "����ִ��ģʽ��"
         Height          =   270
         Left            =   120
         TabIndex        =   47
         Top             =   4980
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblExecuteRoomName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ִ�м����ƣ�"
         Height          =   180
         Left            =   240
         TabIndex        =   23
         Top             =   3600
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "��Ҫ����ҳ�棺"
         Height          =   180
         Left            =   3840
         TabIndex        =   21
         Top             =   4680
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmTechnicSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Public mlng����ID As Long 'IN:��ǰִ�п���ID
Public mblnOK As Boolean
Public mlngModul As Long
Public mstrPrivs As String 'ģ��Ȩ��

'�����б���������
Private mTitleFont As StdFont       '�����б��ͷ����
Private mTextFont As StdFont        '�����б���������


Private Sub cmd3DSetup_Click()
    frm3DSetup.ShowMe Me, mstrPrivs
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdDevSet_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1101)
End Sub

Private Sub CmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub InitExecuteRooms()
    Dim rsExecute As New ADODB.Recordset

    Dim strSql As String

    strSql = "select ִ�м�,�豸�� from ҽ��ִ�з��� a, Ӱ���豸Ŀ¼ b Where a.����豸=b.�豸��(+) and ����ID=[1]"
    Set rsExecute = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID)

    '��տؼ�
    cboExecuteRooms.Clear

    If rsExecute.RecordCount <= 0 Then Exit Sub

    'ѭ������ִ�м���
    Do While Not rsExecute.EOF
        cboExecuteRooms.AddItem rsExecute!ִ�м� & "-" & Nvl(rsExecute!�豸��)
        rsExecute.MoveNext
    Loop

    If cboExecuteRooms.ListCount > 0 Then cboExecuteRooms.ListIndex = 0
End Sub

Private Sub cmdOK_Click()
On Error GoTo errHandle
    Dim strPar As String, i As Long
    '��������
    Dim GrossNum As Long, NormalNum As Long, ImmuneNum As Long, SpecialNum As Long, MoleculeNum As Long
    Dim strSql As String
    Dim rsExpression As ADODB.Recordset

    

    zlDatabase.SetPara "������ҳ", cbxMainPage.Text, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0

    zlDatabase.SetPara "����ʱ��Ƭ", chkView.value, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
    zlDatabase.SetPara "�����Ǽ�����", chkBatchInput.value, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
    zlDatabase.SetPara "���˸���", chkPatTrack.value, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
    
    zlDatabase.SetPara "��ʼ����Զ��򿪱���", ChkOpenReport.value, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
    zlDatabase.SetPara "����ʾ��Ӱ��", chkReagent.value, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
    zlDatabase.SetPara "����ʾ��������", chkAddons.value, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
    zlDatabase.SetPara "¼����Ժ��Ϣ", chkInputOutInfo.value, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
    zlDatabase.SetPara "�Ǽ�ģʽ", IIf(optCheckInMode(1).value = True, 1, 2), glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
    zlDatabase.SetPara "����ʾ��ȡ���ĵǼ�", chkCancelCheck.value, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
    zlDatabase.SetPara "�������Զ���ӡ���뵥", chkAutoPrint.value, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
    
    zlDatabase.SetPara "����ִ�м�����", cboExecuteRooms.Text, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
    
    
    Call zlDatabase.SetPara("PACS����ǩ�����˳�", chkExitAfterSign.value, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0)
    
    If mlngModul = 1291 Then
        Call zlDatabase.SetPara("�ɼ�����ִ��ģʽ", cbxMoneyExeModle.ListIndex, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0)
    End If
    
    If mlngModul <> 1290 Then
        Call zlDatabase.SetPara("����ı�ɼ������С", chkStartVideoCapture.value, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0)
    End If
    
    
    If mlngModul = G_LNG_PATHOLSYS_NUM Then
        '���没��ϵͳ��ز���
        zlDatabase.SetPara "�Ѹ���������", chkDecalin.value, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
        zlDatabase.SetPara "���Ѽ��ʱ��", txtDecalinHintTime.Text, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
        zlDatabase.SetPara "������������", chkChangGui.value, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
        zlDatabase.SetPara "����ʯ����������", chkKuaiPian.value, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
        zlDatabase.SetPara "������������", chkBingDong.value, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
        zlDatabase.SetPara "ϸ����������", chkXiBao.value, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
        zlDatabase.SetPara "������������", chkHuiZhen.value, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
        zlDatabase.SetPara "ʬ����������", chkShiJian.value, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
        
        '�����Ƿ�ֱ�Ӵ�ӡ����
        zlDatabase.SetPara "�Ƿ�ֱ�Ӵ�ӡ", chkIsDirectPrint.value, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
        
        strSql = "select ���� from �����ʾ����"
        Set rsExpression = zlDatabase.OpenSQLRecord(strSql, Me.Caption)

        For i = 1 To rsExpression.RecordCount
            
            If txtGrossDescribe.Text <> "" Then
                '����û�����ķ�������ݿ�ƥ�� �򽫲������浽���ݿ���
                If rsExpression("����").value = txtGrossDescribe.Text Then
                    'ִ�б������
                    zlDatabase.SetPara "�޼�����ģ��", txtGrossDescribe.Text, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
                Else
                    GrossNum = GrossNum + 1
                End If
                
                If GrossNum = rsExpression.RecordCount Then
                    MsgBox "�޼�����ģ���Ӧ�ķ��࣬���ݿ��в����ڣ�"
                End If
            Else
                zlDatabase.SetPara "�޼�����ģ��", txtGrossDescribe.Text, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
            End If
            
             If txtNormalReport.Text <> "" Then
                '����û�����ķ�������ݿ�ƥ�� �򽫲������浽���ݿ���
                If rsExpression("����").value = txtNormalReport.Text Then
                    'ִ�б������
                    zlDatabase.SetPara "���汨��ģ��", txtNormalReport.Text, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
                Else
                    NormalNum = NormalNum + 1
                End If
                
                If NormalNum = rsExpression.RecordCount Then
                    MsgBox "���汨��ģ���Ӧ�ķ��࣬���ݿ��в����ڣ�"
                End If
            Else
                zlDatabase.SetPara "���汨��ģ��", txtNormalReport.Text, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
            End If
            
            If txtImmuneReport.Text <> "" Then
                '����û�����ķ�������ݿ�ƥ�� �򽫲������浽���ݿ���
                If rsExpression("����").value = txtImmuneReport.Text Then
                    'ִ�б������
                    zlDatabase.SetPara "���߱���ģ��", txtImmuneReport.Text, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
                Else
                    ImmuneNum = ImmuneNum + 1
                End If
                
                If ImmuneNum = rsExpression.RecordCount Then
                    MsgBox "���߱���ģ���Ӧ�ķ��࣬���ݿ��в����ڣ�"
                End If
            Else
                zlDatabase.SetPara "���߱���ģ��", txtImmuneReport.Text, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
            End If
            
            If txtSpecialStainReport.Text <> "" Then
                '����û�����ķ�������ݿ�ƥ�� �򽫲������浽���ݿ���
                If rsExpression("����").value = txtSpecialStainReport.Text Then
                    'ִ�б������
                    zlDatabase.SetPara "��Ⱦ����ģ��", txtSpecialStainReport.Text, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
                Else
                    SpecialNum = SpecialNum + 1
                End If
                
                If SpecialNum = rsExpression.RecordCount Then
                    MsgBox "��Ⱦ����ģ���Ӧ�ķ��࣬���ݿ��в����ڣ�"
                End If
             Else
                zlDatabase.SetPara "��Ⱦ����ģ��", txtSpecialStainReport.Text, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
            End If
            
            If txtMoleculeReport.Text <> "" Then
                '����û�����ķ�������ݿ�ƥ�� �򽫲������浽���ݿ���
                If rsExpression("����").value = txtMoleculeReport.Text Then
                    'ִ�б������
                    zlDatabase.SetPara "���ӱ���ģ��", txtMoleculeReport.Text, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
                Else
                    MoleculeNum = MoleculeNum + 1
                End If
                
                If MoleculeNum = rsExpression.RecordCount Then
                    MsgBox "���ӱ���ģ���Ӧ�ķ��࣬���ݿ��в����ڣ�"
                End If
            Else
                zlDatabase.SetPara "���ӱ���ģ��", txtMoleculeReport.Text, glngSys, mlngModul, InStr(mstrPrivs, ";��������;") > 0
            End If
            
            If Not rsExpression.EOF Then
                rsExpression.MoveNext
            End If

        Next
    End If
    
    mblnOK = True
    Unload Me
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Call CmdHelp_Click
    ElseIf KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub InitFaceScheme()
    Dim Item As TabControlItem
    
    If mlngModul = 1290 And InStr(mstrPrivs, "��ά�ؽ�����") <> 0 Then
        cmd3DSetup.Visible = True
    Else
        cmd3DSetup.Visible = False
    End If
    
    
    '����ǲ���ϵͳ���򲻽���ִ�м������
    If mlngModul = G_LNG_PATHOLSYS_NUM Then
        chkReagent.Visible = False
        chkInputOutInfo.Visible = False     '����ϵͳ�����еǼ���Ժ�ͼ쵥λ���ͼ�ҽ��
    Else
        frmPatholParameter.Visible = False
        Form_Resize
    End If
End Sub


Private Sub Form_Load()
    InitFaceScheme
    mblnOK = False
    Dim intTemp As Integer
    Dim strTemp As String
    Dim i As Integer
    Dim blnChkVisible As Boolean
    
    '����ı�ɼ������С���Ʋ�����Ĭ��Ϊ��
    blnChkVisible = True
    
    labMoneyExemodel.Visible = IIf(mlngModul = G_LNG_VIDEOSTATION_MODULE, True, False)
    cbxMoneyExeModle.Visible = IIf(mlngModul = G_LNG_VIDEOSTATION_MODULE, True, False)
    
        
    '���ݲ�ͬ�Ĺ���վ���ز�ͬ�� ��Ҫ����ҳ�� ����
    Select Case mlngModul

        Case 1290 'ҽ������վ
            cbxMainPage.Clear
            cbxMainPage.AddItem ("")
            cbxMainPage.AddItem ("Ӱ��")
            cbxMainPage.AddItem ("����")
            cbxMainPage.AddItem ("����")
            cbxMainPage.AddItem ("ҽ��")
            cbxMainPage.AddItem ("����")
            blnChkVisible = False
            
        Case 1291 '�ɼ�����վ
            cbxMainPage.Clear
            cbxMainPage.AddItem ("")
            cbxMainPage.AddItem ("�ɼ�")
            cbxMainPage.AddItem ("����")
            cbxMainPage.AddItem ("����")
            cbxMainPage.AddItem ("ҽ��")
            cbxMainPage.AddItem ("����")
            
            blnChkVisible = True
            
            cbxMoneyExeModle.ListIndex = Val(zlDatabase.GetPara("�ɼ�����ִ��ģʽ", glngSys, mlngModul, 0, Array(cbxMoneyExeModle), InStr(mstrPrivs, ";��������;") > 0))
        Case 1294 '������վ
            cbxMainPage.Clear
            cbxMainPage.AddItem ("")
            cbxMainPage.AddItem ("�ɼ�")
            cbxMainPage.AddItem ("����")
            cbxMainPage.AddItem ("ȡ��")
            cbxMainPage.AddItem ("��Ƭ")
            cbxMainPage.AddItem ("����")
            cbxMainPage.AddItem ("���")
            cbxMainPage.AddItem ("����")
            cbxMainPage.AddItem ("����")
            cbxMainPage.AddItem ("ҽ��")
            cbxMainPage.AddItem ("����")
            blnChkVisible = True
            
    End Select
    
    chkStartVideoCapture.Visible = blnChkVisible
    
    '���ó�ʼ������ִ�м����ƿؼ�
    InitExecuteRooms
    
    CmdDevSet.Enabled = InStr(mstrPrivs, ";��������;") > 0
    cmdOK.Enabled = InStr(mstrPrivs, ";��������;") > 0
    chkView.value = Val(zlDatabase.GetPara("����ʱ��Ƭ", glngSys, mlngModul, 0, Array(chkView), InStr(mstrPrivs, ";��������;") > 0))
    
    chkBatchInput.value = Val(zlDatabase.GetPara("�����Ǽ�����", glngSys, mlngModul, 0, Array(chkBatchInput), InStr(mstrPrivs, ";��������;") > 0))
    chkPatTrack.value = Val(zlDatabase.GetPara("���˸���", glngSys, mlngModul, 0, Array(chkPatTrack), InStr(mstrPrivs, ";��������;") > 0))
    ChkOpenReport.value = Val(zlDatabase.GetPara("��ʼ����Զ��򿪱���", glngSys, mlngModul, 0, Array(ChkOpenReport), InStr(mstrPrivs, ";��������;") > 0))
    chkReagent.value = Val(zlDatabase.GetPara("����ʾ��Ӱ��", glngSys, mlngModul, 0, Array(chkReagent), InStr(mstrPrivs, ";��������;") > 0))
    chkAddons.value = Val(zlDatabase.GetPara("����ʾ��������", glngSys, mlngModul, 0, Array(chkAddons), InStr(mstrPrivs, ";��������;") > 0))
    chkInputOutInfo.value = Val(zlDatabase.GetPara("¼����Ժ��Ϣ", glngSys, mlngModul, 0, Array(chkInputOutInfo), InStr(mstrPrivs, ";��������;") > 0))
    intTemp = Val(zlDatabase.GetPara("�Ǽ�ģʽ", glngSys, mlngModul, 0, Array(optCheckInMode(1)), InStr(mstrPrivs, ";��������;") > 0))
    intTemp = Val(zlDatabase.GetPara("�Ǽ�ģʽ", glngSys, mlngModul, 0, Array(optCheckInMode(2)), InStr(mstrPrivs, ";��������;") > 0))
    If intTemp = 1 Then
        optCheckInMode(1).value = True
    Else
        optCheckInMode(2).value = True
    End If
    ''
    If Val(GetDeptPara(mlng����ID, "�����Ŷӽк�", 0)) = 1 Then
        If Val(GetDeptPara(mlng����ID, "�Ŷӽкŷ�ʽ", 0)) = 1 Then
        '�Ŷӽкŷ�ʽ��1 ���ʾΪ �������Ŷ�  �����ñ���ִ�м����ƿؼ� ��֮����
            cboExecuteRooms.Enabled = True
        Else
            cboExecuteRooms.Enabled = False
        End If
    Else
        cboExecuteRooms.Enabled = False
    End If

    strTemp = zlDatabase.GetPara("����ִ�м�����", glngSys, mlngModul, "", Array(chkCancelCheck), InStr(mstrPrivs, ";��������;") > 0)

    For i = 0 To cboExecuteRooms.ListCount - 1
        If cboExecuteRooms.list(i) = strTemp Then
            cboExecuteRooms.ListIndex = i
            Exit For
        End If
    Next
    
    chkCancelCheck.value = Val(zlDatabase.GetPara("����ʾ��ȡ���ĵǼ�", glngSys, mlngModul, 0, Array(chkCancelCheck), InStr(mstrPrivs, ";��������;") > 0))
    chkAutoPrint.value = Val(zlDatabase.GetPara("�������Զ���ӡ���뵥", glngSys, mlngModul, 0, Array(chkAutoPrint), InStr(mstrPrivs, ";��������;") > 0))
    chkExitAfterSign.value = Val(zlDatabase.GetPara("PACS����ǩ�����˳�", glngSys, mlngModul, "1", Array(chkExitAfterSign), InStr(mstrPrivs, ";��������;") > 0))
    cbxMainPage.Text = zlDatabase.GetPara("������ҳ", glngSys, mlngModul, "", Array(cbxMainPage), InStr(mstrPrivs, ";��������;") > 0)
        
            
    If mlngModul <> 1290 Then
        chkStartVideoCapture.value = Val(zlDatabase.GetPara("����ı�ɼ������С", glngSys, mlngModul, "1", Array(chkStartVideoCapture), InStr(mstrPrivs, ";��������;") > 0))
    End If
    
    If mlngModul = G_LNG_PATHOLSYS_NUM Then
        chkDecalin.value = Val(zlDatabase.GetPara("�Ѹ���������", glngSys, mlngModul, 1, Array(chkDecalin), InStr(mstrPrivs, ";��������;") > 0))
        txtDecalinHintTime.Text = Val(zlDatabase.GetPara("���Ѽ��ʱ��", glngSys, mlngModul, "30", Array(txtDecalinHintTime), InStr(mstrPrivs, ";��������;") > 0))
        chkChangGui.value = Val(zlDatabase.GetPara("������������", glngSys, mlngModul, 1, Array(chkChangGui), InStr(mstrPrivs, ";��������;") > 0))
        chkKuaiPian.value = Val(zlDatabase.GetPara("����ʯ����������", glngSys, mlngModul, 1, Array(chkKuaiPian), InStr(mstrPrivs, ";��������;") > 0))
        chkBingDong.value = Val(zlDatabase.GetPara("������������", glngSys, mlngModul, 1, Array(chkBingDong), InStr(mstrPrivs, ";��������;") > 0))
        chkXiBao.value = Val(zlDatabase.GetPara("ϸ����������", glngSys, mlngModul, 1, Array(chkXiBao), InStr(mstrPrivs, ";��������;") > 0))
        chkHuiZhen.value = Val(zlDatabase.GetPara("������������", glngSys, mlngModul, 1, Array(chkHuiZhen), InStr(mstrPrivs, ";��������;") > 0))
        chkShiJian.value = Val(zlDatabase.GetPara("ʬ����������", glngSys, mlngModul, 1, Array(chkShiJian), InStr(mstrPrivs, ";��������;") > 0))
        '��ȡ�Ƿ�ֱ�Ӵ�ӡ������Ϣ
        chkIsDirectPrint.value = Val(zlDatabase.GetPara("�Ƿ�ֱ�Ӵ�ӡ", glngSys, mlngModul, 0, Array(chkIsDirectPrint), InStr(mstrPrivs, ";��������;") > 0))
        '��ȡģ���Ӧ�ʾ�������
        txtGrossDescribe.Text = zlDatabase.GetPara("�޼�����ģ��", glngSys, mlngModul, "", Array(txtGrossDescribe), InStr(mstrPrivs, ";��������;") > 0)
        txtNormalReport.Text = zlDatabase.GetPara("���汨��ģ��", glngSys, mlngModul, "", Array(txtNormalReport), InStr(mstrPrivs, ";��������;") > 0)
        txtImmuneReport.Text = zlDatabase.GetPara("���߱���ģ��", glngSys, mlngModul, "", Array(txtImmuneReport), InStr(mstrPrivs, ";��������;") > 0)
        txtSpecialStainReport.Text = zlDatabase.GetPara("��Ⱦ����ģ��", glngSys, mlngModul, "", Array(txtSpecialStainReport), InStr(mstrPrivs, ";��������;") > 0)
        txtMoleculeReport.Text = zlDatabase.GetPara("���ӱ���ģ��", glngSys, mlngModul, "", Array(txtMoleculeReport), InStr(mstrPrivs, ";��������;") > 0)

    End If
End Sub

Private Sub Form_Resize()
    PicAction.Left = (Me.ScaleWidth - PicAction.Width) / 2
    If mlngModul <> G_LNG_PATHOLSYS_NUM Then
        With PicAction
            .Left = 120
            .Top = 120
            .Width = Me.ScaleWidth - 240
            .Height = 4180
        End With
        
        Me.Height = PicAction.Height + 720
    Else
'        optCheckInMode(1).Top = 360
'        optCheckInMode(2).Top = 360
'        chkAddons.Top = 840
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlng����ID = 0
End Sub

Private Sub TxtLike_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub TxtShowPhotoNumber_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
Private Sub TxtĬ������_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub optCheckInMode_Click(Index As Integer)
    If optCheckInMode(1).value = True Then
        chkReagent.Enabled = False
        chkAddons.Enabled = False
    Else
        chkReagent.Enabled = True
        chkAddons.Enabled = True
    End If
End Sub

Private Sub PicAction_Resize()
    On Error Resume Next
    
    If mlngModul <> G_LNG_PATHOLSYS_NUM Then
        chkAutoPrint.Left = 240
        chkAutoPrint.Top = 120
        
        chkBatchInput.Left = chkBatchInput.Left
        chkBatchInput.Top = chkAutoPrint.Top
        
        chkCancelCheck.Left = 240
        chkCancelCheck.Top = chkAutoPrint.Top + chkAutoPrint.Height + 105
        
        chkView.Left = chkView.Left
        chkView.Top = chkCancelCheck.Top
        
        chkPatTrack.Left = 240
        chkPatTrack.Top = chkCancelCheck.Top + chkCancelCheck.Height + 105
        
        ChkOpenReport.Left = ChkOpenReport.Left
        ChkOpenReport.Top = chkPatTrack.Top
        
        chkExitAfterSign.Left = 240
        chkExitAfterSign.Top = chkPatTrack.Top + chkPatTrack.Height + 105
        
        chkStartVideoCapture.Left = ChkOpenReport.Left
        chkStartVideoCapture.Top = chkExitAfterSign.Top
        
        Load Frame6(1)
        With Frame6(1)
            .Left = Frame6(0).Left
            .Top = chkStartVideoCapture.Top + chkStartVideoCapture.Height + 150
            .Width = Frame6(0).Width
            .Height = 25
            
            .Caption = ""
            .Visible = True
        End With
        
        Frame2.Left = Frame2.Left
        Frame2.Top = Frame6(1).Top + Frame6(1).Height + 200
        Frame2.Height = 1185
        
        labMoneyExemodel.Left = Frame2.Left
        labMoneyExemodel.Top = Frame2.Top + Frame2.Height + 175
        
        cbxMoneyExeModle.Left = labMoneyExemodel.Left + labMoneyExemodel.Width
        cbxMoneyExeModle.Top = labMoneyExemodel.Top - 40
        
        
        lblExecuteRoomName.Left = Label2.Left
        lblExecuteRoomName.Top = Frame6(1).Top + Frame6(1).Height + 150
        
        cboExecuteRooms.Left = lblExecuteRoomName.Left
        cboExecuteRooms.Top = lblExecuteRoomName.Top + lblExecuteRoomName.Height + 120
        
        Label2.Left = Label2.Left
        Label2.Top = cboExecuteRooms.Top + 480
        
        cbxMainPage.Left = cbxMainPage.Left
        cbxMainPage.Top = Label2.Top + Label2.Height + 100
        
        Frame6(0).Left = Frame6(0).Left
        Frame6(0).Top = cbxMoneyExeModle.Top + cbxMoneyExeModle.Height + 150
        
        cmdHelp.Left = cmdHelp.Left
        cmdHelp.Top = Frame6(0).Top + Frame6(0).Height + 150
        
        CmdDevSet.Left = CmdDevSet.Left
        CmdDevSet.Top = Frame6(0).Top + Frame6(0).Height + 150
        
        cmd3DSetup.Left = cmd3DSetup.Left
        cmd3DSetup.Top = Frame6(0).Top + Frame6(0).Height + 150
        
        cmdOK.Left = cmdOK.Left
        cmdOK.Top = Frame6(0).Top + Frame6(0).Height + 150
        
        cmdCancel.Left = cmdCancel.Left
        cmdCancel.Top = Frame6(0).Top + Frame6(0).Height + 150
    End If
End Sub
