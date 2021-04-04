VERSION 5.00
Begin VB.Form frmInputBasicEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����ִʱ༭"
   ClientHeight    =   2310
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4530
   Icon            =   "frmInputBasicEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2100
      TabIndex        =   4
      Top             =   1755
      Width           =   1100
   End
   Begin VB.CommandButton cmdClear 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3255
      TabIndex        =   5
      Top             =   1755
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   1485
      Left            =   105
      TabIndex        =   6
      Top             =   120
      Width           =   4245
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1455
         MaxLength       =   12
         TabIndex        =   3
         Top             =   795
         Width           =   2325
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1455
         MaxLength       =   100
         TabIndex        =   1
         Top             =   345
         Width           =   2325
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�ִʱ���(&M)"
         Height          =   180
         Index           =   1
         Left            =   390
         TabIndex        =   2
         Top             =   840
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�����ִ�(&W)"
         Height          =   180
         Index           =   0
         Left            =   405
         TabIndex        =   0
         Top             =   405
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmInputBasicEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mbytMode As Byte
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngLoop As Long
Private mstrOld�ִ� As String
Private mstrOld������ As String
Private mblnNew As Boolean
Private mbytBasicWord As Byte
Private mstrArray(1 To 26) As String

Private Function CheckChineseWord(ByVal strWord As String) As Boolean
    '------------------------------------------------------------------------------------------------
    '����:�Ƿ�Ϊ�����ַ�
    '------------------------------------------------------------------------------------------------
    Dim strNoString As String
    
    strNoString = "�����������������"
    
    If strWord = " " Then Exit Function
    '��ȫ��ת��Ϊ���
    strWord = StrConv(Trim(strWord), vbNarrow)
    
    If Asc(strWord) >= 0 Then Exit Function
    
    If InStr(strNoString, strWord) > 0 Then Exit Function
    
    CheckChineseWord = True
    
End Function

Private Sub LocationTxt(objTxt As Object)
    objTxt.SetFocus
    objTxt.SelStart = 0
    objTxt.SelLength = Len(objTxt)
End Sub

Private Function GetWordCode(ByVal strWord As String, Optional ByVal bytMode As Byte = 1) As String
    '------------------------------------------------------------------------------------------------
    '����:��ȡ���ֵ�ƴ����������
    '------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim lngLoop As Long
    
    On Error GoTo errHand
    
    strSQL = "SELECT ������ FROM zlWordBasic WHERE �ִ�='" & strWord & "' AND ���뷨=" & bytMode
    rs.Open strSQL, gcnOracle
    If rs.BOF = False Then
        GetWordCode = rs("������").Value
        Exit Function
    End If
    
    For lngLoop = 1 To Len(strWord)
        
        strSQL = "SELECT ������ FROM zlWordBasic WHERE �ִ�='" & Mid(strWord, lngLoop, 1) & "' AND ���뷨=" & bytMode
        rs.Close
        rs.Open strSQL, gcnOracle
        If rs.BOF = False Then
            GetWordCode = GetWordCode & Left(rs("������").Value, 1)
        End If
    Next
    
    Exit Function
    
errHand:
    GetWordCode = ""
End Function

Public Function ShowEdit(ByVal frmMain As Object, ByVal blnNew As Boolean, _
                        ByRef strWord As String, _
                        ByVal strCode As String, _
                        ByVal bytMode As Byte, _
                        Optional bytBasicWord As Byte = 1) As Boolean

    mblnOK = False
    
    mbytMode = bytMode
    Set mfrmMain = frmMain
    mblnNew = blnNew
    mbytBasicWord = bytBasicWord
    
    mstrArray(1) = "a;ai;an;ang;ao"
    mstrArray(2) = "ba;bai;ban;bang;bao;bei;ben;beng;bi;bian;biao;bie;bin;bing;bo;bu"
    mstrArray(3) = "ca;cai;can;cang;cao;ce;cen;ceng;cha;chai;chan;chang;chao;che;chen;cheng;chi;chong;chou;chu;chua;chuai;chuan;chuang;chui;chun;chuo;ci;cong;cou;cu;cuan;cui;cun;cuo"
    mstrArray(4) = "da;dai;dang;dao;de;dei;den;deng;di;dia;dian;diao;die;ding;diu;dong;dou;du;duan;dui;dun;duo"
    mstrArray(5) = "e;ei;en;eng;er"
    mstrArray(6) = "fa;fan;fang;fei;fen;feng;fo;fou;fu"
    mstrArray(7) = "ga;gai;gan;gang;gao;ge;gei;gen;geng;gong;gou;gu;gua;guan;guang;gui;gun;guo"
    mstrArray(8) = "ha;hai;han;hang;hao;he;hei;hen;heng;hng;hong;hou;hu;hua;huai;huan;huang;hui;hun;huo"
    mstrArray(10) = "ji;jia;jian;jiang;jiao;jie;jin;jing;jiong;jiu;ju;juan;jue;jun"
    mstrArray(11) = "ka;kai;kan;kang;kao;ke;kei;ken;keng;kong;kou;ku;kua;kuai;kuan;kuang;kui;kun;kuo"
    mstrArray(12) = "la;lai;lan;lang;lao;le;lei;leng;li;lia;lian;liang;liao;lie;lin;ling;liu;lo;long;lou;lu;luan;lun;luo;lv;lve"
    mstrArray(13) = "m;ma;mai;man;mang;mao;me;mei;men;meng;mi;mian;miao;mie;min;ming;miu;mo;mou;mu"
    mstrArray(14) = "ngn;na;nai;nan;nang;nao;ne;nei;nen;neng;ng;ni;nian;niang;niao;nie;nin;ning;niu;nong;nou;nu;nuan;nun;nuo;nv;nve"
    mstrArray(15) = "o;ou"
    mstrArray(16) = "pa;pai;pan;pang;pao;pei;pen;peng;pi;pian;piao;pie;pin;ping;po;pou;pu"
    mstrArray(17) = "qi;qia;qian;qiang;qianwa;qiao;qie;qin;qing;qiong;qiu;qu;quan;que;qun"
    mstrArray(18) = "ran;rang;rao;re;ren;reng;ri;rong;rou;ru;ruan;rui;run;ruo"
    mstrArray(19) = "sa;sai;san;sang;sao;se;sen;seng;sha;shai;shan;shang;shao;she;shei;shen;sheng;shi;shou;shu;shua;shuai;shuan;shuang;shui;shun;shuo;si;song;sou;su;suan;sui;sun;suo"
    mstrArray(20) = "ta;tai;tan;tang;tao;te;teng;ti;tian;tiao;tie;ting;tong;tou;tu;tuan;tui;tun;tuo"
    mstrArray(23) = "wa;wai;wan;wang;wei;wen;weng;wo;wu"
    mstrArray(24) = "xi;xia;xian;xiang;xiao;xie;xin;xing;xiong;xu;xuan;xue;xun"
    mstrArray(25) = "ya;yan;yang;yao;ye;yi;yin;ying;yingli;yo;yong;you;yu;yuan;yue;yun"
    mstrArray(26) = "za;zai;zan;zang;zao;ze;zei;zen;zeng;zha;zhai;zhan;zhang;zhao;zhe;zhei;zhen;zheng;zhi;zhong;zhou;zhu;zhua;zhuai;zhuan;zhuang;zhui;zhun;zhuo;zi;zong;zou;zu;zuan;zui;zun;zuo"
    
    mstrOld�ִ� = ""
    mstrOld������ = ""

    If mbytMode = 1 Then
        
        If mbytBasicWord = 1 Then
            Me.Caption = "ƴ��������ֱ༭"
            lbl(0).Caption = "������(&W)"
            txt(0).MaxLength = 2
        Else
            Me.Caption = "ƴ��������ʱ༭"
            lbl(0).Caption = "������(&W)"
            txt(0).MaxLength = 100
        End If
        
        lbl(1).Caption = "ƴ����(&M)"
    Else
        If mbytBasicWord = 1 Then
            Me.Caption = "���������ֱ༭"
            lbl(0).Caption = "������(&W)"
            txt(0).MaxLength = 2
        Else
            Me.Caption = "���������ʱ༭"
            lbl(0).Caption = "������(&W)"
            txt(0).MaxLength = 100
        End If
        lbl(1).Caption = "�����(&M)"
    End If
    
    If blnNew = False Then
        mstrOld�ִ� = strWord
        mstrOld������ = strCode
        txt(0).Text = strWord
        txt(1).Text = strCode
    Else
        txt(0).Text = strWord
        txt(1).Text = strCode
    End If
    
    cmdOK.Tag = ""
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Sub cmdClear_Click()
    Unload Me
End Sub

Private Function CheckDataValid() As Boolean
    Dim rs As New ADODB.Recordset
    Dim strError As String
    
    If Trim(txt(0).Text) = "" Then
        strError = "������������ִʣ�"
        Call LocationTxt(txt(0))
        GoTo errHand
    End If
    
    If Trim(txt(1).Text) = "" Then
        strError = "������������ִʵı��룡"
        Call LocationTxt(txt(1))
        GoTo errHand
    End If
    
    For mlngLoop = 1 To Len(txt(1).Text)
        If InStr("ZXCVBNMASDFGHJKLQWERTYUIOP", UCase(Mid(txt(1).Text, mlngLoop, 1))) = 0 Then
            strError = "�����ִʵı����������ĸa-z֮�䣡"
            Call LocationTxt(txt(1))
            GoTo errHand
        End If
    Next
    
    For mlngLoop = 1 To Len(txt(0).Text)
        
        If CheckChineseWord(Mid(txt(0).Text, mlngLoop, 1)) = False Then
            strError = "������ֻ�ʱ���Ϊ�����ַ���"
            Call LocationTxt(txt(0))
            GoTo errHand
        End If
        
    Next
    
    If mbytBasicWord <> 1 Then
                    
        If Len(txt(0).Text) < 2 Then
            strError = "�����ʱ������һ�����֣�"
            Call LocationTxt(txt(0))
            GoTo errHand
        End If
        
        '������,��һ������������ǵ�һ�����ֵĵ��ױ����
        gstrSQL = "SELECT 1 FROM zlWordBasic WHERE �ִ�='" & Left(txt(0).Text, 1) & "' and ���뷨=" & mbytMode & " and ������ like '" & Left(txt(1).Text, 1) & "%'"
        rs.Open gstrSQL, gcnOracle
        If rs.BOF Then
            strError = "��������ַ�������""" & Left(txt(0).Text, 1) & """�����ַ���ͬ��"
            Call LocationTxt(txt(1))
            GoTo errHand
        End If
        
        '���ÿһ�����Ƿ��ڻ�����֮��
        For mlngLoop = 2 To Len(txt(0).Text)
            gstrSQL = "SELECT 1 FROM zlWordBasic WHERE �ִ�='" & Mid(txt(0).Text, mlngLoop, 1) & "' and �Ƿ���=1 and ���뷨=" & mbytMode
            If rs.State = adStateOpen Then rs.Close
            rs.Open gstrSQL, gcnOracle
            If rs.BOF Then
                strError = "��" & mlngLoop & "���ֲ��ڻ�����֮�У�"
                Call LocationTxt(txt(0))
                GoTo errHand
            End If
        Next
    Else
        
        If mbytMode = 1 Then
            
            '�������ֵ�ƴ���Ƿ���Ч
            For mlngLoop = 1 To 26
                If InStr(";" & mstrArray(mlngLoop) & ";", ";" & txt(1).Text & ";") > 0 Then
                    Exit For
                End If
            Next
            
            If mlngLoop > 26 Then
                strError = "�����ֵ�ƴ���벻��ȷ�����������ã�"
                Call LocationTxt(txt(1))
                GoTo errHand
            End If
            
        End If
    End If
    
    CheckDataValid = True
    
    Exit Function
    
errHand:
    MsgBox strError, vbInformation, gstrSysName
End Function

Private Sub cmdOK_Click()
    Dim rs As New ADODB.Recordset
    
    If cmdOK.Tag = "" Then Exit Sub
    
    If CheckDataValid = False Then Exit Sub
    
    If txt(0).Text <> mstrOld�ִ� Or txt(1).Text <> mstrOld������ Then
    
        gstrSQL = "SELECT 1 FROM zlWordBasic WHERE �ִ�='" & txt(0).Text & "' AND ������='" & txt(1).Text & "' and ���뷨=" & mbytMode
        rs.Open gstrSQL, gcnOracle
        If rs.BOF = False Then
            MsgBox "������Ļ����ִʼ�������Ѿ����ڣ�", vbInformation, gstrSysName
            txt(1).SetFocus
            Exit Sub
        End If
        
        On Error GoTo errHand
        gcnOracle.BeginTrans
        If mblnNew Then
            gcnOracle.Execute "ZL_zlWordBasic_INSERT('" & txt(0).Text & "','" & LCase(txt(1).Text) & "'," & mbytMode & "," & mbytBasicWord & ")", , adCmdStoredProc
        Else
            gcnOracle.Execute "ZL_zlWordBasic_UPDATE('" & txt(0).Text & "','" & LCase(txt(1).Text) & "'," & mbytMode & "," & mbytBasicWord & ",'" & mstrOld�ִ� & "','" & mstrOld������ & "')", , adCmdStoredProc
            'gcnOracle.Execute "UPDATE zlWordBasic SET �ִ�='" & txt(0).Text & "',������='" & LCase(txt(1).Text) & "' WHERE �ִ�='" & mstrOld�ִ� & "' AND ������='" & mstrOld������ & "' AND ���뷨=" & mbytMode
        End If
        gcnOracle.CommitTrans
        
        On Error Resume Next
        
        mblnOK = True
        
        Call mfrmMain.RefreshData(txt(0).Text, LCase(txt(1).Text))
    End If
    
    If mblnNew Then
        txt(0).Text = ""
        txt(1).Text = ""
        cmdOK.Tag = ""
        txt(0).SetFocus
    Else
        cmdOK.Tag = ""
        Unload Me
    End If
    
    Exit Sub
    
errHand:
    gcnOracle.RollbackTrans
    MsgBox "�༭�����ִ�ʧ�ܣ�" & vbNewLine & Err.Description, vbInformation, gstrSysName
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdOK.Tag <> "" Then
        Cancel = (MsgBox("�޸ĺ�Ļ����ִʱ��뱣������Ч���Ƿ�������棿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo)
    End If
End Sub

Private Sub txt_Change(Index As Integer)
    cmdOK.Tag = "Changed"
End Sub

Private Sub txt_GotFocus(Index As Integer)
    txt(Index).SelStart = 0
    txt(Index).SelLength = Len(txt(Index).Text)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    Else
        If Index = 1 Then KeyAscii = Asc(LCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    If Index = 0 And Trim(txt(0).Text) = "" Then
        txt(1).Text = GetWordCode(txt(0).Text, mbytMode)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub
