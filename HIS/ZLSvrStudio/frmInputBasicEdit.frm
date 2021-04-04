VERSION 5.00
Begin VB.Form frmInputBasicEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "基本字词编辑"
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
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2100
      TabIndex        =   4
      Top             =   1755
      Width           =   1100
   End
   Begin VB.CommandButton cmdClear 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
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
         Caption         =   "字词编码(&M)"
         Height          =   180
         Index           =   1
         Left            =   390
         TabIndex        =   2
         Top             =   840
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "基本字词(&W)"
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
Private mstrOld字词 As String
Private mstrOld输入码 As String
Private mblnNew As Boolean
Private mbytBasicWord As Byte
Private mstrArray(1 To 26) As String

Private Function CheckChineseWord(ByVal strWord As String) As Boolean
    '------------------------------------------------------------------------------------------------
    '功能:是否为汉字字符
    '------------------------------------------------------------------------------------------------
    Dim strNoString As String
    
    strNoString = "ⅠⅡⅢⅣⅤⅥⅧⅧⅨ°"
    
    If strWord = " " Then Exit Function
    '将全角转换为半角
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
    '功能:获取汉字的拼音码或五笔码
    '------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim lngLoop As Long
    
    On Error GoTo errHand
    
    strSQL = "SELECT 输入码 FROM zlWordBasic WHERE 字词='" & strWord & "' AND 输入法=" & bytMode
    rs.Open strSQL, gcnOracle
    If rs.BOF = False Then
        GetWordCode = rs("输入码").Value
        Exit Function
    End If
    
    For lngLoop = 1 To Len(strWord)
        
        strSQL = "SELECT 输入码 FROM zlWordBasic WHERE 字词='" & Mid(strWord, lngLoop, 1) & "' AND 输入法=" & bytMode
        rs.Close
        rs.Open strSQL, gcnOracle
        If rs.BOF = False Then
            GetWordCode = GetWordCode & Left(rs("输入码").Value, 1)
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
    
    mstrOld字词 = ""
    mstrOld输入码 = ""

    If mbytMode = 1 Then
        
        If mbytBasicWord = 1 Then
            Me.Caption = "拼音码基本字编辑"
            lbl(0).Caption = "基本字(&W)"
            txt(0).MaxLength = 2
        Else
            Me.Caption = "拼音码基本词编辑"
            lbl(0).Caption = "基本词(&W)"
            txt(0).MaxLength = 100
        End If
        
        lbl(1).Caption = "拼音码(&M)"
    Else
        If mbytBasicWord = 1 Then
            Me.Caption = "五笔码基本字编辑"
            lbl(0).Caption = "基本字(&W)"
            txt(0).MaxLength = 2
        Else
            Me.Caption = "五笔码基本词编辑"
            lbl(0).Caption = "基本词(&W)"
            txt(0).MaxLength = 100
        End If
        lbl(1).Caption = "五笔码(&M)"
    End If
    
    If blnNew = False Then
        mstrOld字词 = strWord
        mstrOld输入码 = strCode
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
        strError = "必须输入基本字词！"
        Call LocationTxt(txt(0))
        GoTo errHand
    End If
    
    If Trim(txt(1).Text) = "" Then
        strError = "必须输入基本字词的编码！"
        Call LocationTxt(txt(1))
        GoTo errHand
    End If
    
    For mlngLoop = 1 To Len(txt(1).Text)
        If InStr("ZXCVBNMASDFGHJKLQWERTYUIOP", UCase(Mid(txt(1).Text, mlngLoop, 1))) = 0 Then
            strError = "基本字词的编码必须在字母a-z之间！"
            Call LocationTxt(txt(1))
            GoTo errHand
        End If
    Next
    
    For mlngLoop = 1 To Len(txt(0).Text)
        
        If CheckChineseWord(Mid(txt(0).Text, mlngLoop, 1)) = False Then
            strError = "输入的字或词必须为汉字字符！"
            Call LocationTxt(txt(0))
            GoTo errHand
        End If
        
    Next
    
    If mbytBasicWord <> 1 Then
                    
        If Len(txt(0).Text) < 2 Then
            strError = "基本词必须大于一个汉字！"
            Call LocationTxt(txt(0))
            GoTo errHand
        End If
        
        '基本词,第一个编码符必须是第一个汉字的的首编码符
        gstrSQL = "SELECT 1 FROM zlWordBasic WHERE 字词='" & Left(txt(0).Text, 1) & "' and 输入法=" & mbytMode & " and 输入码 like '" & Left(txt(1).Text, 1) & "%'"
        rs.Open gstrSQL, gcnOracle
        If rs.BOF Then
            strError = "编码的首字符必须与""" & Left(txt(0).Text, 1) & """的首字符相同！"
            Call LocationTxt(txt(1))
            GoTo errHand
        End If
        
        '检查每一个字是否在基本字之列
        For mlngLoop = 2 To Len(txt(0).Text)
            gstrSQL = "SELECT 1 FROM zlWordBasic WHERE 字词='" & Mid(txt(0).Text, mlngLoop, 1) & "' and 是否字=1 and 输入法=" & mbytMode
            If rs.State = adStateOpen Then rs.Close
            rs.Open gstrSQL, gcnOracle
            If rs.BOF Then
                strError = "第" & mlngLoop & "个字不在基本字之中！"
                Call LocationTxt(txt(0))
                GoTo errHand
            End If
        Next
    Else
        
        If mbytMode = 1 Then
            
            '检查基本字的拼音是否有效
            For mlngLoop = 1 To 26
                If InStr(";" & mstrArray(mlngLoop) & ";", ";" & txt(1).Text & ";") > 0 Then
                    Exit For
                End If
            Next
            
            If mlngLoop > 26 Then
                strError = "基本字的拼音码不正确，请重新设置！"
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
    
    If txt(0).Text <> mstrOld字词 Or txt(1).Text <> mstrOld输入码 Then
    
        gstrSQL = "SELECT 1 FROM zlWordBasic WHERE 字词='" & txt(0).Text & "' AND 输入码='" & txt(1).Text & "' and 输入法=" & mbytMode
        rs.Open gstrSQL, gcnOracle
        If rs.BOF = False Then
            MsgBox "你输入的基本字词及其编码已经存在！", vbInformation, gstrSysName
            txt(1).SetFocus
            Exit Sub
        End If
        
        On Error GoTo errHand
        gcnOracle.BeginTrans
        If mblnNew Then
            gcnOracle.Execute "ZL_zlWordBasic_INSERT('" & txt(0).Text & "','" & LCase(txt(1).Text) & "'," & mbytMode & "," & mbytBasicWord & ")", , adCmdStoredProc
        Else
            gcnOracle.Execute "ZL_zlWordBasic_UPDATE('" & txt(0).Text & "','" & LCase(txt(1).Text) & "'," & mbytMode & "," & mbytBasicWord & ",'" & mstrOld字词 & "','" & mstrOld输入码 & "')", , adCmdStoredProc
            'gcnOracle.Execute "UPDATE zlWordBasic SET 字词='" & txt(0).Text & "',输入码='" & LCase(txt(1).Text) & "' WHERE 字词='" & mstrOld字词 & "' AND 输入码='" & mstrOld输入码 & "' AND 输入法=" & mbytMode
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
    MsgBox "编辑基本字词失败！" & vbNewLine & Err.Description, vbInformation, gstrSysName
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdOK.Tag <> "" Then
        Cancel = (MsgBox("修改后的基本字词必须保存后才生效，是否放弃保存？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo)
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
