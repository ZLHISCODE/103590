VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPubSel 
   AutoRedraw      =   -1  'True
   Caption         =   "ѡ����"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6840
   Icon            =   "frmPubSel.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      ScaleHeight     =   450
      ScaleWidth      =   6840
      TabIndex        =   9
      Top             =   0
      Width           =   6840
      Begin VB.CommandButton cmdFind 
         Caption         =   "����"
         Height          =   300
         Left            =   5880
         TabIndex        =   13
         Top             =   97
         Width           =   855
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   4320
         TabIndex        =   12
         Top             =   97
         Width           =   1455
      End
      Begin VB.CheckBox chkShowChild 
         Caption         =   "�����¼���Ŀ"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѡ��һ����Ŀ,Ȼ����ȷ��"
         Height          =   180
         Left            =   180
         TabIndex        =   10
         Top             =   157
         Width           =   2430
      End
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   6840
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3840
      Width           =   6840
      Begin VB.CommandButton cmdSelALL 
         Caption         =   "ȫѡ(&A)"
         Height          =   360
         Left            =   300
         TabIndex        =   4
         ToolTipText     =   "Ctrl+A"
         Top             =   100
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton cmdClear 
         Cancel          =   -1  'True
         Caption         =   "ȫ��(&R)"
         Height          =   360
         Left            =   1320
         TabIndex        =   5
         ToolTipText     =   "Ctrl+R"
         Top             =   100
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   5295
         TabIndex        =   3
         Top             =   105
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   4170
         TabIndex        =   2
         Top             =   105
         Width           =   1100
      End
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3240
      Left            =   2205
      TabIndex        =   1
      Top             =   555
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   5715
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView tvw_s 
      Height          =   3240
      Left            =   15
      TabIndex        =   0
      Top             =   540
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   5715
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3210
      Left            =   2145
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3210
      ScaleWidth      =   45
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   540
      Width           =   45
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   4725
      Top             =   1425
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubSel.frx":058A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1170
      Left            =   2400
      ScaleHeight     =   1110
      ScaleWidth      =   2220
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   675
      Width           =   2280
   End
End
Attribute VB_Name = "frmPubSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrKey As String
Private mblnMulti As Boolean
Private mstrFind As String
Private mlngFindIndex As Long

'��ڲ���
Private mstrTitle As String
Private mstrNote As String
Private mbytStyle As Byte
Private mstrSeek As String
Private mblnĩ�� As Boolean
Private mblnShowSub As Boolean
Private mblnShowRoot As Boolean
Private mblnMultiOne As Boolean
Private mstrColWith As String '�п����ò���
Private mstrTipCol As String   '������ʾ����
Private mbytSize As Byte '�����С

Private mstrSaveTag As String 'ע������ּ�
Private mstrSQL As String
Private mstrDetail As String
Private marrInput() As Variant
Private marrHideCols()  As Variant '�������ص��е�����
Private mblnSearch As Boolean '�Ƿ�ͨ�������кż���
Private mblnNotShowNon As Boolean '����ʾû������ķ��࣬bytStyle=2
Private mstrHeadCap As String '������չʾ
Private mblnMultiCheckReturn As Boolean '��ѡʱ��ֻ����ѡ������˫����
Private mblnHideNullCols As Boolean '�Ƿ����� Null as  ��
Private mblnNoneWin As Boolean
Private mlngX As Long, mlngY As Long, mlngTxtH As Long
Private mstrCheck As String
Private mblnHaveCheck As Boolean '�ж�˫�б�ģʽ�£������ֶ����Ƿ���Check�ֶ�
Private mstrFields As String     '��¼��¼��ԭʼ�ֶ�
'���ڲ���
Private mrsSel As ADODB.Recordset
Private mrsDetail As ADODB.Recordset 'Ҷ�Ӽ���ϸSQL
Private mlngMaxPar As Long '��չ��Ĳ�������滻�Ĳ�����
'�������
Private mblnOK As Boolean

Public Function ShowSelect(frmParent As Object, ByVal strSQL As String, ByVal strDetail As String, bytStyle As Byte, _
    ByVal strTitle As String, blnĩ�� As Boolean, _
    ByVal strSeek As String, ByVal strNote As String, _
    ByVal blnShowSub As Boolean, blnShowRoot As Boolean, _
    ByVal blnNoneWin As Boolean, ByVal X As Long, _
    ByVal Y As Long, ByVal txtH As Long, _
    ByRef Cancel As Boolean, ByVal blnMultiOne As Boolean, _
    ByVal blnSearch As Boolean, ByVal blnMulti As Boolean, _
    Optional arrInput As Variant) As ADODB.Recordset
'���ܣ��๦��ѡ����
'������
'     frmParent=��ʾ�ĸ�����
'     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
'     bytStyle=ѡ�������
'       Ϊ0ʱ:�б���:ID,��
'       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
'       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
'             ˫��������������ں�Check��β���ֶΣ�����ֶ���Ϊ�Ƿ�ѡ��ֵ�洢�ֶΡ�=1Ϊ��ѡ��0=����ѡ��
'             ˫���������������*���ƣ�*���룬*����ģ�����ʾ���ϽǵĲ�ѯ���ܣ��Թ���ѯ��Ŀ��
'                    �����б�������ƥ�䣬ƥ��ɹ���λ���÷���ĸ���Ŀ�ϣ���F3֧�ֲ�����һ����
'     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
'     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
'     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
'             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
'             bytStyle=1ʱ,�����Ǳ��������
'     strNote=ѡ������˵������
'     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
'     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
'     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
'     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
'     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
'     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
'     blnMulti=�Ƿ������ѡ
'     arrInput=��Ӧ�ĸ���SQL����ֵ,��˳����,����Ϊ��ȷ����
'     arrInput��һ����������ǡ�����ʾû������ķ��ࡱ������ʾû������ķ���
'     arrInput�У�
'               ��ʽΪ��"bytSize=?"��ʾ���������С(0-С����,1-������;С����Ϊ9����,������Ϊ12����),Ĭ��С���塣
'               ��ʽΪ��ColSet:...ʱ��ʾ�п�����,ColSet��ʽ:�п�����|����1,���1;����2,���2.....|������ʾ|������
'               ��ʽΪ��HeadCap=SQL����1,�б�չʾ����1;SQL����2,�б�չʾ����2������Ŀ�����ֹ�ָ��SQL�����б���չʾ���ƣ�һ�����ڱ��������У����ǲ��ı��е�Key
'               ��ʽΪ��MultiCheckReturn=0,1����ѡʱֻ���ع�ѡ�У����ڶ�ѡ��ȷ��Ĭ�Ϸ��ص�ǰ���������Ӹò������ƣ��ÿ������ú󣬲�֧��Ĭ���еķ��أ������Ծ�֧��˫�����Զ����ء�
'               ��ʽΪ��HideNullCols=0,1;�Ƿ�����SQl�е�null as д������
'���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
'˵����
'     1.ID���ϼ�ID����Ϊ�ַ�������
'     2.ĩ�����ֶβ�Ҫ����ֵ
'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    
    Dim i As Long, j As Integer, arrTmp As Variant
    Dim blnPara As Boolean
    
    mstrSQL = strSQL
    mstrDetail = strDetail
    mstrColWith = "": mstrTipCol = ""
    mblnNotShowNon = False
    mbytSize = 0
    marrInput = Array()
    '�Ӳ��������н�����������
    If TypeName(arrInput) <> "Error" Then
        '�ӿɱ�����зָ��ֲ���
        If UBound(arrInput) >= 0 Then
            For i = LBound(arrInput) To UBound(arrInput)
                If TypeName(arrInput(i)) = "Error" Then arrInput(i) = "" '��û���Ĳ�����ת��Ϊ�մ�����Ȼʹ�û����
                blnPara = True
                
                If arrInput(i) Like "*=*" Then
                    If UCase(arrInput(i)) Like "BYTSIZE*=*" Then
                        mbytSize = Val(Split(arrInput(i), "=")(1)): blnPara = False
                    ElseIf UCase(arrInput(i)) Like "HEADCAP=*" Then
                        mstrHeadCap = Trim(Split(arrInput(i), "=")(1)): blnPara = False
                    ElseIf UCase(arrInput(i)) Like "MULTICHECKRETURN=*" Then
                        mblnMultiCheckReturn = Val(Split(arrInput(i), "=")(1)) = 1: blnPara = False
                    ElseIf UCase(arrInput(i)) Like "HIDENULLCOLS=*" Then
                        mblnHideNullCols = Val(Split(arrInput(i), "=")(1)) = 1: blnPara = False
                    End If
                End If
                If blnPara Then
                    If UCase(arrInput(i)) Like "COLSET:*" Then  'COLSET�������һλ
                        blnPara = False
                        arrTmp = Split(arrInput(i), ":")
                        arrTmp = Split(arrTmp(1), "|")
                        For j = LBound(arrTmp) To UBound(arrTmp) Step 2
                            If arrTmp(j) = "�п�����" Then
                                mstrColWith = arrTmp(j + 1)
                            ElseIf arrTmp(j) = "������ʾ" Then
                                mstrTipCol = arrTmp(j + 1)
                            End If
                        Next
                    ElseIf bytStyle = 2 And i = 0 Then '����ʾû������ķ�����ڵ�һλ
                        If arrInput(i) = "����ʾû������ķ���" Then mblnNotShowNon = True: blnPara = False
                    End If
                End If
                If blnPara Then
                    ReDim Preserve marrInput(UBound(marrInput) + 1)
                    marrInput(UBound(marrInput)) = arrInput(i)
                End If
            Next
        End If
    End If

    marrHideCols = Array()
    If mblnHideNullCols Then
        Call GetHideCols '��ȡ��������
    End If
    mstrTitle = strTitle
    mstrNote = strNote
    mbytStyle = bytStyle
    mblnĩ�� = blnĩ��
    mstrSeek = strSeek
    mblnShowSub = blnShowSub
    mblnShowRoot = blnShowRoot
    mblnMultiOne = blnMultiOne
    mblnNoneWin = blnNoneWin
    mlngX = X: mlngY = Y: mlngTxtH = txtH
    mblnSearch = blnSearch
    mblnMulti = blnMulti
    
    If Not frmParent Is Nothing Then
        mstrSaveTag = frmParent.Name & "_" & strTitle & "_" & bytStyle & IIF(blnNoneWin, 0, 1)
    Else
        mstrSaveTag = strTitle & "_" & bytStyle & IIF(blnNoneWin, 0, 1)
    End If
    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    If mblnOK Then
        Cancel = False
        Set ShowSelect = mrsDetail
    Else
        Cancel = True
        Set ShowSelect = Nothing
    End If
End Function

Public Function ShowSelectV2(frmParent As Object, ByVal objControl As Object, ByVal strSQL As String, bytStyle As Byte, _
                                                ByVal strTitle As String, ByVal blnĩ�� As Boolean, ByVal strSeek As String, ByVal strNote As String, _
                                                ByVal blnShowSub As Boolean, ByVal blnShowRoot As Boolean, ByVal blnNoneWin As Boolean, ByRef Cancel As Boolean, _
                                                Optional ByVal blnMultiOne As Boolean, Optional ByVal blnSearch As Boolean, Optional ByVal blnMulti As Boolean, _
                                                Optional ByVal strOtherInfo As String, Optional arrInput As Variant) As ADODB.Recordset
'���ܣ��๦��ѡ����
'������
'     frmParent=��ʾ�ĸ�����
'     objControl=���ý��������
'     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
'     bytStyle=ѡ�������
'       Ϊ0ʱ:�б���:ID,��
'       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
'       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
'             ˫��������������ں�Check��β���ֶΣ�����ֶ���Ϊ�Ƿ�ѡ��ֵ�洢�ֶΡ�=1Ϊ��ѡ��0=����ѡ��
'             ˫���������������*���ƣ�*���룬*����ģ�����ʾ���ϽǵĲ�ѯ���ܣ��Թ���ѯ��Ŀ��
'                    �����б�������ƥ�䣬ƥ��ɹ���λ���÷���ĸ���Ŀ�ϣ���F3֧�ֲ�����һ����
'     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
'     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
'     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
'             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
'             bytStyle=1ʱ,�����Ǳ��������
'     strNote=ѡ������˵������
'     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
'     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
'     blnNoneWin=����ɷǴ�����
'     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
'     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
'     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
'     blnMulti=�Ƿ������ѡ
'     strOtherInfo=��ʽΪ����Ŀ����1=����1#��Ŀ2=����2#......
'                ��ǰ��Ŀ�У�bytSize=0,1;�����С(0-С����,1-������;С����Ϊ9����,������Ϊ12����),Ĭ��С����
'                            ColSet=�п�����|����1,���1,0;����2,���2,1;.....|������ʾ|������ ���п�Ⱥ����һ��������ʾ���еĶ��뷽ʽ,0��1��2�ֱ��ʾ����롢�Ҷ�����м����
'                            NotShowNon=0,1;0-Ĭ�ϴ�����ʾû������ķ��࣬1-����ʾû������ķ���;bytStyle=2������
'                            HeadCap=SQL����1,�б�չʾ����1;SQL����2,�б�չʾ����2������Ŀ�����ֹ�ָ��SQL�����б���չʾ���ƣ�һ�����ڱ��������У����ǲ��ı��е�Key
'                            MultiCheckReturn=0,1����ѡʱֻ���ع�ѡ�У����ڶ�ѡ��ȷ��Ĭ�Ϸ��ص�ǰ���������Ӹò������ƣ��ÿ������ú󣬲�֧��Ĭ���еķ��أ������Ծ�֧��˫�����Զ����ء�
'                            HideNullCols=0,1;�Ƿ�����SQl�е�null as д������
'     arrInput=��Ӧ�ĸ���SQL����ֵ,��˳����,����Ϊ��ȷ����
'���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
'˵����
'     1.ID���ϼ�ID����Ϊ�ַ�������
'     2.ĩ�����ֶβ�Ҫ����ֵ
'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    
    Dim arrInfo As Variant, arrTmp As Variant, arrTmp2 As Variant
    Dim i As Long, j As Long
    Dim lngH As Long, lngW As Long, vRect As RECT, sngX As Single, sngY As Single
    Dim vPoint As PointAPI
    
    mstrSQL = strSQL
    mstrColWith = ""
    mstrTipCol = ""
    mblnNotShowNon = False
    mbytSize = 0
    '����strOtherInfoInfo
    arrInfo = Split(strOtherInfo, "#")
    For i = LBound(arrInfo) To UBound(arrInfo)
        If Trim(arrInfo(i)) <> "" Then
            arrTmp = Split(Trim(arrInfo(i)), "=")
            If UBound(arrTmp) = 1 Then
                Select Case UCase(arrTmp(0))
                    Case "BYTSIZE" '����
                        mbytSize = Val(arrTmp(1))
                    Case "COLSET" '�п�������������
                        arrTmp2 = Split(arrTmp(1), "|")
                        For j = LBound(arrTmp) To UBound(arrTmp) Step 2
                            If arrTmp2(j) = "�п�����" And bytStyle <> 1 Then
                                mstrColWith = arrTmp2(j + 1)
                            ElseIf arrTmp2(j) = "������ʾ" Then
                                mstrTipCol = arrTmp2(j + 1)
                            End If
                        Next
                    Case "NOTSHOWNON" '����ʾû������ķ���
                        If bytStyle = 2 Then mblnNotShowNon = Val(arrTmp(1))
                    Case "HEADCAP"
                        mstrHeadCap = arrTmp(1)
                    Case "MULTICHECKRETURN"
                        mblnMultiCheckReturn = Val(arrTmp(1))
                    Case "HIDENULLCOLS"
                        mblnHideNullCols = Val(arrTmp(1))
                End Select
            End If
        End If
    Next
    'ͨ��Api������ؼ������������Ϣ
    If Not objControl Is Nothing Then
        Select Case UCase(TypeName(objControl))
            Case UCase("VSFlexGrid")
                vPoint = zlControl.GetClientPoint(objControl.hwnd)
                sngX = vPoint.X
                sngY = vPoint.Y + objControl.Height
                lngH = objControl.CellHeight
                lngW = objControl.CellWidth
                sngY = sngY - lngH
            Case UCase("BILLEDIT")
                vPoint = zlControl.GetClientPoint(objControl.MsfObj.hwnd)
                sngX = vPoint.X
                sngY = vPoint.Y + objControl.MsfObj.Height
                lngH = objControl.MsfObj.CellHeight
                lngW = objControl.MsfObj.CellWidth
            Case Else
                vRect = zlControl.GetControlRect(objControl.hwnd)
                sngX = vRect.Left - 15
                sngY = vRect.Top
                lngH = objControl.Height
                lngW = objControl.Width
        End Select
    End If
    mlngX = sngX: mlngY = sngY: mlngTxtH = lngH
    marrInput = arrInput
    marrHideCols = Array()
    If mblnHideNullCols Then
        Call GetHideCols '��ȡ��������
    End If
    mstrTitle = strTitle
    mstrNote = strNote
    mbytStyle = bytStyle
    mblnĩ�� = blnĩ��
    mstrSeek = strSeek
    mblnShowSub = blnShowSub
    mblnShowRoot = blnShowRoot
    mblnMultiOne = blnMultiOne
    mblnNoneWin = blnNoneWin
    mblnSearch = blnSearch
    mblnMulti = blnMulti
    If Not frmParent Is Nothing Then
        mstrSaveTag = frmParent.Name & "_" & strTitle & "_" & bytStyle & IIF(blnNoneWin, 0, 1)
    Else
        mstrSaveTag = strTitle & "_" & bytStyle & IIF(blnNoneWin, 0, 1)
    End If
    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    If mblnOK Then
        Cancel = False
        Set ShowSelectV2 = mrsSel
    Else
        Cancel = True
        Set ShowSelectV2 = Nothing
    End If
End Function

Private Sub chkShowChild_Click()
    mblnShowSub = chkShowChild.value = 1
    If Not tvw_s.SelectedItem Is Nothing Then mstrKey = "": Call tvw_s_NodeClick(tvw_s.SelectedItem)
End Sub

Private Sub cmdCancel_Click()
    Set mrsSel = Nothing
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Dim i As Long
    
    For i = 1 To lvw.ListItems.Count
        lvw.ListItems(i).Checked = False
        If mbytStyle = 2 Then
            mrsSel.Filter = "ID='" & Split(lvw.ListItems(i).Key, "_")(1) & "' And ĩ��=1"
        Else
            mrsSel.Filter = "ID='" & Split(lvw.ListItems(i).Key, "_")(1) & "'"
        End If
        If mrsSel.RecordCount > 0 Then
            mrsSel.Update mstrCheck, 0
        End If
    Next
    
    If Not lvw.SelectedItem Is Nothing Then
        Call lvw_ItemClick(lvw.SelectedItem)
    End If
End Sub

Private Sub cmdFind_Click()
    Dim strFind As String
    Dim int���� As Integer, i As Long, j As Long, k As Long
    Dim strFilter As String
    Dim strTmp As String, strTemp As String
    
    If txtFind.Text <> "" And mlngFindIndex > 0 Then
        With mrsDetail
'            strFilter = 0
'            .Filter = "ĩ��=1"
            If .RecordCount > 0 Then .AbsolutePosition = mlngFindIndex
            strFind = UCase(Trim(txtFind.Text))
            If zlcommfun.IsCharChinese(txtFind.Text) Then
                '���ĵ�ֻ������
                int���� = 1
            ElseIf zlcommfun.IsCharAlpha(txtFind.Text) Then
                'Ӣ�Ĳ����ƺͼ���
                int���� = 2
            Else
                '��������Ƽ���ͱ���
                int���� = 3
            End If
            For i = mlngFindIndex To .RecordCount
                If int���� = 1 Then
                    For j = 0 To UBound(Split(mstrFind, ","))
                        If Split(mstrFind, ",")(j) Like "*����" Then
                            If .Fields(Split(mstrFind, ",")(j)).value & "" Like "*" & strFind & "*" Then
                                mlngFindIndex = i + 1
                                strTmp = CStr(!ID)
                                '75926,Ƚ����,2014-7-28
                                strTemp = "_" & zlcommfun.NVL(!�ϼ�ID)
                                If strTemp = "_" Then strTemp = "Root"
                                tvw_s.Nodes(strTemp).Selected = True
                                tvw_s.Nodes(strTemp).EnsureVisible
                                tvw_s.Nodes(strTemp).Tag = "ֱ�ӵ���"
                                Call tvw_s_NodeClick(tvw_s.Nodes(strTemp))
                                If mblnSearch = True Then
                                    For k = 1 To lvw.ListItems.Count
                                        If lvw.ListItems.Item(k).Key Like "*_" & strTmp Then
                                            lvw.ListItems.Item(k).Selected = True
                                            lvw.ListItems.Item(k).EnsureVisible
                                            Exit For
                                        End If
                                    Next
                                Else
                                    lvw.ListItems("_" & strTmp).Selected = True
                                    lvw.ListItems("_" & strTmp).EnsureVisible
                                End If
                                .Filter = strFilter
                                Exit Sub
                            End If
                        End If
                    Next
                ElseIf int���� = 2 Then
                    For j = 0 To UBound(Split(mstrFind, ","))
                        If Split(mstrFind, ",")(j) Like "*����" Or Split(mstrFind, ",")(j) Like "*����" Then
                            If .Fields(Split(mstrFind, ",")(j)).value & "" Like "*" & strFind & "*" Then
                                mlngFindIndex = i + 1
                                strTmp = CStr(!ID)
                                '75926,Ƚ����,2014-7-28
                                strTemp = "_" & zlcommfun.NVL(!�ϼ�ID)
                                If strTemp = "_" Then strTemp = "Root"
                                tvw_s.Nodes(strTemp).Selected = True
                                tvw_s.Nodes(strTemp).EnsureVisible
                                tvw_s.Nodes(strTemp).Tag = "ֱ�ӵ���"
                                Call tvw_s_NodeClick(tvw_s.Nodes(strTemp))
                                If mblnSearch = True Then
                                    For k = 1 To lvw.ListItems.Count
                                        If lvw.ListItems.Item(k).Key Like "*_" & strTmp Then
                                            lvw.ListItems.Item(k).Selected = True
                                            lvw.ListItems.Item(k).EnsureVisible
                                            Exit For
                                        End If
                                    Next
                                Else
                                    lvw.ListItems("_" & strTmp).Selected = True
                                    lvw.ListItems("_" & strTmp).EnsureVisible
                                End If
                                .Filter = strFilter
                                Exit Sub
                            End If
                        End If
                    Next
                Else
                    For j = 0 To UBound(Split(mstrFind, ","))
                        If Split(mstrFind, ",")(j) Like "*����" Or Split(mstrFind, ",")(j) Like "*����" Then
                            If .Fields(Split(mstrFind, ",")(j)).value & "" Like "*" & strFind & "*" Then
                                mlngFindIndex = i + 1
                                strTmp = CStr(!ID)
                                '75926,Ƚ����,2014-7-28
                                strTemp = "_" & zlcommfun.NVL(!�ϼ�ID)
                                If strTemp = "_" Then strTemp = "Root"
                                tvw_s.Nodes(strTemp).Selected = True
                                tvw_s.Nodes(strTemp).EnsureVisible
                                tvw_s.Nodes(strTemp).Tag = "ֱ�ӵ���"
                                Call tvw_s_NodeClick(tvw_s.Nodes(strTemp))
                                If mblnSearch = True Then
                                    For k = 1 To lvw.ListItems.Count
                                        If lvw.ListItems.Item(k).Key Like "*_" & strTmp Then
                                            lvw.ListItems.Item(k).Selected = True
                                            lvw.ListItems.Item(k).EnsureVisible
                                            Exit For
                                        End If
                                    Next
                                Else
                                    lvw.ListItems("_" & strTmp).Selected = True
                                    lvw.ListItems("_" & strTmp).EnsureVisible
                                End If
                                .Filter = strFilter
                                Exit Sub
                            End If
                        ElseIf Split(mstrFind, ",")(j) Like "*����" Then
                            If .Fields(Split(mstrFind, ",")(j)).value & "" = strFind Then
                                mlngFindIndex = i + 1
                                strTmp = CStr(!ID)
                                '75926,Ƚ����,2014-7-28
                                strTemp = "_" & zlcommfun.NVL(!�ϼ�ID)
                                If strTemp = "_" Then strTemp = "Root"
                                tvw_s.Nodes(strTemp).Selected = True
                                tvw_s.Nodes(strTemp).EnsureVisible
                                tvw_s.Nodes(strTemp).Tag = "ֱ�ӵ���"
                                Call tvw_s_NodeClick(tvw_s.Nodes(strTemp))
                                If mblnSearch = True Then
                                    For k = 1 To lvw.ListItems.Count
                                        If lvw.ListItems.Item(k).Key Like "*_" & strTmp Then
                                            lvw.ListItems.Item(k).Selected = True
                                            lvw.ListItems.Item(k).EnsureVisible
                                            Exit For
                                        End If
                                    Next
                                Else
                                    lvw.ListItems("_" & strTmp).Selected = True
                                    lvw.ListItems("_" & strTmp).EnsureVisible
                                End If
                                .Filter = strFilter
                                Exit Sub
                            End If
                        End If
                    Next
                End If
                .MoveNext
            Next
            If mlngFindIndex = 1 Then
                MsgBox "δ�ҵ�����ѯ����Ŀ��", vbInformation, Me.Caption
            ElseIf mlngFindIndex <> 1 Then
                MsgBox "�Ѿ����������һ����Ŀ�ˡ�", vbInformation, Me.Caption
                mlngFindIndex = 1
            End If
            .Filter = strFilter
        End With
    End If
End Sub

Private Sub cmdOK_Click()
    If mrsDetail Is Nothing Then Exit Sub
    If mrsDetail.RecordCount = 0 Then Exit Sub
    
    If mblnĩ�� And mbytStyle = 1 Then
        If mrsDetail!ĩ�� <> 1 Then Exit Sub
    End If
    
    If mbytStyle = 1 Then
        mrsDetail.Update mstrCheck, 1
    ElseIf mblnMulti Then
        mrsDetail.Filter = mstrCheck & "= 1"
    ElseIf Not lvw.SelectedItem Is Nothing Then
        If mbytStyle = 2 Then
            mrsDetail.Filter = "ID='" & Split(lvw.SelectedItem.Key, "_")(1) & "' And ĩ��=1"
        Else
            mrsDetail.Filter = "ID='" & Split(lvw.SelectedItem.Key, "_")(1) & "'"
        End If
'        mrsDetail.Update mstrCheck, 1
    End If
    If mblnHaveCheck = False Then
        Set mrsDetail = zlDatabase.CopyNewRec(mrsDetail, , mstrFields)
    End If
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdSelALL_Click()
    Dim i As Long
    
    For i = 1 To lvw.ListItems.Count
        lvw.ListItems(i).Checked = True
        If mbytStyle = 2 Then
            mrsSel.Filter = "ID='" & Split(lvw.ListItems(i).Key, "_")(1) & "' And ĩ��=1"
        Else
            mrsSel.Filter = "ID='" & Split(lvw.ListItems(i).Key, "_")(1) & "'"
        End If
        If mrsSel.RecordCount > 0 Then
            mrsSel.Update mstrCheck, 1
        End If
    Next
    
    If Not lvw.SelectedItem Is Nothing Then
        Call lvw_ItemClick(lvw.SelectedItem)
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If lvw.Visible Then
        If lvw.ListItems.Count = 0 And tvw_s.Visible = True Then
            tvw_s.SetFocus
        Else
            lvw.SetFocus
        End If
    Else
        tvw_s.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And cmdOK.Enabled And Me.ActiveControl.Name <> "txtFind" And Me.ActiveControl.Name <> "cmdFind" Then
        cmdOK_Click
    ElseIf KeyCode = vbKeyEscape And cmdCancel.Enabled Then
        cmdCancel_Click
    ElseIf KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If lvw.Checkboxes Then Call cmdSelALL_Click
    ElseIf (KeyCode = vbKeyR Or KeyCode = vbKeyC) And Shift = vbCtrlMask Then
        If lvw.Checkboxes Then Call cmdClear_Click
    ElseIf KeyCode = vbKeyF3 Then
        cmdFind_Click
    End If
End Sub

Private Sub Form_Load()
    Dim lngScrW As Long, lngScrH As Long
    Dim lngColW As Long, i As Integer
    Dim blnSame As Boolean, strItemID As String
    Dim strCode As String, strName As String
    Dim objNode As Node
    Dim lngIndex As Long
    Dim arrCols As Variant, arrTmp As Variant
    Dim blnLike As Boolean '�Ƿ�������ƥ��
    
    Screen.MousePointer = 11
    
    On Error GoTo errH
    
    mblnOK = False
    mstrKey = ""
    mlngFindIndex = 1
    
    '���ÿؼ������С
    Call SetFontSize(Me, mbytSize)
    '��SQL���
    Set mrsSel = zlDatabase.OpenSQLRecordByArray(mstrSQL, Me.Caption, marrInput)
    If mrsSel.RecordCount > 0 Then mrsSel.MoveFirst
    '�̶���չ1λ������Ԥ��
    mlngMaxPar = UBound(marrInput) + 1
    '����չ10�������ַ�������4000�����
    For lngIndex = 0 To 9
        ReDim Preserve marrInput(UBound(marrInput) + 1)
        marrInput(UBound(marrInput)) = ""
    Next
    
    'û�������򷵻�
    If mrsSel.EOF Then
        Screen.MousePointer = 0
        Set mrsSel = Nothing
        mblnOK = True: Unload Me: Exit Sub
    End If
    
    If mstrSQL Like "*%*" Then
        blnLike = True
    Else
        For i = LBound(marrInput) To UBound(marrInput)
            If marrInput(i) Like "*%*" Then
                blnLike = True: Exit For
            End If
        Next
    End If
    '����ƥ��ʱ�Զ����ص����
    If blnLike Then
        If mrsSel.RecordCount = 1 Then 'ֻ��һ������
            Screen.MousePointer = 0
            mblnOK = True: Unload Me: Exit Sub
        ElseIf mblnMultiOne And mbytStyle = 0 Then '������ͬ����
            blnSame = True
            For i = 1 To mrsSel.RecordCount
                If i = 1 Then
                    strItemID = mrsSel!ID
                Else
                    If mrsSel!ID <> strItemID Then blnSame = False: Exit For
                End If
                mrsSel.MoveNext
            Next
            mrsSel.MoveFirst
            If blnSame Then
                Screen.MousePointer = 0
                mblnOK = True: Unload Me: Exit Sub
            End If
        End If
    End If
    
    '��¼��¼��ԭʼ�ֶΣ����ж��Ƿ���Check�ֶ�
    strCode = "": strName = ""
    mblnHaveCheck = False
    For i = 0 To mrsSel.Fields.Count - 1
        'ȷ�������ֶ�
        If mrsSel.Fields(i).Name = "����" Then
            strCode = "����"
        ElseIf mrsSel.Fields(i).Name = "����" Then
            strName = mrsSel.Fields(i).Name
        ElseIf mrsSel.Fields(i).Name = "����" And strName = "" Then
            strName = mrsSel.Fields(i).Name
        End If
        '�ж��Ƿ���Check�ֶ�
        If UCase(mrsSel.Fields(i).Name) = "CHECKID" Or UCase(mrsSel.Fields(i).Name) Like "*CHECK" Then
            mstrCheck = mrsSel.Fields(i).Name
            mblnHaveCheck = True
        End If
        mstrFields = IIF(mstrFields = "", "", mstrFields & ",") & mrsSel.Fields(i).Name
    Next
    If strName = "" Then strName = "����"
    '��û��Check�ֶΣ���ʹ��CopyNewRec���һ��Check�ֶΣ�
    '����Check�ֶΣ�ҲҪʹ��CopyNewRec����Ϊ����Ҫ��mrsSel���в�����Ҫ��Ϊ��̬�ġ�
    mrsSel.Filter = ""
    If mstrCheck = "" Then
        Set mrsSel = zlDatabase.CopyNewRec(mrsSel, , , Array("Zl_Check", adInteger, 1, Empty))
        mstrCheck = "Zl_Check"
    Else
        Set mrsSel = zlDatabase.CopyNewRec(mrsSel)
    End If
    
     'ɾ��û������ķ���
    If mblnNotShowNon Then Call DeleteNotHave
    
    If mstrNote <> "" And mbytStyle = 2 Then
        If InStr(1, UCase(mstrNote), "[COUNT]") > 0 Then
            mrsSel.Filter = "ĩ��=1"
            mstrNote = Replace(UCase(mstrNote), "[COUNT]", "[" & mrsSel.RecordCount & "]")
        End If
        For i = 0 To mrsSel.Fields.Count - 1
            If InStr(1, mstrNote, "[" & mrsSel.Fields(i).Name & "=") > 0 Then
                lngIndex = InStr(1, mstrNote, "[" & mrsSel.Fields(i).Name & "=") + Len(mrsSel.Fields(i).Name) + 1
                strCode = Mid(mstrNote, lngIndex)
                strCode = Mid(strCode, 1, InStr(1, strCode, "]") - 1)
                mrsSel.Filter = "ĩ��=1 And " & mrsSel.Fields(i).Name & strCode
                mstrNote = Replace(mstrNote, "[" & mrsSel.Fields(i).Name & strCode & "]", "[" & mrsSel.RecordCount & "]")
            End If
        Next i
        mrsSel.Filter = ""
    End If
    
    '���������֮ǰ����CheckBox��ʽ
    If mbytStyle <> 1 And mblnMulti Then
        lvw.Checkboxes = True
        cmdSelALL.Visible = True
        cmdClear.Visible = True
    End If
    
    '�������
    Select Case mbytStyle
        Case 0
            '������ͷ
            lvw.ColumnHeaders.Clear
            For i = 0 To mrsSel.Fields.Count - 1
                If (Not mrsSel.Fields(i).Name Like "*ID" Or mrsSel.Fields(i).Name = "����ID") And mrsSel.Fields(i).Name <> "ĩ��" And mrsSel.Fields(i).Name <> mstrCheck Then
                    lvw.ColumnHeaders.Add , "_" & mrsSel.Fields(i).Name, mrsSel.Fields(i).Name
                    If mrsSel.Fields(i).Name Like "*��*" Or mrsSel.Fields(i).Name Like "*��*" Then
                        lvw.ColumnHeaders(lvw.ColumnHeaders.Count).Alignment = lvwColumnRight
                    End If
                End If
            Next
            '��������
            arrCols = Split(mstrHeadCap, ";")
            For i = LBound(arrCols) To UBound(arrCols)
                arrTmp = Split(arrCols(i), ",")
                lvw.ColumnHeaders("_" & Trim(arrTmp(0))).Text = arrTmp(1)
            Next
            
            If mblnSearch Then lvw.ColumnHeaders.Add , "_��", "��", , 2
            
            lvw.Sorted = False
            If mblnSearch Then lvw.ColumnHeaders(lvw.ColumnHeaders.Count).Position = 1
            
            lvw.ListItems.Clear
            Call FillList
        Case 1
            '������������
            Set objNode = tvw_s.Nodes.Add(, , "Root", "����" & mstrTitle, 1)
            objNode.Expanded = True
            objNode.Selected = True
            objNode.Sorted = True
            
            If Not mrsSel.EOF Then
                For i = 1 To mrsSel.RecordCount
                    If strCode <> "" Then
                        If IsNull(mrsSel!�ϼ�ID) Then
                            Set objNode = tvw_s.Nodes.Add("Root", 4, "_" & mrsSel!ID, IIF(IsNull(mrsSel!����), "", "[" & mrsSel!���� & "]") & mrsSel.Fields(strName).value, 1)
                        Else
                            Set objNode = tvw_s.Nodes.Add("_" & mrsSel!�ϼ�ID, 4, "_" & mrsSel!ID, IIF(IsNull(mrsSel!����), "", "[" & mrsSel!���� & "]") & mrsSel.Fields(strName).value, 1)
                        End If
                    Else
                        If IsNull(mrsSel!�ϼ�ID) Then
                            Set objNode = tvw_s.Nodes.Add("Root", 4, "_" & mrsSel!ID, mrsSel.Fields(strName).value, 1)
                        Else
                            Set objNode = tvw_s.Nodes.Add("_" & mrsSel!�ϼ�ID, 4, "_" & mrsSel!ID, mrsSel.Fields(strName).value, 1)
                        End If
                    End If
                    If objNode.Text Like "*" & mstrSeek & "*" And mstrSeek <> "" Then
                        objNode.Selected = True
                        objNode.Parent.Expanded = True
                    End If
                    objNode.Sorted = True
                    mrsSel.MoveNext
                Next
                If tvw_s.SelectedItem.Index = 1 Then tvw_s.Nodes(1).Child.Selected = True
            End If
            tvw_s.SelectedItem.EnsureVisible
            Call tvw_s_NodeClick(tvw_s.SelectedItem)
        Case 2
            '��ĩ����������
            Set objNode = tvw_s.Nodes.Add(, , "Root", "����" & mstrTitle, 1)
            objNode.Expanded = True
            objNode.Selected = True
            objNode.Sorted = True
            
            If Not mrsSel.EOF Then
                mrsSel.Filter = "ĩ��=0"
                For i = 1 To mrsSel.RecordCount
                    If strCode <> "" Then
                        If IsNull(mrsSel!�ϼ�ID) Then
                            Set objNode = tvw_s.Nodes.Add("Root", 4, "_" & mrsSel!ID, IIF(IsNull(mrsSel!����), "", "[" & mrsSel!���� & "]") & mrsSel.Fields(strName).value, 1)
                        Else
                            Set objNode = tvw_s.Nodes.Add("_" & mrsSel!�ϼ�ID, 4, "_" & mrsSel!ID, IIF(IsNull(mrsSel!����), "", "[" & mrsSel!���� & "]") & mrsSel.Fields(strName).value, 1)
                        End If
                    Else
                        If IsNull(mrsSel!�ϼ�ID) Then
                            Set objNode = tvw_s.Nodes.Add("Root", 4, "_" & mrsSel!ID, mrsSel.Fields(strName).value, 1)
                        Else
                            Set objNode = tvw_s.Nodes.Add("_" & mrsSel!�ϼ�ID, 4, "_" & mrsSel!ID, mrsSel.Fields(strName).value, 1)
                        End If
                    End If
                    objNode.Sorted = True
                    mrsSel.MoveNext
                Next
                If Not tvw_s.Nodes(1).Child Is Nothing Then tvw_s.Nodes(1).Child.Selected = True
            End If
            
            '������ͷ
            lvw.ColumnHeaders.Clear
            For i = 0 To mrsSel.Fields.Count - 1
                If (Not mrsSel.Fields(i).Name Like "*ID" Or mrsSel.Fields(i).Name = "����ID") And mrsSel.Fields(i).Name <> "ĩ��" And mrsSel.Fields(i).Name <> mstrCheck Then
                    lvw.ColumnHeaders.Add , "_" & mrsSel.Fields(i).Name, mrsSel.Fields(i).Name
                    If mrsSel.Fields(i).Name Like "*��*" Or mrsSel.Fields(i).Name Like "*��*" Then
                        lvw.ColumnHeaders(lvw.ColumnHeaders.Count).Alignment = lvwColumnRight
                    End If
                    If mrsSel.Fields(i).Name Like "*����" Or mrsSel.Fields(i).Name Like "*����" Or mrsSel.Fields(i).Name Like "*����" Then
                        mstrFind = mstrFind & "," & mrsSel.Fields(i).Name
                    End If
                End If
            Next
            
            '��������
            arrCols = Split(mstrHeadCap, ";")
            For i = LBound(arrCols) To UBound(arrCols)
                arrTmp = Split(arrCols(i), ",")
                lvw.ColumnHeaders("_" & Trim(arrTmp(0))).Text = arrTmp(1)
            Next
            
            mstrFind = Mid(mstrFind, 2)
            If mblnSearch Then lvw.ColumnHeaders.Add , "_��", "��", , 2
            lvw.Sorted = False
            If mblnSearch Then lvw.ColumnHeaders(lvw.ColumnHeaders.Count).Position = 1
            
            Call tvw_s_NodeClick(tvw_s.SelectedItem)
    End Select
    
    '���ÿؼ��ɼ���
    '---------------------------------------------------------------
    If mstrTitle <> "" Then
        Me.Caption = mstrTitle & "ѡ��"
    End If
    If mstrNote <> "" Then
        lblInfo.Caption = mstrNote
    End If
    If mblnNoneWin Then
        pic.Width = 30
        pic.BackColor = vbBlack
        pic.ZOrder
        picInfo.Visible = mbytStyle = 2 And mstrNote <> ""
        picCMD.Visible = False
        lvw.Appearance = ccFlat
        lvw.BorderStyle = ccFixedSingle
        tvw_s.Appearance = ccFlat
        tvw_s.BorderStyle = ccFixedSingle
    Else
        If mbytStyle <> 2 Then Me.Width = 5400 'ȱʡ���
        '������ʱ�������ؼ�λ��
        If mbytSize = 1 Then
            If mbytStyle <> 2 Then Me.Width = 7500: Me.Height = 5000
            If mbytStyle = 2 Then Me.Width = 9000: Me.Height = 6000
            
            picInfo.Height = picInfo.Height + 60
            lblInfo.Top = lblInfo.Top + 15
            
            chkShowChild.Top = chkShowChild.Top + 30
            chkShowChild.Left = lblInfo.Left + lblInfo.Width + 200
            
            txtFind.Height = 360: txtFind.Left = chkShowChild.Left + chkShowChild.Width + 200
            
            cmdFind.Height = 420: cmdFind.Width = 1300
            cmdFind.Top = cmdFind.Top - 50: cmdFind.Left = txtFind.Left + txtFind.Width + 50
            
            picCMD.Height = picCMD.Height + 30
            cmdSelALL.Height = 420: cmdSelALL.Width = 1500
            cmdSelALL.Top = cmdSelALL.Top - 30
            
            cmdClear.Height = 420: cmdClear.Width = 1500
            cmdClear.Top = cmdClear.Top - 30: cmdClear.Left = cmdSelALL.Left + cmdSelALL.Width + 20
            
            cmdOK.Height = 420: cmdOK.Width = 1500
            cmdOK.Top = cmdOK.Top - 30:
            
            cmdCancel.Height = 420: cmdCancel.Width = 1500
            cmdCancel.Top = cmdCancel.Top - 30
        End If
        Call RestoreWinState(Me, App.ProductName, mstrSaveTag)
    End If
    Select Case mbytStyle
        Case 0
            lvw.Visible = True
            tvw_s.Visible = False
            pic.Visible = False
            cmdFind.Visible = False
            txtFind.Visible = False
        Case 1
            lvw.Visible = False
            tvw_s.Visible = True
            pic.Visible = False
            cmdFind.Visible = False
            txtFind.Visible = False
        Case 2
            lvw.Visible = True
            tvw_s.Visible = True
            pic.Visible = True
            chkShowChild.Visible = True
            If mstrFind <> "" Then
                cmdFind.Visible = True
                txtFind.Visible = True
            End If
    End Select
    
    '��������ߴ�
    '---------------------------------------------------------------
    If mblnNoneWin Then
        Call zlControl.FormSetCaption(Me, False, False)
        Me.Left = mlngX
        
        arrCols = Split(mstrColWith, ";")
        For i = LBound(arrCols) To UBound(arrCols)
            arrTmp = Split(arrCols(i), ",")
            If Val(arrTmp(1)) <> 0 Then
                lvw.ColumnHeaders("_" & arrTmp(0)).Width = Val(arrTmp(1))
            End If
        Next

        If mbytStyle = 1 Then
            Me.Width = 3100
            If mbytSize = 1 Then Me.Width = Me.Width + 500
        Else
            If mbytSize = 1 Then tvw_s.Width = tvw_s.Width + 500
            lngScrW = GetSystemMetrics(SM_CXVSCROLL) * 15 + 75
            For i = 1 To lvw.ColumnHeaders.Count
                lngColW = lngColW + lvw.ColumnHeaders(i).Width
            Next
            If mbytStyle = 2 Then
                If lngColW < 1.5 * tvw_s.Width Then lngColW = 1.5 * tvw_s.Width
                lngColW = lngColW + tvw_s.Width
                If mstrNote <> "" Then '��ʾ�����ޱ߿�ʱ��ʾpicInfo�߿�
                    picInfo.BorderStyle = 1
                    If Me.Width < picInfo.Width Then
                        Me.Width = picInfo.Width
                    End If
                    If Me.Left + Me.Width > Screen.Width Then
                        Me.Left = Screen.Width - Me.Width
                    End If
                End If
            End If
            
            If Me.Left + lngColW + lngScrW > Screen.Width Then
                lngColW = 0
                For i = 1 To lvw.ColumnHeaders.Count
                    If InStr(";" & mstrColWith, ";" & lvw.ColumnHeaders(i).Text & ",") = 0 Then
                        If lvw.ColumnHeaders(i).Width > IIF(mbytSize = 1, 2400, 1800) Then lvw.ColumnHeaders(i).Width = IIF(mbytSize = 1, 2400, 1800)
                    End If
                    lngColW = lngColW + lvw.ColumnHeaders(i).Width
                Next
                If Me.Left + lngColW + lngScrW > Screen.Width Then
                    Me.Width = Screen.Width - Me.Left
                Else
                    Me.Width = lngColW + lngScrW
                End If
            Else
                If mstrNote <> "" And mbytStyle = 2 Then '�ޱ߿����в��ҵģ������п���Զ���Ӧ
                
                Else
                    Me.Width = lngColW + lngScrW
                End If
            End If
        End If
        
        Me.Height = 3240
        lngScrH = GetSystemMetrics(SM_CYFULLSCREEN) * 15 '��Ļ���ø߶�
        If mlngY + mlngTxtH + Me.Height > lngScrH Then
            Me.Top = mlngY - Me.Height
        Else
            Me.Top = mlngY + mlngTxtH
        End If
        Call RestoreListViewState(lvw, App.ProductName & "\" & Me.Name & mstrSaveTag, "")
    End If
    arrCols = Split(mstrColWith, ";")
    For i = LBound(arrCols) To UBound(arrCols)
        arrTmp = Split(arrCols(i), ",")
        If UBound(arrTmp) > 1 Then
            lvw.ColumnHeaders("_" & arrTmp(0)).Alignment = Val(arrTmp(2))
        End If
    Next
    Call Form_Resize
    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Unload Me
End Sub

Private Sub DeleteNotHave()
'���ܣ�ɾ��û������ķ���
    Dim i As Long
    Dim strFilter As String
    Dim rsTmp As Recordset
    Dim rsTmp1 As Recordset
    
    strFilter = mrsSel.Filter
    mrsSel.Filter = "ĩ��=1"
    Set rsTmp = zlDatabase.CopyNewRec(mrsSel)
    mrsSel.Filter = "ĩ��=0"
    Set rsTmp1 = zlDatabase.CopyNewRec(mrsSel)
    If mrsSel.RecordCount > 0 Then mrsSel.MoveFirst
    For i = mrsSel.RecordCount To 1 Step -1
        mrsSel.AbsolutePosition = i
        rsTmp1.Filter = "�ϼ�ID=" & mrsSel!ID & " And ID<>-1"
        rsTmp.Filter = "�ϼ�ID=" & mrsSel!ID
        If rsTmp1.RecordCount = 0 And rsTmp.RecordCount = 0 Then
            rsTmp1.Filter = "ID=" & mrsSel!ID
            rsTmp1!ID = "-1"
            mrsSel!ID = "-1"
        End If
    Next
    mrsSel.Filter = "ID=-1"
    Do While Not mrsSel.EOF
        mrsSel.Delete
        If mrsSel.RecordCount >= 0 Then mrsSel.MoveNext
    Loop
    mrsSel.Filter = IIF(strFilter = "0", 0, strFilter)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    
    Select Case mbytStyle
        Case 0 'ListView
            lvw.Top = IIF(picInfo.Visible, picInfo.Height, 0)
            lvw.Left = 0
            lvw.Width = Me.ScaleWidth
            lvw.Height = Me.ScaleHeight - IIF(picInfo.Visible, picInfo.Height, 0) - IIF(picCMD.Visible, picCMD.Height, 0)
        Case 1
            tvw_s.Top = IIF(picInfo.Visible, picInfo.Height, 0)
            tvw_s.Left = 0
            tvw_s.Width = Me.ScaleWidth
            tvw_s.Height = Me.ScaleHeight - IIF(picInfo.Visible, picInfo.Height, 0) - IIF(picCMD.Visible, picCMD.Height, 0)
        Case 2
            tvw_s.Left = 0
            tvw_s.Top = IIF(picInfo.Visible, picInfo.Height, 0)
            tvw_s.Height = Me.ScaleHeight - IIF(picInfo.Visible, picInfo.Height, 0) - IIF(picCMD.Visible, picCMD.Height, 0)
            
            pic.Top = tvw_s.Top
            pic.Height = tvw_s.Height
            lvw.Top = tvw_s.Top
            lvw.Height = tvw_s.Height
            
            If mblnNoneWin Then
                pic.Left = tvw_s.Width - pic.Width / 2
                lvw.Left = tvw_s.Width
                lvw.Width = Me.ScaleWidth - tvw_s.Width
            Else
                pic.Left = tvw_s.Width
                lvw.Left = tvw_s.Width + pic.Width
                lvw.Width = Me.ScaleWidth - tvw_s.Width - pic.Width
            End If
    End Select
    
    picBack.Left = lvw.Left
    picBack.Top = lvw.Top
    picBack.Width = lvw.Width
    picBack.Height = lvw.Height
    
    If Me.ScaleWidth - cmdCancel.Width * 1.3 - cmdOK.Width >= cmdClear.Left + cmdClear.Width * 1.3 Then
        cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width * 1.3
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 20
    End If
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, mstrSaveTag)
End Sub

Private Sub lvw_DblClick()
    '��ѡ����£�˫����ĿΪѡ����Ŀ
    If Not lvw.SelectedItem Is Nothing Then
        If mbytStyle = 2 Then
            mrsDetail.Filter = "ID='" & Split(lvw.SelectedItem.Key, "_")(1) & "' And ĩ��=1"
        Else
            mrsDetail.Filter = "ID='" & Split(lvw.SelectedItem.Key, "_")(1) & "'"
        End If
'        If mrsDetail.RecordCount > 0 Then
'            mrsDetail.Update mstrCheck, 1
'        End If
    End If
    cmdOK.Enabled = mrsDetail.RecordCount > 0
    If cmdOK.Enabled And Not lvw.SelectedItem Is Nothing Then
        cmdOK_Click
    End If
End Sub

Private Sub lvw_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    If Not Item Is Nothing Then
        If mbytStyle = 2 Then
            mrsSel.Filter = "ID='" & Split(Item.Key, "_")(1) & "' And ĩ��=1"
        Else
            mrsSel.Filter = "ID='" & Split(Item.Key, "_")(1) & "'"
        End If
        If mrsSel.RecordCount > 0 Then
            mrsSel.Update mstrCheck, IIF(Item.Checked, 1, 0)
        End If
    End If

    cmdOK.Enabled = mrsSel.RecordCount > 0
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmdOK.Enabled = mrsSel.RecordCount > 0
End Sub

Private Sub lvw_KeyPress(KeyAscii As Integer)
    Static strIdx As String
    Static sngTim As Single
    
    If mblnSearch Then
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            If Abs(Timer - sngTim) > 0.5 Then
                strIdx = ""
            End If
            sngTim = Timer
            strIdx = strIdx & Chr(KeyAscii)
            KeyAscii = 0
            
            If Len(strIdx) > 4 Then strIdx = Left(strIdx, 4)
            
            If lvw.ListItems.Count >= CInt(strIdx) And CInt(strIdx) > 0 Then
                lvw.ListItems(CInt(strIdx)).Selected = True
                lvw.SelectedItem.EnsureVisible
                Call lvw_ItemClick(lvw.SelectedItem)
            End If
        End If
    End If
End Sub

Private Sub lvw_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Dim objItem As ListItem
        If mstrTipCol <> "" Then
            Set objItem = lvw.HitTest(X, Y)
            If Not objItem Is Nothing Then
                Call zlcommfun.ShowTipInfo(lvw.hwnd, objItem.SubItems(lvw.ColumnHeaders("_" & mstrTipCol).Index - 1), True)
            Else
                Call zlcommfun.ShowTipInfo(lvw.hwnd, "")
            End If
        End If
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If tvw_s.Width + X < 1000 Or lvw.Width - X < 1000 Then Exit Sub
        pic.Left = pic.Left + X
        tvw_s.Width = tvw_s.Width + X
        lvw.Left = lvw.Left + X
        lvw.Width = lvw.Width - X
        picBack.Left = picBack.Left + X
        picBack.Width = picBack.Width - X
        Me.Refresh
    End If
End Sub

Private Sub FillList()
'���ܣ�װ��ListView����
    Dim i As Long, j As Long
    Dim objItem As ListItem
    Dim arrCols As Variant
    Dim arrTmp As Variant
    
    lvw.Visible = False
    Screen.MousePointer = 11
    For i = 1 To mrsSel.RecordCount
        For j = 0 To mrsSel.Fields.Count - 1
            If (Not mrsSel.Fields(j).Name Like "*ID" Or mrsSel.Fields(j).Name = "����ID") And mrsSel.Fields(j).Name <> "ĩ��" And mrsSel.Fields(j).Name <> mstrCheck Then
                If lvw.ColumnHeaders("_" & mrsSel.Fields(j).Name).Index = 1 Then
                    If mblnSearch Then '�ؼ��ּ����к�
                        Set objItem = lvw.ListItems.Add(, i & "_" & mrsSel!ID, IIF(IsNull(mrsSel.Fields(j).value), "", mrsSel.Fields(j).value), , 1)
                    Else
                        Set objItem = lvw.ListItems.Add(, "_" & mrsSel!ID, IIF(IsNull(mrsSel.Fields(j).value), "", mrsSel.Fields(j).value), , 1)
                    End If
                    If objItem.Text Like "*" & mstrSeek & "*" And mstrSeek <> "" Then objItem.Selected = True
                Else
                    objItem.SubItems(lvw.ColumnHeaders("_" & mrsSel.Fields(j).Name).Index - 1) = IIF(IsNull(mrsSel.Fields(j).value), "", mrsSel.Fields(j).value)
                End If
            End If
        Next
        If mblnSearch Then objItem.SubItems(lvw.ColumnHeaders("_��").Index - 1) = i
        If mstrCheck <> "" Then
            objItem.Checked = Val(mrsSel.Fields(mstrCheck).value & "")
        End If
        mrsSel.MoveNext
    Next
    
    Call LvwSetColWidth(lvw, , mbytSize)
    '20031013:���������
    If lvw.Width > Screen.Width / 2 Then
        For i = 1 To lvw.ColumnHeaders.Count
            If InStr(";" & mstrColWith, ";" & lvw.ColumnHeaders(i).Text & ",") = 0 Then
                If lvw.ColumnHeaders(i).Width > 1800 Then lvw.ColumnHeaders(i).Width = 1800
            End If
        Next
    End If
    '����һЩ��
    For i = LBound(marrHideCols) To UBound(marrHideCols)
        lvw.ColumnHeaders("_" & marrHideCols(j)).Width = 0
    Next
    '�����п��Լ��ж��뷽ʽ
    arrCols = Split(mstrColWith, ";")
    For i = LBound(arrCols) To UBound(arrCols)
        arrTmp = Split(arrCols(i), ",")
        lvw.ColumnHeaders("_" & arrTmp(0)).Width = Val(arrTmp(1))
    Next

    If Not lvw.SelectedItem Is Nothing Then
        cmdOK.Enabled = True
        
        lvw.SelectedItem.EnsureVisible
        Call lvw_ItemClick(lvw.SelectedItem)
    Else
        cmdOK.Enabled = False
    End If
    lvw.Refresh
    lvw.Visible = True
    Screen.MousePointer = 0
End Sub

Private Sub tvw_s_DblClick()
    If cmdOK.Enabled And mbytStyle = 1 Then cmdOK_Click
End Sub

Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim strKeys As String, i As Integer
    Dim strFilter As String
    Dim strSQL As String
    Dim dbl����ID As Double
    Dim strIDs As String
    Dim varTmp As Variant
    Dim blnҩƷ As Boolean
    Dim strPar As String
    Dim strParTable As String
    Dim strTable As String
    Dim varArr As Variant
    
    If mstrKey = Node.Key Then Exit Sub
    mstrKey = Node.Key
    If Node.Tag = "ֱ�ӵ���" Then
        Node.Tag = ""
    Else
        mlngFindIndex = 1
    End If
    
    If mbytStyle = 1 Then
        If Node.Key <> "Root" Then
            If mrsSel.Fields("ID").Type = adVarChar Then
                mrsSel.Filter = "ID='" & Mid(Node.Key, 2) & "'"
            Else
                mrsSel.Filter = "ID=" & Mid(Node.Key, 2)
            End If
            If mblnĩ�� Then
                cmdOK.Enabled = (mrsSel!ĩ�� = 1)
            Else
                cmdOK.Enabled = True
            End If
        Else
            cmdOK.Enabled = False
        End If
    ElseIf mbytStyle = 2 Then
        lvw.ListItems.Clear
        If Node.Key = "Root" Then
            '�����֧����,���κ�����
            If mblnShowRoot Then
                mrsSel.Filter = "ĩ��=1" '��������ʱ����
            Else
                mrsSel.Filter = "ĩ��=-1"
            End If
            If Visible Then lvw.SetFocus
        Else
            'ֻ���������,=id,instr(id
            'ע��������ҩƷ��������������������ϼ�ID,ҩƷ���Ǹ�����������������
            If mblnShowSub Then
                strKeys = GetSubTree(Node) '��������ʱ����
            Else
                strKeys = Mid(Node.Key, 2)
            End If
            varTmp = Split(strKeys, ",")
            For i = 0 To UBound(varTmp)
                If InStr(varTmp(i), "999999999") = 0 Then
                    blnҩƷ = Val(varTmp(i)) > 0 ' �ϼ�idΪ����ʱ
                    strIDs = strIDs & "," & Abs(Val(varTmp(i)))
                End If
            Next
            
            If strIDs <> "" Then
                If UBound(varTmp) = 0 Then
                    If blnҩƷ Then
                        strSQL = Replace(mstrDetail, "[ѡ���滻�Ĺ�����1]", " and (a.����id=[" & mlngMaxPar & "] or e.����id=0)")
                        strSQL = Replace(strSQL, "[ѡ���滻�Ĺ�����2]", " and e.����id=0")
                    Else
                        strSQL = Replace(mstrDetail, "[ѡ���滻�Ĺ�����1]", " and (a.����id=0 or e.����id=[" & mlngMaxPar & "])")
                        strSQL = Replace(strSQL, "[ѡ���滻�Ĺ�����2]", " and e.����id=[" & mlngMaxPar & "]")
                    End If
                    marrInput(mlngMaxPar - 1) = Val(Mid(strIDs, 2))
                Else
                    '��Ҫ���ǳ���4000���ȵ����
                    strPar = Mid(strIDs, 2)
                    marrInput(mlngMaxPar - 1) = Mid(strIDs, 2)
                    strParTable = "Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([" & mlngMaxPar & "]) As zlTools.t_Numlist)) X"
                    If Len(strPar) >= 4000 Then
                        varArr = Array()
                        varArr = GetParTable(strPar, strParTable, strTable)
                        For i = 0 To UBound(varArr)
                            marrInput(mlngMaxPar - 1 + i) = CStr(varArr(i))
                        Next
                        strParTable = strTable
                    End If
                    If blnҩƷ Then
                        strSQL = Replace(mstrDetail, "[ѡ���滻�Ĺ�����1]", "  and (a.����id in ( " & strParTable & " ) or e.����id=0)")
                        strSQL = Replace(strSQL, "[ѡ���滻�Ĺ�����2]", " and e.����id=0")
                    Else
                        strSQL = Replace(mstrDetail, "[ѡ���滻�Ĺ�����1]", " and (a.����id=0 or e.����id in (" & strParTable & " ))")
                        strSQL = Replace(strSQL, "[ѡ���滻�Ĺ�����2]", "  and e.����id in (" & strParTable & " )")
                    End If
                End If
                Set mrsDetail = zlDatabase.OpenSQLRecordByArray(strSQL, Me.Caption, marrInput)
                If Not mrsDetail.EOF Then Call FillListSQL
            End If
        End If
    End If
End Sub

Private Function GetSubTree(ByVal objNode As Node) As String
'���ܣ�����һ��������������Key(���ý��)
    Dim strKeys As String
    Dim objTmp As Node
    
    strKeys = "," & Mid(objNode.Key, 2) & strKeys
    Set objTmp = objNode.Child
    Do While Not objTmp Is Nothing
        If objTmp.Children > 0 Then
            strKeys = "," & GetSubTree(objTmp) & strKeys
        Else
            strKeys = "," & Mid(objTmp.Key, 2) & strKeys
        End If
        Set objTmp = objTmp.Next
    Loop
    GetSubTree = Mid(strKeys, 2)
End Function

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If mblnSearch And ColumnHeader.Key = "_��" Then Exit Sub
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvw.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvw.SortOrder = lvwDescending
    Else
        lvw.SortOrder = lvwAscending
    End If
    lvw.Sorted = True
        
    If mblnSearch Then
        For intIdx = 1 To lvw.ListItems.Count
            lvw.ListItems(intIdx).SubItems(lvw.ColumnHeaders("_��").Index - 1) = intIdx
        Next
    End If
    intIdx = ColumnHeader.Index
    If Not lvw.SelectedItem Is Nothing Then lvw.SelectedItem.EnsureVisible
End Sub

Private Sub txtFind_Change()
    mlngFindIndex = 1
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdFind_Click
    End If
End Sub

Private Sub GetHideCols()
'���ܣ�����SQL����ȡ�����ص���
'           NUll ���� �� NULL AS ���� �ſ�������
    Dim arrFileds As Variant
    Dim i As Long
    Dim strSQLTmp As String
    Dim arrTmp As Variant
    
    strSQLTmp = Replace(mstrSQL, vbCrLf, " ")
    strSQLTmp = Replace(strSQLTmp, vbLf, " ")
    strSQLTmp = Replace(strSQLTmp, vbCr, " ")
    strSQLTmp = Trim(Replace(strSQLTmp, vbTab, " "))
    'ȥ���ո�
    i = 5
    Do While i > 1
        strSQLTmp = Replace(strSQLTmp, String(i, " "), " ")
        If InStr(strSQLTmp, String(i, " ")) = 0 Then i = i - 1
    Loop
    strSQLTmp = UCase(strSQLTmp)
    arrFileds = Split(strSQLTmp, ",")
    '������ֵ��
    For i = LBound(arrFileds) To UBound(arrFileds)
        '���ֿ�������
        If Trim(arrFileds(i)) Like "NULL ?*" Or Trim(arrFileds(i)) Like "NULL AS ?*" Then
            arrTmp = Split(Trim(arrFileds(i)), " ")
            If arrTmp(UBound(arrTmp)) <> "" Then
                If Not arrTmp(UBound(arrTmp)) Like "*ID" And arrTmp(UBound(arrTmp)) <> "ĩ��" Or arrTmp(UBound(arrTmp)) = "����ID" Then
                    ReDim Preserve marrHideCols(UBound(marrHideCols) + 1)
                    marrHideCols(UBound(marrHideCols)) = arrTmp(UBound(arrTmp))
                End If
            End If
        End If
    Next
    
End Sub

Private Sub SetFontSize(ByVal objForm As Object, ByVal bytSize As Byte)
'���ܣ����ý���ؼ������С
'��Σ�objForm-�������
'      bytSize-�����С: 0-С����,1-������;С����Ϊ9����,������Ϊ12����
    Dim objCtl As Control
    
    On Error Resume Next
    For Each objCtl In objForm.Controls
        '0-С����,1-������;С����Ϊ9����,������Ϊ12����
        objCtl.Font.Size = IIF(bytSize = 1, 12, 9)
    Next
End Sub

Private Sub LvwSetColWidth(objLVW As Object, Optional blnHideNullCol As Boolean, Optional ByVal bytSize As Byte = 0)
'���ܣ�����ListView�е�ǰ�������Զ�������Ϊ��Сƥ����,���������ٿ�����ʾ��ͷ���ֵĿ��
'������objLvw=Ҫ������ListView����
'      blnHideNullCol=�Ƿ�����û���κ����ݵ���
'      bytSize=�����С��0-С����(9��) 1-������(12��)
    Dim i As Integer, lngAvgW As Long
    
    lngAvgW = IIF(bytSize = 1, 115, 90)
    For i = 1 To objLVW.ColumnHeaders.Count
        SendMessage objLVW.hwnd, LVM_SETCOLUMNWIDTH, i - 1, LVSCW_AUTOSIZE
        If blnHideNullCol Then If objLVW.ColumnHeaders(i).Width < 200 Then objLVW.ColumnHeaders(i).Width = 0
        If objLVW.ColumnHeaders(i).Width < (zlcommfun.ActualLen(objLVW.ColumnHeaders(i).Text) + 2) * lngAvgW And objLVW.ColumnHeaders(i).Width <> 0 Then
            objLVW.ColumnHeaders(i).Width = (zlcommfun.ActualLen(objLVW.ColumnHeaders(i).Text) + 2) * lngAvgW
        End If
    Next
End Sub

Private Sub FillListSQL()
'���ܣ�װ��ListView����
    Dim i As Long, j As Long
    Dim objItem As ListItem
    Dim arrCols As Variant
    Dim arrTmp As Variant
    
    lvw.Visible = False
    Screen.MousePointer = 11
    For i = 1 To mrsDetail.RecordCount
        For j = 0 To mrsDetail.Fields.Count - 1
            If (Not mrsDetail.Fields(j).Name Like "*ID" Or mrsDetail.Fields(j).Name = "����ID") And mrsDetail.Fields(j).Name <> "ĩ��" And mrsDetail.Fields(j).Name <> mstrCheck Then
                If lvw.ColumnHeaders("_" & mrsDetail.Fields(j).Name).Index = 1 Then
                    If mblnSearch Then '�ؼ��ּ����к�
                        Set objItem = lvw.ListItems.Add(, i & "_" & mrsDetail!ID, IIF(IsNull(mrsDetail.Fields(j).value), "", mrsDetail.Fields(j).value), , 1)
                    Else
                        Set objItem = lvw.ListItems.Add(, "_" & mrsDetail!ID, IIF(IsNull(mrsDetail.Fields(j).value), "", mrsDetail.Fields(j).value), , 1)
                    End If
                    If objItem.Text Like "*" & mstrSeek & "*" And mstrSeek <> "" Then objItem.Selected = True
                Else
                    objItem.SubItems(lvw.ColumnHeaders("_" & mrsDetail.Fields(j).Name).Index - 1) = IIF(IsNull(mrsDetail.Fields(j).value), "", mrsDetail.Fields(j).value)
                End If
            End If
        Next
        mrsDetail.MoveNext
    Next
    
    Call LvwSetColWidth(lvw, , mbytSize)
    '20031013:���������
    If lvw.Width > Screen.Width / 2 Then
        For i = 1 To lvw.ColumnHeaders.Count
            If InStr(";" & mstrColWith, ";" & lvw.ColumnHeaders(i).Text & ",") = 0 Then
                If lvw.ColumnHeaders(i).Width > 1800 Then lvw.ColumnHeaders(i).Width = 1800
            End If
        Next
    End If
    '����һЩ��
    For i = LBound(marrHideCols) To UBound(marrHideCols)
        lvw.ColumnHeaders("_" & marrHideCols(j)).Width = 0
    Next
    '�����п��Լ��ж��뷽ʽ
    arrCols = Split(mstrColWith, ";")
    For i = LBound(arrCols) To UBound(arrCols)
        arrTmp = Split(arrCols(i), ",")
        lvw.ColumnHeaders("_" & arrTmp(0)).Width = Val(arrTmp(1))
    Next

    If Not lvw.SelectedItem Is Nothing Then
        cmdOK.Enabled = True
        
        lvw.SelectedItem.EnsureVisible
        Call lvw_ItemClick(lvw.SelectedItem)
    Else
        cmdOK.Enabled = False
    End If
    lvw.Refresh
    lvw.Visible = True
    Screen.MousePointer = 0
End Sub

Private Function GetParTable(ByVal strPar As String, ByVal strParTable As String, ByRef strTableOut As String) As Variant
'���ܣ����ڶ�̬�ڴ��İ󶨲��������Ĵ���
'������strPar ��������strParTable �ڴ����ʽҪ����
'���أ�һ���ַ������飬10��Ԫ��
    Dim n As Long, p As Long
    Dim varPar(0 To 9) As String
    Dim strTable As String, strThis As String
    Dim intNum As Integer '������
    
    For n = 0 To 9
        varPar(n) = ""
    Next
    
    p = InStr(strParTable, "[") + 1
    intNum = Mid(strParTable, p, 1)
    
    n = 0
    Do While True
        If Len(strPar) < 4000 Then
            p = Len(strPar) + 1
        Else
            p = InStrRev(Mid(strPar, 1, 4000), ",")
        End If
        
        strThis = Mid(strPar, 1, p - 1)
        
        If n > 9 Then
            strTable = strTable & vbNewLine & " Union All " & Replace(strParTable, "[" & intNum & "]", "'" & strThis & "'")
        Else
            varPar(n) = strThis
            If n = 0 Then
                strTable = strParTable
            Else
                strTable = strTable & vbNewLine & " Union All " & Replace(strParTable, "[" & intNum & "]", "[" & (n + intNum) & "]")
            End If
        End If
        
        n = n + 1
        
        strPar = Mid(strPar, p + 1)
        
        If strPar = "" Then Exit Do
    Loop
    
    strTableOut = strTable
    GetParTable = varPar
    
End Function
