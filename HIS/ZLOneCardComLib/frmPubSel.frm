VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
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
Private marrHideCols   As Variant '�������ص��е�����

Private mstrSaveTag As String 'ע������ּ�
Private mrsBindings As ADODB.Recordset
Private marrInput() As Variant
Private mblnSearch As Boolean '�Ƿ�ͨ�������кż���
Private mblnNotShowNon As Boolean '����ʾû������ķ��࣬bytStyle=2
Private mstrHeadCap As String '������չʾ
Private mblnMultiCheckReturn As Boolean '��ѡʱ��ֻ����ѡ������˫����
Private mblnNoneWin As Boolean
Private mlngX As Long, mlngY As Long, mlngTxtH As Long
Private mstrCheck As String
Private mblnHaveCheck As Boolean '�ж�˫�б�ģʽ�£������ֶ����Ƿ���Check�ֶ�
Private mstrFields As String     '��¼��¼��ԭʼ�ֶ�
'���ڲ���
Private mrsSel As ADODB.Recordset
'�������
Private mblnOK As Boolean

Public Function ShowSelect(frmParent As Object, rsBindings As ADODB.Recordset, bytStyle As Byte, _
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
'     rsBindings=������Դ,��ͬ����ѡ�������ֶ��в�ͬҪ��
'     bytStyle=ѡ�������
'       Ϊ0ʱ:�б����:ID,��
'       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
'       Ϊ2ʱ:˫�����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
'             ˫���������������ں�Check��β���ֶΣ�����ֶ���Ϊ�Ƿ�ѡ��ֵ�洢�ֶΡ�=1Ϊ��ѡ��0=����ѡ��
'             ˫����������������*���ƣ�*���룬*����ģ�����ʾ���ϽǵĲ�ѯ���ܣ��Թ���ѯ��Ŀ��
'                    �����б�������ƥ�䣬ƥ��ɹ���λ���÷���ĸ���Ŀ�ϣ���F3֧�ֲ�����һ����
'     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
'     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
'     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
'             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
'             bytStyle=1ʱ,�����Ǳ��������
'     strNote=ѡ������˵������
'     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
'     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
'     blnNoneWin,X,Y,txtH=�����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
'     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
'     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
'     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
'     blnMulti=�Ƿ�������ѡ
'     arrInput=�������á����У�
'     arrInput��һ����������ǡ�����ʾû������ķ��ࡱ������ʾû������ķ���
'     arrInput�У�
'               ��ʽΪ��"bytSize=?"��ʾ���������С(0-С����,1-������;С����Ϊ9����,������Ϊ12����),Ĭ��С���塣
'               ��ʽΪ��ColSet:...ʱ��ʾ�п�����,ColSet��ʽ:�п�����|����1,����1;����2,����2.....|������ʾ|������
'               ��ʽΪ��HeadCap=SQL����1,�б�չʾ����1;SQL����2,�б�չʾ����2������Ŀ�����ֹ�ָ��SQL�����б���չʾ���ƣ�һ�����ڱ��������У����ǲ��ı��е�Key
'               ��ʽΪ��MultiCheckReturn=0,1����ѡʱֻ���ع�ѡ�У����ڶ�ѡ��ȷ��Ĭ�Ϸ��ص�ǰ���������Ӹò������ƣ��ÿ������ú󣬲�֧��Ĭ���еķ��أ������Ծ�֧��˫�����Զ����ء�
'               ��ʽΪ��HideCols=?���������ƣ������Ӣ�Ķ��ŷָ�����ʽ:����1,����2,...
'���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
'˵����
'     1.ID���ϼ�ID����Ϊ�ַ�������
'     2.ĩ�����ֶβ�Ҫ����ֵ
'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б��ȡ�
    
    Dim i As Long, j As Integer, arrTmp As Variant
    Dim blnPara As Boolean
    
    Set mrsBindings = rsBindings
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
                    ElseIf UCase(arrInput(i)) Like "HIDECOLS=*" Then
                        marrHideCols = Split(Split(arrInput(i), "=")(1), ","): blnPara = False
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
        mstrSaveTag = frmParent.Name & "_" & strTitle & "_" & bytStyle & IIf(blnNoneWin, 0, 1)
    Else
        mstrSaveTag = strTitle & "_" & bytStyle & IIf(blnNoneWin, 0, 1)
    End If
    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    If mblnOK Then
        Cancel = False
        Set ShowSelect = mrsSel
    Else
        Cancel = True
        Set ShowSelect = Nothing
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
    
    For i = 1 To lvw.ListItems.count
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
        With mrsSel
            strFilter = .Filter
            .Filter = "ĩ��=1"
            If .RecordCount > 0 Then .AbsolutePosition = mlngFindIndex
            strFind = UCase(Trim(txtFind.Text))
            If gobjComLib.zlCommFun.IsCharChinese(txtFind.Text) Then
                '���ĵ�ֻ������
                int���� = 1
            ElseIf gobjComLib.zlCommFun.IsCharAlpha(txtFind.Text) Then
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
                            If .fields(Split(mstrFind, ",")(j)).value & "" Like "*" & strFind & "*" Then
                                mlngFindIndex = i + 1
                                strTmp = CStr(!id)
                                '75926,Ƚ����,2014-7-28
                                strTemp = "_" & gobjComLib.zlCommFun.NVL(!�ϼ�ID)
                                If strTemp = "_" Then strTemp = "Root"
                                tvw_s.Nodes(strTemp).Selected = True
                                tvw_s.Nodes(strTemp).EnsureVisible
                                tvw_s.Nodes(strTemp).Tag = "ֱ�ӵ���"
                                Call tvw_s_NodeClick(tvw_s.Nodes(strTemp))
                                If mblnSearch = True Then
                                    For k = 1 To lvw.ListItems.count
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
                            If .fields(Split(mstrFind, ",")(j)).value & "" Like "*" & strFind & "*" Then
                                mlngFindIndex = i + 1
                                strTmp = CStr(!id)
                                '75926,Ƚ����,2014-7-28
                                strTemp = "_" & gobjComLib.zlCommFun.NVL(!�ϼ�ID)
                                If strTemp = "_" Then strTemp = "Root"
                                tvw_s.Nodes(strTemp).Selected = True
                                tvw_s.Nodes(strTemp).EnsureVisible
                                tvw_s.Nodes(strTemp).Tag = "ֱ�ӵ���"
                                Call tvw_s_NodeClick(tvw_s.Nodes(strTemp))
                                If mblnSearch = True Then
                                    For k = 1 To lvw.ListItems.count
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
                            If .fields(Split(mstrFind, ",")(j)).value & "" Like "*" & strFind & "*" Then
                                mlngFindIndex = i + 1
                                strTmp = CStr(!id)
                                '75926,Ƚ����,2014-7-28
                                strTemp = "_" & gobjComLib.zlCommFun.NVL(!�ϼ�ID)
                                If strTemp = "_" Then strTemp = "Root"
                                tvw_s.Nodes(strTemp).Selected = True
                                tvw_s.Nodes(strTemp).EnsureVisible
                                tvw_s.Nodes(strTemp).Tag = "ֱ�ӵ���"
                                Call tvw_s_NodeClick(tvw_s.Nodes(strTemp))
                                If mblnSearch = True Then
                                    For k = 1 To lvw.ListItems.count
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
                            If .fields(Split(mstrFind, ",")(j)).value & "" = strFind Then
                                mlngFindIndex = i + 1
                                strTmp = CStr(!id)
                                '75926,Ƚ����,2014-7-28
                                strTemp = "_" & gobjComLib.zlCommFun.NVL(!�ϼ�ID)
                                If strTemp = "_" Then strTemp = "Root"
                                tvw_s.Nodes(strTemp).Selected = True
                                tvw_s.Nodes(strTemp).EnsureVisible
                                tvw_s.Nodes(strTemp).Tag = "ֱ�ӵ���"
                                Call tvw_s_NodeClick(tvw_s.Nodes(strTemp))
                                If mblnSearch = True Then
                                    For k = 1 To lvw.ListItems.count
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
    If mrsSel Is Nothing Then Exit Sub
        SavePicture img16.ListImages(1).Picture, "D:\temp.bmp"
    If mbytStyle = 1 Then
    ElseIf mblnMulti Then
        mrsSel.Filter = mstrCheck & "= 1"
    ElseIf Not lvw.SelectedItem Is Nothing Then
        If mbytStyle = 2 Then
            mrsSel.Filter = "ID='" & Split(lvw.SelectedItem.Key, "_")(1) & "' And ĩ��=1"
        Else
            mrsSel.Filter = "ID='" & Split(lvw.SelectedItem.Key, "_")(1) & "'"
        End If
        If Not mrsSel.EOF Then
            mrsSel.Update mstrCheck, 1
        End If
    End If
    If mrsSel.RecordCount = 0 Then Exit Sub
    If mbytStyle = 1 Then
        If mblnĩ�� Then
            If mrsSel!ĩ�� <> 1 Then Exit Sub
        End If
        mrsSel.Update mstrCheck, 1
    End If
    
    If mblnHaveCheck = False Then
        Set mrsSel = gobjComLib.zlDatabase.CopyNewRec(mrsSel, , mstrFields)
    End If
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdSelALL_Click()
    Dim i As Long
    
    For i = 1 To lvw.ListItems.count
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
        If lvw.ListItems.count = 0 And tvw_s.Visible = True Then
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
    Dim strSort As String
    
    Screen.MousePointer = 11
    
    On Error GoTo ErrH
    
    mblnOK = False
    mstrKey = ""
    mlngFindIndex = 1
    
    '���ÿؼ������С
    Call SetFontSize(Me, mbytSize)
    
    Set mrsSel = mrsBindings
    If mrsSel.RecordCount > 0 Then mrsSel.MoveFirst
    'û�������򷵻�
    If mrsSel.EOF Then
        Screen.MousePointer = 0
        Set mrsSel = Nothing
        mblnOK = True: Unload Me: Exit Sub
    End If
    
    '��¼��¼��ԭʼ�ֶΣ����ж��Ƿ���Check�ֶ�
    strCode = "": strName = ""
    mblnHaveCheck = False
    For i = 0 To mrsSel.fields.count - 1
        'ȷ�������ֶ�
        If mrsSel.fields(i).Name = "����" Then
            strCode = "����"
        ElseIf mrsSel.fields(i).Name = "����" Then
            strName = mrsSel.fields(i).Name
        ElseIf mrsSel.fields(i).Name = "����" And strName = "" Then
            strName = mrsSel.fields(i).Name
        End If
        '�ж��Ƿ���Check�ֶ�
        If UCase(mrsSel.fields(i).Name) = "CHECKID" Or UCase(mrsSel.fields(i).Name) Like "*CHECK" Then
            mstrCheck = mrsSel.fields(i).Name
            mblnHaveCheck = True
        End If
        mstrFields = IIf(mstrFields = "", "", mstrFields & ",") & mrsSel.fields(i).Name
    Next
    If strName = "" Then strName = "����"
    '��û��Check�ֶΣ���ʹ��CopyNewRec����һ��Check�ֶΣ�
    '����Check�ֶΣ�ҲҪʹ��CopyNewRec����Ϊ����Ҫ��mrsSel���в�����Ҫ��Ϊ��̬�ġ�
    mrsSel.Filter = ""
    strSort = mrsSel.Sort
    If mstrCheck = "" Then
        Set mrsSel = gobjComLib.zlDatabase.CopyNewRec(mrsSel, , , Array("Zl_Check", adInteger, 1, Empty))
        mstrCheck = "Zl_Check"
    Else
        Set mrsSel = gobjComLib.zlDatabase.CopyNewRec(mrsSel)
    End If
    mrsSel.Sort = strSort '�ָ�����
    
     'ɾ��û������ķ���
    If mblnNotShowNon Then Call DeleteNotHave
    
    If mstrNote <> "" And mbytStyle = 2 Then
        If InStr(1, UCase(mstrNote), "[COUNT]") > 0 Then
            mrsSel.Filter = "ĩ��=1"
            mstrNote = Replace(UCase(mstrNote), "[COUNT]", "[" & mrsSel.RecordCount & "]")
        End If
        For i = 0 To mrsSel.fields.count - 1
            If InStr(1, mstrNote, "[" & mrsSel.fields(i).Name & "=") > 0 Then
                lngIndex = InStr(1, mstrNote, "[" & mrsSel.fields(i).Name & "=") + Len(mrsSel.fields(i).Name) + 1
                strCode = Mid(mstrNote, lngIndex)
                strCode = Mid(strCode, 1, InStr(1, strCode, "]") - 1)
                mrsSel.Filter = "ĩ��=1 And " & mrsSel.fields(i).Name & strCode
                mstrNote = Replace(mstrNote, "[" & mrsSel.fields(i).Name & strCode & "]", "[" & mrsSel.RecordCount & "]")
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
            For i = 0 To mrsSel.fields.count - 1
                If (Not mrsSel.fields(i).Name Like "*ID" Or mrsSel.fields(i).Name = "����ID") And mrsSel.fields(i).Name <> "ĩ��" And mrsSel.fields(i).Name <> mstrCheck Then
                    lvw.ColumnHeaders.Add , "_" & mrsSel.fields(i).Name, mrsSel.fields(i).Name
                    If mrsSel.fields(i).Name Like "*��*" Or mrsSel.fields(i).Name Like "*��*" Then
                        lvw.ColumnHeaders(lvw.ColumnHeaders.count).Alignment = lvwColumnRight
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
            If mblnSearch Then lvw.ColumnHeaders(lvw.ColumnHeaders.count).Position = 1
            
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
                            Set objNode = tvw_s.Nodes.Add("Root", 4, "_" & mrsSel!id, IIf(IsNull(mrsSel!����), "", "[" & mrsSel!���� & "]") & mrsSel.fields(strName).value, 1)
                        Else
                            Set objNode = tvw_s.Nodes.Add("_" & mrsSel!�ϼ�ID, 4, "_" & mrsSel!id, IIf(IsNull(mrsSel!����), "", "[" & mrsSel!���� & "]") & mrsSel.fields(strName).value, 1)
                        End If
                    Else
                        If IsNull(mrsSel!�ϼ�ID) Then
                            Set objNode = tvw_s.Nodes.Add("Root", 4, "_" & mrsSel!id, mrsSel.fields(strName).value, 1)
                        Else
                            Set objNode = tvw_s.Nodes.Add("_" & mrsSel!�ϼ�ID, 4, "_" & mrsSel!id, mrsSel.fields(strName).value, 1)
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
                            Set objNode = tvw_s.Nodes.Add("Root", 4, "_" & mrsSel!id, IIf(IsNull(mrsSel!����), "", "[" & mrsSel!���� & "]") & mrsSel.fields(strName).value, 1)
                        Else
                            Set objNode = tvw_s.Nodes.Add("_" & mrsSel!�ϼ�ID, 4, "_" & mrsSel!id, IIf(IsNull(mrsSel!����), "", "[" & mrsSel!���� & "]") & mrsSel.fields(strName).value, 1)
                        End If
                    Else
                        If IsNull(mrsSel!�ϼ�ID) Then
                            Set objNode = tvw_s.Nodes.Add("Root", 4, "_" & mrsSel!id, mrsSel.fields(strName).value, 1)
                        Else
                            Set objNode = tvw_s.Nodes.Add("_" & mrsSel!�ϼ�ID, 4, "_" & mrsSel!id, mrsSel.fields(strName).value, 1)
                        End If
                    End If
                    objNode.Sorted = True
                    mrsSel.MoveNext
                Next
                If Not tvw_s.Nodes(1).Child Is Nothing Then tvw_s.Nodes(1).Child.Selected = True
            End If
            
            '������ͷ
            lvw.ColumnHeaders.Clear
            For i = 0 To mrsSel.fields.count - 1
                If (Not mrsSel.fields(i).Name Like "*ID" Or mrsSel.fields(i).Name = "����ID") And mrsSel.fields(i).Name <> "ĩ��" And mrsSel.fields(i).Name <> mstrCheck Then
                    lvw.ColumnHeaders.Add , "_" & mrsSel.fields(i).Name, mrsSel.fields(i).Name
                    If mrsSel.fields(i).Name Like "*��*" Or mrsSel.fields(i).Name Like "*��*" Then
                        lvw.ColumnHeaders(lvw.ColumnHeaders.count).Alignment = lvwColumnRight
                    End If
                    If mrsSel.fields(i).Name Like "*����" Or mrsSel.fields(i).Name Like "*����" Or mrsSel.fields(i).Name Like "*����" Then
                        mstrFind = mstrFind & "," & mrsSel.fields(i).Name
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
            If mblnSearch Then lvw.ColumnHeaders(lvw.ColumnHeaders.count).Position = 1
            
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
        picCmd.Visible = False
        lvw.Appearance = ccFlat
        lvw.BorderStyle = ccFixedSingle
        tvw_s.Appearance = ccFlat
        tvw_s.BorderStyle = ccFixedSingle
    Else
        If mbytStyle <> 2 Then Me.Width = 5400 'ȱʡ����
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
            
            picCmd.Height = picCmd.Height + 30
            cmdSelALL.Height = 420: cmdSelALL.Width = 1500
            cmdSelALL.Top = cmdSelALL.Top - 30
            
            cmdClear.Height = 420: cmdClear.Width = 1500
            cmdClear.Top = cmdClear.Top - 30: cmdClear.Left = cmdSelALL.Left + cmdSelALL.Width + 20
            
            cmdOK.Height = 420: cmdOK.Width = 1500
            cmdOK.Top = cmdOK.Top - 30:
            
            cmdCancel.Height = 420: cmdCancel.Width = 1500
            cmdCancel.Top = cmdCancel.Top - 30
        End If
        Call gobjComLib.RestoreWinState(Me, App.ProductName, mstrSaveTag)
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
        Call gobjComLib.zlControl.FormSetCaption(Me, False, False)
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
            For i = 1 To lvw.ColumnHeaders.count
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
                For i = 1 To lvw.ColumnHeaders.count
                    If InStr(";" & mstrColWith, ";" & lvw.ColumnHeaders(i).Text & ",") = 0 Then
                        If lvw.ColumnHeaders(i).Width > IIf(mbytSize = 1, 2400, 1800) Then lvw.ColumnHeaders(i).Width = IIf(mbytSize = 1, 2400, 1800)
                    End If
                    lngColW = lngColW + lvw.ColumnHeaders(i).Width
                Next
                If Me.Left + lngColW + lngScrW > Screen.Width Then
                    Me.Width = Screen.Width - Me.Left
                Else
                    Me.Width = lngColW + lngScrW
                End If
            Else
                If mstrNote <> "" And mbytStyle = 2 Then '�ޱ߿����в��ҵģ������п����Զ���Ӧ
                
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
        Call gobjComLib.RestoreListViewState(lvw, App.ProductName & "\" & Me.Name & mstrSaveTag, "")
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
ErrH:
    Screen.MousePointer = 0
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
    Unload Me
End Sub

Private Sub DeleteNotHave()
'���ܣ�ɾ��û������ķ���
    Dim i As Long
    Dim strFilter As String
    Dim rsTmp As Recordset
    Dim rstmp1 As Recordset
    
    strFilter = mrsSel.Filter
    mrsSel.Filter = "ĩ��=1"
    Set rsTmp = gobjComLib.zlDatabase.CopyNewRec(mrsSel)
    mrsSel.Filter = "ĩ��=0"
    Set rstmp1 = gobjComLib.zlDatabase.CopyNewRec(mrsSel)
    If mrsSel.RecordCount > 0 Then mrsSel.MoveFirst
    For i = mrsSel.RecordCount To 1 Step -1
        mrsSel.AbsolutePosition = i
        rstmp1.Filter = "�ϼ�ID=" & mrsSel!id & " And ID<>-1"
        rsTmp.Filter = "�ϼ�ID=" & mrsSel!id
        If rstmp1.RecordCount = 0 And rsTmp.RecordCount = 0 Then
            rstmp1.Filter = "ID=" & mrsSel!id
            rstmp1!id = "-1"
            mrsSel!id = "-1"
        End If
    Next
    mrsSel.Filter = "ID=-1"
    Do While Not mrsSel.EOF
        mrsSel.Delete
        If mrsSel.RecordCount >= 0 Then mrsSel.MoveNext
    Loop
    mrsSel.Filter = IIf(strFilter = "0", 0, strFilter)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    
    Select Case mbytStyle
        Case 0 'ListView
            lvw.Top = IIf(picInfo.Visible, picInfo.Height, 0)
            lvw.Left = 0
            lvw.Width = Me.ScaleWidth
            lvw.Height = Me.ScaleHeight - IIf(picInfo.Visible, picInfo.Height, 0) - IIf(picCmd.Visible, picCmd.Height, 0)
        Case 1
            tvw_s.Top = IIf(picInfo.Visible, picInfo.Height, 0)
            tvw_s.Left = 0
            tvw_s.Width = Me.ScaleWidth
            tvw_s.Height = Me.ScaleHeight - IIf(picInfo.Visible, picInfo.Height, 0) - IIf(picCmd.Visible, picCmd.Height, 0)
        Case 2
            tvw_s.Left = 0
            tvw_s.Top = IIf(picInfo.Visible, picInfo.Height, 0)
            tvw_s.Height = Me.ScaleHeight - IIf(picInfo.Visible, picInfo.Height, 0) - IIf(picCmd.Visible, picCmd.Height, 0)
            
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
    Call gobjComLib.SaveWinState(Me, App.ProductName, mstrSaveTag)
End Sub

Private Sub lvw_DblClick()
    '��ѡ����£�˫����ĿΪѡ����Ŀ
    If Not lvw.SelectedItem Is Nothing Then
        If mbytStyle = 2 Then
            mrsSel.Filter = "ID='" & Split(lvw.SelectedItem.Key, "_")(1) & "' And ĩ��=1"
        Else
            mrsSel.Filter = "ID='" & Split(lvw.SelectedItem.Key, "_")(1) & "'"
        End If
        If mrsSel.RecordCount > 0 Then
            mrsSel.Update mstrCheck, 1
        End If
    End If
    cmdOK.Enabled = mrsSel.RecordCount > 0
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
            mrsSel.Update mstrCheck, IIf(Item.Checked, 1, 0)
        End If
    End If

    cmdOK.Enabled = mrsSel.RecordCount > 0
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Not mblnMulti Then
        cmdOK.Enabled = Not Item Is Nothing
    Else
        cmdOK.Enabled = mrsSel.RecordCount > 0
    End If
End Sub

Private Sub lvw_KeyPress(KeyAscii As Integer)
    Static strIdx As String
    Static sngTim As Single
    
    If mblnSearch Then
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            If Abs(timer - sngTim) > 0.5 Then
                strIdx = ""
            End If
            sngTim = timer
            strIdx = strIdx & Chr(KeyAscii)
            KeyAscii = 0
            
            If Len(strIdx) > 4 Then strIdx = Left(strIdx, 4)
            
            If lvw.ListItems.count >= CInt(strIdx) And CInt(strIdx) > 0 Then
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
                Call gobjComLib.zlCommFun.ShowTipInfo(lvw.hWnd, objItem.SubItems(lvw.ColumnHeaders("_" & mstrTipCol).Index - 1), True)
            Else
                Call gobjComLib.zlCommFun.ShowTipInfo(lvw.hWnd, "")
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
        For j = 0 To mrsSel.fields.count - 1
            If (Not mrsSel.fields(j).Name Like "*ID" Or mrsSel.fields(j).Name = "����ID") And mrsSel.fields(j).Name <> "ĩ��" And mrsSel.fields(j).Name <> mstrCheck Then
                If lvw.ColumnHeaders("_" & mrsSel.fields(j).Name).Index = 1 Then
                    If mblnSearch Then '�ؼ��ּ����к�
                        Set objItem = lvw.ListItems.Add(, i & "_" & mrsSel!id, IIf(IsNull(mrsSel.fields(j).value), "", mrsSel.fields(j).value), , 1)
                    Else
                        Set objItem = lvw.ListItems.Add(, "_" & mrsSel!id, IIf(IsNull(mrsSel.fields(j).value), "", mrsSel.fields(j).value), , 1)
                    End If
                    If objItem.Text Like "*" & mstrSeek & "*" And mstrSeek <> "" Then objItem.Selected = True
                Else
                    objItem.SubItems(lvw.ColumnHeaders("_" & mrsSel.fields(j).Name).Index - 1) = IIf(IsNull(mrsSel.fields(j).value), "", mrsSel.fields(j).value)
                End If
            End If
        Next
        If mblnSearch Then objItem.SubItems(lvw.ColumnHeaders("_��").Index - 1) = i
        If mstrCheck <> "" Then
            objItem.Checked = Val(mrsSel.fields(mstrCheck).value & "")
        End If
        mrsSel.MoveNext
    Next
    
    Call LvwSetColWidth(lvw, , mbytSize)
    '20031013:����������
    If lvw.Width > Screen.Width / 2 Then
        For i = 1 To lvw.ColumnHeaders.count
            If InStr(";" & mstrColWith, ";" & lvw.ColumnHeaders(i).Text & ",") = 0 Then
                If lvw.ColumnHeaders(i).Width > 1800 Then lvw.ColumnHeaders(i).Width = 1800
            End If
        Next
    End If
    '����һЩ��
    On Error Resume Next
    If TypeName(marrHideCols) <> "Empty" Then
        For i = LBound(marrHideCols) To UBound(marrHideCols)
            lvw.ColumnHeaders("_" & marrHideCols(i)).Width = 0
        Next
    End If
    
    '�����п��Լ��ж��뷽ʽ
    arrCols = Split(mstrColWith, ";")
    For i = LBound(arrCols) To UBound(arrCols)
        arrTmp = Split(arrCols(i), ",")
        lvw.ColumnHeaders("_" & arrTmp(0)).Width = Val(arrTmp(1))
    Next
    If Err.Number <> 0 Then Err.Clear
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
    Dim mstrKeys As String, i As Integer
    Dim strFilter As String
    
    If mstrKey = Node.Key Then Exit Sub
    mstrKey = Node.Key
    If Node.Tag = "ֱ�ӵ���" Then
        Node.Tag = ""
    Else
        mlngFindIndex = 1
    End If
    
    If mbytStyle = 1 Then
        If Node.Key <> "Root" Then
            If IsType(mrsSel.fields("ID").type, adVarChar) Then
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
            If mblnShowRoot Then
                mrsSel.Filter = "ĩ��=1" '��������ʱ����
            Else
                mrsSel.Filter = "ĩ��=-1"
            End If
            If Visible Then lvw.SetFocus
        Else
            If mblnShowSub Then
                mstrKeys = GetSubTree(Node) '��������ʱ����
            Else
                mstrKeys = Mid(Node.Key, 2)
            End If
            For i = 0 To UBound(Split(mstrKeys, ","))
                If IsType(mrsSel.fields("�ϼ�ID").type, adVarChar) Then
                    strFilter = strFilter & " Or (ĩ��=1 And �ϼ�ID='" & Split(mstrKeys, ",")(i) & "')"
                Else
                    strFilter = strFilter & " Or (ĩ��=1 And �ϼ�ID=" & Split(mstrKeys, ",")(i) & ")"
                End If
            Next
            strFilter = Mid(strFilter, 5)
            mrsSel.Filter = strFilter
        End If
        If Not mrsSel.EOF Then Call FillList
    End If
End Sub

Private Function GetSubTree(ByVal objNode As Node) As String
'���ܣ�����һ��������������Key(���ý��)
    Dim mstrKeys As String
    Dim objTmp As Node
    
    mstrKeys = "," & Mid(objNode.Key, 2) & mstrKeys
    Set objTmp = objNode.Child
    Do While Not objTmp Is Nothing
        If objTmp.children > 0 Then
            mstrKeys = "," & GetSubTree(objTmp) & mstrKeys
        Else
            mstrKeys = "," & Mid(objTmp.Key, 2) & mstrKeys
        End If
        Set objTmp = objTmp.Next
    Loop
    GetSubTree = Mid(mstrKeys, 2)
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
        For intIdx = 1 To lvw.ListItems.count
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

Private Sub SetFontSize(ByVal objForm As Object, ByVal bytSize As Byte)
'���ܣ����ý���ؼ������С
'��Σ�objForm-�������
'      bytSize-�����С: 0-С����,1-������;С����Ϊ9����,������Ϊ12����
    Dim objCtl As Control
    
    On Error Resume Next
    For Each objCtl In objForm.Controls
        '0-С����,1-������;С����Ϊ9����,������Ϊ12����
        objCtl.Font.Size = IIf(bytSize = 1, 12, 9)
    Next
End Sub

Private Sub LvwSetColWidth(objLvw As Object, Optional blnHideNullCol As Boolean, Optional ByVal bytSize As Byte = 0)
'���ܣ�����ListView�е�ǰ�������Զ�������Ϊ��Сƥ�����,���������ٿ�����ʾ��ͷ���ֵĿ���
'������objLvw=Ҫ������ListView����
'      blnHideNullCol=�Ƿ�����û���κ����ݵ���
'      bytSize=�����С��0-С����(9��) 1-������(12��)
    Dim i As Integer, lngAvgW As Long
    
    lngAvgW = IIf(bytSize = 1, 115, 90)
    For i = 1 To objLvw.ColumnHeaders.count
        SendMessage objLvw.hWnd, LVM_SETCOLUMNWIDTH, i - 1, LVSCW_AUTOSIZE
        If blnHideNullCol Then If objLvw.ColumnHeaders(i).Width < 200 Then objLvw.ColumnHeaders(i).Width = 0
        If objLvw.ColumnHeaders(i).Width < (gobjComLib.zlCommFun.ActualLen(objLvw.ColumnHeaders(i).Text) + 2) * lngAvgW And objLvw.ColumnHeaders(i).Width <> 0 Then
            objLvw.ColumnHeaders(i).Width = (gobjComLib.zlCommFun.ActualLen(objLvw.ColumnHeaders(i).Text) + 2) * lngAvgW
        End If
    Next
End Sub
