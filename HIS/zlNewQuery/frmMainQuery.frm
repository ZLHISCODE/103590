VERSION 5.00
Begin VB.Form frmMainQuery 
   BorderStyle     =   0  'None
   ClientHeight    =   5925
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8850
   Icon            =   "frmMainQuery.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   8850
   StartUpPosition =   3  '����ȱʡ
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrCheckConnect 
      Interval        =   60000
      Left            =   645
      Top             =   1035
   End
   Begin VB.Timer tmrHome 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   210
      Top             =   2535
   End
   Begin zl9NewQuery.ctlDefaultFrame FrameDefault 
      Height          =   4470
      Left            =   1065
      TabIndex        =   0
      Top             =   630
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   7885
   End
End
Attribute VB_Name = "frmMainQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mvarlngHome As Long                          '��ѯҳ�淵����ҳ��ʱ����
Private mvarBlnFirst As Boolean                      '�Ƿ��Ǹս��뱾ģ��

Public mvarHomeInternal As Long
Private mvarHomeLong As Long
Private mvarCheckConnectInternal As Long
Private mvarCheckConnectCounter As Long
Private mobjRegister As Object
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Form_Activate()
    On Error GoTo ErrHandle

    If mvarBlnFirst = False Then Exit Sub
    mvarBlnFirst = False
     
    mvarHomeInternal = 0
    tmrHome.Enabled = IIf(mvarHomeLong = 0, False, True)
    
    mvarCheckConnectInternal = Val(GetPara("����������Ӽ��ʱ��", "30"))
    tmrCheckConnect.Enabled = IIf(mvarCheckConnectInternal = 0, False, True)
    
    FrameDefault.AllowEdit = (InStr(gstrPrivs, "��Ϣά��") > 0)
    FrameDefault.AllowSelfRegist = (InStr(gstrPrivs, "�����Һ�") > 0)
    FrameDefault.AllowSelfPrint = (InStr(gstrPrivs, "������ӡ") > 0)
    FrameDefault.AllowFreeRegist = (InStr(gstrPrivs, "�����Һ�") > 0)
    
    Set gfrmMain = Me
    
    '2.װ�ز���ʾ��ҳ��
    Call FrameDefault.InitLoad
    
    Call FrameDefault.ShowHome
    
    DoEvents
    'zyk add 200410
    Call FrameDefault.showwww
    
    Dim wwwurl As String
    wwwurl = GetPara("ҽԺ��ҳ", "")
    If Not wwwurl = "" Then
        ShellExecute hwnd, "open", "iexplore.exe", "-k " & wwwurl, "", 1
        'Sleep 5000   'API������ʱ5000����
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '�������Esc,���˳�����ʾģ��
    Select Case KeyCode
    Case vbKeyEscape
        
        If Shift = vbShiftMask Then
            If Val(GetPara("�رղ�ѯ�������¼����", "0")) = 1 Then
                If frmExitPsw.ShowPsw(Me) Then
                    Unload Me
                End If
            Else
                Unload Me
            End If
        End If
        
    Case vbKey0, vbKey1, vbKey2, vbKey3, vbKey4, vbKey5, vbKey6, vbKey7, vbKey8, vbKey9, vbKeyNumpad0, vbKeyNumpad1, vbKeyNumpad2, vbKeyNumpad3, vbKeyNumpad4, vbKeyNumpad5, vbKeyNumpad6, vbKeyNumpad7, vbKeyNumpad8, vbKeyNumpad9
        'ֱ�ӵ��ò��˷��ò�ѯ
        
        gstrSQL = "select 1 from ��ѯҳ������ A,��ѯҳ��Ŀ¼ B where A.ҳ��=B.ҳ����� and B.ҳ�����=2"
        
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If gRs.BOF = False Then
            '��ʾ���ò�ѯ
            Call FrameDefault.ShowSpecPage(2)
            Select Case KeyCode
            Case vbKey0, vbKeyNumpad0
                Call FrameDefault.FirstChar("0")
            Case vbKey1, vbKeyNumpad1
                Call FrameDefault.FirstChar("1")
            Case vbKey2, vbKeyNumpad2
                Call FrameDefault.FirstChar("2")
            Case vbKey3, vbKeyNumpad3
                Call FrameDefault.FirstChar("3")
            Case vbKey4, vbKeyNumpad4
                Call FrameDefault.FirstChar("4")
            Case vbKey5, vbKeyNumpad5
                Call FrameDefault.FirstChar("5")
            Case vbKey6, vbKeyNumpad6
                Call FrameDefault.FirstChar("6")
            Case vbKey7, vbKeyNumpad7
                Call FrameDefault.FirstChar("7")
            Case vbKey8, vbKeyNumpad8
                Call FrameDefault.FirstChar("8")
            Case vbKey9, vbKeyNumpad9
                Call FrameDefault.FirstChar("9")
            End Select
        End If
        
    End Select
End Sub

Private Sub Form_Load()
    On Error GoTo errHand
    mvarBlnFirst = True
    
    '1.���������ϵ�ͼƬ�Ƿ��Ѿ����£�����Ѿ����£�����±���ͼƬ
    Call CheckPicture
    
'    Me.Width = 12000
'    Me.Height = 9000
    '2.��ȡ������ҳ���ʱ����
    mvarHomeLong = Val(GetPara("������ҳ���", "0"))
    Exit Sub
    
    
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call ResizeControl(FrameDefault, 0, 0, Me.ScaleWidth, Me.ScaleHeight)
End Sub

Private Sub LoadPageItemList(ByVal PageNo As Long)
'����:����ҳ���ÿһ��ѯ��Ŀ
'����:PageNo            ҳ�����
'˵��:���ǲ�ѯ������ʾ�����岿��,��ʾ��ѯ����
    Dim FileName As String
    Dim W As Single
    Dim H As Single
    Dim vFont As New StdFont
    Dim i As Long
    Dim j As Long
    Dim strTmp As String
    Dim rs As New ADODB.Recordset
    Dim vRs As New ADODB.Recordset
    Dim vNextY As Single
    Dim vNextX As Single
    Dim objDraw As ctlQueryItem
    Dim vWidth As Single
    Dim vHeight As Single
    Dim vTmp As Single
    Dim vTmp1 As Single
    Dim vMaxWidth As Single
    Dim vVisible As Boolean
    Dim strText As String
    
    On Error GoTo errHand
    i = 1
    vNextY = 60 + (i - 1) * 600
    vNextX = 120
    vMaxWidth = 120
            
    ShowFlatFlash "���Ժ���������ҳ��...", Me
    DoEvents
    
    Set objDraw = FrameDefault.ClientObj
    objDraw.ClientVisible = False
    Call objDraw.ClearAllPageItem
    
    '��ȡҳ��ı������������
    gstrSQL = "select B.����,B.���� from ��ѯҳ��Ŀ¼ A,��ѯͼƬԪ�� B where A.��������=B.��� and A.ҳ�����=[1]"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo)
    If gRs.BOF = False Then FrameDefault.AdviceMovie = IIf(IsNull(gRs!����), "", App.Path & "\ͼ��\" & gRs!���� & IIf(gRs!���� <> 2, ".pic", ".swf"))
                    
    '��ʼ�����Զ����ѯҳ��
    gstrSQL = "select ҳ�����,�������,��������,�����ı�,����ͼ��,��������,����λ��,��������,����ҳ��,��������,������,���λ��,��ͼ���,��ͼλ�� from ��ѯ����Ŀ¼ where ҳ�����=[1] order by �������"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo)
    If gRs.BOF = False Then
        While Not gRs.EOF
            strTmp = IIf(IsNull(gRs!��������), "����;12;0;0;0", gRs!��������)
            vFont.Name = Split(strTmp, ";")(0)
            vFont.Size = Val(Split(strTmp, ";")(1))
            vFont.Bold = Val(Split(strTmp, ";")(2))
            vFont.Italic = Val(Split(strTmp, ";")(3))
                                    
            FileName = ""
            '1.���ر������ݼ�����ͼ��
            vVisible = IIf(IsNull(gRs!��������), 1, gRs!��������)
            
            gstrSQL = "select ���� from ��ѯͼƬԪ�� where ���=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(IIf(IsNull(gRs!����ͼ��), 0, gRs!����ͼ��)))
            If rs.BOF = False Then
                FileName = GetFileName(IIf(IsNull(gRs!����ͼ��), 0, gRs!����ͼ��), W, H)
            End If
            Call objDraw.AddPageItemTitle(i, vNextY, IIf(IsNull(gRs!�����ı�), "", gRs!�����ı�), Val(Split(strTmp, ";")(4)), vFont, FileName, PageNo, IIf(IsNull(gRs!�������), 0, gRs!�������), vWidth, vHeight, Not vVisible, IIf(IsNull(gRs!����λ��), 0, gRs!����λ��))
                                                                                    
            If Not vVisible = True Then vNextY = vNextY + vHeight + 150

            Select Case zlCommFun.Nvl(gRs("��������").Value, 0)
            '----------------------------------------------------------------------------------------------------------
            Case 0      '���ı�����
                strTmp = IIf(IsNull(gRs!��������), "����;12;0;0;0", gRs!��������)
                j = objDraw.NextTxtIndex
                
                vWidth = FrameDefault.ClientWidth - 330
                
                'strText = gRs!�����ı�
                strText = Sys.ReadLob(glngSys, 29, PageNo & "," & gRs("�������").Value, "", 1)
                                
                Call objDraw.AddPageItemTxt(j, vNextX, vNextY, strText & Chr(13) & Chr(10), strTmp, vWidth, vHeight)
                vNextY = vNextY + vHeight + 150
            '----------------------------------------------------------------------------------------------------------
            Case 1      '���������
                vHeight = 0
                Call InsertGrid(objDraw, IIf(IsNull(gRs!������), 0, gRs!������), vNextX, vNextY, vWidth, vHeight)
                If vHeight > 0 Then vNextY = vNextY + vHeight + 150
            '----------------------------------------------------------------------------------------------------------
            Case 2      '��ͼ������
                FileName = GetFileName(IIf(IsNull(gRs!��ͼ���), 0, gRs!��ͼ���), W, H)
                Call objDraw.AddPageItemPic(objDraw.NextPicIndex, 0, vNextY, FileName, vWidth, vHeight, W, H)
                vNextY = vNextY + vHeight + 150
            '----------------------------------------------------------------------------------------------------------
            Case 3      '����������
                gstrSQL = "select A.����ҳ��,A.ҳ�ڶκ� from ��ѯ�������� A Where A.ҳ����� =[1] And A.������� = [2]"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo, Val(IIf(IsNull(gRs!�������), 0, gRs!�������)))
                If rs.BOF = False Then
                    While Not rs.EOF
                        If IIf(IsNull(rs!ҳ�ڶκ�), 0, rs!ҳ�ڶκ�) = 0 Then
                            'ֻ���ӵ�ҳ�棬û��ָ��ҳ���ڵľ�����Ŀ
                            gstrSQL = "select C.ҳ������ as �����ı� from ��ѯҳ��Ŀ¼ C Where C.ҳ�����=[1]"
                            Set vRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(IIf(IsNull(rs!����ҳ��), 0, rs!����ҳ��)))
                            If vRs.BOF = False Then
                                Call objDraw.AddPageItemConnect(objDraw.NextConnectIndex, vNextX + 150, vNextY, IIf(IsNull(vRs!�����ı�), "", vRs!�����ı�), IIf(IsNull(rs!����ҳ��), 0, rs!����ҳ��), 0, vWidth, vHeight)
                                vNextY = vNextY + 300
                            End If
                        Else
                            '���ӵ�ҳ���ڵľ�����Ŀ
                            gstrSQL = "select C.ҳ������||decode(B.�����ı�,NULL,'','��'||B.�����ı�) as �����ı� from ��ѯ����Ŀ¼ B,��ѯҳ��Ŀ¼ C Where C.ҳ������<>'ר�ҽ���' and B.ҳ�����=C.ҳ����� and C.ҳ�����=[1] and B.�������=[2]"
                            Set vRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(IIf(IsNull(rs!����ҳ��), 0, rs!����ҳ��)), Val(IIf(IsNull(rs!ҳ�ڶκ�), 0, rs!ҳ�ڶκ�)))
                            If vRs.BOF = False Then
                                Call objDraw.AddPageItemConnect(objDraw.NextConnectIndex, vNextX + 150, vNextY, IIf(IsNull(vRs!�����ı�), "", vRs!�����ı�), IIf(IsNull(rs!����ҳ��), 0, rs!����ҳ��), 0, vWidth, vHeight)
                                vNextY = vNextY + 300
                            Else
                                gstrSQL = "select B.����||'('||C.����||')' as ���� from ��Ա�� B,������Ա A,���ű� C Where B.id=A.��Աid And (b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.����ʱ�� Is Null) and A.����id=C.id and A.ȱʡ=1 and B.id=[1]"
                                Set vRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(IIf(IsNull(rs!ҳ�ڶκ�), 0, rs!ҳ�ڶκ�)))
                                If vRs.BOF = False Then
                                    Call objDraw.AddPageItemConnect(objDraw.NextConnectIndex, vNextX + 150, vNextY, IIf(IsNull(vRs!����), "", "ר�ҽ��ܣ�" & vRs!����), IIf(IsNull(rs!����ҳ��), 0, rs!����ҳ��), IIf(IsNull(rs!ҳ�ڶκ�), 0, rs!ҳ�ڶκ�), vWidth, vHeight)
                                    vNextY = vNextY + 300
                                End If
                            End If
                        End If
                        rs.MoveNext
                    Wend
                    vNextY = vNextY + 150
                End If
            '----------------------------------------------------------------------------------------------------------
            Case 4      '�ı��ͱ��
                strTmp = IIf(IsNull(gRs!��������), "����;12;0;0;0", gRs!��������)
                j = objDraw.NextTxtIndex
                
                strText = Sys.ReadLob(glngSys, 29, PageNo & "," & gRs("�������").Value, "", 1)
                
                Select Case IIf(IsNull(gRs!���λ��), 0, gRs!���λ��)
                Case 0
                    vHeight = 0
                    Call InsertGrid(objDraw, IIf(IsNull(gRs!������), 0, gRs!������), 0, vNextY, vTmp1, vTmp)
                    vWidth = FrameDefault.ClientWidth - vTmp1 - 120 - 120
                    Call objDraw.AddPageItemTxt(j, vNextX + vTmp1 + 60, vNextY, strText & Chr(13) & Chr(10), strTmp, vWidth, vHeight)
                Case 1
                    Call objDraw.AddPageItemTxt(j, vNextX, vNextY, strText & Chr(13) & Chr(10), strTmp, vTmp1, vHeight)
                    Call InsertGrid(objDraw, IIf(IsNull(gRs!������), 0, gRs!������), 1, vNextY, vWidth, vTmp)
                End Select
                vNextY = vNextY + IIf(vTmp > vHeight, vTmp, vHeight) + 150
            '----------------------------------------------------------------------------------------------------------
            Case 5      '�ı���ͼ��
            
                FileName = GetFileName(IIf(IsNull(gRs!��ͼ���), 0, gRs!��ͼ���), W, H)
                
                strText = Sys.ReadLob(glngSys, 29, PageNo & "," & gRs("�������").Value, "", 1)
                
                strTmp = IIf(IsNull(gRs!��������), "����;12;0;0;0", gRs!��������)
                j = objDraw.NextTxtIndex
                Select Case IIf(IsNull(gRs!��ͼλ��), 0, gRs!��ͼλ��)
                Case 0
                    Call objDraw.AddPageItemPic(objDraw.NextPicIndex, 0, vNextY, FileName, vTmp1, vTmp, W, H)
                    vWidth = FrameDefault.ClientWidth - vTmp1 - 120 - 120
                    Call objDraw.AddPageItemTxt(j, vNextX + vTmp1 + 60, vNextY, strText & Chr(13) & Chr(10), strTmp, vWidth, vHeight)
                Case 1
                    Call objDraw.AddPageItemPic(objDraw.NextPicIndex, 1, vNextY, FileName, vWidth, vTmp, W, H)
                    vTmp1 = FrameDefault.ClientWidth - vWidth - 60 - 90
                    Call objDraw.AddPageItemTxt(j, vNextX, vNextY, strText & Chr(13) & Chr(10), strTmp, vTmp1, vHeight)
                End Select
                vNextY = vNextY + IIf(vTmp > vHeight, vTmp, vHeight) + 150
                
            End Select
                      
            '8.���÷���ҳ�ױ�־
            If IIf(IsNull(gRs!����ҳ��), 0, gRs!����ҳ��) = 1 Then
                vHeight = 0
                Call objDraw.AddReturnFlag(vNextX, vNextY, vHeight)
                If vHeight > 0 Then vNextY = vNextY + vHeight + 150
            End If
            
            i = i + 1
            gRs.MoveNext
        Wend
    End If
        
    Call objDraw.ResizePage(FrameDefault.ClientWidth, vNextY)
    Call FrameDefault.InitNavigator(FrameDefault.ClientWidth, vNextY)
    
    '��ȡ����������ҳ�汳��
    gstrSQL = "select B.����,B.����,B.���,B.�߶� from ��ѯҳ��Ŀ¼ A,��ѯͼƬԪ�� B where A.ҳ�汳��=B.��� and A.ҳ�����=[1]"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo)
    If gRs.BOF = False Then
        Call objDraw.BackPicture(IIf(IsNull(gRs!����), "", App.Path & "\ͼ��\" & gRs!���� & IIf(gRs!���� <> 2, ".pic", ".swf")), IIf(IsNull(gRs!���), 0, gRs!���) * Screen.TwipsPerPixelX, IIf(IsNull(gRs!�߶�), 0, gRs!�߶�) * Screen.TwipsPerPixelY)
    End If
    
    
'    '��ȡ���������ļ�
'    FrameDefault.MusicFile = ""
'
'    Set gRs = OpenRecord(gRs, "select B.����,B.���� from ��ѯҳ��Ŀ¼ A,��ѯͼƬԪ�� B where A.��������=B.��� and A.ҳ�����=" & PageNo, Me.Caption)
'    If gRs.BOF = False Then
'        If IsNull(gRs!����) = False Then FrameDefault.MusicFile = App.Path & "\ͼ��\" & gRs!���� & ".mid"
'    End If
                
    Call objDraw.InitLoad
    objDraw.ClientVisible = True
    
    StopFlatFlash
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub LoadDoctorMsg(ByVal PageNo As Long)
'����:����ר�ҽ���ҳ������
'����:PageNo            ҳ�����
'˵��:���ǹ̶����ݲ��ݣ�����ZLHIS9����ȡ����Ա������Ϣ
    Dim FileName As String
    Dim W As Single
    Dim H As Single
    
    Dim vFont As New StdFont
    Dim i As Long
    Dim j As Long
    Dim strTmp As String
    Dim rs As New ADODB.Recordset
    Dim vNextY As Single
    Dim vNextX As Single
    Dim objDraw As ctlQueryItem
    Dim vWidth As Single
    Dim vHeight As Single
    Dim vTmp As Single
    Dim vTmp1 As Single
    Dim vMaxWidth As Single
    
    On Error GoTo errHand
    i = 1
    vNextY = 60 + (i - 1) * 600
    vNextX = 120
    vMaxWidth = 120
    
    Set objDraw = FrameDefault.ClientObj
    Call objDraw.ClearAllPageItem
    
    gstrSQL = "select A.��Աid,B.����||'('||D.����||')' as ���� from ��ѯר���嵥 A,��Ա�� B,������Ա C,���ű� D where B.ID=C.��Աid And (b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.����ʱ�� Is Null) and C.����id=D.ID and C.ȱʡ=1 and A.��ԱID=B.ID order by A.���"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If gRs.BOF = False Then
        While Not gRs.EOF
            strTmp = "����;12;1;0;0"
            vFont.Name = Split(strTmp, ";")(0)
            vFont.Size = Val(Split(strTmp, ";")(1))
            vFont.Bold = Val(Split(strTmp, ";")(2))
            vFont.Italic = Val(Split(strTmp, ";")(3))
                                                            
            '1.���ر������ݼ�����ͼ��
            Call objDraw.AddPageItemTitle(i, vNextY, IIf(IsNull(gRs!����), "", gRs!����), Val(Split(strTmp, ";")(4)), vFont, "", PageNo, IIf(IsNull(gRs!��ԱID), 0, gRs!��ԱID), vWidth, vHeight, True, 0)
            vNextY = vNextY + vHeight + 150
            
            '2.��Ƭ�����ֻ������
            gstrSQL = "select A.����, A.���˼�� from ��Ա�� A where   (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) and A.ID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(gRs!��ԱID))
            strTmp = "����;12;0;0;0"
            
            FileName = ""
            vTmp = 0
            If rs.BOF = False Then
                If IsNull(rs!����) = False Then FileName = App.Path & "\ͼ��\" & rs!���� & ".pic"
                If Dir(FileName) <> "" And FileName <> "" Then
                    
                    '��������Ƭ��С��ʾ����еȱ�����С ��2940*0.6? ��2280*0.6?   3.33 ����
                    '��Ƭ����ι涨�߶ȺͿ�ȵ�?
                    
                    Call objDraw.AddPageItemPic(objDraw.NextPicIndex, 0, vNextY, FileName, vTmp1, vTmp, 1368, 1764)
                    
                End If
                                
                vWidth = FrameDefault.ClientWidth - vTmp1 - 120 - 120
                j = objDraw.NextTxtIndex
                Call objDraw.AddPageItemTxt(j, vNextX + vTmp1 + 60, vNextY, IIf(IsNull(rs!���˼��), "", rs!���˼��) & Chr(13) & Chr(10), strTmp, vWidth, vHeight)
                vNextY = vNextY + IIf(vTmp > vHeight, vTmp, vHeight) + 150
            End If
            
'            '8.���÷���ҳ�ױ�־
'
'            vHeight = 0
'            Call objDraw.AddReturnFlag(vNextX, vNextY, vHeight)
'            If vHeight > 0 Then vNextY = vNextY + vHeight + 150
'
            
            i = i + 1
            gRs.MoveNext
        Wend
    End If
    
    Call objDraw.ResizePage(FrameDefault.ClientWidth, vNextY)
    Call FrameDefault.InitNavigator(FrameDefault.ClientWidth, vNextY)
    
    gstrSQL = "select B.����,B.����,B.���,B.�߶� from ��ѯҳ��Ŀ¼ A,��ѯͼƬԪ�� B where A.ҳ�汳��=B.��� and A.ҳ�����=[1]"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo)
    If gRs.BOF = False Then
        Call objDraw.BackPicture(IIf(IsNull(gRs!����), "", App.Path & "\ͼ��\" & gRs!���� & IIf(gRs!���� <> 2, ".pic", ".swf")), IIf(IsNull(gRs!���), 0, gRs!���) * Screen.TwipsPerPixelX, IIf(IsNull(gRs!�߶�), 0, gRs!�߶�) * Screen.TwipsPerPixelY)
    End If
    
'    '��ȡ���������ļ�
'    FrameDefault.MusicFile = ""
'
''    Call MusicClose
'    Set gRs = OpenRecord(gRs, "select B.����,B.���� from ��ѯҳ��Ŀ¼ A,��ѯͼƬԪ�� B where A.��������=B.��� and A.ҳ�����=" & PageNo, Me.Caption)
'    If gRs.BOF = False Then
'        If IsNull(gRs!����) = False Then FrameDefault.MusicFile = App.Path & "\ͼ��\" & gRs!���� & ".mid"
'    End If
                
    
    Call objDraw.InitLoad
        
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call MusicClose
    Set grs�Һ����� = Nothing   '67045
End Sub

Private Sub FrameDefault_ExitNewQuery(blnCancel As Boolean)
    If GetPara("����ָ���˳���ѯ", "0") = "0" Then
        blnCancel = False
    Else
        blnCancel = True
        Unload Me
    End If
End Sub

Private Sub FrameDefault_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub FrameDefault_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'tmrHome.Enabled = False
    'tmrHome.Enabled = True
    'tmrHome.Interval = mvarlngHome
    
    mvarHomeInternal = 0
    mvarCheckConnectCounter = 0
    
End Sub

Private Sub FrameDefault_ShowPage(ByVal PageNo As Long, ByVal CusomFormat As String)
'����:��ʾ��ѯҳ��
'����:PageNo            ҳ���
'     CusomFormat       ��ʽ,������ʱֻ�����֣�һ��"ר�ҽ���";����"�Զ���"

    If CusomFormat = "" Then
        Call LoadPageItemList(PageNo)
    Else
        Call LoadDoctorMsg(PageNo)
    End If
    
End Sub

Private Sub tmrCheckConnect_Timer()
    mvarCheckConnectCounter = mvarCheckConnectCounter + 1
    If mvarCheckConnectCounter >= mvarCheckConnectInternal Then

        '������ݿ�����״̬
        If gcnOracle.State = adStateOpen Then gcnOracle.Close

        Dim strErr As String
        
        If gobjRegister Is Nothing Then
            Set gobjRegister = gobjLogin.Register
        End If
        Set gcnOracle = gobjRegister.ReGetConnection(0, strErr)
        InitCommon gcnOracle
    End If
End Sub

Private Sub tmrHome_Timer()
    mvarHomeInternal = mvarHomeInternal + 1
    If mvarHomeInternal < mvarHomeLong Then Exit Sub
    mvarHomeInternal = 0
    
    On Error Resume Next
    Unload frmHelp
    Unload frmCardPass
    Unload frmSelect
    Unload frmIdentify����
    On Error GoTo 0
    
    Call FrameDefault.ShowHome
End Sub

Public Sub RefreshParamer(ByVal lngHomeLong As Long, ByVal lngCheckConnect As Long)
    mvarHomeLong = lngHomeLong
    tmrHome.Enabled = IIf(mvarHomeLong = 0, False, True)
    mvarCheckConnectInternal = lngCheckConnect
    
    'zyk add 200410
    Call FrameDefault.showwww
End Sub
