VERSION 5.00
Begin VB.Form frmDockExpense 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "ҽ�����ѹ���"
   ClientHeight    =   5280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7785
   Icon            =   "frmDockExpense.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin zlPublicExpense.ctlDockExpense dkeExpense 
      Height          =   4860
      Left            =   240
      TabIndex        =   0
      Top             =   60
      Width           =   6660
      _ExtentX        =   12330
      _ExtentY        =   8625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmDockExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------
'����¼�
Public Event Activate() '���Ѽ���ʱ
Public Event RequestRefresh() 'Ҫ��������ˢ��
Event StatusTextUpdate(ByVal bytType As Byte, ByVal Text As String) 'Ҫ�����������״̬������
'bytType:1-����ִ��,2-����ȡ��
Public Event zlPopupMenu(lngҽ��ID As Long, lng���ͺ� As Long, strNO As String, int��¼���� As Integer, X As Single, Y As Single)
'------------------------------------------


Private mfrmParent As Object
Public Enum ҽԺҵ��
    support����Ԥ�� = 0
    
    support�����˷� = 1
    supportԤ���˸����ʻ� = 2
    support�����˸����ʻ� = 3
    
    support�շ��ʻ�ȫ�Է� = 4       '�����շѺ͹Һ��Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�ȫ�Էѣ�ָͳ�����Ϊ0�Ľ��򳬳��޼۵Ĵ�λ�Ѳ���
    support�շ��ʻ������Ը� = 5     '�����շѺ͹Һ��Ƿ��ø����ʻ�֧�������Ը����֡������Ը�����1-ͳ�������* ���
    
    support�����ʻ�ȫ�Է� = 6       'סԺ���������������Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�
    support�����ʻ������Ը� = 7     'סԺ���������������Ƿ��ø����ʻ�֧�������Ը����֡�
    support�����ʻ����� = 8         'סԺ���������������Ƿ��ø����ʻ�֧�����޲��֡�
    
    support����ʹ�ø����ʻ� = 9     '����ʱ��ʹ�ø����ʻ�֧��
    supportδ�����Ժ = 10          '�����˻���δ�����ʱ��Ժ
    
    support���ﲿ�����ֽ� = 11      'ֻ��������ҽ����֧���˷Ѳ�ʹ�ñ�������Ҳ����˵�����ֽ�ʱ�ſ��ǲ�������񣬶��˻ص������ʻ���ҽ�������������˷ѡ�
    support��������ҽ����Ŀ = 12  '�ڽ���ʱ�����Ը��շ�ϸĿ�Ƿ�����ҽ����Ŀ���м��
    
    support������봫����ϸ = 13    '�����շѺ͹Һ��Ƿ���봫����ϸ
    
    support�����ϴ� = 14            'סԺ���ʷ�����ϸʵʱ����
    support���������ϴ� = 15        'סԺ�����˷�ʵʱ����

    support��Ժ���˽������� = 16    '�����Ժ���˽�������
    support������Ժ = 17            '���������˳�Ժ
    support����¼�������� = 18    '������Ժ���Ժʱ������¼�������
    support������ɺ��ϴ� = 19      'Ҫ���ϴ��ڼ��������ύ���ٽ���
    support��Ժ��������Ժ = 20    '���˽���ʱ���ѡ���Ժ���ʣ��ͼ������Ժ�ſ��Խ���
    
    support�Һ�ʹ�ø����ʻ� = 21    'ʹ��ҽ���Һ�ʱ�Ƿ�ʹ�ø����ʻ�����֧��

    support���������շ� = 22        '�����������֤�󣬿ɽ��ж���շѲ���
    support�����շ���ɺ���֤ = 23  '�������շ���ɣ��Ƿ��ٴε��������֤
    
    supportҽ���ϴ� = 24            'ҽ����������ʱ�Ƿ�ʵʱ����
    support�ֱҴ��� = 25            'ҽ�������Ƿ���ֱ�
    support��;������������ϴ����� = 26 '�ṩ�����ϴ��������ݵĽ��㹦��
    support��������ѽ��ʵļ��ʵ��� = 27 '�Ƿ�����������ʵ��ݣ�����õ����Ѿ�����
    
    support�����ݳ������� = 28
    support��Ժ��ʵ�ʽ��� = 29 '��Ժ�ӿ����Ƿ�Ҫ��ӿ��̽��н���
    support�������� = 35            '�Ƿ����������ʣ�����Ա����Ҫӵ�и������ʵ�Ȩ�ޡ��˲���ȱʡΪ�棬��֧�ֵĽӿ��赥������
    supportҽ���ӿڴ�ӡƱ�� = 46    'HIS��ֻ��Ʊ�ݺŵ�������ӡ��ҽ���ӿ�(����)�д�ӡ
    supportҽ��ȷ���������� = 48
    supportסԺ���˲�����׼��Ŀ���� = 50            'ͬһ�ֲ�,��סԺʱ����¼�����е���Ŀ
    support���ﲡ�˲�����׼��Ŀ���� = 51            '����������ĳ������¿���¼��������Ŀ
    supportʵʱ��� = 60
    
    support�ϴ����ﵵ�� = 70                    '������ҽ������ʱ���Ƿ����TranElecDossier����������ﲡ�˵��Ӿ���/���ӵ������ϴ�
    support�ҺŲ���ȡ������ = 81    '�ڹҺ�ʱ����ʹ��ҽ����ȡ������
    support�Һż����Ŀ = 86
End Enum
Private mbytFontSize As Byte
Private mlngOptModule As Long '����ģ���
Private mlngPlugInID As Long '�Զ�ִ�еĲ������ID
Private mrsPlugInBar As ADODB.Recordset '�˵���ʽ
Private mstrPreAdviceIdAndPayNums As String
Private mobjSaveData As Object
Private mblnFirst As Boolean
Private mstrMainPrivs As String
Private mcbsMain As Object
Private mobjSquareCard As Object

Public Sub SetFontSize(ByVal bytSize As Byte)
      '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:���˺�
    '����:2012-06-18 16:50:35
    '����:50793
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    '����vsFlexGrid�ؼ���ʹ�ø��Ի�����ʱ��Ӵ��п�����ڴ�����μ����ǲ���������,��Ҫ��getForm��������
    
    dkeExpense.FontSize = mbytFontSize
End Sub
 
Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal cbsMain As Object, ByRef objSquareCard As Object)
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    
    Set mfrmParent = frmParent
    If cbsMain Is Nothing Then Exit Sub
    
    If Not mblnFirst Then
        mblnFirst = True
        If objSquareCard Is Nothing Then
            Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
            If mobjSquareCard.zlInitComponents(Me, pҽ�����ѹ���, glngSys, gstrDBUser, gcnOracle, False) = False Then
                Set mobjSquareCard = Nothing
                MsgBox "ҽ�ƿ�������zl9CardSquare����ʼ��ʧ��!", vbInformation, gstrSysName
            End If
        Else
            Set mobjSquareCard = objSquareCard
        End If
        Set mcbsMain = cbsMain
        Set cbsMain.Icons = gobjCommFun.GetPubIcons
        Call GetPlugInBar(pҽ�����ѹ���, -1, mrsPlugInBar)
    End If
    
    'ҽ���˵�:���ڹ���˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    Else
        Call DefCommandPlugInPopup(objMenu.CommandBar.Controls, mrsPlugInBar)
    End If
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "����(&M)", objMenu.Index + 1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Append, "����������(&N)") '���Ƽ�ʱ��ʾΪ:����������,�ֹ��Ƽ�ʱ��ʾΪ:����������
        Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_NewItem, "���丽�ӷ���(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸ķ���(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ������(&D)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ExtraFeeMove, "����ת��(&T)")
        objControl.IconId = conMenu_Edit_CollectMan
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ExtraFeeExe, "����ִ��(&E)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ExtraFeeUnExe, "����ȡ��ִ��(&F)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ChargeDelApply, "��������(&L)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ChargeDelAudit, "�������(&U)")
        '��Ҳ˵�
        Call DefCommandPlugInPopup(objMenu.CommandBar.Controls, mrsPlugInBar)
    End With
    
    '���߲˵�:���������û��,���ڰ����˵�ǰ��
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", objMenu.Index, False)
        objMenu.ID = conMenu_ToolPopup
    End If
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Option, "ҽ������ѡ��(&O)"): objControl.BeginGroup = True
        objControl.IconId = conMenu_File_Parameter
    End With

    '����������:���ļ�������˵������ť֮��ʼ����
    '-----------------------------------------------------
    Set objBar = cbsMain(2)
    For Each objControl In objBar.Controls '�����ǰ������һ��Control
        If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
            Set objControl = objBar.Controls(objControl.Index - 1): Exit For
        End If
    Next
    With objBar.Controls
        'Set objControl = .Find(, conMenu_File_Preview) '��Ԥ����ť֮��ʼ����
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Append, "��������", objControl.Index + 1): objControl.BeginGroup = True
        Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_NewItem, "������", objControl.Index + 1): objPopup.BeginGroup = True
            objPopup.ID = conMenu_Edit_NewItem: objPopup.IconId = conMenu_Edit_NewItem
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�ķ�", objPopup.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��", objControl.Index + 1)
                
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ExtraFeeExe, "����ִ��", objControl.Index + 1): objControl.BeginGroup = True
        objControl.IconId = conMenu_Edit_Leave_Modify
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ExtraFeeUnExe, "ȡ��ִ��", objControl.Index + 1)
        objControl.IconId = conMenu_Edit_Leave_Delete
    End With
    
    '����Ŀ����
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyE, conMenu_Edit_Append '����������
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '�޸ĸ��ӷ���
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_Delete 'ɾ�����ӷ���
    End With

    '���ò���������
    '-----------------------------------------------------
    With cbsMain.Options
    End With
    
    
    '��ҳ�������ʼ��
    Call DefCommandPlugIn(cbsMain, mrsPlugInBar)
End Sub

Private Sub DefCommandPlugIn(ByRef cbsMain As Object, ByRef rsBar As ADODB.Recordset)
'���ܣ���Ҳ����˵����롣
'˵�����жϹؼ���     InTool �����˵���ʽ
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim i As Long
    Dim lngTmp As Long
    
    If rsBar Is Nothing Then Exit Sub
    rsBar.Filter = 0
    If rsBar.RecordCount = 0 Then Exit Sub
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    '������ť
    rsBar.Filter = "IsInTool=1  and BarType=1"
    If Not rsBar.EOF Then
        rsBar.Sort = "���"
        If Not objMenu Is Nothing Then
            With objMenu.CommandBar.Controls
                For i = 1 To rsBar.RecordCount
                    Set objControl = .Add(xtpControlButton, rsBar!����ID, rsBar!�˵���)
                        objControl.IconId = rsBar!ͼ��ID
                        objControl.Parameter = rsBar!������
                        objControl.Style = xtpButtonIconAndCaption
                    If Val(rsBar!IsGroup) = 1 Then
                        objControl.BeginGroup = True
                    End If
                    rsBar.MoveNext
                Next
            End With
        End If
    End If
    
    '������ť�����ֻ��һ����ť��Ҳ����������ť
    rsBar.Filter = "IsInTool=0 and BarType=1"
    If Not rsBar.EOF Then
        rsBar.Sort = "���"
        If Not objMenu Is Nothing Then
            Set objPopup = objMenu.CommandBar.Controls.Add(xtpControlButtonPopup, conMenu_Tool_PlugIn, "��չ����", , False)
                objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                For i = 1 To rsBar.RecordCount
                    Set objControl = .Add(xtpControlButton, rsBar!����ID, rsBar!�˵���)
                    objControl.IconId = rsBar!ͼ��ID
                    objControl.Parameter = rsBar!������
                    objControl.Style = xtpButtonIconAndCaption
                    If Val(rsBar!IsGroup) = 1 Then
                        objControl.BeginGroup = True
                    End If
                    rsBar.MoveNext
                Next
            End With
        End If
    End If
    
    '��������ť
    Set objBar = cbsMain(2)
    Set objControl = objBar.FindControl(, conMenu_Help_Help)
    If Not objControl Is Nothing Then
        objControl.BeginGroup = True
        lngTmp = objControl.Index - 1
    Else
        lngTmp = -1
    End If
    rsBar.Filter = "IsInTool=1 and BarType=2"
    If Not rsBar.EOF Then
        With objBar.Controls
            For i = 1 To rsBar.RecordCount
                Set objControl = .Add(xtpControlButton, rsBar!����ID, rsBar!������, lngTmp + 1)
                    objControl.IconId = rsBar!ͼ��ID
                    objControl.Parameter = rsBar!������
                    objControl.Style = xtpButtonIconAndCaption
                lngTmp = objControl.Index
                If Val(rsBar!IsGroup) = 1 Then objControl.BeginGroup = True
                rsBar.MoveNext
            Next
            objControl.BeginGroup = True
        End With
    End If
    rsBar.Filter = "IsInTool=0 and BarType=2"
    If Not rsBar.EOF Then
        rsBar.Sort = "���"
        Set objPopup = objBar.Controls.Add(xtpControlPopup, conMenu_Tool_PlugIn, "��չ����", lngTmp + 1, False)
            objPopup.ID = conMenu_Tool_PlugIn
            objPopup.IconId = conMenu_Tool_PlugIn
            objPopup.BeginGroup = True
            objPopup.Style = xtpButtonIconAndCaption
        lngTmp = objPopup.Index
        With objPopup.CommandBar.Controls
            For i = 1 To rsBar.RecordCount
                Set objControl = .Add(xtpControlButton, rsBar!����ID, rsBar!�˵���, lngTmp + 1)
                objControl.IconId = rsBar!ͼ��ID
                objControl.Parameter = rsBar!������
                If Val(rsBar!IsGroup) = 1 Then objControl.BeginGroup = True
                lngTmp = objPopup.Index
                rsBar.MoveNext
            Next
        End With
    End If
    '�Զ�ִ�еĹ���
    rsBar.Filter = "IsAuto=1"
    If Not rsBar.EOF Then mlngPlugInID = rsBar!����ID
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim strNO As String
    
    Select Case Control.ID
    Case conMenu_File_PrintSet '��ӡ����
        Call zlPrintSet
    Case conMenu_File_Preview 'Ԥ�������嵥
        Call dkeExpense.zlPrintData(2)
    Case conMenu_File_Print '��ӡ�����嵥
        Call dkeExpense.zlPrintData(1)
    Case conMenu_File_Excel '��������嵥
        Call dkeExpense.zlPrintData(3)
        
    Case conMenu_Help_Help '����
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Edit_Append '����������
        Call dkeExpense.zlBuildMainExpense(Me)
    Case conMenu_Edit_NewItem * 10# + 1 '����:�շѵ���
         Call dkeExpense.zlFuncFeeNewPrice(Me)
    Case conMenu_Edit_NewItem * 10# + 2 '����:���ʵ���
        Call dkeExpense.zlFuncFeeNewBilling(Me)
    Case conMenu_Edit_NewItem * 10# + 3 '����:��ķ���
        Call dkeExpense.zlFuncFeeNewNull(Me)
    Case conMenu_Edit_NewItem * 10# + 4  '����:�������Ļ����շ�
       Call dkeExpense.zlFuncStuffCharge(Me, 1)
    Case conMenu_Edit_NewItem * 10# + 5 ''�������ļ���
       Call dkeExpense.zlFuncStuffCharge(Me, 2)
    Case conMenu_Edit_Modify '�޸ĸ���
        Call dkeExpense.zlFuncFeeModi(Me)
    Case conMenu_Edit_Delete 'ɾ������
        Call dkeExpense.zlFuncFeeDel(Me)
    Case conMenu_Edit_ExtraFeeMove '����ת��
        Call dkeExpense.zlFuncExtraFeeMove(Me)
    Case conMenu_Edit_ExtraFeeExe   '����ִ��
        Call dkeExpense.zlFuncExtraFeeExe(Me, 1, mstrMainPrivs)
    Case conMenu_Edit_ExtraFeeUnExe '����ȡ��ִ��
        Call dkeExpense.zlFuncExtraFeeExe(Me, 0, mstrMainPrivs)
    Case conMenu_Edit_ChargeDelApply
        Call dkeExpense.zlFuncAdviceReCharge(1, Me)
    Case conMenu_Edit_ChargeDelAudit   '�����������
        Call dkeExpense.zlFuncAdviceReCharge(2, Me)
    Case conMenu_Tool_Option 'ҽ������ѡ��
        If frmExpenseSetup.zlEditCard(mfrmParent) Then
            'ˢ�·�����ϸ
            dkeExpense.Refresh
        End If
    Case conMenu_Tool_PlugIn_Item + 1 To conMenu_Tool_PlugIn_Item + 99 '��ҹ���ִ��
        Call dkeExpense.zlFuncPlugIn(Me, Control)
    Case Else
    End Select
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    Call dkeExpense.zlUpdateCommandBars(mcbsMain, Control)
End Sub
Public Sub zlPopupCommandBars(ByVal CommandBar As CommandBar)
    Dim objControl As CommandBarControl

    If CommandBar.Parent Is Nothing Then Exit Sub
    
    Select Case CommandBar.Parent.ID
    Case conMenu_Edit_NewItem '����
        With CommandBar.Controls
            .DeleteAll
            '��1λ,Ϊ��ʹ�ÿ�ݼ�
            Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem * 10# + 1, "�շѵ���(&1)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem * 10# + 2, "���ʵ���(&2)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem * 10# + 3, "��ķ���(&3)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem * 10# + 4, "���������շ�(&3)"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem * 10# + 5, "�������ļ���(&4)")
            With mcbsMain.KeyBindings
                .Add FCONTROL, vbKeyN, conMenu_Edit_NewItem * 10# + 1
                .Add FCONTROL, vbKeyB, conMenu_Edit_NewItem * 10# + 2
                .Add FCONTROL, vbKeyS, conMenu_Edit_NewItem * 10# + 4
            End With
        End With
    End Select
End Sub
Private Sub dkeExpense_Activate()
    RaiseEvent Activate
End Sub
Private Sub dkeExpense_RequestRefresh()
    RaiseEvent RequestRefresh
End Sub
 
Private Sub dkeExpense_StatusTextUpdate(ByVal bytType As Byte, ByVal Text As String)
    RaiseEvent StatusTextUpdate(bytType, Text)
End Sub

Private Sub dkeExpense_zlPopupMenu(lngҽ��ID As Long, lng���ͺ� As Long, strNO As String, int��¼���� As Integer, X As Single, Y As Single)
    RaiseEvent zlPopupMenu(lngҽ��ID, lng���ͺ�, strNO, int��¼����, X, Y)
End Sub

Private Sub Form_Load()
    mbytFontSize = 9
    mblnFirst = False
    Set mrsPlugInBar = Nothing
    Call dkeExpense.zlInitCommon(mobjSquareCard)
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With dkeExpense
        .Left = ScaleLeft
        .Top = ScaleTop
        .Width = ScaleWidth
        .Height = ScaleHeight
    End With
End Sub

Public Function zlRefresh(ByVal lng����id As Long, ByVal strAdviceIdAndPayNums As String, _
    Optional ByVal blnMoved As Boolean = False, Optional ByVal strNos As String, _
    Optional ByVal byt��¼���� As Byte, Optional ByVal byt������Դ As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������ˢ��
    '���:lng����id-����ID
    '     strAdviceIdAndPayNums-ҽ��ID�ͷ��ͺźͶ���ִ�б�־(ҽ��ID1:���ͺ�1:����ִ��,ҽ��ID2:���ͺ�2:����ִ��,...)
    '     strNos:���ݺ�(�������ʱ,�ö��ŷ���)
    '     byt��¼����:ҽ��ID����ʱ,�Ŵ���,��������(1-�շѵ�;2-���ʵ�)
    '     byt������Դ-1-����;2-סԺ
    '     blnMoved -�ò��˵������Ƿ���ת��
    '     bln����ִ��-���ڼ�����Ŀ��һ���ɼ���һ����Ŀ���Ƿ�������е�ĳһ������ִ��
    '����:
    '����:���˺�
    '����:2014-04-10 11:02:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objControl As CommandBarControl
    
    zlRefresh = dkeExpense.zlRefresh(Me, lng����id, strAdviceIdAndPayNums, blnMoved, strNos, byt��¼����, byt������Դ)
    
    If mstrPreAdviceIdAndPayNums <> strAdviceIdAndPayNums And mlngPlugInID <> 0 Then
        Set objControl = mcbsMain.FindControl(, mlngPlugInID, , True)
        If Not objControl Is Nothing Then objControl.Execute
        mstrPreAdviceIdAndPayNums = strAdviceIdAndPayNums
    End If
End Function

Public Function zlBuildMainExpense() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������
    '����:���ɳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 13:37:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlBuildMainExpense = dkeExpense.zlBuildMainExpense
End Function

Public Function zlAddChargeExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal bytӦ�ó��� As Byte, _
    Optional ByVal lng����ID As Long, _
    Optional ByVal lng��������id As Long, Optional ByVal lng���˿���ID As Long, _
    Optional ByRef strOutNos As String, _
    Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���շѷ���
    '���:frmMain-����������
    '     lngModule-ģ���
    '     bytӦ�ó���:0-ҽ������;1-��첹��(��ѡ����)
    '����:strOutNos-�ɹ�����ĵ��ݺ�
    '����:���շѷ���,����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 13:37:47
    '˵��:
    '    1.ҽ������ִ�о���Ĺ���(�μ�:zlCisKernel.dockExpense)
    '       ����Ҫ���벡��ID; �������Ҽ����˿���ID
    '    2.��첹��ʱ,��Ҫ����lng����ID,��������ID,���˿���ID
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If bytӦ�ó��� = 0 Then
        zlAddChargeExpense = dkeExpense.zlFuncFeeNewPrice(frmMain, strOutNos, objSaveData)
        Exit Function
    End If
    If frmTechnicExpense.EditCard(frmMain, GetInsidePrivs(pҽ�����ѹ���), 0, 0, 0, lng����ID, 0, _
         1, 1, lng��������id, lng���˿���ID, 0, "", "", "", "", , , False, strOutNos, bytӦ�ó���, objSaveData, mobjSquareCard) Then
         zlAddChargeExpense = True
    End If
End Function

Public Function zlAddBillingExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal bytӦ�ó��� As Byte, ByVal int������Դ As Integer, _
    Optional ByVal lng����ID As Long, Optional lng��ҳId As Long, _
    Optional ByVal lng��������id As Long, _
    Optional ByVal lng���˿���ID As Long, Optional ByRef strOutNos As String, _
    Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ʷ���
    '���:frmMain -����������
    '    lngModule -ģ���
    '    bytӦ�ó���:0-ҽ������;1-��첹��(��ѡ����)
    '    int������Դ:1-���ﲡ��,2-סԺ����
    '����:strOutNos-�ɹ�����ĵ��ݺ�
    '����:�����ʷ���,���ѳɹ�����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 13:37:47
    '˵��:
    '    1.ҽ������ִ�о���Ĺ���(�μ�:zlCisKernel.dockExpense)
    '       ����Ҫ���벡��ID;�������Ҽ����˿���ID
    '    2.��첹��ʱ,��Ҫ����lng����ID,��������ID,���˿���ID
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If bytӦ�ó��� = 0 Then
        '99053:���ϴ�,2016/7/29�����ò����˷��ýӿ�
        zlAddBillingExpense = dkeExpense.zlFuncFeeNewBilling(frmMain, strOutNos)
        Exit Function
    End If
    zlAddBillingExpense = frmTechnicExpense.EditCard(frmMain, GetInsidePrivs(pҽ�����ѹ���), 0, 0, 0, lng����ID, lng��ҳId, _
          int������Դ, 2, lng��������id, lng���˿���ID, 0, "", "", "", "", , , False, strOutNos, bytӦ�ó���, objSaveData, mobjSquareCard)
End Function

Public Function zlAddZeroExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal bytӦ�ó��� As Byte, ByVal int������Դ As Integer, _
    Optional ByVal lng����ID As Long, Optional ByVal lng��ҳId As Long, _
    Optional ByVal lng��������id As Long, _
    Optional ByVal lng���˿���ID As Long, _
    Optional ByRef strOutNos As String, _
    Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '���:frmMain -����������
    '    lngModule -ģ���
    '    bytӦ�ó���:0-ҽ������;1-��첹��(��ѡ����)
    '    int������Դ:1-���ﲡ��,2-סԺ����
    '    lng��ҳID -ҽ�����Ѻ����ﲡ�˴���0,סԺ���˱���ת��
    '����:strOutNos-������ķѵ���,����ö��ŷ���
    '����:�����ʷ���,���ѳɹ�����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 13:37:47
    '˵��:
    '    1.ҽ������ִ�о���Ĺ���(�μ�:zlCisKernel.dockExpense)
    '       ����Ҫ���벡��ID;�������Ҽ����˿���ID
    '    2.��첹��ʱ,��Ҫ����lng����ID,��������ID,���˿���ID
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If bytӦ�ó��� = 0 Then
        zlAddZeroExpense = dkeExpense.zlFuncFeeNewNull(frmMain, strOutNos, objSaveData)
        Exit Function
    End If
    zlAddZeroExpense = frmTechnicExpense.EditCard(frmMain, GetInsidePrivs(pҽ�����ѹ���), 0, 0, 0, lng����ID, lng��ҳId, _
          int������Դ, 2, lng��������id, lng���˿���ID, 0, "", "", "", "", , , True, strOutNos, bytӦ�ó���, objSaveData, mobjSquareCard)
End Function

Public Function zlAddStuffChargeExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal bytӦ�ó��� As Byte, ByVal int������Դ As Integer, _
    Optional ByVal lng����ID As Long, _
    Optional ByVal lng��������id As Long, _
    Optional ByVal lng���˿���ID As Long, _
    Optional ByRef strOutNos As String, _
    Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������շѷ���
    '���:frmMain -����������
    '    lngModule -ģ���
    '    bytӦ�ó���:0-ҽ������;1-��첹��(��ѡ����)
    '    int������Դ:1-���ﲡ��,2-סԺ����
    '����:strNo-���ر������ĵ���,����ö��ŷ���
    '����:�����ʷ���,���ѳɹ�����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 13:37:47
    '˵��:
    '    1.ҽ������ִ�о���Ĺ���(�μ�:zlCisKernel.dockExpense)
    '       ����Ҫ���벡��ID;�������Ҽ����˿���ID
    '    2.��첹��ʱ,��Ҫ����lng����ID,��������ID,���˿���ID
    '---------------------------------------------------------------------------------------------------------------------------------------------
   If bytӦ�ó��� = 0 Then
        Set mobjSaveData = objSaveData
        zlAddStuffChargeExpense = dkeExpense.zlFuncStuffCharge(frmMain, 1, strOutNos, objSaveData)
        Exit Function
    End If
    zlAddStuffChargeExpense = frmStuffCharge.zlBillEdit(frmMain, 0, lngModule, GetInsidePrivs(pҽ�����ѹ���), 1, "", _
         1, lng����ID, 0, lng��������id, lng���˿���ID, _
          0, "", False, "", 0, 0, "", , , strOutNos, bytӦ�ó���, objSaveData, mobjSquareCard) = True
End Function

Public Function zlAddStuffBillingExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal bytӦ�ó��� As Byte, ByVal int������Դ As Integer, _
    Optional ByVal lng����ID As Long, _
    Optional ByVal lng��ҳId As Long, _
    Optional ByVal lng��������id As Long, _
    Optional ByVal lng���˿���ID As Long, _
    Optional ByRef strOutNos As String, _
    Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������ļ��ʷ���
    '���:frmMain -����������
    '    lngModule -ģ���
    '    bytӦ�ó���:0-ҽ������;1-��첹��(��ѡ����)
    '    int������Դ:1-���ﲡ��,2-סԺ����
    '    lng��ҳID-��ҳID���Բ���(��סԺ����һ��Ҫ����)
    '����:strNo-���ر������ĵ���,����ö��ŷ���
    '����:�����ʷ���,���ѳɹ�����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 13:37:47
    '˵��:
    '    1.ҽ������ִ�о���Ĺ���(�μ�:zlCisKernel.dockExpense)
    '       ����Ҫ���벡��ID;�������Ҽ����˿���ID
    '    2.��첹��ʱ,��Ҫ����lng����ID,��������ID,���˿���ID
    '---------------------------------------------------------------------------------------------------------------------------------------------
   If bytӦ�ó��� = 0 Then
        Set mobjSaveData = objSaveData
        zlAddStuffBillingExpense = dkeExpense.zlFuncStuffCharge(frmMain, 2, strOutNos, objSaveData)
        Exit Function
    End If
    zlAddStuffBillingExpense = frmStuffCharge.zlBillEdit(frmMain, 0, lngModule, GetInsidePrivs(pҽ�����ѹ���), 1, "", _
         2, lng����ID, 0, lng��������id, lng���˿���ID, _
          0, "", False, "", 0, 0, "", , , strOutNos, bytӦ�ó���, objSaveData, mobjSquareCard) = True
End Function
Public Function zlIsFunValied(ByVal bytType As Byte, ByVal bytPrivsCheck As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�м��ĳ�����Ƿ���Ч
    '���: bytType- 1-�޸ĸ���;2-ɾ������;3-����ת��;4-����ִ��;5-����ȡ��ִ��;6-��������;7-�������
    '      bytPrivsCheck -���Ȩ��:0-�����Ȩ��;1-������ݺ�Ȩ��;2-�����Ȩ��
    '����:
    '����:������Ч,����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 17:00:52
    '˵��:
    '   1.���ݸ����б��е�����,���ĳ����Ƿ���Ч
    '   2.����Ȩ�޼��ĳ����Ƿ���Ч
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlIsFunValied = dkeExpense.IsFunValied(bytType, bytPrivsCheck)
End Function

Public Function zlModifyExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal int������Դ As Integer, ByVal int��¼���� As Integer, ByVal strNO As String, Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�޸ĸ���
    '���:frmMain -����������
    '    lngModule -ģ���
    '    int������Դ-������Դ
    '    int��¼����
    '    strNO
    '����:�޸ĸ��ѳɹ�����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 13:37:47
    '˵��:
    '    1.ִ�о���Ĺ���(�μ�:zlCisKernel.dockExpense),����Ҫ�����¼���ʺ�NO
    '    2.������ʱ,��Ҫ�����¼���ʺ�NO
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlModifyExpense = dkeExpense.zlFuncFeeModi(frmMain, int������Դ, int��¼����, strNO, objSaveData)
End Function
Public Function zlDelExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long, Optional int������Դ As Integer, _
    Optional ByVal int��¼���� As Integer, _
    Optional ByVal strNO As String, Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ɾ������
    '���:frmMain -����������
    '    int������Դ-1-����;2-סԺ
    '    lngModule -ģ���
    '    int��¼����
    '    strNO
    '����:�޸ĸ��ѳɹ�����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 13:37:47
    '˵��:
    '    1.ִ�о���Ĺ���(�μ�:zlCisKernel.dockExpense),����Ҫ�����¼���ʺ�NO
    '    2.������ʱ,��Ҫ����,������Դ,��¼���ʺ�NO
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlDelExpense = dkeExpense.zlFuncFeeDel(frmMain, int������Դ, int��¼����, strNO, objSaveData)
End Function
Public Function zlMoveExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ƶ�
    '���:frmMain -����������
    '    lngModule -ģ���
    '����:�����ƶ��ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 13:37:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlMoveExpense = dkeExpense.zlFuncExtraFeeMove(frmMain)
End Function
Public Function zlExcuteExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long, Optional blnȡ��ִ�� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ִ��
    '���:frmMain -����������
    '    lngModule -ģ���
    '����:����ִ�гɹ�����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 13:37:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    'bytType=0-ȡ��ִ��,1-ִ��
    zlExcuteExpense = dkeExpense.zlFuncExtraFeeExe(frmMain, IIf(blnȡ��ִ��, 0, 1), mstrMainPrivs)
End Function


Public Function zlApplyExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '���:frmMain -����������
    '    lngModule -ģ���
    '����:��������ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 13:37:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlApplyExpense = dkeExpense.zlFuncAdviceReCharge(1, frmMain)
End Function
 
Public Function zlAuditExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '���:frmMain -����������
    '    lngModule -ģ���
    '����:����������˳ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 13:37:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlAuditExpense = dkeExpense.zlFuncAdviceReCharge(2, frmMain)
End Function
Public Function zlParaOptionSet(ByVal frmMain As Object, _
    ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '���:frmMain -����������
    '    lngModule -ģ���
    '����:����������˳ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2014-04-10 13:37:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
   zlParaOptionSet = frmExpenseSetup.zlEditCard(frmMain)
End Function

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    If Not mobjSaveData Is Nothing Then Set mobjSaveData = Nothing
End Sub

Private Sub DefCommandPlugInPopup(ByVal objBar As Object, ByRef rsBar As ADODB.Recordset)
'���ܣ���ҽ�����Ҽ������˵�
    Dim i As Long
    Dim objControl As CommandBarControl
    Dim objCtl As CommandBarControl
    Dim objPopup As CommandBarPopup
    
    If rsBar Is Nothing Then Exit Sub
    rsBar.Filter = 0
    If rsBar.RecordCount = 0 Then Exit Sub
    
    '������ť
    rsBar.Filter = "IsInTool=1 and BarType=3"
    If Not rsBar.EOF Then
        rsBar.Sort = "���"
        For i = 1 To rsBar.RecordCount
            Set objControl = objBar.Add(xtpControlButton, rsBar!����ID, rsBar!������)
            objControl.IconId = rsBar!ͼ��ID
            objControl.Parameter = rsBar!������
            objControl.Style = xtpButtonIconAndCaption
            If Val(rsBar!IsGroup) = 1 Then
                objControl.BeginGroup = True
            End If
            rsBar.MoveNext
        Next
    End If
    
    rsBar.Filter = "IsInTool=0 and BarType=3"
    If Not rsBar.EOF Then
        rsBar.Sort = "���"
        Set objPopup = objBar.Add(xtpControlButtonPopup, conMenu_Tool_PlugIn, "��չ����")
            objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            For i = 1 To rsBar.RecordCount
                Set objControl = .Add(xtpControlButton, rsBar!����ID, rsBar!�˵���)
                objControl.IconId = rsBar!ͼ��ID
                objControl.Parameter = rsBar!������
                objControl.Style = xtpButtonIconAndCaption
                If Val(rsBar!IsGroup) = 1 Then
                    objControl.BeginGroup = True
                End If
                rsBar.MoveNext
            Next
        End With
    End If
End Sub

Private Function GetPlugInBar(ByVal lngģ�� As Long, ByVal int���� As Integer, rsBar As ADODB.Recordset) As String
'���ܣ���֯��Ҳ����Ĳ˵�����ť
    Dim strFunc As String
    Dim strXML As String
    Call CreatePlugIn(lngģ��, int����)
    If gobjPlugIn Is Nothing Then Exit Function
    On Error Resume Next
    strFunc = gobjPlugIn.GetFuncNames(glngSys, lngģ��, int����, strXML)
    Call zlPlugInErrH(Err, "GetFuncNames")
    Err.Clear: On Error GoTo 0
    Call MakePlugInBar(strFunc, strXML, rsBar)
    GetPlugInBar = strFunc
End Function

Private Sub MakePlugInBar(ByVal strFunc As String, ByVal strXML As String, rsBar As ADODB.Recordset)
'���ܣ���֯�˵������ؼ�¼���У�ע����ϰ汾�ļ��ݴ���
'������strFunc �ϰ汾�����д���strXML��������Ϣ�Ĺ��ܴ�
    Dim strM As String
    Dim strB As String
    Dim strP As String
    Dim strTag As String
    Dim i As Long
    Dim strTmp As String
    Dim lngS As Long, lngE As Long
    Dim rsBarFuncID As ADODB.Recordset
    
    If strXML = "" And strFunc <> "" Then
        '������ǰ�ϰ汾�ķ�ʽ
        Call InitPlugInRsBar(rsBar)
        Call AddPlugInBarRs(rsBar, strFunc, 1)
        Call AddPlugInBarRs(rsBar, strFunc, 2)
        Call AddPlugInBarRs(rsBar, strFunc, 3)
        Call SetPlugInBar(rsBar, 1)
        Exit Sub
    End If
    
    On Error GoTo errH
    strXML = Trim(strXML)
    '�ݶ�Ϊ200����չ���ܲ������ֹ��ѭ��
    For i = 0 To 200
        lngS = InStr(strXML, "<")
        lngE = InStr(strXML, ">")
        strTag = Mid(strXML, lngS + 1, lngE - lngS - 1)
        If strTag = "menubar" Then
            lngS = lngE
            lngE = InStr(strXML, "</menubar>")
            strTmp = Mid(strXML, lngS + 1, lngE - lngS - 1)
            If strTmp <> "" Then strM = strM & "," & strTmp
            strXML = Mid(strXML, lngE + 10)
        ElseIf strTag = "toolbar" Then
            lngS = lngE
            lngE = InStr(strXML, "</toolbar>")
            strTmp = Mid(strXML, lngS + 1, lngE - lngS - 1)
            If strTmp <> "" Then strB = strB & "," & strTmp
            strXML = Mid(strXML, lngE + 10)
        ElseIf strTag = "popbar" Then
            lngS = lngE
            lngE = InStr(strXML, "</popbar>")
            strTmp = Mid(strXML, lngS + 1, lngE - lngS - 1)
            If strTmp <> "" Then strP = strP & "," & strTmp
            strXML = Mid(strXML, lngE + 9)
        End If
        If strXML = "" Then
            Exit For
        End If
    Next
    If strM = "" Then Exit Sub
    strM = Mid(strM, 2)
    strB = Mid(strB, 2)
    strP = Mid(strP, 2)

    Call InitPlugInRsBar(rsBar)
    Call AddPlugInBarRs(rsBar, strM, 1)
    Call AddPlugInBarRs(rsBar, strB, 2)
    Call AddPlugInBarRs(rsBar, strP, 3)
    Call SetPlugInBar(rsBar, 2)
    Exit Sub
errH:
    If 2 = 1 Then
        Resume
    End If
End Sub

Private Sub AddPlugInBarRs(ByRef rsBar As ADODB.Recordset, ByVal strFunc As String, ByVal intType As Integer)
'���ܣ������ܴ�ת��Ϊ��¼����ʽ
'������strFunc ���ܴ���intType ���ܰ�ť������һ�� 1-�˵�����2-��������3-�����
    Dim varFunc As Variant
    Dim i As Long
    Dim strFuncName As String
    Dim blnFirstTool As Boolean
    If strFunc = "" Then Exit Sub
    varFunc = Split(strFunc, ",")
    With rsBar
        For i = 0 To UBound(varFunc)
            strFuncName = varFunc(i)
            .AddNew
            !BarType = intType
            If InStr(strFuncName, "Auto:") > 0 Then
                !IsAuto = 1
                strFuncName = Replace(strFuncName, "Auto:", "")
            Else
                !IsAuto = 0
            End If
            
            If InStr(strFuncName, "InTool:") > 0 Then
                !IsInTool = 1
                strFuncName = Replace(strFuncName, "InTool:", "")
            Else
                !IsInTool = 0
            End If
            If InStr(strFuncName, "|:") > 0 Then
                !IsGroup = 1
                strFuncName = Replace(strFuncName, "|:", "")
            Else
                !IsGroup = 0
                If Not blnFirstTool And !IsInTool = 1 Then
                    '��һ��������ť��ʾ�ָ���
                    blnFirstTool = True
                    !IsGroup = 1
                End If
            End If
            !������ = strFuncName
            !�˵��� = strFuncName
            .Update
        Next
    End With
End Sub

Private Function SetPlugInBar(ByRef rsBar As ADODB.Recordset, ByVal lngV As Long) As String
'���ܣ����书��ID���Ӳ˵����
'������lngV �汾��1-�ϰ棬2-�°�
'���أ��ַ�������ǰ�Ͱ汾��ʽ�Ĺ��ܴ�
    Dim i As Long
    '���书��ID��ͼ��ID
    With rsBar
        .Filter = 0
        If .EOF Then Exit Function
        .MoveFirst
        For i = 1 To .RecordCount
            !��� = i
            !����ID = conMenu_Tool_PlugIn_Item + i
            !ͼ��ID = conMenu_Tool_PlugIn_Item
            If lngV = 1 Then
                !IsInTool = 0
                !IsGroup = 0
            End If
            .Update
            .MoveNext
        Next
    End With
    Call SetPlugInBarKey(rsBar, 1, lngV)
    Call SetPlugInBarKey(rsBar, 2, lngV)
    Call SetPlugInBarKey(rsBar, 3, lngV)
    rsBar.Filter = 0
End Function

Private Sub SetPlugInBarKey(rsBar As ADODB.Recordset, ByVal intType As Integer, ByVal lngV As Long)
'���ܣ��趨���
'������lngV �汾��1-�ϰ棬2-�°� intType ���ܰ�ť������һ�� 1-�˵�����2-��������3-�����
    Dim i As Long
    With rsBar
        .Filter = "IsInTool=0 and BarType=" & intType
        If .RecordCount = 1 And lngV = 2 Then
            '���ֻ��һ����Ҳ��Ϊ������ť
            !IsInTool = 1
            .Update
        Else
            For i = 1 To .RecordCount
                If i <= 35 Then
                    If i <= 9 Then
                        !�˵��� = !�˵��� & "(&" & i & ")"
                    Else
                        !�˵��� = !�˵��� & "(&" & Chr(55 + i) & ")"
                    End If
                    .Update
                    .MoveNext
                Else
                    Exit For
                End If
            Next
        End If
        
        .Filter = "IsInTool=1 and BarType=" & intType
        For i = 1 To .RecordCount
            If i <= 35 Then
                If i <= 9 Then
                    !�˵��� = !�˵��� & "(&" & i & ")"
                Else
                    !�˵��� = !�˵��� & "(&" & Chr(55 + i) & ")"
                End If
                .Update
                .MoveNext
            Else
                Exit For
            End If
        Next
    End With
End Sub

Private Sub InitPlugInRsBar(rsBar As ADODB.Recordset)
    Set rsBar = New ADODB.Recordset
    rsBar.Fields.Append "���", adBigInt '��������
    rsBar.Fields.Append "����ID", adBigInt '�˵���ť Control.ID
    rsBar.Fields.Append "ͼ��ID", adBigInt
    rsBar.Fields.Append "������", adVarChar, 1000 'ȥ���ؼ���֮��� ���� ���������ϵİ�ť����
    rsBar.Fields.Append "�˵���", adVarChar, 1000 '�˵���/�Ҽ��˵� ����
    rsBar.Fields.Append "IsAuto", adInteger '�Ƿ��Զ�ִ�й���
    rsBar.Fields.Append "IsGroup", adInteger '�Ƿ�ָ���
    rsBar.Fields.Append "IsInTool", adInteger '�Ƿ������ʾ
    rsBar.Fields.Append "BarType", adInteger '1-�˵�����2����������3��������
    rsBar.CursorLocation = adUseClient
    rsBar.LockType = adLockOptimistic
    rsBar.CursorType = adOpenStatic
    rsBar.Open
End Sub
