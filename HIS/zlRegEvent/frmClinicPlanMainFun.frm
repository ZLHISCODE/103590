VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmClinicPlanMainFun 
   BorderStyle     =   0  'None
   Caption         =   "���ܲ˵�"
   ClientHeight    =   7740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picPlanBack 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   4965
      Left            =   7590
      ScaleHeight     =   4965
      ScaleWidth      =   3345
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   690
      Width           =   3345
      Begin VB.Frame frmMoveSplitY 
         Height          =   25
         Left            =   -150
         MousePointer    =   7  'Size N S
         TabIndex        =   5
         Top             =   2670
         Width           =   3735
      End
      Begin MSComctlLib.TreeView tvwPlan 
         Height          =   1425
         Left            =   390
         TabIndex        =   6
         Top             =   3270
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   2514
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   88
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imgPlan16"
         Appearance      =   0
         OLEDragMode     =   1
      End
      Begin MSComctlLib.TreeView tvwPlanTemplet 
         Height          =   1815
         Left            =   240
         TabIndex        =   7
         Top             =   810
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   3201
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   88
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imgPlan16"
         Appearance      =   0
         OLEDragMode     =   1
      End
      Begin VB.Image imgYearSelect 
         Height          =   120
         Left            =   2250
         Picture         =   "frmClinicPlanMainFun.frx":0000
         Top             =   2880
         Width           =   120
      End
      Begin VB.Label lblYearSelect 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2016��"
         Height          =   180
         Left            =   1680
         TabIndex        =   10
         Top             =   2850
         Width           =   540
      End
      Begin XtremeSuiteControls.ShortcutCaption sccPlan 
         Height          =   360
         Left            =   0
         TabIndex        =   9
         Top             =   2760
         Width           =   3165
         _Version        =   589884
         _ExtentX        =   5583
         _ExtentY        =   635
         _StockProps     =   6
         Caption         =   "���ﰲ��"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin XtremeSuiteControls.ShortcutCaption sccPlanTemplet 
         Height          =   360
         Left            =   0
         TabIndex        =   8
         Top             =   390
         Width           =   2505
         _Version        =   589884
         _ExtentX        =   4419
         _ExtentY        =   635
         _StockProps     =   6
         Caption         =   "����ģ��"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox picBaseSetBack 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   4905
      Left            =   270
      ScaleHeight     =   4905
      ScaleWidth      =   2985
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   570
      Width           =   2985
      Begin XtremeSuiteControls.TaskPanel tplFunBase 
         Height          =   3045
         Left            =   540
         TabIndex        =   11
         Top             =   1110
         Width           =   1815
         _Version        =   589884
         _ExtentX        =   3201
         _ExtentY        =   5371
         _StockProps     =   64
         ItemLayout      =   2
         HotTrackStyle   =   1
      End
      Begin XtremeSuiteControls.ShortcutCaption sccFunBase 
         Height          =   360
         Left            =   0
         TabIndex        =   3
         Top             =   30
         Width           =   2505
         _Version        =   589884
         _ExtentX        =   4419
         _ExtentY        =   635
         _StockProps     =   6
         Caption         =   "��������"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox picFunBack 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   4965
      Left            =   3510
      ScaleHeight     =   4965
      ScaleWidth      =   3555
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   3555
      Begin XtremeSuiteControls.ShortcutBar scbFunc 
         Height          =   4155
         Left            =   60
         TabIndex        =   1
         Top             =   210
         Width           =   3225
         _Version        =   589884
         _ExtentX        =   5689
         _ExtentY        =   7329
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.ImageList imgPlan16 
      Left            =   9630
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":04AA
            Key             =   "RootPlan"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":07FC
            Key             =   "StopPlan"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":705E
            Key             =   "TempletPlan"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":75F8
            Key             =   "FixedPlan"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":7B92
            Key             =   "InvalidFixedPlan"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":812C
            Key             =   "InvalidPublishedFixedPlan"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":86C6
            Key             =   "PublishedFixedPlan"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":8C60
            Key             =   "InvalidMonthPlan"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":91FA
            Key             =   "InvalidPublishedMonthPlan"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":9794
            Key             =   "MonthPlan"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":9D2E
            Key             =   "PublishedMonthPlan"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":A2C8
            Key             =   "InvalidPublishedWeekPlan"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":A862
            Key             =   "InvalidWeekPlan"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":ADFC
            Key             =   "PublishedWeekPlan"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":B396
            Key             =   "WeekPlan"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":B930
            Key             =   "MonthTemplet"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":BECA
            Key             =   "MonthTempletDay"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":C464
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":C9FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":CF98
            Key             =   "WeekTemplet"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.ImageManager imgIcons 
      Left            =   2640
      Top             =   6000
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmClinicPlanMainFun.frx":D532
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   210
      Top             =   90
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmClinicPlanMainFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mcbsMain As Object            'CommandBar�ؼ�
Private mstrPrivs As String
Private mlngModule As Long

Private Enum ShortItemID
    ID_BaseItem = 10
    ID_PlanItem = 20
End Enum

Private mintCurYear As Integer '��ǰ��ʾ���
Private mstrYear As String '��ѡ��ݣ������"|"�ָ�
Private mblnNotClick As Boolean

Private mcllVisitTable As Collection 'Array(����ID,�������,�Ű෽ʽ,ģ������)

Public Sub InitCommVariable(frmParent As Form, cbsThis As Object, _
    ByVal strPrivs As String, ByVal lngModule As Long)
    '��ʼ������
    Set mfrmMain = frmParent
    Set mcbsMain = cbsThis
    mstrPrivs = strPrivs
    mlngModule = lngModule
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Val(lblYearSelect.Tag) = Val(Control.Parameter) Then Exit Sub
    
    lblYearSelect.Caption = Val(Control.Parameter) & "��"
    lblYearSelect.Tag = Val(Control.Parameter)
    mintCurYear = Val(Control.Parameter)
    Call LoadVisitTable
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Call mfrmMain.ActiveFormChange(Me)
End Sub

Private Sub Form_Load()
    Dim tpGroup As TaskPanelGroup
    Err = 0: On Error GoTo errHandler
    
    mintCurYear = Year(Now)
    lblYearSelect.Caption = mintCurYear & "��"
    With tplFunBase
        .Behaviour = xtpTaskPanelBehaviourList
        .HotTrackStyle = xtpTaskPanelHighlightItem
        .SelectItemOnFocus = True
        .Icons.AddIcons imgIcons.Icons
        .SetIconSize 32, 32
        .ItemLayout = xtpTaskItemLayoutImagesWithTextBelow
        .SetMargins 1, 0, 0, 1, 2
        
        Set tpGroup = .Groups.Add(10, "��������")
        tpGroup.Items.Add Pane_WorkTime, "�ϰ�ʱ�����", xtpTaskItemTypeLink, 12
        tpGroup.Items.Add Pane_Holiday, "�ڼ��չ���", xtpTaskItemTypeLink, 13
        tpGroup.Items.Add Pane_DoctorOffice, "�������ҹ���", xtpTaskItemTypeLink, 14
        tpGroup.Items.Add Pane_SignalSource, "�ٴ���Դ����", xtpTaskItemTypeLink, 15
        
        tpGroup.CaptionVisible = False
        tpGroup.Expanded = True
    End With
    Call CreateShortcutBar
    Call LoadYear
    Call LoadVisitTable
    
    With scbFunc
        .Tag = ID_BaseItem
        mblnNotClick = True
        .Selected = .Item(1) 'Ҫ�л�һ�£���֤�ؼ��󶨵�λ\
        mblnNotClick = False
        .Selected = .Item(0)
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function LoadVisitTable() As Boolean
    '���ܣ����س�����б�
    '˵�����ɽڵ�Keyֵ��ʽ���ֳ�������
    '       ͣ�ﰲ��:K_StopPlan
    '       �����ģ��ڵ㣺K0_����ID
    '       �̶����ŵĹ̶��ڵ㣺K_FixedRoot
    '       �̶������ڵ㣺K1_����ID
    '       XX������ڵ㣺K_���
    '       XX�³����ڵ㣺K2_���_�·�
    '       XX�ܳ����ڵ㣺K3_���_�·�_����
    Dim strSQL As String, strWhere As String, rsVisitTable As ADODB.Recordset
    Dim objYearNode As Node, objMonthNode As Node, objCurNode As Node
    Dim strKey As String, objNode As Node
    
    Err = 0: On Error GoTo errHandler
    Set mcllVisitTable = New Collection
    tvwPlanTemplet.Nodes.Clear
    tvwPlan.Nodes.Clear
    '�������"ͣ������"��"ͣ������"Ȩ�ޣ��ڳ��ﰲ��������"ͣ�ﰲ��"�ڵ�
    If zlStr.IsHavePrivs(mstrPrivs, "ͣ������") Or zlStr.IsHavePrivs(mstrPrivs, "ͣ������") Then
        tvwPlan.Nodes.Add , , "K_StopPlan", "ͣ�ﰲ��", "StopPlan"
    End If
    tvwPlan.Nodes.Add , , "K_FixedRoot", "�̶�����", "RootPlan"
    
'    'û�����п��ұ�ʾ�ٴ��Ű�
'    If zlStr.IsHavePrivs(mstrPrivs, "���п���") = False Then
'        '���ݹҺŰ��ŵĺ�Դȥ�ж�
'        strWhere = "And Exists" & vbNewLine & _
'                    "       (Select 1" & vbNewLine & _
'                    "         From �ٴ����ﰲ�� M, �ٴ������Դ N" & vbNewLine & _
'                    "         Where m.����id = a.Id And m.��Դid + 0 = n.Id And Nvl(n.�Ƿ��ٴ��Ű�, 0) = 1" & vbNewLine & _
'                    "               And Exists(Select 1 From ������Ա Where ����id = n.����id And ��Աid = [2]))" & vbNewLine
'    End If
    strSQL = "Select a.ID, a.�������, a.�Ű෽ʽ, a.���, a.�·�, a.����," & vbNewLine & _
            "       Decode(a.����ʱ��, Null, 0, 1) As �Ƿ񷢲�,Nvl(ģ������,0) As ģ������," & vbNewLine & _
            "       Nvl((Select 1 From �ٴ����ﰲ�� Where ����id = a.Id And ��ֹʱ�� >= Trunc(Sysdate) And Rownum < 2), 0) As �Ƿ���Ч" & vbNewLine & _
            " From �ٴ������ A" & vbNewLine & _
            " Where ((Nvl(a.�Ű෽ʽ, 0) = 3 And (nvl(a.Ӧ�÷�Χ,0)=2 or nvl(a.Ӧ�÷�Χ,0)=0 and a.������=[3]" & vbNewLine & _
            "        Or (Nvl(a.Ӧ�÷�Χ, 0) = 1 And a.����id In (Select ����id From ������Ա Where ��Աid = [2]))))" & vbNewLine & _
            "       Or (Nvl(a.�Ű෽ʽ, 0) In (1, 2) And a.��� = [1]) Or Nvl(a.�Ű෽ʽ, 0) = 0)" & vbNewLine & _
            "       And Nvl(վ��,'-')=Nvl([4],'-')" & vbNewLine & _
            strWhere & vbNewLine & _
            " Order By Decode(a.�Ű෽ʽ, 0, 0, 3, 3, 1), a.���, a.�·�, a.����, a.ID"
    Set rsVisitTable = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mintCurYear, UserInfo.id, UserInfo.����, gstrNodeNo)
    If rsVisitTable Is Nothing Then GoTo ClearFixedRoot:
    If rsVisitTable.RecordCount = 0 Then GoTo ClearFixedRoot:
    
    '�Ű෽ʽ��0-�̶��Ű�;1-�����Ű�;2-�����Ű�;3-ģ��
    With rsVisitTable
        Do While Not .EOF
            'Array(����ID,�������,�Ű෽ʽ,ģ������)
            mcllVisitTable.Add Array(Nvl(!id), Nvl(!�������), Nvl(!�Ű෽ʽ), Val(Nvl(!ģ������))), "K0_" & Nvl(!id)
            If Nvl(!�Ű෽ʽ) = 3 Then  'ģ��,Keyֵ��ʽ��K0_����ID
                strKey = "K0_" & Nvl(!id)
                Set objNode = tvwPlanTemplet.Nodes.Add(, , strKey, Nvl(!�������), Decode(Val(Nvl(!ģ������)), 1, "MonthTemplet", 2, "MonthTempletDay", "WeekTemplet"))
                objNode.Tag = Val(Nvl(!id))
            ElseIf Nvl(!�Ű෽ʽ) = 0 Then  '�̶��Ű�,Keyֵ��ʽ��K1_����ID
                strKey = "K1_" & Nvl(!id)
                If Not FindNodeByKey(tvwPlan.Nodes, "K_FixedRoot") Is Nothing Then
                    Set objNode = tvwPlan.Nodes("K_FixedRoot")
                    objNode.Text = Nvl(!�������)
                    objNode.Key = strKey
                    objNode.Tag = Val(Nvl(!id))
                    objNode.Image = GetIconIndex(1, Val(Nvl(!�Ƿ񷢲�)) = 1, Val(Nvl(!�Ƿ���Ч)) = 0)
                End If
            Else '���Ű�����Ű�
                '1.��ݽڵ�,Keyֵ��ʽ��K_���
                strKey = "K_" & Nvl(!���)
                If FindNodeByKey(tvwPlan.Nodes, strKey) Is Nothing Then
                    Set objYearNode = tvwPlan.Nodes.Add(, , strKey, Nvl(!���) & "����ﰲ��", "RootPlan")
                Else
                    Set objYearNode = tvwPlan.Nodes(strKey)
                End If
                '2.�·ݽڵ�,Keyֵ��ʽ��K2_���_�·�
                strKey = "K2_" & Nvl(!���) & "_" & Nvl(!�·�)
                If FindNodeByKey(tvwPlan.Nodes, strKey) Is Nothing Then
                    Set objMonthNode = tvwPlan.Nodes.Add(objYearNode, tvwChild, strKey, Nvl(!�·�) & "�³����", "InvalidMonthPlan")
                    If Val(Nvl(!����)) = 0 Then
                        objMonthNode.Tag = Val(Nvl(!id))
                        objMonthNode.Image = GetIconIndex(2, Val(Nvl(!�Ƿ񷢲�)) = 1, Val(Nvl(!�Ƿ���Ч)) = 0)
                    End If
                Else
                    Set objMonthNode = tvwPlan.Nodes(strKey)
                    If Val(Nvl(!����)) = 0 Then
                        objMonthNode.Tag = Val(Nvl(!id))
                        objMonthNode.Image = GetIconIndex(2, Val(Nvl(!�Ƿ񷢲�)) = 1, Val(Nvl(!�Ƿ���Ч)) = 0)
                    End If
                End If
                '3.�����ڵ㣬Keyֵ��ʽ��K3_���_�·�_����
                If Nvl(!�Ű෽ʽ) = 2 Then  '���Ű�
                    strKey = "K3_" & Nvl(!���) & "_" & Nvl(!�·�) & "_" & Nvl(!����)
                    Set objNode = tvwPlan.Nodes.Add(objMonthNode, tvwChild, strKey, "��" & Nvl(!����) & "�ܳ����", "InvalidWeekPlan")
                    objNode.Tag = Val(Nvl(!id))
                    objNode.Image = GetIconIndex(3, Val(Nvl(!�Ƿ񷢲�)) = 1, Val(Nvl(!�Ƿ���Ч)) = 0)
                End If
            End If
             .MoveNext
        Loop
    End With
    
    'չ���ڵ�
    For Each objCurNode In tvwPlan.Nodes
        objCurNode.Expanded = True
        
        '�����³����ڵ㣬ȷ��ͼ��
        'ֻ�������ӽڵ㼰�Լ���Ϊ��Чʱ����ʾΪ��Ч�ڵ�
        If InStr(objCurNode.Key, "_") > 0 Then
            If Split(objCurNode.Key, "_")(0) = "K2" Then
                If Val(objCurNode.Tag) <> 0 Then
                    '���³���������³����Ϊ׼
                    '����³������Ч�������ܳ����϶�Ҳ��Ч
                Else
                    'ֻ���ܳ�������³����
                    Set objNode = objCurNode.Child
                    Do While Not objNode Is Nothing
                        If objNode.Image = "WeekPlan" Or objCurNode.Image = "PublishedWeekPlan" Then
                            'һ���ܳ������Ч�����³����ڵ���Ч
                            objCurNode.Image = "MonthPlan"
                        End If
                        Set objNode = objNode.Next
                    Loop
                End If
            End If
        End If
    Next
    
ClearFixedRoot:
    LoadVisitTable = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetIconIndex(ByVal bytMode As Byte, _
    Optional ByVal blnPublished As Boolean, Optional ByVal blnInvalid As Boolean) As String
    '��ȡ���ﰲ�Žڵ�ͼ������
    '��Σ�
    '   bytMode 1-�̶������,2-�³����,3-�ܳ����
    Select Case bytMode
    Case 1
        If blnPublished Then
            If blnInvalid Then
                GetIconIndex = "InvalidPublishedFixedPlan"
            Else
                GetIconIndex = "PublishedFixedPlan"
            End If
        Else
            If blnInvalid Then
                GetIconIndex = "InvalidFixedPlan"
            Else
                GetIconIndex = "FixedPlan"
            End If
        End If
    Case 2
        If blnPublished Then
            If blnInvalid Then
                GetIconIndex = "InvalidPublishedMonthPlan"
            Else
                GetIconIndex = "PublishedMonthPlan"
            End If
        Else
            If blnInvalid Then
                GetIconIndex = "InvalidMonthPlan"
            Else
                GetIconIndex = "MonthPlan"
            End If
        End If
    Case 3
    If blnPublished Then
            If blnInvalid Then
                GetIconIndex = "InvalidPublishedWeekPlan"
            Else
                GetIconIndex = "PublishedWeekPlan"
            End If
        Else
            If blnInvalid Then
                GetIconIndex = "InvalidWeekPlan"
            Else
                GetIconIndex = "WeekPlan"
            End If
        End If
    Case Else
        GetIconIndex = "FixedPlan"
    End Select
End Function

Private Sub CreateShortcutBar()
    Err = 0: On Error GoTo errHandler
    With scbFunc
        .Icons = imgIcons.Icons
        
        'ͼ����������ID��ͬ
        .AddItem ID_PlanItem, "���ﰲ��", picPlanBack.Hwnd
        .AddItem ID_BaseItem, "��������", picBaseSetBack.Hwnd
        .ExpandedLinesCount = .ItemCount 'Ĭ��չ��
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picFunBack.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub frmMoveSplitY_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Err = 0: On Error Resume Next
    If Button <> vbLeftButton Then Exit Sub
    If tvwPlanTemplet.Height + Y < 1200 Or tvwPlan.Height - Y < 1500 Then Exit Sub
    
    tvwPlanTemplet.Height = tvwPlanTemplet.Height + Y
    frmMoveSplitY.Top = frmMoveSplitY.Top + Y
    sccPlan.Top = sccPlan.Top + Y
    lblYearSelect.Top = sccPlan.Top + (sccPlan.Height - lblYearSelect.Height) / 2
    imgYearSelect.Top = sccPlan.Top + (sccPlan.Height - imgYearSelect.Height) / 2
    tvwPlan.Top = tvwPlan.Top + Y
    tvwPlan.Height = tvwPlan.Height - Y
End Sub

Private Sub imgYearSelect_Click()
    Call lblYearSelect_Click
End Sub

Private Sub lblYearSelect_Click()
    Call ShowPopuYear
End Sub

Public Function RefreshVisitTable(Optional ByVal strKey As String) As Boolean
    'ˢ�³����
    Dim objNode As Node
    Dim strDeletedPreviouNodeKey As String
    
    Err = 0: On Error GoTo errHandler
    If scbFunc.Selected.id <> ID_PlanItem Then Exit Function
    
    If strKey = "" Then
        If Me.ActiveControl Is tvwPlanTemplet Then
            If Not tvwPlanTemplet.SelectedItem Is Nothing Then Set objNode = tvwPlanTemplet.SelectedItem
        Else
            If Not tvwPlan.SelectedItem Is Nothing Then Set objNode = tvwPlan.SelectedItem
        End If
        'ɾ���ڵ�ʱ��ȷ��ѡ�нڵ��Keyֵ
        If Not objNode Is Nothing Then
            strKey = objNode.Key
            If Not objNode.Previous Is Nothing Then
                strDeletedPreviouNodeKey = objNode.Previous.Key
            ElseIf Not objNode.Next Is Nothing Then
                strDeletedPreviouNodeKey = objNode.Next.Key
            ElseIf Not objNode.Parent Is Nothing Then
                strDeletedPreviouNodeKey = objNode.Parent.Key
            End If
        End If
    Else
        If Left(strKey, 2) = "K0" Then
            If tvwPlanTemplet.Visible And tvwPlanTemplet.Enabled Then tvwPlanTemplet.SetFocus
        Else
            If tvwPlan.Visible And tvwPlan.Enabled Then tvwPlan.SetFocus
        End If
    End If
    
    Call LoadYear
    Call LoadVisitTable '���¼��س����
    
    'ģ��
    If Me.ActiveControl Is tvwPlanTemplet Then
        '�ȶ�λ����ѡ����
        Set objNode = FindNodeByKey(tvwPlanTemplet.Nodes, strKey)
        If Not objNode Is Nothing Then
            tvwPlanTemplet.Tag = ""
            objNode.Selected = True
            tvwPlanTemplet_NodeClick objNode
            tvwPlanTemplet.SetFocus
            RefreshVisitTable = True: Exit Function
        Else
            '��λ����һ��
            Set objNode = FindNodeByKey(tvwPlanTemplet.Nodes, strDeletedPreviouNodeKey)
            If Not objNode Is Nothing Then
                tvwPlanTemplet.Tag = ""
                objNode.Selected = True
                tvwPlanTemplet_NodeClick objNode
                tvwPlanTemplet.SetFocus
                RefreshVisitTable = True: Exit Function
            End If
        End If
    Else
        '���ﰲ��
        Set objNode = FindNodeByKey(tvwPlan.Nodes, strKey)
        If Not objNode Is Nothing Then
            tvwPlan.Tag = ""
            objNode.Selected = True
            tvwPlan_NodeClick objNode
            tvwPlan.SetFocus
            RefreshVisitTable = True: Exit Function
        Else
            Set objNode = FindNodeByKey(tvwPlan.Nodes, strDeletedPreviouNodeKey)
            If Not objNode Is Nothing Then
                tvwPlan.Tag = ""
                objNode.Selected = True
                tvwPlan_NodeClick objNode
                tvwPlan.SetFocus
                RefreshVisitTable = True: Exit Function
            End If
        End If
    End If
    
    Call tvwPlan_GotFocus
    RefreshVisitTable = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub picBaseSetBack_Resize()
    Err = 0: On Error Resume Next
    sccFunBase.Move 0, -10, picBaseSetBack.ScaleWidth
    With tplFunBase
        .Left = 0
        .Top = sccFunBase.Top + sccFunBase.Height
        .Width = picBaseSetBack.ScaleWidth - .Left
        .Height = picBaseSetBack.ScaleHeight - .Top
    End With
End Sub

Private Sub picPlanBack_Resize()
    Err = 0: On Error Resume Next
    sccPlanTemplet.Move 0, -10, picPlanBack.ScaleWidth
    tvwPlanTemplet.Move 0, sccPlanTemplet.Top + sccPlanTemplet.Height, picPlanBack.ScaleWidth
    frmMoveSplitY.Move -25, tvwPlanTemplet.Top + tvwPlanTemplet.Height, picPlanBack.ScaleWidth + 100
    
    sccPlan.Move 0, frmMoveSplitY.Top + frmMoveSplitY.Height, picPlanBack.ScaleWidth
    imgYearSelect.Top = sccPlan.Top + (sccPlan.Height - imgYearSelect.Height) / 2
    imgYearSelect.Left = sccPlan.Width - imgYearSelect.Width - 10
    lblYearSelect.Top = sccPlan.Top + (sccPlan.Height - lblYearSelect.Height) / 2
    lblYearSelect.Left = imgYearSelect.Left - lblYearSelect.Width - 10
    tvwPlan.Move 0, sccPlan.Top + sccPlan.Height, picPlanBack.ScaleWidth, picPlanBack.ScaleHeight - (sccPlan.Top + sccPlan.Height)
End Sub

Private Sub picFunBack_Resize()
    Err = 0: On Error Resume Next
    scbFunc.Move 0, 0, picFunBack.ScaleWidth, picFunBack.ScaleHeight
End Sub

Private Sub scbFunc_ExpandButtonDown(CancelMenu As Boolean)
    CancelMenu = True
End Sub

Private Sub scbFunc_SelectedChanged(ByVal Item As XtremeSuiteControls.IShortcutBarItem)
    Dim tpGroup As TaskPanelGroup
    Dim tpItem As TaskPanelGroupItem
    Dim blnFind As Boolean, tpItemWork As TaskPanelGroupItem
    
    Err = 0: On Error GoTo errHandler
    If mblnNotClick Then Exit Sub
    If Val(scbFunc.Tag) = Item.id Then Exit Sub
    
    tvwPlanTemplet.Tag = ""
    tvwPlan.Tag = ""
    
    '����Ĭ��ѡ�нڵ�
    picBaseSetBack.Visible = False
    picPlanBack.Visible = False
    If Item.id = ID_BaseItem Then
        scbFunc.Tag = ID_BaseItem
        If tplFunBase.Tag = "" Then tplFunBase.Tag = "�ٴ���Դ����" 'ȱʡѡ���ٴ���Դ����
        For Each tpGroup In tplFunBase.Groups
            For Each tpItem In tpGroup.Items
                If tpItem.Caption = tplFunBase.Tag Then Set tpItemWork = tpItem
                If tplFunBase.Tag = tpItem.Caption Then
                    tpItem.Selected = True: blnFind = True
                    tplFunBase.Tag = "": tplFunBase_ItemClick tpItem
                Else
                    tpItem.Selected = False
                End If
            Next
        Next
        If blnFind = False Then
            'ȱʡѡ���ϰ�ʱ��
            tpItemWork.Selected = True
            tplFunBase.Tag = "": tplFunBase_ItemClick tpItemWork
        End If
        picBaseSetBack.Visible = True
    Else
        scbFunc.Tag = ID_PlanItem
        Call tvwPlan_GotFocus
        
        picPlanBack.Visible = True
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub tplFunBase_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    Err = 0: On Error GoTo errHandler
    If tplFunBase.Tag = Item.Caption Then Exit Sub
    tplFunBase.Tag = Item.Caption
    
    Select Case Item.Caption
    Case "�ϰ�ʱ�����"
        Call mfrmMain.SelectedChange(Pane_WorkTime)
    Case "�ڼ��չ���"
        Call mfrmMain.SelectedChange(Pane_Holiday)
    Case "�������ҹ���"
        Call mfrmMain.SelectedChange(Pane_DoctorOffice)
    Case "�ٴ���Դ����"
        Call mfrmMain.SelectedChange(Pane_SignalSource)
    End Select
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub tvwPlan_GotFocus()
    Err = 0: On Error GoTo errHandler
    
    tvwPlanTemplet.Tag = ""
    tvwPlanTemplet.HideSelection = True
    tvwPlan.HideSelection = False
    
    If tvwPlan.Nodes.Count > 0 Then
        If tvwPlan.SelectedItem Is Nothing Then
            tvwPlan.Tag = ""
            If tvwPlan.Nodes.Count > 1 Then
                tvwPlan.Nodes(2).Selected = True
            Else
                tvwPlan.Nodes(1).Selected = True
            End If
        End If
        tvwPlan_NodeClick tvwPlan.SelectedItem
    Else
        Call mfrmMain.SelectedChange(Pane_FixedPlan)
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub tvwPlan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '�����Ҽ��˵�
    Dim cbCommandBar As CommandBar
    
    Err = 0: On Error GoTo errHandler
    If Not (Button = vbRightButton) Then Exit Sub
    If Not (tvwPlan.Visible And tvwPlan.Enabled) Then Exit Sub
    tvwPlan.SetFocus: Call mfrmMain.ActiveFormChange(Me)
    Call tvwPlan_GotFocus
    
    Set cbCommandBar = mfrmMain.GetPopupCommandBarSub()
    If cbCommandBar Is Nothing Then Exit Sub
    If cbCommandBar.Controls.Count = 0 Then Exit Sub
    
    cbCommandBar.ShowPopup
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub tvwPlan_NodeClick(ByVal Node As MSComctlLib.Node)
    Err = 0: On Error GoTo errHandler
    If tvwPlan.Tag = Node.Key Then Exit Sub

    tvwPlanTemplet.Tag = ""
    tvwPlan.Tag = Node.Key
    tvwPlan.HideSelection = False

    Select Case Split(Node.Key, "_")(0)
    Case "K1" '�̶�����
        Call mfrmMain.SelectedChange(Pane_FixedPlan, Val(Node.Tag))
    Case "K2" '���Ű�
        Call mfrmMain.SelectedChange(Pane_MonthPlan, Val(Node.Tag), _
            Val(Split(Node.Key, "_")(1)), Val(Split(Node.Key, "_")(2)), _
            Val(Split(Node.Key, "_")(1)) & "��" & Val(Split(Node.Key, "_")(2)) & "��")
    Case "K3" '���Ű�
        Call mfrmMain.SelectedChange(Pane_WeekPlan, Val(Node.Tag))
    Case Else
        If Node.Key = "K_StopPlan" Then 'ͣ�����
            Call mfrmMain.SelectedChange(Pane_StopPlan)
        ElseIf Node.Parent Is Nothing Then
            If Node.Key = "K_FixedRoot" Then
                Call mfrmMain.SelectedChange(Pane_FixedPlan, Val(Node.Tag))
            Else
                Call mfrmMain.SelectedChange(Pane_WeekPlan, Val(Node.Tag), 0, 0, "�����")
            End If
        End If
    End Select
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub tvwPlanTemplet_GotFocus()
    Err = 0: On Error GoTo errHandler
    
    tvwPlan.Tag = ""
    tvwPlanTemplet.HideSelection = False
    tvwPlan.HideSelection = True
    
    If tvwPlanTemplet.Nodes.Count > 0 Then
        If tvwPlanTemplet.SelectedItem Is Nothing Then
            tvwPlanTemplet.Tag = ""
            tvwPlanTemplet.Nodes(1).Selected = True
        End If
        tvwPlanTemplet_NodeClick tvwPlanTemplet.SelectedItem
    Else
        Call mfrmMain.SelectedChange(Pane_PlanTemplet)
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub tvwPlanTemplet_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '�����Ҽ��˵�
    Dim cbCommandBar As CommandBar
    
    Err = 0: On Error GoTo errHandler
    If Not (Button = vbRightButton) Then Exit Sub
    If Not (tvwPlanTemplet.Visible And tvwPlanTemplet.Enabled) Then Exit Sub
    tvwPlanTemplet.SetFocus: Call mfrmMain.ActiveFormChange(Me)
    Call tvwPlanTemplet_GotFocus
    
    Set cbCommandBar = mfrmMain.GetPopupCommandBarSub()
    If cbCommandBar Is Nothing Then Exit Sub
    If cbCommandBar.Controls.Count = 0 Then Exit Sub
    
    cbCommandBar.ShowPopup
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub tvwPlanTemplet_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim bytTempletType As Byte
    
    Err = 0: On Error GoTo errHandler
    If tvwPlanTemplet.Tag = Node.Key Then Exit Sub
    
    tvwPlan.Tag = ""
    tvwPlanTemplet.Tag = Node.Key
    tvwPlanTemplet.HideSelection = False
    
    'Array(����ID,�������,�Ű෽ʽ,ģ������)
    If CollExitsValue(mcllVisitTable, Node.Key) Then
        bytTempletType = Val(mcllVisitTable(Node.Key)(3))
    End If
    If bytTempletType = 2 Then
        Call mfrmMain.SelectedChange(Pane_MonthTemplet, Val(Node.Tag))
    Else
        Call mfrmMain.SelectedChange(Pane_PlanTemplet, Val(Node.Tag), 0, 0, "", bytTempletType)
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub LoadYear()
    '���ؿ�ѡ���
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strYear As String
    Dim blnFind As Boolean
    
    Err = 0: On Error GoTo errHandler
    mstrYear = ""
    strSQL = "Select Distinct ��� From �ٴ������ Where �Ű෽ʽ In (1, 2) And Nvl(վ��,'-') = Nvl([1],'-') Order By ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, gstrNodeNo)
    If Not rsTemp.EOF Then
        Do While Not rsTemp.EOF
            mstrYear = mstrYear & "|" & Nvl(rsTemp!���)
            If Val(Nvl(rsTemp!���)) = mintCurYear Then blnFind = True
            rsTemp.MoveNext
        Loop
    End If
    If mstrYear = "" Then
        mstrYear = mintCurYear
    Else
        mstrYear = Mid(mstrYear, 2)
    End If
    If blnFind = False Then
        mintCurYear = Split(mstrYear, "|")(UBound(Split(mstrYear, "|")))
        lblYearSelect.Caption = mintCurYear & "��"
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function CreatePopuMenu(ByVal strYear As String) As CommandBar
    '����:������ʱ�˵�
    Dim objCommandBar As CommandBar
    Dim objControl As CommandBarControl
    Dim i As Integer, varYear As Variant
    
    If strYear = "" Then Exit Function
    
    cbsThis.DeleteAll
    Set objCommandBar = cbsThis.Add("PopupYear", xtpBarPopup)
    With objCommandBar.Controls
        varYear = Split(strYear, "|")
        For i = 0 To UBound(varYear)
            Set objControl = .Add(xtpControlButton, 1000 + i, Val(varYear(i)) & "��")
            objControl.Parameter = Val(varYear(i))
            If Val(varYear(i)) = mintCurYear Then
                objControl.Checked = True
            End If
        Next
    End With
    Set CreatePopuMenu = objCommandBar
    Set objCommandBar = Nothing
End Function

Private Sub ShowPopuYear()
    Dim objCommandBar As CommandBar
    Dim vRect As RECT
    
    vRect = zlControl.GetControlRect(picPlanBack.Hwnd)
    vRect.Left = vRect.Left + lblYearSelect.Left - 2
    vRect.Top = vRect.Top + lblYearSelect.Top + 2
    Set objCommandBar = CreatePopuMenu(mstrYear)
    If objCommandBar Is Nothing Then Exit Sub
    
    Call objCommandBar.ShowPopup(, vRect.Left, vRect.Top + lblYearSelect.Height)
    Set objCommandBar = Nothing
End Sub


