VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmElements 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "���ʵ�Ԫ��"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   9525
   Icon            =   "frmElements.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txt���� 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2700
      TabIndex        =   10
      Top             =   1500
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3120
      TabIndex        =   11
      Top             =   1170
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.ComboBox cmb��� 
      Height          =   300
      Left            =   3090
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   825
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "��"
      Height          =   300
      Left            =   4500
      TabIndex        =   12
      Top             =   1140
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Frame fra��ѡ 
      Caption         =   "��ѡ��Ŀ"
      Height          =   3285
      Left            =   120
      TabIndex        =   6
      Top             =   1590
      Width           =   2745
      Begin VB.ListBox lst��Ŀ 
         Height          =   2790
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   270
         Width           =   2475
      End
   End
   Begin VB.Frame fraFix 
      Caption         =   "�̶���Ŀ"
      Height          =   1275
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   2745
      Begin VB.ComboBox cmb���� 
         Height          =   300
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   750
         Width           =   1485
      End
      Begin MSComCtl2.UpDown ud��Ŀ�� 
         Height          =   300
         Left            =   2385
         TabIndex        =   3
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txt��Ŀ��"
         BuddyDispid     =   196618
         OrigLeft        =   2370
         OrigTop         =   450
         OrigRight       =   2610
         OrigBottom      =   765
         Max             =   100
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txt��Ŀ�� 
         Height          =   300
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   885
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&D)"
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   810
         Width           =   990
      End
      Begin VB.Label lbl��Ŀ���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�շ���Ŀ����(&N)"
         Height          =   180
         Left            =   120
         TabIndex        =   1
         Top             =   420
         Width           =   1350
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6570
      TabIndex        =   15
      Tag             =   "����"
      Top             =   5010
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5250
      TabIndex        =   14
      Tag             =   "����"
      Top             =   5010
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   360
      TabIndex        =   16
      Tag             =   "����"
      Top             =   5010
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh�շ���Ŀ 
      Height          =   4395
      Left            =   2970
      TabIndex        =   13
      Top             =   480
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   7752
      _Version        =   393216
      Rows            =   12
      RowHeightMin    =   320
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      GridColor       =   8421504
      GridColorFixed  =   8421504
      GridLinesFixed  =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lbl�շ���Ŀ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�շ���Ŀ(&S)"
      Height          =   180
      Left            =   3030
      TabIndex        =   8
      Top             =   210
      Width           =   990
   End
End
Attribute VB_Name = "frmElements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum RowSign
    row�շ����ֵ = 1
    row�շ�ϸĿֵ = 2
    row����ֵ = 3
    row�շ���� = 4
    rowϸĿѡ�� = 5
    row���㵥λ = 6
    row���� = 7
    row��׼���� = 8
    rowӦ�ս�� = 9
    rowʵ�ս�� = 10
    rowִ�в��� = 11
    row���ӱ�־ = 12
End Enum

Dim mblnNew As Boolean  '��ǰ�޸ĵĵ����Ƿ���������
Dim mcolBill As Elements
Dim mblnOK As Boolean
Dim mblnChange As Boolean
Dim mlngItemCount As Long

Public Function ModifyElement(colBill As Elements, ItemCount As Long, Optional ByVal ���� As Boolean) As Boolean
'�޸ġ����ʵ�Ԫ�ء��Ĺ������
    Dim lngCount As Long, lngRow As Long
    Dim lngID As Long, strTemp As String
    
    mblnOK = False
    mblnNew = ����
    mlngItemCount = ItemCount
    Set mcolBill = colBill
    
    
    Call InitCont
    If ���� = True Then
        cmb����.ListIndex = 0
        For lngCount = 0 To lst��Ŀ.ListCount - 1
            lst��Ŀ.Selected(lngCount) = True
        Next
        
        msh�շ���Ŀ.TextMatrix(row����ֵ, 1) = "1"
        For lngCount = row�շ���� To msh�շ���Ŀ.Rows - 1
            msh�շ���Ŀ.TextMatrix(lngCount, 1) = "��"
        Next
    Else
        '�̶���Ŀ
        ud��Ŀ��.Value = ItemCount '�Զ���ı��������
        
        If Mid(colBill("��������").Value, 1, 1) = "C" Then
            cmb����.ListIndex = Mid(colBill("��������").Value, 2, 1)
        ElseIf colBill("��������").Value <> "" Then
            lngID = Val(colBill("��������").Value)
            For lngCount = 4 To cmb����.ListCount - 1
                If cmb����.ItemData(lngCount) = lngID Then
                    cmb����.ListIndex = lngCount
                    Exit For
                End If
            Next
            If cmb����.ListIndex < 0 Then cmb����.ListIndex = 0
        Else
            cmb����.ListIndex = 0
        End If
        
        '��ѡ��Ŀ
        For lngCount = 0 To lst��Ŀ.ListCount - 1
            lst��Ŀ.Selected(lngCount) = colBill(lst��Ŀ.List(lngCount)).Visible
        Next
        '�շ���Ŀ
        With msh�շ���Ŀ
            For lngCount = 1 To ud��Ŀ��.Value
                For lngRow = 1 To .Rows - 1
                    Select Case .TextMatrix(lngRow, 0)
                        Case "�շ����ֵ"
                            .TextMatrix(lngRow, lngCount) = GetClassName(colBill("�շ����" & "_" & lngCount).Value)
                        Case "�շ�ϸĿֵ"
                            lngID = colBill("�շ�ϸĿ" & "_" & lngCount).Value
                            strTemp = GetItemName(Abs(lngID))
                            If strTemp <> "" Then
                                .ColData(lngCount) = lngID
                                .TextMatrix(lngRow, lngCount) = strTemp
                            End If
                        Case "����ֵ"
                            .TextMatrix(lngRow, lngCount) = colBill("����" & "_" & lngCount).Value
                        Case Else
                            .TextMatrix(lngRow, lngCount) = IIf(colBill(.TextMatrix(lngRow, 0) & "_" & lngCount).Visible, "��", "")
                    End Select
                Next
            Next
            .Row = row�շ����
            .LeftCol = 1
            
        End With
        
    End If
    
    txt��Ŀ��.Text = ud��Ŀ��.Value
    cmb���.Visible = False
    mblnChange = False
    frmElements.Show vbModal, frmDesign
    ModifyElement = mblnOK
    '��������
    If mblnOK = True Then
        ItemCount = mlngItemCount
    End If
End Function

Private Sub cmb����_Click()
    '����ֻ����Ϊ�ָ���ʹ��
    mblnChange = True
    If cmb����.Text = "����������������������������" Then cmb����.ListIndex = 0
End Sub

Private Sub cmb���_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dim blnCancel As Boolean
        
        Call cmb���_Validate(blnCancel)
        msh�շ���Ŀ.Row = 2
        Call msh�շ���Ŀ_EnterCell
    End If
End Sub

Private Sub cmb���_Validate(Cancel As Boolean)
    If msh�շ���Ŀ.Text <> cmb���.Text Then
        msh�շ���Ŀ.Text = cmb���.Text
        msh�շ���Ŀ.TextMatrix(row�շ�ϸĿֵ, msh�շ���Ŀ.Col) = ""
        msh�շ���Ŀ.ColData(msh�շ���Ŀ.Col) = 0
        If Left(cmb���.Text, 1) = "0" Then
            msh�շ���Ŀ.TextMatrix(row�շ����, msh�շ���Ŀ.Col) = "��"
        End If
    End If
    
    mblnChange = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
ShowHelp "zl9custacc", Me.hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    cmb���.Visible = False
    Call SaveControls
    mblnOK = True
    mblnChange = False
    mlngItemCount = ud��Ŀ��.Value
    Unload Me
End Sub

Private Sub SaveControls()
    Dim ctlTemp As Control
    Dim lngCount As Long
    Dim lngRow As Long
    Dim blnVisible As Boolean
    Dim strTemp As String, strControl As String
    Dim lngLeft As Long, lngWidth As Long
    Dim arrItems As Variant
    
    If mblnNew = True Then
        '���Ȱ���Щû�����ڶԻ����У������Ǳ���Ŀؼ�����
        mcolBill.Clear
        '����
            Set ctlTemp = LoadControl("Label", 225, 180, , , 0)
            mcolBill.Add "��ǩ_0", ctlTemp, , True
        'No
            Set ctlTemp = LoadControl("Label", 8685, 720)
            mcolBill.Add "��ǩ_" & ctlTemp.Index, ctlTemp, , True
            ctlTemp.Caption = "���ݺ�"
            Set ctlTemp = LoadControl("ComboBox", 9465, 660, 1425, , 0)
            mcolBill.Add "NO", ctlTemp, , True
        '����ʱ��
            Set ctlTemp = LoadControl("Label", 8160, 5580)
            mcolBill.Add "��ǩ_" & ctlTemp.Index, ctlTemp, , True
            ctlTemp.Caption = "ʱ��"
            Set ctlTemp = LoadControl("TextBox", 8610, 5520, 2400, , 0)
            mcolBill.Add "����ʱ��", ctlTemp, , True
        '����
            Set ctlTemp = LoadControl("Label", 75, 1125)
            mcolBill.Add "��ǩ_" & ctlTemp.Index, ctlTemp, , True
            ctlTemp.Caption = "����"
            Set ctlTemp = LoadControl("TextBox", 555, 1065, 1365)
            mcolBill.Add "����", ctlTemp, , True
        '��������id
            Set ctlTemp = LoadControl("Label", 8760, 1140)
            mcolBill.Add "��ǩ_" & ctlTemp.Index, ctlTemp, , True
            ctlTemp.Caption = "����"
            Set ctlTemp = LoadControl("ComboBox", 9195, 1080, 2055)
            mcolBill.Add "��������", ctlTemp, , True
        '������
            Set ctlTemp = LoadControl("Label", 5250, 5585)
            mcolBill.Add "��ǩ_" & ctlTemp.Index, ctlTemp, , True
            ctlTemp.Caption = "ҽ��"
            Set ctlTemp = LoadControl("ComboBox", 5625, 5520, 2085)
            mcolBill.Add "������", ctlTemp, , True
        'ȷ��
            Set ctlTemp = LoadControl("CommandButton", 8295, 6120, , , 1)
            mcolBill.Add "ȷ��", ctlTemp, , True
        'ȡ��
            Set ctlTemp = LoadControl("CommandButton", 9690, 6120, , , 2)
            mcolBill.Add "ȡ��", ctlTemp, , True
        '��
            Set ctlTemp = LoadControl("CheckBox", 10890, 660, 400, , 1)
            mcolBill.Add "��", ctlTemp, , True
            
        'Ȼ�����ѡȡ��Ŀ
        For lngCount = 0 To lst��Ŀ.ListCount - 1
            blnVisible = lst��Ŀ.Selected(lngCount)
            Select Case lst��Ŀ.List(lngCount)
               Case "����ID"
                    If blnVisible = True Then
                        Set ctlTemp = LoadControl("Label", 75, 1510)
                        mcolBill.Add "��ǩ_" & ctlTemp.Index, ctlTemp, , True
                        ctlTemp.Caption = "����ID"
                    End If
                    Set ctlTemp = LoadControl("TextBox", 685, 1455, 1005)
                    mcolBill.Add "����ID", ctlTemp, , blnVisible
               Case "��ʶ��"
                    If blnVisible = True Then
                        Set ctlTemp = LoadControl("Label", 1875, 1510)
                        mcolBill.Add "��ǩ_" & ctlTemp.Index, ctlTemp, , True
                        ctlTemp.Caption = "��ʶ��"
                    End If
                    Set ctlTemp = LoadControl("TextBox", 2475, 1455, 1005)
                    mcolBill.Add "��ʶ��", ctlTemp, , blnVisible
               Case "��Ժ����"
                    If blnVisible = True Then
                        Set ctlTemp = LoadControl("Label", 3585, 1510)
                        mcolBill.Add "��ǩ_" & ctlTemp.Index, ctlTemp, , True
                        ctlTemp.Caption = "��Ժ����"
                    End If
                    Set ctlTemp = LoadControl("TextBox", 4350, 1455, 1005)
                    mcolBill.Add "��Ժ����", ctlTemp, , blnVisible
               Case "�Ա�"
                    If blnVisible = True Then
                        Set ctlTemp = LoadControl("Label", 2085, 1125)
                        mcolBill.Add "��ǩ_" & ctlTemp.Index, ctlTemp, , True
                        ctlTemp.Caption = "�Ա�"
                    End If
                    Set ctlTemp = LoadControl("ComboBox", 2475, 1065, 1005)
                    mcolBill.Add "�Ա�", ctlTemp, , blnVisible
               Case "����"
                    If blnVisible = True Then
                        Set ctlTemp = LoadControl("Label", 3615, 1125)
                        mcolBill.Add "��ǩ_" & ctlTemp.Index, ctlTemp, , True
                        ctlTemp.Caption = "����"
                    End If
                    Set ctlTemp = LoadControl("TextBox", 3990, 1065, 660)
                    mcolBill.Add "����", ctlTemp, , blnVisible
               Case "����"
                    If blnVisible = True Then
                        Set ctlTemp = LoadControl("Label", 6765, 1125)
                        mcolBill.Add "��ǩ_" & ctlTemp.Index, ctlTemp, , True
                        ctlTemp.Caption = "����"
                    End If
                    Set ctlTemp = LoadControl("TextBox", 7185, 1065, 1005)
                    mcolBill.Add "����", ctlTemp, , blnVisible
               Case "���˲���"
                    If blnVisible = True Then
                        Set ctlTemp = LoadControl("Label", 5775, 1510)
                        mcolBill.Add "��ǩ_" & ctlTemp.Index, ctlTemp, , True
                        ctlTemp.Caption = "���˲���"
                    End If
                    Set ctlTemp = LoadControl("TextBox", 6570, 1455, 1500)
                    mcolBill.Add "���˲���", ctlTemp, , blnVisible
               Case "���˿���"
                    If blnVisible = True Then
                        Set ctlTemp = LoadControl("Label", 8265, 1510)
                        mcolBill.Add "��ǩ_" & ctlTemp.Index, ctlTemp, , True
                        ctlTemp.Caption = "���˿���"
                    End If
                    Set ctlTemp = LoadControl("TextBox", 9030, 1455, 1500)
                    mcolBill.Add "���˿���", ctlTemp, , blnVisible
               Case "�ѱ�"
                    If blnVisible = True Then
                        Set ctlTemp = LoadControl("Label", 4800, 1125)
                        mcolBill.Add "��ǩ_" & ctlTemp.Index, ctlTemp, , True
                        ctlTemp.Caption = "�ѱ�"
                    End If
                    Set ctlTemp = LoadControl("ComboBox", 5190, 1065)
                    mcolBill.Add "�ѱ�", ctlTemp, , blnVisible
               Case "�Ӱ��־"
                    Set ctlTemp = LoadControl("CheckBox", 350, 5630, 800, , 0)
                    ctlTemp.Caption = "�Ӱ�"
                    mcolBill.Add "�Ӱ��־", ctlTemp, , blnVisible
               Case "Ӥ����"
                    If blnVisible = True Then
                        Set ctlTemp = LoadControl("Label", 1220, 5630)
                        mcolBill.Add "��ǩ_" & ctlTemp.Index, ctlTemp, , True
                        ctlTemp.Caption = "Ӥ����"
                    End If
                    Set ctlTemp = LoadControl("ComboBox", 1950, 5565, 1435)
                    mcolBill.Add "Ӥ����", ctlTemp, , blnVisible
               Case "Ӧ�պϼ�"
                    If blnVisible = True Then
                        Set ctlTemp = LoadControl("Label", 330, 6150)
                        mcolBill.Add "��ǩ_" & ctlTemp.Index, ctlTemp, , True
                        ctlTemp.Caption = "Ӧ��"
                    End If
                    Set ctlTemp = LoadControl("TextBox", 720, 6105, 2085)
                    mcolBill.Add "Ӧ�պϼ�", ctlTemp, , blnVisible
               Case "ʵ�պϼ�"
                    If blnVisible = True Then
                        Set ctlTemp = LoadControl("Label", 3000, 6150)
                        mcolBill.Add "��ǩ_" & ctlTemp.Index, ctlTemp, , True
                        ctlTemp.Caption = "ʵ��"
                    End If
                    Set ctlTemp = LoadControl("TextBox", 3390, 6105, 2085)
                    mcolBill.Add "ʵ�պϼ�", ctlTemp, , blnVisible
            End Select
        Next
        '�շ���Ŀ
        With msh�շ���Ŀ
            For lngCount = 1 To ud��Ŀ��.Value
                '�ӵ�2�п�ʼ
                For lngRow = 2 To .Rows - 1
                    strTemp = .TextMatrix(lngRow, 0)
                    blnVisible = .TextMatrix(lngRow, lngCount) = "��"
                    If lngCount = 1 Then
                        If blnVisible = True Or strTemp = "�շ�ϸĿֵ" Then
                            Set ctlTemp = LoadControl("Label", 405, 2100)
                            mcolBill.Add "��ǩ_" & ctlTemp.Index, ctlTemp, , True
                            ctlTemp.Caption = Replace(.TextMatrix(lngRow, 0), "ֵ", "")
                        End If
                    End If
                    
                    
                    Select Case strTemp
                        Case "�շ����"
                            lngLeft = 120
                            lngWidth = 795
                            strControl = "ComboBox"
                        Case "�շ�ϸĿֵ"
                            lngLeft = 915
                            lngWidth = 2115
                            strTemp = "�շ�ϸĿ"
                            strControl = "TextBox"
                            blnVisible = True '�շ���Ŀ�ǿ϶�Ҫ��ʾ��
                        Case "ϸĿѡ��"
                            lngLeft = 3030
                            lngWidth = frmDesign.cmd(0).Width
                            strControl = "CommandButton"
                        Case "���㵥λ"
                            lngLeft = 3435
                            lngWidth = 1035
                            strControl = "TextBox"
                        Case "����"
                            lngLeft = 4470
                            lngWidth = 705
                            strControl = "TextBox"
                        Case "��׼����"
                            lngLeft = 5175
                            lngWidth = 915
                            strControl = "TextBox"
                        Case "Ӧ�ս��"
                            lngLeft = 6090
                            lngWidth = 1425
                            strControl = "TextBox"
                        Case "ʵ�ս��"
                            lngLeft = 7515
                            lngWidth = 1185
                            strControl = "TextBox"
                        Case "ִ�в���"
                            lngLeft = 8700
                            lngWidth = 1485
                            strControl = "ComboBox"
                        Case "���ӱ�־"
                            lngLeft = 10215
                            lngWidth = 1065
                            strControl = "CheckBox"
                    End Select
                    
                    If strTemp <> "����ֵ" Then
                        '��ӿؼ���Ԫ��
                        If lngCount = 1 Then
                            If blnVisible = True Then
                                ctlTemp.Left = lngLeft '��ǩ����߾�
                            End If
                            
                            If strTemp = "ϸĿѡ��" Then
                                Set ctlTemp = LoadControl(strControl, lngLeft, 2430, , , 0)
                            Else
                                Set ctlTemp = LoadControl(strControl, lngLeft, 2430, lngWidth)
                            End If
                        Else
                            Set ctlTemp = LoadControl(strControl, lngLeft, _
                                mcolBill(strTemp & "_" & lngCount - 1).Control.Top + mcolBill(strTemp & "_" & lngCount - 1).Control.Height, _
                                mcolBill(strTemp & "_" & lngCount - 1).Control.Width)
                        End If
                        
                        If strTemp = "���ӱ�־" Then
                            ctlTemp.Caption = "��������"
                        End If
                        
                        mcolBill.Add strTemp & "_" & lngCount, ctlTemp, , blnVisible
                    End If
                    'Ԥ��ֵ
                    If strTemp = "�շ����" Then
                        mcolBill(strTemp & "_" & lngCount).Value = Mid(.TextMatrix(1, lngCount), 1, 1)
                    ElseIf strTemp = "�շ�ϸĿ" Then
                        mcolBill(strTemp & "_" & lngCount).Value = .ColData(lngCount)
                    ElseIf strTemp = "����" Then
                        mcolBill(strTemp & "_" & lngCount).Value = .TextMatrix(row����ֵ, lngCount)
                    End If
                Next
            Next
        End With
        
    Else
        '�����пؼ������޸�
        
        '�����ѡȡ��Ŀ
        For lngCount = 0 To lst��Ŀ.ListCount - 1
            mcolBill(lst��Ŀ.List(lngCount)).Visible = lst��Ŀ.Selected(lngCount)
        Next
        '�����շ�ϸĿ
        
        arrItems = Split("�շ����,�շ�ϸĿ,ϸĿѡ��,���㵥λ,����,��׼����,Ӧ�ս��,ʵ�ս��,ִ�в���,���ӱ�־", ",")
        With msh�շ���Ŀ
            If mlngItemCount > ud��Ŀ��.Value Then
                '���ڱ���ǰ�����ˣ�Ҫɾ��һЩ
                For lngCount = ud��Ŀ��.Value + 1 To mlngItemCount
                    For lngRow = LBound(arrItems) To UBound(arrItems)
                        strTemp = arrItems(lngRow) & "_" & lngCount
                        '��ɾ��֮ǰ��Ҫжװ�ؼ�
                        Set ctlTemp = mcolBill(strTemp).Control
                        Unload ctlTemp
                        mcolBill.Remove strTemp
                    Next
                Next
            Else
                For lngCount = mlngItemCount + 1 To ud��Ŀ��.Value
                    For lngRow = LBound(arrItems) To UBound(arrItems)
                        strTemp = arrItems(lngRow)
                        Select Case strTemp
                            Case "�շ����"
                                blnVisible = .TextMatrix(row�շ����, lngCount) = "��"
                                Set ctlTemp = LoadControl("ComboBox", mcolBill(strTemp & "_" & lngCount - 1).Control.Left, mcolBill(strTemp & "_" & lngCount - 1).Control.Top + 300, 795)
                                mcolBill.Add strTemp & "_" & lngCount, ctlTemp, , blnVisible
                                mcolBill(strTemp & "_" & lngCount).Value = Mid(.TextMatrix(row�շ����ֵ, lngCount), 1, 1)
                            Case "�շ�ϸĿ"
                                Set ctlTemp = LoadControl("TextBox", mcolBill(strTemp & "_" & lngCount - 1).Control.Left, mcolBill(strTemp & "_" & lngCount - 1).Control.Top + 300, 2115)
                                mcolBill.Add strTemp & "_" & lngCount, ctlTemp
                                mcolBill(strTemp & "_" & lngCount).Value = Mid(.TextMatrix(1, lngRow), 1, 1)
                                mcolBill(strTemp & "_" & lngCount).Value = .ColData(lngCount)
                            Case "ϸĿѡ��"
                                blnVisible = .TextMatrix(rowϸĿѡ��, lngCount) = "��"
                                Set ctlTemp = LoadControl("CommandButton", mcolBill(strTemp & "_" & lngCount - 1).Control.Left, mcolBill("ϸĿѡ��" & "_" & lngCount - 1).Control.Top + 300)
                                mcolBill.Add strTemp & "_" & lngCount, ctlTemp, , blnVisible
                            Case "���㵥λ"
                                blnVisible = .TextMatrix(row���㵥λ, lngCount) = "��"
                                Set ctlTemp = LoadControl("TextBox", mcolBill(strTemp & "_" & lngCount - 1).Control.Left, mcolBill(strTemp & "_" & lngCount - 1).Control.Top + 300, 1035)
                                mcolBill.Add strTemp & "_" & lngCount, ctlTemp, , blnVisible
                            Case "����"
                                blnVisible = .TextMatrix(row����, lngCount) = "��"
                                Set ctlTemp = LoadControl("TextBox", mcolBill(strTemp & "_" & lngCount - 1).Control.Left, mcolBill(strTemp & "_" & lngCount - 1).Control.Top + 300, 705)
                                mcolBill.Add strTemp & "_" & lngCount, ctlTemp, , blnVisible
                                mcolBill(strTemp & "_" & lngCount).Value = .TextMatrix(row����ֵ, lngCount)
                            Case "��׼����"
                                blnVisible = .TextMatrix(row��׼����, lngCount) = "��"
                                Set ctlTemp = LoadControl("TextBox", mcolBill(strTemp & "_" & lngCount - 1).Control.Left, mcolBill(strTemp & "_" & lngCount - 1).Control.Top + 300, 915)
                                mcolBill.Add strTemp & "_" & lngCount, ctlTemp, , blnVisible
                            Case "Ӧ�ս��"
                                blnVisible = .TextMatrix(rowӦ�ս��, lngCount) = "��"
                                Set ctlTemp = LoadControl("TextBox", mcolBill(strTemp & "_" & lngCount - 1).Control.Left, mcolBill(strTemp & "_" & lngCount - 1).Control.Top + 300, 1425)
                                mcolBill.Add strTemp & "_" & lngCount, ctlTemp, , blnVisible
                            Case "ʵ�ս��"
                                blnVisible = .TextMatrix(rowʵ�ս��, lngCount) = "��"
                                Set ctlTemp = LoadControl("TextBox", mcolBill(strTemp & "_" & lngCount - 1).Control.Left, mcolBill(strTemp & "_" & lngCount - 1).Control.Top + 300, 1185)
                                mcolBill.Add strTemp & "_" & lngCount, ctlTemp, , blnVisible
                            Case "ִ�в���"
                                blnVisible = .TextMatrix(rowִ�в���, lngCount) = "��"
                                Set ctlTemp = LoadControl("ComboBox", mcolBill(strTemp & "_" & lngCount - 1).Control.Left, mcolBill(strTemp & "_" & lngCount - 1).Control.Top + 300, 1485)
                                mcolBill.Add strTemp & "_" & lngCount, ctlTemp, , blnVisible
                            Case "���ӱ�־"
                                blnVisible = .TextMatrix(row���ӱ�־, lngCount) = "��"
                                Set ctlTemp = LoadControl("CheckBox", mcolBill(strTemp & "_" & lngCount - 1).Control.Left, mcolBill(strTemp & "_" & lngCount - 1).Control.Top + 300)
                                ctlTemp.Caption = "��������"
                                mcolBill.Add strTemp & "_" & lngCount, ctlTemp, , blnVisible
                        End Select
                    Next
                Next
            End If
            '����������
            For lngCount = 1 To IIf(mlngItemCount < ud��Ŀ��.Value, mlngItemCount, ud��Ŀ��.Value) 'ȡ����С��
                For lngRow = LBound(arrItems) To UBound(arrItems)
                    strTemp = arrItems(lngRow)
                    Select Case strTemp
                        Case "�շ����"
                            blnVisible = .TextMatrix(row�շ����, lngCount) = "��"
                            mcolBill(strTemp & "_" & lngCount).Visible = blnVisible
                            mcolBill(strTemp & "_" & lngCount).Value = Mid(.TextMatrix(row�շ����ֵ, lngCount), 1, 1)
                        Case "�շ�ϸĿ"
                            mcolBill(strTemp & "_" & lngCount).Value = .ColData(lngCount)
                        Case Else
                            blnVisible = .TextMatrix(lngRow + 3, lngCount) = "��"
                            mcolBill(strTemp & "_" & lngCount).Visible = blnVisible
                            
                            If strTemp = "����" Then
                                mcolBill(strTemp & "_" & lngCount).Value = .TextMatrix(row����ֵ, lngCount)
                            End If
                    End Select
                Next
            Next
        End With
    End If
    
    '���濪�����ҵ�ֵ
    If cmb����.ListIndex >= 0 And cmb����.ListIndex < 3 Then
        If cmb����.ListIndex = 0 Or cmb����.ListIndex = 3 Then
            mcolBill("��������").Value = ""
        Else
            mcolBill("��������").Value = "C" & cmb����.ListIndex
        End If
    Else
        mcolBill("��������").Value = cmb����.ItemData(cmb����.ListIndex)
    End If
    
End Sub

Private Function LoadControl(ByVal ControlType As String, ByVal Left As Single, ByVal Top As Single, _
    Optional ByVal Width As Single, Optional ByVal Height As Single, Optional ByVal Index As Long = -1) As Control
    
    Dim ctl As Control
    'װ�ؿؼ�
    Select Case ControlType
        Case "ComboBox"
            If Index = -1 Then
               Load frmDesign.cmb(frmDesign.cmb.UBound + 1)
                Set ctl = frmDesign.cmb(frmDesign.cmb.UBound)
            Else
                Set ctl = frmDesign.cmb(Index)
            End If
        Case "CommandButton"
            If Index = -1 Then
               Load frmDesign.cmd(frmDesign.cmd.UBound + 1)
                Set ctl = frmDesign.cmd(frmDesign.cmd.UBound)
            Else
                Set ctl = frmDesign.cmd(Index)
            End If
        Case "CheckBox"
            If Index = -1 Then
               Load frmDesign.chk(frmDesign.chk.UBound + 1)
                Set ctl = frmDesign.chk(frmDesign.chk.UBound)
            Else
                Set ctl = frmDesign.chk(Index)
            End If
        Case "Label"
            If Index = -1 Then
               Load frmDesign.lbl(frmDesign.lbl.UBound + 1)
                Set ctl = frmDesign.lbl(frmDesign.lbl.UBound)
            Else
                Set ctl = frmDesign.lbl(Index)
            End If
        Case "TextBox"
            If Index = -1 Then
               Load frmDesign.txt(frmDesign.txt.UBound + 1)
                Set ctl = frmDesign.txt(frmDesign.txt.UBound)
            Else
                Set ctl = frmDesign.txt(Index)
            End If
    End Select
    '��������
    Set ctl.Container = frmDesign.picForm
    '����λ��
    ctl.Left = Left
    ctl.Top = Top
    If Width > 0 Then
        ctl.Width = Width
    End If
    If Height > 0 And ControlType <> "ComboBox" Then
        ctl.Height = Height
    End If
    '�����ؼ�����������������ͬ
    SetFont ctl, frmDesign.picForm

    Set LoadControl = ctl
End Function

Private Sub InitCont()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    

    '���ÿ�������
    cmb����.Clear
    cmb����.AddItem "δָ��"
    cmb����.AddItem "�������ڿ���"
    cmb����.AddItem "����Ա���ڿ���"
    cmb����.AddItem "����������������������������"
    
    Set rsTmp = GetDepartments("'�ٴ�','����'", "1,2,3")
    Do Until rsTmp.EOF
        cmb����.AddItem rsTmp("����") & "-" & rsTmp("����")
        cmb����.ItemData(cmb����.NewIndex) = rsTmp("ID")
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    '���ÿ�ѡ��Ŀ�б�
    lst��Ŀ.Clear
    lst��Ŀ.AddItem "����ID"
    lst��Ŀ.AddItem "��ʶ��"
    lst��Ŀ.AddItem "��Ժ����"
    lst��Ŀ.AddItem "�Ա�"
    lst��Ŀ.AddItem "����"
    lst��Ŀ.AddItem "����"
    lst��Ŀ.AddItem "���˲���"
    lst��Ŀ.AddItem "���˿���"
    lst��Ŀ.AddItem "�ѱ�"
    lst��Ŀ.AddItem "�Ӱ��־"
    lst��Ŀ.AddItem "Ӥ����"
    lst��Ŀ.AddItem "Ӧ�պϼ�"
    lst��Ŀ.AddItem "ʵ�պϼ�"
    '�����շ���Ŀ���
    With msh�շ���Ŀ
        .Rows = 13
        .ColWidth(0) = 1300
        .ColWidth(1) = 1000
        .ColAlignmentFixed(0) = 1
        .ColAlignment(1) = 1
        .Row = 0: .Col = 0: .Text = "��Ŀ": .CellAlignment = 4
        .Row = 0: .Col = 1: .Text = "����(1)": .CellAlignment = 4
        .Col = 1
        .TextMatrix(row�շ����ֵ, 0) = "�շ����ֵ"
        .TextMatrix(row�շ�ϸĿֵ, 0) = "�շ�ϸĿֵ"
        .TextMatrix(row����ֵ, 0) = "����ֵ"
        .TextMatrix(row�շ����, 0) = "�շ����"
        .TextMatrix(rowϸĿѡ��, 0) = "ϸĿѡ��"
        .TextMatrix(row���㵥λ, 0) = "���㵥λ"
        .TextMatrix(row����, 0) = "����"
        .TextMatrix(row��׼����, 0) = "��׼����"
        .TextMatrix(rowӦ�ս��, 0) = "Ӧ�ս��"
        .TextMatrix(rowʵ�ս��, 0) = "ʵ�ս��"
        .TextMatrix(rowִ�в���, 0) = "ִ�в���"
        .TextMatrix(row���ӱ�־, 0) = "���ӱ�־"
    End With
    
    strSQL = "Select ����,���� as ��� From �շ���Ŀ��� Where ���� Not In('1','4','5','6','7') Order by ���"
    Set rsTmp = zlDatabase.OpensqlRecord(strSQL, Me.Caption)
    
    cmb���.Clear
    cmb���.AddItem "0-δָ��"
    Do Until rsTmp.EOF
        cmb���.AddItem rsTmp("����") & "-" & rsTmp("���")
        rsTmp.MoveNext
    Loop
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdSelect_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim str��� As String, lng��ĿID As Long
    Dim strSQL As String
    
    str��� = Left(msh�շ���Ŀ.TextMatrix(row�շ����ֵ, msh�շ���Ŀ.Col), 1)
    If str��� = "0" Then str��� = ""
    If str��� <> "" Then str��� = "'" & str��� & "'"
    
    lng��ĿID = frmItemSelect.ShowSelect(Me, gstrPrivs, 0, 0, str���)
    If lng��ĿID <> 0 Then
        strSQL = "Select A.ID,A.���||'-'||B.���� as ���,A.���� From �շ���ĿĿ¼ A,�շ���Ŀ��� B Where A.���=B.���� And A.ID=[1]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpensqlRecord(strSQL, Me.Caption, lng��ĿID)
        msh�շ���Ŀ.Text = rsTmp!����
        msh�շ���Ŀ.ColData(msh�շ���Ŀ.Col) = rsTmp!ID
        msh�շ���Ŀ.TextMatrix(row�շ����ֵ, msh�շ���Ŀ.Col) = rsTmp!���
    End If
    msh�շ���Ŀ.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    If Not InDesign Then
        glngOldProc = GetWindowLong(Me.hwnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf wndProc)
    End If
End Sub

Private Sub Form_Resize()
    cmdHelp.Top = ScaleHeight - cmdHelp.Height - 200
    cmdOK.Top = cmdHelp.Top
    cmdCancel.Top = cmdHelp.Top
    
    cmdCancel.Left = ScaleWidth - cmdCancel.Width - 200
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 100
    
    fra��ѡ.Height = cmdHelp.Top - fra��ѡ.Top - 100
    lst��Ŀ.Height = fra��ѡ.Height - 300
    msh�շ���Ŀ.Height = cmdOK.Top - msh�շ���Ŀ.Top - 100
    msh�շ���Ŀ.Width = ScaleWidth - msh�շ���Ŀ.Left - 60
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = True Then
        If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    If Not InDesign Then
        Call SetWindowLong(Me.hwnd, GWL_WNDPROC, glngOldProc)
    End If
End Sub

Private Sub lst��Ŀ_ItemCheck(Item As Integer)
    mblnChange = True
End Sub

Private Sub msh�շ���Ŀ_DblClick()
    Dim lngRow As Long, lngCol As Long
        
    With msh�շ���Ŀ
        lngRow = .Row
        lngCol = .Col
        If lngCol < 1 Or lngCol > .Cols - 1 Then Exit Sub
        If lngRow < 2 Or lngRow > .Rows - 1 Then Exit Sub
        msh�շ���Ŀ_KeyPress vbKeySpace
    End With
End Sub

Private Sub msh�շ���Ŀ_GotFocus()
    If msh�շ���Ŀ.Row = 1 Then
        cmb���.Visible = True
        cmb���.SetFocus
    End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If InputϸĿ() = True Then
            txtEdit.Visible = False
            KeyAscii = 0
        End If
    ElseIf KeyAscii = vbKeyEscape Then
        txtEdit.Visible = False
        msh�շ���Ŀ.SetFocus
    End If
    
End Sub

Private Sub txtEdit_Validate(Cancel As Boolean)
    If txtEdit.Visible = False Then Exit Sub
    If txtEdit.Text = "" Then
        txtEdit.Visible = False
    Else
        If InputϸĿ = False Then
            Beep
            Cancel = True
        Else
            txtEdit.Visible = False
        End If
    End If
End Sub

Private Function InputϸĿ() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim str��� As String, strSQL As String
    Dim str���� As String, lng��ĿID As Long
    
    str���� = UCase(Replace(txtEdit.Text, "'", "''"))
    str��� = Left(msh�շ���Ŀ.TextMatrix(row�շ����ֵ, msh�շ���Ŀ.Col), 1)
    If str��� = "0" Then str��� = ""
    If str��� <> "" Then str��� = "'" & str��� & "'"
    
    lng��ĿID = frmItemSelect.ShowSelect(Me, gstrPrivs, 0, str���, str����, txtEdit.hwnd)
    If lng��ĿID <> 0 Then
        strSQL = "Select ID,���,���� From �շ���ĿĿ¼ Where ID=[1]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpensqlRecord(strSQL, Me.Caption, lng��ĿID)
        With msh�շ���Ŀ
            .ColData(.Col) = rsTmp!ID
            .TextMatrix(row�շ�ϸĿֵ, .Col) = rsTmp!����
            .TextMatrix(row�շ����ֵ, .Col) = GetClassName(rsTmp!���)
            .Row = row����ֵ
        End With
        mblnChange = True
        InputϸĿ = True
    Else
        zlControl.TxtSelAll txtEdit
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Input���� = True Then
            KeyAscii = 0
        End If
    ElseIf KeyAscii = vbKeyEscape Then
        txt����.Visible = False
        msh�շ���Ŀ.SetFocus
    ElseIf InStr("0123456789" & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    If txt����.Visible = False Then Exit Sub
    
    If txt����.Text = "" Then
        txt����.Visible = False
    Else
        If Input���� = False Then
            Beep
            Cancel = True
        Else
            txt����.Visible = False
        End If
    End If
End Sub

Private Function Input����() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strClass As String
    
    Dim strEdit As String
    
    
    strEdit = UCase(Replace(txt����.Text, "'", "''"))
    If IsNumeric(strEdit) = False Then
        MsgBox "������Ϸ�����ֵ��", vbExclamation, gstrSysName
        txt����.Text = ""
        Exit Function
    End If
    If Val(strEdit) < 0 Or Val(strEdit) > 1000 Then
        MsgBox "������1000���ڵ�����", vbExclamation, gstrSysName
        zlControl.TxtSelAll txt����
        Exit Function
    End If
    With msh�շ���Ŀ
        .TextMatrix(row����ֵ, .Col) = Format(strEdit, "0")
        If Val(strEdit) = 0 Then
            .TextMatrix(row����, .Col) = "��"
        End If
        .Row = row�շ����
    End With
    txt����.Visible = False
    mblnChange = True
    Input���� = True
End Function

Private Sub txt����_LostFocus()
    txt����.Visible = False
End Sub

Private Sub txt��Ŀ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cmb����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub lst��Ŀ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub msh�շ���Ŀ_KeyPress(KeyAscii As Integer)
    With msh�շ���Ŀ
        Select Case KeyAscii
            Case vbKeyReturn
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                Else
                    If .Col = .Cols - 1 Then
                        SendKeys "{TAB}"
                    Else
                        .Row = 1: .Col = .Col + 1
                    End If
                End If
                Call msh�շ���Ŀ_EnterCell
            Case vbKeySpace
                If .Row > row����ֵ Then
                    If .Row = row�շ���� Then
                        If .TextMatrix(row�շ����ֵ, .Col) = "" Or Mid(.TextMatrix(row�շ����ֵ, .Col), 1, 1) = "0" Then
                            .TextMatrix(row�շ����, .Col) = "��"
                            Exit Sub
                        End If
                    ElseIf .Row = row���� Then
                        If Val(.TextMatrix(row����ֵ, .Col)) = 0 Then
                            .TextMatrix(row����, .Col) = "��"
                            Exit Sub
                        End If
                    End If
                    .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = "��", "", "��")
                    mblnChange = True
                ElseIf .Row = row�շ�ϸĿֵ Then
                    txtEdit.Text = .Text
                    zlControl.TxtSelAll txtEdit
                    Call ShowTxtEdit
                ElseIf .Row = row����ֵ Then
                    txt����.Text = .Text
                    zlControl.TxtSelAll txt����
                    Call Show����
                End If
            Case Asc("*")
                If .Row = row�շ�ϸĿֵ Then
                    Call cmdSelect_Click
                End If
            Case Else
                If .Row = row�շ�ϸĿֵ Then
                    txtEdit.Text = Chr(KeyAscii)
                    txtEdit.SelStart = Len(txtEdit.Text)
                    Call ShowTxtEdit
                ElseIf .Row = row����ֵ And InStr("0123456789", Chr(KeyAscii)) > 0 Then
                    txt����.Text = Chr(KeyAscii)
                    txt����.SelStart = Len(txt����.Text)
                    Call Show����
                End If
        End Select
    End With
End Sub

Private Sub msh�շ���Ŀ_Scroll()
    cmdSelect.Visible = False
    cmb���.Visible = False
    If msh�շ���Ŀ.Row = 1 Then msh�շ���Ŀ.Row = 2
'    Call msh�շ���Ŀ_EnterCell
End Sub

Private Sub msh�շ���Ŀ_EnterCell()
    Dim lngCount As Long
    
    cmb���.Visible = False
    cmdSelect.Visible = False
    
    With msh�շ���Ŀ
        If .Row = row�շ����ֵ Then
            cmb���.Left = .Left + .CellLeft
            If .Row = row�շ�ϸĿֵ Then
                '�����˹������������� msh�շ���Ŀ_Scroll �¼�
                cmdSelect.Left = .Left + .CellLeft + GetCellWidth - cmdSelect.Width
                cmdSelect.Visible = True
                Exit Sub
            End If
            cmb���.Width = GetCellWidth()
            
            For lngCount = 0 To cmb���.ListCount - 1
                If cmb���.List(lngCount) = .Text Then
                    cmb���.ListIndex = lngCount
                    Exit For
                End If
            Next
            If lngCount = cmb���.ListCount Then
                cmb���.ListIndex = 0
            End If
            cmb���.Visible = True
            If cmb���.Visible = True Then cmb���.SetFocus
        ElseIf .Row = row�շ�ϸĿֵ Then
            cmdSelect.Left = .Left + .CellLeft + GetCellWidth - cmdSelect.Width
            cmdSelect.Visible = True
        End If
    End With
End Sub

Private Sub ShowTxtEdit()
    With msh�շ���Ŀ
        cmdSelect.Visible = False
        txt����.Visible = False
        txtEdit.Left = .Left + .CellLeft + 30
        txtEdit.Width = GetCellWidth() - 30
        txtEdit.Top = .Top + .CellTop + 45
        txtEdit.Visible = True
        txtEdit.SetFocus
    End With
End Sub

Private Sub Show����()
    With msh�շ���Ŀ
        cmdSelect.Visible = False
        txtEdit.Visible = False
        txt����.Left = .Left + .CellLeft + 30
        txt����.Width = GetCellWidth() - 30
        txt����.Top = .Top + .CellTop + 45
        txt����.ZOrder
        txt����.Visible = True
        txt����.SetFocus
    End With
End Sub

Private Function GetCellWidth() As Long
'�õ���ǰ��Ԫ����ʾ�����Ŀ��
    With msh�շ���Ŀ
        If .CellLeft + .CellWidth > .Width Then
            '��������������
            GetCellWidth = .Width - .CellLeft - 30
        Else
            GetCellWidth = .CellWidth - 30
        End If
    End With
    If GetCellWidth < 0 Then GetCellWidth = 0
End Function

Private Sub ud��Ŀ��_Change()
    Dim lngRow As Long
    
    With msh�շ���Ŀ
        If .Cols < ud��Ŀ��.Value + 1 Then
            .Cols = ud��Ŀ��.Value + 1
            
            For lngRow = 1 To .Cols - 1
                .Row = 0: .Col = lngRow: .Text = "����(" & lngRow & ")": .CellAlignment = 4
                .ColAlignment(.Col) = 1
                .ColWidth(.Col) = .ColWidth(.Col - 1)
            Next
            .Row = 1
            .TextMatrix(row�շ����ֵ, ud��Ŀ��) = .TextMatrix(row�շ����ֵ, ud��Ŀ�� - 1)
            For lngRow = row����ֵ To .Rows - 1
                .TextMatrix(lngRow, ud��Ŀ��) = .TextMatrix(lngRow, ud��Ŀ�� - 1)
            Next
        Else
            .Cols = ud��Ŀ��.Value + 1
        End If
    End With
    mblnChange = True
    Call msh�շ���Ŀ_EnterCell
End Sub

Private Function GetClassName(ByVal str���� As String) As String
'���������룬�õ�����ȫ��
    Dim lngCount As Long
    For lngCount = 1 To cmb���.ListCount - 1
        If Mid(cmb���.List(lngCount), 1, 1) = str���� Then
            GetClassName = cmb���.List(lngCount)
            Exit Function
        End If
    Next
    GetClassName = cmb���.List(0)
End Function

Private Function GetItemName(ByVal strID As String) As String
'���������룬�õ�����ȫ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ���� From �շ���ĿĿ¼ Where ID=[1]"
    Set rsTmp = zlDatabase.OpensqlRecord(strSQL, Me.Caption, strID)
    If Not rsTmp.EOF Then GetItemName = rsTmp!����
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
