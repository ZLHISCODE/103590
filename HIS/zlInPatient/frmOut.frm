VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���˳�Ժ"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8520
   Icon            =   "frmOut.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleMode       =   0  'User
   ScaleWidth      =   6712.303
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7185
      TabIndex        =   18
      Top             =   795
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7185
      TabIndex        =   17
      Top             =   360
      Width           =   1100
   End
   Begin VB.Frame fraInfo 
      Height          =   5535
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   7035
      Begin VB.ComboBox cbo��Ժ��� 
         Height          =   300
         Left            =   5790
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1020
         Width           =   1110
      End
      Begin VB.CheckBox chk���� 
         Alignment       =   1  'Right Justify
         Caption         =   "����"
         Height          =   195
         Left            =   2955
         TabIndex        =   13
         Top             =   5130
         Width           =   660
      End
      Begin VB.TextBox txt���� 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4290
         MaxLength       =   3
         TabIndex        =   14
         Top             =   5070
         Width           =   525
      End
      Begin VB.CheckBox chkʬ�� 
         Alignment       =   1  'Right Justify
         Caption         =   "ʬ��"
         Enabled         =   0   'False
         Height          =   195
         Left            =   5910
         TabIndex        =   12
         Top             =   4740
         Width           =   660
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   630
         Width           =   2055
      End
      Begin VB.TextBox txtסԺ�� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   630
         Width           =   2130
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5685
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txt�Ա� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   1275
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   2130
      End
      Begin VB.TextBox txt��Ժ��� 
         Height          =   300
         Left            =   960
         MaxLength       =   200
         TabIndex        =   2
         Top             =   1020
         Width           =   3900
      End
      Begin VB.CheckBox chk���� 
         Alignment       =   1  'Right Justify
         Caption         =   "ȷ��"
         Height          =   195
         Left            =   2955
         TabIndex        =   9
         Top             =   4740
         Width           =   660
      End
      Begin VB.TextBox txt��ҽ��� 
         Height          =   300
         Left            =   960
         MaxLength       =   200
         TabIndex        =   5
         Top             =   2820
         Width           =   3900
      End
      Begin VB.ComboBox cbo��ҽ��Ժ��� 
         Height          =   300
         Left            =   5790
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2820
         Width           =   1110
      End
      Begin VB.ComboBox cbo��Ժ��ʽ 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   4680
         Width           =   1830
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         ItemData        =   "frmOut.frx":030A
         Left            =   5040
         List            =   "frmOut.frx":031D
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   5070
         Width           =   1095
      End
      Begin MSComCtl2.UpDown UD���� 
         Height          =   300
         Left            =   4800
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   5070
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txt����"
         BuddyDispid     =   196614
         OrigLeft        =   3945
         OrigTop         =   645
         OrigRight       =   4185
         OrigBottom      =   930
         Max             =   99999
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   300
         Left            =   960
         TabIndex        =   11
         Top             =   5070
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   19
         Format          =   "yyyy-MM-dd hh:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VSFlex8Ctl.VSFlexGrid vfg��ҽ 
         Height          =   1335
         Left            =   960
         TabIndex        =   4
         Top             =   1400
         Width           =   5895
         _cx             =   10398
         _cy             =   2355
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16635590
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   280
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vfg��ҽ 
         Height          =   1335
         Left            =   960
         TabIndex        =   7
         Top             =   3200
         Width           =   5895
         _cx             =   10398
         _cy             =   2355
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16635590
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   280
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin MSMask.MaskEdBox txtOkDate 
         Height          =   300
         Left            =   3720
         TabIndex        =   10
         Top             =   4680
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   19
         Format          =   "yyyy-MM-dd hh:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label lbl��ҽ���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   180
         TabIndex        =   37
         Top             =   3240
         Width           =   720
      End
      Begin VB.Label lbl��ҽ���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   180
         TabIndex        =   36
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ���"
         Height          =   180
         Left            =   4995
         TabIndex        =   35
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժʱ��"
         Height          =   180
         Left            =   180
         TabIndex        =   34
         Top             =   5130
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   3900
         TabIndex        =   33
         Top             =   5130
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ��λ"
         Height          =   180
         Left            =   3975
         TabIndex        =   32
         Top             =   690
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   5205
         TabIndex        =   31
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   3345
         TabIndex        =   30
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   525
         TabIndex        =   29
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "סԺ��"
         Height          =   180
         Left            =   360
         TabIndex        =   28
         Top             =   690
         Width           =   540
      End
      Begin VB.Label lbl��Ժ��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ���"
         Height          =   180
         Left            =   180
         TabIndex        =   27
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label lbl��ҽ��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ҽ���"
         Height          =   180
         Left            =   180
         TabIndex        =   26
         Top             =   2880
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ���"
         Height          =   180
         Left            =   4995
         TabIndex        =   25
         Top             =   2880
         Width           =   720
      End
      Begin VB.Label lbl��Ժ��ʽ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ��ʽ"
         Height          =   180
         Left            =   180
         TabIndex        =   24
         Top             =   4740
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   7185
      TabIndex        =   0
      Top             =   4950
      Width           =   1100
   End
End
Attribute VB_Name = "frmOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Public mstrPrivs As String
Public mlng����ID As Long, mlng��ҳID As Long

Private mrsPatiInfo As ADODB.Recordset
Private mintĬ����� As Integer
Private mblnOutDeath As Boolean
Private mstrOldName As String
Private mdteDeathDate As Date
Private mintDeath As Integer
Private mfrmParent As Object

Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML

Private Sub cbo��Ժ��ʽ_Click()
    Dim i As Integer
    If InStr(cbo��Ժ��ʽ.Text, "����") > 0 Then
        txt����.Text = ""
        chk����.Value = 0

        txt����.Enabled = (chk����.Value = 1)
        UD����.Enabled = txt����.Enabled

        chk����.Enabled = False
    
        chkʬ��.Enabled = True
    Else
        chk����.Enabled = True
        
        chkʬ��.Value = 0
        chkʬ��.Enabled = False
        If mrsPatiInfo Is Nothing Then Exit Sub
        If mrsPatiInfo.RecordCount = 0 Then Exit Sub
         '49163,������,2012-09-07,���������־����������
        If Not IsNull(mrsPatiInfo!��������) Then
            chk����.Value = 1
            txt����.Text = Nvl(mrsPatiInfo!��������)
            i = cbo.FindIndex(cbo����, Decode(Val(Nvl(mrsPatiInfo!�����־)), 1, "��", 2, "��", 3, "��", 4, "��", 9, "����", ""), True)
            If i <> -1 Then cbo����.ListIndex = i
            
            txt����.Enabled = False
            UD����.Enabled = False
            chk����.Enabled = False
            cbo����.Enabled = False
        End If
    End If
End Sub

Private Sub cbo��Ժ���_Click()
    Dim i As Integer
    If InStr(cbo��Ժ���.Text, "����") > 0 Then
        i = cbo.FindIndex(cbo��Ժ��ʽ, "����", True)
        If i <> -1 Then cbo��Ժ��ʽ.ListIndex = i
    End If
End Sub

Private Sub cbo����_Click()
    txt����.Enabled = (cbo����.ItemData(cbo����.ListIndex) <> 9)
    UD����.Enabled = txt����.Enabled
End Sub

Private Sub cbo��ҽ��Ժ���_Click()
    Dim i As Integer
    If InStr(cbo��ҽ��Ժ���.Text, "����") > 0 Then
        i = cbo.FindIndex(cbo��Ժ��ʽ, "����", True)
        If i <> -1 Then cbo��Ժ��ʽ.ListIndex = i
    End If
End Sub

Private Sub chk����_Click()
    txt����.Enabled = (chk����.Value = 1)
    UD����.Enabled = txt����.Enabled
    cbo����.Enabled = txt����.Enabled
    Call zlCommFun.PressKey(vbKeyTab)
End Sub
'����28982 by lesfeng 2010-06-09
Private Sub chk����_Click()
    If chk����.Value = 1 Then
        txtOkDate.Enabled = True
    Else
        txtOkDate.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '����28139 by lesfeng 2010-03-02 ����������ϴ���
    If KeyCode = 13 Then
        If Not Me.ActiveControl Is txt��Ժ��� _
            And Not Me.ActiveControl Is txt��ҽ��� And Not Me.ActiveControl Is vfg��ҽ And Not Me.ActiveControl Is vfg��ҽ Then
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '����28139 by lesfeng 2010-03-02 ����������ϴ���
    If KeyAscii = Asc("'") And Not (Me.ActiveControl Is txt��Ժ��� Or Me.ActiveControl Is txt��ҽ��� Or Me.ActiveControl Is vfg��ҽ Or Me.ActiveControl Is vfg��ҽ) Then KeyAscii = 0       '��������п�����'��
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, strSQL As String, str���� As String, str����� As String
    Dim dMax As Date, intԭ�� As Integer, int���� As Integer
    Dim rsDiagnosis As ADODB.Recordset
    Dim str��Ժ��� As String, str��ҽ��Ժ��� As String, strTmp As String
    '����28138 by lesfeng 2010-03-01
    mintĬ����� = Val(zlDatabase.GetPara("Ĭ�����", glngSys, glngModul))
    '63706:������,2014-08-11
    mblnOutDeath = (Val(zlDatabase.GetPara("��Ժ����", glngSys, glngModul)) = 1)
    '����28612 by lesfeng 2010-07-05
    mintDeath = 0
    mdteDeathDate = GetdeathTime(mlng����ID, mlng��ҳID)

    gblnOK = False
    Set mrsPatiInfo = GetPatiInfo(mlng����ID, mlng��ҳID)
    int���� = Val("" & mrsPatiInfo!����)
    
    'ҽ�����˳�Ժ���
    If int���� <> 0 Then '�Ƿ�����δ�����Ժ
        If Not gclsInsure.GetCapability(supportδ�����Ժ, mlng����ID, int����) Then
            Set rsTmp = GetMoneyInfo(mlng����ID, , , , , , , mlng��ҳID)
            If Not rsTmp Is Nothing Then
                If Nvl(rsTmp!�������, 0) <> 0 Then
                    MsgBox "�ñ��ղ��˵ķ�����δ����,���Ƚ��ʺ��ٳ�Ժ��", vbInformation, gstrSysName
                    Unload Me: Exit Sub
                End If
            End If
        End If
    End If

    '����������
    strTmp = inBlackList(mlng����ID)
    If strTmp <> "" Then
        If MsgBox("�ò��������ⲡ�������С�" & vbCrLf & vbCrLf & "ԭ��" & vbCrLf & vbCrLf & "����" & strTmp & vbCrLf & vbCrLf & "Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Unload Me: Exit Sub
        End If
    End If
    
    If gblnҽ��������ܳ�Ժ Then
        If Not Checkҽ���´��Ժҽ��(mlng����ID, mlng��ҳID) Then
            MsgBox "ҽ���Բ����´��Ժ(��תԺ������)ҽ����������Ժ��", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
    End If
    
    '���˴�����������ʱ����������Ժ
    If gbln��Ժ���˲�׼��Ժ���� Then
        strTmp = ""
        '56323:������,2013-02-18,��ǿ����Ϊ��˵��ݵ���ʾ��Ϣ����
        strTmp = Check��������(mlng����ID, mlng��ҳID)
        If strTmp <> "" Then
            MsgBox "�ò��˴�������δ��˵��������뵥�ݣ�" & vbCrLf & vbCrLf & strTmp & vbCrLf & vbCrLf & "���ܰ����Ժ��", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
    End If
    
    With mrsPatiInfo
        txt����.Text = !����
        txt�Ա�.Text = "" & !�Ա�
        txt����.Text = "" & !����
        txtסԺ��.Text = "" & !סԺ��
        
        txt��ҽ���.Enabled = (InStr(1, "," & GetDepCharacter(Val("" & !��Ժ����id)) & ",", ",��ҽ��,") > 0)
        txt��ҽ���.ToolTipText = "ֻ�е��������ڿ��ҵ�����Ϊ��ҽ��ʱ������������ҽ���!"
        cbo��ҽ��Ժ���.Enabled = txt��ҽ���.Enabled
    End With
        
    Set rsTmp = GetPatiBeds(mlng����ID)
    str����� = ""
    If rsTmp.RecordCount = 0 Then
        str���� = "��ͥ����"
    Else
        Do While Not rsTmp.EOF
            str���� = str���� & "," & rsTmp!����
            If Nvl(rsTmp!����) = Nvl(mrsPatiInfo!��Ҫ����) And Nvl(rsTmp!����ID) = Nvl(mrsPatiInfo!��ס����id) Then
                str����� = Nvl(rsTmp!�����)
            End If
            rsTmp.MoveNext
        Loop
        str���� = Mid(str����, 2)
    End If
    txt����.Text = str����
    txt����.Tag = str�����
    
    txtDate.Text = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    dMax = GetMaxDate(mlng����ID, mlng��ҳID, intԭ��)
    If intԭ�� = 10 Then
        '59094:������,2013-04-24,�޸�Ϊֻ��1s,ԭ��Ϊ1m
        txtDate.Text = Format(dMax + 1 / 24 / 60 / 60, "yyyy-MM-dd HH:mm:ss")
    Else
        If dMax > CDate(txtDate.Text) Then
            txtDate.Text = Format(dMax + 1 / 24 / 60, "yyyy-MM-dd HH:mm:ss")
        End If
    End If
    '����28612 by lesfeng 2010-07-05
    If mintDeath = 1 Then
        txtDate.Text = Format(mdteDeathDate, "yyyy-MM-dd HH:mm:ss")
    End If
    
    '75308,û�е�����Ժʱ��Ȩ��,�������޸ĳ�Ժʱ��
    If InStr(1, ";" & mstrPrivs & ";", ";������Ժʱ��;") = 0 Then
        txtDate.Enabled = False
    End If
    '��ʾ������ϼ�¼
    Set rsDiagnosis = GetDiagnosticInfo(mlng����ID, mlng��ҳID, "1,11,2,12,3,13", "2,3")
    If Not rsDiagnosis Is Nothing Then
        'a.��ҽ���
        rsDiagnosis.Filter = "�������=3 and ��¼��Դ=3"            '��ȡ��ҳ����ĳ�Ժ���
        If Not rsDiagnosis.EOF Then
            txt��Ժ���.Text = Nvl(rsDiagnosis!�������): txt��Ժ���.Tag = Nvl(rsDiagnosis!����ID, rsDiagnosis!���ID & ";"): lbl��Ժ���.Tag = txt��Ժ���.Text
            str��Ժ��� = "" & rsDiagnosis!��Ժ���
            '����28982 by lesfeng 2010-06-09
            chk����.Value = IIf(Val("" & rsDiagnosis!�Ƿ�����) = 1, 0, 1)
        Else
            '����28483 by lesfeng 2010-03-01
            rsDiagnosis.Filter = "�������=3 and ��¼��Դ=2"        '��ȡ��Ժ�Ǽǵĳ�Ժ���
            If Not rsDiagnosis.EOF Then
                txt��Ժ���.Text = Nvl(rsDiagnosis!�������): txt��Ժ���.Tag = Nvl(rsDiagnosis!����ID, rsDiagnosis!���ID & ";"): lbl��Ժ���.Tag = txt��Ժ���.Text
                str��Ժ��� = "" & rsDiagnosis!��Ժ���
                '����28982 by lesfeng 2010-06-09
                chk����.Value = IIf(Val("" & rsDiagnosis!�Ƿ�����) = 1, 0, 1)
            Else
                '����28138 by lesfeng 2010-03-01 ����Ĭ����ϵ��ж� ����ȡ������ϼ���Ժ���
                If mintĬ����� = 1 Then
                    rsDiagnosis.Filter = "�������=2 and ��¼��Դ=2"        '��ȡ��Ժ�Ǽǵ���Ժ���
                    If Not rsDiagnosis.EOF Then
                        txt��Ժ���.Text = Nvl(rsDiagnosis!�������): txt��Ժ���.Tag = Nvl(rsDiagnosis!����ID, rsDiagnosis!���ID & ";"): lbl��Ժ���.Tag = txt��Ժ���.Text
                    Else
                        rsDiagnosis.Filter = "�������=1 and ��¼��Դ=2"    '���ȡ��Ժ�Ǽǵ��������
                        If Not rsDiagnosis.EOF Then
                            txt��Ժ���.Text = Nvl(rsDiagnosis!�������): txt��Ժ���.Tag = Nvl(rsDiagnosis!����ID, rsDiagnosis!���ID & ";"): lbl��Ժ���.Tag = txt��Ժ���.Text
                        End If
                    End If
                End If
            End If
        End If
        
        'b.��ҽ���
        If txt��ҽ���.Enabled Then
            rsDiagnosis.Filter = "�������=13 and ��¼��Դ=3"            '��ȡ��ҳ����ĳ�Ժ���
            If Not rsDiagnosis.EOF Then
                txt��ҽ���.Text = Nvl(rsDiagnosis!�������): txt��ҽ���.Tag = Nvl(rsDiagnosis!����ID, rsDiagnosis!���ID & ";"): lbl��ҽ���.Tag = txt��ҽ���.Text
                str��ҽ��Ժ��� = "" & rsDiagnosis!��Ժ���
            Else
                '����28483 by lesfeng 2010-03-01
                rsDiagnosis.Filter = "�������=13 and ��¼��Դ=2"        '��ȡ��Ժ�Ǽǵĳ�Ժ���
                If Not rsDiagnosis.EOF Then
                    txt��ҽ���.Text = Nvl(rsDiagnosis!�������): txt��ҽ���.Tag = Nvl(rsDiagnosis!����ID, rsDiagnosis!���ID & ";"): lbl��ҽ���.Tag = txt��ҽ���.Text
                    str��ҽ��Ժ��� = "" & rsDiagnosis!��Ժ���
                Else
                    '����28138 by lesfeng 2010-03-01 ����Ĭ����ϵ��ж� ����ȡ������ϼ���Ժ���
                    If mintĬ����� = 1 Then
                        rsDiagnosis.Filter = "�������=12 and ��¼��Դ=2"        '��ȡ��Ժ�Ǽǵ���Ժ���
                        If Not rsDiagnosis.EOF Then
                            txt��ҽ���.Text = Nvl(rsDiagnosis!�������): txt��ҽ���.Tag = Nvl(rsDiagnosis!����ID, rsDiagnosis!���ID & ";"): lbl��ҽ���.Tag = txt��ҽ���.Text
                        Else
                            rsDiagnosis.Filter = "�������=11 and ��¼��Դ=2"    '���ȡ��Ժ�Ǽǵ��������
                            If Not rsDiagnosis.EOF Then
                                txt��ҽ���.Text = Nvl(rsDiagnosis!�������): txt��ҽ���.Tag = Nvl(rsDiagnosis!����ID, rsDiagnosis!���ID & ";"): lbl��ҽ���.Tag = txt��ҽ���.Text
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    '����28982 by lesfeng 2010-06-09
    If Not IsNull(mrsPatiInfo!ȷ������) Then
        txtOkDate.Text = Format(mrsPatiInfo!ȷ������, "yyyy-MM-dd HH:mm:ss")
        chk����.Value = IIf(Val("" & mrsPatiInfo!�Ƿ�ȷ��) = 1, 1, 0)
        If chk����.Value = 0 Then chk����.Value = 1
        chk����.Enabled = False
        txtOkDate.Enabled = False
    End If
    
    '��Ժ���
    cbo��Ժ���.AddItem "": cbo��Ժ���.ListIndex = cbo��Ժ���.NewIndex
    If cbo��ҽ��Ժ���.Enabled Then cbo��ҽ��Ժ���.AddItem "": cbo��ҽ��Ժ���.ListIndex = cbo��ҽ��Ժ���.NewIndex
    
     On Error GoTo errH
    strSQL = "Select ����,����,Nvl(ȱʡ��־,0) as ȱʡ From ���ƽ�� Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo��Ժ���.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 Then
                If txt��Ժ���.Text <> "" Then cbo��Ժ���.ListIndex = cbo��Ժ���.NewIndex
                cbo��Ժ���.ItemData(cbo��Ժ���.NewIndex) = 1
            End If
            
            If cbo��ҽ��Ժ���.Enabled Then
                cbo��ҽ��Ժ���.AddItem rsTmp!���� & "-" & rsTmp!����
                If rsTmp!ȱʡ = 1 Then
                    If txt��ҽ���.Text <> "" Then cbo��ҽ��Ժ���.ListIndex = cbo��ҽ��Ժ���.NewIndex
                    cbo��ҽ��Ժ���.ItemData(cbo��ҽ��Ժ���.NewIndex) = 1
                End If
            End If
            rsTmp.MoveNext
        Next
    End If
    Call cbo.Locate(cbo��Ժ���, str��Ժ���)
    Call cbo.Locate(cbo��ҽ��Ժ���, str��ҽ��Ժ���)
    
    '��Ժ��ʽ
    strSQL = "Select ����,����,Nvl(ȱʡ��־,0) as ȱʡ From ��Ժ��ʽ Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo��Ժ��ʽ.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 Then cbo��Ժ��ʽ.ListIndex = cbo��Ժ��ʽ.NewIndex
            rsTmp.MoveNext
        Next
    End If
    If (Nvl(mrsPatiInfo!��Ժ��ʽ) <> "") Then
        Call cbo.Locate(cbo��Ժ��ʽ, Nvl(mrsPatiInfo!��Ժ��ʽ))
    End If
    
    '47955:������,2012-09-18,�����������ҽ������Ժ��ʽĬ��ѡ��"����"
    If InStr(cbo��Ժ���.Text, "����") > 0 Or InStr(cbo��ҽ��Ժ���.Text, "����") > 0 Or mintDeath = 1 Then
        i = cbo.FindIndex(cbo��Ժ��ʽ, "����", True)
        If i <> -1 Then cbo��Ժ��ʽ.ListIndex = i
    End If
    chkʬ��.Value = IIf(Val(Nvl(mrsPatiInfo!ʬ���־)) = 1, 1, 0)
    
    cbo����.ListIndex = 0
    '49163,������,2012-09-07,���������־����������
    If InStr(cbo��Ժ��ʽ.Text, "����") = 0 And Not IsNull(mrsPatiInfo!��������) Then
        chk����.Value = 1
        txt����.Text = Nvl(mrsPatiInfo!��������)
        i = cbo.FindIndex(cbo����, Decode(Val(Nvl(mrsPatiInfo!�����־)), 1, "��", 2, "��", 3, "��", 4, "��", 9, "����", ""), True)
        If i <> -1 Then cbo����.ListIndex = i
        
        txt����.Enabled = False
        UD����.Enabled = False
        chk����.Enabled = False
        cbo����.Enabled = False
    End If
    '����28139 by lesfeng 2010-03-02
    Call LoadVfgData(vfg��ҽ, 1)
    Call LoadVfgData(vfg��ҽ, 2)
    If chk����.Enabled = True Then Call chk����_Click
    If chk����.Enabled Then Call chk����_Click
    
    '������Ϣ����
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, 1132, mstrPrivs, mfrmParent.hWnd)
    Call AddMipModule(mclsMipModule)
    Set mclsXML = New zl9ComLib.clsXML
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Dim dMax As Date, Curdate As Date, i As Integer
    Dim strSQL As String, strInfo As String, blnTrans As Boolean
    Dim lng��ҽ����ID As Long, lng��ҽ����ID As Long
    Dim lng��ҽ���ID As Long, lng��ҽ���ID As Long
    Dim int���� As Integer, int���� As Integer
    Dim strTmp As String, str������� As String, str������� As String
    Dim int���� As Integer, int������� As Integer
    Dim int��ϴ��� As Integer
    Dim strICD���� As String
    Dim strȷ������  As String
    Dim str��Ժʱ�� As String
    Dim lngRow As Long, strRow������� As String, strRowICD���� As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    
    '��Ժ���
    If Not CheckLen(txt��Ժ���, txt��Ժ���.MaxLength) Then Exit Sub
    If Not CheckLen(txt��ҽ���, txt��ҽ���.MaxLength) Then Exit Sub
    
    int���� = Val("" & mrsPatiInfo!����)
    If int���� <> 0 Then
        If gclsInsure.GetCapability(support����¼��������, mlng����ID, int����) Then
            If txt��Ժ���.Text = "" Then
                MsgBox "����д�ò��˵ĳ�Ժ��ϣ�", vbInformation, gstrSysName
                txt��Ժ���.SetFocus: Exit Sub
            End If
        End If
    End If
    If txt��Ժ���.Text <> "" And cbo��Ժ���.Text = "" Then
        MsgBox "��ѡ���Ժ��ϵĳ�Ժ�����", vbInformation, gstrSysName
        cbo��Ժ���.SetFocus: Exit Sub
    End If
    If txt��ҽ���.Text <> "" And cbo��ҽ��Ժ���.Text = "" And cbo��ҽ��Ժ���.Enabled Then
        MsgBox "��ѡ����ҽ��Ժ��ϵĳ�Ժ�����", vbInformation, gstrSysName
        cbo��ҽ��Ժ���.SetFocus: Exit Sub
    End If
    
    '����28139 by lesfeng 2010-03-03 �����ж�
    strTmp = Replace(txt��Ժ���.Text, "'", "''")
    With vfg��ҽ
        For int���� = 1 To .Rows - 1
            str������� = Trim(.TextMatrix(int����, .ColIndex("�������")))
            str������� = Trim(.TextMatrix(int����, .ColIndex("��Ժ���")))
            strICD���� = Trim(.TextMatrix(int����, .ColIndex("ICD����")))
            If str������� <> "" And strTmp = "" Then
                MsgBox "����д�ò��˵ĳ�Ժ��ϣ�������д������ϣ�", vbInformation, gstrSysName
                txt��Ժ���.SetFocus: Exit Sub
            End If
            If str������� <> "" And str������� = "" Then
                MsgBox "��ѡ��������Ժ��ϵĳ�Ժ���", vbInformation, gstrSysName
                vfg��ҽ.SetFocus
                .Select int����, .ColIndex("��Ժ���")
                Exit Sub
            End If
            If strICD���� <> "" Then
                str������� = "(" & strICD���� & ")" & str�������
            End If
            If str������� = strTmp And strTmp <> "" Then
                MsgBox "�ò��˵ĳ�Ժ������Ժ�����ͬ��������ٱ��棡", vbInformation, gstrSysName
                .Select int����, .ColIndex("�������")
                Exit Sub
            End If
            '50337:������,2012-09-18,�����������Ƿ��ظ�
            For lngRow = int���� + 1 To .Rows - 1
                strRow������� = Trim(.TextMatrix(lngRow, .ColIndex("�������")))
                strRowICD���� = Trim(.TextMatrix(lngRow, .ColIndex("ICD����")))
                If strRowICD���� <> "" Then
                    strRow������� = "(" & strRowICD���� & ")" & strRow�������
                End If
                If str������� = strRow������� And str������� <> "" And strRow������� <> "" Then
                    MsgBox "�ò��˵ĳ�Ժ��������б��е�" & int���� & "��" & lngRow & "�е������ͬ��������ڱ��棡", vbInformation, gstrSysName
                    .Select lngRow, .ColIndex("�������")
                    Exit Sub
                End If
            Next lngRow
        Next
    End With
    
    If cbo��ҽ��Ժ���.Enabled Then
        strTmp = Replace(txt��ҽ���.Text, "'", "''")
        With vfg��ҽ
            For int���� = 1 To .Rows - 1
                str������� = Trim(.TextMatrix(int����, .ColIndex("�������")))
                str������� = Trim(.TextMatrix(int����, .ColIndex("��Ժ���")))
                strICD���� = Trim(.TextMatrix(int����, .ColIndex("��ҽ����")))
                If str������� <> "" And strTmp = "" Then
                    MsgBox "����д�ò��˵ĳ�Ժ��ϣ�������д������ϣ�", vbInformation, gstrSysName
                    txt��ҽ���.SetFocus: Exit Sub
                End If
                If str������� <> "" And str������� = "" Then
                    MsgBox "��ѡ��������Ժ��ϵĳ�Ժ���", vbInformation, gstrSysName
                    vfg��ҽ.SetFocus
                    .Select int����, .ColIndex("��Ժ���")
                    Exit Sub
                End If
                If strICD���� <> "" Then
                    str������� = "(" & strICD���� & ")" & str�������
                End If
                If str������� = strTmp And strTmp <> "" Then
                    MsgBox "��д�ò��˵ĳ�Ժ�������ҽ�����ͬ��������ٱ��棡", vbInformation, gstrSysName
                    .Select int����, .ColIndex("�������")
                    Exit Sub
                End If
                
                '50337:������,2012-09-18,�����������Ƿ��ظ�
                For lngRow = int���� + 1 To .Rows - 1
                    strRow������� = Trim(.TextMatrix(lngRow, .ColIndex("�������")))
                    strRowICD���� = Trim(.TextMatrix(lngRow, .ColIndex("��ҽ����")))
                    If strRowICD���� <> "" Then
                        strRow������� = "(" & strRowICD���� & ")" & strRow�������
                    End If
                    If str������� = strRow������� And str������� <> "" And strRow������� <> "" Then
                        MsgBox "�ò��˵���ҽ��������б��е�" & int���� & "��" & lngRow & "�е������ͬ��������ڱ��棡", vbInformation, gstrSysName
                        .Select lngRow, .ColIndex("�������")
                        Exit Sub
                    End If
                Next lngRow
            Next
        End With
    End If
    
    If Not IsDate(txtDate.Text) Then
        MsgBox "��������ȷ�Ĳ��˳�Ժʱ�䣡", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    
    'ʱ�䲻�ܳ�����ǰʱ��̫��(һ��)
    Curdate = zlDatabase.Currentdate
    If CDate(txtDate.Text) > Curdate Then
        If CDate(txtDate.Text) - Curdate > 7 Then
            MsgBox "��Ժʱ��ȵ�ǰʱ���ù���,���飡", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        If MsgBox("��Ժʱ������˵�ǰϵͳʱ��,ȷʵҪ��Ժ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txtDate.SetFocus: Exit Sub
        End If
    End If
    
    '��鲡���Ƿ���δִ����ɵ�������Ŀ��δ��ҩƷ
    strInfo = ""
    If gbyt��Ժʱ���δִ�� <> 0 Then
        strInfo = ExistWaitExe(mlng����ID, mlng��ҳID)
        If strInfo <> "" Then
            If gbyt��Ժʱ���δִ�� = 1 Then
                If MsgBox("���ֲ���" & txt����.Text & "������δִ����ɵ����ݣ�" & _
                    vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "ȷ��Ҫ��Ժ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            Else
                MsgBox "���ֲ���" & txt����.Text & "������δִ����ɵ����ݣ�" & vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "�������Ժ.", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    '����30208 by lesfeng 2010-08-02 ���ֲ���22��32 ����154��155
    strInfo = ""
    If gbyt��Ժʱ���ҩƷδִ�� <> 0 Then
        strInfo = ExistWaitDrug(mlng����ID, mlng��ҳID)
        If strInfo <> "" Then
            If gbyt��Ժʱ���ҩƷδִ�� = 1 Then
                If MsgBox("���ֲ���" & txt����.Text & strInfo & vbCrLf & vbCrLf & "ȷ��Ҫ��Ժ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            Else
                MsgBox "���ֲ���" & txt����.Text & strInfo & vbCrLf & vbCrLf & "�������Ժ��", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        '30339:������,2012-09-14,����Ƿ�Ѫ
        strInfo = ExistWaitBool(mlng����ID, mlng��ҳID)
        If strInfo <> "" Then
            If gbyt��Ժʱ���ҩƷδִ�� = 1 Then
                If MsgBox("���ֲ���" & txt����.Text & strInfo & vbCrLf & vbCrLf & "ȷ��Ҫ��Ժ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            Else
                MsgBox "���ֲ���" & txt����.Text & strInfo & vbCrLf & vbCrLf & "�������Ժ��", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    
    If GetUnAuditReFee(mlng����ID, mlng��ҳID) Then
        If MsgBox("����" & txt����.Text & "�����������˷ѵ�δ��˵ļ�¼,ȷ��Ҫ��Ժ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    dMax = GetMaxDate(mlng����ID, mlng��ҳID)
    If Format(txtDate.Text, "yyyyMMddHHmmss") <= Format(dMax, "yyyyMMddHHmmss") Then
        MsgBox "���˳�Ժʱ�������ڸò����ϴα䶯ʱ�� " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & " ��", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    
    dMax = GetLastAdviceTime(mlng����ID, mlng��ҳID)
    If Format(txtDate.Text, "yyyyMMddHHmmss") < Format(dMax, "yyyyMMddHHmmss") Then
        If MsgBox("��Ժʱ��С�ڸò��������Чҽ����ʱ�� " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & ",ȷʵҪ��Ժ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txtDate.SetFocus: Exit Sub
        End If
    End If
    '����28612 by lesfeng 2010-07-05
    If InStr(cbo��Ժ��ʽ.Text, "����") = 0 And mintDeath = 1 Then
        If MsgBox("�ò��˴�����Ч�ٴ�����ҽ��,������ҽ����ʱ�� " & Format(mdteDeathDate, "yyyy-MM-dd HH:mm:ss") & ",����Ժ��ʽ��Ϊ����,ȷʵҪ��Ժ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            cbo��Ժ��ʽ.SetFocus: Exit Sub
        End If
    End If
    
    If InStr(cbo��Ժ��ʽ.Text, "����") > 0 And mintDeath = 1 Then
        If Format(txtDate.Text, "yyyyMMddHHmmss") <> Format(mdteDeathDate, "yyyyMMddHHmmss") Then
            If MsgBox("��Ժʱ�䲻���ڸò�����Ч�ٴ�����ҽ����ʱ�� " & Format(mdteDeathDate, "yyyy-MM-dd HH:mm:ss") & ",ȷʵҪ��Ժ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                txtDate.SetFocus: Exit Sub
            End If
        End If
    End If
    
    If InStr(cbo��Ժ��ʽ.Text, "����") > 0 And mintDeath = 0 And mblnOutDeath = True Then
        MsgBox "�ò��˳�Ժ��ʽΪ����,����������Ч�ٴ�����ҽ��,�������Ժ!", vbInformation, gstrSysName
        cbo��Ժ��ʽ.SetFocus: Exit Sub
    End If
    
    '68953:������,2012-09-14
    strInfo = ""
    If gbyt��Ժʱ���ڻ������ݼ�� <> 0 Then
        strInfo = ExistNurseData(mlng����ID, mlng��ҳID, CDate(Format(txtDate.Text, "YYYY-MM-DD HH:mm:ss")))
        If strInfo <> "" Then
            If strInfo = "OK" Then
                '�ϰ�
                If gbyt��Ժʱ���ڻ������ݼ�� = 1 Then
                    If MsgBox("���ֲ���" & txt����.Text & "���ڳ�Ժʱ��֮��Ļ������ݣ�ȷ��Ҫ��Ժ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Sub
                    End If
                Else
                    MsgBox "���ֲ���" & txt����.Text & "���ڳ�Ժʱ��֮��Ļ������ݣ��������Ժ.", vbInformation, gstrSysName
                    Exit Sub
                End If
            Else
                '�°�
                If gbyt��Ժʱ���ڻ������ݼ�� = 1 Then
                    If MsgBox("���ֲ���" & txt����.Text & "���ڳ�Ժʱ��֮��Ļ������ݣ�" & _
                        vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "ȷ��Ҫ��Ժ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Sub
                    End If
                Else
                    MsgBox "���ֲ���" & txt����.Text & "���ڳ�Ժʱ��֮��Ļ������ݣ�" & vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "�������Ժ.", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
    End If
    
    '����28982 by lesfeng 2010-06-09
    strȷ������ = ""
    If chk����.Value = 1 Then
        str��Ժʱ�� = Format(mrsPatiInfo!��Ժʱ��, "yyyy-MM-dd HH:mm:ss")
        If Not IsDate(txtOkDate.Text) Then
            MsgBox "��������ȷ�Ĳ���ȷ��ʱ�䣡", vbInformation, gstrSysName
            If txtOkDate.Enabled Then txtOkDate.SetFocus: Exit Sub
        End If
        If Format(txtOkDate.Text, "yyyyMMddHHmmss") >= Format(txtDate.Text, "yyyyMMddHHmmss") Then
            MsgBox "ȷ��ʱ�����С�ڲ��˳�Ժʱ�� " & Format(txtDate.Text, "yyyy-MM-dd HH:mm:ss") & " ��", vbInformation, gstrSysName
            If txtOkDate.Enabled Then txtOkDate.SetFocus: Exit Sub
        End If
        
        If Format(str��Ժʱ��, "yyyyMMddHHmmss") > Format(txtOkDate.Text, "yyyyMMddHHmmss") Then
            MsgBox "ȷ��ʱ�������ڵ��ڲ�����Ժʱ�� " & Format(str��Ժʱ��, "yyyy-MM-dd HH:mm:ss") & " ��", vbInformation, gstrSysName
            If txtOkDate.Enabled Then txtOkDate.SetFocus: Exit Sub
        End If
        
        strȷ������ = Format(txtOkDate.Text, "yyyy-MM-dd HH:mm:ss")
    End If
   
    If cbo����.ListIndex <> -1 Then int���� = cbo����.ItemData(cbo����.ListIndex)
    
    If InStr(1, txt��Ժ���.Tag, ";") <= 0 Then
        lng��ҽ����ID = Val(txt��Ժ���.Tag)
    Else
        lng��ҽ���ID = Val(txt��Ժ���.Tag)
    End If
    If InStr(1, txt��ҽ���.Tag, ";") <= 0 Then
        lng��ҽ����ID = Val(txt��ҽ���.Tag)
    Else
        lng��ҽ���ID = Val(txt��ҽ���.Tag)
    End If
    '����28982 by lesfeng 2010-06-09 ����ȷ������
    strSQL = "zl_���˱䶯��¼_Out(" & mlng����ID & "," & mlng��ҳID & "," & _
        ZVal(lng��ҽ����ID) & "," & ZVal(lng��ҽ���ID) & ",'" & Replace(txt��Ժ���.Text, "'", "''") & "','" & zlCommFun.GetNeedName(cbo��Ժ���.Text) & "'," & _
        ZVal(lng��ҽ����ID) & "," & ZVal(lng��ҽ���ID) & ",'" & Replace(txt��ҽ���.Text, "'", "''") & "','" & zlCommFun.GetNeedName(cbo��ҽ��Ժ���.Text) & "'," & _
        chk����.Value & ",'" & zlCommFun.GetNeedName(cbo��Ժ��ʽ.Text) & "',To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS')," & _
        IIf(chk����.Value = 1, int����, 0) & "," & IIf(chk����.Value = 1 And int���� <> 9, Val(txt����.Text), "Null") & "," & IIf(chkʬ��.Enabled, chkʬ��.Value, "NULL") & "," & _
        "'" & UserInfo.��� & "','" & UserInfo.���� & "'," & IIf(strȷ������ = "", "NULL", "To_Date('" & strȷ������ & "','YYYY-MM-DD HH24:MI:SS')") & ")"
    
    gcnOracle.BeginTrans
    blnTrans = True
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
     '����28139 by lesfeng 2010-03-03 �����ж�
    With vfg��ҽ
        For int���� = 1 To .Rows - 1
            str������� = Trim(.TextMatrix(int����, .ColIndex("�������")))
            str������� = Trim(.TextMatrix(int����, .ColIndex("��Ժ���")))
            
            lng��ҽ����ID = Val(.TextMatrix(int����, .ColIndex("����ID")))
            lng��ҽ���ID = Val(.TextMatrix(int����, .ColIndex("���ID")))
            int������� = 3
            int��ϴ��� = int���� + 1
            If str������� <> "" Then
                '����id,��ҳid,�������,��ϴ���,����id,���id,��Ժ���,������Ϣ,�Ƿ�����
                strSQL = "Zl_����������_Other(" & mlng����ID & "," & mlng��ҳID & "," & int������� & "," & int��ϴ��� & _
                        "," & ZVal(lng��ҽ����ID) & "," & ZVal(lng��ҽ���ID) & _
                        ",'" & zlCommFun.GetNeedName(str�������) & "','" & Replace(str�������, "'", "��") & _
                        "'" & IIf(.TextMatrix(int����, .ColIndex("����")) <> "", ",1", ",0") & ",'" & UserInfo.���� & "')"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
        Next
    End With
    If cbo��ҽ��Ժ���.Enabled Then
        With vfg��ҽ
            For int���� = 1 To .Rows - 1
                str������� = Trim(.TextMatrix(int����, .ColIndex("�������")))
                str������� = Trim(.TextMatrix(int����, .ColIndex("��Ժ���")))
                lng��ҽ����ID = Val(.TextMatrix(int����, .ColIndex("����ID")))
                lng��ҽ���ID = Val(.TextMatrix(int����, .ColIndex("���ID")))
                int������� = 13
                int��ϴ��� = int���� + 1
                If str������� <> "" Then
                    strSQL = "Zl_����������_Other(" & mlng����ID & "," & mlng��ҳID & "," & int������� & "," & int��ϴ��� & _
                        "," & ZVal(lng��ҽ����ID) & "," & ZVal(lng��ҽ���ID) & _
                        ",'" & zlCommFun.GetNeedName(str�������) & "','" & Replace(str�������, "'", "��") & "',0,'" & UserInfo.���� & "')"
                    zlDatabase.ExecuteProcedure strSQL, Me.Caption
                End If
            Next
        End With
    End If
    
    'ҽ���Ķ�
    If int���� <> 0 Then
        If Not gclsInsure.LeaveSwap(mlng����ID, mlng��ҳID, int����) Then
            gcnOracle.RollbackTrans: Exit Sub
        End If
    End If
    gcnOracle.CommitTrans
    blnTrans = False
    
    On Error Resume Next
    '��Ժ�ɹ��󴥷�����Ϣ
    If mclsMipModule.IsConnect = True Then
        mclsXML.ClearXmlText '��������е�XML
        '--������Ϣ��װ
        '������Ϣ
        mclsXML.AppendNode "in_patient"
        'patient_id      ����id  1   N
        mclsXML.appendData "patient_id", mlng����ID, xsNumber  '����ID
        'page_id     ��ҳid  1   N
        mclsXML.appendData "page_id", mlng��ҳID, xsNumber '��ҳID
        'patient_name        ����    1   S
        mclsXML.appendData "patient_name", txt����.Text, xsString  '����
        'patient_sex     �Ա�    0..1    S
        mclsXML.appendData "patient_sex", txt�Ա�.Text, xsString  '�Ա�
        'in_number       סԺ��  1   S
        mclsXML.appendData "in_number", txtסԺ��.Text, xsString 'סԺ��
        mclsXML.AppendNode "in_patient", True
        
        strSQL = "Select ID �䶯id,��ʼʱ�� �䶯ʱ�� From ���˱䶯��¼ where ����ID=[1] And ��ҳId=[2] And ��ֹԭ��=1 And ��ֹʱ�� IS NOT NULL And NVL(���Ӵ�λ,0)=0 "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���˱䶯��¼", mlng����ID, mlng��ҳID)
        
        'out_hospital        ���˳�Ժ    1
        mclsXML.AppendNode "out_hospital"
        'change_id       ��Ժ���id  1   N
        mclsXML.appendData "change_id", rsTmp!�䶯ID, xsNumber
        'out_date        ���ʱ��    1   s
        mclsXML.appendData "out_date", Format(rsTmp!�䶯ʱ��, "YYYY-MM-DD HH:mm:ss"), xsString
        'out_area_id     ��ǰ����id  0..1    N
        mclsXML.appendData "out_area_id", Nvl(mrsPatiInfo!��ǰ����ID, 0), xsNumber
        'out_area_title      ��ǰ����    0..1    S
        mclsXML.appendData "out_area_title", Nvl(mrsPatiInfo!��ǰ����), xsString
        'out_dept_id     ��ǰ����id    1   N
        mclsXML.appendData "out_dept_id", Nvl(mrsPatiInfo!��Ժ����id, 0), xsNumber
        'out_dept_title      ��ǰ����  1   S
        mclsXML.appendData "out_dept_title", Nvl(mrsPatiInfo!��ǰ����), xsString
        'out_room        ��ǰ����    0..1    S
        mclsXML.appendData "out_room", txt����.Tag, xsString
        'out_bed     ��ǰ����    1   S
        mclsXML.appendData "out_bed", Nvl(mrsPatiInfo!��Ҫ����), xsString
        'out_way     ��Ժ��ʽ    1   S
        mclsXML.appendData "out_way", zlCommFun.GetNeedName(cbo��Ժ��ʽ.Text), xsString
        'treat_state     �������    1   S
        mclsXML.appendData "treat_state", zlCommFun.GetNeedName(cbo��Ժ���.Text), xsString
        
        mclsXML.AppendNode "out_hospital", True
        mclsMipModule.CommitMessage "ZLHIS_PATIENT_010", mclsXML.XmlText
    End If
    If Err <> 0 Then Err.Clear
    
     '������ҽӿ�
    If CreatePlugInOK(glngModul) Then
        On Error Resume Next
        Call gobjPlugIn.InPatiCheckOutAfter(mlng����ID, mlng��ҳID)
        Call zlPlugInErrH(Err, "InPatiCheckOutAfter")
        Err.Clear: On Error GoTo 0
    End If
    
    On Error GoTo errH
    
    Dim strOut As String
    Call zlExcuteUploadSwap(mlng����ID, strOut) '�����˵�������һ��ͨ�ϴ�����
    
    '��Ժ���Զ����㲡�˵Ĵ�λ���úͻ������(������ڳ�Ժǰִ�У���ʹ�ð���ģʽʱ�����������)
    strSQL = "ZL1_AUTOCPTPATI(" & mlng����ID & "," & mlng��ҳID & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
       
    gblnOK = True
    
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'ж����Ϣ����
    If Not (mclsMipModule Is Nothing) Then
        Call mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    If Not (mclsXML Is Nothing) Then
        Set mclsXML = Nothing
    End If
    
    Call SaveHead(vfg��ҽ, 1)
    Call SaveHead(vfg��ҽ, 2)
End Sub

Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_LostFocus()
    If Not IsDate(txtDate.Text) Then txtDate.SetFocus
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt��Ժ���_GotFocus()
    zlControl.TxtSelAll txt��Ժ���
End Sub

Private Sub txt��ҽ���_GotFocus()
    zlControl.TxtSelAll txt��ҽ���
End Sub

Private Sub txt��Ժ���_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInput As String, strSex As String, lngTxtHeight As Long, vPoint As POINTAPI
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not RequestCode Then
            '����25785 by lesfeng 2009-10-20 ������������¼�����
            '************************************************
            If gintסԺ������� = 1 Then
                strInput = UCase(txt��Ժ���.Text)
                strSex = zlCommFun.GetNeedName(txt�Ա�.Text)
                
                If zlCommFun.IsCharChinese(strInput) Then
                    strSQL = "���� Like [2] or '('||����||')'||���� Like [2]" '���뺺��ʱֻƥ������
                Else
                    strSQL = "���� Like [1] Or ���� Like [2] Or " & IIf(gbytCode = 0, "����", "�����") & " Like [2]"
                End If
                
                strSQL = _
                        " Select ID,ID as ��ĿID,����,����,����," & IIf(gbytCode = 0, "����", "����� as ����") & ",˵��" & _
                        " From ��������Ŀ¼ Where Instr([3],���)>0 And (" & strSQL & ")" & _
                        IIf(strSex <> "", " And (�Ա�����=[4] Or �Ա����� is NULL)", "") & _
                        " And (����ʱ�� is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " Order by ����"
                '����27613 by lesfeng 2010-01-21
                '����¼��ʱ�ж��ƥ��(����)������ѡ��,���ּ���ĸ�����ѡ��
                If zlCommFun.IsCharChinese(strInput) Then
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", gstrLike & strInput & "%", "D", strSex, gbytCode + 1)
                    If rsTmp.EOF Then
                        Set rsTmp = Nothing
                    ElseIf rsTmp.RecordCount > 1 Then
                        Set rsTmp = Nothing '����¼��ʱ�ж��ƥ�䲻����ѡ��
                    End If
                Else
                    vPoint = zlControl.GetCoordPos(fraInfo.hWnd, txt��Ժ���.Left, txt��Ժ���.Top)
                    strInput = UCase(txt��Ժ���.Text)
                    strSex = zlCommFun.GetNeedName(txt�Ա�.Text)
                    lngTxtHeight = txt��Ժ���.Height
                    Set rsTmp = GetDiseaseCodeNew(Me, blnCancel, strInput, strSex, "D", vPoint.X, vPoint.Y, lngTxtHeight)
                    If Not rsTmp Is Nothing Then
                        If rsTmp.EOF Then
                            Set rsTmp = Nothing
                        End If
                    End If
                End If
                If Not rsTmp Is Nothing Then
                    '���ݿ���ֻ��һ��ƥ����Ŀ�����Ը�ƥ�����ĿΪ׼
                    txt��Ժ���.Tag = rsTmp!ID
                    txt��Ժ���.Text = "(" & rsTmp!���� & ")" & rsTmp!���� '
                    lbl��Ժ���.Tag = txt��Ժ���.Text '���ڻָ���ʾ
                Else
                    '���������ƥ����Ŀʱ���������Ϊ׼
                    txt��Ժ���.Tag = ""
                    lbl��Ժ���.Tag = txt��Ժ���.Text '���ڻָ���ʾ
                End If
            End If
            '************************************************
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt��Ժ���.Text = lbl��Ժ���.Tag And txt��Ժ���.Text <> "" Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt��Ժ���.Text = "" Then
            txt��Ժ���.Tag = "": lbl��Ժ���.Tag = ""
            Call zlCommFun.PressKey(vbKeyTab) '��������
        Else
            vPoint = zlControl.GetCoordPos(fraInfo.hWnd, txt��Ժ���.Left, txt��Ժ���.Top)
            strInput = UCase(txt��Ժ���.Text)
            strSex = zlCommFun.GetNeedName(txt�Ա�.Text)
            lngTxtHeight = txt��Ժ���.Height
            Set rsTmp = GetDiseaseCode(Me, blnCancel, strInput, strSex, "D", vPoint.X, vPoint.Y, lngTxtHeight)
            If Not rsTmp Is Nothing Then
                txt��Ժ���.Tag = rsTmp!ID
                txt��Ժ���.Text = "(" & rsTmp!���� & ")" & rsTmp!����
                lbl��Ժ���.Tag = txt��Ժ���.Text '���ڻָ���ʾ
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                If Not blnCancel Then
                    MsgBox "û���ҵ�ƥ��ļ������롣", vbInformation, gstrSysName
                End If
                If lbl��Ժ���.Tag <> "" Then txt��Ժ���.Text = lbl��Ժ���.Tag
                Call txt��Ժ���_GotFocus
                txt��Ժ���.SetFocus
            End If
        End If
    Else
        CheckInputLen txt��Ժ���, KeyAscii
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt��ҽ���_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInput As String, strSex As String, lngTxtHeight As Long, vPoint As POINTAPI
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not RequestCode Then
            '����25785 by lesfeng 2009-10-20 ������������¼�����
            '************************************************
            If gintסԺ������� = 1 Then
                strInput = UCase(txt��ҽ���.Text)
                strSex = zlCommFun.GetNeedName(txt�Ա�.Text)
                
                If zlCommFun.IsCharChinese(strInput) Then
                    strSQL = "���� Like [2] or '('||����||')'||���� Like [2]" '���뺺��ʱֻƥ������
                Else
                    strSQL = "���� Like [1] Or ���� Like [2] Or " & IIf(gbytCode = 0, "����", "�����") & " Like [2]"
                End If
                
                strSQL = _
                        " Select ID,ID as ��ĿID,����,����,����," & IIf(gbytCode = 0, "����", "����� as ����") & ",˵��" & _
                        " From ��������Ŀ¼ Where Instr([3],���)>0 And (" & strSQL & ")" & _
                        IIf(strSex <> "", " And (�Ա�����=[4] Or �Ա����� is NULL)", "") & _
                        " And (����ʱ�� is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " Order by ����"
                '����27613 by lesfeng 2010-01-21
                '����¼��ʱ�ж��ƥ��(����)������ѡ��,���ּ���ĸ�����ѡ��
                If zlCommFun.IsCharChinese(strInput) Then
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", gstrLike & strInput & "%", "B", strSex, gbytCode + 1)
                    If rsTmp.EOF Then
                        Set rsTmp = Nothing
                    ElseIf rsTmp.RecordCount > 1 Then
                        Set rsTmp = Nothing '����¼��ʱ�ж��ƥ�䲻����ѡ��
                    End If
                Else
                    vPoint = zlControl.GetCoordPos(fraInfo.hWnd, txt��ҽ���.Left, txt��ҽ���.Top)
                    strInput = UCase(txt��ҽ���.Text)
                    strSex = zlCommFun.GetNeedName(txt�Ա�.Text)
                    lngTxtHeight = txt��ҽ���.Height
                    Set rsTmp = GetDiseaseCodeNew(Me, blnCancel, strInput, strSex, "B", vPoint.X, vPoint.Y, lngTxtHeight)
                    If Not rsTmp Is Nothing Then
                        If rsTmp.EOF Then
                            Set rsTmp = Nothing
                        End If
                    End If
                End If
                If Not rsTmp Is Nothing Then
                    '���ݿ���ֻ��һ��ƥ����Ŀ�����Ը�ƥ�����ĿΪ׼
                    txt��ҽ���.Tag = rsTmp!ID
                    txt��ҽ���.Text = "(" & rsTmp!���� & ")" & rsTmp!���� '
                    lbl��ҽ���.Tag = txt��ҽ���.Text '���ڻָ���ʾ
                Else
                    '���������ƥ����Ŀʱ���������Ϊ׼
                    txt��ҽ���.Tag = ""
                    lbl��ҽ���.Tag = txt��ҽ���.Text '���ڻָ���ʾ
                End If
            End If
            '************************************************
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt��ҽ���.Text = lbl��ҽ���.Tag And txt��ҽ���.Text <> "" Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt��ҽ���.Text = "" Then
            txt��ҽ���.Tag = "": lbl��ҽ���.Tag = ""
            Call zlCommFun.PressKey(vbKeyTab) '��������
        Else
            vPoint = zlControl.GetCoordPos(fraInfo.hWnd, txt��ҽ���.Left, txt��ҽ���.Top)
            strInput = UCase(txt��ҽ���.Text)
            strSex = zlCommFun.GetNeedName(txt�Ա�.Text)
            lngTxtHeight = txt��ҽ���.Height
            Set rsTmp = GetDiseaseCode(Me, blnCancel, strInput, strSex, "B", vPoint.X, vPoint.Y, lngTxtHeight)
            
            If Not rsTmp Is Nothing Then
                txt��ҽ���.Tag = rsTmp!ID
                txt��ҽ���.Text = "(" & rsTmp!���� & ")" & rsTmp!����
                lbl��ҽ���.Tag = txt��ҽ���.Text '���ڻָ���ʾ
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                If Not blnCancel Then
                    MsgBox "û���ҵ�ƥ��ļ������롣", vbInformation, gstrSysName
                End If
                If lbl��ҽ���.Tag <> "" Then txt��ҽ���.Text = lbl��ҽ���.Tag
                Call txt��ҽ���_GotFocus
                txt��ҽ���.SetFocus
            End If
        End If
    Else
        CheckInputLen txt��ҽ���, KeyAscii
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt��Ժ���_Validate(Cancel As Boolean)
    If Val(txt��Ժ���.Tag) > 0 And txt��Ժ���.Text <> lbl��Ժ���.Tag Then
        txt��Ժ���.Text = lbl��Ժ���.Tag
    ElseIf Val(txt��Ժ���.Tag) = 0 And RequestCode Then
        txt��Ժ���.Text = ""
    End If
    
    If txt��Ժ���.Text <> "" And cbo��Ժ���.Text = "" Then
        cbo��Ժ���.ListIndex = cbo.FindIndex(cbo��Ժ���, 1)
        If cbo��Ժ���.ListIndex = -1 Then cbo��Ժ���.ListIndex = 0
    ElseIf txt��Ժ���.Text = "" And cbo��Ժ���.Text <> "" Then
        cbo��Ժ���.ListIndex = 0
    End If
End Sub

Private Sub txt��ҽ���_Validate(Cancel As Boolean)
    If Val(txt��ҽ���.Tag) > 0 And txt��ҽ���.Text <> lbl��ҽ���.Tag Then
        txt��ҽ���.Text = lbl��ҽ���.Tag
    ElseIf Val(txt��ҽ���.Tag) = 0 And RequestCode Then
        txt��ҽ���.Text = ""
    End If
    
    If txt��ҽ���.Text <> "" And cbo��ҽ��Ժ���.Text = "" Then
        cbo��ҽ��Ժ���.ListIndex = cbo.FindIndex(cbo��ҽ��Ժ���, 1)
        If cbo��ҽ��Ժ���.ListIndex = -1 Then cbo��ҽ��Ժ���.ListIndex = 0
    ElseIf txt��ҽ���.Text = "" And cbo��ҽ��Ժ���.Text <> "" Then
        cbo��ҽ��Ժ���.ListIndex = 0
    End If
End Sub

Private Function RequestCode() As Boolean
    RequestCode = gintסԺ������� = 2 Or (gintסԺ������� = 3 And Val("" & mrsPatiInfo!����) <> 0)
End Function

'����28139 by lesfeng 2010-03-02
Private Sub initvfgHeadTitle(ByVal vsGrid As VSFlexGrid, ByVal intFlag As Integer)
    Dim strHead As String
    If intFlag = 1 Then
        strHead = "���,500,4,1;�������,2200,1,1;ICD����,1000,1,1;��Ժ���,1000,1,1;����,800,4,0;���ID,0,1,-1;����ID,0,1,-1"
    Else
        strHead = "���,500,4,1;�������,2800,1,1;��ҽ����,1200,1,1;��Ժ���,1000,1,1;���ID,0,1,-1;����ID,0,1,-1"
    End If
        Call SetVsFlexGridChangeHead(strHead, vsGrid, 1)
End Sub

Private Sub SetVfgNo(ByVal vsGrid As VSFlexGrid)
    Dim i As Long
    With vsGrid
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("�������"))) <> "" Then
                .TextMatrix(i, .ColIndex("���")) = i
            End If
        Next
    End With
End Sub

Private Sub SetInitVfgFormat(ByVal vsGrid As VSFlexGrid, ByVal intFlag As Integer)
    Dim i As Long
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "Select ����,����,����||'-'||Nvl(����,'') as ��Ŀ,Nvl(ȱʡ��־,0) as ȱʡ From ���ƽ�� Order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)

    With vsGrid
        .ColComboList(.ColIndex("��Ժ���")) = .BuildComboList(rsTemp, "��Ŀ", "����")
        If rsTemp.RecordCount = 1 Then
            .ColData(.ColIndex("��Ժ���")) = Nvl(rsTemp!����) & ";" & Nvl(rsTemp!��Ŀ)
        Else
            rsTemp.Filter = "ȱʡ=1"
            If rsTemp.EOF = False Then
                .ColData(.ColIndex("��Ժ���")) = Nvl(rsTemp!����) & ";" & Nvl(rsTemp!��Ŀ)
            Else
                .ColData(.ColIndex("��Ժ���")) = ";"
            End If
        End If
        .ExplorerBar = flexExSortShowAndMove
        .SelectionMode = flexSelectionByRow
        If (txt��ҽ���.Enabled And intFlag = 2) Or intFlag = 1 Then
            .Editable = flexEDKbdMouse
        Else
            .Editable = flexEDNone
        End If
    End With
    rsTemp.Close
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadVfgData(ByVal vsGrid As VSFlexGrid, ByVal intFlag As Integer)
    Dim strTemp As String
    Dim i As Long
    Dim rsDiagnosisOther As ADODB.Recordset

    If intFlag = 1 Then
        Set rsDiagnosisOther = GetDiagnosticOtherInfo(mlng����ID, mlng��ҳID, "1,2,3", "2,3")
    Else
        Set rsDiagnosisOther = GetDiagnosticOtherInfo(mlng����ID, mlng��ҳID, "11,12,13", "2,3")
    End If
            
    With vsGrid
        .Clear
        Call initvfgHeadTitle(vsGrid, intFlag)
        If Not rsDiagnosisOther Is Nothing Then
            If intFlag = 1 Then
                'a.��ҽ���
                rsDiagnosisOther.Filter = "�������=3 and ��¼��Դ=3"            '��ȡ��ҳ����ĳ�Ժ���
                If Not rsDiagnosisOther.EOF Then
                    .Rows = IIf(rsDiagnosisOther.EOF, 0, rsDiagnosisOther.RecordCount) + 1
                Else
                    '����28483 by lesfeng 2010-03-01
                    rsDiagnosisOther.Filter = "�������=3 and ��¼��Դ=2"        '��ȡ��Ժ�Ǽǵĳ�Ժ���
                    If Not rsDiagnosisOther.EOF Then
                        .Rows = IIf(rsDiagnosisOther.EOF, 0, rsDiagnosisOther.RecordCount) + 1
                    Else
                        '����28138 by lesfeng 2010-03-01 ����Ĭ����ϵ��ж� ����ȡ������ϼ���Ժ���
                        If mintĬ����� = 1 Then
                            rsDiagnosisOther.Filter = "�������=2 and ��¼��Դ=2"        '��ȡ��Ժ�Ǽǵ���Ժ���
                            If Not rsDiagnosisOther.EOF Then
                                .Rows = IIf(rsDiagnosisOther.EOF, 0, rsDiagnosisOther.RecordCount) + 1
                            Else
                                rsDiagnosisOther.Filter = "�������=1 and ��¼��Դ=2"    '���ȡ��Ժ�Ǽǵ��������
                                If Not rsDiagnosisOther.EOF Then
                                    .Rows = IIf(rsDiagnosisOther.EOF, 0, rsDiagnosisOther.RecordCount) + 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                'b.��ҽ���
                rsDiagnosisOther.Filter = "�������=13 and ��¼��Դ=3"            '��ȡ��ҳ����ĳ�Ժ���
                If Not rsDiagnosisOther.EOF Then
                    .Rows = IIf(rsDiagnosisOther.EOF, 0, rsDiagnosisOther.RecordCount) + 1
                Else
                    '����28483 by lesfeng 2010-03-01
                    rsDiagnosisOther.Filter = "�������=13 and ��¼��Դ=2"        '��ȡ��Ժ�Ǽǵĳ�Ժ���
                    If Not rsDiagnosisOther.EOF Then
                        .Rows = IIf(rsDiagnosisOther.EOF, 0, rsDiagnosisOther.RecordCount) + 1
                    Else
                        '����28138 by lesfeng 2010-03-01 ����Ĭ����ϵ��ж� ����ȡ������ϼ���Ժ���
                        If mintĬ����� = 1 Then
                            rsDiagnosisOther.Filter = "�������=12 and ��¼��Դ=2"        '��ȡ��Ժ�Ǽǵ���Ժ���
                            If Not rsDiagnosisOther.EOF Then
                                .Rows = IIf(rsDiagnosisOther.EOF, 0, rsDiagnosisOther.RecordCount) + 1
                            Else
                                rsDiagnosisOther.Filter = "�������=11 and ��¼��Դ=2"    '���ȡ��Ժ�Ǽǵ��������
                                If Not rsDiagnosisOther.EOF Then
                                    .Rows = IIf(rsDiagnosisOther.EOF, 0, rsDiagnosisOther.RecordCount) + 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            
            '�������,��¼��Դ,�������,����ID,���ID,��Ժ���,��¼����,�Ƿ�����
            If Not rsDiagnosisOther.EOF Then
                For i = 1 To .Rows - 1
                    .TextMatrix(i, .ColIndex("�������")) = IIf(IsNull(rsDiagnosisOther!�������), "", rsDiagnosisOther!�������)
                    .TextMatrix(i, .ColIndex("��Ժ���")) = IIf(IsNull(rsDiagnosisOther!��Ժ���), "", rsDiagnosisOther!��Ժ���)
                    If intFlag = 1 Then
                       .TextMatrix(i, .ColIndex("ICD����")) = IIf(IsNull(rsDiagnosisOther!����), "", rsDiagnosisOther!����)
                        .TextMatrix(i, .ColIndex("����")) = IIf(IsNull(rsDiagnosisOther!�Ƿ�����), "", IIf(rsDiagnosisOther("�Ƿ�����") = 1, "��", ""))
                    Else
                        .TextMatrix(i, .ColIndex("��ҽ����")) = IIf(IsNull(rsDiagnosisOther!����), "", rsDiagnosisOther!����)
                    End If
                    .TextMatrix(i, .ColIndex("����ID")) = IIf(IsNull(rsDiagnosisOther!����ID), 0, rsDiagnosisOther!����ID)
                    .TextMatrix(i, .ColIndex("���ID")) = IIf(IsNull(rsDiagnosisOther!���ID), 0, rsDiagnosisOther!���ID)
                    rsDiagnosisOther.MoveNext
                Next
                .Rows = .Rows + 1
    
'            Else
'                .Rows = .Rows + 1
            End If
            
'            If .Rows > 1 Then
'                .Select 1, .ColIndex("vsGrid")
'            End If
        End If
    End With
    Call SetVfgNo(vsGrid)
    Call SetInitVfgFormat(vsGrid, intFlag)
    Call RestoreHead(vsGrid, intFlag)
End Sub

Private Sub vfg��ҽ_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strValue As String
    With vfg��ҽ
        Select Case Col
           Case .ColIndex("�������")
                strValue = Trim(.TextMatrix(Row, .ColIndex("�������")))
                If Not IsNull(strValue) Then
                    If Not GetDiagnosis(vfg��ҽ, strValue, 1) Then
                        .Select Row, Col
                        .TextMatrix(Row, .ColIndex("�������")) = mstrOldName
                        mstrOldName = ""
                        Exit Sub
                    End If
                    Call SetVfgNo(vfg��ҽ)
                End If
            Case .ColIndex("��Ժ���")
                If .ComboIndex < 0 Then Exit Sub
                .Cell(flexcpData, Row, Col) = .ComboData(.ComboIndex)
                .TextMatrix(Row, Col) = .ComboItem(.ComboIndex)
        End Select
    End With
End Sub

Private Sub vfg��ҽ_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vfg��ҽ.Editable = flexEDNone Then
        Cancel = True
        Exit Sub
    End If
    Select Case Col
        Case vfg��ҽ.ColIndex("�������")
            mstrOldName = Trim(vfg��ҽ.TextMatrix(Row, vfg��ҽ.ColIndex("�������")))
            Cancel = False
            Exit Sub
        Case vfg��ҽ.ColIndex("��Ժ���")  ', vfg��ҽ.ColIndex("����")
            Cancel = False
            Exit Sub
        Case Else
            Cancel = True
            Exit Sub
    End Select
End Sub

Private Sub vfg��ҽ_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    Select Case Col
        Case vfg��ҽ.ColIndex("���")
            Position = -1
            Exit Sub
    End Select
    If Position = 0 Then
        Position = Col
    End If
End Sub

Private Sub vfg��ҽ_BeforeSort(ByVal Col As Long, Order As Integer)
    Call zl_VsGridBeforeSort(vfg��ҽ, Col, Order)
    Call SetVfgNo(vfg��ҽ)
End Sub

Private Sub vfg��ҽ_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim RecReturn As Recordset
    Dim i As Long
    Dim j As Long
    Dim strTemp As String
    Dim lngRow As Long
    Dim lngCurrRow As Long
    Dim blnRow As Boolean
    Dim strValue As String
    
    strValue = ""
'    If InStr(vfg��ҽ.Cell(flexcpText, 0, Col), "ICD����") > 0 Then ' And mintDblick = 0
'         Err = 0: On Error GoTo ErrHand:
'        If Not GetDiagnosis(vfg��ҽ, strValue, 2) Then
'            vfg��ҽ.Select Row, Col
'            Exit Sub
'        End If
'        Exit Sub
'    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vfg��ҽ_DblClick()
    Dim strTemp As String
    Dim intCount As Integer

    If vfg��ҽ.Editable = flexEDKbdMouse Then
        With vfg��ҽ
            If .Row > 0 Then
                If .Col = .ColIndex("����") Then
                    If Trim(.TextMatrix(.Row, .ColIndex("����"))) = "" Then
                        .TextMatrix(.Row, .ColIndex("����")) = "��"
                    Else
                        .TextMatrix(.Row, .ColIndex("����")) = ""
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub vfg��ҽ_EnterCell()
    Dim strTemp As String
    Dim intCount As Integer
    Dim strKey As String
    Dim strValue As String
    
    If vfg��ҽ.Editable = flexEDKbdMouse Then
        With vfg��ҽ
            If .Row > 0 Then
                If .Col = .ColIndex("��Ժ���") Then
                    strTemp = .TextMatrix(.Row, .Col)
                    strKey = .ColData(.ColIndex("��Ժ���"))
                    strValue = Trim(.TextMatrix(.Row, .ColIndex("�������")))
                    If strTemp = "" And strValue <> "" Then
                        .TextMatrix(.Row, .Col) = Mid(strKey, InStr(1, strKey, ";") + 1)
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub vfg��ҽ_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
     If vfg��ҽ.Editable = flexEDKbdMouse Then
        If KeyCode = vbKeyDelete Then
            If vfg��ҽ.Row > 0 Then
                If MsgBox("��Ҫɾ����ǰ��¼��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                    vfg��ҽ.RemoveItem (vfg��ҽ.Row)
                    If vfg��ҽ.Row = 0 Then
                        vfg��ҽ.Rows = vfg��ҽ.Rows + 1
                        vfg��ҽ.Select vfg��ҽ.Rows - 1, vfg��ҽ.Col
                    End If
                End If
            End If
        End If
        
        If KeyCode = vbKeyInsert Then
            With vfg��ҽ
                If MsgBox("��Ҫ���Ӽ�¼��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                   vfg��ҽ.Rows = vfg��ҽ.Rows + 1
                   .Select vfg��ҽ.Rows - 1, vfg��ҽ.Col
                End If
            End With
        End If
    End If
    If KeyCode = vbKeyReturn Then
        lngRow = vfg��ҽ.Row
        If vfg��ҽ.Editable = flexEDKbdMouse Then ''�������,2200,1,1;��Ժ���,1000,1,1;����
            Call zlPvVsMoveGridCell(vfg��ҽ, vfg��ҽ.ColIndex("�������"), vfg��ҽ.ColIndex("����"), True, lngRow, SetHeadCodeData(vfg��ҽ))
        Else
            Call zlPvVsMoveGridCell(vfg��ҽ, vfg��ҽ.ColIndex("�������"), vfg��ҽ.ColIndex("����"), False, lngRow, SetHeadCodeData(vfg��ҽ))
        End If
    End If
    Call SetVfgNo(vfg��ҽ)
'    If KeyCode <> vbKeyReturn Then
'        vfg��ҽ.ColComboList(vfg��ҽ.ColIndex("ICD����")) = ""
'    End If
End Sub

Private Sub vfg��ҽ_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim lngRow As Long
    If KeyCode = vbKeyReturn Then
        lngRow = Row
        If vfg��ҽ.Editable = flexEDKbdMouse Then
            Call zlPvVsMoveGridCell(vfg��ҽ, vfg��ҽ.ColIndex("�������"), vfg��ҽ.ColIndex("����"), True, lngRow, SetHeadCodeData(vfg��ҽ))
        Else
            Call zlPvVsMoveGridCell(vfg��ҽ, vfg��ҽ.ColIndex("�������"), vfg��ҽ.ColIndex("����"), False, lngRow, SetHeadCodeData(vfg��ҽ))
        End If
    End If
End Sub

Private Sub vfg��ҽ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0 'Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub vfg��ҽ_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
        Case vbKeyBack, vbKeyEscape, 3, 22: Exit Sub
        Case Else
'            Select Case Col
'                Case vfgInDetail.ColIndex("����")
'                    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
'                        Exit Sub
'                    Else
'                        KeyAscii = 0
'                    End If
''            End Select

    End Select
End Sub

'Private Sub vfg��ҽ_KeyUp(KeyCode As Integer, Shift As Integer)
'    vfg��ҽ.ColComboList(vfg��ҽ.ColIndex("ICD����")) = "..."
'End Sub
'
'Private Sub vfg��ҽ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    vfg��ҽ.ColComboList(vfg��ҽ.ColIndex("ICD����")) = "..."
'End Sub

Private Sub SaveHead(ByVal vsGrid As VSFlexGrid, ByVal intFlag As Integer)
    If intFlag = 1 Then
        zl_VsGrid_SaveToPara vsGrid, Me.Caption, glngModul, "��ҽ�����ͷ��Ϣ", True, True
    Else
        zl_VsGrid_SaveToPara vsGrid, Me.Caption, glngModul, "��ҽ�����ͷ��Ϣ", True, True
    End If
End Sub

Private Sub RestoreHead(ByVal vsGrid As VSFlexGrid, ByVal intFlag As Integer)
    If intFlag = 1 Then
        zl_VsGrid_FromParaRestore vsGrid, Me.Caption, glngModul, "��ҽ�����ͷ��Ϣ", True, True
    Else
        zl_VsGrid_FromParaRestore vsGrid, Me.Caption, glngModul, "��ҽ�����ͷ��Ϣ", True, True
    End If
End Sub

Private Function SetHeadCodeData(ByRef vsGrid As VSFlexGrid) As String
    Dim i As Long
    Dim strTemp As String
    
    SetHeadCodeData = ""
    With vsGrid
        For i = 0 To .Cols - 1
            If vsGrid.Editable = flexEDKbdMouse Then
'                If i = .ColIndex("ICD����") Then
                    If IsNull(strTemp) Or strTemp = "" Then
                        strTemp = i & "||0"
                    Else
                        strTemp = strTemp & ";" & i & "||0"
                    End If
'                End If
            End If
        Next
    End With
    SetHeadCodeData = strTemp
End Function

Private Function GetDiagnosis(ByRef vsGrid As VSFlexGrid, ByVal strSearch As String, ByVal intFlag As Integer) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '����:��������,������
    '����:strSearch-��������ֵ,
    '����:��ֻ����һ��ֵʱ����True,���򷵻�False
    '--------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    Dim lngHeigth As Long
    Dim lngTop As Long
    Dim vRect As RECT
    Dim blnCancel As Boolean
    Dim StrCodeName As String
    Dim lng���Id As Long
    Dim strInput As String
    Dim strSex As String
    Dim str��� As String
    
    Dim lngRow As Long
    Dim i As Long
    Dim j As Long
    
    GetDiagnosis = False
    If strSearch = "" Then Exit Function
    strInput = UCase(strSearch)
    
    If intFlag = 1 Then
        str��� = "D"
    Else
        str��� = "B"
    End If
    
    On Error GoTo errHandle
    
    If Not RequestCode Then
        If gintסԺ������� = 1 Then
            strSex = zlCommFun.GetNeedName(txt�Ա�.Text)
            If zlCommFun.IsCharChinese(strInput) Then
                strSQL = "���� Like [2] or '('||����||')'||���� Like [2]" '���뺺��ʱֻƥ������
            Else
                strSQL = "���� Like [1] Or ���� Like [2] Or " & IIf(gbytCode = 0, "����", "�����") & " Like [2]"
            End If
            
            strSQL = _
                    " Select ID,ID as ��ĿID,����,����,����," & IIf(gbytCode = 0, "����", "����� as ����") & ",˵��" & _
                    " From ��������Ŀ¼ Where Instr([3],���)>0 And (" & strSQL & ")" & _
                    IIf(strSex <> "", " And (�Ա�����=[4] Or �Ա����� is NULL)", "") & _
                    " And (����ʱ�� is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Order by ����"
            '����¼��ʱ�ж��ƥ��(����)������ѡ��,���ּ���ĸ�����ѡ��
            If zlCommFun.IsCharChinese(strInput) Then
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput & "%", gstrLike & strInput & "%", str���, strSex, gbytCode + 1)
                If rsTemp.EOF Then
                    Set rsTemp = Nothing
                ElseIf rsTemp.RecordCount > 1 Then
                    Set rsTemp = Nothing '����¼��ʱ�ж��ƥ�䲻����ѡ��
                End If
            Else
                vRect = zlControl.GetControlRect(vsGrid.hWnd)
                lngTop = vRect.Top + vsGrid.CellTop + vsGrid.CellHeight
                Set rsTemp = GetDiseaseCodeNew(Me, blnCancel, strInput, strSex, str���, vRect.Left - 15, lngTop, lngHeigth)
'                A.ID,A.����,A.����,A.����,A.����,A.�����,A.˵��,A.�Ա�����
                If Not rsTemp Is Nothing Then
                    If rsTemp.EOF Then
                        Set rsTemp = Nothing
                    End If
                End If
            End If
            If Not rsTemp Is Nothing Then
                '���ݿ���ֻ��һ��ƥ����Ŀ�����Ը�ƥ�����ĿΪ׼
                i = 1
                With rsTemp
                    If UCase(TypeName(vsGrid)) = "VSFLEXGRID" Then
                        With vsGrid
                            lng���Id = IIf(IsNull(rsTemp!ID), 0, rsTemp!ID)
                            '�˶��ظ�
                            If Not ExamineInputRepeat(vsGrid, lng���Id, intFlag, .Row) Then
                               
                                If Not IsNull(.TextMatrix(j, .ColIndex("�������"))) And .TextMatrix(j, .ColIndex("�������")) <> "" Then
                                    If i <> 1 Then
                                        .Row = .Rows - 1
                                        If .Row + 1 = .Rows Then .Rows = .Rows + 1
                                    End If
                                Else
                                    If .Row + 1 = .Rows Then .Rows = .Rows + 1
                                End If
                    
                                If intFlag = 1 Then
                                    .TextMatrix(.Row, .ColIndex("ICD����")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                                Else
                                    .TextMatrix(.Row, .ColIndex("��ҽ����")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                                    .TextMatrix(.Row, .ColIndex("���ID")) = IIf(IsNull(rsTemp!ID), 0, rsTemp!ID)
                                End If
                                .TextMatrix(.Row, .ColIndex("����ID")) = IIf(IsNull(rsTemp!ID), 0, rsTemp!ID)
                                .TextMatrix(.Row, .ColIndex("�������")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                                If .Row + 1 = .Rows Then .Rows = .Rows + 1
                    '                    .Row = .Row + 1
                    
                                If intFlag = 1 Then
                                    .Select .Row, .ColIndex("ICD����")
                                Else
                                    .Select .Row, .ColIndex("��ҽ����")
                                End If
                            Else
                                .TextMatrix(.Row, .ColIndex("�������")) = mstrOldName
                            End If
                        End With
                    Else
                        .Close
                        If vsGrid.Enabled Then vsGrid.SetFocus
                        zlCommFun.PressKey vbKeyTab
                    End If
                    .Close
                End With
            Else
                '���������ƥ����Ŀʱ���������Ϊ׼
                GetDiagnosis = True
                Exit Function
            End If
        End If
    ElseIf strSearch = mstrOldName Then
'        Call zlCommFun.PressKey(vbKeyTab)
    Else
        strSex = zlCommFun.GetNeedName(txt�Ա�.Text)
        
        vRect = zlControl.GetControlRect(vsGrid.hWnd)
        lngTop = vRect.Top + vsGrid.CellTop + vsGrid.CellHeight
        Set rsTemp = GetDiseaseCode(Me, blnCancel, strInput, strSex, str���, vRect.Left - 15, lngTop, lngHeigth)
        If Not rsTemp Is Nothing Then
            i = 1
            With rsTemp
                If UCase(TypeName(vsGrid)) = "VSFLEXGRID" Then
                    With vsGrid
                        lng���Id = IIf(IsNull(rsTemp!ID), 0, rsTemp!ID)
                        If Not ExamineInputRepeat(vsGrid, lng���Id, intFlag, .Row) Then
                           
                            If Not IsNull(.TextMatrix(j, .ColIndex("�������"))) And .TextMatrix(j, .ColIndex("�������")) <> "" Then
                                If i <> 1 Then
                                    .Row = .Rows - 1
                                    If .Row + 1 = .Rows Then .Rows = .Rows + 1
                                End If
                            Else
                                If .Row + 1 = .Rows Then .Rows = .Rows + 1
                            End If
                
                            If intFlag = 1 Then
                                .TextMatrix(.Row, .ColIndex("ICD����")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                            Else
                                .TextMatrix(.Row, .ColIndex("��ҽ����")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                                '50337:������,2012-09-18,�����ID�и�ֵ����Ȼ�����ж�����ظ�
                                .TextMatrix(.Row, .ColIndex("���ID")) = IIf(IsNull(rsTemp!ID), 0, rsTemp!ID)
                            End If
                            .TextMatrix(.Row, .ColIndex("����ID")) = IIf(IsNull(rsTemp!ID), 0, rsTemp!ID)
                            .TextMatrix(.Row, .ColIndex("�������")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                            If .Row + 1 = .Rows Then .Rows = .Rows + 1
                '                    .Row = .Row + 1
                
                            If intFlag = 1 Then
                                .Select .Row, .ColIndex("ICD����")
                            Else
                                .Select .Row, .ColIndex("��ҽ����")
                            End If
                        Else
                            .TextMatrix(.Row, .ColIndex("�������")) = mstrOldName
                        End If
                    End With
                Else
                    .Close
                    If vsGrid.Enabled Then vsGrid.SetFocus
                    zlCommFun.PressKey vbKeyTab
                End If
                .Close
            End With
        Else
            If Not blnCancel Then
                MsgBox "û���ҵ�ƥ��ļ������롣", vbInformation, gstrSysName
            End If
            If vsGrid.Enabled Then
                GetDiagnosis = False
                Exit Function
            End If
        End If
    End If
    GetDiagnosis = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ExamineInputRepeat(ByRef vsGrid As VSFlexGrid, ByVal lng���Id As Long, ByVal intFlag As Integer, ByVal CurrRow As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '����:��֤�����ݵ��Ƿ����ظ�
    '����:���ظ�����true,����false
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
           
    ExamineInputRepeat = True
    
    With vsGrid
        For i = 1 To .Rows - 1
            If i <> CurrRow Then
                If intFlag = 1 Then
                    If Val(.TextMatrix(i, .ColIndex("����ID"))) = lng���Id Then
                        MsgBox "¼��������б��е�" & i & "�������ͬ����¼��������ϵ����ݣ�", vbInformation, gstrSysName
                        Exit Function
                    End If
                Else
                    If Val(.TextMatrix(i, .ColIndex("���ID"))) = lng���Id Then
                        MsgBox "¼��������б��е�" & i & "�������ͬ����¼��������ϵ����ݣ�", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        Next
    End With
    ExamineInputRepeat = False
End Function

Private Sub vfg��ҽ_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strValue As String
    With vfg��ҽ
        Select Case Col
           Case .ColIndex("�������")
                strValue = Trim(.TextMatrix(Row, .ColIndex("�������")))
                If Not IsNull(strValue) Then
                    If Not GetDiagnosis(vfg��ҽ, strValue, 2) Then
                        .Select Row, Col
                        .TextMatrix(Row, .ColIndex("�������")) = mstrOldName
                        mstrOldName = ""
                        Exit Sub
                    End If
                    Call SetVfgNo(vfg��ҽ)
                End If
            Case .ColIndex("��Ժ���")
                If .ComboIndex < 0 Then Exit Sub
                .Cell(flexcpData, Row, Col) = .ComboData(.ComboIndex)
                .TextMatrix(Row, Col) = .ComboItem(.ComboIndex)
        End Select
    End With
End Sub

Private Sub vfg��ҽ_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vfg��ҽ.Editable = flexEDNone Then
        Cancel = True
        Exit Sub
    End If
    Select Case Col
        Case vfg��ҽ.ColIndex("�������")
            mstrOldName = Trim(vfg��ҽ.TextMatrix(Row, vfg��ҽ.ColIndex("�������")))
            Cancel = False
            Exit Sub
        Case vfg��ҽ.ColIndex("��Ժ���") ', vfg��ҽ.ColIndex("����")
            Cancel = False
            Exit Sub
        Case Else
            Cancel = True
            Exit Sub
    End Select
End Sub

Private Sub vfg��ҽ_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    Select Case Col
        Case vfg��ҽ.ColIndex("���")
            Position = -1
            Exit Sub
    End Select
    If Position = 0 Then
        Position = Col
    End If
End Sub

Private Sub vfg��ҽ_BeforeSort(ByVal Col As Long, Order As Integer)
    Call zl_VsGridBeforeSort(vfg��ҽ, Col, Order)
    Call SetVfgNo(vfg��ҽ)
End Sub

Private Sub vfg��ҽ_EnterCell()
    Dim strTemp As String
    Dim intCount As Integer
    Dim strKey As String
    Dim strValue As String
    
    If vfg��ҽ.Editable = flexEDKbdMouse Then
        With vfg��ҽ
            If .Row > 0 Then
                If .Col = .ColIndex("��Ժ���") Then
                    strTemp = .TextMatrix(.Row, .Col)
                    strKey = .ColData(.ColIndex("��Ժ���"))
                    strValue = Trim(.TextMatrix(.Row, .ColIndex("�������")))
                    If strTemp = "" And strValue <> "" Then
                        .TextMatrix(.Row, .Col) = Mid(strKey, InStr(1, strKey, ";") + 1)
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub vfg��ҽ_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
     If vfg��ҽ.Editable = flexEDKbdMouse Then
        If KeyCode = vbKeyDelete Then
            If vfg��ҽ.Row > 0 Then
                If MsgBox("��Ҫɾ����ǰ��¼��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                    vfg��ҽ.RemoveItem (vfg��ҽ.Row)
                    If vfg��ҽ.Row = 0 Then
                        vfg��ҽ.Rows = vfg��ҽ.Rows + 1
                        vfg��ҽ.Select vfg��ҽ.Rows - 1, vfg��ҽ.Col
                    End If
                End If
            End If
        End If
        
        If KeyCode = vbKeyInsert Then
            With vfg��ҽ
                If MsgBox("��Ҫ���Ӽ�¼��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                   vfg��ҽ.Rows = vfg��ҽ.Rows + 1
                   .Select vfg��ҽ.Rows - 1, vfg��ҽ.Col
                End If
            End With
        End If
    End If
    If KeyCode = vbKeyReturn Then
        lngRow = vfg��ҽ.Row
        If vfg��ҽ.Editable = flexEDKbdMouse Then ''�������,2200,1,1;��Ժ���,1000,1,1;����
            Call zlPvVsMoveGridCell(vfg��ҽ, vfg��ҽ.ColIndex("�������"), vfg��ҽ.ColIndex("��Ժ���"), True, lngRow, SetHeadCodeData(vfg��ҽ))
        Else
            Call zlPvVsMoveGridCell(vfg��ҽ, vfg��ҽ.ColIndex("�������"), vfg��ҽ.ColIndex("��Ժ���"), False, lngRow, SetHeadCodeData(vfg��ҽ))
        End If
    End If
    Call SetVfgNo(vfg��ҽ)
End Sub

Private Sub vfg��ҽ_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim lngRow As Long
    If KeyCode = vbKeyReturn Then
        lngRow = Row
        If vfg��ҽ.Editable = flexEDKbdMouse Then
            Call zlPvVsMoveGridCell(vfg��ҽ, vfg��ҽ.ColIndex("�������"), vfg��ҽ.ColIndex("��Ժ���"), True, lngRow, SetHeadCodeData(vfg��ҽ))
        Else
            Call zlPvVsMoveGridCell(vfg��ҽ, vfg��ҽ.ColIndex("�������"), vfg��ҽ.ColIndex("��Ժ���"), False, lngRow, SetHeadCodeData(vfg��ҽ))
        End If
    End If
End Sub

Private Sub vfg��ҽ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0 'Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub vfg��ҽ_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
        Case vbKeyBack, vbKeyEscape, 3, 22: Exit Sub
        Case Else
'            Select Case Col
'                Case vfgInDetail.ColIndex("����")
'                    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
'                        Exit Sub
'                    Else
'                        KeyAscii = 0
'                    End If
''            End Select

    End Select
End Sub
'����28612 by lesfeng 2010-07-05
Private Function GetdeathTime(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Date
'���ܣ���ȡָ�������Ƿ��������ҽ�������ڳ�Ժʱ��Ϊ����ʱ���1��
'˵�������ڻ�ȡ��������ʱ��Ϊ��Ժʱ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    GetdeathTime = CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"))
    
    On Error GoTo errH
    '47955:������,2012-09-18,�ı������е�A.���=0ΪA.Ӥ��=0
    '59094:������,2013-04-24,�޸�Ϊֻ��1s,ԭ��Ϊ1m
    strSQL = "Select Max(Nvl(A.ִ����ֹʱ��, Nvl(A.�ϴ�ִ��ʱ��, A.��ʼִ��ʱ��)) + 1 / 24 / 60 / 60 ) As ʱ�� " & _
             "  From ����ҽ����¼ A, ������ĿĿ¼ B " & _
             " Where A.������� = B.��� And A.������Ŀid = B.ID And B.�������� = 11 And B.��� = 'Z' And A.ҽ��״̬ In (3, 8, 9) And nvl(A.Ӥ��,0)=0 And " & _
             "       A.����ID = [1] And A.��ҳID = [2] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
    
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!ʱ��) Then
            GetdeathTime = rsTmp!ʱ��
            mintDeath = 1
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ShowMe(frmParent As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strPrivs As String) As Boolean
    On Error Resume Next
    Set mfrmParent = frmParent
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mstrPrivs = strPrivs
    
    Me.Show IIf(gfrmMain Is Nothing, 0, 1), frmParent
    
    ShowMe = gblnOK
End Function
