VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmModifOut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�µĳ�Ժʱ��"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8520
   Icon            =   "frmModifOut.frx":0000
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
      TabIndex        =   16
      Top             =   795
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7185
      TabIndex        =   15
      Top             =   360
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   7185
      TabIndex        =   17
      Top             =   4710
      Width           =   1100
   End
   Begin VB.Frame fraInfo 
      Height          =   5535
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   7035
      Begin VB.TextBox txt���ﵥλ 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   6480
         MaxLength       =   3
         TabIndex        =   31
         Top             =   4440
         Width           =   405
      End
      Begin VB.ComboBox cbo��Ժ��� 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         Left            =   5550
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   660
         Width           =   1350
      End
      Begin VB.CheckBox chk���� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "����"
         Enabled         =   0   'False
         Height          =   195
         Left            =   4875
         TabIndex        =   12
         Top             =   4500
         Width           =   660
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   6000
         MaxLength       =   3
         TabIndex        =   13
         Top             =   4440
         Width           =   405
      End
      Begin VB.CheckBox chkʬ�� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "ʬ��"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2910
         TabIndex        =   10
         Top             =   4980
         Width           =   660
      End
      Begin VB.TextBox txtסԺ�� 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   1290
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4170
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   690
      End
      Begin VB.TextBox txt�Ա� 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   675
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txt��Ժ��� 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         Left            =   960
         MaxLength       =   200
         TabIndex        =   4
         Top             =   660
         Width           =   3660
      End
      Begin VB.CheckBox chk���� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "ȷ��"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2235
         TabIndex        =   9
         Top             =   4500
         Width           =   660
      End
      Begin VB.TextBox txt��ҽ��� 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         Left            =   960
         MaxLength       =   200
         TabIndex        =   6
         Top             =   2580
         Width           =   3660
      End
      Begin VB.ComboBox cbo��ҽ��Ժ��� 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         Left            =   5550
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2580
         Width           =   1350
      End
      Begin VB.ComboBox cbo��Ժ��ʽ 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   4440
         Width           =   1230
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   300
         Left            =   960
         TabIndex        =   11
         Top             =   4950
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   14737632
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   19
         Format          =   "yyyy-MM-dd hh:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtNewDate 
         Height          =   300
         Left            =   4920
         TabIndex        =   14
         Top             =   4950
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   529
         _Version        =   393216
         ForeColor       =   16711680
         AutoTab         =   -1  'True
         MaxLength       =   19
         Format          =   "yyyy-MM-dd hh:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VSFlex8Ctl.VSFlexGrid vfg��ҽ 
         Height          =   1455
         Left            =   960
         TabIndex        =   32
         Top             =   1080
         Width           =   5895
         _cx             =   10398
         _cy             =   2566
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
         TabIndex        =   34
         Top             =   3000
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
         Left            =   3000
         TabIndex        =   36
         Top             =   4440
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   14737632
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
         TabIndex        =   35
         Top             =   3045
         Width           =   720
      End
      Begin VB.Label lbl��ҽ���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   180
         TabIndex        =   33
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������Ժʱ��"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   3720
         TabIndex        =   30
         Top             =   5010
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ���"
         Height          =   180
         Left            =   4755
         TabIndex        =   29
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժʱ��"
         Height          =   180
         Left            =   180
         TabIndex        =   28
         Top             =   5010
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   5580
         TabIndex        =   27
         Top             =   4500
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   3765
         TabIndex        =   26
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   2385
         TabIndex        =   25
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   525
         TabIndex        =   24
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "סԺ��"
         Height          =   180
         Left            =   4920
         TabIndex        =   23
         Top             =   300
         Width           =   540
      End
      Begin VB.Label lbl��Ժ��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ���"
         Height          =   180
         Left            =   180
         TabIndex        =   22
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lbl��ҽ��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ҽ���"
         Height          =   180
         Left            =   180
         TabIndex        =   21
         Top             =   2640
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ���"
         Height          =   180
         Left            =   4755
         TabIndex        =   20
         Top             =   2640
         Width           =   720
      End
      Begin VB.Label lbl��Ժ��ʽ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ��ʽ"
         Height          =   180
         Left            =   180
         TabIndex        =   19
         Top             =   4500
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmModifOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Public mstrPrivs As String
Public mlng����ID As Long, mlng��ҳID As Long
Private mintĬ����� As Integer
Private mrsPatiInfo As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, strSQL As String, str���� As String
    Dim dMax As Date, intԭ�� As Integer, int���� As Integer
    Dim rsDiagnosis As ADODB.Recordset
    Dim str��Ժ��� As String, str��ҽ��Ժ��� As String, strTmp As String
    Dim str��Ժ��ʽ As String
    Dim int�����־ As Integer
    
    
    
    gblnOK = False
    Set mrsPatiInfo = GetPatiInfoModiOut(mlng����ID, mlng��ҳID)
    mintĬ����� = Val(zlDatabase.GetPara("Ĭ�����", glngSys, glngModul))
    If mrsPatiInfo.EOF Then
        MsgBox "���˳�Ժ��Ϣ�����ڣ���˲� ��", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    With mrsPatiInfo
        txt����.Text = !����
        txt�Ա�.Text = "" & !�Ա�
        txt����.Text = "" & !����
        txtסԺ��.Text = "" & !סԺ��
        txtDate.Text = "" & !��Ժ����
        txt��ҽ���.Enabled = (InStr(1, "," & GetDepCharacter(Val("" & !��Ժ����id)) & ",", ",��ҽ��,") > 0)
        txt��ҽ���.ToolTipText = "ֻ�е��������ڿ��ҵ�����Ϊ��ҽ��ʱ������������ҽ���!"
        cbo��ҽ��Ժ���.Enabled = txt��ҽ���.Enabled
        '����28982 by lesfeng 2010-06-09
        chk����.Value = IIf(IsNull(!�Ƿ�ȷ��), 0, IIf(!�Ƿ�ȷ�� = 1, 1, 0))
        chkʬ��.Value = IIf(IsNull(!ʬ���־), 0, IIf(!ʬ���־ = 1, 1, 0))
        chk����.Value = IIf(IsNull(!�����־), 0, IIf(!�����־ >= 1, 1, 0))
        txt����.Text = IIf(IsNull(!��������), "", !��������)
        int�����־ = IIf(IsNull(!�����־), 0, !�����־)
        '����28982 by lesfeng 2010-06-09
        If chk����.Value = 1 And Not IsNull(!ȷ������) Then txtOkDate.Text = IIf(IsNull(!ȷ������), "3000-01-01 00:00:00", Format(!ȷ������, "yyyy-MM-dd HH:mm:ss"))
        Select Case int�����־
            Case 0
                txt���ﵥλ.Text = ""
            Case 1
                txt���ﵥλ.Text = "��"
            Case 2
                txt���ﵥλ.Text = "��"
            Case 3
                txt���ﵥλ.Text = "��"
            Case 4
                txt���ﵥλ.Text = "��"
            Case 9
                txt���ﵥλ.Text = "����"
        End Select
    End With
    
    txtNewDate.Text = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    dMax = GetMaxOutDate(mlng����ID, mlng��ҳID, intԭ��)
    If intԭ�� = 10 Then
        '59094:������,2013-04-24,�޸�Ϊֻ��1s,ԭ��Ϊ1m
        txtNewDate.Text = Format(dMax + 1 / 24 / 60 / 60, "yyyy-MM-dd HH:mm:ss")
    Else
        If dMax > CDate(txtNewDate.Text) Then
            txtNewDate.Text = Format(dMax + 1 / 24 / 60, "yyyy-MM-dd HH:mm:ss")
        End If
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
    str��Ժ��ʽ = IIf(IsNull(mrsPatiInfo!��Ժ��ʽ), "", mrsPatiInfo!��Ժ��ʽ)
    strSQL = "Select ����,����,Nvl(ȱʡ��־,0) as ȱʡ From ��Ժ��ʽ Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo��Ժ��ʽ.AddItem rsTmp!���� & "-" & rsTmp!����
            If str��Ժ��ʽ = "" Then
                If rsTmp!ȱʡ = 1 Then cbo��Ժ��ʽ.ListIndex = cbo��Ժ��ʽ.NewIndex
            Else
                '����31294 by lesfeng 2010-07-07 rsTmp!���� ��Ϊ rsTmp!����
                If rsTmp!���� = str��Ժ��ʽ Then cbo��Ժ��ʽ.ListIndex = cbo��Ժ��ʽ.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If
    '����28139 by lesfeng 2010-03-02
    Call LoadVfgData(vfg��ҽ, 1)
    Call LoadVfgData(vfg��ҽ, 2)
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
    
    On Error GoTo errH
    
    If Not IsDate(txtNewDate.Text) Then
        MsgBox "��������ȷ�Ĳ����µĳ�Ժʱ�䣡", vbInformation, gstrSysName
        txtNewDate.SetFocus: Exit Sub
    End If
    
    'ʱ�䲻�ܳ�����ǰʱ��̫��(һ��)
    Curdate = zlDatabase.Currentdate
    If CDate(txtNewDate.Text) > Curdate Then
        If CDate(txtNewDate.Text) - Curdate > 7 Then
            MsgBox "��Ժʱ��ȵ�ǰʱ���ù���,���飡", vbInformation, gstrSysName
            txtNewDate.SetFocus: Exit Sub
        End If
        If MsgBox("��Ժʱ������˵�ǰϵͳʱ��,ȷʵҪ�޸ĳ�Ժʱ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txtNewDate.SetFocus: Exit Sub
        End If
    End If
       
    dMax = GetMaxOutDate(mlng����ID, mlng��ҳID)
    If Format(txtNewDate.Text, "yyyyMMddHHmmss") <= Format(dMax, "yyyyMMddHHmmss") Then
        MsgBox "���˳�Ժʱ�������ڸò����ϴα䶯ʱ�� " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & " ��", vbInformation, gstrSysName
        txtNewDate.SetFocus: Exit Sub
    End If
    
    dMax = GetLastAdviceTime(mlng����ID, mlng��ҳID)
    If Format(txtNewDate.Text, "yyyyMMddHHmmss") < Format(dMax, "yyyyMMddHHmmss") Then
        If MsgBox("��Ժʱ��С�ڸò��������Чҽ����ʱ�� " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & ",ȷʵ��Ҫ�޸ĳ�Ժʱ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txtNewDate.SetFocus: Exit Sub
        End If
    End If
             
    strSQL = "Zl_���˱䶯��¼_ModifOut(" & mlng����ID & "," & mlng��ҳID & ",To_Date('" & txtNewDate.Text & "','YYYY-MM-DD HH24:MI:SS')," & _
        "'" & UserInfo.��� & "','" & UserInfo.���� & "')"
''
    gcnOracle.BeginTrans
        blnTrans = True
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    gcnOracle.CommitTrans
    blnTrans = False
''
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

Private Sub txtNewDate_GotFocus()
    zlControl.TxtSelAll txtNewDate
End Sub

Private Sub txtNewDate_LostFocus()
    If Not IsDate(txtNewDate.Text) Then txtNewDate.SetFocus
End Sub

Private Function GetMaxOutDate(lng����ID As Long, lng��ҳID As Long, Optional intԭ�� As Integer) As Date
'���ܣ���ȡת�Ʋ��������ϴα䶯ʱ��
'������intԭ��=�����ϴα䶯��ԭ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    GetMaxOutDate = #1/1/1900#
    intԭ�� = 0
    
    strSQL = "Select ��ʼʱ��,��ʼԭ�� From ���˱䶯��¼" & _
        " Where ��ʼʱ�� is Not NULL And ��ֹʱ�� is not  NULL AND ��ֹԭ�� = 1 " & _
        " And ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
    If Not rsTmp.EOF Then
        GetMaxOutDate = IIf(IsNull(rsTmp!��ʼʱ��), GetMaxOutDate, rsTmp!��ʼʱ��)
        intԭ�� = Nvl(rsTmp!��ʼԭ��, 0)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
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
        .Editable = flexEDNone
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
                    rsDiagnosisOther.Filter = "�������=3 and ��¼��Դ=2"        '��ȡ��Ժ�Ǽǵĳ�Ժ���
                    If Not rsDiagnosisOther.EOF Then
                        .Rows = IIf(rsDiagnosisOther.EOF, 0, rsDiagnosisOther.RecordCount) + 1
                    Else
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
            Else
                'b.��ҽ���
                rsDiagnosisOther.Filter = "�������=13 and ��¼��Դ=3"            '��ȡ��ҳ����ĳ�Ժ���
                If Not rsDiagnosisOther.EOF Then
                    .Rows = IIf(rsDiagnosisOther.EOF, 0, rsDiagnosisOther.RecordCount) + 1
                Else
                    rsDiagnosisOther.Filter = "�������=13 and ��¼��Դ=2"        '��ȡ��Ժ�Ǽǵĳ�Ժ���
                    If Not rsDiagnosisOther.EOF Then
                        .Rows = IIf(rsDiagnosisOther.EOF, 0, rsDiagnosisOther.RecordCount) + 1
                    Else
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
    
            Else
                .Rows = .Rows + 1
            End If
            
            If .Rows > 1 Then
                .Select 1, .ColIndex("��Ժ���")
            End If
        End If
    End With
    Call SetVfgNo(vsGrid)
    Call SetInitVfgFormat(vsGrid, intFlag)
    Call RestoreHead(vsGrid, intFlag)
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
End Sub

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

Public Function ShowMe(frmParent As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strPrivs As String) As Boolean
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mstrPrivs = strPrivs
    
    Me.Show IIf(gfrmMain Is Nothing, 0, 1), frmParent
    
    ShowMe = gblnOK
End Function
