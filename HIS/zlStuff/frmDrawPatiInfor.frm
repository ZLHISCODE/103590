VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmDrawPatiInfor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���Ĳ��˸�����Ϣ"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frmDrawPatiInfor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VSFlex8Ctl.VSFlexGrid vsfInfo 
      Height          =   2295
      Left            =   930
      TabIndex        =   46
      Top             =   1920
      Visible         =   0   'False
      Width           =   7335
      _cx             =   12938
      _cy             =   4048
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   18
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDrawPatiInfor.frx":000C
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
   Begin VB.CheckBox chk���Կ��� 
      Caption         =   "���Բ������ڿ��һ���"
      Height          =   180
      Left            =   4200
      TabIndex        =   45
      Top             =   360
      Width           =   3855
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   17
      Left            =   6465
      Locked          =   -1  'True
      TabIndex        =   34
      Tag             =   "סԺ��"
      Top             =   3405
      Width           =   1590
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   16
      Left            =   4260
      Locked          =   -1  'True
      TabIndex        =   32
      Tag             =   "�����"
      Top             =   3405
      Width           =   1050
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   15
      Left            =   930
      Locked          =   -1  'True
      TabIndex        =   30
      Tag             =   "��ǰ����"
      Top             =   3405
      Width           =   1395
   End
   Begin VB.Frame fra 
      Height          =   60
      Index           =   1
      Left            =   30
      TabIndex        =   44
      Top             =   4590
      Width           =   9000
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   14
      Left            =   930
      Locked          =   -1  'True
      TabIndex        =   38
      Tag             =   "��ǰ����"
      Top             =   4245
      Width           =   7140
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   13
      Left            =   930
      Locked          =   -1  'True
      TabIndex        =   36
      Tag             =   "��ǰ����"
      Top             =   3840
      Width           =   7140
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   12
      Left            =   930
      Locked          =   -1  'True
      TabIndex        =   26
      Tag             =   "����"
      Top             =   2985
      Width           =   1395
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   11
      Left            =   6465
      Locked          =   -1  'True
      TabIndex        =   24
      Tag             =   "����״��"
      Top             =   2550
      Width           =   1605
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   10
      Left            =   4245
      Locked          =   -1  'True
      TabIndex        =   22
      Tag             =   "���"
      Top             =   2550
      Width           =   1035
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   9
      Left            =   930
      Locked          =   -1  'True
      TabIndex        =   20
      Tag             =   "ѧ��"
      Top             =   2550
      Width           =   1395
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   8
      Left            =   4245
      Locked          =   -1  'True
      TabIndex        =   16
      Tag             =   "����"
      Top             =   2115
      Width           =   1035
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   7
      Left            =   6465
      Locked          =   -1  'True
      TabIndex        =   18
      Tag             =   "���֤��"
      Top             =   2115
      Width           =   1605
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   6
      Left            =   4245
      Locked          =   -1  'True
      TabIndex        =   28
      Tag             =   "�����ص�"
      Top             =   2985
      Width           =   3825
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   5
      Left            =   930
      Locked          =   -1  'True
      TabIndex        =   14
      Tag             =   "��������"
      Top             =   2115
      Width           =   1395
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   4
      Left            =   6465
      Locked          =   -1  'True
      TabIndex        =   12
      Tag             =   "����"
      Top             =   1680
      Width           =   1605
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   3
      Left            =   4245
      Locked          =   -1  'True
      TabIndex        =   10
      Tag             =   "�Ա�"
      Top             =   1680
      Width           =   1035
   End
   Begin VB.CommandButton cmdPati 
      Caption         =   "��"
      Height          =   300
      Left            =   3090
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1680
      Width           =   270
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   1
      Left            =   930
      MaxLength       =   20
      TabIndex        =   3
      Tag             =   "��������"
      Top             =   1260
      Width           =   2430
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   0
      Left            =   930
      Locked          =   -1  'True
      TabIndex        =   1
      Tag             =   "������Ϣ"
      Top             =   810
      Width           =   7140
   End
   Begin VB.Frame fra 
      Height          =   60
      Index           =   0
      Left            =   -30
      TabIndex        =   43
      Top             =   645
      Width           =   9000
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   90
      Picture         =   "frmDrawPatiInfor.frx":0258
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   4785
      Width           =   1100
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   6975
      TabIndex        =   40
      Top             =   4785
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   -405
      Top             =   6090
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawPatiInfor.frx":03A2
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawPatiInfor.frx":093C
            Key             =   "ItemStop"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawPatiInfor.frx":0ED6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5880
      TabIndex        =   39
      Top             =   4785
      Width           =   1100
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   2
      Left            =   930
      TabIndex        =   7
      Tag             =   "����"
      Top             =   1680
      Width           =   2160
   End
   Begin MSMask.MaskEdBox MakTxtEdit 
      Height          =   300
      Left            =   6480
      TabIndex        =   5
      Top             =   1260
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   10
      Format          =   "yyyy-MM-DD"
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "סԺ��"
      Height          =   180
      Index           =   17
      Left            =   5880
      TabIndex        =   33
      Top             =   3465
      Width           =   540
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "�����"
      Height          =   180
      Index           =   16
      Left            =   3675
      TabIndex        =   31
      Top             =   3465
      Width           =   540
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "��ǰ����"
      Height          =   180
      Index           =   15
      Left            =   150
      TabIndex        =   29
      Top             =   3465
      Width           =   720
   End
   Begin VB.Label lblEdit 
      Caption         =   "��ǰ����"
      Height          =   210
      Index           =   14
      Left            =   120
      TabIndex        =   37
      Top             =   4290
      Width           =   765
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "��ǰ����"
      Height          =   180
      Index           =   13
      Left            =   120
      TabIndex        =   35
      Top             =   3900
      Width           =   720
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   12
      Left            =   480
      TabIndex        =   25
      Top             =   3045
      Width           =   360
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����״��"
      Height          =   180
      Index           =   11
      Left            =   5685
      TabIndex        =   23
      Top             =   2610
      Width           =   720
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "���"
      Height          =   180
      Index           =   10
      Left            =   3780
      TabIndex        =   21
      Top             =   2610
      Width           =   360
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "ѧ��"
      Height          =   180
      Index           =   9
      Left            =   480
      TabIndex        =   19
      Top             =   2610
      Width           =   360
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   8
      Left            =   3780
      TabIndex        =   15
      Top             =   2175
      Width           =   360
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "���֤��"
      Height          =   180
      Index           =   7
      Left            =   5685
      TabIndex        =   17
      Top             =   2175
      Width           =   720
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "�����ص�"
      Height          =   180
      Index           =   6
      Left            =   3420
      TabIndex        =   27
      Top             =   3045
      Width           =   720
   End
   Begin VB.Label lblEditDate 
      AutoSize        =   -1  'True
      Caption         =   "ʹ��ʱ��"
      Height          =   180
      Left            =   5685
      TabIndex        =   4
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   5
      Left            =   150
      TabIndex        =   13
      Top             =   2175
      Width           =   720
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   4
      Left            =   6045
      TabIndex        =   11
      Top             =   1740
      Width           =   360
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "�Ա�"
      Height          =   180
      Index           =   3
      Left            =   3780
      TabIndex        =   9
      Top             =   1740
      Width           =   360
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   2
      Left            =   150
      TabIndex        =   6
      Top             =   1740
      Width           =   720
   End
   Begin VB.Label lblEdit 
      Caption         =   "��������"
      Height          =   210
      Index           =   1
      Left            =   150
      TabIndex        =   2
      Top             =   1305
      Width           =   765
   End
   Begin VB.Label lblEdit 
      Caption         =   "������Ϣ"
      Height          =   210
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   855
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   330
      Picture         =   "frmDrawPatiInfor.frx":2BE0
      Top             =   105
      Width           =   480
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    ������ָ���������ϵĸ�����Ϣ."
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   705
      TabIndex        =   41
      Top             =   315
      Width           =   2970
   End
End
Attribute VB_Name = "frmDrawPatiInfor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnFirst As Boolean
Private mbln�༭ As Boolean
Private mlng�շ�ID As Long
Private mlng����ID As Long                  '����ID
Private mlng����id As Long                  '����ID
Private mlng��ǰ����ID As Long
Private mstrʹ��ʱ�� As String                  'ʹ��ʱ��:yyyy-mm-dd
Private mstr���� As String                  '�������ϵ�����
Private mstr���� As String
Private mblnChange As Boolean
Private mblnOk As Boolean
Private mstr�������� As String

Private Sub cmdClose_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub


Private Function MulitSelectPati(ByVal strKey As String) As Boolean
    '----------------------------------------------------------------------------------
    '����:ѡ�����ò����µĲ�����Ϣ
    '����:strKey-ѡ��Ĳ���ID(-),סԺ��(+),����
    '����:���ѡ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/08/20
    '----------------------------------------------------------------------------------
    
    Dim strSearchKey As String, strWhere As String
    Dim rsTemp As ADODB.Recordset, blnCancel As Boolean
    Dim lngH As Long
    Dim vRect  As RECT
    
    strWhere = ""
    strSearchKey = ""
    If strKey <> "" Then
        If Not IsNumeric(Mid(strKey, 2)) Then
            strWhere = " And  a.���� like [2]"
            strSearchKey = GetMatchingSting(strKey)
        Else
            Select Case Mid(strKey, 1, 1)
            Case "-"  '����Ĳ���ID
                strWhere = " And A.����id=[1]"
                strSearchKey = Mid(strKey, 2)
            Case "+"  '�����סԺ��
                strWhere = " And B.סԺ��=[1]  "
                strSearchKey = Mid(strKey, 2)
            Case Else   '����ģ������
                strWhere = " And  a.���� like [2]"
                strSearchKey = GetMatchingSting(strKey)
            End Select
        End If
    End If
    
    If strKey <> "" Then
        If Mid(strKey, 1, 1) = "-" Then
            '������Ϣ,���ܴ�������
            gstrSQL = "" & _
                "   Select Distinct decode(m.����,NULL,NULL,'['||m.����||']'||m.����) ��ǰ����,decode(C.����,NULL,NULL,'['||c.����||']'||c.����) as ��ǰ����,A.����ID As ID,A.����,A.�Ա�, A.����,to_char(A.��������,'yyyy-mm-dd') ��������,A.����,A.���֤��,A.ѧ��,A.���,A.����״��," & _
                "         A.����,A.�����ص�, a.��ǰ���� As ��ǰ����,A.�����,A.סԺ��" & _
                "   From  ������Ϣ A,���ű� C,���ű� M " & _
                "   Where  A.��ǰ����ID=C.id(+) and  a.��ǰ����ID=M.id(+) " & _
                "           " & strWhere
                If mlng��ǰ����ID <> 0 And chk���Կ���.Value = 0 Then
                    '����:13415
                    If mstr�������� <> "����" Then
                        gstrSQL = gstrSQL & " And A.��ǰ����ID=[3]"
                    Else
                        '�ٴ�
                        gstrSQL = gstrSQL & " And A.��ǰ����ID=[3]"
                    End If
                End If
        Else
            '������Ϣ
            gstrSQL = "" & _
                "   Select Distinct decode(m.����,NULL,NULL,'['||m.����||']'||m.����) ��ǰ����,decode(C.����,NULL,NULL,'['||c.����||']'||c.����) as ��ǰ����,A.����ID As ID,A.����,A.�Ա�, b.����,to_char(A.��������,'yyyy-mm-dd') ��������,A.����,A.���֤��,A.ѧ��,A.���,A.����״��," & _
                "         A.����,A.�����ص�, B.��Ժ���� As ��ǰ����,A.�����,A.סԺ��" & _
                "   From ������ҳ B, ������Ϣ A,���ű� C,���ű� M " & _
                "   Where A.����id = B.����id And A.��ҳid=B.��ҳid and B.��Ժ����ID=C.id(+) and  a.��ǰ����ID=M.id(+) and B.��Ժ���� Is Not Null  " & _
                "           " & strWhere
                If mlng��ǰ����ID <> 0 And chk���Կ���.Value = 0 Then
                    '����:13415
                    If mstr�������� <> "����" Then
                        gstrSQL = gstrSQL & " And B.��Ժ����ID=[3]"
                    Else
                        '�ٴ�
                        gstrSQL = gstrSQL & " And B.��ǰ����ID=[3]"
                    End If
                End If
        End If
    Else
        '������Ϣ
        gstrSQL = "" & _
            "   Select Distinct Decode(M.����,NULL,NULL,'['||m.����||']'||m.����) ��ǰ����,decode(C.����,NULL,NULL,'['||c.����||']'||c.���� ) as ��ǰ����,A.����ID As ID,A.����,A.�Ա�, b.����,to_char(A.��������,'yyyy-mm-dd') ��������,A.����,A.���֤��,A.ѧ��,A.���,A.����״��," & _
            "         A.����,A.�����ص�, B.��Ժ���� As ��ǰ����,A.�����,A.סԺ��" & _
            "   From ������ҳ B, ������Ϣ A,���ű� C,���ű� M " & _
            "   Where A.����id = B.����id And A.��ҳid=B.��ҳid and B.��Ժ����ID=C.id(+) and  a.��ǰ����ID=M.id(+) And B.��Ժ���� Is Not Null " & _
            "          " & strWhere
                If mlng��ǰ����ID <> 0 And chk���Կ���.Value = 0 Then
                    '����:13415
                    If mstr�������� <> "����" Then
                        gstrSQL = gstrSQL & " And B.��Ժ����ID=[3]"
                    Else
                        '�ٴ�
                        gstrSQL = gstrSQL & " And B.��ǰ����ID=[3]"
                    End If
                End If
    End If
    vRect = zlControl.GetControlRect(txtEdit(2).hwnd)
    lngH = txtEdit(2).Height
    
'    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "����ѡ����", False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strSearchKey, CStr(UCase(strSearchKey)), mlng��ǰ����ID)
'    If blnCancel = True Then
'        If txtEdit(2).Enabled Then txtEdit(2).SetFocus
'        Exit Function
'    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "������Ϣ", strSearchKey, CStr(UCase(strSearchKey)), mlng��ǰ����ID)
    
    If rsTemp.RecordCount = 0 Then
        ShowMsgBox "û�����������Ĳ�����Ϣ,����!"
        If txtEdit(2).Enabled Then txtEdit(2).SetFocus
        Exit Function
    Else
        vsfInfo.Visible = True
        vsfInfo.Rows = 1
        Set vsfInfo.DataSource = rsTemp
        With vsfInfo
            Do While rsTemp.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, .ColIndex("��ǰ����")) = rsTemp!��ǰ����
                .TextMatrix(.Rows - 1, .ColIndex("��ǰ����")) = rsTemp!��ǰ����
                .TextMatrix(.Rows - 1, .ColIndex("Id")) = rsTemp!Id
                .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTemp!����
                .TextMatrix(.Rows - 1, .ColIndex("�Ա�")) = rsTemp!�Ա�
                .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTemp!����
                .TextMatrix(.Rows - 1, .ColIndex("��������")) = rsTemp!��������
                .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTemp!����
                .TextMatrix(.Rows - 1, .ColIndex("���֤��")) = rsTemp!���֤��
                .TextMatrix(.Rows - 1, .ColIndex("ѧ��")) = rsTemp!ѧ��
                .TextMatrix(.Rows - 1, .ColIndex("���")) = rsTemp!���
                .TextMatrix(.Rows - 1, .ColIndex("����״��")) = rsTemp!����״��
                .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTemp!����
                .TextMatrix(.Rows - 1, .ColIndex("�����ص�")) = rsTemp!�����ص�
                .TextMatrix(.Rows - 1, .ColIndex("��ǰ����")) = rsTemp!��ǰ����
                .TextMatrix(.Rows - 1, .ColIndex("�����")) = rsTemp!�����
                .TextMatrix(.Rows - 1, .ColIndex("סԺ��")) = rsTemp!סԺ��
                rsTemp.MoveNext
            Loop
        End With
    End If
'    txtEdit(2).Text = zlStr.Nvl(rsTemp!����)
'    cmdPati.Tag = zlStr.Nvl(rsTemp!Id)
'    mlng����id = Val(zlStr.Nvl(rsTemp!Id))
    Dim i As Integer
'    For i = 2 To txtEdit.UBound
'        txtEdit(i).Text = zlStr.Nvl(rsTemp.Fields(txtEdit(i).Tag))
'    Next
    
    MulitSelectPati = True
End Function
   
Private Function Init������Ϣ() As Boolean
    '------------------------------------------------------------------------------
    '����:��ʼ��������Ϣ
    '����:
    '����:��ʼ�ɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2007/08/20
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim i1 As Long
    
    On Error GoTo ErrHandle
    gstrSQL = "Select id,����,����,���,���� From �շ���ĿĿ¼ where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID)
    If rsTemp.EOF Then
        ShowMsgBox "������ָ������������,����!"
        Exit Function
    End If
    
    
    txtEdit(0).Text = "[" & zlStr.Nvl(rsTemp!����) & "]" & zlStr.Nvl(rsTemp!����) & Space(5) & zlStr.Nvl(rsTemp!���) & Space(5) & zlStr.Nvl(rsTemp!����)
    txtEdit(1).Text = mstr����
    If mstrʹ��ʱ�� = "" And mbln�༭ = True Then
        MakTxtEdit.Text = Format(sys.Currentdate, "yyyy-mm-dd")
    ElseIf mstrʹ��ʱ�� = "" Then
    Else
        MakTxtEdit.Text = mstrʹ��ʱ��
    End If
        
    gstrSQL = "Select ID,����,���� From ���ű� where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng��ǰ����ID)
    If rsTemp.EOF Then
        ShowMsgBox "δѡ�����ò���,����!"
        Exit Function
    End If
    For i1 = 0 To txtEdit.UBound
        If txtEdit(i1).Tag = "��ǰ����" Then
            txtEdit(i1).Text = "[" & zlStr.Nvl(rsTemp!����) & "]" & zlStr.Nvl(rsTemp!����)
            Exit For
        End If
    Next
    '���ز�����Ϣ
    If mlng����id <> 0 Then
        
        If mlng�շ�ID <> 0 Then
            gstrSQL = "" & _
                "   Select Distinct A.����ID As ID,Q.����,Q.�Ա�, Q.����,to_char(A.��������,'yyyy-mm-dd') ��������,A.����,A.���֤��,A.ѧ��,A.���,A.����״��," & _
                "         A.����,A.�����ص�, Q.���� As ��ǰ����,A.�����,b.סԺ��,decode(C.����,NULL,NULL,'['||c.����||']'||c.����) as ��ǰ����,decode(M.����,NULL,NULL,'['||m.����||']'||m.����) ��ǰ����, " & _
                "        to_char(Q.ʹ��ʱ��,'yyyy-mm-dd') ʹ��ʱ��,Q.����" & _
                "   From  ������Ϣ A,����������Ϣ Q,������ҳ B,���ű� C,���ű� M" & _
                "   Where A.����id = Q.����id And A.��ҳid=Q.��ҳid And q.����id = b.����id And q.��ҳid = b.��ҳid and Q.��ǰ����ID=C.id(+) and  Q.��ǰ����ID=M.id(+) " & _
                "           And Q.�շ�ID= [2] "
        Else
            gstrSQL = "" & _
                "   Select Distinct A.����ID As ID,A.����,A.�Ա�, A.����,to_char(A.��������,'yyyy-mm-dd') ��������,A.����,A.���֤��,A.ѧ��,A.���,A.����״��," & _
                "         A.����,A.�����ص�, a.��ǰ���� As ��ǰ����,A.�����,A.סԺ��,decode(C.����,NULL,NULL,'['||c.����||']'||c.����) as ��ǰ����,decode(m.����,NULL,NULL,'['||m.����||']'||m.����) ��ǰ����" & _
                "   From  ������Ϣ A,���ű� C,���ű� M" & _
                "   Where A.��ǰ����ID=C.id(+) and  a.��ǰ����ID=M.id(+) " & _
                "           And A.����ID= [1] "
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����id, mlng�շ�ID)
        If rsTemp.EOF Then
            GoTo Init:
            Exit Function
        End If
        txtEdit(2).Text = zlStr.Nvl(rsTemp!����)
        cmdPati.Tag = zlStr.Nvl(rsTemp!Id)
        Dim i As Integer
        For i = 2 To txtEdit.UBound
            txtEdit(i).Text = zlStr.Nvl(rsTemp.Fields(txtEdit(i).Tag))
        Next
        If mlng�շ�ID <> 0 Then
            txtEdit(1).Text = zlStr.Nvl(rsTemp!����)
            If zlStr.Nvl(rsTemp!ʹ��ʱ��) <> "" Then
                MakTxtEdit.Text = zlStr.Nvl(rsTemp!ʹ��ʱ��)
            End If
        End If
    End If
Init:
    Init������Ϣ = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ISValid() As Boolean
    '------------------------------------------------------------------------------------------
    '����:��������Ĳ����Ƿ���Ч
    '����:
    '����ֵ:��Ч����True,����ΪFalse
    '------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim strTemp As String
        
    If Val(cmdPati.Tag) = 0 Then
        ShowMsgBox "δѡ����صĲ���,����!"
        If txtEdit(2).Enabled Then txtEdit(2).SetFocus
        Exit Function
    End If
    If MakTxtEdit.Text <> "____-__-__" Then
        If IsDate(MakTxtEdit.Text) = False Then
            ShowMsgBox "�������ڸ�ʽ,������!"
            If MakTxtEdit.Enabled Then MakTxtEdit.SetFocus
            Exit Function
        End If
    End If
    If zlCommFun.ActualLen(txtEdit(1).Text) > txtEdit(1).MaxLength Then
        ShowMsgBox "���벻�ܴ���" & txtEdit(1).MaxLength & " ���ַ���" & txtEdit(1).MaxLength / 2 & "������!"
        If txtEdit(1).Enabled Then txtEdit(1).SetFocus
        Exit Function
    End If
    
    ISValid = True
End Function

Private Sub cmdPati_Click()
        If MulitSelectPati("") = False Then
            If txtEdit(2).Enabled Then txtEdit(2).SetFocus
            Exit Sub
        End If
        OS.PressKey vbKeyTab
End Sub

Private Sub CmdSave_Click()
    '����:��֤��ص���Ϣ
   
    If ISValid() = False Then Exit Sub
        
    If MakTxtEdit.Text = "____-__-__" Then '
        mstrʹ��ʱ�� = ""
    Else
        mstrʹ��ʱ�� = MakTxtEdit.Text
    End If
    mlng����id = Val(cmdPati.Tag)
    mstr���� = Trim(txtEdit(2).Text)
    mstr���� = txtEdit(1).Text
    mblnOk = True
    Unload Me
End Sub
Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    If Init������Ϣ() = False Then Unload Me: Exit Sub
    Call SetTxtCtrColor
    mblnChange = False
End Sub
Private Sub Form_Load()
    mblnFirst = True
End Sub
 
Public Function ShowEdit(ByVal frmMain As Object, ByVal lng�շ�ID As Long, ByVal lng��ǰ����ID As Long, ByVal str�������� As String, ByVal lng����ID As Long, ByVal bln�༭ As Boolean, _
    ByRef lng����id As Long, ByRef str���� As String, ByRef strʹ��ʱ�� As String, ByRef str���� As String) As Boolean
    '-------------------------------------------------------------------------------------------------
    '����:��ʾ�༭����,�������
    '���:frmMain-������
    '     lng�շ�ID-�շ�ID(Ϊ��ʱ,��ʾ��������,���շ�ID,��Ϊ���ʾ,�޸Ĵ����շ���¼�еĲ��������Ϣ)
    '     lng��ǰ����ID-��ǰ����ID
    '     lng����id-����ID
    '     lng����ID -����ID
    '     str����-����
    '     strʹ��ʱ��
    '     bln�༭=�Ƿ���Ա༭
    '����:
    '     lng����ID -����ID
    '     str����-����
    '     strʹ��ʱ��
    '     str����
    '����:��ȷ������true,���򷵻�False
    '����:���˺�
    '����:2007/08/20
    '-------------------------------------------------------------------------------------------------
    '����:13415
    mstr�������� = str��������
    mblnFirst = True
    mlng�շ�ID = lng�շ�ID
    mlng����id = lng����id
    mlng����ID = lng����ID
    mlng��ǰ����ID = lng��ǰ����ID
    mstr���� = str����
    mstrʹ��ʱ�� = strʹ��ʱ��
    mbln�༭ = bln�༭
    Me.Show 1, frmMain
    lng����id = mlng����id
    str���� = mstr����
    str���� = mstr����
    strʹ��ʱ�� = mstrʹ��ʱ��
    ShowEdit = mblnOk
    
End Function

Private Sub CtlEnableSet()
    '---------------------------------------------------------------------------------------------------------------------
    '����:������ؿؼ���Enable
    '����:
    '����:���˺�
    '����:2007/08/20
    '---------------------------------------------------------------------------------------------------------------------
    cmdSave.Enabled = Val(cmdPati.Tag) <> 0
    
End Sub
Private Sub SetTxtCtrColor()
    '----------------------------------------------------------------------------------------------------------------------
    '����:���ò��ɱ༭���ı���ı���ɫ
    '����:���˺�
    '����:2007/08/20
    '----------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To txtEdit.UBound
        If txtEdit(i).Locked Or mbln�༭ = False Then
            txtEdit(i).Enabled = False
            txtEdit(i).BackColor = &H8000000F
        End If
    Next
    cmdSave.Visible = mbln�༭
    cmdPati.Enabled = mbln�༭
    MakTxtEdit.Enabled = mbln�༭
    If mbln�༭ = False Then
        MakTxtEdit.BackColor = &H8000000F
    End If
End Sub
 
Private Sub MakTxtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txtEDIT_Change(Index As Integer)
    If txtEdit(Index).Tag = "����" Then
        cmdPati.Tag = ""
    End If
    mblnChange = True
End Sub

Private Sub txtEDIT_GotFocus(Index As Integer)
     zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEDIT_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Select Case txtEdit(Index).Tag
    Case "����"
        If MulitSelectPati(txtEdit(Index).Text) = False Then Exit Sub
        If cmdSave.Enabled Then cmdSave.SetFocus
    Case Else
        OS.PressKey vbKeyTab
    End Select
End Sub

Private Sub txtEDIT_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = Asc("|") Then
        KeyAscii = 0
    End If
End Sub
Private Sub MakTxtEdit_Validate(Cancel As Boolean)
    Dim strFormat As Date
    
    If MakTxtEdit.Text = "____-__-__" Then Exit Sub
    
    err = 0:    On Error GoTo ErrHand:
    strFormat = CDate(MakTxtEdit.Text)
    Exit Sub
ErrHand:
    MsgBox "�������ڸ�ʽ���������룡", vbInformation, gstrSysName
    Cancel = True
End Sub

Private Sub vsfInfo_DblClick()
    With vsfInfo
        mlng����id = .TextMatrix(.Row, .ColIndex("id"))
        txtEdit(2).Text = .TextMatrix(.Row, .ColIndex("����"))
        txtEdit(3).Text = .TextMatrix(.Row, .ColIndex("�Ա�"))
        txtEdit(4).Text = .TextMatrix(.Row, .ColIndex("����"))
        txtEdit(5).Text = .TextMatrix(.Row, .ColIndex("��������"))
        txtEdit(6).Text = .TextMatrix(.Row, .ColIndex("�����ص�"))
        txtEdit(7).Text = .TextMatrix(.Row, .ColIndex("���֤��"))
        txtEdit(8).Text = .TextMatrix(.Row, .ColIndex("����"))
        txtEdit(9).Text = .TextMatrix(.Row, .ColIndex("ѧ��"))
        txtEdit(10).Text = .TextMatrix(.Row, .ColIndex("���"))
        txtEdit(11).Text = .TextMatrix(.Row, .ColIndex("����״��"))
        txtEdit(12).Text = .TextMatrix(.Row, .ColIndex("����"))
        txtEdit(13).Text = .TextMatrix(.Row, .ColIndex("��ǰ����"))
        txtEdit(14).Text = .TextMatrix(.Row, .ColIndex("��ǰ����"))
        txtEdit(15).Text = .TextMatrix(.Row, .ColIndex("��ǰ����"))
        txtEdit(16).Text = .TextMatrix(.Row, .ColIndex("�����"))
        txtEdit(17).Text = .TextMatrix(.Row, .ColIndex("סԺ��"))
        cmdPati.Tag = .TextMatrix(.Row, .ColIndex("id"))
        vsfInfo.Visible = False
    End With
End Sub

Private Sub vsfInfo_LostFocus()
    vsfInfo.Visible = False
End Sub


