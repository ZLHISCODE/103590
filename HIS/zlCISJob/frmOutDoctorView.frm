VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOutDoctorView 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "��������һ��"
   ClientHeight    =   8160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   15855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox PicCenter 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7695
      Left            =   -360
      ScaleHeight     =   7695
      ScaleWidth      =   15630
      TabIndex        =   1
      Top             =   0
      Width           =   15630
      Begin VB.PictureBox picBottom 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         FillColor       =   &H00C0E0FF&
         ForeColor       =   &H80000008&
         Height          =   6570
         Left            =   480
         ScaleHeight     =   6570
         ScaleWidth      =   15630
         TabIndex        =   4
         Top             =   1080
         Width           =   15630
         Begin VSFlex8Ctl.VSFlexGrid vsView 
            Height          =   6570
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   15630
            _cx             =   27570
            _cy             =   11589
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            MouseIcon       =   "frmOutDoctorView.frx":0000
            BackColor       =   -2147483643
            ForeColor       =   0
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483641
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   14737632
            GridColorFixed  =   10526880
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   0
            HighLight       =   0
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   10
            Cols            =   4
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmOutDoctorView.frx":0162
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   -1  'True
            MergeCells      =   1
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   1
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   -1  'True
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   1
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   -1  'True
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   2
            FrozenCols      =   1
            AllowUserFreezing=   0
            BackColorFrozen =   16777215
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin MSComctlLib.ImageList imgFlag 
            Left            =   13920
            Top             =   -120
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   8
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   12
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutDoctorView.frx":0260
                  Key             =   "��������"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutDoctorView.frx":077A
                  Key             =   "����"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutDoctorView.frx":0C94
                  Key             =   "�ϴο���"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutDoctorView.frx":219C
                  Key             =   "�ϴθ���"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutDoctorView.frx":36A4
                  Key             =   "�ϴβ�����"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutDoctorView.frx":4200
                  Key             =   "�´θ���"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutDoctorView.frx":5708
                  Key             =   "�´ο���"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutDoctorView.frx":6C10
                  Key             =   "�´β�����"
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutDoctorView.frx":776C
                  Key             =   "ֻ��ʾ����"
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutDoctorView.frx":A3E4
                  Key             =   "ֻ��ʾ���Ƹ���"
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutDoctorView.frx":D05C
                  Key             =   "��ʾ���о���"
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmOutDoctorView.frx":FCD4
                  Key             =   "��ʾ���о������"
               EndProperty
            EndProperty
         End
      End
      Begin VB.PictureBox picTop 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1005
         Left            =   360
         ScaleHeight     =   1634.146
         ScaleMode       =   0  'User
         ScaleWidth      =   15630
         TabIndex        =   2
         Top             =   0
         Width           =   15630
         Begin VB.PictureBox pitBtn 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   2
            Left            =   6360
            Picture         =   "frmOutDoctorView.frx":1294C
            ScaleHeight     =   480
            ScaleWidth      =   2400
            TabIndex        =   8
            Top             =   1560
            Width           =   2400
         End
         Begin VB.PictureBox pitBtn 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   1
            Left            =   3480
            Picture         =   "frmOutDoctorView.frx":1A690
            ScaleHeight     =   480
            ScaleWidth      =   2400
            TabIndex        =   7
            Top             =   1440
            Width           =   2400
         End
         Begin VB.PictureBox pitBtn 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Index           =   0
            Left            =   360
            Picture         =   "frmOutDoctorView.frx":223D4
            ScaleHeight     =   480
            ScaleWidth      =   2400
            TabIndex        =   6
            Top             =   1440
            Width           =   2400
         End
         Begin VB.Image imgBtn 
            Height          =   375
            Index           =   1
            Left            =   13800
            Picture         =   "frmOutDoctorView.frx":2A118
            Stretch         =   -1  'True
            Top             =   480
            Width           =   1050
         End
         Begin VB.Image imgBtn 
            Height          =   375
            Index           =   2
            Left            =   120
            Picture         =   "frmOutDoctorView.frx":2B610
            Stretch         =   -1  'True
            Top             =   480
            Width           =   2250
         End
         Begin VB.Image imgBtn 
            Height          =   375
            Index           =   0
            Left            =   12480
            Picture         =   "frmOutDoctorView.frx":2E278
            Stretch         =   -1  'True
            Top             =   480
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����һ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   6360
            TabIndex        =   3
            Top             =   240
            Width           =   1335
         End
      End
   End
   Begin VB.Label lblW 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   90
   End
End
Attribute VB_Name = "frmOutDoctorView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'��ڲ���
Private mlng����ID As Long
Private mlng����ID As Long

'�������
Private mrs����ʱ�� As ADODB.Recordset        '���没�����ξ���ʱ��:���ֶ�:��š�����ID���Һŵ���ִ��ʱ�䣩
Private mrs��� As ADODB.Recordset            '�����Ϣ
Private mrsDrug As ADODB.Recordset
Private mrs������ As ADODB.Recordset
Private mrs�������� As ADODB.Recordset
Private mrsEMR As ADODB.Recordset

Private mrs����ҽ�� As ADODB.Recordset

Private mlngPrev As Long   'ǰһ��   ȱʡ 4
Private mlngNext As Long   '��һ��   ȱʡ 1
Private mbytFontSize As Long   '�����С

Private mcolCate As Collection

Private mstr����ID As String
Private mstr�Һŵ� As String
Private mstrFontUnderLine As String   '����»�����  �к�|��1|��2
Private mintIndex  As Integer        '��ǵ�ǰ����ͼ��
Private mbytShow As Byte                '0-��ʾ���о����¼;1-����ʾ�����Ҿ����¼ (ȱʡ��ʾ���о����¼)

Private Const mstr���� As String = "����ʱ��|���|��ҩ����|��������|���|����|����ҽ��"
Private Const mlngSubCol  As Long = 5 '��ע(��\��\��)|ͼ��(������)|����(ҩƷ����)|�Ƴ�|�����

'��ɫ���:ǳ��,����,�ʺ�,���,����
Private Enum CONST_COLOR
    '��ɫ
    COLOR_ǳ�� = &HC0C0FF          'ǳ��
    COLOR_���� = &H8080FF
    COLOR_�ʺ� = &HFF&
    COLOR_��� = &HC0&
    COLOR_���� = &H80&
    '��ɫ
    COLOR_ǳ�� = &HC0E0FF
    COLOR_���� = &H80C0FF
    COLOR_�ʳ� = &H80FF&
    COLOR_��� = &H40C0&
    COLOR_���� = &H4080&
    '��ɫ
    COLOR_ǳ�� = &HC0FFFF
    COLOR_���� = &H80FFFF
    COLOR_�ʻ� = &HFFFF&
    COLOR_��� = &HC0C0&
    COLOR_���� = &H8080&
    '��ɫ
    COLOR_ǳ�� = &HC0FFC0
    COLOR_���� = &H80FF80
    COLOR_���� = &HFF00&
    COLOR_���� = &HC000&
    COLOR_���� = &H8000&
    '��ɫ
    COLOR_ǳ�� = &HFFFFC0
    COLOR_���� = &HFFFF80
    COLOR_���� = &HFFFF00
    COLOR_���� = &HC0C000
    COLOR_���� = &H808000
    '��ɫ
    COLOR_ǳ�� = &HFFC0C0
    COLOR_��ɫ = &HFF8080
    COLOR_���� = &HFF0000
    COLOR_���� = &HC00000
    COLOR_���� = &H800000
    '��ɫ
    COLOR_ǳ�� = &HFFC0FF
    COLOR_���� = &HFF80FF
    COLOR_���� = &HFF00FF
    COLOR_���� = &HC000C0
    COLOR_���� = &H800080
    '��ɫ
    COLOR_��ɫ = &H80000005
    COLOR_FORMBK = &H8000000B
    COLOR_CENTERBK = &H808000
End Enum


Private Enum CONST_CATEGORY
    CATE_����ʱ�� = 0
    CATE_��� = 1
    CATE_��ҩ���� = 2
    CATE_�������� = 3
    CATE_��� = 4
    CATE_���� = 5
    CATE_����ҽ�� = 6
End Enum

Private Enum CONST_IX_CMD
    CMD_PREV = 0
    CMD_NEXT = 1
    CMD_ORTHER = 2
End Enum

Private Enum CONST_SUBCOL
    SUBCOL_��ע = 0
    SUBCOL_ͼ�� = 1
    SUBCOL_���� = 2
    SUBCOL_�Ƴ� = 3
    SUBCOL_��� = 4
End Enum

Public Function zlRefresh(frmParent As Object, ByVal lng����ID As Long, ByVal lng����ID As Long) As Boolean
'���ܣ�
    If lng����ID = 0 Then
        Exit Function
    End If
    mlng����ID = lng����ID
    mlng����ID = lng����ID
    mlngPrev = 3     'Ĭ��ֵ
    mlngNext = 1     'Ĭ��ֵ
    Call SubRefresh
End Function

Private Sub LoadView()
'����:������ͼ
    Dim strTmp As String
    Dim i As Long, k As Long, j As Long
    Dim lng���ID As Long, lngCount As Long
    Dim lngColor As Long, lngContW As Long
    Dim lng�Һ�id As Long
    Dim lngDay As Long
    Dim strFilter As String, strContent As String
    Dim strDrug As String
    Dim strType As String
    Dim strDay As String
    Dim strMerge As String     '��¼һ����ҩ��Ҫ�ϲ����к�
    Dim lngRow As Long, lngMergeRow As Long
    Dim lngDrugCol As Long
    Dim rsDrug As ADODB.Recordset
    Dim blnAddRow As Boolean
    Dim blnMerge As Boolean
    With vsView
        If mrs����ʱ�� Is Nothing Then Exit Sub
        
        .Redraw = flexRDNone
        mrs����ʱ��.Filter = "��� >=" & mlngNext & " And ��� <= " & mlngPrev
        mrs����ʱ��.MoveLast
        For i = 1 To mrs����ʱ��.RecordCount
            '����ʱ��
            If NVL(mrs����ʱ��!���, 0) = 1 Then
                strTmp = "���ξ���"
            ElseIf NVL(mrs����ʱ��!���, 0) = 2 Then
                strTmp = "�ϴξ���"
            Else
                strTmp = ""
            End If
            strTmp = strTmp & " " & Format(mrs����ʱ��!ִ��ʱ�� & "", "YYYY-MM-DD hh:mm") & IIf(Val(mrs����ʱ��!ִ�в���ID & "") <> mlng����ID, "[" & mrs����ʱ��!�������� & "]", "")
            .Cell(flexcpText, 0, (i - 1) * mlngSubCol + 1, 0, i * mlngSubCol) = strTmp
            .ColData((i - 1) * mlngSubCol + SUBCOL_���� + 1) = CLng(mrs����ʱ��!����id) ' ��¼ÿ�еľ���ID  'ҽ��������
            .Cell(flexcpData, mcolCate("_" & CATE_����ʱ��).lngBeginRow, (i - 1) * mlngSubCol + SUBCOL_���� + 1) = CStr(mrs����ʱ��!�Һŵ�) '��¼�¹Һŵ�
            
            '�����ʽ����
            .Cell(flexcpAlignment, 0, 1, 0, .Cols - 1) = flexAlignCenterCenter '����ʱ��
            .Cell(flexcpAlignment, 1, 1, 1, .Cols - 1) = flexAlignLeftCenter   '���
   
            '�������
            mrs���.Filter = "����ID =" & mrs����ʱ��!����id
            strTmp = ""
            For j = 1 To mrs���.RecordCount
                strTmp = strTmp & "," & mrs���!������� & ""
                mrs���.MoveNext
            Next
            lngContW = .ColWidth(SUBCOL_��ע) + .ColWidth(SUBCOL_����) + .ColWidth(SUBCOL_�Ƴ�) - 300
            If strTmp <> "" Then
                strTmp = Mid(strTmp, 2)
                .Cell(flexcpData, mcolCate("_" & CATE_���).lngBeginRow, (i - 1) * mlngSubCol + SUBCOL_���� + 1) = strTmp
                strTmp = GetSubString(strTmp, lngContW)
                .Cell(flexcpText, mcolCate("_" & CATE_���).lngBeginRow, (i - 1) * mlngSubCol + 1, mcolCate("_" & CATE_���).lngBeginRow, i * mlngSubCol) = strTmp
            Else
                .Cell(flexcpText, mcolCate("_" & CATE_���).lngBeginRow, (i - 1) * mlngSubCol + 1, mcolCate("_" & CATE_���).lngBeginRow, i * mlngSubCol) = IIf(i Mod 2 = 0, " ", "  ") '��������еĺϲ�����
            End If
            
            '��ҩ����
            lngDrugCol = (i - 1) * mlngSubCol + SUBCOL_���� + 1
            strFilter = "�Һŵ� ='" & mrs����ʱ��!�Һŵ� & "' And ������� <> 'E'"
            mrsDrug.Filter = strFilter
            Set rsDrug = zlDatabase.CopyNewRec(mrsDrug)
            lngRow = mcolCate("_" & CATE_��ҩ����).lngBeginRow
            lngCount = 0
            For k = 1 To rsDrug.RecordCount
                blnAddRow = True: blnMerge = False
                '�������=5,6
                'ҽ������:ҩƷ����(��,��,��)vbTabҩƷ���� vbTAB �Ƴ�(��λΪ��,����7��Ĭ��Ϊ7��)
                If rsDrug!������� & "" = "5" Or rsDrug!������� & "" = "6" Then
                    .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_ͼ�� + 1) = CLng(rsDrug!���ID & "")      '���ڱ�ʶͬһ��ҩƷ
                    '���һ����ҩ������
                    mrsDrug.Filter = "���ID=" & CLng(rsDrug!���ID & "")
                    
                    If mrsDrug.RecordCount > 1 Then
                        If lngMergeRow = 0 Then lngMergeRow = 1
                        blnMerge = True:
                        If lngMergeRow = 1 Then
                            .TextMatrix(lngRow, (i - 1) * mlngSubCol + SUBCOL_ͼ�� + 1) = "��"
                            lngMergeRow = lngMergeRow + 1
                            lngCount = lngCount + 1
                        ElseIf lngMergeRow = mrsDrug.RecordCount Then
                            .TextMatrix(lngRow, (i - 1) * mlngSubCol + SUBCOL_ͼ�� + 1) = "��"
                            lngMergeRow = 0
                        ElseIf lngMergeRow <> 1 And lngMergeRow <> mrsDrug.RecordCount Then
                            .TextMatrix(lngRow, (i - 1) * mlngSubCol + SUBCOL_ͼ�� + 1) = "��"
                            lngMergeRow = lngMergeRow + 1
                        End If
                    Else
                        lngMergeRow = 0
                        lngCount = lngCount + 1
                    End If
                        
                    If lng���ID <> CLng(rsDrug!���ID & "") Then '����һ����ҩ�ظ�ȡ
                        lng���ID = CLng(rsDrug!���ID & "")

                        mrsDrug.Filter = "ID=" & lng���ID     '��ҩ;��
                        '0-�����������,1-��Һ��,2-ע����,3-Ƥ��,4-�ڷ�
                        If NVL(mrsDrug!ִ�з���) = "1" Then
                            strType = "��"
                            lngColor = COLOR_ǳ��
                        ElseIf NVL(mrsDrug!ִ�з���) = "2" Then
                            strType = "��"
                            lngColor = COLOR_ǳ��
                        ElseIf NVL(mrsDrug!ִ�з���) = "3" Then
                            strType = "Ƥ"
                            lngColor = COLOR_ǳ��
                        ElseIf NVL(mrsDrug!ִ�з���) = "4" Then
                            strType = "��"
                            lngColor = COLOR_ǳ��
                        Else
                            strType = " "
                            lngColor = COLOR_ǳ��
                        End If
       
                        '�Ƴ�
                        If IsNull(rsDrug!����) Then
                            'ͨ������,������������
                            If NVL(rsDrug!��������, 0) <> 0 Then
                                lngDay = CalcȱʡҩƷ����(NVL(rsDrug!�ܸ�����, 0), NVL(rsDrug!��������, 0), _
                                        NVL(rsDrug!Ƶ�ʴ���, 0), NVL(rsDrug!Ƶ�ʼ��, 0), NVL(rsDrug!�����λ, 0), _
                                        NVL(rsDrug!����ϵ��, 0), NVL(rsDrug!�����װ, 0), _
                                        NVL(rsDrug!�ɷ����, 0))
                            Else
                                lngDay = 7  'δ���õ���ʱȱʡ��Ϊ7��
                            End If
                            
                        Else
                            'ֱ��ȡ����
                            lngDay = NVL(rsDrug!����, 0)
                        End If
                        strDay = "(" & lngDay & "��)"
                    End If
                    '�ϲ�ͬ��ҩƷ�µı�ע
                    .Cell(flexcpText, lngRow, (i - 1) * mlngSubCol + SUBCOL_��ע + 1) = strType & IIf(lngCount Mod 2 = 0, "", vbTab)
                    .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_��ע + 1) = lngColor
                    'һ����ҩ��ע��Ҫ�ϲ�
                    .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_���� + 1) = rsDrug!ҽ������ & ""
                    If Not blnMerge Then
                        .Cell(flexcpText, lngRow, (i - 1) * mlngSubCol + SUBCOL_�Ƴ� + 1) = strDay  '�ϲ���Draw_Cell����
                    End If
                    .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_�Ƴ� + 1) = lngDay    '���ڱ���ɫ������ʾ����
              
                    
                ElseIf rsDrug!������� & "" = "7" Then
                    '��ҩ�䷽
                    If NVL(rsDrug!���ID, 0) <> lng���ID Then
                        lng���ID = CLng(rsDrug!���ID & "")
                        lngCount = lngCount + 1
                        mrsDrug.Filter = "ID=" & lng���ID
                        .Cell(flexcpData, lngRow, lngDrugCol) = lng���ID    '���ڱ�ʶͬһ��ҩƷ
                        .Cell(flexcpText, lngRow, (i - 1) * mlngSubCol + SUBCOL_��ע + 1) = "��" & IIf(lngCount Mod 2 = 0, "", vbTab)
                        .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_��ע + 1) = COLOR_ǳ��
                        .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_���� + 1) = mrsDrug!ҽ������ & ""
                        
                        .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_�Ƴ� + 1) = 7  'ȱʡ����
                    Else
                        blnAddRow = False
                    End If
                    
                End If
                If blnAddRow Then
                    lngRow = lngRow + 1
                End If
                rsDrug.MoveNext
            Next
            '���
            strFilter = "�Һŵ� ='" & mrs����ʱ��!�Һŵ� & "' and ������� ='D' "
            mrs������.Filter = strFilter
            lngRow = mcolCate("_" & CATE_���).lngBeginRow
            
            lngContW = .ColWidth(SUBCOL_���� + 1)
            
            For k = 1 To mrs������.RecordCount
                .MergeCol((i - 1) * mlngSubCol + SUBCOL_�Ƴ� + 1) = False
                '�ı����ݼ�¼��������������ʾ
                strContent = mrs������!ҽ������ & ""
                .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_ͼ�� + 1) = strContent
                strContent = GetSubString(strContent, lngContW)
                .Cell(flexcpText, lngRow, (i - 1) * mlngSubCol + SUBCOL_ͼ�� + 1, lngRow, (i - 1) * mlngSubCol + SUBCOL_���� + 1) = strContent
                .Cell(flexcpAlignment, lngRow, (i - 1) * mlngSubCol + SUBCOL_ͼ�� + 1, lngRow, (i - 1) * mlngSubCol + SUBCOL_���� + 1) = flexAlignLeftCenter
                .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_���� + 1) = mrs������!ID & ""  '��¼�¼��ҽ����ID
                
                '�ѳ����� ����ɫ������ʾ���ӱ���ͼ��
                If Not (Val(mrs������!����ID & "") = 0 And Val(mrs������!��鱨��ID & "") = 0) Then
                    .Cell(flexcpForeColor, lngRow, (i - 1) * mlngSubCol + SUBCOL_ͼ�� + 1, lngRow, (i - 1) * mlngSubCol + SUBCOL_���� + 1) = COLOR_����
                    Set .Cell(flexcpPicture, lngRow, (i - 1) * mlngSubCol + SUBCOL_��ע + 1) = imgFlag.ListImages("����").Picture
                    .Cell(flexcpPictureAlignment, lngRow, (i - 1) * mlngSubCol + SUBCOL_��ע + 1) = flexAlignCenterCenter
                    .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_��ע + 1) = mrs������!����ID & "|" & mrs������!��鱨��ID
                End If
                .TextMatrix(lngRow, (i - 1) * mlngSubCol + SUBCOL_�Ƴ� + 1) = "��Ƭ" & IIf(lngRow Mod 2 = 0, "", vbTab)
                .Cell(flexcpForeColor, lngRow, (i - 1) * mlngSubCol + SUBCOL_�Ƴ� + 1) = COLOR_����
            
                .MergeRow(lngRow) = True
                lngRow = lngRow + 1
                mrs������.MoveNext
            Next
            '����
            strFilter = "�Һŵ� ='" & mrs����ʱ��!�Һŵ� & "' and ������� ='E' And ��������='6' "
            mrs������.Filter = strFilter
            lngRow = mcolCate("_" & CATE_����).lngBeginRow
            lngContW = .ColWidth(SUBCOL_���� + 1)
                
            For k = 1 To mrs������.RecordCount
                strContent = mrs������!ҽ������ & ""
                .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_ͼ�� + 1) = strContent
                strContent = GetSubString(strContent, lngContW)
                .Cell(flexcpText, lngRow, (i - 1) * mlngSubCol + SUBCOL_ͼ�� + 1, lngRow, (i - 1) * mlngSubCol + SUBCOL_���� + 1) = strContent
                .TextMatrix(lngRow, (i - 1) * mlngSubCol + SUBCOL_�Ƴ� + 1) = IIf(lngRow Mod 2 = 0, "", vbTab)
                .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_���� + 1) = mrs������!ID & ""  '��¼ID
                .Cell(flexcpAlignment, lngRow, (i - 1) * mlngSubCol + SUBCOL_ͼ�� + 1, lngRow, (i - 1) * mlngSubCol + SUBCOL_���� + 1) = flexAlignLeftCenter
                
                 '�ѳ����� ����ɫ������ʾ���ӱ���ͼ��
                If Val(mrs������!����ID & "") <> 0 Then
                    .Cell(flexcpForeColor, lngRow, (i - 1) * mlngSubCol + SUBCOL_ͼ�� + 1, lngRow, (i - 1) * mlngSubCol + SUBCOL_���� + 1) = COLOR_����
                    Set .Cell(flexcpPicture, lngRow, (i - 1) * mlngSubCol + SUBCOL_��ע + 1) = imgFlag.ListImages("����").Picture
                    .Cell(flexcpPictureAlignment, lngRow, (i - 1) * mlngSubCol + SUBCOL_��ע + 1) = flexAlignCenterCenter
                    .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_��ע + 1) = mrs������!����ID & ""
                End If
                
                .MergeRow(lngRow) = True
                lngRow = lngRow + 1
                mrs������.MoveNext
            Next
            '����
            
            strFilter = "�Һŵ� ='" & mrs����ʱ��!�Һŵ� & "'"
            mrs����ҽ��.Filter = strFilter
            lngRow = mcolCate("_" & CATE_����ҽ��).lngBeginRow
            lngContW = .ColWidth(SUBCOL_���� + 1)
            
            For k = 1 To mrs����ҽ��.RecordCount
                strContent = mrs����ҽ��!ҽ������ & ""
                .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_ͼ�� + 1) = strContent
                strContent = GetSubString(strContent, lngContW)
                .Cell(flexcpText, lngRow, (i - 1) * mlngSubCol + SUBCOL_ͼ�� + 1, lngRow, (i - 1) * mlngSubCol + SUBCOL_���� + 1) = strContent
                .Cell(flexcpAlignment, lngRow, (i - 1) * mlngSubCol + SUBCOL_ͼ�� + 1, lngRow, (i - 1) * mlngSubCol + SUBCOL_�Ƴ� + 1) = flexAlignLeftCenter
                .MergeRow(lngRow) = True
                lngRow = lngRow + 1
                mrs����ҽ��.MoveNext
            Next
            
        
            '����
            strFilter = "�Һ�ID=" & mrs����ʱ��!����id
            mrs��������.Filter = strFilter
            lngRow = mcolCate("_" & CATE_��������).lngBeginRow
            lngContW = .ColWidth(SUBCOL_���� + 1)
            '�ɰ没��
            For k = 1 To mrs��������.RecordCount
                strContent = mrs��������!�������� & ""
                .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_ͼ�� + 1) = strContent
                strContent = GetSubString(strContent, lngContW)
                .Cell(flexcpText, lngRow, (i - 1) * mlngSubCol + SUBCOL_ͼ�� + 1, lngRow, (i - 1) * mlngSubCol + SUBCOL_���� + 1) = strContent
                .Cell(flexcpAlignment, lngRow, (i - 1) * mlngSubCol + SUBCOL_ͼ�� + 1, lngRow, (i - 1) * mlngSubCol + SUBCOL_���� + 1) = flexAlignLeftCenter
                .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_���� + 1) = mrs��������!ID & ""    '��¼�²�����¼ID
                
                .MergeRow(lngRow) = True
                lngRow = lngRow + 1
                mrs��������.MoveNext
            Next
            '�°没��
            mrsEMR.Filter = "�Һ�ID=" & mrs����ʱ��!����id
            
            For k = 1 To mrsEMR.RecordCount
                strContent = mrsEMR!�������� & ""
                .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_ͼ�� + 1) = strContent
                strContent = GetSubString(strContent, lngContW)
                .Cell(flexcpText, lngRow, (i - 1) * mlngSubCol + SUBCOL_ͼ�� + 1, lngRow, (i - 1) * mlngSubCol + SUBCOL_���� + 1) = strContent
                .Cell(flexcpAlignment, lngRow, (i - 1) * mlngSubCol + SUBCOL_ͼ�� + 1, lngRow, (i - 1) * mlngSubCol + SUBCOL_���� + 1) = flexAlignLeftCenter
                .Cell(flexcpData, lngRow, (i - 1) * mlngSubCol + SUBCOL_���� + 1) = mrsEMR!ID & ""    '��¼�²�����¼ID
                .MergeRow(lngRow) = True
                lngRow = lngRow + 1
                mrsEMR.MoveNext
            Next
            mrs����ʱ��.MovePrevious
        Next
        .Redraw = flexRDDirect
        .Row = 0 '��ȡ����
    End With

    
End Sub

Private Sub ReadRegister()
'����:��ȡ�Ǽ���Ϣ
    Dim strSql As String
    Dim i As Long, j As Long
    Dim rsEmr As ADODB.Recordset
    Dim strMsg As String
    
    On Error GoTo errH
    '1-����ʱ��
    strSql = "Select Rownum As ���, b.����ID,b.�Һŵ�,b.ִ��ʱ��,b.ִ��״̬,b.ִ�в���ID,b.��������  " & vbNewLine & _
            "From (Select a.Id As ����id,a.NO as �Һŵ�,a.ִ��ʱ��,a.ִ��״̬,a.ִ�в���ID,d.���� as �������� " & vbNewLine & _
            "       From ���˹Һż�¼ A,���ű� D " & vbNewLine & _
            "       Where a.ִ�в���ID =d.Id(+) and a.����id = [1] And a.��¼���� = 1 And a.��¼״̬ = 1 " & IIf(mbytShow = 0, "", " And a.ִ�в���ID =[2]") & vbNewLine & _
            "       Order By a.ִ��ʱ�� Desc) B"
    
    Set mrs����ʱ�� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mlng����ID)
    
    '2-��ϼ�¼
    mrs����ʱ��.Filter = "��� >=" & mlngNext & " And ��� <= " & mlngPrev
    
    mstr�Һŵ� = "": mstr����ID = ""
    For i = 1 To mrs����ʱ��.RecordCount
        mstr����ID = mstr����ID & "," & mrs����ʱ��!����id
        mstr�Һŵ� = mstr�Һŵ� & "," & mrs����ʱ��!�Һŵ�
        mrs����ʱ��.MoveNext
    Next
    mstr����ID = mstr����ID & ","
    mstr�Һŵ� = mstr�Һŵ� & ","
    
    strSql = "Select a.��ҳid As ����id, a.�������" & vbNewLine & _
            "From ������ϼ�¼ A" & vbNewLine & _
            "Where ����id = [1] And Instr([2], ',' || ��ҳid || ',') > 0 " & vbNewLine & _
            " order by a.��ҳid,a.��ϴ���"
            
    Set mrs��� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mstr����ID)
    
    '3-ҽ����¼
    strSql = "Select a.�Һŵ�, a.Id, a.���id, a.���, a.ҽ����Ч, a.ҽ������, a.�걾��λ, a.�������, a.������Ŀid, a.����, a.��������, a.�ܸ�����, a.ִ��Ƶ��, b.��������, b.ִ�з���," & vbNewLine & _
            "       a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.�����λ, b.���㵥λ As ������λ, c.����ϵ��, c.�����װ, c.����ɷ���� As �ɷ���� " & vbNewLine & _
            "From ����ҽ����¼ A, ������ĿĿ¼ B, ҩƷ��� C" & vbNewLine & _
            "Where a.������Ŀid = b.Id And a.�շ�ϸĿid = c.ҩƷid(+) And a.����id = [1] And a.ҽ��״̬ = 8 And" & vbNewLine & _
            "      (a.������� In ('5', '6', '7') Or (a.������� = 'E' And b.�������� In ('1', '2', '3', '4'))) And" & vbNewLine & _
            "      Instr([2], ',' || a.�Һŵ� || ',') > 0" & vbNewLine & _
            "Order By a.�Һŵ�, a.���"

    Set mrsDrug = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mstr�Һŵ�)
    '������
    strSql = "Select a.�Һŵ�,a.ID,a.ҽ������,a.������� ,b.��������, Max(c.����id) As ����id, Max(c.��鱨��id) As ��鱨��id," & vbNewLine & _
        "       Decode(Max(Nvl(c.����״̬, 0)), Min(Nvl(c.����״̬, 0)), Max(Nvl(c.����״̬, 0)), 2) As ����״̬" & vbNewLine & _
        "From ����ҽ����¼ A, ������ĿĿ¼ B, ����ҽ������ C" & vbNewLine & _
        "Where a.������Ŀid = b.Id And a.Id = c.ҽ��id(+) And a.����id = [1] And a.ҽ��״̬ = 8 And" & vbNewLine & _
        "      (a.������� = 'D' And a.���id Is Null Or a.������� = 'E' And b.�������� = '6') And" & vbNewLine & _
        "      Instr([2], ',' || a.�Һŵ� || ',') > 0" & vbNewLine & _
        "Group By a.�Һŵ�,a.ID,a.ҽ������,a.���,a.�������, b.��������" & vbNewLine & _
        "Order By a.�Һŵ�, a.���"

    Set mrs������ = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mstr�Һŵ�)
    '����ҽ��
    '1-�ų� ������E��2-��ҩ����(��ҩ);3-��ҩ�巨;4-��ҩ��(��)��;6-�ɼ�����(����);8-��Ѫ;�� ;ֻ����;0-��ͨ,1-��������,5-��������
    '2-�ų� ����,����ͬ����һ��ʹ��
    '3-�ų� ����,���
    '4-�ų� 5-��ҩ,6-�г�ҩ,7-��ҩ
    '5-����\��Ѫ ֻ������ҽ���� ���ID IS NULL
    '6-������ Z
    strSql = "Select a.�Һŵ�, a.ҽ������, a.�������" & vbNewLine & _
        "From ����ҽ����¼ A, ������ĿĿ¼ B" & vbNewLine & _
        "Where a.������Ŀid = b.Id And a.����id = [1] And a.ҽ��״̬ = 8 And Not a.������� In ('G', 'D', 'C', '5', '6', '7') And" & vbNewLine & _
        "      Not (NVL(b.��������,0) In ('2', '3', '4', '6', '8') And a.������� = 'E') And NVL(���id,0)=0 And" & vbNewLine & _
        "      Instr([2], ',' || a.�Һŵ� || ',') > 0" & vbNewLine & _
        "Order By a.�Һŵ�, a.���"
    Set mrs����ҽ�� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mstr�Һŵ�)
    
    '���ﲡ��
    strSql = "Select ID , Nvl(��ҳid, 0) as �Һ�ID,��������, �������� " & vbNewLine & _
            "From ���Ӳ�����¼" & vbNewLine & _
            "Where ������Դ = 1 And (�������� In (1, 6) Or (�������� = 5 And �༭��ʽ <> 2)) And ����id = [1] And Instr([2], ',' || Nvl(��ҳid, 0) || ',') > 0 " & vbNewLine & _
            "Order By  Nvl(��ҳid, 0),��������, ���, ����ʱ��"
    Set mrs�������� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mstr����ID)
    '�°没��
    Set mrsEMR = InitRS
    If Not gobjEmr Is Nothing Then
        mrs����ʱ��.MoveFirst
        For i = 1 To mrs����ʱ��.RecordCount
            '�°没���ṩ�ӿڣ�GetOutEPRRecord(�Һ�ID)����ÿ�ξ���Ĳ��������ID,Title����
            On Error Resume Next
            strMsg = gobjEmr.GetOutEPRRecord(mrs����ʱ��!����id & "", rsEmr)
            err.Clear: On Error GoTo 0
            If Not rsEmr Is Nothing Then
                For j = 1 To rsEmr.RecordCount
                    mrsEMR.AddNew
                    mrsEMR!�Һ�ID = CLng(mrs����ʱ��!����id & "")
                    mrsEMR!ID = rsEmr!ID
                    mrsEMR!�������� = rsEmr!Title
                    mrsEMR.Update
                    rsEmr.MoveNext
                Next
            End If
            mrs����ʱ��.MoveNext
        Next
    End If
    If mrsEMR.RecordCount > 0 Then mrsEMR.MoveFirst
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SetViewRows() As Integer
'����:������ͼ��ʾ������
    Dim udtCate As TYPE_CATE
    Dim arrTmp As Variant
    Dim lngTmp As Long
    Dim lngRows As Long
    Dim i As Long
    
    Set mcolCate = New Collection
    arrTmp = Split(mstr����, "|")
    lngRows = 0
    For i = CATE_����ʱ�� To CATE_����ҽ��
        '����֮��������һ����Ϊ�����
        If i >= CATE_��� And i <= CATE_����ҽ�� Then
            lngRows = lngRows + 1
        End If
        
        lngTmp = CalculateMaxRows(i)
        
        udtCate.strName = arrTmp(i)
        udtCate.lngBeginRow = lngRows
        udtCate.lngEndRow = lngRows + lngTmp - 1
        
        mcolCate.Add udtCate, "_" & i
        lngRows = lngRows + lngTmp
    Next

    SetViewRows = lngRows
End Function

Private Function CalculateMaxRows(ByVal bytFun As Byte) As Integer
'---------------------------------------------------------------------------------------------
'����:�����������
'����:���ظ�����Ӧ����ʾ����
'---------------------------------------------------------------------------------------------
    Dim intNum As Integer
    Dim intMaxNum As Integer
    Dim i As Long, j As Long
    
    Dim blnDo As Boolean
    
    Dim str�Һŵ� As String
    Dim str���ID As String
    Dim strErr As String
    
    
    Dim rsEmr As ADODB.Recordset
    Dim arrTmp As Variant
    
    blnDo = Not mrs����ʱ�� Is Nothing
    
    Select Case bytFun
    Case CATE_����ʱ��, CATE_���
        intMaxNum = 1
    Case CATE_��ҩ����
        intMaxNum = 10
        If blnDo Then
            mrsDrug.Filter = "������� ='5' or ������� ='6' or ������� ='7'"
            For j = 1 To mrsDrug.RecordCount
                If str�Һŵ� <> mrsDrug!�Һŵ� & "" Then
                    str�Һŵ� = mrsDrug!�Һŵ�
                    If intMaxNum < intNum Then intMaxNum = intNum '��¼�����
                    intNum = 0
                End If
                '-��ҩ�䷽,���ҩƷֻ��ʾһ��
                If NVL(mrsDrug!�������, "") = "7" And str���ID <> mrsDrug!���ID & "" Then  '��ҩ�䷽
                    str���ID = mrsDrug!���ID & ""       '��������ж�
                    intNum = intNum + 1
                ElseIf InStr(",5,6,", mrsDrug!������� & "") > 1 Then  '��ҩ����ҩ
                '-һ����ҩ,�м���ҩƷ��ռ�ü���
                    intNum = intNum + 1
                End If
                '���һ��ʱ,��һ�μ�¼�����
                If j = mrsDrug.RecordCount Then
                    If intMaxNum <= intNum Then intMaxNum = intNum + 1 '��¼�����
                End If
                mrsDrug.MoveNext
            Next
        End If
     Case CATE_���
        intMaxNum = 5
        If blnDo Then
            mrs����ʱ��.MoveFirst
            For j = 1 To mrs����ʱ��.RecordCount
                mrs������.Filter = "�Һŵ� ='" & mrs����ʱ��!�Һŵ� & "' And ������� ='D'"
                If intMaxNum <= mrs������.RecordCount Then intMaxNum = mrs������.RecordCount + 1
                
                mrs����ʱ��.MoveNext
            Next
        End If
     Case CATE_��������
        intMaxNum = 5
        '����
        If blnDo Then
            mrs����ʱ��.MoveFirst
            For j = 1 To mrs����ʱ��.RecordCount
                mrs��������.Filter = "�Һ�ID =" & mrs����ʱ��!����id
                mrsEMR.Filter = "�Һ�ID =" & mrs����ʱ��!����id
                If intMaxNum <= (mrs��������.RecordCount + mrsEMR.RecordCount) Then intMaxNum = (mrs��������.RecordCount + mrsEMR.RecordCount) + 1
                mrs����ʱ��.MoveNext
            Next
        End If
    Case CATE_����
        intMaxNum = 5
        If blnDo Then
            mrs����ʱ��.MoveFirst
            For j = 1 To mrs����ʱ��.RecordCount
                mrs������.Filter = "�Һŵ� ='" & mrs����ʱ��!�Һŵ� & "' And ������� ='E' And �������� = '6'"   '������ʾ�ɼ�������
                If intMaxNum <= mrs������.RecordCount Then intMaxNum = mrs������.RecordCount + 1
                
                mrs����ʱ��.MoveNext
            Next
        End If
    Case CATE_����ҽ��
        intMaxNum = 2
        If blnDo Then
            mrs����ʱ��.MoveFirst
            For j = 1 To mrs����ʱ��.RecordCount
                mrs����ҽ��.Filter = "�Һŵ� ='" & mrs����ʱ��!�Һŵ� & "'"
                If intMaxNum <= mrs����ҽ��.RecordCount Then intMaxNum = mrs����ҽ��.RecordCount + 1
                mrs����ʱ��.MoveNext
            Next
        End If
    End Select
    CalculateMaxRows = intMaxNum
End Function

Private Sub SubRefresh(Optional ByVal Index As Integer = -1)
'����:ˢ��

    If Index <> -1 Then
        If Index = CMD_PREV Then
            If imgBtn(CMD_PREV).Enabled = False Then Exit Sub
            mlngPrev = mlngPrev + 1
            mlngNext = mlngNext + 1
        ElseIf Index = CMD_NEXT Then
            If imgBtn(CMD_NEXT).Enabled = False Then Exit Sub
            mlngPrev = mlngPrev - 1
            mlngNext = mlngNext - 1
        ElseIf Index = CMD_ORTHER Then
            mlngPrev = 3
            mlngNext = 1
            If mbytShow = 1 Then
                Set imgBtn(CMD_ORTHER).Picture = imgFlag.ListImages("��ʾ���о���").Picture
                 mbytShow = 0
            Else
                Set imgBtn(CMD_ORTHER).Picture = imgFlag.ListImages("ֻ��ʾ����").Picture
                mbytShow = 1
            End If
        End If
    Else
        If CheckRegister Then
            imgBtn(CMD_ORTHER).Visible = True
        Else
            imgBtn(CMD_ORTHER).Visible = False
        End If
    End If
    
    mstrFontUnderLine = ""
    
    Call ReadRegister
    Call InitVsView
    Call ResizeVsView
    Call LoadView
    mrs����ʱ��.Filter = ""
    imgBtn(CMD_PREV).Enabled = mlngPrev < mrs����ʱ��.RecordCount
    imgBtn(CMD_NEXT).Enabled = mlngNext > 1
 
    Set imgBtn(CMD_PREV).Picture = IIf(imgBtn(CMD_PREV).Enabled, imgFlag.ListImages("�ϴο���").Picture, imgFlag.ListImages("�ϴβ�����").Picture)
    Set imgBtn(CMD_NEXT).Picture = IIf(imgBtn(CMD_NEXT).Enabled, imgFlag.ListImages("�´ο���").Picture, imgFlag.ListImages("�´β�����").Picture)
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyRight Then
        Call SubRefresh(CMD_NEXT)
    ElseIf KeyCode = vbKeyLeft Then
        Call SubRefresh(CMD_PREV)
    End If
End Sub

Private Sub Form_Load()
    '���ؽ���
    '����ʼ��
    mbytFontSize = 9
    mstrFontUnderLine = ""
    imgBtn(CMD_PREV).Enabled = False
    imgBtn(CMD_NEXT).Enabled = False
    imgBtn(CMD_ORTHER).Visible = False
    Set imgBtn(CMD_PREV).Picture = IIf(imgBtn(CMD_PREV).Enabled, imgFlag.ListImages("�ϴο���").Picture, imgFlag.ListImages("�ϴβ�����").Picture)
    Set imgBtn(CMD_NEXT).Picture = IIf(imgBtn(CMD_NEXT).Enabled, imgFlag.ListImages("�´ο���").Picture, imgFlag.ListImages("�´β�����").Picture)
    Set imgBtn(CMD_ORTHER).Picture = IIf(mbytShow = 0, imgFlag.ListImages("ֻ��ʾ����").Picture, imgFlag.ListImages("��ʾ���о���").Picture)
    '���ȱʡ����
    
    Call InitVsView
End Sub

Private Sub InitVsView()
'--------------------------------------------------------------------------------------------------------------------------------------------
'����:��ʼ��������
'--------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim arrTmp As Variant
    
    With vsView
        .Cols = 0: .Rows = 0
        .Cols = 3 * mlngSubCol + 1
        .Rows = SetViewRows
        .FixedRows = mcolCate("_" & CATE_���).lngEndRow + 1  '����ʱ��,���
        .FixedCols = 1
        .ExtendLastCol = False
        .RowHeightMin = 300
        .BackColorFixed = COLOR_��ɫ
        .FixedAlignment(0) = flexAlignCenterCenter
        .ColAlignment(0) = flexAlignCenterCenter
        .MergeCol(0) = True
        .MergeRow(mcolCate("_" & CATE_����ʱ��).lngBeginRow) = True
        .MergeRow(mcolCate("_" & CATE_���).lngBeginRow) = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .MergeCells = flexMergeFree
        .MergeCellsFixed = flexMergeRestrictRows
        .MergeCompare = flexMCExact
        .SelectionMode = flexSelectionFree
        .GridLines = flexGridFlat
 
        '���ط�����
        For i = 0 To mcolCate.Count - 1
            .Cell(flexcpText, mcolCate("_" & i).lngBeginRow, 0, mcolCate("_" & i).lngEndRow, 0) = mcolCate("_" & i).strName
            .Cell(flexcpFontBold, mcolCate("_" & i).lngBeginRow, 0, mcolCate("_" & i).lngEndRow, 0) = flexcpFontBold
            .Cell(flexcpFontSize, mcolCate("_" & i).lngBeginRow, 0, mcolCate("_" & i).lngEndRow, 0) = mbytFontSize
            .Cell(flexcpBackColor, mcolCate("_" & i).lngBeginRow, 0, mcolCate("_" & i).lngEndRow, 0) = &H8000000F
        Next
    End With
End Sub

Private Sub InitFrom()
'����:����ȹ̶�
    Dim lngWFrm As Long
    Dim lngHTop As Long
    
    On Error Resume Next
    '����ߴ硢��С
    PicCenter.Move 0, 0, Me.Width, Me.Height
    lngWFrm = PicCenter.Width
    If lngWFrm < 7035 Then
        lngWFrm = 7035
    End If
    
    lngHTop = 1000
    picTop.Move 0, 0, lngWFrm, lngHTop
    picBottom.Move 0, lngHTop + 45, lngWFrm, Me.Height - lngHTop - 45
    vsView.Move 0, 0, picBottom.Width, picBottom.Height
    imgBtn(CMD_NEXT).Width = 1050: imgBtn(CMD_NEXT).Height = 610
    imgBtn(CMD_PREV).Width = 1050: imgBtn(CMD_PREV).Height = 610
    imgBtn(CMD_ORTHER).Width = 2250: imgBtn(CMD_ORTHER).Height = 610
    
    lbl(0).Move lngWFrm / 2 - lbl(0).Width / 2, lngHTop / 2
    imgBtn(CMD_PREV).Move lngWFrm - (imgBtn(CMD_PREV).Width + imgBtn(CMD_NEXT).Width + 200), 850
    imgBtn(CMD_NEXT).Move lngWFrm - (imgBtn(CMD_NEXT).Width + 100), 850
    imgBtn(CMD_ORTHER).Move 45, 850
    
    '������ɫ����
    PicCenter.BackColor = COLOR_CENTERBK
    Me.BackColor = COLOR_FORMBK      '���ڱ�����ɫ
    'VS���²��� ˢ�½������½�ȡ�ַ�
    If Me.Visible Then
        Call LoadView
        Call ResizeVsView
    End If
End Sub

Private Sub Form_Resize()
    Dim lngSpace As Long
    
    On Error Resume Next
    Call InitFrom
End Sub

Private Sub imgBtn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim strKey As String
    Dim strRePic As String
    
    If imgBtn(Index).Enabled = False Then Exit Sub
    
    Select Case Index
    Case CMD_PREV
        strKey = "�ϴθ���"
    Case CMD_NEXT
        strKey = "�´θ���"
    Case CMD_ORTHER
        If mbytShow = 0 Then
            strKey = "ֻ��ʾ���Ƹ���"
        Else
            strKey = "��ʾ���о������"
        End If
    End Select
    Set imgBtn(Index).Picture = imgFlag.ListImages(strKey).Picture
    mintIndex = Index
End Sub

Private Sub ShowTipInfo(ByVal objHwnd As Long, ByVal strInfo As String)
    If strInfo <> "" Then
        Call zlCommFun.ShowTipInfo(objHwnd, strInfo, True, , 4500)
    Else
        Call zlCommFun.ShowTipInfo(0, strInfo)
    End If
End Sub


Private Sub imgBtn_Click(Index As Integer)
    SubRefresh Index
End Sub

Private Sub ResizeVsView()
    Dim i As Long, j As Long
    Dim lngMainW As Long
    Dim lngSubW As Long
    Dim lngCount As Long
    Dim lngRow As Long
    Dim strSpace As String
    Dim objImage As Object
    
    Dim udtCate As TYPE_CATE
    
    On Error Resume Next
    
    With vsView
        '�����п�
        strSpace = "  "    'һ���ո�
        .FontSize = mbytFontSize
        .RowHeightMin = IIf(mbytFontSize = 9, 300, 400)
        .Cell(flexcpFontSize, 0, 0, .Rows - 1, .Cols - 1) = mbytFontSize
        .ColWidth(0) = IIf(mbytFontSize = 9, 1000, 1200)
        lngMainW = (.Width - .ColWidth(0) - 480) / 3
        If lngMainW < 2800 Then lngMainW = 2800
        For i = 1 To .Cols - 1
            Select Case i Mod mlngSubCol
            Case SUBCOL_��ע + 1 '��ע�У���|��|�룩
                .ColWidth(i) = 300
                .ColAlignment(i) = flexAlignCenterCenter
                .MergeCol(i) = True      '�����ע�кϲ�
            Case SUBCOL_ͼ�� + 1  'һ����ҩ��ʶ��(������)
                .ColWidth(i) = IIf(mbytFontSize = 9, 200, 250)
                .ColAlignment(i) = flexAlignRightCenter
            Case SUBCOL_���� + 1   'ҩƷ��
                .ColWidth(i) = lngMainW - IIf(mbytFontSize = 9, 1160, 1215)
                .ColAlignment(i) = flexAlignLeftCenter
            Case SUBCOL_�Ƴ� + 1  '�Ƴ���Ϣ
                 .ColWidth(i) = 700
                .ColAlignment(i) = flexAlignCenterCenter
            Case Else
                '�����
                .ColWidth(i) = 15
                Call .Select(mcolCate("_" & CATE_��������).lngBeginRow, i, mcolCate("_" & CATE_����ҽ��).lngEndRow, i)
                Call .CellBorder(.GridColorFixed, 1, 1, 1, 1, -1, -1)
            End Select
            
        Next
        '����д���
        For i = CATE_��� To CATE_����ҽ��
            lngRow = mcolCate("_" & i).lngBeginRow - 1
            .RowHidden(lngRow) = True
        Next
        
        '��������ʱ���кϲ�����
        For i = 1 To 3
            udtCate = mcolCate("_" & CATE_����ʱ��)
            If .Cell(flexcpData, udtCate.lngBeginRow, (i - 1) * mlngSubCol + SUBCOL_���� + 1) = "" Then   '�Һŵ�Ϊ��ʱ�������û������
                udtCate = mcolCate("_" & CATE_����ʱ��)
                .Cell(flexcpText, udtCate.lngBeginRow, (i - 1) * 5 + 1, udtCate.lngEndRow, i * 5) = IIf(i Mod 2, strSpace, strSpace & strSpace)
                udtCate = mcolCate("_" & CATE_���)
                .Cell(flexcpText, udtCate.lngBeginRow, (i - 1) * 5 + 1, udtCate.lngEndRow, i * 5) = IIf(i Mod 2, strSpace, strSpace & strSpace)
            End If
        Next
        
        udtCate = mcolCate("_" & CATE_����ʱ��)
        .Cell(flexcpAlignment, udtCate.lngBeginRow, 1, udtCate.lngEndRow, .Cols - 1) = flexAlignCenterCenter   '����ʱ��
        udtCate = mcolCate("_" & CATE_���)
        .Cell(flexcpAlignment, udtCate.lngBeginRow, 1, udtCate.lngEndRow, .Cols - 1) = flexAlignLeftCenter     '���
        
        
        '��ҩ����
        udtCate = mcolCate("_" & CATE_��ҩ����)
        Call .Select(udtCate.lngEndRow, 1, udtCate.lngEndRow, .Cols - 1)
        Call .CellBorder(.GridColorFixed, -1, -1, -1, 1, -1, -1)
        '��������
        udtCate = mcolCate("_" & CATE_��������)
        Call .Select(udtCate.lngEndRow, 1, udtCate.lngEndRow, .Cols - 1)
        Call .CellBorder(.GridColorFixed, -1, -1, -1, 1, -1, -1)
        '���
        udtCate = mcolCate("_" & CATE_���)
        Call .Select(udtCate.lngEndRow, 1, udtCate.lngEndRow, .Cols - 1)
        Call .CellBorder(.GridColorFixed, -1, -1, -1, 1, -1, -1)
        '����
        udtCate = mcolCate("_" & CATE_����)
        Call .Select(udtCate.lngEndRow, 1, udtCate.lngEndRow, .Cols - 1)
        Call .CellBorder(.GridColorFixed, -1, -1, -1, 1, -1, -1)


        '���һ��
        Call .Select(0, 0, .Rows - 1, .Cols - 1)
        Call .CellBorder(.GridColorFixed, -1, -1, 1, 1, -1, -1)

        .AutoSize 0, .Cols - 1, , 45
        .Row = 0   '��ý���
        
    End With
End Sub

Private Sub picTop_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If mintIndex >= 0 Then
        SetImageDefault
    End If
End Sub

Private Sub vsView_Click()
    Dim lngҽ��ID As Long
    Dim lng����ID As Long
    Dim lng����ID As Long
    Dim str��鱨��ID As String
    Dim lngCol As Long
    Dim strMsg As String
    Dim strTmp As String
    
    Dim blnMoved As Boolean
    With vsView
         If .Col <= .FixedCols - 1 Then Exit Sub
         If .Row <= .FixedRows - 1 Then Exit Sub
         .Redraw = flexRDNone
         lngCol = IIf(.Col Mod mlngSubCol = 0, .Col - mlngSubCol, (.Col \ mlngSubCol) * mlngSubCol)
         If .Row >= mcolCate("_" & CATE_���).lngBeginRow And .Row <= mcolCate("_" & CATE_���).lngEndRow Then
            lng����ID = CLng(.ColData(lngCol + SUBCOL_���� + 1))
            lngҽ��ID = CLng(.Cell(flexcpData, .Row, lngCol + SUBCOL_���� + 1))
            If .Col Mod mlngSubCol = SUBCOL_�Ƴ� + 1 And .Cell(flexcpFontUnderline, .Row, .Col) = True Then
                '��Ƭ
                mrs����ʱ��.Filter = "����ID=" & lng����ID
                If NVL(mrs����ʱ��!ִ��״̬, 0) = 1 Then '��ɾ�������,��������Ƿ�ת��
                    blnMoved = zlDatabase.NOMoved("���˹Һż�¼", mrs����ʱ��!�Һŵ� & "")
                End If
               
                If CreateObjectPacs(gobjPublicPacs) Then
                    Call gobjPublicPacs.ShowImage(lngҽ��ID, Me, blnMoved)
                End If
            ElseIf .Col Mod mlngSubCol = SUBCOL_���� And Not .Cell(flexcpPicture, .Row, lngCol + SUBCOL_��ע + 1) Is Nothing Then
                '���ı���
                strTmp = .Cell(flexcpData, .Row, lngCol + SUBCOL_��ע + 1)
                lng����ID = CLng(Split(strTmp, "|")(0))
                str��鱨��ID = Split(strTmp, "|")(1)
                Call FuncEPRReport(Me, lngҽ��ID, "D", lng����ID, str��鱨��ID, 1)
            End If
         ElseIf .Row >= mcolCate("_" & CATE_����).lngBeginRow And .Row <= mcolCate("_" & CATE_����).lngEndRow Then
            If .Col Mod mlngSubCol = SUBCOL_���� And Not .Cell(flexcpPicture, .Row, lngCol + SUBCOL_��ע + 1) Is Nothing Then
                lngҽ��ID = CLng(.Cell(flexcpData, .Row, lngCol + SUBCOL_���� + 1))
                lng����ID = CLng(.Cell(flexcpData, .Row, lngCol + SUBCOL_��ע + 1))
                Call FuncEPRReport(Me, lngҽ��ID, "", lng����ID, , 1)
            End If
         ElseIf .Row >= mcolCate("_" & CATE_��������).lngBeginRow And .Row <= mcolCate("_" & CATE_��������).lngEndRow And .MousePointer = flexCustom Then
             
            strTmp = CStr(.Cell(flexcpData, .Row, lngCol + SUBCOL_���� + 1))
            If strTmp = "" Then Exit Sub
            If Len(strTmp) < 32 Then
                lng����ID = CLng(strTmp) '�ϰ没���鿴
                Call gobjRichEPR.ViewDocument(Me, lng����ID, False)
            ElseIf Len(strTmp) = 32 And Not gobjEmr Is Nothing Then
                '�°没��
                On Error Resume Next
                strMsg = gobjEmr.OpenOutEPR(strTmp)
                err.Clear: On Error GoTo 0
            End If
         End If
         .Redraw = flexRDDirect
    End With
End Sub

Private Sub vsView_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim lngColor As Long, lngclrg, k As Long, n As Long
    Dim r1 As Integer, g1 As Integer, b1 As Integer
    Dim r2 As Integer, g2 As Integer, b2 As Integer
    Dim rg As Integer, gg As Integer, bg As Integer
    
    Dim lngFontW As Long
    Dim lng��ID As Long
    Dim strContent As String
    
    Dim vRect As RECT, vRect1 As RECT, vRect2 As RECT

    If mcolCate Is Nothing Then Exit Sub
    
    With vsView
        If .RowHidden(Row) = True Then Exit Sub
        '�����б���ɫ����
        If mcolCate("_" & CATE_����ʱ��).lngBeginRow <= Row And Row <= mcolCate("_" & CATE_����ҽ��).lngEndRow And Col = 0 Then
            '��ȡ���ο�
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Right - 1
            vRect.Bottom = Bottom - 1
            'draw frame
            lngColor = SetBkColor(hDC, 0)
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, k

            ' get colors
            r1 = 250: g1 = 250: b1 = 250   '������ʼ
            r2 = 229: g2 = 229: b2 = 229   '������ֹ
            ' show color
            vRect2 = vRect
            vRect2.Bottom = vRect.Bottom - (vRect.Bottom - vRect.Top) / 2
            vRect1 = vRect2

            For k = vRect2.Top To vRect2.Bottom
                rg = r1 + (k - vRect2.Top) * (r2 - r1) / (vRect2.Bottom - vRect2.Top)
                gg = g1 + (k - vRect2.Top) * (g2 - g1) / (vRect2.Bottom - vRect2.Top)
                bg = b1 + (k - vRect2.Top) * (b2 - b1) / (vRect2.Bottom - vRect2.Top)
                lngclrg = RGB(rg, gg, bg)
                SetBkColor hDC, lngclrg
                vRect1.Top = k
                ExtTextOut hDC, vRect1.Left, vRect1.Top, ETO_OPAQUE, vRect1, " ", 1, k
            Next
            ' get colors
            r1 = 229: g1 = 229: b1 = 229   '������ʼ
            r2 = 250: g2 = 250: b2 = 250   '������ֹ
            ' show color
            vRect2 = vRect
            vRect2.Top = vRect.Top + (vRect.Bottom - vRect.Top) / 2
            vRect1 = vRect2
            For k = vRect2.Top To vRect2.Bottom
                rg = r1 + (k - vRect2.Top) * (r2 - r1) / (vRect2.Bottom - vRect2.Top)
                gg = g1 + (k - vRect2.Top) * (g2 - g1) / (vRect2.Bottom - vRect2.Top)
                bg = b1 + (k - vRect2.Top) * (b2 - b1) / (vRect2.Bottom - vRect2.Top)
                lngclrg = RGB(rg, gg, bg)
                SetBkColor hDC, lngclrg
                vRect1.Top = k
                ExtTextOut hDC, vRect1.Left, vRect1.Top, ETO_OPAQUE, vRect1, " ", 1, k
            Next

            SetBkColor hDC, lngColor
            '����Ԫ������浽��������
            strContent = .Cell(flexcpText, Row, Col)
            lblW.Caption = strContent: lblW.AutoSize = True
            vRect1.Top = vRect.Top + (vRect.Bottom - vRect.Top) / 2 - (lblW.Height / 2) / Screen.TwipsPerPixelY

            vRect1.Left = vRect.Left + (vRect.Right - vRect.Left) / 2 - (lblW.Width / 2) / Screen.TwipsPerPixelX

            TextOut hDC, vRect1.Left, vRect1.Top, strContent, LenB(StrConv(strContent, vbFromUnicode))
        End If

        If Not (Col >= 1 And Col < vsView.Cols - 1) Then Exit Sub
        If mcolCate("_" & CATE_��ҩ����).lngBeginRow <= Row And Row <= mcolCate("_" & CATE_��ҩ����).lngEndRow Then
            If Col Mod mlngSubCol = SUBCOL_��ע Then Exit Sub
           '����ұ���
            vRect.Left = Right - 1
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Bottom - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
            If Col Mod mlngSubCol = SUBCOL_���� + 1 Then
                If .Cell(flexcpData, Row, (Col \ mlngSubCol) * mlngSubCol + SUBCOL_���� + 1) <> "" Then
                    vRect.Left = Left
                    vRect.Top = Top + 1
                    vRect.Right = Left + (Right - Left) / 7 * Val(.Cell(flexcpData, Row, (Col \ mlngSubCol) * mlngSubCol + SUBCOL_�Ƴ� + 1))
                    vRect.Bottom = Bottom - 2
                    SetBkColor hDC, OS.SysColor2RGB(.Cell(flexcpData, Row, (Col \ mlngSubCol) * mlngSubCol + SUBCOL_��ע + 1))
                    '���þ������򱳾�ɫ
                    ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, "", 0, 0
                     '�ָ�����ɫΪ���屳��
                    SetBkColor hDC, OS.SysColor2RGB(.BackColor)
                    vRect.Left = Left + 1
                    vRect.Top = Top + 4
                    vRect.Right = Right - 1
                    vRect.Bottom = Bottom - 1

                    '���峤�ȳ����п�ʱ ��ȡ����+��...���ķ�ʽ��ʾ
                    strContent = .Cell(flexcpData, Row, Col)
                    lngFontW = .Cell(flexcpWidth, Row, Col)
                    strContent = GetSubString(strContent, lngFontW)
                    '����Ԫ������浽��������
                    TextOut hDC, vRect.Left, vRect.Top, strContent, LenB(StrConv(.Cell(flexcpData, Row, Col), vbFromUnicode))
                End If
            ElseIf Col Mod mlngSubCol = SUBCOL_�Ƴ� + 1 Then
                 'һ����ҩ�Ƴ̺ϲ���ʾ,Ҫ��������Ϻϲ�����,����ϲ����ݱ��������ڸ�
                If .Cell(flexcpText, Row, (Col \ mlngSubCol) * mlngSubCol + SUBCOL_ͼ�� + 1) = "��" Then
                    lngEnd = 0
                    lng��ID = CLng(.Cell(flexcpData, Row, (Col \ mlngSubCol) * mlngSubCol + SUBCOL_ͼ�� + 1))
                    For k = 1 To .Rows - 1
                        If lng��ID <> CLng(.Cell(flexcpData, Row - k, (Col \ mlngSubCol) * mlngSubCol + SUBCOL_ͼ�� + 1)) Then
                             k = k - 1
                             Exit For
                        End If
                    Next
                    
                    vRect.Top = Top - (Bottom - Top) * k
                    vRect.Left = Left
                    vRect.Right = Right
                    vRect.Bottom = Bottom - 1
                    If vRect.Top < (.RowPos(mcolCate("_" & CATE_���).lngEndRow) / Screen.TwipsPerPixelY + (Bottom - Top)) Then
                    '�ϲ����ο򳬹��̶���ʱ,ȡ�̶��б�Եֵ
                        vRect.Top = (.RowPos(mcolCate("_" & CATE_���).lngEndRow) / Screen.TwipsPerPixelY + (Bottom - Top))
                    End If
                    '���þ������򱳾�ɫ
                    ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, "", 0, 0
                    
                    strContent = "(" & .Cell(flexcpData, Row, Col) & "��)"
                    lblW.Caption = strContent: lblW.AutoSize = True
                    vRect.Top = vRect.Top + (vRect.Bottom - vRect.Top) / 2 - (lblW.Height / 2) / Screen.TwipsPerPixelY
                    vRect.Left = vRect.Left + (vRect.Right - vRect.Left) / 2 - (lblW.Width / 2) / Screen.TwipsPerPixelX
                    
                    TextOut hDC, vRect.Left, vRect.Top, strContent, LenB(StrConv(strContent, vbFromUnicode))
                    
                End If
            End If
        End If

        If mcolCate("_" & CATE_��������).lngBeginRow <= Row And Row <= mcolCate("_" & CATE_����ҽ��).lngEndRow Then
            If Col Mod mlngSubCol = SUBCOL_��ע Then Exit Sub
           '����ұ���
            vRect.Left = Right - 1
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Bottom - 1

            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

        End If

    End With
End Sub

Private Sub vsView_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngColor As Long
    
    Dim strInfo As String
    
    Dim arrTmp As Variant
    
    With vsView
        If mcolCate Is Nothing Then Exit Sub
        lngRow = .MouseRow: lngCol = .MouseCol
        If lngRow = -1 Or lngCol = -1 Then Exit Sub
        .MousePointer = flexDefault
        If mstrFontUnderLine <> "" Then
            arrTmp = Split(mstrFontUnderLine, "|")
            If UBound(arrTmp) >= 3 Then
                lngColor = Val(arrTmp(3))
            Else
                lngColor = vbBlack
            End If
            .Cell(flexcpForeColor, arrTmp(0), arrTmp(1), arrTmp(0), arrTmp(2)) = lngColor
            .Cell(flexcpFontUnderline, arrTmp(0), arrTmp(1), arrTmp(0), arrTmp(2)) = False
        End If
        
        If lngRow >= mcolCate("_" & CATE_���).lngBeginRow And lngRow <= mcolCate("_" & CATE_���).lngEndRow And lngCol > 0 Then
            If lngCol Mod mlngSubCol >= 1 And lngCol Mod mlngSubCol <= mlngSubCol - 1 Then
                If .Cell(flexcpText, lngRow, lngCol) <> "" And Right(.Cell(flexcpText, lngRow, lngCol), 3) = "..." Then
                    strInfo = .Cell(flexcpData, lngRow, (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_���� + 1)
                    ShowTipInfo .hwnd, strInfo
                    Exit Sub
                End If
            End If
            
        ElseIf lngRow >= mcolCate("_" & CATE_��ҩ����).lngBeginRow And lngRow <= mcolCate("_" & CATE_��ҩ����).lngEndRow Then
            If lngCol Mod mlngSubCol = (SUBCOL_���� + 1) Then
                If .Cell(flexcpData, lngRow, lngCol) <> "" Then
                    strInfo = .Cell(flexcpData, lngRow, lngCol)
                    ShowTipInfo .hwnd, strInfo
                    Exit Sub
                End If
            End If
        ElseIf lngRow >= mcolCate("_" & CATE_���).lngBeginRow And lngRow <= mcolCate("_" & CATE_���).lngEndRow Then
            If lngCol Mod mlngSubCol = SUBCOL_���� + 1 Then
                If .Cell(flexcpText, lngRow, lngCol) <> "" Then
                    strInfo = .Cell(flexcpData, lngRow, (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_ͼ�� + 1)
                    ShowTipInfo .hwnd, strInfo
                End If
                
                If Not .Cell(flexcpPicture, lngRow, (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_��ע + 1) Is Nothing Then
                    .Cell(flexcpFontUnderline, lngRow, (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_ͼ�� + 1, lngRow, lngCol) = True
                    mstrFontUnderLine = lngRow & "|" & (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_ͼ�� + 1 & "|" & lngCol & "|" & COLOR_����
                    .MousePointer = flexCustom
                End If
                Exit Sub
            ElseIf lngCol Mod mlngSubCol = (SUBCOL_�Ƴ� + 1) And Replace(.TextMatrix(lngRow, lngCol), vbTab, "") = "��Ƭ" Then
                'ÿ������ �Ƴ�һ��
                .Cell(flexcpFontUnderline, lngRow, lngCol) = True
                mstrFontUnderLine = lngRow & "|" & lngCol & "|" & lngCol & "|" & COLOR_����
                .MousePointer = flexCustom
            Else
                .MousePointer = flexDefault
            End If
        ElseIf lngRow >= mcolCate("_" & CATE_����).lngBeginRow And lngRow <= mcolCate("_" & CATE_����).lngEndRow Then
            If lngCol Mod mlngSubCol = (SUBCOL_���� + 1) Then
                If .Cell(flexcpText, lngRow, lngCol) <> "" Then
                    strInfo = .Cell(flexcpData, lngRow, (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_ͼ�� + 1)
                    ShowTipInfo .hwnd, strInfo
                End If
                    
                If Not .Cell(flexcpPicture, lngRow, (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_��ע + 1) Is Nothing Then
                    .Cell(flexcpFontUnderline, lngRow, (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_ͼ�� + 1, lngRow, lngCol) = True
                    mstrFontUnderLine = lngRow & "|" & (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_ͼ�� + 1 & "|" & lngCol & "|" & COLOR_����
                    .MousePointer = flexCustom
                End If
            End If
            Exit Sub
        ElseIf lngRow >= mcolCate("_" & CATE_��������).lngBeginRow And lngRow <= mcolCate("_" & CATE_��������).lngEndRow Then
            If lngCol Mod mlngSubCol = SUBCOL_���� + 1 And .TextMatrix(lngRow, lngCol) <> "" Then
                strInfo = .Cell(flexcpData, lngRow, (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_ͼ�� + 1)
                ShowTipInfo .hwnd, strInfo

                .Cell(flexcpFontUnderline, lngRow, (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_ͼ�� + 1, lngRow, lngCol) = True
                mstrFontUnderLine = lngRow & "|" & (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_ͼ�� + 1 & "|" & lngCol
                .Cell(flexcpForeColor, lngRow, (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_ͼ�� + 1, lngRow, (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_���� + 1) = COLOR_���
                .MousePointer = flexCustom
            End If
            Exit Sub
        ElseIf lngRow >= mcolCate("_" & CATE_����ҽ��).lngBeginRow And lngRow <= mcolCate("_" & CATE_����ҽ��).lngEndRow Then
            If .Cell(flexcpText, lngRow, lngCol) <> "" Then
                strInfo = .Cell(flexcpData, lngRow, (lngCol \ mlngSubCol) * mlngSubCol + SUBCOL_ͼ�� + 1)
                Call ShowTipInfo(.hwnd, strInfo)
            End If
            Exit Sub
        End If
        
        If strInfo = "" Then
            ShowTipInfo 0, strInfo
        End If
    
        If mintIndex >= 0 Then
            SetImageDefault
        End If
    End With
End Sub

Private Function InitRS() As ADODB.Recordset
'����:�����¼��
    Dim rs As ADODB.Recordset
    Dim strFields As String
    Dim strFieldName As String
    Dim lngLen As Long
    Dim FieldType As DataTypeEnum
    Dim i As Long, j As Long
    
    Dim arrField As Variant
    Dim arrSubFeld As Variant '�ֶ�����|�ֶ�����|�ֶγ��� ȱʡ�ֶ����� ΪadVarChar
    

    strFields = "ID|adVarChar|32,�Һ�ID|adBigInt|18,��������||100"

    
    Set rs = New ADODB.Recordset
    '-----------------------------------------
    With rs.Fields
        arrField = Split(strFields, ",")
        For i = LBound(arrField) To UBound(arrField)
            arrSubFeld = Split(arrField(i), "|")
            strFieldName = arrSubFeld(0)
            Select Case UCase(arrSubFeld(1) & "")
            Case UCase("adVarChar")
                FieldType = adVarChar
            Case UCase("adBigInt")
                FieldType = adBigInt
            Case Else
                FieldType = adVarChar
            End Select
            lngLen = Val(arrSubFeld(2))
            .Append strFieldName, FieldType, lngLen
        Next
    End With
    '---------------------------------------
    rs.CursorLocation = adUseClient
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenStatic
    rs.Open
    '----------------------------------
    Set InitRS = rs
End Function

Public Sub SetFontSize(ByVal bytSize As Byte)
    If mbytFontSize = IIf(bytSize = 0, 9, 12) Then Exit Sub
    mbytFontSize = IIf(bytSize = 0, 9, 12)
    lblW.FontSize = mbytFontSize
    Call ResizeVsView
    Call LoadView     '���¼�������
End Sub

Private Function GetSubString(ByVal strSource As String, ByVal lngShowW As Long) As String
'----------------------------------------------------------------------
'���ܣ��ַ���������ʾ���ʱ��ȡ������ʾ
'����:strSource-��Ҫ��ȡ�ĳ���
'    lngShowW-��ʾ���
'����:��ȡ���ַ��� ��ʽ����ȡ�ַ��� + ��...��
'----------------------------------------------------------------------
    Dim strRet As String
    Dim lngSingleWord As Long, lngSumWord As Long
    Dim lngSumLen As Long, i As Long
    
    Dim lngPosBegin As Long, lngPosEnd As Long, lngPosMid As Long
    Dim blnTag As Boolean
    
    If strSource = "" Then Exit Function
    
    lblW.AutoSize = True
    lblW.FontSize = mbytFontSize
    lblW.Caption = strSource
    lngSumWord = lblW.Width - 15     'ʵ���ַ����
    '�����ȴ�����ʾʱ��ȡ
    If lngSumWord > lngShowW Then
        lblW.Caption = "\"
        lngSingleWord = lblW.Width - 15         '�����ַ����
        lngShowW = lngShowW - lngSingleWord * 3      'Ԥ��ʡ�Ժ�"..."����
        lngSumLen = Len(strSource)              '���ַ��ַ�����

        lngPosBegin = 1: lngPosEnd = lngSumLen
        
        For i = 1 To lngSumLen
            lngPosMid = (lngPosBegin + lngPosEnd) \ 2
            lblW.Caption = Mid(strSource, 1, lngPosMid)
            
            If lblW.Width < lngShowW Then
                lngPosBegin = lngPosMid
                blnTag = True
            ElseIf lblW.Width > lngShowW Then
                lngPosEnd = lngPosMid
                blnTag = False
            End If
            
            If (lngPosBegin + lngPosEnd) \ 2 = lngPosMid Then
                lngPosMid = IIf(blnTag, lngPosMid, lngPosMid - 1)
                strRet = Mid(strSource, 1, lngPosMid) & "..."
                Exit For
            End If
        Next
        
    Else
        strRet = strSource
    End If
    GetSubString = strRet
End Function

Private Function CheckRegister() As Boolean
'����:��鲡�˾����¼���Ƿ�����������ҵľ����¼
'����:T-����;F-������
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    strSql = " Select 1  " & vbNewLine & _
        "       From ���˹Һż�¼ A " & vbNewLine & _
        "       Where a.����id = [1] And a.��¼���� = 1 And a.��¼״̬ = 1 And a.ִ�в���ID<>[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, mlng����ID)
    
    CheckRegister = rsTmp.RecordCount > 0
End Function

Private Sub SetImageDefault()
'����:����ȱʡͼƬ
    If imgBtn(mintIndex).Enabled = False Then Exit Sub
    Select Case mintIndex
    Case CMD_PREV
        Set imgBtn(mintIndex).Picture = imgFlag.ListImages("�ϴο���").Picture
    Case CMD_NEXT
        Set imgBtn(mintIndex).Picture = imgFlag.ListImages("�´ο���").Picture
    Case CMD_ORTHER
        Set imgBtn(mintIndex).Picture = imgFlag.ListImages(IIf(mbytShow = 0, "ֻ��ʾ����", "��ʾ���о���")).Picture
    End Select
    mintIndex = -1
End Sub
