VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSchemeEditEx 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4140
   ControlBox      =   0   'False
   Icon            =   "frmSchemeEditEx.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame fraMethod 
      BackColor       =   &H8000000E&
      Height          =   2175
      Left            =   1560
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
      Begin VB.CommandButton cmdMethodOK 
         Caption         =   "ȷ��"
         Height          =   300
         Left            =   1065
         TabIndex        =   14
         Top             =   1800
         Width           =   975
      End
      Begin VSFlex8Ctl.VSFlexGrid vsMethod 
         Height          =   1815
         Left            =   0
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   0
         Width           =   2055
         _cx             =   1993543209
         _cy             =   1993542785
         Appearance      =   0
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
         BackColorSel    =   4210752
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   0
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmSchemeEditEx.frx":000C
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   0   'False
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
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   45
      Index           =   0
      Left            =   495
      MousePointer    =   7  'Size N S
      TabIndex        =   11
      Top             =   2310
      Width           =   615
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   45
      Index           =   2
      Left            =   495
      TabIndex        =   10
      Top             =   2580
      Width           =   615
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   330
      Index           =   3
      Left            =   405
      TabIndex        =   9
      Top             =   2295
      Width           =   45
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   330
      Index           =   1
      Left            =   1155
      MousePointer    =   9  'Size W E
      TabIndex        =   8
      Top             =   2310
      Width           =   45
   End
   Begin VB.CommandButton cmdData 
      Caption         =   "��"
      Height          =   240
      Left            =   2475
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "ѡ����Ŀ(*)"
      Top             =   1950
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Left            =   525
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   315
      Left            =   3555
      Picture         =   "frmSchemeEditEx.frx":0048
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "ȡ��(Esc)"
      Top             =   1920
      Width           =   450
   End
   Begin VSFlex8Ctl.VSFlexGrid vsExt 
      Height          =   1845
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   4080
      _cx             =   7197
      _cy             =   3254
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
      BackColorSel    =   4210752
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   7
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSchemeEditEx.frx":05D2
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
      Begin MSComctlLib.ImageList img16 
         Left            =   1650
         Top             =   975
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSchemeEditEx.frx":06CD
               Key             =   "c0"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSchemeEditEx.frx":0C67
               Key             =   "c1"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSchemeEditEx.frx":1201
               Key             =   "o0"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSchemeEditEx.frx":179B
               Key             =   "o1"
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmd 
         Caption         =   "��"
         Height          =   240
         Left            =   3435
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "ѡ����Ŀ(*)"
         Top             =   1035
         Visible         =   0   'False
         Width           =   270
      End
   End
   Begin VB.ComboBox cbo�걾 
      Height          =   300
      Left            =   525
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.CommandButton cmdOK 
      Height          =   315
      Left            =   3015
      Picture         =   "frmSchemeEditEx.frx":1D35
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "ȷ��(F2)"
      Top             =   1920
      Width           =   450
   End
   Begin VB.Label Label1 
      Caption         =   "��___ζ"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Line lin 
      Index           =   0
      X1              =   2070
      X2              =   2745
      Y1              =   2385
      Y2              =   2385
   End
   Begin VB.Line lin 
      Index           =   1
      X1              =   2070
      X2              =   2745
      Y1              =   2415
      Y2              =   2415
   End
   Begin VB.Line lin 
      Index           =   2
      X1              =   2070
      X2              =   2745
      Y1              =   2445
      Y2              =   2445
   End
   Begin VB.Line lin 
      Index           =   3
      X1              =   2070
      X2              =   2745
      Y1              =   2475
      Y2              =   2475
   End
   Begin VB.Line lin 
      Index           =   4
      X1              =   2070
      X2              =   2745
      Y1              =   2505
      Y2              =   2505
   End
   Begin VB.Line lin 
      Index           =   5
      X1              =   2070
      X2              =   2745
      Y1              =   2535
      Y2              =   2535
   End
   Begin VB.Line lin 
      Index           =   6
      X1              =   2070
      X2              =   2745
      Y1              =   2565
      Y2              =   2565
   End
   Begin VB.Line lin 
      Index           =   7
      X1              =   2070
      X2              =   2745
      Y1              =   2595
      Y2              =   2595
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Left            =   105
      TabIndex        =   2
      Top             =   1980
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "frmSchemeEditEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'��ڲ�����
Private mlngHwnd As Long '���ڶ�λ�Ŀؼ����
Private mint������� As Integer '1-����,2-סԺ,3-�����סԺ
Private mint��Ч As Integer

'0-������,1-��������,4-�������
Private mintType As Integer

'��:��������ĿID
Private mlng��ĿID As Long

'��/��:���Ӷ�������,����ʱһ��Ϊ��
'      ���="��λ��1;������1,������2|��λ��2;������1,������2|...<vbTab>0-����/1-����/2-����"
'      ����="����ID1,����ID2,...;����ID",���п���û�и�������������
'      �������="��ĿID1,��ĿID2,...;����걾" ������°�LIS��ģʽ���ǣ�"��ĿID1|ָ��1|ָ��2...,��ĿID2|ָ��1|ָ��2...,...;����걾"
Private mstrExtData As String
Private mblnNew As Boolean  '�ж��Ƿ����¿�������Ŀʱ���룬����Ϊ���¼�ͷ����
'�룺�жϼ�������Ƿ�ʹ���°�LIS�ļ������ģʽ
Private mblnNewLIS As Boolean
'���ڲ�����
Private mblnOK As Boolean '��


Private mblnReturn As Boolean '�Ƿ��˻س�ȷ��
Private mblnNotAddNew As Boolean '�Ƿ���������
Private mfrmParent As Object

Private mblnChangeSel As Boolean

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long

'-----------------------------------------------------------------------------------------------------
Public Function ShowMe(ByVal frmParent As Object, ByVal lngHwnd As Long, ByVal intType As Integer, ByVal int��Ч As Integer, ByVal int������� As Integer, _
            Optional ByVal blnNewLIS As Boolean, Optional ByVal blnNew As Boolean, Optional ByVal lng��Ŀid As Long, Optional ByRef strExtData As String) As Boolean
'����:
'     frmParent         ������
'     lngHwnd           ���ڶ�λ�Ŀؼ����,�����øô���Ŀؼ�
'     intType           0-������,1-��������,4-������ϣ�5-��Ѫ����������Ҫ��д���븽���
'     int��Ч           ��Ҫ�����ҽ����Ч 0-������1-����
'     int�������       ��ҽ��Ҫ����Ĳ������� 1-����������ﲡ�ˣ���첡�ˣ��������˵�) 2-סԺ��ֻ��סԺ���ˣ�
'     blnNewLIS         �жϼ�������Ƿ�ʹ���°�LIS�ļ������ģʽ
'     blnNew            �ж��Ƿ����¿�������Ŀʱ���룬����Ϊ���¼�ͷ���롣 true-�¿�������Ŀʱ���룬 false-���¼�ͷ���루����ֻ��Լ��飬ֻ���°�LIS��ʹ�ã�blnNewLIS=true)��
'     lng��Ŀid         ��������ĿID
'���أ�
'     strExtData        ���Ӷ������� , ����ʱһ��Ϊ��
'                       ��� = "��λ��1;������1,������2|��λ��2;������1,������2|...<vbTab>0-����/1-����/2-����"
'                       ����="����ID1,����ID2,...;����ID",���п���û�и�������������
'                       �������="��ĿID1,��ĿID2,...;����걾" ������°�LIS��ģʽ���ǣ�"��ĿID1|ָ��1|ָ��2...,��ĿID2|ָ��1|ָ��2...,...;����걾"


    Set mfrmParent = frmParent
    mlngHwnd = lngHwnd
    mintType = intType
    mint��Ч = int��Ч
    mint������� = int�������
    mblnNewLIS = blnNewLIS
    mblnNew = blnNew
    mlng��ĿID = lng��Ŀid
    mstrExtData = strExtData
    mblnOK = False
    On Error Resume Next
    Me.Show 1, frmParent
    
    strExtData = mstrExtData
    ShowMe = mblnOK
End Function

Private Sub cbo�걾_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo�걾.ListIndex <> -1 Then
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        lngIdx = Cbo.MatchIndex(cbo�걾.Hwnd, KeyAscii)
        If lngIdx = -1 And cbo�걾.ListCount > 0 Then lngIdx = 0
        cbo�걾.ListIndex = lngIdx
    End If
End Sub

Private Sub cmd_Click()
'���ܣ�����Ŀѡ����
    Dim rsTmp As ADODB.Recordset, i As Long
    Dim strSql As String, strSQLItem As String
    Dim vPoint As PointAPI, blnCancel As Boolean, strҩƷ As String
    Dim strSamples As String
    
    On Error GoTo errH
    
    If mintType = 1 Then
        '���븽������:���ﲻ�ǵ���Ӧ��,��˲�����
        strSQLItem = _
            " From ������ĿĿ¼ A Where A.���='F' And A.ID<>" & mlng��ĿID & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                " And (A.������� IN([1],3) Or [1]=3 And Nvl(A.�������,0)<>0) And Nvl(A.ִ��Ƶ��,0) IN(0,[2])"
        
        strSql = "Select Distinct 0 as ĩ��,ID,�ϼ�ID,����,����,NULL as ��λ,NULL as ��ģ" & _
            " From ���Ʒ���Ŀ¼ Where ����=5 And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Start With ID In (Select ����ID" & strSQLItem & ") Connect by Prior �ϼ�ID=ID"
        strSql = strSql & " Union ALL" & _
            " Select Distinct 1 as ĩ��,A.ID,����ID as �ϼ�ID,A.����,A.����,A.���㵥λ as ��λ,A.�������� as ��ģ" & _
            strSQLItem & " Order By ����"
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 2, "����", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False, _
            mint�������, IIF(mint��Ч = 0, 2, 1))
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "δ�ҵ����õ�������Ŀ�����ȵ�������Ŀ���������á�", vbInformation, gstrSysName
            End If
            Exit Sub
        End If
        
        '����ظ�����
        i = vsExt.FindRow(CLng(rsTmp!ID))
        If i <> -1 And i <> vsExt.Row Then
            MsgBox "�ø��������Ѿ���������¼�롣", vbInformation, gstrSysName
            Exit Sub
        End If
        
        Call Set��������(vsExt.Row, rsTmp)
    ElseIf mintType = 4 Then
        '������Ŀ
        With Me.cbo�걾
            For i = 0 To .ListCount - 1
                strSamples = strSamples & ",'" & .List(i) & "'"
            Next
        End With
        If Len(strSamples) > 0 Then
            strSamples = Mid(strSamples, 2)
        Else
            strSamples = "''"
        End If
        
        strSQLItem = "From ������ĿĿ¼ A,������Ŀ�ο� C,���鱨����Ŀ D " & _
            "Where A.ID=D.������Ŀid(+) And D.������ĿID=C.��Ŀid(+)" & _
            " And A.���='C' And Nvl(A.����Ӧ��,0)=1" & _
            " And (A.������� IN([1],3) Or [1]=3 And Nvl(A.�������,0)<>0)" & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
            " And (C.�걾���� In (" & strSamples & ") Or C.�걾���� Is Null)"
        
        strSql = "Select Distinct 0 as ĩ��,ID,�ϼ�ID,����,����,' ' As ��������,' ' As �걾��λ" & _
            " From ���Ʒ���Ŀ¼ Where ����=5 And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Start With ID In (Select A.����ID " & strSQLItem & ") Connect by Prior �ϼ�ID=ID"
        strSql = strSql & " Union ALL" & _
            " Select Distinct 1 as ĩ��,A.ID,����ID as �ϼ�ID,A.����,A.����,A.�������� as ��������,A.�걾��λ " & strSQLItem & " Order By ����"
        
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 2, "������Ŀ", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False, mint�������)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "δ�ҵ����õļ�����Ŀ�����ȵ�������Ŀ���������á�", vbInformation, gstrSysName
            End If
            Exit Sub
        End If
        If rsTmp!�������� = "΢����" And vsExt.Rows > 2 Then
            If vsExt.RowData(2) <> 0 Or vsExt.Row > 1 Then '��������ֻ�ܿ�һ��΢������Ŀ
                MsgBox "΢������Ŀֻ�ܵ������룡", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        '����ظ�����
        i = vsExt.FindRow(CLng(rsTmp!ID))
        If i <> -1 And i <> vsExt.Row Then
            MsgBox "�ü�����Ŀ�Ѿ�¼�룡", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '�����������Ƿ���ͬ
        For i = 1 To vsExt.Rows - 1
            If vsExt.RowData(i) <> 0 And i <> vsExt.Row Then
                If Not (vsExt.TextMatrix(i, 1) = Nvl(rsTmp!��������) _
                    Or vsExt.TextMatrix(i, 1) = "" Or Nvl(rsTmp!��������) = "") Then
                    MsgBox "��������ͬ�������͵���Ŀ����������Ŀ�ļ�������Ϊ""" & vsExt.TextMatrix(i, 1) & """��", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        Next
        
        '���³�ʼ�걾
        If Not InitCombox(rsTmp!ID, Nvl(rsTmp!�걾��λ)) Then Exit Sub
        
        Call Set������Ŀ(vsExt.Row, rsTmp)
        If rsTmp("��������") = "΢����" Then
            mblnNotAddNew = True
            vsExt.Rows = 2
        Else
            mblnNotAddNew = False
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdData_Click()
'���ܣ�����Ŀѡ����
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim strSQLItem As String
    
    If mintType = 1 Then
        '����������Ŀ:���ﲻ�ǵ���Ӧ��,��˲�����
        strSQLItem = " From ������ĿĿ¼ A Where A.���='G'" & _
                " And (A.������� IN([2],3) Or [2]=3 And Nvl(A.�������,0)<>0) And A.ID<>[1]" & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)"

        strSql = "Select Distinct 0 as ĩ��,ID,�ϼ�ID,����,����,NULL as ��λ,NULL as ��������" & _
            " From ���Ʒ���Ŀ¼ Where ����=5 And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Start With ID In (Select ����ID" & strSQLItem & ") Connect by Prior �ϼ�ID=ID"
        strSql = strSql & " Union ALL" & _
            " Select Distinct 1 as ĩ��,A.ID,����ID as �ϼ�ID,A.����,A.����,A.���㵥λ as ��λ,A.�������� as ��������" & _
            strSQLItem & " Order By ����"
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 2, "������Ŀ", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False, _
            mlng��ĿID, mint�������)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "δ�ҵ�ƥ����Ŀ��", vbInformation, gstrSysName
            End If
            txtData.SetFocus: Exit Sub
        End If
        txtData.Tag = rsTmp!ID
        txtData.Text = "[" & rsTmp!���� & "]" & rsTmp!����
        cmdData.Tag = txtData.Text
        
        txtData.SetFocus
    ElseIf mintType = 4 Then
        '����걾
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdMethodOK_Click()
    Call vsMethod_KeyPress(vbKeyReturn)
End Sub

Private Sub cmdOK_Click()
    Dim blnSkip As Boolean
    Dim strMsg As String, strTmp As String
    Dim strSql As String, i As Long, j As Long
    Dim rsTmp As ADODB.Recordset
    
    
    Dim lngBegin As Long, lngEnd As Long
    Dim strData As String
    
    If mintType = 0 Then '��鲿λ���
        '�ռ���λ�����������
        With vsExt
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, 1) = 1 Then
                    If .TextMatrix(i, 2) = "" Then
                        .Row = i: .ShowCell .Row, .Col
                        MsgBox "û��Ϊ��鲿λ""" & .TextMatrix(i, 1) & """ȷ����鷽����", vbInformation, gstrSysName
                        vsExt.SetFocus: Exit Sub
                    End If
                    
                    strTmp = strTmp & "|" & .TextMatrix(i, 1) & ";" & .TextMatrix(i, 2)
                End If
            Next
            If strTmp = "" And vsExt.Editable <> flexEDNone Then
                MsgBox "������ѡ��һ����鲿λ��", vbInformation, gstrSysName
                vsExt.SetFocus: Exit Sub
            End If
            strTmp = Mid(strTmp, 2) & vbTab & 0
        End With
    ElseIf mintType = 1 Or mintType = 4 Then '����������������Ŀ��������Ŀ���걾
        If mintType = 1 Or mintType = 4 And mblnNewLIS = False Then
            For i = 1 To vsExt.Rows - 1
                If vsExt.RowData(i) <> 0 Then
                    strTmp = strTmp & "," & vsExt.RowData(i)
                End If
            Next
        ElseIf mintType = 4 And mblnNewLIS Then
            For i = 1 To vsExt.Rows - 1
                If vsExt.RowData(i) <> 0 And (Val(vsExt.Cell(flexcpChecked, i, 0)) = 1 Or Val(vsExt.TextMatrix(i, 3)) = 0) Then
                    strTmp = strTmp & IIF(Val(vsExt.TextMatrix(i, 3)) = 1, "|", ",") & vsExt.RowData(i)
                End If
            Next
        End If
        strTmp = Mid(strTmp, 2)
        If strTmp = "" And mintType = 4 Then
            MsgBox "����Ҫѡ��һ��������Ŀ��", vbInformation, gstrSysName
            vsExt.SetFocus: Exit Sub
        End If
        strTmp = strTmp & ";" & IIF(mintType = 4, Me.cbo�걾.Text, IIF(Val(txtData.Tag) = 0, "", Val(txtData.Tag)))
    End If
    
    mstrExtData = strTmp
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    If KeyCode = vbKeyEscape Then
        If fraMethod.Visible Then
            fraMethod.Visible = False
            vsExt.SetFocus
        Else
            Call cmdCancel_Click
        End If
    ElseIf KeyCode = vbKeyF2 Then
        If cmdOK.Enabled And cmdOK.Visible Then Call cmdOK_Click
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then 'CTRL+A
        If mintType = 0 Then
            vsExt.Cell(flexcpData, vsExt.FixedRows, 1, vsExt.Rows - 1, 1) = 1
            Set vsExt.Cell(flexcpPicture, vsExt.FixedRows, 1, vsExt.Rows - 1, 1) = img16.ListImages("c1").Picture
        End If
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyR Then 'CTRL+R
        If mintType = 0 Then
            vsExt.Cell(flexcpData, vsExt.FixedRows, 1, vsExt.Rows - 1, 1) = 0
            Set vsExt.Cell(flexcpPicture, vsExt.FixedRows, 1, vsExt.Rows - 1, 1) = img16.ListImages("c0").Picture
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '����������ָ�����������
    If InStr(",;|'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Resize()
    Dim lngMinRows As Long
    Dim lngRows As Long, i As Long
    
    On Error Resume Next
    
    fraBorder(0).Left = 0
    fraBorder(0).Top = 0
    fraBorder(0).Width = Me.ScaleWidth
    fraBorder(1).Top = fraBorder(0).Height
    fraBorder(1).Left = Me.ScaleWidth - fraBorder(1).Width
    fraBorder(1).Height = Me.ScaleHeight - fraBorder(0).Height * 2
    fraBorder(2).Left = 0
    fraBorder(2).Top = Me.ScaleHeight - fraBorder(2).Height
    fraBorder(2).Width = Me.ScaleWidth
    fraBorder(3).Top = fraBorder(0).Height
    fraBorder(3).Left = 0
    fraBorder(3).Height = Me.ScaleHeight - fraBorder(0).Height * 2
    
    vsExt.Left = fraBorder(3).Width
    vsExt.Top = fraBorder(0).Height + fraBorder(0).Height
    vsExt.Width = Me.ScaleWidth - fraBorder(3).Width * 2
    
    vsExt.Height = Me.ScaleHeight - fraBorder(2).Height * 2 - (cbo�걾.Height + 200)
    
    cbo�걾.Top = Me.ScaleHeight - fraBorder(2).Height - ((Me.ScaleHeight - fraBorder(0).Height * 2 - vsExt.Height) - cbo�걾.Height) / 2 - cbo�걾.Height
    
    txtData.Top = cbo�걾.Top
    lblData.Top = cbo�걾.Top + (cbo�걾.Height - lblData.Height) / 2
    cmdOK.Top = cbo�걾.Top + (cbo�걾.Height - cmdOK.Height) / 2
    cmdCancel.Top = cmdOK.Top
    
    lblData.Left = 200
    cbo�걾.Left = lblData.Left + lblData.Width + fraBorder(3).Width
    txtData.Left = cbo�걾.Left
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - cmdCancel.Height
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - fraBorder(1).Width * 3
    
            
    cbo�걾.Width = cmdOK.Left - cbo�걾.Left - 200
    txtData.Width = cbo�걾.Width
    cmdData.Top = txtData.Top + 30
    cmdData.Left = txtData.Left + txtData.Width - cmdData.Width - 45
    
    Me.Refresh
End Sub

Private Sub Form_Load()
    Dim blnMulti As Boolean, vRect As RECT
    Dim str���� As String, i As Long
    
    Me.Height = 2325
    
    '�߿�����
    For i = 0 To fraBorder.UBound
        fraBorder(i).BackColor = vbButtonFace
    Next
    Set lin(0).Container = fraBorder(0): Set lin(1).Container = fraBorder(0)
    Set lin(2).Container = fraBorder(1): Set lin(3).Container = fraBorder(1)
    Set lin(4).Container = fraBorder(2): Set lin(5).Container = fraBorder(2)
    Set lin(6).Container = fraBorder(3): Set lin(7).Container = fraBorder(3)
    lin(0).X1 = 0: lin(0).Y1 = 0: lin(0).X2 = Screen.Width: lin(0).Y2 = lin(0).Y1: lin(0).BorderColor = &H8000000F
    lin(1).X1 = 0: lin(1).Y1 = Screen.TwipsPerPixelY: lin(1).X2 = Screen.Width: lin(1).Y2 = lin(1).Y1: lin(1).BorderColor = &H8000000E
    lin(2).X1 = fraBorder(1).Width - Screen.TwipsPerPixelX: lin(2).Y1 = 0: lin(2).X2 = lin(2).X1: lin(2).Y2 = Screen.Height: lin(2).BorderColor = &H80000011
    lin(3).X1 = fraBorder(1).Width - Screen.TwipsPerPixelX * 2: lin(3).Y1 = 0: lin(3).X2 = lin(3).X1: lin(3).Y2 = Screen.Height: lin(3).BorderColor = &H80000010
    lin(4).X1 = 0: lin(4).Y1 = fraBorder(2).Height - Screen.TwipsPerPixelY: lin(4).X2 = Screen.Width: lin(4).Y2 = lin(4).Y1: lin(4).BorderColor = &H80000011
    lin(5).X1 = 0: lin(5).Y1 = fraBorder(2).Height - Screen.TwipsPerPixelY * 2: lin(5).X2 = Screen.Width: lin(5).Y2 = lin(5).Y1: lin(5).BorderColor = &H80000010
    lin(6).X1 = 0: lin(6).Y1 = 0: lin(6).X2 = lin(6).X1: lin(6).Y2 = Screen.Height: lin(6).BorderColor = &H8000000F
    lin(7).X1 = Screen.TwipsPerPixelX: lin(7).Y1 = 0: lin(7).X2 = lin(7).X1: lin(7).Y2 = Screen.Height: lin(7).BorderColor = &H8000000E
    
 
    If mint������� = 0 Then mint������� = 3
    mblnOK = False
    mblnNotAddNew = False
                
    '��ʼ�������ʽ
    If mintType = 0 Then
        If Not Init������ Then Unload Me: Exit Sub
    ElseIf mintType = 1 Then
        lblData.Visible = True
        txtData.Visible = True
        cmdData.Visible = True
        lblData.Caption = "����"
        If Not Init������Ŀ Then Unload Me: Exit Sub
    ElseIf mintType = 4 Then
        lblData.Visible = True
        lblData.Caption = "�걾"
        With cbo�걾
            .Left = txtData.Left: .Top = txtData.Top: .Width = txtData.Width
            .Visible = True
        End With
        If Not Init������� Then Unload Me: Exit Sub
        If Not InitCombox(DefaultValue:=Me.txtData) Then Unload Me: Exit Sub
    End If
    
    '��������
    If mintType = 0 Then
        If vsExt.Rows = vsExt.FixedRows + 1 Then
            If vsExt.Editable = flexEDNone Then
                'û�����ò�λʱ�����Զ�ȷ��
                Call cmdOK_Click: Exit Sub
            ElseIf vsExt.TextMatrix(vsExt.FixedRows, 1) <> "" Then
                'ֻ��һ����λ���Ҳ�λֻ��һ��������ѡʱ���Զ�ȷ��
                If vsExt.TextMatrix(vsExt.FixedRows, 1) <> "" Then
                    'ֻ��һ����λ���Զ�ѡ�иò�λ
                    vsExt.Cell(flexcpData, vsExt.FixedRows, 1) = 1
                    Set vsExt.Cell(flexcpPicture, vsExt.FixedRows, 1) = img16.ListImages("c1").Picture
                    '���û��Ĭ�Ϸ�����ֻ��һ������Ҳѡ��
                    str���� = GetOnlyOneMethod(vsExt.Cell(flexcpData, vsExt.FixedRows, 2))
                    If vsExt.TextMatrix(vsExt.FixedRows, 2) = "" And str���� <> "" Then
                        vsExt.TextMatrix(vsExt.FixedRows, 2) = str����
                    End If
                    If vsExt.TextMatrix(vsExt.FixedRows, 2) <> "" Then vsExt.TabStop = False
                    
                    'ֻ��һ��������ѡʱ���������Ҫ�������븽������Ҳ������
                    If vsExt.TextMatrix(vsExt.FixedRows, 2) <> "" And str���� <> "" Then
                        Call cmdOK_Click: Exit Sub
                    End If
                End If
            End If
        End If
    ElseIf mintType = 4 Then
        '������������⴦��
        blnMulti = Val(zlDatabase.GetPara(84, glngSys)) = 1 '�Ƿ�����һ��ҽ��������������Ŀ
        If Len(Trim(mstrExtData)) > 0 Then
            If Len(Trim(Split(mstrExtData, ";")(0))) > 0 And Not blnMulti Then
                vsExt.Enabled = False
                '���ֻ��һ���걾����ʾ������
                If cbo�걾.ListCount < 2 Then cmdOK_Click: Exit Sub
            End If
        End If
    End If
    
    '�ָ����Ի�
    Call RestoreWinState(Me, App.ProductName, mintType)
    
    
    '���嶨λ
    GetWindowRect mlngHwnd, vRect
    Me.Left = (vRect.Left - 1) * Screen.TwipsPerPixelX
    Me.Top = (vRect.Top - 1) * Screen.TwipsPerPixelY - Me.Height
    Call Form_Resize
    
End Sub

Private Function Init������Ŀ() As Boolean
'���ܣ���ʼ����������ʽ������
'������mstrExtData=��������������������Ŀ����Ϣ,���п���û�и���������Ϊ��ʱ��ʾ������������Ŀ
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, lng����ID As Long
    Dim arr����IDs As Variant, str����IDs As String
    Dim i As Long, j As Long
    
    On Error GoTo errH
    
    strSql = mstrExtData
    If strSql = "" Then strSql = ";"
    str����IDs = CStr(Split(strSql, ";")(0))
    lng����ID = Val(Split(strSql, ";")(1))
    
    '��������
    If str����IDs <> "" Then
        strSql = "Select /*+ Rule*/ A.ID,A.����,A.����,A.��������" & _
            " From ������ĿĿ¼ A" & _
            " Where A.���='F' And A.ID IN(Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & _
            " Order by A.����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, str����IDs)
        i = rsTmp.RecordCount
    End If
        
    vsExt.Clear
    vsExt.Rows = IIF(i = 0, 2, i + 1)
    vsExt.Cols = 2
    vsExt.FixedRows = 1: vsExt.FixedCols = 0
    vsExt.TextMatrix(0, 0) = "��������"
    vsExt.TextMatrix(0, 1) = "��ģ"
    vsExt.ColWidth(0) = 3200: vsExt.ColWidth(1) = 800
    vsExt.FixedAlignment(0) = 4: vsExt.FixedAlignment(1) = 4
    vsExt.ColAlignment(0) = 1: vsExt.ColAlignment(1) = 1
    vsExt.Editable = flexEDKbdMouse
    
    If str����IDs <> "" And i <> 0 Then
        arr����IDs = Split(str����IDs, ",") '����ԭ������˳��
        For i = 0 To UBound(arr����IDs)
            rsTmp.Filter = "ID=" & CStr(arr����IDs(i))
            If Not rsTmp.EOF Then
                j = j + 1
                vsExt.RowData(j) = CLng(rsTmp!ID)
                vsExt.TextMatrix(j, 0) = "[" & rsTmp!���� & "]" & rsTmp!����
                vsExt.Cell(flexcpData, j, 0) = vsExt.TextMatrix(j, 0) '���ڻָ���ʾ
                vsExt.TextMatrix(j, 1) = Nvl(rsTmp!��������, 0)
            End If
        Next
    End If
    
    '������Ŀ
    If lng����ID <> 0 Then
        strSql = "Select A.ID,A.����,A.����,�������� From ������ĿĿ¼ A Where A.���='G' And A.ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID)
        If rsTmp.Filter <> 0 Then rsTmp.Filter = 0
        If Not rsTmp.EOF Then
            txtData.Tag = rsTmp!ID
            txtData.Text = "[" & rsTmp!���� & "]" & rsTmp!����
            cmdData.Tag = txtData.Text '���ڻָ���ʾ
        End If
    End If
    
    vsExt.Row = 1: vsExt.Col = 0
    Init������Ŀ = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Init������() As Boolean
'���ܣ���ʼ����鲿λ����ʽ������
'������mstrExtData=������鲿λ����Ϣ,Ϊ��ʱ��ʾ�������������Ŀ
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, lngIdx As Long, i As Integer
    Dim str���� As String, str���� As String
    Dim arrData As Variant, strNoneRegion As String
    Dim blnNone As Boolean
    
    On Error GoTo errH
    
    '��ȡ�����Ŀ������Ϣ
    strSql = "Select ����,��������,ִ�б�� From ������ĿĿ¼ Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng��ĿID)
    str���� = rsTmp!��������
    str���� = rsTmp!����
    
    '��ȡ��鲿λ��Ϣ
    strSql = "Select B.����,A.��λ,A.����,A.Ĭ��,B.��ע,B.���� as ��鷽�� From ������Ŀ��λ A,���Ƽ�鲿λ B" & _
        " Where A.����=B.���� And A.��λ=B.���� And A.��ĿID=[1] And A.����=[2] Order by B.����,B.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng��ĿID, str����)
    blnNone = rsTmp.EOF
'    If rsTmp.EOF Then
'        '����ü����Ŀ��û�����ü�鲿λ,�������еĹ�ѡ��
'        strSQL = "Select ����,���� as ��λ,Null as ����,Null as Ĭ��,��ע,���� as ��鷽�� From ���Ƽ�鲿λ Where ����=[1] Order by ����,����"
'        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����)
'        If rsTmp.EOF Then
'            MsgBox "����Ŀ�ļ������""" & str���� & """����û�������κμ�鲿λ�����ȵ���鲿λ�����н������á�", vbInformation, gstrSysName
'            Exit Function
'        End If
'    End If
    With vsExt
        '��ʾ��׼�Ĳ�λ��Ĭ�Ϸ���
        If blnNone Then
            .HighLight = flexHighlightNever
            .Editable = flexEDNone
            .TabStop = False
        Else
            .HighLight = flexHighlightAlways
            .Editable = flexEDKbdMouse
        End If
        .WordWrap = True
        .FocusRect = flexFocusNone
        .BackColorSel = &HFFCC99
        .ForeColorSel = &H0&
        .FixedRows = 1: .FixedCols = 0
        .Rows = .FixedRows + 1: .Cols = 4
        .MergeCellsFixed = flexMergeFree: .MergeRow(0) = True
        .MergeCells = flexMergeFree: .MergeCol(0) = True
        If str���� = "����" Then
            .TextMatrix(0, 0) = "�걾����"
            .TextMatrix(0, 1) = "�걾����"
            .TextMatrix(0, 2) = "�������"
        Else
            .TextMatrix(0, 0) = "��鲿λ"
            .TextMatrix(0, 1) = "��鲿λ"
            .TextMatrix(0, 2) = "��鷽��"
        End If
        
        .TextMatrix(0, 3) = "��ע"
        .RowHeight(0) = 300
        .ColComboList(2) = "..."
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = 4: .ColAlignment(i) = 1
        Next
        Do While Not rsTmp.EOF
            If .TextMatrix(.Rows - 1, 1) <> rsTmp!��λ Then
                If .TextMatrix(.Rows - 1, 1) <> "" Then
                    .Rows = .Rows + 1
                End If
                .TextMatrix(.Rows - 1, 0) = zlCommFun.GetNeedName("" & rsTmp!����)
                .TextMatrix(.Rows - 1, 1) = rsTmp!��λ
                Set .Cell(flexcpPicture, .Rows - 1, 1) = img16.ListImages("c0").Picture
                .Cell(flexcpData, .Rows - 1, 2) = CStr(Nvl(rsTmp!��鷽��)) '������ѡ����ʹ��
                .TextMatrix(.Rows - 1, 3) = Nvl(rsTmp!��ע)
            End If
            If Nvl(rsTmp!Ĭ��, 0) = 1 Then '��"������1,������2,..."�ķ�ʽ��ʾ��λ��鷽��
                .TextMatrix(.Rows - 1, 2) = .TextMatrix(.Rows - 1, 2) & "," & Nvl(rsTmp!����)
                If Left(.TextMatrix(.Rows - 1, 2), 1) = "," Then
                    .TextMatrix(.Rows - 1, 2) = Mid(.TextMatrix(.Rows - 1, 2), 2)
                End If
            End If
            rsTmp.MoveNext
        Loop
        
        '�޸�ʱ�������е�����
        '  ���Ϊ�գ�Ҳ��������ǰ�ĵ���λ�����Ŀ����ʱҪ�������ķ�ʽ����ѡ��λ
        '  ���߶�����ǰ�ĵ���λ��Ŀ��ǿ�д�����ǰ�Ĳ�λ(û�з���)���ֻ�������ͬ����λ
        If mstrExtData <> "" Then
            arrData = Split(Split(mstrExtData, vbTab)(0), "|")
            For i = 0 To UBound(arrData)
                lngIdx = .FindRow(CStr(Split(arrData(i), ";")(0)), 1, 1, , True)
                If lngIdx <> -1 Then
                    '�ò�λ�ķ���:������ǰ������ֻ�в�λû�з���
                    If UBound(Split(arrData(i), ";")) >= 1 Then
                        .TextMatrix(lngIdx, 2) = Split(arrData(i), ";")(1)
                    Else
                        .TextMatrix(lngIdx, 2) = ""
                    End If
                    .Cell(flexcpData, lngIdx, 1) = 1 '�����ò�λ��ѡ��
                    Set .Cell(flexcpPicture, lngIdx, 1) = img16.ListImages("c1").Picture
                Else
                    '�ò�λ�����Ѳ�����
                    strNoneRegion = strNoneRegion & "," & Split(arrData(i), ";")(0)
                End If
            Next
        End If
        
        .Row = 1: .Col = 1
        .ShowCell .Row, .Col
        
        'ȷ�����ߴ�
        .AutoSize 0, .Cols - 1
        If .ColWidth(0) < 500 Then .ColWidth(0) = 500
        If .ColWidth(0) > 850 Then .ColWidth(0) = 850
        If .ColWidth(1) < 800 Then .ColWidth(1) = 800
        If .ColWidth(1) > 1600 Then .ColWidth(1) = 1600
        If .ColWidth(2) < 2500 Then .ColWidth(2) = 2500
        If .ColWidth(2) > 3500 Then .ColWidth(2) = 3500
        If .ColWidth(3) < 800 Then .ColWidth(3) = 800
        If .ColWidth(3) > 2000 Then .ColWidth(3) = 2000
        
        lngIdx = 0
        For i = 0 To .Cols - 1
            lngIdx = lngIdx + .ColWidth(i) + 15
        Next
        Me.Width = lngIdx + 90
        
        .Height = (.Rows - 1) * (.RowHeightMin + 15) + .RowHeight(0) + 60
        If Not blnNone Then
            If .Height < 1590 Then .Height = 1590 '����5�в�λ
            If .Height > 2865 + 50 Then .Height = 2865 + 50 '���10�в�λ
        End If
    End With
    
    Me.Height = (vsExt.Height + 90) + cmdOK.Height + (cmdOK.Height * 0.65)
    
    '�Ѳ����ڵĲ�λ��ʾ
    If strNoneRegion <> "" Then
        If str���� = "����" Then
            MsgBox "���²���걾����Ŀ�������Ѳ����ڣ�" & vbCrLf & Mid(strNoneRegion, 2), vbInformation, gstrSysName
        Else
            MsgBox "���¼�鲿λ����Ŀ�������Ѳ����ڣ�" & vbCrLf & Mid(strNoneRegion, 2), vbInformation, gstrSysName
        End If
    End If
    
    Init������ = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Init�������() As Boolean
'���ܣ���ʼ��������Ŀ
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, blnLis As Boolean
    Dim arrItems As Variant, strItems As String
    Dim i As Long, j As Long
    Dim strLIS As String
    Dim strTmp As String
    Dim colTmp As New Collection
    Dim strItemTmp As String
    Dim lng��ID As Long
    Dim Y As Long
    
    On Error GoTo errH
    
    strSql = mstrExtData
    If strSql = "" Then strSql = IIF(mlng��ĿID <> 0, mlng��ĿID, "") & ";"
    strItems = CStr(Split(strSql, ";")(0))
    Me.txtData.Text = Split(strSql, ";")(1)
    cmdData.Tag = txtData.Text
    
    If strItems <> "" Then
        '�ж��Ƿ����°�LISģʽ�������Ŀ
        If Not gobjLIS Is Nothing Then
            blnLis = gobjLIS.CheckLisSate
        End If
        If mblnNewLIS And blnLis Then
            strLIS = " Union All" & vbNewLine & _
                    "       Select e.Id, e.����, e.����, e.��������, ���������Ŀ.���� As ���,���������Ŀ.id as ��ID " & vbNewLine & _
                    "       From ���������Ŀ, ���鱨����Ŀ C, ���鱨����Ŀ D, ������ĿĿ¼ E" & vbNewLine & _
                    "       Where ���������Ŀ.Id = c.������Ŀid And c.������Ŀid = d.������Ŀid And d.������Ŀid = e.Id And e.�����Ŀ <> 1 And ���������Ŀ.Id <> e.Id"
            '�ֽ�����
            For i = 0 To UBound(Split(strItems, ","))
                strTmp = Split(strItems, ",")(i)
                If InStr(strTmp, "|") > 0 Then
                    colTmp.Add Mid(strTmp, InStr(strTmp, "|") + 1), "_" & Mid(strTmp, 1, InStr(strTmp, "|") - 1)
                    strItemTmp = strItemTmp & "," & Mid(strTmp, 1, InStr(strTmp, "|") - 1)
                Else
                    strItemTmp = strItemTmp & "," & strTmp
                End If
            Next
            strItems = Mid(strItemTmp, 2)
            Me.Height = Me.Height + 1200
            vsExt.Height = vsExt.Height + 1200
        End If
        strSql = "Select * From (With ���������Ŀ As (Select /*+ Rule*/ A.ID,A.����,A.����,A.��������, a.���� As ���,null as ��ID  From ������ĿĿ¼ A " & _
            " Where A.���='C' And Nvl(A.����Ӧ��,0)=1" & _
            " And (A.������� IN([2],3) Or [2]=3 And Nvl(A.�������,0)<>0)" & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
            " And A.ID In(Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))))" & _
            " Select * from ���������Ŀ" & _
            strLIS & _
            ") Order by ���,����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strItems, mint�������)
    End If
        
    vsExt.Clear
    If strItems <> "" Then
        vsExt.Rows = IIF(rsTmp.RecordCount = 0, 2, rsTmp.RecordCount + 1)
    Else
        vsExt.Rows = 2
    End If
    vsExt.Cols = 4
    vsExt.FixedRows = 1: vsExt.FixedCols = 0
    vsExt.TextMatrix(0, 2) = "������Ŀ"
    If mblnNewLIS Then
        vsExt.ColWidth(2) = 3700
        vsExt.ColWidth(0) = 300
    Else
        vsExt.ColWidth(2) = 4000
        vsExt.ColHidden(0) = True
    End If
    vsExt.ColHidden(1) = True
    vsExt.ColHidden(3) = True
    vsExt.FixedAlignment(2) = 4
    vsExt.ColAlignment(2) = 1
    vsExt.Editable = flexEDKbdMouse
    
    If Not rsTmp.EOF Then
        arrItems = Split(strItems, ",") '����ԭ������˳��
        For i = 0 To UBound(arrItems)
            rsTmp.Filter = "ID=" & arrItems(i)
            If Not rsTmp.EOF Then
                Y = vsExt.FindRow(CLng(rsTmp!ID))
                '�ظ���ָ�겻����
                If Y = -1 Then
                    j = j + 1
                    vsExt.RowData(j) = CLng(rsTmp!ID)
                    '����Ĭ�Ϲ�ѡ���Ҳ���ȡ��
                    vsExt.TextMatrix(j, 0) = " "
                    vsExt.Cell(flexcpBackColor, j, 0) = &H8000000F
                    vsExt.TextMatrix(j, 2) = "[" & rsTmp!���� & "]" & rsTmp!����
                    vsExt.Cell(flexcpData, j, 2) = vsExt.TextMatrix(j, 2) '���ڻָ���ʾ
                    vsExt.TextMatrix(j, 1) = Nvl(rsTmp!��������)
                    vsExt.TextMatrix(j, 3) = 0   '����
'                    If Nvl(rsTmp!��������) = "΢����" Then mblnNotAddNew = True '΢����ֻ�ܿ�һ��������Ŀ
                End If
                If mblnNewLIS Then
                    lng��ID = CLng(rsTmp!ID)
                    rsTmp.Filter = "��ID=" & CLng(rsTmp!ID)
                    Do While Not rsTmp.EOF
                        Y = vsExt.FindRow(CLng(rsTmp!ID))
                        '�ظ���ָ�겻����
                        If Y = -1 Then
                            j = j + 1
                            vsExt.RowData(j) = CLng(rsTmp!ID)
                            On Error Resume Next
                            strItemTmp = ""
                            strItemTmp = colTmp("_" & lng��ID)
                            On Error GoTo errH
                            If InStr("|" & strItemTmp & "|", "|" & CLng(rsTmp!ID) & "|") > 0 Then
                                vsExt.Cell(flexcpChecked, j, 0) = 1
                            ElseIf strItemTmp = "" And mblnNew Then  '��һ�ν���Ĭ�Ϲ�ѡ
                                vsExt.Cell(flexcpChecked, j, 0) = 1
                            Else
                                vsExt.Cell(flexcpChecked, j, 0) = 2
                            End If
                            '��������
                            vsExt.TextMatrix(j, 2) = "    [" & rsTmp!���� & "]" & rsTmp!����
                            vsExt.Cell(flexcpData, j, 2) = vsExt.TextMatrix(j, 2) '���ڻָ���ʾ
                            vsExt.TextMatrix(j, 1) = Nvl(rsTmp!��������)
                            If Nvl(rsTmp!��������) = "΢����" Then mblnNotAddNew = True '΢����ֻ�ܿ�һ��������Ŀ
                            vsExt.TextMatrix(j, 3) = 1    '����
                        Else
                            '����ظ���ָ�깴ѡ��ǰ���ָ��δ��ѡ����ɾ��ǰ���ָ����غ����ָ��
                            On Error Resume Next
                            strItemTmp = ""
                            strItemTmp = colTmp("_" & lng��ID)
                            On Error GoTo errH
                            If vsExt.Cell(flexcpChecked, Y, 0) = 1 And InStr("|" & strItemTmp & "|", "|" & CLng(rsTmp!ID) & "|") > 0 Then
                                vsExt.RemoveItem Y
                                vsExt.AddItem ""
                                vsExt.RowData(j) = CLng(rsTmp!ID)
                                vsExt.Cell(flexcpChecked, j, 0) = 1
                                '��������
                                vsExt.TextMatrix(j, 2) = "    [" & rsTmp!���� & "]" & rsTmp!����
                                vsExt.Cell(flexcpData, j, 2) = vsExt.TextMatrix(j, 2) '���ڻָ���ʾ
                                vsExt.TextMatrix(j, 1) = Nvl(rsTmp!��������)
                                If Nvl(rsTmp!��������) = "΢����" Then mblnNotAddNew = True '΢����ֻ�ܿ�һ��������Ŀ
                                vsExt.TextMatrix(j, 3) = 1    '����
                            End If
                        End If
                        rsTmp.MoveNext
                    Loop
                End If
            End If
        Next
    End If
    If j > 0 Then vsExt.Rows = j + 1
    
    vsExt.Row = 1: vsExt.Col = 2
    Init������� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitCombox(Optional ByVal strNewItemID As String = "", Optional ByVal DefaultValue As String = "") As Boolean
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, strTmp As String, lngItemCount As Long
    InitCombox = False
    
    On Error GoTo DBError
    strTmp = "": lngItemCount = 0
    For i = 1 To vsExt.Rows - 1
        If vsExt.RowData(i) <> 0 And (i <> vsExt.Row Or Len(strNewItemID) = 0) Then
            lngItemCount = lngItemCount + 1
            strTmp = strTmp & "," & vsExt.RowData(i)
        End If
    Next
    If Len(strNewItemID) > 0 Then
        lngItemCount = lngItemCount + 1
        strTmp = strTmp & "," & strNewItemID
    End If
    If Len(strTmp) > 0 Then strTmp = Mid(strTmp, 2)

    If lngItemCount = 0 Then
        strSql = "Select ���� From ���Ƽ���걾"
    Else
        strSql = "Select /*+ Rule*/ �걾����,Sum(1) From (" & _
            "   Select Distinct A.ID,B.���� As �걾����" & _
            "   From ������ĿĿ¼ A,���Ƽ���걾 B,������Ŀ�ο� C,���鱨����Ŀ D" & _
            "   Where A.ID=D.������ĿID(+) And D.������ĿID=C.��ĿID(+)" & _
            "        And (NVL(C.�걾����,'') Is Null Or NVL( C.�걾����,'')=B.����)  And A.ID In (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & _
            " ) Group By �걾���� Having Sum(1)=[2]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strTmp, lngItemCount)
    If rsTmp.EOF Then
        MsgBox Switch(lngItemCount = 0, "δ���ü���걾���뵽�ֵ�����������á�", _
            lngItemCount = 1, "ѡȡ�ļ�����Ŀδ�������걾�����ȵ�������Ŀ����������", _
            lngItemCount > 1, "ѡȡ�ļ�����Ŀ�ļ���걾��������Ŀ�Ĳ�һ�£����ȵ�������Ŀ����������"), vbInformation, gstrSysName
        Exit Function
    End If
    
    With cbo�걾
        strTmp = .Text
        
        .Clear
        Do While Not rsTmp.EOF
            .AddItem rsTmp(0)
            rsTmp.MoveNext
        Loop
        .ListIndex = 0
        On Error Resume Next
        If Len(DefaultValue) > 0 Then
            .Text = DefaultValue
        Else
            .Text = strTmp
        End If
    End With
    InitCombox = True
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, mintType)
    
    mlngHwnd = 0
    mintType = 0
    mint��Ч = 0
    mint������� = 0
    mblnNewLIS = False
    mblnNew = False
    mlng��ĿID = 0
    Set mfrmParent = Nothing
End Sub

Private Sub fraBorder_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If Index = 0 Then
            If Me.Height - Y < 2355 Or Me.Height - Y > 7200 Then Exit Sub
            Me.Top = Me.Top + Y
            Me.Height = Me.Height - Y
        ElseIf Index = 1 Then
            If Me.Width + x < 4140 Or Me.Width + x > 9600 Then Exit Sub
            Me.Width = Me.Width + x
        End If
        Call Form_Resize
    End If
End Sub

Private Sub txtData_GotFocus()
    zlControl.TxtSelAll txtData
End Sub

Private Sub txtData_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset, vRect As RECT
    Dim strSql As String, strLike As String
    Dim blnCancel As Boolean
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtData.Text = "" Then
            If mintType = 1 Then '�������Բ�����������Ŀ
                Call zlCommFun.PressKey(vbKeyTab)
            End If
            Exit Sub
        ElseIf txtData.Text = cmdData.Tag Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        
        '�Ż�
        strLike = gstrLike
        If Len(txtData.Text) < 2 Then strLike = ""
        
        If mintType = 1 Then
            '����������Ŀ
            strSql = _
                " Select Distinct A.ID,A.����,A.����,A.���㵥λ as ��λ,A.�������� as ��������" & _
                " From ������ĿĿ¼ A,������Ŀ���� B" & _
                " Where A.ID=B.������ĿID And A.���='G' And (A.������� IN([3],3) Or [3]=3 And Nvl(A.�������,0)<>0)" & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                    " And (A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]) And B.����=[4]" & _
                " Order by A.����"
            vRect = zlControl.GetControlRect(txtData.Hwnd)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "������Ŀ", False, "", "", False, False, True, vRect.Left, vRect.Top, txtData.Height, blnCancel, False, True, _
                UCase(txtData.Text) & "%", strLike & UCase(txtData.Text) & "%", mint�������, gbytCode + 1)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "δ�ҵ�ƥ����Ŀ��", vbInformation, gstrSysName
                End If
                txtData.Text = cmdData.Tag
                zlControl.TxtSelAll txtData
                Exit Sub
            End If
            txtData.Tag = rsTmp!ID
            txtData.Text = "[" & rsTmp!���� & "]" & rsTmp!����
            cmdData.Tag = txtData.Text
            
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf mintType = 4 Then
            '����걾
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        Call cmdData_Click
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtData_Validate(Cancel As Boolean)
'���ܣ��ָ���ʾԭ����
    If txtData.Text <> cmdData.Tag Then
        txtData.Text = cmdData.Tag
    End If
End Sub

Private Sub vsExt_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'����:��ʾѡ��ť,����֤��ǰ��Ԫ��ɼ�
     Dim strKey As String, lngҩ��ID As Long
     
    If mblnChangeSel = True Then Exit Sub
    '��֤��ǰ��Ԫ��ɼ�
    If NewRow >= vsExt.FixedRows And NewRow <= vsExt.Rows - 1 Then
        If vsExt.LeftCol >= vsExt.FixedCols And vsExt.LeftCol <= vsExt.Cols - 1 Then
            Call vsExt.ShowCell(NewRow, vsExt.LeftCol)
        End If
    End If
    
    If mintType = 1 Or mintType = 4 Then
        '��ʾ/��������ѡ��ť
        If NewCol = 0 And mintType = 1 Or NewCol = 2 And mintType = 4 Then
            cmd.Height = vsExt.CellHeight - 30
            cmd.Left = vsExt.CellLeft + vsExt.CellWidth - cmd.Width - 15
            cmd.Top = vsExt.CellTop + 15
            
            If mintType = 4 And mblnNewLIS Then
                If vsExt.TextMatrix(NewRow, 3) = "1" Then
                    cmd.Visible = False
                Else
                    cmd.Visible = True
                End If
            Else
                cmd.Visible = True
            End If
        Else
            cmd.Visible = False
        End If
        If cmd.Visible Then
            vsExt.FocusRect = flexFocusSolid
        Else
            vsExt.FocusRect = flexFocusLight
        End If
    End If
    
End Sub

Private Sub vsExt_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
'����:����ĳЩ�п�ķ�Χ
    If Row = -1 Then
        If mintType = 1 Or mintType = 4 Then
            Call vsExt_AfterRowColChange(-1, -1, vsExt.Row, vsExt.Col) 'ʹ��ť�ɼ���������ťλ��
        End If
    End If
End Sub

Private Sub vsExt_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If mintType = 0 Then
        If NewCol = 0 Or NewCol = 3 Then
            Cancel = True
            If NewRow <> OldRow Then vsExt.Row = NewRow
        End If
    End If
End Sub

Private Sub vsExt_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    If cmd.Visible Then cmd.Visible = False
    If fraMethod.Visible Then fraMethod.Visible = False
End Sub

Private Function GetOnlyOneMethod(ByVal strMethod As String) As String
'���ܣ����ݲ�λ�ķ������壬���ֻ��һ��������ѡ���򷵻ظ÷���
    Dim strTmp As String
    
    If strMethod = "" Then Exit Function
    strTmp = strMethod
    
    strTmp = Replace(strTmp, vbTab, ";")
    strTmp = Replace(strTmp, ",", ";")
    strTmp = Replace(strTmp, ";;", ";")
    strTmp = "<spdel>" & strTmp & "<spdel>"
    strTmp = Replace(strTmp, "<spdel>;", "")
    strTmp = Replace(strTmp, ";<spdel>", "")
    strTmp = Replace(strTmp, "<spdel>", "")
    
    If InStr(strTmp, ";") = 0 Then GetOnlyOneMethod = Mid(strTmp, 2)        'ȥ��ǰ��λ��Ӱ���
End Function

Private Sub vsExt_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strMethod As String, i As Long, j As Long
    Dim arrMethod As Variant, arrSub As Variant
    
    strMethod = vsExt.Cell(flexcpData, Row, Col)
    If strMethod = "" Then
        MsgBox "�ü�鲿λû�����ÿɹ�ѡ��ļ�鷽����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With vsMethod
        .Rows = 0
        arrMethod = Split(Replace(strMethod, vbTab, ";" & vbTab), ";")
        For i = 0 To UBound(arrMethod)
            arrSub = Split(arrMethod(i), ",")
            For j = 0 To UBound(arrSub)
                .Rows = .Rows + 1
                If j = 0 Then
                    If InStr(1, arrMethod(i), vbTab) > 0 Then
                        .MergeRow(.Rows - 1) = True
                        .RowData(.Rows - 1) = 2 '�����ǹ�ѡ��
                        
                        .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, 1) = Mid(arrSub(j), 3) '��һλ����Ӱ����־
                        If InStr("," & vsExt.TextMatrix(vsExt.Row, 2) & ",", "," & Mid(arrSub(j), 3) & ",") > 0 Then
                            Set .Cell(flexcpPicture, .Rows - 1, 0, .Rows - 1, 1) = img16.ListImages("c1").Picture
                            .Cell(flexcpData, .Rows - 1, 0) = 1
                        Else
                            Set .Cell(flexcpPicture, .Rows - 1, 0, .Rows - 1, 1) = img16.ListImages("c0").Picture
                            .Cell(flexcpData, .Rows - 1, 0) = 0
                        End If
                    Else
                        '�ų���
                        .MergeRow(.Rows - 1) = True
                        .RowData(.Rows - 1) = 1 '�������ų���
                        
                        .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, 1) = Mid(arrSub(j), 2) '��һλ����Ӱ����־
                        If InStr("," & vsExt.TextMatrix(vsExt.Row, 2) & ",", "," & Mid(arrSub(j), 2) & ",") > 0 Then
                            Set .Cell(flexcpPicture, .Rows - 1, 0, .Rows - 1, 1) = img16.ListImages("o1").Picture
                            .Cell(flexcpData, .Rows - 1, 0) = 1 '1Ϊѡ��
                        Else
                            Set .Cell(flexcpPicture, .Rows - 1, 0, .Rows - 1, 1) = img16.ListImages("o0").Picture
                            .Cell(flexcpData, .Rows - 1, 0) = 0
                        End If
                    End If
                Else
                    '��ѡ����
                    .RowData(.Rows - 1) = 3 '�����ǹ�ѡ����
                    
                    .Cell(flexcpText, .Rows - 1, 1) = Mid(arrSub(j), 2)
                    If InStr("," & vsExt.TextMatrix(vsExt.Row, 2) & ",", "," & Mid(arrSub(j), 2) & ",") > 0 Then
                        Set .Cell(flexcpPicture, .Rows - 1, 1) = img16.ListImages("c1").Picture
                        .Cell(flexcpData, .Rows - 1, 0) = 1
                    Else
                        Set .Cell(flexcpPicture, .Rows - 1, 1) = img16.ListImages("c0").Picture
                        .Cell(flexcpData, .Rows - 1, 0) = 0
                    End If
                End If
            Next
        Next
        
        .Row = 0: .Col = 1
        
        .Height = .Rows * (.RowHeightMin + 15) + 30
        If .Height > Me.ScaleHeight - 100 Then .Height = Me.ScaleHeight - 100
        If .Height < 3 * (.RowHeightMin + 15) + 30 Then .Height = 3 * (.RowHeightMin + 15) + 30
        
        .Width = (vsExt.Width - 30) - (vsExt.CellLeft + 15)
        .Left = vsExt.Left + vsExt.CellLeft + 15
        .Top = vsExt.Top + vsExt.CellTop + vsExt.CellHeight + 15
        
        If .Top + .Height > Me.ScaleHeight Then
            .Top = Me.ScaleHeight - .Height
        End If
        
        fraMethod.Top = .Top: .Top = 0
        fraMethod.Left = .Left: .Left = 0
        fraMethod.Width = .Width
        fraMethod.Height = .Height + cmdMethodOK.Height + 20
        cmdMethodOK.Top = .Height
        cmdMethodOK.Left = .Width - cmdMethodOK.Width - 20
        
        fraMethod.ZOrder
        fraMethod.Visible = True
        If fraMethod.Visible Then .SetFocus
    End With
End Sub

Private Sub vsExt_DblClick()
    If mintType = 0 Then
        If vsExt.Editable <> flexEDNone And vsExt.MouseCol = 1 And vsExt.MouseRow >= vsExt.FixedRows Then
            Call vsExt_KeyPress(vbKeySpace)
        End If
    End If
End Sub

Private Sub vsExt_GotFocus()
    If fraMethod.Visible Then fraMethod.Visible = False
    Call vsExt_AfterRowColChange(-1, -1, vsExt.Row, vsExt.Col) 'ʹ��ť�ɼ�
End Sub

Private Sub vsExt_KeyDown(KeyCode As Integer, Shift As Integer)
'���ܣ�ɾ��������
    Dim i As Long, j As Long, k As Long
    Dim intRow As Integer        '��Ч��
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim strKey As String, lngҩƷID As Long
    
   If KeyCode = vbKeyDelete Then
        If (mintType = 1 Or mintType = 4) And vsExt.RowData(vsExt.Row) <> 0 Then
            '������°�LIS�����Ŀģʽ��������ɾ������
            If mintType = 4 And mblnNewLIS Then
                If vsExt.TextMatrix(vsExt.Row, 3) = "1" Then Exit Sub
            End If
            If MsgBox("Ҫɾ����ǰ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            '��������Ŀģʽ����ͬʱɾ������
            If mintType = 4 And mblnNewLIS Then
                lngBegin = vsExt.Row + 1
                For j = vsExt.Row + 1 To vsExt.Rows - 1
                    If vsExt.TextMatrix(j, 3) <> "1" Then Exit For
                    lngEnd = j
                Next
                For j = lngEnd To lngBegin Step -1
                    vsExt.RowData(j) = 0
                    For i = 0 To vsExt.Cols - 1
                        vsExt.TextMatrix(j, i) = ""
                        vsExt.Cell(flexcpData, j, i) = ""
                    Next
                    If Not (vsExt.Rows = vsExt.FixedRows + 1 And j = vsExt.FixedRows) Then
                        vsExt.RemoveItem j
                    End If
                Next
            End If
            vsExt.RowData(vsExt.Row) = 0
            For i = 0 To vsExt.Cols - 1
                vsExt.TextMatrix(vsExt.Row, i) = ""
                vsExt.Cell(flexcpData, vsExt.Row, i) = ""
            Next
            If Not (vsExt.Rows = vsExt.FixedRows + 1 And vsExt.Row = vsExt.FixedRows) Then
                vsExt.RemoveItem vsExt.Row
            End If
            
            '���³�ʼ�걾
            If mintType = 4 Then InitCombox
        End If
    End If
End Sub

Private Sub vsExt_LostFocus()
    If Not ActiveControl Is cmd Then cmd.Visible = False
End Sub

Private Sub vsExt_KeyPress(KeyAscii As Integer)
'���ܣ��Ǳ༭״̬ʱ���Զ��ƶ���Ԫ��
    If KeyAscii = 13 Then
        KeyAscii = 0
        '��λ����һӦ���뵥Ԫ��
        If mintType = 0 Then
            If vsExt.Col <= 1 Then
                vsExt.Col = vsExt.Col + 1
            ElseIf vsExt.Col = 2 And vsExt.Row <= vsExt.Rows - 2 Then
                vsExt.Row = vsExt.Row + 1
                vsExt.Col = 1
            ElseIf vsExt.Col = 2 And vsExt.Row = vsExt.Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
                Exit Sub
            End If
        ElseIf mintType = 1 Or mintType = 4 Then
            If vsExt.Row = vsExt.Rows - 1 Then
                If vsExt.RowData(vsExt.Row) = 0 Or mblnNotAddNew Then
                    Call zlCommFun.PressKey(vbKeyTab)
                    Exit Sub
                Else
                    vsExt.AddItem ""
                End If
            End If
            If vsExt.Row + 1 <= vsExt.Rows - 1 Then
                vsExt.Row = vsExt.Row + 1
                If mintType = 1 Then
                    vsExt.Col = 0
                Else
                    vsExt.Col = 2
                End If
            End If
        End If
    ElseIf KeyAscii = Asc("*") Then
        If mintType = 0 Then
            If vsExt.Col = 2 Then
                Call vsExt_CellButtonClick(vsExt.Row, vsExt.Col)
            End If
        ElseIf mintType = 1 Or mintType = 4 Then
            KeyAscii = 0
            If cmd.Visible Then cmd_Click
        End If
    ElseIf KeyAscii = vbKeySpace Then
        If mintType = 0 Then
            If vsExt.Editable <> flexEDNone Then
                If vsExt.Col = 1 Then
                    If vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col) = 1 Then
                        vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col) = 0
                        Set vsExt.Cell(flexcpPicture, vsExt.Row, vsExt.Col) = img16.ListImages("c0").Picture
                    Else
                        vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col) = 1
                        Set vsExt.Cell(flexcpPicture, vsExt.Row, vsExt.Col) = img16.ListImages("c1").Picture
                        
                        '�Զ���������ѡ����
                        vsExt.Col = 2
                        Call vsExt_CellButtonClick(vsExt.Row, vsExt.Col)
                    End If
                ElseIf vsExt.Col = 2 Then
                    Call vsExt_CellButtonClick(vsExt.Row, vsExt.Col)
                End If
            End If
        End If
    End If
End Sub

Private Sub vsExt_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'���ܣ��ǻس�ȷ�����༭�Ĵ���(����Text:=EditText,��ValidateEdit�¼��л�û��)
    Dim strKey As String, lngҩ��ID As Long, i As Long
    
    If Not mblnReturn Then
        If mintType = 1 Or mintType = 4 Then
            If Col = 0 And mintType = 1 Or Col = 2 And mintType = 4 Then
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
                            
                '���³�ʼ�걾
                If mintType = 4 Then InitCombox
            End If
        End If
    End If
End Sub


Private Sub vsExt_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'���ܣ���������ȷ��
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, strSamples As String
    Dim blnCancel As Boolean, i As Long
    Dim vPoint As PointAPI, strLike As String, strҩƷ As String
    Dim strKey As String, lngҩ��ID As Long
    
    If KeyAscii = 13 Then
        mblnReturn = True '����ǰ��س�ȷ�ϱ༭
        KeyAscii = 0
        
        '�Ż�
        strLike = gstrLike
        If Len(vsExt.EditText) < 2 Then strLike = ""
        
        On Error GoTo errH
        
        If mintType = 1 Then
            '���븽������:���ﲻ�ǵ���Ӧ��,��˲�����
            strSql = _
                " Select Distinct A.ID,A.����,A.����,A.���㵥λ as ��λ,A.�������� as ��ģ" & _
                " From ������ĿĿ¼ A,������Ŀ���� B" & _
                " Where A.ID=B.������ĿID And A.���='F' And A.ID<>[3]" & IIF(strLike = "", "", " And Rownum<=100") & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                    " And (A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]) And B.����=[4]" & _
                    " And (A.������� IN([5],3) Or [5]=3 And Nvl(A.�������,0)<>0) And Nvl(A.ִ��Ƶ��,0) IN(0,[6])" & _
                " Order by A.����"
            vPoint = zlControl.GetCoordPos(vsExt.Hwnd, vsExt.CellLeft, vsExt.CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "����", False, "", "", False, False, True, vPoint.x, vPoint.Y, vsExt.CellHeight, blnCancel, False, True, _
                UCase(vsExt.EditText) & "%", strLike & UCase(vsExt.EditText) & "%", mlng��ĿID, gbytCode + 1, mint�������, IIF(mint��Ч = 0, 2, 1))
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "δ�ҵ�ƥ����Ŀ��", vbInformation, gstrSysName
                End If
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
                Exit Sub
            End If
            
            '����ظ�����
            i = vsExt.FindRow(CLng(rsTmp!ID))
            If i <> -1 And i <> Row Then
                MsgBox "�ø��������Ѿ���������¼�롣", vbInformation, gstrSysName
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
                Exit Sub
            End If
            
            Call Set��������(Row, rsTmp)
        ElseIf mintType = 4 Then
            '������Ŀ
            With Me.cbo�걾
                For i = 0 To .ListCount - 1
                    strSamples = strSamples & ",'" & .List(i) & "'"
                Next
            End With
            If Len(strSamples) > 0 Then
                strSamples = Mid(strSamples, 2)
            Else
                strSamples = "''"
            End If
            strSql = "Select A.ID,A.����,A.����,A.��������,A.�걾��λ" & _
                " From ������ĿĿ¼ A,������Ŀ���� C Where A.ID=C.������ĿID" & _
                " And (A.���� Like [1] Or C.���� Like [2] Or C.���� Like [2]) And C.����=[3]" & _
                " And A.���='C' And Nvl(A.����Ӧ��,0)=1" & _
                " And (A.������� IN([4],3) Or [4]=3 And Nvl(A.�������,0)<>0)" & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)"
            If strLike = "" Then
                '���������ü�������ʱ(����ƥ��),�����(+)����,����ҪGroup Byһ��(���)
                strSql = strSql & " Group by A.ID,A.����,A.����,A.��������,A.�걾��λ"
            End If
            
            strSql = "Select Distinct A.ID,A.����,A.����,A.�������� as ��������,A.�걾��λ" & _
                " From ������Ŀ�ο� D,���鱨����Ŀ E,(" & strSql & ") A" & _
                " Where A.ID=E.������Ŀid(+) And E.������ĿID=D.��Ŀid(+)" & _
                " And (D.�걾���� In (" & strSamples & ") Or D.�걾���� Is Null)" & _
                " Order by A.����"

            vPoint = zlControl.GetCoordPos(vsExt.Hwnd, vsExt.CellLeft, vsExt.CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "������Ŀ", False, "", "", False, False, True, vPoint.x, vPoint.Y, vsExt.CellHeight, blnCancel, False, True, _
                UCase(vsExt.EditText) & "%", strLike & UCase(vsExt.EditText) & "%", gbytCode + 1, mint�������)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "δ�ҵ�ƥ����Ŀ��", vbInformation, gstrSysName
                End If
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
                Exit Sub
            End If
            If rsTmp!�������� = "΢����" And vsExt.Rows > 2 Then
                If vsExt.RowData(2) <> 0 Or vsExt.Row > 1 Then '��������ֻ�ܿ�һ��΢������Ŀ
                    MsgBox "΢������Ŀֻ�ܵ������룡", vbInformation, gstrSysName
                    vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                    Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
                    Exit Sub
                End If
            End If
            
            '����ظ�����
            i = vsExt.FindRow(CLng(rsTmp!ID))
            If i <> -1 And i <> Row Then
                MsgBox "�ü�����Ŀ�Ѿ�¼�룡", vbInformation, gstrSysName
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
                Exit Sub
            End If
            
            '�����������Ƿ���ͬ
            For i = 1 To vsExt.Rows - 1
                If vsExt.RowData(i) <> 0 And i <> Row Then
                    If Not (vsExt.TextMatrix(i, 1) = Nvl(rsTmp!��������) _
                        Or vsExt.TextMatrix(i, 1) = "" Or Nvl(rsTmp!��������) = "") Then
                        MsgBox "��������ͬ�������͵���Ŀ����������Ŀ�ļ�������Ϊ""" & vsExt.TextMatrix(i, 1) & """��", vbInformation, gstrSysName
                        vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                        Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
                        Exit Sub
                    End If
                End If
            Next
            
            '���³�ʼ�걾
            If Not InitCombox(rsTmp!ID, Nvl(rsTmp!�걾��λ)) Then
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
                Exit Sub
            End If
            
            Call Set������Ŀ(Row, rsTmp)
            If rsTmp!�������� = "΢����" Then
                mblnNotAddNew = True
                vsExt.Rows = 2
            Else
                mblnNotAddNew = False
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Set��������(ByVal lngRow As Long, rsInput As ADODB.Recordset)
    '��������
    vsExt.EditText = "[" & rsInput!���� & "]" & rsInput!���� '��������ֱ��ƥ��ʱ�б�Ҫ
    
    vsExt.RowData(lngRow) = CLng(rsInput!ID)
    vsExt.TextMatrix(lngRow, 0) = "[" & rsInput!���� & "]" & rsInput!����
    vsExt.Cell(flexcpData, lngRow, 0) = vsExt.TextMatrix(lngRow, 0)
    vsExt.TextMatrix(lngRow, 1) = Nvl(rsInput!��ģ)
    
    '��һ������
    If vsExt.RowData(vsExt.Rows - 1) <> 0 And Not mblnNotAddNew Then vsExt.AddItem ""
    vsExt.Row = vsExt.Rows - 1: vsExt.Col = 0
End Sub

Private Sub Set������Ŀ(ByVal lngRow As Long, rsInput As ADODB.Recordset)
    Dim strSql As String, rsTmp As Recordset
    Dim i As Long, j As Long
    Dim lngBegin As Long, lngEnd As Long
    
    '������Ŀ
    '�����LIS�����Ŀģʽ����ɾ��������·��
    '��������Ŀģʽ����ͬʱɾ������
    If mblnNewLIS Then
        lngBegin = lngRow + 1
        For j = lngRow + 1 To vsExt.Rows - 1
            If vsExt.TextMatrix(j, 3) <> "1" Then Exit For
            lngEnd = j
        Next
        For j = lngEnd To lngBegin Step -1
            vsExt.RowData(j) = 0
            For i = 0 To vsExt.Cols - 1
                vsExt.TextMatrix(j, i) = ""
                vsExt.Cell(flexcpData, j, i) = ""
            Next
            If Not (vsExt.Rows = vsExt.FixedRows + 1 And j = vsExt.FixedRows) Then
                vsExt.RemoveItem j
            End If
        Next
    End If
    vsExt.EditText = "[" & rsInput!���� & "]" & rsInput!���� '��������ֱ��ƥ��ʱ�б�Ҫ
    
    vsExt.RowData(lngRow) = CLng(rsInput!ID)
    vsExt.TextMatrix(lngRow, 2) = "[" & rsInput!���� & "]" & rsInput!����
    vsExt.Cell(flexcpData, lngRow, 2) = vsExt.TextMatrix(lngRow, 2)
    vsExt.TextMatrix(lngRow, 1) = Nvl(rsInput!��������)
    vsExt.TextMatrix(lngRow, 0) = " "
    vsExt.Cell(flexcpBackColor, lngRow, 0) = &H8000000F
    vsExt.TextMatrix(lngRow, 3) = 0 '����
    
    If mblnNewLIS Then
        strSql = "" & vbNewLine & _
            "       Select e.Id, e.����, e.����, e.��������, a.���� As ���, a.Id As ��id" & vbNewLine & _
            "       From ������ĿĿ¼ a, ���鱨����Ŀ C, ���鱨����Ŀ D, ������ĿĿ¼ E" & vbNewLine & _
            "       Where a.Id = c.������Ŀid And c.������Ŀid = d.������Ŀid And d.������Ŀid = e.Id And e.�����Ŀ <> 1 And a.Id <> e.Id and a.id=[1]" & vbNewLine & _
            "       Order By ���, ����"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CLng(rsInput!ID))
        Do While Not rsTmp.EOF
            i = vsExt.FindRow(CLng(rsTmp!ID))
            '�ظ���ָ�겻����
            If i = -1 Then
                If vsExt.RowData(vsExt.Rows - 1) & "" <> "" Then vsExt.AddItem ""
                vsExt.RowData(vsExt.Rows - 1) = CLng(rsTmp!ID)
                vsExt.Cell(flexcpChecked, vsExt.Rows - 1, 0) = 1
                '��������
                vsExt.TextMatrix(vsExt.Rows - 1, 2) = "    [" & rsTmp!���� & "]" & rsTmp!����
                vsExt.Cell(flexcpData, vsExt.Rows - 1, 2) = vsExt.TextMatrix(vsExt.Rows - 1, 2) '���ڻָ���ʾ
                vsExt.TextMatrix(vsExt.Rows - 1, 1) = Nvl(rsTmp!��������)
    '                       If Nvl(rsTmp!��������) = "΢����" Then mblnNotAddNew = True '΢����ֻ�ܿ�һ��������Ŀ
                vsExt.TextMatrix(vsExt.Rows - 1, 3) = 1  '����
            End If
            
            rsTmp.MoveNext
        Loop
    End If
    
    '��һ������
    If vsExt.RowData(vsExt.Rows - 1) <> 0 And Not mblnNotAddNew Then vsExt.AddItem ""
    vsExt.Row = vsExt.Rows - 1: vsExt.Col = 2
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsExt_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim lngRow As Long, lngCol As Long
    Dim strTip As String
    
    If mintType = 0 Then
        lngRow = vsExt.MouseRow: lngCol = vsExt.MouseCol
        If Between(lngRow, 0, vsExt.Rows - 1) And Between(lngCol, 0, vsExt.Cols - 1) Then
            If vsExt.Cell(flexcpPicture, lngRow, lngCol) Is Nothing Then
                If Me.TextWidth(vsExt.TextMatrix(lngRow, lngCol)) > vsExt.ColWidth(lngCol) - 15 Then
                    strTip = vsExt.TextMatrix(lngRow, lngCol)
                End If
            Else
                If Me.TextWidth(vsExt.TextMatrix(lngRow, lngCol)) > vsExt.ColWidth(lngCol) - 15 - 240 Then
                    strTip = vsExt.TextMatrix(lngRow, lngCol)
                End If
            End If
        End If
        vsExt.ToolTipText = strTip
    End If
End Sub

Private Sub vsExt_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If mintType = 0 Then
        If vsExt.Col = 1 And vsExt.MouseCol = 1 Then
            If x <= vsExt.CellLeft + 250 Then
                Call vsExt_KeyPress(vbKeySpace)
            End If
        End If
    End If
End Sub

Private Sub vsExt_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsExt.EditSelStart = 0
    vsExt.EditSelLength = zlCommFun.ActualLen(vsExt.EditText)
End Sub

Private Sub vsExt_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'���ܣ�����ĳЩ�в�����༭(���¼�����BeforeEdit,��EditText��ֵ֮ǰ)
    mblnReturn = False
        
    If mintType = 0 Then
        'ֻ����ѡ���鷽��
        If Col <> 2 Then Cancel = True
    ElseIf mintType = 1 Or mintType = 4 Then
        'ֻ����༭��������
        If cmd.Visible Then cmd.Visible = False '��ʼ�༭�������ذ�ť
        If Col <> 0 And mintType = 1 Or Col <> 2 And Col <> 0 And mintType = 4 Then Cancel = True
        '����������°�LIS�������Ŀģʽ�������������
        If mblnNewLIS And mintType = 4 And Col = 2 Then
            If vsExt.TextMatrix(Row, 3) = "1" Then Cancel = True
        ElseIf mblnNewLIS And mintType = 4 And Col = 0 Then
            If Val(vsExt.TextMatrix(Row, 3)) = 0 Then Cancel = True
        End If
    End If
End Sub

Private Sub vsMethod_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewCol = 0 And NewRow <> -1 Then
        If vsMethod.TextMatrix(NewRow, 0) = "" Then
            Cancel = True
            vsMethod.Col = 1
        End If
    End If
End Sub

Private Sub vsMethod_Click()
    If fraMethod.Visible And vsMethod.Row >= 0 And vsMethod.Col >= 0 Then Call vsMethod_KeyPress(vbKeySpace)
End Sub

Private Sub vsMethod_KeyPress(KeyAscii As Integer)
    Dim strMethod As String
    Dim i As Long, j As Long
    Dim blnDo As Boolean
    
    With vsMethod
        If KeyAscii = 13 Then
            '��鷽����ȷ��
            For i = 0 To .Rows - 1
                If .Cell(flexcpData, i, 0) = 1 Then
                    strMethod = strMethod & "," & .TextMatrix(i, 1)
                End If
            Next
            If strMethod = "" Then Exit Sub
            vsExt.TextMatrix(vsExt.Row, 2) = Mid(strMethod, 2)
            vsExt.Cell(flexcpData, vsExt.Row, 1) = 1 '�������ú��Զ�ѡ�иò�λ
            Set vsExt.Cell(flexcpPicture, vsExt.Row, 1) = img16.ListImages("c1").Picture
            
            fraMethod.Visible = False
            vsExt.SetFocus
        ElseIf KeyAscii = vbKeySpace Then
            '��鷽����ѡ����ȡ��
            If .Cell(flexcpData, .Row, 0) = 1 Then
                '��ѡ��ĿǰҲ����ȡ��ѡ��
                .Cell(flexcpData, .Row, 0) = 0
                Set .Cell(flexcpPicture, .Row, IIF(.RowData(.Row) = 3, 1, 0), .Row, 1) = img16.ListImages(IIF(.RowData(.Row) = 1, "o0", "c0")).Picture
                'ͬʱȡ���õ�ѡ�������
                If .RowData(.Row) = 1 Then
                    For i = .Row + 1 To .Rows - 1
                        If .RowData(i) = 3 Then
                            If .Cell(flexcpData, i, 0) = 1 Then
                                .Cell(flexcpData, i, 0) = 0
                                Set .Cell(flexcpPicture, i, 1) = img16.ListImages("c0").Picture
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
            Else
                blnDo = True
                If .RowData(.Row) = 3 Then
                    '����û��ѡ��ʱ,�����ѡ��
                    For i = .Row - 1 To 0 Step -1
                        If .RowData(i) <> 3 Then
                            If .Cell(flexcpData, i, 0) = 0 Then blnDo = False
                            Exit For
                        End If
                    Next
                End If
                If blnDo Then
                    .Cell(flexcpData, .Row, 0) = 1
                    Set .Cell(flexcpPicture, .Row, IIF(.RowData(.Row) = 3, 1, 0), .Row, 1) = img16.ListImages(IIF(.RowData(.Row) = 1, "o1", "c1")).Picture
                    If .RowData(.Row) = 1 Then '��ѡ��ѡ��ʱ��ȡ��������ѡ��
                        For i = 0 To .Rows - 1
                            If i <> .Row And .RowData(i) = 1 Then
                                .Cell(flexcpData, i, 0) = 0
                                Set .Cell(flexcpPicture, i, 0, i, 1) = img16.ListImages("o0").Picture
                                For j = i + 1 To .Rows - 1 'ͬʱȡ���õ�ѡ�������
                                    If .RowData(j) = 3 Then
                                        If .Cell(flexcpData, j, 0) = 1 Then
                                            .Cell(flexcpData, j, 0) = 0
                                            Set .Cell(flexcpPicture, j, 1) = img16.ListImages("c0").Picture
                                        End If
                                    Else
                                        Exit For
                                    End If
                                Next
                            End If
                        Next
                    End If
                End If
            End If
        End If
    End With
End Sub
