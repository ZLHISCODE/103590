VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmBillSelect 
   Caption         =   "����ѡ����"
   ClientHeight    =   6240
   ClientLeft      =   1548
   ClientTop       =   1896
   ClientWidth     =   10800
   Icon            =   "FrmBillSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   10800
   StartUpPosition =   2  '��Ļ����
   Tag             =   "82"
   Begin VB.CommandButton cmdˢ�� 
      Caption         =   "ˢ��(&R)"
      Height          =   350
      Left            =   9660
      TabIndex        =   15
      Top             =   95
      Width           =   1100
   End
   Begin VB.CommandButton cmdDeptSel 
      Caption         =   "��"
      Height          =   300
      Left            =   7620
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   120
      Width           =   270
   End
   Begin VB.TextBox txtDept 
      Height          =   300
      Left            =   5940
      TabIndex        =   13
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "ȫѡ(&A)"
      Height          =   350
      Left            =   5565
      TabIndex        =   12
      Top             =   5415
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "ȫ��(&L)"
      Height          =   350
      Left            =   6675
      TabIndex        =   11
      Top             =   5415
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   120
      TabIndex        =   8
      Top             =   5430
      Width           =   1100
   End
   Begin VB.CommandButton Cmd���� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   8190
      TabIndex        =   7
      Top             =   5430
      Width           =   1100
   End
   Begin VB.CommandButton Cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   9330
      TabIndex        =   6
      Top             =   5415
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtp��ʼ���� 
      Height          =   285
      Left            =   930
      TabIndex        =   3
      Top             =   128
      Width           =   1665
      _ExtentX        =   2942
      _ExtentY        =   508
      _Version        =   393216
      CustomFormat    =   "yyyy��MM��dd��"
      Format          =   111214595
      CurrentDate     =   36734
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   2595
      Left            =   30
      TabIndex        =   0
      Top             =   2730
      Width           =   10785
      _ExtentX        =   19029
      _ExtentY        =   4572
      _Version        =   393216
      BackColor       =   16777215
      FixedCols       =   0
      FocusRect       =   0
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.DTPicker dtp�������� 
      Height          =   285
      Left            =   3780
      TabIndex        =   4
      Top             =   128
      Width           =   1665
      _ExtentX        =   2942
      _ExtentY        =   508
      _Version        =   393216
      CustomFormat    =   "yyyy��MM��dd��"
      Format          =   111214595
      CurrentDate     =   36734
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   5880
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2350
            MinWidth        =   882
            Picture         =   "FrmBillSelect.frx":0E42
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14012
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshHead 
      Height          =   2055
      Left            =   30
      TabIndex        =   9
      Top             =   480
      Width           =   10785
      _ExtentX        =   19029
      _ExtentY        =   3620
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label LblDepartment 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5565
      TabIndex        =   5
      Top             =   180
      Width           =   360
   End
   Begin VB.Label Lbl�������� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2970
      TabIndex        =   2
      Top             =   180
      Width           =   720
   End
   Begin VB.Label Lbl��ʼ���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ʼ����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   720
   End
   Begin VB.Image ImgLine_S 
      Height          =   45
      Left            =   30
      MousePointer    =   7  'Size N S
      Top             =   2670
      Width           =   10755
   End
End
Attribute VB_Name = "FrmBillSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnBootUp As Boolean '�����ɹ���־
Private mblnFirst As Boolean
Private mstrFind As String '��������
Private mstrUserPart As String '��������
Private mintLastRow As Integer '��һ��
Private mblnOpenCheckCbo As Boolean '�Ƿ������벿�ű���
Private mlngSelectCount As Long
Private mstrStart As String
Private mstrEnd As String
Private Const mlngModule = 1724
Private mintUnit As Integer                 '0��ɢװ��λ��1����װ��λ
Private mblnSuccess As Boolean
Private mstr���ķ��� As String
Private mint�ƻ����� As Integer
Private mlng�ⷿid As Long
Private mstrSelectNO As String
Private mstrStartDate As String, mstrEndDate As String
Private msngOldY As String
Private Const mstrCaption As String = "����ѡ����"

'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/12/27
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString

'----------------------------------------------------------------------------------------------------------

Public Function ShowCard(ByVal str���Ϸ��� As String, ByVal lng�ⷿID As Long, _
        ByVal int�ƻ����� As Integer, ByRef strSelectNo As String, _
        ByRef strStartDate As String, ByRef strEndDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------
    '����:��ָ���깺������ѡ��.
    '����:str���Ϸ���-��������(��ID����Ϊ׼)
    '     int�ƻ�����-�ƻ�����
    '����:strSelectNo-��ѡ��ĵ��ݺ�
    '     strStartDate-ѡ��Ŀ�ʼ����
    '     strEndDate-ѡ��Ľ�������
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '�޸�:2007/12/26
    '---------------------------------------------------------------------------------------------------------
    mlng�ⷿid = lng�ⷿID
    mblnSuccess = False
    mstr���ķ��� = str���Ϸ���
    mint�ƻ����� = int�ƻ�����
    mstrSelectNO = ""
    
    Me.Show vbModal
    
    ShowCard = mblnSuccess
    strSelectNo = mstrSelectNO
    strStartDate = mstrStartDate
    strEndDate = mstrEndDate
End Function
  
Private Sub cmdAllCls_Click()
    Dim intRow As Integer, intCol As Integer
    
    intCol = GetCol(mshHead, "ѡ��")
    
    mlngSelectCount = 0
    With mshHead
          For intRow = 1 To .Rows - 1
              If Trim(.TextMatrix(intRow, 0)) <> "" Then
                  .TextMatrix(intRow, intCol) = ""
              End If
          Next
      End With
   If mlngSelectCount = 0 Then
        Cmd����.Enabled = False
    Else
        Cmd����.Enabled = True
    End If
End Sub

Private Sub cmdAllSel_Click()
    Dim intRow As Integer, intCol As Integer
    intCol = GetCol(mshHead, "ѡ��")

    mlngSelectCount = 0
    With mshHead
          For intRow = 1 To .Rows - 1
              If Trim(.TextMatrix(intRow, 0)) <> "" Then
                  .TextMatrix(intRow, intCol) = "��"
                  mlngSelectCount = mlngSelectCount + 1
              End If
          Next
    End With
   If mlngSelectCount = 0 Then
        Cmd����.Enabled = False
    Else
        Cmd����.Enabled = True
    End If
End Sub

Private Sub cmdDeptSel_Click()
    If Select����("") = False Then Exit Sub
    OS.PressKey vbKeyTab
 
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, 4)
End Sub

Private Sub Cmd����_Click()
    Dim intRow As Integer
    Dim intCol As Integer
    intCol = GetCol(mshHead, "ѡ��")
    mstrSelectNO = ""
    With mshHead
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, intCol) <> "" Then
                mstrSelectNO = IIf(mstrSelectNO = "", "", mstrSelectNO & ",") _
                    & "'" & .TextMatrix(intRow, 0) & "'"
            End If
        Next
    End With
    mstrStartDate = Format(dtp��ʼ����.Value, "yyyy-mm-dd")
    mstrEndDate = Format(dtp��������.Value, "yyyy-mm-dd")
    mblnSuccess = True
    Unload Me
End Sub

Private Sub Cmdȡ��_Click()
    mblnSuccess = False
    mstrStartDate = "1991-01-01"
    mstrEndDate = "1991-01-01"
    Unload Me
End Sub

Private Sub cmdˢ��_Click()
    Call GetList
End Sub

Private Sub Dtp��������_Change()
    If Me.dtp��������.Value < Me.dtp��ʼ����.Value Then Me.dtp��������.Value = Me.dtp��ʼ����.Value
    mstrEnd = Format(Me.dtp��������.Value, "yyyy-MM-dd")
End Sub
Private Sub Dtp��ʼ����_Change()
    If Me.dtp��ʼ����.Value > Me.dtp��������.Value Then Me.dtp��ʼ����.Value = Me.dtp��������.Value
    mstrStart = Format(Me.dtp��ʼ����.Value, "yyyy-MM-dd")
End Sub

Private Sub Form_Activate()
    
    If mblnBootUp = False Then
        Unload Me
        Exit Sub
    End If
End Sub
Private Function Select����(ByVal strSeach As String) As Boolean
    '--------------------------------------------------------------------------------------------
    '����:ѡ����
    '����:strKey-��������
    '����:ѡ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/12/26
    '--------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long
    Dim objCtl As Object: Dim vRect As RECT
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    
    
    Set objCtl = txtDept
    vRect = zlControl.GetControlRect(objCtl.hwnd)
    lngH = objCtl.Height
    strKey = GetMatchingSting(strSeach)
    strTittle = "�깺����ѡ��"
    If strSeach = "" Then
        gstrSQL = "" & _
            "   Select ID,�ϼ�ID,����,����,����,λ��,to_char(����ʱ��,'yyyy-mm-dd') as ����ʱ��" & _
            "   From ���ű� " & _
            "   Where to_char(����ʱ��,'yyyy-MM-dd')='3000-01-01' and (վ��=[1] or վ�� is null) " & _
            "   start with �ϼ�id is null connect by prior id=�ϼ�id "
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 1, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, gstrNodeNo)
    Else
        gstrSQL = "" & _
            "   Select ID,�ϼ�ID,����,����,����,λ��,to_char(����ʱ��,'yyyy-mm-dd') as ����ʱ��" & _
            "   From ���ű� " & _
            "   Where to_char(����ʱ��,'yyyy-MM-dd')='3000-01-01' " & _
            "         and  (���� like [1] or ����  like [1] or  ����  like  [1]) and (վ��=[2] or վ�� is null) " & _
            "   order by ����"
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strKey, gstrNodeNo)
    End If
    
    If blnCancel = True Then
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgBox "û�������������깺����,����!"
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    objCtl.Text = zlStr.Nvl(rsTemp!����) & "-" & zlStr.Nvl(rsTemp!����)
    objCtl.Tag = zlStr.Nvl(rsTemp!Id)
     If objCtl.Enabled Then objCtl.SetFocus
    Select���� = True
End Function
Private Sub Form_Load()
    Dim rsDepend As New Recordset
    Dim strSQL As String
    
    mintUnit = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngModule, "0"))
    
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
    End With
    
    With mOraFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���, True)
        .FM_��� = GetFmtString(mintUnit, g_���, True)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�, True)
        .FM_���� = GetFmtString(mintUnit, g_����, True)
    End With
    
    mblnBootUp = False
    Select Case mint�ƻ�����
        Case 1
            mstrStart = Format(GetFirstDate(FirstDayOfMonth, sys.Currentdate), "yyyy-mm-dd")
        Case 2
            mstrStart = Format(GetFirstDate(FirstDayOfQuarter, sys.Currentdate), "yyyy-mm-dd")
        Case 3
            mstrStart = Format(GetFirstDate(FirstDayOfyear, sys.Currentdate), "yyyy-mm-dd")
        Case 4
            mstrStart = Format(GetFirstDate(FirstDayOfWeek, sys.Currentdate), "yyyy-mm-dd")
    End Select
    mblnFirst = True
    Me.dtp��ʼ���� = mstrStart
    Me.dtp��ʼ����.MaxDate = sys.Currentdate
    mstrEnd = Format(sys.Currentdate, "yyyy-MM-dd")
    Me.dtp�������� = sys.Currentdate
    Me.dtp��������.MaxDate = sys.Currentdate
    mblnBootUp = True
    Call SetDetal
    Call GetList
    RestoreWinState Me, App.ProductName, mstrCaption
    mblnFirst = False
End Sub

Public Function SetColWidth()
    Dim intCol As Integer
    With mshHead
                
        For intCol = 0 To .Cols - 1
            .ColAlignment(intCol) = flexAlignLeftCenter
            .ColAlignmentFixed(intCol) = 4
            If mblnFirst Then
                  .ColWidth(intCol) = 1000
            End If
        Next
        If mblnFirst Then
            .ColWidth(0) = 1000
            .ColWidth(1) = 1000
            .ColWidth(2) = 1200
            .ColWidth(3) = 2000
            .ColWidth(4) = 1000
            .ColWidth(6) = 1000
            
            .ColWidth(9) = 500
        End If
        .ColAlignment(9) = flexAlignCenterCenter
        
        .ColAlignment(GetCol(mshHead, "�ɹ����")) = flexAlignRightCenter
    End With
    
End Function

Private Sub Form_Resize()
    err = 0: On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    If Me.Height < 6380 Then
        Me.Height = 6380
    End If
    If Me.Width < 10000 Then
        Me.Width = 10000
    End If
    
    If Me.ImgLine_S.Top >= Me.ScaleHeight - cmdAllCls.Height * 3 - 100 Then
        Me.ImgLine_S.Top = Me.ScaleHeight - cmdAllCls.Height * 3 - 100
    End If
    If Me.ImgLine_S.Top <= 2000 Then
        Me.ImgLine_S.Top = 2000
    End If
    
    With cmdˢ��
        .Left = Me.ScaleWidth - .Width - 50
    End With
    With ImgLine_S
        .Left = 0
        .Width = Me.ScaleWidth
    End With
   
    With mshHead
        .Left = 50
        .Width = Me.ScaleWidth - 100
        .Height = ImgLine_S.Top - .Top
    End With
    
    With Cmdȡ��
        .Top = Me.ScaleHeight - .Height - IIf(stbThis.Visible, stbThis.Height, 0) - 100
        .Left = Me.ScaleWidth - .Width - 50
        Cmd����.Top = .Top
        Cmd����.Left = .Left - Cmd����.Width - 50
        cmdAllCls.Top = .Top
        cmdAllCls.Left = Cmd����.Left - cmdAllCls.Width * 2
        cmdAllSel.Top = .Top
        cmdAllSel.Left = cmdAllCls.Left - cmdAllSel.Width - 50
        cmdHelp.Top = .Top
    End With
    
    With mshDetail
        .Top = ImgLine_S.Top + ImgLine_S.Height + 100
        .Left = 50
        If Cmd����.Top - .Top - 50 <= 0 Then
            .Height = 0
        Else
            .Height = Cmd����.Top - .Top - 50
        End If
        .Width = mshHead.Width
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mstrCaption
End Sub

 Private Sub ImgLine_S_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
        msngOldY = y
End Sub

Private Sub ImgLine_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '�ָ�������
    
    If Button <> 1 Then Exit Sub
    
    With ImgLine_S
        If .Top + y < 2000 Then Exit Sub
        If .Top + y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + y - msngOldY
    End With
    Call Form_Resize
End Sub

Private Sub ImgLine_S_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
        msngOldY = 0
End Sub

Private Sub GetList()
    If mblnBootUp = False Then Exit Sub
    Dim rsList As New Recordset
    Dim strSQL As String
    Dim lng����ID As Long
    Dim dtStartDate As Date, dtEndDate As Date
    dtStartDate = Format(dtp��ʼ����.Value, "yyyy-mm-dd")
    dtEndDate = CDate(Format(dtp��������.Value, "yyyy-mm-dd") & " 23:59:59")
    lng����ID = Val(txtDept.Tag)
    
    On Error GoTo ErrHandle
    If mstr���ķ��� <> "" Then
        strSQL = "" & _
            " Select /*+ Rule*/ distinct A.NO,B.���� as ����,decode(�ƻ�����,1,'�¶ȼƻ�',2,'���ȼƻ�','��ȼƻ�') as �ƻ�����,rtrim(ltrim(to_char(Sum(nvl(c.���,0))," & mOraFMT.FM_��� & "))) as �ɹ����," & _
            "       A.����˵��,A.������ as ������,to_char(A.��������,'yyyy-MM-dd') as ��������,A.�����," & _
            "       to_char(A.�������,'yyyy-MM-dd') as �������,'' as ѡ�� " & _
            " From  ���ϲɹ��ƻ� A,���ű� B,���ϼƻ����� c,�������� d,������ĿĿ¼ M," & _
            "       Table(Cast(f_Num2List([1]) As zlTools.t_NumList)) J " & _
            " Where A.����id=B.id and a.id=c.�ƻ�id and c.����id=d.����id and (b.վ��=[6] or b.վ�� is null)" & _
            "       And A.����ID is not NULL  and a.����=1 And d.����id=M.id and (M.վ��=[6] or M.վ�� is null) and M.����id=J.Column_Value"
    Else
        strSQL = "" & _
            " Select distinct A.NO,B.���� as ����, decode(�ƻ�����,1,'�¶ȼƻ�',2,'���ȼƻ�','��ȼƻ�') as �ƻ�����,rtrim(ltrim(to_char(Sum(nvl(c.���,0))," & mOraFMT.FM_��� & "))) as �ɹ����," & _
            "       A.����˵��,A.������ as ������,to_char(A.��������,'yyyy-MM-dd') as ��������,A.�����," & _
            "       to_char(A.�������,'yyyy-MM-dd') as �������,'' as ѡ�� " & _
            " From ���ϲɹ��ƻ� A,���ű� B,���ϼƻ����� c,�������� d " & _
            " Where A.����id=B.id and a.id=c.�ƻ�id and c.����id=d.����id and (b.վ��=[6] or b.վ�� is null)" & _
            "       and a.����=1 "
    End If
    strSQL = strSQL & _
    "       And (A.������� between [2] and [3] )  And nvl(A.�ⷿid,[5])=[5] " & _
            IIf(lng����ID = 0, "", " And a.����id=[4]") & _
    " Group by A.no,B.����,A.����˵��,A.������,A.��������,A.�����,A.�������,A.�ƻ�����" & _
    " Order by to_char(A.�������,'yyyy-MM-dd') Desc,A.NO"
    
    Set rsList = zlDatabase.OpenSQLRecord(strSQL, mstrCaption, mstr���ķ���, dtStartDate, dtEndDate, lng����ID, mlng�ⷿid, gstrNodeNo)
    stbThis.Panels(2).Text = "��ǰ����" & rsList.RecordCount & "�ŵ���"
    
    With mshHead
        .Redraw = False
        Set mshHead.Recordset = rsList
        If .Rows = 1 Then
            .Rows = 2
            .Row = 1
        End If
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
    End With
    Call SetColWidth
    Call SetDetal
    mshHead_EnterCell
    mintLastRow = 0
    mlngSelectCount = 0
    mshHead.Redraw = True
    Cmd����.Enabled = False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshHead_EnterCell()
    Dim strSQL As String
    Dim rsDetail As New Recordset
    Dim lngLop As Integer
    If mintLastRow = mshHead.Row Then Exit Sub
    mintLastRow = mshHead.Row
    
    On Error GoTo ErrHandle
    If mstr���ķ��� <> "" Then
        strSQL = "" & _
                "Select /*+ Rule*/ A.���, ('['|| q.���� || ']' || q.����) as ������Ϣ,q.���,q.����," & _
                "       ltrim(to_char(A.�빺����/" & IIf(mintUnit = 0, "1", "b.����ϵ��") & "," & mOraFMT.FM_���� & "))  as �빺����," & _
                "       ltrim(to_char(A.�ƻ�����/" & IIf(mintUnit = 0, "1", "b.����ϵ��") & "," & mOraFMT.FM_���� & ")) as ��������," & _
                        IIf(mintUnit = 0, "Q.���㵥λ", "B.��װ��λ") & " as ��λ, " & _
                "       ltrim(to_char((" & IIf(mintUnit = 0, "1", "b.����ϵ��") & " * A.����)," & mOraFMT.FM_�ɱ��� & ")) as ����," & _
                "       ltrim(to_char(A.���," & mOraFMT.FM_��� & ")) as ��� " & _
                "   From ���ϼƻ����� A,�������� B,���ϲɹ��ƻ� c,�շ���ĿĿ¼ Q,������ĿĿ¼ M, " & _
                "       Table(Cast(f_Num2List([1]) As zlTools.t_NumList)) J " & _
                "   Where A.����id=B.����id and A.����id=q.id And a.�ƻ�id=c.id and c.����=1 And C.No =[2]  " & _
                "         And B.����id=M.id and M.����id=J.Column_Value" & _
                "   Order by A.���"
    Else
        strSQL = "" & _
                "Select A.���, ('['|| q.���� || ']' || q.����) as ������Ϣ,q.���,q.����," & _
                "       ltrim(to_char(A.�빺����/" & IIf(mintUnit = 0, "1", "b.����ϵ��") & "," & mOraFMT.FM_���� & "))  as �빺����," & _
                "       ltrim(to_char(A.�ƻ�����/" & IIf(mintUnit = 0, "1", "b.����ϵ��") & "," & mOraFMT.FM_���� & ")) as ��������," & _
                        IIf(mintUnit = 0, "Q.���㵥λ", "B.��װ��λ") & " as ��λ, " & _
                "       ltrim(to_char((" & IIf(mintUnit = 0, "1", "b.����ϵ��") & " * A.����)," & mOraFMT.FM_�ɱ��� & ")) as ����," & _
                "       ltrim(to_char(A.���," & mOraFMT.FM_��� & ")) as ��� " & _
                "   From ���ϼƻ����� A,�������� B,���ϲɹ��ƻ� c,�շ���ĿĿ¼ Q " & _
                "   Where A.����id=B.����id and A.����id=q.id And a.�ƻ�id=c.id and c.����=1 And C.No =[2]  " & _
                "   Order by A.���"
    End If
    Set rsDetail = zlDatabase.OpenSQLRecord(strSQL, mstrCaption, mstr���ķ���, mshHead.TextMatrix(mshHead.Row, 0))
    With rsDetail
        mshDetail.Rows = 2
        mshDetail.Redraw = False
        For lngLop = 0 To mshDetail.Cols - 1
            mshDetail.TextMatrix(1, lngLop) = ""
            mshDetail.ColAlignment(lngLop) = 1
        Next
        
        If Not .EOF Then
            For lngLop = 1 To .RecordCount
                mshDetail.TextMatrix(lngLop, 0) = zlStr.Nvl(!������Ϣ)
                mshDetail.TextMatrix(lngLop, 1) = zlStr.Nvl(!���)
                mshDetail.TextMatrix(lngLop, 2) = zlStr.Nvl(!����)
                mshDetail.TextMatrix(lngLop, 3) = zlStr.Nvl(!��λ)
                mshDetail.TextMatrix(lngLop, 4) = zlStr.Nvl(!�빺����)
                mshDetail.TextMatrix(lngLop, 5) = zlStr.Nvl(!��������)
                mshDetail.TextMatrix(lngLop, 6) = zlStr.Nvl(!����)
                mshDetail.TextMatrix(lngLop, 7) = zlStr.Nvl(!���)
                If lngLop = mshDetail.Rows - 1 Then mshDetail.Rows = mshDetail.Rows + 1
                .MoveNext
            Next
            If .RecordCount > 0 Then
                .MoveFirst
                mshDetail.Row = 1
                mshDetail.Col = 0
                mshDetail.ColSel = mshDetail.Cols - 1
            End If
        End If
        mshDetail.Redraw = True
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function SetDetal()
    Dim i As Long
    With mshDetail
        .Clear
        .Rows = 2
        .Cols = 8
        .TextMatrix(0, 0) = "������Ϣ"
        .TextMatrix(0, 1) = "���"
        .TextMatrix(0, 2) = "����"
        .TextMatrix(0, 3) = "��λ"
        .TextMatrix(0, 4) = "�깺����"
        .TextMatrix(0, 5) = "��������"
        .TextMatrix(0, 6) = "����"
        .TextMatrix(0, 7) = "���"
        For i = 0 To .Cols - 1
            .ColAlignment(i) = IIf(i <= 3, 1, 7)
            .ColAlignmentFixed(i) = 4
        Next
        If mblnFirst Then
            .ColWidth(0) = 2500
            .ColWidth(1) = 1000
            .ColWidth(2) = 1200
            .ColWidth(3) = 500
            .ColWidth(4) = 1000
            .ColWidth(5) = 1000
            .ColWidth(6) = 1000
            .ColWidth(7) = 1000
        End If
    End With
End Function
Private Sub mshHead_DblClick()
    Dim lngCol As Long
    If mshHead.TextMatrix(mshHead.Row, 0) = "" Then Exit Sub
    lngCol = GetCol(mshHead, "ѡ��")
    If mshHead.TextMatrix(mshHead.Row, lngCol) = "" Then
        mshHead.TextMatrix(mshHead.Row, lngCol) = "��"
        mlngSelectCount = mlngSelectCount + 1
    Else
        mshHead.TextMatrix(mshHead.Row, lngCol) = ""
        mlngSelectCount = mlngSelectCount - 1
    End If
   
    If mlngSelectCount = 0 Then
        Cmd����.Enabled = False
    Else
        Cmd����.Enabled = True
    End If
End Sub


Private Sub mshHead_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then mshHead_DblClick
End Sub

Private Sub txtDept_Change()
    txtDept.Tag = ""
End Sub

Private Sub txtDept_GotFocus()
    OS.OpenIme True
End Sub

Private Sub txtDept_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtDept.Tag) <> "" Or Trim(txtDept.Text) = "" Then
        OS.PressKey vbKeyTab
        Exit Sub
    End If
    If Select����(Trim(txtDept.Text)) = False Then Exit Sub
    OS.PressKey vbKeyTab
End Sub

Private Sub txtDept_LostFocus()
    OS.OpenIme False
End Sub
