VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmLCDShow 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12315
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   12315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picFace 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1695
      ScaleWidth      =   11535
      TabIndex        =   1
      Top             =   7200
      Width           =   11535
      Begin VB.Label labInf 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   975
         Left            =   1320
         TabIndex        =   2
         Top             =   120
         Width           =   11775
      End
   End
   Begin VB.Timer timerLCD 
      Interval        =   2000
      Left            =   7320
      Top             =   2040
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgCallingData 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      _cx             =   20770
      _cy             =   12303
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   0
      ForeColor       =   65280
      BackColorFixed  =   16777215
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   0
      BackColorAlternate=   0
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483627
      TreeColor       =   -2147483633
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmLCDShow.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
End
Attribute VB_Name = "frmLCDShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstr��������() As String
Private mint��Ч���� As Integer
Private mstr�������� As String, mstrҽ������ As String, mstrExcludeData As String

Private mlngCallItemsCount As Long    '���ж�����ʾ������
Private mintViewDataType As Integer '������ʾ����
Private mblnComeBackFirst As Boolean    '���ﲡ���Ƿ������Ŷ�
Private mlngLedLoopQueryTime As Long '��ѯ���ʱ�䳤��
Private mlngCallingColor As Long
Private mlngCalledColor As Long
Private mstrDelString As String '��Ҫɾ�����ַ���ʹ�á�,���ŷָ�




'��ʾ����

'----------------------------------------------------------------------------------------------------------------------------------
'|   ����    ����      ����      ����     ״̬
'|
'|   7001    ����      ����       CT      ������
'|   7002    ����      ����       DR      ������
'|
'|
'|
'|
'|
'|
'|
'|
'|  |-------------------------|-------------------------|-------------------------|-------------------------|
'|  |       �����            |       �����            |       Ƥ����            |       ������            |
'|  |-------------------------|-------------------------|-------------------------|-------------------------|
'|  |  (����)����     ����    |  (����)����     ����    |  (����)����     ����    |  (����)����     ����    |
'|  |-------------------------|-------------------------|-------------------------|-------------------------|
'|  |  (7001)����      CT(��) |                         |                         |                         |
'|  |  (7002)����      DR     |                         |                         |                         |
'|  |                         |                         |                         |                         |
'|  |                         |                         |                         |                         |
'|  |                         |                         |                         |                         |
'|  |                         |                         |                         |                         |
'|  |                         |                         |                         |                         |
'|  |                         |                         |                         |                         |
'|  |                         |                         |                         |                         |
'|  ---------------------------------------------------------------------------------------------------------
'|
'|      ף�����տ�����2010-06-04 13:51 ����һ
'|
'|
'|
'|
'|
'----------------------------------------------------------------------------------------------------------------------------------


Private Const colSplit = &H808080



'�ж������Ƿ�Ϊ��
Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long



Public Function zlShowMe(cnOracle As ADODB.Connection, str��������() As String, _
    Optional str���� As String = "", Optional strҽ�� As String = "", _
    Optional strExcludeData As String = "", Optional intViewDataType As Integer = 0, _
    Optional blnComeBackFirst As Boolean = True) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ʾ�Ŷ����
    '��Σ�str��������():�����ָ����������(��1��ʼ)
    '         strCur��������-��ǰ��������
    '         lngCurҵ��ID-ҵ��ID
    '         str����-����Ϊָ��������,����Ϊ�������:��"һ����,������,..."
    '         strҽ��-����Ϊ�ƶ���ҽ��,���Դ����ҽ��,�ö��ŷָ�,��"����,����,..."
    '         strExcludeData-�Ŷӵ�ָ��ҵ��ID
    '         intViewDataType������ʾ���ͣ�0��ʾ��ǰ�����µ��������ݣ�
    '                                      1��ʾ����Ϊ��ǰ������ҽ������Ϊ�գ�����ҽ���������ڵ�ǰҽ������������Ϊ�պ�ҽ��Ϊ�յ�����
    '                                      2��ʾ����Ϊ��ǰ���Һ�ҽ������Ϊ�ջ�ҽ���������ڵ�ǰҽ��������
    '                                      3��ʾ��ǰҽ��������
    '         blnComebackFirst���ﲡ���Ƿ������Ŷ�
    '���ƣ����˺�
    '���ڣ�2010-06-11 20:54:55
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    mstr�������� = str��������
    mstr�������� = str����
    mstrҽ������ = strҽ��
    mstrExcludeData = strExcludeData
    mintViewDataType = intViewDataType
    mblnComeBackFirst = blnComeBackFirst
    
    Me.Show
End Function

Public Function zlSetPara(str��������() As String, _
    Optional str���� As String = "", Optional strҽ�� As String = "", _
    Optional strExcludeData As String, Optional blnComeBackFirst As Boolean = True) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ʾ�Ŷ����
    '��Σ�str��������():�����ָ����������(��1��ʼ)
     '         str����-����Ϊָ��������,����Ϊ�������:��"һ����,������,..."
    '         strҽ��-����Ϊ�ƶ���ҽ��,���Դ����ҽ��,�ö��ŷָ�,��"����,����,..."
    '         strExcludeData-�Ŷӵ�ָ��ҵ��ID
    '���ƣ����˺�
    '���ڣ�2010-06-11 20:54:55
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    mstr�������� = str��������
    mstr�������� = str����
    mstrҽ������ = strҽ��
    mstrExcludeData = strExcludeData
    mblnComeBackFirst = blnComeBackFirst
End Function


Private Sub Form_Load()
    Dim i As Integer
    Dim strReg As String
    
    strReg = "����ģ��\�Ŷӽк�\Һ������"

    mlngLedLoopQueryTime = Val(GetSetting("ZLSOFT", strReg, "LED��ѯʱ��", "2"))
    timerLCD.Interval = mlngLedLoopQueryTime * 1000
    
    mlngCallItemsCount = Val(GetSetting("ZLSOFT", strReg, "���м�¼��ʾ��", "6"))
    
    mlngCallingColor = GetSetting("ZLSOFT", strReg, "��������ɫ", vbGreen)
    mlngCalledColor = GetSetting("ZLSOFT", strReg, "�Ѻ�����ɫ", &H408000)
    mstrDelString = GetSetting("ZLSOFT", strReg, "ɾ���ַ�", "")

    Me.BackColor = vbBlack
    
    '������ʾ����
    Call SetFaceFont
    '������ʾλ��
    Call SetFacePostion
    
    Call InitFace(mlngCallItemsCount)
End Sub


Private Sub Form_Resize()
    Call InitFace(mlngCallItemsCount)
End Sub


Public Sub SetFaceFont()
'************************************************************************************
'
'���ý�����ʾ��������ʽ
'
'
'************************************************************************************
    Dim strReg As String
    Dim curFontSize As Currency
    
    On Error Resume Next
    '��ע����У���ȡ��ʾ����
    strReg = "����ģ��\�Ŷӽк�\Һ������"

    vfgCallingData.Font.Name = GetSetting("ZLSOFT", strReg, "����", "����")
    vfgCallingData.Font.Bold = GetSetting("ZLSOFT", strReg, "����", "False")
    vfgCallingData.Font.Italic = GetSetting("ZLSOFT", strReg, "б��", "False")
    
    curFontSize = GetSetting("ZLSOFT", strReg, "�ֺ�", "14")
    vfgCallingData.Font.Size = curFontSize * 4.5 / 5
    
    
    labInf.Font.Name = vfgCallingData.Font.Name
    labInf.Font.Bold = vfgCallingData.Font.Bold
    labInf.Font.Italic = vfgCallingData.Font.Italic
    
    If curFontSize > 48 Then
        labInf.Font.Size = curFontSize - 16
    Else
        labInf.Font.Size = curFontSize
    End If
End Sub



Public Sub SetFacePostion()
'************************************************************************************
'
'���ý������ʾλ��
'
'
'************************************************************************************
    Dim strReg As String
    
    On Error Resume Next
        
    '��ע����У���ȡ��ʾ����
    strReg = "����ģ��\�Ŷӽк�\Һ������"
    
    '������ʾ����
    Me.Left = GetSetting("ZLSOFT", strReg, "��", "1024") * Screen.TwipsPerPixelX
    Me.Top = GetSetting("ZLSOFT", strReg, "��", "0") * Screen.TwipsPerPixelY
    Me.Width = GetSetting("ZLSOFT", strReg, "���", "1024") * Screen.TwipsPerPixelX
    Me.Height = GetSetting("ZLSOFT", strReg, "�߶�", "768") * Screen.TwipsPerPixelY
End Sub



Private Sub InitFace(ByVal lngCallItemsCount As Long)
'************************************************************************************
'
'��ʼ��������ʾ
'
'lngRoomsCount: ÿ��Ļ�ܹ���ʾ�Ŀ�������
'
'lngCallItemsCount���кŶ�����ʾ��������
'lngWaitItemsCount�����������ʾ��������
'
'************************************************************************************
    Dim i As Integer
    
    On Error GoTo errHandle
    
    '���ú�����ʾ-------------------------------------
    vfgCallingData.Top = 0
    vfgCallingData.Left = 140
    vfgCallingData.Width = Me.ScaleWidth - 140
    vfgCallingData.Height = Round(Me.ScaleHeight * 0.9)
    vfgCallingData.Rows = lngCallItemsCount + 3
    vfgCallingData.ForeColor = vbGreen
    vfgCallingData.Enabled = False
    
    vfgCallingData.Cols = 3
'
''    vfgCallingData.Cell(flexcpText, 0, 0) = "��  ��"
'    vfgCallingData.Cell(flexcpAlignment, 0, 0) = flexAlignLeftCenter
''    vfgCallingData.ColWidth(0) = Round(Me.Width / 5) + 1000
'
''    vfgCallingData.Cell(flexcpText, 0, 1) = "��  ��"
'    vfgCallingData.Cell(flexcpAlignment, 0, 1) = flexAlignLeftCenter
''    vfgCallingData.ColWidth(1) = Round(Me.Width / 5) - 1000
'
''    vfgCallingData.Cell(flexcpText, 0, 2) = "�������"
'    vfgCallingData.Cell(flexcpAlignment, 0, 2) = flexAlignLeftCenter
''    vfgCallingData.ColWidth(2) = Round(Me.Width / 5) * 3
'
'
'    Call vfgCallingData.AutoSize(0, vfgCallingData.Cols - 1)

    
    picFace.Top = vfgCallingData.Height + 40
    picFace.Left = 0
    picFace.Width = Me.ScaleWidth
    picFace.Height = Round(Me.ScaleHeight * 0.1)
        

        
    '��ʾ��Ϣ--------------------------------------
    labInf.Caption = "ף�����տ�����  " & Format(Now, "yyyy-mm-dd  hh:mm") & "  ����" & GetTodayNum
    
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub




Private Sub GetDepartKey(str��������() As String, strDepartKey() As String)
'**************************************************************************
'
'ȡ���Ŷӽк�ϵͳ���漰���Ŀ�������
'
'str��������()���к�ϵͳ�еĶ�����������
'
'strDepartKey()�������������
'
'**************************************************************************

    Dim strSQL As String
    Dim i As Integer
    Dim rsDepart As ADODB.Recordset
    Dim strDepartId As String
    
    On Error GoTo errHandle
    
    If SafeArrayGetDim(str��������) <= 0 Then
        Exit Sub
    End If
    
    If UBound(str��������) <= 0 Then
        Exit Sub
    End If
    
    'ȡ����Ҫ�����Ŀ���ID
    strDepartId = ""
    For i = 1 To UBound(str��������)
            
        If Trim(str��������(i)) <> "" Then
                                    
            If Trim(strDepartId) <> "" Then strDepartId = strDepartId & ","
                                    
            strDepartId = strDepartId & Mid(str��������(i) & ":", 1, InStr(1, str��������(i) & ":", ":") - 1)
        End If
    Next i
    
    If Trim(strDepartId) = "" Then
      Exit Sub
    End If
    
    strSQL = "select /*+ Rule*/ distinct ����, id from ���ű� a, Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) b where a.id =b.Column_Value order by Id"
    Set rsDepart = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", strDepartId)

    
    If rsDepart.RecordCount <= 0 Then
        Exit Sub
    End If
    
    ReDim strDepartKey(rsDepart.RecordCount)
    
    For i = 1 To rsDepart.RecordCount
        strDepartKey(i) = rsDepart!����
        rsDepart.MoveNext
    Next i
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub








'������ʽ��ʾ
Private Sub MultiRoomsDisplay()
'**************************************************************************
'
'��ʾ������ҵĽк���Ϣ
'
'
'**************************************************************************
    Dim i As Integer, j As Integer
    Dim blnSwitchPage As Boolean
    Dim blnAllowRoll As Boolean    '�жϵ�ǰ����ʾ�������Ƿ���Ҫ�����������������С�ڵ�ǰ�ܹ���ʾ�ļ�¼�������򲻹���
    Dim strCurRoomKey As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim rsSource As ADODB.Recordset   '��������
    Dim strValues(0 To 10) As String, strValue As String, strUninTable As String
    Dim strFilter As String, strCurExcludeData() As String
    Dim aryDelStr() As String
    Dim strRoomName As String
    
    
    err = 0: On Error GoTo errHandle
    If SafeArrayGetDim(mstr��������) > 0 Then
        j = 0
        
        strFilter = ""
        strValue = ""
        strUninTable = ""
        
        For i = 1 To UBound(mstr��������)
            If j > 10 Then
                strFilter = strFilter & " Or A.�������� ='" & mstr��������(i) & "'"
            Else
                If zlCommFun.ActualLen(strValue) > 2000 Then
                     strValues(j) = Mid(strValue, 2)
                     strUninTable = strUninTable & " Union ALL  Select  Column_Value as �������� From Table(Cast(f_Str2list([" & j + 3 & "]) As zlTools.t_Strlist))  " & vbCrLf
                     strValue = "": j = j + 1
                End If
                strValue = strValue & "," & mstr��������(i)
            End If
        Next i
        If strValue <> "" Then
            strValues(j) = Mid(strValue, 2)
            strUninTable = strUninTable & " Union ALL  Select  Column_Value as �������� From Table(Cast(f_Str2list([" & j + 3 & "]) As zlTools.t_Strlist))  " & vbCrLf
        End If
    End If
        
    
    If strUninTable <> "" Then
        strUninTable = Mid(strUninTable, 11)
    Else
       strUninTable = " Select  Column_Value as �������� From Table(Cast(f_Str2list([3]) As zlTools.t_Strlist)) "
    End If
    If strFilter <> "" Then strFilter = "( " & Mid(strFilter, 4) & ")"
    
    
    '�����ݿ��в�ѯ�����Ŷ����,ֻ��ѯ�����к��Ѻ��е�����
    
    strSQL = "" & _
    "   Select /*+ Rule*/  to_Number(a.ID) as ID, a.��������, b.���� as ��������, a.�ŶӺ��� As ����, a.�ŶӺ���, a.��������, a.ҽ������, a.����, " & _
    "               decode (a.�Ŷ�״̬,0,'�Ŷ���',1,'������',7,'�Ѻ���') as �Ŷ�״̬, to_Number(a.����) as ����, a.�Ŷ�ʱ��, a.����ʱ��, to_Number(�������) as �������, to_Number(a.ҵ������) as ҵ������, a.ҵ��ID, " & _
                    IIf(mblnComeBackFirst, "to_Number(Nvl(a.�������, 9999999999)) as ���������", "0 as ���������") & _
    " From �ŶӽкŶ��� a, ���ű� b , (" & strUninTable & ") E " & _
                IIf(mintViewDataType = 1, ", Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) C, Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) D ", "") & _
                IIf(mintViewDataType = 2, ", Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) C, Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) D ", "") & _
                IIf(mintViewDataType = 3, ", Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) D", "") & _
    " Where a.����id = b.Id  And a. ��������=E.��������  And (a.�Ŷ�״̬ = 1 or a.�Ŷ�״̬=7) and �Ŷ�ʱ�� <= trunc(sysdate + 1) - 1/24/60/60 " & _
                IIf(mintViewDataType = 1, " and  ((a.����=C.Column_Value and a.ҽ������ is null) or a.ҽ������=D.Column_Value or (a.���� is null and a.ҽ������ is null))", "") & _
                IIf(mintViewDataType = 2, " and ((a.����=C.Column_Value and a.ҽ������ is Null) or a.ҽ������=D.Column_Value) ", "") & _
                IIf(mintViewDataType = 3, " and a.ҽ������=D.Column_Value", "") & _
    " Order By a.�Ŷ�״̬, a.����ʱ�� desc"
    
    Set rsSource = zlDatabase.OpenSQLRecord(strSQL, "��ʾ�Ŷ����", mstr��������, mstrҽ������, strValues(0), strValues(1), strValues(2), strValues(3), strValues(4), strValues(5), strValues(6), strValues(7), strValues(8), strValues(9), strValues(10))
    
    On Error GoTo errCopyData
        Set rsTemp = zlDatabase.CopyNewRec(rsSource)
        GoTo readData
        
errCopyData:
        If Not rsTemp Is Nothing Then Set rsTemp = Nothing
        
        Call SaveErrLog
        
        Exit Sub
readData:
        
    If rsTemp.EOF Then
        vfgCallingData.Cell(flexcpText, 0, 0, vfgCallingData.Rows - 1, vfgCallingData.Cols - 1) = ""
        Exit Sub
    End If
        
        
    While Not rsTemp.EOF
        If InStr(1, mstrExcludeData, rsTemp!ҵ������ & ":" & rsTemp!ҵ��ID) > 0 Then
            rsTemp.Delete
        End If
        
        rsTemp.MoveNext
    Wend
    
        
    '�ڵ�����Ļ����ʾǰ�������е�����

    rsTemp.Sort = "�Ŷ�״̬ asc, ����ʱ�� desc"
    If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
    
    '��ȡ��Ҫɾ�����ַ���������
    If Trim(mstrDelString) <> "" Then
        mstrDelString = mstrDelString & ","
        aryDelStr() = Split(mstrDelString, ",")
    End If
    
    vfgCallingData.Redraw = flexRDNone
    For i = 0 To mlngCallItemsCount - 1
        'ɾ������ʱ��Ϊ�յ�����
        While Not rsTemp.EOF
            If Nvl(rsTemp!����ʱ��) = "" Then
                rsTemp.MoveNext
            Else
                GoTo AddCallingData
            End If
        Wend
        
AddCallingData:
        
        If Not rsTemp.EOF Then
            vfgCallingData.Cell(flexcpText, i, 0) = "�� " & rsTemp!���� & "��"
            vfgCallingData.Cell(flexcpText, i, 1) = rsTemp!��������
            
            If Trim(mstrDelString) <> "" Then
                strRoomName = rsTemp!�������� & rsTemp!���� & "����" & IIf(Nvl(rsTemp!�������, 0) = 0, "", "(��)")
                
                
                For j = LBound(aryDelStr()) To UBound(aryDelStr())
                    strRoomName = Replace(strRoomName, aryDelStr(j), "")
                Next j
                
                vfgCallingData.Cell(flexcpText, i, 2) = "��" & strRoomName
            Else
                vfgCallingData.Cell(flexcpText, i, 2) = "��" & rsTemp!�������� & rsTemp!���� & "����" & IIf(Nvl(rsTemp!�������, 0) = 0, "", "(��)")
            End If

            vfgCallingData.Cell(flexcpAlignment, i, 0, i, 2) = flexAlignLeftCenter
            
            If rsTemp!�Ŷ�״̬ = "������" Then
                vfgCallingData.Cell(flexcpForeColor, i, 0, i, 2) = mlngCallingColor
            Else
                vfgCallingData.Cell(flexcpForeColor, i, 0, i, 2) = mlngCalledColor
            End If
            
            Call vfgCallingData.AutoSize(0, vfgCallingData.Cols - 1)
            
            rsTemp.MoveNext
        Else
            vfgCallingData.Cell(flexcpText, i, 0) = ""
            vfgCallingData.Cell(flexcpText, i, 1) = ""
            vfgCallingData.Cell(flexcpText, i, 2) = ""

            vfgCallingData.Cell(flexcpAlignment, i, 0, i, 2) = flexAlignLeftCenter
        End If
    Next i
    
    vfgCallingData.Redraw = flexRDBuffered

    
    If Not rsTemp Is Nothing Then Set rsTemp = Nothing
    
    Exit Sub
errHandle:
    If Not rsTemp Is Nothing Then Set rsTemp = Nothing
    '�����ﲻ�ܽ��д�����ʾ��������ɳ�������������
    Call SaveErrLog
End Sub


Private Function GetTodayNum()
    On Error Resume Next
    
    Select Case Weekday(Date, vbMonday)
        Case 1: GetTodayNum = "һ"
        Case 2: GetTodayNum = "��"
        Case 3: GetTodayNum = "��"
        Case 4: GetTodayNum = "��"
        Case 5: GetTodayNum = "��"
        Case 6: GetTodayNum = "��"
        Case 7: GetTodayNum = "��"
    End Select
End Function


Private Sub picFace_Resize()
    labInf.Top = 40
    labInf.Left = 0
    labInf.Width = picFace.Width
    labInf.Height = picFace.Height - 40
End Sub

Public Sub timerLCD_Timer()
        On Error GoTo errHandle
        Dim blnTimer As Boolean
                
        
        blnTimer = timerLCD.Enabled
        timerLCD.Enabled = False
        
        labInf.Caption = "ף�����տ�����  " & Format(Now, "yyyy-mm-dd  hh:mm") & "  ����" & GetTodayNum
        
        Call MultiRoomsDisplay
        
        timerLCD.Enabled = blnTimer
    Exit Sub
errHandle:
    Call SaveErrLog
    
    timerLCD.Enabled = blnTimer
End Sub

Private Function setTextWidth(strText As String, iLen As Integer, intWay As Integer) As String
'**************************************************************************
'
'�����ı����ȴﵽ�ƶ����ȣ�������㣬�򲹳�ո�
'
'strText����Ҫ���õ��ı���
'
'iLen���ı�����
'
'intWay�����뷽��
'
'**************************************************************************
    
    On Error GoTo errHandle
    
    If Len(strText) >= iLen Then
        setTextWidth = Mid(strText, 1, iLen)
        Exit Function
    End If
    
    Select Case intWay
      Case 1
        setTextWidth = Space(iLen - Len(strText)) & strText
      Case 2
        setTextWidth = strText & Space(iLen - Len(strText))
      Case 3
        setTextWidth = Space((iLen - Len(strText)) - Int((iLen - Len(strText)) / 2)) & strText & Space(Int((iLen - Len(strText)) / 2))
    End Select
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
End Function

