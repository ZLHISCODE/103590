VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Begin VB.Form frmApplyCustom 
   Caption         =   "Form1"
   ClientHeight    =   9330
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11550
   Icon            =   "frmApplyCustom.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9330
   ScaleWidth      =   11550
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picTmp 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3120
      ScaleHeight     =   375
      ScaleWidth      =   855
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin SHDocVwCtl.WebBrowser webSub 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   11535
      ExtentX         =   20346
      ExtentY         =   14631
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   720
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmApplyCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintType As Integer '=0������=1�޸ģ�=2�鿴
Private mint���� As Integer '0 סԺҽ������վ��1 ����ҽ������վ��
Private mint���ó��� As Integer '���뵥���ó��ϣ�0��ҽ��վ���ã�1��ҽ���༭������á�Ϊ1ʱ�����û������ݼ��ؽ���ͱ���Ϊ�������ݡ�
Private mint�������� As Integer  '1-����,2-סԺ
Private mint������� As Integer '1-����,2-סԺ

Private mrsAppend As ADODB.Recordset
Private mobjFile As New FileSystemObject     '�ļ���������
Private mlng����ID As Long
Private mstr�Һŵ� As String
Private mlng��ҳID As Long
Private mlng����ID As Long
Private mlng���˿���id As Long '���˿���id/�Һ�ִ�п���id
Private mlng�������� As Long   '0-סԺ��1-����
Private mlng��������ID As Long
Private mstr�������� As String '�����������
Private mstr�������� As String, mstr����� As String, mstrסԺ�� As String, mstr���� As String
Private mrsPati As Recordset

Private mintPState As Integer
Private mdatTurn As Date
Private mstr��Ժʱ�� As String
Private mlng������� As Long  '������ţ��޸ģ��鿴ʱ���룬����Ϊ��
Private mlng��ĿID As Long
Private mlngҽ��ID As Long '�޸Ĳ鿴ʱ����
Private mlngXML�ļ�ID As Long 'XML��Ӧ���ļ�ID
Private mlngFileID As Long   '���뵥�ļ�ID
Private mintҽ��״̬ As Integer

Private mclsMipModule As zl9ComLib.clsMipModule '��Ϣƽ̨����

Private mlngǰ��ID As Long
Private mbytBaby As Byte  'Ӥ�����
Private mstrXSL As String, mstrXSLPath As String, mstrXSLFileName As String
Private mstrXML As String, mstrXMLPath As String, mstrXMLFileName As String
Private mstrHTML As String, mstrHTMLPath As String, mstrHTMLFileName As String
Private mstrFold As String
Private mblnOK As Boolean
Private mrsDefine As Recordset, mobjVBA As Object, mobjScript As clsScript
Private mint���� As Integer
Private mobjEmrInterface As Object
Private mobjXML As Object
Private mlngOutҽ��ID As Long
Private astr128Code()
Private astr128A() As Variant
Private astr128B() As Variant
Private astr128C() As Variant
Private astr128ID() As Variant
Public Enum ImageFileFormat
    Bmp = 1
    Jpg = 2
    Png = 3
    Gif = 4
End Enum
Private Const EncoderQuality             As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"

Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Enum EncoderParameterValueType
    EncoderParameterValueTypeByte = 1
    EncoderParameterValueTypeASCII = 2
    EncoderParameterValueTypeShort = 3
    EncoderParameterValueTypeLong = 4
    EncoderParameterValueTypeRational = 5
    EncoderParameterValueTypeLongRange = 6
    EncoderParameterValueTypeUndefined = 7
    EncoderParameterValueTypeRationalRange = 8
End Enum

Private Type EncoderParameter
    GUID(0 To 3)        As Long
    NumberOfValues      As Long
    Type                As EncoderParameterValueType
    value               As Long
End Type

Private Type EncoderParameters
    Count               As Long
    Parameter           As EncoderParameter
End Type

Private Type ImageCodecInfo
    ClassID(0 To 3)     As Long
    FormatID(0 To 3)    As Long
    CodecName           As Long
    DllName             As Long
    FormatDescription   As Long
    FilenameExtension   As Long
    MimeType            As Long
    Flags               As Long
    Version             As Long
    SigCount            As Long
    SigSize             As Long
    SigPattern          As Long
    SigMask             As Long
End Type

Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal hImage As Long, ByVal sFilename As Long, clsidEncoder As Any, encoderParams As Any) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hPal As Long, Bitmap As Long) As Long
Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (numEncoders As Long, Size As Long) As Long
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal Size As Long, Encoders As Any) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, pCLSID As Any) As Long
Private Declare Function GdipBitmapSetResolution Lib "gdiplus" (ByVal Bitmap As Long, ByVal xdpi As Single, ByVal ydpi As Single) As Long

'---------------------------------------------------------------------------------------------------
'ҽ���������(��xslt��Լ���Ĺ̶�����)

Private mrsAdvice As New Recordset
Private mintWebCompLete As Integer '������ɺ�Ĳ�����0-�޲�����7-Ԥ����6-��ӡ

Public Function ShowMe(frmParent As Object, ByVal int���� As Integer, ByVal intType As Integer, ByVal lng����ID As Long, ByVal str����ID As String, ByVal lng�������� As Long, _
    Optional ByVal lngFileID As Long, Optional ByRef lng������� As Long, Optional ByVal lng����id As Long, Optional ByVal lng��������ID As Long, _
    Optional ByVal lng����ID As Long, Optional ByVal rsDefine As Recordset, Optional ByVal intPState As Integer, Optional ByVal datTurn As Date, Optional ByVal int���ó��� As Integer, _
    Optional ByRef objMip As Object, Optional ByVal lngǰ��ID As Long, Optional ByVal bytBaby As Byte, Optional ByVal int���� As Integer, Optional ByRef lngOutҽ��ID As Long, Optional ByRef lng��Ŀid As Long) As Boolean
'���ܣ������ӿ�
'������frmParent �������壻int���� 0 סԺҽ������վ��1 ����ҽ������վ�� lng�������� 0-סԺ��1-���lng����ID��
'      str����ID ���� int���� �жϣ���ҳid/�Һŵ���
'      intType ��������   0-������1-�޸ģ�2-�鿴,3-ҽ���༭���ã�lngҽ��ID �����ҽ��ID��
'      lng����ID  ���˿���id/�Һ�ִ�п���id��int���ó��� 0������վ���棬1�� ҽ���༭���棻 lng��������ID ��������id��
'      strDefine ҽ�����ݸ�ʽ����
'      lng����ID��intPState��datTurn סԺ���У�objMip ��Ϣ����������Ϣ���� סԺ���У�
'      lng��ĿID-ҽ���༭����������ĿID
'      lngOutҽ��ID=���Σ������������޸ĵ�ҽ��ID�����ڶ�λ
    
    mint���� = int����
    Set mrsAdvice = Nothing
    If mint���� = 0 Then
        mlng��ҳID = Val(str����ID)
        mint�������� = 2
        mint������� = 2
    Else
        mstr�Һŵ� = str����ID
        mint�������� = 1
        mint������� = 1
    End If
    mint���ó��� = int���ó���
    mlng����ID = lng����ID
    mlng�������� = lng��������
    mlng���˿���id = lng����id
    mlng����ID = lng����ID
    mlng��������ID = lng��������ID
    mlngFileID = lngFileID
    mlng������� = lng�������
    mintPState = intPState
    mintType = intType
    mdatTurn = datTurn
    mint���� = int����
    mlng��ĿID = lng��Ŀid
    Set mrsDefine = rsDefine

    mlngǰ��ID = lngǰ��ID
    mbytBaby = bytBaby
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    
    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    ShowMe = mblnOK
    lngOutҽ��ID = mlngOutҽ��ID
End Function

Private Sub SetXMLForLoad(ByRef strText As String)
'���ܣ�XMLԤ����
    Dim strSQL As String, rsTmp As Recordset
    Dim objNodes As Object
    Dim objNode As Object
    Dim objNewNode As Object
    Dim objAttribute As Object
    Dim objAttributeNode As Object
    Dim lngִ�п���ID As Long

    On Error GoTo errH
    strSQL = "Select I.ID, I.������" & vbNewLine & _
        "From ����������Ŀ I, ������������ K" & vbNewLine & _
        "Where I.����id = K.ID And (K.���� = 1 And K.���� = '06' And" & vbNewLine & _
        "      I.������ Not In ('�������', 'һ��סԺ���', '����סԺ���', '�ϴ�סԺ���')  Or k.���� = 6)" & vbNewLine & _
        "Order By I.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    Call mobjXML.loadXML(strText)
    On Error Resume Next
    mobjXML.selectSingleNode(".//" & "����").Text = mstr��������
    mobjXML.selectSingleNode(".//" & "�Ա�").Text = mrsPati!�Ա� & ""
    mobjXML.selectSingleNode(".//" & "����").Text = mrsPati!���� & ""
    mobjXML.selectSingleNode(".//" & "����״��").Text = mrsPati!����״�� & ""
    mobjXML.selectSingleNode(".//" & "����").Text = mrsPati!��ǰ���� & ""
    mobjXML.selectSingleNode(".//" & "סԺ��").Text = mrsPati!סԺ�� & ""
    mobjXML.selectSingleNode(".//" & "�����").Text = mrsPati!����� & ""
    mobjXML.selectSingleNode(".//" & "����").Text = mstr��������
    mobjXML.selectSingleNode(".//" & "�Ʊ�").Text = mstr��������
    mobjXML.selectSingleNode(".//" & "����").Text = mrsPati!���� & ""
    mobjXML.selectSingleNode(".//" & "��ַ").Text = mrsPati!��ͥ��ַ & ""
    mobjXML.selectSingleNode(".//" & "������Դ").Text = IIF(mint�������� = 1, "����", "סԺ")
    mobjXML.selectSingleNode(".//" & "������Դ").Text = IIF(mint�������� = 1, "����", "סԺ")
    mobjXML.selectSingleNode(".//" & "����").Text = mrsPati!���� & ""
    mobjXML.selectSingleNode(".//" & "ְҵ").Text = mrsPati!ְҵ & ""
    mobjXML.selectSingleNode(".//" & "���֤��").Text = mrsPati!���֤�� & ""
    mobjXML.selectSingleNode(".//" & "���֤").Text = mrsPati!���֤�� & ""
    mobjXML.selectSingleNode(".//" & "��ϵ�˵绰").Text = mrsPati!��ϵ�˵绰 & ""
    mobjXML.selectSingleNode(".//" & "��ϵ�绰").Text = mrsPati!��ϵ�˵绰 & ""
    mobjXML.selectSingleNode(".//" & "��ͥ�绰").Text = mrsPati!��ͥ�绰 & ""
    If Val(mrsPati!��ǰ����ID & "") <> 0 Then
        mobjXML.selectSingleNode(".//" & "��ǰ����").Text = Sys.RowValue("���ű�", Val(mrsPati!��ǰ����ID & ""), "����")
        mobjXML.selectSingleNode(".//" & "����").Text = Sys.RowValue("���ű�", Val(mrsPati!��ǰ����ID & ""), "����")
    End If
    mobjXML.selectSingleNode(".//" & "�ͼ�����").Text = Format(zlDatabase.Currentdate, "YYYY��MM��DD��")
    mobjXML.selectSingleNode(".//" & "�ͼ�ҽʦ").Text = UserInfo.����
    mobjXML.selectSingleNode(".//" & "����ҽʦ").Text = UserInfo.����
    mobjXML.selectSingleNode(".//" & "��Чʱ��").Text = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm")
    mobjXML.selectSingleNode(".//" & "ִ��ʱ��").Text = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm")
    mobjXML.selectSingleNode(".//" & "�ɼ�ʱ��").Text = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm")
    
    
    '�������ݣ�֧�����븽������ݣ�Ҫ��XML�������������ڵ㣩
    Do While Not rsTmp.EOF
        err.Clear
        mobjXML.selectSingleNode(".//" & rsTmp!������).Text = ""
        '�������˵��û�����Ҫ�أ��Ͳ���ȡ
        If err.Number = 0 Then
            mobjXML.selectSingleNode(".//" & rsTmp!������).Text = GetAppendItemValue(rsTmp!������, rsTmp!ID, rsTmp!������)
        End If
        rsTmp.MoveNext
    Loop
    '��̬��������ĿID
    strSQL = "select C.ID,C.���,C.����,C.ִ�п���,D.���� from  ��������Ӧ�� B,������ĿĿ¼ C,������Ŀ���� D where  B.������ĿID=C.ID And C.ID=D.������ĿID AND D.����=1 and B.�����ļ�ID=[1] and b.Ӧ�ó���=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngFileID, mint��������)
    If rsTmp.RecordCount > 0 Then
        Set objNodes = mobjXML.selectSingleNode(".//" & "������ĿID")
        objNodes.Text = ""
        Do While Not rsTmp.EOF
            Set objNode = mobjXML.createNode("1", "��Ŀ", "")
            Set objNewNode = mobjXML.createNode("1", "ID", "")
            objNewNode.Text = rsTmp!ID & ""
            objNode.appendChild objNewNode
            Set objNewNode = mobjXML.createNode("1", "����", "")
            objNewNode.Text = rsTmp!���� & ""
            objNode.appendChild objNewNode
            Set objNewNode = mobjXML.createNode("1", "����", "")
            objNewNode.Text = rsTmp!���� & ""
            objNode.appendChild objNewNode
            lngִ�п���ID = Get����ִ�п���ID(mlng����ID, mlng��ҳID, rsTmp!��� & "", Val(rsTmp!ID & ""), 0, Val(rsTmp!ִ�п��� & ""), mlng���˿���id, mlng��������ID, 1, mint��������, , , mint��������)
            Set objNewNode = mobjXML.createNode("1", "ȱʡ����ID", "")
            objNewNode.Text = lngִ�п���ID
            objNode.appendChild objNewNode
            
            If mlng��ĿID <> 0 And mlng��ĿID = Val(rsTmp!ID & "") Then
                Set objAttributeNode = mobjXML.selectSingleNode("root/indt")
                Set objNewNode = mobjXML.createNode("1", "������ĿID", "")
                objNewNode.Text = rsTmp!ID & ""
                objAttributeNode.appendChild objNewNode
                
                Set objAttributeNode = mobjXML.selectSingleNode("root/indt")
                Set objNewNode = mobjXML.createNode("1", "������Ŀ����", "")
                objNewNode.Text = rsTmp!���� & ""
                objAttributeNode.appendChild objNewNode
                
            End If
            
            objNodes.appendChild objNode
            rsTmp.MoveNext
        Loop
    End If
    
    strText = mobjXML.xml
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetAppendItemValue(ByVal str��Ŀ As String, ByVal lngҪ��ID As Long, ByVal str������ As String) As String
'���ܣ���ȡָ�������븽��ֵ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strText As String
    Dim arrItem As Variant, i As Long
    Dim lng����ID As Long
    Dim intType As Integer '1-���2��סԺ
    
    On Error GoTo errH
    
    '4.δȡ����δ��ӦҪ�صģ��Ӳ���֮ǰ�ѱ����ҽ������ȡ,�������д��Ϊ׼
    strSQL = " Select ���� From (" & _
        " Select B.���� From ����ҽ����¼ A,����ҽ������ B" & _
        " Where A.ID=B.ҽ��ID And A.����ID=[1] And Nvl(A.Ӥ��,0)=[4]" & _
        IIF(mint�������� = 1, " And A.�Һŵ�=[2]", " And A.��ҳID=[3]") & _
        " And B.��Ŀ=[5] And B.���� is Not Null" & _
        " Order by A.����ʱ�� Desc) Where Rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�, mlng��ҳID, mbytBaby, str��Ŀ)
    If Not rsTmp.EOF Then strText = NVL(rsTmp!����)
    
    '1.����ж�ӦҪ�أ���Ҫ����ȡ������ȡ
    If lngҪ��ID <> 0 And strText = "" Then
        '���ϰ棬���°�
        If mint�������� = 1 Then '����
            strSQL = "Select Zl_Replace_Element_Value(B.������,[1],A.ID,1) as ����" & _
                " From ���˹Һż�¼ A,����������Ŀ B Where A.NO=[2] And B.ID=[3] And a.��¼����=1 And a.��¼״̬=1"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�, lngҪ��ID)
        Else
            strSQL = "Select Zl_Replace_Element_Value(������,[1],[2],2) as ���� From ����������Ŀ Where ID=[3]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, lngҪ��ID)
        End If
        If Not rsTmp.EOF Then strText = NVL(rsTmp!����)
        If strText = "" Then
            
            If mint�������� = 1 Then
                strSQL = "select a.id From ���˹Һż�¼ A Where A.NO=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�Һŵ�)
                lng����ID = Val(rsTmp!ID & "")
                intType = 1
            Else
                lng����ID = mlng��ҳID
                intType = 2
            End If
            strText = GetOrderInspectInfo(mlng����ID, str������, intType, lng����ID)
        End If
    End If
    
    GetAppendItemValue = strText
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetOrderInspectInfo(ByVal lng����ID As Long, ByVal strCondition As String, ByVal intType As Integer, ByVal lng����ID As Long) As String
'���ܣ���ȡָ�����˵�ָ������ڲ�����д����Ϣ�����磺���ߣ���ϵ�
    Dim strText As String
    On Error Resume Next
    If mobjEmrInterface Is Nothing Then
        Set mobjEmrInterface = CreateObject("zl9EmrInterface.ClsEmrInterface")
    End If
    If Not mobjEmrInterface Is Nothing Then
        strText = mobjEmrInterface.GetOrderInspectInfoEx(intType, lng����ID, lng����ID, strCondition)
        If err.Number <> 0 Then
            strText = mobjEmrInterface.GetOrderInspectInfo(lng����ID, strCondition)
        End If
    End If
    GetOrderInspectInfo = strText
End Function

Private Function LoadFile(ByVal lng�ļ�ID As Long, ByVal strFile As String, ByVal int���� As Integer, ByVal lngҽ��ID As Long) As Boolean
'���ܣ���ȡ�ʹ��������ļ�
    Dim strText As String
    Dim stmTmp As Stream
    Dim strFilename As String
    Dim objNodes As Object
    Dim objNewNode As Object
    
    strFilename = strFile
    strFile = mstrFold & "\" & strFile
    If mobjFile.FileExists(strFile) Then mobjFile.DeleteFile strFile, True
    
    If mintType = 0 Or lngҽ��ID = 0 And mintType = 2 Then
        strText = Sys.ReadLob(glngSys, 24, mlngFileID & "," & int����, strFile, 1)
    Else
        strText = Sys.ReadLob(glngSys, 25, lngҽ��ID & "," & int����, strFile, 1)
    End If
    'XMLԤ����
    If int���� = 2 Then
        '�¿�ʱ����Ԥ�����ȡ��Ϣ
        If mintType = 0 Then Call SetXMLForLoad(strText)
        On Error Resume Next
        If Not gobjPlugIn Is Nothing Then
            If gobjPlugIn.AdviceLoadApplyCustom(glngSys, IIF(mint���� = 0, pסԺҽ��վ, p����ҽ��վ), mlng����ID, IIF(mstr�Һŵ� = "", mlng��ҳID, mstr�Һŵ�), lng�ļ�ID, strText, lngҽ��ID) = False Then
                If err.Number = 0 Then Exit Function
            End If
            Call zlPlugInErrH(err, "AdviceLoadApplyCustom")
        End If
        'ҽ��վ�޸��ѷ��͵�����
        If mintҽ��״̬ <> 1 And mintType = 1 Then
            Set objNodes = Nothing
            Call mobjXML.loadXML(strText)
            Set objNodes = mobjXML.selectSingleNode("root/�ѷ���")
            If objNodes Is Nothing Then
                Set objNodes = mobjXML.selectSingleNode("root")
                Set objNewNode = mobjXML.createNode("1", "�ѷ���", "")
                objNewNode.Text = 1
                objNodes.appendChild objNewNode
            Else
                objNodes.Text = 1
            End If
            strText = mobjXML.xml
        End If
        If err.Number <> 0 Then err.Clear
        On Error GoTo 0
    End If
    If Not mobjFile.FolderExists(mstrFold) Then Call mobjFile.CreateFolder(mstrFold)
    If Not mobjFile.FileExists(strFile) Then Call mobjFile.CreateTextFile(strFile, True)
    Set stmTmp = New Stream
    stmTmp.Open
    stmTmp.Charset = "UTF-8"
    stmTmp.WriteText strText
    stmTmp.SaveToFile strFile, adSaveCreateOverWrite
    stmTmp.Close
    If int���� = 1 Then
        mstrXSL = strText
        mstrXSLPath = strFile
        mstrXSLFileName = strFilename
    ElseIf int���� = 2 Then
        mstrXML = strText
        mstrXMLPath = strFile
        mstrXMLFileName = strFilename
    ElseIf int���� = 3 Then
        mstrHTML = strText
        mstrHTMLPath = strFile
        mstrHTMLFileName = strFilename
    End If

    If Not mobjFile.FileExists(strFile) Then
        MsgBox "�ļ����ݶ�ȡʧ�ܣ�", vbInformation, gstrSysName:
        Screen.MousePointer = 0: Exit Function
    End If
    LoadFile = True
End Function

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_Edit_SaveExit  '����
            Control.Enabled = mintType <> 2
    End Select
End Sub

Private Sub Form_Load()
    Dim strSQL As String, rsTmp As Recordset
    Dim strFileXSL As String, strFileXML As String, strFileHTML As String
    
    InitCommandBar
    
    mblnOK = False
    mintWebCompLete = 0
    On Error GoTo errH
    Set mobjXML = CreateObject("MSXML2.DOMDocument")
    If mobjXML Is Nothing Then
        MsgBox "����MSXML2.DOMDocument����ʧ��", vbExclamation, Me.Caption
        Unload Me
        Exit Sub
    End If
    
    mintҽ��״̬ = 1
    
    If mlng������� = 0 Then
        strSQL = "Select A.�ļ�ID,A.�ļ���, A.���,0 as ҽ��ID,B.����,1 AS ҽ��״̬ From �Զ������뵥�ļ� A,�����ļ��б� B Where A.�ļ�ID=B.ID And B.ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngFileID)
    Else
        strSQL = "Select a.�ļ�ID,a.�ļ���, a.���,A.ҽ��ID,B.����,C.ҽ��״̬" & vbNewLine & _
                "From ҽ�����뵥�ļ� A,�����ļ��б� B,����ҽ����¼ C " & vbNewLine & _
                "Where A.�ļ�ID=B.ID And A.ҽ��ID=C.ID And C.������� = [1] And C.���id Is Null"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�������)
    End If
    
    Screen.MousePointer = 11
    If mlng����ID <> 0 Then
        strSQL = "select 0 ������ĿID,0 �շ�ϸĿID,'' ҽ������,0 ִ�п���ID ,0 �ɼ�����ID,sysdate ��Чʱ��,sysdate �ɼ�ʱ��,0 �ɼ���ʽID,'' �걾��λ from dual"
        Set mrsAdvice = zlDatabase.CopyNewRec(zlDatabase.OpenSQLRecord(strSQL, Me.Caption))
        strSQL = "Select ����,�Ա�,����,סԺ��,�����,��ǰ����,����״��,����,��ͥ��ַ,����,ְҵ,���֤��,��ϵ�˵绰,��ͥ�绰,��ǰ����ID from ������Ϣ Where ����ID=[1]"
        Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
        mstr�������� = mrsPati!����
        mstrסԺ�� = mrsPati!סԺ�� & ""
        mstr����� = mrsPati!����� & ""
        mstr���� = mrsPati!��ǰ���� & ""
        mstr�������� = Sys.RowValue("���ű�", mlng��������ID, "����")
    End If
    
    mstrFold = mobjFile.GetSpecialFolder(TemporaryFolder) & "\" & Decode(mintType, 0, "����", 1, "�޸�", 2, "�鿴") & "_" & mlngFileID & "_" & mlng�������
    If rsTmp.RecordCount > 0 Then
        mlngҽ��ID = Val(rsTmp!ҽ��ID)
        mlngXML�ļ�ID = Val(rsTmp!�ļ�ID & "")
        Me.Caption = rsTmp!���� & ""
        mintҽ��״̬ = Val(rsTmp!ҽ��״̬ & "")
    End If
    Do While Not rsTmp.EOF
        If LoadFile(Val(rsTmp!�ļ�ID & ""), rsTmp!�ļ��� & "", Val(rsTmp!���), Val(rsTmp!ҽ��ID)) = False Then Exit Sub
        rsTmp.MoveNext
    Loop
    
    '������ҳ
    webSub.Navigate mstrHTMLPath
    
    Me.Width = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & mlngFileID & "", "W", Me.Width)
    Me.Height = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & mlngFileID & "", "H", Me.Height)

    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    webSub.Top = 530
    webSub.Move 0, webSub.Top, Me.Width - 240, Me.Height - webSub.Top - 580
End Sub

Private Sub InitCommandBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objCbo As CommandBarComboBox
    
    '������----------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    '���ɹ�����
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����")
        objControl.IconId = 815
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ�˳�")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_SaveExit, " �����˳�(&S)")
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, " �˳�(&X)"): objControl.BeginGroup = True
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbsMain.KeyBindings
        .Add FALT, vbKeyX, conMenu_File_Exit
        .Add FALT, vbKeyS, conMenu_Edit_Save
    End With

End Sub

Private Sub PrintPreView(ByVal intType As Integer)
'���ܣ�Ԥ����ӡ
'������intType=6��ӡ��=7Ԥ����=8��ӡ����
    Dim stmTmp As Stream
    Dim b() As Byte
    Dim strTmp As String
    Dim picSet As StdPicture
    Dim objNode As Object
    
    If (intType = 7 Or intType = 6) Then
        If mintType <> 2 Then
            If SaveData = False Then Exit Sub
        End If
        
        mintWebCompLete = intType
        webSub.Navigate mstrHTMLPath
        If mintType = 0 Then mintType = 1
    Else
        Call webSub.ExecWB(intType, 1)
    End If

End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_File_PrintSet: PrintPreView 8
        Case conMenu_File_Preview: PrintPreView 7
        Case conMenu_File_Print: PrintPreView 6
        Case conMenu_Edit_SaveExit  '����
            If SaveData = True Then Unload Me
        Case conMenu_File_Exit
            Unload Me
    End Select
End Sub

Private Function SetXML(ByVal vTag As Object) As Boolean
    Dim objNodes As Object
    Dim objNode As Object
    Dim objNewNode As Object
    Dim objAttribute As Object
    Dim i As Long
    Dim j As Long
    Dim blnSelect As Boolean
    Dim blnIsSel As Boolean '�Ƿ����""ѡ��""�ڵ�
    Dim strNodeName As String
    Dim strGroupName As String
    Dim objNewNodeGroup As Object
    
    On Error Resume Next
    SetXML = True
    'ID��Ϊ�գ���������ZLREAD���ԵĲŽ���
    If vTag.ID = "" Or vTag.zlread <> "1" Then Exit Function
    If vTag.zlnotnull = "1" Then
        If err.Number <> 0 Then
            err.Clear
        Else
            If vTag.value = "" Then
                MsgBox "��Ŀ��[" & vTag.ID & "]Ϊ������Ŀ�����鲢��д��", vbExclamation, Me.Caption
                SetXML = False
                Exit Function
            End If
        End If
    End If
    
    If UCase(vTag.tagName) = "INPUT" And (vTag.Type = "text" Or vTag.Type = "password" Or vTag.Type = "hidden") Or UCase(vTag.tagName) = "TEXTAREA" And UCase(vTag.Type) = "TEXTAREA" Then
        '��һ�ı���¼��Ĵ���
        err.Clear
        mobjXML.selectSingleNode("root/indt/" & vTag.ID).Text = vTag.value
        If err.Number <> 0 Then
            '�����ڵ�
            strNodeName = "root/indt"
            Set objNodes = mobjXML.selectSingleNode(strNodeName)
            For i = 0 To UBound(Split(vTag.ID, "/"))
                Set objNodes = mobjXML.selectSingleNode(strNodeName & "/" & Split(vTag.ID, "/")(i))
                If objNodes Is Nothing Then
                    Set objNewNode = mobjXML.createNode("1", Split(vTag.ID, "/")(i), "")
                    If i = UBound(Split(vTag.ID, "/")) Then objNewNode.Text = vTag.value
                    Set objNodes = mobjXML.selectSingleNode(strNodeName)
                    objNodes.appendChild objNewNode
                End If
                strNodeName = strNodeName & "/" & Split(vTag.ID, "/")(i)
            Next
            err.Clear
        End If
    ElseIf UCase(vTag.tagName) = "SELECT" And vTag.Type = "select-one" Then
        '��ѡ��ѡ����
        Set objNodes = mobjXML.selectSingleNode("root/indt/" & vTag.ID)
        If Not objNodes Is Nothing And err.Number = 0 Then
            objNodes.Text = vTag.Options(vTag.selectedIndex).Text
            objNodes.Attributes.getNamedItem("valueid").value = vTag.value
        ElseIf vTag.ID <> "" Then
            Set objNodes = mobjXML.selectSingleNode("root/indt")
            Set objNewNode = mobjXML.createNode("1", vTag.ID, "")
            objNewNode.Text = vTag.Options(vTag.selectedIndex).Text
            Set objAttribute = mobjXML.createAttribute("valueid")
            objAttribute.value = vTag.value
            objNewNode.Attributes.setNamedItem objAttribute
            objNodes.appendChild objNewNode
        End If
    ElseIf UCase(vTag.tagName) = "INPUT" And (vTag.Type = "radio" Or vTag.Type = "checkbox") Then
        '��ѡ��Ͷ�ѡ��ѡ����
        strGroupName = ""
        strGroupName = vTag.groupname
        err.Clear
        mobjXML.selectSingleNode("root/indt/" & IIF(strGroupName = "", "", strGroupName & "/") & vTag.ID).Text = IIF(vTag.Checked, 1, 0)
        If err.Number <> 0 Then
            '�����ڵ�
            strNodeName = "root/indt"
            Set objNodes = mobjXML.selectSingleNode(strNodeName)
            For i = 0 To UBound(Split(vTag.ID, "/"))
                Set objNodes = mobjXML.selectSingleNode(strNodeName & "/" & IIF(strGroupName = "", "", strGroupName & "/") & Split(vTag.ID, "/")(i))
                If objNodes Is Nothing Then
                    Set objNewNode = mobjXML.createNode("1", Split(vTag.ID, "/")(i), "")
                    If i = UBound(Split(vTag.ID, "/")) Then objNewNode.Text = IIF(vTag.Checked, 1, 0)
                    If strGroupName <> "" Then
                        Set objNodes = mobjXML.selectSingleNode(strNodeName & "/" & strGroupName)
                        If objNodes Is Nothing Then
                            Set objNewNodeGroup = mobjXML.createNode("1", strGroupName, "")
                            Set objNodes = mobjXML.selectSingleNode(strNodeName)
                            objNodes.appendChild objNewNodeGroup
                            Set objNodes = mobjXML.selectSingleNode(strNodeName & "/" & strGroupName)
                        End If
                    Else
                        Set objNodes = mobjXML.selectSingleNode(strNodeName)
                    End If
                    
                    objNodes.appendChild objNewNode
                End If
                strNodeName = strNodeName & "/" & Split(vTag.ID, "/")(i)
            Next
            err.Clear
        End If
    End If
End Function

Private Function SaveData() As Boolean
    Dim vDoc, vTag
    Dim i As Long, j As Long
    Dim strSQL As String
    Dim rs��Ŀ As Recordset, rs�ɼ� As Recordset
    Dim datCur As Date, strDate As String
    Dim arrSQL As Variant
    Dim str����Ƽ����� As String, str�ɼ��Ƽ����� As String, str����ִ������ As String, str�ɼ�ִ������ As String
    Dim strҽ������ As String, lng��� As Long, strժҪ As String, lngҽ��ID As Long, lng���ID As Long
    Dim blnCancel As Boolean
    Dim str��ʼʱ�� As String, str�ɼ�ʱ�� As String
    Dim rsTmp As Recordset, blnTrans As Boolean
    Dim arrTmp() As String
    Dim stmTmp As Stream
    Dim b() As Byte
    Dim strTmp As String
    Dim picSet As StdPicture
    Dim objNode As Object
    
    On Error GoTo errH
    Set vDoc = webSub.Document
    '��֯XML
    Call mobjXML.loadXML(mstrXML)
    For i = 0 To vDoc.All.Length - 1
        If UCase(vDoc.All(i).tagName) = "INPUT" Or UCase(vDoc.All(i).tagName) = "SELECT" Or UCase(vDoc.All(i).tagName) = "TEXTAREA" Then
            Set vTag = vDoc.All(i)
            If vTag.Type = "text" Or vTag.Type = "password" Or vTag.Type = "hidden" Or vTag.Type = "select-one" Or UCase(vTag.Type) = "TEXTAREA" Then
            '----------------------------------------------------------------------------
            '��ȡxsltԼ������Ŀֵ(�Ժ�����������)
                If vTag.ID = "������ĿID" Then
                    mrsAdvice!������ĿID = Val(vTag.value)
                ElseIf vTag.ID = "�շ�ϸĿID" Then
                    mrsAdvice!�շ�ϸĿID = Val(vTag.value)
                ElseIf vTag.ID = "ҽ������" Then
                    mrsAdvice!ҽ������ = vTag.value
                ElseIf vTag.ID = "ִ�п���ID" Then
                    mrsAdvice!ִ�п���ID = Val(vTag.value)
                ElseIf vTag.ID = "�ɼ�����ID" Then
                    mrsAdvice!�ɼ�����ID = Val(vTag.value)
                ElseIf vTag.ID = "��Чʱ��" Then
                    If IsDate(vTag.value) Then mrsAdvice!��Чʱ�� = CDate(vTag.value)
                ElseIf vTag.ID = "�ɼ�ʱ��" Then
                    If IsDate(vTag.value) Then mrsAdvice!�ɼ�ʱ�� = CDate(vTag.value)
                ElseIf vTag.ID = "�ɼ���ʽID" Then
                    mrsAdvice!�ɼ���ʽID = Val(vTag.value)
                ElseIf vTag.ID = "�걾��λ" Then
                    mrsAdvice!�걾��λ = vTag.value
                End If
                If SetXML(vTag) = False Then Exit Function
            ElseIf vTag.Type = "submit" Then
                vTag.Click
            ElseIf vTag.Type = "radio" Or vTag.Type = "checkbox" Then
                If SetXML(vTag) = False Then Exit Function
            End If
        End If
    Next i
    mstrXML = mobjXML.xml
    webSub.Stop
    
    
    'ת���Ͷ�ȡ����
    If mrsAdvice!������ĿID = 0 Then
        MsgBox "δѡ��������Ŀ����ѡ��һ����Ŀ��", vbExclamation, Me.Caption
        Exit Function
    End If
    Set rs��Ŀ = Get������Ŀ��¼(mrsAdvice!������ĿID)
    
    datCur = zlDatabase.Currentdate
    If mrsAdvice!��Чʱ�� = CDate(0) Then mrsAdvice!��Чʱ�� = datCur
    If mrsAdvice!�ɼ�ʱ�� = CDate(0) Then mrsAdvice!�ɼ�ʱ�� = datCur
    If mrsAdvice!�걾��λ & "" = "" Then mrsAdvice!�걾��λ = rs��Ŀ!�걾��λ
    mrsAdvice.Update
    
    
    '���ݼ��
    If CheckData(rs��Ŀ) = False Then Exit Function
    
    '���ݱ���
    arrSQL = Array()
    If mlng������� <> 0 Then
        strSQL = "select a.id from ����ҽ����¼ a where a.�������=[1] and A.���ID is null"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�������)
        For i = 1 To rsTmp.RecordCount
            If Not (mintҽ��״̬ <> 1 And mintType = 1) Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Delete(" & rsTmp!ID & ",1)"
            End If
            lng���ID = Val(rsTmp!ID)
            rsTmp.MoveNext
        Next
    End If
    If Not (mintҽ��״̬ <> 1 And mintType = 1) Then
        str��ʼʱ�� = "To_Date('" & Format(mrsAdvice!��Чʱ��, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
        strDate = "To_Date('" & Format(datCur, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
        If mlng������� = 0 Then
            strSQL = "Select ����ҽ����¼_�������.Nextval as ������� From Dual"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            mlng������� = Val(rsTmp!�������)
        End If
        lng��� = GetMaxAdviceNO(mlng����ID, mlng��ҳID, mbytBaby, mstr�Һŵ�)
        If rs��Ŀ!��� = "C" Then
            str�ɼ�ʱ�� = "To_Date('" & Format(mrsAdvice!�ɼ�ʱ��, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
            str����Ƽ����� = Val("" & rs��Ŀ!�Ƽ�����)
            str����ִ������ = IIF("" & rs��Ŀ!ִ�п��� = "", "NULL", "" & rs��Ŀ!ִ�п���)
            strҽ������ = rs��Ŀ!����
            lng��� = lng��� + 1
            strժҪ = ""
            strժҪ = gclsInsure.GetItemInfo(mint����, mlng����ID, 0, "", 0, "", CStr(mrsAdvice!������ĿID) & "||" & IIF(mlng�������� = 1, 1, 2))
            blnCancel = CheckLISAppAdvice(mint��������, mlng����ID, mlng��ҳID, mint����, "C", mrsAdvice!������ĿID, mlng��������ID, UserInfo.����, mrsAdvice!ִ�п���ID, Val(rs��Ŀ!ִ�п��� & ""), strժҪ & "||0||0|| ||0")
            If Not blnCancel Then Exit Function
            
            lngҽ��ID = zlDatabase.GetNextID("����ҽ����¼")
            If lng���ID = 0 Then lng���ID = zlDatabase.GetNextID("����ҽ����¼")
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Insert(" & lngҽ��ID & "," & lng���ID & "," & lng��� & "," & mint�������� & "," & mlng����ID & "," & _
                ZVal(mlng��ҳID) & "," & mbytBaby & ",1,1,'C'," & mrsAdvice!������ĿID & ",Null,Null,Null,1," & _
                "'" & strҽ������ & "',Null," & "'" & mrsAdvice!�걾��λ & "','һ����',Null," & _
                "Null,Null,Null," & str����Ƽ����� & "," & mrsAdvice!ִ�п���ID & _
                "," & str����ִ������ & "," & 0 & "," & str��ʼʱ�� & ",Null," & mlng���˿���id & "," & _
                mlng��������ID & ",'" & UserInfo.���� & "'," & strDate & "," & IIF(mstr�Һŵ� = "", "NULL", "'" & mstr�Һŵ� & "'") & "," & ZVal(mlngǰ��ID) & "," & _
                "Null,0,Null," & IIF(strժҪ = "", "Null", "'" & strժҪ & "'") & ",'" & UserInfo.���� & "'" & _
                ",Null,Null,Null,Null," & mlng������� & ")"
                
            '�ɼ���ʽ
            Set rsTmp = Get������Ŀ��¼(mrsAdvice!�ɼ���ʽID)
            str�ɼ��Ƽ����� = Val("" & rsTmp!�Ƽ�����)
            str�ɼ�ִ������ = "" & rsTmp!ִ�п���
            strҽ������ = AdviceTextMake(rs��Ŀ)
            lng��� = lng��� + 1
            strժҪ = ""
            strժҪ = gclsInsure.GetItemInfo(mint����, mlng����ID, 0, "", 0, "", CStr(mrsAdvice!�ɼ���ʽID) & "||" & IIF(mlng�������� = 1, 1, 2))
            blnCancel = CheckLISAppAdvice(mint��������, mlng����ID, mlng��ҳID, mint����, "E", mrsAdvice!�ɼ���ʽID, mlng��������ID, UserInfo.����, mrsAdvice!�ɼ�����ID, Val(rsTmp!ִ�п��� & ""), strժҪ & "||0||0|| ||0")
            If Not blnCancel Then Exit Function
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Insert(" & lng���ID & ",Null," & lng��� & "," & mint�������� & "," & mlng����ID & "," & _
                ZVal(mlng��ҳID) & "," & mbytBaby & ",1,1,'E'," & mrsAdvice!�ɼ���ʽID & ",Null,Null,Null,1," & _
                "'" & strҽ������ & "','" & mrsAdvice!ҽ������ & "'," & "'" & mrsAdvice!�걾��λ & "','һ����',Null," & _
                "Null,Null,Null," & str�ɼ��Ƽ����� & "," & mrsAdvice!�ɼ�����ID & _
                "," & str�ɼ�ִ������ & "," & 0 & "," & str�ɼ�ʱ�� & ",Null," & mlng���˿���id & "," & _
                mlng��������ID & ",'" & UserInfo.���� & "'," & strDate & "," & IIF(mstr�Һŵ� = "", "NULL", "'" & mstr�Һŵ� & "'") & "," & ZVal(mlngǰ��ID) & "," & _
                "Null,0,Null," & IIF(strժҪ = "", "Null", "'" & strժҪ & "'") & ",'" & UserInfo.���� & "'" & _
                ",Null,Null,Null,Null," & mlng������� & ")"
        Else
            str����Ƽ����� = Val("" & rs��Ŀ!�Ƽ�����)
            str����ִ������ = IIF("" & rs��Ŀ!ִ�п��� = "", "NULL", "" & rs��Ŀ!ִ�п���)
            strҽ������ = AdviceTextMake(rs��Ŀ)
            lng��� = lng��� + 1
            strժҪ = ""
            strժҪ = gclsInsure.GetItemInfo(mint����, mlng����ID, 0, "", 0, "", CStr(mrsAdvice!������ĿID) & "||" & IIF(mlng�������� = 1, 1, 2))
            blnCancel = CheckLISAppAdvice(mint��������, mlng����ID, mlng��ҳID, mint����, "C", mrsAdvice!������ĿID, mlng��������ID, UserInfo.����, mrsAdvice!ִ�п���ID, Val(rs��Ŀ!ִ�п��� & ""), strժҪ & "||0||0|| ||0")
            If Not blnCancel Then Exit Function
            
            If lng���ID = 0 Then lng���ID = zlDatabase.GetNextID("����ҽ����¼")
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Insert(" & lng���ID & ",NULL," & lng��� & "," & mint�������� & "," & mlng����ID & "," & _
                ZVal(mlng��ҳID) & "," & mbytBaby & ",1,1,'" & rs��Ŀ!��� & "'," & mrsAdvice!������ĿID & ",Null,Null,Null,1," & _
                "'" & strҽ������ & "',Null," & "'" & mrsAdvice!�걾��λ & "','һ����',Null," & _
                "Null,Null,Null," & str����Ƽ����� & "," & mrsAdvice!ִ�п���ID & _
                "," & str����ִ������ & "," & 0 & "," & str��ʼʱ�� & ",Null," & mlng���˿���id & "," & _
                mlng��������ID & ",'" & UserInfo.���� & "'," & strDate & "," & IIF(mstr�Һŵ� = "", "NULL", "'" & mstr�Һŵ� & "'") & "," & ZVal(mlngǰ��ID) & "," & _
                "Null,0,Null," & IIF(strժҪ = "", "Null", "'" & strժҪ & "'") & ",'" & UserInfo.���� & "'" & _
                ",Null,Null,Null,Null," & mlng������� & ")"
        End If
        
        If mstr�Һŵ� <> "" Then
            '����Ĭ�ϰ������������
            strSQL = "Select A.ID from ������ϼ�¼ A,���˹Һż�¼ B Where A.����ID=B.����ID AND A.��ҳID=B.ID AND A.����ID=[1] and B.NO=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�)
            Do While Not rsTmp.EOF
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_�������ҽ��_Insert(NULL,NULL," & rsTmp!ID & ",'" & lng���ID & "')"
                rsTmp.MoveNext
            Loop
        End If
        
        If mintType = 0 Then
            Call mobjXML.loadXML(mstrXML)
            On Error Resume Next
            err.Clear
            Set objNode = mobjXML.selectSingleNode(".//" & "����")
            If Not objNode Is Nothing Then
                Set picSet = DrawBarCode128Auto(pictmp, "1" & lng���ID, 0, 2, True)
                'IE8��JPG�м������⣬���Ը�Ϊ��PNG
                SaveStdPicToFile picSet, mstrFold & "����.png", Png
                Open mstrFold & "����.png" For Binary As #1
                ReDim b(LOF(1) - 1)
                Get #1, , b
                Close #1
                strTmp = Base64Encode(b)
                objNode.Text = "data:image/jpeg;base64," & strTmp
                
                mstrXML = mobjXML.xml
                
            End If
        End If
    End If
    '���´洢�����أ��Ա��ڴ�ӡԤ��
    If Not mobjFile.FolderExists(mstrFold) Then Exit Function
    If mobjFile.FileExists(mstrXMLPath) Then Call mobjFile.DeleteFile(mstrXMLPath): Call mobjFile.CreateTextFile(mstrXMLPath, True)
    Set stmTmp = New Stream
    stmTmp.Open
    stmTmp.Charset = "UTF-8"
    stmTmp.WriteText mstrXML
    stmTmp.SaveToFile mstrXMLPath, adSaveCreateOverWrite
    stmTmp.Close
                
                
    '�����ļ�SQL
    On Error Resume Next
    If Not gobjPlugIn Is Nothing Then
        If gobjPlugIn.AdviceSaveApplyCustom(glngSys, IIF(mint���� = 0, pסԺҽ��վ, p����ҽ��վ), mlng����ID, IIF(mstr�Һŵ� = "", mlng��ҳID, mstr�Һŵ�), mlngXML�ļ�ID, mstrXML, webSub, mlngҽ��ID) = False Then
            If err.Number = 0 Then Exit Function
        End If
        Call zlPlugInErrH(err, "AdviceSaveApplyCustom")
    End If
    If err.Number <> 0 Then err.Clear
    On Error GoTo errH
    For j = 1 To 3
         ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_ҽ�����뵥�ļ�_Edit(" & mlngFileID & ",'" & Decode(j, 1, mstrXSLFileName, 2, mstrXMLFileName, 3, mstrHTMLFileName) & "'," & j & "," & lng���ID & ")"
        ReDim arrTmp(0)
        If Not Sys.GetLobSql(glngSys, 25, lng���ID & "," & j, Replace(Decode(j, 1, mstrXSL, 2, mstrXML, 3, mstrHTML), "'", "''"), arrTmp(), 1) Then
            MsgBox "�ļ���ӵ����ݿ�ʧ�ܣ��޷�����ҽ����", vbExclamation, Me.Caption
            Exit Function
        End If
        For i = LBound(arrTmp) To UBound(arrTmp)
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = arrTmp(i)
        Next
    Next

    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    If Not (mintҽ��״̬ <> 1 And mintType = 1) Then
        Call ZLHIS_CIS_001(mclsMipModule, mlng����ID, mstr��������, mstrסԺ��, , IIF(mlng�������� = 1, 1, 2), mlng��ҳID, mlng����ID, , mlng���˿���id, "", , mstr����, _
           lng���ID, 0, 1, rs��Ŀ!���, "", UserInfo.����, str��ʼʱ��, mlng��������ID, "", , , "")
    End If
    mlngOutҽ��ID = lng���ID
    mblnOK = True
    SaveData = True
    Exit Function
errH:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckData(rsSQL As Recordset) As Boolean
'���ܣ����ݼ��
    If rsSQL!��� & "" = "C" Then
        If mrsAdvice!�ɼ���ʽID = 0 Then
            MsgBox "������Ŀ����ѡ��һ���ɼ���ʽ����ѡ��һ���ɼ���ʽ��", vbExclamation, Me.Caption
            Exit Function
        End If
    End If
    '��������
    On Error Resume Next
    If Not gobjPlugIn Is Nothing Then
        If gobjPlugIn.AdviceCheckApplyCustom(glngSys, IIF(mint���� = 0, pסԺҽ��վ, p����ҽ��վ), mlng����ID, IIF(mstr�Һŵ� = "", mlng��ҳID, mstr�Һŵ�), mlngXML�ļ�ID, mstrXML, webSub, mlngҽ��ID) = False Then
            If err.Number = 0 Then Exit Function
        End If
        Call zlPlugInErrH(err, "AdviceCheckApplyCustom")
    End If
    If err.Number <> 0 Then err.Clear
    On Error GoTo 0
    CheckData = True
End Function

Private Sub Form_Unload(Cancel As Integer)
    If mobjFile.FolderExists(mstrFold) Then mobjFile.DeleteFolder mstrFold, True
    Set mobjFile = Nothing
    Set mobjVBA = Nothing
    Set mobjEmrInterface = Nothing
    '�����ļ�ID���洢�����С
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & mlngFileID & "", "W", Me.Width
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & mlngFileID & "", "H", Me.Height
    
End Sub

Private Function AdviceTextMake(rsSQL As Recordset) As String
'���ܣ���ȡҽ�������ı�
    Dim rsTmp As New ADODB.Recordset
    Dim blnDefine As Boolean
    Dim strText As String, strSQL As String
    Dim strField As String, intƵ�ʷ�Χ As Integer
    Dim i As Long, k As Long
    Dim blnDo As Boolean, str��� As String
    Dim str��ҩ As String, str�巨 As String, str��̬ As String
    Dim str���� As String, str���� As String
    Dim str���� As String, str�걾 As String
    Dim str��λ As String, str��λLast As String, str���� As String
    Dim dbl���� As Double, str��ҩ���� As String
    Dim str��ҩ������ĿIDS As String, strSame As String
    
    On Error GoTo errH
    If mobjVBA Is Nothing Then
        On Error Resume Next
        Set mobjVBA = CreateObject("ScriptControl")
        err.Clear: On Error GoTo 0
        
        If Not mobjVBA Is Nothing Then
            mobjVBA.Language = "VBScript"
            Set mobjScript = New clsScript
            mobjVBA.AddObject "clsScript", mobjScript, True
        End If
    End If
    
    'ȷ���Ƿ���
    str��� = rsSQL!���
    blnDefine = Not mrsDefine Is Nothing And Not mobjVBA Is Nothing
    If blnDefine Then
        mrsDefine.Filter = "�������='" & str��� & "'"
        If mrsDefine.EOF Then
            blnDefine = False
        ElseIf Trim(NVL(mrsDefine!ҽ������)) = "" Then
            blnDefine = False
        End If
    End If
    
ReDoDefault: '���ڰ����幫ʽ����ʧ�ܣ����°�ȱʡ���������֯
    strText = ""
    If blnDefine Then strText = mrsDefine!ҽ������
    
    '����ҽ������
    Select Case str���
    Case "C" '����-------------------------------------------------------------
        str���� = "": str�걾 = ""

        str���� = rsSQL!����
        str�걾 = mrsAdvice!�걾��λ & ""
        If str�걾 = "" Then str�걾 = rsSQL!�걾��λ & ""
        
        If Not blnDefine Then
            strText = str���� & IIF(str�걾 <> "", "(" & str�걾 & ")", "")
        Else
            If InStr(strText, "[������Ŀ]") > 0 Then
                strField = str����
                strText = Replace(strText, "[������Ŀ]", """" & strField & """")
            End If
            If InStr(strText, "[����걾]") > 0 Then
                strField = str�걾
                strText = Replace(strText, "[����걾]", """" & strField & """")
            End If
            If InStr(strText, "[�ɼ�����]") > 0 Then
                If mrsAdvice!�ɼ���ʽID <> 0 Then
                    strSQL = "select ���,����,�걾��λ from ������ĿĿ¼ where ID=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mrsAdvice!�ɼ���ʽID)
                    strField = rsTmp!���� & ""
                Else
                    strField = ""
                End If
                
                strText = Replace(strText, "[�ɼ�����]", """" & strField & """")
            End If
        End If
    Case "D" '���-------------------------------------------------------------
        str��λ = "": str���� = ""
        
        strText = rsSQL!����
    Case "F" '����-------------------------------------------------------------
        str���� = "": str���� = ""
        strText = rsSQL!����

    Case "8" '��ҩ�䷽---------------------------------------------------------
        str��ҩ = "": str�巨 = "": str��ҩ������ĿIDS = "": strSame = ""
        strText = rsSQL!����
    Case "4" '����------------------------------------------------------------
        If Val(mrsAdvice!�շ�ϸĿID & "") <> 0 Then
            strSQL = "Select ����,���,���� From �շ���ĿĿ¼ Where ID=[1]"
            Set rsTmp = New ADODB.Recordset '���Filter
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mrsAdvice!�շ�ϸĿID & ""))
            
            If Not blnDefine Then
                strText = rsSQL!����
                If Not IsNull(rsTmp!���) Then
                    strText = strText & " " & rsTmp!���
                End If
            Else
                If InStr(strText, "[��������]") > 0 Then
                    strField = rsTmp!����
                    strText = Replace(strText, "[��������]", """" & strField & """")
                End If
                If InStr(strText, "[���]") > 0 Then
                    strField = NVL(rsTmp!���)
                    strText = Replace(strText, "[���]", """" & strField & """")
                End If
                If InStr(strText, "[����]") > 0 Then
                    strField = NVL(rsTmp!����)
                    strText = Replace(strText, "[����]", """" & strField & """")
                End If
            End If
        End If
    Case "5", "6" '����ҩ���г�ҩ---------------------------------------------
        strText = rsSQL!����
    Case "K" '��Ѫҽ��
        strText = rsSQL!����
    Case Else '�����������-----------------------------------------------------
        If Not blnDefine Then
            strText = rsSQL!����
        Else
            If InStr(strText, "[������Ŀ]") > 0 Then
                strField = rsSQL!����
                strText = Replace(strText, "[������Ŀ]", """" & strField & """")
            End If
        End If
        '����ҽ��������ʾ
        If str��� = "Z" And (Val(rsSQL!�������� & "") = 4 Or Val(rsSQL!�������� & "") = 14) Then
            strText = "������" & strText & "������"
        End If
        'ת��ҽ��������ʾ
        If str��� = "Z" And Val(rsSQL!�������� & "") = 3 Then
            strText = "������" & strText & "������"
        End If
    End Select
    
    '�����ֶλ���Թ���������ֶ�-------------------------------------------
    If blnDefine Then
        If InStr(strText, "[��ʼʱ��]") > 0 Then
            strField = mrsAdvice!��Чʱ��
            strText = Replace(strText, "[��ʼʱ��]", """" & strField & """")
        End If
        If InStr(strText, "[ҽ������]") > 0 Then
            strField = mrsAdvice!ҽ������ & ""
            If mrsAdvice!ҽ������ & "" <> "" Then
                If strField <> "" Then
                    strField = strField & "," & mrsAdvice!ҽ������
                Else
                    strField = mrsAdvice!ҽ������
                End If
            End If
            strText = Replace(strText, "[ҽ������]", """" & strField & """")
        End If
    End If
            
    '����ҽ������
    If blnDefine Then
        On Error Resume Next
        strText = mobjVBA.Eval(strText)
        If mobjVBA.Error.Number <> 0 Then
            err.Clear: On Error GoTo errH
            blnDefine = False: GoTo ReDoDefault
        End If
    End If
    AdviceTextMake = strText
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub webSub_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    If mintWebCompLete <> 0 Then
        Call webSub.ExecWB(mintWebCompLete, 2)
        If mintWebCompLete = 6 Then Unload Me
        mintWebCompLete = 0
    End If
End Sub

Private Sub WebSub_DownloadBegin()
    webSub.Silent = True
End Sub

Private Sub WebSub_DownloadComplete()
    webSub.Silent = True
End Sub


Public Function DrawBarCode128Auto(ByVal PicObj As Object, ByVal strBarCode As String, sngPrintWidth As Single, _
                            Optional ByVal intLineWidth As Integer = 2, Optional ByVal blnShowBarCodeTxt As Boolean = True) As StdPicture
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����                       ����ͼƬ����ͼ��
    '����                       PicObj              Picture����(���ڻ�ͼ��
    '                           strBarCode          �������������
    '                           sngPrintWidth       ��ӡʱʹ�õĹ̶����ȵ�λ��mm)(���أ�
    '                   ��ѡ
    '                           intLineWidth        �߿����ڿ�������Ŀ�� Ĭ��Ϊ2
    '                           blnShowBarCodeTxt   �Ƿ���ʾ�������ݣ�Ĭ��TrueΪ��ʾ
    '����                       Image����
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim intLoop As Integer
    Dim strGetBarCode As String
    Dim intval As Integer
    Dim lngTxtHeight As Long
    
    If intLineWidth = 0 Then intLineWidth = 2
    
    With PicObj
        .Cls
        .BackColor = vbWhite
        .AutoRedraw = True
        .ScaleMode = vbPixels
        .DrawWidth = intLineWidth
        .Height = 1335
    End With
    
    strGetBarCode = GetBarCode(strBarCode)
    PicObj.Width = (2 + Len(strGetBarCode) * intLineWidth) * Screen.TwipsPerPixelX
    
    If blnShowBarCodeTxt Then
        PicObj.FontSize = 12
        PicObj.FontName = "Arial"
        lngTxtHeight = PicObj.TextHeight(strBarCode)
        
        PicObj.CurrentX = (PicObj.ScaleWidth / 2) - PicObj.TextWidth(strBarCode) / 2   ' ˮƽ����
        PicObj.CurrentY = PicObj.ScaleHeight - lngTxtHeight - 1 ' ��ֱ����
        PicObj.Print strBarCode
    End If
    
    For intLoop = 1 To Len(strGetBarCode)
        intval = Mid$(strGetBarCode, intLoop, 1)
        If blnShowBarCodeTxt Then
            PicObj.Line (1 + (intLoop * intLineWidth), 1)-(1 + (intLoop * intLineWidth), PicObj.ScaleHeight - lngTxtHeight - 2), IIF(intval = 1, vbBlack, vbWhite), BF
        Else
            PicObj.Line (1 + (intLoop * intLineWidth), 1)-(1 + (intLoop * intLineWidth), PicObj.ScaleHeight - 1), IIF(intval = 1, vbBlack, vbWhite), BF
        End If
    Next
    
    PicObj.Width = (2 + intLoop * intLineWidth) * Screen.TwipsPerPixelX
    
    sngPrintWidth = PicObj.ScaleWidth / 8
    
    Set DrawBarCode128Auto = PicObj.Image
    
    '�ָ�ȱʡ�Ա���������ͼ��Ӱ��
    PicObj.ScaleMode = vbTwips
    PicObj.DrawWidth = 1
    PicObj.FontName = "����"
    PicObj.FontSize = 9
End Function

Private Sub initBarCode()
    '//539
    '��ʼ���൱����
    astr128Code = Array( _
             "11011001100", "11001101100", "11001100110", "10010011000", "10010001100", "10001001100", _
             "10011001000", "10011000100", "10001100100", "11001001000", "11001000100", "11000100100", _
             "10110011100", "10011011100", "10011001110", "10111001100", "10011101100", "10011100110", _
             "11001110010", "11001011100", "11001001110", "11011100100", "11001110100", "11101101110", _
             "11101001100", "11100101100", "11100100110", "11101100100", "11100110100", "11100110010", _
             "11011011000", "11011000110", "11000110110", "10100011000", "10001011000", "10001000110", _
             "10110001000", "10001101000", "10001100010", "11010001000", "11000101000", "11000100010", _
             "10110111000", "10110001110", "10001101110", "10111011000", "10111000110", "10001110110", _
             "11101110110", "11010001110", "11000101110", "11011101000", "11011100010", "11011101110", _
             "11101011000", "11101000110", "11100010110", "11101101000", "11101100010", "11100011010", _
             "11101111010", "11001000010", "11110001010", "10100110000", "10100001100", "10010110000", _
             "10010000110", "10000101100", "10000100110", "10110010000", "10110000100", "10011010000", _
             "10011000010", "10000110100", "10000110010", "11000010010", "11001010000", "11110111010", _
             "11000010100", "10001111010", "10100111100", "10010111100", "10010011110", "10111100100", _
             "10011110100", "10011110010", "11110100100", "11110010100", "11110010010", "11011011110", _
             "11011110110", "11110110110", "10101111000", "10100011110", "10001011110", "10111101000", _
             "10111100010", "11110101000", "11110100010", "10111011110", "10111101110", "11101011110", _
             "11110101110", "11010000100", "11010010000", "11010011100", "1100011101011" _
             )
             
    astr128A = Array( _
             "SP", "!", """", "#", "$", "%", _
             "&", "'", "(", ")", "*", "+", _
             ",", "-", ".", "/", "0", "1", _
             "2", "3", "4", "5", "6", "7", _
             "8", "9", ":", ";", "<", "=", _
             ">", "?", "@", "A", "B", "C", _
             "D", "E", "F", "G", "H", "I", _
             "J", "K", "L", "M", "N", "O", _
             "P", "Q", "R", "S", "T", "U", _
             "V", "W", "X", "Y", "Z", "[", _
             "\", "]", "^", "_", "NUL", "SOH", _
             "STX", "ETX", "EOT", "ENQ", "ACK", "BEL", _
             "BS", "HT", "LF", "VT", "FF", "CR", _
             "SO", "SI", "DLE", "DC1", "DC2", "DC3", _
             "DC4", "NAK", "SYN", "ETB", "CAN", "EM", _
             "SUB", "ESC", "FS", "GS", "RS", "US", _
             "FNC3", "FNC2", "SHIFT", "CODEC", "CODEB", "FNC4", _
             "FNC1", "StartA", "StartB", "StartC", "Stop" _
             )
    
    astr128B = Array( _
             "SP", "!", """", "#", "$", "%", _
             "&", "'", "(", ")", "*", "+", _
             ",", "-", ".", "/", "0", "1", _
             "2", "3", "4", "5", "6", "7", _
             "8", "9", ":", ";", "<", "=", _
             ">", "?", "@", "A", "B", "C", _
             "D", "E", "F", "G", "H", "I", _
             "J", "K", "L", "M", "N", "O", _
             "P", "Q", "R", "S", "T", "U", _
             "V", "W", "X", "Y", "Z", "[", _
             "\", "]", "^", "_", "`", "a", _
             "b", "c", "d", "e", "f", "g", _
             "h", "i", "j", "k", "I", "m", _
             "n", "o", "p", "q", "r", "s", _
             "t", "u", "v", "w", "x", "y", _
             "z", "{", "|", "}", "~", "DEL", _
             "FNC3", "FNC2", "SHIFT", "CODEC", "FNC4", "CODEA", _
             "FNC1", "StartA", "StartB", "StartC", "Stop" _
             )
             
    astr128C = Array( _
             "0", "1", "2", "3", "4", "5", _
             "6", "7", "8", "9", "10", "11", _
             "12", "13", "14", "15", "16", "17", _
             "18", "19", "20", "21", "22", "23", _
             "24", "25", "26", "27", "28", "29", _
             "30", "31", "32", "33", "34", "35", _
             "36", "37", "38", "39", "40", "41", _
             "42", "43", "44", "45", "46", "47", _
             "48", "49", "50", "51", "52", "53", _
             "54", "55", "56", "57", "58", "59", _
             "60", "61", "62", "63", "64", "65", _
             "66", "67", "68", "69", "70", "71", _
             "72", "73", "74", "75", "76", "77", _
             "78", "79", "80", "81", "82", "83", _
             "84", "85", "86", "87", "88", "89", _
             "90", "91", "92", "93", "94", "95", _
             "96", "97", "98", "99", "CODEB", "CODEA", _
             "FNC1", "StartA", "StartB", "StartC", "Stop" _
             )
    astr128ID = Array( _
             "0", "1", "2", "3", "4", "5", _
             "6", "7", "8", "9", "10", "11", _
             "12", "13", "14", "15", "16", "17", _
             "18", "19", "20", "21", "22", "23", _
             "24", "25", "26", "27", "28", "29", _
             "30", "31", "32", "33", "34", "35", _
             "36", "37", "38", "39", "40", "41", _
             "42", "43", "44", "45", "46", "47", _
             "48", "49", "50", "51", "52", "53", _
             "54", "55", "56", "57", "58", "59", _
             "60", "61", "62", "63", "64", "65", _
             "66", "67", "68", "69", "70", "71", _
             "72", "73", "74", "75", "76", "77", _
             "78", "79", "80", "81", "82", "83", _
             "84", "85", "86", "87", "88", "89", _
             "90", "91", "92", "93", "94", "95", _
             "96", "97", "98", "99", "100", "101", _
             "102", "103", "104", "105", "106" _
             )
End Sub

Private Function FindArray(strChar As String, strArray() As Variant) As Integer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����           �����ַ��������е�λ��
    '����           strChar �����ַ�
    '               strArray() Ҫ���ҵ�����
    '����           λ��index
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim intLoop As Integer
    FindArray = 0
    For intLoop = 0 To UBound(strArray)
        If strChar = strArray(intLoop) Then
            FindArray = intLoop
            Exit For
        End If
    Next
    
End Function


Private Function FindCode(intType As Integer, strChar As String) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����:              ��������ҹ̶����ַ�
    '����:              intType(1=128A,2=128B,3=128C,4=ID)
    '                   strChar �����ַ�
    '����:              ��Ӧ�ı������
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim intIndex As Integer
    
    Call initBarCode
    
    If intType = 1 Then
        intIndex = FindArray(strChar, astr128A)
        FindCode = astr128Code(intIndex)
    End If
    
    If intType = 3 Then
        intIndex = FindArray(strChar, astr128C)
        FindCode = astr128Code(intIndex)
    End If
    
    If intType = 4 Then
        intIndex = FindArray(strChar, astr128ID)
        FindCode = astr128Code(intIndex)
    End If
End Function

Private Function GetBarCode(strChar As String) As String
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����           �������ַ����������߹���
    '����           strChar = �����ַ���
    '����           ��������Ĺ���
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim intLoop As Integer
    Dim strTmp As String
    Dim intType As Integer
    Dim lngCheckCount As Long
    Dim intRow As Integer
    intType = 0
    intRow = 0
    
    For intLoop = 1 To Len(strChar) Step 2
        strTmp = Mid$(strChar, intLoop, 2)
        If Len(strTmp) = 2 And IsNumeric(strTmp) = True Then
            '128C��
            If intType = 0 Then
                GetBarCode = FindCode(3, "StartC")
                lngCheckCount = FindID(3, "StartC")
                intRow = intRow + 1
                intType = 3
            End If
            If intType <> 3 Then
                'תΪ128C��
                GetBarCode = GetBarCode & FindCode(1, "CODEC")
                lngCheckCount = lngCheckCount + intRow * FindID(1, "CODEC")
                intRow = intRow + 1
            End If
            
            
            GetBarCode = GetBarCode & FindCode(3, Val(strTmp))
            lngCheckCount = lngCheckCount + intRow * FindID(3, Val(strTmp))
            intRow = intRow + 1
        
            intType = 3
        Else
             '128A��
            If intType = 0 Then
                GetBarCode = FindCode(1, "StartA")
                lngCheckCount = FindID(1, "StartA")
                intRow = intRow + 1
                intType = 1
            End If
            If intType <> 1 Then
                'תΪ128A��
                GetBarCode = GetBarCode & FindCode(3, "CODEA")
                lngCheckCount = lngCheckCount + intRow * FindID(3, "CODEA")
                intRow = intRow + 1
            End If
            If Len(strTmp) = 1 Then
                GetBarCode = GetBarCode & FindCode(1, strTmp)
                lngCheckCount = lngCheckCount + intRow * FindID(1, strTmp)
                intRow = intRow + 1
            Else
                GetBarCode = GetBarCode & FindCode(1, Mid(strTmp, 1, 1))
                lngCheckCount = lngCheckCount + intRow * FindID(1, Mid(strTmp, 1, 1))
                intRow = intRow + 1
              
                GetBarCode = GetBarCode & FindCode(1, Mid(strTmp, 2, 1))
                lngCheckCount = lngCheckCount + intRow * FindID(1, Mid(strTmp, 2, 1))
                intRow = intRow + 1
            End If
            intType = 1
        End If
    Next
    lngCheckCount = lngCheckCount Mod 103
    GetBarCode = GetBarCode & FindCode(4, CStr(lngCheckCount)) & FindCode(3, "Stop")
End Function

Private Function FindID(intType As Integer, strChar As String) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����:              ��������ҹ̶����ַ���ID
    '����:              intType(1=128A,2=128B,3=128C,4=ID)
    '                   strChar �����ַ�
    '����:              ��Ӧ�ı����ID
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim intIndex As Integer
    
    Call initBarCode
    
    If intType = 1 Then
        intIndex = FindArray(strChar, astr128A)
        FindID = Val(astr128ID(intIndex))
    End If
    
    If intType = 3 Then
        intIndex = FindArray(strChar, astr128C)
        FindID = Val(astr128ID(intIndex))
    End If
End Function


Function Base64Encode(Str() As Byte) As String                                  'Base64 ����
    On Error GoTo over                                                          '�Ŵ�
    Dim buf() As Byte, Length As Long, mods As Long
    Const B64_CHAR_DICT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
    mods = (UBound(Str) + 1) Mod 3   '����3������
    Length = UBound(Str) + 1 - mods
    ReDim buf(Length / 3 * 4 + IIF(mods <> 0, 4, 0) - 1)
    Dim i As Long
    For i = 0 To Length - 1 Step 3
        buf(i / 3 * 4) = (Str(i) And &HFC) / &H4
        buf(i / 3 * 4 + 1) = (Str(i) And &H3) * &H10 + (Str(i + 1) And &HF0) / &H10
        buf(i / 3 * 4 + 2) = (Str(i + 1) And &HF) * &H4 + (Str(i + 2) And &HC0) / &H40
        buf(i / 3 * 4 + 3) = Str(i + 2) And &H3F
    Next
    If mods = 1 Then
        buf(Length / 3 * 4) = (Str(Length) And &HFC) / &H4
        buf(Length / 3 * 4 + 1) = (Str(Length) And &H3) * &H10
        buf(Length / 3 * 4 + 2) = 64
        buf(Length / 3 * 4 + 3) = 64
    ElseIf mods = 2 Then
        buf(Length / 3 * 4) = (Str(Length) And &HFC) / &H4
        buf(Length / 3 * 4 + 1) = (Str(Length) And &H3) * &H10 + (Str(Length + 1) And &HF0) / &H10
        buf(Length / 3 * 4 + 2) = (Str(Length + 1) And &HF) * &H4
        buf(Length / 3 * 4 + 3) = 64
    End If
    For i = 0 To UBound(buf)
        Base64Encode = Base64Encode + Mid(B64_CHAR_DICT, buf(i) + 1, 1)
    Next
over:
End Function

Public Function SaveStdPicToFile(Stdpic As StdPicture, ByVal FileName As String, _
                              Optional ByVal FileFormat As ImageFileFormat = Jpg, _
                              Optional ByVal JpgQuality As Long = 80, _
                              Optional Resolution As Single) As Boolean
                              
    Dim CLSID(3)        As Long
    Dim Bitmap          As Long
    Dim Token           As Long
    Dim Gsp             As GdiplusStartupInput

    Gsp.GdiplusVersion = 1                      'GDI+ 1.0�汾
    GdiplusStartup Token, Gsp                   '��ʼ��GDI+
    GdipCreateBitmapFromHBITMAP Stdpic.Handle, Stdpic.hPal, Bitmap
    If Bitmap <> 0 Then                          '˵�����ǳɹ��Ľ�StdPic����ת��ΪGDI+��Bitmap������
        GdipBitmapSetResolution Bitmap, Resolution, Resolution
        Select Case FileFormat
        Case ImageFileFormat.Bmp
            If Not GetEncoderClsID("Image/bmp", CLSID) = -1 Then
                SaveStdPicToFile = (GdipSaveImageToFile(Bitmap, StrPtr(FileName), CLSID(0), ByVal 0) = 0)
            End If
        Case ImageFileFormat.Jpg                    'JPG��ʽ�������ñ��������
            Dim aEncParams()        As Byte
            Dim uEncParams          As EncoderParameters
            If GetEncoderClsID("Image/jpeg", CLSID) <> -1 Then
                uEncParams.Count = 1                                        ' �����Զ���ı������������Ϊ1������
                If JpgQuality < 0 Then
                    JpgQuality = 0
                ElseIf JpgQuality > 100 Then
                    JpgQuality = 100
                End If
                ReDim aEncParams(1 To Len(uEncParams))
                With uEncParams.Parameter
                    .NumberOfValues = 1
                    .Type = EncoderParameterValueTypeLong                   ' ���ò���ֵ����������Ϊ������
                    Call CLSIDFromString(StrPtr(EncoderQuality), .GUID(0))  ' ���ò���Ψһ��־��GUID������Ϊ����Ʒ��
                    .value = VarPtr(JpgQuality)                                ' ���ò�����ֵ��Ʒ�ʵȼ������Ϊ100��ͼ���ļ���С��Ʒ�ʳ�����
                End With
                CopyMemory aEncParams(1), uEncParams, Len(uEncParams)
                SaveStdPicToFile = (GdipSaveImageToFile(Bitmap, StrPtr(FileName), CLSID(0), aEncParams(1)) = 0)
            End If
        Case ImageFileFormat.Png
            If Not GetEncoderClsID("Image/png", CLSID) = -1 Then
                SaveStdPicToFile = (GdipSaveImageToFile(Bitmap, StrPtr(FileName), CLSID(0), ByVal 0) = 0)
            End If
        Case ImageFileFormat.Gif
            If Not GetEncoderClsID("Image/gif", CLSID) = -1 Then                '���ԭʼ��ͼ����24λ����������������ϵͳ�ĵ�ɫ������ͼ��ת��Ϊ8λ��ת����Ч���᲻������,��Ҳ�п���ϵͳ���Զ�ת��������ʧ��
                SaveStdPicToFile = (GdipSaveImageToFile(Bitmap, StrPtr(FileName), CLSID(0), ByVal 0) = 0)
            End If
        End Select
    End If
    GdipDisposeImage Bitmap      'ע���ͷ���Դ
    GdiplusShutdown Token       '�ر�GDI+��
End Function

Private Function GetEncoderClsID(strMimeType As String, ClassID() As Long) As Long
    Dim Num         As Long
    Dim Size        As Long
    Dim i           As Long
    Dim Info()      As ImageCodecInfo
    Dim Buffer()    As Byte
    GetEncoderClsID = -1
    GdipGetImageEncodersSize Num, Size               '�õ�����������Ĵ�С
    If Size <> 0 Then
       ReDim Info(1 To Num) As ImageCodecInfo       '�����鶯̬�����ڴ�
       ReDim Buffer(1 To Size) As Byte
       GdipGetImageEncoders Num, Size, Buffer(1)            '�õ�������ַ�����
       CopyMemory Info(1), Buffer(1), (Len(Info(1)) * Num)     '������ͷ
       For i = 1 To Num             'ѭ��������н���
           If (StrComp(PtrToStrW(Info(i).MimeType), strMimeType, vbTextCompare) = 0) Then         '�����ָ��ת���ɿ��õ��ַ�
               CopyMemory ClassID(0), Info(i).ClassID(0), 16  '�������ID
               GetEncoderClsID = i      '���سɹ�������ֵ
               Exit For
           End If
       Next
    End If
End Function

Private Function PtrToStrW(ByVal lpsz As Long) As String
    Dim Out         As String
    Dim Length      As Long
    Length = lstrlenW(lpsz)
    If Length > 0 Then
        Out = StrConv(String$(Length, vbNullChar), vbUnicode)
        CopyMemory ByVal Out, ByVal lpsz, Length * 2
        PtrToStrW = StrConv(Out, vbFromUnicode)
    End If
End Function

