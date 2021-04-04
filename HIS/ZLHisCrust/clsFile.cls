VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'===========================================================================
'                0��ģ������빫������
'===========================================================================
'�ļ�����
Public Enum FileType
    FT_Public = 0                   '��Ʒ��������
    FT_Apply = 1                    '��ƷӦ�ò���
    FT_Help = 2                     '��Ʒ�����ļ�
    FT_AdditionFile = 3             '��Ʒ�����ļ�
    FT_Other = 4                    '������Ʒ�ļ�
    FT_System = 5                   'ϵͳ�ļ�
End Enum

'===========================================================================
'                1���������Եı���
'===========================================================================
'FileName�ļ�����(R/W)
Private mstrFileName                As String
'FileSTName�ļ���׼���ƣ�ȫ��д(R)
Private mstrFileSTName              As String
'FileType�ļ�����(R/W)
Private mftFileType                 As FileType
'FilePath�ļ�ȫ·��(R/W)
Private mstrFilePath                As String
'SetupPath��װ·��(R/W)
Private mstrSetupPath               As String
'FileVerion�ļ��汾(R/W)
Private mstrFileVerion              As String
'ModifyDate����޸�ʱ��(R/W)
Private mdatModifyDate              As Date
'MD5�ļ�MD5(R/W)
Private mstrMD5                     As String
'IsForceCoverǿ�Ƹ���(R/W)
Private mblnForceCover              As Boolean
'RFTRegע��(R/W)
Private mrftReg                     As RegFileType
'===========================================================================
'                2�����Զ���
'===========================================================================
'FileName�ļ�����(R)
Public Property Get FileName() As String
    FileName = mstrFileName
End Property
'FileSTName�ļ���׼���ƣ�ȫ��д(R)
Public Property Get FileSTName() As String
    FileSTName = mstrFileSTName
End Property
'FileType�ļ�����(R)
Public Property Get FileType() As FileType
    FileType = mftFileType
End Property
'FilePath�ļ�ȫ·��(R)
Public Property Get FilePath() As String
    FilePath = mstrFilePath
End Property
'SetupPath��װ·��(R)
Public Property Get SetupPath() As String
    SetupPath = mstrSetupPath
End Property
'FileVerion�ļ��汾(R)
Public Property Get FileVerion() As String
    FileVerion = mstrFileVerion
End Property
'ModifyDate����޸�ʱ��(R)
Public Property Get ModifyDate() As Date
    ModifyDate = mdatModifyDate
End Property
'MD5�ļ�MD5(R)
Public Property Get MD5() As String
    MD5 = mstrMD5
End Property
'IsForceCoverǿ�Ƹ���(R)
Public Property Get IsForceCover() As Boolean
    IsForceCover = mblnForceCover
End Property
'ע��Type(R)
Public Property Get RegType() As RegFileType
    RegType = mrftReg
End Property
'===========================================================================
'                3����������
'===========================================================================
Public Function InitFile(ByVal blnUpBase As Boolean, ByVal strFileName As String, ByVal ftFileType As FileType, ByVal strSetupPath As String, ByVal strFileVerion As String, _
                        ByVal datModifyDate As Date, ByVal strMD5 As String, Optional ByVal blnForceCover As Boolean, Optional ByVal rftRegType As RegFileType, _
                        Optional ByVal strRelatedComponent As String) As Boolean
'���ܣ���ʼ���ļ���Ϣ
'������
'       blnUpBase:�Ƿ���������Ҫ�Ļ�������������������Ԥ����ʱ���������ǰ����Ŀ¼������ʽ�����Ĳ�����
'               ������������������ZLRUNAS.EXE,ZLHISCRUST.EXE,REGCOM.DLL,7Z.EXE,7Z.DLL
'               ZLRUNAS.EXE-ʹ�Զ������Թ���Ա���еĹ��ߡ�ZLHISCRUST.EXE-�Զ��������ߡ�REGCOM.DLL-ע�Ჿ���Ĳ������Ѿ������ˣ�
'               7Z.EXE,7Z.DLL:ѹ����ѹ���������з����������Ľ�ѹ��
'       strFileName:�ļ�����
'       ftFileType���ļ�����
'       strSetupPath:��װ·��
'       strFileVerion:�ļ��汾
'       datModifyDate:����޸�ʱ��
'       strMD5:�ļ�MD5
'       blnForceCover:ǿ�Ƹ���
'       blnReg:ע��
'���أ��Ƿ�ɹ�
    On Error GoTo ErrH
    mstrFileName = strFileName
    mstrFileSTName = UCase(strFileName)
    mftFileType = ftFileType
    mstrSetupPath = strSetupPath
    mstrFileVerion = strFileVerion
    mdatModifyDate = datModifyDate
    mstrMD5 = strMD5
    mblnForceCover = blnForceCover
    mrftReg = rftRegType
    mstrFilePath = GetActualPath(strSetupPath, ftFileType, strFileName)
    InitFile = mstrFilePath = ""
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
'===========================================================================
'                4��˽�з���
'===========================================================================
