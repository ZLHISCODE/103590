Attribute VB_Name = "mdlZip"
Option Explicit

Public Type ZIPnames
    S(0 To 99) As String
End Type



'Structure ZCL - not used by VB
'Private Type ZCL
'    argc As Long            'number of files
'    filename As String      'Name of the Zip file
'    fileArray As ZIPnames   'The array of filenames
'End Type

' Call back "string" (sic)
' Callback large "string" (sic)
Private Type CBChar
    ch(4096) As Byte
End Type

' Callback small "string" (sic)
Private Type CBCh
    ch(256) As Byte
End Type


'ȡIP��API
Public Const MAX_ADAPTER_NAME_LENGTH         As Long = 256
Public Const MAX_ADAPTER_DESCRIPTION_LENGTH  As Long = 128
Public Const MAX_ADAPTER_ADDRESS_LENGTH      As Long = 8

Public Type IP_ADDRESS_STRING
    IpAddr(0 To 15)  As Byte
End Type
Public Type IP_MASK_STRING
    IpMask(0 To 15)  As Byte
End Type

Public Type IP_ADDR_STRING
    dwNext     As Long
    IpAddress  As IP_ADDRESS_STRING
    IpMask     As IP_MASK_STRING
    dwContext  As Long
End Type

Public Type IP_ADAPTER_INFO
  dwNext                As Long
  ComboIndex            As Long  '����
  sAdapterName(0 To (MAX_ADAPTER_NAME_LENGTH + 3))        As Byte
  sDescription(0 To (MAX_ADAPTER_DESCRIPTION_LENGTH + 3)) As Byte
  dwAddressLength       As Long
  sIPAddress(0 To (MAX_ADAPTER_ADDRESS_LENGTH - 1))       As Byte
  dwIndex               As Long
  uType                 As Long
  uDhcpEnabled          As Long
  CurrentIpAddress      As Long
  IpAddressList         As IP_ADDR_STRING
  GatewayList           As IP_ADDR_STRING
  DhcpServer            As IP_ADDR_STRING
  bHaveWins             As Long
  PrimaryWinsServer     As IP_ADDR_STRING
  SecondaryWinsServer   As IP_ADDR_STRING
  LeaseObtained         As Long
  LeaseExpires          As Long
End Type

Public Function FnPtr(ByVal lp As Long) As Long
'���ܣ�ȡ�ú�����ָ��ֵ
    FnPtr = lp
End Function

' Callback for unzip32.dll
Public Sub ReceiveDllMessage(ByVal ucsize As Long, _
    ByVal csiz As Long, _
    ByVal cfactor As Integer, _
    ByVal mo As Integer, _
    ByVal dy As Integer, _
    ByVal yr As Integer, _
    ByVal HH As Integer, _
    ByVal mm As Integer, _
    ByVal C As Byte, ByRef fname As CBCh, _
    ByRef meth As CBCh, ByVal crc As Long, _
    ByVal fCrypt As Byte)

'���ս�ѹ�����з��ص���Ϣ
    Dim strTemp As String, lngCount As Long
    Dim strInfo As String * 80

    ' always put this in callback routines!
    On Error Resume Next
    strInfo = Space(80)
'    If vbzipnum = 0 Then
'        Mid$(strInfo, 1, 50) = "Filename:"
'        Mid$(strInfo, 53, 4) = "Size"
'        Mid$(strInfo, 62, 4) = "Date"
'        Mid$(strInfo, 71, 4) = "Time"
'        vbzipmes = strInfo + vbCrLf
'        strInfo = Space(80)
'    End If
    strTemp = ""
    For lngCount = 0 To 255
        If fname.ch(lngCount) = 0 Then
            lngCount = 99999
        Else
            strTemp = strTemp & Chr$(fname.ch(lngCount))
        End If
    Next lngCount
    Mid$(strInfo, 1, 50) = Mid$(strTemp, 1, 50)
    Mid$(strInfo, 51, 7) = Right$("        " + str$(ucsize), 7)
    Mid$(strInfo, 60, 3) = Right$(str$(dy), 2) + "/"
    Mid$(strInfo, 63, 3) = Right$("0" + Trim$(str$(mo)), 2) + "/"
    Mid$(strInfo, 66, 2) = Right$("0" + Trim$(str$(yr)), 2)
    Mid$(strInfo, 70, 3) = Right$(str$(HH), 2) + ":"
    Mid$(strInfo, 73, 2) = Right$("0" + Trim$(str$(mm)), 2)
    ' Mid$(strInfo, 75, 2) = Right$(" " + Str$(cfactor), 2)
    ' Mid$(strInfo, 78, 8) = Right$("        " + Str$(csiz), 8)
    ' strTemp = ""
    ' For lngCount = 0 To 255
    '     If meth.ch(lngCount) = 0 Then lngCount = 99999 Else strTemp = strTemp + Chr(meth.ch(lngCount))
    ' Next lngCount
    '��ѹ���ļ�����
'    vbzipmes = vbzipmes + strInfo + vbCrLf
'    vbzipnum = vbzipnum + 1
End Sub

' Callback for unzip32.dll
Public Function DllPrnt(ByRef fname As CBChar, ByVal lngLength As Long) As Long
    Dim strTemp As String, lngCount As Long

    ' always put this in callback routines!
    On Error Resume Next
    strTemp = ""
    For lngCount = 0 To lngLength
        If fname.ch(lngCount) = 0 Then
            lngCount = 99999
        Else
            strTemp = strTemp + Chr(fname.ch(lngCount))
        End If
    Next lngCount
    DllPrnt = 0
End Function

' Callback for unzip32.dll
Public Function DllPass(ByRef s1 As Byte, X As Long, _
    ByRef s2 As Byte, _
    ByRef s3 As Byte) As Long

    ' always put this in callback routines!
    On Error Resume Next
    ' not supported - always return 1
    DllPass = 1
End Function

Public Function DllRep(ByRef fname As CBChar) As Long
'���ܣ��ļ�����ʱ�����֡��Ƿ��滻�ļ�������Ϣ
'      ��unzip32.dll����

    Dim strTemp As String, lngCount As Long
    
    On Error Resume Next
    
    DllRep = 100 ' 100=do not overwrite - keep asking user
    '����ļ���
    strTemp = ""
    For lngCount = 0 To 255
        If fname.ch(lngCount) = 0 Then
            lngCount = 99999
        Else
            strTemp = strTemp + Chr(fname.ch(lngCount))
        End If
    Next lngCount
    
    lngCount = MsgBox("�ļ���" + strTemp + "���Ѿ����ڣ��Ƿ��滻��", vbQuestion Or vbYesNoCancel, gstrSysName)
    
    If lngCount = vbNo Then Exit Function
    If lngCount = vbCancel Then
        DllRep = 104 ' 104=overwrite none
        Exit Function
    End If
    DllRep = 102 ' 102=overwrite 103=overwrite all
End Function

Public Function szTrim(szString As String) As String
'���ܣ�ȥ��\0�Ժ���ַ���ASCIIZ to String
    
    Dim pos As Integer, ln As Integer

    pos = InStr(szString, Chr$(0))
    ln = Len(szString)
    Select Case pos
        Case Is > 1
            szTrim = Trim(Left(szString, pos - 1))
        Case 1
            szTrim = ""
        Case Else
            szTrim = Trim(szString)
    End Select
End Function

' Callback for zip32.dll
Public Function DllComm(ByRef s1 As CBChar) As CBChar
    
    ' always put this in callback routines!
    On Error Resume Next
    ' not supported always return \0
    s1.ch(0) = vbNullString
    DllComm = s1
End Function

' Main subroutine
Public Function VBUnzip(fname As String, extdir As String, _
    prom As Integer, over As Integer, _
    mess As Integer, dirs As Integer, numfiles As Long, numxfiles As Long, _
    vbzipnam As ZIPnames, vbxnames As ZIPnames) As Boolean
'���ܣ���ѹ����
'����˵��
'    zipfile    ҪUnzip���ļ�
'    unzipdir   ���ý�ѹ���ļ���Ŀ¼
'    prom       1 = ���ڸ��ǽ�����ʾ
'    over       1 = ���Ǹ���
'    mess       1 = ֻ�г��ļ�����  0 = ��ѹ
'    dirs       1 = ����ZIP�ļ��е�·��
'    vbzipnam  ��ѡ�Ľ�ѹ���ļ�
'    vbxnames  Ҫ���ų��Ľ�ѹ�ļ�
    
    Dim lngCount As Long ' , s1 As String * 20, s2 As String * 256
    
    Dim MYUSER As USERFUNCTION
    Dim MYDCL As DCLIST
    Dim MYVER As UZPVER

    ' Set options
    With MYDCL
        .ExtractOnlyNewer = 0      ' 1=extract only newer
        .SpaceToUnderscore = 0     ' 1=convert space to underscore
        .PromptToOverwrite = prom  ' 1=prompt to overwrite required
        .fQuiet = 0                ' 2=no messages 1=less 0=all
        .ncflag = 0                ' 1=write to stdout
        .ntflag = 0                ' 1=test zip
        .nvflag = mess             ' 0=extract 1=list contents
        .nUflag = 0                ' 1=extract only newer
        .nzflag = 0                ' 1=display zip file comment
        .ndflag = dirs             ' 1=honour directories
        .noflag = over              ' 1=overwrite files
        .naflag = 0                ' 1=convert CR to CRLF
        .nZIflag = 0               ' 1=Zip Info Verbose
        .C_flag = 0                ' 1=Case insensitivity, 0=Case Sensitivity
        .fPrivilege = 0            ' 1=ACL 2=priv
        .Zip = fname               ' ZIP name
        .ExtractDir = extdir       ' Extraction directory, NULL if extracting
    End With                              ' to current directory
    
    '�����ڲ������ĵ�ַ
    With MYUSER
        .DllPrnt = FnPtr(AddressOf DllPrnt)
        .DLLSND = 0& ' not supported
        .DLLREPLACE = FnPtr(AddressOf DllRep)
        .DLLPASSWORD = FnPtr(AddressOf DllPass)
        .DLLMESSAGE = FnPtr(AddressOf ReceiveDllMessage)
        .DLLSERVICE = 0& ' not coded yet :)
    End With
    ' Set Version space
    ' Do not change
    With MYVER
        .structlen = Len(MYVER)
        .beta = Space(9) & vbNullChar
        .date = Space(19) & vbNullChar
        .zlib = Space(9) & vbNullChar
    End With
    
    On Error Resume Next
    ' Get version
    Call UzpVersion2(MYVER)
    
    ' Go for it!
    lngCount = windll_unzip(numfiles, vbzipnam, _
        numxfiles, vbxnames, MYDCL, MYUSER)
    If err <> 0 Then
        'û������DLL
        err.Clear
        VBUnzip = False
        MsgBox "�����ļ� " & fname & " ��ѹʧ�ܡ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    If lngCount = 0 Then
        VBUnzip = True
    Else
        VBUnzip = False
        MsgBox "�����ļ� " & fname & " ��ѹʧ�ܡ�", vbInformation, gstrSysName
    End If
End Function

'Main Subroutine
Public Function VBZip(argc As Integer, zipname As String, _
        mynames As ZIPnames, junk As Integer, _
        recurse As Integer, updat As Integer, _
        freshen As Integer, basename As String) As Boolean
        
'���ܣ�ѹ���ļ�
'������argc         �ļ�����
'      zipname      ZIP�ļ���
'      mynames      Ҫѹ�����ļ��б�
'      junk         1 �׿�Ŀ¼��
'      recurse      ZIP�ļ���
'      updat        ZIP�ļ���
    Dim hMem As Long, lngCount As Integer
    Dim retcode As Long
    Dim MYOPT As ZPOPT
    Dim MYUSER As ZIPUSERFUNCTIONS
    
    On Error Resume Next ' nothing will go wrong :-)
    
    '�����ڲ������ĵ�ַ
    With MYUSER
        .DllPrnt = FnPtr(AddressOf DllPrnt)
        .DLLPASSWORD = FnPtr(AddressOf DllPass)
        .DLLCOMMENT = FnPtr(AddressOf DllComm)
        .DLLSERVICE = 0& ' not coded yet :-)
    End With
    retcode = ZpInit(MYUSER)
    
    '����ѹ��ѡ��
    With MYOPT
        .fSuffix = 0        ' include suffixes (not yet implemented)
        .fEncrypt = 0       ' 1 if encryption wanted
        .fSystem = 0        ' 1 to include system/hidden files
        .fVolume = 0        ' 1 if storing volume label
        .fExtra = 0         ' 1 if including extra attributes
        .fNoDirEntries = 0  ' 1 if ignoring directory entries
        .fExcludeDate = 0   ' 1 if excluding files earlier than a specified date
        .fIncludeDate = 0   ' 1 if including files earlier than a specified date
        .fVerbose = 0       ' 1 if full messages wanted
        .fQuiet = 0         ' 1 if minimum messages wanted
        .fCRLF_LF = 0       ' 1 if translate CR/LF to LF
        .fLF_CRLF = 0       ' 1 if translate LF to CR/LF
        .fJunkDir = junk    ' 1 if junking directory names
        .fRecurse = recurse ' 1 if recursing into subdirectories
        .fGrow = 0          ' 1 if allow appending to zip file
        .fForce = 0         ' 1 if making entries using DOS names
        .fMove = 0          ' 1 if deleting files added or updated
        .fDeleteEntries = 0 ' 1 if files passed have to be deleted
        .fUpdate = updat    ' 1 if updating zip file--overwrite only if newer
        .fFreshen = freshen ' 1 if freshening zip file--overwrite only
        .fJunkSFX = 0       ' 1 if junking sfx prefix
        .fLatestTime = 0    ' 1 if setting zip file time to time of latest file in archive
        .fComment = 0       ' 1 if putting comment in zip file
        .fOffsets = 0       ' 1 if updating archive offsets for sfx Files
        .fPrivilege = 0     ' 1 if not saving privelages
        .fEncryption = 0    'Read only property!
        .fRepair = 0        ' 1=> fix archive, 2=> try harder to fix
        .flevel = 0         ' compression level - should be 0!!!
        .date = vbNullString ' "12/31/79"? US Date?
        .szRootDir = basename
    End With
    ' Set options
    retcode = ZpSetOptions(MYOPT)
    
    ' ZCL not needed in VB
    ' MYZCL.argc = 2
    ' MYZCL.filename = "c:\wiz\new.zip"
    ' MYZCL.fileArray = MYNAMES
    
    ' Go for it!
    retcode = ZpArchive(argc, zipname, mynames)
    If err <> 0 Then
        'û������DLL
        err.Clear
        VBZip = False
        MsgBox "�ļ� " & zipname & " ѹ��ʧ�ܡ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    If retcode = 0 Then
        VBZip = True
    Else
        VBZip = False
        MsgBox "�ļ� " & zipname & " ѹ��ʧ�ܡ�", vbInformation, gstrSysName
    End If
End Function

Public Function AnalyseIP() As String
    '������:����
    '����:������������IP��ַ
    Dim cbRequired  As Long
    Dim buff()      As Byte
    Dim Adapter     As IP_ADAPTER_INFO
    Dim AdapterStr  As IP_ADDR_STRING
    Dim ptr1        As Long
    Dim sIPAddr     As String
    Dim found       As Boolean
    Call GetAdaptersInfo(ByVal 0&, cbRequired)
    If cbRequired > 0 Then
        ReDim buff(0 To cbRequired - 1) As Byte
        If GetAdaptersInfo(buff(0), cbRequired) = ERROR_SUCCESS Then
            '��ȡ�����buff()�е����ݵ�ָ��
            ptr1 = VarPtr(buff(0))
            Do While (ptr1 <> 0)
                '����һ������������ת����IP_ADAPTER_INFO�ṹ��
                 CopyMemory Adapter, ByVal ptr1, LenB(Adapter)
                 With Adapter
                    'IpAddress.IpAddr��Ա������DHCP��IP��ַ
                    sIPAddr = TrimNull(StrConv(.IpAddressList.IpAddress.IpAddr, vbUnicode))
                    If Len(sIPAddr) > 0 Then
                        found = True
                        Exit Do
                    End If
                    ptr1 = .dwNext
                 End With  'With Adapter
            '����������ʱ��ptr1��ֵΪ0
            Loop  'Do While (ptr1 <> 0)
        End If  'If GetAdaptersInfo
    End If  'If cbRequired > 0
    '���ؽ���ַ���
    AnalyseIP = sIPAddr
End Function

Public Function TrimNull(Item As String)
    Dim pos As Integer
    pos = InStr(Item, Chr$(0))
    If pos Then
          TrimNull = Left$(Item, pos - 1)
    Else: TrimNull = Item
    End If
End Function

Public Function GetMyCompterName() As String
    '����:��ȡ�������
    '��ȡ�������
    Dim strComputerName As String * 256
    err = 0
    On Error Resume Next
    
    Call GetComputerName(strComputerName, 255)
    GetMyCompterName = Trim(Replace(strComputerName, Chr(0), ""))
End Function

