VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIDCard"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mfrmReadInfo As frmReadInfo
Private mlngHwnd As Long

Public Event ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
                            
Friend Function GetParent() As Long
    GetParent = mlngHwnd
End Function

Public Sub SetParent(ByVal lngHwnd As Long)
    mlngHwnd = lngHwnd
    If Not mfrmReadInfo Is Nothing Then mfrmReadInfo.glngParentHwnd = GetParentHwnd(lngHwnd)
End Sub
                            
Private Sub Class_Initialize()
    If Init(Me) Then
        lngTmp = Val(GetSetting("ZLSOFT", "公共全局\IDCard", "自动识别间隔", 300))
        Set mfrmReadInfo = New frmReadInfo
        Load mfrmReadInfo
        Set mfrmReadInfo.mobjIDCard = Me
        mfrmReadInfo.Timer1.Interval = IIf(lngTmp < 300 Or lngTmp > 65535, 300, lngTmp) '加载窗体不显示
        mfrmReadInfo.Timer1.Enabled = False
    End If
End Sub

Private Sub Class_Terminate()
    Terminate
    If Not mfrmReadInfo Is Nothing Then
        Set mfrmReadInfo.mobjIDCard = Nothing
        Unload mfrmReadInfo
        Set mfrmReadInfo = Nothing
    End If
End Sub

Public Sub ParameterSet()
    frmParameterSet.Show vbModal
End Sub

Friend Sub ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    RaiseEvent ShowIDCardInfo(strID, strName, strSex, strNation, datBirthday, strAddress)
End Sub

Public Function GetPhotoAsStdPicture(picPhoto As StdPicture)
    If mfrmReadInfo Is Nothing Then Exit Function
    Set picPhoto = mfrmReadInfo.ReadPhotoInfo
End Function

Public Sub ReadIDCard(strID As String, strName As String, strSex As String, _
                            strNation As String, datBirthday As Date, strAddress As String, picPhoto As StdPicture)
    If mfrmReadInfo Is Nothing Then Exit Sub
    Call mfrmReadInfo.Read_Card(strID, strName, strSex, strNation, datBirthday, strAddress)
    Set picPhoto = mfrmReadInfo.ReadPhotoInfo
End Sub

Public Sub SetEnabled(ByVal blnEnabled As Boolean)
    If mfrmReadInfo Is Nothing Then Exit Sub
    Call SetAutoRead(mfrmReadInfo.Timer1, blnEnabled)
    '108802:李南春，2017/5/10，激活时重新赋予对象
    If blnEnabled = False Then
        Set mfrmReadInfo.mobjIDCard = Nothing
    Else
        Set mfrmReadInfo.mobjIDCard = Me
        mfrmReadInfo.glngParentHwnd = GetParentHwnd(mlngHwnd)
    End If
End Sub

