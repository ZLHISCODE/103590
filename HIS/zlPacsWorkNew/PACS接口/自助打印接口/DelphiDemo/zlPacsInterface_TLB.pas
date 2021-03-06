unit zlPacsInterface_TLB;

// ************************************************************************ //
// WARNING                                                                    
// -------                                                                    
// The types declared in this file were generated from data read from a       
// Type Library. If this type library is explicitly or indirectly (via        
// another type library referring to this type library) re-imported, or the   
// 'Refresh' command of the Type Library Editor activated while editing the   
// Type Library, the contents of this file will be regenerated and all        
// manual modifications will be lost.                                         
// ************************************************************************ //

// PASTLWTR : 1.2
// File generated on 2011/7/28 17:45:11 from Type Library described below.

// ************************************************************************  //
// Type Lib: E:\ZLHIS\ZLPacsWork\通用接口\zlPacsInterface.dll (1)
// LIBID: {AB5BB424-129E-43A0-A797-0A9819B954E4}
// LCID: 0
// Helpfile: 
// HelpString: 
// DepndLst: 
//   (1) v2.0 stdole, (C:\Windows\system32\stdole2.tlb)
//   (2) v2.6 ADODB, (C:\Program Files\Common Files\System\ado\msado26.tlb)
// Errors:
//   Error creating palette bitmap of (TclsPacsInterface) : Server E:\ZLHIS\ZLPacsWork\通用接口\zlPacsInterface.dll contains no icons
// ************************************************************************ //
// *************************************************************************//
// NOTE:                                                                      
// Items guarded by $IFDEF_LIVE_SERVER_AT_DESIGN_TIME are used by properties  
// which return objects that may need to be explicitly created via a function 
// call prior to any access via the property. These items have been disabled  
// in order to prevent accidental use from within the object inspector. You   
// may enable them by defining LIVE_SERVER_AT_DESIGN_TIME or by selectively   
// removing them from the $IFDEF blocks. However, such items must still be    
// programmatically created via a method of the appropriate CoClass before    
// they can be used.                                                          
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
interface

uses Windows, ActiveX, ADODB_TLB, Classes, Graphics, OleServer, StdVCL, Variants;
  


// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  zlPacsInterfaceMajorVersion = 1;
  zlPacsInterfaceMinorVersion = 0;

  LIBID_zlPacsInterface: TGUID = '{AB5BB424-129E-43A0-A797-0A9819B954E4}';

  IID__clsPacsInterface: TGUID = '{19AA3D5A-CD0C-4ACF-9BD4-43512F9A34C3}';
  CLASS_clsPacsInterface: TGUID = '{E0A958B7-9A2E-449C-B4B0-EE18BF1812ED}';

// *********************************************************************//
// Declaration of Enumerations defined in Type Library                    
// *********************************************************************//
// Constants for enum TErrorShowType
type
  TErrorShowType = TOleEnum;
const
  estNoDisplay = $00000001;
  estShowMsg = $00000002;

// Constants for enum TPatientWhereType
type
  TPatientWhereType = TOleEnum;
const
  pwtPatientId = $00000001;
  pwtInHospital = $00000002;
  pwtOutPatient = $00000003;
  pwtSickCard = $00000004;
  pwtIdCard = $00000005;
  pwtHealthNum = $00000006;
  pwtPatientName = $00000007;

// Constants for enum TRequestWhereType
type
  TRequestWhereType = TOleEnum;
const
  rwtPatientId = $00000001;
  rwtInHospital = $00000002;
  rwtOutPatient = $00000003;
  rwtSickCard = $00000004;
  rwtIdCard = $00000005;
  rwtHealthNum = $00000006;
  rwtPatientName = $00000007;
  rwtAdviceId = $00000008;

type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  _clsPacsInterface = interface;
  _clsPacsInterfaceDisp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  clsPacsInterface = _clsPacsInterface;


// *********************************************************************//
// Declaration of structures, unions and aliases.                         
// *********************************************************************//
  TCusTable = packed record
    strDatas: PSafeArray;
    strColumns: PSafeArray;
  end;


// *********************************************************************//
// Interface: _clsPacsInterface
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {19AA3D5A-CD0C-4ACF-9BD4-43512F9A34C3}
// *********************************************************************//
  _clsPacsInterface = interface(IDispatch)
    ['{19AA3D5A-CD0C-4ACF-9BD4-43512F9A34C3}']
    function GetErrShowType: TErrorShowType; safecall;
    function GetSplitChar: WideString; safecall;
    function GetSysNo: Integer; safecall;
    function GetSysOwner: WideString; safecall;
    function Get_Tables: TCusTable; safecall;
    function GetNullValue: WideString; safecall;
    function InitInterface(const strServerName: WideString; const strUserName: WideString; 
                           const strUserPwd: WideString; SysNo: Integer; 
                           const SysOwner: WideString; const NullValue: WideString; 
                           const SplitChar: WideString; errType: TErrorShowType): WordBool; safecall;
    function GetLastError: WideString; safecall;
    function BeginTrans: WordBool; safecall;
    function CommitTrans: WordBool; safecall;
    function RollbackTrans: WordBool; safecall;
    function ExecutePacsProcedure(const strProcedureName: WideString): WordBool; safecall;
    function GetPacsCursor(const strProcedureName: WideString; const strFilterValue: WideString; 
                           var blnIsNoParamerer: WordBool): WordBool; safecall;
    function GetAdoData(const strProcedureName: WideString; const strFilterValue: WideString; 
                        var blnIsNoParamerer: WordBool): _Recordset; safecall;
    function GetRecordValueByColumnName(var strDatas: PSafeArray; var strColumns: PSafeArray; 
                                        lngRecordIndex: Integer; const strCurColumn: WideString): WideString; safecall;
    function GetRecordValueByColumnIndex(var strDatas: PSafeArray; lngRecordIndex: Integer; 
                                         lngColumnIndex: Integer): WideString; safecall;
    function GetCurValueByColumnName(lngRecordIndex: Integer; const strCurColumn: WideString): WideString; safecall;
    function GetCurRecordCount: Integer; safecall;
    function GetCurColumnCount: Integer; safecall;
    function GetCurColumnIndex(const strCurColumn: WideString): Integer; safecall;
    function GetCurValueByColumnIndex(lngRecordIndex: Integer; lngColumnIndex: Integer): WideString; safecall;
    function GetCurRecordData(lngRecordIndex: Integer): WideString; safecall;
    function GetRecordData(var strDatas: PSafeArray; lngRecordIndex: Integer): WideString; safecall;
    function GetRecordCount(var strDatas: PSafeArray): Integer; safecall;
    function GetColumnCount(var strColumns: PSafeArray): Integer; safecall;
    function GetColumnIndex(var strColumns: PSafeArray; const strCurColumn: WideString): Smallint; safecall;
    function GetColumnName(var strColumns: PSafeArray; columnIndex: Smallint): WideString; safecall;
    function GetCurColumnName(columnIndex: Smallint): WideString; safecall;
    function GetDeptItems(const strFilter: WideString): WordBool; safecall;
    function GetChargeTypes(const strFilter: WideString): WordBool; safecall;
    function GetPacsItems(const strFilter: WideString): WordBool; safecall;
    function GetAdviceItems(lngAdviceKey: Integer): WordBool; safecall;
    function GetAdviceFees(lngAdviceKey: Integer): WordBool; safecall;
    function GetPatientInfo(const strQueryKey: WideString; lngWhereType: TPatientWhereType): WordBool; safecall;
    function GetRequestInfo(const strQueryKey: WideString; lngWhereType: TRequestWhereType): WordBool; safecall;
    function GetRequestExecuteStatus(lngAdviceKey: Integer): Integer; safecall;
    function GetRequestAdviceStatus(lngAdviceKey: Integer): Integer; safecall;
    function GetRequestExeProcedureStatus(lngAdviceKey: Integer): Integer; safecall;
    function CancelRequest(lngAdviceKey: Integer; lngExecOne: Integer): WordBool; safecall;
    function RecevieRequest(lngAdviceKey: Integer; const strExeRoom: WideString; 
                            lngStudyNo: Integer; const strDevice: WideString; lngHeight: Integer; 
                            lngWeight: Integer; const strStudyDoc: WideString; 
                            StrExeDate: TDateTime; const strExeDes: WideString; lngExecOne: Integer): WordBool; safecall;
    function ModifyRequest(lngAdviceKey: Integer; const strExeRoom: WideString; 
                           lngStudyNo: Integer; const strDevice: WideString; lngHeight: Integer; 
                           lngWeight: Integer; const strStudyDoc: WideString; 
                           StrExeDate: TDateTime; const strExeDes: WideString; lngExecOne: Integer): WordBool; safecall;
    function DeleteReport(lngAdviceKey: Integer): WordBool; safecall;
    function DeleteElectrocardioReport(lngAdviceKey: Integer): WordBool; safecall;
    function SendReport(lngAdviceKey: Integer; const strReportView: WideString; 
                        const strReportAdvice: WideString; const strReportDoctor: WideString; 
                        const strAuditingDoctor: WideString): WordBool; safecall;
    function SendElectrocardioReport(lngAdviceKey: Integer; const strReportTitle: WideString; 
                                     const strReportImgFiles: WideString; 
                                     const strReportResult: WideString; 
                                     const strReportAdvice: WideString; 
                                     const strReportDoctor: WideString; 
                                     const strAuditingDoctor: WideString): WordBool; safecall;
    function SendReportImages(lngAdviceKey: Integer; const strImgFiles: WideString): WordBool; safecall;
    function SendReportAffix(lngAdviceKey: Integer; const strAffixFiles: WideString): WordBool; safecall;
    property Tables: TCusTable read Get_Tables;
  end;

// *********************************************************************//
// DispIntf:  _clsPacsInterfaceDisp
// Flags:     (4560) Hidden Dual NonExtensible OleAutomation Dispatchable
// GUID:      {19AA3D5A-CD0C-4ACF-9BD4-43512F9A34C3}
// *********************************************************************//
  _clsPacsInterfaceDisp = dispinterface
    ['{19AA3D5A-CD0C-4ACF-9BD4-43512F9A34C3}']
    function GetErrShowType: TErrorShowType; dispid 1610809345;
    function GetSplitChar: WideString; dispid 1610809346;
    function GetSysNo: Integer; dispid 1610809347;
    function GetSysOwner: WideString; dispid 1610809348;
    property Tables: {??TCusTable}OleVariant readonly dispid 1745027072;
    function GetNullValue: WideString; dispid 1610809349;
    function InitInterface(const strServerName: WideString; const strUserName: WideString; 
                           const strUserPwd: WideString; SysNo: Integer; 
                           const SysOwner: WideString; const NullValue: WideString; 
                           const SplitChar: WideString; errType: TErrorShowType): WordBool; dispid 1610809350;
    function GetLastError: WideString; dispid 1610809352;
    function BeginTrans: WordBool; dispid 1610809353;
    function CommitTrans: WordBool; dispid 1610809354;
    function RollbackTrans: WordBool; dispid 1610809355;
    function ExecutePacsProcedure(const strProcedureName: WideString): WordBool; dispid 1610809356;
    function GetPacsCursor(const strProcedureName: WideString; const strFilterValue: WideString; 
                           var blnIsNoParamerer: WordBool): WordBool; dispid 1610809357;
    function GetAdoData(const strProcedureName: WideString; const strFilterValue: WideString; 
                        var blnIsNoParamerer: WordBool): _Recordset; dispid 1610809358;
    function GetRecordValueByColumnName(var strDatas: {??PSafeArray}OleVariant; 
                                        var strColumns: {??PSafeArray}OleVariant; 
                                        lngRecordIndex: Integer; const strCurColumn: WideString): WideString; dispid 1610809359;
    function GetRecordValueByColumnIndex(var strDatas: {??PSafeArray}OleVariant; 
                                         lngRecordIndex: Integer; lngColumnIndex: Integer): WideString; dispid 1610809360;
    function GetCurValueByColumnName(lngRecordIndex: Integer; const strCurColumn: WideString): WideString; dispid 1610809361;
    function GetCurRecordCount: Integer; dispid 1610809362;
    function GetCurColumnCount: Integer; dispid 1610809363;
    function GetCurColumnIndex(const strCurColumn: WideString): Integer; dispid 1610809364;
    function GetCurValueByColumnIndex(lngRecordIndex: Integer; lngColumnIndex: Integer): WideString; dispid 1610809365;
    function GetCurRecordData(lngRecordIndex: Integer): WideString; dispid 1610809366;
    function GetRecordData(var strDatas: {??PSafeArray}OleVariant; lngRecordIndex: Integer): WideString; dispid 1610809367;
    function GetRecordCount(var strDatas: {??PSafeArray}OleVariant): Integer; dispid 1610809368;
    function GetColumnCount(var strColumns: {??PSafeArray}OleVariant): Integer; dispid 1610809369;
    function GetColumnIndex(var strColumns: {??PSafeArray}OleVariant; const strCurColumn: WideString): Smallint; dispid 1610809370;
    function GetColumnName(var strColumns: {??PSafeArray}OleVariant; columnIndex: Smallint): WideString; dispid 1610809371;
    function GetCurColumnName(columnIndex: Smallint): WideString; dispid 1610809372;
    function GetDeptItems(const strFilter: WideString): WordBool; dispid 1610809373;
    function GetChargeTypes(const strFilter: WideString): WordBool; dispid 1610809374;
    function GetPacsItems(const strFilter: WideString): WordBool; dispid 1610809375;
    function GetAdviceItems(lngAdviceKey: Integer): WordBool; dispid 1610809376;
    function GetAdviceFees(lngAdviceKey: Integer): WordBool; dispid 1610809377;
    function GetPatientInfo(const strQueryKey: WideString; lngWhereType: TPatientWhereType): WordBool; dispid 1610809378;
    function GetRequestInfo(const strQueryKey: WideString; lngWhereType: TRequestWhereType): WordBool; dispid 1610809379;
    function GetRequestExecuteStatus(lngAdviceKey: Integer): Integer; dispid 1610809380;
    function GetRequestAdviceStatus(lngAdviceKey: Integer): Integer; dispid 1610809381;
    function GetRequestExeProcedureStatus(lngAdviceKey: Integer): Integer; dispid 1610809382;
    function CancelRequest(lngAdviceKey: Integer; lngExecOne: Integer): WordBool; dispid 1610809383;
    function RecevieRequest(lngAdviceKey: Integer; const strExeRoom: WideString; 
                            lngStudyNo: Integer; const strDevice: WideString; lngHeight: Integer; 
                            lngWeight: Integer; const strStudyDoc: WideString; 
                            StrExeDate: TDateTime; const strExeDes: WideString; lngExecOne: Integer): WordBool; dispid 1610809384;
    function ModifyRequest(lngAdviceKey: Integer; const strExeRoom: WideString; 
                           lngStudyNo: Integer; const strDevice: WideString; lngHeight: Integer; 
                           lngWeight: Integer; const strStudyDoc: WideString; 
                           StrExeDate: TDateTime; const strExeDes: WideString; lngExecOne: Integer): WordBool; dispid 1610809385;
    function DeleteReport(lngAdviceKey: Integer): WordBool; dispid 1610809386;
    function DeleteElectrocardioReport(lngAdviceKey: Integer): WordBool; dispid 1610809387;
    function SendReport(lngAdviceKey: Integer; const strReportView: WideString; 
                        const strReportAdvice: WideString; const strReportDoctor: WideString; 
                        const strAuditingDoctor: WideString): WordBool; dispid 1610809388;
    function SendElectrocardioReport(lngAdviceKey: Integer; const strReportTitle: WideString; 
                                     const strReportImgFiles: WideString; 
                                     const strReportResult: WideString; 
                                     const strReportAdvice: WideString; 
                                     const strReportDoctor: WideString; 
                                     const strAuditingDoctor: WideString): WordBool; dispid 1610809389;
    function SendReportImages(lngAdviceKey: Integer; const strImgFiles: WideString): WordBool; dispid 1610809392;
    function SendReportAffix(lngAdviceKey: Integer; const strAffixFiles: WideString): WordBool; dispid 1610809393;
  end;

// *********************************************************************//
// The Class CoclsPacsInterface provides a Create and CreateRemote method to          
// create instances of the default interface _clsPacsInterface exposed by              
// the CoClass clsPacsInterface. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoclsPacsInterface = class
    class function Create: _clsPacsInterface;
    class function CreateRemote(const MachineName: string): _clsPacsInterface;
  end;


// *********************************************************************//
// OLE Server Proxy class declaration
// Server Object    : TclsPacsInterface
// Help String      : 
// Default Interface: _clsPacsInterface
// Def. Intf. DISP? : No
// Event   Interface: 
// TypeFlags        : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TclsPacsInterfaceProperties= class;
{$ENDIF}
  TclsPacsInterface = class(TOleServer)
  private
    FIntf:        _clsPacsInterface;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TclsPacsInterfaceProperties;
    function      GetServerProperties: TclsPacsInterfaceProperties;
{$ENDIF}
    function      GetDefaultInterface: _clsPacsInterface;
  protected
    procedure InitServerData; override;
    function Get_Tables: TCusTable;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: _clsPacsInterface);
    procedure Disconnect; override;
    function GetErrShowType: TErrorShowType;
    function GetSplitChar: WideString;
    function GetSysNo: Integer;
    function GetSysOwner: WideString;
    function GetNullValue: WideString;
    function InitInterface(const strServerName: WideString; const strUserName: WideString; 
                           const strUserPwd: WideString; SysNo: Integer; 
                           const SysOwner: WideString; const NullValue: WideString; 
                           const SplitChar: WideString; errType: TErrorShowType): WordBool;
    function GetLastError: WideString;
    function BeginTrans: WordBool;
    function CommitTrans: WordBool;
    function RollbackTrans: WordBool;
    function ExecutePacsProcedure(const strProcedureName: WideString): WordBool;
    function GetPacsCursor(const strProcedureName: WideString; const strFilterValue: WideString; 
                           var blnIsNoParamerer: WordBool): WordBool;
    function GetAdoData(const strProcedureName: WideString; const strFilterValue: WideString; 
                        var blnIsNoParamerer: WordBool): _Recordset;
    function GetRecordValueByColumnName(var strDatas: PSafeArray; var strColumns: PSafeArray; 
                                        lngRecordIndex: Integer; const strCurColumn: WideString): WideString;
    function GetRecordValueByColumnIndex(var strDatas: PSafeArray; lngRecordIndex: Integer; 
                                         lngColumnIndex: Integer): WideString;
    function GetCurValueByColumnName(lngRecordIndex: Integer; const strCurColumn: WideString): WideString;
    function GetCurRecordCount: Integer;
    function GetCurColumnCount: Integer;
    function GetCurColumnIndex(const strCurColumn: WideString): Integer;
    function GetCurValueByColumnIndex(lngRecordIndex: Integer; lngColumnIndex: Integer): WideString;
    function GetCurRecordData(lngRecordIndex: Integer): WideString;
    function GetRecordData(var strDatas: PSafeArray; lngRecordIndex: Integer): WideString;
    function GetRecordCount(var strDatas: PSafeArray): Integer;
    function GetColumnCount(var strColumns: PSafeArray): Integer;
    function GetColumnIndex(var strColumns: PSafeArray; const strCurColumn: WideString): Smallint;
    function GetColumnName(var strColumns: PSafeArray; columnIndex: Smallint): WideString;
    function GetCurColumnName(columnIndex: Smallint): WideString;
    function GetDeptItems(const strFilter: WideString): WordBool;
    function GetChargeTypes(const strFilter: WideString): WordBool;
    function GetPacsItems(const strFilter: WideString): WordBool;
    function GetAdviceItems(lngAdviceKey: Integer): WordBool;
    function GetAdviceFees(lngAdviceKey: Integer): WordBool;
    function GetPatientInfo(const strQueryKey: WideString; lngWhereType: TPatientWhereType): WordBool;
    function GetRequestInfo(const strQueryKey: WideString; lngWhereType: TRequestWhereType): WordBool;
    function GetRequestExecuteStatus(lngAdviceKey: Integer): Integer;
    function GetRequestAdviceStatus(lngAdviceKey: Integer): Integer;
    function GetRequestExeProcedureStatus(lngAdviceKey: Integer): Integer;
    function CancelRequest(lngAdviceKey: Integer; lngExecOne: Integer): WordBool;
    function RecevieRequest(lngAdviceKey: Integer; const strExeRoom: WideString; 
                            lngStudyNo: Integer; const strDevice: WideString; lngHeight: Integer; 
                            lngWeight: Integer; const strStudyDoc: WideString; 
                            StrExeDate: TDateTime; const strExeDes: WideString; lngExecOne: Integer): WordBool;
    function ModifyRequest(lngAdviceKey: Integer; const strExeRoom: WideString; 
                           lngStudyNo: Integer; const strDevice: WideString; lngHeight: Integer; 
                           lngWeight: Integer; const strStudyDoc: WideString; 
                           StrExeDate: TDateTime; const strExeDes: WideString; lngExecOne: Integer): WordBool;
    function DeleteReport(lngAdviceKey: Integer): WordBool;
    function DeleteElectrocardioReport(lngAdviceKey: Integer): WordBool;
    function SendReport(lngAdviceKey: Integer; const strReportView: WideString; 
                        const strReportAdvice: WideString; const strReportDoctor: WideString; 
                        const strAuditingDoctor: WideString): WordBool;
    function SendElectrocardioReport(lngAdviceKey: Integer; const strReportTitle: WideString; 
                                     const strReportImgFiles: WideString; 
                                     const strReportResult: WideString; 
                                     const strReportAdvice: WideString; 
                                     const strReportDoctor: WideString; 
                                     const strAuditingDoctor: WideString): WordBool;
    function SendReportImages(lngAdviceKey: Integer; const strImgFiles: WideString): WordBool;
    function SendReportAffix(lngAdviceKey: Integer; const strAffixFiles: WideString): WordBool;
    property DefaultInterface: _clsPacsInterface read GetDefaultInterface;
    property Tables: TCusTable read Get_Tables;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TclsPacsInterfaceProperties read GetServerProperties;
{$ENDIF}
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// OLE Server Properties Proxy Class
// Server Object    : TclsPacsInterface
// (This object is used by the IDE's Property Inspector to allow editing
//  of the properties of this server)
// *********************************************************************//
 TclsPacsInterfaceProperties = class(TPersistent)
  private
    FServer:    TclsPacsInterface;
    function    GetDefaultInterface: _clsPacsInterface;
    constructor Create(AServer: TclsPacsInterface);
  protected
    function Get_Tables: TCusTable;
  public
    property DefaultInterface: _clsPacsInterface read GetDefaultInterface;
  published
  end;
{$ENDIF}


procedure Register;

resourcestring
  dtlServerPage = 'ActiveX';

  dtlOcxPage = 'ActiveX';

implementation

uses ComObj;

class function CoclsPacsInterface.Create: _clsPacsInterface;
begin
  Result := CreateComObject(CLASS_clsPacsInterface) as _clsPacsInterface;
end;

class function CoclsPacsInterface.CreateRemote(const MachineName: string): _clsPacsInterface;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_clsPacsInterface) as _clsPacsInterface;
end;

procedure TclsPacsInterface.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{E0A958B7-9A2E-449C-B4B0-EE18BF1812ED}';
    IntfIID:   '{19AA3D5A-CD0C-4ACF-9BD4-43512F9A34C3}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TclsPacsInterface.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as _clsPacsInterface;
  end;
end;

procedure TclsPacsInterface.ConnectTo(svrIntf: _clsPacsInterface);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TclsPacsInterface.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TclsPacsInterface.GetDefaultInterface: _clsPacsInterface;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call ''Connect'' or ''ConnectTo'' before this operation');
  Result := FIntf;
end;

constructor TclsPacsInterface.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TclsPacsInterfaceProperties.Create(Self);
{$ENDIF}
end;

destructor TclsPacsInterface.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function TclsPacsInterface.GetServerProperties: TclsPacsInterfaceProperties;
begin
  Result := FProps;
end;
{$ENDIF}

function TclsPacsInterface.Get_Tables: TCusTable;
begin
    Result := DefaultInterface.Tables;
end;

function TclsPacsInterface.GetErrShowType: TErrorShowType;
begin
  Result := DefaultInterface.GetErrShowType;
end;

function TclsPacsInterface.GetSplitChar: WideString;
begin
  Result := DefaultInterface.GetSplitChar;
end;

function TclsPacsInterface.GetSysNo: Integer;
begin
  Result := DefaultInterface.GetSysNo;
end;

function TclsPacsInterface.GetSysOwner: WideString;
begin
  Result := DefaultInterface.GetSysOwner;
end;

function TclsPacsInterface.GetNullValue: WideString;
begin
  Result := DefaultInterface.GetNullValue;
end;

function TclsPacsInterface.InitInterface(const strServerName: WideString; 
                                         const strUserName: WideString; 
                                         const strUserPwd: WideString; SysNo: Integer; 
                                         const SysOwner: WideString; const NullValue: WideString; 
                                         const SplitChar: WideString; errType: TErrorShowType): WordBool;
begin
  Result := DefaultInterface.InitInterface(strServerName, strUserName, strUserPwd, SysNo, SysOwner, 
                                           NullValue, SplitChar, errType);
end;

function TclsPacsInterface.GetLastError: WideString;
begin
  Result := DefaultInterface.GetLastError;
end;

function TclsPacsInterface.BeginTrans: WordBool;
begin
  Result := DefaultInterface.BeginTrans;
end;

function TclsPacsInterface.CommitTrans: WordBool;
begin
  Result := DefaultInterface.CommitTrans;
end;

function TclsPacsInterface.RollbackTrans: WordBool;
begin
  Result := DefaultInterface.RollbackTrans;
end;

function TclsPacsInterface.ExecutePacsProcedure(const strProcedureName: WideString): WordBool;
begin
  Result := DefaultInterface.ExecutePacsProcedure(strProcedureName);
end;

function TclsPacsInterface.GetPacsCursor(const strProcedureName: WideString; 
                                         const strFilterValue: WideString; 
                                         var blnIsNoParamerer: WordBool): WordBool;
begin
  Result := DefaultInterface.GetPacsCursor(strProcedureName, strFilterValue, blnIsNoParamerer);
end;

function TclsPacsInterface.GetAdoData(const strProcedureName: WideString; 
                                      const strFilterValue: WideString; 
                                      var blnIsNoParamerer: WordBool): _Recordset;
begin
  Result := DefaultInterface.GetAdoData(strProcedureName, strFilterValue, blnIsNoParamerer);
end;

function TclsPacsInterface.GetRecordValueByColumnName(var strDatas: PSafeArray; 
                                                      var strColumns: PSafeArray; 
                                                      lngRecordIndex: Integer; 
                                                      const strCurColumn: WideString): WideString;
begin
  Result := DefaultInterface.GetRecordValueByColumnName(strDatas, strColumns, lngRecordIndex, 
                                                        strCurColumn);
end;

function TclsPacsInterface.GetRecordValueByColumnIndex(var strDatas: PSafeArray; 
                                                       lngRecordIndex: Integer; 
                                                       lngColumnIndex: Integer): WideString;
begin
  Result := DefaultInterface.GetRecordValueByColumnIndex(strDatas, lngRecordIndex, lngColumnIndex);
end;

function TclsPacsInterface.GetCurValueByColumnName(lngRecordIndex: Integer; 
                                                   const strCurColumn: WideString): WideString;
begin
  Result := DefaultInterface.GetCurValueByColumnName(lngRecordIndex, strCurColumn);
end;

function TclsPacsInterface.GetCurRecordCount: Integer;
begin
  Result := DefaultInterface.GetCurRecordCount;
end;

function TclsPacsInterface.GetCurColumnCount: Integer;
begin
  Result := DefaultInterface.GetCurColumnCount;
end;

function TclsPacsInterface.GetCurColumnIndex(const strCurColumn: WideString): Integer;
begin
  Result := DefaultInterface.GetCurColumnIndex(strCurColumn);
end;

function TclsPacsInterface.GetCurValueByColumnIndex(lngRecordIndex: Integer; lngColumnIndex: Integer): WideString;
begin
  Result := DefaultInterface.GetCurValueByColumnIndex(lngRecordIndex, lngColumnIndex);
end;

function TclsPacsInterface.GetCurRecordData(lngRecordIndex: Integer): WideString;
begin
  Result := DefaultInterface.GetCurRecordData(lngRecordIndex);
end;

function TclsPacsInterface.GetRecordData(var strDatas: PSafeArray; lngRecordIndex: Integer): WideString;
begin
  Result := DefaultInterface.GetRecordData(strDatas, lngRecordIndex);
end;

function TclsPacsInterface.GetRecordCount(var strDatas: PSafeArray): Integer;
begin
  Result := DefaultInterface.GetRecordCount(strDatas);
end;

function TclsPacsInterface.GetColumnCount(var strColumns: PSafeArray): Integer;
begin
  Result := DefaultInterface.GetColumnCount(strColumns);
end;

function TclsPacsInterface.GetColumnIndex(var strColumns: PSafeArray; const strCurColumn: WideString): Smallint;
begin
  Result := DefaultInterface.GetColumnIndex(strColumns, strCurColumn);
end;

function TclsPacsInterface.GetColumnName(var strColumns: PSafeArray; columnIndex: Smallint): WideString;
begin
  Result := DefaultInterface.GetColumnName(strColumns, columnIndex);
end;

function TclsPacsInterface.GetCurColumnName(columnIndex: Smallint): WideString;
begin
  Result := DefaultInterface.GetCurColumnName(columnIndex);
end;

function TclsPacsInterface.GetDeptItems(const strFilter: WideString): WordBool;
begin
  Result := DefaultInterface.GetDeptItems(strFilter);
end;

function TclsPacsInterface.GetChargeTypes(const strFilter: WideString): WordBool;
begin
  Result := DefaultInterface.GetChargeTypes(strFilter);
end;

function TclsPacsInterface.GetPacsItems(const strFilter: WideString): WordBool;
begin
  Result := DefaultInterface.GetPacsItems(strFilter);
end;

function TclsPacsInterface.GetAdviceItems(lngAdviceKey: Integer): WordBool;
begin
  Result := DefaultInterface.GetAdviceItems(lngAdviceKey);
end;

function TclsPacsInterface.GetAdviceFees(lngAdviceKey: Integer): WordBool;
begin
  Result := DefaultInterface.GetAdviceFees(lngAdviceKey);
end;

function TclsPacsInterface.GetPatientInfo(const strQueryKey: WideString; 
                                          lngWhereType: TPatientWhereType): WordBool;
begin
  Result := DefaultInterface.GetPatientInfo(strQueryKey, lngWhereType);
end;

function TclsPacsInterface.GetRequestInfo(const strQueryKey: WideString; 
                                          lngWhereType: TRequestWhereType): WordBool;
begin
  Result := DefaultInterface.GetRequestInfo(strQueryKey, lngWhereType);
end;

function TclsPacsInterface.GetRequestExecuteStatus(lngAdviceKey: Integer): Integer;
begin
  Result := DefaultInterface.GetRequestExecuteStatus(lngAdviceKey);
end;

function TclsPacsInterface.GetRequestAdviceStatus(lngAdviceKey: Integer): Integer;
begin
  Result := DefaultInterface.GetRequestAdviceStatus(lngAdviceKey);
end;

function TclsPacsInterface.GetRequestExeProcedureStatus(lngAdviceKey: Integer): Integer;
begin
  Result := DefaultInterface.GetRequestExeProcedureStatus(lngAdviceKey);
end;

function TclsPacsInterface.CancelRequest(lngAdviceKey: Integer; lngExecOne: Integer): WordBool;
begin
  Result := DefaultInterface.CancelRequest(lngAdviceKey, lngExecOne);
end;

function TclsPacsInterface.RecevieRequest(lngAdviceKey: Integer; const strExeRoom: WideString; 
                                          lngStudyNo: Integer; const strDevice: WideString; 
                                          lngHeight: Integer; lngWeight: Integer; 
                                          const strStudyDoc: WideString; StrExeDate: TDateTime; 
                                          const strExeDes: WideString; lngExecOne: Integer): WordBool;
begin
  Result := DefaultInterface.RecevieRequest(lngAdviceKey, strExeRoom, lngStudyNo, strDevice, 
                                            lngHeight, lngWeight, strStudyDoc, StrExeDate, 
                                            strExeDes, lngExecOne);
end;

function TclsPacsInterface.ModifyRequest(lngAdviceKey: Integer; const strExeRoom: WideString; 
                                         lngStudyNo: Integer; const strDevice: WideString; 
                                         lngHeight: Integer; lngWeight: Integer; 
                                         const strStudyDoc: WideString; StrExeDate: TDateTime; 
                                         const strExeDes: WideString; lngExecOne: Integer): WordBool;
begin
  Result := DefaultInterface.ModifyRequest(lngAdviceKey, strExeRoom, lngStudyNo, strDevice, 
                                           lngHeight, lngWeight, strStudyDoc, StrExeDate, 
                                           strExeDes, lngExecOne);
end;

function TclsPacsInterface.DeleteReport(lngAdviceKey: Integer): WordBool;
begin
  Result := DefaultInterface.DeleteReport(lngAdviceKey);
end;

function TclsPacsInterface.DeleteElectrocardioReport(lngAdviceKey: Integer): WordBool;
begin
  Result := DefaultInterface.DeleteElectrocardioReport(lngAdviceKey);
end;

function TclsPacsInterface.SendReport(lngAdviceKey: Integer; const strReportView: WideString; 
                                      const strReportAdvice: WideString; 
                                      const strReportDoctor: WideString; 
                                      const strAuditingDoctor: WideString): WordBool;
begin
  Result := DefaultInterface.SendReport(lngAdviceKey, strReportView, strReportAdvice, 
                                        strReportDoctor, strAuditingDoctor);
end;

function TclsPacsInterface.SendElectrocardioReport(lngAdviceKey: Integer; 
                                                   const strReportTitle: WideString; 
                                                   const strReportImgFiles: WideString; 
                                                   const strReportResult: WideString; 
                                                   const strReportAdvice: WideString; 
                                                   const strReportDoctor: WideString; 
                                                   const strAuditingDoctor: WideString): WordBool;
begin
  Result := DefaultInterface.SendElectrocardioReport(lngAdviceKey, strReportTitle, 
                                                     strReportImgFiles, strReportResult, 
                                                     strReportAdvice, strReportDoctor, 
                                                     strAuditingDoctor);
end;

function TclsPacsInterface.SendReportImages(lngAdviceKey: Integer; const strImgFiles: WideString): WordBool;
begin
  Result := DefaultInterface.SendReportImages(lngAdviceKey, strImgFiles);
end;

function TclsPacsInterface.SendReportAffix(lngAdviceKey: Integer; const strAffixFiles: WideString): WordBool;
begin
  Result := DefaultInterface.SendReportAffix(lngAdviceKey, strAffixFiles);
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TclsPacsInterfaceProperties.Create(AServer: TclsPacsInterface);
begin
  inherited Create;
  FServer := AServer;
end;

function TclsPacsInterfaceProperties.GetDefaultInterface: _clsPacsInterface;
begin
  Result := FServer.DefaultInterface;
end;

function TclsPacsInterfaceProperties.Get_Tables: TCusTable;
begin
    Result := DefaultInterface.Tables;
end;

{$ENDIF}

procedure Register;
begin
  RegisterComponents(dtlServerPage, [TclsPacsInterface]);
end;

end.
