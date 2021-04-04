{*******************************************************************************
*
*  对一些过程函数进行修改或者扩展(和视频采集相关的函数过程)
*  创建人：TJH
*  创建日前：2009-11-26
*
*******************************************************************************}
unit DirectShow9Ex;

interface

uses
  Windows, Sysutils, Classes, Graphics, DirectShow9, DSUtil, stdCtrls, ZLDSVideoProcess_TLB;


const
  IID_IPropertyBag          : TGUID = '{55272A00-42CB-11CE-8135-00AA004BB851}';
  IID_ISpecifyPropertyPages : TGUID = '{B196B28B-BAB4-101A-B69C-00AA00341D07}';
  IID_IPersistStream        : TGUID = '{00000109-0000-0000-C000-000000000046}';
  IID_IMoniker              : TGUID = '{0000000F-0000-0000-C000-000000000046}';  

Type
  //属性页类型  (在DSUTIL单元中已经定义)
  {TPropertyPage = (
    ppDefault,       // Simple property page.
    ppVFWCapDisplay, // Capture Video source dialog box.
    ppVFWCapFormat,  // Capture Video format dialog box.
    ppVFWCapSource,  // Capture Video source dialog box.
    ppVFWCompConfig, // Compress Configure dialog box.
    ppVFWCompAbout   // Compress About Dialog box.
  );
  }

  //对Directshow9单元的扩展类
  TDS9Ex = class(TObject)
  private
    {private description}
  public
    //查找指定类型的未连接的PIN
    class function FindUnConnectedPin(Filter: IBaseFilter;
      dir: PIN_DIRECTION; out pin: IPin): HResult;

    //查找指定类型的连接的PIN
    class function FindConnectedPin(Filter: IBaseFilter;
      dir: PIN_DIRECTION; out pin: IPin): HResult;

    //查找上层filter的输出连接PIN
    class function FindLastPin(Filter: IBaseFilter; out lastPin: IPin): HRESULT;

    //查找顶层的filter
    class function FindTopFilter(curFilter: IBaseFilter;
      out topFilter, nextFilter: IBaseFilter):  HRESULT;  

    //判断指定的PIN是否已经连接   2010-06-23暂时未测试
    class function IsConnectedPin(pin: IPin; dir: PIN_DIRECTION): HRESULT;

    //连接filter
    class function ConnectFilters(Graph: IGraphBuilder;
      Src: IBaseFilter; Dst: IBaseFilter; IsSmartTee: Boolean): HResult;

    //配置分辨率和色深
    class function ConfigCaptureScale(Filter: IBaseFilter;
      Width, Height, Bits: Integer): HResult;

    //显示PIN属性页
    class function ShowPinPropertyPage(const titleName: WideString;
      parent: THandle; Pin: IPin): HRESULT;

    class function ShowPinPropertyPage1(const titleName: WideString; parent: THandle;
      iCapGraphicBuilder2: ICaptureGraphBuilder2; capSourceFilter: IBaseFilter; out videoSize: TVideoSize): HRESULT;

    //显示VIDEO CROSSBAR属性页  
    class function ShowVideoCrossbarPropertyPage(const titleName: WideString; parent: THandle;
      iCapGraphicBuilder2: ICaptureGraphBuilder2; capSourceFilter: IBaseFilter): HRESULT;      

    //显示VFW指定属性页
    class function ShowFilterPropertyPage(const titleName: WideString;
      parent: THandle; Filter: IBaseFilter; PropertyPage: TPropertyPage = ppDefault): HRESULT;


    //显示编码器属性
    class function ShowEncoderFilterProperty(
      const encoderName: WideString; parent: THandle): HRESULT;

    //显示采集源FILTER属性  
    class function ShowCaptureFilterProperty(
      const capDeviceName: WideString; parent: THandle): HRESULT;

    //根据Clsid创建FILTER到GraphBuilder中
    class function AddFilterToGraphBuilder(Graph: IGraphBuilder; const clsid: TGUID;
      out Filter: IBaseFilter; wName: WideString): HResult;

    //根据采集设备名称创建FILTER
    class function CreateFilterByDeviceName(const clsIdDeviceClass: TGUID;
      const deviceName: WideString; out filter: IBaseFilter): HRESULT;

    //取得指定设备目录的所有设备名称  
    class function GetDeviceNames(const clsidDeviceClass: TGUID;
      var deviceNames: TStringList): HRESULT;

    //根据设备类型目录返回指定filter
    class function CreateFilterByDeviceCategory(const clsidDeviceClass: TGUID;
      const iidFilter: TGuid; out filter : IBaseFilter) : HRESULT;

    //根据FILTER名称查找FILTER并返回
    class function FindFilterByName(Category: TGUID;
      out Filter: IBaseFilter; wName: WideString): HResult;

    //设置SampleGrabber的媒体类型
    class function SetSampleGrabberMediaType(Grabber : ISampleGrabber;
      majortype : TGUID; iBitDepth : integer) : HRESULT;

    //获取图像
    class function GrabBitmap(Grabber : ISampleGrabber;
      bitmap : TBitmap) : HRESULT;

    //取得pin的类型名称
    class function GetCrossbarPinTypeName(const pinType : integer): WideString;

    //判断是否vfw的设备
    class function IsVfwDevice(const deviceName: WideString): Boolean;

    //取得媒体类型名称
    class function GetMediaGuidName(guid: TGUID): WideString;
    
    //取得时间格式名称
    class function GetTimeFormatName(guid: TGUID): WideString;

    //取得当前filter的连接PIN，并取得连接pin相关的连接属性
    class procedure GetConnectedPin(filter: IBaseFilter; out connectionPin, destPin: Ipin; out pmt: TAMMediaType);
  end;

implementation

uses
  Activex, CaptureDebug, Dialogs;


//释放媒体指针数据  
procedure FreeMediaType(mt: PAMMediaType);
begin
  if (mt^.cbFormat <> 0) then
  begin
    CoTaskMemFree(mt^.pbFormat);
    // Strictly unnecessary but tidier
    mt^.cbFormat := 0;
    mt^.pbFormat := nil;
  end;
  if (mt^.pUnk <> nil) then mt^.pUnk := nil;
end;


//删除媒体类型指针  
procedure DeleteMediaType(pmt: PAMMediaType);
begin
    // allow nil pointers for coding simplicity
  if (pmt = nil) then exit;
  FreeMediaType(pmt);
  CoTaskMemFree(pmt);
end;

//查找指定的未连接PIN
class function TDS9Ex.FindUnConnectedPin(Filter: IBaseFilter; dir: PIN_DIRECTION; out pin: IPin): HResult;
var
  enum: IEnumPins;
  pinTmp: IPin;
  dirTmp: PIN_DIRECTION;
begin
  Result := E_FAIL;

  if not Assigned(Filter) then Exit;

  Result := Filter.EnumPins(enum);
  if Result <> S_OK then Exit;

  try
    while enum.Next(1, pin, nil) = S_OK do begin
      Result := pin.QueryDirection(dirTmp);

      if SUCCEEDED(Result) and (dirTmp = dir) then begin
        pinTmp := nil;
        Result := pin.ConnectedTo(pinTmp);

        if FAILED(Result) then begin
          pinTmp := nil;
          Result := S_OK;
          Exit;
        end;

      end;
    end;

    Result := E_FAIL;
  finally
    enum := nil;
  end;
end;

//连接各个filter pin
class function TDS9Ex.ConnectFilters(Graph: IGraphBuilder;
  Src: IBaseFilter; Dst: IBaseFilter; IsSmartTee: Boolean): HResult;
var
  pinOut, pinIn: IPin;
  pins: TPinList;
  i: Integer;
  found: Boolean;
begin
  Result := E_FAIL;

  if not (Assigned(Graph) and Assigned(Src) and Assigned(Dst)) then Exit;

  if not IsSmartTee then begin
    Result := FindUnConnectedPin(Src, PINDIR_OUTPUT, pinOut);
    if FAILED(Result) then Exit;
  end else begin
    pins := TPinList.Create(Src);
    try
      found := False;
      for i := 0 to pins.Count - 1 do begin
        if pins.PinInfo[i].dir <> PINDIR_OUTPUT then Continue;
        //Capture Pin在输出端子中序号为0，Preview Pin在输出端子中序号为1
        if found then pinOut := pins[i];

        found := True;
      end;
    finally
      FreeAndNil(pins);
    end;
  end;

  Result := FindUnConnectedPin(Dst, PINDIR_INPUT, pinIn);
  if FAILED(Result) then Exit;

  Result := Graph.Connect(pinOut, pinIn);
end;

class function TDS9Ex.ConfigCaptureScale(Filter: IBaseFilter; Width, Height, Bits: Integer): HResult;
var
  pin: IPin;
  amStreamConfig: IAMStreamConfig;
  pmt: PAMMediaType;
  pvih: PVideoInfoHeader;
begin
  Result := E_FAIL;

  if not Assigned(Filter) then Exit;

  //如果该filter有多个输出pin时，如何进行设置呢？？（2010-12-15 无物理采集卡测试，不过每个采集源filter应该至少有一个capture pin）
  //在调用该方法之前，实际上已经断开了所有filter之间的连接，因此查找未连接PIN时，实际上找到的就是第一个pin
  //如果未断开filter之间的连接，则在设置分辨率时，应该查找已经连接的pin
  Result := FindUnConnectedPin(Filter, PINDIR_OUTPUT, pin);
  if FAILED(Result) then Exit;

  try
    Result := pin.QueryInterface(IID_IAMStreamConfig, amStreamConfig);
    if FAILED(Result) then Exit;
             
    Result := amStreamConfig.GetFormat(pmt);   //取得默认视频格式
    if FAILED(Result) then Exit;

    try
      pvih := pmt.pbFormat;
      pvih.bmiHeader.biBitCount := Bits;
      pvih.bmiHeader.biWidth := Width;
      pvih.bmiHeader.biHeight := Height;
      pvih.bmiHeader.biSize := (Width * Height * Bits) div 8;
      //pmt.subtype := MEDIASUBTYPE_RGB24; //如果没有指定，则自动匹配颜色空间

      Result := amStreamConfig.SetFormat(pmt^);

      DeleteMediaType(pmt);
    finally
      amStreamConfig := nil;
    end;  
  finally
    pin := nil;
  end;
end;

class function TDS9Ex.ShowPinPropertyPage(const titleName: WideString; parent: THandle; Pin: IPin): HRESULT;
var
  SpecifyPropertyPages: ISpecifyPropertyPages;
  CAGUID :TCAGUID;
  PinInfo: TPinInfo;
begin
  result := E_FAIL;
  
  if Pin = nil then exit;
  result := Pin.QueryInterface(IID_ISpecifyPropertyPages, SpecifyPropertyPages);
  
  if result <> S_OK then exit;
  
  result := SpecifyPropertyPages.GetPages(CAGUID);
  
  if result <> S_OK then begin
    SpecifyPropertyPages := nil;
    Exit;
  end;
  
  result := Pin.QueryPinInfo(PinInfo);
  if result <> S_OK then exit;
  try
    result := OleCreatePropertyFrame(parent, 0, 0, PWideChar(titleName), 1, @Pin,
                                     CAGUID.cElems, CAGUID.pElems, 0, 0, nil);
  finally
    CoTaskMemFree(CAGUID.pElems);
    PinInfo.pFilter := nil;
    
    SpecifyPropertyPages := nil;
  end;
end;


class function TDS9Ex.ShowFilterPropertyPage(const titleName: WideString;
  parent: THandle; Filter: IBaseFilter; PropertyPage: TPropertyPage = ppDefault): HRESULT;
var
  SpecifyPropertyPages : ISpecifyPropertyPages;
  CaptureDialog : IAMVfwCaptureDialogs;
  CompressDialog: IAMVfwCompressDialogs;
  CAGUID  :TCAGUID;
  FilterInfo: TFilterInfo;
  Code: Integer;
begin
  result := E_FAIL;
  code := 0;
  if Filter = nil then exit;

  ZeroMemory(@FilterInfo, SizeOf(TFilterInfo));

  case PropertyPage of
    ppVFWCapDisplay: code := VfwCaptureDialog_Display;
    ppVFWCapFormat : code := VfwCaptureDialog_Format;
    ppVFWCapSource : code := VfwCaptureDialog_Source;
    ppVFWCompConfig: code := VfwCompressDialog_Config;
    ppVFWCompAbout : code := VfwCompressDialog_About;
  end;

  case PropertyPage of
    ppDefault:
      begin
        result := Filter.QueryInterface(IID_ISpecifyPropertyPages, SpecifyPropertyPages);
        if result <> S_OK then exit;
        result := SpecifyPropertyPages.GetPages(CAGUID);
        if result <> S_OK then exit;
        result := Filter.QueryFilterInfo(FilterInfo);
        if result = S_OK then
        begin
          result := OleCreatePropertyFrame(parent, 0, 0, PWideChar(titleName), 1, @Filter, CAGUID.cElems, CAGUID.pElems, 0, 0, nil );
          FilterInfo.pGraph := nil;
        end;
        if Assigned(CAGUID.pElems) then CoTaskMemFree(CAGUID.pElems);
        SpecifyPropertyPages := nil;
      end;
    ppVFWCapDisplay..ppVFWCapSource:
      begin
        result := Filter.QueryInterface(IID_IAMVfwCaptureDialogs,CaptureDialog);
        if (result <> S_OK) then exit;
        result := CaptureDialog.HasDialog(code);
        if result <> S_OK then exit;
        result := CaptureDialog.ShowDialog(code,parent);
        CaptureDialog := nil;
      end;
    ppVFWCompConfig..ppVFWCompAbout:
      begin
        result := Filter.QueryInterface(IID_IAMVfwCompressDialogs, CompressDialog);
        if (result <> S_OK) then exit;
        case PropertyPage of
          ppVFWCompConfig: result := CompressDialog.ShowDialog(VfwCompressDialog_QueryConfig, 0);
          ppVFWCompAbout : result := CompressDialog.ShowDialog(VfwCompressDialog_QueryAbout, 0);
        end;
        if result = S_OK then result := CompressDialog.ShowDialog(code,parent);
        CompressDialog := nil;
      end;
  end;
end;

class function TDS9Ex.AddFilterToGraphBuilder(Graph: IGraphBuilder; const clsid: TGUID;
  out Filter: IBaseFilter; wName: WideString): HResult;
begin
  Result := E_FAIL;

  if not Assigned(Graph) then Exit;

  Result := CoCreateInstance(clsid, nil, CLSCTX_INPROC_SERVER, IID_IBaseFilter, Filter);
  if Succeeded(Result) then begin
    Result := Graph.AddFilter(Filter, PWideChar(wName));
  end else begin
    Filter := nil;
  end;
end;


class function TDS9Ex.FindFilterByName(Category: TGUID;
  out Filter: IBaseFilter; wName: WideString): HResult;
var
  EnumDev: ICreateDevEnum;
  EnumMonik: IEnumMoniker;
  Moniker: IMoniker;
  bFound: Boolean;
  PropBag: IPropertyBag;
  vName: OleVariant;
begin
  Result := CoCreateInstance(CLSID_SystemDeviceEnum, nil, CLSCTX_INPROC_SERVER, IID_ICreateDevEnum, EnumDev);
  if FAILED(Result) then Exit;

  try
    Result := EnumDev.CreateClassEnumerator(Category, EnumMonik, 0);
    if FAILED(Result) then Exit;

    if not Assigned(EnumMonik) then begin
      Result := E_FAIL;
      Exit;
    end;

    while (EnumMonik.Next(1, Moniker, nil) = S_OK) do
    begin
      Result := Moniker.BindToStorage(nil, nil, IID_IPropertyBag, PropBag);
      if FAILED(Result) then continue;

      Result := PropBag.Read('FriendlyName', vName, nil);
      if FAILED(Result) then continue;

      if wName <> '' then
        bFound := SUCCEEDED(Result) and (wName = vName)
      else
        bFound := True;

      if bFound then begin
        Result := Moniker.BindToObject(nil, nil, IID_IBaseFilter, Filter);
        if FAILED(Result) then
          continue
        else
          Exit;
      end;
    end;
  finally
    EnumDev := nil;
  end;
end;

class function TDS9Ex.SetSampleGrabberMediaType(Grabber : ISampleGrabber;
  majortype : TGUID; iBitDepth : integer) : HRESULT;
var
  AMMediaType : PAMMediaType;
begin
  New(AMMediaType);
  ZeroMemory(AMMediaType, SizeOf(TAMMediaType));
  AMMediaType.lSampleSize := 1;
  AMMediaType.bFixedSizeSamples := TRUE;

  AMMediaType.majortype := majortype;
  case iBitDepth of
    8 : AMMediaType.subtype := MEDIASUBTYPE_RGB8;
    16: AMMediaType.subtype := {MEDIASUBTYPE_RGB24}MEDIASUBTYPE_RGB555;//注：因16位有BUG，暂时先把16位改为24位;
    24: AMMediaType.subtype := MEDIASUBTYPE_RGB24;
    32: AMMediaType.subtype := MEDIASUBTYPE_RGB32;
  end;

  Result := Grabber.SetMediaType(AMMediaType^);
  Dispose(AMMediaType);
end;


class function TDS9Ex.GrabBitmap(Grabber : ISampleGrabber;
  bitmap: TBitmap) : HRESULT;
var
  hr          : HRESULT;

  BIHeaderPtr : PBitmapInfoHeader;
  MediaType   : TAMMediaType;
  BitmapHandle: HBitmap;
  DIBPtr      : Pointer;
  DIBSize     : LongInt;
  BufferLen   : Longint;
  function GetDIBLineSize(BitCount, Width: Integer): Integer;
  begin
    if BitCount = 15 then
      BitCount := 16;
    Result := ((BitCount * Width + 31) div 32) * 4;
  end;
begin
  Result := E_FAIL;
  if not Assigned(bitmap) then Exit;

  hr := Grabber.GetConnectedMediaType(MediaType);
  if Failed(hr) then Exit;

  try
    if IsEqualGUID(MediaType.majortype, MEDIATYPE_Video) then
    begin
      BIHeaderPtr := nil;
      if IsEqualGUID(MediaType.formattype, FORMAT_VideoInfo) then
        begin
          if MediaType.cbFormat = SizeOf(TVideoInfoHeader) then  // check size
            BIHeaderPtr := @(PVideoInfoHeader(MediaType.pbFormat)^.bmiHeader);
        end
      else if IsEqualGUID(MediaType.formattype, FORMAT_VideoInfo2) then
        begin
          if MediaType.cbFormat = SizeOf(TVideoInfoHeader2) then  // check size
            BIHeaderPtr := @(PVideoInfoHeader2(MediaType.pbFormat)^.bmiHeader);
        end;

      // check, whether format is supported by SampleGrabber
      if not Assigned(BIHeaderPtr) then Exit;
      BitmapHandle := CreateDIBSection(0, PBitmapInfo(BIHeaderPtr)^,
                                       DIB_RGB_COLORS, DIBPtr, 0, 0);
      if BitmapHandle <> 0 then
      begin
        try
          if DIBPtr = nil then Exit;

          // get DIB size
          DIBSize := BIHeaderPtr^.biSizeImage;
          if DIBSize = 0 then
            with BIHeaderPtr^ do
              DIBSize := GetDIBLineSize(biBitCount, biWidth) * biHeight * biPlanes;

          // copy DIB
          // get buffer size
          BufferLen := 0;
          hr := Grabber.GetCurrentBuffer(BufferLen, nil);
          if Failed(hr) or (BufferLen <= 0) then Exit;

          // copy buffer to DIB
          if BufferLen > DIBSize then  // copy Min(BufferLen, DIBSize)
            BufferLen := DIBSize;
          hr := Grabber.GetCurrentBuffer(BufferLen, DIBPtr);
          if Failed(hr) then Exit;

          bitmap.Handle := BitmapHandle;
          Result := S_OK;
        finally
          if bitmap.Handle <> BitmapHandle then  // preserve for any changes in Graphics.pas
            DeleteObject(BitmapHandle);
        end;
      end;
    end;
  finally
    FreeMediaType(@MediaType);
  end;
end;


class function TDS9Ex.CreateFilterByDeviceCategory(const clsidDeviceClass: TGUID;
  const iidFilter: TGuid; out filter : IBaseFilter) : HRESULT;
var
  sysDevEnum: ICreateDevEnum;
  EnumMoniker  : IEnumMoniker;
  Moniker      : IMoniker;
  tmpFilter    : IBaseFilter;
begin
  filter := nil;
  
  Result := CoCreateInstance(CLSID_SystemDeviceEnum, nil, CLSCTX_INPROC_SERVER, IID_ICreateDevEnum, sysDevEnum);
  if not Succeeded(Result) then Exit;

  try
    Result := sysDevEnum.CreateClassEnumerator(clsidDeviceClass, EnumMoniker, 0);
    if Failed(Result) then Exit;

    if EnumMoniker = nil then begin
      Result := E_FAIL;
      Exit;
    end;

    while EnumMoniker.Next(1, Moniker, nil) = S_OK do begin
      Result := Moniker.BindToObject(nil, nil, IID_IBaseFilter, tmpFilter);

      if Failed(Result) then begin
        Moniker := nil;
        Continue;
      end;

      Result := tmpFilter.QueryInterface(iidFilter, Filter);
      if Succeeded(Result) then begin
        Moniker := nil;
        Result := S_OK;
        Exit;
      end;
    end;
  finally
    sysDevEnum := nil;
  end;
end;


class function TDS9Ex.GetCrossbarPinTypeName(
  const pinType: integer): WideString;
begin
  case pinType of
    {PhysConn_Video_Tuner}
    PhysConn_Video_Tuner: Result := 'Video Tuner';
    {PhysConn_Video_Composite}
    PhysConn_Video_Composite: Result := 'Video Composite';
    {PhysConn_Video_SVideo}
    PhysConn_Video_SVideo: Result	:= 'Video SVideo';
    {PhysConn_Video_RGB}
    PhysConn_Video_RGB: Result := 'Video RGB';
    {PhysConn_Video_YRYBY}
    PhysConn_Video_YRYBY: Result := 'Video YRYBY';
    {PhysConn_Video_SerialDigital}
    PhysConn_Video_SerialDigital: Result := 'Video SerialDigital';
    {PhysConn_Video_ParallelDigital}
    PhysConn_Video_ParallelDigital: Result := 'Video ParallelDigital';
    {PhysConn_Video_SCSI}
    PhysConn_Video_SCSI: Result := 'Video SCSI';
    {PhysConn_Video_AUX}
    PhysConn_Video_AUX: Result := 'Video AUX';
    {PhysConn_Video_1394}
    PhysConn_Video_1394: Result := 'Video 1394';
    {PhysConn_Video_USB}
    PhysConn_Video_USB: Result := 'Video USB';
    {PhysConn_Video_VideoDecoder}
    PhysConn_Video_VideoDecoder: Result := 'Video VideoDecoder';
    {PhysConn_Video_VideoEncoder}
    PhysConn_Video_VideoEncoder: Result := 'Video VideoEncoder';
    {PhysConn_Video_SCART}
    PhysConn_Video_SCART: Result := 'Video SCART';
    {PhysConn_Video_Black}
    PhysConn_Video_Black: Result := 'Video Black';


    {PhysConn_Audio_Tuner}
    PhysConn_Audio_Tuner: Result := 'Audio Tuner';
    {PhysConn_Audio_Line}
    PhysConn_Audio_Line: Result := 'Audio Line';
    {PhysConn_Audio_Mic}
    PhysConn_Audio_Mic: Result := 'Audio Mic';
    {PhysConn_Audio_AESDigital}
    PhysConn_Audio_AESDigital: Result := 'Audio AESDigital';
    {PhysConn_Audio_SPDIFDigital}
    PhysConn_Audio_SPDIFDigital: Result := 'Audio SPDIFDigital';
    {PhysConn_Audio_SCSI}
    PhysConn_Audio_SCSI: Result := 'Audio SCSI';
    {PhysConn_Audio_AUX}
    PhysConn_Audio_AUX: Result := 'Audio AUX';
    {PhysConn_Audio_1394}
    PhysConn_Audio_1394: Result := 'Audio 1394';
    {PhysConn_Audio_USB}
    PhysConn_Audio_USB: Result := 'Audio USB';
    {PhysConn_Audio_AudioDecoder}
    PhysConn_Audio_AudioDecoder: Result := 'Audio AudioDecoder';
  end;
end;

class function TDS9Ex.CreateFilterByDeviceName(const clsIdDeviceClass: TGUID;
  const deviceName: WideString; out filter: IBaseFilter): HRESULT;
var
  capEnum: ICreateDevEnum;
  EnumMoniker: IEnumMoniker;
  Moniker: IMoniker;
  propertyBag: IPropertyBag;
  vatFriendlyName: OleVariant;
  strFriendlyName: WideString;
begin
  Filter := nil;

  Result := CocreateInstance(CLSID_SystemDeviceEnum, nil, CLSCTX_INPROC_SERVER, IID_ICreateDevEnum, capEnum);
  if Failed(Result) then Exit;

  try
    Result := capEnum.CreateClassEnumerator(clsIdDeviceClass, EnumMoniker, 0);
    if Failed(Result) then Exit;

    if not Assigned(EnumMoniker) then begin
      Result := E_FAIL;
      Exit;
    end;

    while EnumMoniker.Next(1, Moniker, nil) = S_OK do begin
      Result := Moniker.BindToStorage(nil, nil, IID_IPropertyBag, propertyBag);

      if Failed(Result) then begin
        Moniker := nil;
        Continue;
      end;

      //取得设备的友好显示名称
      Result := propertyBag.Read('FriendlyName', vatFriendlyName, nil);
      if Failed(Result) then begin
        Moniker := nil;
        propertyBag := nil;
        Continue;
      end;

      strfriendlyName := UpperCase(vatFriendlyName);
      VariantClear(vatFriendlyName);

      //判断当前设备是否为指定的设备
      if strFriendlyName <> UpperCase(deviceName) then begin
        Moniker := nil;
        propertyBag := nil;
        Continue;
      end;

      Result := Moniker.BindToObject(nil, nil, IID_IBaseFilter, Filter);
      if Failed(Result) then begin
        Moniker := nil;
        propertyBag := nil;
        Continue;
      end;

      Moniker := nil;
      propertyBag := nil;

      Exit;
    end;
  finally
    capEnum := nil;
  end;
end;


class function TDS9Ex.IsVfwDevice(const deviceName: WideString): Boolean;
var
  hr: HRESULT;
  sourceFilter: IBaseFilter;
  vfwCfg: IAMVfwCaptureDialogs;
begin
  hr := CreateFilterByDeviceName(CLSID_VideoInputDeviceCategory, deviceName, sourceFilter);
  if hr <> S_OK then begin
    Result := false;
    Exit;
  end;
  
  try
    hr := sourceFilter.QueryInterface(IID_IAMVfwCaptureDialogs, vfwCfg);
    Result := Succeeded(hr);
  finally
    sourceFilter := nil;
    vfwCfg := nil;
  end;
end;

class function TDS9Ex.GetMediaGuidName(guid: TGUID): WideString;
var
  guidStr: WideString;
begin
  guidStr := GUIDToString(guid);
  if guidStr = GUIDToString(MEDIATYPE_NULL) then begin
    Result := 'NULL';
    Exit;
  end;

  //----------------------------------------------------
  
  if guidStr = '{30355844-0000-0010-8000-00AA00389B71}' then begin
    Result := 'Divx 5.2.1';
    Exit;
  end;

  if guidStr = '{64697663-0000-0010-8000-00AA00389B71}' then begin
    Result := 'Cinepak Codec By Radius';
    Exit;
  end;

  if guidStr = '{33564944-0000-0010-8000-00AA00389B71}' then begin
    Result := 'Divx MPEG-4';
    Exit;
  end;

  if guidStr = '{34363248-0000-0010-8000-00AA00389B71}' then begin
    Result := 'H264';
    Exit;
  end;

  if guidStr = '{55594648-0000-0010-8000-00AA00389B71}' then begin
    Result := 'Huffyuv v2.1.1';
    Exit;
  end;

  if guidStr = '{3447504D-0000-0010-8000-00AA00389B71}' then begin
    Result := 'Microsoft MPEG-4 Video Codec V1';
    Exit;
  end;

  if guidStr = '{3334504D-0000-0010-8000-00AA00389B71}' then begin
    Result := 'Microsoft MPEG-4 Video Codec V3';
    Exit;
  end;

  if guidStr = '{4D415243-0000-0010-8000-00AA00389B71}' then begin
    Result := 'Microsoft Video 1';
    Exit;
  end;

  if guidStr = '{33564D57-0000-0010-8000-00AA00389B71}' then begin
    Result := 'Microsoft Windows Media Video 9';
    Exit;
  end;

  if guidStr = '{30365056-0000-0010-8000-00AA00389B71}' then begin
    Result := 'VP60';
    Exit;
  end;

  if guidStr = '{31365056-0000-0010-8000-00AA00389B71}' then begin
    Result := 'VP61';
    Exit;
  end;

  if guidStr = '{32365056-0000-0010-8000-00AA00389B71}' then begin
    Result := 'VP62';
    Exit;
  end;

  if guidStr = '{30375056-0000-0010-8000-00AA00389B71}' then begin
    Result := 'VP70';
    Exit;
  end;

  if guidStr = '{44495658-0000-0010-8000-00AA00389B71}' then begin
    Result := 'XVID MPEG-4';
    Exit;
  end;

  //----------------------------------------------------

  if guidStr = GUIDToString(MEDIATYPE_Video) then begin
    Result := 'Video';
    Exit;
  end;

  if guidStr = GUIDToString(MEDIATYPE_Audio) then begin
    Result := 'Audio';
    Exit;
  end;

  if guidStr = GUIDToString(MEDIATYPE_Text) then begin
    Result := 'Text';
    Exit;
  end;

  if guidStr = GUIDToString(MEDIATYPE_Midi) then begin
    Result := 'Midi';
    Exit;
  end;

  if guidStr = GUIDToString(MEDIATYPE_Stream) then begin
    Result := 'Stream';
    Exit;
  end;

  if guidStr = GUIDToString(MEDIATYPE_Interleaved) then begin
    Result := 'Interleaved';
    Exit;
  end;

  if guidStr = GUIDToString(MEDIATYPE_File) then begin
    Result := 'File';
    Exit;
  end;

  if guidStr = GUIDToString(MEDIATYPE_ScriptCommand) then begin
    Result := 'ScriptCommand';
    Exit;
  end;

  if guidStr = GUIDToString(MEDIATYPE_AUXLine21Data) then begin
    Result := 'AUXLine21Data';
    Exit;
  end;

  if guidStr = GUIDToString(MEDIATYPE_VBI) then begin
    Result := 'VBI';
    Exit;
  end;

  if guidStr = GUIDToString(MEDIATYPE_Timecode) then begin
    Result := 'Timecode';
    Exit;
  end;

  if guidStr = GUIDToString(MEDIATYPE_LMRT) then begin
    Result := 'LMRT';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIATYPE_URL_STREAM) then begin
    Result := 'URL_STREAM';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIATYPE_MPEG1SystemStream) then begin
    Result := 'MPEG1SystemStream';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIATYPE_AnalogAudio) then begin
    Result := 'AnalogAudio';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIATYPE_AnalogVideo) then begin
    Result := 'AnalogVideo';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIATYPE_MPEG2_PACK) then begin
    Result := 'MPEG2_PACK';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIATYPE_MPEG2_PES) then begin
    Result := 'MPEG2_PES';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIATYPE_CONTROL) then begin
    Result := 'CONTROL';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIATYPE_MPEG2_SECTIONS) then begin
    Result := 'MPEG2_SECTIONS';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIATYPE_DVD_ENCRYPTED_PACK) then begin
    Result := 'DVD_ENCRYPTED_PACK';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIATYPE_DVD_NAVIGATION) then begin
    Result := 'DVD_NAVIGATION';
    Exit;
  end;

  //sub type
  
  if guidStr = GUIDToString(MEDIASUBTYPE_MP42) then begin
    Result := 'MEDIASUBTYPE_MP42';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_DIVX) then begin
    Result := 'DIVX';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_VOXWARE) then begin
    Result := 'VOXWARE';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_Ogg) then begin
    Result := 'Ogg';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_Vorbis) then begin
    Result := 'Vorbis';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_NULL) then begin
    Result := 'NULL';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_None) then begin
    Result := 'None';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_CLPL) then begin
    Result := 'CLPL';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_YUYV) then begin
    Result := 'YUYV';
    Exit;
  end;

  if guidStr = GUIDToString(MEDIASUBTYPE_IYUV) then begin
    Result := 'IYUV';
    Exit;
  end;

  if guidStr = GUIDToString(MEDIASUBTYPE_YVU9) then begin
    Result := 'YVU9';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_Y411) then begin
    Result := 'Y411';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_Y41P) then begin
    Result := 'Y41P';
    Exit;
  end;

  if guidStr = GUIDToString(MEDIASUBTYPE_YUY2) then begin
    Result := 'YUY2';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_YVYU) then begin
    Result := 'YVYU';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_UYVY) then begin
    Result := 'UYVY';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_Y211) then begin
    Result := 'Y211';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_CLJR) then begin
    Result := 'CLJR';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_IF09) then begin
    Result := 'IF09';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_CPLA) then begin
    Result := 'CPLA';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_MJPG) then begin
    Result := 'MJPG';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_TVMJ) then begin
    Result := 'TVMJ';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_WAKE) then begin
    Result := 'WAKE';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_CFCC) then begin
    Result := 'CFCC';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_IJPG) then begin
    Result := 'IJPG';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_Plum) then begin
    Result := 'Plum';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_DVCS) then begin
    Result := 'DVCS';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_DVSD) then begin
    Result := 'DVSD';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_MDVF) then begin
    Result := 'MDVF';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_RGB1) then begin
    Result := 'RGB1';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_RGB4) then begin
    Result := 'RGB4';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_RGB8) then begin
    Result := 'RGB8';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_RGB565) then begin
    Result := 'RGB565';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_RGB555) then begin
    Result := 'RGB555';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_RGB24) then begin
    Result := 'RGB24';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_RGB32) then begin
    Result := 'RGB32';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_ARGB1555) then begin
    Result := 'ARGB1555';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_ARGB4444) then begin
    Result := 'ARGB4444';
    Exit;
  end;

  if guidStr = GUIDToString(MEDIASUBTYPE_ARGB32) then begin
    Result := 'ARGB32';
    Exit;
  end;

  if guidStr = GUIDToString(MEDIASUBTYPE_A2R10G10B10) then begin
    Result := 'A2R10G10B10';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_A2B10G10R10) then begin
    Result := 'A2B10G10R10';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_AYUV) then begin
    Result := 'AYUV';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_AI44) then begin
    Result := 'AI44';
    Exit;
  end;

  if guidStr = GUIDToString(MEDIASUBTYPE_IA44) then begin
    Result := 'IA44';
    Exit;
  end;

  if guidStr = GUIDToString(MEDIASUBTYPE_RGB32_D3D_DX7_RT) then begin
    Result := 'RGB32_D3D_DX7_RT';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_RGB16_D3D_DX7_RT) then begin
    Result := 'RGB16_D3D_DX7_RT';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_ARGB32_D3D_DX7_RT) then begin
    Result := 'ARGB32_D3D_DX7_RT';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_ARGB4444_D3D_DX7_RT) then begin
    Result := 'ARGB4444_D3D_DX7_RT';
    Exit;
  end;

  if guidStr = GUIDToString(MEDIASUBTYPE_ARGB1555_D3D_DX7_RT) then begin
    Result := 'ARGB1555_D3D_DX7_RT';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_RGB32_D3D_DX9_RT) then begin
    Result := 'RGB32_D3D_DX9_RT';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_RGB16_D3D_DX9_RT) then begin
    Result := 'RGB16_D3D_DX9_RT';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_ARGB32_D3D_DX9_RT) then begin
    Result := 'ARGB32_D3D_DX9_RT';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_ARGB4444_D3D_DX9_RT) then begin
    Result := 'ARGB4444_D3D_DX9_RT';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_ARGB1555_D3D_DX9_RT) then begin
    Result := 'ARGB1555_D3D_DX9_RT';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_YV12) then begin
    Result := 'YV12';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_NV12) then begin
    Result := 'NV12';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_IMC1) then begin
    Result := 'IMC1';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_IMC2) then begin
    Result := 'IMC2';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_IMC3) then begin
    Result := 'IMC3';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_IMC4) then begin
    Result := 'IMC4';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_S340) then begin
    Result := 'S340';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_S342) then begin
    Result := 'S342';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_Overlay) then begin
    Result := 'Overlay';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_MPEG1Packet) then begin
    Result := 'MPEG1Packet';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_MPEG1Payload) then begin
    Result := 'MPEG1Payload';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_MPEG1AudioPayload) then begin
    Result := 'MPEG1AudioPayload';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIATYPE_MPEG1SystemStream) then begin
    Result := 'MEDIATYPE_MPEG1SystemStream';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_MPEG1System) then begin
    Result := 'MPEG1System';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_MPEG1VideoCD) then begin
    Result := 'MPEG1VideoCD';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_MPEG1Video) then begin
    Result := 'MPEG1Video';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_MPEG1Audio) then begin
    Result := 'MPEG1Audio';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_Avi) then begin
    Result := 'Avi';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_Asf) then begin
    Result := 'Asf';
    Exit;
  end;

  if guidStr = GUIDToString(MEDIASUBTYPE_QTMovie) then begin
    Result := 'QTMovie';
    Exit;
  end;

  if guidStr = GUIDToString(MEDIASUBTYPE_QTRpza) then begin
    Result := 'QTRpza';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_QTSmc) then begin
    Result := 'QTSmc';
    Exit;
  end;
    
  if guidStr = GUIDToString(MEDIASUBTYPE_QTRle) then begin
    Result := 'QTRle';
    Exit;
  end;
    
  if guidStr = GUIDToString(MEDIASUBTYPE_QTJpeg) then begin
    Result := 'QTJpeg';
    Exit;
  end;
    
  if guidStr = GUIDToString(MEDIASUBTYPE_PCMAudio_Obsolete) then begin
    Result := 'PCMAudio_Obsolete';
    Exit;
  end;
    
  if guidStr = GUIDToString(MEDIASUBTYPE_PCM) then begin
    Result := 'PCM';
    Exit;
  end;
    
  if guidStr = GUIDToString(MEDIASUBTYPE_WAVE) then begin
    Result := 'WAVE';
    Exit;
  end;
    
  if guidStr = GUIDToString(MEDIASUBTYPE_AU) then begin
    Result := 'AU';
    Exit;
  end;
    
  if guidStr = GUIDToString(MEDIASUBTYPE_AIFF) then begin
    Result := 'AIFF';
    Exit;
  end;
    
  if guidStr = GUIDToString(MEDIASUBTYPE_dvsd_) then begin
    Result := 'dvsd_';
    Exit;
  end;
    
  if guidStr = GUIDToString(MEDIASUBTYPE_dvhd) then begin
    Result := 'dvhd';
    Exit;
  end;
    
  if guidStr = GUIDToString(MEDIASUBTYPE_dvsl) then begin
    Result := 'dvsl';
    Exit;
  end;
    
  if guidStr = GUIDToString(MEDIASUBTYPE_dv25) then begin
    Result := 'dv25';
    Exit;
  end;
    
  if guidStr = GUIDToString(MEDIASUBTYPE_dv50) then begin
    Result := 'dv50';
    Exit;
  end;
    
  if guidStr = GUIDToString(MEDIASUBTYPE_dvh1) then begin
    Result := 'dvh1';
    Exit;
  end;
    
  if guidStr = GUIDToString(MEDIASUBTYPE_Line21_BytePair) then begin
    Result := 'Line21_BytePair';
    Exit;
  end;
    
  if guidStr = GUIDToString(MEDIASUBTYPE_Line21_GOPPacket) then begin
    Result := 'Line21_GOPPacket';
    Exit;
  end;

  if guidStr = GUIDToString(MEDIASUBTYPE_Line21_VBIRawData) then begin
    Result := 'Line21_VBIRawData';
    Exit;
  end;
    
  if guidStr = GUIDToString(MEDIASUBTYPE_TELETEXT) then begin
    Result := 'TELETEXT';
    Exit;
  end;
    
  if guidStr = GUIDToString(MEDIASUBTYPE_WSS) then begin
    Result := 'WSS';
    Exit;
  end;

  if guidStr = GUIDToString(MEDIASUBTYPE_VPS) then begin
    Result := 'VPS';
    Exit;
  end;
            
  if guidStr = GUIDToString(MEDIASUBTYPE_DRM_Audio) then begin
    Result := 'DRM_Audio';
    Exit;
  end;
      
  if guidStr = GUIDToString(MEDIASUBTYPE_IEEE_FLOAT) then begin
    Result := 'IEEE_FLOAT';
    Exit;
  end;
      
  if guidStr = GUIDToString(MEDIASUBTYPE_DOLBY_AC3_SPDIF) then begin
    Result := 'DOLBY_AC3_SPDIF';
    Exit;
  end;
      
  if guidStr = GUIDToString(MEDIASUBTYPE_RAW_SPORT) then begin
    Result := 'RAW_SPORT';
    Exit;
  end;
      
  if guidStr = GUIDToString(MEDIASUBTYPE_SPDIF_TAG_241h) then begin
    Result := 'SPDIF_TAG_241h';
    Exit;
  end;
      
  if guidStr = GUIDToString(MEDIASUBTYPE_DssVideo) then begin
    Result := 'DssVideo';
    Exit;
  end;
      
  if guidStr = GUIDToString(MEDIASUBTYPE_DssAudio) then begin
    Result := 'DssAudio';
    Exit;
  end;
      
  if guidStr = GUIDToString(MEDIASUBTYPE_VPVideo) then begin
    Result := 'VPVideo';
    Exit;
  end;
      
  if guidStr = GUIDToString(MEDIASUBTYPE_VPVBI) then begin
    Result := 'VPVBI';
    Exit;
  end;

  if guidStr = GUIDToString(MEDIASUBTYPE_AnalogVideo_NTSC_M) then begin
    Result := 'AnalogVideo_NTSC_M';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_AnalogVideo_PAL_B) then begin
    Result := 'AnalogVideo_PAL_B';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_AnalogVideo_PAL_D) then begin
    Result := 'AnalogVideo_PAL_D';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_AnalogVideo_PAL_G) then begin
    Result := 'AnalogVideo_PAL_G';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_AnalogVideo_PAL_H) then begin
    Result := 'AnalogVideo_PAL_H';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_AnalogVideo_PAL_I) then begin
    Result := 'AnalogVideo_PAL_I';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_AnalogVideo_PAL_M) then begin
    Result := 'AnalogVideo_PAL_M';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_AnalogVideo_PAL_N) then begin
    Result := 'AnalogVideo_PAL_N';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_AnalogVideo_PAL_N_COMBO) then begin
    Result := 'AnalogVideo_PAL_N_COMBO';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_AnalogVideo_SECAM_B) then begin
    Result := 'AnalogVideo_SECAM_B';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_AnalogVideo_SECAM_D) then begin
    Result := 'AnalogVideo_SECAM_D';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_AnalogVideo_SECAM_G) then begin
    Result := 'AnalogVideo_SECAM_G';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_AnalogVideo_SECAM_H) then begin
    Result := 'AnalogVideo_SECAM_H';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_AnalogVideo_SECAM_K) then begin
    Result := 'AnalogVideo_SECAM_K';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_AnalogVideo_SECAM_K1) then begin
    Result := 'AnalogVideo_SECAM_K1';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_AnalogVideo_SECAM_L) then begin
    Result := 'AnalogVideo_SECAM_L';
    Exit;
  end;

  if guidStr = GUIDToString(MEDIASUBTYPE_MPEG2_PROGRAM) then begin
    Result := 'MPEG2_PROGRAM';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_MPEG2_TRANSPORT) then begin
    Result := 'MPEG2_TRANSPORT';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_MPEG2_AUDIO) then begin
    Result := 'MPEG2_AUDIO';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_DOLBY_AC3) then begin
    Result := 'DOLBY_AC3';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_DVD_SUBPICTURE) then begin
    Result := 'DVD_SUBPICTURE';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_DVD_LPCM_AUDIO) then begin
    Result := 'DVD_LPCM_AUDIO';
    Exit;
  end;

  if guidStr = GUIDToString(MEDIASUBTYPE_DTS) then begin
    Result := 'DTS';
    Exit;
  end;
    
  if guidStr = GUIDToString(MEDIASUBTYPE_SDDS) then begin
    Result := 'SDDS';
    Exit;
  end;

    
  if guidStr = GUIDToString(MEDIATYPE_DVD_ENCRYPTED_PACK) then begin
    Result := 'MEDIATYPE_DVD_ENCRYPTED_PACK';
    Exit;
  end;

    
  if guidStr = GUIDToString(MEDIATYPE_DVD_NAVIGATION) then begin
    Result := 'MEDIATYPE_DVD_NAVIGATION';
    Exit;
  end;

    
  if guidStr = GUIDToString(MEDIASUBTYPE_DVD_NAVIGATION_PCI) then begin
    Result := 'DVD_NAVIGATION_PCI';
    Exit;
  end;

    
  if guidStr = GUIDToString(MEDIASUBTYPE_DVD_NAVIGATION_DSI) then begin
    Result := 'DVD_NAVIGATION_DSI';
    Exit;
  end;
  
  if guidStr = GUIDToString(MEDIASUBTYPE_DVD_NAVIGATION_PROVIDER) then begin
    Result := 'DVD_NAVIGATION_PROVIDER';
    Exit;
  end;

  if UpperCase(guidStr) = UpperCase('{32564933-0000-0010-8000-00aa00389b71}') then begin
    Result := '3IV2';
    Exit;
  end;

  if UpperCase(guidStr) = UpperCase('{30355649-0000-0010-8000-00aa00389b71}') then begin
    Result := 'IV50';
    Exit;
  end;

  if UpperCase(guidStr) = UpperCase('{31345649-0000-0010-8000-00aa00389b71}') then begin
    Result := 'IV41';
    Exit;
  end;

  if UpperCase(guidStr) = UpperCase('{31435657-0000-0010-8000-00AA00389B71}') then begin
    Result := 'WVC1';
    Exit;
  end;
  

  //format type
  if guidStr = GUIDToString(FORMAT_VorbisFormat) then begin
    Result := 'VorbisFormat';
    Exit;
  end;

  if guidStr = GUIDToString(FORMAT_None) then begin
    Result := 'None';
    Exit;
  end;

  if guidStr = GUIDToString(FORMAT_VideoInfo) then begin
    Result := 'VideoInfo';
    Exit;
  end;
  
  if guidStr = GUIDToString(FORMAT_VideoInfo2) then begin
    Result := 'VideoInfo2';
    Exit;
  end;
  
  if guidStr = GUIDToString(FORMAT_WaveFormatEx) then begin
    Result := 'WaveFormatEx';
    Exit;
  end;
  
  if guidStr = GUIDToString(FORMAT_MPEGVideo) then begin
    Result := 'MPEGVideo';
    Exit;
  end;
  
  if guidStr = GUIDToString(FORMAT_MPEGStreams) then begin
    Result := 'MPEGStreams';
    Exit;
  end;
  
  if guidStr = GUIDToString(FORMAT_DvInfo) then begin
    Result := 'DvInfo';
    Exit;
  end;
  
  if guidStr = GUIDToString(FORMAT_AnalogVideo) then begin
    Result := 'AnalogVideo';
    Exit;
  end;
  
  if guidStr = GUIDToString(FORMAT_MPEG2_VIDEO) then begin
    Result := 'MPEG2_VIDEO';
    Exit;
  end;

  if guidStr = GUIDToString(FORMAT_MPEG2Video) then begin
    Result := 'MPEG2Video';
    Exit;
  end;
    
  if guidStr = GUIDToString(FORMAT_DolbyAC3) then begin
    Result := 'DolbyAC3';
    Exit;
  end;

  if guidStr = GUIDToString(FORMAT_MPEG2Audio) then begin
    Result := 'MPEG2Audio';
    Exit;
  end;
    
  if guidStr = GUIDToString(FORMAT_DVD_LPCMAudio) then begin
    Result := 'DVD_LPCMAudio';
    Exit;
  end;

  Result := GUIDToString(guid);
end;

class function TDS9Ex.GetTimeFormatName(guid: TGUID): WideString;
var
  curTimeFormat: WideString;
begin
  curTimeFormat := GUIDToString(guid);

  if curTimeFormat = GUIDToString(TIME_FORMAT_NONE) then begin
    Result := 'TIME_FORMAT_NONE';
    Exit;
  end;

  if curTimeFormat = GUIDToString(TIME_FORMAT_FRAME) then begin
    Result := 'TIME_FORMAT_FRAME';
    Exit;
  end;

  if curTimeFormat = GUIDToString(TIME_FORMAT_BYTE) then begin
    Result := 'TIME_FORMAT_BYTE';
    Exit;
  end;

  if curTimeFormat = GUIDToString(TIME_FORMAT_SAMPLE) then begin
    Result := 'TIME_FORMAT_SAMPLE';
    Exit;
  end;

  if curTimeFormat = GUIDToString(TIME_FORMAT_FIELD) then begin
    Result := 'TIME_FORMAT_FIELD';
    Exit;
  end;

  if curTimeFormat = GUIDToString(TIME_FORMAT_MEDIA_TIME) then begin
    Result := 'TIME_FORMAT_MEDIA_TIME';
    Exit;
  end;

  Result := GUIDToString(guid);
end;

class function TDS9Ex.GetDeviceNames(const clsidDeviceClass: TGUID;
  var deviceNames: TStringList): HRESULT;
var
  capEnum: ICreateDevEnum;
  EnumMoniker: IEnumMoniker;
  Moniker: IMoniker;
  propertyBag: IPropertyBag;
  vatFriendlyName: OleVariant;
begin
  Result := CocreateInstance(CLSID_SystemDeviceEnum, nil, CLSCTX_INPROC_SERVER, IID_ICreateDevEnum, capEnum);
  if Failed(Result) then Exit;
  
  try
    Result := capEnum.CreateClassEnumerator(clsIdDeviceClass, EnumMoniker, 0);
    if Failed(Result) then Exit;

    if not Assigned(EnumMoniker) then begin
      Result := E_FAIL;
      Exit;
    end;

    while EnumMoniker.Next(1, Moniker, nil) = S_OK do begin
      Result := Moniker.BindToStorage(nil, nil, IID_IPropertyBag, propertyBag);
      if Failed(Result) then begin
        Moniker := nil;
        Continue;
      end;

      //取得设备的友好显示名称
      Result := propertyBag.Read('FriendlyName', vatFriendlyName, nil);
      if Failed(Result) then begin
        Moniker := nil;
        propertyBag := nil;
        Continue;
      end;

      deviceNames.Append(vatFriendlyName);

      VariantClear(vatFriendlyName);

      Moniker := nil;
      propertyBag := nil;
    end;

    Result := S_OK;
  finally
    capEnum := nil;
  end;
end;

class function TDS9Ex.ShowEncoderFilterProperty(
  const encoderName: WideString; parent: THandle): HRESULT;
var
  filterGraph: IFilterGraph;
  encoderFilter: IBaseFilter;
begin
  Result := CoCreateInstance(CLSID_FilterGraph, nil, CLSCTX_INPROC_SERVER, IID_IFilterGraph2, filterGraph);
  if Failed(Result) then Exit;

  try
    Result := CreateFilterByDeviceName(CLSID_VideoCompressorCategory, encoderName, encoderFilter);
    if Failed(Result) then Exit;

    try
      filterGraph.AddFilter(encoderFilter, 'FBC6923FC3584F5798EF6E573038BFEE');

      Result := ShowFilterPropertyPage('编码器', parent, encoderFilter, ppVFWCompConfig);
      
      if Failed(Result) then
        Result := ShowFilterPropertyPage('编码器', parent, encoderFilter);
    finally
      encoderFilter := nil;
    end;
  finally
    filterGraph := nil;
  end;
end;


class function TDS9Ex.ShowCaptureFilterProperty(
  const capDeviceName: WideString; parent: THandle): HRESULT;
var
  filterGraph: IFilterGraph;
  captureFilter: IBaseFilter;
begin
  Result := CoCreateInstance(CLSID_FilterGraph, nil, CLSCTX_INPROC_SERVER, IID_IFilterGraph2, filterGraph);
  if Failed(Result) then Exit;

  try
    Result := CreateFilterByDeviceName(CLSID_VideoInputDeviceCategory, capDeviceName, captureFilter);
    if Failed(Result) then Exit;

    try
      filterGraph.AddFilter(captureFilter, 'C92B19B81CBC4E4BB09F64FC1BB4C171');

      Result := ShowFilterPropertyPage('视频源', parent, captureFilter);
    finally
      captureFilter := nil;
    end;
  finally
    filterGraph := nil;
  end;
end;

class function TDS9Ex.IsConnectedPin(pin: IPin; dir: PIN_DIRECTION): HRESULT;
var
  pinTmp: IPin;
  dirTmp: PIN_DIRECTION;
begin
  Result := E_FAIL;

  if not Assigned(pin) then Exit;

  try
    Result := pin.QueryDirection(dirTmp);

    if SUCCEEDED(Result) and (dirTmp = dir) then begin
      Result := pin.ConnectedTo(pinTmp);

      if Succeeded(Result) then Result := S_OK;
    end;
  finally
    pinTmp := nil;
  end;
end;

class function TDS9Ex.ShowPinPropertyPage1(const titleName: WideString; parent: THandle;
  iCapGraphicBuilder2: ICaptureGraphBuilder2; capSourceFilter: IBaseFilter; out videoSize: TVideoSize): HRESULT;
var
  iAMStreamCfg: IAMStreamConfig;
  iPropertyPages: ISpecifyPropertyPages;
  cid: TCAGUID;
  //pmt: PAMMediaType;
  //pvih: PVideoInfoHeader;
  connectedPin: IPin;
  destPin: IPin;
  connectedPmt: TAMMediaType;
begin
  //该处理程序由amcap的源程序转换成的delphi代码
  Result := E_FAIL;

  if not Assigned(iCapGraphicBuilder2) then Exit;

  try
    //断开当前连接的采集端口
    TDS9Ex.GetConnectedPin(capSourceFilter, connectedPin, destPin, connectedPmt);
    if Assigned(connectedPin) then begin
      connectedPin.Disconnect();
    end;

    try

      Result := iCapGraphicBuilder2.FindInterface(@PIN_CATEGORY_CAPTURE, @MEDIATYPE_Interleaved, capSourceFilter, IID_IAMStreamConfig, iAMStreamCfg);

      if Result <> S_OK then
        Result := iCapGraphicBuilder2.FindInterface(@PIN_CATEGORY_CAPTURE, @MEDIATYPE_Video, capSourceFilter, IID_IAMStreamConfig, iAMStreamCfg);

      if Failed(Result) then exit;

      Result := iAMStreamCfg.QueryInterface(IID_ISpecifyPropertyPages, iPropertyPages);

      if Succeeded(Result) then begin
        Result := iPropertyPages.GetPages(cid);
        if Failed(Result) then exit;

        Result := OleCreatePropertyFrame(parent, 0, 0, PWideChar(titleName), 1, @iAMStreamCfg, cid.cElems, cid.pElems, 0, 0, nil);

        //取得设置的视频格式大小
        {iAMStreamCfg.GetFormat(pmt);
        try
          pvih := pmt.pbFormat;
        finally
          DeleteMediaType(pmt);
        end;//}
      end;
    finally
      if Assigned( connectedPin) then begin
        connectedPin.ReceiveConnection(destPin, connectedPmt);
        connectedPin := nil;
        destPin := nil;
      end;
    end;
  finally
    CoTaskMemFree(cid.pElems);
    iPropertyPages := nil;
    iAMStreamCfg := nil;
  end;
end;

class procedure TDS9Ex.GetConnectedPin(filter: IBaseFilter;
  out connectionPin, destPin: Ipin; out pmt: TAMMediaType);
var
  pinList: TPinList;
  i: Integer;
begin
  pinList := TPinList.Create(filter);
  try
    connectionPin := nil;
    destPin := nil;
  
    for i := 0 to pinList.Count - 1 do begin
      if pinList.PinInfo[i].dir <> PINDIR_OUTPUT then continue;
      if not pinList.Connected[i] then continue;

      connectionPin := pinList.Items[i];
      connectionPin.ConnectedTo(destPin);
      connectionPin.ConnectionMediaType(pmt);

      exit;
    end;
  finally
    FreeAndNil(pinList);
  end;
end;

class function TDS9Ex.ShowVideoCrossbarPropertyPage(const titleName: WideString; parent: THandle;
  iCapGraphicBuilder2: ICaptureGraphBuilder2; capSourceFilter: IBaseFilter): HRESULT;
var
  iAMCrossbarObj: IAMCrossbar;
  iPropertyPages: ISpecifyPropertyPages;
  cid: TCAGUID;
begin
  //该处理程序由amcap的源程序转换成的delphi代码
  Result := E_FAIL;

  if not Assigned(iCapGraphicBuilder2) then Exit;

  try
    Result := iCapGraphicBuilder2.FindInterface(@PIN_CATEGORY_CAPTURE, @MEDIATYPE_Interleaved, capSourceFilter, IID_IAMCrossbar, iAMCrossbarObj);

    if Result <> S_OK then
      Result := iCapGraphicBuilder2.FindInterface(@PIN_CATEGORY_CAPTURE, @MEDIATYPE_Video, capSourceFilter, IID_IAMCrossbar, iAMCrossbarObj);

    if Failed(Result) then exit;

    Result := iAMCrossbarObj.QueryInterface(IID_ISpecifyPropertyPages, iPropertyPages);

    if Succeeded(Result) then begin
      Result := iPropertyPages.GetPages(cid);
      if Failed(Result) then exit;

      Result := OleCreatePropertyFrame(parent, 0, 0, PWideChar(titleName), 1, @iAMCrossbarObj, cid.cElems, cid.pElems, 0, 0, nil);
    end;
  finally
    CoTaskMemFree(cid.pElems);
    iPropertyPages := nil;
    iAMCrossbarObj := nil;
  end;
end;

//查找已经连接的PIN
class function TDS9Ex.FindConnectedPin(Filter: IBaseFilter;
  dir: PIN_DIRECTION; out pin: IPin): HResult;
var
  enum: IEnumPins;
  pinTmp: IPin;
  dirTmp: PIN_DIRECTION;
begin
  Result := E_FAIL;

  if not Assigned(Filter) then Exit;

  Result := Filter.EnumPins(enum);
  if Result <> S_OK then Exit;

  try
    while enum.Next(1, pin, nil) = S_OK do begin
      Result := pin.QueryDirection(dirTmp);

      if SUCCEEDED(Result) and (dirTmp = dir) then begin
        pinTmp := nil;
        Result := pin.ConnectedTo(pinTmp);
                            
        if Succeeded(Result) then begin
          pinTmp := nil;
          Result := S_OK;
          
          Exit;
        end;
      end;
    end;

    Result := E_FAIL;
  finally
    enum := nil;
  end;
end;


class function TDS9Ex.FindTopFilter(curFilter: IBaseFilter; out topFilter,
  nextFilter: IBaseFilter): HRESULT;
var
  lastFilter: IBaseFilter;
  lastPin: IPin;
  pinInfo: _PinInfo;
begin
  topFilter := nil;
  nextFilter := nil;

  Result := E_FAIL;

  try

    lastFilter := curFilter;

    while Assigned(lastFilter) do begin
      Result := FindLastPin(curFilter, lastPin);
      if Failed(Result) then Exit;


      if not Assigned(lastPin) then Exit;


      Result := lastPin.QueryPinInfo(pinInfo);
      if Failed(Result) then Exit;

      nextFilter := lastFilter;

      lastFilter := pinInfo.pFilter;
      topFilter := lastFilter;
    end;
  finally
    lastPin := nil;
  end;
end;

class function TDS9Ex.FindLastPin(Filter: IBaseFilter;
  out lastPin: IPin): HRESULT;
var
  enum: IEnumPins;
  pinTmp: IPin;
  dirTmp: PIN_DIRECTION;
  pin: IPin;
begin
  lastPin := nil;
  Result := E_FAIL;

  if not Assigned(Filter) then Exit;

  Result := Filter.EnumPins(enum);
  if Result <> S_OK then Exit;

  try
    while enum.Next(1, pin, nil) = S_OK do begin
      Result := pin.QueryDirection(dirTmp);

      if SUCCEEDED(Result) and (dirTmp = PINDIR_INPUT) then begin
        pinTmp := nil;
        Result := pin.ConnectedTo(pinTmp);

        if Succeeded(Result) then begin
          lastPin := pinTmp;
          Result := S_OK;
          
          Exit;
        end;
      end;
    end;

    Result := E_FAIL;
  finally
    enum := nil;
    pin := nil;
  end;
end;

end.
