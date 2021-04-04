library ZLDSVideoProcess;

uses
  ComServ,
  Windows,
  ZLDSVideoProcess_TLB in 'ZLDSVideoProcess_TLB.pas',
  DSCaptureImpl1 in 'DSCaptureImpl1.pas' {DSCapture: TActiveForm} {DSCapture: CoClass},
  CapParameterCfg in 'CapParameterCfg.pas' {frmCapParameterCfg},
  VideoProcessDefine in 'VideoProcessDefine.pas',
  SizerControl in 'Component\SizerControl.pas',
  GraphicProcess in 'GraphicProcess.pas',
  DirectShow9Ex in 'DirectShow9Ex.pas',
  CaptureDebug in 'CaptureDebug.pas',
  DSPlayImpl1 in 'DSPlayImpl1.pas' {DSPlay: TActiveForm} {DSPlay: CoClass},
  VideoInfWindow in 'VideoInfWindow.pas' {frmVideoInf},
  FullScreenWindow in 'FullScreenWindow.pas' {frmFullScreen},
  DSCapParameterConfigObj in 'DSCapParameterConfigObj.pas' {DSCapParameterEnum: CoClass},
  GIFImage in 'GifImage.pas',
  TMCIAudioImpl1 in 'TMCIAudioImpl1.pas' {TMCIAudio: CoClass},
  TMCIPlayerImpl1 in 'TMCIPlayerImpl1.pas' {TMCIPlayer: CoClass},
  AxCtrls in 'axctrls.pas';

{$E ocx}

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer;

{$R *.TLB}

{$R *.RES}

procedure DllEntry(reason: DWord);
begin
  case reason of
    DLL_PROCESS_ATTACH: begin
      //...
    end;
    DLL_THREAD_ATTACH: begin
    end;
    DLL_THREAD_DETACH: begin
    end;
    DLL_PROCESS_DETACH: begin
      //...      
    end;
  end;
end;

begin
  //DllProc := @DllEntry;
  //DllEntry(DLL_PROCESS_ATTACH);
end.
