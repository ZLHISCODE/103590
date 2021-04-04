{*******************************************************************************
*
*  ��Ƶ��ȫ����ʾ��װ
*  �����ˣ�TJH
*  ������ǰ��2009-11-26
*  ˵����ȫ������ʾʹ�������ַ�ʽ�����֧��VMR������VMR��ȫ����ʾ��
*        ����ʹ�ÿؼ������ȫ����ʾ��ʽ������Ҫע����ǣ�DSPACK�����֧�ֶ�����
*        ��ȫ����ʾ��ʼ�ս�ȫ����ʾ�ڵ�һ����ʾ���У���˶�DSPACK��Ԫ����ʾ����
*        �����޸ģ��Ա㴫��ָ������ʾ���ھ����Ȼ���ٽ�����Ƶ��ȫ����ʾ��
*
*******************************************************************************}
unit FullScreenWindow;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DirectShow9, ZLDSVideoProcess_TLB, DSPack, AppEvnts, ExtCtrls,
  StdCtrls;

type
  TfrmFullScreen = class(TForm)
    timeFocus: TTimer;
    Label1: TLabel;
    timeRefreshVideo: TTimer;
    procedure FormHide(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure timeFocusTimer(Sender: TObject);
    procedure FormDeactivate(Sender: TObject);
    procedure timeRefreshVideoTimer(Sender: TObject);
  private
    { Private declarations }

    //������ʾ����λ��
    class procedure SetFullWindowPos(fullForm: TForm; const monitorIndex: Integer);

    //��ʹ��VMR��ȫ����ʾ����
    class procedure FullScreenWithNotUserVmr(fullForm: TfrmFullScreen;
      videoWindow: TVideoWindow; ShowModel: TShowModel; const monitorIndex: Integer);

    //ʹ��VMR��ȫ����ʾ����
    class procedure FullScreenWithUserVmr(fullForm: TfrmFullScreen;
      videoWindow: TVideoWindow; ShowModel: TShowModel; const monitorIndex: Integer);

  public
    { Public declarations }
    FullScreenVideoWindow: TVideoWindow; //��Ƶ��ʾ����
    FullScreenMonitorIndex: Integer;
    FullScreenShowModel: TShowModel;

    //ȫ����ʾ
    class function ShowFullScreen(vw: TVideoWindow;
                                  const showModel: TShowModel;
                                  const monitorIndex: Integer
                                  ): WideString;

    //�˳�ȫ��
    class function QuitFullScreen(): WideString;
    //ȡ���Ƿ��Ѿ�����ȫ����ʾ
    class function GetFullScreenState(): Boolean;

  end;

implementation

{$R *.dfm}

uses
  CaptureDebug, Math;

var
  frmFullScreen: TfrmFullScreen = nil; //����ȫ������ʾ����
  IsExitApp: Boolean = false;          //�ж��Ƿ�ʼ�˳�Ӧ�ó�������ǣ���ִ�д��ڵ�����(HIDE)�¼�

{ TfrmFullScreen }


class function TfrmFullScreen.QuitFullScreen(): WideString;
begin
  try
    Result := '';

    if not Assigned(frmFullScreen) then Exit;

    if frmFullScreen.Showing then frmFullScreen.Hide;
  except
    on e: Exception do begin
      Result := e.Message;
      
      TDebug.DebugMsg('TfrmFullScreen', 'QuitFullScreen', e.Message);
    end;
  end;
end;


class function TfrmFullScreen.ShowFullScreen(
  vw: TVideoWindow;
  const showModel: TShowModel;
  const monitorIndex: Integer): WideString;
var
  vmrWindowLessCtrl9: IVMRWindowlessControl9;
  hr: HRESULT;
begin
  try
    Result := '';

    if not Assigned(frmFullScreen) then
      frmFullScreen := TfrmFullScreen.Create(Application{nil});

    frmFullScreen.OnClick := vw.OnClick;
    frmFullScreen.OnDblClick := vw.OnDblClick;
    frmFullScreen.OnKeyDown := vw.OnKeyDown;
    frmFullScreen.OnKeyPress := vw.OnKeyPress;
    frmFullScreen.OnKeyUp := vw.OnKeyUp;
    frmFullScreen.OnMouseDown := vw.OnMouseDown;
    frmFullScreen.OnMouseMove := vw.OnMouseMove;
    frmFullScreen.OnMouseUp := vw.OnMouseUp;
    frmFullScreen.OnPaint := vw.OnPaint;

    frmFullScreen.Color := vw.Color;
    frmFullScreen.FullScreenVideoWindow := vw;
    frmFullScreen.FullScreenMonitorIndex := monitorIndex;
    frmFullScreen.FullScreenShowModel := showModel;

    frmFullScreen.Show;

    SetFullWindowPos(frmFullScreen, monitorIndex);

    hr := (vw as IBaseFilter).QueryInterface(IID_IVMRWindowlessControl9, vmrWindowLessCtrl9);
    if Failed(hr) then begin
      //��ʹ��VMR��ȫ������
      FullScreenWithNotUserVmr(frmFullScreen, vw, showModel, monitorIndex);
    end else begin
      //ʹ��VMR��ʾȫ��
      vmrWindowLessCtrl9 := nil;
      FullScreenWithUserVmr(frmFullScreen, vw, showModel, monitorIndex);
    end;
  except
    on e: Exception do begin
      Result := e.Message;
      if Assigned(frmFullScreen) then FreeAndNil(frmFullScreen);
    end;
  end;
end;


procedure TfrmFullScreen.FormHide(Sender: TObject);
var
  vmrWindowLessCtr9: IVMRWindowlessControl9;
  hr: HRESULT;
  videoRect, localRect: TRect;
  localWidth, localHeight, arWidth, arHeight: Integer;
begin
  try
    if IsExitApp then Exit;
    if (FullScreenVideoWindow = nil) then Exit;
    
    hr := (FullScreenVideoWindow as IBaseFilter).QueryInterface(IID_IVMRWindowlessControl9, vmrWindowLessCtr9);
    if Failed(hr) then begin
      FullScreenVideoWindow.FullScreen := False;
      Exit;
    end;

    try
      if (vmrWindowLessCtr9 = nil) then Exit;
      
      vmrWindowLessCtr9.GetNativeVideoSize(localWidth, localHeight, arWidth, arHeight);

      videoRect := Rect(0, 0, FullScreenVideoWindow.Width, FullScreenVideoWindow.Height);
      localRect := Rect(0, 0, localWidth, localHeight);

      vmrWindowLessCtr9.SetVideoClippingWindow(FullScreenVideoWindow.Handle);
      vmrWindowLessCtr9.SetVideoPosition(@localRect, @videoRect);

      FullScreenVideoWindow.VMROptions.KeepAspectRatio := False;

      FullScreenVideoWindow := nil;
    finally
      vmrWindowLessCtr9 := nil;
    end;
  except
    on e: Exception do begin
      TDebug.DebugMsg('TfrmFullScreen', 'FormHide', e.Message);
    end;
  end;
end;

procedure TfrmFullScreen.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  frmFullScreen := nil;
  Action := caFree;
end;


class function TfrmFullScreen.GetFullScreenState(): Boolean;
begin
  try
    Result := False;

    if not Assigned(frmFullScreen) then Exit;

    Result := frmFullScreen.Showing;
  except
    Result := False;
  end;
end;

class procedure TfrmFullScreen.SetFullWindowPos(fullForm: TForm; const monitorIndex: Integer);
begin
  fullForm.SetBounds(Screen.Monitors[monitorIndex].Left,
                     Screen.Monitors[monitorIndex].Top,
                     Screen.Monitors[monitorIndex].Width,
                     Screen.Monitors[monitorIndex].Height);
end;

class procedure TfrmFullScreen.FullScreenWithNotUserVmr(fullForm: TfrmFullScreen;
  videoWindow: TVideoWindow; ShowModel: TShowModel; const monitorIndex: Integer);
begin
  videoWindow.IsStreach := ShowModel = smStretch;
  videoWindow.MonitorIndex := MonitorIndex;
  videoWindow.FullHandle := frmFullScreen.Handle;
  videoWindow.FullScreen := True;
end;

class procedure TfrmFullScreen.FullScreenWithUserVmr(fullForm: TfrmFullScreen;
  videoWindow: TVideoWindow; ShowModel: TShowModel; const monitorIndex: Integer);
var
  hr: HRESULT;
  screenRect, videoRect: TRect;
  vmrWindowLessCtr9: IVMRWindowlessControl9;
  localWidth, localHeight, arWidth, arHeight: Integer;
begin
  try
    hr := (videoWindow as IBaseFilter).QueryInterface(IID_IVMRWindowlessControl9, vmrWindowLessCtr9);
    if Failed(hr) then begin
      fullForm.Label1.Visible := True;
      Exit;
    end;

    fullForm.Label1.Visible := False;
    
    case showModel of
      smNormal, smFit, smAutoFitCut, smWindAutoFit: begin
        VideoWindow.VMROptions.KeepAspectRatio := True;
      end;
      smStretch: begin
        VideoWindow.VMROptions.KeepAspectRatio := False;
      end;
    end;

    vmrWindowLessCtr9.GetNativeVideoSize(localWidth, localHeight, arWidth, arHeight);

    screenRect := Rect(0, 0, Screen.Monitors[MonitorIndex].Width, Screen.Monitors[MonitorIndex].Height);
    videoRect := Rect(0, 0, localWidth, localHeight);

    vmrWindowLessCtr9.SetVideoClippingWindow(fullForm.Handle);
    vmrWindowLessCtr9.SetVideoPosition(@videoRect, @screenRect);
  finally
    vmrWindowLessCtr9 := nil;
  end;
end;

procedure TfrmFullScreen.timeFocusTimer(Sender: TObject);
begin
  //�����ø�����ĳ�����ʹ��ȫ��ʱ�������ȫ������Ƶ�����е������갴����
  //��ô��ڽ�ʧȥ���㣬�����Ҫʹ�ü�������ȡʧȥ�Ľ��㣬
  //��Deactivate�¼����߽ػ���Ϣ����������Ч���ػ�ô��ڵĽ���
  if Self.Visible then begin
    SetFocus;

    if Focused then begin
      timeFocus.Enabled := False;
    end;  
  end;
end;

procedure TfrmFullScreen.FormDeactivate(Sender: TObject);
begin
  timeFocus.Enabled := True;
end;

procedure TfrmFullScreen.timeRefreshVideoTimer(Sender: TObject);
begin
  if Self.Visible then begin
    //ʹ�ü�����ˢ��VMRȫ����������Ϊ����������С��ȫ��������ʾʱ����Ƶֻ��ʾ��һС������
    FullScreenWithUserVmr(Self,
                          FullScreenVideoWindow,
                          FullScreenShowModel,
                          FullScreenMonitorIndex);
  end;
end;

initialization

finalization
  if Assigned(frmFullScreen) then begin
    IsExitApp := True; //������ﲻ���ó�����˳�����ΪTRUE����ʹ�ø������Ӧ�ó����˳�ʱ���ڸô��ڵ�HIDE�¼��н��������쳣��
    FreeAndNil(frmFullScreen);
  end;  

end.
