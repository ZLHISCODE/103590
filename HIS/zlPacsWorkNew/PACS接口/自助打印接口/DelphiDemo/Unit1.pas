unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, zlPacsPrtInterface_TLB, StdCtrls, ComCtrls, ExtCtrls;

type
  TForm1 = class(TForm)
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    Label1: TLabel;
    lvRequestInf: TListView;
    Button1: TButton;
    Edit1: TEdit;
    Label2: TLabel;
    GroupBox1: TGroupBox;
    edtOracleInstanceName: TEdit;
    Label3: TLabel;
    Label4: TLabel;
    edtOracleUserName: TEdit;
    Label5: TLabel;
    edtSysNum: TEdit;
    Label6: TLabel;
    edtDbOwner: TEdit;
    Label7: TLabel;
    edtOraclePwd: TEdit;
    Button2: TButton;
    lvPatientInf: TListView;
    butQueryPatientInf: TButton;
    cbxQueryType: TComboBox;
    Label9: TLabel;
    Label8: TLabel;
    cbxPatientQueryType: TComboBox;
    Label10: TLabel;
    edtPatientValue: TEdit;
    TabSheet3: TTabSheet;
    lvPacsStudyBodyPart: TListView;
    Button3: TButton;
    TabSheet4: TTabSheet;
    TabSheet5: TTabSheet;
    GroupBox2: TGroupBox;
    GroupBox3: TGroupBox;
    GroupBox4: TGroupBox;
    Button4: TButton;
    Label11: TLabel;
    memStudyView: TMemo;
    Label12: TLabel;
    memAdvice: TMemo;
    memReportImage: TMemo;
    Button5: TButton;
    Label13: TLabel;
    edtAdviceId: TEdit;
    memAffix: TMemo;
    Button6: TButton;
    OpenDialog1: TOpenDialog;
    Label14: TLabel;
    edtReportDoctor: TEdit;
    TabSheet6: TTabSheet;
    lvDeptItemInf: TListView;
    Button7: TButton;
    GroupBox5: TGroupBox;
    Label15: TLabel;
    edtExeRoom: TEdit;
    Label16: TLabel;
    edtStudyNo: TEdit;
    Label17: TLabel;
    edtDevice: TEdit;
    Label18: TLabel;
    edtHeight: TEdit;
    Label19: TLabel;
    edtWeight: TEdit;
    Label20: TLabel;
    edtStudyDoctor: TEdit;
    Label21: TLabel;
    dtpRequestDate: TDateTimePicker;
    Label22: TLabel;
    edtExeDes: TEdit;
    Label24: TLabel;
    lvAdviceFees: TListView;
    lvAdviceItems: TListView;
    GroupBox6: TGroupBox;
    Label23: TLabel;
    edtRecevieAdviceId: TEdit;
    butRecevieAdvice: TButton;
    Button8: TButton;
    Button9: TButton;
    GroupBox7: TGroupBox;
    Label25: TLabel;
    edtReceiveAdviceIDOne: TEdit;
    btnReceiveRequestOne: TButton;
    btnModifyReqestOne: TButton;
    btnCancelRequestOne: TButton;
    btnDeleteReport: TButton;
    TabSheet7: TTabSheet;
    GroupBox8: TGroupBox;
    Label26: TLabel;
    Label27: TLabel;
    memECGResult: TMemo;
    memECGAdvice: TMemo;
    GroupBox9: TGroupBox;
    memECGImage: TMemo;
    Label28: TLabel;
    edtECGAdviceId: TEdit;
    Label29: TLabel;
    edtECGReport: TEdit;
    Button11: TButton;
    Button12: TButton;
    Button10: TButton;
    Label30: TLabel;
    edtECGName: TEdit;
    Shape1: TShape;
    Label31: TLabel;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    Label32: TLabel;
    Button13: TButton;
    TabSheet8: TTabSheet;
    lvFinishedStudy: TListView;
    Label33: TLabel;
    txtPatientId: TEdit;
    Button14: TButton;
    Button15: TButton;
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure lvRequestInfClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure butQueryPatientInfClick(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure butRecevieAdviceClick(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
    procedure btnDeleteReportClick(Sender: TObject);
    procedure Button10Click(Sender: TObject);
    procedure Button11Click(Sender: TObject);
    procedure Button12Click(Sender: TObject);
    procedure Button13Click(Sender: TObject);
    procedure Button14Click(Sender: TObject);
    procedure Button15Click(Sender: TObject);
  private
    { Private declarations }
    ipacs: _clsPacsInterface;

    procedure ReadQueryData(var objPacsInterface: _clsPacsInterface; var objCurListView: TListView);
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

uses ActiveX, StrUtils;

{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
begin
  ipacs := CoclsPacsInterface.create;

  PageControl1.ActivePageIndex := 0;
end;

procedure TForm1.Button1Click(Sender: TObject);
//**********************************************************
//
//????????????????????????
//
//**********************************************************
var
  strError: WideString;
begin
  Screen.Cursor := crHourGlass;

  try
    //??????????????
    ipacs.GetRequestInfo(Edit1.Text, cbxQueryType.ItemIndex + 1);
    strError := ipacs.GetLastError;
    if strError <> '' then begin
      ShowMessage(strError);
      Exit;
    end;

    ReadQueryData(ipacs, lvRequestInf);

    lvAdviceItems.Items.Clear;
    lvAdviceFees.Items.Clear;
    if lvRequestInf.Items.Count > 0 then begin
      lvRequestInf.ItemIndex := 0;
      lvRequestInfClick(nil);
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TForm1.lvRequestInfClick(Sender: TObject);
var
  iRecordCount: Integer;
  lngAdviceId: Longint;
  i: Integer;
  strError: WideString;
begin
  if lvRequestInf.Selected = nil then exit;

  lngAdviceId := StrToInt(lvRequestInf.Selected.Caption);

  //????????????????
  ipacs.GetAdviceItems(lngAdviceId);
  strError := ipacs.GetLastError;
  if strError <> '' then begin
    ShowMessage(strError);
    Exit;
  end;
  ReadQueryData(ipacs,lvAdviceItems);

  //????????????
  ipacs.GetAdviceFees(lngAdviceId);
  strError := ipacs.GetLastError;
  if strError <> '' then
    begin
      ShowMessage(strError);
      Exit;
    end;
  ReadQueryData(ipacs, lvAdviceFees);
end;

procedure TForm1.Button2Click(Sender: TObject);
//**********************************************************
//
//??????ipacs????
//
//**********************************************************
var
  sErr: String;
begin
  //??????????????"#"??????????????????????????","????
  ipacs.InitInterface(edtOracleInstanceName.Text, edtOracleUserName.Text,
    edtOraclePwd.Text, StrToInt(edtSysNum.Text), edtDbOwner.Text, '', '~', estNoDisplay);


  sErr := ipacs.GetLastError;
  if sErr <> '' then
    ShowMessage(sErr)
  else
    ShowMessage(Pchar('????????????Oracle????????' + IfThen(edtOracleInstanceName.Text = '', 'Local', edtOracleInstanceName.Text)));
end;

procedure TForm1.butQueryPatientInfClick(Sender: TObject);
//**********************************************************
//
//??????????????????????
//
//**********************************************************
var
  strError: WideString;
begin
  Screen.Cursor := crHourGlass;

  try
    //??????????????????????
    ipacs.GetPatientInfo(edtPatientValue.Text, cbxPatientQueryType.ItemIndex + 1);
    strError := ipacs.GetLastError;
    if strError <> '' then begin
      ShowMessage(strError);
      Exit;
    end;

    ReadQueryData(ipacs, lvPatientInf);
  finally
    Screen.Cursor := crDefault;
  end;  
end;

procedure TForm1.Button3Click(Sender: TObject);
//**********************************************************
//
//??????????????????????
//
//**********************************************************
var
  strError: WideString;
begin
  screen.Cursor := crHourGlass;

  try
    //????pacs????????????
    ipacs.GetPacsItems('');
    strError := ipacs.GetLastError;
    if strError <> '' then begin
      ShowMessage(strError);
      Exit;
    end;  

    ReadQueryData(ipacs, lvPacsStudyBodyPart);
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TForm1.Button5Click(Sender: TObject);
var
  i: Integer;
begin
  OpenDialog1.Filter := '(*.bmp)|*.bmp|(*.jpg)|*.jpg|(*.*)|*.*';
  OpenDialog1.DefaultExt := '*.bmp';
  OpenDialog1.Options := [ofHideReadOnly,ofAllowMultiSelect,ofEnableSizing];

  if not OpenDialog1.Execute then Exit;

  for i := 0 to OpenDialog1.Files.Count - 1 do begin
    if Trim(memReportImage.Text) <> '' then
      memReportImage.Text := memReportImage.Text + ipacs.GetSplitChar;

    memReportImage.Text := memReportImage.Text + OpenDialog1.Files[i];
  end;
end;

procedure TForm1.Button6Click(Sender: TObject);
begin
  OpenDialog1.Filter := '(*.pdf)|*.pdf|(*.*)|*.*';
  OpenDialog1.DefaultExt := '*.pdf';
  OpenDialog1.Options := [ofHideReadOnly,ofEnableSizing];

  if OpenDialog1.Execute then memAffix.Text := OpenDialog1.FileName;
end;

procedure TForm1.Button4Click(Sender: TObject);
//**********************************************************
//
//????????
//
//**********************************************************
var
  sErr: WideString;
begin
  if Trim(edtAdviceId.Text) = '' then begin
    ShowMessage('????????????????ID??');
    Exit;
  end;

  Screen.Cursor := crHourGlass;
  try

    //????????????????????????,
    ipacs.DeleteReport(StrToInt(edtAdviceId.Text));

    //????????????????
    ipacs.SendReport(StrToInt(edtAdviceId.Text), memStudyView.Text, memAdvice.Text, edtReportDoctor.Text, '');
    sErr := ipacs.GetLastError;
    if Trim(sErr) <> '' then begin
      ShowMessage(sErr);
      Exit;
    end;

    //????????????(????????????????????SendReport)
    ipacs.SendReportImages(StrToInt(edtAdviceId.Text), memReportImage.Text);
    sErr := ipacs.GetLastError;
    if Trim(sErr) <> '' then begin
      ShowMessage(sErr);
    end;

    //????????????(????????????????????SendReport)
    ipacs.SendReportAffix(StrToInt(edtAdviceId.Text), memAffix.Text);
    sErr := ipacs.GetLastError;
    if Trim(sErr) <> '' then begin
      ShowMessage(sErr);
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TForm1.Button7Click(Sender: TObject);
//**********************************************************
//
//??????????????????????????
//
//**********************************************************
var
  strError: WideString; 
begin
  screen.Cursor := crHourGlass;

  try
    //????????????????
    ipacs.GetDeptItems('');
    strError := ipacs.GetLastError;
    if strError <> '' then begin
      ShowMessage(strError);
      Exit;
    end;

    ReadQueryData(ipacs, lvDeptItemInf);
  finally
    Screen.Cursor := crDefault;
  end;
end;





procedure TForm1.ReadQueryData(var objPacsInterface: _clsPacsInterface; var objCurListView: TListView);
//**********************************************************
//
//??objPacsInterface??????????????????????????
//
//**********************************************************
var
  iRecordCount: Integer;
  i, j: Integer;
  itemData: TListItem;

  iColumnCount: Integer;
  columnData: TListColumn;
begin
  objCurListView.Columns.BeginUpdate;
  try
    if objCurListView.Columns.Count <= 0 then begin
      objCurListView.Columns.Clear;
      //??????????????
      iColumnCount := objPacsInterface.GetCurColumnCount;
      for i := 0 to iColumnCount - 1 do begin
        columnData := objCurListView.Columns.Add;
        columnData.Caption := objPacsInterface.GetCurColumnName(i);
        columnData.Width := 80;
      end;
    end;  
  finally
    objCurListView.Columns.EndUpdate;
  end;



  iRecordCount := objPacsInterface.GetCurRecordCount;

  objCurListView.Items.BeginUpdate;
  try
    //????????????????
    objCurListView.Clear;
    for i := 0 to iRecordCount - 1 do begin
      itemData := objCurListView.Items.Add;
      itemData.Caption := objPacsInterface.GetCurValueByColumnName(i, objCurListView.Columns[0].Caption);

      for j := 1 to objCurListView.Columns.Count - 1 do begin
        itemData.SubItems.Add(objPacsInterface.GetCurValueByColumnName(i, objCurListView.Columns[j].Caption));
      end;
    end;
  finally
    objCurListView.Items.EndUpdate;
  end;

  if objCurListView.Items.Count > 0 then begin
    objCurListView.ItemIndex := 0;
  end;
end;

procedure TForm1.butRecevieAdviceClick(Sender: TObject);
//**********************************************************
//
//????????
//
//**********************************************************
var
  sErr: WideString;
  intExecOne : Integer;
  intAdviceID : Integer;
begin
  if Sender = butRecevieAdvice then
    begin
      if Trim(edtRecevieAdviceId.Text) = '' then begin
        ShowMessage('????????????????ID??');
        Exit;
      end;
      intAdviceID := StrToInt(edtRecevieAdviceId.Text);
      intExecOne := 0;
    end
  else
    begin
      if Trim(edtReceiveAdviceIDOne.Text) = '' then begin
        ShowMessage('????????????????ID??');
        Exit;
      end;
      intAdviceID := StrToInt(edtReceiveAdviceIDOne.Text);
      intExecOne := 1;
    end;

  Screen.Cursor := crHourGlass;
  try
    ipacs.RecevieRequest(intAdviceID, edtExeRoom.Text, StrToInt(edtStudyNo.Text), edtDevice.Text,
      StrToInt(edtHeight.Text), StrToInt(edtWeight.Text), edtStudyDoctor.Text, dtpRequestDate.Date, edtExeDes.Text,intExecOne);

    sErr := ipacs.GetLastError;
    if Trim(sErr) <> '' then begin
      ShowMessage(sErr);
    end else begin
      ShowMessage('??????????????????');
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TForm1.Button8Click(Sender: TObject);
//**********************************************************
//
//????????
//
//**********************************************************
var
  sErr: WideString;
  intExecOne : Integer;
  intAdviceID : int64;
begin
  if Sender = Button8 then
    begin
      if Trim(edtRecevieAdviceId.Text) = '' then begin
        ShowMessage('????????????????ID??');
        Exit;
      end;
      intAdviceID := StrToInt64(edtRecevieAdviceId.Text);
      intExecOne := 0;
    end
  else
    begin
       if Trim(edtReceiveAdviceIDOne.Text) = '' then begin
        ShowMessage('????????????????ID??');
        Exit;
      end;
      intAdviceID := StrToInt64(edtReceiveAdviceIDOne.Text);
      intExecOne := 1;
    end;


  Screen.Cursor := crHourGlass;
  try
    ipacs.ModifyRequest(intAdviceID, edtExeRoom.Text, StrToInt(edtStudyNo.Text), edtDevice.Text,
      StrToInt(edtHeight.Text), StrToInt(edtWeight.Text), edtStudyDoctor.Text, dtpRequestDate.Date, edtExeDes.Text,intExecOne);

    sErr := ipacs.GetLastError;
    if Trim(sErr) <> '' then begin
      ShowMessage(sErr);
    end else begin
      ShowMessage('??????????????????');
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TForm1.Button9Click(Sender: TObject);
//**********************************************************
//
//????????
//
//**********************************************************
var
  sErr: WideString;
  intAdviceID : Int64;
  intExecOne : Integer;
begin

  if Sender = Button9 then
    begin
      if Trim(edtRecevieAdviceId.Text) = '' then begin
        ShowMessage('????????????????ID??');
        Exit;
      end;
      intAdviceID := StrToInt64(edtRecevieAdviceId.Text);
      intExecOne := 0;
    end
  else
    begin
      if Trim(edtReceiveAdviceIDOne.Text) = '' then begin
        ShowMessage('????????????????ID??');
        Exit;
      end;
      intAdviceID := StrToInt64(edtReceiveAdviceIDOne.Text);
      intExecOne := 1;
    end;

  Screen.Cursor := crHourGlass;
  try
    ipacs.CancelRequest(intAdviceID,intExecOne);

    sErr := ipacs.GetLastError;
    if Trim(sErr) <> '' then begin
      ShowMessage(sErr);
    end else begin
      ShowMessage('??????????????????');
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TForm1.btnDeleteReportClick(Sender: TObject);
  var
  sErr: WideString;
begin
  if Trim(edtAdviceId.Text) = '' then begin
    ShowMessage('????????????????ID??');
    Exit;
  end;

  Screen.Cursor := crHourGlass;
  try

    //????????????????????????,
    ipacs.DeleteReport(StrToInt(edtAdviceId.Text));

  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TForm1.Button10Click(Sender: TObject);
var
  i: Integer;
begin
  OpenDialog1.Filter := '(*.bmp)|*.bmp|(*.jpg)|*.jpg|(*.*)|*.*';
  OpenDialog1.DefaultExt := '*.bmp';
  OpenDialog1.Options := [ofHideReadOnly,ofAllowMultiSelect,ofEnableSizing];

  if not OpenDialog1.Execute then Exit;

  for i := 0 to OpenDialog1.Files.Count - 1 do begin
    if Trim(memECGImage.Text) <> '' then
      memECGImage.Text := memECGImage.Text + ipacs.GetSplitChar;

    memECGImage.Text := memECGImage.Text + OpenDialog1.Files[i];
  end;
end;

procedure TForm1.Button11Click(Sender: TObject);
var
  sErr: WideString;
begin
  if Trim(edtECGAdviceId.Text) = '' then begin
    ShowMessage('????????????????ID??');
    Exit;
  end;

  Screen.Cursor := crHourGlass;
  try

    //????????????????????????,
    ipacs.DeleteElectrocardioReport(StrToInt(edtECGAdviceId.Text));

    //????????????????
    ipacs.SendElectrocardioReport(StrToInt(edtECGAdviceId.Text), edtECGName.Text, memECGImage.Text,
                memECGResult.Text, memECGAdvice.Text, edtECGReport.Text, '');
    sErr := ipacs.GetLastError;
    if Trim(sErr) <> '' then begin
      ShowMessage(sErr);
      Exit;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TForm1.Button12Click(Sender: TObject);
  var
  sErr: WideString;
begin
  if Trim(edtECGAdviceId.Text) = '' then begin
    ShowMessage('????????????????ID??');
    Exit;
  end;

  Screen.Cursor := crHourGlass;
  try

    //????????????????????????,
    ipacs.DeleteElectrocardioReport(StrToInt(edtECGAdviceId.Text));

  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TForm1.Button13Click(Sender: TObject);
//**********************************************************
//
//????????????????????????
//
//**********************************************************
var
  strError: WideString;
begin
  Screen.Cursor := crHourGlass;

  try
    //??????????????
    ipacs.GetRequestInfo1(DateToStr( DateTimePicker1.Date ), DateToStr(DateTimePicker2.Date ), '????');
    strError := ipacs.GetLastError;
    if strError <> '' then begin
      ShowMessage(strError);
      Exit;
    end;

    ReadQueryData(ipacs, lvRequestInf);

    lvAdviceItems.Items.Clear;
    lvAdviceFees.Items.Clear;
    if lvRequestInf.Items.Count > 0 then begin
      lvRequestInf.ItemIndex := 0;
      lvRequestInfClick(nil);
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TForm1.Button14Click(Sender: TObject);
//**********************************************************
//
//????????????????????????????????
//
//**********************************************************
var
  strError: WideString;
begin
  Screen.Cursor := crHourGlass;

  try
    //??????????????????????
    ipacs.GetFinishedRequestInfo(StrToInt( txtPatientId.Text));
    strError := ipacs.GetLastError;
    if strError <> '' then begin
      ShowMessage(strError);
      Exit;
    end;

    ReadQueryData(ipacs, lvFinishedStudy);
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TForm1.Button15Click(Sender: TObject);
//**********************************************************
//
//??????????
//
//**********************************************************
var
  strError: WideString;
  lngAdviceId: Integer;
begin
  if not Assigned( lvFinishedStudy.Selected) then begin
    ShowMessage('????????????????????????????');
    exit;
  end;

  lngAdviceId := StrToInt(lvFinishedStudy.Selected.Caption);
  //??????????????????????
  ipacs.PrintReport(lngAdviceId,true) ;
  strError := ipacs.GetLastError;
  if strError <> '' then begin
    ShowMessage(strError);
  end;
end;

end.

