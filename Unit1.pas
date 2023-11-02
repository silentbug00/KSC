unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ComCtrls, System.JSON, System.NetEncoding,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdHTTP, Excel4Delphi, Excel4Delphi.Stream,
  Vcl.ExtCtrls, LbCipher, LbClass, Math;

type
  TForm1 = class(TForm)
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    Panel1: TPanel;
    Panel2: TPanel;
    Memo1: TMemo;
    OpenDialog1: TOpenDialog;
    Label1: TLabel;
    Edit1: TEdit;
    Edit2: TEdit;
    Edit3: TEdit;
    Edit4: TEdit;
    Label4: TLabel;
    Label3: TLabel;
    Label2: TLabel;
    Button1: TButton;
    LbBlowfish1: TLbBlowfish;
    Button2: TButton;
    Button3: TButton;
    procedure Button2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Button3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    LanjutProses: boolean;
    StartDTTM: TDateTime;
    EndDTTM: TDateTime;
    function DoConfig(AUser, APass, ABaseURL, AFileName: string): boolean;
    function DTIndex(AStr: string): integer;
    function FCIndex(AStr: string): integer;
    procedure WriteLog(AText: string);
  end;

const
  cSubDir = 'Register type';
  cAddtAddr = '/v1/project';
  cChannel = 'channels';
  cDevice = 'devices';
  cTag = 'tags';
  cChannelStart = 4;
  cTagStart = 2;
  cDT: array[0..29] of string = ('string', 'boolean', 'char', 'byte', 'short',
    'word', 'long', 'dword', 'float', 'double', 'bcd', 'lbcd', 'date', 'llong',
    'qword', 'string array', 'boolean array', 'char array', 'byte array',
    'short array', 'word array', 'long array', 'dword array', 'float array',
    'double array', 'bcd array', 'lbcd array', 'date array', 'llong array', 'qword array');
  cFC: array[0..5] of string = ('none', 'dtr', 'rts', 'rts. dtr', 'rts always', 'rts manual');

  cChNmOld      = 'C';
  cChNm         = 'D';
  cChDesc       = 'E';
  cChType       = 'F';
  cChPort       = 'G';
  cChCom        = 'H';
  cChBaud       = 'I';
  cChDataBt     = 'J';
  cChPrty       = 'K';
  cChStopBt     = 'L';
  cChFlowC      = 'M';
  cDvNm         = 'N';
  cDvDesc       = 'O';
  cDvID         = 'P';
  cDvZeroBs     = 'Q';
  cDvRegTyp     = 'R';
  cDvInpCoil    = 'S';
  cDvIntReg     = 'T';
  cDvHoldReg    = 'U';
  cDvIPAddr     = 'V';
  cDvDataColc   = 'W';
  cDvSimul      = 'X';
  cDvInterReqDM = 'Y';


var
  Form1: TForm1;

implementation

{$R *.dfm}

uses uhttpreq;

{ TForm1 }

procedure TForm1.Button1Click(Sender: TObject);
begin
  if OpenDialog1.Execute then
    Edit4.Text := OpenDialog1.FileName;
end;

procedure TForm1.Button2Click(Sender: TObject);
var
  TimerLen: TDateTime;
  h,n,s: Integer;
begin
  Panel1.Enabled := False;
  Button2.Enabled := False;
  Button3.Left := Button2.Left;
  Button3.Top := Button2.Top;
  Button3.Visible := True;
  LanjutProses := True;
  Application.ProcessMessages;
  DoConfig(Edit2.Text, Edit3.Text, Edit1.Text, Edit4.Text);
  Button3.Visible := False;
  Button2.Enabled := True;
  Panel1.Enabled := True;
  TimerLen := EndDTTM - StartDTTM;
  h := Floor(TimerLen) * 24 + StrToInt(FormatDateTime('hh', TimerLen));
  n := StrToInt(FormatDateTime('nn', TimerLen));
  s := StrToInt(FormatDateTime('ss', TimerLen));
  if LanjutProses then
    Memo1.Lines.Add('PROCESS COMPLETED in '+ IntToStr(h) + ' hour(s) ' + IntToStr(n) + ' minute(s) ' + IntToStr(s) + ' second(s) ' + #13+#10)
  else
    Memo1.Lines.Add('PROCESS CANCELED after '+ IntToStr(h) + ' hour(s) ' + IntToStr(n) + ' minute(s) ' + IntToStr(s) + ' second(s) ' + #13+#10);
end;

procedure TForm1.Button3Click(Sender: TObject);
begin
  LanjutProses := False;
end;

function TForm1.DoConfig(AUser, APass, ABaseURL, AFileName: string): boolean;
var
  xwb1: TZWorkBook;
  xRTF, xRecList: TStringList;
  xRootDir: string;
  xRTFileName: string;
  xOldChannel, xNewChannel, xNewDevice, xChannelType: string;
  xRetMsg: string;
  xJSONBody: TJSONObject;
  i,j: integer;
  xtext: string;
  cdscnt, cdfcnt, chscnt, chfcnt, dvscnt, dvfcnt, tgscnt, tgfcnt, otfcnt: integer;
begin
  StartDTTM := Now;
  cdscnt := 0;
  cdfcnt := 0;
  chscnt := 0;
  chfcnt := 0;
  dvscnt := 0;
  dvfcnt := 0;
  tgscnt := 0;
  tgfcnt := 0;
  otfcnt := 0;
  xRootDir := ExtractFileDir(AFileName);
  Result := False;
  xOldChannel := '';
  xwb1 := TZWorkBook.Create(Self);
  try
    WriteLog('Opening file ' + AFileName);
    xwb1.LoadFromFile(AFileName);
    WriteLog('Success');
    i := cChannelStart - 1;
    while (xwb1.Sheets[0].CellRef[cChNmOld, i].AsString <> '') and (LanjutProses) do begin
      if xOldChannel <> xwb1.Sheets[0].CellRef[cChNmOld, i].AsString then begin

        // Delete Old Channel
        xOldChannel := xwb1.Sheets[0].CellRef[cChNmOld, i].AsString;
        WriteLog('Deleting Channel ' + xOldChannel);
        if not SendHTTPCommand(AUser, APass,
                        ABaseURL + cAddtAddr + '/' + cChannel + '/' + xOldChannel,
                        nil, 0, xRetMsg) then begin
          WriteLog('Failed : ' + xRetMsg);
          Inc(cdfcnt);
        end
        else begin
          WriteLog('Success');
          Inc(cdscnt);
        end;

        // Create New Channel
        try
          xNewChannel := xwb1.Sheets[0].CellRef[cChNm, i].AsString;
        except
          on E: Exception do begin
            WriteLog('Failed : ' + 'Column "New Channel" not found');
            Inc(otfcnt);
            xNewChannel := '';
          end;
        end;
        xJSONBody := TJSONObject.Create;
        try
          xJSONBody.AddPair('common.ALLTYPES_NAME', xNewChannel);
          xJSONBody.AddPair('common.ALLTYPES_DESCRIPTION', xwb1.Sheets[0].CellRef[cChDesc, i].AsString);
          xChannelType := xwb1.Sheets[0].CellRef[cChType, i].AsString;
          xJSONBody.AddPair('servermain.MULTIPLE_TYPES_DEVICE_DRIVER', xChannelType);
          if Pos('tcp', LowerCase(xChannelType)) > 0 then begin
            xJSONBody.AddPair('modbus_ethernet.CHANNEL_ETHERNET_PORT_NUMBER', xwb1.Sheets[0].CellRef[cChPort, i].AsInteger);
          end
          else begin
            xJSONBody.AddPair('servermain.CHANNEL_SERIAL_COMMUNICATIONS_COM_ID', xwb1.Sheets[0].CellRef[cChCom, i].AsInteger);
            xJSONBody.AddPair('servermain.CHANNEL_SERIAL_COMMUNICATIONS_BAUD_RATE', xwb1.Sheets[0].CellRef[cChBaud, i].AsInteger);
            xJSONBody.AddPair('servermain.CHANNEL_SERIAL_COMMUNICATIONS_DATA_BITS', xwb1.Sheets[0].CellRef[cChDataBt, i].AsInteger);
            if LowerCase(xwb1.Sheets[0].CellRef[cChPrty, i].AsString) = 'odd' then
              xJSONBody.AddPair('servermain.CHANNEL_SERIAL_COMMUNICATIONS_PARITY', 79)
            else if LowerCase(xwb1.Sheets[0].CellRef['K', i].AsString) = 'even' then
              xJSONBody.AddPair('servermain.CHANNEL_SERIAL_COMMUNICATIONS_PARITY', 69)
            else
              xJSONBody.AddPair('servermain.CHANNEL_SERIAL_COMMUNICATIONS_PARITY', 78);
            xJSONBody.AddPair('servermain.CHANNEL_SERIAL_COMMUNICATIONS_STOP_BITS', xwb1.Sheets[0].CellRef[cChStopBt, i].AsInteger);
            xJSONBody.AddPair('servermain.CHANNEL_SERIAL_COMMUNICATIONS_FLOW_CONTROL', FCIndex(LowerCase(xwb1.Sheets[0].CellRef[cChFlowC, i].AsString)));
          end;
          xtext := xJSONBody.ToJSON;
          WriteLog('Creating Channel ' + xNewChannel);
          if not SendHTTPCommand(AUser, APass,
                          ABaseURL + cAddtAddr + '/' + cChannel,
                          xJSONBody, 1, xRetMsg) then begin
            WriteLog('Failed : ' + xRetMsg);
            Inc(chfcnt) ;
          end
          else begin
            WriteLog('Success');
            Inc(chscnt);
          end;
        except
          on E: Exception do begin
            WriteLog('Failed : ' + E.Message);
            Inc(otfcnt);
          end;
        end;
        FreeAndNil(xJSONBody);
      end;

      // Create New Device
      try
        xNewDevice := xwb1.Sheets[0].CellRef[cDvNm, i].AsString;
      except
        on E: Exception do begin
          WriteLog('Failed : ' + 'Column "New Device" not found');
          Inc(otfcnt);
          xNewDevice := '';
        end;
      end;
      xJSONBody := TJSONObject.Create;
      try
        xJSONBody.AddPair('common.ALLTYPES_NAME', xNewDevice);
        xJSONBody.AddPair('common.ALLTYPES_DESCRIPTION', xwb1.Sheets[0].CellRef[cDvDesc, i].AsString);
        xJSONBody.AddPair('servermain.MULTIPLE_TYPES_DEVICE_DRIVER', xChannelType);
        xJSONBody.AddPair('servermain.DEVICE_INTER_REQUEST_DELAY_MILLISECONDS', xwb1.Sheets[0].CellRef[cDvInterReqDM, i].AsInteger);
        if Pos('tcp', LowerCase(xChannelType)) > 0 then begin
          xJSONBody.AddPair('servermain.DEVICE_ID_STRING', '<' + xwb1.Sheets[0].CellRef[cDvIPAddr, i].AsString + '>.' + xwb1.Sheets[0].CellRef[cDvID, i].AsString);
          xJSONBody.AddPair('modbus_ethernet.DEVICE_ZERO_BASED_BIT_ADDRESSING', LowerCase(xwb1.Sheets[0].CellRef[cDvZeroBs, i].AsString) = 'true');
        end
        else begin
          xJSONBody.AddPair('servermain.DEVICE_ID_DECIMAL', xwb1.Sheets[0].CellRef[cDvID, i].AsInteger);
          xJSONBody.AddPair('modbus.DEVICE_INPUT_COILS', xwb1.Sheets[0].CellRef[cDvInpCoil, i].AsInteger);
          xJSONBody.AddPair('modbus.DEVICE_INTERNAL_REGISTERS', xwb1.Sheets[0].CellRef[cDvIntReg, i].AsInteger);
          xJSONBody.AddPair('modbus.DEVICE_HOLDING_REGISTERS', xwb1.Sheets[0].CellRef[cDvHoldReg, i].AsInteger);
          xJSONBody.AddPair('modbus.DEVICE_ZERO_BASED_BIT_ADDRESSING', LowerCase(xwb1.Sheets[0].CellRef[cDvZeroBs, i].AsString) = 'true');
        end;
        xJSONBody.AddPair('servermain.DEVICE_DATA_COLLECTION', LowerCase(xwb1.Sheets[0].CellRef[cDvDataColc, i].AsString) = 'enable');
        xJSONBody.AddPair('servermain.DEVICE_SIMULATED', LowerCase(xwb1.Sheets[0].CellRef[cDvSimul, i].AsString) = 'yes');

        WriteLog('Creating Device ' + xNewDevice);
        if not SendHTTPCommand(AUser, APass,
                        ABaseURL + cAddtAddr + '/' + cChannel + '/' + xNewChannel + '/' + cDevice,
                        xJSONBody, 1, xRetMsg) then begin
          WriteLog('Failed : ' + xRetMsg);
          Inc(dvfcnt);
        end
        else begin
          WriteLog('Success');
          Inc(dvscnt);
        end;
      except
        on E: Exception do begin
          WriteLog('Failed : ' + E.Message);
          Inc(otfcnt);
        end;
      end;
      FreeAndNil(xJSONBody);

      // Create New Tag
      try
        xRTFileName := xRootDir + '\' + cSubDir + '\' + xwb1.Sheets[0].CellRef[cDvRegTyp, i].AsString + '.csv';
      except
        on E: Exception do begin
          WriteLog('Failed : ' + 'Column "Register Type" not found');
          Inc(otfcnt);
        end;
      end;
      xRTF := TStringList.Create;
      try
        WriteLog('Opening file ' + xRTFileName);
        xRTF.LoadFromFile(xRTFileName);
        WriteLog('Success');
        xRecList := TStringList.Create;
        xRecList.Delimiter := ',';
        xRecList.QuoteChar := '"';
        try
          j := cTagStart - 1;
          while (j < xRTF.Count) and (LanjutProses) do begin
            if xRTF[j] <> '' then begin
              xRecList.DelimitedText := FixCSVLine(xRTF[j]);

              xJSONBody := TJSONObject.Create;
              try
                xJSONBody.AddPair('common.ALLTYPES_NAME', xRecList[0]);
                xJSONBody.AddPair('common.ALLTYPES_DESCRIPTION', xRecList[15]);
                xJSONBody.AddPair('servermain.TAG_ADDRESS', xRecList[1]);
                if xRecList[2] <> '' then
                  xJSONBody.AddPair('servermain.TAG_DATA_TYPE', DTIndex(LowerCase(xRecList[2])));

                if LowerCase(xRecList[4]) = 'ro' then
                  xJSONBody.AddPair('servermain.TAG_READ_WRITE_ACCESS', 0)
                else
                  xJSONBody.AddPair('servermain.TAG_READ_WRITE_ACCESS', 1);

                if xRecList[5] <> '' then
                  xJSONBody.AddPair('servermain.TAG_SCAN_RATE_MILLISECONDS', StrToInt(xRecList[5]));

                if xRecList[6] <> '' then
                  if LowerCase(xRecList[6]) = 'linear' then
                    xJSONBody.AddPair('servermain.TAG_SCALING_TYPE', 1)
                  else
                    xJSONBody.AddPair('servermain.TAG_SCALING_TYPE', 2);

                if xRecList[7] <> '' then
                  xJSONBody.AddPair('servermain.TAG_SCALING_RAW_LOW', StrToInt(xRecList[7]));

                if xRecList[8] <> '' then
                  xJSONBody.AddPair('servermain.TAG_SCALING_RAW_HIGH', StrToInt(xRecList[8]));

                if xRecList[9] <> '' then
                  xJSONBody.AddPair('servermain.TAG_SCALING_SCALED_LOW', StrToInt(xRecList[9]));

                if xRecList[10] <> '' then
                  xJSONBody.AddPair('servermain.TAG_SCALING_SCALED_HIGH', StrToInt(xRecList[10]));

                if xRecList[11] <> '' then
                  xJSONBody.AddPair('servermain.TAG_SCALING_SCALED_DATA_TYPE', DTIndex(LowerCase(xRecList[11])));

                if xRecList[12] <> '' then
                  xJSONBody.AddPair('servermain.TAG_SCALING_CLAMP_LOW', StrToInt(xRecList[12]));

                if xRecList[13] <> '' then
                  xJSONBody.AddPair('servermain.TAG_SCALING_CLAMP_HIGH', StrToInt(xRecList[13]));

                if xRecList[14] <> '' then
                  xJSONBody.AddPair('servermain.TAG_SCALING_UNITS', StrToInt(xRecList[14]));

                if xRecList[16] <> '' then
                  xJSONBody.AddPair('servermain.TAG_SCALING_NEGATE_VALUE', StrToInt(xRecList[16]));

                WriteLog('Creating Tag ' + xRecList[0]);
                if not SendHTTPCommand(AUser, APass,
                                ABaseURL + cAddtAddr + '/' + cChannel + '/' + xNewChannel + '/' + cDevice + '/' + xNewDevice + '/' + cTag,
                                xJSONBody, 1, xRetMsg) then begin
                  WriteLog('Failed : ' + xRetMsg);
                  Inc(tgfcnt);
                end
                else begin
                  WriteLog('Success');
                  Inc(tgscnt);
                end;
              except
                on E: Exception do begin
                  WriteLog('Failed : ' + E.Message);
                  Inc(otfcnt);
                end;
              end;
              FreeAndNil(xJSONBody);
            end;
            Inc(j);
            Application.ProcessMessages;
          end;
        except
          on E: Exception do begin
            WriteLog('Failed : ' + E.Message);
            Inc(otfcnt);
          end;
        end;
        xRecList.Free;
      except
        on E: Exception do begin
          WriteLog('Failed : ' + E.Message);
          Inc(otfcnt);
        end;
      end;
      xRTF.Free;
      Inc(i);
      Application.ProcessMessages;
    end;

    Result := True;
  except
    on E: Exception do begin
      WriteLog('Failed : ' + E.Message);
      Inc(otfcnt);
    end;
  end;
  xwb1.Free;
  EndDTTM := Now;

  Memo1.Lines.Add('');
  Memo1.Lines.Add('Delete Channel Success: ' + IntToStr(cdscnt));
  Memo1.Lines.Add('Delete Channel Failed: ' + IntToStr(cdfcnt));
  Memo1.Lines.Add('Create Channel Success: ' + IntToStr(chscnt));
  Memo1.Lines.Add('Create Channel Failed: ' + IntToStr(chfcnt));
  Memo1.Lines.Add('Create Device Success: ' + IntToStr(dvscnt));
  Memo1.Lines.Add('Create Device Failed: ' + IntToStr(dvfcnt));
  Memo1.Lines.Add('Create Tag Success: ' + IntToStr(tgscnt));
  Memo1.Lines.Add('Create Tag Failed: ' + IntToStr(tgfcnt));
  Memo1.Lines.Add('Other process failed: ' + IntToStr(otfcnt));

end;

function TForm1.DTIndex(AStr: string): integer;
var
  i: integer;
begin
  Result := -1;
  i := 0;
  while (i <= High(cDT)) and (Result = -1) do begin
    if AStr = cDT[i] then
      Result := i;
    Inc(i)
  end;
end;

function TForm1.FCIndex(AStr: string): integer;
var
  i: integer;
begin
  Result := -1;
  i := 0;
  while (i <= High(cFC)) and (Result = -1) do begin
    if AStr = cFC[i] then
      Result := i;
    Inc(i)
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
var
  i: integer;
  xIniF: TStringList;
begin
  Caption := 'KepServer Config Uploader';
  PageControl1.Pages[0].Caption := Caption;
  LbBlowfish1.GenerateKey(Caption);
  for i := 0 to Form1.ComponentCount-1 do begin
    if Form1.Components[i] is TCustomEdit then begin
      TCustomEdit(Form1.Components[i]).Clear;
    end;
  end;
  xIniF := TStringList.Create;
  try
    xIniF.LoadFromFile('KSC.conf');
    Edit1.Text := xIniF.Values['URL'];
    Edit2.Text := xIniF.Values['User'];
    Edit3.Text := LbBlowfish1.DecryptString(xIniF.Values['Pass']);
  except
  end;
  xIniF.Free;
end;

procedure TForm1.FormDestroy(Sender: TObject);
var
  xIniF: TStringList;
begin
  xIniF := TStringList.Create;
  try
    xIniF.Values['URL'] := Edit1.Text;
    xIniF.Values['User'] := Edit2.Text;
    xIniF.Values['Pass'] := LbBlowfish1.EncryptString(Edit3.Text);
    xIniF.SaveToFile('KSC.conf');
  except
  end;
  xIniF.Free;
end;

procedure TForm1.WriteLog(AText: string);
begin
  Memo1.Lines.Add(FormatDateTime('[dd/mm/yy-hh:nn:ss.zzz]', Now) + ' ' + AText);
end;

end.
