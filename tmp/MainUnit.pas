unit MainUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, OleCtrls, ComCtrls, ExtCtrls, Buttons, ADODB,
  Grids, IniFiles, DB;

type
  TfrmMain = class(TForm)
    dlgOpenPrn: TOpenDialog;
    pcMain: TPageControl;
    tsPrint: TTabSheet;
    tsLoad: TTabSheet;
    tsService: TTabSheet;
    Panel5: TPanel;
    gbFiles: TGroupBox;
    Label5: TLabel;
    edBooks: TEdit;
    gbBase: TGroupBox;
    Label4: TLabel;
    Label3: TLabel;
    Label2: TLabel;
    Label1: TLabel;
    lblBaseStatus: TLabel;
    btnConnect: TSpeedButton;
    Label18: TLabel;
    edBasePath: TEdit;
    edBaseName: TEdit;
    edBaseUser: TEdit;
    edBasePass: TEdit;
    edBaseDriver: TEdit;
    Label6: TLabel;
    edOrders: TEdit;
    Label7: TLabel;
    edBlocks: TEdit;
    Label8: TLabel;
    edCovers: TEdit;
    pnSQL: TPanel;
    Panel1: TPanel;
    btnSql: TSpeedButton;
    btnSQL2Excel: TSpeedButton;
    cbTemplate: TComboBox;
    sgTemplate: TStringGrid;
    conMain: TADOConnection;
    qTemp: TADOQuery;
    sgOrder: TStringGrid;
    pnButtons: TPanel;
    btnCovers: TSpeedButton;
    btnBlocks: TSpeedButton;
    tsLog: TTabSheet;
    GroupBox1: TGroupBox;
    mmErrors: TMemo;
    GroupBox2: TGroupBox;
    mmLog: TMemo;
    procedure MakeError(sText, sTitle: String);
    function  UpdateCfgFile(sFile, sCopies: String): Boolean;
    function  PutCoverToQueue(sFile: WideString; sCopies: String): Integer;
    procedure btnCoversClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure cbTemplateChange(Sender: TObject);
    procedure DeleteRow(Grid: TStringGrid; iRow: Integer);
    procedure btnConnectClick(Sender: TObject);
    procedure conMainAfterConnect(Sender: TObject);
    procedure conMainAfterDisconnect(Sender: TObject);
    procedure cbTemplateDropDown(Sender: TObject);
    function  LoadOrder(sFileName: String): Integer;
    function  LoadTemplate(sFile: String): Integer;
    function  GetParam(sParam: String; iIdx: Integer): String;
    procedure CorrectPaths;
    procedure btnBlocksClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;
  sLogFile:     String;

const
  iOffset = 32;
  iMaxRow = 10;

implementation

uses  ExcelUnit;

{$R *.dfm}

function TrimSpecial(strInput: String): String;
var
  i: Integer;
begin
Result := strInput;
for i := 0 to 13 do
  Result := StringReplace(Result, Chr(i), '', [rfReplaceAll, rfIgnoreCase]);
end;

function StringReplaceExt(const S : string; OldPattern, NewPattern:  array of string; Flags: TReplaceFlags):string;
var
 i : integer;
begin
   Assert(Length(OldPattern)=(Length(NewPattern)));
   Result:=S;
   for  i:= Low(OldPattern) to High(OldPattern) do
    Result:=StringReplace(Result,OldPattern[i], NewPattern[i], Flags);
end;

function	IsNumber(str: String): Boolean;
var
  i: Integer;
begin
i := 1;
Result := Length(str) > 0;
while ((i < Length(str)) and (Result)) do begin
  Result := (Ord(str[i]) in [48..57]);
  Inc(i);
end;
end;

procedure WriteLog(sLine: string);
var
  FileStream : TFileStream;
begin
FileStream := nil;
try
  if not FileExists(sLogFile) then
  try
    try
      FileStream := TFileStream.Create(sLogFile, fmCreate);
    finally
      FileStream.Free;
    end;
  except
  end;
  try
    try
      FileStream := TFileStream.Create(sLogFile, fmOpenReadWrite + fmShareDenyNone);
      FileStream.Position := FileStream.Size;
      sLine := FormatDateTime('dd-mm-yyyy hh:nn:ss ', Now) + sLine;
      FileStream.Write(sLine[1], Length(sLine));
    finally
      FileStream.Write(#$D#$A, 2);
      FileStream.Free;
      frmMain.mmLog.Lines.Add(sLine);
    end;
  except
  end;
except
end;
end;

function HexToInt(sStr: String): Integer;
var
  i: Integer;
  r: Integer;
begin
  val('$' + Trim(sStr), r, i);
  if i <> 0 then HexToInt := 0
  else HexToInt := r;
end;

function DecodeLine(sLine: String): String;
var
  i:    Integer;
  iSym: Integer;
begin
Result := '';
i := 1;
while (i < Length(sLine)) do
begin
  iSym := HexToInt(sLine[i] + sLine[i + 1]);
  if (iSym - iOffset < 32) then
    Result := Result + Chr(iSym - iOffset + 94)
  else
    Result := Result + Chr(iSym - iOffset);
  i := i + 2;
end;
end;

function EncodeLine(sLine: String): String;
var
  i: Integer;
begin
Result := '';
for i := 1 to Length(sLine) do
begin
  if (Ord(sLine[i]) + iOffset > 126) then
    Result := Result + IntToHex(Ord(sLine[i]) + iOffset - 94, 2)
  else
    Result := Result + IntToHex(Ord(sLine[i]) + iOffset, 2)
end;
end;

function  FindLine(sl: TStrings; sLine: String): Integer;
var
  i:      Integer;
  sLLine: String;
  sSl:    String;
begin
Result := -1;
sLLine := AnsiLowerCase(sLine);
i := 0;
while ((i < sl.Count - 1) and (Result = -1)) do
begin
  sSl := AnsiLowerCase(sl[i]);
  if (Pos(sLLine, sSl) > 0) then
    Result := i;
  Inc(i);
end;
end;

function SearchFile(sFile: WideString): WideString;
var
  fdData: TSearchRec;
begin
Result := '';
if (sFile[Length(sFile)] = '\') then
  sFile[Length(sFile)] := #0;
if (FindFirst(sFile, faAnyFile, fdData) = 0) then
    Result := fdData.Name;
FindClose(fdData);
end;

procedure TfrmMain.MakeError(sText, sTitle: String);
begin
if (sTitle = '') then sTitle := 'Ошибка';
MessageBox(Application.Handle, PWideChar(sText),
          PWideChar(sTitle), MB_OK + MB_ICONHAND);
end;

function TfrmMain.GetParam(sParam: String; iIdx: Integer): String;
var
  i: Integer;
begin
Result := '000000000000000';
for i := 0 to iMaxRow do
  if (sgTemplate.Cells[0, i] = sParam) then
  begin
    Result := sgOrder.Cells[StrToInt(sgTemplate.Cells[1, i]) - 1, iIdx];
    Break;
  end;
end;

function TfrmMain.LoadOrder(sFileName: String): Integer;
var
  Rows:       Integer;
  Cols:       Integer;
  i:          Integer;
  j:          Integer;
  WorkSheet:  OLEVariant;
  FData:      OLEVariant;
begin
Result := -1;
if (not RunExcel(False, False)) then Exit;
try
  Screen.Cursor := crHourGlass;
  sgOrder.Enabled := False;
  for i := 0 to sgOrder.ColCount - 1 do
    sgOrder.Cols[i].Clear;
  MSExcel.Workbooks.Open(sFileName);
  WorkSheet := MSExcel.ActiveWorkbook.ActiveSheet;
  Cols := 2;
  while not (Unassigned = WorkSheet.Cells[1, Cols].Value) do
    Inc(Cols);
  i := 0;
  Rows := 2;
  while (i < 10) do
  begin
    if (WorkSheet.Cells[Rows, 1].Value = Unassigned) then
      Inc(i)
    else
      i := 0;
    Inc(Rows);
  end;
  Dec(Cols);
  FData := WorkSheet.UsedRange.Value;
  sgOrder.ColCount := Cols;
  sgOrder.RowCount := Rows;
  for i := 0 to Rows - 1 do begin
    for j := 0 to Cols - 1 do
      sgOrder.Cells[j, i] := FData[i + 1, j + 1];
    Application.ProcessMessages;
  end;
  Result := 0;
finally
  ExcelUnit.StopExcel;
  Screen.Cursor := crDefault;
  frmMain.sgOrder.Enabled := True;
end;
end;

function  TfrmMain.UpdateCfgFile(sFile, sCopies: String): Boolean;
var
  slCfg:  TStringList;
  iPos:   Integer;
begin
Result := False;
slCfg := TStringList.Create;
slCfg.LoadFromFile(sFile);
iPos := FindLine(slCfg, '<keyvalue key="num copies"');
if (iPos = -1) then begin
  iPos := FindLine(slCfg, '</JobSetting>');
  if (iPos > -1) then
    slCfg.Insert(iPos, 'Inserted line');
end;
if (iPos > -1) then
begin
  slCfg[iPos] := '    <KEYVALUE Key="num copies" Value="' +
                 sCopies +
                 '" KeyDisplayValue="РљРѕРїРёРё" LocDisplayValue="1"/>';
  try
    slCfg.SaveToFile(sFile);
    WriteLog('Сохранены настройки для файла ' + sFile + '. Ждем принтер...');
    Result := True;
  except
  end;
end;
slCfg.Free;
end;

function TfrmMain.PutCoverToQueue(sFile: WideString; sCopies: String): Integer;
var
  sCfg:     String;
  sOutFile: WideString;
begin
Result := -1;
sCfg := edCovers.Text + '2_na_list\[_EFI_HotFolder_]\Folder.cfg';
if (not UpdateCfgFile(sCfg, sCopies)) then begin
  WriteLog('Ошибка сохранения настроек для ' + sFile);
  Exit;
end;
sOutFile := ExtractFileName(sFile);
sOutFile := edCovers.Text + '2_na_list\' + sOutFile;
if (not CopyFileW(PWideChar(sFile), PWideChar(sOutFile), True)) then
  WriteLog('Ошибка копирования файла ' + sFile)
else
begin
  while FileExists(sOutFile) do
  begin
    Application.ProcessMessages;
    Sleep(500);
  end;
  Result := 1;
end;
end;

procedure TfrmMain.btnCoversClick(Sender: TObject);
var
  sFile:      WideString;
  sDir:       WideString;
  sPdfFile:   WideString;
  sTemplate:  String;
  sIsbn:      String;
  sCopies:    String;
  i:          Integer;
  iFound:     Integer;
begin
sFile := edOrders.Text;
if SearchFile(sFile) <> '' then
  dlgOpenPrn.InitialDir := edOrders.Text;
if (not dlgOpenPrn.Execute) then Exit;
iFound := 0;
sFile := ExtractFileName(dlgOpenPrn.FileName);
for i := 0 to cbTemplate.Items.Count - 1 do
begin
  sTemplate := Copy(cbTemplate.Items[i], 1, Length(cbTemplate.Items[i]) - 5);
  if (Pos(WideString(sTemplate), sFile) > 0) then
    if (LoadTemplate(ExtractFilePath(ParamStr(0)) +
        'templates\' + cbTemplate.Items[i]) = 0) then
    begin
      iFound := 1;
      Break;
    end;
end;
if (iFound > 0) then
begin
  LoadOrder(dlgOpenPrn.FileName);
  for i := 0 to sgOrder.RowCount - 1 do
    if IsNumber(sgOrder.Cells[0, i]) then
    begin
      sFile := ExtractFileName(dlgOpenPrn.FileName);
      sFile := Copy(sFile, 1, LastDelimiter('.', sFile) - 1);
      sDir := edBooks.Text;
      sIsbn := GetParam('Isbn', i);
      sPdfFile := SearchFile(sDir + '*' + sIsbn + '*.*');
      if (sPdfFile <> '') then
      begin
        sDir := sDir + sPdfFile;
        if (sDir[Length(sDir)] <> '\') then
          sDir := sDir + '\';
        sPdfFile := SearchFile(sDir + '*' + Trim(sIsbn) + '-cover.pdf');
        if (sPdfFile <> '') then
        begin
          sFile := edCovers.Text + sFile + '_' + GetParam('BookFormat', i) +
                  '_' + sPdfFile;
          sCopies := GetParam('NumberOfCopies', i);
          if (sCopies = '0') then
            WriteLog('Пропускаем файл, тираж 0: ' + sDir + sPdfFile)
          else begin
            WriteLog('Копируем файл, тираж ' + sCopies + ': ' +
                     sDir + sPdfFile + ' в ' + sFile);
            PutCoverToQueue(sDir + sPdfFile, sCopies);
          end;
        end
        else
        begin
          WriteLog('Не удалось найти файл ' + sIsbn);
          mmErrors.Lines.Add(sIsbn + ' - не удалось найти файл');
        end;
      end
      else
      begin
        WriteLog('Нет директории макетов ' + sIsbn);
        mmErrors.Lines.Add(sIsbn + ' - нет директории макетов');
      end;
      Application.ProcessMessages;
    end;
end
else
  MakeError('Не найден шаблон загрузки для файла ' + dlgOpenPrn.FileName, '');
end;

procedure TfrmMain.btnBlocksClick(Sender: TObject);
var
  sFile:      WideString;
  sDir:       WideString;
  sPdfFile:   WideString;
  sTemplate:  String;
  sIsbn:      String;
  sCopies:    String;
  i:          Integer;
  iFound:     Integer;
begin
sFile := edOrders.Text;
if SearchFile(sFile) <> '' then
  dlgOpenPrn.InitialDir := edOrders.Text;
if (not dlgOpenPrn.Execute) then Exit;
iFound := 0;
sFile := ExtractFileName(dlgOpenPrn.FileName);
for i := 0 to cbTemplate.Items.Count - 1 do
begin
  sTemplate := Copy(cbTemplate.Items[i], 1, Length(cbTemplate.Items[i]) - 5);
  if (Pos(WideString(sTemplate), sFile) > 0) then
    if (LoadTemplate(ExtractFilePath(ParamStr(0)) +
        'templates\' + cbTemplate.Items[i]) = 0) then
    begin
      iFound := 1;
      Break;
    end;
end;
if (iFound > 0) then
begin
  LoadOrder(dlgOpenPrn.FileName);
  for i := 0 to sgOrder.RowCount - 1 do
    if IsNumber(sgOrder.Cells[0, i]) then
    begin
      sFile := ExtractFileName(dlgOpenPrn.FileName);
      sFile := Copy(sFile, 1, LastDelimiter('.', sFile) - 1);
      sDir := edBooks.Text;
      sIsbn := GetParam('Isbn', i);
      sPdfFile := SearchFile(sDir + '*' + sIsbn + '*.*');
      if (sPdfFile <> '') then
      begin
        sDir := sDir + sPdfFile;
        if (sDir[Length(sDir)] <> '\') then
          sDir := sDir + '\';
        sPdfFile := SearchFile(sDir + '*' + Trim(sIsbn) + '.pdf');
        if (sPdfFile <> '') then
        begin
          sFile := edBlocks.Text + sFile + '_' + GetParam('BookFormat', i) +
                  '_' + sPdfFile;
          if (sCopies = '0') then
            WriteLog('Пропускаем файл, тираж 0: ' + sDir + sPdfFile)
          else begin
            WriteLog('Копируем файл, тираж ' + sCopies + ': ' +
                     sDir + sPdfFile + ' в ' + sFile);
            if (not CopyFileW(PWideChar(sDir + sPdfFile), PWideChar(sFile), True)) then
              WriteLog('Ошибка копирования файла ' + sFile);
          end;
        end
        else
        begin
          WriteLog('Не удалось найти файл ' + sIsbn);
          mmErrors.Lines.Add(sIsbn + ' - не удалось найти файл');
        end;
      end
      else
      begin
        WriteLog('Нет директории макетов ' + sIsbn);
        mmErrors.Lines.Add(sIsbn + ' - нет директории макетов');
      end;
    end;
end
else
  MakeError('Не найден шаблон загрузки для файла ' + dlgOpenPrn.FileName, '');
end;

procedure TfrmMain.CorrectPaths;
begin
if (edBooks.Text[Length(edBooks.Text)] <> '\') then
  edBooks.Text := edBooks.Text + '\';
if (edOrders.Text[Length(edOrders.Text)] <> '\') then
  edOrders.Text := edOrders.Text + '\';
if (edBlocks.Text[Length(edBlocks.Text)] <> '\') then
  edBlocks.Text := edBlocks.Text + '\';
if (edCovers.Text[Length(edCovers.Text)] <> '\') then
  edCovers.Text := edCovers.Text + '\';
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var
  ini:      TIniFile;
  sIniFile: String;
begin
sIniFile := ExtractFilePath(ParamStr(0)) + 'bookexpertprn.ini';
ini := TIniFile.Create(sIniFile);
edBaseDriver.Text := ini.ReadString('Base', 'Driver', 'MySQL ODBC 8.0 Unicode Driver');
edBasePath.Text := ini.ReadString('Base', 'Path', 'localhost');
edBaseName.Text := ini.ReadString('Base', 'BaseName', '');
edBaseUser.Text := ini.ReadString('Base', 'User', '');
edBasePass.Text := DecodeLine(ini.ReadString('Base', 'Pass', ''));
edBooks.Text := ini.ReadString('Paths', 'Books',  '');
edOrders.Text := ini.ReadString('Paths', 'Orders', '');
edBlocks.Text := ini.ReadString('Paths', 'Blocks', '');
edCovers.Text := ini.ReadString('Paths', 'Covers', '');
case ini.ReadInteger('Settings', 'Window', 0) of
0:  frmMain.WindowState := wsNormal;
1:  frmMain.WindowState := wsMinimized;
2:  frmMain.WindowState := wsMaximized;
end;
ini.Free;
CreateDir(ExtractFilePath(ParamStr(0)) + 'logs');
sLogFile := ExtractFilePath(ParamStr(0)) + 'logs\' + FormatDateTime('ddmmyy', Now) + 'prn.log';
conMain.ConnectionString := 'Provider=MSDASQL.1;Persist Security Info=False;' +
                            'Extended Properties="Driver={' +
                            edBaseDriver.Text + '};SERVER=' +
                            edBasePath.Text + ';UID=' +
                            edBaseUser.Text + ';Pwd=' +
                            edBasePass.Text + ';DATABASE=' +
                            edBaseName.Text + ';PORT=3306"';
cbTemplateDropDown(nil);
conMain.Connected := True;
CorrectPaths;
end;

procedure TfrmMain.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if (ssCtrl in Shift) and (Key = VK_BACK) then tsService.TabVisible := True;
if (ssCtrl in Shift) and (Key = VK_RETURN) then tsLoad.TabVisible := True;
end;

procedure TfrmMain.FormClose(Sender: TObject; var Action: TCloseAction);
var
  ini:      TIniFile;
  sIniFile: String;
begin
CorrectPaths;
sIniFile := ExtractFilePath(ParamStr(0)) + 'BookExpertPrn.ini';
ini := TIniFile.Create(sIniFile);
ini.WriteString('Base', 'Driver', edBaseDriver.Text);
ini.WriteString('Base', 'Path', edBasePath.Text);
ini.WriteString('Base', 'BaseName', edBaseName.Text);
ini.WriteString('Base', 'User', edBaseUser.Text);
ini.WriteString('Base', 'Pass', EncodeLine(edBasePass.Text));
ini.WriteString('Paths', 'Books', edBooks.Text);
ini.WriteString('Paths', 'Orders', edOrders.Text);
ini.WriteString('Paths', 'Blocks', edBlocks.Text);
ini.WriteString('Paths', 'Covers', edCovers.Text);
case frmMain.WindowState of
wsNormal: ini.WriteInteger('Settings', 'Window', 0);
wsMinimized: ini.WriteInteger('Settings', 'Window', 1);
wsMaximized: ini.WriteInteger('Settings', 'Window', 2);
end;
ini.Free;
end;

procedure TfrmMain.DeleteRow(Grid: TStringGrid; iRow: Integer);
var
  i:  Integer;
  j:  Integer;
begin
with sgTemplate do
begin
  for i := iRow + 1 to RowCount - 1 do
    for j := 0 to ColCount - 1 do
      Cells[j, i - 1] := Cells[j, i];
  for i := 0 to ColCount - 1 do
    Cells[i, RowCount - 1] := '';
end;
end;

function TfrmMain.LoadTemplate(sFile: String): Integer;
var
  sCell: String;
  i:     Integer;
begin
Result := 0;
with sgTemplate do
begin
  for i := 0 to RowCount - 1 do
    Rows[i].Clear;
  RowCount := iMaxRow;
  Cols[0].LoadFromFile(sFile);
  i := 0;
  while ((i < iMaxRow) and (Cells[0, i] <> '')) do
  begin
    if (Pos(':', Cells[0, i]) = 0) then
      DeleteRow(sgTemplate, i)
    else begin
      sCell := Cells[0, i];
      sCell := Copy(sCell, Pos(':', sCell) + 1, Length(sCell));
      sCell := StringReplaceExt(sCell, [' ', '"', ','], ['', '', ''], [rfReplaceAll]);
      if (not IsNumber(sCell)) then
      begin
        MakeError('Не корректный шаблон ' + sFile, '');
        Result := -1;
        Break;
      end;
      Cells[1, i] := sCell;
      sCell := Cells[0, i];
      sCell := Copy(sCell, 1, Pos(':', sCell) - 1);
      sCell := StringReplaceExt(sCell, [' ', '"', ','], ['', '', ''], [rfReplaceAll]);
      Cells[0, i] := sCell;
      Inc(i);
    end;
    if (Cells[0, i] = '}') then
    begin
      RowCount := i;
      Break;
    end;
  end;
end;
end;

procedure TfrmMain.cbTemplateChange(Sender: TObject);
begin
LoadTemplate(ExtractFilePath(ParamStr(0)) + 'templates\' +
  cbTemplate.Items[cbTemplate.ItemIndex]);
end;

procedure TfrmMain.cbTemplateDropDown(Sender: TObject);
var
  srchFiles: TSearchRec;
  sPath:     String;
begin
cbTemplate.Items.Clear;
sPath := ExtractFilePath(ParamStr(0));
if FindFirst(sPath + 'templates\*.*', faAnyFile - faDirectory, srchFiles) = 0 then
begin
    repeat if (srchFiles.Name <> '.') and (srchFiles.Name <> '..') then
      cbTemplate.Items.Add(srchFiles.Name);
    until FindNext(srchFiles) <> 0;
    FindClose(srchFiles);
end;
end;

procedure TfrmMain.btnConnectClick(Sender: TObject);
begin
conMain.Connected := (btnConnect.Caption = '&Подключить');
end;

procedure TfrmMain.conMainAfterConnect(Sender: TObject);
begin
lblBaseStatus.Caption := 'Статус: подключено';
btnConnect.Caption := '&Отключить';
end;

procedure TfrmMain.conMainAfterDisconnect(Sender: TObject);
begin
lblBaseStatus.Caption := 'Статус: отключено';
btnConnect.Caption := '&Подключить';
end;

end.
