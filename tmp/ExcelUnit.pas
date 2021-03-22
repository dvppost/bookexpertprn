unit ExcelUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, ActiveX, ComObj, Forms,
  Graphics, Controls;

function RunExcel(DisableAlerts: Boolean = True; Visible: Boolean = False): Boolean;
function StopExcel: Boolean;
function AddWorkBook(AutoRun: Boolean = True): Boolean;

var
  MSExcel: OleVariant;

const
  ExcelApp = 'Excel.Application';

implementation

uses
  MainUnit, HashUnit;

function CheckExcelInstall: Boolean;
var
  ClassID: TCLSID;
begin
  Result := CLSIDFromProgID(PWideChar(WideString(ExcelApp)), ClassID) = S_OK;
end;

function CheckExcelRun: boolean;
begin
  try
    MSExcel := GetActiveOleObject(ExcelApp);
    Result := True;
  except
    Result := False;
  end;
end;

function RunExcel(DisableAlerts: Boolean = True; Visible: Boolean = False): Boolean;
begin
  try
    if CheckExcelInstall then
      begin
        MSExcel := CreateOleObject(ExcelApp);
        MSExcel.Application.EnableEvents := DisableAlerts;
        MSExcel.Visible := Visible;
        Result := True;
      end
    else
      begin
        frmMain.MakeError('Приложение MS Excel не установлено на этом компьютере', 'Ошибка');
        Result := False;
      end;
  except
    Result := False;
  end;
end;

function AddWorkBook(AutoRun: Boolean = True): Boolean;
begin
  Result := CheckExcelRun;
  if (not Result) and (AutoRun) then
  begin
    RunExcel;
    Result := CheckExcelRun;
  end;
  if Result then
    MSExcel.WorkBooks.Add;
end;

function StopExcel: Boolean;
begin
  try
    if MSExcel.Visible then MSExcel.Visible := False;
    MSExcel.Quit;
    MSExcel := Unassigned;
    Result := True;
  except
    Result := False;
  end;
end;

end.
