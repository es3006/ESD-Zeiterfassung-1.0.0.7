unit uExcelFunktionen;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, System.Zip, StdCtrls, ComCtrls, ExtCtrls, DateUtils, ShellApi,
  FileCtrl, System.UITypes, CryptBase, AESObj, MMSystem, System.IOUtils,
  VCL.FlexCel.Core, FlexCel.XlsAdapter, Vcl.Imaging.pngimage, Generics.Collections,
  FireDAC.Stan.Param, FireDAC.Phys.SQLite, Data.DB, FireDAC.Comp.DataSet, FireDAC.Comp.Client;


function IsXlsFile(const FileName: string): Boolean;
procedure LoadExcelBlattNamenToComboBox(const FileName: string; ComboBox: TComboBox);





implementation


uses
  uMain;


function IsXlsFile(const FileName: string): Boolean;
var
  FileExt: string;
begin
  // Hole die Dateiendung
  FileExt := ExtractFileExt(FileName);

  // Prüfe, ob die Dateiendung ".xlsx" oder ".xlsm" ist (unempfindlich gegenüber Groß-/Kleinschreibung)
  Result := SameText(FileExt, '.xlsx') or SameText(FileExt, '.xlsm');
end;




//Alle Blattnamen der übergebenen Exceldatei in ComboBox ausgeben
procedure LoadExcelBlattNamenToComboBox(const FileName: string; ComboBox: TComboBox);
var
  xls: TXlsFile;
  i, a: Integer;
begin
  xls := TXlsFile.Create;
  try
    xls.Open(FileName);
    xls.IgnoreFormulaText := true;
    if OFZEXLDATEI = 1 then a := 2 else a := 1;
    ComboBox.Clear;
    for i := a to xls.SheetCount do
    begin
      ComboBox.Items.Add(xls.GetSheetName(i));
    end;
  finally
    xls.Free;
  end;
end;




















end.
