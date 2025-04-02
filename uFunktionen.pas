unit uFunktionen;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, System.Zip, StdCtrls, ComCtrls, ExtCtrls, DateUtils, ShellApi,
  FileCtrl, System.UITypes, CryptBase, AESObj, MMSystem, System.IOUtils,
  VCL.FlexCel.Core, FlexCel.XlsAdapter, Vcl.Imaging.pngimage;



function SubtractOneSecond(TimeStr: string): string;
procedure PlayResourceMP3(ResEntryName, TempFileName: string);
procedure SaveResourceToFile(ResourceName, FileName: string);
procedure LoadImageFromResource(const ResourceName: string; TargetImage: TImage);
procedure ZipDir(const Dir: string);
function ConvertGermanDateToSQLDate(const GermanDate: string; ShowTime: boolean = false): string;
function ConvertSQLDateToGermanDate(const SQLDate: string; ShowTime: boolean = true; ShortYear: boolean = false): string;
function ReplaceUmlauteWithHtmlEntities(const InputString: string): string;
function DeleteFiles(const AFile: string): boolean;
function IsFileZeroSize(const FileName: string): Boolean;
procedure ReadExcelRange(const FileName: string; const StartRow, StartCol, EndRow, EndCol: Integer; Output: TStrings);
function CountEntriesInColumnB(Xls: TXlsFile; StartRow, EndRow: Integer): Integer;
function ExtractField(const s: string; const delimiter: char; index: integer): string;



implementation

uses
  uMain;


function SubtractOneSecond(TimeStr: string): string;
var
  TimeValue: TDateTime;
begin
  // String in TDateTime umwandeln
  TimeValue := StrToTime(TimeStr);

  // Eine Sekunde abziehen
  TimeValue := TimeValue - EncodeTime(0, 0, 1, 0);

  // Ergebnis zurückgeben
  Result := FormatDateTime('hh:nn:ss', TimeValue);
end;



procedure PlayResourceMP3(ResEntryName, TempFileName: string);
var
  ResStream: TResourceStream;
  MemStream: TMemoryStream;
begin
  // Prüfen, ob die Datei bereits vorhanden ist
  if not FileExists(TempFileName) then
  begin
    // Datei aus der Resource holen und speichern
    ResStream := TResourceStream.Create(HInstance, ResEntryName, RT_RCDATA);
    try
      MemStream := TMemoryStream.Create;
      try
        MemStream.LoadFromStream(ResStream);
        MemStream.SaveToFile(TempFileName);
      finally
        MemStream.Free;
      end;
    finally
      ResStream.Free;
    end;
  end;

  // Datei abspielen
  PlaySound(PChar(TempFileName), 0, SND_FILENAME or SND_ASYNC);
end;



procedure SaveResourceToFile(ResourceName, FileName: string);
var
  ResStream: TResourceStream;
  FileStream: TFileStream;
begin
  if not FileExists(FileName) then
  begin
    ResStream := TResourceStream.Create(HInstance, ResourceName, RT_RCDATA); // RT_RCDATA ist ein gängiger Typ für benutzerdefinierte Ressourcen
    try
      FileStream := TFileStream.Create(FileName, fmCreate);
      try
        FileStream.CopyFrom(ResStream, 0); // Kopiere den Ressourceninhalt in die Datei
      finally
        FileStream.Free;
      end;
    finally
      ResStream.Free;
    end;
  end;
end;




procedure LoadImageFromResource(const ResourceName: string; TargetImage: TImage);
var
  ResourceStream: TResourceStream;
  PngImage: TPngImage;
begin
  // Erstelle den Stream, um die angegebene Ressource zu laden
  ResourceStream := TResourceStream.Create(HInstance, ResourceName, RT_RCDATA);
  try
    // Erstelle ein TPngImage-Objekt, um die PNG-Daten zu laden
    PngImage := TPngImage.Create;
    try
      // Lade die PNG-Daten aus dem Stream in das TPngImage-Objekt
      PngImage.LoadFromStream(ResourceStream);
      // Weise das geladene Bild der angegebenen TImage-Komponente zu
      TargetImage.Picture.Graphic := PngImage;
    finally
      PngImage.Free;
    end;
  finally
    ResourceStream.Free;
  end;
end;








procedure ZipDir(const Dir: string);
var
  ZipFile: TZipFile;
  SearchRec: TSearchRec;
  FilePath: string;
  FullDirPath: string;
begin
  // Sicherstellen, dass der Pfad ohne zusätzliche Leerzeichen oder ungültige Zeichen ist
  FullDirPath := Trim(Dir);

  ZipFile := TZipFile.Create;
  try
    ZipFile.Open(FullDirPath + '.zip', zmWrite);

    // Verzeichnisinhalt durchsuchen und Dateien zur Zip-Datei hinzufügen
    if FindFirst(IncludeTrailingPathDelimiter(FullDirPath) + '*.sql', faAnyFile, SearchRec) = 0 then
    begin
      repeat
        FilePath := IncludeTrailingPathDelimiter(FullDirPath) + SearchRec.Name;
        ZipFile.Add(FilePath, ExtractFileName(FilePath));
      until FindNext(SearchRec) <> 0;
      FindClose(SearchRec);
    end;

    ZipFile.Close;
  finally
    ZipFile.Free;
  end;

  // Nach erfolgreicher Zip-Erstellung Verzeichnis löschen
  if SysUtils.DirectoryExists(FullDirPath) then
    TDirectory.Delete(FullDirPath, True);
end;







function IsFileZeroSize(const FileName: string): Boolean;
var
  FileInfo: TSearchRec;
begin
  Result := False;
  if FindFirst(FileName, faAnyFile, FileInfo) = 0 then
  try
    Result := FileInfo.Size = 0;
  finally
    FindClose(FileInfo);
  end;
end;







// Funktion zum Ersetzen von Umlauten durch HTML-Entities
function ReplaceUmlauteWithHtmlEntities(const InputString: string): string;
begin
  // Ersetzen Sie die Umlaute durch die entsprechenden HTML-Entities
  Result := StringReplace(InputString, 'ä', '&auml;', [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, 'Ä', '&Auml;', [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, 'ö', '&ouml;', [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, 'Ö', '&Ouml;', [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, 'ü', '&uuml;', [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, 'Ü', '&Uuml;', [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, 'ß', '&szlig;', [rfReplaceAll, rfIgnoreCase]);
end;






function DeleteFiles(const AFile: string): boolean;
var
  sh: SHFileOpStruct;
begin
  ZeroMemory(@sh, SizeOf(sh));
  with sh do
  begin
    Wnd := Application.Handle;
    wFunc := FO_DELETE;
    pFrom := PChar(AFile +#0);
    fFlags := FOF_SILENT or FOF_NOCONFIRMATION;
  end;
  result := SHFileOperation(sh) = 0;
end;










{
  Aktuelle Version um ein deutsches Dateum im Format DD.MM.YY oder DD.MM.YYYY HH:NN:SS in ein
  SQL-Datum im Format YYYY-MM-DD oder YYYY-MM-DD HH:NN:SS umzuwandeln

  Aufruf mit
  ConvertGermanDateToSQLDate('30.06.1975'); //Nur Datum
  ConvertGermanDateToSQLDate('30.06.1975 10:30:20'); //Datum und Zeit
}
function ConvertGermanDateToSQLDate(const GermanDate: string; ShowTime: boolean = false): string;
var
  DateValue: TDateTime;
  FormatedDate: string;
begin
  if(GermanDate = '') then
  begin
    result := '';
    exit;
  end;

  // Prüfen und Konvertieren des Datumsformats von DD.MM.YYYY HH:NN:SS zu einem TDateTime-Wert
  if TryStrToDateTime(GermanDate, DateValue, FormatSettings) then
  begin
    if(ShowTime = true) then
      FormatedDate := FormatDateTime('yyyy-mm-dd hh:nn', DateValue, FormatSettings)
    else
      FormatedDate := FormatDateTime('yyyy-mm-dd', DateValue, FormatSettings);

    Result := FormatedDate;
  end
  else
  begin
    // Bei Fehler wird ein leerer String zurückgegeben
    Result := '';
    ShowMessage('Ungültiges Datumsformat: ' + GermanDate);
    abort;
  end;
end;







function ConvertSQLDateToGermanDate(const SQLDate: string; ShowTime: boolean = true; ShortYear: boolean = false): string;
var
  DateValue: TDateTime;
  FormattedDate: string;
  FormatSettings: TFormatSettings;
begin
  // Spezifische FormatSettings für die Konvertierung von Datums- und Zeitwerten konfigurieren
  FormatSettings := TFormatSettings.Create;
  FormatSettings.DateSeparator   := '-';
  FormatSettings.TimeSeparator   := ':';
  FormatSettings.ShortDateFormat := 'yyyy-mm-dd';
  FormatSettings.LongTimeFormat  := 'hh:nn:ss';

  if(SQLDate = '') then
  begin
    result := '';
    exit;
  end;
  // Konvertieren des Datumsformats von YYYY-MM-DD HH:NN:SS zu einem TDateTime-Wert
  if TryStrToDateTime(SQLDate, DateValue, FormatSettings) then
  begin
    // Spezifische FormatSettings für die deutsche Schreibweise konfigurieren
    FormatSettings.DateSeparator   := '.';
    FormatSettings.ShortDateFormat := 'dd.mm.yyyy';
    FormatSettings.LongTimeFormat  := 'hh:nn';

    // Das Datum im Format DD.MM.YYYY HH:NN formatiert
    if(ShowTime = true) then
      FormattedDate := FormatDateTime('dd.mm.yyyy hh:nn', DateValue, FormatSettings)
    else
      FormattedDate := FormatDateTime('dd.mm.yyyy', DateValue, FormatSettings);

    // Das Datum im Format DD.MM.YY formatiert
    if(ShortYear = true) then
      FormattedDate := FormatDateTime('dd.mm.yy', DateValue, FormatSettings)
    else
      FormattedDate := FormatDateTime('dd.mm.yyyy', DateValue, FormatSettings);



    Result := FormattedDate;
  end
  else
  begin
    // Bei Fehler wird ein leerer String zurückgegeben
    Result := '';
    ShowMessage('Ungültiges Datumsformat: ' + SQLDate);
    Abort;
  end;
end;






procedure ReadExcelRange(const FileName: string; const StartRow, StartCol, EndRow, EndCol: Integer; Output: TStrings);
var
  xls: TXlsFile;
  Row, Col: Integer;
  CellValue: string;
  RowOutput: string;
  IsRowEmpty: Boolean;
begin
  xls := TXlsFile.Create;
  try
    xls.Open(FileName);
    xls.ActiveSheetByName := selSheet;

    for Row := StartRow to EndRow do
    begin
      RowOutput := '';
      IsRowEmpty := True; // Angenommen, die Zeile ist leer

      for Col := StartCol to EndCol do
      begin
        CellValue := xls.GetCellValue(Row, Col).ToString;
        if Trim(CellValue) <> '' then
          IsRowEmpty := False; // Zeile enthält nicht-leere Zellen

        if Col = StartCol then
          RowOutput := CellValue
        else
          RowOutput := RowOutput + #9 + CellValue; // Tabulator als Trenner
      end;

      // Entfernen von Leerzeichen am Ende der Zeile
      RowOutput := TrimRight(RowOutput);

      // Füge die Zeile nur hinzu, wenn sie nicht leer ist
      if not IsRowEmpty then
        Output.Add(RowOutput);
    end;
  finally
    xls.Free;
  end;
end;




//Anzahl EInträge in der Exceldatei in einem bestimmten Bereich ermitteln
function CountEntriesInColumnB(Xls: TXlsFile; StartRow, EndRow: Integer): Integer;
var
  Row: Integer;
  CellValue: TCellValue;
  Count: Integer;
begin
  Count := 0;

  // Durchlaufe die Zeilen von StartRow bis EndRow in Spalte B
  for Row := StartRow to EndRow do
  begin
    CellValue := Xls.GetCellValue(Row, SPALTEMA); // Spalte B ist die 2. Spalte

    // Prüfe, ob die Zelle einen Wert enthält (kein leerer String oder Leerwert)
    if not CellValue.IsEmpty then
    begin
      Inc(Count);
    end;
  end;

  Result := Count;
end;






function ExtractField(const s: string; const delimiter: char; index: integer): string;
var
  fields: TArray<string>;
begin
  fields := s.Split([delimiter]);
  if (index >= Low(fields)) and (index <= High(fields)) then
  begin
    Result := fields[index];
  end
  else
    Result := '';
end;





end.
