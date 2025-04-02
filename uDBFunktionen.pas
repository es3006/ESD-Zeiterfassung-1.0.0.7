unit uDBFunktionen;

interface

uses
  Windows, Classes, Forms, SysUtils, Vcl.StdCtrls, Vcl.ComCtrls, Dialogs, Controls, ExtCtrls, DateUtils,
  Graphics, StrUtils, ShellApi, System.UITypes, System.Zip, System.IOUtils,
  FireDAC.Stan.Param, FireDAC.Phys.SQLite, Data.DB, FireDAC.Comp.DataSet, FireDAC.Comp.Client;


type
  TSchichtData = record
    Vertragsnr: string;
    Einsatz: string;
    Leistungsart: string;
    UhrzeitVon: string;
    UhrzeitBis: string;
  end;


procedure CreateDatabaseTables;
procedure BackupSQLiteTable(const TableName: string; const BackupDir: string);
procedure ImportSQLiteTable(const SQLFileName: string);
procedure LoadSettingsFromDB(Jahr: integer);
function GetSchichtData(Schicht: string): TSchichtData;
function SchichtExists(Schicht: string): Boolean;
procedure DeleteLegendeFromDB;


implementation

uses
  uMain, uFunktionen;










procedure CreateDatabaseTables;
var
  FDQuery: TFDQuery;
begin
  FDQuery := TFDQuery.Create(nil);
  try
    FDQuery.Connection := fMain.FDConnection1;


    FDQuery.SQL.Text := 'SELECT name FROM sqlite_master WHERE type="table" AND name="einstellungen"';
    FDQuery.Open();
    if not FDQuery.IsEmpty then
      Exit
    else
    begin
      //OffizielleDienstplanExcelDateiArt (0 = Keine Standard Excel Dienstplan Datei, 1 = 30 Zeilen, 2 = 40 Zeilen)
      FDQuery.SQL.Text := '''
        CREATE TABLE einstellungen
        (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          Jahr INTEGER,
          Exceldatei TEXT,
          ExcelPnSpalte INTEGER,
          ExcelMaSpalte INTEGER,
          ExcelDienstSpalteVon INTEGER,
          ExcelDienstSpalteBis INTEGER,
          ExcelZeileVon INTEGER,
          ExcelZeileBis INTEGER,
          ExcelBemerkungenSpalteVon INTEGER,
          ExcelBemerkungenSpalteBis INTEGER,
          ExcelBemerkungenZeileVon INTEGER,
          ExcelBemerkungenZeileBis INTEGER,
          Objektname TEXT,
          ObjektNr TEXT,
          OffizielleDienstplanExcelDatei INTEGER,
          OffizielleDienstplanExcelDateiArt INTEGER
        );
      ''';
      FDQuery.ExecSQL;
    end;




    FDQuery.SQL.Text := 'SELECT name FROM sqlite_master WHERE type="table" AND name="legende"';
    FDQuery.Open();
    if not FDQuery.IsEmpty then
      Exit
    else
    begin
      FDQuery.SQL.Text := '''
        CREATE TABLE legende
        (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          schicht TEXT,
          einsatz TEXT,
          leistungsart TEXT,
          vertragsnr TEXT,
          beschreibung TEXT,
          UhrzeitVon,
          UhrzeitBis
        );
      ''';
      FDQuery.ExecSQL;
    end;
  finally
    FDQuery.Free;
  end;
end;
























procedure BackupSQLiteTable(const TableName: string; const BackupDir: string);
var
  FDQuery: TFDQuery;
  SQLFile: TStringList;
  FileName: string;
  i: integer;
  Field: TField;
  FieldType: string;
  InsertSQL, CreateTableSQL: string;
begin
  FDQuery := TFDquery.Create(nil);
  SQLFile := TStringList.Create;
  try
    with FDQuery do
    begin
      Connection := fMain.FDConnection1;
      SQL.Text := Format('SELECT * FROM %s', [TableName]);
      Open;

      // Erzeuge die CREATE TABLE-Anweisung
      CreateTableSQL := Format('DROP TABLE IF EXISTS %s; CREATE TABLE %s (', [TableName, TableName]);
      for i := 0 to FDQuery.FieldCount - 1 do
      begin
        Field := FDQuery.Fields[i];

        // Überprüfen, ob es sich um die id-Spalte handelt
        if (Field.FieldName = 'id') and (Field.DataType in [ftInteger, ftAutoInc]) then
          FieldType := 'INTEGER PRIMARY KEY AUTOINCREMENT'
        else
        begin
        case Field.DataType of
            ftString, ftMemo, ftWideString, ftWideMemo:
              FieldType := 'TEXT';
            ftInteger, ftSmallint, ftWord, ftAutoInc:
              FieldType := 'INTEGER';
            ftFloat, ftCurrency, ftBCD:
              FieldType := 'REAL';
            ftDate, ftTime, ftDateTime, ftTimeStamp:
              FieldType := 'TEXT'; // SQLite speichert Datumswerte als TEXT
            else
              FieldType := 'BLOB';
          end;
        end;

        if i > 0 then
          CreateTableSQL := CreateTableSQL + ', ';

        CreateTableSQL := CreateTableSQL + Format('%s %s', [Field.FieldName, FieldType]);
      end;
      CreateTableSQL := CreateTableSQL + ');';
      SQLFile.Add(CreateTableSQL);

      // Füge die INSERT-Anweisungen hinzu
      FDQuery.First;
      while not FDQuery.Eof do
      begin
        InsertSQL := Format('INSERT INTO %s VALUES (', [TableName]);
        for i := 0 to FDQuery.FieldCount - 1 do
        begin
          if i > 0 then
            InsertSQL := InsertSQL + ', ';

          if FDQuery.Fields[i].IsNull then
            InsertSQL := InsertSQL + 'NULL'
          else
            InsertSQL := InsertSQL + QuotedStr(FDQuery.Fields[i].AsString);
        end;
        InsertSQL := InsertSQL + ');';
        SQLFile.Add(InsertSQL);
        FDQuery.Next;
      end;

      // Speichere die SQL-Anweisungen in einer Datei
      FileName := Format('%s.sql', [TableName]);
      SQLFile.SaveToFile(IncludeTrailingPathDelimiter(BackupDir) + FileName);
    end;
  finally
    FDQuery.Free;
    SQLFile.Free;
  end;
end;
















procedure ImportSQLiteTable(const SQLFileName: string);
var
  FDQuery: TFDQuery;
  SQLFile: TStringList;
  TableName, InsertSQL: string;
  HighestID: Integer;
begin
  FDQuery := TFDQuery.Create(nil);
  SQLFile := TStringList.Create;
  try
    // 1. Lade die SQL-Datei
    SQLFile.LoadFromFile(SQLFileName);

    // 2. Extrahiere den Tabellennamen aus der CREATE TABLE-Anweisung
    TableName := '';
    if SQLFile.Count > 0 then
    begin
      InsertSQL := SQLFile[0];
      if Pos('CREATE TABLE ', InsertSQL) = 1 then
      begin
        InsertSQL := Copy(InsertSQL, 14, MaxInt);
        TableName := Copy(InsertSQL, 1, Pos(' ', InsertSQL) - 1);
      end;
    end;

    // 3. Führe die SQL-Befehle aus
    FDQuery.Connection := fMain.FDConnection1;
    FDQuery.SQL.Text := SQLFile.Text;
    FDQuery.ExecSQL;

    // 4. Optional: Aktualisiere sqlite_sequence
    if TableName <> '' then
    begin
      FDQuery.SQL.Text := Format('SELECT MAX(id) FROM %s', [TableName]);
      FDQuery.Open;
      HighestID := FDQuery.Fields[0].AsInteger;

      FDQuery.SQL.Text := 'INSERT OR REPLACE INTO sqlite_sequence (name, seq) VALUES (:TableName, :Seq)';
      FDQuery.Params.ParamByName('TableName').AsString := TableName;
      FDQuery.Params.ParamByName('Seq').AsInteger := HighestID;
      FDQuery.ExecSQL;
    end;
  finally
    FDQuery.Free;
    SQLFile.Free;
  end;
end;











procedure LoadSettingsFromDB(Jahr: integer);
var
  FDQuery: TFDQuery;
begin
  FDQuery := TFDquery.Create(nil);
  try
    with FDQuery do
    begin
      Connection := fMain.FDConnection1;

      SQL.Text := '''
        SELECT Exceldatei, ExcelPnSpalte, ExcelMaSpalte, ExcelDienstSpalteVon, ExcelDienstSpalteBis,
        ExcelZeileVon, ExcelZeileBis, ExcelBemerkungenSpalteVon, ExcelBemerkungenSpalteBis,
        ExcelBemerkungenZeileVon, ExcelBemerkungenZeileBis, Objektname, ObjektNr,
        OffizielleDienstplanExcelDatei, OffizielleDienstplanExcelDateiArt
        FROM einstellungen
       ;
      ''';

      // WHERE jahr = :JAHR
      //  LIMIT 1
   //   Params.ParamByName('JAHR').AsInteger  := Jahr;
      Open;

      if(not FDQuery.IsEmpty) then
      begin
        OBJEKTNAME     := FieldByName('objektname').AsString;
        EXCELDATEI     := FieldByName('Exceldatei').AsString;
        ZEILEVON       := FieldByName('ExcelZeileVon').AsInteger;
        ZEILEBIS       := FieldByName('ExcelZeileBis').AsInteger;
        SPALTEPN       := FieldByName('ExcelPnSpalte').AsInteger;;
        SPALTEMA       := FieldByName('ExcelMaSpalte').AsInteger;
        SPALTEVON      := FieldByName('ExcelDienstSpalteVon').AsInteger;
        SPALTEBIS      := FieldByName('ExcelDienstSpalteBis').AsInteger;
        OBJEKTNR       := FieldByName('ObjektNr').AsString;
        BEMSPALTEVON   := FieldByName('ExcelBemerkungenSpalteVon').AsInteger;
        BEMSPALTEBIS   := FieldByName('ExcelBemerkungenSpalteBis').AsInteger;
        BEMZEILEVON    := FieldByName('ExcelBemerkungenZeileVon').AsInteger;
        BEMZEILEBIS    := FieldByName('ExcelBemerkungenZeileBis').AsInteger;
        OFZEXLDATEI    := FieldByName('OffizielleDienstplanExcelDatei').AsInteger;
        OFZEXLDATEIART := FieldByName('OffizielleDienstplanExcelDateiArt').AsInteger;

        with fMain.StatusBar1 do
        begin
          Panels[0].Text := 'ObjektNr: ' + OBJEKTNR;
          Panels[1].Text := OBJEKTNAME;
        end;
      end
      else
      begin
        exit;
      end;
    end;
  finally
    FDQuery.free;
  end;
end;





procedure DeleteLegendeFromDB;
var
  i, id: integer;
  FDQuery: TFDQuery;
begin
  FDQuery := TFDQuery.Create(nil);
  try
    with FDQuery do
    begin
      Connection := fMain.FDConnection1;
      SQL.Text := 'DELETE FROM legende;';
      try
        ExecSQL;
      except
        on E: Exception do
        begin
          ShowMessage('Fehler beim löschen der Legende aus der Datenbank: ' + E.Message);
        end;
      end;
    end;
  finally
    FDQuery.Free;
  end;
 end;




function SchichtExists(Schicht: string): Boolean;
var
  FDQuery: TFDQuery;
begin
 // Result := False; // Standardmäßig auf False setzen
  FDQuery := TFDQuery.Create(nil);
  try
    with FDQuery do
    begin
      Connection := fMain.FDConnection1;
      SQL.Text := 'SELECT COUNT(*) AS CNT FROM legende WHERE schicht = :SCHICHT';
      Params.ParamByName('SCHICHT').AsString := Schicht;
      Open;
      Result := FieldByName('CNT').AsInteger > 0; // True, wenn CNT > 0
    end;
  finally
    FDQuery.Free;
  end;
end;






{
  Einsatz, Leistungsart, UhrzeitVon und UhrzeitBis anhand des Schichtkürzels ermitteln
}
function GetSchichtData(Schicht: string): TSchichtData;
var
  FDQuery: TFDQuery;
  SchichtData: TSchichtData;
begin
  FDQuery := TFDQuery.Create(nil);
  try
    with FDQuery do
    begin
      Connection := fMain.FDConnection1;

      SQL.Text := 'SELECT einsatz, leistungsart, vertragsnr, uhrzeitvon, uhrzeitbis FROM legende WHERE schicht = :SCHICHT;';
      Params.ParamByName('SCHICHT').AsString := Schicht;
      Open;

      if not FDQuery.IsEmpty then
      begin
        SchichtData.Einsatz      := FieldByName('einsatz').AsString;
        SchichtData.Leistungsart := FieldByName('leistungsart').AsString;
        SchichtData.VertragsNr   := FieldByName('vertragsnr').AsString;
        SchichtData.UhrzeitVon   := FieldByName('uhrzeitvon').AsString;
        SchichtData.UhrzeitBis   := FieldByName('uhrzeitbis').AsString;
      end;
    end;
  finally
    FDQuery.Free;
  end;
  Result := SchichtData;
end;











end.
