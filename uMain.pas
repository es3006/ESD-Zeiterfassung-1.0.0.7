unit uMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtDlgs, Vcl.ComCtrls, AdvListV,
  Vcl.StdCtrls, AdvUtil, Vcl.Grids, AdvObj, BaseGrid, AdvGrid, AdvGridCSVPager,
  System.Generics.Collections, ComObj, System.IOUtils, FlexCel.Core, FlexCel.XlsAdapter, VCL.FlexCel.Core,
  FlexCel.Render, FlexCel.Preview, Vcl.Buttons, Vcl.Menus,
  Vcl.ExtCtrls, FireDAC.Stan.ExprFuncs, FireDAC.Phys.SQLiteDef, FireDAC.Stan.Intf,
  FireDAC.Stan.Option, FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf,
  FireDAC.Stan.Def, FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys,
  FireDAC.VCLUI.Wait, FireDAC.Stan.Param, FireDAC.DatS, FireDAC.DApt.Intf,
  FireDAC.DApt, FireDAC.Comp.DataSet, FireDAC.Comp.Client, FireDAC.Phys.SQLite,
  Data.DB, DateUtils, Vcl.Imaging.pngimage, System.UITypes,
  FireDAC.Phys.SQLiteWrapper.Stat, AdvPageControl;






type
  TfMain = class(TForm)
    lvZeiterfassung: TAdvListView;
    OpenTextFileDialog1: TOpenTextFileDialog;
    OpenDialog1: TOpenDialog;
    Panel1: TPanel;
    Panel2: TPanel;
    cbBlattnamen: TComboBox;
    Label2: TLabel;
    Label3: TLabel;
    MainMenu1: TMainMenu;
    Dazei1: TMenuItem;
    Beenden1: TMenuItem;
    Einstellungen1: TMenuItem;
    Einstellungen2: TMenuItem;
    FDPhysSQLiteDriverLink1: TFDPhysSQLiteDriverLink;
    FDConnection1: TFDConnection;
    Legende1: TMenuItem;
    Image1: TImage;
    Mitarbeiter1: TMenuItem;
    SaveDialog1: TSaveDialog;
    N1: TMenuItem;
    N2: TMenuItem;
    Mitarbeiterlisteexportieren1: TMenuItem;
    ExportLegende: TMenuItem;
    ImportLegende: TMenuItem;
    ExportMitarbeiter: TMenuItem;
    ImportMitarbeiter: TMenuItem;
    StatusBar1: TStatusBar;
    Einstellungen3: TMenuItem;
    ExportEinstellungen: TMenuItem;
    ImportEinstellungen: TMenuItem;
    N3: TMenuItem;
    Splitter1: TSplitter;
    gpBemerkungen: TGridPanel;
    AdvPageControl1: TAdvPageControl;
    AdvTabSheet1: TAdvTabSheet;
    mBemerkungenDienstplan: TMemo;
    Hinweis1: TMenuItem;
    Programmhinweis1: TMenuItem;
    Dienstplanvorlageladen1: TMenuItem;
    N25Zeilen1: TMenuItem;
    N25Zeilen2: TMenuItem;
    btnConvertListViewZuExcelZeiterfassung: TButton;
    N20ZeilenDienstplanVorlagespeichern1: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure cbBlattnamenSelect(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Legende1Click(Sender: TObject);
    procedure Beenden1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ExportLegendeClick(Sender: TObject);
    procedure ImportLegendeClick(Sender: TObject);
    procedure ExportEinstellungenClick(Sender: TObject);
    procedure ImportEinstellungenClick(Sender: TObject);
    procedure Einstellungen2Click(Sender: TObject);
    procedure Programmhinweis1Click(Sender: TObject);
    procedure N25Zeilen1Click(Sender: TObject);
    procedure N25Zeilen2Click(Sender: TObject);
    procedure btnConvertListViewZuExcelZeiterfassungClick(Sender: TObject);
    procedure N20ZeilenDienstplanVorlagespeichern1Click(Sender: TObject);
  private
    procedure ShowSelectedDienstplanInListView;
    procedure ConvertMonthYearFromBlattName(cb: TComboBox);
    procedure SaveFiles(MonatJahr: String; OBJEKTNAME: string; mBemerkungen: TMemo; xls2: TXlsFile);
  public

  end;




var
  fMain: TfMain;
  PATH, OBJEKTNAME, OBJEKTNR, EXCELDATEI: string;
  selMonth, selYear: integer;
  selSheet: string;
  JAHR, ZEILEVON, ZEILEBIS, SPALTEPN, SPALTEMA, SPALTEVON, SPALTEBIS: integer;
  BEMSPALTEVON, BEMSPALTEBIS, BEMZEILEVON, BEMZEILEBIS, OFZEXLDATEI, OFZEXLDATEIART: integer;
  DBHOST, DBUSR, DBPW, DBNAME, DBPROT, DBLIBLOC: string;
  FIRSTSTART, ENCRYPTDB: boolean;
  GlobalFormatSettings: TFormatSettings;





const
  ENCRYPTIONKEY = 'mdklwuje90321iks,2moijlwödmeu3290dnu2i1p,sdim1239';
  PROGRAMMVERSION = '1.0.0.7';
  LASTCHANGEDATE  = '15.03.2025';


implementation

{$R *.dfm}
{$R MyResources.RES}

uses uSettings, uFunktionen, uDBFunktionen, uLegende, uMitarbeiter, uExcelFunktionen,
  uAbout;




procedure TfMain.Legende1Click(Sender: TObject);
begin
  fLegende.Show;
end;






procedure TfMain.N20ZeilenDienstplanVorlagespeichern1Click(Sender: TObject);
var
  dateiname: string;
begin
  dateiname := PATH + 'Dienstplan Vorlage 20 Zeilen.xlsm';

  if(FileExists(dateiname)) then
  begin
    if MessageDlg('Die Datei "' + ExtractFilename(dateiname) + '" befindet sich bereits im Programmverzeichnis.'+#13#10+'Wollen Sie diese überschreiben?', mtConfirmation, [mbYes, mbNo], 0, mbYes) = mrYes then
    begin
      SaveResourceToFile('20ZEILENVORLAGE', dateiname);
      if(FileExists(dateiname)) then
        showmessage('Die Datei "' + ExtractFilename(dateiname) + '" wurde im Programmverzeichnis gespeichert');
    end
  end
  else
  begin
    SaveResourceToFile('30ZEILENVORLAGE', dateiname);
    if(FileExists(dateiname)) then
      showmessage('Die Datei "' + ExtractFilename(dateiname) + '" wurde im Programmverzeichnis gespeichert');
  end;
end;


procedure TfMain.N25Zeilen1Click(Sender: TObject);
var
  dateiname: string;
begin
  dateiname := PATH + 'Dienstplan Vorlage 30 Zeilen.xlsm';

  if(FileExists(dateiname)) then
  begin
    if MessageDlg('Die Datei "' + ExtractFilename(dateiname) + '" befindet sich bereits im Programmverzeichnis.'+#13#10+'Wollen Sie diese überschreiben?', mtConfirmation, [mbYes, mbNo], 0, mbYes) = mrYes then
    begin
      SaveResourceToFile('30ZEILENVORLAGE', dateiname);
      if(FileExists(dateiname)) then
        showmessage('Die Datei "' + ExtractFilename(dateiname) + '" wurde im Programmverzeichnis gespeichert');
    end
  end
  else
  begin
    SaveResourceToFile('30ZEILENVORLAGE', dateiname);
    if(FileExists(dateiname)) then
      showmessage('Die Datei "' + ExtractFilename(dateiname) + '" wurde im Programmverzeichnis gespeichert');
  end;
end;


procedure TfMain.N25Zeilen2Click(Sender: TObject);
var
  dateiname: string;
begin
  dateiname := PATH + 'Dienstplan Vorlage 40 Zeilen.xlsm';

  if(FileExists(dateiname)) then
  begin
    if MessageDlg('Die Datei "' + ExtractFilename(dateiname) + '" befindet sich bereits im Programmverzeichnis.'+#13#10+'Wollen Sie diese überschreiben?', mtConfirmation, [mbYes, mbNo], 0, mbYes) = mrYes then
    begin
      SaveResourceToFile('40ZEILENVORLAGE', dateiname);
      if(FileExists(dateiname)) then
        showmessage('Die Datei "' + ExtractFilename(dateiname) + '" wurde im Programmverzeichnis gespeichert');
    end
  end
  else
  begin
    SaveResourceToFile('40ZEILENVORLAGE', dateiname);
    if(FileExists(dateiname)) then
      showmessage('Die Datei "' + ExtractFilename(dateiname) + '" wurde im Programmverzeichnis gespeichert');
  end;
end;




procedure TfMain.Programmhinweis1Click(Sender: TObject);
begin
  fAbout.Show;
end;

procedure TfMain.Einstellungen2Click(Sender: TObject);
begin
  fSettings.Show;
end;

procedure TfMain.ExportLegendeClick(Sender: TObject);
begin
  BackupSQLiteTable('Legende', PATH + 'DBDUMP');
end;

procedure TfMain.ImportLegendeClick(Sender: TObject);
begin
  if FileExists(PATH + 'DBDUMP\Legende.sql') then
    ImportSQLiteTable(PATH + 'DBDUMP\Legende.sql');
end;

procedure TfMain.ExportEinstellungenClick(Sender: TObject);
begin
  BackupSQLiteTable('einstellungen', PATH + 'DBDUMP');
end;




procedure TfMain.ImportEinstellungenClick(Sender: TObject);
begin
  if FileExists(PATH + 'DBDUMP\Einstellungen.sql') then
  begin
    ImportSQLiteTable(PATH + 'DBDUMP\Einstellungen.sql');
   // LoadSettingsFromDB(YearOf(Now));
    if(cbBlattnamen.ItemIndex <> -1) then
      cbBlattnamenSelect(nil);
  end;
end;














procedure TfMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if FDConnection1.Connected then
    FDConnection1.Close; // Verbindung schließen
end;

procedure TfMain.FormCreate(Sender: TObject);
begin
  PATH := ExtractFilePath(ParamStr(0));
  ForceDirectories(PATH + 'DBDUMP');

  SaveResourceToFile('SQLITE3DLL64BIT', 'sqlite3.dll');
  SaveResourceToFile('SSLEAY32', 'ssleay32.dll');
  SaveResourceToFile('LIBEAY32', 'libeay32.dll');

  if not System.SysUtils.DirectoryExists('TEMP') then
  begin
    CreateDirectory(PChar('TEMP'), nil);
  end;

  DBNAME     := 'zeiterfassung.s3db';
  ENCRYPTDB  := false; //im Produktivmodus auf true setzen damit Datenbank verschlüsselt wird
  FIRSTSTART := false;


  //Wenn FirstStart abgebrochen wurde und so kein Admin angelegt werden konnte,
  //die Datenbankdatei löschen
  if(FileExists(DBNAME)) then
  begin
    if IsFileZeroSize(DBNAME) then
    begin
      if(FDConnection1.Connected) then FDConnection1.Connected := false;
      DeleteFile(DBNAME);
    end;
  end;

  //Wenn Datenbank noch nicht vorhanden Formular zum anlegen eines neuen Nutzers anzeigen
  if not FileExists(DBNAME) then
  begin
    FIRSTSTART := true;
  end;

  FDConnection1.Connected := False; // Sicherstellen, dass die Verbindung geschlossen ist
  FDConnection1.DriverName := 'SQLite';
  FDConnection1.Params.Values['Database'] := DBNAME; // Datenbankname/Pfad
  FDConnection1.Params.Values['CharacterSet'] := 'utf8';
  if(ENCRYPTDB) then
  begin
    FDConnection1.Params.Values['EncryptionMode'] := 'Aes128'; // Verschlüsselung (optional)
    FDConnection1.Params.Values['Password'] := ENCRYPTIONKEY; // Passwort (optional)
  end;

  FDConnection1.Connected := true;
  FDConnection1.Open(); // Verbindung öffnen
end;












procedure TfMain.FormDestroy(Sender: TObject);
begin
  if FDConnection1.Connected then
    FDConnection1.Close; // Verbindung schließen
end;






procedure TfMain.FormShow(Sender: TObject);
begin
  fMain.Caption := 'ESD Zeiterfassung ' + PROGRAMMVERSION + '    -     Letzte Änderung: ' + LASTCHANGEDATE;


  if(FIRSTSTART = true) then
  begin
    CreateDatabaseTables;
    fSettings.ShowModal;
  end
  else
  begin
    JAHR := YearOf(Now);
    LoadSettingsFromDB(JAHR);

    if(FileExists(EXCELDATEI)) then
    begin
      LoadExcelBlattNamenToComboBox(EXCELDATEI, cbBlattnamen);
    end
    else
    begin
      ShowMessage('Die Dienstplan Exceldatei wurde nicht gefunden!' + #13#10 +
                  'Sie können das Programm erst benutzen, wenn Sie eine ' + #13#10 +
                  'Dienstplan-Exceldatei in den Einstellungen angegeben ' + #13#10 +
                  'und die Legende mit den Schichtkürzeln und Einsatznummern ausgefüllt haben!');

      fSettings.Show;
      fSettings.edDienstplan.SetFocus;
    end;
  end;

  StatusBar1.Panels[0].Text := 'ObjektNr: ' + OBJEKTNR;
  StatusBar1.Panels[1].Text := OBJEKTNAME;
end;






procedure TfMain.ShowSelectedDienstplanInListView;
var
  xls: TXlsFile;
  TagEntries, NachtEntries, CombinedEntries: TStringList;
  i, j, colIndex: Integer;
  Mitarbeiter, PersonalNr, tagValue, currentDate, dateCaption: string;
  item: TListItem;
  day: string;
  SchichtData: TSchichtData;
  NAVISIONPatchBisDatum: string;   //Bis Datum bei 24 Stunden anpassen dass anstatt von 18:00:00 bis 18:00:00 bis 17:59:59 wird damit das NAVISION Programm im Büro funktioniert
begin
  TagEntries      := TStringList.Create;
  NachtEntries    := TStringList.Create;
  CombinedEntries := TStringList.Create;
  try
    xls := TXlsFile.Create(TPath.Combine(TPath.GetDocumentsPath, EXCELDATEI));
    try
      xls.ActiveSheetByName := selSheet;
      xls.IgnoreFormulaText := true;

      //Anzahl der eingetragenen Mitarbeiter in Spalte B ermitteln
      if(CountEntriesInColumnB(Xls, ZEILEVON, ZEILEBIS) = 0) then
      begin
        lvZeiterfassung.Items.Clear;
        ShowMessage('Im Dienstplan "' + selSheet + '" stehen keine Mitarbeiter!');
        exit;
      end;

      // Durch die Spalten der Exceldatei iterieren
      for colIndex := SPALTEVON to SPALTEBIS do
      begin
        // Durch die Zeilen der StringList iterieren
        for j := ZEILEVON to ZEILEBIS do // Zeilen B8 bis AG31
        begin
          Mitarbeiter := xls.GetCellValue(j, SPALTEMA).ToString; // Spalte B (2)
          PersonalNr  := xls.GetCellValue(j, SPALTEPN).ToString; // Spalte A (1)
          tagValue    := xls.GetCellValue(j, colIndex).ToString; // Schichtkürzel

          day := IntToStr(colIndex - 2); // Die Spalte 3 (C) entspricht Tag 1
          dateCaption := Format('%.2d.%.2d.%.4d', [StrToInt(day), selMonth, selYear]);

          if SchichtExists(tagValue) then
          begin
            SchichtData := GetSchichtData(tagValue);

            if(SchichtData.UhrzeitVon < SchichtData.UhrzeitBis) then
            begin
              //Tagschicht
              TagEntries.Add(Format('%s;%s;%s;%s;%s;%s;%s', [dateCaption, PersonalNr, SchichtData.UhrzeitVon, SchichtData.UhrzeitBis, SchichtData.Vertragsnr, SchichtData.Einsatz, SchichtData.Leistungsart]));
            end
            else if(SchichtData.UhrzeitVon > SchichtData.UhrzeitBis) then
            begin
              //Nachtschicht
              NachtEntries.Add(Format('%s;%s;%s;%s;%s;%s;%s', [dateCaption, PersonalNr, SchichtData.UhrzeitVon, SchichtData.UhrzeitBis, SchichtData.Vertragsnr, SchichtData.Einsatz, SchichtData.Leistungsart]));
            end
            else if(SchichtData.UhrzeitVon = SchichtData.UhrzeitBis) then
            begin
              //24 Std Dienst

              NAVISIONPatchBisDatum := SubtractOneSecond(SchichtData.UhrzeitBis);
              TagEntries.Add(Format('%s;%s;%s;%s;%s;%s;%s', [dateCaption, PersonalNr, SchichtData.UhrzeitVon, NAVISIONPatchBisDatum, SchichtData.Vertragsnr, SchichtData.Einsatz, SchichtData.Leistungsart]));
            //  NachtEntries.Add(Format('%s;%s;%s;%s;%s;%s;%s', [dateCaption, PersonalNr, SchichtData.UhrzeitVon, SchichtData.UhrzeitBis, SchichtData.Vertragsnr, SchichtData.Einsatz, SchichtData.Leistungsart]));
            end
          end;
        end;
      end;


      currentDate := '';

      // CombinedEntries erstellen
      // Enthält alle Tag und Nachtschichten eines Tages
      // Sortiert nach Tag und Nacht
      for colIndex := SPALTEVON to SPALTEBIS do
      begin
        day := IntToStr(colIndex - 2); // Die Spalte 3 (C) entspricht Tag 1
        dateCaption := Format('%.2d.%.2d.%.4d', [StrToInt(day), selMonth, selYear]);

        if currentDate <> dateCaption then
        begin
          currentDate := dateCaption;

          // Add TagEntries for the current date
          for i := 0 to TagEntries.Count - 1 do
          begin
            if ExtractField(TagEntries[i], ';', 0) = currentDate then
            begin
              CombinedEntries.Add(TagEntries[i]);
            end;
          end;

          // Add NachtEntries for the current date
          for i := 0 to NachtEntries.Count - 1 do
          begin
            if ExtractField(NachtEntries[i], ';', 0) = currentDate then
            begin
              CombinedEntries.Add(NachtEntries[i]);
            end;
          end;
        end;
      end;


      lvZeiterfassung.Items.Clear;

      // Add the CombinedEntries to the ListView
      for i := 0 to CombinedEntries.Count - 1 do
      begin
        item := lvZeiterfassung.Items.Add;
        item.Caption := ExtractField(CombinedEntries[i],   ';', 0); //Datum
        item.SubItems.Add(ExtractField(CombinedEntries[i], ';', 1)); //PersonalNr
        item.SubItems.Add(ExtractField(CombinedEntries[i], ';', 2)); //Uhrzeit Von
        item.SubItems.Add(ExtractField(CombinedEntries[i], ';', 3)); //Uhrzeit Bis
        item.SubItems.Add(ExtractField(CombinedEntries[i], ';', 4)); //BelegNr
        item.SubItems.Add(ExtractField(CombinedEntries[i], ';', 5)); //EinsatzNr
        item.SubItems.Add(ExtractField(CombinedEntries[i], ';', 6)); //Leistungsart
      end;
    finally
      xls.Free;
    end;
  finally
    TagEntries.Free;
    NachtEntries.Free;
    CombinedEntries.Free;
  end;
end;










procedure TfMain.ConvertMonthYearFromBlattName(cb: TComboBox);
var
  MonthStr: string;
  YearStr: integer;
begin
  MonthStr := cb.Text;
  YearStr  := JAHR;
  selYear  := YearStr;

  // Monatsnamen in Zahl umwandeln
  if MonthStr = 'Januar' then
    selMonth := 1
  else if MonthStr = 'Februar' then
    selMonth := 2
  else if MonthStr = 'März' then
    selMonth := 3
  else if MonthStr = 'April' then
    selMonth := 4
  else if MonthStr = 'Mai' then
    selMonth := 5
  else if MonthStr = 'Juni' then
    selMonth := 6
  else if MonthStr = 'Juli' then
    selMonth := 7
  else if MonthStr = 'August' then
    selMonth := 8
  else if MonthStr = 'September' then
    selMonth := 9
  else if MonthStr = 'Oktober' then
    selMonth := 10
  else if MonthStr = 'November' then
    selMonth := 11
  else if MonthStr = 'Dezember' then
    selMonth := 12
  else
    selMonth := 0;
end;












procedure TfMain.SaveFiles(MonatJahr: string; OBJEKTNAME: string; mBemerkungen: TMemo; xls2: TXlsFile);
var
  SaveDialog: TSaveDialog;
  FileName: string;
  FilePath: string;
  PATH: string;
  Bemerkungen: TStringList;
begin
  PATH := 'C:\'; // Definieren Sie den Pfad entsprechend Ihren Anforderungen

  // Erstellen und konfigurieren Sie den SaveDialog
  SaveDialog := TSaveDialog.Create(nil);
  Bemerkungen := TStringList.Create;
  try
    SaveDialog.Filter := 'Excel Files (*.xlsx)|*.xlsx|Text Files (*.txt)|*.txt';
    SaveDialog.DefaultExt := 'xlsx';
    SaveDialog.FileName := Format('SHD Zeiterfassung %s - %s.xlsx', [MonatJahr, OBJEKTNAME]);

    // Zeigen Sie den Dialog an und überprüfen Sie, ob der Benutzer auf 'Speichern' geklickt hat
    if SaveDialog.Execute then
    begin
      // Holen Sie sich den vom Benutzer gewählten Pfad
      FileName := SaveDialog.FileName;

      // Überprüfen, ob die Datei bereits existiert
      if FileExists(FileName) then
      begin
        // Benutzer fragen, ob die Datei überschrieben werden soll
        if MessageDlg(Format('Die Datei %s existiert bereits. Möchten Sie diese überschreiben?', [FileName]),
          mtConfirmation, [mbYes, mbNo], 0) = mrNo then
        begin
          Exit; // Abbrechen, falls der Benutzer nicht überschreiben möchte
        end;
      end;

      // Speichern Sie die Excel-Datei
      xls2.Save(FileName);

      // Speichern Sie die Bemerkungen als Textdatei im selben Verzeichnis
      FilePath := TPath.ChangeExtension(FileName, '.txt');

      Bemerkungen.Add(mBemerkungen.Text);
      Bemerkungen.SaveToFile(FilePath);
    end;
  finally
    SaveDialog.Free;
    Bemerkungen.Free;
  end;
end;





procedure TfMain.Beenden1Click(Sender: TObject);
begin
  close;
end;







procedure TfMain.btnConvertListViewZuExcelZeiterfassungClick(Sender: TObject);
var
  xls2: TXlsFile;
  zeile, i: Integer;
begin
  if(cbBlattnamen.ItemIndex > -1) then
  begin
    if(lvZeiterfassung.Items.Count > 0) then
    begin
      PlayResourceMP3('CLICK', 'TEMP\click.wav');

      zeile := 1;

      xls2 := TXlsFile.Create(1, TExcelFileFormat.v2019, true);
      try
        xls2.SetCellValue(1, 1, 'Datum');
        xls2.SetCellValue(1, 2, 'Mitarbeiter');
        xls2.SetCellValue(1, 3, 'Von Uhrzeit');
        xls2.SetCellValue(1, 4, 'Bis Uhrzeit');
        xls2.SetCellValue(1, 5, 'BelegNr');
        xls2.SetCellValue(1, 6, 'Einsatz');
        xls2.SetCellValue(1, 7, 'Leistungsart');

        if Assigned(lvZeiterfassung) then
        begin
          for i := 0 to lvZeiterfassung.Items.Count - 1 do
          begin
            inc(zeile);

            xls2.SetCellValue(zeile, 1, lvZeiterfassung.Items[i].Caption);
            if lvZeiterfassung.Items[i].SubItems.Count >= 6 then
            begin
              xls2.SetCellValue(zeile, 2, lvZeiterfassung.Items[i].SubItems[0]);
              xls2.SetCellValue(zeile, 3, lvZeiterfassung.Items[i].SubItems[1]);
              xls2.SetCellValue(zeile, 4, lvZeiterfassung.Items[i].SubItems[2]);
              xls2.SetCellValue(zeile, 5, lvZeiterfassung.Items[i].SubItems[3]);
              xls2.SetCellValue(zeile, 6, lvZeiterfassung.Items[i].SubItems[4]);
              xls2.SetCellValue(zeile, 7, lvZeiterfassung.Items[i].SubItems[5]);
            end;
          end;
        end;

        SaveFiles(cbBlattnamen.Text, OBJEKTNAME, mBemerkungenDienstplan, xls2);

      finally
        xls2.Free;
      end;
    end
    else
    begin
      PlayResourceMP3('WRONGPW', 'TEMP\wrongpw.WAV');
      showmessage('In der Liste befinden sich keine Daten!'+#13#10+'Wurden die Schichtkürzel aus dem Dienstplan in der Legende eingetragen?');
      exit;
    end;
  end
  else
  begin
    PlayResourceMP3('WRONGPW', 'TEMP\wrongpw.WAV');
    showmessage('Bitte wählen Sie erst einen Monat aus?');
    exit;
  end;
end;









procedure TfMain.cbBlattnamenSelect(Sender: TObject);
var
  Output: TStringList;
  i: integer;
begin
  i := cbBlattnamen.ItemIndex;
  if(i <> -1) then
  begin
    selSheet := cbBlattnamen.Items[i];

    PlayResourceMP3('WHOOSH', 'TEMP\whoosh.wav');

    ConvertMonthYearFromBlattName(cbBlattnamen); //Juli 2024

    ShowSelectedDienstplanInListView;

    //Alle Namen aus dem Dienstplan ändern durch die PersonalNummern aus der Datenbank
  //  ReplaceNamesWithPersonalNr;

    Output := TStringList.Create;
    try
      // Die Zellen im Bereich A40 bis S50 auslesen
      ReadExcelRange(EXCELDATEI, BEMZEILEVON, BEMSPALTEVON, BEMZEILEBIS, BEMSPALTEBIS, Output); // A=1, S=19
      mBemerkungenDienstplan.Lines.Assign(Output);
    finally
      Output.Free;
    end;
  end;
end;











end.
