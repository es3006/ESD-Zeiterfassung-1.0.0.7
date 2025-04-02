unit uSettings;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.Buttons, IniFiles,
  Vcl.ComCtrls, FireDAC.Stan.ExprFuncs, FireDAC.Phys.SQLiteDef, FireDAC.Stan.Intf,
  FireDAC.Stan.Option, FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf,
  FireDAC.Stan.Def, FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys,
  FireDAC.VCLUI.Wait, FireDAC.Stan.Param, FireDAC.DatS, FireDAC.DApt.Intf,
  FireDAC.DApt, FireDAC.Comp.DataSet, FireDAC.Comp.Client, FireDAC.Phys.SQLite,
  Data.DB, DateUtils, Vcl.Imaging.pngimage, Vcl.Mask, System.UITypes, Vcl.Menus,
  VCL.FlexCel.Core, FlexCel.XlsAdapter, AdvEdit, AdvEdBtn, AdvFileNameEdit;

type
  TfSettings = class(TForm)
    OpenDialog1: TOpenDialog;
    lbDienstplanHinweis: TLabel;
    edDienstplan: TLabeledEdit;
    Bevel3: TBevel;
    edObjektname: TLabeledEdit;
    edObjektNr: TLabeledEdit;
    edJahr: TLabeledEdit;
    btnSaveDienstplanSettings: TButton;
    sbLoadDienstplan: TSpeedButton;
    procedure FormCreate(Sender: TObject);
    procedure sbLoadDienstplanClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnSaveDienstplanSettingsClick(Sender: TObject);
    procedure edDienstplanClickBtn(Sender: TObject);
  private
    procedure LoadStart;
    procedure LeseUndPruefeWerte(const DateiPfad: string; var Zeilen, Jahr: integer; var Objektname, Objektnr: string);
  public
    { Public-Deklarationen }
  end;

var
  fSettings: TfSettings;
  SelLegendeEntry: integer;
  offizielleExceldatei: boolean;
  NEWDPLOADED: boolean;


implementation

{$R *.dfm}

uses
  uMain, uFunktionen, uExcelFunktionen, uDBFunktionen, uSettingsHelp, uLegende;






procedure TfSettings.edDienstplanClickBtn(Sender: TObject);
var
  Zeilen, Jahr: integer;
  Objektname, Objektnr: string;
begin
  edObjektname.Clear;
  edObjektNr.Clear;
  edJahr.Clear;

  //if OpenDialog1.Execute then
  begin
    edDienstplan.Text := OpenDialog1.FileName;
    LeseUndPruefeWerte(OpenDialog1.FileName, Zeilen, Jahr, Objektname, Objektnr);

    edObjektname.Text := Objektname;
    edObjektNr.Text := ObjektNr;
    edJahr.Text := IntToStr(Jahr);
  end;
end;


procedure TfSettings.btnSaveDienstplanSettingsClick(Sender: TObject);
var
  FDQuery: TFDQuery;
  ZEILEVON, ZEILEBIS, BEMZEILEVON, BEMZEILEBIS: integer;
  ZEILEN, JAHR, SPALTEPN, SPALTEMA, SPALTEVON, SPALTEBIS, BEMSPALTEVON, BEMSPALTEBIS: integer;
  Objektname, Objektnr: string;
begin
  ZEILEVON       := 0;
  ZEILEBIS       := 0;
  BEMZEILEVON    := 0;
  BEMZEILEBIS    := 0;
  SPALTEPN       := 0;
  SPALTEMA       := 0;
  SPALTEVON      := 0;
  SPALTEBIS      := 0;
  BEMSPALTEVON   := 0;
  BEMSPALTEBIS   := 0;
  OFZEXLDATEI    := 0;
  OFZEXLDATEIART := 0;

  EXCELDATEI := trim(edDienstplan.Text);
  if(EXCELDATEI <> '') AND FileExists(EXCELDATEI) then
  begin
    try
      LeseUndPruefeWerte(EXCELDATEI, Zeilen, Jahr, Objektname, Objektnr);

      if(Zeilen = 20) then
      begin
        ZEILEVON       := 10;
        ZEILEBIS       := 29;
        BEMZEILEVON    := 68;
        BEMZEILEBIS    := 76;
        SPALTEPN       := 1;
        SPALTEMA       := 2;
        SPALTEVON      := 3;
        SPALTEBIS      := 33;
        BEMSPALTEVON   := 2;
        BEMSPALTEBIS   := 14;
        OFZEXLDATEI    := 1;
        OFZEXLDATEIART := 1;
      end
      else if(Zeilen = 30) then
      begin
        ZEILEVON       := 10;
        ZEILEBIS       := 39;
        BEMZEILEVON    := 68;
        BEMZEILEBIS    := 76;
        SPALTEPN       := 1;
        SPALTEMA       := 2;
        SPALTEVON      := 3;
        SPALTEBIS      := 33;
        BEMSPALTEVON   := 2;
        BEMSPALTEBIS   := 14;
        OFZEXLDATEI    := 1;
        OFZEXLDATEIART := 2;
      end
      else if(Zeilen = 40) then
      begin
        ZEILEVON       := 10;
        ZEILEBIS       := 49;
        BEMZEILEVON    := 68;
        BEMZEILEBIS    := 76;
        SPALTEPN       := 1;
        SPALTEMA       := 2;
        SPALTEVON      := 3;
        SPALTEBIS      := 33;
        BEMSPALTEVON   := 2;
        BEMSPALTEBIS   := 14;
        OFZEXLDATEI    := 1;
        OFZEXLDATEIART := 3;
      end
      else
      begin
        showmessage('Sie benutzen keine offizielle Dienstplanvorlage');
        ZEILEVON       := 0;
        ZEILEBIS       := 0;
        BEMZEILEVON    := 0;
        BEMZEILEBIS    := 0;
        SPALTEPN       := 0;
        SPALTEMA       := 0;
        SPALTEVON      := 0;
        SPALTEBIS      := 0;
        BEMSPALTEVON   := 0;
        BEMSPALTEBIS   := 0;
        OFZEXLDATEI    := 0;
        OFZEXLDATEIART := 0;
        exit;
      end;

      edJahr.Text       := IntToStr(JAHR);
      edObjektName.Text := Objektname;
      edObjektNr.Text   := Objektnr;
    except
      on E: Exception do
        ShowMessage('Fehler: ' + E.Message);
    end;


    if(OFZEXLDATEI = 1) then
    begin
      FDQuery := TFDquery.Create(nil);
      try
        with FDQuery do
        begin
          Connection := fMain.FDConnection1;

          if FIRSTSTART = false then
          begin
            SQL.Text := '''
              UPDATE einstellungen SET
              Jahr = :JAHR,
              Exceldatei = :EXCELDATEI,
              Objektname = :OBJEKTNAME,
              ObjektNr = :OBJEKTNR,
              OffizielleDienstplanExcelDatei = :OFZEXLDATEI,
              OffizielleDienstplanExcelDateiArt = :OFZEXLDATEIART,
              ExcelPnSpalte = :SPALTEPN,
              ExcelMaSpalte = :SPALTEMA,
              ExcelDienstSpalteVon = :SPALTEVON,
              ExcelDienstSpalteBis = :SPALTEBIS,
              ExcelZeileVon = :ZEILEVON,
              ExcelZeileBis = :ZEILEBIS,
              ExcelBemerkungenSpalteVon = :BEMSPALTEVON,
              ExcelBemerkungenSpalteBis = :BEMSPALTEBIS,
              ExcelBemerkungenZeileVon = :BEMZEILEVON,
              ExcelBemerkungenZeileBis = :BEMZEILEBIS;
            ''';
            ParamByName('JAHR').AsInteger           := JAHR;
            ParamByName('EXCELDATEI').AsString      := EXCELDATEI;
            ParamByName('OBJEKTNAME').AsString      := OBJEKTNAME;
            ParamByName('OBJEKTNR').AsString        := OBJEKTNR;
            ParamByName('OFZEXLDATEI').AsInteger    := OFZEXLDATEI;
            ParamByName('OFZEXLDATEIART').AsInteger := OFZEXLDATEIART;  //1 = 20 Zeilen, 2 = 30 Zeilen, 3 = 40 Zeilen
            ParamByName('SPALTEPN').AsInteger       := SPALTEPN;
            ParamByName('SPALTEMA').AsInteger       := SPALTEMA;
            ParamByName('SPALTEVON').AsInteger      := SPALTEVON;
            ParamByName('SPALTEBIS').AsInteger      := SPALTEBIS;
            ParamByName('ZEILEVON').AsInteger       := ZEILEVON;
            ParamByName('ZEILEBIS').AsInteger       := ZEILEBIS;
            ParamByName('BEMSPALTEVON').AsInteger   := BEMSPALTEVON;
            ParamByName('BEMSPALTEBIS').AsInteger   := BEMSPALTEBIS;
            ParamByName('BEMZEILEVON').AsInteger    := BEMZEILEVON;
            ParamByName('BEMZEILEBIS').AsInteger    := BEMZEILEBIS;
            try
              ExecSQL;
              FIRSTSTART := false;
              LoadSettingsFromDB(JAHR);
             except
              on E: Exception do
                ShowMessage('Fehler beim Ändern der Dienstplan-Einstellungen in der Datenbank: ' + E.Message);
            end;
          end //FIRSTSTART = FALSE
        else
        begin
          SQL.Text := '''
            INSERT INTO einstellungen
            (
              Jahr,
              Exceldatei,
              Objektname,
              ObjektNr,
              OffizielleDienstplanExcelDatei,
              OffizielleDienstplanExcelDateiArt,
              ExcelPnSpalte,
              ExcelMaSpalte,
              ExcelDienstSpalteVon,
              ExcelDienstSpalteBis,
              ExcelZeileVon,
              ExcelZeileBis,
              ExcelBemerkungenSpalteVon,
              ExcelBemerkungenSpalteBis,
              ExcelBemerkungenZeileVon,
              ExcelBemerkungenZeileBis
            )
            VALUES
            (
              :JAHR,
              :EXCELDATEI,
              :OBJEKTNAME,
              :OBJEKTNR,
              :OFZEXLDATEI,
              :OFZEXLDATEIART,
              :SPALTEPN,
              :SPALTEMA,
              :SPALTEVON,
              :SPALTEBIS,
              :ZEILEVON,
              :ZEILEBIS,
              :BEMSPALTEVON,
              :BEMSPALTEBIS,
              :BEMZEILEVON,
              :BEMZEILEBIS
            );
          ''';
          ParamByName('JAHR').AsInteger           := Jahr;
          ParamByName('EXCELDATEI').AsString      := EXCELDATEI;
          ParamByName('OBJEKTNAME').AsString      := OBJEKTNAME;
          ParamByName('OBJEKTNR').AsString        := OBJEKTNR;
          ParamByName('OFZEXLDATEI').AsInteger    := OFZEXLDATEI;
          ParamByName('OFZEXLDATEIART').AsInteger := OFZEXLDATEIART;
          ParamByName('SPALTEPN').AsInteger       := SPALTEPN;
          ParamByName('SPALTEMA').AsInteger       := SPALTEMA;
          ParamByName('SPALTEVON').AsInteger      := SPALTEVON;
          ParamByName('SPALTEBIS').AsInteger      := SPALTEBIS;
          ParamByName('ZEILEVON').AsInteger       := ZEILEVON;
          ParamByName('ZEILEBIS').AsInteger       := ZEILEBIS;
          ParamByName('BEMSPALTEVON').AsInteger   := BEMSPALTEVON;
          ParamByName('BEMSPALTEBIS').AsInteger   := BEMSPALTEBIS;
          ParamByName('BEMZEILEVON').AsInteger    := BEMZEILEVON;
          ParamByName('BEMZEILEBIS').AsInteger    := BEMZEILEBIS;
          try
            ExecSQL;
            LoadSettingsFromDB(JAHR);
          except
            on E: Exception do
              ShowMessage('Fehler beim Hinzufügen der Dienstplan-Einstellungen in die Datenbank: ' + E.Message);
          end;
        end; //FIRSTSTART = TRUE
      end; //with FDQuery do
    finally
      FDQuery.free;
    end;
  end
  else
  begin
    showmessage('Bitte füllen Sie alle Eingabefelder aus!');
    exit;
  end;

  LoadSettingsFromDB(JAHR);

  if(NEWDPLOADED = true) then
  begin
    if MessageDlg('Um das Programm nutzen zu können, müssen im nächsten Schritt alle Kürzel aus dem Dienstplan in der Legende eingetragen werden.'+#13#10+#13#10+'Wollen Sie die Kürzel jetzt eintragen?', mtConfirmation, [mbYes, mbNo], 0, mbYes) = mrYes then
    begin
      fLegende.LEGENDELADEN := true;
      fLegende.Show;
    end
    else
    begin
      fLegende.LEGENDELADEN := false;
    end;
  end;

  LoadExcelBlattNamenToComboBox(EXCELDATEI, fMain.cbBlattnamen);

  fMain.lvZeiterfassung.Items.Clear;
  fMain.cbBlattnamen.ItemIndex := -1;
  fMain.cbBlattnamen.SetFocus;

  FirstStart := False;

  close;
  end;
end;







procedure TfSettings.FormCreate(Sender: TObject);
begin
  NEWDPLOADED := false;
  lbDienstplanHinweis.Caption  := 'Wählen Sie hier die Exceldatei, in der die Dienstpläne aller Monate eines Jahres stehen.';
end;




procedure TfSettings.FormShow(Sender: TObject);
begin
  Jahr := YearOf(Now);
  LoadStart;
end;




procedure TfSettings.LoadStart;
begin
  edDienstplan.Text := EXCELDATEI;
  edObjektname.Text := OBJEKTNAME;
  edObjektNr.Text   := OBJEKTNR;
  edJahr.Text       := IntToStr(JAHR);
end;









procedure TfSettings.sbLoadDienstplanClick(Sender: TObject);
var
  Zeilen, Jahr: integer;
  Objektname, Objektnr: string;
begin
  NEWDPLOADED := true;

  edObjektname.Clear;
  edObjektNr.Clear;
  edJahr.Clear;

  OpenDialog1.InitialDir := PATH;

  if OpenDialog1.Execute then
  begin
    edDienstplan.Text := OpenDialog1.FileName;
    LeseUndPruefeWerte(OpenDialog1.FileName, Zeilen, Jahr, Objektname, Objektnr);

    if(offizielleExceldatei = false) then
    begin
      showmessage('Sie benutzen keine offizielle Dienstplan Exceldatei.'+#13#10+'Nutzen Sie stattdessen das Zeitverwaltungsprogramm Version 1.0.0.5, mit dem Sie jede beliebige Dienstplan Exceldatei verwenden können.');
    end;

    edObjektname.Text := Objektname;
    edObjektNr.Text   := ObjektNr;
    edJahr.Text       := IntToStr(Jahr);
  end;
end;




procedure TfSettings.LeseUndPruefeWerte(const DateiPfad: string; var Zeilen, Jahr: integer; var Objektname, Objektnr: string);
var
  Xls: TXlsFile;
  WertJahr, WertZeilen: integer;
  WertObjektname, WertObjektNr: string;
  SheetIndex, i: Integer;
  SheetFound: Boolean;
begin
  // Initialisiere die Excel-Datei
  Xls := TXlsFile.Create(DateiPfad);
  try
    // Standardmäßig setzen wir den Indikator auf "nicht gefunden"
    SheetFound := False;
    offizielleExceldatei := False;

    // Durchlaufe alle Blätter und prüfe, ob eines "Einstellungen" heißt
    for i := 1 to Xls.SheetCount do
    begin
      if Xls.GetSheetName(i) = 'Einstellungen' then
      begin
        SheetIndex := i;
        SheetFound := True;
        offizielleExceldatei := True;
        Break;
      end;
    end;

    // Wenn das Blatt "Einstellungen" nicht gefunden wurde, Exception werfen
    if not SheetFound then
    begin
      offizielleExceldatei := false;
     // raise Exception.Create('Das Blatt "Einstellungen" existiert nicht in der Excel-Datei!');
    end;


    if(SheetFound = True) then
    begin
      // Wechsle zum Blatt "Einstellungen", falls gefunden
      Xls.ActiveSheet := SheetIndex;

      // Lese die Werte aus den Zellen E3, J3 und P3
      if(Xls.GetCellValue(1, 1).ToString <> '') then
        WertZeilen     := StrToInt(Xls.GetCellValue(1, 1).ToString)  // Zelle A1
      else
        WertZeilen     := 0;

      if(Xls.GetCellValue(3, 2).ToString <> '') then
        WertJahr       := StrToInt(Xls.GetCellValue(3, 2).ToString)  // Zelle B3
      else
        WertJahr       := 0;


      WertObjektname := Xls.GetCellValue(3, 5).ToString;  // Zelle E3
      WertObjektNr   := Xls.GetCellValue(3, 11).ToString; // Zelle K3

      // Prüfe, ob alle erforderlichen Werte vorhanden sind
      if (WertObjektname.Trim <> '') and (WertObjektNr.Trim <> '') AND (WertJahr <> 0) AND (WertZeilen <> 0) then
      begin
        // Werte in Variablen speichern
        Jahr       := WertJahr;
        Zeilen     := WertZeilen;
        Objektname := WertObjektname;
        Objektnr   := WertObjektNr;
        offizielleExceldatei := True;
      end
      else
      begin
        offizielleExceldatei := False;
      //  raise Exception.Create('Eine oder mehrere der erforderlichen Zellen auf dem Blatt "Einstellungen" sind leer.');
      end;
    end;
  finally
    Xls.Free;
  end;
end;












end.
