unit uLegende;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls, AdvListV, Vcl.StdCtrls,
  Vcl.ExtCtrls, FireDAC.Stan.ExprFuncs, FireDAC.Phys.SQLiteDef, FireDAC.Stan.Intf,
  FireDAC.Stan.Option, FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf,
  FireDAC.Stan.Def, FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys,
  FireDAC.VCLUI.Wait, FireDAC.Stan.Param, FireDAC.DatS, FireDAC.DApt.Intf,
  FireDAC.DApt, FireDAC.Comp.DataSet, FireDAC.Comp.Client, FireDAC.Phys.SQLite,
  Data.DB, DateUtils, System.UITypes, Vcl.Mask, Vcl.Menus, IdHTTP, System.JSON, IdSSLOpenSSL,
  System.Generics.Collections;

type
  TfLegende = class(TForm)
    lbSchicht: TLabeledEdit;
    lbEinsatz: TLabeledEdit;
    lbLeistungsart: TLabeledEdit;
    lbBeschreibung: TLabeledEdit;
    btnNewLegende: TButton;
    btnAddLegende: TButton;
    lvLegende: TAdvListView;
    edUhrzeitVon: TLabeledEdit;
    edUhrzeitBis: TLabeledEdit;
    StatusBar1: TStatusBar;
    Panel1: TPanel;
    lbHinweis: TLabel;
    MainMenu1: TMainMenu;
    Legende1: TMenuItem;
    Exportieern1: TMenuItem;
    Importieren1: TMenuItem;
    lbVertragsNr: TLabeledEdit;
    btnLoadLegendeFromWeb: TButton;
    procedure lvLegendeSelectItem(Sender: TObject; Item: TListItem; Selected: Boolean);
    procedure btnNewLegendeClick(Sender: TObject);
    procedure btnAddLegendeClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure lvLegendeRightClickCell(Sender: TObject; iItem, iSubItem: Integer);
    procedure Exportieern1Click(Sender: TObject);
    procedure Importieren1Click(Sender: TObject);
    procedure btnLoadLegendeFromWebClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    procedure LoadLegendeInListView(lv: TListView);
    procedure InsertLegendeEntry;
    procedure UpdateLegendeEntry(id: integer);
    procedure InsertDataIntoSQLite(const ObjektNr: string);
    procedure LoadDataFromSQLite(const ObjektNr: string);
  public
    LEGENDELADEN: boolean;
  end;

var
  fLegende: TfLegende;
  SelLegendeEntry: integer;



implementation

{$R *.dfm}

uses uMain, uDBFunktionen;




procedure TfLegende.UpdateLegendeEntry(id: integer);
var
  FDQuery: TFDQuery;
  i: integer;
begin
  i := lvLegende.ItemIndex;
  if i <> -1 then
  begin
    FDQuery := TFDquery.Create(nil);
    try
      with FDQuery do
      begin
        Connection := fMain.FDConnection1;

        SQL.Text := 'UPDATE legende SET schicht = :SCHICHT, einsatz = :EINSATZ, ' +
                    'leistungsart = :LEISTUNGSART, vertragsnr = :VERTRAGSNR, beschreibung = :BESCHREIBUNG, UhrzeitVon = :UHRZEITVON, ' +
                    'UhrzeitBis = :UHRZEITBIS WHERE id = :ID;';
        Params.ParamByName('ID').AsInteger          := SelLegendeEntry;
        Params.ParamByName('SCHICHT').AsString      := lbSchicht.Text;
        Params.ParamByName('EINSATZ').AsString      := lbEinsatz.Text;
        Params.ParamByName('LEISTUNGSART').AsString := lbLeistungsart.Text;
        Params.ParamByName('VERTRAGSNR').AsString   := lbVertragsNr.Text;
        Params.ParamByName('BESCHREIBUNG').AsString := lbBeschreibung.Text;
        Params.ParamByName('UHRZEITVON').AsString   := edUhrzeitVon.Text;
        Params.ParamByName('UHRZEITBIS').AsString   := edUhrzeitBis.Text;
        try
          ExecSQL;

          lvLegende.Items[i].SubItems[0] := lbSchicht.Text;
          lvLegende.Items[i].SubItems[1] := lbEinsatz.Text;
          lvLegende.Items[i].SubItems[2] := lbLeistungsart.Text;
          lvLegende.Items[i].SubItems[3] := lbVertragsNr.Text;
          lvLegende.Items[i].SubItems[4] := lbBeschreibung.Text;
          lvLegende.Items[i].SubItems[5] := edUhrzeitVon.Text;
          lvLegende.Items[i].SubItems[6] := edUhrzeitBis.Text;
        except
          on E: Exception do
          begin
            ShowMessage('Fehler beim �ndern des Wertes in der Tabelle legende: ' + E.Message);
          end;
        end;
      end;
    finally
      FDQuery.Free;
    end;
  end;
end;




procedure TfLegende.Importieren1Click(Sender: TObject);
begin
  if FileExists(PATH + 'DBDUMP\Legende.sql') then
  begin
    ImportSQLiteTable(PATH + 'DBDUMP\Legende.sql');
    LoadLegendeInListView(lvLegende);
  end;
end;






procedure TfLegende.InsertLegendeEntry;
var
  FDQuery: TFDQuery;
  IDQuery: TFDQuery;
  l: TListItem;
  insertedID: Integer;
begin
  insertedID := -1;
  FDQuery    := TFDQuery.Create(nil);
  IDQuery    := TFDQuery.Create(nil);
  try
    with FDQuery do
    begin
      Connection := fMain.FDConnection1;

      SQL.Text := 'INSERT INTO legende (schicht, einsatz, leistungsart, vertragsnr, beschreibung, uhrzeitVon, uhrzeitBis) ' +
                  'VALUES (:SCHICHT, :EINSATZ, :LEISTUNGSART, :VERTRAGSNR, :BESCHREIBUNG, :UHRZEITVON, :UHRZEITBIS);';
      Params.ParamByName('SCHICHT').AsString      := lbSchicht.Text;
      Params.ParamByName('EINSATZ').AsString      := lbEinsatz.Text;
      Params.ParamByName('LEISTUNGSART').AsString := lbLeistungsart.Text;
      Params.ParamByName('VERTRAGSNR').AsString   := lbVertragsNr.Text;
      Params.ParamByName('BESCHREIBUNG').AsString := lbBeschreibung.Text;
      Params.ParamByName('UHRZEITVON').AsString   := edUhrzeitVon.Text;
      Params.ParamByName('UHRZEITBIS').AsString   := edUhrzeitBis.Text;

      try
        ExecSQL;

        // Ermitteln der zuletzt eingef�gten ID
        with IDQuery do
        begin
          Connection := fMain.FDConnection1;
          SQL.Text := 'SELECT last_insert_rowid() as last_id;';
          Open;
          insertedID := FieldByName('last_id').AsInteger;
        end;

      except
        on E: Exception do
        begin
          ShowMessage('Fehler beim Einf�gen in die Tabelle legende: ' + E.Message);
        end;
      end;
    end;
  finally
    FDQuery.Free;
    IDQuery.Free;

    l := lvLegende.Items.Add;
    l.Caption := IntToStr(insertedID);
    l.SubItems.Add(lbSchicht.Text);
    l.SubItems.Add(lbEinsatz.Text);
    l.SubItems.Add(lbLeistungsart.Text);
    l.SubItems.Add(lbVertragsNr.Text);
    l.SubItems.Add(lbBeschreibung.Text);
    l.SubItems.Add(edUhrzeitVon.Text);
    l.SubItems.Add(edUhrzeitBis.Text);
  end;
end;







procedure TfLegende.btnAddLegendeClick(Sender: TObject);
var
  schicht, einsatz, leistungsart, vertragsnr, beschreibung, uhrzeitvon, uhrzeitbis: string;
begin
  schicht      := trim(lbSchicht.Text);
  einsatz      := trim(lbeinsatz.Text);
  leistungsart := trim(lbLeistungsart.Text);
  vertragsnr   := trim(lbVertragsNr.Text);
  beschreibung := trim(lbBeschreibung.Text);
  uhrzeitvon   := trim(edUhrzeitVon.Text);
  uhrzeitbis   := trim(edUhrzeitBis.Text);

  if(schicht <> '') AND (einsatz <> '') AND (leistungsart <> '') AND (vertragsnr <> '') AND (beschreibung <> '') then
  begin
    if(SelLegendeEntry = -1) then
    begin
      InsertLegendeEntry;
    end
    else if(SelLegendeEntry <> -1) then
    begin
      UpdateLegendeEntry(SelLegendeEntry);
    end;
  end
  else
  begin
    showmessage('Bitte f�llen Sie alle Eingabefelder aus!');
  end;
  btnNewLegendeClick(nil);
end;





procedure TfLegende.btnNewLegendeClick(Sender: TObject);
begin
  lvLegende.ItemIndex := -1;
  lbSchicht.Clear;
  lbEinsatz.Clear;
  lbLeistungsart.Clear;
  lbVertragsNr.Clear;
  lbBeschreibung.Clear;
  edUhrzeitVon.Clear;
  edUhrzeitBis.Clear;
  btnAddLegende.Caption := 'Hinzuf�gen';
  lbSchicht.SetFocus;
  SelLegendeEntry := -1;
end;






procedure TfLegende.btnLoadLegendeFromWebClick(Sender: TObject);
begin
  InsertDataIntoSQLite(OBJEKTNR);  // F�ge die Daten aus dem Web in die SQLite-Datenbank ein
  LoadDataFromSQLite(OBJEKTNR);     // Lade die Daten in die ListView
  btnLoadLegendeFromWeb.Visible := false;

 // showmessage('Bitte erg�nzen Sie alle Eintr�ge durch Eingabe des Schichtk�rzels, so wie dieses im Dienstplan angegeben ist, damit die Dienste im Dienstplan gefunden werden k�nnen.');
end;

procedure TfLegende.Exportieern1Click(Sender: TObject);
begin
  BackupSQLiteTable('Legende', PATH + 'DBDUMP');
  showmessage('Legende unter dem Namen "Legende.sql" im Verzeichnis "DBDUMP" gespeichert');
end;







procedure TfLegende.FormCreate(Sender: TObject);
begin
  LEGENDELADEN := false;
end;

procedure TfLegende.FormShow(Sender: TObject);
begin
  SelLegendeEntry := -1;
  btnNewLegendeClick(nil);

  if(LEGENDELADEN = true) then
  begin
    DeleteLegendeFromDB;

   // lvLegende.BeginUpdate;
    lvLegende.Items.Clear;
   // lvLegende.EndUpdate;

    btnLoadLegendeFromWeb.Visible := true;
    btnLoadLegendeFromWebClick(Self);

    LEGENDELADEN := false;
  end
  else
  begin
    LoadLegendeInListView(lvLegende);
  end;

  lvLegende.SortColumn := 1;
  lvLegende.AlphaSort;
  lvLegende.Sort;

  if(lvLegende.Items.Count <= 0) then
  begin
    btnLoadLegendeFromWeb.Visible := true;
  end;
end;






procedure TfLegende.lvLegendeRightClickCell(Sender: TObject; iItem, iSubItem: Integer);
var
  i, id: integer;
  FDQuery: TFDQuery;
begin
  if MessageDlg('Wollen Sie diesen Eintrag wirklich l�schen?', mtConfirmation, [mbYes, mbNo], 0, mbYes) = mrYes then
  begin
    i := lvLegende.ItemIndex;
    if i <> -1 then
    begin
      id := StrToInt(lvLegende.Items[i].Caption);

      FDQuery := TFDQuery.Create(nil);
      try
        with FDQuery do
        begin
          Connection := fMain.FDConnection1;

          SQL.Text := 'DELETE FROM legende WHERE id = :ID;';
          Params.ParamByName('ID').AsInteger := id;
          try
            ExecSQL;
          except
            on E: Exception do
            begin
              ShowMessage('Fehler beim l�schen des Eintrages aus der Tabelle legende: ' + E.Message);
            end;
          end;
        end;
      finally
        FDQuery.Free;
        lvLegende.DeleteSelected;
      end;
    end;
  end;

  if(lvLegende.Items.Count <= 0) then
  begin
    btnLoadLegendeFromWeb.Visible := true;
  end
  else
  begin
    btnLoadLegendeFromWeb.Visible := false;
  end;

end;






procedure TfLegende.lvLegendeSelectItem(Sender: TObject; Item: TListItem; Selected: Boolean);
var
  i: integer;
begin
  i := lvLegende.ItemIndex;

  if i <> -1 then
  begin
    SelLegendeEntry     := StrToInt(lvLegende.Items[i].Caption);
    lbSchicht.Text      := lvLegende.Items[i].SubItems[0];
    lbEinsatz.Text      := lvLegende.Items[i].SubItems[1];
    lbLeistungsart.Text := lvLegende.Items[i].SubItems[2];
    lbVertragsNr.Text   := lvLegende.Items[i].SubItems[3];
    lbBeschreibung.Text := lvLegende.Items[i].SubItems[4];
    edUhrzeitVon.Text   := lvLegende.Items[i].SubItems[5];
    edUhrzeitBis.Text   := lvLegende.Items[i].SubItems[6];

    btnAddLegende.Caption      := 'Speichern';
  end
  else
  begin
    btnNewLegendeClick(nil);
  end;
end;






procedure TfLegende.LoadLegendeInListView(lv: TListView);
var
  l: TListItem;
  FDQuery: TFDQuery;
begin
  lv.Items.Clear;

  FDQuery := TFDQuery.Create(nil);
  try
    with FDQuery do
    begin
      Connection := fMain.FDConnection1;

      SQL.Text := 'SELECT id, schicht, einsatz, leistungsart, vertragsnr, beschreibung, UhrzeitVon, UhrzeitBis ' +
                  'FROM legende ORDER BY schicht;';
      Open;

      while not Eof do
      begin
        l := lv.Items.Add;
        l.Caption := FieldByName('id').AsString;
        l.SubItems.Add(FieldByName('schicht').AsString);
        l.SubItems.Add(FieldByName('einsatz').AsString);
        l.SubItems.Add(FieldByName('leistungsart').AsString);
        l.SubItems.Add(FieldByName('vertragsnr').AsString);
        l.SubItems.Add(FieldByName('beschreibung').AsString);
        l.SubItems.Add(FieldByName('UhrzeitVon').AsString);
        l.SubItems.Add(FieldByName('UhrzeitBis').AsString);
        next;
      end;
    end;
  finally
    FDQuery.Free;
    btnNewLegendeClick(nil);
  end;
end;





{
procedure TfLegende.LoadDataToListView(const ObjektNr: string);
var
  HTTP: TIdHTTP;
  SSLHandler: TIdSSLIOHandlerSocketOpenSSL;
  Response: string;
  JSONArray: TJSONArray;
  JSONObj: TJSONObject;
  ListItem: TListItem;
  i: Integer;
begin
  if(lvLegende.Items.Count <= 0) then
  begin
    HTTP := TIdHTTP.Create(nil);

    SSLHandler := TIdSSLIOHandlerSocketOpenSSL.Create(nil);
    SSLHandler.SSLOptions.SSLVersions := [sslvTLSv1_2];  // TLS 1.2 und 1.3 aktivieren
    HTTP.IOHandler := SSLHandler;

   // SSLHandler := TIdSSLIOHandlerSocketOpenSSL.Create(nil);
    try
      HTTP.IOHandler := SSLHandler;  // SSL f�r HTTPS
      HTTP.HandleRedirects := True;  // Umleiten erlauben, falls n�tig

      // PHP-Seite aufrufen mit ID als Parameter
      Response := HTTP.Get('https://esd.developercorner.de/Scripts/getDeploymentNumbersFromWeb.php?objektnr=' + OBJEKTNR);

      // JSON parsen
      JSONArray := TJSONObject.ParseJSONValue(Response) as TJSONArray;

      // ListView leeren
      lvLegende.Items.Clear;

      // Daten in die ListView einf�gen
      for i := 0 to JSONArray.Count - 1 do
      begin
        JSONObj := JSONArray.Items[i] as TJSONObject;
        ListItem := lvLegende.Items.Add;

        ListItem.Caption := '';

        while ListItem.SubItems.Count < 7 do
          ListItem.SubItems.Add('');  // Leere Eintr�ge hinzuf�gen


        ListItem.SubItems[1] := JSONObj.GetValue('Einsatznummer').Value;
        ListItem.SubItems[2] := JSONObj.GetValue('Leistungsart').Value;
        ListItem.SubItems[3] := JSONObj.GetValue('vertragsnr').Value;
        ListItem.SubItems[4] := JSONObj.GetValue('Bezeichnung').Value;
        ListItem.SubItems[5] := JSONObj.GetValue('VonUhrzeit').Value;
        ListItem.SubItems[6] := JSONObj.GetValue('BisUhrzeit').Value;
      end;

    finally
      HTTP.Free;
      SSLHandler.Free;
    end;
  end
  else
  begin
    showmessage('Die Funktion des automatischen Datenabgleichs kann nur genutzt werden wenn die Liste noch leer ist.');
  end;
end;
}




procedure TfLegende.InsertDataIntoSQLite(const ObjektNr: string);
var
  HTTP: TIdHTTP;
  SSLHandler: TIdSSLIOHandlerSocketOpenSSL;
  Response: string;
  JSONArray: TJSONArray;
  JSONObj: TJSONObject;
  i: Integer;
  FDQuery: TFDQuery;
begin
  if(lvLegende.Items.Count <= 0) then
  begin
    HTTP := TIdHTTP.Create(nil);
    SSLHandler := TIdSSLIOHandlerSocketOpenSSL.Create(nil);
    SSLHandler.SSLOptions.SSLVersions := [sslvTLSv1_2];
    HTTP.IOHandler := SSLHandler;

    try
      HTTP.HandleRedirects := True;  // Umleiten erlauben, falls n�tig

      try
        // PHP-Seite aufrufen mit objektnr als Parameter
        Response := HTTP.Get('https://esd.developercorner.de/Scripts/getDeploymentNumbersFromWeb.php?objektnr=' + ObjektNr);

        // JSON parsen
        JSONArray := TJSONObject.ParseJSONValue(Response) as TJSONArray;

        if JSONArray = nil then
        begin
          raise Exception.Create('Die Antwort enth�lt keine g�ltigen Daten.');
        end;

        // SQLite-Datenbankverbindung herstellen
        FDQuery := TFDQuery.Create(nil);
        try
          FDQuery.Connection := fMain.FDConnection1;

          // Daten in die SQLite-Datenbank einf�gen
          for i := 0 to JSONArray.Count - 1 do
          begin
            JSONObj := JSONArray.Items[i] as TJSONObject;

            with FDQuery do
            begin
              SQL.Text := 'INSERT INTO Legende (schicht, einsatz, leistungsart, vertragsnr, beschreibung, UhrzeitVon, UhrzeitBis) VALUES (:schicht, :einsatz, :leistungsart, :vertragsnr, :beschreibung, :UhrzeitVon, :UhrzeitBis)';

              ParamByName('schicht').AsString := JSONObj.GetValue('Kuerzel').Value;
              ParamByName('einsatz').AsString := JSONObj.GetValue('Einsatznummer').Value;
              ParamByName('leistungsart').AsString := JSONObj.GetValue('Leistungsart').Value;
              ParamByName('vertragsnr').AsString := JSONObj.GetValue('vertragsnr').Value;
              ParamByName('beschreibung').AsString := JSONObj.GetValue('Bezeichnung').Value;
              ParamByName('UhrzeitVon').AsString := JSONObj.GetValue('VonUhrzeit').Value;
              ParamByName('UhrzeitBis').AsString := JSONObj.GetValue('BisUhrzeit').Value;

              ExecSQL;
            end;
          end;
        finally
          FDQuery.Free;
        end;

      except
        on E: EIdHTTPProtocolException do
        begin
          if E.ErrorCode = 500 then
            ShowMessage('Im PHP-Script "getDeploymentNumbersFromWeb.php" ist ein Fehler aufgetreten.')
          else if E.ErrorCode = 404 then
            ShowMessage('Die Datei "getDeploymentNumbersFromWeb.php" konnte auf dem Server nicht gefunden werden.')
          else
            ShowMessage('HTTP-Fehler: ' + IntToStr(E.ErrorCode) + ' - ' + E.Message);
        end;
        on E: Exception do
        begin
          // Alle anderen Fehler behandeln
          ShowMessage('Fehler beim Abrufen der Daten: ' + E.Message);
        end;
      end;

    finally
      HTTP.Free;
      SSLHandler.Free;
    end;
  end
  else
  begin
    ShowMessage('Die Funktion des automatischen Datenabgleichs kann nur genutzt werden, wenn die Liste noch leer ist.');
  end;
end;










procedure TfLegende.LoadDataFromSQLite(const ObjektNr: string);
var
  FDQuery: TFDQuery;
  ListItem: TListItem;
begin
  FDQuery := TFDQuery.Create(nil);
  try
    FDQuery.Connection := fMain.FDConnection1;

    // Abfrage, um die Daten aus der Tabelle Legende zu laden
    with FDQuery do
    begin
      SQL.Text := 'SELECT * FROM Legende ORDER BY vertragsNr ASC, UhrzeitVon ASC';
      Open;
    end;

    // ListView leeren
    lvLegende.Items.Clear;

    // Daten in die ListView einf�gen
    while not FDQuery.Eof do
    begin
      ListItem := lvLegende.Items.Add;
      ListItem.Caption := FDQuery.FieldByName('id').AsString;

      // F�ge leere SubItems hinzu, um sicherzustellen, dass die Liste korrekt gef�llt ist
      while ListItem.SubItems.Count < 7 do
        ListItem.SubItems.Add('');

      // F�ge die Werte zu den SubItems hinzu
      ListItem.SubItems[0] := FDQuery.FieldByName('schicht').AsString;
      ListItem.SubItems[1] := FDQuery.FieldByName('einsatz').AsString;
      ListItem.SubItems[2] := FDQuery.FieldByName('leistungsart').AsString;
      ListItem.SubItems[3] := FDQuery.FieldByName('vertragsnr').AsString;
      ListItem.SubItems[4] := FDQuery.FieldByName('beschreibung').AsString;
      ListItem.SubItems[5] := FDQuery.FieldByName('UhrzeitVon').AsString;
      ListItem.SubItems[6] := FDQuery.FieldByName('UhrzeitBis').AsString;

      FDQuery.Next;  // N�chster Datensatz
    end;
  finally
    FDQuery.Free;
  end;
end;




end.
