unit uMitarbeiter;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls,
  Vcl.ComCtrls, AdvListV, FireDAC.Stan.ExprFuncs, FireDAC.Phys.SQLiteDef, FireDAC.Stan.Intf,
  FireDAC.Stan.Option, FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf,
  FireDAC.Stan.Def, FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys,
  FireDAC.VCLUI.Wait, FireDAC.Stan.Param, FireDAC.DatS, FireDAC.DApt.Intf,
  FireDAC.DApt, FireDAC.Comp.DataSet, FireDAC.Comp.Client, FireDAC.Phys.SQLite,
  Data.DB, DateUtils, System.UITypes, Vcl.Mask, Vcl.Menus;



type
  TfMitarbeiter = class(TForm)
    lvMitarbeiter: TAdvListView;
    edNachname: TLabeledEdit;
    edVorname: TLabeledEdit;
    edPersonalNr: TLabeledEdit;
    btnNewMitarbeiter: TButton;
    btnSave: TButton;
    Panel1: TPanel;
    lbHinweis: TLabel;
    MainMenu1: TMainMenu;
    Mitarbeiter1: TMenuItem;
    Exportieren1: TMenuItem;
    Exportieren2: TMenuItem;
    procedure btnNewMitarbeiterClick(Sender: TObject);
    procedure lvMitarbeiterSelectItem(Sender: TObject; Item: TListItem; Selected: Boolean);
    procedure FormShow(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure lvMitarbeiterRightClickCell(Sender: TObject; iItem,
      iSubItem: Integer);
    procedure Exportieren1Click(Sender: TObject);
    procedure Exportieren2Click(Sender: TObject);
  private
    procedure LoadMitarbeiterInListView(lv: TListView);
    procedure UpdateMitarbeiterInDB(id: integer);
    procedure InsertNewMitarbeiterInDB;
  public
    { Public-Deklarationen }
  end;

var
  fMitarbeiter: TfMitarbeiter;
  SelEntry: integer;


implementation

{$R *.dfm}

uses uMain, uDBFunktionen;





procedure TfMitarbeiter.btnNewMitarbeiterClick(Sender: TObject);
begin
  edNachname.Clear;
  edVorname.Clear;
  edPersonalNr.Clear;
  edNachname.SetFocus;
  lvMitarbeiter.ItemIndex := -1;
  btnSave.Caption := 'Hinzufügen';
  SelEntry := -1;
end;





procedure TfMitarbeiter.btnSaveClick(Sender: TObject);
var
  nachname, vorname, personalnr: string;
begin
  nachname := trim(edNachname.Text);
  vorname := trim(edVorname.Text);
  personalnr := trim(edPersonalNr.Text);


  if(nachname <> '') then
  begin
    if(SelEntry = -1) then
    begin
      InsertNewMitarbeiterInDB;
    end
    else if(SelEntry <> -1) then
    begin
      UpdateMitarbeiterInDB(SelEntry);
    end;

    btnNewMitarbeiterClick(nil);
  end
  else
  begin
    showmessage('Bitte füllen Sie alle Eingabefelder aus!');
  end;
end;






procedure TfMitarbeiter.Exportieren1Click(Sender: TObject);
begin
  BackupSQLiteTable('Mitarbeiter', PATH + 'DBDUMP');
  showmessage('Mitarbeiterliste unter dem Namen "Mitarbeiter.sql" im Verzeichnis "DBDUMP" gespeichert');
end;

procedure TfMitarbeiter.Exportieren2Click(Sender: TObject);
begin
  if FileExists(PATH + 'DBDUMP\Mitarbeiter.sql') then
  begin
    ImportSQLiteTable(PATH + 'DBDUMP\Mitarbeiter.sql');
    LoadMitarbeiterInListView(lvMitarbeiter);
  end;
end;

procedure TfMitarbeiter.UpdateMitarbeiterInDB(id: integer);
var
  FDQuery: TFDQuery;
  i: integer;
begin
  i := lvMitarbeiter.ItemIndex;
  if(i <> -1) then
  begin
    FDQuery := TFDquery.Create(nil);
    try
      with FDQuery do
      begin
        Connection := fMain.FDConnection1;

        SQL.Text := 'UPDATE mitarbeiter SET nachname = :NACHNAME, vorname = :VORNAME, personalnr = :PERSONALNR ' +
                    'WHERE id = :ID;';
        Params.ParamByName('ID').AsInteger        := SelEntry;
        Params.ParamByName('NACHNAME').AsString   := edNachname.Text;
        Params.ParamByName('VORNAME').AsString    := edVorname.Text;
        Params.ParamByName('PERSONALNR').AsString := edPersonalNr.Text;
        try
          ExecSQL;
        except
          on E: Exception do
          begin
            ShowMessage('Fehler beim ändern des Mitarbeiters in der Tabelle mitarbeiter: ' + E.Message);
          end;
        end;
      end;
    finally
      FDQuery.Free;

      lvMitarbeiter.Items[i].SubItems[0] := edNachname.Text;
      lvMitarbeiter.Items[i].SubItems[1] := edVorname.Text;
      lvMitarbeiter.Items[i].SubItems[2] := edPersonalNr.Text;

      edNachname.Clear;
      edVorname.Clear;
      edPersonalNr.Clear;
    end;
  end;
end;




procedure TfMitarbeiter.InsertNewMitarbeiterInDB;
var
  FDQuery: TFDQuery;
  l: TListItem;
  insertedID: Integer;
begin
  insertedID := -1;

  FDQuery := TFDQuery.Create(nil);
  try
    with FDQuery do
    begin
      Connection := fMain.FDConnection1;

      SQL.Text := 'INSERT INTO mitarbeiter (nachname, vorname, personalnr) ' +
                  'VALUES (:NACHNAME, :VORNAME, :PERSONALNR);';
      Params.ParamByName('NACHNAME').AsString   := edNachname.Text;
      Params.ParamByName('VORNAME').AsString    := edVorname.Text;
      Params.ParamByName('PERSONALNR').AsString := edPersonalNr.Text;

      try
        ExecSQL;

        SQL.Text := 'SELECT last_insert_rowid() as last_id;';
        Open;
        insertedID := FieldByName('last_id').AsInteger;
      except
        on E: Exception do
        begin
          ShowMessage('Fehler beim Einfügen in die Tabelle mitarbeiter: ' + E.Message);
        end;
      end;
    end;
  finally
    FDQuery.Free;

    l := lvMitarbeiter.Items.Add;
    l.Caption := IntToStr(insertedID);
    l.SubItems.Add(edNachname.Text);
    l.SubItems.Add(edVorname.Text);
    l.SubItems.Add(edPersonalNr.Text);
  end;
end;





procedure TfMitarbeiter.FormShow(Sender: TObject);
begin
  SelEntry := -1;
  btnNewMitarbeiterClick(nil);

  LoadMitarbeiterInListView(lvMitarbeiter);

  lvMitarbeiter.SortColumn := 1;
  lvMitarbeiter.AlphaSort;
  lvMitarbeiter.Sort;
end;






procedure TfMitarbeiter.lvMitarbeiterRightClickCell(Sender: TObject; iItem, iSubItem: Integer);
var
  i, id: integer;
  FDQuery: TFDQuery;
begin
  if MessageDlg('Wollen Sie diesen Eintrag wirklich löschen?', mtConfirmation, [mbYes, mbNo], 0, mbYes) = mrYes then
  begin
    i := lvMitarbeiter.ItemIndex;
    if i <> -1 then
    begin
      id := StrToInt(lvMitarbeiter.Items[i].Caption);

      FDQuery := TFDQuery.Create(nil);
      try
        with FDQuery do
        begin
          Connection := fMain.FDConnection1;

          SQL.Text := 'DELETE FROM mitarbeiter WHERE id = :ID;';
          Params.ParamByName('ID').AsInteger := id;
          try
            ExecSQL;
          except
            on E: Exception do
            begin
              ShowMessage('Fehler beim löschen des Eintrages aus der Tabelle mitarbeiter: ' + E.Message);
            end;
          end;
        end;
      finally
        FDQuery.Free;
        lvMitarbeiter.DeleteSelected;
      end;
    end;
  end;
end;







procedure TfMitarbeiter.lvMitarbeiterSelectItem(Sender: TObject; Item: TListItem; Selected: Boolean);
var
  i: integer;
begin
  i := lvMitarbeiter.ItemIndex;

  if i <> -1 then
  begin
    SelEntry          := StrToInt(lvMitarbeiter.Items[i].Caption);
    edNachname.Text   := lvMitarbeiter.Items[i].SubItems[0];
    edVorname.Text    := lvMitarbeiter.Items[i].SubItems[1];
    edPersonalNr.Text := lvMitarbeiter.Items[i].SubItems[2];

    btnSave.Caption   := 'Speichern';
  end
  else
  begin
    btnNewMitarbeiterClick(nil);
  end;
end;






procedure TfMitarbeiter.LoadMitarbeiterInListView(lv: TListView);
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

      SQL.Text := 'SELECT id, nachname, vorname, personalnr FROM mitarbeiter ORDER BY nachname;';
      Open;

      while not Eof do
      begin
        l := lv.Items.Add;
        l.Caption := FieldByName('id').AsString;
        l.SubItems.Add(FieldByName('nachname').AsString);
        l.SubItems.Add(FieldByName('vorname').AsString);
        l.SubItems.Add(FieldByName('personalnr').AsString);
        next;
      end;
    end;
  finally
    FDQuery.Free;
    btnNewMitarbeiterClick(nil);
  end;
end;




end.
