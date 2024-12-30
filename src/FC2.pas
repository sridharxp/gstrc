(*
Copyright (C) 2022, Sridharan S

This file is part of GSTR Utilities

GSTR Utilities is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

Exce; to Tally is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License version 3
along with GSTR Utilities. If not, see <http://www.gnu.org/licenses/>.
*)
unit FC2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, StdCtrls,
  ShellApi;


type
  TfrmFC2 = class(TForm)
    lbl1: TLabel;
    edtHost: TEdit;
    lbl2: TLabel;
    edtPort: TEdit;
    mm1: TMainMenu;
    mniool1: TMenuItem;
    mniHelp1: TMenuItem;
    mniTestTally: TMenuItem;
    mniFindIP1: TMenuItem;
    mniFC2: TMenuItem;
    mniImportGSTRjson1: TMenuItem;
    mniReconcile2A2B1: TMenuItem;
    mniImportTallyData1: TMenuItem;
    procedure mniFC2Click(Sender: TObject);
    procedure mniFindIP1Click(Sender: TObject);
    procedure mniTestTallyClick(Sender: TObject);
    procedure mniImportGSTRjson1Click(Sender: TObject);
    procedure mniReconcile2A2B1Click(Sender: TObject);
    procedure mniImportTallyData1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmFC2: TfrmFC2;

implementation

uses FindIP, OffLineP3, OffTlyP, TlyInput, GSTRJSON, Jsn2Xls;

{$R *.dfm}

procedure TfrmFC2.mniFC2Click(Sender: TObject);
begin
    MessageDlg('Sridharan S' + #10 + '+91 98656 82910' + #10 +'excel2tallyerp@gmail.com'
  + #10 + 'Telegram: https://t.me/excel2tallyerp',
    mtInformation, [mbOK], 0);
end;

procedure TfrmFC2.mniFindIP1Click(Sender: TObject);
begin
  frmFindIP.Show;
end;

procedure TfrmFC2.mniTestTallyClick(Sender: TObject);
  var
    url: string;
begin
  url := 'http://' + edtHost.Text + ':' + edtPort.Text;
  URL := StringReplace(URL, '"', '%22', [rfReplaceAll]);
  ShellExecute(0, 'open', PChar(URL), nil, nil, SW_SHOWNORMAL);
end;
procedure TfrmFC2.mniImportGSTRjson1Click(Sender: TObject);
begin
  frmImpGSTR.Show;
end;
procedure TfrmFC2.mniReconcile2A2B1Click(Sender: TObject);
begin
  frMOffLineP.Show;
end;

procedure TfrmFC2.mniImportTallyData1Click(Sender: TObject);
begin
  frmOfTlyP.Show;
end;

end.
