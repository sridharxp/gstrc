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
unit GSTRJSON;
{ TbjClient is not used }

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, OleCtrls, StdCtrls,
  ExtCtrls, Menus,
  ShellAPI,
  bjxml33,
  Jsn2Xls,
  Easysize;

type
  TfrmImpGSTR = class(TForm)
    OpenDialog1: TOpenDialog;
    frmEzReSizer: TFormResizer;
    edtReqFileName: TEdit;
    lbl1: TLabel;
    lbl2: TLabel;
    cmbdbName: TComboBox;
    Info: TLabel;
    btn1: TButton;
    edt1: TEdit;
    edtdbName: TEdit;
    Info1: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure mniSourceCode1Click(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
  protected
    StartTime, EndTime, Elapsed: double;
    Hrs, Mins, Secs, Msecs: word;
  public
    { Public declarations }
    Host: string;
  end;

procedure UpdateMsg(const msg: string);

var
  frmImpGSTR: TfrmImpGSTR;

implementation
{$R *.dfm}
procedure TfrmImpGSTR.FormCreate(Sender: TObject);
begin
  Host := 'http://127.0.0.1:9000';
  frmEzReSizer.InitializeForm;
end;

procedure TfrmImpGSTR.FormResize(Sender: TObject);
begin
  frmEzReSizer.ResizeAll;
end;

procedure TfrmImpGSTR.mniSourceCode1Click(Sender: TObject);
begin
  MessageDlg('git@github.com:sridharxp/connect2tally.git',
    mtInformation, [mbOK], 0);
end;

procedure TfrmImpGSTR.btn1Click(Sender: TObject);
var
  Obj: TbjJsn2Xls;
  Timestr: string;
begin
  if (Length(edtReqFileName.Text) = 0) or (not FileExists(EdtReqFileName.Text)) then
    if not OpenDialog1.Execute then
      Abort
  else
    EdtReqFileName.Text := OpenDialog1.FileName;
  if not FileExists(EdtReqFileName.Text) then
    Exit;
  StartTime := Time;
  Obj := TbjJsn2Xls.Create;
  Obj.AppPath := ExtractFilePath(Application.ExeName);
  Obj.JFileName := EdtReqFileName.Text;
  Obj.FUpdate := UpdateMsg;
  Obj.Process;
  EndTime := Time;
  Elapsed := EndTime - StartTime;
  DecodeTime(Elapsed, Hrs, Mins, Secs, MSecs);
  if Mins > 0 then
  timestr := InttoStr(Mins) + ' . ' + InttoStr(Secs) + ' Minutes'
  else if Secs > 0 then
  timestr := InttoStr(Secs) + ' . ' + InttoStr(MSecs) + ' Seconds'
  else
  timestr := InttoStr(MSecs) + ' MSecs';
  Info.Caption := timestr;
  MessageDlg(IntToStr(Obj.nInv) + ' Invoices Imported in JSON WS', mtInformation, [mbOK], 0);
  Obj.Free;
end;
procedure TfrmImpGSTR.FormDestroy(Sender: TObject);
begin
//    req.Free;
end;


procedure UpdateMsg(const msg: string);
begin
  if Length(msg) > 0 then
    frmImpGSTR.Info.Caption := Msg;
  Application.ProcessMessages;
end;

end.

