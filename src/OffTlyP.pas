(*
Copyright (C) 2013, Sridharan S

This file is part of Excel to Tally.

Excel to Tally is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

Exce; to Tally is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License version 3
along with Excel to Tally. If not, see <http://www.gnu.org/licenses/>.
*)

unit OffTlyP;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls,
  ShellApi,
  TlyInput,
  CMPInput,
  PClientFns,
  ObjCollFn,
  StrFn,

  MstListImp, Easysize;

type
  TfrmOfTlyP = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    btnImport: TButton;
    DateTimePicker2: TDateTimePicker;
    Info: TLabel;
    cmbLedGroup: TComboBox;
    Label4: TLabel;
    cmbVendor: TComboBox;
    btnRefresh: TButton;
    cmbFirm: TComboBox;
    Label6: TLabel;
    Label8: TLabel;
    DateTimePicker1: TDateTimePicker;
    frmEzReSizer: TFormResizer;
    cbFaxInv: TCheckBox;
    procedure btnImportClick(Sender: TObject);
    procedure btnRefreshClick(Sender: TObject);
    //procedure btnReconcileClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormResize(Sender: TObject);
  private
    { Private declarations }
  protected
        StartTime, EndTime, Elapsed: double;
    Hrs, Mins, Secs, Msecs: word;
  public
    { Public declarations }
    rStr: string;
  end;

procedure UpdateMsg(Msg: string);

var
  frmOfTlyP: TfrmOfTlyP;

implementation

uses
  FC2;

{$R *.dfm}

procedure TfrmOfTlyP.btnImportClick(Sender: TObject);
var
  xdb: TbjxLedVch4P;
  Timestr: string;
begin
  btnImport.Enabled := False;
  btnRefresh.Enabled := False;
  StartTime := Time;
//  btnReconcile.Enabled := False;
//  if (Length(cmbVendor.Text) = 0) and (Length(cmbLedger.Text) = 0) then
//    cmbLedger.Text := 'Purchase Accounts';
  xdb := TbjxLedVch4P.Create;
    try
    xdb.FUpdate := UpdateMsg;
    xdb.FrmDt := FormatDateTime('yyyyMMDD',DateTimePicker1.date);
    xdb.ToDt := FormatDateTime('yyyyMMDD',DateTimePicker2.date);

    if Length(Trim(xdb.FrmDt)) = 0 then
      xdb.FrmDt := FormatDateTime('yyyyMMDD',Now);
    if Length(Trim(xdb.ToDt)) = 0 then
      xdb.ToDt := FormatDateTime('yyyyMMDD',Now);
    if Length(cmbLedGroup.Text) > 0 then
      xdb.psType := cmbLedGroup.Text;

//  xdb.ToSaveXMLFile := False;
//    xdb.ToSaveXMLFile := True;
    xdb.Vendor := cmbVendor.Text;
    xdb.VchType := cmbLedGroup.Text;
    if cbFaxInv.Checked then
      xdb.IsTaxInv := True;
//    xdb.Ledger := cmbLedger.Text;
    xdb.Host := 'http://' + frmFC2.edtHost.Text + ':'+
      frmFC2.edtPort.Text;
//    ShowMessage(xdb.Host);
    xdb.Process;
  finally
    xdb.Free;
    btnImport.Enabled := True;
    btnRefresh.Enabled := True;
//    btnReconcile.Enabled := True;
  end;
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

    MessageDlg(IntToStr(xdb.ndups) + ' Invoice lines imported',
      mtInformation, [mbOK],0);

  ShellExecute(frmFC2.Handle, Pchar('Open'), PChar(
      '.\Data\Tally2AB.xls'),
      nil,
      nil, SW_SHOWNORMAL);
end;

procedure UpdateMsg;
begin
  if length(msg) > 0 then
    frMOfTlyP.Info.Caption := Msg
  else
    frMOfTlyP.Info.Caption := 'Done';
  Application.ProcessMessages;
end;

procedure TfrmOfTlyP.btnRefreshClick(Sender: TObject);
var
  i: integer;
  CMPList: TStringlist;
  LedList: TStringlist;
  PartyList: TStringlist;
  Obj: TbjMstListImp;
  str, rStr: String;
  rLCount: Integer;
begin
//  btnReconcile.Enabled := False;
  btnImport.Enabled := False;
  btnRefresh.Enabled := False;
  CMPList := TStringList.Create;
  LedList := TStringList.Create;
  PartyList := TStringList.Create;
  try
  Obj := TbjMstListImp.Create;
  Obj.ToPack := False;
//  ClientFns.Host := 'http://' + frmFC2.edtHost.Text + ':' +
//    frmFC2.edtPort.Text;
  Obj.Host := 'http://' + frmFC2.edtHost.Text + ':' +
    frmFC2.edtPort.Text;
  try
    try
      CMPList.Text := Obj.GetCMPText;
      CMPList.Sorted := True;
//      case cmbLedGroup.ItemIndex of
//      2 :

        LedList.Text := Obj.GetGrpLedText('Purchase Accounts');
        PartyList.Text := Obj.GetPartyText(False);
{
        rStr := FullListEx(#13, 'Ledger', '$Name');
        rStr := StringReplace(rStr, #13#10, '', [rfReplaceAll, rfIgnoreCase]);
        rStr := StringReplace(rStr, #3#32, #13#10, [rfReplaceAll, rfIgnoreCase]);
        LedList.Text := rStr;
}
//      else
//        LedList.Text := Obj.GetPartyText(True);
//      end;
      LedList.Sorted := True;
    except
      MessageDlg('Error Connecting to Tally', mtError, [mbOK], 0);
    end;
  finally
    Obj.Free;
  end;

  cmbFirm.Items.Add('');
  cmbFirm.Clear;
  for i:= 0 to CMPList.Count-1 do
    cmbFirm.Items.Add(CMPList.Strings[i]);
{
  cmbLedger.Clear;
  for i:= 0 to LedList.Count-1 do
    cmbLedger.Items.Add(LedList.Strings[i]);
}
  cmbVendor.Clear;
//  cmbVendor.Items.Add('');
  rLCount :=  PartyList.Count;
  for i:= 0 to PartyList.Count-1 do
    cmbVendor.Items.Add(PartyList.Strings[i]);
//  cmbParty.Clear;
//  for i:= 0 to LedList.Count-1 do
//    cmbParty.Items.Add(LedList.Strings[i]);
  finally
//  btnReconcile.Enabled := True;
  btnImport.Enabled := True;
  btnRefresh.Enabled := True;
  CMPList.Clear;
  CMPList.Free;
  PartyList.Clear;
  PartyList.Free;
  LedList.Clear;
  LedList.Free;
  end;
  MessageDlg(IntToStr(rLCount) + ' Ledgers imported', mtInformation, [mbOK], 0);
end;
{
procedure TfrmOfTlyP.btnReconcileClick(Sender: TObject);
var
  Obj: TbjVerifyOLine;
begin
  btnImport.Enabled := False;
  btnRefresh.Enabled := False;
//  btnReconcile.Enabled := False;

  Obj := TbjVerifyOLine.Create;
  try
    obj.FUpdate := UpdateMsg;
    Obj.VerifyOline;
  finally
    Obj.Free;
  end;

  MessageDlg(IntToStr(obj.nNInp) + ' lines found',
    mtInformation, [mbOK],0);

  btnImport.Enabled := True;
  btnRefresh.Enabled := True;
//  btnReconcile.Enabled := True;
   ShellExecute(frmFC2.Handle, Pchar('Open'), PChar(
      '.\Data\Tally2AB.xls'),
      nil,
      nil, SW_SHOWNORMAL);
end;
}
procedure TfrmOfTlyP.FormCreate(Sender: TObject);
begin
  frmEzReSizer.InitializeForm;
end;

procedure TfrmOfTlyP.FormResize(Sender: TObject);
begin
  frmEzReSizer.ResizeAll;
end;

end.
