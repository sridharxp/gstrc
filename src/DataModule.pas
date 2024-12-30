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
unit DataModule;

interface

uses
  SysUtils, Classes,
  xlstbl313,
  Dialogs;

type
  TDM = class(TDataModule)
    procedure DataModuleCreate(Sender: TObject);
    procedure CreateTally2AB;
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  DM: TDM;
  Flds: TStringList;

implementation

{$R *.dfm}

procedure TDM.DataModuleCreate(Sender: TObject);
begin
ShortDateFormat := 'DD-MM-YYYY';
DateSeparator := '-';

if not DirectoryExists('.\Data\') then
  if not CreateDir('.\Data\') then
    raise Exception.Create('Cannot create folder Data');
{
if not DirectoryExists('.\Report\') then
  if not CreateDir('.\Report\') then
    raise Exception.Create('Cannot create folder Report');
}
  CreateTally2AB;
end;

procedure TDM.CreateTally2AB;
var
  XL: TbjXLSTable;
  rXLName: string;
begin
  rXLName := 'Tally2AB.xls';

  Flds := TStringList.Create;
//  Flds.Clear;
  Flds.Add('PERIOD');
  Flds.Add('GSTIN');
  Flds.Add('Invoice_No');
  Flds.Add('Invoice_TYPE');
  Flds.Add('Invoice_Date');
  Flds.Add('Invoice_Value');
  Flds.Add('PLACE_OF_SUPPLY');
  Flds.Add('Supplier');
  Flds.Add('Taxable_Value');
  Flds.Add('SGST');
  Flds.Add('CGST');
  Flds.Add('IGST');
  Flds.Add('CESS');
  Flds.Add('Total_Tax');
  Flds.Add('Tax_Rate');

  if FileExists('.\Data\' + rXLName) then
    Exit;

  XL := TbjXLSTable.Create;
  XL.XLSFileName := '.\Data\' + rXLName;

  XL.SetSheet('PORTAL');
  XL.SetFields(flds, True);
  XL.Save;

  XL.SetSheet('BOOKS');
  XL.SetFields(flds, True);
  XL.Save;

  XL.SetSheet('JSON');
  XL.SetFields(flds, True);
  XL.Save;
{$IFDEF SM}
  XL.SetFieldFormat('PERIOD', 14);
  XL.SetFieldFormat('Invoice_Date', 14);
{$ENDIF SM}
{$IFDEF NXL}
  XL.SetFieldFormat('PERIOD', '@');
  XL.SetFieldFormat('Invoice_Date', 'DD-MMM-YY');
  XL.SetFieldFormat('INVOICE_VALUE', '###,##0.00');
  XL.SetFieldFormat('Taxable_Value', '###,##0.00');
  XL.SetFieldFormat('SGST', '###,##0.00');
  XL.SetFieldFormat('CGST', '###,##0.00');
  XL.SetFieldFormat('IGST', '###,##0.00');
{$ENDIF NXL}

  XL.Close;
  XL.Free;
end;

procedure TDM.DataModuleDestroy(Sender: TObject);
begin
  Flds.Clear;
  Flds.Free;
end;

end.
