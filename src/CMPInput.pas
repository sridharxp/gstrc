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
unit CMPInput;

interface

uses
  Classes, SysUtils, DateUtils,
  Dialogs,
  DateFns,
  Variants,
  xlstbl313;

type
  Tfnupdate = procedure(msg: string);

  PV = Record
    PERIOD: string;
    GSTN: string;
    Invoice_No: string;
    Invoice_Date: string;
    Invoice_Value: string;
    PLACE_OF_SUPPLY: string;
    Supplier: string;
    Taxable_Value: string;
    SGST: string;
    CGST: string;
    IGST: string;
    CESS: string;
    Tax: string;
    Rate: string;
  end;
  pPV = ^PV;

type
  TbjVerifyOLine = class
  private
    { Private declarations }
  protected
  public
    { Public declarations }
    nNInp: integer;
    FrmDt, ToDt: string;
    SupplierGsTN: string;
    IsSupplierGSTN: Boolean;
    XL: TbjXLSTable;
    FUpdate: TfnUpdate;
    OnList: TList;
    OffList: TList;
    VList: TList;
    procedure VerifyOline;
    procedure LoadList;
    procedure LoadRec(Itm: pPV);
    function UnAvailable(item: pPV): Boolean ;
    function Available(item: pPV): Boolean ;
    function Pending(item: pPV): Boolean ;
    procedure Writexls(item: pPV);
    constructor Create;
    destructor Destroy; override;
  end;

  function VarIsSame(_A, _B: Variant): Boolean;
  function VarCompare(_A, _B: Variant): Boolean;

function DateTimeStrToSTOD(const aDate : String) : string;

implementation

uses DataModule;

Constructor TbjVerifyOLine.Create;
//var
//  Obj: TxlsBlank;
begin
  Inherited;
  OnList := TList.Create;
  OFFList := TList.Create;
  XL := TbjXLSTable.Create;
  XL.XLSFileName := '.\Data\Tally2AB.xls';
  xl.ToSaveFile := True;
  nNInp := 0;
end;

destructor TbjVerifyOLine.Destroy;
var
  item: pPV;
  i: integer;
begin
  for i := 0 to OnList.Count - 1 do
  begin
    item := OnList.Items[i];
    Dispose(item);
  end;
  OnList.Clear;
  OnList.Free;
  for i := 0 to OffList.Count - 1 do
  begin
    item := OffList.Items[i];
    Dispose(item);
  end;
  OffList.Clear;
  OffList.Free;
  XL.Close;
  XL.Free;
  inherited;
end;

procedure TbjVerifyOLine.VerifyOline;
var
  item_L: pPV;
  i: integer;
  rMatched: Boolean;
begin
  LoadList;

  XL.DropSheet('NOINPUT');
  XL.SetSheet('NOINPUT');
  XL.SetFields(flds, True);
  XL.GetFields(flds);
  XL.First;
  for i := 0 to OffList.Count-1 do
  begin
    rMatched := False;
    item_L := OffList.Items[i];
    FUpdate('Checking Not in Portal' + item_L.Invoice_No);
    rMatched := UnAvailable(item_L);
    if not rMatched then
    begin
      WriteXls(item_L);
    end;
  end;
//  WriteBlanks;
//  WriteBlanks;
  XL.Save;

  XL.DropSheet('INPUT');
  XL.SetSheet('INPUT');
  XL.SetFields(flds, True);
//  XL.SetSheet('INPUT');
  XL.GetFields(flds);
  XL.First;
  for i := 0 to OffList.Count-1 do
  begin
    rMatched := True;
    item_L := OffList.Items[i];
    FUpdate('Checking Available' + item_L.Invoice_No);
    rMatched := Available(item_L);
    if rMatched then
    begin
      WriteXls(item_L);
//      nNInp := nNInp + 1;
    end;
 end;
  XL.Save;

  XL.DropSheet('NOBOOKS');
  XL.SetSheet('NOBOOKS');
  XL.SetFields(flds, True);
//  XL.SetSheet('NOBOOKS');
  XL.GetFields(flds);
  XL.First;
  for i := 0 to OnList.Count-1 do
  begin
    rMatched := False;
    item_L := OnList.Items[i];
    FUpdate('Checking Not in Books' + item_L.Invoice_No);
    rMatched := Pending(item_L);
    if not rMatched then
    begin
      WriteXls(item_L);
    end;
  end;
//  WriteBlanks;
//  WriteBlanks;
  XL.Save;
end;


function TbjVerifyOLine.UnAvailable(item: pPV): Boolean;
var
  item_R: pPV;
  i: integer;
begin
  Result := False;
  for i := 0 to OnList.Count-1 do
  begin
    item_R := OnList.Items[i];
    if item.GSTN <> item_R.GSTN then
      Continue;
    if item.Invoice_Date <> item_R.Invoice_Date then
      Continue;
    if item.Invoice_No <> item_R.Invoice_No then
      Continue;
    Result := True;
    Break;
  end;
end;

function TbjVerifyOLine.Available(item: pPV): Boolean;
var
  item_R: pPV;
  i: integer;
begin
  Result := False;
  for i := 0 to OnList.Count-1 do
  begin
    item_R := OnList.Items[i];
    if (item.GSTN = item_R.GSTN) and
    (item.Invoice_Date = item_R.Invoice_Date) and
    (item.Invoice_No = item_R.Invoice_No) then
    begin
      Result := True;
      if Length(item.PLACE_OF_SUPPLY) > 0 then
      if item.PLACE_OF_SUPPLY <> item_R.PLACE_OF_SUPPLY then
      begin
        item.PLACE_OF_SUPPLY := '*';
      end;
      if Length(item.Invoice_Value) > 0 then
      if item.Invoice_Value <> item_R.Invoice_Value then
      begin
        item.Invoice_Value := '*';
      end;
      if Length(item.Taxable_Value) > 0 then
      if item.Taxable_Value <> item_R.Taxable_Value then
      begin
        item.Taxable_Value := '*';
      end;
      if Length(item.SGST) > 0 then
      if item.SGST <> item_R.SGST then
      begin
        item.SGST := '*'
      end;
      if Length(item.CGST) > 0 then
      if item.CGST <> item_R.CGST then
        begin
        item.CGST := '*';
      end;
      if Length(item.IGST) > 0 then
      if item.IGST <> item_R.IGST then
      begin
        item.IGST := '*';
      end;
      Break;
    end;
  end;
end;

function TbjVerifyOLine.Pending(item: pPV): Boolean;
var
  item_R: pPV;
  i: integer;
begin
  Result := False;
  for i := 0 to OffList.Count-1 do
  begin
    item_R := OffList.Items[i];
    if item.GSTN <> item_R.GSTN then
      Continue;
    if item.Invoice_Date <> item_R.Invoice_Date then
      Continue;
    if item.Invoice_No <> item_R.Invoice_No then
      Continue;
    Result := True;
    Break;
  end;
end;

procedure TbjVerifyOLine.Writexls(item: pPV);
var
  rDt: string;
begin
  if (item.Invoice_Date < FrmDt) or (item.Invoice_Date > ToDt) then
    Exit;
  if IsSupplierGSTN then
  if item.GSTN <> SupplierGsTN then
    Exit;
  if (Length(item.SGST) = 0) and (Length(item.CGST) = 0) and (Length(item.IGST) = 0) then
  Exit;
  XL.SetFieldVal('PERIOD', item.PERIOD);
  XL.SetFieldVal('GSTiN', item.GSTN);
  XL.SetFieldVal('INVOICE_NO', item.Invoice_No);
  rDt := Copy(item.Invoice_Date, 7, 2) + '-' + Copy(item.Invoice_Date, 5, 2) +
  '-'  + Copy(item.Invoice_Date, 3, 2);
//  XL.SetFieldVal('INVOICE_DATE', item.Invoice_Date);
  XL.SetFieldVal('INVOICE_DATE', rDt);
  XL.SetFieldVal('INVOICE_VALUE', item.Invoice_Value);
  XL.SetFieldVal('SUPPLIER', item.Supplier);
  XL.SetFieldVal('Taxable_Value', item.Taxable_Value);
  XL.SetFieldVal('SGST', item.SGST);
  XL.SetFieldVal('CGST', item.CGST);
  XL.SetFieldVal('IGST', item.IGST);
  XL.SetFieldVal('Tax_Rate', item.Rate);
  XL.CurrentRow := XL.CurrentRow + 1;
  nNInp := nNInp + 1;
end;


procedure TbjVerifyOLine.LoadList;
var
  item: pPV;
begin
//  onLine := TbjXLSTable.Create;
  XL.SetSheet('PORTAL');
  XL.GetFields(flds);
{$IFDEF SM}
  XL.SetFieldFormat('PERIOD', 14);
  XL.SetFieldFormat('Invoice_Date', 14);
{$ENDIF SM}
{$IFDEF NXL}
  XL.SetFieldFormat('PERIOD', '@');
  XL.SetFieldFormat('Invoice_Date', 'DD-MMM-YY');
{$ENDIF NXL}

  XL.First;
  if XL.EOF then
    MessageDlg('Portal worksheet is empty', mtWarning, [mbOK], 0);
  while not XL.EOF do
  begin
    item := New(pPV);
    LoadRec(item);
    OnList.Add(item);
    XL.Next;
  end;

  XL.SetSheet('BOOKS');
  XL.GetFields(flds);
{$IFDEF SM}
  XL.SetFieldFormat('PERIOD', 14);
  XL.SetFieldFormat('Invoice_Date', 14);
{$ENDIF SM}
{$IFDEF NXL}
  XL.SetFieldFormat('PERIOD', '@');
  XL.SetFieldFormat('Invoice_Date', 'DD-MMM-YY');
{$ENDIF NXL}
  XL.First;
  if XL.EOF then
    MessageDlg('Books worksheet is empty', mtWarning, [mbOK], 0);
  XL.First;
  while not XL.EOF do
  begin
    item := New(pPV);
    LoadRec(item);
    OffList.Add(item);
    XL.Next;
  end;
//  ShowMessage(IntToStr(OnList.Count));
end;

procedure TbjVerifyOLine.LoadRec(Itm: pPV);
begin
    Itm.PERIOD := XL.GetFieldString('PERIOD');
    itm.GSTN := XL.GetFieldString('GSTIN');
    itm.Invoice_No := XL.GetFieldString('INVOICE_NO');
    itm.Invoice_Date := XL.GetFieldSDate('INVOICE_DATE');
    itm.Invoice_Date := DateTimeStrToSTOD(itm.Invoice_Date);
    itm.Invoice_Value := XL.GetFieldString('INVOICE_Value');
    itm.Supplier := XL.GetFieldString('SUPPLIER');
    itm.Taxable_Value := XL.GetFieldString('Taxable_Value');
    itm.SGST := XL.GetFieldString('SGST');
    itm.CGST := XL.GetFieldString('CGST');
    itm.IGST := XL.GetFieldString('IGST');
    itm.Tax := XL.GetFieldString('TOTAL_TAX');
    itm.Rate := XL.GetFieldString('Tax_Rate');
end;

function VarIsSame(_A, _B: Variant): Boolean;
var
  LA, LB: TVarData;
begin
  LA := FindVarData(_A)^;
  LB := FindVarData(_B)^;
  if LA.VType <> LB.VType then
    Result := False
  else
    Result := (_A = _B);
end;

function VarCompare(_A, _B: Variant): Boolean;
begin
  Result := VarIsSame(_A, _B);
  if not Result then
    Result := VarSameValue(_A, _B);
end;

function DateTimeStrToSTOD(const aDate : String) : string;
var
  Dt: TDateTime;
  DateTimeFormat, DateTimeStr: string;
  Separators: set of AnsiChar;
begin
  DateTimeStr := aDate;
  Separators := ['/', '-', '.'];
  DateTimeFormat := 'Err';
  if (DateTimeStr[3] in Separators) and (DateTimeStr[6] in Separators) then
  begin
    if (Length(DateTimeStr) = 8) then
    DateTimeFormat := 'DD.MM.YY';
    if (Length(DateTimeStr) = 10) then
    DateTimeFormat := 'DD.MM.YYYY';
  end;
  if (DateTimeStr[3] in Separators) and (DateTimeStr[7] in Separators) then
  begin
    if (Length(DateTimeStr) = 9) then
    DateTimeFormat := 'DD.MMM.YY';
    if (Length(DateTimeStr) = 11) then
    DateTimeFormat := 'DD.MMM.YYYY';
  end;
  if DateTimeFormat = 'Err' then
  begin
    Result := DateTimeStr;
    Exit;
  end;
  Dt := DateTimeStrEval(DateTimeFormat, DateTimeStr);
  Result := FormatDateTime('YYYYMMDD', Dt)
end;

end.
