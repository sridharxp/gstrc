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
unit Jsn2Xls;

interface

uses
  Classes, SysUtils,
  bjxml33,
  xlstbl313,
  Dialogs;

type
  Tfnupdate = procedure(const msg: string);

  PV = Record
    PERIOD: string;
    GSTN: string;
    Invoice_No: string;
    Invoice_Type: string;
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
  TbjJsn2Xls = class
  private
    { Private declarations }
  protected
    jStr: IbjXml;
    cnNode: IbjXml;

    docdataNode, GSTNNode: IbjXml;
    b2bNode, invNode: IbjXml;
    SupNode, SupGSTNNode, PeriodNode: IbjXml;
    InvdtNode, InvvalNode, InvinumNode: IbjXml;
    itemsNode, ItmsNode, ItmDetNode: IbjXml;
    InvTypeNode: IbjXml;

    docdataStr, b2bStr, invStr: UTF8String;
    SupStr, SupGSTNStr, PeriodStr: UTF8String;
    InvdtStr, InvvalStr, InvinumStr: UTF8String;
    typStr: UTF8String;
    itemsStr, txvalStr, rtStr: UTF8String;
    sgstStr, cgstStr, igstStr, cessStr: UTF8String;

    GSTNStr: UTF8String;
    InvTyppeStr: UTF8String;
    IDtStr, ItmsStr, ItmDetStr: UTF8String;
    sAmtStr, cAmtStr, iAmtStr, csAmtStr: UTF8String;
  public
    { Public declarations }
    JFileName: string;
    SaveXMLFile: boolean;
    OnList: TList;
    AppPath: String;
    nInv: Integer;
    IsPurchase: Boolean;
    FUpdate: TfnUpdate;
    XL: TbjXLSTable;
    procedure Process;
    function GetGSTRData: String;
    procedure ParseGSTRStr;
    procedure Load_P_Rec;
    procedure Load_S_Rec;
    procedure Writexls;
    constructor Create;
    destructor Destroy; override;
  end;

implementation

uses GSTRJSON, DataModule;

Constructor TbjJsn2Xls.Create;
begin
  Inherited;
  jStr := CreatebjXmlDocument;
  OnList := TList.Create;
  XL := TbjXLSTable.Create;
  xl.ToSaveFile := True;

  docdataStr := 'docdata';
  b2bStr := 'b2b';
  invStr := 'inv';
  SupStr := 'trdnm';
  SupGSTNStr := 'ctin';
  PeriodStr := 'supprd';
  itemsStr := 'items';
  txvalStr := 'txval';
  rtStr := 'rt';

  InvdtStr := 'dt';
  InvvalStr := 'val';
  InvinumStr := 'inum';
  typStr := 'typ';

  sgstStr := 'sgst';
  cgstStr := 'cgst';
  igstStr := 'igst';
  cessStr := 'cess';

  GSTNStr := 'gstin';
  InvTyppeStr := 'inv_typ';
  IDtStr := 'idt';
  ItmsStr := 'itms';
  ItmDetStr := 'itm_det';

  sAmtStr := 'samt';
  cAmtStr := 'camt';
  iAmtStr := 'iamt';
  csAmtStr := 'csamt';
end;

destructor TbjJsn2Xls.Destroy;
var
  item: pPV;
  i: integer;
begin
  jStr := nil;
  for i := 0 to OnList.Count - 1 do
  begin
    item := OnList.Items[i];
    Dispose(item);
  end;
  OnList.Clear;
  OnList.Free;
  XL.Close;
  XL.Free;
  inherited;
end;

function TbjJsn2Xls.GetGSTRData: String;
begin
  FUpdate('Loading JSON file...');
  jStr.LoadJsnFile(JFileName);
  cnNode := JStr.CloneNode;
  JStr := cnNode.ShortenTree;
  JStr.SaveXmlFile(AppPath + '\Data\Period.xml');
end;

procedure TbjJsn2Xls.process;
begin
  XL.XLSFileName := AppPath + '\Data\Tally2AB.xls';
  XL.SetSheet('JSON');
{.$IFNDEF LXW}
{.$ENDIF}
  XL.SetFields(DataModule.Flds, True);
  XL.GetFields(DataModule.flds);
  GetGSTRData;
  ParseGSTRStr;
  Writexls;
end;

procedure TbjJsn2Xls.ParseGSTRStr;
var
  Item: pPV;
begin
  FUpdate('Parsing JSON file...');
  docdataNode := jStr.SearchForTag(nil, docdataStr);
  if Assigned(docdataNode) then
  begin
    IsPurchase := True;
    b2bNode := docdataNode.SearchForNode(nil, b2bStr);
  end;
  if not Assigned(docdataNode) then
  begin
    GSTNNode := jStr.SearchForTag(nil, GSTNStr);
    if not Assigned(GSTNNode) then
    Exit;
    IsPurchase := False;
    b2bNode := jStr.SearchForNode(GSTNNode, b2bStr)
  end;
  if IsPurchase then
  while Assigned(b2bNode) do
  begin
    invNode := b2bNode.SearchForTag(nil, invStr);
    if not Assigned(invNode) then
      Exit;

    Load_P_Rec;
    b2bNode := docdataNode.SearchForNode(b2bNode, b2bStr);
  end;
  if not IsPurchase then
  while Assigned(b2bNode) do
  begin
    invNode := b2bNode.SearchForTag(nil, invStr);
//    if not Assigned(invNode) then
//      Exit;

    Load_S_Rec;
    b2bNode := jStr.SearchForNode(b2bNode, b2bStr);
  end;
end;

procedure TbjJsn2Xls.Load_P_Rec;
var
  Item: pPV;
  rInvType: UTF8String;
  rInval, rInvDt, rInvNo: UTF8String;
  rSupName, rSupGSTN, rSupPeriod: UTF8String;
  rItemRate, rTaxable, rSGST, rCGST, rIGST, rCESS: UTF8String;
begin
  Item := New(pPV);
  SupNode :=  b2bNode.SearchForTag(nil, SupStr);
  if Assigned(SupNode) then
  rSupName := SupNode.Text;
  Item^.Supplier := rSupName;

  PeriodNode :=  b2bNode.SearchForTag(nil, PeriodStr);
  if Assigned(PeriodNode) then
  rSupPeriod := PeriodNode.Text;
  Item^.PERIOD := rSupPeriod;

  SupGSTNnODE :=  b2bNode.SearchForTag(nil, SupGSTNStr);
  if Assigned(SupGSTNNode) then
  rSupGSTN := SupGSTNNode.Text;
  Item^.GSTN := rSupGSTN;

  rInvDt := invNode.GetChildText(InvdtStr);
  Item^.Invoice_Date := rInvDt;

  rInvaL := invNode.GetChildText(InvvalStr);
  Item^.Invoice_Value := rInval;

  rinvType := invNode.GetChildText(TypStr);
  Item^.Invoice_Type := rinvType;

  rInvNo := invNode.GetChildText(InvinumStr);
  Item^.Invoice_No := rInvNo;

  itemsNode := invNode.SearchForTag(nil, itemsStr);
  while Assigned(itemsNode) do
  begin
    rTaxable := itemsNode.GetChildText(txvalStr);
    Item^.Taxable_Value := rTaxable;
    rItemRate := itemsNode.GetChildText(rtStr);
    Item^.Rate := rItemRate;
    rCGST :=  itemsNode.GetChildText(cgstStr);
    Item^.CGST := rCGST;
    rSGST :=  itemsNode.GetChildText(sgstStr);
    Item^.SGST := rSGST;
    rIGST :=  itemsNode.GetChildText(igstStr);
    Item^.iGST := riGST;
    rCESS :=  itemsNode.GetChildText(cessStr);
    Item^.CESS := rCESS;

    itemsNode := invNode.SearchForTag(itemsNode, itemsStr);
    if Assigned(itemsNode) then
    begin
      OnList.Add(item);
      Item := New(pPV);
      Item^.Supplier := rSupName;
      Item.PERIOD := rSupPeriod;
      Item^.GSTN := rSupGSTN;
      Item^.Invoice_Value := rInval;
      Item^.Invoice_Date := rInvDt;
      Item^.Invoice_No := rInvNo;
    end;
  end;
  OnList.Add(item);
end;

procedure TbjJsn2Xls.Load_S_Rec;
var
  Item: pPV;
  rInvType: UTF8String;
  rInvVal, rInvDt, rInvNo: UTF8String;
  rSupName, rSupGSTN, rSupPeriod: UTF8String;
  rItemRate, rTaxable: UTF8String;
  rSGST, rCGST, rIGST, rCESS: UTF8String;
begin
  Item := New(pPV);
  SupGSTNnODE :=  b2bNode.SearchForTag(nil, SupGSTNStr);
  if Assigned(SupGSTNNode) then
  rSupGSTN := SupGSTNNode.Text;
  Item^.GSTN := rSupGSTN;

  InvvalNode := invNode.SearchForTag(nil, InvvalStr);
  if Assigned(InvvalNode) then
    rInvVal := InvvalNode.Text;
  Item^.Invoice_Value := rInvVal;

  InvTypeNode := invNode.SearchForTag(nil, InvTyppeStr);
  if Assigned(InvTypeNode) then
    rInvType := InvTypeNode.Text;
  Item^.Invoice_Type := rInvType;

  InvDtnODE := invNode.SearchForTag(nil, IDtstr);
  if Assigned(InvDtNode) then
    rInvDt := InvDtNode.Text;
  Item^.Invoice_Date := rInvDt;

  InvinumNode := invNode.SearchForTag(nil, InvinumStr);
  if Assigned(InvinumNode) then
    rInvNo := InvinumNode.Text;
  Item^.Invoice_No := rInvNo;

  ItmsNode := invNode.SearchForTag(nil, itmsStr);
  ItmDetNode := invNode.SearchForTag(ItmsNode, ItmDetStr);
  while Assigned(ItmDetNode) do
  begin
    rTaxable := ItmDetNode.GetChildText(txvalStr);
    Item^.Taxable_Value := rTaxable;
    rItemRate := ItmDetNode.GetChildText(rtStr);
    Item^.Rate := rItemRate;
    rCGST :=  ItmDetNode.GetChildText(cAmtStr);
    Item^.CGST := rCGST;
    rSGST :=  ItmDetNode.GetChildText(sAmtStr);
    Item^.SGST := rSGST;
    rIGST :=  ItmDetNode.GetChildText(iAmtStr);
    Item^.iGST := riGST;
    rCESS :=  ItmDetNode.GetChildText(csAmtStr);
    Item^.CESS := rCESS;

    ItmDetNode := invNode.SearchForTag(ItmDetNode, ItmDetStr);
    if Assigned(ItmDetNode) then
    begin
      OnList.Add(item);
      Item := New(pPV);
      Item^.GSTN := rSupGSTN;
      Item^.Invoice_Value := rInvVal;
      Item^.Invoice_Date := rInvDt;
      Item^.Invoice_No := rInvNo;
    end;
  end;
  OnList.Add(item);
end;

procedure TbjJsn2Xls.Writexls;
var
  i: Integer;
  Item: pPV;
begin
  XL.First;
  for i := 0 to OnList.Count-1 do
  begin
    ppV(item) := OnList.Items[i];

    XL.SetFieldVal('PERIOD', item.PERIOD);
    XL.SetFieldVal('GSTiN', item.GSTN);
    XL.SetFieldVal('INVOICE_NO', item.Invoice_No);
    XL.SetFieldVal('INVOICE_TYPE', item.Invoice_Type);
    XL.SetFieldVal('INVOICE_DATE', item.Invoice_Date);
    if Length(item.Invoice_Value) > 0 then
    XL.SetFieldVal('INVOICE_VALUE', StrToFloat(item.Invoice_Value));
//    XL.SetFieldStr('SUPPLIER', item.Supplier);
    XL.SetFieldVal('SUPPLIER', item.Supplier);
    if Length(item.Taxable_Value) > 0 then
    XL.SetFieldVal('Taxable_Value', StrToFloat(item.Taxable_Value));
    if Length(item.SGST) > 0 then
    XL.SetFieldVal('SGST', StrToFloat(item.SGST));
    if Length(item.CGST) > 0 then
    XL.SetFieldVal('CGST', StrToFloat(item.CGST));
    if Length(item.IGST) > 0 then
    XL.SetFieldVal('IGST', StrToFloat(item.IGST));
    if StrToFloat(item.CESS) <> 0 then
    XL.SetFieldVal('CESS', StrToFloat(item.CESS));
    XL.SetFieldVal('Tax_Rate', item.Rate);
    XL.CurrentRow := XL.CurrentRow + 1;
  end;
  nInv := OnList.Count;
  XL.Save;
end;

end.
