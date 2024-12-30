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
unit TlyInput;

interface

uses
  Classes, SysUtils, DateUtils,
  Dialogs,
  DateFns,
  StrUtils,
  VchLib,
  Client,
  PClientFns,
  ObjCollFn,
  xlstbl313,
  MstListImp,
  bjxml33;

type
  Tfnupdate = procedure(msg: string);

type
  TbjxLedVch4P = class
  private
    { Private declarations }
    FXReq: wideString;
    FHost: string;
    FFirm: String;
    FFrmDt: string;
    FToDT: string;
    FPLedger: string;
    FVendor: string;
    FVchType: string;
    FMasterID: string;
    FPSType: string;
  protected
    xDB: string;
    VendorGSTN: string;
    dmyDate: String;
    VDate, InvNo: string;
    InvAmt: Double;
    lName, pName, lGSTN, lAmt: string;
    Basic, SGST, CGST, IGST, CESS: Double;
    Client: TbjClient;
    Vdoc: IbjXml;
    lID: Ibjxml;
    MList: TbjMstListImp;
    PuList, TXList, PartyList: TStringList;
    FIsTaxInv: Boolean;
  public
    { Public declarations }
    ndups: Integer;
    DBKDt: string;
    CurDt: string;
    FUpdate: TfnUpdate;
    dbkDoc: Ibjxml;
    kadb: TbjXLSTable;
    tgDoc: Ibjxml;
    lDoc: Ibjxml;
    procedure GetLedVchRgtr(const aLedger: string);
    procedure GetVouXML;
    procedure CalcGSTN;
    procedure CalcAmt;
    procedure WriteXls;
    procedure Process;
    procedure ParseVch;
    procedure LoadDefault;
    procedure SetHost(const aHost: string);
    constructor Create;
    destructor Destroy; override;
  published
    property Firm: String read FFirm write FFirm;
    property Host: String read FHost write SetHost;
    property FrmDt: string Read FFrmDt write FFrmDt;
    property ToDt: string Read FToDt write FToDt;
    property Ledger: string Read FPLedger write FPLedger;
    property Vendor: string Read FVendor write FVendor;
    property VchType: string Read FVchType write FVchType;
    property MasterId: string Read FMasterId write FMasterId;
    property psType: string read FPSType write FPSType;
    property IsTaxInv: Boolean read FIsTaxInv write FIsTaxInv;
  end;

implementation

uses DataModule;

Constructor TbjxLedVch4P.Create;
//var
{  Obj: TxlsBlank; }
begin
  FHost := 'http://127.0.0.1:9000';
  Client := TbjClient.Create;
{
  Obj := TxlsBlank.Create;
  Obj.IsRecreateOn := True;
  Obj.CreateTally2A;
  Obj.Free;
}
  kadb := TbjXLSTable.Create;
  MList := TbjMstListImp.Create;
  TxList := TStringList.Create;
  PuList := TStringList.Create;
  PartyList := TStringList.Create;
  kadb.SetXLSFile('.\Data\Tally2AB.xls');
  kadb.SetSheet('BOOKS');
{.$IFDEF LXW }
//  kadb.SetFields(flds, True);
{.$ENDIF LXW }
  kadb.SetFields(flds, True);
  kadb.GetFields(flds);
  kadb.First;
  CurDt := '19560229';
  if not Assigned(ClientFns) then
     ClientFns := TbjGSTClientFns.Create;
  dbkDoc := CreatebjXmlDocument;
  tgDoc :=  CreatebjXmlDocument;;
  lDoc :=  CreatebjXmlDocument;;
end;

destructor TbjxLedVch4P.Destroy;
begin
  Client.Free;
  PartyList.Free;
  PuList.Clear;
  PuList.Free;
  TxList.Clear;
  TxList.Free;
  MList.Free;
  kadb.Free;
  dbkDoc :=nil;
  inherited;
end;

procedure TbjxLedVch4P.SetHost(const aHost: string);
begin
  Client.Host := aHost;
  ClientFns.Host := aHost;
  FHost := aHost;
end;

procedure TbjxLedVch4P.Process;
var
  rID: Ibjxml;
  VTypeNode, LedNode: IbjXml;
  DtNode, grpMIdNode: IbjXml;
  rDt: string;
begin
  PartyList.Text := MList.GetPartyText(True);
  PartyList.Sorted := True;
  if psType = 'Purchase' then
    PuList.Text := MList.GetGrpLedText('Purchase Accounts');
//  ShowMessage(PuList.Text);  
  if psType = 'Sales' then
    PuList.Text := MList.GetGrpLedText('Sales Accounts');
  PuList.Sorted := True;
  TxList.Text := MList.GetTaxText;
  TxList.Sorted := True;
  if Length(Vendor) > 0 then
    VendorGSTN := GetLedgerGSTIN(Vendor);
  GetLedVchRgtr(Ledger);
  if Length(VendorGSTN) = 0 then
  if IsTaxInv then
  Exit;
  rID := lDoc.SearchForTag(nil, 'COLLECTION');
  if Assigned(rID) then
//    rID := lDoc.SearchForNode(rID, 'VOUCHER', False);
    rID := lDoc.SearchForNode(rID, 'VOUCHER');
  while Assigned(rID) do
  begin
    VTypeNode := rID.SearchForTag(nil, 'VOUCHERTYPENAME');
    if Assigned(VTypeNode) then
    if VTypeNode.GetContent <> VchType then
    begin
      rID := lDoc.SearchForNode(rID, 'VOUCHER');
      Continue;
    end;

    LedNode := rID.SearchForTag(nil, 'LEDGERNAME');
    if Assigned(LedNode) then
    if LedNode.GetContent <> Vendor then
    begin
      rID := lDoc.SearchForNode(rID, 'VOUCHER');
      Continue;
    end;

    DtNode := rID.SearchForTag(nil, 'DATE');
    if Assigned(DtNode) then
      rDt := DtNode.Text;
    if (rDt < FrmDt) or (rDt > ToDt)then
    begin
      rID := lDoc.SearchForNode(rID, 'VOUCHER');
//      FUpdate('Searched Date: ' + rDt);
      Continue;
    end;
    CurDt := rDt;

    grpMIdNode := rID.SearchForTag(nil, 'MASTERID');
    if Assigned(grpMIdNode) then
      MasterId := grpMIdNode.Text;

    FUpdate('Checking Date: ' + CurDt);
    GetVouXML;
    ParseVch;;
    rID := lDoc.SearchForNode(rID, 'VOUCHER');
  end;
end;

procedure TbjxLedVch4P.GetLedVchRgtr(const aLedger: string);
var
  lDB: string;
  Debug: Ibjxml;
  VchColName: string;
begin
  if Length(Vendor) > 0 then
    VchColName := 'vouchers of Ledger'
  ELSE
    VchColName := 'Vouchers of Group';
  FxReq := '<ENVELOPE><HEADER><VERSION>1</VERSION>'+
  '<TALLYREQUEST>Export</TALLYREQUEST>' +
  '<TYPE>COLLECTION</TYPE><ID>' + VchColName +'</ID></HEADER>'+
  '<BODY><DESC><STATICVARIABLES>';

  FxReq := FXReq + '<SVEXPORTFORMAT>' + '$$SysName:XML' + '</SVEXPORTFORMAT>';

  if (Length(FFirm) <> 0) then
    FxReq := FXReq + '<SVCURRENTCOMPANY>' + FFirm + '</SVCURRENTCOMPANY>';
  if ( Length(FFrmDt) <> 0 ) and ( Length(FToDt) <> 0 ) then
  begin
    FxReq := FXReq + '<SVFROMDATE Type="Date">' + FFrmDt + '</SVFROMDATE>';
    FxReq := FXReq + '<SVTODATE Type="Date">' + FToDt + '</SVTODATE>';
  end;

  if Length(Vendor) > 0 then
  FxReq := FXReq + '<LEDGERNAME>' +  Vendor + '</LEDGERNAME>'
  else
  FxReq := FXReq + '<GROUPNAME>Duties &amp; Taxes</GROUPNAME>';

  FxReq := FXReq +
  '</STATICVARIABLES><TDL><TDLMESSAGE>'+
  '<Add>SVFromDATE</Add><Add>SVToDATE</Add>'+
  '<COLLECTION NAME="' + VchColName + '" ISINITIALIZE="Yes">'+
  '<Type>Voucher</Type>'+
  '</COLLECTION>'+
  '</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>';

  Client.Host := FHost;
  Client.xmlRequestString := Fxreq;
  Client.post;
  ldb := Client.xmlResponseString;
  lDoc.LoadXML(lDB);
{
  Debug := CreatebjXmlDocument;
  Debug.LoadXml(ldb);
  Debug.SaveXmlFile('LedVchRes.xml');
}
end;


procedure TbjxLedVch4P.GetVouXML;
var
  Debug: Ibjxml;
begin
  if DBKDt = CurDt then
    Exit;
  FxReq := '<ENVELOPE><HEADER><VERSION>1</VERSION>';
  FxReq := FxReq + '<TALLYREQUEST>EXPORT</TALLYREQUEST>';
  FxReq := FxReq + '<TYPE>DATA</TYPE>';
  FxReq := FxReq + '<ID>Daybook</ID>';
  FxReq := FxReq + '</HEADER><BODY><DESC>';

  FxReq := FxReq + '<STATICVARIABLES>';

  FxReq := FXReq + '<SVEXPORTFORMAT>' + '$$SysName:XML' + '</SVEXPORTFORMAT>';

  if (Length(FFirm) <> 0) then
    FxReq := FXReq + '<SVCURRENTCOMPANY>' + FFirm + '</SVCURRENTCOMPANY>';

  FxReq := FXReq + '<SVCURRENTDATE>' + CurDt + '</SVCURRENTDATE>';

  if VchType = 'Purchase' then
    FxReq := FXReq + '<VoucherTypeName>' +  '$$VchTypePurchase' + '</VoucherTypeName>';
  if VchType = 'Sales' then
    FxReq := FXReq + '<VoucherTypeName>' +  '$$VchTypeSales' + '</VoucherTypeName>';
  if Length(Vendor) > 0 then
  FxReq := FXReq + '<LEDGERNAME>' +  Vendor + '</LEDGERNAME>';

  FxReq := FXReq + '</STATICVARIABLES>';

  FxReq := FXReq + '</DESC></BODY></ENVELOPE>';

  Client.Host := FHost;
  Client.xmlRequestString := Fxreq;

  Client.post;

  xdb := Client.xmlResponseString;
{
  Debug := CreatebjXmlDocument;
//  Debug.LoadXML(xDB);
//  Debug.SaveXmlFile('Daybook.xml');

  Debug.LoadXML(FXReq);
  Debug.SaveXmlFile('Data.xml');
}
end;

{ rfReplaceAll rfIgnoreCase }
procedure TbjxLedVch4P.ParseVch;
var
//  item: pPV;
  xDate: Ibjxml;
  dbkMIDNode: IbjXml;
  IsFound: Boolean;
  rLevel, r2Level: Integer;
begin
  rlevel := 0;
  if DBKDt <> CurDt then
  begin
    DBKDt := CurDt;
    dbkDoc.Clear;
    dbkDoc.LoadXML(xDB);
    DBKDoc.Prune(nil);
  end;
  vDoc := dbkDoc.SearchForTag(nil, 'TALLYMESSAGE');
  VDoc := dbkDoc.SearchForTag(VDoc, 'VOUCHER');
//  ShowMessage(IntToStr(rLevel));
  while Assigned(vDoc) do
  begin
    xDate := VdOC.SearchForTag(nil, 'DATE');
    VDate := xDate.GetContent;
    if (Length(VDate) = 0) or (VDate < FrmDt) or (VDate > ToDt) then
    begin
//      vDoc := dbkDoc.BFSearchForTag(VDoc, 'VOUCHER', rLevel);
      vDoc := dbkDoc.SearchForNode(VDoc, 'VOUCHER');
      Continue;
    end;
    dbkMIDNode := vDoc.SearchForTag(nil, 'MASTERID');
    IF Assigned(dbkMIDNode) then
    if dbkMIDNode.Text <> MasterId then
    begin
//      vDoc := dbkDoc.BFSearchForTag(VDoc, 'VOUCHER', rLevel);
      vDoc := dbkDoc.SearchForNode(VDoc, 'VOUCHER');
      Continue;
    end;
    IsFound := True;
    FUpdate('Importing Date: ' + CurDt + ' MasterID: ' +MasterId);

    InvNo := '';
    lName := '';
    pName := '';
    lGSTN := '';
    lAmt := '';
    Basic := 0;
    SGST := 0;
    CGST := 0;
    IGST := 0;
    CESS := 0;

    r2Level := 0;
    CalcGSTN;
    lID := VDoc.SearchForTag(nil, 'LEDGERNAME');
    //GSTNNode := Vdoc.SearchForTag(nil, 'PARTYNAME');
    while Assigned(lID) do
    begin
      CalcAmt;
      lID := VDoc.BFSearchForTag(lID, 'LEDGERNAME', r2Level);
    end;
//    ShowMessage(IntToStr(r2Level));
    WriteXls;
    if IsFound then
      Break;
//    vDoc := dbkDoc.BFSearchForTag(VDoc, 'VOUCHER', rLevel);
    vDoc := dbkDoc.SearchForNode(VDoc, 'VOUCHER');
  end;

  kadb.Save;
end;

procedure TbjxLedVch4P.CalcGSTN;
var
  GSTNNode: IbjXml;
  AmtID: Ibjxml;
  xRef: IbjXml;
  idx: integer;
begin
  GSTNNode := Vdoc.SearchForTag(nil, 'PARTYLEDGERNAME');
  if Assigned(GSTNNode) then
    PName := GSTNNode.Text;

  GSTNNode := Vdoc.SearchForTag(nil, 'PARTYGSTIN');
  if Assigned(GSTNNode) then
    lGSTN := GSTNNode.Text;
  if Length(lGSTN) > 0 then
    VendorGSTN := lGSTN;

  xRef := VDoc.SearchForTag(nil, 'REFERENCE');
  if Assigned(xRef) then
    InvNo := xRef.GetContent;
end;

procedure TbjxLedVch4P.CalcAmt;
var
//  Obj: TxlsBlank;
  GSTNNode: IbjXml;
  AmtID: Ibjxml;
  xRef: IbjXml;
  idx: integer;
begin

  lName := lID.GetContent;
  AmtID := VDoc.SearchForTag(lID, 'AMOUNT');
  if Assigned(AmtID) then
    lAmt := AmtID.GetContent;
  if (lName = Vendor) or (lName = pName) then
    InvAmt := StrToFloat(lAmt);
  if (TxList.Find(PackStr(lName),  idx)) and (Pos('SGST', UpperCase(lName)) > 0)  then
  begin
      SGST := SGST + StrToFloat(lAmt);
    end;
  if (TxList.Find(PackStr(lName), idx)) and (Pos('CGST', UpperCase(lName)) > 0)  then
  begin
      CGST := CGST + StrToFloat(lAmt);
    end;
  if (TxList.Find(PackStr(lName), idx)) and (Pos('IGST', UpperCase(lName)) > 0)  then
    begin
      IGST := IGST + StrToFloat(lAmt);
    end;
  if (TxList.Find(PackStr(lName), idx)) and (Pos('CESS', UpperCase(lName)) > 0)  then
    CESS := CESS + StrToFloat(lAmt);

//  if PuList.Find(PackStr(lName), idx) then
  if PuList.Find(lName, idx) then
  begin
      Basic := Basic + StrToFloat(lAmt);
  end;
end;

procedure TbjxLedVch4P.WriteXls;
var
  mthDate: string;
begin
  dmyDate := Copy(VDate, 7, 2) + '-' + Copy(VDate, 5, 2) +
  '-'  + Copy(VDate, 3, 2);
//  mthDate := LongMonthNames(Copy(VDate, 5, 2));
  FUpdate('Writing Date: ' + CurDt + ' MasterID: ' +MasterId);

  IF Basic <> 0 then
//    kadb.SetFieldVal('Taxable_Value', FloatToStr(Abs(Basic)));
//  FUpdate('Writing Date: ' + CurDt + ' MasterID: ' +MasterId + ' Taxable_Value');
//  IF SGST > 0 then
//    kadb.SetFieldVal('SGST', FloatToStr(Abs(SGST)));
//  IF CGST > 0 then
//    kadb.SetFieldVal('CGST', FloatToStr(Abs(CGST)));
//  IF IGST > 0 then
    kadb.SetFieldVal('Taxable_Value', Abs(Basic));
//    kadb.SetFieldVal('IGST', Abs(IGST));
    kadb.SetFieldVal('SGST', Abs(SGST));
    kadb.SetFieldVal('CGST', Abs(CGST));
    IF IGST <> 0 then
    kadb.SetFieldVal('IGST', Abs(IGST));
    IF CESS <> 0 then
    kadb.SetFieldVal('CESS', Abs(CESS));
      LoadDefault;
      kadb.SetFieldVal('Tax_Rate', '');
      kadb.CurrentRow := kadb.CurrentRow + 1;
      ndups := ndups + 1;
  FUpdate('Writing Date: ' + CurDt + ' MasterID: ' +MasterId + ' Done');
end;

procedure TbjxLedVch4P.LoadDefault;
var
  rVal: TVarRec;
begin
  FUpdate('Writing Date: ' + CurDt + ' MasterID: ' +MasterId + ' Default');
//  kadb.SetFieldVal('PERIOD', dmyDate);
  kadb.SetFieldVal('PERIOD', '');

  kadb.SetFieldVal('GSTIN', VendorGSTN);
{
  rVal.VType := vtPChar;
  rVal.VPChar := PChar(VendorGSTN);
  kadb.SetFieldVal('GSTIN', rVal, nil);
}
  kadb.SetFieldVal('GSTIN', VendorGSTN);
  kadb.SetFieldVal('Invoice_No', InvNo);
  kadb.SetFieldVal('Invoice_Date', dmyDate);
//  kadb.SetFieldVal('Invoice_Value', FloatToStr(Abs(InvAmt)));
  kadb.SetFieldVal('Invoice_Value', Abs(InvAmt));
  if Length(Vendor) > 0 then
  kadb.SetFieldVal('Supplier', Vendor);
  if Length(pName) > 0 then
  kadb.SetFieldVal('Supplier', pName);
  kadb.SetFieldVal('Tax_Rate', '');
end;

end.
