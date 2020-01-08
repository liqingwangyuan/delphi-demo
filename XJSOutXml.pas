unit XJSOutXml;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, UBaseAAC, PrnDbgeh, DB, DBClient, Menus, ImgList, StdActns,
  StdCtrls, DBActns, ActnList, XPMan, fcdbtreeview, Buttons, Grids,
  DBGridEh, ExtCtrls, ComCtrls, ToolWin,UBase,UBaseA, CommProc,
  sysBaseBL, UPubVar, USysComDef,sysPubProc,Waiting, XMLDoc, XMLIntf;

type
  TfrmXJSOutXml = class(TfrmBaseAAC)
    dtpBeg: TDateTimePicker;
    dtpEnd: TDateTimePicker;
    btn1: TToolButton;
    lbl1: TLabel;
    lbl2: TLabel;
    procedure FormShow(Sender: TObject);
    procedure dgrdsearchKeyPress(Sender: TObject; var Key: Char);
    procedure dgrdsearchCellClick(Column: TColumnEh);
    procedure btn1Click(Sender: TObject);
    procedure actRefreshExecute(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  function ToXMLFile(cdsOutData: TClientDataSet; sDate,sNSRMC,sNSRSBH: string; sNo: Integer): string;
  end;

var
  frmXJSOutXml: TfrmXJSOutXml;
  
function CreatefrmXJSOutXml(app: TApplication; pCall: Pointer): TfrmBase; stdCall;


implementation

{$R *.dfm}

function CreatefrmXJSOutXml(app: TApplication; pCall: Pointer): TfrmBase; stdcall;
begin
  Application := app;
  result := TfrmXJSOutXml.create(application);
  result.FrmFreeCallBack := pCall;
end;

procedure TfrmXJSOutXml.FormShow(Sender: TObject);
var
  i:Integer;
  vData: OleVariant;
  iRtnCode: Integer;
  sMsg : WideString;
begin
  TableMasterID := 5505;
  AutoRank:=false;
  FPCDSPropertyMain.Fclientdataset.IsSelect := True;
  DoAddSearch(FPCDSPropertyMain.Fclientdataset,GBWhere);
  DoAddDbgrid(FPCDSPropertyMain.Fclientdataset,dgrdSearch);
  inherited;
  fcDBTreeView1.Visible:=False;
  dtpBeg.DateTime:=Now;     
  dtpEnd.DateTime:=Now;
  actRefreshExecute(nil);
end;

procedure TfrmXJSOutXml.dgrdsearchKeyPress(Sender: TObject; var Key: Char);
begin
  inherited;
 { if Key = #32 then
  begin
    FPCDSPropertyMain.Fclientdataset.Edit;
      FPCDSPropertyMain.Fclientdataset.FieldByName('Is_Sel$').AsString := '1';
  end;}
end;

procedure TfrmXJSOutXml.dgrdsearchCellClick(Column: TColumnEh);
begin
  inherited;
  if UpperCase(Column.FieldName) = 'IS_SEL$' then
    FPCDSPropertyMain.Fclientdataset.SelectValue := not FPCDSPropertyMain.Fclientdataset.SelectValue;
end;

procedure TfrmXJSOutXml.btn1Click(Sender: TObject);
var
  vData,vOutData: OleVariant;
  sMsg,sXmlMsg: WideString;
  iRtnCode,iAllCount,iRowCount,iOutCount: Integer;
  sSql,sOutSql,sOrgCode,sNowDate,sBegDate,sEndDate,sMonDate: string;
  cdsTmpData,cdsOutData: TClientDataSet;
  sNSRSBH,sNSRMC: string;

  Xml:   TXMLDocument;
  ROOT,HZINFO,BASEINFO: IXMLNode;
  NSRSBH,NSRMC,QSYF,ZZYF,ZCFPYL,SHZFSL: IXMLNode;
begin
  if FormatDateTime('yyyyMM',dtpBeg.Date)<>FormatDateTime('yyyyMM',dtpEnd.Date) then
  begin
    UApp.ErrorMsgBox('导出数据不允许跨月！请按月导出');
    Abort;
  end;
  if not UApp.ConfirmMsgDlg('您确认要按选中组织和设置日期导出文件吗？') then Exit;
  try
    sNSRSBH:='00X';
    sNSRMC:='**有限公司';
    sNowDate:=FormatDateTime('yyyyMMdd',Now());
    FPCDSPropertyMain.Fclientdataset.First;
    while not FPCDSPropertyMain.Fclientdataset.Eof do
    begin
      if FPCDSPropertyMain.Fclientdataset.FieldByName('Is_Sel$').AsString = '1' then
      begin
        if sOrgCode = '' then
          sOrgCode := ''''+FPCDSPropertyMain.Fclientdataset.FieldByName('OrgCode').AsString+''''
        else
          sOrgCode := sOrgCode + ',''' + FPCDSPropertyMain.Fclientdataset.FieldByName('OrgCode').AsString + '''';
      end;
      FPCDSPropertyMain.Fclientdataset.Next;
    end;
    if sOrgCode = '' then
    begin
      UApp.ErrorMsgBox('请先选择组织再执行导出！');
      Abort;
    end;
    sBegDate := FormatDateTime('yyyy-MM-dd',dtpBeg.Date);
    sEndDate := FormatDateTime('yyyy-MM-dd',dtpEnd.Date);
    sMonDate := FormatDateTime('yyyyMM',dtpBeg.Date);
    sSql := 'select A.OrgCode||A.PosNo||A.SaleNo as WYM, to_char(to_date(A.JzDate,''yyyy-mm-dd''),''yyyymmdd'') as KPRQ, '
          + '       A.XsDate,count(1) as SPHSL,Sum(B.SsTotal) as SPHJJE '
          + '  from tSalSale'+sMonDate+' A, tSalSalePlu'+sMonDate+' B '
          + ' where A.TranType = ''1'' and A.OrgCode = B.OrgCode  and A.SaleNo = B.SaleNo '
          + '   and A.JzDate between '''+sBegDate+''' and  '''+sEndDate+''' '
          + '   and A.OrgCode in(' + sOrgCode+ ')  '
          + ' group by A.OrgCode, A.PosNo, A.SaleNo, A.JzDate, A.XsDate';
    iRtnCode := GetIcommDisp.OpenQuery(CurLoginId, sMsg, sSql, vData, false);
    if (iRtnCode = 1) then
    begin
      cdsTmpData := TClientDataSet.Create(nil);
      try
        cdsTmpData.Data := UVDeCodeClient(vData);
        iAllCount:=cdsTmpData.RecordCount;
        if cdsTmpData.IsEmpty then
        begin
          UApp.ErrorMsgBox('不存在待导出数据！');
          Abort;
        end;

        //导出XML主表
        Xml:=TXMLDocument.Create(nil);
        try
            //加入版本信息 ‘<?xml version="1.0" encoding="UTF-8" ?> ’
            Xml.Active := True;
            Xml.Version := '1.0';
            Xml.Encoding :='GBK';
            ROOT := Xml.CreateNode('ROOT');
            Xml.DocumentElement := ROOT;

            HZINFO := Xml.CreateNode('HZINFO');
            HZINFO.SetAttribute('class','HZINFO');
            ROOT.ChildNodes.Add(HZINFO);
      
            BASEINFO := Xml.CreateNode('BASEINFO');
            BASEINFO.SetAttribute('class','BASEINFO');
            HZINFO.ChildNodes.Add(BASEINFO);

            NSRSBH := Xml.CreateNode('NSRSBH');
            NSRSBH.NodeValue:=sNSRSBH;
            BASEINFO.ChildNodes.Add(NSRSBH);

            NSRMC := Xml.CreateNode('NSRMC');
            NSRMC.NodeValue:=sNSRMC;
            BASEINFO.ChildNodes.Add(NSRMC);

            QSYF := Xml.CreateNode('QSYF');
            QSYF.NodeValue:=FormatDateTime('yyyyMM',dtpBeg.Date);
            BASEINFO.ChildNodes.Add(QSYF);
      
            ZZYF := Xml.CreateNode('ZZYF');
            ZZYF.NodeValue:=FormatDateTime('yyyyMM',dtpEnd.Date);
            BASEINFO.ChildNodes.Add(ZZYF);
      
            ZCFPYL := Xml.CreateNode('ZCFPYL');
            ZCFPYL.NodeValue:=inttostr(iAllCount);
            BASEINFO.ChildNodes.Add(ZCFPYL);

            SHZFSL := Xml.CreateNode('SHZFSL');
            SHZFSL.NodeValue:='0';
            BASEINFO.ChildNodes.Add(SHZFSL);

            Xml.SaveToFile('c:\TAX_FPXX_HZ_'+sNowDate+'_'+sNSRSBH+'.xml');
            xml.Active := False;
        finally
          Xml.Free;
        end;

        //导出XML明细
        sOutSql := 'select A.OrgCode||A.PosNo||A.SaleNo as WYM, A.JzDate as KPRQ, '
          + '       A.XsDate,count(1) as SPHSL,Sum(B.SsTotal) as SPHJJE '
          + '  from tSalSale A, tSalSalePlu B '
          + ' where 1=2  '
          + ' group by A.OrgCode, A.PosNo, A.SaleNo, A.JzDate, A.XsDate';
        iRtnCode := GetIcommDisp.OpenQuery(CurLoginId, sMsg, sOutSql, vOutData, false);
        if (iRtnCode <> 1) then
        begin
          UApp.ErrorMsgBox('获取导出数据失败！');
          Abort;
        end;
        cdsOutData := TClientDataSet.Create(nil);
        cdsOutData.Data := UVDeCodeClient(vOutData);

        iRowCount:=0;   //XML导出行数
        iOutCount:=0;   //导出明细XML文件的序号
        cdsTmpData.First;
        while not cdsTmpData.Eof do
        begin
          iRowCount := iRowCount + 1;
          cdsOutData.Append;
          cdsOutData.FieldByName('WYM').AsString:=cdsTmpData.FieldByName('WYM').AsString;
          cdsOutData.FieldByName('KPRQ').AsString:=cdsTmpData.FieldByName('KPRQ').AsString;
          cdsOutData.FieldByName('SPHSL').AsString:=cdsTmpData.FieldByName('SPHSL').AsString;
          cdsOutData.FieldByName('SPHJJE').AsString:=cdsTmpData.FieldByName('SPHJJE').AsString;
          cdsOutData.Post;
          if iRowCount=100000 then
          begin
            iOutCount := iOutCount + 1;
            iRowCount := 0;
            sXmlMsg := ToXMLFile(cdsOutData, sNowDate, sNSRMC, sNSRSBH, iOutCount);
            if sXmlMsg <> '1' then
            begin
              UApp.ErrorMsgBox('导出数据失败！' + sXmlMsg);
              Abort;
            end;
            cdsOutData.First;
            while not cdsOutData.Eof do
            begin
              cdsOutData.Delete;
            end;
          end;
          cdsTmpData.Next;
        end;
        if iRowCount > 0 then
        begin
          iOutCount := iOutCount + 1;
          sXmlMsg := ToXMLFile(cdsOutData, sNowDate, sNSRMC, sNSRSBH, iOutCount);
          if sXmlMsg <> '1' then
          begin
            UApp.ErrorMsgBox('导出数据失败！' + sXmlMsg);
            Abort;
          end;
        end;
      finally
        cdsTmpData.Free;
      end;
    end
    else
    begin
      UApp.ErrorMsgBox('查询导出数据失败！');
      Abort;
    end;
    UApp.MsgBox('导出数据成功！');
  finally
    
  end;
end;

function TfrmXJSOutXml.ToXMLFile(cdsOutData: TClientDataSet; sDate,sNSRMC,sNSRSBH: string; sNo: Integer): string;
var
  vData: OleVariant;
  sMsg: widestring;
  sSql: string;
  iCount: Integer;
  
  Xml:   TXMLDocument;
  ROOT,UPINVINFO,BASEINFO,FPHZXX_JLS,FPHZXX_JL,FPHZXX: IXMLNode;
  NSRSBH,NSRMC,WYM,KPRQ,SPHSL,SPHJJE: IXMLNode;
  FPDM,FPHM,YFPDM,YFPHM,FKFDM,FKFMC,FKTKZT,DBFPSL,SJKPJE: IXMLNode;
begin
  Result := '0';
  try
    iCount:=cdsOutData.RecordCount;
    cdsOutData.First;
    with cdsOutData do
    begin
      Xml:=TXMLDocument.Create(nil);
      try
        //加入版本信息 ‘<?xml version="1.0" encoding="UTF-8" ?> ’
        Xml.Active := True;
        Xml.Version := '1.0';
        Xml.Encoding :='GBK';
        ROOT := Xml.CreateNode('ROOT');
        Xml.DocumentElement := ROOT;

        UPINVINFO := Xml.CreateNode('UPINVINFO');
        UPINVINFO.SetAttribute('class','UPINVINFO');
        ROOT.ChildNodes.Add(UPINVINFO);

        BASEINFO := Xml.CreateNode('BASEINFO');
        BASEINFO.SetAttribute('class','BASEINFO');
        BASEINFO.SetAttribute('version','1.0');
        UPINVINFO.ChildNodes.Add(BASEINFO);

        NSRSBH := Xml.CreateNode('NSRSBH');
        NSRSBH.NodeValue:=sNSRSBH;
        BASEINFO.ChildNodes.Add(NSRSBH);

        NSRMC := Xml.CreateNode('NSRMC');
        NSRMC.NodeValue:=sNSRMC;
        BASEINFO.ChildNodes.Add(NSRMC);

        FPHZXX_JLS := Xml.CreateNode('FPHZXX_JLS');
        FPHZXX_JLS.SetAttribute('class','FPHZXX_JL;');
        FPHZXX_JLS.SetAttribute('size',IntToStr(iCount));
        UPINVINFO.ChildNodes.Add(FPHZXX_JLS);
        while not Eof do
        begin
          FPHZXX_JL := Xml.CreateNode('FPHZXX_JL');
          FPHZXX_JLS.ChildNodes.Add(FPHZXX_JL);

          FPHZXX := Xml.CreateNode('FPHZXX');
          FPHZXX.SetAttribute('class','FPHZXX');
          FPHZXX_JL.ChildNodes.Add(FPHZXX);

          WYM := Xml.CreateNode('WYM');
          WYM.NodeValue:=cdsOutData.FieldByName('WYM').AsString;
          FPHZXX.ChildNodes.Add(WYM);

          FPDM := Xml.CreateNode('FPDM');
          FPDM.NodeValue:='';
          FPHZXX.ChildNodes.Add(FPDM);
          
          FPHM := Xml.CreateNode('FPHM');
          FPHM.NodeValue:='';
          FPHZXX.ChildNodes.Add(FPHM);

          YFPDM := Xml.CreateNode('YFPDM');
          YFPDM.NodeValue:='';
          FPHZXX.ChildNodes.Add(YFPDM);
          
          YFPHM := Xml.CreateNode('YFPHM');
          YFPHM.NodeValue:='';
          FPHZXX.ChildNodes.Add(YFPHM);

          KPRQ := Xml.CreateNode('KPRQ');
          KPRQ.NodeValue:=cdsOutData.FieldByName('KPRQ').AsString;
          FPHZXX.ChildNodes.Add(KPRQ);

          FKFDM := Xml.CreateNode('FKFDM');
          FKFDM.NodeValue:='';
          FPHZXX.ChildNodes.Add(FKFDM);

          FKFMC := Xml.CreateNode('FKFMC');
          FKFMC.NodeValue:='';
          FPHZXX.ChildNodes.Add(FKFMC);

          FKTKZT := Xml.CreateNode('FKTKZT');
          if (cdsOutData.FieldByName('SPHJJE').AsCurrency < 0) then
          begin
            FKTKZT.NodeValue:='2';
            YFPDM.NodeValue:='有效证明';
            YFPHM.NodeValue:='有效证明';
          end
          else
          begin
            FKTKZT.NodeValue:='1';
          end;
          FPHZXX.ChildNodes.Add(FKTKZT);
          
          DBFPSL := Xml.CreateNode('DBFPSL');
          DBFPSL.NodeValue:='1';
          FPHZXX.ChildNodes.Add(DBFPSL);

          SPHSL := Xml.CreateNode('SPHSL');
          SPHSL.NodeValue:=cdsOutData.FieldByName('SPHSL').AsString;
          FPHZXX.ChildNodes.Add(SPHSL);

          SPHJJE := Xml.CreateNode('SPHJJE');
          SPHJJE.NodeValue:=cdsOutData.FieldByName('SPHJJE').AsString;
          FPHZXX.ChildNodes.Add(SPHJJE);

          SJKPJE := Xml.CreateNode('SJKPJE');
          SJKPJE.NodeValue:=cdsOutData.FieldByName('SPHJJE').AsString;
          FPHZXX.ChildNodes.Add(SJKPJE);
          Next;
        end;
        Xml.SaveToFile('c:\TAX_FPXX_'+sDate+'_'+sNSRSBH+'_'+inttostr(sNo)+'.xml');
        xml.Active := False;
      finally
        Xml.Free;
      end;
    end;
    result := '1';
  except
    on E: Exception do
    begin
      UApp.ErrorMsgBox(E.Message);
      result := E.Message;
    end;
  end;
end;

procedure TfrmXJSOutXml.actRefreshExecute(Sender: TObject);
begin
  inherited;
  FPCDSPropertyMain.Fclientdataset.First;
  while not FPCDSPropertyMain.Fclientdataset.Eof do
  begin
    FPCDSPropertyMain.Fclientdataset.SelectValue := True;
    FPCDSPropertyMain.Fclientdataset.Next;
  end;
  FPCDSPropertyMain.Fclientdataset.First;
end;

end.
