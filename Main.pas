unit Main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Uni, Data.DB, MemDS, DBAccess,
  Vcl.ComCtrls, Vcl.ToolWin, UniProvider, OracleUniProvider, System.ImageList,
  Vcl.ImgList, cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters,
  cxStyles, dxSkinsCore, dxSkinscxPCPainter, cxCustomData, cxFilter, cxData,
  cxDataStorage, cxEdit, cxNavigator, cxDBData, cxCurrencyEdit, cxContainer,
  dxCore, cxDateUtils, Vcl.Menus, Vcl.StdCtrls, cxButtons, cxTextEdit,
  cxMaskEdit, cxDropDownEdit, cxCalendar, cxGridCustomTableView, cxGridExportLink,
  cxGridTableView, cxGridDBTableView, cxGridLevel, cxClasses, cxGridCustomView,
  cxGrid, Vcl.Buttons, cxGridCustomPopupMenu, cxGridPopupMenu, dxDateRanges,
  cxDataControllerConditionalFormattingRulesManagerDialog, dxBarBuiltInMenu;

type
  TfrmExtReyWin = class(TForm)
    pc1: TPageControl;
    tb1: TToolBar;
    tsSeleccion: TTabSheet;
    tsGrilla: TTabSheet;
    tbtImportar: TToolButton;
    tbtSalir: TToolButton;
    cnn: TUniConnection;
    qPERIODOS_TRABAJADOS: TUniQuery;
    qSPImportar: TUniStoredProc;
    OracleUniProvider1: TOracleUniProvider;
    ImageList1: TImageList;
    cxGrid1DBTableView1: TcxGridDBTableView;
    cxGrid1Level1: TcxGridLevel;
    cxGrid1: TcxGrid;
    dsPERIODOS_TRABAJADOS: TDataSource;
    cxGrid1DBTableView1PER_ID: TcxGridDBColumn;
    cxGrid1DBTableView1PER_DESDE: TcxGridDBColumn;
    cxGrid1DBTableView1PER_HASTA: TcxGridDBColumn;
    cxGrid1DBTableView1PER_LEGAJO: TcxGridDBColumn;
    cxGrid1DBTableView1EVE_ID_E: TcxGridDBColumn;
    cxGrid1DBTableView1EVE_ID_S: TcxGridDBColumn;
    cxGrid1DBTableView1PER_HORAS: TcxGridDBColumn;
    cxGrid1DBTableView1LEG_NOMBRE: TcxGridDBColumn;
    Label1: TLabel;
    Label2: TLabel;
    cxFechaDesde: TcxDateEdit;
    cxFechaHasta: TcxDateEdit;
    fsDialog: TFileSaveDialog;
    tbtRango: TToolButton;
    cdbRangoFechasOK: TSpeedButton;
    qPERIODOS_TRABAJADOSPER_ID: TFloatField;
    qPERIODOS_TRABAJADOSPER_DESDE: TDateTimeField;
    qPERIODOS_TRABAJADOSPER_HASTA: TDateTimeField;
    qPERIODOS_TRABAJADOSPER_LEGAJO: TStringField;
    qPERIODOS_TRABAJADOSEVE_ID_E: TFloatField;
    qPERIODOS_TRABAJADOSEVE_ID_S: TFloatField;
    qPERIODOS_TRABAJADOSPER_HORAS: TFloatField;
    qPERIODOS_TRABAJADOSLEG_NOMBRE: TStringField;
    cxGridPopupMenu1: TcxGridPopupMenu;
    procedure ImportarEventos;
    procedure tbtSalirClick(Sender: TObject);
    procedure tbtImportarClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure tbtRangoClick(Sender: TObject);
    procedure cdbRangoFechasOKClick(Sender: TObject);
    procedure qPERIODOS_TRABAJADOS_AjustarPeriodo(Sender: TField);

  private
    { Private declarations }
    sActualPhase:  string;
    procedure FormManage(sPhase: string);
    procedure ExportarPeriodos ;
  public
    { Public declarations }
  end;

var
  frmExtReyWin: TfrmExtReyWin;

implementation

{$R *.dfm}


procedure TfrmExtReyWin.FormShow(Sender: TObject);
begin
  FormManage('SHOWGRID');
end;

procedure TfrmExtReyWin.ImportarEventos;
begin
  try
    begin
      Screen.Cursor:=crHourGlass;
      qSPImportar.Prepare;
      qSPImportar.ExecProc;
      cnn.Commit;
    end;
  finally
   Screen.Cursor := crDefault;
  end;

end;


procedure TfrmExtReyWin.tbtSalirClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmExtReyWin.tbtImportarClick(Sender: TObject);
begin
 // if sActualPhase = 'SHOW'     then    ImportarEventos;

  if sActualPhase = 'SHOWGRID' then    ExportarPeriodos;


end;


procedure TfrmExtReyWin.tbtRangoClick(Sender: TObject);
begin
  FormManage('SHOW');
end;

procedure TfrmExtReyWin.cdbRangoFechasOKClick(Sender: TObject);
begin
  qPERIODOS_TRABAJADOS.Close;
  qPERIODOS_TRABAJADOS.ParamByName('P_DESDE').AsString := cxFechaDesde.Text;
  qPERIODOS_TRABAJADOS.ParamByName('P_HASTA').AsString := cxFechaHasta.Text;
  qPERIODOS_TRABAJADOS.Open;
  FormManage('SHOWGRID');
end;

procedure TfrmExtReyWin.FormManage(sPhase: string);
begin
  sActualPhase := sPhase;

  if sPhase = 'SHOW' then
  begin
     tsSeleccion.TabVisible := true;
     tsGrilla.TabVisible := false;
     pc1.ActivePage := tsSeleccion;

     tbtImportar.Caption := 'Importar';
     tbtImportar.ImageIndex := 1;
     tbtRango.Visible    := false;
     cxFechaDesde.SetFocus;
  end;

   if sPhase = 'SHOWGRID' then
  begin
     tsSeleccion.TabVisible := false ;
     tsGrilla.TabVisible := true;
     pc1.ActivePage := tsGrilla;
     tbtImportar.Caption := 'Exportar';
     tbtImportar.ImageIndex := 2;
     tbtRango.Visible    := true;

  end;


end;


procedure TfrmExtReyWin.ExportarPeriodos ;
var

sDesde,sHasta : String;
begin
  sDesde :=  cxFechaDesde.Text;
  sHasta :=  cxFechaHasta.Text;

  fsDialog.FileName := 'Periodo - ' + sDesde.Replace('/','-') + ' a ' + sHasta.Replace('/','-')  ;

  if fsDialog.Execute then
    begin
      ExportGridToExcel( fsDialog.FileName, cxGrid1);
    end;

end;


procedure TfrmExtReyWin.qPERIODOS_TRABAJADOS_AjustarPeriodo(Sender: TField);
begin
  qPERIODOS_TRABAJADOS.Edit;
  qPERIODOS_TRABAJADOS.FieldByName('PER_HORAS').AsFloat := (qPERIODOS_TRABAJADOS.FieldByName('PER_HASTA').AsDateTime - qPERIODOS_TRABAJADOS.FieldByName('PER_DESDE').AsDateTime)*24;
  qPERIODOS_TRABAJADOS.Post;
  cnn.Commit;
  qPERIODOS_TRABAJADOS.RefreshRecord();
end;


end.


