using Qlik.Engine;

try {
  var uri = new Uri("https://xxxx.yy.qlikcloud.com");
  var location = QcsLocation.FromUri(uri);
  location.AsApiKey("eyJhbGci....");
  using IApp app = location.App("72a3da4b-1093-4c4c-840d-1ee44fbcbb91", SessionToken.Unique(), false);

  app.ClearAll();
  var field = app.GetField("支店名");
  var select_ok = field.SelectValues([new(){Text = "関東支店"}, new(){Text = "関西支店"}]);

  var listobject_def = new ListObjectDef{
    Def = new NxInlineDimensionDef{
      FieldDefs = ["支店名"],
      FieldLabels = ["支店名"],
      SortCriterias = [new(){SortByLoadOrder = SortDirection.Ascending}]
    },
    FrequencyMode = NxFrequencyMode.NX_FREQUENCY_VALUE,
    ShowAlternatives = true
  };
  var lo_gp = new GenericObjectProperties{
    Info = new NxInfo{Type = "my-list-object"}
  };
  lo_gp.Set<ListObjectDef>("qListObjectDef", listobject_def);
  var lo_hypercube = app.CreateGenericSessionObject(lo_gp);
  var lo_pager = lo_hypercube.GetListObjectPager("/qListObjectDef");
  int lo_row_len = lo_pager.NumberOfRows;
  int lo_col_len = lo_pager.NumberOfColumns;
  int lo_page_height = 1;
  if(lo_col_len != 0)
    lo_page_height = (int)Math.Floor(10000/(double)lo_col_len);
  int lo_row_top = 0;
  ListObject lo_layout = lo_hypercube.Layout.Get<ListObject>("qListObject");
  System.Console.Out.WriteLine(lo_layout.DimensionInfo.FallbackTitle);
  while(true) {
    IEnumerable<NxCellRows> rows = lo_pager.GetAllData(
      new NxPage{Top = lo_row_top, Left = 0, Height = lo_page_height, Width = lo_col_len}
    );
    foreach(var row in rows) {
      foreach(NxCell cell in row) {
        string field_data = cell.ElemNumber + ",";
        if(cell.State == StateEnumType.SELECTED)
          field_data += "(Selected)";
        if(cell.ElemNumber == -2) // -2: the cell is a Null cell.
          field_data += "-";
        else if(!Double.IsNaN(cell.Num))
          field_data += cell.Num;
        else if(cell.Text.Length > 0)
          field_data += cell.Text;
        else
          field_data += "";
        lo_row_top++;
        System.Console.Out.WriteLine(field_data);
      }
    }
    if(lo_row_top >= lo_row_len)
      break;
  }
  app.DestroyGenericObject(lo_hypercube.Id);

  var hypercube_def = new HyperCubeDef(){
    Dimensions = [new(){
      Def = new NxInlineDimensionDef(){
        FieldDefs = ["営業員名"],
        FieldLabels = ["営業員名"]
      },
      NullSuppression = true
    }],
    Measures = [new(){
      Def = new NxInlineMeasureDef(){
        Def = "Sum([販売価格])",
        Label = "実績",
        NumFormat = new(){Type = FieldAttrType.MONEY, UseThou = 1, Thou = ","}
      },
      SortBy = new SortCriteria(){
        SortByState      = 0,
        SortByFrequency  = 0,
        SortByNumeric    = SortDirection.Descending, // ソート: 0=無し, 1=昇順, -1=降順
        SortByAscii      = 0,
        SortByLoadOrder  = 0,
        SortByExpression = 0,
        Expression       = ""
      }
    }],
    SuppressZero = false,
    SuppressMissing = false,
    Mode = NxHypercubeMode.DATA_MODE_STRAIGHT,
    InterColumnSortOrder = [1,0], // ソート順: 1=実績, 0=営業員名
    StateName = "$"
  };
  var hc_gp = new GenericObjectProperties {
    Info = new NxInfo{Type = "my-straight-hypercube"}
  };
  hc_gp.Set<HyperCubeDef>("qHyperCubeDef", hypercube_def);
  var hc_hypercube = app.CreateGenericSessionObject(hc_gp);
  var hc_pager = hc_hypercube.GetHyperCubePager("/qHyperCubeDef");
  int hc_row_len = hc_pager.NumberOfRows;
  int hc_col_len = hc_pager.NumberOfColumns;
  int hc_page_height = 1;
  if(hc_col_len != 0)
    hc_page_height = (int)Math.Floor(10000/(double)hc_col_len);
  int hc_row_top = 0;
  HyperCube hc_layaout = hc_hypercube.Layout.Get<HyperCube>("qHyperCube");
  foreach(var dim in hc_layaout.DimensionInfo)
    System.Console.Out.WriteLine(dim.FallbackTitle);
  foreach(var mes in hc_layaout.MeasureInfo)
    System.Console.Out.WriteLine(mes.FallbackTitle);
  while(true) {
    IEnumerable<NxCellRows> rows = hc_pager.GetAllData(
      new NxPage{Top = hc_row_top, Left = 0, Height = hc_page_height, Width = hc_col_len}
    );
    foreach(var row in rows) {
      foreach(NxCell cell in row) {
        string field_data = "";
        if(cell.ElemNumber == -2) // -2: the cell is a Null cell.
          field_data += "-";
        else if(cell.Text.Length > 0)
          field_data += cell.Text;
        else if(!Double.IsNaN(cell.Num))
          field_data += cell.Num;
        else
          field_data += "";
        hc_row_top++;
        System.Console.Out.WriteLine(field_data);
      }
    }
    if(hc_row_top >= hc_row_len)
      break;
  }
  app.DestroyGenericObject(hc_hypercube.Id);
}
catch (System.Exception e) {
  System.Console.Write(e);
}
