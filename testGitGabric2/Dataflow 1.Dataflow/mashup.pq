[StagingDefinition = [Kind = "FastCopy"]]
section Section1;
shared #"Sample file" = let
  Source = SharePoint.Files("https://datatrainingteam.sharepoint.com/sites/HK", [ApiVersion = 15]),
  #"Filtered hidden files" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),
  Navigation = #"Filtered hidden files"{0}[Content]
in
  Navigation;
shared Parameter = let
  Parameter = #"Sample file" meta [IsParameterQuery = true, IsParameterQueryRequired = false, Type = type binary, BinaryIdentifier = #"Sample file"]
in
  Parameter;
shared #"Transform Sample file" = let
  Source = Excel.Workbook(Parameter, null, true),
  Navigation = Source{[Item = "Items", Kind = "Sheet"]}[Data]
in
  Navigation;
[FunctionQueryBinding = "{""exemplarFormulaName"":""Transform Sample file""}"]
shared #"Transform file" = (Parameter as binary) => let
  Source = Excel.Workbook(Parameter, null, true),
  Navigation = Source{[Item = "Items", Kind = "Sheet"]}[Data]
in
  Navigation;
[DataDestinations = {[Definition = [Kind = "Reference", QueryName = "Query_DataDestination", IsNewTarget = true], Settings = [Kind = "Automatic", TypeSettings = [Kind = "Table"]]]}]
shared Query = let
  Source = SharePoint.Files("https://datatrainingteam.sharepoint.com/sites/HK", [ApiVersion = 15]),
  #"Filtered hidden files" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),
  #"Invoke custom function" = Table.AddColumn(#"Filtered hidden files", "Transform file (2)", each #"Transform file (2)"([Content])),
  #"Renamed columns" = Table.RenameColumns(#"Invoke custom function", {{"Name", "Source.Name"}}),
  #"Removed other columns" = Table.SelectColumns(#"Renamed columns", {"Source.Name", "Transform file (2)"}),
  #"Expanded table column" = Table.ExpandTableColumn(#"Removed other columns", "Transform file (2)", Table.ColumnNames(#"Transform file (2)"(#"Sample file (2)"))),
  #"Changed column type" = Table.TransformColumnTypes(#"Expanded table column", {{"Column1", type text}})
in
  #"Changed column type";
shared #"Sample file (2)" = let
  Source = SharePoint.Files("https://datatrainingteam.sharepoint.com/sites/HK", [ApiVersion = 15]),
  #"Filtered hidden files" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),
  Navigation = #"Filtered hidden files"{0}[Content]
in
  Navigation;
shared #"Parameter (2)" = let
  #"Parameter (2)" = #"Sample file (2)" meta [IsParameterQuery = true, IsParameterQueryRequired = false, Type = type binary, BinaryIdentifier = #"Sample file (2)"]
in
  #"Parameter (2)";
shared #"Transform Sample file (2)" = let
  Source = Excel.Workbook(#"Parameter (2)", null, true),
  Navigation = Source{[Item = "Sheet1", Kind = "Sheet"]}[Data]
in
  Navigation;
[FunctionQueryBinding = "{""exemplarFormulaName"":""Transform Sample file (2)""}"]
shared #"Transform file (2)" = (#"Parameter (2)" as binary) => let
  Source = Excel.Workbook(#"Parameter (2)", null, true),
  Navigation = Source{[Item = "Sheet1", Kind = "Sheet"]}[Data]
in
  Navigation;
shared Query_DataDestination = let
  Pattern = Lakehouse.Contents([CreateNavigationProperties = false, EnableFolding = false, HierarchicalNavigation = null]),
  Navigation_1 = Pattern{[workspaceId = "4ea41ae6-8f37-4604-82c9-5c3e90ce43f0"]}[Data],
  Navigation_2 = Navigation_1{[lakehouseId = "1c96ab69-fe40-40b9-8170-eba1e8a7ecbe"]}[Data],
  TableNavigation = Navigation_2{[Id = "Query", ItemKind = "Table"]}?[Data]?
in
  TableNavigation;
