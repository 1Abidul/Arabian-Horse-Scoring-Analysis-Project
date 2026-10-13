// Power Query (M) — optional. Paste into Home > Transform data > New Source > Blank Query > Advanced Editor
// to load all CSVs from one folder parameter. Set FolderPath to where you saved powerbi/data.
// Create one query per table by changing the FileName, or use Get Data > Text/CSV for each file.

let
    FolderPath = "C:\HorseProject\powerbi\data\",      // <-- change this
    FileName   = "fact_trait_scores.csv",             // <-- change per table
    Source     = Csv.Document(File.Contents(FolderPath & FileName), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Promoted   = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed      = Table.TransformColumnTypes(Promoted,
                    List.Transform(
                        List.Intersect({Table.ColumnNames(Promoted),
                            {"Score","TotalScore","AvgTotal","JudgeSpread","SireScore","DamSireScore","MidParentScore",
                             "Type","Head_Neck","Body_Topline","Legs","Movement","Balance","ParentShowScore"}}),
                        each {_, type number}) &
                    List.Transform(
                        List.Intersect({Table.ColumnNames(Promoted), {"BirthYear","Age","Rank","SortOrder","MaxScore"}}),
                        each {_, Int64.Type}))
in
    Typed
