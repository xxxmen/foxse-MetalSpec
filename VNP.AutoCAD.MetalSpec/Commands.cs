using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using Autodesk.AutoCAD.Colors;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System.IO;

namespace MetalSpec
{
    public class Commands
    {
        //[CommandMethod("S2T")]
        //static public void UpdateTableFromSpreadsheet()
        //{
        //    var doc =
        //      Application.DocumentManager.MdiActiveDocument;
        //    var db = doc.Database;
        //    var ed = doc.Editor;

        //    var opt = new PromptEntityOptions("\nВыберите таблицу для обновления");
        //    opt.SetRejectMessage("\nВыбранный объект не таблица.");
        //    opt.AddAllowedClass(typeof(Table), false);

        //    var per = ed.GetEntity(opt);
        //    if (per.Status != PromptStatus.OK)
        //        return;

        //    using (var tr = db.TransactionManager.StartTransaction())
        //    {
        //        try
        //        {
        //            var obj = tr.GetObject(per.ObjectId, OpenMode.ForRead);
        //            var tb = obj as Table;

        //            // It should always be a table
        //            // but we'll check, just in case

        //            if (tb != null)
        //            {
        //                // The table must be open for write

        //                tb.UpgradeOpen();

        //                // Update data link from the spreadsheet

        //                var dlIds = tb.Cells.GetDataLink();

        //                foreach (ObjectId dlId in dlIds)
        //                {
        //                    var dl =
        //                      (DataLink)tr.GetObject(dlId, OpenMode.ForWrite);
        //                    dl.Update(
        //                      UpdateDirection.SourceToData,
        //                      UpdateOption.None
        //                    );

        //                    // And the table from the data link

        //                    tb.UpdateDataLink(
        //                      UpdateDirection.SourceToData,
        //                      UpdateOption.None
        //                    );
        //                }
        //            }
        //            tr.Commit();
        //            ed.WriteMessage(
        //              "\nОбновление таблицы завершено."
        //            );
        //        }
        //        catch (Autodesk.AutoCAD.Runtime.Exception ex)
        //        {
        //            ed.WriteMessage(
        //              "\nОшибка: {0}",
        //              ex.Message
        //            );
        //        }
        //    }
        //}

        //[CommandMethod("T2S")]
        //static public void UpdateSpreadsheetFromTable()
        //{
        //    var doc =
        //      Application.DocumentManager.MdiActiveDocument;
        //    var db = doc.Database;
        //    var ed = doc.Editor;

        //    var opt =
        //      new PromptEntityOptions(
        //        "\nВыберите таблицу с листом для обновления"
        //      );
        //    opt.SetRejectMessage(
        //      "\nВыбранный объект не таблица."
        //    );
        //    opt.AddAllowedClass(typeof(Table), false);

        //    var per = ed.GetEntity(opt);
        //    if (per.Status != PromptStatus.OK)
        //        return;

        //    Transaction tr =
        //      db.TransactionManager.StartTransaction();
        //    using (tr)
        //    {
        //        try
        //        {
        //            DBObject obj =
        //              tr.GetObject(per.ObjectId, OpenMode.ForRead);
        //            Table tb = obj as Table;

        //            // It should always be a table
        //            // but we'll check, just in case

        //            if (tb != null)
        //            {
        //                // The table must be open for write

        //                tb.UpgradeOpen();

        //                // Update the data link from the table

        //                tb.UpdateDataLink(
        //                  UpdateDirection.DataToSource,
        //                  UpdateOption.ForceFullSourceUpdate
        //                );

        //                // And the spreadsheet from the data link

        //                var dlIds = tb.Cells.GetDataLink();
        //                foreach (ObjectId dlId in dlIds)
        //                {
        //                    var dl =
        //                      (DataLink)tr.GetObject(dlId, OpenMode.ForWrite);
        //                    dl.Update(
        //                      UpdateDirection.DataToSource,
        //                      UpdateOption.ForceFullSourceUpdate
        //                    );
        //                }
        //            }
        //            tr.Commit();

        //            ed.WriteMessage(
        //              "\nДокумент обновлён."
        //            );
        //        }
        //        catch (Autodesk.AutoCAD.Runtime.Exception ex)
        //        {
        //            ed.WriteMessage("\nОшибка: {0}", ex.Message);
        //        }
        //    }
        //}

        static public List<string> GetSheetNames(string excelFileName)
        {
            var listSheets = new List<string>();

            var excel = new Excel.Application();
            var wbs = excel.Workbooks.Open(excelFileName);
            foreach (Excel.Worksheet sheet in wbs.Worksheets)
            {
                listSheets.Add(sheet.Name);
            }
            excel.Quit();

            return listSheets;
        }

        //[CommandMethod("TFS")]
        //static public void TableFromSpreadsheet()
        //{
        //    const string dlName = "Импорт данных из таблицы Excell";

        //    var doc =
        //      Application.DocumentManager.MdiActiveDocument;
        //    var db = doc.Database;
        //    var ed = doc.Editor;

        //    // Ask the user to select an XLS(X) file

        //    var ofd =
        //      new OpenFileDialog(
        //        "Выберите документ для связи",
        //        null,
        //        "xls; xlsx; xlsm",
        //        "Файл Excel для связи",
        //        OpenFileDialog.OpenFileDialogFlags.
        //          DoNotTransferRemoteFiles
        //      );

        //    var dr = ofd.ShowDialog();

        //    if (dr != System.Windows.Forms.DialogResult.OK)
        //        return;

        //    // Display the name of the file and the contained sheets

        //    ed.WriteMessage(
        //      "\nВыбранный файл \"{0}\" содержит листы:",
        //      ofd.Filename
        //    );

        //    // First we get the sheet names

        //    var sheetNames = GetSheetNames(ofd.Filename);

        //    if (sheetNames.Count == 0)
        //    {
        //        ed.WriteMessage(
        //          "\nКнига не содержит листов."
        //        );
        //        return;
        //    }

        //    // And loop through, printing their names

        //    for (int i = 0; i < sheetNames.Count; i++)
        //    {
        //        var name = sheetNames[i];

        //        ed.WriteMessage("\n{0} - {1}", i + 1, name);
        //    }

        //    // Ask the user to select one

        //    var pio = new PromptIntegerOptions("\nВыберите лист");
        //    pio.AllowNegative = false;
        //    pio.AllowZero = false;
        //    pio.DefaultValue = 1;
        //    pio.UseDefaultValue = true;
        //    pio.LowerLimit = 1;
        //    pio.UpperLimit = sheetNames.Count;

        //    var pir = ed.GetInteger(pio);
        //    if (pir.Status != PromptStatus.OK)
        //        return;

        //    // Ask for the insertion point of the table

        //    var ppr = ed.GetPoint("\nУкажите точку вставки таблицы");
        //    if (ppr.Status != PromptStatus.OK)
        //        return;

        //    // Remove any Data Link, if one exists already

        //    var dlm = db.DataLinkManager;
        //    var dlId = dlm.GetDataLink(dlName);
        //    if (dlId != ObjectId.Null)
        //    {
        //        dlm.RemoveDataLink(dlId);
        //    }

        //    // Create and add the new Data Link, this time with
        //    // a direction connection to the selected sheet

        //    var dl = new DataLink();
        //    dl.DataAdapterId = "AcExcel";
        //    dl.Name = dlName;
        //    dl.Description = "Источник данных для спецификации металлопроката.";
        //    dl.ConnectionString =
        //      ofd.Filename + "!" + sheetNames[pir.Value - 1];
        //    dl.DataLinkOption =
        //      DataLinkOption.PersistCache;
        //    dl.UpdateOption |=
        //      (int)UpdateOption.AllowSourceUpdate;

        //    dlId = dlm.AddDataLink(dl);

        //    using (var tr = doc.TransactionManager.StartTransaction())
        //    {
        //        tr.AddNewlyCreatedDBObject(dl, true);

        //        var bt =
        //          (BlockTable)tr.GetObject(
        //            db.BlockTableId,
        //            OpenMode.ForRead
        //          );

        //        // Create our table

        //        var tb = new Table();
        //        tb.TableStyle = db.Tablestyle;
        //        tb.Position = ppr.Value;
        //        tb.Cells.SetDataLink(dlId, true);
        //        tb.GenerateLayout();
        //        double w = tb.Width;
        //        double h = tb.Height;
        //        tb.Width = h;
        //        tb.Height = w;
        //        // Add it to the drawing

        //        var btr =
        //          (BlockTableRecord)tr.GetObject(
        //            db.CurrentSpaceId,
        //            OpenMode.ForWrite
        //          );

        //        btr.AppendEntity(tb);
        //        tr.AddNewlyCreatedDBObject(tb, true);
        //        tr.Commit();

        //        ed.Regen();
        //    }
        //}

        //[CommandMethod("TFSR")]
        //static public void TableFromSpreadsheetRange()
        //{
        //    const string dlName = "Import table from Excel demo";

        //    var doc =
        //      Application.DocumentManager.MdiActiveDocument;
        //    var db = doc.Database;
        //    var ed = doc.Editor;

        //    // Ask the user to select an XLS(X) file

        //    var ofd =
        //      new OpenFileDialog(
        //        "Select Excel spreadsheet to link",
        //        null,
        //        "xls; xlsx; xlsm",
        //        "Файл Excel для связи",
        //        OpenFileDialog.OpenFileDialogFlags.
        //          DoNotTransferRemoteFiles
        //      );

        //    var dr = ofd.ShowDialog();

        //    if (dr != System.Windows.Forms.DialogResult.OK)
        //        return;

        //    // Display the name of the file and the contained sheets

        //    ed.WriteMessage(
        //      "\nFile selected was \"{0}\". Contains these sheets:",
        //      ofd.Filename
        //    );

        //    // First we get the sheet names

        //    var sheetNames = GetSheetNames(ofd.Filename);

        //    if (sheetNames.Count == 0)
        //    {
        //        ed.WriteMessage(
        //          "\nWorkbook doesn't contain any sheets."
        //        );
        //        return;
        //    }

        //    // And loop through, printing their names

        //    for (int i = 0; i < sheetNames.Count; i++)
        //    {
        //        var name = sheetNames[i];

        //        ed.WriteMessage("\n{0} - {1}", i + 1, name);
        //    }

        //    // Ask the user to select one

        //    var pio = new PromptIntegerOptions("\nSelect a sheet");
        //    pio.AllowNegative = false;
        //    pio.AllowZero = false;
        //    pio.DefaultValue = 1;
        //    pio.UseDefaultValue = true;
        //    pio.LowerLimit = 1;
        //    pio.UpperLimit = sheetNames.Count;

        //    var pir = ed.GetInteger(pio);
        //    if (pir.Status != PromptStatus.OK)
        //        return;

        //    // Ask the user to select a range of cells in the spreadsheet

        //    // We'll use a Regular Expression that matches a column (with
        //    // one or more letters) followed by a numeric row (which we're
        //    // naming "row1" so we can validate it's > 0 later),
        //    // followed by a colon and then the same (but with "row2")

        //    const string rangeExp =
        //      "^[A-Z]+(?<row1>[0-9]+):[A-Z]+(?<row2>[0-9]+)$";
        //    bool done = false;
        //    string range = "";

        //    do
        //    {
        //        var psr = ed.GetString("\nEnter cell range <entire sheet>");
        //        if (psr.Status != PromptStatus.OK)
        //            return;

        //        if (String.IsNullOrEmpty(psr.StringResult))
        //        {
        //            // Default is to select entire sheet

        //            done = true;
        //        }
        //        else
        //        {
        //            // If a string was entered, make sure it's a
        //            // valid cell range, which means it matches the
        //            // Regular Expression and has positive (non-zero)
        //            // row numbers

        //            var m =
        //              Regex.Match(
        //                psr.StringResult, rangeExp, RegexOptions.IgnoreCase
        //              );
        //            if (
        //              m.Success &&
        //              Int32.Parse(m.Groups["row1"].Value) > 0 &&
        //              Int32.Parse(m.Groups["row2"].Value) > 0
        //            )
        //            {
        //                done = true;
        //                range = psr.StringResult.ToUpper();
        //            }
        //            else
        //            {
        //                ed.WriteMessage("\nInvalid range, please try again.");
        //            }
        //        }
        //    } while (!done);

        //    // Ask for the insertion point of the table

        //    var ppr = ed.GetPoint("\nEnter table insertion point");
        //    if (ppr.Status != PromptStatus.OK)
        //        return;

        //    try
        //    {
        //        // Remove any Data Link, if one exists already

        //        var dlm = db.DataLinkManager;
        //        var dlId = dlm.GetDataLink(dlName);
        //        if (dlId != ObjectId.Null)
        //        {
        //            dlm.RemoveDataLink(dlId);
        //        }

        //        // Create and add the new Data Link, this time with
        //        // a direction connection to the selected sheet

        //        var dl = new DataLink();
        //        dl.DataAdapterId = "AcExcel";
        //        dl.Name = dlName;
        //        dl.Description = "Excel fun with Through the Interface";
        //        dl.ConnectionString =
        //          ofd.Filename +
        //          "!" + sheetNames[pir.Value - 1] +
        //          (String.IsNullOrEmpty(range) ? "" : "!" + range);
        //        dl.DataLinkOption = DataLinkOption.PersistCache;
        //        dl.UpdateOption |= (int)UpdateOption.AllowSourceUpdate;

        //        dlId = dlm.AddDataLink(dl);

        //        using (var tr = doc.TransactionManager.StartTransaction())
        //        {
        //            tr.AddNewlyCreatedDBObject(dl, true);

        //            var bt =
        //              (BlockTable)tr.GetObject(
        //                db.BlockTableId,
        //                OpenMode.ForRead
        //              );

        //            // Create our table

        //            var tb = new Table();
        //            tb.TableStyle = db.Tablestyle;
        //            tb.Position = ppr.Value;
        //            tb.Cells.SetDataLink(dlId, true);
        //            tb.GenerateLayout();

        //            double w = tb.Width;
        //            double h = tb.Height;
        //            tb.Width = h;
        //            tb.Height = w;
        //            // Add it to the drawing

        //            var btr =
        //              (BlockTableRecord)tr.GetObject(
        //                db.CurrentSpaceId,
        //                OpenMode.ForWrite
        //              );

        //            btr.AppendEntity(tb);
        //            tr.AddNewlyCreatedDBObject(tb, true);
        //            tr.Commit();
        //        }
        //    }
        //    catch (Autodesk.AutoCAD.Runtime.Exception ex)
        //    {
        //        ed.WriteMessage("\nException: {0}", ex.Message);
        //    }
        //}

        //[CommandMethod("STF")]
        //public static void StripTableFormatting()
        //{
        //    Document doc =
        //      Application.DocumentManager.MdiActiveDocument;
        //    Editor ed = doc.Editor;

        //    // Ask the user to select a table

        //    var peo = new PromptEntityOptions("\nSelect table");
        //    peo.SetRejectMessage("\nMust be a table.");
        //    peo.AddAllowedClass(typeof(Table), false);
        //    peo.AllowNone = false;
        //    var per = ed.GetEntity(peo);
        //    if (per.Status != PromptStatus.OK)
        //        return;

        //    var tabId = per.ObjectId;

        //    using (var tr = doc.TransactionManager.StartTransaction())
        //    {
        //        // Open the table for write

        //        var tab = tr.GetObject(tabId, OpenMode.ForWrite) as Table;
        //        if (tab == null)
        //            return;

        //        // Start by clearing any style overrides

        //        tab.Cells.ClearStyleOverrides();

        //        // We'll use an MText object to strip embedded formatting

        //        using (var mt = new MText())
        //        {
        //            for (int r = 0; r < tab.Rows.Count; r++)
        //            {
        //                for (int c = 0; c < tab.Columns.Count; c++)
        //                {
        //                    // Get the cell and its contents

        //                    var cell = tab.Cells[r, c];
        //                    mt.Contents = cell.TextString;

        //                    // Explode the text fragments

        //                    mt.ExplodeFragments(
        //                      (a, b) =>
        //                      {
        //                          if (a != null)
        //                          {
        //                              // Assuming we have a fragment, use its text
        //                              // to set the cell contents

        //                              cell.TextString = a.Text;
        //                          }
        //                          return MTextFragmentCallbackStatus.Continue;
        //                      }
        //                    );
        //                }
        //            }
        //        }
        //        tr.Commit();
        //    }
        //}

        [CommandMethod("TFMS")]
        static public void TableFromMetalSpec()
        {
            const string dlName = "MetalSpec";

            var json = File.ReadAllText(Environment.GetEnvironmentVariable("TEMP") + @"\metalSpec_sd.json");

            var tempSheetData =
                JsonConvert.DeserializeObject<SortedDictionary<string, int>>(json);

            var fileName = tempSheetData.Where(s => s.Key.Contains("_SPEC_FILE_PATH_")).First().Key.Replace("_SPEC_FILE_PATH_", "");

            var sheetData = tempSheetData.Where(s => !s.Key.Contains("_SPEC_FILE_PATH_")).ToList();

            var doc =
              Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            // Ask the user to select an XLS(X) file

            //var ofd =
            //  new OpenFileDialog(
            //    "Select Excel spreadsheet to link",
            //    null,
            //    "xls; xlsx; xlsm",
            //    "Файл Excel для связи",
            //    OpenFileDialog.OpenFileDialogFlags.
            //      DoNotTransferRemoteFiles
            //  );

            //var dr = ofd.ShowDialog();

            //if (dr != System.Windows.Forms.DialogResult.OK)
            //    return;

            if (!File.Exists(fileName))
                return;

            // Display the name of the file and the contained sheets

            ed.WriteMessage(
              "\nВыбранный файл \"{0}\" содержал листы:",
              //ofd.Filename
              fileName
            );

            // First we get the sheet names

            var sheetNames = GetSheetNames(fileName/*ofd.Filename*/);

            if (sheetNames.Count == 0)
            {
                ed.WriteMessage(
                  "\nВ книге нет листов."
                );
                return;
            }

            // And loop through, printing their names

            for (int i = 0; i < sheetNames.Count; i++)
            {
                var name = sheetNames[i];

                ed.WriteMessage("\n{0} - {1}", i + 1, name);
            }

            var ppr = ed.GetPoint("\nУкажите точку вставки таблицы");
            if (ppr.Status != PromptStatus.OK)
                return;
            string firstCell = "A1";
            int lim = 4;
            try
            {
                // Remove any Data Link, if one exists already
                ObjectId newTableId;
                var dlm = db.DataLinkManager;

                Autodesk.AutoCAD.Geometry.Point3d pos = ppr.Value;

                // Create and add the new Data Link, this time with
                // a direction connection to the selected sheet
                foreach (var sheet in sheetData)
                {
                    var dlId = dlm.GetDataLink(dlName + " " + sheet.Key);
                    if (dlId != ObjectId.Null)
                    {
                        dlm.RemoveDataLink(dlId);
                    }

                    var dl = new DataLink();
                    dl.DataAdapterId = "AcExcel";
                    dl.Name = dlName + " " + sheet.Key;
                    dl.Description = "";
                    dl.ConnectionString =
                      fileName +
                      "!" + sheet.Key +
                      ((sheet.Value == 0) ? "" : "!"+ firstCell + ":" + "S" + sheet.Value);
                    dl.DataLinkOption = DataLinkOption.PersistCache;
                    dl.UpdateOption |= (int)UpdateOption.AllowSourceUpdate;

                    dlId = dlm.AddDataLink(dl);

                    using (var tr = doc.TransactionManager.StartTransaction())
                    {
                        tr.AddNewlyCreatedDBObject(dl, true);

                        var bt =
                          (BlockTable)tr.GetObject(
                            db.BlockTableId,
                            OpenMode.ForRead
                          );

                        DBDictionary sd =
                           (DBDictionary)tr.GetObject(
                             db.TableStyleDictionaryId,
                             OpenMode.ForRead
                           );

                        // Use the style if it already exists
                        ObjectId tsId = ObjectId.Null;

                        // Create our table

                        var tb = new Table();

                        if (sd.Contains("ROM35"))
                        {
                            tsId = sd.GetAt("ROM35");
                            tb.TableStyle = tsId;
                        }
                        else
                        {
                            TextStyleTable newTextStyleTable = tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;

                            if (!newTextStyleTable.Has("ROM35"))  //The TextStyle is currently not in the database
                            {
                                newTextStyleTable.UpgradeOpen();
                                var newTextStyleTableRecord = new TextStyleTableRecord();
                                newTextStyleTableRecord.FileName = "romans.shx";
                                newTextStyleTableRecord.Name = "ROM35";
                                newTextStyleTableRecord.XScale = 0.8;
                                newTextStyleTableRecord.TextSize = 3.5;
                                //Autodesk.AutoCAD.GraphicsInterface.FontDescriptor myNewTextStyle = new Autodesk.AutoCAD.GraphicsInterface.FontDescriptor("ROMANS", false, false, 0, 0);
                                //newTextStyleTableRecord.Font = myNewTextStyle;
                                newTextStyleTable.Add(newTextStyleTableRecord);
                                tr.AddNewlyCreatedDBObject(newTextStyleTableRecord, true);
                            }

                            // Otherwise we have to create it
                            TableStyle ts = new TableStyle();
                            #region Тут всякие цвета - шрифты


                            // With yellow text everywhere (yeuch :-)

                            ts.SetColor(
                              Color.FromColorIndex(ColorMethod.ByAci, 2),
                              (int)(RowType.TitleRow |
                                    RowType.HeaderRow |
                                    RowType.DataRow)
                            );

                            // And now with magenta outer grid-lines

                            ts.SetGridColor(
                              Color.FromColorIndex(ColorMethod.ByAci, 6),
                              (int)GridLineType.OuterGridLines,
                              (int)(RowType.TitleRow |
                                    RowType.HeaderRow |
                                    RowType.DataRow)
                            );

                            // And red inner grid-lines

                            ts.SetGridColor(
                              Color.FromColorIndex(ColorMethod.ByAci, 1),
                              (int)GridLineType.InnerGridLines,
                              (int)(RowType.TitleRow |
                                    RowType.HeaderRow |
                                    RowType.DataRow)
                            );

                            if (newTextStyleTable.Has("ROM35"))
                            // And we'll make the grid-lines nice and chunky
                            {
                                ts.SetTextStyle(newTextStyleTable["ROM35"], (int)RowType.TitleRow); // title row 
                                ts.SetTextStyle(newTextStyleTable["ROM35"], (int)RowType.HeaderRow); // header row 
                                ts.SetTextStyle(newTextStyleTable["ROM35"], (int)RowType.DataRow); // data row 
                            }
                            // Add our table style to the dictionary
                            //  and to the transaction
                            #endregion 
                            tsId = ts.PostTableStyleToDatabase(db, "ROM35");
                            tr.AddNewlyCreatedDBObject(ts, true);
                        }

                        tb.Position = pos;
                        tb.Cells.SetDataLink(dlId, true);
                        tb.GenerateLayout();

                        double w = tb.Width;
                        double h = tb.Height;
                        tb.Width = 340;
                        // Table Height
                        if (lim == 4)
                            tb.Height = 8 + 22 + 4 + 8 * sheet.Value;
                        else
                            tb.Height = 14 + 4 + 8 * sheet.Value;
                        // Add it to the drawing

                        /* */

                        var btr =
                          (BlockTableRecord)tr.GetObject(
                            db.CurrentSpaceId,
                            OpenMode.ForWrite
                          );

                        btr.AppendEntity(tb);
                        tr.AddNewlyCreatedDBObject(tb, true);
                        newTableId = tb.ObjectId;
                        tb.Cells.ClearStyleOverrides();
                        tb.TableStyle = sd.GetAt("ROM35");
                        tb.RemoveDataLink();
                        tb.UpgradeOpen();

                        for (int c = 0; c < tb.Columns.Count; c++)
                        {
                            if (c < 3)
                            {
                                tb.Columns[c].Width = 30;
                            }
                            else if (c == 3)
                            {
                                tb.Columns[c].Width = 10;
                            }
                            else if (c != tb.Columns.Count - 1)
                            {
                                tb.Columns[c].Width = 15;
                            }
                            else
                            {
                                tb.Columns[c].Width = 25;
                            }
                        }

                        using (var mt = new MText())
                        {
                            for (int r = 0; r < tb.Rows.Count; r++)
                            {
                                tb.Rows[r].Height = 8;
                                for (int c = 0; c < tb.Columns.Count; c++)
                                {
                                    // Get the cell and its contents
                                    var cell = tb.Cells[r, c];
                                    mt.Contents = cell.TextString;
                                    mt.Height = 3.5;
                                    cell.ContentColor = Color.FromColorIndex(ColorMethod.ByColor, 2);
                                    cell.TextHeight = 3.5;
                                    cell.Alignment = CellAlignment.MiddleCenter;
                                    if (r < lim )
                                    {
                                        cell.Borders.Bottom.Color = Color.FromColorIndex(ColorMethod.ByColor, 6);
                                        cell.Borders.Top.Color = Color.FromColorIndex(ColorMethod.ByColor, 6);
                                    }
                                    else
                                    {
                                        if (tb.Cells[cell.Row, 1].TextString.Contains("Итого"))
                                        {
                                            cell.Borders.Bottom.Color = Color.FromColorIndex(ColorMethod.ByColor, 2);
                                            cell.Borders.Top.Color = Color.FromColorIndex(ColorMethod.ByColor, 2);
                                            if (c > 3)
                                                cell.ContentColor = Color.FromColorIndex(ColorMethod.ByColor, 3);
                                        }
                                        else if (tb.Cells[cell.Row, 0].TextString.Contains("Всего масса металла") ||
                                            tb.Cells[cell.Row, 0].TextString.Contains("B том числе по маркам"))
                                        {
                                            cell.Borders.Bottom.Color = Color.FromColorIndex(ColorMethod.ByColor, 6);
                                            cell.Borders.Top.Color = Color.FromColorIndex(ColorMethod.ByColor, 6);
                                            if (c > 3)
                                                cell.ContentColor = Color.FromColorIndex(ColorMethod.ByColor, 6);
                                        }
                                        else if (tb.Cells[cell.Row, 0].TextString.Contains("Всего профиля"))
                                        {
                                            cell.Borders.Bottom.Color = Color.FromColorIndex(ColorMethod.ByColor, 6);
                                            cell.Borders.Top.Color = Color.FromColorIndex(ColorMethod.ByColor, 6);
                                            if (c > 3)
                                                cell.ContentColor = Color.FromColorIndex(ColorMethod.ByColor, 5);
                                        }
                                        else if (tb.Cells[cell.Row - 1, 0].TextString.Contains("Всего профиля"))
                                        {
                                            cell.Borders.Bottom.Color = Color.FromColorIndex(ColorMethod.ByColor, 1);
                                            cell.Borders.Top.Color = Color.FromColorIndex(ColorMethod.ByColor, 6);
                                        }
                                        else if (tb.Cells[cell.Row - 1, 1].TextString.Contains("Итого"))
                                        {
                                            cell.Borders.Bottom.Color = Color.FromColorIndex(ColorMethod.ByColor, 2);
                                            cell.Borders.Top.Color = Color.FromColorIndex(ColorMethod.ByColor, 2);
                                            if (c > 3)
                                                cell.ContentColor = Color.FromColorIndex(ColorMethod.ByColor, 3);
                                        }
                                        else
                                        {
                                            if (tb.Cells[cell.Row - 1, 1].TextString.Contains("Итого"))
                                            {
                                                cell.Borders.Top.Color = Color.FromColorIndex(ColorMethod.ByColor, 2);
                                            }
                                            else
                                            {
                                                cell.Borders.Top.Color = Color.FromColorIndex(ColorMethod.ByColor, 1);
                                            }
                                            cell.Borders.Bottom.Color = Color.FromColorIndex(ColorMethod.ByColor, 1);
                                            
                                        }
                                    }
                                    
                                    cell.Borders.Left.Color = Color.FromColorIndex(ColorMethod.ByColor, 6);
                                    cell.Borders.Right.Color = Color.FromColorIndex(ColorMethod.ByColor, 6);
                                    //tb.Cells[r, c].Style = "_HEADER";
                                    //cell.Style = "_HEADER";

                                    mt.Contents = @"{\W0.75;" + mt.Text + "}";
                                    mt.Width = tb.Columns[c].Width;
                                    // Explode the text fragments
                                    cell.TextString = mt.Contents;
                                    
                                }
                            }
                        }
                        if (firstCell == "A1")
                            tb.Rows[1].Height = 55.0;
                        pos = new Autodesk.AutoCAD.Geometry.Point3d(new double[] { pos.X + tb.Width + 20, pos.Y, pos.Z });
                        tb.DowngradeOpen();
                        tr.Commit();
                    }
                    firstCell = "A3";
                    lim = 2;
                }

            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                ed.WriteMessage("\nИсключение: {0}", ex.Message);
            }
        }
    }
}
