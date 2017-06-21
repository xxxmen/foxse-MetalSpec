using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.IO;
using System;
using Newtonsoft.Json;
using MetalSpec.DataAdapter;
using System.Collections.Generic;
using System.Linq;

namespace MetalSpec
{
    class ExcelTable
    {
        internal static ExcelWorksheet CreateWorksheet(ExcelPackage package, int sheetIndex, double firstRowHeight, double secondRowHeight, double thirdRowHeight, double generalRowHeight)
        {
            string gostFontName = "GOST type A";
            string fontName = "Arial";
            using (Font fontTester = new Font(gostFontName, 12, FontStyle.Regular, GraphicsUnit.Pixel))
            {
                if (fontTester.Name == gostFontName)
                {
                    fontName = "GOST type A";
                }
            }

            ExcelWorksheet ws = package.Workbook.Worksheets.Add($"Лист {sheetIndex.ToString()}");
            ws.Cells.Style.Font.Size = 11;
            ws.Cells.Style.Font.Name = fontName;
            ws.PrinterSettings.PaperSize = ePaperSize.A3;
            ws.PrinterSettings.Orientation = eOrientation.Landscape;
            ws.PrinterSettings.LeftMargin = 0.787M;
            ws.PrinterSettings.TopMargin = 0.25M;
            ws.PrinterSettings.RightMargin = 0.05M;
            ws.PrinterSettings.BottomMargin = 0.05M;
            ws.View.PageLayoutView = true;

            const double width10mm = 5.2;
            const double width15mm = 7.7;
            const double width25mm = 12.7;
            const double width30mm = 15.2;
            ws.Column(1).Width = width30mm;
            ws.Column(2).Width = width30mm;
            ws.Column(3).Width = width30mm;
            ws.Column(4).Width = width10mm;
            ws.Column(5).Width = width15mm;
            ws.Column(6).Width = width15mm;
            ws.Column(7).Width = width15mm;
            ws.Column(8).Width = width15mm;
            ws.Column(9).Width = width15mm;
            ws.Column(10).Width = width15mm;
            ws.Column(11).Width = width15mm;
            ws.Column(12).Width = width15mm;
            ws.Column(13).Width = width15mm;
            ws.Column(14).Width = width15mm;
            ws.Column(15).Width = width15mm;
            ws.Column(16).Width = width15mm;
            ws.Column(17).Width = width15mm;
            ws.Column(18).Width = width15mm;
            ws.Column(19).Width = width25mm;
            ws.Column(20).Width = width25mm;
            ws.Column(21).Width = 13.5;

            ws.Row(1).Height = firstRowHeight;
            ws.Row(2).Height = secondRowHeight;
            ws.Row(3).Height = thirdRowHeight;
            ws.Row(4).Height = generalRowHeight;
            ws.Row(5).Height = generalRowHeight;
            ws.Row(6).Height = generalRowHeight;
            ws.Row(7).Height = generalRowHeight;
            ws.Row(8).Height = generalRowHeight;
            ws.Row(9).Height = generalRowHeight;
            ws.Row(10).Height = generalRowHeight;
            ws.Row(11).Height = generalRowHeight;
            ws.Row(12).Height = generalRowHeight;
            ws.Row(13).Height = generalRowHeight;
            ws.Row(14).Height = generalRowHeight;
            ws.Row(15).Height = generalRowHeight;
            ws.Row(16).Height = generalRowHeight;
            ws.Row(17).Height = generalRowHeight;
            ws.Row(18).Height = generalRowHeight;
            ws.Row(19).Height = generalRowHeight;
            ws.Row(20).Height = generalRowHeight;
            ws.Row(21).Height = generalRowHeight;
            ws.Row(22).Height = generalRowHeight;
            ws.Row(23).Height = generalRowHeight;
            ws.Row(24).Height = generalRowHeight;
            ws.Row(25).Height = generalRowHeight;
            ws.Row(26).Height = generalRowHeight;
            ws.Row(27).Height = generalRowHeight;
            ws.Row(28).Height = generalRowHeight;
            ws.Row(29).Height = generalRowHeight;
            ws.Row(30).Height = generalRowHeight;
            ws.Row(31).Height = generalRowHeight;
            ws.Row(32).Height = generalRowHeight;
            ws.Row(33).Height = generalRowHeight;
            ws.Row(34).Height = 30;

            ws.Cells[$"A1:U34"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

            return ws;
        }

        internal static void DrawWorksheetHeader(string units, string[] categories, ExcelWorksheet ws)
        {
            using (ExcelRange h = ws.Cells["A1:S3"])
            {
                h.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                h.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                using (ExcelRange h2 = ws.Cells[1, 1, 2, 19])
                {
                    h2.Style.WrapText = true;

                    ws.Cells[1, 1].Value = "Наименование профиля ГОСТ, ТУ";
                    ws.Cells[1, 1, 2, 1].Merge = true;
                    ws.Cells[1, 2].Value = "Наименование или марка металла ГОСТ, ТУ";
                    ws.Cells[1, 2, 2, 2].Merge = true;
                    ws.Cells[1, 3].Value = "Номер или размеры профиля, мм";
                    ws.Cells[1, 3, 2, 3].Merge = true;
                    ws.Cells[1, 4].Value = "№ п.п.";
                    ws.Cells[1, 4, 2, 4].Merge = true;
                    ws.Cells[1, 5].Value = $"Масса металла по элементам конструкций, {units}";
                    ws.Cells[1, 5, 1, 18].Merge = true;

                    using (ExcelRange h3 = ws.Cells[2, 5, 2, 19])
                    {
                        h3.Style.TextRotation = 90;

                        if (categories != null)
                        {
                            if (categories.Length > 0) ws.Cells[2, 5].Value = categories[0];
                            if (categories.Length > 1) ws.Cells[2, 6].Value = categories[1];
                            if (categories.Length > 2) ws.Cells[2, 7].Value = categories[2];
                            if (categories.Length > 3) ws.Cells[2, 8].Value = categories[3];
                            if (categories.Length > 4) ws.Cells[2, 9].Value = categories[4];
                            if (categories.Length > 5) ws.Cells[2, 10].Value = categories[5];
                            if (categories.Length > 6) ws.Cells[2, 11].Value = categories[6];
                            if (categories.Length > 7) ws.Cells[2, 12].Value = categories[7];
                            if (categories.Length > 8) ws.Cells[2, 13].Value = categories[8];
                            if (categories.Length > 8) ws.Cells[2, 14].Value = categories[9];
                            if (categories.Length > 10) ws.Cells[2, 15].Value = categories[10];
                            if (categories.Length > 11) ws.Cells[2, 16].Value = categories[11];
                            if (categories.Length > 12) ws.Cells[2, 17].Value = categories[12];
                            if (categories.Length > 13) ws.Cells[2, 18].Value = categories[13];
                        }
                        ws.Cells[1, 19].Value = $"Общая масса, {units}";
                        ws.Cells[1, 19, 2, 19].Merge = true;
                        ws.Cells["E1:R1"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        ws.Cells["A3:S3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    }
                }

                ws.Cells[3, 1].Value = "1";
                ws.Cells[3, 2].Value = "2";
                ws.Cells[3, 3].Value = "3";
                ws.Cells[3, 4].Value = "4";
                ws.Cells[3, 5].Value = "5";
                ws.Cells[3, 6].Value = "6";
                ws.Cells[3, 7].Value = "7";
                ws.Cells[3, 8].Value = "8";
                ws.Cells[3, 9].Value = "9";
                ws.Cells[3, 10].Value = "10";
                ws.Cells[3, 11].Value = "11";
                ws.Cells[3, 12].Value = "12";
                ws.Cells[3, 13].Value = "13";
                ws.Cells[3, 14].Value = "14";
                ws.Cells[3, 15].Value = "15";
                ws.Cells[3, 16].Value = "16";
                ws.Cells[3, 17].Value = "17";
                ws.Cells[3, 18].Value = "18";
                ws.Cells[3, 19].Value = "19";

                ws.Cells[$"A1:A3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"B1:B3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"C1:C3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"D1:D3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"E2:E3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"F2:F3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"G2:G3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"H1:H3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"I1:I3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"J1:J3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"K1:K3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"L1:L3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"M1:M3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"N1:N3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"O1:O3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"P1:P3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"Q1:Q3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"R1:R3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"S1:S3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            }
        }

        internal static void DrawStamp(ExcelWorkbook workbook, StampData data, int sheetsCount)
        {
            const string invisibleChar = " "; //alt + 0160

            ExcelWorksheet ws = workbook.Worksheets.Add("Штамп");
            ws.Cells.Style.Font.Size = 11;
            ws.Cells.Style.Font.Name = "GOST type A";

            const double width10mm = 5.36; //5.35
            ws.Column(1).Width = width10mm;
            ws.Column(2).Width = width10mm;
            ws.Column(3).Width = width10mm;
            ws.Column(4).Width = width10mm;
            ws.Column(5).Width = width10mm * 1.5;
            ws.Column(6).Width = width10mm;
            ws.Column(7).Width = width10mm * 7;
            ws.Column(8).Width = width10mm * 1.5;
            ws.Column(9).Width = width10mm * 1.5;
            ws.Column(10).Width = width10mm;
            ws.Column(11).Width = width10mm;

            const double height5mm = 15; //15.12
            ws.Row(1).Height = height5mm;
            ws.Row(2).Height = height5mm;
            ws.Row(3).Height = height5mm;
            ws.Row(4).Height = height5mm;
            ws.Row(5).Height = height5mm;
            ws.Row(6).Height = height5mm;
            ws.Row(7).Height = height5mm;
            ws.Row(8).Height = height5mm;
            ws.Row(9).Height = height5mm;
            ws.Row(10).Height = height5mm;
            ws.Row(11).Height = height5mm;

            ws.Cells["A1:K11"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            ws.Cells["A1:F1"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            ws.Cells["A2:F2"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            ws.Cells["A3:F3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            ws.Cells["A4:F4"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            ws.Cells["A5:F5"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            ws.Cells["A6:F6"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            ws.Cells["A7:F7"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            ws.Cells["A8:F8"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            ws.Cells["A9:F9"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            ws.Cells["A10:F10"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            ws.Cells["A11:F11"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            ws.Cells["A1:B1"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            ws.Cells["C1:D1"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            ws.Cells["A6:B11"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            ws.Cells["E1:E11"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            ws.Cells["F1:F1"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            ws.Cells["A1:A5"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            ws.Cells["C1:C5"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            ws.Cells["G1:K2"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            ws.Cells["G6:G8"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            ws.Cells["G9:G11"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            ws.Cells["H6:K11"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            ws.Cells["H7:K8"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            ws.Cells["H6:H8"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            ws.Cells["J6:K8"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

            ws.Cells["A6:B6"].Merge = true;
            ws.Cells["A7:B7"].Merge = true;
            ws.Cells["A8:B8"].Merge = true;
            ws.Cells["A9:B9"].Merge = true;
            ws.Cells["A10:B10"].Merge = true;
            ws.Cells["A11:B11"].Merge = true;
            ws.Cells["C6:D6"].Merge = true;
            ws.Cells["C7:D7"].Merge = true;
            ws.Cells["C8:D8"].Merge = true;
            ws.Cells["C9:D9"].Merge = true;
            ws.Cells["C10:D10"].Merge = true;
            ws.Cells["C11:D11"].Merge = true;
            ws.Cells["G1:K2"].Merge = true;
            ws.Cells["G3:K5"].Merge = true;
            ws.Cells["I7:I8"].Merge = true;
            ws.Cells["G9:G11"].Merge = true;
            ws.Cells["J6:K6"].Merge = true;
            ws.Cells["G6:G8"].Merge = true;
            ws.Cells["H7:H8"].Merge = true;
            ws.Cells["J7:K8"].Merge = true;
            ws.Cells["H9:K11"].Merge = true;

            ws.Cells["A5"].Value = "Изм.";
            ws.Cells["A5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["B5"].Value = "Кол.";
            ws.Cells["B5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["C5"].Value = "Лист";
            ws.Cells["C5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["D5"].Value = "№ док.";
            ws.Cells["D5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["E5"].Value = "Подп.";
            ws.Cells["E5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["F5"].Value = "Дата";
            ws.Cells["F5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A6"].Value = "Разраб.";
            ws.Cells["A10"].Value = "Н. контр.";
            ws.Cells["H6"].Value = "Стадия";
            ws.Cells["H6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["I6"].Value = "Лист";
            ws.Cells["I6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["I7"].Value = $"{invisibleChar}1";
            ws.Cells["I7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["I7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["J6"].Value = "Листов";
            ws.Cells["J6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["J7"].Value = $"{invisibleChar}{sheetsCount.ToString()}";
            ws.Cells["J7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["J7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["G1"].Value = data.Cipher;
            ws.Cells["G1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["G1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["G3"].Value = $"{data.BuildingObject1}\r\n{data.BuildingObject2}\r\n{data.BuildingObject3}";
            ws.Cells["G3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["G3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["G3"].Style.WrapText = true;
            ws.Cells["G6"].Value = $"{data.BuildingName1}\r\n{data.BuildingName2}\r\n{data.BuildingName3}";
            ws.Cells["G6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["G6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["G6"].Style.WrapText = true;
            if (data.AttName6 != null)
            {
                if (data.AttName6.Length > 0)
                {
                    ws.Cells["A6"].Value = data.AttName6;
                }
            }
            ws.Cells["A7"].Value = data.AttName7;
            ws.Cells["A8"].Value = data.AttName8;
            ws.Cells["A9"].Value = data.AttName9;
            if (data.AttName10 != null)
            {
                if (data.AttName10.Length > 0)
                {
                    ws.Cells["A10"].Value = data.AttName10;
                }
            }
            ws.Cells["A11"].Value = data.AttName11;
            ws.Cells["C6"].Value = data.AttValue6;
            ws.Cells["C7"].Value = data.AttValue7;
            ws.Cells["C8"].Value = data.AttValue8;
            ws.Cells["C9"].Value = data.AttValue9;
            ws.Cells["C10"].Value = data.AttValue10;
            ws.Cells["C11"].Value = data.AttValue11;
            ws.Cells["H7"].Value = data.Stage;
            ws.Cells["H7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["H7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["J7"].Value = sheetsCount.ToString();
            ws.Cells["J7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["J7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            if (data.Sheets != 0)
            {
                ws.Cells["J7"].Value = data.Sheets.ToString();
            }
            ws.Cells["H9"].Value = $"{data.OrganizationName1}\r\n{data.OrganizationName2}\r\n{data.OrganizationName3}";
            ws.Cells["H9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["H9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["H9"].Style.WrapText = true;

            for (int i = 0; i < sheetsCount - 1; i++)
            {
                ws.Row((i * 3) + 12).Height = height5mm;
                ws.Row((i * 3) + 13).Height = height5mm;
                ws.Row((i * 3) + 14).Height = height5mm;

                ws.Cells[$"A{((i * 3) + 12).ToString()}:K{((i * 3) + 14).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"A{((i * 3) + 12).ToString()}:F{((i * 3) + 12).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"A{((i * 3) + 13).ToString()}:F{((i * 3) + 13).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"A{((i * 3) + 14).ToString()}:F{((i * 3) + 14).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"A{((i * 3) + 12).ToString()}:A{((i * 3) + 14).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"C{((i * 3) + 12).ToString()}:C{((i * 3) + 14).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"E{((i * 3) + 12).ToString()}:E{((i * 3) + 14).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"K{((i * 3) + 12).ToString()}:K{((i * 3) + 14).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"K{((i * 3) + 12).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                ws.Cells[$"G{((i * 3) + 12).ToString()}:J{((i * 3) + 14).ToString()}"].Merge = true;
                ws.Cells[$"K{((i * 3) + 13).ToString()}:K{((i * 3) + 14).ToString()}"].Merge = true;

                ws.Cells[$"A{((i * 3) + 14).ToString()}"].Value = "Изм.";
                ws.Cells[$"A{((i * 3) + 14).ToString()}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[$"B{((i * 3) + 14).ToString()}"].Value = "Кол.";
                ws.Cells[$"B{((i * 3) + 14).ToString()}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[$"C{((i * 3) + 14).ToString()}"].Value = "Лист";
                ws.Cells[$"C{((i * 3) + 14).ToString()}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[$"D{((i * 3) + 14).ToString()}"].Value = "№ док.";
                ws.Cells[$"D{((i * 3) + 14).ToString()}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[$"E{((i * 3) + 14).ToString()}"].Value = "Подп.";
                ws.Cells[$"E{((i * 3) + 14).ToString()}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[$"F{((i * 3) + 14).ToString()}"].Value = "Дата";
                ws.Cells[$"F{((i * 3) + 14).ToString()}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[$"K{((i * 3) + 12).ToString()}"].Value = "Лист";
                ws.Cells[$"K{((i * 3) + 14).ToString()}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[$"K{((i * 3) + 13).ToString()}"].Value = $"{invisibleChar}{i + 2}";
                ws.Cells[$"K{((i * 3) + 13).ToString()}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[$"K{((i * 3) + 13).ToString()}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                ws.Cells[$"G{((i * 3) + 12).ToString()}"].Value = data.Cipher;
                ws.Cells[$"G{((i * 3) + 12).ToString()}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[$"G{((i * 3) + 12).ToString()}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            }

            ws.Hidden = eWorkSheetHidden.VeryHidden;
        }

        internal void DrawStampInline(ExcelWorksheet ws, int sheetIndex, int lastRowOnPage)
        {
            int stampUpRow = 0;

            if (sheetIndex == 1)
            {
                stampUpRow = lastRowOnPage - 10;
                ws.Cells[$"M{stampUpRow.ToString()}:X{lastRowOnPage.ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                ws.Cells[$"M{(stampUpRow + 4).ToString()}"].Value = "Изм.";
                ws.Cells[$"N{(stampUpRow + 4).ToString()}"].Value = "Кол. уч.";
                ws.Cells[$"O{(stampUpRow + 4).ToString()}"].Value = "Лист";
                ws.Cells[$"P{(stampUpRow + 4).ToString()}"].Value = "№ док.";
                ws.Cells[$"Q{(stampUpRow + 4).ToString()}"].Value = "Подп.";
                ws.Cells[$"R{(stampUpRow + 4).ToString()}"].Value = "Дата";
                ws.Cells[$"U{(stampUpRow + 5).ToString()}"].Value = "Стадия";
                ws.Cells[$"V{(stampUpRow + 5).ToString()}"].Value = "Лист";
                ws.Cells[$"W{(stampUpRow + 5).ToString()}"].Value = "Листов";

                ws.Cells[$"M{stampUpRow.ToString()}:R{stampUpRow.ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"M{(stampUpRow + 1).ToString()}:R{(stampUpRow + 1).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"M{(stampUpRow + 2).ToString()}:R{(stampUpRow + 2).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"M{(stampUpRow + 3).ToString()}:R{(stampUpRow + 3).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"M{(stampUpRow + 4).ToString()}:R{(stampUpRow + 4).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"M{(stampUpRow + 5).ToString()}:R{(stampUpRow + 5).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"M{(stampUpRow + 6).ToString()}:R{(stampUpRow + 6).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"M{(stampUpRow + 7).ToString()}:R{(stampUpRow + 7).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"M{(stampUpRow + 8).ToString()}:R{(stampUpRow + 8).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"M{(stampUpRow + 9).ToString()}:R{(stampUpRow + 9).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"M{(stampUpRow + 10).ToString()}:R{(stampUpRow + 10).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                ws.Cells[$"M{stampUpRow.ToString()}:M{(stampUpRow + 10).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"N{stampUpRow.ToString()}:N{(stampUpRow + 10).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"O{stampUpRow.ToString()}:O{(stampUpRow + 4).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"P{stampUpRow.ToString()}:P{(stampUpRow + 4).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"Q{stampUpRow.ToString()}:Q{(stampUpRow + 10).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"R{stampUpRow.ToString()}:R{(stampUpRow + 10).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                ws.Cells[$"S{stampUpRow.ToString()}:X{(stampUpRow + 1).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"S{(stampUpRow + 2).ToString()}:X{(stampUpRow + 4).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"S{(stampUpRow + 5).ToString()}:X{(stampUpRow + 7).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"S{(stampUpRow + 8).ToString()}:X{(stampUpRow + 10).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                ws.Cells[$"U{(stampUpRow + 5).ToString()}:X{(stampUpRow + 5).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"U{(stampUpRow + 5).ToString()}:X{(stampUpRow + 10).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"U{(stampUpRow + 5).ToString()}:U{(stampUpRow + 7).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"V{(stampUpRow + 5).ToString()}:V{(stampUpRow + 7).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"W{(stampUpRow + 5).ToString()}:X{(stampUpRow + 7).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            }
            else
            {
                stampUpRow = lastRowOnPage - 2;
                ws.Cells[$"M{stampUpRow.ToString()}:X{lastRowOnPage.ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                ws.Cells[$"M{(stampUpRow + 2).ToString()}"].Value = "Изм.";
                ws.Cells[$"N{(stampUpRow + 2).ToString()}"].Value = "Кол. уч.";
                ws.Cells[$"O{(stampUpRow + 2).ToString()}"].Value = "Лист";
                ws.Cells[$"P{(stampUpRow + 2).ToString()}"].Value = "№ док.";
                ws.Cells[$"Q{(stampUpRow + 2).ToString()}"].Value = "Подп.";
                ws.Cells[$"R{(stampUpRow + 2).ToString()}"].Value = "Дата";
                ws.Cells[$"X{(stampUpRow).ToString()}"].Value = "Лист";

                ws.Cells[$"M{stampUpRow.ToString()}:R{stampUpRow.ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"M{(stampUpRow + 1).ToString()}:R{(stampUpRow + 1).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"M{(stampUpRow + 2).ToString()}:R{(stampUpRow + 2).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                ws.Cells[$"M{stampUpRow.ToString()}:M{(stampUpRow + 2).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"N{stampUpRow.ToString()}:N{(stampUpRow + 2).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"O{stampUpRow.ToString()}:O{(stampUpRow + 2).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"P{stampUpRow.ToString()}:P{(stampUpRow + 2).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"Q{stampUpRow.ToString()}:Q{(stampUpRow + 2).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"R{stampUpRow.ToString()}:R{(stampUpRow + 2).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                ws.Cells[$"X{stampUpRow.ToString()}:X{(stampUpRow + 2).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                ws.Cells[$"X{stampUpRow.ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            }
        }

        internal static void InsertStampVba(ExcelPackage package)
        {
            package.Workbook.CreateVBAProject();
            StringBuilder vbaCode = new StringBuilder();
            package.Workbook.CodeModule.Code += "Private Sub Workbook_Open()\r\nInsertStamp (2)\r\nInsertStamp (3)\r\nInsertStamp (4)\r\nInsertStamp (5)\r\nInsertStamp (6)\r\nInsertStamp (7)\r\nInsertStamp (8)\r\nInsertStamp (9)\r\nInsertStamp (10)\r\nInsertStamp (11)\r\nInsertStamp (12)\r\nInsertStamp (13)\r\nInsertStamp (14)\r\nInsertStamp (15)\r\nInsertStamp (16)\r\nInsertStamp (17)\r\nInsertStamp (18)\r\nInsertStamp (19)\r\nInsertStamp (20)\r\nSheets(\"Штамп\").Activate\r\nIf SheetExists(\"Лист 1\", ThisWorkbook) = True Then\r\nSheets(\"Штамп\").Range(\"A1:K11\").Select\r\nSelection.CopyPicture Appearance:=xlScreen, Format:=xlPicture\r\nRange(\"A1\").Select\r\nSheets(\"Лист 1\").Activate\r\nIf ActiveSheet.Pictures.Count = 0 Then\r\nRange(\"L28:U34\").PasteSpecial _\r\nOperation:=xlPasteSpecialOperationAdd\r\nActiveSheet.Shapes.Range(Array(\"Picture 1\")).Select\r\nSelection.ShapeRange.IncrementLeft 1.5\r\nSelection.ShapeRange.IncrementTop 12.6\r\nEnd If\r\nRange(\"A1\").Select\r\nEnd If\r\nEnd Sub\r\nFunction InsertStamp(shtNum As Integer)\r\nSheets(\"Штамп\").Activate\r\nIf SheetExists(\"Лист \" & CStr(shtNum), ThisWorkbook) = True Then\r\nSheets(\"Штамп\").Range(\"A\" & CStr(shtNum * 3 + 6) & \":K\" & CStr(shtNum * 3 + 8)).Select\r\nSelection.CopyPicture Appearance:=xlScreen, Format:=xlPicture\r\nRange(\"A1\").Select\r\nSheets(\"Лист \" & CStr(shtNum)).Activate\r\nIf ActiveSheet.Pictures.Count = 0 Then\r\nRange(\"L33:U34\").PasteSpecial _\r\nOperation:=xlPasteSpecialOperationAdd\r\nActiveSheet.Shapes.Range(Array(\"Picture 1\")).Select\r\nSelection.ShapeRange.IncrementLeft 1.5\r\nSelection.ShapeRange.IncrementTop 10.6\r\nEnd If\r\nRange(\"A1\").Select\r\nEnd If\r\nEnd Function\r\nFunction SheetExists(shtName As String, Optional wb As Workbook) As Boolean\r\nDim sht As Worksheet\r\nIf wb Is Nothing Then Set wb = ThisWorkbook\r\nOn Error Resume Next\r\nSet sht = wb.Sheets(shtName)\r\nOn Error GoTo 0\r\nSheetExists = Not sht Is Nothing\r\nEnd Function\r\n";
        }

        internal static SortedDictionary<string, int> CreateTable(string file, StampData stampData, string path, bool insertStamp, string units = "т", int unitDigits = 2)
        {
            List<MaterialSumCell> materialSumCells = new List<MaterialSumCell>();
            SortedDictionary<string, int> sheetData = new SortedDictionary<string, int>();

            double coeff = 1;
            if (units == "кг")
            {
                coeff = .0001;
            }
            string format = string.Empty;
            string formatTmp = string.Empty;
            if (unitDigits < 1)
            {
                formatTmp = "0";
            }
            else
            {
                formatTmp = "0.";
            }
            for (int i = 0; i < unitDigits; i++)
            {
                formatTmp += "0";
            }
            format = $"{formatTmp};-{formatTmp};\"\"";
            string ext = "xlsx";
            if (insertStamp == true)
            {
                ext = "xlsm";
            }

            string filePath = $@"{path}\Спецификация {DateTime.Now.ToString("dd.MM.yy H.mm")}.{ext}";
            FileInfo newFile = new FileInfo(filePath);
            if (newFile.Exists)
            {
                newFile.Delete();
                newFile = new FileInfo(filePath);
            }

            Specification spec = JsonConvert.DeserializeObject<Specification>(File.ReadAllText(file));

            List<Detail> details = new List<Detail>();
            List<Detail> unknownDetails = new List<Detail>();
            string[] categories = spec.Headers;

            const int firstSheetLineCount = 22;
            const int secondSheetLineCount = 27;
            int sheetNumber = 1;
            int lineCount = 0;
            List<List<Document>> s = new List<List<Document>>();
            List<Document> t = new List<Document>();
            foreach (Document curStandardGroup in spec.Documents)
            {
                int standardGroupLines = 0;
                foreach (Material curMaterialGroup in curStandardGroup.Materials)
                {
                    standardGroupLines += curMaterialGroup.Profiles.Count;
                    standardGroupLines++;
                }

                int linesCount = 0;
                if (curStandardGroup.Name != null)
                {
                    if (curStandardGroup.Name.Length <= 18)
                    {
                        linesCount = 1;
                    }
                    else if (curStandardGroup.Name.Length <= 43)
                    {
                        linesCount = 2;
                    }
                    else if (curStandardGroup.Name.Length <= 62)
                    {
                        linesCount = 3;
                    }
                    else if (curStandardGroup.Name.Length <= 100)
                    {
                        linesCount = 4;
                    }
                    else
                    {
                        linesCount = 5;
                    }
                }
                curStandardGroup.lines = standardGroupLines;
                curStandardGroup.linesMin = linesCount;
                if (standardGroupLines < linesCount)
                {
                    lineCount += linesCount - standardGroupLines;
                }

                standardGroupLines++;
                lineCount += standardGroupLines;

                if (sheetNumber == 1)
                {
                    if (lineCount / firstSheetLineCount < 1)
                    {
                        t.Add(curStandardGroup);
                    }
                    else
                    {
                        s.Add(t);
                        t = new List<Document>();
                        t.Add(curStandardGroup);
                        sheetNumber = 2;
                        lineCount = standardGroupLines;
                    }
                }
                else
                {
                    if (lineCount / secondSheetLineCount < 1)
                    {
                        t.Add(curStandardGroup);
                    }
                    else
                    {
                        s.Add(t);
                        t = new List<Document>();
                        t.Add(curStandardGroup);
                        sheetNumber++;
                        lineCount = standardGroupLines;
                    }
                }
            }
            if (t.Count > 0)
            {
                s.Add(t);
            }

            if (s.Count > 0)
            {
                using (ExcelPackage package = new ExcelPackage(newFile))
                {
                    package.Workbook.Properties.Author = "VNP";
                    package.Workbook.Properties.Title = "Metal Specification";

                    int sheetIndex = 0;
                    ExcelWorksheet lastWs = null;
                    int lastI = 0;
                    int index = 0;

                    string weight5Total = string.Empty;
                    string weight6Total = string.Empty;
                    string weight7Total = string.Empty;
                    string weight8Total = string.Empty;
                    string weight9Total = string.Empty;
                    string weight10Total = string.Empty;
                    string weight11Total = string.Empty;
                    string weight12Total = string.Empty;
                    string weight13Total = string.Empty;
                    string weight14Total = string.Empty;
                    string weight15Total = string.Empty;
                    string weight16Total = string.Empty;
                    string weight17Total = string.Empty;
                    string weight18Total = string.Empty;

                    double firstRowHeight = 23.0; //8mm
                    double secondRowHeight = 63.0; //22mm
                    double thirdRowHeight = 12.0; //4mm
                    double generalRowHeight = 23.5; //8mm

                    foreach (List<Document> item in s)
                    {
                        sheetIndex++;

                        ExcelWorksheet ws = CreateWorksheet(package, sheetIndex, firstRowHeight, secondRowHeight, thirdRowHeight, generalRowHeight);
                        DrawWorksheetHeader(units, categories, ws);

                        int i = 0;

                        foreach (Document curStandardGroup in item)
                        {
                            ws.Cells[i + 4, 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            ws.Cells[i + 4, 1].Style.WrapText = true;
                            ws.Cells[i + 4, 1].Value = curStandardGroup.Name; //Стандарт
                            string standardCellFirst = ws.Cells[i + 4, 1].Address;
                            string standardCellSecond = string.Empty;

                            string weight5All = string.Empty;
                            string weight6All = string.Empty;
                            string weight7All = string.Empty;
                            string weight8All = string.Empty;
                            string weight9All = string.Empty;
                            string weight10All = string.Empty;
                            string weight11All = string.Empty;
                            string weight12All = string.Empty;
                            string weight13All = string.Empty;
                            string weight14All = string.Empty;
                            string weight15All = string.Empty;
                            string weight16All = string.Empty;
                            string weight17All = string.Empty;
                            string weight18All = string.Empty;

                            foreach (Material curMaterialGroup in curStandardGroup.Materials)
                            {
                                ws.Cells[i + 4, 2].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                ws.Cells[i + 4, 2].Style.WrapText = true;
                                ws.Cells[i + 4, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                ws.Cells[i + 4, 2].Value = curMaterialGroup.Name; //Материал
                                string materialCellFirst = ws.Cells[i + 4, 2].Address;
                                string materialCellSecond = string.Empty;

                                string weight5Sum = string.Empty;
                                string weight6Sum = string.Empty;
                                string weight7Sum = string.Empty;
                                string weight8Sum = string.Empty;
                                string weight9Sum = string.Empty;
                                string weight10Sum = string.Empty;
                                string weight11Sum = string.Empty;
                                string weight12Sum = string.Empty;
                                string weight13Sum = string.Empty;
                                string weight14Sum = string.Empty;
                                string weight15Sum = string.Empty;
                                string weight16Sum = string.Empty;
                                string weight17Sum = string.Empty;
                                string weight18Sum = string.Empty;

                                foreach (Profile curProfileGroup in curMaterialGroup.Profiles)
                                {
                                    if (curProfileGroup.Name != null)
                                    {
                                        if ((curProfileGroup.Name.Length > 0) && (curProfileGroup.Name.Contains(' ')) && (!curProfileGroup.Name.Contains("?")))
                                        {
                                            string symbol = string.Empty;
                                            string[] split = curProfileGroup.Name.Split(' ');
                                            string start = split[0].ToUpper();
                                            string second = split[1].ToUpper();
                                            string main = string.Empty;
                                            //I[∟┬□○▬#♯
                                            switch (start)
                                            {
                                                case ("ДВУТАВР"):
                                                    symbol = "I";
                                                    main = curProfileGroup.Name.Substring(curProfileGroup.Name.IndexOf(' ') + 1);
                                                    ws.Cells[i + 4, 3].IsRichText = true;

                                                    ExcelRichText ert1 = ws.Cells[i + 4, 3].RichText.Add(symbol);
                                                    ert1.FontName = "Consolas";
                                                    ert1 = ws.Cells[i + 4, 3].RichText.Add(main);
                                                    ert1.FontName = "GOST type A";
                                                    break;
                                                case ("ШВЕЛЛЕР"):
                                                    main = curProfileGroup.Name.Substring(curProfileGroup.Name.IndexOf(' ') + 1);
                                                    if (curStandardGroup.Name.Contains("ГОСТ 8278"))
                                                    {
                                                        symbol = "[";

                                                        ws.Cells[i + 4, 3].IsRichText = true;
                                                        ExcelRichText ert2 = ws.Cells[i + 4, 3].RichText.Add("Гн.");
                                                        ert2.FontName = "GOST type A";
                                                        ert2 = ws.Cells[i + 4, 3].RichText.Add(symbol);
                                                        ert2.FontName = "Consolas";
                                                        ert2 = ws.Cells[i + 4, 3].RichText.Add(main);
                                                        ert2.FontName = "GOST type A";
                                                    }
                                                    else
                                                    {
                                                        symbol = "[";

                                                        ws.Cells[i + 4, 3].IsRichText = true;
                                                        ExcelRichText ert2 = ws.Cells[i + 4, 3].RichText.Add(symbol);
                                                        ert2.FontName = "Consolas";
                                                        ert2 = ws.Cells[i + 4, 3].RichText.Add(main);
                                                        ert2.FontName = "GOST type A";
                                                    }
                                                    break;
                                                case ("УГОЛОК"):
                                                    main = "∟" + curProfileGroup.Name.Substring(curProfileGroup.Name.IndexOf(' ') + 1);
                                                    if ((curStandardGroup.Name.Contains("ГОСТ 19771")) || (curStandardGroup.Name.Contains("ГОСТ 19772")))
                                                    {
                                                        main = "Гн.∟" + curProfileGroup.Name.Substring(curProfileGroup.Name.IndexOf(' ') + 1);
                                                    }

                                                    ws.Cells[i + 4, 3].Value = main;
                                                    break;
                                                case ("ТРУБА"):
                                                    main = "Т" + curProfileGroup.Name.Substring(curProfileGroup.Name.IndexOf(' ') + 1);
                                                    if (curStandardGroup.Name.Contains("ГОСТ 8732"))
                                                    {
                                                        main = "ТБ" + curProfileGroup.Name.Substring(curProfileGroup.Name.IndexOf(' ') + 1);
                                                    }
                                                    else if (curStandardGroup.Name.Contains("ГОСТ 10704"))
                                                    {
                                                        main = "ТЭ" + curProfileGroup.Name.Substring(curProfileGroup.Name.IndexOf(' ') + 1);
                                                    }
                                                    else if (curStandardGroup.Name.Contains("ТУ 67-2287"))
                                                    {
                                                        main = "Гн.□" + curProfileGroup.Name.Substring(curProfileGroup.Name.IndexOf(' ') + 4);
                                                    }
                                                    else if (curStandardGroup.Name.Contains("ТУ 36-2287"))
                                                    {
                                                        main = "Гн." + curProfileGroup.Name.Substring(curProfileGroup.Name.IndexOf(' ') + 4);
                                                    }
                                                    else if (curStandardGroup.Name.Contains("ГОСТ Р 54157"))
                                                    {
                                                        main = "ПК" + curProfileGroup.Name.Substring(curProfileGroup.Name.IndexOf(' ') + 4);
                                                    }
                                                    else if (curStandardGroup.Name.Contains("ГОСТ 30245"))
                                                    {
                                                        main = "Гн." + curProfileGroup.Name.Substring(curProfileGroup.Name.IndexOf(' ') + 7);
                                                        if (second == "(ПР.)")
                                                        {
                                                            main = "Гн.□" + curProfileGroup.Name.Substring(curProfileGroup.Name.IndexOf(' ') + 7);
                                                        }
                                                    }

                                                    ws.Cells[i + 4, 3].Value = main;
                                                    break;
                                                case ("ТАВР"):
                                                    main = "Т" + curProfileGroup.Name.Substring(curProfileGroup.Name.IndexOf(' ') + 1);

                                                    ws.Cells[i + 4, 3].Value = main;
                                                    break;
                                                case ("КРУГ"):
                                                    main = "○" + curProfileGroup.Name.Substring(curProfileGroup.Name.IndexOf(' ') + 1);

                                                    ws.Cells[i + 4, 3].Value = main;
                                                    break;
                                                case ("КВАДРАТ"):
                                                    main = "□" + curProfileGroup.Name.Substring(curProfileGroup.Name.IndexOf(' ') + 1);

                                                    ws.Cells[i + 4, 3].Value = main;
                                                    break;
                                                case ("ЛИСТ"):
                                                    main = "▬" + curProfileGroup.Name.Substring(curProfileGroup.Name.IndexOf(' ') + 1);

                                                    ws.Cells[i + 4, 3].Value = main;
                                                    break;
                                                case ("ПВ"):
                                                    ws.Cells[i + 4, 3].Value = curProfileGroup.Name;
                                                    break;
                                                case ("РИФ"):
                                                    ws.Cells[i + 4, 3].Value = curProfileGroup.Name;
                                                    break;
                                                case ("ЧРИФ"):
                                                    ws.Cells[i + 4, 3].Value = curProfileGroup.Name;
                                                    break;
                                                default:
                                                    break;
                                            }
                                        }
                                        else
                                        {
                                            if (curProfileGroup.Name.StartsWith("КР"))
                                            {
                                                ws.Cells[i + 4, 3].Value = curProfileGroup.Name;
                                            }
                                            else if (curProfileGroup.Name.StartsWith("С10-") || curProfileGroup.Name.StartsWith("С15-") || curProfileGroup.Name.StartsWith("С18-") ||
                                                curProfileGroup.Name.StartsWith("С21-") || curProfileGroup.Name.StartsWith("С44-") || curProfileGroup.Name.StartsWith("Н57-") ||
                                                curProfileGroup.Name.StartsWith("Н60-") || curProfileGroup.Name.StartsWith("Н75-") || curProfileGroup.Name.StartsWith("Н114-") ||
                                                curProfileGroup.Name.StartsWith("НС35-") || curProfileGroup.Name.StartsWith("НС44-"))
                                            {
                                                ws.Cells[i + 4, 3].Value = curProfileGroup.Name;
                                            }
                                            else
                                            {
                                                ws.Cells[i + 4, 3].Value = $"{curProfileGroup.Name}"; //Профиль
                                            }
                                        }
                                    }
                                    ws.Cells[i + 4, 3].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                                    double? weight5 = 0.0;
                                    double? weight6 = 0.0;
                                    double? weight7 = 0.0;
                                    double? weight8 = 0.0;
                                    double? weight9 = 0.0;
                                    double? weight10 = 0.0;
                                    double? weight11 = 0.0;
                                    double? weight12 = 0.0;
                                    double? weight13 = 0.0;
                                    double? weight14 = 0.0;
                                    double? weight15 = 0.0;
                                    double? weight16 = 0.0;
                                    double? weight17 = 0.0;
                                    double? weight18 = 0.0;

                                    foreach (Construction curCategoryGroup in curProfileGroup.ConstructionTypes[0].Constructions)
                                    {
                                        ws.Row(i + 4).Height = generalRowHeight;
                                        ws.Row(i + 4).Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                        switch (curCategoryGroup.ID) //Категория
                                        {
                                            case 5:
                                                weight5 += curCategoryGroup.Count;
                                                break;
                                            case 6:
                                                weight6 += curCategoryGroup.Count;
                                                break;
                                            case 7:
                                                weight7 += curCategoryGroup.Count;
                                                break;
                                            case 8:
                                                weight8 += curCategoryGroup.Count;
                                                break;
                                            case 9:
                                                weight9 += curCategoryGroup.Count;
                                                break;
                                            case 10:
                                                weight10 += curCategoryGroup.Count;
                                                break;
                                            case 11:
                                                weight11 += curCategoryGroup.Count;
                                                break;
                                            case 12:
                                                weight12 += curCategoryGroup.Count;
                                                break;
                                            case 13:
                                                weight13 += curCategoryGroup.Count;
                                                break;
                                            case 14:
                                                weight14 += curCategoryGroup.Count;
                                                break;
                                            case 15:
                                                weight15 += curCategoryGroup.Count;
                                                break;
                                            case 16:
                                                weight16 += curCategoryGroup.Count;
                                                break;
                                            case 17:
                                                weight17 += curCategoryGroup.Count;
                                                break;
                                            case 18:
                                                weight18 += curCategoryGroup.Count;
                                                break;
                                            default:
                                                break;
                                        }

                                        ws.Cells[i + 4, 4, i + 4, 19].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                                        ws.Cells[i + 4, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        ws.Cells[i + 4, 4].Value = (index + 1).ToString();

                                        using (ExcelRange r3 = ws.Cells[i + 4, 5, i + 4, 19])
                                        {
                                            r3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                            r3.Style.Numberformat.Format = format;

                                            ws.Cells[i + 4, 5].Value = weight5 / coeff;
                                            ws.Cells[i + 4, 6].Value = weight6 / coeff;
                                            ws.Cells[i + 4, 7].Value = weight7 / coeff;
                                            ws.Cells[i + 4, 8].Value = weight8 / coeff;
                                            ws.Cells[i + 4, 9].Value = weight9 / coeff;
                                            ws.Cells[i + 4, 10].Value = weight10 / coeff;
                                            ws.Cells[i + 4, 11].Value = weight11 / coeff;
                                            ws.Cells[i + 4, 12].Value = weight12 / coeff;
                                            ws.Cells[i + 4, 13].Value = weight13 / coeff;
                                            ws.Cells[i + 4, 14].Value = weight14 / coeff;
                                            ws.Cells[i + 4, 15].Value = weight15 / coeff;
                                            ws.Cells[i + 4, 16].Value = weight16 / coeff;
                                            ws.Cells[i + 4, 17].Value = weight17 / coeff;
                                            ws.Cells[i + 4, 18].Value = weight18 / coeff;
                                            ws.Cells[i + 4, 19].CreateArrayFormula($"SUM(ROUND(({ws.Cells[i + 4, 5].Address}:{ws.Cells[i + 4, 18].Address}),2))");
                                        }
                                    }

                                    weight5Sum += $"ROUND({ws.Cells[i + 4, 5].Address},{unitDigits}),";
                                    weight6Sum += $"ROUND({ws.Cells[i + 4, 6].Address},{unitDigits}),";
                                    weight7Sum += $"ROUND({ws.Cells[i + 4, 7].Address},{unitDigits}),";
                                    weight8Sum += $"ROUND({ws.Cells[i + 4, 8].Address},{unitDigits}),";
                                    weight9Sum += $"ROUND({ws.Cells[i + 4, 9].Address},{unitDigits}),";
                                    weight10Sum += $"ROUND({ws.Cells[i + 4, 10].Address},{unitDigits}),";
                                    weight11Sum += $"ROUND({ws.Cells[i + 4, 11].Address},{unitDigits}),";
                                    weight12Sum += $"ROUND({ws.Cells[i + 4, 12].Address},{unitDigits}),";
                                    weight13Sum += $"ROUND({ws.Cells[i + 4, 13].Address},{unitDigits}),";
                                    weight14Sum += $"ROUND({ws.Cells[i + 4, 14].Address},{unitDigits}),";
                                    weight15Sum += $"ROUND({ws.Cells[i + 4, 15].Address},{unitDigits}),";
                                    weight16Sum += $"ROUND({ws.Cells[i + 4, 16].Address},{unitDigits}),";
                                    weight17Sum += $"ROUND({ws.Cells[i + 4, 17].Address},{unitDigits}),";
                                    weight18Sum += $"ROUND({ws.Cells[i + 4, 18].Address},{unitDigits}),";

                                    materialCellSecond = ws.Cells[i + 4, 2].Address;
                                    i++;
                                    index++;
                                }

                                bool isRowAdded = false;
                                if (curStandardGroup.lines < curStandardGroup.linesMin)
                                {
                                    for (int j = 0; j < (curStandardGroup.linesMin - curStandardGroup.lines); j++)
                                    {
                                        ws.Cells[i + 4, 3, i + 4, 19].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                                        ws.Cells[i + 4, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                        ws.Cells[i + 4, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        materialCellSecond = ws.Cells[i + 4, 2].Address;
                                        i++;
                                    }
                                    isRowAdded = true;
                                }

                                if ((curMaterialGroup.Profiles.Count == 1) && (isRowAdded == false))
                                {
                                    ws.Cells[i + 4, 3, i + 4, 19].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                                    ws.Cells[i + 4, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    ws.Cells[i + 4, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    materialCellSecond = ws.Cells[i + 4, 2].Address;
                                    i++;
                                }

                                ws.Cells[$"{materialCellFirst}:{materialCellSecond}"].Merge = true;

                                ws.Row(i + 4).Height = generalRowHeight;
                                ws.Row(i + 4).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                ws.Cells[i + 4, 2, i + 4, 3].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                                ws.Cells[i + 4, 2].Value = "Итого:";

                                ws.Cells[i + 4, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                ws.Cells[i + 4, 4].Value = (index + 1).ToString();

                                using (ExcelRange r4 = ws.Cells[i + 4, 5, i + 4, 19])
                                {
                                    r4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    r4.Style.Numberformat.Format = format;

                                    ws.Cells[i + 4, 5].Formula = $"SUM({weight5Sum.Substring(0, weight5Sum.Length - 1)})";
                                    ws.Cells[i + 4, 6].Formula = $"SUM({weight6Sum.Substring(0, weight6Sum.Length - 1)})";
                                    ws.Cells[i + 4, 7].Formula = $"SUM({weight7Sum.Substring(0, weight7Sum.Length - 1)})";
                                    ws.Cells[i + 4, 8].Formula = $"SUM({weight8Sum.Substring(0, weight8Sum.Length - 1)})";
                                    ws.Cells[i + 4, 9].Formula = $"SUM({weight9Sum.Substring(0, weight9Sum.Length - 1)})";
                                    ws.Cells[i + 4, 10].Formula = $"SUM({weight10Sum.Substring(0, weight10Sum.Length - 1)})";
                                    ws.Cells[i + 4, 11].Formula = $"SUM({weight11Sum.Substring(0, weight11Sum.Length - 1)})";
                                    ws.Cells[i + 4, 12].Formula = $"SUM({weight12Sum.Substring(0, weight12Sum.Length - 1)})";
                                    ws.Cells[i + 4, 13].Formula = $"SUM({weight13Sum.Substring(0, weight13Sum.Length - 1)})";
                                    ws.Cells[i + 4, 14].Formula = $"SUM({weight14Sum.Substring(0, weight14Sum.Length - 1)})";
                                    ws.Cells[i + 4, 15].Formula = $"SUM({weight15Sum.Substring(0, weight15Sum.Length - 1)})";
                                    ws.Cells[i + 4, 16].Formula = $"SUM({weight16Sum.Substring(0, weight16Sum.Length - 1)})";
                                    ws.Cells[i + 4, 17].Formula = $"SUM({weight17Sum.Substring(0, weight17Sum.Length - 1)})";
                                    ws.Cells[i + 4, 18].Formula = $"SUM({weight18Sum.Substring(0, weight18Sum.Length - 1)})";
                                    ws.Cells[i + 4, 19].CreateArrayFormula($"SUM(ROUND(({ws.Cells[i + 4, 5].Address}:{ws.Cells[i + 4, 18].Address}),2))");
                                }

                                weight5All += $"{ws.Cells[i + 4, 5].Address},";
                                weight6All += $"{ws.Cells[i + 4, 6].Address},";
                                weight7All += $"{ws.Cells[i + 4, 7].Address},";
                                weight8All += $"{ws.Cells[i + 4, 8].Address},";
                                weight9All += $"{ws.Cells[i + 4, 9].Address},";
                                weight10All += $"{ws.Cells[i + 4, 10].Address},";
                                weight11All += $"{ws.Cells[i + 4, 11].Address},";
                                weight12All += $"{ws.Cells[i + 4, 12].Address},";
                                weight13All += $"{ws.Cells[i + 4, 13].Address},";
                                weight14All += $"{ws.Cells[i + 4, 14].Address},";
                                weight15All += $"{ws.Cells[i + 4, 15].Address},";
                                weight16All += $"{ws.Cells[i + 4, 16].Address},";
                                weight17All += $"{ws.Cells[i + 4, 17].Address},";
                                weight18All += $"{ws.Cells[i + 4, 18].Address},";

                                bool isMaterialExists = MaterialSumCell.ContainsMaterial(materialSumCells, curMaterialGroup.Name);
                                if (isMaterialExists == false)
                                {
                                    MaterialSumCell newMaterialSumCell = new MaterialSumCell(curMaterialGroup.Name);
                                    materialSumCells.Add(newMaterialSumCell);
                                }
                                foreach (MaterialSumCell materialSumCell in materialSumCells)
                                {
                                    if (materialSumCell.Material == curMaterialGroup.Name)
                                    {
                                        materialSumCell.Category5Addresses += $"ROUND({ws.Cells[i + 4, 5].FullAddress},2),";
                                        materialSumCell.Category6Addresses += $"ROUND({ws.Cells[i + 4, 6].FullAddress},2),";
                                        materialSumCell.Category7Addresses += $"ROUND({ws.Cells[i + 4, 7].FullAddress},2),";
                                        materialSumCell.Category8Addresses += $"ROUND({ws.Cells[i + 4, 8].FullAddress},2),";
                                        materialSumCell.Category9Addresses += $"ROUND({ws.Cells[i + 4, 9].FullAddress},2),";
                                        materialSumCell.Category10Addresses += $"ROUND({ws.Cells[i + 4, 10].FullAddress},2),";
                                        materialSumCell.Category11Addresses += $"ROUND({ws.Cells[i + 4, 11].FullAddress},2),";
                                        materialSumCell.Category12Addresses += $"ROUND({ws.Cells[i + 4, 12].FullAddress},2),";
                                        materialSumCell.Category13Addresses += $"ROUND({ws.Cells[i + 4, 13].FullAddress},2),";
                                        materialSumCell.Category14Addresses += $"ROUND({ws.Cells[i + 4, 14].FullAddress},2),";
                                        materialSumCell.Category15Addresses += $"ROUND({ws.Cells[i + 4, 15].FullAddress},2),";
                                        materialSumCell.Category16Addresses += $"ROUND({ws.Cells[i + 4, 16].FullAddress},2),";
                                        materialSumCell.Category17Addresses += $"ROUND({ws.Cells[i + 4, 17].FullAddress},2),";
                                        materialSumCell.Category18Addresses += $"ROUND({ws.Cells[i + 4, 18].FullAddress},2),";
                                    }
                                }

                                standardCellSecond = ws.Cells[i + 4, 1].Address;
                                i++;
                                index++;
                            }

                            ws.Cells[$"{standardCellFirst}:{standardCellSecond}"].Merge = true;

                            ws.Row(i + 4).Height = generalRowHeight;
                            ws.Row(i + 4).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            ws.Cells[i + 4, 1, i + 4, 19].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            ws.Cells[i + 4, 1].Value = "Всего профиля:";
                            ws.Cells[i + 4, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            ws.Cells[i + 4, 4].Value = (index + 1).ToString();

                            using (ExcelRange r5 = ws.Cells[i + 4, 5, i + 4, 19])
                            {
                                r5.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                r5.Style.Numberformat.Format = format;

                                ws.Cells[i + 4, 5].Formula = $"SUM({weight5All.Substring(0, weight5All.Length - 1)})";
                                ws.Cells[i + 4, 6].Formula = $"SUM({weight6All.Substring(0, weight6All.Length - 1)})";
                                ws.Cells[i + 4, 7].Formula = $"SUM({weight7All.Substring(0, weight7All.Length - 1)})";
                                ws.Cells[i + 4, 8].Formula = $"SUM({weight8All.Substring(0, weight8All.Length - 1)})";
                                ws.Cells[i + 4, 9].Formula = $"SUM({weight9All.Substring(0, weight9All.Length - 1)})";
                                ws.Cells[i + 4, 10].Formula = $"SUM({weight10All.Substring(0, weight10All.Length - 1)})";
                                ws.Cells[i + 4, 11].Formula = $"SUM({weight11All.Substring(0, weight11All.Length - 1)})";
                                ws.Cells[i + 4, 12].Formula = $"SUM({weight12All.Substring(0, weight12All.Length - 1)})";
                                ws.Cells[i + 4, 13].Formula = $"SUM({weight13All.Substring(0, weight13All.Length - 1)})";
                                ws.Cells[i + 4, 14].Formula = $"SUM({weight14All.Substring(0, weight14All.Length - 1)})";
                                ws.Cells[i + 4, 15].Formula = $"SUM({weight15All.Substring(0, weight15All.Length - 1)})";
                                ws.Cells[i + 4, 16].Formula = $"SUM({weight16All.Substring(0, weight16All.Length - 1)})";
                                ws.Cells[i + 4, 17].Formula = $"SUM({weight17All.Substring(0, weight17All.Length - 1)})";
                                ws.Cells[i + 4, 18].Formula = $"SUM({weight18All.Substring(0, weight18All.Length - 1)})";
                                ws.Cells[i + 4, 19].CreateArrayFormula($"SUM(ROUND(({ws.Cells[i + 4, 5].Address}:{ws.Cells[i + 4, 18].Address}),2))");
                            }

                            weight5Total += $"ROUND({ws.Cells[i + 4, 5].FullAddress},2),";
                            weight6Total += $"ROUND({ws.Cells[i + 4, 6].FullAddress},2),";
                            weight7Total += $"ROUND({ws.Cells[i + 4, 7].FullAddress},2),";
                            weight8Total += $"ROUND({ws.Cells[i + 4, 8].FullAddress},2),";
                            weight9Total += $"ROUND({ws.Cells[i + 4, 9].FullAddress},2),";
                            weight11Total += $"ROUND({ws.Cells[i + 4, 11].FullAddress},2),";
                            weight12Total += $"ROUND({ws.Cells[i + 4, 12].FullAddress},2),";
                            weight13Total += $"ROUND({ws.Cells[i + 4, 13].FullAddress},2),";
                            weight14Total += $"ROUND({ws.Cells[i + 4, 14].FullAddress},2),";
                            weight15Total += $"ROUND({ws.Cells[i + 4, 15].FullAddress},2),";
                            weight16Total += $"ROUND({ws.Cells[i + 4, 16].FullAddress},2),";
                            weight10Total += $"ROUND({ws.Cells[i + 4, 10].FullAddress},2),";
                            weight17Total += $"ROUND({ws.Cells[i + 4, 17].FullAddress},2),";
                            weight18Total += $"ROUND({ws.Cells[i + 4, 18].FullAddress},2),";

                            i++;
                            index++;
                        }

                        ws.Cells[$"A4:A{(i + 3).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        ws.Cells[$"B4:B{(i + 3).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        ws.Cells[$"C4:C{(i + 3).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        ws.Cells[$"D4:D{(i + 3).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        ws.Cells[$"E4:E{(i + 3).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        ws.Cells[$"F4:F{(i + 3).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        ws.Cells[$"G4:G{(i + 3).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        ws.Cells[$"H4:H{(i + 3).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        ws.Cells[$"I4:I{(i + 3).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        ws.Cells[$"J4:J{(i + 3).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        ws.Cells[$"K4:K{(i + 3).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        ws.Cells[$"L4:L{(i + 3).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        ws.Cells[$"M4:M{(i + 3).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        ws.Cells[$"N4:N{(i + 3).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        ws.Cells[$"O4:O{(i + 3).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        ws.Cells[$"P4:P{(i + 3).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        ws.Cells[$"Q4:Q{(i + 3).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        ws.Cells[$"R4:R{(i + 3).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        ws.Cells[$"S4:S{(i + 3).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                        lastWs = ws;
                        lastI = i;

                        sheetData.Add(ws.Name, i + 3);
                    }

                    bool newWs = false;
                    if ((lastI + materialSumCells.Count) > secondSheetLineCount)
                    {
                        ExcelWorksheet ws = CreateWorksheet(package, sheetIndex + 1, firstRowHeight, secondRowHeight, thirdRowHeight, generalRowHeight);
                        DrawWorksheetHeader(units, categories, ws);

                        lastWs = ws;
                        lastI = 0;
                        sheetNumber++;
                        newWs = true;
                    }

                    lastWs.Row(lastI + 4).Height = generalRowHeight;
                    lastWs.Row(lastI + 4).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    lastWs.Cells[lastI + 4, 1, lastI + 4, 19].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    lastWs.Cells[lastI + 4, 1].Value = "Всего масса металла:";
                    lastWs.Cells[lastI + 4, 1].Style.WrapText = true;
                    lastWs.Cells[lastI + 4, 1, lastI + 4, 2].Merge = true;
                    lastWs.Cells[lastI + 4, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    lastWs.Cells[lastI + 4, 4].Value = (index + 1).ToString();

                    using (ExcelRange r6 = lastWs.Cells[lastI + 4, 5, lastI + 4, 19])
                    {
                        r6.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        r6.Style.Numberformat.Format = format;

                        lastWs.Cells[lastI + 4, 5].Formula = ($"SUM({weight5Total.Substring(0, weight5Total.Length - 1)})");
                        lastWs.Cells[lastI + 4, 6].Formula = ($"SUM({weight6Total.Substring(0, weight6Total.Length - 1)})");
                        lastWs.Cells[lastI + 4, 7].Formula = ($"SUM({weight7Total.Substring(0, weight7Total.Length - 1)})");
                        lastWs.Cells[lastI + 4, 8].Formula = ($"SUM({weight8Total.Substring(0, weight8Total.Length - 1)})");
                        lastWs.Cells[lastI + 4, 9].Formula = ($"SUM({weight9Total.Substring(0, weight9Total.Length - 1)})");
                        lastWs.Cells[lastI + 4, 10].Formula = ($"SUM({weight10Total.Substring(0, weight10Total.Length - 1)})");
                        lastWs.Cells[lastI + 4, 11].Formula = ($"SUM({weight11Total.Substring(0, weight11Total.Length - 1)})");
                        lastWs.Cells[lastI + 4, 12].Formula = ($"SUM({weight12Total.Substring(0, weight12Total.Length - 1)})");
                        lastWs.Cells[lastI + 4, 13].Formula = ($"SUM({weight13Total.Substring(0, weight13Total.Length - 1)})");
                        lastWs.Cells[lastI + 4, 14].Formula = ($"SUM({weight14Total.Substring(0, weight14Total.Length - 1)})");
                        lastWs.Cells[lastI + 4, 15].Formula = ($"SUM({weight15Total.Substring(0, weight15Total.Length - 1)})");
                        lastWs.Cells[lastI + 4, 16].Formula = ($"SUM({weight16Total.Substring(0, weight16Total.Length - 1)})");
                        lastWs.Cells[lastI + 4, 17].Formula = ($"SUM({weight17Total.Substring(0, weight17Total.Length - 1)})");
                        lastWs.Cells[lastI + 4, 18].Formula = ($"SUM({weight18Total.Substring(0, weight18Total.Length - 1)})");
                        lastWs.Cells[lastI + 4, 19].CreateArrayFormula($"SUM(ROUND(({lastWs.Cells[lastI + 4, 5].Address}:{lastWs.Cells[lastI + 4, 18].Address}),2))");
                    }

                    lastWs.Row(lastI + 5).Height = generalRowHeight;
                    lastWs.Row(lastI + 6).Height = generalRowHeight;
                    lastWs.Row(lastI + 5).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    lastWs.Cells[lastI + 5, 1].Value = "В том числе по маркам или наименованиям:";
                    lastWs.Cells[lastI + 5, 1].Style.WrapText = true;
                    lastWs.Cells[lastI + 5, 1, lastI + 6, 2].Merge = true;
                    lastWs.Cells[lastI + 5, 3, lastI + 6, 3].Merge = true;
                    lastWs.Cells[lastI + 5, 4, lastI + 6, 4].Merge = true;
                    lastWs.Cells[lastI + 5, 5, lastI + 6, 5].Merge = true;
                    lastWs.Cells[lastI + 5, 6, lastI + 6, 6].Merge = true;
                    lastWs.Cells[lastI + 5, 7, lastI + 6, 7].Merge = true;
                    lastWs.Cells[lastI + 5, 8, lastI + 6, 8].Merge = true;
                    lastWs.Cells[lastI + 5, 9, lastI + 6, 9].Merge = true;
                    lastWs.Cells[lastI + 5, 10, lastI + 6, 10].Merge = true;
                    lastWs.Cells[lastI + 5, 11, lastI + 6, 11].Merge = true;
                    lastWs.Cells[lastI + 5, 12, lastI + 6, 12].Merge = true;
                    lastWs.Cells[lastI + 5, 13, lastI + 6, 13].Merge = true;
                    lastWs.Cells[lastI + 5, 14, lastI + 6, 14].Merge = true;
                    lastWs.Cells[lastI + 5, 15, lastI + 6, 15].Merge = true;
                    lastWs.Cells[lastI + 5, 16, lastI + 6, 16].Merge = true;
                    lastWs.Cells[lastI + 5, 17, lastI + 6, 17].Merge = true;
                    lastWs.Cells[lastI + 5, 18, lastI + 6, 18].Merge = true;
                    lastWs.Cells[lastI + 5, 19, lastI + 6, 19].Merge = true;
                    lastWs.Cells[lastI + 5, 19].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    lastWs.Cells[lastI + 5, 19].Style.Numberformat.Format = format;

                    var materialSumCellsOrdered = materialSumCells.OrderBy(x => x.Material);
                    string materialSumCellFirst = lastWs.Cells[lastI + 7, 19].Address;
                    string materialSumCellLast = string.Empty;
                    foreach (MaterialSumCell materialSumCell in materialSumCellsOrdered)
                    {
                        lastWs.Row(lastI + 7).Height = generalRowHeight;
                        lastWs.Row(lastI + 7).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        lastWs.Row(lastI + 7).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        if (materialSumCell.Material != null)
                        {
                            int indexOf = materialSumCell.Material.IndexOf("\r\n");
                            if (indexOf != -1)
                            {
                                lastWs.Cells[lastI + 7, 1].Value = materialSumCell.Material.Substring(0, indexOf);
                            }
                            else
                            {
                                lastWs.Cells[lastI + 7, 1].Value = materialSumCell.Material;
                            }
                        }
                        lastWs.Cells[lastI + 7, 1].Style.WrapText = true;
                        lastWs.Cells[lastI + 7, 1, lastI + 7, 2].Merge = true;
                        lastWs.Cells[lastI + 7, 1, lastI + 7, 19].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        lastWs.Cells[lastI + 7, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        lastWs.Cells[lastI + 7, 4].Value = (index + 2).ToString();
                        using (ExcelRange r7 = lastWs.Cells[lastI + 7, 5, lastI + 7, 18])
                        {
                            r7.Style.Numberformat.Format = format;
                            r7.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                            lastWs.Cells[lastI + 7, 5].Formula = ($"SUM({materialSumCell.Category5Addresses.Substring(0, materialSumCell.Category5Addresses.Length - 1)})");
                            lastWs.Cells[lastI + 7, 6].Formula = ($"SUM({materialSumCell.Category6Addresses.Substring(0, materialSumCell.Category6Addresses.Length - 1)})");
                            lastWs.Cells[lastI + 7, 7].Formula = ($"SUM({materialSumCell.Category7Addresses.Substring(0, materialSumCell.Category7Addresses.Length - 1)})");
                            lastWs.Cells[lastI + 7, 8].Formula = ($"SUM({materialSumCell.Category8Addresses.Substring(0, materialSumCell.Category8Addresses.Length - 1)})");
                            lastWs.Cells[lastI + 7, 9].Formula = ($"SUM({materialSumCell.Category9Addresses.Substring(0, materialSumCell.Category9Addresses.Length - 1)})");
                            lastWs.Cells[lastI + 7, 10].Formula = ($"SUM({materialSumCell.Category10Addresses.Substring(0, materialSumCell.Category10Addresses.Length - 1)})");
                            lastWs.Cells[lastI + 7, 11].Formula = ($"SUM({materialSumCell.Category11Addresses.Substring(0, materialSumCell.Category11Addresses.Length - 1)})");
                            lastWs.Cells[lastI + 7, 12].Formula = ($"SUM({materialSumCell.Category12Addresses.Substring(0, materialSumCell.Category12Addresses.Length - 1)})");
                            lastWs.Cells[lastI + 7, 13].Formula = ($"SUM({materialSumCell.Category13Addresses.Substring(0, materialSumCell.Category13Addresses.Length - 1)})");
                            lastWs.Cells[lastI + 7, 14].Formula = ($"SUM({materialSumCell.Category14Addresses.Substring(0, materialSumCell.Category14Addresses.Length - 1)})");
                            lastWs.Cells[lastI + 7, 15].Formula = ($"SUM({materialSumCell.Category15Addresses.Substring(0, materialSumCell.Category15Addresses.Length - 1)})");
                            lastWs.Cells[lastI + 7, 16].Formula = ($"SUM({materialSumCell.Category16Addresses.Substring(0, materialSumCell.Category16Addresses.Length - 1)})");
                            lastWs.Cells[lastI + 7, 17].Formula = ($"SUM({materialSumCell.Category17Addresses.Substring(0, materialSumCell.Category17Addresses.Length - 1)})");
                            lastWs.Cells[lastI + 7, 18].Formula = ($"SUM({materialSumCell.Category18Addresses.Substring(0, materialSumCell.Category18Addresses.Length - 1)})");
                        }
                        lastWs.Cells[lastI + 7, 19].CreateArrayFormula($"SUM(ROUND(({lastWs.Cells[lastI + 7, 5].Address}:{lastWs.Cells[lastI + 7, 18].Address}),2))");

                        lastWs.Cells[$"A{(lastI + 4).ToString()}:B{(lastI + 7).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        lastWs.Cells[$"C{(lastI + 4).ToString()}:C{(lastI + 7).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        lastWs.Cells[$"D{(lastI + 4).ToString()}:D{(lastI + 7).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        lastWs.Cells[$"E{(lastI + 4).ToString()}:E{(lastI + 7).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        lastWs.Cells[$"F{(lastI + 4).ToString()}:F{(lastI + 7).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        lastWs.Cells[$"G{(lastI + 4).ToString()}:G{(lastI + 7).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        lastWs.Cells[$"H{(lastI + 4).ToString()}:H{(lastI + 7).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        lastWs.Cells[$"I{(lastI + 4).ToString()}:I{(lastI + 7).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        lastWs.Cells[$"J{(lastI + 4).ToString()}:J{(lastI + 7).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        lastWs.Cells[$"K{(lastI + 4).ToString()}:K{(lastI + 7).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        lastWs.Cells[$"L{(lastI + 4).ToString()}:L{(lastI + 7).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        lastWs.Cells[$"M{(lastI + 4).ToString()}:M{(lastI + 7).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        lastWs.Cells[$"N{(lastI + 4).ToString()}:N{(lastI + 7).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        lastWs.Cells[$"O{(lastI + 4).ToString()}:O{(lastI + 7).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        lastWs.Cells[$"P{(lastI + 4).ToString()}:P{(lastI + 7).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        lastWs.Cells[$"Q{(lastI + 4).ToString()}:Q{(lastI + 7).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        lastWs.Cells[$"R{(lastI + 4).ToString()}:R{(lastI + 7).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        lastWs.Cells[$"S{(lastI + 4).ToString()}:S{(lastI + 7).ToString()}"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                        materialSumCellLast = lastWs.Cells[lastI + 7, 19].Address;

                        lastI++;
                        index++;
                    }

                    lastWs.Row(lastI + 7).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    lastWs.Cells[lastI + 7, 1].Value = "Общая масса металла:";
                    lastWs.Cells[lastI + 7, 1].Style.WrapText = true;
                    lastWs.Cells[lastI + 7, 1, lastI + 7, 2].Merge = true;

                    lastWs.Cells[lastI + 7, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    lastWs.Cells[lastI + 7, 4].Value = (index + 2).ToString();

                    lastWs.Cells[lastI + 7, 19].Formula = $"SUM({materialSumCellFirst}:{materialSumCellLast})";
                    lastWs.Cells[lastI + 7, 19].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    lastWs.Cells[lastI + 7, 19].Style.Numberformat.Format = format;

                    lastWs.Cells[lastI + 7, 1, lastI + 7, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    lastWs.Cells[lastI + 7, 3].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    lastWs.Cells[lastI + 7, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    lastWs.Cells[lastI + 7, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    lastWs.Cells[lastI + 7, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    lastWs.Cells[lastI + 7, 7].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    lastWs.Cells[lastI + 7, 8].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    lastWs.Cells[lastI + 7, 9].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    lastWs.Cells[lastI + 7, 10].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    lastWs.Cells[lastI + 7, 11].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    lastWs.Cells[lastI + 7, 12].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    lastWs.Cells[lastI + 7, 13].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    lastWs.Cells[lastI + 7, 14].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    lastWs.Cells[lastI + 7, 15].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    lastWs.Cells[lastI + 7, 16].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    lastWs.Cells[lastI + 7, 17].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    lastWs.Cells[lastI + 7, 18].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                    lastWs.Cells[lastI + 7, 19].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                    if (newWs == true)
                    {
                        sheetData.Add(lastWs.Name, lastI);
                    }
                    else
                    {
                        string lastKey = sheetData.Keys.Last();
                        sheetData[lastKey] += lastI;
                    }

                    if (insertStamp == true)
                    {
                        DrawStamp(package.Workbook, stampData, sheetNumber);
                        InsertStampVba(package);
                    }

                    if (unknownDetails.Count() > 0)
                    {
                        ExcelWorksheet unknownWs = package.Workbook.Worksheets.Add("Неизвестные детали");

                        unknownWs.Cells[1, 1].Value = "Категория";
                        unknownWs.Cells[1, 2].Value = "GUID";
                        unknownWs.Cells[1, 3].Value = "Стандарт";
                        unknownWs.Cells[1, 4].Value = "Профиль";
                        unknownWs.Cells[1, 5].Value = "Материал";
                        unknownWs.Cells[1, 6].Value = "Масса, кг";

                        int unknownIndex = 2;
                        foreach (Detail unknownDetail in unknownDetails)
                        {
                            unknownWs.Cells[unknownIndex, 1].Value = unknownDetail.Category;
                            unknownWs.Cells[unknownIndex, 2].Value = string.Empty;
                            unknownWs.Cells[unknownIndex, 3].Value = unknownDetail.Standard;
                            unknownWs.Cells[unknownIndex, 4].Value = unknownDetail.Profile;
                            unknownWs.Cells[unknownIndex, 5].Value = unknownDetail.Material;
                            unknownWs.Cells[unknownIndex, 6].Value = unknownDetail.Weight;

                            unknownIndex++;
                        }
                    }

                    package.Save();
                }
            }
            sheetData.Add("_SPEC_FILE_PATH_" + filePath, 0);
            return sheetData;
        }
    }
}
