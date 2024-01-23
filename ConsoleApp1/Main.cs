using System;
using MySql.Data.MySqlClient;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System.Collections.Generic;


class ConsoleApp1
{
    static void Main()
    {
        List<Records> dataArray = new List<Records>();

        string connectionString = "Server=localhost;Database=mydb;User ID=root;Password=jimmy2052;";
        MySqlConnection connection = new MySqlConnection(connectionString);
            try
            {

                connection.Open();

                string query = "select DISTINCT accountCode, accountName,documentDate,srcDocNO, description , debitAmount, creditAmount from mytable;";

                using (MySqlCommand command = new MySqlCommand(query, connection))
                {
                    MySqlDataReader reader = command.ExecuteReader();

                    if (reader.HasRows)
                    {

                        readData(reader, dataArray);
                    }
                    else
                    {
                        Console.WriteLine("No data found.");
                    }

                    reader.Dispose();
                }

            } catch (Exception ex) {
                Console.WriteLine($"Error: {ex.Message}");
            }
            connection.Dispose();
        
        createExcel(dataArray);
        Console.WriteLine("Finished");
        Console.ReadKey();
    }    

    static void readData(MySqlDataReader reader, List<Records> dataArray)
    {
        while (reader.Read())
        {
            string column1Value = reader.GetString(0);
            string column2Value = reader.GetString(1);
            string column3Value = reader.GetString(2);
            string column4Value = reader.GetString(3);
            string column5Value = reader.GetString(4);
            double column6Value = reader.GetDouble(5);
            double column7Value = reader.GetDouble(6);

            Records data = new Records
            {
                accountCode = column1Value,
                accountName = column2Value,
                documentDate = column3Value,
                documentTitle = column4Value,
                Description = column5Value,
                debitAmount = column6Value,
                creditAmount = column7Value
            };


            dataArray.Add(data);
        }
    }
    static void createExcel(List<Records> dataArray)
    {
        string filePath = "C:\\Users\\JIMMY\\OneDrive\\Desktop\\ExcelFile.xlsx";
        var spreadsheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
        
            /** BOILERPLATE**/
            // Add a workbook part
            var workbookPart = spreadsheetDocument.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            // Add a worksheet part
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            worksheetPart.Worksheet = new Worksheet(sheetData);

            // Add a sheets element to the workbook
            var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
            var sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Ledger" };
            sheets.Append(sheet);
            WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            workbookStylesPart.Stylesheet = GenerateStylesheet();
            /** BOILERPLATE**/

            /** HEADER ROW **/
            var headerRow = new Row();
            string[] headers = { "#", "Doc.Date", "Doc.NO", "Description", "Debit", "Credit", "Balance" };

            foreach (var header in headers)
            {
                var cell = new Cell(new InlineString(new Text(header))) { DataType = CellValues.InlineString };
                headerRow.AppendChild(cell);
            }

            sheetData.AppendChild(headerRow);
            /** HEADER ROW **/

            var emptyRow = new Row();
            var emptyCell = new Cell(new InlineString(new Text(""))) { DataType = CellValues.InlineString };
            emptyRow.AppendChild(emptyCell);
            sheetData.AppendChild(emptyRow);

            var mergeCells = worksheetPart.Worksheet.Elements<MergeCells>().FirstOrDefault();
            if (mergeCells == null)
            {
                mergeCells = new MergeCells();
                worksheetPart.Worksheet.AppendChild(mergeCells);
            }

            int num = 0;
            for (int numRecord = 0; numRecord < dataArray.Count; numRecord++)
            {
                int groupingNo = 0;
                /** GROUPING  **/
                if (numRecord + 1 < dataArray.Count)
                {
                    
                    if (dataArray[numRecord].accountCode == dataArray[numRecord + 1].accountCode)
                    {
                        for (int i = numRecord + 1; i < dataArray.Count; i++)
                        {
                            if (dataArray[numRecord].accountCode == dataArray[i].accountCode)
                            {
                                groupingNo = i;
                            }
                        }

                        var rowGroupingTitle = new Row();
                        var mergeGroupCell = new MergeCell() { Reference = new StringValue("A" + (sheetData.Elements<Row>().Count() + 1) + ":H" + (sheetData.Elements<Row>().Count() + 1)) };
                        mergeCells.Append(mergeGroupCell);
                        var groupingCell = new Cell(new InlineString(new Text(dataArray[numRecord].accountCode + " - " + dataArray[numRecord].accountName))) { DataType = CellValues.InlineString };
                        rowGroupingTitle.AppendChild(groupingCell);
                        sheetData.AppendChild(rowGroupingTitle);

                        double groupingBalance = 0;
                        double groupingDebit = 0;
                        double groupingCredit = 0;
                        bool firstTime = false;
                        for (int i = numRecord; i <= groupingNo; i++)
                        {
                            num++; // increase the count
                            var rowGroupingChild = new Row();
                            var groupingChildCell = new Cell(new CellValue(num)) { DataType = CellValues.Number, CellReference = "A" + (sheetData.Elements<Row>().Count() + 1) };
                            rowGroupingChild.AppendChild(groupingChildCell);
                            var groupingChildCell2 = new Cell(new InlineString(new Text(dataArray[i].documentDate))) { DataType = CellValues.InlineString };
                            rowGroupingChild.AppendChild(groupingChildCell2);
                            var groupingChildCell3 = new Cell(new InlineString(new Text(dataArray[i].documentTitle))) { DataType = CellValues.InlineString };
                            rowGroupingChild.AppendChild(groupingChildCell3);
                            var groupingChildCell4 = new Cell(new InlineString(new Text(dataArray[i].Description))) { DataType = CellValues.InlineString };
                            rowGroupingChild.AppendChild(groupingChildCell4);
                            // skip calculate 1st row
                            if(i != numRecord)
                            {
                                groupingCredit += dataArray[i].creditAmount;
                                groupingDebit += dataArray[i].debitAmount;
                            }                          
                            groupingBalance += dataArray[i].debitAmount;
                            groupingBalance -= dataArray[i].creditAmount;
                            
                            
                            if (firstTime == false)
                            {
                                var groupingChildCell5 = new Cell(new CellValue("0")) { DataType = CellValues.Number, StyleIndex = 3 };
                                rowGroupingChild.AppendChild(groupingChildCell5);
                                var groupingChildCell6 = new Cell(new CellValue("0")) { DataType = CellValues.Number, StyleIndex = 3 };
                                rowGroupingChild.AppendChild(groupingChildCell6);
                                firstTime = true;
                            }
                            else
                            {
                                var groupingChildCell5 = new Cell(new CellValue(dataArray[i].debitAmount)) { DataType = CellValues.Number, StyleIndex = 3 };
                                rowGroupingChild.AppendChild(groupingChildCell5);
                                var groupingChildCell6 = new Cell(new CellValue(dataArray[i].creditAmount)) { DataType = CellValues.Number, StyleIndex = 3 };
                                rowGroupingChild.AppendChild(groupingChildCell6);
                            }

                            var groupingChildCell7 = new Cell(new CellValue(groupingBalance)) { DataType = CellValues.Number, StyleIndex = 3 };
                            rowGroupingChild.AppendChild(groupingChildCell7);
                            sheetData.AppendChild(rowGroupingChild);
                        }
                        
                        var rowGroupingChild3 = new Row();
                        //var rowGroupingChild4 = new Row();
                        var rowGroupingChild5 = new Row();
                        for (int i = 0; i < 4; i++)
                        {
                            var cell = new Cell(new InlineString(new Text(""))) { DataType = CellValues.InlineString };
                            rowGroupingChild3.AppendChild(cell);
                        }

                        var gr3c5 = new Cell(new CellValue(groupingDebit)) { DataType = CellValues.Number, StyleIndex = 3 };
                        rowGroupingChild3.AppendChild(gr3c5);
                        var gr3c6 = new Cell(new CellValue(groupingCredit)) { DataType = CellValues.Number, StyleIndex = 3 };
                        rowGroupingChild3.AppendChild(gr3c6);
                        var gr3c7= new Cell(new CellValue(groupingBalance)) { DataType = CellValues.Number, StyleIndex = 3 };
                        rowGroupingChild3.AppendChild(gr3c7);
                        sheetData.AppendChild(rowGroupingChild3);

                        //for (int i = 0; i < 4; i++)
                        //{
                        //    var cell = new Cell(new InlineString(new Text(""))) { DataType = CellValues.InlineString };
                        //    rowGroupingChild4.AppendChild(cell);
                        //}

                        //var gr4c5 = new Cell(new CellValue(groupingDebit)) { DataType = CellValues.Number };
                        //rowGroupingChild4.AppendChild(gr4c5);
                        //var gr4c6 = new Cell(new CellValue(groupingCredit)) { DataType = CellValues.Number };
                        //rowGroupingChild4.AppendChild(gr4c6);
                        //var gr4c7 = new Cell(new CellValue(groupingBalance)) { DataType = CellValues.Number };
                        //rowGroupingChild4.AppendChild(gr4c7);
                        //sheetData.AppendChild(rowGroupingChild4);
                        sheetData.AppendChild(rowGroupingChild5);

                        numRecord += (groupingNo - numRecord);
                        continue;
                    }
                }
                /** Finished GROUPING  **/
                num++;

                // Create new rows for each iteration
                var row1 = new Row();
                var row2 = new Row();
                var row3 = new Row();
                //var row4 = new Row();
                var row5 = new Row();

                // Add cells for the first two columns and merge them from A to H
                var mergeCell = new MergeCell() { Reference = new StringValue("A" + (sheetData.Elements<Row>().Count() + 1) + ":H" + (sheetData.Elements<Row>().Count() + 1)) };
                mergeCells.Append(mergeCell);

                var r1c1 = new Cell(new InlineString(new Text(dataArray[numRecord].accountCode + " - " + dataArray[numRecord].accountName))) { DataType = CellValues.InlineString };
                row1.AppendChild(r1c1);
                sheetData.AppendChild(row1);

                // Add cell for the sequential number
                var r2c1 = new Cell(new CellValue(num.ToString())) { DataType = CellValues.Number, CellReference = "A" + (sheetData.Elements<Row>().Count() + 1) };
                row2.AppendChild(r2c1);

                var r2c2 = new Cell(new InlineString(new Text(dataArray[numRecord].documentDate))) { DataType = CellValues.InlineString };
                row2.AppendChild(r2c2);

                var r2c3 = new Cell(new InlineString(new Text(dataArray[numRecord].documentTitle))) { DataType = CellValues.InlineString };
                row2.AppendChild(r2c3);

                var r2c4 = new Cell(new InlineString(new Text(dataArray[numRecord].Description))) { DataType = CellValues.InlineString };
                row2.AppendChild(r2c4);


                double balance = 0;

                balance += dataArray[numRecord].debitAmount;
                balance -= dataArray[numRecord].creditAmount;



                var r2c5 = new Cell(new CellValue("0")) { DataType = CellValues.Number, StyleIndex = 3 };
                row2.AppendChild(r2c5);
                var r2c6 = new Cell(new CellValue("0")) { DataType = CellValues.Number , StyleIndex = 3 };
                row2.AppendChild(r2c6);

                var r2c7 = new Cell(new CellValue(balance)) { DataType = CellValues.Number, StyleIndex = 3 };
                row2.AppendChild(r2c7);
                sheetData.AppendChild(row2);

                for(int i = 0; i < 4; i++)
                {
                    var cell = new Cell(new InlineString(new Text(""))) { DataType = CellValues.InlineString };
                    row3.AppendChild(cell);
                }

                var r3c5 = new Cell(new CellValue("0")) { DataType = CellValues.Number, StyleIndex = 3 };
                row3.AppendChild(r3c5);
                var r3c6 = new Cell(new CellValue("0")) { DataType = CellValues.Number, StyleIndex = 3 };
                row3.AppendChild(r3c6);
                
                var r3c7 = new Cell(new CellValue(balance)) { DataType = CellValues.Number , StyleIndex = 3 };
                row3.AppendChild(r3c7);
                sheetData.AppendChild(row3);


                //for (int i = 0; i < 4; i++)
                //{
                //    var cell = new Cell(new InlineString(new Text(""))) { DataType = CellValues.InlineString };
                //    row4.AppendChild(cell);
                //}

                //var r4c5 = new Cell(new CellValue("0")) { DataType = CellValues.Number };
                //row4.AppendChild(r4c5);
                //var r4c6 = new Cell(new CellValue("0")) { DataType = CellValues.Number };
                //row4.AppendChild(r4c6);
                //var r4c7 = new Cell(new CellValue(balance)) { DataType = CellValues.Number };
                //row4.AppendChild(r4c7);
                //sheetData.AppendChild(row4);
                sheetData.AppendChild(row5);

            }
        workbookStylesPart.Stylesheet.Save();
        spreadsheetDocument.Save();
            spreadsheetDocument.Dispose();

        Console.WriteLine($"Excel file created at: {filePath}");
        
    }

    static Stylesheet GenerateStylesheet()
    {
        Stylesheet ss = new Stylesheet();

        Fonts fts = new Fonts();
        DocumentFormat.OpenXml.Spreadsheet.Font ft = new DocumentFormat.OpenXml.Spreadsheet.Font();
        FontName ftn = new FontName();
        ftn.Val = "Calibri";
        FontSize ftsz = new FontSize();
        ftsz.Val = 11;
        ft.FontName = ftn;
        ft.FontSize = ftsz;
        fts.Append(ft);
        fts.Count = (uint)fts.ChildElements.Count;

        Fills fills = new Fills();
        Fill fill;
        PatternFill patternFill;
        fill = new Fill();
        patternFill = new PatternFill();
        patternFill.PatternType = PatternValues.None;
        fill.PatternFill = patternFill;
        fills.Append(fill);
        fill = new Fill();
        patternFill = new PatternFill();
        patternFill.PatternType = PatternValues.Gray125;
        fill.PatternFill = patternFill;
        fills.Append(fill);
        fills.Count = (uint)fills.ChildElements.Count;

        Borders borders = new Borders();
        Border border = new Border();
        border.LeftBorder = new LeftBorder();
        border.RightBorder = new RightBorder();
        border.TopBorder = new TopBorder();
        border.BottomBorder = new BottomBorder();
        border.DiagonalBorder = new DiagonalBorder();
        borders.Append(border);
        borders.Count = (uint)borders.ChildElements.Count;

        CellStyleFormats csfs = new CellStyleFormats();
        CellFormat cf = new CellFormat();
        cf.NumberFormatId = 0;
        cf.FontId = 0;
        cf.FillId = 0;
        cf.BorderId = 0;
        csfs.Append(cf);
        csfs.Count = (uint)csfs.ChildElements.Count;

        uint iExcelIndex = 164;
        NumberingFormats nfs = new NumberingFormats();
        CellFormats cfs = new CellFormats();

        cf = new CellFormat();
        cf.NumberFormatId = 0;
        cf.FontId = 0;
        cf.FillId = 0;
        cf.BorderId = 0;
        cf.FormatId = 0;
        cfs.Append(cf);

        NumberingFormat nf;


        // #,##0.00 is also Excel style index 4
        nf = new NumberingFormat();
        nf.NumberFormatId = iExcelIndex++;
        nf.FormatCode = "#,##0.00";
        nfs.Append(nf);
        cf = new CellFormat();
        cf.NumberFormatId = nf.NumberFormatId;
        cf.FontId = 0;
        cf.FillId = 0;
        cf.BorderId = 0;
        cf.FormatId = 0;
        cf.ApplyNumberFormat = true;
        cfs.Append(cf);

        // @ is also Excel style index 49
        nf = new NumberingFormat();
        nf.NumberFormatId = iExcelIndex++;
        nf.FormatCode = "@";
        nfs.Append(nf);
        cf = new CellFormat();
        cf.NumberFormatId = nf.NumberFormatId;
        cf.FontId = 0;
        cf.FillId = 0;
        cf.BorderId = 0;
        cf.FormatId = 0;
        cf.ApplyNumberFormat = true;
        cfs.Append(cf);

        nfs.Count = (uint)nfs.ChildElements.Count;
        cfs.Count = (uint)cfs.ChildElements.Count;

        ss.Append(nfs);
        ss.Append(fts);
        ss.Append(fills);
        ss.Append(borders);
        ss.Append(csfs);
        ss.Append(cfs);

        CellStyles css = new CellStyles();
        CellStyle cs = new CellStyle();
        cs.Name = "Normal";
        cs.FormatId = 0;
        cs.BuiltinId = 0;
        css.Append(cs);
        css.Count = (uint)css.ChildElements.Count;
        ss.Append(css);

        DifferentialFormats dfs = new DifferentialFormats();
        dfs.Count = 0;
        ss.Append(dfs);

        TableStyles tss = new TableStyles();
        tss.Count = 0;
        tss.DefaultTableStyle = "TableStyleMedium9";
        tss.DefaultPivotStyle = "PivotStyleLight16";
        ss.Append(tss);

        return ss;
    }

}

