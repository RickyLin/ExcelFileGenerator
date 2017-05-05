using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace RL.FileGenerators.Excel
{
    public static class ExcelFileGenerator
    {
        #region Create Excel File from entity collection

        public static MemoryStream CreateStream<T>(IEnumerable<T> entities, string sheetName)
        {
            Type entityType = typeof(T);
            IOrderedEnumerable<ExcelColumnProperty> excelColProperties = GetExcelColumnProperties(entityType);
            MemoryStream ms = new MemoryStream();
            CreateExcelStream(entities, excelColProperties, ms, sheetName);
            return ms;
        }

        public static void CreateFile<T>(IEnumerable<T> entities, string sheetName, string fileName)
        {
            Type entityType = typeof(T);
            IOrderedEnumerable<ExcelColumnProperty> excelColProperties = GetExcelColumnProperties(entityType);
            using (FileStream fs = File.Create(fileName))
            {
                CreateExcelStream(entities, excelColProperties, fs, sheetName);
            }
        }

        private static IOrderedEnumerable<ExcelColumnProperty> GetExcelColumnProperties(Type entityType)
        {
            List<ExcelColumnProperty> excelColumnProperties = new List<ExcelColumnProperty>();
            ExcelColumnAttribute customAttribute;
            foreach (PropertyInfo pi in entityType.GetTypeInfo().DeclaredProperties)
            {
                customAttribute = pi.GetCustomAttribute<ExcelColumnAttribute>(true);
                if (customAttribute == null)
                    continue;

                excelColumnProperties.Add(new ExcelColumnProperty
                {
                    PropertyInfo = pi,
                    ExcelColumnAttr = customAttribute
                });
            }
            IOrderedEnumerable<ExcelColumnProperty> excelColProperties = excelColumnProperties.OrderBy(p => p.ExcelColumnAttr.Index);
            return excelColProperties;
        }

        private static void CreateExcelStream<T>(IEnumerable<T> entities, IOrderedEnumerable<ExcelColumnProperty> excelColProperties, Stream excelStream, string sheetName)
        {
            SetDefaultNumberFormatId(excelColProperties);

            using (SpreadsheetDocument doc = SpreadsheetDocument.Create(excelStream, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = doc.AddWorkbookPart();

                SharedStringTablePart sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
                SharedStringTable sharedStrTbl = new SharedStringTable();
                sharedStringPart.SharedStringTable = sharedStrTbl;

                WorkbookStylesPart stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylePart.Stylesheet = CreateStylesheet(excelColProperties);
                stylePart.Stylesheet.Save();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                Worksheet worksheet = new Worksheet();
                worksheetPart.Worksheet = worksheet;

                SheetData sheetData = new SheetData();
                sheetData.Append(CreateHeader(excelColProperties, sharedStrTbl));
                sheetData.Append(CreateContent(entities, excelColProperties, sharedStrTbl));
                worksheet.Append(sheetData);
                worksheet.Save();

                Workbook workbook = new Workbook();
                workbookPart.Workbook = workbook;

                FileVersion version = new FileVersion
                {
                    ApplicationName = "Microsoft Office Excel"
                };
                workbook.Append(version);

                Sheets sheets = new Sheets();
                Sheet sheet = new Sheet
                {
                    Name = sheetName,
                    SheetId = 1,
                    Id = workbookPart.GetIdOfPart(worksheetPart)
                };
                sheets.Append(sheet);
                workbook.Append(sheets);

                sharedStrTbl.Save();
                workbook.Save();
                doc.Close();
            }

        }

        private static void SetDefaultNumberFormatId(IOrderedEnumerable<ExcelColumnProperty> excelColProperties)
        {
            foreach (ExcelColumnProperty p in excelColProperties)
            {
                if (p.ExcelColumnAttr.NumberFormatId == uint.MaxValue && string.IsNullOrEmpty(p.ExcelColumnAttr.NumberingFormatString))
                {
                    p.ExcelColumnAttr.NumberFormatId = 0;
                }
            }
        }

        private static Stylesheet CreateStylesheet(IOrderedEnumerable<ExcelColumnProperty> excelColProperties)
        {
            Stylesheet ss = new Stylesheet();

            // set default styles

            Fonts fonts = new Fonts();
            Font font = new Font()
            {
                FontName = new FontName() { Val = "Calibri" },
                FontSize = new FontSize() { Val = 11 }
            };
            fonts.Append(font);
            fonts.Count = (uint)fonts.ChildElements.Count;

            Fills fills = new Fills();
            Fill fill;
            fill = new Fill()
            {
                PatternFill = new PatternFill() { PatternType = PatternValues.None }
            };
            fills.Append(fill);
            fill = new Fill()
            {
                PatternFill = new PatternFill() { PatternType = PatternValues.Gray125 }
            };
            fills.Append(fill);
            fills.Count = (uint)fills.ChildElements.Count;

            Borders borders = new Borders();
            borders.Append(new Border()
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder(),
                BottomBorder = new BottomBorder()
            });
            borders.Count = (uint)borders.ChildElements.Count;

            CellStyles cellStyles = new CellStyles();
            cellStyles.Append(new CellStyle()
            {
                Name = "Normal",
                BuiltinId = 0,
                FormatId = 0
            });
            cellStyles.Count = (uint)cellStyles.ChildElements.Count;

            TableStyles tableStyles = new TableStyles()
            {
                Count = 0,
                DefaultTableStyle = "TableStyleMedium2",
                DefaultPivotStyle = "PivotStyleMedium9"
            };

            DifferentialFormats dxfs = new DifferentialFormats()
            {
                Count = 0
            };

            CellStyleFormats styleFormats = new CellStyleFormats();
            styleFormats.Append(new CellFormat()
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0
            });
            styleFormats.Count = (uint)styleFormats.ChildElements.Count;

            // set styles according to the entity setting.
            Dictionary<uint, int> embeddedStyle = new Dictionary<uint, int>();   // key: NumberFormatId, value: cell format index
            uint nbrFmtId;
            int cellFormatIndex = -1;
            uint customizedNumberFormatId = 164;
            NumberingFormats nbrFormats = new NumberingFormats();
            NumberingFormat nbrFmt;
            Dictionary<string, uint> nbrFmtIds = new Dictionary<string, uint>();
            CellFormats cellFormats = new CellFormats();
            CellFormat cellFmt;

            foreach (ExcelColumnProperty excelColProperty in excelColProperties)
            {
                // handle custom number format
                if (string.IsNullOrEmpty(excelColProperty.ExcelColumnAttr.NumberingFormatString) == false)
                {
                    // customized format
                    if (nbrFmtIds.ContainsKey(excelColProperty.ExcelColumnAttr.NumberingFormatString) == false)
                    {
                        nbrFmt = new NumberingFormat
                        {
                            NumberFormatId = customizedNumberFormatId++,
                            FormatCode = excelColProperty.ExcelColumnAttr.NumberingFormatString
                        };
                        nbrFormats.Append(nbrFmt);
                        nbrFmtIds.Add(nbrFmt.FormatCode, nbrFmt.NumberFormatId);
                    }

                    nbrFmtId = nbrFmtIds[excelColProperty.ExcelColumnAttr.NumberingFormatString];
                }
                else
                    nbrFmtId = excelColProperty.ExcelColumnAttr.NumberFormatId;

                if (embeddedStyle.ContainsKey(nbrFmtId) == false)
                {
                    cellFormatIndex++;
                    cellFmt = new CellFormat
                    {
                        NumberFormatId = nbrFmtId,
                        FontId = 0,
                        FillId = 0,
                        BorderId = 0
                    };
                    cellFormats.Append(cellFmt);
                    embeddedStyle.Add(nbrFmtId, cellFormatIndex);
                }

                excelColProperty.CellFormatIndex = embeddedStyle[nbrFmtId];
            }
            nbrFormats.Count = (uint)nbrFormats.ChildElements.Count;
            cellFormats.Count = (uint)cellFormats.ChildElements.Count;

            // the sequence of appending is significant
            if (nbrFormats.Count > 0)
                ss.Append(nbrFormats);
            ss.Append(fonts);
            ss.Append(fills);
            ss.Append(borders);
            ss.Append(styleFormats);
            ss.Append(cellFormats);
            ss.Append(cellStyles);
            ss.Append(dxfs);
            ss.Append(tableStyles);

            return ss;
        }

        private static Row[] CreateContent<T>(IEnumerable<T> entities, IOrderedEnumerable<ExcelColumnProperty> excelColProperties, SharedStringTable sharedStrTbl)
        {
            Row row;
            Cell cell;
            List<Row> rows = new List<Row>(entities.Count());
            uint rowIndex = 1;
            string colName;
            int colIndex;
            object cellValue;
            foreach (T entity in entities)
            {
                rowIndex++;
                row = new Row
                {
                    RowIndex = rowIndex,
                    Spans = new ListValue<StringValue>() { InnerText = "1:" + excelColProperties.Count().ToString() }
                };
                rows.Add(row);
                colIndex = 0;
                foreach (ExcelColumnProperty property in excelColProperties)
                {
                    colIndex++;
                    cellValue = property.PropertyInfo.GetValue(entity, null);
                    if (cellValue == null)
                        continue;
                    colName = GetExcelColumnName(colIndex);
                    cell = new Cell
                    {
                        CellReference = string.Format("{0}{1}", colName, rowIndex)
                    };
                    SetCellTypeAndValue(cell, property, cellValue, sharedStrTbl);
                    row.Append(cell);
                }
            }

            return rows.ToArray();
        }

        private static Row CreateHeader(IOrderedEnumerable<ExcelColumnProperty> excelColProperties, SharedStringTable sharedStrTbl)
        {
            Row header = new Row
            {
                RowIndex = 1
            };
            Cell headerCell;
            string colName;
            int colIndex = 0;
            int sharedStrIndex;
            foreach (ExcelColumnProperty property in excelColProperties)
            {
                colIndex++;
                colName = GetExcelColumnName(colIndex);
                headerCell = new Cell
                {
                    CellReference = string.Format("{0}{1}", colName, '1'),
                    DataType = CellValues.SharedString
                };
                sharedStrIndex = GetSharedStringIndex(sharedStrTbl, property.ExcelColumnAttr.Title);
                headerCell.CellValue = new CellValue(sharedStrIndex.ToString());
                header.Append(headerCell);
            }

            return header;
        }

        private static int GetSharedStringIndex(SharedStringTable sharedStrTbl, string str)
        {
            int i = 0;
            foreach (SharedStringItem item in sharedStrTbl.Elements<SharedStringItem>())
            {
                if (item.InnerText == str)
                    return i;
                i++;
            }
            sharedStrTbl.AppendChild(new SharedStringItem(new Text(str)));

            return i;
        }

        #endregion

        #region Create Excel file by Excel template

        /// <summary>
        /// Create Excel stream from an Excel template
        /// </summary>
        /// <typeparam name="T">the type of Entity</typeparam>
        /// <param name="entity">the entity where we get data from</param>
        /// <param name="excelTemplateFileName">the full path and name of the Excel template file</param>
        /// <returns></returns>
        public static MemoryStream CreateStreamFromTemplate<T>(T entity, string excelTemplateFileName)
        {
            // copy the template file to memory, so changes won't affect the template file.
            MemoryStream ms;
            using (FileStream fs = File.Open(excelTemplateFileName, FileMode.Open))
            {
                ms = new MemoryStream((int)fs.Length);
                fs.CopyTo(ms);
            }

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(ms, true))
            {
                WorkbookPart wbPart = doc.WorkbookPart;
                List<DefinedNameData> definedNameDataSet = GetDefinedNames(wbPart);
                WorksheetPart wsPart = null;
                string previousSheetName = null;
                IEnumerable<PropertyInfo> entityProperties = entity.GetType().GetTypeInfo().DeclaredProperties;
                foreach (DefinedNameData dn in definedNameDataSet)
                {
                    if (previousSheetName != dn.SheetName)
                    {
                        wsPart = GetWorksheetPart(wbPart, dn.SheetName);
                        previousSheetName = dn.SheetName;
                    }

                    object cellValue = GetEntityPropertyValue(entity, entityProperties, dn.Name);
                    Cell cell = GetCell(wsPart, dn, cellValue);
                    SetCellTypeAndValue(cell, cellValue, wbPart.SharedStringTablePart.SharedStringTable);
                }
                doc.Close();
            }

            return ms;
        }

        public static void CreateFileFromTemplate<T>(T entity, string sheetName, string fileName)
        {
            //Type entityType = typeof(T);
            //IOrderedEnumerable<ExcelColumnProperty> excelColProperties = GetExcelColumnProperties(entityType);
            //using (FileStream fs = File.Create(fileName))
            //{
            //	CreateExcelStream(entity, excelColProperties, fs, sheetName);
            //	fs.Close();
            //}
            throw new NotImplementedException();
        }

        private static List<DefinedNameData> GetDefinedNames(WorkbookPart wbPart)
        {
            DefinedNames definedNames = wbPart.Workbook.DefinedNames;
            List<DefinedNameData> definedNameDataSet = new List<DefinedNameData>(definedNames.Count());
            DefinedNameData definedNameData;
            string sheetName, columnName;
            uint rowIndex;
            string[] referenceSegments;
            foreach (DefinedName dn in definedNames)
            {
                // assume none of these defined names are cell range (e.g. "A1", not "A1:B1").
                referenceSegments = dn.Text.Split('!');
                sheetName = referenceSegments[0].Trim('\'');
                referenceSegments = referenceSegments[1].Split('$');
                columnName = referenceSegments[1];
                rowIndex = uint.Parse(referenceSegments[2]);
                definedNameData = new DefinedNameData(dn.Name, sheetName, columnName, rowIndex);
                definedNameDataSet.Add(definedNameData);
            }

            return definedNameDataSet;
        }

        private static WorksheetPart GetWorksheetPart(WorkbookPart wbPart, string sheetName)
        {
            Sheet sheet = wbPart.Workbook.Descendants<Sheet>().First(s => s.Name == sheetName);
            return wbPart.GetPartById(sheet.Id) as WorksheetPart;
        }

        private static Cell GetCell(WorksheetPart wsPart, DefinedNameData dn, object cellValue)
        {
            Cell cell = wsPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference == dn.Reference);
            if (cell != null)
                return cell;

            Worksheet ws = wsPart.Worksheet;
            SheetData sheetData = ws.GetFirstChild<SheetData>();

            Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == dn.RowIndex);
            if (row == null)
            {
                row = new Row() { RowIndex = dn.RowIndex };
                sheetData.Append(row);
            }

            Type cellValueType = cellValue == null ? typeof(int) : cellValue.GetType();
            cell = new Cell() { CellReference = dn.Reference, DataType = GetCellDataType(cellValueType) };

            Cell refCell = null;
            IEnumerable<Cell> existingCells = row.Elements<Cell>();
            foreach (Cell cellInRow in existingCells)
            {
                if (string.Compare(cellInRow.CellReference.Value, dn.Reference) > 0)
                {
                    refCell = cellInRow;
                    break;
                }
            }

            if (refCell != null)
            {
                row.InsertBefore(cell, refCell);
            }
            else if (existingCells.Count() > 0)
            {
                row.InsertAfter(cell, existingCells.Last());
            }
            else
            {
                row.InsertAt(cell, 0);
            }

            return cell;
        }

        private static object GetEntityPropertyValue<T>(T entity, IEnumerable<PropertyInfo> entityProperties, string propertyName)
        {
            PropertyInfo entityProperty = entityProperties.FirstOrDefault(pi => pi.Name.Equals(propertyName, StringComparison.CurrentCultureIgnoreCase));
            if (entityProperty == null)
                return null;

            return entityProperty.GetValue(entity);
        }

        #endregion

        #region Utilities

        private static EnumValue<CellValues> GetCellDataType(PropertyInfo pi)
        {
            Type propertyType = pi.PropertyType;
            return GetCellDataType(propertyType);
        }

        private static EnumValue<CellValues> GetCellDataType(Type typeOfValue)
        {
            if (typeOfValue == typeof(string))
            {
                return CellValues.SharedString;
            }
            if (typeOfValue == typeof(bool) || typeOfValue == typeof(Nullable<bool>))
            {
                return CellValues.Boolean;
            }

            return CellValues.Number;
        }

        private static void SetCellTypeAndValue(Cell cell, object value, SharedStringTable sharedStrTbl)
        {
            SetCellTypeAndValue(cell, null, value, sharedStrTbl);
        }

        private static void SetCellTypeAndValue(Cell cell, ExcelColumnProperty property, object value, SharedStringTable sharedStrTbl)
        {
            if (property != null)
            {
                cell.DataType = GetCellDataType(property.PropertyInfo);
                if (property.CellFormatIndex > 0)
                    cell.StyleIndex = (uint)property.CellFormatIndex;
            }

            if (value == null)
                return;

            string strCellValue;
            int sharedStrIndex;

            if (value is DateTime)
            {
                strCellValue = ((DateTime)value).ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
            else if (cell.DataType == null)
            {
                strCellValue = value.ToString();
            }
            else if (cell.DataType == CellValues.SharedString)
            {
                string theValue;
                if (value is bool)
                {
                    theValue = (bool)value ? "Yes" : "No";
                }
                else
                    theValue = value.ToString();
                sharedStrIndex = GetSharedStringIndex(sharedStrTbl, theValue);
                strCellValue = sharedStrIndex.ToString();
            }
            else if (cell.DataType.Value == CellValues.Boolean)
            {
                strCellValue = BooleanValue.FromBoolean((bool)value).ToString();
            }
            else
                strCellValue = value.ToString();

            cell.CellValue = new CellValue(strCellValue);
        }

        // this method is from http://stackoverflow.com/questions/181596/how-to-convert-a-column-number-eg-127-into-an-excel-column-eg-aa
        // the columnNumber starts from 1
        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        #endregion
    }
}
