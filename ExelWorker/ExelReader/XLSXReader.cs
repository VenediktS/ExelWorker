using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ExelWorker.Models;

namespace ExelWorker.ExelReader
{
    internal class XLSXReader
    {
        private List<ExelPropertyModel> _exelProperyModel;

        internal XLSXReader(List<ExelPropertyModel> exelPropertyModel) 
        {
            _exelProperyModel = exelPropertyModel;
        }
        internal Stack<Dictionary<string, string>> ReadAllCellValues(Stream fileStream)
        {
            var resultList = new Stack<Dictionary<string, string>>();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileStream, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;

                foreach (WorksheetPart worksheetPart in workbookPart.WorksheetParts)
                {
                    OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);

                    int indexOfRow = 1;

                    while (reader.Read())
                    {
                        if (reader.ElementType == typeof(Row))
                        {
                            // Если это линия с заголовками

                            var row = ReadSingleRow(reader, workbookPart);
                            if (indexOfRow == 1)
                            {
                                SetIdsOfExelPropertyToModel(row);
                            }
                            else
                            {
                                resultList.Push(row);
                            }

                            indexOfRow++;
                        }
                    }
                }
            }

            return resultList;
        }

        private Dictionary<string, string> ReadSingleRow(OpenXmlReader reader, WorkbookPart workbookPart)
        {
            var resultDictionary = new Dictionary<string, string>();

            reader.ReadFirstChild();
            do
            {
                if (reader.ElementType == typeof(Cell))
                {

                    Cell c = (Cell)reader.LoadCurrentElement();

                    string cellValue;

                    #region Read cell value
                    if (c.DataType != null && c.DataType == CellValues.SharedString)
                    {
                        SharedStringItem ssi = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(c.CellValue.InnerText));

                        cellValue = ssi.Text.Text;
                    }
                    else
                    {
                        if (c.CellValue != null)
                        {
                            cellValue = c.CellValue.InnerText;
                        }
                        else
                        {
                            cellValue = "";
                        }

                    }
                    #endregion

                    resultDictionary.Add(Regex.Replace(c.CellReference, @"[\d]", string.Empty), cellValue);

                }
            } while (reader.ReadNextSibling());

            return resultDictionary;
        }

        private void SetIdsOfExelPropertyToModel(Dictionary<string, string> rowOfTitle)
        {
            foreach (var row in rowOfTitle)
            {
                _exelProperyModel.FirstOrDefault(item => item.ExelPropertyName == row.Value).ExelPropertyId = row.Key;
            }
        }
    }
}
