using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using ExelWorker.Models;
using ExelWorker.App;

namespace ExelWorker.ExelWorker
{
    public class ExelWorker : IExelWorker
    {
        private List<ExelPropertyModel> exelProperyModel;
        public ExelWorker()
        {

        }

        public List<TModel> GetModelFromExel<TModel>(FileStream fileStream)
            where TModel : class
        {
            ExelModelService exelModelService = new ExelModelService();

            var model = Activator.CreateInstance(typeof(TModel));

            exelProperyModel = exelModelService.getClassProperties<TModel>((TModel)model);

            var bookValues = ReadAllCellValues(fileStream);

            return MapBookValueToModel<TModel>(bookValues);
        }

        private List<Dictionary<string, string>> ReadAllCellValues(Stream fileStream)
        {
            var resultList = new List<Dictionary<string, string>>();

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
                                resultList.Add(row);
                            }

                            indexOfRow++;
                        }
                    }
                }
            }

            return resultList;
        }

        private void SetIdsOfExelPropertyToModel(Dictionary<string, string> rowOfTitle)
        {
            foreach (var row in rowOfTitle)
            {
                exelProperyModel.FirstOrDefault(item => item.ExelPropertyName == row.Value).ExelPropertyId = row.Key;
            }
        }

        private List<TModel> MapBookValueToModel<TModel>(List<Dictionary<string, string>> bookValues)
            where TModel : class
        {
            var resultList = new List<TModel>();

            foreach (var row in bookValues)
            {
                var valModel = Activator.CreateInstance(typeof(TModel));

                foreach (var value in row)
                {
                    var property = exelProperyModel.FirstOrDefault(prop => value.Key == prop.ExelPropertyId);
                    if (property != null)
                    {
                        SetProperty(valModel, property.ModelPropertyName, value.Value);
                    }
                }

                resultList.Add((TModel)valModel);
                valModel = null;
            }

            return resultList;
        }

        private void SetProperty(object obj, string property, object value)
        {
            var prop = obj.GetType().GetProperty(property, BindingFlags.Public | BindingFlags.Instance);
            if (prop != null && prop.CanWrite)
            {
                try
                {
                    prop.SetValue(obj, Convert.ChangeType(value, prop.PropertyType, null));
                }
                catch (Exception e) 
                {
                }
            }
        }

        private Dictionary<string, string> ReadSingleRow(OpenXmlReader reader, WorkbookPart workbookPart)
        {
            var resultDictionary = new Dictionary<string, string>();
            int index = 1;


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

                    ++index;

                }
            } while (reader.ReadNextSibling());

            return resultDictionary;
        }
    }
}
