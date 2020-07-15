using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using ExelWorker.Models;
using ExelWorker.App;
using ExelWorker.ExelReader;

namespace ExelWorker.ExelWorker
{
    public class ExelWorker : IExelWorker
    {
        public ExelWorker()
        {

        }

        public List<TModel> GetModelFromExel<TModel>(FileStream fileStream)
            where TModel : class
        {
            if (Path.GetExtension(fileStream.Name) != ".xlsx")
            {
                throw new FileFormatException("Incorrect file format, work is possible only with .xlsx");
            }

            ExelModelService exelModelService = new ExelModelService();

            var model = Activator.CreateInstance(typeof(TModel));

            List<ExelPropertyModel> exelPropertyModels;

            exelPropertyModels = exelModelService.getClassProperties<TModel>((TModel)model);

            XLSXReader reader = new XLSXReader(exelPropertyModels);

            var bookValues = reader.ReadAllCellValues(fileStream);

            return MapBookValueToModel<TModel>(bookValues, exelPropertyModels);
        }

        private List<TModel> MapBookValueToModel<TModel>(Stack<Dictionary<string, string>> bookValues, List<ExelPropertyModel> exelPropertyModels)
            where TModel : class
        {
            var resultList = new List<TModel>();

            foreach (var row in bookValues)
            {
                var valModel = Activator.CreateInstance(typeof(TModel));

                foreach (var value in row)
                {
                    var property = exelPropertyModels.FirstOrDefault(prop => value.Key == prop.ExelPropertyId);
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
                if (value.GetType() == prop.PropertyType)
                {
                    prop.SetValue(obj, Convert.ChangeType(value, prop.PropertyType, null));
                }
            }
        }    
    }
}
