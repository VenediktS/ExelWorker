using ExelWorker.Attributes;
using ExelWorker.Models;
using System.Collections.Generic;

namespace ExelWorker.App
{
    internal class ExelModelService
    {
        public ExelModelService()
        {

        }

        internal List<ExelPropertyModel> getClassProperties<TClass>(TClass ExelDataModel)
            where TClass : class
        {
            List<ExelPropertyModel> result = new List<ExelPropertyModel>();

            var properties = ExelDataModel.GetType().GetProperties();

            foreach (var property in properties)
            {
                object[] attrs = property.GetCustomAttributes(typeof(ExelColumnTitleAttribute), false);

                if (attrs.Length > 0)
                {
                    var attr = (ExelColumnTitleAttribute)attrs[0];
                    result.Add(new ExelPropertyModel
                    {
                        ModelPropertyName = property.Name,
                        ExelPropertyName = attr.Title
                    });
                }
            }

            return result;
        }
    }
}
