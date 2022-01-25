using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ReadExcel_SimpleVer
{
    public static class Validation
    {

        public static void Validate<T>(this List<T> listModel)
        {
            foreach (var item in listModel)
            {
                SetPropertyValue(item, "CREATE_USER", "UserA");
                SetPropertyValue(item, "CREATE_DATE", DateTime.Now);
                SetPropertyValue(item, "IMPORT_DATE", DateTime.Now);

                var context = new ValidationContext(item, serviceProvider: null, items: null);
                var validationResults = new List<ValidationResult>();
                bool isValid = Validator.TryValidateObject(item, context, validationResults, true);


                SetPropertyValue(item, "IsValidModel", isValid);
                SetPropertyValue(item, "ValidationResults", validationResults);

            }
        }

        private static void SetPropertyValue<T>(T model, string propName, object param)
        {
            var type = model.GetType();
            var prop = type.GetProperty(propName);
            prop.SetValue(model, param);
        }
        private static object GetPropByName<T>(T tempModel, string name)
        {
            return tempModel.GetType().GetProperty(name).GetValue(tempModel);
        }

        private static object GetValueByName<T>(T tempModel, string name)
        {
            return false;
        }
    }
}
