using System;
using System.Collections.Generic;
using System.Text;
using ReadExcel_SimpleVer.Model;
using Dapper;
using System.Data.SqlClient;
using System.Linq;

namespace ReadExcel_SimpleVer
{
    public class DB
    {
        private static string connectionString = @"data source = '.\SQLEXPRESS01'; initial catalog = 'Northwind'; User Id = 'test'; Password = '123456'; Max Pool Size = 300;";
        public static bool CreateTable(string sql)
        {
            bool result = false;
            using (var connection = new SqlConnection(connectionString))
            {
                int n = connection.Execute(sql);

                //var num = connection.Query("select * from tablename").ToList();
                //if (num.Count() > 0)
                //{
                //    result = true;
                //}
                result = true;
            }
            return result;
        }

        public static bool Insert<T>(string command, T model) where T : class
        {
            try
            {


                using (var connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    var ddd = connection.Execute(command, model);

                    return true;

                }

            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
                return false;
            }

        }

        public static List<Customer> GetCustomer(Customer Customer)
        {
            try
            {
                var sqlCommand = new StringBuilder();
                sqlCommand.Append(" Select * From " + Customer.TableName);
                sqlCommand.Append(" Where CUSTOMER_NO = @CUSTOMER_NO");

                using (var connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    return connection.Query<Customer>(sqlCommand.ToString(), Customer).ToList();

                }

            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
            return null;
        }

        internal static bool UpdateCustomer(Customer customer)
        {
            try
            {
                var sqlCommand = new StringBuilder();
                sqlCommand.Append("Update CU_CUSTOMER_MAIN SET  CUSTOMER_NO=@CUSTOMER_NO ,CUSTOMER_NAME=@CUSTOMER_NAME ,INVOICE_NAME=@INVOICE_NAME ,CREATE_USER=@CREATE_USER ,CREATE_DATE=@CREATE_DATE ,IMPORT_DATE=@IMPORT_DATE ,UPDATE_USER=@UPDATE_USER ,UPDATE_DATE=@UPDATE_DATE ");
                sqlCommand.Append("Where CUSTOMER_NO=@CUSTOMER_NO ");


                using (var connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    var res = connection.Execute(sqlCommand.ToString(), customer);
                    if (res != 0)
                    {
                        return true;
                    }

                }

            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
            return false;
        }

        private static object GetPropByName<T>(T tempModel, string name)
        {
            return tempModel.GetType().GetProperty(name).GetValue(tempModel);
        }

        private static void SetPropertyValue<T>(T model, string propName, object param)
        {
            var type = model.GetType();
            var prop = type.GetProperty(propName);
            prop.SetValue(model, param);
        }



    }
}
