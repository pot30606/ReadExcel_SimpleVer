using Newtonsoft.Json;
using ReadExcel_SimpleVer.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;


namespace ReadExcel_SimpleVer
{
    class Program
    {
        static void Main(string[] args)
        {
            var filePath_TableContent = AppDomain.CurrentDomain.BaseDirectory + @"測試資料err.xlsx";

            //Excel 匯入
            ReadExcel(filePath_TableContent);
        }

        private static void ReadExcel(string filePath)
        {
            List<string> errorMesages = new List<string>();
            try
            {

                List<Customer> listTables = new List<Customer>();

                //讀取Excel
                listTables = MyNPOIReader.ReadExcel<Customer>(filePath);

                //驗證資料
                listTables.Validate();

                //新增或修改資料
                listTables = InsertOrUpdateData(errorMesages, listTables);

                //紀錄Log
                WriteLog(listTables);

                #region 顯示錯誤

                DisplayErrors(errorMesages, listTables);

                Console.ReadLine();
                #endregion


            }
            catch (Exception ex)
            {
                Console.WriteLine("發生錯誤:" + ex.Message);
            }
            Console.WriteLine("完成");
        }


        private static List<Customer> InsertOrUpdateData(List<string> errorMesages, List<Customer> listTables)
        {
            foreach (var table in listTables)
            {
                // 其中一筆Model有格式不對
                if (table.IsValidModel != true)
                {
                    string errors = table.ValidationResults.Aggregate("", (current, next) => current + next.MemberNames.FirstOrDefault() + next.ErrorMessage + ", ");
                    errorMesages.Add($"第{table.row}列格式錯誤:{errors}");
                    continue;
                }
                //這裡要檢查是否有資料存在資料庫
                if (IsExist(table))
                {
                    //這裡要更新資料
                    bool updateResult = false;
                    if (updateResult is false)
                    {
                        errorMesages.Add($"第{table.row}列更新資料失敗");
                    }
                    else
                    {
                        table.IsSuccess = true;
                    }
                }
                else
                {
                    //這裡要新增資料
                    bool insertResult = false;
                    if (insertResult is false)
                    {
                        errorMesages.Add($"第{table.row}列新增資料失敗");
                    }
                    else
                    {
                        table.IsSuccess = true;
                    }
                }

            }
            return listTables;
        }

        private static void DisplayErrors(List<string> errorMesages, List<Customer> listTables)
        {
            var fail = listTables.Where((x) =>
            {
                return x.IsSuccess == false;
            });

            var success = listTables.Where((x) =>
            {
                return x.IsSuccess == true;
            });

            Console.WriteLine("總共失敗筆數:" + fail.Count());
            Console.WriteLine("總共成功筆數:" + success.Count());


            foreach (var errorItem in errorMesages)
            {
                Console.WriteLine(errorItem);
            }
        }


        #region 紀錄Log
        private static void WriteLog(object data)
        {
            StringBuilder sb = new StringBuilder();
            var guid = Guid.NewGuid().ToString();
            string strDatetime = DateTime.Now.ToString(" yyyy-MM-dd-HHmmss-ffff");
            string strData = JsonConvert.SerializeObject(data);
            sb.Append("Datetime:");
            sb.Append(strDatetime);
            sb.Append("\r\n");
            sb.Append("Data:");
            sb.Append(strData);
            string pathlolg = @"C:\Users\pot\Documents\sql作業\log-" + guid + strDatetime + ".txt";
            File.AppendAllText(pathlolg, sb.ToString());
            Console.WriteLine("write log" + pathlolg);
            sb.Clear();
        }
        #endregion

        private static bool IsExist(object table)
        {
            bool isExist = true;
            string className = table.GetType().Name;
            switch (className)
            {
                case "Customer":
                    if (DB.GetCustomer((Customer)table).Count() == 0)
                    {
                        isExist = false;
                    }
                    break;
              
                default:
                    Console.WriteLine("找不到相應的Model " + className);
                    break;
            }

            return isExist;
        }



    }
}
