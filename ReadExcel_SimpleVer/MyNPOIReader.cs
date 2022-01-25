using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ReadExcel_SimpleVer
{
    public class MyNPOIReader
    {
        public static List<T> ReadExcel<T>(string filePath) where T : class, new()
        {
            int HeaderRow = 1;
            int StartFrom_Row = 2;
            int StartFrom_Col = 0;


            List<T> listModel = new List<T>();
            ISheet sheet;
            using (var stream = new FileStream(filePath, FileMode.Open))
            {
                stream.Position = 0;
                XSSFWorkbook xssWorkbook = new XSSFWorkbook(stream);
                //取得Model Property :　SheetIndex　位置
                var sheetIndex = (int)GetPropByName(new T(), "SheetIndex");
                //取得指定位置的試算表sheet
                sheet = xssWorkbook.GetSheetAt(sheetIndex);
                //取得Header的一整個Row
                IRow headerRow = sheet.GetRow(HeaderRow);
                //Header Row長度
                var lastCellNum = headerRow.LastCellNum;

                //每個Row
                for (int i = StartFrom_Row; i <= sheet.LastRowNum; i++)
                {
                    //實體化model
                    var model = new T();
                    //取得model的type
                    var type = model.GetType();
                    //讀取第i個Row資料
                    IRow row = sheet.GetRow(i);
                    //紀錄現在是第幾列 要顯示給user看哪一列有錯所以+1
                    SetPropertyValue(model, "row", i + 1);
                    //若該Row都是Blank或沒有Cell 則跳過這一列
                    if (row.Cells.All(d => d.CellType == CellType.Blank) || row.Cells.Count() <= 0)
                    {
                        continue;
                    };

                    //讀取該Row的每個Column
                    for (var j = 0 + StartFrom_Col; j < lastCellNum; j++)
                    {
                        //若該Column是空值或Blank則跳過
                        if (row.GetCell(j) is null || row.GetCell(j).CellType == CellType.Blank)
                            continue;
                        //取得Model的Property 找到該Column的Header欄位的字串(如該欄位是:CUSTOMER_NO => headerRow.GetCell(j).StringCellValue 就會是CUSTOMER_NO) 去取得Property名稱
                        var property = type.GetProperty(headerRow.GetCell(j).StringCellValue);
                        if (property is null)
                            continue;
                        //取得Cell的值
                        var cellValue = row.GetCell(j)?.ToString();

                        //若Type是bool需要做轉換: 1=True , 0=False
                        Type paramType = property.PropertyType;
                        if (paramType == typeof(bool))
                        {
                            cellValue = ConvertBooleanToStringValue(cellValue);
                        }

                        //若Cell值不為空 則賦予給Property
                        if (cellValue != null && cellValue.Length > 0)
                        {
                            //將值轉換為Property的型態
                            var value = ConvertType(cellValue, paramType);
                            if (value != null)
                            {
                                //設定值給Property
                                property.SetValue(model, value);
                            }
                        }
                    }
                    listModel.Add(model);
                }
            }
            return listModel;
        }

        /// <summary>
        /// 將字串(0或1)轉換為bool字串
        /// </summary>
        /// <param name="cellValue">字串(0或1)</param>
        /// <returns>字串"true"或"false"</returns>
        private static string ConvertBooleanToStringValue(string cellValue)
        {
            int myVal = 0;
            if (int.TryParse(cellValue, out myVal))
            {
                if (myVal == 1)
                {
                    cellValue = "true";
                }
                else
                {
                    cellValue = "false";
                }
            }
            else
            {
                cellValue = "false";
            }

            return cellValue;
        }

        /// <summary>
        /// 嘗試將物件轉換為指定的型別
        /// </summary>
        /// <param name="cellValue">物件</param>
        /// <param name="type">型別</param>
        /// <returns></returns>
        public static object ConvertType(object cellValue, Type type)
        {
            try
            {
                return Convert.ChangeType(cellValue, type);
            }
            catch (Exception ex)
            {
                return null;
            }

        }

        /// <summary>
        /// 取得Model底下的Property
        /// </summary>
        /// <typeparam name="T">Model</typeparam>
        /// <param name="tempModel">Model</param>
        /// <param name="name">Property名稱</param>
        /// <returns>物件</returns>
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

        private static ISheet GetSheetByName(XSSFWorkbook xssfWorkBook, string sheetName)
        {
            List<WB> wBs = new List<WB>();
            var num = xssfWorkBook.NumberOfSheets;
            for (var i = 0; i < num; i++)
            {
                WB wb = new WB();
                wb.Index = i;
                wb.SheetName = xssfWorkBook.GetSheetName(i);
                wBs.Add(wb);
            }

            //大小寫相同
            var listWBs = wBs.Where(x => x.SheetName.ToLower().Contains(sheetName.ToLower())).ToList();
            if (listWBs.Count <= 0)
                throw new Exception("發生錯誤,找不到工作表:" + sheetName + "請確認工作表名稱");
            if (listWBs.Count > 1)
                throw new Exception("發生錯誤,找到多個工作表名稱重複:" + sheetName + "請修改重複的名稱");

            return xssfWorkBook.GetSheetAt(listWBs.First().Index);

        }

        private class WB
        {
            public int Index { get; set; }
            public string SheetName { get; set; }
        }
    }
}
