using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace ReadExcel_SimpleVer.Model
{
    public class Customer : ExcelImport
    {
        public int SheetIndex { get; } = 0;
        public string SheetName { get; } = "CUSTOMER";
        public string TableName { get; } = "CU_CUSTOMER_MAIN";
        #region Excel

        /// <summary>
        /// 顧客編號
        /// </summary>
        [Required(ErrorMessage = "顧客編號不可為空值")]
        public long CUSTOMER_NO { get; set; }

        /// <summary>
        /// 顧客名
        /// </summary>
        
        [Required(ErrorMessage = "顧客名不可為空值")]
        [MaxLength(20, ErrorMessage = "顧客名最大長度為 20")]
        public string CUSTOMER_NAME { get; set; }

        /// <summary>
        /// 發票名
        /// </summary>
        
        [Required(ErrorMessage = "發票名不可為空值")]
        [MaxLength(40, ErrorMessage = "發票名最大長度為 40")]
        public string INVOICE_NAME { get; set; }

        #endregion



    }

}
