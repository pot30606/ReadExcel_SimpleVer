using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace ReadExcel_SimpleVer.Model
{
    public class ExcelImport
    {

        public bool IsSuccess { get; set; } = false;

        public int row { get; set; }


        #region excel
        /// <summary>
        /// 使用者代碼(USER_ID)
        /// </summary>
        
        [Required(ErrorMessage = "使用者代碼(USER_ID)不可為空值")]
        [MaxLength(12, ErrorMessage = "使用者代碼(USER_ID) 不可為空值")]
        public string CREATE_USER { get; set; }

        /// <summary>
        /// 資料新增的系統日期時間
        /// </summary>
        
        [Required(ErrorMessage = "資料新增的系統日期時間不可為空值")]
        public DateTime CREATE_DATE { get; set; }

        /// <summary>
        /// 使用者代碼(USER_ID)
        /// </summary>
        
        [MaxLength(12, ErrorMessage = "使用者代碼(USER_ID) 不可為空值")]
        public string UPDATE_USER { get; set; }

        /// <summary>
        /// 資料修改的系統日期時間
        /// </summary>
        
        public DateTime? UPDATE_DATE { get; set; }

        /// <summary>
        /// 匯入APLDATE
        /// </summary>
        
        public DateTime? IMPORT_DATE { get; set; }


        /// <summary>
        /// 是否符合Model限制條件
        /// </summary>
        public bool IsValidModel { get; set; }
        /// <summary>
        /// 不符合條件結果說明
        /// </summary>
        public List<ValidationResult> ValidationResults { get; set; }


        #endregion


    }
}
