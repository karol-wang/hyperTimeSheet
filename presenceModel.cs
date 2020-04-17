using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using System.Collections;

namespace hyperTimeSheet
{
    /// <summary>
    /// 出勤物件
    /// </summary>
    class presenceModel:pathCheck
    {
        public List<exceptionCaseModel> exceptionCases;
        public presenceModel(int month)
        {
            this.FileExts = ConfigurationManager.AppSettings["presenceFileExts"].Split(',').ToList();
            this.FileFormats = ConfigurationManager.AppSettings["presenceFileFormats"].Split(',').ToList();
            this.FileRegex = ConfigurationManager.AppSettings["presenceFileRegex"];
            string presenceFilePath = ConfigurationManager.AppSettings["presenceFilePath"] == "" ? $@"{System.Environment.CurrentDirectory}\files\" : ConfigurationManager.AppSettings["presenceFilePath"];
            this.path = presenceFilePath.EndsWith(@"\") ? presenceFilePath : $@"{presenceFilePath}\";
            check(this.path);
            this.month = month;
            List<DateInfo> tmp = new List<DateInfo>();

            #region 讀取範本->取當月日期資訊
            string path = $@"{System.Environment.CurrentDirectory}\template";
            using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                IWorkbook book = new HSSFWorkbook(stream);
                ISheet sheet = book.GetSheet($"{month}月份");
                IEnumerator rows = sheet.GetRowEnumerator();
                while (rows.MoveNext())
                {
                    IRow row = (IRow)rows.Current;
                    if (row.LastCellNum > 0 && row.GetCell(0).ToString().Equals(""))
                    {
                        continue;
                    }

                    if (row.LastCellNum > 0 && row.GetCell(0).ToString().Equals("小計"))
                    {
                        //讀row到"小計" 強制離開
                        break;
                    }

                    if (row.RowNum >= 8)
                    {
                        // TimeSheet內容
                        //0 出勤日期  1 上班  2 下班  3 正常時數  4 加班時數   5 on-site單位  6 說明(假日, 休假, 病假, 事假, 其他)
                        int colorCode = row.GetCell(0).CellStyle.FillForegroundColor;
                        DateTime date = row.GetCell(0).DateCellValue;
                        string week = date.ToString("ddd").Substring(1,1); //週五 => 五
                        string note = row.GetCell(6).StringCellValue;
                        bool workingDay = colorCode == 11 ? false : true;
                        // 上班日底色代碼：64 -> auto
                        // 9 -> white
                        // 假日底色代碼：11 -> light green

                        tmp.Add(new DateInfo() { date = date, week = week, note = note, workingDay = workingDay});
                    }
                }
            }
            #endregion

            Dates = tmp;
        }

        /// <summary>
        /// 檔名格式
        /// </summary>
        public List<string> FileFormats { get; set; }

        /// <summary>
        /// 解析檔名的正規式
        /// </summary>
        public string FileRegex { get; set; }

        /// <summary>
        /// 支援的檔案副檔名
        /// </summary>
        public List<string> FileExts { get; set; }
        
        /// <summary>
        /// 資料夾路徑
        /// </summary>
        public string path { get; set; }

        /// <summary>
        /// 月份
        /// </summary>
        public int month { get; set; }

        /// <summary>
        /// 每日資訊
        /// </summary>
        public List<DateInfo> Dates { get; set; }

        /// <summary>
        /// 日期資訊：幾月幾號、週幾、是否為工作天，當天的員工資訊
        /// </summary>
        public class DateInfo
        {
            /// <summary>
            /// 日期
            /// </summary>
            public DateTime date { get; set; }

            /// <summary>
            /// 是否為上班日
            /// </summary>
            public bool workingDay { get; set; }

            public string note { get; set; }
            
            /// <summary>
            /// 星期
            /// </summary>
            public string week { get; set; }

            public List<employee> employees = new List<employee>();

            /// <summary>
            /// 取得當天請假員工列表，若當天無員工請假，則為null
            /// </summary>
            /// <param name="absence"></param>
            /// <returns></returns>
            public List<absenceModel.employee> GetAbsenceEmploees(absenceModel absence)
            {
                var absenceDate = absence.Dates.SingleOrDefault(x => x.date == date);
                return absenceDate == null ? null : absenceDate.employees;
            }
        }

        public class employee
        {
            private string no;
            /// <summary>
            /// 姓名
            /// </summary>
            public string name { get; set; }

            /// <summary>
            /// 英文名
            /// </summary>
            public string ename { get; set; }

            /// <summary>
            /// 4位數員編,ex:0001
            /// </summary>
            public string No
            {
                get => no;
                set
                {
                    no = value.Length > 4 ? value.Substring(value.Length - 4, 4) : value.PadLeft(4, '0');
                }
            }

            /// <summary>
            /// 各天上下班資訊
            /// </summary>
            public info infos = new info();

            /// <summary>
            /// 檢查上班時間
            /// </summary>
            /// <param name="exceptionCases"></param>
            /// <returns></returns>
            public bool CheckClockInTime(List<exceptionCaseModel> exceptionCases)
            {
                bool inTime = false;
                string dtime = ConfigurationManager.AppSettings["defaultClockInTime"];
                DateTime defaultTime = default(DateTime).Add(DateTime.Parse(dtime).TimeOfDay);
                if (exceptionCases != null)
                {
                    var ecp = exceptionCases.SingleOrDefault(x => x.No == No);
                    DateTime time = ecp == null ? defaultTime : ecp.specifiedClockIn;
                    inTime = infos.clockin.CompareTo(time) <= 0; //實際打卡時間小於等於設定時間 = 準時
                }

                return inTime;
            }

            /// <summary>
            /// 取得請假員工資訊
            /// </summary>
            /// <returns></returns>
            public absenceModel.employee GetAbsenceEmployee(List<absenceModel.employee> aEmps){
                return aEmps == null ? null : aEmps.SingleOrDefault(x => x.No == No);
            }
        }

        public class info
        {
            /// <summary>
            /// 上班時間
            /// </summary>
            public DateTime clockin { get; set; }

            /// <summary>
            /// 下班時間
            /// </summary>
            public DateTime clockout { get; set; }

            /// <summary>
            /// 上班時數
            /// </summary>
            public double hours { get; set; }

            /// <summary>
            /// 加班時數
            /// </summary>
            public double overtime { get; set; }

            /// <summary>
            /// 駐點單位
            /// </summary>
            public string onsite { get; set; }

            /// <summary>
            /// 說明(其他)
            /// </summary>
            public string note { get; set; }
        }

        /// <summary>
        /// 取得工作天數
        /// </summary>
        /// <returns></returns>
        public int GetWorkingDays()
        {
            int count = 0;
            count = this.Dates.Where(x => x.workingDay == true).Count();
            return count;
        }

        public List<presenceModel.employee> GetEmployees()
        {
            return this.Dates.FirstOrDefault(x => x.workingDay == true && x.employees.Count > 0).employees;
        }
        
    }

    /// <summary>
    /// 例外人員
    /// </summary>
    class exceptionCaseModel
    {
        private string no;
        public exceptionCaseModel(string No, DateTime specifiedClockIn)
        {
            //以0補齊左側至4位數
            this.No = No;
            this.specifiedClockIn = specifiedClockIn;
        }

        /// <summary>
        /// 4位數員編,ex:0001
        /// </summary>
        public string No
        {
            get => no; 
            set
            {
                no = value.Length > 4 ? value.Substring(value.Length - 4, 4) : value.PadLeft(4, '0');
            }
        }

        /// <summary>
        /// 指定上班時間
        /// </summary>
        public DateTime specifiedClockIn { get; set; }
    }

    /// <summary>
    /// 請假明細物件
    /// </summary>
    class absenceModel:pathCheck
    {
        public absenceModel(int month)
        {
            this.FileExts = ConfigurationManager.AppSettings["absenceFileExts"].Split(',').ToList();
            this.FileFormats = ConfigurationManager.AppSettings["absenceFileFormats"].Split(',').ToList();
            this.FileRegex = ConfigurationManager.AppSettings["absenceFileRegex"];
            string absenceFilePath = ConfigurationManager.AppSettings["absenceFilePath"] == "" ? $@"{System.Environment.CurrentDirectory}\absenceForms\" : ConfigurationManager.AppSettings["absenceFilePath"];
            this.path = absenceFilePath.EndsWith(@"\") ? absenceFilePath : $@"{absenceFilePath}\";
            this.month = month;
            check(this.path);
            Dates = new List<DateInfo>();
        }

        /// <summary>
        /// 檔名格式
        /// </summary>
        public List<string> FileFormats { get; set; }

        /// <summary>
        /// 解析檔名的正規式
        /// </summary>
        public string FileRegex { get; set; }

        /// <summary>
        /// 支援的檔案副檔名
        /// </summary>
        public List<string> FileExts { get; set; }

        /// <summary>
        /// 資料夾路徑
        /// </summary>
        public string path { get; set; }

        /// <summary>
        /// 月份
        /// </summary>
        public int month { get; set; }

        /// <summary>
        /// 每日資訊
        /// </summary>
        public List<DateInfo> Dates { get; set; }

        /// <summary>
        /// 日期資訊：幾月幾號，當天有假單的員工資訊
        /// </summary>
        public class DateInfo
        {
            /// <summary>
            /// 日期
            /// </summary>
            public DateTime date { get; set; }

            public List<employee> employees = new List<employee>();
        }

        public class employee
        {
            private string no;

            /// <summary>
            /// 姓名
            /// </summary>
            public string name { get; set; }

            /// <summary>
            /// 英文名
            /// </summary>
            public string ename { get; set; }

            /// <summary>
            /// 4位數員編,ex:0001
            /// </summary>
            public string No
            {
                get => no;
                set
                {
                    no = value.Length > 4 ? value.Substring(value.Length - 4, 4) : value.PadLeft(4, '0');
                }
            }

            /// <summary>
            /// 請假時數
            /// </summary>
            public double hours { get; set; }
        }
    }


    class ResultModel
    {
        /// <summary>
        /// 姓名+4位數員編,ex:0001
        /// </summary>
        public string NoAndName { get; set; }

        /// <summary>
        /// 出勤
        /// </summary>
        public string presenceMessage { get; set; }

        /// <summary>
        /// 請假
        /// </summary>
        public string absenceMessage { get; set; }

        /// <summary>
        /// 加班
        /// </summary>
        public string overtimeMessage { get; set; }
    }

    /// <summary>
    /// 繼承測式
    /// </summary>
    class pathCheck
    {
        public static void check(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }
    }
}
