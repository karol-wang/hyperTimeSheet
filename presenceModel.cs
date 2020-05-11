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
        public static List<exceptionCaseModel> ExceptionCases { get; set; }
        public presenceModel(int month)
        {
            FileExts = ConfigurationManager.AppSettings["presenceFileExts"].Split(',').ToList();
            FileFormats = ConfigurationManager.AppSettings["presenceFileFormats"].Split(',').ToList();
            FileRegex = ConfigurationManager.AppSettings["presenceFileRegex"];
            string presenceFilePath = ConfigurationManager.AppSettings["presenceFilePath"] == "" ? $@"{System.Environment.CurrentDirectory}\files\" : ConfigurationManager.AppSettings["presenceFilePath"];
            Path = presenceFilePath.EndsWith(@"\") ? presenceFilePath : $@"{presenceFilePath}\";
            Check(Path);
            Month = month;
            string dtime = ConfigurationManager.AppSettings["defaultClockInTime"];
            DefaultClockInTime = default(DateTime).Add(DateTime.Parse(dtime).TimeOfDay);
            Employees = new List<Employee>();

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

                        tmp.Add(new DateInfo() { Date = date, Week = week, Note = note, WorkingDay = workingDay});
                    }
                }
            }
            #endregion

            DateInfos = tmp;
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
        public string Path { get; set; }

        /// <summary>
        /// 月份
        /// </summary>
        public int Month { get; set; }

        /// <summary>
        /// 預設上班時間
        /// </summary>
        public static DateTime DefaultClockInTime { get; set; }

        /// <summary>
        /// 員工出勤紀錄
        /// </summary>
        public List<Employee> Employees { get; set; }

        public static List<DateInfo> DateInfos { get; set; }

        /// <summary>
        /// 日期資訊：幾月幾號、週幾、是否為工作天，當天上下班資訊
        /// </summary>
        public class DateInfo
        {
            /// <summary>
            /// 日期
            /// </summary>
            public DateTime Date { get; set; }

            /// <summary>
            /// 是否為上班日
            /// </summary>
            public bool WorkingDay { get; set; }

            public string Note { get; set; }
            
            /// <summary>
            /// 星期
            /// </summary>
            public string Week { get; set; }

            /// <summary>
            /// 當天上下班資訊
            /// </summary>
            public Info Infos = new Info();


            ///// <summary>
            ///// 取得當天請假員工列表，若當天無員工請假，則為null
            ///// </summary>
            ///// <param name="absence"></param>
            ///// <returns></returns>
            //public List<absenceModel.Employee> GetAbsenceEmploees(absenceModel absence)
            //{
            //    var absenceDate = absence.Dates.SingleOrDefault(x => x.date == Date);
            //    return absenceDate == null ? null : absenceDate.employees;
            //}
        }

        public class Employee
        {
            private string no;

            public Employee(string No)
            {
                var e = ExceptionCases?.SingleOrDefault(x => x.No == No);
                SpecifiedClockInTime = e?.SpecifiedClockIn ?? default;
                Dates = new List<DateInfo>();
            }

            /// <summary>
            /// 姓名
            /// </summary>
            public string Name { get; set; }

            /// <summary>
            /// 英文名
            /// </summary>
            public string Ename { get; set; }

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
            public DateTime SpecifiedClockInTime { get; set; }

            /// <summary>
            /// 每日資訊
            /// </summary>
            public List<DateInfo> Dates { get; set; }

            /// <summary>
            /// 檢查上班時間
            /// </summary>
            /// <param name="ClockInTime"></param>
            /// <returns></returns>
            public bool CheckClockInTime(DateInfo dateInfo)
            {
                DateTime time = SpecifiedClockInTime == default ? DefaultClockInTime : SpecifiedClockInTime;
                //若有請假，判斷時間則延後請假之時數。ex:原9:30，請假1小時，則延後至10:30
                time = dateInfo.Infos.AbsenceHours > 0 ? time.AddHours(dateInfo.Infos.AbsenceHours) : time;
                bool inTime = dateInfo.Infos.Clockin.CompareTo(time) <= 0;
                return inTime;
            }

            /// <summary>
            /// 取得請假員工資訊
            /// </summary>
            /// <returns></returns>
            public absenceModel.Employee GetAbsenceEmployee(List<absenceModel.Employee> aEmps){
                return aEmps == null ? null : aEmps.SingleOrDefault(x => x.No == No);
            }
        }

        public class Info
        {
            /// <summary>
            /// 上班時間
            /// </summary>
            public DateTime Clockin { get; set; }

            /// <summary>
            /// 下班時間
            /// </summary>
            public DateTime Clockout { get; set; }

            /// <summary>
            /// 上班時數
            /// </summary>
            public double Hours { get; set; }

            /// <summary>
            /// 加班時數
            /// </summary>
            public double Overtime { get; set; }

            /// <summary>
            /// 駐點單位
            /// </summary>
            public string Onsite { get; set; }

            /// <summary>
            /// 說明(其他)
            /// </summary>
            public string Note { get; set; }

            /// <summary>
            /// 請假時數
            /// </summary>
            public double AbsenceHours { get; set; }
        }

        /// <summary>
        /// 取得工作天數
        /// </summary>
        /// <returns></returns>
        public int GetWorkingDays()
        {
            int count = 0;
            count = DateInfos.Where(x => x.WorkingDay == true).Count();
            return count;
        }

        /// <summary>
        /// 取得員工數
        /// </summary>
        /// <returns></returns>
        public int GetEmployees()
        {
            return Employees.Count;
        }

        //public List<presenceModel.Employee> GetEmployees()
        //{
        //    return this.Dates.FirstOrDefault(x => x.workingDay == true && x.employees.Count > 0).employees;
        //}
        
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
            this.SpecifiedClockIn = specifiedClockIn;
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
        public DateTime SpecifiedClockIn { get; set; }
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
            this.Path = absenceFilePath.EndsWith(@"\") ? absenceFilePath : $@"{absenceFilePath}\";
            this.Month = month;
            Check(this.Path);
            Employees = new List<Employee>();
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
        public string Path { get; set; }

        /// <summary>
        /// 月份
        /// </summary>
        public int Month { get; set; }

        /// <summary>
        /// 員工列表
        /// </summary>
        public List<Employee> Employees { get; set; }

        /// <summary>
        /// 日期資訊：幾月幾號，當天有假單的員工資訊
        /// </summary>
        public class DateInfo
        {
            /// <summary>
            /// 日期
            /// </summary>
            public DateTime Date { get; set; }

            /// <summary>
            /// 請假時數
            /// </summary>
            public double Hours { get; set; }
        }

        public class Employee
        {
            private string no;

            /// <summary>
            /// 姓名
            /// </summary>
            public string Name { get; set; }

            /// <summary>
            /// 英文名
            /// </summary>
            public string Ename { get; set; }

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
            /// 每日資訊
            /// </summary>
            public List<DateInfo> Dates = new List<DateInfo>();
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
        public string PresenceMessage { get; set; }

        /// <summary>
        /// 請假
        /// </summary>
        public string AbsenceMessage { get; set; }

        /// <summary>
        /// 加班
        /// </summary>
        public string OvertimeMessage { get; set; }
    }

    /// <summary>
    /// 繼承測式
    /// </summary>
    class pathCheck
    {
        public static void Check(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }
    }
}
