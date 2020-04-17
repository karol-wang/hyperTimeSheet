using System;
using System.Collections.Generic;
using System.Linq;
using System.Configuration;
using System.IO;
using System.Reflection;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.HSSF.UserModel;
using System.Collections;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace hyperTimeSheet
{
    class Program
    {
        static string root = System.Environment.CurrentDirectory;
        static int thisYear = DateTime.Now.Year;
        static int thisMonth = 0;
        static string log = null;
        static readonly string title = $"HyperTimeSheet Ver 2";
        //static string currentFile = "";
        static DateTime defaultTime = default; //設定的上班時間
        static presenceModel presenceModel = null;
        static absenceModel absenceModel = null;
        static List<exceptionCaseModel> exceptionCases = null;
        static List<ResultModel> Results = null;

        static void Main(string[] args)
        {
            Console.Title = title;

            #region initialization 
            try
            {
                ini_Log();

                Console.Write("輸入月份:");
                string tmp = Console.ReadLine();
                while (!int.TryParse(tmp, out thisMonth) || (thisMonth > 12 || thisMonth < 1))
                {
                    Console.WriteLine("輸入錯誤！");
                    Console.Write("輸入月份:");
                }
                Console.Title = $"{title} ({thisYear}.{thisMonth})";

                //初使化出勤物件
                presenceModel = new presenceModel(thisMonth);
                //初使化請假明細物件
                absenceModel = new absenceModel(thisMonth);
                //初使化例外列表
                presenceModel.exceptionCases = exceptionCases = GetExceptionCase();
                //初使化結果列表
                //Results = new List<ResultModel>();
                //初使化預設上班時間
                string time = ConfigurationManager.AppSettings["defaultClockInTime"];
                defaultTime = default(DateTime).Add(DateTime.Parse(time).TimeOfDay);
            }
            catch (Exception e)
            {
                Console.BackgroundColor = ConsoleColor.White;
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("糟了！程式被你弄壞了！");
                Log(LogMessagesType.error, $"初使化發生錯誤。{e.Message}", e.LineNumber(), MethodBase.GetCurrentMethod().Name);
                Console.ResetColor();
            }
            #endregion

            try
            {
                absenceModel = GetAbsenceModel(absenceModel);
                presenceModel = GetPresenceModel(presenceModel);
            }
            catch (Exception e)
            {
                Console.BackgroundColor = ConsoleColor.White;
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("糟了！程式被你弄壞了！");
                Console.ResetColor();
                Log(LogMessagesType.error, $"建立物件發生錯誤。{e.Message}", e.LineNumber(), MethodBase.GetCurrentMethod().Name);
            }

            try
            {
                ExportExcel(presenceModel, absenceModel, exceptionCases);
            }
            catch
            {
                Console.BackgroundColor = ConsoleColor.White;
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("糟了！程式被你弄壞了！");
                Console.ResetColor();
            }

            Console.ResetColor();
            Console.Write(@"檢查完成，請按""Enter""關閉程式...");
            Console.ReadLine();
        }

        #region LOG
        /// <summary>
        /// 初使化Log物件
        /// </summary>
        private static void ini_Log()
        {
            string LogPath = $@"{root}\Log\";
            log = $@"{LogPath}log{DateTime.Today.ToString("yyyy-MM-dd")}.txt";
            if (!Directory.Exists(LogPath))
            {
                Directory.CreateDirectory(LogPath);
            }

            if (!File.Exists(log))
            {
                File.Create(log).Close();
            }
        }

        /// <summary>
        /// 寫入Log
        /// </summary>
        /// <param name="type"></param>
        /// <param name="logMessage"></param>
        /// <param name="errorLineNumber"></param>
        /// <param name="inMethod"></param>
        private static void Log(LogMessagesType type, string logMessage, int errorLineNumber, string inMethod)
        {
            if (string.IsNullOrEmpty(log))
            {
                throw new Exception("Log物件尚未初使化");
            }
            StreamWriter sw = File.AppendText(log);
            if (type == LogMessagesType.error)
            {
                sw.WriteLine($"{DateTime.Now.ToString("yyyy-MM-dd  hh:mm:ss")} | {type} on line:{errorLineNumber} at {inMethod} | {logMessage.Replace('\r', ' ')}");
            }
            else
            {
                sw.WriteLine($"{DateTime.Now.ToString("yyyy-MM-dd  hh:mm:ss")} | {type} | {logMessage.Replace('\r', ' ')}");
            }
            sw.Close();
        }
        #endregion

        #region 例外名單列表
        /// <summary>
        /// 取得例外名單列表
        /// </summary>
        /// <returns></returns>
        public static List<exceptionCaseModel> GetExceptionCase()
        {
            List<exceptionCaseModel> list = new List<exceptionCaseModel>();
            string filePath = $@"{root}\exceptionCase.txt";
            string line = string.Empty;
            try
            {
                StreamReader stream = new StreamReader(filePath);
                while ((line = stream.ReadLine()) != null)
                {
                    string[] tmp = line.Split(',');
                    DateTime time = default;
                    //Ex: 0001,9:30
                    //Ex: 0001,0005,0009,9:00
                    if (tmp.Length < 2)
                    {
                        throw new Exception("Exception case 格式有誤");
                    }

                    if (!DateTime.TryParse(tmp[tmp.Length - 1], out time))
                    {
                        throw new Exception($"{tmp[tmp.Length - 1]} 無法轉換時間");
                    }

                    for (int i = 0; i < tmp.Length - 1; i++)
                    {
                        //最後一筆為時間，不應以迴圈讀取
                        list.Add(new exceptionCaseModel(tmp[i], default(DateTime).Add(time.TimeOfDay)));
                    }
                }
            }
            catch (Exception e)
            {
                Log(LogMessagesType.error, $"讀取Exception Case發生錯誤，已略過讀取\"{line}\"。原因：{e.Message}", e.LineNumber(), MethodBase.GetCurrentMethod().Name);
            }

            return list;
        }
        #endregion

        #region 讀取假單
        /// <summary>
        /// 讀取指定路徑假單檔案(doc|docx)，並回傳absenceModel
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        private static absenceModel GetAbsenceModel(absenceModel model)
        {
            string currentFile = "";
            try
            {
                List<string> exts = model.FileExts;
                List<string> formats = model.FileFormats;
                List<string> files = Directory.GetFiles(absenceModel.path).ToList();

                string fileNamePattern = model.FileRegex;
                //List<absenceModel.DateInfo> dates = new List<absenceModel.DateInfo>();
                int current = 1;
                int count = files.Count;
                foreach (var file in files)
                {
                    try
                    {
                        string ext = Path.GetExtension(file);
                        string fileName = Path.GetFileNameWithoutExtension(file);
                        currentFile = file;
                        if (exts.Contains(ext))
                        {
                            //檢查 & 讀取檔名
                            Regex regex = new Regex(fileNamePattern);
                            Match match = regex.Match(fileName);

                            if (!match.Success)
                            {
                                throw new Exception("檔名格式有誤");
                            }

                            //{no},{name},{year},{month},{day},{day2},{hours}
                            //regex group第一組為 Full match (須略過要 +1)
                            int num_no = formats.IndexOf("{no}") + 1;
                            int num_name = formats.IndexOf("{name}") + 1;
                            int num_year = formats.IndexOf("{year}") + 1;
                            int num_month = formats.IndexOf("{month}") + 1;
                            int num_day = formats.IndexOf("{day}") + 1;
                            int num_day2 = formats.IndexOf("{day2}") + 1;
                            int num_hours = formats.IndexOf("{hours}") + 1;
    
                            //取值
                            string No = match.Groups[num_no].ToString();
                            string name = match.Groups[num_name].ToString();
                            int y = int.Parse(match.Groups[num_year].ToString()) + 1911; // 轉成西元年
                            int m = int.Parse(match.Groups[num_month].ToString());
                            int d = int.Parse(match.Groups[num_day].ToString());
                            int m2 = 0;
                            int d2 = 0;
                            string day2 = match.Groups[num_day2].ToString();

                            #region 處理月份不等於當月，則跳過
                            if (model.month != m || thisYear != y)
                            {
                                Log(LogMessagesType.info, $"{currentFile}不屬於{thisYear}.{model.month}月，已略過讀取。", 0, "");
                                continue;
                            }
                            #endregion

                            int hours = match.Groups[num_hours].ToString() == "" ? 8 : int.Parse(match.Groups[num_hours].ToString());

                            #region 處理第2組日期
                            if (int.TryParse(day2, out d2))
                            {
                                //若第2組日期為2碼(僅有日)，第2組月份則與第1組相同。Ex : 09
                                m2 = m;
                                if (day2.Length > 2)
                                {
                                    //Ex : 0309
                                    m2 = int.Parse(day2.Substring(0, 2));
                                    d2 = int.Parse(day2.Substring(2, 2));
                                }
                            }
                            else if (day2 == "")
                            {
                                //若第2組日期為空，表示僅有1組日期
                                m2 = m;
                                d2 = d;
                            }
                            else
                            {
                                throw new Exception($"第2組日期轉換時失敗！");
                            }
                            #endregion

                            List<DateTime> absenceDates = GetAllDays(new DateTime(y, m, d), new DateTime(y, m2, d2));

                            foreach (var date in absenceDates)
                            {
                                absenceModel.DateInfo dateinfo = model.Dates.SingleOrDefault(x => x.date == date);
                                if (dateinfo == null)
                                {
                                    //model中若沒有 date => 新增
                                    dateinfo = new absenceModel.DateInfo { date = date };
                                    dateinfo.employees.Add(new absenceModel.employee { No = No, ename = name, hours = hours });
                                    model.Dates.Add(dateinfo);
                                }
                                else
                                {
                                    dateinfo.employees.Add(new absenceModel.employee { No = No, ename = name, hours = hours });
                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Log(LogMessagesType.error, $"讀取假單發生錯誤，已略過讀取{currentFile}。原因：{e.Message}", e.LineNumber(), MethodBase.GetCurrentMethod().Name);
                        Console.WriteLine($"讀取假單發生錯誤，已略過讀取{currentFile}。原因：{e.Message}");
                    }
                    finally
                    {
                        proccess("假單檔案處理中", current, count);
                        current++;
                    }
                }
            }
            catch (Exception e)
            {
                Log(LogMessagesType.error, $"讀取假單發生嚴重錯誤，已終止讀取。原因：{e.Message}", e.LineNumber(), MethodBase.GetCurrentMethod().Name);
                throw;
            }
            return model;
        }
        #endregion

        #region 讀取TimeSheet
        /// <summary>
        /// 讀取指定路徑TimeSheet(xls|xlsx)，並回傳presenceModel
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        private static presenceModel GetPresenceModel(presenceModel model)
        {
            string currentFile = "";
            try
            {
                #region 讀取Excel
                List<string> exts = presenceModel.FileExts;
                List<string> formats = model.FileFormats;
                List<string> files = Directory.GetFiles(presenceModel.path).ToList();

                string fileNamePattern = model.FileRegex;
                int current = 1;
                int count = files.Count();
                foreach (var file in files)
                {
                    try
                    {
                        string fileName = Path.GetFileNameWithoutExtension(file);
                        string ext = Path.GetExtension(file);
                        currentFile = file;
                        IWorkbook book = null;
                        if (exts.Contains(ext))
                        {
                            //檢查 & 讀取 excel 檔案
                            Regex regex = new Regex(fileNamePattern);
                            Match match = regex.Match(fileName);

                            if (!match.Success)
                            {
                                throw new Exception("檔名格式有誤");
                            }
                            //{year},{no},{name}
                            //regex group第一組為 Full match (須略過要 +1)
                            int num_year = formats.IndexOf("{year}") + 1;
                            int num_no = formats.IndexOf("{no}") + 1;
                            int num_name = formats.IndexOf("{name}") + 1;

                            //取值
                            int year = int.Parse(match.Groups[num_year].ToString());
                            string No = match.Groups[num_no].ToString();
                            string name = match.Groups[num_name].ToString();

                            using (FileStream stream = new FileStream(file, FileMode.Open, FileAccess.Read))
                            {
                                if (ext == ".xls") { book = new HSSFWorkbook(stream); }
                                else if (ext == ".xlsx") { book = new XSSFWorkbook(stream); }
                            }
                            ISheet sheet = book.GetSheet($"{model.month}月份");
                            
                            #region 若序列名稱找不到sheet(前後有空白)，則逐筆搜尋sheet，比對去空白後的sheet名稱
                            if (sheet == null)
                            {
                                IEnumerator sheets = book.GetEnumerator();
                                while (sheets.MoveNext())
                                {
                                    ISheet sheet1 = (ISheet)sheets.Current;
                                    if (sheet1.SheetName.Trim() == $"{model.month}月份")
                                    {
                                        sheet = sheet1;
                                        break;
                                    }
                                }
                            } 
                            #endregion
                            
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

                                //if (row.RowNum == 4)
                                //{
                                //    // TimeSheet月份 & 名字
                                //    //Console.ForegroundColor = ConsoleColor.Yellow;
                                //    //Console.WriteLine("".PadLeft(30, '='));
                                //    //Console.WriteLine(row.GetCell(0).ToString());
                                //    //Console.WriteLine("".PadLeft(30, '='));
                                //    //Console.ResetColor();
                                //}

                                if (row.RowNum >= 8)
                                {
                                    int colorCode = row.GetCell(0).CellStyle.FillForegroundColor;
                                    // TimeSheet內容
                                    //0 出勤日期  1 上班  2 下班  3 正常時數  4 加班時數   5 on-site單位  6 說明(假日, 休假, 病假, 事假, 其他)
                                    DateTime date = getValue<DateTime>(row.Cells[0]);
                                    //取得當月所有日期資訊
                                    presenceModel.DateInfo dateInfo = model.Dates.SingleOrDefault(x => x.date == date);
                                    if (dateInfo.workingDay)
                                    {
                                        #region columns to model.dates.dateinfo.employee.info
                                        presenceModel.employee employee = new presenceModel.employee
                                        {
                                            No = No,
                                            name = name,
                                            infos = new presenceModel.info
                                            {
                                                clockin = getValue<DateTime>(row.Cells[1]),
                                                clockout = getValue<DateTime>(row.Cells[2]),
                                                hours = getValue<double>(row.Cells[3]),
                                                overtime = getValue<double>(row.Cells[4]),
                                                onsite = getValue<string>(row.Cells[5]),
                                                note = getValue<string>(row.Cells[6])
                                            }
                                        };

                                        dateInfo.employees.Add(employee);
                                        #endregion
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Log(LogMessagesType.error, $"讀取TimeSheet發生錯誤，已略過讀取{currentFile}。原因：{e.Message}", e.LineNumber(), MethodBase.GetCurrentMethod().Name);
                        Console.WriteLine($"讀取TimeSheet發生錯誤，已略過讀取{currentFile}。原因：{e.Message}");
                    }
                    finally
                    {
                        proccess("TimeSheet讀取中", current, count);
                        current++;
                    }
                }
                #endregion
            }
            catch (Exception e)
            {
                Log(LogMessagesType.error, $"讀取TimeSheet發生嚴重錯誤，已終止讀取。原因：{e.Message}", e.LineNumber(), MethodBase.GetCurrentMethod().Name);
                throw;
            }
            return model;

            #region 取cell值(指定型別)
            dynamic getValue<T>(ICell cell)
            {
                var type = typeof(T);
                double d = 0d;
                DateTime dateTime = default;
                dynamic temp;
                switch (cell.CellType)
                {
                    case CellType.Numeric:  // 數值格式
                        if (DateUtil.IsCellDateFormatted(cell))
                        {   // 日期格式=> 大於等於1：1900/1/1以後的日期；小於0：表示時間(乘上24)
                            if (cell.NumericCellValue >= 1)
                            {
                                temp = cell.DateCellValue;
                            }
                            else
                            {
                                temp = DateTime.MinValue.AddDays(cell.NumericCellValue);
                            }
                            break;
                        }
                        else
                        {   // 數值格式
                            temp = cell.NumericCellValue;
                            break;
                        }
                    case CellType.String:   // 字串格式
                        temp = cell.StringCellValue;
                        break;
                    case CellType.Formula:
                        IFormulaEvaluator iFormula = WorkbookFactory.CreateFormulaEvaluator(cell.Sheet.Workbook);
                        var formulaType = iFormula.Evaluate(cell).CellType;
                        if (formulaType == CellType.Numeric)
                        {
                            if (DateUtil.IsCellDateFormatted(cell))
                            {   // 日期格式=> 大於等於1：1900/1/1以後的日期；小於0：表示時間(乘上24)
                                if (cell.NumericCellValue >= 1)
                                {
                                    temp = cell.DateCellValue;
                                }
                                else
                                {
                                    temp = DateTime.MinValue.AddDays(cell.NumericCellValue);
                                }
                                break;
                            }
                            else
                            {   // 數值格式
                                temp = cell.NumericCellValue;
                                break;
                            }
                        }
                        else
                        {
                            temp = cell.StringCellValue;
                            break;
                        }
                    default:
                        temp = cell.ToString();
                        break;
                }

                if (type == temp.GetType())
                {
                    return temp;
                }
                else if (type == typeof(string))
                {
                    return temp.ToString();
                }
                else if (type == typeof(DateTime))
                {
                    DateTime.TryParse(temp.ToString(), out dateTime);
                    temp = dateTime;
                    return temp;
                }
                else if (type == typeof(double))
                {
                    double.TryParse(temp.ToString(), out d);
                    temp = d;
                    return temp;
                }
                else
                {
                    return temp;
                }
            }
            #endregion
        }
        #endregion

        #region 檢查上班時間
        /// <summary>
        /// 檢查emp的上班時間，若emp在例外清單中，則以例外清單所設定的時間為準，否則以預設時間設定。回傳是否準時
        /// </summary>
        /// <param name="emp"></param>
        /// <returns></returns>
        private static bool checkClockInTime(presenceModel.employee emp)
        {
            bool inTime = false;

            if (exceptionCases != null)
            {
                var ecp = exceptionCases.SingleOrDefault(x => x.No == emp.No);
                DateTime time = ecp == null ? defaultTime : ecp.specifiedClockIn;
                inTime = emp.infos.clockin.CompareTo(time) <= 0; //實際打卡時間小於等於設定時間 = 準時
            }

            return inTime;
        } 
        #endregion

        #region console進度
        /// <summary>
        /// (message)... x/y 
        /// </summary>
        /// <param name="current"></param>
        /// <param name="count"></param>
        private static void proccess(string messgae, int current, int count)
        {
            int currentTop = Console.CursorTop;
            Console.SetCursorPosition(0, currentTop);
            Console.Write($"{messgae}...{current}/{count}");
            if (current == count)
            {
                Console.SetCursorPosition(0, currentTop + 1);
                Console.WriteLine("完成");
                Console.WriteLine("");
            }
        }
        #endregion

        #region 取得指定範圍內所有日期
        /// <summary> 
        /// 取得指定範圍內所有日期，並回傳List
        /// </summary>  
        /// <param name="dt1">開始日期</param>  
        /// <param name="dt2">結束日期</param>  
        /// <returns></returns>  
        private static List<DateTime> GetAllDays(DateTime dt1, DateTime dt2)
        {
            List<DateTime> listDays = new List<DateTime>();
            DateTime dtDay = new DateTime();
            for (dtDay = dt1; dtDay.CompareTo(dt2) <= 0; dtDay = dtDay.AddDays(1))
            {
                listDays.Add(dtDay);
            }
            return listDays;
        }
        #endregion

        #region 匯出EXCEL

        /// <summary>
        /// 匯出至EXCEL
        /// </summary>
        private static void ExportExcel(presenceModel presence, absenceModel absence, List<exceptionCaseModel> exceptions)
        {
            var wb = new XSSFWorkbook();
            ISheet sheetResult = wb.CreateSheet("測試輸出結果");
            sheetResult.SetColumnWidth(0, 15 * 256);
            for (int i = 1; i <= 3; i++)
            {
                sheetResult.SetColumnWidth(i, 30 * 256);
            }
            sheetResult.CreateFreezePane(0, 1);
            ISheet sheetArrange = wb.CreateSheet("出勤一覽");
            sheetArrange.SetColumnWidth(0, 15 * 256);
            sheetArrange.CreateFreezePane(0, 1);

            try
            {
                #region 粗底細邊(左右)+底色
                XSSFCellStyle firstRowStyle = (XSSFCellStyle)wb.CreateCellStyle();
                XSSFColor color = new XSSFColor();
                color.SetRgb(new byte[] { (byte)136, (byte)160, (byte)226 });
                firstRowStyle.FillForegroundColorColor = color;
                firstRowStyle.FillPattern = FillPattern.SolidForeground;
                firstRowStyle.BorderBottom = BorderStyle.Thick;
                firstRowStyle.BorderRight = BorderStyle.Thin;
                firstRowStyle.BorderLeft = BorderStyle.Thin;
                #endregion

                IRow firstRow = sheetArrange.CreateRow(0);
                int count = presence.GetWorkingDays();
                firstRow.CreateCell(0).SetCellValue($"上班日共{count}日");
                firstRow.GetCell(0).CellStyle = firstRowStyle;
                int index = 1;
                foreach (var date in presence.Dates)
                {
                    if (date.workingDay)
                    {
                        firstRow.CreateCell(index).SetCellValue($"{date.date.ToString("MM-dd")}({date.week})");
                        firstRow.GetCell(index).CellStyle = firstRowStyle;
                        sheetArrange.SetColumnWidth(index, 10 * 256);
                        foreach (var emp in date.employees.Select((v, i) => new { index = i, value = v }))
                        {
                            var absenceEmps = date.GetAbsenceEmploees(absence);
                            var absenceEmp = emp.value.GetAbsenceEmployee(absenceEmps);
                            writeEmlpoyeeByDate(sheetArrange, emp.index, index, emp.value, absenceEmp);
                        }
                        proccess($"寫入工作天", index, count);
                        index++;
                    }
                }

                IRow firstRow2 = sheetResult.CreateRow(0);
                firstRow2.CreateCell(0);
                firstRow2.CreateCell(1).SetCellValue("出勤");
                firstRow2.CreateCell(2).SetCellValue("請假");
                firstRow2.CreateCell(3).SetCellValue("加班");
                for (int i = 0; i <= 3; i++)
                {
                    firstRow2.GetCell(i).CellStyle = firstRowStyle;
                }
                #region 粗底細邊
                XSSFCellStyle RowStyle = (XSSFCellStyle)wb.CreateCellStyle();
                RowStyle.BorderBottom = BorderStyle.Thick;
                RowStyle.BorderLeft = BorderStyle.Thin;
                RowStyle.BorderRight = BorderStyle.Thin;
                RowStyle.WrapText = true;
                RowStyle.VerticalAlignment = VerticalAlignment.Center;
                #endregion
                var emplist = presenceModel.GetEmployees();
                int index2 = 1;
                int count2 = emplist.Count;
                foreach (var emp in emplist)
                {
                    var result = (GetResult(presenceModel, absenceModel, exceptions, emp.No, emp.name));
                    IRow row = sheetResult.CreateRow(index2);
                    row.CreateCell(0).SetCellValue(result.NoAndName);
                    row.CreateCell(1);
                    row.CreateCell(2);
                    row.CreateCell(3).SetCellValue(result.overtimeMessage == "" ? "無" : result.overtimeMessage);
                    row.GetCell(0).CellStyle = RowStyle;
                    #region 出勤(1) 異常=>紅字
                    if (result.presenceMessage != "")
                    {
                        XSSFCellStyle presenceRowStyle = (XSSFCellStyle)wb.CreateCellStyle();
                        presenceRowStyle.BorderBottom = BorderStyle.Thick;
                        presenceRowStyle.BorderLeft = BorderStyle.Thin;
                        presenceRowStyle.BorderRight = BorderStyle.Thin;
                        presenceRowStyle.WrapText = true;
                        presenceRowStyle.VerticalAlignment = VerticalAlignment.Center;
                        XSSFFont redFont = (XSSFFont)wb.CreateFont();
                        redFont.Color = HSSFColor.Red.Index;
                        redFont.FontHeightInPoints = 11;
                        presenceRowStyle.SetFont(redFont);
                        row.GetCell(1).SetCellValue(result.presenceMessage);
                        row.GetCell(1).CellStyle = presenceRowStyle;
                    }
                    else
                    {
                        row.GetCell(1).SetCellValue("正常");
                        row.GetCell(1).CellStyle = RowStyle;
                    }
                    #endregion
                    #region 請假(2) =>綠字
                    if (result.absenceMessage != "")
                    {
                        XSSFCellStyle absenceRowStyle = (XSSFCellStyle)wb.CreateCellStyle();
                        absenceRowStyle.BorderBottom = BorderStyle.Thick;
                        absenceRowStyle.BorderLeft = BorderStyle.Thin;
                        absenceRowStyle.BorderRight = BorderStyle.Thin;
                        absenceRowStyle.WrapText = true;
                        absenceRowStyle.VerticalAlignment = VerticalAlignment.Center;
                        XSSFFont greenFont = (XSSFFont)wb.CreateFont();
                        greenFont.Color = HSSFColor.Green.Index;
                        greenFont.FontHeightInPoints = 11;
                        absenceRowStyle.SetFont(greenFont);
                        row.GetCell(2).SetCellValue(result.absenceMessage);
                        row.GetCell(2).CellStyle = absenceRowStyle;
                    }
                    else
                    {
                        row.GetCell(2).SetCellValue("無");
                        row.GetCell(2).CellStyle = RowStyle;
                    }
                    #endregion
                    row.GetCell(3).CellStyle = RowStyle;
                    proccess("輸出彙整結果", index2, count2);
                    index2++;
                }

                var fs = new FileStream($"{thisYear}.{thisMonth}.xlsx", FileMode.Create);
                wb.Write(fs);
                fs.Close();
            }
            catch (Exception e)
            {
                Log(LogMessagesType.error, $"匯出excel時發生錯誤。原因：{e.Message}", e.LineNumber(), MethodBase.GetCurrentMethod().Name);
                throw;
            }

            #region 寫入員工資料
            void writeEmlpoyeeByDate(ISheet sheet, int index, int dateCol, presenceModel.employee emp, absenceModel.employee aEmp)
            {
                #region 員工資料格式
                XSSFCellStyle empStyle = (XSSFCellStyle)wb.CreateCellStyle();
                empStyle.BorderBottom = BorderStyle.Thick;
                empStyle.BorderRight = BorderStyle.Thin;
                empStyle.BorderLeft = BorderStyle.Thin;
                #endregion
                #region 打卡資訊格式
                XSSFCellStyle infoStyle = (XSSFCellStyle)wb.CreateCellStyle();
                infoStyle.BorderBottom = BorderStyle.Thin;
                infoStyle.BorderRight = BorderStyle.Thin;
                infoStyle.BorderLeft = BorderStyle.Thin;
                #endregion
                #region 打卡資訊第一列
                XSSFCellStyle firstRowStyle = (XSSFCellStyle)wb.CreateCellStyle();
                firstRowStyle.BorderBottom = BorderStyle.Thin;
                firstRowStyle.BorderLeft = BorderStyle.Thin;
                firstRowStyle.BorderRight = BorderStyle.Thin;
                if (!emp.CheckClockInTime(exceptionCases))
                {
                    //超過時間打卡 => 上班時間設定紅字
                    XSSFFont redFont = (XSSFFont)wb.CreateFont();
                    redFont.Color = HSSFColor.Red.Index;
                    redFont.FontHeightInPoints = 11;
                    firstRowStyle.SetFont(redFont);
                }
                #endregion
                #region 打卡資訊最後一列
                XSSFCellStyle lastRowStyle = (XSSFCellStyle)wb.CreateCellStyle();
                lastRowStyle.BorderBottom = BorderStyle.Thick;
                lastRowStyle.BorderLeft = BorderStyle.Thin;
                lastRowStyle.BorderRight = BorderStyle.Thin;
                lastRowStyle.WrapText = true;
                lastRowStyle.Alignment = HorizontalAlignment.Left;
                lastRowStyle.VerticalAlignment = VerticalAlignment.Top;
                #endregion

                #region 醒目底色
                if (index % 2 == 0)
                {
                    XSSFColor color = new XSSFColor();
                    color.SetRgb(new byte[] { (byte)230, (byte)230, (byte)230 });
                    empStyle.FillForegroundColorColor = color;
                    empStyle.FillPattern = FillPattern.SolidForeground;
                    infoStyle.FillForegroundColorColor = color;
                    infoStyle.FillPattern = FillPattern.SolidForeground;
                    firstRowStyle.FillForegroundColorColor = color;
                    firstRowStyle.FillPattern = FillPattern.SolidForeground;
                    lastRowStyle.FillForegroundColorColor = color;
                    lastRowStyle.FillPattern = FillPattern.SolidForeground;
                }
                #endregion

                #region 請假資訊
                int absenceHours = 0; //請假時數
                string absenceHoursStr = "";
                if (aEmp != null)
                {
                    //表示該員工有請假資訊
                    absenceHours = (int)aEmp.hours;
                    absenceHoursStr = $"\n({aEmp.hours})";
                    XSSFColor color = new XSSFColor();
                    color.SetRgb(new byte[] { (byte)147, (byte)255, (byte)147 });
                    infoStyle.FillForegroundColorColor = color;
                    infoStyle.FillPattern = FillPattern.SolidForeground;
                    firstRowStyle.FillForegroundColorColor = color;
                    firstRowStyle.FillPattern = FillPattern.SolidForeground;
                    lastRowStyle.FillForegroundColorColor = color;
                    lastRowStyle.FillPattern = FillPattern.SolidForeground;
                }
                #endregion

                #region 出勤異常加粉紅底色
                if (absenceHours < 8 && (emp.infos.clockin == DateTime.MinValue || emp.infos.clockout == DateTime.MinValue || (emp.infos.hours < 8 - absenceHours)))
                {
                    XSSFColor color = new XSSFColor();
                    color.SetRgb(new byte[] { (byte)255, (byte)147, (byte)147 });
                    infoStyle.FillForegroundColorColor = color;
                    infoStyle.FillPattern = FillPattern.SolidForeground;
                    firstRowStyle.FillForegroundColorColor = color;
                    firstRowStyle.FillPattern = FillPattern.SolidForeground;
                    lastRowStyle.FillForegroundColorColor = color;
                    lastRowStyle.FillPattern = FillPattern.SolidForeground;
                }
                #endregion

                //每個人皆5列
                int rowspan = 5;
                try
                {
                    for (int i = 1; i <= rowspan; i++)
                    {
                        //每人跨列的列數 * 第？個人 + 該人第幾列
                        int r = rowspan * index + i;
                        IRow row = sheet.GetRow(r) == null ? sheet.CreateRow(r) : sheet.GetRow(r);

                        #region 寫入第一個日期的打卡資料時，同時在第一欄cell(0)寫入員編&名字
                        if (dateCol == 1)
                        {
                            empStyle.VerticalAlignment = VerticalAlignment.Center;
                            empStyle.WrapText = true;
                            var exception = exceptionCases.SingleOrDefault(x => x.No == emp.No);
                            string addition = exception == null ? "" : $"({exception.specifiedClockIn.ToString("HH:mm")})";
                            switch (i)
                            {
                                case 1:
                                    row.CreateCell(0).SetCellValue($"{emp.No}-{emp.name}\r{addition}");
                                    sheet.AddMergedRegion(new CellRangeAddress(rowspan * index + 1, rowspan * index + 5, 0, 0));
                                    row.GetCell(0).CellStyle = empStyle;
                                    break;
                                default:
                                    // 2~4
                                    row.CreateCell(0).CellStyle = empStyle;
                                    break;
                                case 5:
                                    //每人最後一列
                                    row.CreateCell(0).CellStyle = empStyle;
                                    break;
                            }
                        }
                        #endregion

                        #region 寫入當天打卡資料
                        switch (i)
                        {
                            case 1:
                                row.CreateCell(dateCol).SetCellValue(DateTime.MinValue == emp.infos.clockin ? "" : emp.infos.clockin.ToString("HH:mm"));
                                row.GetCell(dateCol).CellStyle = firstRowStyle;
                                break;
                            case 2:
                                row.CreateCell(dateCol).SetCellValue(DateTime.MinValue == emp.infos.clockout ? "" : emp.infos.clockout.ToString("HH:mm"));
                                row.GetCell(dateCol).CellStyle = infoStyle;
                                break;
                            case 3:
                                row.CreateCell(dateCol).SetCellValue(emp.infos.hours);
                                row.GetCell(dateCol).SetCellType(CellType.String);
                                row.GetCell(dateCol).CellStyle = infoStyle;
                                break;
                            case 4:
                                row.CreateCell(dateCol).SetCellValue(emp.infos.onsite);
                                row.GetCell(dateCol).CellStyle = infoStyle;
                                break;
                            case 5:
                                row.CreateCell(dateCol).SetCellValue($"{emp.infos.note}{absenceHoursStr}");
                                row.GetCell(dateCol).CellStyle = lastRowStyle;
                                break;
                        }
                        #endregion
                    }
                }
                catch (Exception e)
                {
                    Log(LogMessagesType.error, $"寫入employee時發生錯誤。原因：{e.Message}", e.LineNumber(), MethodBase.GetCurrentMethod().Name);
                    throw;
                }
            }
            #endregion

            #region 彙整員工當月出勤結果
            ResultModel GetResult(presenceModel presenceModel, absenceModel absenceModel, List<exceptionCaseModel> exceptionCases, string no, string name)
            {
                //取得該員工當月所有出勤資料
                var res = presenceModel.Dates.SelectMany(
                    x => x.employees,
                    (d, e) => new
                    {
                        d.date,
                        dateStr = $"{d.date.ToString("MM/dd")}({d.week})",
                        employee = e
                    }
                ).Where(x => x.employee.No == no).Select(
                    x => new
                    {
                        x.date,
                        x.dateStr,
                        x.employee
                    }
                );

                //取得該員工當月所有請假資料
                var absenceRes = absenceModel.Dates.SelectMany(
                    x => x.employees,
                    (d, e) => new
                    {
                        d.date,
                        employee = e
                    }
                ).Where(x => x.employee.No == no).Select(
                    x => new
                    {
                        x.date,
                        x.employee.hours
                    }
                );
                var exception = exceptionCases.SingleOrDefault(x => x.No == no);
                var Result = new ResultModel();
                Result.NoAndName = exception == null ? $"{no}-{name}" : $"{no}-{name}\n({exception.specifiedClockIn.ToString("HH:mm")})";
                var tmp = new Dictionary<string, List<string>>();
                //Key：presence、absence、overtime
                //處理單一員工當月每個工作天
                foreach (var item in res)
                {
                    //取得當天請假時數
                    var ab = absenceRes.SingleOrDefault(x => x.date == item.date);
                    int absenceHours = ab == null ? 0 : (int)ab.hours;
                    #region 出勤
                    if (absenceHours >= 8)
                    {
                        //請假時數大於等於8 => 請全天不處理
                    }
                    else if (item.employee.infos.clockin == DateTime.MinValue && item.employee.infos.clockout == DateTime.MinValue)
                    {
                        //無上班 & 下班紀錄 (未請假)
                        string m = $"{item.dateStr} 未請假";
                        if (tmp.ContainsKey("presence"))
                        {
                            tmp["presence"].Add(m);
                        }
                        else
                        {
                            tmp.Add("presence", new List<string> { m });
                        }
                    }
                    else if (item.employee.infos.clockin > DateTime.MinValue && item.employee.infos.clockout > DateTime.MinValue && item.employee.infos.hours < (8 - absenceHours))
                    {
                        //時數未滿 ( 8 - 當天請假時數 ) 
                        //檢查是否遲到
                        string late = item.employee.CheckClockInTime(exceptionCases) ? "" : "遲到、";
                        string m = $"{item.dateStr} {late}時數未滿 {8 - absenceHours} 小時";
                        if (tmp.ContainsKey("presence"))
                        {
                            tmp["presence"].Add(m);
                        }
                        else
                        {
                            tmp.Add("presence", new List<string> { m });
                        }
                    }
                    else if (item.employee.infos.clockin == DateTime.MinValue)
                    {
                        //無上班紀錄
                        string m = $"{item.dateStr} 無上班紀錄";
                        if (tmp.ContainsKey("presence"))
                        {
                            tmp["presence"].Add(m);
                        }
                        else
                        {
                            tmp.Add("presence", new List<string> { m });
                        }
                    }
                    else if (item.employee.infos.clockout == DateTime.MinValue)
                    {
                        //無下班紀錄
                        //檢查是否遲到
                        string late = item.employee.CheckClockInTime(exceptionCases) ? "" : "遲到、";
                        string m = $"{item.dateStr} {late}無下班紀錄";
                        if (tmp.ContainsKey("presence"))
                        {
                            tmp["presence"].Add(m);
                        }
                        else
                        {
                            tmp.Add("presence", new List<string> { m });
                        }
                    }
                    else if (!item.employee.CheckClockInTime(exceptionCases))
                    {
                        string m = $"{item.dateStr} 遲到";
                        if (tmp.ContainsKey("presence"))
                        {
                            tmp["presence"].Add(m);
                        }
                        else
                        {
                            tmp.Add("presence", new List<string> { m });
                        }
                    }


                    #endregion

                    #region 請假
                    if (absenceHours > 0)
                    {
                        string m = $"{item.dateStr} ({absenceHours})";
                        if (tmp.ContainsKey("absence"))
                        {
                            tmp["absence"].Add(m);
                        }
                        else
                        {
                            tmp.Add("absence", new List<string> { m });
                        }
                    }
                    #endregion

                    #region 加班
                    if (item.employee.infos.overtime > 0)
                    {
                        string m = $"{item.dateStr} ({item.employee.infos.overtime})";
                        if (tmp.ContainsKey("overtime"))
                        {
                            tmp["overtime"].Add(m);
                        }
                        else
                        {
                            tmp.Add("overtime", new List<string> { m });
                        }
                    }
                    #endregion
                }
                Result.presenceMessage = tmp.ContainsKey("presence") ? $"{tmp["presence"].Count}天異常\n{string.Join("\n", tmp["presence"])}" : "";
                Result.absenceMessage = tmp.ContainsKey("absence") ? $"{tmp["absence"].Count}天\n{string.Join("\n", tmp["absence"])}" : "";
                Result.overtimeMessage = tmp.ContainsKey("overtime") ? $"{tmp["overtime"].Count}天\n{string.Join("\n", tmp["overtime"])}" : "";

                return Result;
            } 
            #endregion
        } 
        #endregion
    }

    public enum LogMessagesType
    {
        info,
        error,
    }

    public static class ExceptionHelper
    {
        public static int LineNumber(this Exception e)
        {

            int linenum = 0;
            try
            {
                linenum = Convert.ToInt32(e.StackTrace.Substring(e.StackTrace.LastIndexOf(' ')));
            }
            catch
            {
                //Stack trace is not available!
            }
            return linenum;
        }
    }


}
