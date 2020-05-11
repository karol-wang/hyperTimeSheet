using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Text.RegularExpressions;

namespace hyperTimeSheet
{
    class Program
    {
        static string root = System.Environment.CurrentDirectory;
        static int thisYear = DateTime.Now.Year;
        static int thisMonth = 0;
        static string log = null;
        static readonly string title = $"HyperTimeSheet Ver 2.1";
        static presenceModel presenceModel = null;
        static absenceModel absenceModel = null;

        static void Main(string[] args)
        {
            Console.Title = title;

            #region initialization 
            try
            {
                ini_Log();
                Log(LogMessagesType.info, $"{string.Empty.PadLeft(10, '-')} Process Started {string.Empty.PadRight(10, '-')}");
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
                presenceModel.ExceptionCases = GetExceptionCase();
            }
            catch (Exception e)
            {
                Console.BackgroundColor = ConsoleColor.White;
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("糟了！程式被你弄壞了！");
                Console.ResetColor();
                Log(LogMessagesType.error, $"初使化發生錯誤。{e.GetFullStackTracesString()}"); 
            }
            #endregion

            try
            {
                absenceModel = GetAbsenceModel(absenceModel);
                presenceModel = GetPresenceModel(presenceModel, absenceModel);
            }
            catch (Exception e)
            {
                Console.BackgroundColor = ConsoleColor.White;
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("糟了！程式被你弄壞了！");
                Console.ResetColor();
                Log(LogMessagesType.error, $"建立物件發生錯誤。原因：{e.GetFullStackTracesString()}");
            }

            try
            {
                ExportExcel(presenceModel);
            }
            catch(Exception e)
            {
                Console.BackgroundColor = ConsoleColor.White;
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("糟了！程式被你弄壞了！");
                Console.ResetColor();
                Log(LogMessagesType.error, $"匯出Excel發生錯誤。原因：{e.GetFullStackTracesString()}");
            }

            Log(LogMessagesType.info, $"{string.Empty.PadLeft(10, '-')} Process Completed {string.Empty.PadRight(10, '-')}");
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
        private static void Log(LogMessagesType type, string logMessage, ExceptionHelper.StackTraceModel stackTrace = null)
        {
            if (string.IsNullOrEmpty(log))
            {
                throw new Exception("Log物件尚未初使化");
            }
            StreamWriter sw = File.AppendText(log);
            if (stackTrace != null)
            {
                sw.WriteLine($"{DateTime.Now:yyyy-MM-dd  hh:mm:ss} | {type} at {stackTrace?.namespace_}.{stackTrace?.method} in {stackTrace?.fileName}:Line {stackTrace?.line} | {logMessage.Replace('\r', ' ')}");
            }
            else
            {
                sw.WriteLine($"{DateTime.Now:yyyy-MM-dd  hh:mm:ss} | {type} | {logMessage.Replace('\r', ' ')}");
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
            StreamReader stream = new StreamReader(filePath);
            string line;
            while ((line = stream.ReadLine()) != null)
            {
                try
                {
                    line = line.Trim();
                    if (line.StartsWith("#"))
                    {
                        continue;
                    }
                    string[] tmp = line.Split(',');
                    //Ex: 0001,9:30
                    //Ex: 0001,0005,0009,9:00
                    if (tmp.Length < 2)
                    {
                        throw new Exception("Exception case 格式有誤");
                    }

                    if (!DateTime.TryParse(tmp[tmp.Length - 1], out DateTime time))
                    {
                        throw new Exception($"{tmp[tmp.Length - 1]} 無法轉換時間");
                    }

                    for (int i = 0; i < tmp.Length - 1; i++)
                    {
                        //最後一筆為時間，不應以迴圈讀取
                        list.Add(new exceptionCaseModel(tmp[i], default(DateTime).Add(time.TimeOfDay)));
                    }
                }
                catch (Exception e)
                {
                    string message = $"\"{line}\"已略過讀取。原因：";
                    Log(LogMessagesType.error, $"{message}{e.GetFullStackTracesString()}");
                    Console.SetCursorPosition(0, Console.CursorTop);
                    Console.WriteLine($"{message}{e.Message}");
                }
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
                List<string> files = Directory.GetFiles(absenceModel.Path).ToList();

                string fileNamePattern = model.FileRegex;
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
                            if (model.Month != m || thisYear != y)
                            {
                                Log(LogMessagesType.info, $"{currentFile}不屬於{thisYear}.{model.Month}月，已略過讀取。");
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
                                absenceModel.Employee employee = model.Employees.SingleOrDefault(x => x.No == No);
                                //absenceModel.DateInfo dateinfo = model.Dates.SingleOrDefault(x => x.date == date);
                                if (employee == null)
                                {
                                    //model中若沒有 employee => 新增
                                    employee = new absenceModel.Employee { No = No, Ename = name };
                                    employee.Dates.Add(new absenceModel.DateInfo { Date = date, Hours = hours });
                                    model.Employees.Add(employee);
                                }
                                else
                                {
                                    employee.Dates.Add(new absenceModel.DateInfo { Date = date, Hours = hours });
                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        //僅略過，不中斷
                        string message = $"讀取假單發生錯誤，已略過讀取{currentFile}。原因：";
                        Log(LogMessagesType.error, $"{message}{e.GetFullStackTracesString()}");
                        Console.SetCursorPosition(0, Console.CursorTop);
                        Console.WriteLine($"{message}{e.Message}");
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
                ExceptionDispatchInfo.Capture(e.InnerException ?? e).Throw();
            }
            return model;
        }
        #endregion

        #region 讀取TimeSheet
        /// <summary>
        /// 讀取指定路徑TimeSheet(xls|xlsx)，並回傳presenceModel
        /// </summary>
        /// <param name="presence"></param>
        /// <returns></returns>
        private static presenceModel GetPresenceModel(presenceModel presence, absenceModel absence)
        {
            string currentFile = "";
            try
            {
                #region 讀取Excel
                List<string> exts = presenceModel.FileExts;
                List<string> formats = presence.FileFormats;
                List<string> files = Directory.GetFiles(presenceModel.Path).ToList();

                string fileNamePattern = presence.FileRegex;
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
                            ISheet sheet = book.GetSheet($"{presence.Month}月份");
                            
                            #region 若序列名稱找不到sheet(前後有空白)，則逐筆搜尋sheet，比對去空白後的sheet名稱
                            if (sheet == null)
                            {
                                IEnumerator sheets = book.GetEnumerator();
                                while (sheets.MoveNext())
                                {
                                    ISheet sheet1 = (ISheet)sheets.Current;
                                    if (sheet1.SheetName.Trim() == $"{presence.Month}月份")
                                    {
                                        sheet = sheet1;
                                        break;
                                    }
                                }
                            }
                            #endregion

                            presenceModel.Employee emp = new presenceModel.Employee(No)
                            {
                                No = No,
                                Name = name
                            };

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
                                    int colorCode = row.GetCell(0).CellStyle.FillForegroundColor;
                                    // TimeSheet內容
                                    //0 出勤日期  1 上班  2 下班  3 正常時數  4 加班時數   5 on-site單位  6 說明(假日, 休假, 病假, 事假, 其他)
                                    DateTime date = getValue<DateTime>(row.Cells[0]);
                                    //取得當月所有日期資訊
                                    //var dateInfo = presenceModel.DateInfos.SingleOrDefault(x => x.Date == date);
                                    var dateInfo = presenceModel.DateInfos.Select(x =>
                                        {
                                            return new presenceModel.DateInfo { Date = x.Date, Note = x.Note, Week = x.Week, WorkingDay = x.WorkingDay };
                                        }
                                    ).SingleOrDefault(x => x.Date == date);
                                    var absenceEmp = absence.Employees.SingleOrDefault(x => x.No == No);
                                    if (dateInfo.WorkingDay)
                                    {
                                        #region columns to model.dates.dateinfo.employee.info
                                        dateInfo.Infos.Clockin = getValue<DateTime>(row.Cells[1]);
                                        dateInfo.Infos.Clockout = getValue<DateTime>(row.Cells[2]);
                                        dateInfo.Infos.Hours = getValue<double>(row.Cells[3]);
                                        dateInfo.Infos.Overtime = getValue<double>(row.Cells[4]);
                                        dateInfo.Infos.Onsite = getValue<string>(row.Cells[5]);
                                        dateInfo.Infos.Note = getValue<string>(row.Cells[6]);
                                        dateInfo.Infos.AbsenceHours = absenceEmp?.Dates.SingleOrDefault(x => x.Date == date)?.Hours ?? 0;

                                        emp.Dates.Add(dateInfo);
                                        #endregion
                                    }
                                }
                            }
                            presence.Employees.Add(emp);
                        }
                    }
                    catch (Exception e)
                    {
                        //僅略過，不中斷
                        string message = $"讀取TimeSheet發生錯誤，已略過讀取{currentFile}。原因：";
                        Log(LogMessagesType.error, $"{message}{e.GetFullStackTracesString()}");
                        Console.SetCursorPosition(0, Console.CursorTop);
                        Console.WriteLine($"{message}{e.Message}");
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
                ExceptionDispatchInfo.Capture(e.InnerException ?? e).Throw();
            }
            return presence;

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
        private static void ExportExcel(presenceModel presence)
        {
            var wb = new XSSFWorkbook();
            ISheet sheetResult = wb.CreateSheet("測試輸出結果");
            sheetResult.SetColumnWidth(0, 15 * 256);
            for (int i = 1; i <= 3; i++)
            {
                sheetResult.SetColumnWidth(i, 40 * 256);
            }
            sheetResult.CreateFreezePane(0, 1);
            ISheet sheetArrange = wb.CreateSheet("出勤一覽");
            sheetArrange.SetColumnWidth(0, 15 * 256);
            sheetArrange.CreateFreezePane(0, 1);

            try
            {
                #region 出勤一覽
                IRow firstRow = sheetArrange.CreateRow(0);
                #region 格式(粗底細邊(左右)+底色)
                XSSFCellStyle firstRowStyle = (XSSFCellStyle)wb.CreateCellStyle();
                XSSFColor color = new XSSFColor();
                color.SetRgb(new byte[] { (byte)136, (byte)160, (byte)226 });
                firstRowStyle.FillForegroundColorColor = color;
                firstRowStyle.FillPattern = FillPattern.SolidForeground;
                firstRowStyle.BorderBottom = BorderStyle.Thick;
                firstRowStyle.BorderRight = BorderStyle.Thin;
                firstRowStyle.BorderLeft = BorderStyle.Thin;
                #endregion
                int count = presence.GetWorkingDays();
                firstRow.CreateCell(0).SetCellValue($"上班日共{count}日");
                firstRow.GetCell(0).CellStyle = firstRowStyle;
                int index = 1;
                foreach (var date in presenceModel.DateInfos)
                {
                    if (date.WorkingDay)
                    {
                        firstRow.CreateCell(index).SetCellValue($"{date.Date:MM-dd}({date.Week})");
                        firstRow.GetCell(index).CellStyle = firstRowStyle;
                        sheetArrange.SetColumnWidth(index, 11 * 256);
                        proccess($"寫入工作天", index, count);
                        index++;
                    }
                }

                int empCount = presence.GetEmployees();
                foreach (var emp in presence.Employees.Select((v, i) => new { index = i, list = v }))
                {
                    writeEmlpoyeeByDate(sheetArrange, emp.index, emp.list);
                    proccess($"寫入員工出勤紀錄", emp.index + 1, empCount);
                }
                #endregion

                #region 彙總結果
                IRow firstRow2 = sheetResult.CreateRow(0);
                firstRow2.CreateCell(0);
                firstRow2.CreateCell(1).SetCellValue("出勤");
                firstRow2.CreateCell(2).SetCellValue("請假");
                firstRow2.CreateCell(3).SetCellValue("加班");
                for (int i = 0; i <= 3; i++)
                {
                    firstRow2.GetCell(i).CellStyle = firstRowStyle;
                }
                #region 格式(粗底細邊)
                XSSFCellStyle RowStyle = (XSSFCellStyle)wb.CreateCellStyle();
                RowStyle.BorderBottom = BorderStyle.Thick;
                RowStyle.BorderLeft = BorderStyle.Thin;
                RowStyle.BorderRight = BorderStyle.Thin;
                RowStyle.WrapText = true;
                RowStyle.VerticalAlignment = VerticalAlignment.Center;
                #endregion
                foreach (var emp in presence.Employees.Select((v, i) => new { index = i, list = v }))
                {
                    var result = GetResult(emp.list);
                    IRow row = sheetResult.CreateRow(emp.index + 1);
                    row.CreateCell(0).SetCellValue(result.NoAndName);
                    row.CreateCell(1);
                    row.CreateCell(2);
                    row.CreateCell(3).SetCellValue(result.OvertimeMessage == "" ? "無" : result.OvertimeMessage);
                    row.GetCell(0).CellStyle = RowStyle;
                    #region 出勤(1) 異常=>紅字
                    if (result.PresenceMessage != "")
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
                        row.GetCell(1).SetCellValue(result.PresenceMessage);
                        row.GetCell(1).CellStyle = presenceRowStyle;
                    }
                    else
                    {
                        row.GetCell(1).SetCellValue("出勤正常");
                        row.GetCell(1).CellStyle = RowStyle;
                    }
                    #endregion
                    #region 請假(2) =>綠字
                    if (result.AbsenceMessage != "")
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
                        row.GetCell(2).SetCellValue(result.AbsenceMessage);
                        row.GetCell(2).CellStyle = absenceRowStyle;
                    }
                    else
                    {
                        row.GetCell(2).SetCellValue("無");
                        row.GetCell(2).CellStyle = RowStyle;
                    }
                    #endregion
                    row.GetCell(3).CellStyle = RowStyle;
                    proccess("輸出彙整結果", emp.index + 1, empCount);
                }
                #endregion

                var fs = new FileStream($"{thisYear}.{thisMonth}.xlsx", FileMode.Create);
                wb.Write(fs);
                fs.Close();
            }
            catch (Exception e)
            {
                ExceptionDispatchInfo.Capture(e.InnerException ?? e).Throw();
            }

            #region 寫入員工資料(void)
            void writeEmlpoyeeByDate(ISheet sheet, int index, presenceModel.Employee employee)
            {
                //每個人皆5列
                int rowspan = 5;
                int colnumber = 1;
                foreach (var dateInfo in employee.Dates)
                {
                    bool IsFirst = dateInfo.Equals(employee.Dates.First());
                    bool IsNormal = !(dateInfo.Infos.AbsenceHours < 8 && (dateInfo.Infos.Clockin == default || dateInfo.Infos.Clockout == default || (dateInfo.Infos.Hours < 8 - dateInfo.Infos.AbsenceHours)));
                    bool IsAbsent = dateInfo.Infos.AbsenceHours > 0;
                    #region 員工資料格式
                    XSSFCellStyle empStyle = (XSSFCellStyle)wb.CreateCellStyle();
                    empStyle.BorderBottom = BorderStyle.Thick;
                    empStyle.BorderRight = BorderStyle.Thin;
                    empStyle.BorderLeft = BorderStyle.Thin;
                    empStyle.VerticalAlignment = VerticalAlignment.Center;
                    empStyle.WrapText = true;
                    #endregion
                    #region 出勤資訊格式
                    XSSFCellStyle infoStyle = (XSSFCellStyle)wb.CreateCellStyle();
                    infoStyle.BorderBottom = BorderStyle.Thin;
                    infoStyle.BorderRight = BorderStyle.Thin;
                    infoStyle.BorderLeft = BorderStyle.Thin;
                    infoStyle.Alignment = HorizontalAlignment.Left;
                    #endregion
                    #region 出勤資訊第一列
                    XSSFCellStyle firstRowStyle = (XSSFCellStyle)wb.CreateCellStyle();
                    firstRowStyle.BorderBottom = BorderStyle.Thin;
                    firstRowStyle.BorderLeft = BorderStyle.Thin;
                    firstRowStyle.BorderRight = BorderStyle.Thin;
                    #endregion
                    #region 出勤資訊最後一列
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

                    if (!employee.CheckClockInTime(dateInfo))
                    {
                        //判斷是否遲到
                        #region 遲到加紅字
                        XSSFFont redFont = (XSSFFont)wb.CreateFont();
                        redFont.Color = HSSFColor.Red.Index;
                        redFont.FontHeightInPoints = 11;
                        firstRowStyle.SetFont(redFont);
                        #endregion
                    }

                    if (IsAbsent && IsNormal)
                    {
                        //有請假
                        #region 請假加亮綠底色
                        XSSFColor color = new XSSFColor();
                        color.SetRgb(new byte[] { (byte)147, (byte)255, (byte)147 });
                        infoStyle.FillForegroundColorColor = color;
                        infoStyle.FillPattern = FillPattern.SolidForeground;
                        firstRowStyle.FillForegroundColorColor = color;
                        firstRowStyle.FillPattern = FillPattern.SolidForeground;
                        lastRowStyle.FillForegroundColorColor = color;
                        lastRowStyle.FillPattern = FillPattern.SolidForeground;
                        #endregion
                    }
                    else if (!IsNormal)
                    {
                        //出勤異常
                        #region 出勤異常加粉紅底色
                        XSSFColor color = new XSSFColor();
                        color.SetRgb(new byte[] { (byte)255, (byte)147, (byte)147 });
                        infoStyle.FillForegroundColorColor = color;
                        infoStyle.FillPattern = FillPattern.SolidForeground;
                        firstRowStyle.FillForegroundColorColor = color;
                        firstRowStyle.FillPattern = FillPattern.SolidForeground;
                        lastRowStyle.FillForegroundColorColor = color;
                        lastRowStyle.FillPattern = FillPattern.SolidForeground;
                        #endregion
                    }

                    for (int i = 1; i <= rowspan; i++)
                    {
                        //每人跨列的列數 * 第？個人 + 該人第幾列
                        int r = rowspan * index + i;
                        IRow row = sheet.GetRow(r) ?? sheet.CreateRow(r);

                        #region 第一欄cell(0)寫入員編&名字
                        if (IsFirst)
                        {
                            string SpecifiedClockInTime = employee.SpecifiedClockInTime == default ? "" : $"\r({employee.SpecifiedClockInTime:HH:mm})";
                            switch (i)
                            {
                                case 1:
                                    row.CreateCell(0).SetCellValue($"{employee.No}-{employee.Name}{SpecifiedClockInTime}");
                                    sheet.AddMergedRegion(new CellRangeAddress(rowspan * index + 1, rowspan * index + 5, 0, 0));
                                    row.GetCell(0).CellStyle = empStyle;
                                    break;
                                default:
                                    // 2~5
                                    row.CreateCell(0).CellStyle = empStyle;
                                    break;
                                    //case 5:
                                    //    //每人最後一列
                                    //    row.CreateCell(0).CellStyle = empStyle;
                                    //    break;
                            }
                        }
                        #endregion

                        #region 寫入各工作天出勤資料
                        switch (i)
                        {
                            case 1:
                                row.CreateCell(colnumber).SetCellValue(dateInfo.Infos.Clockin == default ? "" : $"{dateInfo.Infos.Clockin:HH:mm}");
                                row.GetCell(colnumber).CellStyle = firstRowStyle;
                                break;
                            case 2:
                                row.CreateCell(colnumber).SetCellValue(dateInfo.Infos.Clockout == default ? "" : $"{dateInfo.Infos.Clockout:HH:mm}");
                                row.GetCell(colnumber).CellStyle = infoStyle;
                                break;
                            case 3:
                                row.CreateCell(colnumber).SetCellValue(dateInfo.Infos.Hours);
                                //row.GetCell(colnumber).SetCellType(CellType.String);
                                row.GetCell(colnumber).CellStyle = infoStyle;
                                break;
                            case 4:
                                row.CreateCell(colnumber).SetCellValue(dateInfo.Infos.Onsite);
                                row.GetCell(colnumber).CellStyle = infoStyle;
                                break;
                            case 5:
                                string absenceString = dateInfo.Infos.AbsenceHours == 0 ? "" : $"\n(請假:{dateInfo.Infos.AbsenceHours})";
                                row.CreateCell(colnumber).SetCellValue($"{dateInfo.Infos.Note}{absenceString}");
                                row.GetCell(colnumber).CellStyle = lastRowStyle;
                                break;
                        }
                        #endregion
                    }
                    colnumber++;
                }
            }
            #endregion

            #region 彙整員工當月出勤結果(void)
            ResultModel GetResult(presenceModel.Employee employee)
            {
                if (employee != null)
                {
                    var Result = new ResultModel
                    {
                        NoAndName = employee.SpecifiedClockInTime == default ? $"{employee.No}-{employee.Name}" : $"{employee.No}-{employee.Name}\n({employee.SpecifiedClockInTime:HH:mm})"
                    };
                    var tmp = new Dictionary<string, List<string>>();
                    //Key：presence、absence、overtime
                    //處理單一員工當月每個工作天
                    foreach (var dateInfo in employee.Dates)
                    {
                        string DateAndWeek = $"{dateInfo.Date:MM/dd}({dateInfo.Week})";
                        bool IsInTime = employee.CheckClockInTime(dateInfo);

                        #region 出勤
                        if (dateInfo.Infos.AbsenceHours >= 8)
                        {
                            //請假時數大於等於8 => 請全天不處理
                        }
                        else if (dateInfo.Infos.Clockin == default && dateInfo.Infos.Clockout == default)
                        {
                            //無上班 & 下班紀錄 (未請假)
                            string m = $"{DateAndWeek} 未請假";
                            if (tmp.ContainsKey("presence"))
                            {
                                tmp["presence"].Add(m);
                            }
                            else
                            {
                                tmp.Add("presence", new List<string> { m });
                            }
                        }
                        else if (dateInfo.Infos.Clockin > default(DateTime) && dateInfo.Infos.Clockout > default(DateTime) && dateInfo.Infos.Hours < (8 - dateInfo.Infos.AbsenceHours))
                        {
                            //時數未滿 ( 8 - 當天請假時數 ) 
                            //檢查是否遲到
                            string late = IsInTime ? "" : "遲到、";
                            string m = $"{DateAndWeek} {late}時數未滿 {8 - dateInfo.Infos.AbsenceHours} 小時";
                            if (tmp.ContainsKey("presence"))
                            {
                                tmp["presence"].Add(m);
                            }
                            else
                            {
                                tmp.Add("presence", new List<string> { m });
                            }
                        }
                        else if (dateInfo.Infos.Clockin == default)
                        {
                            //無上班紀錄
                            string m = $"{DateAndWeek} 無上班紀錄";
                            if (tmp.ContainsKey("presence"))
                            {
                                tmp["presence"].Add(m);
                            }
                            else
                            {
                                tmp.Add("presence", new List<string> { m });
                            }
                        }
                        else if (dateInfo.Infos.Clockout == default)
                        {
                            //無下班紀錄
                            //檢查是否遲到
                            string late = IsInTime ? "" : "遲到、";
                            string m = $"{DateAndWeek} {late}無下班紀錄";
                            if (tmp.ContainsKey("presence"))
                            {
                                tmp["presence"].Add(m);
                            }
                            else
                            {
                                tmp.Add("presence", new List<string> { m });
                            }
                        }
                        else if (!IsInTime)
                        {
                            string m = $"{DateAndWeek} 遲到";
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
                        if (dateInfo.Infos.AbsenceHours > 0)
                        {
                            string m = $"{DateAndWeek} ({dateInfo.Infos.AbsenceHours})";
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
                        if (dateInfo.Infos.Overtime > 0)
                        {
                            string m = $"{DateAndWeek} ({dateInfo.Infos.Overtime})";
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
                    Result.PresenceMessage = tmp.ContainsKey("presence") ? $"{tmp["presence"].Count}天異常\n{string.Join("\n", tmp["presence"])}" : "";
                    Result.AbsenceMessage = tmp.ContainsKey("absence") ? $"{tmp["absence"].Count}天\n{string.Join("\n", tmp["absence"])}" : "";
                    Result.OvertimeMessage = tmp.ContainsKey("overtime") ? $"{tmp["overtime"].Count}天\n{string.Join("\n", tmp["overtime"])}" : "";

                    return Result;
                }
                else
                {
                    throw new Exception($"彙整結果發生問題。原因：員工出勤紀錄為空");
                }
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

        /// <summary>
        /// 取得例外Full Stack Traces
        /// </summary>
        /// <param name="e"></param>
        /// <returns></returns>
        public static List<StackTraceModel> GetStackTraces (this Exception e)
        {
            List<StackTraceModel> StackTraces = new List<StackTraceModel>();
            foreach (var item in (new StackTrace(e, true)).GetFrames())
            {
                int line = item.GetFileLineNumber();
                string method = $"{item.GetMethod().Name}({string.Join(",", item.GetMethod().GetParameters().Select(x => x.ParameterType.Name))})";
                string namespace_ = item.GetMethod().DeclaringType.FullName;
                string fileName = item.GetFileName()?.Split('\\').LastOrDefault() ?? "";
                //Console.WriteLine($"at {namespace_}.{method} in {fileName}:Line {line}");
                StackTraces.Add(new StackTraceModel
                {
                    fileName = fileName,
                    line = line,
                    method = method,
                    namespace_ = namespace_
                });
            }
            return StackTraces;
        }

        /// <summary>
        /// 取得Full Stack Traces的訊息字串
        /// </summary>
        /// <param name="e"></param>
        /// <returns></returns>
        public static string GetFullStackTracesString(this Exception e)
        {
            List<string> m = new List<string> { $"\"{e.Message}\"" };
            foreach (var item in (new StackTrace(e, true)).GetFrames())
            {
                if (item.GetMethod().Name != "Throw" && item.GetMethod().DeclaringType.Name != "ExceptionDispatchInfo")
                {
                    int line = item.GetFileLineNumber();
                    string method = $"{item.GetMethod().Name}({string.Join(",", item.GetMethod().GetParameters().Select(x => x.ParameterType.Name))})";
                    string namespace_ = item.GetMethod().DeclaringType.FullName;
                    string fileName = item.GetFileName()?.Split('\\').LastOrDefault() ?? "";
                    m.Add($" at {namespace_}.{method} in {fileName}:Line {line}");
                }
            }
            return string.Join("\n", m);
        }

        public class StackTraceModel
        {
            public string method;
            public string namespace_;
            public string fileName;
            public int line = 0;
        }
    }
}
