using AZReport.Model;
using AZReport.Services;
using AZReport.Services.IServices;
using AZReport.ViewModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace AZReport
{
    public partial class Form1 : Form
    {
        IWorkbook _iWorkbook;
        IProgramService _iProgramService;
        IScheduleService _iScheduleService;
        ISaleService _iSaleService;
        IReportService _iReportService;
        ILevelService _iLevelService;
        ITimeSettingService _iTimeSettingService;
        DateTime start;
        DateTime end;
        DateTime reportStart;
        DateTime reportEnd;
        IFormatProvider culture;
        List<ProductivityViewModel> totalResult;
        DateTime minDate;
        DateTime maxDate;
        int mode;

        public Form1(IProgramService iProgramService, IScheduleService iScheduleService, ISaleService iSaleService,
            IReportService iReportService, ILevelService iLevelService, ITimeSettingService iTimeSettingService)
        {
            InitializeComponent();
            _iProgramService = iProgramService;
            _iScheduleService = iScheduleService;
            _iSaleService = iSaleService;
            _iReportService = iReportService;
            _iLevelService = iLevelService;
            _iTimeSettingService = iTimeSettingService;
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker3.Format = DateTimePickerFormat.Custom;
            dateTimePicker4.Format = DateTimePickerFormat.Custom;            
            dateTimePicker6.Format = DateTimePickerFormat.Custom;
            dateTimePicker7.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "dd-MM-yyyy";
            dateTimePicker1.CustomFormat = "dd-MM-yyyy";
            dateTimePicker2.CustomFormat = "dd-MM-yyyy";
            dateTimePicker3.CustomFormat = "dd-MM-yyyy";
            dateTimePicker4.CustomFormat = "dd-MM-yyyy";
            dateTimePicker6.CustomFormat = "dd-MM-yyyy";
            dateTimePicker7.CustomFormat = "dd-MM-yyyy";

            dateTimePicker5.Format = DateTimePickerFormat.Time;
            dateTimePicker5.ShowUpDown = true;

            var all = _iProgramService.GetAll().ToList();
            start = DateTime.Today;
            end = DateTime.Today;
            reportStart = DateTime.Today;
            reportEnd = DateTime.Today;
            culture = new System.Globalization.CultureInfo("fr-FR", true);
            var levelList = _iLevelService.GetAll().ToList();
            var a = levelList.Where(x => x.Name == "A").OrderByDescending(x=>x.UpdateDate).FirstOrDefault();
            if (a != null && a.Begin.HasValue) { 
                textBox4.Text = a.Begin.Value.ToString("N", CultureInfo.CreateSpecificCulture("en-US"));
                textBox7.Text = a.End.Value.ToString("N", CultureInfo.CreateSpecificCulture("en-US"));

                var b = levelList.Where(x => x.Name == "B").OrderByDescending(x => x.UpdateDate).FirstOrDefault();
                textBox5.Text = b.Begin.Value.ToString("N", CultureInfo.CreateSpecificCulture("en-US"));
                textBox8.Text = b.End.Value.ToString("N", CultureInfo.CreateSpecificCulture("en-US"));
                var c = levelList.Where(x => x.Name == "C").OrderByDescending(x => x.UpdateDate).FirstOrDefault();
                textBox6.Text = c.Begin.Value.ToString("N", CultureInfo.CreateSpecificCulture("en-US"));
                textBox9.Text = c.End.Value.ToString("N", CultureInfo.CreateSpecificCulture("en-US"));
            }
            var timesetting = _iTimeSettingService.GetAll().OrderByDescending(x=>x.UpdateDate).FirstOrDefault();
            if (timesetting != null)
            {
                dateTimePicker5.Value = timesetting.time;
            }
            else
            {
                dateTimePicker5.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 6, 0, 0);
            }
            changeTimeGraph();

            ComboboxItem item1 = new ComboboxItem();
            item1.Text = "Quantity";
            item1.Value = 1;
            ComboboxItem item2 = new ComboboxItem();
            item2.Text = "Efficiency";
            item2.Value = 2;
            comboBox1.Items.Add(item1);
            comboBox1.Items.Add(item2);
            comboBox1.SelectedIndex = 0;
            mode = 1;
        }

        public class ComboboxItem
        {
            public string Text { get; set; }
            public object Value { get; set; }

            public override string ToString()
            {
                return Text;
            }
        }

        private void changeTimeGraph()
        {
            minDate = new DateTime(dateTimePicker6.Value.Year, dateTimePicker6.Value.Month, dateTimePicker6.Value.Day).AddSeconds(-1);
            maxDate = new DateTime(dateTimePicker7.Value.Year, dateTimePicker7.Value.Month, dateTimePicker7.Value.Day); // or DateTime.Now;
            checkedListBox1.Items.Clear();           
            var result = _iReportService.GetProductivity(new DateTime(minDate.Year, minDate.Month, minDate.Day, 0, 0, 0), new DateTime(maxDate.Year, maxDate.Month, maxDate.Day, 23, 59, 59));               
            foreach (var i in result)
            {
                if (Convert.ToDateTime(i.Duration).Minute > 4)
                {
                    checkedListBox1.Items.Add(i.Code + "-" + i.Name + "-" + i.Note, false);
                }
            }
        }

        private void ReadProgram(IWorkbook _iWordbook)
        {
            for (int i = 0; i < _iWorkbook.NumberOfSheets; i++)
            {
                if (_iWorkbook.GetSheetName(i).Equals("MABANG"))
                {
                    ISheet proSheet = _iWorkbook.GetSheetAt(i);
                    List<Program> listProgram = new List<Program>();
                    for (int j = 4; j <= proSheet.LastRowNum; j++)
                    {
                        var row = proSheet.GetRow(j);
                        Program program = new Program();
                        if (row.GetCell(0) == null || row.GetCell(0).NumericCellValue == 0)
                            break;
                        program.Name = row.GetCell(1).StringCellValue.ToString();
                        program.Duration = row.GetCell(2).DateCellValue.ToString();
                        program.Code = row.GetCell(3).StringCellValue.ToString();
                        program.Category = row.GetCell(4).StringCellValue.ToString();
                        program.Price = row.GetCell(5).NumericCellValue.ToString();
                        program.Note = row.GetCell(6).StringCellValue.ToString();
                        if (checkBox1.Checked)
                        {
                            _iProgramService.CheckAndUpdate(program);
                        }
                        else
                        {
                            _iProgramService.CheckAndCreate(program);
                        }
                    }
                    _iProgramService.Save();
                }
            }
        }

        private void ReadSchedule(IWorkbook _iWordbook)
        {
            for (int i = 0; i < _iWorkbook.NumberOfSheets; i++)
            {
                ISheet sheet = _iWorkbook.GetSheetAt(i);
                if (CheckSchedule(sheet))
                {
                    var listTime = GetTimeList(sheet);
                    var startDay = listTime.FirstOrDefault();
                    var lastDay = listTime.LastOrDefault();
                    var startPoint = StartPoint(sheet);
                    while (startDay <= lastDay)
                    {
                        if (checkBox2.Checked)
                        {
                            //delete old data
                            if (_iScheduleService.CheckExistDate(startDay))
                            {
                                _iScheduleService.DeleteOldDate(startDay);
                            }
                            // add update data
                            for (int j = startPoint.First() + 1; j <= sheet.LastRowNum; j++)
                            {
                                var row = sheet.GetRow(j);
                                Schedule schedule = new Schedule();
                                if (row.GetCell(startPoint.Last()) == null || row.GetCell(startPoint.Last()).NumericCellValue == 0)
                                    break;
                                schedule.Code = row.GetCell(startPoint.Last() + 4).StringCellValue.ToString();
                                var mytime = row.GetCell(startPoint.Last() + 1).DateCellValue;
                                schedule.Date = new DateTime(startDay.Year, startDay.Month, startDay.Day, mytime.Hour, mytime.Minute, mytime.Second);
                                _iScheduleService.Create(schedule);
                            }
                        }
                        else
                        {
                            if (_iScheduleService.CheckExistDate(startDay))
                            {
                                //do nothing
                            }
                            else
                            {
                                for (int j = startPoint.First() + 1; j <= sheet.LastRowNum; j++)
                                {
                                    var row = sheet.GetRow(j);
                                    Schedule schedule = new Schedule();
                                    if (row.GetCell(startPoint.Last()) == null || row.GetCell(startPoint.Last()).NumericCellValue == 0)
                                        break;
                                    schedule.Code = row.GetCell(startPoint.Last() + 4).StringCellValue.ToString();
                                    var mytime = row.GetCell(startPoint.Last() + 1).DateCellValue;
                                    schedule.Date = new DateTime(startDay.Year, startDay.Month, startDay.Day, mytime.Hour, mytime.Minute, mytime.Second);
                                    _iScheduleService.Create(schedule);
                                }
                            }
                        }
                        startDay = startDay.AddDays(1);
                    }
                    _iScheduleService.Save();
                }
            }
        }

        private bool CheckSchedule(ISheet sheet)
        {            
            var IsSheet = false;
            for (int j = 0; j <= 4; j++)
            {
                var row = sheet.GetRow(j);
                if (row != null)
                {
                    for (int k = 0; k < 3; k++)
                    {
                        if (row.GetCell(k) != null && row.GetCell(k).CellType == CellType.String)
                        {
                            var mystring = row.GetCell(k).StringCellValue.ToString();
                            if (!string.IsNullOrWhiteSpace(mystring))
                            {
                                if (mystring.Contains("LỊCH PHÁT SÓNG QUẢNG CÁO"))
                                {
                                    IsSheet = true;
                                    break;
                                }
                            }
                        }

                    }
                    if (IsSheet == true)
                        break;
                }
            }
            return IsSheet;
        }

        private List<DateTime> GetTimeList(ISheet sheet)
        {
            var datestring = "";
            for (int j = 0; j <= 4; j++)
            {
                var row = sheet.GetRow(j);
                if (row != null)
                {
                    for (int k = 0; k < 3; k++)
                    {
                        if (row.GetCell(k) != null && row.GetCell(k).CellType == CellType.String)
                        {
                            var mystring = row.GetCell(k).StringCellValue.ToString();
                            if (!string.IsNullOrWhiteSpace(mystring))
                            {
                                if (mystring.Contains("Ngày phát sóng"))
                                {
                                    datestring = mystring;
                                    break;
                                }
                            }
                        }
                        if (!string.IsNullOrEmpty(datestring))
                        {
                            break;
                        }
                    }

                }
                
            }
            var date = datestring.Split(new string[] { "Ngày phát sóng:" }, StringSplitOptions.None).LastOrDefault();
            List<DateTime> TimeList = new List<DateTime>();
            DateTime startday = new DateTime(1990,1,1);
            DateTime lastday = new DateTime(1990, 1, 1);
            var format = "d/M/yyyy";
            var provider = new CultureInfo("fr-FR");
            if (date.Contains('-'))
            {
                var datelist = date.Split('-');                
                startday = DateTime.ParseExact(datelist.FirstOrDefault().Trim(), format, provider);
                lastday = DateTime.ParseExact(datelist.LastOrDefault().Trim(), format, provider);               
            }
            else if ((date.Contains(',')||date.Contains(';'))&&!date.Contains('&'))
            {
                var datelist = date.Split(new char[]{';',','});
                lastday = DateTime.ParseExact(datelist.LastOrDefault().Trim(), format, provider);
                startday = new DateTime(lastday.Year, lastday.Month, Convert.ToInt32(datelist.FirstOrDefault().Trim()));
            }
            else
            {
                var datelist = date.Split(new char[] { ';', ',','&' });
                if (datelist.Count() == 1)
                {
                    lastday = DateTime.ParseExact(datelist.LastOrDefault().Trim(), format, provider);
                    startday = DateTime.ParseExact(datelist.FirstOrDefault().Trim(), format, provider);
                }
                else if (datelist.Count() > 2)
                {
                    lastday = DateTime.ParseExact(datelist.LastOrDefault().Trim(), format, provider);
                    if (lastday.Month != 1)
                    {
                        startday = new DateTime(lastday.Year, lastday.Month - 1, Convert.ToInt32(datelist.FirstOrDefault().Trim()));
                    }
                    else
                    {
                        startday = new DateTime(lastday.Year-1, 12, Convert.ToInt32(datelist.FirstOrDefault().Trim()));
                    }
                }
                else
                {

                
                }
            }
            if (startday.Year != 1990)
            {
                while (startday <= lastday)
                {
                    TimeList.Add(startday);
                    startday = startday.AddDays(1);
                }
            }       
            return TimeList;
        }

        private List<int> StartPoint(ISheet sheet)
        {
            List<int> result = new List<int>();
            for (int j = 0; j <= 7; j++)
            {
                var row = sheet.GetRow(j);
                if (row != null)
                {
                    for (int k = 0; k < 3; k++)
                    {
                        if (row.GetCell(k) != null && row.GetCell(k).CellType == CellType.String)
                        {
                            var mystring = row.GetCell(k).StringCellValue.ToString();
                            if (!string.IsNullOrWhiteSpace(mystring))
                            {
                                if (mystring.Contains("STT"))
                                {
                                    result.Add(j);
                                    result.Add(k);
                                    break;
                                }
                            }
                        }
                        if (result.Count() != 0)
                        {
                            break;
                        }
                    }

                }

            }
            return result;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx";
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                try
                {
                    string file = openFileDialog1.FileName;
                    textBox1.Text = file;
                    using (FileStream pr = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        if (file.Contains(".xlsx"))
                        {
                            _iWorkbook = new XSSFWorkbook(pr);
                        }
                        else if (file.Contains(".xls"))
                        {
                            _iWorkbook = new HSSFWorkbook(pr);
                        }
                        else
                        {
                            return;
                        }
                    }
                }
                catch (Exception ex)
                {
                    var errorForm = new ErrorForm(ex.Message);
                    errorForm.ShowDialog();
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog2.Filter = "Excel Files|*.xls;*.xlsx";
            DialogResult result = openFileDialog2.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                try
                {
                    string file = openFileDialog2.FileName;
                    textBox2.Text = file;
                    using (FileStream pr = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        if (file.Contains(".xlsx"))
                        {
                            _iWorkbook = new XSSFWorkbook(pr);
                        }
                        else if (file.Contains(".xls"))
                        {
                            _iWorkbook = new HSSFWorkbook(pr);
                        }
                        else
                        {
                            return;
                        }
                    }
                }
                catch (Exception ex)
                {
                    var errorForm = new ErrorForm(ex.Message);
                    errorForm.ShowDialog();
                }
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            start = dateTimePicker1.Value;
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            end = dateTimePicker2.Value;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Excel|*.xls;*.xlsx";
            string title = "";
            if (dateTimePicker3.Value.ToShortDateString() != dateTimePicker4.Value.ToShortDateString())
            {
                title = "AZ_Quantity_" + dateTimePicker1.Value.Day.ToString() + "~" + dateTimePicker2.Value.ToString("d.MM.yyyy");
            }
            else
            {
                title = "AZ_Quantity_" + dateTimePicker2.Value.ToString("d.MM.yyyy");
            }
            saveFileDialog1.FileName = title;
            saveFileDialog1.DefaultExt = "xlsx";
            saveFileDialog1.ShowDialog();    
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            try
            {
                var result = _iReportService.GetProductivity(new DateTime(start.Year, start.Month, start.Day, 0, 0, 0), new DateTime(end.Year, end.Month, end.Day, 23, 59, 59));
                var quantityViewModel = new List<SaleViewModel>();
                foreach (var productivity in result)
                {
                    var temp = new SaleViewModel();
                    temp.Sales = _iSaleService.GetQuantity(new DateTime(start.Year, start.Month, start.Day, 0, 0, 0), new DateTime(end.Year, end.Month, end.Day, 23, 59, 59), productivity.Code);
                    temp.Code = productivity.Code;
                    quantityViewModel.Add(temp);
                }
                string name = saveFileDialog1.FileName;
                var wb = new XSSFWorkbook();
                // tab name
                ISheet sheet = wb.CreateSheet("Bao cao SL ban ra hang ngay");
                // header
                IRow row = sheet.CreateRow(0);
                ICell cell = row.CreateCell(0);
                cell.SetCellValue("BÁO CÁO SẢN PHẨM HÀNG NGÀY CÔNG TY ATZ");
                NPOI.SS.Util.CellRangeAddress cra = new NPOI.SS.Util.CellRangeAddress(0, 0, 0, 2);
                sheet.AddMergedRegion(cra);
                // column header 
                IRow row3 = sheet.CreateRow(2);
                ICell cell0 = row3.CreateCell(0);
                cell0.SetCellValue("STT");
                ICell cell1 = row3.CreateCell(1);
                cell1.SetCellValue("MÃ CHƯƠNG TRÌNH");
                ICell cell2 = row3.CreateCell(2);
                cell2.SetCellValue("CHƯƠNG TRÌNH");
                //ICell cell3 = row3.CreateCell(3);
                //cell3.SetCellValue("DURATION");
                ICell cell4 = row3.CreateCell(3);
                cell4.SetCellValue("CATEGORY");
                //ICell cell5 = row3.CreateCell(5);
                //cell5.SetCellValue("GIÁ SẢN PHẨM");
                ICell cell6 = row3.CreateCell(4);
                cell6.SetCellValue("Ghi chú");
                var tempStart = start;
                var k = 5;
                while (DateTime.Compare(tempStart, end) <= 0)
                {
                    ICell cellk = row3.CreateCell(k);
                    cellk.SetCellValue(tempStart.ToString("dd/MM/yyyy"));
                    tempStart = tempStart.AddDays(1);
                    k++;
                }
                // add Program Code
                int i = 3;
                foreach (var item in result)
                {
                    var time = Convert.ToDateTime(item.Duration);
                    if (time.Minute > 4)
                    {
                        IRow row_temp = sheet.CreateRow(i);
                        ICell cell_temp0 = row_temp.CreateCell(0);
                        cell_temp0.SetCellValue(i - 2);
                        ICell cell_temp1 = row_temp.CreateCell(1);
                        cell_temp1.SetCellValue(item.Code);
                        ICell cell_temp2 = row_temp.CreateCell(2);
                        cell_temp2.SetCellValue(item.Name);
                        //ICell cell_temp3 = row_temp.CreateCell(3);
                        //DateTime time1 = DateTime.Today;
                        //time1 = time1.AddMinutes(time.Minute).AddSeconds(time.Second);
                        //cell_temp3.SetCellValue(time1);
                        //ICellStyle style = wb.CreateCellStyle();
                        //cell_temp3.CellStyle = style;
                        //IDataFormat dataFormatCustom = wb.CreateDataFormat();
                        //cell_temp3.CellStyle.DataFormat = dataFormatCustom.GetFormat("mm:ss");
                        ICell cell_temp4 = row_temp.CreateCell(3);
                        cell_temp4.SetCellValue(item.Category);
                        //ICell cell_temp5 = row_temp.CreateCell(5);
                        //cell_temp5.SetCellValue(item.Price);
                        ICell cell_temp6 = row_temp.CreateCell(4);
                        cell_temp6.SetCellValue(item.Note);
                        var tempStart1 = start;
                        var k1 = 5;
                        var sales = quantityViewModel.Where(x=> x.Code == item.Code).FirstOrDefault();
                        while (DateTime.Compare(tempStart1, end) <= 0)
                        {
                            ICell cellk = row_temp.CreateCell(k1);
                            if (sales != null)
                            {
                                var q = sales.Sales.Where(y => y.Date.Year == tempStart1.Year && y.Date.Month == tempStart1.Month && y.Date.Day == tempStart1.Day).FirstOrDefault();
                                if (q != null)
                                {
                                    cellk.SetCellValue(q.Quantity);
                                }
                                else
                                {
                                    cellk.SetCellValue(0);
                                }
                            }
                            else
                            {
                                cellk.SetCellValue(0);
                            }
                            tempStart1 = tempStart1.AddDays(1);
                            k1++;
                        }
                        i++;
                    }
                }

                for (int l = 0; l < row3.LastCellNum; l++)
                {
                    sheet.AutoSizeColumn(l);
                }

                using (FileStream stream = new FileStream(name, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
                {
                    wb.Write(stream);
                    stream.Close();
                }
                var successForm = new SuccessForm();
                successForm.ShowDialog();
            }
            catch (Exception ex)
            {
                var errorForm = new ErrorForm(ex.Message);
                errorForm.ShowDialog();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            openFileDialog3.Filter = "Excel Files|*.xls;*.xlsx";
            DialogResult result = openFileDialog3.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                try
                {
                    string file = openFileDialog3.FileName;
                    textBox3.Text = file;
                    using (FileStream pr = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        if (file.Contains(".xlsx"))
                        {
                            _iWorkbook = new XSSFWorkbook(pr);
                        }
                        else if (file.Contains(".xls"))
                        {
                            _iWorkbook = new HSSFWorkbook(pr);
                        }
                        else
                        {
                            return;
                        }
                    }
                }
                catch (Exception ex)
                {
                    var errorForm = new ErrorForm(ex.Message);
                    errorForm.ShowDialog();
                }
            }
        }

        private void ReadQuantity(IWorkbook _iWordbook)
        {
            for (int i = 0; i < _iWorkbook.NumberOfSheets; i++)
            {
                if (_iWorkbook.GetSheetName(i).Equals("Bao cao SL ban ra hang ngay"))
                {
                    ISheet proSheet = _iWorkbook.GetSheetAt(i);
                    List<Program> listProgram = new List<Program>();
                    for (int j = 3; j <= proSheet.LastRowNum; j++)
                    {
                        var row = proSheet.GetRow(j);
                        Sale sale = new Sale();
                        if (row.GetCell(0) == null || row.GetCell(0).NumericCellValue == 0)
                            break;
                        sale.Code = row.GetCell(1).StringCellValue.ToString();
                        for (int k = 5; k < proSheet.GetRow(2).LastCellNum; k++)
                        {
                            sale.Quantity = row.GetCell(k) != null ? Convert.ToInt32(row.GetCell(k).NumericCellValue) : 0;
                            sale.Date = DateTime.Parse(proSheet.GetRow(2).GetCell(k).StringCellValue.ToString(), culture, System.Globalization.DateTimeStyles.AssumeLocal);
                            _iSaleService.CheckAndUpdate(sale);
                            _iSaleService.Save();
                        }                        
                    }
                    
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            saveFileDialog2.Filter = "Excel|*.xls;*.xlsx";
            string title = "";
            if (dateTimePicker3.Value.ToShortDateString() != dateTimePicker4.Value.ToShortDateString())
            {
                title = "AZ_Efficiency_" + dateTimePicker3.Value.Day.ToString() + "~" + dateTimePicker4.Value.ToString("d.MM.yyyy");
            }
            else
            {
                title = "AZ_Efficiency_" + dateTimePicker4.Value.ToString("d.MM.yyyy");
            }
            saveFileDialog2.FileName = title;
            saveFileDialog2.DefaultExt = "xlsx";
            saveFileDialog2.ShowDialog();
        }

        private void saveFileDialog2_FileOk(object sender, CancelEventArgs e)
        {
            try
            {
                string name = saveFileDialog2.FileName;
                var wb = new XSSFWorkbook();
                using (FileStream stream = new FileStream(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TemplateEfficiency.xlsx"), FileMode.Open, FileAccess.Read))
                {
                    wb = new XSSFWorkbook(stream);
                    stream.Close();
                }
                ISheet itemListSheet = wb.GetSheetAt(1);
                createItemList(itemListSheet, reportStart, reportEnd);
                ISheet timeTableSheet = wb.GetSheetAt(0);
                createTimeTable(timeTableSheet, reportStart, reportEnd);
                using (FileStream stream = new FileStream(name, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
                {
                    wb.Write(stream);
                    stream.Close();
                }
                var successForm = new SuccessForm();
                successForm.ShowDialog();
            }
            catch (Exception ex)
            {
                var errorForm = new ErrorForm(ex.Message);
                errorForm.ShowDialog();
            }
        }

        private void createItemList(ISheet sheet, DateTime startDay, DateTime endDay)
        {
            ICell topCel3 = sheet.GetRow(0).GetCell(3);
            if (startDay.Year == endDay.Year && startDay.Year == endDay.Year && startDay.Year == endDay.Year)
            {
                topCel3.SetCellValue(startDay.ToString("dd/MM/yyyy"));
            }
            else
            {
                topCel3.SetCellValue(startDay.ToString("dd/MM/yyyy") + " - " + endDay.ToString("dd/MM/yyyy"));
            }

            var startDayTem = new DateTime(startDay.Year, startDay.Month, startDay.Day, 0, 0, 0);            
            IRow dayHeaderList = sheet.GetRow(1);
            for (int m = 0; m < 11; m++)
            {
                ICell dayHeader = dayHeaderList.GetCell(8 + m);
                dayHeader.SetCellValue(startDayTem.ToString("ddd"));
                startDayTem = startDayTem.AddDays(1);
            }
            var timesetting = _iTimeSettingService.GetAll().OrderByDescending(x => x.UpdateDate).FirstOrDefault();
            var time1 = timesetting.time;
            var result = _iReportService.GetProductivity(new DateTime(startDay.Year, startDay.Month, startDay.Day, 0, 0, 0), new DateTime(endDay.Year, endDay.Month, endDay.Day, 23, 59, 59));
            var quantityList = _iReportService.GetQuantity(new DateTime(startDay.Year, startDay.Month, startDay.Day, 0, 0, 0), new DateTime(endDay.Year, endDay.Month, endDay.Day, 23, 59, 59));
            var freqList = _iReportService.GetFreq(new DateTime(startDay.Year, startDay.Month, startDay.Day, 0, 0, 0), new DateTime(endDay.Year, endDay.Month, endDay.Day, 23, 59, 59), time1);
            IRow row2 = sheet.GetRow(1);
            int rowIndex = 2;
            foreach (var item in result)
            {
                DateTime time = Convert.ToDateTime(item.Duration);
                if (time.Minute > 4)
                {
                    IRow rowEff = sheet.GetRow(rowIndex);
                    ICell eff_cell0 = rowEff.GetCell(0);
                    eff_cell0.SetCellValue(rowIndex - 1);
                    ICell eff_cell1 = rowEff.GetCell(1);
                    eff_cell1.SetCellValue(item.Code);
                    ICell eff_cell2 = rowEff.GetCell(2);
                    eff_cell2.SetCellValue(item.Name);
                    ICell eff_cell4 = rowEff.GetCell(4);

                    eff_cell4.SetCellValue(time.Minute+ time.Second/60.0);
                    ICell eff_cell5 = rowEff.GetCell(5);
                    eff_cell5.SetCellValue(item.Category);
                    ICell eff_cell6 = rowEff.GetCell(6);
                    eff_cell6.SetCellValue(item.Price);                   

                    var tempDate = startDay;
                    int l = 1;
                    while (tempDate <= endDay)
                    {
                        ICell eff_cellweek = rowEff.GetCell(7 + l);
                        eff_cellweek.SetCellValue(freqList.Where(x=>x.Code == item.Code && x.Date == new DateTime(startDay.Year, startDay.Month, startDay.Day, 0, 0, 0)).FirstOrDefault() != null? freqList.Where(x=>x.Code == item.Code && x.Date == new DateTime(startDay.Year, startDay.Month, startDay.Day, 0, 0, 0)).FirstOrDefault().Freq: 0);
                        tempDate = tempDate.AddDays(1);
                        l++;
                    }

                    int q = quantityList.Where(x => x.Code == item.Code).FirstOrDefault() != null ?(int)quantityList.Where(x => x.Code == item.Code).FirstOrDefault().Quantity: 0;
                    ICell eff_cell10 = rowEff.CreateCell(27);
                    eff_cell10.SetCellValue(q);
                    ICell eff_cell11 = rowEff.CreateCell(28);
                    var amount = q * Convert.ToInt32(item.Price);
                    eff_cell11.SetCellValue(amount);
                    ICell eff_cell12 = rowEff.CreateCell(29);
                    var totalTime = freqList.Where(x => x.Code == item.Code).Sum(x => x.Freq) * (time.Minute + time.Second / 60.0);
                    eff_cell12.SetCellValue(totalTime);

                    ICell eff_cell7 = rowEff.GetCell(7);
                    eff_cell7.SetCellValue(amount/totalTime);

                    ICell eff_cell13 = rowEff.CreateCell(3);
                    item.Group = calculateGroup(amount / totalTime);
                    eff_cell13.SetCellValue(item.Group);
                    rowIndex++;
                }              
            }

            //for (int l = 0; l < row2.LastCellNum; l++)
            //{
            //    sheet.AutoSizeColumn(l);
            //}

            totalResult = result;
        }

        private void createTimeTable(ISheet sheetTime, DateTime startDay, DateTime endDay)
        {
            #region header
            string title = "";
            if (dateTimePicker3.Value.ToShortDateString() != dateTimePicker4.Value.ToShortDateString())
            {
                title = "Weekly programming SCTV10 channel \n " + dateTimePicker3.Value.Day.ToString() + " ~ " + dateTimePicker4.Value.ToString("d.MM.yyyy");
            }
            else
            {
                title = "Weekly programming SCTV10 channel \n " + dateTimePicker4.Value.ToString("d.MM.yyyy");
            }
            IRow rowTime0 = sheetTime.GetRow(0);
            rowTime0.GetCell(2).SetCellValue(title);
            IRow rowTime1 = sheetTime.GetRow(1);
            IRow rowTime2 = sheetTime.GetRow(2);
            var tempDate = startDay;
            for (int p = 1; p < 12; p++)
            {
                var startWeek = tempDate.ToString("ddd");
                var cellWeek1 = rowTime1.GetCell((p - 1) * 6 + 2);
                var cellWeek2 = rowTime2.GetCell((p - 1) * 6 + 2);
                cellWeek1.SetCellValue(tempDate.ToString("dd.MM"));
                cellWeek2.SetCellValue(startWeek.ToUpper());
                tempDate = tempDate.AddDays(1);
            }
            #endregion
            var timesetting = _iTimeSettingService.GetAll().OrderByDescending(x => x.UpdateDate).FirstOrDefault();
            int beginHour = timesetting.time.Hour;
            int index1 = 4;
            for (int p = beginHour; p < 24; p++)
            {
                IRow rowTime = sheetTime.GetRow(index1);
                if (rowTime == null)
                {
                    rowTime = sheetTime.CreateRow(index1);
                }
                var myCell1 = rowTime.GetCell(1);
                if (myCell1 == null)
                {
                    myCell1 = rowTime.CreateCell(1);
                }
                myCell1.SetCellValue(beginHour);
                var myCell68 = rowTime.GetCell(68);
                if (myCell68 == null)
                {
                    myCell68 = rowTime.CreateCell(68);
                }
                myCell68.SetCellValue(beginHour++);
                index1 = index1 + 4;
            }
            var scheduleList = _iScheduleService.GetByDate(startDay, endDay);
            var timeTemp = startDay;
            int dd = 0;
            while (timeTemp <= endDay)
            {
                var schedulePerDay = scheduleList.Where(x => x.Date.ToShortDateString() == timeTemp.ToShortDateString()).ToList();
                schedulePerDay = schedulePerDay.Where(x=>x.Date.TimeOfDay > Convert.ToDateTime(dateTimePicker5.Value).TimeOfDay).OrderBy(x=>x.Date).ToList();
                List<ScheduleViewModel> groupHour = GroupHour(schedulePerDay).OrderBy(x => x.Hour).ToList();
                int index = 4;
                foreach (var i in groupHour)
                {   
                    // add hour
                    // add program
                    foreach (var program in i.programs)
                    {
                        var item = totalResult.Where(x => x.Code == program.Code).FirstOrDefault();
                        if (Convert.ToDateTime(item.Duration).Minute > 4)
                        {
                            IRow rowTime4 = sheetTime.GetRow(index++);
                            if (rowTime4 == null)
                            {
                                rowTime4 = sheetTime.CreateRow(index++);
                            }
                            var myCell = rowTime4.GetCell(7 + dd * 6);
                            if (myCell == null)
                            {
                                myCell = rowTime4.CreateCell(7 + dd * 6);
                            }
                            myCell.SetCellValue(item.Name);
                            var myCell2 = rowTime4.GetCell(3 + dd * 6);
                            if (myCell2 == null)
                            {
                                myCell2 = rowTime4.CreateCell(3 + dd * 6);
                            }
                            myCell2.SetCellValue(Convert.ToDateTime(item.Duration).Minute);
                          
                            var myCell4 = rowTime4.GetCell(4 + dd * 6);
                            if (myCell4 == null)
                            {
                                myCell4 = rowTime4.CreateCell(4 + dd * 6);
                            }
                            myCell4.SetCellValue(item.Group);
                            var myCell3 = rowTime4.GetCell(5 + dd * 6);
                            if (myCell3 == null)
                            {
                                myCell3 = rowTime4.CreateCell(5 + dd * 6);
                            }
                            myCell3.SetCellValue(item.Category);
                            //var myCell5 = rowTime4.GetCell(1 + dd * 6);
                            //if (myCell5 == null)
                            //{
                            //    myCell5 = rowTime4.CreateCell(1 + dd * 6);
                            //}
                            //myCell5.SetCellValue(Convert.ToDateTime(i.Date).Hour);
                        }
                    }
                    if ((index - 4) % 4 != 0)
                    {
                        int tem = (index - 4) / 4;
                        index = (tem + 1) * 4 + 4;
                    }
                }
                timeTemp = timeTemp.AddDays(1);
                dd++;

            }          
        }

        private List<ScheduleViewModel> GroupHour(List<Schedule> originalList)
        {
            List<ScheduleViewModel> result = new List<ScheduleViewModel>();
            foreach (var item in originalList)
            {
                addItemToScheduleList(item, result);
            }
            return result;
        }

        private List<ScheduleViewModel> addItemToScheduleList(Schedule item, List<ScheduleViewModel> result)
        {           
            var item_temp = totalResult.Where(x => x.Code == item.Code).FirstOrDefault();
            DateTime time = Convert.ToDateTime(item_temp.Duration);
            if (time.Minute > 4)
            {
                DateTime endTime = item.Date;
                endTime = endTime.AddMinutes(time.Minute).AddSeconds(time.Second);
                DateTime seprateTime = new DateTime(endTime.Year, endTime.Month, endTime.Day, endTime.Hour, 0, 0);
                DateTime choosenTime = new DateTime();
                if (item.Date <= seprateTime)
                {
                    if ((seprateTime - item.Date) >= (endTime - seprateTime))
                    {
                        choosenTime = item.Date;
                    }
                    else
                    {
                        choosenTime = endTime;
                    }
                }
                else
                    choosenTime = endTime;
                var choosenItem = result.Where(x => x.Hour == choosenTime.Hour).FirstOrDefault();
                if (choosenItem != null)
                {
                    choosenItem.programs.Add(item);
                    choosenItem.programs = choosenItem.programs.OrderBy(x => x.Date).ToList();
                }
                else
                {
                    choosenItem = new ScheduleViewModel();
                    choosenItem.Hour = choosenTime.Hour;
                    choosenItem.programs = new List<Schedule>();
                    choosenItem.programs.Add(item);
                    result.Add(choosenItem);
                }
            }
            return result;
        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {
            reportStart = dateTimePicker3.Value;
        }

        private void dateTimePicker4_ValueChanged(object sender, EventArgs e)
        {
            reportEnd = dateTimePicker4.Value;
        }

        private string calculateGroup(double eff)
        {
            var levelList = _iLevelService.GetAll().ToList();
            foreach (var i in levelList)
            {
                if (i.Begin != null)
                {
                    if ((int)i.Begin <= eff && (int)i.End >= eff)
                    {
                        return i.Name;
                    }
                }
            }
            return "";
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                var levelList = _iLevelService.GetAll().ToList();
               
                var a = new Level();
                a.Name = "A";
                a.UpdateDate = DateTime.Now;
                a.Begin = (int)Convert.ToDouble(textBox4.Text);
                a.End = (int)Convert.ToDouble(textBox7.Text);
               
                var b = new Level();
                b.Name = "B";
                b.UpdateDate = DateTime.Now;
                b.Begin = (int)Convert.ToDouble(textBox5.Text);
                b.End = (int)Convert.ToDouble(textBox8.Text);
                _iLevelService.Create(b);
               
                var c = new Level();
                c.Name = "C";
                c.UpdateDate = DateTime.Now;
                c.Begin = (int)Convert.ToDouble(textBox6.Text);
                c.End = (int)Convert.ToDouble(textBox9.Text);
                _iLevelService.Create(c);
               
                _iLevelService.Save();
                var successForm = new SuccessForm();
                successForm.ShowDialog();
            }
            catch (Exception ex)
            {
                var errorForm = new ErrorForm(ex.Message);
                errorForm.ShowDialog();
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            //textBox4.Text = textBox4.Text.ToString("N", CultureInfo.CreateSpecificCulture("en-US"));
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                var timesetting = _iTimeSettingService.GetAll().OrderByDescending(x => x.UpdateDate).FirstOrDefault();
                if (timesetting != null)
                {
                    timesetting.time = dateTimePicker5.Value;
                    timesetting.UpdateDate = DateTime.Now;
                    _iTimeSettingService.Update(timesetting);
                }
                else
                {
                    timesetting = new TimeSetting();
                    timesetting.UpdateDate = DateTime.Now;
                    timesetting.time = dateTimePicker5.Value;
                    _iTimeSettingService.Create(timesetting);
                }
                _iTimeSettingService.Save();
                var successForm = new SuccessForm();
                successForm.ShowDialog();
            }
            catch (Exception ex)
            {
                var errorForm = new ErrorForm(ex.Message);
                errorForm.ShowDialog();
            }
        }

        private void groupBox5_Enter(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrWhiteSpace(textBox1.Text))
            {
                try
                {
                    ReadProgram(_iWorkbook);
                    var successForm = new SuccessForm();
                    successForm.ShowDialog();
                }
                catch(Exception ex){
                    var errorForm = new ErrorForm(ex.Message);
                    errorForm.ShowDialog();
                }

            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox2.Text) && !string.IsNullOrWhiteSpace(textBox2.Text))
            {
                try
                {
                    ReadSchedule(_iWorkbook);
                    var successForm = new SuccessForm();
                    successForm.ShowDialog();
                }
                catch (Exception ex)
                {
                    var errorForm = new ErrorForm(ex.Message);
                    errorForm.ShowDialog();
                }
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox3.Text) && !string.IsNullOrWhiteSpace(textBox3.Text))
            {
                try
                {
                    ReadQuantity(_iWorkbook);
                    var successForm = new SuccessForm();
                    successForm.ShowDialog();
                }
                catch (Exception ex)
                {
                    var errorForm = new ErrorForm(ex.Message);
                    errorForm.ShowDialog();
                }
            }
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void panel1_Scroll(object sender, ScrollEventArgs e)
        {

        }

        private void button11_Click(object sender, EventArgs e)
        {
            DateTime minDate = new DateTime(2015, 01, 01).AddSeconds(-1);
            DateTime maxDate = new DateTime(2015, 01, 05); // or DateTime.Now;
            var choosenList = checkedListBox1.CheckedItems; 
            var graphForm = new Graph(choosenList, minDate, maxDate, _iSaleService, mode,_iReportService,_iTimeSettingService, _iProgramService);
            graphForm.Show();
                  
        }

        private void dateTimePicker6_ValueChanged(object sender, EventArgs e)
        {
            changeTimeGraph();
        }

        private void dateTimePicker7_ValueChanged(object sender, EventArgs e)
        {
            changeTimeGraph();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            mode = Convert.ToInt32(comboBox1.SelectedValue);
        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            mode = Convert.ToInt32(comboBox1.SelectedValue);
        }
    }    
}
