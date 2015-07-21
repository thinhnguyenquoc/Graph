using AZReport.Model;
using AZReport.Services;
using AZReport.Services.IServices;
using AZReport.ViewModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace AZReport
{
    public partial class Graph : Form
    {
        private System.Windows.Forms.CheckedListBox.CheckedItemCollection _data;
        private ISaleService _iSaleService;
        IReportService _iReportService;
        ITimeSettingService _iTimeSettingService;
        IProgramService _iProgramService;
        DateTime _minDate;
        DateTime _maxDate;
        private List<Color> colorList = new List<Color>();

        public Graph(System.Windows.Forms.CheckedListBox.CheckedItemCollection data, DateTime minDate, DateTime maxDate, ISaleService iSaleService, int mode, IReportService iReportService, ITimeSettingService iTimeSettingService, IProgramService iProgramService)
        {
            InitializeComponent();
            fillColor(colorList);
            _data = data;
            _minDate = minDate;
            _maxDate = maxDate;
            _iSaleService = iSaleService;
            _iReportService = iReportService;
            _iTimeSettingService = iTimeSettingService;
            _iProgramService = iProgramService;

            this.chart1.ChartAreas[0].AxisX.MajorGrid.LineDashStyle = System.Windows.Forms.DataVisualization.Charting.ChartDashStyle.Dash;
            chart1.ChartAreas.Add("area");
            chart1.ChartAreas[0].AxisX.LabelStyle.Format = "dd-MM";
            chart1.ChartAreas["area"].AxisX.Interval = 1;
            chart1.ChartAreas["area"].AxisX.IntervalType = DateTimeIntervalType.Days;
            chart1.ChartAreas["area"].AxisX.IntervalOffset = 1;
            
            chart1.ChartAreas["area"].AxisX.Minimum = minDate.ToOADate();
            chart1.ChartAreas["area"].AxisX.Maximum = maxDate.ToOADate();
            if (mode == 1)
            {
                drawQuantity();
            }
            else
            {
                drawEfficiency();
            }
        }

        private void drawEfficiency()
        {
            int index = 0;
            foreach (string j in _data)
            {
                var code = j.Split('-')[0].ToString();
                chart1.Series.Add(code);
                chart1.Series[index].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                if (index > colorList.Count() - 1)
                    chart1.Series[index].Color = colorList[0];
                else
                    chart1.Series[index].Color = colorList[index];
                chart1.Series[index].BorderWidth = 2;
                chart1.Series[index].XValueType = ChartValueType.DateTime;
                List<Sale> saleResult = _iSaleService.GetQuantity(_minDate, _maxDate, code);
                var timesetting = _iTimeSettingService.GetAll().ToList().FirstOrDefault();
                var time1 = timesetting.time;
                var freq = _iReportService.GetItemFreq(_minDate, _maxDate, time1, code);
                var item = _iProgramService.FindItem(code);
                foreach (var i in saleResult)
                {
                    var amount = Convert.ToInt32(item.Price) * i.Quantity;
                    var totalTime = freq[index].Freq * (Convert.ToDateTime(item.Duration).Minute + Convert.ToDateTime(item.Duration).Second / 60.0); 
                    chart1.Series[index].Points.AddXY(i.Date, amount/totalTime);
                }
                index++;
            }
        }

        private void drawQuantity()
        {
            int index = 0;
            foreach (string j in _data)
            {
                var code = j.Split('-')[0].ToString();
                chart1.Series.Add(code);
                chart1.Series[index].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
                if (index > colorList.Count() - 1)
                    chart1.Series[index].Color = colorList[0];
                else
                    chart1.Series[index].Color = colorList[index];
                chart1.Series[index].BorderWidth = 2;
                chart1.Series[index].XValueType = ChartValueType.DateTime;
                var saleResult = _iSaleService.GetQuantity(_minDate, _maxDate, code);
                foreach (var i in saleResult)
                {                    
                    chart1.Series[index].Points.AddXY(i.Date, i.Quantity);
                }
                index++;
            }
        }

        private void fillColor(List<Color> colorList)
        {
            colorList.Add(Color.Black);
            colorList.Add(Color.Brown);
            colorList.Add(Color.Blue);
            colorList.Add(Color.BlueViolet);
            colorList.Add(Color.Green);
            colorList.Add(Color.GreenYellow);           
            colorList.Add(Color.Yellow);
            colorList.Add(Color.Orange);
            colorList.Add(Color.OrangeRed);
            colorList.Add(Color.Red);
            colorList.Add(Color.Purple);
            colorList.Add(Color.Pink);
        }
    }
}
