/* Title:           Computer Mean and SD
 * Date:            3-23-18
 * Author:          Terry Holmes
 * 
 * Description:     This form computes Mean and SD for Tasks */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using NewEventLogDLL;
using DateSearchDLL;
using EmployeeProjectAssignmentDLL;
using WorkTaskDLL;
using Microsoft.Win32;

namespace ComputeTaskMeanSD
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //setting up the classes
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        DateSearchClass TheDataSearchClass = new DateSearchClass();
        EmployeeProjectAssignmentClass TheEmployeeProjectAssignmentClass = new EmployeeProjectAssignmentClass();
        WorkTaskClass TheWorkTaskClass = new WorkTaskClass();

        FindLaborHoursByDateRangeDataSet TheFindLaborHoursByDateRangeDataSet = new FindLaborHoursByDateRangeDataSet();
        WorkTaskStatsDataSet TheWorkTaskStatsDataSet = new WorkTaskStatsDataSet();
        FindWorkTaskHoursDataSet TheFindWorkTaskHoursDataSet = new FindWorkTaskHoursDataSet();

        decimal gdecTotal;
        decimal gdecMean;
        int gintCounter;
        decimal gdecVariance;
        decimal gdecStandardDeviation;
        decimal gdecLimiter;
        int gintTaskUpperLimit;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            DateTime datStartDate = DateTime.Now;
            DateTime datEndDate = DateTime.Now;
            int intCounter;
            int intNumberOfRecords;
            double douHours;
            int intTaskCounter;
            bool blnItemFound;
            int intWorkTaskID;
            decimal decMean;
            decimal decVariance;
            decimal dectotalHours;
            int intItems;
            string strWorkTask;
            double douTaskHours;
            decimal decStandardDeviation;

            try
            {
                datEndDate = TheDataSearchClass.RemoveTime(datEndDate);

                datStartDate = TheDataSearchClass.SubtractingDays(datEndDate, 60);

                gintCounter = 0;
                gdecTotal = 0;
                gdecVariance = 0;
                gintTaskUpperLimit = 0;

                TheFindLaborHoursByDateRangeDataSet = TheEmployeeProjectAssignmentClass.FindLaborHoursByDateRange(datStartDate, datEndDate);

                intNumberOfRecords = TheFindLaborHoursByDateRangeDataSet.FindLaborHoursByDateRange.Rows.Count - 1;

                for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    if(TheFindLaborHoursByDateRangeDataSet.FindLaborHoursByDateRange[intCounter].TotalHours > 0)
                    {
                        intWorkTaskID = TheFindLaborHoursByDateRangeDataSet.FindLaborHoursByDateRange[intCounter].WorkTaskID;
                        blnItemFound = false;

                        if (gintTaskUpperLimit > 0)
                        {
                            for(intTaskCounter = 0; intTaskCounter < gintTaskUpperLimit; intTaskCounter++)
                            {
                                if(intWorkTaskID == TheWorkTaskStatsDataSet.worktaskstats[intTaskCounter].WorkTaskID)
                                {
                                    blnItemFound = true;
                                    TheWorkTaskStatsDataSet.worktaskstats[intTaskCounter].HoursPerTask += TheFindLaborHoursByDateRangeDataSet.FindLaborHoursByDateRange[intCounter].TotalHours;
                                    TheWorkTaskStatsDataSet.worktaskstats[intTaskCounter].ItemCounter++;
                                }
                            }
                        }

                        gdecTotal += TheFindLaborHoursByDateRangeDataSet.FindLaborHoursByDateRange[intCounter].TotalHours;
                        gintCounter++;

                        if(blnItemFound == false)
                        {
                            WorkTaskStatsDataSet.worktaskstatsRow NewTaskRow = TheWorkTaskStatsDataSet.worktaskstats.NewworktaskstatsRow();

                            NewTaskRow.HoursPerTask = TheFindLaborHoursByDateRangeDataSet.FindLaborHoursByDateRange[intCounter].TotalHours; ;
                            NewTaskRow.ItemCounter = 1;
                            NewTaskRow.Limiter = 0;
                            NewTaskRow.Mean = 0;
                            NewTaskRow.StandardDeviation = 0;
                            NewTaskRow.Variance = 0;
                            NewTaskRow.WorkTask = TheFindLaborHoursByDateRangeDataSet.FindLaborHoursByDateRange[intCounter].WorkTask;
                            NewTaskRow.WorkTaskID = TheFindLaborHoursByDateRangeDataSet.FindLaborHoursByDateRange[intCounter].WorkTaskID;

                            TheWorkTaskStatsDataSet.worktaskstats.Rows.Add(NewTaskRow);
                            gintTaskUpperLimit++;
                        }
                    }
                }

                decVariance = 0;

                for(intTaskCounter = 0; intTaskCounter < gintTaskUpperLimit; intTaskCounter++)
                {
                    intItems = TheWorkTaskStatsDataSet.worktaskstats[intTaskCounter].ItemCounter;
                    dectotalHours = TheWorkTaskStatsDataSet.worktaskstats[intTaskCounter].HoursPerTask;
                    intWorkTaskID = TheWorkTaskStatsDataSet.worktaskstats[intTaskCounter].WorkTaskID;
                    strWorkTask = TheWorkTaskStatsDataSet.worktaskstats[intTaskCounter].WorkTask;

                    decMean = dectotalHours / intItems;

                    decMean = Math.Round(decMean, 4);

                    TheWorkTaskStatsDataSet.worktaskstats[intTaskCounter].Mean = decMean;

                    TheFindWorkTaskHoursDataSet = TheWorkTaskClass.FindWorkTaskHours(strWorkTask, datStartDate, datEndDate);

                    intNumberOfRecords = TheFindWorkTaskHoursDataSet.FindWorkTaskHours.Rows.Count - 1;

                    for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                    {
                        if(TheFindWorkTaskHoursDataSet.FindWorkTaskHours[intCounter].TotalHours > 0)
                        {
                            douTaskHours = Convert.ToDouble(TheFindWorkTaskHoursDataSet.FindWorkTaskHours[intCounter].TotalHours - decMean);

                            decVariance += Convert.ToDecimal(Math.Pow(douTaskHours, 2));
                        }
                    }

                    decVariance = decVariance / intItems;

                    decVariance = Math.Round(decVariance, 4);

                    decStandardDeviation = Math.Round(Convert.ToDecimal(Math.Sqrt(Convert.ToDouble(decVariance))), 4);

                    TheWorkTaskStatsDataSet.worktaskstats[intTaskCounter].Variance = decVariance;

                    TheWorkTaskStatsDataSet.worktaskstats[intTaskCounter].StandardDeviation = decStandardDeviation;

                    TheWorkTaskStatsDataSet.worktaskstats[intTaskCounter].Limiter = decMean + (5 * decStandardDeviation);
                }

                gdecMean = gdecTotal / gintCounter;

                gdecMean = Math.Round(gdecMean, 4);

                txtTaskMean.Text = Convert.ToString(gdecMean);

                intNumberOfRecords = TheFindLaborHoursByDateRangeDataSet.FindLaborHoursByDateRange.Rows.Count - 1;

                for (intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    if (TheFindLaborHoursByDateRangeDataSet.FindLaborHoursByDateRange[intCounter].TotalHours > 0)
                    {
                        douHours = Convert.ToDouble(TheFindLaborHoursByDateRangeDataSet.FindLaborHoursByDateRange[intCounter].TotalHours - gdecMean);

                        gdecVariance += Convert.ToDecimal(Math.Pow(douHours, 2));
                    }
                }

                gdecVariance = gdecVariance / gintCounter;

                gdecStandardDeviation = Convert.ToDecimal(Math.Sqrt(Convert.ToDouble(gdecVariance)));

                gdecStandardDeviation = Math.Round(gdecStandardDeviation, 4);

                txtSD.Text = Convert.ToString(gdecStandardDeviation);

                dgrResults.ItemsSource = TheWorkTaskStatsDataSet.worktaskstats;

                gdecLimiter = gdecMean + (gdecStandardDeviation * 5);

                txtLimiter.Text = Convert.ToString(gdecLimiter);
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Compute Task Mean SD // Window Loaded " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

        private void btnExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            int intRowCounter;
            int intRowNumberOfRecords;
            int intColumnCounter;
            int intColumnNumberOfRecords;

            // Creating a Excel object. 
            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            try
            {

                worksheet = workbook.ActiveSheet;

                worksheet.Name = "OpenOrders";

                int cellRowIndex = 1;
                int cellColumnIndex = 1;
                intRowNumberOfRecords = TheWorkTaskStatsDataSet.worktaskstats.Rows.Count;
                intColumnNumberOfRecords = TheWorkTaskStatsDataSet.worktaskstats.Columns.Count;

                for (intColumnCounter = 0; intColumnCounter < intColumnNumberOfRecords; intColumnCounter++)
                {
                    worksheet.Cells[cellRowIndex, cellColumnIndex] = TheWorkTaskStatsDataSet.worktaskstats.Columns[intColumnCounter].ColumnName;

                    cellColumnIndex++;
                }

                cellRowIndex++;
                cellColumnIndex = 1;

                //Loop through each row and read value from each column. 
                for (intRowCounter = 0; intRowCounter < intRowNumberOfRecords; intRowCounter++)
                {
                    for (intColumnCounter = 0; intColumnCounter < intColumnNumberOfRecords; intColumnCounter++)
                    {
                        worksheet.Cells[cellRowIndex, cellColumnIndex] = TheWorkTaskStatsDataSet.worktaskstats.Rows[intRowCounter][intColumnCounter].ToString();

                        cellColumnIndex++;
                    }
                    cellColumnIndex = 1;
                    cellRowIndex++;
                }

                //Getting the location and file name of the excel to save from user. 
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveDialog.FilterIndex = 1;

                saveDialog.ShowDialog();

                workbook.SaveAs(saveDialog.FileName);
                MessageBox.Show("Export Successful");

            }
            catch (System.Exception ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Vehicle Reports // Vehicle Problem Report // Export to Excel " + ex.Message);

                MessageBox.Show(ex.ToString());
            }
            finally
            {
                excel.Quit();
                workbook = null;
                excel = null;
            }
        }
    }
}
