using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Windows.Forms;
using LiveCharts;
using LiveCharts.Wpf;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace slimlineCharts
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
            dateEnd.Value = dateEnd.Value.AddDays(1);
            //string sql = "SELECT forename +' ' + surname,id FROM dbo.[user] WHERE[grouping] = 5 and[current] = 1 and id<> 314";
            //using (SqlConnection connTemp = new SqlConnection(ConnectionStrings.ConnectionStringUser))
            //{
            //    connTemp.Open();
            //    using (SqlCommand cmd = new SqlCommand(sql, connTemp))
            //    {
            //        DataTable dt = new DataTable();
            //        SqlDataAdapter da = new SqlDataAdapter(cmd);
            //        da.Fill(dt);
            //        listAutoStaff.Items.Clear();
            //        foreach (DataRow row in dt.Rows)
            //        {
            //            listAutoStaff.Items.Add(row[0].ToString());
            //        }
            //    }
            //}
        }


        public void getData()
        {
            DateTime tempDate = dateStart.Value;
            List<double> workAvail = new List<double>();
            List<double> workDone = new List<double>();
            List<double> goals = new List<double>();
            List<string> dates = new List<string>();
           

            //build the sql string

            while (tempDate.ToString("yyyy-MM-dd") != dateEnd.Value.AddDays(1).ToString("yyyy-MM-dd"))
            {
                //in here call procedure that sums everything
                //usp_slimline_charts_department_workload
                using (SqlConnection conn = new SqlConnection(ConnectionStrings.ConnectionString))
                {
                    conn.Open();
                    //vvvvvvvvvvvvvvvvvv this one is for raw data
                    SqlCommand cmd = new SqlCommand("usp_slimline_charts_department_workload", conn); //
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.Add("@date", SqlDbType.DateTime).Value = tempDate;
                    cmd.Parameters.Add("@dailyGoals", SqlDbType.Int).Value = 0;
                    SqlDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        workAvail.Add(reader.GetDouble(0));
                        dates.Add(tempDate.ToString("yyyy-MM-dd"));
                    }
                    reader.Close();
                    //vvvvvvvvvvvvvvvvvvvvvvvvv this one is for daily goals
                    SqlCommand cmd2 = new SqlCommand("usp_slimline_charts_department_workload", conn); //@dailyGoals
                    cmd2.CommandType = CommandType.StoredProcedure;

                    cmd2.Parameters.Add("@date", SqlDbType.DateTime).Value = tempDate;
                    cmd2.Parameters.Add("@dailyGoals", SqlDbType.Int).Value = 1;
                    SqlDataReader reader2 = cmd2.ExecuteReader();
                    while (reader2.Read())
                    {
                        goals.Add(reader2.GetDouble(0));

                    }
                    reader2.Close();
                    //vvvvvvvvvvvvvvvvvvvvvvvvv this one is for work done
                    SqlCommand cmd3 = new SqlCommand("usp_slimline_charts_department_workload", conn); //@dailyGoals
                    cmd3.CommandType = CommandType.StoredProcedure;

                    cmd3.Parameters.Add("@date", SqlDbType.DateTime).Value = tempDate;
                    cmd3.Parameters.Add("@dailyGoals", SqlDbType.Int).Value = 2;
                    SqlDataReader reader3 = cmd3.ExecuteReader();
                    while (reader3.Read())
                    {
                        workDone.Add(reader3.GetDouble(0));

                    }
                    reader3.Close();
                    tempDate = tempDate.AddDays(1);
                }
            }
            while (cartesianChart1.Series.Count > 0) { cartesianChart1.Series.RemoveAt(0); }
            cartesianChart1.Series.Clear();
            cartesianChart1.AxisY.Clear();
            cartesianChart1.AxisX.Clear();
            var tempHoursAvail = new ChartValues<double>();
            var tempHoursDone = new ChartValues<double>();
            var tempGoals = new ChartValues<double>();

            for (int i = 0; i < workAvail.Count; i++)
            {
                //MessageBox.Show(data[value].ToString());
                tempHoursAvail.Add(workAvail[i]);
                //  MessageBox.Show(data[value].ToString()) ;
            }
            for (int i = 0; i < goals.Count; i++)
            {
                //MessageBox.Show(data[value].ToString());
                tempGoals.Add(goals[i]);
                //  MessageBox.Show(data[value].ToString()) ;
            }
            for (int i = 0; i < workDone.Count; i++)
            {
                //MessageBox.Show(data[value].ToString());
                tempHoursDone.Add(workDone[i]);
                //  MessageBox.Show(data[value].ToString()) ;
            }

            cartesianChart1.Series.Clear();
            cartesianChart1.Series.Add(new ColumnSeries
            {

                Values = tempHoursAvail,
                DataLabels = true,

                Title = "Hours Available"
            });
            cartesianChart1.Series.Add(new LineSeries
            {

                Values = tempGoals,
                DataLabels = true,

                Title = "Goal Hours"
            });
            cartesianChart1.Series.Add(new ColumnSeries
            {

                Values = tempHoursDone,
                DataLabels = true,

                Title = "Hours Done"
            });

            cartesianChart1.AxisX.Add(new Axis
            {
                Title = "Dates",
                FontSize = 10,
                Labels = dates,
                Separator = new Separator { Step = 1 }
            });

            cartesianChart1.AxisY.Add(new Axis
            {
                Title = "Hours of work",
                FontSize = 10,

            });
            cartesianChart1.LegendLocation = LegendLocation.Bottom;

        }

        private void btnSearch_Click(object sender, EventArgs e)
        {

            if (Convert.ToDateTime(dateEnd.Value.ToString("yyyy-MM-dd")) < Convert.ToDateTime(dateStart.Value.ToString("yyyy-MM-dd")))
            {
                MessageBox.Show("End date can not be less than start date!", "ERROR", MessageBoxButtons.OK);
                dateEnd.Value = dateStart.Value.AddDays(1);
                return;
            }
            getData();
        }

        private void btnEmail_Click(object sender, EventArgs e)
        {
            try
            {

                var frm = Form.ActiveForm;
                using (var bmp = new Bitmap(frm.Width, frm.Height))
                {
                    frm.DrawToBitmap(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height));
                    bmp.Save(@"C:\Temp\temp.png");
                }

                }
            catch
            {

            }





            Outlook.Application outlookApp = new Outlook.Application();
            Outlook.MailItem mailItem = outlookApp.CreateItem(Outlook.OlItemType.olMailItem);
            mailItem.Subject = "";
            mailItem.To = "";
            string imageSrc = @"C:\Temp\temp.png"; // Change path as needed

            var attachments = mailItem.Attachments;
            var attachment = attachments.Add(imageSrc);
            attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x370E001F", "image/jpeg");
            attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "myident"); // Image identifier found in the HTML code right after cid. Can be anything.
            mailItem.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/id/{00062008-0000-0000-C000-000000000046}/8514000B", true);

            // Set body format to HTML
            try
            {
                mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
                mailItem.Attachments.Add(imageSrc);
                string msgHTMLBody = "";
                mailItem.HTMLBody = msgHTMLBody;
                mailItem.Display(true);
                //mailItem.Send();
            }
            catch
            {

            }
        }
    }
}
