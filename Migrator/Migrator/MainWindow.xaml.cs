using MahApps.Metro.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
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
using Excel = Microsoft.Office.Interop.Excel;

namespace Migrator
{
    public partial class MainWindow : MetroWindow
    {
        #region Private Members

        private int _errorCount = 0;
        private DateTime _startTime;

        private int _lastInsertId = -1;
        private bool _isPayment = false;
        
        private string _destFile;
        private string _sourceFile;
        private readonly string _paymentfile = "payments_insert.sql";
        private readonly string _schedulefile = "schedules_insert.sql";

        private Microsoft.Win32.OpenFileDialog _dlg = new Microsoft.Win32.OpenFileDialog();

        #endregion

        #region Constructor

        public MainWindow()
        {
            this.DataContext = this;

            InitializeComponent();

            cbx_mt.SelectedIndex = 0;
        }

        #endregion

        #region Event Handlers

        private void cbx_mt_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _destFile = (cbx_mt.SelectedIndex == 0) ? _paymentfile : _schedulefile;
            _isPayment = (cbx_mt.SelectedIndex == 0);
            this.lastinsertid.IsEnabled = true;
        }

        private void lastinsertid_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double?> e)
        {
            if (e.NewValue.ToString() != "")
            {
                _lastInsertId = int.Parse(e.NewValue.ToString());
                btn_migrate.IsEnabled = true;
            }
            else
            {
                _lastInsertId = -1;
                btn_migrate.IsEnabled = false;
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            _dlg.DefaultExt = ".xlsx";
            _dlg.Filter = "Excel files |*.xlsx;*.xls";
            _dlg.Title = "Please select the file to process";

            Nullable<bool> result = _dlg.ShowDialog();

            if (result == true)
            {
                _sourceFile = _dlg.FileName;
                btn_migrate.IsEnabled = false;

                BackgroundWorker worker = new BackgroundWorker();
                worker.DoWork += new DoWorkEventHandler(ProcessFile);
                worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(ProcessFileComplete);

                //start work
                spinner.IsActive = true;
                livefeed.Content = "Loading files from " + _dlg.FileName;

                _startTime = DateTime.Now;
                worker.RunWorkerAsync();
            }
        }

        private void ProcessFile(object sender, DoWorkEventArgs e)
        {
            try
            {
                int sheet = 1;
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(_sourceFile);
                Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[sheet];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                ObjectSqlGenerator obj;
                if (_isPayment) { obj = new Payment(xlRange); } else { obj = new Schedule(xlRange); }
         
                using (StreamWriter file = new System.IO.StreamWriter(_destFile))
                {                   
                    file.Write(obj.IntroSql());
                    bool alldone = false;

                    for (int row = 2; row <= xlRange.Rows.Count; row++)
                    {
                        //Console.WriteLine("Row: " + row);
                        if(obj.GetValues(row, (_lastInsertId + (row-1)), out alldone))
                        {
                            if (alldone)
                            {
                                break;
                            }
                            else
                            {
                                file.Write(obj.Sql());
                            }
                        }
                        else 
                        {
                          _errorCount++;
                        }

                        livefeed.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Normal, new Action(delegate()
                        {
                            livefeed.Content = string.Format("Working on {0}: Row {1}", xlWorksheet.Name, row);
                        }));
                    }

                    file.WriteLine(obj.ExitSql());
                }

                xlWorkbook.Close(0);
                xlApp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.GetBaseException().ToString(), "An Error Occured", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ProcessFileComplete(object sender, RunWorkerCompletedEventArgs e)
        {
            spinner.IsActive = false;
            DateTime now = DateTime.Now;
            var timetaken = now.Subtract(_startTime);
            livefeed.Content = string.Format("Done with {3} errrors! Took {0} hrs {1} mins and {2} secs.", timetaken.Hours, timetaken.Minutes, timetaken.Seconds, _errorCount);
            btn_migrate.IsEnabled = true;
        }

        #endregion
    }

    #region Objects

    class ObjectSqlGenerator
    {
        #region Private members

        private string[] _headerPos;
        private Excel.Range _xlRange;
        private List<Tuple<string, string>> _headers;

        #endregion

        #region Public members

        public List<Tuple<string, string>> Headers
        {
            get { return _headers;  }
            set { _headers = value;  }
        }
        
        public string[] HeaderPos 
        { 
            get { return _headerPos; }
        }

        public Excel.Range xlRange
        {
            set { _xlRange = value;  }
            get { return _xlRange; }
        }

        #endregion

        #region Constructor
        
        public ObjectSqlGenerator(Excel.Range range)
        {
            xlRange = range;
            _headers = new List<Tuple<string, string>>();
            SetHeaderPositions();
        }

        #endregion

        #region Public Methods

        public string GetValue(int row, string colname)
        {
            string val = "";
            int pos = -1;
            
            try
            {
               pos = Array.IndexOf(HeaderPos, colname);
            }
            catch (Exception) { }

            if (pos > -1)
            {
                string type = GetHeaderType(colname);

                try
                {
                    var value = (xlRange.Cells[row, pos] as Excel.Range).Value2;

                    //Console.WriteLine(colname + " pos:" + pos.ToString() + " val:" + value);

                    if (value != null)
                    {
                        if (type == "String")
                        {
                            val = value.ToString();
                        }
                        else if (type == "Double")
                        {
                            val = ((double)value).ToString();
                        }
                        else if (type == "Amount")
                        {
                            val = value.ToString();   
                        }
                        else if (type == "Date")
                        {
                            if (Convert.ToString(value).Contains("to"))
                            {
                                val = Convert.ToString(value);
                            }
                            else
                            {
                                DateTime dt = DateTime.FromOADate(Convert.ToDouble(value));
                                val = dt.ToString();
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    string msg = string.Format("Could not process sheet 1 row {0}, col '{1}'. {2}", row, colname, e.Message);
                    LogUtil.LogInfo("MainWindow", "GetValues", msg);
                    val = "error";
                }
            }

            return val.Replace("'", "''").Trim();
        }

        public virtual bool GetValues(int row, int id, out bool alldone) { throw new NotImplementedException(); }

        public virtual string IntroSql() { throw new NotImplementedException(); }

        public virtual string Sql() { throw new NotImplementedException(); }

        public virtual string ExitSql() { throw new NotImplementedException(); }

        #endregion

        #region Private methods

        protected string GetMonth(string date)
        {
            if (date != "error" || date != "")
            {
                string[] dd = date.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
                return dd[0];
            }
            return "";
        }

        protected string GetYear(string date)
        {
            if (date != "error" || date != "")
            {
                string[] dd = date.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
                return dd[2].Substring(0,4);
            }
            return "";
        }
        
        private string GetHeaderType(string header)
        {
            string type = "";
            
            foreach(var t in _headers)
            {
                if (t.Item1.ToLower().Equals(header.ToLower())) { type = t.Item2;  }
            }
            
            return type;
        }

        private void SetHeaderPositions()
        {
            int colCount = xlRange.Columns.Count;

            _headerPos = new String[colCount+1];

            for (int col = 1; col <= colCount; col++)
            {
                string title = (string)(xlRange.Cells[1, col] as Excel.Range).Value2;
                _headerPos[col] = title;
            }
        }

        #endregion
    }

    class Payment : ObjectSqlGenerator
    {
        #region Private members

        // DB values
        private string ID { get; set; }
        private string Transaction_Reference { get; set; }
        private string Transaction_Detail { get; set; }
        private string Tier { get; set; }
        private string Contribution_Date { get; set; }
        private string Value_Date { get; set; }
        private string Subscription_Value_Date { get; set; }
        private string Transaction_Amount { get; set; }
        private string Subscription_Amount { get; set; }
        private string Company_Code { get; set; }
        private string Company_Name { get; set; }
        private string Company_ID { get; set; }
        private string Savings_Booster { get; set; }
        private string Savings_Booster_Customer_ID { get; set; }
        private string Contribution_Month_1 { get; set; }
        private string Contribution_Type_1 { get; set; }
        private string Contribution_Type_ID_1 { get; set; }
        private string Contribution_Month_2 { get; set; }
        private string Contribution_Type_2 { get; set; }
        private string Contribution_Type_ID_2 { get; set; }
        private string Contribution_Month_3 { get; set; }
        private string Contribution_Type_3 { get; set; }
        private string Contribution_Type_ID_3 { get; set; }
        private string Contribution_Month_4 { get; set; }
        private string Contribution_Type_4 { get; set; }
        private string Contribution_Type_ID_4 { get; set; }
        private string Contribution_Month_5 { get; set; }
        private string Contribution_Type_5 { get; set; }
        private string Contribution_Type_ID_5 { get; set; }
        private string Contribution_Month_6 { get; set; }
        private string Contribution_Type_6 { get; set; }
        private string Contribution_Type_ID_6 { get; set; }
        private string Contribution_Month_7 { get; set; }
        private string Contribution_Type_7 { get; set; }
        private string Contribution_Type_ID_7 { get; set; }
        private string Contribution_Month_8 { get; set; }
        private string Contribution_Type_8 { get; set; }
        private string Contribution_Type_ID_8 { get; set; }
        private string Contribution_Month_9 { get; set; }
        private string Contribution_Type_9 { get; set; }
        private string Contribution_Type_ID_9 { get; set; }
        private string Contribution_Month_10 { get; set; }
        private string Contribution_Type_10 { get; set; }
        private string Contribution_Type_ID_10 { get; set; }
        private string Contribution_Month_11 { get; set; }
        private string Contribution_Type_11 { get; set; }
        private string Contribution_Type_ID_11 { get; set; }
        private string Contribution_Month_12 { get; set; }
        private string Contribution_Type_12 { get; set; }
        private string Contribution_Type_ID_12 { get; set; }
        // End DB values

        #endregion

        #region Constructor

        public Payment(Excel.Range xlRange) : base(xlRange) 
        {
            Headers.Add(Tuple.Create("Transaction Reference", "String"));
            Headers.Add(Tuple.Create("Transaction Detail", "String"));
            Headers.Add(Tuple.Create("Tier", "String"));
            Headers.Add(Tuple.Create("Contribution Date", "Date"));
            Headers.Add(Tuple.Create("Value Date", "Date"));
            Headers.Add(Tuple.Create("Subscription Value Date", "Date"));
            Headers.Add(Tuple.Create("Transaction Amount", "Amount"));
            Headers.Add(Tuple.Create("Subscription Amount", "Amount"));
            Headers.Add(Tuple.Create("Company Code", "String"));
            Headers.Add(Tuple.Create("Company Name", "String"));
            Headers.Add(Tuple.Create("Company ID", "String"));
            Headers.Add(Tuple.Create("Savings Booster", "String"));
            Headers.Add(Tuple.Create("Savings Booster Customer ID", "String"));
            Headers.Add(Tuple.Create("Contribution Month 1", "Date"));
            Headers.Add(Tuple.Create("Contribution Type 1", "String"));
            Headers.Add(Tuple.Create("Contribution Type ID 1", "String"));
            Headers.Add(Tuple.Create("Contribution Month 2", "Date"));
            Headers.Add(Tuple.Create("Contribution Type 2", "String"));
            Headers.Add(Tuple.Create("Contribution Type ID 2", "String"));
            Headers.Add(Tuple.Create("Contribution Month 3", "Date"));
            Headers.Add(Tuple.Create("Contribution Type 3", "String"));
            Headers.Add(Tuple.Create("Contribution Type ID 3", "String"));
            Headers.Add(Tuple.Create("Contribution Month 4", "Date"));
            Headers.Add(Tuple.Create("Contribution Type 4", "String"));
            Headers.Add(Tuple.Create("Contribution Type ID 4", "String"));
            Headers.Add(Tuple.Create("Contribution Month 5", "Date"));
            Headers.Add(Tuple.Create("Contribution Type 5", "String"));
            Headers.Add(Tuple.Create("Contribution Type ID 5", "String"));
            Headers.Add(Tuple.Create("Contribution Month 6", "Date"));
            Headers.Add(Tuple.Create("Contribution Type 6", "String"));
            Headers.Add(Tuple.Create("Contribution Type ID 6", "String"));
            Headers.Add(Tuple.Create("Contribution Month 7", "Date"));
            Headers.Add(Tuple.Create("Contribution Type 7", "String"));
            Headers.Add(Tuple.Create("Contribution Type ID 7", "String"));
            Headers.Add(Tuple.Create("Contribution Month 8", "Date"));
            Headers.Add(Tuple.Create("Contribution Type 8", "String"));
            Headers.Add(Tuple.Create("Contribution Type ID 8", "String"));
            Headers.Add(Tuple.Create("Contribution Month 9", "Date"));
            Headers.Add(Tuple.Create("Contribution Type 9", "String"));
            Headers.Add(Tuple.Create("Contribution Type ID 9", "String"));
            Headers.Add(Tuple.Create("Contribution Month 10", "Date"));
            Headers.Add(Tuple.Create("Contribution Type 10", "String"));
            Headers.Add(Tuple.Create("Contribution Type ID 10", "String"));
            Headers.Add(Tuple.Create("Contribution Month 11", "Date"));
            Headers.Add(Tuple.Create("Contribution Type 11", "String"));
            Headers.Add(Tuple.Create("Contribution Type ID 11", "String"));
            Headers.Add(Tuple.Create("Contribution Month 12", "Date"));
            Headers.Add(Tuple.Create("Contribution Type 12", "String"));
            Headers.Add(Tuple.Create("Contribution Type ID 12", "String"));
        }

        #endregion

        #region Public methods 
        
        public override bool GetValues(int row, int id, out bool alldone) 
        {
            ID = id.ToString();
            Transaction_Reference = GetValue(row, "Transaction Reference");
            Transaction_Detail = GetValue(row, "Transaction Detail");
            Tier = GetValue(row, "Tier");
            Contribution_Date = GetValue(row, "Contribution Date");
            Value_Date = GetValue(row, "Value Date");
            Subscription_Value_Date = GetValue(row, "Subscription Value Date");
            Transaction_Amount = GetValue(row, "Transaction Amount");
            Subscription_Amount = GetValue(row, "Subscription Amount");
            Company_Code = GetValue(row, "Company Code");
            Company_Name = GetValue(row, "Company Name");
            Company_ID = GetValue(row, "Company ID");
            Savings_Booster = GetValue(row, "Savings Booster");
            Savings_Booster_Customer_ID = GetValue(row, "Savings Booster Customer ID");
            Contribution_Month_1 = GetValue(row, "Contribution Month 1");
            Contribution_Type_1 = GetValue(row, "Contribution Type 1");
            Contribution_Type_ID_1 = GetValue(row, "Contribution Type ID 1");
            Contribution_Month_2 = GetValue(row, "Contribution Month 2");
            Contribution_Type_2 = GetValue(row, "Contribution Type 2");
            Contribution_Type_ID_2 = GetValue(row, "Contribution Type ID 2");
            Contribution_Month_3 = GetValue(row, "Contribution Month 3");
            Contribution_Type_3 = GetValue(row, "Contribution Type 3");
            Contribution_Type_ID_3 = GetValue(row, "Contribution Type ID 3");
            Contribution_Month_4 = GetValue(row, "Contribution Month 4");
            Contribution_Type_4 = GetValue(row, "Contribution Type 4");
            Contribution_Type_ID_4 = GetValue(row, "Contribution Type ID 4");
            Contribution_Month_5 = GetValue(row, "Contribution Month 5");
            Contribution_Type_5 = GetValue(row, "Contribution Type 5");
            Contribution_Type_ID_5 = GetValue(row, "Contribution Type ID 5");
            Contribution_Month_6 = GetValue(row, "Contribution Month 6");
            Contribution_Type_6 = GetValue(row, "Contribution Type 6");
            Contribution_Type_ID_6 = GetValue(row, "Contribution Type ID 6");
            Contribution_Month_7 = GetValue(row, "Contribution Month 7");
            Contribution_Type_7 = GetValue(row, "Contribution Type 7");
            Contribution_Type_ID_7 = GetValue(row, "Contribution Type ID 7");
            Contribution_Month_8 = GetValue(row, "Contribution Month 8");
            Contribution_Type_8 = GetValue(row, "Contribution Type 8");
            Contribution_Type_ID_8 = GetValue(row, "Contribution Type ID 8");
            Contribution_Month_9 = GetValue(row, "Contribution Month 9");
            Contribution_Type_9 = GetValue(row, "Contribution Type 9");
            Contribution_Type_ID_9 = GetValue(row, "Contribution Type ID 9");
            Contribution_Month_10 = GetValue(row, "Contribution Month 10");
            Contribution_Type_10 = GetValue(row, "Contribution Type 10");
            Contribution_Type_ID_10 = GetValue(row, "Contribution Type ID 10");
            Contribution_Month_11 = GetValue(row, "Contribution Month 11");
            Contribution_Type_11 = GetValue(row, "Contribution Type 11");
            Contribution_Type_ID_11 = GetValue(row, "Contribution Type ID 11");
            Contribution_Month_12 = GetValue(row, "Contribution Month 12");
            Contribution_Type_12 = GetValue(row, "Contribution Type 12");
            Contribution_Type_ID_12 = GetValue(row, "Contribution Type ID 12");

            alldone = allEmpty();

            return (errorExists()) ? false : true;
        }

        public override string IntroSql()
        {
            return  "SET IDENTITY_INSERT [Petra_tracker].[dbo].[Jobs] ON;\n" +
                    "\nBEGIN\n" +
                    "\tIF NOT EXISTS (SELECT * FROM [Petra_tracker].[dbo].[Jobs] WHERE id=1) \n" +
                    "\tBEGIN\n" +
                    "\t\tINSERT [Petra_tracker].[dbo].[Jobs] ([id],[job_type], [job_description], [status], [owner], [modified_by], [created_at], [updated_at]) VALUES (1,N'Subscription', N'Payments Migration', N'In Progress', 1, 1, getdate(), getdate());\n" +
                    "\tEND\nEND\n"+
                    "\nSET IDENTITY_INSERT [Petra_tracker].[dbo].[Jobs] OFF;" +
                    "\n\nSET IDENTITY_INSERT [Petra_tracker].[dbo].[PPayments] ON;\n";
        }

        public override string Sql() 
        {
            string status = (Company_Code != "") ? "Identified and Approved" : "Unidentified";

            string one = string.Format("\nINSERT [Petra_tracker].[dbo].[PPayments] ([id], [job_id], [transaction_ref_no], [transaction_details], [transaction_date], [value_date], [transaction_amount], [subscription_value_date], [subscription_amount], [tier], [company_code], [company_name], [company_id], [savings_booster],[savings_booster_client_code], [status], [owner], [modified_by], [created_at], [updated_at]) VALUES ({0},1,'{1}',N'{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}',1,1,getdate(),getdate());",
                                 ID, Transaction_Reference, Transaction_Detail, Contribution_Date, Value_Date, Transaction_Amount, Subscription_Value_Date, Subscription_Amount, Tier, Company_Code, Company_Name, Company_ID, Savings_Booster, Savings_Booster_Customer_ID, status);

            one += DealDescriptionSql(ID, Contribution_Month_1, Contribution_Type_1, Contribution_Type_ID_1);
            one += DealDescriptionSql(ID, Contribution_Month_2, Contribution_Type_2, Contribution_Type_ID_2);
            one += DealDescriptionSql(ID, Contribution_Month_3, Contribution_Type_3, Contribution_Type_ID_3);
            one += DealDescriptionSql(ID, Contribution_Month_4, Contribution_Type_4, Contribution_Type_ID_4);
            one += DealDescriptionSql(ID, Contribution_Month_5, Contribution_Type_5, Contribution_Type_ID_5);
            one += DealDescriptionSql(ID, Contribution_Month_6, Contribution_Type_6, Contribution_Type_ID_6);
            one += DealDescriptionSql(ID, Contribution_Month_7, Contribution_Type_7, Contribution_Type_ID_7);
            one += DealDescriptionSql(ID, Contribution_Month_8, Contribution_Type_8, Contribution_Type_ID_8);
            one += DealDescriptionSql(ID, Contribution_Month_9, Contribution_Type_9, Contribution_Type_ID_9);
            one += DealDescriptionSql(ID, Contribution_Month_10, Contribution_Type_10, Contribution_Type_ID_10);
            one += DealDescriptionSql(ID, Contribution_Month_11, Contribution_Type_11, Contribution_Type_ID_11);
            one += DealDescriptionSql(ID, Contribution_Month_12, Contribution_Type_12, Contribution_Type_ID_12); 

            return one+"\n";
        }

        private string DealDescriptionSql(string payment_id, string date, string ct, string ctid)
        {
            if (ct != "" && date !="" & date != "error")
            {
                if (date.Contains("to"))
                {
                    //  do nothing. They should have split the dates across the contributions.
                }
                else
                {
                    string month = GetMonth(date);
                    string year = GetYear(date);
                    return string.Format("\nINSERT [Petra_tracker].[dbo].[PDealDescriptions] ([payment_id], [month], [year], [contribution_type_id], [contribution_type], [owner], [modified_by], [created_at], [updated_at]) VALUES ({0},'{1}','{2}','{3}','{4}',1,1, getdate(), getdate());\n",
                                       payment_id, month, year, ctid, ct);
                }
            }

            return "";
        }

        public override string ExitSql()
        {
            return "\nSET IDENTITY_INSERT [Petra_tracker].[dbo].[PPayments] OFF;";
        }

        #endregion

        #region Private methods

        private bool errorExists()
        {
            if (Transaction_Reference == "error" ||
                Transaction_Detail == "error" ||
                Tier == "error" ||
                Contribution_Date == "error" ||
                Value_Date == "error" ||
                Subscription_Value_Date == "error" ||
                Transaction_Amount == "error" ||
                Subscription_Amount == "error" ||
                Company_Code == "error" ||
                Company_Name == "error" ||
                Company_ID == "error" ||
                Savings_Booster == "error" ||
                Savings_Booster_Customer_ID == "error" ||
                Contribution_Month_1 == "error" ||
                Contribution_Type_1 == "error" ||
                Contribution_Type_ID_1 == "error" ||
                Contribution_Month_2 == "error" ||
                Contribution_Type_2 == "error" ||
                Contribution_Type_ID_2 == "error" ||
                Contribution_Month_3 == "error" ||
                Contribution_Type_3 == "error" ||
                Contribution_Type_ID_3 == "error" ||
                Contribution_Month_4 == "error" ||
                Contribution_Type_4 == "error" ||
                Contribution_Type_ID_4 == "error" ||
                Contribution_Month_5 == "error" ||
                Contribution_Type_5 == "error" ||
                Contribution_Type_ID_5 == "error" ||
                Contribution_Month_6 == "error" ||
                Contribution_Type_6 == "error" ||
                Contribution_Type_ID_6 == "error" ||
                Contribution_Month_7 == "error" ||
                Contribution_Type_7 == "error" ||
                Contribution_Type_ID_7 == "error" ||
                Contribution_Month_8 == "error" ||
                Contribution_Type_8 == "error" ||
                Contribution_Type_ID_8 == "error" ||
                Contribution_Month_9 == "error" ||
                Contribution_Type_9 == "error" ||
                Contribution_Type_ID_9 == "error" ||
                Contribution_Month_10 == "error" ||
                Contribution_Type_10 == "error" ||
                Contribution_Type_ID_10 == "error" ||
                Contribution_Month_11 == "error" ||
                Contribution_Type_11 == "error" ||
                Contribution_Type_ID_11 == "error" ||
                Contribution_Month_12 == "error" ||
                Contribution_Type_12 == "error" ||
                Contribution_Type_ID_12 == "error") { return true; }
            return false;
        }

        private bool allEmpty()
        {
            if (Transaction_Reference == "" &&
                Transaction_Detail == "" &&
                Tier == "" &&
                Contribution_Date == "" &&
                Value_Date == "" &&
                Subscription_Value_Date == "" &&
                Transaction_Amount == "" &&
                Subscription_Amount == "" &&
                Company_Code == "" &&
                Company_Name == "" &&
                Company_ID == "" &&
                Savings_Booster == "" &&
                Savings_Booster_Customer_ID == "" &&
                Contribution_Month_1 == "" &&
                Contribution_Type_1 == "" &&
                Contribution_Type_ID_1 == "" &&
                Contribution_Month_2 == "" &&
                Contribution_Type_2 == "" &&
                Contribution_Type_ID_2 == "" &&
                Contribution_Month_3 == "" &&
                Contribution_Type_3 == "" &&
                Contribution_Type_ID_3 == "" &&
                Contribution_Month_4 == "" &&
                Contribution_Type_4 == "" &&
                Contribution_Type_ID_4 == "" &&
                Contribution_Month_5 == "" &&
                Contribution_Type_5 == "" &&
                Contribution_Type_ID_5 == "" &&
                Contribution_Month_6 == "" &&
                Contribution_Type_6 == "" &&
                Contribution_Type_ID_6 == "" &&
                Contribution_Month_7 == "" &&
                Contribution_Type_7 == "" &&
                Contribution_Type_ID_7 == "" &&
                Contribution_Month_8 == "" &&
                Contribution_Type_8 == "" &&
                Contribution_Type_ID_8 == "" &&
                Contribution_Month_9 == "" &&
                Contribution_Type_9 == "" &&
                Contribution_Type_ID_9 == "" &&
                Contribution_Month_10 == "" &&
                Contribution_Type_10 == "" &&
                Contribution_Type_ID_10 == "" &&
                Contribution_Month_11 == "" &&
                Contribution_Type_11 == "" &&
                Contribution_Type_ID_11 == "" &&
                Contribution_Month_12 == "" &&
                Contribution_Type_12 == "" &&
                Contribution_Type_ID_12 == "") { return true; }
            return false;
        }

        #endregion
    }

    class Schedule : ObjectSqlGenerator
    {
        #region Private members
        
        // DB Values
        private string ID { get; set; }
        private string Company_Code { get; set; }
        private string Company_Name { get; set; }
        private string Company_ID { get; set; }
        private string Contribution_Type { get; set; }
        private string Contribution_Type_ID { get; set; }
        private string Contribution_Month { get; set; }
        private string Tier { get; set; }
        // End DB Values

        #endregion

        #region Constructor

        public Schedule(Excel.Range xlRange) : base(xlRange) 
        {
            Headers.Add(Tuple.Create("Company Code", "String"));
            Headers.Add(Tuple.Create("Company Name", "String"));
            Headers.Add(Tuple.Create("Company ID", "Double"));
            Headers.Add(Tuple.Create("Contribution Type", "String"));
            Headers.Add(Tuple.Create("Contribution Type ID", "Double"));
            Headers.Add(Tuple.Create("Contribution Month", "Date"));
            Headers.Add(Tuple.Create("Tier", "String"));
        }

        #endregion

        #region Public methods

        public override bool GetValues(int row, int id, out bool alldone)
        {
            ID = id.ToString();
            Company_Code = GetValue(row, "Company Code");
            Company_Name = GetValue(row, "Company Name");
            Company_ID = GetValue(row, "Company ID");
            Contribution_Type = GetValue(row, "Contribution Type");
            Contribution_Type_ID = GetValue(row, "Contribution Type ID");
            Contribution_Month = GetValue(row, "Contribution Month");
            Tier = GetValue(row, "Tier");

            alldone = allEmpty();

            return (errorExists()) ? false : true;
        }

        public override string IntroSql()
        {
            return "SET IDENTITY_INSERT [Petra_tracker].[dbo].[Schedules] ON;\n";
        }

        public override string Sql()
        {
            if (Contribution_Month.Contains("to"))
            {
                //  do nothing. They should been in the right format.
            }
            else
            {
                string month = GetMonth(Contribution_Month);
                string year = GetYear(Contribution_Month);
                return string.Format("\nINSERT INTO [Petra_tracker].[dbo].[Schedules] ([id] ,[company_id] ,[company] ,[tier] ,[amount] ,[contributiontype] ,[contributiontypeid] ,[month] ,[year] ,[validated] ,[validation_status] ,[file_downloaded] ,[file_uploaded] ,[receipt_sent] ,[workflow_status] ,[workflow_summary] ,[parent_id] ,[modified_by] ,[created_at] ,[updated_at] ,[ptas_fund_deal_id]) VALUES ({0},N'{1}','{2}','{3}',0.00,'{4}','{5}','{6}','{7}',0,'Not Validated',0,0,0,'Not Validated','Processing of this schedule has not begun',0,1,(getdate()),(getdate()),0);",
                                      ID, Company_ID, Company_Name, Tier, Contribution_Type, Contribution_Type_ID, month, year);
            }
            return "";
        }

        public override string ExitSql()
        {
            return "\n\nSET IDENTITY_INSERT [Petra_tracker].[dbo].[Schedules] OFF;";
        }

        #endregion

        #region Private methods

        private bool errorExists()
        {
            if (Company_Code == "error" ||
                Company_Name == "error" ||
                Company_ID == "error" ||
                Contribution_Type == "error" ||
                Contribution_Type_ID == "error" ||
                Contribution_Month == "error" ||
                Tier == "error") { return true; }
            return false;
        }

        private bool allEmpty()
        {
            if (Company_Code == "" &&
                Company_Name == "" &&
                Company_ID == "" &&
                Contribution_Type == "" &&
                Contribution_Type_ID == "" &&
                Contribution_Month == "" &&
                Tier == "") { return true; }
            return false;
        }

        #endregion
    }

    #endregion
}
