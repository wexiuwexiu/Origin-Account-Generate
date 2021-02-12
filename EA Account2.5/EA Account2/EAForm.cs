using OpenQA.Selenium;
using OpenQA.Selenium.Support.Extensions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Drawing;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace selenium
{
    public partial class EAForm : MetroFramework.Forms.MetroForm
    {
        public EAForm()
        {            
            InitializeComponent();
        }
        
        Random random = new Random();
        
        
        //通过SelectElement，并Sendkey进行选择
        public void webselect(IWebDriver driver, string xpath, string num)
        {
            var selector = driver.FindElement(By.XPath(xpath));
            SelectElement select = new SelectElement(selector);
            select.SelectByIndex(Convert.ToInt32(num));
        }     
        
        /// <summary>
        /// 日志记录
        /// </summary>
        /// <param name="slog"></param>
        public void xlog(string slog){
            richTextBox1.Text += slog + "\n\n";                   
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            Sqlite.Sqlinit();
            timer1.Start();                       
            metroLabel1.Text = "版本：" + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
            Sqlite.CreateTable("OutTable", new string[] { "ID","Email", "Name" ,"Password","CreatTime"}, new string[] { "integer", "text", "text", "text", "text" });
            Untils.Filldgv("OutTable", dgv_OL);         
            Sqlite.CreateTable("EATable", new string[] { "ID", "Email", "Name", "Password", "CreatTime" }, new string[] { "integer", "text", "text", "text", "text" });
            Untils.Filldgv("EATable", dgv_EA);
            Sqlite.CreateTable("ProtonTable", new string[] { "ID", "Email", "Name", "Password", "CreatTime" }, new string[] { "integer", "text", "text", "text", "text" });
            Untils.Filldgv("ProtonTable", dgv_pro);
        }
        [Obsolete]
        public void ProtonMail(object i)
        {
            IWebDriver driver = Untils.FireFoxDriver();
            driver.Manage().Cookies.DeleteAllCookies();
            try
            {
                string Accountmail = Untils.getRandString(12);
                string recovermail = Untils.getRandString(12)+"@163.com";
                string Accountname = Untils.getRandStringAll(8);
                string Password = Untils.ranpass();
                driver.Manage().Window.Size = new Size(1000, 1200);
                driver.Navigate().GoToUrl("https://mail.protonmail.com/create/new?language=en");
                WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, 8000));

                wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("password")));
                //密码和确认密码
                driver.FindElement(By.Id("password")).SendKeys(Password);
                driver.FindElement(By.Id("passwordc")).SendKeys(Password);
                driver.SwitchTo().Frame(1);
                wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("username")));
                //用户名
                driver.FindElement(By.Id("username")).SendKeys(Accountmail);
                driver.SwitchTo().DefaultContent();
                driver.SwitchTo().Frame(0);
                //恢复邮箱
                driver.FindElement(By.Id("notificationEmail")).SendKeys(recovermail);               
                driver.FindElement(By.Name("submitBtn")).Click();
                //wait.Until(ExpectedConditions.ElementIsVisible(By.Id("pm_loading")));//*[@id="id-signup-radio-email"]
                //Thread.Sleep(2000);
                //MessageBox.Show("1");
                //driver.FindElement(By.XPath("//*[@id=\"id-signup-radio-email\"]")).Click();
                wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("pm_loading")));
                Sqlite.InsertValues("ProtonTable", new string[] { dgv_pro.Rows.Count.ToString(), Accountmail, Accountname, Password, DateTime.Now.ToString() });
                ReFreshAll();
                driver.Quit();
            }
            catch (Exception ex)
            {
                driver.Quit();
                MessageBox.Show(ex.Message);
                xlog(ex.Message);
            }
        }
        [Obsolete]
        public void OutlookSignup(object i)
        {
            IWebDriver driver = Untils.FireFoxDriver();
            driver.Manage().Cookies.DeleteAllCookies();
            try
            {
                string Accountmail = Untils.getRandString(12) + "@outlook.com";// "@outlook.com";;//邮箱
                string Accountname = Untils.getRandStringAll(8);//账户名
                string Password = Untils.ranpass();//密码
                
                driver.Manage().Window.Size = new Size(1000, 700);
                driver.Navigate().GoToUrl("https://signup.live.com/signup?");              
                WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, 8000));
                //账号
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id=\"MemberName\"]")));
                driver.FindElement(By.XPath("//*[@id=\"MemberName\"]")).SendKeys(Accountmail);              
                driver.FindElement(By.XPath("//*[@id=\"iSignupAction\"]")).Click();
                //密码
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id=\"PasswordInput\"]")));
                driver.FindElement(By.XPath("//*[@id=\"PasswordInput\"]")).SendKeys(Password);
                driver.FindElement(By.XPath("//*[@id=\"iSignupAction\"]")).Click();
                //名字
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id=\"LastName\"]")));
                driver.FindElement(By.XPath("//*[@id=\"LastName\"]")).SendKeys(Accountname);
                driver.FindElement(By.XPath("//*[@id=\"FirstName\"]")).SendKeys(Untils.getRandStringAll(8));
                driver.FindElement(By.XPath("//*[@id=\"iSignupAction\"]")).Click();
                //地区
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id=\"Country\"]")));
                driver.FindElement(By.XPath("//*[@id=\"Country\"]")).SendKeys(random.Next(1, 80).ToString());

                var birthy = driver.FindElement(By.XPath("//*[@id=\"BirthYear\"]"));
                while (birthy.GetAttribute("value") == String.Empty)
                {
                    birthy.SendKeys("19" + random.Next(20, 80).ToString());
                    Thread.Sleep(50);
                }
                var birthm = driver.FindElement(By.XPath("//*[@id=\"BirthMonth\"]"));
                while (birthm.GetAttribute("value") == String.Empty)
                {
                    birthm.SendKeys(random.Next(1, 12).ToString());
                    Thread.Sleep(50);
                }
                var birthd = driver.FindElement(By.XPath("//*[@id=\"BirthDay\"]"));
                while (birthd.GetAttribute("value") == String.Empty)
                {
                    birthd.SendKeys(random.Next(1, 30).ToString());
                    Thread.Sleep(50);
                }
                driver.FindElement(By.XPath("//*[@id=\"iSignupAction\"]")).Click();
                
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id=\"mectrl_headerPicture\"]")));
                Sqlite.InsertValues("OutTable", new string[] { dgv_OL.Rows.Count.ToString(), Accountmail, Accountname , Password,DateTime.Now.ToString()});//插入数据到数据库中                
                ReFreshAll();
                driver.Quit();
                driver.Close();
            }
            catch (Exception ex)
            {
                driver.Quit();
                //MessageBox.Show(ex.Message);
                xlog(ex.Message );
            }
        }
        //通过点击option的序号进行选择
        public void optionclick(IWebDriver driver, string xpath, string num)
        {
            driver.FindElement(By.XPath(String.Format(xpath + "/option[{0}]", num))).Click();
        }
        [Obsolete]
        public void EAsignup(string email,string name,string Password)
        {
            IWebDriver driver = Untils.FireFoxDriver();
            driver.Manage().Cookies.DeleteAllCookies();
            try
            {               
                driver.Manage().Window.Size = new Size(800, 900);
                String url = String.Format("https://signin.ea.com/p/originX/create?execution=e1{0}s6&initref=https%3A%2F%2Faccounts.ea.com%3A443%2Fconnect%2Fauth%3Fresponse_type%3Dcode%26client_id%3DORIGIN_SPA_ID%26display%3DoriginXWeb%252Fcreate", random.Next(10000, 99999).ToString());
                driver.Navigate().GoToUrl(url);
                WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, 8000));
                //driver.ExecuteJavaScript(Properties.Resources.EASignup);
                wait.Until(ExpectedConditions.ElementExists(By.XPath("//*[@id=\"clientreg_country-selctrl\"]")));
                //地区
                optionclick(driver, "//*[@id=\"clientreg_country-selctrl\"]", random.Next(1, 100).ToString());
                //webselect(driver, "//*[@id=\"clientreg_country-selctrl\"]", getRandStringAll(2, 0));
                //日期
                optionclick(driver, "//*[@id=\"clientreg_dobyear-selctrl\"]", random.Next(20, 100).ToString());
                optionclick(driver, "//*[@id=\"clientreg_dobmonth-selctrl\"]", random.Next(1, 12).ToString());
                optionclick(driver, "//*[@id=\"clientreg_dobday-selctrl\"]", random.Next(1, 30).ToString());
                //Thread.Sleep(502222);
                var btn = driver.FindElement(By.XPath("//*[@id=\"alternativeContent\"]"));
                driver.ExecuteJavaScript("arguments[0].click();", btn);
                driver.FindElement(By.XPath("//*[@id=\"countryDobNextBtn\"]")).Click();
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id=\"email\"]")));
                //Thread.Sleep(1000);
                //邮箱
                driver.FindElement(By.XPath("//*[@id=\"email\"]")).SendKeys(email);
                //密码
                driver.FindElement(By.XPath("//*[@id=\"password\"]")).SendKeys(Password);
                //账号名
                driver.FindElement(By.XPath("//*[@id=\"originId\"]")).SendKeys(name);

                driver.FindElement(By.XPath("//*[@id=\"basicInfoNextBtn\"]")).Click();
                //wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(("//*[@id=\"home_children_button\"]"))));
                //driver.FindElement(By.XPath("//*[@id=\"home_children_button\"]")).Click();
                //图片验证结束//*[@id="victoryScreen"]
                //wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(("//*[@id=\"securityQuestion\"]"))));
                //driver.FindElement(By.XPath("//*[@id=\"basicInfoNextBtn\"]")).Click();
                //安全问题
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(("//*[@id=\"securityQuestion\"]/option[5]"))));
                optionclick(driver, "//*[@id=\"securityQuestion\"]", random.Next(1, 8).ToString());
                driver.FindElement(By.XPath("//*[@id=\"securityAnswer\"]")).SendKeys(name);
                driver.FindElement(By.XPath("//*[@id=\"submitBtn\"]")).Click();

                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(("//*[@id=\"continueDoneBtn\"]"))));
                Sqlite.InsertValues("EATable", new string[] { dgv_EA.Rows.Count.ToString (), email, name, Password,Untils.GetNetworkTime().ToString()});//插入数据到数据库中
                Sqlite.DeleteValuesOR("OutTable", new string[] { "Email", "Name" }, new string[] { email, name });
                ReFreshAll();                
                driver.Quit();
                driver.Close();
            }
            catch (Exception ex)
            {
                driver.Quit();
                //MessageBox.Show(ex.Message);
                xlog(ex.Message);               
            }
        }
        [Obsolete]
        public void EAsignuptemp(string temail)
        {
            IWebDriver driver = Untils.FireFoxDriver();
            driver.Manage().Cookies.DeleteAllCookies();
            
            string name = Untils.getRandStringAll(8);//账户名
            string Password = Untils.ranpass();//密码
            try
            {
                driver.Manage().Window.Size = new Size(800, 900);
                String url = String.Format("https://signin.ea.com/p/originX/create?execution=e1{0}s6&initref=https%3A%2F%2Faccounts.ea.com%3A443%2Fconnect%2Fauth%3Fresponse_type%3Dcode%26client_id%3DORIGIN_SPA_ID%26display%3DoriginXWeb%252Fcreate", random.Next(10000, 99999).ToString());
                driver.Navigate().GoToUrl(url);
                WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, 8000));
                //driver.ExecuteJavaScript(Properties.Resources.EASignup);
                //wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("clientreg_country-selctrl")));
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(("//*[@id=\"clientreg_dobmonth-selctrl\"]/option[12]"))));
                Thread.Sleep(500);
                //地区
                optionclick(driver, "//*[@id=\"clientreg_country-selctrl\"]", random.Next(1, 100).ToString());
                //driver.FindElement(By.XPath(String.Format(xpath + "/option[{0}]", num))).Click();

                //日期
                optionclick(driver, "//*[@id=\"clientreg_dobyear-selctrl\"]", random.Next(20, 100).ToString());
                optionclick(driver, "//*[@id=\"clientreg_dobmonth-selctrl\"]", random.Next(1, 12).ToString());
                optionclick(driver, "//*[@id=\"clientreg_dobday-selctrl\"]", random.Next(1, 30).ToString());
                Thread.Sleep(1000);
                var btn = driver.FindElement(By.Id("alternativeContent"));
                driver.ExecuteJavaScript("arguments[0].click();", btn);
                Thread.Sleep(1500);
                driver.FindElement(By.Id("countryDobNextBtn")).Click();

                wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.Id("email")));
                Thread.Sleep(500);
                //邮箱
                
                driver.FindElement(By.Id("email")).SendKeys(temail);
                //密码
                driver.FindElement(By.Id("password")).SendKeys(Password);
                //账号名
                driver.FindElement(By.Id("originId")).SendKeys(name);
                Thread.Sleep(1000);
                driver.FindElement(By.Id("basicInfoNextBtn")).Click();
              
                //安全问题
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(("//*[@id=\"securityQuestion\"]/option[8]"))));
                Thread.Sleep(500);
                optionclick(driver, "//*[@id=\"securityQuestion\"]", random.Next(1, 8).ToString());
                driver.FindElement(By.XPath("//*[@id=\"securityAnswer\"]")).SendKeys(name);
                Thread.Sleep(1000);
                driver.FindElement(By.Id("submitBtn")).Click();

                wait.Until(ExpectedConditions.ElementToBeClickable(By.Id(("continueDoneBtn"))));
                Sqlite.InsertValues("EATable", new string[] { dgv_EA.Rows.Count.ToString(), temail, name, Password, Untils.GetNetworkTime().ToString() });//插入数据到数据库中
                //Sqlite.DeleteValuesOR("OutTable", new string[] { "Email", "Name" }, new string[] { temail, name });
                ReFreshAll();
                driver.Quit();
                driver.Close();
            }
            catch (Exception ex)
            {
                driver.Quit();
                //MessageBox.Show(ex.Message);
                xlog(ex.Message);
            }
        }
        
        private void metroButton12_Click(object sender, EventArgs e)
        {
            string temail = "";
            if (metroTextBox2.Text!=string.Empty)
            {
                Control.CheckForIllegalCrossThreadCalls = false;
                string[] str = new string[metroTextBox2.Lines.Length];

                for (int i = 0; i < metroTextBox2.Lines.Length-1; i++)
                {
                    temail = metroTextBox2.Lines[i];
                    //MessageBox.Show(temail);
                    Thread td = new Thread(() =>
                    
                        EAsignuptemp(temail)
                    );
                    td.IsBackground = true;
                    td.Start();
                    Thread.Sleep(5000);
                }
            }
        }
        [Obsolete]
        private void getmail(object j)
        {
            IWebDriver driver = Untils.FireFoxDriver();
            WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, 8000));
            try
            {               
                driver.Manage().Cookies.DeleteAllCookies();
                driver.Manage().Window.Size = new Size(1300, 800);
                //driver.Navigate().GoToUrl("https://temp-mail.io/");
                for (int i = 1; i <= 10; i++)
                {
                    driver.ExecuteJavaScript("window.open('https://temp-mail.io')");
                }


                for (int i = 1; i <= 10; i++)
                {
                    driver.SwitchTo().Window(driver.WindowHandles[i]);
                    //Thread.Sleep(2000);
                    wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("email")));

                    var ma = driver.FindElement(By.Id("email"));

                    driver.FindElement(By.CssSelector("button.header-btn:nth-child(3)")).Click();
                    Thread.Sleep(1000);
                    wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.Id("email")));
                    var em = driver.FindElement(By.Id("email"));
                    //MessageBox.Show(em.GetAttribute("value"));
                    metroTextBox2.Text += em.GetAttribute("value") + "\r\n";
                }
            }
            catch{
                driver.Quit();
            }
            
        }
        private void metroButton13_Click(object sender, EventArgs e)
        {
            Control.CheckForIllegalCrossThreadCalls = false;
            ThreadPool.QueueUserWorkItem(getmail, 1);
        }
        [Obsolete]
        public void LoginOutlook(string email,string pass)
        {
            IWebDriver driver = Untils.FireFoxDriver();
            try
            {
                driver.Manage().Window.Position = new Point(1200, 50);
                driver.Manage().Window.Size = new Size(1000, 700);

                string url = "https://login.live.com/login.srf";
                driver.Navigate().GoToUrl(url);
                WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, 8000));
                wait.Until(ExpectedConditions.ElementExists(By.XPath("//*[@id=\"i0116\"]")));
                //输入邮箱
                driver.FindElement(By.XPath("//*[@id=\"i0116\"]")).SendKeys(email);
                driver.FindElement(By.XPath("//*[@id=\"idSIButton9\"]")).Click();
                //输入密码               
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id=\"i0118\"]")));
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("//*[@id=\"i0118\"]")).SendKeys(pass);
                wait.Until(ExpectedConditions.ElementExists(By.XPath("//*[@id=\"idSIButton9\"]")));
                driver.FindElement(By.XPath("//*[@id=\"idSIButton9\"]")).Click();
                Thread.Sleep(3000);
                driver.Navigate().GoToUrl("https://outlook.live.com/mail/0/inbox");                
            }
            catch (Exception ex)
            {
                driver.Quit();
                //MessageBox.Show(ex.Message);               
            }
        }
        private int rowIndex = 0;
        private int columnIndex = 0;
        DataGridView dgvw = new DataGridView();
        private void metroGrid1_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (e.RowIndex>=0)
                {
                    dgvw = sender as DataGridView;
                    //MessageBox.Show(dgvw.Name);
                    dgvw.Rows[e.RowIndex].Selected = true;
                    rowIndex = e.RowIndex;
                    columnIndex = e.ColumnIndex;
                    dgvw.CurrentCell = dgvw.Rows[e.RowIndex].Cells[1];
                    //contextMenuStrip1.Show(dgv, e.Location);
                    metroContextMenu1.Show(Cursor.Position);
                }              
            }
        }
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Sqlite.DeleteValuesOR("OutTable", new string[] { "Email", "Name" }, new string[] { dgvw.Rows[rowIndex].Cells[columnIndex].Value.ToString(),dgvw.Rows[rowIndex].Cells[columnIndex].Value.ToString() });
            dgvw.Rows.RemoveAt(rowIndex);
        }
        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(dgvw.GetClipboardContent());
        }

        public static string email = "";
        public static string name = "";
        public static string pass = "";
        
        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(rowIndex.ToString() + "  " + columnIndex.ToString()+"  "+dgvw.Rows[rowIndex].Cells[columnIndex].Value.ToString());
            email = dgvw.Rows[rowIndex].Cells[1].Value.ToString();
            name = dgvw.Rows[rowIndex].Cells[2].Value.ToString();
            pass = dgvw.Rows[rowIndex].Cells[3].Value.ToString();
            ////MessageBox.Show(Email+ Name);
            Control.CheckForIllegalCrossThreadCalls = false;
            Thread td = new Thread(() =>
            {
                EAsignup(email, name, pass);
            });
            td.IsBackground = true;
            td.Start();

            Thread td1 = new Thread(() =>
            {
                LoginOutlook(email, pass);
            });
            td1.IsBackground = true;
            td1.Start();

        }
        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            Control.CheckForIllegalCrossThreadCalls = false;
            email = dgvw.Rows[rowIndex].Cells[1].Value.ToString();
            name = dgvw.Rows[rowIndex].Cells[2].Value.ToString();
            pass = dgvw.Rows[rowIndex].Cells[3].Value.ToString();
            Thread td = new Thread(() =>
            {
                LoginOutlook(email, pass);
            });
            td.IsBackground = true;
            td.Start();
        }
        private void ReFreshAll()
        {
            Untils.Filldgv("EATable", dgv_EA);
            Untils.Filldgv("OutTable", dgv_OL);
            Untils.Filldgv("ProtonTable", dgv_OL);
        }    
        private void metroButton1_Click(object sender, EventArgs e)
        {
            ReFreshAll();
        }
        string sql = "";
        private void timer1_Tick(object sender, EventArgs e)
        {
            sql = DateTime.Now+"\r\n"+Sqlite.sqlstr;
            metroLabel2.Text = DateTime.Now.ToString();
            metroTextBox1.Text = sql;
        }
        private void metroButton2_Click(object sender, EventArgs e)
        {
            Untils.CloseByName("geckodriver");
            Untils.CloseByName("firefox");
        }
        private void metroButton4_Click(object sender, EventArgs e)
        {
            Control.CheckForIllegalCrossThreadCalls = false;
            ThreadPool.QueueUserWorkItem(OutlookSignup, 2);
            ThreadPool.QueueUserWorkItem(OutlookSignup, 3);
            ThreadPool.QueueUserWorkItem(OutlookSignup, 4);
            ThreadPool.QueueUserWorkItem(OutlookSignup, 5);
        }
        private void metroButton5_Click(object sender, EventArgs e)
        {
            Untils.savedig("OutLook账号", dgv_OL);
        }

        private void metroButton6_Click(object sender, EventArgs e)
        {
            Sqlite.DeleteAllValues("OutTable");
            Untils.Filldgv("OutTable", dgv_OL);
        }
        private void metroButton8_Click(object sender, EventArgs e)
        {
            Untils.savedig("EA账号", dgv_EA);
        }
        private void metroButton9_Click(object sender, EventArgs e)
        {
            Sqlite.DeleteAllValues("EATable");
            Untils.Filldgv("EATable", dgv_EA);
        }
        private void metroButton10_Click(object sender, EventArgs e)
        {
            ListViewItem lvi = new ListViewItem();
            lvi.Text = "1";
            lvi.SubItems.Add("bdsbsdb@outlook.com");
            lvi.SubItems.Add("Ab34555");
            lvi.SubItems.Add("aaa");
            listView1.Items.Add(lvi);
            listView1.BackColor = Color.LightSteelBlue;
        }
        
        private void EAForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            string path = Application.StartupPath + "/log.txt";
            if (!System.IO.File.Exists(path))
            {
                using (System.IO.File.Create(path)) { }
            }
            StreamWriter sw = new StreamWriter(path, true, Encoding.Default);
            sw.Write(richTextBox1.Text);
            sw.Close();

            string path2 = Application.StartupPath + "/logsql.txt";
            
            if (!System.IO.File.Exists(path2))
            {
                using(System.IO.File.Create(path2)){ }
            }
            StreamWriter sw2 = new StreamWriter(path2, true, Encoding.Default);
            sw2.Write(sql);
            sw2.Close();
            Environment.Exit(0);
        }

        private void metroButton11_Click(object sender, EventArgs e)
        {
            Control.CheckForIllegalCrossThreadCalls = false;
            ThreadPool.QueueUserWorkItem(ProtonMail, 2);
            //ThreadPool.QueueUserWorkItem(ProtonMail, 3);
            //ThreadPool.QueueUserWorkItem(ProtonMail, 4);
            //ThreadPool.QueueUserWorkItem(ProtonMail, 5);
        }

        private void metroButton14_Click(object sender, EventArgs e)
        {
            Untils.dataGridViewToCSV(dgv_EA);
        }

        private void metroButton15_Click(object sender, EventArgs e)
        {
            MessageBox.Show(Untils.getExternalIp());
        }
    }
}
