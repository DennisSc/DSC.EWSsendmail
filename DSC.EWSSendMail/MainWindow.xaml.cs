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
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Win32;
using DataAccess;
using System.Collections.ObjectModel;
using System.Threading;
using System.ComponentModel;
using System.Configuration;

namespace DSC.EWSSendMail
{
    

    class recipient
    {
        public string smtp { get; set; }
        public string name { get; set; }
        public string var1 { get; set; }
        public string var2 { get; set; }
        public string var3 { get; set; }
        public string var4 { get; set; }
        public string var5 { get; set; }
    }

    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        

        DataTable CSV_dt;
        ObservableCollection<recipient> RecipientsList = new ObservableCollection<recipient>();
        ObservableCollection<recipient> SelectedRecipientsList = new ObservableCollection<recipient>();

        BackgroundWorker loadInfo;
        BackgroundWorker loadInfo2;

        public MainWindow()
        {
            InitializeComponent();

            App.Current.Properties["MsgBody"] = Properties.Settings.Default.MsgBody; // "Hello, ###recipientName###! This is a test email.";
            App.Current.Properties["MsgSubject"] = Properties.Settings.Default.MsgSubject;//"test email";

            loadInfo = new BackgroundWorker();
            loadInfo.DoWork += loadInfo_DoWork;
            loadInfo.RunWorkerCompleted += loadInfo_RunWorkerCompleted;
            
            loadInfo.WorkerReportsProgress = true;
            loadInfo.WorkerSupportsCancellation = true;

            loadInfo2 = new BackgroundWorker();
            loadInfo2.DoWork += loadInfo2_DoWork;
            loadInfo2.RunWorkerCompleted += loadInfo_RunWorkerCompleted;

            loadInfo2.WorkerReportsProgress = true;
            loadInfo2.WorkerSupportsCancellation = true;

        }


        public void loadInfo_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            UpdateLabelDone();
        }


        public void loadInfo2_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                this.Dispatcher.BeginInvoke((Action)(() => Cogwheel1.Visibility = Visibility.Visible));
                this.Dispatcher.BeginInvoke((Action)(() => Cogwheel1.BeginInit()));
                this.Dispatcher.BeginInvoke((Action)(() => Cogwheel1.UpdateLayout()));

                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
                service.Credentials = new WebCredentials(ConfigurationSettings.AppSettings["EWSuser"], ConfigurationSettings.AppSettings["EWSpassword"]);
                service.Url = new System.Uri("https://outlook.office365.com/ews/exchange.asmx");
                // service.AutodiscoverUrl("user1@contoso.com");


                int sendcounter = 0;

                foreach (recipient _recipient in SelectedRecipientsList)
                {
                    SendEmail(service, _recipient);
                    sendcounter++;

                    Console.WriteLine("Message " + sendcounter.ToString() + " of " + SelectedRecipientsList.Count.ToString() + " sent (selection)");
                    
                    
                    this.Dispatcher.BeginInvoke((Action)(() =>
                       UpdateLabelRunning(sendcounter.ToString(), SelectedRecipientsList.Count.ToString())
                    ));
                    //Thread.Sleep(1000);
                }
                Thread.Sleep(500);
                this.Dispatcher.BeginInvoke((Action)(() => Cogwheel1.Visibility = Visibility.Hidden));
                this.Dispatcher.BeginInvoke((Action)(() => Cogwheel1.EndInit()));
                this.Dispatcher.BeginInvoke((Action)(() => Cogwheel1.UpdateLayout()));
                this.Dispatcher.BeginInvoke((Action)(() => label1.Content = "done."));
            }
            catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                    Console.ReadLine();
                }
            
        }


        void SendEmail(ExchangeService service, recipient _recipient)
        {
            EmailMessage message = new EmailMessage(service);
            message.Sender = new EmailAddress(ConfigurationSettings.AppSettings["EWSsendAsUser"]);
            message.ToRecipients.Add(_recipient.smtp);

            string _msgSubject = App.Current.Properties["MsgSubject"].ToString();
            _msgSubject = _msgSubject.Replace("###recipientName###", _recipient.name);
            _msgSubject = _msgSubject.Replace("###var1###", _recipient.var1);
            _msgSubject = _msgSubject.Replace("###var2###", _recipient.var2);
            _msgSubject = _msgSubject.Replace("###var3###", _recipient.var3);
            _msgSubject = _msgSubject.Replace("###var4###", _recipient.var4);
            _msgSubject = _msgSubject.Replace("###var5###", _recipient.var5);

            message.Subject = _msgSubject;

            string _msgBody = App.Current.Properties["MsgBody"].ToString();
            _msgBody = _msgBody.Replace("###recipientName###", _recipient.name);
            _msgBody = _msgBody.Replace("###var1###", _recipient.var1);
            _msgBody = _msgBody.Replace("###var2###", _recipient.var2);
            _msgBody = _msgBody.Replace("###var3###", _recipient.var3);
            _msgBody = _msgBody.Replace("###var4###", _recipient.var4);
            _msgBody = _msgBody.Replace("###var5###", _recipient.var5);

            _msgBody = _msgBody.Replace(Environment.NewLine, "<br />");


            message.Body = @"<meta http-equiv='Content-Type' content='text/html; charset=us-ascii'>";
            message.Body += "<span style='font-family: Calibri, Arial, Helvetica, sans-serif; font-size: 11pt;'>";
            message.Body += _msgBody;
            message.Body += "</span>";

            
            message.Body.BodyType = BodyType.HTML;
            
            if (ConfigurationSettings.AppSettings["SaveCopyInSentItems"] == "yes")
                message.SendAndSaveCopy();
            else 
                message.Send();
        }


        public void loadInfo_DoWork(object sender, DoWorkEventArgs e)
        {
            this.Dispatcher.BeginInvoke((Action)(() => Cogwheel1.Visibility = Visibility.Visible));
            this.Dispatcher.BeginInvoke((Action)(() => Cogwheel1.BeginInit()));
            this.Dispatcher.BeginInvoke((Action)(() => Cogwheel1.UpdateLayout()));



            try
            {
                // Connect to Exchange Web Services as user1 at contoso.com.
                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
                service.Credentials = new WebCredentials(ConfigurationSettings.AppSettings["EWSuser"], ConfigurationSettings.AppSettings["EWSpassword"]);
                service.Url = new System.Uri("https://outlook.office365.com/ews/exchange.asmx");
                // service.AutodiscoverUrl("user1@contoso.com");

                int sendcounter = 0;


                foreach (recipient _recipient in RecipientsList)
                {


                    // Create the e-mail message, set its properties, and send it to user2@contoso.com, saving a copy to the Sent Items folder. 
                    SendEmail(service, _recipient);

                    sendcounter++;
                    // Write confirmation message to console window.
                    Console.WriteLine("Message " + sendcounter.ToString() + " of " + RecipientsList.Count.ToString() + " sent.");

                    this.Dispatcher.BeginInvoke((Action)(() =>
                        UpdateLabelRunning(sendcounter.ToString(), RecipientsList.Count.ToString())
                    ));



                }
               

                    
                    
                
                Thread.Sleep(500);
                this.Dispatcher.BeginInvoke((Action)(() => Cogwheel1.Visibility = Visibility.Hidden));
                this.Dispatcher.BeginInvoke((Action)(() => Cogwheel1.EndInit()));
                this.Dispatcher.BeginInvoke((Action)(() => Cogwheel1.UpdateLayout()));
                this.Dispatcher.BeginInvoke((Action)(() => label1.Content = "done."));

                //label1.Content = "done.";
                //Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                Console.ReadLine();
            }

            

        }


        private void SendMail_button_Click(object sender, RoutedEventArgs e)
        {
            label1.Content = "";
            if (checkBox1.IsChecked == false)
            {
                loadInfo.RunWorkerAsync();

                loadInfo = new BackgroundWorker();
                loadInfo.DoWork += loadInfo_DoWork;
                loadInfo.RunWorkerCompleted += loadInfo_RunWorkerCompleted;

                loadInfo.WorkerReportsProgress = true;
                loadInfo.WorkerSupportsCancellation = true;
            }
            else
            {
                List<recipient> recListTemp = new List<recipient>();

                foreach (recipient rec in listBox.SelectedItems)
                {
                    //MessageBox.Show(rec.name);
                    recListTemp.Add(rec);
                }
                SelectedRecipientsList =  new ObservableCollection<recipient>(recListTemp);
                recListTemp = null;

                loadInfo2.RunWorkerAsync();

                loadInfo2 = new BackgroundWorker();
                loadInfo2.DoWork += loadInfo2_DoWork;
                loadInfo2.RunWorkerCompleted += loadInfo_RunWorkerCompleted;

                loadInfo2.WorkerReportsProgress = true;
                loadInfo2.WorkerSupportsCancellation = true;
            }
            
        }

        private void OpenCSV_button_Click(object sender, RoutedEventArgs e)
        {

            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
            if (openFileDialog1.ShowDialog() == true)
            {
                /*System.IO.StreamReader sr = new System.IO.StreamReader(openFileDialog1.FileName);
                string _wholecsvstring = await sr.ReadToEndAsync();
                sr.Close();
                MessageBox.Show(_wholecsvstring);
                */
                CSV_dt = DataTable.New.ReadCsv(@openFileDialog1.FileName);
                RecipientsList.Clear();

                foreach (Row row in CSV_dt.Rows)
                {
                    recipient newObject = new recipient
                    {
                        smtp = row.GetValueOrEmpty("recipientSMTPAddress"),
                        name = row.GetValueOrEmpty("recipientName"),
                        var1 = row.GetValueOrEmpty("var1"),
                        var2 = row.GetValueOrEmpty("var2"),
                        var3 = row.GetValueOrEmpty("var3"),
                        var4 = row.GetValueOrEmpty("var4"),
                        var5 = row.GetValueOrEmpty("var5")
                    };
                    RecipientsList.Add(newObject);
                    
                }
                listBox.ItemsSource = RecipientsList;
                label1.Content = "";
            }
        }


        void UpdateLabelRunning(string num1, string num2)
        {
            label1.Content = ("Sending message " + num1 + " of " + num2 + "...");
        }

        void UpdateLabelDone()
        {
            label1.Content = "done :)";
        }

    
        private void SetTextBody_Click_1(object sender, RoutedEventArgs e)
        {
            Window main = new MsgTemplateDialog();
            App.Current.MainWindow = main;
            //this.Close();
            main.Show();
        }
    }
}
