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
using System.Windows.Shapes;

namespace DSC.EWSSendMail
{
    /// <summary>
    /// Interaktionslogik für MsgBodyDialog.xaml
    /// </summary>
    public partial class MsgTemplateDialog : Window
    {
        public MsgTemplateDialog()
        {
            InitializeComponent();
            
            textBox.Text = App.Current.Properties["MsgBody"].ToString();
            textBox1.Text = App.Current.Properties["MsgSubject"].ToString();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            App.Current.Properties["MsgBody"] = textBox.Text;
            Properties.Settings.Default.MsgBody = textBox.Text;
            App.Current.Properties["MsgSubject"] = textBox1.Text;
            Properties.Settings.Default.MsgSubject = textBox1.Text;

            Properties.Settings.Default.Save();

            this.Close();
            //main.Show();
        }

        private void buttonCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
