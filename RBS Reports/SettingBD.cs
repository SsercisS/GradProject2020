using System;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Win32;
using System.Threading;

namespace RBS_Reports
{
    public partial class SettingBD : Form
    {

        public static bool firstOp = true;


        Setting s = new Setting();
        public SettingBD()
        {
            InitializeComponent();
            TopMost = true;
            //if (firstOp)
            //{
            //    Thread t = new Thread(new ThreadStart(Splash));
            //    t.Start();

            //    Thread.Sleep(2000);
            //    t.Abort();
            //}
            
            
        }

        //void Splash()  Загрузочный экран, с ним иногда дропает код
        //{
        //    SplashScreen.SplashForm splash = new SplashScreen.SplashForm();
        //    splash.AppName = "RPS Reports v 1.4.2";
        //    Application.Run(splash);
        //}

        private void SettingBD_Load(object sender, EventArgs e)
        {
            RegistryKey rk = Registry.CurrentUser;
            if (rk.OpenSubKey("RBS_Setting_menu") != null)
            {
                RegistryKey config = rk.OpenSubKey("RBS_Setting_menu");
                textBox1.Text = Crypting.decryptAES(config.GetValue("ip").ToString());
                textBox2.Text = Crypting.decryptAES(config.GetValue("nameDB").ToString());
                textBox3.Text = Crypting.decryptAES(config.GetValue("login").ToString());
                textBox4.Text = Crypting.decryptAES(config.GetValue("pass").ToString());
            }
        }

       

        private void button3_Click(object sender, EventArgs e)
        {
           
            s.Memory(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text);

        }

       

        private void button2_Click(object sender, EventArgs e)
        {
            s.Connection(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text);
            try
            {
                Setting.cnt.Open();
            }
            catch(SqlException ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                MessageBox.Show("Подключено успешно");
                Setting.cnt.Close();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (firstOp)
            {
                s.Connection(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text);
                MainWindow Form1 = new MainWindow();
                Form1.Show();
                Hide();
            }
            else { Hide(); }
            
        }
    }
}
