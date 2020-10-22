using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using Microsoft.Win32;
using System.Windows.Forms;

namespace RBS_Reports
{
    class Setting
    {
        public static SqlConnection cnt = new SqlConnection();
        public static SqlCommand command = new SqlCommand("",cnt);
        public static DataTable table = new DataTable();

        public Crypting crypting = new Crypting();
        public string remarks= "";

        public void Memory(string ip, string nameBD, string login, string pass)
        {
            RegistryKey rk = Registry.CurrentUser;
            if (rk.OpenSubKey("RBS_Setting_menu") == null)//Елси такого в реестре нет, то создаем такой раздел
            {
                RegistryKey user = rk.CreateSubKey("RBS_Setting_menu", RegistryKeyPermissionCheck.ReadWriteSubTree);
                user.SetValue("ip", Crypting.encryptAES(ip));
                user.SetValue("nameDB", Crypting.encryptAES(nameBD));
                user.SetValue("login", Crypting.encryptAES(login));
                user.SetValue("pass", Crypting.encryptAES(pass));
                user.Close();
            }
            else
            {
                RegistryKey user = rk.OpenSubKey("RBS_Setting_menu", true);
                user.DeleteValue("ip");
                user.DeleteValue("nameDB");
                user.DeleteValue("login");
                user.DeleteValue("pass");

                user.SetValue("ip", Crypting.encryptAES(ip));
                user.SetValue("nameDB", Crypting.encryptAES(nameBD));
                user.SetValue("login", Crypting.encryptAES(login));
                user.SetValue("pass", Crypting.encryptAES(pass));
                user.Close();
            }
            rk.Close();
            MessageBox.Show("Сохранено успешно");
        }

        public void LoadRemarks(string dateStart, string dateEnd, DataTable table)
        {
                       
            command.CommandText = "SELECT        Tenders.Number as 'Номер закупки'"+
            ", Tenders.Subject as 'Объект закупки'" +
            ", Endorsement.Descriprion as 'Название события'" +
            ", EndorsementStep.Name as 'Название шага'" +
            ", EndorsementStepDetail.Name as 'Детали'" +
            ", EnodrsementStepDetailResolution.ResolTime as 'Дата проставления замечания'" +
            ", EnodrsementStepDetailResolution.Comment as 'Комментарий к замечанию'" +
            "FROM Endorsement " +
                         "INNER JOIN EndorsementStep ON Endorsement.id = EndorsementStep.EndorsementId " +
                        "INNER JOIN EndorsementStepDetail ON EndorsementStep.id = EndorsementStepDetail.StepId " +
                         "INNER JOIN EnodrsementStepDetailResolution ON EndorsementStepDetail.id = EnodrsementStepDetailResolution.ResolutionId " +
                         "INNER JOIN Tenders ON Endorsement.AnyId = Tenders.id " +
                         "where EnodrsementStepDetailResolution.ResolTime between '"+ dateStart + "' and '"+ dateEnd + "'";

            command.Notification = null;
            cnt.Open();
            table.Load(command.ExecuteReader());
            cnt.Close();

        }

        

        public void LoadConnect()
        {
            RegistryKey rk = Registry.CurrentUser;
            if (rk.OpenSubKey("RBS_Setting_menu") != null)
            {
                // MessageBox.Show(Encrypt.decryptAES("1MaI6AK2+fLzOHjvvebFzxzrO87PKysd/p5E2o9mECAWjkQmSAXqmVixHAsTNdzZ"));
                RegistryKey config = rk.OpenSubKey("BD_configs");

                string ip = Crypting.decryptAES(config.GetValue("ip").ToString());
                string nameDB = Crypting.decryptAES(config.GetValue("nameDB").ToString());
                string login = Crypting.decryptAES(config.GetValue("login").ToString());
                string pass = Crypting.decryptAES(config.GetValue("pass").ToString());

                Connection(ip, nameDB, login, pass);
                //MessageBox.Show(sql.ConnectionString);
            }
        }

        public void Connection(string ip, string nameDB, string login, string pass)
        {
           try
           {
                cnt.ConnectionString = "Data Source = " + ip + "; Initial Catalog = " + nameDB + "; Persist Security Info=True; " +
                "User ID = " + login + " ; Password= \"" + pass + "\"";
           }
           catch (Exception ex)
           {
                MessageBox.Show(
                             "Ошибка при формировании строки подключения.",
                              ex.ToString(),
                              MessageBoxButtons.OK,
                              MessageBoxIcon.Error);
            }
        }
    }
}
