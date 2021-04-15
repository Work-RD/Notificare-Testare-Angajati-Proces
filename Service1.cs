using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.ServiceProcess;
using System.Timers;

namespace Notificare_Testare_Angajati
{
    public partial class Service1 : ServiceBase
    {
        private SqlConnection myConnection = new SqlConnection(myConnectionString);

        public string rezultat = "";
        public string nr_marca = "";
        public string nr_marca_superv = "";
        public string nr_cartela = "";
        public string nume = "";
        public string prenume = "";
        public string motiv = "";
        public string email_superv = "";
        public string mesaj = "";
        public string email_server = "";
        public string email_port = "";
        public string email_ssl = "";
        public string email_from = "";
        public string email_user = "";
        public string email_pass = "";

        System.Timers.Timer timer = new System.Timers.Timer();

        public static string myConnectionString = string.Concat(new string[]
        {
            "Data Source=",
            ConfigurationManager.AppSettings["Server"],
            ",",
            ConfigurationManager.AppSettings["Port"],
            ";Network Library=DBMSSOCN;Initial Catalog=",
            ConfigurationManager.AppSettings["Username"],
            ";User ID=",
            ConfigurationManager.AppSettings["Password"],
            ";Password=",
            ConfigurationManager.AppSettings["Database"],
            ";"
        });

        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);

            timer.Interval = 1000;

            timer.Enabled = true;
        }

        protected override void OnStop()
        {
            timer.Enabled = false;
        }

        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            TimeSpan start1 = new TimeSpan(13, 30, 0);
            TimeSpan end1 = new TimeSpan(13, 30, 1);
            TimeSpan start2 = new TimeSpan(22, 00, 0);
            TimeSpan end2 = new TimeSpan(22, 00, 1);
            TimeSpan start3 = new TimeSpan(05, 0, 0);
            TimeSpan end3 = new TimeSpan(05, 0, 1);
            TimeSpan startA = new TimeSpan(15, 30, 0);
            TimeSpan endA = new TimeSpan(15, 30, 1);
            TimeSpan startR = new TimeSpan(18, 00, 0);
            TimeSpan endR = new TimeSpan(18, 00, 1);

            bool TimeBetween(DateTime datetime, TimeSpan start, TimeSpan end)
            {
                TimeSpan now = datetime.TimeOfDay;
                if (start < end)
                    return start <= now && now <= end;
                return !(end < now && now < start);
            }

            bool sch1 = TimeBetween(DateTime.Now, start1, end1);
            bool sch2 = TimeBetween(DateTime.Now, start2, end2);
            bool sch3 = TimeBetween(DateTime.Now, start3, end3);
            bool schA = TimeBetween(DateTime.Now, startA, endA);
            bool schR = TimeBetween(DateTime.Now, startR, endR);

            if (ConfigurationManager.AppSettings["SendEmail"] == "da")
            {
                if (sch1)
                {
                    send_email("1");
                }
                if (sch2)
                {
                    send_email("2");
                }
                if (sch3)
                {
                    send_email("3");
                }
                if (schA)
                {
                    send_email("A");
                }
            }

            if (ConfigurationManager.AppSettings["SendRaport"] == "da")
            {
                if (schR)
                {
                    send_raport();
                }
            }
        }

        public void send_email(string schimb)
        {
            nr_marca = "";
            nr_marca_superv = "";
            nume = "";
            prenume = "";
            motiv = "";
            email_superv = "";
            mesaj = "";

            if (schimb == "1")
            {
                DataTable dataTable = new DataTable();

                string query = "SELECT DISTINCT t3.nr_marca_superv, t3.email_superv FROM scanare t1 LEFT JOIN rezultat t2 ON t2.nr_cartela = t1.nr_cartela LEFT JOIN personal t3 ON t3.nr_cartela = t1.nr_cartela WHERE CONVERT(varchar, t1.data_scanare, 101) = CONVERT(varchar, GETDATE(), 101) AND (DATEDIFF(day, t2.data_testare, GetDate()) > 3 OR t2.data_testare is null) AND t3.tesa like 'nu' AND CAST(t1.data_scanare as time) >= CAST('05:00' as time)";
                SqlCommand cmd = new SqlCommand(query, myConnection);
                myConnection.Open();
                SqlDataAdapter dadapter = new SqlDataAdapter(cmd);
                dadapter.Fill(dataTable);
                myConnection.Close();
                dadapter.Dispose();

                foreach (DataRow row in dataTable.Rows)
                {
                    mesaj = "Pentru următoarele persoane timpul de 72 de ore de la ultima testare a expirat.\nÎncă nu s-au prezentat la testare astăzi:\n";
                    nr_marca_superv = row["nr_marca_superv"].ToString();
                    email_superv = row["email_superv"].ToString();

                    DataTable dataTable2 = new DataTable();

                    string query2 = "SELECT DISTINCT t3.nr_marca, t3.nume, t3.prenume, t2.alte_motive FROM scanare t1 LEFT JOIN rezultat t2 ON t2.nr_cartela = t1.nr_cartela LEFT JOIN personal t3 ON t3.nr_cartela = t1.nr_cartela WHERE CONVERT(varchar, t1.data_scanare, 101) = CONVERT(varchar, GETDATE(), 101) AND (DATEDIFF(day, t2.data_testare, GetDate()) > 3 OR t2.data_testare is null) AND t3.tesa like 'nu' AND CAST(t1.data_scanare as time) >= CAST('05:00' as time) AND t3.nr_marca_superv = '" + nr_marca_superv + "'";
                    SqlCommand cmd2 = new SqlCommand(query2, myConnection);
                    myConnection.Open();
                    SqlDataAdapter dadapter2 = new SqlDataAdapter(cmd2);
                    dadapter2.Fill(dataTable2);
                    myConnection.Close();
                    dadapter2.Dispose();

                    foreach (DataRow row2 in dataTable2.Rows)
                    {
                        nr_marca = row2["nr_marca"].ToString();
                        nume = row2["nume"].ToString();
                        prenume = row2["prenume"].ToString();
                        motiv = row2["alte_motive"].ToString();
                        mesaj = mesaj + nr_marca + " - " + nume + " " + prenume + " - " + motiv + "\n";
                    }
                    email(email_superv, mesaj);
                }
            }
            if (schimb == "2")
            {
                DataTable dataTable = new DataTable();

                string query = "SELECT DISTINCT t3.nr_marca_superv, t3.email_superv FROM scanare t1 LEFT JOIN rezultat t2 ON t2.nr_cartela = t1.nr_cartela LEFT JOIN personal t3 ON t3.nr_cartela = t1.nr_cartela WHERE CONVERT(varchar, t1.data_scanare, 101) = CONVERT(varchar, GETDATE(), 101) AND (DATEDIFF(day, t2.data_testare, GetDate()) > 3 OR t2.data_testare is null) AND t3.tesa like 'nu' AND (CAST(t1.data_scanare as time) >= CAST('13:30' as time) AND CAST(t1.data_scanare as time) < CAST('21:30' as time))";
                SqlCommand cmd = new SqlCommand(query, myConnection);
                myConnection.Open();
                SqlDataAdapter dadapter = new SqlDataAdapter(cmd);
                dadapter.Fill(dataTable);
                myConnection.Close();
                dadapter.Dispose();

                foreach (DataRow row in dataTable.Rows)
                {
                    mesaj = "Pentru următoarele persoane timpul de 72 de ore de la ultima testare a expirat.\nÎncă nu s-au prezentat la testare astăzi:\n";
                    nr_marca_superv = row["nr_marca_superv"].ToString();
                    email_superv = row["email_superv"].ToString();

                    DataTable dataTable2 = new DataTable();

                    string query2 = "SELECT DISTINCT t3.nr_marca, t3.nume, t3.prenume, t2.alte_motive FROM scanare t1 LEFT JOIN rezultat t2 ON t2.nr_cartela = t1.nr_cartela LEFT JOIN personal t3 ON t3.nr_cartela = t1.nr_cartela WHERE CONVERT(varchar, t1.data_scanare, 101) = CONVERT(varchar, GETDATE(), 101) AND (DATEDIFF(day, t2.data_testare, GetDate()) > 3 OR t2.data_testare is null) AND t3.tesa like 'nu' AND (CAST(t1.data_scanare as time) >= CAST('13:30' as time) AND CAST(t1.data_scanare as time) < CAST('21:30' as time)) AND t3.nr_marca_superv = '" + nr_marca_superv + "'";
                    SqlCommand cmd2 = new SqlCommand(query2, myConnection);
                    myConnection.Open();
                    SqlDataAdapter dadapter2 = new SqlDataAdapter(cmd2);
                    dadapter2.Fill(dataTable2);
                    myConnection.Close();
                    dadapter2.Dispose();

                    foreach (DataRow row2 in dataTable2.Rows)
                    {
                        nr_marca = row2["nr_marca"].ToString();
                        nume = row2["nume"].ToString();
                        prenume = row2["prenume"].ToString();
                        motiv = row2["alte_motive"].ToString();
                        mesaj = mesaj + nr_marca + " - " + nume + " " + prenume + " - " + motiv + "\n";
                    }
                    email(email_superv, mesaj);
                }
            }
            if (schimb == "3")
            {
                DataTable dataTable = new DataTable();

                string query = "SELECT DISTINCT t3.nr_marca_superv, t3.email_superv FROM scanare t1 LEFT JOIN rezultat t2 ON t2.nr_cartela = t1.nr_cartela LEFT JOIN personal t3 ON t3.nr_cartela = t1.nr_cartela WHERE (CONVERT(varchar, t1.data_scanare, 101) = CONVERT(varchar, GETDATE(), 101) OR CONVERT(varchar, t1.data_scanare, 101) =  CONVERT(varchar, DATEADD(day,DATEDIFF(day,1,GETDATE()),0), 101)) AND (DATEDIFF(day, t2.data_testare, GetDate()) > 3 OR t2.data_testare is null) AND t3.tesa like 'nu' AND (CAST(t1.data_scanare as time) >= CAST('21:30' as time) OR CAST(t1.data_scanare as time) < CAST('05:00' as time))";
                SqlCommand cmd = new SqlCommand(query, myConnection);
                myConnection.Open();
                SqlDataAdapter dadapter = new SqlDataAdapter(cmd);
                dadapter.Fill(dataTable);
                myConnection.Close();
                dadapter.Dispose();

                foreach (DataRow row in dataTable.Rows)
                {
                    mesaj = "Pentru următoarele persoane timpul de 72 de ore de la ultima testare a expirat.\nÎncă nu s-au prezentat la testare astăzi:\n";
                    nr_marca_superv = row["nr_marca_superv"].ToString();
                    email_superv = row["email_superv"].ToString();

                    DataTable dataTable2 = new DataTable();

                    string query2 = "SELECT DISTINCT t3.nr_marca, t3.nume, t3.prenume, t2.alte_motive FROM scanare t1 LEFT JOIN rezultat t2 ON t2.nr_cartela = t1.nr_cartela LEFT JOIN personal t3 ON t3.nr_cartela = t1.nr_cartela WHERE (CONVERT(varchar, t1.data_scanare, 101) = CONVERT(varchar, GETDATE(), 101) OR CONVERT(varchar, t1.data_scanare, 101) =  CONVERT(varchar, DATEADD(day,DATEDIFF(day,1,GETDATE()),0), 101)) AND (DATEDIFF(day, t2.data_testare, GetDate()) > 3 OR t2.data_testare is null) AND t3.tesa like 'nu' AND (CAST(t1.data_scanare as time) >= CAST('21:30' as time) OR CAST(t1.data_scanare as time) < CAST('05:00' as time)) AND t3.nr_marca_superv = '" + nr_marca_superv + "'";
                    SqlCommand cmd2 = new SqlCommand(query2, myConnection);
                    myConnection.Open();
                    SqlDataAdapter dadapter2 = new SqlDataAdapter(cmd2);
                    dadapter2.Fill(dataTable2);
                    myConnection.Close();
                    dadapter2.Dispose();

                    foreach (DataRow row2 in dataTable2.Rows)
                    {
                        nr_marca = row2["nr_marca"].ToString();
                        nume = row2["nume"].ToString();
                        prenume = row2["prenume"].ToString();
                        motiv = row2["alte_motive"].ToString();
                        mesaj = mesaj + nr_marca + " - " + nume + " " + prenume + " - " + motiv + "\n";
                    }
                    email(email_superv, mesaj);
                }
            }
            if (schimb == "A")
            {
                DataTable dataTable = new DataTable();

                string query = "SELECT DISTINCT t3.nr_marca_superv, t3.email_superv FROM scanare t1 LEFT JOIN rezultat t2 ON t2.nr_cartela = t1.nr_cartela LEFT JOIN personal t3 ON t3.nr_cartela = t1.nr_cartela WHERE CONVERT(varchar, t1.data_scanare, 101) = CONVERT(varchar, GETDATE(), 101) AND (DATEDIFF(day, t2.data_testare, GetDate()) > 3 OR t2.data_testare is null) AND t3.tesa like 'da'";
                SqlCommand cmd = new SqlCommand(query, myConnection);
                myConnection.Open();
                SqlDataAdapter dadapter = new SqlDataAdapter(cmd);
                dadapter.Fill(dataTable);
                myConnection.Close();
                dadapter.Dispose();

                foreach (DataRow row in dataTable.Rows)
                {
                    mesaj = "Pentru următoarele persoane timpul de 72 de ore de la ultima testare a expirat.\nÎncă nu s-au prezentat la testare astăzi:\n";
                    nr_marca_superv = row["nr_marca_superv"].ToString();
                    email_superv = row["email_superv"].ToString();

                    DataTable dataTable2 = new DataTable();

                    string query2 = "SELECT DISTINCT t3.nr_marca, t3.nume, t3.prenume, t2.alte_motive FROM scanare t1 LEFT JOIN rezultat t2 ON t2.nr_cartela = t1.nr_cartela LEFT JOIN personal t3 ON t3.nr_cartela = t1.nr_cartela WHERE CONVERT(varchar, t1.data_scanare, 101) = CONVERT(varchar, GETDATE(), 101) AND (DATEDIFF(day, t2.data_testare, GetDate()) > 3 OR t2.data_testare is null) AND t3.tesa like 'da' AND nr_marca_superv = '" + nr_marca_superv + "'";
                    SqlCommand cmd2 = new SqlCommand(query2, myConnection);
                    myConnection.Open();
                    SqlDataAdapter dadapter2 = new SqlDataAdapter(cmd2);
                    dadapter2.Fill(dataTable2);
                    myConnection.Close();
                    dadapter2.Dispose();

                    foreach (DataRow row2 in dataTable2.Rows)
                    {
                        nr_marca = row2["nr_marca"].ToString();
                        nume = row2["nume"].ToString();
                        prenume = row2["prenume"].ToString();
                        motiv = row2["alte_motive"].ToString();
                        mesaj = mesaj + nr_marca + " - " + nume + " " + prenume + " - " + motiv + "\n";
                    }
                    email(email_superv, mesaj);
                }
            }
        }

        public void send_raport()
        {
            SQLToCSV(@"C:\Notificare Testare Angajati\Raport Testare " + DateTime.Now.ToString("dd-MM-yyyy") + ".csv");
            //exportExcel();

            string persoane_testate = "";
            mesaj = "";

            myConnection.Open();
            SqlDataReader sqlDataReader = new SqlCommand("SELECT COUNT (DISTINCT nr_cartela) as Numar FROM rezultat WHERE data_testare > DATEADD(day,DATEDIFF(day,1,GETDATE()),'18:00')", myConnection).ExecuteReader();
            bool flag = sqlDataReader.Read();
            if (flag)
            {
                persoane_testate = sqlDataReader.GetValue(0).ToString();
            }
            myConnection.Close();
            mesaj = "În ultimele 24 de ore s-au testat: " + persoane_testate + " persoane";
            emailWithAttachment(ConfigurationManager.AppSettings["EmailRaport"], mesaj);
        }

        public void email(string email_to, string mesaj)
        {
            email_server = "mailhub.de.web-int.net";
            email_port = "25";
            email_ssl = "false";
            email_from = "TestareAngajati@webasto.com";
            email_user = "user";
            email_pass = "user";

            MailMessage mailMessage = new MailMessage(email_from, email_to);
            SmtpClient smtpClient = new SmtpClient();
            smtpClient.Port = int.Parse(email_port);
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtpClient.UseDefaultCredentials = false;
            smtpClient.Host = email_server;
            NetworkCredential credentials = new NetworkCredential(email_user, email_pass);
            bool flag = email_ssl.ToLower() == "true";
            if (flag)
            {
                smtpClient.EnableSsl = true;
            }
            else
            {
                smtpClient.EnableSsl = false;
            }
            bool flag2 = this.email_user != "";
            if (flag2)
            {
                smtpClient.Credentials = credentials;
            }
            mailMessage.Subject = "Testare Angajați";
            mailMessage.Body = mesaj;
            smtpClient.Send(mailMessage);
        }

        public void emailWithAttachment(string email_to, string mesaj)
        {
            email_server = "mailhub.de.web-int.net";
            email_port = "25";
            email_ssl = "false";
            email_from = "TestareAngajati@webasto.com";
            email_user = "user";
            email_pass = "user";

            MailMessage mailMessage = new MailMessage(email_from, email_to);
            SmtpClient smtpClient = new SmtpClient();
            smtpClient.Port = int.Parse(email_port);
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtpClient.UseDefaultCredentials = false;
            smtpClient.Host = email_server;
            NetworkCredential credentials = new NetworkCredential(email_user, email_pass);
            bool flag = email_ssl.ToLower() == "true";
            if (flag)
            {
                smtpClient.EnableSsl = true;
            }
            else
            {
                smtpClient.EnableSsl = false;
            }
            bool flag2 = this.email_user != "";
            if (flag2)
            {
                smtpClient.Credentials = credentials;
            }
            mailMessage.Subject = "Testare Angajați";
            mailMessage.Body = mesaj;
            mailMessage.Attachments.Add(new Attachment(@"C:\Notificare Testare Angajati\Raport Testare " + DateTime.Now.ToString("dd-MM-yyyy") + ".csv"));
            smtpClient.Send(mailMessage);
        }

        private void SQLToCSV(string Filename)
        {
            myConnection.Open();
            SqlDataReader sqlDataReader = new SqlCommand("SELECT t1.rezultat as 'Rezultat', t1.data_testare as 'Data recoltare', '' as'Data trimiterii', t1.data_testare as 'Data rezultat', '' as'Tip proba', '' as'Proba dispusa', '' as'Solicitant proba', '' as'Cod judet', '' as'Personal medical', t2.CNP, t2.nume as 'Nume', t2.prenume as 'Prenume', t2.email as 'Email', t2.nr_telefon as 'Numar telefon', '' as'Cod judet rezidenta', t2.localitate as 'Localitate rezidenta', t2.adresa as 'Adresa rezidenta', 'Laborator Laser' as'Solicitant detalii' FROM rezultat t1 LEFT JOIN personal t2 ON t2.nr_cartela = t1.nr_cartela WHERE data_testare > DATEADD(day,DATEDIFF(day,1,GETDATE()),'18:00')", myConnection).ExecuteReader();

            using (System.IO.StreamWriter fs = new System.IO.StreamWriter(Filename))
            {
                // Loop through the fields and add headers
                for (int i = 0; i < sqlDataReader.FieldCount; i++)
                {
                    string name = sqlDataReader.GetName(i);
                    if (name.Contains(","))
                        name = "\"" + name + "\"";

                    fs.Write(name + ",");
                }
                fs.WriteLine();

                // Loop through the rows and output the data
                while (sqlDataReader.Read())
                {
                    for (int i = 0; i < sqlDataReader.FieldCount; i++)
                    {
                        string value = sqlDataReader[i].ToString();
                        if (value.Contains(","))
                            value = "\"" + value + "\"";

                        fs.Write("=" + "\"" + value + "\"" + ",");
                    }
                    fs.WriteLine();
                }
            }
            myConnection.Close();
        }

        /*private void exportExcel()
        {
            DataTable dt = new DataTable();

            string query = "SELECT t1.rezultat as 'Rezultat', t1.data_testare as 'Data recoltare', '' as'Data trimiterii', t1.data_testare as 'Data rezultat', '' as'Tip proba', '' as'Proba dispusa', '' as'Solicitant proba', '' as'Cod judet', '' as'Personal medical', t2.CNP, t2.nume as 'Nume', t2.prenume as 'Prenume', t2.email as 'Email', t2.nr_telefon as 'Numar telefon', '' as'Cod judet rezidenta', t2.localitate as 'Localitate rezidenta', t2.adresa as 'Adresa rezidenta', 'Laborator Laser' as'Solicitant detalii' FROM rezultat t1 LEFT JOIN personal t2 ON t2.nr_cartela = t1.nr_cartela WHERE data_testare > DATEADD(day,DATEDIFF(day,1,GETDATE()),'18:00')";
            SqlCommand cmd = new SqlCommand(query, myConnection);
            myConnection.Open();
            SqlDataAdapter dadapter = new SqlDataAdapter(cmd);
            dadapter.Fill(dt);
            myConnection.Close();
            dadapter.Dispose();

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            try
            {
                string path = AppDomain.CurrentDomain.BaseDirectory;
                object misValue = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Excel.Workbook book = excel.Workbooks.Add(misValue);
                Microsoft.Office.Interop.Excel.Worksheet sheet = new Microsoft.Office.Interop.Excel.Worksheet();
                Microsoft.Office.Interop.Excel.Range format;

                sheet = (Microsoft.Office.Interop.Excel.Worksheet)book.Sheets["Sheet1"];
                int colIndex = 0;
                int rowIndex = 1;

                format = sheet.get_Range("A1");
                format.EntireRow.Font.Bold = true;
                format.EntireRow.Font.Size = 14;
                format = sheet.get_Range("B:D");
                format.NumberFormat = "dd/mm/yyyy";
                format = sheet.get_Range("J:J");
                format.NumberFormat = "0";
                format = sheet.get_Range("N:N");
                format.NumberFormat = "+0";

                foreach (DataColumn dc in dt.Columns)
                {
                    colIndex++;
                    sheet.Cells[1, colIndex] = dc.ColumnName;
                }
                foreach (DataRow dr in dt.Rows)
                {
                    rowIndex++;
                    colIndex = 0;

                    foreach (DataColumn dc in dt.Columns)
                    {
                        colIndex++;
                        sheet.Cells[rowIndex, colIndex] = dr[dc.ColumnName];
                    }
                }

                sheet.Columns.AutoFit();
                string filepath = @"C:\Testare Angajati\Raport Testare " + DateTime.Now.ToString("dd-MM-yyyy") + ".xlsx";

                book.SaveAs(filepath);
                book.Close();
                excel.Quit();
                releaseObject(sheet);
                releaseObject(book);
                releaseObject(excel);
                GC.Collect();
            }
            catch (Exception ex)
            {
                WriteToFile("ERROR!!!" + ex.Message + ex.StackTrace);
                excel.Quit();
            }
        }*/

        /*private void releaseObject(object o)
        {
            try
            {
                while (System.Runtime.InteropServices.Marshal.ReleaseComObject(o) > 0)
                {
                }
            }
            catch
            {
            }
            finally
            {
                o = null;
            }
        }*/

        /*private void WriteToFile(string text)
        {
            string path = "C:\\Testare Angajati\\Log.txt";
            using (StreamWriter writer = new StreamWriter(path, true))
            {
                writer.WriteLine(string.Format(text, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")));
                writer.Close();
            }
        }*/
    }
}
