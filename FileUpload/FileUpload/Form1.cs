using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Data.OleDb;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using ExcelDataReader;
using System.Xml;
using System.Security.Cryptography;
using System.Text.RegularExpressions;

namespace FileUpload
{

    public partial class Form1 : Form
    {
        public Form1()
        {

            InitializeComponent();
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt = tableCollection[comboBox1.SelectedItem.ToString()];
            dataGridView1.DataSource = dt;
        }
        DataTableCollection tableCollection;
        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    textBox1.Text = openFileDialog.FileName;
                    using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.ReadWrite))
                    {
                        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                            });
                            tableCollection = result.Tables;
                            comboBox1.Items.Clear();
                            foreach (DataTable table in tableCollection)
                                comboBox1.Items.Add(table.TableName);


                        }
                    }
                }
            }

        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            DataTable dt = tableCollection[comboBox1.SelectedItem.ToString()];
            dataGridView1.DataSource = dt;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //XmlWriterSettings writerSettings = new XmlWriterSettings();
            //writerSettings.Indent = true;
            //writerSettings.NewLineHandling = NewLineHandling.None;


            //using (var writer = new CustomXmlTextWriter(Path.Combine("aleph_data.xml")))
            //{
            DateTime _date = DateTime.Now;
            var _dateString = _date.ToString("dd_MM_yyyy");
            using (XmlWriter writer = XmlWriter.Create("aleph_"+_dateString+".xml"))
            {

                writer.WriteStartElement("p-file-20");
                for (int rows = 0; rows < dataGridView1.Rows.Count; rows++)
                    {
                    string akbe = dataGridView1.Rows[rows].Cells[4].Value.ToString();
                    string abc = ".";
                    string ak;
                    string lang = "LIT";
                    string aleph = "ALEPH";
                    string gender = "";
                    string zz = dataGridView1.Rows[rows].Cells[6].Value.ToString();
                    string birthday = Regex.Replace(zz, @"[^0-9]", "");
                    ak = dataGridView1.Rows[rows].Cells[4].Value.ToString();
                    bool stringtest = ak.All(char.IsDigit);
                    if (dataGridView1.Rows[rows].Cells[4].Value.ToString() == "moteris")
                    {
                        gender = "F";
                    }
                    else
                    {
                        gender = "M";
                    }
                        if (stringtest == true)
                        {
                        



                        long pirmasSk = Convert.ToInt64(ak);
                            while (pirmasSk >= 10)
                            {
                                pirmasSk = pirmasSk / 10;
                            }
                            if (ak.Length == 11 && pirmasSk == 3 || ak.Length == 11 && pirmasSk == 4 || ak.Length == 11 && pirmasSk == 5 || ak.Length == 11 && pirmasSk == 6)
                            {
                            ak = 0 + ak;
                                lang = "LIT";
                                aleph = "ALEPH";
                            }
                           


                            if (pirmasSk == 4 || pirmasSk == 6)
                            {
                                gender = "F";
                            }
                        }
                        else
                        {
                   
                            lang = "ENG";
                            aleph = "ALEPH-ENG";
                        if (ak.ToString().Contains(abc))
                        {
                            ak = birthday;
                        }
                    }
                    string fakultetas="";
                    if (dataGridView1.Rows[rows].Cells[8].Value.ToString()== "Elektronikos ir informatikos fakultetas")
                    {
                        fakultetas = "VKBEI";
                    }
                    else if (dataGridView1.Rows[rows].Cells[8].Value.ToString() == "Verslo vadybos fakultetas")
                    {
                        fakultetas = "VKBVV";
                    }
                    else if (dataGridView1.Rows[rows].Cells[8].Value.ToString() == "Pedagogikos fakultetas")
                    {                                       
                        fakultetas = "VKBPD";               
                    }                                       
                    else if (dataGridView1.Rows[rows].Cells[8].Value.ToString() == "Sveikatos priežiūros fakultetas")
                    {                                       
                        fakultetas = "VKBSP";               
                    }                                       
                    else if (dataGridView1.Rows[rows].Cells[8].Value.ToString() == "Ekonomikos fakultetas")
                    {                                       
                        fakultetas = "VKBEK";               
                    }                                       
                    else if (dataGridView1.Rows[rows].Cells[8].Value.ToString() == "Agrotechnologijų fakultetas")
                    {                                       
                        fakultetas = "VKBAT";
                    }
                    else if (dataGridView1.Rows[rows].Cells[8].Value.ToString() == "Menų ir kūrybinių technologijų fakultetas")
                    {
                        fakultetas = "VKBMT";
                    }
                    



                    writer.WriteStartElement("patron-record");
                        writer.WriteStartElement("z303");
                        writer.WriteElementString("record-action", "A");
                        writer.WriteElementString("match-id-type", "01");
                        writer.WriteElementString("match-id",  ak);
                        writer.WriteElementString("z303-id",  ak);
                        writer.WriteElementString("z303-user-type", "REG");
                        writer.WriteElementString("z303-con-lng", lang);
                        writer.WriteElementString("z303-alpha", "L");
                        writer.WriteElementString("z303-first-name", dataGridView1.Rows[rows].Cells[2].Value.ToString());
                        writer.WriteElementString("z303-last-name", dataGridView1.Rows[rows].Cells[3].Value.ToString());
                        writer.WriteElementString("z303-title", "Stud.");
                        writer.WriteElementString("z303-delinq-1", "00");
                        writer.WriteElementString("z303-delinq-n-1", "");

                        writer.WriteElementString("z303-delinq-3", "00");
                        writer.WriteElementString("z303-delinq-n-3", "+");
                        writer.WriteElementString("z303-budget", "");
                        writer.WriteElementString("z303-profile-id", aleph);
                         writer.WriteElementString("z303-ill-library", "");
                         writer.WriteElementString("z303-home-library", fakultetas);
                        writer.WriteElementString("z303-note-1", "+");
                        writer.WriteElementString("z303-ill-total-limit", "0000");
                        writer.WriteElementString("z303-ill-active-limit", "0000");
                        writer.WriteElementString("z303-birth-date", birthday.ToString());
                        writer.WriteElementString("z303-export-consent", "Y");
                        writer.WriteElementString("z303-proxy-id-type", "00");
                        writer.WriteElementString("z303-send-all-letters", "Y");
                        writer.WriteElementString("z303-plain-html", "H");
                        writer.WriteElementString("z303-want-sms", "N");
                        writer.WriteElementString("z303-title-req-limit", "0099");
                        writer.WriteElementString("z303-gender", gender);
                        writer.WriteElementString("z303-birthplace", "");
                        writer.WriteEndElement();


                        writer.WriteStartElement("z304");
                        writer.WriteElementString("record-action", "A");
                        writer.WriteElementString("z304-id",  ak);
                        writer.WriteElementString("z304-sequence", "01");
                        writer.WriteElementString("z304-address-0", dataGridView1.Rows[rows].Cells[2].Value.ToString() + " " + dataGridView1.Rows[rows].Cells[3].Value.ToString());
                        writer.WriteElementString("z304-address-1", dataGridView1.Rows[rows].Cells[7].Value.ToString());
                        writer.WriteElementString("z304-address-2", " Vilnius");
                        writer.WriteElementString("z304-address-3", "");
                        writer.WriteElementString("z304-address-4", "");
                        writer.WriteElementString("z304-zip", "");
                        writer.WriteElementString("z304-email-address", dataGridView1.Rows[rows].Cells[9].Value.ToString());
                        writer.WriteElementString("z304-telephone", "");
                        writer.WriteElementString("z304-date-from", "20210901");
                        writer.WriteElementString("z304-date-to", "20220901");
                        writer.WriteElementString("z304-address-type", "01");
                        writer.WriteElementString("z304-telephone-2", "");
                        writer.WriteElementString("z304-telephone-3", "");
                        writer.WriteElementString("z304-telephone-4", "");
                        writer.WriteEndElement();


                        writer.WriteStartElement("z304");
                        writer.WriteElementString("record-action", "A");
                        writer.WriteElementString("z304-id", ak);
                        writer.WriteElementString("z304-sequence", "02");
                        writer.WriteElementString("z304-address-0", dataGridView1.Rows[rows].Cells[2].Value.ToString() + " " + dataGridView1.Rows[rows].Cells[3].Value.ToString());
                        writer.WriteElementString("z304-address-1", dataGridView1.Rows[rows].Cells[8].Value.ToString());
                        writer.WriteElementString("z304-address-2", "Profesinis bakalauras");
                        writer.WriteElementString("z304-address-3", dataGridView1.Rows[rows].Cells[11].Value.ToString());
                        writer.WriteElementString("z304-address-4", dataGridView1.Rows[rows].Cells[10].Value.ToString());
                        writer.WriteElementString("z304-zip", "");
                        writer.WriteElementString("z304-email-address", dataGridView1.Rows[rows].Cells[9].Value.ToString());
                        writer.WriteElementString("z304-telephone", "");
                        writer.WriteElementString("z304-date-from", "20210901");
                        writer.WriteElementString("z304-date-to", "20220901");
                        writer.WriteElementString("z304-address-type", "02");
                        writer.WriteElementString("z304-telephone-2", "");
                        writer.WriteElementString("z304-telephone-3", "");
                        writer.WriteElementString("z304-telephone-4", "");
                        writer.WriteEndElement();


                        writer.WriteStartElement("z305");
                        writer.WriteElementString("record-action", "A");
                        writer.WriteElementString("z305-id", ak);
                        
                        writer.WriteElementString("z305-sub-library", "VKB50");
                        writer.WriteElementString("z305-bor-type", "ST");
                        writer.WriteElementString("z305-bor-status", "40");
                        writer.WriteElementString("z305-registration-date", "00000000");
                        writer.WriteElementString("z305-expiry-date", "00000000");
                        writer.WriteEndElement();

                    writer.WriteStartElement("z305");
                    writer.WriteElementString("record-action", "A");
                    writer.WriteElementString("z305-id", ak);
                    
                    writer.WriteElementString("z305-sub-library", "VKBCB");
                    writer.WriteElementString("z305-bor-type", "ST");
                    writer.WriteElementString("z305-bor-status", "40");
                    writer.WriteElementString("z305-registration-date", "00000000");
                    writer.WriteElementString("z305-expiry-date", "00000000");
                    writer.WriteEndElement();

                    writer.WriteStartElement("z305");
                    writer.WriteElementString("record-action", "A");
                    writer.WriteElementString("z305-id", ak);
                    
                    writer.WriteElementString("z305-sub-library", fakultetas);
                    writer.WriteElementString("z305-bor-type", "ST");
                    writer.WriteElementString("z305-bor-status", "40");
                    writer.WriteElementString("z305-registration-date", "00000000");
                    writer.WriteElementString("z305-expiry-date", "00000000");
                    writer.WriteEndElement();


                    const string valid = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";


                        string s = "";
                        using (RNGCryptoServiceProvider provider = new RNGCryptoServiceProvider())
                        {
                            while (s.Length != 8)
                            {
                                byte[] oneByte = new byte[1];
                                provider.GetBytes(oneByte);
                                char character = (char)oneByte[0];
                                if (valid.Contains(character))
                                {
                                    s += character;
                                }
                            }
                        }
                        writer.WriteStartElement("z308");
                        writer.WriteElementString("record-action", "A");
                        writer.WriteElementString("z308-key-type", "01");
                        writer.WriteElementString("z308-key-data", ak);
                        writer.WriteElementString("z308-verification", s);
                        writer.WriteElementString("z308-verification-type", "00");
                        writer.WriteElementString("z308-status", "AC");
                        writer.WriteElementString("z308-encryption", "H");
                        writer.WriteEndElement();

                        writer.WriteStartElement("z308");
                        writer.WriteElementString("record-action", "A");
                        writer.WriteElementString("z308-key-type", "02");
                        writer.WriteElementString("z308-key-data", akbe);
                        writer.WriteElementString("z308-verification", s);
                        writer.WriteElementString("z308-verification-type", "00");
                        writer.WriteElementString("z308-status", "AC");
                        writer.WriteElementString("z308-encryption", "H");
                        writer.WriteEndElement();

                    string key = dataGridView1.Rows[rows].Cells[1].Value.ToString();
                    key = key.ToUpper();
                        writer.WriteStartElement("z308");
                        writer.WriteElementString("record-action", "A");
                        writer.WriteElementString("z308-key-type", "07");
                        writer.WriteElementString("z308-key-data", key);
                        writer.WriteElementString("z308-verification", s);
                        writer.WriteElementString("z308-verification-type", "00");
                        writer.WriteElementString("z308-status", "AC");
                        writer.WriteElementString("z308-encryption", "H");
                        writer.WriteEndElement();

                        writer.WriteEndElement();
                        Console.WriteLine(rows + " done");

                    }

                    writer.WriteEndElement();





                    Console.WriteLine("done");
                }

            MessageBox.Show("Failas suformuotas!");

        }

    }
    }

