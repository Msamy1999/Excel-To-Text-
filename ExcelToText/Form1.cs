using ExcelDataReader;
using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace ExcelToText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            CheckForIllegalCrossThreadCalls = false;
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {

           
            FolderBrowserDialog theDialog = new FolderBrowserDialog();
            if (theDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string folder = theDialog.SelectedPath;  //selected folder path


                //theDialog.Title = "Open Excel File";
                foreach (string file in Directory.EnumerateFiles(folder, "*.xlsx"))
                {
                    
                        label1.Visible = true;
                        label1.Text = "Loading";
                        button1.Enabled = false;
                        Thread th = new Thread(() => ExcelToText(file));
                        th.IsBackground = true;
                        th.Start();
                      
                    
                }
            }
            }
            catch (Exception x )
            {
                MessageBox.Show(x.ToString());
                throw;
            }
        }

        private void ExcelToText(string excelFilePath)
        {
            try
            {
                string seprator = txtSperator.Text;
                using (StreamWriter sw = new StreamWriter(Path.Combine(Path.GetDirectoryName(excelFilePath), $"{Path.GetFileNameWithoutExtension(excelFilePath)}.txt")))
                {
                    using (var stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            // 1. Use the reader methods
                            do
                            {
                                while (reader.Read())
                                {
                                    for (int i = 0; i < reader.FieldCount; i++)
                                    {
                                        sw.Write(reader.GetValue(i)?.ToString().Replace("\n", " ").Replace("\r", " "));
                                        if (i < reader.FieldCount - 1)
                                        {
                                            sw.Write(seprator);
                                        }
                                    }

                                    sw.WriteLine("");

                                    sw.Flush();
                                }
                            } while (reader.NextResult());

                            button1.Enabled = true;
                            label1.Text = "Done";
                            GC.Collect();
                        }
                    }
                }
            }
            catch (Exception x)
            {
                MessageBox.Show("Wrong file extnsion or " + x.Message);
                button1.Enabled = true;
                label1.Text = "Error";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string folder = txtPath.Text;
            foreach (string file in Directory.EnumerateFiles(folder, "*.xlsx"))
            {
                

                label1.Visible = true;
                label1.Text = "Loading";
                button1.Enabled = false;
                Thread th = new Thread(() => ExcelToText(file));
                th.IsBackground = true;
                th.Start();
                

            }
        }
    }
}