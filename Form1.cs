using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Net;
using System.IO;

namespace Laba_2
{
    public partial class Form1 : Form
    {
        List<Danger> listSt = new List<Danger>();
       
        

       public Form1()
        {
            InitializeComponent();
            
            listSt = FilledUp(listSt);
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();

            foreach (var item in listSt)
            {
                listBox1.Items.Add(" identifier: " + item.identifier + " | " + " name: " + item.name);
                listBox1.Items.Add(" description: " + "  " + item.description);
                listBox1.Items.Add(" source: " + "  " + item.source);
                listBox1.Items.Add(" impacts: " + "  " + item.impacts);
                listBox1.Items.Add(" privacyPolicy: " + "  " + item.privacyPolicy + " integrity: " + "  " + item.integrity + " availability: " + "  " + item.availability);


            }
    
        }

        private void button2_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();

            foreach (var item in listSt)
            {
                listBox1.Items.Add(" identifier: " + item.identifier + " | " + " name: " + item.name);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
           int a = textBox1_TextChanged(1,e);

            listBox1.Items.Add(listSt[a].identifier);
            listBox1.Items.Add(listSt[a].name);
            listBox1.Items.Add(listSt[a].description);
            listBox1.Items.Add(listSt[a].source);
            listBox1.Items.Add(listSt[a].privacyPolicy);
            listBox1.Items.Add(listSt[a].identifier + listSt[a].integrity + listSt[a].availability);

        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                string url = "https://bdu.fstec.ru/documents/files/thrlist.xlsx";
                string save_path = "C:\\io\\thrlist2.xlsx";
                //string name = "hrlist.xlsx";
                WebClient wc = new WebClient();
                wc.DownloadFile(url, save_path);
            }
            catch (System.Net.WebException)
            {
                MessageBox.Show("не тыкай много раз !");

            }
          

            try
            {
                FileStream fileStream = new FileStream("C:\\io\\thrlist2.xlsx", FileMode.Open);
                FileStream fileStream2 = new FileStream("C:\\io\\thrlist.xlsx", FileMode.Open);
                StreamReader reader = new StreamReader(fileStream);
                StreamReader reader2 = new StreamReader(fileStream2);
                string str = reader.ReadToEnd();
                string str2 = reader2.ReadToEnd();
                fileStream.Close();
                fileStream.Close();

                if (str2.Equals(str))
                {
                    listBox1.Items.Clear();
                    listBox1.Items.Add("Обнавление файла не требуется !!!");
                }
                else
                {
                    listBox1.Items.Clear();
                    listBox1.Items.Add("Требуется обнавеление!!!");
                }
            }
            catch (System.IO.IOException)
            {
                MessageBox.Show("не тыкай много раз !!!");           
            }   
        
            
        }

        static List<Danger> FilledUp(List<Danger> list)
        {


            try
            {
                string path = "C:\\io\\thrlist.xlsx";

                Application excel = new Application();
                Workbook wb = excel.Workbooks.Open(path);
                Worksheet excelSheet = wb.ActiveSheet;
           

            for (int i = 3; i < 215; i++)
            {
                Danger danger = new Danger();

                list.Add(danger);

                danger.identifier = excelSheet.Cells[i, 1].Value.ToString();
                danger.name = excelSheet.Cells[i, 2].Value.ToString();
                danger.description = excelSheet.Cells[i, 3].Value.ToString();
                danger.source = excelSheet.Cells[i, 4].Value.ToString();
                danger.impacts = excelSheet.Cells[i, 5].Value.ToString();
                danger.privacyPolicy = excelSheet.Cells[i, 6].Value.ToString();
                danger.integrity = excelSheet.Cells[i, 7].Value.ToString();
                danger.availability = excelSheet.Cells[i, 8].Value.ToString();

            }

            wb.Close();

            return list;

            }
            catch (System.Runtime.InteropServices.COMException a)
            {
              

                MessageBox.Show("Ээээээээээээээээээй!!!\n  Нету файла, нужно его скачать на диск С:\\io :P");
                throw a;


            }
        }

        private int textBox1_TextChanged(object sender, EventArgs e)
        {
            int a = Convert.ToInt32(Console.ReadLine());

            return a;
        }


    }
}
