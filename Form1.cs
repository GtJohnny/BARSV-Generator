using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using System.Windows.Input;
using Word=Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Security.Cryptography.X509Certificates;


namespace BARSV_Generator
{

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }



        private void Form1_Load(object sender, EventArgs e)
        {



        }

        private string WriteParagraph(Paragraph p,ref Dictionary<string, string> dict)
        {

            string str = p.Range.Text;
            int i = str.Length-1;
            while (i >0)
            {
                i = str.LastIndexOf("}}", i);
                if (i < 0) break;
                int j = str.LastIndexOf("{{", i - 2);
                if (j < 0) break;
                string key = str.Substring(j+2, i - j - 2);
                
                if (dict.ContainsKey(key))
                {
                    string v = "{{" + key + "}}";
                    string n = dict[key];

                    str=str.Replace("{{" + key + "}}", dict[key]);
                }
                i = j;
                
            }
            textBox1.Text += str+"\r\n";
            return str;
        }



        public void WriteDocument(ref string path,ref Dictionary<string,string> dict)
        {

            Word.Application app = new Word.Application();
            object miss = System.Reflection.Missing.Value;
            object readOnly = true;
            Document doc = app.Documents.Open(path,miss,readOnly);
           
            int a = 0;
            for (int i = 1; i < doc.Paragraphs.Count; i++)
            {
                 if (doc.Paragraphs[i].Range.Text == "\r\a") continue;
                 doc.Paragraphs[i].Range.Text = WriteParagraph(doc.Paragraphs[i], ref dict);
            }
            doc.SaveAs2(Environment.CurrentDirectory + "\\Test1.docx");
            doc.Close();

        }


        private void Form1_Shown(object sender, EventArgs e)
        {

            string fis = "#Ordonanta_CFL_Template.docx";
            string path = Environment.CurrentDirectory+ "\\"+fis;

            
            Word.Application app = new Word.Application();
            object miss= System.Reflection.Missing.Value;
            object readOnly = true;


            string ora_accident = "23:29";
            string criminal = "Marcel Gruiu";
            Dictionary<string,string> dict = new Dictionary<string,string>();
            dict["nr_penal"] = "69";
            dict["anul_acc"] = "2024";
            dict["luna_acc"] = "4";
            dict["ziua_acc"] = "29";
            dict["ora_accident"] = "23:29";
            dict["locul_accidentului"] = "Craiova";
            dict["criminalist1"] = "Cezar";
            dict["criminalist2"] = "Ghergu";
            dict["nume_agent1"] = "Salam";
            dict["articol_penal"] = "17";
            dict["agent1"] = "Mihnea";
            dict["anul_doc"] = "2024";
            dict["luna_doc"] = "4";
            dict["ziua_doc"] = "29";
            dict["ora_disp"] = "00:49";

            /*  multiple calls */

            WriteDocument(ref path, ref dict);

            

            


        }
    }
}
