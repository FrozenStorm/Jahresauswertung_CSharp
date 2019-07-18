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
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace Jahresauswertung
{// mehrere gleiche ränge, erstellungsdatum
    public partial class Form1 : Form
    {

        OpenFileDialog ofd1 = new OpenFileDialog();
        String[] files;
        List<Rennfahrer> jahresrangliste = new List<Rennfahrer>();
        List<Lauf> läufe = new List<Lauf>();
        string klasse;
        string jahr;
        string path;

        public Form1()
        {
            InitializeComponent();

        }

        public int rangZuPunkte(int rang)
        {
            int[] punktetabelle = { 100, 90, 82, 74, 66, 60, 54, 50, 46, 42, 38, 36, 34, 32, 30, 28, 26, 24, 22, 20, 18, 16, 14, 12, 10, 8, 6, 4, 2, 1 };
            return punktetabelle[rang - 1];
        }

        public int läufeZuGewerteteLäufe(int anzahlLäufe)
        {
            int[] zuWertetndeLäufe = { 1, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7 };
            return zuWertetndeLäufe[anzahlLäufe - 1];
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ofd1.InitialDirectory = "C:\\Users\\Daniel Zwygart\\Dropbox\\SEC Renndaten\\2019_Expert";
            ofd1.RestoreDirectory = true;
            ofd1.Multiselect = true;
            if (DialogResult.OK == ofd1.ShowDialog())
            {
                files = ofd1.FileNames;
                comboBox1.Items.Clear();
                comboBox1.Items.AddRange(files);
                comboBox1.SelectedIndex = 0;
            }


            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;

            jahresrangliste.Clear();
            läufe.Clear();
            dataGridView1.Rows.Clear();

            /*try
            {*/


            //Start Excel and get Application object.
            oXL = new Excel.Application();
            oXL.Visible = false;



            // Daten einlesen

            foreach (string lauf in files)
            {
                bool added = false;
                int rang = 0;
                string nachname;
                string vorname;
                Rennfahrer rennfahrer;

                int rCnt;
                int cCnt;
                int rw = 0;
                int cl = 0;

                int platzSpalte = 1;
                int vornameSpalte = 4;
                int nachnameSpalte = 3;

                string[] words = lauf.Split('_');

                string laufName = words[words.Length - 1].Split('.').First();
                int laufNummer = Convert.ToInt16(words[words.Length - 2]);
                klasse = words[words.Length - 3];
                jahr = words[words.Length - 4].Split('\\').Last();

                path = Path.GetDirectoryName(lauf);


                label4.Text = "Jahresrangliste " + klasse + " " + jahr;

                läufe.Add(new Lauf(laufName, laufNummer));
                /*try
                {*/

                //Get a the workbook.
                oWB = oXL.Workbooks.Open(lauf, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                oSheet = (Excel.Worksheet)oWB.Worksheets.get_Item(1);

                oRng = oSheet.UsedRange;
                rw = oRng.Rows.Count;
                cl = oRng.Columns.Count;

                for (cCnt = 1; cCnt <= cl; cCnt++)
                {
                    string header = (string)(oRng.Cells[1, cCnt] as Excel.Range).Value2;
                    if (header == "Platz") platzSpalte = cCnt;
                    if (header == "Nachname") nachnameSpalte = cCnt;
                    if (header == "Vachname") vornameSpalte = cCnt;
                }

                for (rCnt = 2; rCnt <= rw; rCnt++)
                {
                    added = false;
                    if (Convert.ToInt16((oRng.Cells[rCnt, 1] as Excel.Range).Value2) != 0) rang = Convert.ToInt16((oRng.Cells[rCnt, 1] as Excel.Range).Value2); //Bei mehreren gleichen Rängen den Rang von vorher übernehmen
                    nachname = (string)(oRng.Cells[rCnt, nachnameSpalte] as Excel.Range).Value2;
                    vorname = (string)(oRng.Cells[rCnt, vornameSpalte] as Excel.Range).Value2;

                    foreach (Rennfahrer r in jahresrangliste)
                    {
                        if (r.nachname == nachname && r.vorname == vorname)
                        {
                            r.Add(laufName, rangZuPunkte(rang));
                            added = true;
                            break;
                        }
                    }

                    if (added == false)
                    {
                        rennfahrer = new Rennfahrer(vorname, nachname, läufeZuGewerteteLäufe(files.Length));
                        rennfahrer.Add(laufName, rangZuPunkte(rang));
                        jahresrangliste.Add(rennfahrer);
                    }

                }

                oWB.Close();
                /*}
                catch
                {

                }*/
            }

            // Daten verarbeiten
            jahresrangliste.Sort();
            jahresrangliste.Reverse();

            läufe.Sort();

            // Daten ausgeben             

            dataGridView1.ColumnCount = 4 + läufe.Count;
            dataGridView1.Columns[0].Name = "Rang";
            dataGridView1.Columns[1].Name = "Nachname";
            dataGridView1.Columns[2].Name = "Vorname";



            for (int i = 0; i < läufe.Count; i++)
            {
                dataGridView1.Columns[i + 3].Name = läufe[i].name;
            }

            dataGridView1.Columns[dataGridView1.ColumnCount - 1].Name = "Total";

            foreach (Rennfahrer r in jahresrangliste)
            {
                List<string> row = new List<string> { (jahresrangliste.IndexOf(r) + 1).ToString(), r.nachname, r.vorname };
                for (int i = 0; i < läufe.Count; i++)
                {
                    bool found = false;
                    for (int n = 0; n < r.läufe.Count; n++)
                    {
                        if (r.läufe[n].name == läufe[i].name)
                        {
                            found = true;
                            row.Add(r.läufe[n].punkte.ToString());
                            break;
                        }
                    }
                    if (found == false) row.Add(null);
                }
                row.Add(r.getPoints().ToString());
                dataGridView1.Rows.Add(row.ToArray<string>());
            }

            /*}
            catch
            {

            }*/
        }


        private void button1_Click(object sender, EventArgs e)
        {
            // Daten speichern
            //try { 
            // creating Excel Application  
            Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            // creating new WorkBook within Excel application  
            Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            // creating new Excelsheet in workbook  
            Excel._Worksheet worksheet = null;
            // see the excel sheet behind the program  
            app.Visible = false;
            // get the reference of first sheet. By default its name is Sheet1.  
            // store its reference to worksheet  
            worksheet = workbook.Worksheets.get_Item(1);
            worksheet = workbook.ActiveSheet;
            // changing the name of active sheet  
            worksheet.Name = "Jahresrangliste";
            // storing title in Excel
            worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, dataGridView1.Columns.Count]].Merge(false);
            worksheet.Cells[1, 1] = label4.Text;
            worksheet.Cells[1, 1].Font.Bold = true;
            worksheet.Cells[1, 1].Font.Size = 20;
            worksheet.Cells[1, 1].Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            worksheet.Cells[1, 1].Interior.Color = ColorTranslator.ToOle(Color.Gray);
            worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, dataGridView1.Columns.Count]].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            // storing header part in Excel  
            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {
                if (i > 3 && i < dataGridView1.Columns.Count) worksheet.Cells[3, i] = "Lauf " + (i - 3).ToString();
                worksheet.Cells[2, i].Font.Bold = true;
                worksheet.Cells[2, i].Font.Size = 12;
                worksheet.Cells[2, i].Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                worksheet.Cells[2, i].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);



                if (i > 3 && i < dataGridView1.Columns.Count) worksheet.Cells[2, i] = dataGridView1.Columns[i - 1].HeaderText;
                else worksheet.Cells[3, i] = dataGridView1.Columns[i - 1].HeaderText;
                worksheet.Cells[3, i].Font.Bold = true;
                worksheet.Cells[3, i].Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                worksheet.Cells[3, i].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

            }
            // storing Each row and column value to excel sheet  
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    try
                    {
                        worksheet.Cells[i + 4, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                    catch //Fängt Ausnahmen ab wenn Zellen leer sind
                    {

                    }
                    worksheet.Cells[i + 4, j + 1].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                }
            }
            worksheet.Range[worksheet.Cells[dataGridView1.Rows.Count + 4, 1], worksheet.Cells[dataGridView1.Rows.Count + 4, dataGridView1.Columns.Count]].Merge(false);
            worksheet.Cells[dataGridView1.Rows.Count + 4, 1] = DateTime.Today.ToString("s");
            worksheet.UsedRange.Columns.AutoFit();

            // save the application  
            workbook.SaveAs(path + "\\" + jahr + "_" + klasse + "_0_Jahresrangliste" + ".xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);


            // export PDF
            worksheet.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, path + "\\" + jahr + "_" + klasse + "_0_Jahresrangliste" + ".pdf");

            // Exit from the application  
            app.Quit();

            /*
            catch
            {
            }
            */

        }
    }

    public class Rennfahrer : IComparable
    {
        private int anzahlGewerteteLäufe;
        public string vorname;
        public string nachname;
        public List<Lauf> läufe;

        public Rennfahrer(string vorname,string nachname,int anzahlGewerteteLäufe)
        {
            this.anzahlGewerteteLäufe = anzahlGewerteteLäufe;
            this.vorname = vorname;
            this.nachname = nachname;
            läufe = new List<Lauf>();
        }

        public int CompareTo(object obj)
        {
            Rennfahrer that = obj as Rennfahrer;
            int res;
            if (that == null)
            {
                return 0;
            }

            // "CompareTo()" method 
            res = this.getPoints().CompareTo(that.getPoints());

            if (res == 0)
            {
                int lthis = this.läufe.Count;
                int lthat = that.läufe.Count;

                for (int c = anzahlGewerteteLäufe + 1; c <= lthis && c <= lthat; c++)
                {
                    res = this.getPoints(c).CompareTo(that.getPoints(c));
                    if (res != 0) break;
                }
            }

            return res;
        }

        public void Add(string name, int punkte)
        {
            this.läufe.Add(new Lauf(name, punkte));
            läufe.Sort();
            läufe.Reverse();
        }

        public int getPoints()
        {
            return getPoints(this.anzahlGewerteteLäufe);
        }
        public int getPoints(int anzahlGewerteteLäufe)
        {
            int sum = 0;
            if (anzahlGewerteteLäufe > this.läufe.Count) anzahlGewerteteLäufe = this.läufe.Count;
            for(int i=0;i < anzahlGewerteteLäufe; i++)
            {
                sum += läufe[i].punkte;
            }
            return sum;
        }
    }

    public class Lauf : IComparable
    {
        public string name;
        public int punkte;

        public Lauf(string name, int punkte)
        {
            this.name = name;
            this.punkte = punkte;
        }

        public int CompareTo(object obj)
        {
            Lauf that = obj as Lauf;
            return this.punkte.CompareTo(that.punkte);
        }
    }
}
