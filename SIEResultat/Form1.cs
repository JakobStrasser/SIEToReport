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
using System.Text.RegularExpressions;
using Microsoft.Office.Interop;
using Excel = Microsoft.Office.Interop.Excel;



namespace SIEResultat
{
    public partial class Form1 : Form
    { 
        
        Microsoft.Office.Interop.Excel.Application myApp;
        private List<string> lista;

        private SortedDictionary<string, string> konton = new SortedDictionary<string, string>();
        private Dictionary<string, string> dimensioner = new Dictionary<string, string>();
        private List<Objekt> objekten = new List<Objekt>();
        public List<string> valda = new List<string>();
        private List<Transaktion> transaktioner = new List<Transaktion>();
        private string regexDelning = "([^\"]\\S*|\".+?\"|[^\\s*$])\\s*";
        private string fNamn = "";
        Excel.Workbook wb;


        public Form1()
        {
            InitializeComponent();
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;


        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string fileContents = "";

                OpenFileDialog openFileDialog1 = new OpenFileDialog
                {
                    Filter = "SIE 4|*.SI",
                    Title = "Select a SIE File"
                };

                // Show the Dialog.  
                // If the user clicked OK in the dialog and  
                // a .CUR file was selected, open it.  
                if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    // Assign the cursor in the Stream to the Form's Cursor property.  
                    // REGEXP ("([^\"]\\S*|\".+?\")\\s*")   

                    byte[] indata = File.ReadAllBytes(openFileDialog1.FileName);

                    fileContents = convertUnicode(indata);
                    //fileContents = indata.ToString();

                }
                if (fileContents.Length > 1)
                {
                    lista = new List<string>(Regex.Split(fileContents, Environment.NewLine));

                }
                else
                {
                    throw new Exception("Fel vid filinläsning");
                }
                // MessageBox.Show("Rader: " + list.Count);
                button3.Enabled = true;
                this.Update();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fel: " + ex.Message + " \n" + ex.StackTrace);
            }
        }
        private void SaveTransactions()
        {

            try { 

            myApp = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook wb = myApp.Workbooks.Add();
            Excel.Worksheet sheet;
            sheet = wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]);



            sheet.Name = "Transaktioner";
            sheet.Activate();
            sheet.Cells[1, 1] = "Konto";
            sheet.Cells[1, 2] = "Resultatenhet";
            sheet.Cells[1, 3] = "Projekt";
            sheet.Cells[1, 4] = "Belopp";
            sheet.Cells[1, 5] = "Transaktionsdatum";
            // sheet.Cells[1, 6] = "Månad";

            sheet.Cells[1, 6] = "Transaktionstext";
            int r = 2;
            toolStripProgressBar1.Value = 1;
            toolStripProgressBar1.Maximum = transaktioner.Count;

            foreach (Transaktion t in transaktioner)
            {
                toolStripStatusLabel1.Text = "Skapar exceldokument rad: " + toolStripProgressBar1.Value + " av " + toolStripProgressBar1.Maximum;
                toolStripProgressBar1.Increment(1);
                sheet.Cells[r, 1] = t.Kontonr;
                if (t.Objekt.ContainsKey("1"))
                    sheet.Cells[r, 2] = "'" + t.Objekt["1"];
                if (t.Objekt.ContainsKey("6"))
                    sheet.Cells[r, 3] = "'" + t.Objekt["6"];
                sheet.Cells[r, 4] = t.Belopp;
                DateTime datum = t.Transaktionsdatum;
                sheet.Cells[r, 5] = datum;
                //sheet.Cells[r, 5] = t.Transaktionsdatum.Substring(0, 4) + "-" + t.Transaktionsdatum.Substring(4, 2);
                sheet.Cells[r, 6] = t.Transtext;
                r++;
            }
            r--;
            Microsoft.Office.Interop.Excel.Range ran = sheet.Range["A2", "G" + r.ToString()];
            ran.Sort(ran.Columns[5], Excel.XlSortOrder.xlAscending);
            //sheet.Names.Add("Transaktioner", ran);

            sheet = wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]);
            r = 2;
            sheet.Name = "Konton";
            sheet.Cells[1, 1] = "Konto";
            sheet.Cells[1, 2] = "Kontobenämning";
            foreach (var key in konton.Keys)
            {
                sheet.Cells[r, 1] = key.ToString();
                sheet.Cells[r, 2] = konton[key];
                r++;
            }
            r--;
            Microsoft.Office.Interop.Excel.Range ran2 = sheet.Range["A2", "B" + r.ToString()];
            ran2.Sort(ran2.Columns[1], Excel.XlSortOrder.xlAscending);
            //sheet.Names.Add("Konton", ran);
            sheet = wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]);
            r = 2;
            sheet.Name = "Objekt";
            sheet.Cells[1, 1] = "Typ";
            sheet.Cells[1, 2] = "Objekt";
            sheet.Cells[1, 3] = "Benämning";
            objekten.Sort();
            foreach (Objekt o in objekten)
            {

                sheet.Cells[r, 1] = o.Typ;
                sheet.Cells[r, 2] = "'" + o.Id;
                sheet.Cells[r, 3] = o.Namn;

                r++;
            }
            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Excel | *.xlsx"
            };
            sfd.ShowDialog();
            wb.SaveAs(sfd.FileName);
            wb.Close();
            myApp.Quit();
            toolStripStatusLabel1.Text = "Excelfil sparad som " + sfd.FileName;
        }
            catch (Exception ex)
            {
                myApp.Quit();
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace);
            }
}
        private void button2_Click(object sender, EventArgs e)
        {

            try
            {
                myApp = new Microsoft.Office.Interop.Excel.Application();

                wb = myApp.Workbooks.Add();
            
               
                objekten.Sort();
                
              
                foreach(Objekt o in objekten)
                {
                    if (valda.Contains(o.Id))
                    {
                        Total(wb, o);
                      
                        toolStripStatusLabel1.Text = "Skapar flik " + o.Id + " " + o.Namn;
                    }
                }
                toolStripStatusLabel1.Text = "Skapar flik TOTAL";
                Total(wb, null);
               
                wb.Sheets["Blad1"].Delete();
                wb.Sheets["Blad2"].Delete();
                wb.Sheets["Blad3"].Delete();
               

                SaveFileDialog sfd = new SaveFileDialog
                {
                    Filter = "Excel | *.xlsx"
                };
                toolStripStatusLabel1.Text = "Excelfil sparad som " + sfd.FileName;
                sfd.ShowDialog();
                wb.SaveAs(sfd.FileName);
                myApp.Quit();
            }
            catch (Exception ex)
            {
                myApp.Quit();
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace);
            }
        }

        private string convertUnicode(byte[] indata)
        {
            try
            {

                string utdata = "";



                Encoding asciiEncoding = Encoding.GetEncoding(437);
                utdata = asciiEncoding.GetString(indata);
                return utdata;


            }
            catch (Exception ex)
            {

                MessageBox.Show("Fel: " + ex.Message + " \n" + ex.StackTrace);
                return "";
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            toolStripProgressBar1.Maximum = lista.Count;
            string verdatum = "";
            string vertext = "";
            List<string> monthList = new List<string>();
            string tokenString = "";
            toolStripProgressBar1.Value = 0;
            try
            {
                foreach (string currentRow in lista)
                {
                    int hashMark = currentRow.IndexOf("#");
                    if (hashMark >= 0)
                    {
                        int nextSpace = currentRow.Substring(hashMark).IndexOf(" ") + hashMark;
                        tokenString = currentRow.Substring(hashMark, nextSpace - hashMark);
                        //textBox1.Text += Environment.NewLine + tokenString;
                        
                        switch (tokenString)
                        {
                            case "#FNAMN":
                                {
                                   
                                    string[] konto = Regex.Split(currentRow, regexDelning);
                                    fNamn = konto[3].ToString().Replace('"',' ').Trim();
                                    this.label2.Text = "Företag: " + fNamn;
                                    break;
                                }
                          
                            case "#KONTO":
                                {
                                    char[] charSeparators = new char[] { ' ' };
                                    string[] konto = Regex.Split(currentRow, regexDelning);
                                    List<string> rensakonton = new List<string>();
                                    foreach (string s in konto)
                                    {
                                        if (!s.Equals(""))
                                            rensakonton.Add(s);

                                    }
                                    if (!konton.ContainsKey(rensakonton[1].Replace('"', ' ').Trim()))
                                        konton.Add(rensakonton[1].Replace('"', ' ').Trim(), rensakonton[2].Replace('"', ' ').Trim());
                                    break;
                                }
                            case "#DIM":
                                {
                                    char[] charSeparators = new char[] { ' ' };
                                    string[] dim = Regex.Split(currentRow, regexDelning);
                                    List<string> rensaDim = new List<string>();
                                    foreach (string s in dim)
                                    {
                                        if (!s.Equals(""))
                                            rensaDim.Add(s);

                                    }
                                    if (!dimensioner.ContainsKey(rensaDim[1].Replace('"', ' ').Trim()))
                                        dimensioner.Add(rensaDim[1].Replace('"', ' ').Trim(), rensaDim[2].Replace('"', ' ').Trim());
                                    break;
                                }
                            case "#OBJEKT":
                                {
                                    char[] charSeparators = new char[] { ' ' };
                                    string[] dim = Regex.Split(currentRow, regexDelning);
                                    List<string> rensaDim = new List<string>();
                                    
                                    foreach (string s in dim)
                                    {
                                        if (!s.Equals(""))
                                            rensaDim.Add(s);

                                    }
                                    string obj_typ = rensaDim[1].Replace('"', ' ').Trim();
                                    string obj_id = rensaDim[2].Replace('"', ' ').Trim();
                                    string obj_namn = rensaDim[3].Replace('"', ' ').Trim();
                                    bool add = true;
                                    foreach (Objekt o in objekten)
                                    {
                                        if (o.Typ.Equals(obj_typ) & o.Id.Equals(obj_id))
                                            add = false;

                                    }
                                    if (add)
                                    {
                                        objekten.Add(new Objekt(obj_typ, obj_id, obj_namn));
                                        
                                    }
                                    break;
                                }
                            case "#TRANS":
                                {
                                    string[] dim;
                                    Dictionary<string, string> objekt = new Dictionary<string, string>();
                                    string brackets = Regex.Match(currentRow, "({.*})").ToString();
                                    brackets = brackets.Substring(1, brackets.Length - 2);

                                    if (brackets.Length > 0)
                                    {
                                        dim = Regex.Split(brackets, regexDelning);
                                        for (int i = 1; i < dim.Length - 3; i = i + 4)
                                        {

                                            objekt.Add(dim[i].Replace('"', ' ').Trim(), dim[i + 2].Replace('"', ' ').Trim());
                                        }
                                    }
                                    string utanObjekt = currentRow.Substring(1, currentRow.IndexOf('{') - 1) + currentRow.Substring(currentRow.IndexOf('}') + 1);
                                    utanObjekt = utanObjekt.Trim();
                                    string[] delar = Regex.Split(utanObjekt, regexDelning);
                                    List<string> rensaDelar = new List<string>();
                                    foreach (string s in delar)
                                    {
                                        if (!s.Equals(""))
                                            rensaDelar.Add(s);
                                    }

                                    String transaktionsdatum = "";
                                    if (transaktionsdatum.Equals(""))
                                        transaktionsdatum = verdatum;
                                    if (!monthList.Contains(transaktionsdatum.Substring(0, 4)+"-"+transaktionsdatum.Substring(4,2)))
                                        monthList.Add(transaktionsdatum.Substring(0, 4) + "-" + transaktionsdatum.Substring(4, 2));
                                    String kontonr = rensaDelar[1].Replace('"', ' ').Trim();
                                    string b = rensaDelar[2].Replace('.', ',').Replace('"', ' ').Trim();
                                    double belopp = -double.Parse(b);
                                    string transtext = "";
                                    if (transtext.Equals(""))
                                        transtext = vertext;
                                    double kvantitet = 0.0;
                                    string sign = "";
                                    DateTime transaktionsdate = DateTime.Parse(transaktionsdatum.Substring(0, 4) + "-" + transaktionsdatum.Substring(4, 2) + "-" + transaktionsdatum.Substring(6, 2));
                                    transaktioner.Add(new Transaktion(objekt, transaktionsdate, kontonr, belopp, transtext, kvantitet, sign));
                                    break;
                                }
                            case "#VER":
                                {
                                    char[] charSeparators = new char[] { ' ' };
                                    string[] dim = Regex.Split(currentRow, regexDelning);
                                    verdatum = dim[7].Replace('"', ' ').Trim();
                                    vertext = dim[9].Replace('"', ' ').Trim();
                                    break;
                                }
                            default:
                                {
                                    break;
                                }
                        } //switch
                    } //if
                   
                   
                    this.Update();
                    toolStripProgressBar1.Increment(1);
                    toolStripStatusLabel1.Text = "Bearbetar rad " + toolStripProgressBar1.Value + " av " + toolStripProgressBar1.Maximum;
                } //foreach

                monthList.Sort();
                foreach(string s in monthList)
                {
                    comboBox1.Items.Add(s);
                  
                }
                comboBox1.SelectedItem = comboBox1.Items[comboBox1.Items.Count - 1];
                button4.Enabled = true;
                SelectResultatenhet();

            }//try
            catch (Exception ex)
            {
              
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace);
            }//Catch

        } //button3_click

        private void SelectResultatenhet()
        {
            foreach (Objekt o in objekten)
            {
                if (o.Typ.Equals("1"))
                    checkedListBox1.Items.Add(o.Id);

            }

            for(int i = 0; i<checkedListBox1.Items.Count;i++)
            {
                checkedListBox1.SetItemChecked(i, true);
            }
           
        }


        private void Total(Excel.Workbook wb, Objekt o)
        {
            try
            {

                string resenh;
                if (o == null)
                    resenh = "*";
                else
                    resenh = o.Id;
                DateTime startMonth = DateTime.Parse(comboBox1.SelectedItem.ToString() + "-01");
                DateTime endMonth = endOfMonth(DateTime.Parse(comboBox1.SelectedItem.ToString() + "-01"));
                DateTime startYear = new DateTime(startMonth.Year, 5, 1);
                DateTime startPreviousYear = new DateTime(startMonth.Year - 1, 5, 1);
                DateTime endPreviousYear = new DateTime(startMonth.Year - 1, startMonth.AddYears(-1).Month, endOfMonth(startMonth.AddYears(-1)).Day);
                DateTime startPrevMonth = startMonth.AddMonths(-1);
                DateTime endPrevMonth = endOfMonth(startMonth.AddMonths(-1));
                int intaktsRad;
                int rorelseKostnadsRad;
                int bruttoRad;
                List<int> kostnadsRader = new List<int>();
                
                int kostnadsRad;
               


                Excel.Worksheet totalSheet = wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]);
                if (resenh.Contains("*"))
                    totalSheet.Name = "TOTAL";
                else
                    totalSheet.Name = resenh;
                totalSheet.Cells[1, 2] = fNamn;
                totalSheet.Cells[1, 3] = resenh;
                if(o != null)
                    totalSheet.Cells[1, 4] = o.Namn;


                totalSheet.Cells[2, 1] = "Från:";
                totalSheet.Cells[2, 2] = startMonth;
                totalSheet.Cells[3, 1] = "Till:";
                totalSheet.Cells[3, 2] = endMonth;

                //Skapa tabellerna
                Dictionary<string, double> sumMonth = SumTransaktion(startMonth, endMonth, resenh);
                Dictionary<string, double> sumLastMonth = SumTransaktion(startPrevMonth, endPrevMonth, resenh);
                Dictionary<string, double> sumYTD = SumTransaktion(startYear, endMonth, resenh);
                Dictionary<string, double> sumLastYTD = SumTransaktion(startPreviousYear, endPreviousYear, resenh);


                int row = 6;
                totalSheet.Cells[5, 1] = "Konto";
                totalSheet.Cells[5, 2] = "Benämning";
                totalSheet.Cells[5, 3] = "Månadens resultat";
                totalSheet.Cells[5, 4] = "Fg månads resultat";
                totalSheet.Cells[5, 5] = "Differens";
                totalSheet.Cells[5, 7] = "Ack resultat";
                totalSheet.Cells[5, 8] = "Ack resultat fg år";
                totalSheet.Cells[5, 9] = "Differens";

                row++;
                totalSheet.Cells[row, 2] = "Intäkter";
                row++;
                //3000-3999
                int startrow = row;
                int[] sorteradeKonton = new int[konton.Count];
                int i = 0;
                foreach(string s in konton.Keys)
                {
                    sorteradeKonton[i] = Int32.Parse(s.Trim());
                    i++;
                }

                Array.Sort<int>(sorteradeKonton);
            
                foreach (int s in sorteradeKonton)
                {

                    int konto = s;
                    if (konto >= 3000 && konto <= 3999)
                    {
                        totalSheet.Cells[row, 1] = konto;
                        totalSheet.Cells[row, 2] = konton[s.ToString()];

                        if (sumMonth.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 3] = sumMonth[konto.ToString()];
                        else
                            totalSheet.Cells[row, 3] = 0.0;
                        if (sumLastMonth.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 4] = sumLastMonth[konto.ToString()];
                        else
                            totalSheet.Cells[row, 4] = 0.0;
                         if (sumMonth.ContainsKey(konto.ToString()) & sumLastMonth.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 5] = sumMonth[konto.ToString()] - sumLastMonth[konto.ToString()];
                         else
                            totalSheet.Cells[row, 5] = 0.0;

                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 7] = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells[row, 7] = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 8] = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells[row, 8] = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 9] = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells[row, 9] = 0.0;
                        row++;
                    }
                }
                String rangeStr = startrow + ":" + (row-1);
                Excel.Range myRange = totalSheet.Rows[rangeStr] as Excel.Range;
                myRange.Group();
                if (startrow == row)
                    row++;
                totalSheet.Cells[row, 2] = "Summa intäkter";
                totalSheet.Cells[row, 3] = "=SUM(C" + startrow + ":C" + (row - 1);
                totalSheet.Cells[row, 4] = "=SUM(D" + startrow + ":D" + (row - 1);
                totalSheet.Cells[row, 5] = "=SUM(E" + startrow + ":E" + (row - 1);
                totalSheet.Cells[row, 7] = "=SUM(G" + startrow + ":G" + (row - 1);
                totalSheet.Cells[row, 8] = "=SUM(H" + startrow + ":H" + (row - 1);
                totalSheet.Cells[row, 9] = "=SUM(I" + startrow + ":I" + (row - 1);
                intaktsRad = row;



                row++;
                row++;

                //4000-4999 Rörelsens kostnader
                totalSheet.Cells[row, 2] = "Rörelsens kostnader";
                row++;
                startrow = row;
                foreach (string s in konton.Keys)
                {

                    int konto = Int32.Parse(s);
                    if (konto >= 4000 && konto <= 4999)
                    {
                        totalSheet.Cells[row, 1] = konto;
                        totalSheet.Cells[row, 2] = konton[s];

                        if (sumMonth.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 3] = sumMonth[konto.ToString()];
                        else
                            totalSheet.Cells[row, 3] = 0.0;
                        if (sumLastMonth.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 4] = sumLastMonth[konto.ToString()];
                        else
                            totalSheet.Cells[row, 4] = 0.0;
                        if (sumMonth.ContainsKey(konto.ToString()) & sumLastMonth.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 5] = sumMonth[konto.ToString()] - sumLastMonth[konto.ToString()];
                        else
                            totalSheet.Cells[row, 5] = 0.0;

                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 7] = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells[row, 7] = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 8] = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells[row, 8] = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 9] = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells[row, 9] = 0.0;
                        row++;
                    }
                }
                rangeStr = startrow + ":" + (row -1);
                myRange = totalSheet.Rows[rangeStr] as Excel.Range;
                myRange.Group();
                if (startrow == row)
                    row++;
                totalSheet.Cells[row, 2] = "Summa rörelsens kostnader";
                totalSheet.Cells[row, 3] = "=SUM(C" + startrow + ":C" + (row - 1);
                totalSheet.Cells[row, 4] = "=SUM(D" + startrow + ":D" + (row - 1);
                totalSheet.Cells[row, 5] = "=SUM(E" + startrow + ":E" + (row - 1);
                totalSheet.Cells[row, 7] = "=SUM(G" + startrow + ":G" + (row - 1);
                totalSheet.Cells[row, 8] = "=SUM(H" + startrow + ":H" + (row - 1);
                totalSheet.Cells[row, 9] = "=SUM(I" + startrow + ":I" + (row - 1);
                rorelseKostnadsRad = row;
                row++;
                row++;
                bruttoRad = row;
                totalSheet.Cells[row, 2] = "Bruttovinst";
                totalSheet.Cells[row, 3] = "=C" + intaktsRad + "+C" + rorelseKostnadsRad;
                totalSheet.Cells[row, 4] = "=D" + intaktsRad + "+D" + rorelseKostnadsRad;
                totalSheet.Cells[row, 5] = "=E" + intaktsRad + "+E" + rorelseKostnadsRad;
                totalSheet.Cells[row, 7] = "=G" + intaktsRad + "+G" + rorelseKostnadsRad;
                totalSheet.Cells[row, 8] = "=H" + intaktsRad + "+H" + rorelseKostnadsRad;
                totalSheet.Cells[row, 9] = "=I" + intaktsRad + "+I" + rorelseKostnadsRad;
                row++;
                row++;


                //5000-6999 Externa kostnader
                totalSheet.Cells[row, 2] = "Externa kostnader";
                row++;
                startrow = row;
                foreach (string s in konton.Keys)
                {

                    int konto = Int32.Parse(s);
                    if (konto >= 5000 && konto <= 6999)
                    {
                        totalSheet.Cells[row, 1] = konto;
                        totalSheet.Cells[row, 2] = konton[s];

                        if (sumMonth.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 3] = sumMonth[konto.ToString()];
                        else
                            totalSheet.Cells[row, 3] = 0.0;
                        if (sumLastMonth.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 4] = sumLastMonth[konto.ToString()];
                        else
                            totalSheet.Cells[row, 4] = 0.0;
                        if (sumMonth.ContainsKey(konto.ToString()) & sumLastMonth.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 5] = sumMonth[konto.ToString()] - sumLastMonth[konto.ToString()];
                        else
                            totalSheet.Cells[row, 5] = 0.0;

                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 7] = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells[row, 7] = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 8] = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells[row, 8] = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 9] = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells[row, 9] = 0.0;

                        row++;
                    }
                }
                rangeStr = startrow + ":" + (row -1);
                myRange = totalSheet.Rows[rangeStr] as Excel.Range;
                myRange.Group();
                if (startrow == row)
                    row++;
                kostnadsRader.Add(row);
                totalSheet.Cells[row, 2] = "Summa externa kostnader";
                totalSheet.Cells[row, 3] = "=SUM(C" + startrow + ":C" + (row - 1);
                totalSheet.Cells[row, 4] = "=SUM(D" + startrow + ":D" + (row - 1);
                totalSheet.Cells[row, 5] = "=SUM(E" + startrow + ":E" + (row - 1);
                totalSheet.Cells[row, 7] = "=SUM(G" + startrow + ":G" + (row - 1);
                totalSheet.Cells[row, 8] = "=SUM(H" + startrow + ":H" + (row - 1);
                totalSheet.Cells[row, 9] = "=SUM(I" + startrow + ":I" + (row - 1);
                row++;
                
                row++;


                //7000-7799 Personalkostnader
                totalSheet.Cells[row, 2] = "Personalkostnader";
                row++;
                startrow = row;
                foreach (string s in konton.Keys)
                {

                    int konto = Int32.Parse(s);
                    if (konto >= 7000 && konto <= 7799)
                    {
                        totalSheet.Cells[row, 1] = konto;
                        totalSheet.Cells[row, 2] = konton[s];

                        if (sumMonth.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 3] = sumMonth[konto.ToString()];
                        else
                            totalSheet.Cells[row, 3] = 0.0;
                        if (sumLastMonth.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 4] = sumLastMonth[konto.ToString()];
                        else
                            totalSheet.Cells[row, 4] = 0.0;
                        if (sumMonth.ContainsKey(konto.ToString()) & sumLastMonth.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 5] = sumMonth[konto.ToString()] - sumLastMonth[konto.ToString()];
                        else
                            totalSheet.Cells[row, 5] = 0.0;

                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 7] = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells[row, 7] = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 8] = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells[row, 8] = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 9] = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells[row, 9] = 0.0;
                        row++;
                    }
                }
                rangeStr = startrow + ":" + (row -1);
                myRange = totalSheet.Rows[rangeStr] as Excel.Range;
                myRange.Group();
                if (startrow == row)
                    row++;
                kostnadsRader.Add(row);
                totalSheet.Cells[row, 2] = "Summa personalkostnader";
                totalSheet.Cells[row, 3] = "=SUM(C" + startrow + ":C" + (row - 1);
                totalSheet.Cells[row, 4] = "=SUM(D" + startrow + ":D" + (row - 1);
                totalSheet.Cells[row, 5] = "=SUM(E" + startrow + ":E" + (row - 1);
                totalSheet.Cells[row, 7] = "=SUM(G" + startrow + ":G" + (row - 1);
                totalSheet.Cells[row, 8] = "=SUM(H" + startrow + ":H" + (row - 1);
                totalSheet.Cells[row, 9] = "=SUM(I" + startrow + ":I" + (row - 1);
                row++;
               
                row++;

                //7800-7999 Avskrivningar
                totalSheet.Cells[row, 2] = "Avskrivningar";
                row++;
                startrow = row;
                foreach (string s in konton.Keys)
                {

                    int konto = Int32.Parse(s);
                    if (konto >= 7800 && konto <= 7999)
                    {
                        totalSheet.Cells[row, 1] = konto;
                        totalSheet.Cells[row, 2] = konton[s];

                        if (sumMonth.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 3] = sumMonth[konto.ToString()];
                        else
                            totalSheet.Cells[row, 3] = 0.0;
                        if (sumLastMonth.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 4] = sumLastMonth[konto.ToString()];
                        else
                            totalSheet.Cells[row, 4] = 0.0;
                        if (sumMonth.ContainsKey(konto.ToString()) & sumLastMonth.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 5] = sumMonth[konto.ToString()] - sumLastMonth[konto.ToString()];
                        else
                            totalSheet.Cells[row, 5] = 0.0;

                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 7] = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells[row, 7] = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 8] = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells[row, 8] = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 9] = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells[row, 9] = 0.0;
                        row++;
                    }
                }
                rangeStr = startrow + ":" + (row -1);
                myRange = totalSheet.Rows[rangeStr] as Excel.Range;
                myRange.Group();
                if (startrow == row)
                    row++;
                kostnadsRader.Add(row);
                totalSheet.Cells[row, 2] = "Summa avskrivningar";
                totalSheet.Cells[row, 3] = "=SUM(C" + startrow + ":C" + (row - 1);
                totalSheet.Cells[row, 4] = "=SUM(D" + startrow + ":D" + (row - 1);
                totalSheet.Cells[row, 5] = "=SUM(E" + startrow + ":E" + (row - 1);
                totalSheet.Cells[row, 7] = "=SUM(G" + startrow + ":G" + (row - 1);
                totalSheet.Cells[row, 8] = "=SUM(H" + startrow + ":H" + (row - 1);
                totalSheet.Cells[row, 9] = "=SUM(I" + startrow + ":I" + (row - 1);
                row++;
               
                row++;

                //8000-8799 Finansiella kostnader
                totalSheet.Cells[row, 2] = "Finansiella kostnader";
                row++;
                startrow = row;
                foreach (string s in konton.Keys)
                {

                    int konto = Int32.Parse(s);
                    if (konto >= 8000 && konto <= 8799)
                    {
                        totalSheet.Cells[row, 1] = konto;
                        totalSheet.Cells[row, 2] = konton[s];

                        if (sumMonth.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 3] = sumMonth[konto.ToString()];
                        else
                            totalSheet.Cells[row, 3] = 0.0;
                        if (sumLastMonth.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 4] = sumLastMonth[konto.ToString()];
                        else
                            totalSheet.Cells[row, 4] = 0.0;
                        if (sumMonth.ContainsKey(konto.ToString()) & sumLastMonth.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 5] = sumMonth[konto.ToString()] - sumLastMonth[konto.ToString()];
                        else
                            totalSheet.Cells[row, 5] = 0.0;

                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 7] = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells[row, 7] = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 8] = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells[row, 8] = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 9] = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells[row, 9] = 0.0;
                        row++;
                    }
                }
                rangeStr = startrow + ":" + (row -1);
                myRange = totalSheet.Rows[rangeStr] as Excel.Range;
                myRange.Group();
                if (startrow == row)
                    row++;
                kostnadsRader.Add(row);
                totalSheet.Cells[row, 2] = "Summa finansiella kostnader";
                totalSheet.Cells[row, 3] = "=SUM(C" + startrow + ":C" + (row - 1);
                totalSheet.Cells[row, 4] = "=SUM(D" + startrow + ":D" + (row - 1);
                totalSheet.Cells[row, 5] = "=SUM(E" + startrow + ":E" + (row - 1);
                totalSheet.Cells[row, 7] = "=SUM(G" + startrow + ":G" + (row - 1);
                totalSheet.Cells[row, 8] = "=SUM(H" + startrow + ":H" + (row - 1);
                totalSheet.Cells[row, 9] = "=SUM(I" + startrow + ":I" + (row - 1);
                row++;
                
                row++;

                //8800-8999 Bokslutsdispositioner
                totalSheet.Cells[row, 2] = "Bokslutsdispositioner";
                row++;
                startrow = row;
                foreach (string s in konton.Keys)
                {

                    int konto = Int32.Parse(s);
                    if (konto >= 8800 && konto <= 8999)
                    {
                        totalSheet.Cells[row, 1] = konto;
                        totalSheet.Cells[row, 2] = konton[s];

                        if (sumMonth.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 3] = sumMonth[konto.ToString()];
                        else
                            totalSheet.Cells[row, 3] = 0.0;
                        if (sumLastMonth.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 4] = sumLastMonth[konto.ToString()];
                        else
                            totalSheet.Cells[row, 4] = 0.0;
                        if (sumMonth.ContainsKey(konto.ToString()) & sumLastMonth.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 5] = sumMonth[konto.ToString()] - sumLastMonth[konto.ToString()];
                        else
                            totalSheet.Cells[row, 5] = 0.0;

                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 7] = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells[row, 7] = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 8] = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells[row, 8] = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells[row, 9] = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells[row, 9] = 0.0;
                        row++;
                    }
                }
                rangeStr = startrow + ":" + (row -1);
                myRange = totalSheet.Rows[rangeStr] as Excel.Range;
                myRange.Group();
               
                
                
                if (startrow == row)
                    row++;
                kostnadsRader.Add(row);
                totalSheet.Cells[row, 2] = "Summa bokslutsdispositioner";
                totalSheet.Cells[row, 3] = "=SUM(C" + startrow + ":C" + (row - 1);
                totalSheet.Cells[row, 4] = "=SUM(D" + startrow + ":D" + (row - 1);
                totalSheet.Cells[row, 5] = "=SUM(E" + startrow + ":E" + (row - 1);
                totalSheet.Cells[row, 7] = "=SUM(G" + startrow + ":G" + (row - 1);
                totalSheet.Cells[row, 8] = "=SUM(H" + startrow + ":H" + (row - 1);
                totalSheet.Cells[row, 9] = "=SUM(I" + startrow + ":I" + (row - 1);
                row++;
                row++;
                string summaKostnader = "=SUM(C";
                foreach (int n in kostnadsRader)
                {
                    summaKostnader += n + "+C";
                }
                summaKostnader = summaKostnader.Substring(0, summaKostnader.Length - 2) + ")";
                kostnadsRad = row;
                totalSheet.Cells[row, 2] = "Summa kostnader";
                totalSheet.Cells[row, 3] = summaKostnader;
                totalSheet.Cells[row, 4] = summaKostnader.Replace("C","D") ;
                totalSheet.Cells[row, 5] = summaKostnader.Replace("C", "E");
                totalSheet.Cells[row, 7] = summaKostnader.Replace("C", "G");
                totalSheet.Cells[row, 8] = summaKostnader.Replace("C", "H");
                totalSheet.Cells[row, 9] = summaKostnader.Replace("C", "I");


                row++;
                row++;
                totalSheet.Cells[row, 2] = "Resultat";
                totalSheet.Cells[row, 3] = "=C"+bruttoRad+"+C"+kostnadsRad;
                totalSheet.Cells[row, 4] = "=D" + bruttoRad + "+D" + kostnadsRad;
                totalSheet.Cells[row, 5] = "=E" + bruttoRad + "+E" + kostnadsRad;
                totalSheet.Cells[row, 7] = "=G" + bruttoRad + "+G" + kostnadsRad;
                totalSheet.Cells[row, 8] = "=H" + bruttoRad + "+H" + kostnadsRad;
                totalSheet.Cells[row, 9] = "=I" + bruttoRad + "+I" + kostnadsRad;


                totalSheet.Range["C7:I" + row].NumberFormat = "# ##0,;[Red]-# ##0,";
                totalSheet.Columns.AutoFit();
                totalSheet.Activate();
                totalSheet.Application.ActiveWindow.SplitRow = 6;
                totalSheet.Application.ActiveWindow.SplitColumn = 2;
                totalSheet.Application.ActiveWindow.FreezePanes = true;
                totalSheet.Outline.ShowLevels(1, 1);
                




            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace);
            }
        }

        private DateTime endOfMonth(DateTime datum)
        {
            int[] endDay;
            if (DateTime.IsLeapYear(datum.Year))
            {
                endDay = new int[]{ 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
            }
            else
            {
                 endDay = new int[]{ 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
            }
            if (DateTime.TryParse(datum.Year + "-" + datum.Month + "-" + endDay[datum.Month - 1], out DateTime returnDate))
                return returnDate;
            else
                return datum;
        }


        private void skapaTabeller(DateTime from, DateTime to)
        {
            Dictionary<string, double> sum = SumTransaktion(from, to);

        }

        private Dictionary<string, double> SumTransaktion(DateTime from, DateTime to)
        {
            Dictionary<string, double> sum = new Dictionary<string, double>();

            foreach (Transaktion t in transaktioner)
            {


                if (!sum.ContainsKey(t.Kontonr))
                    sum.Add(t.Kontonr, 0.0); //Add to dictionary if not already in it
                if (t.Transaktionsdatum >= from & t.Transaktionsdatum <= to)
                    sum[t.Kontonr] += t.Belopp; // Add t.Belopp to Sum if resenh == t.Id


            }


            return sum;
        }

        private Dictionary<string,double> SumTransaktion(DateTime from, DateTime to, string resenh)
        {
            Dictionary<string, double> sum = new Dictionary<string, double>();

            foreach(Transaktion t in transaktioner)
            {
                

                if (!sum.ContainsKey(t.Kontonr))
                    sum.Add(t.Kontonr, 0.0); //Add to dictionary if not already in it
                if (t.Objekt.ContainsKey("1") )
                    if(t.Objekt["1"].Equals(resenh) || resenh.Equals("*") )
                        if(t.Transaktionsdatum >= from & t.Transaktionsdatum <= to)
                            sum[t.Kontonr] += t.Belopp; // Add t.Belopp to Sum if resenh == t.Id


            }


            return sum;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            foreach(object o in checkedListBox1.CheckedItems)
            {
                
                valda.Add(o.ToString());
            }
            button2.Enabled = true;
            button5.Enabled = true;


        }

        private void button5_Click(object sender, EventArgs e)
        {
            SaveTransactions();
        }
    }//class
}//namespace

