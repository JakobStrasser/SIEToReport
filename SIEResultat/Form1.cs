
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;




namespace EPPlusResultat
{
    public partial class Form1 : Form
    {


        private List<string> lista;
        private readonly SortedDictionary<string, string> konton = new SortedDictionary<string, string>();
        private readonly Dictionary<string, string> dimensioner = new Dictionary<string, string>();
        private readonly List<Objekt> objekten = new List<Objekt>();
        private readonly List<string> valda = new List<string>();
        private readonly List<Transaktion> transaktioner = new List<Transaktion>();
        private readonly string regexDelning = "([^\"]\\S*|\".+?\"|[^\\s*$])\\s*";
        private string fNamn = "";
        private string selectedType = "";
        private ExcelPackage wb;


        public Form1()
        {
            InitializeComponent();
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;


        }

        ~Form1()
        {
            wb.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string fileContents = "";

                OpenFileDialog openFileDialog1 = new OpenFileDialog
                {
                    Filter = "SIE 4|*.SE",
                    Title = "Select a SIE File"
                };

                // Show the Dialog.  
                // If the user clicked OK in the dialog and  

                if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {



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
                button6.Enabled = true;
                Update();
                openFileDialog1.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fel: " + ex.Message + " \n" + ex.StackTrace);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            try
            {
                // myApp = new Microsoft.Office.Interop.Excel.Application();

                wb = new ExcelPackage();


                objekten.Sort();
                string valdTyp = comboBox2.SelectedItem.ToString();
                foreach (string v in valda)
                {
                    foreach (Objekt o in objekten)
                    {
                        if (o.Id.Equals(v) && dimensioner[o.Typ].Equals(valdTyp))
                        {
                            if (checkBox1.Checked)
                                YearlyTotal(wb, o);
                            else Total(wb, o);
                            Console.WriteLine("Skapar flik " + o.Id + " " + o.Namn);
                            toolStripStatusLabel1.Text = "Skapar flik " + o.Id + " " + o.Namn;
                        }

                    }


                }
                Console.WriteLine("Antal flikar " + wb.Workbook.Worksheets.Count);

                toolStripStatusLabel1.Text = "Skapar flik TOTAL";

                if (checkBox1.Checked)
                    YearlyTotal(wb, null);
                else Total(wb, null);
                SkapaResultatsida(wb);
                wb.Workbook.Worksheets["TOTAL"].Select();
                wb.Workbook.Calculate();

                SaveFileDialog sfd = new SaveFileDialog
                {
                    Filter = "Excel | *.xlsx"
                };
                toolStripStatusLabel1.Text = "Excelfil sparad som " + sfd.FileName;

                sfd.ShowDialog();


                using (var fileData = new FileStream(sfd.FileName, FileMode.Create))
                {
                    wb.SaveAs(fileData);
                }
                sfd.Dispose();
            }
            catch (Exception ex)
            {

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
            string tokenString;
            toolStripProgressBar1.Value = 0;
            objekten.Add(new Objekt("1", "Saknas", "Saknas"));
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
                                    fNamn = konto[3].ToString().Replace('"', ' ').Trim();
                                    label2.Text = "Företag: " + fNamn;
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
                                    string obj_id = rensaDim[2].Replace('"', ' ').Trim().ToUpper();
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
                                    bool addedAtLeastOne = false;
                                    if (brackets.Length > 0)
                                    {
                                        dim = Regex.Split(brackets, regexDelning);

                                        for (int i = 1; i < dim.Length - 3; i += 4)
                                        {
                                            addedAtLeastOne = true;
                                            objekt.Add(dim[i].Replace('"', ' ').Trim(), dim[i + 2].Replace('"', ' ').Trim().ToUpper());

                                        }
                                        if (!addedAtLeastOne)
                                        {
                                            if (!objekt.ContainsKey("Saknas"))
                                            {
                                                objekt.Add("1", "Saknas");


                                            }
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
                                    if (!monthList.Contains(transaktionsdatum.Substring(0, 4) + "-" + transaktionsdatum.Substring(4, 2)))
                                        monthList.Add(transaktionsdatum.Substring(0, 4) + "-" + transaktionsdatum.Substring(4, 2));
                                    String kontonr = rensaDelar[1].Replace('"', ' ').Trim();
                                    string b = rensaDelar[2].Replace('.', ',').Replace('"', ' ').Trim();
                                    double belopp = -double.Parse(b);

                                    string transtext = "";
                                    if (rensaDelar.Count > 4)
                                        transtext = rensaDelar[4];
                                    if (transtext.Equals(""))
                                        transtext = vertext;
                                    double kvantitet = 0.0;
                                    string sign = "";
                                    DateTime transaktionsdate = DateTime.Parse(transaktionsdatum.Substring(0, 4) + "-" + transaktionsdatum.Substring(4, 2) + "-" + transaktionsdatum.Substring(6, 2));
                                    transaktioner.Add(new Transaktion(objekt, transaktionsdate, kontonr, belopp / 100, transtext, kvantitet, sign));
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


                    Update();
                    toolStripProgressBar1.Increment(1);
                    toolStripStatusLabel1.Text = "Bearbetar rad " + toolStripProgressBar1.Value + " av " + toolStripProgressBar1.Maximum;
                } //foreach

                monthList.Sort();
                foreach (string s in monthList)
                {
                    comboBox1.Items.Add(s);

                }

                foreach (string k in dimensioner.Keys)
                {
                    comboBox2.Items.Add(dimensioner[k]);

                }
                comboBox2.SelectedIndex = comboBox2.Items.IndexOf("Resultatenhet");
                comboBox1.SelectedItem = comboBox1.Items[comboBox1.Items.Count - 1];
                button4.Enabled = true;
                SelectResultatenhet(comboBox2.SelectedItem.ToString());

            }//try
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace);
            }//Catch

        } //button3_click

        private void SelectResultatenhet(string typ)
        {
            checkedListBox1.Items.Clear();
            string myKey = dimensioner.FirstOrDefault(x => x.Value.ToString().Equals(typ)).Key;
            foreach (Objekt o in objekten)
            {
                if (o.Typ.Equals(myKey))
                    checkedListBox1.Items.Add(o.Id);

            }


            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, true);
            }
            selectedType = myKey;

        }

        private void YearlyTotal(ExcelPackage wb, Objekt o)
        {
            try
            {
                char[] alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();
                string resenh;
                if (o == null)
                    resenh = "*";
                else
                    resenh = o.Id;
                Console.WriteLine(resenh);
                DateTime startMonth = DateTime.Parse("2018-05-01");
                DateTime endMonth = endOfMonth(DateTime.Parse("2018-05-31"));
                DateTime startYear = DateTime.Parse("2018-05-01");
                DateTime endYear = DateTime.Parse("2019-04-30");
                DateTime startPreviousYear = DateTime.Parse("2017-05-01");

                DateTime endPreviousYear = DateTime.Parse("2018-04-30");



                int intaktsRad;
                int rorelseKostnadsRad;
                int bruttoRad;
                List<int> kostnadsRader = new List<int>();

                int kostnadsRad;


                string newname = resenh;
                ExcelWorksheet totalSheet;

                if (resenh.Contains("*"))
                    totalSheet = wb.Workbook.Worksheets.Add("TOTAL");
                else
                {
                    foreach (ExcelWorksheet ws in wb.Workbook.Worksheets)
                    {
                        if (ws.Name.Equals(newname))
                            newname += "_";
                    }

                    totalSheet = wb.Workbook.Worksheets.Add(newname);
                }

                totalSheet.Cells["B1"].Value = fNamn;
                totalSheet.Cells["C1"].Value = resenh;
                if (o != null)
                    totalSheet.Cells["C1"].Value = o.Namn;


                totalSheet.Cells["A2"].Value = "Från:";
                totalSheet.Cells["B2"].Value = startMonth;
                totalSheet.Cells["B2"].Style.Numberformat.Format = "yyyy-mm-dd";
                totalSheet.Cells["A3"].Value = "Till:";
                totalSheet.Cells["B3"].Value = endYear;
                totalSheet.Cells["B3"].Style.Numberformat.Format = "yyyy-mm-dd";

                //Skapa tabellerna

                Dictionary<string, double>[] months = new Dictionary<string, double>[12];
                for (int m = 0; m < 12; m++)
                {
                    // Console.WriteLine(m + ":" + startMonth.AddMonths(m).ToShortDateString() + " -- " + endMonth.AddMonths(m).ToShortDateString());
                    months[m] = SumTransaktion(new DateTime(startMonth.AddMonths(m).Ticks), new DateTime(endMonth.AddMonths(m).Ticks), resenh);
                }

                Dictionary<string, double> sumYTD = SumTransaktion(startYear, endYear, resenh);
                Dictionary<string, double> sumLastYTD = SumTransaktion(startPreviousYear, endPreviousYear, resenh);


                int row = 6;
                totalSheet.Cells["A5"].Value = "Konto";
                totalSheet.Cells["B5"].Value = "Benämning";
                for (int m = 0; m < 12; m++)
                    totalSheet.Cells[alpha[m + 3].ToString() + "5"].Value = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(startMonth.AddMonths(m).Month);
                totalSheet.Cells["O5"].Value = "Ack resultat";
                totalSheet.Cells["P5"].Value = "Ack resultat fg år";
                totalSheet.Cells["Q5"].Value = "Differens";

                row++;
                totalSheet.Cells["B" + row].Value = "Intäkter";
                row++;
                //3000-3999
                int startrow = row;
                int[] sorteradeKonton = new int[konton.Count];
                int i = 0;
                foreach (string s in konton.Keys)
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
                        totalSheet.Cells["A" + row].Value = konto;
                        totalSheet.Cells["B" + row].Value = konton[s.ToString()];
                        for (int m = 0; m < 12; m++)
                        {
                            if (months[m].ContainsKey(konto.ToString()))
                                totalSheet.Cells[alpha[m + 3].ToString() + row].Value = months[m][konto.ToString()];
                            else
                                totalSheet.Cells[alpha[m + 3].ToString() + row].Value = 0.0;


                        }

                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["O" + row].Value = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells["O" + row].Value = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["P" + row].Value = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["P" + row].Value = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["Q" + row].Value = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["Q" + row].Value = 0.0;
                        row++;
                    }
                }

                for (int r = startrow; r <= (row - 1); r++)
                {
                    totalSheet.Row(r).OutlineLevel = 1;
                    totalSheet.Row(r).Collapsed = true;
                }
                if (startrow == row)
                    row++;
                totalSheet.Cells["B" + row].Value = "Summa intäkter";
                for (int c = 3; c < 18; c++)
                {
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "SUM(" + alpha[c - 1] + startrow + ":" + alpha[c - 1] + (row - 1) + ")";

                }

                intaktsRad = row;



                row++;
                row++;

                //4000-4999 Rörelsens kostnader
                totalSheet.Cells["B" + row].Value = "Rörelsens kostnader";
                row++;
                startrow = row;
                foreach (string s in konton.Keys)
                {

                    int konto = Int32.Parse(s);
                    if (konto >= 4000 && konto <= 4999)
                    {
                        totalSheet.Cells["A" + row].Value = konto;
                        totalSheet.Cells["B" + row].Value = konton[s.ToString()];
                        for (int m = 0; m < 12; m++)
                        {
                            if (months[m].ContainsKey(konto.ToString()))
                                totalSheet.Cells[alpha[m + 3] + "" + row].Value = months[m][konto.ToString()];
                            else
                                totalSheet.Cells["C" + row].Value = 0.0;


                        }

                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["O" + row].Value = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells["O" + row].Value = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["P" + row].Value = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["P" + row].Value = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["Q" + row].Value = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["Q" + row].Value = 0.0;
                        row++;
                    }
                }
                for (int r = startrow; r <= (row - 1); r++)
                {
                    totalSheet.Row(r).OutlineLevel = 1;
                    totalSheet.Row(r).Collapsed = true;
                }
                if (startrow == row)
                    row++;
                totalSheet.Cells["B" + row].Value = "Summa rörelsens kostnader";
                for (int c = 3; c < 18; c++)
                {
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "SUM(" + alpha[c - 1] + startrow + ":" + alpha[c - 1] + (row - 1) + ")";

                }
                rorelseKostnadsRad = row;
                row++;
                row++;
                bruttoRad = row;
                totalSheet.Cells["B" + row].Value = "Bruttovinst";
                for (int c = 3; c < 18; c++)
                {
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "SUM(" + alpha[c - 1] + intaktsRad + "+" + alpha[c - 1] + rorelseKostnadsRad + ")";

                }

                row++;
                row++;


                //5000-6999 Externa kostnader
                totalSheet.Cells["B" + row].Value = "Externa kostnader";
                row++;
                startrow = row;
                foreach (string s in konton.Keys)
                {

                    int konto = Int32.Parse(s);
                    if (konto >= 5000 && konto <= 6999)
                    {
                        totalSheet.Cells["A" + row].Value = konto;
                        totalSheet.Cells["B" + row].Value = konton[s.ToString()];
                        for (int m = 0; m < 12; m++)
                        {
                            if (months[m].ContainsKey(konto.ToString()))
                                totalSheet.Cells[alpha[m + 3] + "" + row].Value = months[m][konto.ToString()];
                            else
                                totalSheet.Cells["C" + row].Value = 0.0;


                        }

                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["O" + row].Value = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells["O" + row].Value = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["P" + row].Value = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["P" + row].Value = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["Q" + row].Value = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["Q" + row].Value = 0.0;
                        row++;
                    }
                }
                for (int r = startrow; r <= (row - 1); r++)
                {
                    totalSheet.Row(r).OutlineLevel = 1;
                    totalSheet.Row(r).Collapsed = true;
                }
                if (startrow == row)
                    row++;
                kostnadsRader.Add(row);
                totalSheet.Cells["B" + row].Value = "Summa externa kostnader";
                for (int c = 3; c < 18; c++)
                {
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "SUM(" + alpha[c - 1] + startrow + ":" + alpha[c - 1] + (row - 1) + ")";

                }
                row++;

                row++;


                //7000-7799 Personalkostnader
                totalSheet.Cells["B" + row].Value = "Personalkostnader";
                row++;
                startrow = row;
                foreach (string s in konton.Keys)
                {

                    int konto = Int32.Parse(s);
                    if (konto >= 7000 && konto <= 7799)
                    {
                        totalSheet.Cells["A" + row].Value = konto;
                        totalSheet.Cells["B" + row].Value = konton[s.ToString()];
                        for (int m = 0; m < 12; m++)
                        {
                            if (months[m].ContainsKey(konto.ToString()))
                                totalSheet.Cells[alpha[m + 3] + "" + row].Value = months[m][konto.ToString()];
                            else
                                totalSheet.Cells["C" + row].Value = 0.0;


                        }

                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["O" + row].Value = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells["O" + row].Value = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["P" + row].Value = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["P" + row].Value = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["Q" + row].Value = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["Q" + row].Value = 0.0;
                        row++;
                    }
                }
                for (int r = startrow; r <= (row - 1); r++)
                {
                    totalSheet.Row(r).OutlineLevel = 1;
                    totalSheet.Row(r).Collapsed = true;
                }
                if (startrow == row)
                    row++;
                kostnadsRader.Add(row);
                totalSheet.Cells["B" + row].Value = "Summa personalkostnader";
                for (int c = 3; c < 18; c++)
                {
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "SUM(" + alpha[c - 1] + startrow + ":" + alpha[c - 1] + (row - 1) + ")";

                }
                row++;

                row++;

                //7800-7999 Avskrivningar
                totalSheet.Cells["B" + row].Value = "Avskrivningar";
                row++;
                startrow = row;
                foreach (string s in konton.Keys)
                {

                    int konto = Int32.Parse(s);
                    if (konto >= 7800 && konto <= 7999)
                    {
                        totalSheet.Cells["A" + row].Value = konto;
                        totalSheet.Cells["B" + row].Value = konton[s.ToString()];
                        for (int m = 0; m < 12; m++)
                        {
                            if (months[m].ContainsKey(konto.ToString()))
                                totalSheet.Cells[alpha[m + 3] + "" + row].Value = months[m][konto.ToString()];
                            else
                                totalSheet.Cells["C" + row].Value = 0.0;


                        }

                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["O" + row].Value = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells["O" + row].Value = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["P" + row].Value = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["P" + row].Value = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["Q" + row].Value = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["Q" + row].Value = 0.0;
                        row++;
                    }
                }
                for (int r = startrow; r <= (row - 1); r++)
                {
                    totalSheet.Row(r).OutlineLevel = 1;
                    totalSheet.Row(r).Collapsed = true;
                }
                if (startrow == row)
                    row++;
                kostnadsRader.Add(row);
                totalSheet.Cells["B" + row].Value = "Summa avskrivningar";
                for (int c = 3; c < 18; c++)
                {
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "SUM(" + alpha[c - 1] + startrow + ":" + alpha[c - 1] + (row - 1) + ")";

                }
                row++;

                row++;

                //8000-8799 Finansiella kostnader
                totalSheet.Cells["B" + row].Value = "Finansiella kostnader";
                row++;
                startrow = row;
                foreach (string s in konton.Keys)
                {

                    int konto = Int32.Parse(s);
                    if (konto >= 8000 && konto <= 8799)
                    {
                        totalSheet.Cells["A" + row].Value = konto;
                        totalSheet.Cells["B" + row].Value = konton[s.ToString()];
                        for (int m = 0; m < 12; m++)
                        {
                            if (months[m].ContainsKey(konto.ToString()))
                                totalSheet.Cells[alpha[m + 3] + "" + row].Value = months[m][konto.ToString()];
                            else
                                totalSheet.Cells["C" + row].Value = 0.0;


                        }

                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["O" + row].Value = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells["O" + row].Value = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["P" + row].Value = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["P" + row].Value = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["Q" + row].Value = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["Q" + row].Value = 0.0;
                        row++;
                    }
                }
                for (int r = startrow; r <= (row - 1); r++)
                {
                    totalSheet.Row(r).OutlineLevel = 1;
                    totalSheet.Row(r).Collapsed = true;
                }
                if (startrow == row)
                    row++;
                kostnadsRader.Add(row);
                totalSheet.Cells["B" + row].Value = "Summa finansiella kostnader";
                for (int c = 3; c < 18; c++)
                {
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "SUM(" + alpha[c - 1] + startrow + ":" + alpha[c - 1] + (row - 1) + ")";

                }
                row++;

                row++;

                //8800-8999 Bokslutsdispositioner
                totalSheet.Cells["B" + row].Value = "Bokslutsdispositioner";
                row++;
                startrow = row;
                foreach (string s in konton.Keys)
                {

                    int konto = Int32.Parse(s);
                    if (konto >= 8800 && konto <= 8999)
                    {
                        totalSheet.Cells["A" + row].Value = konto;
                        totalSheet.Cells["B" + row].Value = konton[s.ToString()];
                        for (int m = 0; m < 12; m++)
                        {
                            if (months[m].ContainsKey(konto.ToString()))
                                totalSheet.Cells[alpha[m + 3] + "" + row].Value = months[m][konto.ToString()];
                            else
                                totalSheet.Cells["C" + row].Value = 0.0;


                        }

                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["O" + row].Value = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells["O" + row].Value = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["P" + row].Value = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["P" + row].Value = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["Q" + row].Value = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["Q" + row].Value = 0.0;
                        row++;
                    }
                }
                for (int r = startrow; r <= (row - 1); r++)
                {
                    totalSheet.Row(r).OutlineLevel = 1;
                    totalSheet.Row(r).Collapsed = true;
                }
                if (startrow == row)
                    row++;
                kostnadsRader.Add(row);
                totalSheet.Cells["B" + row].Value = "Summa bokslutsdispositioner";
                for (int c = 3; c < 18; c++)
                {
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "SUM(" + alpha[c - 1] + startrow + ":" + alpha[c - 1] + (row - 1) + ")";

                }
                row++;
                row++;

                kostnadsRad = row;
                totalSheet.Cells["B" + row].Value = "Summa kostnader";
                for (int c = 3; c < 18; c++)
                {
                    string summaKostnader = "SUM(" + alpha[c - 1];
                    foreach (int n in kostnadsRader)
                    {
                        summaKostnader += n + "+" + alpha[c - 1];
                    }
                    summaKostnader = summaKostnader.Substring(0, summaKostnader.Length - 2) + ")";
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = summaKostnader;

                }


                row++;
                row++;
                totalSheet.Cells["B" + row].Value = "Resultat";
                for (int c = 3; c < 18; c++)
                {
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "SUM(" + alpha[c - 1] + bruttoRad + "+" + alpha[c - 1] + kostnadsRad + ")";

                }
                for (int r = 3; r <= 14; r++)
                {
                    totalSheet.Column(r).OutlineLevel = 1;
                    totalSheet.Column(r).Collapsed = true;
                }
                totalSheet.Cells["C7:Z" + row].Style.Numberformat.Format = "#,### ;[Red]-#,### ";
                totalSheet.Cells["A1:Z" + row].AutoFitColumns();
                totalSheet.View.FreezePanes(6, 2);






            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace);
            }
        }

        private void Total(ExcelPackage wb, Objekt o)
        {
            try
            {
                char[] alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();
                string resenh;
                if (o == null)
                    resenh = "*";
                else
                    resenh = o.Id;
                Console.WriteLine(resenh);
                DateTime startMonth = DateTime.Parse(comboBox1.SelectedItem.ToString() + "-01");
                DateTime endMonth = endOfMonth(DateTime.Parse(comboBox1.SelectedItem.ToString() + "-01"));
                DateTime startYear;
                DateTime endYear;
                DateTime startPreviousYear;

                if (startMonth.Month < 5)
                {
                    startYear = new DateTime(startMonth.Year - 1, 5, 1);
                    endYear = new DateTime(startMonth.Year, 4, 30);
                }
                else
                {
                    startYear = new DateTime(startMonth.Year, 5, 1);
                    endYear = new DateTime(startMonth.Year + 1, 4, 30);
                }
                if (startMonth.Month < 5)
                    startPreviousYear = new DateTime(startMonth.Year - 2, 5, 1);
                else
                    startPreviousYear = new DateTime(startMonth.Year - 1, 5, 1);

                if (checkBox1.Checked)
                {
                    startMonth = new DateTime(startYear.Year, startYear.Month, 1);
                }
                DateTime endPreviousYear = new DateTime(startMonth.Year - 1, startMonth.AddYears(-1).Month, endOfMonth(startMonth.AddYears(-1)).Day);
                DateTime startPrevMonth = startMonth.AddMonths(-1);
                DateTime endPrevMonth = endOfMonth(startMonth.AddMonths(-1));
                if (checkBox1.Checked)
                {
                    startPrevMonth = startPreviousYear;
                    endPrevMonth = endPreviousYear;
                }



                int intaktsRad;
                int rorelseKostnadsRad;
                int bruttoRad;
                List<int> kostnadsRader = new List<int>();

                int kostnadsRad;


                string newname = resenh;
                ExcelWorksheet totalSheet;

                if (resenh.Contains("*"))
                    totalSheet = wb.Workbook.Worksheets.Add("TOTAL");
                else
                {
                    foreach (ExcelWorksheet ws in wb.Workbook.Worksheets)
                    {
                        if (ws.Name.Equals(newname))
                            newname += "_";
                    }

                    totalSheet = wb.Workbook.Worksheets.Add(newname);
                }

                totalSheet.Cells["B1"].Value = fNamn;
                totalSheet.Cells["C1"].Value = resenh;
                if (o != null)
                    totalSheet.Cells["C1"].Value = o.Namn;


                totalSheet.Cells["A2"].Value = "Från:";
                totalSheet.Cells["B2"].Value = startMonth;
                totalSheet.Cells["A3"].Value = "Till:";
                totalSheet.Cells["B3"].Value = endYear;

                //Skapa tabellerna

                Dictionary<string, double>[] months = new Dictionary<string, double>[12];
                for (int m = 0; m < 12; m++)
                {
                    // Console.WriteLine(m + ":" + startMonth.AddMonths(m).ToShortDateString() + " -- " + endMonth.AddMonths(m).ToShortDateString());
                    months[m] = SumTransaktion(new DateTime(startMonth.AddMonths(m).Ticks), new DateTime(endMonth.AddMonths(m).Ticks), resenh);
                }

                Dictionary<string, double> sumYTD = SumTransaktion(startYear, endYear, resenh);
                Dictionary<string, double> sumLastYTD = SumTransaktion(startPreviousYear, endPreviousYear, resenh);


                int row = 6;
                totalSheet.Cells["A5"].Value = "Konto";
                totalSheet.Cells["B5"].Value = "Benämning";
                for (int m = 0; m < 12; m++)
                    totalSheet.Cells[alpha[m + 3].ToString() + "5"].Value = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(startMonth.AddMonths(m).Month);
                totalSheet.Cells["O5"].Value = "Ack resultat";
                totalSheet.Cells["P5"].Value = "Ack resultat fg år";
                totalSheet.Cells["Q5"].Value = "Differens";

                row++;
                totalSheet.Cells["B" + row].Value = "Intäkter";
                row++;
                //3000-3999
                int startrow = row;
                int[] sorteradeKonton = new int[konton.Count];
                int i = 0;
                foreach (string s in konton.Keys)
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
                        totalSheet.Cells["A" + row].Value = konto;
                        totalSheet.Cells["B" + row].Value = konton[s.ToString()];
                        for (int m = 0; m < 12; m++)
                        {
                            if (months[m].ContainsKey(konto.ToString()))
                                totalSheet.Cells[alpha[m + 3].ToString() + row].Value = months[m][konto.ToString()];
                            else
                                totalSheet.Cells[alpha[m + 3].ToString() + row].Value = 0.0;


                        }

                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["O" + row].Value = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells["O" + row].Value = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["P" + row].Value = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["P" + row].Value = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["Q" + row].Value = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["Q" + row].Value = 0.0;
                        row++;
                    }
                }

                for (int r = startrow; r <= (row - 1); r++)
                {
                    totalSheet.Row(r).OutlineLevel = 1;
                    totalSheet.Row(r).Collapsed = true;
                }
                if (startrow == row)
                    row++;
                totalSheet.Cells["B" + row].Value = "Summa intäkter";
                for (int c = 3; c < 18; c++)
                {
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "SUM(" + alpha[c - 1] + startrow + ":" + alpha[c - 1] + (row - 1) + ")";

                }

                intaktsRad = row;



                row++;
                row++;

                //4000-4999 Rörelsens kostnader
                totalSheet.Cells["B" + row].Value = "Rörelsens kostnader";
                row++;
                startrow = row;
                foreach (string s in konton.Keys)
                {

                    int konto = Int32.Parse(s);
                    if (konto >= 4000 && konto <= 4999)
                    {
                        totalSheet.Cells["A" + row].Value = konto;
                        totalSheet.Cells["B" + row].Value = konton[s.ToString()];
                        for (int m = 0; m < 12; m++)
                        {
                            if (months[m].ContainsKey(konto.ToString()))
                                totalSheet.Cells[alpha[m + 3] + "" + row].Value = months[m][konto.ToString()];
                            else
                                totalSheet.Cells["C" + row].Value = 0.0;


                        }

                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["O" + row].Value = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells["O" + row].Value = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["P" + row].Value = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["P" + row].Value = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["Q" + row].Value = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["Q" + row].Value = 0.0;
                        row++;
                    }
                }
                for (int r = startrow; r <= (row - 1); r++)
                {
                    totalSheet.Row(r).OutlineLevel = 1;
                    totalSheet.Row(r).Collapsed = true;
                }
                if (startrow == row)
                    row++;
                totalSheet.Cells["B" + row].Value = "Summa rörelsens kostnader";
                for (int c = 3; c < 18; c++)
                {
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "SUM(" + alpha[c - 1] + startrow + ":" + alpha[c - 1] + (row - 1) + ")";

                }
                rorelseKostnadsRad = row;
                row++;
                row++;
                bruttoRad = row;
                totalSheet.Cells["B" + row].Value = "Bruttovinst";
                for (int c = 3; c < 18; c++)
                {
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "SUM(" + alpha[c - 1] + intaktsRad + "+" + alpha[c - 1] + rorelseKostnadsRad + ")";

                }

                row++;
                row++;


                //5000-6999 Externa kostnader
                totalSheet.Cells["B" + row].Value = "Externa kostnader";
                row++;
                startrow = row;
                foreach (string s in konton.Keys)
                {

                    int konto = Int32.Parse(s);
                    if (konto >= 5000 && konto <= 6999)
                    {
                        totalSheet.Cells["A" + row].Value = konto;
                        totalSheet.Cells["B" + row].Value = konton[s.ToString()];
                        for (int m = 0; m < 12; m++)
                        {
                            if (months[m].ContainsKey(konto.ToString()))
                                totalSheet.Cells[alpha[m + 3] + "" + row].Value = months[m][konto.ToString()];
                            else
                                totalSheet.Cells["C" + row].Value = 0.0;


                        }

                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["O" + row].Value = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells["O" + row].Value = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["P" + row].Value = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["P" + row].Value = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["Q" + row].Value = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["Q" + row].Value = 0.0;
                        row++;
                    }
                }
                for (int r = startrow; r <= (row - 1); r++)
                {
                    totalSheet.Row(r).OutlineLevel = 1;
                    totalSheet.Row(r).Collapsed = true;
                }
                if (startrow == row)
                    row++;
                kostnadsRader.Add(row);
                totalSheet.Cells["B" + row].Value = "Summa externa kostnader";
                for (int c = 3; c < 18; c++)
                {
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "SUM(" + alpha[c - 1] + startrow + ":" + alpha[c - 1] + (row - 1) + ")";

                }
                row++;

                row++;


                //7000-7799 Personalkostnader
                totalSheet.Cells["B" + row].Value = "Personalkostnader";
                row++;
                startrow = row;
                foreach (string s in konton.Keys)
                {

                    int konto = Int32.Parse(s);
                    if (konto >= 7000 && konto <= 7799)
                    {
                        totalSheet.Cells["A" + row].Value = konto;
                        totalSheet.Cells["B" + row].Value = konton[s.ToString()];
                        for (int m = 0; m < 12; m++)
                        {
                            if (months[m].ContainsKey(konto.ToString()))
                                totalSheet.Cells[alpha[m + 3] + "" + row].Value = months[m][konto.ToString()];
                            else
                                totalSheet.Cells["C" + row].Value = 0.0;


                        }

                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["O" + row].Value = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells["O" + row].Value = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["P" + row].Value = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["P" + row].Value = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["Q" + row].Value = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["Q" + row].Value = 0.0;
                        row++;
                    }
                }
                for (int r = startrow; r <= (row - 1); r++)
                {
                    totalSheet.Row(r).OutlineLevel = 1;
                    totalSheet.Row(r).Collapsed = true;
                }
                if (startrow == row)
                    row++;
                kostnadsRader.Add(row);
                totalSheet.Cells["B" + row].Value = "Summa personalkostnader";
                for (int c = 3; c < 18; c++)
                {
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "SUM(" + alpha[c - 1] + startrow + ":" + alpha[c - 1] + (row - 1) + ")";

                }
                row++;

                row++;

                //7800-7999 Avskrivningar
                totalSheet.Cells["B" + row].Value = "Avskrivningar";
                row++;
                startrow = row;
                foreach (string s in konton.Keys)
                {

                    int konto = Int32.Parse(s);
                    if (konto >= 7800 && konto <= 7999)
                    {
                        totalSheet.Cells["A" + row].Value = konto;
                        totalSheet.Cells["B" + row].Value = konton[s.ToString()];
                        for (int m = 0; m < 12; m++)
                        {
                            if (months[m].ContainsKey(konto.ToString()))
                                totalSheet.Cells[alpha[m + 3] + "" + row].Value = months[m][konto.ToString()];
                            else
                                totalSheet.Cells["C" + row].Value = 0.0;


                        }

                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["O" + row].Value = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells["O" + row].Value = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["P" + row].Value = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["P" + row].Value = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["Q" + row].Value = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["Q" + row].Value = 0.0;
                        row++;
                    }
                }
                for (int r = startrow; r <= (row - 1); r++)
                {
                    totalSheet.Row(r).OutlineLevel = 1;
                    totalSheet.Row(r).Collapsed = true;
                }
                if (startrow == row)
                    row++;
                kostnadsRader.Add(row);
                totalSheet.Cells["B" + row].Value = "Summa avskrivningar";
                for (int c = 3; c < 18; c++)
                {
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "SUM(" + alpha[c - 1] + startrow + ":" + alpha[c - 1] + (row - 1) + ")";

                }
                row++;

                row++;

                //8000-8799 Finansiella kostnader
                totalSheet.Cells["B" + row].Value = "Finansiella kostnader";
                row++;
                startrow = row;
                foreach (string s in konton.Keys)
                {

                    int konto = Int32.Parse(s);
                    if (konto >= 8000 && konto <= 8799)
                    {
                        totalSheet.Cells["A" + row].Value = konto;
                        totalSheet.Cells["B" + row].Value = konton[s.ToString()];
                        for (int m = 0; m < 12; m++)
                        {
                            if (months[m].ContainsKey(konto.ToString()))
                                totalSheet.Cells[alpha[m + 3] + "" + row].Value = months[m][konto.ToString()];
                            else
                                totalSheet.Cells["C" + row].Value = 0.0;


                        }

                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["O" + row].Value = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells["O" + row].Value = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["P" + row].Value = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["P" + row].Value = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["Q" + row].Value = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["Q" + row].Value = 0.0;
                        row++;
                    }
                }
                for (int r = startrow; r <= (row - 1); r++)
                {
                    totalSheet.Row(r).OutlineLevel = 1;
                    totalSheet.Row(r).Collapsed = true;
                }
                if (startrow == row)
                    row++;
                kostnadsRader.Add(row);
                totalSheet.Cells["B" + row].Value = "Summa finansiella kostnader";
                for (int c = 3; c < 18; c++)
                {
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "SUM(" + alpha[c - 1] + startrow + ":" + alpha[c - 1] + (row - 1) + ")";

                }
                row++;

                row++;

                //8800-8999 Bokslutsdispositioner
                totalSheet.Cells["B" + row].Value = "Bokslutsdispositioner";
                row++;
                startrow = row;
                foreach (string s in konton.Keys)
                {

                    int konto = Int32.Parse(s);
                    if (konto >= 8800 && konto <= 8999)
                    {
                        totalSheet.Cells["A" + row].Value = konto;
                        totalSheet.Cells["B" + row].Value = konton[s.ToString()];
                        for (int m = 0; m < 12; m++)
                        {
                            if (months[m].ContainsKey(konto.ToString()))
                                totalSheet.Cells[alpha[m + 3] + "" + row].Value = months[m][konto.ToString()];
                            else
                                totalSheet.Cells["C" + row].Value = 0.0;


                        }

                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["O" + row].Value = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells["O" + row].Value = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["P" + row].Value = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["P" + row].Value = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["Q" + row].Value = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["Q" + row].Value = 0.0;
                        row++;
                    }
                }
                for (int r = startrow; r <= (row - 1); r++)
                {
                    totalSheet.Row(r).OutlineLevel = 1;
                    totalSheet.Row(r).Collapsed = true;
                }
                if (startrow == row)
                    row++;
                kostnadsRader.Add(row);
                totalSheet.Cells["B" + row].Value = "Summa bokslutsdispositioner";
                for (int c = 3; c < 18; c++)
                {
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "SUM(" + alpha[c - 1] + startrow + ":" + alpha[c - 1] + (row - 1) + ")";

                }
                row++;
                row++;

                kostnadsRad = row;
                totalSheet.Cells["B" + row].Value = "Summa kostnader";
                for (int c = 3; c < 18; c++)
                {
                    string summaKostnader = "SUM(" + alpha[c - 1];
                    foreach (int n in kostnadsRader)
                    {
                        summaKostnader += n + "+" + alpha[c - 1];
                    }
                    summaKostnader = summaKostnader.Substring(0, summaKostnader.Length - 2) + ")";
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = summaKostnader;

                }


                row++;
                row++;
                totalSheet.Cells["B" + row].Value = "Resultat";
                for (int c = 3; c < 18; c++)
                {
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "SUM(" + alpha[c - 1] + bruttoRad + "+" + alpha[c - 1] + kostnadsRad + ")";

                }
                for (int r = 3; r <= 14; r++)
                {
                    totalSheet.Column(r).OutlineLevel = 1;
                    totalSheet.Column(r).Collapsed = true;
                }
                totalSheet.Cells["C7:Z" + row].Style.Numberformat.Format = "#,### ;[Red]-#,### ";
                totalSheet.Cells["A1:Z" + row].AutoFitColumns();
                totalSheet.View.FreezePanes(6, 2);






            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace);
            }
        }

        private static DateTime endOfMonth(DateTime datum)
        {
            int[] endDay;
            if (DateTime.IsLeapYear(datum.Year))
            {
                endDay = new int[] { 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
            }
            else
            {
                endDay = new int[] { 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
            }
            if (DateTime.TryParse(datum.Year + "-" + datum.Month + "-" + endDay[datum.Month - 1], out DateTime returnDate))
                return returnDate;
            else
                return datum;
        }

        private void SkapaResultatsida(ExcelPackage wb)
        {
            //För varje flik skapa en rad med 
            //Kolumner Summa Intäkt, summa rörelsekostnader, bruttovinst, summa externa kostnader, Summa personalkostnader, summa finansiella kostnader, summa kostnader, Resultat
            ExcelWorksheet overSheet = wb.Workbook.Worksheets.Add("Översikt");


            overSheet.Cells["C1"].Value = "Summa Intäkter";
            overSheet.Cells["D1"].Value = "Summa rörelsens kostnader";
            overSheet.Cells["E1"].Value = "Bruttovinst";
            overSheet.Cells["F1"].Value = "Summa externa kostnader";
            overSheet.Cells["G1"].Value = "Summa personalkostnader";
            overSheet.Cells["H1"].Value = "Summa avskrivningar";
            overSheet.Cells["I1"].Value = "Summa finansiella kostnader";
            overSheet.Cells["J1"].Value = "Summa bokslutsdispositioner";
            overSheet.Cells["K1"].Value = "Summa kostnader";
            overSheet.Cells["L1"].Value = "Resultat";

            int i = 2;
            foreach (string s in valda)
            {
                overSheet.Cells["A" + i].Value = s;

                overSheet.Cells["C" + i].Formula = "VLOOKUP(C1,'" + s + "'!B:O,14,0)";
                overSheet.Cells["D" + i].Formula = "VLOOKUP(D1,'" + s + "'!B:O,14,0)";
                overSheet.Cells["E" + i].Formula = "VLOOKUP(E1,'" + s + "'!B:O,14,0)";
                overSheet.Cells["F" + i].Formula = "VLOOKUP(F1,'" + s + "'!B:O,14,0)";
                overSheet.Cells["G" + i].Formula = "VLOOKUP(G1,'" + s + "'!B:O,14,0)";
                overSheet.Cells["H" + i].Formula = "VLOOKUP(H1,'" + s + "'!B:O,14,0)";
                overSheet.Cells["I" + i].Formula = "VLOOKUP(I1,'" + s + "'!B:O,14,0)";
                overSheet.Cells["J" + i].Formula = "VLOOKUP(J1,'" + s + "'!B:O,14,0)";
                overSheet.Cells["K" + i].Formula = "VLOOKUP(K1,'" + s + "'!B:O,14,0)";
                overSheet.Cells["L" + i].Formula = "VLOOKUP(L1,'" + s + "'!B:O,14,0)";
                i++;
            }
            overSheet.Cells["A" + i].Value = "TOTAL";

            overSheet.Cells[i, 3].Formula = "VLOOKUP(C1," + "TOTAL" + "!B:O,14,0)";
            overSheet.Cells[i, 4].Formula = "VLOOKUP(D1," + "TOTAL" + "!B:O,14,0)";
            overSheet.Cells[i, 5].Formula = "VLOOKUP(E1," + "TOTAL" + "!B:O,14,0)";
            overSheet.Cells[i, 6].Formula = "VLOOKUP(F1," + "TOTAL" + "!B:O,14,0)";
            overSheet.Cells[i, 7].Formula = "VLOOKUP(G1," + "TOTAL" + "!B:O,14,0)";
            overSheet.Cells[i, 8].Formula = "VLOOKUP(H1," + "TOTAL" + "!B:O,14,0)";
            overSheet.Cells[i, 9].Formula = "VLOOKUP(I1," + "TOTAL" + "!B:O,14,0)";
            overSheet.Cells[i, 10].Formula = "VLOOKUP(J1," + "TOTAL" + "!B:O,14,0)";
            overSheet.Cells[i, 11].Formula = "VLOOKUP(K1," + "TOTAL" + "!B:O,14,0)";
            overSheet.Cells[i, 12].Formula = "VLOOKUP(L1," + "TOTAL" + "!B:O,14,0)";

            overSheet.Cells["C2:L" + i].Style.Numberformat.Format = "#,### ;[Red]-#,### ";
            overSheet.Cells["A1:L" + i].AutoFitColumns();

        }

        private Dictionary<string, double> SumTransaktion(DateTime from, DateTime to, string resenh)
        {
            Dictionary<string, double> sum = new Dictionary<string, double>();

            foreach (Transaktion t in transaktioner)
            {


                if (!sum.ContainsKey(t.Kontonr))
                    sum.Add(t.Kontonr, 0.0); //Add to dictionary if not already in it
                if (t.Objekt.ContainsKey(selectedType))
                    if (t.Objekt[selectedType].Equals(resenh) || resenh.Equals("*"))
                        if (t.Transaktionsdatum >= from & t.Transaktionsdatum <= to)
                            sum[t.Kontonr] += t.Belopp; // Add t.Belopp to Sum if resenh == t.Id


            }


            return sum;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            valda.Clear();
            foreach (object o in checkedListBox1.CheckedItems)
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

        private void SaveTransactions()
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                string fileContents = "";

                OpenFileDialog openFileDialog1 = new OpenFileDialog
                {
                    Filter = "SIE 4|*.SE",
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
                    List<string> nuvarandelista = new List<string>(Regex.Split(fileContents, Environment.NewLine));
                    foreach (string s in nuvarandelista)
                        lista.Add(s);

                }
                else
                {
                    throw new Exception("Fel vid filinläsning");
                }
                // MessageBox.Show("Rader: " + list.Count);
                button3.Enabled = true;
                Update();
                openFileDialog1.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Fel: " + ex.Message + " \n" + ex.StackTrace);
            }
        }

        private void Label3_Click(object sender, EventArgs e)
        {

        }

        private void Label5_Click(object sender, EventArgs e)
        {

        }

        private void Label6_Click(object sender, EventArgs e)
        {

        }

        private void ComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            SelectResultatenhet(comboBox2.SelectedItem.ToString());

        }
    }//class
}//namespace

