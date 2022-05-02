
using OfficeOpenXml;
using OfficeOpenXml.Style;
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

        DateTime startMonth;
        DateTime endMonth;
        DateTime startYear;
        DateTime endYear;
        DateTime startPreviousYear;
        DateTime endPreviousYear = DateTime.Parse("2018-04-30");

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

                using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
                {

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
                        throw new IOException("Fel vid filinläsning");
                    }
                    // MessageBox.Show("Rader: " + list.Count);
                    button6.Enabled = true;
                    button1.Enabled = false;
                    Update();
                }
            }
            catch (IOException ex)
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


                //objekten.Sort();
                string valdTyp = comboBox2.SelectedItem.ToString();
                foreach (string v in valda)
                {
                    foreach (Objekt o in objekten)
                    {
                        if (o.Id.Equals(v) && dimensioner[o.Typ].Equals(valdTyp))
                        {

                            YearlyTotal(wb, o);

                            Console.WriteLine("Skapar flik " + o.Id + " " + o.Namn);
                            toolStripStatusLabel1.Text = "Skapar flik " + o.Id + " " + o.Namn;
                        }

                    }


                }
                Console.WriteLine("Antal flikar " + wb.Workbook.Worksheets.Count);

                toolStripStatusLabel1.Text = "Skapar flik TOTAL";

                YearlyTotal(wb, null);
                SkapaResultatsida(wb);
                SkapaResultatsidaPerManad(wb);
                wb.Workbook.Worksheets["TOTAL"].Select();
                //wb.Workbook.Calculate();

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
                            case "#RAR":
                                {
                                    string[] rar = currentRow.Split(' ');
                                    if (rar[1].Equals("0"))
                                    {
                                        startMonth = DateTime.Parse(rar[2].Substring(0, 4) + "-" + rar[2].Substring(4, 2) + "-" + rar[2].Substring(6, 2));

                                        endMonth = DateTime.Parse(rar[3].Substring(0, 4) + "-" + rar[3].Substring(4, 2) + "-" + rar[3].Substring(6, 2));
                                        endYear = endMonth;
                                    }
                                    if (rar[1].Equals("-1"))
                                    {
                                        startPreviousYear = DateTime.Parse(rar[2].Substring(0, 4) + "-" + rar[2].Substring(4, 2) + "-" + rar[2].Substring(6, 2));
                                        endPreviousYear = DateTime.Parse(rar[3].Substring(0, 4) + "-" + rar[3].Substring(4, 2) + "-" + rar[3].Substring(6, 2));
                                    }
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
                button3.Enabled = false;
                SelectResultatenhet(comboBox2.SelectedItem.ToString());

            }//try
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace);
            }//Catch

        } //button3_click

        private void SelectResultatenhet(string typ)
        {
            DateTime start = new DateTime(2018, 5, 1);
            DateTime end = new DateTime(2020, 4, 30);
            checkedListBox1.Items.Clear();
            string myKey = dimensioner.FirstOrDefault(x => x.Value.ToString().Equals(typ)).Key;
            foreach (Objekt o in objekten)
            {
                if (o.Typ.Equals(myKey))
                    checkedListBox1.Items.Add(o.Id, false);

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

                DateTime startYear = startMonth;




                int intaktsRad;
                int rorelseKostnadsRad;
                int bruttoRad;
                List<int> kostnadsRader = new List<int>();

                int kostnadsRad;

                string newname = "";

                if (resenh == null || o == null)
                {
                    newname = "Saknas";

                }
                else
                {
                    if (resenh.Equals("Saknas"))
                    {
                        newname = "Saknas";
                    }
                    else
                    {
                        newname = resenh + " " + o.Namn;
                        newname = newname.Substring(0, Math.Min(newname.Length - 1, 25));
                    }
                }
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
                totalSheet.Cells["A4"].Value = resenh;
                if (o != null)
                    totalSheet.Cells["B4"].Value = o.Namn;


                totalSheet.Cells["A2"].Value = "Från:";
                totalSheet.Cells["B2"].Value = startMonth;
                totalSheet.Cells["B2"].Style.Numberformat.Format = "yyyy-mm-dd";
                totalSheet.Cells["A3"].Value = "Till:";
                totalSheet.Cells["B3"].Value = endYear;
                totalSheet.Cells["B3"].Style.Numberformat.Format = "yyyy-mm-dd";
                int currmonth = endYear.Month;
                if (currmonth < 5)
                    currmonth = 7 + currmonth;
                else
                    currmonth = currmonth - 5;
                char currm = alpha[currmonth + 3];
                DateTime monthPrevYear = new DateTime(endYear.Year, endYear.Month, 1).AddYears(-1);
                //Skapa tabellerna

                Dictionary<string, double>[] months = new Dictionary<string, double>[12];
                Dictionary<string, double> monthPrev = new Dictionary<string, double>();
                for (int m = 0; m < 12; m++)
                {
                    // Console.WriteLine(m + ":" + startMonth.AddMonths(m).ToShortDateString() + " -- " + endMonth.AddMonths(m).ToShortDateString());

                    months[m] = SumTransaktion(new DateTime(startMonth.AddMonths(m).Ticks), EndOfMonth(startMonth.AddMonths(m)), resenh);
                }
                // Console.WriteLine(m + ":" + startMonth.AddMonths(m).ToShortDateString() + " -- " + endMonth.AddMonths(m).ToShortDateString());

                monthPrev = SumTransaktion(monthPrevYear, EndOfMonth(monthPrevYear), resenh);


                Dictionary<string, double> sumYTD = SumTransaktion(startYear, endYear, resenh);
                Dictionary<string, double> sumLastYTD = SumTransaktion(startPreviousYear, endYear.AddYears(-1), resenh);


                int row = 6;
                totalSheet.Cells["A5"].Value = "Konto";
                totalSheet.Cells["B5"].Value = "Benämning";
                for (int m = 0; m < 12; m++)
                    totalSheet.Cells[alpha[m + 3].ToString() + "5"].Value = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(startMonth.AddMonths(m).Month);
                totalSheet.Cells["P5"].Value = "Resultat nuvarande månad";
                totalSheet.Cells["Q5"].Value = "Diff mån fg mån";
                totalSheet.Cells["R5"].Value = "Ack resultat";
                totalSheet.Cells["S5"].Value = "Ack resultat fg år";
                totalSheet.Cells["T5"].Value = "Differens";
                totalSheet.Cells["U4"].Value = "3";
                totalSheet.Cells["U5"].Value = "Prognos månader";
                totalSheet.Cells["V5"].Value = "Tillägg";
                totalSheet.Cells["W5"].Value = "Prognos hela året";



               row++;
                totalSheet.Cells["B" + row].Value = "Intäkter";
                row++;
                //3000-3999
                int startrow = row;
                List<int> sorteradeKonton = new List<int>();

                foreach (string s in konton.Keys)
                {

                    for (int m = 0; m < 12; m++)
                    {
                        if (months[m].ContainsKey(s) && Math.Abs(months[m][s]) >= 0.1)
                        {
                            int val = Int32.Parse(s.Trim());
                            if (!sorteradeKonton.Contains(val) && val >= 3000 && val <= 8999)
                                sorteradeKonton.Add(val);
                        }


                    }

                }

                sorteradeKonton.Sort();
                int[] sorteradArray = sorteradeKonton.ToArray();

                foreach (int s in sorteradArray)
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

                        double prev = 0;
                        double curr = 0;
                        monthPrev.TryGetValue(konto.ToString(), out prev);
                        months[currmonth].TryGetValue(konto.ToString(), out curr);
                        totalSheet.Cells["P" + row].Value = curr;
                        totalSheet.Cells["Q" + row].Value = curr - prev;

                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["R" + row].Value = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells["R" + row].Value = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["S" + row].Value = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["S" + row].Value = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["T" + row].Value = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["T" + row].Value = 0.0;
                        totalSheet.Cells["U" + row].Formula = "P" + row + "*$U$4";
                        totalSheet.Cells["W" + row].Formula = "R"+ row + "+U" + row + "+V"+row;
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
                for (int c = 3; c < 24; c++)
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
                foreach (int s in sorteradArray)
                {

                    int konto = s;
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

                        double prev = 0;
                        double curr = 0;
                        monthPrev.TryGetValue(konto.ToString(), out prev);
                        months[currmonth].TryGetValue(konto.ToString(), out curr);
                        totalSheet.Cells["P" + row].Value = curr;
                        totalSheet.Cells["Q" + row].Value = curr - prev;
                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["R" + row].Value = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells["R" + row].Value = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["S" + row].Value = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["S" + row].Value = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["T" + row].Value = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["T" + row].Value = 0.0;
                        totalSheet.Cells["U" + row].Formula = "P" + row + "*$U$4";
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
                for (int c = 3; c < 24; c++)
                {
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "SUM(" + alpha[c - 1] + startrow + ":" + alpha[c - 1] + (row - 1) + ")";

                }
                rorelseKostnadsRad = row;
                row++;
                row++;
                bruttoRad = row;
                totalSheet.Cells["B" + row].Value = "Bruttovinst";
                for (int c = 3; c < 24; c++)
                {
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "SUM(" + alpha[c - 1] + intaktsRad + "+" + alpha[c - 1] + rorelseKostnadsRad + ")";

                }

                row++;
                row++;
                totalSheet.Cells["B" + row].Value = "TB1";
                for (int c = 3; c < 24; c++)
                {
                    if (c!=17 && c!=20 && c != 22)
                        totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "IFERROR("+ alpha[c - 1].ToString() + (row - 2) + "/" + alpha[c - 1].ToString() + rorelseKostnadsRad + ",0)";

                }
                int TB1Rad = row;
                row++;
                row++;

                //5000-6999 Externa kostnader
                totalSheet.Cells["B" + row].Value = "Externa kostnader";
                row++;
                startrow = row;
                foreach (int s in sorteradArray)
                {

                    int konto = s;
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

                        double prev = 0;
                        double curr = 0;
                        monthPrev.TryGetValue(konto.ToString(), out prev);
                        months[currmonth].TryGetValue(konto.ToString(), out curr);
                        totalSheet.Cells["P" + row].Value = curr;
                        totalSheet.Cells["Q" + row].Value = curr - prev;
                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["R" + row].Value = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells["R" + row].Value = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["S" + row].Value = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["S" + row].Value = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["T" + row].Value = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["T" + row].Value = 0.0;
                        totalSheet.Cells["U" + row].Formula = "P" + row + "*$U$4";
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
                for (int c = 3; c < 24; c++)
                {
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "SUM(" + alpha[c - 1] + startrow + ":" + alpha[c - 1] + (row - 1) + ")";

                }
                row++;

                row++;


                //7000-7799 Personalkostnader
                totalSheet.Cells["B" + row].Value = "Personalkostnader";
                row++;
                startrow = row;
                foreach (int s in sorteradArray)
                {

                    int konto = s;
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

                        double prev = 0;
                        double curr = 0;
                        monthPrev.TryGetValue(konto.ToString(), out prev);
                        months[currmonth].TryGetValue(konto.ToString(), out curr);
                        totalSheet.Cells["P" + row].Value = curr;
                        totalSheet.Cells["Q" + row].Value = curr - prev;
                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["R" + row].Value = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells["R" + row].Value = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["S" + row].Value = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["S" + row].Value = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["T" + row].Value = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["T" + row].Value = 0.0;
                        totalSheet.Cells["U" + row].Formula = "P" + row + "*$U$4";
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
                for (int c = 3; c < 24; c++)
                {
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "SUM(" + alpha[c - 1] + startrow + ":" + alpha[c - 1] + (row - 1) + ")";

                }
                row++;

                row++;

                //7800-7999 Avskrivningar
                totalSheet.Cells["B" + row].Value = "Avskrivningar";
                row++;
                startrow = row;
                foreach (int s in sorteradArray)
                {

                    int konto = s;
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
                        double prev = 0;
                        double curr = 0;
                        monthPrev.TryGetValue(konto.ToString(), out prev);
                        months[currmonth].TryGetValue(konto.ToString(), out curr);
                        totalSheet.Cells["P" + row].Value = curr;
                        totalSheet.Cells["Q" + row].Value = curr - prev;
                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["R" + row].Value = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells["R" + row].Value = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["S" + row].Value = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["S" + row].Value = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["T" + row].Value = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["T" + row].Value = 0.0;
                        totalSheet.Cells["U" + row].Formula = "P" + row + "*$U$4";
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
                for (int c = 3; c < 24; c++)
                {
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "SUM(" + alpha[c - 1] + startrow + ":" + alpha[c - 1] + (row - 1) + ")";

                }
                row++;

                row++;

                //8000-8799 Finansiella kostnader
                totalSheet.Cells["B" + row].Value = "Finansiella kostnader";
                row++;
                startrow = row;
                foreach (int s in sorteradArray)
                {

                    int konto = s;
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

                        double prev = 0;
                        double curr = 0;
                        monthPrev.TryGetValue(konto.ToString(), out prev);
                        months[currmonth].TryGetValue(konto.ToString(), out curr);
                        totalSheet.Cells["P" + row].Value = curr;
                        totalSheet.Cells["Q" + row].Value = curr - prev;
                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["R" + row].Value = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells["R" + row].Value = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["S" + row].Value = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["S" + row].Value = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["T" + row].Value = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["T" + row].Value = 0.0;
                        totalSheet.Cells["U" + row].Formula = "P" + row + "*$U$4";
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
                for (int c = 3; c < 24; c++)
                {
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "SUM(" + alpha[c - 1] + startrow + ":" + alpha[c - 1] + (row - 1) + ")";

                }
                row++;

                row++;

                //8800-8999 Bokslutsdispositioner
                totalSheet.Cells["B" + row].Value = "Bokslutsdispositioner";
                row++;
                startrow = row;
                foreach (int s in sorteradArray)
                {

                    int konto = s;
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
                        double prev = 0;
                        double curr = 0;
                        monthPrev.TryGetValue(konto.ToString(), out prev);
                        months[currmonth].TryGetValue(konto.ToString(), out curr);
                        totalSheet.Cells["P" + row].Value = curr;
                        totalSheet.Cells["Q" + row].Value = curr - prev;

                        if (sumYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["R" + row].Value = sumYTD[konto.ToString()];
                        else
                            totalSheet.Cells["R" + row].Value = 0.0;
                        if (sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["S" + row].Value = sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["S" + row].Value = 0.0;
                        if (sumYTD.ContainsKey(konto.ToString()) & sumLastYTD.ContainsKey(konto.ToString()))
                            totalSheet.Cells["T" + row].Value = sumYTD[konto.ToString()] - sumLastYTD[konto.ToString()];
                        else
                            totalSheet.Cells["T" + row].Value = 0.0;
                        totalSheet.Cells["U" + row].Formula = "P" + row + "*$U$4";
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
                for (int c = 3; c < 24; c++)
                {
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "SUM(" + alpha[c - 1] + startrow + ":" + alpha[c - 1] + (row - 1) + ")";

                }
                row++;
                row++;

                kostnadsRad = row;
                totalSheet.Cells["B" + row].Value = "Summa kostnader";
                for (int c = 3; c < 24; c++)
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
                for (int c = 3; c < 24; c++)
                {
                    totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "SUM(" + alpha[c - 1] + bruttoRad + "+" + alpha[c - 1] + kostnadsRad + ")";

                }

                //Group and hide months
                for (int r = 3; r < 16; r++)
                {
                    totalSheet.Column(r).OutlineLevel = 1;
                    totalSheet.Column(r).Collapsed = true;
                }
                totalSheet.Cells["C7:Z" + row].Style.Numberformat.Format = "#,### ;[Red]-#,### ";
                totalSheet.Cells["A1:Z" + row].AutoFitColumns();
                row++;
                row++;
                //Calculate TB2 in this row
                totalSheet.Cells["B" + row].Value = "TB2";
                for (int c = 3; c < 24; c++)
                {
                    if (c!=17 && c!=20 && c!=22)
                     totalSheet.Cells[alpha[c - 1] + "" + row].Formula = "IFERROR(" + alpha[c - 1] + (row-2) + "/" + alpha[c - 1] + intaktsRad + ",0)";

                }
                //Set borders
                using (var range = totalSheet.Cells["P5:Q" + row])
                {
                    range.Style.Border.BorderAround(ExcelBorderStyle.Thick);
                }
                using (var range = totalSheet.Cells["R5:T" + row])
                {
                    range.Style.Border.BorderAround(ExcelBorderStyle.Thick);
                }
                using (var range = totalSheet.Cells["U5:W" + row])
                {
                    range.Style.Border.BorderAround(ExcelBorderStyle.Thick);
                }
                //Add column gaps for easier reading
                totalSheet.InsertColumn(18, 1);
                totalSheet.InsertColumn(22, 1);

                //Style TB1 row
                totalSheet.Cells["D" + TB1Rad + ":Y" + TB1Rad].Style.Numberformat.Format = "0%";
                totalSheet.Cells["B" + TB1Rad + ":Y" + TB1Rad].Style.Fill.PatternType = ExcelFillStyle.Solid;
                totalSheet.Cells["B" + TB1Rad + ":Y" + TB1Rad].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                //Style TB2 row
                totalSheet.Cells["D" + row + ":Y" + row].Style.Numberformat.Format = "0%";
                totalSheet.Cells["B" + row + ":Y" + row].Style.Fill.PatternType = ExcelFillStyle.Solid;
                totalSheet.Cells["B" + row + ":Y" + row].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                totalSheet.View.FreezePanes(6, 2);
                wb.Workbook.Calculate();
                //Ta bort om resultat för båda åren saknas


                double sumq = (double)totalSheet.Cells["S" + (row-2)].Value;
                double sumr = (double)totalSheet.Cells["T" + (row-2)].Value;

                if (sumr == 0 && sumq == 0)
                    wb.Workbook.Worksheets.Delete(totalSheet);




            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + Environment.NewLine + ex.StackTrace);
            }
        }


        private static DateTime EndOfMonth(DateTime datum)
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
            ExcelWorksheet overSheet = wb.Workbook.Worksheets.Add("Översikt YTD");


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

            foreach (ExcelWorksheet e in wb.Workbook.Worksheets)
            {
                string s = e.Name;
                if (s.Contains("Översikt") || s.Contains("TOTAL"))
                {

                }
                else
                {
                    overSheet.Cells["A" + i].Value = s;

                    overSheet.Cells["C" + i].Formula = "VLOOKUP(C1,'" + s + "'!B:S, 18,0)";
                    overSheet.Cells["D" + i].Formula = "VLOOKUP(D1,'" + s + "'!B:S, 18,0)";
                    overSheet.Cells["E" + i].Formula = "VLOOKUP(E1,'" + s + "'!B:S, 18,0)";
                    overSheet.Cells["F" + i].Formula = "VLOOKUP(F1,'" + s + "'!B:S, 18,0)";
                    overSheet.Cells["G" + i].Formula = "VLOOKUP(G1,'" + s + "'!B:S, 18,0)";
                    overSheet.Cells["H" + i].Formula = "VLOOKUP(H1,'" + s + "'!B:S, 18,0)";
                    overSheet.Cells["I" + i].Formula = "VLOOKUP(I1,'" + s + "'!B:S, 18,0)";
                    overSheet.Cells["J" + i].Formula = "VLOOKUP(J1,'" + s + "'!B:S, 18,0)";
                    overSheet.Cells["K" + i].Formula = "VLOOKUP(K1,'" + s + "'!B:S, 18,0)";
                    overSheet.Cells["L" + i].Formula = "VLOOKUP(L1,'" + s + "'!B:S, 18,0)";
                    i++;
                }
            }
            overSheet.Cells["A" + i].Value = "TOTAL";

            overSheet.Cells[i, 3].Formula = "VLOOKUP(C1," + "TOTAL" + "!B:S, 18,0)";
            overSheet.Cells[i, 4].Formula = "VLOOKUP(D1," + "TOTAL" + "!B:S, 18,0)";
            overSheet.Cells[i, 5].Formula = "VLOOKUP(E1," + "TOTAL" + "!B:S, 18,0)";
            overSheet.Cells[i, 6].Formula = "VLOOKUP(F1," + "TOTAL" + "!B:S, 18,0)";
            overSheet.Cells[i, 7].Formula = "VLOOKUP(G1," + "TOTAL" + "!B:S, 18,0)";
            overSheet.Cells[i, 8].Formula = "VLOOKUP(H1," + "TOTAL" + "!B:S, 18,0)";
            overSheet.Cells[i, 9].Formula = "VLOOKUP(I1," + "TOTAL" + "!B:S, 18,0)";
            overSheet.Cells[i, 10].Formula = "VLOOKUP(J1," + "TOTAL" + "!B:S, 18,0)";
            overSheet.Cells[i, 11].Formula = "VLOOKUP(K1," + "TOTAL" + "!B:S, 18,0)";
            overSheet.Cells[i, 12].Formula = "VLOOKUP(L1," + "TOTAL" + "!B:S, 18,0)";

            overSheet.Cells["C2:L" + i].Style.Numberformat.Format = "#,### ;[Red]-#,### ";
            overSheet.Calculate();
            overSheet.Cells["A1:L" + i].AutoFitColumns();

        }
        private void SkapaResultatsidaPerManad(ExcelPackage wb)
        {
            //För varje flik skapa en rad med 
            //Kolumner Summa Intäkt, summa rörelsekostnader, bruttovinst, summa externa kostnader, Summa personalkostnader, summa finansiella kostnader, summa kostnader, Resultat
            ExcelWorksheet overSheet = wb.Workbook.Worksheets.Add("Översikt per månad");


            overSheet.Cells["C1"].Formula = "TOTAL!D5";
            overSheet.Cells["D1"].Formula = "TOTAL!E5";
            overSheet.Cells["E1"].Formula = "TOTAL!F5";
            overSheet.Cells["F1"].Formula = "TOTAL!G5";
            overSheet.Cells["G1"].Formula = "TOTAL!H5";
            overSheet.Cells["H1"].Formula = "TOTAL!I5";
            overSheet.Cells["I1"].Formula = "TOTAL!J5";
            overSheet.Cells["J1"].Formula = "TOTAL!K5";
            overSheet.Cells["K1"].Formula = "TOTAL!L5";
            overSheet.Cells["L1"].Formula = "TOTAL!M5";
            overSheet.Cells["M1"].Formula = "TOTAL!N5";
            overSheet.Cells["N1"].Formula = "TOTAL!O5";
            overSheet.Cells["O1"].Value = "Summa";
            int i = 2;
            foreach (ExcelWorksheet e in wb.Workbook.Worksheets)
            {
                string s = e.Name;
                if (s.Contains("Översikt") || s.Contains("TOTAL"))
                {

                }
                else
                {
                    overSheet.Cells["A" + i].Value = s;

                    overSheet.Cells["C" + i].Formula = "VLOOKUP(\"Resultat\",'" + s + "'!B:O,3,0)";
                    overSheet.Cells["D" + i].Formula = "VLOOKUP(\"Resultat\",'" + s + "'!B:O,4,0)";
                    overSheet.Cells["E" + i].Formula = "VLOOKUP(\"Resultat\",'" + s + "'!B:O,5,0)";
                    overSheet.Cells["F" + i].Formula = "VLOOKUP(\"Resultat\",'" + s + "'!B:O,6,0)";
                    overSheet.Cells["G" + i].Formula = "VLOOKUP(\"Resultat\",'" + s + "'!B:O,7,0)";
                    overSheet.Cells["H" + i].Formula = "VLOOKUP(\"Resultat\",'" + s + "'!B:O,8,0)";
                    overSheet.Cells["I" + i].Formula = "VLOOKUP(\"Resultat\",'" + s + "'!B:O,9,0)";
                    overSheet.Cells["J" + i].Formula = "VLOOKUP(\"Resultat\",'" + s + "'!B:O,10,0)";
                    overSheet.Cells["K" + i].Formula = "VLOOKUP(\"Resultat\",'" + s + "'!B:O,11,0)";
                    overSheet.Cells["L" + i].Formula = "VLOOKUP(\"Resultat\",'" + s + "'!B:O,12,0)";
                    overSheet.Cells["M" + i].Formula = "VLOOKUP(\"Resultat\",'" + s + "'!B:O,13,0)";
                    overSheet.Cells["N" + i].Formula = "VLOOKUP(\"Resultat\",'" + s + "'!B:O,14,0)";
                    overSheet.Cells["O" + i].Formula = "SUM(C" + i + ":N" + i + ")";
                    i++;
                }
            }
            overSheet.Cells["A" + i].Value = "TOTAL";

            overSheet.Cells[i, 3].Formula = "VLOOKUP(\"Resultat\"," + "TOTAL" + "!B:O,3,0)";
            overSheet.Cells[i, 4].Formula = "VLOOKUP(\"Resultat\"," + "TOTAL" + "!B:O,4,0)";
            overSheet.Cells[i, 5].Formula = "VLOOKUP(\"Resultat\"," + "TOTAL" + "!B:O,5,0)";
            overSheet.Cells[i, 6].Formula = "VLOOKUP(\"Resultat\"," + "TOTAL" + "!B:O,6,0)";
            overSheet.Cells[i, 7].Formula = "VLOOKUP(\"Resultat\"," + "TOTAL" + "!B:O,7,0)";
            overSheet.Cells[i, 8].Formula = "VLOOKUP(\"Resultat\"," + "TOTAL" + "!B:O,8,0)";
            overSheet.Cells[i, 9].Formula = "VLOOKUP(\"Resultat\"," + "TOTAL" + "!B:O,9,0)";
            overSheet.Cells[i, 10].Formula = "VLOOKUP(\"Resultat\"," + "TOTAL" + "!B:O,10,0)";
            overSheet.Cells[i, 11].Formula = "VLOOKUP(\"Resultat\"," + "TOTAL" + "!B:O,11,0)";
            overSheet.Cells[i, 12].Formula = "VLOOKUP(\"Resultat\"," + "TOTAL" + "!B:O,12,0)";
            overSheet.Cells[i, 13].Formula = "VLOOKUP(\"Resultat\"," + "TOTAL" + "!B:O,13,0)";
            overSheet.Cells[i, 14].Formula = "VLOOKUP(\"Resultat\"," + "TOTAL" + "!B:O,14,0)";
            overSheet.Cells[i, 15].Formula = "SUM(C" + i + ":N" + i + ")";


            overSheet.Cells["C2:O" + i].Style.Numberformat.Format = "#,### ;[Red]-#,### ";
            overSheet.Calculate();
            overSheet.Cells["A1:O" + i].AutoFitColumns();

        }

        private Dictionary<string, double> SumTransaktion(DateTime from, DateTime to, string resenh)
        {

            Dictionary<string, double> sum = new Dictionary<string, double>();

            foreach (Transaktion t in transaktioner)
            {

                if (to <= endMonth)
                {
                    if (!sum.ContainsKey(t.Kontonr))
                        sum.Add(t.Kontonr, 0.0); //Add to dictionary if not already in it
                    if (t.Objekt.ContainsKey(selectedType))
                        if (t.Objekt[selectedType].Equals(resenh) || resenh.Equals("*"))
                            if (t.Transaktionsdatum >= from & t.Transaktionsdatum <= to)
                                sum[t.Kontonr] += t.Belopp; // Add t.Belopp to Sum if resenh == t.Id

                }
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
            button4.Enabled = false;

            //Ställ in datum
            DateTime newEndMonth = DateTime.Parse(comboBox1.SelectedItem.ToString() + "-01");
            endMonth = new DateTime(newEndMonth.Year, newEndMonth.Month, DateTime.DaysInMonth(newEndMonth.Year, newEndMonth.Month));
            endYear = endMonth;
            Console.WriteLine(newEndMonth + "|" + endMonth);




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

                using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
                {

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
                    button6.Enabled = false;
                }
            }
            catch (IOException ex)
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

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }//class
}//namespace

