using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using MySql.Data.MySqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.IO;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace actualGOST
{
    public partial class Form1 : Form
    {
        bool writeInfo;
        string upd;   //переменная для хранения информации по ГОСТу
        //int numCol = Convert.ToInt32(textBox2.Text); //переменная для номера колонки

        //Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();
        //Microsoft.Office.Interop.Word.Document docWord = new Microsoft.Office.Interop.Word.Document();
        MySqlConnection mycon;
        MySqlCommand mycom;

        string constr = "Server=servgost;" +
                                "port=3306;" +
                                "Database=gost;" +
                                "Uid=admin;" +
                                "Pwd=;" +
                                "CharSet = cp1251;" +
                                "Allow Zero Datetime=true; ";
        public Form1()
        {
            InitializeComponent();
            
            mycon = new MySqlConnection(constr);
        }

        public void Find_old()
        {
            label2.Text = "";
            richTextBox1.AppendText("\n" + textBox1.Text + ":\n");
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1[1, i].FormattedValue.ToString().Contains(textBox1.Text.Trim()))
                {
                    label2.Text = "1";
                    upd = textBox1.Text + " | " + dataGridView1[1, i].Value + " ";

                    //else
                    //{
                    //MessageBox.Show(dataGridView1.CurrentCell.Value + " || " + textBox1.Text.Trim());
                    if (Convert.ToString(dataGridView1[4, i].Value) == "" &             // проверка на непустые значения во всех(!) ячейках
                        Convert.ToString(dataGridView1[5, i].Value) == "" &             // ГОСТы с пустыми ячейками, не находятся в "Представительстве"
                        Convert.ToString(dataGridView1[6, i].Value) == "" &
                        Convert.ToString(dataGridView1[11, i].Value) == "" &
                        Convert.ToString(dataGridView1[12, i].Value) == "" &
                        Convert.ToString(dataGridView1[13, i].Value) == "" &
                        Convert.ToString(dataGridView1[14, i].Value) == "")
                    {
                        richTextBox1.AppendText("Информация по ГОСТу не найдена...\n");
                        //richTextBox1.AppendText("aughdjkhjkhhhhhhhhhhhhhhhhh\n");
                    }

                    else                                                                // иначе проходим по всем ячейкам в строке
                    {
                        // проверяем на непустые значения
                        if (Convert.ToString(dataGridView1[4, i].Value) != "")
                        {
                            richTextBox1.AppendText("Дата Регистрации : " + dataGridView1[4, i].Value + "\n");
                            upd = upd + "Дата Регистрации : " + dataGridView1[4, i].Value + "; ";
                        }

                        if (Convert.ToString(dataGridView1[5, i].Value) != "")
                        {
                            richTextBox1.AppendText("Дата начала : " + dataGridView1[5, i].Value + "\n");
                            upd = upd + "Дата начала : " + dataGridView1[5, i].Value + "; ";
                        }

                        if (Convert.ToString(dataGridView1[6, i].Value) != "")
                        {
                            richTextBox1.AppendText("Дата окончания : " + dataGridView1[6, i].Value + "\n");
                            upd = upd + "Дата окончания : " + dataGridView1[6, i].Value + "; ";
                        }

                        if (Convert.ToString(dataGridView1[11, i].Value) != "")
                        {
                            richTextBox1.AppendText("Взамен прошлого : " + dataGridView1[11, i].Value + "\n");
                            upd = upd + "Взамен прошлого : " + dataGridView1[11, i].Value + "; ";
                        }

                        if (Convert.ToString(dataGridView1[12, i].Value) != "")
                        {
                            richTextBox1.AppendText("Заменяющий : " + dataGridView1[12, i].Value + "\n");
                            upd = upd + "Заменяющий : " + dataGridView1[12, i].Value + "; ";
                        }

                        if (Convert.ToString(dataGridView1[13, i].Value) != "")
                        {
                            richTextBox1.AppendText("Частично заменяющий : " + dataGridView1[13, i].Value + "\n");
                            upd = upd + "Частично заменяющий : " + dataGridView1[13, i].Value + "; ";
                        }

                        if (Convert.ToString(dataGridView1[14, i].Value) != "")
                        {
                            richTextBox1.AppendText("Частично замененный : " + dataGridView1[14, i].Value + "\n");
                            upd = upd + "Частично замененный : " + dataGridView1[14, i].Value + "; ";
                        }

                        // действующий/недействующий
                        if (Convert.ToString(dataGridView1[6, i].Value) == "")
                        {
                            richTextBox1.AppendText("Действующий НД\n");
                            upd = upd + "";
                            //upd = "Действующий НД";
                        }
                        else
                            if (Convert.ToDateTime(dataGridView1[6, i].Value) < DateTime.Now)
                            {
                                label2.Text = "2";
                                richTextBox1.AppendText("Недействующий НД\n");
                                upd = "Недействующий НД  " + upd;

                            }
                            else
                            {
                                richTextBox1.AppendText("Действующий НД\n");
                                upd = upd + "";
                                //upd = "Действующий НД";
                            }
                    }
                }
            }
        }

        private void Form1_Load_old(object sender, EventArgs e)
        {
            try
            {
                //ServGost:8080
                //string constr = "Server=servgost;" +
                //                "port=3306;" +
                //                "Database=gost;" +
                //                "Uid=admin;" +
                //                "Pwd=;" +
                //                "CharSet = cp1251; ";
                mycon = new MySqlConnection(constr);
                mycon.Open();

                mycom = new MySqlCommand(@"SELECT * FROM gost.standard ", mycon);
                //        MessageBox.Show("CONNECTED !");
                mycom.CommandType = CommandType.Text;

                MySqlDataAdapter adapter = new MySqlDataAdapter();

                adapter.TableMappings.Add("Table", "standard");

                adapter.SelectCommand = mycom;

                DataTable dataTable = new DataTable();
                DataSet dataSet = new DataSet("standard");
                adapter.Fill(dataTable);
                dataGridView1.DataSource = dataTable;                
            }

            catch (InvalidCastException ee)
            {
                MessageBox.Show("Нет подключения к серверу" + ee.Message);
            }
        }

        private void Form1_FormClosing_old(object sender, System.Windows.Forms.FormClosingEventArgs e)
        {
            mycon.Close();
        }

        private void button2_Click_old(object sender, EventArgs e)
        {
            if (textBox2.Text != "")
            {
                int numCol = Convert.ToInt32(textBox2.Text); //переменная для номера колонки

                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Filter = "Доступные форматы (*.xls ; *.xlsx)|*.xls; *.xlsx";
                openFileDialog1.FilterIndex = 2;
                openFileDialog1.RestoreDirectory = true;
                openFileDialog1.Title = "Select an Excel File";

                if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    //таймер
                    //считает время выполнения программы
                    Stopwatch testStopwatch = new Stopwatch();
                    testStopwatch.Start();

                    button2.Enabled = false;
                    button2.Enabled = false;
                    textBox1.ReadOnly = true;
                    textBox2.ReadOnly = true;
                    richTextBox1.Clear();

                    // цикл по вытаскиванию строчки из xls
                    string fileName = openFileDialog1.FileName;

                    Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель
                    Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(fileName, Type.Missing,
                               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                               Type.Missing, Type.Missing, Type.Missing); //открыть файл
                    Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист

                    var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку
                    string[] str = new string[lastCell.Row]; // массив значений с листа равен по размеру листу

                    for (int j = 0; j < lastCell.Row; j++) // по всем строкам
                    {
                        str[j] = ObjWorkSheet.Cells[j + 1, numCol].Text.ToString();//считываем текст в строку

                        //вытаскиваем регулярками ГОСТ
                        string input = str[j];
                        string pattern = @"((\bГОСТ\b)(\D)*([^А-я,(])*)";
                        Regex regex = new Regex(pattern);

                        // Получаем совпадения в экземпляре класса Match
                        Match match = regex.Match(input);
                        // отображаем все совпадения
                        //MessageBox.Show("" + match.Value);
                        //записываем в textBox
                        if (match.Value != null)
                            textBox1.Text = match.Value;

                        if (textBox1.Text != "") // проверка на пустые строки
                        {
                            Find();              // функция выполнения поиска госта
                                                 //WriteToRTB();        // функция записи в rtb

                            if (label2.Text == "2")
                            {
                                ObjWorkSheet.Cells[j + 1, numCol + 2].Value = upd;
                                ObjWorkSheet.Cells[j + 1, numCol + 2].Font.Color = Excel.XlRgbColor.rgbRed;
                            }
                            else if (label2.Text == "1")
                            {
                                ObjWorkSheet.Cells[j + 1, numCol + 2].Value = textBox1.Text + " действующий";
                            }
                            else
                            {
                                richTextBox1.AppendText("Информация по ГОСТу не найдена...\n");
                                ObjWorkSheet.Cells[j + 1, numCol + 2].Value = textBox1.Text + "  Информация по ГОСТу не найдена...";
                                ObjWorkSheet.Cells[j + 1, numCol + 2].Font.Color = Excel.XlRgbColor.rgbOrangeRed;
                            }
                        }
                    }

                    ObjWorkExcel.Visible = true;
                    button2.Enabled = true;
                    button2.Enabled = true;
                    textBox1.ReadOnly = false;
                    textBox2.ReadOnly = false;
                    //object missing = Type.Missing;

                    //останавливаем таймер
                    testStopwatch.Stop();
                    TimeSpan tSpan; tSpan = testStopwatch.Elapsed;
                    MessageBox.Show("Время выполнения операции - " + tSpan.ToString());
                }

                //object missing = Type.Missing;
                //docWord = appWord.Documents.Add(ref missing, false, ref missing, true);
                //appWord.ActiveDocument.Content.FormattedText.InsertAfter(richTextBox1.Text);         // функция записи в doc
                //appWord.Visible = true;
            }
            else
                MessageBox.Show("Укажите номер колонки!");
        }







        //НОВАЯ ПРОГРАМУЛИНА
        private void Form1_Load_1(object sender, EventArgs e)
        {
            progressBar1.Visible = false;
            try
            {
                mycon = new MySqlConnection(constr);
                mycon.Open();

                mycom = new MySqlCommand(@"SELECT * FROM gost.s_service ", mycon);
                //        MessageBox.Show("CONNECTED !");
                mycom.CommandType = CommandType.Text;

                MySqlDataAdapter adapter = new MySqlDataAdapter();

                adapter.TableMappings.Add("Table", "s_service");

                adapter.SelectCommand = mycom;

                DataTable dataTable = new DataTable();
                DataSet dataSet = new DataSet("s_service");
                adapter.Fill(dataTable);
                dataGridView1.DataSource = dataTable;
            }

            catch (InvalidCastException ee)
            {
                MessageBox.Show("Нет подключения к серверу" + ee.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            progressBar1.Value = 0;
            progressBar1.Visible = true;
            if (textBox2.Text != "")
            {
                int numCol = Convert.ToInt32(textBox2.Text); //переменная для номера колонки

                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Filter = "Доступные форматы (*.xls ; *.xlsx)|*.xls; *.xlsx";
                openFileDialog1.FilterIndex = 2;
                openFileDialog1.RestoreDirectory = true;
                openFileDialog1.Title = "Select an Excel File";

                if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    //таймер
                    //считает время выполнения программы
                    Stopwatch testStopwatch = new Stopwatch();
                    testStopwatch.Start();

                    button2.Enabled = false;
                    //button2.Enabled = false;
                    textBox1.ReadOnly = true;
                    textBox2.ReadOnly = true;
                    richTextBox1.Clear();

                    // цикл по вытаскиванию строчки из xls
                    string fileName = openFileDialog1.FileName;

                    Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель
                    Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(fileName, Type.Missing,
                               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                               Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                               Type.Missing, Type.Missing, Type.Missing); //открыть файл
                    Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист

                    var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку
                    string[] str = new string[lastCell.Row]; // массив значений с листа равен по размеру листу

                    for (int j = 0; j < lastCell.Row; j++) // по всем строкам
                    {
                        str[j] = ObjWorkSheet.Cells[j + 1, numCol].Text.ToString();//считываем текст в строку

                        //вытаскиваем регулярками ГОСТ
                        string input = str[j];
                        string pattern = @"((\bГОСТ\b)(\D)*([^А-я,(])*)";
                        Regex regex = new Regex(pattern);

                        // Получаем совпадения в экземпляре класса Match
                        Match match = regex.Match(input);
                        // отображаем все совпадения
                        //MessageBox.Show("" + match.Value);
                        //записываем в textBox
                        if (match.Value != null)
                            textBox1.Text = match.Value;

                        if (textBox1.Text != "") // проверка на пустые строки
                        {
                            Find();              // функция выполнения поиска госта
                                                 //WriteToRTB();        // функция записи в rtb

                            ObjWorkSheet.Cells[j + 1, numCol + 2].Value = upd;

                            if (label2.Text == "1")
                            {
                                ObjWorkSheet.Cells[j + 1, numCol + 2].Value = upd /*textBox1.Text + " действующий"*/;
                            }

                            else if (label2.Text == "2")
                            {
                                ObjWorkSheet.Cells[j + 1, numCol + 2].Value = upd;
                                ObjWorkSheet.Cells[j + 1, numCol + 2].Font.Color = Excel.XlRgbColor.rgbRed;
                            }

                            else
                            {
                                //richTextBox1.AppendText("Информация по ГОСТу не найдена...\n");
                                ObjWorkSheet.Cells[j + 1, numCol + 2].Value = textBox1.Text + "  Информация по ГОСТу не найдена...";
                                ObjWorkSheet.Cells[j + 1, numCol + 2].Font.Color = Excel.XlRgbColor.rgbOrangeRed;
                            }

                        }

                        progressBar1.Maximum = lastCell.Row;
                        progressBar1.Value = j;
                        //System.Threading.Thread.Sleep(100);  //задержка, но ну ее нахер...из часа тридцати сделала 3 часа
                    }


                    ObjWorkExcel.Visible = true;
                    button2.Enabled = true;
                    button2.Enabled = true;
                    textBox1.ReadOnly = false;
                    textBox2.ReadOnly = false;
                    //object missing = Type.Missing;

                    //останавливаем таймер
                    testStopwatch.Stop();
                    TimeSpan tSpan; tSpan = testStopwatch.Elapsed;
                    MessageBox.Show("Время выполнения операции - " + tSpan.ToString());
                }

                //object missing = Type.Missing;
                //docWord = appWord.Documents.Add(ref missing, false, ref missing, true);
                //appWord.ActiveDocument.Content.FormattedText.InsertAfter(richTextBox1.Text);         // функция записи в doc2
                //appWord.Visible = true;
            }
            else
                MessageBox.Show("Укажите номер колонки!");

            progressBar1.Visible = false;
        }

        public void Find()
        {
            label2.Text = "";
            richTextBox1.AppendText(textBox1.Text + "\n");
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1[2, i].FormattedValue.ToString().Contains(textBox1.Text.Trim()))
                {
                    upd = Convert.ToString(dataGridView1[2, i].Value) + " ";
                    //if (Convert.ToString(dataGridView1[8, i].Value) == "")
                    //{
                    //label2.Text = "0";
                    //richTextBox1.AppendText("Нет информации\n");
                    //upd = upd + "Нет информации";
                    //return;
                    //}

                    if (Convert.ToString(dataGridView1[8, i].Value) == "Утратил силу в РФ" | Convert.ToString(dataGridView1[8, i].Value) == "Отменен" | Convert.ToString(dataGridView1[8, i].Value) == "Заменен")
                    {
                        label2.Text = "2";
                        //richTextBox1.AppendText(Convert.ToString(dataGridView1[8, i].Value));
                        upd = upd + Convert.ToString(dataGridView1[8, i].Value);
                        return;
                    }

                    else
                    {
                        label2.Text = "1";
                        //richTextBox1.AppendText("Действующий НД\n");
                        upd = upd + Convert.ToString(dataGridView1[8, i].Value);
                        //upd = "Действующий НД";
                        return;
                    }
                }
            }
        }

        private void Form1_FormClosing(object sender, System.Windows.Forms.FormClosingEventArgs e)
        {
            mycon.Close();
        }
    }
}

/*
                string constr = "Server=servgost;" +
                                "port=3306;" +
                                "Database=gost;" +
                                "Uid=admin;" +
                                "Pwd=;" +
                                "CharSet = cp1251; ";
                mycon = new MySqlConnection(constr);
                mycon.Open();

                mycom = new MySqlCommand(@"SELECT * FROM gost.standard ", mycon);
                //        MessageBox.Show("CONNECTED !");
                mycom.CommandType = CommandType.Text;

                MySqlDataAdapter adapter = new MySqlDataAdapter();

                adapter.TableMappings.Add("Table", "standard");

                adapter.SelectCommand = mycom;

                DataTable dataTable = new DataTable();
                DataSet dataSet = new DataSet("standard");
                adapter.Fill(dataTable);
                dataGridView1.DataSource = dataTable;
*/