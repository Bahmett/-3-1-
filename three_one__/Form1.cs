using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using Microsoft.Office.Interop.Excel;
using System.IO;
using EXX = Microsoft.Office.Interop.Excel;


namespace three_one__
{
    public partial class Form1 : Form
    {


        private Microsoft.Office.Interop.Excel.Application ObjExcel;
        private Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
        private Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
        public static int chst,a, nehvatka, otschet;
        public double[] got_prod;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Файл за период|*.xls; *.xlsx";
            openDialog.ShowDialog();

            try
            {
                ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                //Книга.
                ObjWorkBook = ObjExcel.Workbooks.Open(openDialog.FileName);
                //Таблица.
                ObjWorkSheet = ObjExcel.ActiveSheet as Worksheet;
                Range rg = null;
                int LastRowNumber = ObjWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
                int LastColNumber = ObjWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Column;
                dataGridView1.ColumnCount = LastColNumber;
                dataGridView1.RowCount = LastRowNumber;
                // MessageBox.Show(LastRowNumber.ToString());
                Int32 row = 1;
                dataGridView1.Rows.Clear();
                List<String> arr = new List<string>();
                while (row != LastRowNumber + 1) //ObjWorkSheet.get_Range("a" + row, "a" + row).Value != null)
                {
                    // Читаем данные из ячейки
                    rg = ObjWorkSheet.get_Range("a" + row, "bb" + row);
                    foreach (Range item in rg)
                    {
                        try
                        {
                            arr.Add(item.Value.ToString().Trim());
                        }
                        catch
                        {
                            arr.Add("");
                        }
                    }
                    dataGridView1.Rows.Add(arr[0], arr[1], arr[2], arr[3], arr[4], arr[5], arr[6], arr[7], arr[8],
                        arr[9], arr[10], arr[11], arr[12], arr[13], arr[14], arr[15], arr[16], arr[17], arr[18], arr[19],
                        arr[20], arr[21], arr[22], arr[23], arr[24], arr[25], arr[26], arr[27], arr[28], arr[29], arr[30],
                        arr[31], arr[32], arr[33], arr[34], arr[35], arr[36], arr[37], arr[38], arr[39], arr[40], arr[41],
                        arr[42], arr[43], arr[44], arr[45]);
                    arr.Clear();
                    row++;
                }

                MessageBox.Show("Файл успешно считан!", "Считывания excel файла", MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message, "Ошибка при считывании excel файла", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                ObjWorkBook.Close(false, "", null);
                // Закрытие приложения Excel.
                ObjExcel.Quit();
                ObjWorkBook = null;
                ObjWorkSheet = null;
                ObjExcel = null;
                GC.Collect();
            }

           // button1.Enabled = false;
            button4.Enabled = true;
            button1.Enabled = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {

            int n = Convert.ToInt32(textBox1.Text);
            // обработка 1 и 3 грида
            
            #region
            /*
            
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1.Rows[i].Cells[1].Value.ToString().Contains("Подразделение"))
                {
                    dataGridView1.Rows[i + 1].Cells[1].Value = "XX";
                    dataGridView1.Rows[i + 2].Cells[1].Value = "XX";
                }
            }

         /*   for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Value.ToString() == "Остаток ГП (вагоны под погрузкой)")
                    {
                        int x = i;
                        int y = j;
                        if ((dataGridView1.Rows[x + 3].Cells[y].Value.ToString() != ""))
                        {
                            nehvatka = Convert.ToInt32(dataGridView1.Rows[x + 3].Cells[y].Value.ToString());
                        }
                        else 
                        {
                            nehvatka = 0;
                        }
                    }
                }

            }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1.Rows[i].Cells[1].Value.ToString().Equals(""))
                {
                    dataGridView1.Rows.RemoveAt(i);
                    i--;
                }
            }
            dataGridView1.Columns.RemoveAt(0);
            chst = dataGridView1.ColumnCount;

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (
                    (!(dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бр. Коростелевых ул., 52, терр. з-да \"Гидропресс\"") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бугуруслан г., Восточное ш. 1") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бузулук г., ул. Промышленная, 6") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Донгузская ул, 20") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Илек с., ул. Шоссейная, 54Б") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Краснохолм с., ул. Шоссейная 1А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Кувандык г., ул. Дзержинского, 1А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Курманаевка п.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Медногорск г., ул. Комсомольская, 40") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Новосергиевский п.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Октябрьское, ул. Транспортная, 1") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Орск, ул.Строителей, 44") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Переволоцкий п., ул. Ленинская, 2А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Подольск п., Промышленная ул., 3") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Сакмарский р-н, Сакмарская ст., терр. СМП-639") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Саракташ п., ул. Производственная, 4") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Соль-Илецк г., ул. Гонтаренко, 1А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Сорочинск г., ул. Пролетарская, 3") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Терешковой  ул., 287 д.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Тоцкое с. Тоцкий р-н, Автомобилистов ул., 1Е") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Центральная ул., 1 д.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Шильда, Топсклад") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("ИТОГО:") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Транзит") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("Отчет") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("Подразделение") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("Из расчета") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("XX"))))
                {
                    dataGridView1.Rows.RemoveAt(i);
                    i--;
                }
            }
            int bb;
            bb = 2*chst;
            DataGridViewTextBoxColumn[] column = new DataGridViewTextBoxColumn[chst];

            for (int i = 0; i < chst; i++)
            {
                column[i] = new DataGridViewTextBoxColumn();
            }
            dataGridView1.Columns.AddRange(column);

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Length > 17 &&
                    dataGridView1.Rows[i].Cells[0].Value.ToString().Substring(0, 17) == "Отчет по отгрузке")
                {
                    a = i;
                   
                    for (int j = a; j < dataGridView1.RowCount; j++)
                    {
                        int t = 0;
                        for (int k = chst; k < bb; k++)
                        {
                            dataGridView1.Rows[j - a].Cells[k].Value = dataGridView1.Rows[j].Cells[t].Value;
                            t++;
                        }
                    }
                }
            }


            for (int i = dataGridView1.RowCount - 1; i >= a + 1; i--)
            {
                dataGridView1.Rows.RemoveAt(i);
            }
            for (int i = dataGridView1.ColumnCount - 1; i >= 0; i--)
            {
                if (dataGridView1.Rows[2].Cells[i].Value.ToString().Equals("") &&
                    dataGridView1.Rows[3].Cells[i].Value.ToString().Equals("") &&
                    dataGridView1.Rows[4].Cells[i].Value.ToString().Equals("") &&
                    dataGridView1.Rows[5].Cells[i].Value.ToString().Equals("") &&
                    dataGridView1.Rows[6].Cells[i].Value.ToString().Equals("") &&
                    dataGridView1.Rows[7].Cells[i].Value.ToString().Equals(""))
                {
                    dataGridView1.Columns.RemoveAt(i);
                }
            }

            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                dataGridView1.Columns[i].HeaderText = (i + 1).ToString();
            }

          
            


             for (int i = dataGridView1.ColumnCount - 1; i >= 0; i--)
            {

                if (dataGridView1.Columns[i].HeaderText.Equals("30")||
                    dataGridView1.Columns[i].HeaderText.Equals("28")||
                    dataGridView1.Columns[i].HeaderText.Equals("27")||
                    dataGridView1.Columns[i].HeaderText.Equals("20")||
                    dataGridView1.Columns[i].HeaderText.Equals("18")||
                    dataGridView1.Columns[i].HeaderText.Equals("17")||
                    dataGridView1.Columns[i].HeaderText.Equals("15")||
                    dataGridView1.Columns[i].HeaderText.Equals("14")||
                    dataGridView1.Columns[i].HeaderText.Equals("13")||
                    dataGridView1.Columns[i].HeaderText.Equals("12")||
                    dataGridView1.Columns[i].HeaderText.Equals("10")||
                    dataGridView1.Columns[i].HeaderText.Equals("7")||
                    dataGridView1.Columns[i].HeaderText.Equals("4"))
                {
                   dataGridView1.Columns.RemoveAt(i); 
                }
            }
            

            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                dataGridView1.Columns[i].HeaderText = (i + 1).ToString();
            }

            for (int i = 5; i < dataGridView1.RowCount-1; i++)
                for (int j = 5; j < dataGridView1.RowCount-1; j++)
                {
                    if (dataGridView1.Rows[i].Cells[0].Value == dataGridView1.Rows[j].Cells[10])

                    {
                        string s = String.Empty;
                        s = dataGridView1.Rows[j].Cells[10].Value.ToString();
                        dataGridView1.Rows[j].Cells[10].Value = "";
                        dataGridView1.Rows[j].Cells[10].Value = dataGridView1.Rows[i].Cells[10].Value;
                        dataGridView1.Rows[i].Cells[10].Value = "";
                        dataGridView1.Rows[i].Cells[10].Value = s;
                    }

                }
            dataGridView1.Rows[2].Cells[2].Value = "";
            dataGridView1.Rows[3].Cells[2].Value = "";
            dataGridView1.Rows[2].Cells[4].Value = "";
            dataGridView1.Rows[3].Cells[4].Value = "";
            dataGridView1.Rows[2].Cells[6].Value = "";
            dataGridView1.Rows[3].Cells[6].Value = "";
            dataGridView1.Rows[2].Cells[8].Value = "";
            dataGridView1.Rows[3].Cells[8].Value = "";


            //4 button вторая обработка 1 грида
            
            int ii = 5;
            for (int j = 5; j < dataGridView1.RowCount - 1; j++)
            {
                if (dataGridView1.Rows[ii].Cells[0].Value.ToString() == dataGridView1.Rows[j].Cells[9].Value.ToString())
                {
                    string s1 = String.Empty;
                    string s2 = String.Empty;
                    string s3 = String.Empty;
                    string s4 = String.Empty;
                    string s5 = String.Empty;
                    string s6 = String.Empty;
                    string s7 = String.Empty;
                    string s8 = String.Empty;

                    s1 = dataGridView1.Rows[j].Cells[9].Value.ToString();
                    s2 = dataGridView1.Rows[j].Cells[10].Value.ToString();
                    s3 = dataGridView1.Rows[j].Cells[11].Value.ToString();
                    s4 = dataGridView1.Rows[j].Cells[12].Value.ToString();
                    s5 = dataGridView1.Rows[j].Cells[13].Value.ToString();
                    s6 = dataGridView1.Rows[j].Cells[14].Value.ToString();
                    s7 = dataGridView1.Rows[j].Cells[15].Value.ToString();
                    s8 = dataGridView1.Rows[j].Cells[16].Value.ToString();

                    dataGridView1.Rows[j].Cells[9].Value = "";
                    dataGridView1.Rows[j].Cells[10].Value = "";
                    dataGridView1.Rows[j].Cells[11].Value = "";
                    dataGridView1.Rows[j].Cells[12].Value = "";
                    dataGridView1.Rows[j].Cells[13].Value = "";
                    dataGridView1.Rows[j].Cells[14].Value = "";
                    dataGridView1.Rows[j].Cells[15].Value = "";
                    dataGridView1.Rows[j].Cells[16].Value = "";

                    dataGridView1.Rows[j].Cells[9].Value = dataGridView1.Rows[ii].Cells[9].Value;
                    dataGridView1.Rows[j].Cells[10].Value = dataGridView1.Rows[ii].Cells[10].Value;
                    dataGridView1.Rows[j].Cells[11].Value = dataGridView1.Rows[ii].Cells[11].Value;
                    dataGridView1.Rows[j].Cells[12].Value = dataGridView1.Rows[ii].Cells[12].Value;
                    dataGridView1.Rows[j].Cells[13].Value = dataGridView1.Rows[ii].Cells[13].Value;
                    dataGridView1.Rows[j].Cells[14].Value = dataGridView1.Rows[ii].Cells[14].Value;
                    dataGridView1.Rows[j].Cells[15].Value = dataGridView1.Rows[ii].Cells[15].Value;
                    dataGridView1.Rows[j].Cells[16].Value = dataGridView1.Rows[ii].Cells[16].Value;

                    dataGridView1.Rows[ii].Cells[9].Value = "";
                    dataGridView1.Rows[ii].Cells[10].Value = "";
                    dataGridView1.Rows[ii].Cells[11].Value = "";
                    dataGridView1.Rows[ii].Cells[12].Value = "";
                    dataGridView1.Rows[ii].Cells[13].Value = "";
                    dataGridView1.Rows[ii].Cells[14].Value = "";
                    dataGridView1.Rows[ii].Cells[15].Value = "";
                    dataGridView1.Rows[ii].Cells[16].Value = "";

                    dataGridView1.Rows[ii].Cells[9].Value = s1;
                    dataGridView1.Rows[ii].Cells[10].Value = s2;
                    dataGridView1.Rows[ii].Cells[11].Value = s3;
                    dataGridView1.Rows[ii].Cells[12].Value = s4;
                    dataGridView1.Rows[ii].Cells[13].Value = s5;
                    dataGridView1.Rows[ii].Cells[14].Value = s6;
                    dataGridView1.Rows[ii].Cells[15].Value = s7;
                    dataGridView1.Rows[ii].Cells[16].Value = s8;

                    ii++;
                    j = 5;


                }

            }

            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[0].Value = "";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[0].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[0].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[1].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[1].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[2].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[2].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[3].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[3].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[4].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[4].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[5].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[5].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[6].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[6].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[7].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[7].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[8].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[8].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[9].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[9].Value;

            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[0].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[1].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[2].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[3].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[4].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[5].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[6].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[7].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[8].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[9].Value = "-";

            //dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[9].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[10].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[11].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[12].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[13].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[14].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[15].Value = "-";

            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[0].Value = "Транзит"; //dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[9].Value;

            //dataGridView1.Rows[0].Cells[11].Value = dataGridView1.Rows[0].Cells[10].Value.ToString();
           // dataGridView1.Rows[1].Cells[11].Value = dataGridView1.Rows[1].Cells[10].Value.ToString();


            dataGridView1.Columns.RemoveAt(9);

            dataGridView1.Rows[2].Cells[0].Value = "Подразделение";

            for (int j = 0; j < dataGridView1.ColumnCount; j++)
            {
                dataGridView1.Columns[j].HeaderText = (j + 1).ToString();
            }

            string q1, q2, q3, q4, q5;
            q1 = dataGridView1.Rows[4].Cells[7].Value.ToString();
            dataGridView1.Rows[2].Cells[7].Value = "Откл-е на дату:";
            dataGridView1.Rows[4].Cells[7].Value = "";
            dataGridView1.Rows[4].Cells[7].Value = q1 + ",тонн";

            dataGridView1.Rows[4].Cells[8].Value = "";
            dataGridView1.Rows[4].Cells[8].Value ="Кол-во, %";

            dataGridView1.Rows[3].Cells[8].Value = dataGridView1.Rows[3].Cells[7].Value;
            dataGridView1.Rows[2].Cells[8].Value = "% вып-я на дату:";

            q2 = dataGridView1.Rows[4].Cells[8].Value.ToString();

            dataGridView1.Rows[2].Cells[9].Value = "План на период:";

            q3 = dataGridView1.Rows[3].Cells[10].Value.ToString();
            dataGridView1.Rows[3].Cells[10].Value = "";
            dataGridView1.Rows[3].Cells[10].Value = "План на: " + q3;

            q4 = dataGridView1.Rows[3].Cells[11].Value.ToString();
            dataGridView1.Rows[3].Cells[11].Value = "";
            dataGridView1.Rows[3].Cells[11].Value = "Отгрузка на: " + q4;

            q5 = dataGridView1.Rows[3].Cells[15].Value.ToString();
            dataGridView1.Rows[3].Cells[15].Value = "";
            dataGridView1.Rows[3].Cells[15].Value = "Ост.  " + q5;
            dataGridView1.Rows[1].Cells[10].Value = "";

           // dataGridView1.Rows[3].Cells[13].Value = dataGridView1.Rows[3].Cells[13].Value.ToString().Substring(13, (dataGridView1.Rows[3].Cells[13].Value.ToString().Length - 13));



                ///   //Обработка 1 и 3 грида
            */
#endregion


            #region

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1.Rows[i].Cells[1].Value.ToString().Contains("Подразделение"))
                {
                    dataGridView1.Rows[i + 1].Cells[1].Value = "XX";
                    dataGridView1.Rows[i + 2].Cells[1].Value = "XX";
                }
            }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1.Rows[i].Cells[1].Value.ToString().Equals(""))
                {
                    dataGridView1.Rows.RemoveAt(i);
                    i--;
                }
            }
            dataGridView1.Columns.RemoveAt(0);
            chst = dataGridView1.ColumnCount;

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("Отчет по отгрузке лома за период:"))
                {
                    otschet = i;
                }
            }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (
                    (!(dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бр. Коростелевых ул., 52, терр. з-да \"Гидропресс\"") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бугуруслан г., Восточное ш. 1") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бузулук г., ул. Промышленная, 6") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Донгузская ул, 20") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Илек с., ул. Шоссейная, 54Б") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Краснохолм с., ул. Шоссейная 1А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Кувандык г., ул. Дзержинского, 1А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Курманаевка п.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Медногорск г., ул. Комсомольская, 40") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Новосергиевский п.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Октябрьское, ул. Транспортная, 1") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Орск, ул.Строителей, 44") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Переволоцкий п., ул. Ленинская, 2А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Подольск п., Промышленная ул., 3") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Сакмарский р-н, Сакмарская ст., терр. СМП-639") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Саракташ п., ул. Производственная, 4") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Соль-Илецк г., ул. Гонтаренко, 1А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Сорочинск г., ул. Пролетарская, 3") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Терешковой  ул., 287 д.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Тоцкое с. Тоцкий р-н, Автомобилистов ул., 1Е") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Центральная ул., 1 д.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Шильда, Топсклад") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("ИТОГО:") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Транзит") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("Отчет") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("Подразделение") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("Из расчета") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("XX") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("2А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("2А1") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("3А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("3А1") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("3А2") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("3АР") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("3АР2") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("8А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("9А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("10А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("41 счет"))))
                {
                    dataGridView1.Rows.RemoveAt(i);
                    i--;
                }
            }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("Отчет по отгрузке лома за период:"))
                {
                    otschet = i;
                }
            }


            for (int i = otschet; i < dataGridView1.RowCount; i++)
            {
                if (
                    (!(dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бр. Коростелевых ул., 52, терр. з-да \"Гидропресс\"") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бугуруслан г., Восточное ш. 1") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бузулук г., ул. Промышленная, 6") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Донгузская ул, 20") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Илек с., ул. Шоссейная, 54Б") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Краснохолм с., ул. Шоссейная 1А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Кувандык г., ул. Дзержинского, 1А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Курманаевка п.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Медногорск г., ул. Комсомольская, 40") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Новосергиевский п.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Октябрьское, ул. Транспортная, 1") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Орск, ул.Строителей, 44") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Переволоцкий п., ул. Ленинская, 2А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Подольск п., Промышленная ул., 3") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Сакмарский р-н, Сакмарская ст., терр. СМП-639") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Саракташ п., ул. Производственная, 4") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Соль-Илецк г., ул. Гонтаренко, 1А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Сорочинск г., ул. Пролетарская, 3") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Терешковой  ул., 287 д.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Тоцкое с. Тоцкий р-н, Автомобилистов ул., 1Е") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Центральная ул., 1 д.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Шильда, Топсклад") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("ИТОГО:") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Транзит") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("Отчет") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("Подразделение") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("Из расчета") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("XX"))))
                {
                    dataGridView1.Rows.RemoveAt(i);
                    i--;
                }
            }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("41 счет"))
                {
                    dataGridView1.Rows.RemoveAt(i + 1);
                    dataGridView1.Rows.RemoveAt(i);
                    i--;
                }
            }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("2А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("2А1") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("3А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("3А1") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("3А2") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("3АР") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("3АР2") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("8А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("9А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("10А"))
                {
                    for (int j = 1; j < dataGridView1.ColumnCount - 1; j++)
                    {
                        dataGridView1.Rows[i].Cells[j].Value = "";
                    }
                }
            }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString() == "" || dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value == null || dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString().Length == 0)
                {
                    dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value = 0;
                }

            }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("Отчет по отгрузке лома за период:"))
                {
                    otschet = i;
                }
            }

            got_prod = new double[22];

            //double got_prod_i = 0;



            //1
            int verh = 0, niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бр. Коростелевых ул., 52, терр. з-да \"Гидропресс\""))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бугуруслан г., Восточное ш. 1"))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[0] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[0] += Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }


            //2
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бугуруслан г., Восточное ш. 1"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бузулук г., ул. Промышленная, 6"))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[1] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[1] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }
            //3
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бузулук г., ул. Промышленная, 6"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Донгузская ул, 20"))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[2] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[2] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            //4
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Донгузская ул, 20"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Илек с., ул. Шоссейная, 54Б"))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[3] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[3] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }
            //5
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Илек с., ул. Шоссейная, 54Б"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Краснохолм с., ул. Шоссейная 1А"))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[4] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[4] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }
            //6
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Краснохолм с., ул. Шоссейная 1А"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Кувандык г., ул. Дзержинского, 1А"))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[5] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[5] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }
            //7
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Кувандык г., ул. Дзержинского, 1А"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Курманаевка п."))
                {
                    niz = i;
                }

            }
            if (verh - niz == -1)
            {
                got_prod[6] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[6] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            //8
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Курманаевка п."))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Медногорск г., ул. Комсомольская, 40"))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[7] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[7] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            //9
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Медногорск г., ул. Комсомольская, 40"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Новосергиевский п."))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[8] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[8] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            //10
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Новосергиевский п."))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Октябрьское, ул. Транспортная, 1"))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[9] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[9] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            //11
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Октябрьское, ул. Транспортная, 1"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Орск, ул.Строителей, 44"))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[10] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[10] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            //12
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Орск, ул.Строителей, 44"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Переволоцкий п., ул. Ленинская, 2А"))
                {
                    niz = i;
                }

            }
            if (verh - niz == -1)
            {
                got_prod[11] = 0.0;
            }
            else
            {

                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[11] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }
            //13
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Переволоцкий п., ул. Ленинская, 2А"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Подольск п., Промышленная ул., 3"))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[12] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[12] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            //14
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Подольск п., Промышленная ул., 3"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Сакмарский р-н, Сакмарская ст., терр. СМП-639"))
                {
                    niz = i;
                }

            }
            if (verh - niz == -1)
            {
                got_prod[13] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[13] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }
            //15
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Сакмарский р-н, Сакмарская ст., терр. СМП-639"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Саракташ п., ул. Производственная, 4"))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[14] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[14] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            //16
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Саракташ п., ул. Производственная, 4"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Соль-Илецк г., ул. Гонтаренко, 1А"))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[15] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[15] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            //17
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Соль-Илецк г., ул. Гонтаренко, 1А"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Сорочинск г., ул. Пролетарская, 3"))
                {
                    niz = i;
                }

            }
            if (verh - niz == -1)
            {
                got_prod[16] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[16] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            //18
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Сорочинск г., ул. Пролетарская, 3"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Терешковой  ул., 287 д."))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[17] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[17] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            //19
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Терешковой  ул., 287 д."))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Тоцкое с. Тоцкий р-н, Автомобилистов ул., 1Е"))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[18] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[18] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            //20
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Тоцкое с. Тоцкий р-н, Автомобилистов ул., 1Е"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Центральная ул., 1 д."))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[19] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[19] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }
            //21
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Центральная ул., 1 д."))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Шильда, Топсклад"))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[20] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[20] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            //22
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Шильда, Топсклад"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("ИТОГО:"))
                {
                    niz = i;
                }

            }
            if (verh - niz == -1)
            {
                got_prod[21] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[21] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (
                    (!(dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бр. Коростелевых ул., 52, терр. з-да \"Гидропресс\"") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бугуруслан г., Восточное ш. 1") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бузулук г., ул. Промышленная, 6") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Донгузская ул, 20") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Илек с., ул. Шоссейная, 54Б") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Краснохолм с., ул. Шоссейная 1А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Кувандык г., ул. Дзержинского, 1А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Курманаевка п.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Медногорск г., ул. Комсомольская, 40") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Новосергиевский п.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Октябрьское, ул. Транспортная, 1") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Орск, ул.Строителей, 44") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Переволоцкий п., ул. Ленинская, 2А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Подольск п., Промышленная ул., 3") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Сакмарский р-н, Сакмарская ст., терр. СМП-639") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Саракташ п., ул. Производственная, 4") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Соль-Илецк г., ул. Гонтаренко, 1А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Сорочинск г., ул. Пролетарская, 3") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Терешковой  ул., 287 д.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Тоцкое с. Тоцкий р-н, Автомобилистов ул., 1Е") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Центральная ул., 1 д.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Шильда, Топсклад") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("ИТОГО:") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Транзит") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("Отчет") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("Подразделение") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("Из расчета") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("XX"))))
                {
                    dataGridView1.Rows.RemoveAt(i);
                    i--;
                }
            }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                dataGridView1.Rows[i].HeaderCell.Value = (i + 1).ToString();
            }

            int bb;
            bb = 2 * chst;
            DataGridViewTextBoxColumn[] column = new DataGridViewTextBoxColumn[chst];

            for (int i = 0; i < chst; i++)
            {
                column[i] = new DataGridViewTextBoxColumn();
            }
            dataGridView1.Columns.AddRange(column);

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Length > 17 &&
                    dataGridView1.Rows[i].Cells[0].Value.ToString().Substring(0, 17) == "Отчет по отгрузке")
                {
                    a = i;

                    for (int j = a; j < dataGridView1.RowCount; j++)
                    {
                        int t = 0;
                        for (int k = chst; k < bb; k++)
                        {
                            dataGridView1.Rows[j - a].Cells[k].Value = dataGridView1.Rows[j].Cells[t].Value;
                            t++;
                        }
                    }
                }
            }

            for (int i = dataGridView1.RowCount - 1; i >= a + 1; i--)
            {
                dataGridView1.Rows.RemoveAt(i);
            }

            for (int i = dataGridView1.ColumnCount - 1; i >= 0; i--)
            {
                if (dataGridView1.Rows[2].Cells[i].Value.ToString().Equals("") &&
                    dataGridView1.Rows[3].Cells[i].Value.ToString().Equals("") &&
                    dataGridView1.Rows[4].Cells[i].Value.ToString().Equals("") &&
                    dataGridView1.Rows[5].Cells[i].Value.ToString().Equals("") &&
                    dataGridView1.Rows[6].Cells[i].Value.ToString().Equals("") &&
                    dataGridView1.Rows[7].Cells[i].Value.ToString().Equals(""))
                {
                    dataGridView1.Columns.RemoveAt(i);
                }
            }

            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                dataGridView1.Columns[i].HeaderText = (i + 1).ToString();
            }

            for (int i = dataGridView1.ColumnCount - 1; i >= 0; i--)
            {

                if (dataGridView1.Columns[i].HeaderText.Equals("31") ||
                    dataGridView1.Columns[i].HeaderText.Equals("30") ||
                    dataGridView1.Columns[i].HeaderText.Equals("28") ||
                    dataGridView1.Columns[i].HeaderText.Equals("20") ||
                    dataGridView1.Columns[i].HeaderText.Equals("18") ||
                    dataGridView1.Columns[i].HeaderText.Equals("17") ||
                    dataGridView1.Columns[i].HeaderText.Equals("15") ||
                    dataGridView1.Columns[i].HeaderText.Equals("14") ||
                    dataGridView1.Columns[i].HeaderText.Equals("13") ||
                    dataGridView1.Columns[i].HeaderText.Equals("12") ||
                    dataGridView1.Columns[i].HeaderText.Equals("10") ||
                    dataGridView1.Columns[i].HeaderText.Equals("7") ||
                    dataGridView1.Columns[i].HeaderText.Equals("4"))
                {
                    dataGridView1.Columns.RemoveAt(i);
                }
            }


            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                dataGridView1.Columns[i].HeaderText = (i + 1).ToString();
            }

            for (int i = 5; i < dataGridView1.RowCount - 1; i++)
                for (int j = 5; j < dataGridView1.RowCount - 1; j++)
                {
                    if (dataGridView1.Rows[i].Cells[0].Value == dataGridView1.Rows[j].Cells[10])
                    {
                        string s = String.Empty;
                        s = dataGridView1.Rows[j].Cells[10].Value.ToString();
                        dataGridView1.Rows[j].Cells[10].Value = "";
                        dataGridView1.Rows[j].Cells[10].Value = dataGridView1.Rows[i].Cells[10].Value;
                        dataGridView1.Rows[i].Cells[10].Value = "";
                        dataGridView1.Rows[i].Cells[10].Value = s;
                    }

                }
            dataGridView1.Rows[2].Cells[2].Value = "";
            dataGridView1.Rows[3].Cells[2].Value = "";
            dataGridView1.Rows[2].Cells[4].Value = "";
            dataGridView1.Rows[3].Cells[4].Value = "";
            dataGridView1.Rows[2].Cells[6].Value = "";
            dataGridView1.Rows[3].Cells[6].Value = "";
            dataGridView1.Rows[2].Cells[8].Value = "";
            dataGridView1.Rows[3].Cells[8].Value = "";

            int ii = 5;
            for (int j = 5; j < dataGridView1.RowCount - 1; j++)
            {
                if (dataGridView1.Rows[ii].Cells[0].Value.ToString() == dataGridView1.Rows[j].Cells[9].Value.ToString())
                {
                    string s1 = String.Empty;
                    string s2 = String.Empty;
                    string s3 = String.Empty;
                    string s4 = String.Empty;
                    string s5 = String.Empty;
                    string s6 = String.Empty;
                    string s7 = String.Empty;
                    string s8 = String.Empty;
                    string s9 = String.Empty;

                    s1 = dataGridView1.Rows[j].Cells[9].Value.ToString();
                    s2 = dataGridView1.Rows[j].Cells[10].Value.ToString();
                    s3 = dataGridView1.Rows[j].Cells[11].Value.ToString();
                    s4 = dataGridView1.Rows[j].Cells[12].Value.ToString();
                    s5 = dataGridView1.Rows[j].Cells[13].Value.ToString();
                    s6 = dataGridView1.Rows[j].Cells[14].Value.ToString();
                    s7 = dataGridView1.Rows[j].Cells[15].Value.ToString();
                    s8 = dataGridView1.Rows[j].Cells[16].Value.ToString();
                    s9 = dataGridView1.Rows[j].Cells[17].Value.ToString();

                    dataGridView1.Rows[j].Cells[9].Value = "";
                    dataGridView1.Rows[j].Cells[10].Value = "";
                    dataGridView1.Rows[j].Cells[11].Value = "";
                    dataGridView1.Rows[j].Cells[12].Value = "";
                    dataGridView1.Rows[j].Cells[13].Value = "";
                    dataGridView1.Rows[j].Cells[14].Value = "";
                    dataGridView1.Rows[j].Cells[15].Value = "";
                    dataGridView1.Rows[j].Cells[16].Value = "";
                    dataGridView1.Rows[j].Cells[17].Value = "";

                    dataGridView1.Rows[j].Cells[9].Value = dataGridView1.Rows[ii].Cells[9].Value;
                    dataGridView1.Rows[j].Cells[10].Value = dataGridView1.Rows[ii].Cells[10].Value;
                    dataGridView1.Rows[j].Cells[11].Value = dataGridView1.Rows[ii].Cells[11].Value;
                    dataGridView1.Rows[j].Cells[12].Value = dataGridView1.Rows[ii].Cells[12].Value;
                    dataGridView1.Rows[j].Cells[13].Value = dataGridView1.Rows[ii].Cells[13].Value;
                    dataGridView1.Rows[j].Cells[14].Value = dataGridView1.Rows[ii].Cells[14].Value;
                    dataGridView1.Rows[j].Cells[15].Value = dataGridView1.Rows[ii].Cells[15].Value;
                    dataGridView1.Rows[j].Cells[16].Value = dataGridView1.Rows[ii].Cells[16].Value;
                    dataGridView1.Rows[j].Cells[17].Value = dataGridView1.Rows[ii].Cells[17].Value;

                    dataGridView1.Rows[ii].Cells[9].Value = "";
                    dataGridView1.Rows[ii].Cells[10].Value = "";
                    dataGridView1.Rows[ii].Cells[11].Value = "";
                    dataGridView1.Rows[ii].Cells[12].Value = "";
                    dataGridView1.Rows[ii].Cells[13].Value = "";
                    dataGridView1.Rows[ii].Cells[14].Value = "";
                    dataGridView1.Rows[ii].Cells[15].Value = "";
                    dataGridView1.Rows[ii].Cells[16].Value = "";
                    dataGridView1.Rows[ii].Cells[17].Value = "";

                    dataGridView1.Rows[ii].Cells[9].Value = s1;
                    dataGridView1.Rows[ii].Cells[10].Value = s2;
                    dataGridView1.Rows[ii].Cells[11].Value = s3;
                    dataGridView1.Rows[ii].Cells[12].Value = s4;
                    dataGridView1.Rows[ii].Cells[13].Value = s5;
                    dataGridView1.Rows[ii].Cells[14].Value = s6;
                    dataGridView1.Rows[ii].Cells[15].Value = s7;
                    dataGridView1.Rows[ii].Cells[16].Value = s8;
                    dataGridView1.Rows[ii].Cells[17].Value = s9;

                    ii++;
                    j = 5;


                }

            }

            for (int i = 4; i < dataGridView1.RowCount - 1; i++)
            {
                if ((dataGridView1.Rows[i].Cells[16].Value.ToString() == "") || (dataGridView1.Rows[i].Cells[16].Value == null) || (dataGridView1.Rows[i].Cells[16].Value.ToString().Length == 0))
                {
                    dataGridView1.Rows[i].Cells[16].Value = "0";
                }
            }

            for (int i = 4; i < dataGridView1.RowCount - 1; i++)
            {
                dataGridView1.Rows[i].Cells[16].Value = (Convert.ToInt32(dataGridView1.Rows[i].Cells[16].Value) + Convert.ToInt32(dataGridView1.Rows[i].Cells[17].Value));
            }

            dataGridView1.Columns.RemoveAt(dataGridView1.ColumnCount - 1);

            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[0].Value = "";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[0].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[0].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[1].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[1].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[2].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[2].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[3].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[3].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[4].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[4].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[5].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[5].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[6].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[6].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[7].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[7].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[8].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[8].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[9].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[9].Value;

            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[0].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[1].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[2].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[3].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[4].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[5].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[6].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[7].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[8].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[9].Value = "-";

            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[10].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[11].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[12].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[13].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[14].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[15].Value = "-";

            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[0].Value = "Транзит";

            dataGridView1.Columns.RemoveAt(9);

            dataGridView1.Rows[2].Cells[0].Value = "Подразделение";

            for (int j = 0; j < dataGridView1.ColumnCount; j++)
            {
                dataGridView1.Columns[j].HeaderText = (j + 1).ToString();
            }

            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[9].Value = dataGridView1.Rows[4].Cells[9].Value;
            dataGridView1.Rows[4].Cells[9].Value = "";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[10].Value = dataGridView1.Rows[4].Cells[10].Value;
            dataGridView1.Rows[4].Cells[10].Value = "";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[11].Value = dataGridView1.Rows[4].Cells[11].Value;
            dataGridView1.Rows[4].Cells[11].Value = "";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[12].Value = dataGridView1.Rows[4].Cells[12].Value;
            dataGridView1.Rows[4].Cells[12].Value = "";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[13].Value = dataGridView1.Rows[4].Cells[13].Value;
            dataGridView1.Rows[4].Cells[13].Value = "";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[14].Value = dataGridView1.Rows[4].Cells[14].Value;
            dataGridView1.Rows[4].Cells[14].Value = "";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[15].Value = dataGridView1.Rows[4].Cells[15].Value;
            dataGridView1.Rows[4].Cells[15].Value = "";

            string q1, q3, q4;
            q1 = dataGridView1.Rows[4].Cells[7].Value.ToString();
            dataGridView1.Rows[2].Cells[7].Value = "Вагоны";
            dataGridView1.Rows[4].Cells[7].Value = "Кол-во";
            dataGridView1.Rows[4].Cells[8].Value = "";

            dataGridView1.Rows[3].Cells[8].Value = dataGridView1.Rows[3].Cells[7].Value;
            dataGridView1.Rows[2].Cells[8].Value = "Денежные средства";
            dataGridView1.Rows[4].Cells[8].Value = "Тысяч руб.";

            dataGridView1.Rows[2].Cells[9].Value = "План на период:";

            q3 = dataGridView1.Rows[3].Cells[10].Value.ToString();
            dataGridView1.Rows[3].Cells[10].Value = "";
            dataGridView1.Rows[3].Cells[10].Value = "План на: " + q3;

            q4 = dataGridView1.Rows[3].Cells[11].Value.ToString();
            dataGridView1.Rows[3].Cells[11].Value = "";
            dataGridView1.Rows[3].Cells[11].Value = "Отгрузка на: " + q4;

            dataGridView1.Rows[2].Cells[15].Value = "";
            dataGridView1.Rows[2].Cells[15].Value = "Остаток на текущую дату";
            dataGridView1.Rows[1].Cells[10].Value = "";
            dataGridView1.Rows[2].Cells[13].Value = "";
            dataGridView1.Rows[2].Cells[14].Value = "";
            dataGridView1.Rows[3].Cells[13].Value = "";

            dataGridView1.Rows[3].Cells[13].Value = "На дату: " + dataGridView1.Rows[3].Cells[12].Value;
            dataGridView1.Rows[2].Cells[13].Value = "Кол-во готового лома";

            dataGridView1.Rows[4].Cells[9].Value = "Кол-во, тонн";
            dataGridView1.Rows[4].Cells[11].Value = "Кол-во, тонн";
            dataGridView1.Rows[4].Cells[13].Value = "Кол-во, тонн";

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                dataGridView1.Rows[i].Cells[14].Value = "";
            }

            for (int i = 5; i < dataGridView1.RowCount; i++)
            {
                dataGridView1.Rows[i].Cells[13].Value = "";
                dataGridView1.Rows[i].Cells[8].Value = "";
                dataGridView1.Rows[i].Cells[7].Value = "";
            }
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[7].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[8].Value = "-";


            for (int i = 5; i < got_prod.Length + 5; i++)
            {
                dataGridView1.Rows[i].Cells[13].Value = got_prod[i - 5];
            }

            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[13].Value = got_prod.Sum();

            dataGridView3.Rows.RemoveAt(0);
            dataGridView3.Columns.RemoveAt(0);

            for (int i = 0; i < 24; i++)
            {
                dataGridView1.Rows[i + 5].Cells[7].Value = dataGridView3.Rows[i].Cells[0].Value;
                dataGridView1.Rows[i + 5].Cells[8].Value = dataGridView3.Rows[i].Cells[1].Value;
            }

           

            
            

            #endregion














            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                if (dataGridView2.Rows[i].Cells[1].Value.ToString().Contains("Подразделение"))
                {
                    dataGridView2.Rows[i + 1].Cells[1].Value = "XX";
                    dataGridView2.Rows[i + 2].Cells[1].Value = "XX";
                    break;
                }
            }


            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                if (dataGridView2.Rows[i].Cells[1].Value.ToString().Equals(""))
                {
                    dataGridView2.Rows.RemoveAt(i);
                    i--;
                }
            }
            dataGridView2.Columns.RemoveAt(0);
            chst = dataGridView2.ColumnCount;

            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                if (
                    (!(dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Бр. Коростелевых ул., 52, терр. з-да \"Гидропресс\"") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Бугуруслан г., Восточное ш. 1") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Бузулук г., ул. Промышленная, 6") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Донгузская ул, 20") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Илек с., ул. Шоссейная, 54Б") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Краснохолм с., ул. Шоссейная 1А") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Кувандык г., ул. Дзержинского, 1А") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Курманаевка п.") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Медногорск г., ул. Комсомольская, 40") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Новосергиевский п.") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Октябрьское, ул. Транспортная, 1") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Орск, ул.Строителей, 44") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Переволоцкий п., ул. Ленинская, 2А") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Подольск п., Промышленная ул., 3") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Сакмарский р-н, Сакмарская ст., терр. СМП-639") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Саракташ п., ул. Производственная, 4") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Соль-Илецк г., ул. Гонтаренко, 1А") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Сорочинск г., ул. Пролетарская, 3") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Терешковой  ул., 287 д.") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Тоцкое с. Тоцкий р-н, Автомобилистов ул., 1Е") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Центральная ул., 1 д.") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Шильда, Топсклад") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("ИТОГО:") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Транзит") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Contains("Отчет") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Contains("Подразделение") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Contains("Из расчета") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("XX"))))
                {
                    dataGridView2.Rows.RemoveAt(i);
                    i--;
                }
            }

            for (int i = dataGridView2.ColumnCount - 1; i >= 0; i--)
            {
                if (!(i == 8 || i == 7))
                {
                    dataGridView2.Columns.RemoveAt(i);
                }
            }

            DataGridViewTextBoxColumn[] cccolumn = new DataGridViewTextBoxColumn[2];

            for (int i = 0; i < 2; i++)
            {
                cccolumn[i] = new DataGridViewTextBoxColumn();
            }
            dataGridView1.Columns.AddRange(cccolumn);

            for (int i = dataGridView1.RowCount - 1; i >= 0; i--)
                for (int j = dataGridView1.ColumnCount - 3; j >= 5; j--)
                {
                    dataGridView1.Rows[i].Cells[j + 2].Value = dataGridView1.Rows[i].Cells[j].Value;
                    dataGridView1.Rows[i].Cells[j].Value = "";
                }


            for (int i = 0; i < dataGridView2.RowCount; i++)
                for (int j = 0; j < dataGridView2.ColumnCount; j++)
                {
                    dataGridView1.Rows[i].Cells[j + 5].Value =
                        dataGridView2.Rows[i].Cells[j].Value;
                }


            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[5].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[5].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[5].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[6].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[6].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[6].Value = "-";

            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[dataGridView1.ColumnCount - 1].Value = "-";

            dataGridView1.Rows[2].Cells[5].Value = "Заготовлено за дату:";

            dataGridView1.Rows[3].Cells[12].Value = dataGridView1.Rows[3].Cells[12].Value.ToString().Substring(9, 8);
            dataGridView1.Rows[4].Cells[2].Value = "Ср. цена";
            dataGridView1.Rows[0].Cells[11].Value = "Отчет по отгрузке " + dataGridView1.Rows[0].Cells[0].Value.ToString().Substring(19, (dataGridView1.Rows[0].Cells[0].Value.ToString().Length - 20));
            dataGridView1.Rows[3].Cells[13].Value = dataGridView1.Rows[3].Cells[13].Value.ToString().Substring(13, 19);





            for (int i = 10; i < 16; i++)
            {
                if ((dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[i].Value.ToString() == "") || (dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[i].Value == null) || (dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[i].Value.ToString().Length == 0))
                {
                    dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[i].Value = "-";
                }
            }




            
          /*  if (dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[dataGridView1.ColumnCount - 1].Value == null)
            {
                dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[dataGridView1.ColumnCount - 1].Value = 0;
            }

            MessageBox.Show(dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
            MessageBox.Show(nehvatka.ToString());
            int oj1 = Convert.ToInt32(dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[dataGridView1.ColumnCount - 1].Value = "";
            int oj2 = oj1 + nehvatka;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[dataGridView1.ColumnCount - 1].Value = oj2.ToString();
            */



            //----------------------------------------------------------------------------------- Download to Excel from dataGrid






            EXX.Application exApp = new EXX.Application();

            exApp.Visible = true;
            exApp.Workbooks.Add();

            Worksheet workSheet = (Worksheet)exApp.ActiveSheet;

            workSheet.PageSetup.LeftMargin = exApp.Application.InchesToPoints(0.1);
            workSheet.PageSetup.RightMargin = exApp.Application.InchesToPoints(0.1);
            workSheet.PageSetup.TopMargin = exApp.Application.InchesToPoints(0.1);
            workSheet.PageSetup.BottomMargin = exApp.Application.InchesToPoints(0.1);


            int rowExcel = 1;

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                workSheet.Cells[rowExcel, "A"] = dataGridView1.Rows[i].Cells[0].Value;
                workSheet.Cells[rowExcel, "B"] = dataGridView1.Rows[i].Cells[1].Value;
                workSheet.Cells[rowExcel, "C"] = dataGridView1.Rows[i].Cells[2].Value;
                workSheet.Cells[rowExcel, "D"] = dataGridView1.Rows[i].Cells[3].Value;
                workSheet.Cells[rowExcel, "E"] = dataGridView1.Rows[i].Cells[4].Value;
                workSheet.Cells[rowExcel, "F"] = dataGridView1.Rows[i].Cells[5].Value;
                workSheet.Cells[rowExcel, "G"] = dataGridView1.Rows[i].Cells[6].Value;
                workSheet.Cells[rowExcel, "H"] = dataGridView1.Rows[i].Cells[7].Value;
                workSheet.Cells[rowExcel, "I"] = dataGridView1.Rows[i].Cells[8].Value;
                workSheet.Cells[rowExcel, "J"] = dataGridView1.Rows[i].Cells[9].Value;
                workSheet.Cells[rowExcel, "K"] = dataGridView1.Rows[i].Cells[10].Value;
                workSheet.Cells[rowExcel, "L"] = dataGridView1.Rows[i].Cells[11].Value;
                workSheet.Cells[rowExcel, "M"] = dataGridView1.Rows[i].Cells[12].Value;
                workSheet.Cells[rowExcel, "N"] = dataGridView1.Rows[i].Cells[13].Value;
                workSheet.Cells[rowExcel, "O"] = dataGridView1.Rows[i].Cells[14].Value;
                workSheet.Cells[rowExcel, "P"] = dataGridView1.Rows[i].Cells[15].Value;
                workSheet.Cells[rowExcel, "Q"] = dataGridView1.Rows[i].Cells[16].Value;
                workSheet.Cells[rowExcel, "R"] = dataGridView1.Rows[i].Cells[17].Value;
                ++rowExcel;
            }

            // Переименование названий площадок
            workSheet.Cells[6, 1] = "Гидропресс";
            workSheet.Cells[7, 1] = "Бугуруслан";
            workSheet.Cells[8, 1] = "Бузулук";
            workSheet.Cells[9, 1] = "Донгузская";
            workSheet.Cells[10, 1] = "Илек";
            workSheet.Cells[11, 1] = "Краснохолм";
            workSheet.Cells[12, 1] = "Кувандык";
            workSheet.Cells[13, 1] = "Курманаевка";
            workSheet.Cells[14, 1] = "Медногорск";
            workSheet.Cells[15, 1] = "Новосергиевка";
            workSheet.Cells[16, 1] = "Октябрьское"; 
            workSheet.Cells[17, 1] = "Орск";
            workSheet.Cells[18, 1] = "Переволоцк";
            workSheet.Cells[19, 1] = "Подольск";
            workSheet.Cells[20, 1] = "Сакмара";
            workSheet.Cells[21, 1] = "Саракташ";
            workSheet.Cells[22, 1] = "Соль-Илецк";
            workSheet.Cells[23, 1] = "Сорочинск";
            workSheet.Cells[24, 1] = "Терешковой";
            workSheet.Cells[25, 1] = "Тоцкое";
            workSheet.Cells[26, 1] = "Центральная";
            workSheet.Cells[27, 1] = "Шильда";


           // Обработка отчётов + ВАГОНЫ+ДЕНЬГИ
            workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[1, 11]].Merge(Type.Missing);
            workSheet.Range[workSheet.Cells[2, 1], workSheet.Cells[2, 18]].Merge(Type.Missing);

            workSheet.Range[workSheet.Cells[3, 2], workSheet.Cells[3, 3]].Merge(Type.Missing);
            workSheet.Range[workSheet.Cells[4, 2], workSheet.Cells[4, 3]].Merge(Type.Missing);

            workSheet.Range[workSheet.Cells[3, 4], workSheet.Cells[3, 5]].Merge(Type.Missing);
            workSheet.Range[workSheet.Cells[4, 4], workSheet.Cells[4, 5]].Merge(Type.Missing);

            workSheet.Range[workSheet.Cells[3, 6], workSheet.Cells[3, 7]].Merge(Type.Missing);
            workSheet.Range[workSheet.Cells[4, 6], workSheet.Cells[4, 7]].Merge(Type.Missing);

            workSheet.Range[workSheet.Cells[3, 8], workSheet.Cells[3, 9]].Merge(Type.Missing);
            workSheet.Range[workSheet.Cells[4, 8], workSheet.Cells[4, 9]].Merge(Type.Missing);

            // workSheet.Range[workSheet.Cells[3, 10], workSheet.Cells[3, 11]].Merge(Type.Missing);
            //workSheet.Range[workSheet.Cells[4, 10], workSheet.Cells[4, 11]].Merge(Type.Missing);

            workSheet.Range[workSheet.Cells[1, 12], workSheet.Cells[1, 18]].Merge(Type.Missing); //otgruzka

            //  workSheet.Range[workSheet.Cells[2, 11], workSheet.Cells[2, 15]].Merge(Type.Missing);
            workSheet.Range[workSheet.Cells[3, 12], workSheet.Cells[3, 13]].Merge(Type.Missing);
            workSheet.Range[workSheet.Cells[3, 14], workSheet.Cells[3, 15]].Merge(Type.Missing);
            workSheet.Range[workSheet.Cells[3, 16], workSheet.Cells[3, 17]].Merge(Type.Missing);
            //workSheet.Range[workSheet.Cells[3, 18], workSheet.Cells[4, 18]].Merge(Type.Missing);

            // workSheet.Range[workSheet.Cells[1, 16], workSheet.Cells[2, 17]].Merge(Type.Missing);
            //workSheet.Range[workSheet.Cells[3, 16], workSheet.Cells[4, 17]].Merge(Type.Missing);



            for (int i = 0; i < dataGridView1.RowCount; i++)
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    (workSheet.Cells[i + 1, j + 1] as Range).Font.Size = 9;
                    (workSheet.Cells[i + 1, j + 1] as Range).Font.Name = "Arial";
                }

            workSheet.Rows.RowHeight = 19.50;
            workSheet.Rows.VerticalAlignment = XlHAlign.xlHAlignDistributed;

            workSheet.Columns.ColumnWidth = 15.14;

            workSheet.Range[workSheet.Cells[1, 2], workSheet.Cells[10, dataGridView1.RowCount]].ColumnWidth = 6.86;
            workSheet.Columns.HorizontalAlignment = XlHAlign.xlHAlignDistributed;
            workSheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;

            workSheet.Cells.Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlDot;
            workSheet.Cells.Borders[XlBordersIndex.xlInsideVertical].Weight = XlBorderWeight.xlThin;// внутренние вертикальные
            workSheet.Cells.Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlDot;
            workSheet.Cells.Borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlThin;// внутренние горизонтальные
            workSheet.Cells.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlDot;
            workSheet.Cells.Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;// верхняя внешняя
            workSheet.Cells.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlDot;
            workSheet.Cells.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;// правая внешняя
            workSheet.Cells.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlDot;
            workSheet.Cells.Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;// левая внешняя
            workSheet.Cells.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDot;
            workSheet.Cells.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;// нижняя внешняя

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                (workSheet.Cells[i + 1, 12] as Range).Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlDouble;
                (workSheet.Cells[i + 1, 12] as Range).Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThick;
            }

            for (int i = 6; i < dataGridView1.RowCount; i++)
            {
                (workSheet.Cells[i, 5] as Range).Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                (workSheet.Cells[i, 5] as Range).Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;
                (workSheet.Cells[i, 5] as Range).Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                (workSheet.Cells[i, 5] as Range).Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
                (workSheet.Cells[i, 5] as Range).Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                (workSheet.Cells[i, 5] as Range).Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;

                (workSheet.Cells[i, 6] as Range).Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                (workSheet.Cells[i, 6] as Range).Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
                (workSheet.Cells[i, 6] as Range).Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                (workSheet.Cells[i, 6] as Range).Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;
                (workSheet.Cells[i, 6] as Range).Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                (workSheet.Cells[i, 6] as Range).Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;

                (workSheet.Cells[i, 8] as Range).Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                (workSheet.Cells[i, 8] as Range).Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
                (workSheet.Cells[i, 8] as Range).Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                (workSheet.Cells[i, 8] as Range).Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;
                (workSheet.Cells[i, 8] as Range).Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                (workSheet.Cells[i, 8] as Range).Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
            }

            (workSheet.Cells[6, 5] as Range).Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            (workSheet.Cells[6, 5] as Range).Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;
            (workSheet.Cells[6, 6] as Range).Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            (workSheet.Cells[6, 6] as Range).Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;
            (workSheet.Cells[6, 8] as Range).Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            (workSheet.Cells[6, 8] as Range).Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;

            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                (workSheet.Cells[dataGridView1.RowCount, i + 1] as Range).Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                (workSheet.Cells[dataGridView1.RowCount, i + 1] as Range).Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;
                (workSheet.Cells[dataGridView1.RowCount, i + 1] as Range).Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                (workSheet.Cells[dataGridView1.RowCount, i + 1] as Range).Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;

                (workSheet.Cells[dataGridView1.RowCount, i + 1] as Range).Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                (workSheet.Cells[dataGridView1.RowCount, i + 1] as Range).Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
                (workSheet.Cells[dataGridView1.RowCount, i + 1] as Range).Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                (workSheet.Cells[dataGridView1.RowCount, i + 1] as Range).Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;
            }

            workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[dataGridView1.RowCount, dataGridView1.ColumnCount]].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            /*for (int i = 5; i < dataGridView1.RowCount; i++)
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    if (i%2 == 1)
                    {
                        (workSheet.Cells[i + 1, j + 1] as Range).Interior.ColorIndex = 19;
                    }
                }*/

            for (int i = 2; i < 4; i++)
                for (int j = 1; j < dataGridView1.ColumnCount; j++)
                {
                    (workSheet.Cells[i + 1, j + 1] as Range).Font.Size = 8;
                    (workSheet.Cells[i + 1, j + 1] as Range).Font.Name = "Arial";
                }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                (workSheet.Cells[i + 1, 12] as Range).ColumnWidth = 5.86;
                (workSheet.Cells[i + 1, 14] as Range).ColumnWidth = 6.71;

                (workSheet.Cells[i + 1, 13] as Range).ColumnWidth = 7.71;
                (workSheet.Cells[i + 1, 15] as Range).ColumnWidth = 7.71;

                (workSheet.Cells[i + 1, 16] as Range).ColumnWidth = 6.00;
                (workSheet.Cells[i + 1, 17] as Range).ColumnWidth = 5.43;
            }

            for (int i = 3; i < dataGridView1.RowCount; i++)
            {
                workSheet.Range[workSheet.Cells[i+1, 16], workSheet.Cells[i+1, 17]].Merge(Type.Missing);
            }

            workSheet.Range[workSheet.Cells[5, 14], workSheet.Cells[5, 15]].Merge(Type.Missing);
            workSheet.Range[workSheet.Cells[5, 12], workSheet.Cells[5, 13]].Merge(Type.Missing);
            workSheet.Range[workSheet.Cells[3, 18], workSheet.Cells[5, 18]].Merge(Type.Missing);

            for (int i = 2; i < dataGridView1.RowCount-1; i++)
            {
                (workSheet.Cells[i + 1, 10] as Range).Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlDouble;
                (workSheet.Cells[i + 1, 10] as Range).Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThick;
            }


            workSheet.PrintOut(1, 1, n, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);


            
            button1.Enabled = true;
            button2.Enabled = false;
           


        }

        private void button3_Click(object sender, EventArgs e)
        {
            EXX.Application exApp = new EXX.Application();

            exApp.Visible = true;
            exApp.Workbooks.Add();

            Worksheet workSheet = (Worksheet)exApp.ActiveSheet;

            workSheet.PageSetup.LeftMargin = exApp.Application.InchesToPoints(0.1);
            workSheet.PageSetup.RightMargin = exApp.Application.InchesToPoints(0.1);
            workSheet.PageSetup.TopMargin = exApp.Application.InchesToPoints(0.1);
            workSheet.PageSetup.BottomMargin = exApp.Application.InchesToPoints(0.1);


            int rowExcel = 1;

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                workSheet.Cells[rowExcel, "A"] = dataGridView1.Rows[i].Cells[0].Value;
                workSheet.Cells[rowExcel, "B"] = dataGridView1.Rows[i].Cells[1].Value;
                workSheet.Cells[rowExcel, "C"] = dataGridView1.Rows[i].Cells[2].Value;
                workSheet.Cells[rowExcel, "D"] = dataGridView1.Rows[i].Cells[3].Value;
                workSheet.Cells[rowExcel, "E"] = dataGridView1.Rows[i].Cells[4].Value;
                workSheet.Cells[rowExcel, "F"] = dataGridView1.Rows[i].Cells[5].Value;
                workSheet.Cells[rowExcel, "G"] = dataGridView1.Rows[i].Cells[6].Value;
                workSheet.Cells[rowExcel, "H"] = dataGridView1.Rows[i].Cells[7].Value;
                workSheet.Cells[rowExcel, "I"] = dataGridView1.Rows[i].Cells[8].Value;
                workSheet.Cells[rowExcel, "J"] = dataGridView1.Rows[i].Cells[9].Value;
                workSheet.Cells[rowExcel, "K"] = dataGridView1.Rows[i].Cells[10].Value;
                workSheet.Cells[rowExcel, "L"] = dataGridView1.Rows[i].Cells[11].Value;
                workSheet.Cells[rowExcel, "M"] = dataGridView1.Rows[i].Cells[12].Value;
                workSheet.Cells[rowExcel, "N"] = dataGridView1.Rows[i].Cells[13].Value;
                workSheet.Cells[rowExcel, "O"] = dataGridView1.Rows[i].Cells[14].Value;
                workSheet.Cells[rowExcel, "P"] = dataGridView1.Rows[i].Cells[15].Value;
                workSheet.Cells[rowExcel, "Q"] = dataGridView1.Rows[i].Cells[16].Value;
                workSheet.Cells[rowExcel, "R"] = dataGridView1.Rows[i].Cells[17].Value;
               ++rowExcel;
            }

            workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[1, 11]].Merge(Type.Missing);
            workSheet.Range[workSheet.Cells[2, 1], workSheet.Cells[2, 18]].Merge(Type.Missing);

            workSheet.Range[workSheet.Cells[3, 2], workSheet.Cells[3, 3]].Merge(Type.Missing);
            workSheet.Range[workSheet.Cells[4, 2], workSheet.Cells[4, 3]].Merge(Type.Missing);

            workSheet.Range[workSheet.Cells[3, 4], workSheet.Cells[3, 5]].Merge(Type.Missing);
            workSheet.Range[workSheet.Cells[4, 4], workSheet.Cells[4, 5]].Merge(Type.Missing);

            workSheet.Range[workSheet.Cells[3, 6], workSheet.Cells[3, 7]].Merge(Type.Missing);
            workSheet.Range[workSheet.Cells[4, 6], workSheet.Cells[4, 7]].Merge(Type.Missing);

            workSheet.Range[workSheet.Cells[3, 8], workSheet.Cells[3, 9]].Merge(Type.Missing);
            workSheet.Range[workSheet.Cells[4, 8], workSheet.Cells[4, 9]].Merge(Type.Missing);

           // workSheet.Range[workSheet.Cells[3, 10], workSheet.Cells[3, 11]].Merge(Type.Missing);
            //workSheet.Range[workSheet.Cells[4, 10], workSheet.Cells[4, 11]].Merge(Type.Missing);
           
            workSheet.Range[workSheet.Cells[1, 12], workSheet.Cells[1, 18]].Merge(Type.Missing); //otgruzka

          //  workSheet.Range[workSheet.Cells[2, 11], workSheet.Cells[2, 15]].Merge(Type.Missing);
            workSheet.Range[workSheet.Cells[3, 12], workSheet.Cells[3, 13]].Merge(Type.Missing);
            workSheet.Range[workSheet.Cells[3, 14], workSheet.Cells[3, 15]].Merge(Type.Missing);
            workSheet.Range[workSheet.Cells[3, 16], workSheet.Cells[3, 17]].Merge(Type.Missing);
            workSheet.Range[workSheet.Cells[3, 18], workSheet.Cells[4, 18]].Merge(Type.Missing);
            
           // workSheet.Range[workSheet.Cells[1, 16], workSheet.Cells[2, 17]].Merge(Type.Missing);
            //workSheet.Range[workSheet.Cells[3, 16], workSheet.Cells[4, 17]].Merge(Type.Missing);
            
            

            for (int i = 0; i < dataGridView1.RowCount; i++)
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    (workSheet.Cells[i+1, j+1] as Range).Font.Size = 9;
                    (workSheet.Cells[i+1, j+1] as Range).Font.Name = "Arial";
                }

            workSheet.Rows.RowHeight = 20.25;
            workSheet.Rows.VerticalAlignment = XlHAlign.xlHAlignDistributed;
            
            workSheet.Columns.ColumnWidth = 15.14;

            workSheet.Range[workSheet.Cells[1, 2], workSheet.Cells[10, dataGridView1.RowCount]].ColumnWidth = 6.86;            
            workSheet.Columns.HorizontalAlignment = XlHAlign.xlHAlignDistributed;
            workSheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;

            workSheet.Cells.Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlDot;
            workSheet.Cells.Borders[XlBordersIndex.xlInsideVertical].Weight = XlBorderWeight.xlThin;// внутренние вертикальные
            workSheet.Cells.Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlDot;
            workSheet.Cells.Borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlThin;// внутренние горизонтальные
            workSheet.Cells.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlDot;
            workSheet.Cells.Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;// верхняя внешняя
            workSheet.Cells.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlDot;
            workSheet.Cells.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;// правая внешняя
            workSheet.Cells.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlDot;
            workSheet.Cells.Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;// левая внешняя
            workSheet.Cells.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDot;
            workSheet.Cells.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;// нижняя внешняя

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                (workSheet.Cells[i + 1, 12] as Range).Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlDouble;
                (workSheet.Cells[i + 1, 12] as Range).Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThick;
           }

            for (int i = 6; i < dataGridView1.RowCount; i++)
            {
                (workSheet.Cells[i, 5] as Range).Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                (workSheet.Cells[i, 5] as Range).Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;
                (workSheet.Cells[i, 5] as Range).Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                (workSheet.Cells[i, 5] as Range).Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
                (workSheet.Cells[i, 5] as Range).Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                (workSheet.Cells[i, 5] as Range).Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
                
                (workSheet.Cells[i, 6] as Range).Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                (workSheet.Cells[i, 6] as Range).Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
                (workSheet.Cells[i, 6] as Range).Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                (workSheet.Cells[i, 6] as Range).Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;
                (workSheet.Cells[i, 6] as Range).Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                (workSheet.Cells[i, 6] as Range).Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;

                (workSheet.Cells[i, 8] as Range).Borders[XlBordersIndex.xlEdgeRight].LineStyle =XlLineStyle.xlContinuous;
                (workSheet.Cells[i, 8] as Range).Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
                (workSheet.Cells[i, 8] as Range).Borders[XlBordersIndex.xlEdgeLeft].LineStyle =XlLineStyle.xlContinuous;
                (workSheet.Cells[i, 8] as Range).Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;
                (workSheet.Cells[i, 8] as Range).Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                (workSheet.Cells[i, 8] as Range).Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
            }

            (workSheet.Cells[6, 5] as Range).Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            (workSheet.Cells[6, 5] as Range).Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;
            (workSheet.Cells[6, 6] as Range).Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            (workSheet.Cells[6, 6] as Range).Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;
            (workSheet.Cells[6, 8] as Range).Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            (workSheet.Cells[6, 8] as Range).Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;

            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                (workSheet.Cells[dataGridView1.RowCount, i + 1] as Range).Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                (workSheet.Cells[dataGridView1.RowCount, i + 1] as Range).Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;
                (workSheet.Cells[dataGridView1.RowCount, i + 1] as Range).Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                (workSheet.Cells[dataGridView1.RowCount, i + 1] as Range).Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
                
                (workSheet.Cells[dataGridView1.RowCount, i + 1] as Range).Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                (workSheet.Cells[dataGridView1.RowCount, i + 1] as Range).Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
                (workSheet.Cells[dataGridView1.RowCount, i + 1] as Range).Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                (workSheet.Cells[dataGridView1.RowCount, i + 1] as Range).Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;
            }
            
            workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[dataGridView1.RowCount, dataGridView1.ColumnCount]].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            /*for (int i = 5; i < dataGridView1.RowCount; i++)
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    if (i%2 == 1)
                    {
                        (workSheet.Cells[i + 1, j + 1] as Range).Interior.ColorIndex = 19;
                    }
                }*/

            for (int i = 2; i < 4; i++)
                for (int j = 1; j < dataGridView1.ColumnCount; j++)
                {
                    (workSheet.Cells[i + 1, j + 1] as Range).Font.Size = 8;
                    (workSheet.Cells[i + 1, j + 1] as Range).Font.Name = "Arial";
                }

            for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    (workSheet.Cells[i + 1, 12] as Range).ColumnWidth = 5.86;
                    (workSheet.Cells[i + 1, 14] as Range).ColumnWidth = 6.71;

                    (workSheet.Cells[i + 1, 13] as Range).ColumnWidth = 7.71;
                    (workSheet.Cells[i + 1, 15] as Range).ColumnWidth = 7.71;

                    (workSheet.Cells[i + 1, 16] as Range).ColumnWidth = 6.00;
                    (workSheet.Cells[i + 1, 17] as Range).ColumnWidth = 6.00;


                }

            (workSheet.Cells[dataGridView1.RowCount, dataGridView1.ColumnCount] as Range).PrintOut(1, 1, 1, false, "", false, false, true);


        }

        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Файл за ДЕНЬ|*.xls; *.xlsx";
            openDialog.ShowDialog();

            try
            {
                ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                //Книга.
                ObjWorkBook = ObjExcel.Workbooks.Open(openDialog.FileName);
                //Таблица.
                ObjWorkSheet = ObjExcel.ActiveSheet as Worksheet;
                Range rg = null;
                int LastRowNumber = ObjWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
                int LastColNumber = ObjWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Column;
                dataGridView2.ColumnCount = LastColNumber;
                dataGridView2.RowCount = LastRowNumber;
                // MessageBox.Show(LastRowNumber.ToString());
                Int32 row = 1;
                dataGridView2.Rows.Clear();
                List<String> arr = new List<string>();
                while (row != LastRowNumber + 1) //ObjWorkSheet.get_Range("a" + row, "a" + row).Value != null)
                {
                    // Читаем данные из ячейки
                    rg = ObjWorkSheet.get_Range("a" + row, "bb" + row);
                    foreach (Range item in rg)
                    {
                        try
                        {
                            arr.Add(item.Value.ToString().Trim());
                        }
                        catch
                        {
                            arr.Add("");
                        }
                    }
                    dataGridView2.Rows.Add(arr[0], arr[1], arr[2], arr[3], arr[4], arr[5], arr[6], arr[7], arr[8],
                        arr[9], arr[10], arr[11], arr[12], arr[13], arr[14], arr[15], arr[16], arr[17], arr[18], arr[19],
                        arr[20], arr[21], arr[22], arr[23], arr[24], arr[25], arr[26], arr[27], arr[28], arr[29], arr[30]);
                    arr.Clear();
                    row++;
                }

                MessageBox.Show("Файл успешно считан!", "Считывания excel файла", MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message, "Ошибка при считывании excel файла", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                ObjWorkBook.Close(false, "", null);
                // Закрытие приложения Excel.
                ObjExcel.Quit();
                ObjWorkBook = null;
                ObjWorkSheet = null;
                ObjExcel = null;
                GC.Collect();
            }

          //  button4.Enabled = false;
            button4.Enabled = false;
            button6.Enabled = true;

        }

        private void button5_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                if (dataGridView2.Rows[i].Cells[1].Value.ToString().Contains("Подразделение"))
                {
                    dataGridView2.Rows[i + 1].Cells[1].Value = "XX";
                    dataGridView2.Rows[i + 2].Cells[1].Value = "XX";
                    break;
                }
            }


            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                if (dataGridView2.Rows[i].Cells[1].Value.ToString().Equals(""))
                {
                    dataGridView2.Rows.RemoveAt(i);
                    i--;
                }
            }
            dataGridView2.Columns.RemoveAt(0);
            chst = dataGridView2.ColumnCount;

            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                if (
                    (!(dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Бр. Коростелевых ул., 52, терр. з-да \"Гидропресс\"") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Бугуруслан г., Восточное ш. 1") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Бузулук г., ул. Промышленная, 6") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Донгузская ул, 20") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Илек с., ул. Шоссейная, 54Б") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Краснохолм с., ул. Шоссейная 1А") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Кувандык г., ул. Дзержинского, 1А") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Курманаевка п.") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Медногорск г., ул. Комсомольская, 40") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Новосергиевский п.") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Октябрьское, ул. Транспортная, 1") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Орск, ул.Строителей, 44") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Переволоцкий п., ул. Ленинская, 2А") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Подольск п., Промышленная ул., 3") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Сакмарский р-н, Сакмарская ст., терр. СМП-639") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Саракташ п., ул. Производственная, 4") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Соль-Илецк г., ул. Гонтаренко, 1А") ||                  
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Сорочинск г., ул. Пролетарская, 3") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Терешковой  ул., 287 д.") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Тоцкое с. Тоцкий р-н, Автомобилистов ул., 1Е") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Центральная ул., 1 д.") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Шильда, Топсклад") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("ИТОГО:") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("Транзит") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Contains("Отчет") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Contains("Подразделение") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Contains("Из расчета") ||
                       dataGridView2.Rows[i].Cells[0].Value.ToString().Equals("XX"))))
                {
                    dataGridView2.Rows.RemoveAt(i);
                    i--;
                }
            }

            for (int i = dataGridView2.ColumnCount - 1; i >= 0; i--)
            {
                if (!(i == 8 || i == 7))
                {
                    dataGridView2.Columns.RemoveAt(i);
                }
            }

            DataGridViewTextBoxColumn[] cccolumn = new DataGridViewTextBoxColumn[2];

            for (int i = 0; i < 2; i++)
            {
                cccolumn[i] = new DataGridViewTextBoxColumn();
            }
            dataGridView1.Columns.AddRange(cccolumn);

            for (int i = dataGridView1.RowCount - 1; i >= 0; i--)
                for (int j = dataGridView1.ColumnCount - 3; j >= 5;  j--)
                {
                    dataGridView1.Rows[i].Cells[j+2].Value = dataGridView1.Rows[i].Cells[j].Value;
                    dataGridView1.Rows[i].Cells[j].Value = "";
                }
            

            for (int i = 0; i < dataGridView2.RowCount; i++)
                for (int j = 0; j < dataGridView2.ColumnCount; j++)
                {
                    dataGridView1.Rows[i].Cells[j+5].Value =
                        dataGridView2.Rows[i].Cells[j].Value;
                }

            
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[5].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[5].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[5].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[6].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[6].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[6].Value = "-";

            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[dataGridView1.ColumnCount - 1].Value = "-";
          
            dataGridView1.Rows[2].Cells[5].Value = "Заготовлено за дату:";

            dataGridView1.Rows[3].Cells[12].Value = dataGridView1.Rows[3].Cells[12].Value.ToString().Substring(9, 8);
           dataGridView1.Rows[4].Cells[2].Value = "Ср. цена";
           dataGridView1.Rows[0].Cells[11].Value = "Отчет по отгрузке " + dataGridView1.Rows[0].Cells[0].Value.ToString().Substring(19, (dataGridView1.Rows[0].Cells[0].Value.ToString().Length - 20));
            dataGridView1.Rows[3].Cells[13].Value = dataGridView1.Rows[3].Cells[13].Value.ToString().Substring(13, 19);
             //dataGridView1.Rows[3].Cells[13].Value.ToString().Substring(13, 19);

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            button2.Enabled = false;
            button6.Enabled = false;
            button4.Enabled = false;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Файл за ДЕНЬ|*.xls; *.xlsx";
            openDialog.ShowDialog();

            try
            {
                ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                //Книга.
                ObjWorkBook = ObjExcel.Workbooks.Open(openDialog.FileName);
                //Таблица.
                ObjWorkSheet = ObjExcel.ActiveSheet as Worksheet;
                Range rg = null;
                int LastRowNumber = ObjWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
                int LastColNumber = ObjWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Column;
                dataGridView3.ColumnCount = LastColNumber;
                dataGridView3.RowCount = LastRowNumber;
                // MessageBox.Show(LastRowNumber.ToString());
                Int32 row = 1;
                dataGridView3.Rows.Clear();
                List<String> arr = new List<string>();
                while (row != LastRowNumber + 1) //ObjWorkSheet.get_Range("a" + row, "a" + row).Value != null)
                {
                    // Читаем данные из ячейки
                    rg = ObjWorkSheet.get_Range("a" + row, "bb" + row);
                    foreach (Range item in rg)
                    {
                        try
                        {
                            arr.Add(item.Value.ToString().Trim());
                        }
                        catch
                        {
                            arr.Add("");
                        }
                    }
                    dataGridView3.Rows.Add(arr[0], arr[1], arr[2], arr[3], arr[4], arr[5], arr[6], arr[7], arr[8],
                        arr[9], arr[10], arr[11], arr[12], arr[13], arr[14], arr[15], arr[16], arr[17], arr[18], arr[19],
                        arr[20], arr[21], arr[22], arr[23], arr[24], arr[25], arr[26], arr[27], arr[28], arr[29], arr[30]);
                    arr.Clear();
                    row++;
                }

                MessageBox.Show("Файл успешно считан!", "Считывания excel файла", MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message, "Ошибка при считывании excel файла", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                ObjWorkBook.Close(false, "", null);
                // Закрытие приложения Excel.
                ObjExcel.Quit();
                ObjWorkBook = null;
                ObjWorkSheet = null;
                ObjExcel = null;
                GC.Collect();
            }
            button2.Enabled = true;
            button6.Enabled = false;
        }













        //Новая обработка 1 грида!!!!



        private void button7_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1.Rows[i].Cells[1].Value.ToString().Contains("Подразделение"))
                {
                    dataGridView1.Rows[i + 1].Cells[1].Value = "XX";
                    dataGridView1.Rows[i + 2].Cells[1].Value = "XX";
                }
            }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1.Rows[i].Cells[1].Value.ToString().Equals(""))
                {
                    dataGridView1.Rows.RemoveAt(i);
                    i--;
                }
            }
            dataGridView1.Columns.RemoveAt(0);
            chst = dataGridView1.ColumnCount;

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("Отчет по отгрузке лома за период:"))
                {
                    otschet = i;
                }
            }
            
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (
                    (!(dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бр. Коростелевых ул., 52, терр. з-да \"Гидропресс\"") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бугуруслан г., Восточное ш. 1") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бузулук г., ул. Промышленная, 6") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Донгузская ул, 20") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Илек с., ул. Шоссейная, 54Б") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Краснохолм с., ул. Шоссейная 1А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Кувандык г., ул. Дзержинского, 1А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Курманаевка п.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Медногорск г., ул. Комсомольская, 40") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Новосергиевский п.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Октябрьское, ул. Транспортная, 1") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Орск, ул.Строителей, 44") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Переволоцкий п., ул. Ленинская, 2А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Подольск п., Промышленная ул., 3") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Сакмарский р-н, Сакмарская ст., терр. СМП-639") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Саракташ п., ул. Производственная, 4") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Соль-Илецк г., ул. Гонтаренко, 1А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Сорочинск г., ул. Пролетарская, 3") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Терешковой  ул., 287 д.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Тоцкое с. Тоцкий р-н, Автомобилистов ул., 1Е") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Центральная ул., 1 д.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Шильда, Топсклад") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("ИТОГО:") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Транзит") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("Отчет") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("Подразделение") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("Из расчета") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("XX") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("2А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("2А1") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("3А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("3А1") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("3А2") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("3АР") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("3АР2") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("8А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("9А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("10А")||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("41 счет"))))
                {
                    dataGridView1.Rows.RemoveAt(i);
                    i--;
                }
            }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("Отчет по отгрузке лома за период:"))
                {
                    otschet = i;
                }
            }

            
            for (int i = otschet; i < dataGridView1.RowCount; i++)
            {
                if (
                    (!(dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бр. Коростелевых ул., 52, терр. з-да \"Гидропресс\"") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бугуруслан г., Восточное ш. 1") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бузулук г., ул. Промышленная, 6") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Донгузская ул, 20") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Илек с., ул. Шоссейная, 54Б") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Краснохолм с., ул. Шоссейная 1А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Кувандык г., ул. Дзержинского, 1А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Курманаевка п.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Медногорск г., ул. Комсомольская, 40") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Новосергиевский п.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Октябрьское, ул. Транспортная, 1") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Орск, ул.Строителей, 44") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Переволоцкий п., ул. Ленинская, 2А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Подольск п., Промышленная ул., 3") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Сакмарский р-н, Сакмарская ст., терр. СМП-639") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Саракташ п., ул. Производственная, 4") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Соль-Илецк г., ул. Гонтаренко, 1А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Сорочинск г., ул. Пролетарская, 3") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Терешковой  ул., 287 д.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Тоцкое с. Тоцкий р-н, Автомобилистов ул., 1Е") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Центральная ул., 1 д.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Шильда, Топсклад") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("ИТОГО:") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Транзит") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("Отчет") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("Подразделение") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("Из расчета") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("XX"))))
                {
                    dataGridView1.Rows.RemoveAt(i);
                    i--;
                }
            }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("41 счет"))
                {
                    dataGridView1.Rows.RemoveAt(i+1);
                    dataGridView1.Rows.RemoveAt(i);
                    i--;
                }
            }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (   dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("2А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("2А1") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("3А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("3А1") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("3А2") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("3АР") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("3АР2") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("8А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("9А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("10А"))
                {
                    for (int j = 1; j < dataGridView1.ColumnCount-1; j++)
                    {
                        dataGridView1.Rows[i].Cells[j].Value = "";
                    }
                }
            }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString() == "" || dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value == null || dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString().Length == 0)
                {
                    dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value = 0;
                }
                
            }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("Отчет по отгрузке лома за период:"))
                {
                    otschet = i;
                }
            }

            got_prod = new double[22];
            
            //double got_prod_i = 0;



            //1
            int verh = 0, niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бр. Коростелевых ул., 52, терр. з-да \"Гидропресс\""))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бугуруслан г., Восточное ш. 1"))
                {
                    niz = i;
                }
                
            }

            if (verh-niz == -1)
            {
                got_prod[0] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[0] += Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                } 
            }
            
            
            //2
            verh = 0; 
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бугуруслан г., Восточное ш. 1"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бузулук г., ул. Промышленная, 6"))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[1] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[1] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }
            //3
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бузулук г., ул. Промышленная, 6"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Донгузская ул, 20"))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[2] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[2] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            //4
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Донгузская ул, 20"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Илек с., ул. Шоссейная, 54Б"))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[3] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[3] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }
            //5
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Илек с., ул. Шоссейная, 54Б"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Краснохолм с., ул. Шоссейная 1А"))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[4] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[4] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }
            //6
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Краснохолм с., ул. Шоссейная 1А"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Кувандык г., ул. Дзержинского, 1А"))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[5] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[5] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }
            //7
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Кувандык г., ул. Дзержинского, 1А"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Курманаевка п."))
                {
                    niz = i;
                }

            }
            if (verh - niz == -1)
            {
                got_prod[6] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[6] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            //8
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Курманаевка п."))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Медногорск г., ул. Комсомольская, 40"))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[7] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[7] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            //9
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Медногорск г., ул. Комсомольская, 40"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Новосергиевский п."))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[8] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[8] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            //10
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Новосергиевский п."))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Октябрьское, ул. Транспортная, 1"))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[9] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[9] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            //11
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Октябрьское, ул. Транспортная, 1"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Орск, ул.Строителей, 44"))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[10] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[10] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            //12
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Орск, ул.Строителей, 44"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Переволоцкий п., ул. Ленинская, 2А"))
                {
                    niz = i;
                }

            }
            if (verh - niz == -1)
            {
                got_prod[11] = 0.0;
            }
            else
            {

                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[11] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }
            //13
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Переволоцкий п., ул. Ленинская, 2А"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Подольск п., Промышленная ул., 3"))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[12] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[12] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            //14
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Подольск п., Промышленная ул., 3"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Сакмарский р-н, Сакмарская ст., терр. СМП-639"))
                {
                    niz = i;
                }

            }
            if (verh - niz == -1)
            {
                got_prod[13] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[13] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }
            //15
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Сакмарский р-н, Сакмарская ст., терр. СМП-639"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Саракташ п., ул. Производственная, 4"))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[14] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[14] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            //16
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Саракташ п., ул. Производственная, 4"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Соль-Илецк г., ул. Гонтаренко, 1А"))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[15] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[15] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            //17
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Соль-Илецк г., ул. Гонтаренко, 1А"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Сорочинск г., ул. Пролетарская, 3"))
                {
                    niz = i;
                }

            }
            if (verh - niz == -1)
            {
                got_prod[16] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[16] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            //18
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Сорочинск г., ул. Пролетарская, 3"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Терешковой  ул., 287 д."))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[17] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[17] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            //19
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Терешковой  ул., 287 д."))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Тоцкое с. Тоцкий р-н, Автомобилистов ул., 1Е"))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[18] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[18] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            //20
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Тоцкое с. Тоцкий р-н, Автомобилистов ул., 1Е"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Центральная ул., 1 д."))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[19] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[19] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }
            //21
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Центральная ул., 1 д."))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Шильда, Топсклад"))
                {
                    niz = i;
                }

            }

            if (verh - niz == -1)
            {
                got_prod[20] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[20] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            //22
            verh = 0;
            niz = 0;
            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Шильда, Топсклад"))
                {
                    verh = i;
                }
            }

            for (int i = 0; i < otschet; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("ИТОГО:"))
                {
                    niz = i;
                }

            }
            if (verh - niz == -1)
            {
                got_prod[21] = 0.0;
            }
            else
            {
                for (int i = verh + 1; i < niz; i++)
                {
                    got_prod[21] +=
                        Convert.ToDouble(dataGridView1.Rows[i].Cells[dataGridView1.ColumnCount - 1].Value.ToString());
                }
            }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (
                    (!(dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бр. Коростелевых ул., 52, терр. з-да \"Гидропресс\"") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бугуруслан г., Восточное ш. 1") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Бузулук г., ул. Промышленная, 6") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Донгузская ул, 20") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Илек с., ул. Шоссейная, 54Б") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Краснохолм с., ул. Шоссейная 1А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Кувандык г., ул. Дзержинского, 1А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Курманаевка п.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Медногорск г., ул. Комсомольская, 40") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Новосергиевский п.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Октябрьское, ул. Транспортная, 1") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Орск, ул.Строителей, 44") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Переволоцкий п., ул. Ленинская, 2А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Подольск п., Промышленная ул., 3") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Сакмарский р-н, Сакмарская ст., терр. СМП-639") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Саракташ п., ул. Производственная, 4") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Соль-Илецк г., ул. Гонтаренко, 1А") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Сорочинск г., ул. Пролетарская, 3") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Терешковой  ул., 287 д.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Тоцкое с. Тоцкий р-н, Автомобилистов ул., 1Е") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Центральная ул., 1 д.") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Шильда, Топсклад") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("ИТОГО:") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("Транзит") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("Отчет") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("Подразделение") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Contains("Из расчета") ||
                       dataGridView1.Rows[i].Cells[0].Value.ToString().Equals("XX"))))
                {
                    dataGridView1.Rows.RemoveAt(i);
                    i--;
                }
            }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                dataGridView1.Rows[i].HeaderCell.Value = (i + 1).ToString();
            }

            int bb;
            bb = 2 * chst;
            DataGridViewTextBoxColumn[] column = new DataGridViewTextBoxColumn[chst];

            for (int i = 0; i < chst; i++)
            {
                column[i] = new DataGridViewTextBoxColumn();
            }
            dataGridView1.Columns.AddRange(column);

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value.ToString().Length > 17 &&
                    dataGridView1.Rows[i].Cells[0].Value.ToString().Substring(0, 17) == "Отчет по отгрузке")
                {
                    a = i;

                    for (int j = a; j < dataGridView1.RowCount; j++)
                    {
                        int t = 0;
                        for (int k = chst; k < bb; k++)
                        {
                            dataGridView1.Rows[j - a].Cells[k].Value = dataGridView1.Rows[j].Cells[t].Value;
                            t++;
                        }
                    }
                }
            }

            for (int i = dataGridView1.RowCount - 1; i >= a + 1; i--)
            {
                dataGridView1.Rows.RemoveAt(i);
            }

            for (int i = dataGridView1.ColumnCount - 1; i >= 0; i--)
            {
                if (dataGridView1.Rows[2].Cells[i].Value.ToString().Equals("") &&
                    dataGridView1.Rows[3].Cells[i].Value.ToString().Equals("") &&
                    dataGridView1.Rows[4].Cells[i].Value.ToString().Equals("") &&
                    dataGridView1.Rows[5].Cells[i].Value.ToString().Equals("") &&
                    dataGridView1.Rows[6].Cells[i].Value.ToString().Equals("") &&
                    dataGridView1.Rows[7].Cells[i].Value.ToString().Equals(""))
                {
                    dataGridView1.Columns.RemoveAt(i);
                }
            }

            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                dataGridView1.Columns[i].HeaderText = (i + 1).ToString();
            }

            for (int i = dataGridView1.ColumnCount - 1; i >= 0; i--)
            {

                if (dataGridView1.Columns[i].HeaderText.Equals("31") ||
                    dataGridView1.Columns[i].HeaderText.Equals("30") ||
                    dataGridView1.Columns[i].HeaderText.Equals("28") ||
                    dataGridView1.Columns[i].HeaderText.Equals("20") ||
                    dataGridView1.Columns[i].HeaderText.Equals("18") ||
                    dataGridView1.Columns[i].HeaderText.Equals("17") ||
                    dataGridView1.Columns[i].HeaderText.Equals("15") ||
                    dataGridView1.Columns[i].HeaderText.Equals("14") ||
                    dataGridView1.Columns[i].HeaderText.Equals("13") ||
                    dataGridView1.Columns[i].HeaderText.Equals("12") ||
                    dataGridView1.Columns[i].HeaderText.Equals("10") ||
                    dataGridView1.Columns[i].HeaderText.Equals("7") ||
                    dataGridView1.Columns[i].HeaderText.Equals("4"))
                {
                    dataGridView1.Columns.RemoveAt(i);
                }
            }


            for (int i = 0; i < dataGridView1.ColumnCount; i++)
            {
                dataGridView1.Columns[i].HeaderText = (i + 1).ToString();
            }

        for (int i = 5; i < dataGridView1.RowCount-1; i++)
                for (int j = 5; j < dataGridView1.RowCount-1; j++)
                {
                    if (dataGridView1.Rows[i].Cells[0].Value == dataGridView1.Rows[j].Cells[10])

                    {
                        string s = String.Empty;
                        s = dataGridView1.Rows[j].Cells[10].Value.ToString();
                        dataGridView1.Rows[j].Cells[10].Value = "";
                        dataGridView1.Rows[j].Cells[10].Value = dataGridView1.Rows[i].Cells[10].Value;
                        dataGridView1.Rows[i].Cells[10].Value = "";
                        dataGridView1.Rows[i].Cells[10].Value = s;
                    }

                }
            dataGridView1.Rows[2].Cells[2].Value = "";
            dataGridView1.Rows[3].Cells[2].Value = "";
            dataGridView1.Rows[2].Cells[4].Value = "";
            dataGridView1.Rows[3].Cells[4].Value = "";
            dataGridView1.Rows[2].Cells[6].Value = "";
            dataGridView1.Rows[3].Cells[6].Value = "";
            dataGridView1.Rows[2].Cells[8].Value = "";
            dataGridView1.Rows[3].Cells[8].Value = "";

            int ii = 5;
            for (int j = 5; j < dataGridView1.RowCount - 1; j++)
            {
                if (dataGridView1.Rows[ii].Cells[0].Value.ToString() == dataGridView1.Rows[j].Cells[9].Value.ToString())
                {
                    string s1 = String.Empty;
                    string s2 = String.Empty;
                    string s3 = String.Empty;
                    string s4 = String.Empty;
                    string s5 = String.Empty;
                    string s6 = String.Empty;
                    string s7 = String.Empty;
                    string s8 = String.Empty;
                    string s9 = String.Empty;

                    s1 = dataGridView1.Rows[j].Cells[9].Value.ToString();
                    s2 = dataGridView1.Rows[j].Cells[10].Value.ToString();
                    s3 = dataGridView1.Rows[j].Cells[11].Value.ToString();
                    s4 = dataGridView1.Rows[j].Cells[12].Value.ToString();
                    s5 = dataGridView1.Rows[j].Cells[13].Value.ToString();
                    s6 = dataGridView1.Rows[j].Cells[14].Value.ToString();
                    s7 = dataGridView1.Rows[j].Cells[15].Value.ToString();
                    s8 = dataGridView1.Rows[j].Cells[16].Value.ToString();
                    s9 = dataGridView1.Rows[j].Cells[17].Value.ToString();

                    dataGridView1.Rows[j].Cells[9].Value = "";
                    dataGridView1.Rows[j].Cells[10].Value = "";
                    dataGridView1.Rows[j].Cells[11].Value = "";
                    dataGridView1.Rows[j].Cells[12].Value = "";
                    dataGridView1.Rows[j].Cells[13].Value = "";
                    dataGridView1.Rows[j].Cells[14].Value = "";
                    dataGridView1.Rows[j].Cells[15].Value = "";
                    dataGridView1.Rows[j].Cells[16].Value = "";
                    dataGridView1.Rows[j].Cells[17].Value = "";

                    dataGridView1.Rows[j].Cells[9].Value = dataGridView1.Rows[ii].Cells[9].Value;
                    dataGridView1.Rows[j].Cells[10].Value = dataGridView1.Rows[ii].Cells[10].Value;
                    dataGridView1.Rows[j].Cells[11].Value = dataGridView1.Rows[ii].Cells[11].Value;
                    dataGridView1.Rows[j].Cells[12].Value = dataGridView1.Rows[ii].Cells[12].Value;
                    dataGridView1.Rows[j].Cells[13].Value = dataGridView1.Rows[ii].Cells[13].Value;
                    dataGridView1.Rows[j].Cells[14].Value = dataGridView1.Rows[ii].Cells[14].Value;
                    dataGridView1.Rows[j].Cells[15].Value = dataGridView1.Rows[ii].Cells[15].Value;
                    dataGridView1.Rows[j].Cells[16].Value = dataGridView1.Rows[ii].Cells[16].Value;
                    dataGridView1.Rows[j].Cells[17].Value = dataGridView1.Rows[ii].Cells[17].Value;

                    dataGridView1.Rows[ii].Cells[9].Value = "";
                    dataGridView1.Rows[ii].Cells[10].Value = "";
                    dataGridView1.Rows[ii].Cells[11].Value = "";
                    dataGridView1.Rows[ii].Cells[12].Value = "";
                    dataGridView1.Rows[ii].Cells[13].Value = "";
                    dataGridView1.Rows[ii].Cells[14].Value = "";
                    dataGridView1.Rows[ii].Cells[15].Value = "";
                    dataGridView1.Rows[ii].Cells[16].Value = "";
                    dataGridView1.Rows[ii].Cells[17].Value = "";

                    dataGridView1.Rows[ii].Cells[9].Value = s1;
                    dataGridView1.Rows[ii].Cells[10].Value = s2;
                    dataGridView1.Rows[ii].Cells[11].Value = s3;
                    dataGridView1.Rows[ii].Cells[12].Value = s4;
                    dataGridView1.Rows[ii].Cells[13].Value = s5;
                    dataGridView1.Rows[ii].Cells[14].Value = s6;
                    dataGridView1.Rows[ii].Cells[15].Value = s7;
                    dataGridView1.Rows[ii].Cells[16].Value = s8;
                    dataGridView1.Rows[ii].Cells[17].Value = s9;

                    ii++;
                    j = 5;


                }

            }

            for (int i = 4; i < dataGridView1.RowCount-1; i++)
            {
                if ((dataGridView1.Rows[i].Cells[16].Value.ToString() == "") || (dataGridView1.Rows[i].Cells[16].Value == null) || (dataGridView1.Rows[i].Cells[16].Value.ToString().Length == 0))
                {
                    dataGridView1.Rows[i].Cells[16].Value = "0";
                }
            }

            for (int i = 4; i < dataGridView1.RowCount - 1; i++)
            {
                dataGridView1.Rows[i].Cells[16].Value = (Convert.ToInt32(dataGridView1.Rows[i].Cells[16].Value) + Convert.ToInt32(dataGridView1.Rows[i].Cells[17].Value));
            }

            dataGridView1.Columns.RemoveAt(dataGridView1.ColumnCount-1);

            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[0].Value = "";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[0].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[0].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[1].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[1].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[2].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[2].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[3].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[3].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[4].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[4].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[5].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[5].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[6].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[6].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[7].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[7].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[8].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[8].Value;
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[9].Value = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[9].Value;

            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[0].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[1].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[2].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[3].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[4].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[5].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[6].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[7].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[8].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[9].Value = "-";

            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[10].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[11].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[12].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[13].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[14].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[15].Value = "-";

            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[0].Value = "Транзит"; 

            dataGridView1.Columns.RemoveAt(9);

            dataGridView1.Rows[2].Cells[0].Value = "Подразделение";

            for (int j = 0; j < dataGridView1.ColumnCount; j++)
            {
                dataGridView1.Columns[j].HeaderText = (j + 1).ToString();
            }

            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[9].Value = dataGridView1.Rows[4].Cells[9].Value;
            dataGridView1.Rows[4].Cells[9].Value = "";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[10].Value = dataGridView1.Rows[4].Cells[10].Value;
            dataGridView1.Rows[4].Cells[10].Value = "";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[11].Value = dataGridView1.Rows[4].Cells[11].Value;
            dataGridView1.Rows[4].Cells[11].Value = "";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[12].Value = dataGridView1.Rows[4].Cells[12].Value;
            dataGridView1.Rows[4].Cells[12].Value = "";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[13].Value = dataGridView1.Rows[4].Cells[13].Value;
            dataGridView1.Rows[4].Cells[13].Value = "";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[14].Value = dataGridView1.Rows[4].Cells[14].Value;
            dataGridView1.Rows[4].Cells[14].Value = "";
            dataGridView1.Rows[dataGridView1.RowCount - 1].Cells[15].Value = dataGridView1.Rows[4].Cells[15].Value;
            dataGridView1.Rows[4].Cells[15].Value = "";

            string q1, q3, q4;




            q1 = dataGridView1.Rows[4].Cells[7].Value.ToString();
            dataGridView1.Rows[2].Cells[7].Value = "Вагоны";
            dataGridView1.Rows[4].Cells[7].Value = "Кол-во, шт.";
            dataGridView1.Rows[4].Cells[8].Value = "";
            
            dataGridView1.Rows[3].Cells[8].Value = dataGridView1.Rows[3].Cells[7].Value;
            dataGridView1.Rows[2].Cells[8].Value = "Денежные средства";
            dataGridView1.Rows[4].Cells[8].Value = "Кол-во, руб.";

           dataGridView1.Rows[2].Cells[9].Value = "План на период:";

            q3 = dataGridView1.Rows[3].Cells[10].Value.ToString();
            dataGridView1.Rows[3].Cells[10].Value = "";
            dataGridView1.Rows[3].Cells[10].Value = "План на: " + q3;

            q4 = dataGridView1.Rows[3].Cells[11].Value.ToString();
            dataGridView1.Rows[3].Cells[11].Value = "";
            dataGridView1.Rows[3].Cells[11].Value = "Отгрузка на: " + q4;

            dataGridView1.Rows[2].Cells[15].Value = "";
            dataGridView1.Rows[2].Cells[15].Value = "Остаток на текущую дату";
            dataGridView1.Rows[1].Cells[10].Value = "";
            dataGridView1.Rows[2].Cells[13].Value = "";
            dataGridView1.Rows[2].Cells[14].Value = "";
            dataGridView1.Rows[3].Cells[13].Value = "";

            dataGridView1.Rows[3].Cells[13].Value = "На дату: " + dataGridView1.Rows[3].Cells[12].Value;
            dataGridView1.Rows[2].Cells[13].Value = "Кол-во готового лома, тонн";

            dataGridView1.Rows[4].Cells[9].Value = "Кол-во, тонн";
            dataGridView1.Rows[4].Cells[11].Value = "Кол-во, тонн";
            dataGridView1.Rows[4].Cells[13].Value = "Кол-во, тонн";

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                dataGridView1.Rows[i].Cells[14].Value = "";
            }

            for (int i = 5; i < dataGridView1.RowCount; i++)
            {
                dataGridView1.Rows[i].Cells[13].Value = "";
                dataGridView1.Rows[i].Cells[8].Value = "";
                dataGridView1.Rows[i].Cells[7].Value = "";
            }

            dataGridView1.Rows[dataGridView1.RowCount-2].Cells[7].Value = "-";
            dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[8].Value = "-";


            for (int i = 5; i < got_prod.Length+5; i++)
            {
                dataGridView1.Rows[i].Cells[13].Value = got_prod[i - 5];
            }

            dataGridView1.Rows[dataGridView1.RowCount-1].Cells[13].Value = got_prod.Sum();

            dataGridView3.Rows.RemoveAt(0);
            dataGridView3.Columns.RemoveAt(0);

            for (int i = 0; i < 24; i++)
            {
                dataGridView1.Rows[i + 5].Cells[7].Value = dataGridView3.Rows[i].Cells[0].Value;
                dataGridView1.Rows[i + 5].Cells[8].Value = dataGridView3.Rows[i].Cells[1].Value;
           }
        }
     }
  }
