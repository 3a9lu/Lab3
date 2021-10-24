using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net;
using System.Data;

namespace Lab1
{
    public partial class Метод : Form
    {
        public static List<Point> steps = new List<Point>(); // Список для точек
        public static List<Point> pointsE = new List<Point>(); // Список для точек
        public static List<Point> pointsS = new List<Point>(); // Список для точек

        public class Point // Сохраниние точек
        {
            public double x, y;
            public Point(double X, double Y)
            {
                this.x = X;
                this.y = Y;
            }
        }
        public Метод()
        {
            InitializeComponent();
            StartPosition = FormStartPosition.CenterScreen;  // Вывод формы по центру экрана
            label1.Text = "";
            label2.Text = "";
        }

        private void очиститьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox5.Clear(); // Очистка X
            textBox6.Clear(); // Очистка Y
            dataGridView1.Rows.Clear();
            chart1.Series[0].Points.Clear(); // Очистка графика
            chart1.Series[1].Points.Clear(); // Очистка точки минимума
            label1.Text = "";
            label2.Text = "";
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        async void Excel()
        {
            await Task.Run(() =>
            {
                label1.Text = "";
                label2.Text = "";
                chart1.Series[0].Points.Clear();  // Очистка точек
                chart1.Series[1].Points.Clear();
                pointsE.Clear(); // Очистка списка
                dataGridView1.Rows.Clear();

                string path = @"D:\Проекты\Lab1\Lab1\Excel.xlsx";
                Excel.Application ObjExcel = new Excel.Application();

                Workbook ObjWorkBook = ObjExcel.Workbooks.Open(path); // Открываем книгу
                Worksheet ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1]; // Выбираем лист

                Range xRange = ObjWorkSheet.UsedRange.Columns[1]; // Первый столбец
                Range yRange = ObjWorkSheet.UsedRange.Columns[2]; // Второй столбец
                Array xCells = (Array)xRange.Cells.Value2;
                Array yCells = (Array)yRange.Cells.Value2;

                string[] xColumn = xCells.OfType<object>().Select(o => o.ToString()).ToArray();
                string[] yColumn = yCells.OfType<object>().Select(o => o.ToString()).ToArray();

                for (int i = 0; i < xColumn.Length; ++i)
                {
                    Point point = new Point(double.Parse(xColumn[i]), double.Parse(yColumn[i]));
                    pointsE.Add(point);
                    dataGridView1.Rows.Add(pointsE[i].x, pointsE[i].y);
                    chart1.Series[1].Points.AddXY(pointsE[i].x, pointsE[i].y);
                }
                ObjWorkBook.Close(); // Закрытие книги
                ObjExcel.Quit(); // Выход из Excel
            });
        }

        async void GoogleSheets()
        {
            await Task.Run(() =>
            {
                string path = @"D:\Проекты\Lab1\Lab1\Sheets.xlsx";
                System.IO.File.Delete(path);
                string link = "https://docs.google.com/spreadsheets/d/1PVFSai9ncjEyAPu6ARgD1du4Hye9IPevvrMRBvgAQ5s/export?format=xlsx";

                using (var client = new WebClient()) // Скачивание файла
                {
                    client.DownloadFile(new Uri(link), path);
                }
                chart1.Series[0].Points.Clear();  // Очистка точек
                chart1.Series[1].Points.Clear();
                pointsS.Clear(); // Очистка списка
                dataGridView1.Rows.Clear();

                Excel.Application ObjExcel = new Excel.Application();

                Workbook ObjWorkBook = ObjExcel.Workbooks.Open(path); // Открываем книгу
                Worksheet ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1]; // Выбираем лист

                Range xRange = ObjWorkSheet.UsedRange.Columns[1]; // Первый столбец
                Range yRange = ObjWorkSheet.UsedRange.Columns[2]; // Второй столбец
                Array xCells = (Array)xRange.Cells.Value2;
                Array yCells = (Array)yRange.Cells.Value2;

                string[] xColumn = xCells.OfType<object>().Select(o => o.ToString()).ToArray();
                string[] yColumn = yCells.OfType<object>().Select(o => o.ToString()).ToArray();

                for (int i = 0; i < xColumn.Length; ++i)
                {
                    Point point = new Point(double.Parse(xColumn[i]), double.Parse(yColumn[i]));
                    pointsS.Add(point);
                    dataGridView1.Rows.Add(pointsS[i].x, pointsS[i].y);
                    chart1.Series[1].Points.AddXY(pointsS[i].x, pointsS[i].y);
                }
                ObjWorkBook.Close(); // Закрытие книги
                ObjExcel.Quit(); // Выход из Excel
            });
        }

        private void считатьСExcelToolStripMenuItem_Click(object sender, EventArgs e) // Считывание с Excel асинхронно
        {
            Excel();
        }

        private void считатьСGoogleSheetsToolStripMenuItem_Click(object sender, EventArgs e) // Считывание с Google Sheets асинхронно
        {
            GoogleSheets();
        }

        async void MNK()
        {
            await Task.Run(() =>
            {
                steps.Clear(); // Очистка списка

                int n = Convert.ToInt32(dataGridView1.Rows.Count - 1);
                int nn = n; double x = 0; double sx = 0; double y = 0; double sy = 0; double sxy = 0; double sx2 = 0;
                double b = 0; double a = 0; double c = 0; double xmax = 0; double xmin = 0; double y1 = 0; double y2 = 0;
                double q = 0; double st = 0; double stp = 0; double sp = 0; double st2 = 0; double sx3 = 0;
                double sx4 = 0; double sx2y = 0; double D = 0; double Da = 0; double Db = 0; double Dc = 0;

                for (int i = 0; i < n; ++i)
                {
                    x = Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                    y = Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                    sx += x;
                    sy += y;
                    sx2 += x * x;
                    sxy += x * y;
                    sx3 += x * x * x;
                    st += (1 / x);
                    sx4 += x * x * x * x;
                    st2 += (1 / (x * x));
                    stp += (y / x);
                    sp += (1 / y);
                    sx2y += (x * x) * y;

                    if (x > xmax)
                    {
                        xmax = x;
                    }
                    if (x < xmin)
                    {
                        xmin = x;
                    }
                }

                if (comboBox1.SelectedIndex == 0)
                {
                    a = (nn * sxy - sx * sy) / (nn * sx2 - sx * sx);
                    b = (sy - a * sx) / nn;
                    y1 = b + a * xmin;
                    y2 = b + a * xmax;
                    chart1.Series[0].Points.AddXY(xmin, y1);
                    chart1.Series[0].Points.AddXY(xmax, y2);
                    for (int i = 0; i < n; ++i)
                    {
                        x = Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                        y = Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                        q += Math.Pow((y - (a * x + b)), 2);
                    }
                    label1.Text = "y = " + Math.Round(a, 3).ToString() + " * x + " + Math.Round(b, 3).ToString();
                }

                else if (comboBox1.SelectedIndex == 1)
                {
                    D = st * st - nn * st2;
                    Da = sy * st - stp * nn;
                    Db = st * stp - st2 * sy;
                    a = Da / D;
                    b = Db / D;
                    for (int i = 0; i < n; ++i)
                    {
                        x = Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                        y = Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                        y1 = b + a / x;
                        chart1.Series[0].Points.AddXY(x, y1);
                        q += Math.Pow((y - (a / x + b)), 2);
                    }
                    label1.Text = "y = " + Math.Round(a, 3).ToString() + " / x + " + Math.Round(b, 3).ToString();
                }

                if (comboBox1.SelectedIndex == 2) // Метод Крамера
                {
                    D = sx2 * sx2 * sx2 + sx * sx * sx4 + nn * sx3 * sx3 - nn * sx2 * sx4 - sx * sx3 * sx2 - sx2 * sx * sx3;
                    Da = sy * sx2 * sx2 + sx * sx * sx2y + nn * sxy * sx3 - nn * sx2 * sx2y - sx * sxy * sx2 - sy * sx * sx3;
                    Db = sx2 * sxy * sx2 + sy * sx * sx4 + nn * sx3 * sx2y - nn * sxy * sx4 - sy * sx3 * sx2 - sx2 * sx * sx2y;
                    Dc = sx2 * sx2 * sx2y + sx * sxy * sx4 + sy * sx3 * sx3 - sy * sx2 * sx4 - sx * sx3 * sx2y - sx2 * sxy * sx3;
                    a = Da / D;
                    b = Db / D;
                    c = Dc / D;
                    for (int i = 0; i < n; ++i)
                    {
                        x = Convert.ToDouble(dataGridView1.Rows[i].Cells[0].Value);
                        y = Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value);
                        y1 = c + b * x + a * (x * x);
                        chart1.Series[0].Points.AddXY(x, y1);
                        q += Math.Pow((y - (a * (x * x) + b * x + c)), 2);
                    }
                    label1.Text = "y = " + Math.Round(a, 3).ToString() + " * x^2 + " + Math.Round(b, 3).ToString() + " * x + " + Math.Round(c, 3).ToString();
                }
                label2.Text = q.ToString();
            });
        }

        private void рассчитатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            chart1.Series[0].Points.Clear();

            try
            {
                MNK(); // Асинхронный метод наименьших квадратов
            }
            catch
            {
                MessageBox.Show("Что-то пошло не так", "Ошибка!");
            }
        }

        private void button3_Click(object sender, EventArgs e) // Добавление в DataGridView точек
        {
            if (textBox5.Text == "" || textBox6.Text == "")
            {
                MessageBox.Show("Заполните оба поля!", "Ошибка!");
            }
            else
            {
                Point xy = new Point(Convert.ToDouble(textBox5.Text), Convert.ToDouble(textBox6.Text));
                steps.Add(xy);
                dataGridView1.Rows.Add(xy.x, xy.y);
                textBox5.Text = "";
                textBox6.Text = "";
                chart1.Series[1].Points.AddXY(xy.x, xy.y);
            }
        }

        private void button4_Click(object sender, EventArgs e) // Удаление таблицы
        {
            if (dataGridView1.Rows.Count > 0)
            {
                label1.Text = "";
                label2.Text = "";
                steps.Clear(); // Очистка списка
                chart1.Series[0].Points.Clear(); // Очистка графика
                chart1.Series[1].Points.Clear();
                dataGridView1.Rows.Clear();
            }
            else
            {
                MessageBox.Show("Таблица пустая!", "Ошибка");
            }
        }
    }
}


//ln((e^x+5^2)/x)
//log(5,4*x-3)-ln(2*x)+log10(3.14)
//x^5 - 120*x^4 + 749*x^3-1530*x^2+4823*x-14393
