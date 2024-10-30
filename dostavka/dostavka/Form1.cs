using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

// Чтобы читать данные из файла Excel
using ExcelDataReader;
using Excel = Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;


namespace dostavka
{
    public partial class Form1 : Form
    {
        private string fileName = string.Empty;
        // Для хранения листов в Excel
        private DataTableCollection tableCollection = null;
        
        FileStream stream;
        IExcelDataReader reader;
        DataSet ds;
        DataTable table;
        DateTime today = new DateTime(2024, 10, 24, 12, 0, 0);

        public Form1()
        {
            InitializeComponent();
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult result = openFileDialog1.ShowDialog();
                // Читиение файла Excel
                if (result == DialogResult.OK)
                {
                    fileName = openFileDialog1.FileName;
                    Text = fileName;
                    OpenExcelFile(fileName);
                }
                else // Вывод ошибки
                {
                    throw new Exception("Файл не выбран!");
                }
            }
            // Вывод ошибки
            catch ( Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        // Открытие и считывание таблицы Excel (путь)
        private void OpenExcelFile(string path)
        {
            stream = File.Open(path, FileMode.Open, FileAccess.Read);
            // Приводим значение .CreateReader к IExcelDataReader 
            reader = ExcelReaderFactory.CreateReader(stream);
            // Создание DataSet
            ds = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                // Лямбда выражение, инициализация класса
                ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
                {
                    // Считывание первой строки как названия стоблцов
                    UseHeaderRow = false
                }
            });
            // БД создана из Excel файла
            // Все листы этого файла передаем в tableCollection
            tableCollection = ds.Tables;
            // Очистка ComboBox
            toolStripComboBox1.Items.Clear();
            // Переберам названия всех листов, которые может выбрать пользователь
            foreach (DataTable table in tableCollection)
            {
                toolStripComboBox1.Items.Add(table.TableName);
            }
            // По умолчанию выбран первый лист таблицы
            toolStripComboBox1.SelectedIndex = 0;
        }

        // Обработка изменений в ComboBox
        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Отображение листа Excel в dataGridView
            // В соотвтствии с названием листа - он отображается
            table = tableCollection[Convert.ToString(toolStripComboBox1.SelectedItem)];
            ResultGrid.DataSource = table;

            toolStripComboBox2.Items.Clear();

            // Создание списка для фильтрации по району
            List<string> list = new List<string>();
            foreach (DataGridViewRow row in ResultGrid.Rows)
            {
                if (row.Cells.Count >= 2 && row.Cells[2].Value != null)
                {
                    list.Add(row.Cells[2].Value.ToString());
                }
            }
            // Убираем дубликаты из списка
            List<string> resList = list.Distinct().ToList();
            // Добавление списка в combobox
            toolStripComboBox2.Items.Add("Все результаты");
            for (int i = 1; i <= resList.Count-1; i++)
            {
                toolStripComboBox2.Items.Add(resList[i].ToString());
            }
        }

        private void toolStripComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Для отображения нужных строк
            CurrencyManager cm = (CurrencyManager)BindingContext[ResultGrid.DataSource];
            cm.SuspendBinding();
            ResultGrid.ReadOnly = true;
            cm.ResumeBinding();

            // Отображаем нужные строки в таблице после выбранного района
            if (toolStripComboBox2.SelectedIndex != 0)
            {
                foreach (DataGridViewRow row in ResultGrid.Rows)
                {
                    row.Visible = true;
                }
                foreach (DataGridViewRow row in ResultGrid.Rows)
                {
                    if (row.Index > 0)
                    {
                        if (row.Cells[2].Value != null && (string)row.Cells[2].Value != toolStripComboBox2.Text)
                        {
                            row.Visible = false;
                        }
                    }
                }
                foreach (DataGridViewRow row in ResultGrid.Rows)
                {
                    if (row.Index > 0)
                    {
                        DateTime dt = DateTime.Parse(row.Cells[3].Value.ToString());
                        DateTime first_order_dt = new DateTime(2024, 10, 25, 12, 30, 00);
                        if (dt > first_order_dt)
                        {
                            row.Visible = false;
                        }
                    }
                }
            }
            else
            {
                foreach (DataGridViewRow row in ResultGrid.Rows)
                {
                    row.Visible = true;
                }
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = new Excel.Application();
            exApp.Workbooks.Add();
            Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;

            int i, j;
            for (i = 0; i <= ResultGrid.RowCount - 2; i++)
            {
                for (j = 0; j <= ResultGrid.ColumnCount - 1; j++)
                {
                    if (ResultGrid.Rows[i].Visible)
                    {
                        wsh.Cells[1, j + 1] = ResultGrid.Columns[j].HeaderText.ToString();
                        wsh.Cells[i + 1, j + 1] = ResultGrid[j, i].Value.ToString();
                    }
                }
            }

            exApp.Visible = true;
        }
    }
}
