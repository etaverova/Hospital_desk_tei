using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Hospital_desk_tei
{
    public partial class Form1 : Form
    {
        private Excel.Application excelapp; // Программа Excel
        private Excel.Window excelWindow; // Окно программы Excel
        private Excel.Workbooks excelappworkbooks; // Рабочие книги
        private Excel.Workbook excelappworkbook; // Рабочая книга
        private Excel.Sheets excelsheets; // Рабочие листы
        private Excel.Worksheet excelworksheet; // Рабочий лист
        private Excel.Range excelcells; // Диапазон ячеек или ячейка

        private Word.Application WordApp; // Программа Word
        private Word.Documents WordDocuments; // Документы
        private Word.Document WordDocument; // Документ
        private Word.Paragraphs WordParagraphs; // Параграфы
        private Word.Paragraph WordParagraph; // Параграф
        private Word.Range WordRange; // Выделенный диапазон
        public Form1()
        {
            InitializeComponent();
        }

        private void n_pal()
        {
            int i, k, t;
            DataRowView r, r1;
            this.больные1BindingSource.MoveFirst();
            i = 0;
            while (i < this.больные1BindingSource.Count)
            {
                r = (DataRowView)this.больные1BindingSource.Current;
                k = (int)r["Код_палаты"];
                this.палаты1BindingSource.Filter = "Код_палаты=" + Convert.ToString(k);
                r1 = (DataRowView)this.палаты1BindingSource.Current;
                if (r1 != null)
                    t = (int)r1["Номер_палаты"];
                else 
                    t = 0;
                r["Номер_палаты"] = t;
                i = i + 1;
                this.больные1BindingSource.MoveNext();
            }
            this.больные1BindingSource.Filter = "";
            this.палаты1BindingSource.Filter = "";
        }

        private void fio()
        {
            int i, k;
            string t;
            DataRowView r, r1;
            this.процедуры1BindingSource.MoveFirst();
            i = 0;
            while (i < this.процедуры1BindingSource.Count)
            {
                r = (DataRowView)this.процедуры1BindingSource.Current;
                k = (int)r["Код_ФИО"];
                this.больные1BindingSource.Filter = "Код_ФИО=" + Convert.ToString(k);
                r1 = (DataRowView)this.больные1BindingSource.Current;
                if (r1 != null)
                    t = (string)r1["ФИО_больного"];
                else
                    t = "";
                r["ФИО_больного"] = t;
                i = i + 1;
                this.процедуры1BindingSource.MoveNext();
            }
            this.процедуры1BindingSource.Filter = "";
            this.больные1BindingSource.Filter = "";
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "больницаDataSet.Отделения". При необходимости она может быть перемещена или удалена.
            this.отделенияTableAdapter.Fill(this.больницаDataSet.Отделения);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "больницаDataSet.Отделения". При необходимости она может быть перемещена или удалена.
            this.отделенияTableAdapter.Fill(this.больницаDataSet.Отделения);
            DataRowView drv = отделенияBindingSource.Current as DataRowView;
            // TODO: данная строка кода позволяет загрузить данные в таблицу "больницаDataSet.Палаты". При необходимости она может быть перемещена или удалена.
            this.палаты1TableAdapter.Fill(this.больницаDataSet.Палаты1, Convert.ToInt32(drv.Row["Код_отделения"]));
            // TODO: данная строка кода позволяет загрузить данные в таблицу "больницаDataSet.Палаты". При необходимости она может быть перемещена или удалена.
            this.больные1TableAdapter.Fill(this.больницаDataSet.Больные1, Convert.ToInt32(drv.Row["Код_отделения"]));
            // TODO: данная строка кода позволяет загрузить данные в таблицу "больницаDataSet.Палаты". При необходимости она может быть перемещена или удалена.
            this.процедуры1TableAdapter.Fill(this.больницаDataSet.Процедуры1, Convert.ToInt32(drv.Row["Код_отделения"]));
            n_pal();
            fio();

        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                int i;
                if (dataGridView1.SelectedRows.Count != 0)
                {
                    i = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value);
                }
                else
                {
                    dataGridView1.Rows[0].Selected = true;
                    i = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value);
                }
                this.палаты1TableAdapter.Fill(this.больницаDataSet.Палаты1, i);
                this.больные1TableAdapter.Fill(this.больницаDataSet.Больные1, i);
                this.процедуры1TableAdapter.Fill(this.больницаDataSet.Процедуры1, i);
                n_pal();
                fio();
            }
            catch
            {

            }
        }

        //----------------------------------------------------Отделения---------------------------------------------------------

        //Добавить
        private void button1_Click(object sender, EventArgs e)
        {
            Form2 f = new Form2();
            f.ShowDialog();
            if (f.Rc)
            {
                this.отделенияTableAdapter.Insert(f.nm);
                this.отделенияTableAdapter.Fill(this.больницаDataSet.Отделения);
            }
        }

        //Изменить
        private void button2_Click(object sender, EventArgs e)
        {
            DataRowView rdw = отделенияBindingSource.Current as DataRowView;
            Int32 ID = Convert.ToInt32(rdw.Row["Код_отделения"]);
            String nm = rdw.Row["Название_отделения"].ToString();

            Form2 f = new Form2();
            f.nm = nm;

            f.ShowDialog();
            if (f.Rc)
            {
                this.отделенияTableAdapter.Update(f.nm, ID);
                this.отделенияTableAdapter.Fill(this.больницаDataSet.Отделения);
            }

        }

        //Удалить
        private void button3_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("Вы точно хотите удалить услугу из списка?", "Подтверждение действий", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                DataRowView rdw = отделенияBindingSource.Current as DataRowView;
                Int32 ID = Convert.ToInt32(rdw.Row["Код_отделения"]);
                try
                {
                    отделенияTableAdapter.Delete(ID);
                    this.отделенияTableAdapter.Fill(this.больницаDataSet.Отделения);
                }
                catch
                {
                    MessageBox.Show("Ошибка удаления! Данную услугу удалить невозможно, данные о ней содержатся в журнале заказов");
                }
            }
        }

        //-------------------------------------------------------Палаты-------------------------------------------------------------

        //Добавить
        private void button4_Click(object sender, EventArgs e)
        {
            DataRowView rdw = отделенияBindingSource.Current as DataRowView;
            Form3 f = new Form3();
            f.ShowDialog();

            if (f.Rc)
            {
                try
                {
                    палаты1TableAdapter.Insert(Convert.ToInt32(rdw.Row["Код_отделения"]), f.Nm, f.Vm);
                    this.палаты1TableAdapter.Fill(this.больницаDataSet.Палаты1, Convert.ToInt32(rdw.Row["Код_отделения"]));
                }
                catch (Exception)
                {
                    MessageBox.Show("Невозможно добавить данные о палате");
                }
            }
        }

        //Изменить
        private void button5_Click(object sender, EventArgs e)
        {
            DataRowView rdw = отделенияBindingSource.Current as DataRowView;

            if (dataGridView2.SelectedRows.Count != 0)
            {
                Int32 ID = Convert.ToInt32(dataGridView2.CurrentRow.Cells[0].Value);
                Int32 Nm = Convert.ToInt32(dataGridView2.CurrentRow.Cells[1].Value);
                Int32 Vm = Convert.ToInt32(dataGridView2.CurrentRow.Cells[2].Value);

                Form3 f = new Form3();
                f.Nm = Nm;
                f.Vm = Vm;


                f.ShowDialog();
                if (f.Rc)
                {
                    палаты1TableAdapter.Update(Convert.ToInt32(rdw.Row["Код_отделения"]), f.Nm, f.Vm, ID);
                    this.палаты1TableAdapter.Fill(this.больницаDataSet.Палаты1, Convert.ToInt32(rdw.Row["Код_отделения"]));
                }
            }
            else
            {
                MessageBox.Show("Для изменения выбрано пустое поле! Невозможно произвести изменение!");
            }
        }

        //Удалить
        private void button6_Click(object sender, EventArgs e)
        {
            DataRowView rdw = отделенияBindingSource.Current as DataRowView;

            if (dataGridView2.SelectedRows.Count != 0)
            {
                var result = MessageBox.Show("Вы точно хотите удалить клиент из заказа?", "Подтверждение действий", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    DataGridViewRow dgvr = dataGridView2.CurrentRow;
                    Int32 ID = Convert.ToInt32(dgvr.Cells[0].Value);
                    палаты1TableAdapter.Delete(ID);
                    this.палаты1TableAdapter.Fill(this.больницаDataSet.Палаты1, Convert.ToInt32(rdw.Row["Код_отделения"]));
                }
            }
            else
            {
                MessageBox.Show("Для удаления выбрано пустое поле! Невозможно произвести удаление!");
            }
        }

        //-------------------------------------------------------Больные-------------------------------------------------------------

        //Добавить
        private void button7_Click(object sender, EventArgs e)
        {
            List<string> list = new List<string>();

            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                list.Add(dataGridView2.Rows[i].Cells[1].Value.ToString());
            }
            list.Sort();
            DataRowView rdw = отделенияBindingSource.Current as DataRowView;
            int pl = 0;
            Form4 f = new Form4(list);
            f.ShowDialog();

            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                if (dataGridView2.Rows[i].Cells[1].Value.ToString() == f.Pal)
                    pl = Convert.ToInt32(dataGridView2.Rows[i].Cells[0].Value);
            }

            if (f.Rc)
            {
                try
                {
                    больные1TableAdapter.Insert(Convert.ToInt32(rdw.Row["Код_отделения"]), pl, f.FIO, f.Date1, f.Date2, f.Di);
                    this.больные1TableAdapter.Fill(this.больницаDataSet.Больные1, Convert.ToInt32(rdw.Row["Код_отделения"]));
                    n_pal();
                }
                catch (Exception)
                {
                    MessageBox.Show("Невозможно добавить данные о палате");
                }
            }
        }

        //Изменить
        private void button8_Click(object sender, EventArgs e)
        {
            List <string> list = new List <string> ();

            for (int i = 0; i < dataGridView2.Rows.Count; i ++)
            {
                list.Add(dataGridView2.Rows[i].Cells[1].Value.ToString());
            }
            list.Sort();
            DataRowView rdw = отделенияBindingSource.Current as DataRowView;

            if (dataGridView3.SelectedRows.Count != 0)
            {
                Int32 ID = Convert.ToInt32(dataGridView3.CurrentRow.Cells[0].Value);
                String Pal = dataGridView3.CurrentRow.Cells[1].Value.ToString();
                int pl = 0;
                String FIO = dataGridView3.CurrentRow.Cells[2].Value.ToString();
                DateTime Date1 = Convert.ToDateTime(dataGridView3.CurrentRow.Cells[3].Value);
                DateTime Date2 = Convert.ToDateTime(dataGridView3.CurrentRow.Cells[4].Value);
                String Di = dataGridView3.CurrentRow.Cells[5].Value.ToString();

                Form4 f = new Form4(list);
                f.Pal = Pal;
                f.FIO = FIO;
                f.Date1 = Date1;
                f.Date2 = Date2;
                f.Di = Di;

                f.ShowDialog();

                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    if (dataGridView2.Rows[i].Cells[1].Value.ToString() == f.Pal)
                        pl = Convert.ToInt32(dataGridView2.Rows[i].Cells[0].Value);
                }
                if (f.Rc)
                {
                    больные1TableAdapter.Update(Convert.ToInt32(rdw.Row["Код_отделения"]), pl, f.FIO, f.Date1, f.Date2, f.Di, ID);
                    this.больные1TableAdapter.Fill(this.больницаDataSet.Больные1, Convert.ToInt32(rdw.Row["Код_отделения"]));
                    n_pal();
                }
            }
            else
            {
                MessageBox.Show("Для изменения выбрано пустое поле! Невозможно произвести изменение!");
            }
        }

        //Удалить
        private void button9_Click(object sender, EventArgs e)
        {
            DataRowView rdw = отделенияBindingSource.Current as DataRowView;

            if (dataGridView3.SelectedRows.Count != 0)
            {
                var result = MessageBox.Show("Вы точно хотите удалить клиент из заказа?", "Подтверждение действий", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    DataGridViewRow dgvr = dataGridView3.CurrentRow;
                    Int32 ID = Convert.ToInt32(dgvr.Cells[0].Value);
                    больные1TableAdapter.Delete(ID);
                    this.больные1TableAdapter.Fill(this.больницаDataSet.Больные1, Convert.ToInt32(rdw.Row["Код_отделения"]));
                    n_pal();
                }
            }
            else
            {
                MessageBox.Show("Для удаления выбрано пустое поле! Невозможно произвести удаление!");
            }

        }
        //-------------------------------------------------------Процедуры-------------------------------------------------------------

        //Добавить

        private void button10_Click(object sender, EventArgs e)
        {
            List<string> list = new List<string>();

            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                list.Add(dataGridView3.Rows[i].Cells[2].Value.ToString());
            }
            list.Sort();

            DataRowView rdw = отделенияBindingSource.Current as DataRowView;
            Int32 pal = 0;
            Int32 fi = 0;

            Form5 f = new Form5(list);

            f.ShowDialog();

            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                if (dataGridView3.Rows[i].Cells[2].Value.ToString() == f.FIO)
                {
                    fi = Convert.ToInt32(dataGridView3.Rows[i].Cells[0].Value);
                    pal = Convert.ToInt32(dataGridView3.Rows[i].Cells[6].Value);
                }
            }

            if (f.Rc)
            {
                try
                {
                    процедуры1TableAdapter.Insert(Convert.ToInt32(rdw.Row["Код_отделения"]), pal, fi, f.Pr, f.Date, f.Cost);
                    this.процедуры1TableAdapter.Fill(this.больницаDataSet.Процедуры1, Convert.ToInt32(rdw.Row["Код_отделения"]));
                    fio();
                }
                catch (Exception)
                {
                    MessageBox.Show("Невозможно добавить данные о палате");
                }
            }

        }

        //Изменить
        private void button11_Click(object sender, EventArgs e)
        {
            List<string> list = new List<string>();

            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                list.Add(dataGridView3.Rows[i].Cells[2].Value.ToString());
            }
            list.Sort();
            DataRowView rdw = отделенияBindingSource.Current as DataRowView;



            if (dataGridView4.SelectedRows.Count != 0)
            {
                Int32 ID = Convert.ToInt32(dataGridView4.CurrentRow.Cells[0].Value);
                Int32 pal = 0;
                Int32 fi = 0;
                String FIO = dataGridView4.CurrentRow.Cells[1].Value.ToString();
                String Pr = dataGridView4.CurrentRow.Cells[2].Value.ToString();
                DateTime Date = Convert.ToDateTime(dataGridView4.CurrentRow.Cells[3].Value);
                Int32 Cost = Convert.ToInt32(dataGridView4.CurrentRow.Cells[4].Value);

                Form5 f = new Form5(list);
                f.FIO = FIO;
                f.Pr = Pr;
                f.Date = Date;
                f.Cost = Cost;

                f.ShowDialog();

                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    if (dataGridView3.Rows[i].Cells[2].Value.ToString() == f.FIO)
                    {
                        fi = Convert.ToInt32(dataGridView3.Rows[i].Cells[0].Value);
                        pal = Convert.ToInt32(dataGridView3.Rows[i].Cells[6].Value);
                    }
                }
                if (f.Rc)
                {
                    процедуры1TableAdapter.Update(Convert.ToInt32(rdw.Row["Код_отделения"]), pal, fi, f.Pr, f.Date, f.Cost, ID);
                    this.процедуры1TableAdapter.Fill(this.больницаDataSet.Процедуры1, Convert.ToInt32(rdw.Row["Код_отделения"]));
                    fio();
                }
            }
            else
            {
                MessageBox.Show("Для изменения выбрано пустое поле! Невозможно произвести изменение!");
            }
        }

        //Удалить
        private void button12_Click(object sender, EventArgs e)
        {
            DataRowView rdw = отделенияBindingSource.Current as DataRowView;

            if (dataGridView4.SelectedRows.Count != 0)
            {
                var result = MessageBox.Show("Вы точно хотите удалить клиент из заказа?", "Подтверждение действий", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    DataGridViewRow dgvr = dataGridView4.CurrentRow;
                    Int32 ID = Convert.ToInt32(dgvr.Cells[0].Value);
                    процедуры1TableAdapter.Delete(ID);
                    this.процедуры1TableAdapter.Fill(this.больницаDataSet.Процедуры1, Convert.ToInt32(rdw.Row["Код_отделения"]));
                    fio();
                }
            }
            else
            {
                MessageBox.Show("Для удаления выбрано пустое поле! Невозможно произвести удаление!");
            }

        }

        //Вывод отчетав в Word
        private void button15_Click(object sender, EventArgs e)
        {
            // Запускаем Word
            WordApp = new Word.Application();
            // Делаем Word видимым
            WordApp.Visible = true;

            WordDocuments = WordApp.Documents;
            // Добавляем документ
            WordDocument = WordDocuments.Add();
            // Получаем доступ к объекту все параграфы
            WordParagraphs = WordDocument.Content.Paragraphs;
            // Получаем доступ к объекту первый параграф
            WordParagraph = WordParagraphs[1];
            // Устанавливаем выравнивание по центру
            WordParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            // Получаем доступ к объекту выделенный участок
            WordRange = WordParagraph.Range;
            // Добавим текст в выделенный участок
            WordRange.InsertAfter("База данных таблицы\n");
            // Сделаем шрифт выделенного участка жирным
            WordRange.Font.Bold = 1;
            // Сделаем размер шрифта выделенного участка равным 16
            WordRange.Font.Size = 16;
            // Сбросим выделение участка
            WordRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            // Сейчас выделенным участком будет пустой участок в конце текста
            WordRange = WordParagraph.Range;
            // Добавим текст, он будет выделенным участком.
            WordRange.InsertAfter(DateTime.Now.ToLongDateString() + "\n");
            // Сделаем шрифт выделенного участка нежирным
            WordRange.Font.Bold = 0;
            // Сделаем размер шрифта выделенного участка равным 14
            WordRange.Font.Size = 14;
            int i = 0, j = 0, k = 0;


            // Получаем доступ к объекту текущая запись таблицы route
            DataRowView rdw = отделенияBindingSource.Current as DataRowView;
            Int32 ID = Convert.ToInt32(rdw.Row["Код_отделения"]);
            String Nam = rdw.Row["Название_отделения"].ToString();

            // Добавим параграф
            WordParagraph = WordParagraphs.Add();
            // Устанавливаем выравнивание по левой границе
            WordParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            // Получим доступ к выделенному участку нового параграфа
            WordRange = WordParagraph.Range;
            // Установим шрифт выделенного участка нового параграфа
            WordRange.Font.Bold = 0;
            WordRange.Font.Size = 9;
            // Добавим текст в новый параграф
            WordRange.InsertAfter("Наименование Отделения: " + Nam);


            j = 0;
            // Заполним набор данных DataTable1
            this.больные1TableAdapter.Fill(this.больницаDataSet.Больные1, ID);
            // Добавим параграф
            WordParagraph = WordParagraphs.Add();
            WordParagraph = WordParagraphs.Add();
            // Устанавливаем выравнивание по левой границе
            WordParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            // Получим доступ к выделенному участку нового параграфа
            WordRange = WordParagraph.Range;
            Word.Table table = WordDocument.Tables.Add(WordRange, dataGridView3.RowCount + 1, 5);
            table.Range.Cells.Borders.InsideColor = Word.WdColor.wdColorBlack;
            table.Range.Cells.Borders.OutsideColor = Word.WdColor.wdColorBlack;
            string[] zag = new string[5] { "Номер палаты", "ФИО", "Дата поступления", "Дата выписки", "Диагноз" };
            table.Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleOutset;
            table.Range.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleOutset;
            table.Range.Font.Size = 10;
            table.Range.Bold = 0;

            table.Cell(1, 1).Range.Text = zag[0];
            table.Cell(1, 2).Range.Text = zag[1];
            table.Cell(1, 3).Range.Text = zag[2];
            table.Cell(1, 4).Range.Text = zag[3];
            table.Cell(1, 5).Range.Text = zag[4];

            // Цикл по записям таблицы DataTable1
            for (больные1BindingSource.MoveFirst(); j < больные1BindingSource.Count; больные1BindingSource.MoveNext())
            {
                DataRowView drv = больные1BindingSource.Current as DataRowView;
                for (int u = 0; u < dataGridView2.Rows.Count; u++)
                {
                    if (dataGridView2.Rows[u].Cells[0].Value.ToString() == drv.Row["Код_палаты"].ToString())
                    {
                        table.Cell(j + 2, 1).Range.Text = dataGridView2.Rows[u].Cells[1].Value.ToString();
                    }
                }
                table.Cell(j + 2, 2).Range.Text = drv.Row["ФИО_больного"].ToString();
                table.Cell(j + 2, 3).Range.Text = (Convert.ToDateTime(drv.Row["Дата_поступления"])).ToShortDateString();
                table.Cell(j + 2, 4).Range.Text = (Convert.ToDateTime(drv.Row["Дата_выписки"])).ToShortDateString();
                table.Cell(j + 2, 5).Range.Text = drv.Row["Диагноз"].ToString();
                j++;
            } // for
            WordRange.Font.Bold = 0;
            WordRange.Font.Size = 12;

            k = 0;
            // Заполним набор данных DataTable1
            this.процедуры1TableAdapter.Fill(this.больницаDataSet.Процедуры1, ID);
            WordParagraph = WordParagraphs.Add();
            // Устанавливаем выравнивание по левой границе
            WordParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            // Получим доступ к выделенному участку нового параграфа
            WordRange = WordParagraph.Range;
            // Установим шрифт выделенного участка нового параграфа
            WordRange.Font.Bold = 0;
            WordRange.Font.Size = 12;
            // Добавим текст в новый параграф
            WordRange.InsertAfter("Процедуры:");

            // Цикл по записям таблицы DataTable1
            for (процедуры1BindingSource.MoveFirst(); k < процедуры1BindingSource.Count; процедуры1BindingSource.MoveNext())
            {
                // Получаем доступ к объекту текущая запись DataTable1
                DataRowView drv3 = процедуры1BindingSource.Current as DataRowView;
                // Получаем значения полей StopNumb, Name
                string _Name = "";
                for (int u = 0; u < dataGridView3.Rows.Count; u++)
                {
                    if (dataGridView3.Rows[u].Cells[0].Value.ToString() == drv3.Row["Код_ФИО"].ToString())
                    {
                        _Name = dataGridView3.Rows[u].Cells[2].Value.ToString();
                    }
                }
                string _pr = drv3.Row["Процедура"].ToString();
                string _cs = drv3.Row["Цена"].ToString();
                // Добавим параграф
                WordParagraph = WordParagraphs.Add();
                // Устанавливаем выравнивание по левой границе
                WordParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                // Получим доступ к выделенному участку нового параграфа
                WordRange = WordParagraph.Range;
                // Установим шрифт выделенного участка нового параграфа
                WordRange.Font.Bold = 0;
                WordRange.Font.Size = 12;
                // Добавим текст в новый параграф
                WordRange.InsertAfter("\t" + (k + 1).ToString() + ". " + _Name + "   " + _pr + "   " + _cs + " руб.");
                k++;
            } // for
            WordParagraph = WordParagraphs.Add();
            // Устанавливаем выравнивание по левой границе
            WordParagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            // Получим доступ к выделенному участку нового параграфа
            WordRange = WordParagraph.Range;
            // Установим шрифт выделенного участка нового параграфа
            WordRange.Font.Bold = 0;
            WordRange.Font.Size = 12;
            n_pal();
            fio();
        }

        //Вывод отчетав в Excel
        private void button14_Click(object sender, EventArgs e)
        {
            //Передать в Excel
            // Запустим Excel
            excelapp = new Excel.Application();
            // Сделаем Excel видимым
            excelapp.Visible = true;
            // В книге, которую создадим позже, будет 3 листа
            excelapp.SheetsInNewWorkbook = 3;
            // Создадим книгу
            excelapp.Workbooks.Add(Type.Missing);
            // Получаем набор ссылок на объекты Workbook (на созданные книги)
            excelappworkbooks = excelapp.Workbooks;
            //Получаем ссылку на книгу 1 - нумерация от 1
            excelappworkbook = excelappworkbooks[1];
            // Получаем ссылку на рабочие листы книги
            excelsheets = excelappworkbook.Worksheets;
            //Получаем ссылку на лист 1
            excelworksheet = (Excel.Worksheet)excelsheets[1];
            // Сделаем первый лист активным
            excelworksheet.Activate();

            DataRowView rdw = отделенияBindingSource.Current as DataRowView;
            excelcells = (Excel.Range)excelworksheet.get_Range("B3", "E3").Cells;
            excelcells.Merge(Type.Missing);
            excelcells.Value2 = rdw.Row["Название_отделения"].ToString();
            excelcells.Interior.Color = Color.PeachPuff;
            excelcells.HorizontalAlignment = Excel.Constants.xlCenter;


            excelcells = (Excel.Range)excelworksheet.get_Range("B4", "E4").Cells;
            excelcells.Merge(Type.Missing);
            excelcells.Value2 = "Больные";
            excelcells.Interior.Color = Color.LightGray;
            excelcells.HorizontalAlignment = Excel.Constants.xlCenter;

            excelcells = excelworksheet.get_Range("B5", "B5");
            excelcells.Value2 = "ФИО больного";

            excelcells = excelworksheet.get_Range("C5", "C5");
            excelcells.Value2 = "Дата поступления";

            excelcells = excelworksheet.get_Range("D5", "D5");
            excelcells.Value2 = "Дата выписки";

            excelcells = excelworksheet.get_Range("E5", "E5");
            excelcells.Value2 = "Диагноз";

            string[] c = { "B", "C", "D", "E" };
            // Цикл по строкам таблицы
            for (int i = 0; i < dataGridView3.RowCount; i++)
            {
                // Цикл по столбцам
                for (int j = 2; j < dataGridView3.ColumnCount - 2; j++)
                {
                    // Вывод в ячейку i+2, j Excel-я содержимого соответствующей ячейки
                    // dataGridView1 
                    excelcells = excelworksheet.get_Range(c[j - 2] + Convert.ToString(i + 6), c[j - 2] + Convert.ToString(i + 6));
                    if ((j == 3) | (j == 4))
                        excelcells.Value2 = (Convert.ToDateTime(dataGridView3.Rows[i].Cells[j].Value)).ToShortDateString();
                    else
                        excelcells.Value2 = dataGridView3.Rows[i].Cells[j].Value.ToString();
                } // for
            } // for



            excelcells = excelworksheet.get_Range("B3", "E" + Convert.ToSingle(dataGridView3.RowCount + 5));
            Excel.XlBordersIndex bi = Excel.XlBordersIndex.xlInsideVertical;
            excelcells.Borders[bi].LineStyle = 1;
            bi = Excel.XlBordersIndex.xlInsideHorizontal;
            excelcells.Borders[bi].LineStyle = 1;
            bi = Excel.XlBordersIndex.xlEdgeLeft;
            excelcells.Borders[bi].LineStyle = 1;
            bi = Excel.XlBordersIndex.xlEdgeTop;
            excelcells.Borders[bi].LineStyle = 1;
            bi = Excel.XlBordersIndex.xlEdgeBottom;
            excelcells.Borders[bi].LineStyle = 1;
            bi = Excel.XlBordersIndex.xlEdgeRight;
            excelcells.Borders[bi].LineStyle = 1;

            // Размер 12
            excelcells.Font.Size = 12;
            // Выравнивание по центру
            //excelcells.HorizontalAlignment = Excel.Constants.xlCenter; //-4108;

            // Выделим всю таблицу
            excelcells = excelworksheet.get_Range("B5", "E" + Convert.ToSingle(dataGridView2.RowCount + 5));
            // Подгоним ширины столбцов
            excelcells.Columns.AutoFit();

        }
    }
}
