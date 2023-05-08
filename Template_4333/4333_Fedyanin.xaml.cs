using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Numerics;
using System.IO;
using System.Text.Json;

namespace Template_4333
{
    /// <summary>
    /// Логика взаимодействия для _4333_Fedyanin.xaml
    /// </summary>
    public partial class _4333_Fedyanin : Window
    {
        private const int _sheetsCount = 3;
        public _4333_Fedyanin()
        {
            InitializeComponent();
        }
        class Employee
        {
            public int Id { get; set; }
            public string CodeStaff { get; set; }
            public string Position { get; set; }
            public string FullName { get; set; }
            public string Log { get; set; }
            public string Password { get; set; }
            public string LastEnter { get; set; }
            public string TypeEnter { get; set; }
        }
        private void import_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            string[,] list;
            Excel.Application ObjWorkExcel = new
            Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
            {
                for (int i = 0; i < _rows; i++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                }
            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using (IsrpoLr2Entities usersEntities = new IsrpoLr2Entities())
            {
                for (int i = 0; i < _rows; i++)
                {
                    usersEntities.table1.Add(new table1()
                    {
                        idEmployee = list[i, 0],
                        Post = list[i, 1],
                        FIO = list[i, 2],
                        Login = list[i, 3],
                        Password = list[i, 4],
                        LastInput = list[i, 5],
                        TypeInput = list[i, 6]
                    });
                }
                usersEntities.SaveChanges();
            }

        }
        private void export_Click(object sender, RoutedEventArgs e)
        {
            Dictionary<string, List<table1>> ByData = new Dictionary<string, List<table1>>();
            using (IsrpoLr2Entities usersEntities = new IsrpoLr2Entities())
            {
                var allWorkers = usersEntities.table1.ToList().GroupBy(w => w.Post);

                foreach (var group in allWorkers)
                {
                    ByData[group.Key] = group.ToList();
                }
            }
            var app = new Excel.Application();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            app.Visible = true;

            foreach (var worker in ByData)
            {
                string type = worker.Key;
                List<table1> workers = worker.Value;
                Excel.Worksheet worksheet = app.Worksheets.Add();

                if (type != "")
                {
                    worksheet.Name = type;

                    worksheet.Cells[1, 1] = "Код сотрудника";
                    worksheet.Cells[1, 2] = "ФИО";
                    worksheet.Cells[1, 3] = "Должность";
                }
                int rowIndex = 2;

                foreach (table1 work in workers)
                {
                    worksheet.Cells[rowIndex, 1] = work.idEmployee;
                    worksheet.Cells[rowIndex, 2] = work.FIO;
                    worksheet.Cells[rowIndex, 3] = work.Login;
                    rowIndex++;
                }
            }
        }
        private async void importJSON_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "файл Json |*.json",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            List<Employee> list;
            using (FileStream fs = new FileStream(ofd.FileName, FileMode.OpenOrCreate))
            {
                list = await JsonSerializer.DeserializeAsync<List<Employee>>(fs);
            }
            using (IsrpoLr2Entities usersEntities = new IsrpoLr2Entities())
            {
                foreach (Employee employee in list)
                {
                    usersEntities.table1.Add(new table1()
                    {
                        idEmployee = employee.CodeStaff,
                        Post = employee.Position,
                        FIO = employee.FullName,
                        Login = employee.Log,
                        Password = employee.Password,
                        LastInput = employee.LastEnter,
                        TypeInput = employee.TypeEnter
                    });
                }
                usersEntities.SaveChanges();
            }
        }
        private void exportWord_Click(object sender, RoutedEventArgs e)
        {
            List<table1> allOrder;
            using (IsrpoLr2Entities entities = new IsrpoLr2Entities())
            {
                allOrder = entities.table1.ToList();
            }
            var app = new Word.Application();
            Word.Document document = app.Documents.Add();

            for (int i = 0; i < _sheetsCount; i++)
            {
                Word.Paragraph paragraph = document.Paragraphs.Add();
                Word.Range range = paragraph.Range;
                Word.Range range1 = paragraph.Range;
                List<string[]> PostEmployee = new List<string[]>() { //for sheets name
                    new string[]{ "Продавец" },
                    new string[]{ "Администратор" },
                    new string[]{ "Старший смены" },
                };




                var data = i == 0 ? allOrder.Where(o => o.Post == "Продавец")
                        : i == 1 ? allOrder.Where(o => o.Post == "Администратор")
                        : i == 2 ? allOrder.Where(o => o.Post == "Старший смены")
                        : allOrder; //sort for task
                List<table1> currentStreet = data.ToList();
                int countStreetInCategory = currentStreet.Count();
                Word.Paragraph tableParagraph = document.Paragraphs.Add();
                Word.Range tableRange = tableParagraph.Range;
                Word.Table strettTable = document.Tables.Add(tableRange, countStreetInCategory + 1, 3);
                strettTable.Borders.InsideLineStyle =
                strettTable.Borders.OutsideLineStyle =
                Word.WdLineStyle.wdLineStyleSingle;
                strettTable.Range.Cells.VerticalAlignment =
                Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                Word.Range cellRange = strettTable.Cell(1, 1).Range;
                cellRange.Text = "Код сотрудника";

                cellRange = strettTable.Cell(1, 2).Range;
                cellRange.Text = "ФИО";
                cellRange = strettTable.Cell(1, 3).Range;
                cellRange.Text = "Логин";

                strettTable.Rows[1].Range.Bold = 1;
                strettTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                int j = 1;
                foreach (var currentStaff in currentStreet.OrderBy(a => a.Post))
                {
                    range1.Text = Convert.ToString($"Кол-во сотрудников - {currentStreet.OrderBy(a => a.Post).Count()}");
                    range1.InsertParagraphAfter();
                    cellRange = strettTable.Cell(j + 1, 1).Range;
                    cellRange.Text = $"{currentStaff.idEmployee}";
                    cellRange.ParagraphFormat.Alignment =
                    Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    cellRange = strettTable.Cell(j + 1, 2).Range;
                    cellRange.Text = $"{currentStaff.FIO}";
                    cellRange.ParagraphFormat.Alignment =
                    Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    cellRange = strettTable.Cell(j + 1, 3).Range;
                    cellRange.Text = currentStaff.Login;
                    cellRange.ParagraphFormat.Alignment =
                    Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    j++;
                }

                for (int t = 0; t < _sheetsCount; t++)
                {
                    range.Text = Convert.ToString($"Должность - {PostEmployee[i][0]}");
                    range.InsertParagraphAfter();
                    //range.InsertParagraphBefore();
                }

                if (i > 0)
                {
                    range.InsertBreak(Word.WdBreakType.wdPageBreak);
                }
            }
            app.Visible = true;
        }
    }
}
