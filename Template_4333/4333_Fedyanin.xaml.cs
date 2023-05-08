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

namespace Template_4333
{
    /// <summary>
    /// Логика взаимодействия для _4333_Fedyanin.xaml
    /// </summary>
    public partial class _4333_Fedyanin : Window
    {
        public _4333_Fedyanin()
        {
            InitializeComponent();
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
    }
}
