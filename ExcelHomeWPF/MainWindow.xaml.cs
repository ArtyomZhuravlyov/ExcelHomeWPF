using System;
using System.Windows;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelHomeWPF
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Excel.Application excelapp;

        private Excel.Workbook workbook;
        private Excel.Sheets worksheets;
        private Excel.Worksheet worksheet;
        private Excel.Range cells1;
        private Excel.Range cells2;

        //книга , листы и ячейки
        private Excel.Workbook workbooknew;
        private Excel.Worksheet worksheetnew;
        private Excel.Range cellsnew;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void ListBox_DragEnter(object sender, DragEventArgs e)
        {

            e.Effects = DragDropEffects.All;
            //DragDrop.DoDragDrop(Type.Missing, Type.Missing , DragDropEffects.Copy);


        }

        private void ListBox_Drop(object sender, DragEventArgs e)
        {
         //   image1.Visibility = Visibility.Visible;

            string[] file = (string[])e.Data.GetData(DataFormats.FileDrop);
            //ListBox1.Items.Add(System.IO.Path.GetFileName(file[0]));
            string nameFile = file[0]; //только для отображения в листбокс
            if (nameFile.EndsWith(".xlxs")) nameFile = nameFile.Replace(".xlxs", "");
            if (nameFile.EndsWith(".xls")) nameFile = nameFile.Replace(".xls", "");
            ListBox1.Items[0] = System.IO.Path.GetFileName(nameFile);

            excelapp = new Excel.Application();

            workbook = excelapp.Workbooks.Open(file[0],
                       Type.Missing, Type.Missing, Type.Missing,
                       "WWWWW", "WWWWW", Type.Missing, Type.Missing, Type.Missing,
                       Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                       Type.Missing, Type.Missing);

            worksheets = workbook.Worksheets; //приписали список всех листов выбранной книги
            worksheet = (Excel.Worksheet)worksheets.get_Item(3); //взяли конкретный лист
            
            //cells1 = worksheet.get_Range("С16", "С16"); //откуда смотрим ячейки
            cells1 = worksheet.get_Range("I16", "I420"); 
            cells2 = worksheet.get_Range("C16", "C16"); //для проверки на конец файла
            cells1.Formula = "=C16 * F16";
            int countEmpty = 0;
            int temp = 1;
            int ss = 0;

            // cells1[430, 1].Value2 = cells1[430, 1].Value2 ?? 0; //для определения пустых ячеек в них null(скорее всего)
            // MessageBox.Show(cells1[430, 1].Value2.ToString(), "пустая");

            int i=1;
            for ( i = 1; i < 430 && countEmpty < 7; i++)
            {
                //  проверка на пустые клетки(конец файла)
                if (cells2[i, 1].Value2 == null)
                {
                 //   cells1[i, 1].Value2 = cells1[i, 1].Value2 ?? 0;
                    countEmpty++;
                }
                else
                {
                    countEmpty = 0;
                }


             //   cells1[i, 1].Value2 = cells1[i, 1].Value2 ?? 0; //для определения пустых ячеек в них null(скорее всего)
                if (cells1[i, 1].Value2 < 0) cells1[i, 1].Value2 = 0; //если в ячейке #знач! то там лежит макс отрииц число для int32
            }

            cells2 = worksheet.get_Range("J421", "K422");
            cells2.Merge(Type.Missing);
            cells2.Value2 = "Итог";
            cells2.get_Characters(1, 10).Font.Size = 24;

            cells2 = worksheet.get_Range("J423", "K424");
            cells2.Merge(Type.Missing);
            //Устанавливаем цвет обводки
            cells2.Borders.ColorIndex = 3;
            //Устанавливаем стиль и толщину линии
            cells2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            cells2.Borders.Weight = Excel.XlBorderWeight.xlThin;

            //Делаем все содержимое ячейки жирным - пример форматирования данных во всей ячейки
            //cells2.EntireRow.Font.Bold = true;
            //Объект rng соответствует содержимому ячейки
            //Range rng = excelapp.ActiveCell;
            
            //Выполняем форматирование фрагментов текста
            //rng.get_Characters(start, end).Font.Bold = false;
            //rng.get_Characters(start, end).Font.Italic = false;
            cells2.get_Characters(1, 10).Font.Size = 24;
            //rng.get_Characters(start, end).Font.ColorIndex = 4;


            //cells2.NumberFormat = "$0.00";
            cells2.Formula = "=SUM(I16:I" + (i+15-8).ToString() +")";
            //cells2 = worksheet.get_Range("H16", "H16");
            //cells2.Formula = "=SUM(C16:C20)";

            // cells2 = worksheet.get_Range("J423", "J423");
            excelapp.UserControl = true;
            excelapp.Visible = true;







        }

        private void Window_Closed(object sender, EventArgs e)
        {
            try
            {
                excelapp.Quit();
                Close();
            }
            catch
            {
                Close();
            }
        }
    }
}
