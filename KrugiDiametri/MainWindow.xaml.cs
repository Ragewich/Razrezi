using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using ClosedXML.Excel;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using Path = System.IO.Path;



namespace KrugiDiametri
{


    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

        }

        private void OpenExcelButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel(*.xlsx)|*.xlsx";
            if (openFileDialog.ShowDialog() == true)
            {
                ExcelFilePathBox.Text = openFileDialog.FileName;
                SavePath.Text = openFileDialog.FileName.Replace(openFileDialog.SafeFileName, "");
            }
        }

        private void SelectSaveButton_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog saveDialog = new FolderBrowserDialog();
            saveDialog.Description = "Выберите папку куда требуется сохранить файл";
            saveDialog.RootFolder = Environment.SpecialFolder.Desktop;

            if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                SavePath.Text = saveDialog.SelectedPath;
            }
        }

        private async void CalcCirclesButton_Click(object sender, RoutedEventArgs e)
        {
            ProgressBar.Value = 0;
            XLWorkbook workbook;
            try
            {
                workbook = new XLWorkbook(@$"{ExcelFilePathBox.Text}");
            }
            catch (Exception exception)
            {
                MessageBox.Show("Выбранный файл эксель в данный момент открыт на этом или другом компьютере\nЗакройте файл или перейдите в режим только для чтения перед запуском процедуры\n" +
                                $"Error: {exception.Message}",
                    "Ошибка открытия файла", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.Cancel);
                return;
            }

            ProgressBar.IsIndeterminate = true;

            await Task.Run(() =>
            {
                var ws = workbook.Worksheets.First(a => a.Position == int.Parse(NList.Text));
                var circles = new List<Circle>();
                foreach (var row in ws.RowsUsed())
                {
                    if (row.RowNumber() == 1)
                    {
                        continue;
                    }
                    var cellValue = row.Cell("E").GetString();
                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        double value = double.Parse(cellValue, System.Globalization.CultureInfo.InvariantCulture);
                        if (value % 1 == 0)
                        {
                            value += 0.1;
                        }
                        if (value < 100)
                        {
                            value *= 10;
                        }
                        int intValue = (int)Math.Round(value); // round to the nearest integer
                        circles.Add(new Circle(intValue, row.Cell("A").GetString()));
                    }
                }

                circles.Sort((c1, c2) => c2.Diametr.CompareTo(c1.Diametr));

                Dictionary<int, int> userCircles = new Dictionary<int, int>();
                var diameters = circles.DistinctBy(a => a.Diametr).Select(a => a.Diametr).ToList();
                foreach (var diameter in diameters)
                {
                    userCircles[diameter] = circles.Count(a => a.Diametr == diameter);
                }

                // Create an instance of CirclePacking
                CirclePacking cp = new CirclePacking(sheetW: int.Parse(SheetW.Text), sheetH: int.Parse(SheetH.Text), userCircles: userCircles, weldingW: int.Parse(WeldingW.Text));

                // Perform circle packing
                cp.Packing(Status, ws, circles, int.Parse(NBox.Text));

                // Draw sheets with circles
                for (int idx = 0; idx < cp.Sheets.Count; idx++)
                {
                    string sheetFilename = Path.Combine(SavePath.Text, $"sheet_{idx}.svg");
                    var image = cp.DrawSheetWithCircles(idx, scale: 0.01f, filename: sheetFilename);
                    // Преобразование объекта Bitmap в BitmapImage
                    BitmapImage bitmapImage = new BitmapImage();
                    using (MemoryStream memory = new MemoryStream())
                    {
                        image.Save(memory, System.Drawing.Imaging.ImageFormat.Bmp);
                        memory.Position = 0;
                        bitmapImage.BeginInit();
                        bitmapImage.StreamSource = memory;
                        bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                        bitmapImage.EndInit();
                    }
                    //Почему-то не сохраняет txt файл, что нужно исправить?
                    //SaveCircleCentersToFile(cp.Sheets[idx], Path.Combine(SavePath.Text, $"sheet_{idx}_centers.txt"));

                    Dispatcher.Invoke(() =>
                    {
                        ImageContainer.Source = bitmapImage;
                    });

                    string centersFilename = $"sheet_{idx}_centers.txt";
                    string path = Path.Combine(SavePath.Text, centersFilename);
                    try
                    {
                        Console.WriteLine($"Попытка сохранить файл: {path}"); // Добавьте эту строку
                        SaveCircleCentersToFile(cp.Sheets[idx], path);
                        Console.WriteLine($"Файл успешно сохранён: {path}"); // И эту строку
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при сохранении файла: {ex.Message}");
                    }

                }

                Dispatcher.Invoke(() =>
                {
                    MessageBox.Show("Готово!", "Операция выполнена", MessageBoxButton.OK, MessageBoxImage.Information,
                        MessageBoxResult.OK);
                    Status.Content = ("Process completed!");
                    ProgressBar.IsIndeterminate = false;
                    ProgressBar.Value = 100;
                });
            });
        }

        static void SaveCircleCentersToFile(Sheet sheet, string filename)
        {
            using (StreamWriter writer = new StreamWriter(filename))
            {
                writer.WriteLine($"{sheet.Width/10} {sheet.Height/10}");
                foreach (Circle circle in sheet.Circles)
                {
                    double diameter = circle.Diametr > 100 ? circle.Diametr / 10.0 : circle.Diametr / 1.0;
                    writer.WriteLine($"\"{circle.Name}\" {circle.Cx/10 } {circle.Cy /10} {diameter.ToString().Replace(',', '.')}");
                }
            }
        }

        private void SelectInfoButton_Click(object sender, RoutedEventArgs e)
        {
         

            MainWindow mainWindow = new MainWindow();
            MessageBox.Show("Привет, это небольшая инструкция по созданию разреза" + Environment.NewLine + Environment.NewLine +
                "1.Необходимо выбрать Excel файл, в excel файле нужно добавить новые листы с разрезами, можно посмотреть образец, нажав кнопку 'Образец'"
                + Environment.NewLine + "2.Выбрать куда сохранить файл, лучше всего в C:/Autodesk/(по умолчанию так и сохраняет)" + Environment.NewLine +
                "Выбрать ширину и высотув лотка" + Environment.NewLine + "и нажать Просчитать координаты" + Environment.NewLine + "3. N-обходов - это точность, своего рода. Чем > тем лучше, но дольше " + Environment.NewLine +
                "4. Нумерация листа, как в Excel, то есть нужен второй лист." + Environment.NewLine +
                "5.Далее необходимо запустить CAD приложение и написать CREATESHAPES или нажать кнопку Razrezi" + Environment.NewLine + 
                "Если команда CREATESHAPES не работает, то нужно подгрузить lisp, для этого нужно написать appload и найти папку с программой, там будет main.lsp");


        }

        private void OpenObrazecButton_Click(object sender, RoutedEventArgs e)
        {
        
            string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;

    
            string exeDirectory = System.IO.Path.GetDirectoryName(exePath);

            // Добавляем дополнительную папку в путь
            string targetDirectory = System.IO.Path.Combine(exeDirectory, "ОБРАЗЕЦ");

            // Открываем папку в проводнике
            System.Diagnostics.Process.Start("explorer.exe", targetDirectory);
        }
        

    }
}
