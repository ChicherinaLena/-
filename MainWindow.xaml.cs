using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Parser
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            try
            {
                InitLongTable();
                InitShortTable();
            }
            catch
            {
                MessageBoxResult result = MessageBox.Show("Загрузить файл?", "Файл не обнаружен!!!", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                switch (result)
                {
                    case MessageBoxResult.Yes:
                        WebClient myWebClient = new WebClient();
                        myWebClient.DownloadFile("https://bdu.fstec.ru/files/documents/thrlist.xlsx", Environment.CurrentDirectory + "/thrlist.xlsx");
                        InitLongTable();
                        InitShortTable();
                        break;
                    case MessageBoxResult.No:
                        
                        Application.Current.Shutdown();

                        break;
                }

            }
        }

        public void InitLongTable()
        {
            var metrics = ExcelClass.EnumerateMetrics(Environment.CurrentDirectory + "/thrlist.xlsx").ToList();
            TableLong.ItemsSource = metrics;
        }
        public void InitShortTable()
        {
            var shortmetrics = ExcelClass.EnumerateMetricsShort(Environment.CurrentDirectory + "/thrlist.xlsx").ToList();
            TableShort.ItemsSource = shortmetrics;
        }


        private void Forward_Click(object sender, RoutedEventArgs e)
        {
            Back.IsEnabled = true;
            Metric.rowscount++;
            InitLongTable();
            if (Metric.maxrowscount == false)
            {
                Forward.IsEnabled = false;
            }
        }

        private void Back_Click(object sender, RoutedEventArgs e)
        {
            Forward.IsEnabled = true;
            Metric.rowscount--;
            InitLongTable();
            Metric.maxrowscount = true;
            if (Metric.rowscount == 0)
            {
                Back.IsEnabled = false;
            }
        }
        private void TableShort_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var firstSelectedCellContent = this.TableShort.Columns[0].GetCellContent(this.TableShort.SelectedItem);
            var firstSelectedCell = firstSelectedCellContent != null ? firstSelectedCellContent.Parent as DataGridCell : null;
            string s = ExcelClass.Find(Environment.CurrentDirectory + "/thrlist.xlsx", Convert.ToString(firstSelectedCell).Substring(42));
            MessageBox.Show(s);
            
        }
    }
}
