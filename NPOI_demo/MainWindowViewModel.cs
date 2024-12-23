using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.CommandWpf;
using Microsoft.Win32;
using System.Data;
using System.Windows.Input;

namespace NPOI_demo
{
    public class MainWindowViewModel:ViewModelBase
    {
        public ICommand ExportCommand { get; set; }
        public MainWindowViewModel() 
        {
            ExportCommand = new RelayCommand(Export);
        }

        private void Export()
        {
            var saveDialog = new SaveFileDialog
            {
                DefaultExt = ".xlsx",
                Title = "Save Excel (.xlsx)",
                Filter = "Excel Files|*.xlsx|All Files|*",
                CheckPathExists = true,
                OverwritePrompt = true,
                RestoreDirectory = true
            };
            if (saveDialog.ShowDialog() != true) return;
            // 创建DataTable
            DataTable dt = CreateFakeDataTable(30);

            // 导出到Excel文件
            string filePath = "FakeData"; // 你可以根据需要修改文件路径和名称
            NPOIHelper.ExportExcel(dt, saveDialog.FileName, filePath);
        }

        private static DataTable CreateFakeDataTable(int rowsCount)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("Age", typeof(int));
            dt.Columns.Add("Email", typeof(string));
            dt.Columns.Add("RegistrationDate", typeof(DateTime));

            Random random = new Random();
            for (int i = 0; i < rowsCount; i++)
            {
                dt.Rows.Add(
                    i + 1,
                    $"Name{i + 1}",
                    random.Next(18, 70),
                    $"email{i + 1}@example.com",
                    DateTime.Now.AddDays(-i)
                );
            }

            return dt;
        }
    }
}
