using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;

using JetBrains.Annotations;
using Microsoft.Win32;
using OfficeOpenXml;

namespace Task1 {

    public sealed class MainViewModel : INotifyPropertyChanged {
        
        private bool _openExcelAfterGeneration;
        
        public ICommand GenerateExcelCommand { get; }
        public ICommand CloseCommand { get; }

        public MainViewModel() {
            GenerateExcelCommand = new RelayCommand(GenerateExcel);
            CloseCommand = new RelayCommand(CloseWindow);
        }

        private void GenerateExcel() {
            try {
                var saveFileDialog = new SaveFileDialog {
                    Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                    FilterIndex = 1,
                    FileName = "PC_Info.xlsx"
                };
                
                if (saveFileDialog.ShowDialog() != true) {
                    return;
                }

                string fileName = saveFileDialog.FileName;

                var fileInfo = new FileInfo(fileName);

                using (var excelPackage = new ExcelPackage(fileInfo)) {
                    const string sheetName = "PC Info";
                    ExcelWorksheet existingWorksheet = excelPackage.Workbook.Worksheets[sheetName];
                    if (existingWorksheet != null) {
                        excelPackage.Workbook.Worksheets.Delete(sheetName);
                    }

                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(sheetName);

                    worksheet.Cells["A1"].Value = "OS Version";
                    worksheet.Cells["B1"].Value = Environment.OSVersion.VersionString;

                    worksheet.Cells["A2"].Value = "Computer Name";
                    worksheet.Cells["B2"].Value = Environment.MachineName;

                    worksheet.Cells["A3"].Value = "Processor Type";
                    worksheet.Cells["B3"].Value = $"{Environment.GetEnvironmentVariable("PROCESSOR_IDENTIFIER")}";

                    worksheet.Cells["A4"].Value = "RAM";
                    worksheet.Cells["B4"].Value = $"{(Environment.WorkingSet / 1024) / 1024} MB";

                    worksheet.Cells["A5"].Value = "Screen Resolution";
                    worksheet.Cells["B5"].Value = $"{SystemParameters.PrimaryScreenWidth}x{SystemParameters.PrimaryScreenHeight}";

                    excelPackage.Save();
                }
                ShowMessage("Success!",$"Excel file saved successfully:\n{fileName}");

                
                switch (OpenExcelAfterGeneration) {
                    case true when File.Exists(fileName):
                        Process.Start(new ProcessStartInfo(fileName) { UseShellExecute = true });
                        break;
                    case true:
                        ShowMessage("Error!","Error: Unable to open saved file.");
                        break;
                }

            } catch (Exception ex) {
                ShowMessage("Error!",$"Error: {ex.Message}");
            }
        }
        
        private void CloseWindow() {
            Application.Current.MainWindow?.Close();
        }

        public bool OpenExcelAfterGeneration {
            get => _openExcelAfterGeneration;
            set {
                if (_openExcelAfterGeneration == value) {
                    return;
                }
                _openExcelAfterGeneration = value;
                OnPropertyChanged(nameof(OpenExcelAfterGeneration));
            }
        }
        
        private static void ShowMessage(string title,string message) {
            var viewModel = new MessageDialogViewModel(message);

            var dialog = new MessageDialog {
                DataContext = viewModel
            };

            var window = new Window {
                Title = title,
                Content = dialog,
                SizeToContent = SizeToContent.WidthAndHeight,
                Margin = new Thickness(0),
                Padding = new Thickness(0),
                ResizeMode = ResizeMode.NoResize,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = Application.Current.MainWindow
            };

            window.ShowDialog();
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        [NotifyPropertyChangedInvocator]
        private void OnPropertyChanged([CallerMemberName] string? propertyName = null) {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
