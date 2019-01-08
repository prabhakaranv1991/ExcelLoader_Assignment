using System;
using System.Collections.Generic;
using System.Windows.Input;
using Prism.Commands;
using System.Windows.Forms;
using SampleApplication.Model;
using Microsoft.Office.Interop.Excel;
using System.Collections.ObjectModel;
using SampleApplication.Services;
using System.ComponentModel;
using System.Linq;

namespace SampleApplication.ViewModel
{
    public class ExcelDataLoaderViewModel : INotifyPropertyChanged
    {
        public ICommand BrowseCommand { get; set; }
        public ICommand RefreshCommand { get; set; }

        ObservableCollection<ExcelLoaderClientModel> excelDataItemSocurce;

        ExcelLoaderModuleController controller;

        IExcelLoaderApplicationService _excelLoaderApplicationService;

        public event PropertyChangedEventHandler PropertyChanged;

        public ObservableCollection<ExcelLoaderClientModel> ExcelDataItemSocurce
        {
            get { return excelDataItemSocurce; }
            set { excelDataItemSocurce = value; OnPropertyChanged("ExcelDataItemSocurce"); }
        }

        private string path;
        public string Path
        {
            get { return path; }
            set { path = value; OnPropertyChanged("Path"); }
        }

        private string commodityCodeFilter;
        public string CommodityCodeFilter
        {
            get { return commodityCodeFilter; }
            set { commodityCodeFilter = value; OnPropertyChanged("CommodityCodeFilter"); }
        }

        private string validFromFilter;
        public string ValidFromFilter
        {
            get { return validFromFilter; }
            set { validFromFilter = value; OnPropertyChanged("ValidFromFilter"); }
        }

        private DateTime maxDate;
        public DateTime MaxDate
        {
            get { return maxDate; }
            set { maxDate = value; OnPropertyChanged("MaxDate"); }
        }

        private void OnPropertyChanged(string propertyName = null)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public ExcelDataLoaderViewModel()
        {
            controller = new ExcelLoaderModuleController();
            _excelLoaderApplicationService = controller.Resolve<ExcelLoaderApplicationService>("ExcelLoaderApplicationService");

            BrowseCommand = new DelegateCommand(SaveExcelData);
            RefreshCommand = new DelegateCommand(LoadExcelData);
            ExcelDataItemSocurce = new ObservableCollection<ExcelLoaderClientModel>();
            validFromFilter = DateTime.Now.ToShortDateString();
            MaxDate = DateTime.Now;
        }


        private void LoadExcelData()
        {
            var source = _excelLoaderApplicationService.GetComoditityData();

            DateTime dt;
            if (DateTime.TryParse(validFromFilter, out dt))
            {

                ExcelDataItemSocurce = new ObservableCollection<ExcelLoaderClientModel>
                        (source.Where(x =>
                                        (string.IsNullOrWhiteSpace(commodityCodeFilter) ? true : x.CommodityCode.ToUpper() == commodityCodeFilter.ToUpper())
                                        && (string.IsNullOrWhiteSpace(validFromFilter) ? true : x.ValidFrom >= Convert.ToDateTime(validFromFilter))));
            }
            else
            {
                MessageBox.Show("Choose valid Date");
            }
        }

        public void SaveExcelData()
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "(*.xlsx)|*.xlsx";
            var browseFile = fileDialog.ShowDialog();

            if (!string.IsNullOrWhiteSpace(fileDialog.FileName))
            {
                Path = fileDialog.FileName;
                _excelLoaderApplicationService.SaveExcelToSQL(fileDialog.FileName);
            }

            var result = MessageBox.Show(fileDialog.FileName, "Alert", MessageBoxButtons.OKCancel);
        }

    }
}
