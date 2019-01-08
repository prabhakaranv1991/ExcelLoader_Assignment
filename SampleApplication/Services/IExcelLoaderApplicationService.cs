using SampleApplication.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SampleApplication.Services
{
    public interface IExcelLoaderApplicationService
    {
        void SaveExcelToSQL(string filePath);
        ObservableCollection<ExcelLoaderClientModel> GetComoditityData();
    }
}
