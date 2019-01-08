using CommonModule.Domain.Entity;
using ExcelLoaderService.Interface;
using Microsoft.Office.Interop.Excel;
using SampleApplication.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;

namespace SampleApplication.Services
{
    public class ExcelLoaderApplicationService : IExcelLoaderApplicationService
    {
        IService _doaminService;
        public ExcelLoaderApplicationService(IService doaminService)
        {
            _doaminService = doaminService;
        }

        public ObservableCollection<ExcelLoaderClientModel> GetComoditityData()
        {
            ObservableCollection<ExcelLoaderClientModel> commoditityList = new ObservableCollection<ExcelLoaderClientModel>();
            try
            {
                IList<ExcelDataLoader> commoditities = _doaminService.GetComoditityData();
                ConvertDomainToClient(ref commoditityList,commoditities);
                return commoditityList;
            }
            catch(Exception ex)
            {
                throw ex;
            }
            
        }

        [ExcludeFromCodeCoverage]
        private void ConvertDomainToClient(ref ObservableCollection<ExcelLoaderClientModel> commoditityList, IList<ExcelDataLoader> commoditities)
        {
            foreach(var commoditity in commoditities)
            {
                commoditityList.Add(new ExcelLoaderClientModel()
                {
                    AllMonthLimit = commoditity.AllMonthLimit,
                    AnyOneMonthLimit = commoditity.AnyOneMonthLimit,
                    CommodityCode = commoditity.CommodityCode,
                    DiminishingBalanceContract = commoditity.DiminishingBalanceContract,
                    ExpiryMonthLimit = commoditity.ExpiryMonthLimit,
                    ValidFrom = commoditity.ValidFrom
                });
            }
        }

        [ExcludeFromCodeCoverage]
        public void SaveExcelToSQL(string filePath)
        {
            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook excelBook = excelApp.Workbooks.Open(filePath);
            try
            {
                IList<ExcelLoaderClientModel> excelData = new List<ExcelLoaderClientModel>();
                Worksheet excelSheet = (Worksheet)excelBook.Worksheets.get_Item(1);
                Range excelRange = excelSheet.UsedRange;

                for (int i = 2; i <= excelSheet.UsedRange.Rows.Count; i++)
                {
                    double value;
                    DateTime dtValue;
                    excelData.Add(new ExcelLoaderClientModel
                    {
                        CommodityCode = (excelSheet.Cells[i, 1].Text).ToString(),
                        DiminishingBalanceContract = (excelSheet.Cells[i, 2].Text).ToString(),
                        ExpiryMonthLimit = Double.TryParse(excelSheet.Cells[i, 3].Text, out value) ? value : 0,
                        AllMonthLimit = Double.TryParse(excelSheet.Cells[i, 4].Text, out value) ? value : 0,
                        AnyOneMonthLimit = Double.TryParse(excelSheet.Cells[i, 5].Text, out value) ? value : 0,
                        ValidFrom = DateTime.TryParse(excelSheet.Cells[i, 6].Text, out dtValue) ? dtValue : DateTime.Now
                    });
                }

                _doaminService.SaveExcelToSQL(ConvertToDTO(excelData));

            }
            catch(Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (excelBook != null)
                {
                    excelBook.Close();
                    excelApp.Quit();
                }

            }
        }

        [ExcludeFromCodeCoverage]
        private IList<ExcelDataLoader> ConvertToDTO(IList<ExcelLoaderClientModel> excelData)
        {
            IList<ExcelDataLoader> excelDto = new List<ExcelDataLoader>();

            foreach(var data in excelData)
            {
                excelDto.Add(new ExcelDataLoader()
                {
                    AllMonthLimit = data.AllMonthLimit,
                    AnyOneMonthLimit = data.AnyOneMonthLimit,
                    CommodityCode = data.CommodityCode,
                    DiminishingBalanceContract = data.DiminishingBalanceContract,
                    ExpiryMonthLimit = data.ExpiryMonthLimit,
                    ValidFrom = data.ValidFrom
                });
            }
            return excelDto;
        }
    }
}
