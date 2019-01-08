using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SampleApplication.Services;
using ExcelLoaderService.Interface;
using Rhino.Mocks;
using System.IO;
using SampleApplication.Model;
using Microsoft.Office.Interop.Excel;
using CommonModule.Domain.Entity;

namespace SimpleApplication.UnitTests.Services
{
    /// <summary>
    /// Summary description for ExcelLoaderApplicationServiceTests
    /// </summary>
    [TestClass]
    public class ExcelLoaderApplicationServiceTests
    {
        private ExcelLoaderApplicationService _excelApplicationService = null;
        private IService _doaminService = null;
        private string inputFilePath;

        IList<ExcelLoaderClientModel> expectedResult;
        IList<ExcelDataLoader> mockValue;

        [TestInitialize]
        public void Test_Initialize()
        {
            _doaminService = MockRepository.GenerateMock<IService>();

            _excelApplicationService = new ExcelLoaderApplicationService(_doaminService);
            expectedResult = new List<ExcelLoaderClientModel>();
            mockValue = new List<ExcelDataLoader>();

            var tempPath = Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory);
            inputFilePath = @"C:\Users\balaji.shanmugam01\Desktop\ExcelLoader-master\SimpleApplication.UnitTests\InputFile\Chapter5.xlsx";

            MockValues();
        }

        private void MockValues()
        {
            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook excelBook = excelApp.Workbooks.Open(inputFilePath);
            Worksheet excelSheet = (Worksheet)excelBook.Worksheets.get_Item(1);
            Range excelRange = excelSheet.UsedRange;

            for (int i = 2; i <= excelSheet.UsedRange.Rows.Count; i++)
            {
                double value;
                DateTime dtValue;
                expectedResult.Add(new ExcelLoaderClientModel
                {
                    CommodityCode = (excelSheet.Cells[i, 1].Text).ToString(),
                    DiminishingBalanceContract = (excelSheet.Cells[i, 2].Text).ToString(),
                    ExpiryMonthLimit = Double.TryParse(excelSheet.Cells[i, 3].Text, out value) ? value : 0,
                    AllMonthLimit = Double.TryParse(excelSheet.Cells[i, 4].Text, out value) ? value : 0,
                    AnyOneMonthLimit = Double.TryParse(excelSheet.Cells[i, 5].Text, out value) ? value : 0,
                    ValidFrom = DateTime.TryParse(excelSheet.Cells[i, 6].Text, out dtValue) ? dtValue : DateTime.Now
                });
            }

            if (excelBook != null)
            {
                excelBook.Close();
                excelApp.Quit();
            }

            ConvertClientToDomain(expectedResult);
        }

        private void ConvertClientToDomain(IList<ExcelLoaderClientModel> commoditityList)
        {
            foreach (var commoditity in commoditityList)
            {
                mockValue.Add(new ExcelDataLoader()
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

        [TestMethod]
        public void Test_SaveExcelToSQL()
        {
            _doaminService.Stub(x => x.GetComoditityData()).Return(mockValue);

            var actulResults = _excelApplicationService.GetComoditityData();

            Assert.AreEqual(actulResults.Count, expectedResult.Count);
            for (int i=0; i<actulResults.Count;i++)
            {
                Assert.AreEqual(actulResults[i].AllMonthLimit, expectedResult[i].AllMonthLimit);
            }

        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void Test_SaveExcelToSQL_Negative()
        {
            _doaminService.Stub(x => x.GetComoditityData()).Throw(new Exception());

            var actulResults = _excelApplicationService.GetComoditityData();


        }
    }
}
