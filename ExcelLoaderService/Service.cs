using CommonModule.Domain.Entity;
using ExcelLoaderRepository;
using ExcelLoaderService.Interface;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLoaderService
{
    [ExcludeFromCodeCoverage]
    public class Service : IService
    {
        IExcelDataLoaderRepository _excelRepository;
        public Service(IExcelDataLoaderRepository excelRepository)
        {
            _excelRepository = excelRepository;
        }

        public IList<ExcelDataLoader> GetComoditityData()
        {
            return _excelRepository.GetComoditityData();
        }

        public void SaveExcelToSQL(IList<ExcelDataLoader> excelData)
        {
            _excelRepository.SaveExcelToSQL(excelData);
        }
    }
}
