using CommonModule.Domain.Entity;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLoaderService.Interface
{
    public interface IService
    {
        void SaveExcelToSQL(IList<ExcelDataLoader> excelData);
        IList<ExcelDataLoader> GetComoditityData();
    }
}
