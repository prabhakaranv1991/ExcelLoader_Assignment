using CommonModule;
using ExcelLoaderRepository;
using ExcelLoaderService;
using ExcelLoaderService.Interface;
using SampleApplication.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Unity;

namespace SampleApplication
{
    public class ExcelLoaderModuleController : CommonModuleController
    {
        public ExcelLoaderModuleController()
        {
            RegisterType<IExcelLoaderApplicationService, ExcelLoaderApplicationService>();
            RegisterType<IService, Service>();
            RegisterType<IExcelDataLoaderRepository, ExcelDataLoaderRepository>();
        }

        public  TType Resolve<TType>(string key)
        {
            return ResolveType<TType>(key);
        }
    }
}
