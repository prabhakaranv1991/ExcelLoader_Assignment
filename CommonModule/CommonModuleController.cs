using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Unity;

namespace CommonModule
{
    [ExcludeFromCodeCoverage]
    public class CommonModuleController
    {
        IUnityContainer _container = new UnityContainer();

        public void RegisterType<TFrom, TTo>(string key = null) where TTo : TFrom
        {
            if (string.IsNullOrWhiteSpace(key))
                _container.RegisterType<TFrom, TTo>();
            else
                _container.RegisterType<TFrom, TTo>(key);
        }

        public TType ResolveType<TType>(string key)
        {
            return _container.Resolve<TType>(key);
        }

    }
}
