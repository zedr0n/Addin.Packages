using System;
using System.Linq;
using System.Reflection;
using IoC;

namespace ExcelInterfaces
{
    public class PublicFactory<TPublic> : IFactory<TPublic> where TPublic : class
    {
        private readonly IContainerService _containerService;
        protected PublicFactory(IContainerService containerService)
        {
            _containerService = containerService;
        }

        public TPublic Bind(string name, object instance)
        {
            var oObject = (IPublicObject)Activator.CreateInstance(typeof(TPublic), name, instance);

            foreach (var prop in oObject.GetType()
                .GetRuntimeProperties()
                .Where(p => typeof(IInjectable).GetTypeInfo().IsAssignableFrom(p.PropertyType.GetTypeInfo())))

            {
                var service = _containerService.GetInstance(prop.PropertyType);
                prop.SetValue(oObject, service);
            }

            return oObject as TPublic;
        }
    }
}