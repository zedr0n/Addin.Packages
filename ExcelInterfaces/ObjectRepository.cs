using System.Linq;
using System.Reflection;
using IoC;

namespace ExcelInterfaces
{
    public class ObjectRepository : IObjectRepository
    {
        private readonly IContainerService _containerService;
        private readonly IRegistrationService _registrationService;
        public ObjectRepository(IContainerService containerService)
        {
            _containerService = containerService;
            _registrationService = containerService.GetInstance<IRegistrationService>();
        }

        public IPublicObject Get(string handle)
        {
            if (handle == "")
                // #RegistrationService get the handle associated with the button
                handle = _registrationService.GetAssociatedHandle();

            var oObject = ExcelInterfaces.Public.This(handle);

            foreach (var prop in oObject.GetType()
                .GetRuntimeProperties()
                .Where(p => typeof(IInjectable).GetTypeInfo().IsAssignableFrom(p.PropertyType.GetTypeInfo())))

            {
                var service = _containerService.GetInstance(prop.PropertyType);
                prop.SetValue(oObject, service);
            }

            return oObject;

        }
    }
}