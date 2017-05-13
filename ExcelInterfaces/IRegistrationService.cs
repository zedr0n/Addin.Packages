using IoC;

namespace ExcelInterfaces
{
    public interface IRegistrationService : IInjectable
    {
        IStatusService StatusService { get; set; }
        bool RegisterButton(string buttonName, string functionName,string handle);
        string GetAssociatedHandle();
    }
}