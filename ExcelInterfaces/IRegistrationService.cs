namespace ExcelInterfaces
{
    public interface IRegistrationService
    {
        IStatusService StatusService { get; set; }
        bool RegisterButton(string buttonName, string functionName,string handle);
        string GetAssociatedHandle();
    }
}