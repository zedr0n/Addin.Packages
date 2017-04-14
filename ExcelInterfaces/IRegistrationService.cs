namespace ExcelInterfaces
{
    public interface IRegistrationService
    {
        bool RegisterButton(string buttonName, string functionName,string handle);
        string GetAssociatedHandle();
    }
}