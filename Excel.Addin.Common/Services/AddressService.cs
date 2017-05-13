using ExcelDna.Integration;
using ExcelInterfaces;

namespace CommonAddin
{
    public class AddressService : IAddressService
    {
        public string GetAddress()
        {
            var reference = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
            if (reference == null)
                return null;
            var cellReference = (string)XlCall.Excel(XlCall.xlfAddress, 1 + reference.RowFirst,
                1 + reference.ColumnFirst);

            var sheetName = (string)XlCall.Excel(XlCall.xlSheetNm,
                reference);
            cellReference = sheetName + cellReference;
            return cellReference;

        }
    }
}