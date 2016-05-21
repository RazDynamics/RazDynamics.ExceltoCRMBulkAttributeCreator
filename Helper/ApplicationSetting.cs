using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CRMConsultants.CreateAttributes
{
    internal static class ApplicationSetting
    {
        public static EntityData SelectedEntity { get; set; }
        public static DataTable ExcelData { get; set; }
        public static string[] ColumnsListToVarify { get; set; }
        public static string FileToImport { get; set; }
    }
}
//Attribute Type	Attribute Schema Name	Option Set Values	Attribute Display Name	Description	RequiredLevel	Boolean Default  Value	String Format	String Length	Date Type Format	Integer Format	Integer Minimum Value	Integer Maximum Value	Floating Number Precision	Float Min Value	Float Max Value	Decimal Precision	Decimal Min Value	Decimal Max Value	Currency precision	Currency Min Value	Currency Max Value	 IME Mode	AuditEnable	IsValidForAdvancedFind
