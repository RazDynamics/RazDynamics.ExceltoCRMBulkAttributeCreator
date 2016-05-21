using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CRMConsultants.CreateAttributes
{
    internal class EntityData
    {
        public EntityData(string logicalName, string displayName,string primaryIdAttribute)
        {
            LogicalName = logicalName;
            DisplayName = displayName;
            PrimaryIdAttribute = primaryIdAttribute;
        }

        public string DisplayName { get;  set; }
        public string LogicalName { get;  set; }
        public string PrimaryIdAttribute { get; set; }
        public override string ToString()
        {
            return DisplayName;
        }
    }
}
