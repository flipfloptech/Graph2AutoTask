using System;
using System.Collections.Generic;
using System.Text;

namespace AutotaskPSA
{
    [AttributeUsage(AttributeTargets.Field)]
    class AutotaskIntegrationKeyAttribute : Attribute
    {
        private string _integrationkey;
        private string _description;
        public AutotaskIntegrationKeyAttribute(string IntegrationKey, string Description)
        {
            this._integrationkey = IntegrationKey;
            this._description = Description;
        }
        public virtual string IntegrationKey
        {
            get { return _integrationkey;  }
        }
        public virtual string Description
        {
            get { return _description; }
        }

    }
}
