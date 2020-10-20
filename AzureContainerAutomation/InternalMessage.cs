using System;
using System.Collections.Generic;
using System.Text;

namespace AzureContainerAutomation
{
    public class InternalMessage
    {
        public string ID { get; set; }
       public  Microsoft.Graph.Message Message { get; set; }
    }
}
