using System;
using System.Collections.Generic;
using System.Text;

namespace AzureContainerAutomation.ApiQueue
{
    public enum ApiQueueJobResult
    {
        QUEUE_FAILED,
        QUEUE_RETRY,
        QUEUE_SUCCESS
    }
}
