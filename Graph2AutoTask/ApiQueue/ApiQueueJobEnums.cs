using System;
using System.Collections.Generic;
using System.Text;

namespace Graph2AutoTask.ApiQueue
{
    public enum ApiQueueJobResult
    {
        QUEUE_FAILED,
        QUEUE_RETRY,
        QUEUE_SUCCESS
    }
}
