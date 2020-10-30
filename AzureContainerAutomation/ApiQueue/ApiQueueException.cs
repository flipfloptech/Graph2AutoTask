using System;
using System.Collections.Generic;
using System.Text;

namespace Graph2AutoTask.ApiQueue
{
    public class ApiQueueException : Exception
    {
        public ApiQueueException()
        {
        }
        public ApiQueueException(string _exceptionMessage) : base(_exceptionMessage)
        {
        }
        public ApiQueueException(string _exceptionMessage, Exception _innerException) : base (_exceptionMessage,_innerException)
        {
        }
    }
}
