using AutotaskPSA;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Text;

namespace AzureContainerAutomation.ApiQueue
{
    public class ApiQueueJob
    {
        Action<Dictionary<string, object>> _internalTask = null;
        private Dictionary<string, object> _internalArguments = new Dictionary<string, object>();
        private string _internalID = null;
        private ApiQueueException _internalException = null;
        private UInt64 _internalMaxRetries = 0;
        private TimeSpan _internalRetryDelay = new TimeSpan(0);
        
        public ApiQueueJob(string ID, Action<Dictionary<string, object>> Task, Dictionary<string, object> Arguments)
        {
            _internalID = ID;
            _internalTask = Task;
            _internalArguments = Arguments;
        }

        public ApiQueueJob(string ID, Action<Dictionary<string, object>> Task, Dictionary<string, object> Arguments, UInt64 MaxRetries, TimeSpan RetryDelay)
        {
            _internalID = ID;
            _internalTask = Task;
            _internalArguments = Arguments;
            _internalMaxRetries = MaxRetries;
            _internalRetryDelay = RetryDelay;
        }
        private UInt64 _internalRetries = 0;
        public UInt64 MaxRetries { get { return _internalMaxRetries; } }
        public TimeSpan RetryDelay { get { return _internalRetryDelay; } }
        public Dictionary<string,object> Arguments { get { return _internalArguments; } }
        public ApiQueueException Exception { get { return _internalException; } }
        public string ID { get { return _internalID; } }
        public ApiQueueJobResult Execute()
        {
            try
            {
                _internalTask(Arguments);
                return ApiQueueJobResult.QUEUE_SUCCESS;
            }
            catch (Exception _ex)
            {
                _internalRetries++;
                if (_internalRetries < MaxRetries)
                    return ApiQueueJobResult.QUEUE_RETRY;
                else
                {
                    _internalException = new ApiQueueException("Max Retries Reached", _ex);
                    return ApiQueueJobResult.QUEUE_FAILED;
                }
            }
        }
    }
}
