using AutotaskPSA;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Text;

namespace Graph2AutoTask.ApiQueue
{
    public class ApiQueueJob
    {
        Action<Dictionary<string, object>> _internalTask = null;
        private Dictionary<string, object> _internalArguments = new Dictionary<string, object>();
        private string _internalID = null;
        private bool _internalAlertable = false;
        private ApiQueueException _internalException = null;
        private UInt64 _internalMaxRetries = 3;
        private TimeSpan _internalRetryDelay = new TimeSpan(0,0,30);

        public ApiQueueJob(string ID, Action<Dictionary<string, object>> Task, Dictionary<string, object> Arguments)
        {
            _internalID = new String(ID);
            _internalTask = new Action<Dictionary<string, object>>(Task);
            _internalArguments = new Dictionary<string, object>(Arguments);
        }

        public ApiQueueJob(string ID, Action<Dictionary<string, object>> Task, Dictionary<string, object> Arguments, UInt64 MaxRetries, TimeSpan RetryDelay) : this(ID, Task, Arguments)
        {
            _internalMaxRetries = MaxRetries;
            _internalRetryDelay = RetryDelay;
        }
        public ApiQueueJob(string ID, Action<Dictionary<string, object>> Task, Dictionary<string, object> Arguments, UInt64 MaxRetries, TimeSpan RetryDelay, bool Alertable) : this(ID, Task, Arguments, MaxRetries, RetryDelay)
        {
            _internalAlertable = Alertable;
        }
        private UInt64 _internalRetries = 0;
        public UInt64 MaxRetries { get { return _internalMaxRetries; } }
        public TimeSpan RetryDelay { get { return (_internalRetryDelay*(_internalRetries > 0 ? _internalRetries : 1 )); } }
        public UInt64 RetryCount { get { return _internalRetries; } }
        public Dictionary<string,object> Arguments { get { return _internalArguments; } }
        public ApiQueueException Exception { get { return _internalException; } }
        public string ID { get { return _internalID; } }
        public Action<Dictionary<string, object>> Task { get { return _internalTask; } }
        public bool Alertable { get { return _internalAlertable; } }
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
