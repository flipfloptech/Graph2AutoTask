using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace AzureContainerAutomation.ApiQueue
{
    class ApiQueueManager
    {
        private Queue<ApiQueueJob> _jobs = new Queue<ApiQueueJob>();
        private bool _delegateQueuedOrRunning = false;
        private CancellationToken? _internalToken = null;
        private CancellationToken _defaultToken = new CancellationToken(canceled: false);
        public ApiQueueManager()
        {

        }
        public ApiQueueManager(CancellationToken? Token)
        {
            if (Token != null)
            {
                _internalToken = Token;
            }
            else
                _internalToken = null;
        }
        public void Enqueue(ApiQueueJob newJob)
        {
            lock (_jobs)
            {
                _jobs.Enqueue(newJob);
                if (!_delegateQueuedOrRunning)
                {
                    _delegateQueuedOrRunning = true;
                    ThreadPool.UnsafeQueueUserWorkItem(ProcessQueuedItems, _internalToken);
                }
            }
        }
        public bool HasJobWithMessageID(string ID)
        {
            bool _result = false;
            if (String.IsNullOrWhiteSpace(ID))
                return _result;
            lock (_jobs)
            {
                Parallel.ForEach<ApiQueueJob>(_jobs, (_job,_state) => { 
                    if (_job.ID.ToLower() == ID.ToLower())
                    {
                        _result = true;
                        _state.Break();
                    }
                });
                return _result;
            }
        }

        private void ProcessQueuedItems(object Token)
        {
            if (Token is CancellationToken) { _internalToken = (CancellationToken)Token; }
            while (!_internalToken.GetValueOrDefault(_defaultToken).IsCancellationRequested)
            {
                ApiQueueJob _job;
                lock (_jobs)
                {
                    if (_jobs.Count == 0)
                    {
                        _delegateQueuedOrRunning = false;
                        break;
                    }
                    _job = _jobs.Peek();
                }
                try
                {
                    switch(_job.Execute())
                    {
                        case ApiQueueJobResult.QUEUE_SUCCESS:
                            //dequeue item
                            _jobs.Dequeue();
                            break;
                        case ApiQueueJobResult.QUEUE_RETRY:
                            //leave for retry // wait
                            System.Threading.Thread.Sleep(_job.RetryDelay);
                            break;
                        case ApiQueueJobResult.QUEUE_FAILED:
                            _jobs.Dequeue();
                            //we retried x times, over x time and it still failed, 
                            break;
                    }
                }
                catch
                {
                    ThreadPool.UnsafeQueueUserWorkItem(ProcessQueuedItems, _internalToken);
                    throw;
                }
            }
        }
    }
}
