using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Graph2AutoTask.ApiQueue
{
    class ApiQueueManager
    {
        int _maxthreads = Environment.ProcessorCount;
        int _currentthreads = 0;
        private Queue<ApiQueueJob> _jobs = new Queue<ApiQueueJob>();
        private CancellationToken? _internalToken = null;
        private CancellationToken _defaultToken = new CancellationToken(canceled: false);
        private readonly ILogger<MailMonitorWorker> _logger = null;
        private readonly OpsGenieApi.OpsGenieClient _opsGenieClient = null;
        private readonly MailboxConfig _configuration = null;
        public ApiQueueManager()
        {

        }
        public ApiQueueManager(CancellationToken? Token, MailboxConfig Configuration, ILogger<MailMonitorWorker> Logger, OpsGenieApi.OpsGenieClient OpsGenie)
        {
            if (Configuration == null || Logger == null || OpsGenie == null)
                throw new ArgumentNullException();
            _configuration = Configuration;
            _logger = Logger;
            _opsGenieClient = OpsGenie;
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
                _logger.LogInformation($"[{_configuration.MailBox}] - Enqueued Job {newJob.ID} at: {DateTimeOffset.Now}");
                if (_currentthreads < _maxthreads)
                {
                    _currentthreads++;
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
                        if (_currentthreads > 0)
                            _currentthreads--;
                        break;
                    }
                    _job = _jobs.Dequeue();
                }
                try
                {
                    if (_job.RetryCount > 0)
                    {
                        _logger.LogInformation($"[{_configuration.MailBox}] - Delaying Job {_job.ID} at: {DateTimeOffset.Now} for: {_job.RetryDelay} reason: QUEUE_RETRY_DELAY");
                        System.Threading.Thread.Sleep(_job.RetryDelay);
                    }
                    switch (_job.Execute())
                    {
                        case ApiQueueJobResult.QUEUE_SUCCESS:
                            //dequeue item
                            _logger.LogInformation($"[{_configuration.MailBox}] - Dequeued Job {_job.ID} at: {DateTimeOffset.Now} reason: QUEUE_SUCCESS");
                            break;
                        case ApiQueueJobResult.QUEUE_RETRY:
                            //leave for retry // wait
                            _jobs.Enqueue(_job);
                            _logger.LogInformation($"[{_configuration.MailBox}] - Requeued Job {_job.ID} at: {DateTimeOffset.Now} reason: QUEUE_RETRY[{_job.RetryCount}]");
                            break;
                        case ApiQueueJobResult.QUEUE_FAILED:
                            //we retried x times, over x time and it still failed, 
                            _logger.LogInformation($"[{_configuration.MailBox}] - Dequeued Job {_job.ID} at: {DateTimeOffset.Now} reason: QUEUE_FAIL_MAXRETRY");
                            if (_job.Alertable)
                            {
                                try
                                {
                                    _opsGenieClient.Raise(new OpsGenieApi.Model.Alert()
                                    {
                                        Alias = _job.ID,
                                        Source = "AzureTicketProcessor",
                                        Message = $"There has been a critical failure in {_job.Task.Method.Name} reason: QUEUE_FAIL_MAXRETRY"
                                    }).GetAwaiter().GetResult();
                                    _logger.LogInformation($"[{_configuration.MailBox}] - Job {_job.ID} Sent OpsGenie Alert at: {DateTimeOffset.Now} for task: {_job.Task.Method.Name}");
                                }
                                catch
                                {
                                    _logger.LogInformation($"[{_configuration.MailBox}] - Job {_job.ID} Failed to Send OpsGenie Alert at: {DateTimeOffset.Now} for task: {_job.Task.Method.Name}");
                                }
                            }
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
