using System;
using System.Collections.Generic;
using System.Text;

namespace AutotaskPSA
{
    public enum AutotaskIntegrationKey
    {
        [AutotaskIntegrationKey("TVW2PL6MBU2YYNQ2ZIEQZSIR2A", "nCentral Embedded Integration Key")]
        nCentral,
        [AutotaskIntegrationKey("BBYW7R7JOZPKULDOD2HK3ZKWB7", "nCentral NCOD Embedded Integration Key")]
        nCentral_NCOD
    }

    public static class AutotaskIntegrationKeyExtensions
    {
        public static string GetIntegrationKey(this AutotaskIntegrationKey _AutotaskIntegrationKey)
        {
            string _result = string.Empty;
            try
            {
                AutotaskIntegrationKeyAttribute[] _attributes = (AutotaskIntegrationKeyAttribute[])_AutotaskIntegrationKey.GetType().GetField(_AutotaskIntegrationKey.ToString()).GetCustomAttributes(typeof(AutotaskIntegrationKeyAttribute), false);
                if (_attributes.Length > 0)
                {
                    if (!String.IsNullOrWhiteSpace(_attributes[0].IntegrationKey))
                    {
                        _result = _attributes[0].IntegrationKey;
                    }
                    else
                        _result = string.Empty;
                }
            }
            catch
            {
                _result = string.Empty;
            }
            return _result;
        }
    }
}
