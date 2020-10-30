using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Primitives;
using Newtonsoft.Json;
namespace Graph2AutoTask
{
   
    public class MimeTypeDatabase
    {
        internal class mimeTypeInfo
        {
            public string source=String.Empty;
            public bool compressible=false;
            public string[] extensions=null;
            public bool allowable=true;
        }
        public class MimeTypeInfo
        {
            internal string _mimeType = String.Empty;
            internal string _source = String.Empty;
            internal bool _compressible = false;
            internal string[] _extensions = Array.Empty<string>();
            internal bool _allowable = true;
            public MimeTypeInfo(string MimeType, string Source, bool Compressible, bool Allowable, string[] Extensions)
            {
                if (string.IsNullOrWhiteSpace(MimeType))
                    throw new ArgumentNullException();
                if (string.IsNullOrWhiteSpace(Source))
                    Source = "Unknown";
                _mimeType = MimeType;
                _source = Source;
                _compressible = Compressible;
                _extensions = Extensions;
                _allowable = Allowable;
            }
            public string MimeType { get { return _mimeType; } }
            public string Source { get { return _source; } }
            public bool Compressible { get { return _compressible; } }
            public string[] Extensions { get { return _extensions; } }
            public bool Allowable { get { return _allowable; } }
        }
        internal Dictionary<string, mimeTypeInfo> _internal = new Dictionary<string, mimeTypeInfo>();
        internal List<MimeTypeInfo> _external = new List<MimeTypeInfo>();
        public MimeTypeDatabase()
        {
            try
            {
                _internal = (Dictionary<string, mimeTypeInfo>)JsonConvert.DeserializeObject<Dictionary<string, mimeTypeInfo>>(System.IO.File.ReadAllText($"{System.IO.Path.GetDirectoryName(new System.Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath)}/configs/mimetypes.json"));
                _external.Clear();
                Parallel.ForEach(_internal, (_mimeType, state) =>
                 {
                     lock (_external)
                     {
                         _external.Add(new MimeTypeInfo(
                             MimeType: _mimeType.Key,
                             Source: _mimeType.Value.source,
                             Compressible: _mimeType.Value.compressible,
                             Extensions: _mimeType.Value.extensions,
                             Allowable: _mimeType.Value.allowable
                             ));
                     }
                 });
            }
            catch
            {
                // add some defaults.
                _external.Clear();
                _external.Add(new MimeTypeInfo(
                    MimeType: "image/png",
                    Source: "iana",
                    Compressible: false,
                    Extensions: new string[] { "png" },
                    Allowable: true));
            }
        }
        public MimeTypeInfo GetMimeTypeInfoFromExtension(string Extension)
        {
            List<MimeTypeInfo> _results = new List<MimeTypeInfo>();
            MimeTypeInfo _return = null;
            if (!string.IsNullOrWhiteSpace(Extension))
            {
                Parallel.ForEach(_external, (_mimetype, state) => {
                    lock (_results)
                    {
                        if (_mimetype.Extensions != null)
                        {
                            if (_mimetype.Extensions.Contains(Extension.ToLower()))
                                _results.Add(_mimetype);
                        }
                    }
                });
                foreach(MimeTypeInfo _result in _results)
                {
                    _return = _result;
                    if (_return.Allowable == false)
                        break;
                }
            }
            return _return;
        }
    }
}
