using OpsGenieApi.Helpers;
using Newtonsoft.Json;

namespace OpsGenieApi
{
    public class OpsGenieSerializer : IJsonSerializer
    {
        public T DeserializeFromString<T>(string json)
        {
            //provide you deserializer
            return JsonConvert.DeserializeObject<T>(json);
        }

        public string SerializeToString<T>(T data)
        {
            //provide your serializer
            return JsonConvert.SerializeObject(data);
        }
    }


}