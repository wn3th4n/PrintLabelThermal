
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;

namespace PrintLabelThermal
{
    public class StringArrayConvert : JsonConverter<string[]>
    {
        public override string[] ReadJson(JsonReader reader, Type objectType, string[] existingValue, bool hasExistingValue, JsonSerializer serializer)
        {
            JToken token = JToken.Load(reader);
            if (token.Type == JTokenType.String)
            {
                return token.ToString().Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            }
            return new string[0];
        }

        public override void WriteJson(JsonWriter writer, string[] value, JsonSerializer serializer)
        {
            writer.WriteValue(string.Join(",", value));
        }
    }
}
