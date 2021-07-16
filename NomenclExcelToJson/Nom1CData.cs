﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace NomenclExcelToJson
{
    public partial class Nom1CData
    {
        [JsonProperty("Name")]
        public string Name { get; set; }

        [JsonProperty("GUID1C")]
        public string GUID1C { get; set; }

        [JsonProperty("VendorCode")]
        public string VendorCode { get; set; }

        [JsonProperty("Code1С")]
        public string Code1С { get; set; }

        [JsonProperty("AdditionalInfo")]
        public string AdditionalInfo { get; set; }

        [JsonProperty("OrderMultiplicity")]
        public string OrderMultiplicity { get; set; }

        [JsonProperty("Multiplicity")]
        public string Multiplicity { get; set; }

        [JsonProperty("MarketingBrand")]
        public string MarketingBrand { get; set; }

        [JsonProperty("MarketingCategory")]
        public string MarketingCategory { get; set; }

        [JsonProperty("MarketingResponsible")]
        public string MarketingResponsible { get; set; }

        [JsonProperty("RF")]
        public string RF { get; set; }

        [JsonProperty("SNG")]
        public string SNG { get; set; }

        [JsonProperty("Marketing")]
        public string Marketing { get; set; }

        [JsonProperty("Groups")]
        public string Groups { get; set; }

        [JsonProperty("InPack")]
        public string InPack { get; set; }

        [JsonProperty("InRow")]
        public string InRow { get; set; }

        [JsonProperty("InPallet")]
        public string InPallet { get; set; }

        [JsonProperty("Piece")]
        public string Piece { get; set; }

        [JsonProperty("Weight")]
        public string Weight { get; set; }
    }

    public partial class Nom1CData
    {
        public static List<Nom1CData> FromJson(string json) => JsonConvert.DeserializeObject<List<Nom1CData>>(json, Converter.Settings);
    }

    public static class Serialize
    {
        public static string ToJson(this List<Nom1CData> self) => JsonConvert.SerializeObject(self, NomenclExcelToJson.Converter.Settings);
    }

    internal static class Converter
    {
        public static readonly JsonSerializerSettings Settings = new JsonSerializerSettings
        {
            MetadataPropertyHandling = MetadataPropertyHandling.Ignore,
            DateParseHandling = DateParseHandling.None,
            Converters =
            {
                new IsoDateTimeConverter { DateTimeStyles = DateTimeStyles.AssumeUniversal }
            },
        };
    }

    internal class ParseStringConverter : JsonConverter
    {
        public override bool CanConvert(Type t) => t == typeof(long) || t == typeof(long?);

        public override object ReadJson(JsonReader reader, Type t, object existingValue, JsonSerializer serializer)
        {
            if (reader.TokenType == JsonToken.Null) return null;
            var value = serializer.Deserialize<string>(reader);
            if (long.TryParse(value, out long l))
            {
                return l;
            }
            return null;
        }

        public override void WriteJson(JsonWriter writer, object untypedValue, JsonSerializer serializer)
        {
            if (untypedValue == null)
            {
                serializer.Serialize(writer, null);
                return;
            }
            var value = (long)untypedValue;
            serializer.Serialize(writer, value.ToString());
            return;
        }

        public static readonly ParseStringConverter Singleton = new ParseStringConverter();
    }
}
