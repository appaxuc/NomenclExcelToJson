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
    public partial class ElmaObject
    {
        [JsonProperty("Name")]
        public string Name { get; set; }

        [JsonProperty("VendorCode")]
        public string VendorCode { get; set; }

        [JsonProperty("Brand")]
        public string Brand { get; set; }

        [JsonProperty("Categories")]
        public string Categories { get; set; }

        [JsonProperty("Guid1C")]
        public string Guid1C { get; set; }

        [JsonProperty("Code1С")]
        public string Code1С { get; set; }

        [JsonProperty("Group")]
        public string Group { get; set; }

        [JsonProperty("Bestseller")]
        public string Bestseller { get; set; }

        [JsonProperty("Kratnost")]
        public string Kratnost { get; set; }

        [JsonProperty("Liquidation")]
        public string Liquidation { get; set; }

        [JsonProperty("Action")]
        public string Action { get; set; }

        [JsonProperty("Neww")]
        public string Neww { get; set; }

        [JsonProperty("NomGroup")]
        public string NomGroup { get; set; }

        [JsonProperty("VolumePer1St")]
        public string VolumePer1St { get; set; }

        [JsonProperty("Promo")]
        public string Promo { get; set; }

        [JsonProperty("Sale")]
        public string Sale { get; set; }

        [JsonProperty("ElmaUid")]
        public string ElmaUid { get; set; }

        [JsonProperty("InPallet")]
        public string InPallet { get; set; }

        [JsonProperty("InRow")]
        public string InRow { get; set; }

        [JsonProperty("InPack")]
        public string InPack { get; set; }

        [JsonProperty("Exclusive")]
        public string Exclusive { get; set; }
    }

    public partial class ElmaObject
    {
        public static List<ElmaObject> FromJsonElma(string json) => JsonConvert.DeserializeObject<List<ElmaObject>>(json, Converter.Settings);
    }
}