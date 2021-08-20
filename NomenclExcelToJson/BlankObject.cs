using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace NomenclExcelToJson
{
    public partial class BlankObject
    {
        [JsonProperty("Name")]
        public string Name { get; set; }

        [JsonProperty("VendorCode")]
        public string VendorCode { get; set; }

        [JsonProperty("Pallet")]
        public string Pallet { get; set; }

        [JsonProperty("Row")]
        public string Row { get; set; }

        [JsonProperty("Pack")]
        public string Pack { get; set; }

        [JsonProperty("MinSale")]
        public string Exclusive { get; set; }

        [JsonProperty("Weight")]
        public string Weight { get; set; }
    }

    public partial class BlankObject
    {
        public static List<BlankObject> FromJsonBlank(string json) => JsonConvert.DeserializeObject<List<BlankObject>>(json, Converter.Settings);
    }
}
