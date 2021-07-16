using Newtonsoft.Json;

namespace NomenclExcelToJson
{
    public partial class JointObject
    {
        [JsonProperty("Name")]
        public string Name { get; set; }

        [JsonProperty("IsWrongName")]
        public string IsWrongName { get; set; }

        [JsonProperty("GUID1C")]
        public string GUID1C { get; set; }

        [JsonProperty("BlankVendorCode")]
        public string BlankVendorCode { get; set; }

        [JsonProperty("NomVendorCode")]
        public string NomVendorCode { get; set; }

        [JsonProperty("ElmaVendorCode")]
        public string ElmaVendorCode { get; set; }

        [JsonProperty("Code1С")]
        public string Code1С { get; set; }

        [JsonProperty("AdditionalInfo")]
        public string AdditionalInfo { get; set; }

        [JsonProperty("NomOrderMultiplicity")]
        public string NomOrderMultiplicity { get; set; }

        [JsonProperty("NomMultiplicity")]
        public string NomMultiplicity { get; set; }

        [JsonProperty("ElmaMultiplicity")]
        public string ElmaMultiplicity { get; set; }

        [JsonProperty("MarketingBrand")]
        public string MarketingBrand { get; set; }

        [JsonProperty("MarketingCategory")]
        public string MarketingCategory { get; set; }

        [JsonProperty("MarketingResponsible")]
        public string MarketingResponsible { get; set; }

        [JsonProperty("BlankRF")]
        public string BlankRF { get; set; }

        [JsonProperty("NomRF")]
        public string NomRF { get; set; }

        [JsonProperty("SNG")]
        public string SNG { get; set; }

        [JsonProperty("Marketing")]
        public string Marketing { get; set; }

        [JsonProperty("Groups")]
        public string Groups { get; set; }

        [JsonProperty("BlankInPack")]
        public string BlankInPack { get; set; }

        [JsonProperty("BlankInRow")]
        public string BlankInRow { get; set; }

        [JsonProperty("BlankInPallet")]
        public string BlankInPallet { get; set; }

        [JsonProperty("ElmaInPack")]
        public string ElmaInPack { get; set; }

        [JsonProperty("ElmaInRow")]
        public string ElmaInRow { get; set; }

        [JsonProperty("ElmaInPallet")]
        public string ElmaInPallet { get; set; }

        [JsonProperty("Piece")]
        public string Piece { get; set; }

        [JsonProperty("Weight")]
        public string Weight { get; set; }

        [JsonProperty("NomInPack")]
        public string NomInPack { get; set; }

        [JsonProperty("NomInRow")]
        public string NomInRow { get; set; }

        [JsonProperty("NomInPallet")]
        public string NomInPallet { get; set; }
    }
}