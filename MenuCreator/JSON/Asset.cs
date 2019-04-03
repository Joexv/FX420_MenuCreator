using System;

namespace MenuCreator.JSON
{
    public class MimeTypeConverter
    {
        public object Convert(object value)
        {
            if (value is string)
            {
                switch ((string)value)
                {
                    case "video": return "\uE116";
                    case "image": return "\uEB9F";
                    case "webpage": return "\uE128";
                    default: return "\uEB90";
                }
            }
            return new object();
        }
    }

    internal class Asset
    {
        [Newtonsoft.Json.JsonProperty(PropertyName = "asset_id")]
        public string AssetID { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "mimetype")]
        public string Mimetype { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "name")]
        public string Name { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "end_date")]
        public DateTime EndDate { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "is_enabled")]
        public Int32 isEnabled { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "is_processing")]
        public Int32? isProcessing { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "skip_asset_check")]
        public Int32 SkipAssetCheck { get; set; }

        [Newtonsoft.Json.JsonIgnore]
        public bool isEnabledSwitch
        {
            get
            {
                return isEnabled.Equals(1) ? true : false;
            }
        }

        [Newtonsoft.Json.JsonProperty(PropertyName = "nocache")]
        public Int32 NoCache { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "is_active")]
        public Int32 isActive { get; set; }

        public string _Uri;

        [Newtonsoft.Json.JsonProperty(PropertyName = "uri")]
        public string Uri
        {
            get { return _Uri; }
            set { _Uri = System.Net.WebUtility.UrlEncode(value); }
        }

        [Newtonsoft.Json.JsonIgnore]
        public string ReadableUri
        {
            get
            {
                return System.Net.WebUtility.UrlDecode(this.Uri);
            }
        }

        [Newtonsoft.Json.JsonProperty(PropertyName = "duration")]
        public string Duration { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "play_order")]
        public Int32 PlayOrder { get; set; }

        [Newtonsoft.Json.JsonProperty(PropertyName = "start_date")]
        public DateTime StartDate { get; set; }
    }
}