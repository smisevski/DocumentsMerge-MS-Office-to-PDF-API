using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using Newtonsoft.Json;

namespace DocumentsMerger.Models
{
    public class DocumentsData
    {

        [JsonProperty("unique_filepath")]
        public string unique_filepath { get; set; }

        [JsonProperty("filenames")]
        public string[] filenames { get; set; }

        [JsonProperty("filepath")]
        public string filepath { get; set; }            

    }
}