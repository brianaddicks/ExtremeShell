using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Xml;
using System.Web;

namespace ExtremeShell {
	
    public class OptionModule {
		public string Model { get; set; }
		public string PartNumber { get; set; }
		public string SerialNumber { get; set; }
		public string BoardRevision { get; set; }
		public string Location { get; set; }
    }
}