using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Xml;
using System.Web;

namespace ExtremeShell {
    public class Neighbor {
		public string LocalPort { get; set; }
		public string DeviceId { get; set; }
		public string RemotePort { get; set; }
		public string Type { get; set; }
		public string IpAddress { get; set; }
    }
}