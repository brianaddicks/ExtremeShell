using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Xml;
using System.Web;

namespace ExtremeShell {
	
    public class Port {
		public string Name;
		public string Alias;
		public string OperStatus;
		public bool Enabled;
		public string Speed;
		public string Duplex;
		public string Type;
    }
}