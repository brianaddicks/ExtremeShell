using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Xml;
using System.Web;

namespace ExtremeShell {
	
    public class Chassis {
		public string Type { get; set; }
		public string SerialNumber { get; set; }
		public string Version { get; set; }
		public string FanStatus { get; set; }
		public int Number { get; set; }
		public string PartNumber { get; set; }
		
		public List<Slot> Slots { get; set; }
		public List<PowerSupply> PowerSupplies { get; set; }
    }
}