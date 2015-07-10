using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Xml;
using System.Web;

namespace ExtremeShell {
	
    public class Slot {
		public int Number { get; set; }
		
		public string Model { get; set; }
		public string PartNumber { get; set; }
		public string SerialNumber { get; set; }
		
		public string HardwareVersion { get; set; }
		public string FirmwareVersion { get; set; }
		public string BootCodeVersion { get; set; }
		public string BootPromVersion { get; set; }
		
		public string Class { get; set; }
		public List<OptionModule> OptionModules { get; set; }
    }
}