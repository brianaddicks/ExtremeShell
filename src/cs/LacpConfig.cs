using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Xml;
using System.Web;

namespace ExtremeShell {
    public class LacpConfig {
		public int SystemPriority;
		public List<LagPort> LagPorts;
		public bool SinglePortLag;
		public bool FlowRegeneration;
		public string OutportLocalPreference;
    }
}