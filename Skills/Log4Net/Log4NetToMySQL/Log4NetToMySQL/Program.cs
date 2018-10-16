using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using log4net;

namespace Log4NetToMySQL {
	class Program {

		static void Main(string[] args) {
			ILog log = LogManager.GetLogger("Custom Logger");
			log.Error("test", new Exception("Test it"));
			Console.WriteLine("Hit enter");
			log.Error("Test it again");
			Console.ReadLine();
		}
	}
}
