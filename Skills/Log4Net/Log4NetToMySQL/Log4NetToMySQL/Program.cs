using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using log4net;



namespace Log4NetToMySQL {
	class Program {

		//private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		static void Main(string[] args) {
			//log4net.Config.XmlConfigurator.Configure();
			ILog log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
			log.Error("test", new Exception("test"));
			Console.WriteLine("Hit enter");
			log.Error("Hello logging world!");
			Console.ReadLine();
		}
	}
}
