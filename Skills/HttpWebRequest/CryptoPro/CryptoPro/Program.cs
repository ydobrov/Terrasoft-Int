using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace CryptoPro
{
	class Program
	{
		static void Main(string[] args) {
			XmlDocument xDoc = new XmlDocument();
			xDoc.Load(@"C:\Projects\FMS\Request1.xml");
			byte[] XmlRequest = Encoding.ASCII.GetBytes(xDoc.OuterXml);
			HttpWebRequest request = (HttpWebRequest)WebRequest.Create("http://188.254.16.92:7777/gateway/services/SID0003110");
			request.Method = "POST";
			request.Credentials = CredentialCache.DefaultCredentials;
			request.ContentType = "text/xml;charset=utf-8";//; charset=Windows-1251";
			request.ContentLength = XmlRequest.Length;
			//request.Headers.Add("SOAPAction", /*имя функции*/);
			Stream dataStream = request.GetRequestStream();
			dataStream.Write(XmlRequest, 0, XmlRequest.Length);
			dataStream.Close();
			XDocument resultStream;
			try {
				WebResponse response = request.GetResponse();
				dataStream = response.GetResponseStream();
				XDocument myResult = XDocument.Load(dataStream);
				byte[] result = null;
				using (MemoryStream ms = new MemoryStream()) {
					int bufferSize = 1024;
					byte[] buffer = new byte[bufferSize];
					int bytes = 0;
					while ((bytes = dataStream.Read(buffer, 0, bufferSize)) > 0) {
						ms.Write(buffer, 0, bytes);
						ms.Flush();
					}
					result = ms.GetBuffer();
				}
			} catch (WebException ex) {
				WebException wex = (WebException)ex;
				var s = wex.Response.GetResponseStream();
				resultStream = XDocument.Load(s);
				string ss = "";
				int lastNum = 0;
				do {
					lastNum = s.ReadByte();
					ss += (char)lastNum;
				} while (lastNum != -1);
				resultStream.Save(@"C:\Projects\FMS\SMEV+FMS\SMEV+FMS\serviceResponse.xml");
				s.Close();
			}
		}
	}
}
