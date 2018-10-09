using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace RequestExample {
	class Program {
		static string url = "https://jsonplaceholder.typicode.com/todos/1";

		static void Main(string[] args) {
			string text;
			/* #1 HttpWebRequest */
			HttpWebRequest http = (HttpWebRequest)WebRequest.Create(url);
			WebResponse response = http.GetResponse();
			using(StreamReader streamReader = new StreamReader(response.GetResponseStream(), Encoding.UTF8)) {
				text = streamReader.ReadToEnd();
			}
			response.Close();

			/* #2 WebClient */
			using(var client = new WebClient()) {
				text = client.DownloadString(url);
			}

			/* #3 Async HttpWebRequest */
			HttpWebRequestAsync().GetAwaiter().GetResult();

			/* #4 HttpRequest */
			HttpClientRequest().GetAwaiter().GetResult();
			Console.ReadLine();
		}

		private static async Task HttpWebRequestAsync() {
			HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
			HttpWebResponse response = (HttpWebResponse)await request.GetResponseAsync();
			using(Stream stream = response.GetResponseStream()) {
				using(StreamReader reader = new StreamReader(stream)) {
					Console.WriteLine(reader.ReadToEnd());
				}
			}
			response.Close();
		}

		private static async Task HttpClientRequest() {
			HttpClient ht = new HttpClient();
			HttpResponseMessage response = await ht.GetAsync(url);
			if(response.IsSuccessStatusCode) {
				Console.WriteLine(response.Content.ReadAsStringAsync().Result);
			}
		}
	}
}