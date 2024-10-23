using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace AutoCC_Main
{
    public static class HttpToServer
	{
        public static async Task<HttpResponseMessage> Upload(NameValueCollection strings, NameValueCollection files)
        {
            string url = PathHelper.ServerAddress + "/upload_pdf";
            var formContent = new MultipartFormDataContent();
            formContent.Headers.ContentType.MediaType = "multipart/form-data";
            // Strings
            foreach (string key in strings.Keys)
            {
                string inputName = key;
                string content = strings[key];

                formContent.Add(new StringContent(content), inputName);
            }

            // Files
            foreach (string key in files.Keys)
            {
                string inputName = key;
                string fullPathToFile = files[key];

                //FileStream fileStream = File.ReadAllBytes(fullPathToFile);
                //var streamContent = new StreamContent();
                //streamContent.Headers.Add("Content-Type", "application/pdf");

                // Initialize byte array
                byte[] buff = null;
                FileStream fs = new FileStream(fullPathToFile, FileMode.Open, FileAccess.Read);
                BinaryReader br = new BinaryReader(fs);
                long numBytes = new FileInfo(fullPathToFile).Length;

                // Load input image into Byte Array
                buff = br.ReadBytes((int)numBytes);

                var fileContent = new ByteArrayContent(buff, 0, buff.Length);
                formContent.Add(fileContent, inputName, Path.GetFileName(fullPathToFile));
            }

            var myHttpClient = new HttpClient();
            var response = myHttpClient.PostAsync(url, formContent).Result;
            var responseContent = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            return response;
        }
    }
}
