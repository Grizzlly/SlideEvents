using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.IO;
using System.Runtime.CompilerServices;
using SlideEvents.Properties;
using System.Text.Json;
using System.Security.Cryptography;
using System.Net;
using System.Net.Http;
using System.Diagnostics;

namespace SlideEvents
{
    public partial class ThisAddIn
    {
        public Dictionary<int, string> urlPairs;
        private string filePath;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            
            Application.SlideShowNextSlide += Application_SlideShowNextSlide;
            Application.PresentationOpen += Application_PresentationOpen;
        }

        private void Application_PresentationOpen(PowerPoint.Presentation Pres)
        {
            string appdata = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string path = Path.Combine(appdata, "slideevents");
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            string jsonname = Pres.Name + ".json";
            string settings = Path.Combine(path, jsonname);
            if (!File.Exists(settings))
            {
                var stream = File.Create(settings);
                byte[] data = JsonSerializer.SerializeToUtf8Bytes(new Dictionary<int, string>());
                stream.Write(data, 0, data.Length);
                stream.Close();
            }

            filePath = settings;
            var readStream = File.OpenRead(settings);
            urlPairs = JsonSerializer.Deserialize<Dictionary<int, string>>(readStream);
            readStream.Close();
        }

        private async void Application_SlideShowNextSlide(PowerPoint.SlideShowWindow Wn)
        {
            int index = Wn.View.CurrentShowPosition + 1;
            if (urlPairs.TryGetValue(index, out string url))
            {
                using (HttpClient httpClient = new HttpClient())
                {
                    await httpClient.GetAsync(url);
                }    
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            var writeStream = File.Open(filePath, FileMode.Truncate);
            JsonSerializer.Serialize(writeStream, urlPairs);
            writeStream.Close();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
