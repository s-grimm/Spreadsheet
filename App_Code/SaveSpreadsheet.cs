using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Web.Script.Services;
using System.IO;

    /// <summary>
    /// Summary description for SpreadsheetOperations
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    [System.Web.Script.Services.ScriptService]
public class SaveSpreadsheet : System.Web.Services.WebService
    {
        public const string SAVE_DIRECTORY = @"C:\Saves\";

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json, UseHttpGet = false)]
        public bool Save(string name, string saveData)
        {
            try
            {
                System.IO.File.WriteAllText(string.Format("{0}{1}.jss", SAVE_DIRECTORY, name), saveData);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json, UseHttpGet = false)]
        public string Load(string fileName)
        {
            try
            {
                string saveData;
                string loadPath = Directory.GetFiles(SAVE_DIRECTORY, "*.jss").ToList().FirstOrDefault(x => Path.GetFileNameWithoutExtension(x) == fileName);
                saveData = System.IO.File.ReadAllText(loadPath);
                return saveData;
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json, UseHttpGet = false)]
        public string GetSavedFilesList()
        {
            try
            {
                List<string> savedFiles = new List<string>();
                Directory.GetFiles(SAVE_DIRECTORY, "*.jss").ToList().ForEach(x =>
                {
                    savedFiles.Add(Path.GetFileNameWithoutExtension(x));
                });
                System.Web.Script.Serialization.JavaScriptSerializer jSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                return jSerializer.Serialize(savedFiles);
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }
    }