using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Script.Services;
using System.Web.Services;
using System.Web.Script.Serialization;
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
    private const string file_location = "C:\\Saves\\";

    [WebMethod]
    [ScriptMethod(ResponseFormat = ResponseFormat.Json, UseHttpGet = false)]
    public bool Save(string name, string saveData)
    {
        try
        {
            File.WriteAllText(string.Format("{0}{1}.shane", file_location, name), saveData);
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
            var loadPath = Directory.GetFiles(file_location, "*.shane").ToList().FirstOrDefault(file => Path.GetFileNameWithoutExtension(file) == fileName);
            return System.IO.File.ReadAllText(loadPath);
        }
        catch (Exception)
        {
            return "";
        }
    }

    [WebMethod]
    [ScriptMethod(ResponseFormat = ResponseFormat.Json, UseHttpGet = false)]
    public string GetList()
    {
        try
        {
            var saved_files = new List<string>();
            Directory.GetFiles(file_location, "*.shane").ToList().ForEach(file =>
            {
                saved_files.Add(Path.GetFileNameWithoutExtension(file));
            });
            return new JavaScriptSerializer().Serialize(saved_files);
        }
        catch (Exception)
        {
            return "";
        }
    }
}