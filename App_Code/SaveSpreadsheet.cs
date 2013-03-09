using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Script.Services;
using System.Web.Services;

/// <summary>
/// Summary description for SaveSpreadsheet
/// </summary>
[WebService(Namespace = "http://localhost:61070/SpreadsheetApplication")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
// To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
[System.Web.Script.Services.ScriptService]
public class SaveSpreadsheet : System.Web.Services.WebService {

    public SaveSpreadsheet () {

        //Uncomment the following line if using designed components 
        //InitializeComponent(); 
    }

    [WebMethod]
    [ScriptMethod(ResponseFormat=ResponseFormat.Json,UseHttpGet=false)]
    public void HelloWorld(string name)
    {
        using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\_Development\WriteLines2.txt"))
        {
            file.WriteLine(name);
        }
        return;
    }
    
}
