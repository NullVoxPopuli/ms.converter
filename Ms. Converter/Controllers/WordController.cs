using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

namespace Ms.Converter.Controllers
{
    public class WordController : ApiController
    {
        // list of formats that can be converted by this controller
        private static string[] acceptableFormats = { "docx" };


        // The default POST action
        // POST: api/Word
        //public void Post([FromBody] string output) {}


        // POST: api/Word 
        //
        // From the APS.NET Web API File Upload Sample
        // With Modifications:
        // - don't write file to disk
        // - render UnsupportedMediaType if non-word file uploaded
        public async Task<HttpResponseMessage> Post()
        {
            HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);

            // Check if the request contains multipart/form-data.
            if (!Request.Content.IsMimeMultipartContent())
            {
                throw new HttpResponseException(HttpStatusCode.UnsupportedMediaType);
            }

            try
            {


                string root = HttpContext.Current.Server.MapPath("~/App_Data");
                var provider = new MultipartFormDataStreamProvider(root);

                // Read the form data and return an async task.
                await Request.Content.ReadAsMultipartAsync(provider);

                // This illustrates how to get the file names for uploaded files.
                var firstFile = provider.FileData[0];
                foreach (var file in provider.FileData)
                {
                    string fileName = file.Headers.ContentDisposition.FileName;
                    fileName = fileName.Replace("\"", "");
                    string ext = Path.GetExtension(fileName);
                    FileInfo fileInfo = new FileInfo(file.LocalFileName);


                    // Path.GetExtension leaves the . on the extension.
                    // this removes the dot. 
                    ext = ext.Substring(1, ext.Length - 1);

                    if (!acceptableFormats.Contains(ext))
                    {
                        throw new HttpResponseException(HttpStatusCode.UnsupportedMediaType);
                    }

                    byte[] fileAsBytes = File.ReadAllBytes(file.LocalFileName);
                    WmlDocument doc = new WmlDocument(fileName, fileAsBytes);
                    System.Xml.Linq.XElement html = HtmlConverter.ConvertToHtml(doc, new HtmlConverterSettings());

                    // http://msdn.microsoft.com/en-us/library/office/ff628051(v=office.14).aspx#XHtml_Using
                    // 
                    // Note: the XHTML returned by ConvertToHtmlTransform contains objects of type
                    // XEntity. PtOpenXmlUtil.cs defines the XEntity class. See
                    // http://blogs.msdn.com/ericwhite/archive/2010/01/21/writing-entity-references-using-linq-to-xml.aspx
                    // for detailed explanation.
                    //
                    // If you further transform the XML tree returned by ConvertToHtmlTransform, you
                    // must do it correctly, or entities do not serialize properly.
                    result.Content = new StringContent(html.ToStringNewLineOnAttributes());
                    result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");

                    // only support one file for now
                    break;
                }


                return result;
            }
            catch (System.Exception e)
            {
                return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, e);
            }
        }
    }



}
