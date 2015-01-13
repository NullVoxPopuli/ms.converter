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
        private static string[] acceptableFormats = { "doc", "docx" };


        // The default POST action
        // POST: api/Word
        //public void Post([FromBody] string output) {}


        public static byte[] ReadFully(Stream input)
        {
            byte[] buffer = new byte[16 * 1024];
            using (MemoryStream ms = new MemoryStream())
            {
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    ms.Write(buffer, 0, read);
                }
                return ms.ToArray();
            }
        }

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

            //try
            //{
                // Multipart content can be read into any of the concrete implementations of 
                // the abstract MultipartStreamProvider class
                // - see http://aspnetwebstack.codeplex.com/SourceControl/changeset/view/8fda60945d49#src%2fSystem.Net.Http.Formatting%2fMultipartStreamProvider.cs
                var streamProvider = new MultipartMemoryStreamProvider();
                Request.Content.LoadIntoBufferAsync().Wait();

                var task = Request.Content.ReadAsMultipartAsync(streamProvider).ContinueWith(t =>
                {
                    MultipartMemoryStreamProvider provider = t.Result;

                    foreach (HttpContent content in provider.Contents)
                    {
                        Stream stream = content.ReadAsStreamAsync().Result;

                        //uploadedFile = streamProvider.FileData[0];
                        string fileName = content.Headers.ContentDisposition.FileName;
                        // some browsers send the file name in extra quotes
                        fileName = fileName.Replace("\"", "");
                        var file = ReadFully(stream);
                        WmlDocument doc = new WmlDocument(fileName, file);
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

                        break;
                    }
                });

                return result;
            //}
            //catch ( System.Exception e)
            //{
            //    return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, e);
            //}
        }
    }



}
