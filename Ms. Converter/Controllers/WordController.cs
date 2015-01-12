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

namespace Ms.Converter.Controllers
{
    public class WordController : ApiController
    {
        // list of formats that can be converted by this controller
        private static string[] acceptableFormats = { "doc", "docx" };


        // The default POST action
        // POST: api/Word
        //public void Post([FromBody] string output) {}


        // POST: api/Word 
        //
        // From the APS.NET Web API File Upload Sample
        // With Modifications:
        // - don't write file to disk
        // - render UnsupportedMediaType if non-word file uploaded
        public async Task<HttpResponseMessage> PostFile()
        {
            // Check if the request contains multipart/form-data.
            if (!Request.Content.IsMimeMultipartContent())
            {
                throw new HttpResponseException(HttpStatusCode.UnsupportedMediaType);
            }

            // server upload path
            string root = HttpContext.Current.Server.MapPath("~/App_Data");
            var provider = new MultipartFormDataStreamProvider(root);
            //var provider = await Request.Content.ReadAsMultipartAsync<InMemoryMultipartFormDataStreamProvider>(new InMemoryMultipartFormDataStreamProvider());


            try
            {
                // Read the form data and return an async task.
                await Request.Content.ReadAsMultipartAsync(provider);

                // This illustrates how to get the form data.
                //foreach (var key in provider.FormData.AllKeys)
                //{
                //    foreach (var val in provider.FormData.GetValues(key))
                //    {
                //        sb.Append(string.Format("{0}: {1}\n", key, val));
                //    }
                //}

                // This illustrates how to get the file names for uploaded files.
                var firstFile = provider.FileData[0];
                foreach (var file in provider.FileData)
                {
                    string fileName = UnquoteToken(file.Headers.ContentDisposition.FileName);
                    string ext = Path.GetExtension(fileName);

                    // Path.GetExtension leaves the . on the extension.
                    // this removes the dot. 
                    ext = ext.Substring(1, ext.Length - 1);

                    if (!acceptableFormats.Contains(ext))
                    {
                        throw new HttpResponseException(HttpStatusCode.UnsupportedMediaType);
                    }
                }

                FileInfo fileInfo = new FileInfo(firstFile.LocalFileName);

                HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK);
                var stream = new FileStream(firstFile.LocalFileName, FileMode.Open);
                result.Content = new StreamContent(stream);
                result.Content.Headers.ContentType =
                    new MediaTypeHeaderValue("application/octet-stream");
                return result;
                
            }
            catch (System.Exception e)
            {
                return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, e);
            }
        }

        /// <summary>
        /// Remove bounding quotes on a token if present
        /// </summary>
        /// <param name="token">Token to unquote.</param>
        /// <returns>Unquoted token.</returns>
        private static string UnquoteToken(string token)
        {
            if (String.IsNullOrWhiteSpace(token))
            {
                return token;
            }

            if (token.StartsWith("\"", StringComparison.Ordinal) && token.EndsWith("\"", StringComparison.Ordinal) && token.Length > 1)
            {
                return token.Substring(1, token.Length - 2);
            }

            return token;
        }
    }



}
