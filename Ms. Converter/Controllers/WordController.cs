using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;
using Ms.Converter.Controllers.Concerns;

namespace Ms.Converter.Controllers
{
    public class WordController : ApiController
    {

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

                    result.Content = DocumentProcessor.convert(file.LocalFileName, fileName);
                    result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                    result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                    {
                        FileName = DocumentProcessor.getTargetNameFor(fileName)
                    };

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
