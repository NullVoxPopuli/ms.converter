using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace Ms.Converter.Controllers.Concerns
{
    public static class DocumentProcessor
    {

        // list of formats that can be converted by this controller
        private static string[] acceptableFormats = { "docx" };

        public static StringContent convert(string pathName, string sourceFileName)
        {
            // parse out the extension
            string ext = Path.GetExtension(sourceFileName);
            FileInfo fileInfo = new FileInfo(pathName);

            // Path.GetExtension leaves the . on the extension.
            // this removes the dot. 
            ext = ext.Substring(1, ext.Length - 1);

            if (!acceptableFormats.Contains(ext))
            {
                throw new HttpResponseException(HttpStatusCode.UnsupportedMediaType);
            }

            // read in the uploaded file and create an object that HTMLConverter can work with
            byte[] fileAsBytes = File.ReadAllBytes(pathName);
            WmlDocument doc = new WmlDocument(sourceFileName, fileAsBytes);


            // Custom settings / overrides
            HtmlConverterSettings settings = new HtmlConverterSettings()
            {
                // PageTitle = "some title"
            };


            // Tell HTMLConverter what to do with images, should it come across them
            // - TODO: find a way for the uploader to specify an S3 resource to upload to
            //   for when we want to display the resulting HTML online

            // notes on HTMLConverter
            // http://msdn.microsoft.com/en-us/library/office/ff628051(v=office.14).aspx
            System.Xml.Linq.XElement html = HtmlConverter.ConvertToHtml(doc, settings);

            // http://msdn.microsoft.com/en-us/library/office/ff628051(v=office.14).aspx#XHtml_Using
            // 
            // Note: the XHTML returned by ConvertToHtmlTransform contains objects of type
            // XEntity. PtOpenXmlUtil.cs defines the XEntity class. See
            // http://blogs.msdn.com/ericwhite/archive/2010/01/21/writing-entity-references-using-linq-to-xml.aspx
            // for detailed explanation.
            //
            // If you further transform the XML tree returned by ConvertToHtmlTransform, you
            // must do it correctly, or entities do not serialize properly.
            return new StringContent(html.ToStringNewLineOnAttributes());
        }

        public static string getTargetNameFor(string sourceName)
        {
            return sourceName.Replace("docx", "html");
        }

    }
}