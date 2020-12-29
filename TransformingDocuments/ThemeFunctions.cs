using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.SharePoint.Client;
using System.Text.RegularExpressions;
using OfficeDevPnP.Core;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;

namespace TransformingDocuments
{
    public static partial class ThemeFunctions
    {
        // Sample Azure Function call URL:
        //      https://<host_domain>/api/swaptheme?name=Original_Transform_Presentation
        //      where: swaptheme is this function
        //             name= is a query string containing the filename (w/o extenstion) of the presentation to process.
        //             <host_domain> is the domain for your Azure Function like: my_tenant.azurewebsites.net
        [FunctionName("swaptheme")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = "{name}")] HttpRequest req,
            ILogger log)
        {
            string name = req.Query["name"];
            if (name == null)
                return new BadRequestObjectResult("No presentation name provided.");

            if (name.ToLower().Contains("template"))
                return new BadRequestObjectResult("Templates are not processed.");

            // Logging is helpful when debugging locally.
            log.LogInformation("C# HTTP trigger function processed a request for file:" + name + ".pptx.");

            #region access_sharepoint_site
            // IMPORTANT: Replace the SPRootURL string with your Sharepoint site library root.
            string SPRootURL = @"https://jebosoft.sharepoint.com/sites/InteropEvents";
            string FileURL = @"/Shared Documents/" + name + ".pptx";
            string templateFileURL = @"/Shared Documents/Microsoft 365 DevDays Presentation Template.potx";

            // This Azure Function uses Azure AD app-only permissions to access the Sharepoint document list.
            // Read about setting up and granting Azure AD app-only permissions here:
            // https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azuread
            // 
            // IMPORTANT: the .pfx file below is not included in this repo. Use the above resource to generate this
            // certificate and then import the resulting .pfx into the Visual Studio project as a resource.
            // Remember to change the name here to your resource name.
            var assembly = Assembly.GetExecutingAssembly();
            var allResourceNames = assembly.GetManifestResourceNames();
            var resourceName = "TransformingDocuments.JeboSoft.pfx";
            X509Certificate2 x509 = null;

            using (Stream pfx = assembly.GetManifestResourceStream(resourceName))
            using (StreamReader reader = new StreamReader(pfx))
            {
                var certbytes = new byte[pfx.Length];
                pfx.Read(certbytes, 0, certbytes.Length);
                // IMPORTANT: The 2nd parameter in the X509Certificate2 ctor is the password you used to generate the 
                // certificate with the above resource. Remember to replace this password.
                x509 = new X509Certificate2(certbytes, "25MmFjQLy4An");
            }

            ClientContext clientContext = new AuthenticationManager().GetAzureADAppOnlyAuthenticatedContext(SPRootURL, "c819dd6f-7552-44a1-b7e9-a550ef4f9369", "jebosoft.onmicrosoft.com", x509);
            #endregion

            #region retrieve_files
            Web web = clientContext.Web;
            clientContext.Load(web, website => website.ServerRelativeUrl);
            clientContext.ExecuteQuery();
            Regex regex = new Regex(SPRootURL, RegexOptions.IgnoreCase);
            string strSiteRelavtiveURL = regex.Replace(FileURL, string.Empty);
            string strSiteRelavtiveTemplateURL = regex.Replace(templateFileURL, string.Empty);
            string strServerRelativeURL = CombineUrl(web.ServerRelativeUrl, strSiteRelavtiveURL);
            string strServerRelativeTemplateURL = CombineUrl(web.ServerRelativeUrl, strSiteRelavtiveTemplateURL);
            Microsoft.SharePoint.Client.File oFile = web.GetFileByServerRelativeUrl(strServerRelativeURL);
            Microsoft.SharePoint.Client.File oTemplateFile = web.GetFileByServerRelativeUrl(strServerRelativeTemplateURL);
            clientContext.Load(oFile);
            clientContext.Load(oTemplateFile);
            #endregion

            #region open_file_streams

            // Download file from SP online
            ClientResult<Stream> stream = oFile.OpenBinaryStream();
            ClientResult<Stream> templatestream = oTemplateFile.OpenBinaryStream();
            clientContext.ExecuteQuery();
            string strSiteRelavtiveURL1 = null;
            using (Stream memstrPresentation = new MemoryStream())
            using (Stream memstrTemplate = new MemoryStream())
            {
                CopyStream(stream.Value, memstrPresentation, oFile.Length);
                CopyStream(templatestream.Value, memstrTemplate, oTemplateFile.Length);
                memstrPresentation.Position = 0;
                memstrTemplate.Position = 0;
                #endregion

                #region swaptheme
                StreamReader reader = new StreamReader(memstrPresentation);
                StreamReader readertemplate = new StreamReader(memstrTemplate);
                string data = reader.ReadToEnd();
                //
                // Step 1.
                // Open the existing PPTX document as a PresentationDocument object
                // Open the template POTX as a PresentationDocument object
                //
                using (PresentationDocument presentationFile = PresentationDocument.Open(memstrPresentation, true))
                using (PresentationDocument templateFile = PresentationDocument.Open(memstrTemplate, false))
                {
                    ApplyThemeToPresentation(presentationFile, templateFile);
                }
                #endregion
                #region upload_fixed_presentation
                //
                // Step 8.
                // Upload the "fixed" presentation file to SP online with "_fixed" added to the name.
                //
                memstrPresentation.Position = 0;

                FileCreationInformation newfile = new FileCreationInformation();
                newfile.Url = name + "_fixed.pptx";
                newfile.ContentStream = memstrPresentation;
                newfile.Overwrite = true;

                List library = clientContext.Web.Lists.GetByTitle("Documents");
                Microsoft.SharePoint.Client.File uploadFile = library.RootFolder.Files.Add(newfile);

                //Update the metadata for a field having name "DocType"
                uploadFile.ListItemAllFields["DocType"] = "Updated Theme";

                uploadFile.ListItemAllFields.Update();
                clientContext.Load(library);
                clientContext.Load(uploadFile);
                clientContext.ExecuteQuery();
                #endregion
            }

            string responseMessage = string.IsNullOrEmpty(name)
                ? "swaptheme executed successfully but is missing a query string. Pass a name in the query string or in the request body for a personalized response."
                : $"swaptheme executed successfully on {name}.";
            return new OkObjectResult(responseMessage);
        }
    }
}
