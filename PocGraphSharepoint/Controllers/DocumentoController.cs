using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace PocGraphSharepoint.Controllers
{
    public class DocumentoController : Controller
    {
        // GET: Documento
        public async Task<ActionResult> Index()
        {
            var appId = ConfigurationManager.AppSettings["AppId"];
            var tenantId = ConfigurationManager.AppSettings["TenantId"];
            var clientSecret = ConfigurationManager.AppSettings["ClientSecret"];
            var tenantUrl = ConfigurationManager.AppSettings["TenantUrl"];
            var sitio = ConfigurationManager.AppSettings["Sitio"];
            var nombreBiblioteca = ConfigurationManager.AppSettings["NombreBiblioteca"];
            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(appId)
                .WithTenantId(tenantId)
                .WithClientSecret(clientSecret)
                .WithAuthority(AadAuthorityAudience.AzureAdMyOrg)
                .Build();
            ClientCredentialProvider authenticationProvider = new ClientCredentialProvider(confidentialClientApplication);
            var scopes = new[] { "https://graph.microsoft.com/.default" };
            GraphServiceClient graphClient = new GraphServiceClient(authenticationProvider);


            var sitioRaiz = await graphClient.Sites[tenantUrl].Sites[sitio]
            .Request()
            .GetAsync();

            var stream = new System.IO.MemoryStream(Encoding.UTF8.GetBytes(@"The contents of the file goes here."));

            var biblioteca = (await graphClient.Sites[sitioRaiz.Id].Drives.Request().GetAsync()).ToList().Where(l => l.Name == nombreBiblioteca).FirstOrDefault();

            //Crear carpeta
            //https://docs.microsoft.com/en-us/graph/api/driveitem-post-children?view=graph-rest-1.0&tabs=csharp
            var driveItem = new DriveItem
            {
                Name = "New Folder",
                Folder = new Folder
                {
                },
                AdditionalData = new Dictionary<string, object>()
                {
                    {"@microsoft.graph.conflictBehavior", "replace"}
                }
            };
            var fileAbsoluteUrl = "https://latinshare.sharepoint.com/sites/dev/Documentos%20compartidos/New%20Folder%203/subfolder/";
            string base64Value = System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(fileAbsoluteUrl));
            string encodedUrl = "u!" + base64Value.TrimEnd('=').Replace('/', '_').Replace('+', '-');
            //var asdId = await graphClient.Shares[encodedUrl].DriveItem.Request().GetAsync();



            var carpeta = await graphClient.Sites[sitioRaiz.Id].Drives[biblioteca.Id].Root.Children
                .Request()
                .AddAsync(driveItem);
            driveItem.Name = "SubFolder";
            var subCarpeta = await graphClient.Sites[sitioRaiz.Id].Drives[biblioteca.Id].Items[carpeta.Id].Children
                .Request()
                .AddAsync(driveItem);

            //subir archivo
            //https://docs.microsoft.com/en-us/graph/api/driveitem-put-content?view=graph-rest-1.0&tabs=http
            var sitioDrivexasd = await graphClient.Sites[sitioRaiz.Id].Drives[biblioteca.Id].Items[subCarpeta.Id]
                .ItemWithPath("asda.txt")
                .Content
                .Request()
                .PutAsync<DriveItem>(stream);

            //subir archivo >4mb

            using (var fileStream = System.IO.File.OpenRead("C:\\Users\\bpalma\\Desktop\\libros\\Developing-Windows-Azure-and-Web-Services.pdf"))
            {
                DriveItemUploadableProperties driveItemUploadableProperties = new DriveItemUploadableProperties
                {
                    //Name = "",
                    ODataType = null,
                    AdditionalData = new Dictionary<string, object>
                    {
                        { "@microsoft.graph.conflictBehavior", "replace" }
                    }
                };
                var uploadSession = await graphClient.Sites[sitioRaiz.Id].Drives[biblioteca.Id].Items[subCarpeta.Id]
                .ItemWithPath(Path.GetFileName(fileStream.Name))
                .CreateUploadSession(driveItemUploadableProperties)
                .Request()
                .PostAsync();


                // Max slice size must be a multiple of 320 KiB
                int maxSliceSize = 320 * 1024;
                var fileUploadTask =
                    new LargeFileUploadTask<DriveItem>(uploadSession, fileStream, maxSliceSize);

                // Create a callback that is invoked after each slice is uploaded
                IProgress<long> progress = new Progress<long>(prog =>
                {
                    Console.WriteLine($"Uploaded {prog} bytes of {fileStream.Length} bytes");
                });

                try
                {
                    // Upload the file
                    var uploadResult = await fileUploadTask.UploadAsync(progress);

                    if (uploadResult.UploadSucceeded)
                    {
                        // The ItemResponse object in the result represents the
                        // created item.
                        Console.WriteLine($"Upload complete, item ID: {uploadResult.ItemResponse.Id}");
                    }
                    else
                    {
                        Console.WriteLine("Upload failed");
                    }
                }
                catch (ServiceException ex)
                {
                    Console.WriteLine($"Error uploading: {ex.ToString()}");
                }



            }




            return View();
        }
    }
}