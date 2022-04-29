using Microsoft.Graph;
using Azure.Identity;
using System;
using System.Threading.Tasks;
using System.Text.Json;

namespace Graph1
{
    class Program
    {
        static async Task Main()
        {
            // The client credentials flow requires that you request the
            // /.default scope, and preconfigure your permissions on the
            // app registration in Azure. An administrator must grant consent
            // to those permissions beforehand.
            var scopes = new[] { "https://graph.microsoft.com/.default" };

            // Multi-tenant apps can use "common",
            // single-tenant apps must use the tenant ID from the Azure portal
            var tenantId = "5d1a9f9d-201f-4a10-b983-451cf65cbc1e";

            // Values from app registration
            var clientId = "163155e9-8829-48d9-bdca-501ce3ff5fce";
            var clientSecret = "wXU8Q~dmOu4VLso_hDBOEq36~BTguuNH.aWZwads";

            // using Azure.Identity;
            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            // https://docs.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
            var clientSecretCredential = new ClientSecretCredential(
                tenantId, clientId, clientSecret, options);

            var graphClient = new GraphServiceClient(clientSecretCredential, scopes);


            // GET-https://graph.microsoft.com/v1.0/me/drive/items/0134CST65F5YHJOHQ2F5H3V3CT5BNNNC3H/workbook/worksheets('Arbeit')/range(address='E55')

            WorkbookRange range = await graphClient.Users["fd82005b-9ac2-4a48-ad36-e79cf3c366c7"]
                .Drive.Items["01SZPLU2I7EUC7NJFOKJA3QWJLZ2I74CEA"]
                .Workbook
                .Worksheets["Arbeit"]
                .Cell(54, 4)
                .Request()
                .GetAsync();

            
            // Update Cells

            var a = JsonDocument.Parse(@"""Frank Lampard""");

           
            WorkbookRange nRange = new WorkbookRange();
            nRange.Values = a;


            await graphClient.Users["fd82005b-9ac2-4a48-ad36-e79cf3c366c7"]
                .Drive.Items["01SZPLU2I7EUC7NJFOKJA3QWJLZ2I74CEA"]
                .Workbook
                .Worksheets["Arbeit"]                
                .Range("A27")
                .Request()
                .PatchAsync(nRange);






        }
    }
}
