using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using MimeTypes;
using Microsoft.Extensions.Configuration;

namespace Valo.Teams.StaticFiles
{
    public class ServeStaticFile
    {   
        private const string staticFilesFolder = "wwwroot\\assets";
        private readonly string defaultPage;
        // The configuration is available for injection.
        // The used settings can be in any config (environment, host.json local.settings.json)
        public ServeStaticFile(IConfiguration configuration)
        {
            this.defaultPage = configuration.GetValue<string>("DEFAULT_PAGE", "index.html");
        }

        [FunctionName("ServeStaticFile")]
        public async Task<IActionResult> Run(
          [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = null)] HttpRequest req, ExecutionContext context, ILogger log)
        {
            try
            {
                var filePath = GetFilePath(context, req.Query["file"]);
                if (File.Exists(filePath))
                {
                    var stream = File.OpenRead(filePath);
                    return new FileStreamResult(stream, GetMimeType(filePath))
                    {
                        LastModified = File.GetLastWriteTime(filePath)
                    };
                } else
                {
                    return new NotFoundResult();
                }
            }
            catch
            {
                return new BadRequestResult();
            }
        }

        private string GetFilePath(ExecutionContext context, string pathValue)
        {
            var path = pathValue ?? "";
            string contentRoot = Path.Combine(context.FunctionAppDirectory, staticFilesFolder);
            string fullPath = Path.GetFullPath(Path.Combine(contentRoot, pathValue));
            if (!IsInDirectory(contentRoot, fullPath))
            {
                throw new ArgumentException("Invalid path");
            }

            if (Directory.Exists(fullPath))
            {
                fullPath = Path.Combine(fullPath, defaultPage);
            }
            return fullPath;
        }

        private static bool IsInDirectory(string parentPath, string childPath) => childPath.StartsWith(parentPath);
        
        private static string GetMimeType(string filePath)
        {
            var fileInfo = new FileInfo(filePath);
            return MimeTypeMap.GetMimeType(fileInfo.Extension);
        }
    }
}
