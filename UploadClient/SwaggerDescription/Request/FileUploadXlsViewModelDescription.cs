using Swashbuckle.AspNetCore.Filters;
using UploadClient.Models.Files;

namespace UploadClient.SwaggerDescription.Request
{
    public class FileUploadXlsViewModelDescription : IExamplesProvider<FileUploadViewModelXls>
    {
        public FileUploadViewModelXls GetExamples()
        {
            return new FileUploadViewModelXls
            {
                Extension = "xls"
            };
        }
    }
}
