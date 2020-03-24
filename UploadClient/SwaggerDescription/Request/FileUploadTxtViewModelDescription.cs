using Microsoft.AspNetCore.Http;
using Swashbuckle.AspNetCore.Filters;
using UploadClient.Models.Files;

namespace UploadClient.SwaggerDescription.Request
{
    public class FileUploadTxtViewModelDescription : IExamplesProvider<FileUploadViewModelTxt>
    {
        public FileUploadViewModelTxt GetExamples()
        {

            return new FileUploadViewModelTxt
            {
               Extension="txt"
            };
        }
    }
}
