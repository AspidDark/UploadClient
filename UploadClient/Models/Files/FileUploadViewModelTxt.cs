using Microsoft.AspNetCore.Http;

namespace UploadClient.Models.Files
{
    public class FileUploadViewModelTxt
    {
        public IFormFile File { get; set; }
        public string Extension { get; set; }
    }
}
