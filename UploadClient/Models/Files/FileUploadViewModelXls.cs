using Microsoft.AspNetCore.Http;

namespace UploadClient.Models.Files
{
    public class FileUploadViewModelXls
    {
        public IFormFile File { get; set; }
        public string Extension { get; set; }
    }
}
