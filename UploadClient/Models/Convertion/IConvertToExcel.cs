using Microsoft.AspNetCore.Http;
using System.IO;

namespace UploadClient.Models.Convertion
{
    public interface IConvertToExcel
    {
        public byte[] Convert(IFormFile formFile);
    }
}
