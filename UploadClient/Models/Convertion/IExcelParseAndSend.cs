using Microsoft.AspNetCore.Http;
using System.Threading.Tasks;

namespace UploadClient.Models.Convertion
{
    public interface IExcelParseAndSend
    {
        Task<string> ParseAndSend(IFormFile file);
    }
}
