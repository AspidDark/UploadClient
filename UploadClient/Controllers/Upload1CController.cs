using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using System.IO;
using Microsoft.AspNetCore.Hosting;
using UploadClient.Models.Files;

namespace UploadClient.Controllers
{
    public class Upload1CController : Controller
    {
        IWebHostEnvironment _appEnvironment;

        public Upload1CController(IWebHostEnvironment appEnvironment)
        {
            _appEnvironment = appEnvironment;
        }

        [HttpPost("api/Upload1CFile")]
        public async Task<IActionResult> AddFile(IFormFile uploadedFile)
        {
            if (uploadedFile != null)
            {
                // путь к папке Files
                string path = "/Files/" + uploadedFile.FileName;
                // сохраняем файл в папку Files в каталоге wwwroot
                using (var fileStream = new FileStream(_appEnvironment.WebRootPath + path, FileMode.Create))
                {
                    await uploadedFile.CopyToAsync(fileStream);
                }
                FileModel file = new FileModel { Name = uploadedFile.FileName, Path = path };

            }
            return Ok();
        }

  
    }
}
