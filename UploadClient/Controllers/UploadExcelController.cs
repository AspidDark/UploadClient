using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using System.IO;
using UploadClient.Models.Files;

namespace UploadClient.Controllers
{
    public class UploadExcelController : Controller
    {
        [HttpPost("api/Upload")]
        public async Task<IActionResult> Upload(FileUploadViewModel model)
        {
            var file = model.File;

            if (file.Length > 0)
            {
                string path = Path.Combine(@"/Files/", "uploadFiles");
                using (var fs = new FileStream(Path.Combine(path, file.FileName), FileMode.Create))
                {
                    await file.CopyToAsync(fs);
                }

                model.source = $"/uploadFiles{file.FileName}";
                model.Extension = Path.GetExtension(file.FileName).Substring(1);
            }
            return BadRequest();
        }

    }
}
