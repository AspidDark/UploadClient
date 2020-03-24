using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using UploadClient.Contracts.Response;
using UploadClient.Models.Files;

namespace UploadClient.Controllers
{
    [Produces("application/json")]
    public class UploadExcelController : Controller
    {
        [HttpPost("api/UploadExcel")]
        [ProducesResponseType(typeof(IActionResult), 200)]
        [ProducesResponseType(typeof(ErrorResponse), 400)]
        public async Task<IActionResult> Upload(FileUploadViewModelXls model)
        {

            if (model.File.Length > 0)
            {

                return Ok();
            }
            return BadRequest(new ErrorResponse());
        }

    }
}
