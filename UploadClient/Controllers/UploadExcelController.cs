using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using UploadClient.Contracts.Response;
using UploadClient.Models.Convertion;
using UploadClient.Models.Files;

namespace UploadClient.Controllers
{
    [Produces("application/json")]
    public class UploadExcelController : Controller
    {
        private readonly IExcelParseAndSend _excelParseAndSend;
        public UploadExcelController(IExcelParseAndSend excelParseAndSend)
        {
            _excelParseAndSend = excelParseAndSend;
        }
        [HttpPost("api/UploadExcel")]
        [ProducesResponseType(typeof(IActionResult), 200)]
        [ProducesResponseType(typeof(ErrorResponse), 400)]
        public async Task<IActionResult> Upload(FileUploadViewModelXls model)
        {

            if (model.File.Length > 0)
            {
                var response = await _excelParseAndSend.ParseAndSend(model.File);
                return Ok();
            }
            return BadRequest(new ErrorResponse());
        }

    }
}
