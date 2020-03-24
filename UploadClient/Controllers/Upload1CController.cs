using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using System.IO;
using UploadClient.Models.Files;
using UploadClient.Models.Convertion;
using UploadClient.Contracts.Response;

namespace UploadClient.Controllers
{
    [Produces("application/octet-stream")]
    public class Upload1CController : Controller
    {
        private readonly IConvertToExcel _convertToExcel;
        public Upload1CController(IConvertToExcel convertToExcel)
        {
            _convertToExcel = convertToExcel;
        }

        [HttpPost("api/Upload1CFile")]
        [ProducesResponseType(typeof(Stream), 200)]
        [ProducesResponseType(typeof(ErrorResponse), 400)]
        public async Task<IActionResult> AddFile(FileUploadViewModelTxt model)
        {
            var file = model.File;

            if (file.Length > 0)
            {

             //   var response = _convertToExcel.Convert(model.File);

                Stream stream =  _convertToExcel.Convert(model.File);

                if (stream == null)
                    return BadRequest(new ErrorResponse()); // returns a NotFoundResult with Status404NotFound response.




                return File(stream, "application/octet-stream"); // returns a FileStreamResult
                //return Ok();
            }
            return BadRequest(new ErrorResponse());
        }

  
    }
}
