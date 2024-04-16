using Excel.Repository;
using Microsoft.AspNetCore.Mvc;

namespace ExcelColumns.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelController : ControllerBase
    {
       private readonly ILogger<ExcelController> _logger;

        public ExcelController(ILogger<ExcelController> logger)
        {
            _logger = logger;
        }

        [HttpGet("GetColumnLetterToColumnNumber")]
        public IActionResult GetColumnLetterToColumnNumber(string columnTitle)
        {
            if (string.IsNullOrEmpty(columnTitle) )
            {
                return BadRequest("Column Title can not be empty");
            }

            var resultNumber = Tools.ConvertColumnLetterToColumnNumber(columnTitle);
            return Ok(resultNumber);
        }

        [HttpGet("GetColumnNumberToColumnTitle")]
        public IActionResult GetColumnNumberToColumnTitle(int number)
        {
            if (number <= 0)
            {
                return BadRequest("Number Column must have a value");
            }

            string resultValue = Tools.GetColumnNumberToTitleColumns(number);
            return Ok(resultValue);
        }
    }
}
