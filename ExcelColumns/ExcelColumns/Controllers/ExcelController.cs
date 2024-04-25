using Excel.Interface;
using Microsoft.AspNetCore.Mvc;

namespace ExcelColumns.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelController : ControllerBase
    {
        private readonly ILogger<ExcelController> _logger;
        private readonly IConvertColumn _convertColumn;

        public ExcelController(ILogger<ExcelController> logger, IConvertColumn convertColumn)
        {
            _logger = logger;
            _convertColumn = convertColumn;
        }

        [HttpGet("GetColumnLetterToColumnNumber")]
        public IActionResult GetColumnLetterToColumnNumber(string columnTitle)
        {
            if (string.IsNullOrEmpty(columnTitle))
            {
                return BadRequest("Column Title can not be empty");
            }

            var resultNumber = _convertColumn.ConvertColumnLetterToColumnNumber(columnTitle);
            return Ok(resultNumber);
        }

        [HttpGet("GetColumnNumberToColumnTitle")]
        public IActionResult GetColumnNumberToColumnTitle(int number)
        {
            if (number <= 0)
            {
                return BadRequest("Number Column must have a value");
            }

            string resultValue = _convertColumn.GetColumnNumberToTitleColumns(number);
            return Ok(resultValue);
        }
    }
}
