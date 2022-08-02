using ExcelGenerate.DTO;
using ExcelGenerate.Helper;
using ExcelGenerate.IRepository;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelGenerate.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class PoductController : ControllerBase
    {
        private readonly IProduct IRep;
        public PoductController(IProduct IRep)
        {
            this.IRep = IRep;
        }

        [HttpGet]
        [Route("GetProduct")]
        public async Task<IActionResult> GetProduct()
        {
            var data = await IRep.GetProduct();
            return await DownloadXL.GetExcel<ProductDTO>("ProductReport", data);
        }
    }
}
