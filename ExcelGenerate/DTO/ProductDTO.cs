using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelGenerate.DTO
{
    public class ProductDTO
    {
        public int ProductId { get; set; }
        public string ProductName { get; set; }
        public int Price { get; set; }
        public string ProductDescription { get; set; }
    }
}
