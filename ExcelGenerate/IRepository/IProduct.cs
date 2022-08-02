using ExcelGenerate.DTO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelGenerate.IRepository
{
    public interface IProduct
    {
        public Task<List<ProductDTO>> GetProduct();
    }
}
