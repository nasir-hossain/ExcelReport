using ExcelGenerate.DbContexts;
using ExcelGenerate.DTO;
using ExcelGenerate.IRepository;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelGenerate.Repository
{
    public class ProductRepository : IProduct
    {
        private readonly ExportContext context;
        public ProductRepository(ExportContext context)
        {
            this.context = context;
        }
        public async Task<List<ProductDTO>> GetProduct()
        {
            try
            {
                List<ProductDTO> GetProduct =await Task.FromResult((from a in context.Products
                                                               select new ProductDTO
                                                               {
                                                                   ProductId = a.ProductId,
                                                                   ProductName = a.ProductName,
                                                                   Price = a.Price,
                                                                   ProductDescription = a.ProductDescription
                                                               }).ToList());

                return GetProduct;


            }
            catch(Exception ex)
            {
                throw ex;
            }
        }
    }
}


