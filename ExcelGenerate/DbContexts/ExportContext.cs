using System;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata;
using ExcelGenerate.Models;

// Code scaffolded by EF Core assumes nullable reference types (NRTs) are not used or disabled.
// If you have enabled NRTs for your project, then un-comment the following line:
// #nullable disable

namespace ExcelGenerate.DbContexts
{
    public partial class ExportContext : DbContext
    {
        public ExportContext()
        {
        }

        public ExportContext(DbContextOptions<ExportContext> options)
            : base(options)
        {
        }

        public virtual DbSet<Products> Products { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (!optionsBuilder.IsConfigured)
            {
                //#warning To protect potentially sensitive information in your connection string, you should move it out of source code. See http://go.microsoft.com/fwlink/?LinkId=723263 for guidance on storing connection strings.
                optionsBuilder.UseSqlServer("Data Source=NASIR\\SQLEXPRESS; Initial Catalog=IBOS2; Connect Timeout=30; Encrypt=False; Trusted_Connection=True; ApplicationIntent=ReadWrite;");
            }
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Products>(entity =>
            {
                entity.HasKey(e => e.ProductId)
                    .HasName("PK_dbo.Products");

                entity.Property(e => e.ProductId).HasColumnName("ProductID");
            });

            OnModelCreatingPartial(modelBuilder);
        }

        partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
    }
}
