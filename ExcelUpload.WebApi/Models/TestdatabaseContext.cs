using System;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata;

namespace ExcelUpload.WebApi.Models
{
    public partial class TestdatabaseContext : DbContext
    {
        public TestdatabaseContext()
        {
        }

        public TestdatabaseContext(DbContextOptions<TestdatabaseContext> options)
            : base(options)
        {
        }

        public virtual DbSet<EmployeeAddress> EmployeeAddress { get; set; }
        public virtual DbSet<Employees> Employees { get; set; }
        public virtual DbSet<ExcelTable> ExcelTable { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (!optionsBuilder.IsConfigured)
            {
  //#warning To protect potentially sensitive information in your connection string, you should move it out of source code. See http://go.microsoft.com/fwlink/?LinkId=723263 for guidance on storing connection strings.
                  //optionsBuilder.UseSqlServer("Server=172.16.0.241;Database=Testdatabase;user id=traininguser;password=Sevina@123;MultipleActiveResultSets=True;");
            }
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<EmployeeAddress>(entity =>
            {
                entity.HasKey(e => e.AddressId);

                entity.Property(e => e.City)
                    .IsRequired()
                    .HasMaxLength(50);

                entity.Property(e => e.Country)
                    .IsRequired()
                    .HasMaxLength(50);

                entity.Property(e => e.State)
                    .IsRequired()
                    .HasMaxLength(50);

                entity.HasOne(d => d.Employee)
                    .WithMany(p => p.EmployeeAddress)
                    .HasForeignKey(d => d.EmployeeId)
                    .HasConstraintName("FK_EmployeeAddress_Employees");
            });

            modelBuilder.Entity<Employees>(entity =>
            {
                entity.HasKey(e => e.EmployeeId);

                entity.Property(e => e.Dob).HasColumnType("datetime");

                entity.Property(e => e.FirstName)
                    .IsRequired()
                    .HasMaxLength(50);

                entity.Property(e => e.Gender)
                    .IsRequired()
                    .HasMaxLength(10)
                    .IsFixedLength();

                entity.Property(e => e.LastName)
                    .IsRequired()
                    .HasMaxLength(50);
            });

            modelBuilder.Entity<ExcelTable>(entity =>
            {
                entity.HasKey(e => e.ExcelId);

                entity.Property(e => e.Address).HasMaxLength(100);

                entity.Property(e => e.Age).HasMaxLength(50);

                entity.Property(e => e.Name).HasMaxLength(50);

                entity.Property(e => e.PatientId).HasMaxLength(50);
            });

            OnModelCreatingPartial(modelBuilder);
        }

        partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
    }
}
