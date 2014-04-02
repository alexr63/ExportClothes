using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExportToExcel;
using SelectedHotelsModel;

namespace ExportClothes
{
    class Program
    {
        static void Main(string[] args)
        {
            List<ClothView> clothViews = new List<ClothView>();
            using (SelectedHotelsEntities db = new SelectedHotelsEntities())
            {
                var clothes = db.Products.Where(p => !p.IsDeleted).OfType<Cloth>() as IQueryable<Cloth>;
                foreach (Cloth cloth in clothes)
                {
                    ClothView clothView = new ClothView();
                    clothView.Id = cloth.Id;
                    clothView.Name = cloth.Name;
                    clothView.Number = cloth.Number;
                    clothView.UnitCost = cloth.UnitCost;
                    clothView.Categories = String.Join(", ", cloth.Categories.Select(c => c.Name));
                    clothView.MerchantCategory = cloth.MerchantCategory.Name;
                    clothView.Brand = cloth.Brand.Name;
                    clothView.Styles = String.Join(", ", cloth.Styles.Select(s => s.Name));
                    clothView.Departments = String.Join(", ", cloth.Departments.Select(d => d.Name));
                    clothView.Colour = cloth.Colour.Name;
                    clothView.Gender = cloth.Gender.Name;
                    clothView.Sizes = String.Join(", ", cloth.Sizes.Select(s => s.Name));
                    clothViews.Add(clothView);
                }
                CreateExcelFile.CreateExcelDocument(clothViews, "Clothes.xlsx");
            }
        }

        private partial class ClothView
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public string Number { get; set; }
            public Nullable<decimal> UnitCost { get; set; }
            public string Categories { get; set; }
            public string MerchantCategory { get; set; }
            public string Brand { get; set; }
            public string Styles { get; set; }
            public string Departments { get; set; }
            public string Colour { get; set; }
            public string Gender { get; set; }
            public string Sizes { get; set; }
        }

    }
}
