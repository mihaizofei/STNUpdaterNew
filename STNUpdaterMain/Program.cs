using System;
using System.Collections.Generic;
using System.Linq;
using STNUpdater.Models;

namespace STNUpdater
{
    class Program
    {
        [STAThread]
        static void Main()
        {
            var source = new FileSource();
            Console.WriteLine("The file is opened!");

            var repository = new DbRepository();
            var products = source.GetProducts();
            var existingProductsName = repository.GetProductsNames();

            Console.WriteLine("Eliminating existing products...");
            products = products.Where(p => existingProductsName.All(pn => string.Compare(pn, p.Name, StringComparison.OrdinalIgnoreCase) != 0)).ToList();

            var categories = repository.GetCategories();
            var makers = repository.GetMakers();
            PopulateCategoryIds(products, categories);

            products = FilterNoCategoryItems(products);

            PopulateMakerIds(products, makers);
            repository.InsertProducts(products);

            PrintFinishText(products.Count);
        }

        private static void PrintFinishText(int productsAmount)
        {
            Console.WriteLine();
            Console.WriteLine("*****************************************");
            Console.WriteLine($"{productsAmount} products inserted.");
            Console.WriteLine("All done!!! Press any key...");
            Console.ReadLine();
        }

        private static List<Product> FilterNoCategoryItems(List<Product> products)
        {
            Console.WriteLine("Filter no categories items...");
            return products.Where(p => p.CategoryId != 0).ToList();
        }

        private static void PopulateMakerIds(List<Product> products, List<Maker> makers)
        {
            Console.WriteLine("Populate makers ids...");
            products.ForEach(p =>
            {
                var maker = makers.FirstOrDefault(m => string.Equals(m.Name, p.Maker, StringComparison.CurrentCultureIgnoreCase));
                if (maker != null)
                {
                    p.MakerId = maker.Id;
                }
            });
        }

        private static void PopulateCategoryIds(List<Product> products, List<Category> categories)
        {
            Console.WriteLine("Populate categories ids...");
            products.ForEach(p =>
            {
                var category = categories.FirstOrDefault(c => string.Equals(c.Name, p.Category, StringComparison.CurrentCultureIgnoreCase));
                if (category != null)
                {
                    p.CategoryId = category.Id;
                }
            });
        }
    }
}
