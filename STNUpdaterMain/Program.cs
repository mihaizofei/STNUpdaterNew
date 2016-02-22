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
            var repository = new DbRepository();

            var products = source.GetProducts();
            var existingProductsName = repository.GetProductsNames();

            products = products.Where(p => existingProductsName
                                                .All(pn => string.Compare(pn, p.Name, StringComparison.OrdinalIgnoreCase) != 0)).ToList();
            
            var categories = repository.GetCategories();
            var makers = repository.GetMakers();

            PopulateCategoryIds(products, categories);
            products = FilterNoCategoryItems(products);
            PopulateMakerIds(products, makers);

            repository.InsertProducts(products);

            Console.WriteLine($"{products.Count} products inserted!! All done!!!");
            Console.ReadLine();
        }

        private static List<Product> FilterNoCategoryItems(List<Product> products)
        {
            return products.Where(p => p.CategoryId != 0).ToList();
        }

        private static void PopulateMakerIds(List<Product> products, List<Maker> makers)
        {
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
