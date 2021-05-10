using Bogus;
using ClosedXML.Report;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordTableExtractor
{
    public class Person
    {
        public string Firstname { get; set; }
        public string Lastname { get; set; }
        public string Email { get; set; }
        public DateTime? Birthday { get; set; }

        public string Name
        {
            get
            {
                return string.Join(' ', Firstname, Lastname);
            }
        }
    }

    public class Customer
    {
        public int CustNo { get; set; }
        public string Company { get; set; }

        public List<Order> Orders { get; set; } = new List<Order>();
    }

    public class Order
    {
        public int OrderNo { get; set; }
        public DateTime SaleDate { get; set; }
        public DateTime ShipDate { get; set; }
        public double ItemsTotal { get; set; }
        public double AmountPaid { get; set; }
        public List<OrderItem> Items { get; set; } = new List<OrderItem>();
    }

    public class OrderItem
    {
        public int Qty { get; set; }
        public double Discount { get; set; }
        public Part Part { get; set; }
    }

    public class Part
    {
        public double ListPrice { get; set; }
        public string Description { get; set; }
    }

    public class ReportGenerator
    {
        private string _output = @"C:\Daten\Temp\LH_Word2Excel\Report-Test\Report.xlsx";

        //private string _template = @"C:\Daten\Temp\LH_Word2Excel\Report-Test\Template2.xlsx";
        private string _template = @"C:\Users\q271043\Downloads\Subranges_Simple_tMD1.xlsx";
        //private string _template = @"C:\Daten\Temp\LH_Word2Excel\Report-Test\SimpleTemplate.xlsx";
        //private string _template = @"C:\Users\q271043\Downloads\3.xlsx";

        public void Run()
        {
            var template = new XLTemplate(_template);

            //var persons = GeneratePersons();
            //template.AddVariable(new { Persons = persons });
            //template.AddVariable("Persons", persons);

            var cust = GenerateCustomers();
            template.AddVariable(new { Customers = cust });

            template.Generate();

            //var ws = template.Workbook.Worksheets.First();
            //ws.Range("2:3").AddToNamed("TEST");

            template.SaveAs(_output);
        }

        private List<Person> GeneratePersons(int count = 10)
        {
            var persons = new List<Person>();
            var faker = new Faker();

            for(int i = 0; i < count; i++)
            {
                var firstname = faker.Name.FirstName();
                var lastname = faker.Name.LastName();

                persons.Add(new Person
                {
                    Firstname = firstname,
                    Lastname = lastname,
                    Email = $"{firstname}.{lastname}@gmail.com",
                    Birthday = faker.Date.Between(new DateTime(1980, 1, 1), new DateTime(2000, 12, 31))
                });
            }

            return persons;            
        }

        private List<Customer> GenerateCustomers(int count = 3)
        {
            var faker = new Faker();

            var customers = new List<Customer>();

            for(int i = 0; i < count; i++)
            {

                var orders = new List<Order>();
                for(int j = 0; j < count; j++)
                {
                    var items = new List<OrderItem>();
                    for(int k = 0; k < count; k++)
                    {
                        items.Add(new OrderItem { 
                            Discount = 0,
                            Qty = faker.Random.Number(1, 5),
                            Part = new Part
                            {
                                Description = faker.Lorem.Sentence(),
                                ListPrice = faker.Random.Double(10, 20)
                            }
                        });
                    }

                    orders.Add(new Order { 
                        OrderNo = j + 1,
                        AmountPaid = faker.Random.Double(10, 100),
                        ItemsTotal = faker.Random.Double(10, 100),
                        SaleDate = faker.Date.Recent(10),
                        ShipDate = faker.Date.Recent(5),
                        Items = items
                    });
                }

                customers.Add(new Customer
                {
                    CustNo = faker.Random.Number(1, 100),
                    Company = faker.Company.CompanyName(),
                    Orders = orders
                });
            }

            return customers;
        }
    }
}
