using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Data.Sql;

namespace WebScrapper
{
    class Program
    {
        static void Main(string[] args)
        {
            SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\Administrator\source\repos\WebScrapper\WebScrapper\Database1.mdf;Integrated Security=True");
            IWebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl("https://www.daraz.pk/");
            var element = driver.FindElement(By.XPath("//*[@id=\"q\"]"));
            element.SendKeys("Mobile");
            element.Submit();
            //var price = driver.FindElements(By.ClassName("c3gUW0"));
            var items = driver.FindElements(By.ClassName("c16H9d"));
            var sales = driver.FindElements(By.ClassName("c15YQ9"));
            string rating = "";
            string price = "";
            string name = "";
            string sku = "";

            for (int i = 1; i <= items.Count; i++)
            {
                driver.FindElement(By.XPath("//*[@id=\"root\"]/div/div[2]/div[1]/div/div[1]/div[2]/div["+ i +"]/div/div/div[2]/div[2]")).Click();
                name = driver.FindElement(By.ClassName("pdp-mod-product-badge-title")).Text.ToString();
                sku = driver.FindElement(By.XPath("//*[@id=\"module_product_detail\"]/div/div/div[3]/div[1]/ul/li[2]/div")).Text.ToString();
                if (sku == null)
                {
                    sku = driver.FindElement(By.XPath("//*[@id=\"module_product_detail\"]/div/div[1]/div[3]/div[1]/ul/li[2]/div")).Text.ToString();
                }
                else if (sku == "") 
                {
                    sku = driver.FindElement(By.XPath("//*[@id=\"module_product_detail\"]/div/div[1]/div[3]/div[1]/ul/li[2]/div")).Text.ToString();
                    //*[@id="module_product_detail"]/div/div[1]/div[3]/div[1]/ul/li[2]/div
                }
                else 
                { }
                price = driver.FindElement(By.ClassName("pdp-price")).Text.ToString();
                rating = driver.FindElement(By.ClassName("count")).Text.ToString();
                Console.WriteLine(name);
                Console.WriteLine(price);
                Console.WriteLine(sku);
                if(rating == null || rating == "")
                {
                    Console.WriteLine(0);
                }
                else
                {
                    Console.WriteLine(rating);
                }
                driver.Navigate().Back();

                string Query = "INSERT INTO [Table](Sku,Name,Price,Rating) VALUES('"+sku+ "', '" + name + "', '" + price + "', '" + rating + "')";
                SqlCommand cmd = new SqlCommand(Query, con);
                SqlDataReader myreader;
                try 
                {
                    con.Open();
                    myreader = cmd.ExecuteReader();
                    Console.WriteLine("Data Saved");
                    while (myreader.Read()) { }
                }
                catch(Exception a)
                {
                    Console.WriteLine(a.Message);
                }
                con.Close();
            } 
            /*foreach(var item in items)
            {
                str = item.Text.ToString();
                Console.WriteLine(str);
                driver.FindElement(By.LinkText(str)).Click();
                driver.Navigate().Back();
            }*/
            //foreach(var p in price)
            //{
            //    Console.WriteLine(p.Text);
            //    file.WriteLine(p.Text);
            //}
            //foreach (var item in items)
            //{
            //    Console.WriteLine(item.Text);
            //    file.WriteLine(item.Text);
            //}
           
            
        }
    }
}
