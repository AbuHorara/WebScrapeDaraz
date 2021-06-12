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
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.PageObjects;
using SeleniumExtras.WaitHelpers;

namespace WebScrapper
{
    class Program
    {
        static void Main(string[] args)
        {
            //SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\Administrator\source\repos\WebScrapper\WebScrapper\Database1.mdf;Integrated Security=True");
            IWebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl("https://www.daraz.pk/");
            var element = driver.FindElement(By.XPath("//*[@id=\"q\"]"));
            var cat = "Electronics";
            element.SendKeys(cat);
            element.Submit();
            //var price = driver.FindElements(By.ClassName("c3gUW0"));
            var items = driver.FindElements(By.ClassName("c16H9d"));
            //var sales = driver.FindElements(By.ClassName("c15YQ9"));
            string rating = "";
            string price = "";
            string name = "";
            string sku = "";
            string rat = "";
            string month = "";
            string pri = "";
            string subcat = "";
            string[] months = { "January", "Feburary", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" };
            int count = 1;
            int jan = 0, feb = 0, mar = 0, apr = 0, may = 0, jun = 0, jul = 0, aug = 0, sep = 0, oct = 0, nov = 0, dec = 0;

            for (int i = 1; i <= items.Count; i++)
            {
                try
                {
                    System.Threading.Thread.Sleep(500);
                    driver.FindElement(By.XPath("//*[@id=\"root\"]/div/div[2]/div[1]/div/div[1]/div[2]/div[" + i + "]/div/div/div[2]/div[2]")).Click();
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(300);
                    name = driver.FindElement(By.ClassName("pdp-mod-product-badge-wrapper")).Text.ToString();
                    //System.Threading.Thread.Sleep(1000);

                    //sku = driver.FindElement(By.XPath("//*[@id=\"module_product_detail\"]/div/div[1]/div[3]/div/ul/li[2]/div")).Text.ToString();
                    try
                    {
                        sku = driver.FindElements(By.ClassName("key-value"))[1].Text.ToString();
                    }
                    catch(Exception)
                    {
                        sku = driver.FindElements(By.ClassName("key-value"))[0].Text.ToString();
                    }
                    //System.Threading.Thread.Sleep(1000);
                    price = driver.FindElement(By.ClassName("pdp-price")).Text.ToString();
                    //System.Threading.Thread.Sleep(1000);
                    rating = driver.FindElement(By.ClassName("count")).Text.ToString();
                    Console.WriteLine("Name: " + name);

                    int indx1;
                    pri = price.Substring(4);
                    string pstr = "";
                    if (pri.Contains(','))
                    {
                        string p1 = (pri.Split(','))[0];
                        string p2 = (pri.Split(','))[1];
                        pstr = p1 + p2;
                        Console.WriteLine("Price: " + pstr);
                    }
                    else
                    {
                        pstr = pri;
                        Console.WriteLine("Price: " + pstr);
                    }
                    indx1 = rating.IndexOf(' ');
                    rat = rating.Substring(0, indx1 + 1);
                    int a = Convert.ToInt32(rat);
                    Console.WriteLine("Rating: " + rat);
                    //System.Threading.Thread.Sleep(1000);

                    //if (sku == null || sku == "")
                    //{
                    //    sku = driver.FindElements(By.ClassName("key-value"))[0].Text.ToString();
                    //    Console.WriteLine(sku);
                    //}
                    //else
                    //{
                    //    Console.WriteLine(sku);
                    //}
                    if (rating == null || rating == "" || rating == "0")
                    {
                        Console.WriteLine(0);
                    }
                    else if (a <= 5)
                    {
                        WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                        IWebElement SearchResult = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.XPath("//*[@id=\"module_product_review\"]/div/div[3]/div[1]/div[" + count + "]/div[1]/span")));
                        month = driver.FindElement(By.XPath("//*[@id=\"module_product_review\"]/div/div[3]/div[1]/div[" + count + "]/div[1]/span")).Text.ToString();
                        //month = driver.FindElements(By.CssSelector(".title.right"))[count].Text.ToString();
                        System.Threading.Thread.Sleep(1000);
                        string m = month.Substring(3, 3);
                        
                        count = 1;
                        jan = 0; feb = 0; mar = 0; apr = 0; may = 0; jun = 0; jul = 0; aug = 0; sep = 0; oct = 0; nov = 0; dec = 0;
                        if (m.Equals("Jan"))
                        {
                            jan++;
                        }
                        else if (m.Equals("Feb"))
                        {
                            feb++;
                        }
                        else if (m.Equals("Mar"))
                        {
                            mar++;
                        }
                        else if (m.Equals("Apr"))
                        {
                            apr++;
                        }
                        else if (m.Equals("May"))
                        {
                            may++;
                        }
                        else if (m.Equals("Jun"))
                        {
                            jun++;
                        }
                        else if (m.Equals("Jul"))
                        {
                            jul++;
                        }
                        else if (m.Equals("Aug"))
                        {
                            aug++;
                        }
                        else if (m.Equals("Sep"))
                        {
                            sep++;
                        }
                        else if (m.Equals("Oct"))
                        {
                            oct++;
                        }
                        else if (m.Equals("Nov"))
                        {
                            nov++;
                        }
                        else if (m.Equals("Dec"))
                        {
                            dec++;
                        }
                        else { }
                    }
                    else
                    {

                        for (int j = 0; j < a; j++)
                        {
                            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                            IWebElement SearchResult = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.XPath("//*[@id=\"module_product_review\"]/div/div[3]/div[1]/div[" + count + "]/div[1]/span")));
                            month = driver.FindElement(By.XPath("//*[@id=\"module_product_review\"]/div/div[3]/div[1]/div[" + count + "]/div[1]/span")).Text.ToString();
                            //month = driver.FindElements(By.CssSelector(".title.right"))[count].Text.ToString();
                            //foreach(char k in month)
                            //{
                            //    Console.WriteLine(month);
                            //}
                            string m = month.Substring(3, 3);
                            if (j % 5 == 0)
                            {

                                var button = driver.FindElement(By.ClassName("next-icon-last"));
                                System.Threading.Thread.Sleep(1000);
                                button.Click();
                                count = 1;
                            }

                            else
                            {
                                count++;
                            }
                            if (m.Equals("Jan"))
                            {
                                jan++;
                            }
                            else if (m.Equals("Feb"))
                            {
                                feb++;
                            }
                            else if (m.Equals("Mar"))
                            {
                                mar++;
                            }
                            else if (m.Equals("Apr"))
                            {
                                apr++;
                            }
                            else if (m.Equals("May"))
                            {
                                may++;
                            }
                            else if (m.Equals("Jun"))
                            {
                                jun++;
                            }
                            else if (m.Equals("Jul"))
                            {
                                jul++;
                            }
                            else if (m.Equals("Aug"))
                            {
                                aug++;
                            }
                            else if (m.Equals("Sep"))
                            {
                                sep++;
                            }
                            else if (m.Equals("Oct"))
                            {
                                oct++;
                            }
                            else if (m.Equals("Nov"))
                            {
                                nov++;
                            }
                            else if (m.Equals("Dec"))
                            {
                                dec++;
                            }
                            else { }

                        }

                    }
                    subcat = driver.FindElement(By.ClassName("breadcrumb")).Text.ToString();
                    Console.WriteLine("Category: " + subcat);
                    System.Threading.Thread.Sleep(4000);
                    int[] mo = { jan, feb, mar, apr, may, jun, jul, aug, sep, oct, nov, dec };
                    jan = 0; feb = 0; mar = 0; apr = 0; may = 0; jun = 0; jul = 0; aug = 0; sep = 0; oct = 0; nov = 0; dec = 0;
                    try
                    {

                        using (StreamWriter w = new StreamWriter("d.csv", true))
                        {
                            for (int k = 0; k < 12; k++)
                            {
                                var line = string.Format("{0},{1},{2},{3},{4},{5},{6}", name, pstr, sku, cat, subcat, mo[k], months[k]);
                                w.WriteLine(line);
                                w.Flush();
                            }
                        }
                    }
                    catch(Exception)
                    {
                    }
                    driver.Navigate().Back();

                }
                catch(StaleElementReferenceException e)
                {
                    Console.WriteLine(e);
                }
                catch(ArgumentOutOfRangeException e)
                {
                    Console.WriteLine(e);
                }
                catch(Exception e)
                {
                    Console.WriteLine(e);
                }
                }
        }
    }
}
