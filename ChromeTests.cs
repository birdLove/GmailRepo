using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using System.Windows;
using Xunit;
using System.Runtime.InteropServices;
using OfficeOpenXml;
using Microsoft.WindowsAzure.Storage;
using System.IO.Compression;
using Sikuli4Net.sikuli_REST;
using Sikuli4Net.sikuli_UTIL;
using System.Drawing;
using OpenQA.Selenium.Support.UI;
using System.Net;
using RestSharp;
using Newtonsoft.Json.Linq;
using Nancy.Json;

namespace Selenium.NetCore.Test
{
    public class ChromeTests : IDisposable
    {
        private static IWebDriver driver;
        APILauncher launcher = new APILauncher(true);
        public object MessageBox { get; private set; }
		public static string homePath = Environment.GetEnvironmentVariable("HOMEPATH").ToString();


        public ChromeTests()
        {
            var directory = Directory.GetCurrentDirectory();
            var pathDrivers = directory + "/../../../../drivers";
            //Create a instance of ChromeOptions class
            ChromeOptions options = new ChromeOptions();

            //Add chrome switch to disable notification - "**--disable-notifications**"
            options.AddArgument("--disable-notifications");        

            driver = new ChromeDriver(pathDrivers, options);
            launcher.Start();
        }


        [Fact]
        public void SikulixAPI()
        {
            Pattern pattern = new Pattern(@"C:\" + homePath + @"t\Desktop\SeleniumProject1\mypng.png");
            Screen screen = new Screen();
            screen.Find(pattern);
            screen.Click(pattern);
            screen.Type(pattern, "Hi");
           // driver.Manage().Window.Position(-2000, 0);


            string zipPath = @"C:\"+homePath+@"\Desktop\SeleniumProject1\Docs.zip";
            string extractPath = @".\extract";
            string pathString2 = @"C:\"+homePath+@"\Desktop\SeleniumProject1\Docs";

            ZipFile.CreateFromDirectory(pathString2, zipPath);
            ZipFile.ExtractToDirectory(zipPath, pathString2);
        }

        [Fact]
        public void Program()
        {
            /*
            ProgramPage x = new ProgramPage();
                var query = from c in set where c.Name == name select c;
                Used obj = new Used();
                obj.Id = 123;
                obj.Name = "?";
            x.ItemList.Add(obj);
            x.SaveChanges();
            */
            

        }

        [Fact]
        public void ReadEmailsByKeyword()
        {
			var homePath = Environment.GetEnvironmentVariable("HOMEPATH").ToString();

            GmailLogin();
            WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, 10));

            //gb_ff
            wait.Until(d => d.FindElements(By.XPath("//*[@aria-label='Search mail']")).FirstOrDefault().Displayed);

            IWebElement searchBar = driver.FindElement(By.XPath("//*[@aria-label='Search mail']"));
            searchBar.Clear();
            searchBar.SendKeys("from: x@gmail.com");

            var searchBarClick = driver.FindElement(By.XPath("//button[@aria-label='Search Mail']"));
            wait.Until(d => searchBarClick.Enabled);
            searchBarClick.Click();
            Thread.Sleep(10000);

            string pathString2 = @"C:\" + homePath + @"\Desktop\SeleniumProject1\Docs";
            if (!System.IO.Directory.Exists(pathString2))
            {
                System.IO.Directory.CreateDirectory(pathString2);
            }
            string pathString = @"C:\" + homePath + @"\Desktop\SeleniumProject1\Docs\ReadAllEmailsByKeyword.txt";
            System.IO.File.Delete(pathString);
            Thread.Sleep(6000);

            File.AppendAllText(pathString, "Today:" + DateTime.Now);
            List<IWebElement> inboxEmails = driver.FindElements(By.XPath("//*[starts-with(@class,'zA')]")).ToList<IWebElement>();
            foreach (IWebElement email in inboxEmails)
            {
                if (email.Displayed)
                {
                    email.Click();
                    Thread.Sleep(2000);
                    try
                    {

                        //adn ads

                        IWebElement label = driver.FindElement(By.XPath("//*[starts-with(@class,'adn ads')]"));
                        Thread.Sleep(5000);
                        Console.WriteLine("Label: " + label.Text);
                        File.AppendAllText(pathString, label.Text);
                        Thread.Sleep(5000);

                        driver.Navigate().Back();
                        Thread.Sleep(3000);
                    }
                    catch(Exception e)
                    {
                        driver.Navigate().Back();
                        Thread.Sleep(3000);
                    }
                }

            }
        }

        [Fact]
        public void ReadEmails()
        {
            GmailLogin();
            WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, 10));

            List<IWebElement> inboxEmails = driver.FindElements(By.XPath("//*[starts-with(@class,'zA')]")).ToList<IWebElement>();
            File.WriteAllText(@"C:\" + homePath + @"\Desktop\SeleniumProject1\Docs\ReadAllEmails.txt", "Today:" + DateTime.Now);
            int index = 0;
            int totalMails = 0;
            foreach (IWebElement email in inboxEmails) {
                totalMails++;
                if (email.Displayed && email.Text.Contains("o l"))
                {
                    email.Click();
                    Thread.Sleep(2000);
                    //adn ads
                    IWebElement label = driver.FindElement(By.XPath("//*[starts-with(@class,'adn ads')]"));
                    Thread.Sleep(2000);
                    Console.WriteLine("Label: " + label.Text);
                    File.AppendAllText(@"C:\"+homePath+@"\Desktop\SeleniumProject1\Docs\ReadAllEmails.txt", label.Text);
                    Thread.Sleep(2000);
                    driver.Navigate().Back();
                    index++;
                }
                if(index > 10)
                {
                    break;
                }
                //I'm not sure the exact number
                if(totalMails > 50)
                {
                    driver.FindElement(By.Id(":l2")).Click();
                    Thread.Sleep(2000);
                }
            }
        }

        public void GmailLogin(String email = "x@gmail.com", String password = "xyz")
        {
            WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, 20));
            driver.Navigate().GoToUrl("https://gmail.com/");
            driver.Manage().Window.Maximize();

            IWebElement loginBox = driver.FindElement(By.Id("identifierId"));
            wait.Until(d => loginBox.Displayed);
            loginBox.SendKeys(email);

            IWebElement nextBtn = driver.FindElement(By.Id("identifierNext"));
            wait.Until(d => nextBtn.Enabled);
            nextBtn.Click();

            wait.Until(d => d.FindElement(By.Id("forgotPassword")).Text.Contains("Forgot password?"));
            IWebElement pwBox = driver.FindElement(By.Name("password"));
            wait.Until(d => pwBox.Displayed);
            pwBox.SendKeys(password);

            IWebElement signinBtn = driver.FindElement(By.Id("passwordNext"));
            wait.Until(d => signinBtn.Enabled);
            signinBtn.Click();

            wait.Until(d => d.FindElements(By.TagName("tr")).Count > 5);
            wait.Until(d => d.FindElements(By.XPath("//*[starts-with(@class,'zA')]")).FirstOrDefault().Displayed);


        }

        [Fact]
        public void DeleteEmailsByKeyword()
        {
            GmailLogin();
            WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, 10));

            List<String> itemArray = new List<String>();
            //itemArray.Add("label:unread");
            //itemArray.Add("in:social");
            itemArray.Add("from: y");

            wait.Until(d => d.FindElements(By.XPath("//*[starts-with(@class,'zA')]")).FirstOrDefault().Displayed);

            foreach (String item in itemArray)
            {
               
                    wait.Until(d => d.FindElements(By.XPath("//*[@aria-label='Search mail']")).FirstOrDefault().Displayed);

                    IWebElement searchBar = driver.FindElement(By.XPath("//*[@aria-label='Search mail']"));
                    searchBar.Clear();
                    searchBar.SendKeys(item);

                    IWebElement searchBarClick = driver.FindElement(By.XPath("//button[@aria-label='Search Mail']"));
                    wait.Until(d => searchBarClick.Enabled);
                    searchBarClick.Click();
                    Thread.Sleep(10000);

                try
                {
                        do
                        {
                            IWebElement checkBox = driver.FindElement(By.XPath("/html[1]/body[1]/div[7]/div[3]/div[1]/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/span[1]"));
                            checkBox.Click();
                            Thread.Sleep(4000);
                            IWebElement deleteThem = driver.FindElement(By.XPath("/html[1]/body[1]/div[7]/div[3]/div[1]/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]/div[3]"));
                            wait.Until(d => deleteThem.Enabled);
                            deleteThem.Click();
                            searchBar.Clear();

                            Thread.Sleep(10000);
                        } while (driver.FindElements(By.XPath("//*[starts-with(@class,'zA')]")).Count > 0);
                }
                catch (Exception e)
                {
                    
                }
            }

        }

        [Fact]
        public void ReadXcelFile()
        {
            File.Move(@"C:\"+homePath+@"\Desktop\SeleniumProject1\Book1.xlsx", @"C:\"+homePath+@"\Desktop\SeleniumProject1\Book1.csv");
            using (var reader = new StreamReader(@"C:\"+homePath+@"\Desktop\SeleniumProject1\Book1.csv"))
            {
                List<string> listA = new List<string>();
                List<string> listB = new List<string>();
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(';');
                    Console.WriteLine(line);

                    listA.Add(values[0]);
                    listB.Add(values[1]);
                }
                Console.WriteLine(listA[0]);
            }
        }

        public void MessengerLogin(String email = "x@gmail.com", String password = "y")
        {
            WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, 22));

            driver.Navigate().GoToUrl("https://www.messenger.com/");
            driver.Manage().Window.Maximize();

            IWebElement loginBox = driver.FindElement(By.Id("email"));
            wait.Until(d => loginBox.Displayed);
            loginBox.SendKeys(email);

            IWebElement pwBox = driver.FindElement(By.Id("pass"));
            wait.Until(d => pwBox.Displayed);
            pwBox.SendKeys(password);

            IWebElement nextBtn = driver.FindElement(By.Id("loginbutton"));
            wait.Until(d => nextBtn.Enabled);
            nextBtn.Click();
            Thread.Sleep(5000);
        }

        [Fact]
        public void ReadMyFile()
        {
            FileInfo existingFile = new FileInfo(@"C:\"+homePath+@"\Desktop\SeleniumProject1\Book1.xlsx");
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count
                string y = "";
                for (int row = 1; row <= rowCount; row++)
                {
                    for (int col = 1; col <= colCount; col++)
                    {
                        //You can update code here to add each cell value to DataTable.
                        string x = worksheet.Cells[row, col].GetValue<String>();

                        if(x != null)
                        {
                            y = y + x;
                            Console.WriteLine(x);
                        }
                    }
                }
                File.AppendAllText(@"C:\"+homePath+@"\Desktop\SeleniumProject1\Docs\excelFile.txt", y);
            }
        }


        [Fact]
        public void MessengerBot()
        {
            WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, 22));

            MessengerLogin();

            List<String> nameArray = new List<String>();
            // nameArray.Add("NewName");

            foreach (String name in nameArray)
            {
                try
                {
                    wait.Until(d => driver.FindElement(By.XPath("//input[@aria-label='Search Messenger']")).Displayed);
                    IWebElement nameBox = driver.FindElement(By.XPath("//input[@aria-label='Search Messenger']"));
                    nameBox.Clear();
                    nameBox.SendKeys(name);

                    /*
                    wait.Until(d => driver.FindElements(By.TagName("li")).ToList<IWebElement>().FirstOrDefault().Displayed);
                    List<IWebElement> myList = driver.FindElements(By.TagName("li")).ToList<IWebElement>();
                    myList.FirstOrDefault().Click();
                    */
                    Thread.Sleep(2000);
                    String search = "//div//div//div//div//div//div[2]//ul[1]//li[1]//a[1]//div[1]//div[2]//div[1]//div[1]";
                    wait.Until(d => driver.FindElement(By.XPath(search)).Displayed);
                    IWebElement myClick = driver.FindElement(By.XPath(search));
                    myClick.Click();


                    wait.Until(d => driver.FindElement(By.XPath("//*[starts-with(@class,'notranslate _5rpu')]")).Displayed);
                    String dir = Directory.GetCurrentDirectory();
                    Console.WriteLine(dir);
                    String[] buffer = System.IO.File.ReadAllLines(@"C:\"+homePath+@"\Desktop\SeleniumProject1\Docs\Messengerbot.txt");
                    String myBuffer = "";
                    //myBuffer = myBuffer + "Hi " + name + ",\n";

                    for (int index = 0; index < buffer.Length; index++)
                    {
                        myBuffer = myBuffer + buffer[index] + "\n";
                    }
                    IWebElement message = driver.FindElement(By.XPath("//*[starts-with(@class,'notranslate _5rpu')]"));
                    message.Clear();
                    message.SendKeys(myBuffer + Keys.Enter);
                    Thread.Sleep(4000);
                }
                catch(Exception e)
                {

                }
            }
        }

        [Fact]
        public void TagBot()
        {
            driver.Navigate().GoToUrl("https://facebook.com/");
            driver.Manage().Window.Maximize();
            Thread.Sleep(3000);

            IWebElement userName = driver.FindElement(By.Id("email"));
            userName.SendKeys("x@gmail.edu");
            Thread.Sleep(3000);

            IWebElement passwordElement = driver.FindElement(By.Id("pass"));
            passwordElement.SendKeys("y");
            Thread.Sleep(3000);

            IWebElement buttonLogin = driver.FindElement(By.Id("loginbutton"));
            buttonLogin.Click();
            Thread.Sleep(5000);

            List<String> names = new List<String>();
            names.Add("xyz");


            foreach (var name in names)
            {
                IWebElement nameSearch = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[1]/div[1]/div[1]/input[2]"));
                nameSearch.Clear();
                nameSearch.SendKeys(name);
                Thread.Sleep(3000);

                IWebElement click1 = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/button[1]"));
                click1.Click();
                Thread.Sleep(3000);

                IWebElement click2 = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[3]/div[1]/div[1]/div[3]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/a[1]"));
                click2.Click();
                Thread.Sleep(3000);

                //Actually post
                IWebElement postingBox = driver.FindElement(By.XPath("//div[@class='_5yk2']//div[@class='notranslate _5rpu']"));
                postingBox.Click();
                Thread.Sleep(4000);

                String[] buffer = System.IO.File.ReadAllLines(@"C:\"+homePath+@"\Desktop\SeleniumProject1\Docs\Tagbot.txt");
                String allbuff = "";
                allbuff = allbuff + "Today: " + DateTime.Now + "\n";
                allbuff = allbuff + "Hi " + name + ",\n";
                for (int index = 0; index < buffer.Length; index++)
                {
                    allbuff = allbuff + buffer[index] + "\n";
                }
                Thread.Sleep(4000);
                Random rand = new Random();
                int num = rand.Next();
                allbuff = allbuff + num;
                postingBox.SendKeys(allbuff);
                Thread.Sleep(7000);
                IWebElement submittingButton = driver.FindElement(By.XPath("//button[@class='_1mf7 _4jy0 _4jy3 _4jy1 _51sy selected _42ft']//span[contains(text(),'Post')]"));
                submittingButton.Click();
                Thread.Sleep(4000);
            }

        }

        [Fact]
        public void Emailbot()
        {
            GmailLogin();
            WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, 10));

            //Uncomment to send email to each person
            List<String> array = new List<String>();
            array.Add("a@gmail.com");

            List<String> nameArray = new List<String>();
            nameArray.Add("a");
            
            String[] buffer = System.IO.File.ReadAllLines(@"C:\"+homePath+@"\Desktop\SeleniumProject1\Docs\Emailbot.txt");
            String allbuff = "";
            String subText = "";
            for (int index = 0; index < buffer.Length; index++)
            {
                if (index == 0)
                {
                    subText = (buffer[index]);
                }
                allbuff = allbuff + buffer[index] + "\n";
            }

            for (int numEmails = 0; numEmails < array.Count; numEmails++)
                {
                    wait.Until(d => d.FindElement(By.XPath("//div[contains(text(),'Compose')]")).Displayed);
                    driver.FindElement(By.XPath("//div[contains(text(),'Compose')]")).Click();

                    wait.Until(d => d.FindElement(By.XPath("//textarea[@name='to']")).Displayed);

                    IWebElement subject = driver.FindElement(By.Name("subjectbox"));
                    IWebElement body = driver.FindElement(By.XPath("//div[@aria-label='Message Body']"));

                    String myBuffer ="Today: " + DateTime.Now + "\n";
                    myBuffer = myBuffer + "Dear " + nameArray[numEmails] + ",\n" + allbuff;
                    myBuffer = myBuffer + "\n -Haley";
                    IWebElement input = driver.FindElement(By.XPath("//textarea[@name='to']"));
                    input.Clear();
                    input.SendKeys(array[numEmails]);
                    Thread.Sleep(7000);
                    subject.Clear();
                    subject.SendKeys(subText);
                    body.Clear();
                    // Read the file and display it line by line.  
                    body.SendKeys(myBuffer);
                    Thread.Sleep(4000);

                //wG J-Z-I
                /* NOT WORKING ATTACH FILE
                driver.FindElement(By.XPath("//*[starts-with(@class,'wG J-Z-I')]")).Click();
                Thread.Sleep(4000);
                Actions action = new actions;
                action.SendKeys(@"C:\"+homePath+"\Desktop\SeleniumProject1\Invite.txt" + Keys.Enter).Build().Perform();
                Thread.Sleep(4000);
                */

                //starts with T-I J-J5-Ji aoO v7 T-I-atl L3
                IWebElement sendEmail = driver.FindElement(By.XPath("//*[starts-with(@class,'dC')]"));
                wait.Until(d=>sendEmail.Enabled);
                sendEmail.Click();
                Thread.Sleep(5000);
                }
            
        }

        [Fact]
        public void RestCall()
        {
            GetTokenChooseUser("x@gmail.com", "y", "http://gmail.com");
        }

        public String GetTokenChooseUser(String username, String password, String url)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
            var client = new RestClient("");
            var request = new RestRequest(url);

            request.Method = Method.POST;
            request.AddHeader("Accept", "application/json");
            request.Parameters.Clear();
            string body = new JavaScriptSerializer().Serialize(new
            {
                username,
                password
            });
            request.AddParameter("application/json", body, ParameterType.RequestBody);

            var response = client.Execute(request);
            var content = response.Content;

            JObject accessToken = JObject.Parse(response.Content);
            String myToken = accessToken["token"].ToString();
            return myToken;
        }

        public IRestResponse GetResultsEverything(string token, string url)
        {
            RestClient client2 = new RestClient(url);
            RestRequest getRequest = new RestRequest(Method.GET);
            getRequest.AddHeader("Content-Type", "application/json");
            getRequest.AddHeader("Authorization", "Bearer " + token);

            IRestResponse getResponse = client2.Execute(getRequest);
            return getResponse;
        }


        public IRestResponse PostResultsEverything(string token, string url)
        {
            RestClient client2 = new RestClient(url);
            RestRequest getRequest = new RestRequest(Method.POST);
            getRequest.AddHeader("Content-Type", "application/json");
            getRequest.AddHeader("Authorization", "Bearer " + token);

            IRestResponse getResponse = client2.Execute(getRequest);
            return getResponse;
        }

        public void Dispose()
        {
            driver?.Dispose();
            launcher.Stop();
        }
    }
}
