using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Remote;
using System.Net;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Text.RegularExpressions;
using WDSE.Decorators;
using WDSE.ScreenshotMaker;
using OpenQA.Selenium.Support.Extensions;
using WDSE;
using System.Drawing.Imaging;
using System.Threading;
using System.Windows.Forms;
using System.Data;
using OpenQA.Selenium.Support.UI;

namespace Patient_Master
{
    public class clsScrapper
    {        
        IWebDriver driver;
        private IWebElement element;
        LogWriter logger = new LogWriter();        
        public clsScrapper(string oPath)
        {
            logger.LogWrite(":clsScrapper: initiating chrome browser");
            oPath = Application.StartupPath;
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--start-maximized");
            options.AddArgument("headless");
            driver = new ChromeDriver(oPath,options);
            
            driver.Manage().Window.Size = new Size(429, 489);
            logger.LogWrite(":clsScrapper: chrome browser initiated successfully");
        }
        public clsMap ScrapData( string oUrl, string patientName, bool isWordOutput=false)
        {
            clsMap oMap = new clsMap();
            string FailUrl = "Unable to load Url";
            logger.LogWrite(":ScrapData: Begin");
            try
            {
                logger.LogWrite(":ScrapData: Navigating " + oUrl + " begin");
                driver.Url = oUrl;
                if(WaitUntilElementExists(By.CssSelector("div[class='abstract style-scope patent-text']")) == false)
                {
                    logger.LogWrite(":ScrapData: Unable to load url: " + oUrl);
                    oMap.Abstract = FailUrl;                    
                    return oMap;
                }
                logger.LogWrite(":ScrapData: Url loaded successfully");

                //logger.LogWrite(":ScrapData: Getting Title");
                //if (IsElementPresent(By.Id("title")))
                //    oMap.title = driver.FindElement(By.Id("title")).Text;
                //else
                //    oMap.title = string.Empty;

                logger.LogWrite(":ScrapData: Getting Abstract");
                if (IsElementPresent(By.CssSelector("div[class='abstract style-scope patent-text']")))
                    oMap.Abstract = driver.FindElement(By.CssSelector("div[class='abstract style-scope patent-text']")).Text;
                else
                    oMap.Abstract = string.Empty;

                logger.LogWrite(":ScrapData: Getting Current Assignee");
                if (IsElementPresent(By.TagName("dl")))
                {
                    try
                    {
                        List<IWebElement> els = driver.FindElements(By.TagName("dl"))[0].FindElements(By.TagName("dd")).ToList();
                    
                        oMap.Current_Assignee = els[els.Count - 1].Text;
                    }
                    catch (Exception ex)
                    {
                        logger.LogWrite(":ScrapData: Gettng Current Assignee >> Error: " + ex.Message);
                    }

                }
                else
                    oMap.Current_Assignee = string.Empty;

                logger.LogWrite(":ScrapData: Getting Status");
                if (IsElementPresent(By.CssSelector("div[class='flex title style-scope application-timeline']")))
                {
                    try
                    {
                        List<IWebElement> elmts = driver.FindElements(By.CssSelector("div[class='flex title style-scope application-timeline']")).ToList();
                    
                        oMap.status = elmts[elmts.Count - 2].Text;
                    }
                    catch (Exception ex)
                    {
                        logger.LogWrite(":ScrapData: Getting Status >> Error: " + ex.Message);
                    }

                }
                else
                    oMap.status = string.Empty;

                logger.LogWrite(":ScrapData: Getting Classification");
                if (IsElementPresent(By.CssSelector("div[class='layout horizontal wrap style-scope classification-tree']")))
                    oMap.Classfication = driver.FindElement(By.CssSelector("div[class='layout horizontal wrap style-scope classification-tree']")).Text;
                else
                    oMap.Classfication = string.Empty;

                logger.LogWrite(":ScrapData: Getting Claims");
                if (IsElementPresent(By.Id("claims")))
                {
                    string claims= driver.FindElements(By.Id("claims"))[1].Text;
                    try
                    {
                        claims = claims.Substring(claims.IndexOf("\n1.") + 4);
                    }
                    catch (Exception ex)
                    {
                        logger.LogWrite(":ScrapData: Getting Claims >> Error: " + ex.Message);
                    }
                    
                    oMap.Claim = claims;
                }
                else
                    oMap.Claim = string.Empty;

                logger.LogWrite(":ScrapData: Date Of Anticipated Expiration");
                if (IsElementPresent(By.CssSelector("div[class='event layout horizontal style-scope application-timeline']")))
                {
                    oMap.anticipationExpiry = driver.FindElement(By.CssSelector("div[class='event layout horizontal style-scope application-timeline']")).Text.Split('\n')[0].Trim();
                }

                logger.LogWrite(":ScrapData: Getting Description");
                if(IsElementPresent(By.CssSelector("section[class='flex style-scope patent-text']")))
                {                    
                    oMap.Description= driver.FindElement(By.CssSelector("section[class='flex style-scope patent-text']")).Text;
                }
                
                try
                {
                    if (isWordOutput)
                    {
                        logger.LogWrite(":ScrapData: Capturing Image");
                        element = driver.FindElement(By.CssSelector("img[class='style-scope image-carousel']"));
                        Actions act = new Actions(driver);
                        act.MoveToElement(element).Click().Build().Perform();
                        Thread.Sleep(5000);

                        ((ITakesScreenshot)driver).GetScreenshot().SaveAsFile(Path.Combine(Path.GetTempPath(), patientName), ScreenshotImageFormat.Png);
                    }
                }
                catch (Exception ex)
                {
                    logger.LogWrite(":ScrapData: Capturing Image >> Error: " + ex.Message);
                }                             
                return oMap;
            }
            catch (Exception ex)
            {
                //MessageBox.Show("clsScrapper: " + ex.Message);
                logger.LogWrite(":ScrapData: Error: " + ex.Message);
                //driver.Close();
                //driver.Quit();
                oMap.Abstract = FailUrl;
                return oMap;
            }
        }
        //this will search for the element until a timeout is reached
        public bool WaitUntilElementExists(By elementLocator, int timeout = 20)
        {
            try
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeout));
                IWebElement  elm=wait.Until(ExpectedConditions.ElementExists(elementLocator));
                if (elm != null)
                    return true;
                else
                    return false;
            }
            catch (NoSuchElementException)
            {
                Console.WriteLine("Element with locator: '" + elementLocator + "' was not found in current context page.");
                return false;
            }
        }
        public List<patentCitation> GetPatent_Citations()
        {
            try
            {
                List<patentCitation> dtPatentCitation = new List<patentCitation>();
                if (IsElementPresent(By.CssSelector("div[class='tbody style-scope patent-result']")))
                {
                    if (IsElementPresent(By.CssSelector("div[class='tr style-scope patent-result']")))
                    {
                        List<IWebElement> patentRows = driver.FindElements(By.CssSelector("div[class='tr style-scope patent-result']")).ToList();
                        foreach (IWebElement item in patentRows)
                        {
                            patentCitation pCitation = new patentCitation();
                            if (IsElementPresent(By.CssSelector("span[class='td nowrap style-scope patent-result']"), item))
                            {
                                try
                                {
                                    pCitation.publication_number = item.FindElement(By.CssSelector("span[class='td nowrap style-scope patent-result']")).Text;
                                }
                                catch (Exception)
                                {
                                }
                                try
                                {
                                    pCitation.priority_date = item.FindElements(By.CssSelector("span[class='td style-scope patent-result']"))[0].Text;
                                }
                                catch (Exception)
                                {
                                }
                                try
                                {
                                    pCitation.publication_date = item.FindElements(By.CssSelector("span[class='td style-scope patent-result']"))[1].Text;
                                }
                                catch (Exception)
                                {
                                }
                                try
                                {
                                    pCitation.assignee = item.FindElements(By.CssSelector("span[class='td style-scope patent-result']"))[2].Text;
                                }
                                catch (Exception)
                                {
                                }
                                try
                                {
                                    pCitation.title = item.FindElements(By.CssSelector("span[class='td style-scope patent-result']"))[3].Text;
                                }
                                catch (Exception)
                                {
                                }
                                dtPatentCitation.Add(pCitation);
                            }
                        }
                    }
                }
                return dtPatentCitation;
            }
            catch (Exception)
            {
                return null;
            }
        }
        void CropImage(string imagepath)
        {
            Bitmap sourceImg = (Bitmap)Bitmap.FromFile(Path.Combine(Path.GetTempPath(),imagepath));
            using(var absentRectangleImage = sourceImg)
            {
                using(var currentTile=new Bitmap(256, 256))
                {
                    currentTile.SetResolution(absentRectangleImage.HorizontalResolution, absentRectangleImage.VerticalResolution);
                    using(var currentTileGraphics = Graphics.FromImage(currentTile))
                    {
                        currentTileGraphics.Clear(Color.Black);
                        var absentRectangleArea = new Rectangle(3, 8963, 256, 256);
                        currentTileGraphics.DrawImage(absentRectangleImage, 0, 0, absentRectangleArea, GraphicsUnit.Pixel);
                    }
                    currentTile.Save(Path.Combine(Path.GetTempPath(), "__" + imagepath));
                }
            }
        }
        public void CloseSession()
        {
            logger.LogWrite(":CloseSession: Begin");
            try
            {
                driver.Close();
                driver.Quit();
            }
            catch (Exception ex)
            {
                logger.LogWrite(":CloseSession: Error: " + ex.Message);
            }
            logger.LogWrite(":CloseSession: Ends");
        }
        private bool IsElementPresent(By by)
        {
            try
            {
                driver.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
        private bool IsElementPresent(By by, IWebElement elms)
        {
            try
            {
                elms.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
        public void SaveImage(string url, string patientName)
        {
            using (WebClient client = new WebClient())
            {
                client.DownloadFile(new Uri(url), Path.GetTempPath() + patientName);
            }
            //try
            //{

            //    //resize
            //    Image img = Image.FromFile(Path.GetTempPath() + patientName);
            //    Bitmap b = new Bitmap(img);
            //    Image resizediMg = resizeImage(b, new Size(500, 500));
            //    //File.Delete(Path.GetTempPath() + patientName);
            //    resizediMg.Save(Path.GetTempPath() + "_" + patientName);
            //}
            //catch (Exception ex)
            //{


            //}

        }
        private System.Drawing.Image resizeImage(System.Drawing.Image imgToResize, Size size)
        {
            //Get the image current width  
            int sourceWidth = imgToResize.Width;
            //Get the image current height  
            int sourceHeight = imgToResize.Height;
            float nPercent = 0;
            float nPercentW = 0;
            float nPercentH = 0;
            //Calulate  width with new desired size  
            nPercentW = ((float)size.Width / (float)sourceWidth);
            //Calculate height with new desired size  
            nPercentH = ((float)size.Height / (float)sourceHeight);
            if (nPercentH < nPercentW)
                nPercent = nPercentH;
            else
                nPercent = nPercentW;
            //New Width  
            int destWidth = (int)(sourceWidth * nPercent);
            //New Height  
            int destHeight = (int)(sourceHeight * nPercent);
            Bitmap b = new Bitmap(destWidth, destHeight);
            Graphics g = Graphics.FromImage((System.Drawing.Image)b);
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            // Draw image with new width and height  
            g.DrawImage(imgToResize, 0, 0, destWidth, destHeight);
            g.Dispose();
            b.SetResolution(258, 342);
            return (System.Drawing.Image)b;
        }
    }
}
