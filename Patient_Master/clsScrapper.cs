using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using OpenQA.Selenium.Remote;
using System.Net;

namespace Patient_Master
{
    public class clsScrapper
    {        
        IWebDriver driver;
        private IWebElement element;
        
        public clsMap ScrapData(string oPath, string oUrl, string patientName)
        {
            clsMap oMap = new clsMap();
            oPath = oPath.Replace("\\bin", "").Replace("\\Debug", "").Replace("\\Patient_Master.EXE", "");
            driver = new ChromeDriver(oPath);
            driver.Url = oUrl;
            element = driver.FindElement(By.Id("text"));
            string Abstract = element.Text;
            element = driver.FindElement(By.CssSelector("search-result.style-scope.search-app:nth-child(7) search-ui.style-scope.search-result.x-scope.search-ui-0 div.style-scope.search-ui div.layout.horizontal.style-scope.search-ui:nth-child(2) div.style-scope.search-ui:nth-child(2) div.style-scope.search-result div.vertical.layout.style-scope.search-result result-container.style-scope.search-result:nth-child(2) patent-result.style-scope.result-container:nth-child(2) div.layout.horizontal.style-scope.patent-result div.style-scope.patent-result:nth-child(1) div.style-scope.patent-result div.horizontal.layout.style-scope.patent-result:nth-child(2) div.flex-2.style-scope.patent-result section.knowledge-card.style-scope.patent-result dl.important-people.style-scope.patent-result:nth-child(2) > dd.style-scope.patent-result:nth-child(7)"));
            string Current_Assignee = element.Text;
            element = driver.FindElement(By.CssSelector("search-result.style-scope.search-app:nth-child(7) search-ui.style-scope.search-result.x-scope.search-ui-0 div.style-scope.search-ui div.layout.horizontal.style-scope.search-ui:nth-child(2) div.style-scope.search-ui:nth-child(2) div.style-scope.search-result div.vertical.layout.style-scope.search-result result-container.style-scope.search-result:nth-child(2) patent-result.style-scope.result-container:nth-child(2) div.layout.horizontal.style-scope.patent-result div.style-scope.patent-result:nth-child(1) div.style-scope.patent-result div.horizontal.layout.style-scope.patent-result:nth-child(2) div.flex-2.style-scope.patent-result section.knowledge-card.style-scope.patent-result application-timeline.style-scope.patent-result:nth-child(4) div.wrap.style-scope.application-timeline div.event.layout.horizontal.style-scope.application-timeline:nth-child(12) div.flex.title.style-scope.application-timeline > span.title-text.style-scope.application-timeline:nth-child(4)"));
            string status = element.Text;
            string Classfication = driver.FindElement(By.CssSelector("search-result.style-scope.search-app:nth-child(7) search-ui.style-scope.search-result.x-scope.search-ui-0 div.style-scope.search-ui div.layout.horizontal.style-scope.search-ui:nth-child(2) div.style-scope.search-ui:nth-child(2) div.style-scope.search-result div.vertical.layout.style-scope.search-result result-container.style-scope.search-result:nth-child(2) patent-result.style-scope.result-container:nth-child(2) div.layout.horizontal.style-scope.patent-result div.style-scope.patent-result:nth-child(1) div.style-scope.patent-result div.horizontal.layout.style-scope.patent-result:nth-child(2) div.flex-3.style-scope.patent-result section.style-scope.patent-result:nth-child(4) classification-viewer.style-scope.patent-result:nth-child(3) div.table.style-scope.classification-viewer classification-tree.style-scope.classification-viewer:nth-child(1) div.style-scope.classification-tree > div.layout.horizontal.wrap.style-scope.classification-tree")).Text;
            string Claim = driver.FindElement(By.CssSelector("search-result.style-scope.search-app:nth-child(7) search-ui.style-scope.search-result.x-scope.search-ui-0 div.style-scope.search-ui div.layout.horizontal.style-scope.search-ui:nth-child(2) div.style-scope.search-ui:nth-child(2) div.style-scope.search-result div.vertical.layout.style-scope.search-result result-container.style-scope.search-result:nth-child(2) patent-result.style-scope.result-container:nth-child(2) div.layout.horizontal.style-scope.patent-result div.style-scope.patent-result:nth-child(1) div.style-scope.patent-result div.horizontal.layout.style-scope.patent-result:nth-child(3) div.flex.flex-width.style-scope.patent-result:nth-child(2) section.style-scope.patent-result patent-text.style-scope.patent-result:nth-child(2) div.layout.horizontal.style-scope.patent-text:nth-child(3) section.flex.style-scope.patent-text > div.claims.style-scope.patent-text")).Text;
            List<IWebElement> elms = driver.FindElements(By.TagName("img")).ToList();
            string imgUrl = elms[1].GetAttribute("src");
            
            //SaveImage(imgUrl, patientName);
            driver.Close();
            return oMap;
        }
       void SaveImage(string url, string patientName)
        {
            using (WebClient client = new WebClient())
            {
                client.DownloadFile(new Uri(url), Path.GetTempPath() + patientName);
            }
        }
    }
}
