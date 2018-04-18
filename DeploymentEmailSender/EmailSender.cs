using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Net.Mail;
using System.Net.Mime;


namespace DeploymentEmailSender
{
    public static class EmailSender
    {
        static IWebDriver driver = null;
        static string filePath_uxp_table;
        static string filePath_uxp_graph;
        static string website_url = "####website_url####";

        public static void Main(string[] args)
        {
            filePath_uxp_table = "####file_path####";
            filePath_uxp_graph = "####file_path####";
            ScreenShot_UXP();
            SendEmail();
        }

        public static void ScreenShot_UXP()
        {
            driver = new ChromeDriver();

            driver.Navigate().GoToUrl(website_url);
            driver.Manage().Window.Maximize();
            System.Threading.Thread.Sleep(5000);

            //Login Jira
            DoLogin();
            ((IJavaScriptExecutor)driver).ExecuteScript("document.body.style.zoom='90%';");

            //Screenshot table 
            ClickMaximizeTableButton();
            TakeWholePageScreenShot(filePath_uxp_table);

            //Screenshot graph
            ClickMaximizeTableButton();
            ClickMaximizeGraphButton();
            TakeWholePageScreenShot(filePath_uxp_graph);

            driver.Close();
        }

        private static void HoverHeaderTitle(string element)
        {
            MouseHoverByLocator(driver, By.CssSelector(element));
            System.Threading.Thread.Sleep(2000);
        }

        private static void DoLogin()
        {
            driver.FindElement(By.CssSelector(".aui-nav-link.login-link")).Click();
            driver.FindElement(By.CssSelector("#login-form-username")).SendKeys("####username####");
            driver.FindElement(By.CssSelector("#login-form-password")).SendKeys("####password####");
            driver.FindElement(By.CssSelector("#login-form-submit")).Click();
            System.Threading.Thread.Sleep(5000);
            driver.FindElement(By.CssSelector(".aui-icon.icon-close")).Click();
        }

        private static void TakeWholePageScreenShot(string path)
        {
            var totalWidth = (int)(long)((IJavaScriptExecutor)driver).ExecuteScript("return Math.max(document.body.scrollWidth, document.body.offsetWidth, document.documentElement.clientWidth, document.documentElement.scrollWidth, document.documentElement.offsetWidth);");
            var totalHeight = (int)(long)((IJavaScriptExecutor)driver).ExecuteScript("return Math.max(document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight);");

            // Get the size of the viewport
            bool isJQuery = (bool)((IJavaScriptExecutor)driver).ExecuteScript("return window.jQuery != undefined");
            //int viewportWidth = 0;
            int viewportHeight = 0;
            if (isJQuery)
            {
                try
                {
                    viewportHeight = (int)(long)((IJavaScriptExecutor)driver).ExecuteScript("return $(window).height();");
                }
                catch
                {
                    viewportHeight = (int)(long)((IJavaScriptExecutor)driver).ExecuteScript("return window.innerHeight");
                }
            }
            else
            {
                viewportHeight = (int)(long)((IJavaScriptExecutor)driver).ExecuteScript("return window.innerHeight");
            }

            // Split the screen in multiple Rectangles
            var rectangles = new List<Rectangle>();
            // Loop until the totalHeight is reached
            for (var y = 0; y < totalHeight; y += viewportHeight)
            {
                var newHeight = viewportHeight;
                // Fix if the height of the element is too big
                if (y + viewportHeight > totalHeight)
                {
                    newHeight = totalHeight - y;
                }
                // Loop until the totalWidth is reached
                for (var x = 0; x < totalWidth; x += totalWidth)
                {
                    var newWidth = totalWidth;
                    // Fix if the Width of the Element is too big
                    if (x + totalWidth > totalWidth)
                    {
                        newWidth = totalWidth - x;
                    }
                    // Create and add the Rectangle
                    var currRect = new Rectangle(x, y, newWidth, newHeight);
                    rectangles.Add(currRect);
                }
            }
            // Build the Image
            var stitchedImage = new Bitmap(totalWidth, totalHeight);
            // Get all Screenshots and stitch them together
            var previous = Rectangle.Empty;
            foreach (var rectangle in rectangles)
            {
                // Calculate the scrolling (if needed)
                if (previous != Rectangle.Empty)
                {
                    var xDiff = rectangle.Right - previous.Right;
                    var yDiff = rectangle.Bottom - previous.Bottom;
                    // Scroll
                    ((IJavaScriptExecutor)driver).ExecuteScript(String.Format("window.scrollBy({0}, {1})", xDiff, yDiff));
                }
                // Take Screenshot

                var screenshot = ((ITakesScreenshot)driver).GetScreenshot();
                // Build an Image out of the Screenshot
                var screenshotImage = ScreenshotToImage(screenshot);
                // Calculate the source Rectangle
                var sourceRectangle = new Rectangle(totalWidth - rectangle.Width, viewportHeight - rectangle.Height, rectangle.Width, rectangle.Height);
                // Copy the Image
                using (var graphics = Graphics.FromImage(stitchedImage))
                {
                    graphics.DrawImage(screenshotImage, rectangle, sourceRectangle, GraphicsUnit.Pixel);
                }
                // Set the Previous Rectangle
                previous = rectangle;
            }

            stitchedImage.Save(path, System.Drawing.Imaging.ImageFormat.Jpeg);
        }

        private static Image ScreenshotToImage(Screenshot screenshot)
        {
            Image screenshotImage;
            using (var memStream = new MemoryStream(screenshot.AsByteArray))
            {
                screenshotImage = Image.FromStream(memStream);
            }
            return screenshotImage;
        }

        private static void ClickMaximizeTableButton()
        {
            HoverHeaderTitle("#gadget-17658-title");
            if (driver.FindElement(By.CssSelector("#gadget-17658-maximize")).Displayed)
            {
            driver.ClickElementByJavascript(driver.FindElement(By.CssSelector("#gadget-17658-maximize")));
            }
            else { return; }
        }

        private static void ClickMaximizeGraphButton()
        {
            HoverHeaderTitle("#gadget-17661-title");
            if (driver.FindElement(By.CssSelector("#gadget-17661-maximize")).Displayed)
            {

                driver.ClickElementByJavascript(driver.FindElement(By.CssSelector("#gadget-17661-maximize")));
            }
            else { return; }
        }

        private static void MouseHoverByLocator(this IWebDriver driver, By locator)
        {
            var builder = new Actions(driver);
            builder.MoveToElement(driver.FindElement(locator)).Build().Perform();
        }

        private static void ClickElementByJavascript(this IWebDriver driver, IWebElement element)
        {
            (driver as IJavaScriptExecutor).ExecuteScript("arguments[0].click();", element);
        }

        // send email 
        public static void SendEmail()
        {
            LinkedResource inline_uxp_table = new LinkedResource(filePath_uxp_table, MediaTypeNames.Image.Jpeg);
            inline_uxp_table.ContentId = Guid.NewGuid().ToString();

            LinkedResource inline_uxp_graph = new LinkedResource(filePath_uxp_graph, MediaTypeNames.Image.Jpeg);
            inline_uxp_graph.ContentId = Guid.NewGuid().ToString();

            string htmlBody = string.Format("<html><body><h1>UX Research Tracking</h1><br><img src=\"cid:{0}\"><img src=\"cid:{1}\"></body></html>", inline_uxp_table.ContentId, inline_uxp_graph.ContentId);
            AlternateView altView = AlternateView.CreateAlternateViewFromString(htmlBody, null, System.Net.Mime.MediaTypeNames.Text.Html);
            altView.LinkedResources.Add(inline_uxp_table);
            altView.LinkedResources.Add(inline_uxp_graph);

            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient("####mail_server####");
            mail.From = new MailAddress("####sender_email####");
            mail.To.Add("####reciever_email####");        
            mail.Subject = "UX Research Tracking: Date " + DateTime.Now.ToString("M/d/yyyy"); ;
            mail.AlternateViews.Add(altView);

            Attachment uxp_table = new Attachment(filePath_uxp_table);
            uxp_table.ContentDisposition.Inline = true;
            mail.Attachments.Add(uxp_table);
            Attachment uxp_graph = new Attachment(filePath_uxp_graph);
            uxp_graph.ContentDisposition.Inline = true;
            mail.Attachments.Add(uxp_graph);

            SmtpServer.Port = 25;
            SmtpServer.DeliveryMethod = SmtpDeliveryMethod.Network;
            SmtpServer.UseDefaultCredentials = false;

            try
            {
                SmtpServer.Send(mail);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }

}
