using LocalWincontrolSrv.Web;
using Microsoft.Owin.Hosting;
using System;
using System.Net;
using System.Net.Http;
using System.Threading;

namespace LocalWincontrolSrv
{
    class Program
    {
        static void Main(string[] args)
        {
            string hostName = Dns.GetHostName();
            string myIP = Dns.GetHostByName(hostName).AddressList[0].ToString();
            string baseAddress = $"http://{myIP}:8080/";

            // Start OWIN host 
            try
            {
                WebApp.Start<Startup>(url: baseAddress);
                Console.WriteLine($"Server started.. Can be found on url: {baseAddress}");
                
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }


            //testApp(baseAddress);
            Console.ReadLine();
        }

        public static void testApp(string baseAddress)
        {
            // Create HttpCient and make a request to api/values 
            HttpClient client = new HttpClient();

            //Check API
            var response = client.GetAsync(baseAddress + "api/powerpoint").Result;

            Console.WriteLine(response);
            Console.WriteLine("");
            Console.WriteLine("Server responding.. Testing api");

            Console.WriteLine("powerpoint root");
            Thread.Sleep(3000);
            Console.WriteLine(response.Content.ReadAsStringAsync().Result);
            Console.WriteLine("");

            Console.WriteLine("");
            Console.WriteLine("Api responding.. Testing powerpoint");

            //Check status
            var status = client.GetAsync(baseAddress + "api/powerpoint/status").Result;

            Console.WriteLine("status");
            Thread.Sleep(3000);
            Console.WriteLine(status.Content.ReadAsStringAsync().Result);
            Console.WriteLine("");
            

            //Go to first slide
            var firstslide = client.GetAsync(baseAddress + "api/powerpoint/firstslide").Result;

            Console.WriteLine("going to first slide");
            Thread.Sleep(3000);
            Console.WriteLine(firstslide.Content.ReadAsStringAsync().Result);
            Console.WriteLine("");

            //Go to next slide
            var nextslide1 = client.GetAsync(baseAddress + "api/powerpoint/nextslide").Result;

            Console.WriteLine("goint to next slide");
            Thread.Sleep(3000);
            Console.WriteLine(nextslide1.Content.ReadAsStringAsync().Result);
            Console.WriteLine("");
            

            //Go to next slide
            var nextslide2 = client.GetAsync(baseAddress + "api/powerpoint/nextslide").Result;

            Console.WriteLine("goint to next slide");
            Thread.Sleep(3000);
            Console.WriteLine(nextslide2.Content.ReadAsStringAsync().Result);
            Console.WriteLine("");

            //Go to prev slide
            var prevslide = client.GetAsync(baseAddress + "api/powerpoint/prevslide").Result;

            Console.WriteLine("going to previous slide");
            Thread.Sleep(3000);
            Console.WriteLine(prevslide.Content.ReadAsStringAsync().Result);
            Console.WriteLine("");
            

            //Go to last slide
            var lastslide = client.GetAsync(baseAddress + "api/powerpoint/lastslide").Result;

            Console.WriteLine("going to last slide");
            Thread.Sleep(3000);
            Console.WriteLine(lastslide.Content.ReadAsStringAsync().Result);
            Console.WriteLine("");

            //Go to prev slide
            var prevslide1 = client.GetAsync(baseAddress + "api/powerpoint/prevslide").Result;

            Console.WriteLine("going back one slide");
            Thread.Sleep(3000);
            Console.WriteLine(prevslide1.Content.ReadAsStringAsync().Result);
            Console.WriteLine("");

            //Get current slide
            var currSlide = client.GetAsync(baseAddress + "api/powerpoint/getcurrentslide").Result;

            Console.WriteLine("Getting current slide");
            Thread.Sleep(3000);
            Console.WriteLine(currSlide.Content.ReadAsStringAsync().Result);
            Console.WriteLine("");

            
        }
    }
}
