using LocalWincontrolSrv.OfficeControl.Powerpoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http;

namespace LocalWincontrolSrv.Web.Controllers
{
    public class PowerpointController : ApiController
    {
        // GET api/powerpoint 
        public string Get()
        {
            return "Commands /nextSlide /prevSlide /firstSlide /lastSlide /getCurrentSlide /status";
        }

        // GET api/powerpoint/action
        public string Get(string id)
        {
            var returnValue = "";

            switch (id.ToLower())
            {
                case "nextslide":
                    returnValue = PowerpointCtrl.nextSlide();
                    break;
                case "prevslide":
                    returnValue = PowerpointCtrl.prevSlide();
                    break;
                case "firstslide":
                    returnValue = PowerpointCtrl.firstSlide();
                    break;
                case "lastslide":
                    returnValue = PowerpointCtrl.lastSlide();
                    break;
                case "getcurrentslide":
                    returnValue = PowerpointCtrl.currentSlide();
                    break;
                case "status":
                    returnValue = PowerpointCtrl.checkStatus();
                    break;
            }
                
                    
            return returnValue;
        }

        // POST api/powerpoint 
        public void Post([FromBody]string value)
        {
        }

        // PUT api/powerpoint/5 
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE api/powerpoint/5 
        public void Delete(int id)
        {
        }
    }
}
