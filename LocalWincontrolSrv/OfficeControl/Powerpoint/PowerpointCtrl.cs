using System;
using PPt = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;

namespace LocalWincontrolSrv.OfficeControl.Powerpoint
{
    public static class PowerpointCtrl
    {
        public static string checkStatus()
        {
            var status = "";
            try
            {
                var pptApplication = Marshal.GetActiveObject("PowerPoint.Application") as PPt.Application;
                var nextslideStatus = nextSlide();
                status = prevSlide();
                if (nextslideStatus == "false")
                {
                    status = nextSlide();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                status = "false";
                
            }
            
            return status;
        }

        public static string nextSlide()
        {
            var status = "";
            try
            {
                var pptApplication = Marshal.GetActiveObject("PowerPoint.Application") as PPt.Application;

                // Get Presentation Object 
                var presentation = pptApplication.ActivePresentation;
                // Get Slide collection object 
                var slides = presentation.Slides;
                // Get Slide count 
                var slidescount = slides.Count;
                // Get current selected slide  
                try
                {
                    // Get selected slide object in normal view 
                    var slide = slides[pptApplication.ActiveWindow.Selection.SlideRange.SlideNumber];

                    var slideIndex = slide.SlideIndex + 1;
                    if (slideIndex > slidescount)
                    {
                        status = "false";
                    }
                    else
                    {
                        slide = slides[slideIndex];
                        slides[slideIndex].Select();

                        status = "true";
                    }

                }
                catch
                {
                    // Get selected slide object in reading view 
                    var slide = pptApplication.SlideShowWindows[1].View.Slide;

                    var slideIndex = slide.SlideIndex + 1;
                    if (slideIndex > slidescount)
                    {
                        status = "false";
                    }
                    else
                    {
                        pptApplication.SlideShowWindows[1].View.Next();
                        slide = pptApplication.SlideShowWindows[1].View.Slide;
                     
                        status = "true";
                    }
                }

                if (status == "false")
                {
                    Console.WriteLine("next slide failed to execute");
                }

            }
            catch (Exception e)
            {
                Console.WriteLine("Next slide failed:"+e.Message);
                status = "false";

            }

            return status;
        }

        public static string prevSlide()
        {
            var status = "";
            try
            {
                var pptApplication = Marshal.GetActiveObject("PowerPoint.Application") as PPt.Application;

                // Get Presentation Object 
                var presentation = pptApplication.ActivePresentation;
                // Get Slide collection object 
                var slides = presentation.Slides;
                // Get Slide count 
                var slidescount = slides.Count;
                // Get current selected slide  
                try
                {
                    // Get selected slide object in normal view 
                    var slide = slides[pptApplication.ActiveWindow.Selection.SlideRange.SlideNumber];

                    var slideIndex = slide.SlideIndex - 1;
                    if (slideIndex >= 1)
                    {
                        slide = slides[slideIndex];
                        slides[slideIndex].Select();
                        status = "true";
                    }
                    else
                    {
                         status = "false";
                    }

                }
                catch
                {
                    // Get selected slide object in reading view 
                    var slide = pptApplication.SlideShowWindows[1].View.Slide;

                    var slideIndex = slide.SlideIndex - 1;
                    if (slideIndex >= 1)
                    {
                        pptApplication.SlideShowWindows[1].View.Previous();
                        slide = pptApplication.SlideShowWindows[1].View.Slide;
                        status = "true";
                    }
                    else
                    {
                        status = "false";
                    }
                }


            }
            catch (Exception)
            {
                status = "false";

            }

            return status;
        }

        public static string firstSlide()
        {
            var status = "";
            try
            {
                var pptApplication = Marshal.GetActiveObject("PowerPoint.Application") as PPt.Application;               

                // Get Presentation Object 
                var presentation = pptApplication.ActivePresentation;
                // Get Slide collection object 
                var slides = presentation.Slides;
                // Get Slide count 
                var slidescount = slides.Count;
                // Get current selected slide  
                try
                {
                    // Get selected slide object in normal view 
                    var slide = slides[pptApplication.ActiveWindow.Selection.SlideRange.SlideNumber];
                        
                    // Call Select method to select first slide in normal view 
                    slides[1].Select();
                    slide = slides[1];
                    status = "true";
                }
                catch
                {
                    // Get selected slide object in reading view 
                    var slide = pptApplication.SlideShowWindows[1].View.Slide;

                    // Transform to first page in reading view 
                    pptApplication.SlideShowWindows[1].View.First();
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                    status = "true";
                }
            }
            catch (Exception)
            {
                status = "false";
            }

            return status;
        }

        public static string lastSlide()
        {
            var status = "";
            try
            {
                var pptApplication = Marshal.GetActiveObject("PowerPoint.Application") as PPt.Application;

                // Get Presentation Object 
                var presentation = pptApplication.ActivePresentation;
                // Get Slide collection object 
                var slides = presentation.Slides;
                // Get Slide count 
                var slidescount = slides.Count;
                // Get current selected slide  
                try
                {
                    // Get selected slide object in normal view 
                    var slide = slides[pptApplication.ActiveWindow.Selection.SlideRange.SlideNumber];

                    slides[slidescount].Select();
                    slide = slides[slidescount];

                    status = "true";

                }
                catch
                {
                    // Get selected slide object in reading view 
                    var slide = pptApplication.SlideShowWindows[1].View.Slide;

                    pptApplication.SlideShowWindows[1].View.Last();
                    slide = pptApplication.SlideShowWindows[1].View.Slide;

                    status = "true";
                }


            }
            catch (Exception)
            {
                status = "false";

            }

            return status;
        }
        
        public static string currentSlide()
        {
            var status = "";
            try
            {
                var pptApplication = Marshal.GetActiveObject("PowerPoint.Application") as PPt.Application;
                

                // Get Presentation Object 
                var presentation = pptApplication.ActivePresentation;
                // Get Slide collection object 
                var slides = presentation.Slides;
                // Get Slide count 
                var slidescount = slides.Count;
                // Get current selected slide  
                try
                {
                    // Get selected slide object in normal view 
                    var slide = slides[pptApplication.ActiveWindow.Selection.SlideRange.SlideNumber];

                    status = slide.SlideNumber.ToString();

                }
                catch
                {
                    // Get selected slide object in reading view 
                    var slide = pptApplication.SlideShowWindows[1].View.Slide;

                    status = slide.SlideNumber.ToString();
                }


            }
            catch (Exception)
            {
                status = "false";

            }

            return status;
        }
    }
}
