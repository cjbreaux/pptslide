using PixabaySharp;
using PixabaySharp.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace PowerPointWebAddIn2Web.Controllers
{
    public class PixabayController : ApiController
    {
        public async System.Threading.Tasks.Task<object> GetPhotosAsync()
        {
            var client = new PixabaySharpClient("13128735-e8d30252cb1da0dadc588a8d8");
            var result = await client.QueryImagesAsync(new ImageQueryBuilder()
            {
                Query = "car", // Need to pass in input here
                Page = 1,
                PerPage = 5
            });
            return result;
        }
         
    }
}
