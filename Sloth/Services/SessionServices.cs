using Microsoft.AspNetCore.Http;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;

namespace Sloth.Services
{
    public static class SessionExtensionMethods
    {
        public static void SetObject(this ISession session, string key, object value)
        {
            string stringValue = JsonConvert.SerializeObject(value);
            session.SetString(key, stringValue);
        }

        public static T GetObject<T>(this ISession session, string key)
        {
            string stringValue = session.GetString(key);
            T value = JsonConvert.DeserializeObject<T>(stringValue);
            return value;
        }
    }


    public static class BCFServices
    {
        public static string GetImageFromBytes(byte[] bytes, bool compress)
        {
            System.IO.MemoryStream myMemStream = new System.IO.MemoryStream(bytes);
            System.Drawing.Image fullsizeImage = System.Drawing.Image.FromStream(myMemStream);

            if (compress)
            {
                System.Drawing.Image newImage = fullsizeImage.GetThumbnailImage(512, 512, null, IntPtr.Zero);
                myMemStream = new System.IO.MemoryStream();
                newImage.Save(myMemStream, System.Drawing.Imaging.ImageFormat.Jpeg);  //Or whatever format you want.
            }


            string base64 = Convert.ToBase64String(myMemStream.ToArray());
            return String.Format("data:image/gif;base64,{0}", base64);
        }

    }
}
