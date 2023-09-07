using Newtonsoft.Json;
using RESTCountries.NET.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.Linq;
using System.Web;

namespace WebApplication2.Models
{
    public class formsubmito
    {
        [Display(Name = "الاٍسم")]
        public string Name { get; set; }
        [Display(Name = "رقم الملف")]
        public string FileNumber { get; set; }
        [Display(Name = "العمر")]
        public string Age { get; set; }
        [Display(Name = "الجنسية")]
        public string Nationality { get; set; }
        [Display(Name = "التاريخ")]
        public string Date { get; set; }
        [Display(Name = "العيادة")]
        public string Clinic{ get; set; }


        public IEnumerable<Color> Colors = GetCountryes();

        private static IEnumerable<Color> GetCountryes()
        {
            var query2 = from c in RestCountriesService.GetAllCountries()
                         select new Color
                         {
                            name = c.Translations["ara"].Common
                         };
            return query2;
        }
    }
    public class Color
    {
        public string name { get; set; }
    }
    // Root myDeserializedClass = JsonConvert.DeserializeObject<List<Root>>(myJsonResponse);
    public class Ara
    {
        public string official { get; set; }
        public string common { get; set; }
    }

}