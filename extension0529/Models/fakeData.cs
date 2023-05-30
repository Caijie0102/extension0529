using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Web;

namespace extension0529.Models
{
    public class fakeData
    {
        [DisplayName("Name")]
        public string Name { get; set; }

        [DisplayName("Age")]
        public int Age { get; set; }

        [DisplayName("Email")]
        public string Email { get; set; }
    }

    public class BigView1
    {
        public List<fakeData> Report1 { get; set; }
        public List<fakeData> Report2 { get; set; }
        public BigView1()
        {
            this.Report1 = new List<fakeData>();
            this.Report2 = new List<fakeData>();
        }
    }
}