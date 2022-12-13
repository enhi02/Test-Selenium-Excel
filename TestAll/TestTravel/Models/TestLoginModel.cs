using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestTravel.Models
{
    public class TestLoginModel
    {
        public string TaiKhoanDung { get; set; }
        public string MatKhauDung { get; set; }

    }
    public class TestLoginFailedModel
    {
        public string TaiKhoanSai { get; set; }
        public string MatKhauSai { get; set; }

    }
}
