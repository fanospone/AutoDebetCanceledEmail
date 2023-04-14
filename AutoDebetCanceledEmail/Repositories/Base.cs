using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoDebetCanceledEmail.Repositories
{
    public class Base
    {
        public string conn;
        public Base()
        {
            conn = ConfigurationManager.ConnectionStrings["connectionString"].ConnectionString;
        }
    }
}
