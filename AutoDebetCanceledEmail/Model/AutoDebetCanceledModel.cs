using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoDebetCanceledEmail.Model
{
    class AutoDebetCanceledModel
    {
        public string DeliveryBy { get; set; }
        public int? SuccessSent { get; set; }
        public int? SuccessDelivered { get; set; }
    }
}
