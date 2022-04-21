using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LeerExcel.Models
{
    class Plantilla
    {
        public int CodEmpleado { get; set; }
        public string Nombre { get; set; }
        public int TotalDias { get; set; }
        public int DiasTrabajados { get; set; }
        public int DiasNoTrabajados { get; set; }
    }
}
