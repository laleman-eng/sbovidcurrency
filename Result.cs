using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SBO_VID_Currency
{
   public class Result
   {
       public Id id { get; set; }
       public string fecha { get; set; }
       public string serie { get; set; }
       public object descripcion { get; set; }
       public string codigo { get; set; }
       public double valor { get; set; }

       public class Id
       {
           public int timestamp { get; set; }
           public int machine { get; set; }
           public int pid { get; set; }
           public int increment { get; set; }
           public DateTime creationTime { get; set; }
       }

   }
}
