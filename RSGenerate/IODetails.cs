using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RSGenerate
{
    class IOPoint
    {
        public string Bit { get; set; }
        public string Description { get; set; }
    }

    class IOCard
    {
        public string CardType { get; set; }
        public string Rack { get; set; }
        public string Module { get; set; }
        public List<IOPoint> Points { get; set; }

        public IOCard()
        {
            Points = new List<IOPoint>();
        }
    }
}
