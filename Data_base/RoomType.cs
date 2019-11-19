using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Data_base
{
    [Serializable]
    public class RoomType
    {
        public string room_type { get; set; }
        public string room_name { get; set; }
        public string room_spe { get; set; }
        public string room_price { get; set; }

        public RoomType(string room_type, string room_name, string room_spe, string room_price)
        {
            this.room_type = room_type;
            this.room_name = room_name;
            this.room_spe = room_spe;
            this.room_price = room_price;
        }
    }
}
