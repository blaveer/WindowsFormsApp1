using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Event_handing
{
    class FireEventArgs : EventArgs
    {
        public FireEventArgs(string room, int ferocity)
        {
            this.room = room;
            this.ferocity = ferocity;
        }

        public string room;
        public int ferocity;
    }
    class FireAlarm
    {
        public delegate void FireEventHandler(object sender, FireEventArgs fe);
        public event FireEventHandler FireEvent;

        public void ActivateFireAlarm(string room, int ferocity)
        {
            FireEventArgs fireArgs = new FireEventArgs(room, ferocity);
            //执行对象事件处理函数指针，必须保证处理函数要和声明代理时的参数列表相同
            //相当于调用函数
            FireEvent(this, fireArgs);
        }
    }
}
