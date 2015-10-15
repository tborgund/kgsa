using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KGSA
{
    class TimeWatch
    {
        DateTime time;
        DateTime update;

        public TimeWatch()
        {
        }

        public void Start()
        {
            time = DateTime.Now;
            update = DateTime.Now;
        }

        public double Stop()
        {
            double ts = (DateTime.Now - time).TotalSeconds;
            return Math.Round(ts, 2);
        }

        public bool ReadyForRefresh()
        {
            double tms = (DateTime.Now - update).TotalMilliseconds;
            update = DateTime.Now;
            if (tms > 900)
                return true;
            return false;
        }
    }
}
