using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KGSA
{
    public class TimeWatch
    {
        DateTime currentTime;
        DateTime update;

        public string show
        {
            get
            {
                double ts = (DateTime.Now - currentTime).TotalSeconds;
                return string.Format("{0:n0}", Math.Round(ts, 2));
            }
        }

        public TimeWatch()
        {
        }

        public void Start()
        {
            currentTime = DateTime.Now;
            update = DateTime.Now;
        }

        public double Stop()
        {
            double ts = (DateTime.Now - currentTime).TotalSeconds;
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
