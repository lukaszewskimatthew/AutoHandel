using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoHandel
{
    class Lunch
    {
        string sName;
        string bFase;
        string status;
        string dist;
        string meal;
        string addit;
        string milk;

        public Lunch(string inName, string inFast, string inStat, string inDist,
                     string inMeal, string inAddit, string inMilk)
        {
            sName = inName;
            bFase = inFast;
            status = inStat;
            dist = inDist;
            meal = inMeal;
            addit = inAddit;
            milk = inMilk;
        }

        public string Name { get { return sName; } }

        public override string ToString()
        {
            return sName +
                   "\n   Meal: " + bFase +
                   "\n   Payment Code: " + status +
                   "\n   Distract: " + dist +
                   "\n   Food Order: " + meal +
                   "\n   Additional Options: " + addit +
                   "\n   Milk: " + milk;
        }
    }
}
