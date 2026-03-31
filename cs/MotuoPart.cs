using System;

namespace Game.Config
{
    [Serializable]
    public class MotuoPart
    {
        public int Id;
        public int Type;
        public int Level;
        public int GoldCost;
        public string OtherCost;
        public int IsAd;
    }
}
