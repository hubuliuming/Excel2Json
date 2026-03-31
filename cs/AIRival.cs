using System;

namespace Game.Config
{
    [Serializable]
    public class AIRival
    {
        public int Id;
        public int AiLevel;
        public float AiAdd;
        public int CommandTime;
        public float SensingDistance;
        public int ChangeTime;
        public float PlayerDistance;
        public int SpeedTime;
        public int SpeedTimeCD;
    }
}
