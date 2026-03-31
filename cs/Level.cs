using System;

namespace Game.Config
{
    [Serializable]
    public class Level
    {
        public int Id;
        public int Stage;
        public int IsBoss;
        public string Get;
        public string Map;
        public string AiNum;
        public string AiSpeed;
        public string AiControl;
        public string AiDurability;
        public string AiAttack;
    }
}
