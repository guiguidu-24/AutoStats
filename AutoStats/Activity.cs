using System;

namespace AutoStats
{
    /// <summary>
    /// Represent the sport data of an activity
    /// </summary>
    public struct Activity
    {
        public int? HeartRate { get; set; }
        public int? HeartMax { get; set; }
        public int? Distance { get; set; }
        public int Time { get; set; }
        public string Name { get; set; }
        
        
        private double? avrSpeed;
        public double? AvrSpeed
        {
            get
            {
                return avrSpeed;
            }
            set
            {
                if (value != null)
                    avrSpeed = Math.Round((double) value, 1); //round the value with one digit
                else
                    avrSpeed = null;
            }
        }
    }
}