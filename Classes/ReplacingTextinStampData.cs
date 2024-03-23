using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PositionApplicability.Classes
{
    internal class ReplacingTextinStampData
    {
        private string profileFind = "";
        private string thicknessFind = "";
        private string gostProfileFind = "";
        private string steelFind = "";
        private string gostSteelFind = "";

        private string profileReplace = "";
        private string thicknessReplace = "";
        private string gostProfileReplace = "";
        private string steelReplace = "";
        private string gostSteelReplace = "";

        private double height = 5;


        private bool isProfile = false;
        private bool isThickness = false;
        private bool isGostProfile = false;
        private bool isSteel = false;
        private bool isGostSteel = false;

        public bool IsProfile { get => isProfile; set => isProfile = value; }
        public bool IsThickness { get => isThickness; set => isThickness = value; }
        public bool IsGostProfile { get => isGostProfile; set => isGostProfile = value; }
        public bool IsSteel { get => isSteel; set => isSteel = value; }
        public bool IsGostSteel { get => isGostSteel; set => isGostSteel = value; }
        public string ProfileFind { get => profileFind; set => profileFind = value; }
        public string ThicknessFind { get => thicknessFind; set => thicknessFind = value; }
        public string GostProfileFind { get => gostProfileFind; set => gostProfileFind = value; }
        public string SteelFind { get => steelFind; set => steelFind = value; }
        public string GostSteelFind { get => gostSteelFind; set => gostSteelFind = value; }
        public string ProfileReplace { get => profileReplace; set => profileReplace = value; }
        public string ThicknessReplace { get => thicknessReplace; set => thicknessReplace = value; }
        public string GostProfileReplace { get => gostProfileReplace; set => gostProfileReplace = value; }
        public string SteelReplace { get => steelReplace; set => steelReplace = value; }
        public string GostSteelReplace { get => gostSteelReplace; set => gostSteelReplace = value; }
        public double Height { get => height; set => height = value; }
    }
}
