using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using Tekla.Structures;
using Tekla.Structures.Model;
using Tekla.Structures.Geometry3d;

namespace ZeroTouchTekla
{
    public class Element
    {
        public Element(Beam beam)
        {
            GetProfilePointsAndParameters(beam);
            SetRebarCreatorProperties(beam);
        }
        public void Create()
        { }
        public void CreateSingle(string barName)
        { }
        protected static void GetProfilePointsAndParameters(Beam beam)
        {
            ProfilePoints = new List<List<Point>>();
            ProfileParameters = new Dictionary<string, double>();
        }
        protected static void SetRebarCreatorProperties(Beam beam)
        {
            beam.GetUserProperty("__CovThickSides", ref SideCover);
            beam.GetUserProperty("__CovThickBottom", ref BottomCover);
            beam.GetUserProperty("__CovThickTop ", ref TopCover);
        }
        protected static string[] GetProfileValues(Beam beam)
        {
            Profile profile = beam.Profile;
            string profileName = profile.ProfileString;
            profileName = Regex.Replace(profileName, "[A-Za-z ]", "");
            string[] profileValues = profileName.Split('*');
            return profileValues;
        }
        protected static double GetBendingRadious(double diameter)
        {
            if (diameter > 16)
            {
                return 7 * diameter;
            }
            else
            {
                return 4 * diameter;
            }
        }
        protected static int SetClass(double diameter)
        {
            switch (diameter)
            {
                case 10:
                    return 1;
                case 12:
                    return 2;
                case 16:
                    return 6;
                case 20:
                    return 5;
                case 25:
                    return 11;
                case 28:
                    return 3;
                case 32:
                    return 7;
                default:
                    return 8;
            }
        }

        protected static List<List<Point>> ProfilePoints;
        protected static Dictionary<string, double> ProfileParameters;
        public static double SideCover = 0;
        public static double BottomCover = 0;
        public static double TopCover = 0;
        enum RebarType
        {
        };
    }
    public class BaseParameter
    {
        public const string None = "None";
    }
}
