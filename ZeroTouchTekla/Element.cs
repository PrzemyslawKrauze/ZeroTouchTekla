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
        public Element(Beam part)
        {
            GetProfilePointsAndParameters(part);
            SetRebarCreatorProperties(part);
        }
        public void Create()
        { }
        public void CreateSingle(string barName)
        { }
        protected static void GetProfilePointsAndParameters(Beam beam)
        {
            ProfilePoints = new List<List<Point>>();
            ProfileParameters = new Dictionary<string, double>();
            ElementFace = new ElementFace(ProfilePoints);
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
        public static ElementFace ElementFace;
        enum RebarType
        {
        };
    }
    public class ElementFace
    {
        public ElementFace(List<List<Point>> profilePoints)
        {
            if (profilePoints.Count != 0)
            {
                rebarLegFaces = new List<RebarLegFace>();
                RebarLegFace startFace = new RebarLegFace();
                int numberOfPoints = profilePoints[0].Count;
                for (int i = 0; i < numberOfPoints; i++)
                {
                    startFace.Contour.AddContourPoint(new ContourPoint(profilePoints[0][i], null));
                }
                rebarLegFaces.Add(startFace);

                for (int i = 0; i < numberOfPoints - 1; i++)
                {
                    Point firstPoint = profilePoints[0][i];
                    Point secondPoint = profilePoints[1][i];
                    Point thirdPoint = profilePoints[1][i + 1];
                    Point fourthPoint = profilePoints[0][i + 1];

                    var rebarLegFace = new RebarLegFace();
                    rebarLegFace.Contour.AddContourPoint(new ContourPoint(firstPoint, null));
                    rebarLegFace.Contour.AddContourPoint(new ContourPoint(secondPoint, null));
                    rebarLegFace.Contour.AddContourPoint(new ContourPoint(thirdPoint, null));
                    rebarLegFace.Contour.AddContourPoint(new ContourPoint(fourthPoint, null));
                    rebarLegFaces.Add(rebarLegFace);
                }


                RebarLegFace face = new RebarLegFace();
                face.Contour.AddContourPoint(new ContourPoint(profilePoints[0][numberOfPoints - 1], null));
                face.Contour.AddContourPoint(new ContourPoint(profilePoints[1][numberOfPoints - 1], null));
                face.Contour.AddContourPoint(new ContourPoint(profilePoints[1][0], null));
                face.Contour.AddContourPoint(new ContourPoint(profilePoints[0][0], null));
                rebarLegFaces.Add(face);

                RebarLegFace endFace = new RebarLegFace();
                for (int i = 0; i < numberOfPoints; i++)
                {
                    endFace.Contour.AddContourPoint(new ContourPoint(profilePoints[1][i], null));
                }
                rebarLegFaces.Add(endFace);
            }
        }
        public RebarLegFace GetRebarLegFace(int faceNumber)
        {
            return rebarLegFaces[faceNumber];
        }

        List<RebarLegFace> rebarLegFaces;

    }
    public class BaseParameter
    {
        public const string None = "None";
    }
}
