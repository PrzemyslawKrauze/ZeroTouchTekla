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
    public static class TeklaUtils
    {
        public static RebarSet CreateDefaultRebarSet(string name, int rebarSize)
        {
            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = name;
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(rebarSize);
            rebarSet.RebarProperties.Size = rebarSize.ToString();
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(rebarSize);
            rebarSet.LayerOrderNumber = 1;
            return rebarSet;
        }
        public static int SetClass(double diameter)
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
        public static double GetBendingRadious(double diameter)
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
    }
    public abstract class Element
    {
        public Element(Part part)
        {
            SetRebarCreatorProperties(part);
        }
        public Element(List<Part> parts)
        {
        }
        //Abstract methods
        public abstract void Create();
        public abstract void CreateSingle(string barName);
        //Public methods
        public void SetLocalPlane()
        {
            Model model = new Model();
            TransformationPlane localPlane = new TransformationPlane(BasePart.GetCoordinateSystem());
            model.GetWorkPlaneHandler().SetCurrentTransformationPlane(localPlane);
        }
       
        protected void SetRebarCreatorProperties(Part beam)
        {
            beam.GetUserProperty("__CovThickSides", ref SideCover);
            beam.GetUserProperty("__CovThickBottom", ref BottomCover);
            beam.GetUserProperty("__CovThickTop ", ref TopCover);
        }
        protected static string[] GetProfileValues(Part beam)
        {
            Profile profile = beam.Profile;
            string profileName = profile.ProfileString;
            profileName = Regex.Replace(profileName, "[A-Za-z ]", "");
            string[] profileValues = profileName.Split('*');
            return profileValues;
        }
       
        protected static double GetHookLength(double diameter)
        {
            return 10 * diameter;
        }        
        protected void PostRebarCreationMethod(RebarSet rebarSet, System.Reflection.MethodBase methodBase)
        {
            rebarSet.SetUserProperty(RebarCreator.FATHER_ID_NAME, RebarCreator.FatherID);
            rebarSet.SetUserProperty(RebarCreator.METHOD_NAME, methodBase.Name);
            string diameter = rebarSet.RebarProperties.Size;
            rebarSet.SetUserProperty("__MIN_BAR_LENTYPE", 0);
            rebarSet.SetUserProperty("__MIN_BAR_LENGTH", RebarCreator.MinLengthCoefficient * Convert.ToDouble(diameter));

        }
        protected void PostRebarCreationMethod(RebarSet rebarSet, System.Reflection.MethodBase methodBase, int input)
        {
            rebarSet.SetUserProperty(RebarCreator.FATHER_ID_NAME, RebarCreator.FatherID);
            rebarSet.SetUserProperty(RebarCreator.METHOD_NAME, methodBase.Name);
            rebarSet.SetUserProperty(RebarCreator.MethodInput, input);
            string diameter = rebarSet.RebarProperties.Size;
            rebarSet.SetUserProperty("__MIN_BAR_LENTYPE", 0);
            rebarSet.SetUserProperty("__MIN_BAR_LENGTH", RebarCreator.MinLengthCoefficient * Convert.ToDouble(diameter));
        }

        private List<List<Point>> profilePoints = new List<List<Point>>();
        protected Dictionary<string, double> ProfileParameters = new Dictionary<string, double>();
        private  Dictionary<int, int[]> layerDictionary = new Dictionary<int, int[]>();
        private Part basePart;
        public double SideCover = 0;
        public double BottomCover = 0;
        public double TopCover = 0;
        public ElementFace ElementFace;

       
        public  Dictionary<int, int[]> LayerDictionary { get => layerDictionary; }
        protected Part BasePart { get => basePart; set => basePart = value; }
        public List<List<Point>> ProfilePoints { get => profilePoints; set => profilePoints = value; }
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

}
