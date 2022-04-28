﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using Tekla.Structures;
using Tekla.Structures.Solid;
using Tekla.Structures.Model;
using Tekla.Structures.Geometry3d;

using ZeroTouchTekla.Profiles;

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
        public static Face[] GetPartEndFaces(Part part)
        {
            Solid soild = part.GetSolid();
            FaceEnumerator faceEnumerator = soild.GetFaceEnumerator();
            List<Face> faces = GetFacesFromFaceEnumerator(faceEnumerator);
            double maxVertex = 0;
            double minVertex = 0;
            Face startFace = null;
            Face endFace = null;
            for (int i = 0; i < faces.Count; i++)
            {
                Face currentFace = faces[i];
                LoopEnumerator loopEnumerator = currentFace.GetLoopEnumerator();
                List<Point> points = GetPointsFromLoopEnumerator(loopEnumerator);
                double vertexXsum = 0;
                foreach (Point p in points)
                {
                    vertexXsum += p.X;
                }
                if (vertexXsum >= maxVertex)
                {
                    maxVertex = vertexXsum;
                    endFace = currentFace;
                }
                if (vertexXsum <= minVertex)
                {
                    minVertex = vertexXsum;
                    startFace = currentFace;
                }
            }
            Face[] faceArray = new Face[] { startFace, endFace };
            return faceArray;
        }
        public static List<List<Point>> GetPointsFromFaces(Face[] faces)
        {
            List<List<Point>> pointLists = new List<List<Point>>();
            for (int i = 0; i < faces.Count(); i++)
            {
                List<Point> points = new List<Point>();
                LoopEnumerator loopEnum = faces[i].GetLoopEnumerator();
                loopEnum.MoveNext();
                Loop loop = loopEnum.Current;
                VertexEnumerator vertexEnumerator = loop.GetVertexEnumerator();
                while (vertexEnumerator.MoveNext())
                {
                    points.Add(vertexEnumerator.Current);
                }
                pointLists.Add(points);
            }
            return pointLists;
        }
        public static List<List<Point>> SortPoints(List<List<Point>> pointsList)
        {
            List<List<Point>> sortedPoints = new List<List<Point>>();
            foreach(List<Point> points in pointsList)
            {
                List<Point> sortedList = (from Point p in points
                                 orderby p.X, p.Y, p.Z ascending
                                 select p).ToList();
                sortedPoints.Add(sortedList);
            }
            return sortedPoints;
        }
        public static List<Face> GetFacesFromFaceEnumerator(FaceEnumerator faceEnumerator)
        {
            List<Face> faces = new List<Face>();
            while (faceEnumerator.MoveNext())
            {
                faces.Add(faceEnumerator.Current);
            }
            return faces;
        }
        public static List<Point> GetPointsFromLoopEnumerator(LoopEnumerator loopEnumerator)
        {
            List<Point> points = new List<Point>();
            loopEnumerator.MoveNext();
            VertexEnumerator vertexEnumerator = loopEnumerator.Current.GetVertexEnumerator();

            while (vertexEnumerator.MoveNext())
            {
                points.Add(vertexEnumerator.Current);
            }
            return points;
        }
        public static List<List<Point>> GetSortedPointsFromPart(Part part)
        {
            Face[] faces = TeklaUtils.GetPartEndFaces(part);
            List<List<Point>> points = TeklaUtils.GetPointsFromFaces(faces);
            points = TeklaUtils.SortPoints(points);
            return points;
        }
    }
    public abstract class Element
    {
        //Constant
        private string COV_THICK_SIDES = "__CovThickSides";
        private string COV_THICK_BOTTOM = "__CovThickBottom";
        private string COV_THICK_TOP = "__CovThickTop";
        //Fields
        private List<List<Point>> profilePoints = new List<List<Point>>();
        private Dictionary<string, double> profileParameters = new Dictionary<string, double>();
        private Dictionary<int, int[]> layerDictionary = new Dictionary<int, int[]>();
        private Part[] baseParts;
        //Constructors
        public static Element Create(params Part[] parts)
        {
            string partName = parts[0].Profile.ProfileString;

            ProfileType profileType = GetProfileType(partName);
            Element element;
            switch (profileType)
            {
                case ProfileType.FTG:
                    element = new FTG(parts);
                    break;
                case ProfileType.RTW:
                    if (parts.Length > 1)
                    {
                        element = new DRTW(parts);
                    }
                    else
                    {
                        element = new RTW(parts);
                    }
                    break;
                case ProfileType.DRTW:
                    element = new DRTW(parts);
                    break;
                case ProfileType.RTWS:
                    element = new RTWS(parts);
                    break;
                case ProfileType.RCLMN:
                    element = new CLMN(parts);
                    break;
                case ProfileType.ABT:
                    element = new ABT(parts);
                    break;
                case ProfileType.APS:
                    element = new APS(parts);
                    break;
                default:
                    throw new Exception("Profile type doesn't match");
            }
            element.SetCover(parts[0]);

            element.Create();


            return element;
        }
        protected Element() { }
        public static Element.ProfileType GetProfileType(string profileString)
        {
            switch (profileString)
            {
                case var _ when profileString.StartsWith("FTG"):
                    return ProfileType.FTG;
                case var _ when profileString.StartsWith("RTWS"):
                    return ProfileType.RTWS;
                case var _ when profileString.StartsWith("RTW"):
                    return ProfileType.RTW;
                case var _ when profileString.StartsWith("RCLMN"):
                    return ProfileType.RCLMN;
                case var _ when profileString.StartsWith("ABT"):
                    return ProfileType.ABT;
                case var _ when profileString.StartsWith("TABT"):
                    return ProfileType.TABT;
                case var _ when profileString.StartsWith("APS"):
                    return ProfileType.APS;
                default:
                    return ProfileType.None;
            }
        }
        //Properties
        public enum ProfileType
        {
            None,
            FTG,
            RTW,
            DRTW,
            RTWS,
            RCLMN,
            ABT,
            TABT,
            WING,
            APS
        }
        public Dictionary<int, int[]> LayerDictionary { get => layerDictionary; }
        protected Part[] BaseParts { get => baseParts; set => baseParts = value; }
        public List<List<Point>> ProfilePoints { get => profilePoints; set => profilePoints = value; }
        protected Dictionary<string, double> ProfileParameters { get => profileParameters; set => profileParameters = value; }
        public double SideCover = 0;
        public double BottomCover = 0;
        public double TopCover = 0;
        public ElementFace ElementFace;
        //Initialization methods
        protected void SetCover(Part beam)
        {
            beam.GetUserProperty(COV_THICK_SIDES, ref SideCover);
            beam.GetUserProperty(COV_THICK_BOTTOM, ref BottomCover);
            beam.GetUserProperty(COV_THICK_TOP, ref TopCover);
        }
        protected string[] GetProfileValues(Part beam)
        {
            Profile profile = beam.Profile;
            string profileName = profile.ProfileString;
            profileName = Regex.Replace(profileName, "[A-Za-z ]", "");
            string[] profileValues = profileName.Split('*');
            return profileValues;
        }

        //Abstract methods
        public abstract void Create();
        public abstract void CreateSingle(string barName);
        //Public methods
        public void SetLocalPlane()
        {
            Model model = new Model();
            TransformationPlane localPlane = new TransformationPlane(BaseParts.FirstOrDefault().GetCoordinateSystem());
            model.GetWorkPlaneHandler().SetCurrentTransformationPlane(localPlane);
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
        private RebarLegFace GetRebarLegFace(int faceNumber)
        {
            return rebarLegFaces[faceNumber];
        }

        List<RebarLegFace> rebarLegFaces;
       
    }

}
