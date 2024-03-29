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
                return 7 * diameter/2.0;
            }
            else
            {
                return 4 * diameter/2.0;
            }
        }
        public static Face[] GetPartEndFaces(Part part)
        {
            Solid soild = part.GetSolid(Solid.SolidCreationTypeEnum.RAW);
            FaceEnumerator faceEnumerator = soild.GetFaceEnumerator();
            List<Face> faces = GetFacesFromFaceEnumerator(faceEnumerator);
            double maxVertex = Double.MinValue;
            double minVertex = Double.MaxValue;
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
        public static List<Point> SortPoints(List<Point> points)
        {
            List<Point> sortedList = (from Point p in points
                                      orderby Math.Round(p.X) , Math.Round(p.Y), Math.Round(p.Z) ascending
                                      select p).ToList();
            return sortedList;
        }
        public static List<List<Point>> SortPoints(List<List<Point>> pointsList)
        {
            List<List<Point>> sortedPoints = new List<List<Point>>();
            foreach (List<Point> points in pointsList)
            {
                sortedPoints.Add(SortPoints(points));
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
        public static List<List<Point>> GetSortedPointsFromEndFaces(Part part)
        {
            Face[] faces = TeklaUtils.GetPartEndFaces(part);
            List<List<Point>> points = TeklaUtils.GetPointsFromFaces(faces);
            points = TeklaUtils.SortPoints(points);
            return points;
        }
        public static List<List<Point>> GetSortedPointsFromEndFaces(Part[] parts)
        {
            List<List<Point>> sortedPoints = new List<List<Point>>();
            for (int i = 0; i < parts.Count(); i++)
            {
                List<List<Point>> points = GetSortedPointsFromEndFaces(parts[i]);
                if (i == 0)
                {
                    sortedPoints.Add(points.First());
                }
                sortedPoints.Add(points.Last());
            }
            return sortedPoints;
        }
        public static List<List<Point>> GetSortedPointsFromPart(Part part)
        {
            Solid soild = part.GetSolid(Solid.SolidCreationTypeEnum.NORMAL_WITHOUT_EDGECHAMFERS);
            FaceEnumerator faceEnumerator = soild.GetFaceEnumerator();
            List<Face> faces = GetFacesFromFaceEnumerator(faceEnumerator);
            Face[] faceArray = faces.ToArray();
            List<List<Point>> facePoints = GetPointsFromFaces(faceArray);
            List<Point> unsortedPoints = new List<Point>();
            foreach (List<Point> point in facePoints)
            {
                unsortedPoints.AddRange(point);
            }
            unsortedPoints = RemoveDuplicatedPoints(unsortedPoints);
            List<List<Point>> sortedPoints = new List<List<Point>>();
            double currentXValue = Double.MinValue;
            List<Point> currentList = new List<Point>();
            for (int i = 0; i < unsortedPoints.Count; i++)
            {
                if (Math.Abs(unsortedPoints[i].X -currentXValue)<1)
                {
                    currentList.Add(unsortedPoints[i]);
                }
                else
                {
                    currentXValue = unsortedPoints[i].X;
                    if (i != 0)
                    {
                        sortedPoints.Add(currentList);
                        currentList = new List<Point>();
                    }
                    currentList.Add(unsortedPoints[i]);

                }
            }
            sortedPoints.Add(currentList);
            return sortedPoints;
        }
        public static List<Point> RemoveDuplicatedPoints(List<Point> pointList)
        {
            pointList = SortPoints(pointList);
            for (int i = 0; i < pointList.Count - 1; i++)
            {
                if (Distance.PointToPoint(pointList[i], pointList[i + 1]) <= 1)
                {
                    pointList.RemoveAt(i);
                    i--;
                }
            }
            return pointList;
        }
        public static List<Line> GetLinesFromPolygonPoints(Polygon polygon)
        {
            List<Line> lines = new List<Line>();
            for(int i=0;i< polygon.Points.Count-1;i++)
            {
                Point firstPoint = polygon.Points[i] as Point;
                Point secondPoint = polygon.Points[i + 1] as Point;
                Line line = new Line(firstPoint, secondPoint);
                lines.Add(line);
            }
            return lines;
        }
        public static Plane GetPlaneFromContour(Contour contour)
        {
            Plane plane = new Plane();
            plane.Origin = contour.ContourPoints[0] as Point;
            Vector axisX = Utility.GetVectorFromTwoPoints(contour.ContourPoints[0] as Point, contour.ContourPoints[1] as Point);
            plane.AxisX = axisX;
            plane.AxisY = Utility.GetVectorFromTwoPoints(contour.ContourPoints[0] as Point,contour.ContourPoints[2] as Point);
            return plane;
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
       private List<RebarLegFace> rebarLegFaces = new List<RebarLegFace>();
        //Constructors
        public static Element Initialize(params Part[] parts)
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
                    element = new RCLMN(parts);
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
        public List<RebarLegFace> RebarLegFaces { get => rebarLegFaces; set => rebarLegFaces = value; }

        protected double SideCover = 0;
        protected double BottomCover = 0;
        protected double TopCover = 0;
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
        //Protected methods
        protected RebarLegFace GetRebarLegFace(int number)
        {
            RebarLegFace faceToCopy = RebarLegFaces[number];
            RebarLegFace rebarLegFace = new RebarLegFace();
            rebarLegFace.Contour = faceToCopy.Contour;
            return rebarLegFace;
        }
        protected void SetLocalPlane()
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
        protected static RebarSet CreateDefaultRebarSet(string name, int rebarSize)
        {
            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = name;
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = TeklaUtils.SetClass(rebarSize);
            rebarSet.RebarProperties.Size = rebarSize.ToString();
            rebarSet.RebarProperties.BendingRadius = TeklaUtils.GetBendingRadious(rebarSize);
            rebarSet.LayerOrderNumber = 1;
            return rebarSet;
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
