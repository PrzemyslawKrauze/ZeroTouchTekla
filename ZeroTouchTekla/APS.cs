using System;
using System.Collections.Generic;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model;
using System.Linq;
using System.Reflection;

namespace ZeroTouchTekla
{
    public class APS : Element
    {
        public APS(Beam part) : base(part)
        {
            GetProfilePointsAndParameters(part);
        }
        public void GetProfilePointsAndParameters(Beam beam)
        {
            //ABSV Height*Width*Height2*TopWidth*Thickness*SkewWidth*Drop*HorizontalOffset*VerticaOffset

            FirstBeamProperties(beam);
            ElementFace = new ElementFace(ProfilePoints);
        }
        new public void Create()
        {
            /*
            LongitudinalRebar(0);
            LongitudinalRebar(1);
            LongitudinalRebar(2);
            LongitudinalRebar(3);
            LongitudinalRebar(4);
            */
            FirstPerpendicularRebar(0);
        }
        #region PrivateMethods
        void FirstBeamProperties(Beam beam)
        {
            string[] profileValues = GetProfileValues(beam);
            Height = Convert.ToDouble(profileValues[0]);
            Width = Convert.ToDouble(profileValues[1]);
            Height2 = Convert.ToDouble(profileValues[2]);
            TopWidth = Convert.ToDouble(profileValues[3]);
            SkewWidth = Convert.ToDouble(profileValues[4]);
            Drop = Convert.ToDouble(profileValues[5]);
            Thickness = Convert.ToDouble(profileValues[6]);
            HorizontalOffset = Convert.ToDouble(profileValues[7]);
            VerticalOffset = Convert.ToDouble(profileValues[8]);
            Length = Distance.PointToPoint(beam.StartPoint, beam.EndPoint);
            FullLength = Length;
            Angle = Math.Atan(Drop / 100);
            FrontHeight = Thickness / Math.Cos(Angle);
            double bottomHeight = Width * Drop / 100.0;

            double distanceToMid = (Height2 + bottomHeight) / 2.0;

            Point p0 = new Point(0, -distanceToMid, -Width / 2.0);
            Point p1 = new Point(0, p0.Y + FrontHeight, p0.Z);
            Point p2 = new Point(0, p1.Y + (Width - TopWidth - SkewWidth) * Drop / 100, p1.Z + Width - TopWidth - SkewWidth);
            Point p3 = new Point(0, p0.Y + bottomHeight + Height2, p1.Z + Width - TopWidth);
            Point p4 = new Point(0, p0.Y + bottomHeight + Height, p0.Z + Width);
            Point p5 = new Point(0, p0.Y + bottomHeight, p4.Z);
            List<Point> firstProfile = new List<Point> { p0, p1, p2, p3, p4, p5 };

            Point n0 = new Point(Length, -distanceToMid, -Width / 2.0);
            Point n1 = new Point(Length, n0.Y + FrontHeight, n0.Z);
            Point n2 = new Point(Length, n1.Y + (Width - TopWidth - SkewWidth) * Drop / 100, n1.Z + Width - TopWidth - SkewWidth);
            Point n3 = new Point(Length, n0.Y + bottomHeight + Height2, n1.Z + Width - TopWidth);
            Point n4 = new Point(Length, n0.Y + bottomHeight + Height, n0.Z + Width);
            Point n5 = new Point(Length, n0.Y + bottomHeight, n4.Z);
            List<Point> secondProfile = new List<Point> { n0, n1, n2, n3, n4, n5 };

            if (HorizontalOffset != 0)
            {
                foreach (Point p in firstProfile)
                {
                    p.Translate(0, 0, -HorizontalOffset / 2.0);
                }
                foreach (Point p in secondProfile)
                {
                    p.Translate(0, 0, HorizontalOffset / 2.0);
                }
            }

            ProfilePoints.Add(firstProfile);
            ProfilePoints.Add(secondProfile);
        }
        void LongitudinalRebar(int faceNumber)
        {
            string rebarSize = Program.ExcelDictionary["LR_Diameter"];
            string spacing = Program.ExcelDictionary["LR_Spacing"];

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "APS_LR_" + faceNumber;
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            int f, s;
            switch (faceNumber)
            {
                case 0:
                    f = 0;
                    s = 5;
                    break;
                case 1:
                    f = 1;
                    s = 2;
                    break;
                case 2:
                    f = 2;
                    s = 3;
                    break;
                case 3:
                    f = 3;
                    s = 4;
                    break;
                case 4:
                    f = 4;
                    s = 5;
                    break;
                default:
                    f = 0;
                    s = 0;
                    break;
            }

            for (int i = 0; i < ProfilePoints.Count - 1; i++)
            {
                Point p1 = ProfilePoints[i][f];
                Point p2 = ProfilePoints[i + 1][f];
                Point p3 = ProfilePoints[i + 1][s];
                Point p4 = ProfilePoints[i][s];
                var face = new RebarLegFace();
                face.Contour.AddContourPoint(new ContourPoint(p1, null));
                face.Contour.AddContourPoint(new ContourPoint(p2, null));
                face.Contour.AddContourPoint(new ContourPoint(p3, null));
                face.Contour.AddContourPoint(new ContourPoint(p4, null));
                rebarSet.LegFaces.Add(face);
            }

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing),
                SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = 100;
            guideline.Spacing.EndOffset = 100;

            Vector normal = Utility.GetVectorFromTwoPoints(ProfilePoints[0][0], ProfilePoints[1][0]).GetNormal();
            GeometricPlane backwallPlane = new GeometricPlane(ProfilePoints[0][0], normal);

            Line sLine = new Line(ProfilePoints[0][f], ProfilePoints[1][f]);
            Line eLine = new Line(ProfilePoints[0][s], ProfilePoints[1][s]);
            Point startGL = Utility.GetExtendedIntersection(sLine, backwallPlane, 2);
            Point endGL = Utility.GetExtendedIntersection(eLine, backwallPlane, 2);

            guideline.Curve.AddContourPoint(new ContourPoint(startGL, null));
            guideline.Curve.AddContourPoint(new ContourPoint(endGL, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            int[] faceLayer = new int[rebarSet.LegFaces.Count];
            for (int i = 0; i < faceLayer.Length; i++)
            {
                faceLayer[i] = 2;
            }
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, faceLayer);
        }
        void FirstPerpendicularRebar(int number)
        {
            string rebarSize = Program.ExcelDictionary["FPR_Diameter"];
            string spacing = Program.ExcelDictionary["FPR_Spacing"];
            double skewOffset = Convert.ToDouble(Program.ExcelDictionary["FPR_SkewOffset"]);

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "APS_LR_" + number;
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            Point sp0 = ProfilePoints[number][0];
            Point sp1 = ProfilePoints[number][1];
            Point sp2 = ProfilePoints[number][2];
            Point ep0 = ProfilePoints[number + 1][0];
            Point ep1 = ProfilePoints[number + 1][1];
            Point ep2 = ProfilePoints[number + 1][2];
            Vector xAxis = Utility.GetVectorFromTwoPoints(sp1, ep1).GetNormal();
            Vector yAxis = Utility.GetVectorFromTwoPoints(sp0, sp1).GetNormal();
            GeometricPlane endPlane = new GeometricPlane(ProfilePoints[number][5], xAxis,yAxis);
            Line startLine12 = new Line(ProfilePoints[number][1], ProfilePoints[number][2]);
            Line endLine12 = new Line(ProfilePoints[number + 1][1], ProfilePoints[number + 1][2]);
            Point startCorrecterP4 = Utility.GetExtendedIntersection(startLine12, endPlane, 2);
            Point endCorrectedP4 = Utility.GetExtendedIntersection(endLine12, endPlane, 2);
           
            Vector perpVector = Utility.GetVectorFromTwoPoints(sp1, sp2).GetNormal();
            Point startLeftTopSkewPoint = Utility.TranslePointByVectorAndDistance(sp1, perpVector, skewOffset);
            Point endLeftTopSkewPoint = Utility.TranslePointByVectorAndDistance(ep1, perpVector, skewOffset);
            Point startRightTopSkewPoint = Utility.TranslePointByVectorAndDistance(startCorrecterP4, perpVector, -skewOffset);
            Point endRightTopSkewPoint = Utility.TranslePointByVectorAndDistance(endCorrectedP4, perpVector, -skewOffset);
            double b = Thickness / Math.Tan(Math.PI / 4.0);
            Point startSkewPlaneOrigin = Utility.TranslePointByVectorAndDistance(startLeftTopSkewPoint, perpVector, b);
            Point endSkewPlaneOrigin = Utility.Translate(startRightTopSkewPoint, perpVector, -b);
            GeometricPlane startSkewPlane = new GeometricPlane(startSkewPlaneOrigin, xAxis, yAxis);
            GeometricPlane endSkewPlane = new GeometricPlane(endSkewPlaneOrigin, xAxis, yAxis);
            Line startLine05 = new Line(ProfilePoints[number][0], ProfilePoints[number][5]);
            Line endLine05 = new Line(ProfilePoints[number + 1][0], ProfilePoints[number + 1][5]);
            Point startLeftBottomSkewPoint = Utility.GetExtendedIntersection(startLine05, startSkewPlane, 2);
            Point endLeftBottomSkewPoint = Utility.GetExtendedIntersection(endLine05, startSkewPlane, 2);
            Point startRightBottomSkewPoint = Utility.GetExtendedIntersection(startLine05, endSkewPlane, 2);
            Point endRightBottomSkewPoint = Utility.GetExtendedIntersection(endLine05, endSkewPlane, 2);

            var firstFace = new RebarLegFace();
            firstFace.Contour.AddContourPoint(new ContourPoint(sp1, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(ep1, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(endLeftTopSkewPoint, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(startLeftTopSkewPoint, null));
            rebarSet.LegFaces.Add(firstFace);
            
            var secondFace = new RebarLegFace();
            secondFace.Contour.AddContourPoint(new ContourPoint(startLeftTopSkewPoint, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(endLeftTopSkewPoint, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(endLeftBottomSkewPoint, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(startLeftBottomSkewPoint, null));
            rebarSet.LegFaces.Add(secondFace);

            var thirdFace = new RebarLegFace();
            thirdFace.Contour.AddContourPoint(new ContourPoint(startLeftBottomSkewPoint, null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(endLeftBottomSkewPoint, null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(endRightBottomSkewPoint, null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(startRightBottomSkewPoint, null));
            rebarSet.LegFaces.Add(thirdFace);
            
            var fourthFace = new RebarLegFace();
            fourthFace.Contour.AddContourPoint(new ContourPoint(startRightBottomSkewPoint, null));
            fourthFace.Contour.AddContourPoint(new ContourPoint(endRightBottomSkewPoint, null));
            fourthFace.Contour.AddContourPoint(new ContourPoint(endRightTopSkewPoint, null));
            fourthFace.Contour.AddContourPoint(new ContourPoint(startRightTopSkewPoint, null));
            rebarSet.LegFaces.Add(fourthFace);

            var fifthFace = new RebarLegFace();
            fifthFace.Contour.AddContourPoint(new ContourPoint(startRightTopSkewPoint, null));
            fifthFace.Contour.AddContourPoint(new ContourPoint(endRightTopSkewPoint, null));
            fifthFace.Contour.AddContourPoint(new ContourPoint(endCorrectedP4, null));
            fifthFace.Contour.AddContourPoint(new ContourPoint(startCorrecterP4, null));
            rebarSet.LegFaces.Add(fifthFace);

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing),
                SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = 100;
            guideline.Spacing.EndOffset = 100;

            Point startGL = new Point(ProfilePoints[0][1]);
            Point endGL = Utility.TranslePointByVectorAndDistance(startGL, new Vector(1, 0, 0), Length);
            guideline.Curve.AddContourPoint(new ContourPoint(startGL, null));
            guideline.Curve.AddContourPoint(new ContourPoint(endGL, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            int[] faceLayer = new int[rebarSet.LegFaces.Count];
            for (int i = 0; i < faceLayer.Length; i++)
            {
                faceLayer[i] = 2;
            }
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, faceLayer);
        }
        #endregion
        #region Fields
        public static double Height;
        public static double Height2;
        public static double Width;
        public static double TopWidth;
        public static double Thickness;
        public static double SkewWidth;
        public static double Drop;
        public static double HorizontalOffset;
        public static double VerticalOffset;
        public static double Length;
        public static double FullLength;
        public static double FrontHeight;
        public static double Angle;
        #endregion
    }
}
