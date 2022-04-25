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
        public APS(params Part[] parts) : base()
        {
            List<Part> partList = new List<Part>(parts);
            SetLocalPlane();
            GetProfilePointsAndParameters(partList);
        }
        public void GetProfilePointsAndParameters(List<Part> parts)
        {
            //ABSV Height*Width*Height2*TopWidth*Thickness*SkewWidth*Drop*HorizontalOffset*VerticaOffset

            FirstBeamProperties(parts[0]);
            for (int i = 1; i < parts.Count; i++)
            {
                NextBeamProperties(parts[i]);
            }
        }
        public override void Create()
        {
            LongitudinalRebar(0);
            LongitudinalRebar(1);
            LongitudinalRebar(3);

            for (int i = 0; i < ProfilePoints.Count - 1; i++)
            {
                PerpendicularRebar(i, true);

                PerpendicularRebar(i, false);
                CantileverRebar(i);
                LongitudinalEndRebar(i);
                TopPerpendicularRebar(i);
                BottomPerpendicularRebar(i);
                ClosingBackRebar(i);
                CShapeRebarCore(i);

            }
            LongitudlSkewRebar();
            ClosingRebar(0, true);
            ClosingRebar(ProfilePoints.Count - 2, false);

        }
        public override void CreateSingle(string barName)
        {
            throw new NotImplementedException();
        }
        #region PrivateMethods
        void FirstBeamProperties(Part part)
        {
            Beam beam = part as Beam;
            string[] profileValues = GetProfileValues(beam);
            Height = Convert.ToDouble(profileValues[0]);
            Width = Convert.ToDouble(profileValues[1]);
            Height2 = Convert.ToDouble(profileValues[2]);
            TopWidth = Convert.ToDouble(profileValues[3]);
            SkewWidth = Convert.ToDouble(profileValues[4]);
            Drop = Convert.ToDouble(profileValues[5]);
            Thickness = Convert.ToDouble(profileValues[6]);
            HorizontalOffset.Add( Convert.ToDouble(profileValues[7]));
            VerticalOffset.Add( Convert.ToDouble(profileValues[8]));
            Length.Add(Distance.PointToPoint(beam.StartPoint, beam.EndPoint));
            FullLength += Length.Last();
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

            Point n0 = new Point(Length[0], -distanceToMid, -Width / 2.0);
            Point n1 = new Point(Length[0], n0.Y + FrontHeight, n0.Z);
            Point n2 = new Point(Length[0], n1.Y + (Width - TopWidth - SkewWidth) * Drop / 100, n1.Z + Width - TopWidth - SkewWidth);
            Point n3 = new Point(Length[0], n0.Y + bottomHeight + Height2, n1.Z + Width - TopWidth);
            Point n4 = new Point(Length[0], n0.Y + bottomHeight + Height, n0.Z + Width);
            Point n5 = new Point(Length[0], n0.Y + bottomHeight, n4.Z);
            List<Point> secondProfile = new List<Point> { n0, n1, n2, n3, n4, n5 };

            if (HorizontalOffset[0] != 0)
            {
                foreach (Point p in firstProfile)
                {
                    p.Translate(0, 0, -HorizontalOffset[0] / 2.0);
                }
                foreach (Point p in secondProfile)
                {
                    p.Translate(0, 0, HorizontalOffset[0] / 2.0);
                }
            }
            if (VerticalOffset[0] != 0)
            {
                foreach (Point p in firstProfile)
                {
                    p.Translate(0, -VerticalOffset[0] / 2.0, 0);
                }
                foreach (Point p in secondProfile)
                {
                    p.Translate(0, VerticalOffset[0] / 2.0, 0);
                }
            }

            ProfilePoints.Add(firstProfile);
            ProfilePoints.Add(secondProfile);
        }
        void NextBeamProperties(Part part)
        {
            Beam secondBeam = part as Beam;
            string[] nextProfileValues = GetProfileValues(secondBeam);
            Length.Add(Distance.PointToPoint(secondBeam.StartPoint, secondBeam.EndPoint));
            HorizontalOffset.Add(Convert.ToDouble(nextProfileValues[7]));
            VerticalOffset.Add( Convert.ToDouble(nextProfileValues[8]));
            FullLength += Length.Last();

            double bottomHeight = Width * Drop / 100.0;
            double distanceToMid = (Height2 + bottomHeight) / 2.0;

            Point p0 = new Point(FullLength, -distanceToMid, -Width / 2.0);
            Point p1 = new Point(FullLength, p0.Y + FrontHeight, p0.Z);
            Point p2 = new Point(FullLength, p1.Y + (Width - TopWidth - SkewWidth) * Drop / 100, p1.Z + Width - TopWidth - SkewWidth);
            Point p3 = new Point(FullLength, p0.Y + bottomHeight + Height2, p1.Z + Width - TopWidth);
            Point p4 = new Point(FullLength, p0.Y + bottomHeight + Height, p0.Z + Width);
            Point p5 = new Point(FullLength, p0.Y + bottomHeight, p4.Z);
            List<Point> thirdProfile = new List<Point> { p0, p1, p2, p3, p4, p5 };

            if (HorizontalOffset.Last() != 0)
            {
                double horizontalOffsetSum = 0;
                for(int i=0;i<HorizontalOffset.Count;i++)
                {
                    horizontalOffsetSum += i == 0 ? 0.5 * HorizontalOffset[i] : HorizontalOffset[i];
                }

                foreach (Point p in thirdProfile)
                {
                    p.Translate(0, 0,horizontalOffsetSum);
                }
            }
            if (VerticalOffset.Last() != 0)
            {
                double verticalOffsetSum = 0;
                for (int i = 0; i < VerticalOffset.Count; i++)
                {
                    verticalOffsetSum += i == 0 ? 0.5 * VerticalOffset[i] : VerticalOffset[i];
                }

                foreach (Point p in thirdProfile)
                {
                    p.Translate(0, verticalOffsetSum, 0);
                }
            }

            ProfilePoints.Add(thirdProfile);
        }    
        void LongitudinalEndRebar(int number)
        {
            string rebarSize = Program.ExcelDictionary["LR_Diameter"];
            string spacing = Program.ExcelDictionary["LR_Spacing"];

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "APS_LR_";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = TeklaUtils.SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = TeklaUtils.GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            int f, s;
            f = 4;
            s = 5;

            Point p1 = ProfilePoints[number][f];
            Point p2 = ProfilePoints[number + 1][f];
            Point p3 = ProfilePoints[number + 1][s];
            Point p4 = ProfilePoints[number][s];
            var face = new RebarLegFace();
            face.Contour.AddContourPoint(new ContourPoint(p1, null));
            face.Contour.AddContourPoint(new ContourPoint(p2, null));
            face.Contour.AddContourPoint(new ContourPoint(p3, null));
            face.Contour.AddContourPoint(new ContourPoint(p4, null));
            rebarSet.LegFaces.Add(face);

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing),
                SpacingType = RebarSpacingZone.SpacingEnum.TARGET,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = 100;
            guideline.Spacing.EndOffset = 100;

            Vector normal = Utility.GetVectorFromTwoPoints(ProfilePoints[number][0], ProfilePoints[number + 1][0]).GetNormal();
            GeometricPlane backwallPlane = new GeometricPlane(ProfilePoints[number][0], normal);

            Line sLine = new Line(ProfilePoints[number][f], ProfilePoints[number + 1][f]);
            Line eLine = new Line(ProfilePoints[number][s], ProfilePoints[number + 1][s]);
            Point startGL = Utility.GetExtendedIntersection(sLine, backwallPlane, 2);
            Point endGL = Utility.GetExtendedIntersection(eLine, backwallPlane, 2);

            guideline.Curve.AddContourPoint(new ContourPoint(startGL, null));
            guideline.Curve.AddContourPoint(new ContourPoint(endGL, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            if (ProfilePoints.Count>2&&number!=ProfilePoints.Count-2)
            {
                var endModifier = new RebarEndDetailModifier();
                endModifier.Father = rebarSet;
                endModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.END_OFFSET;
                endModifier.RebarLengthAdjustment.AdjustmentLength = 10 * Convert.ToInt32(rebarSize);
                endModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[number+1][4], null));
                endModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[number+1][5], null));
                endModifier.Insert();
            }
            if(number!=0)
            {
                var startModifier = new RebarEndDetailModifier();
                startModifier.Father = rebarSet;
                startModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.END_OFFSET;
                startModifier.RebarLengthAdjustment.AdjustmentLength = 10 * Convert.ToInt32(rebarSize);
                startModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[number][4], null));
                startModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[number][5], null));
                startModifier.Insert();
            }

            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FATHER_ID_NAME, RebarCreator.FatherID);
            int[] faceLayer = new int[rebarSet.LegFaces.Count];
            for (int i = 0; i < faceLayer.Length; i++)
            {
                faceLayer[i] = 2;
            }
           LayerDictionary.Add(rebarSet.Identifier.ID, faceLayer);
        }
        void LongitudlSkewRebar()
        {
            string rebarSize = Program.ExcelDictionary["LR_Diameter"];
            string spacing = Program.ExcelDictionary["LR_Spacing"];

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "APS_LR_";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = TeklaUtils.SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = TeklaUtils.GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            int f, s;
            f = 2;
            s = 3;

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

            Vector normal = new Vector(0, 1, 0);
            GeometricPlane plane = new GeometricPlane(ProfilePoints[0][0], normal);
            Point projectedSP2 = Projection.PointToPlane(ProfilePoints[0][2], plane);
            Point projectedSP3 = Projection.PointToPlane(ProfilePoints[0][3], plane);

            double normalDistance = Distance.PointToPoint(ProfilePoints[0][2], ProfilePoints[0][3]);
            double projectedDistance = Distance.PointToPoint(projectedSP2, projectedSP3);
            double spacingCoefficient = projectedDistance / normalDistance;

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing) * spacingCoefficient,
                SpacingType = RebarSpacingZone.SpacingEnum.TARGET,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = 100;
            guideline.Spacing.EndOffset = 100;

            guideline.Curve.AddContourPoint(new ContourPoint(projectedSP2, null));
            guideline.Curve.AddContourPoint(new ContourPoint(projectedSP3, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FATHER_ID_NAME, RebarCreator.FatherID);
            int[] faceLayer = new int[rebarSet.LegFaces.Count];
            for (int i = 0; i < faceLayer.Length; i++)
            {
                faceLayer[i] = 2;
            }
           LayerDictionary.Add(rebarSet.Identifier.ID, faceLayer);
        }
        void LongitudinalRebar(int faceNumber)
        {
            string rebarSize = Program.ExcelDictionary["LR_Diameter"];
            string spacing = Program.ExcelDictionary["LR_Spacing"];

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "APS_LR_" + faceNumber;
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = TeklaUtils.SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = TeklaUtils.GetBendingRadious(Convert.ToDouble(rebarSize));
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
                SpacingType = RebarSpacingZone.SpacingEnum.TARGET,
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
            rebarSet.SetUserProperty(RebarCreator.FATHER_ID_NAME, RebarCreator.FatherID);
            int[] faceLayer = new int[rebarSet.LegFaces.Count];
            for (int i = 0; i < faceLayer.Length; i++)
            {
                faceLayer[i] = 2;
            }
           LayerDictionary.Add(rebarSet.Identifier.ID, faceLayer);
        }
        void PerpendicularRebar(int number, bool isFirst)
        {
            string rebarSizeName = isFirst ? "FPR_Diameter" : "SPR_Diameter";
            string spacingName = isFirst ? "FPR_Spacing" : "SPR_Spacing";
            string offsetName = isFirst ? "FPR_SkewOffset" : "SPR_SkewOffset";
            string rebarSize = Program.ExcelDictionary[rebarSizeName];
            string spacing = Program.ExcelDictionary[spacingName];
            double skewOffset = Convert.ToDouble(Program.ExcelDictionary[offsetName]);
            double dSPacing = Convert.ToDouble(spacing);

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "APS_LR_" + number;
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = TeklaUtils.SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = TeklaUtils.GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            Point sp0 = ProfilePoints[number][0];
            Point sp1 = ProfilePoints[number][1];
            Point sp2 = ProfilePoints[number][2];
            Point ep0 = ProfilePoints[number + 1][0];
            Point ep1 = ProfilePoints[number + 1][1];
            Point ep2 = ProfilePoints[number + 1][2];
            Vector xAxis = Utility.GetVectorFromTwoPoints(sp1, ep1).GetNormal();
            Vector yAxis = Utility.GetVectorFromTwoPoints(sp0, sp1).GetNormal();
            GeometricPlane endPlane = new GeometricPlane(ProfilePoints[number][5], xAxis, yAxis);
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
                SpacingType = RebarSpacingZone.SpacingEnum.TARGET,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = isFirst ? dSPacing / 2.0 : dSPacing;
            guideline.Spacing.StartOffsetType = number == 0 ? RebarSpacing.OffsetEnum.MINIMUM : RebarSpacing.OffsetEnum.EXACT;
            guideline.Spacing.EndOffset = isFirst ? dSPacing / 2.0 : dSPacing;
            guideline.Spacing.EndOffsetType = number != 0 ? RebarSpacing.OffsetEnum.MINIMUM : RebarSpacing.OffsetEnum.EXACT;

            Point startGL = new Point(ProfilePoints[number][1]);
            Point endGL = Utility.TranslePointByVectorAndDistance(startGL, new Vector(1, 0, 0), Length[number]);
            guideline.Curve.AddContourPoint(new ContourPoint(startGL, null));
            guideline.Curve.AddContourPoint(new ContourPoint(endGL, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FATHER_ID_NAME, RebarCreator.FatherID);
            int[] faceLayer = new int[rebarSet.LegFaces.Count];
            for (int i = 0; i < faceLayer.Length; i++)
            {
                faceLayer[i] = 1;
            }
           LayerDictionary.Add(rebarSet.Identifier.ID, faceLayer);
        }
        void TopPerpendicularRebar(int number)
        {
            string rebarSize = Program.ExcelDictionary["TPR_Diameter"];
            string spacing = Program.ExcelDictionary["TPR_Spacing"];
            double dSPacing = Convert.ToDouble(spacing);

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "APS_TPR_" + number;
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = TeklaUtils.SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = TeklaUtils.GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            Point sp0 = ProfilePoints[number][0];
            Point sp1 = ProfilePoints[number][1];
            Point sp2 = ProfilePoints[number][2];
            Point sp4 = ProfilePoints[number][4];
            Point sp5 = ProfilePoints[number][5];
            Point ep0 = ProfilePoints[number + 1][0];
            Point ep1 = ProfilePoints[number + 1][1];
            Point ep2 = ProfilePoints[number + 1][2];
            Point ep5 = ProfilePoints[number + 1][5];

            Line startLine12 = new Line(sp1, sp2);
            Line endLine12 = new Line(ep1, ep2);

            Vector xAxis = Utility.GetVectorFromTwoPoints(sp5, ep5);
            Vector yAxis = Utility.GetVectorFromTwoPoints(sp5, sp4);
            GeometricPlane geometricPlane = new GeometricPlane(sp5, xAxis, yAxis);

            Point correctedSP2 = Utility.GetExtendedIntersection(startLine12, geometricPlane, 2);
            Point correctedEP2 = Utility.GetExtendedIntersection(endLine12, geometricPlane, 2);

            var firstFace = new RebarLegFace();
            firstFace.Contour.AddContourPoint(new ContourPoint(sp1, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(ep1, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(correctedEP2, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(correctedSP2, null));
            rebarSet.LegFaces.Add(firstFace);

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing),
                SpacingType = RebarSpacingZone.SpacingEnum.TARGET,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = dSPacing / 2.0;
            guideline.Spacing.EndOffset = dSPacing / 2.0;

            Point startGL = new Point(ProfilePoints[number][1]);
            Point endGL = Utility.TranslePointByVectorAndDistance(startGL, new Vector(1, 0, 0), Length[number]);
            guideline.Curve.AddContourPoint(new ContourPoint(startGL, null));
            guideline.Curve.AddContourPoint(new ContourPoint(endGL, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FATHER_ID_NAME, RebarCreator.FatherID);
            int[] faceLayer = new int[rebarSet.LegFaces.Count];
            for (int i = 0; i < faceLayer.Length; i++)
            {
                faceLayer[i] = 1;
            }
           LayerDictionary.Add(rebarSet.Identifier.ID, faceLayer);
        }
        void BottomPerpendicularRebar(int number)
        {
            string rebarSize = Program.ExcelDictionary["BPR_Diameter"];
            string spacing = Program.ExcelDictionary["BPR_Spacing"];
            double dSPacing = Convert.ToDouble(spacing);

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "APS_LR_" + number;
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = TeklaUtils.SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = TeklaUtils.GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            Point sp0 = ProfilePoints[number][0];
            Point sp5 = ProfilePoints[number][5];
            Point ep0 = ProfilePoints[number + 1][0];
            Point ep5 = ProfilePoints[number + 1][5];

            var firstFace = new RebarLegFace();
            firstFace.Contour.AddContourPoint(new ContourPoint(sp0, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(ep0, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(ep5, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(sp5, null));
            rebarSet.LegFaces.Add(firstFace);

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing),
                SpacingType = RebarSpacingZone.SpacingEnum.TARGET,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = dSPacing / 2.0;
            guideline.Spacing.EndOffset = dSPacing / 2.0;

            Point startGL = new Point(ProfilePoints[number][1]);
            Point endGL = Utility.TranslePointByVectorAndDistance(startGL, new Vector(1, 0, 0), Length[number]);
            guideline.Curve.AddContourPoint(new ContourPoint(startGL, null));
            guideline.Curve.AddContourPoint(new ContourPoint(endGL, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FATHER_ID_NAME, RebarCreator.FatherID);
            int[] faceLayer = new int[rebarSet.LegFaces.Count];
            for (int i = 0; i < faceLayer.Length; i++)
            {
                faceLayer[i] = 1;
            }
           LayerDictionary.Add(rebarSet.Identifier.ID, faceLayer);
        }
        void CantileverRebar(int number)
        {
            string rebarSize = Program.ExcelDictionary["CR_Diameter"];
            string spacing = Program.ExcelDictionary["CR_Spacing"];
            double diameter = Convert.ToDouble(rebarSize);

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "APS_CR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = TeklaUtils.SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = TeklaUtils.GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            Point sp0 = ProfilePoints[number][0];
            Point sp2 = ProfilePoints[number][2];
            Point sp3 = ProfilePoints[number][3];
            Point sp4 = ProfilePoints[number][4];
            Point sp5 = ProfilePoints[number][5];
            Point ep0 = ProfilePoints[number + 1][0];
            Point ep2 = ProfilePoints[number + 1][2];
            Point ep3 = ProfilePoints[number + 1][3];
            Point ep4 = ProfilePoints[number + 1][4];
            Point ep5 = ProfilePoints[number + 1][5];

            Line startLine23 = new Line(sp2, sp3);
            Line endLine23 = new Line(ep2, ep3);

            Vector xAxis = Utility.GetVectorFromTwoPoints(sp0, ep0);
            Vector yAxis = Utility.GetVectorFromTwoPoints(sp0, sp5);
            GeometricPlane bottomPlane = new GeometricPlane(sp0, xAxis, yAxis);

            Point startIntersection23 = Utility.GetExtendedIntersection(startLine23, bottomPlane, 10);
            Point endIntersection23 = Utility.GetExtendedIntersection(endLine23, bottomPlane, 10);

            Vector perpVector = Utility.GetVectorFromTwoPoints(sp5, sp0).GetNormal();

            Point correctedStartIntersection23 = Utility.TranslePointByVectorAndDistance(startIntersection23, perpVector, 20 * diameter);
            Point correctedEndIntersection23 = Utility.TranslePointByVectorAndDistance(endIntersection23, perpVector, 20 * diameter);
            Point correctedSP5 = Utility.TranslePointByVectorAndDistance(sp5, perpVector, 20 * diameter);
            Point correctedEP5 = Utility.TranslePointByVectorAndDistance(ep5, perpVector, 20 * diameter);

            var zeroFace = new RebarLegFace();
            zeroFace.Contour.AddContourPoint(new ContourPoint(correctedStartIntersection23, null));
            zeroFace.Contour.AddContourPoint(new ContourPoint(startIntersection23, null));
            zeroFace.Contour.AddContourPoint(new ContourPoint(endIntersection23, null));
            zeroFace.Contour.AddContourPoint(new ContourPoint(correctedEndIntersection23, null));
            rebarSet.LegFaces.Add(zeroFace);

            var firstFace = new RebarLegFace();
            firstFace.Contour.AddContourPoint(new ContourPoint(startIntersection23, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(sp3, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(ep3, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(endIntersection23, null));
            rebarSet.LegFaces.Add(firstFace);

            var secondFace = new RebarLegFace();
            secondFace.Contour.AddContourPoint(new ContourPoint(sp3, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(ep3, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(ep4, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(sp4, null));
            rebarSet.LegFaces.Add(secondFace);

            var thirdFace = new RebarLegFace();
            thirdFace.Contour.AddContourPoint(new ContourPoint(sp4, null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(ep4, null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(ep5, null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(sp5, null));
            rebarSet.LegFaces.Add(thirdFace);

            var fourthace = new RebarLegFace();
            fourthace.Contour.AddContourPoint(new ContourPoint(sp5, null));
            fourthace.Contour.AddContourPoint(new ContourPoint(ep5, null));
            fourthace.Contour.AddContourPoint(new ContourPoint(correctedEP5, null));
            fourthace.Contour.AddContourPoint(new ContourPoint(correctedSP5, null));
            rebarSet.LegFaces.Add(fourthace);

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

            Point startGL = new Point(ProfilePoints[number][3]);
            Point endGL = Utility.TranslePointByVectorAndDistance(startGL, new Vector(1, 0, 0), Length[number]);
            guideline.Curve.AddContourPoint(new ContourPoint(startGL, null));
            guideline.Curve.AddContourPoint(new ContourPoint(endGL, null));
            rebarSet.Guidelines.Add(guideline);

            bool succes = rebarSet.Insert();
            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FATHER_ID_NAME, RebarCreator.FatherID);
            int[] faceLayer = new int[rebarSet.LegFaces.Count];
            for (int i = 0; i < faceLayer.Length; i++)
            {
                faceLayer[i] = 1;
            }
           LayerDictionary.Add(rebarSet.Identifier.ID, faceLayer);
        }
        void CShapeRebarCore(int number)
        {
            string firstOffset = Program.ExcelDictionary["SPR_SkewOffset"];
            string rowSpacing = Program.ExcelDictionary["CSR_RowSpacing"];
            double dRowSpacing = Convert.ToDouble(rowSpacing);
            double dFirstOffset = Convert.ToDouble(firstOffset) + 500;
            double length = Distance.PointToPoint(ProfilePoints[0][0], ProfilePoints[0][5]);

            double reaminingLength = length - dFirstOffset * 2;
            int numberOfRows = (int)Math.Ceiling(reaminingLength / dRowSpacing);
            for (int i = 0; i < numberOfRows; i++)
            {
                double offset = dFirstOffset + (i) * dRowSpacing;
                CShapeRebar(number, offset);
            }

        }
        void CShapeRebar(int number, double offset)
        {
            string rebarSize = Program.ExcelDictionary["CSR_Diameter"];
            double dRebarSize = Convert.ToDouble(rebarSize);
            string spacing = Program.ExcelDictionary["CSR_Spacing"];
            double dSPacing = Convert.ToDouble(spacing);

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "APS_CSR_" + number;
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = TeklaUtils.SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = TeklaUtils.GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            Point sp0 = ProfilePoints[number][0];
            Point sp1 = ProfilePoints[number][1];
            Point sp2 = ProfilePoints[number][2];
            Point sp5 = ProfilePoints[number][5];
            Point ep0 = ProfilePoints[number + 1][0];
            Point ep1 = ProfilePoints[number + 1][1];
            Point ep2 = ProfilePoints[number + 1][2];
            Point ep5 = ProfilePoints[number + 1][5];

            Vector perpVector = Utility.GetVectorFromTwoPoints(ProfilePoints[number][0], ProfilePoints[number][5]).GetNormal();

            sp0 = Utility.TranslePointByVectorAndDistance(sp0, perpVector, offset);
            sp1 = Utility.TranslePointByVectorAndDistance(sp1, perpVector, offset);
            ep0 = Utility.TranslePointByVectorAndDistance(ep0, perpVector, offset);
            ep1 = Utility.TranslePointByVectorAndDistance(ep1, perpVector, offset);

            Point correctedSP0 = Utility.TranslePointByVectorAndDistance(sp0, perpVector, 20 * dRebarSize);
            Point correctedSP1 = Utility.TranslePointByVectorAndDistance(sp1, perpVector, 20 * dRebarSize);
            Point correctedEP0 = Utility.TranslePointByVectorAndDistance(ep0, perpVector, 20 * dRebarSize);
            Point correctedEP1 = Utility.TranslePointByVectorAndDistance(ep1, perpVector, 20 * dRebarSize);

            var firstFace = new RebarLegFace();
            firstFace.Contour.AddContourPoint(new ContourPoint(sp1, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(ep1, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(ep0, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(sp0, null));
            rebarSet.LegFaces.Add(firstFace);

            var secondFace = new RebarLegFace();
            secondFace.Contour.AddContourPoint(new ContourPoint(sp1, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(ep1, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(correctedEP1, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(correctedSP1, null));
            rebarSet.LegFaces.Add(secondFace);

            var thirdFace = new RebarLegFace();
            thirdFace.Contour.AddContourPoint(new ContourPoint(sp0, null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(ep0, null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(correctedEP0, null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(correctedSP0, null));
            rebarSet.LegFaces.Add(thirdFace);

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing),
                SpacingType = RebarSpacingZone.SpacingEnum.TARGET,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = dSPacing;
            guideline.Spacing.EndOffset = dSPacing;

            Point startGL = new Point(sp1);
            Point endGL = Utility.TranslePointByVectorAndDistance(startGL, new Vector(1, 0, 0), Length[number]);
            guideline.Curve.AddContourPoint(new ContourPoint(startGL, null));
            guideline.Curve.AddContourPoint(new ContourPoint(endGL, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FATHER_ID_NAME, RebarCreator.FatherID);
            int[] faceLayer = new int[rebarSet.LegFaces.Count];
            for (int i = 0; i < faceLayer.Length; i++)
            {
                faceLayer[i] = 1;
            }
           LayerDictionary.Add(rebarSet.Identifier.ID, faceLayer);
        }
        void ClosingBackRebar(int number)
        {
            string rebarSize = Program.ExcelDictionary["ClR_Diameter"];
            double dRebarSize = Convert.ToDouble(rebarSize);
            string spacing = Program.ExcelDictionary["ClR_Spacing"];
            double dSPacing = Convert.ToDouble(spacing);

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "APS_ClR_" + number;
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = TeklaUtils.SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = TeklaUtils.GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            Point sp0 = ProfilePoints[number][0];
            Point sp1 = ProfilePoints[number][1];
            Point sp2 = ProfilePoints[number][2];
            Point sp5 = ProfilePoints[number][5];
            Point ep0 = ProfilePoints[number + 1][0];
            Point ep1 = ProfilePoints[number + 1][1];
            Point ep2 = ProfilePoints[number + 1][2];
            Point ep5 = ProfilePoints[number + 1][5];

            Vector perpVector = Utility.GetVectorFromTwoPoints(ProfilePoints[number][0], ProfilePoints[number][5]).GetNormal();

            Point correctedSP0 = Utility.TranslePointByVectorAndDistance(sp0, perpVector, 20 * dRebarSize);
            Point correctedSP1 = Utility.TranslePointByVectorAndDistance(sp1, perpVector, 20 * dRebarSize);
            Point corretedEP0 = Utility.TranslePointByVectorAndDistance(ep0, perpVector, 20 * dRebarSize);
            Point correctedEP1 = Utility.TranslePointByVectorAndDistance(ep1, perpVector, 20 * dRebarSize);

            var firstFace = new RebarLegFace();
            firstFace.Contour.AddContourPoint(new ContourPoint(sp1, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(ep1, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(ep0, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(sp0, null));
            rebarSet.LegFaces.Add(firstFace);

            var secondFace = new RebarLegFace();
            secondFace.Contour.AddContourPoint(new ContourPoint(sp1, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(ep1, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(correctedEP1, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(correctedSP1, null));
            rebarSet.LegFaces.Add(secondFace);

            var thirdFace = new RebarLegFace();
            thirdFace.Contour.AddContourPoint(new ContourPoint(sp0, null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(ep0, null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(corretedEP0, null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(correctedSP0, null));
            rebarSet.LegFaces.Add(thirdFace);

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing),
                SpacingType = RebarSpacingZone.SpacingEnum.TARGET,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = dSPacing / 2.0;
            guideline.Spacing.EndOffset = dSPacing / 2.0;

            Point startGL = new Point(ProfilePoints[number][1]);
            Point endGL = Utility.TranslePointByVectorAndDistance(startGL, new Vector(1, 0, 0), Length[number]);
            guideline.Curve.AddContourPoint(new ContourPoint(startGL, null));
            guideline.Curve.AddContourPoint(new ContourPoint(endGL, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FATHER_ID_NAME, RebarCreator.FatherID);
            int[] faceLayer = new int[rebarSet.LegFaces.Count];
            for (int i = 0; i < faceLayer.Length; i++)
            {
                faceLayer[i] = 1;
            }
           LayerDictionary.Add(rebarSet.Identifier.ID, faceLayer);
        }
        void ClosingRebar(int number, bool isStart)
        {
            string rebarSize = Program.ExcelDictionary["ClR_Diameter"];
            double dRebarSize = Convert.ToDouble(rebarSize);
            string spacing = Program.ExcelDictionary["ClR_Spacing"];
            double dSPacing = Convert.ToDouble(spacing);

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "APS_ClR_" + number;
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = TeklaUtils.SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = TeklaUtils.GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            int firstNumber = isStart ? number : number + 1;
            int secondNumber = isStart ? number + 1 : number;

            Point sp0 = ProfilePoints[firstNumber][0];
            Point sp1 = ProfilePoints[firstNumber][1];
            Point sp2 = ProfilePoints[firstNumber][2];
            Point sp3 = ProfilePoints[firstNumber][3];
            Point sp4 = ProfilePoints[firstNumber][4];
            Point sp5 = ProfilePoints[firstNumber][5];

            Vector longVector = Utility.GetVectorFromTwoPoints(ProfilePoints[firstNumber][0], ProfilePoints[secondNumber][0]).GetNormal();
            Point ep0 = Utility.TranslePointByVectorAndDistance(sp0, longVector, 20 * dRebarSize);
            Point ep1 = Utility.TranslePointByVectorAndDistance(sp1, longVector, 20 * dRebarSize);
            Point ep2 = Utility.TranslePointByVectorAndDistance(sp2, longVector, 20 * dRebarSize);
            Point ep3 = Utility.TranslePointByVectorAndDistance(sp3, longVector, 20 * dRebarSize);
            Point ep4 = Utility.TranslePointByVectorAndDistance(sp4, longVector, 20 * dRebarSize);
            Point ep5 = Utility.TranslePointByVectorAndDistance(sp5, longVector, 20 * dRebarSize);

            var firstFace = new RebarLegFace();
            firstFace.Contour.AddContourPoint(new ContourPoint(sp0, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(sp5, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(sp4, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(sp3, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(sp2, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(sp1, null));
            rebarSet.LegFaces.Add(firstFace);

            var secondFace = new RebarLegFace();
            secondFace.Contour.AddContourPoint(new ContourPoint(sp0, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(sp5, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(ep5, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(ep0, null));
            rebarSet.LegFaces.Add(secondFace);

            var thirdFace = new RebarLegFace();
            thirdFace.Contour.AddContourPoint(new ContourPoint(sp4, null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(ep4, null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(ep3, null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(sp3, null));
            rebarSet.LegFaces.Add(thirdFace);

            var fourthFace = new RebarLegFace();
            fourthFace.Contour.AddContourPoint(new ContourPoint(sp3, null));
            fourthFace.Contour.AddContourPoint(new ContourPoint(sp2, null));
            fourthFace.Contour.AddContourPoint(new ContourPoint(ep2, null));
            fourthFace.Contour.AddContourPoint(new ContourPoint(ep3, null));
            rebarSet.LegFaces.Add(fourthFace);

            var fifthFace = new RebarLegFace();
            fifthFace.Contour.AddContourPoint(new ContourPoint(sp1, null));
            fifthFace.Contour.AddContourPoint(new ContourPoint(sp2, null));
            fifthFace.Contour.AddContourPoint(new ContourPoint(ep2, null));
            fifthFace.Contour.AddContourPoint(new ContourPoint(ep1, null));
            rebarSet.LegFaces.Add(fifthFace);

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing),
                SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = dSPacing / 2.0;
            guideline.Spacing.EndOffset = dSPacing / 2.0;

            Point startGL = new Point(ProfilePoints[number][1]);
            Point endGL = Utility.TranslePointByVectorAndDistance(startGL, new Vector(0, 0, 1), Width);
            guideline.Curve.AddContourPoint(new ContourPoint(startGL, null));
            guideline.Curve.AddContourPoint(new ContourPoint(endGL, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FATHER_ID_NAME, RebarCreator.FatherID);
            int[] faceLayer = new int[rebarSet.LegFaces.Count];
            for (int i = 0; i < faceLayer.Length; i++)
            {
                faceLayer[i] = 2;
            }
           LayerDictionary.Add(rebarSet.Identifier.ID, faceLayer);
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
        public static List<double> HorizontalOffset = new List<double>();
        public static List<double> VerticalOffset = new List<double>();
        //   public static double HorizontalOffset;
        //  public static double VerticalOffset;
        public static List<double> Length = new List<double>();
       // public static double VerticalOffset2;
      //  public static double Length2;
        public static double FullLength;
        public static double FrontHeight;
        public static double Angle;
        #endregion
    }
}
