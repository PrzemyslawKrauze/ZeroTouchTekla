﻿using System;
using System.Collections.Generic;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model;
using System.Linq;

namespace ZeroTouchTekla.Profiles
{
    public class DABT : Element
    {
        #region Constructor
        public DABT(Beam part) : base(part)
        {
        }
        public DABT(Beam part, Beam secondPart) : base(part)
        {
            GetProfilePointsAndParameters(part, secondPart);
        }
        new public void Create()
        {
            OuterVerticalRebar();
            InnerVerticalRebar();
            CantileverVerticalRebar(0);
            CantileverVerticalRebar(1);
            BackwallTopVerticalRebar(0);
            BackwallTopVerticalRebar(1);
            BackwallOuterVerticalRebar(0);
            BackwallOuterVerticalRebar(1);
            BackwallInnerVerticalRebar(0);
            BackwallInnerVerticalRebar(1);
            ShelfHorizontalRebar(0);
            ShelfHorizontalRebar(1);
            OuterLongitudinalRebar();
            InnerLongitudinalRebar();
            CantileverLongitudinalRebar(0);
            CantileverLongitudinalRebar(1);
            CantileverLongitudinalRebar(2);
            BackwallLongitudinalRebar(0);
            BackwallLongitudinalRebar(1);
            BackwallLongitudinalRebar(2);
            ShelfLongitudinalRebar();
            ClosingCShapeRebarBottom(0);
            ClosingCShapeRebarBottom(2);
            ClosingCShapeRebarMid(0);
            ClosingCShapeRebarMid(2);
            ClosingCShapeRebarTop(0);
            ClosingCShapeRebarTop(2);
            CShapeRebarCommon();
            ClosingLongitudianlRebarBottom(0);
            ClosingLongitudianlRebarBottom(2);
            ClosingLongitudianlRebarMid(0);
            ClosingLongitudianlRebarMid(2);
            ClosingLongitudianlRebarTop(0);
            ClosingLongitudianlRebarTop(2);
            ClosingLongitudianlRebarTop2(0);
            ClosingLongitudianlRebarTop2(2);
        }
        #endregion
        #region PrivateMethods  
        public static void GetProfilePointsAndParameters(Beam beam, Beam secondBeam)
        {
            string[] profileValues = GetProfileValues(beam);
            string[] secondProfileValues = GetProfileValues(secondBeam);
            //ABT Width*Height*FrontHeight*ShelfHeight*ShelfWidth*BackwallWidth*CantileverWidth*BackwallTopHeight*CantileverHeight*BackwallBottomHeight*SkewHeight
            //ABTV W*H*FH*SH*SW*BWW*CW*BWTH*CH*BWBH*SH*H*FH*BWTH*BWBH

            Width = Convert.ToDouble(profileValues[0]);
            Height = Convert.ToDouble(profileValues[1]);
            FrontHeight = Convert.ToDouble(profileValues[2]);
            ShelfHeight = Convert.ToDouble(profileValues[3]);
            ShelfWidth = Convert.ToDouble(profileValues[4]);
            BackwallWidth = Convert.ToDouble(profileValues[5]);
            CantileverWidth = Convert.ToDouble(profileValues[6]);
            BackwallTopHeight = Convert.ToDouble(profileValues[7]);
            CantileverHeight = Convert.ToDouble(profileValues[8]);
            BackwallBottomHeight = Convert.ToDouble(profileValues[9]);
            SkewHeight = Convert.ToDouble(profileValues[10]);
            FullWidth = ShelfWidth + BackwallWidth + CantileverWidth;
            Length = Distance.PointToPoint(beam.StartPoint, beam.EndPoint);
            Length2 = Distance.PointToPoint(secondBeam.StartPoint, secondBeam.EndPoint);
            double fullLength = Length + Length2;

            if (profileValues.Length > 11)
            {
                Height2 = Convert.ToDouble(profileValues[11]);
                FrontHeight2 = Convert.ToDouble(profileValues[12]);
                BackwallTopHeight2 = Convert.ToDouble(profileValues[13]);
                BackwallBottomHeight2 = Convert.ToDouble(profileValues[14]);
            }
            else
            {
                Height2 = Height;
                FrontHeight2 = FrontHeight;
                BackwallTopHeight2 = BackwallBottomHeight;
                BackwallBottomHeight2 = BackwallBottomHeight;
            }

            if (secondProfileValues.Length > 11)
            {
                Height3 = Convert.ToDouble(secondProfileValues[11]);
                FrontHeight3 = Convert.ToDouble(secondProfileValues[12]);
                BackwallTopHeight3 = Convert.ToDouble(secondProfileValues[13]);
                BackwallBottomHeight3 = Convert.ToDouble(secondProfileValues[14]);
            }
            else
            {
                Height3 = Height2;
                FrontHeight3 = FrontHeight2;
                BackwallTopHeight3 = BackwallBottomHeight2;
                BackwallBottomHeight3 = BackwallBottomHeight2;
            }


            double distanceToMid = Height > Height2 ? Height / 2.0 : Height2 / 2.0;

            Point p0 = new Point(0, -distanceToMid, FullWidth / 2.0);
            Point p1 = new Point(0, p0.Y + FrontHeight, p0.Z);
            Point p2 = new Point(0, p1.Y + ShelfHeight, p1.Z - ShelfWidth);
            Point p3 = new Point(0, -distanceToMid + Height, p2.Z);
            Point p4 = new Point(0, -distanceToMid + Height, p3.Z - BackwallWidth);
            Point p5 = new Point(0, p4.Y - BackwallTopHeight, p4.Z);
            Point p6 = new Point(0, p5.Y - CantileverHeight, p5.Z - CantileverWidth);
            Point p7 = new Point(0, p6.Y - BackwallBottomHeight, p6.Z);
            Point p8 = new Point(0, p7.Y - SkewHeight, FullWidth / 2.0 - Width);
            Point p9 = new Point(0, -distanceToMid, FullWidth / 2.0 - Width);

            List<Point> firstProfile = new List<Point> { p0, p1, p2, p3, p4, p5, p6, p7, p8, p9 };

            List<Point> secondProfile = new List<Point>();
            if (Height2 == Height)
            {
                foreach (Point p in firstProfile)
                {
                    Point secondPoint = new Point(p.X, p.Y, p.Z);
                    secondPoint.Translate(Length, 0, 0);
                    secondProfile.Add(secondPoint);
                }
            }
            else
            {
                Point n0 = new Point(Length, -distanceToMid, FullWidth / 2.0);
                Point n1 = new Point(Length, n0.Y + FrontHeight2, n0.Z);
                Point n2 = new Point(Length, n1.Y + ShelfHeight, n1.Z - ShelfWidth);
                Point n3 = new Point(Length, -distanceToMid + Height2, n2.Z);
                Point n4 = new Point(Length, n3.Y, n3.Z - BackwallWidth);
                Point n5 = new Point(Length, n4.Y - BackwallTopHeight2, n4.Z);
                Point n6 = new Point(Length, n5.Y - CantileverHeight, n5.Z - CantileverWidth);
                Point n7 = new Point(Length, n6.Y - BackwallBottomHeight2, n6.Z);
                Point n8 = new Point(Length, n7.Y - SkewHeight, FullWidth / 2.0 - Width);
                Point n9 = new Point(Length, -distanceToMid, FullWidth / 2.0 - Width);
                secondProfile.AddRange(new List<Point> { n0, n1, n2, n3, n4, n5, n6, n7, n8, n9 });
            }
            List<Point> thirdProfile = new List<Point>();
            if (Height3 == Height2)
            {
                foreach (Point p in secondProfile)
                {
                    Point secondPoint = new Point(p.X, p.Y, p.Z);
                    secondPoint.Translate(Length2, 0, 0);
                    thirdProfile.Add(secondPoint);
                }
            }
            else
            {
                Point n0 = new Point(fullLength, -distanceToMid, FullWidth / 2.0);
                Point n1 = new Point(fullLength, n0.Y + FrontHeight3, n0.Z);
                Point n2 = new Point(fullLength, n1.Y + ShelfHeight, n1.Z - ShelfWidth);
                Point n3 = new Point(fullLength, -distanceToMid + Height3, n2.Z);
                Point n4 = new Point(fullLength, n3.Y, n3.Z - BackwallWidth);
                Point n5 = new Point(fullLength, n4.Y - BackwallTopHeight3, n4.Z);
                Point n6 = new Point(fullLength, n5.Y - CantileverHeight, n5.Z - CantileverWidth);
                Point n7 = new Point(fullLength, n6.Y - BackwallBottomHeight3, n6.Z);
                Point n8 = new Point(fullLength, n7.Y - SkewHeight, FullWidth / 2.0 - Width);
                Point n9 = new Point(fullLength, -distanceToMid, FullWidth / 2.0 - Width);
                thirdProfile.AddRange(new List<Point> { n0, n1, n2, n3, n4, n5, n6, n7, n8, n9 });
            }

            List<List<Point>> beamPoints = new List<List<Point>> { firstProfile, secondProfile, thirdProfile };
            ProfilePoints.AddRange(beamPoints);

        }
        #endregion
        #region PrivateMethods   
        void OuterVerticalRebar()
        {
            string rebarSize = Program.ExcelDictionary["OVR_Diameter"];
            int rebarDiameter = Convert.ToInt32(rebarSize);
            string spacing = Program.ExcelDictionary["OVR_Spacing"];
            int addSplitter = Convert.ToInt32(Program.ExcelDictionary["OVR_AddSplitter"]);
            string secondRebarSize = Program.ExcelDictionary["OVR_SecondDiameter"];
            double spliterOffset = Convert.ToDouble(Program.ExcelDictionary["OVR_SplitterOffset"]) + Convert.ToDouble(rebarSize) * 20;

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "ABT_OVR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][1], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
            rebarSet.LegFaces.Add(mainFace);

            Point offsetedStartPoint = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y, ProfilePoints[0][0].Z + 40 * Convert.ToInt32(rebarSize));
            Point offsetedEndPoint = new Point(ProfilePoints[2][0].X, ProfilePoints[2][0].Y, ProfilePoints[2][0].Z + 40 * Convert.ToInt32(rebarSize));

            var bottomFace = new RebarLegFace();
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][0], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(offsetedEndPoint, null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(offsetedStartPoint, null));
            rebarSet.LegFaces.Add(bottomFace);

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

            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[2][0], null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();

            var innerEndDetailModifier = new RebarEndDetailModifier();
            innerEndDetailModifier.Father = rebarSet;
            innerEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            innerEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = GetHookLength(rebarDiameter);
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(offsetedStartPoint, null));
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(offsetedEndPoint, null));
            innerEndDetailModifier.Insert();
            new Model().CommitChanges();

            if (addSplitter == 1)
            {
                var bottomSpliter = new RebarSplitter();
                bottomSpliter.Father = rebarSet;
                bottomSpliter.Lapping.LappingType = RebarLapping.LappingTypeEnum.STANDARD_LAPPING;
                bottomSpliter.Lapping.LapSide = RebarLapping.LapSideEnum.LAP_MIDDLE;
                bottomSpliter.Lapping.LapPlacement = RebarLapping.LapPlacementEnum.ON_LEG_FACE;
                bottomSpliter.BarsAffected = BaseRebarModifier.BarsAffectedEnum.EVERY_SECOND_BAR;
                bottomSpliter.FirstAffectedBar = 1;

                Point startIntersection = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y + spliterOffset, ProfilePoints[0][0].Z);
                Point endIntersection = new Point(ProfilePoints[2][0].X, ProfilePoints[2][0].Y + spliterOffset, ProfilePoints[2][0].Z);
                bottomSpliter.Curve.AddContourPoint(new ContourPoint(startIntersection, null));
                bottomSpliter.Curve.AddContourPoint(new ContourPoint(endIntersection, null));
                bottomSpliter.Insert();

                var topSpliter = new RebarSplitter();
                topSpliter.Father = rebarSet;
                topSpliter.Lapping.LappingType = RebarLapping.LappingTypeEnum.STANDARD_LAPPING;
                topSpliter.Lapping.LapSide = RebarLapping.LapSideEnum.LAP_MIDDLE;
                topSpliter.Lapping.LapPlacement = RebarLapping.LapPlacementEnum.ON_LEG_FACE;
                topSpliter.BarsAffected = BaseRebarModifier.BarsAffectedEnum.EVERY_SECOND_BAR;
                topSpliter.FirstAffectedBar = 2;

                Point startTopIntersection = new Point(startIntersection.X, startIntersection.Y + 1.3 * 40 * rebarDiameter, startIntersection.Z);
                Point endTopIntersection = new Point(endIntersection.X, endIntersection.Y + 1.3 * 40 * rebarDiameter, endIntersection.Z);

                topSpliter.Curve.AddContourPoint(new ContourPoint(startTopIntersection, null));
                topSpliter.Curve.AddContourPoint(new ContourPoint(endTopIntersection, null));
                topSpliter.Insert();

                if (rebarSize != secondRebarSize)
                {
                    var propertyModifier = new RebarPropertyModifier();
                    propertyModifier.Father = rebarSet;
                    propertyModifier.BarsAffected = BaseRebarModifier.BarsAffectedEnum.ALL_BARS;
                    propertyModifier.RebarProperties.Size = secondRebarSize;
                    propertyModifier.RebarProperties.Class = SetClass(Convert.ToDouble(secondRebarSize));
                    propertyModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
                    propertyModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[2][1], null));
                    propertyModifier.Insert();
                }
                new Model().CommitChanges();
            }

            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 3 });
        }
        void InnerVerticalRebar()
        {
            string rebarSize = Program.ExcelDictionary["IVR_Diameter"];
            int rebarDiameter = Convert.ToInt32(rebarSize);
            string spacing = Program.ExcelDictionary["IVR_Spacing"];
            int addSplitter = Convert.ToInt32(Program.ExcelDictionary["IVR_AddSplitter"]);
            string secondRebarSize = Program.ExcelDictionary["IVR_SecondDiameter"];
            double spliterOffset = Convert.ToDouble(Program.ExcelDictionary["IVR_SplitterOffset"]) + Convert.ToDouble(rebarSize) * 20;

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "ABT_IVR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            Point p1 = ProfilePoints[0][9];
            Point p2 = ProfilePoints[2][9];
            Point p3 = ProfilePoints[2][8];
            Point p4 = ProfilePoints[1][8];
            Point p5 = ProfilePoints[0][8];

            Point p3o = new Point(p3.X, p3.Y + 40 * rebarDiameter, p3.Z);
            Point p4o = new Point(p4.X, p4.Y + 40 * rebarDiameter, p4.Z);
            Point p5o = new Point(p5.X, p5.Y + 40 * rebarDiameter, p5.Z);

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(p1, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p2, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p3o, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p4o, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p5o, null));
            rebarSet.LegFaces.Add(mainFace);

            Point p1o = new Point(p1.X, p1.Y, p1.Z - 40 * Convert.ToInt32(rebarSize));
            Point p2o = new Point(p2.X, p2.Y, p2.Z - 40 * Convert.ToInt32(rebarSize));

            var bottomFace = new RebarLegFace();
            bottomFace.Contour.AddContourPoint(new ContourPoint(p1, null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(p2, null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(p2o, null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(p1o, null));
            rebarSet.LegFaces.Add(bottomFace);

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

            guideline.Curve.AddContourPoint(new ContourPoint(p1, null));
            guideline.Curve.AddContourPoint(new ContourPoint(p2, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();

            var innerEndDetailModifier = new RebarEndDetailModifier();
            innerEndDetailModifier.Father = rebarSet;
            innerEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            innerEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = GetHookLength(rebarDiameter);
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(p1o, null));
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(p2o, null));
            innerEndDetailModifier.Insert();
            new Model().CommitChanges();

            if (addSplitter == 1)
            {
                var bottomSpliter = new RebarSplitter();
                bottomSpliter.Father = rebarSet;
                bottomSpliter.Lapping.LappingType = RebarLapping.LappingTypeEnum.STANDARD_LAPPING;
                bottomSpliter.Lapping.LapSide = RebarLapping.LapSideEnum.LAP_MIDDLE;
                bottomSpliter.Lapping.LapPlacement = RebarLapping.LapPlacementEnum.ON_LEG_FACE;
                bottomSpliter.BarsAffected = BaseRebarModifier.BarsAffectedEnum.EVERY_SECOND_BAR;
                bottomSpliter.FirstAffectedBar = 1;

                Point startIntersection = new Point(p1.X, p1.Y + spliterOffset, p1.Z);
                Point endIntersection = new Point(p2.X, p2.Y + spliterOffset, p2.Z);
                bottomSpliter.Curve.AddContourPoint(new ContourPoint(startIntersection, null));
                bottomSpliter.Curve.AddContourPoint(new ContourPoint(endIntersection, null));
                bottomSpliter.Insert();

                var topSpliter = new RebarSplitter();
                topSpliter.Father = rebarSet;
                topSpliter.Lapping.LappingType = RebarLapping.LappingTypeEnum.STANDARD_LAPPING;
                topSpliter.Lapping.LapSide = RebarLapping.LapSideEnum.LAP_MIDDLE;
                topSpliter.Lapping.LapPlacement = RebarLapping.LapPlacementEnum.ON_LEG_FACE;
                topSpliter.BarsAffected = BaseRebarModifier.BarsAffectedEnum.EVERY_SECOND_BAR;
                topSpliter.FirstAffectedBar = 2;

                Point startTopIntersection = new Point(startIntersection.X, startIntersection.Y + 1.3 * 40 * rebarDiameter, startIntersection.Z);
                Point endTopIntersection = new Point(endIntersection.X, endIntersection.Y + 1.3 * 40 * rebarDiameter, endIntersection.Z);

                topSpliter.Curve.AddContourPoint(new ContourPoint(startTopIntersection, null));
                topSpliter.Curve.AddContourPoint(new ContourPoint(endTopIntersection, null));
                topSpliter.Insert();

                if (rebarSize != secondRebarSize)
                {
                    var propertyModifier = new RebarPropertyModifier();
                    propertyModifier.Father = rebarSet;
                    propertyModifier.BarsAffected = BaseRebarModifier.BarsAffectedEnum.ALL_BARS;
                    propertyModifier.RebarProperties.Size = secondRebarSize;
                    propertyModifier.RebarProperties.Class = SetClass(Convert.ToDouble(secondRebarSize));
                    propertyModifier.Curve.AddContourPoint(new ContourPoint(p3o, null));
                    propertyModifier.Curve.AddContourPoint(new ContourPoint(p4o, null));
                    propertyModifier.Curve.AddContourPoint(new ContourPoint(p5o, null));
                    propertyModifier.Insert();
                }
                new Model().CommitChanges();
            }

            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 3 });
        }
        void CantileverVerticalRebar(int number)
        {
            string rebarSize = Program.ExcelDictionary["CtVR_Diameter"];
            int rebarDiameter = Convert.ToInt32(rebarSize);
            string spacing = Program.ExcelDictionary["CtVR_Spacing"];

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "ABT_IVR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            int f = number == 0 ? 0 : 1;
            int s = number == 0 ? 1 : 2;

            Point p8s = ProfilePoints[f][8];
            Point p8e = ProfilePoints[s][8];
            Point p7s = ProfilePoints[f][7];
            Point p7e = ProfilePoints[s][7];
            Point p6s = ProfilePoints[f][6];
            Point p6e = ProfilePoints[s][6];
            Point p5s = ProfilePoints[f][5];
            Point p5e = ProfilePoints[s][5];
            Point p2s = ProfilePoints[f][2];
            Point p2e = ProfilePoints[s][2];

            Vector xAxis = Utility.GetVectorFromTwoPoints(ProfilePoints[f][2], ProfilePoints[s][2]).GetNormal();
            Vector yAxis = Utility.GetVectorFromTwoPoints(ProfilePoints[f][2], ProfilePoints[f][3]).GetNormal();
            GeometricPlane backwallPlane = new GeometricPlane(ProfilePoints[f][2], xAxis, yAxis);

            Line sLine = new Line(p6s, p5s);
            Line eLine = new Line(p6e, p5e);
            Point p3s = Utility.GetExtendedIntersection(sLine, backwallPlane, 5);
            Point p3e = Utility.GetExtendedIntersection(eLine, backwallPlane, 5);

            var firstFace = new RebarLegFace();
            firstFace.Contour.AddContourPoint(new ContourPoint(p8s, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(p8e, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(p7e, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(p7s, null));
            rebarSet.LegFaces.Add(firstFace);

            var secondFace = new RebarLegFace();
            secondFace.Contour.AddContourPoint(new ContourPoint(p7s, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(p7e, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(p6e, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(p6s, null));
            rebarSet.LegFaces.Add(secondFace);

            var thirdFace = new RebarLegFace();
            thirdFace.Contour.AddContourPoint(new ContourPoint(p6s, null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(p6e, null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(p3e, null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(p3s, null));
            rebarSet.LegFaces.Add(thirdFace);

            var fourthFace = new RebarLegFace();
            fourthFace.Contour.AddContourPoint(new ContourPoint(p3s, null));
            fourthFace.Contour.AddContourPoint(new ContourPoint(p3e, null));
            fourthFace.Contour.AddContourPoint(new ContourPoint(p2e, null));
            fourthFace.Contour.AddContourPoint(new ContourPoint(p2s, null));
            rebarSet.LegFaces.Add(fourthFace);

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing),
                SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = number == 0 ? 100 : 50;
            guideline.Spacing.StartOffsetType = number == 0 ? RebarSpacing.OffsetEnum.MINIMUM : RebarSpacing.OffsetEnum.EXACT;
            guideline.Spacing.EndOffset = number == 0 ? 50 : 100;
            guideline.Spacing.EndOffsetType = number == 0 ? RebarSpacing.OffsetEnum.EXACT : RebarSpacing.OffsetEnum.MINIMUM;

            guideline.Curve.AddContourPoint(new ContourPoint(p7s, null));
            guideline.Curve.AddContourPoint(new ContourPoint(p7e, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            var bottomLengthModifier = new RebarEndDetailModifier();
            bottomLengthModifier.Father = rebarSet;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.END_OFFSET;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentLength = 40 * rebarDiameter;
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(p8s, null));
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(p8e, null));
            bottomLengthModifier.Insert();

            var topEndModifier = new RebarEndDetailModifier();
            topEndModifier.Father = rebarSet;
            topEndModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            topEndModifier.RebarLengthAdjustment.AdjustmentLength = 20 * rebarDiameter;
            topEndModifier.Curve.AddContourPoint(new ContourPoint(p2s, null));
            topEndModifier.Curve.AddContourPoint(new ContourPoint(p2e, null));
            topEndModifier.Insert();

            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 1, 1, 1 });
        }
        void BackwallTopVerticalRebar(int number)
        {
            string rebarSize = Program.ExcelDictionary["BVR_Diameter"];
            int rebarDiameter = Convert.ToInt32(rebarSize);
            string spacing = Program.ExcelDictionary["BVR_Spacing"];

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "ABT_BTVR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            int f = number == 0 ? 0 : 1;
            int s = number == 0 ? 1 : 2;

            Point p4s = ProfilePoints[f][4];
            Point p4e = ProfilePoints[s][4];
            Point p3s = ProfilePoints[f][3];
            Point p3e = ProfilePoints[s][3];
            Point p2s = ProfilePoints[f][2];
            Point p2e = ProfilePoints[s][2];
            Point p5s = new Point(p4s.X, p2s.Y, p4s.Z);
            Point p5e = new Point(p4e.X, p2e.Y, p4e.Z);

            var firstFace = new RebarLegFace();
            firstFace.Contour.AddContourPoint(new ContourPoint(p2s, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(p2e, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(p3e, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(p3s, null));
            rebarSet.LegFaces.Add(firstFace);

            var secondFace = new RebarLegFace();
            secondFace.Contour.AddContourPoint(new ContourPoint(p3s, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(p3e, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(p4e, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(p4s, null));
            rebarSet.LegFaces.Add(secondFace);

            var fourthFace = new RebarLegFace();
            fourthFace.Contour.AddContourPoint(new ContourPoint(p4s, null));
            fourthFace.Contour.AddContourPoint(new ContourPoint(p4e, null));
            fourthFace.Contour.AddContourPoint(new ContourPoint(p5e, null));
            fourthFace.Contour.AddContourPoint(new ContourPoint(p5s, null));
            rebarSet.LegFaces.Add(fourthFace);


            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing),
                SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = number == 0 ? 100 : 50;
            guideline.Spacing.StartOffsetType = number == 0 ? RebarSpacing.OffsetEnum.MINIMUM : RebarSpacing.OffsetEnum.EXACT;
            guideline.Spacing.EndOffset = number == 0 ? 50 : 100;
            guideline.Spacing.EndOffsetType = number == 0 ? RebarSpacing.OffsetEnum.EXACT : RebarSpacing.OffsetEnum.MINIMUM;

            Point correctedP3 = new Point(p3e.X, p3s.Y, p3s.Z);
            guideline.Curve.AddContourPoint(new ContourPoint(p3s, null));
            guideline.Curve.AddContourPoint(new ContourPoint(correctedP3, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 1, 1 });
        }
        void BackwallOuterVerticalRebar(int number)
        {
            string rebarSize = Program.ExcelDictionary["BVR_Diameter"];
            int rebarDiameter = Convert.ToInt32(rebarSize);
            string spacing = Program.ExcelDictionary["BVR_Spacing"];

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "ABT_BOVR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(rebarDiameter);
            rebarSet.LayerOrderNumber = 1;

            int f = number == 0 ? 0 : 1;
            int s = number == 0 ? 1 : 2;

            Point p8s = ProfilePoints[f][8];
            Point p8e = ProfilePoints[s][8];
            Point p7s = ProfilePoints[f][7];
            Point p7e = ProfilePoints[s][7];
            Point p5s = ProfilePoints[f][5];
            Point p5e = ProfilePoints[s][5];
            Point p4s = ProfilePoints[f][4];
            Point p4e = ProfilePoints[s][4];
            Point p3s = ProfilePoints[f][3];
            Point p3e = ProfilePoints[s][3];
            Point p2s = ProfilePoints[f][2];
            Point p2e = ProfilePoints[s][2];

            Point startTopPoint = new Point(p4s.X, p5s.Y, p4s.Z);
            Point endTopPoint = new Point(p4e.X, p5e.Y, p4e.Z);

            Vector xAxis = Utility.GetVectorFromTwoPoints(p7s, p8s).GetNormal();
            Vector yAxis = Utility.GetVectorFromTwoPoints(p7s, p7e).GetNormal();
            GeometricPlane cornicePlane = new GeometricPlane(p7s, xAxis, yAxis);

            Line innerStartLine = new Line(p4s, p5s);
            Line innerEndLine = new Line(p4e, p5e);
            Point innerStartPoint = Utility.GetExtendedIntersection(innerStartLine, cornicePlane, 5);
            Point innerEndPoint = Utility.GetExtendedIntersection(innerEndLine, cornicePlane, 5);

            var fourthFace = new RebarLegFace();
            fourthFace.Contour.AddContourPoint(new ContourPoint(startTopPoint, null));
            fourthFace.Contour.AddContourPoint(new ContourPoint(endTopPoint, null));
            fourthFace.Contour.AddContourPoint(new ContourPoint(innerEndPoint, null));
            fourthFace.Contour.AddContourPoint(new ContourPoint(innerStartPoint, null));
            rebarSet.LegFaces.Add(fourthFace);

            var fifthFace = new RebarLegFace();
            fifthFace.Contour.AddContourPoint(new ContourPoint(innerStartPoint, null));
            fifthFace.Contour.AddContourPoint(new ContourPoint(innerEndPoint, null));
            fifthFace.Contour.AddContourPoint(new ContourPoint(p8e, null));
            fifthFace.Contour.AddContourPoint(new ContourPoint(p8s, null));
            rebarSet.LegFaces.Add(fifthFace);

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing),
                SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = number == 0 ? 100 : 50;
            guideline.Spacing.StartOffsetType = number == 0 ? RebarSpacing.OffsetEnum.MINIMUM : RebarSpacing.OffsetEnum.EXACT;
            guideline.Spacing.EndOffset = number == 0 ? 50 : 100;
            guideline.Spacing.EndOffsetType = number == 0 ? RebarSpacing.OffsetEnum.EXACT : RebarSpacing.OffsetEnum.MINIMUM;
            Point correctedP3 = new Point(p3e.X, p3s.Y, p3s.Z);
            guideline.Curve.AddContourPoint(new ContourPoint(p3s, null));
            guideline.Curve.AddContourPoint(new ContourPoint(correctedP3, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            var bottomLengthModifier = new RebarEndDetailModifier();
            bottomLengthModifier.Father = rebarSet;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentLength = GetHookLength(rebarDiameter);
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(p8s, null));
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(p8e, null));
            bottomLengthModifier.Insert();

            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 1, 1, 1, 1 });
        }
        void BackwallInnerVerticalRebar(int number)
        {
            string rebarSize = Program.ExcelDictionary["BVR_Diameter"];
            int rebarDiameter = Convert.ToInt32(rebarSize);
            string spacing = Program.ExcelDictionary["BVR_Spacing"];

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "ABT_BIVR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            int f = number == 0 ? 0 : 1;
            int s = number == 0 ? 1 : 2;

            Point p8s = ProfilePoints[f][8];
            Point p8e = ProfilePoints[s][8];
            Point p7s = ProfilePoints[f][7];
            Point p7e = ProfilePoints[s][7];
            Point p5s = ProfilePoints[f][5];
            Point p5e = ProfilePoints[s][5];
            Point p4s = ProfilePoints[f][4];
            Point p4e = ProfilePoints[s][4];
            Point p3s = ProfilePoints[f][3];
            Point p3e = ProfilePoints[s][3];
            Point p2s = new Point(p3s.X, p5s.Y, p3s.Z);
            Point p2e = new Point(p3e.X, p5e.Y, p3e.Z);

            Vector xAxis = Utility.GetVectorFromTwoPoints(p7s, p8s).GetNormal();
            Vector yAxis = Utility.GetVectorFromTwoPoints(p7s, p7e).GetNormal();
            GeometricPlane cornicePlane = new GeometricPlane(p7s, xAxis, yAxis);

            Line innerStartLine = new Line(p4s, p5s);
            Line outerStartLine = new Line(p3s, p2s);
            Line innerEndLine = new Line(p4e, p5e);
            Line outerEndLine = new Line(p3e, p2e);
            Point innerStartPoint = Utility.GetExtendedIntersection(innerStartLine, cornicePlane, 5);
            Point outerStartPoint = Utility.GetExtendedIntersection(outerStartLine, cornicePlane, 5);
            Point innerEndPoint = Utility.GetExtendedIntersection(innerEndLine, cornicePlane, 5);
            Point outerEndPoint = Utility.GetExtendedIntersection(outerEndLine, cornicePlane, 5);

            var firstFace = new RebarLegFace();
            firstFace.Contour.AddContourPoint(new ContourPoint(p8s, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(p8e, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(outerEndPoint, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(outerStartPoint, null));
            rebarSet.LegFaces.Add(firstFace);

            var secondFace = new RebarLegFace();
            secondFace.Contour.AddContourPoint(new ContourPoint(outerStartPoint, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(outerEndPoint, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(p2e, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(p2s, null));
            rebarSet.LegFaces.Add(secondFace);

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing),
                SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = number == 0 ? 100 : 50;
            guideline.Spacing.StartOffsetType = number == 0 ? RebarSpacing.OffsetEnum.MINIMUM : RebarSpacing.OffsetEnum.EXACT;
            guideline.Spacing.EndOffset = number == 0 ? 50 : 100;
            guideline.Spacing.EndOffsetType = number == 0 ? RebarSpacing.OffsetEnum.EXACT : RebarSpacing.OffsetEnum.MINIMUM;
            Point correctedP2 = new Point(p2e.X, p2s.Y, p2s.Z);
            guideline.Curve.AddContourPoint(new ContourPoint(p2s, null));
            guideline.Curve.AddContourPoint(new ContourPoint(correctedP2, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            var innerLengthModifier = new RebarEndDetailModifier();
            innerLengthModifier.Father = rebarSet;
            innerLengthModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            innerLengthModifier.RebarLengthAdjustment.AdjustmentLength = GetHookLength(rebarDiameter);
            innerLengthModifier.Curve.AddContourPoint(new ContourPoint(innerStartPoint, null));
            innerLengthModifier.Curve.AddContourPoint(new ContourPoint(innerEndPoint, null));
            innerLengthModifier.Insert();

            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 1 });
        }
        void ShelfHorizontalRebar(int number)
        {
            string rebarSize = Program.ExcelDictionary["SHR_Diameter"];
            int rebarDiameter = Convert.ToInt32(rebarSize);
            string spacing = Program.ExcelDictionary["SHR_Spacing"];

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "ABT_SHR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            int f = number == 0 ? 0 : 1;
            int s = number == 0 ? 1 : 2;

            Point p0s = ProfilePoints[f][0];
            Point p0e = ProfilePoints[s][0];
            Point p1s = ProfilePoints[f][1];
            Point p1e = ProfilePoints[s][1];
            Point p2s = ProfilePoints[f][2];
            Point p2e = ProfilePoints[s][2];
            Point p6s = ProfilePoints[f][6];
            Point p6e = ProfilePoints[s][6];
            Point p7s = ProfilePoints[f][7];
            Point p7e = ProfilePoints[s][7];

            Vector xAxis = Utility.GetVectorFromTwoPoints(p6s, p6e).GetNormal();
            Vector yAxis = Utility.GetVectorFromTwoPoints(p7s, p6s).GetNormal();
            GeometricPlane backwallPlane = new GeometricPlane(p7s, xAxis, yAxis);

            Line sLine = new Line(p1s, p2s);
            Line eLine = new Line(p1e, p2e);
            Point startIntersection = Utility.GetExtendedIntersection(sLine, backwallPlane, 5);
            Point endIntersection = Utility.GetExtendedIntersection(eLine, backwallPlane, 5);

            var firstFace = new RebarLegFace();
            firstFace.Contour.AddContourPoint(new ContourPoint(p0s, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(p0e, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(p1e, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(p1s, null));
            rebarSet.LegFaces.Add(firstFace);

            var secondFace = new RebarLegFace();
            secondFace.Contour.AddContourPoint(new ContourPoint(p1s, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(p1e, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(endIntersection, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(startIntersection, null));
            rebarSet.LegFaces.Add(secondFace);

            var thirdFace = new RebarLegFace();
            thirdFace.Contour.AddContourPoint(new ContourPoint(p6s, null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(p6e, null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(endIntersection, null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(startIntersection, null));
            rebarSet.LegFaces.Add(thirdFace);


            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing),
                SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = number == 0 ? 100 : 50;
            guideline.Spacing.StartOffsetType = number == 0 ? RebarSpacing.OffsetEnum.MINIMUM : RebarSpacing.OffsetEnum.EXACT;
            guideline.Spacing.EndOffset = number == 0 ? 50 : 100;
            guideline.Spacing.EndOffsetType = number == 0 ? RebarSpacing.OffsetEnum.EXACT : RebarSpacing.OffsetEnum.MINIMUM; ;

            guideline.Curve.AddContourPoint(new ContourPoint(p1s, null));
            guideline.Curve.AddContourPoint(new ContourPoint(p1e, null));
            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            var bottomLengthModifier = new RebarEndDetailModifier();
            bottomLengthModifier.Father = rebarSet;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentLength = GetHookLength(rebarDiameter);
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(p0s, null));
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(p0e, null));
            bottomLengthModifier.Insert();

            var topLengthModifier = new RebarEndDetailModifier();
            topLengthModifier.Father = rebarSet;
            topLengthModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            topLengthModifier.RebarLengthAdjustment.AdjustmentLength = GetHookLength(rebarDiameter);
            topLengthModifier.Curve.AddContourPoint(new ContourPoint(p6s, null));
            topLengthModifier.Curve.AddContourPoint(new ContourPoint(p6e, null));
            topLengthModifier.Insert();

            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 1, 1 });
        }
        void OuterLongitudinalRebar()
        {
            string rebarSize = Program.ExcelDictionary["OLR_Diameter"];
            int rebarDiameter = Convert.ToInt32(rebarSize);
            string spacing = Program.ExcelDictionary["OLR_Spacing"];
            string secondRebarSize = Program.ExcelDictionary["OLR_SecondDiameter"];
            double startOffset = Convert.ToDouble(Program.ExcelDictionary["OLR_StartOffset"]);
            double secondDiameterLength = Convert.ToDouble(Program.ExcelDictionary["OLR_SecondDiameterLength"]);

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "ABT_OLR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            Point p0s = ProfilePoints[0][0];
            Point p0e = ProfilePoints[2][0];
            Point p1s = ProfilePoints[0][1];
            Point p1m = ProfilePoints[1][1];
            Point p1e = ProfilePoints[2][1];
            Point p0m = new Point(p1m.X, p0s.Y, p1m.Z);

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(p0s, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p0e, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p1e, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p1m, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p1s, null));
            rebarSet.LegFaces.Add(mainFace);

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing),
                SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = startOffset;
            guideline.Spacing.EndOffset = 100;

            Point startGL, endGL;
            if (Height > Height2 && Height > Height3)
            {
                startGL = p0s;
                endGL = p1s;
            }
            else if (Height2 > Height3)
            {
                startGL = p0m;
                endGL = p1m;
            }
            else
            {
                startGL = p0e;
                endGL = p1e;
            }

            guideline.Curve.AddContourPoint(new ContourPoint(startGL, null));
            guideline.Curve.AddContourPoint(new ContourPoint(endGL, null));
            rebarSet.Guidelines.Add(guideline);

            bool succes = rebarSet.Insert();
            new Model().CommitChanges();

            Point ms = new Point(p0e.X, p0e.Y + startOffset + secondDiameterLength, p0e.Z);
            var innerEndDetailModifier = new RebarPropertyModifier();
            innerEndDetailModifier.Father = rebarSet;
            innerEndDetailModifier.BarsAffected = BaseRebarModifier.BarsAffectedEnum.ALL_BARS;
            innerEndDetailModifier.RebarProperties.Size = secondRebarSize;
            innerEndDetailModifier.RebarProperties.Class = SetClass(Convert.ToDouble(secondRebarSize));
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(p0e, null));
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ms, null));
            innerEndDetailModifier.Insert();
            new Model().CommitChanges();

            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2 });
        }
        void InnerLongitudinalRebar()
        {
            string rebarSize = Program.ExcelDictionary["ILR_Diameter"];
            int rebarDiameter = Convert.ToInt32(rebarSize);
            string spacing = Program.ExcelDictionary["ILR_Spacing"];
            string secondRebarSize = Program.ExcelDictionary["ILR_SecondDiameter"];
            double startOffset = Convert.ToDouble(Program.ExcelDictionary["ILR_StartOffset"]);
            double secondDiameterLength = Convert.ToDouble(Program.ExcelDictionary["ILR_SecondDiameterLength"]);

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "ABT_ILR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            Point p9s = ProfilePoints[0][9];
            Point p9m = ProfilePoints[1][9];
            Point p9e = ProfilePoints[2][9];
            Point p8s = ProfilePoints[0][8];
            Point p8m = ProfilePoints[1][8];
            Point p8e = ProfilePoints[2][8];

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(p9s, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p9e, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p8e, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p8m, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p8s, null));
            rebarSet.LegFaces.Add(mainFace);

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing),
                SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = startOffset;
            guideline.Spacing.EndOffset = 0;

            Point startGL, endGL;
            if (Height > Height2 && Height > Height3)
            {
                startGL = p9s;
                endGL = p8s;
            }
            else if (Height2 > Height3)
            {
                startGL = p9m;
                endGL = p8m;
            }
            else
            {
                startGL = p9e;
                endGL = p8e;
            }

            guideline.Curve.AddContourPoint(new ContourPoint(startGL, null));
            guideline.Curve.AddContourPoint(new ContourPoint(endGL, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();

            Point ms = new Point(p9e.X, p9e.Y + startOffset + secondDiameterLength, p9e.Z);
            var innerEndDetailModifier = new RebarPropertyModifier();
            innerEndDetailModifier.Father = rebarSet;
            innerEndDetailModifier.BarsAffected = BaseRebarModifier.BarsAffectedEnum.ALL_BARS;
            innerEndDetailModifier.RebarProperties.Size = secondRebarSize;
            innerEndDetailModifier.RebarProperties.Class = SetClass(Convert.ToDouble(secondRebarSize));
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(p9e, null));
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ms, null));
            innerEndDetailModifier.Insert();
            new Model().CommitChanges();

            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2 });
        }
        void CantileverLongitudinalRebar(int number)
        {
            string rebarSize = Program.ExcelDictionary["CtLR_Diameter"];
            string spacing = Program.ExcelDictionary["CtLR_Spacing"];

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "ABT_CtLR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            Point p8s = ProfilePoints[0][8];
            Point p8m = ProfilePoints[1][8];
            Point p8e = ProfilePoints[2][8];
            Point p7s = ProfilePoints[0][7];
            Point p7m = ProfilePoints[1][7];
            Point p7e = ProfilePoints[2][7];
            Point p6s = ProfilePoints[0][6];
            Point p6m = ProfilePoints[1][6];
            Point p6e = ProfilePoints[2][6];
            Point p5s = ProfilePoints[0][5];
            Point p5m = ProfilePoints[1][5];
            Point p5e = ProfilePoints[2][5];

            Point startGL, endGL;
            if (number == 1)
            {
                var firstFace = new RebarLegFace();
                firstFace.Contour.AddContourPoint(new ContourPoint(p8s, null));
                firstFace.Contour.AddContourPoint(new ContourPoint(p8m, null));
                firstFace.Contour.AddContourPoint(new ContourPoint(p8e, null));
                firstFace.Contour.AddContourPoint(new ContourPoint(p7e, null));
                firstFace.Contour.AddContourPoint(new ContourPoint(p7m, null));
                firstFace.Contour.AddContourPoint(new ContourPoint(p7s, null));
                rebarSet.LegFaces.Add(firstFace);
                if (Height > Height2 && Height > Height3)
                {
                    startGL = p8s;
                    endGL = p7s;
                }
                else if (Height2 > Height3)
                {
                    startGL = p8m;
                    endGL = p7m;
                }
                else
                {
                    startGL = p8e;
                    endGL = p7e;
                }
            }
            else if (number == 2)
            {
                var secondFace = new RebarLegFace();
                secondFace.Contour.AddContourPoint(new ContourPoint(p7s, null));
                secondFace.Contour.AddContourPoint(new ContourPoint(p7m, null));
                secondFace.Contour.AddContourPoint(new ContourPoint(p7e, null));
                secondFace.Contour.AddContourPoint(new ContourPoint(p6e, null));
                secondFace.Contour.AddContourPoint(new ContourPoint(p6m, null));
                secondFace.Contour.AddContourPoint(new ContourPoint(p6s, null));
                rebarSet.LegFaces.Add(secondFace);
                if (Height > Height2 && Height > Height3)
                {
                    startGL = p7s;
                    endGL = p6s;
                }
                else if (Height2 > Height3)
                {
                    startGL = p7m;
                    endGL = p6m;
                }
                else
                {
                    startGL = p7e;
                    endGL = p6e;
                }
            }
            else
            {
                var thirdFace = new RebarLegFace();
                thirdFace.Contour.AddContourPoint(new ContourPoint(p6s, null));
                thirdFace.Contour.AddContourPoint(new ContourPoint(p6m, null));
                thirdFace.Contour.AddContourPoint(new ContourPoint(p6e, null));
                thirdFace.Contour.AddContourPoint(new ContourPoint(p5e, null));
                thirdFace.Contour.AddContourPoint(new ContourPoint(p5m, null));
                thirdFace.Contour.AddContourPoint(new ContourPoint(p5s, null));
                rebarSet.LegFaces.Add(thirdFace);
                if (Height > Height2 && Height > Height3)
                {
                    startGL = p6s;
                    endGL = p5s;
                }
                else if (Height2 > Height3)
                {
                    startGL = p6m;
                    endGL = p5m;
                }
                else
                {
                    startGL = p6e;
                    endGL = p5e;
                }
            }

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing),
                SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = number == 1 ? 0 : 100;
            guideline.Spacing.EndOffset = number == 3 ? 0 : 100;
            guideline.Curve.AddContourPoint(new ContourPoint(startGL, null));
            guideline.Curve.AddContourPoint(new ContourPoint(endGL, null));
            rebarSet.Guidelines.Add(guideline);

            bool succes = rebarSet.Insert();

            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2 });
        }
        void BackwallLongitudinalRebar(int number)
        {
            string rebarSize = Program.ExcelDictionary["BLR_Diameter"];
            int rebarDiameter = Convert.ToInt32(rebarSize);
            string spacing = Program.ExcelDictionary["BLR_Spacing"];

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "ABT_BLR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            Point p4s = ProfilePoints[0][4];
            Point p4m = ProfilePoints[1][4];
            Point p4e = ProfilePoints[2][4];
            Point p3s = ProfilePoints[0][3];
            Point p3m = ProfilePoints[1][3];
            Point p3e = ProfilePoints[2][3];
            Point p2s = ProfilePoints[0][2];
            Point p2m = ProfilePoints[1][2];
            Point p2e = ProfilePoints[2][2];
            Point p5s = ProfilePoints[0][5];
            Point p5m = ProfilePoints[1][5];
            Point p5e = ProfilePoints[2][5];

            Point startGL;
            Point endGL;

            if (number == 1)
            {
                var firstFace = new RebarLegFace();
                firstFace.Contour.AddContourPoint(new ContourPoint(p2s, null));
                firstFace.Contour.AddContourPoint(new ContourPoint(p2m, null));
                firstFace.Contour.AddContourPoint(new ContourPoint(p2e, null));
                firstFace.Contour.AddContourPoint(new ContourPoint(p3e, null));
                firstFace.Contour.AddContourPoint(new ContourPoint(p3m, null));
                firstFace.Contour.AddContourPoint(new ContourPoint(p3s, null));
                rebarSet.LegFaces.Add(firstFace);
                if (Height > Height2 && Height > Height3)
                {
                    startGL = p2s;
                    endGL = p3s;
                }
                else if (Height2 > Height3)
                {
                    startGL = p2m;
                    endGL = p3m;
                }
                else
                {
                    startGL = p2e;
                    endGL = p3e;
                }
            }
            else if (number == 2)
            {
                var secondFace = new RebarLegFace();
                secondFace.Contour.AddContourPoint(new ContourPoint(p3s, null));
                secondFace.Contour.AddContourPoint(new ContourPoint(p3m, null));
                secondFace.Contour.AddContourPoint(new ContourPoint(p4m, null));
                secondFace.Contour.AddContourPoint(new ContourPoint(p4s, null));
                rebarSet.LegFaces.Add(secondFace);
                var secondFace2 = new RebarLegFace();
                secondFace2.Contour.AddContourPoint(new ContourPoint(p3m, null));
                secondFace2.Contour.AddContourPoint(new ContourPoint(p3e, null));
                secondFace2.Contour.AddContourPoint(new ContourPoint(p4e, null));
                secondFace2.Contour.AddContourPoint(new ContourPoint(p4m, null));
                rebarSet.LegFaces.Add(secondFace2);
                if (Height > Height2 && Height > Height3)
                {
                    startGL = p3s;
                    endGL = p4s;
                }
                else if (Height2 > Height3)
                {
                    startGL = p3m;
                    endGL = p4m;
                }
                else
                {
                    startGL = p3e;
                    endGL = p4e;
                }
            }
            else
            {
                var thirdFace = new RebarLegFace();
                thirdFace.Contour.AddContourPoint(new ContourPoint(p4s, null));
                thirdFace.Contour.AddContourPoint(new ContourPoint(p4m, null));
                thirdFace.Contour.AddContourPoint(new ContourPoint(p4e, null));
                thirdFace.Contour.AddContourPoint(new ContourPoint(p5e, null));
                thirdFace.Contour.AddContourPoint(new ContourPoint(p5m, null));
                thirdFace.Contour.AddContourPoint(new ContourPoint(p5s, null));
                rebarSet.LegFaces.Add(thirdFace);
                double minY;
                if (p5s.Y < p5m.Y && p5s.Y < p5e.Y)
                {
                    minY = p5s.Y;
                }
                else if (p5m.Y < p5e.Y)
                {
                    minY = p5m.Y;
                }
                else
                {
                    minY = p5e.Y;
                }
                if (Height > Height2 && Height > Height3)
                {
                    startGL = p4s;
                    endGL = new Point(p5s.X, minY, p5s.Z);
                }
                else if (Height2 > Height3)
                {
                    startGL = p4m;
                    endGL = new Point(p5m.X, minY, p5m.Z);
                }
                else
                {
                    startGL = p4e;
                    endGL = new Point(p5e.X, minY, p5e.Z);
                }
            }

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing),
                SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = number == 1 ? 0 : 100;
            guideline.Spacing.EndOffset = number == 3 ? 0 : 100;

            guideline.Curve.AddContourPoint(new ContourPoint(startGL, null));
            guideline.Curve.AddContourPoint(new ContourPoint(endGL, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            if (number == 2)
            {
                RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2, 2 });
            }
            else
            {
                RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2 });
            }
        }
        void ShelfLongitudinalRebar()
        {
            string rebarSize = Program.ExcelDictionary["SLR_Diameter"];
            int rebarDiameter = Convert.ToInt32(rebarSize);
            string spacing = Program.ExcelDictionary["SLR_Spacing"];

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "ABT_SLR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 2;

            Point p1s = ProfilePoints[0][1];
            Point p1m = ProfilePoints[1][1];
            Point p1e = ProfilePoints[2][1];
            Point p2s = ProfilePoints[0][2];
            Point p2m = ProfilePoints[1][2];
            Point p2e = ProfilePoints[2][2];

            var firstFace = new RebarLegFace();
            firstFace.Contour.AddContourPoint(new ContourPoint(p1s, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(p1m, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(p2m, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(p2s, null));
            rebarSet.LegFaces.Add(firstFace);

            var secondFace = new RebarLegFace();
            secondFace.Contour.AddContourPoint(new ContourPoint(p1m, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(p1e, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(p2e, null));
            secondFace.Contour.AddContourPoint(new ContourPoint(p2m, null));
            rebarSet.LegFaces.Add(secondFace);

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing),
                SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = 100;
            guideline.Spacing.EndOffset = 0;

            guideline.Curve.AddContourPoint(new ContourPoint(p1s, null));
            guideline.Curve.AddContourPoint(new ContourPoint(p2s, null));
            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2,2 });
        }
        /// <summary>
        /// Create closing C shape rebar
        /// </summary>
        /// <param name="number">0 for start section, 1 for end</param>
        void ClosingCShapeRebarBottom(int number)
        {
            string rebarSize = Program.ExcelDictionary["CCSR_Diameter"];
            int rebarDiameter = Convert.ToInt32(rebarSize);
            string spacing = Program.ExcelDictionary["CCSR_Spacing"];
            double offset = Convert.ToDouble(Program.ExcelDictionary["ILR_StartOffset"]);

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "ABT_CCSRB";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 2;

            int secondNumber = number == 0 ? 2 : 0;
            Point p0 = ProfilePoints[number][0];
            Point p1 = ProfilePoints[number][1];
            Point p3 = ProfilePoints[number][3];
            Point p4 = ProfilePoints[number][4];
            Point p5 = ProfilePoints[number][5];
            Point p6 = ProfilePoints[number][6];
            Point p7 = ProfilePoints[number][7];
            Point p8 = ProfilePoints[number][8];
            Point p9 = ProfilePoints[number][9];
            Point p2 = new Point(p7.X, p1.Y, p7.Z);

            List<Point> listPoint = new List<Point> { p0, p1, p2, p3, p4, p5, p6, p7, p8, p9 };

            Point p0e = ProfilePoints[secondNumber][0];

            Vector sectionDir = Utility.GetVectorFromTwoPoints(p0, p0e).GetNormal();
            List<Point> offsetedPoints = new List<Point>();
            for (int i = 0; i < listPoint.Count; i++)
            {
                Point correctedPoint = Utility.TranslePointByVectorAndDistance(listPoint[i], sectionDir, 2 * GetHookLength(rebarDiameter));
                offsetedPoints.Add(correctedPoint);
            }

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(p0, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p1, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p2, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p7, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p8, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p9, null));
            rebarSet.LegFaces.Add(mainFace);

            var firstFace = new RebarLegFace();
            firstFace.Contour.AddContourPoint(new ContourPoint(p0, null));
            firstFace.Contour.AddContourPoint(new ContourPoint(offsetedPoints[0], null));
            firstFace.Contour.AddContourPoint(new ContourPoint(offsetedPoints[1], null));
            firstFace.Contour.AddContourPoint(new ContourPoint(p1, null));
            rebarSet.LegFaces.Add(firstFace);

            var fourthFace = new RebarLegFace();
            fourthFace.Contour.AddContourPoint(new ContourPoint(p9, null));
            fourthFace.Contour.AddContourPoint(new ContourPoint(offsetedPoints[9], null));
            fourthFace.Contour.AddContourPoint(new ContourPoint(offsetedPoints[8], null));
            fourthFace.Contour.AddContourPoint(new ContourPoint(p8, null));
            rebarSet.LegFaces.Add(fourthFace);

            var fifthFace = new RebarLegFace();
            fifthFace.Contour.AddContourPoint(new ContourPoint(p8, null));
            fifthFace.Contour.AddContourPoint(new ContourPoint(offsetedPoints[8], null));
            fifthFace.Contour.AddContourPoint(new ContourPoint(offsetedPoints[7], null));
            fifthFace.Contour.AddContourPoint(new ContourPoint(p7, null));
            rebarSet.LegFaces.Add(fifthFace);

            var sixthFace = new RebarLegFace();
            sixthFace.Contour.AddContourPoint(new ContourPoint(p7, null));
            sixthFace.Contour.AddContourPoint(new ContourPoint(offsetedPoints[7], null));
            sixthFace.Contour.AddContourPoint(new ContourPoint(offsetedPoints[2], null));
            sixthFace.Contour.AddContourPoint(new ContourPoint(p2, null));
            rebarSet.LegFaces.Add(sixthFace);

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

            guideline.Curve.AddContourPoint(new ContourPoint(p0, null));
            guideline.Curve.AddContourPoint(new ContourPoint(p1, null));
            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            var firstEndDetailModifier = new RebarEndDetailModifier();
            firstEndDetailModifier.Father = rebarSet;
            firstEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            firstEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = GetHookLength(rebarDiameter);
            firstEndDetailModifier.Curve.AddContourPoint(new ContourPoint(offsetedPoints[0], null));
            firstEndDetailModifier.Curve.AddContourPoint(new ContourPoint(offsetedPoints[1], null));
            firstEndDetailModifier.Insert();

            var secondEndDetailModifier = new RebarEndDetailModifier();
            secondEndDetailModifier.Father = rebarSet;
            secondEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            secondEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = GetHookLength(rebarDiameter);
            secondEndDetailModifier.Curve.AddContourPoint(new ContourPoint(offsetedPoints[9], null));
            secondEndDetailModifier.Curve.AddContourPoint(new ContourPoint(offsetedPoints[8], null));
            secondEndDetailModifier.Curve.AddContourPoint(new ContourPoint(offsetedPoints[7], null));
            secondEndDetailModifier.Curve.AddContourPoint(new ContourPoint(offsetedPoints[2], null));
            secondEndDetailModifier.Insert();

            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 2, 2, 2, 2 });
        }
        void ClosingCShapeRebarMid(int number)
        {
            string rebarSize = Program.ExcelDictionary["CCSR_Diameter"];
            int rebarDiameter = Convert.ToInt32(rebarSize);
            string spacing = Program.ExcelDictionary["CCSR_Spacing"];

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "ABT_CCSRM";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 2;

            int secondNumber = number == 0 ? 2 : 0;
            Point p0 = ProfilePoints[number][0];
            Point p1 = ProfilePoints[number][1];
            Point p2 = ProfilePoints[number][2];
            Point p3 = ProfilePoints[number][3];
            Point p4 = ProfilePoints[number][4];
            Point p5 = ProfilePoints[number][5];
            Point p6 = ProfilePoints[number][6];
            Point p7 = ProfilePoints[number][7];
            Point p8 = ProfilePoints[number][8];
            Point p9 = ProfilePoints[number][9];
            p2 = new Point(p2.X, p1.Y, p2.Z);
            p7 = new Point(p7.X, p1.Y, p7.Z);
            p3 = new Point(p3.X, p6.Y, p3.Z);

            List<Point> listPoint = new List<Point> { p0, p1, p2, p3, p4, p5, p6, p7, p8, p9 };

            Point p0e = ProfilePoints[secondNumber][0];

            Vector sectionDir = Utility.GetVectorFromTwoPoints(p0, p0e).GetNormal();
            List<Point> offsetedPoints = new List<Point>();
            for (int i = 0; i < listPoint.Count; i++)
            {
                Point correctedPoint = Utility.TranslePointByVectorAndDistance(listPoint[i], sectionDir, 2 * GetHookLength(rebarDiameter));
                offsetedPoints.Add(correctedPoint);
            }

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(p2, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p3, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p6, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p7, null));
            rebarSet.LegFaces.Add(mainFace);

            var thirdFace = new RebarLegFace();
            thirdFace.Contour.AddContourPoint(new ContourPoint(p2, null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(offsetedPoints[2], null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(offsetedPoints[3], null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(p3, null));
            rebarSet.LegFaces.Add(thirdFace);

            var sixthFace = new RebarLegFace();
            sixthFace.Contour.AddContourPoint(new ContourPoint(p7, null));
            sixthFace.Contour.AddContourPoint(new ContourPoint(offsetedPoints[7], null));
            sixthFace.Contour.AddContourPoint(new ContourPoint(offsetedPoints[6], null));
            sixthFace.Contour.AddContourPoint(new ContourPoint(p6, null));
            rebarSet.LegFaces.Add(sixthFace);

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing),
                SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = 0;
            guideline.Spacing.EndOffset = 100;

            guideline.Curve.AddContourPoint(new ContourPoint(p2, null));
            guideline.Curve.AddContourPoint(new ContourPoint(p3, null));
            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            var thirdEndDetailModifier = new RebarEndDetailModifier();
            thirdEndDetailModifier.Father = rebarSet;
            thirdEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            thirdEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = GetHookLength(rebarDiameter);
            thirdEndDetailModifier.Curve.AddContourPoint(new ContourPoint(offsetedPoints[2], null));
            thirdEndDetailModifier.Curve.AddContourPoint(new ContourPoint(offsetedPoints[3], null));
            thirdEndDetailModifier.Insert();

            var secondEndDetailModifier = new RebarEndDetailModifier();
            secondEndDetailModifier.Father = rebarSet;
            secondEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            secondEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = GetHookLength(rebarDiameter);
            secondEndDetailModifier.Curve.AddContourPoint(new ContourPoint(offsetedPoints[7], null));
            secondEndDetailModifier.Curve.AddContourPoint(new ContourPoint(offsetedPoints[6], null));
            secondEndDetailModifier.Insert();

            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 2, 2 });
        }
        void ClosingCShapeRebarTop(int number)
        {
            string rebarSize = Program.ExcelDictionary["CCSR_Diameter"];
            int rebarDiameter = Convert.ToInt32(rebarSize);
            string spacing = Program.ExcelDictionary["CCSR_Spacing"];

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "ABT_CCSRT";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 2;

            int secondNumber = number == 0 ? 2 : 0;
            Point p0 = ProfilePoints[number][0];
            Point p1 = ProfilePoints[number][1];
            Point p2 = ProfilePoints[number][2];
            Point p3 = ProfilePoints[number][3];
            Point p4 = ProfilePoints[number][4];
            Point p5 = ProfilePoints[number][5];
            Point p6 = ProfilePoints[number][6];
            Point p7 = ProfilePoints[number][7];
            Point p8 = ProfilePoints[number][8];
            Point p9 = ProfilePoints[number][9];
            p2 = new Point(p2.X, p6.Y, p2.Z);
            p5 = new Point(p5.X, p6.Y, p5.Z);

            List<Point> listPoint = new List<Point> { p0, p1, p2, p3, p4, p5, p6, p7, p8, p9 };

            Point p0e = ProfilePoints[secondNumber][0];

            Vector sectionDir = Utility.GetVectorFromTwoPoints(p0, p0e).GetNormal();
            List<Point> offsetedPoints = new List<Point>();
            for (int i = 0; i < listPoint.Count; i++)
            {
                Point correctedPoint = Utility.TranslePointByVectorAndDistance(listPoint[i], sectionDir, 2 * GetHookLength(rebarDiameter));
                offsetedPoints.Add(correctedPoint);
            }

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(p2, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p3, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p4, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p5, null));
            rebarSet.LegFaces.Add(mainFace);

            var thirdFace = new RebarLegFace();
            thirdFace.Contour.AddContourPoint(new ContourPoint(p2, null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(offsetedPoints[2], null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(offsetedPoints[3], null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(p3, null));
            rebarSet.LegFaces.Add(thirdFace);

            var eighthFace = new RebarLegFace();
            eighthFace.Contour.AddContourPoint(new ContourPoint(p5, null));
            eighthFace.Contour.AddContourPoint(new ContourPoint(offsetedPoints[5], null));
            eighthFace.Contour.AddContourPoint(new ContourPoint(offsetedPoints[4], null));
            eighthFace.Contour.AddContourPoint(new ContourPoint(p4, null));
            rebarSet.LegFaces.Add(eighthFace);

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing),
                SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = 0;
            guideline.Spacing.EndOffset = 100;

            guideline.Curve.AddContourPoint(new ContourPoint(p2, null));
            guideline.Curve.AddContourPoint(new ContourPoint(p3, null));
            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            var thirdEndDetailModifier = new RebarEndDetailModifier();
            thirdEndDetailModifier.Father = rebarSet;
            thirdEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            thirdEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = GetHookLength(rebarDiameter);
            thirdEndDetailModifier.Curve.AddContourPoint(new ContourPoint(offsetedPoints[2], null));
            thirdEndDetailModifier.Curve.AddContourPoint(new ContourPoint(offsetedPoints[3], null));
            thirdEndDetailModifier.Insert();

            var secondEndDetailModifier = new RebarEndDetailModifier();
            secondEndDetailModifier.Father = rebarSet;
            secondEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            secondEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = GetHookLength(rebarDiameter);
            secondEndDetailModifier.Curve.AddContourPoint(new ContourPoint(offsetedPoints[5], null));
            secondEndDetailModifier.Curve.AddContourPoint(new ContourPoint(offsetedPoints[4], null));
            secondEndDetailModifier.Insert();

            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 2, 2 });
        }
        void CShapeRebarCommon()
        {
            string rebarSize = Program.ExcelDictionary["CSR_Diameter"];
            int rebarDiameter = Convert.ToInt32(rebarSize);
            double horizontalSpacing = Convert.ToDouble(Program.ExcelDictionary["CSR_HorizontalSpacing"]);
            string verticalSpacing = Program.ExcelDictionary["CSR_VerticalSpacing"];
            double startOffset = Convert.ToDouble(Program.ExcelDictionary["OLR_StartOffset"]);

            List<double> min1YList = new List<double> { ProfilePoints[0][1].Y, ProfilePoints[1][1].Y, ProfilePoints[2][1].Y };
            double min1Y = (from double y in min1YList
                           orderby Math.Abs(y) ascending
                           select y).First();

            double height = Math.Abs(min1Y) + Math.Abs(ProfilePoints[0][0].Y);
            double correctedHeight = height - startOffset;
            int corTotalNumberOfRows = (int)Math.Ceiling(correctedHeight / Convert.ToDouble(verticalSpacing));
            double offset = startOffset + 10 * Convert.ToInt32(rebarSize);
            Vector dir = Utility.GetVectorFromTwoPoints(ProfilePoints[0][8], ProfilePoints[0][9]).GetNormal();

            List<double> min8YList = new List<double> { ProfilePoints[0][8].Y, ProfilePoints[1][8].Y, ProfilePoints[2][8].Y };
            double simpleHeight = (from double y in min8YList
                                   orderby Math.Abs(y) ascending
                                   select y).First();
            simpleHeight += Math.Abs(ProfilePoints[0][0].Y);
            for (int i = 0; i < corTotalNumberOfRows; i++)
            {
                double newoffset = offset + i * Convert.ToDouble(verticalSpacing);
                if (newoffset <= simpleHeight)
                {
                    CShapeRebarSimple(rebarDiameter, newoffset, dir, horizontalSpacing);
                }
                else
                {
                    CShapeRebarComplex(rebarDiameter, newoffset, horizontalSpacing);
                }
            }
        }
        void CShapeRebarSimple(int rebarDiameter, double newoffset, Vector dir, double horizontalSpacing)
        {
            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "ABT_CSRS";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarDiameter));
            rebarSet.RebarProperties.Size = Convert.ToString(rebarDiameter);
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarDiameter));
            rebarSet.LayerOrderNumber = 1;

            Point startOuterPoint = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y + newoffset, ProfilePoints[0][0].Z);
            Point endOuterPoint = new Point(ProfilePoints[2][0].X, ProfilePoints[2][0].Y + newoffset, ProfilePoints[2][0].Z);
            Point startInnerPoint = new Point(ProfilePoints[0][9].X, ProfilePoints[0][9].Y + newoffset, ProfilePoints[0][9].Z);
            Point endInnerPoint = new Point(ProfilePoints[2][9].X, ProfilePoints[2][9].Y + newoffset, ProfilePoints[2][9].Z);

            Point bottomStartOuterPoint = Utility.Translate(startOuterPoint, dir, 2 * GetHookLength(rebarDiameter));
            Point bottomStartInnerPoint = Utility.Translate(startInnerPoint, dir, 2 * GetHookLength(rebarDiameter));
            Point bottomEndtOuterPoint = Utility.Translate(endOuterPoint, dir, 2 * GetHookLength(rebarDiameter));
            Point bottomEndInnerPoint = Utility.Translate(endInnerPoint, dir, 2 * GetHookLength(rebarDiameter));

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(startOuterPoint, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(endOuterPoint, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(endInnerPoint, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(startInnerPoint, null));
            rebarSet.LegFaces.Add(mainFace);

            var innerFace = new RebarLegFace();
            innerFace.Contour.AddContourPoint(new ContourPoint(startOuterPoint, null));
            innerFace.Contour.AddContourPoint(new ContourPoint(endOuterPoint, null));
            innerFace.Contour.AddContourPoint(new ContourPoint(bottomEndtOuterPoint, null));
            innerFace.Contour.AddContourPoint(new ContourPoint(bottomStartOuterPoint, null));
            rebarSet.LegFaces.Add(innerFace);

            var outerFace = new RebarLegFace();
            outerFace.Contour.AddContourPoint(new ContourPoint(startInnerPoint, null));
            outerFace.Contour.AddContourPoint(new ContourPoint(endInnerPoint, null));
            outerFace.Contour.AddContourPoint(new ContourPoint(bottomEndInnerPoint, null));
            outerFace.Contour.AddContourPoint(new ContourPoint(bottomStartInnerPoint, null));
            rebarSet.LegFaces.Add(outerFace);

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(horizontalSpacing),
                SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = 100;
            guideline.Spacing.EndOffset = 100;

            guideline.Curve.AddContourPoint(new ContourPoint(startOuterPoint, null));
            guideline.Curve.AddContourPoint(new ContourPoint(endOuterPoint, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();

            //Create RebarEndDetailModifier
            var innerEndDetailModifier = new RebarEndDetailModifier();
            innerEndDetailModifier.Father = rebarSet;
            innerEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            innerEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = GetHookLength(rebarDiameter);
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(bottomStartOuterPoint, null));
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(bottomEndtOuterPoint, null));
            innerEndDetailModifier.Insert();

            var outerEndDetailModifier = new RebarEndDetailModifier();
            outerEndDetailModifier.Father = rebarSet;
            outerEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            outerEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = GetHookLength(rebarDiameter);
            outerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(bottomStartInnerPoint, null));
            outerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(bottomEndInnerPoint, null));
            outerEndDetailModifier.Insert();
            new Model().CommitChanges();

            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 1, 1 });
        }
        void CShapeRebarComplex(int rebarDiameter, double newoffset, double horizontalSpacing)
        {
            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "ABT_CSRC";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarDiameter));
            rebarSet.RebarProperties.Size = Convert.ToString(rebarDiameter);
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarDiameter));
            rebarSet.LayerOrderNumber = 1;


            Point startOuterPoint = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y + newoffset, ProfilePoints[0][0].Z);
            Point midOuterPoint = new Point(ProfilePoints[1][0].X, ProfilePoints[1][0].Y + newoffset, ProfilePoints[1][0].Z);
            Point endOuterPoint = new Point(ProfilePoints[2][0].X, ProfilePoints[2][0].Y + newoffset, ProfilePoints[2][0].Z);
            Point startInnerPoint = new Point(ProfilePoints[0][9].X, ProfilePoints[0][9].Y + newoffset, ProfilePoints[0][9].Z);
            Point midInnerPoint = new Point(ProfilePoints[1][9].X, ProfilePoints[1][9].Y + newoffset, ProfilePoints[1][9].Z);
            Point endInnerPoint = new Point(ProfilePoints[2][9].X, ProfilePoints[2][9].Y + newoffset, ProfilePoints[2][9].Z);


            Line startLine = new Line(startOuterPoint, startInnerPoint);
            Line endLine = new Line(endOuterPoint, endInnerPoint);
            Line midLine = new Line(midOuterPoint, midInnerPoint);

            Vector outerXdir = Utility.GetVectorFromTwoPoints(ProfilePoints[0][0], ProfilePoints[1][0]).GetNormal();
            Vector outerYdir = Utility.GetVectorFromTwoPoints(ProfilePoints[0][0], ProfilePoints[0][1]).GetNormal();
            Vector innerXdirFirst, innerYdirFirst, innerXdirSecond, innerYdirSecond;
            Point innerPlaneOriginFirst,innerPlaneOriginSecond;
            if (newoffset < Math.Abs(ProfilePoints[0][9].Y) + Math.Abs(ProfilePoints[0][7].Y))
            {
                innerXdirFirst = Utility.GetVectorFromTwoPoints(ProfilePoints[0][8], ProfilePoints[1][8]).GetNormal();
                innerYdirFirst = Utility.GetVectorFromTwoPoints(ProfilePoints[0][8], ProfilePoints[0][7]).GetNormal();
                innerXdirSecond = Utility.GetVectorFromTwoPoints(ProfilePoints[2][8], ProfilePoints[1][8]).GetNormal();
                innerYdirSecond = Utility.GetVectorFromTwoPoints(ProfilePoints[2][8], ProfilePoints[2][7]).GetNormal();
                innerPlaneOriginFirst = ProfilePoints[0][8];
                innerPlaneOriginSecond = ProfilePoints[2][8];
            }
            else
            {
                innerXdirFirst = Utility.GetVectorFromTwoPoints(ProfilePoints[0][7], ProfilePoints[1][7]).GetNormal();
                innerYdirFirst = Utility.GetVectorFromTwoPoints(ProfilePoints[0][7], ProfilePoints[0][6]).GetNormal();
                innerXdirSecond = Utility.GetVectorFromTwoPoints(ProfilePoints[2][7], ProfilePoints[1][7]).GetNormal();
                innerYdirSecond = Utility.GetVectorFromTwoPoints(ProfilePoints[2][7], ProfilePoints[2][6]).GetNormal();
                innerPlaneOriginFirst = ProfilePoints[0][7];
                innerPlaneOriginSecond = ProfilePoints[2][7];
            }
            GeometricPlane outerPlane = new GeometricPlane(ProfilePoints[0][0], outerXdir, outerYdir);
            GeometricPlane innerPlaneFirst = new GeometricPlane(innerPlaneOriginFirst, innerXdirFirst, innerYdirFirst);
            GeometricPlane innerPlaneSecond = new GeometricPlane(innerPlaneOriginSecond, innerXdirSecond, innerYdirSecond);

            Point startOuterIntersection = Utility.GetExtendedIntersection(startLine, outerPlane, 10);
            Point startInnerIntersection = Utility.GetExtendedIntersection(startLine, innerPlaneFirst, 10);
            Point midInnerIntersection = Utility.GetExtendedIntersection(midLine, innerPlaneFirst, 10);
            Point endOuterIntersection = Utility.GetExtendedIntersection(endLine, outerPlane, 10);
            Point endInnerIntersection = Utility.GetExtendedIntersection(endLine, innerPlaneSecond, 10);

            Point bottomStartOuterPoint = Utility.Translate(startOuterIntersection, outerYdir, -2 * GetHookLength(rebarDiameter));
            Point bottomStartInnerPoint = Utility.Translate(startInnerIntersection, innerYdirFirst, 2 * GetHookLength(rebarDiameter));
            Point bottomMidInnerPoint = Utility.Translate(midInnerIntersection, innerYdirFirst, 2 * GetHookLength(rebarDiameter));
            Point bottomEndtOuterPoint = Utility.Translate(endOuterIntersection, outerYdir, -2 * GetHookLength(rebarDiameter));
            Point bottomEndInnerPoint = Utility.Translate(endInnerIntersection, innerYdirFirst, 2 * GetHookLength(rebarDiameter));

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(startOuterPoint, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(endOuterPoint, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(endInnerIntersection, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(midInnerIntersection, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(startInnerIntersection, null));
            rebarSet.LegFaces.Add(mainFace);

            var innerFace = new RebarLegFace();
            innerFace.Contour.AddContourPoint(new ContourPoint(startOuterPoint, null));
            innerFace.Contour.AddContourPoint(new ContourPoint(endOuterPoint, null));
            innerFace.Contour.AddContourPoint(new ContourPoint(bottomEndtOuterPoint, null));
            innerFace.Contour.AddContourPoint(new ContourPoint(bottomStartOuterPoint, null));
            rebarSet.LegFaces.Add(innerFace);

            var outerFace = new RebarLegFace();
            outerFace.Contour.AddContourPoint(new ContourPoint(startInnerIntersection, null));
            outerFace.Contour.AddContourPoint(new ContourPoint(midInnerIntersection, null));
            outerFace.Contour.AddContourPoint(new ContourPoint(endInnerIntersection, null));
            outerFace.Contour.AddContourPoint(new ContourPoint(bottomEndInnerPoint, null));
            outerFace.Contour.AddContourPoint(new ContourPoint(bottomMidInnerPoint, null));
            outerFace.Contour.AddContourPoint(new ContourPoint(bottomStartInnerPoint, null));
            rebarSet.LegFaces.Add(outerFace);

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(horizontalSpacing),
                SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = 100;
            guideline.Spacing.EndOffset = 100;

            guideline.Curve.AddContourPoint(new ContourPoint(startOuterPoint, null));
            guideline.Curve.AddContourPoint(new ContourPoint(endOuterPoint, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();

            //Create RebarEndDetailModifier
            var innerEndDetailModifier = new RebarEndDetailModifier();
            innerEndDetailModifier.Father = rebarSet;
            innerEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            innerEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = GetHookLength(rebarDiameter);
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(bottomStartOuterPoint, null));
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(bottomEndtOuterPoint, null));
            innerEndDetailModifier.Insert();

            var outerEndDetailModifier = new RebarEndDetailModifier();
            outerEndDetailModifier.Father = rebarSet;
            outerEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            outerEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = GetHookLength(rebarDiameter);
            outerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(bottomStartInnerPoint, null));
            outerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(bottomMidInnerPoint, null));
            outerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(bottomEndInnerPoint, null));
            outerEndDetailModifier.Insert();
            new Model().CommitChanges();

            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 1, 1 });
        }
        void ClosingLongitudianlRebarBottom(int number)
        {
            string rebarSize = Program.ExcelDictionary["CLR_Diameter"];
            int rebarDiameter = Convert.ToInt32(rebarSize);
            string spacing = Program.ExcelDictionary["CLR_Spacing"];

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "ABT_CLR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 2;

            int secondNumber = number == 0 ? 2 : 0;
            Point p0 = ProfilePoints[number][0];
            Point p1 = ProfilePoints[number][1];
            Point p2 = ProfilePoints[number][2];
            Point p3 = ProfilePoints[number][3];
            Point p4 = ProfilePoints[number][4];
            Point p5 = ProfilePoints[number][5];
            Point p6 = ProfilePoints[number][6];
            Point p7 = ProfilePoints[number][7];
            Point p8 = ProfilePoints[number][8];
            Point p9 = ProfilePoints[number][9];
            Point p0e = ProfilePoints[secondNumber][0];

            Point p1e = ProfilePoints[secondNumber][1];

            Vector xDirTop = Utility.GetVectorFromTwoPoints(p1, p2);
            Vector yDirTop = Utility.GetVectorFromTwoPoints(p1, p1e);
            GeometricPlane geometricPlaneTop = new GeometricPlane(p1, xDirTop, yDirTop);

            Line line = new Line(p9, p8);
            Point correctedP2 = Utility.GetExtendedIntersection(line, geometricPlaneTop, 10);

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(p0, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p1, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(correctedP2, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p9, null));
            rebarSet.LegFaces.Add(mainFace);

            int c = number == 0 ? 1 : -1;

            Point offsetP0 = new Point(p0.X + c * 40 * rebarDiameter, p0.Y, p0.Z);
            Point offsetP9 = new Point(p9.X + c * 40 * rebarDiameter, p9.Y, p9.Z);

            var offsetFace = new RebarLegFace();
            offsetFace.Contour.AddContourPoint(new ContourPoint(p0, null));
            offsetFace.Contour.AddContourPoint(new ContourPoint(offsetP0, null));
            offsetFace.Contour.AddContourPoint(new ContourPoint(offsetP9, null));
            offsetFace.Contour.AddContourPoint(new ContourPoint(p9, null));
            rebarSet.LegFaces.Add(offsetFace);

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

            guideline.Curve.AddContourPoint(new ContourPoint(p0, null));
            guideline.Curve.AddContourPoint(new ContourPoint(p9, null));
            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            var firstEndDetailModifier = new RebarEndDetailModifier();
            firstEndDetailModifier.Father = rebarSet;
            firstEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            firstEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = GetHookLength(rebarDiameter);
            firstEndDetailModifier.Curve.AddContourPoint(new ContourPoint(offsetP0, null));
            firstEndDetailModifier.Curve.AddContourPoint(new ContourPoint(offsetP9, null));
            firstEndDetailModifier.Insert();

            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2, 2 });
        }
        void ClosingLongitudianlRebarMid(int number)
        {
            string rebarSize = Program.ExcelDictionary["CLR_Diameter"];
            int rebarDiameter = Convert.ToInt32(rebarSize);
            string spacing = Program.ExcelDictionary["CLR_Spacing"];

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "ABT_CLR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 2;

            int secondNumber = 1;
            Point p1 = ProfilePoints[number][1];
            Point p2 = ProfilePoints[number][2];
            Point p3 = ProfilePoints[number][3];
            Point p7 = ProfilePoints[number][7];
            Point p8 = ProfilePoints[number][8];
            Point p9 = ProfilePoints[number][9];

            Point p1e = ProfilePoints[secondNumber][1];
            Point p2e = ProfilePoints[secondNumber][2];
            Point p3e = ProfilePoints[secondNumber][3];
            Point p9e = ProfilePoints[secondNumber][9];
            Point p8e = ProfilePoints[secondNumber][8];

            Vector xDirTop = Utility.GetVectorFromTwoPoints(p1, p2);
            Vector yDirTop = Utility.GetVectorFromTwoPoints(p1, p1e);
            GeometricPlane geometricPlaneTop = new GeometricPlane(p1, xDirTop, yDirTop);

            Vector xDirBottom = Utility.GetVectorFromTwoPoints(p8, p7);
            Vector yDirBottom = Utility.GetVectorFromTwoPoints(p8, p8e);
            GeometricPlane geometricPlaneBottom = new GeometricPlane(p8, xDirBottom, yDirBottom);

            Line startLine89 = new Line(p9, p8);
            Line endLine89 = new Line(p9e, p8e);
            Line startLine32 = new Line(p3, p2);
            Line endLine32 = new Line(p3e, p2e);

            Point correctedP8s = Utility.GetExtendedIntersection(startLine89, geometricPlaneTop, 2);
            Point correctedP8e = Utility.GetExtendedIntersection(endLine89, geometricPlaneTop, 2);
            Point correctedP2s = Utility.GetExtendedIntersection(startLine32, geometricPlaneBottom, 4);
            Point correctedP2e = Utility.GetExtendedIntersection(endLine32, geometricPlaneBottom, 4);

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(p8, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(correctedP8s, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p2, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(correctedP2s, null));
            rebarSet.LegFaces.Add(mainFace);

            var bottomFace = new RebarLegFace();
            bottomFace.Contour.AddContourPoint(new ContourPoint(p8, null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(p8e, null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(correctedP2e, null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(correctedP2s, null));
            rebarSet.LegFaces.Add(bottomFace);

            var topFace = new RebarLegFace();
            topFace.Contour.AddContourPoint(new ContourPoint(correctedP8s, null));
            topFace.Contour.AddContourPoint(new ContourPoint(p2, null));
            topFace.Contour.AddContourPoint(new ContourPoint(p2e, null));
            topFace.Contour.AddContourPoint(new ContourPoint(correctedP8e, null));
            rebarSet.LegFaces.Add(topFace);

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing),
                SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = 0;
            guideline.Spacing.EndOffset = 0;

            Point startGL = new Point(p8.X, p2.Y, p8.Z);

            guideline.Curve.AddContourPoint(new ContourPoint(startGL, null));
            guideline.Curve.AddContourPoint(new ContourPoint(p2, null));
            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            var topEndDetailModifier = new RebarEndDetailModifier();
            topEndDetailModifier.Father = rebarSet;
            topEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            topEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = GetHookLength(rebarDiameter);
            topEndDetailModifier.Curve.AddContourPoint(new ContourPoint(correctedP8e, null));
            topEndDetailModifier.Curve.AddContourPoint(new ContourPoint(p2e, null));
            topEndDetailModifier.Insert();

            var bottomEndDetailModifier = new RebarEndDetailModifier();
            bottomEndDetailModifier.Father = rebarSet;
            bottomEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            bottomEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = GetHookLength(rebarDiameter);
            bottomEndDetailModifier.Curve.AddContourPoint(new ContourPoint(p8e, null));
            bottomEndDetailModifier.Curve.AddContourPoint(new ContourPoint(correctedP2e, null));
            bottomEndDetailModifier.Insert();

            rebarSet.SetUserProperty("__MIN_BAR_LENTYPE", 0);
            rebarSet.SetUserProperty("__MIN_BAR_LENGTH", 30 * rebarDiameter);

            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2, 2, 2 });
        }
        void ClosingLongitudianlRebarTop(int number)
        {
            string rebarSize = Program.ExcelDictionary["CLR_Diameter"];
            int rebarDiameter = Convert.ToInt32(rebarSize);
            string spacing = Program.ExcelDictionary["CLR_Spacing"];

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "ABT_CLR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 2;

            int secondNumber = 1;
            Point p1 = ProfilePoints[number][1];
            Point p2 = ProfilePoints[number][2];
            Point p3 = ProfilePoints[number][3];
            Point p4 = ProfilePoints[number][4];
            Point p5 = ProfilePoints[number][5];
            Point p7 = ProfilePoints[number][7];
            Point p8 = ProfilePoints[number][8];
            Point p9 = ProfilePoints[number][9];

            Point p2e = ProfilePoints[secondNumber][2];
            Point p3e = ProfilePoints[secondNumber][3];
            Point p4e = ProfilePoints[secondNumber][4];
            Point p5e = ProfilePoints[secondNumber][5];
            Point p7e = ProfilePoints[secondNumber][7];
            Point p8e = ProfilePoints[secondNumber][8];

            Vector xDirBottom = Utility.GetVectorFromTwoPoints(p8, p7);
            Vector yDirBottom = Utility.GetVectorFromTwoPoints(p8, p8e);
            GeometricPlane geometricPlaneBottom = new GeometricPlane(p8, xDirBottom, yDirBottom);

            Line startLine45 = new Line(p4, p5);
            Line endLine45 = new Line(p4e, p5e);
            Line startLine32 = new Line(p3, p2);
            Line endLine32 = new Line(p3e, p2e);

            Point correctedP5s = Utility.GetExtendedIntersection(startLine45, geometricPlaneBottom, 10);
            Point correctedP5e = Utility.GetExtendedIntersection(endLine45, geometricPlaneBottom, 10);
            Point correctedP2s = Utility.GetExtendedIntersection(startLine32, geometricPlaneBottom, 10);
            Point correctedP2e = Utility.GetExtendedIntersection(endLine32, geometricPlaneBottom, 10);

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(p3, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p4, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(correctedP5s, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(correctedP2s, null));
            rebarSet.LegFaces.Add(mainFace);

            var bottomFace = new RebarLegFace();
            bottomFace.Contour.AddContourPoint(new ContourPoint(p3, null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(p4, null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(p4e, null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(p3e, null));
            rebarSet.LegFaces.Add(bottomFace);

            var topFace = new RebarLegFace();
            topFace.Contour.AddContourPoint(new ContourPoint(correctedP2s, null));
            topFace.Contour.AddContourPoint(new ContourPoint(correctedP5s, null));
            topFace.Contour.AddContourPoint(new ContourPoint(correctedP5e, null));
            topFace.Contour.AddContourPoint(new ContourPoint(correctedP2e, null));
            rebarSet.LegFaces.Add(topFace);

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

            Point startGL = new Point(p3.X, p2.Y, p3.Z);
            Point endGL = new Point(p4.X, p2.Y, p4.Z);

            guideline.Curve.AddContourPoint(new ContourPoint(startGL, null));
            guideline.Curve.AddContourPoint(new ContourPoint(endGL, null));
            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            var topEndDetailModifier = new RebarEndDetailModifier();
            topEndDetailModifier.Father = rebarSet;
            topEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            topEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = GetHookLength(rebarDiameter);
            topEndDetailModifier.Curve.AddContourPoint(new ContourPoint(p3e, null));
            topEndDetailModifier.Curve.AddContourPoint(new ContourPoint(p4e, null));
            topEndDetailModifier.Insert();

            var bottomEndDetailModifier = new RebarEndDetailModifier();
            bottomEndDetailModifier.Father = rebarSet;
            bottomEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            bottomEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = GetHookLength(rebarDiameter);
            bottomEndDetailModifier.Curve.AddContourPoint(new ContourPoint(correctedP2e, null));
            bottomEndDetailModifier.Curve.AddContourPoint(new ContourPoint(correctedP5e, null));
            bottomEndDetailModifier.Insert();


            rebarSet.SetUserProperty("__MIN_BAR_LENTYPE", 0);
            rebarSet.SetUserProperty("__MIN_BAR_LENGTH", 30 * rebarDiameter);

            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2, 2, 2 });
        }
        void ClosingLongitudianlRebarTop2(int number)
        {
            string rebarSize = Program.ExcelDictionary["CLR_Diameter"];
            int rebarDiameter = Convert.ToInt32(rebarSize);
            string spacing = Program.ExcelDictionary["CLR_Spacing"];

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "ABT_CLR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 2;

            int secondNumber = 1;
            Point p2 = ProfilePoints[number][2];
            Point p3 = ProfilePoints[number][3];
            Point p4 = ProfilePoints[number][4];
            Point p5 = ProfilePoints[number][5];
            Point p6 = ProfilePoints[number][6];
            Point p7 = ProfilePoints[number][7];
            Point p8 = ProfilePoints[number][8];

            Point p2e = ProfilePoints[secondNumber][2];
            Point p3e = ProfilePoints[secondNumber][3];
            Point p4e = ProfilePoints[secondNumber][4];
            Point p5e = ProfilePoints[secondNumber][5];
            Point p6e = ProfilePoints[secondNumber][6];
            Point p7e = ProfilePoints[secondNumber][7];
            Point p8e = ProfilePoints[secondNumber][8];

            Vector xDirBottom = Utility.GetVectorFromTwoPoints(p8, p7);
            Vector yDirBottom = Utility.GetVectorFromTwoPoints(p8, p8e);
            GeometricPlane geometricPlaneBottom = new GeometricPlane(p8, xDirBottom, yDirBottom);

            Line startLine45 = new Line(p4, p5);
            Line endLine45 = new Line(p4e, p5e);
            Line startLine32 = new Line(p3, p2);
            Line endLine32 = new Line(p3e, p2e);

            Point correctedP5s = Utility.GetExtendedIntersection(startLine45, geometricPlaneBottom, 10);
            Point correctedP5e = Utility.GetExtendedIntersection(endLine45, geometricPlaneBottom, 10);
            Point correctedP2s = Utility.GetExtendedIntersection(startLine32, geometricPlaneBottom, 10);
            Point correctedP2e = Utility.GetExtendedIntersection(endLine32, geometricPlaneBottom, 10);

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(p5, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(correctedP5s, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p7, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p6, null));
            rebarSet.LegFaces.Add(mainFace);

            var bottomFace = new RebarLegFace();
            bottomFace.Contour.AddContourPoint(new ContourPoint(correctedP5s, null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(p7, null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(p7e, null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(correctedP5e, null));
            rebarSet.LegFaces.Add(bottomFace);

            var topFace = new RebarLegFace();
            topFace.Contour.AddContourPoint(new ContourPoint(p5, null));
            topFace.Contour.AddContourPoint(new ContourPoint(p6, null));
            topFace.Contour.AddContourPoint(new ContourPoint(p6e, null));
            topFace.Contour.AddContourPoint(new ContourPoint(p5e, null));
            rebarSet.LegFaces.Add(topFace);

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing),
                SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = 0;
            guideline.Spacing.EndOffset = 100;

            Point startGL = new Point(p5.X, p6.Y, p5.Z);

            guideline.Curve.AddContourPoint(new ContourPoint(startGL, null));
            guideline.Curve.AddContourPoint(new ContourPoint(p6, null));
            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            var topEndDetailModifier = new RebarEndDetailModifier();
            topEndDetailModifier.Father = rebarSet;
            topEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            topEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = GetHookLength(rebarDiameter);
            topEndDetailModifier.Curve.AddContourPoint(new ContourPoint(p5e, null));
            topEndDetailModifier.Curve.AddContourPoint(new ContourPoint(p6e, null));
            topEndDetailModifier.Insert();

            var bottomEndDetailModifier = new RebarEndDetailModifier();
            bottomEndDetailModifier.Father = rebarSet;
            bottomEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            bottomEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = GetHookLength(rebarDiameter);
            bottomEndDetailModifier.Curve.AddContourPoint(new ContourPoint(correctedP5e, null));
            bottomEndDetailModifier.Curve.AddContourPoint(new ContourPoint(p7e, null));
            bottomEndDetailModifier.Insert();
            rebarSet.SetUserProperty("__MIN_BAR_LENTYPE", 0);
            rebarSet.SetUserProperty("__MIN_BAR_LENGTH", 30 * rebarDiameter);

            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2, 2, 2 });
        }
        #endregion
        #region Fields
        public static double Height;
        public static double Width;
        public static double FrontHeight;
        public static double ShelfWidth;
        public static double ShelfHeight;
        public static double BackwallWidth;
        public static double CantileverWidth;
        public static double BackwallTopHeight;
        public static double CantileverHeight;
        public static double BackwallBottomHeight;
        public static double SkewHeight;
        public static double Height2;
        public static double FrontHeight2;
        public static double BackwallTopHeight2;
        public static double BackwallBottomHeight2;
        public static double Height3;
        public static double FrontHeight3;
        public static double BackwallTopHeight3;
        public static double BackwallBottomHeight3;
        public static double Length;
        public static double Length2;
        public static double FullWidth;
        #endregion
    }
}