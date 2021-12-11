﻿using System;
using System.Collections.Generic;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model;

namespace ZeroTouchTekla
{
    public class ABT : Element
    {

        /*
         *              4-----3
         *              |     |
         *              |     |
         *              |     |
         *         6----5     |
         *         |          |
         *         |          2--------1
         *         |                   |
         *         7                   |
         *          \                  |
         *           \                 |
         *            \                |
         *             8               |
         *             |               |
         *             |               |
         *             |               |
         *             |               |
         *             9---------------0
         *
         */
        public ABT(Beam part) : base(part)
        {
            GetProfilePointsAndParameters(part);
        }
        new public static void GetProfilePointsAndParameters(Beam beam)
        {
            string[] profileValues = GetProfileValues(beam);
            //ABT Width*Height*FrontHeight*ShelfHeight*ShelfWidth*BackwallWidth*CantileverWidth*BackwallTopHeight*CantileverHeight*BackwallBottomHeight*SkewHeight

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
            Height2 = 0;

            Point p0 = new Point(0, -Height / 2.0, FullWidth / 2.0);
            Point p1 = new Point(0, p0.Y + FrontHeight, p0.Z);
            Point p2 = new Point(0, p1.Y + ShelfHeight, p1.Z - ShelfWidth);
            Point p3 = new Point(0, Height / 2.0, p2.Z);
            Point p4 = new Point(0, Height / 2.0, p3.Z - BackwallWidth);
            Point p5 = new Point(0, p4.Y - BackwallTopHeight, p4.Z);
            Point p6 = new Point(0, p5.Y - CantileverHeight, p5.Z - CantileverWidth);
            Point p7 = new Point(0, p6.Y - BackwallBottomHeight, p6.Z);
            Point p8 = new Point(0, p7.Y - SkewHeight, FullWidth / 2.0 - Width);
            Point p9 = new Point(0, -Height / 2.0, FullWidth / 2.0 - Width);

            List<Point> firstProfile = new List<Point> { p0, p1, p2, p3, p4, p5, p6, p7, p8, p9 };

            List<Point> secondProfile = new List<Point>();
            if (Height2 == 0)
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

            }
            List<List<Point>> beamPoints = new List<List<Point>> { firstProfile, secondProfile };
            ProfilePoints = beamPoints;
            ElementFace = new ElementFace(ProfilePoints);
        }
        new public void Create()
        {
            OuterVerticalRebar();
            InnerVerticalRebar();
            CantileverVerticalRebar();
            BackwallTopVerticalRebar();
            BackwallOuterVerticalRebar();
            BackwallInnerVerticalRebar();
            ShelfHorizontalRebar();
            OuterLongitudinalRebar();
            InnerLongitudinalRebar();
            CantileverLongitudinalRebar(1);
            CantileverLongitudinalRebar(2);
            CantileverLongitudinalRebar(3);
            BackwallLongitudinalRebar(1);
            BackwallLongitudinalRebar(2);
            BackwallLongitudinalRebar(3);
        }
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
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
            rebarSet.LegFaces.Add(mainFace);

            Point offsetedStartPoint = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y, ProfilePoints[0][0].Z + 40 * Convert.ToInt32(rebarSize));
            Point offsetedEndPoint = new Point(ProfilePoints[1][0].X, ProfilePoints[1][0].Y, ProfilePoints[1][0].Z + 40 * Convert.ToInt32(rebarSize));

            var bottomFace = new RebarLegFace();
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
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
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();

            var innerEndDetailModifier = new RebarEndDetailModifier();
            innerEndDetailModifier.Father = rebarSet;
            innerEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            innerEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = 10 * Convert.ToInt32(rebarSize);
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
                Point endIntersection = new Point(ProfilePoints[1][0].X, ProfilePoints[1][0].Y + spliterOffset, ProfilePoints[1][0].Z);
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
                    propertyModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
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
            Point p2 = ProfilePoints[1][9];
            Point p3 = ProfilePoints[1][8];
            Point p4 = ProfilePoints[0][8];

            Point p3o = new Point(p3.X, p3.Y + 40 * rebarDiameter, p3.Z);
            Point p4o = new Point(p4.X, p4.Y + 40 * rebarDiameter, p4.Z);

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(p1, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p2, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p3o, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p4o, null));
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
            innerEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = 10 * Convert.ToInt32(rebarSize);
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
                    propertyModifier.Insert();
                }
                new Model().CommitChanges();
            }

            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 3 });
        }
        void CantileverVerticalRebar()
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

            Point p8s = ProfilePoints[0][8];
            Point p8e = ProfilePoints[1][8];
            Point p7s = ProfilePoints[0][7];
            Point p7e = ProfilePoints[1][7];
            Point p6s = ProfilePoints[0][6];
            Point p6e = ProfilePoints[1][6];
            Point p5s = ProfilePoints[0][5];
            Point p5e = ProfilePoints[1][5];
            Point p2s = ProfilePoints[0][2];
            Point p2e = ProfilePoints[1][2];

            Vector xAxis = Utility.GetVectorFromTwoPoints(ProfilePoints[0][2], ProfilePoints[1][2]).GetNormal();
            Vector yAxis = Utility.GetVectorFromTwoPoints(ProfilePoints[0][2], ProfilePoints[0][3]).GetNormal();
            GeometricPlane backwallPlane = new GeometricPlane(ProfilePoints[0][2], xAxis, yAxis);

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
            guideline.Spacing.StartOffset = 100;
            guideline.Spacing.EndOffset = 100;

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
        void BackwallTopVerticalRebar()
        {
            string rebarSize = Program.ExcelDictionary["BVR_Diameter"];
            int rebarDiameter = Convert.ToInt32(rebarSize);
            string spacing = Program.ExcelDictionary["BVR_Spacing"];

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "ABT_BVR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;


            Point p4s = ProfilePoints[0][4];
            Point p4e = ProfilePoints[1][4];
            Point p3s = ProfilePoints[0][3];
            Point p3e = ProfilePoints[1][3];
            Point p2s = ProfilePoints[0][2];
            Point p2e = ProfilePoints[1][2];
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
            guideline.Spacing.StartOffset = 100;
            guideline.Spacing.EndOffset = 100;

            guideline.Curve.AddContourPoint(new ContourPoint(p3s, null));
            guideline.Curve.AddContourPoint(new ContourPoint(p3e, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 1, 1 });
        }
        void BackwallOuterVerticalRebar()
        {
            string rebarSize = Program.ExcelDictionary["BVR_Diameter"];
            int rebarDiameter = Convert.ToInt32(rebarSize);
            string spacing = Program.ExcelDictionary["BVR_Spacing"];

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "ABT_BVR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            Point p8s = ProfilePoints[0][8];
            Point p8e = ProfilePoints[1][8];
            Point p7s = ProfilePoints[0][7];
            Point p7e = ProfilePoints[1][7];
            Point p5s = ProfilePoints[0][5];
            Point p5e = ProfilePoints[1][5];
            Point p4s = ProfilePoints[0][4];
            Point p4e = ProfilePoints[1][4];
            Point p3s = ProfilePoints[0][3];
            Point p3e = ProfilePoints[1][3];
            Point p2s = ProfilePoints[0][2];
            Point p2e = ProfilePoints[1][2];

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
            guideline.Spacing.StartOffset = 100;
            guideline.Spacing.EndOffset = 100;

            guideline.Curve.AddContourPoint(new ContourPoint(p3s, null));
            guideline.Curve.AddContourPoint(new ContourPoint(p3e, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            var bottomLengthModifier = new RebarEndDetailModifier();
            bottomLengthModifier.Father = rebarSet;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentLength = 10 * rebarDiameter;
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(p8s, null));
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(p8e, null));
            bottomLengthModifier.Insert();

            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 1, 1, 1, 1 });
        }
        void BackwallInnerVerticalRebar()
        {
            string rebarSize = Program.ExcelDictionary["BVR_Diameter"];
            int rebarDiameter = Convert.ToInt32(rebarSize);
            string spacing = Program.ExcelDictionary["BVR_Spacing"];

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "ABT_BVR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            Point p8s = ProfilePoints[0][8];
            Point p8e = ProfilePoints[1][8];
            Point p7s = ProfilePoints[0][7];
            Point p7e = ProfilePoints[1][7];
            Point p5s = ProfilePoints[0][5];
            Point p5e = ProfilePoints[1][5];
            Point p4s = ProfilePoints[0][4];
            Point p4e = ProfilePoints[1][4];
            Point p3s = ProfilePoints[0][3];
            Point p3e = ProfilePoints[1][3];
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
            guideline.Spacing.StartOffset = 100;
            guideline.Spacing.EndOffset = 100;

            guideline.Curve.AddContourPoint(new ContourPoint(p2s, null));
            guideline.Curve.AddContourPoint(new ContourPoint(p2e, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            var innerLengthModifier = new RebarEndDetailModifier();
            innerLengthModifier.Father = rebarSet;
            innerLengthModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            innerLengthModifier.RebarLengthAdjustment.AdjustmentLength = 10 * rebarDiameter;
            innerLengthModifier.Curve.AddContourPoint(new ContourPoint(innerStartPoint, null));
            innerLengthModifier.Curve.AddContourPoint(new ContourPoint(innerEndPoint, null));
            innerLengthModifier.Insert();

            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 1 });
        }
        void ShelfHorizontalRebar()
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

            Point p0s = ProfilePoints[0][0];
            Point p0e = ProfilePoints[1][0];
            Point p1s = ProfilePoints[0][1];
            Point p1e = ProfilePoints[1][1];
            Point p2s = ProfilePoints[0][2];
            Point p2e = ProfilePoints[1][2];
            Point p6s = ProfilePoints[0][6];
            Point p6e = ProfilePoints[1][6];
            Point p7s = ProfilePoints[0][7];
            Point p7e = ProfilePoints[1][7];

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
            guideline.Spacing.StartOffset = 100;
            guideline.Spacing.EndOffset = 100;

            guideline.Curve.AddContourPoint(new ContourPoint(p1s, null));
            guideline.Curve.AddContourPoint(new ContourPoint(p1e, null));
            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            var bottomLengthModifier = new RebarEndDetailModifier();
            bottomLengthModifier.Father = rebarSet;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentLength = 10 * rebarDiameter;
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(p0s, null));
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(p0e, null));
            bottomLengthModifier.Insert();

            var topLengthModifier = new RebarEndDetailModifier();
            topLengthModifier.Father = rebarSet;
            topLengthModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            topLengthModifier.RebarLengthAdjustment.AdjustmentLength = 10 * rebarDiameter;
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
            Point p0e = ProfilePoints[1][0];
            Point p1s = ProfilePoints[0][1];
            Point p1e = ProfilePoints[1][1];

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(p0s, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p0e, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p1e, null));
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

            guideline.Curve.AddContourPoint(new ContourPoint(p0s, null));
            guideline.Curve.AddContourPoint(new ContourPoint(p1s, null));

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
            Point p9e = ProfilePoints[1][9];
            Point p8s = ProfilePoints[0][8];
            Point p8e = ProfilePoints[1][8];

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(p9s, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p9e, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(p8e, null));
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

            guideline.Curve.AddContourPoint(new ContourPoint(p9s, null));
            guideline.Curve.AddContourPoint(new ContourPoint(p8s, null));

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
            Point p8e = ProfilePoints[1][8];
            Point p7s = ProfilePoints[0][7];
            Point p7e = ProfilePoints[1][7];
            Point p6s = ProfilePoints[0][6];
            Point p6e = ProfilePoints[1][6];
            Point p5s = ProfilePoints[0][5];
            Point p5e = ProfilePoints[1][5];

            Point startGL;
            Point endGL;
            if (number == 1)
            {
                var firstFace = new RebarLegFace();
                firstFace.Contour.AddContourPoint(new ContourPoint(p8s, null));
                firstFace.Contour.AddContourPoint(new ContourPoint(p8e, null));
                firstFace.Contour.AddContourPoint(new ContourPoint(p7e, null));
                firstFace.Contour.AddContourPoint(new ContourPoint(p7s, null));
                rebarSet.LegFaces.Add(firstFace);
                startGL = p8s;
                endGL = p7s;
            }
            else if (number == 2)
            {
                var secondFace = new RebarLegFace();
                secondFace.Contour.AddContourPoint(new ContourPoint(p7s, null));
                secondFace.Contour.AddContourPoint(new ContourPoint(p7e, null));
                secondFace.Contour.AddContourPoint(new ContourPoint(p6e, null));
                secondFace.Contour.AddContourPoint(new ContourPoint(p6s, null));
                rebarSet.LegFaces.Add(secondFace);
                startGL = p7s;
                endGL = p6s;
            }
            else
            {
                var thirdFace = new RebarLegFace();
                thirdFace.Contour.AddContourPoint(new ContourPoint(p6s, null));
                thirdFace.Contour.AddContourPoint(new ContourPoint(p6e, null));
                thirdFace.Contour.AddContourPoint(new ContourPoint(p5e, null));
                thirdFace.Contour.AddContourPoint(new ContourPoint(p5s, null));
                rebarSet.LegFaces.Add(thirdFace);
                startGL = p6s;
                endGL = p5s;
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
            Point p4e = ProfilePoints[1][4];
            Point p3s = ProfilePoints[0][3];
            Point p3e = ProfilePoints[1][3];
            Point p2s = ProfilePoints[0][2];
            Point p2e = ProfilePoints[1][2];
            Point p5s = ProfilePoints[0][5];
            Point p5e = ProfilePoints[1][5];

            Point startGL;
            Point endGL;

            if (number == 1)
            {
                var firstFace = new RebarLegFace();
                firstFace.Contour.AddContourPoint(new ContourPoint(p2s, null));
                firstFace.Contour.AddContourPoint(new ContourPoint(p2e, null));
                firstFace.Contour.AddContourPoint(new ContourPoint(p3e, null));
                firstFace.Contour.AddContourPoint(new ContourPoint(p3s, null));
                rebarSet.LegFaces.Add(firstFace);
                startGL = p2s;
                endGL = p3s;
            }
            else if (number == 2)
            {
                var secondFace = new RebarLegFace();
                secondFace.Contour.AddContourPoint(new ContourPoint(p3s, null));
                secondFace.Contour.AddContourPoint(new ContourPoint(p3e, null));
                secondFace.Contour.AddContourPoint(new ContourPoint(p4e, null));
                secondFace.Contour.AddContourPoint(new ContourPoint(p4s, null));
                rebarSet.LegFaces.Add(secondFace);
                startGL = p3s;
                endGL = p4s;
            }
            else
            {
                var thirdFace = new RebarLegFace();
                thirdFace.Contour.AddContourPoint(new ContourPoint(p4s, null));
                thirdFace.Contour.AddContourPoint(new ContourPoint(p4e, null));
                thirdFace.Contour.AddContourPoint(new ContourPoint(p5e, null));
                thirdFace.Contour.AddContourPoint(new ContourPoint(p5s, null));
                rebarSet.LegFaces.Add(thirdFace);
                startGL = p4s;
                endGL = p5s;
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
            rebarSet.LayerOrderNumber = 1;

            Point p0s = ProfilePoints[0][0];
            Point p0e = ProfilePoints[1][0];
            Point p1s = ProfilePoints[0][1];
            Point p1e = ProfilePoints[1][1];
            Point p2s = ProfilePoints[0][2];
            Point p2e = ProfilePoints[1][2];
            Point p6s = ProfilePoints[0][6];
            Point p6e = ProfilePoints[1][6];
            Point p7s = ProfilePoints[0][7];
            Point p7e = ProfilePoints[1][7];

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
            guideline.Spacing.StartOffset = 100;
            guideline.Spacing.EndOffset = 100;

            guideline.Curve.AddContourPoint(new ContourPoint(p1s, null));
            guideline.Curve.AddContourPoint(new ContourPoint(p1e, null));
            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            var bottomLengthModifier = new RebarEndDetailModifier();
            bottomLengthModifier.Father = rebarSet;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentLength = 10 * rebarDiameter;
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(p0s, null));
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(p0e, null));
            bottomLengthModifier.Insert();

            var topLengthModifier = new RebarEndDetailModifier();
            topLengthModifier.Father = rebarSet;
            topLengthModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            topLengthModifier.RebarLengthAdjustment.AdjustmentLength = 10 * rebarDiameter;
            topLengthModifier.Curve.AddContourPoint(new ContourPoint(p6s, null));
            topLengthModifier.Curve.AddContourPoint(new ContourPoint(p6e, null));
            topLengthModifier.Insert();

            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 1, 1 });
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
        public static double Length;
        public static double FullWidth;
        #endregion
    }
}
