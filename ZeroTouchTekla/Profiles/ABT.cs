using System;
using System.Collections.Generic;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model;
using System.Linq;
using System.Reflection;

namespace ZeroTouchTekla.Profiles
{
    public class ABT : Element
    {
        #region Constructor
        public ABT(List<Part> parts) : base(parts)
        {
            SetLocalPlane(parts[0]);
            GetProfilePointsAndParameters(parts);
        }
        new public void Create()
        {
            OuterVerticalRebar();
            InnerVerticalRebar();
            for (int i = 0; i < ProfilePoints.Count - 1; i++)
            {
                CantileverVerticalRebar(i);
                BackwallTopVerticalRebar(i);
                BackwallOuterVerticalRebar(i);
                BackwallInnerVerticalRebar(i);
                ShelfHorizontalRebar(i);

            }
            CantileverLongitudinalRebar(0);
            CantileverLongitudinalRebar(1);
            CantileverLongitudinalRebar(2);
            OuterLongitudinalRebar();
            InnerLongitudinalRebar();
            BackwallLongitudinalRebar(1);
            BackwallLongitudinalRebar(2);
            BackwallLongitudinalTopRebar();
            ShelfLongitudinalRebar();
            ClosingCShapeRebarBottom(0);
            ClosingCShapeRebarBottom(ProfilePoints.Count - 1);
            ClosingCShapeRebarMid(0);
            ClosingCShapeRebarMid(ProfilePoints.Count - 1);
            ClosingCShapeRebarTop(0);
            ClosingCShapeRebarTop(ProfilePoints.Count - 1);
            CShapeRebarCommon();
            ClosingLongitudianlRebarBottom(0);
            ClosingLongitudianlRebarBottom(ProfilePoints.Count - 1);
            ClosingLongitudianlRebarMid(0);
            ClosingLongitudianlRebarMid(ProfilePoints.Count - 1);
            ClosingLongitudianlRebarTop(0);
            ClosingLongitudianlRebarTop(ProfilePoints.Count - 1);
            ClosingLongitudianlRebarTop2(0);
            ClosingLongitudianlRebarTop2(ProfilePoints.Count - 1);
        }
        #endregion
        #region PrivateMethods  
        public void GetProfilePointsAndParameters(List<Part> parts)
        {
            //ABT Width*Height*FrontHeight*ShelfHeight*ShelfWidth*BackwallWidth*CantileverWidth*BackwallTopHeight*CantileverHeight*BackwallBottomHeight*SkewHeight
            //ABTV W*H*FH*SH*SW*BWW*CW*BWTH*CH*BWBH*SH*H*FH*BWTH*BWBH
            //ABTVSK W*H*FH*SH*SW*BWW*CW*BWTH*CH*BWBH*SH*H*FH*BWTH*BWBH*HO

            FirstBeamProperties(parts[0]);
            for (int i = 1; i < parts.Count; i++)
            {
                NextBeamProperties(parts[i], i);
            }
        }
        public void FirstBeamProperties(Part part)
        {
            Beam beam = part as Beam;
            string[] profileValues = GetProfileValues(beam);
            Width = Convert.ToDouble(profileValues[0]);
            Height.Add(Convert.ToDouble(profileValues[1]));
            FrontHeight.Add(Convert.ToDouble(profileValues[2]));
            ShelfHeight = Convert.ToDouble(profileValues[3]);
            ShelfWidth = Convert.ToDouble(profileValues[4]);
            BackwallWidth = Convert.ToDouble(profileValues[5]);
            CantileverWidth = Convert.ToDouble(profileValues[6]);
            BackwallTopHeight.Add(Convert.ToDouble(profileValues[7]));
            CantileverHeight = Convert.ToDouble(profileValues[8]);
            BackwallBottomHeight.Add(Convert.ToDouble(profileValues[9]));
            SkewHeight = Convert.ToDouble(profileValues[10]);
            FullWidth = ShelfWidth + BackwallWidth + CantileverWidth;
            Length.Add(Distance.PointToPoint(beam.StartPoint, beam.EndPoint));
            FullLength += Length.Last();
            HorizontalOffset = 0;
            string firstProfileName = beam.Profile.ProfileString;

            if (firstProfileName.Contains("V"))
            {
                Height.Add(Convert.ToDouble(profileValues[11]));
                FrontHeight.Add(Convert.ToDouble(profileValues[12]));
                BackwallTopHeight.Add(Convert.ToDouble(profileValues[13]));
                BackwallBottomHeight.Add(Convert.ToDouble(profileValues[14]));
                HorizontalOffset = Convert.ToDouble(profileValues[15]);
            }
            else
            {
                Height.Add(Height[0]);
                FrontHeight.Add(FrontHeight[0]);
                BackwallTopHeight.Add(BackwallTopHeight.Last());
                BackwallBottomHeight.Add(BackwallBottomHeight[0]);
                HorizontalOffset = Convert.ToDouble(profileValues[11]);
            }

            double distanceToMid = Height[0] > Height[1] ? Height[0] / 2.0 : Height[1] / 2.0;

            Point p0 = new Point(0, -distanceToMid, FullWidth / 2.0);
            Point p1 = new Point(0, p0.Y + FrontHeight[0], p0.Z);
            Point p2 = new Point(0, p1.Y + ShelfHeight, p1.Z - ShelfWidth);
            Point p3 = new Point(0, -distanceToMid + Height[0], p2.Z);
            Point p4 = new Point(0, -distanceToMid + Height[0], p3.Z - BackwallWidth);
            Point p5 = new Point(0, p4.Y - BackwallTopHeight[0], p4.Z);
            Point p6 = new Point(0, p5.Y - CantileverHeight, p5.Z - CantileverWidth);
            Point p7 = new Point(0, p6.Y - BackwallBottomHeight[0], p6.Z);
            Point p8 = new Point(0, p7.Y - SkewHeight, FullWidth / 2.0 - Width);
            Point p9 = new Point(0, -distanceToMid, FullWidth / 2.0 - Width);
            List<Point> firstProfile = new List<Point> { p0, p1, p2, p3, p4, p5, p6, p7, p8, p9 };

            Point n0 = new Point(Length.Last(), -distanceToMid, FullWidth / 2.0);
            Point n1 = new Point(Length.Last(), n0.Y + FrontHeight[1], n0.Z);
            Point n2 = new Point(Length.Last(), n1.Y + ShelfHeight, n1.Z - ShelfWidth);
            Point n3 = new Point(Length.Last(), -distanceToMid + Height[1], n2.Z);
            Point n4 = new Point(Length.Last(), n3.Y, n3.Z - BackwallWidth);
            Point n5 = new Point(Length.Last(), n4.Y - BackwallTopHeight[1], n4.Z);
            Point n6 = new Point(Length.Last(), n5.Y - CantileverHeight, n5.Z - CantileverWidth);
            Point n7 = new Point(Length.Last(), n6.Y - BackwallBottomHeight[1], n6.Z);
            Point n8 = new Point(Length.Last(), n7.Y - SkewHeight, FullWidth / 2.0 - Width);
            Point n9 = new Point(Length.Last(), -distanceToMid, FullWidth / 2.0 - Width);
            List<Point> secondProfile = new List<Point> { n0, n1, n2, n3, n4, n5, n6, n7, n8, n9 };

            if (HorizontalOffset != 0)
            {
                foreach (Point p in firstProfile)
                {
                    p.Translate(0, 0, HorizontalOffset / 2.0);
                }
                foreach (Point p in secondProfile)
                {
                    p.Translate(0, 0, -HorizontalOffset / 2.0);
                }
            }

            ProfilePoints.Add(firstProfile);
            ProfilePoints.Add(secondProfile);
        }
        public void NextBeamProperties(Part part, int partNumber)
        {
            Beam nextBeam = part as Beam;
            string[] secondProfileValues = GetProfileValues(nextBeam);
            string nextProfileName = nextBeam.Profile.ProfileString;
            Length.Add(Distance.PointToPoint(nextBeam.StartPoint, nextBeam.EndPoint));
            FullLength += Length.Last();

            if (nextProfileName.Contains("V"))
            {
                Height.Add(Convert.ToDouble(secondProfileValues[11]));
                FrontHeight.Add(Convert.ToDouble(secondProfileValues[12]));
                BackwallTopHeight.Add(Convert.ToDouble(secondProfileValues[13]));
                BackwallBottomHeight.Add(Convert.ToDouble(secondProfileValues[14]));
            }
            else
            {
                Height.Add(Height.Last());
                FrontHeight.Add(FrontHeight.Last());
                BackwallTopHeight.Add(BackwallTopHeight.Last());
                BackwallBottomHeight.Add(BackwallBottomHeight.Last());
            }

            double distanceToMid = Height[0] > Height[1] ? Height[0] / 2.0 : Height[1] / 2.0;

            Point n0 = new Point(FullLength, -distanceToMid, FullWidth / 2.0);
            Point n1 = new Point(FullLength, n0.Y + FrontHeight.Last(), n0.Z);
            Point n2 = new Point(FullLength, n1.Y + ShelfHeight, n1.Z - ShelfWidth);
            Point n3 = new Point(FullLength, -distanceToMid + Height.Last(), n2.Z);
            Point n4 = new Point(FullLength, n3.Y, n3.Z - BackwallWidth);
            Point n5 = new Point(FullLength, n4.Y - BackwallTopHeight.Last(), n4.Z);
            Point n6 = new Point(FullLength, n5.Y - CantileverHeight, n5.Z - CantileverWidth);
            Point n7 = new Point(FullLength, n6.Y - BackwallBottomHeight.Last(), n6.Z);
            Point n8 = new Point(FullLength, n7.Y - SkewHeight, FullWidth / 2.0 - Width);
            Point n9 = new Point(FullLength, -distanceToMid, FullWidth / 2.0 - Width);
            List<Point> nextProfile = new List<Point> { n0, n1, n2, n3, n4, n5, n6, n7, n8, n9 };

            if (HorizontalOffset != 0)
            {
                foreach (Point p in nextProfile)
                {
                    p.Translate(0, 0, -HorizontalOffset * (partNumber + 0.5));
                }
            }
            ProfilePoints.Add(nextProfile);
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

            RebarSet rebarSet = InitializeRebarSet("ABT_OVR",rebarSize);

            //Itarate throught ProfilePoints
            int profilePointsMax = ProfilePoints.Count - 1;
            var mainFace = new RebarLegFace();
            for (int i = 0; i < ProfilePoints.Count; i++)
            {
                mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[i][0], null));
            }
            for (int i = ProfilePoints.Count - 1; i > -1; i--)
            {
                mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[i][1], null));
            }
            rebarSet.LegFaces.Add(mainFace);

            Point offsetedStartPoint = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y, ProfilePoints[0][0].Z + 40 * Convert.ToInt32(rebarSize));
            Point offsetedEndPoint = new Point(ProfilePoints[profilePointsMax][0].X, ProfilePoints[profilePointsMax][0].Y, ProfilePoints[profilePointsMax][0].Z + 40 * Convert.ToInt32(rebarSize));

            var bottomFace = new RebarLegFace();
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[profilePointsMax][0], null));
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
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[profilePointsMax][0], null));

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
                Point endIntersection = new Point(ProfilePoints[profilePointsMax][0].X, ProfilePoints[profilePointsMax][0].Y + spliterOffset, ProfilePoints[profilePointsMax][0].Z);
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
                    propertyModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[profilePointsMax][1], null));
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
            int profilePointsMax = ProfilePoints.Count - 1;

            var mainFace = new RebarLegFace();
            for (int i = 0; i < ProfilePoints.Count; i++)
            {
                mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[i][9], null));
            }
            for (int i = ProfilePoints.Count - 1; i > -1; i--)
            {
                mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[i][8], null));
            }
            rebarSet.LegFaces.Add(mainFace);

            Point p1 = ProfilePoints[0][9];
            Point p2 = ProfilePoints[profilePointsMax][9];
            Point p1o = new Point(ProfilePoints[0][9].X, ProfilePoints[0][9].Y, ProfilePoints[0][9].Z - 40 * Convert.ToInt32(rebarSize));
            Point p2o = new Point(ProfilePoints[profilePointsMax][9].X, ProfilePoints[profilePointsMax][9].Y, ProfilePoints[profilePointsMax][9].Z - 40 * Convert.ToInt32(rebarSize));

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
                    for (int i = ProfilePoints.Count - 1; i > -1; i--)
                    {
                        propertyModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[i][8], null));
                    }
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

            int f = number;
            int s = number + 1;

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

            Vector normal = Utility.GetVectorFromTwoPoints(ProfilePoints[0][0], ProfilePoints[1][0]).GetNormal();
            GeometricPlane correctedEndPlane = new GeometricPlane(p8e, normal);
            GeometricPlane correctedStartPlane = new GeometricPlane(p8s, normal);
            Line line8 = new Line(p8s, p8e);
            Line line7 = new Line(p7s, p7e);
            Line line6 = new Line(p6s, p6e);
            Line line5 = new Line(p5s, p5e);
            Line line3 = new Line(p3s, p3e);
            Line line2 = new Line(p2s, p2e);
            p8e = Utility.GetExtendedIntersection(line8, correctedEndPlane, 2);
            p7e = Utility.GetExtendedIntersection(line7, correctedEndPlane, 2);
            p6e = Utility.GetExtendedIntersection(line6, correctedEndPlane, 2);
            p5e = Utility.GetExtendedIntersection(line5, correctedEndPlane, 2);
            p3e = Utility.GetExtendedIntersection(line3, correctedEndPlane, 2);
            p2e = Utility.GetExtendedIntersection(line2, correctedEndPlane, 2);
            p8s = Utility.GetExtendedIntersection(line8, correctedStartPlane, 2);
            p7s = Utility.GetExtendedIntersection(line7, correctedStartPlane, 2);
            p6s = Utility.GetExtendedIntersection(line6, correctedStartPlane, 2);
            p5s = Utility.GetExtendedIntersection(line5, correctedStartPlane, 2);
            p3s = Utility.GetExtendedIntersection(line3, correctedStartPlane, 2);
            p2s = Utility.GetExtendedIntersection(line2, correctedStartPlane, 2);

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

            var guideline = GetPresetGuideline(spacing, number);

            Point correctedP8e = new Point(p8e.X, p8s.Y, p8e.Z);
            guideline.Curve.AddContourPoint(new ContourPoint(p8s, null));
            guideline.Curve.AddContourPoint(new ContourPoint(correctedP8e, null));

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

            int f = number;
            int s = number + 1;

            Point p4s = ProfilePoints[f][4];
            Point p4e = ProfilePoints[s][4];
            Point p3s = ProfilePoints[f][3];
            Point p3e = ProfilePoints[s][3];
            Point p2s = ProfilePoints[f][2];
            Point p2e = ProfilePoints[s][2];
            Point p5s = new Point(p4s.X, p2s.Y, p4s.Z);
            Point p5e = new Point(p4e.X, p2e.Y, p4e.Z);

            Vector normal = Utility.GetVectorFromTwoPoints(ProfilePoints[0][0], ProfilePoints[1][0]).GetNormal();
            GeometricPlane correctedEndPlane = new GeometricPlane(p3e, normal);
            GeometricPlane correctedStartPlane = new GeometricPlane(p3s, normal);
            Line line5 = new Line(p5s, p5e);
            Line line4 = new Line(p4s, p4e);
            Line line3 = new Line(p3s, p3e);
            Line line2 = new Line(p2s, p2e);
            p5e = Utility.GetExtendedIntersection(line5, correctedEndPlane, 2);
            p4e = Utility.GetExtendedIntersection(line4, correctedEndPlane, 2);
            p3e = Utility.GetExtendedIntersection(line3, correctedEndPlane, 2);
            p2e = Utility.GetExtendedIntersection(line2, correctedEndPlane, 2);
            p5s = Utility.GetExtendedIntersection(line5, correctedStartPlane, 2);
            p4s = Utility.GetExtendedIntersection(line4, correctedStartPlane, 2);
            p3s = Utility.GetExtendedIntersection(line3, correctedStartPlane, 2);
            p2s = Utility.GetExtendedIntersection(line2, correctedStartPlane, 2);

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

            var guideline = GetPresetGuideline(spacing, number);
            Point correctedP3 = new Point(p3e.X, p3s.Y, p3e.Z);
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

            int f = number;
            int s = number + 1;

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

            Point startTopPoint = new Point(p4s.X, p5s.Y, p4s.Z);
            Point endTopPoint = new Point(p4e.X, p5e.Y, p4e.Z);

            Vector xAxis = Utility.GetVectorFromTwoPoints(p7s, p8s).GetNormal();
            Vector yAxis = Utility.GetVectorFromTwoPoints(p7s, p7e).GetNormal();
            GeometricPlane cornicePlane = new GeometricPlane(p7s, xAxis, yAxis);

            Line innerStartLine = new Line(p4s, p5s);
            Line innerEndLine = new Line(p4e, p5e);
            Point innerStartPoint = Utility.GetExtendedIntersection(innerStartLine, cornicePlane, 5);
            Point innerEndPoint = Utility.GetExtendedIntersection(innerEndLine, cornicePlane, 5);

            Vector normal = Utility.GetVectorFromTwoPoints(ProfilePoints[0][0], ProfilePoints[1][0]).GetNormal();
            GeometricPlane correctedEndPlane = new GeometricPlane(p4e, normal);
            GeometricPlane correctedStartPlane = new GeometricPlane(p4s, normal);
            Line topLine = new Line(startTopPoint, endTopPoint);
            Line innerLine = new Line(innerStartPoint, innerEndPoint);
            Line line8 = new Line(p8s, p8e);
            endTopPoint = Utility.GetExtendedIntersection(topLine, correctedEndPlane, 2);
            innerEndPoint = Utility.GetExtendedIntersection(innerLine, correctedEndPlane, 2);
            p8e = Utility.GetExtendedIntersection(line8, correctedEndPlane, 2);
            startTopPoint = Utility.GetExtendedIntersection(topLine, correctedStartPlane, 2);
            innerStartPoint = Utility.GetExtendedIntersection(innerLine, correctedStartPlane, 2);
            p8s = Utility.GetExtendedIntersection(line8, correctedStartPlane, 2);

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

            var guideline = GetPresetGuideline(spacing, number);
            Point correctedP4 = new Point(p4e.X, p4s.Y, p4e.Z);
            guideline.Curve.AddContourPoint(new ContourPoint(p4s, null));
            guideline.Curve.AddContourPoint(new ContourPoint(correctedP4, null));

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

            int f = number;
            int s = number + 1;

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

            Line outerStartLine = new Line(p3s, p2s);
            Line outerEndLine = new Line(p3e, p2e);

            Point outerStartPoint = Utility.GetExtendedIntersection(outerStartLine, cornicePlane, 5);
            Point outerEndPoint = Utility.GetExtendedIntersection(outerEndLine, cornicePlane, 5);

            Vector normal = Utility.GetVectorFromTwoPoints(ProfilePoints[0][0], ProfilePoints[1][0]).GetNormal();
            GeometricPlane correctedEndPlane = new GeometricPlane(p4e, normal);
            GeometricPlane correctedStartPlane = new GeometricPlane(p4s, normal);
            Line outerLine = new Line(outerStartPoint, outerEndPoint);
            Line line8 = new Line(p8s, p8e);
            Line line2 = new Line(p2s, p2e);
            outerEndPoint = Utility.GetExtendedIntersection(outerLine, correctedEndPlane, 2);
            p8e = Utility.GetExtendedIntersection(line8, correctedEndPlane, 2);
            p2e = Utility.GetExtendedIntersection(line2, correctedEndPlane, 5);
            outerStartPoint = Utility.GetExtendedIntersection(outerLine, correctedStartPlane, 2);
            p8s = Utility.GetExtendedIntersection(line8, correctedStartPlane, 2);
            p2s = Utility.GetExtendedIntersection(line2, correctedStartPlane, 5);

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

            var guideline = GetPresetGuideline(spacing, number);
            Point correctedP2 = new Point(p2e.X, p2s.Y, p2e.Z);
            guideline.Curve.AddContourPoint(new ContourPoint(p2s, null));
            guideline.Curve.AddContourPoint(new ContourPoint(correctedP2, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            var innerLengthModifier = new RebarEndDetailModifier();
            innerLengthModifier.Father = rebarSet;
            innerLengthModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            innerLengthModifier.RebarLengthAdjustment.AdjustmentLength = GetHookLength(rebarDiameter);
            innerLengthModifier.Curve.AddContourPoint(new ContourPoint(outerStartPoint, null));
            innerLengthModifier.Curve.AddContourPoint(new ContourPoint(outerEndPoint, null));
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

            int f = number;
            int s = number + 1; ;

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

            Vector normal = Utility.GetVectorFromTwoPoints(ProfilePoints[0][0], ProfilePoints[1][0]).GetNormal();
            GeometricPlane correctedEndPlane = new GeometricPlane(ProfilePoints[s][3], normal);
            GeometricPlane correctedStartPlane = new GeometricPlane(ProfilePoints[f][3], normal);
            Line line0 = new Line(p0s, p0e);
            Line line1 = new Line(p1s, p1e);
            Line line6 = new Line(p6s, p6e);
            Line intersectionLine = new Line(startIntersection, endIntersection);
            p0e = Utility.GetExtendedIntersection(line0, correctedEndPlane, 2);
            p1e = Utility.GetExtendedIntersection(line1, correctedEndPlane, 2);
            p6e = Utility.GetExtendedIntersection(line6, correctedEndPlane, 2);
            endIntersection = Utility.GetExtendedIntersection(intersectionLine, correctedEndPlane, 2);
            p0s = Utility.GetExtendedIntersection(line0, correctedStartPlane, 2);
            p1s = Utility.GetExtendedIntersection(line1, correctedStartPlane, 2);
            p6s = Utility.GetExtendedIntersection(line6, correctedStartPlane, 2);
            startIntersection = Utility.GetExtendedIntersection(intersectionLine, correctedStartPlane, 2);

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

            var guideline = GetPresetGuideline(spacing, number);

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

            int profilePointsMax = ProfilePoints.Count - 1;
            Point p0e = ProfilePoints[profilePointsMax][0];

            var mainFace = new RebarLegFace();
            double minY = 0;
            double maxY = 0;

            for (int i = 0; i < ProfilePoints.Count; i++)
            {
                Point p = ProfilePoints[i][0];
                mainFace.Contour.AddContourPoint(new ContourPoint(p, null));
                if (minY > p.Y)
                {
                    minY = p.Y;
                }
            }
            for (int i = ProfilePoints.Count - 1; i > -1; i--)
            {
                Point p = ProfilePoints[i][1];
                mainFace.Contour.AddContourPoint(new ContourPoint(p, null));
                if (maxY < p.Y)
                {
                    maxY = p.Y;
                }
            }
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

            Point startGL = new Point(ProfilePoints[0][0].X, minY, ProfilePoints[0][0].Z);
            Point endGL = new Point(ProfilePoints[0][1].X, maxY, ProfilePoints[0][1].Z);

            guideline.Curve.AddContourPoint(new ContourPoint(startGL, null));
            guideline.Curve.AddContourPoint(new ContourPoint(endGL, null));
            rebarSet.Guidelines.Add(guideline);

            bool succes = rebarSet.Insert();
            new Model().CommitChanges();

            Point correctedP0e = new Point(p0e.X, p0e.Y + startOffset + secondDiameterLength, p0e.Z);
            var innerEndDetailModifier = new RebarPropertyModifier();
            innerEndDetailModifier.Father = rebarSet;
            innerEndDetailModifier.BarsAffected = BaseRebarModifier.BarsAffectedEnum.ALL_BARS;
            innerEndDetailModifier.RebarProperties.Size = secondRebarSize;
            innerEndDetailModifier.RebarProperties.Class = SetClass(Convert.ToDouble(secondRebarSize));
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(p0e, null));
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(correctedP0e, null));
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
            int profilePointsMax = ProfilePoints.Count - 1;

            Point p9e = ProfilePoints[profilePointsMax][9];
            double minY = 0;
            double maxY = 0;
            var mainFace = new RebarLegFace();
            for (int i = 0; i < ProfilePoints.Count; i++)
            {
                Point p = ProfilePoints[i][9];
                mainFace.Contour.AddContourPoint(new ContourPoint(p, null));
                if (minY > p.Y)
                {
                    minY = p.Y;
                }
            }
            for (int i = ProfilePoints.Count - 1; i > -1; i--)
            {
                Point p = ProfilePoints[i][8];
                mainFace.Contour.AddContourPoint(new ContourPoint(p, null));
                if (maxY < p.Y)
                {
                    maxY = p.Y;
                }
            }
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

            Point startGL = new Point(ProfilePoints[0][8].X, minY, ProfilePoints[0][8].Z);
            Point endGL = new Point(ProfilePoints[0][9].X, maxY, ProfilePoints[0][9].Z);

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

            Point startGL, endGL;
            if (number == 1)
            {
                for (int i = 0; i < ProfilePoints.Count - 1; i++)
                {
                    Point p1 = ProfilePoints[i][8];
                    Point p2 = ProfilePoints[i + 1][8];
                    Point p3 = ProfilePoints[i + 1][7];
                    Point p4 = ProfilePoints[i][7];
                    var face = new RebarLegFace();
                    face.Contour.AddContourPoint(new ContourPoint(p1, null));
                    face.Contour.AddContourPoint(new ContourPoint(p2, null));
                    face.Contour.AddContourPoint(new ContourPoint(p3, null));
                    face.Contour.AddContourPoint(new ContourPoint(p4, null));
                    rebarSet.LegFaces.Add(face);
                }

                Vector normal = Utility.GetVectorFromTwoPoints(ProfilePoints[0][8], ProfilePoints[1][8]).GetNormal();
                GeometricPlane geometricPlane = new GeometricPlane(ProfilePoints[0][8], normal);

                Line firstLine = new Line(ProfilePoints[0][8], ProfilePoints[1][8]);
                Line secondLine = new Line(ProfilePoints[0][7], ProfilePoints[1][7]);

                startGL = Utility.GetExtendedIntersection(firstLine, geometricPlane, 2);
                endGL = Utility.GetExtendedIntersection(secondLine, geometricPlane, 2);

            }
            else if (number == 2)
            {
                List<Point> list6 = new List<Point>();
                List<Point> list7 = new List<Point>();

                for (int i = 0; i < ProfilePoints.Count - 1; i++)
                {
                    Point p1 = ProfilePoints[i][6];
                    Point p2 = ProfilePoints[i + 1][6];
                    Point p3 = ProfilePoints[i + 1][7];
                    Point p4 = ProfilePoints[i][7];
                    p1 = Utility.Translate(p1, new Vector(0, -50, 0));
                    p2 = Utility.Translate(p2, new Vector(0, -50, 0));
                    var face = new RebarLegFace();
                    face.Contour.AddContourPoint(new ContourPoint(p1, null));
                    face.Contour.AddContourPoint(new ContourPoint(p2, null));
                    face.Contour.AddContourPoint(new ContourPoint(p3, null));
                    face.Contour.AddContourPoint(new ContourPoint(p4, null));
                    rebarSet.LegFaces.Add(face);
                    list6.Add(p1);
                    list6.Add(p2);
                    list7.Add(p3);
                    list7.Add(p4);
                }

                double maxZ = (from Point p in list6
                               orderby p.Y ascending
                               select p).LastOrDefault().Y;
                double minZ = (from Point p in list7
                               orderby p.Y ascending
                               select p).FirstOrDefault().Y;

                startGL = new Point(ProfilePoints[0][6].X, maxZ, ProfilePoints[0][6].Z);
                endGL = new Point(ProfilePoints[0][7].X, minZ, ProfilePoints[0][7].Z);
            }
            else
            {

                for (int i = 0; i < ProfilePoints.Count - 1; i++)
                {
                    Point p1 = ProfilePoints[i][6];
                    Point p2 = ProfilePoints[i + 1][6];
                    Point p3 = ProfilePoints[i + 1][5];
                    Point p4 = ProfilePoints[i][5];
                    var face = new RebarLegFace();
                    face.Contour.AddContourPoint(new ContourPoint(p1, null));
                    face.Contour.AddContourPoint(new ContourPoint(p2, null));
                    face.Contour.AddContourPoint(new ContourPoint(p3, null));
                    face.Contour.AddContourPoint(new ContourPoint(p4, null));
                    rebarSet.LegFaces.Add(face);
                }

                Vector normal = Utility.GetVectorFromTwoPoints(ProfilePoints[0][0], ProfilePoints[1][0]).GetNormal();
                GeometricPlane geometricPlane = new GeometricPlane(ProfilePoints[0][5], normal);
                Line firstLine = new Line(ProfilePoints[0][6], ProfilePoints[1][6]);
                Line secondLine = new Line(ProfilePoints[0][5], ProfilePoints[1][5]);

                startGL = Utility.GetExtendedIntersection(firstLine, geometricPlane, 2);
                endGL = Utility.GetExtendedIntersection(secondLine, geometricPlane, 2);
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
            guideline.Spacing.EndOffset = number == 3 ? 0 : 50;

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

            Point startGL;
            Point endGL;

            if (number == 1)
            {
                List<Point> list2 = new List<Point>();
                List<Point> list3 = new List<Point>();
                for (int i = 0; i < ProfilePoints.Count - 1; i++)
                {
                    Point p1 = ProfilePoints[i][2];
                    Point p2 = ProfilePoints[i + 1][2];
                    Point p3 = ProfilePoints[i + 1][3];
                    Point p4 = ProfilePoints[i][3];
                    var face = new RebarLegFace();
                    face.Contour.AddContourPoint(new ContourPoint(p1, null));
                    face.Contour.AddContourPoint(new ContourPoint(p2, null));
                    face.Contour.AddContourPoint(new ContourPoint(p3, null));
                    face.Contour.AddContourPoint(new ContourPoint(p4, null));
                    rebarSet.LegFaces.Add(face);
                    list2.Add(p1);
                    list2.Add(p2);
                    list3.Add(p3);
                    list3.Add(p4);
                }

                double maxZ = (from Point p in list3
                               orderby p.Y ascending
                               select p).LastOrDefault().Y;
                double minZ = (from Point p in list2
                               orderby p.Y ascending
                               select p).FirstOrDefault().Y;

                startGL = new Point(ProfilePoints[0][3].X, maxZ, ProfilePoints[0][3].Z);
                endGL = new Point(ProfilePoints[0][2].X, minZ, ProfilePoints[0][2].Z);
            }
            else
            {
                List<Point> list4 = new List<Point>();
                List<Point> list5 = new List<Point>();
                for (int i = 0; i < ProfilePoints.Count - 1; i++)
                {
                    Point p1 = ProfilePoints[i][4];
                    Point p2 = ProfilePoints[i + 1][4];
                    Point p3 = ProfilePoints[i + 1][5];
                    Point p4 = ProfilePoints[i][5];
                    p1 = Utility.Translate(p1, new Vector(0, -50, 0));
                    p2 = Utility.Translate(p2, new Vector(0, -50, 0));
                    var face = new RebarLegFace();
                    face.Contour.AddContourPoint(new ContourPoint(p1, null));
                    face.Contour.AddContourPoint(new ContourPoint(p2, null));
                    face.Contour.AddContourPoint(new ContourPoint(p3, null));
                    face.Contour.AddContourPoint(new ContourPoint(p4, null));
                    rebarSet.LegFaces.Add(face);
                    list4.Add(p1);
                    list4.Add(p2);
                    list5.Add(p3);
                    list5.Add(p4);
                }
                double maxZ = (from Point p in list4
                               orderby p.Y ascending
                               select p).LastOrDefault().Y;
                double minZ = (from Point p in list5
                               orderby p.Y ascending
                               select p).FirstOrDefault().Y;

                startGL = new Point(ProfilePoints[0][4].X, maxZ, ProfilePoints[0][4].Z);
                endGL = new Point(ProfilePoints[0][5].X, minZ, ProfilePoints[0][5].Z);
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
            guideline.Spacing.EndOffset = number == 2 ? 0 : 100;

            guideline.Curve.AddContourPoint(new ContourPoint(startGL, null));
            guideline.Curve.AddContourPoint(new ContourPoint(endGL, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            new Model().CommitChanges();
            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);

            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2 });

        }
        void BackwallLongitudinalTopRebar()
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

            Point startGL, endGL;

            for (int i = 0; i < ProfilePoints.Count - 1; i++)
            {
                Point p1 = ProfilePoints[i][3];
                Point p2 = ProfilePoints[i + 1][3];
                Point p3 = ProfilePoints[i + 1][4];
                Point p4 = ProfilePoints[i][4];
                var face = new RebarLegFace();
                face.Contour.AddContourPoint(new ContourPoint(p1, null));
                face.Contour.AddContourPoint(new ContourPoint(p2, null));
                face.Contour.AddContourPoint(new ContourPoint(p3, null));
                face.Contour.AddContourPoint(new ContourPoint(p4, null));
                rebarSet.LegFaces.Add(face);
            }

            Vector normal = Utility.GetVectorFromTwoPoints(ProfilePoints[0][0], ProfilePoints[1][0]).GetNormal();
            GeometricPlane backwallPlane = new GeometricPlane(ProfilePoints[0][4], normal);

            Line sLine = new Line(ProfilePoints[0][3], ProfilePoints[1][3]);
            Line eLine = new Line(ProfilePoints[0][4], ProfilePoints[1][4]);
            startGL = Utility.GetExtendedIntersection(sLine, backwallPlane, 2);
            endGL = Utility.GetExtendedIntersection(eLine, backwallPlane, 2);

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

            for (int i = 0; i < ProfilePoints.Count - 1; i++)
            {
                Point p1 = ProfilePoints[i][1];
                Point p2 = ProfilePoints[i + 1][1];
                Point p3 = ProfilePoints[i + 1][2];
                Point p4 = ProfilePoints[i][2];
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
            guideline.Spacing.EndOffset = 0;

            Vector normal = Utility.GetVectorFromTwoPoints(ProfilePoints[0][0], ProfilePoints[1][0]).GetNormal();
            GeometricPlane backwallPlane = new GeometricPlane(ProfilePoints[0][2], normal);

            Line sLine = new Line(ProfilePoints[0][1], ProfilePoints[1][1]);
            Line eLine = new Line(ProfilePoints[0][2], ProfilePoints[1][2]);
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

            int secondNumber = number == 0 ? 1 : ProfilePoints.Count - 2;
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
            Point p1e = ProfilePoints[secondNumber][1];
            Point p3e = ProfilePoints[secondNumber][3];
            Point p4e = ProfilePoints[secondNumber][4];
            Point p5e = ProfilePoints[secondNumber][5];
            Point p6e = ProfilePoints[secondNumber][6];
            Point p7e = ProfilePoints[secondNumber][7];
            Point p8e = ProfilePoints[secondNumber][8];
            Point p9e = ProfilePoints[secondNumber][9];
            Point p2e = new Point(p7e.X, p1e.Y, p7e.Z);
            List<Point> secondListPoint = new List<Point> { p0e, p1e, p2e, p3e, p4e, p5e, p6e, p7e, p8e, p9e };

            List<Point> offsetedPoints = new List<Point>();
            for (int i = 0; i < listPoint.Count; i++)
            {
                Vector sectionDir = Utility.GetVectorFromTwoPoints(listPoint[i], secondListPoint[i]).GetNormal();
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

            int secondNumber = number == 0 ? 1 : ProfilePoints.Count - 2;

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

            int secondNumber = number == 0 ? 1 : ProfilePoints.Count - 2;
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

            List<Point> list1 = new List<Point>();
            List<Point> list8 = new List<Point>();
            for (int i = 0; i < ProfilePoints.Count - 1; i++)
            {
                Point p1 = ProfilePoints[i][1];
                Point p8 = ProfilePoints[i][8];
                list1.Add(p1);
                list8.Add(p8);
            }

            double min1Y = (from Point p in list1
                            orderby p.Y ascending
                            select p).First().Y;
            double simpleHeight = (from Point p in list8
                                   orderby p.Y ascending
                                   select p).First().Y;

            double height = Math.Abs(min1Y) + Math.Abs(ProfilePoints[0][0].Y);
            double correctedHeight = height - startOffset;
            int corTotalNumberOfRows = (int)Math.Ceiling(correctedHeight / Convert.ToDouble(verticalSpacing));
            double offset = startOffset + 10 * Convert.ToInt32(rebarSize);
            Vector dir = Utility.GetVectorFromTwoPoints(ProfilePoints[0][8], ProfilePoints[0][9]).GetNormal();

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
                    for (int j = 0; j < ProfilePoints.Count - 1; j++)
                    {
                        CShapeRebarComplex(rebarDiameter, newoffset, horizontalSpacing, j);
                    }
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

            int sn = ProfilePoints.Count - 1;
            Point startOuterPoint = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y + newoffset, ProfilePoints[0][0].Z);
            Point endOuterPoint = new Point(ProfilePoints[sn][0].X, ProfilePoints[sn][0].Y + newoffset, ProfilePoints[sn][0].Z);
            Point startInnerPoint = new Point(ProfilePoints[0][9].X, ProfilePoints[0][9].Y + newoffset, ProfilePoints[0][9].Z);
            Point endInnerPoint = new Point(ProfilePoints[sn][9].X, ProfilePoints[sn][9].Y + newoffset, ProfilePoints[sn][9].Z);

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
        void CShapeRebarComplex(int rebarDiameter, double newoffset, double horizontalSpacing, int number)
        {
            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "ABT_CSRC";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarDiameter));
            rebarSet.RebarProperties.Size = Convert.ToString(rebarDiameter);
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarDiameter));
            rebarSet.LayerOrderNumber = 1;

            int f = number;
            int s = number + 1;
            if (ProfilePoints[f][8].Y > ProfilePoints[s][8].Y)
            {
                int temp = f;
                f = s;
                s = temp;
            }

            Point endOuterPoint, endInnerPoint;
            Point startOuterPoint = new Point(ProfilePoints[f][0].X, ProfilePoints[f][0].Y + newoffset, ProfilePoints[f][0].Z);
            Point startInnerPoint = new Point(ProfilePoints[f][9].X, ProfilePoints[f][9].Y + newoffset, ProfilePoints[f][9].Z);

            Vector longitudinalVector = Utility.GetVectorFromTwoPoints(ProfilePoints[f][8], ProfilePoints[s][8]);

            endOuterPoint = Utility.Translate(startOuterPoint, longitudinalVector);
            endInnerPoint = Utility.Translate(startInnerPoint, longitudinalVector);

            Line startLine = new Line(startOuterPoint, startInnerPoint);
            Line endLine = new Line(endOuterPoint, endInnerPoint);

            Vector outerXdir = Utility.GetVectorFromTwoPoints(ProfilePoints[f][0], ProfilePoints[s][0]).GetNormal();
            Vector outerYdir = Utility.GetVectorFromTwoPoints(ProfilePoints[f][0], ProfilePoints[s][1]).GetNormal();
            Vector innerXdir, innerYdir;
            Point innerPlaneOrigin;
            if (newoffset < Math.Abs(ProfilePoints[f][9].Y) + Math.Abs(ProfilePoints[f][7].Y))
            {
                innerXdir = Utility.GetVectorFromTwoPoints(ProfilePoints[f][8], ProfilePoints[s][8]).GetNormal();
                innerYdir = Utility.GetVectorFromTwoPoints(ProfilePoints[f][8], ProfilePoints[f][7]).GetNormal();
                innerPlaneOrigin = ProfilePoints[f][8];
            }
            else
            {
                innerXdir = Utility.GetVectorFromTwoPoints(ProfilePoints[f][7], ProfilePoints[s][7]).GetNormal();
                innerYdir = Utility.GetVectorFromTwoPoints(ProfilePoints[f][7], ProfilePoints[f][6]).GetNormal();
                innerPlaneOrigin = ProfilePoints[f][7];
            }
            GeometricPlane outerPlane = new GeometricPlane(ProfilePoints[f][0], outerXdir, outerYdir);
            GeometricPlane innerPlane = new GeometricPlane(innerPlaneOrigin, innerXdir, innerYdir);

            Point startOuterIntersection = Utility.GetExtendedIntersection(startLine, outerPlane, 10);
            Point startInnerIntersection = Utility.GetExtendedIntersection(startLine, innerPlane, 10);
            Point endOuterIntersection = Utility.GetExtendedIntersection(endLine, outerPlane, 10);
            Point endInnerIntersection = Utility.GetExtendedIntersection(endLine, innerPlane, 10);

            Point bottomStartOuterPoint = Utility.Translate(startOuterIntersection, outerYdir, -2 * GetHookLength(rebarDiameter));
            Point bottomStartInnerPoint = Utility.Translate(startInnerIntersection, innerYdir, 2 * GetHookLength(rebarDiameter));
            Point bottomEndtOuterPoint = Utility.Translate(endOuterIntersection, outerYdir, -2 * GetHookLength(rebarDiameter));
            Point bottomEndInnerPoint = Utility.Translate(endInnerIntersection, innerYdir, 2 * GetHookLength(rebarDiameter));

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(startOuterPoint, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(endOuterPoint, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(endInnerIntersection, null));
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
            outerFace.Contour.AddContourPoint(new ContourPoint(endInnerIntersection, null));
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

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
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

            int secondNumber = number == 0 ? 1 : ProfilePoints.Count - 2;
            Point p0 = ProfilePoints[number][0];
            Point p1 = ProfilePoints[number][1];
            Point p2 = ProfilePoints[number][2];
            Point p8 = ProfilePoints[number][8];
            Point p9 = ProfilePoints[number][9];

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

            Vector longitudinal = Utility.GetVectorFromTwoPoints(ProfilePoints[0][0], ProfilePoints[1][0]).GetNormal();

            Point offsetP0 = Utility.TranslePointByVectorAndDistance(p0, longitudinal, c * 40 * rebarDiameter);
            Point offsetP9 = Utility.TranslePointByVectorAndDistance(p9, longitudinal, c * 40 * rebarDiameter);

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
            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
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

            int secondNumber = number == 0 ? 1 : ProfilePoints.Count - 2;
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
            guideline.Spacing.EndOffset = 50;

            Point correctedP8 = new Point(p8.X, p2.Y, p8.Z);
            Point correctedP8e2 = new Point(p8e.X, p2e.Y, p8e.Z);

            Vector normal = Utility.GetVectorFromTwoPoints(p2, p2e).GetNormal();
            GeometricPlane backwallPlane = new GeometricPlane(p2, normal);

            Line sLine = new Line(p2, p2e);
            Line eLine = new Line(correctedP8, correctedP8e2);
            Point startIntersection = Utility.GetExtendedIntersection(sLine, backwallPlane, 2);
            Point endIntersection = Utility.GetExtendedIntersection(eLine, backwallPlane, 2);

            guideline.Curve.AddContourPoint(new ContourPoint(startIntersection, null));
            guideline.Curve.AddContourPoint(new ContourPoint(endIntersection, null));
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

            new Model().CommitChanges();
            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
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

            int secondNumber = number == 0 ? 1 : ProfilePoints.Count - 2;
            Point p2 = ProfilePoints[number][2];
            Point p3 = ProfilePoints[number][3];
            Point p4 = ProfilePoints[number][4];
            Point p5 = ProfilePoints[number][5];
            Point p7 = ProfilePoints[number][7];
            Point p8 = ProfilePoints[number][8];

            Point p2e = ProfilePoints[secondNumber][2];
            Point p3e = ProfilePoints[secondNumber][3];
            Point p4e = ProfilePoints[secondNumber][4];
            Point p5e = ProfilePoints[secondNumber][5];
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
            guideline.Spacing.StartOffset = 50;
            guideline.Spacing.EndOffset = 50;

            Vector normal = Utility.GetVectorFromTwoPoints(p2, p2e).GetNormal();
            GeometricPlane backwallPlane = new GeometricPlane(p2, normal);

            Line sLine = new Line(p3, p3e);
            Line eLine = new Line(p4, p4e);
            Point startIntersection = Utility.GetExtendedIntersection(sLine, backwallPlane, 2);
            Point endIntersection = Utility.GetExtendedIntersection(eLine, backwallPlane, 2);

            guideline.Curve.AddContourPoint(new ContourPoint(startIntersection, null));
            guideline.Curve.AddContourPoint(new ContourPoint(endIntersection, null));
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

            new Model().CommitChanges();
            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
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

            int secondNumber = number == 0 ? 1 : ProfilePoints.Count - 2;
            Point p2 = ProfilePoints[number][2];
            Point p4 = ProfilePoints[number][4];
            Point p5 = ProfilePoints[number][5];
            Point p6 = ProfilePoints[number][6];
            Point p7 = ProfilePoints[number][7];
            Point p8 = ProfilePoints[number][8];

            Point p2e = ProfilePoints[secondNumber][2];
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

            Point correctedP5s = Utility.GetExtendedIntersection(startLine45, geometricPlaneBottom, 10);
            Point correctedP5e = Utility.GetExtendedIntersection(endLine45, geometricPlaneBottom, 10);

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
            guideline.Spacing.StartOffset = 50;
            guideline.Spacing.EndOffset = 100;

            Point startGL = new Point(p5.X, p6.Y, p5.Z);
            Point endGL = new Point(p5e.X, p6e.Y, p5e.Z);

            Vector normal = Utility.GetVectorFromTwoPoints(p2, p2e).GetNormal();
            GeometricPlane backwallPlane = new GeometricPlane(p2, normal);

            Line sLine = new Line(startGL, endGL);
            Line eLine = new Line(p6, p6e);
            Point startIntersection = Utility.GetExtendedIntersection(sLine, backwallPlane, 2);
            Point endIntersection = Utility.GetExtendedIntersection(eLine, backwallPlane, 2);

            guideline.Curve.AddContourPoint(new ContourPoint(startIntersection, null));
            guideline.Curve.AddContourPoint(new ContourPoint(endIntersection, null));
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

            new Model().CommitChanges();
            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2, 2, 2 });
        }
        RebarGuideline GetPresetGuideline(string spacing, int number)
        {
            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing),
                SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            bool isLast = ProfilePoints.Count - 2 == number;
            guideline.Spacing.StartOffset = number == 0 ? 100 : 50;
            guideline.Spacing.StartOffsetType = number == 0 ? RebarSpacing.OffsetEnum.MINIMUM : RebarSpacing.OffsetEnum.EXACT;
            guideline.Spacing.EndOffset = isLast ? 100 : 50;
            guideline.Spacing.EndOffsetType = isLast ? RebarSpacing.OffsetEnum.MINIMUM : RebarSpacing.OffsetEnum.EXACT;
            return guideline;
        }
        #endregion
        #region Fields
        //  public static double Height;
        public static double Width;
        public static List<double> FrontHeight = new List<double>();
        public static double ShelfWidth;
        public static double ShelfHeight;
        public static double BackwallWidth;
        public static double CantileverWidth;
        public static double CantileverHeight;
        public static List<double> BackwallTopHeight = new List<double>();
        public static List<double> BackwallBottomHeight = new List<double>();
        public static double SkewHeight;
        public static List<double> Height = new List<double>();
        public static List<double> Length = new List<double>();
        public static double FullWidth;
        public static double FullLength = 0;
        public static double HorizontalOffset;
        #endregion
    }
}
