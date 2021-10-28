using System;
using System.Collections.Generic;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model;

namespace ZeroTouchTekla
{
    public class DRTW : Element
    {
        #region Constructor
        public DRTW(Beam part) : base(part)
        {
        }
        public DRTW(Beam part, Beam secondPart) : base(part)
        {
            GetProfilePointsAndParameters(part, secondPart);
        }
        #endregion
        #region PublicMethods
        public static void GetProfilePointsAndParameters(Beam beam1, Beam beam2)
        {
            string[] profileValues1 = GetProfileValues(beam1);
            string[] profileValues2 = GetProfileValues(beam2);
            //RTW Height*CorniceHeight*BottomWidth*TopWidth*CorniceWidth
            //RTWV Height*CorniceHeight*BottomWidth*TopWidth*CorniceWidth*Height2
            double height1 = Convert.ToDouble(profileValues1[0]);
            double corniceHeight = Convert.ToDouble(profileValues1[1]);
            double bottomWidth = Convert.ToDouble(profileValues1[2]);
            double topWidth = Convert.ToDouble(profileValues1[3]);
            double corniceWidth = Convert.ToDouble(profileValues1[4]);
            double length1 = Distance.PointToPoint(beam1.StartPoint, beam1.EndPoint);
            double length2 = Distance.PointToPoint(beam2.StartPoint, beam2.EndPoint);
            double fullWidth = corniceWidth + bottomWidth;
            double hToW = (bottomWidth - (topWidth - corniceWidth)) / height1;
            double height2, height3;
            if (profileValues1.Length > 5)
            {
                height2 = Convert.ToDouble(profileValues1[5]);
                height3 = Convert.ToDouble(profileValues2[0]);

            }
            else
            {
                height2 = Convert.ToDouble(profileValues2[0]);
                height3 = Convert.ToDouble(profileValues2[5]);
            }
            double bottomWidth2 = hToW * height2 + (topWidth - corniceWidth);
            double bottomWidth3 = hToW * height3 + (topWidth - corniceWidth);

            Height1 = height1;
            CorniceHeight = corniceHeight;
            BottomWidth1 = bottomWidth;
            TopWidth = topWidth;
            CorniceWidth = corniceWidth;
            Length1 = length1;
            Length2 = length2;
            Height2 = height2;
            Height3 = height3;
            BottomWidth2 = bottomWidth2;
            BottomWidth3 = bottomWidth3;

            Point p0 = new Point(0, -height1 / 2.0, fullWidth / 2.0 - corniceWidth);
            Point p1 = new Point(0, height1 / 2.0 - corniceHeight, p0.Z);
            Point p2 = new Point(0, p1.Y, p1.Z + corniceWidth);
            Point p3 = new Point(0, p2.Y + corniceHeight, p2.Z);
            Point p4 = new Point(0, p3.Y, p3.Z - topWidth);
            Point p5 = new Point(0, -height1 / 2.0, -fullWidth / 2.0);

            List<Point> firstProfile = new List<Point> { p0, p1, p2, p3, p4, p5 };

            Point s0 = new Point(length1, -height1 / 2.0, fullWidth / 2.0 - corniceWidth);
            Point s1 = new Point(length1, s0.Y + height2 - corniceHeight, s0.Z);
            Point s2 = new Point(length1, s1.Y, p1.Z + corniceWidth);
            Point s3 = new Point(length1, s2.Y + corniceHeight, s2.Z);
            Point s4 = new Point(length1, s3.Y, s3.Z - topWidth);
            Point s5 = new Point(length1, -height1 / 2.0, s0.Z - bottomWidth2);
            List<Point> secondProfile = new List<Point> { s0, s1, s2, s3, s4, s5 };

            Point t0 = new Point(length1 + length2, -height1 / 2.0, fullWidth / 2.0 - corniceWidth);
            Point t1 = new Point(length1 + length2, t0.Y + height3 - corniceHeight, t0.Z);
            Point t2 = new Point(length1 + length2, t1.Y, p1.Z + corniceWidth);
            Point t3 = new Point(length1 + length2, t2.Y + corniceHeight, t2.Z);
            Point t4 = new Point(length1 + length2, t3.Y, t3.Z - topWidth);
            Point t5 = new Point(length1 + length2, -height1 / 2.0, t0.Z - bottomWidth3);
            List<Point> thirdProfile = new List<Point> { t0, t1, t2, t3, t4, t5 };

            List<List<Point>> beamPoints = new List<List<Point>> { firstProfile, secondProfile, thirdProfile };
            ProfilePoints = beamPoints;
            ElementFace = new ElementFace(ProfilePoints);
            List<double> heightList = new List<double> { height1, height2, height3 };
            heightList.Sort();
            maxHeight = heightList[heightList.Count - 1];
            minHeight = heightList[0];
        }
        new public void Create()
        {
            InnerVerticalRebar();
            OuterVerticalRebar();
            InnerLongitudinalRebar();
            OuterLongitudinalRebar();
            CornicePerpendicularRebar();
            FirstCorniceLongitudinalRebar();
            SecondCorniceLongitudinalRebar();
            ClosingCShapeRebar(true);
            ClosingCShapeRebar(false);

            ClosingLongitudinalRebar(true);
            ClosingLongitudinalRebar(false);

            FirstCShapeRebar();
            SecondCShapeRebar();

        }
        new public void CreateSingle(string rebarName)
        {
            rebarName = rebarName.Split('_')[1];
            RebarType rType;
            Enum.TryParse(rebarName, out rType);
            switch (rType)
            {
                case RebarType.IVR:
                    InnerVerticalRebar();
                    break;
                case RebarType.OVR:
                    OuterVerticalRebar();
                    break;
                case RebarType.ILR:
                    InnerLongitudinalRebar();
                    break;
                case RebarType.OLR:
                    OuterLongitudinalRebar();
                    break;
                case RebarType.CrPR:
                    CornicePerpendicularRebar();
                    break;
                case RebarType.CrLR:
                    FirstCorniceLongitudinalRebar();
                    break;
                case RebarType.CCSR:
                    ClosingCShapeRebar(true);
                    ClosingCShapeRebar(false);
                    break;
                case RebarType.CLR:
                    ClosingLongitudinalRebar(true);
                    ClosingLongitudinalRebar(false);
                    break;
            }
        }
        #endregion
        #region PrivateMethods
        void InnerVerticalRebar()
        {
            string rebarSize = Program.ExcelDictionary["IVR_Diameter"];
            string spacing = Program.ExcelDictionary["IVR_Spacing"];
            int addSplitter = Convert.ToInt32(Program.ExcelDictionary["IVR_AddSplitter"]);
            string secondRebarSize = Program.ExcelDictionary["IVR_SecondDiameter"];
            double spliterOffset = Convert.ToDouble(Program.ExcelDictionary["IVR_SplitterOffset"]) + Convert.ToDouble(rebarSize) * 20;
            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "RTW_IVR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            Point startp3bis = new Point(ProfilePoints[0][1].X, ProfilePoints[0][1].Y + CorniceHeight, ProfilePoints[0][1].Z);
            Point midp3pis = new Point(ProfilePoints[1][1].X, ProfilePoints[1][1].Y + CorniceHeight, ProfilePoints[1][1].Z);
            Point endp3pis = new Point(ProfilePoints[2][1].X, ProfilePoints[2][1].Y + CorniceHeight, ProfilePoints[2][1].Z);

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(endp3pis, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(midp3pis, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(startp3bis, null));
            rebarSet.LegFaces.Add(mainFace);

            Point offsetedStartPoint = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y, ProfilePoints[0][0].Z + 1000);
            Point offsetedEndPoint = new Point(ProfilePoints[2][0].X, ProfilePoints[2][0].Y, ProfilePoints[2][0].Z + 1000);

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

                Point startBottomPoint = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y + spliterOffset, ProfilePoints[0][0].Z);
                Point endBottomPoint = new Point(ProfilePoints[2][0].X, ProfilePoints[2][0].Y + spliterOffset, ProfilePoints[2][0].Z);

                bottomSpliter.Curve.AddContourPoint(new ContourPoint(startBottomPoint, null));
                bottomSpliter.Curve.AddContourPoint(new ContourPoint(endBottomPoint, null));
                bottomSpliter.Insert();

                var topSpliter = new RebarSplitter();
                topSpliter.Father = rebarSet;
                topSpliter.Lapping.LappingType = RebarLapping.LappingTypeEnum.STANDARD_LAPPING;
                topSpliter.Lapping.LapSide = RebarLapping.LapSideEnum.LAP_MIDDLE;
                topSpliter.Lapping.LapPlacement = RebarLapping.LapPlacementEnum.ON_LEG_FACE;
                topSpliter.BarsAffected = BaseRebarModifier.BarsAffectedEnum.EVERY_SECOND_BAR;
                topSpliter.FirstAffectedBar = 2;

                Point startTopPoint = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y, ProfilePoints[0][0].Z);
                startTopPoint.Translate(0, spliterOffset + 1.3 * 40 * Convert.ToDouble(rebarSize), 0);

                Point endTopPoint = new Point(ProfilePoints[2][0].X, ProfilePoints[2][0].Y, ProfilePoints[2][0].Z);
                endTopPoint.Translate(0, spliterOffset + 1.3 * 40 * Convert.ToDouble(rebarSize), 0);

                topSpliter.Curve.AddContourPoint(new ContourPoint(startTopPoint, null));
                topSpliter.Curve.AddContourPoint(new ContourPoint(endTopPoint, null));
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
                    propertyModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[2][1], null));
                    propertyModifier.Insert();
                }
                new Model().CommitChanges();
            }

            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 3 });
        }
        void OuterVerticalRebar()
        {
            string rebarSize = Program.ExcelDictionary["OVR_Diameter"];
            string spacing = Program.ExcelDictionary["OVR_Spacing"];
            int addSplitter = Convert.ToInt32(Program.ExcelDictionary["OVR_AddSplitter"]);
            string secondRebarSize = Program.ExcelDictionary["OVR_SecondDiameter"];
            double spliterOffset = Convert.ToDouble(Program.ExcelDictionary["OVR_SplitterOffset"]) + Convert.ToDouble(rebarSize) * 20;

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "RTW_OVR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][5], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][5], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][4], null));
            rebarSet.LegFaces.Add(mainFace);

            var mainFace2 = new RebarLegFace();
            mainFace2.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][5], null));
            mainFace2.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][5], null));
            mainFace2.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][4], null));
            mainFace2.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
            rebarSet.LegFaces.Add(mainFace2);

            Point offsetedStartPoint = new Point(ProfilePoints[0][5].X, ProfilePoints[0][5].Y, ProfilePoints[0][5].Z - 40 * Convert.ToInt32(rebarSize));
            Point offsetedEndPoint = new Point(ProfilePoints[2][5].X, ProfilePoints[2][5].Y, ProfilePoints[2][5].Z - 40 * Convert.ToInt32(rebarSize));

            var bottomFace = new RebarLegFace();
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][5], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][5], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][5], null));
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

            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][5], null));
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][5], null));
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[2][5], null));

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

                Vector startVector = Utility.GetVectorFromTwoPoints(ProfilePoints[0][5], ProfilePoints[0][4]).GetNormal();
                Vector midVector = Utility.GetVectorFromTwoPoints(ProfilePoints[1][5], ProfilePoints[1][4]).GetNormal();
                Vector endVector = Utility.GetVectorFromTwoPoints(ProfilePoints[2][5], ProfilePoints[2][4]).GetNormal();

                double startCoefficient = 1.0 / startVector.Y;
                double midCoefficient = 1.0 / midVector.Y;
                double endCoefficient = 1.0 / endVector.Y;

                Point startIntersection = new Point(ProfilePoints[0][5].X, ProfilePoints[0][5].Y, ProfilePoints[0][5].Z);
                startIntersection.Translate(spliterOffset * startCoefficient * startVector.X, spliterOffset * startCoefficient * startVector.Y, spliterOffset * startCoefficient * startVector.Z);
                Point midIntersection = new Point(ProfilePoints[1][5].X, ProfilePoints[1][5].Y, ProfilePoints[1][5].Z);
                midIntersection.Translate(spliterOffset * midCoefficient * midVector.X, spliterOffset * midCoefficient * midVector.Y, spliterOffset * midCoefficient * midVector.Z);
                Point endIntersection = new Point(ProfilePoints[2][5].X, ProfilePoints[2][5].Y, ProfilePoints[2][5].Z);
                endIntersection.Translate(spliterOffset * endCoefficient * endVector.X, spliterOffset * endCoefficient * endVector.Y, spliterOffset * endCoefficient * endVector.Z);

                bottomSpliter.Curve.AddContourPoint(new ContourPoint(startIntersection, null));
                bottomSpliter.Curve.AddContourPoint(new ContourPoint(midIntersection, null));
                bottomSpliter.Curve.AddContourPoint(new ContourPoint(endIntersection, null));
                bottomSpliter.Insert();

                var topSpliter = new RebarSplitter();
                topSpliter.Father = rebarSet;
                topSpliter.Lapping.LappingType = RebarLapping.LappingTypeEnum.STANDARD_LAPPING;
                topSpliter.Lapping.LapSide = RebarLapping.LapSideEnum.LAP_MIDDLE;
                topSpliter.Lapping.LapPlacement = RebarLapping.LapPlacementEnum.ON_LEG_FACE;
                topSpliter.BarsAffected = BaseRebarModifier.BarsAffectedEnum.EVERY_SECOND_BAR;
                topSpliter.FirstAffectedBar = 2;

                double secondSplitterOffset = spliterOffset + 1.3 * Convert.ToDouble(rebarSize) * 40;
                Point startIntersection2 = new Point(ProfilePoints[0][5].X, ProfilePoints[0][5].Y, ProfilePoints[0][5].Z);
                startIntersection2.Translate(secondSplitterOffset * startCoefficient * startVector.X, secondSplitterOffset * startCoefficient * startVector.Y, secondSplitterOffset * startCoefficient * startVector.Z);
                Point midIntersection2 = new Point(ProfilePoints[1][5].X, ProfilePoints[1][5].Y, ProfilePoints[1][5].Z);
                midIntersection2.Translate(secondSplitterOffset * midCoefficient * midVector.X, secondSplitterOffset * midCoefficient * midVector.Y, secondSplitterOffset * midCoefficient * midVector.Z);
                Point endIntersection2 = new Point(ProfilePoints[2][5].X, ProfilePoints[2][5].Y, ProfilePoints[2][5].Z);
                endIntersection2.Translate(secondSplitterOffset * endCoefficient * endVector.X, secondSplitterOffset * endCoefficient * endVector.Y, secondSplitterOffset * endCoefficient * endVector.Z);

                topSpliter.Curve.AddContourPoint(new ContourPoint(startIntersection2, null));
                topSpliter.Curve.AddContourPoint(new ContourPoint(midIntersection2, null));
                topSpliter.Curve.AddContourPoint(new ContourPoint(endIntersection2, null));
                topSpliter.Insert();

                if (rebarSize != secondRebarSize)
                {
                    var propertyModifier = new RebarPropertyModifier();
                    propertyModifier.Father = rebarSet;
                    propertyModifier.BarsAffected = BaseRebarModifier.BarsAffectedEnum.ALL_BARS;
                    propertyModifier.RebarProperties.Size = secondRebarSize;
                    propertyModifier.RebarProperties.Class = SetClass(Convert.ToDouble(secondRebarSize));
                    propertyModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][4], null));
                    propertyModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
                    propertyModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[2][4], null));
                    propertyModifier.Insert();
                }
                new Model().CommitChanges();
            }

            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 1, 3 });
        }
        void InnerLongitudinalRebar()
        {
            string rebarSize = Program.ExcelDictionary["ILR_Diameter"];
            string secondRebarSize = Program.ExcelDictionary["ILR_SecondDiameter"];
            string spacing = Program.ExcelDictionary["ILR_Spacing"];
            double startOffset = Convert.ToDouble(Program.ExcelDictionary["ILR_StartOffset"]);
            double firstLength = Convert.ToDouble(Program.ExcelDictionary["ILR_SecondDiameterLength"]);
            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "RTW_ILR";
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

            Vector vector = Utility.GetVectorFromTwoPoints(ProfilePoints[0][0], ProfilePoints[0][1]).GetNormal();
            Point sp = new Point(ProfilePoints[0][0].X + vector.X * maxHeight, ProfilePoints[0][0].Y + vector.Y * maxHeight, ProfilePoints[0][0].Z + vector.Z * maxHeight);
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            guideline.Curve.AddContourPoint(new ContourPoint(sp, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            if (rebarSize != secondRebarSize)
            {
                var propertyModifier = new RebarPropertyModifier();
                propertyModifier.Father = rebarSet;
                propertyModifier.BarsAffected = BaseRebarModifier.BarsAffectedEnum.ALL_BARS;
                propertyModifier.RebarProperties.Size = secondRebarSize;
                propertyModifier.RebarProperties.Class = SetClass(Convert.ToDouble(secondRebarSize));

                Point secondPoint = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y + startOffset + firstLength, ProfilePoints[0][0].Z);
                propertyModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
                propertyModifier.Curve.AddContourPoint(new ContourPoint(secondPoint, null));
                propertyModifier.Insert();
            }
            new Model().CommitChanges();

            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2 });
        }
        void OuterLongitudinalRebar()
        {
            string rebarSize = Program.ExcelDictionary["OLR_Diameter"];
            string secondRebarSize = Program.ExcelDictionary["OLR_SecondDiameter"];
            string spacing = Program.ExcelDictionary["OLR_Spacing"];
            double startOffset = Convert.ToDouble(Program.ExcelDictionary["OLR_StartOffset"]);
            double firstLength = Convert.ToDouble(Program.ExcelDictionary["OLR_SecondDiameterLength"]);
            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "RTW_OLR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            var firstMainFace = new RebarLegFace();
            firstMainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][5], null));
            firstMainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][5], null));
            firstMainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
            firstMainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][4], null));
            rebarSet.LegFaces.Add(firstMainFace);

            var secondMainFace = new RebarLegFace();
            secondMainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][5], null));
            secondMainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][5], null));
            secondMainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][4], null));
            secondMainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
            rebarSet.LegFaces.Add(secondMainFace);

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

            Vector vector = Utility.GetVectorFromTwoPoints(ProfilePoints[0][5], ProfilePoints[0][4]).GetNormal();
            Point sp = Utility.TranslePointByVectorAndDistance(ProfilePoints[0][5], vector, maxHeight);
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][5], null));
            guideline.Curve.AddContourPoint(new ContourPoint(sp, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            if (rebarSize != secondRebarSize)
            {
                var propertyModifier = new RebarPropertyModifier();
                propertyModifier.Father = rebarSet;
                propertyModifier.BarsAffected = BaseRebarModifier.BarsAffectedEnum.ALL_BARS;
                propertyModifier.RebarProperties.Size = secondRebarSize;
                propertyModifier.RebarProperties.Class = SetClass(Convert.ToDouble(secondRebarSize));

                Point secondPoint = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y + startOffset + firstLength, ProfilePoints[0][0].Z);
                propertyModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
                propertyModifier.Curve.AddContourPoint(new ContourPoint(secondPoint, null));
                propertyModifier.Insert();
            }

            new Model().CommitChanges();

            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2, 2 });
        }
        void CornicePerpendicularRebar()
        {
            string rebarSize = Program.ExcelDictionary["CPR_Diameter"];
            string spacing = Program.ExcelDictionary["CPR_Spacing"];
            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "DRTW_CrPR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            var firstMainFace = new RebarLegFace();
            firstMainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            firstMainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
            firstMainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][3], null));
            firstMainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            rebarSet.LegFaces.Add(firstMainFace);

            var secondMainFace = new RebarLegFace();
            secondMainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
            secondMainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][2], null));
            secondMainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][3], null));
            secondMainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][3], null));
            rebarSet.LegFaces.Add(secondMainFace);

            var firstTopFace = new RebarLegFace();
            firstTopFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            firstTopFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][3], null));
            firstTopFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
            firstTopFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][4], null));
            rebarSet.LegFaces.Add(firstTopFace);

            var secondTopFace = new RebarLegFace();
            secondTopFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][3], null));
            secondTopFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][3], null));
            secondTopFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][4], null));
            secondTopFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
            rebarSet.LegFaces.Add(secondTopFace);

            var firstOuterFace = new RebarLegFace();
            firstOuterFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][4], null));
            firstOuterFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
            firstOuterFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][5], null));
            firstOuterFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][5], null));
            rebarSet.LegFaces.Add(firstOuterFace);

            var secondOuterFace = new RebarLegFace();
            secondOuterFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
            secondOuterFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][4], null));
            secondOuterFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][5], null));
            secondOuterFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][5], null));
            rebarSet.LegFaces.Add(secondOuterFace);

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

            var hookModifier = new RebarEndDetailModifier();
            hookModifier.Father = rebarSet;
            hookModifier.EndType = RebarEndDetailModifier.EndTypeEnum.HOOK;
            hookModifier.RebarHook.Shape = RebarHookData.RebarHookShapeEnum.HOOK_90_DEGREES;
            hookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            hookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
            hookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[2][2], null));
            hookModifier.Insert();
            new Model().CommitChanges();

            var bottomLengthModifier = new RebarEndDetailModifier();
            bottomLengthModifier.Father = rebarSet;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentLength = 40 * Convert.ToInt32(rebarSize);
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][5], null));
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[2][5], null));
            bottomLengthModifier.Insert();

            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 1, 1, 1, 1, 1 });
        }
        void FirstCorniceLongitudinalRebar()
        {
            string rebarSize = Program.ExcelDictionary["ILR_Diameter"];
            string spacing = Program.ExcelDictionary["ILR_Spacing"];
            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "DRTW_FCrLR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            var firstMainFace = new RebarLegFace();
            firstMainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            firstMainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
            firstMainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][3], null));
            firstMainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            rebarSet.LegFaces.Add(firstMainFace);

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
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            rebarSet.Guidelines.Add(guideline);

            var secondaryGuideLine = new RebarGuideline();
            secondaryGuideLine.Spacing.InheritFromPrimary = true;
            secondaryGuideLine.Spacing.StartOffset = 100;
            secondaryGuideLine.Spacing.EndOffset = 100;
            secondaryGuideLine.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
            secondaryGuideLine.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][3], null));
            rebarSet.Guidelines.Add(secondaryGuideLine);

            bool succes = rebarSet.Insert();
            new Model().CommitChanges();

            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2 });
        }
        void SecondCorniceLongitudinalRebar()
        {
            string rebarSize = Program.ExcelDictionary["ILR_Diameter"];
            string spacing = Program.ExcelDictionary["ILR_Spacing"];
            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "DRTW_FCrLR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            var secondMainFace = new RebarLegFace();
            secondMainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
            secondMainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][2], null));
            secondMainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][3], null));
            secondMainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][3], null));
            rebarSet.LegFaces.Add(secondMainFace);

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
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][3], null));
            rebarSet.Guidelines.Add(guideline);

            var thirdGuideLine = new RebarGuideline();
            thirdGuideLine.Spacing.InheritFromPrimary = true;
            thirdGuideLine.Spacing.StartOffset = 100;
            thirdGuideLine.Spacing.EndOffset = 100;
            thirdGuideLine.Curve.AddContourPoint(new ContourPoint(ProfilePoints[2][2], null));
            thirdGuideLine.Curve.AddContourPoint(new ContourPoint(ProfilePoints[2][3], null));
            rebarSet.Guidelines.Add(thirdGuideLine);

            bool succes = rebarSet.Insert();
            new Model().CommitChanges();

            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2 });
        }
        void ClosingCShapeRebar(bool isStart)
        {
            string rebarSize = Program.ExcelDictionary["CCSR_Diameter"];
            string spacing = Program.ExcelDictionary["CCSR_Spacing"];
            double startOffset = Convert.ToDouble(Program.ExcelDictionary["OLR_StartOffset"]);
            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "RTW_CCSR_" + isStart;
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;
            rebarSet.SetUserProperty("User field 1", 0);

            Point leftBottom, rightBottom, rightTop, leftTop;
            Point endLeftBottom, endRightBottom, endRightTop, endLeftTop;
            if (isStart)
            {
                leftBottom = ProfilePoints[0][0];
                rightBottom = ProfilePoints[0][5];
                rightTop = ProfilePoints[0][4];
                leftTop = new Point(ProfilePoints[0][1].X, ProfilePoints[0][1].Y + CorniceHeight, ProfilePoints[0][1].Z);
                endLeftBottom = ProfilePoints[1][0];
                endRightBottom = ProfilePoints[1][5];
                endRightTop = ProfilePoints[1][4];
                endLeftTop = new Point(ProfilePoints[1][1].X, ProfilePoints[1][1].Y + CorniceHeight, ProfilePoints[1][1].Z);
            }
            else
            {
                leftBottom = ProfilePoints[2][0];
                rightBottom = ProfilePoints[2][5];
                rightTop = ProfilePoints[2][4];
                leftTop = new Point(ProfilePoints[2][1].X, ProfilePoints[2][1].Y + CorniceHeight, ProfilePoints[2][1].Z);
                endLeftBottom = ProfilePoints[1][0];
                endRightBottom = ProfilePoints[1][5];
                endRightTop = ProfilePoints[1][4];
                endLeftTop = new Point(ProfilePoints[1][1].X, ProfilePoints[1][1].Y + CorniceHeight, ProfilePoints[1][1].Z);
            }

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(leftBottom, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(rightBottom, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(rightTop, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(leftTop, null));
            rebarSet.LegFaces.Add(mainFace);

            var innerFace = new RebarLegFace();
            innerFace.Contour.AddContourPoint(new ContourPoint(leftBottom, null));
            innerFace.Contour.AddContourPoint(new ContourPoint(endLeftBottom, null));
            innerFace.Contour.AddContourPoint(new ContourPoint(endLeftTop, null));
            innerFace.Contour.AddContourPoint(new ContourPoint(leftTop, null));
            rebarSet.LegFaces.Add(innerFace);

            var outerFace = new RebarLegFace();
            outerFace.Contour.AddContourPoint(new ContourPoint(rightBottom, null));
            outerFace.Contour.AddContourPoint(new ContourPoint(endRightBottom, null));
            outerFace.Contour.AddContourPoint(new ContourPoint(endRightTop, null));
            outerFace.Contour.AddContourPoint(new ContourPoint(rightTop, null));
            rebarSet.LegFaces.Add(outerFace);

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

            guideline.Curve.AddContourPoint(new ContourPoint(leftBottom, null));
            guideline.Curve.AddContourPoint(new ContourPoint(leftTop, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            rebarSet.SetUserProperty("__MIN_BAR_LENTYPE", 0);
            rebarSet.SetUserProperty("__MIN_BAR_LENGTH", 30 * Convert.ToDouble(rebarSize));
            new Model().CommitChanges();

            //Create RebarEndDetailModifier
            var innerEndDetailModifier = new RebarEndDetailModifier();
            innerEndDetailModifier.Father = rebarSet;
            innerEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            innerEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = 10 * Convert.ToInt32(rebarSize);
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(endLeftBottom, null));
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(endLeftTop, null));
            if (Height2 != 0)
            {
                innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(leftTop, null));
            }
            innerEndDetailModifier.Insert();

            var outerEndDetailModifier = new RebarEndDetailModifier();
            outerEndDetailModifier.Father = rebarSet;
            outerEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            outerEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = 10 * Convert.ToInt32(rebarSize);
            outerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(endRightBottom, null));
            outerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(endRightTop, null));
            if (Height2 != 0)
            {
                outerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(rightTop, null));
            }
            outerEndDetailModifier.Insert();
            new Model().CommitChanges();

            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 2, 2 });
        }
        void ClosingLongitudinalRebar(bool isStart)
        {
            string rebarSize = Program.ExcelDictionary["CLR_Diameter"];
            string spacing = Program.ExcelDictionary["CLR_Spacing"];
            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "RTW_CLR_" + isStart;
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;


            Point leftBottom, rightBottom, rightTop, leftTop;
            if (isStart)
            {
                leftBottom = ProfilePoints[0][0];
                rightBottom = ProfilePoints[0][5];
                rightTop = ProfilePoints[0][4];
                leftTop = new Point(ProfilePoints[0][1].X, ProfilePoints[0][1].Y + CorniceHeight, ProfilePoints[0][1].Z);
            }
            else
            {
                leftBottom = ProfilePoints[2][0];
                rightBottom = ProfilePoints[2][5];
                rightTop = ProfilePoints[2][4];
                leftTop = new Point(ProfilePoints[2][1].X, ProfilePoints[2][1].Y + CorniceHeight, ProfilePoints[2][1].Z);
            }

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(leftBottom, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(rightBottom, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(rightTop, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(leftTop, null));
            rebarSet.LegFaces.Add(mainFace);

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

            guideline.Curve.AddContourPoint(new ContourPoint(leftBottom, null));
            guideline.Curve.AddContourPoint(new ContourPoint(rightBottom, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();

            rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2 });
        }
        void FirstCShapeRebar()
        {
            string rebarSize = Program.ExcelDictionary["CSR_Diameter"];
            string horizontalSpacing = Program.ExcelDictionary["CSR_HorizontalSpacing"];
            string verticalSpacing = Program.ExcelDictionary["CSR_VerticalSpacing"];
            double startOffset = Convert.ToDouble(Program.ExcelDictionary["OLR_StartOffset"]);

            double height = Height1 > Height2 ? Height2 : Height1;
            double corniceHeight = CorniceHeight;


            double correctedHeight = height - startOffset - corniceHeight - 10 * Convert.ToInt32(rebarSize);
            int correctedNumberOfRows = (int)Math.Floor(correctedHeight / Convert.ToDouble(verticalSpacing));
            double offset = startOffset + 10 * Convert.ToInt32(rebarSize);

            for (int i = 0; i < correctedNumberOfRows; i++)
            {
                double newoffset = offset + i * Convert.ToDouble(verticalSpacing);
                var rebarSet = new RebarSet();
                rebarSet.RebarProperties.Name = "RTW_CSR";
                rebarSet.RebarProperties.Grade = "B500SP";
                rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
                rebarSet.RebarProperties.Size = rebarSize;
                rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
                rebarSet.LayerOrderNumber = 1;

                Point startLeftTopPoint = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y + newoffset, ProfilePoints[0][0].Z);
                Point endLeftTopPoint = new Point(ProfilePoints[1][0].X, ProfilePoints[1][0].Y + newoffset, ProfilePoints[1][0].Z);

                Point tempSLP = new Point(startLeftTopPoint.X, startLeftTopPoint.Y, startLeftTopPoint.Z - BottomWidth1 * 2);
                Point tempELP = new Point(endLeftTopPoint.X, endLeftTopPoint.Y, endLeftTopPoint.Z - BottomWidth1 * 2);

                Line startLine = new Line(startLeftTopPoint, tempSLP);
                Line endLine = new Line(endLeftTopPoint, tempELP);

                Vector xAxis = Utility.GetVectorFromTwoPoints(ProfilePoints[0][5], ProfilePoints[0][4]).GetNormal();
                Vector yAxis = Utility.GetVectorFromTwoPoints(ProfilePoints[0][5], ProfilePoints[1][5]).GetNormal();
                GeometricPlane plane = new GeometricPlane(ProfilePoints[0][5], xAxis, yAxis);
                Point startIntersection = Intersection.LineToPlane(startLine, plane) as Point;
                Point endIntersection = Intersection.LineToPlane(endLine, plane) as Point;

                var mainFace = new RebarLegFace();
                mainFace.Contour.AddContourPoint(new ContourPoint(startLeftTopPoint, null));
                mainFace.Contour.AddContourPoint(new ContourPoint(endLeftTopPoint, null));
                mainFace.Contour.AddContourPoint(new ContourPoint(endIntersection, null));
                mainFace.Contour.AddContourPoint(new ContourPoint(startIntersection, null));
                rebarSet.LegFaces.Add(mainFace);

                var innerFace = new RebarLegFace();
                innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
                innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
                innerFace.Contour.AddContourPoint(new ContourPoint(endLeftTopPoint, null));
                innerFace.Contour.AddContourPoint(new ContourPoint(startLeftTopPoint, null));
                rebarSet.LegFaces.Add(innerFace);

                var outerFace = new RebarLegFace();
                outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][5], null));
                outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][5], null));
                outerFace.Contour.AddContourPoint(new ContourPoint(endIntersection, null));
                outerFace.Contour.AddContourPoint(new ContourPoint(startIntersection, null));
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
                guideline.Spacing.EndOffset = Convert.ToDouble(horizontalSpacing) / 2.0;

                guideline.Curve.AddContourPoint(new ContourPoint(startLeftTopPoint, null));
                guideline.Curve.AddContourPoint(new ContourPoint(endLeftTopPoint, null));

                rebarSet.Guidelines.Add(guideline);
                bool succes = rebarSet.Insert();
                rebarSet.SetUserProperty("__MIN_BAR_LENGTH", 20 * Convert.ToDouble(rebarSize));
                new Model().CommitChanges();

                //Create RebarEndDetailModifier
                var innerEndDetailModifier = new RebarEndDetailModifier();
                innerEndDetailModifier.Father = rebarSet;
                innerEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
                innerEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = 10 * Convert.ToInt32(rebarSize);
                innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
                innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
                innerEndDetailModifier.Insert();

                var outerEndDetailModifier = new RebarEndDetailModifier();
                outerEndDetailModifier.Father = rebarSet;
                outerEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
                outerEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = 10 * Convert.ToInt32(rebarSize);
                outerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][5], null));
                outerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][5], null));
                outerEndDetailModifier.Insert();
                new Model().CommitChanges();

                rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
                RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 1, 1 });
            }
        }
        void SecondCShapeRebar()
        {
            string rebarSize = Program.ExcelDictionary["CSR_Diameter"];
            string horizontalSpacing = Program.ExcelDictionary["CSR_HorizontalSpacing"];
            string verticalSpacing = Program.ExcelDictionary["CSR_VerticalSpacing"];
            double startOffset = Convert.ToDouble(Program.ExcelDictionary["OLR_StartOffset"]);

            double height = Height2 > Height3 ? Height3 : Height2;
            double corniceHeight = CorniceHeight;

            double correctedHeight = height - startOffset - corniceHeight - 10 * Convert.ToInt32(rebarSize);
            int correctedNumberOfRows = (int)Math.Floor(correctedHeight / Convert.ToDouble(verticalSpacing))+1;
            double offset = startOffset + 10 * Convert.ToInt32(rebarSize);

            for (int i = 0; i < correctedNumberOfRows; i++)
            {
                double newoffset = offset + i * Convert.ToDouble(verticalSpacing);
                var rebarSet = new RebarSet();
                rebarSet.RebarProperties.Name = "RTW_CSR";
                rebarSet.RebarProperties.Grade = "B500SP";
                rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
                rebarSet.RebarProperties.Size = rebarSize;
                rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
                rebarSet.LayerOrderNumber = 1;

                Point startLeftTopPoint = new Point(ProfilePoints[1][0].X, ProfilePoints[1][0].Y + newoffset, ProfilePoints[1][0].Z);
                Point endLeftTopPoint = new Point(ProfilePoints[2][0].X, ProfilePoints[2][0].Y + newoffset, ProfilePoints[2][0].Z);

                Point tempSLP = new Point(startLeftTopPoint.X, startLeftTopPoint.Y, startLeftTopPoint.Z - BottomWidth1 * 2);
                Point tempELP = new Point(endLeftTopPoint.X, endLeftTopPoint.Y, endLeftTopPoint.Z - BottomWidth1 * 2);

                Line startLine = new Line(startLeftTopPoint, tempSLP);
                Line endLine = new Line(endLeftTopPoint, tempELP);

                Vector xAxis = Utility.GetVectorFromTwoPoints(ProfilePoints[1][5], ProfilePoints[1][4]).GetNormal();
                Vector yAxis = Utility.GetVectorFromTwoPoints(ProfilePoints[1][5], ProfilePoints[2][5]).GetNormal();
                GeometricPlane plane = new GeometricPlane(ProfilePoints[1][5], xAxis, yAxis);
                Point startIntersection = Intersection.LineToPlane(startLine, plane) as Point;
                Point endIntersection = Intersection.LineToPlane(endLine, plane) as Point;

                var mainFace = new RebarLegFace();
                mainFace.Contour.AddContourPoint(new ContourPoint(startLeftTopPoint, null));
                mainFace.Contour.AddContourPoint(new ContourPoint(endLeftTopPoint, null));
                mainFace.Contour.AddContourPoint(new ContourPoint(endIntersection, null));
                mainFace.Contour.AddContourPoint(new ContourPoint(startIntersection, null));
                rebarSet.LegFaces.Add(mainFace);

                var innerFace = new RebarLegFace();
                innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
                innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][0], null));
                innerFace.Contour.AddContourPoint(new ContourPoint(endLeftTopPoint, null));
                innerFace.Contour.AddContourPoint(new ContourPoint(startLeftTopPoint, null));
                rebarSet.LegFaces.Add(innerFace);

                var outerFace = new RebarLegFace();
                outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][5], null));
                outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][5], null));
                outerFace.Contour.AddContourPoint(new ContourPoint(endIntersection, null));
                outerFace.Contour.AddContourPoint(new ContourPoint(startIntersection, null));
                rebarSet.LegFaces.Add(outerFace);

                var guideline = new RebarGuideline();
                guideline.Spacing.Zones.Add(new RebarSpacingZone
                {
                    Spacing = Convert.ToInt32(horizontalSpacing),
                    SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                    Length = 100,
                    LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
                });
                guideline.Spacing.StartOffset = Convert.ToDouble(horizontalSpacing) / 2.0;
                guideline.Spacing.EndOffset = 100;

                guideline.Curve.AddContourPoint(new ContourPoint(startLeftTopPoint, null));
                guideline.Curve.AddContourPoint(new ContourPoint(endLeftTopPoint, null));

                rebarSet.Guidelines.Add(guideline);
                bool succes = rebarSet.Insert();
                rebarSet.SetUserProperty("__MIN_BAR_LENGTH", 20 * Convert.ToDouble(rebarSize));
                new Model().CommitChanges();

                //Create RebarEndDetailModifier
                var innerEndDetailModifier = new RebarEndDetailModifier();
                innerEndDetailModifier.Father = rebarSet;
                innerEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
                innerEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = 10 * Convert.ToInt32(rebarSize);
                innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
                innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[2][0], null));
                innerEndDetailModifier.Insert();

                var outerEndDetailModifier = new RebarEndDetailModifier();
                outerEndDetailModifier.Father = rebarSet;
                outerEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
                outerEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = 10 * Convert.ToInt32(rebarSize);
                outerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][5], null));
                outerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[2][5], null));
                outerEndDetailModifier.Insert();
                new Model().CommitChanges();

                rebarSet.SetUserProperty(RebarCreator.FatherIDName, RebarCreator.FatherID);
                RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 1, 1 });
            }
        }
        #endregion
        #region Properties
        #endregion
        #region Fields
        static double maxHeight;
        static double minHeight;
        enum RebarType
        {
            IVR,
            OVR,
            ILR,
            OLR,
            CrPR,
            CrLR,
            CCSR,
            CLR
        }

        public static double Height1;
        public static double CorniceHeight;
        public static double BottomWidth1;
        public static double TopWidth;
        public static double CorniceWidth;
        public static double Length1;
        public static double Length2;
        public static double Height2;
        public static double BottomWidth2;
        public static double Height3;
        public static double BottomWidth3;
        #endregion
    }
}
