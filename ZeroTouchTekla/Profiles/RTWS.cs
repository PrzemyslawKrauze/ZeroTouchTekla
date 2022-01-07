using System;
using System.Collections.Generic;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model;
using System.Reflection;


namespace ZeroTouchTekla.Profiles
{
    public class RTWS : Element
    {
        #region Constructor
        public RTWS(Beam part) : base(part)
        {
            GetProfilePointsAndParameters(part);
        }
        #endregion
        #region PublicMethods
        new public void GetProfilePointsAndParameters(Beam beam)
        {
            string[] profileValues = GetProfileValues(beam);
            //RTWS Height*BottomHeight*SkewHeight*BottomWidth*TopWidth*CorniceWidth*CorniceHeight*Height2
            _height = Convert.ToDouble(profileValues[0]);
            _bottomHeight = Convert.ToDouble(profileValues[1]);
            _skewHeight = Convert.ToDouble(profileValues[2]);
            _bottomWidth = Convert.ToDouble(profileValues[3]);
            _topWidth = Convert.ToDouble(profileValues[4]);
            _corniceWidth = Convert.ToDouble(profileValues[5]);
            _corniceHeight = Convert.ToDouble(profileValues[6]);
            _length = Distance.PointToPoint(beam.StartPoint, beam.EndPoint);
            _height2 = 0;
            if (profileValues.Length > 7)
            {
                _height2 = Convert.ToDouble(profileValues[7]);
            }

            _maxHeight = _height>_height2 ?_height: _height2;
            _minHeight = _height > _height2 ? _height2 : _height;

            double fullWidth = _bottomWidth + _corniceWidth;
            double topHeight = _height - _bottomHeight - _skewHeight;

            Point p0 = new Point(0, -_height / 2.0, fullWidth / 2.0 - _corniceWidth);
            Point p1 = new Point(0, _height / 2.0 - _corniceHeight, p0.Z);
            Point p2 = new Point(0, p1.Y, p1.Z + _corniceWidth);
            Point p3 = new Point(0, p2.Y + _corniceHeight, p2.Z);
            Point p4 = new Point(0, p3.Y, p3.Z - _topWidth);
            Point p5 = new Point(0, p4.Y - topHeight, p4.Z);
            Point p7 = new Point(0, -_height / 2.0, -fullWidth / 2.0);
            Point p6 = new Point(0, -_height / 2.0 + _bottomHeight, -fullWidth / 2.0);

            List<Point> firstProfile = new List<Point> { p0, p1, p2, p3, p4, p5, p6, p7 };

            List<Point> secondProfile = new List<Point>();
            if (_height2 == 0)
            {
                foreach (Point p in firstProfile)
                {
                    Point secondPoint = new Point(p.X, p.Y, p.Z);
                    secondPoint.Translate(_length, 0, 0);
                    secondProfile.Add(secondPoint);
                }
            }
            else
            {
                Point s0 = new Point(_length, -_height / 2.0, fullWidth / 2.0 - _corniceWidth);
                Point s1 = new Point(_length, s0.Y + _height2 - _corniceHeight, s0.Z);
                Point s2 = new Point(_length, s1.Y, s1.Z + _corniceWidth);
                Point s3 = new Point(_length, s2.Y + _corniceHeight, s2.Z);
                Point s4 = new Point(_length, s3.Y, p3.Z - _topWidth);
                Point s5 = new Point(_length, p4.Y - topHeight, p4.Z);
                Point s7 = new Point(_length, -_height / 2.0, -fullWidth / 2.0);
                Point s6 = new Point(_length, -_height / 2.0 + _bottomHeight, -fullWidth / 2.0);
                secondProfile = new List<Point> { s0, s1, s2, s3, s4, s5, s6, s7 };
            }

            List<List<Point>> beamPoints = new List<List<Point>> { firstProfile, secondProfile };
            ProfilePoints = beamPoints;
            ElementFace = new ElementFace(ProfilePoints);

        }
        new public void Create()
        {
            InnerVerticalRebar();

            BottomOuterVerticalRebar();

            TopOuterVerticalRebar();

            SkewVerticalRebar();
            InnerLongitudinalRebar();

            TopOuterLongitudinalRebar();
            BottomOuterLongitudinalRebar();

            SkewLongitudinalRebar();
            CornicePerpendicularRebar();

            CorniceLongitudinalRebar();
            BottomClosingCShapeRebar(true);
            BottomClosingCShapeRebar(false);

            TopClosingCShapeRebar(true);
            TopClosingCShapeRebar(false);
            SkewClosingCShapeRebar(true);
            SkewClosingCShapeRebar(false);

            ClosingLongitudinalRebar(true);
            ClosingLongitudinalRebar(false);

            BottomCShapeRebar();
            TopCShapeRebar();

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
                    BottomOuterVerticalRebar();
                    break;
                case RebarType.ILR:
                    InnerLongitudinalRebar();
                    break;
                case RebarType.TOLR:
                    TopOuterLongitudinalRebar();
                    break;
                case RebarType.CrPR:
                    CornicePerpendicularRebar();
                    break;
                case RebarType.CrLR:
                    CorniceLongitudinalRebar();
                    break;
                case RebarType.CCSR:
                    BottomClosingCShapeRebar(true);
                    BottomClosingCShapeRebar(false);
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

            Point startp3bis = new Point(ProfilePoints[0][1].X, ProfilePoints[0][1].Y + _corniceHeight, ProfilePoints[0][1].Z);
            Point endp3bis = new Point(ProfilePoints[1][1].X, ProfilePoints[1][1].Y + _corniceHeight, ProfilePoints[1][1].Z);

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(endp3bis, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(startp3bis, null));
            rebarSet.LegFaces.Add(mainFace);

            Point offsetedStartPoint = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y, ProfilePoints[0][0].Z + 1000);
            Point offsetedEndPoint = new Point(ProfilePoints[1][0].X, ProfilePoints[1][0].Y, ProfilePoints[1][0].Z + 1000);

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

                Point startBottomPoint = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y + spliterOffset, ProfilePoints[0][0].Z);
                Point endBottomPoint = new Point(ProfilePoints[1][0].X, ProfilePoints[1][0].Y + spliterOffset, ProfilePoints[1][0].Z);

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

                Point endTopPoint = new Point(ProfilePoints[1][0].X, ProfilePoints[1][0].Y, ProfilePoints[1][0].Z);
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
                    propertyModifier.Insert();
                }
                new Model().CommitChanges();
            }

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 3 });
        }
        void BottomOuterVerticalRebar()
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
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][7], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][7], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][6], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][6], null));
            rebarSet.LegFaces.Add(mainFace);

            Point offsetedStartPoint = new Point(ProfilePoints[0][7].X, ProfilePoints[0][7].Y, ProfilePoints[0][7].Z - 40 * Convert.ToInt32(rebarSize));
            Point offsetedEndPoint = new Point(ProfilePoints[1][7].X, ProfilePoints[1][7].Y, ProfilePoints[1][7].Z - 40 * Convert.ToInt32(rebarSize));

            var bottomFace = new RebarLegFace();
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][7], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][7], null));
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

            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][7], null));
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][7], null));

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

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 3 });
        }
        void TopOuterVerticalRebar()
        {
            string rebarSize = Program.ExcelDictionary["OVR_SecondDiameter"];
            string spacing = Program.ExcelDictionary["OVR_Spacing"];

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

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();

            var innerEndDetailModifier = new RebarEndDetailModifier();
            innerEndDetailModifier.Father = rebarSet;
            innerEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.END_OFFSET;
            innerEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = 40 * Convert.ToInt32(rebarSize);
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][5], null));
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][5], null));
            innerEndDetailModifier.Insert();
            new Model().CommitChanges();

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1 });
        }
        void SkewVerticalRebar()
        {
            string rebarSize = Program.ExcelDictionary["OVR_SecondDiameter"];
            string spacing = Program.ExcelDictionary["OVR_Spacing"];

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "RTW_OVR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            var topFace = new RebarLegFace();
            topFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][5], null));
            topFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][5], null));
            topFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
            topFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][4], null));
            rebarSet.LegFaces.Add(topFace);

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][6], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][6], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][5], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][5], null));
            rebarSet.LegFaces.Add(mainFace);

            var bottomFace = new RebarLegFace();
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][7], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][7], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][6], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][6], null));
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

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();

            var topEndDetailModifier = new RebarEndDetailModifier();
            topEndDetailModifier.Father = rebarSet;
            topEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            topEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = 40 * Convert.ToInt32(rebarSize);
            topEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][4], null));
            topEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
            topEndDetailModifier.Insert();

            var bottomEndDetailModifier = new RebarEndDetailModifier();
            bottomEndDetailModifier.Father = rebarSet;
            bottomEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            bottomEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = 40 * Convert.ToInt32(rebarSize);
            bottomEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][7], null));
            bottomEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][7], null));
            bottomEndDetailModifier.Insert();

            new Model().CommitChanges();

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 1, 1 });
        }
        void InnerLongitudinalRebar()
        {
            string rebarSize = Program.ExcelDictionary["ILR_SecondDiameter"];
            string secondRebarSize = Program.ExcelDictionary["ILR_Diameter"];
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
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
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

            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));

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

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2 });
        }
        void TopOuterLongitudinalRebar()
        {
            string rebarSize = Program.ExcelDictionary["OLR_SecondDiameter"];
            string spacing = Program.ExcelDictionary["OLR_Spacing"];
            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "RTW_OLR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            var mianFace = new RebarLegFace();
            mianFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][5], null));
            mianFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][5], null));
            mianFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
            mianFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][4], null));
            rebarSet.LegFaces.Add(mianFace);

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
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][4], null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2 });
        }
        void SkewLongitudinalRebar()
        {
            string rebarSize = Program.ExcelDictionary["OLR_SecondDiameter"];
            string spacing = Program.ExcelDictionary["OLR_Spacing"];
            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "RTW_OLR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            var mianFace = new RebarLegFace();
            mianFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][6], null));
            mianFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][6], null));
            mianFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][5], null));
            mianFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][5], null));
            rebarSet.LegFaces.Add(mianFace);

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

            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][6], null));
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][5], null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2 });
        }
        void BottomOuterLongitudinalRebar()
        {
            string rebarSize = Program.ExcelDictionary["OLR_Diameter"];
            string spacing = Program.ExcelDictionary["OLR_Spacing"];
            double startOffset = Convert.ToDouble(Program.ExcelDictionary["OLR_StartOffset"]);
            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "RTW_OLR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            var mianFace = new RebarLegFace();
            mianFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][7], null));
            mianFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][7], null));
            mianFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][6], null));
            mianFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][6], null));
            rebarSet.LegFaces.Add(mianFace);

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

            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][7], null));
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][6], null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2 });
        }
        void CornicePerpendicularRebar()
        {
            string rebarSize = Program.ExcelDictionary["CPR_Diameter"];
            string spacing = Program.ExcelDictionary["CPR_Spacing"];
            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "RTW_CrPR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][3], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            rebarSet.LegFaces.Add(mainFace);

            var topFace = new RebarLegFace();
            topFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            topFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][3], null));
            topFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
            topFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][4], null));
            rebarSet.LegFaces.Add(topFace);

            var outerFace = new RebarLegFace();
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][4], null));
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][5], null));
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][5], null));
            rebarSet.LegFaces.Add(outerFace);

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

            var hookModifier = new RebarEndDetailModifier();
            hookModifier.Father = rebarSet;
            hookModifier.EndType = RebarEndDetailModifier.EndTypeEnum.HOOK;
            hookModifier.RebarHook.Shape = RebarHookData.RebarHookShapeEnum.HOOK_90_DEGREES;
            hookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            hookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
            hookModifier.Insert();
            new Model().CommitChanges();

            var bottomLengthModifier = new RebarEndDetailModifier();
            bottomLengthModifier.Father = rebarSet;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentLength = 40 * Convert.ToInt32(rebarSize);
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][5], null));
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][5], null));
            bottomLengthModifier.Insert();

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 1, 1 });
        }
        void CorniceLongitudinalRebar()
        {
            string rebarSize = Program.ExcelDictionary["ILR_SecondDiameter"];
            string spacing = Program.ExcelDictionary["ILR_Spacing"];
            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "RTW_CrLR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][3], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
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
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            rebarSet.Guidelines.Add(guideline);

            if (_height2 != 0)
            {
                var secondaryGuideLine = new RebarGuideline();
                secondaryGuideLine.Spacing.InheritFromPrimary = true;
                secondaryGuideLine.Spacing.StartOffset = 100;
                secondaryGuideLine.Spacing.EndOffset = 100;
                secondaryGuideLine.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
                secondaryGuideLine.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][3], null));
                rebarSet.Guidelines.Add(secondaryGuideLine);
            }

            bool succes = rebarSet.Insert();
            new Model().CommitChanges();

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2 });
        }
        void BottomClosingCShapeRebar(bool isStart)
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

            Point leftBottom, rightBottom, rightTop, leftTop;
            Point endLeftBottom, endRightBottom, endRightTop, endLeftTop;
            if (isStart)
            {
                leftBottom = ProfilePoints[0][0];
                rightBottom = ProfilePoints[0][7];
                rightTop = ProfilePoints[0][6];
                leftTop = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y + _bottomHeight, ProfilePoints[0][0].Z);
                endLeftBottom = ProfilePoints[1][0];
                endRightBottom = ProfilePoints[1][7];
                endRightTop = ProfilePoints[1][6];
                endLeftTop = new Point(ProfilePoints[1][0].X, ProfilePoints[1][0].Y + _bottomHeight, ProfilePoints[1][0].Z);
            }
            else
            {
                leftBottom = ProfilePoints[1][0];
                rightBottom = ProfilePoints[1][7];
                rightTop = ProfilePoints[1][6];
                leftTop = new Point(ProfilePoints[1][0].X, ProfilePoints[1][0].Y + _bottomHeight, ProfilePoints[1][0].Z);
                endLeftBottom = ProfilePoints[0][0];
                endRightBottom = ProfilePoints[0][7];
                endRightTop = ProfilePoints[0][6];
                endLeftTop = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y + _bottomHeight, ProfilePoints[0][0].Z);
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
            guideline.Spacing.EndOffset = Convert.ToDouble(spacing) / 2.0;

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
            innerEndDetailModifier.Insert();

            var outerEndDetailModifier = new RebarEndDetailModifier();
            outerEndDetailModifier.Father = rebarSet;
            outerEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            outerEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = 10 * Convert.ToInt32(rebarSize);
            outerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(endRightBottom, null));
            outerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(endRightTop, null));
            outerEndDetailModifier.Insert();
            new Model().CommitChanges();

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 2, 2 });
        }
        void SkewClosingCShapeRebar(bool isStart)
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
                leftBottom = new Point(ProfilePoints[0][0].X, ProfilePoints[0][6].Y, ProfilePoints[0][0].Z);
                rightBottom = ProfilePoints[0][6];
                rightTop = ProfilePoints[0][5];
                leftTop = new Point(ProfilePoints[0][0].X, ProfilePoints[0][5].Y, ProfilePoints[0][0].Z);
                endLeftBottom = new Point(ProfilePoints[1][0].X, ProfilePoints[1][6].Y, ProfilePoints[1][0].Z);
                endRightBottom = ProfilePoints[1][6];
                endRightTop = ProfilePoints[1][5];
                endLeftTop = new Point(ProfilePoints[1][0].X, ProfilePoints[1][5].Y, ProfilePoints[1][0].Z);
            }
            else
            {
                leftBottom = new Point(ProfilePoints[1][0].X, ProfilePoints[1][6].Y, ProfilePoints[1][0].Z);
                rightBottom = ProfilePoints[1][6];
                rightTop = ProfilePoints[1][5];
                leftTop = new Point(ProfilePoints[1][0].X, ProfilePoints[1][5].Y, ProfilePoints[1][0].Z);
                endLeftBottom = new Point(ProfilePoints[0][0].X, ProfilePoints[0][6].Y, ProfilePoints[0][0].Z);
                endRightBottom = ProfilePoints[0][6];
                endRightTop = ProfilePoints[0][5];
                endLeftTop = new Point(ProfilePoints[0][0].X, ProfilePoints[0][5].Y, ProfilePoints[0][0].Z);
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
            guideline.Spacing.StartOffset = Convert.ToDouble(spacing) / 2.0;
            guideline.Spacing.EndOffset = Convert.ToDouble(spacing) / 2.0;

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
            innerEndDetailModifier.Insert();

            var outerEndDetailModifier = new RebarEndDetailModifier();
            outerEndDetailModifier.Father = rebarSet;
            outerEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            outerEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = 10 * Convert.ToInt32(rebarSize);
            outerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(endRightBottom, null));
            outerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(endRightTop, null));
            outerEndDetailModifier.Insert();
            new Model().CommitChanges();

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 2, 2 });
        }
        void TopClosingCShapeRebar(bool isStart)
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
                leftBottom = new Point(ProfilePoints[0][0].X, ProfilePoints[0][5].Y, ProfilePoints[0][0].Z);
                rightBottom = ProfilePoints[0][5];
                rightTop = ProfilePoints[0][4];
                leftTop = new Point(ProfilePoints[0][0].X, ProfilePoints[0][4].Y, ProfilePoints[0][0].Z);
                endLeftBottom = new Point(ProfilePoints[1][0].X, ProfilePoints[1][5].Y, ProfilePoints[1][0].Z);
                endRightBottom = ProfilePoints[1][5];
                endRightTop = ProfilePoints[1][4];
                endLeftTop = new Point(ProfilePoints[1][0].X, ProfilePoints[1][4].Y, ProfilePoints[1][0].Z);
            }
            else
            {
                leftBottom = new Point(ProfilePoints[1][0].X, ProfilePoints[1][5].Y, ProfilePoints[1][0].Z);
                rightBottom = ProfilePoints[1][5];
                rightTop = ProfilePoints[1][4];
                leftTop = new Point(ProfilePoints[1][0].X, ProfilePoints[1][4].Y, ProfilePoints[1][0].Z);
                endLeftBottom = new Point(ProfilePoints[0][0].X, ProfilePoints[0][5].Y, ProfilePoints[0][0].Z);
                endRightBottom = ProfilePoints[0][5];
                endRightTop = ProfilePoints[0][4];
                endLeftTop = new Point(ProfilePoints[0][0].X, ProfilePoints[0][4].Y, ProfilePoints[0][0].Z);
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
            guideline.Spacing.StartOffset = Convert.ToDouble(spacing) / 2.0;
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
            innerEndDetailModifier.Insert();

            var outerEndDetailModifier = new RebarEndDetailModifier();
            outerEndDetailModifier.Father = rebarSet;
            outerEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            outerEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = 10 * Convert.ToInt32(rebarSize);
            outerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(endRightBottom, null));
            outerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(endRightTop, null));
            outerEndDetailModifier.Insert();
            new Model().CommitChanges();

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
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

            Point leftBottom, rightBottom, rightTop, leftTop, rightMid1, rightMid2; ;
            if (isStart)
            {
                leftBottom = ProfilePoints[0][0];
                rightBottom = ProfilePoints[0][7];
                rightMid1 = ProfilePoints[0][6];
                rightMid2 = ProfilePoints[0][5];
                rightTop = ProfilePoints[0][4];
                leftTop = new Point(ProfilePoints[0][0].X, ProfilePoints[0][4].Y, ProfilePoints[0][0].Z);
            }
            else
            {
                leftBottom = ProfilePoints[1][0];
                rightBottom = ProfilePoints[1][7];
                rightMid1 = ProfilePoints[1][6];
                rightMid2 = ProfilePoints[1][5];
                rightTop = ProfilePoints[1][4];
                leftTop = new Point(ProfilePoints[1][0].X, ProfilePoints[1][4].Y, ProfilePoints[1][0].Z);
            }

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(leftBottom, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(rightBottom, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(rightMid1, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(rightMid2, null));
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

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2 });
        }
        void BottomCShapeRebar()
        {
            string rebarSize = Program.ExcelDictionary["CSR_Diameter"];
            string horizontalSpacing = Program.ExcelDictionary["CSR_HorizontalSpacing"];
            string verticalSpacing = Program.ExcelDictionary["CSR_VerticalSpacing"];
            double startOffset = Convert.ToDouble(Program.ExcelDictionary["OLR_StartOffset"]);

            double height = _bottomHeight;
            double correctedHeight = height - startOffset - 10 * Convert.ToInt32(rebarSize);
            int correctedNumberOfRows = (int)Math.Floor(correctedHeight / Convert.ToDouble(verticalSpacing)) + 1;
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
                Point startRightTopPoint = new Point(ProfilePoints[0][7].X, ProfilePoints[0][7].Y + newoffset, ProfilePoints[0][7].Z);
                Point endRightTopPoint = new Point(ProfilePoints[1][7].X, ProfilePoints[1][7].Y + newoffset, ProfilePoints[1][7].Z);

                var mainFace = new RebarLegFace();
                mainFace.Contour.AddContourPoint(new ContourPoint(startLeftTopPoint, null));
                mainFace.Contour.AddContourPoint(new ContourPoint(endLeftTopPoint, null));
                mainFace.Contour.AddContourPoint(new ContourPoint(endRightTopPoint, null));
                mainFace.Contour.AddContourPoint(new ContourPoint(startRightTopPoint, null));
                rebarSet.LegFaces.Add(mainFace);

                var innerFace = new RebarLegFace();
                innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
                innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
                innerFace.Contour.AddContourPoint(new ContourPoint(endLeftTopPoint, null));
                innerFace.Contour.AddContourPoint(new ContourPoint(startLeftTopPoint, null));
                rebarSet.LegFaces.Add(innerFace);

                var outerFace = new RebarLegFace();
                outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][7], null));
                outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][7], null));
                outerFace.Contour.AddContourPoint(new ContourPoint(endRightTopPoint, null));
                outerFace.Contour.AddContourPoint(new ContourPoint(startRightTopPoint, null));
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
                outerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][7], null));
                outerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][7], null));
                outerEndDetailModifier.Insert();
                new Model().CommitChanges();

                PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
                RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 1, 1 });
            }
        }
        void TopCShapeRebar()
        {
            string rebarSize = Program.ExcelDictionary["CSR_Diameter"];
            string horizontalSpacing = Program.ExcelDictionary["CSR_HorizontalSpacing"];
            string verticalSpacing = Program.ExcelDictionary["CSR_VerticalSpacing"];

            double height = _minHeight - _bottomHeight - _skewHeight;
            double correctedHeight = height - 10 * Convert.ToInt32(rebarSize);
            int correctedNumberOfRows = (int)Math.Floor(correctedHeight / Convert.ToDouble(verticalSpacing)) + 1;
            double offset = 10 * Convert.ToInt32(rebarSize);

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

                Point startLeftTopPoint = new Point(ProfilePoints[0][0].X, ProfilePoints[0][5].Y + newoffset, ProfilePoints[0][0].Z);
                Point endLeftTopPoint = new Point(ProfilePoints[1][0].X, ProfilePoints[1][5].Y + newoffset, ProfilePoints[1][0].Z);
                Point startRightTopPoint = new Point(ProfilePoints[0][5].X, ProfilePoints[0][5].Y + newoffset, ProfilePoints[0][5].Z);
                Point endRightTopPoint = new Point(ProfilePoints[1][5].X, ProfilePoints[1][5].Y + newoffset, ProfilePoints[1][5].Z);

                var mainFace = new RebarLegFace();
                mainFace.Contour.AddContourPoint(new ContourPoint(startLeftTopPoint, null));
                mainFace.Contour.AddContourPoint(new ContourPoint(endLeftTopPoint, null));
                mainFace.Contour.AddContourPoint(new ContourPoint(endRightTopPoint, null));
                mainFace.Contour.AddContourPoint(new ContourPoint(startRightTopPoint, null));
                rebarSet.LegFaces.Add(mainFace);

                var innerFace = new RebarLegFace();
                innerFace.Contour.AddContourPoint(new ContourPoint(new Point(ProfilePoints[0][0].X, ProfilePoints[0][5].Y, ProfilePoints[0][0].Z), null));
                innerFace.Contour.AddContourPoint(new ContourPoint(new Point(ProfilePoints[1][0].X, ProfilePoints[1][5].Y + newoffset, ProfilePoints[1][0].Z), null));
                innerFace.Contour.AddContourPoint(new ContourPoint(endLeftTopPoint, null));
                innerFace.Contour.AddContourPoint(new ContourPoint(startLeftTopPoint, null));
                rebarSet.LegFaces.Add(innerFace);

                var outerFace = new RebarLegFace();
                outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][5], null));
                outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][5], null));
                outerFace.Contour.AddContourPoint(new ContourPoint(endRightTopPoint, null));
                outerFace.Contour.AddContourPoint(new ContourPoint(startRightTopPoint, null));
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
                innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(new Point(ProfilePoints[0][0].X, ProfilePoints[0][5].Y, ProfilePoints[0][0].Z), null));
                innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(new Point(ProfilePoints[1][0].X, ProfilePoints[1][5].Y, ProfilePoints[1][0].Z), null));
                innerEndDetailModifier.Insert();

                var outerEndDetailModifier = new RebarEndDetailModifier();
                outerEndDetailModifier.Father = rebarSet;
                outerEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
                outerEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = 10 * Convert.ToInt32(rebarSize);
                outerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][5], null));
                outerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][5], null));
                outerEndDetailModifier.Insert();
                new Model().CommitChanges();

                PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
                RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 1, 1 });
            }
        }
        #endregion
        #region Properties
        public double TopWidth { get { return _topWidth; } }
        public double CorniceWidth { get { return _corniceWidth; } }
        public double BottomWidth { get { return _bottomWidth; } }
        public double BottomHeight { get { return _bottomHeight; } }
        public double SkewHeight { get { return _skewHeight; } }
        public double CorniceHeight { get { return _corniceHeight; } }
        public double Height { get { return _height; } }
        public double Height2 { get { return _height2; } }
        public double Length { get { return _length; } }
        public double MaxHeight { get { return _maxHeight; } }
        public double MinHeight { get { return _minHeight; } }
        #endregion
        #region Fields
        enum RebarType
        {
            IVR,
            OVR,
            ILR,
            TOLR,
            BOLR,
            CrPR,
            CrLR,
            CCSR,
            CLR
        }
        static double _topWidth;
        static double _corniceWidth;
        static double _bottomWidth;
        static double _bottomHeight;
        static double _skewHeight;
        static double _corniceHeight;
        static double _height;
        static double _height2;
        static double _length;
        static double _maxHeight;
        static double _minHeight;
        #endregion
    }
}
