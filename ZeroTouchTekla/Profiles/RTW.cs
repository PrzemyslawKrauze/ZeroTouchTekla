using System;
using System.Collections.Generic;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model;
using System.Reflection;

namespace ZeroTouchTekla.Profiles
{
    public class RTW : Element
    {
        #region Fields
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
        double _height;
        double _corniceHeight;
        double _bottomWidth;
        double _topWidth;
        double _corniceWidth;
        double _length;
        double _height2;
        #endregion
        #region Constructors
        public RTW(Part part) : base(part)
        {
            base.BasePart = part;
            SetLocalPlane();
            SetFields();
            SetProfilePoints();
        }
        #endregion
        #region Properties
        public double Height { get { return _height; } }
        public double CorniceHeight { get { return _corniceHeight; } }
        public double BottomWidth { get { return _bottomWidth; } }
        public double TopWidth { get { return _topWidth; } }
        public double CorniceWidth { get { return _corniceWidth; } }
        public double Length { get { return _length; } }
        public double Height2 { get { return _height2; } }
        #endregion
        #region PublicMethods
        private void SetFields()
        {
            Beam beam = base.BasePart as Beam;
            //RTW Height*CorniceHeight*BottomWidth*TopWidth*CorniceWidth
            //RTWVR Height*CorniceHeight*BottomWidth*TopWidth*CorniceWidth*Height2
            string[] profileValues = GetProfileValues(beam);            
            _height = Convert.ToDouble(profileValues[0]);
            _corniceHeight = Convert.ToDouble(profileValues[1]);
            _bottomWidth = Convert.ToDouble(profileValues[2]);
            _topWidth = Convert.ToDouble(profileValues[3]);
            _corniceWidth = Convert.ToDouble(profileValues[4]);
            _length = Distance.PointToPoint(beam.StartPoint, beam.EndPoint);
           
            if (profileValues.Length > 5)
            {
                _height2 = Convert.ToDouble(profileValues[5]);
            }
            else
            {
                _height2 = _height;
            }
        }
        private void SetProfilePoints()
        {            
            double hToW = (BottomWidth - (TopWidth - CorniceWidth)) / Height;
            double bottomWidth2 = hToW * Height2 + (TopWidth - CorniceWidth);
            double distanceToMid = Height > Height2 ? Height / 2.0 : Height2 / 2.0;
            double fullWidth = _corniceWidth + (bottomWidth2>_bottomWidth?bottomWidth2:_bottomWidth);

            Point p0 = new Point(0, -distanceToMid, fullWidth / 2.0 - _corniceWidth);
            Point p1 = new Point(0, p0.Y+Height - CorniceHeight, p0.Z);
            Point p2 = new Point(0, p1.Y, p1.Z + CorniceWidth);
            Point p3 = new Point(0, p2.Y + CorniceHeight, p2.Z);
            Point p4 = new Point(0, p3.Y, p3.Z - TopWidth);
            Point p5 = new Point(0, -distanceToMid, p0.Z-_bottomWidth);

            List<Point> firstProfile = new List<Point> { p0, p1, p2, p3, p4, p5 };

            List<Point> secondProfile = new List<Point>();
            if (Height2 == Height)
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
                Point s0 = new Point(_length, -distanceToMid, fullWidth / 2.0 - _corniceWidth);
                Point s1 = new Point(_length, s0.Y + Height2 - CorniceHeight, s0.Z);
                Point s2 = new Point(_length, s1.Y, p1.Z + CorniceWidth);
                Point s3 = new Point(_length, s2.Y + CorniceHeight, s2.Z);
                Point s4 = new Point(_length, s3.Y, s3.Z - TopWidth);
                Point s5 = new Point(_length, -distanceToMid, s0.Z - bottomWidth2);
                secondProfile = new List<Point> { s0, s1, s2, s3, s4, s5 };
            }
            List<List<Point>> beamPoints = new List<List<Point>> { firstProfile, secondProfile };
            ProfilePoints = beamPoints;
        }
        public override void Create()
        {
            OuterVerticalRebar();
            InnerVerticalRebar();
            OuterLongitudinalRebar();
            InnerLongitudinalRebar();
            CornicePerpendicularRebar();
            CorniceLongitudinalRebar();
            ClosingCShapeRebar(0);
            ClosingCShapeRebar(1);
            ClosingLongitudinalRebar(0);
            ClosingLongitudinalRebar(1);
            CShapeRebar();
        }
        public override void CreateSingle(string rebarName)
        {
            rebarName = rebarName.Split('_')[1];
            RebarType rType;
            Enum.TryParse(rebarName, out rType);
            switch (rType)
            {
                case RebarType.IVR:
                    OuterVerticalRebar();
                    break;
                case RebarType.OVR:
                    InnerVerticalRebar();
                    break;
                case RebarType.ILR:
                    OuterLongitudinalRebar();
                    break;
                case RebarType.OLR:
                    InnerLongitudinalRebar();
                    break;
                case RebarType.CrPR:
                    CornicePerpendicularRebar();
                    break;
                case RebarType.CrLR:
                    CorniceLongitudinalRebar();
                    break;
                case RebarType.CCSR:
                    ClosingCShapeRebar(0);
                    ClosingCShapeRebar(1);
                    break;
                case RebarType.CLR:
                    ClosingLongitudinalRebar(0);
                    ClosingLongitudinalRebar(1);
                    break;
            }
        }
        #endregion
        #region PrivateMethods
        void OuterVerticalRebar()
        {
            int rebarSize = Convert.ToInt32(Program.ExcelDictionary["OVR_Diameter"]);
            double spacing =Convert.ToDouble(Program.ExcelDictionary["OVR_Spacing"]);
            int addSplitter = Convert.ToInt32(Program.ExcelDictionary["OVR_AddSplitter"]);
            int secondRebarSize = Convert.ToInt32(Program.ExcelDictionary["OVR_SecondDiameter"]);
            double spliterOffset = Convert.ToDouble(Program.ExcelDictionary["OVR_SplitterOffset"]) + Convert.ToDouble(rebarSize) * 20;

            var rebarSet = TeklaUtils.CreateDefaultRebarSet("RTW_OVR", rebarSize);

            Point startp3bis = new Point(ProfilePoints[0][1].X, ProfilePoints[0][1].Y + CorniceHeight, ProfilePoints[0][1].Z);
            Point endp3bis = new Point(ProfilePoints[1][1].X, ProfilePoints[1][1].Y + CorniceHeight, ProfilePoints[1][1].Z);

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
                    propertyModifier.RebarProperties.Size = secondRebarSize.ToString();
                    propertyModifier.RebarProperties.Class = TeklaUtils.SetClass(secondRebarSize);
                    propertyModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
                    propertyModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
                    propertyModifier.Insert();
                }
                new Model().CommitChanges();
            }

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
            LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 3 });
        }
        void InnerVerticalRebar()
        {
            int rebarSize = Convert.ToInt32(Program.ExcelDictionary["IVR_Diameter"]);
            string spacing = Program.ExcelDictionary["IVR_Spacing"];
            int addSplitter = Convert.ToInt32(Program.ExcelDictionary["IVR_AddSplitter"]);
            int secondRebarSize = Convert.ToInt32(Program.ExcelDictionary["IVR_SecondDiameter"]);
            double spliterOffset = Convert.ToDouble(Program.ExcelDictionary["IVR_SplitterOffset"]) + Convert.ToDouble(rebarSize) * 20;

            var rebarSet = TeklaUtils.CreateDefaultRebarSet("RTW_IVR", rebarSize);

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][5], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][5], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][4], null));
            rebarSet.LegFaces.Add(mainFace);

            Point offsetedStartPoint = new Point(ProfilePoints[0][5].X, ProfilePoints[0][5].Y, ProfilePoints[0][5].Z - 40 * Convert.ToInt32(rebarSize));
            Point offsetedEndPoint = new Point(ProfilePoints[1][5].X, ProfilePoints[1][5].Y, ProfilePoints[1][5].Z - 40 * Convert.ToInt32(rebarSize));

            var bottomFace = new RebarLegFace();
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][5], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][5], null));
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

                GeometricPlane facePlane = Utility.GetPlaneFromFace(mainFace);

                Point offsetedStartP5 = new Point(ProfilePoints[0][5].X, ProfilePoints[0][5].Y, ProfilePoints[0][5].Z);
                offsetedStartP5.Translate(0, spliterOffset, 0);
                Point startP5Bis = new Point(offsetedStartP5.X, offsetedStartP5.Y, offsetedStartP5.Z - BottomWidth * 2);
                Line startline = new Line(offsetedStartP5, startP5Bis);

                Point offsetedEndP5 = new Point(ProfilePoints[1][5].X, ProfilePoints[1][5].Y, ProfilePoints[1][5].Z);
                offsetedEndP5.Translate(0, spliterOffset, 0);
                Point endP5Bis = new Point(offsetedEndP5.X, offsetedEndP5.Y, offsetedEndP5.Z - BottomWidth * 2);
                Line endline = new Line(offsetedEndP5, endP5Bis);

                Point startIntersection = Intersection.LineToPlane(startline, facePlane);
                Point endIntersection = Intersection.LineToPlane(endline, facePlane);

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

                Point offsetedStartTopP5 = new Point(ProfilePoints[0][5].X, ProfilePoints[0][5].Y, ProfilePoints[0][5].Z);
                offsetedStartTopP5.Translate(0, spliterOffset + 1.3 * 40 * Convert.ToDouble(rebarSize), 0);
                Point startTopP5Bis = new Point(offsetedStartTopP5.X, offsetedStartTopP5.Y, offsetedStartTopP5.Z - BottomWidth * 2);
                Line startTopline = new Line(offsetedStartTopP5, startTopP5Bis);

                Point offsetedTopEndP5 = new Point(ProfilePoints[1][5].X, ProfilePoints[1][5].Y, ProfilePoints[1][5].Z);
                offsetedTopEndP5.Translate(0, spliterOffset + 1.3 * 40 * Convert.ToDouble(rebarSize), 0);
                Point endTopP5Bis = new Point(offsetedTopEndP5.X, offsetedTopEndP5.Y, offsetedTopEndP5.Z - BottomWidth * 2);
                Line endTopLine = new Line(offsetedTopEndP5, endTopP5Bis);

                Point startTopIntersection = Intersection.LineToPlane(startTopline, facePlane);
                Point endTopIntersection = Intersection.LineToPlane(endTopLine, facePlane);

                topSpliter.Curve.AddContourPoint(new ContourPoint(startTopIntersection, null));
                topSpliter.Curve.AddContourPoint(new ContourPoint(endTopIntersection, null));
                topSpliter.Insert();

                if (rebarSize != secondRebarSize)
                {
                    var propertyModifier = new RebarPropertyModifier();
                    propertyModifier.Father = rebarSet;
                    propertyModifier.BarsAffected = BaseRebarModifier.BarsAffectedEnum.ALL_BARS;
                    propertyModifier.RebarProperties.Size = secondRebarSize.ToString();
                    propertyModifier.RebarProperties.Class = TeklaUtils.SetClass(Convert.ToDouble(secondRebarSize));
                    propertyModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][4], null));
                    propertyModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
                    propertyModifier.Insert();
                }
                new Model().CommitChanges();
            }

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod()); ;
            LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 3 });
        }
        void OuterLongitudinalRebar()
        {
            int rebarSize = Convert.ToInt32(Program.ExcelDictionary["OLR_SecondDiameter"]);
            int secondRebarSize = Convert.ToInt32(Program.ExcelDictionary["OLR_Diameter"]);
            double spacing = Convert.ToDouble(Program.ExcelDictionary["OLR_Spacing"]);
            double startOffset = Convert.ToDouble(Program.ExcelDictionary["OLR_StartOffset"]);
            double firstLength = Convert.ToDouble(Program.ExcelDictionary["OLR_SecondDiameterLength"]);
            var rebarSet = TeklaUtils.CreateDefaultRebarSet("RTW_OLR", rebarSize);

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
                propertyModifier.RebarProperties.Size = secondRebarSize.ToString();
                propertyModifier.RebarProperties.Class = TeklaUtils.SetClass(Convert.ToDouble(secondRebarSize));

                Point secondPoint = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y + startOffset + firstLength, ProfilePoints[0][0].Z);
                propertyModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
                propertyModifier.Curve.AddContourPoint(new ContourPoint(secondPoint, null));
                propertyModifier.Insert();
            }
            new Model().CommitChanges();

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
            LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2 });
        }
        void InnerLongitudinalRebar()
        {
            int rebarSize = Convert.ToInt32(Program.ExcelDictionary["ILR_SecondDiameter"]);
            int secondRebarSize = Convert.ToInt32(Program.ExcelDictionary["ILR_Diameter"]);
            string spacing = Program.ExcelDictionary["ILR_Spacing"];
            double startOffset = Convert.ToDouble(Program.ExcelDictionary["ILR_StartOffset"]);
            double firstLength = Convert.ToDouble(Program.ExcelDictionary["ILR_SecondDiameterLength"]);
            var rebarSet = TeklaUtils.CreateDefaultRebarSet("RTW_ILR", rebarSize);

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
            guideline.Spacing.StartOffset = startOffset;
            guideline.Spacing.EndOffset = 100;

            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][5], null));
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][4], null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            if (rebarSize != secondRebarSize)
            {
                var propertyModifier = new RebarPropertyModifier();
                propertyModifier.Father = rebarSet;
                propertyModifier.BarsAffected = BaseRebarModifier.BarsAffectedEnum.ALL_BARS;
                propertyModifier.RebarProperties.Size = secondRebarSize.ToString();
                propertyModifier.RebarProperties.Class = TeklaUtils.SetClass(Convert.ToDouble(secondRebarSize));
                Point origin = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y + startOffset + firstLength, ProfilePoints[0][0].Z);
                GeometricPlane plane = new GeometricPlane(origin, new Vector(0, 1, 0));
                Line line = new Line(ProfilePoints[0][5], ProfilePoints[0][4]);
                Point intersection = Utility.GetExtendedIntersection(line, plane, 1);

                propertyModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][5], null));
                propertyModifier.Curve.AddContourPoint(new ContourPoint(intersection, null));
                propertyModifier.Insert();
            }

            new Model().CommitChanges();

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
            LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2 });
        }
        void CornicePerpendicularRebar()
        {
            int rebarSize = Convert.ToInt32(Program.ExcelDictionary["CPR_Diameter"]);
            string spacing = Program.ExcelDictionary["CPR_Spacing"];
            var rebarSet = TeklaUtils.CreateDefaultRebarSet("RTW_CrPR", rebarSize);

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

            Vector xAxis = Utility.GetVectorFromTwoPoints(ProfilePoints[0][4], ProfilePoints[1][4]);
            Vector yAxis = Utility.GetVectorFromTwoPoints(ProfilePoints[0][4], ProfilePoints[0][3]);
            Point origin = new Point(ProfilePoints[0][4].X, ProfilePoints[0][4].Y - 1000, ProfilePoints[0][4].Z);
            GeometricPlane plane = new GeometricPlane(origin, xAxis, yAxis);
            Line startLine = new Line(ProfilePoints[0][4], ProfilePoints[0][5]);
            Line endLine = new Line(ProfilePoints[1][4], ProfilePoints[1][5]);

            Point startIntersection = Utility.GetExtendedIntersection(startLine, plane, 1);
            Point endIntersection = Utility.GetExtendedIntersection(endLine, plane, 1);
            var outerFace = new RebarLegFace();
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][4], null));
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
            outerFace.Contour.AddContourPoint(new ContourPoint(endIntersection, null));
            outerFace.Contour.AddContourPoint(new ContourPoint(startIntersection, null));
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
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(startIntersection, null));
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(endIntersection, null));
            bottomLengthModifier.Insert();

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
            LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 1, 1 });
        }
        void CorniceLongitudinalRebar()
        {
            int rebarSize = Convert.ToInt32(Program.ExcelDictionary["ILR_Diameter"]);
            string spacing = Program.ExcelDictionary["ILR_Spacing"];
            var rebarSet = TeklaUtils.CreateDefaultRebarSet("RTW_CrLR", rebarSize);

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

            if (Height2 != Height)
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
            LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2 });
        }
        void ClosingCShapeRebar(int number)
        {
            int rebarSize = Convert.ToInt32(Program.ExcelDictionary["CCSR_Diameter"]);
            string spacing = Program.ExcelDictionary["CCSR_Spacing"];
            double startOffset = Convert.ToDouble(Program.ExcelDictionary["OLR_StartOffset"]);
            var rebarSet = TeklaUtils.CreateDefaultRebarSet("RTW_CCSR", rebarSize);

            Point leftBottom, rightBottom, rightTop, leftTop;
            Point endLeftBottom, endRightBottom, endRightTop, endLeftTop;
            if (number==0)
            {
                leftBottom = ProfilePoints[0][0];
                rightBottom = ProfilePoints[0][5];
                rightTop = ProfilePoints[0][4];
                leftTop = new Point(ProfilePoints[0][1].X, ProfilePoints[0][1].Y + _corniceHeight, ProfilePoints[0][1].Z);
                endLeftBottom = ProfilePoints[1][0];
                endRightBottom = ProfilePoints[1][5];
                endRightTop = ProfilePoints[1][4];
                endLeftTop = new Point(ProfilePoints[1][1].X, ProfilePoints[1][1].Y + _corniceHeight, ProfilePoints[1][1].Z);
            }
            else
            {
                leftBottom = ProfilePoints[1][0];
                rightBottom = ProfilePoints[1][5];
                rightTop = ProfilePoints[1][4];
                leftTop = new Point(ProfilePoints[1][1].X, ProfilePoints[1][1].Y + _corniceHeight, ProfilePoints[1][1].Z);
                endLeftBottom = ProfilePoints[0][0];
                endRightBottom = ProfilePoints[0][5];
                endRightTop = ProfilePoints[0][4];
                endLeftTop = new Point(ProfilePoints[0][1].X, ProfilePoints[0][1].Y + _corniceHeight, ProfilePoints[0][1].Z);
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
            new Model().CommitChanges();

            //Create RebarEndDetailModifier
            var innerEndDetailModifier = new RebarEndDetailModifier();
            innerEndDetailModifier.Father = rebarSet;
            innerEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            innerEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = 10 * Convert.ToInt32(rebarSize);
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(endLeftBottom, null));
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(endLeftTop, null));
            if (Height2 != Height)
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
            if (Height2 != Height)
            {
                outerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(rightTop, null));
            }
            outerEndDetailModifier.Insert();
            new Model().CommitChanges();

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod(),number);
            LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 2, 2 });
        }
        void ClosingLongitudinalRebar(int number)
        {
            int rebarSize = Convert.ToInt32(Program.ExcelDictionary["CLR_Diameter"]);
            string spacing = Program.ExcelDictionary["CLR_Spacing"];
            var rebarSet = TeklaUtils.CreateDefaultRebarSet("RTW_CLR", rebarSize);

            Point leftBottom, rightBottom, rightTop, leftTop;
            if (number==0)
            {
                leftBottom = ProfilePoints[0][0];
                rightBottom = ProfilePoints[0][5];
                rightTop = ProfilePoints[0][4];
                leftTop = new Point(ProfilePoints[0][1].X, ProfilePoints[0][1].Y + _corniceHeight, ProfilePoints[0][1].Z);
            }
            else
            {
                leftBottom = ProfilePoints[1][0];
                rightBottom = ProfilePoints[1][5];
                rightTop = ProfilePoints[1][4];
                leftTop = new Point(ProfilePoints[1][1].X, ProfilePoints[1][1].Y + _corniceHeight, ProfilePoints[1][1].Z);
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

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod(), number);
            LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2 });
        }
        void CShapeRebar()
        {
            int rebarSize = Convert.ToInt32(Program.ExcelDictionary["CSR_Diameter"]);
            string horizontalSpacing = Program.ExcelDictionary["CSR_HorizontalSpacing"];
            string verticalSpacing = Program.ExcelDictionary["CSR_VerticalSpacing"];
            double startOffset = Convert.ToDouble(Program.ExcelDictionary["OLR_StartOffset"]);

            double correctedHeight = Height - startOffset - CorniceHeight - 10 * Convert.ToInt32(rebarSize);
            int correctedNumberOfRows = (int)Math.Floor(correctedHeight / Convert.ToDouble(verticalSpacing));
            double offset = startOffset + 10 * Convert.ToInt32(rebarSize);

            for (int i = 0; i < correctedNumberOfRows; i++)
            {
                double newoffset = offset + i * Convert.ToDouble(verticalSpacing);
                var rebarSet = TeklaUtils.CreateDefaultRebarSet("RTW_CSR", rebarSize);

                Point startLeftTopPoint = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y + newoffset, ProfilePoints[0][0].Z);
                Point endLeftTopPoint = new Point(ProfilePoints[1][0].X, ProfilePoints[1][0].Y + newoffset, ProfilePoints[1][0].Z);

                Point tempSLP = new Point(startLeftTopPoint.X, startLeftTopPoint.Y, startLeftTopPoint.Z - _bottomWidth * 2);
                Point tempELP = new Point(endLeftTopPoint.X, endLeftTopPoint.Y, endLeftTopPoint.Z - _bottomWidth * 2);

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
                guideline.Spacing.EndOffset = 100;

                guideline.Curve.AddContourPoint(new ContourPoint(startLeftTopPoint, null));
                guideline.Curve.AddContourPoint(new ContourPoint(endLeftTopPoint, null));

                rebarSet.Guidelines.Add(guideline);
                bool succes = rebarSet.Insert();
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

                PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
                LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 1, 1 });
            }
        }      
        #endregion
     
      
    }
}
