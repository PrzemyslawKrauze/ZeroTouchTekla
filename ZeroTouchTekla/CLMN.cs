using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using Tekla.Structures;
using Tekla.Structures.Model;
using Tekla.Structures.Geometry3d;


namespace ZeroTouchTekla
{
    class CLMN : Element
    {
        public CLMN(Beam beam) : base(beam)
        {
            GetProfilePointsAndParameters(beam);
        }
        new public void Create()
        {
            OuterStirrups();
            MainRebar(1);
            MainRebar(2);
            MainRebar(3);
            MainRebar(4);
        }
        new public static void GetProfilePointsAndParameters(Beam beam)
        {
            string[] profileValues = GetProfileValues(beam);
            //FTG Width*FirstHeight*SecondHeight*AsymWidth
            double width = Convert.ToDouble(profileValues[0]);
            double height = Convert.ToDouble(profileValues[1]);
            double chamfer = Convert.ToDouble(profileValues[2]);
            double length = Distance.PointToPoint(beam.StartPoint, beam.EndPoint);

            ProfileParameters.Add(CLMNParameter.Width, width); ;
            ProfileParameters.Add(CLMNParameter.Height, height);
            ProfileParameters.Add(CLMNParameter.Chamfer, chamfer);
            ProfileParameters.Add(CLMNParameter.Length, length);

            Point p1 = new Point(0, -height / 2.0, -width / 2.0);
            Point p2 = new Point(0, p1.Y + height, p1.Z);
            Point p3 = new Point(0, p2.Y, p2.Z + width);
            Point p4 = new Point(0, p3.Y - height, p3.Z);

            List<Point> firstProfile = new List<Point> { p1, p2, p3, p4 };

            List<Point> secondProfile = new List<Point>();
            foreach (Point p in firstProfile)
            {
                Point secondPoint = new Point(p.X, p.Y, p.Z);
                secondPoint.Translate(length, 0, 0);
                secondProfile.Add(secondPoint);
            }
            List<List<Point>> beamPoints = new List<List<Point>> { firstProfile, secondProfile };
            ProfilePoints = beamPoints;
            ElementFace = new ElementFace(ProfilePoints);
        }
        void OuterStirrups()
        {
            double startOffset = Convert.ToDouble(Program.ExcelDictionary["M_StartOffset"]);
            string rebarSize = Program.ExcelDictionary["S_Diameter"];
            int variableSpacing = Convert.ToInt32(Program.ExcelDictionary["S_VariableSpacing"]);
            int firstSpacing = Convert.ToInt32(Program.ExcelDictionary["S_FirstSpacing"]);
            double firstSpacingLength = Convert.ToDouble(Program.ExcelDictionary["S_FirstSpacingLength"]);
            int secondSpacing = Convert.ToInt32(Program.ExcelDictionary["S_SecondSpacing"]);

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "FullStirrup";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            var leftFace = ElementFace.GetRebarLegFace(1);
            rebarSet.LegFaces.Add(leftFace);

            var bottomFace = ElementFace.GetRebarLegFace(2);
            rebarSet.LegFaces.Add(bottomFace);

            var rightFace = ElementFace.GetRebarLegFace(3);
            rebarSet.LegFaces.Add(rightFace);

            var topEndFace = ElementFace.GetRebarLegFace(4);
            rebarSet.LegFaces.Add(topEndFace);

            var guideline = new RebarGuideline();
            if (variableSpacing == 0)
            {
                guideline.Spacing.Zones.Add(new RebarSpacingZone
                {
                    Spacing = firstSpacing,
                    SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                    Length = 100,
                    LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
                });
            }
            else
            {
                guideline.Spacing.Zones.Add(new RebarSpacingZone
                {
                    Spacing = firstSpacing,
                    SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                    Length = firstSpacingLength + startOffset,
                    LengthType = RebarSpacingZone.LengthEnum.ABSOLUTE,
                });
                guideline.Spacing.Zones.Add(new RebarSpacingZone
                {
                    Spacing = secondSpacing,
                    SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                    Length = 100,
                    LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
                });
            }

            guideline.Spacing.StartOffset = startOffset + firstSpacing;
            guideline.Spacing.EndOffset = 100;
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            //Create RebarEndDetailModifier
            var leftHookModifier = new RebarEndDetailModifier
            {
                Father = rebarSet,
                EndType = RebarEndDetailModifier.EndTypeEnum.HOOK
            };
            leftHookModifier.RebarHook.Shape = RebarHookData.RebarHookShapeEnum.HOOK_90_DEGREES;
            leftHookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            leftHookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
            leftHookModifier.Insert();
            new Model().CommitChanges();

            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 1, 1, 1 });
        }
        void MainRebar(int faceNumber)
        {
            double startOffset = Convert.ToDouble(Program.ExcelDictionary["M_StartOffset"]);
            int rebarSize = Convert.ToInt32(Program.ExcelDictionary["MR_Diameter"]);
            int spacing = Convert.ToInt32(Program.ExcelDictionary["MR_Spacing"]);
            int addSplitter = Convert.ToInt32(Program.ExcelDictionary["MR_AddSplitter"]);

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "FullStirrup";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(rebarSize);
            rebarSet.RebarProperties.Size = rebarSize.ToString(); ;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(rebarSize);
            rebarSet.LayerOrderNumber = 1;

            RebarLegFace mainFace;
            Point startPoint, endPoint;
            Point oStartPoint, oEndPoint;
            switch (faceNumber)
            {
                case 1:
                    mainFace = ElementFace.GetRebarLegFace(1);
                    startPoint = ProfilePoints[0][0];
                    endPoint = ProfilePoints[0][1];
                    oStartPoint = new Point(startPoint.X, startPoint.Y, startPoint.Z - 40 * rebarSize);
                    oEndPoint = new Point(endPoint.X, endPoint.Y, endPoint.Z - 40 * rebarSize);
                    break;
                case 2:
                    mainFace = ElementFace.GetRebarLegFace(2);
                    startPoint = ProfilePoints[0][1];
                    endPoint = ProfilePoints[0][2];
                    oStartPoint = new Point(startPoint.X, startPoint.Y+40*rebarSize, startPoint.Z);
                    oEndPoint = new Point(endPoint.X, endPoint.Y + 40 * rebarSize, endPoint.Z);
                    break;
                case 3:
                    mainFace = ElementFace.GetRebarLegFace(3);
                    startPoint = ProfilePoints[0][2];
                    endPoint = ProfilePoints[0][3];
                    oStartPoint = new Point(startPoint.X, startPoint.Y, startPoint.Z + 40 * rebarSize);
                    oEndPoint = new Point(endPoint.X, endPoint.Y, endPoint.Z + 40 * rebarSize);
                    break;
                default:
                    mainFace = ElementFace.GetRebarLegFace(4);
                    startPoint = ProfilePoints[0][3];
                    endPoint = ProfilePoints[0][0];
                    oStartPoint = new Point(startPoint.X, startPoint.Y - 40 * rebarSize, startPoint.Z);
                    oEndPoint = new Point(endPoint.X, endPoint.Y - 40 * rebarSize, endPoint.Z);
                    break;
            }
            rebarSet.LegFaces.Add(mainFace);

            var bottomFace = new RebarLegFace();
            bottomFace.Contour.AddContourPoint(new ContourPoint(startPoint, null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(endPoint, null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(oEndPoint, null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(oStartPoint, null));
            rebarSet.LegFaces.Add(bottomFace);

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = spacing,
                SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = 100;
            guideline.Spacing.EndOffset = 100;
            guideline.Curve.AddContourPoint(new ContourPoint(startPoint, null));
            guideline.Curve.AddContourPoint(new ContourPoint(endPoint, null));
            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            var bottomLengthModifier = new RebarEndDetailModifier();
            bottomLengthModifier.Father = rebarSet;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentLength = 10 * Convert.ToInt32(rebarSize);
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(oStartPoint, null));
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(oEndPoint, null));
            bottomLengthModifier.Insert();

            if (addSplitter == 1)
            {
                var bottomSpliter = new RebarSplitter();
                bottomSpliter.Father = rebarSet;
                bottomSpliter.Lapping.LappingType = RebarLapping.LappingTypeEnum.STANDARD_LAPPING;
                bottomSpliter.Lapping.LapSide = RebarLapping.LapSideEnum.LAP_MIDDLE;
                bottomSpliter.Lapping.LapPlacement = RebarLapping.LapPlacementEnum.ON_LEG_FACE;
                bottomSpliter.BarsAffected = BaseRebarModifier.BarsAffectedEnum.EVERY_SECOND_BAR;
                bottomSpliter.FirstAffectedBar = 1;

                Point bottomStartPoint = new Point(startPoint.X + startOffset + 20 * rebarSize, startPoint.Y, startPoint.Z);
                Point bottomEndPoint = new Point(endPoint.X + startOffset + 20 * rebarSize, endPoint.Y, endPoint.Z);

                bottomSpliter.Curve.AddContourPoint(new ContourPoint(bottomStartPoint, null));
                bottomSpliter.Curve.AddContourPoint(new ContourPoint(bottomEndPoint, null));
                bottomSpliter.Insert();

                var topSpliter = new RebarSplitter();
                topSpliter.Father = rebarSet;
                topSpliter.Lapping.LappingType = RebarLapping.LappingTypeEnum.STANDARD_LAPPING;
                topSpliter.Lapping.LapSide = RebarLapping.LapSideEnum.LAP_MIDDLE;
                topSpliter.Lapping.LapPlacement = RebarLapping.LapPlacementEnum.ON_LEG_FACE;
                topSpliter.BarsAffected = BaseRebarModifier.BarsAffectedEnum.EVERY_SECOND_BAR;
                topSpliter.FirstAffectedBar = 2;

                Point topStartPoint = new Point(bottomStartPoint.X + 1.3 * 40 * rebarSize, bottomStartPoint.Y, bottomStartPoint.Z);
                Point topEndPoint = new Point(bottomEndPoint.X + 1.3 * 40 * rebarSize, bottomEndPoint.Y, bottomEndPoint.Z);
                topSpliter.Curve.AddContourPoint(new ContourPoint(topStartPoint, null));
                topSpliter.Curve.AddContourPoint(new ContourPoint(topEndPoint, null));
                topSpliter.Insert();            
            }
            new Model().CommitChanges();

            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2,2});
        }
        void TopClosingRebar(int faceNumber)
        {
            double startOffset = Convert.ToDouble(Program.ExcelDictionary["M_StartOffset"]);
            int rebarSize = Convert.ToInt32(Program.ExcelDictionary["MR_Diameter"]);
            int spacing = Convert.ToInt32(Program.ExcelDictionary["MR_Spacing"]);
            int addSplitter = Convert.ToInt32(Program.ExcelDictionary["MR_AddSplitter"]);

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "FullStirrup";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(rebarSize);
            rebarSet.RebarProperties.Size = rebarSize.ToString(); ;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(rebarSize);
            rebarSet.LayerOrderNumber = 1;

            RebarLegFace mainFace = ElementFace.GetRebarLegFace(5);
            rebarSet.LegFaces.Add(mainFace);

            RebarLegFace leftFace, rightFace;
            Point startPoint, endPoint;
            Point oStartPoint, oEndPoint;
            switch (faceNumber)
            {
                case 1:
                    leftFace = ElementFace.GetRebarLegFace(1);
                    rightFace = ElementFace.GetRebarLegFace(3);
                    startPoint = ProfilePoints[1][0];
                    endPoint = ProfilePoints[1][1];
                    oStartPoint = ProfilePoints[0][0];
                    oEndPoint = ProfilePoints[0][1];
                    break;                
                default:
                    mainFace = ElementFace.GetRebarLegFace(4);
                    startPoint = ProfilePoints[0][3];
                    endPoint = ProfilePoints[0][0];
                    oStartPoint = new Point(startPoint.X, startPoint.Y - 40 * rebarSize, startPoint.Z);
                    oEndPoint = new Point(endPoint.X, endPoint.Y - 40 * rebarSize, endPoint.Z);
                    break;
            }
            

            var bottomFace = new RebarLegFace();
            bottomFace.Contour.AddContourPoint(new ContourPoint(startPoint, null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(endPoint, null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(oEndPoint, null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(oStartPoint, null));
            rebarSet.LegFaces.Add(bottomFace);

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = spacing,
                SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = 100;
            guideline.Spacing.EndOffset = 100;
            guideline.Curve.AddContourPoint(new ContourPoint(startPoint, null));
            guideline.Curve.AddContourPoint(new ContourPoint(endPoint, null));
            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            var bottomLengthModifier = new RebarEndDetailModifier();
            bottomLengthModifier.Father = rebarSet;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentLength = 10 * Convert.ToInt32(rebarSize);
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(oStartPoint, null));
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(oEndPoint, null));
            bottomLengthModifier.Insert();
            new Model().CommitChanges();

            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2, 2 });
        }
        public enum RebarType
        {
            Stirrup,
            MainRebar
        }
        class CLMNParameter : BaseParameter
        {
            public const string Width = "Width";
            public const string Height = "Height";
            public const string Chamfer = "Chamfer";
            public const string Length = "Length";
        }
    }
}
