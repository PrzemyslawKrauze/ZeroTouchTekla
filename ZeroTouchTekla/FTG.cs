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
    public class FTG : Element
    {
        public FTG(Beam beam) : base(beam)
        {
            GetProfilePointsAndParameters(beam);
        }
        new public static void GetProfilePointsAndParameters(Beam beam)
        {
            string[] profileValues = GetProfileValues(beam);
            //FTG Width*FirstHeight*SecondHeight*AsymWidth
            double width = Convert.ToDouble(profileValues[0]);
            double firstHeight = Convert.ToDouble(profileValues[1]);
            double secondHeight = Convert.ToDouble(profileValues[2]);
            double length = Distance.PointToPoint(beam.StartPoint, beam.EndPoint);

            ProfileParameters.Add(FTGParameter.Width, width); ;
            ProfileParameters.Add(FTGParameter.FirstHeight, firstHeight);
            ProfileParameters.Add(FTGParameter.SecondHeight, secondHeight);
            ProfileParameters.Add(FTGParameter.Length, length);

            Point p1 = new Point(0, -(firstHeight + secondHeight) / 2.0, -width / 2.0);
            Point p2 = new Point(0, p1.Y + firstHeight, p1.Z);
            Point p3 = new Point(0, p2.Y + secondHeight, 0);
            Point p4 = new Point(0, p2.Y, width / 2.0);
            Point p5 = new Point(0, p1.Y, p4.Z);

            if (profileValues.Length > 3)
            {
                double asymWidth = Convert.ToDouble(profileValues[3]);
                ProfileParameters.Add(FTGParameter.AsymWiidth, asymWidth);
                p3 = new Point(0, p3.Y, p2.Z + asymWidth);
            }

            List<Point> firstProfile = new List<Point> { p1, p2, p3, p4, p5 };

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
        void FullStirrups()
        {
            string rebarSize = Program.ExcelDictionary["S_Diameter"];
            int rowSpacing = Convert.ToInt32(Program.ExcelDictionary["S_RowSpacing"]);
            int barSpacing = Convert.ToInt32(Program.ExcelDictionary["S_BarSpacing"]);
            int stirrupSpacing = Convert.ToInt32(Program.ExcelDictionary["S_StirrupSpacing"]);

            double length = ProfileParameters[FTGParameter.Length];
            int numberOfStirrupSets = (int)Math.Floor((length - 2 * SideCover) / (stirrupSpacing + barSpacing));
            double leftover = length - barSpacing * numberOfStirrupSets - stirrupSpacing * (numberOfStirrupSets - 1);

            for (int i = 0; i < numberOfStirrupSets; i++)
            {
                var rebarSet = new RebarSet();
                rebarSet.RebarProperties.Name = "FullStirrup";
                rebarSet.RebarProperties.Grade = "B500SP";
                rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
                rebarSet.RebarProperties.Size = rebarSize;
                rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
                rebarSet.LayerOrderNumber = 1;

                //  Point startMidPoint = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y, ProfilePoints[0][2].Z);
                //Point endMidPoint = new Point(ProfilePoints[1][0].X, ProfilePoints[1][0].Y, ProfilePoints[1][2].Z);       

                List<List<Point>> correctedPoints = new List<List<Point>>();
                double additionalOffset = (leftover / 2.0 + i * (barSpacing + stirrupSpacing));

                List<Point> cP = new List<Point>();
                foreach (Point p in ProfilePoints[0])
                {
                    Point correctedPoint = new Point(p.X + additionalOffset, p.Y, p.Z);
                    cP.Add(correctedPoint);
                }
                correctedPoints.Add(cP);

                List<Point> cP2 = new List<Point>();
                foreach (Point p in cP)
                {
                    Point correctedPoint = new Point(p.X + stirrupSpacing, p.Y, p.Z);
                    cP2.Add(correctedPoint);
                }
                correctedPoints.Add(cP2);

                var leftFace = new RebarLegFace();                
                leftFace.Contour.AddContourPoint(new ContourPoint(correctedPoints[1][0], null));
                leftFace.Contour.AddContourPoint(new ContourPoint(correctedPoints[1][4], null));
                leftFace.Contour.AddContourPoint(new ContourPoint(correctedPoints[1][3], null));
                leftFace.Contour.AddContourPoint(new ContourPoint(correctedPoints[1][2], null));
                leftFace.Contour.AddContourPoint(new ContourPoint(correctedPoints[1][1], null));
                
                rebarSet.LegFaces.Add(leftFace);

                var bottomFace = new RebarLegFace();
                bottomFace.Contour.AddContourPoint(new ContourPoint(correctedPoints[0][0], null));
                bottomFace.Contour.AddContourPoint(new ContourPoint(correctedPoints[1][0], null));
                bottomFace.Contour.AddContourPoint(new ContourPoint(correctedPoints[1][4], null));
                bottomFace.Contour.AddContourPoint(new ContourPoint(correctedPoints[0][4], null));
                rebarSet.LegFaces.Add(bottomFace);

                var rightFace = new RebarLegFace();
                // rightFace.AdditonalOffset = length - (leftover / 2.0 + barSpacing) - i * (barSpacing + stirrupSpacing);
                rightFace.Contour.AddContourPoint(new ContourPoint(correctedPoints[0][0], null));
                rightFace.Contour.AddContourPoint(new ContourPoint(correctedPoints[0][4], null));
                rightFace.Contour.AddContourPoint(new ContourPoint(correctedPoints[0][3], null));
                rightFace.Contour.AddContourPoint(new ContourPoint(correctedPoints[0][2], null));
                rightFace.Contour.AddContourPoint(new ContourPoint(correctedPoints[0][1], null));
                rebarSet.LegFaces.Add(rightFace);

                var topEndFace = new RebarLegFace();
                topEndFace.Contour.AddContourPoint(new ContourPoint(correctedPoints[0][2], null));
                topEndFace.Contour.AddContourPoint(new ContourPoint(correctedPoints[1][2], null));
                topEndFace.Contour.AddContourPoint(new ContourPoint(correctedPoints[1][3], null));
                topEndFace.Contour.AddContourPoint(new ContourPoint(correctedPoints[0][3], null));
                rebarSet.LegFaces.Add(topEndFace);

                var topStartFace = new RebarLegFace();
                topStartFace.Contour.AddContourPoint(new ContourPoint(correctedPoints[0][1], null));
                topStartFace.Contour.AddContourPoint(new ContourPoint(correctedPoints[1][1], null));
                topStartFace.Contour.AddContourPoint(new ContourPoint(correctedPoints[1][2], null));
                topStartFace.Contour.AddContourPoint(new ContourPoint(correctedPoints[0][2], null));
                rebarSet.LegFaces.Add(topStartFace);


                var guideline = new RebarGuideline();
                guideline.Spacing.Zones.Add(new RebarSpacingZone
                {
                    Spacing = rowSpacing,
                    SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                    Length = 100,
                    LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
                });
                guideline.Spacing.StartOffset = 100;
                guideline.Spacing.EndOffset = 100;

                guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
                guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][4], null));

                rebarSet.Guidelines.Add(guideline);
                bool succes = rebarSet.Insert();

                //Create RebarEndDetailModifier
                var leftHookModifier = new RebarEndDetailModifier
                {
                    Father = rebarSet,
                    EndType = RebarEndDetailModifier.EndTypeEnum.HOOK
                };
                leftHookModifier.RebarHook.Shape = RebarHookData.RebarHookShapeEnum.HOOK_90_DEGREES;
                leftHookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
                leftHookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
                leftHookModifier.Insert();

                var rightHookModifier = new RebarEndDetailModifier
                {
                    Father = rebarSet,
                    EndType = RebarEndDetailModifier.EndTypeEnum.HOOK
                };
                rightHookModifier.RebarHook.Shape = RebarHookData.RebarHookShapeEnum.HOOK_90_DEGREES;
                rightHookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
                rightHookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][3], null));
                rightHookModifier.Insert();
                new Model().CommitChanges();

                RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 1, 1, 1, 1 });
            }
        }
        void Stirrups()
        {
            string rebarSize = Program.ExcelDictionary["S_Diameter"];
            int rowSpacing = Convert.ToInt32(Program.ExcelDictionary["S_RowSpacing"]);
            int barSpacing = Convert.ToInt32(Program.ExcelDictionary["S_BarSpacing"]);

            double length = ProfileParameters[FTGParameter.Length];
            int numberOfStirrupSets = (int)Math.Floor((length - 2 * SideCover) / (barSpacing));
            double leftover = length - barSpacing * (numberOfStirrupSets - 1);

            for (int i = 0; i < numberOfStirrupSets; i++)
            {
                var rebarSet = new RebarSet();
                rebarSet.RebarProperties.Name = "Stirrup";
                rebarSet.RebarProperties.Grade = "B500SP";
                rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
                rebarSet.RebarProperties.Size = rebarSize;
                rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
                rebarSet.LayerOrderNumber = 1;

                Point startMidPoint = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y, ProfilePoints[0][2].Z);
                Point endMidPoint = new Point(ProfilePoints[1][0].X, ProfilePoints[1][0].Y, ProfilePoints[1][2].Z);

                List<List<Point>> correctedPoints = new List<List<Point>>();
                double additionalOffset = (leftover / 2.0 + i * barSpacing);

                List<Point> cP = new List<Point>();
                foreach (Point p in ProfilePoints[0])
                {
                    Point correctedPoint = new Point(p.X + additionalOffset, p.Y, p.Z);
                    cP.Add(correctedPoint);
                }
                correctedPoints.Add(cP);

                var leftFace = new RebarLegFace();
                // double additionalOffset = leftover / 2.0 + i * (barSpacing + stirrupSpacing) + sideCover;
                leftFace.Contour.AddContourPoint(new ContourPoint(correctedPoints[0][0], null));
                leftFace.Contour.AddContourPoint(new ContourPoint(correctedPoints[0][4], null));
                leftFace.Contour.AddContourPoint(new ContourPoint(correctedPoints[0][3], null));
                leftFace.Contour.AddContourPoint(new ContourPoint(correctedPoints[0][2], null));
                leftFace.Contour.AddContourPoint(new ContourPoint(correctedPoints[0][1], null));
                rebarSet.LegFaces.Add(leftFace);


                var guideline = new RebarGuideline();
                guideline.Spacing.Zones.Add(new RebarSpacingZone
                {
                    Spacing = rowSpacing,
                    SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                    Length = 100,
                    LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
                });
                guideline.Spacing.StartOffset = 100;
                guideline.Spacing.EndOffset = 100;

                guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
                guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][4], null));

                rebarSet.Guidelines.Add(guideline);
                bool succes = rebarSet.Insert();
                new Model().CommitChanges();

                //Create RebarEndDetailModifier
                var leftTopHookModifier = new RebarEndDetailModifier
                {
                    Father = rebarSet,
                    EndType = RebarEndDetailModifier.EndTypeEnum.HOOK
                };
                leftTopHookModifier.RebarHook.Shape = RebarHookData.RebarHookShapeEnum.HOOK_90_DEGREES;
                leftTopHookModifier.RebarHook.Rotation = 90;
                leftTopHookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
                leftTopHookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
                leftTopHookModifier.Insert();

                var rightTopHookModifier = new RebarEndDetailModifier
                {
                    Father = rebarSet,
                    EndType = RebarEndDetailModifier.EndTypeEnum.HOOK
                };
                rightTopHookModifier.RebarHook.Shape = RebarHookData.RebarHookShapeEnum.HOOK_90_DEGREES;
                rightTopHookModifier.RebarHook.Rotation = -90;
                rightTopHookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
                rightTopHookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][3], null));
                rightTopHookModifier.Insert();

                var leftBottomHookModifier = new RebarEndDetailModifier
                {
                    Father = rebarSet,
                    EndType = RebarEndDetailModifier.EndTypeEnum.HOOK
                };
                leftBottomHookModifier.RebarHook.Shape = RebarHookData.RebarHookShapeEnum.HOOK_90_DEGREES;
                leftBottomHookModifier.RebarHook.Rotation = -90;
                leftBottomHookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
                leftBottomHookModifier.Curve.AddContourPoint(new ContourPoint(endMidPoint, null));
                leftBottomHookModifier.Insert();

                var rightBottomHookModifier = new RebarEndDetailModifier
                {
                    Father = rebarSet,
                    EndType = RebarEndDetailModifier.EndTypeEnum.HOOK
                };
                rightBottomHookModifier.RebarHook.Shape = RebarHookData.RebarHookShapeEnum.HOOK_90_DEGREES;
                rightBottomHookModifier.RebarHook.Rotation = 90;
                rightBottomHookModifier.Curve.AddContourPoint(new ContourPoint(endMidPoint, null));
                rightBottomHookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
                rightBottomHookModifier.Insert();

                new Model().CommitChanges();

                RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1 });
            }
        }
        void TopPerpendicularRebar()
        {
            string rebarSize = Program.ExcelDictionary["TPR_Diameter"];
            string spacing = Program.ExcelDictionary["TPR_Spacing"];
            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "TPR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            var leftFace = ElementFace.GetRebarLegFace(1);         
            rebarSet.LegFaces.Add(leftFace);

            var topLeftFace = ElementFace.GetRebarLegFace(2);
            rebarSet.LegFaces.Add(topLeftFace);

            var topRightFace = ElementFace.GetRebarLegFace(3);
            rebarSet.LegFaces.Add(topRightFace);

            var rightFace = ElementFace.GetRebarLegFace(4);
            rebarSet.LegFaces.Add(rightFace);

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
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();

            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 2, 2, 1 });
        }
        void BottomPerpendicularRebar()
        {
            string rebarSize = Program.ExcelDictionary["BPR_Diameter"];
            string spacing = Program.ExcelDictionary["BPR_Spacing"];
            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "BPR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            var legFace1 = ElementFace.GetRebarLegFace(5);
            rebarSet.LegFaces.Add(legFace1);

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
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));

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

            var rightHookModifier = new RebarEndDetailModifier
            {
                Father = rebarSet,
                EndType = RebarEndDetailModifier.EndTypeEnum.HOOK
            };
            rightHookModifier.RebarHook.Shape = RebarHookData.RebarHookShapeEnum.HOOK_90_DEGREES;
            rightHookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][4], null));
            rightHookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
            rightHookModifier.Insert();
            new Model().CommitChanges();

            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2 });
        }
        void BottomLongitudinalRebar()
        {
            string rebarSize = Program.ExcelDictionary["BLR_Diameter"];
            string spacing = Program.ExcelDictionary["BLR_Spacing"];
            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "BLR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 2;

            var legFace1 = ElementFace.GetRebarLegFace(5);
            rebarSet.LegFaces.Add(legFace1);

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
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][4], null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();

            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 3 });
        }
        void TopLongitudinalLeftRebar()
        {
            string rebarSize = Program.ExcelDictionary["TLR_Diameter"];
            string spacing = Program.ExcelDictionary["TLR_Spacing"];
            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "TLR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            var legFace1 = ElementFace.GetRebarLegFace(2);
            rebarSet.LegFaces.Add(legFace1);

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing),
                SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = 150;
            guideline.Spacing.EndOffset = Convert.ToDouble(spacing) / 2.0;

            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();

            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 3 });
        }
        void TopLongitudinalRightRebar()
        {
            string rebarSize = Program.ExcelDictionary["TLR_Diameter"];
            string spacing = Program.ExcelDictionary["TLR_Spacing"];
            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "TLR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            var legFace1 = ElementFace.GetRebarLegFace(3);
            rebarSet.LegFaces.Add(legFace1);

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing),
                SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = Convert.ToDouble(spacing) / 2.0;
            guideline.Spacing.EndOffset = 150;

            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();

            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 3 });
        }
        void ClosingCShapeRebar(int faceNumber)
        {
            //1 - StartLeft, 2- StarRight, 3 - EndLeft, 4 -EndRight
            string rebarSize = Program.ExcelDictionary["CR_Diameter"];
            string spacing = Program.ExcelDictionary["CR_Spacing"];
            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "CCSR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            Point bottomLeft, bottomRight;
            Point endBottomLeft, endTopLeft, endTopMid, endTopRight, endBottomRight;

            RebarLegFace mainFace;
            switch (faceNumber)
            {
                case 1:
                    mainFace = ElementFace.GetRebarLegFace(0);
                    bottomLeft = ProfilePoints[0][0];
                    bottomRight = ProfilePoints[0][4];
                    endBottomLeft = ProfilePoints[1][0];
                    endTopLeft = ProfilePoints[1][1];
                    endTopMid = ProfilePoints[1][2];
                    endBottomRight = ProfilePoints[1][4];
                    endTopRight = ProfilePoints[1][3];
                    break;
                default:
                    mainFace = ElementFace.GetRebarLegFace(6);
                    bottomLeft = ProfilePoints[1][0];
                    bottomRight = ProfilePoints[1][4];
                    endBottomLeft = ProfilePoints[0][0];
                    endTopLeft = ProfilePoints[0][1];
                    endTopMid = ProfilePoints[0][2];
                    endBottomRight = ProfilePoints[0][4];
                    endTopRight = ProfilePoints[0][3];
                    break;
            }

            rebarSet.LegFaces.Add(mainFace);

            var bottomFace = ElementFace.GetRebarLegFace(5);
            rebarSet.LegFaces.Add(bottomFace);

            var topFaceLeft = ElementFace.GetRebarLegFace(2);
            rebarSet.LegFaces.Add(topFaceLeft);

            var topFaceRight = ElementFace.GetRebarLegFace(3);
            rebarSet.LegFaces.Add(topFaceRight);

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing),
                SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = 150;
            guideline.Spacing.EndOffset = 150;

            guideline.Curve.AddContourPoint(new ContourPoint(bottomLeft, null));
            guideline.Curve.AddContourPoint(new ContourPoint(bottomRight, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            //Create RebarEndDetailModifier
            var bottomLengthModifier = new RebarEndDetailModifier();
            bottomLengthModifier.Father = rebarSet;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentLength = 10 * Convert.ToInt32(rebarSize);
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(endBottomLeft, null));
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(endBottomRight, null));
            bottomLengthModifier.Insert();

            var topLengthModifierLeft = new RebarEndDetailModifier();
            topLengthModifierLeft.Father = rebarSet;
            topLengthModifierLeft.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            topLengthModifierLeft.RebarLengthAdjustment.AdjustmentLength = 10 * Convert.ToInt32(rebarSize);
            topLengthModifierLeft.Curve.AddContourPoint(new ContourPoint(endTopLeft, null));
            topLengthModifierLeft.Curve.AddContourPoint(new ContourPoint(endTopMid, null));
            topLengthModifierLeft.Insert();

            var topLengthModifierRight = new RebarEndDetailModifier();
            topLengthModifierRight.Father = rebarSet;
            topLengthModifierRight.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            topLengthModifierRight.RebarLengthAdjustment.AdjustmentLength = 10 * Convert.ToInt32(rebarSize);
            topLengthModifierRight.Curve.AddContourPoint(new ContourPoint(endTopMid, null));
            topLengthModifierRight.Curve.AddContourPoint(new ContourPoint(endTopRight, null));
            topLengthModifierRight.Insert();
            new Model().CommitChanges();

            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 3, 3, 3 });
        }
        void ClosingLongitudinalRebar(int faceNumber)
        {
            //1 - Start, 2- Left, 3 - Right, 4 - End
            string rebarSize = Program.ExcelDictionary["CR_Diameter"];
            string spacing = Program.ExcelDictionary["CR_Spacing"];
            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "CLR";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            Point bottomLeft, topLeft, bottomRight, topRight;
            Point endBottomLeft, endTopLeft, endBottomRight, endTopRight;
            switch (faceNumber)
            {
                case 1:
                    bottomLeft = ProfilePoints[0][0];
                    topLeft = ProfilePoints[0][1];
                    bottomRight = ProfilePoints[0][4];
                    topRight = ProfilePoints[0][3];
                    endBottomLeft = ProfilePoints[1][0];
                    endTopLeft = ProfilePoints[1][1];
                    endBottomRight = ProfilePoints[1][4];
                    endTopRight = ProfilePoints[1][3];
                    break;
                case 2:
                    bottomLeft = ProfilePoints[1][0];
                    topLeft = ProfilePoints[1][1];
                    bottomRight = ProfilePoints[0][0];
                    topRight = ProfilePoints[0][1];
                    endBottomLeft = ProfilePoints[1][4];
                    endTopLeft = ProfilePoints[1][3];
                    endBottomRight = ProfilePoints[0][4];
                    endTopRight = ProfilePoints[0][3];
                    break;
                case 3:
                    bottomLeft = ProfilePoints[0][4];
                    topLeft = ProfilePoints[0][3];
                    bottomRight = ProfilePoints[1][4];
                    topRight = ProfilePoints[1][3];
                    endBottomLeft = ProfilePoints[0][0];
                    endTopLeft = ProfilePoints[0][1];
                    endBottomRight = ProfilePoints[1][0];
                    endTopRight = ProfilePoints[1][1];
                    break;
                default:
                    bottomLeft = ProfilePoints[1][0];
                    topLeft = ProfilePoints[1][1];
                    bottomRight = ProfilePoints[1][4];
                    topRight = ProfilePoints[1][3];
                    endBottomLeft = ProfilePoints[0][0];
                    endTopLeft = ProfilePoints[0][1];
                    endBottomRight = ProfilePoints[0][4];
                    endTopRight = ProfilePoints[0][3];
                    break;
            }

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(bottomLeft, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(bottomRight, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(topRight, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(topLeft, null));
            rebarSet.LegFaces.Add(mainFace);

            var leftFace = new RebarLegFace();
            leftFace.Contour.AddContourPoint(new ContourPoint(bottomLeft, null));
            leftFace.Contour.AddContourPoint(new ContourPoint(endBottomLeft, null));
            leftFace.Contour.AddContourPoint(new ContourPoint(endTopLeft, null));
            leftFace.Contour.AddContourPoint(new ContourPoint(topLeft, null));
            rebarSet.LegFaces.Add(leftFace);

            var rightFace = new RebarLegFace();
            rightFace.Contour.AddContourPoint(new ContourPoint(bottomRight, null));
            rightFace.Contour.AddContourPoint(new ContourPoint(endBottomRight, null));
            rightFace.Contour.AddContourPoint(new ContourPoint(endTopRight, null));
            rightFace.Contour.AddContourPoint(new ContourPoint(topRight, null));
            rebarSet.LegFaces.Add(rightFace);

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = Convert.ToInt32(spacing),
                SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });
            guideline.Spacing.StartOffset = 100;
            guideline.Spacing.EndOffset = Convert.ToDouble(spacing) / 2.0;

            guideline.Curve.AddContourPoint(new ContourPoint(bottomLeft, null));
            guideline.Curve.AddContourPoint(new ContourPoint(topLeft, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();

            //Create RebarEndDetailModifier
            var bottomLengthModifier = new RebarEndDetailModifier();
            bottomLengthModifier.Father = rebarSet;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentLength = 10 * Convert.ToInt32(rebarSize);
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(endBottomLeft, null));
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(endTopLeft, null));
            bottomLengthModifier.Insert();

            var topLengthModifier = new RebarEndDetailModifier();
            topLengthModifier.Father = rebarSet;
            topLengthModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            topLengthModifier.RebarLengthAdjustment.AdjustmentLength = 10 * Convert.ToInt32(rebarSize);
            topLengthModifier.Curve.AddContourPoint(new ContourPoint(endBottomRight, null));
            topLengthModifier.Curve.AddContourPoint(new ContourPoint(endTopRight, null));
            topLengthModifier.Insert();
            new Model().CommitChanges();

            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2, 2, 2 });
        }
        new public void Create()
        {
            if (Convert.ToInt32(Program.ExcelDictionary["S_FullStirrups"]) == 1)
            {
                FullStirrups();
            }
            else
            {
                Stirrups();
            }
            TopPerpendicularRebar();
            BottomPerpendicularRebar();
            BottomLongitudinalRebar();
            TopLongitudinalLeftRebar();
            TopLongitudinalRightRebar();
            ClosingCShapeRebar(1);
            ClosingCShapeRebar(2);
            ClosingLongitudinalRebar(1);
            ClosingLongitudinalRebar(2);
            ClosingLongitudinalRebar(3);
            ClosingLongitudinalRebar(4);
        }
        new public void CreateSingle(string rebarName)
        {
            RebarType rType;
            Enum.TryParse(rebarName, out rType);
            switch (rType)
            {
                case RebarType.Stirrup:
                    Stirrups();
                    break;
                case RebarType.FullStirrup:
                    FullStirrups();
                    break;
                case RebarType.TPR:
                    TopPerpendicularRebar();
                    break;
                case RebarType.BPR:
                    BottomPerpendicularRebar();
                    break;
                case RebarType.BLR:
                    BottomLongitudinalRebar();
                    break;
                case RebarType.TLR:
                    TopLongitudinalLeftRebar();
                    TopLongitudinalRightRebar();
                    break;
                case RebarType.CCSR:
                    ClosingCShapeRebar(1);
                    ClosingCShapeRebar(2);
                    break;
                case RebarType.CLR:
                    ClosingLongitudinalRebar(1);
                    ClosingLongitudinalRebar(2);
                    ClosingLongitudinalRebar(3);
                    ClosingLongitudinalRebar(4);
                    break;
            }
        }
        public enum RebarType
        {
            FullStirrup,
            Stirrup,
            TPR,
            BPR,
            BLR,
            TLR,
            CCSR,
            CLR
        }
        class FTGParameter : BaseParameter
        {
            public const string Width = "Width";
            public const string FirstHeight = "FirstHeight";
            public const string SecondHeight = "SecondHeight";
            public const string AsymWiidth = "AsymWidth";
            public const string Length = "Length";
        }
    }
}
