﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections;

using Tekla.Structures;
using Tekla.Structures.Model;
using Tekla.Structures.Geometry3d;
using System.Reflection;
using Tekla.Structures.Solid;


namespace ZeroTouchTekla.Profiles
{
    public class FTG : Element
    {
        #region Fields
        enum RebarType
        {
            FS,
            S,
            TPR,
            BPR,
            BLR,
            TLLR,
            TLRR,
            CCSR,
            CLR
        }
        double width;
        double firstHeight;
        double secondHeight;
        double asymWidth;
        double length;
        double horizontalOffset;
        double verticalOffset;

        #endregion
        #region Constructor
        public FTG(params Part[] parts)
        {
            if (parts.Count() > 1)
            {
                throw new ArgumentException("Multiple parts not supported");
            }
            else
            {
                base.BaseParts = parts;
                SetLocalPlane();
                //GetProfilePointsAndParameters(parts[0]);
                SetProfileParameters(parts[0]);
                SetProfilePoints(parts[0]);
                SetRebarLegFaces();
            }
        }
        #endregion
        #region PublicMathods
        void SetProfileParameters(Part part)
        {
            Beam beam = part as Beam;
            //Get beam local plane

            string[] profileValues = GetProfileValues(beam);
            string profileName = beam.Profile.ProfileString;
            //FTG Width*FirstHeight*SecondHeight
            //FTGASYM Width*FirstHeight*SecondHeight*AsymWidth
            //FTGSK Width*FirstHeight*SecondHeight*HorizontalOfset*VerticalOffset
            double width = Convert.ToDouble(profileValues[0]);
            double firstHeight = Convert.ToDouble(profileValues[1]);
            double secondHeight = Convert.ToDouble(profileValues[2]);
            double length = Distance.PointToPoint(beam.StartPoint, beam.EndPoint);
            length -= beam.StartPointOffset.Dx;
            length += beam.EndPointOffset.Dx;

            this.width = width;
            this.firstHeight = firstHeight;
            this.secondHeight = secondHeight;
            this.length = length;

            if (profileName.Contains("ASYM"))
            {
                double asymWidth = Convert.ToDouble(profileValues[3]);
                this.asymWidth = asymWidth;
            }
            else
            {
                asymWidth = 0;
            }

            if (profileName.Contains("V"))
            {
                if (asymWidth == 0)
                {
                    horizontalOffset = Convert.ToDouble(profileValues[3]);
                    verticalOffset = Convert.ToDouble(profileValues[4]);
                }
                else
                {
                    horizontalOffset = Convert.ToDouble(profileValues[4]);
                    verticalOffset = Convert.ToDouble(profileValues[5]);
                }
            }
        }
        void SetProfilePoints(Part part)
        {
            base.ProfilePoints = TeklaUtils.GetSortedPointsFromEndFaces(part);
        }
        void SetRebarLegFaces()
        {
            RebarLegFace startFace = new RebarLegFace();
            startFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            startFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            startFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][4], null));
            startFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            startFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
            RebarLegFaces.Add(startFace);

            RebarLegFace firstFace = new RebarLegFace();
            firstFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            firstFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
            firstFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
            firstFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));          
            RebarLegFaces.Add(firstFace);

            RebarLegFace secondFace = new RebarLegFace();
            secondFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            secondFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
            secondFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
            secondFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][4], null));
            RebarLegFaces.Add(secondFace);

            RebarLegFace thirdFace = new RebarLegFace();
            thirdFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][4], null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][3], null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            RebarLegFaces.Add(thirdFace);

            RebarLegFace fourthFace = new RebarLegFace();
            fourthFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            fourthFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][3], null));
            fourthFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
            fourthFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
            RebarLegFaces.Add(fourthFace);

            RebarLegFace fifthFace = new RebarLegFace();
            fifthFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
            fifthFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
            fifthFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
            fifthFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            RebarLegFaces.Add(fifthFace);

            RebarLegFace endFace = new RebarLegFace();
            endFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
            endFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
            endFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
            endFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][3], null));
            endFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
            RebarLegFaces.Add(endFace);
        }      
        public override void Create()
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
        public override void CreateSingle(string rebarName)
        {
            rebarName = rebarName.Split('_')[1];
            RebarType rType;
            Enum.TryParse(rebarName, out rType);
            switch (rType)
            {
                case RebarType.FS:
                    Stirrups();
                    break;
                case RebarType.S:
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
                case RebarType.TLLR:
                    TopLongitudinalLeftRebar();
                    break;
                case RebarType.TLRR:
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
        #endregion
        #region PrivateMethods
        void FullStirrups()
        {
            int rebarSize = Convert.ToInt32(Program.ExcelDictionary["S_Diameter"]);
            double rowSpacing = Convert.ToDouble(Program.ExcelDictionary["S_RowSpacing"]) + Convert.ToDouble(rebarSize) / 2.0;
            double barSpacing = Convert.ToDouble(Program.ExcelDictionary["S_BarSpacing"]) - Convert.ToDouble(rebarSize) / 2.0;
            int stirrupSpacing = Convert.ToInt32(Program.ExcelDictionary["S_StirrupSpacing"]);

            double length = this.length;
            int numberOfStirrupSets = (int)Math.Floor((length - 2 * SideCover) / (stirrupSpacing + barSpacing));
            double leftover = length - barSpacing * numberOfStirrupSets - stirrupSpacing * (numberOfStirrupSets - 1);
            if (horizontalOffset != 0)
            {
                numberOfStirrupSets++;
            }

            for (int i = 0; i < numberOfStirrupSets; i++)
            {
                string name = CreateRebarName(RebarType.FS);
                var rebarSet = CreateDefaultRebarSet(name, rebarSize);

                List<List<Point>> correctedPoints = new List<List<Point>>();
                double additionalOffset = (leftover / 2.0 + i * (barSpacing + stirrupSpacing));

                Vector skewVector = Utility.GetVectorFromTwoPoints(ProfilePoints[0][0], ProfilePoints[1][0]).GetNormal();
                Point origin = new Point(ProfilePoints[0][0].X + additionalOffset,
                    ProfilePoints[0][0].Y + (1 / skewVector.X) * additionalOffset * skewVector.Y,
                    ProfilePoints[0][0].Z + (1 / skewVector.X) * additionalOffset * skewVector.Z);
                Vector normal = skewVector.GetNormal();
                GeometricPlane geometricPlane = new GeometricPlane(origin, normal);

                Line line0 = new Line(ProfilePoints[0][0], ProfilePoints[1][0]);
                Line line1 = new Line(ProfilePoints[0][2], ProfilePoints[1][2]);
                Line line2 = new Line(ProfilePoints[0][4], ProfilePoints[1][4]);
                Line line3 = new Line(ProfilePoints[0][3], ProfilePoints[1][3]);
                Line line4 = new Line(ProfilePoints[0][1], ProfilePoints[1][1]);
                Point point0 = Intersection.LineToPlane(line0, geometricPlane);
                Point point1 = Intersection.LineToPlane(line1, geometricPlane);
                Point point2 = Intersection.LineToPlane(line2, geometricPlane);
                Point point3 = Intersection.LineToPlane(line3, geometricPlane);
                Point point4 = Intersection.LineToPlane(line4, geometricPlane);
                List<Point> cP = new List<Point> { point0, point1, point2, point3, point4 };
                correctedPoints.Add(cP);

                Point secondOrigin = new Point(origin.X + stirrupSpacing, origin.Y + (1 / skewVector.X) * stirrupSpacing * skewVector.Y, origin.Z + (1 / skewVector.X) * stirrupSpacing * skewVector.Z);
                GeometricPlane secondGP = new GeometricPlane(secondOrigin, normal);

                Point secondPoint0 = Intersection.LineToPlane(line0, secondGP);
                Point secondPoint1 = Intersection.LineToPlane(line1, secondGP);
                Point secondPoint2 = Intersection.LineToPlane(line2, secondGP);
                Point secondPoint3 = Intersection.LineToPlane(line3, secondGP);
                Point secondPoint4 = Intersection.LineToPlane(line4, secondGP);
                List<Point> cP2 = new List<Point> { secondPoint0, secondPoint1, secondPoint2, secondPoint3, secondPoint4 };
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

                Point startGL = correctedPoints[0][0];
                Point endGL = correctedPoints[0][4];


                if (i == 0 && horizontalOffset != 0)
                {
                    GeometricPlane startplane = new GeometricPlane(ProfilePoints[0][0], new Vector(1, 0, 0));
                    Line gl = new Line(startGL, endGL);
                    Point intersection = Intersection.LineToPlane(gl, startplane);
                    endGL = intersection;
                }
                if (i == numberOfStirrupSets - 1 && horizontalOffset != 0)
                {
                    startGL = correctedPoints[1][0];
                    endGL = correctedPoints[1][4];
                    GeometricPlane endPlane = new GeometricPlane(ProfilePoints[1][0], new Vector(1, 0, 0));
                    Line gl = new Line(startGL, endGL);
                    Point intersection = Intersection.LineToPlane(gl, endPlane);
                    startGL = intersection;
                }

                guideline.Curve.AddContourPoint(new ContourPoint(startGL, null));
                guideline.Curve.AddContourPoint(new ContourPoint(endGL, null));

                rebarSet.Guidelines.Add(guideline);
                bool succes = rebarSet.Insert();

                //Create RebarEndDetailModifier
                var leftHookModifier = new RebarEndDetailModifier
                {
                    Father = rebarSet,
                    EndType = RebarEndDetailModifier.EndTypeEnum.HOOK
                };
                leftHookModifier.RebarHook.Shape = RebarHookData.RebarHookShapeEnum.HOOK_90_DEGREES;
                leftHookModifier.Curve.AddContourPoint(new ContourPoint(correctedPoints[1][1], null));
                leftHookModifier.Curve.AddContourPoint(new ContourPoint(correctedPoints[1][2], null));
                leftHookModifier.Insert();

                var rightHookModifier = new RebarEndDetailModifier
                {
                    Father = rebarSet,
                    EndType = RebarEndDetailModifier.EndTypeEnum.HOOK
                };
                rightHookModifier.RebarHook.Shape = RebarHookData.RebarHookShapeEnum.HOOK_90_DEGREES;
                rightHookModifier.Curve.AddContourPoint(new ContourPoint(correctedPoints[1][2], null));
                rightHookModifier.Curve.AddContourPoint(new ContourPoint(correctedPoints[1][3], null));
                rightHookModifier.Insert();
                new Model().CommitChanges();

                PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
                LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 1, 1, 1, 1 });
            }
        }
        void Stirrups()
        {
            int rebarSize = Convert.ToInt32(Program.ExcelDictionary["S_Diameter"]);
            int rowSpacing = Convert.ToInt32(Program.ExcelDictionary["S_RowSpacing"]);
            int barSpacing = Convert.ToInt32(Program.ExcelDictionary["S_BarSpacing"]);

            double length = this.length;
            int numberOfStirrupSets = (int)Math.Floor((length - 2 * SideCover) / (barSpacing));
            double leftover = length - barSpacing * (numberOfStirrupSets - 1);

            for (int i = 0; i < numberOfStirrupSets; i++)
            {
                string name = CreateRebarName(RebarType.S);
                var rebarSet = CreateDefaultRebarSet(name, rebarSize);

                Point startMidPoint = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y, ProfilePoints[0][4].Z);
                Point endMidPoint = new Point(ProfilePoints[1][0].X, ProfilePoints[1][0].Y, ProfilePoints[1][4].Z);

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
                guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));

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
                leftTopHookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
                leftTopHookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
                leftTopHookModifier.Insert();

                var rightTopHookModifier = new RebarEndDetailModifier
                {
                    Father = rebarSet,
                    EndType = RebarEndDetailModifier.EndTypeEnum.HOOK
                };
                rightTopHookModifier.RebarHook.Shape = RebarHookData.RebarHookShapeEnum.HOOK_90_DEGREES;
                rightTopHookModifier.RebarHook.Rotation = -90;
                rightTopHookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
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
                rightBottomHookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
                rightBottomHookModifier.Insert();

                new Model().CommitChanges();

                PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
                LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1 });
            }
        }
        void TopPerpendicularRebar()
        {
            int rebarSize = Convert.ToInt32( Program.ExcelDictionary["TPR_Diameter"]);
            string spacing = Program.ExcelDictionary["TPR_Spacing"];

            string name = CreateRebarName(RebarType.TPR);
            var rebarSet = CreateDefaultRebarSet(name, rebarSize);

            var leftFace = GetRebarLegFace(1);
            rebarSet.LegFaces.Add(leftFace);

            var topLeftFace = GetRebarLegFace(2);
            rebarSet.LegFaces.Add(topLeftFace);

            var topRightFace = GetRebarLegFace(3);
            rebarSet.LegFaces.Add(topRightFace);

            var rightFace = GetRebarLegFace(4);
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

            Vector normal = Utility.GetVectorFromTwoPoints(ProfilePoints[0][4], ProfilePoints[1][4]);

            double firstDistance = Distance.PointToPoint(ProfilePoints[0][0], ProfilePoints[1][1]);
            double secondDistance = Distance.PointToPoint(ProfilePoints[0][1], ProfilePoints[1][0]);
            Point startPoint, endPoint;
            if (firstDistance > secondDistance)
            {
                startPoint = ProfilePoints[0][0];
                endPoint = ProfilePoints[1][1];
            }
            else
            {
                startPoint = ProfilePoints[0][1];
                endPoint = ProfilePoints[1][0];
            }

            GeometricPlane startPlane = new GeometricPlane(startPoint, normal);
            GeometricPlane endPlane = new GeometricPlane(endPoint, normal);
            Line longitudinal = new Line(ProfilePoints[0][4], ProfilePoints[1][4]);
            Point startIntersection = Utility.GetExtendedIntersection(longitudinal, startPlane, 2);
            Point endIntersection = Utility.GetExtendedIntersection(longitudinal, endPlane, 2);

            guideline.Curve.AddContourPoint(new ContourPoint(startIntersection, null));
            guideline.Curve.AddContourPoint(new ContourPoint(endIntersection, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
            LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 2, 2, 1 });
        }
        void BottomPerpendicularRebar()
        {
            int rebarSize = Convert.ToInt32(Program.ExcelDictionary["BPR_Diameter"]);
            string spacing = Program.ExcelDictionary["BPR_Spacing"];

            string name = CreateRebarName(RebarType.BPR);
            var rebarSet = CreateDefaultRebarSet(name, rebarSize);

            var legFace1 = GetRebarLegFace(5);
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

            Vector normal = Utility.GetVectorFromTwoPoints(ProfilePoints[0][4], ProfilePoints[1][4]);

            double firstDistance = Distance.PointToPoint(ProfilePoints[0][0], ProfilePoints[1][1]);
            double secondDistance = Distance.PointToPoint(ProfilePoints[0][1], ProfilePoints[1][0]);
            Point startPoint, endPoint;
            if (firstDistance > secondDistance)
            {
                startPoint = ProfilePoints[0][0];
                endPoint = ProfilePoints[1][1];
            }
            else
            {
                startPoint = ProfilePoints[0][1];
                endPoint = ProfilePoints[1][0];
            }

            GeometricPlane startPlane = new GeometricPlane(startPoint, normal);
            GeometricPlane endPlane = new GeometricPlane(endPoint, normal);
            Line longitudinal = new Line(ProfilePoints[0][4], ProfilePoints[1][4]);
            Point startIntersection = Utility.GetExtendedIntersection(longitudinal, startPlane, 2);
            Point endIntersection = Utility.GetExtendedIntersection(longitudinal, endPlane, 2);

            guideline.Curve.AddContourPoint(new ContourPoint(startIntersection, null));
            guideline.Curve.AddContourPoint(new ContourPoint(endIntersection, null));

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
            rightHookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
            rightHookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
            rightHookModifier.Insert();
            new Model().CommitChanges();

            if (horizontalOffset != 0)
            {
                var startHookModifier = new RebarEndDetailModifier
                {
                    Father = rebarSet,
                    EndType = RebarEndDetailModifier.EndTypeEnum.HOOK
                };
                startHookModifier.RebarHook.Shape = RebarHookData.RebarHookShapeEnum.HOOK_90_DEGREES;
                startHookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
                startHookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
                startHookModifier.Insert();

                var endHookModifier = new RebarEndDetailModifier
                {
                    Father = rebarSet,
                    EndType = RebarEndDetailModifier.EndTypeEnum.HOOK
                };
                endHookModifier.RebarHook.Shape = RebarHookData.RebarHookShapeEnum.HOOK_90_DEGREES;
                endHookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
                endHookModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
                endHookModifier.Insert();
            }

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
            LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2 });
        }
        void BottomLongitudinalRebar()
        {
            int rebarSize = Convert.ToInt32(Program.ExcelDictionary["BLR_Diameter"]);
            string spacing = Program.ExcelDictionary["BLR_Spacing"];

            string name = CreateRebarName(RebarType.BLR);
            var rebarSet = CreateDefaultRebarSet(name, rebarSize);

            var legFace1 = GetRebarLegFace(5);
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

            Vector normal = Utility.GetVectorFromTwoPoints(ProfilePoints[0][4], ProfilePoints[1][4]).GetNormal();
            Point origin = new Point((ProfilePoints[0][4].X + ProfilePoints[1][4].X) / 2.0, ProfilePoints[0][4].Y, ProfilePoints[0][4].Z);
            GeometricPlane geometricPlane = new GeometricPlane(origin, normal);

            Line leftLine = new Line(ProfilePoints[0][0], ProfilePoints[1][0]);
            Line rightLine = new Line(ProfilePoints[0][1], ProfilePoints[1][1]);

            Point startIntersection = Intersection.LineToPlane(leftLine, geometricPlane);
            Point rightIntersection = Intersection.LineToPlane(rightLine, geometricPlane);

            guideline.Curve.AddContourPoint(new ContourPoint(startIntersection, null));
            guideline.Curve.AddContourPoint(new ContourPoint(rightIntersection, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
            LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 3 });
        }
        void TopLongitudinalLeftRebar()
        {
            int rebarSize = Convert.ToInt32(Program.ExcelDictionary["TLR_Diameter"]);
            string spacing = Program.ExcelDictionary["TLR_Spacing"];

            string name = CreateRebarName(RebarType.TLLR);
            var rebarSet = CreateDefaultRebarSet(name, rebarSize);

            var legFace1 = GetRebarLegFace(2);
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

            Vector normal = Utility.GetVectorFromTwoPoints(ProfilePoints[0][4], ProfilePoints[1][4]).GetNormal();
            Point origin = new Point((ProfilePoints[0][4].X + ProfilePoints[1][4].X) / 2.0, ProfilePoints[0][4].Y, ProfilePoints[0][4].Z);
            GeometricPlane geometricPlane = new GeometricPlane(origin, normal);

            Line leftLine = new Line(ProfilePoints[0][2], ProfilePoints[1][2]);
            Line rightLine = new Line(ProfilePoints[0][4], ProfilePoints[1][4]);

            Point startIntersection = Intersection.LineToPlane(leftLine, geometricPlane);
            Point rightIntersection = Intersection.LineToPlane(rightLine, geometricPlane);

            guideline.Curve.AddContourPoint(new ContourPoint(startIntersection, null));
            guideline.Curve.AddContourPoint(new ContourPoint(rightIntersection, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
            LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 3 });
        }
        void TopLongitudinalRightRebar()
        {
            int rebarSize = Convert.ToInt32(Program.ExcelDictionary["TLR_Diameter"]);
            string spacing = Program.ExcelDictionary["TLR_Spacing"];

            string name = CreateRebarName(RebarType.TLRR);
            var rebarSet = CreateDefaultRebarSet(name, rebarSize);

            var legFace1 = GetRebarLegFace(3);
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

            Vector normal = Utility.GetVectorFromTwoPoints(ProfilePoints[0][4], ProfilePoints[1][4]).GetNormal();
            Point origin = new Point((ProfilePoints[0][4].X + ProfilePoints[1][4].X) / 2.0, ProfilePoints[0][4].Y, ProfilePoints[0][4].Z);
            GeometricPlane geometricPlane = new GeometricPlane(origin, normal);

            Line leftLine = new Line(ProfilePoints[0][4], ProfilePoints[1][4]);
            Line rightLine = new Line(ProfilePoints[0][3], ProfilePoints[1][3]);

            Point startIntersection = Intersection.LineToPlane(leftLine, geometricPlane);
            Point rightIntersection = Intersection.LineToPlane(rightLine, geometricPlane);

            guideline.Curve.AddContourPoint(new ContourPoint(startIntersection, null));
            guideline.Curve.AddContourPoint(new ContourPoint(rightIntersection, null));

            rebarSet.Guidelines.Add(guideline);
            bool succes = rebarSet.Insert();
            new Model().CommitChanges();

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
            LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 3 });
        }
        void ClosingCShapeRebar(int faceNumber)
        {
            //1 - StartLeft, 2- StarRight, 3 - EndLeft, 4 -EndRight
            int rebarSize = Convert.ToInt32(Program.ExcelDictionary["CR_Diameter"]);
            string spacing = Program.ExcelDictionary["CR_Spacing"];

            string name = CreateRebarName(RebarType.CCSR);
            var rebarSet = CreateDefaultRebarSet(name, rebarSize);

            Point bottomLeft, bottomRight;
            Point endBottomLeft, endTopLeft, endTopMid, endTopRight, endBottomRight;

            RebarLegFace mainFace;
            switch (faceNumber)
            {
                case 1:
                    mainFace = GetRebarLegFace(0);
                    bottomLeft = ProfilePoints[0][0];
                    bottomRight = ProfilePoints[0][1];
                    endBottomLeft = ProfilePoints[1][0];
                    endTopLeft = ProfilePoints[1][2];
                    endTopMid = ProfilePoints[1][4];
                    endBottomRight = ProfilePoints[1][1];
                    endTopRight = ProfilePoints[1][3];
                    break;
                default:
                    mainFace = GetRebarLegFace(6);
                    bottomLeft = ProfilePoints[1][0];
                    bottomRight = ProfilePoints[1][1];
                    endBottomLeft = ProfilePoints[0][0];
                    endTopLeft = ProfilePoints[0][2];
                    endTopMid = ProfilePoints[0][4];
                    endBottomRight = ProfilePoints[0][1];
                    endTopRight = ProfilePoints[0][3];
                    break;
            }

            rebarSet.LegFaces.Add(mainFace);

            var bottomFace = GetRebarLegFace(5);
            rebarSet.LegFaces.Add(bottomFace);

            var topFaceLeft = GetRebarLegFace(2);
            rebarSet.LegFaces.Add(topFaceLeft);

            var topFaceRight = GetRebarLegFace(3);
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

            Vector normal = Utility.GetVectorFromTwoPoints(ProfilePoints[0][4], ProfilePoints[1][4]).GetNormal();
            Point origin = new Point((ProfilePoints[0][4].X + ProfilePoints[1][4].X) / 2.0, ProfilePoints[0][4].Y, ProfilePoints[0][4].Z);
            GeometricPlane geometricPlane = new GeometricPlane(origin, normal);

            Line leftLine = new Line(bottomLeft, endBottomLeft);
            Line rightLine = new Line(bottomRight, endBottomRight);

            Point startIntersection = Intersection.LineToPlane(leftLine, geometricPlane);
            Point rightIntersection = Intersection.LineToPlane(rightLine, geometricPlane);

            guideline.Curve.AddContourPoint(new ContourPoint(startIntersection, null));
            guideline.Curve.AddContourPoint(new ContourPoint(rightIntersection, null));
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

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
            LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 1, 3, 3, 3 });
        }
        void ClosingLongitudinalRebar(int faceNumber)
        {
            //1 - Start, 2- Left, 3 - Right, 4 - End
            int rebarSize = Convert.ToInt32(Program.ExcelDictionary["CR_Diameter"]);
            string spacing = Program.ExcelDictionary["CR_Spacing"];

            string name = CreateRebarName(RebarType.CLR);
            var rebarSet = CreateDefaultRebarSet(name, rebarSize);

            Point bottomLeft, topLeft, bottomRight, topRight;
            Point endBottomLeft, endTopLeft, endBottomRight, endTopRight;
            Vector normal;
            Line bottomLine;
            Line topLine;
            switch (faceNumber)
            {
                case 1:
                    bottomLeft = ProfilePoints[0][0];
                    topLeft = ProfilePoints[0][2];
                    bottomRight = ProfilePoints[0][1];
                    topRight = ProfilePoints[0][3];
                    endBottomLeft = ProfilePoints[1][0];
                    endTopLeft = ProfilePoints[1][2];
                    endBottomRight = ProfilePoints[1][1];
                    endTopRight = ProfilePoints[1][3];
                    normal = Utility.GetVectorFromTwoPoints(bottomLeft, endBottomLeft).GetNormal();
                    bottomLine = new Line(bottomLeft, endBottomLeft);
                    topLine = new Line(topLeft, endTopLeft);
                    break;
                case 2:
                    bottomLeft = ProfilePoints[1][0];
                    topLeft = ProfilePoints[1][2];
                    bottomRight = ProfilePoints[0][0];
                    topRight = ProfilePoints[0][2];
                    endBottomLeft = ProfilePoints[1][1];
                    endTopLeft = ProfilePoints[1][3];
                    endBottomRight = ProfilePoints[0][1];
                    endTopRight = ProfilePoints[0][3];
                    normal = Utility.GetVectorFromTwoPoints(bottomLeft, bottomRight).GetNormal();
                    bottomLine = new Line(bottomLeft, bottomRight);
                    topLine = new Line(topLeft, topRight);
                    break;
                case 3:
                    bottomLeft = ProfilePoints[0][1];
                    topLeft = ProfilePoints[0][3];
                    bottomRight = ProfilePoints[1][1];
                    topRight = ProfilePoints[1][3];
                    endBottomLeft = ProfilePoints[0][0];
                    endTopLeft = ProfilePoints[0][2];
                    endBottomRight = ProfilePoints[1][0];
                    endTopRight = ProfilePoints[1][2];
                    normal = Utility.GetVectorFromTwoPoints(bottomLeft, bottomRight).GetNormal();
                    bottomLine = new Line(bottomLeft, bottomRight);
                    topLine = new Line(topLeft, topRight);
                    break;
                default:
                    bottomLeft = ProfilePoints[1][0];
                    topLeft = ProfilePoints[1][2];
                    bottomRight = ProfilePoints[1][1];
                    topRight = ProfilePoints[1][3];
                    endBottomLeft = ProfilePoints[0][0];
                    endTopLeft = ProfilePoints[0][2];
                    endBottomRight = ProfilePoints[0][1];
                    endTopRight = ProfilePoints[0][3];
                    normal = Utility.GetVectorFromTwoPoints(bottomLeft, endBottomLeft).GetNormal();
                    bottomLine = new Line(bottomLeft, endBottomLeft);
                    topLine = new Line(topLeft, endTopLeft);
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

            GeometricPlane geometricPlane = new GeometricPlane(bottomLeft, normal);

            Point startGL = Utility.GetExtendedIntersection(bottomLine, geometricPlane, 2);
            Point endGL = Utility.GetExtendedIntersection(topLine, geometricPlane, 2);

            guideline.Curve.AddContourPoint(new ContourPoint(startGL, null));
            guideline.Curve.AddContourPoint(new ContourPoint(endGL, null));

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

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod());
            LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2, 2, 2 });
        }
        string CreateRebarName(RebarType rebarType)
        {
            string name = "FTG_"+rebarType.ToString();
            return name;
        }
        #endregion

    }
}
