﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using Tekla.Structures;
using Tekla.Structures.Model;
using Tekla.Structures.Geometry3d;


namespace ZeroTouchTekla
{
    class RCLMN : Element
    {
        #region Fields
        enum RebarType
        {
            OS,
            IS,
            MR,
            IMR,
            TCR
        }
        double width;
        double height;
        double chamfer;
        double length;
        List<RebarLegFace> rebarLegFaces = new List<RebarLegFace>();      
        #endregion
        #region Constructor
        public RCLMN(params Part[] parts) : base()
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
        void SetProfileParameters(Part part)
        {
            Beam beam = part as Beam;
            string[] profileValues = GetProfileValues(beam);
            //FTG Width*FirstHeight*SecondHeight*AsymWidth
            double width = Convert.ToDouble(profileValues[0]);
            double height = Convert.ToDouble(profileValues[1]);
            double chamfer = Convert.ToDouble(profileValues[2]);
            double length = Distance.PointToPoint(beam.StartPoint, beam.EndPoint);
            length -= beam.StartPointOffset.Dx;
            length += beam.EndPointOffset.Dx;

            this.width = width;
            this.height = height;
            this.chamfer = chamfer;
            this.length = length;
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
            secondFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][3], null));
            secondFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            RebarLegFaces.Add(secondFace);

            RebarLegFace thirdFace = new RebarLegFace();
            thirdFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][3], null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
            thirdFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
            RebarLegFaces.Add(thirdFace);

            RebarLegFace fourthFace = new RebarLegFace();
            fourthFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
            fourthFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
            fourthFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
            fourthFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            RebarLegFaces.Add(fourthFace);

            RebarLegFace endFace = new RebarLegFace();
            endFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
            endFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
            endFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][3], null));
            endFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
            RebarLegFaces.Add(endFace);
        }
       
        #region RebarCreationMethods
        public override void Create()
        {
            OuterStirrups(false);
            if (Convert.ToInt32(Program.ExcelDictionary["OS_Doubled"]) == 1)
            {
                OuterStirrups(true);
            }
            MainRebar(1);
            MainRebar(2);
            MainRebar(3);
            MainRebar(4);
            if (Convert.ToInt32(Program.ExcelDictionary["MR_AddTopHooks"]) != 1)
            {
                TopClosingRebar(1);
                TopClosingRebar(2);
            }
            InnerStirrups();
            InnerMainRebar(1);
            InnerMainRebar(2);
            InnerMainRebar(3);
            InnerMainRebar(4);
        }
        public override void CreateSingle(string rebarName)
        {
            string rebarType = rebarName.Split('_')[1];

            RebarType rType;
            Enum.TryParse(rebarType, out rType);
            switch (rType)
            {
                case RebarType.OS:
                    bool isSecond = Convert.ToBoolean(rebarName.Split('_')[2]);
                    OuterStirrups(isSecond);
                    break;
                case RebarType.IS:
                    InnerStirrups();
                    break;
                case RebarType.MR:
                    int p1 = Convert.ToInt32(rebarName.Split('_')[2]);
                    MainRebar(p1);
                    break;
                case RebarType.IMR:
                    int p2 = Convert.ToInt32(rebarName.Split('_')[2]);
                    InnerMainRebar(p2);
                    break;
                case RebarType.TCR:
                    int p3 = Convert.ToInt32(rebarName.Split('_')[2]);
                    TopClosingRebar(p3);
                    break;
            }
        }
        void OuterStirrups(bool isSecond)
        {
            double startOffset = Convert.ToDouble(Program.ExcelDictionary["M_StartOffset"]);
            string rebarSize = Program.ExcelDictionary["OS_Diameter"];
            int variableSpacing = Convert.ToInt32(Program.ExcelDictionary["OS_VariableSpacing"]);
            int firstSpacing = Convert.ToInt32(Program.ExcelDictionary["OS_FirstSpacing"]);
            double firstSpacingLength = Convert.ToDouble(Program.ExcelDictionary["OS_FirstSpacingLength"]);
            int secondSpacing = Convert.ToInt32(Program.ExcelDictionary["OS_SecondSpacing"]);
            int doubled = Convert.ToInt32(Program.ExcelDictionary["OS_Doubled"]);
            double doubledOffset = Convert.ToDouble(Program.ExcelDictionary["OS_DoubledOffset"]);
            double width = this.width;
            double height = this.height;

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "CLMN_OS_" + isSecond;
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = TeklaUtils.SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = TeklaUtils.GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;            

            var leftFace = GetRebarLegFace(1);
            rebarSet.LegFaces.Add(leftFace);

            var bottomFace = GetRebarLegFace(2);
            rebarSet.LegFaces.Add(bottomFace);

            var rightFace = GetRebarLegFace(3);
            rebarSet.LegFaces.Add(rightFace);

            var topEndFace = GetRebarLegFace(4);
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

            guideline.Spacing.StartOffset = isSecond ? startOffset + 50 : startOffset;
            guideline.Spacing.EndOffset = 100;
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
            rebarSet.Guidelines.Add(guideline);

            int[] offsetArray = new int[] { 1, 1, 1, 1 };
            if (doubled == 1)
            {
                int faceToOffset;
                if (isSecond)
                {
                    if (width > height)
                    {
                        faceToOffset = 3;
                    }
                    else
                    {
                        faceToOffset = 2;
                    }
                }
                else
                {
                    if (width > height)
                    {
                        faceToOffset = 1;
                    }
                    else
                    {
                        faceToOffset = 4;
                    }
                }
                rebarSet.LegFaces[faceToOffset - 1].AdditonalOffset = doubledOffset;
                offsetArray[faceToOffset - 1] = 5;
            }

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

            rebarSet.SetUserProperty(RebarCreator.FATHER_ID_NAME, RebarCreator.FatherID);
            LayerDictionary.Add(rebarSet.Identifier.ID, offsetArray);
        }
        void InnerStirrups()
        {
            double startOffset = Convert.ToDouble(Program.ExcelDictionary["M_StartOffset"]);
            string rebarSize = Program.ExcelDictionary["IS_Diameter"];
            int firstSpacing = Convert.ToInt32(Program.ExcelDictionary["IS_Spacing"]);
            double spacingLength = Convert.ToDouble(Program.ExcelDictionary["IS_Length"]);

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "CLMN_IS";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = TeklaUtils.SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = TeklaUtils.GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            var leftFace = GetRebarLegFace(1);
            rebarSet.LegFaces.Add(leftFace);

            var bottomFace = GetRebarLegFace(2);
            rebarSet.LegFaces.Add(bottomFace);

            var rightFace = GetRebarLegFace(3);
            rebarSet.LegFaces.Add(rightFace);

            var topEndFace = GetRebarLegFace(4);
            rebarSet.LegFaces.Add(topEndFace);

            var guideline = new RebarGuideline();
            guideline.Spacing.Zones.Add(new RebarSpacingZone
            {
                Spacing = firstSpacing,
                SpacingType = RebarSpacingZone.SpacingEnum.EXACT,
                Length = 100,
                LengthType = RebarSpacingZone.LengthEnum.RELATIVE,
            });


            guideline.Spacing.StartOffset = startOffset;
            guideline.Spacing.EndOffset = 100;
            guideline.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            guideline.Curve.AddContourPoint(new ContourPoint(new Point(ProfilePoints[0][0].X + spacingLength, ProfilePoints[0][0].Y, ProfilePoints[0][0].Z), null));
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

            rebarSet.SetUserProperty(RebarCreator.FATHER_ID_NAME, RebarCreator.FatherID);
            LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 3, 3, 3, 3 });
        }
        void MainRebar(int faceNumber)
        {
            double startOffset = Convert.ToDouble(Program.ExcelDictionary["M_StartOffset"]);
            int rebarSize = Convert.ToInt32(Program.ExcelDictionary["MR_Diameter"]);
            int spacing = Convert.ToInt32(Program.ExcelDictionary["MR_Spacing"]);
            int addSplitter = Convert.ToInt32(Program.ExcelDictionary["MR_AddSplitter"]);
            int addTopHooks = Convert.ToInt32(Program.ExcelDictionary["MR_AddTopHooks"]);
            double topHookLength = Convert.ToDouble(Program.ExcelDictionary["MR_TopHooksLength"]);

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "CLMN_MR_" + faceNumber;
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = TeklaUtils.SetClass(rebarSize);
            rebarSet.RebarProperties.Size = rebarSize.ToString(); ;
            rebarSet.RebarProperties.BendingRadius = TeklaUtils.GetBendingRadious(rebarSize);
            rebarSet.LayerOrderNumber = 1;

            RebarLegFace mainFace;
            Point startPoint, endPoint;
            Point oStartPoint, oEndPoint;
            switch (faceNumber)
            {
                case 1:
                    mainFace = GetRebarLegFace(1);
                    startPoint = ProfilePoints[0][0];
                    endPoint = ProfilePoints[0][2];
                    oStartPoint = new Point(startPoint.X, startPoint.Y, startPoint.Z - 40 * rebarSize);
                    oEndPoint = new Point(endPoint.X, endPoint.Y, endPoint.Z - 40 * rebarSize);
                    break;
                case 2:
                    mainFace = GetRebarLegFace(2);
                    startPoint = ProfilePoints[0][2];
                    endPoint = ProfilePoints[0][3];
                    oStartPoint = new Point(startPoint.X, startPoint.Y + 40 * rebarSize, startPoint.Z);
                    oEndPoint = new Point(endPoint.X, endPoint.Y + 40 * rebarSize, endPoint.Z);
                    break;
                case 3:
                    mainFace = GetRebarLegFace(3);
                    startPoint = ProfilePoints[0][3];
                    endPoint = ProfilePoints[0][1];
                    oStartPoint = new Point(startPoint.X, startPoint.Y, startPoint.Z + 40 * rebarSize);
                    oEndPoint = new Point(endPoint.X, endPoint.Y, endPoint.Z + 40 * rebarSize);
                    break;
                default:
                    mainFace = GetRebarLegFace(4);
                    startPoint = ProfilePoints[0][1];
                    endPoint = ProfilePoints[0][0];
                    oStartPoint = new Point(startPoint.X, startPoint.Y - 40 * rebarSize, startPoint.Z);
                    oEndPoint = new Point(endPoint.X, endPoint.Y - 40 * rebarSize, endPoint.Z);
                    break;
            }


            if (addTopHooks == 1)
            {
                Vector vector = new Vector(topHookLength, 0, 0);
                Utility.StretchLegFace(ref mainFace, 1, vector);
                Utility.StretchLegFace(ref mainFace, 2, vector);

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
            guideline.Spacing.StartOffset = 150;
            guideline.Spacing.EndOffset = 150;
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

            rebarSet.SetUserProperty(RebarCreator.FATHER_ID_NAME, RebarCreator.FatherID);
            LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2, 2 });
        }
        void InnerMainRebar(int faceNumber)
        {
            double startOffset = Convert.ToDouble(Program.ExcelDictionary["M_StartOffset"]);
            int rebarSize = Convert.ToInt32(Program.ExcelDictionary["IMR_Diameter"]);
            int spacing = Convert.ToInt32(Program.ExcelDictionary["IMR_Spacing"]);
            double innerLength = Convert.ToInt32(Program.ExcelDictionary["IMR_Length"]);

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "CLMN_IMR_" + faceNumber;
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = TeklaUtils.SetClass(rebarSize);
            rebarSet.RebarProperties.Size = rebarSize.ToString(); ;
            rebarSet.RebarProperties.BendingRadius = TeklaUtils.GetBendingRadious(rebarSize);
            rebarSet.LayerOrderNumber = 1;

            RebarLegFace mainFace;
            Point topStartPoint, topEndPoint;
            Point startPoint, endPoint;
            Point oStartPoint, oEndPoint;
            switch (faceNumber)
            {
                case 1:
                    mainFace = GetRebarLegFace(1);
                    startPoint = ProfilePoints[0][0];
                    endPoint = ProfilePoints[0][2];
                    oStartPoint = new Point(startPoint.X, startPoint.Y, startPoint.Z + 40 * rebarSize);
                    oEndPoint = new Point(endPoint.X, endPoint.Y, endPoint.Z + 40 * rebarSize);
                    topStartPoint = ProfilePoints[1][0];
                    topEndPoint = ProfilePoints[1][2];
                    break;
                case 2:
                    mainFace = GetRebarLegFace(2);
                    startPoint = ProfilePoints[0][2];
                    endPoint = ProfilePoints[0][3];
                    oStartPoint = new Point(startPoint.X, startPoint.Y - 40 * rebarSize, startPoint.Z);
                    oEndPoint = new Point(endPoint.X, endPoint.Y - 40 * rebarSize, endPoint.Z);
                    topStartPoint = ProfilePoints[1][2];
                    topEndPoint = ProfilePoints[1][3];
                    break;
                case 3:
                    mainFace = GetRebarLegFace(3);
                    startPoint = ProfilePoints[0][3];
                    endPoint = ProfilePoints[0][1];
                    oStartPoint = new Point(startPoint.X, startPoint.Y, startPoint.Z - 40 * rebarSize);
                    oEndPoint = new Point(endPoint.X, endPoint.Y, endPoint.Z - 40 * rebarSize);
                    topStartPoint = ProfilePoints[1][3];
                    topEndPoint = ProfilePoints[1][1];
                    break;
                default:
                    mainFace = GetRebarLegFace(4);
                    startPoint = ProfilePoints[0][1];
                    endPoint = ProfilePoints[0][0];
                    oStartPoint = new Point(startPoint.X, startPoint.Y + 40 * rebarSize, startPoint.Z);
                    oEndPoint = new Point(endPoint.X, endPoint.Y + 40 * rebarSize, endPoint.Z);
                    topStartPoint = ProfilePoints[1][1];
                    topEndPoint = ProfilePoints[1][0];
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
            guideline.Spacing.StartOffset = 200;
            guideline.Spacing.EndOffset = 200;
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

            var topLengthModifier = new RebarEndDetailModifier();
            topLengthModifier.Father = rebarSet;
            topLengthModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            topLengthModifier.RebarLengthAdjustment.AdjustmentLength = startOffset + innerLength;
            topLengthModifier.Curve.AddContourPoint(new ContourPoint(topStartPoint, null));
            topLengthModifier.Curve.AddContourPoint(new ContourPoint(topEndPoint, null));
            topLengthModifier.Insert();
            new Model().CommitChanges();

            rebarSet.SetUserProperty(RebarCreator.FATHER_ID_NAME, RebarCreator.FatherID);
            LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 4, 2 });
        }
        void TopClosingRebar(int faceNumber)
        {
            int rebarSize = Convert.ToInt32(Program.ExcelDictionary["TCR_Diameter"]);
            int spacing = Convert.ToInt32(Program.ExcelDictionary["TCR_Spacing"]);

            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "CLMN_TCR_" + faceNumber;
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = TeklaUtils.SetClass(rebarSize);
            rebarSet.RebarProperties.Size = rebarSize.ToString(); ;
            rebarSet.RebarProperties.BendingRadius = TeklaUtils.GetBendingRadious(rebarSize);
            rebarSet.LayerOrderNumber = 1;

            RebarLegFace mainFace = rebarLegFaces[5];
            rebarSet.LegFaces.Add(mainFace);

            RebarLegFace leftFace, rightFace;
            Point startPoint, endPoint;
            Point leftFaceSP, leftFaceEP, rightFaceSP, rightFaceEP;
            switch (faceNumber)
            {
                case 1:
                    leftFace = GetRebarLegFace(1);
                    rightFace = GetRebarLegFace(3);
                    startPoint = ProfilePoints[1][0];
                    endPoint = ProfilePoints[1][2];
                    leftFaceSP = ProfilePoints[0][0];
                    leftFaceEP = ProfilePoints[0][2];
                    rightFaceSP = ProfilePoints[0][1];
                    rightFaceEP = ProfilePoints[0][3];
                    break;
                default:
                    leftFace = GetRebarLegFace(2);
                    rightFace = rebarLegFaces[4];
                    startPoint = ProfilePoints[0][1];
                    endPoint = ProfilePoints[0][0];
                    leftFaceSP = ProfilePoints[0][2];
                    leftFaceEP = ProfilePoints[0][3];
                    rightFaceSP = ProfilePoints[0][0];
                    rightFaceEP = ProfilePoints[0][1];
                    break;
            }


            rebarSet.LegFaces.Add(leftFace);
            rebarSet.LegFaces.Add(rightFace);

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

            var leftFaceEndModifier = new RebarEndDetailModifier();
            leftFaceEndModifier.Father = rebarSet;
            leftFaceEndModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            leftFaceEndModifier.RebarLengthAdjustment.AdjustmentLength = 10 * Convert.ToInt32(rebarSize);
            leftFaceEndModifier.Curve.AddContourPoint(new ContourPoint(leftFaceSP, null));
            leftFaceEndModifier.Curve.AddContourPoint(new ContourPoint(leftFaceEP, null));
            leftFaceEndModifier.Insert();

            var rightFaceEndModifier = new RebarEndDetailModifier();
            rightFaceEndModifier.Father = rebarSet;
            rightFaceEndModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            rightFaceEndModifier.RebarLengthAdjustment.AdjustmentLength = 10 * Convert.ToInt32(rebarSize);
            rightFaceEndModifier.Curve.AddContourPoint(new ContourPoint(rightFaceSP, null));
            rightFaceEndModifier.Curve.AddContourPoint(new ContourPoint(rightFaceEP, null));
            rightFaceEndModifier.Insert();
            new Model().CommitChanges();

            rebarSet.SetUserProperty(RebarCreator.FATHER_ID_NAME, RebarCreator.FatherID);
            LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { faceNumber, 2, 2 });
        }
        #endregion
        #region Properties
        #endregion

    }
}
