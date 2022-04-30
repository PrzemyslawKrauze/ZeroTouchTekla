using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model;

namespace ZeroTouchTekla.Profiles
{
    public class WING : Element
    {
        List<List<Point>> CompomentProfilePoints = new List<List<Point>>();
        public WING(Part part)
        {
            base.BaseParts = new Part[] { part };
            SetLocalPlane();
            GetProfilePointsAndParameters(part);
            CompomentProfilePoints = TeklaUtils.GetSortedPointsFromPart(part);
        }

        public static void GetProfilePointsAndParameters(Part part)
        {
            Beam beam = part as Beam;
            Profile profile = beam.Profile;
            string profileName = profile.ProfileString;
            profileName = Regex.Replace(profileName, "[0-9]", "");
            profileName = profileName.Replace("*", "");
            if (profileName.Contains("RTWS"))
            {
                _profileType = ProfileType.RTWS;
            }
            else
            {
                _profileType = ProfileType.RTW;
            }
            BaseComponent baseComponent = beam.GetFatherComponent();
            System.Collections.Hashtable hashtable = new System.Collections.Hashtable();
            baseComponent.GetAllUserProperties(ref hashtable);
            _soch = (double)hashtable["SOCH"];
            _scd = (double)hashtable["SCD"];
            _sich = (double)hashtable["SICH"];
            _ctth = (double)hashtable["CtTH"];
            _cd = (double)hashtable["CD"];
            _ctbh = (double)hashtable["CtBH"];
        }
        public override void Create()
        {
            Model model = new Model();
            string componentName =  BaseParts[0].GetFatherComponent().Name;
            Element element;
            if (_profileType == ProfileType.RTW)
            {
                RTW rtw = new RTW(BaseParts[0]);
                rtw.Create();
                element = rtw;
            }
            else
            {
                RTWS rtws = new RTWS(BaseParts[0]);
                rtws.Create();
                element = rtws;
            }

            Type[] Types = new Type[] { typeof(RebarSet) };
            ModelObjectEnumerator moe = model.GetModelObjectSelector().GetAllObjectsWithType(Types);
            var rebarList = Utility.ToList(moe);

            List<RebarSet> selectedRebars = (from RebarSet r in rebarList
                                             where Utility.GetUserProperty(r, RebarCreator.FATHER_ID_NAME) == BaseParts[0].Identifier.ID
                                             select r).ToList();

            if (_profileType == ProfileType.RTW)
            {
                RTW rtw = element as RTW;
                if (componentName.Contains("Reversed"))
                {
                    SetRTWProfilePointsReversed(element);
                    _isReversed = true;
                }
                else
                {
                    SetRTWProfilePoints(element);
                    _isReversed = false;
                }
                EditRTWRebar(rtw, selectedRebars);
            }
            else
            {
                RTWS rtws = element as RTWS;
                if (componentName.Contains("Reversed"))
                {
                    SetRTWSProfilePointsReversed(element);
                    _isReversed = true;
                }
                else
                {
                    SetRTWSProfilePoints(element);
                    _isReversed = false;
                }
                EditRTWSRebar(rtws, selectedRebars);
            }
        }
        public override void CreateSingle(string barName)
        {
            throw new NotImplementedException();
        }       
        #region RTWMethods
        void EditRTWRebar(RTW rtw, List<RebarSet> rebarSets)
        {
            foreach (RebarSet rs in rebarSets)
            {
                string methodName = string.Empty;
                rs.GetUserProperty(RebarCreator.METHOD_NAME, ref methodName);
                switch (methodName)
                {
                    
                    case "InnerVerticalRebar":
                        RTWEditInnerVerticalRebar(rtw, rs);
                        break;
                    
                    case "OuterVerticalRebar":
                        RTWEditOuterVerticalRebar(rtw, rs);
                        break;
                        
                    case "InnerLongitudinalRebar":
                        RTWEditInnerLongitudinalRebar(rtw, rs);
                        break;
                        
                    case "OuterLongitudinalRebar":
                        RTWEditOuterLongitudinalRebar(rtw, rs);
                        break;
                        
                    case "CornicePerpendicularRebar":
                        RTWEditCornicePerpendicularRebar(rtw, rs);
                        break;
                        
                    case "CorniceLongitudinalRebar":
                        RTWEditCorniceLongitudinalRebar(rtw, rs);
                        break;
                        
                    case "ClosingCShapeRebar":
                        int n1 = 0;
                        rs.GetUserProperty(RebarCreator.MethodInput, ref n1);
                        RTWEditClosingCShapeRebar(rtw, rs, n1);
                        break;
                        
                    case "ClosingLongitudinalRebar":
                        int n2 = 0;
                        rs.GetUserProperty(RebarCreator.MethodInput, ref n2);
                        RTWEditClosingLongitudinalRebar(rtw, rs, n2);
                        RTWClosingLongitudinalRebarBottom(n2);
                        break;
                        
                    case "CShapeRebar":
                        RTWEditCShapeRebar(rtw, rs);
                        break;
                        
                }
            }
        }
        void SetRTWProfilePoints(Element element)
        {
            ProfilePoints = CompomentProfilePoints;
        }
        void SetRTWProfilePointsReversed(Element element)
        {
            CompomentProfilePoints.Reverse();
            ProfilePoints = CompomentProfilePoints;
        }
        void RTWEditOuterVerticalRebar(RTW rtw, RebarSet rebarSet)
        {
            List<RebarLegFace> rebarLegFaces = rebarSet.LegFaces;

            Point correctedP32 = new Point(ProfilePoints[3][2].X, ProfilePoints[3][2].Y + rtw.CorniceHeight, ProfilePoints[3][2].Z);
            Point correctedP12 = new Point(ProfilePoints[1][2].X, ProfilePoints[1][2].Y + rtw.CorniceHeight, ProfilePoints[1][2].Z);

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][1], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][3], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][1], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(correctedP32, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(correctedP12, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            rebarLegFaces[0] = mainFace;

            Point offsetedStartPoint = new Point(ProfilePoints[0][1].X, ProfilePoints[0][1].Y, ProfilePoints[0][1].Z + 1000);
            Point offsetedEndPoint = new Point(ProfilePoints[2][1].X, ProfilePoints[2][1].Y, ProfilePoints[2][1].Z + 1000);
            var bottomFace = new RebarLegFace();
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][1], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(offsetedEndPoint, null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(offsetedStartPoint, null));
            rebarLegFaces[1] = bottomFace;
            bool succes = rebarSet.Modify();

            ModelObjectEnumerator modelObjectEnumerator = rebarSet.GetRebarModifiers();
            List<ModelObject> modelObjectList = Utility.ToList(modelObjectEnumerator);
            RebarPropertyModifier rebarPropertyModifier = (from mo in modelObjectList
                                                           where mo.GetType() == typeof(RebarPropertyModifier)
                                                           select mo as RebarPropertyModifier).FirstOrDefault();
            if (rebarPropertyModifier != null)
            {
                var contour = new Contour();
                contour.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
                contour.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
                contour.AddContourPoint(new ContourPoint(correctedP12, null));
                contour.AddContourPoint(new ContourPoint(correctedP32, null));
                rebarPropertyModifier.Curve = contour;
                rebarPropertyModifier.Modify();
            }

            new Model().CommitChanges();
        }
        void RTWEditInnerVerticalRebar(RTW rtw, RebarSet rebarSet)
        {
            List<RebarLegFace> rebarLegFaces = rebarSet.LegFaces;

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][2], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][4], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            rebarLegFaces[0] = mainFace;

            Point offsetedStartPoint = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y, ProfilePoints[0][0].Z - 1000);
            Point offsetedEndPoint = new Point(ProfilePoints[2][0].X, ProfilePoints[2][0].Y, ProfilePoints[2][0].Z - 1000);
            var bottomFace = new RebarLegFace();
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][0], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(offsetedEndPoint, null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(offsetedStartPoint, null));
            rebarLegFaces[1] = bottomFace;

            ModelObjectEnumerator modelObjectEnumerator = rebarSet.GetRebarModifiers();
            List<ModelObject> modelObjectList = Utility.ToList(modelObjectEnumerator);
            RebarPropertyModifier rebarPropertyModifier = (from mo in modelObjectList
                                                           where mo.GetType() == typeof(RebarPropertyModifier)
                                                           select mo as RebarPropertyModifier).FirstOrDefault();
            if (rebarPropertyModifier != null)
            {
                var contour = new Contour();
                contour.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
                contour.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
                contour.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
                contour.AddContourPoint(new ContourPoint(ProfilePoints[3][4], null));
                rebarPropertyModifier.Curve = contour;
                rebarPropertyModifier.Modify();
            }

            bool succes = rebarSet.Modify();
            new Model().CommitChanges();
        }
        void RTWEditOuterLongitudinalRebar(RTW rtw, RebarSet rebarSet)
        {
            List<RebarLegFace> rebarLegFaces = rebarSet.LegFaces;

            Point correctedP32 = new Point(ProfilePoints[3][2].X, ProfilePoints[3][2].Y + rtw.CorniceHeight, ProfilePoints[3][2].Z);
            Point correctedP12 = new Point(ProfilePoints[1][2].X, ProfilePoints[1][2].Y + rtw.CorniceHeight, ProfilePoints[1][2].Z);

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][1], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][3], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][1], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(correctedP32, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(correctedP12, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            rebarLegFaces[0] = mainFace;

            ModelObjectEnumerator modelObjectEnumerator = rebarSet.GetRebarModifiers();
            List<ModelObject> modelObjectList = Utility.ToList(modelObjectEnumerator);
            RebarPropertyModifier rebarPropertyModifier = (from mo in modelObjectList
                                                           where mo.GetType() == typeof(RebarPropertyModifier)
                                                           select mo as RebarPropertyModifier).FirstOrDefault();
            if (rebarPropertyModifier != null)
            {
                double startOffset = Convert.ToDouble(Program.ExcelDictionary["ILR_StartOffset"]);
                double firstLength = Convert.ToDouble(Program.ExcelDictionary["ILR_SecondDiameterLength"]);
                Point secondPoint = new Point(ProfilePoints[0][1].X, ProfilePoints[0][1].Y + startOffset + firstLength, ProfilePoints[0][1].Z);

                var contour = new Contour();
                contour.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
                contour.AddContourPoint(new ContourPoint(secondPoint, null));
                rebarPropertyModifier.Curve = contour;
                rebarPropertyModifier.Modify();
            }

            var bottomFace = new RebarLegFace();
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            rebarLegFaces.Add(bottomFace);

            var midFace = new RebarLegFace();
            midFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            midFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            midFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
            midFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
            rebarLegFaces.Add(midFace);

            Point corrected21Point = new Point(ProfilePoints[1][2].X, ProfilePoints[1][2].Y + rtw.CorniceHeight, ProfilePoints[1][2].Z);
            var topFace = new RebarLegFace();
            topFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
            topFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
            topFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
            topFace.Contour.AddContourPoint(new ContourPoint(corrected21Point, null));
            rebarLegFaces.Add(topFace);

            var innerEndDetailModifier = new RebarEndDetailModifier();
            innerEndDetailModifier.Father = rebarSet;
            innerEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            innerEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = 300;
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
            innerEndDetailModifier.Insert();

            bool succes = rebarSet.Modify();
            new Model().CommitChanges();
            LayerDictionary[rebarSet.Identifier.ID] = new int[] { 2, 2, 2, 2 };
        }
        void RTWEditInnerLongitudinalRebar(RTW rtw, RebarSet rebarSet)
        {
            List<RebarLegFace> rebarLegFaces = rebarSet.LegFaces;

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][2], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][4], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            rebarLegFaces[0] = mainFace;

            if (_isReversed)
            {
                ModelObjectEnumerator modelObjectEnumerator = rebarSet.GetRebarModifiers();
                List<ModelObject> modelObjectList = Utility.ToList(modelObjectEnumerator);
                RebarPropertyModifier rebarPropertyModifier = (from mo in modelObjectList
                                                               where mo.GetType() == typeof(RebarPropertyModifier)
                                                               select mo as RebarPropertyModifier).FirstOrDefault();
                if (rebarPropertyModifier != null)
                {

                    double startOffset = Convert.ToDouble(Program.ExcelDictionary["ILR_StartOffset"]);
                    double firstLength = Convert.ToDouble(Program.ExcelDictionary["ILR_SecondDiameterLength"]);
                    Point origin = new Point(ProfilePoints[0][1].X, ProfilePoints[0][1].Y + startOffset + firstLength, ProfilePoints[0][1].Z);
                    GeometricPlane plane = new GeometricPlane(origin, new Vector(0, 1, 0));
                    Line line = new Line(ProfilePoints[0][0], ProfilePoints[0][2]);
                    Point intersection = Utility.GetExtendedIntersection(line, plane, 1);

                    var contour = new Contour();
                    contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
                    contour.AddContourPoint(new ContourPoint(intersection, null));
                    rebarPropertyModifier.Curve = contour;
                    rebarPropertyModifier.Modify();
                }
            }           

            bool succes = rebarSet.Modify();
            new Model().CommitChanges();            
        }
        void RTWEditCornicePerpendicularRebar(RTW rtw, RebarSet rebarSet)
        {
            RebarGuideline rebarGuideline = rebarSet.Guidelines.FirstOrDefault();
            var contour = new Contour();
            contour.AddContourPoint(new ContourPoint(ProfilePoints[1][5], null));
            contour.AddContourPoint(new ContourPoint(ProfilePoints[3][5], null));
            rebarGuideline.Curve = contour;

            bool succes = rebarSet.Modify();
            new Model().CommitChanges();
        }
        void RTWEditCorniceLongitudinalRebar(RTW rtw, RebarSet rebarSet)
        {
            List<RebarLegFace> rebarLegFaces = rebarSet.LegFaces;

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][5], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][5], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][3], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][3], null));
            rebarLegFaces[0] = mainFace;

            bool succes = rebarSet.Modify();
            new Model().CommitChanges();
        }
        void RTWEditClosingCShapeRebar(RTW rtw, RebarSet rebarSet, int number)
        {
            if (number == 1)
            {
                rebarSet.Delete();
                return;
            }

            Point correctedP32 = new Point(ProfilePoints[3][2].X, ProfilePoints[3][2].Y + rtw.CorniceHeight, ProfilePoints[3][2].Z);
            Point correctedP12 = new Point(ProfilePoints[1][2].X, ProfilePoints[1][2].Y + rtw.CorniceHeight, ProfilePoints[1][2].Z);

            List<RebarLegFace> correctedLegFaces = new List<RebarLegFace>();
            var bottomFace = new RebarLegFace();
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][0], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][1], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][3], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][2], null));
            correctedLegFaces.Add(bottomFace);

            var midFace = new RebarLegFace();
            midFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][2], null));
            midFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][3], null));
            midFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][1], null));
            midFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][0], null));
            correctedLegFaces.Add(midFace);

            var topFace = new RebarLegFace();
            topFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][4], null));
            topFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][0], null));
            topFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][1], null));
            topFace.Contour.AddContourPoint(new ContourPoint(correctedP32, null));
            correctedLegFaces.Add(topFace);

            var innerFace = new RebarLegFace();
            innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
            innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][1], null));
            innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][3], null));
            innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][1], null));
            innerFace.Contour.AddContourPoint(new ContourPoint(correctedP32, null));
            innerFace.Contour.AddContourPoint(new ContourPoint(correctedP12, null));
            innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
            innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            correctedLegFaces.Add(innerFace);

            var outerFace = new RebarLegFace();
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][0], null));
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][2], null));
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][0], null));
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][4], null));
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            correctedLegFaces.Add(outerFace);

            rebarSet.LegFaces = correctedLegFaces;

            ModelObjectEnumerator modelObjectEnumerator = rebarSet.GetRebarModifiers();
            List<ModelObject> modelObjectList = Utility.ToList(modelObjectEnumerator);
            List<RebarEndDetailModifier> rebarEndDetail = (from mo in modelObjectList
                                                     where mo.GetType() == typeof(RebarEndDetailModifier)
                                                     select mo as RebarEndDetailModifier).ToList();

            if (rebarEndDetail.Count>1)
            {
                var contour = new Contour();
                contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
                contour.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
                contour.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
                contour.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
                rebarEndDetail[0].Curve = contour;
                rebarEndDetail[0].Modify();

                var secondContour = new Contour();
                secondContour.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
                secondContour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
                secondContour.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
                secondContour.AddContourPoint(new ContourPoint(correctedP12, null));
                rebarEndDetail[1].Curve = secondContour;
                rebarEndDetail[1].Modify();
            }

            bool succes = rebarSet.Modify();
            new Model().CommitChanges();

            LayerDictionary[rebarSet.Identifier.ID] = new int[] { 1, 1, 1, 2, 2 };
        }
        void RTWEditClosingLongitudinalRebar(RTW rtw, RebarSet rebarSet, int number)
        {
            if (number == 1)
            {
                rebarSet.Delete();
                return;
            }

            Point correctedP32 = new Point(ProfilePoints[3][2].X, ProfilePoints[3][2].Y + rtw.CorniceHeight, ProfilePoints[3][2].Z);

            List<RebarLegFace> correctedLegFaces = new List<RebarLegFace>();

            var midFace = new RebarLegFace();
            midFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][2], null));
            midFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][3], null));
            midFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][1], null));
            midFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][0], null));
            correctedLegFaces.Add(midFace);

            var topFace = new RebarLegFace();
            topFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][4], null));
            topFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][0], null));
            topFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][1], null));
            topFace.Contour.AddContourPoint(new ContourPoint(correctedP32, null));
            correctedLegFaces.Add(topFace);

            rebarSet.LegFaces = correctedLegFaces;

            var bottomLengthModifier = new RebarEndDetailModifier();
            bottomLengthModifier.Father = rebarSet;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.END_OFFSET;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentLength = 40 * Convert.ToDouble(rebarSet.RebarProperties.Size);
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[2][2], null));
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[2][3], null));
            bottomLengthModifier.Insert();

            bool succes = rebarSet.Modify();
            new Model().CommitChanges();

            LayerDictionary[rebarSet.Identifier.ID] = new int[] { 2, 2 };
        }
        void RTWClosingLongitudinalRebarBottom(int number)
        {
            if (number == 1)
            {
                return;
            }
            string rebarSize = Program.ExcelDictionary["CLR_Diameter"];
            string spacing = Program.ExcelDictionary["CLR_Spacing"];
            var rebarSet = new RebarSet();
            rebarSet.RebarProperties.Name = "RTW_CLR_" + "B";
            rebarSet.RebarProperties.Grade = "B500SP";
            rebarSet.RebarProperties.Class = TeklaUtils.SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = TeklaUtils.GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            Point leftBottom, rightBottom, rightTop, leftTop;

            leftBottom = ProfilePoints[2][0];
            rightBottom = ProfilePoints[2][1];
            rightTop = ProfilePoints[2][3];
            leftTop = ProfilePoints[2][2];


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

            var bottomLengthModifier = new RebarEndDetailModifier();
            bottomLengthModifier.Father = rebarSet;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.END_OFFSET;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentLength = 40 * Convert.ToDouble(rebarSet.RebarProperties.Size);
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[2][2], null));
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[2][3], null));
            bool inserted = bottomLengthModifier.Insert();


            new Model().CommitChanges();

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod(), number);
            LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2 });
        }
        void RTWEditCShapeRebar(RTW rtw, RebarSet rebarSet)
        {
            RebarGuideline gl = rebarSet.Guidelines.FirstOrDefault();
            System.Collections.ArrayList arrayList = gl.Curve.ContourPoints;
            var startPoint = arrayList[0] as ContourPoint;
            var endPoint = arrayList[1] as ContourPoint;
            GeometricPlane plane;

            if (startPoint.Y <= ProfilePoints[2][3].Y)
            {
                Vector xAxis = Utility.GetVectorFromTwoPoints(ProfilePoints[2][1], ProfilePoints[2][3]);
                Vector yAxis = Utility.GetVectorFromTwoPoints(ProfilePoints[2][1], ProfilePoints[2][0]);
                plane = new GeometricPlane(ProfilePoints[2][1], xAxis, yAxis);
            }
            else
            {
                Vector xAxis = Utility.GetVectorFromTwoPoints(ProfilePoints[2][3], ProfilePoints[3][1]);
                Vector yAxis = Utility.GetVectorFromTwoPoints(ProfilePoints[2][3], ProfilePoints[2][2]);
                plane = new GeometricPlane(ProfilePoints[2][3], xAxis, yAxis);
            }

            Line glLine = new Line(startPoint, endPoint);
            Point correctedStartPoint = Utility.GetExtendedIntersection(glLine, plane, 2);

            int number = _isReversed ? 0 : 1;
            gl.Curve.ContourPoints[number] = new ContourPoint(correctedStartPoint, null);
            if (_isReversed)
            {
                gl.Spacing.StartOffset = 300;
            }
            else
            {
                gl.Spacing.EndOffset = 300;
            }

            if (startPoint.Y >= ProfilePoints[0][2].Y)
            {
                Vector xAxis2 = new Vector(0, 1, 0);
                Vector yAxis2 = new Vector(0, 0, 1);
                GeometricPlane secondPlane = new GeometricPlane(ProfilePoints[1][1], xAxis2, yAxis2);
                Point correctedEndPoint = Utility.GetExtendedIntersection(glLine, secondPlane, 2);
                int secondNumber = _isReversed ? 1 : 0;
                gl.Curve.ContourPoints[secondNumber] = new ContourPoint(correctedEndPoint, null);
            }

            bool succes = rebarSet.Modify();
            new Model().CommitChanges();
        }
        #endregion
        
        #region RTWSMethods
        void EditRTWSRebar(RTWS rtws, List<RebarSet> rebarSets)
        {
            foreach (RebarSet rs in rebarSets)
            {
                string methodName = string.Empty;
                rs.GetUserProperty(RebarCreator.METHOD_NAME, ref methodName);
                switch (methodName)
                {
                    
                    case "OuterVerticalRebar":
                        RTWSEditOuterVerticalRebar(rtws, rs);
                        break;                        
                    case "BottomInnerVerticalRebar":
                        RTWSEditBottomInnerVerticalRebar(rtws, rs);
                        break;                        
                    case "TopInnerVerticalRebar":
                        RTWSEditTopInnerVerticalRebar(rtws, rs);
                        break;                        
                    case "SkewVerticalRebar":
                        RTWSEditSkewVerticalRebar(rtws, rs);
                        break;
                        
                    case "OuterLongitudinalRebar":
                        RTWSEditOuterLongitudinalRebar(rtws, rs);
                        break;
                        
                    case "TopInnerLongitudinalRebar":
                        RTWSEditTopInnerLongitudinalRebar(rtws, rs);
                        break;
                        
                    case "SkewLongitudinalRebar":
                        RTWSEditSkewLongitudinalRebar(rtws, rs);
                        break;                        
                    case "BottomInnerLongitudinalRebar":
                        RTWSEditBottomInnerLongitudinalRebar(rtws, rs);
                        break;
                        
                    case "CornicePerpendicularRebar":
                        RTWSEditCornicePerpendicularRebar(rtws, rs);
                        break;
                        
                    case "CorniceLongitudinalRebar":
                        RTWSEditCorniceLongitudinalRebar(rtws, rs);
                        break;
                        
                    case "BottomClosingCShapeRebar":
                        int n1 = 0;
                        rs.GetUserProperty(RebarCreator.MethodInput, ref n1);
                        RTWSEditBottomClosingCShapeRebar(rtws, rs,n1);
                        break;
                        
                    case "SkewClosingCShapeRebar":
                        int n2 = 0;
                        rs.GetUserProperty(RebarCreator.MethodInput, ref n2);
                        RTWSEditSkewClosingCShapeRebar(rtws, rs, n2);
                        break;                        
                    case "TopClosingCShapeRebar":
                        int n3 = 0;
                        rs.GetUserProperty(RebarCreator.MethodInput, ref n3);
                        RTWSEditTopClosingCShapeRebar(rtws, rs, n3);
                        break;
                        /*
                    case "ClosingLongitudinalRebar":
                        int n4 = 0;
                        rs.GetUserProperty(RebarCreator.MethodInput, ref n4);
                        RTWSEditClosingLongitudinalRebar(rtws, rs, n4);
                        break;
                    case "BottomCShapeRebar":
                        RTWSEditBottomCShapeRebar(rtws, rs);
                        break;
                    case "TopCShapeRebar":
                        RTWSEditTopCShapeRebar(rtws, rs);
                        break;
                    */
                }
            }
        }
        void SetRTWSProfilePoints(Element element)
        {
            ProfilePoints = CompomentProfilePoints;
        }
        void SetRTWSProfilePointsReversed(Element element)
        {
            RTW rtws = element as RTW;
            double correctedHeight1 = rtws.Height2 - (rtws.Height2 - rtws.Height) * (SCD) / rtws.Length;
            double correctedHeight2 = rtws.Height2 - (rtws.Height2 - rtws.Height) * (rtws.Length - CD) / rtws.Length;
            double s = (rtws.BottomWidth - (rtws.TopWidth - rtws.CorniceWidth)) / rtws.Height;
            double secondBottomWidth = rtws.TopWidth - rtws.CorniceWidth + rtws.Height2 * s;

            List<List<Point>> profilePoints = rtws.ProfilePoints;
            List<List<Point>> correctedPoints = new List<List<Point>>();

            Point p00 = profilePoints[1][0];
            Point p01 = new Point(p00.X, p00.Y + rtws.Height - SOCH, p00.Z);
            Point p03 = profilePoints[1][5];
            Point p02 = new Point(p00.X, p01.Y, p01.Z - secondBottomWidth + s * (rtws.Height - SOCH));
            correctedPoints.Add(new List<Point> { p00, p01, p02, p03 });

            Point p10 = new Point(p00.X - SCD, p00.Y, p00.Z);
            Point p11 = new Point(p10.X, p10.Y + correctedHeight1 - SICH, p10.Z);
            Point p12 = new Point(p10.X, p11.Y, profilePoints[1][4].Z - s * SICH);
            Point p13 = new Point(p10.X, p03.Y, profilePoints[1][4].Z - s * correctedHeight1);
            correctedPoints.Add(new List<Point> { p10, p11, p12, p13 });

            Point p20 = p10;
            Point p21 = new Point(p20.X, p20.Y + correctedHeight1 - rtws.CorniceHeight, p20.Z);
            Point p22 = new Point(p21.X, p21.Y, p21.Z + rtws.CorniceWidth);
            Point p23 = new Point(p22.X, p22.Y + rtws.CorniceHeight, p22.Z);
            Point p24 = new Point(p23.X, p23.Y, p23.Z - rtws.TopWidth);
            Point p25 = new Point(p20.X, p13.Y, p24.Z - s * correctedHeight1);
            correctedPoints.Add(new List<Point> { p20, p21, p22, p23, p24, p25 });

            Point p30 = new Point(profilePoints[0][0].X + CD, p20.Y, p20.Z);
            Point p31 = new Point(p30.X, p30.Y + correctedHeight2 - rtws.CorniceHeight, p30.Z);
            Point p32 = new Point(p30.X, p31.Y, p31.Z + rtws.CorniceWidth);
            Point p33 = new Point(p32.X, p32.Y + rtws.CorniceHeight, p32.Z);
            Point p34 = new Point(p33.X, p33.Y, p33.Z - rtws.TopWidth);
            Point p35 = new Point(p30.X, p25.Y, p34.Z - s * correctedHeight2);
            correctedPoints.Add(new List<Point> { p30, p31, p32, p33, p34, p35 });

            Point p40 = new Point(p30.X, p30.Y + CtBH, p30.Z);
            Point p41 = new Point(p40.X, p30.Y + correctedHeight2 - rtws.CorniceHeight, p40.Z);
            Point p42 = new Point(p41.X, p41.Y, p41.Z + rtws.CorniceWidth);
            Point p43 = new Point(p42.X, p42.Y + rtws.CorniceHeight, p42.Z);
            Point p44 = new Point(p43.X, p43.Y, p43.Z - rtws.TopWidth);
            Point p45 = new Point(p30.X, p40.Y, p44.Z - s * (correctedHeight2 - CtBH));
            correctedPoints.Add(new List<Point> { p40, p41, p42, p43, p44, p45 });

            Point p50 = new Point(profilePoints[0][0].X, p30.Y + rtws.Height - CtTH, p30.Z);
            Point p51 = new Point(p50.X, p30.Y + rtws.Height - rtws.CorniceHeight, p50.Z);
            Point p52 = new Point(p51.X, p51.Y, p51.Z + rtws.CorniceWidth);
            Point p53 = new Point(p52.X, p52.Y + rtws.CorniceHeight, p52.Z);
            Point p54 = new Point(p53.X, p53.Y, p53.Z - rtws.TopWidth);
            Point p55 = new Point(p50.X, p50.Y, p54.Z - s * CtTH);
            correctedPoints.Add(new List<Point> { p50, p51, p52, p53, p54, p55 });

            ProfilePoints = correctedPoints;
        }
        void RTWSEditOuterVerticalRebar(RTWS rtws,RebarSet rebarSet)
        {
            List<RebarLegFace> rebarLegFaces = rebarSet.LegFaces;

            Point correctedP32 = new Point(ProfilePoints[3][2].X, ProfilePoints[3][2].Y + rtws.CorniceHeight, ProfilePoints[3][2].Z);
            Point correctedP12 = new Point(ProfilePoints[1][2].X, ProfilePoints[1][2].Y + rtws.CorniceHeight, ProfilePoints[1][2].Z);

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][1], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][4], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][1], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(correctedP32, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(correctedP12, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][5], null));
            rebarLegFaces[0] = mainFace;

            Point offsetedStartPoint = new Point(ProfilePoints[0][1].X, ProfilePoints[0][1].Y, ProfilePoints[0][1].Z + 1000);
            Point offsetedEndPoint = new Point(ProfilePoints[2][1].X, ProfilePoints[2][1].Y, ProfilePoints[2][1].Z + 1000);
            var bottomFace = new RebarLegFace();
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][1], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(offsetedEndPoint, null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(offsetedStartPoint, null));
            rebarLegFaces[1] = bottomFace;
            bool succes = rebarSet.Modify();

            ModelObjectEnumerator modelObjectEnumerator = rebarSet.GetRebarModifiers();
            List<ModelObject> modelObjectList = Utility.ToList(modelObjectEnumerator);
            RebarPropertyModifier rebarPropertyModifier = (from mo in modelObjectList
                                                           where mo.GetType() == typeof(RebarPropertyModifier)
                                                           select mo as RebarPropertyModifier).FirstOrDefault();
            if (rebarPropertyModifier != null)
            {
                var contour = new Contour();
                contour.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
                contour.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
                contour.AddContourPoint(new ContourPoint(correctedP12, null));
                contour.AddContourPoint(new ContourPoint(correctedP32, null));
                rebarPropertyModifier.Curve = contour;
                rebarPropertyModifier.Modify();
            }

            new Model().CommitChanges();
        }
        void RTWSEditBottomInnerVerticalRebar(RTWS rtw, RebarSet rebarSet)
        {
            List<RebarLegFace> rebarLegFaces = rebarSet.LegFaces;

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][2], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            rebarLegFaces[0] = mainFace;

            Point offsetedStartPoint = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y, ProfilePoints[0][0].Z - 1000);
            Point offsetedEndPoint = new Point(ProfilePoints[2][0].X, ProfilePoints[2][0].Y, ProfilePoints[2][0].Z - 1000);
            var bottomFace = new RebarLegFace();
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][0], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(offsetedEndPoint, null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(offsetedStartPoint, null));
            rebarLegFaces[1] = bottomFace;
        
            bool succes = rebarSet.Modify();
            new Model().CommitChanges();
        }
        void RTWSEditTopInnerVerticalRebar(RTWS rtw, RebarSet rebarSet)
        {
            List<RebarLegFace> rebarLegFaces = rebarSet.LegFaces;

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][3], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][4], null));            
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));            
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));            
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][4], null));            
            rebarLegFaces[0] = mainFace;
            bool succes = rebarSet.Modify();

            ModelObjectEnumerator modelObjectEnumerator = rebarSet.GetRebarModifiers();
            List<ModelObject> modelObjectList = Utility.ToList(modelObjectEnumerator);
            RebarEndDetailModifier rebarEndDetail = (from mo in modelObjectList
                                                                      where mo.GetType() == typeof(RebarEndDetailModifier)
                                                                      select mo as RebarEndDetailModifier).FirstOrDefault();

            if (rebarEndDetail != null)
            {
                var contour = new Contour();
                contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
                contour.AddContourPoint(new ContourPoint(new Point(ProfilePoints[2][3].X-100,ProfilePoints[2][3].Y,ProfilePoints[2][3].Z), null));
                rebarEndDetail.Curve = contour;
                rebarEndDetail.Modify();
            }

            new Model().CommitChanges();
        }
        void RTWSEditSkewVerticalRebar(RTWS rtw, RebarSet rebarSet)
        {
            List<RebarLegFace> rebarLegFaces = rebarSet.LegFaces;

            Line startLine = new Line(ProfilePoints[0][3], ProfilePoints[0][2]);
            Line endLine = new Line(ProfilePoints[2][3], ProfilePoints[2][2]);

            GeometricPlane geometricPlane = new GeometricPlane(ProfilePoints[0][1], new Vector(1, 0, 0), new Vector(0, 1, 0));

            Point startIntersection = Utility.GetExtendedIntersection(startLine, geometricPlane, 10);
            Point endIntersection = Utility.GetExtendedIntersection(endLine, geometricPlane, 10);

            Point cP06 = new Point(ProfilePoints[0][2].X, ProfilePoints[0][2].Y - 1000, ProfilePoints[0][2].Z);
            Point cP05 = new Point(startIntersection.X, startIntersection.Y + 1000, startIntersection.Z);
            Point cP36 = new Point(ProfilePoints[2][2].X, ProfilePoints[2][2].Y - 1000, ProfilePoints[2][2].Z);
            Point cP35 = new Point(endIntersection.X, endIntersection.Y + 1000, endIntersection.Z);

            var topFace = new RebarLegFace();
            topFace.Contour.AddContourPoint(new ContourPoint(cP05, null));
            topFace.Contour.AddContourPoint(new ContourPoint(startIntersection, null));
            topFace.Contour.AddContourPoint(new ContourPoint(endIntersection, null));
            topFace.Contour.AddContourPoint(new ContourPoint(cP35, null));
            rebarLegFaces[0] = topFace;

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(startIntersection, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(endIntersection, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][2], null));
            rebarLegFaces[1] = mainFace;

            var bottomFace = new RebarLegFace();
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(cP06, null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(cP36, null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][2], null));
            rebarLegFaces[2] = bottomFace;
            bool succes = rebarSet.Modify();

            RebarGuideline rebarGuideline = rebarSet.Guidelines.FirstOrDefault();
            var contour = new Contour();
            contour.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            contour.AddContourPoint(new ContourPoint(ProfilePoints[2][2], null));
            rebarGuideline.Curve = contour;

            ModelObjectEnumerator modelObjectEnumerator = rebarSet.GetRebarModifiers();
            List<ModelObject> modelObjectList = Utility.ToList(modelObjectEnumerator);
            List<RebarEndDetailModifier> rebarPropertyModifierList = (from mo in modelObjectList
                                                                 where mo.GetType() == typeof(RebarEndDetailModifier)
                                                                 select mo as RebarEndDetailModifier).ToList();

            if (rebarPropertyModifierList!=null)
            {
                RebarEndDetailModifier rebarPropertyModifier1 = rebarPropertyModifierList[0];
                var contour2 = new Contour();
                contour2.AddContourPoint(new ContourPoint(cP05, null));
                contour2.AddContourPoint(new ContourPoint(cP35, null));
                rebarPropertyModifier1.Curve = contour2;
                rebarPropertyModifier1.Modify();

                RebarEndDetailModifier rebarPropertyModifier2 = rebarPropertyModifierList[1];
                var contour3 = new Contour();
                contour3.AddContourPoint(new ContourPoint(cP06, null));
                contour3.AddContourPoint(new ContourPoint(cP36, null));
                rebarPropertyModifier2.Curve = contour3;
                rebarPropertyModifier2.Modify();
            }

            new Model().CommitChanges();
        }
        void RTWSEditBottomInnerLongitudinalRebar(RTWS rtws, RebarSet rebarSet)
        {
            List<RebarLegFace> rebarLegFaces = rebarSet.LegFaces;

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][2], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][0], null));
            rebarLegFaces[0] = mainFace;

            bool succes = rebarSet.Modify();
            new Model().CommitChanges();
        }
        void RTWSEditOuterLongitudinalRebar(RTWS rtws, RebarSet rebarSet)
        {
            List<RebarLegFace> rebarLegFaces = rebarSet.LegFaces;

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][1], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][4], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][1], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][2], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][5], null));
            rebarLegFaces[0] = mainFace;

            bool succes = rebarSet.Modify();
            new Model().CommitChanges();
        }
        void RTWSEditTopInnerLongitudinalRebar(RTWS rTWS,RebarSet rebarSet)
        {
            List<RebarLegFace> rebarLegFaces = rebarSet.LegFaces;

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][3], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][4], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][4], null));
            rebarLegFaces[0] = mainFace;

            bool succes = rebarSet.Modify();
            new Model().CommitChanges();
        }
        void RTWSEditSkewLongitudinalRebar(RTWS rtws,RebarSet rebarSet)
        {
            List<RebarLegFace> rebarLegFaces = rebarSet.LegFaces;

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][2], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][3], null));
            rebarLegFaces[0] = mainFace;

            bool succes = rebarSet.Modify();
            new Model().CommitChanges();
        }
        void RTWSEditCornicePerpendicularRebar(RTWS rtws, RebarSet rebarSet)
        {
            RebarGuideline rebarGuideline = rebarSet.Guidelines.FirstOrDefault();
            var contour = new Contour();
            contour.AddContourPoint(new ContourPoint(ProfilePoints[1][5], null));
            contour.AddContourPoint(new ContourPoint(ProfilePoints[3][5], null));
            rebarGuideline.Curve = contour;

            bool succes = rebarSet.Modify();
            new Model().CommitChanges();
        }
        void RTWSEditCorniceLongitudinalRebar(RTWS rtws, RebarSet rebarSet)
        {
            List<RebarLegFace> rebarLegFaces = rebarSet.LegFaces;

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][5], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][3], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][3], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][5], null));
            rebarLegFaces[0] = mainFace;

            bool succes = rebarSet.Modify();
            new Model().CommitChanges();
        }
        void RTWSEditBottomClosingCShapeRebar(RTWS rtws, RebarSet rebarSet, int number)
        {
            if (number == 0)
            {
                rebarSet.Delete();
                return;
            }

            List<RebarLegFace> rebarLegFaces = rebarSet.LegFaces;

            Point cP36 = new Point(ProfilePoints[2][1].X, ProfilePoints[2][2].Y, ProfilePoints[2][1].Z);
            Point cP06 = new Point(ProfilePoints[0][1].X, ProfilePoints[0][2].Y, ProfilePoints[0][1].Z);

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][1], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(cP36, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][2], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][0], null));
            rebarLegFaces[0] = mainFace;

            var innerFace =new RebarLegFace();
            innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][0], null));
            innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][2], null));
            rebarLegFaces[1] = innerFace;

            var outerFace = new RebarLegFace();
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][1], null));
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
            outerFace.Contour.AddContourPoint(new ContourPoint(cP06, null));
            outerFace.Contour.AddContourPoint(new ContourPoint(cP36, null));
            rebarLegFaces[2] = outerFace;

            bool succes = rebarSet.Modify();
            new Model().CommitChanges();

        }
        void RTWSEditSkewClosingCShapeRebar(RTWS rtws, RebarSet rebarSet, int number)
        {
            if (number == 0)
            {
                rebarSet.Delete();
                return;
            }

            List<RebarLegFace> rebarLegFaces = rebarSet.LegFaces;

            Point cP22 = new Point(ProfilePoints[2][1].X, ProfilePoints[2][2].Y, ProfilePoints[2][1].Z);
            Point cP02 = new Point(ProfilePoints[0][1].X, ProfilePoints[0][2].Y, ProfilePoints[0][1].Z);
            Point cP35 = new Point(ProfilePoints[2][4].X, ProfilePoints[2][4].Y, ProfilePoints[2][4].Z);
            Point cP05 = new Point(ProfilePoints[0][1].X, ProfilePoints[0][3].Y, ProfilePoints[0][1].Z);

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(cP22, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][4], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][3], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][2], null));
            rebarLegFaces[0] = mainFace;

            var innerFace = new RebarLegFace();
            innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][2], null));
            innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][3], null));
            innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            rebarLegFaces[1] = innerFace;

            var outerFace = new RebarLegFace();
            outerFace.Contour.AddContourPoint(new ContourPoint(cP35, null));
            outerFace.Contour.AddContourPoint(new ContourPoint(cP22, null));
            outerFace.Contour.AddContourPoint(new ContourPoint(cP02, null));
            outerFace.Contour.AddContourPoint(new ContourPoint(cP05, null));
            rebarLegFaces[2] = outerFace;

            bool succes = rebarSet.Modify();

            new Model().CommitChanges();
        }
        void RTWSEditTopClosingCShapeRebar(RTWS rtws, RebarSet rebarSet,int number)
        {
            if (number == 0)
            {
                rebarSet.Delete();
                return;
            }

            List<RebarLegFace> cRebarLegFaces = new List<RebarLegFace>();

            Point cP51 = new Point(ProfilePoints[3][2].X, ProfilePoints[3][2].Y + rtws.CorniceHeight, ProfilePoints[3][2].Z);
            Point cP21 = new Point(ProfilePoints[1][2].X, ProfilePoints[1][2].Y + rtws.CorniceHeight, ProfilePoints[1][2].Z);
            Point cP03 = new Point(ProfilePoints[0][1].X, ProfilePoints[0][3].Y, ProfilePoints[0][1].Z);

            var mainFaceBottom = new RebarLegFace();
            mainFaceBottom.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][4], null));
            mainFaceBottom.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][1], null));
            mainFaceBottom.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][0], null));
            mainFaceBottom.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][3], null));
            cRebarLegFaces.Add(mainFaceBottom);

            var mainFaceTop = new RebarLegFace();
            mainFaceTop.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][0], null));
            mainFaceTop.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][1], null));
            mainFaceTop.Contour.AddContourPoint(new ContourPoint(cP51, null));
            mainFaceTop.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][4], null));
            cRebarLegFaces.Add(mainFaceTop);

            var innerFace = new RebarLegFace();
            innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][4], null));
            innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][0], null));
            innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][3], null));
            innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][4], null));
            innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][0], null));
            innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][4], null));
            cRebarLegFaces.Add(innerFace);

            var outerFace = new RebarLegFace();
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][1], null));
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][4], null));
            outerFace.Contour.AddContourPoint(new ContourPoint(cP03, null));
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][5], null));
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
            outerFace.Contour.AddContourPoint(new ContourPoint(cP21, null));
            outerFace.Contour.AddContourPoint(new ContourPoint(cP51, null));
            cRebarLegFaces.Add(outerFace);

            rebarSet.LegFaces = cRebarLegFaces;

            bool succes = rebarSet.Modify();
            new Model().CommitChanges();

            LayerDictionary[rebarSet.Identifier.ID] = new int[] { 1, 1, 2, 2 };
        }
        void RTWSEditClosingLongitudinalRebar(RTWS rtws,RebarSet rebarSet,int number)
        {
            if (number == 0)
            {
                rebarSet.Delete();
                return;
            }

            List<RebarLegFace> rebarLegFaces = rebarSet.LegFaces;
            Point cP35 = new Point(ProfilePoints[3][0].X, ProfilePoints[3][5].Y, ProfilePoints[3][0].Z);

            var mainFaceBottom = new RebarLegFace();
            mainFaceBottom.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][0], null));
            mainFaceBottom.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][7], null));
            mainFaceBottom.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][6], null));
            mainFaceBottom.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][5], null));
            mainFaceBottom.Contour.AddContourPoint(new ContourPoint(cP35, null));
            rebarLegFaces[0]=mainFaceBottom;

            bool succes = rebarSet.Modify();

            var bottomLengthModifier = new RebarEndDetailModifier();
            bottomLengthModifier.Father = rebarSet;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.END_OFFSET;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentLength = 40 * Convert.ToDouble(rebarSet.RebarProperties.Size);
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(new Point(ProfilePoints[3][5].X,ProfilePoints[3][5].Y,ProfilePoints[3][5].Z+50), null));
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(cP35, null));
            bool inserted = bottomLengthModifier.Insert();

            new Model().CommitChanges();

        }
        void RTWSEditBottomCShapeRebar(RTWS rtws,RebarSet rebarSet)
        {
            GeometricPlane plane = new GeometricPlane(ProfilePoints[3][0], new Vector(0, 1, 0), new Vector(0, 0, 1));

            RebarGuideline gl = rebarSet.Guidelines.FirstOrDefault();
            System.Collections.ArrayList arrayList = gl.Curve.ContourPoints;
            var startPoint = arrayList[0] as ContourPoint;
            var endPoint = arrayList[1] as ContourPoint;
            Line glLine = new Line(startPoint, endPoint);

            Point correctedEndPoint = Utility.GetExtendedIntersection(glLine, plane, 2);

            gl.Curve.ContourPoints[1] = new ContourPoint(correctedEndPoint, null);
            bool succes = rebarSet.Modify();
            new Model().CommitChanges();
        }
        void RTWSEditTopCShapeRebar(RTWS rtws, RebarSet rebarSet)
        {
            Vector longitudinal = Utility.GetVectorFromTwoPoints(ProfilePoints[4][0], ProfilePoints[5][0]);
            Vector perpendicular = Utility.GetVectorFromTwoPoints(ProfilePoints[4][0], ProfilePoints[4][5]);
            GeometricPlane skewPlane = new GeometricPlane(ProfilePoints[4][0], longitudinal,perpendicular);

            GeometricPlane verticalPlane = new GeometricPlane(ProfilePoints[1][0], new Vector(0, 1, 0), new Vector(0, 0, 1));

            RebarGuideline gl = rebarSet.Guidelines.FirstOrDefault();
            System.Collections.ArrayList arrayList = gl.Curve.ContourPoints;
            var startPoint = arrayList[0] as ContourPoint;
            var endPoint = arrayList[1] as ContourPoint;
            Line glLine = new Line(startPoint, endPoint);

            if(startPoint.Y<=ProfilePoints[5][0].Y)
            {
                Point endIntersectionPoint = Utility.GetExtendedIntersection(glLine, skewPlane, 2);
                gl.Curve.ContourPoints[1] = new ContourPoint(endIntersectionPoint, null);
                gl.Spacing.EndOffset = 300;
            }

            if(startPoint.Y>=ProfilePoints[1][1].Y-50)
            {
                Point startIntersectionPoint = Utility.GetExtendedIntersection(glLine, verticalPlane, 2);
                gl.Curve.ContourPoints[0] = new ContourPoint(startIntersectionPoint, null);
            }

            bool succes = rebarSet.Modify();
            new Model().CommitChanges();
        }
        #endregion
        #region Properties
        public double SCD { get { return _scd; } }
        public double SOCH { get { return _soch; } }
        public double SICH { get { return _sich; } }
        public double CD { get { return _cd; } }
        public double CtBH { get { return _ctbh; } }
        public double CtTH { get { return _ctth; } }
        #endregion
        #region Fields
        static ProfileType _profileType;
        enum ProfileType
        {
            RTW,
            RTWS
        }      
        static double _soch, _scd, _sich, _ctth, _ctbh, _cd;
        bool _isReversed;
        #endregion
    }
}
