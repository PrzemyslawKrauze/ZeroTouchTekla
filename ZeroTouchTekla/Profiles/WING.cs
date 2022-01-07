using System;
using System.Collections.Generic;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model;
using System.Text.RegularExpressions;
using System.Linq;
using System.Reflection;

namespace ZeroTouchTekla.Profiles
{
    public class WING : Element
    {
        public WING(Beam part) : base(part)
        {
            _beam = part;
            GetProfilePointsAndParameters(part);
        }
        #region PublicMethods
        public static void GetProfilePointsAndParameters(Beam beam)
        {
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
            int i = 0;

        }
        new public void Create()
        {
            Model model = new Model();
            string componentName = _beam.GetFatherComponent().Name;
            Element element;
            if (_profileType == ProfileType.RTW)
            {
                RTW rtw = new RTW(_beam);
                rtw.Create();
                element = rtw;
            }
            else
            {
                RTWS rtws = new RTWS(_beam);
                rtws.Create();
                element = rtws;
            }

            Type[] Types = new Type[] { typeof(RebarSet) };
            ModelObjectEnumerator moe = model.GetModelObjectSelector().GetAllObjectsWithType(Types);
            var rebarList = Utility.ToList(moe);

            List<RebarSet> selectedRebars = (from RebarSet r in rebarList
                                             where Utility.GetUserProperty(r, RebarCreator.FatherIDName) == _beam.Identifier.ID
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
                RTWS rtws = new RTWS(_beam);
                rtws.Create();
                element = rtws;
            }
        }
        #region RTWMethods
        void EditRTWRebar(RTW rtw, List<RebarSet> rebarSets)
        {
            foreach (RebarSet rs in rebarSets)
            {
                string methodName = string.Empty;
                rs.GetUserProperty(RebarCreator.MethodName, ref methodName);
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
            RTW rtw = element as RTW;
            double correctedHeight1 = rtw.Height - (rtw.Height - rtw.Height2) * (SCD) / rtw.Length;
            double correctedHeight2 = rtw.Height - (rtw.Height - rtw.Height2) * (rtw.Length - CD) / rtw.Length;
            double s = (rtw.BottomWidth - (rtw.TopWidth - rtw.CorniceWidth)) / rtw.Height;

            List<List<Point>> profilePoints = rtw.GetProfilePoints();
            List<List<Point>> correctedPoints = new List<List<Point>>();

            Point p00 = profilePoints[0][0];
            Point p01 = new Point(p00.X, p00.Y + correctedHeight1 - SOCH, p00.Z);
            Point p03 = profilePoints[0][5];
            Point p02 = new Point(p00.X, p01.Y, p01.Z - rtw.BottomWidth + s *  (correctedHeight1 - SOCH));
            correctedPoints.Add(new List<Point> { p00, p01, p02, p03 });

            Point p10 = new Point(p00.X + SCD, p00.Y, p00.Z);
            Point p11 = new Point(p10.X, p10.Y + correctedHeight1 - SICH, p10.Z);
            Point p12 = new Point(p10.X, p11.Y, profilePoints[0][4].Z - s * SICH);
            Point p13 = new Point(p10.X, p03.Y, profilePoints[0][4].Z - s * correctedHeight1);
            correctedPoints.Add(new List<Point> { p10, p11, p12, p13 });

            Point p20 = p10;
            Point p21 = new Point(p20.X, p20.Y + correctedHeight1 - rtw.CorniceHeight, p20.Z);
            Point p22 = new Point(p21.X, p21.Y, p21.Z + rtw.CorniceWidth);
            Point p23 = new Point(p22.X, p22.Y + rtw.CorniceHeight, p22.Z);
            Point p24 = new Point(p23.X, p23.Y, p23.Z - rtw.TopWidth);
            Point p25 = new Point(p20.X, p13.Y, p24.Z - s * correctedHeight1);
            correctedPoints.Add(new List<Point> { p20, p21, p22, p23, p24, p25 });

            Point p30 = new Point(profilePoints[1][0].X - CD, p20.Y, p20.Z);
            Point p31 = new Point(p30.X, p30.Y + correctedHeight2 - rtw.CorniceHeight, p30.Z);
            Point p32 = new Point(p30.X, p31.Y, p31.Z + rtw.CorniceWidth);
            Point p33 = new Point(p32.X, p32.Y + rtw.CorniceHeight, p32.Z);
            Point p34 = new Point(p33.X, p33.Y, p33.Z - rtw.TopWidth);
            Point p35 = new Point(p30.X, p25.Y, p34.Z - s * correctedHeight2);
            correctedPoints.Add(new List<Point> { p30, p31, p32, p33, p34, p35 });

            Point p40 = new Point(p30.X, p30.Y + CtBH, p30.Z);
            Point p41 = new Point(p40.X, p30.Y + correctedHeight2 - rtw.CorniceHeight, p40.Z);
            Point p42 = new Point(p41.X, p41.Y, p41.Z + rtw.CorniceWidth);
            Point p43 = new Point(p42.X, p42.Y + rtw.CorniceHeight, p42.Z);
            Point p44 = new Point(p43.X, p43.Y, p43.Z - rtw.TopWidth);
            Point p45 = new Point(p30.X, p40.Y, p44.Z - s * (correctedHeight2 - CtBH));
            correctedPoints.Add(new List<Point> { p40, p41, p42, p43, p44, p45 });

            Point p50 = new Point(profilePoints[1][0].X, p30.Y + rtw.Height2 - CtTH, p30.Z);
            Point p51 = new Point(p50.X, p30.Y + rtw.Height2 - rtw.CorniceHeight, p50.Z);
            Point p52 = new Point(p51.X, p51.Y, p51.Z + rtw.CorniceWidth);
            Point p53 = new Point(p52.X, p52.Y + rtw.CorniceHeight, p52.Z);
            Point p54 = new Point(p53.X, p53.Y, p53.Z - rtw.TopWidth);
            Point p55 = new Point(p50.X, p50.Y, p54.Z - s * CtTH);
            correctedPoints.Add(new List<Point> { p50, p51, p52, p53, p54, p55 });

            ProfilePoints = correctedPoints;
        }
        void SetRTWProfilePointsReversed(Element element)
        {
            RTW rtw = element as RTW;
            double correctedHeight1 = rtw.Height2 - (rtw.Height2 - rtw.Height) * (SCD) / rtw.Length;
            double correctedHeight2 = rtw.Height2 - (rtw.Height2 - rtw.Height) * (rtw.Length - CD) / rtw.Length;
            double s = (rtw.BottomWidth - (rtw.TopWidth - rtw.CorniceWidth)) / rtw.Height;
            double secondBottomWidth = rtw.TopWidth - rtw.CorniceWidth + rtw.Height2 * s;

            List<List<Point>> profilePoints = rtw.GetProfilePoints();
            List<List<Point>> correctedPoints = new List<List<Point>>();

            Point p00 = profilePoints[1][0];
            Point p01 = new Point(p00.X, p00.Y + correctedHeight1 - SOCH, p00.Z);
            Point p03 = profilePoints[1][5];
            Point p02 = new Point(p00.X, p01.Y, p01.Z - secondBottomWidth + s * (correctedHeight1 - SOCH));
            correctedPoints.Add(new List<Point> { p00, p01, p02, p03 });

            Point p10 = new Point(p00.X - SCD, p00.Y, p00.Z);
            Point p11 = new Point(p10.X, p10.Y + correctedHeight1 - SICH, p10.Z);
            Point p12 = new Point(p10.X, p11.Y, profilePoints[1][4].Z - s * SICH);
            Point p13 = new Point(p10.X, p03.Y, profilePoints[1][4].Z - s * correctedHeight1);
            correctedPoints.Add(new List<Point> { p10, p11, p12, p13 });

            Point p20 = p10;
            Point p21 = new Point(p20.X, p20.Y + correctedHeight1 - rtw.CorniceHeight, p20.Z);
            Point p22 = new Point(p21.X, p21.Y, p21.Z + rtw.CorniceWidth);
            Point p23 = new Point(p22.X, p22.Y + rtw.CorniceHeight, p22.Z);
            Point p24 = new Point(p23.X, p23.Y, p23.Z - rtw.TopWidth);
            Point p25 = new Point(p20.X, p13.Y, p24.Z - s * correctedHeight1);
            correctedPoints.Add(new List<Point> { p20, p21, p22, p23, p24, p25 });

            Point p30 = new Point(profilePoints[0][0].X + CD, p20.Y, p20.Z);
            Point p31 = new Point(p30.X, p30.Y + correctedHeight2 - rtw.CorniceHeight, p30.Z);
            Point p32 = new Point(p30.X, p31.Y, p31.Z + rtw.CorniceWidth);
            Point p33 = new Point(p32.X, p32.Y + rtw.CorniceHeight, p32.Z);
            Point p34 = new Point(p33.X, p33.Y, p33.Z - rtw.TopWidth);
            Point p35 = new Point(p30.X, p25.Y, p34.Z - s * correctedHeight2);
            correctedPoints.Add(new List<Point> { p30, p31, p32, p33, p34, p35 });

            Point p40 = new Point(p30.X, p30.Y + CtBH, p30.Z);
            Point p41 = new Point(p40.X, p30.Y + correctedHeight2 - rtw.CorniceHeight, p40.Z);
            Point p42 = new Point(p41.X, p41.Y, p41.Z + rtw.CorniceWidth);
            Point p43 = new Point(p42.X, p42.Y + rtw.CorniceHeight, p42.Z);
            Point p44 = new Point(p43.X, p43.Y, p43.Z - rtw.TopWidth);
            Point p45 = new Point(p30.X, p40.Y, p44.Z - s * (correctedHeight2 - CtBH));
            correctedPoints.Add(new List<Point> { p40, p41, p42, p43, p44, p45 });

            Point p50 = new Point(profilePoints[0][0].X, p30.Y + rtw.Height - CtTH, p30.Z);
            Point p51 = new Point(p50.X, p30.Y + rtw.Height - rtw.CorniceHeight, p50.Z);
            Point p52 = new Point(p51.X, p51.Y, p51.Z + rtw.CorniceWidth);
            Point p53 = new Point(p52.X, p52.Y + rtw.CorniceHeight, p52.Z);
            Point p54 = new Point(p53.X, p53.Y, p53.Z - rtw.TopWidth);
            Point p55 = new Point(p50.X, p50.Y, p54.Z - s * CtTH);
            correctedPoints.Add(new List<Point> { p50, p51, p52, p53, p54, p55 });

            ProfilePoints = correctedPoints;
        }
        void RTWEditOuterVerticalRebar(RTW rtw, RebarSet rebarSet)
        {
            List<RebarLegFace> rebarLegFaces = rebarSet.LegFaces;

            Point correctedP51 = new Point(ProfilePoints[5][1].X, ProfilePoints[5][1].Y + rtw.CorniceHeight, ProfilePoints[5][1].Z);
            Point correctedP21 = new Point(ProfilePoints[2][1].X, ProfilePoints[2][1].Y + rtw.CorniceHeight, ProfilePoints[2][1].Z);

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[4][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[5][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(correctedP51, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(correctedP21, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
            rebarLegFaces[0] = mainFace;

            Point offsetedStartPoint = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y, ProfilePoints[0][0].Z + 1000);
            Point offsetedEndPoint = new Point(ProfilePoints[3][0].X, ProfilePoints[3][0].Y, ProfilePoints[3][0].Z + 1000);
            var bottomFace = new RebarLegFace();
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][0], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(offsetedEndPoint, null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(offsetedStartPoint, null));
            rebarLegFaces[1] = bottomFace;
            bool succes = rebarSet.Modify();

            new Model().CommitChanges();
        }
        void RTWEditInnerVerticalRebar(RTW rtw, RebarSet rebarSet)
        {
            List<RebarLegFace> rebarLegFaces = rebarSet.LegFaces;

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][5], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[4][5], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[5][5], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[5][4], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][4], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            rebarLegFaces[0] = mainFace;

            Point offsetedStartPoint = new Point(ProfilePoints[0][3].X, ProfilePoints[0][3].Y, ProfilePoints[0][3].Z - 1000);
            Point offsetedEndPoint = new Point(ProfilePoints[3][5].X, ProfilePoints[3][5].Y, ProfilePoints[3][5].Z - 1000);
            var bottomFace = new RebarLegFace();
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][5], null));
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
                contour.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
                contour.AddContourPoint(new ContourPoint(ProfilePoints[2][4], null));
                contour.AddContourPoint(new ContourPoint(ProfilePoints[5][4], null));
                rebarPropertyModifier.Curve = contour;
                rebarPropertyModifier.Modify();
            }




            bool succes = rebarSet.Modify();
            new Model().CommitChanges();
        }
        void RTWEditOuterLongitudinalRebar(RTW rtw, RebarSet rebarSet)
        {
            List<RebarLegFace> rebarLegFaces = rebarSet.LegFaces;

            Point correctedP51 = new Point(ProfilePoints[5][1].X, ProfilePoints[5][1].Y + rtw.CorniceHeight, ProfilePoints[5][1].Z);
            Point correctedP21 = new Point(ProfilePoints[2][1].X, ProfilePoints[2][1].Y + rtw.CorniceHeight, ProfilePoints[2][1].Z);

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[4][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[5][0], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(correctedP51, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(correctedP21, null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
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
                Point secondPoint = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y + startOffset + firstLength, ProfilePoints[0][0].Z);

                var contour = new Contour();
                contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
                contour.AddContourPoint(new ContourPoint(secondPoint, null));
                rebarPropertyModifier.Curve = contour;
                rebarPropertyModifier.Modify();
            }

            var bottomFace = new RebarLegFace();
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
            rebarLegFaces.Add(bottomFace);

            var midFace = new RebarLegFace();
            midFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
            midFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            midFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
            midFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
            rebarLegFaces.Add(midFace);

            Point corrected21Point = new Point(ProfilePoints[2][1].X, ProfilePoints[2][1].Y + rtw.CorniceHeight, ProfilePoints[2][1].Z);
            var topFace = new RebarLegFace();
            topFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
            topFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
            topFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][4], null));
            topFace.Contour.AddContourPoint(new ContourPoint(corrected21Point, null));
            rebarLegFaces.Add(topFace);

            var innerEndDetailModifier = new RebarEndDetailModifier();
            innerEndDetailModifier.Father = rebarSet;
            innerEndDetailModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.LEG_LENGTH;
            innerEndDetailModifier.RebarLengthAdjustment.AdjustmentLength = 300;
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
            innerEndDetailModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[2][4], null));
            innerEndDetailModifier.Insert();

            bool succes = rebarSet.Modify();
            new Model().CommitChanges();
            RebarCreator.LayerDictionary[rebarSet.Identifier.ID] = new int[] { 2, 2, 2, 2 };
        }
        void RTWEditInnerLongitudinalRebar(RTW rtw, RebarSet rebarSet)
        {
            List<RebarLegFace> rebarLegFaces = rebarSet.LegFaces;

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][5], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[4][5], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[5][5], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[5][4], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][4], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
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
                    Point origin = new Point(ProfilePoints[0][0].X, ProfilePoints[0][0].Y + startOffset + firstLength, ProfilePoints[0][0].Z);
                    GeometricPlane plane = new GeometricPlane(origin, new Vector(0, 1, 0));
                    Line line = new Line(ProfilePoints[0][3], ProfilePoints[0][2]);
                    Point intersection = Utility.GetExtendedIntersection(line, plane, 1);

                    var contour = new Contour();
                    contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
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
            contour.AddContourPoint(new ContourPoint(ProfilePoints[2][3], null));
            contour.AddContourPoint(new ContourPoint(ProfilePoints[5][3], null));
            rebarGuideline.Curve = contour;

            bool succes = rebarSet.Modify();
            new Model().CommitChanges();
        }
        void RTWEditCorniceLongitudinalRebar(RTW rtw, RebarSet rebarSet)
        {
            List<RebarLegFace> rebarLegFaces = rebarSet.LegFaces;

            var mainFace = new RebarLegFace();
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[5][3], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][3], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][2], null));
            mainFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[5][2], null));
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

            Point correctedP51 = new Point(ProfilePoints[5][1].X, ProfilePoints[5][1].Y + rtw.CorniceHeight, ProfilePoints[5][1].Z);
            Point correctedP21 = new Point(ProfilePoints[2][1].X, ProfilePoints[2][1].Y + rtw.CorniceHeight, ProfilePoints[2][1].Z);

            List<RebarLegFace> correctedLegFaces = new List<RebarLegFace>();
            var bottomFace = new RebarLegFace();
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][0], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[4][0], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[4][5], null));
            bottomFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][5], null));
            correctedLegFaces.Add(bottomFace);

            var midFace = new RebarLegFace();
            midFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[4][0], null));
            midFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[4][5], null));
            midFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[5][5], null));
            midFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[5][0], null));
            correctedLegFaces.Add(midFace);

            var topFace = new RebarLegFace();
            topFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[5][0], null));
            topFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[5][5], null));
            topFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[5][4], null));
            topFace.Contour.AddContourPoint(new ContourPoint(correctedP51, null));
            correctedLegFaces.Add(topFace);

            var innerFace = new RebarLegFace();
            innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][0], null));
            innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][0], null));
            innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[4][0], null));
            innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[5][0], null));
            innerFace.Contour.AddContourPoint(new ContourPoint(correctedP51, null));
            innerFace.Contour.AddContourPoint(new ContourPoint(correctedP21, null));
            innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][1], null));
            innerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][1], null));
            correctedLegFaces.Add(innerFace);

            var outerFace = new RebarLegFace();
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][3], null));
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[3][5], null));
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[4][5], null));
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[5][5], null));
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[5][4], null));
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[2][4], null));
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[1][2], null));
            outerFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[0][2], null));
            correctedLegFaces.Add(outerFace);

            rebarSet.LegFaces = correctedLegFaces;

            bool succes = rebarSet.Modify();
            new Model().CommitChanges();

            RebarCreator.LayerDictionary[rebarSet.Identifier.ID] = new int[] { 1, 1, 1, 2, 2 };
        }
        void RTWEditClosingLongitudinalRebar(RTW rtw, RebarSet rebarSet, int number)
        {
            if (number == 1)
            {
                rebarSet.Delete();
                return;
            }

            Point correctedP51 = new Point(ProfilePoints[5][1].X, ProfilePoints[5][1].Y + rtw.CorniceHeight, ProfilePoints[5][1].Z);

            List<RebarLegFace> correctedLegFaces = new List<RebarLegFace>();

            var midFace = new RebarLegFace();
            midFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[4][0], null));
            midFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[4][5], null));
            midFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[5][5], null));
            midFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[5][0], null));
            correctedLegFaces.Add(midFace);

            var topFace = new RebarLegFace();
            topFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[5][0], null));
            topFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[5][5], null));
            topFace.Contour.AddContourPoint(new ContourPoint(ProfilePoints[5][4], null));
            topFace.Contour.AddContourPoint(new ContourPoint(correctedP51, null));
            correctedLegFaces.Add(topFace);

            rebarSet.LegFaces = correctedLegFaces;

            var bottomLengthModifier = new RebarEndDetailModifier();
            bottomLengthModifier.Father = rebarSet;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentType = RebarLengthAdjustmentDataNullable.LengthAdjustmentTypeEnum.END_OFFSET;
            bottomLengthModifier.RebarLengthAdjustment.AdjustmentLength = 40 * Convert.ToDouble(rebarSet.RebarProperties.Size);
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[4][0], null));
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[4][5], null));
            bottomLengthModifier.Insert();

            bool succes = rebarSet.Modify();
            new Model().CommitChanges();

            RebarCreator.LayerDictionary[rebarSet.Identifier.ID] = new int[] { 2, 2 };
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
            rebarSet.RebarProperties.Class = SetClass(Convert.ToDouble(rebarSize));
            rebarSet.RebarProperties.Size = rebarSize;
            rebarSet.RebarProperties.BendingRadius = GetBendingRadious(Convert.ToDouble(rebarSize));
            rebarSet.LayerOrderNumber = 1;

            Point leftBottom, rightBottom, rightTop, leftTop;

            leftBottom = ProfilePoints[3][0];
            rightBottom = ProfilePoints[3][5];
            rightTop = ProfilePoints[4][5];
            leftTop = ProfilePoints[4][0];


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
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[4][0], null));
            bottomLengthModifier.Curve.AddContourPoint(new ContourPoint(ProfilePoints[4][5], null));
            bool inserted = bottomLengthModifier.Insert();


            new Model().CommitChanges();

            PostRebarCreationMethod(rebarSet, MethodBase.GetCurrentMethod(), number);
            RebarCreator.LayerDictionary.Add(rebarSet.Identifier.ID, new int[] { 2 });
        }
        void RTWEditCShapeRebar(RTW rtw, RebarSet rebarSet)
        {
            RebarGuideline gl = rebarSet.Guidelines.FirstOrDefault();
            System.Collections.ArrayList arrayList = gl.Curve.ContourPoints;
            var startPoint = arrayList[0] as ContourPoint;
            var endPoint = arrayList[1] as ContourPoint;
            GeometricPlane plane;

            if (startPoint.Y <= ProfilePoints[4][0].Y)
            {
                Vector xAxis = Utility.GetVectorFromTwoPoints(ProfilePoints[3][0], ProfilePoints[3][1]);
                Vector yAxis = Utility.GetVectorFromTwoPoints(ProfilePoints[3][0], ProfilePoints[3][5]);
                plane = new GeometricPlane(ProfilePoints[3][0], xAxis, yAxis);
            }
            else
            {
                Vector xAxis = Utility.GetVectorFromTwoPoints(ProfilePoints[4][0], ProfilePoints[5][0]);
                Vector yAxis = Utility.GetVectorFromTwoPoints(ProfilePoints[4][0], ProfilePoints[4][5]);
                plane = new GeometricPlane(ProfilePoints[4][0], xAxis, yAxis);
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
                GeometricPlane secondPlane = new GeometricPlane(ProfilePoints[1][2], xAxis2, yAxis2);
                Point correctedEndPoint = Utility.GetExtendedIntersection(glLine, secondPlane, 2);
                int secondNumber = _isReversed ? 1 : 0;
                gl.Curve.ContourPoints[secondNumber] = new ContourPoint(correctedEndPoint, null);
            }

            bool succes = rebarSet.Modify();
            new Model().CommitChanges();
        }
        #endregion
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
        static Beam _beam;
        static double _soch, _scd, _sich, _ctth, _ctbh, _cd;
        bool _isReversed;
        #endregion
    }
}
