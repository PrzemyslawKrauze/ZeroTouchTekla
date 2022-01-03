using System;
using System.Collections.Generic;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model;
using System.Text.RegularExpressions;
using System.Linq;

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
                }
                else
                {
                    SetRTWProfilePoints(element);
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
                }
            }
        }
        void SetRTWProfilePoints(Element element)
        {
            RTW rtw = element as RTW;
            double correctedHeight1 = (rtw.Height2 - rtw.Height) * (SCD) / rtw.Length + rtw.Height;
            double correctedHeight2 = (rtw.Height2 - rtw.Height) * (rtw.Length - CD) / rtw.Length + rtw.Height;
            double s = (rtw.BottomWidth - (rtw.TopWidth - rtw.CorniceWidth)) / rtw.Height;

            List<List<Point>> profilePoints = rtw.GetProfilePoints();
            List<List<Point>> correctedPoints = new List<List<Point>>();

            Point p00 = profilePoints[0][0];
            Point p01 = new Point(p00.X, p00.Y + correctedHeight1 - SOCH, p00.Z);
            Point p03 = profilePoints[0][5];
            Point p02 = new Point(p00.X, p01.Y, p01.Z - rtw.BottomWidth + s * Distance.PointToPoint(p00, p01));
            correctedPoints.Add(new List<Point> { p00, p01, p02, p03 });

            Point p10 = new Point(p00.X + SCD, p00.Y, p00.Z);
            Point p11 = new Point(p10.X, p10.Y + correctedHeight1 - SICH, p10.Z);
            Point p13 = new Point(p10.X, p03.Y, p03.Z);
            Point p12 = new Point(p10.X, p01.Y, p10.Z - rtw.BottomWidth + s * Distance.PointToPoint(p10, p11));
            correctedPoints.Add(new List<Point> { p10, p11, p12, p13 });

            Point p20 = p10;
            Point p25 = p13;
            Point p21 = new Point(p20.X, p20.Y + correctedHeight1 - rtw.CorniceHeight, p20.Z);
            Point p22 = new Point(p21.X, p21.Y, p21.Z + rtw.CorniceWidth);
            Point p23 = new Point(p22.X, p22.Y + rtw.CorniceHeight, p22.Z);
            Point p24 = new Point(p23.X, p23.Y, p23.Z - rtw.TopWidth);
            correctedPoints.Add(new List<Point> { p20, p21, p22, p23, p24, p25 });

            Point p30 = new Point(profilePoints[1][0].X - CD, p20.Y, p20.Z);
            Point p31 = new Point(p30.X, p30.Y + correctedHeight2 - rtw.CorniceHeight, p30.Z);
            Point p32 = new Point(p30.X, p31.Y, p31.Z + rtw.CorniceWidth);
            Point p33 = new Point(p32.X, p32.Y + rtw.CorniceHeight, p32.Z);
            Point p34 = new Point(p33.X, p33.Y, p33.Z - rtw.TopWidth);
            Point p35 = new Point(p30.X, p25.Y, p25.Z);
            correctedPoints.Add(new List<Point> { p30, p31, p32, p33, p34, p35 });

            Point p40 = new Point(p30.X, p30.Y + CtBH, p30.Z);
            Point p45 = new Point(p30.X, p40.Y, p40.Z - rtw.BottomWidth + s * CtBH);
            Point p41 = new Point(p40.X, p30.Y + correctedHeight2 - rtw.CorniceHeight, p40.Z);
            Point p42 = new Point(p41.X, p41.Y, p41.Z + rtw.CorniceWidth);
            Point p43 = new Point(p42.X, p42.Y + rtw.CorniceHeight, p42.Z);
            Point p44 = new Point(p43.X, p43.Y, p43.Z - rtw.TopWidth);
            correctedPoints.Add(new List<Point> { p40, p41, p42, p43, p44, p45 });

            Point p50 = new Point(profilePoints[1][0].X, p30.Y + rtw.Height2 - CtTH, p30.Z);
            Point p51 = new Point(p50.X, p30.Y + rtw.Height2 - rtw.CorniceHeight, p50.Z);
            Point p52 = new Point(p51.X, p51.Y, p51.Z - rtw.CorniceWidth);
            Point p53 = new Point(p52.X, p52.Y + rtw.CorniceHeight, p52.Z);
            Point p54 = new Point(p53.X, p53.Y, p53.Z - rtw.TopWidth);
            Point p55 = new Point(p50.X, p50.Y, p30.Z + rtw.BottomWidth - s * (rtw.Height2 - CtTH));
            correctedPoints.Add(new List<Point> { p50, p51, p52, p53, p54, p55 });

            ProfilePoints = correctedPoints;
        }
        void SetRTWProfilePointsReversed(Element element)
        {
            RTW rtw = element as RTW;
            double correctedHeight1 = (rtw.Height2 - rtw.Height) * (SCD) / rtw.Length + rtw.Height;
            double correctedHeight2 = (rtw.Height2 - rtw.Height) * (rtw.Length - CD) / rtw.Length + rtw.Height;
            double s = (rtw.BottomWidth - (rtw.TopWidth - rtw.CorniceWidth)) / rtw.Height;

            List<List<Point>> profilePoints = rtw.GetProfilePoints();
            List<List<Point>> correctedPoints = new List<List<Point>>();

            Point p00 = profilePoints[1][0];
            Point p01 = new Point(p00.X, p00.Y + correctedHeight1 - SOCH, p00.Z);
            Point p03 = profilePoints[1][5];
            Point p02 = new Point(p00.X, p01.Y, p01.Z - rtw.BottomWidth + s * Distance.PointToPoint(p00, p01));
            correctedPoints.Add(new List<Point> { p00, p01, p02, p03 });

            Point p10 = new Point(p00.X - SCD, p00.Y, p00.Z);
            Point p11 = new Point(p10.X, p10.Y + correctedHeight1 - SICH, p10.Z);
            Point p13 = new Point(p10.X, p03.Y, p03.Z);
            Point p12 = new Point(p10.X, p01.Y, p10.Z - rtw.BottomWidth + s * Distance.PointToPoint(p10, p11));
            correctedPoints.Add(new List<Point> { p10, p11, p12, p13 });

            Point p20 = p10;
            Point p25 = p13;
            Point p21 = new Point(p20.X, p20.Y + correctedHeight1 - rtw.CorniceHeight, p20.Z);
            Point p22 = new Point(p21.X, p21.Y, p21.Z + rtw.CorniceWidth);
            Point p23 = new Point(p22.X, p22.Y + rtw.CorniceHeight, p22.Z);
            Point p24 = new Point(p23.X, p23.Y, p23.Z - rtw.TopWidth);
            correctedPoints.Add(new List<Point> { p20, p21, p22, p23, p24, p25 });

            Point p30 = new Point(profilePoints[0][0].X + CD, p20.Y, p20.Z);
            Point p31 = new Point(p30.X, p30.Y + correctedHeight2 - rtw.CorniceHeight, p30.Z);
            Point p32 = new Point(p30.X, p31.Y, p31.Z + rtw.CorniceWidth);
            Point p33 = new Point(p32.X, p32.Y + rtw.CorniceHeight, p32.Z);
            Point p34 = new Point(p33.X, p33.Y, p33.Z - rtw.TopWidth);
            Point p35 = new Point(p30.X, p25.Y, p25.Z);
            correctedPoints.Add(new List<Point> { p30, p31, p32, p33, p34, p35 });

            Point p40 = new Point(p30.X, p30.Y + CtBH, p30.Z);
            Point p45 = new Point(p30.X, p40.Y, p40.Z - rtw.BottomWidth + s * CtBH);
            Point p41 = new Point(p40.X, p30.Y + correctedHeight2 - rtw.CorniceHeight, p40.Z);
            Point p42 = new Point(p41.X, p41.Y, p41.Z + rtw.CorniceWidth);
            Point p43 = new Point(p42.X, p42.Y + rtw.CorniceHeight, p42.Z);
            Point p44 = new Point(p43.X, p43.Y, p43.Z - rtw.TopWidth);
            correctedPoints.Add(new List<Point> { p40, p41, p42, p43, p44, p45 });

            Point p50 = new Point(profilePoints[0][0].X, p30.Y + rtw.Height2 - CtTH, p30.Z);
            Point p51 = new Point(p50.X, p30.Y + rtw.Height2 - rtw.CorniceHeight, p50.Z);
            Point p52 = new Point(p51.X, p51.Y, p51.Z - rtw.CorniceWidth);
            Point p53 = new Point(p52.X, p52.Y + rtw.CorniceHeight, p52.Z);
            Point p54 = new Point(p53.X, p53.Y, p53.Z - rtw.TopWidth);
            Point p55 = new Point(p50.X, p50.Y, p30.Z + rtw.BottomWidth - s * (rtw.Height2 - CtTH));
            correctedPoints.Add(new List<Point> { p50, p51, p52, p53, p54, p55 });

            ProfilePoints = correctedPoints;
        }
        void SetRTWProfilePointsReversed2(Element element)
        {
            RTW rtw = element as RTW;
            double correctedHeight1 = (rtw.Height2 - rtw.Height) * (SCD) / rtw.Length + rtw.Height;
            double correctedHeight2 = (rtw.Height2 - rtw.Height) * (rtw.Length - CD) / rtw.Length + rtw.Height;
            double s = (rtw.BottomWidth - (rtw.TopWidth - rtw.CorniceWidth)) / rtw.Height;

            List<List<Point>> profilePoints = rtw.GetProfilePoints();

            Point p50 = new Point(profilePoints[0][0].X, profilePoints[0][0].Y + rtw.Height - CtTH, profilePoints[0][0].Z);
            Point p51 = new Point(p50.X, profilePoints[0][0].Y + rtw.Height - rtw.CorniceHeight, p50.Z);
            Point p52 = new Point(p51.X, p51.Y, p51.Z - rtw.CorniceWidth);
            Point p53 = new Point(p52.X, p52.Y + rtw.CorniceHeight, p52.Z);
            Point p54 = new Point(p53.X, p53.Y, p53.Z - rtw.TopWidth);
            Point p55 = new Point(p50.X, p50.Y, profilePoints[0][0].Z + rtw.BottomWidth - s * (rtw.Height2 - CtTH));
            profilePoints.Add(new List<Point> { p50, p51, p52, p53, p54, p55 });

            Point p40 = new Point(profilePoints[0][0].X + CD, profilePoints[0][0].Y + CtBH, profilePoints[0][0].Z);
            Point p45 = new Point(p40.X, p40.Y, p40.Z - rtw.BottomWidth + s * CtBH);
            Point p41 = new Point(p40.X, profilePoints[0][0].Y + correctedHeight2 - rtw.CorniceHeight, p40.Z);
            Point p42 = new Point(p41.X, p41.Y, p41.Z + rtw.CorniceWidth);
            Point p43 = new Point(p42.X, p42.Y + rtw.CorniceHeight, p42.Z);
            Point p44 = new Point(p43.X, p43.Y, p43.Z - rtw.TopWidth);
            profilePoints.Add(new List<Point> { p40, p41, p42, p43, p44, p45 });

            Point p30 = new Point(p40.X, p40.Y, p40.Z);
            Point p31 = new Point(p30.X, p30.Y + correctedHeight2 - rtw.CorniceHeight, p30.Z);
            Point p32 = new Point(p30.X, p31.Y, p31.Z + rtw.CorniceWidth);
            Point p33 = new Point(p32.X, p32.Y + rtw.CorniceHeight, p32.Z);
            Point p34 = new Point(p33.X, p33.Y, p33.Z - rtw.TopWidth);
            Point p35 = new Point(p30.X, p45.Y, p45.Z);
            profilePoints.Add(new List<Point> { p30, p31, p32, p33, p34, p35 });

            Point p20 = new Point(p50.X + rtw.Length - SCD, p30.Y, p30.Z);
            Point p25 = new Point(p20.X, p20.Y, p30.Z);
            Point p21 = new Point(p20.X, p20.Y + correctedHeight1 - rtw.CorniceHeight, p20.Z);
            Point p22 = new Point(p21.X, p21.Y, p21.Z + rtw.CorniceWidth);
            Point p23 = new Point(p22.X, p22.Y + rtw.CorniceHeight, p22.Z);
            Point p24 = new Point(p23.X, p23.Y, p23.Z - rtw.TopWidth);
            profilePoints.Add(new List<Point> { p20, p21, p22, p23, p24, p25 });

            Point p10 = p20;
            Point p11 = new Point(p10.X, p10.Y + correctedHeight1 - SICH, p10.Z);
            Point p13 = p25;
            Point p12 = new Point(p10.X, p10.Y + correctedHeight1, p10.Z - rtw.BottomWidth + s * Distance.PointToPoint(p10, p11));
            profilePoints.Add(new List<Point> { p10, p11, p12, p13 });

            Point p00 = profilePoints[1][0];
            Point p01 = new Point(p00.X, p00.Y + correctedHeight1 - SOCH, p00.Z);
            Point p03 = profilePoints[1][5];
            Point p02 = new Point(p00.X, p01.Y, p01.Z - rtw.BottomWidth + s * Distance.PointToPoint(p00, p01));
            profilePoints.Add(new List<Point> { p00, p01, p02, p03 });

            ProfilePoints = profilePoints;
        }
        void RTWEditInnerVerticalRebar(RTW rtw, RebarSet rebarSet)
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
        #endregion
    }
}
