﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using Tekla.Structures;
using Tekla.Structures.Model;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model.Operations;


namespace ZeroTouchTekla
{
    public class Utility
    {
        public static Vector GetVectorFromTwoPoints(Point startPoint, Point endPoint)
        {
            Vector vector = new Vector(endPoint.X - startPoint.X, endPoint.Y - startPoint.Y, endPoint.Z - startPoint.Z);
            return vector;
        }
        public static GeometricPlane GetPlaneFromFace(RebarLegFace face)
        {
            Point firstPoint = face.Contour.ContourPoints[0] as Point;
            Point secondPoint = face.Contour.ContourPoints[1] as Point;
            Point thirdPoint = face.Contour.ContourPoints[2] as Point;

            Vector firstVector = GetVectorFromTwoPoints(firstPoint, secondPoint).GetNormal();
            Vector secondVector = GetVectorFromTwoPoints(firstPoint, thirdPoint).GetNormal();

            GeometricPlane facePlane = new GeometricPlane(firstPoint, firstVector, secondVector);

            return facePlane;
        }
        public static void CopyGuideLine()
        {
            Model model = new Model();
            Operation.DisplayPrompt("Pick rebar set to copy spacing from");
            Tekla.Structures.Model.UI.Picker picker = new Tekla.Structures.Model.UI.Picker();
            Tekla.Structures.Model.UI.Picker.PickObjectEnum pickObjectEnum = Tekla.Structures.Model.UI.Picker.PickObjectEnum.PICK_ONE_REINFORCEMENT;
            RebarSet rebarSet = picker.PickObject(pickObjectEnum) as RebarSet;
            List<RebarGuideline> rebarGuidelines = rebarSet.Guidelines;
            List<RebarSpacing> referenceSpacings = new List<RebarSpacing>();
            foreach (RebarGuideline rebarGuideline in rebarGuidelines)
            {
                RebarSpacing rebarSpacing = rebarGuideline.Spacing;
                referenceSpacings.Add(rebarSpacing);
            }

            Operation.DisplayPrompt("Pick rebar sets to copy spacing to");
            Tekla.Structures.Model.UI.Picker secondPicker = new Tekla.Structures.Model.UI.Picker();
            Tekla.Structures.Model.UI.Picker.PickObjectsEnum pickObjectEnums = Tekla.Structures.Model.UI.Picker.PickObjectsEnum.PICK_N_REINFORCEMENTS;
            ModelObjectEnumerator rebarEnum = picker.PickObjects(pickObjectEnums);
            var rebarList = ToList(rebarEnum);

            List<RebarSet> rs = (from RebarSet r in rebarList
                                 select r).ToList();

            foreach (RebarSet r in rebarList)
            {
                List<RebarGuideline> rgls = r.Guidelines;
                for (int i = 0; i < rgls.Count; i++)
                {
                    rgls[i].Spacing = referenceSpacings[i];
                }
                r.Guidelines = rgls;
                r.Modify();
            }

            model.CommitChanges();
        }
        public static List<ModelObject> ToList(ModelObjectEnumerator enumerator)
        {
            var list = new List<ModelObject>();
            while (enumerator.MoveNext())
            {
                var current = enumerator.Current;
                list.Add(current);
            }
            return list;
        }
    }
}