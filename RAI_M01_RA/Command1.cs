#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

#endregion

namespace RAI_M01_RA
{
    [Transaction(TransactionMode.Manual)]
    public class Command1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // 01. Get room parameters and filter for department names
            FilteredElementCollector roomCollector = new FilteredElementCollector(doc);
            roomCollector.OfCategory(BuiltInCategory.OST_Rooms);
            roomCollector.Where(a => a.LookupParameter("Department") != null);
            
            Element roomInst = roomCollector.FirstElement();

            // 02. Get room parameter values for schedule fields
            Parameter roomNumParam = roomInst.get_Parameter(BuiltInParameter.ROOM_NUMBER);
            Parameter roomNameParam = roomInst.get_Parameter(BuiltInParameter.ROOM_NAME);
            Parameter roomDepartmentParam = roomInst.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT);
            Parameter roomCommentsParam = roomInst.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
            Parameter roomAreaParam = roomInst.get_Parameter(BuiltInParameter.ROOM_AREA);
            Parameter roomLevelParam = roomInst.get_Parameter(BuiltInParameter.ROOM_LEVEL_ID);

            List<string> roomDepartmentNames = new List<string>();

            foreach (SpatialElement filteredRoom in roomCollector)
            {
                // 03. Get department names and add to list
                Parameter roomDepartmentsParam = filteredRoom.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT);
                string roomDepartment = roomDepartmentsParam.AsString();    
                roomDepartmentNames.Add(roomDepartment);
            }

            // 04. Filter the department list for unique items
            List<string> roomDepartmentNamesUnique = roomDepartmentNames.Distinct().ToList();
            roomDepartmentNamesUnique.Sort();

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create Schedules");

                ElementId catId = new ElementId(BuiltInCategory.OST_Rooms);

                foreach (var roomDepartmentNameValue in roomDepartmentNamesUnique)
                    {
                        // 05. Create schedules
                        ViewSchedule newSchedule = ViewSchedule.CreateSchedule(doc, catId);
                        newSchedule.Name = "Dept - " + roomDepartmentNameValue;

                        // 06. Create fields
                        ScheduleField roomNumField = newSchedule.Definition.AddField(ScheduleFieldType.Instance, roomNumParam.Id);
                        ScheduleField roomNameField = newSchedule.Definition.AddField(ScheduleFieldType.Instance, roomNameParam.Id);
                        ScheduleField roomDepartmentField = newSchedule.Definition.AddField(ScheduleFieldType.Instance, roomDepartmentParam.Id);
                        ScheduleField roomCommentsField = newSchedule.Definition.AddField(ScheduleFieldType.Instance, roomCommentsParam.Id);
                        ScheduleField roomAreaField = newSchedule.Definition.AddField(ScheduleFieldType.ViewBased, roomAreaParam.Id);
                        ScheduleField roomLevelField = newSchedule.Definition.AddField(ScheduleFieldType.Instance, roomLevelParam.Id);

                        roomLevelField.IsHidden = true;
                        roomAreaField.DisplayType = ScheduleFieldDisplayType.Totals;

                        // 07. Filter by department
                        ScheduleFilter departmentFilter = new ScheduleFilter(roomDepartmentField.FieldId, ScheduleFilterType.Equal, roomDepartmentNameValue);
                        newSchedule.Definition.AddFilter(departmentFilter);

                        // 08. Group schedule data by level 
                        ScheduleSortGroupField levelSort = new ScheduleSortGroupField(roomLevelField.FieldId);
                        levelSort.ShowHeader = true;
                        levelSort.ShowFooter = true;
                        levelSort.ShowBlankLine = true;
                        newSchedule.Definition.AddSortGroupField(levelSort);

                        // 09. Sort schedule data by room name 
                        ScheduleSortGroupField nameSort = new ScheduleSortGroupField(roomNameField.FieldId);
                        newSchedule.Definition.AddSortGroupField(nameSort);

                        // 10. Set totals to display
                        newSchedule.Definition.IsItemized = true;
                        newSchedule.Definition.ShowGrandTotal = true;
                        newSchedule.Definition.ShowGrandTotalTitle = true;
                        newSchedule.Definition.ShowGrandTotalCount = true;
                    }

                // 11. Create schedules
                ViewSchedule newSchedule2 = ViewSchedule.CreateSchedule(doc, catId);
                newSchedule2.Name = "All Departments";

                // 12. Create fields
                ScheduleField roomDepartmentField1 = newSchedule2.Definition.AddField(ScheduleFieldType.Instance, roomDepartmentParam.Id);
                ScheduleField roomAreaField1 = newSchedule2.Definition.AddField(ScheduleFieldType.ViewBased, roomAreaParam.Id);

                roomAreaField1.DisplayType = ScheduleFieldDisplayType.Totals;

                // 13. Group schedule data by department 
                ScheduleSortGroupField departmentSort = new ScheduleSortGroupField(roomDepartmentField1.FieldId);
                departmentSort.ShowHeader = true;
                departmentSort.ShowFooter = true;
                departmentSort.ShowBlankLine = true;
                newSchedule2.Definition.AddSortGroupField(departmentSort);

                // 10. Set totals to display
                newSchedule2.Definition.IsItemized = false;
                newSchedule2.Definition.ShowGrandTotal = true;
                newSchedule2.Definition.ShowGrandTotalTitle = true;
                newSchedule2.Definition.ShowGrandTotalCount = true;

                t.Commit();

            }

            return Result.Succeeded;
        }

        private Level GetLevelByName(Document doc, string levelName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_Levels);
            collector.WhereElementIsNotElementType();

            foreach (Level curLevel in collector)
            {
                if (curLevel.Name == levelName)
                    return curLevel;
            }
            return null;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData1.Data;
        }
    }
}
