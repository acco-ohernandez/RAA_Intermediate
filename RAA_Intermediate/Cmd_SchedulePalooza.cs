#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Windows.Controls;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

#endregion

namespace RAA_Intermediate
{
    [Transaction(TransactionMode.Manual)]
    public class Cmd_SchedulePalooza : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            //TaskDialog.Show("test", "Module 1 challange");

            try
            {
                // Get All the Rooms in the document orderd by name
                var rooms = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms)
                                                             .WhereElementIsNotElementType()
                                                             .OrderBy(r => r.Name)
                                                             .ToList();

                // Get all the unique department names from the rooms.
                var uniqueDepartments = rooms
                                        .Select(room => GetDepartment(room))   // Project each room to its department using GetDepartment method
                                        .Distinct()                            // Get distinct department names
                                        .ToList();                             // Convert the result to a list

                // List to store the schedules created
                var schedulesCreatedList = new List<ViewSchedule>();
                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Create Department Schedules");

                    // Create Schedule for each department
                    foreach (var departmentName in uniqueDepartments)
                    {
                        var createdSchedule = CreateNewDepatmentSchedule(doc, departmentName);
                        var createdScheduleWithParams = AddParamsToCreatedSchedule(createdSchedule, rooms, departmentName);
                        schedulesCreatedList.Add(createdSchedule);

                    }



                    t.Commit();
                }

            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }

            return Result.Succeeded;
        }

        private ViewSchedule AddParamsToCreatedSchedule(ViewSchedule schedule, List<Element> rooms, string departmentName)
        {
            // Create a copy of the input schedule
            ViewSchedule curSchedule = schedule;
            Element roomInstance = rooms.First();
            Parameter numberParam = roomInstance.LookupParameter("Number");
            Parameter levelParam = roomInstance.LookupParameter("Level");
            Parameter roomNameParam = roomInstance.LookupParameter("Name");
            Parameter DepartmenParam = roomInstance.LookupParameter("Department");
            Parameter commentsParam = roomInstance.LookupParameter("Comments");
            Parameter areaParam = roomInstance.get_Parameter(BuiltInParameter.ROOM_AREA);

            // Get the fields for "Level" and "Room Name" from the schedule
            ScheduleField roomNumberField = curSchedule.Definition.AddField(ScheduleFieldType.Instance, numberParam.Id);
            ScheduleField roomlevelField = curSchedule.Definition.AddField(ScheduleFieldType.Instance, levelParam.Id);
            ScheduleField roomNameField = curSchedule.Definition.AddField(ScheduleFieldType.Instance, roomNameParam.Id);
            ScheduleField roomDeptField = curSchedule.Definition.AddField(ScheduleFieldType.Instance, DepartmenParam.Id);
            ScheduleField roomCommentsField = curSchedule.Definition.AddField(ScheduleFieldType.Instance, commentsParam.Id);
            ScheduleField roomAreaField = curSchedule.Definition.AddField(ScheduleFieldType.ViewBased, areaParam.Id);

            // Hide Level fields
            roomlevelField.IsHidden = true;
            // Show Area Totals
            roomAreaField.DisplayType = ScheduleFieldDisplayType.Totals;
            // Show Department count
            //roomDeptField.DisplayType = ScheduleFieldDisplayType.;

            // Filter by Department name
            ScheduleFilter deptFilter = new ScheduleFilter(roomDeptField.FieldId, ScheduleFilterType.Equal, departmentName);
            curSchedule.Definition.AddFilter(deptFilter);

            // Group schedule data by Level
            ScheduleSortGroupField typeSort = new ScheduleSortGroupField(roomlevelField.FieldId);
            typeSort.ShowHeader = true;
            typeSort.ShowFooter = true;
            typeSort.ShowBlankLine = true;
            curSchedule.Definition.AddSortGroupField(typeSort);

            // Sort schedule data by room name
            ScheduleSortGroupField nameSort = new ScheduleSortGroupField(roomNameField.FieldId);
            curSchedule.Definition.AddSortGroupField(nameSort);

            curSchedule.Definition.IsItemized = true;
            curSchedule.Definition.ShowGrandTotal = true;
            curSchedule.Definition.ShowGrandTotalTitle = true;
            curSchedule.Definition.ShowGrandTotalCount = true;

            return curSchedule;
        }

        /// <summary>
        /// Creates a new department-specific view schedule for rooms in the document.
        /// </summary>
        /// <param name="doc">The document in which to create the schedule.</param>
        /// <param name="department">The name of the department for which to create the schedule.</param>
        /// <returns>
        /// The newly created department-specific view schedule.
        /// </returns>
        private ViewSchedule CreateNewDepatmentSchedule(Document doc, string department)
        {
            // Define the category for rooms.
            var roomsCat = new ElementId(BuiltInCategory.OST_Rooms);

            // Create a new schedule for rooms in the specified category.
            ViewSchedule newSchedule = ViewSchedule.CreateSchedule(doc, roomsCat);

            // Set the name of the new schedule to include the department name.
            newSchedule.Name = $"Dept - {department}";

            // Return the newly created department-specific view schedule.
            return newSchedule;
        }

        private string GetDepartment(Element room)
        {
            // Get the "Department" parameter in the room.
            var department = room.ParametersMap
                        .Cast<Parameter>()
                        .FirstOrDefault(p => p.Definition.Name == "Department");

            // if department not null return department as value string, else returne "No Department"
            return department != null ? department.AsValueString() : "No Department";
        }


        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btn_SchedulePalooza";
            string buttonTitle = "Schedule-Palooza";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "RAA_Intermediate - Module 01 Challenge");

            return myButtonData1.Data;
        }
    }
}
