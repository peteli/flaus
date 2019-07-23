using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace peteli.flaus
{
    /// <summary>
    /// This Class is TOP CLASS in represents application model.
    /// All other class are subsequent classes to this class.
    /// Model initializes importent dictionary (lookup) arrays to calculate fast for good user experiences
    /// </summary>
    public class FlausModel
    {
		public FlausModel(Excel.Workbook workbook)
        {
            this.Workbook = workbook;

            Excel.Worksheet worksheet = workbook.Worksheets[1];

            //worksheet.Change

            InitListObjectsNeeded();
            
        }

        private void InitListObjectsNeeded()
        {
            //throw new NotImplementedException();
            foreach(Excel.Worksheet worksheet in this.Workbook.Worksheets)
            {
                foreach (Excel.ListObject listObject in worksheet.ListObjects)
                {
                    if(listObject.Name == Properties.Settings.Default.ListObjectNameProjects)
                    {
                        Debug.WriteLine("Found project table");
                        TableProjects = listObject;
                        // subscribe to workbook change event
                        worksheet.Change += WorksheetProjects_Change;
                    }
                    if(listObject.Name == Properties.Settings.Default.ListObjectNameFloorGridSizes)
                    {
                        Debug.WriteLine("Found Floor Grid Size table");
                        TableFloorGridSize = listObject;
                        // subscribe to workbook change event
                        worksheet.Change += WorksheetFloorGridSize_Change;
                    }
                    if (listObject.Name == Properties.Settings.Default.ListObjectNameExceptedAllocations)
                    {
                        Debug.WriteLine("Found Exception Allocation table");
                        TableExceptedAllocations = listObject;
                    }
                }
            }
        }

        private bool DetermineModelIntegrity()
        {
            if(TableProjects != null && TableFloorGridSize != null && TableFloorGridSize != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private void WorksheetFloorGridSize_Change(Excel.Range Target)
        {
            Debug.WriteLine("Floor Grid Size Worksheet changed");
        }

        private void WorksheetProjects_Change(Excel.Range Target)
        {
            Debug.WriteLine("Projects Worksheet changed");
        }

     

        #region Properties
        // table projects
        internal SortedList m_objprojects { get; set; }
        // table floor grid sizes
        internal SortedList m_objgridMaster { get; set; }
        // table manual PVBs -deviations from project's baseline grid
        internal SortedList m_objgridDeviation { get; set; }
        // table for index that delivers grid utilization by week and floorCategory
        internal SortedList m_objidxUtilizationByCategoryByWeek { get; set; }
        // table for index that delivers grid allocation by week and gridCoordinate
        internal SortedList m_objidxAllocationByWeekBySegment { get; set; }
        // table for index that delivers grid allocation by week and project
        internal SortedList m_objidxReportByProjectByWeek { get; set; }
        // table for index that delivers space utilization by project by week for power pivot application
        internal SortedList m_objidxReportFloor { get; set; }

        // reference to ListObject Project
        internal Excel.ListObject TableProjects { get; set; }
        // reference to ListObject Floor Grid Sizes
        internal Excel.ListObject TableFloorGridSize { get; set; }
        // referemce to ListObject Exception PVB Assignments 
        internal Excel.ListObject TableExceptedAllocations { get; set; }


        public string Name = "Hello World";
        public bool ModelIntegrity => (TableProjects != null && TableFloorGridSize != null && TableFloorGridSize != null);
        private Excel.Workbook Workbook { get; set; }

        #endregion
    }
}
