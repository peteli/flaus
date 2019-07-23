using System.Diagnostics;
using ExcelDna.Integration;
using Excel=Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Text;
using System;
using System.Collections.Generic;

namespace peteli.flaus
{
    public class MyAddIn : IExcelAddIn
    {
        #region Properties
        internal static Dictionary<Excel.Workbook,FlausModel> FlausModelByWorkbook = new Dictionary<Excel.Workbook, FlausModel>();
        internal static Excel.Application _xlApplication = (Excel.Application)ExcelDnaUtil.Application;
        internal static FlausModel ActiveWorkbookFlausModel;
        #endregion

        public void AutoOpen()
        {
            // get handle of excel application
            Debug.WriteLine("XLL AutoOpen runs");
            InitCollectAllWorkbooks();
            ((Excel.AppEvents_Event)_xlApplication).WorkbookOpen += MyAddIn_WorkbookOpen;
            ((Excel.AppEvents_Event)_xlApplication).WorkbookActivate += MyAddIn_WorkbookActivate;
            ((Excel.AppEvents_Event)_xlApplication).NewWorkbook += MyAddIn_NewWorkbook;

            InitAppsHooks();
            
            
        }

        private void MyAddIn_WorkbookActivate(Excel.Workbook Wb)
        {
            Debug.WriteLine("MyAddIn_WorkbookActivate runs");
            if (FlausModelByWorkbook.ContainsKey(Wb))
            {
                ActiveWorkbookFlausModel = FlausModelByWorkbook[Wb];
                Debug.WriteLine(Wb.Flaus().ModelIntegrity);
            }
            else
            {
                ActiveWorkbookFlausModel = null;
                Debug.WriteLine("Flaus Modell is null");
            }

        }

        private void MyAddIn_WorkbookOpen(Excel.Workbook Wb)
        {
            Debug.WriteLine("MyAddIn_WorkbookOpen runs");
            FlausModel flausModel = new FlausModel(Wb);
            if (flausModel.ModelIntegrity)
            {
                FlausModelByWorkbook.Add(Wb, flausModel);
            }
        }

        private void MyAddIn_NewWorkbook(Excel.Workbook Wb)
        {
            Debug.WriteLine("MyAddIn_NewWorkbook runs");
            FlausModel flausModel = new FlausModel(Wb);
            if (flausModel.ModelIntegrity)
            {
                FlausModelByWorkbook.Add(Wb, flausModel);
            }
        }


        private void InitCollectAllWorkbooks()
        {
            //throw new NotImplementedException();
           
            foreach (Excel.Workbook workbook in _xlApplication.Workbooks)
            {
                FlausModelByWorkbook.Add(workbook, new FlausModel(workbook));
            }
        }

        private void InitAppsHooks()
        {
            //Application _XlApp = (Application)ExcelDnaUtil.Application;
            //_XlApp.Workbooks
            //_XlApp.WorkbookActivate += DeleteCTP;
            //_XlApp.WorkbookDeactivate += DeleteCTP;

            
            

        }

        private static void AddWorkbookToList(Excel.Workbook Wb)
        {
            FlausModelByWorkbook.Add(Wb,new FlausModel(Wb));
        }

        private void DeleteCTP(Excel.Workbook Wb)
        {
            //CTPManager.DeleteCTP();
        }

        public void AutoClose()
        {
            // put code here
            Debug.WriteLine("XLL closes");
        }

    }
}
