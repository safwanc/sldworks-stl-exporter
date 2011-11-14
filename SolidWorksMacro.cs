using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Diagnostics; 
using System;

namespace STLExporter.csproj
{
    public partial class SolidWorksMacro
    {
        public void Main()
        {
            Debug.WriteLine("Starting.."); 

            int swAPIWarnings = 0;
            int swAPIErrors = 0; 

            IModelDoc2 swActiveDoc = swApp.ActiveDoc as IModelDoc2;
            IAssemblyDoc swAssembly = swActiveDoc as IAssemblyDoc;

            Object swComponents = swAssembly.GetComponents(true);
            Array swComponentArray = (Array) swComponents;

            IComponent2 swAssemblyComponent;
            IModelDoc2 swAssemblyModel; 

            foreach (Object swComponent in swComponentArray)
            {
                swAssemblyComponent = swComponent as IComponent2;
                swAssemblyModel = swApp.ActivateDoc2(swAssemblyComponent.GetPathName(), true, ref swAPIWarnings) as IModelDoc2;

                swAssemblyModel.SaveAs4(
                    String.Format("{0}.STL", swAssemblyComponent.Name2), 
                    (int) swSaveAsVersion_e.swSaveAsCurrentVersion, 
                    (int) swSaveAsOptions_e.swSaveAsOptions_Copy, 
                    ref swAPIWarnings, 
                    ref swAPIErrors);

                swApp.CloseDoc(swAssemblyComponent.GetPathName()); 
            }


            Debug.WriteLine("Ending.."); 
        }

        /// <summary>
        ///  The SldWorks swApp variable is pre-assigned for you.
        /// </summary>
        public SldWorks swApp;
    }
}


