using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;

using ProSpace;

namespace BlazorWebViewOnWPF.Services
{
    public class ProspaceService
    {
        public const string Prospace = "ProSpace.Application";

        public string GetCurrentPlanogram()
        {
            Application app = null;
            Planogram plano = null;
            try
            {
                app = GetRunningProspaceInstance();

                if (app == null) return "No active JDA instance";

                return (string)app.ActivePlanogram.PlanogramField[ProSpace.ePlanogramFields.PlanogramName];

            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                if (app != null)
                {
                    Marshal.ReleaseComObject(app);
                    app = null;
                }

                if (plano != null)
                {
                    Marshal.ReleaseComObject(plano);
                    plano = null;
                }
            }
        }

        private static Application? GetRunningProspaceInstance()
        {
            var instances = GetAppInstances(Prospace);
            Application? app = null;

            foreach(var inst in instances)
            {
                app = (Application)inst;
                if (app.Visible)
                    return app;

                Marshal.ReleaseComObject(app);
                app = null;
            }

            return app;
        }

        [DllImport("ole32.dll")]
        static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        [DllImport("ole32.dll")]

        public static extern void GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

        private static List<object> GetAppInstances(string name)
        {
            List<object> runningAppInstances = new List<object>();
            Hashtable runningObjects = GetRunningObjectTable();

            IDictionaryEnumerator rotEnumerator = runningObjects.GetEnumerator();
            while (rotEnumerator.MoveNext())
            {
                string candidateName = (string)rotEnumerator.Key;
                if (!candidateName.StartsWith(name))
                    continue;
                runningAppInstances.Add(rotEnumerator.Value);
            }
            return runningAppInstances;
        }
        private static Hashtable GetRunningObjectTable()
        {
            Hashtable result = new Hashtable();

            IntPtr numFetched = new IntPtr();
            IRunningObjectTable runningObjectTable;
            IEnumMoniker monikerEnumerator;
            IMoniker[] monikers = new IMoniker[1];

            GetRunningObjectTable(0, out runningObjectTable);
            runningObjectTable.EnumRunning(out monikerEnumerator);
            monikerEnumerator.Reset();

            while (monikerEnumerator.Next(1, monikers, numFetched) == 0)
            {
                IBindCtx ctx;
                CreateBindCtx(0, out ctx);

                string runningObjectName;
                monikers[0].GetDisplayName(ctx, null, out runningObjectName);

                object runningObjectVal;
                runningObjectTable.GetObject(monikers[0], out runningObjectVal);

                result[runningObjectName] = runningObjectVal;
            }

            return result;
        }
    }
}
