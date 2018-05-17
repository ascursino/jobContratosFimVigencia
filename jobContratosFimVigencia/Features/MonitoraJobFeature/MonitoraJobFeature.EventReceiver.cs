using System;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Security;
using System.Linq; 


namespace jobContratosFimVigencia.Features.MonitoraJobFeature
{
    /// <summary>
    /// This class handles events raised during feature activation, deactivation, installation, uninstallation, and upgrade.
    /// </summary>
    /// <remarks>
    /// The GUID attached to this class may be used during packaging and should not be modified.
    /// </remarks>

    [Guid("81cd5045-431c-4be8-a869-4cb92a966129")]
    public class MonitoraJobFeatureEventReceiver : SPFeatureReceiver
    {
        // Name of the Timer Job, but not the Title which is displayed in central admin 
        private const string List_JOB_NAME = "jobContratosFimVigencia";

        public override void FeatureActivated(SPFeatureReceiverProperties properties)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                SPSite site = properties.Feature.Parent as SPSite;

                // make sure the job isn't already registered 
                site.WebApplication.JobDefinitions.Where(t => t.Name.Equals(List_JOB_NAME)).ToList().ForEach(j => j.Delete());

                //job por minuto 
                /*
                MonitoraJob listLoggerJob = new MonitoraJob(List_JOB_NAME, site.WebApplication);
                SPMinuteSchedule schedule = new SPMinuteSchedule();
                //schedule.BeginSecond = 0;
                //schedule.EndSecond = 59;
                schedule.Interval = 1;
                */

                //job por dia 
                MonitoraJob listLoggerJob = new MonitoraJob(List_JOB_NAME, site.WebApplication);
                SPDailySchedule schedule = new SPDailySchedule();
                //schedule.BeginSecond = 0;
                //schedule.EndSecond = 59;
                schedule.BeginHour = 5;
                schedule.EndHour = 6;

                listLoggerJob.Schedule = schedule;
                listLoggerJob.Update();
            });
        }

        public override void FeatureDeactivating(SPFeatureReceiverProperties properties)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                SPSite site = properties.Feature.Parent as SPSite;

                // delete the job 
                site.WebApplication.JobDefinitions.Where(t => t.Name.Equals(List_JOB_NAME)).ToList().ForEach(j => j.Delete());
            });
        }
    }
}

