using DevExpress.XtraBars.Helpers;
using DevExpress.XtraBars.Ribbon;
using DevExpress.XtraEditors;
using DevExpress.XtraScheduler;
using DevExpress.XtraScheduler.GoogleCalendar;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GCSync {
    public partial class Form1 : RibbonForm {
        UserCredential credential;
        CalendarService service;
        CalendarList calendarList;
        string activeCalendarId;
        bool allowEventLoad;
        public Form1() {
            InitializeComponent();
            InitSkinGallery();
            schedulerControl.Start = System.DateTime.Now;
            //ricbCalendarList.SelectedIndexChanged += RicbCalendarList_SelectedIndexChanged;
           // bbiSynchronize.ItemClick += BbiSynchronize_ItemClick;
            gcSyncComponent.ConflictDetected += GcSyncComponent_ConflictDetected;

        }

        private void GcSyncComponent_ConflictDetected(object sender, ConflictDetectedEventArgs e) {
            XtraMessageBox.Show("Google '" + e.Event.Summary + "' Event conflicts with the Scheduler '" +
                e.Appointment.Subject + "' Appointment." + Environment.NewLine + "Synchronizing by the Google Event.",
                "Conflict detected", MessageBoxButtons.OK);
            //uncomment the following line to sync by Scheduler Appointments instead
            e.GoogleEventIsValid = true;
        }

        //manually trigger synchronization
        //private void BbiSynchronize_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e) {

        //    this.gcSyncComponent.Synchronize();
        //}

        //assign a selected calendar ID to the component's CalendarID property
        private void RicbCalendarList_SelectedIndexChanged(object sender, EventArgs e) {





            //ComboBoxEdit edit = (ComboBoxEdit)sender;
            //string selectedCalendarSummary = (string)edit.SelectedItem;
            //CalendarListEntry selectedCalendar = this.calendarList.Items.FirstOrDefault(x => x.Summary == selectedCalendarSummary);
            //this.activeCalendarId = selectedCalendar.Id;
            //this.gcSyncComponent.CalendarId = selectedCalendar.Id;
            //this.gcSyncComponent.Synchronize(); 
            //UpdateBbiAvailability();

            //MessageBox.Show(selectedCalendar.Id.ToString());
        }


        #region Authorization
        async protected override void OnLoad(EventArgs e)
        {
            XtraMessageBoxArgs args = new XtraMessageBoxArgs();
            args.AutoCloseOptions.Delay = 3000;
            args.Caption = "Auto-close message";
            args.Text = "This message closes automatically after 3 seconds.";
            args.Buttons = new DialogResult[] { DialogResult.OK, DialogResult.Cancel };



            ////////////:
            base.OnLoad(e);
            Directory.CreateDirectory(Environment.CurrentDirectory + @"\xml");
            try {
            this.credential = await AuthorizeToGoogle();
            this.service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = this.credential,
                ApplicationName = "Other client 1"
            });
            this.gcSyncComponent.CalendarService = this.service;
           // await UpdateCalendarListUI();
            this.allowEventLoad = true;
            UpdateBbiAvailability();
            this.gcSyncComponent.Storage = schedulerStorage;


               // this.gcSyncComponent.CalendarId = "652ilqp8u3ikeur3ac2uqd1t8s@group.calendar.google.com";
               // this.gcSyncComponent.Synchronize();
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message);
               // XtraMessageBox.Show(args).ToString();

            }
        }

async Task<UserCredential> AuthorizeToGoogle() {
            using (FileStream stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read)) {
                string credPath = Environment.GetFolderPath(
                    Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/GoogleSchedulerSync.json");

                return await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    new String[] { CalendarService.Scope.Calendar },
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true));
            }
        }
        #endregion

       // retrieve calendars from a signed user account
        async Task UpdateCalendarListUI()
        {
            CalendarListResource.ListRequest listRequest = this.service.CalendarList.List();
            this.calendarList = await listRequest.ExecuteAsync();
            this.ricbCalendarList.Items.Clear();
            foreach (CalendarListEntry item in this.calendarList.Items)
                if (!ricbCalendarList.Items.Contains(item.Summary))
                {
                    this.ricbCalendarList.Items.Add(item.Summary);
                }
            if (!String.IsNullOrEmpty(this.activeCalendarId))
            {
                CalendarListEntry itemToSelect = this.calendarList.Items.FirstOrDefault(x => x.Id == this.activeCalendarId);
                this.gcSyncComponent.CalendarId = this.activeCalendarId;
                if (this.ricbCalendarList.Items.Contains(itemToSelect.Summary))
                {

                    this.beiCalendarList.EditValue = itemToSelect.Summary;

                }
                else
                    this.activeCalendarId = String.Empty;
            }
            UpdateBbiAvailability();
        }

        //specifies whether the "Sync" button should be enabled
        void UpdateBbiAvailability() {
            this.bbiSynchronize.Enabled = !String.IsNullOrEmpty(this.activeCalendarId) && this.allowEventLoad;
        }

        void InitSkinGallery() {
            SkinHelper.InitSkinGallery(rgbiSkins, true);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'dossierMarcherDataSet.Appointments' table. You can move, or remove it, as needed.
            //  this.appointmentsTableAdapter.Fill(this.dossierMarcherDataSet.Appointments);





        }

        private void bbiSynchronize_ItemClick_1(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {

        }
      

        private void barButtonItem1_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
          //  CalendarService myService = new CalendarService("your calendar name");
          //  myService.setUserCredentials(username, password);

          //  CalendarEntry calendar;

            


            XtraMessageBoxArgs args = new XtraMessageBoxArgs();
            args.AutoCloseOptions.Delay = 3000;
            args.Caption = "Auto-close message";
            args.Text = "This message closes automatically after 3 seconds.";
            args.Buttons = new DialogResult[] { DialogResult.OK, DialogResult.Cancel };


            try
            {
                DataSet ds = new DataSet();
                for (int count = 0; count < ds.Tables.Count; count++)
                {
                    // Get individual datatables here...
                    DataTable table = ds.Tables[count];
                }


                dossierMarcherDataSet.Clear();
                dossierMarcherDataSet.Tables.Clear();

                this.appointmentsTableAdapter.Fill(this.dossierMarcherDataSet.Appointments);

                this.gcSyncComponent.CalendarId = "46o41vhg3e1o0m0d2kmkomlvtg@group.calendar.google.com";
                this.gcSyncComponent.Synchronize();



                

            //    this.gcSyncComponent.CalendarId = "652ilqp8u3ikeur3ac2uqd1t8s@group.calendar.google.com";
              //  this.gcSyncComponent.Synchronize();
            }

            catch (Exception ex)
            {
                XtraMessageBox.Show(args + ex.Message).ToString();

               
            }
        }

        private void schedulerStorage_AppointmentsChanged(object sender, PersistentObjectsEventArgs e)
        {
         


            //bool commitLock = false;
            //void CommitTask()
            //{
            //    if (commitLock)
            //        return;
            //    commitLock = true;
            //    try
            //    {
            //        appointmentsTableAdapter.Update(dossierMarcherDataSet);
            //        dossierMarcherDataSet.AcceptChanges();
            //    }
            //    finally
            //    {
            //        commitLock = false;
            //    }
            //}
        }

        private void ribbonControl_Click(object sender, EventArgs e)
        {

        }

        private void schedulerControl_AllowAppointmentCopy(object sender, AppointmentOperationEventArgs e)
        {
      
        }

        private void barButtonItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {

            /// <summary>
            //that's not delete from google calendar
            /// </summary>
            /// 
            // schedulerStorage.Appointments.Clear();



            //for (int i = 0; i < schedulerControl.SelectedAppointments.Count; i++)
            //{


            /// <summary>
            //that's   delete  selected item from google calendar
            /// </summary>

            //    schedulerStorage.Appointments.Remove(schedulerControl.SelectedAppointments[i]);


            //}
            //}



        }

        private void gcSyncComponent_ConflictDetected_1(object sender, ConflictDetectedEventArgs e)
        {
            e.GoogleEventIsValid = true;

        }

        private void gcSyncComponent_FilterAppointments(object sender, FilterAppointmentsEventArgs e)
        {
            //I test that code but it just stopping add the appointments to google calendar


            //  if (e.Appointment != null && e.Appointment.Description.Contains("1"))
            //        e.Cancel = true;

        }
    }
}