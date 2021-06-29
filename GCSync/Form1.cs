using DevExpress.XtraBars.Helpers;
using DevExpress.XtraBars.Ribbon;
using DevExpress.XtraEditors;
using DevExpress.XtraScheduler;
using DevExpress.XtraScheduler.GoogleCalendar;
using Google;
using Google.Apis;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Management;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Timers; 
using System.Windows.Forms;
using System.Net.NetworkInformation;
using Microsoft.Win32;

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

            RegistryKey reg = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
            reg.SetValue("GCSync", Application.ExecutablePath.ToString());

            WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            notifyIcon1.BalloonTipText = "votre application a été réduite dans la barre d'état système";
            notifyIcon1.ShowBalloonTip(200);
            notifyIcon1.Visible = true;

            XtraMessageBoxArgs args = new XtraMessageBoxArgs();
            args.AutoCloseOptions.Delay = 3000;
            args.Caption = "Auto-close message";
            args.Text = "This message closes automatically after 3 seconds.";
            args.Buttons = new DialogResult[] { DialogResult.OK, DialogResult.Cancel };






        }

        private void bbiSynchronize_ItemClick_1(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
           // schedulerStorage.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("gId", "dossierMarcherDataSet"));
           // schedulerStorage.Appointments.CustomFieldMappings.Add(new AppointmentCustomFieldMapping("etag", "database_ETag_field_name"));
        }
      

        private void barButtonItem1_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {


            //  CalendarService myService = new CalendarService("your calendar name");
            //  myService.setUserCredentials(username, password);

            //  CalendarEntry calendar;



            if (ss() == false  && testcn() == false)
            {
                XtraMessageBoxArgs args = new XtraMessageBoxArgs();
                args.AutoCloseOptions.Delay = 2000;
                args.Caption = "Auto-close message";
                args.Text = "access internet connection deny";
                args.Buttons = new DialogResult[] { DialogResult.OK, DialogResult.Cancel };
                //     MessageBox.Show(ex.Message);
                XtraMessageBox.Show(args).ToString();

            }
            else
            {
                foreach (var item in LoadEvents(service, "ohsaurskmltarugmcflkp758i4@group.calendar.google.com", CancellationToken.None))
                {
                    //CalendarService service = new CalendarService();
                    string calendarId = "ohsaurskmltarugmcflkp758i4@group.calendar.google.com";
                    service.Events.Delete(calendarId, item.Id).Execute();

                    // MessageBox.Show(item.Id);
                    //rest of the fields

                }

                this.appointmentsTableAdapter.Fill(this.dossierMarcherDataSet.Appointments);

                this.gcSyncComponent.CalendarId = "ohsaurskmltarugmcflkp758i4@group.calendar.google.com";
                this.gcSyncComponent.Synchronize();
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
           //SchedulerControl1.AllowAppointmentCopy;
           // e.Appointment.CustomFields.Item(0) = Nothing;
           // e.Appointment.CustomFields.Item(1) = Nothing;
        }



        private void barButtonItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {


            //XtraMessageBoxArgs args = new XtraMessageBoxArgs();
            //args.AutoCloseOptions.Delay = 3000;
            //args.Caption = "Auto-close message";
            //args.Text = "This message closes automatically after 3 seconds.";
            //args.Buttons = new DialogResult[] { DialogResult.OK, DialogResult.Cancel };
            //try
            //{
            //    dossierMarcherDataSet.Appointments.Clear();
            //    appointmentsTableAdapter.Update(this.dossierMarcherDataSet.Appointments);
            //    this.gcSyncComponent.CalendarId = "ohsaurskmltarugmcflkp758i4@group.calendar.google.com";
            //    this.gcSyncComponent.Synchronize();
            //    appointmentsTableAdapter.Update(this.dossierMarcherDataSet.Appointments);
            //}
            //catch (Exception ex)
            //{
            //    XtraMessageBox.Show(args + ex.Message).ToString();
            //}


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

            //var service = new CalendarService(new BaseClientService.Initializer()
            //{
            //    HttpClientInitializer = credential,
            //    ApplicationName = "Other client 1"
            //});

            //AuthenticateOauth("TmqFTGACK6NBurNip0feaK1y", "ohsaurskmltarugmcflkp758i4@group.calendar.google.com");
            //MessageBox.Show(LoadEvents(service, "ohsaurskmltarugmcflkp758i4@group.calendar.google.com", CancellationToken.None).Select(x=>x.Start).ToString());




            //CalendarService service = new CalendarService();
            //string calendarId = "ohsaurskmltarugmcflkp758i4@group.calendar.google.com";
            //service.Events.Delete(calendarId, eventId).Execute();


           // CalendarService service = new CalendarService();

            foreach (var item in LoadEvents(service, "ohsaurskmltarugmcflkp758i4@group.calendar.google.com", CancellationToken.None))
            {
                //CalendarService service = new CalendarService();
                string calendarId = "ohsaurskmltarugmcflkp758i4@group.calendar.google.com";
                service.Events.Delete(calendarId, item.Id).Execute();

               // MessageBox.Show(item.Id);
                //rest of the fields

            }

        }




            private void gcSyncComponent_ConflictDetected_1(object sender, ConflictDetectedEventArgs e)
        {
            e.GoogleEventIsValid = true;

        }

        private void gcSyncComponent_FilterAppointments(object sender, FilterAppointmentsEventArgs e)
        {
            //I test that code but it just stopping add the appointments to google calendar

            //if (e.Appointment != null && e.Appointment.Description.Contains("*"))
            //    e.Cancel = true;


           


            //if (Program.sql_con.State == ConnectionState.Closed) Program.sql_con.Open();
            //Program.sql_cmd = new SqlCommand("select  *  from  Appointments ", Program.sql_con);
            //Program.db = Program.sql_cmd.ExecuteReader();
            //while (Program.db.HasRows)
            //{

            //    if (Program.sql_con.State == ConnectionState.Closed) Program.sql_con.Open();
            //    Program.sql_cmd = new SqlCommand("select  UniqueID  from  Appointments where Description like '%*%'", Program.sql_con);
            //    string id = "";
            //    Program.db = Program.sql_cmd.ExecuteReader();


            //    id = Program.db[0].ToString();


            //    string sql = "UPDATE Appointments SET Description = concat('*',Description) where UniqueID != @id ";
            //    Program.sql_cmd = new SqlCommand(sql, Program.sql_con);
            //    Program.sql_cmd.Parameters.AddWithValue("@id", int.Parse(id));


            //    if (Program.sql_con.State == ConnectionState.Closed) Program.sql_con.Open();
            //    Program.sql_cmd.ExecuteNonQuery();
            //    Program.sql_con.Close();


               

            }

        //holy scrypye

        public IEnumerable<Event> LoadEvents(CalendarService service, string calendarId, CancellationToken cancellationToken)
        {
           
                List<Event> result = new List<Event>();
                String pageToken = null;
                do
                {
                    EventsResource.ListRequest listRequest = service.Events.List(calendarId);
                    listRequest.PageToken = pageToken;
                    listRequest.ShowDeleted = false;
                    Events events = null;
                    try
                    {
                        if (cancellationToken.IsCancellationRequested)
                            return result;
                        events = listRequest.Execute();
                        result.AddRange(events.Items);
                        pageToken = events.NextPageToken;


                    }
                    catch (TaskCanceledException)
                    {
                  
                    //reload function
                  //  sync();
                }
                    catch (GoogleApiException apiException)
                    {
                        //!!!process errors!!!
                        break;
                    }
                } while (pageToken != null);
                return result;


            
        }

        public void sync()
        {
            //holy scrypte

            try
            {



                foreach (var item in LoadEvents(service, "ohsaurskmltarugmcflkp758i4@group.calendar.google.com", CancellationToken.None))
                {
                    //CalendarService service = new CalendarService();
                    string calendarId = "ohsaurskmltarugmcflkp758i4@group.calendar.google.com";
                    service.Events.Delete(calendarId, item.Id).Execute();

                    MessageBox.Show(item.Id);
                    //rest of the fields
                }
            }
            catch (Exception ex)
            {

                XtraMessageBoxArgs args = new XtraMessageBoxArgs();
                args.AutoCloseOptions.Delay = 3000;
                args.Caption = "Auto-close message";
                args.Text = ex.Message;
                args.Buttons = new DialogResult[] { DialogResult.OK, DialogResult.Cancel };
                //     MessageBox.Show(ex.Message);
                XtraMessageBox.Show(args).ToString();
                sync();
            }


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

                //this.appointmentsTableAdapter.Fill(this.dossierMarcherDataSet.Appointments);

                //this.gcSyncComponent.CalendarId = "ohsaurskmltarugmcflkp758i4@group.calendar.google.com";
                //this.gcSyncComponent.Synchronize();





                //    this.gcSyncComponent.CalendarId = "652ilqp8u3ikeur3ac2uqd1t8s@group.calendar.google.com";
                //  this.gcSyncComponent.Synchronize();
            }

            catch (Exception ex)
            {

                XtraMessageBoxArgs args = new XtraMessageBoxArgs();
                args.AutoCloseOptions.Delay = 3000;
                args.Caption = "Auto-close message";
                args.Text = ex.Message;
                args.Buttons = new DialogResult[] { DialogResult.OK, DialogResult.Cancel };
                //     MessageBox.Show(ex.Message);
                XtraMessageBox.Show(args).ToString();
                //reload function
                sync();

            }



            //holy scrypte
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (ss() == false &&  testcn() == false)
            {
                XtraMessageBoxArgs args = new XtraMessageBoxArgs();
                args.AutoCloseOptions.Delay = 2000;
                args.Caption = "Auto-close message";
                args.Text = "access internet connection deny";
                args.Buttons = new DialogResult[] { DialogResult.OK, DialogResult.Cancel };
                //     MessageBox.Show(ex.Message);
                XtraMessageBox.Show(args).ToString();

            }
            else
            {
                foreach (var item in LoadEvents(service, "ohsaurskmltarugmcflkp758i4@group.calendar.google.com", CancellationToken.None))
                {
                    //CalendarService service = new CalendarService();
                    string calendarId = "ohsaurskmltarugmcflkp758i4@group.calendar.google.com";
                    service.Events.Delete(calendarId, item.Id).Execute();

                    // MessageBox.Show(item.Id);
                    //rest of the fields

                }

                this.appointmentsTableAdapter.Fill(this.dossierMarcherDataSet.Appointments);

                this.gcSyncComponent.CalendarId = "ohsaurskmltarugmcflkp758i4@group.calendar.google.com";
                this.gcSyncComponent.Synchronize();
            }

        }

        private void barButtonItem3_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {


        }

        public  bool testcn()
        {
            using (SqlConnection connection = new SqlConnection(@"server =192.168.100.92;database = dossierMarcher ; user id = log1; password=P@ssword1965** ;MultipleActiveResultSets = True;"))
            {
                try
                {
                    connection.Open();
                    return true;
                }
                catch (SqlException)
                {
                    return false;
                }
            }

        }



        public bool ss()
        {
            try
            {
                Ping myPing = new Ping();
                String host = "google.com";
                byte[] buffer = new byte[32];
                int timeout = 1000;
                PingOptions pingOptions = new PingOptions();
                PingReply reply = myPing.Send(host, timeout, buffer, pingOptions);
                return (reply.Status == IPStatus.Success);
            }
            catch (Exception)
            {
                return false;
            }
        }


        //holy scrypye


        

        private void barButtonItem4_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.WindowState = FormWindowState.Normal;
            if (this.WindowState == FormWindowState.Normal)
            {
                this.ShowInTaskbar = true;
                notifyIcon1.Visible = false;
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                WindowState = FormWindowState.Minimized;
                this.ShowInTaskbar = false;
                notifyIcon1.BalloonTipText = "votre application a été réduite dans la barre d'état système";
                notifyIcon1.ShowBalloonTip(1000);
                notifyIcon1.Visible = true;


            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void notifyIcon1_BalloonTipClosed(object sender, EventArgs e)
        {
           
        }

        private void notifyIcon1_BalloonTipClicked(object sender, EventArgs e)
        {
           
        }

        private void Form1_Activated(object sender, EventArgs e)
        {
          
        }
        int counter = 0;

        private void mytimer_Tick(object sender, EventArgs e)
        {
            counter++;
            if (counter == 10)  //or whatever your limit is
            {
                if (ss() == false && testcn() == false)
                {
                    XtraMessageBoxArgs args = new XtraMessageBoxArgs();
                    args.AutoCloseOptions.Delay = 2000;
                    args.Caption = "Auto-close message";
                    args.Text = "access internet connection deny";
                    args.Buttons = new DialogResult[] { DialogResult.OK, DialogResult.Cancel };
                    //     MessageBox.Show(ex.Message);
                    XtraMessageBox.Show(args).ToString();

                }
                else
                {
                    foreach (var item in LoadEvents(service, "ohsaurskmltarugmcflkp758i4@group.calendar.google.com", CancellationToken.None))
                    {
                        //CalendarService service = new CalendarService();
                        string calendarId = "ohsaurskmltarugmcflkp758i4@group.calendar.google.com";
                        service.Events.Delete(calendarId, item.Id).Execute();

                        // MessageBox.Show(item.Id);
                        //rest of the fields

                    }

                    this.appointmentsTableAdapter.Fill(this.dossierMarcherDataSet.Appointments);

                    this.gcSyncComponent.CalendarId = "ohsaurskmltarugmcflkp758i4@group.calendar.google.com";
                    this.gcSyncComponent.Synchronize();
                }

                mytimer.Stop();
            }




            //mytimer.Stop();
            // mytimer.Enabled = false;
        }
    }
}