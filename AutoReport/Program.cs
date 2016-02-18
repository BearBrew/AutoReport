using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using Rally.RestApi;
using Rally.RestApi.Json;
using Rally.RestApi.Response;
using System.Configuration;
using System.Data.SqlClient;

namespace AutoReport
{
    class Program
    {

        // Global level variables
        /// <summary>
        /// Reference to the RallyRestApi object
        /// </summary>
        public static RallyRestApi RallyAPI;
        /// <summary>
        /// Holds on to the full list of all Projects for all Initiatives
        /// </summary>
        public static List<Project> BasicProjectList;
        /// <summary>
        /// Holds on to the full list of all User Stories for all Projects
        /// </summary>
        public static List<UserStory> BasicStoryList;
        /// <summary>
        /// Holds on to the full list of all Defects for all User Stories
        /// </summary>
        public static List<Defect> BasicDefectList;
        /// <summary>
        /// Holds on to full list of Projects along with stories and tasks included
        /// </summary>
        public static List<Project> CompleteProjectList;
        /// <summary>
        /// Holds the final list of all projects and calculated totals
        /// </summary>
        public static List<Project> AllProjectInfo;
        /// <summary>
        /// Configuration item for Rally UserID
        /// </summary>
        public static string ConfigRallyUID;
        /// <summary>
        /// Configuration item for Rally Password
        /// </summary>
        public static string ConfigRallyPWD;
        /// <summary>
        /// Configuration item for UNC Logfile Location
        /// </summary>
        public static string ConfigLogPath;
        /// <summary>
        /// Configuration item for end-of-line indicate in project status update
        /// </summary>
        public static string ConfigStatusEOL;
        /// <summary>
        /// Configuration item for UNC Report Location
        /// </summary>
        public static string ConfigReportPath;
        /// <summary>
        /// Configuration item to enable full debug
        /// </summary>
        public static bool ConfigDebug;
        /// <summary>
        /// Full path to the output text file report
        /// </summary>
        public static string ReportFile;
        /// <summary>
        /// Full path to the logfile
        /// </summary>
        public static string LogFile;
        /// <summary>
        /// Name of the SQL Server the reporting database is running on
        /// </summary>
        public static string ConfigDBServer;
        /// <summary>
        /// Name of the SQL Server database to use for reporting
        /// </summary>
        public static string ConfigDBName;
        /// <summary>
        /// UID for the database connection
        /// </summary>
        public static string ConfigDBUID;
        /// <summary>
        /// Password for database connection
        /// </summary>
        public static string ConfigDBPWD;
        /// <summary>
        /// Reporting "Mode" we are running for
        /// </summary>
        public static OperateMode OperatingMode;
        /// <summary>
        /// Day portion of reporting date
        /// </summary>
        public static int ReportDayNum;
        /// <summary>
        /// Week number of year of reporting date
        /// </summary>
        public static int ReportWeekNum;
        /// <summary>
        /// Month portion of reporting date
        /// </summary>
        public static int ReportMonth;
        /// <summary>
        /// Quarter number of reporting date
        /// </summary>
        public static int ReportQuarter;
        /// <summary>
        /// Year portion of reporting date
        /// </summary>
        public static int ReportYear;
        /// <summary>
        /// Full Date/Time of reporting date
        /// </summary>
        public static DateTime ReportingDay;

        // Global level structures
        public struct Initiative
        {
            /// <summary>
            /// Name of the Owner of the Initiative
            /// </summary>
            public string Owner;
            /// <summary>
            /// Formatted ID Number of the item
            /// </summary>
            public string FormattedID;
            /// <summary>
            /// Long Description of the Initiative
            /// </summary>
            public string Description;
            /// <summary>
            /// Name of the Initiative
            /// </summary>
            public string Name;
            /// <summary>
            /// Planned start date
            /// </summary>
            public System.DateTime PlannedStartDate;
            /// <summary>
            /// Planned End Date
            /// </summary>
            public System.DateTime PlannedEndDate;
            /// <summary>
            /// Actual End Date
            /// </summary>
            public System.DateTime ActualEndDate;
            /// <summary>
            /// Current State of the Initiative
            /// </summary>
            public string State;
        }

        public struct Project
        {
            /// <summary>
            /// Formatted ID of the Project
            /// </summary>
            public string FormattedID;
            /// <summary>
            /// Name of the Project
            /// </summary>
            public string Name;
            /// <summary>
            /// Name of the Owner of the Project
            /// </summary>
            public string Owner;
            /// <summary>
            /// Name of the person responsible for providing status updates
            /// </summary>
            public string UpdateOwner;
            /// <summary>
            /// Planned Start date of the Project
            /// </summary>
            public System.DateTime PlannedStartDate;
            /// <summary>
            /// Planned End date of the Project
            /// </summary>
            public System.DateTime PlannedEndDate;
            /// <summary>
            /// Actual End date of the Project
            /// </summary>
            public System.DateTime ActualEndDate;
            /// <summary>
            /// Last Date/Time the project was updated
            /// </summary>
            public System.DateTime LastUpdateDate;
            /// <summary>
            /// Running status updates of the Project
            /// </summary>
            public string StatusUpdate;
            /// <summary>
            /// Total value of the associated opportunity
            /// </summary>
            public long OpportunityAmount;
            /// <summary>
            /// State of the Project
            /// </summary>
            public string State;
            /// <summary>
            /// Long description of the Project
            /// </summary>
            public string Description;
            /// <summary>
            /// Reason the Project was 'Revoked'
            /// </summary>
            public string RevokedReason;
            /// <summary>
            /// Primary Stakeholder
            /// </summary>
            public string StakeHolder;
            /// <summary>
            /// Parent Initiative for the Project
            /// </summary>
            public string Initiative;
            /// <summary>
            /// Is the project considered a "Top 5"
            /// </summary>
            public bool Expedite;
            /// <summary>
            /// List of Epic Structures for Project
            /// </summary>
            public List<Epic> Epics;
            /// <summary>
            /// List of Story Structures for Project
            /// </summary>
            public List<UserStory> UserStories;
            /// <summary>
            /// List of Defect Structures for Project
            /// </summary>
            public List<Defect> Defects;
            /// <summary>
            /// Total for all 'Estimates' of Stories
            /// </summary>
            public decimal StoryEstimate;
            /// <summary>
            /// Total for all 'ToDo' of Stories
            /// </summary>
            public decimal StoryToDo;
            /// <summary>
            /// Total for all 'Actuals' of Stories
            /// </summary>
            public decimal StoryActual;
            /// <summary>
            /// Total for all 'Estimates' of Defects
            /// </summary>
            public decimal DefectEstimate;
            /// <summary>
            /// Total for all 'ToDo' of Defects
            /// </summary>
            public decimal DefectToDo;
            /// <summary>
            /// Total for all 'Actuals' of Defects
            /// </summary>
            public decimal DefectActual;
            /// <summary>
            /// Which 'Project' or 'Team' is this for?
            /// </summary>
            public string ScrumTeam;
        }

        public struct Epic
        {
            /// <summary>
            /// Formatted ID of the Epic
            /// </summary>
            public string FormattedID;
            /// <summary>
            /// Name of the Epic
            /// </summary>
            public string Name;
            /// <summary>
            /// Name of the Owner of the Epic
            /// </summary>
            public string Owner;
            /// <summary>
            /// Planned Start date of the Epic
            /// </summary>
            public System.DateTime PlannedStartDate;
            /// <summary>
            /// Planned End date of the Epic
            /// </summary>
            public System.DateTime PlannedEndDate;
            /// <summary>
            /// Actual End date of the Epic
            /// </summary>
            public System.DateTime ActualEndDate;
            /// <summary>
            /// State of the Epic
            /// </summary>
            public string State;
            /// <summary>
            /// Long description of the Epic
            /// </summary>
            public string Description;
            /// <summary>
            /// Release this Epic is part of
            /// </summary>
            public string Release;
            /// <summary>
            /// Parent Project this Epic is part of
            /// </summary>
            public string ParentProject;
            /// <summary>
            /// Total for all 'Estimates' of Stories
            /// </summary>
            public decimal StoryEstimate;
            /// <summary>
            /// Total for all 'ToDo' of Stories
            /// </summary>
            public decimal StoryToDo;
            /// <summary>
            /// Total for all 'Actuals' of Stories
            /// </summary>
            public decimal StoryActual;
            /// <summary>
            /// Total for all 'Estimates' of Defects
            /// </summary>
            public decimal DefectEstimate;
            /// <summary>
            /// Total for all 'ToDo' of Defects
            /// </summary>
            public decimal DefectToDo;
            /// <summary>
            /// Total for all 'Actuals' of Defects
            /// </summary>
            public decimal DefectActual;
        }

        public struct UserStory
        {
            /// <summary>
            /// Formatted ID of the Story
            /// </summary>
            public string FormattedID;
            /// <summary>
            /// Name of the Story
            /// </summary>
            public string Name;
            /// <summary>
            /// Name of the Owner of the Story
            /// </summary>
            public string Owner;
            /// <summary>
            /// Planned estimate for the Story
            /// </summary>
            public decimal PlanEstimate;
            /// <summary>
            /// State of the Story
            /// </summary>
            public string State;
            /// <summary>
            /// Long description of the Story
            /// </summary>
            public string Description;
            /// <summary>
            /// Release this story is targeted for
            /// </summary>
            public string Release;
            /// <summary>
            /// Iteration this story is targeted for
            /// </summary>
            public string Iteration;
            /// <summary>
            /// Count of how many children there are with a nested User Story
            /// </summary>
            public Int32 Children;
            /// <summary>
            /// Parent Project this Story is part of
            /// </summary>
            public string ParentProject;
            /// <summary>
            /// List of Task Structures for this Story
            /// </summary>
            public List<Task> Tasks;
            /// <summary>
            /// Is the Story blocked?
            /// </summary>
            public bool Blocked;
            /// <summary>
            /// Name of the Epic this Story is part of
            /// </summary>
            public string ParentEpic;
            /// <summary>
            /// Total for all 'Estimates' of Stories
            /// </summary>
            public decimal StoryEstimate;
            /// <summary>
            /// Total for all 'ToDo' of Stories
            /// </summary>
            public decimal StoryToDo;
            /// <summary>
            /// Total for all 'Actuals' of Stories
            /// </summary>
            public decimal StoryActual;
        }

        public struct Task
        {
            /// <summary>
            /// Formatted ID of the Task
            /// </summary>
            public string FormattedID;
            /// <summary>
            /// Name of the Task
            /// </summary>
            public string Name;
            /// <summary>
            /// Name of the Owner of the Task
            /// </summary>
            public string Owner;
            /// <summary>
            /// Estimated Hours for the Task
            /// </summary>
            public decimal Estimate;
            /// <summary>
            /// Actual Hours for the Task
            /// </summary>
            public decimal Actual;
            /// <summary>
            /// Long description of the Task
            /// </summary>
            public string Description;
            /// <summary>
            /// Remaining Hours for the Task
            /// </summary>
            public decimal ToDo;
            /// <summary>
            /// Current Status for the Task
            /// </summary>
            public string State;
            /// <summary>
            /// Last Update Date/Time for the Task
            /// </summary>
            public DateTime LastUpdate;
        }

        public struct Defect
        {
            /// <summary>
            /// Formatted ID of the Defect
            /// </summary>
            public string FormattedID;
            /// <summary>
            /// Name of the Defect
            /// </summary>
            public string Name;
            /// <summary>
            /// Long description of the Defect
            /// </summary>
            public string Description;
            /// <summary>
            /// Iteration this Defect is targeted for
            /// </summary>
            public string Iteration;
            /// <summary>
            /// Name of the Owner of the Defect
            /// </summary>
            public string Owner;
            /// <summary>
            /// Release this Defect is targeted for
            /// </summary>
            public string Release;
            /// <summary>
            /// State of the Defect
            /// </summary>
            public string State;
            /// <summary>
            /// Story that the Defect belongs to
            /// </summary>
            public string ParentStory;
            /// <summary>
            /// List of Task Structures for the Defect
            /// </summary>
            public List<Task> Tasks;
            /// <summary>
            /// Total for all 'Estimates' of Defects
            /// </summary>
            public decimal DefectEstimate;
            /// <summary>
            /// Total for all 'ToDo' of Defects
            /// </summary>
            public decimal DefectToDo;
            /// <summary>
            /// Total for all 'Actuals' of Defects
            /// </summary>
            public decimal DefectActual;
        }

        public enum ProjectTotal { Estimate = 1, ToDo, Actual };
        public enum OperateMode { Daily = 1, Weekly, Monthly, Quarterly, Annual };

        static void Main(string[] args)
        {

            #region StartUp Section
            // Initialize all variables
            ConfigRallyUID = "";
            ConfigRallyPWD = "";
            ConfigLogPath = "";
            ConfigStatusEOL = "";
            ConfigReportPath = "";
            ReportFile = "";
            LogFile = "";
            ConfigDBServer = "";
            ConfigDBName = "";
            ConfigDBUID = "";
            ConfigDBPWD = "";
            ReportDayNum = 0;
            ReportWeekNum = 0;
            ReportMonth = 0;
            ReportQuarter = 0;
            ReportYear = 0;

            // Get the configuration information from config file
            if (!GetConfigSettings())
            {
                // If we can't get the configuration settings, we can't even log anything, so just terminate
                Environment.Exit(-1);
            }

            // Check for any commandline arguments.  If there are not any, assume a "Daily" operating mode and set
            // the report date to yesterday (we don't want to report on today)
            if (args.Length != 0)
            {
                GetCommandArgs();
            }
            else
            {
                OperatingMode = OperateMode.Daily;
                ReportingDay = DateTime.Today.AddDays(-1);
                ReportDayNum = ReportingDay.Day;
                ReportMonth = ReportingDay.Month;
                ReportYear = ReportingDay.Year;
            }
            #endregion StartUp Section

            // Log the start of processing
            LogOutput("Started processing at " + DateTime.Now.ToLongTimeString() + " on " + DateTime.Now.ToLongDateString(), "Main", false);
            DateTime dtStartTime = DateTime.Now;

            // Log the operating mode
            switch (OperatingMode)
            {
                case OperateMode.Daily:
                    LogOutput("Operating in Daily Mode...Processing for Day " + ReportingDay.ToString("dd-MMM-yyyy"), "Main", false);
                    break;
                case OperateMode.Monthly:
                    LogOutput("Operating in Monthly Mode...Processing for Month " + ReportingDay.ToString("dd-MMM-yyyy"), "Main", false);
                    break;
                case OperateMode.Quarterly:
                    LogOutput("Operating in Quarterly mode...Processing for Quarter Q" + ReportQuarter.ToString() + "Y" + ReportYear.ToString(), "Main", false);
                    break;
                case OperateMode.Annual:
                    LogOutput("Operating in Annual mode...Processing for year " + ReportYear.ToString(), "Main", false);
                    break;
                case OperateMode.Weekly:
                    LogOutput("Operating in Weekly mode...Processing for Day " + ReportingDay.ToString("dd-MMM-yyyy"), "Main", false);
                    break;
                default:
                    LogOutput("Unknown Operating mode...assuming Daily...", "Main", false);
                    break;
            }

            #region Gather from Rally
            // Create the Rally API object
            LogOutput("Creating reference to RallyAPI...", "Main", true);
            RallyAPI = new RallyRestApi();

            // Login to Rally
            LogOutput("Starting connection to Rally...", "Main", true);
            try
            {
                RallyAPI.Authenticate(ConfigRallyUID, ConfigRallyPWD);
                LogOutput("Response from RallyAPI.Authenticate: " + RallyAPI.AuthenticationState.ToString(), "Main", true);
                if (RallyAPI.AuthenticationState.ToString() != "Authenticated")
                {
                    // We did not actually connect
                    LogOutput("Unable to connect to Rally and establish session.  Application will terminate.", "Main", false);
                    Environment.Exit(-1);
                }
                else
                {
                    if (RallyAPI.ConnectionInfo.UserName == null)
                    {
                        LogOutput("RallyAPI.ConnectionInfo: " + RallyAPI.ConnectionInfo.ToString(), "Main", false);
                        LogOutput("Unable to authenticate with Rally.  Application will terminate.", "Main", false);
                        Environment.Exit(-1);
                    }
                    else
                    {
                        LogOutput("Connected to Rally as user " + RallyAPI.ConnectionInfo.UserName, "Main", true);
                    }
                }
            }
            catch (Exception ex)
            {
                LogOutput("Error Connecting to Rally: " + ex.Message, "Main", false);
                LogOutput("Rally Authentication State: " + RallyAPI.AuthenticationState.ToString() +
                    "Rally Connection Info: " + RallyAPI.ConnectionInfo.ToString(), "Main", false);
            }

            // Grab the active Initiatives we want to report on
            LogOutput("Getting all Initiatives...", "Main", false);
            List<Initiative> InitiativeList = new List<Initiative>();
            LogOutput("Calling 'GetInitiativeList'...", "Main", true);
            InitiativeList = GetInitiativeList();
            LogOutput("Done with 'GetInitiativeList'", "Main", true);
            if (InitiativeList.Count == 0)
            {
                // Could not get the Initiatives...or nothing to report on, so stop
                LogOutput("Unable to open Initiative list or no Initiatives to report on.  Application will terminate.", "Main", false);
                InitiativeList.Clear();
                RallyAPI.Logout();      // Disconnect from Rally
                Environment.Exit(-1);   // End Program
            }
            LogOutput("Retrieved " + InitiativeList.Count + " Initiatives to report on", "Main", false);

            // Now iterate through the initiatives and get all the Features or "Projects"
            LogOutput("Getting all Projects for the Initiatives...", "Main", false);
            BasicProjectList = new List<Project>();
            LogOutput("Looping for each Initiative...", "Main", true);
            foreach (Initiative init in InitiativeList)
            {
                // Get the project list for the current initiative ONLY
                List<Project> ProjectList = new List<Project>();
                LogOutput("Calling 'GetProjectsForInitiative' with " + init.Name.Trim() + "...", "Main", true);
                ProjectList = GetProjectsForInitiative(init.Name.Trim());
                LogOutput("Done with 'GetProjectsForInitiative' for " + init.Name.Trim(), "Main", true);

                // Append this list to the FULL list
                LogOutput("Appending " + ProjectList.Count + " projects to object 'BasicProjectList'", "Main", true);
                BasicProjectList.AddRange(ProjectList);
            }
            LogOutput("Retrieved " + BasicProjectList.Count + " Projects total", "Main", false);

            // We need to loop through the project list now and for each project
            // we need to get all the epics.  Then with each epic, we recursively
            // get all user stories
            LogOutput("Getting all User Stories for all projects...", "Main", false);
            // Initialize a new list of projects.  This will become the full list including stories
            // and defects
            CompleteProjectList = new List<Project>();
            LogOutput("Looping for each Project...", "Main", true);
            foreach (Project proj in BasicProjectList)
            {
                // Get all the epics for this project
                LogOutput("~+~+~+~+~+~+~+~+~+~+~+~+~+~+~+~+~+~+", "Main", true);
                LogOutput("Calling 'GetEpicsForProject' with " + proj.Name.Trim() + "...", "Main", true);
                List<Epic> EpicList = new List<Epic>();
                EpicList = GetEpicsForProject(proj.FormattedID.Trim());
                LogOutput("Done with 'GetEpicsForProject' for " + proj.Name.Trim(), "Main", true);

                // Now go through each of the Epics for the current project
                // and recurse through to get all final-level user stories
                LogOutput("Getting all User Stories for " + proj.Name + "...", "Main", true);
                BasicStoryList = new List<UserStory>();
                LogOutput("Looping for each Epic...", "Main", true);
                foreach (Epic epic in EpicList)
                {
                    List<UserStory> StoryList = new List<UserStory>();
                    LogOutput("Calling 'GetUserStoriesPerParent' with " + epic.FormattedID.Trim() + " as Root Parent...", "Main", true);
                    StoryList = GetUserStoriesPerParent(epic.FormattedID.Trim(), epic.Name.Trim(), true);
                    LogOutput("Done with 'GetUserStoriesPerParent' for " + epic.FormattedID.Trim(), "Main", true);
                    BasicStoryList.AddRange(StoryList);
                }
                LogOutput("Retrieved " + BasicStoryList.Count + " User Stories for " + proj.Name, "Main", false);

                // Get any defects there may be for the User Stories
                LogOutput("Getting all Defects for " + proj.Name + "...", "Main", false);
                BasicDefectList = new List<Defect>();
                LogOutput("Looping for each Story...", "Main", true);
                foreach (UserStory story in BasicStoryList)
                {
                    List<Defect> DefectList = new List<Defect>();
                    // Defects will always be attached to a User Story
                    LogOutput("Calling 'GetDefectsForStory' with " + story.Name + "...", "Main", true);
                    DefectList = GetDefectsForStory(story);
                    LogOutput("Done with 'GetDefectsForStory' for " + story.Name, "Main", true);
                    // If there are any defects, add them to the list
                    if (DefectList.Count > 0)
                    {
                        LogOutput("Appending " + DefectList.Count + " defects to object 'BasicDefectList'", "Main", true);
                        BasicDefectList.AddRange(DefectList);
                    }
                }

                // At this point we have the FULL list of User Stories/Defects for the current
                // project.  We now create a "new" project with the same properties, but this time
                // we are able to store the User Stories and Defects.
                LogOutput("Creating new project object and populating with full information...", "Main", true);
                Project newproject = new Project();
                newproject.Description = proj.Description;
                newproject.Expedite = proj.Expedite;
                newproject.FormattedID = proj.FormattedID;
                newproject.Initiative = proj.Initiative;
                newproject.Name = proj.Name;
                newproject.OpportunityAmount = proj.OpportunityAmount;
                newproject.Owner = proj.Owner;
                newproject.UpdateOwner = proj.UpdateOwner;
                newproject.StakeHolder = proj.StakeHolder;
                newproject.PlannedEndDate = proj.PlannedEndDate;
                newproject.PlannedStartDate = proj.PlannedStartDate;
                newproject.RevokedReason = proj.RevokedReason;
                newproject.State = proj.State;
                newproject.StatusUpdate = proj.StatusUpdate;
                newproject.UserStories = BasicStoryList;
                newproject.Defects = BasicDefectList;
                newproject.Epics = EpicList;
                LogOutput("Appending new project object to object 'CompleteProjectList'", "Main", true);
                CompleteProjectList.Add(newproject);
            }
            LogOutput("Done looping through all projects", "Main", false);
            LogOutput("Appending new project object to object 'CompleteProjectList'", "Main", true);
            #endregion Gather from Rally

            #region Calculation Section
            // We now have a full list of all projects with all complete information so
            // at this point we can calculate the Actuals for each project based on the reporting mode we are operating in
            switch (OperatingMode)
            {
                case OperateMode.Daily:
                    // This mode runs every day and simply keeps running total of time with daily inserts for each project
                    AllProjectInfo = new List<Project>();
                    LogOutput("Calling 'CalculateDailyTotals' with " + CompleteProjectList.Count + " complete projects...", "Main", true);
                    AllProjectInfo = CalculateDailyTotals(CompleteProjectList, ReportingDay);
                    LogOutput("Done with 'CalculateDailyTotals'", "Main", true);

                    // Now create the final report
                    LogOutput("Calling 'CreateDailyReport'...", "Main", true);
                    CreateDailyReport(AllProjectInfo, ReportingDay);
                    LogOutput("Done with 'CreateDailyReport'...", "Main", true);
                    break;
                case OperateMode.Monthly:
                    // This mode runs each month and creates very high-level summary info
                    AllProjectInfo = new List<Project>();
                    LogOutput("Calling 'CalculateMonthlyReport' with " + CompleteProjectList.Count + " complete projects...", "Main", true);
                    AllProjectInfo = CalculateMonthlyReport(CompleteProjectList, ReportMonth);
                    LogOutput("Done with 'CalculateMonthlyReport'", "Main", true);

                    // Now create the final report
                    LogOutput("Calling 'CreateMonthlyReport'...", "Main", true);
                    CreateMonthlyReport(AllProjectInfo, ReportMonth);
                    LogOutput("Done with 'CreateMonthlyReport'...", "Main", true);
                    break;
                case OperateMode.Quarterly:
                    AllProjectInfo = new List<Project>();
                    LogOutput("Calling 'CalculateQuarterTotals' with " + CompleteProjectList.Count + " complete projects...", "Main", true);
                    AllProjectInfo = CalculateQuarterTotals(CompleteProjectList, ReportQuarter, ReportYear);
                    LogOutput("Done with 'CalculateQuarterTotals'", "Main", true);

                    // Now create the final report
                    LogOutput("Calling 'CreateQuarterReport'...", "Main", true);
                    CreateQuarterReport(AllProjectInfo, ReportYear.ToString() + 'Q' + ReportQuarter.ToString());
                    LogOutput("Done with 'CreateQuarterReport'...", "Main", true);
                    break;
                case OperateMode.Annual:
                    AllProjectInfo = new List<Project>();
                    LogOutput("Calling 'CalculateAnnualTotals' with " + CompleteProjectList.Count + " complete projects...", "Main", true);
                    AllProjectInfo = CalculateAnnualTotals(CompleteProjectList, ReportYear);
                    LogOutput("Done with 'CalculateAnnualTotals'", "Main", true);

                    // Now create the final report
                    LogOutput("Calling 'CreateAnnualReport'...", "Main", true);
                    CreateAnnualReport(AllProjectInfo, ReportYear);
                    LogOutput("Done with 'CreateAnnualReport'...", "Main", true);
                    break;
                case OperateMode.Weekly:
                    // This mode is intended to run on Sunday in order to run stats for up-to-and-including current week for projects.  The "Project Update"
                    // is also included so the single table can be queried
                    AllProjectInfo = new List<Project>();
                    LogOutput("Calling 'CalculateWeeklyTotals' with " + CompleteProjectList.Count + " complete projects...", "Main", true);
                    AllProjectInfo = CalculateWeeklyTotals(CompleteProjectList, ReportingDay);
                    LogOutput("Done with 'CalculateWeeklyTotals'", "Main", true);

                    // Now create the weekly report
                    LogOutput("Calling 'CreateWeeklyReport'...", "Main", true);
                    CreateWeeklyReport(AllProjectInfo, ReportingDay);
                    LogOutput("Done with 'CreateWeeklyReport'...", "Main", true);
                    break;
            }
            #endregion

            DateTime dtEndTime = DateTime.Now;
            string strTotSeconds = dtEndTime.Subtract(dtStartTime).TotalSeconds.ToString();
            LogOutput("Completed processing at " + DateTime.Now.ToLongTimeString() + " on " + DateTime.Now.ToLongDateString() + " - Total Processing Time: " + strTotSeconds + " seconds", "Main", false);
            Environment.Exit(0);

        }

        /// <summary>
        /// This reads the Initiatives that we do NOT want to report on from configuration
        /// settings to allow ignoring Initiatives for OCTO that we don't we want
        /// </summary>
        public static List<Initiative> GetInitiativeList()
        {

            List<Initiative> listReturn = new List<Initiative>();
            string[] IgnoreList = new string[0];

            // Full path and filename to read the Initiative list from
            LogOutput("Getting list of initiatives from 'ConfigurationManager'...", "GetInitiativeList", true);
            string strIgnoreList = ConfigurationManager.AppSettings["IgnoreList"];
            LogOutput("Configured ignore list is: " + strIgnoreList, "GetInitiativeList", true);
            System.Text.RegularExpressions.Regex myReg = new System.Text.RegularExpressions.Regex(",");
            IgnoreList = myReg.Split(strIgnoreList);
            LogOutput("Looping for each initiative to get full information...", "GetInitiativeList", false);

            // Grab all Information for the given initiative
            LogOutput("Using RallyAPI to request information for all Initiatives...", "GetInitiativeList", true);
            LogOutput("Building Rally Request...", "GetInitiativeList", true);
            Request rallyRequest = new Request("PortfolioItem/Initiative");
            // You can get the project reference for this next part by editing the main project "" and taking the URL.  We want the part after the word 'project'
            // so this is currently https://rally1.rallydev.com/#/22986814983d/detail/project/22986814983
            rallyRequest.Project = "/project/22986814983";  // This must be a reference to the project.  This project equates to "CTO Customer & Innovation Labs"
            rallyRequest.ProjectScopeDown = true;   // Specify that we want any projects under the main one
            rallyRequest.Fetch = new List<string>() { "Name", "FormattedID", "Owner", "PlannedStartDate", "PlannedEndDate", "ActualEndDate", "Description", "State" };
            rallyRequest.Query = new Query("FormattedID", Query.Operator.DoesNotEqual, "I68");  // We really don't want a filter, but you need a query string of some sort
            LogOutput("Running Rally Query request...", "GetInitiativeList", true);
            QueryResult rallyResult = RallyAPI.Query(rallyRequest);
            LogOutput("Looping through Query result...", "GetInitiativeList", true);
            foreach (var result in rallyResult.Results)
            {
                // Only process this entry if NOT in the ignore list
                if (!CheckIgnoreList(IgnoreList, RationalizeData(result["FormattedID"])))
                {
                    // Create and populate the Initiative object
                    LogOutput("Creating new Initiative object and saving...", "GetInitiativeList", true);
                    Initiative initiative = new Initiative();
                    initiative.FormattedID = RationalizeData(result["FormattedID"]);
                    initiative.Name = RationalizeData(result["Name"]);
                    initiative.Owner = RationalizeData(result["Owner"]);
                    initiative.PlannedStartDate = Convert.ToDateTime(result["PlannedStartDate"]);
                    initiative.PlannedEndDate = Convert.ToDateTime(result["PlannedEndDate"]);
                    initiative.ActualEndDate = Convert.ToDateTime(result["ActualEndDate"]);
                    initiative.Description = GetPlainTextFromHtml(RationalizeData(result["Description"]));
                    initiative.State = RationalizeData(result["State"]);
                    LogOutput("Appending new initiative object to return list...", "GetInitiativeList", true);
                    listReturn.Add(initiative);
                }
            }

            LogOutput("Completed processing all initiatives, returning list", "GetInitiativeList", true);

            return listReturn;
        }

        /// <summary>
        /// This will retrieve all projects (Rally "Features") for the passed Initiative
        /// </summary>
        /// <param name="InitiativeID">Formatted ID of the Initiative in Rally</param>
        public static List<Project> GetProjectsForInitiative(string InitiativeID)
        {

            List<Project> listReturn = new List<Project>();

            // Grab all Features under the given initiative
            LogOutput("Using RallyAPI to request project information for all Initiatives...", "GetProjectsForInitiative", true);
            LogOutput("Building Rally Request...", "GetProjectsForInitiative", true);
            Request rallyRequest = new Request("PortfolioItem/Feature");
            rallyRequest.Fetch = new List<string>() { "Name", "FormattedID", "Owner", "PlannedStartDate",
                "ActualEndDate", "c_ProjectUpdate", "c_RevokedReason", "State", "ValueScore", "Description",
                "Expedite", "c_UpdateOwner", "c_Stakeholder", "Project", "LastUpdateDate" };
            rallyRequest.Query = new Query("Parent.Name", Query.Operator.Equals, InitiativeID);
            LogOutput("Running Rally Query request...", "GetProjectsForInitiative", true);
            QueryResult rallyResult = RallyAPI.Query(rallyRequest);
            LogOutput("Looping through Query request...", "GetProjectsForInitiative", true);
            foreach (var result in rallyResult.Results)
            {
                // Create and populate the Project object
                Project proj = new Project();
                proj.Name = RationalizeData(result["Name"]);
                proj.FormattedID = RationalizeData(result["FormattedID"]);
                proj.Owner = RationalizeData(result["Owner"]);
                proj.PlannedEndDate = Convert.ToDateTime(result["PlannedEndDate"]);
                proj.PlannedStartDate = Convert.ToDateTime(result["PlannedStartDate"]);
                proj.ActualEndDate = Convert.ToDateTime(result["ActualEndDate"]);
                proj.LastUpdateDate = Convert.ToDateTime(result["LastUpdateDate"]);
                proj.StatusUpdate = GetPlainTextFromHtml(RationalizeData(result["c_ProjectUpdate"]));
                proj.RevokedReason = RationalizeData(result["c_RevokedReason"]);
                proj.OpportunityAmount = RationalizeData(result["ValueScore"]);
                proj.State = RationalizeData(result["State"]);
                proj.StakeHolder = RationalizeData(result["c_Stakeholder"]);
                proj.Description = GetPlainTextFromHtml(RationalizeData(result["Description"]));
                proj.Expedite = RationalizeData(result["Expedite"]);
                proj.UpdateOwner = RationalizeData(result["c_UpdateOwner"]);
                proj.Initiative = InitiativeID;
                proj.ScrumTeam = RationalizeData(result["Project"]);
                LogOutput("Appending new project object to return list...", "GetProjectsForInitiative", true);
                listReturn.Add(proj);
            }

            LogOutput("Completed processing all projects, returning list", "GetProjectsForInitiative", true);

            return listReturn;
        }

        /// <summary>
        /// Reads all Epics for the indicated Project
        /// </summary>
        /// <param name="ProjectID">Formatted ID of the Project (Rally Feature)</param>
        public static List<Epic> GetEpicsForProject(string ProjectID)
        {

            List<Epic> listReturn = new List<Epic>();

            // Grab all Epics under the given project
            LogOutput("Using RallyAPI to request epic information for all projects...", "GetEpicsForProject", true);
            LogOutput("Building Rally Request...", "GetEpicsForProject", true);
            Request rallyRequest = new Request("PortfolioItem/Epic");
            rallyRequest.Fetch = new List<string>() { "Name", "FormattedID", "Owner", "PlannedEndDate",
                "PlannedStartDate","ActualEndDate", "State", "Description", "Release" };
            rallyRequest.Query = new Query("Parent.FormattedID", Query.Operator.Equals, ProjectID);
            LogOutput("Running Rally Query request...", "GetEpicsForProject", true);
            QueryResult rallyResult = RallyAPI.Query(rallyRequest);
            LogOutput("Looping through Query request...", "GetEpicsForProject", true);
            foreach (var result in rallyResult.Results)
            {
                Epic epic = new Epic();
                epic.Name = RationalizeData(result["Name"]);
                epic.FormattedID = RationalizeData(result["FormattedID"]);
                epic.Owner = RationalizeData(result["Owner"]);
                epic.PlannedEndDate = Convert.ToDateTime(result["PlannedEndDate"]);
                epic.ActualEndDate = Convert.ToDateTime(result["ActualEndDate"]);
                epic.PlannedStartDate = Convert.ToDateTime(result["PlannedStartDate"]);
                epic.State = RationalizeData(result["State"]);
                epic.Description = RationalizeData(result["Description"]);
                epic.ParentProject = ProjectName(ProjectID);
                epic.Release = RationalizeData(result["Release"]);
                LogOutput("Appending new epic object to return list...", "GetEpicsForProject", true);
                listReturn.Add(epic);
            }

            LogOutput("Completed processing all epics, returning list", "GetEpicsForProject", true);

            return listReturn;

        }

        /// <summary>
        /// This reads all defects for the supplied User Story
        /// </summary>
        /// <param name="Parent">User Story to get all defects for</param>
        public static List<Defect> GetDefectsForStory(UserStory Parent)
        {

            List<Defect> listReturn = new List<Defect>();

            // Grab all Defects under the given User Story
            LogOutput("Using RallyAPI to request defect information for all stories...", "GetDefectsForStory", true);
            LogOutput("Building Rally Request...", "GetDefectsForStory", true);
            Request rallyRequest = new Request("Defect");
            rallyRequest.Fetch = new List<string>() { "Name", "FormattedID", "Description", "Iteration", "Owner", "Release", "State" };
            rallyRequest.Query = new Query("Requirement.Name", Query.Operator.Equals, Parent.Name.Trim());
            LogOutput("Running Rally Query request...", "GetDefectsForStory", true);
            QueryResult rallyResult = RallyAPI.Query(rallyRequest);
            LogOutput("Looping through Query request...", "GetDefectsForStory", true);
            foreach (var result in rallyResult.Results)
            {
                Defect defect = new Defect();
                defect.Name = RationalizeData(result["Name"]);
                defect.FormattedID = RationalizeData(result["FormattedID"]);
                defect.Description = RationalizeData(result["Description"]);
                defect.Iteration = RationalizeData(result["Iteration"]);
                defect.Owner = RationalizeData(result["Owner"]);
                defect.Release = RationalizeData(result["Release"]);
                defect.State = RationalizeData(result["State"]);
                defect.ParentStory = Parent.FormattedID.Trim();
                defect.Tasks = GetTasksForUserStory(defect.FormattedID);
                LogOutput("Appending new defect object to return list...", "GetDefectsForStory", true);
                listReturn.Add(defect);
            }

            LogOutput("Completed processing all defects, returning list", "GetDefectsForStory", true);

            return listReturn;

        }

        /// <summary>
        /// Grabs all User Stories for the indicated parent.  The parent can be an Epic or another User Story
        /// </summary>
        /// <param name="ParentID">Formatted ID of the Parent</param>
        /// <param name="ParentName">Name of the parent, specifically the Epic name</param>
        /// <param name="RootParent">Indicates if the parent is the top-level User Story or not.  If looking for sub-stories, things are handled differently</param>
        public static List<UserStory> GetUserStoriesPerParent(string ParentID, string ParentName, bool RootParent)
        {

            List<UserStory> listReturn = new List<UserStory>();

            // Grab all UserStories under the given Epic
            LogOutput("Using RallyAPI to request story information for all parents...", "GetUserStoriesPerParent", true);
            LogOutput("Building Rally Request...", "GetUserStoriesPerParent", true);
            Request rallyRequest = new Request("HierarchicalRequirement");
            rallyRequest.Fetch = new List<string>() { "Name", "FormattedID", "Release", "Iteration", "Blocked",
                "Owner", "ScheduleState", "DirectChildrenCount", "Description", "PlanEstimate" };
            // If this is the "Root" or highest order User Story, then we want to grab by the FormattedID of the actual portfolio
            // item.  If this is a subordinate, then we want to grab everything where the PARENT object is the FormattedID that 
            // we passed into the method
            if (RootParent)
            {
                rallyRequest.Query = new Query("PortfolioItem.FormattedID", Query.Operator.Equals, ParentID);
            }
            else
            {
                rallyRequest.Query = new Query("Parent.FormattedID", Query.Operator.Equals, ParentID);
            }
            LogOutput("Running Rally Query request...", "GetUserStoriesPerParent", true);
            QueryResult rallyResult = RallyAPI.Query(rallyRequest);
            LogOutput("Looping through Query request...", "GetUserStoriesPerParent", true);
            foreach (var result in rallyResult.Results)
            {
                UserStory story = new UserStory();
                story.Name = RationalizeData(result["Name"]);
                story.FormattedID = RationalizeData(result["FormattedID"]);
                story.Owner = RationalizeData(result["Owner"]);
                story.Release = RationalizeData(result["Release"]);
                story.Iteration = RationalizeData(result["Iteration"]);
                story.State = RationalizeData(result["ScheduleState"]);
                story.Children = RationalizeData(result["DirectChildrenCount"]);
                story.Description = RationalizeData(result["Description"]);
                story.ParentProject = ParentID;
                story.ParentEpic = ParentName;
                story.PlanEstimate = RationalizeData(result["PlanEstimate"], true);
                story.Tasks = GetTasksForUserStory(story.FormattedID);
                story.Blocked = RationalizeData(result["Blocked"]);
                LogOutput("Appending new story object to return list...", "GetUserStoriesPerParent", true);
                listReturn.Add(story);
                LogOutput(story.FormattedID + " " + story.Name, "GetUserStoriesPerParent", true);

                // Check for children.  If there are children, then we need to drill down
                // through all children to retrieve the full list of stories
                if (story.Children > 0)
                {
                    // Recursively get child stories until we reach the lowest level
                    listReturn.AddRange(GetUserStoriesPerParent(story.FormattedID.Trim(), "", false));
                }
            }

            LogOutput("Completed processing all stories, returning list", "GetUserStoriesPerParent", true);

            return listReturn;

        }

        /// <summary>
        /// retrieves all Tasks for the indicated User Story
        /// </summary>
        /// <param name="StoryID">Formatted ID of the User Story to get tasks for</param>
        public static List<Task> GetTasksForUserStory(string StoryID)
        {

            List<Task> listReturn = new List<Task>();

            // Grab all Tasks under the given Story
            LogOutput("Using RallyAPI to request task information for all stories...", "GetTasksForUserStory", true);
            LogOutput("Building Rally Request...", "GetTasksForUserStory", true);
            Request rallyRequest = new Request("Tasks");
            rallyRequest.Fetch = new List<string>() { "Name", "FormattedID", "Estimate", "ToDo", "Actuals", "Owner", "State", "Description", "LastUpdateDate" };
            rallyRequest.Query = new Query("WorkProduct.FormattedID", Query.Operator.Equals, StoryID);
            LogOutput("Running Rally Query request...", "GetTasksForUserStory", true);
            QueryResult rallyResult = RallyAPI.Query(rallyRequest);
            LogOutput("Looping through Query request...", "GetTasksForUserStory", true);
            foreach (var result in rallyResult.Results)
            {
                Task task = new Task();
                task.Name = RationalizeData(result["Name"]);
                task.FormattedID = RationalizeData(result["FormattedID"]);
                task.Owner = RationalizeData(result["Owner"]);
                task.Estimate = (decimal)RationalizeData(result["Estimate"], true);
                task.ToDo = (decimal)RationalizeData(result["ToDo"], true);
                task.Actual = (decimal)RationalizeData(result["Actuals"], true);
                task.State = RationalizeData(result["State"]);
                task.Description = RationalizeData(result["Description"]);
                task.LastUpdate = Convert.ToDateTime(RationalizeData(result["LastUpdateDate"]));
                LogOutput("Appending new task object to return list...", "GetTasksForUserStory", true);
                listReturn.Add(task);
            }

            LogOutput("Completed processing all tasks, returning list", "GetTasksForUserStory", true);

            return listReturn;

        }

        /// <summary>
        /// Validates data and returns "cleaned" values
        /// </summary>
        /// <param name="RadicalData">Original Data to be cleaned</param>
        public static object RationalizeData(object RadicalData)
        {

            //  If the incoming data is null, then dont do anything but return empty string
            if (RadicalData == null)
            {
                return "";
            }
            else
            {
                switch (RadicalData.GetType().Name)
                {
                    case ("DynamicJsonObject"):
                        DynamicJsonObject djo = (DynamicJsonObject)RadicalData;
                        if (djo == null)
                        {
                            return "";
                        }
                        else
                        {
                            try
                            {
                                dynamic resp = RallyAPI.GetByReference(djo["_ref"]);
                                switch ((string)djo["_type"])
                                {
                                    case ("User"):
                                        return resp["DisplayName"];
                                    case ("State"):
                                        return resp["Name"];
                                    case ("Release"):
                                        return resp["Name"];
                                    case ("Iteration"):
                                        return resp["Name"];
                                    case ("Project"):
                                        return resp["Name"];
                                    default:
                                        return "";
                                }
                            }
                            catch (Exception ex)
                            {
                                LogOutput("Error occurred getting JSon object from Rally for " + djo["_ref"] + ": " + ex.Message, "RationalizeData", true);
                                return "";
                            }
                        }
                    case ("DateTime"):
                        //
                        DateTime dtTest;
                        try
                        {
                            dtTest = Convert.ToDateTime(RadicalData);
                        }
                        catch (Exception)
                        {
                            return DateTime.MinValue;
                        }
                        if (dtTest == null || dtTest.ToString() == "" ||
                            dtTest < DateTime.MinValue ||
                            dtTest > DateTime.MaxValue)
                        {
                            return DateTime.MinValue;
                        }
                        else
                        {
                            return Convert.ToDateTime(RadicalData.ToString());
                        }
                    case ("Double"):
                        //
                        if (RadicalData == null || RadicalData.ToString() == "")
                        {
                            return 0.0;
                        }
                        else
                        {
                            return Convert.ToDouble(RadicalData.ToString());
                        }
                    case ("Byte"):
                        //
                        if (RadicalData == null || RadicalData.ToString() == "")
                        {
                            return 0;
                        }
                        else
                        {
                            return Convert.ToByte(RadicalData.ToString());
                        }
                    case ("Int32"):
                        //
                        if (RadicalData == null || RadicalData.ToString() == "")
                        {
                            return 0;
                        }
                        else
                        {
                            return Convert.ToInt32(RadicalData.ToString());
                        }
                    case ("Int64"):
                        //
                        if (RadicalData == null || RadicalData.ToString() == "")
                        {
                            return 0;
                        }
                        else
                        {
                            return Convert.ToInt64(RadicalData.ToString());
                        }
                    case ("Single"):
                        //
                        if (RadicalData == null || RadicalData.ToString() == "")
                        {
                            return 0.0;
                        }
                        else
                        {
                            return Convert.ToSingle(RadicalData.ToString());
                        }
                    case ("Decimal"):
                        //
                        if (RadicalData == null || RadicalData.ToString() == "")
                        {
                            return 0.0;
                        }
                        else
                        {
                            return Convert.ToDecimal(RadicalData.ToString());
                        }
                    case ("String"):
                        //
                        if (RadicalData == null || RadicalData.ToString() == "")
                        {
                            return "";
                        }
                        else
                        {
                            return RadicalData.ToString().Trim();
                        }
                    case ("Boolean"):
                        //
                        if (RadicalData == null || RadicalData.ToString() == "")
                        {
                            return false;
                        }
                        else if ((bool)RadicalData)
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                }

            }
            return 0;

        }/// <summary>
         /// Validates data and returns "cleaned" values
         /// </summary>
         /// <param name="RadicalData">Original Data to be cleaned</param>
         /// <param name="ReturnNumeric">Return a numeric value</param>
        public static object RationalizeData(object RadicalData, bool ReturnNumeric)
        {

            // If the incoming data is null, then dont do anything but return empty string
            // However, if the caller is expecting a number, we need to return 0 rather than empty string
            if (RadicalData == null)
            {
                if (ReturnNumeric == true)
                {
                    return 0;
                }
                else
                {
                    return "";
                }
            }
            else
            {
                switch (RadicalData.GetType().Name)
                {
                    case ("DynamicJsonObject"):
                        DynamicJsonObject djo = (DynamicJsonObject)RadicalData;
                        if (djo == null)
                        {
                            return "";
                        }
                        else
                        {
                            try
                            {
                                dynamic resp = RallyAPI.GetByReference(djo["_ref"]);
                                switch ((string)djo["_type"])
                                {
                                    case ("User"):
                                        return resp["DisplayName"];
                                    case ("State"):
                                        return resp["Name"];
                                    case ("Release"):
                                        return resp["Name"];
                                    case ("Iteration"):
                                        return resp["Name"];
                                    case ("Project"):
                                        return resp["Name"];
                                    default:
                                        return "";
                                }
                            }
                            catch (Exception ex)
                            {
                                LogOutput("Error occurred getting JSon object from Rally for " + djo["_ref"] + ": " + ex.Message, "RationalizeData", true);
                                throw;
                            }
                        }
                    case ("DateTime"):
                        //
                        DateTime dtTest;
                        try
                        {
                            dtTest = Convert.ToDateTime(RadicalData);
                        }
                        catch (Exception)
                        {
                            return DateTime.MinValue;
                        }
                        if (dtTest == null || dtTest.ToString() == "" ||
                            dtTest < DateTime.MinValue ||
                            dtTest > DateTime.MaxValue)
                        {
                            return DateTime.MinValue;
                        }
                        else
                        {
                            return Convert.ToDateTime(RadicalData.ToString());
                        }
                    case ("Double"):
                        //
                        if (RadicalData == null || RadicalData.ToString() == "")
                        {
                            return 0.0;
                        }
                        else
                        {
                            return Convert.ToDouble(RadicalData.ToString());
                        }
                    case ("Byte"):
                        //
                        if (RadicalData == null || RadicalData.ToString() == "")
                        {
                            return 0;
                        }
                        else
                        {
                            return Convert.ToByte(RadicalData.ToString());
                        }
                    case ("Int32"):
                        //
                        if (RadicalData == null || RadicalData.ToString() == "")
                        {
                            return 0;
                        }
                        else
                        {
                            return Convert.ToInt32(RadicalData.ToString());
                        }
                    case ("Int64"):
                        //
                        if (RadicalData == null || RadicalData.ToString() == "")
                        {
                            return 0;
                        }
                        else
                        {
                            return Convert.ToInt64(RadicalData.ToString());
                        }
                    case ("Single"):
                        //
                        if (RadicalData == null || RadicalData.ToString() == "")
                        {
                            return 0.0;
                        }
                        else
                        {
                            return Convert.ToSingle(RadicalData.ToString());
                        }
                    case ("Decimal"):
                        //
                        if (RadicalData == null || RadicalData.ToString() == "")
                        {
                            return 0.0;
                        }
                        else
                        {
                            return Convert.ToDecimal(RadicalData.ToString());
                        }
                    case ("String"):
                        //
                        if (RadicalData == null || RadicalData.ToString() == "")
                        {
                            return "";
                        }
                        else
                        {
                            return RadicalData.ToString().Trim();
                        }
                    case ("Boolean"):
                        //
                        if (RadicalData == null || RadicalData.ToString() == "")
                        {
                            return false;
                        }
                        else if ((bool)RadicalData)
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                }

            }

            return 0;

        }

        /// <summary>
        /// Grabs the full name of the Project based on the Formatted ID
        /// </summary>
        /// <param name="FormattedID">Formatted ID of the Project (Rally Feature)</param>
        private static string ProjectName(string FormattedID)
        {

            // Loop through the projects and compare the FormattedID numbers
            foreach (Project proj in BasicProjectList)
            {
                if (proj.FormattedID == FormattedID)
                {
                    return proj.Name;
                }
            }

            return "";

        }

        /// <summary>
        /// Sends processing information to output (screen, file, etc)
        /// </summary>
        /// <param name="Message">Text to output</param>
        /// <param name="Method">Current executing method</param>
        /// <param name="ExtendedLogInfo">Indicates if this extended Debug info</param>
        private static void LogOutput(string Message, string Method, bool ExtendedLogInfo)
        {

            // Open the StreamWriter to write to the file
            System.IO.StreamWriter logfile = new System.IO.StreamWriter(LogFile, true);

            // Check if we are in Debug mode or not
            switch (ConfigDebug)
            {
                case true:
                    // Full debug...This dumps timing information, executing module, and some message about what we are doing
                    string strOutput = System.DateTime.Now.ToString("hh:mm:ss.ffffff") + "\tCURRENT MODULE: " + Method + "\tMESSAGE: " + Message;
                    Console.WriteLine(strOutput);
                    logfile.WriteLine(strOutput);
                    break;
                default:
                    // Basic logging...just dump basic info
                    if (!ExtendedLogInfo)
                    {
                        Console.WriteLine(Message);
                        logfile.WriteLine(Message);
                    }
                    break;
            }

            logfile.Flush();
            logfile.Close();
            logfile.Dispose();

        }

        /// <summary>
        /// Creates the final Daily output report using the supplied project list
        /// </summary>
        /// <param name="projects">Project list, with calculations, to use for report</param>
        /// <param name="RptDate">Date to run report for</param>
        private static void CreateDailyReport(List<Project> projects, DateTime RptDate)
        {

            string strOutLine = "";
            bool bDBConnected = false;

            // Create the SQL Server database connection
            SqlConnection sqlDatabase = new SqlConnection("Data Source=" + ConfigDBServer +
                                                            ";Initial Catalog=" + ConfigDBName +
                                                            ";User ID=" + ConfigDBUID +
                                                            ";Password=" + ConfigDBPWD);

            // Attempt to open the database
            try
            {
                sqlDatabase.Open();
                bDBConnected = true;
            }
            catch (SqlException ex)
            {
                LogOutput("Error connecting to SQL Server:" + ex.Message, "CreateDailyReport", false);
                bDBConnected = false;
            }

            // Write the string to a file.
            ReportFile = ConfigReportPath + "\\" + System.DateTime.Now.ToString("ddMMMyyyy") + "-" + System.DateTime.Now.ToString("HHmm") + "_DailyReport.txt";
            System.IO.StreamWriter reportfile = new System.IO.StreamWriter(ReportFile);

            // Write the header line
            reportfile.WriteLine("*********************************");
            reportfile.WriteLine("Daily Totals report for " + RptDate.ToString("ddMMMyyyy"));
            reportfile.WriteLine("*********************************\n");
            reportfile.WriteLine("Initiative\tProject Name\t\t\tStory\t\t\t\tStory Hours Actual\tStory Hours To Do\tStory State\tReport Date\tCreate Date");
            foreach (Project proj in projects)
            {
                foreach (UserStory story in proj.UserStories)
                {
                    strOutLine = string.Empty;
                    strOutLine = proj.Initiative + "\t" + proj.Name + "\t";
                    strOutLine = strOutLine + story.Name + "\t\t\t\t";
                    strOutLine = strOutLine + story.StoryActual + "\t";
                    strOutLine = strOutLine + story.StoryToDo + "\t";
                    strOutLine = strOutLine + story.State + "\t";
                    strOutLine = strOutLine + RptDate.ToString("ddMMMyyyy") + "\t";
                    strOutLine = strOutLine + DateTime.Now.ToString("ddMMMyyyy") + "\n";
                    reportfile.WriteLine(strOutLine);

                    if (bDBConnected)
                    {
                        SqlCommand sqlCmd = new SqlCommand("INSERT INTO dbo.DailyStats VALUES(@Initiative, @ProjectName, @Story, @HoursStoryActual, " +
                                "@HoursStoryToDo, @State, @ReportDate, @CreateDate)", sqlDatabase);
                        sqlCmd.Parameters.Add(new SqlParameter("Initiative", proj.Initiative));
                        sqlCmd.Parameters.Add(new SqlParameter("ProjectName", proj.Name));
                        sqlCmd.Parameters.Add(new SqlParameter("Story", story.Name));
                        sqlCmd.Parameters.Add(new SqlParameter("HoursStoryActual", story.StoryToDo));
                        sqlCmd.Parameters.Add(new SqlParameter("HoursStoryToDo", story.StoryActual));
                        sqlCmd.Parameters.Add(new SqlParameter("State", story.State));
                        sqlCmd.Parameters.Add(new SqlParameter("ReportDate", RptDate.ToString("ddMMMyyyy")));
                        sqlCmd.Parameters.Add(new SqlParameter("CreateDate", System.DateTime.Now));
                        try
                        {
                            sqlCmd.ExecuteNonQuery();
                        }
                        catch (Exception ex)
                        {
                            LogOutput("Error on insert to SQL Server:" + ex.Message, "CreateDailyReport", false);
                        }
                        sqlCmd.Dispose();
                    }
                }
                strOutLine = string.Empty;
                strOutLine = "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~";
                reportfile.WriteLine(strOutLine);
            }

            // Close the file
            reportfile.Flush();
            reportfile.Close();
            reportfile.Dispose();

            // Close the database
            sqlDatabase.Close();
            sqlDatabase.Dispose();

        }

        /// <summary>
        /// Creates the final Daily output report using the supplied project list
        /// </summary>
        /// <param name="projects">Project list, with calculations, to use for report</param>
        /// <param name="RptYear">Year to run report for</param>
        private static void CreateAnnualReport(List<Project> projects, int RptYear)
        {

            string strOutLine = "";
            bool bDBConnected = false;

            // Create the SQL Server database connection
            SqlConnection sqlDatabase = new SqlConnection("Data Source=" + ConfigDBServer +
                                                            ";Initial Catalog=" + ConfigDBName +
                                                            ";User ID=" + ConfigDBUID +
                                                            ";Password=" + ConfigDBPWD);

            // Attempt to open the database
            try
            {
                sqlDatabase.Open();
                bDBConnected = true;
            }
            catch (SqlException ex)
            {
                LogOutput("Error connecting to SQL Server:" + ex.Message, "CreateAnnualReport", false);
                bDBConnected = false;
            }

            // Write the string to a file.
            ReportFile = ConfigReportPath + "\\" + System.DateTime.Now.ToString("ddMMMyyyy") + "-" + System.DateTime.Now.ToString("HHmm") + "_AnnualReport.txt";
            System.IO.StreamWriter reportfile = new System.IO.StreamWriter(ReportFile);

            // Write the header line
            reportfile.WriteLine("*********************************");
            reportfile.WriteLine(RptYear.ToString() + " year to date totals as of " + System.DateTime.Now.ToString("ddMMMyyyy"));
            reportfile.WriteLine("*********************************\r\n");
            reportfile.WriteLine("Project Name\t\tHours");
            foreach (Project proj in projects)
            {
                strOutLine = proj.Name + "\r\n";
                strOutLine = strOutLine + "   Story Estimate-> " + proj.StoryEstimate + "\t";
                strOutLine = strOutLine + "Story ToDo-> " + proj.StoryToDo + "\t";
                strOutLine = strOutLine + "Story Actual-> " + proj.StoryActual + "\t";
                strOutLine = strOutLine + "Defect Estimate-> " + proj.DefectEstimate + "\t";
                strOutLine = strOutLine + "Defect ToDo-> " + proj.DefectToDo + "\t";
                strOutLine = strOutLine + "Defect Actual-> " + proj.DefectActual + "\r\n";
                strOutLine = strOutLine + " ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" + "\r\n";
                reportfile.WriteLine(strOutLine);

                if (bDBConnected)
                {
                    SqlCommand sqlCmd = new SqlCommand("INSERT INTO dbo.DailyStats VALUES(@Initiative, @ProjectName, @OppAmount, @Owner, @PlannedStartDate, " +
                            "@PlannedEndDate, @Expedite, @State, @CancelledReason, @HoursDefectsEstimate, @HoursDefectsToDo, @HoursDefectsActual, " +
                            "@HoursStoryEstimate, @HoursStoryToDo, @HoursStoryActual, @ReportDate, @UpdateOwner)", sqlDatabase);
                    sqlCmd.Parameters.Add(new SqlParameter("Initiative", proj.Initiative));
                    sqlCmd.Parameters.Add(new SqlParameter("ProjectName", proj.Name));
                    sqlCmd.Parameters.Add(new SqlParameter("OppAmount", proj.OpportunityAmount));
                    sqlCmd.Parameters.Add(new SqlParameter("Owner", proj.Owner));
                    // Check to make sure that PlannedStartDate is within date limits for SQL Server
                    if (proj.PlannedStartDate < Convert.ToDateTime("1/1/2010") || proj.PlannedStartDate > Convert.ToDateTime("1/1/2099"))
                    {
                        sqlCmd.Parameters.Add(new SqlParameter("PlannedStartDate", System.Data.SqlTypes.SqlDateTime.MinValue));
                    }
                    else
                    {
                        sqlCmd.Parameters.Add(new SqlParameter("PlannedStartDate", proj.PlannedStartDate));
                    }
                    // Check to make sure that PlannedEndDate is within date limits for SQL Server
                    if (proj.PlannedEndDate < Convert.ToDateTime("1/1/2010") || proj.PlannedEndDate > Convert.ToDateTime("1/1/2099"))
                    {
                        sqlCmd.Parameters.Add(new SqlParameter("PlannedEndDate", System.Data.SqlTypes.SqlDateTime.MinValue));
                    }
                    else
                    {
                        sqlCmd.Parameters.Add(new SqlParameter("PlannedEndDate", proj.PlannedStartDate));
                    }
                    if (proj.Expedite)
                    {
                        sqlCmd.Parameters.Add(new SqlParameter("Expedite", "Y"));
                    }
                    else
                    {
                        sqlCmd.Parameters.Add(new SqlParameter("Expedite", "N"));
                    }
                    sqlCmd.Parameters.Add(new SqlParameter("State", proj.State));
                    sqlCmd.Parameters.Add(new SqlParameter("CancelledReason", proj.RevokedReason));
                    sqlCmd.Parameters.Add(new SqlParameter("HoursDefectsEstimate", proj.DefectEstimate));
                    sqlCmd.Parameters.Add(new SqlParameter("HoursDefectsToDo", proj.DefectToDo));
                    sqlCmd.Parameters.Add(new SqlParameter("HoursDefectsActual", proj.DefectActual));
                    sqlCmd.Parameters.Add(new SqlParameter("HoursStoryEstimate", proj.StoryEstimate));
                    sqlCmd.Parameters.Add(new SqlParameter("HoursStoryToDo", proj.StoryToDo));
                    sqlCmd.Parameters.Add(new SqlParameter("HoursStoryActual", proj.StoryActual));
                    sqlCmd.Parameters.Add(new SqlParameter("ReportDate", System.DateTime.Now));
                    sqlCmd.Parameters.Add(new SqlParameter("UpdateOwner", proj.UpdateOwner));
                    try
                    {
                        sqlCmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        LogOutput("Error on insert to SQL Server:" + ex.Message, "CreateAnnualReport", false);
                    }
                    sqlCmd.Dispose();
                }
            }

            // Close the file
            reportfile.Flush();
            reportfile.Close();
            reportfile.Dispose();

            // Close the database
            sqlDatabase.Close();
            sqlDatabase.Dispose();

        }

        /// <summary>
        /// Creates the weekly output report using the supplied project list
        /// </summary>
        /// <param name="projects">Project list, with calculations, to use for report</param>
        /// <param name="RptDate">Date to use for calculating Sunday/Saturday end dates</param>
        private static void CreateWeeklyReport(List<Project> projects, DateTime RptDate)
        {

            string strOutLine = "";
            bool bDBConnected = false;

            // Do our rollups to the Epic level


            // Create the SQL Server database connection
            SqlConnection sqlDatabase = new SqlConnection("Data Source=" + ConfigDBServer +
                                                            ";Initial Catalog=" + ConfigDBName +
                                                            ";User ID=" + ConfigDBUID +
                                                            ";Password=" + ConfigDBPWD);

            try
            {
                sqlDatabase.Open();
                bDBConnected = true;
            }
            catch (SqlException ex)
            {
                LogOutput("Error connecting to SQL Server:" + ex.Message, "CreateWeeklyReport", false);
                bDBConnected = false;
            }

            // Write the string to a file.
            ReportFile = ConfigReportPath + "\\" + System.DateTime.Now.ToString("ddMMMyyyy") + "-" + System.DateTime.Now.ToString("HHmm") + "_WeeklyReport.txt";
            System.IO.StreamWriter reportfile = new System.IO.StreamWriter(ReportFile);

            // Write the header line
            reportfile.WriteLine("*********************************");
            reportfile.WriteLine("Weekly report for " + RptDate.ToString("ddMMMyyyy"));
            reportfile.WriteLine("*********************************\r\n");
            strOutLine = "Intiative\tProject Name\t\t\tOpp Amount\tOwner\tPlanned Start Date\tPlanned End Date\tActual End Date\t";
            strOutLine = strOutLine + "Expedite\tState\tCancel Reason\tEpic\t\t\tDefect Hours Est\tDefect Hours To Do\tDefect Hours Actual\t";
            strOutLine = strOutLine + "Story Hours Est\tStory Hours To Do\tStory Hours Actual\tUpdate Owner\tStatus Update\tReport Date";
            reportfile.WriteLine(strOutLine);
            foreach (Project proj in projects)
            {
                foreach (Epic epic in proj.Epics)
                {
                    strOutLine = proj.Initiative + "\t" + proj.Name + "\t" + proj.OpportunityAmount.ToString("$ 0,000,000") + "\t" + proj.Owner + "\t";
                    strOutLine = strOutLine + proj.PlannedStartDate.ToString("ddMMMyyyy") + "\t" + proj.PlannedEndDate.ToString("ddMMMyyyy") + "\t";
                    strOutLine = strOutLine + proj.ActualEndDate.ToString("ddMMMyyyy") + "\t";
                    if (proj.Expedite)
                    {
                        strOutLine = strOutLine + "Yes\t";
                    }
                    else
                    {
                        strOutLine = strOutLine + "No\t";
                    }
                    strOutLine = strOutLine + proj.State + "\t" + proj.RevokedReason + "\t" + epic.Name + "\t" + epic.DefectEstimate.ToString("0.0") + "\t";
                    strOutLine = strOutLine + epic.DefectToDo.ToString("0.0") + "\t" + epic.DefectActual.ToString("0.0") + "\t" + epic.StoryEstimate.ToString("0.0") + "\t";
                    strOutLine = strOutLine + epic.StoryToDo.ToString("0.0") + "\t" + epic.StoryActual.ToString("0.0") + "\t" + proj.UpdateOwner + "\t";
                    strOutLine = strOutLine + proj.StatusUpdate + "\t" + DateTime.Now.ToString("ddMMMyyyy") + "\r\n";
                    //"   " + epic.Name + "   Story Estimate-> " + epic.StoryEstimate + "\r\n";
                    //strOutLine = strOutLine + "   " + epic.Name + "   Story ToDo-> " + epic.StoryToDo + "\r\n";
                    //s/trOutLine = strOutLine + "   " + epic.Name + "   Story Actual-> " + epic.StoryActual + "\r\n";
                    //strOutLine = strOutLine + "   " + epic.Name + "   Defect Estimate-> " + epic.DefectEstimate + "\r\n";
                    //s/trOutLine = strOutLine + "   " + epic.Name + "   Defect ToDo-> " + epic.DefectToDo + "\r\n";
                    //strOutLine = strOutLine + "   " + epic.Name + "   Defect Actual-> " + epic.DefectActual + "\r\n";
                    //strOutLine = strOutLine + " ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" + "\r\n";
                    reportfile.WriteLine(strOutLine);

                    if (bDBConnected)
                    {
                        SqlCommand sqlCmd = new SqlCommand("INSERT INTO dbo.WeeklyStats VALUES(@Initiative, @ProjectName, @OppAmount, @Owner, @PlannedStartDate, " +
                                "@PlannedEndDate, @Expedite, @State, @CancelledReason, @HoursDefectsEstimate, @HoursDefectsToDo, @HoursDefectsActual, " +
                                "@HoursStoryEstimate, @HoursStoryToDo, @HoursStoryActual, @ReportDate, @UpdateOwner, @ProjectUpdate)", sqlDatabase);
                        sqlCmd.Parameters.Add(new SqlParameter("Initiative", proj.Initiative));
                        sqlCmd.Parameters.Add(new SqlParameter("ProjectName", proj.Name));
                        sqlCmd.Parameters.Add(new SqlParameter("OppAmount", proj.OpportunityAmount));
                        sqlCmd.Parameters.Add(new SqlParameter("Owner", proj.Owner));
                        // Check to make sure that PlannedStartDate is within date limits for SQL Server
                        if (proj.PlannedStartDate < Convert.ToDateTime("1/1/2010") || proj.PlannedStartDate > Convert.ToDateTime("1/1/2099"))
                        {
                            sqlCmd.Parameters.Add(new SqlParameter("PlannedStartDate", System.Data.SqlTypes.SqlDateTime.MinValue));
                        }
                        else
                        {
                            sqlCmd.Parameters.Add(new SqlParameter("PlannedStartDate", proj.PlannedStartDate));
                        }
                        // Check to make sure that PlannedEndDate is within date limits for SQL Server
                        if (proj.PlannedEndDate < Convert.ToDateTime("1/1/2010") || proj.PlannedEndDate > Convert.ToDateTime("1/1/2099"))
                        {
                            sqlCmd.Parameters.Add(new SqlParameter("PlannedEndDate", System.Data.SqlTypes.SqlDateTime.MinValue));
                        }
                        else
                        {
                            sqlCmd.Parameters.Add(new SqlParameter("PlannedEndDate", proj.PlannedStartDate));
                        }
                        if (proj.Expedite)
                        {
                            sqlCmd.Parameters.Add(new SqlParameter("Expedite", "Y"));
                        }
                        else
                        {
                            sqlCmd.Parameters.Add(new SqlParameter("Expedite", "N"));
                        }
                        sqlCmd.Parameters.Add(new SqlParameter("State", proj.State));
                        sqlCmd.Parameters.Add(new SqlParameter("CancelledReason", proj.RevokedReason));
                        sqlCmd.Parameters.Add(new SqlParameter("HoursDefectsEstimate", proj.DefectEstimate));
                        sqlCmd.Parameters.Add(new SqlParameter("HoursDefectsToDo", proj.DefectToDo));
                        sqlCmd.Parameters.Add(new SqlParameter("HoursDefectsActual", proj.DefectActual));
                        sqlCmd.Parameters.Add(new SqlParameter("HoursStoryEstimate", proj.StoryEstimate));
                        sqlCmd.Parameters.Add(new SqlParameter("HoursStoryToDo", proj.StoryToDo));
                        sqlCmd.Parameters.Add(new SqlParameter("HoursStoryActual", proj.StoryActual));
                        sqlCmd.Parameters.Add(new SqlParameter("ReportDate", System.DateTime.Now));
                        sqlCmd.Parameters.Add(new SqlParameter("UpdateOwner", proj.UpdateOwner));
                        sqlCmd.Parameters.Add(new SqlParameter("ProjectUpdate", proj.StatusUpdate));
                        try
                        {
                            sqlCmd.ExecuteNonQuery();
                        }
                        catch (Exception ex)
                        {
                            LogOutput("Error on insert to SQL Server:" + ex.Message, "CreateWeeklyReport", false);
                        }
                        sqlCmd.Dispose();
                    }
                }
                strOutLine = strOutLine + " ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" + "\r\n";
            }

            // Close the file
            reportfile.Flush();
            reportfile.Close();
            reportfile.Dispose();

            // Close the database
            sqlDatabase.Close();
            sqlDatabase.Dispose();

        }

        /// <summary>
        /// Creates the weekly output report using the supplied project list
        /// </summary>
        /// <param name="projects">Project list, with calculations, to use for report</param>
        /// <param name="RptMonth">Month to run report for</param>
        private static void CreateMonthlyReport(List<Project> projects, int RptMonth)
        {

            string strOutLine = "";
            bool bDBConnected = false;

            // Create the SQL Server database connection
            SqlConnection sqlDatabase = new SqlConnection("Data Source=" + ConfigDBServer +
                                                            ";Initial Catalog=" + ConfigDBName +
                                                            ";User ID=" + ConfigDBUID +
                                                            ";Password=" + ConfigDBPWD);

            try
            {
                sqlDatabase.Open();
                bDBConnected = true;
            }
            catch (SqlException ex)
            {
                LogOutput("Error connecting to SQL Server:" + ex.Message, "CreateMonthlyReport", false);
                bDBConnected = false;
            }

            // Write the string to a file.
            ReportFile = ConfigReportPath + "\\" + System.DateTime.Now.ToString("ddMMMyyyy") + "-" + System.DateTime.Now.ToString("HHmm") + "_MonthlyReport.txt";
            System.IO.StreamWriter reportfile = new System.IO.StreamWriter(ReportFile);

            // Write the header line
            reportfile.WriteLine("*********************************");
            reportfile.WriteLine("Monthly report for " + RptMonth.ToString("00"));
            reportfile.WriteLine("*********************************\r\n");
            reportfile.WriteLine("Project Name\t\tHours");
            foreach (Project proj in projects)
            {
                strOutLine = proj.Name + "\r\n";
                strOutLine = strOutLine + "   Story Estimate-> " + proj.StoryEstimate + "\t";
                strOutLine = strOutLine + "Story ToDo-> " + proj.StoryToDo + "\t";
                strOutLine = strOutLine + "Story Actual-> " + proj.StoryActual + "\t";
                strOutLine = strOutLine + "Defect Estimate-> " + proj.DefectEstimate + "\t";
                strOutLine = strOutLine + "Defect ToDo-> " + proj.DefectToDo + "\t";
                strOutLine = strOutLine + "Defect Actual-> " + proj.DefectActual + "\r\n";
                strOutLine = strOutLine + " ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" + "\r\n";
                reportfile.WriteLine(strOutLine);

                if (bDBConnected)
                {
                    SqlCommand sqlCmd = new SqlCommand("INSERT INTO dbo.WeeklyStats VALUES(@Initiative, @ProjectName, @OppAmount, @Owner, @PlannedStartDate, " +
                            "@PlannedEndDate, @Expedite, @State, @CancelledReason, @HoursDefectsEstimate, @HoursDefectsToDo, @HoursDefectsActual, " +
                            "@HoursStoryEstimate, @HoursStoryToDo, @HoursStoryActual, @ReportDate, @UpdateOwner, @ProjectUpdate)", sqlDatabase);
                    sqlCmd.Parameters.Add(new SqlParameter("Initiative", proj.Initiative));
                    sqlCmd.Parameters.Add(new SqlParameter("ProjectName", proj.Name));
                    sqlCmd.Parameters.Add(new SqlParameter("OppAmount", proj.OpportunityAmount));
                    sqlCmd.Parameters.Add(new SqlParameter("Owner", proj.Owner));
                    // Check to make sure that PlannedStartDate is within date limits for SQL Server
                    if (proj.PlannedStartDate < Convert.ToDateTime("1/1/2010") || proj.PlannedStartDate > Convert.ToDateTime("1/1/2099"))
                    {
                        sqlCmd.Parameters.Add(new SqlParameter("PlannedStartDate", System.Data.SqlTypes.SqlDateTime.MinValue));
                    }
                    else
                    {
                        sqlCmd.Parameters.Add(new SqlParameter("PlannedStartDate", proj.PlannedStartDate));
                    }
                    // Check to make sure that PlannedEndDate is within date limits for SQL Server
                    if (proj.PlannedEndDate < Convert.ToDateTime("1/1/2010") || proj.PlannedEndDate > Convert.ToDateTime("1/1/2099"))
                    {
                        sqlCmd.Parameters.Add(new SqlParameter("PlannedEndDate", System.Data.SqlTypes.SqlDateTime.MinValue));
                    }
                    else
                    {
                        sqlCmd.Parameters.Add(new SqlParameter("PlannedEndDate", proj.PlannedStartDate));
                    }
                    if (proj.Expedite)
                    {
                        sqlCmd.Parameters.Add(new SqlParameter("Expedite", "Y"));
                    }
                    else
                    {
                        sqlCmd.Parameters.Add(new SqlParameter("Expedite", "N"));
                    }
                    sqlCmd.Parameters.Add(new SqlParameter("State", proj.State));
                    sqlCmd.Parameters.Add(new SqlParameter("CancelledReason", proj.RevokedReason));
                    sqlCmd.Parameters.Add(new SqlParameter("HoursDefectsEstimate", proj.DefectEstimate));
                    sqlCmd.Parameters.Add(new SqlParameter("HoursDefectsToDo", proj.DefectToDo));
                    sqlCmd.Parameters.Add(new SqlParameter("HoursDefectsActual", proj.DefectActual));
                    sqlCmd.Parameters.Add(new SqlParameter("HoursStoryEstimate", proj.StoryEstimate));
                    sqlCmd.Parameters.Add(new SqlParameter("HoursStoryToDo", proj.StoryToDo));
                    sqlCmd.Parameters.Add(new SqlParameter("HoursStoryActual", proj.StoryActual));
                    sqlCmd.Parameters.Add(new SqlParameter("ReportDate", System.DateTime.Now));
                    sqlCmd.Parameters.Add(new SqlParameter("UpdateOwner", proj.UpdateOwner));
                    sqlCmd.Parameters.Add(new SqlParameter("ProjectUpdate", proj.StatusUpdate));
                    try
                    {
                        sqlCmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        LogOutput("Error on insert to SQL Server:" + ex.Message, "CreateMonthlyReport", false);
                    }
                    sqlCmd.Dispose();
                }
            }

            // Close the file
            reportfile.Flush();
            reportfile.Close();
            reportfile.Dispose();

            // Close the database
            sqlDatabase.Close();
            sqlDatabase.Dispose();

        }

        /// <summary>
        /// Creates the final quarter output report using the supplied project list
        /// </summary>
        /// <param name="projects">Project list, with calculations, to use for report</param>
        /// <param name="ReportPeriod">Quarter/Year to report on in format Q#YYYY</param>
        /// <param name="ReportingPeriod">Quarter reporting on</param>
        private static void CreateQuarterReport(List<Project> projects, string ReportPeriod)
        {

            string strOutLine = "";
            bool bDBConnected = false;

            // Create the SQL Server database connection
            SqlConnection sqlDatabase = new SqlConnection("Data Source=" + ConfigDBServer +
                                                            ";Initial Catalog=" + ConfigDBName +
                                                            ";User ID=" + ConfigDBUID +
                                                            ";Password=" + ConfigDBPWD);

            try
            {
                sqlDatabase.Open();
                bDBConnected = true;
            }
            catch (SqlException ex)
            {
                LogOutput("Error connecting to SQL Server:" + ex.Message, "CreateQuarterReport", false);
                bDBConnected = false;
            }

            // Write the string to a file.
            ReportFile = ConfigReportPath + "\\" + System.DateTime.Now.ToString("ddMMMyyyy") + "-" + System.DateTime.Now.ToString("HHmm") + "_QuarterReport.txt";
            System.IO.StreamWriter reportfile = new System.IO.StreamWriter(ReportFile);

            // Write the header line
            reportfile.WriteLine("*********************************");
            reportfile.WriteLine("Quarterly report for " + ReportPeriod);
            reportfile.WriteLine("*********************************\r\n");
            reportfile.WriteLine("Project Name\t\tHours");
            foreach (Project proj in projects)
            {
                strOutLine = proj.Name + "\r\n";
                strOutLine = strOutLine + "   Story Estimate-> " + proj.StoryEstimate + "\t";
                strOutLine = strOutLine + "Story ToDo-> " + proj.StoryToDo + "\t";
                strOutLine = strOutLine + "Story Actual-> " + proj.StoryActual + "\t";
                strOutLine = strOutLine + "Defect Estimate-> " + proj.DefectEstimate + "\t";
                strOutLine = strOutLine + "Defect ToDo-> " + proj.DefectToDo + "\t";
                strOutLine = strOutLine + "Defect Actual-> " + proj.DefectActual + "\r\n";
                strOutLine = strOutLine + " ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" + "\r\n";
                reportfile.WriteLine(strOutLine);

                if (bDBConnected)
                {
                    SqlCommand sqlCmd = new SqlCommand("INSERT INTO dbo.QuarterStats VALUES(@Initiative, @ProjectName, @OppAmount, @Owner, @PlannedStartDate, " +
                            "@PlannedEndDate, @Expedite, @State, @CancelledReason, @HoursDefectsEstimate, @HoursDefectsToDo, @HoursDefectsActual, " +
                            "@HoursStoryEstimate, @HoursStoryToDo, @HoursStoryActual, @ReportDate, @UpdateOwner, @ReportingPeriod)", sqlDatabase);
                    sqlCmd.Parameters.Add(new SqlParameter("Initiative", proj.Initiative));
                    sqlCmd.Parameters.Add(new SqlParameter("ProjectName", proj.Name));
                    sqlCmd.Parameters.Add(new SqlParameter("OppAmount", proj.OpportunityAmount));
                    sqlCmd.Parameters.Add(new SqlParameter("Owner", proj.Owner));
                    // Check to make sure that PlannedStartDate is within date limits for SQL Server
                    if (proj.PlannedStartDate < Convert.ToDateTime("1/1/2010") || proj.PlannedStartDate > Convert.ToDateTime("1/1/2099"))
                    {
                        sqlCmd.Parameters.Add(new SqlParameter("PlannedStartDate", System.Data.SqlTypes.SqlDateTime.MinValue));
                    }
                    else
                    {
                        sqlCmd.Parameters.Add(new SqlParameter("PlannedStartDate", proj.PlannedStartDate));
                    }
                    // Check to make sure that PlannedEndDate is within date limits for SQL Server
                    if (proj.PlannedEndDate < Convert.ToDateTime("1/1/2010") || proj.PlannedEndDate > Convert.ToDateTime("1/1/2099"))
                    {
                        sqlCmd.Parameters.Add(new SqlParameter("PlannedEndDate", System.Data.SqlTypes.SqlDateTime.MinValue));
                    }
                    else
                    {
                        sqlCmd.Parameters.Add(new SqlParameter("PlannedEndDate", proj.PlannedStartDate));
                    }
                    if (proj.Expedite)
                    {
                        sqlCmd.Parameters.Add(new SqlParameter("Expedite", "Y"));
                    }
                    else
                    {
                        sqlCmd.Parameters.Add(new SqlParameter("Expedite", "N"));
                    }
                    sqlCmd.Parameters.Add(new SqlParameter("State", proj.State));
                    sqlCmd.Parameters.Add(new SqlParameter("CancelledReason", proj.RevokedReason));
                    sqlCmd.Parameters.Add(new SqlParameter("HoursDefectsEstimate", proj.DefectEstimate));
                    sqlCmd.Parameters.Add(new SqlParameter("HoursDefectsToDo", proj.DefectToDo));
                    sqlCmd.Parameters.Add(new SqlParameter("HoursDefectsActual", proj.DefectActual));
                    sqlCmd.Parameters.Add(new SqlParameter("HoursStoryEstimate", proj.StoryEstimate));
                    sqlCmd.Parameters.Add(new SqlParameter("HoursStoryToDo", proj.StoryToDo));
                    sqlCmd.Parameters.Add(new SqlParameter("HoursStoryActual", proj.StoryActual));
                    sqlCmd.Parameters.Add(new SqlParameter("ReportDate", System.DateTime.Now));
                    sqlCmd.Parameters.Add(new SqlParameter("UpdateOwner", proj.UpdateOwner));
                    sqlCmd.Parameters.Add(new SqlParameter("ReportingPeriod", ReportPeriod));
                    try
                    {
                        sqlCmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        LogOutput("Error on insert to SQL Server:" + ex.Message, "CreateQuarterReport", false);
                    }
                    sqlCmd.Dispose();
                }
            }

            // Close the file
            reportfile.Flush();
            reportfile.Close();
            reportfile.Dispose();

            // Close the database
            sqlDatabase.Close();
            sqlDatabase.Dispose();

        }

        /// <summary>
        /// Runs through all tasks associated with a project (both User Stories and Defects) and totals the Estimates, ToDo, and Actuals for the day
        /// </summary>
        /// <param name="SourceProjectList">Original Project list to calculate</param>
        /// <param name="RptDay">Full date/time to calculate totals for</param>
        private static List<Project> CalculateDailyTotals(List<Project> SourceProjectList, DateTime RptDay)
        {

            List<Project> DestinationProjectList = new List<Project>();

            // Once again, we need to loop through all projects and build a new list 
            // with the summary information included
            foreach (Project proj in SourceProjectList)
            {
                Project newproject = new Project();
                newproject.Description = proj.Description;
                newproject.Expedite = proj.Expedite;
                newproject.FormattedID = proj.FormattedID;
                newproject.Initiative = proj.Initiative;
                newproject.Name = proj.Name;
                newproject.OpportunityAmount = proj.OpportunityAmount;
                newproject.Owner = proj.Owner;
                newproject.UpdateOwner = proj.UpdateOwner;
                newproject.PlannedEndDate = proj.PlannedEndDate;
                newproject.PlannedStartDate = proj.PlannedStartDate;
                newproject.RevokedReason = proj.RevokedReason;
                newproject.State = proj.State;
                newproject.StatusUpdate = proj.StatusUpdate;
                newproject.UserStories = CreateDailyRollup(proj.UserStories, RptDay);
                newproject.Defects = CreateDailyRollup(proj.Defects, RptDay);
                newproject.StoryEstimate = CreateDailyTotal(newproject.UserStories, ProjectTotal.Estimate, RptDay);
                newproject.StoryToDo = CreateDailyTotal(newproject.UserStories, ProjectTotal.ToDo, RptDay);
                newproject.StoryActual = CreateDailyTotal(newproject.UserStories, ProjectTotal.Actual, RptDay);
                newproject.DefectEstimate = CreateDailyTotal(newproject.Defects, ProjectTotal.Estimate, RptDay);
                newproject.DefectToDo = CreateDailyTotal(newproject.Defects, ProjectTotal.ToDo, RptDay);
                newproject.DefectActual = CreateDailyTotal(newproject.Defects, ProjectTotal.Actual, RptDay);
                if (newproject.StoryEstimate != 0 ||
                    newproject.StoryToDo != 0 ||
                    newproject.StoryActual != 0 ||
                    newproject.DefectEstimate != 0 ||
                    newproject.DefectToDo != 0 ||
                    newproject.DefectActual != 0)
                {
                    DestinationProjectList.Add(newproject);
                }
            }

            return DestinationProjectList;

        }

        /// <summary>
        /// Runs through all tasks associated with a project (both User Stories and Defects) and totals the Estimates, ToDo, and Actuals
        ///  for the current week.  Also saves the project update for the week.
        /// </summary>
        /// <param name="SourceProjectList">Original Project list to calculate</param>
        /// <param name="RptDate">Full date/time to calculate totals for.  This is used to find the Sunday and Saturday of the week for reporting</param>
        private static List<Project> CalculateWeeklyTotals(List<Project> SourceProjectList, DateTime RptDate)
        {

            List<Project> DestinationProjectList = new List<Project>();

            // Once again, we need to loop through all projects and build a new list 
            // with the summary information included
            foreach (Project proj in SourceProjectList)
            {
                Project newproject = new Project();
                newproject.Description = proj.Description;
                newproject.Expedite = proj.Expedite;
                newproject.FormattedID = proj.FormattedID;
                newproject.Initiative = proj.Initiative;
                newproject.Name = proj.Name;
                newproject.OpportunityAmount = proj.OpportunityAmount;
                newproject.Owner = proj.Owner;
                newproject.UpdateOwner = proj.UpdateOwner;
                newproject.PlannedEndDate = proj.PlannedEndDate;
                newproject.PlannedStartDate = proj.PlannedStartDate;
                newproject.RevokedReason = proj.RevokedReason;
                newproject.State = proj.State;
                // Grab just the latest update
                newproject.StatusUpdate = GetLatestStatus(proj.StatusUpdate);
                newproject.UserStories = proj.UserStories;
                newproject.Epics = CreateWeeklyEpicSummary(proj.Epics, newproject.UserStories, newproject.Defects, RptDate);
                newproject.Defects = proj.Defects;
                newproject.StoryEstimate = CreateWeeklyTotal(newproject.UserStories, ProjectTotal.Estimate, RptDate);
                newproject.StoryToDo = CreateWeeklyTotal(newproject.UserStories, ProjectTotal.ToDo, RptDate);
                newproject.StoryActual = CreateWeeklyTotal(newproject.UserStories, ProjectTotal.Actual, RptDate);
                newproject.DefectEstimate = CreateWeeklyTotal(newproject.Defects, ProjectTotal.Estimate, RptDate);
                newproject.DefectToDo = CreateWeeklyTotal(newproject.Defects, ProjectTotal.ToDo, RptDate);
                newproject.DefectActual = CreateWeeklyTotal(newproject.Defects, ProjectTotal.Actual, RptDate);
                DestinationProjectList.Add(newproject);
                //if (newproject.StoryEstimate != 0 ||
                //    newproject.StoryToDo != 0 ||
                //    newproject.StoryActual != 0 ||
                //    newproject.DefectEstimate != 0 ||
                //    newproject.DefectToDo != 0 ||
                //    newproject.DefectActual != 0)
                //{
                //    DestinationProjectList.Add(newproject);
                //}
            }

            return DestinationProjectList;

        }

        /// <summary>
        /// Runs through all tasks associated with a project (both User Stories and Defects) and totals the Estimates, ToDo, and Actuals
        /// for the current year.
        /// </summary>
        /// <param name="SourceProjectList">Original Project list to calculate</param>
        /// <param name="RptYear">Year to calculate totals for</param>
        private static List<Project> CalculateAnnualTotals(List<Project> SourceProjectList, int RptYear)
        {

            List<Project> DestinationProjectList = new List<Project>();

            // Once again, we need to loop through all projects and build a new list 
            // with the summary information included
            foreach (Project proj in SourceProjectList)
            {
                Project newproject = new Project();
                newproject.Description = proj.Description;
                newproject.Expedite = proj.Expedite;
                newproject.FormattedID = proj.FormattedID;
                newproject.Initiative = proj.Initiative;
                newproject.Name = proj.Name;
                newproject.OpportunityAmount = proj.OpportunityAmount;
                newproject.Owner = proj.Owner;
                newproject.UpdateOwner = proj.UpdateOwner;
                newproject.PlannedEndDate = proj.PlannedEndDate;
                newproject.PlannedStartDate = proj.PlannedStartDate;
                newproject.RevokedReason = proj.RevokedReason;
                newproject.State = proj.State;
                newproject.StatusUpdate = proj.StatusUpdate;
                newproject.UserStories = proj.UserStories;
                newproject.Defects = proj.Defects;
                newproject.StoryEstimate = CreateAnnualTotal(newproject.UserStories, ProjectTotal.Estimate, RptYear);
                newproject.StoryToDo = CreateAnnualTotal(newproject.UserStories, ProjectTotal.ToDo, RptYear);
                newproject.StoryActual = CreateAnnualTotal(newproject.UserStories, ProjectTotal.Actual, RptYear);
                newproject.DefectEstimate = CreateAnnualTotal(newproject.Defects, ProjectTotal.Estimate, RptYear);
                newproject.DefectToDo = CreateAnnualTotal(newproject.Defects, ProjectTotal.ToDo, RptYear);
                newproject.DefectActual = CreateAnnualTotal(newproject.Defects, ProjectTotal.Actual, RptYear);
                DestinationProjectList.Add(newproject);
                //if (newproject.StoryEstimate != 0 ||
                //    newproject.StoryToDo != 0 ||
                //    newproject.StoryActual != 0 ||
                //    newproject.DefectEstimate != 0 ||
                //    newproject.DefectToDo != 0 ||
                //   newproject.DefectActual != 0)
                //{
                //    DestinationProjectList.Add(newproject);
                //}
            }

            return DestinationProjectList;

        }

        /// <summary>
        /// Runs through all projects and generates the summary information on monthly basis.  There are no numeric
        /// statistics calculated
        /// </summary>
        /// <param name="SourceProjectList">Original Project list to calculate</param>
        /// <param name="RptMonth">Year to calculate totals for</param>
        private static List<Project> CalculateMonthlyReport(List<Project> SourceProjectList, int RptMonth)
        {

            List<Project> DestinationProjectList = new List<Project>();

            // Once again, we need to loop through all projects and build a new list 
            // with the summary information included
            foreach (Project proj in SourceProjectList)
            {
                Project newproject = new Project();
                newproject.Description = proj.Description;
                newproject.Expedite = proj.Expedite;
                newproject.FormattedID = proj.FormattedID;
                newproject.Initiative = proj.Initiative;
                newproject.Name = proj.Name;
                newproject.OpportunityAmount = proj.OpportunityAmount;
                newproject.Owner = proj.Owner;
                newproject.UpdateOwner = proj.UpdateOwner;
                newproject.PlannedEndDate = proj.PlannedEndDate;
                newproject.PlannedStartDate = proj.PlannedStartDate;
                newproject.RevokedReason = proj.RevokedReason;
                newproject.State = proj.State;
                newproject.StatusUpdate = proj.StatusUpdate;
                // We only want to include this project if the project is still open or was
                // closed DURING the reporting month 
                if (newproject.State == "Done")
                {
                    //if done and last update is during the reporting month, then add
                    if (IsReportMonth(newproject.LastUpdateDate, RptMonth))
                    {
                        DestinationProjectList.Add(newproject);
                    }
                }
                else
                {
                    DestinationProjectList.Add(newproject);
                }
            }

            return DestinationProjectList;

        }

        /// <summary>
        /// Runs through all tasks associated with a project (both User Stories and Defects) and totals the Estimates, ToDo, and Actuals for the quarter
        /// </summary>
        /// <param name="SourceProjectList">Original Project list to calculate</param>
        /// <param name="RptQuarter">Quarter to generate totals for</param>
        /// <param name="RptYear">Year to generate totals for</param>
        private static List<Project> CalculateQuarterTotals(List<Project> SourceProjectList, int RptQuarter, int RptYear)
        {

            List<Project> DestinationProjectList = new List<Project>();

            // Once again, we need to loop through all projects and build a new list 
            // with the summary information included
            foreach (Project proj in SourceProjectList)
            {
                Project newproject = new Project();
                newproject.Description = proj.Description;
                newproject.Expedite = proj.Expedite;
                newproject.FormattedID = proj.FormattedID;
                newproject.Initiative = proj.Initiative;
                newproject.Name = proj.Name;
                newproject.OpportunityAmount = proj.OpportunityAmount;
                newproject.Owner = proj.Owner;
                newproject.UpdateOwner = proj.UpdateOwner;
                newproject.PlannedEndDate = proj.PlannedEndDate;
                newproject.PlannedStartDate = proj.PlannedStartDate;
                newproject.RevokedReason = proj.RevokedReason;
                newproject.State = proj.State;
                newproject.StatusUpdate = proj.StatusUpdate;
                newproject.UserStories = proj.UserStories;
                newproject.Defects = proj.Defects;
                newproject.StoryEstimate = CreateQuarterTotal(newproject.UserStories, ProjectTotal.Estimate, RptQuarter, RptYear);
                newproject.StoryToDo = CreateQuarterTotal(newproject.UserStories, ProjectTotal.ToDo, RptQuarter, RptYear);
                newproject.StoryActual = CreateQuarterTotal(newproject.UserStories, ProjectTotal.Actual, RptQuarter, RptYear);
                newproject.DefectEstimate = CreateQuarterTotal(newproject.Defects, ProjectTotal.Estimate, RptQuarter, RptYear);
                newproject.DefectToDo = CreateQuarterTotal(newproject.Defects, ProjectTotal.ToDo, RptQuarter, RptYear);
                newproject.DefectActual = CreateQuarterTotal(newproject.Defects, ProjectTotal.Actual, RptQuarter, RptYear);
                if (newproject.StoryEstimate != 0 ||
                    newproject.StoryToDo != 0 ||
                    newproject.StoryActual != 0 ||
                    newproject.DefectEstimate != 0 ||
                    newproject.DefectToDo != 0 ||
                    newproject.DefectActual != 0)
                {
                    DestinationProjectList.Add(newproject);
                }
            }

            return DestinationProjectList;

        }

        /// <summary>
        /// Accepts list of all User Stories / Defects and calculates daily totals
        /// </summary>
        /// <param name="stories">List of stories to calculate</param>
        /// <param name="Action">Indicates whether to total for Estimate, ToDo, or Actuals</param>
        /// <param name="DayforDaily">Date to calculate totals for</param>
        private static decimal CreateDailyTotal(List<UserStory> stories, ProjectTotal Action, DateTime DayforDaily)
        {

            decimal decReturn = 0;

            foreach (UserStory story in stories)
            {
                foreach (Task task in story.Tasks)
                {
                    if (DayforDaily.ToShortDateString() == task.LastUpdate.ToShortDateString())
                    {
                        switch (Action)
                        {
                            case ProjectTotal.Estimate:
                                decReturn = decReturn + task.Estimate;
                                break;
                            case ProjectTotal.ToDo:
                                decReturn = decReturn + task.ToDo;
                                break;
                            case ProjectTotal.Actual:
                                decReturn = decReturn + task.Actual;
                                break;
                            default:
                                break;
                        }

                    }
                }
            }

            return decReturn;

        }

        /// <summary>
        /// Accepts list of all User Stories / Defects and calculates daily totals
        /// </summary>
        /// <param name="defects">List of Defects to calculate</param>
        /// <param name="Action">Indicates whether to total for Estimate, ToDo, or Actuals</param>
        /// <param name="DayforDaily">Date to calculate totals for</param>
        private static decimal CreateDailyTotal(List<Defect> defects, ProjectTotal Action, DateTime DayforDaily)
        {

            decimal decReturn = 0;

            foreach (Defect defect in defects)
            {
                foreach (Task task in defect.Tasks)
                {
                    if (DayforDaily.ToShortDateString() == task.LastUpdate.ToShortDateString())
                    {
                        switch (Action)
                        {
                            case ProjectTotal.Estimate:
                                decReturn = decReturn + task.Estimate;
                                break;
                            case ProjectTotal.ToDo:
                                decReturn = decReturn + task.ToDo;
                                break;
                            case ProjectTotal.Actual:
                                decReturn = decReturn + task.Actual;
                                break;
                            default:
                                break;
                        }
                    }
                }
            }

            return decReturn;

        }

        /// <summary>
        /// Accepts list of all User Stories / Defects and calculates quarter totals
        /// </summary>
        /// <param name="stories">List of stories to calculate</param>
        /// <param name="Action">Indicates whether to total for Estimate, ToDo, or Actuals</param>
        /// <param name="RptQuarter">Which quarter to generate totals for</param>
        /// <param name="RptYear">Which Year to generate totals for</param>
        private static decimal CreateQuarterTotal(List<UserStory> stories, ProjectTotal Action, int RptQuarter, int RptYear)
        {

            decimal decReturn = 0;
            string ReportPeriod = "";
            string StoryRptPeriod = "";

            // Set the string to indicate the Reporting Period
            ReportPeriod = "Q" + RptQuarter + RptYear;

            // Loop through the stories
            foreach (UserStory story in stories)
            {
                // Set the string to indicate the story period
                // Releases are quarterly and of the format--> 2016'Q1 CTO C&I Labs
                if (story.Release != "")
                {
                    StoryRptPeriod = "Q" + story.Release.Substring(6, 1) + story.Release.Substring(0, 4);
                }
                else
                {
                    StoryRptPeriod = "";
                }
                if (ReportPeriod == StoryRptPeriod)
                {
                    foreach (Task task in story.Tasks)
                    {
                        switch (Action)
                        {
                            case ProjectTotal.Estimate:
                                decReturn = decReturn + task.Estimate;
                                break;
                            case ProjectTotal.ToDo:
                                decReturn = decReturn + task.ToDo;
                                break;
                            case ProjectTotal.Actual:
                                decReturn = decReturn + task.Actual;
                                break;
                            default:
                                break;
                        }
                    }
                }
            }

            return decReturn;

        }

        /// <summary>
        /// Accepts list of all User Stories / Defects and calculates quarter totals
        /// </summary>
        /// <param name="defects">List of Defects to calculate</param>
        /// <param name="Action">Indicates whether to total for Estimate, ToDo, or Actuals</param>
        /// <param name="RptQuarter">Which quarter to generate totals for</param>
        /// <param name="RptYear">Which Year to generate totals for</param>
        private static decimal CreateQuarterTotal(List<Defect> defects, ProjectTotal Action, int RptQuarter, int RptYear)
        {

            decimal decReturn = 0;
            string ReportPeriod = "";
            string StoryRptPeriod = "";

            // Set the string to indicate the Reporting Period
            ReportPeriod = "Q" + RptQuarter + RptYear;

            // Loop through the stories
            foreach (Defect defect in defects)
            {
                // Set the string to indicate the story period
                // Releases are quarterly and of the format--> 2016'Q1 CTO C&I Labs
                if (defect.Release != "")
                {
                    StoryRptPeriod = "Q" + defect.Release.Substring(6, 1) + defect.Release.Substring(0, 4);
                }
                else
                {
                    StoryRptPeriod = "";
                }
                if (ReportPeriod == StoryRptPeriod)
                {
                    foreach (Task task in defect.Tasks)
                    {
                        switch (Action)
                        {
                            case ProjectTotal.Estimate:
                                decReturn = decReturn + task.Estimate;
                                break;
                            case ProjectTotal.ToDo:
                                decReturn = decReturn + task.ToDo;
                                break;
                            case ProjectTotal.Actual:
                                decReturn = decReturn + task.Actual;
                                break;
                            default:
                                break;
                        }
                    }
                }
            }

            return decReturn;

        }

        /// <summary>
        /// Grabs all configuration information from config file
        /// </summary>
        private static bool GetConfigSettings()
        {

            try
            {
                ConfigRallyUID = ConfigurationManager.AppSettings["RallyUser"];
                ConfigRallyPWD = ConfigurationManager.AppSettings["RallyPass"];
                ConfigStatusEOL = ConfigurationManager.AppSettings["StatusUpdateDiv"];
                ConfigLogPath = ConfigurationManager.AppSettings["LogPath"];
                ConfigReportPath = ConfigurationManager.AppSettings["ReportPath"];
                ConfigDBServer = ConfigurationManager.AppSettings["DBServer"];
                ConfigDBName = ConfigurationManager.AppSettings["DBName"];
                ConfigDBUID = ConfigurationManager.AppSettings["DBUID"];
                ConfigDBPWD = ConfigurationManager.AppSettings["DBPWD"];

                if (ConfigurationManager.AppSettings["Debug"].ToUpper() == "Y")
                {
                    ConfigDebug = true;
                }
                else
                {
                    ConfigDebug = false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to read configuration settings: " + ex.Message);
                return false;
            }

            LogFile = ConfigLogPath + "\\" + System.DateTime.Now.ToString("ddMMMyyyy") + "-" + System.DateTime.Now.ToString("HHmm") + "_Log.txt";

            return true;

        }

        /// <summary>
        /// Gets all Command Line Arguments
        /// </summary>
        private static void GetCommandArgs()
        {

            // Now get any commandline args
            string[] CmdArgs = Environment.GetCommandLineArgs();

            // Check for proper switches
            if (!ValidArgs(CmdArgs))
            {
                // If ValidArgs returns false, then the supplied arguments are invalid
                // therefore, we message the user on proper usage and then UsageMessage terminates the application
                UsageMessage();
            }

            // Loop through all commandline args
            for (int LoopCtl = 0; LoopCtl < CmdArgs.Length; LoopCtl++)
            {
                // Daily Stats = No Switch / -DddMonYYYY
                // Weekly Status Update = -W or -WddMonYYYY
                // Monthly Stats = -M#
                // Quarterly Stats = -Q#YYYY
                // Annual Stats = -AYYYY
                string CommandLinePart = CmdArgs[LoopCtl];

                switch (CommandLinePart.Substring(0, 2).ToUpper())
                {
                    case "/H":
                        // Message the user what the usage is
                        UsageMessage();
                        break;
                    case "-Q":
                        // Format should be -Q<number><Year>
                        OperatingMode = OperateMode.Quarterly;
                        ReportQuarter = Convert.ToInt32(CommandLinePart.Substring(2, 1));
                        ReportYear = Convert.ToInt32(CommandLinePart.Substring(3, 4));
                        break;
                    case "-A":
                        // Format should be -A<Year>
                        OperatingMode = OperateMode.Annual;
                        ReportYear = Convert.ToInt32(CommandLinePart.Substring(2, 4));
                        break;
                    case "-M":
                        // Format should be -M<Number>
                        OperatingMode = OperateMode.Monthly;
                        ReportMonth = Convert.ToInt32(CommandLinePart.Substring(2, 1));
                        break;
                    case "-W":
                        // Format should be -W or -WddMonYYYY
                        OperatingMode = OperateMode.Weekly;
                        if (CommandLinePart.Length > 2)
                        {
                            ReportDayNum = Convert.ToInt32(CommandLinePart.Substring(2, 2));
                            ReportMonth = MonthNumber(CommandLinePart.Substring(4, 3));
                            ReportYear = Convert.ToInt32(CommandLinePart.Substring(7, 4));
                            ReportWeekNum = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(new DateTime(ReportYear, ReportMonth, ReportDayNum),
                                CalendarWeekRule.FirstDay, DayOfWeek.Sunday);
                        }
                        else
                        {
                            ReportDayNum = DateTime.Today.Day;
                            ReportMonth = DateTime.Today.Month;
                            ReportYear = DateTime.Today.Year;
                            ReportingDay = new DateTime(ReportYear, ReportMonth, ReportDayNum);
                            ReportWeekNum = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(ReportingDay, CalendarWeekRule.FirstDay, DayOfWeek.Sunday);
                        }
                        break;
                    case "-D":
                        // Format should be -DddMonYYYY
                        OperatingMode = OperateMode.Daily;
                        ReportDayNum = Convert.ToInt32(CommandLinePart.Substring(2, 2));
                        ReportMonth = MonthNumber(CommandLinePart.Substring(4, 3));
                        ReportYear = Convert.ToInt32(CommandLinePart.Substring(7, 4));
                        ReportingDay = new DateTime(ReportYear, ReportMonth, ReportDayNum);
                        break;
                    default:
                        // If the last part of the commandline contains "EXE" then it is the full path to the executable and we don't care
                        if (CommandLinePart.Contains(".exe"))
                        {
                            break;
                        }
                        // Format should be -DddMonYYYY
                        OperatingMode = OperateMode.Daily;
                        ReportingDay = DateTime.Today.AddDays(-1);
                        ReportDayNum = ReportingDay.Day;
                        ReportMonth = ReportingDay.Month;
                        ReportYear = ReportingDay.Year;
                        break;
                }
            }

        }

        /// <summary>
        /// Loops through the supplied list of Initiatives to ignore to see if the supplied Initiative IDNumber
        /// is in the list.  If it is, returns "True" indicating that we should ignore.
        /// </summary>
        /// <param name="IgnoreList">Array of InitiativeIDs to ignore</param>
        /// <param name="IDToCheck">InitiativeID to check the list against</param>
        private static bool CheckIgnoreList(string[] IgnoreList, string IDToCheck)
        {

            // Loop through the IgnoreList
            foreach (string stringpart in IgnoreList)
            {
                // Compare the Ignore item to the current Initiative ID number
                if (stringpart == IDToCheck)
                {
                    // YES, this should be ignored
                    return true;
                }
            }

            // If we get here, then the ID is not in the Ignore list, so NO, we should not ignore it
            return false;

        }

        /// <summary>
        /// Accepts a string with HTML tags and removes the tags for non-HTML display
        /// </summary>
        /// <param name="HTMLString">String containing HTML tags</param>
        private static string GetPlainTextFromHtml(string HTMLString)
        {

            string htmlTagPattern = "<.*?>";
            var regexCss = new Regex("(\\<script(.+?)\\</script\\>)|(\\<style(.+?)\\</style\\>)", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            HTMLString = regexCss.Replace(HTMLString, string.Empty);
            HTMLString = Regex.Replace(HTMLString, htmlTagPattern, string.Empty);
            HTMLString = Regex.Replace(HTMLString, @"^\s+$[\r\n]*", "", RegexOptions.Multiline);
            HTMLString = HTMLString.Replace("&nbsp;", string.Empty);
            HTMLString = HTMLString.Replace("&lt;", "<");
            HTMLString = HTMLString.Replace("&gt;", ">");

            return HTMLString;

        }

        /// <summary>
        /// Accepts the full status field from Rally, which includes all status updates, and extracts the latest
        /// </summary>
        /// <param name="FullStatus">Complete status history from Rally</param>
        private static string GetLatestStatus(string FullStatus)
        {

            string RecentStatus = "";

            if (FullStatus.IndexOf(ConfigStatusEOL) > 0)
            {
                RecentStatus = FullStatus.Substring(0, FullStatus.IndexOf(ConfigStatusEOL));
            }
            else
            {
                RecentStatus = FullStatus.Trim();
            }

            return RecentStatus;

        }

        /// <summary>
        /// Calculates the numeric value for the month based on the "Short Month"
        /// </summary>
        /// <param name="ShortMonth">"Short Month" input (i.e. 'Jan', 'Feb', etc)</param>
        private static int MonthNumber(string ShortMonth)
        {

            switch (ShortMonth.ToUpper())
            {
                case "JAN":
                    return 1;
                case "FEB":
                    return 2;
                case "MAR":
                    return 3;
                case "APR":
                    return 4;
                case "MAY":
                    return 5;
                case "JUN":
                    return 6;
                case "JUL":
                    return 7;
                case "AUG":
                    return 8;
                case "SEP":
                    return 9;
                case "OCT":
                    return 10;
                case "NOV":
                    return 11;
                case "DEC":
                    return 12;
                default:
                    return 0;
            }

        }

        /// <summary>
        /// This is used to dump out usage instructions to the user on the commandline
        /// </summary>
        public static void UsageMessage()
        {

            Environment.Exit(0);

        }

        /// <summary>
        /// Accepts list of all User Stories / Defects and calculates weekly totals
        /// </summary>
        /// <param name="defects">List of Defects to calculate</param>
        /// <param name="Action">Indicates whether to total for Estimate, ToDo, or Actuals</param>
        /// <param name="DayforWeekly">Date to use for calculating Sunday/Saturday end dates</param>
        /// <param name="RptQuarter">Which quarter to generate totals for</param>
        /// <param name="RptYear">Which Year to generate totals for</param>
        private static decimal CreateWeeklyTotal(List<Defect> defects, ProjectTotal Action, DateTime DayforWeekly)
        {

            decimal decReturn = 0;
            DateTime startDate;
            DateTime endDate;

            // Set the end dates for the Reporting Period.  Should start on Sunday and end on Saturday
            startDate = GetFirstDayOfWeek(DayforWeekly, CultureInfo.CurrentCulture);
            endDate = startDate.AddDays(6);

            // Loop through the stories
            foreach (Defect defect in defects)
            {
                foreach (Task task in defect.Tasks)
                {
                    if ((task.LastUpdate >= startDate) && (task.LastUpdate <= endDate))
                    {
                        switch (Action)
                        {
                            case ProjectTotal.Estimate:
                                decReturn = decReturn + task.Estimate;
                                break;
                            case ProjectTotal.ToDo:
                                decReturn = decReturn + task.ToDo;
                                break;
                            case ProjectTotal.Actual:
                                decReturn = decReturn + task.Actual;
                                break;
                            default:
                                break;
                        }
                    }
                }

            }

            return decReturn;

        }

        /// <summary>
        /// Accepts list of all User Stories / Defects and calculates weekly totals
        /// </summary>
        /// <param name="defects">List of Defects to calculate</param>
        /// <param name="Action">Indicates whether to total for Estimate, ToDo, or Actuals</param>
        /// <param name="YearForAnnual">Which Year to generate totals for</param>
        private static decimal CreateAnnualTotal(List<Defect> defects, ProjectTotal Action, int YearForAnnual)
        {

            decimal decReturn = 0;
            DateTime startDate;
            DateTime endDate;

            // Set the end dates for the Reporting Period.  Should start on Sunday and end on Saturday
            startDate = new DateTime(YearForAnnual, 1, 1);
            endDate = new DateTime(YearForAnnual, 12, 31);

            // Loop through the stories
            foreach (Defect defect in defects)
            {
                foreach (Task task in defect.Tasks)
                {
                    if ((task.LastUpdate >= startDate) && (task.LastUpdate <= endDate))
                    {
                        switch (Action)
                        {
                            case ProjectTotal.Estimate:
                                decReturn = decReturn + task.Estimate;
                                break;
                            case ProjectTotal.ToDo:
                                decReturn = decReturn + task.ToDo;
                                break;
                            case ProjectTotal.Actual:
                                decReturn = decReturn + task.Actual;
                                break;
                            default:
                                break;
                        }
                    }
                }

            }

            return decReturn;

        }

        /// <summary>
        /// Accepts list of all User Stories / Defects and calculates weekly totals
        /// </summary>
        /// <param name="story">List of Defects to calculate</param>
        /// <param name="stories">List of Stories to calculate</param>
        /// <param name="Action">Indicates whether to total for Estimate, ToDo, or Actuals</param>
        /// <param name="YearForAnnual">Which Year to generate totals for</param>
        private static decimal CreateAnnualTotal(List<UserStory> stories, ProjectTotal Action, int YearForAnnual)
        {

            decimal decReturn = 0;
            DateTime startDate;
            DateTime endDate;

            // Set the end dates for the Reporting Period.  Should start on Sunday and end on Saturday
            startDate = new DateTime(YearForAnnual, 1, 1);
            endDate = new DateTime(YearForAnnual, 12, 31);

            // Loop through the stories
            foreach (UserStory story in stories)
            {
                foreach (Task task in story.Tasks)
                {
                    if ((task.LastUpdate >= startDate) && (task.LastUpdate <= endDate))
                    {
                        switch (Action)
                        {
                            case ProjectTotal.Estimate:
                                decReturn = decReturn + task.Estimate;
                                break;
                            case ProjectTotal.ToDo:
                                decReturn = decReturn + task.ToDo;
                                break;
                            case ProjectTotal.Actual:
                                decReturn = decReturn + task.Actual;
                                break;
                            default:
                                break;
                        }
                    }
                }

            }

            return decReturn;

        }

        /// <summary>
        /// Accepts list of all User Stories / Defects and calculates weekly totals
        /// </summary>
        /// <param name="stories">List of Stories to calculate</param>
        /// <param name="Action">Indicates whether to total for Estimate, ToDo, or Actuals</param>
        /// <param name="DayforWeekly">Date to use for calculating Sunday/Saturday end dates</param>
        private static decimal CreateWeeklyTotal(List<UserStory> stories, ProjectTotal Action, DateTime DayforWeekly)
        {

            decimal decReturn = 0;
            DateTime startDate;
            DateTime endDate;

            // Set the end dates for the Reporting Period.  Should start on Sunday and end on Saturday
            startDate = GetFirstDayOfWeek(DayforWeekly, CultureInfo.CurrentCulture);
            endDate = startDate.AddDays(6);

            // Loop through the stories
            foreach (UserStory story in stories)
            {
                foreach (Task task in story.Tasks)
                {
                    if ((task.LastUpdate >= startDate) && (task.LastUpdate <= endDate))
                    {
                        switch (Action)
                        {
                            case ProjectTotal.Estimate:
                                decReturn = decReturn + task.Estimate;
                                break;
                            case ProjectTotal.ToDo:
                                decReturn = decReturn + task.ToDo;
                                break;
                            case ProjectTotal.Actual:
                                decReturn = decReturn + task.Actual;
                                break;
                            default:
                                break;
                        }
                    }
                }

            }

            return decReturn;

        }

        /// <summary>
        /// Returns the first day of the week that the specified date
        /// is in.
        /// </summary>
        /// <param name="DayInWeek">Date to calculate the start of week for</param>
        /// <param name="cultureInfo">CultureInfo object to use for determining start of week day</param>
        public static DateTime GetFirstDayOfWeek(DateTime DayInWeek, CultureInfo cultureInfo)
        {

            DayOfWeek firstDay = cultureInfo.DateTimeFormat.FirstDayOfWeek;
            DateTime firstDayInWeek = DayInWeek.Date;
            while (firstDayInWeek.DayOfWeek != firstDay)
                firstDayInWeek = firstDayInWeek.AddDays(-1);

            return firstDayInWeek;

        }

        /// <summary>
        /// Checks the supplied Commandline Arguments for valid switches
        /// </summary>
        /// <param name="CommandArgs">String Array of arguments</param>
        public static bool ValidArgs(string[] CommandArgs)
        {

            HashSet<string> possibleArgs = new HashSet<string> { "-D", "-W", "-A", "-Q" };
            HashSet<string> passedArgs = new HashSet<string>();

            foreach (string str in CommandArgs)
            {
                passedArgs.Add(str);
            }

            bool isSubset = !passedArgs.IsSubsetOf(possibleArgs);

            return true;
        }

        /// <summary>
        /// Checks if the supplied date occured during the supplied reporting month.  This is used to determine if the last update for a project occured in the month we are currently reporting.
        /// </summary>
        /// <param name="UpdateDate">LastUpdateDate for the project</param>
        /// <param name="RptMonth">Month we are reporting for</param>
        private static bool IsReportMonth(DateTime UpdateDate, int RptMonth)
        {

            DateTime startOfMonth = new DateTime(DateTime.Today.Year, RptMonth, 1);
            DateTime endOfMonth = new DateTime(DateTime.Today.Year, RptMonth, DateTime.DaysInMonth(DateTime.Today.Year, RptMonth));

            if (UpdateDate >= startOfMonth && UpdateDate <= endOfMonth)
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        /// <summary>
        /// Accepts list of all Epics / User Stories / Defects and creates summary total roll-ups at the Epic level
        /// </summary>
        /// <param name="epics">List of all Epics</param>
        /// <param name="stories">List of all Stories</param>
        /// <param name="defects">List of all Defects</param>
        /// <param name="DateToCalc">Date to use for calculating Sunday/Saturday end dates</param>
        public static List<Epic> CreateWeeklyEpicSummary(List<Epic> epics, List<UserStory> stories, List<Defect> defects, DateTime DateToCalc)
        {

            List<Epic> ReturnEpics = new List<Epic>();

            foreach (Epic epic in epics)
            {
                Epic newepic = new Epic();
                newepic.StoryEstimate = 0;
                newepic.StoryToDo = 0;
                newepic.StoryActual = 0;
                newepic.DefectEstimate = 0;
                newepic.DefectToDo = 0;
                newepic.DefectActual = 0;
                newepic.ActualEndDate = epic.ActualEndDate;
                newepic.Description = epic.Description;
                newepic.FormattedID = epic.FormattedID;
                newepic.Name = epic.Name;
                newepic.Owner = epic.Owner;
                newepic.ParentProject = epic.ParentProject;
                newepic.PlannedEndDate = epic.PlannedEndDate;
                newepic.PlannedStartDate = epic.PlannedStartDate;
                newepic.Release = epic.Release;
                newepic.State = epic.State;
                foreach (UserStory story in stories)
                {
                    // Check if the current epic is parent to the current story
                    if (epic.Name == story.ParentEpic)
                    {
                        // Save the info
                        foreach (Task task in story.Tasks)
                        {
                            newepic.StoryEstimate = newepic.StoryEstimate + task.Estimate;
                            newepic.StoryToDo = newepic.StoryToDo + task.ToDo;
                            newepic.StoryActual = newepic.StoryActual + task.Actual;
                            if (defects != null)
                            {
                                foreach (Defect defect in defects)
                                {
                                    if (defect.ParentStory == story.FormattedID)
                                    {
                                        foreach (Task storytask in defect.Tasks)
                                        {
                                            newepic.DefectEstimate = newepic.DefectEstimate + storytask.Estimate;
                                            newepic.DefectToDo = newepic.DefectToDo + storytask.ToDo;
                                            newepic.DefectActual = newepic.DefectActual + storytask.Actual;
                                        }
                                    }
                                }
                            }
                            newepic.DefectEstimate = newepic.DefectEstimate + 0;
                            newepic.DefectToDo = newepic.DefectToDo + 0;
                            newepic.DefectActual = newepic.DefectActual + 0;
                        }
                    }
                }
                ReturnEpics.Add(newepic);
            }

            return ReturnEpics;

        }

        /// <summary>
        /// Calculates the "Short Month" name from the supplied numeric month
        /// </summary>
        /// <param name="MonthNumber">Number for the month to return the name for (i.e. 'Jan', 'Feb', etc)</param>
        private string MonthName(int MonthNumber)
        {

            switch (MonthNumber)
            {
                case 1:
                    return "Jan";
                case 2:
                    return "Feb";
                case 3:
                    return "Mar";
                case 4:
                    return "Apr";
                case 5:
                    return "May";
                case 6:
                    return "Jun";
                case 7:
                    return "Jul";
                case 8:
                    return "Aug";
                case 9:
                    return "Sep";
                case 10:
                    return "Oct";
                case 11:
                    return "Nov";
                case 12:
                    return "Dec";
                default:
                    return "N/A";
            }

        }

        /// <summary>
        /// Generates a view of User Stories for a project rolled up to the story level
        /// </summary>
        /// <param name="stories">Original list of User Stories for project</param>
        /// <param name="DateToCalc">Date to create rollup for</param>
        public static List<UserStory> CreateDailyRollup(List<UserStory> stories, DateTime DateToCalc)
        {

            List<UserStory> rollup = new List<UserStory>();

            foreach (UserStory story in stories)
            {
                UserStory newstory = new UserStory();
                newstory.Blocked = story.Blocked;
                newstory.Children = story.Children;
                newstory.Description = story.Description;
                newstory.FormattedID = story.FormattedID;
                newstory.Iteration = story.Iteration;
                newstory.Name = story.Name;
                newstory.Owner = story.Owner;
                newstory.ParentEpic = story.ParentEpic;
                newstory.ParentProject = story.ParentProject;
                newstory.Release = story.Release;
                newstory.State = story.State;
                newstory.Tasks = story.Tasks;
                newstory.StoryActual = TaskTotal(story.Tasks, ProjectTotal.Actual, DateToCalc);
                newstory.StoryEstimate = TaskTotal(story.Tasks, ProjectTotal.Estimate, DateToCalc);
                newstory.StoryToDo = TaskTotal(story.Tasks, ProjectTotal.ToDo, DateToCalc);
                newstory.PlanEstimate = story.PlanEstimate;
                if (newstory.StoryActual != 0)
                {
                    rollup.Add(newstory);
                }
            }

            return rollup;

        }

        /// <summary>
        /// Generates a view of User Stories for a project rolled up to the story level
        /// </summary>q
        /// <param name="defects">Original list of Defects for project</param>
        /// <param name="DateToCalc">Date to create rollup for</param>
        public static List<Defect> CreateDailyRollup(List<Defect> defects, DateTime DateToCalc)
        {

            List<Defect> rollup = new List<Defect>();

            foreach (Defect defect in defects)
            {
                Defect newdefect = new Defect();
                newdefect.Description = defect.Description;
                newdefect.FormattedID = defect.FormattedID;
                newdefect.Iteration = defect.Iteration;
                newdefect.Name = defect.Name;
                newdefect.Owner = defect.Owner;
                newdefect.ParentStory = defect.ParentStory;
                newdefect.Release = defect.Release;
                newdefect.State = defect.State;
                newdefect.Tasks = defect.Tasks;
                newdefect.DefectActual = TaskTotal(defect.Tasks, ProjectTotal.Actual, DateToCalc);
                newdefect.DefectEstimate = TaskTotal(defect.Tasks, ProjectTotal.Estimate, DateToCalc);
                newdefect.DefectToDo = TaskTotal(defect.Tasks, ProjectTotal.ToDo, DateToCalc);
                if (newdefect.DefectActual != 0)
                {
                    rollup.Add(newdefect);
                }
            }

            return rollup;

        }

        /// <summary>
        /// Accepts User Story and calculates daily totals
        /// </summary>
        /// <param name="tasks">Story to calculate</param>
        /// <param name="Action">Indicates whether to total for Estimate, ToDo, or Actuals</param>
        /// <param name="DayforDaily">Date to calculate totals for</param>
        private static decimal TaskTotal(List<Task> tasks, ProjectTotal Action, DateTime DayforDaily)
        {

            decimal decReturn = 0;

            foreach (Task task in tasks)
            {
                if (DayforDaily.ToShortDateString() == task.LastUpdate.ToShortDateString())
                {
                    switch (Action)
                    {
                        case ProjectTotal.Estimate:
                            decReturn = decReturn + task.Estimate;
                            break;
                        case ProjectTotal.ToDo:
                            decReturn = decReturn + task.ToDo;
                            break;
                        case ProjectTotal.Actual:
                            decReturn = decReturn + task.Actual;
                            break;
                        default:
                            break;
                    }

                }
            }

            return decReturn;

        }
    }
}
