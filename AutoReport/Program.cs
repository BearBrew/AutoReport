using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rally.RestApi;
using Rally.RestApi.Auth;
using Rally.RestApi.Connection;
using Rally.RestApi.Exceptions;
using Rally.RestApi.Json;
using Rally.RestApi.Response;
using System.Configuration;
//using Excel = Microsoft.Office.Interop.Excel;
//using Word = Microsoft.Office.Interop.Word;
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
        /// Configuration item for UNC Report Location
        /// </summary>
        public static string ConfigReportPath;
        /// <summary>
        /// Configuration item to enable full debug
        /// </summary>
        public static bool ConfigDebug;
        public static string ReportFile;
        public static string LogFile;
        public static string ConfigDBServer;
        public static string ConfigDBName;
        public static string ConfigDBUID;
        public static string ConfigDBPWD;

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
            /// Planned Start date of the Project
            /// </summary>
            public System.DateTime PlannedStartDate;
            /// <summary>
            /// Planned End date of the Project
            /// </summary>
            public System.DateTime PlannedEndDate;
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
            /// Parent Initiative for the Project
            /// </summary>
            public string Initiative;
            /// <summary>
            /// Is the project considered a "Top 5"
            /// </summary>
            public bool Expedite;
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
        }

        public enum ProjectTotal { Estimate = 1, ToDo, Actual };

        static void Main(string[] args)
        {

            // Get the configuration information from config file
            if (!GetConfigSettings())
            {
                // If we can't get the configuration settings, we can't even log anything, so just terminate
                return;
            }

            // Check for any commandline arguments
            if (args.Length != 0)
            {
                // 
                GetCommandArgs();
            }

            LogOutput("Started processing at " + DateTime.Now.ToLongTimeString() + " on " + DateTime.Now.ToLongDateString(), "Main", false);
            DateTime dtStartTime = DateTime.Now;
            // Create the Rally API object
            LogOutput("Creating reference to RallyAPI...", "Main", true);
            RallyAPI = new RallyRestApi();

            // Login to Rally
            LogOutput("Starting connection to Rally...", "Main", true);
            // By leaving the server identifier off, it uses the default server
            try
            {
                RallyAPI.Authenticate(ConfigRallyUID, ConfigRallyPWD);
                LogOutput("Response from RallyAPI.Authenticate: " + RallyAPI.AuthenticationState.ToString(), "Main", true);
                if (RallyAPI.AuthenticationState.ToString() != "Authenticated")
                {
                    // We did not actually connect
                    LogOutput("Unable to connect to Rally and establish session.  Application will terminate.", "Main", false);
                    return;
                }
                else
                {
                    if (RallyAPI.ConnectionInfo.UserName == null)
                    {
                        LogOutput("RallyAPI.ConnectionInfo: " + RallyAPI.ConnectionInfo.ToString(), "Main", false);
                        LogOutput("Unable to authenticate with Rally.  Application will terminate.", "Main", false);
                        return;
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
                RallyAPI.Logout();  // Disconnect from Rally
                return;             // End Program
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
                    StoryList = GetUserStoriesPerParent(epic.FormattedID.Trim(), true);
                    LogOutput("Done with 'GetUserStoriesPerParent' for " + epic.FormattedID.Trim(), "Main", true);
                    BasicStoryList.AddRange(StoryList);
                }
                // We don't need the Epics anymore so clean up
                EpicList.Clear();
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
                newproject.PlannedEndDate = proj.PlannedEndDate;
                newproject.PlannedStartDate = proj.PlannedStartDate;
                newproject.RevokedReason = proj.RevokedReason;
                newproject.State = proj.State;
                newproject.StatusUpdate = proj.StatusUpdate;
                newproject.UserStories = BasicStoryList;
                newproject.Defects = BasicDefectList;
                LogOutput("Appending new project object to object 'CompleteProjectList'", "Main", true);
                CompleteProjectList.Add(newproject);
            }
            LogOutput("Done looping through all projects", "Main", false);
            LogOutput("Appending new project object to object 'CompleteProjectList'", "Main", true);

            // We now have a full list of all projects with all complete information so
            // at this point we need to total the Actuals for each project
            AllProjectInfo = new List<Project>();
            LogOutput("Calling 'CalculateTotals' with " + CompleteProjectList.Count + " complete projects...", "Main", true);
            AllProjectInfo = CalculateTotals(CompleteProjectList);
            LogOutput("Done with 'CalculateTotals'", "Main", true);

            // Now create the final report
            LogOutput("Calling 'CreateReport'...", "Main", true);
            CreateReport(AllProjectInfo);
            LogOutput("Done with 'CreateReport'...", "Main", true);

            //CompleteProjectList[0].UserStories[0].Tasks[0].Actual
            DateTime dtEndTime = DateTime.Now;
            string strTotSeconds = dtEndTime.Subtract(dtStartTime).TotalSeconds.ToString();
            LogOutput("Completed processing at " + DateTime.Now.ToLongTimeString() + " on " + DateTime.Now.ToLongDateString() + " - Total Processing Time: " + strTotSeconds + " seconds", "Main", false);

        }

        /// <summary>
        /// This reads the Initiatives that we want to report on from configuration
        /// settings rather than parsing all Initiatives for OCTO so that we can 
        /// report on only the ones we want
        /// </summary>
        public static List<Initiative> GetInitiativeList()
        {

            List<Initiative> listReturn = new List<Initiative>();
            string[] Initiatives = new string[0];

            // Full path and filename to read the Initiative list from
            LogOutput("Getting list of initiatives from 'ConfigurationManager'...", "GetInitiativeList", true);
            string strInitiativeList = ConfigurationManager.AppSettings["ReportList"];
            LogOutput("Full configured list is: " + strInitiativeList, "GetInitiativeList", true);
            System.Text.RegularExpressions.Regex myReg = new System.Text.RegularExpressions.Regex(",");
            Initiatives = myReg.Split(strInitiativeList);
            LogOutput("Looping for each initiative to get full information...", "GetInitiativeList", false);
            foreach (string stringpart in Initiatives)
            {
                // Grab all Information for the given initiative
                LogOutput("Using RallyAPI to request information for " + stringpart + "...", "GetInitiativeList", true);
                LogOutput("Building Rally Request...", "GetInitiativeList", true);
                Request rallyRequest = new Request("PortfolioItem/Initiative");
                rallyRequest.Fetch = new List<string>() { "Name", "FormattedID", "Owner", "PlannedStartDate", "PlannedEndDate", "Description", "State" };
                rallyRequest.Query = new Query("FormattedID", Query.Operator.Equals, stringpart.Trim());
                LogOutput("Running Rally Query request...", "GetInitiativeList", true);
                QueryResult rallyResult = RallyAPI.Query(rallyRequest);
                LogOutput("Looping through Query result...", "GetInitiativeList", true);
                foreach (var result in rallyResult.Results)
                {
                    // Create and populate the Initiative object
                    LogOutput("Creating new Initiative object and saving...", "GetInitiativeList", true);
                    Initiative initiative = new Initiative();
                    initiative.FormattedID = RationalizeData(result["FormattedID"]);
                    initiative.Name = RationalizeData(result["Name"]);
                    initiative.Owner = RationalizeData(result["Owner"]);
                    initiative.PlannedStartDate = Convert.ToDateTime(result["PlannedStartDate"]);
                    initiative.PlannedEndDate = Convert.ToDateTime(result["PlannedEndDate"]);
                    initiative.Description = RationalizeData(result["Description"]);
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
            Request rallyRequest = new Request("PortfolioItem/Feature");
            rallyRequest.Fetch = new List<string>() { "Name", "FormattedID", "Owner", "PlannedStartDate", "PlannedEndDate",
                "c_ProjectUpdate", "c_RevokedReason", "State", "ValueScore", "Description", "Expedite"};
            rallyRequest.Query = new Query("Parent.Name", Query.Operator.Equals, InitiativeID);
            QueryResult rallyResult = RallyAPI.Query(rallyRequest);
            foreach (var result in rallyResult.Results)
            {
                // Create and populate the Project object
                Project proj = new Project();
                proj.Name = RationalizeData(result["Name"]);
                proj.FormattedID = RationalizeData(result["FormattedID"]);
                proj.Owner = RationalizeData(result["Owner"]);
                proj.PlannedEndDate = Convert.ToDateTime(result["PlannedEndDate"]);
                proj.PlannedStartDate = Convert.ToDateTime(result["PlannedStartDate"]);
                proj.StatusUpdate = RationalizeData(result["c_ProjectUpdate"]);
                proj.RevokedReason = RationalizeData(result["c_RevokedReason"]);
                proj.OpportunityAmount = RationalizeData(result["ValueScore"]);
                proj.State = RationalizeData(result["State"]);
                proj.Description = RationalizeData(result["Description"]);
                proj.Expedite = RationalizeData(result["Expedite"]);
                proj.Initiative = InitiativeID;
                listReturn.Add(proj);
            }

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
            Request rallyRequest = new Request("PortfolioItem/Epic");
            rallyRequest.Fetch = new List<string>() { "Name", "FormattedID", "Owner", "PlannedEndDate",
                "PlannedStartDate", "State", "Description", "Release" };
            rallyRequest.Query = new Query("Parent.FormattedID", Query.Operator.Equals, ProjectID);
            QueryResult rallyResult = RallyAPI.Query(rallyRequest);
            foreach (var result in rallyResult.Results)
            {
                Epic epic = new Epic();
                epic.Name = RationalizeData(result["Name"]);
                epic.FormattedID = RationalizeData(result["FormattedID"]);
                epic.Owner = RationalizeData(result["Owner"]);
                epic.PlannedEndDate = Convert.ToDateTime(result["PlannedEndDate"]);
                epic.PlannedStartDate = Convert.ToDateTime(result["PlannedStartDate"]);
                epic.State = RationalizeData(result["State"]);
                epic.Description = RationalizeData(result["Description"]);
                epic.ParentProject = ProjectName(ProjectID);
                epic.Release = RationalizeData(result["State"]);
                listReturn.Add(epic);
            }

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
            Request rallyRequest = new Request("Defect");
            rallyRequest.Fetch = new List<string>() { "Name", "FormattedID", "Description", "Iteration", "Owner", "Release", "State" };
            rallyRequest.Query = new Query("Requirement.Name", Query.Operator.Equals, Parent.Name.Trim());
            QueryResult rallyResult = RallyAPI.Query(rallyRequest);
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
                listReturn.Add(defect);
            }

            return listReturn;

        }

        /// <summary>
        /// Grabs all User Stories for the indicated parent.  The parent can be an Epic or another User Story
        /// </summary>
        /// <param name="ParentID">Formatted ID of the Parent</param>
        /// <param name="RootParent">Indicates if the parent is the top-level User Story or not.  If looking for sub-stories, things are handled differently</param>
        public static List<UserStory> GetUserStoriesPerParent(string ParentID, bool RootParent)
        {

            List<UserStory> listReturn = new List<UserStory>();

            // Grab all UserStories under the given Epic
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
            QueryResult rallyResult = RallyAPI.Query(rallyRequest);
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
                story.ParentProject = "";
                story.PlanEstimate = RationalizeData(result["PlanEstimate"], true);
                story.Tasks = GetTasksForUserStory(story.FormattedID);
                story.Blocked = RationalizeData(result["Blocked"]);
                listReturn.Add(story);
                LogOutput(story.FormattedID + " " + story.Name, "GetUserStoriesPerParent", true);

                // Check for children.  If there are children, then we need to drill down
                // through all children to retrieve the full list of stories
                if (story.Children > 0)
                {
                    // Recursively get child stories until we reach the lowest level
                    listReturn.AddRange(GetUserStoriesPerParent(story.FormattedID.Trim(), false));
                }
            }

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
            Request rallyRequest = new Request("Tasks");
            rallyRequest.Fetch = new List<string>() { "Name", "FormattedID", "Estimate", "ToDo", "Actuals", "Owner", "State", "Description" };
            rallyRequest.Query = new Query("WorkProduct.FormattedID", Query.Operator.Equals, StoryID);
            QueryResult rallyResult = RallyAPI.Query(rallyRequest);
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
                listReturn.Add(task);
            }

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

            System.IO.StreamWriter logfile = new System.IO.StreamWriter(LogFile, true);

            // Check if we are in Debug mode or not
            switch (ConfigDebug)
            {
                case true:
                    // Full debug
                    string strOutput = System.DateTime.Now.ToString("hh:mm:ss.ffffff") + "\tCURRENT MODULE: " + Method + "\tMESSAGE: " + Message;
                    Console.WriteLine(strOutput);
                    logfile.WriteLine(strOutput);
                    break;
                default:
                    // Basic logging
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
        /// Creates the final output report using the supplied project list
        /// </summary>
        /// <param name="projects">Project list, with calculations, to use for report</param>
        private static void CreateReport(List<Project> projects)
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
            }
            catch (SqlException ex)
            {
                LogOutput("Error connecting to SQL Server:" + ex.Message, "CreateReport", false);
                bDBConnected = false;
            }

            // Write the string to a file.
            System.IO.StreamWriter reportfile = new System.IO.StreamWriter(ReportFile);

            foreach (Project proj in projects)
            {
                strOutLine = "Final Totals for: " + proj.Name + "\r\n";
                strOutLine = strOutLine + "   Total Estimate for Stories --> " + proj.StoryEstimate + "\r\n";
                strOutLine = strOutLine + "   Total ToDo for Stories --> " + proj.StoryToDo + "\r\n";
                strOutLine = strOutLine + "   Total Actual for Stories --> " + proj.StoryActual + "\r\n";
                strOutLine = strOutLine + "   Total Estimate for Defects --> " + proj.DefectEstimate + "\r\n";
                strOutLine = strOutLine + "   Total ToDo for Defects --> " + proj.DefectToDo + "\r\n";
                strOutLine = strOutLine + "   Total Actual for Defects --> " + proj.DefectActual + "\r\n";
                strOutLine = strOutLine + " ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" + "\r\n";
                reportfile.WriteLine(strOutLine);

                if (bDBConnected)
                {
                    SqlCommand sqlCmd = new SqlCommand("INSERT INTO dbo.DailyStats VALUES(@Initiative, @ProjectName, @OppAmount, @Owner, @PlannedStartDate, " +
                            "@PlannedEndDate, @Expedite, @State, @CancelledReason, @HoursDefectsEstimate, @HoursDefectsToDo, @HoursDefectsActual, " +
                            "@HoursStoryEstimate, @HoursStoryToDo, @HoursStoryActual, @ReportDate)", sqlDatabase);
                    sqlCmd.Parameters.Add(new SqlParameter("Initiative", proj.Initiative));
                    sqlCmd.Parameters.Add(new SqlParameter("ProjectName", proj.Name));
                    sqlCmd.Parameters.Add(new SqlParameter("OppAmount", proj.OpportunityAmount));
                    sqlCmd.Parameters.Add(new SqlParameter("Owner", proj.Owner));
                    // Check to make sure that PlannedStartDate is within date limits for SQL Server
                    if (proj.PlannedStartDate < System.Data.SqlTypes.SqlDateTime.MinValue || proj.PlannedStartDate > System.Data.SqlTypes.SqlDateTime.MaxValue)
                    {
                        sqlCmd.Parameters.Add(new SqlParameter("PlannedStartDate", System.Data.SqlTypes.SqlDateTime.MinValue));
                    }
                    else
                    {
                        sqlCmd.Parameters.Add(new SqlParameter("PlannedStartDate", proj.PlannedStartDate));
                    }
                    // Check to make sure that PlannedEndDate is within date limits for SQL Server
                    if (proj.PlannedEndDate < System.Data.SqlTypes.SqlDateTime.MinValue || proj.PlannedEndDate > System.Data.SqlTypes.SqlDateTime.MaxValue)
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
                    sqlCmd.ExecuteNonQuery();
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
        /// Runs through all tasks associated with a project (both User Stories and Defects) and totals the Estimates, ToDo, and Actuals
        /// </summary>
        /// <param name="SourceProjectList">Original Project list to calculate</param>
        private static List<Project> CalculateTotals(List<Project> SourceProjectList)
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
                newproject.PlannedEndDate = proj.PlannedEndDate;
                newproject.PlannedStartDate = proj.PlannedStartDate;
                newproject.RevokedReason = proj.RevokedReason;
                newproject.State = proj.State;
                newproject.StatusUpdate = proj.StatusUpdate;
                newproject.UserStories = proj.UserStories;
                newproject.Defects = proj.Defects;
                newproject.StoryEstimate = CreateTotal(newproject.UserStories, ProjectTotal.Estimate);
                newproject.StoryToDo = CreateTotal(newproject.UserStories, ProjectTotal.ToDo);
                newproject.StoryActual = CreateTotal(newproject.UserStories, ProjectTotal.Actual);
                newproject.DefectEstimate = CreateTotal(newproject.Defects, ProjectTotal.Estimate);
                newproject.DefectToDo = CreateTotal(newproject.Defects, ProjectTotal.ToDo);
                newproject.DefectActual = CreateTotal(newproject.Defects, ProjectTotal.Actual);
                DestinationProjectList.Add(newproject);
            }

            return DestinationProjectList;

        }

        /// <summary>
        /// Accepts list of all User Stories / Defects and calculates totals
        /// </summary>
        /// <param name="stories">List of stories to calculate</param>
        /// <param name="action">Indicates whether to total for Estimate, ToDo, or Actuals</param>
        private static decimal CreateTotal(List<UserStory> stories, ProjectTotal action)
        {

            decimal decReturn = 0;

            foreach (UserStory story in stories)
            {
                foreach (Task task in story.Tasks)
                {
                    switch (action)
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
        /// <summary>
        /// Accepts list of all User Stories / Defects and calculates totals
        /// </summary>
        /// <param name="defects">List of Defects to calculate</param>
        /// <param name="action">Indicates whether to total for Estimate, ToDo, or Actuals</param>
        private static decimal CreateTotal(List<Defect> defects, ProjectTotal action)
        {

            decimal decReturn = 0;

            foreach (Defect defect in defects)
            {
                foreach (Task task in defect.Tasks)
                {
                    switch (action)
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

        /// <summary>
        /// Grabs all configuration information from config file
        /// </summary>
        private static bool GetConfigSettings()
        {

            try
            {
                ConfigRallyUID = ConfigurationManager.AppSettings["RallyUser"];
                ConfigRallyPWD = ConfigurationManager.AppSettings["RallyPass"];
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

            ReportFile = ConfigReportPath + "\\" + System.DateTime.Now.ToString("ddMMMyyyy") + "-" + System.DateTime.Now.ToString("HHmm") + "_Report.txt";
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

            for (int LoopCtl = 0; LoopCtl < CmdArgs.Length; LoopCtl++)
            {
                // 
                switch (CmdArgs[LoopCtl])
                {
                    case "":
                        break;
                    case "1":
                        break;
                    case "2":
                        break;
                    default:
                        break;
                }
            }

        }

    }
}
