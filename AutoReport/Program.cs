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
            public List<UserStory> UserStories;
            public List<Defect> Defects;
            public decimal StoryEstimate;
            public decimal StoryToDo;
            public decimal StoryActual;
            public decimal DefectEstimate;
            public decimal DefectToDo;
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
            public List<Task> Tasks;
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
            public List<Task> Tasks;
        }

        public enum ProjectTotal { Estimate = 1, ToDo, Actual };

        static void Main(string[] args)
        {

            // Login credentials for Rally
            string strRallyUID = "james.drager@vce.com";
            string strRallyPWD = "01DeadCat";

            DisplayOutput("Started processing at " + DateTime.Now.ToLongTimeString() + " on " + DateTime.Now.ToLongDateString());
            DateTime dtStartTime = DateTime.Now;
            // Create the Rally API object
            RallyAPI = new RallyRestApi();

            // Login to Rally
            // By leaving the server identifier off, it uses the default server
            RallyAPI.Authenticate(strRallyUID, strRallyPWD);
            DisplayOutput("Connected to Rally...");

            // Grab the active Initiatives we want to report on
            DisplayOutput("Getting all Initiatives...");
            List<Initiative> InitiativeList = new List<Initiative>();
            InitiativeList = GetInitiativeList();
            if (InitiativeList.Count == 0)
            {
                // Could not get the Initiatives...or nothing to report on, so stop
                DisplayOutput("Unable to open Initiative list or no Initiatives to report on.  Application will terminate.");
                RallyAPI.Logout();
                return;
            }
            DisplayOutput("Retrieved " + InitiativeList.Count + " Initiatives to report on");

            // Now iterate through the initiatives and get all the Features or "Projects"
            DisplayOutput("Getting all Projects for the Initiatives...");
            BasicProjectList = new List<Project>();
            foreach (Initiative init in InitiativeList)
            {
                // Get the project list for the current initiative ONLY
                List<Project> ProjectList = new List<Project>();
                ProjectList = GetProjectsForInitiative(init.Name.Trim());

                // Append this list to the FULL list
                BasicProjectList.AddRange(ProjectList);
            }
            DisplayOutput("Retrieved " + BasicProjectList.Count + " Projects");

            // We need to loop through the project list now and for each project
            // we need to get all the epics.  Then with each epic, we recursively
            // get all user stories
            DisplayOutput("Getting all User Stories for all projects...");
            // Initialize a new list of projects.  This will become the full list including stories
            // and defects
            CompleteProjectList = new List<Project>();
            foreach (Project proj in BasicProjectList)
            {
                // Get all the epics for this project
                DisplayOutput("~+~+~+~+~+~+~+~+~+~+~+~+~+~+~+~+~+~+");
                DisplayOutput("Working on Project " + proj.Name + "...");
                List<Epic> EpicList = new List<Epic>();
                EpicList = GetEpicsForProject(proj.FormattedID.Trim());

                // Now go through each of the Epics for the current project
                // and recurse through to get all final-level user stories
                DisplayOutput("Getting all User Stories for " + proj.Name + "...");
                BasicStoryList = new List<UserStory>();
                foreach (Epic epic in EpicList)
                {
                    List<UserStory> StoryList = new List<UserStory>();
                    StoryList = GetUserStoriesPerParent(epic.FormattedID.Trim(), true);
                    BasicStoryList.AddRange(StoryList);
                }
                DisplayOutput("Retrieved " + BasicStoryList.Count + " User Stories");

                // Get any defects there may be for the User Stories
                DisplayOutput("Getting all Defects for " + proj.Name + "...");
                BasicDefectList = new List<Defect>();
                foreach (UserStory story in BasicStoryList)
                {
                    List<Defect> DefectList = new List<Defect>();
                    DefectList = GetDefectsForStory(story);
                    if (DefectList.Count > 0)
                    {
                        BasicDefectList.AddRange(DefectList);
                    }
                }
                DisplayOutput("Retrieved " + BasicDefectList.Count + " Defects");

                // At this point we have the FULL list of User Stories/Defects for the current
                // project.  We now create a "new" project with the same properties, but this time
                // we are able to store the User Stories and Defects.
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
                CompleteProjectList.Add(newproject);
            }

            // We now have a full list of all projects with all complete information so
            // at this point we need to total the Actuals for each project
            AllProjectInfo = new List<Project>();
            AllProjectInfo = CalculateTotals(CompleteProjectList);

            // Now create the final report
            CreateReport(AllProjectInfo);

            //CompleteProjectList[0].UserStories[0].Tasks[0].Actual
            DateTime dtEndTime = DateTime.Now;
            string strTotSeconds = dtEndTime.Subtract(dtStartTime).TotalSeconds.ToString();
            DisplayOutput("Completed processing at " + DateTime.Now.ToLongTimeString() + " on " + DateTime.Now.ToLongDateString() + " - Total Processing Time: " + strTotSeconds + " seconds");

        }

        /// <summary>
        /// This reads the Initiatives that we want to report on from a text file
        /// A text file is used rather than parsing all Initiatives for OCTO so
        /// that we can report on only the ones we want
        /// </summary>
        public static List<Initiative> GetInitiativeList()
        {

            List<Initiative> listReturn = new List<Initiative>();
            Int32 BufferSize = 128; // Setting the buffer size can spead up reading

            // Full path and filename to read the Initiative list from
            string strInitiativeList = AppDomain.CurrentDomain.BaseDirectory + "ReportList.txt";

            try
            {
                // Open the file
                using (var fileStream = System.IO.File.OpenRead(strInitiativeList))
                using (var streamReader = new System.IO.StreamReader(fileStream, Encoding.UTF8, true, BufferSize))
                {
                    String line;    // Input buffer
                    while ((line = streamReader.ReadLine()) != null)
                    {
                        // Grab all Information for the given initiative
                        Request rallyRequest = new Request("PortfolioItem/Initiative");
                        rallyRequest.Fetch = new List<string>() { "Name", "FormattedID", "Owner", "PlannedStartDate", "PlannedEndDate", "Description", "State" };
                        rallyRequest.Query = new Query("FormattedID", Query.Operator.Equals, line.Trim());
                        QueryResult rallyResult = RallyAPI.Query(rallyRequest);
                        foreach (var result in rallyResult.Results)
                        {
                            // Create and populate the Initiative object
                            Initiative initiative = new Initiative();
                            initiative.FormattedID = RationalizeData(result["FormattedID"]);
                            initiative.Name = RationalizeData(result["Name"]);
                            initiative.Owner = RationalizeData(result["Owner"]);
                            initiative.PlannedStartDate = Convert.ToDateTime(result["PlannedStartDate"]);
                            initiative.PlannedEndDate = Convert.ToDateTime(result["PlannedEndDate"]);
                            initiative.Description = RationalizeData(result["Description"]);
                            initiative.State = RationalizeData(result["State"]);
                            listReturn.Add(initiative);
                        }
                    }
                    streamReader.Close();   // Clean up the reader
                    streamReader.Dispose();
                    fileStream.Close(); // Clean up the file stream
                    fileStream.Dispose();
                }

            }
            catch (Exception ex)
            {
                DisplayOutput("Exception caught while attempting to open Report List: " + ex.Message);
                //throw;
            }
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
        /// <param name="StoryID">FormattedID of the User Story</param>
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
            rallyRequest.Fetch = new List<string>() { "Name", "FormattedID", "Release", "Iteration",
                "Owner", "ScheduleState", "DirectChildrenCount", "Description", "PlanEstimate" };
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
                listReturn.Add(story);
                DisplayOutput(story.FormattedID + " " + story.Name);

                // Check for children.  If there are children, then we need to drill down
                // through all children to retrieve the full list of stories
                if (story.Children > 0)
                {
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
        private static void DisplayOutput(string Message)
        {

            Console.WriteLine(Message);

        }

        /// <summary>
        /// Creates the final output report using the supplied project list
        /// </summary>
        /// <param name="projects">Project list, with calculations, to use for report</param>
        private static void CreateReport(List<Project> projects)
        {

            string strOutLine = "";

            // Full path and filename to read the Initiative list from
            string strOutputReport = AppDomain.CurrentDomain.BaseDirectory + "Output.txt";
            //string outputline = "First line.\r\nSecond line.\r\nThird line.";

            // Write the string to a file.
            System.IO.StreamWriter reportfile = new System.IO.StreamWriter(strOutputReport);

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
            }
            reportfile.Close();

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
    }
}
