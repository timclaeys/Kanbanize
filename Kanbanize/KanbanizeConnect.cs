using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.IO;
using System.Web;
using RestSharp;
using System.Xml.Serialization;
using System.Security;
using RestSharp.Contrib;

namespace Kanbanize
{
    class KanbanizeConnect
    {
        #region api urls
        const string KanbanAdresLogin = "http://kanbanize.com/index.php/api/kanbanize/login/email/";
        const string KanbanAdresProjAndBoards = "http://kanbanize.com/index.php/api/kanbanize/get_projects_and_boards/";
        const string KanbanAdresGetBoardStructure = "http://kanbanize.com/index.php/api/kanbanize/get_board_structure/";
        const string KanbanAdresGetAllTasks = "http://kanbanize.com/index.php/api/kanbanize/get_all_tasks/";
        const string kanbanAdresGetTaskDetails = "http://kanbanize.com/index.php/api/kanbanize/get_task_details";
        #endregion 
        
        #region variables
        public string apikey { get; set; }
        public string email { get; set; }
        public string password { get; set; }
        #endregion

        #region constructors
        public KanbanizeConnect()
        {
            apikey = null;
            email = null;
            password = null;
        }
        public KanbanizeConnect(string emailIn, string passwordIn)//most recommended for using
        {
            email = emailIn;
            password = passwordIn;
        } 
        public KanbanizeConnect(string apikeyIn, string emailIn, string passwordIn)
        {
            apikey = apikeyIn;
            email = emailIn;
            password = passwordIn;

        }
        #endregion

        #region connections to api
        public Login GetLogin()
        {
            string adres = KanbanAdresLogin + HttpUtility.UrlEncode(email) + "/pass/" + password;
            
            IRestResponse response = GetResponseLogin(adres);                                    //execute the request

            if(response.Content.Contains("Invalid"))
            {
                return null;
            }
            else
            {                
                string content = response.Content;                                                      //change the content of the respons
                content = content.Replace("<xml>", "<login>");
                content = content.Replace("</xml>", "</login>");

                XmlSerializer deserializer = new XmlSerializer(typeof(Login));                          //deserialize the content to login object
                object obj = deserializer.Deserialize(StringToStream(content));
                Login XmlData = (Login)obj;
                return XmlData;
            }

        }

        public Projects GetProjectsAndBoards()
        {
            IRestResponse response = GetResponse(KanbanAdresProjAndBoards);

            string content = response.Content;
            content = content.Replace("<xml>", "");
            content = content.Replace("</xml>", "");
            string replacement = "";
            bool inBoard = false;
            for (int i = 0; i < content.Length; i++)
            {
                if (i < content.Length - 9)
                {
                    if (content.Substring(i, 8) == "<boards>")
                    {
                        inBoard = true;
                        i += 7;
                    }
                    if (content.Substring(i, 9) == "</boards>")
                    {
                        inBoard = false;
                        i += 8;
                    }
                }
                else
                {
                    inBoard = false;
                }

                if (i < content.Length - 7)
                {
                    if (inBoard)
                    {
                        if (content.Substring(i, 6) == "<item>")
                        {
                            replacement += "<boardid>";
                            i += 5;
                        }
                        else if (content.Substring(i, 7) == "</item>")
                        {
                            replacement += "</boardid>";
                            i += 6;
                        }
                        else
                        {
                            replacement += content.Substring(i, 1);
                        }
                    }
                    else
                    {
                        if (content.Substring(i, 6) == "<item>")
                        {
                            replacement += "<project>";
                            i += 5;
                        }
                        else if (content.Substring(i, 7) == "</item>")
                        {
                            replacement += "</project>";
                            i += 6;
                        }
                        else
                        {
                            replacement += content.Substring(i, 1);
                        }
                    }
                }
                else
                    replacement += content.Substring(i, 1);
            }

            XmlSerializer deserializer = new XmlSerializer(typeof(Projects));                          //deserialize the content to login object
            object obj = deserializer.Deserialize(StringToStream(replacement));
            Projects XmlData = (Projects)obj;
            return XmlData;
        }

        public Board GetTasks(string boardId, bool enableSubTasks)
        {
            string adres = KanbanAdresGetAllTasks + "boardid/" + boardId;
            if (enableSubTasks )
                adres += "/subtasks/yes";
            IRestResponse response = GetResponse(adres);

            string content = response.Content;                                                      //change the content of the respons
            content = content.Replace("<xml>", "<board>");
            content = content.Replace("</xml>", "</board>");
            string replacement = "";
            bool inSubtasks = false;
            for (int i = 0; i < content.Length; i++)
            {
                if(i < content.Length-17)
                {
                    if (content.Substring(i, 16) == "<subtaskdetails>")
                    {
                        inSubtasks = true;
                        i += 15;
                    }
                    if (content.Substring(i, 17) == "</subtaskdetails>")
                    {
                        inSubtasks = false;
                        i += 16;
                    }
                }
                else
                {
                    inSubtasks = false;                 
                }

                if (i < content.Length - 7)
                {
                    if (inSubtasks)
                    {
                        if (content.Substring(i, 6) == "<item>")
                        {
                            replacement += "<subtask>";
                            i += 5;
                        }
                        else if (content.Substring(i, 7) == "</item>")
                        {
                            replacement += "</subtask>";
                            i += 6;
                        }
                        else
                        {
                            replacement += content.Substring(i, 1);
                        }
                    }
                    else
                    {
                        if (content.Substring(i, 6) == "<item>")
                        {
                            replacement += "<task>";
                            i += 5;
                        }
                        else if (content.Substring(i, 7) == "</item>")
                        {
                            replacement += "</task>";
                            i += 6;
                        }
                        else
                        {
                            replacement += content.Substring(i, 1);
                        }
                    }
                }
                else
                    replacement += content.Substring(i, 1);
            }

            XmlSerializer deserializer = new XmlSerializer(typeof(Board ));                          //deserialize the content to login object
            object obj = null;
            try
            {
                obj = deserializer.Deserialize(StringToStream(replacement));
            }
            catch
            {
                obj = new Board();
            }
            Board XmlData = (Board)obj;
            return XmlData;
        }
        public Board GetTasks(string boardId, bool enableSubTasks, string fromDate, string toDate)
        {
            string adres = KanbanAdresGetAllTasks + "boardid/" + boardId;
            if (enableSubTasks)
                adres += "/subtasks/yes";
            adres += "/container/archive";
            adres += "/fromdate/" + fromDate;
            adres += "/todate/" + toDate;

            IRestResponse response = GetResponse(adres);

            string content = response.Content;                                                      //change the content of the respons
            content = content.Replace("<xml>", "<board>");
            content = content.Replace("</xml>", "</board>");
            string replacement = "";
            bool inSubtasks = false;
            for (int i = 0; i < content.Length; i++)
            {
                if (i < content.Length - 17)
                {
                    if (content.Substring(i, 16) == "<subtaskdetails>")
                    {
                        inSubtasks = true;
                        i += 15;
                    }
                    if (content.Substring(i, 17) == "</subtaskdetails>")
                    {
                        inSubtasks = false;
                        i += 16;
                    }
                }
                else
                {
                    inSubtasks = false;
                }

                if (i < content.Length - 7)
                {
                    if (inSubtasks)
                    {
                        if (content.Substring(i, 6) == "<item>")
                        {
                            replacement += "<subtask>";
                            i += 5;
                        }
                        else if (content.Substring(i, 7) == "</item>")
                        {
                            replacement += "</subtask>";
                            i += 6;
                        }
                        else
                        {
                            replacement += content.Substring(i, 1);
                        }
                    }
                    else
                    {
                        if (content.Substring(i, 6) == "<item>")
                        {
                            replacement += "<task>";
                            i += 5;
                        }
                        else if (content.Substring(i, 7) == "</item>")
                        {
                            replacement += "</task>";
                            i += 6;
                        }
                        else
                        {
                            replacement += content.Substring(i, 1);
                        }
                    }
                }
                else
                    replacement += content.Substring(i, 1);
            }

            XmlSerializer deserializer = new XmlSerializer(typeof(Board));                          //deserialize the content to login object
            object obj = null;
            try
            {
                obj = deserializer.Deserialize(StringToStream(replacement));
            }
            catch
            {
                obj = new Board();
            }
            Board XmlData = (Board)obj;
            return XmlData;
        }

        public HistoryDetails GetTaskDetails(String BoardId,String TaskId)
        {
            string adres = kanbanAdresGetTaskDetails + "/boardid/" + BoardId + "/taskid/" + TaskId + "/history/yes";
            IRestResponse response = GetResponse(adres);
            string content = response.Content;
            string replacement = "<?xml version=\"1.0\" encoding=\"utf-8\" ?> ";
            bool inHistory = false;


            for (int i = 0; i < content.Length-17; i++)
            {
                
                if (content.Substring(i, 16) == "<historydetails>")
                {
                    inHistory = true;
                    i += 15;
                    replacement += "<historydetails>";
                }
                else if (content.Substring(i, 17) == "</historydetails>")
                {
                    inHistory = false;
                    i += 16;
                    replacement += "</historydetails>";
                }
                else
                {
                    if (inHistory)
                    {
                        if (content.Substring(i, 6) == "<item>")
                        {
                            replacement += "<taskEvent>";
                            i += 5;
                        }
                        else if (content.Substring(i, 7) == "</item>")
                        {
                            replacement += "</taskEvent>";
                            i += 6;
                        }
                        else
                        {
                            replacement += content.Substring(i, 1);
                        }
                    }
                }
            }
            XmlSerializer deserializer = new XmlSerializer(typeof(HistoryDetails));                          //deserialize the content to login object
            object obj = obj = deserializer.Deserialize(StringToStream(replacement));           
            HistoryDetails XmlData = (HistoryDetails)obj;


            



            return XmlData;

            
        }
        #endregion

        #region private methods
        private MemoryStream StringToStream(string input)
        {
            byte[] byteArray = Encoding.ASCII.GetBytes(input);
            MemoryStream stream = new MemoryStream(byteArray);
            return stream;
        }

        private IRestResponse GetResponse(string adres)
        {
            var client = new RestClient(adres);
            var request = new RestRequest(Method.POST);
            request.AddHeader("apikey", apikey);
            return client.Execute(request);
        }

        private IRestResponse GetResponseLogin(string adres)
        {
            var client = new RestClient(adres);
            var request = new RestRequest(Method.POST);
            return client.Execute(request);
        }
        #endregion

    }

    #region Classes related to xml input from api
    #region login
    public class Login
    {
        public string email { get; set; }
        public string username { get; set; }
        public string realname { get; set; }
        public string companyname { get; set; }
        public string timeZone { get; set; }
        public string apikey { get; set; }

    }
    #endregion
    #region board
    public class Board
    {
        [XmlElement("task")]
        public List<Task> tasklist = new List<Task>();
    }

    public class Task
    {
        public string taskid { get; set; }
        public string position { get; set; }
        public string type { get; set; }
        public string assignee { get; set; }
        public string title { get; set; }
        public string description { get; set; }
        public string subtasks { get; set; }
        public string subtaskscomplete { get; set; }
        public string color { get; set; }
        public string priority { get; set; }
        public string size { get; set; }
        public string deadline { get; set; }
        public string deadlineoriginalformat { get; set; }
        public string extlink { get; set; }
        public string tags { get; set; }
        public string columnid { get; set; }
        public string laneid { get; set; }
        public string leadtime { get; set; }
        public string blocked { get; set; }
        public string blockedreason { get; set; }
        public string columnname{ get; set; }
        public string lanename { get; set; }
        public string columnpath { get; set; }
        public string loggedtime { get; set; }
        [XmlElement("subtask")]
        public List<SubTask> subtasklist= new List<SubTask>();
    }

    public class SubTask
    {
        public string subtaskid { get; set; }
        public string assignee { get; set; }
        public string title { get; set; }
        public string completiondate { get; set; }
    }
    #endregion
    #region projects
    public class Projects
    {
        [XmlElement("project")]
        public List<Project> projectlist = new List<Project>();
        
    }

    public class Project
    {
        public string name { get; set; }
        public string id { get; set; }
        [XmlElement("boardid")]
        public List<Boardid> boardlist = new List<Boardid>();
    }

    public class Boardid 
    {
        public string name { get; set; }
        public string id { get; set; }
    }
    #endregion
    #region history
    public class HistoryDetails
    {
        [XmlElement("taskEvent")]
        public List<TaskEvent> eventList = new List<TaskEvent>();
    }

    public class TaskEvent
    {
        public string eventtype { get; set; }
        public string historyevent { get; set; }
        public string details { get; set; }
        public string author { get; set; }
        public string entrydate { get; set; }
        public string historyid { get; set; }
    }
    #endregion
    #endregion

}
