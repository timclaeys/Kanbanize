using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Security;

/*some classes are called the same like task and yourTask,
 * these are almost the same classes, but the your... classes are used in this form
 * the not other classes are classes used in the kananize api handler
 */

namespace Kanbanize
{
    public partial class Form1 : Form
    {
        #region variables
        private string email = "";   
        private string password = "";
        int currentRow = 2;                     //current row in the excell file 
        private string currentProjectId = "";   
        private string currentBoardId = "";
        private string currentTaksId = "";
        int numberTasks = 0;                    //number of tasks in total
        private string asked_name = "";         //the searchterm
        #endregion


        #region objects
        KanbanizeConnect kanbanizeApiObj;                                           //connection to the kanbanize api handler
                
        Login yourLogin;                                                            //object of kanbanize api handler
        Projects projecten;                                                         //object of kanbanize api handler
        Board Boards;                                                               //object of kanbanize api handler
        HistoryDetails historyDetails;                                              //object of kanbanize api handler                      

        YourProjects allProjects = new YourProjects();                              //all of the projects,board,tasks
        List<YourTask> foundTasks = new List<YourTask>();                           //tasks that have the asked_name or a part of it in the assignees name
        AutoCompleteStringCollection source = new AutoCompleteStringCollection();   //collection of words that occur in the assignee property, this is used s an autocomplete database for the search term
                        
        Excel.Application xlApp;                                                    //the excell application
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        object misValue = System.Reflection.Missing.Value;
        #endregion

        public Form1()
        {
            InitializeComponent();
            #region declareExcelApplication
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            #endregion
        }

        #region methods

        private void FindTasksForDate(DateTime start, DateTime stop)//find the tasks in progress between the start and stop date
        {
            foundTasks = foundTasks.OrderBy(o => o.assignee).ToList();                                          //order the tasks according to the assignee
            int tasksInProgress = 0;                                                                            //not used now, can be used later on
            currentRow = 2;                                                                                     //the current row in the excel file

            xlWorkSheet.Cells[currentRow, 1] = "Periode";                                                       //assigning the titles in the excel file
            xlWorkSheet.Cells[currentRow - 1, 2] = "Start";
            xlWorkSheet.Cells[currentRow, 2] = start.ToString();
            xlWorkSheet.Cells[currentRow - 1, 3] = "Stop";
            xlWorkSheet.Cells[currentRow, 3] = stop.ToString();
            currentRow++;                                                                                       //next row
            xlWorkSheet.Cells[currentRow, 1] = "name";
            xlWorkSheet.Cells[currentRow, 2] = "Task name";
            xlWorkSheet.Cells[currentRow, 3] = "Task id";
            xlWorkSheet.Cells[currentRow, 4] = "Columnname";
            xlWorkSheet.Cells[currentRow, 5] = "Columnpath";
            xlWorkSheet.Cells[currentRow, 6] = "Lanename";
            xlWorkSheet.Cells[currentRow, 7] = "subtasks";
            xlWorkSheet.Cells[currentRow, 8] = "subtaskscomplete";
            xlWorkSheet.Cells[currentRow, 9] = "Board";
            xlWorkSheet.Cells[currentRow, 10] = "Board id";
            xlWorkSheet.Cells[currentRow, 11] = "Project";
            xlWorkSheet.Cells[currentRow, 12] = "project id";
            currentRow++;                                                                                       //next row
            foreach (YourTask yourTaskBuf in foundTasks)                                                        //for every task in foundtasks
            {
                bool searchForStart = false;                                                                    
                if ((yourTaskBuf.columnname == "Done") || (yourTaskBuf.columnname == "Archived"))               //Task is done
                {
                    foreach (YourTaskEvent yourTaskEventBuf in yourTaskBuf.historyDetails)                      //for every task event of the current task
                    {
                        DateTime test = Convert.ToDateTime(yourTaskEventBuf.entrydate);                         //entrydate for the current task event

                        if ((yourTaskEventBuf.historyevent == "Task moved") && (start <= test) && (yourTaskEventBuf.details.Contains("Done"))) //Task is ended after starting time
                        {
                            if (test <= stop)                                                                   //task is done before the stop time                               
                            {
                                writeLineToExcel(yourTaskBuf);                                                  //write the property's to excel
                                tasksInProgress++;
                                break;
                            }
                            else                                                                                //task ended after stop time
                            { 
                                searchForStart = true;                                                          //still need to look if the task started before the stop time
                            }
                        }

                        if ((searchForStart) && (yourTaskEventBuf.historyevent == "Task moved") && (yourTaskEventBuf.details.Contains("In Progress")) && (test <= stop)) //task is started before the stop time
                        {
                            writeLineToExcel(yourTaskBuf);                                                      //write the property's to excel
                            tasksInProgress++;
                            break;
                        }
                    }
                }
                else if ((yourTaskBuf.columnname == "In Progress") || (yourTaskBuf.columnname == "In Test") || (yourTaskBuf.columnname == "In validation") || (yourTaskBuf.columnname == "In development")||(yourTaskBuf.columnpath=="In Progress"))// de taak is op dit moment nog in progress
                {
                    if (yourTaskBuf.historyDetails.Count == 1)                                                  //if there is only 1 task event and the task is in progress it is created in the "In Progress" Column
                    {
                        writeLineToExcel(yourTaskBuf);                                                          //write the property's to excel
                        tasksInProgress++;
                    }
                    else
                    {
                        if (yourTaskBuf.columnpath == "Backlog")                                                //if the task is in backlog
                        {
                            foreach (YourTaskEvent yourTaskEventBuf in yourTaskBuf.historyDetails)              //for every task event in the current task
                            {
                                DateTime test = Convert.ToDateTime(yourTaskEventBuf.entrydate);                 //the entrydate for the current taskevent
                                if((yourTaskEventBuf.historyevent == "Task moved")&&(start <= test)&&( test <= stop)&&(yourTaskEventBuf.details.Contains("In Progress")))//als in progres was en naar backlog gestuurd geweest binnen de periode.
                                {
                                    writeLineToExcel(yourTaskBuf);                                              //write the property's to excel
                                    tasksInProgress++;
                                    break;
                                }
                            }
                        }
                        else
                        {
                            foreach (YourTaskEvent yourTaskEventBuf in yourTaskBuf.historyDetails)              //for every task event in the current task
                            {
                                DateTime test = Convert.ToDateTime(yourTaskEventBuf.entrydate);                 //the entrydate for the current taskevent
                                if ((yourTaskEventBuf.historyevent == "Task moved") && (test <= stop))          // de taak is voor de stop tijd gemoved naar in progress
                                {
                                    writeLineToExcel(yourTaskBuf);                                              //write the property's to excel
                                    tasksInProgress++;
                                    break;
                                }

                            }
                        }
                    }

                }
            }
            
        }           
        private void writeLineToExcel(YourTask yourTaskBuffer)//write a line to the current excell file with the property's of a task
        {
            xlWorkSheet.Cells[currentRow, 1] = yourTaskBuffer.assignee;                             //write the tasks property's to the excel file
            xlWorkSheet.Cells[currentRow, 2] = yourTaskBuffer.title;
            xlWorkSheet.Cells[currentRow, 3] = yourTaskBuffer.taskid;
            xlWorkSheet.Cells[currentRow, 4] = yourTaskBuffer.columnname;
            xlWorkSheet.Cells[currentRow, 5] = yourTaskBuffer.columnpath;
            xlWorkSheet.Cells[currentRow, 6] = yourTaskBuffer.lanename;
            xlWorkSheet.Cells[currentRow, 7] = yourTaskBuffer.subtasks;
            xlWorkSheet.Cells[currentRow, 8] = yourTaskBuffer.subtaskscomplete;
            foreach (YourProject yourProjectBuf in allProjects.yourProjectList)                     //for every project
            {
                foreach (YourBoard yourBoardBuf in yourProjectBuf.boardList)                        //for every board
                {  
                    foreach (YourTask yourTaskBuf in yourBoardBuf.taskList)                         //for every task
                    {
                        if (yourTaskBuf.taskid == yourTaskBuffer.taskid)                            //if the tasks are the same the board and project of the given task(yourtaskbuffer) are found
                        {
                            xlWorkSheet.Cells[currentRow, 9] = yourBoardBuf.name;
                            xlWorkSheet.Cells[currentRow, 10] = yourBoardBuf.id;
                            xlWorkSheet.Cells[currentRow, 11] = yourProjectBuf.name;
                            xlWorkSheet.Cells[currentRow, 12] = yourProjectBuf.id;
                        }
                    }
                }
            }
            currentRow++;                                                                           //next row
            if (yourTaskBuffer.subtaskList.Count > 0)                                               //if there are subtasks write them to the excell file
            {
                foreach (YourSubTask yoursubtaskBuf in yourTaskBuffer.subtaskList)                  //for every subtask in the current task
                {
                    xlWorkSheet.Cells[currentRow, 1] = yoursubtaskBuf.assignee;                     //write the subtask propertys to the excell file
                    xlWorkSheet.Cells[currentRow, 2] = yoursubtaskBuf.title;
                    xlWorkSheet.Cells[currentRow, 3] = yoursubtaskBuf.subtaskid;
                    xlWorkSheet.Cells[currentRow, 4] = yoursubtaskBuf.completiondate;
                    if (yoursubtaskBuf.completiondate != "")                                        //if the subtask is done
                    {
                        xlWorkSheet.Cells[currentRow, 5] = "DONE";                                  //subtask is done
                    }
                    currentRow++;                                                                   //next row
                }
            }
        }                 
        private void getAllInfo()//download all of the projects, boards and tasks from the website "kanbanize.com" via the api handler
        {
            projecten = kanbanizeApiObj.getProjectsAndBoards();                                         //download the list of projects and boards
            foreach (Project Project in projecten.projectlist)                                          //for every project in this list
            {
                YourProject yourProjectBuf = new YourProject();                                         
                yourProjectBuf.name = Project.name;
                yourProjectBuf.id = Project.id;
                foreach (Boardid boardId in Project.boardlist)                                          //for every board in this list
                {
                    YourBoard yourBoardBuf = new YourBoard();
                    yourBoardBuf.name = boardId.name;
                    yourBoardBuf.id = boardId.id;
                    Boards = kanbanizeApiObj.getTasks(boardId.id, true);                                //download all the tasks for a given board
                    foreach (Task taken in Boards.tasklist)                                             //for every task in this board
                    {
                        YourTask yourTaskBuf = new YourTask();
                        yourTaskBuf.taskid = taken.taskid;
                        yourTaskBuf.position = taken.position;
                        yourTaskBuf.type = taken.type;
                        yourTaskBuf.assignee = taken.assignee;
                        yourTaskBuf.title = taken.title;
                        yourTaskBuf.description = taken.description;
                        yourTaskBuf.subtasks = taken.subtasks;
                        yourTaskBuf.subtaskscomplete = taken.subtaskscomplete;
                        yourTaskBuf.color = taken.color;
                        yourTaskBuf.priority = taken.priority;
                        yourTaskBuf.size = taken.size;
                        yourTaskBuf.deadline = taken.deadline;
                        yourTaskBuf.deadlineoriginalformat = taken.deadlineoriginalformat;
                        yourTaskBuf.extlink = taken.extlink;
                        yourTaskBuf.tags = taken.tags;
                        yourTaskBuf.columnid = taken.columnid;
                        yourTaskBuf.laneid = taken.laneid;
                        yourTaskBuf.leadtime = taken.leadtime;
                        yourTaskBuf.blocked = taken.blocked;
                        yourTaskBuf.blockedreason = taken.blockedreason;
                        yourTaskBuf.columnname = taken.columnname;
                        yourTaskBuf.lanename = taken.lanename;
                        yourTaskBuf.columnpath = taken.columnpath;
                        yourTaskBuf.loggedtime = taken.loggedtime;
                        foreach (SubTask subtasks in taken.subtasklist)                               //for every subtask in the current task
                        {
                            YourSubTask yourSubTaskBuf = new YourSubTask();
                            yourSubTaskBuf.assignee = subtasks.assignee;
                            yourSubTaskBuf.completiondate = subtasks.completiondate;
                            yourSubTaskBuf.subtaskid = subtasks.subtaskid;
                            yourSubTaskBuf.title = subtasks.title;
                            yourTaskBuf.subtaskList.Add(yourSubTaskBuf);
                        }
                        yourBoardBuf.taskList.Add(yourTaskBuf);                                       
                        numberTasks++;
                    }
                    yourProjectBuf.boardList.Add(yourBoardBuf);
                }
                allProjects.yourProjectList.Add(yourProjectBuf);                                    //add  everything to "allProjects"
            }
        }   
        private void releaseObject(object obj)//release an object(used for the excel objects)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        private bool SaveToExcel()//save the excel file
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel files (*.xls)|*.xls";
                saveFileDialog.FilterIndex = 0;
                saveFileDialog.RestoreDirectory = true;
                saveFileDialog.CreatePrompt = true;
                saveFileDialog.FileName = "InProgress_" + asked_name + "_start_" + dateTimePicker1.Value.ToString("yyyy-MM-dd_HH-mm-ss") + "_stop_" + dateTimePicker2.Value.ToString("yyyy-MM-dd_HH-mm-ss");
                saveFileDialog.Title = "Save path of the file to be exported";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Save. The selected path can be got with saveFileDialog.FileName.ToString()
                    xlWorkBook.SaveAs(saveFileDialog.FileName.ToString(), Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                }

                foundTasks.Clear();
                return true;
            }
            catch
            {
                return false;
            }

        }

        #region populate gui
        private void populateLogin()//populate the login labels
        {
            labelUsername.Text = "Username: " + yourLogin.username;
            labelRealName.Text = "Realname: " + yourLogin.realname;
            labelCompanyname.Text = "Companyname: " + yourLogin.companyname;
            labelTimezone.Text = "Timezone: " + yourLogin.timeZone;
        }
        private void populateTreeView()//populate the treeview
        {
            foreach (YourProject yourProjectBuf in allProjects.yourProjectList)                                                                                         //every project (level 1)
            {
                treeView1.Nodes.Add("Project_" + yourProjectBuf.id, yourProjectBuf.name);                                                                               //add node to level 1
                foreach (YourBoard yourBoardBuf in yourProjectBuf.boardList)                                                                                            //every board (level2)
                {
                    treeView1.Nodes["Project_" + yourProjectBuf.id].Nodes.Add("Board_" + yourBoardBuf.id, yourBoardBuf.name);                                           //add node to level 2
                    foreach (YourTask yourTaskBuf in yourBoardBuf.taskList)                                                                                             //every task (level 3)
                    {
                        treeView1.Nodes["Project_" + yourProjectBuf.id].Nodes["Board_" + yourBoardBuf.id].Nodes.Add("Task_" + yourTaskBuf.taskid, yourTaskBuf.title);   //add node to level 3
                    }

                }
            }
        }
        private void populateYourProject(YourProject yourProjectBuf, bool clearLabels)//populate project labels
        {
            labelProjectName.Text = (clearLabels) ? "Project name: " : "Project name: " + yourProjectBuf.name;
            labelProjectId.Text = (clearLabels) ? "Project id: " : "Project id: " + yourProjectBuf.id;
            button2.Enabled = false;
        }
        private void populateYourBoard(YourBoard yourBoardBuf, bool clearLabels)//populate board labels
        {
            labelBoardName.Text = (clearLabels) ? "Board name: " : "Board name: " + yourBoardBuf.name;
            labelBoardId.Text = (clearLabels) ? "Board id: " : "Board id: " + yourBoardBuf.id;

        }
        private void populateYourTask(YourTask yourTaskBuf, bool clearLabels)//populate task labels
        {
            labelTaskTitle.Text = (clearLabels) ? "Task title: " : "Task title: " + yourTaskBuf.title;
            labelTaskId.Text = (clearLabels) ? "Task id: " : "Task id: " + yourTaskBuf.taskid;
            labelTaskPosition.Text = (clearLabels) ? "Position: " : "Position: " + yourTaskBuf.position;
            labelTaksType.Text = (clearLabels) ? "Type: " : "Type: " + yourTaskBuf.type;
            labelTaskAssignee.Text = (clearLabels) ? "Assignee: " : "Assignee: " + yourTaskBuf.assignee;
            labelTaskSubtasks.Text = (clearLabels) ? "Subtasks: " : "Subtasks: " + yourTaskBuf.subtasks;
            labelSubtasksComplete.Text = (clearLabels) ? "Subtasks complete: " : "Subtasks complete: " + yourTaskBuf.subtaskscomplete;
            labelTaskColor.Text = (clearLabels) ? "Color: " : "Color: " + yourTaskBuf.color;
            labelTaskSize.Text = (clearLabels) ? "Size: " : "Size: " + yourTaskBuf.size;
            labelTaskDeadline.Text = (clearLabels) ? "Deadline: " : "Deadline: " + yourTaskBuf.deadline;
            labelTaskDeadlineOF.Text = (clearLabels) ? "Deadline original format: " : "Deadline original format: " + yourTaskBuf.deadlineoriginalformat;
            labelTaskExtLink.Text = (clearLabels) ? "Extlink: " : "Extlink: " + yourTaskBuf.extlink;
            labelTaskTags.Text = (clearLabels) ? "Tags: " : "Tags: " + yourTaskBuf.tags;
            labelTaskColumnId.Text = (clearLabels) ? "Column id: " : "Column id: " + yourTaskBuf.columnid;
            labelTaskLaneId.Text = (clearLabels) ? "Lane id: " : "Lane id: " + yourTaskBuf.laneid;
            labelTaskLeadTime.Text = (clearLabels) ? "Lead time: " : "Lead time: " + yourTaskBuf.leadtime;
            labelTaskBlocked.Text = (clearLabels) ? "Blocked: " : "Blocked: " + yourTaskBuf.blocked;
            labelTaskBlockedReason.Text = (clearLabels) ? "Blocked reason: " : "Blocked reason: " + yourTaskBuf.blockedreason;
            labelTaskColumnName.Text = (clearLabels) ? "Column name: " : "Column name: " + yourTaskBuf.columnname;
            labelTaskLaneName.Text = (clearLabels) ? "Lane name: " : "Lane name: " + yourTaskBuf.lanename;
            labelTaskColumnPath.Text = (clearLabels) ? "Column path: " : "Column path: " + yourTaskBuf.columnpath;
            labelTaskLoggedTime.Text = (clearLabels) ? "Logged time: " : "Logged time: " + yourTaskBuf.loggedtime;
            textBox1.Text = (clearLabels) ? "" : yourTaskBuf.description;

            if (clearLabels)
            {
                listBoxSubTasks.Items.Clear();
                listBox1.Items.Clear();
            }
            else
            {
                listBoxSubTasks.Items.Clear();
                foreach (YourSubTask yourSubTaskBuf in yourTaskBuf.subtaskList)
                {

                    listBoxSubTasks.Items.Add(yourSubTaskBuf.title);
                }

                listBox1.Items.Clear();
                foreach (YourTaskEvent yourTaskEventBuf in yourTaskBuf.historyDetails)
                {
                    listBox1.Items.Add(yourTaskEventBuf.historyevent + " " + yourTaskEventBuf.details);
                }
            }

            if (yourTaskBuf.historyDetails.Count > 0)
            {
                button2.Enabled = false;
            }
            else
            {
                button2.Enabled = true;
            }
        }
        private void populateYourSubTask(YourSubTask yourSubTaskBuf, bool clearLabels)//populate subtasks
        {
            labelSubtaskTitle.Text = (clearLabels) ? "Subtask title: " : "Subtask title: " + yourSubTaskBuf.title;
            labelSubtaskId.Text = (clearLabels) ? "Subtask id: " : "Subtask id: " + yourSubTaskBuf.subtaskid;
            labelSubTaskAssignee.Text = (clearLabels) ? "Assignee: " : "Assignee: " + yourSubTaskBuf.assignee;
            labelSubTaskCompletionDate.Text = (clearLabels) ? "Completion date: " : "Completion date: " + yourSubTaskBuf.completiondate;
        }
        private void populateYourTaskEvents(YourTaskEvent yourTaskEventBuf, bool clearlabels)//popullate taskevents
        {
            labelHistoryAuthor.Text = (clearlabels) ? "Author: " : "Author: " + yourTaskEventBuf.author;
            labelHistoryDetails.Text = (clearlabels) ? "Details: " : "Details: " + yourTaskEventBuf.details;
            labelHistoryEntryDate.Text = (clearlabels) ? "Entry date: " : "Entry date: " + yourTaskEventBuf.entrydate;
            labelHistoryEventType.Text = (clearlabels) ? "Event type: " : "Event type: " + yourTaskEventBuf.eventtype;
            labelHistoryId.Text = (clearlabels) ? "History id: " : "Hystory id: " + yourTaskEventBuf.historyid;
            labelHistoryEvent.Text = (clearlabels)? "History event: " : "history even:" +  yourTaskEventBuf.historyevent;
        }
        private void enableSearching()//enable the searchcontrols and assign the autocomplete list
        {
            textBox2.Enabled = true;
            button1.Enabled = true;
            checkBox1.Enabled = true;
            dateTimePicker1.Enabled = true;
            dateTimePicker2.Enabled = true;

            foreach (YourProject yourProjectBuf in allProjects.yourProjectList)         //for every project
            {
                foreach (YourBoard yourBoardBuf in yourProjectBuf.boardList)            //for every board
                {
                    foreach (YourTask yourTaskBuf in yourBoardBuf.taskList)             //for every task
                    {
                        if (!source.Contains(yourTaskBuf.assignee))                     //if the assignee is not already in the autocompletelist
                            source.Add(yourTaskBuf.assignee);                           //add assignee to the autocompletelist
                        if (yourTaskBuf.assignee.Contains(" "))                         //if the assignees name contains multiple blocks split them up and add to the autocompletelist
                        {
                            string test = "";
                            for (int i = 0; i < yourTaskBuf.assignee.Length; i++)       
                            {
                                if (yourTaskBuf.assignee.Substring(i, 1) == " ")
                                {
                                    if (!source.Contains(test))
                                    {
                                        source.Add(test);
                                    }
                                    
                                    test = "";
                                }
                                else
                                {
                                    test += yourTaskBuf.assignee.Substring(i, 1);
                                }
                            }
                            if (!source.Contains(test))
                            {
                                source.Add(test);
                            }

                        }
                    }

                }
            }
            textBox2.AutoCompleteCustomSource = source;                                 //add the autocompletelist to the textbox
            textBox2.AutoCompleteMode = AutoCompleteMode.SuggestAppend;

        }
        #endregion

        #endregion
        
        #region events
        #region treeviews
        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            TreeNode selectedNode = treeView1.SelectedNode;                                                     //get the selectednode
            switch (selectedNode.Level)                                                                         //what is the level of the selected level
            {
                case 0:                                                                                         //level 0 (project selected)
                    foreach (YourProject yourProjectBuf in allProjects.yourProjectList)                         //find the project
                    {
                        if (selectedNode.Name.Substring(8, selectedNode.Name.Length - 8) == yourProjectBuf.id)
                        {
                            populateYourProject(yourProjectBuf, false);                                         //populate the project labels

                            YourBoard yourBoardDummy = new YourBoard();
                            YourTask yourTaskDummy = new YourTask();
                            YourSubTask yourSubTaskDummy = new YourSubTask();
                            populateYourBoard(yourBoardDummy, true);                                            //clear the board labels
                            populateYourTask(yourTaskDummy, true);                                              //clear the task labels
                            populateYourSubTask(yourSubTaskDummy, true);                                        //clear the subtask labels

                            currentProjectId = yourProjectBuf.id;                                               //save the project id
                            currentBoardId = "";
                            currentTaksId = "";

                            button2.Enabled = false;
                            break;
                        }
                    }
                    break;
                case 1:                                                                                         //level 1 (board selected)
                    foreach (YourProject yourProjectBuf in allProjects.yourProjectList)                         //find the board
                    {
                        foreach (YourBoard yourBoardBuf in yourProjectBuf.boardList)
                        {
                            if (selectedNode.Name.Substring(6, selectedNode.Name.Length - 6) == yourBoardBuf.id)
                            {
                                populateYourProject(yourProjectBuf, false);                                     //populate the project labels
                                populateYourBoard(yourBoardBuf, false);                                         //project the board labels

                                YourTask yourTaskDummy = new YourTask();
                                YourSubTask yourSubTaskDummy = new YourSubTask();
                                populateYourTask(yourTaskDummy, true);                                          //clear the task labels
                                populateYourSubTask(yourSubTaskDummy, true);                                    //clear the subtask labels

                                currentProjectId = yourProjectBuf.id;                                           //save the project id        
                                currentBoardId = yourBoardBuf.id;                                               //save the board id
                                currentTaksId = "";

                                button2.Enabled = false;
                                goto BreakLoops1;
                            }
                        }

                    }
                BreakLoops1:
                    break;
                case 2:                                                                                         //level 2 (task selected)
                    foreach (YourProject yourProjectBuf in allProjects.yourProjectList)                         //find the task
                    {
                        foreach (YourBoard yourBoardBuf in yourProjectBuf.boardList)
                        {
                            foreach (YourTask yourTaskBuf in yourBoardBuf.taskList)
                            {
                                if (selectedNode.Name.Substring(5, selectedNode.Name.Length - 5) == yourTaskBuf.taskid)
                                {
                                    populateYourProject(yourProjectBuf, false);                                 //populate project labels
                                    populateYourBoard(yourBoardBuf, false);                                     //populate board labels
                                    button2.Enabled = false;
                                    populateYourTask(yourTaskBuf, false);                                       //populate task labels

                                    YourSubTask yourSubTaskDummy = new YourSubTask();
                                    populateYourSubTask(yourSubTaskDummy, true);                                //clear the subtask labels

                                    currentProjectId = yourProjectBuf.id;                                       //save the project id
                                    currentBoardId = yourBoardBuf.id;                                           //save te board id
                                    currentTaksId = yourTaskBuf.taskid;                                         //save the tak id
                                    goto BreakLoops2;
                                }
                            }
                        }
                    }
                BreakLoops2:
                    break;
                default: break;
            }

        }
        #endregion
        #region textboxes
        private void textBoxPassword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                buttonLogin_Click((object)sender, (EventArgs)e);
            }
        }
        #endregion
        #region listboxes
        private void listBoxSubTasks_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (YourProject yourProjectBuf in allProjects.yourProjectList)                     //find subtask
            {
                foreach (YourBoard yourBoardBuf in yourProjectBuf.boardList)
                {
                    foreach (YourTask yourTaskBuf in yourBoardBuf.taskList)
                    {
                        foreach (YourSubTask yourSubTaskBuf in yourTaskBuf.subtaskList)
                        {
                            if (listBoxSubTasks.SelectedItem.ToString() == yourSubTaskBuf.title)    
                            {
                                populateYourSubTask(yourSubTaskBuf, false);                         //populate subtask labels
                                goto BreakLoops3;
                            }
                        }
                    }
                }
            }
        BreakLoops3: ;
        }
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        #endregion
        #region buttons
        private void button1_Click(object sender, EventArgs e)//search button is pressed
        {
            #region preparing
            asked_name = textBox2.Text;                                                             //search term
            DateTime test1 = dateTimePicker1.Value;
            DateTime test2 = dateTimePicker2.Value;
            toolStripStatusLabel1.Text = "asking for historydetails";
            #endregion

            if ((asked_name == "") && (checkBox1.Checked == false))                             //nothing is filled in and the checkbox is not checked
            {
                MessageBox.Show("No search entered", "error");
            }
            else
            {
                if (test1 > test2)                                                              //this can not be... start before stop
                {
                    MessageBox.Show("your start date is later then your stop date", "error");
                }
                else
                {
                    if (backgroundWorker2.IsBusy == false)
                    {
                        backgroundWorker2.RunWorkerAsync();                                     //download the taskevents that are needed and search wich tasks are in progress for a given periode and a given name
                                                                                                //this download and searccode is executed in a backgroundworker
                                                                                                // because the form is not responsive when downloading when we dont use another thread. like the backgroundworker
                    }

                }

            }



        }
        private void buttonLogin_Click(object sender, EventArgs e)// login button is pressed
        {
            toolStripStatusLabel1.Text = "Asking for Api key and login info";
            email = textBoxEmail.Text;
            password = textBoxPassword.Text;
            kanbanizeApiObj = new KanbanizeConnect(email, password);    //create an object of the kanbanize api handler class
            if (backgroundWorker1.IsBusy == false)
                backgroundWorker1.RunWorkerAsync();                     //download the apikey, login, and all of the tasks (no task events are loaded)
                                                                        //this downloadis executed in a backgroundworker
                                                                        // because the form is not responsive when downloading when we dont use another thread. like the backgroundworker
        }
        private void button2_Click(object sender, EventArgs e)//download task history
        {
            foreach (YourProject yourProjectBuf in allProjects.yourProjectList)                     //find subtask
            {
                foreach (YourBoard yourBoardBuf in yourProjectBuf.boardList)
                {
                    foreach (YourTask yourTaskBuf in yourBoardBuf.taskList)
                    {
                        if (yourTaskBuf.taskid == currentTaksId)
                        {
                            historyDetails = kanbanizeApiObj.getTaskDetails(yourBoardBuf.id, yourTaskBuf.taskid);   //download the task events 
                            foreach (TaskEvent taskEventBuf in historyDetails.eventList)
                            {
                                YourTaskEvent yourTaskEventBuf = new YourTaskEvent();
                                yourTaskEventBuf.author = taskEventBuf.author;
                                yourTaskEventBuf.details = taskEventBuf.details;
                                yourTaskEventBuf.entrydate = taskEventBuf.entrydate;
                                yourTaskEventBuf.eventtype = taskEventBuf.eventtype;
                                yourTaskEventBuf.historyevent = taskEventBuf.historyevent;
                                yourTaskEventBuf.historyid = taskEventBuf.historyid;
                                yourTaskBuf.historyDetails.Add(yourTaskEventBuf);
                            }
                            populateYourProject(yourProjectBuf, false);
                            populateYourBoard(yourBoardBuf, false);
                            populateYourTask(yourTaskBuf, false);
                            button2.Enabled = false;
                            goto breakloop4;
                        }
                    }
                }
            }
            breakloop4:;
        }
        #endregion
        #region form
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //when the form is closing excel needs to be closed down
            #region close Excel                                     
            xlWorkBook.Close(true, misValue, misValue);                     
            xlApp.Quit();
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
            MessageBox.Show("it could be that an excel process is still running, check your task manager", "warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            #endregion
        }
        #endregion
        #region toolstrip
        private void toolStripStatusLabel1_TextChanged(object sender, EventArgs e)
        {
            //when the label gets a new text it must be visible on the gui immediatly
            this.Refresh();
        }
        #endregion
        #region checckboxes
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            //if the checkbox "verybody" is checked no search term can be entered
            if (checkBox1.Checked)
                textBox2.Enabled = false;
            else
                textBox2.Enabled = true;
        }
        #endregion
        #endregion

        #region BackGroundWorker1: Get login and all tasks
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)                             //the work that needs to be done 
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            yourLogin = kanbanizeApiObj.getLogin();                                                         //download the api key and other login info
            if (yourLogin != null)
            {
                kanbanizeApiObj.apikey = yourLogin.apikey;                                                  //sync apikey between api handler and this class
                getAllInfo();                                                                               //download all of the projects,boards and tasks
                e.Result = "OK";                                                                            //the result is good

            }
            else
            {
                e.Result = "NOK";                                                                           //the result is not good probably an invalid e-mail adres or password
            }

        }
        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)     //the backgroundworker is done
        {
            if (e.Result != "OK")                                                                           //if the result si different from ok
            {
                MessageBox.Show("Invalid E-mail or password", "Error");                                     //show the error message to the user
                toolStripStatusLabel1.Text = "Invalid E-mail or password";
            }
            else
            {
                populateLogin();                                                                            //populate the login labels
                populateTreeView();                                                                         //populate the treeview
                populateYourProject(allProjects.yourProjectList[0], false);                                 //select the first project and populate the project labels
                enableSearching();
                toolStripStatusLabel1.Text = "Succesfully logged in and received all tasks";
            }
        }
        #endregion
        
        #region BackGroundWorker2: Get all history details and search
        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;

            DateTime start = dateTimePicker1.Value;                                                                     //get the start and stop date from the gui
            DateTime stop= dateTimePicker2.Value;

            double step = (double)100 / (double)numberTasks;                                                            //how much percent is done when 1 task is done
            double progress = 0;                                                                                        //the progress in % (needs to be send to the toolstripprogresbar

            foreach (YourProject yourProjectBuf in allProjects.yourProjectList)                                         //for every project
            {
                foreach (YourBoard yourBoardBuf in yourProjectBuf.boardList)                                            //for every board
                {
                    foreach (YourTask yourTaskBuf in yourBoardBuf.taskList)                                             // for every task
                    {
                        if (((yourTaskBuf.assignee.Contains(asked_name)) && (yourTaskBuf.columnname != "Requested")) || (checkBox1.Checked == true)) //if the task.assignee containts the searchterm and the task is not in the requested column, or everybody is selected  then download
                        {
                            if (yourTaskBuf.historyDetails.Count == 0)                                                  //if the task events havent been downloaded yet.
                            {
                                historyDetails = kanbanizeApiObj.getTaskDetails(yourBoardBuf.id, yourTaskBuf.taskid);   //download the task events 
                                foreach (TaskEvent taskEventBuf in historyDetails.eventList)
                                {
                                    YourTaskEvent yourTaskEventBuf = new YourTaskEvent();
                                    yourTaskEventBuf.author = taskEventBuf.author;
                                    yourTaskEventBuf.details = taskEventBuf.details;
                                    yourTaskEventBuf.entrydate = taskEventBuf.entrydate;
                                    yourTaskEventBuf.eventtype = taskEventBuf.eventtype;
                                    yourTaskEventBuf.historyevent = taskEventBuf.historyevent;
                                    yourTaskEventBuf.historyid = taskEventBuf.historyid;
                                    yourTaskBuf.historyDetails.Add(yourTaskEventBuf);
                                    

                                }
                                
                            }
                            foundTasks.Add(yourTaskBuf);                                                                //add to found tasks
                            progress+=step;                                                                             //add step to progres
                            if (Convert.ToInt32(progress) == 100)        
                            {
                                worker.ReportProgress(99);
                            }
                            else
                            {
                                worker.ReportProgress(Convert.ToInt32(progress));                                       //report the progress from this backgroundworker
                            }
                        }

                    }
                }
            }
            worker.ReportProgress(100);

            FindTasksForDate(start, stop);
        }

        private void backgroundWorker2_ProgressChanged(object sender, ProgressChangedEventArgs e)//the progress is updated
        {
            if (e.ProgressPercentage == 100)                                                                        //everything is downloaded
            {
                toolStripStatusLabel1.Text = "searching and writing to excel";
                toolStripProgressBar1.Value = 100;

            }
            else
            {
                toolStripStatusLabel1.Text = "asking for historydetails ("+e.ProgressPercentage.ToString()+" %)";   //show the current percentage on the progress bar and the tooltipstatuslabel
                toolStripProgressBar1.Value = e.ProgressPercentage;
            }

        }

        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)//everything is downloaded and the excel file is made
        {
            SaveToExcel();                                          //save the excel file
            toolStripStatusLabel1.Text = "written to excell";
        }

        #endregion

        private void listBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            ListBox test = (ListBox)sender;
            foreach (YourProject yourProjectBuf in allProjects.yourProjectList)                     //find subtask
            {
                foreach (YourBoard yourBoardBuf in yourProjectBuf.boardList)
                {
                    foreach (YourTask yourTaskBuf in yourBoardBuf.taskList)
                    {
                        foreach (YourTaskEvent yourTaskEventBuf in yourTaskBuf.historyDetails)
                        {
                            
                            if (test.SelectedItem.ToString() == (yourTaskEventBuf.historyevent + " " + yourTaskEventBuf.details))
                            {
                                populateYourTaskEvents(yourTaskEventBuf, false);
                                goto BreakLoops5;
                            }
                        }
                    }
                }
            }
        BreakLoops5: ;
        }




    }

    /*yourClasses is the same as the classes in KanbanizeConnect.cs
     * Difference here:
     * everything is in 1 object.
     * this could not be done in KanbanizeConnect because this 1 object exists from multiple xml files received from the api*/
     #region yourClasses
    public class YourProjects
    {
        public List<YourProject> yourProjectList = new List<YourProject>();
    }

    public class YourProject
    {
        public string name { get; set; }
        public string id { get; set; }
        public List<YourBoard> boardList = new List<YourBoard>();
    }

    public class YourBoard
    {
        public string name {get; set;}
        public string id { get; set;}
        public List<YourTask> taskList = new List<YourTask>();
    }

    public class YourTask
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
        public string columnname { get; set; }
        public string lanename { get; set; }
        public string columnpath { get; set; }
        public string loggedtime { get; set; }
        public List<YourSubTask> subtaskList = new List<YourSubTask>();
        public List<YourTaskEvent> historyDetails = new List<YourTaskEvent>();
    }

    public class YourSubTask
    {
        public string subtaskid { get; set; }
        public string assignee { get; set; }
        public string title { get; set; }
        public string completiondate { get; set; }
    }

    public class YourTaskEvent
    {
        public string eventtype { get; set; }
        public string historyevent { get; set; }
        public string details { get; set; }
        public string author { get; set; }
        public string entrydate { get; set; }
        public string historyid { get; set; }
    }
    #endregion
}
