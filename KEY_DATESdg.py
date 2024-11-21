<?xml version="1.0" encoding="UTF-8"?>
<tfb>
  <KeyDates>
    <Init>
      <![CDATA[
import clr
import datetime
import System.Windows.Input as Input

clr.AddReference('mscorlib')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('System.Windows.Forms')
clr.AddReference('IronPython.Wpf')

#from datetime import datetime
from datetime import date, timedelta, datetime    # Errors because we use 'now'
from System import DateTime
from System.Diagnostics import Process
from System.Globalization import DateTimeStyles
from System.Collections.Generic import Dictionary
from System.Windows import Controls, Forms, LogicalTreeHelper
from System.Windows import Data, UIElement, Visibility, Window, FontWeights, GridLength, GridUnitType
from System.Windows.Controls import Button, Canvas, GridView, GridViewColumn, ListView, Orientation, DataGrid, SelectedCellsChangedEventArgs
from System.Windows.Data import Binding, CollectionView, ListCollectionView, PropertyGroupDescription
from System.Windows.Forms import SelectionMode, ListViewItem, MessageBox, MessageBoxButtons, DialogResult
from System.Windows.Input import KeyEventHandler  #, ModifierKeys, MouseButtonState, KeyEventArgs
from System.Windows.Media import Brush, Brushes

# Following a demo to the 'Risk board' (16th Jan 2024), there are a number of changes needed here:
# 1) whether to use 'Dates' or 'Tasks' - no concensus, and have decided to set by department
#    - therefore will need a table where we can set the department preference (and suggest adding to the 'defaults' version of the form for 'Risk'/'IT' to update)
# 2) Terminology-wise, we need to make sure we only call them 'Key Dates' (dashboards etc need to be updated)
# 3) We ONLY want to see 'Key Dates' here (all other tasks/dates shouldn't be visible) - now handling via 'Step Category' of 'KeyDates'
# 4) The 'Description' of a 'Key Date' should include the prefix 'Key Date: ' to aid in distinguishing these from regular reminder (tasks) 
#     - still not fully convinced this is applicable
# 5) Need a button to mark as 'complete' - have added 
# 6) possibly, remove editable datagrid (for 'Date Missed Notes'), and add 'DMN' to bottom section for consistency
#     We have gone the opposite way here in that we are fully editing in datagrid itself!
# 7) Perhaps introduce a 'splitter' control (horizontal) to allow 'Attendee(s)' list to be expanded (obvs only for 'Dates' element)
#     I have added an 'expander' control that fits the bill nicely here

# As of 29th August 2024
# - Created table: Usr_KeyDates_Defaults to store each departments preferences (Dates or Tasks)
# - Updated code so that both Tasks and Dates now work, and will only show the tab according to department preference
# - Updated 'Defaults' code so that it will add to the respective departments preference 
# - Tidied up XAML (put controls into StackPanel to group 'rows' better)
# - Added 'Grouping' to DataGrids, so these now display banner for 'Outstanding' and 'Completed' (and 'Added from default')
# - Added new 'Step Category' in P4W to use for these 'KeyDates' ('KeyD' and 'KeyDD' - latter is for the 'defaults' added... once updated properly, we remove second 'D')
# - WORKING ON: Adding a separate button to update 'Date Missed Notes' and only showing this field and button once the due date has passed
#    No, think the solution here is to split out 'add/save' functions, based on current 'Status' (Group)
# As of 6th September 2024
# - Made DataGrids directly editable - but perhaps could do with some error checking
# - Added 'Add New' button onto 'Tasks' and 'Dates' for manually adding new row (this does the actual add, and then we have the 'CellEditFinished' to update
# - moved attendees into a 'side panel' (an expander control)
# - Investigating 'UpdateThis' column... I had assumed that we set this to '1' to make it sync to users calendar, however, it doesn't appear to get set back to zero
#   and keeps nagging user every 10 minutes (or whatever their sync interval is).
#    Going to try making into a global variable, to make it easier to switch all occurrences on/off
#    Following a support call I had with Advanced on this, 'UpdateThis' is meant to trigger update (when set to 1), but only for Tikit Exchange Connector
#    Note: 'UpdateThis' is exclusively a Tikit Exchange Connector thing... doesn't do anything for 'Third Party Diary' dates

####################################################################################################################################################################
# Global Functions (Line num and name)       |  Task functions                                |  DEFAULTS                                          |  Date functions
#  88 - myOnLoadEvent(s, event)              | 217 - class KeyTasks(object)                   | 1251 - class KDDefaults(object)                    | 1522 - class KeyDates(object)
# 142 - runSQL(codeToRun, showError = False, | 291 - refresh_KeyTasks(s, event)               | 1275 - refresh_KDDefaults_List(s, event)           | 1595 - refresh_KeyDates(s, event)
#     errorMsgText = "", errorMsgTitle = "") | 358 - task_AddNew(s, event)                    | 1364 - addDefaults_ToDiaryDates(s, event)          | 1668 - dg_KeyDates_SelectionChanged(s, event)
# 178 - runSQL_OLD - consider removing       | x388 - task_Validation - to remove             | 1422 - addDefaults_ToTasks(s, event)               | 1745 - class comboTypes(object)
# 1170 - get_FullEntityRef(shortRef)         | x430 - task_AddNewOrSave - to remove           | 1479 - defaultDiaryDates_SelectAllNone(s, event)   | 1760 - populateComboTypes(s, event)
# 1186 - RemC(myInputStr)                    | 562 - global_AddTask(...)                      |                                                    | 1765 - get_TypeOfUnitTypes()
# 1201 - getSQLDate(varDate)                 | x609 - task_SetToFeeEarner(s, event)           |                                                    | 1774 - expand_DateAttendees(s, event)
#                                            | x629 - task_SetToCurrentUser(s, event)         |                                                    | 1779 - contract_DateAttendees(s, event)
#                                            | x642 - task_optAddNew_Clicked                  |                                                    | 1786 - class attendees(object)
#                                            | x655 - task_optEditSelected_Clicked            |                                                    | 1804 - refresh_AttendeeList(usersToPreSelect = '')
#                                            | 672 - get_taskStatusTypes()                    |                                                    | 1899 - populateAttendeesList(s, event)
#                                            | 682 - populate_taskStatus(s, event)            |                                                    | x1907 - diaryDates_ValidationErrCount()
#                                            | 688 - get_taskPriorityTypes()                  |                                                    | x1930 - diaryDates_ValidationMessage()
#                                            | 696 - populate_taskPriority(s, event)          |                                                    | 1954 - addOrUpdate_KeyDate(s, event)
#                                            | 702 - class postponeOptions(object)            |                                                    | 2117 - global_AddDate(...)
#                                            | 717 - populate_taskPostponeOptions             |                                                    | 2162 - date_AddNew(s, event)
#                                            | 734 - task_PostponeNow(s, event)               |                                                    | 2189 - get_DD_CaseStepID(matchingDescription)
#                                            | 770 - task_cellSelection_Changed(s, event)     |                                                    | 2194 - get_DD_CaseStepID2(forID)
#                                            | 875 - deleteTask(s, event)                     |                                                    | 2199 - get_DD_ID(matchingDescription)
#                                            | 912 - revertTask(s, event)  [NEW]              |                                                    | 2206 - getDefaultNextDay()
#                                            | 1001 - dg_KeyTasks_cellEdit_Finished(s, event) |                                                    | 2237 - getTimeFixed(inputToCheck)
#                                            | 
####################################################################################################################################################################
gUpdateThis = 1

################################################################################################################
# # # #   O N - L O A D   F U N C T I O N   # # # #
def myOnLoadEvent(s, event):
  # Here we setup the view according to Department and the 'Type' set to use (whether 'Dates' or 'Tasks')
  # lookup default Type for current Matters' Department (note: we did have this on the XAML but it appears that this code runs before XAML loaded - lbl_DeptDefaultType.Content)
  mType = runSQL("SELECT TypeToUse FROM Usr_KeyDatesDeptSettings WHERE Department = (SELECT CaseTypeGroupRef FROM CaseTypes WHERE Code = (SELECT CaseTypeRef FROM Matters WHERE EntityRef = '{0}' AND Number = {1}))".format(_tikitEntity, _tikitMatter),apostropheHandle=1)
  #lbl_DeptDefaultType.Content = mType
  # this is actually a text box, so needs to be '.Text', not '.Content'
  lbl_DeptDefaultType.Text = mType

  populate_FeeEarnersList(s, event)
  
  if _tikitUser in ('MP', 'LD1'):
    chk_DebugMode.Visibility = Visibility.Visible
  else:
    chk_DebugMode.Visibility = Visibility.Collapsed
    chk_DebugMode.IsChecked = False
  
  if mType == 'Tasks':
    refresh_KeyTasks(s, event)
    populate_taskStatus(s, event)
    populate_taskPriority(s, event)
    populate_taskPostponeOptions(s, event)
    #task_optAddNew_Clicked(s, event)
    ti_Tasks.IsSelected = True

    # hide 'Diary Dates' tab and 'Add to Diary Dates' button (from 'Defaults' tab)
    ti_Dates.Visibility = Visibility.Collapsed
    btn_AddDefaultsToDates.Visibility = Visibility.Collapsed

  elif mType == 'Dates':
    populateComboTypes(s, event)
    refresh_KeyDates(s, event)
    refresh_AttendeeList('')
    contract_DateAttendees(s, event)
    ti_Dates.IsSelected = True

    # hide 'Task Reminders' tab and 'Add to Tasks' button (from 'Defaults' tab)
    ti_Tasks.Visibility = Visibility.Collapsed
    btn_AddDefaultsToTasks.Visibility = Visibility.Collapsed

  #if countOfDGitems() > 0:    # NB: not sure why I don't appear to have known about DataGrid.Items.Count (I created a function to iterate over and count, lol)
  #if dg_KeyTasks.Items.Count > 0:
  #  updateStepCompleted(s, event)
  #  updateMPLinkedFields(s, event)
  refresh_KDDefaults_List(s, event)

  for d in dg_DateDefaults.Items:
    if d.iGroup == 'Available':   # and newTickStatus == True:
      btn_AddAllDefaultTasks.Visibility = Visibility.Visible
      break
      
  return


###################################################################################################################################################
# New April 2024 - 'runSQL()' should replace manual '_tikitResolver.Resolve()' in functions
def runSQL(codeToRun, showError = False, errorMsgText = "", errorMsgTitle = "", apostropheHandle = 0):
  # I'm wondering if there's merit to having a dedicated function for running/executing SQL, as we tend to use same 'try except' wrapper
  # (or ought to be for trapping errors) and therefore could save some lines of code and make code easier to read because we're not having to repeat stuff
  # codeToRun     = Full SQL of code to run. No need to wrap in '[SQL: code_Here]' as we can do that here
  # showError     = True / False. Indicates whether or not to display message upon error
  # errorMsgText  = Text to display in the body of the message box upon error (note: actual SQL will automatically be included, so no need to re-supply that)
  # errorMsgTitle = Text to display in the title bar of the message box upon error
  
  # if no code actually supplied, exit early...
  if len(codeToRun) < 10:
    MessageBox.Show("The supplied 'codeToRun' doesn't appear long enough, please check and update this code if necessary.\nPassed SQL: " + str(codeToRun), "ERROR: runSQL...")
    return
  
  # Add '[SQL: ]' wrapper if not already included...
  if codeToRun[:5] == "[SQL:":
    fCodeToRun = codeToRun
  else:
    fCodeToRun = "[SQL: " + codeToRun + "]"
  
  # try to execute the SQL...
  try:
    tmpValue = _tikitResolver.Resolve(fCodeToRun)
    # Adding apostrophe handler that was added in the MRA to try and fix bug MLC-74 on Jira
    if apostropheHandle == 1:
      tmpValue = tmpValue.replace("'", "''")
    returnVal = str(tmpValue)
    returnVal1 = 'N/A' if returnVal == None else returnVal
  except:
    # there was an error... check to see if opted to show message or not...
    if showError == True:
      MessageBox.Show("{0}\n\nSQL used:\n{1}".format(errorMsgText, codeToRun), errorMsgTitle)
    returnVal = ''
    returnVal1 = "!Error"
    
  # print SQL to run in console
  debugMessage(msgBody = "runSQL(...):\n  CodeToRun: {0}\n  ShowError: {1}\n  ErrorMsgText: '{2}'\n  ErrorMsgTitle: '{3}'\n  > Result: {4}".format(fCodeToRun, showError, errorMsgText, errorMsgTitle, returnVal1))
  return returnVal

###################################################################################################################################################
# This was suggested by ChatGPT when asking how to add validation on datagrid.  It seemed like the right idea, but as we're not actually using
# the MVVM pattern, this isn't working as I had hoped... I did even try amending the __getItem__ for main 'KeyTasks', but we just got error about
# 'IDataErrorInfo' not recognised
# # class YourViewModel(IDataErrorInfo):
#     def __init__(self):
#       self._iDateRemind = None

#     @property
#     def iDateRemind(self):
#       return self._iDateRemind

#     @iDateRemind.setter
#     def iDateRemind(self, value):
#       self._iDateRemind = value

#     def __getitem__(self, column_name):
#       if column_name == 'iDateRemind':
#         if self._iDateRemind is None:
#           return "Reminder Date cannot be empty."
#         if self._iDateRemind < datetime.datetime.now():
#           return "Reminder Date cannot be in the past."
#       return None

#     @property
#     def Error(self):
#       return None
####################################################################################################################################################

class KeyTasks(object):
  def __init__(self, myDesc, myDate, myDateRemind, myAssignedTo, myStatus, myPriority, myPercentComp, myITCode, 
                myAgenda, myCaseStepID, myGroup, myDateMissedN, myKDID, myFEList, myStatusID, myPriorityID, myInclReminder):
    #tmpDateRemind = str(myDateRemind)
    #tmpDateRemind = tmpDateRemind[:10]
    if len(str(myDateRemind)) > 10:
      tmpTime = str(myDateRemind)
      tmpTime = tmpTime[11:]
      tmpTime = tmpTime[:5]
      if tmpTime[4] == ':':
        tmpTime = tmpTime[:4] 
      
      #tmpTime = tmpTime.replace(" AM", "")
      #tmpTime = tmpTime.replace(" PM", "")
      #tmpDate = str(myDate)
      #tmpDate = tmpDate[:10]
    else:
      tmpTime = ''

    self.iDesc = myDesc
    self.iDate = myDate                #tmpDate
    self.inclReminder = True if myInclReminder == 'Y' else False
    self.iDateRemind = myDateRemind    #tmpDateRemind

    self.iDateRemindTime = tmpTime
    if len(tmpTime) > 3:
      self.iHour = tmpTime[:2]
      self.iMins = tmpTime[3:]
    else:
      self.iHour = "09"
      self.iMins = "00"

    self.iAssignedTo = myAssignedTo
    self.iStatus = myStatus
    self.iPriority = myPriority
    self.iPercentComplete = myPercentComp
    
    self.iTCode = myITCode
    self.iTAgenda = myAgenda
    self.iCaseStepID = myCaseStepID
    self.xGrouping = myGroup
    self.xDateMissedNote = myDateMissedN
    #self.iLinkedMPField = myLinkedMPField
    self.iKDid = myKDID
    self.statusID = myStatusID
    self.priorityID = myPriorityID

    # assign list items (combo boxes)
    self.FEList = myFEList
    self.StatusItems = get_taskStatusTypes()
    self.PriorityItems = get_taskPriorityTypes()
    self.iHoursList = get_TimeHours(startHour=7, endHour=19)
    self.iMinsList = get_TimeMins(increment=10)

    return

  def __getitem__(self, index):
    if index == 'Desc': 
      return self.iDesc
    elif index == 'Date': 
      return self.iDate
    elif index == 'ReminderDate': 
      #if self.iDateRemind < datetime.now():
      #  return "Reminder Date cannot be in the past!"
      #elif self.iDateRemind > self.iDate:
      #  return "Reminder Date cannot be greater than the Date of the Task!"
      #else:
      return self.iDateRemind

    elif index == 'ReminderTime':
      # need to only return value from proper fields if something is set
      if self.iHour is None or self.iHour == '00':
        tmpHour = '09'
      else:
        tmpHour = self.iHour
      
      if self.iMins is None or self.iMins == '00':
        tmpMins = '00'
      else:
        tmpMins = self.iMins

      return "{0}:{1}".format(tmpHour, tmpMins)
      # formally the below when allowing for text input
      #return self.iDateRemindTime

    elif index == 'oldRemindTime':
      return self.iDateRemindTime

    #elif index == 'ReminderDateAndTime':
    #  if self.iDateRemind is None or self.iDateRemind == '':
    #    # no reminder date set
    #    return None
    #  else:
    # thought here was to return combined date and time, but may as well do in calling procedure

    elif index == 'AssignedTo':
      return self.iAssignedTo
    elif index == 'Status':
      return self.statusID
    elif index == 'Priority':
      return self.priorityID
    elif index == 'PercentComplete':
      return self.iPercentComplete
    elif index == 'Code':
      return self.iTCode
    elif index == 'Agenda': 
      return self.iTAgenda
    elif index == 'CaseStepID': 
      return self.iCaseStepID
    elif index == 'Group':
      return self.xGrouping
    elif index == 'DateMissedNote':
      return self.xDateMissedNote
    elif index == 'KDid':
      return self.iKDid
    elif index == 'InclReminder':
      if self.inclReminder == True:
        return 'Y'
      else:
        return 'N'


def refresh_KeyTasks(s, event):
  # This function will populate the 'Task Reminders' datagrid

  # we need to get the 'Case History' agenda ID for SQL (unless other departments state they want all, in which case we need to amend this (or add options to 'Dept defaults')
  caseHistoryAgID = get_CaseHistoryAgendaID(titleForError = 'Error: refresh_KeyTasks...')
  # need to pass in list of Fee Earners (cannot run within '__init__' because we end up closing the already open data reader)
  tmpFEs = get_FeeEarnerList()

  # set SQL to get data
  mySQL = """SELECT '0-Desc' = CI.Description, '1-Due Date' = ISNULL(CMS.DiaryDate, KD.Date), 
              '2-Reminder Date' = ISNULL(DT.ReminderDate, KD.ReminderDate), 
              '3-DateMissedNotes' = ISNULL(KD.DateMissedNotes, ''),
              '4-Assigned To' = ISNULL(DT.Username, ISNULL(CMS.AssignedUser, KD.AssignedTo)), 
              '5-Status' = CASE ISNULL(DT.Status, KD.Status) WHEN 0 THEN 'Not Started' WHEN 1 THEN 'In Progress' WHEN 2 THEN 'Completed' WHEN 3 THEN 'Waiting on someone else' WHEN 4 THEN 'Deferred' ELSE 'Completed' END,
              '6-Priority' = CASE ISNULL(DT.Priority, KD.Priority) WHEN 0 THEN 'Low' WHEN 1 THEN 'Normal' WHEN 2 THEN 'High' ELSE '' END, 
              '7-Complete' = ISNULL(DT.Complete, 100), 'H8-Code' = ISNULL(DT.Code,0), 'H9-AgendaRef' = CI.ParentID, 
              'H10-CaseItemRef' = CI.ItemID, 
              'H11-Group' = CASE WHEN CMS.StepCategory = 'KeyDD' THEN '0) Added from Defaults - To set date' ELSE (CASE WHEN CI.CompletionDate IS NULL THEN '1) Outstanding' ELSE '2) Complete' END) END, 
              'H12-KDid' = ISNULL(KD.ID, 0), 'H13-StatusID' = ISNULL(DT.Status, KD.Status), 'H14-PriorityID' = ISNULL(DT.Priority, KD.Priority), 
              '15-InclReminder' = CASE WHEN ISNULL(DT.ReminderDate, '') = '' THEN 'N' ELSE 'Y' END
          FROM Cm_CaseItems CI 
              LEFT OUTER JOIN Cm_Steps CMS ON CI.ItemID = CMS.ItemID 
              LEFT OUTER JOIN Diary_Tasks DT ON CI.ItemID = DT.CaseItemRef 
              LEFT OUTER JOIN Users U ON DT.Username = U.Code 
              LEFT OUTER JOIN Usr_Key_Dates KD ON CI.ItemID = KD.TaskStepID
          WHERE CI.ParentID = {0} AND CMS.Type = 'FreeStyle' AND CMS.StepCategory LIKE 'KeyD%'
          ORDER BY [H11-Group] ASC, CMS.DiaryDate ASC, CI.Description""".format(caseHistoryAgID)

  myItems = []
  _tikitDbAccess.Open(mySQL)

  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          iDesc = '' if dr.IsDBNull(0) else dr.GetString(0)
          iDate = '' if dr.IsDBNull(1) else dr.GetValue(1)
          iDateRemind = '' if dr.IsDBNull(2) else dr.GetValue(2)
          iDateMissedN = '' if dr.IsDBNull(3) else dr.GetString(3)
          iAssignedTo = '' if dr.IsDBNull(4) else dr.GetString(4)
          iStatus = 0 if dr.IsDBNull(5) else dr.GetValue(5)
          iPriority = 0 if dr.IsDBNull(6) else dr.GetValue(6)
          iComplete = 0 if dr.IsDBNull(7) else dr.GetValue(7)
          iCode = 0 if dr.IsDBNull(8) else dr.GetValue(8)
          iAgenda = 0 if dr.IsDBNull(9) else dr.GetValue(9)
          iCaseID = 0 if dr.IsDBNull(10) else dr.GetValue(10)
          iGrouping = '' if dr.IsDBNull(11) else dr.GetString(11)
          iKDid = 0 if dr.IsDBNull(12) else dr.GetValue(12)
          iStID = 0 if dr.IsDBNull(13) else dr.GetValue(13)
          iPrID = 0 if dr.IsDBNull(14) else dr.GetValue(14)
          iinclRem = 'N' if dr.IsDBNull(15) else dr.GetString(15)

          myItems.append(KeyTasks(myDesc = iDesc, myDate = iDate, myDateRemind = iDateRemind, myAssignedTo = iAssignedTo, myStatus = iStatus, myPriority = iPriority, myKDID = iKDid, 
                              myPercentComp = iComplete, myITCode = iCode, myAgenda = iAgenda, myCaseStepID = iCaseID, myGroup = iGrouping, myDateMissedN = iDateMissedN, 
                              myFEList = tmpFEs, myStatusID = iStID, myPriorityID = iPrID, myInclReminder = iinclRem))

    dr.Close()
  _tikitDbAccess.Close()

  # add grouping
  tmpC = ListCollectionView(myItems)
  tmpC.GroupDescriptions.Add(PropertyGroupDescription("xGrouping"))
  dg_KeyTasks.ItemsSource = tmpC # Error currently occuring here, probably due to how tmpc is appearing
  return

def task_AddAllDefaults(s, event):
  # Function tied to the btn_AddAllDefaultTasks on the XAML, designed to add all of the case type defaults to the key dates manager tab
  # This will add all available (not already added based on the "description" field of the task)
  user_confirmed = MessageBox.Show("Are you sure you want to add all Default Tasks?", "Add All Defaults?", MessageBoxButtons.YesNo)
    
  # Process only if the user confirms
  if user_confirmed == DialogResult.Yes:
    for d in dg_DateDefaults.Items:
      if d.iGroup == 'Available':   # and newTickStatus == True:
        dg_DateDefaults.SelectedItems.Add(d)
        btn_AddAllDefaultTasks.Visibility = Visibility.Collapsed
    addDefaults_ToTasks(s, event, 0)
    # Log or print the exception for debugging
    MessageBox.Show("An error occurred while adding default tasks: ")

  
def task_AddNew(s, event):
  # This is the new function to add a new blank Task to the list (NB: we'll give it the 'Default' Category too, so it appears near top)
  # Linked to XAML button.click for: btn_AddNew_Task

  # set defaults
  caseHistoryAgID = get_CaseHistoryAgendaID(titleForError = 'Error: Task_AddNewOrSave - getting Case History agenda ID...')
  currentStepID = get_currentStepID(agendaID = caseHistoryAgID, titleForError = 'Error: Task_AddNewOrSave - getting Current Step ID...')
  dueDate = getDefaultNextDay()
  assignedTo = runSQL("SELECT FeeEarnerRef FROM Matters WHERE EntityRef = '{0}' AND Number = {1}".format(_tikitEntity, _tikitMatter))
  descToUse = getUniqueDescription(desiredDesc = 'New Task - Edit here', forUser = assignedTo, taskORdate = 'Task')

  # we created a generic function to add tasks, so call that passing in values
  global_AddTask(caseHistoryAgendaID = caseHistoryAgID, 
                       currentStepId = currentStepID, 
                     taskDescription = descToUse, 
                         taskDueDate = dueDate, 
                      taskAssignedTo = assignedTo, 
                     taskReminderQty = 15, 
                      taskRemindDate = dueDate, 
                          taskStatus = 0, 
                        taskPriority = 1, 
                      taskPCComplete = 0, 
                       isFromDefault = True)

  # refresh 'Defaults' list and Tasks
  refresh_KeyTasks(s, event)
  
  # select newly added row
  tmpX = -1
  for xRow in dg_KeyTasks.Items:
    tmpX += 1
    if xRow.iDesc == descToUse: # and xRow.iDate == dueDate:
      dg_KeyTasks.SelectedIndex = tmpX
      break
  return


def global_AddTask(caseHistoryAgendaID, currentStepId, taskDescription, taskDueDate, taskAssignedTo, taskRemindDate, 
                      taskPriority, taskReminderQty = 15, taskStatus = 0, taskPCComplete = 0, isFromDefault = False):
  # This function will add a new TASK item with the passed details

  # need to first separate out time as 'sp_InsertStepMP' only wants a date - not time element as well
  # just get the left 10 characters from due date
  spDueDate = taskDueDate[:10]

  # run the in-built stored procedure to add step
    # @AgendaID INT, @InsertWhere VARCHAR(1), @CurrentStepID INT, @Description VARCHAR(260), @Mandatory VARCHAR(1), @DocID INT,
    # @DocType VARCHAR(20), @StartDate VARCHAR(20), @Duration INT, @DurationType VARCHAR(1), @ReminderUnitQty INT, @ReminderTypeOfUnit VARCHAR(1), @Username VARCHAR(12)
  sqlToRun = """EXEC sp_InsertStepMP {0}, 'L', {1}, '{2}', 'N', 0, 'Task', '{3}', 1, 'D', {4}, 'M', 
                '{5}'""".format(caseHistoryAgendaID, currentStepId, taskDescription, spDueDate, taskReminderQty, taskAssignedTo)

  #msg = MessageBox.Show("SQL to be run:\n" + sqlToRun + "\n\nOK to continue?", "Adding Date...", MessageBoxButtons.YesNo)
  #if msg == DialogResult.No:
  #  return
  runSQL(sqlToRun, True, "There was an error using InsertStep to add Task", "Error: global_AddTask") 

  # unfortunately, the InsertStep stored procedure does not add the 'reminder date' for tasks (into Diary_Tasks). Which means we need to manually add it
  # firstly, get the ID from the Diary_Tasks table
  tmpCountDT = runSQL("SELECT COUNT(Code) FROM Diary_Tasks WHERE EntityRef = '{0}' AND MatterNoRef = {1} AND Username = '{2}' AND Description = '{3}'".format(_tikitEntity, _tikitMatter, taskAssignedTo, taskDescription))
  if int(tmpCountDT) == 0:
    # report error - Diary Task doesn't appear to have been added yet 
    appendToDebugLog(textToAppend = "No Diary Table ID (Diary_Tasks) - so not updating ReminderDate, Status, Priority, and Complete Percent.  Also, we don't update the 'StepCategory'", inclEndLine = False, inclTimeStamp = True)
  else:
    # now update it with the actual reminder date, and set UpdateThis to 1 to make sure it feeds through to FE's Outlook
    dtCode = runSQL("SELECT MAX(Code) FROM Diary_Tasks WHERE EntityRef = '{0}' AND MatterNoRef = {1} AND Username = '{2}' AND Description = '{3}'".format(_tikitEntity, _tikitMatter, taskAssignedTo, taskDescription))
    if int(dtCode) > 0:
      sqlToRun = """UPDATE Diary_Tasks SET ReminderDate = '{0}', Status = {1}, Priority = {2}, Complete = {3}, UpdateThis = {4} 
                    WHERE Code = {5}""".format(taskRemindDate, taskStatus, taskPriority, taskPCComplete, gUpdateThis, dtCode)
      runSQL(sqlToRun)

      if isFromDefault == True:
        stepCatID = 'KeyDD'
      else:
        stepCatID = 'KeyD'
        
      # additionally, it doesn't add our 'KeyDate' category to the Cm_Steps table, so we need to add this too
      # get the Case Item ID from the Dairy_ table
      caseID = runSQL("SELECT CaseItemRef FROM Diary_Tasks WHERE Code = {0}".format(dtCode))
      if int(caseID) > 0:
        runSQL("UPDATE Cm_Steps SET StepCategory = '{0}' WHERE ItemID = {1}".format(stepCatID, caseID))

  return


def task_SetToFeeEarner(s, event):
  # This function will set the 'Assigned To' combo box for the Fee Earner of the matter (on 'Task' tab)
  # Linked to XAML button.click: btn_SetTaskAssignee_MatterFE  (this only appears in the 'Debug Mode' - consider deleting)

  # firstly, lookup matter Fee Earner
  matterFE = runSQL("SELECT FeeEarnerRef FROM Matters WHERE EntityRef = '{0}' AND Number = {1}".format(_tikitEntity, _tikitMatter))

  if len(matterFE) == 0:
    MessageBox.Show("There doesn't appear to be a Fee Earner set against this matter - please check the Matter Properties screen")
    return

  # iterate over items in combo box, and select the Fee Earner (when found)
  tmpCount = -1
  for i in cbo_Task_AssignedTo.Items:
    tmpCount += 1
    if i.iCode == matterFE:
      cbo_Task_AssignedTo.SelectedIndex = tmpCount
      break
  return


def task_SetToCurrentUser(s, event):
  # This function will set the 'Assigned To' combox box to the current user (on 'Task' tab)
  # Linked to XAML button.click: btn_SetTaskAssignee_CurrUser  (this only appears in the 'Debug Mode' - consider deleting)

  # iterate over items in combo box, and select current user details (when found)
  tmpCount = -1
  for i in cbo_Task_AssignedTo.Items:
    tmpCount += 1
    if i.iCode == _tikitUser:
      cbo_Task_AssignedTo.SelectedIndex = tmpCount
      break
  return


def task_optAddNew_Clicked(s, event):
  # This is the 'Tasks' tab 'Add New' option/radio button, and clicking this will nullify current details in fields
  # Linked to XAML radio button.click: opt_Task_AddNew  (this only appears in the 'Debug Mode' - consider deleting)
  
  opt_Task_AddNew.FontWeight = FontWeights.Bold
  opt_Task_EditSelected.FontWeight = FontWeights.Normal
  grp_PostponeOptions.Visibility = Visibility.Hidden
  cbo_Task_Status.SelectedIndex = 0
  cbo_Task_Priority.SelectedIndex = 1
  dg_KeyTasks.SelectedIndex = -1
  # task_cellSelection_Changed(s, event)
  return

def task_optEditSelected_Clicked(s, event):
  # This is the 'Tasks' tab 'Edit Selected' option/radio button, and clicking this populates controls with data from DataGrid
  # Linked to XAML radio button.click: opt_Task_EditSelected  (this only appears in the 'Debug Mode' - consider deleting)
  
  opt_Task_EditSelected.FontWeight = FontWeights.Bold
  opt_Task_AddNew.FontWeight = FontWeights.Normal
  grp_PostponeOptions.Visibility = Visibility.Visible
  return


def get_taskStatusTypes():
  # new function Sept 2024 - returns a list of Task Status types
  # Amended into function to return list as this is what we need for populating data grid drop-down items
  
  xItem = []
  xItem.append(AssignToList(0, 'Not Started'))
  xItem.append(AssignToList(1, 'In Progress'))
  xItem.append(AssignToList(2, 'Completed'))
  xItem.append(AssignToList(3, 'Waiting on someone else'))
  xItem.append(AssignToList(4, 'Deferred'))
  return xItem

def populate_taskStatus(s, event):
  # This function populates the 'Task Status' combo box on the 'Tasks' tab
  # Note 'cbo_Task_Status' is in the 'Edit Area' (hidden behind 'Debug Mode') - consider deleting
  
  myItems = get_taskStatusTypes()
  cbo_Task_Status.ItemsSource = myItems
  return

def get_taskPriorityTypes():
  # This function returns a list of Task Priority types
  # Amended into function to return list as this is what we need for populating data grid drop-down items
  
  xItem = []
  xItem.append(AssignToList(0, 'Low'))
  xItem.append(AssignToList(1, 'Normal'))
  xItem.append(AssignToList(2, 'High'))
  return xItem

def populate_taskPriority(s, event):
  # This function populates the 'Task Priority' combo box on the 'Tasks' tab
  # Note 'cbo_Task_Priority' is in the 'Edit Area' (hidden behind 'Debug Mode') - consider deleting
  myItems = get_taskPriorityTypes()
  cbo_Task_Priority.ItemsSource = myItems
  return

class postponeOptions(object):
  def __init__(self, myFText, myDuration, myDType):
    self.pFriendlyText = myFText
    self.pDuration = myDuration
    self.pDurationType = myDType
    return

  def __getitem__(self, index):
    if index == 'FText':
      return self.pFriendlyText
    elif index == 'Duration':
      return self.pDuration
    elif index == 'DType':
      return self.pDurationType

def populate_taskPostponeOptions(s, event):
  # This function populates the 'Postpone' combo box on the 'Tasks' tab

  pItem = []
  pItem.append(postponeOptions('1 day', 1, 'DAY'))
  pItem.append(postponeOptions('2 days', 2, 'DAY'))
  pItem.append(postponeOptions('3 days', 3, 'DAY'))
  pItem.append(postponeOptions('4 days', 4, 'DAY'))
  pItem.append(postponeOptions('1 week', 1, 'WEEK'))
  pItem.append(postponeOptions('2 weeks', 2, 'WEEK'))
  pItem.append(postponeOptions('1 month', 1, 'MONTH'))
  pItem.append(postponeOptions('2 months', 2, 'MONTH'))
  ## NB: was going to use letters for third parameter, however, would make sense to use the number
  ## 0 = Minutes; 1 = Hours; 2 = Days; 3 = Weeks
  cbo_TaskPostpone.ItemsSource = pItem
  return

def task_PostponeNow(s, event):
  # This function is the action button against 'Postpone' options (on 'Tasks' tab) and will postpone the Task by the period specified
  # Linked to XAML button.click: btn_taskPostpone

  # if a postpone option hasn't been selected, alert user and exit
  if cbo_TaskPostpone.SelectedIndex == -1:
    MessageBox.Show("You haven't selected how long to postpone this task for!")
    return
  
  # get initial inputs as variables
  postponeDuration = cbo_TaskPostpone.SelectedItem['Duration']
  postponeType = cbo_TaskPostpone.SelectedItem['DType']
  originalDate = getSQLDate(dp_Task_DateDue.SelectedDate)
  newDueDate = getSQLDate(runSQL("SELECT DATEADD({0}, {1}, '{2}')".format(postponeType, postponeDuration, originalDate)))
  newDueDate = "{0} 09:00:00.000".format(newDueDate)
  currentStepID = dg_KeyTasks.SelectedItem['CaseStepID'] 

  # if Reminder Date is NOT 'none'
  if dp_Task_DateReminder.SelectedDate != None:
    # temp get the entered Reminder date and time
    tmpDate = str(getSQLDate(dp_Task_DateReminder.SelectedDate))
    tmpDate = tmpDate[:10]
    newRemindTime = getTimeFixed(txt_Task_TimeReminder.Text)
    originalRemindDate = "{0} {1}".format(tmpDate, newRemindTime)
    # get the new Reminder date
    newReminderDate = getSQLDate(runSQL("SELECT DATEADD({0}, {1}, '{2}')".format(postponeType, postponeDuration, originalRemindDate)))
    newReminderDate = "{0} {1}".format(newReminderDate, newRemindTime)
    
    # do actual update of Diary table
    runSQL("UPDATE Diary_Tasks SET DateStamp = '{0}', ReminderDate= '{1}', UpdateThis = {3} WHERE Code = {2}".format(newDueDate, newReminderDate, lbl_TaskRowID.Content, gUpdateThis))
  else:
    # do actual update of Diary table (without 'Reminder Date' as none given)
    runSQL("UPDATE Diary_Tasks SET DateStamp = '{0}', ReminderDate= NULL, UpdateThis = {1} WHERE Code = {2}".format(newDueDate, gUpdateThis, lbl_TaskRowID.Content))

  # do update to Case Manager table  
  runSQL("UPDATE Cm_Steps SET DiaryDate = '{0}' WHERE ItemID = {1}".format(newDueDate, currentStepID))

  refresh_KeyTasks(s, event)
  MessageBox.Show("Successfully postponed task")
  return


def task_cellSelection_Changed(s, event):
  # This function triggers when the datagrid selection on the 'Tasks' tab is changed
  # Linked to XAML control: dg_KeyTasks.SelectionChanged
  appendToDebugLog(textToAppend="entering task_cellSelection_Changed() event - copy vals from DG to individual hidden controls to right (visible in 'DebugMode')", inclTimeStamp=True)

  # if nothing selected
  if dg_KeyTasks.SelectedIndex == -1:
    appendToDebugLog(textToAppend = 'Nothing selected - setting values to null', inclTimeStamp = True)
    # nothing is selected - set values to empty/nothing
    cbo_Task_AssignedTo.SelectedIndex = -1
    cbo_Task_Status.SelectedIndex = 0
    cbo_Task_Priority.SelectedIndex = 1
    lbl_TaskRowID.Content = ''
    dp_Task_DateDue.SelectedDate = datetime.now()  #DateTime.Now    # None  # Possible error with: datetime.datetime.now()
    dp_Task_DateReminder.SelectedDate = None    #DateTime.Now
    txt_Task_TimeReminder.Text = ''
    txt_TaskDescription.Text = ''
    txt_Task_PercentComplete.Text = '0'
    txt_DateMissedNotes_Task.Text = ''
    
    # disable buttons that act on a selected item
    btn_MarkAsComplete_Task.IsEnabled = False
    btn_RevertTask.IsEnabled = False
    btn_DeleteTask.IsEnabled = False
    grp_PostponeOptions.Visibility = Visibility.Collapsed
    #tSep1.Visibility = Visibility.Collapsed

  else:
    # something IS selected from the list
    tmpGroup = dg_KeyTasks.SelectedItem['Group']
    
    # if item is marked as 'Completed'
    if tmpGroup == '2) Complete':
      # do not show edit area / set to 'Add new'
      opt_Task_AddNew.IsChecked = True
      btn_MarkAsComplete_Task.IsEnabled = False
      btn_RevertTask.IsEnabled = True
      btn_DeleteTask.IsEnabled = False
      grp_PostponeOptions.Visibility = Visibility.Collapsed
      #tSep1.Visibility = Visibility.Collapsed

    else:
      # item is NOT marked as 'Completed' (eg: is outstanding) - put values from list into 'Edit selected' area at right of form (from the selected DG item)
      opt_Task_EditSelected.IsChecked = True
      btn_MarkAsComplete_Task.IsEnabled = True
      btn_RevertTask.IsEnabled = False
      btn_DeleteTask.IsEnabled = True
      grp_PostponeOptions.Visibility = Visibility.Visible
      #tSep1.Visibility = Visibility.Visible

      #appendToDebugLog(textToAppend = 'Setting Code, Reminder Time, Description, Percent Complete and DateMissed Note', inclTimeStamp = True)
      lbl_TaskRowID.Content = dg_KeyTasks.SelectedItem['Code']
      txt_Task_TimeReminder.Text = str(dg_KeyTasks.SelectedItem['ReminderTime'])
      txt_TaskDescription.Text = str(dg_KeyTasks.SelectedItem['Desc'])
      txt_Task_PercentComplete.Text = str(dg_KeyTasks.SelectedItem['PercentComplete'])
      txt_DateMissedNotes_Task.Text = str(dg_KeyTasks.SelectedItem['DateMissedNote'])

      #appendToDebugLog(textToAppend = 'Setting DueDate', inclTimeStamp = True)
      # if 'DueDate' in DG list is empty, set selected date to 'Select a date' (None)
      tmpDate = dg_KeyTasks.SelectedItem['Date']
      if tmpDate is None or tmpDate == '':
        dp_Task_DateDue.SelectedDate = None
      else:
        # there is a date so use this
        dp_Task_DateDue.SelectedDate = tmpDate

      #appendToDebugLog(textToAppend = 'Setting ReminderDate', inclTimeStamp = True)
      # if 'ReminderDate' in DG list is empty, set selected date to 'Select a date' (None)
      #if dg_KeyTasks.SelectedItem['ReminderDate'] == None:
      tmpRDate = dg_KeyTasks.SelectedItem['ReminderDate']
      if tmpRDate is None or tmpRDate == '':
        dp_Task_DateReminder.SelectedDate = None
        dp_Task_DateReminder.DisplayDateEnd = None
      else:
        # there IS a date, so use this
        dp_Task_DateReminder.SelectedDate = tmpRDate

        #appendToDebugLog(textToAppend = 'Setting DisplayDateEnd for Reminder Date', inclTimeStamp = True)
        if tmpDate is None:
          dp_Task_DateReminder.DisplayDateEnd = None
        else:
          dp_Task_DateReminder.DisplayDateEnd = tmpDate

      # populate 'Assigned To' combo box - first need to extract user code
      uCode = dg_KeyTasks.SelectedItem['AssignedTo']
      #appendToDebugLog(textToAppend = 'Setting AssignedTo: {0}'.format(uCode), inclTimeStamp = True)
      tCount = -1
      for tItem in cbo_Task_AssignedTo.Items:
        tCount += 1
        if tItem.iCode == uCode:
          cbo_Task_AssignedTo.SelectedIndex = tCount
          break

      tStatus = dg_KeyTasks.SelectedItem['Status']
      #appendToDebugLog(textToAppend = 'Setting Status: {0}'.format(tStatus), inclTimeStamp = True)
      tCount = -1
      for tItem in cbo_Task_Status.Items:
        tCount += 1
        if tItem.iCode == tStatus:
          cbo_Task_Status.SelectedIndex = tCount
          break
      
      tPriority = dg_KeyTasks.SelectedItem['Priority']
      #appendToDebugLog(textToAppend = 'Setting Priority: {0}'.format(tPriority), inclTimeStamp = True)
      tCount = -1
      for tItem in cbo_Task_Priority.Items:
        tCount += 1
        if tItem.iCode == tPriority:
          cbo_Task_Priority.SelectedIndex = tCount
          break
  return


def deleteTask(s, event):
  # This function will delete the currently selected Task (on 'Tasks' tab) - after confirmation from user
  # Linked to XAML button.click: btn_DeleteTask

  if dg_KeyTasks.SelectedIndex != -1:
    myMessage = "Are you sure you want to remove the following Task?\n{0}".format(dg_KeyTasks.SelectedItem['Desc'])
    result = MessageBox.Show(myMessage, 'Confirm deletion of Task...', MessageBoxButtons.YesNo)
  
    if result == DialogResult.Yes:
      # if task has a CaseItem ID then we will want to remove from there too (as well as Diary_Tasks)
      # Come to think of it, users will not have ability to delete case items, therefore wondering how this would play out - would it error (we shall do anyway)
      tmpCaseStepID = 0 if dg_KeyTasks.SelectedItem['CaseStepID'] == None else dg_KeyTasks.SelectedItem['CaseStepID']
      tmpAgendaID = 0 if dg_KeyTasks.SelectedItem['Agenda'] == None else dg_KeyTasks.SelectedItem['Agenda']
      tmpKDid = dg_KeyTasks.SelectedItem['RowID']

      if tmpCaseStepID > 0:
        # get the order ID as we'll need this later to update order of case items table
        tmpCIorder = runSQL("SELECT ItemOrder FROM Cm_CaseItems WHERE ItemID = {0}".format(tmpCaseStepID))
        # Update DeleteThis in Diary_Tasks (to delete Outlook item)
        runSQL("UPDATE Diary_Tasks SET DeleteThis = 1 WHERE CaseItemRef = {0}".format(tmpCaseStepID))

        # Now form and run our DELETE FROM tables
        runSQL("DELETE FROM Cm_Steps WHERE ItemID = {0}".format(tmpCaseStepID))
        runSQL("DELETE FROM Cm_CaseItems WHERE ItemID = {0}".format(tmpCaseStepID))
        # for sake of tidying up database, ought to also delete from CaseActionHistory table too
        runSQL("DELETE FROM Cm_Steps_ActionHistory WHERE ItemID = {0}".format(tmpCaseStepID))

        if tmpAgendaID > 0:
          # Update ItemOrder - decrease all items with Order GREATER than selected item
          runSQL("UPDATE Cm_CaseItems SET ItemOrder = (ItemOrder - 1) WHERE ParentID = {0} AND ItemOrder > {1}".format(tmpAgendaID, tmpCIorder))

      if tmpKDid != None and tmpKDid > 0:
        # Delete the item from our Key Dates table (We do this regardless of whether Case item exists or not)
        if tmpCaseStepID > 0:
          runSQL("DELETE FROM Usr_Key_Dates WHERE TaskStepID = {0} AND ID = {1}".format(tmpCaseStepID, tmpKDid))
        else:
          runSQL("DELETE FROM Usr_Key_Dates WHERE ID = {0}".format(tmpKDid))

      # now refresh the datagrid to reflect updates
      refresh_KeyTasks(s, event)
      # and refresh defaults list as an item may now be 'Available'
  return


def revertTask(s, event):
  # This function will revert (re-add) the selected 'Completed' Task
  # Linked to XAML button.click: btn_RevertTask
  # Add to CaseAction History table (with 'reverted' number (2))
  # Add to 'Diary Tasks' table (with applicable Case Item ID)
  # nullify 'completion date' in CaseItems table

  # if nothing selected, exit now
  if dg_KeyTasks.SelectedIndex == -1:
   return

  # if items group is not 'Complete' then exit now (only act upon completed items)
  if dg_KeyTasks.SelectedItem['Group'] != '2) Complete':
    return

  # get variables needed for SQL
  caseItemID = dg_KeyTasks.SelectedItem['CaseStepID']
  feREF = runSQL("SELECT FeeEarnerRef FROM Matters WHERE EntityRef = '{0}' AND Number = {1}".format(_tikitEntity, _tikitMatter))
  agendaID = dg_KeyTasks.SelectedItem['Agenda']
  taskDueDate = getSQLDate(dg_KeyTasks.SelectedItem['Date'])
  # we ought to do a check here to ensure we don't add diary tasks in the past
  # eg: check if above date is before todays date, if not, get tomorrows date instead
  newDueDate = getSQLDate(runSQL("SELECT CASE WHEN '{0}' < GETDATE() THEN DATEADD(day, 1, GETDATE()) ELSE '{0}' END".format(taskDueDate)))

  fullTaskDate = "{0} 09:00:00.000".format(newDueDate)
  taskDesc = getUniqueDescription(desiredDesc = dg_KeyTasks.SelectedItem['Desc'], forUser = feREF, taskORdate = 'Task') 
  taskDescSQLs = sql_safe_string(stringToClean = taskDesc)
  tRemindDate = getSQLDate(dg_KeyTasks.SelectedItem['ReminderDate'])
  tRemindTime = getTimeFixed(dg_KeyTasks.SelectedItem['ReminderTime'])
  # as per above check against due date, check this and if necessary get tomorrows date
  newRemindDate = getSQLDate(runSQL("SELECT CASE WHEN '{0}' < GETDATE() THEN DATEADD(day, 1, GETDATE()) ELSE '{0}' END".format(tRemindDate)))

  tRemindDate = "{0} {1}.000".format(newRemindDate, tRemindTime)
  tDateMissedN = sql_safe_string(stringToClean = dg_KeyTasks.SelectedItem['DateMissedNotes'])

  # likewise with date checks above, do we really want to revert exact same status etc... I'd think we would rather set these to default (not complete etc)
  tStatus = 0      #dg_KeyTasks.SelectedItem['Status']
  tPriority = 1    #dg_KeyTasks.SelectedItem['Priority']
  tPCComplete = 0  #dg_KeyTasks.SelectedItem['PercentComplete']

  kdID = dg_KeyTasks.SelectedItem['RowID']
  if kdID == None or kdID == 0:
    kdID = get_KeyDatesTableID(caseItemID)

  # make a description for Cm Steps History table
  rDesc = "Key Date - Original details: Description: {0}; DueDate: {1}; RemindDate: {2}".format(taskDescSQLs, fullTaskDate, tRemindDate)

  # update Steps ActionHistory table to state current user is reverting this task
  runSQL("INSERT INTO Cm_Steps_ActionHistory (ItemID, UserID, StepActionID, Description) VALUES ({0}, '{1}', {2}, '{3}')".format(caseItemID, _tikitUser, 2, rDesc))

  # add new entry into Diary Tasks table (note: use FeeEarner of matter rather than current user when setting person assigned to)
  tasksSQL = """INSERT INTO Diary_Tasks (Username, DateStamp, [Description], EntityRef, MatterNoRef, AgendaRef, CaseItemRef, UserType, Status, Priority, Complete, ReminderDate) 
                VALUES('{0}', '{1}', '{2}', '{3}', {4}, {5}, {6}, 'A', {7}, {8}, {9}, '{10}')""".format(feREF, fullTaskDate, taskDescSQLs, _tikitEntity, _tikitMatter, agendaID, 
                                                                                                        caseItemID, tStatus, tPriority, tPCComplete, tRemindDate)
  runSQL(tasksSQL)

  # now we need to get the ID of the added item
  taskTblID = runSQL("SELECT Code FROM Diary_Tasks WHERE Username = '{0}' AND Description = '{1}' AND DateStamp = '{2}' AND EntityRef = '{3}' AND MatterNoRef = {4}".format(feREF, taskDesc, fullTaskDate, _tikitEntity, _tikitMatter))

  # remove completion date on Case items table
  runSQL("UPDATE Cm_CaseItems SET CompletionDate = NULL WHERE ItemID = {0}".format(caseItemID))

  # update due date in Steps table (after checking in tables, it is the Task Due Date we want here, not the reminder)
  runSQL("UPDATE Cm_Steps SET DiaryDate = '{0}' WHERE ItemID = {1}".format(fullTaskDate, caseItemID))

  # update Key Dates table (are there any other fields we need to update here)
  updateKD_SQL = """UPDATE Usr_Key_Dates SET AssignedTo = '{0}', Date = '{1}', DurationQty = 1, DurationTypeN = 2, 
                    Description = '{2}', ReminderQty = 15, ReminderTypeN = 0, ReminderDate = '{3}', ReminderTime = '{4}', 
                    Status = {5}, Priority = {6}, PCComplete = {7}, DateMissedNotes = '{8}', TaskOrDate = 'Task' WHERE ID = {9}""".format(feREF, fullTaskDate, taskDescSQLs, tRemindDate, tRemindTime, 
                                                                                                                                          tStatus, tPriority, tPCComplete, tDateMissedN, kdID)

  if debugMessage(msgTitle = 'DEBUG MESSAGE - Testing Task Revert', msgBody = updateKD_SQL) == True:
    runSQL(updateKD_SQL, True, "There was an error updating the KeyDates table", "ERROR - KeyDate Revert...")

  # finally refresh the key tasks datagrid
  refresh_KeyTasks(s, event)

  # it would be nice to select this item again, as it's going to move to 'outstanding' group
  tCount = -1
  for tRow in dg_KeyTasks.Items:
    tCount += 1
    if tRow.iCaseStepID == caseItemID:
      dg_KeyTasks.SelectedIndex = tCount
      break
  return


def dg_KeyTasks_cellEdit_Finished(s, event):
  # New - September 2024 - Louis wants to be able to edit via the DataGrid itself, rather than via the editable controls at the bottom.
  # Therefore, this function splices up the 'addOrUpdate_KeyTask' function (line 318) to ensure we both validate input and update applicable tables.
  # Note - shouldn't allow updating of 'Completed' items other than the 'DateMissedNote'
  # Linked to XAML DataGrid.CellEditFinished: dg_KeyTasks
  
  # firstly, store column name to variable
  tmpColName = event.Column.Header
  appendToDebugLog(textToAppend="entering dg_KeyTasks_cellEdit_Finished() event - column changed: {0}".format(tmpColName), inclTimeStamp=True)

  # setup initial variables (SQL to update tables, get ID for respective table, reset count of updates per table)
  uSQL_Diary = "UPDATE Diary_Tasks SET "
  uSQL_CmS = "UPDATE Cm_Steps SET "
  uSQL_CmCI = "UPDATE Cm_CaseItems SET "
  uSQL_UsrKD = "UPDATE Usr_Key_Dates SET "
  dCount = cmsCount = cmciCount = usrKDCount = newStatus = 0
  tmpDebugTitle = "DEBUGGING - Editing in DataGrid..."

  # values from DataGrid
  mStatus = dg_KeyTasks.SelectedItem['Group']
  diaryID = dg_KeyTasks.SelectedItem['Code']
  cMsID = dg_KeyTasks.SelectedItem['CaseStepID']
  uKDID = dg_KeyTasks.SelectedItem['KDid']

  # Conditionally add parts depending on column updated and whether value has changed
  if tmpColName == 'Description' and mStatus != '2) Complete':
    newName = dg_KeyTasks.SelectedItem['Desc']

    if newName != txt_TaskDescription.Text:
      # make text SQL safe (replace single quotes with double)
      newName = sql_safe_string(stringToClean = newName)
      uSQL_Diary += "Description = '{0}' ".format(newName) 
      dCount += 1
      uSQL_CmCI += "Description = '{0}' ".format(newName) 
      cmciCount += 1
      uSQL_UsrKD += "Description = '{0}' ".format(newName) 
      usrKDCount += 1
      #debugMessage(msgTitle = tmpDebugTitle, msgBody = "Diary SQL:\n{0}\n\nCaseItems SQL:\n{1}\n\nKeyDates SQL:\n{2}".format(uSQL_Diary, uSQL_CmCI, uSQL_UsrKD))

  if tmpColName == 'Assigned To' and mStatus != '2) Complete':
    newAssignedTo = dg_KeyTasks.SelectedItem['AssignedTo']
    #MessageBox.Show("newAssignedTo: {0}".format(newAssignedTo), "DEBUGGING - Editing in DataGrid")
    
    #if newAssignedTo != cbo_Date_AssignedTo.SelectedItem['Code']:
    # following our revelation with the 'Duration' drop-down (changing it in DG immediately updated 'combo box' in 'edit' area), we can't test
    # if value is different to combo box as they will alwats be the same value and code would never trigger. So just updating always.
    uSQL_Diary += "Username = '{0}' ".format(newAssignedTo) 
    dCount += 1
    uSQL_UsrKD += "AssignedTo = '{0}' ".format(newAssignedTo) 
    usrKDCount += 1
    #debugMessage(msgTitle = tmpDebugTitle, msgBody = "Diary SQL:\n{0}\n\n KeyDate SQL:\n{1}".format(uSQL_Diary, uSQL_UsrKD))

  if tmpColName == 'Due Date' and mStatus != '2) Complete':
    newDueDate = getSQLDate(dg_KeyTasks.SelectedItem['Date'])
    tmpDDate = getSQLDate(dp_Task_DateDue.SelectedDate)
    tmpDDate = tmpDDate[:10]
    newTime = "09:00:00.000"
    actualDate = "{0} {1}".format(newDueDate, newTime)
    if newDueDate != tmpDDate:
      uSQL_Diary += "DateStamp = '{0}' ".format(actualDate) 
      dCount += 1
      uSQL_CmS += "DiaryDate = '{0}' ".format(actualDate) 
      cmsCount += 1
      # NB we also update the StepCategory once date has been entered
      if mStatus[:1] == '0':
        uSQL_CmS += ", StepCategory = 'KeyD' " 
      uSQL_UsrKD += "Date = '{0}' ".format(actualDate) 
      usrKDCount += 1
      #debugMessage(msgTitle = tmpDebugTitle, msgBody = "Diary SQL:\n{0}\n\nCmSteps SQL:\n{1}\n\nKey Dates SQL:\n{2}".format(uSQL_Diary, uSQL_CmS, uSQL_UsrKD))

    # check for Reminder Date
    if dg_KeyTasks.SelectedItem['InclReminder'] == 'N':
      if dCount > 0:
        uSQL_Diary += ", "

      uSQL_Diary += "ReminderDate = NULL "
      dCount += 1

      if usrKDCount > 0:
        uSQL_UsrKD += ", "
      uSQL_UsrKD += "ReminderDate = NULL, ReminderTime = NULL "
      usrKDCount += 1

    else:
      # have opted to have a reminder date, so check to ensure this is BEFORE the Due Date
      # get values
      tmpRemDate = dg_KeyTasks.SelectedItem['ReminderDate']
      tmpTime = dg_KeyTasks.SelectedItem['ReminderTime']

      # if current Reminder Date is nothing
      if tmpRemDate is None:
        # use Due Date for Reminder Date
        xNewRemindDate = actualDate
      else:
        # else: get current Reminder Date and strip off any time element
        xNewRemindDate = getSQLDate(tmpRemDate)
        if len(xNewRemindDate) > 10:
          xNewRemindDate = xNewRemindDate[:10]

      # if Reminder Time is empty string or nothing
      if tmpTime is None or tmpTime == '':
        # set newTime to our default
        newTime = getTimeFixed("09:00")
      else:
        # else: use time as-is
        newTime = getTimeFixed(tmpTime)
      # create full Date and Time of Reminder
      finalRemDateTime = "{0} {1}".format(xNewRemindDate, newTime)

      # get number of days between both dates
      # note: a POSITIVE number indicates Reminder Date is AFTER Due Date
      #       a NEGATIVE number indicates Reminder Date is BEFORE Due Date
      daysDiff = runSQL("SELECT DATEDIFF(day, '{0}', '{1}')".format(actualDate, finalRemDateTime))
      
      if int(daysDiff) > 0:
        # Reminder Date is AFTER 'Due Date', so set to 'Due Date'
        if dCount > 0:
          uSQL_Diary += ", "
        uSQL_Diary += "ReminderDate = '{0}' ".format(actualDate)
        dCount += 1
        if usrKDCount > 0:
          uSQL_UsrKD += ", "
        uSQL_UsrKD += "ReminderDate = '{0}', ReminderTime = {1} ".format(actualDate, newTime)
        usrKDCount += 1
        #update_ReminderDate(xDiaryID = diaryID, kdID = uKDID, newRemindDate = newDueDate, newRemindTime = newTime, oldRemindDate = None, oldRemindTime = None)



  # new 'Include reminder' check box...
  if tmpColName == 'Include Reminder?' and mStatus != '2) Complete':
    appendToDebugLog(textToAppend="CellEdit_Finished - Include Reminder changed - current value: '{0}'".format(dg_KeyTasks.SelectedItem['InclReminder']), inclTimeStamp=True)
    if dg_KeyTasks.SelectedItem['InclReminder'] == 'N':
      # set reminder date to null
      uSQL_Diary += "ReminderDate = NULL "
      dCount += 1
      uSQL_UsrKD += "ReminderDate = NULL "
      usrKDCount += 1
      appendToDebugLog(textToAppend="CellEdit_Finished - Include Reminder changed - Setting to null", inclTimeStamp=True)
    else:
      # set reminder time defaults
      #newDueDate = getDefaultNextDay()
      # used to be 'get tomorrow', however, would be better if it matched the 'DueDate'
      tmpDueDate = getSQLDate(dg_KeyTasks.SelectedItem['Date'])
      tmpDueDate = tmpDueDate[:10]
      tmpTime = "09:00:00.000"
      newDueDate = "{0} {1}".format(tmpDueDate, tmpTime)

      appendToDebugLog(textToAppend="CellEdit_Finished - Include Reminder changed - New Reminder date: '{0}'".format(newDueDate), inclTimeStamp=True)
      # slight issue with the above is that whilst ideally we want say 'tomorrows' date, we do have a limitation on the 
      # date (it cannot be after the due date). Therefore, may need to consider how to handle this, because other option is
      # to set to 1 day before the Due Date - - but if that date has passed
      uSQL_Diary += "ReminderDate = '{0}' ".format(newDueDate) 
      dCount += 1
      uSQL_UsrKD += "ReminderDate = '{0}' ".format(newDueDate) 
      usrKDCount += 1


  if tmpColName == 'Reminder Date' and mStatus != '2) Complete':
    # REALLY OUGHT TO ADD SOME LOGIC IN HERE SO THAT THE REMINDER DATE CANNOT BE AFTER THE DUE DATE
    # (have added code to XAML to set the DisplayMaxDate to the same as 'iDate' - works perfectly)
    #update_ReminderDate(xDiaryID = diaryID, kdID = uKDID, newRemindDate = dg_KeyTasks.SelectedItem['ReminderDate'], newRemindTime = dg_KeyTasks.SelectedItem['ReminderTime'], 
    #                                           oldRemindDate = dp_Task_DateReminder.SelectedDate, oldRemindTime = dg_KeyTasks.SelectedItem['oldRemindTime'])

    tmpRemDate = dg_KeyTasks.SelectedItem['ReminderDate']
    tmpTime = dg_KeyTasks.SelectedItem['ReminderTime']
    #appendToDebugLog(textToAppend="CellEdit_Finished - ReminderDate changed - current value: '{0}' (type: {2}) Time: '{1}'".format(tmpRemDate, tmpTime, type(tmpRemDate)), inclTimeStamp=True)
    if tmpRemDate is None:  # or tmpRemDate == '':
      # Reminder Date was empty, so set to 'null'
      uSQL_Diary += "ReminderDate = NULL "
      dCount += 1
      uSQL_UsrKD += "ReminderDate = NULL "
      usrKDCount += 1
    else:
      # Reminder Date was entred - as this comes with time on the end, get just the first 10 characters
      newDueDate = getSQLDate(tmpRemDate)
      if len(newDueDate) > 10:
        newDueDate = newDueDate[:10]
    
      if tmpTime is None or tmpTime == '':
        newTime = getTimeFixed("09:00")
      else:
        newTime = getTimeFixed(tmpTime)
      actualDate = "{0} {1}".format(newDueDate, newTime)
    
      # get old (before changed) date
      tmpOldDate = dp_Task_DateReminder.SelectedDate
      if tmpOldDate is None or tmpOldDate == '':
        # just update as nothing provided before so can't compare against anything
        uSQL_Diary += "ReminderDate = '{0}' ".format(actualDate) 
        dCount += 1
        uSQL_UsrKD += "ReminderDate = '{0}' ".format(actualDate) 
        usrKDCount += 1
      else:
        # there was a date previously, so get that
        oldDueDate = getSQLDate(tmpOldDate)
        if len(oldDueDate) > 10:
          oldDueDate = oldDueDate[:10]
    
        oldTime = getTimeFixed(tmpTime)
        fullOldDate = "{0} {1}".format(oldDueDate, oldTime)
    
        if actualDate != fullOldDate:
          uSQL_Diary += "ReminderDate = '{0}' ".format(actualDate) 
          dCount += 1
          uSQL_UsrKD += "ReminderDate = '{0}' ".format(actualDate) 
          usrKDCount += 1
          #debugMessage(msgTitle = tmpDebugTitle, msgBody = "Diary SQL:\n{0}\n\nCmSteps SQL:\n{1}\n\nKey Dates SQL:\n{2}".format(uSQL_Diary, uSQL_CmS, uSQL_UsrKD))


  if tmpColName == 'Reminder Time' and mStatus != '2) Complete':
    tmpTime = dg_KeyTasks.SelectedItem['ReminderTime']
    oldTime = dg_KeyTasks.SelectedItem['oldRemindTime']
    tmpDate = dp_Task_DateReminder.SelectedDate

    if tmpTime == None or tmpTime == '':
      newTime = getTimeFixed("09:00")
    else:
      newTime = getTimeFixed(tmpTime)
    
    if oldTime == None or oldTime == '':
      oldTime = getTimeFixed("09:00")
    else:
      oldTime = getTimeFixed(oldTime)

    tmpDDate = getSQLDate(tmpDate)
    tmpDDate = tmpDDate[:10]
    actualDate = "{0} {1}".format(tmpDDate, newTime)

    if newTime != oldTime:
      uSQL_Diary += "ReminderDate = '{0}' ".format(actualDate) 
      dCount += 1
      uSQL_UsrKD += "ReminderTime = {0} ".format(newTime) 
      usrKDCount += 1
      #debugMessage(msgTitle = tmpDebugTitle, msgBody = "Diary SQL:\n{0}\n\nCmSteps SQL:\n{1}\n\nKey Dates SQL:\n{2}".format(uSQL_Diary, uSQL_CmS, uSQL_UsrKD))


  if tmpColName == 'Status' and mStatus != '2) Complete':
    newStatus = dg_KeyTasks.SelectedItem['Status']
    oldStatus = cbo_Task_Status.SelectedItem['Code']
    #if newStatus != oldStatus:
    # see comments against other combo boxes in datagrid - cannot test against values in edit area (unless we create new labels to store info)
    uSQL_Diary += "Status = {0} ".format(newStatus)
    dCount += 1
    uSQL_UsrKD += "Status = {0} ".format(newStatus) 
    usrKDCount += 1
    #debugMessage(msgTitle = tmpDebugTitle, msgBody = "Diary SQL:\n{0}\n\nCmSteps SQL:\n{1}\n\nKey Dates SQL:\n{2}".format(uSQL_Diary, uSQL_CmS, uSQL_UsrKD))


  if tmpColName == 'Priority' and mStatus != '2) Complete':
    newPriority = dg_KeyTasks.SelectedItem['Priority']
    oldPriority = cbo_Task_Priority.SelectedItem['Code']
    #if newPriority != oldPriority:
    # see comments against other combo boxes in datagrid - cannot test against values in edit area (unless we create new labels to store info)
    uSQL_Diary += "Priority = {0} ".format(newPriority)
    dCount += 1
    uSQL_UsrKD += "Priority = {0} ".format(newPriority) 
    usrKDCount += 1
    #debugMessage(msgTitle = tmpDebugTitle, msgBody = "Diary SQL:\n{0}\n\nCmSteps SQL:\n{1}\n\nKey Dates SQL:\n{2}".format(uSQL_Diary, uSQL_CmS, uSQL_UsrKD))

  if tmpColName == '% Complete' and mStatus != '2) Complete':
    newName = dg_KeyTasks.SelectedItem['PercentComplete']

    if newName != txt_Task_PercentComplete.Text:
      uSQL_Diary += "Complete = '{0}' ".format(newName) 
      dCount += 1
      uSQL_UsrKD += "PCComplete = '{0}' ".format(newName) 
      usrKDCount += 1
      #debugMessage(msgTitle = tmpDebugTitle, msgBody = "Diary SQL:\n{0}\n\nCaseItems SQL:\n{1}\n\nKeyDates SQL:\n{2}".format(uSQL_Diary, uSQL_CmCI, uSQL_UsrKD))

  if tmpColName == 'Date Missed Notes':
    newDMN = dg_KeyTasks.SelectedItem['DateMissedNote']
    tmpDMN = txt_DateMissedNotes_Task.Text
    
    if newDMN != tmpDMN:
      # make SQL safe text (replace single quote with double)
      newDMN = sql_safe_string(stringToClean = newDMN)
      uSQL_UsrKD += "DateMissedNotes = '{0}' ".format(newDMN)
      usrKDCount += 1
      #debugMessage(msgTitle = tmpDebugTitle, msgBody = "Date Missed Notes SQL:\n{0}".format(uSQL_UsrKD))


  # now for the actual updates
  if dCount > 0: 
    uSQL_Diary += ", UpdateThis = {3} WHERE Code = {0} AND EntityRef = '{1}' AND MatterNoRef = {2}".format(diaryID, _tikitEntity, _tikitMatter, gUpdateThis)
    if debugMessage(msgTitle = tmpDebugTitle, msgBody = "Diary Table SQL:\n{0}".format(uSQL_Diary)) == True:
      runSQL(uSQL_Diary)

  if cmsCount > 0:
    uSQL_CmS += "WHERE ItemID = {0}".format(cMsID)
    if debugMessage(msgTitle = tmpDebugTitle, msgBody = "Cm_Steps Table SQL:\n{0}".format(uSQL_CmS)) == True:
      runSQL(uSQL_CmS)

  if cmciCount > 0:
    uSQL_CmCI += "WHERE ItemID = {0}".format(cMsID)
    if debugMessage(msgTitle = tmpDebugTitle, msgBody = "Cm_CaseItems Table SQL:\n{0}".format(uSQL_CmCI)) == True:
      runSQL(uSQL_CmCI)

  if usrKDCount > 0:
    if int(uKDID) == 0:
      uKDID = get_KeyDatesTableID(caseItemID = cMsID)
    uSQL_UsrKD += "WHERE ID = {0}".format(uKDID)
    if debugMessage(msgTitle = tmpDebugTitle, msgBody = "Usr_Key_Dates Table SQL:\n{0}".format(uSQL_UsrKD)) == True:
      runSQL(uSQL_UsrKD)

  # Louis added this within main function above (line 1023) and wondering if this was causing issues as we hadn't finished actual updates needed first, and as mostly based on 'selectedItem' and we do a refresh there, code here didn't know what's going on
  # If new status = 'Complete', then we ought to do actual 'mark as complete' function
  if newStatus == 2:
    # following function does a refresh on KeyTasks datagrid
    task_MarkComplete(s, event)
  else:
    # as new status isn't 'Complete', refresh the KeyTasks datagrid
    refresh_KeyTasks(s, event)

  # select row again as we refreshed
  xCount = -1
  for xRow in dg_KeyTasks.Items:
    xCount += 1
    if xRow.iCaseStepID == cMsID:
      dg_KeyTasks.SelectedIndex = xCount
      break
  # I like the idea of the above, however, I wonder if causing a circular issue, so switching off for now
  # Think the issue is with the TIME, as getting odd value passed.
  return

def update_ReminderDate(xDiaryID = 0, kdID = 0, newRemindDate = None, newRemindTime = None, oldRemindDate = None, oldRemindTime = None):
  # This function will update the ReminderDate (tasks)
  
  newValueDate = ''
  newValueTime = ''
  newDateAndTime = ''
  oldValueDate = ''
  oldValueTime = ''
  oldDateAndTime = ''
  SQL_Diary = ''
  SQL_UsrKD = ''

  if newRemindDate is None:
    # we set to null
    SQL_Diary = "UPDATE Diary_Tasks SET ReminderDate = NULL, UpdateThis = {0} WHERE Code = {1} AND EntityRef = '{2}' AND MatterNoRef = {3}".format(gUpdateThis, xDiaryID, _tikitEntity, _tikitMatter)
    SQL_UsrKD = "UPDATE Usr_Key_Dates SET ReminderDate = NULL WHERE ID = {0}".format(kdID)

  else:
    # we do have a new date, so get SQL version - as this comes with time on the end, get first 10 chars
    newValueDate = getSQLDate(newRemindDate)
    if len(newValueDate) > 10:
      newValueDate = newValueDate[:10]
    
    # if newRemindTime was not supplied or empty string
    if newRemindTime is None or newRemindTime == '':
      # set time to 9am
      newValueTime = getTimeFixed("09:00")
    else:
      # new Time was provided, so use that
      newValueTime = getTimeFixed(newRemindTime)
    # create final new 'Date and Time'
    newDateAndTime = "{0} {1}".format(newValueDate, newValueTime)

    # if old remind date wasn't supplied or empty string provided
    if oldRemindDate is None or oldRemindDate == '':
      # update as is as nothing to compare against
      SQL_Diary = "UPDATE Diary_Tasks SET ReminderDate = '{0}', UpdateThis = {1}, WHERE Code = {2} AND EntityRef = '{3}' AND MatterNoRef = {4}".format(newDateAndTime, gUpdateThis, xDiaryID, _tikitEntity, _tikitMatter)
      SQL_UsrKD = "UPDATE Usr_Key_Dates SET ReminderDate = '{0}' WHERE ID = {1}".format(newDateAndTime, kdID)

    else:
      # there was a date previously, so get that (and only first 10 chars because of time)
      oldValueDate = getSQLDate(oldRemindDate)
      if len(oldValueDate) > 10:
        oldValueDate = oldValueDate[:10]
      
      # if an old reminder time wasn't supplied or empty string provided
      if oldRemindTime is None or oldRemindTime == '':
        # set old time to 9am
        oldValueTime = getTimeFixed("09:00")
      else:
        # old time was provided so use that
        oldValueTime = getTimeFixed(oldRemindTime)
      # create final 'old Date and Time'
      oldDateAndTime = "{0} {1}".format(oldValueDate, oldValueTime)

      # if new date and time doesn't match old date and time
      if newDateAndTime != oldDateAndTime:
        # update
        SQL_Diary = "UPDATE Diary_Tasks SET ReminderDate = '{0}', UpdateThis = {1}, WHERE Code = {2} AND EntityRef = '{3}' AND MatterNoRef = {4}".format(newDateAndTime, gUpdateThis, xDiaryID, _tikitEntity, _tikitMatter)
        SQL_UsrKD = "UPDATE Usr_Key_Dates SET ReminderDate = '{0}' WHERE ID = {1}".format(newDateAndTime, kdID)

  if len(SQL_Diary) > 0:
    runSQL(SQL_Diary)
  if len(SQL_UsrKD) > 0:
    runSQL(SQL_UsrKD)

  return

# END OF:  T A S K S   section
#################################################################################################################################


def getSQLDate(varDate):
  #Converts the passed varDate into SQL version date (YYYY-MM-DD)

  newDate = ''
  tmpDate = ''
  tmpDay = ''
  tmpMonth = ''
  tmpYear = ''
  mySplit = []
  finalStr = ''
  canContinue = False

  # If passed value is of 'DateTime' then convert to string
  if isinstance(varDate, DateTime) == True:
    tmpDate = varDate.ToString()
    canContinue = True

  # else if a 'datetime'
  #elif isinstance(varDate, datetime.datetime):
  #  tmpDate = varDate.strftime("%Y-%m-%d")
  #  finalStr = tmpDate
  #  canContinue = False

  elif isinstance(varDate, datetime):
    tmpDate = varDate.strftime("%Y-%m-%d")
    finalStr = tmpDate
    canContinue = False

  # else if already a string, assign passed date directly into newDate 
  elif isinstance(varDate, str) == True:
    tmpDate = varDate
    canContinue = True

  if canContinue == True:
    # now to strip out the time element
    mySplit = []
    mySplit = tmpDate.split(' ')
    newDate = mySplit[0]

    #MessageBox.Show('newDate is ' + newDate)
    mySplit = []

    if len(newDate) >= 8:
      mySplit = newDate.split('/')

      tmpDay = mySplit[0]             #newDate.strftime("%d")
      tmpMonth = mySplit[1]           #newDate.strftime("%m")
      tmpYear = mySplit[2]            #newDate.strftime("%Y")

      testStr = '{0}-{1}-{2}'.format(tmpYear, tmpMonth, tmpDay)
        #MessageBox.Show('Original: ' + str(varDate) + '\nFinal: ' + testStr)
        #newDate1 = datetime.datetime(int(tmpYear), int(tmpMonth), int(tmpDay))
        #finalStr = newDate1.strftime("%Y-%m-%d")
      finalStr = testStr

    return finalStr


#################################################################################################################
# # # #   D E F A U L T S   -   S E C T I O N   # # # #

class KDDefaults(object):
  def __init__(self, myTick, myOrder, myDescription, myRemindDays, myLinkedMPfield, myGroup):
    self.iTicked = myTick
    self.iOrder = myOrder
    self.iDesc = myDescription
    self.iLinkedMPField = myLinkedMPfield
    self.iRemindDaysBefore = myRemindDays
    self.iGroup = myGroup
    return

  def __getitem__(self, index):
    if index == 'Ticked':
      return self.iTicked
    elif index == 'Order':
      return self.iOrder
    elif index == 'Desc':
      return self.iDesc
    elif index == 'LinkMPField':
      return self.iLinkedMPField
    elif index == 'RemindDaysBefore':
      return self.iRemindDaysBefore
    elif index == 'Group':
      return self.iGroup

def refresh_KDDefaults_List(s, event):
  # This function will populate the 'Default for Current Case Type' tab on the XAML.
  # Defaults can be set against a specific CaseType if desired, but if nothing specified at CaseType level, this function will display all at Department level

  # firstly get the 'CaseType' and 'Department' for current matter
  currCaseType = runSQL("SELECT Description FROM CaseTypes WHERE Code = (SELECT CaseTypeRef FROM Matters WHERE EntityRef = '{0}' AND Number = {1})".format(_tikitEntity, _tikitMatter))
  currDept = get_CurrentMatterDepartment()
  #currDept = runSQL("SELECT CTG.Name FROM CaseTypes CT JOIN CaseTypeGroups CTG ON CT.CaseTypeGroupRef = CTG.ID WHERE CT.Code = (SELECT CaseTypeRef FROM Matters WHERE EntityRef = '{0}' AND Number = {1})".format(_tikitEntity, _tikitMatter))

  # next, get count of defaults for 'CaseType' and for 'Department'
  countCaseTypeDefs = runSQL("SELECT COUNT(ID) FROM Partner.dbo.Usr_KeyDates_Defaults WHERE CaseType = '{0}'".format(currCaseType))
  #countDeptDefs = runSQL("SELECT COUNT(ID) FROM Partner.dbo.Usr_KeyDates_Defaults WHERE Department = '{0}'".format(currDept))
  tasksOrDates = runSQL("SELECT ISNULL(TypeToUse, 'Dates') FROM Usr_KeyDatesDeptSettings WHERE Department = (SELECT CaseTypeGroupRef FROM CaseTypes WHERE Code = (SELECT CaseTypeRef FROM Matters WHERE EntityRef = '{0}' AND Number = {1}))".format(_tikitEntity, _tikitMatter))

  # if count of CaseType defaults is zero
  if int(countCaseTypeDefs) == 0:
    # set SQL to use to point at the DEPARTMENT level ones
    #mySQL = "SELECT DisplayOrder, Description, LinkedMPField, NotifyDaysBefore FROM Partner.dbo.Usr_KeyDates_Defaults WHERE Department = '{0}' ORDER BY DisplayOrder".format(currDept)
    # Note that above does NOT take into account any already added and I'd like to show this split (see below)
    if tasksOrDates == 'Tasks':
      mySQL = """SELECT KDD.DisplayOrder, KDD.Description, KDD.LinkedMPField, KDD.NotifyDaysBefore,
                       'Group' = CASE WHEN (SELECT COUNT(Code) FROM Diary_Tasks DT WHERE DT.Description = KDD.Description AND EntityRef = '{0}' AND MatterNoRef = {1} AND DeleteThis = 0) > 0 THEN 'Already Added' ELSE 'Available' END
                 FROM Usr_KeyDates_Defaults KDD
                 WHERE KDD.Department = '{2}' 
                 GROUP BY KDD.DisplayOrder, KDD.Description, KDD.LinkedMPField, KDD.NotifyDaysBefore
                 ORDER BY KDD.DisplayOrder""".format(_tikitEntity, _tikitMatter, currDept)
    else:
      mySQL = """SELECT KDD.DisplayOrder, KDD.Description, KDD.LinkedMPField, KDD.NotifyDaysBefore,
                       'Group' = CASE WHEN (SELECT COUNT(Code) FROM Diary_Appointments DA WHERE DA.Description = KDD.Description AND EntityRef = '{0}' AND MatterNoRef = {1} AND DeleteThis = 0) > 0 THEN 'Already Added' ELSE 'Available' END
                 FROM Usr_KeyDates_Defaults KDD
                 WHERE KDD.Department = '{2}' 
                 GROUP BY KDD.DisplayOrder, KDD.Description, KDD.LinkedMPField, KDD.NotifyDaysBefore
                 ORDER BY KDD.DisplayOrder""".format(_tikitEntity, _tikitMatter, currDept)
    
  else:
    # there ARE some specified at CaseType, so set SQL to point to these
    mySQL = "SELECT DisplayOrder, Description, LinkedMPField, NotifyDaysBefore FROM Partner.dbo.Usr_KeyDates_Defaults WHERE CaseType = '{0}' ORDER BY DisplayOrder".format(currCaseType)
    # Note that above does NOT take into account any already added and I'd like to show this split (see below)
    if tasksOrDates == 'Tasks':
      mySQL = """SELECT KDD.DisplayOrder, KDD.Description, KDD.LinkedMPField, KDD.NotifyDaysBefore,
                       'Group' = CASE WHEN (SELECT COUNT(Code) FROM Diary_Tasks DT WHERE DT.Description = KDD.Description AND EntityRef = '{0}' AND MatterNoRef = {1} AND DeleteThis = 0) > 0 THEN 'Already Added' ELSE 'Available' END
                 FROM Usr_KeyDates_Defaults KDD
                 WHERE KDD.CaseType = '{2}' 
                 GROUP BY KDD.DisplayOrder, KDD.Description, KDD.LinkedMPField, KDD.NotifyDaysBefore
                 ORDER BY KDD.DisplayOrder""".format(_tikitEntity, _tikitMatter, currCaseType)
    else:
      mySQL = """SELECT KDD.DisplayOrder, KDD.Description, KDD.LinkedMPField, KDD.NotifyDaysBefore,
                       'Group' = CASE WHEN (SELECT COUNT(Code) FROM Diary_Appointments DA WHERE DA.Description = KDD.Description AND EntityRef = '{0}' AND MatterNoRef = {1} AND DeleteThis = 0) > 0 THEN 'Already Added' ELSE 'Available' END
                 FROM Usr_KeyDates_Defaults KDD
                 WHERE KDD.CaseType = '{2}' 
                 GROUP BY KDD.DisplayOrder, KDD.Description, KDD.LinkedMPField, KDD.NotifyDaysBefore
                 ORDER BY KDD.DisplayOrder""".format(_tikitEntity, _tikitMatter, currCaseType)

  myItems = []
  _tikitDbAccess.Open(mySQL)

  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          xTick = False
          xDOrder = 0 if dr.IsDBNull(0) else dr.GetValue(0)
          xDesc = '' if dr.IsDBNull(1) else dr.GetString(1)
          xLinkMPf = '' if dr.IsDBNull(2) else dr.GetString(2)
          xRemindDB = 0 if dr.IsDBNull(3) else dr.GetValue(3)
          xGroup = '' if dr.IsDBNull(4) else dr.GetString(4)

          myItems.append(KDDefaults(xTick, xDOrder, xDesc, xRemindDB, xLinkMPf, xGroup))

    dr.Close()
  _tikitDbAccess.Close()

  # put items into DataGrid
  tmpC = ListCollectionView(myItems)
  tmpC.GroupDescriptions.Add(PropertyGroupDescription("iGroup"))
  dg_DateDefaults.ItemsSource = tmpC

  # if there are no items in the DataGrid
  if dg_DateDefaults.Items.Count == 0:
    # hide the DataGrid and show the 'no items' help label
    lbl_NoDefaults.Visibility = Visibility.Visible
    dg_DateDefaults.Visibility = Visibility.Collapsed
  else:
    # there ARE items, so hide the 'no items' help label and show the DataGrid
    lbl_NoDefaults.Visibility = Visibility.Collapsed
    dg_DateDefaults.Visibility = Visibility.Visible
  return


def addDefaults_ToDiaryDates(s, event, tab = 0):
  # This function is the 'Add ticked items to Diary Dates' button on the 'Defaults' tab - NOTE: ADDS TO DIARY DATES TAB
  # Linked to XAML button.click: btn_AddDefaultsToDates
  # tab state determines whether it is being caled from key dates manager tab or default tab, 1 = key dates manager, 0 = defaults tab
  
  countAdded = 0
  countTicked = 0
  defaultDate = getDefaultNextDay()
  
  # lookup / get the Agenda ID and latest step ID from said agenda (needed for later)
  caseHistoryAgID = get_CaseHistoryAgendaID(titleForError = 'Error: addDefaults_ToTasks...')
  currentStepID = get_currentStepID(agendaID = caseHistoryAgID, titleForError = 'Error: addDefaults_ToTasks...')
  matterFE = runSQL("SELECT FeeEarnerRef FROM Matters WHERE EntityRef = '{0}' AND Number = {1}".format(_tikitEntity, _tikitMatter))

  # iterate over items in the Defaults list
  itemsToAdd = []
  for x in dg_DateDefaults.SelectedItems:
    # code to get unique version of 'Description'
    tmpDesc = getUniqueDescription(desiredDesc = x.iDesc, forUser = matterFE, taskORdate = 'Date')
    # technically, we'd get 'sql safe string' returned from above 'getUniqueDesc' as need to test 'safe' sql in that function
    #tmpDesc = sql_safe_string(stringToClean = tmpDesc)
    itemsToAdd.append(tmpDesc)
    countTicked += 1

  for y in itemsToAdd:
    # call global add DATE function to this row item
    global_AddDate(caseHistoryAgendaID = caseHistoryAgID, 
                         currentStepId = currentStepID, 
                       dateDescription = y, 
                           dateDueDate = defaultDate, 
                        dateAssignedTo = matterFE, 
                       dateDurationQty = 1, 
                      dateDurationType = 1,
                       dateReminderQty = 15,
                      dateReminderType = 0, 
                         isFromDefault = True)

    countAdded += 1

  # if count of items added matches count of ticked items
  if countAdded == countTicked:
    # successfully added all ticked items - alert user and ask if they wish to edit now
    myMsg = "Successfully added {0} ticked default Diary Date(s).\n\nWould you like to go to 'Diary Dates' tab now to update details?".format(countAdded)
  else:
    # advise user that only so many items were actually copied and ask if they wish to edit details now
    myMsg = "Only added {0} default Diary Dates out of the ticked {1} items due to an error.\n\nWould you like to go to 'Diary Dates' tab now to update details?".format(countAdded, countTicked)

  # refresh 'Defaults' list
  refresh_KDDefaults_List(s, event)
  if tab == 1:
    myResult = MessageBox.Show(myMsg, "Add Default Diary Dates...", MessageBoxButtons.YesNo)
    if myResult == DialogResult.Yes:
      refresh_KeyDates(s, event)
      tc_Main.SelectedIndex = 0
      if countTicked == 1:
        gCount = -1
        for g in dg_KeyDates.Items:
          gCount += 1
          if g.iDesc == tmpDesc:
            dg_KeyDates.SelectedIndex = gCount
            break

  refresh_KeyTasks(s, event)
  return


def addDefaults_ToTasks(s, event, tab = 1):
  # This will add the selected (ticked) default items as TASKS
  # Linked to XAML button.click: btn_AddDefaultsToTasks
  
  countAdded = 0
  countTicked = 0
  defaultDate = getDefaultNextDay()

  
  # lookup / get the Agenda ID and latest step ID from said agenda (needed for later)
  caseHistoryAgID = get_CaseHistoryAgendaID(titleForError = 'Error: addDefaults_ToTasks...')
  currentStepID = get_currentStepID(agendaID = caseHistoryAgID, titleForError = 'Error: addDefaults_ToTasks...')
  matterFE = runSQL("SELECT FeeEarnerRef FROM Matters WHERE EntityRef = '{0}' AND Number = {1}".format(_tikitEntity, _tikitMatter))

  # new - seems to only add one ticked item and doesn't report it didn't add the other(s)
  # wondering if better to add to list first, then loop over list instead perhaps
  itemsToAdd = []
  for x in dg_DateDefaults.SelectedItems:
    # if item is ticked
    #if x.iTicked == True:
    tmpDesc = getUniqueDescription(desiredDesc = x.iDesc, forUser = matterFE, taskORdate = 'Task')
    tmpDesc = sql_safe_string(stringToClean = tmpDesc)
    itemsToAdd.append(tmpDesc)
    countTicked += 1
    #MessageBox.Show("Desc: {0}\nCountTicked: {1}".format(tmpDesc, countTicked), "DEBUG")

  for y in itemsToAdd:
    # call global add task function to this row item
    global_AddTask(caseHistoryAgendaID = caseHistoryAgID, 
                         currentStepId = currentStepID, 
                       taskDescription = y, 
                           taskDueDate = defaultDate, 
                        taskAssignedTo = matterFE, 
                        taskRemindDate = defaultDate, 
                          taskPriority = 2, 
                       taskReminderQty = 15, 
                            taskStatus = 0, 
                        taskPCComplete = 0, 
                         isFromDefault = True)
    countAdded += 1

  # if count of items added matches count of ticked items
  if countAdded == countTicked:
    # successfully added all ticked items - alert user and ask if they wish to edit now
    myMsg = "Successfully added {0} ticked default Task Reminder(s).\n\nWould you like to go to 'Task Reminders' tab now to update details?".format(countAdded)
  else:
    # advise user that only so many items were actually copied and ask if they wish to edit details now
    myMsg = "Only added {0} default Task Reminders out of the ticked {1} items due to an error.\n\nWould you like to go to 'Task Reminders' tab now to update details?".format(countAdded, countTicked)

  # refresh 'Defaults' list
  refresh_KDDefaults_List(s, event)

  if tab == 1:
    myResult = MessageBox.Show(myMsg, "Add Default Key Dates...", MessageBoxButtons.YesNo)
    if myResult == DialogResult.Yes:
      refresh_KeyTasks(s, event)
      tc_Main.SelectedIndex = 1
      if countTicked == 1:
        gCount = -1
        for g in dg_KeyTasks.Items:
          gCount += 1
          if g.iDesc == tmpDesc:
            dg_KeyTasks.SelectedIndex = gCount
            break

  refresh_KeyTasks(s, event)
  return


def defaultDiaryDates_SelectAllNone(s, event):
  # This is the action button for the 'Tick All / None' button on the 'Defaults' tab, and will toggle between ticking all items and removing all ticks
  # Linked to XAML button.click: btn_DefaultsSelectAll

  selectAll_text = "Select All 'Available'"
  unselectAll_text = "De-select All"
  # firstly, deselect everything, as may have one item inadvertantly selected to start with
  dg_DateDefaults.SelectedItems.Clear()
  
  # firstly set the text/content of the button control to opposite state from current
  if tb_DefaultsSelectAll.Text == selectAll_text:
    #newTickStatus = True
    tb_DefaultsSelectAll.Text = unselectAll_text

    # new to work off selected item instead of creating new list
    for d in dg_DateDefaults.Items:
      if d.iGroup == 'Available':   # and newTickStatus == True:
        dg_DateDefaults.SelectedItems.Add(d)
    
  else:
    #newTickStatus = False
    tb_DefaultsSelectAll.Text = selectAll_text

  return


#################################################################################################################
# # # #   K E Y   D A T E S   -   S E C T I O N    # # # # 
class KeyDates(object):
  def __init__(self, myDACode, myDesc, myLocation, myDate, myTime, myDuration, myDurationType, myReminder, myReminderType, myDDAttendees, 
                     myDateCompleted, myDateMissedNotes, myCaseStepID, myLinkedMPField, myRowID, myAssignedTo, myGroup, myDurUnits, myRemindUnits, myFEList, myAgendaID):
    # break up Due Date into date and time
    tmpTime = str(myTime)
    tmpTime = tmpTime[:5]
    tmpDate = str(myDate)
    tmpDate = tmpDate[:10]

    # visible columns
    self.iDAcode = myDACode
    self.iDesc = myDesc
    self.iLocation = myLocation
    self.iDate = myDate
    self.iTime = tmpTime
    if len(tmpTime) > 3:
      self.iHour = tmpTime[:2]
      self.iMins = tmpTime[3:]
    else:
      self.iHour = "09"
      self.iMins = "00"

    self.iDurationFriendly = myDuration
    self.iReminderFriendly = myReminder
    self.iDDAttendees = myDDAttendees
    self.iDCompleted = myDateCompleted
    self.iDMissedNotes = myDateMissedNotes
    self.iCaseStepID = myCaseStepID
    
    # assign list items (combo boxes)
    self.iTypeItems = get_TypeOfUnitTypes()
    self.iHoursList = get_TimeHours(startHour=7, endHour=23)
    self.iMinsList = get_TimeMins(increment=5)
  
    # hidden columns (for use by other functions)
    self.iDateAndTime = "{0} {1}".format(tmpDate, tmpTime)
    self.iDuration = myDurUnits
    self.iDurationType = myDurationType
    self.iReminder = myRemindUnits
    self.iReminderType = myReminderType
    self.iLinkedMPField = myLinkedMPField
    self.iRowID = myRowID
    self.iAssignedTo = myAssignedTo
    self.iGroup = myGroup
    self.FEList = myFEList
    self.iAgenda = myAgendaID
    return

  def __getitem__(self, index):
    if index == 'DAcode':
      return self.iDAcode
    elif index == 'Desc': 
      return self.iDesc
    elif index == 'Date': 
      return self.iDate
    elif index == 'Location': 
      if self.iLocation is None:
        return ''
      else:
        return self.iLocation

    elif index == 'Time':
      if self.iHour is None or self.iHour == '00':
        tmpHour = '09'
      else:
        tmpHour = self.iHour

      if self.iMins is None or self.iMins == '00':
        tmpMins = '00'
      else:
        tmpMins = self.iMins
      
      return "{0}:{1}".format(tmpHour, tmpMins)
      # formally the below when allowing for straight text input
      #return self.iTime
    elif index == 'oldDueTime':
      return self.iTime
    elif index == 'Duration':
      if self.iDuration is None:
        return 1
      else:
        return self.iDuration

    elif index == 'DurType':
      if self.iDurationType is None:
        return 1
      else:
        return self.iDurationType

    elif index == 'Reminder':
      if self.iReminder is None:
        return 15
      else:
        return self.iReminder

    elif index == 'RemType':
      if self.iReminderType is None:
        return 0
      else:
        return self.iReminderType

    elif index == 'Attendees':
      if self.iDDAttendees is None:
        return ''
      #elif len(str(self.iDDAttendees)) > 0:
      else:
        return self.iDDAttendees

    elif index == 'DateCompleted': 
      return self.iDCompleted
    elif index == 'DateMissedNotes': 
      if self.iDMissedNotes is None:
        return ''
      else:
        return self.iDMissedNotes

    elif index == 'CaseStepID': 
      return self.iCaseStepID
    elif index == 'LinkedMPField':
      return self.iLinkedMPField
    elif index == 'RowID': 
      return self.iRowID
    elif index == 'Agenda':
      return self.iAgenda
    elif index == 'AssignedTo':
      return self.iAssignedTo
    elif index == 'Grouping':
      return self.iGroup

def refresh_KeyDates(s, event):
  # This function will refresh the 'Diary Dates' DataGrid
  
  # need to pass in list of Fee Earners (cannot run within '__init__' because we end up closing the already open data reader)
  tmpFEs = get_FeeEarnerList()
  
  # we need to get the 'Case History' agenda ID for SQL (unless other departments state they want all, in which case we need to amend this (or add options to 'Dept defaults')
  caseHistoryAgID = get_CaseHistoryAgendaID(titleForError = 'Error: refresh_KeyDates...')

  # New SQL
  mySQL = """SELECT '0-Desc' = CI.Description, '1-Location' = DA.Location, '2-Due Date' = CMS.StartDate, 
                    '3-DurationType' = CASE CMS.DurationType WHEN 0 THEN 'Minute(s)' WHEN 1 THEN 'Hour(s)' WHEN 2 THEN 'Day(s)' WHEN 3 THEN 'Week(s)' END, 
                    '4-ReminderType' = CASE CMS.ReminderType WHEN 0 THEN 'Minute(s)' WHEN 1 THEN 'Hour(s)' WHEN 2 THEN 'Day(s)' WHEN 3 THEN 'Week(s)' END,  
                    '5-Attendees' = (SELECT STRING_AGG(Fullname, '; ') FROM (SELECT 'Fullname' = UA.Fullname FROM Diary_AppointmentsAttendees DAA JOIN Users UA ON DAA.Username = UA.Code WHERE DAA.AppointmentRef = DA.Code) as tmpT), 
                    '6-CompleteDate' = ISNULL(CONVERT(NVARCHAR, CI.CompletionDate, 103), ''), 
                    '7-DateMissedNotes' = ISNULL(KD.DateMissedNotes, ''), 
                    '8-Assigned To' = ISNULL(DA.Username, ISNULL(CMS.AssignedUser, KD.AssignedTo)), 
                    'H9-Code' = ISNULL(DA.Code,0), 'H10-AgendaRef' = CI.ParentID, 'H11-CaseItemRef' = CI.ItemID, 
                    'H12-Group' = CASE WHEN CMS.StepCategory = 'KeyDD' THEN '0) Added from Defaults - To set date' ELSE (CASE WHEN CI.CompletionDate IS NULL THEN '1) Outstanding' ELSE '2) Complete' END) END, 
                    'x13-DateTime' = CONVERT(varchar(10), CMS.StartDate, 8), 'x14-DurationType' = CMS.DurationType, 'x15-RemindType' = CMS.ReminderType, 'x16-LinkedMP' = KD.LinkedMPField, 
                    'x17-DurationUnits' = CMS.Duration, 'x18-ReminderUnits' = CMS.Reminder, 'x19-KDid' = ISNULL(KD.ID, 0) 
            FROM Cm_CaseItems CI 
              LEFT OUTER JOIN Cm_Steps CMS ON CI.ItemID = CMS.ItemID 
              LEFT OUTER JOIN Diary_Appointments DA ON CI.ItemID = DA.CaseItemRef 
              LEFT OUTER JOIN Usr_Key_Dates KD ON CI.ItemID = KD.TaskStepID
            WHERE CI.ParentID = {0}
              AND CMS.Type = 'Date' AND CMS.StepCategory LIKE 'KeyD%'
            ORDER BY [H12-Group] ASC, CMS.StartDate DESC, CI.Description""".format(caseHistoryAgID)

  #MessageBox.Show(mySQL)
  
  myItems = []
  _tikitDbAccess.Open(mySQL)

  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          iDesc = '-' if dr.IsDBNull(0) else dr.GetString(0)
          iLoc = '' if dr.IsDBNull(1) else dr.GetString(1)
          iDate = '' if dr.IsDBNull(2) else dr.GetValue(2)                    # 2024-08-27 16:00:00.000
          iDuration = '' if dr.IsDBNull(3) else dr.GetString(3)               # 1 hour(s)
          iReminder = '' if dr.IsDBNull(4) else dr.GetString(4)               # 15 minute(s)
          iDDAttendees = '' if dr.IsDBNull(5) else dr.GetString(5)            # MP: Matt Patt; LD1: Louis Debnam
          iDateCompleted = '-' if dr.IsDBNull(6) else dr.GetValue(6)          # 2024-08-27 18:00:00.000
          iDateMissedNotes = '' if dr.IsDBNull(7) else dr.GetString(7)        # text notes
          iAssignedTo = '' if dr.IsDBNull(8) else dr.GetString(8)             # MP        #: Matt Patt
          iDAcode = 0 if dr.IsDBNull(9) else dr.GetValue(9)                   # 965
          iCaseStepID = 0 if dr.IsDBNull(11) else dr.GetValue(11)             # 12345678
          iGrouping = '' if dr.IsDBNull(12) else dr.GetString(12)             # Completed | Outstanding
          iTime = '' if dr.IsDBNull(13) else dr.GetString(13)                 # 16:00:00
          iDurationType = '' if dr.IsDBNull(14) else dr.GetValue(14)          # 0 | 1 | 2 | 3
          iReminderType = '' if dr.IsDBNull(15) else dr.GetValue(15)          # 0 | 1 | 2 | 3
          iLinkedMPField = '' if dr.IsDBNull(16) else dr.GetString(16)        # [myTable.myField]
          iDurationUnits = 0 if dr.IsDBNull(17) else dr.GetValue(17)
          iRemindUnits = 0 if dr.IsDBNull(18) else dr.GetValue(18)
          iKDID = 0 if dr.IsDBNull(19) else dr.GetValue(19)                   # Key Dates table ID

          myItems.append(KeyDates(myDACode = iDAcode, myDesc = iDesc, myLocation = iLoc, myDate = iDate, myDuration = iDuration, myReminder = iReminder, 
                                  myDDAttendees = iDDAttendees, myDateCompleted = iDateCompleted, myDateMissedNotes = iDateMissedNotes, myCaseStepID = iCaseStepID, 
                                  myAssignedTo = iAssignedTo, myTime = iTime, myDurationType = iDurationType, myReminderType = iReminderType, myLinkedMPField = iLinkedMPField, 
                                  myRowID = iKDID, myGroup = iGrouping, myDurUnits = iDurationUnits, myRemindUnits = iRemindUnits, myFEList = tmpFEs, myAgendaID = caseHistoryAgID))
    dr.Close()
  _tikitDbAccess.Close()

  # add grouping
  tmpC = ListCollectionView(myItems)
  tmpC.GroupDescriptions.Add(PropertyGroupDescription("iGroup"))
  dg_KeyDates.ItemsSource = tmpC
  return


def dg_KeyDates_SelectionChanged(s, event):
  # This function runs whenever the selection is changed in the Key Dates (Diary Dates) DataGrid
  # Linked to XAML control: dg_KeyDates.SelectionChanged
  appendToDebugLog(textToAppend="entering dg_KeyDates_SelectionChanged() event", inclTimeStamp=True)

  if dg_KeyDates.Items.Count == 0:
    appendToDebugLog(textToAppend = 'No items in DataGrid', inclTimeStamp = True)
    return

  # if nothing is selected
  if dg_KeyDates.SelectedIndex == -1:
    # set all values in bottom 'edit area' of XAML to default values
    appendToDebugLog(textToAppend = 'Nothing selected - setting values to null', inclTimeStamp = True)
    #btnTB_DD_Save.Text = 'Add New'
    cbo_Date_AssignedTo.SelectedIndex = -1
    txt_DD_Desc.Text = ''
    txt_DD_Location.Text = ''
    dp_DD_Date.SelectedDate = datetime.now()  #DateTime.Now  #_tikitResolver.Resolve("[SQL: SELECT GETDATE()]")
    txt_DD_Time.Text = '09:00'
    txt_DD_Duration.Text = '1'
    cbo_DD_DurType.SelectedIndex = 1
    txt_DD_ReminderQty.Text = '15'
    cbo_DD_RemType.SelectedIndex = 0
    lbl_CaseStepID.Content = ''
    lbl_DateMissedNotes.Content = ''
    lbl_DAcode.Content = ''
    txt_DateMissedNotes.Text = ''
    refresh_AttendeeList()
    
    # disable buttons that act on a selected item
    btn_MarkAsComplete_Date.IsEnabled = False
    btn_RevertDate.IsEnabled = False
    btn_DeleteDate.IsEnabled = False
    
  else:
    # something valid was selected so populate controls on bottom half of form (with selected DG values)
    tmpGroup = dg_KeyDates.SelectedItem['Grouping']

    # if item is complete, show Date Missed Notes
    if tmpGroup == '2) Complete':
      #stk_TaskMissedNotes.Visibility = Visibility.Visible
      btn_MarkAsComplete_Date.IsEnabled = False
      btn_RevertDate.IsEnabled = True
      btn_DeleteDate.IsEnabled = False
    else:
      #stk_TaskMissedNotes.Visibility = Visibility.Collapsed
      btn_MarkAsComplete_Date.IsEnabled = True
      btn_RevertDate.IsEnabled = False
      btn_DeleteDate.IsEnabled = True

      #btnTB_DD_Save.Text = 'Save'
      opt_DD_EditSelected.IsChecked = True
      lbl_DAcode.Content = dg_KeyDates.SelectedItem['DAcode']
      lbl_CaseStepID.Content = dg_KeyDates.SelectedItem['CaseStepID']
      #MessageBox.Show("Updating Label items, DA Code ({0}) and CaseStepID ({1})".format(lbl_DAcode.Content, lbl_CaseStepID.Content), "DEBUGGING")

      txt_DD_Desc.Text = str(dg_KeyDates.SelectedItem['Desc'])
      txt_DD_Location.Text = str(dg_KeyDates.SelectedItem['Location'])
      #MessageBox.Show("Updating Text boxes, Desc ({0}) and Location ({1})".format(txt_DD_Desc.Text, txt_DD_Location.Text), "DEBUGGING")
      txt_DD_Time.Text = str(dg_KeyDates.SelectedItem['Time'])
      txt_DD_Duration.Text = str(dg_KeyDates.SelectedItem['Duration'])
      #MessageBox.Show("Updating Text boxes, Time ({0}) and Duration ({1})".format(txt_DD_Time.Text, txt_DD_Duration.Text), "DEBUGGING")
      txt_DateMissedNotes.Text = str(dg_KeyDates.SelectedItem['DateMissedNotes'])
      #lbl_DateMissedNotes.Content = str(dg_KeyDates.SelectedItem['DateMissedNotes'])
      txt_DD_ReminderQty.Text = str(dg_KeyDates.SelectedItem['Reminder'])
      #MessageBox.Show("Updating Text boxes, Date Missed Notes ({0}) and Reminder Qty ({1})".format(txt_DateMissedNotes.Text, txt_DD_ReminderQty.Text), "DEBUGGING")
      tmpAttendees = dg_KeyDates.SelectedItem['Attendees']
      #MessageBox.Show("Updating Text boxes, Attendees ({0})".format(tmpAttendees), "DEBUGGING")

      tmpDate = dg_KeyDates.SelectedItem['Date']
      if tmpDate is None or tmpDate == '':
        dp_DD_Date.SelectedDate = None
      else: 
        dp_DD_Date.SelectedDate = tmpDate
      #MessageBox.Show("Updating DatePicker, Date ({0})".format(dp_DD_Date.SelectedDate), "DEBUGGING")

      # iterate over combo box items looking for matching value and exit loop once found
      tmpDurType = dg_KeyDates.SelectedItem['DurType']
      dCount = -1
      #isSet = 'No'
      for xItem in cbo_DD_DurType.Items:
        dCount += 1
        if xItem.iCode == tmpDurType:
          cbo_DD_DurType.SelectedIndex = dCount
          #isSet = 'Yes'
          break
      #MessageBox.Show("Updating Combo box, Duration Type ({0}), is set: {1}".format(tmpDurType, isSet), "DEBUGGING")

      tmpRemType = dg_KeyDates.SelectedItem['RemType']
      rCount = -1
      #isSet = 'No'
      for xItem in cbo_DD_RemType.Items:
        rCount += 1
        if xItem.iCode == tmpRemType:
          cbo_DD_RemType.SelectedIndex = rCount
          #isSet = 'Yes'
          break
      #MessageBox.Show("Updating Combo box, Reminder Type ({0}), is set: {1}".format(tmpRemType, isSet), "DEBUGGING")

      #posOfC = tmpUserCode.find(":")
      #tmpUserCode = tmpUserCode[:posOfC]
      ##MessageBox.Show("Current User code: {0}".format(tmpUserCode), "DEBUGGING")
      tmpUserCode = dg_KeyDates.SelectedItem['AssignedTo']
      if tmpUserCode is not None and str(tmpUserCode) != '':
        uCount = -1
        #isSet = 'No'
        for userItem in cbo_Date_AssignedTo.Items:
          uCount += 1
          if userItem.iCode == tmpUserCode:
            cbo_Date_AssignedTo.SelectedIndex = uCount
            #isSet = 'Yes'
            break
        #MessageBox.Show("Updating Combo box, Assigned To ({0}), is set: {1}".format(tmpUserCode, isSet), "DEBUGGING")

      if tmpAttendees is not None and len(tmpAttendees) > 0:
        refresh_AttendeeList(tmpAttendees)
        #MessageBox.Show("Updating Attendees list", "DEBUGGING")
      #MessageBox.Show("End of dg_KeyDate_SelectionChanged", "DEBUGGING")
  return


class comboTypes(object):
  def __init__(self, myText, myCode, myLtr):
    self.iCode = myCode
    self.iText = myText
    self.iLtr = myLtr
    return

  def __getitem__(self, index):
    if index == 'Text':
      return self.iText
    elif index == 'Code':
      return self.iCode
    elif index == 'Letter':
      return self.iLtr

def populateComboTypes(s, event):
  cbo_DD_DurType.ItemsSource = get_TypeOfUnitTypes()
  cbo_DD_RemType.ItemsSource = get_TypeOfUnitTypes()
  return

def get_TypeOfUnitTypes():
  xItem = []
  xItem.append(comboTypes('Minute(s)', 0, 'M'))
  xItem.append(comboTypes('Hour(s)', 1, 'H'))
  xItem.append(comboTypes('Day(s)', 2, 'D'))
  xItem.append(comboTypes('Week(s)', 3, 'W'))
  return xItem


def expand_DateAttendees(s, event):
  #grd_MainDates.ColumnDefinitions[0].Width = System.Windows.GridLength(700)
  col_DatesDG.Width = GridLength(700, GridUnitType.Pixel)
  return

def contract_DateAttendees(s, event):
  #grd_MainDates.ColumnDefinitions[0].Width = System.Windows.GridLength(1040)
  exp_DateAttendees.IsExpanded = False if exp_DateAttendees.IsExpanded == True else False
  col_DatesDG.Width = GridLength(1040, GridUnitType.Pixel)
  return


class attendees(object):
  def __init__(self, myTick, myCode, myName, myJTitle):
    self.iTemTicked = myTick
    self.iCode = myCode
    self.iName = myName
    self.iJobTitle = myJTitle
    return

  def __getitem__(self, index):
    if index == 'Ticked':
      return self.iTemTicked
    elif index == 'Code':
      return self.iCode
    elif index == 'Name':
      return self.iName
    elif index == 'JobTitle':
      return self.iJobTitle    

def refresh_AttendeeList(usersToPreSelect = ''):
  # This function takes the passed 'usersToPreSelect' and starts creating new 'Attendees' list with these people pre-ticked
  # and then proceeds to add all other users in the firm (or department, if checkbox ticked).

  myItems = []
  mySplitUsers = []
  fullNameInString = ''
  
  # (adding actually ticked people first)
  # if length of passed argument is not zero
  if len(usersToPreSelect) != 0:
  
    # split passed argument into a list we can iterate over
    mySplitUsers = usersToPreSelect.split("; ")

    # if more than zero list items
    if len(mySplitUsers) > 0:
      # iterate over list
      for y in mySplitUsers:
        # enclose name in apostrophe (for use later) and add comma
        fullNameInString += "'{0}', ".format(y.replace("'", "''"))

      # once above loop has finished we will have an extra two unwanted characters (comma and space), so lets get rid of these
      strLength = len(fullNameInString) - 2
      fullNameInString = fullNameInString[:strLength]

      # now form SQL to get user data based on these users (using string we created above)
      attList_SQL = "SELECT Code, FullName, ISNULL(JobTitle, '') FROM Users U WHERE U.UserStatus = 0 AND Locked = 0 AND LEN(MailBoxName) > 0 AND FullName IN({0})".format(fullNameInString)

      _tikitDbAccess.Open(attList_SQL)

      if _tikitDbAccess._dr is not None:
        dr = _tikitDbAccess._dr
        if dr.HasRows:
          while dr.Read():
            if not dr.IsDBNull(0):
              iCode = '' if dr.IsDBNull(0) else dr.GetString(0)
              iName = '' if dr.IsDBNull(1) else dr.GetString(1)
              iTitle = '' if dr.IsDBNull(2) else dr.GetString(2)
              iTicked = True

              myItems.append(attendees(iTicked, iCode, iName, iTitle))

        dr.Close()
      _tikitDbAccess.Close()


  # (now add all other users)
  # if 'only show dept users' is ticked
  if chk_DD_OnlyShowDeptUsers.IsChecked == True:
  
    # get current department for current matter
    currentDeptCode = get_CurrentMatterDepartment()
  
    if len(currentDeptCode) > 0:
      # found a department for current matter - form SQL to include only Department members
      userList_SQL = "SELECT Code, FullName, ISNULL(JobTitle, '') FROM Users U WHERE U.UserStatus = 0 AND Locked = 0 AND LEN(MailBoxName) > 0 AND Department = '{0}' ".format(currentDeptCode)
    else:
      # no department found for current matter - form SQL to include everyone
      userList_SQL = "SELECT Code, FullName, ISNULL(JobTitle, '') FROM Users U WHERE U.UserStatus = 0 AND Locked = 0 AND LEN(MailBoxName) > 0 "
  else:
    # dept only users NOT ticked - form SQL to include everyone
    userList_SQL = "SELECT Code, FullName, ISNULL(JobTitle, '') FROM Users U WHERE U.UserStatus = 0 AND Locked = 0 AND LEN(MailBoxName) > 0 "

  # if our string we created before has more than one character
  if len(fullNameInString) != 0:
    # include in SQL to EXCLUDE those users already added to list
    userList_SQL += " AND FullName NOT IN({0}) ".format(fullNameInString)

  # allowing for Text search
  tmpSearchText = txt_AttendeeSearch.Text
  if tmpSearchText is not None:
    if len(tmpSearchText) > 0:
      userList_SQL += " AND FullName LIKE '%{0}%'".format(tmpSearchText)

  _tikitDbAccess.Open(userList_SQL)

  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          iCode = '' if dr.IsDBNull(0) else dr.GetString(0)
          iName = '' if dr.IsDBNull(1) else dr.GetString(1)
          iTitle = '' if dr.IsDBNull(2) else dr.GetString(2)
          iTicked = False

          myItems.append(attendees(iTicked, iCode, iName, iTitle))

    dr.Close()
  _tikitDbAccess.Close()

  # finally, populate attendee datagrid with our list of users we've created
  dg_DD_Attendees.ItemsSource = myItems
  return


def populateAttendeesList(s, event):
  if dg_KeyDates.SelectedIndex > -1:
    refresh_AttendeeList(dg_KeyDates.SelectedItem['Attendees'])
  else:
    refresh_AttendeeList()
  return


# # # #   M A I N   A D D  /  U P D A T E   C O D E   F O R   D A T E S   # # # #

def global_AddDate(caseHistoryAgendaID, currentStepId, dateDescription, dateDueDate, dateAssignedTo, dateDurationQty = 1, dateDurationType = 1, 
                    dateReminderQty = 15, dateReminderType = 0, dateLocation = '', isFromDefault = False):
  # This function will add a new TASK item with the passed details

  # need to first separate out time as 'sp_InsertStepMP' only wants a date - not time element as well
  # just get the left 10 characters from due date
  spDueDate = dateDueDate[:10]

  # run the in-built stored procedure to add step
    # @AgendaID INT, @InsertWhere VARCHAR(1), @CurrentStepID INT, @Description VARCHAR(260), @Mandatory VARCHAR(1), @DocID INT,
    # @DocType VARCHAR(20), @StartDate VARCHAR(20), @Duration INT, @DurationType VARCHAR(1), @ReminderUnitQty INT, @ReminderTypeOfUnit VARCHAR(1), @Username VARCHAR(12)
  sqlToRun = "EXEC sp_InsertStepMP {0}, 'L', {1}, '{2}', 'N', 0, 'Date', '{3}', {4}, '{5}', {6}, '{7}', '{8}'".format(caseHistoryAgendaID, currentStepId, dateDescription, spDueDate, 
                                                                                                                      dateDurationQty, dateDurationType, dateReminderQty, dateReminderType, dateAssignedTo)

  #msg = MessageBox.Show("SQL to be run:\n" + sqlToRun + "\n\nOK to continue?", "Adding Date...", MessageBoxButtons.YesNo)
  #if msg == DialogResult.No:
  #  return
  runSQL(sqlToRun, True, "There was an error using InsertStep to add Task", "Error: global_AddDate") 

  # unfortunately, the InsertStep stored procedure does not add the 'reminder date' for tasks (into Diary_Appointments). Which means we need to manually add it
  # firstly, get the ID from the Diary_Tasks table
  tmpCountDT = runSQL("SELECT COUNT(Code) FROM Diary_Appointments WHERE EntityRef = '{0}' AND MatterNoRef = {1} AND Username = '{2}' AND Description = '{3}'".format(_tikitEntity, _tikitMatter, 
                                                                                                                                                                     dateAssignedTo, dateDescription))
  if int(tmpCountDT) > 0:
    # now update it with the actual reminder date, and set UpdateThis to 1 to make sure it feeds through to FE's Outlook
    dtCode = runSQL("SELECT TOP(1) Code FROM Diary_Appointments WHERE EntityRef = '{0}' AND MatterNoRef = {1} AND Username = '{2}' AND Description = '{3}'".format(_tikitEntity, _tikitMatter, 
                                                                                                                                                                   dateAssignedTo, dateDescription))
    #MessageBox.Show("Count of items matching details: {0}\ndtCode: {1}".format(tmpCountDT, dtCode), "DEBUGGING: Updating Location")
    
    if int(dtCode) > 0:
      sqlToRun = "UPDATE Diary_Appointments SET DateStamp = '{3}', DurationType = 1, Location = '{0}', UpdateThis = {1} WHERE Code = {2}".format(dateLocation, gUpdateThis, dtCode, dateDueDate)
      #MessageBox.Show("SQL to update:\n{0}".format(sqlToRun), "DEBUGGING: Updating Location")
      runSQL(sqlToRun)

      if isFromDefault == True:
        stepCatID = 'KeyDD'
      else:
        stepCatID = 'KeyD'
        
      # additionally, it doesn't add our 'KeyDate' category to the Cm_Steps table, so we need to add this too
      # get the Case Item ID from the Dairy_ table
      caseID = runSQL("SELECT CaseItemRef FROM Diary_Appointments WHERE Code = {0}".format(dtCode))
      if int(caseID) > 0:
        runSQL("UPDATE Cm_Steps SET StartDate = '{2}', DurationType = 1, StepCategory = '{0}' WHERE ItemID = {1}".format(stepCatID, caseID, dateDueDate))

  return

def date_AddAllDefaults(s, event):
  # Function tied to the btn_AddAllDefaultTasks on the XAML, designed to add all of the case type defaults to the key dates manager tab
  # This will add all available (not already added based on the "description" field of the task)
  user_confirmed = MessageBox.Show("Are you sure you want to add all Default dates?", "Add All Defaults?", MessageBoxButtons.YesNo)
    
  # Process only if the user confirms
  if user_confirmed == DialogResult.Yes:
    for d in dg_DateDefaults.Items:
      if d.iGroup == 'Available':   # and newTickStatus == True:
        dg_DateDefaults.SelectedItems.Add(d)
    addDefaults_ToDiaryDates(s, event, 1)



def date_AddNew(s, event):
  # This is the new function to add a new blank Date to the list (NB: we'll give it the 'Default' Category too, so it appears near top)
  # Linked to XAML button.click for: btn_AddNew_Date
  # 'Types': 0 = Minutes; 1 = Hours; 2 = Days; 3 = Weeks

  # set defaults
  caseHistoryAgID = get_CaseHistoryAgendaID(titleForError = 'Error: date_AddNewOrSave - getting Case History agenda ID...')
  currentStepID = get_currentStepID(agendaID = caseHistoryAgID, titleForError = 'Error: date_AddNewOrSave - getting Current Step ID...')
  dueDate = getDefaultNextDay()
  assignedTo = runSQL("SELECT FeeEarnerRef FROM Matters WHERE EntityRef = '{0}' AND Number = {1}".format(_tikitEntity, _tikitMatter))
  descToUse = getUniqueDescription(desiredDesc = 'New Date - Edit here', forUser = assignedTo, taskORdate = 'Date')
  
  # we created a generic function to add dates, so call that passing in values
  global_AddDate(caseHistoryAgendaID = caseHistoryAgID, 
                       currentStepId = currentStepID, 
                     dateDescription = descToUse, 
                         dateDueDate = dueDate, 
                      dateAssignedTo = assignedTo, 
                     dateDurationQty = 1, 
                    dateDurationType = 1,
                     dateReminderQty = 15,
                    dateReminderType = 0,
                        dateLocation = '', 
                       isFromDefault = True)
  
  # refresh 'Dates' datagrid
  refresh_KeyDates(s, event)

  # ought to pre-select item too
  tmpX = -1
  for xRow in dg_KeyDates.Items:
    tmpX += 1
    if xRow.iDesc == descToUse: # and xRow.iDate == dueDate:
      dg_KeyDates.SelectedIndex = tmpX
      break
  return
  

def getDefaultNextDay():
  canExit = False
  tmpIncr = 1
  tmpDayName = ''

  # keep looping next part until we set 'canExit' to True
  while canExit == False:
    # get the next date
    tmpDate = getSQLDate(runSQL("SELECT DATEADD(day, {0}, GETDATE())".format(tmpIncr)))
    #MessageBox.Show("tmpDate: {0}".format(tmpDate), "DEBUG: getDefaultNextDay")
    
    # add time element
    tmpDate1 = "{0} 09:00:00.000".format(tmpDate)
    #MessageBox.Show("tmpDate (in SQL format and with new 9am time): {0}".format(tmpDate1), "DEBUG: getDefaultNextDay")
    
    # get day name
    tmpDayName = runSQL("SELECT DATENAME(dw, '{0}')".format(tmpDate1))
    #MessageBox.Show("tmpDayName: {0}".format(tmpDayName), "DEBUG: getDefaultNextDay")
    
    # if day is a weekday
    if tmpDayName in ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]:
      # we can stop function and return this date
      canExit = True
      break
    else:
      # increment counter and re-run loop
      tmpIncr += 1
  
  #MessageBox.Show("End of function - return value is: {0}".format(tmpDate), "DEBUG: getDefaultNextDay")
  return str(tmpDate1)

def getTimeFixed(inputToCheck):
  # This function takes one parameter and will return a clean 'time'
  if inputToCheck == None:
    return "09:00:00.000"

  # firtly, try casting passed item as as string
  tmpStr = str(inputToCheck)
  # next, remove any leading or trailing spaces
  tmpStr = tmpStr.strip()
  # as a failsafe, replace any period/full stop with time colon
  tmpStr = tmpStr.replace(".", ":")
  # the following function will remove any characters that are not 'time' related
  tmpStr = stripString(sourceString = tmpStr, leaveOnly = '1234567890:')
  newTime = ""
  tmpCount = 0

  # break up string to only get the first part
  mySplit = tmpStr.split(":")

  # this next loop ensures we get a leading zero (in case user entered 9:00 for example)
  for x in mySplit:
    tmpCount += 1
    tmpNo = int(x)
    # new check - first part (hours)
    if tmpCount == 1:
      # if the 'hours' part is greater than 24
      if tmpNo > 24:
        # set to '09' (default 9 o'clock)
        newTime += "09"
      else: 
        # number is 24 or below, use as is
        if len(str(tmpNo)) == 1:
          newTime += "0{0}".format(tmpNo)
        else:
          newTime += str(tmpNo)
    else:
      # new check for 'minutes' part (cannot be greater than 59)
      if tmpNo > 59:
        # number is greater than 59, so set to default '00'
        newTime += "00"
      else:
        # number is NOT greater than 59, so use number as is
        if len(str(tmpNo)) == 1:
          newTime += "0{0}".format(tmpNo)
        else:
          newTime += str(tmpNo)
    newTime += ":"

  lenNewTime = len(str(newTime)) 

  # This last part just add the 'seconds' end part
  if lenNewTime == 3:
    newTime += "00:00"
  elif lenNewTime == 6:
    newTime += "00"
  elif lenNewTime > 8:
    newTime = newTime[:8]
  
  return newTime


class AssignToList(object):
  def __init__(self, myFECode, myFEName):
    self.iCode = myFECode
    self.iName = myFEName
    return

  def __getitem__(self, index):
    if index == 'Code':
      return self.iCode
    elif index == 'Name':
      return self.iName

def get_FeeEarnerList():
  mySQL = "SELECT Code, FullName FROM Users WHERE Locked = 0 AND UserStatus = 0 ORDER BY FullName"  #FeeEarner = 1 AND

  _tikitDbAccess.Open(mySQL)
  myFEitems = []

  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          myCode = '-' if dr.IsDBNull(0) else dr.GetString(0)
          myName = '-' if dr.IsDBNull(1) else dr.GetString(1)

          myFEitems.append(AssignToList(myCode, myName))  

    dr.Close()
  _tikitDbAccess.Close()
  
  return myFEitems


def populate_FeeEarnersList(s, event): 

  tmpList = get_FeeEarnerList()
  cbo_Date_AssignedTo.ItemsSource = tmpList
  cbo_Task_AssignedTo.ItemsSource = tmpList
  return

def assignDate_toMatterFeeEarner(s, event):

  matterFE = _tikitResolver.Resolve("[SQL: SELECT FeeEarnerRef FROM Matters WHERE EntityRef = '{0}' AND Number = {1}]".format(_tikitEntity, _tikitMatter))
  tmpCount = -1
  
  if len(matterFE) == 0:
    MessageBox.Show("There doesn't appear to be a Fee Earner set against this matter - please check the Matter Properties screen")
    return
    
  for i in cbo_Date_AssignedTo.Items:
    tmpCount += 1
    if i.iCode == matterFE:
      cbo_Date_AssignedTo.SelectedIndex = tmpCount
      break
  return


def addignDate_toCurrentUser(s, event):
  tmpCount = -1
  
  for i in cbo_Date_AssignedTo.Items:
    tmpCount += 1
    if i.iCode == _tikitUser:
      cbo_Date_AssignedTo.SelectedIndex = tmpCount
      break  
  return


def dd_EditSelected_Click(s, event):
  opt_DD_EditSelected.FontWeight = FontWeights.Bold
  opt_DD_AddNew.FontWeight = FontWeights.Normal
  #btnTB_DD_Save.Text = 'Save'
  return


def dd_AddNew_Click(s, event):
  opt_DD_AddNew.FontWeight = FontWeights.Bold
  opt_DD_EditSelected.FontWeight = FontWeights.Normal
  # set to nothing selected - should then invoke 'dg_KeyDates_SelectionChanged()'
  dg_KeyDates.SelectedIndex = -1
  refresh_AttendeeList()
  #btnTB_DD_Save.Text = 'Add New'
  return

def get_TypeNo(fromText):
  if fromText == 'M':
    return '0'  
  elif fromText == 'H':
    return '1'
  elif fromText == 'D':
    return '2'
  elif fromText == 'W':
    return '3'


def updateMPLinkedField(fullMergeCode, newDate):
  # This function will break up the passed 'fullMergeCode' into the Table element and the Field element, where we then form SQL to update required field with the 'newDate'
  #MessageBox.Show("fullMergeCode: " + fullMergeCode, "Function: updateMPLinkedField")

  if len(fullMergeCode) == 0:
    return

  tmpGotToDot = False
  myTable = ''
  myField = ''

  for x in fullMergeCode:
    if x == '.':
      tmpGotToDot = True
    else:
      if tmpGotToDot == False:
        myTable += x
      else:
        myField += x
      
  finalTbl = myTable.replace('[', '')
  finalFld = myField.replace(']', '')
  #MessageBox.Show("finalTbl: " + finalTbl + "\nfinalFld: " + finalFld, "Function: updateMPLinkedField")

  if len(finalTbl) > 0 and len(finalFld) > 0:
    # NB we return a date in year 3000 if date is currently null, so can test against that before updating date here
    #myGetSQL = _tikitResolver.Resolve("[SQL: SELECT ISNULL(" + finalFld + ", '3000-01-01') FROM " + finalTbl + " WHERE EntityRef = '" + _tikitEntity + "' AND MatterNo = " + str(_tikitMatter) + "]")
    #if str(myGetSQL) != '3000-01-01 00:00:00.000':
    # update MP Linked  table 
    updateSQL = "[SQL: UPDATE {0} SET {1} = '{2}' WHERE EntityRef = '{3}' AND MatterNo = {4}]".format(finalTbl, finalFld, newDate, _tikitEntity, _tikitMatter)
    #MessageBox.Show("updateSQL: " + updateSQL, "Function: updateMPLinkedField")
    _tikitResolver.Resolve(updateSQL)

    # update Key Dates table
    #updateSQL = "[SQL: UPDATE Usr_Key_Dates SET Date = '{0}' WHERE ID = {1}]".format(getSQLDate(newDate), myRow.iRowID)
    #_tikitResolver.Resolve(updateSQL)
  return


# New generic functions (because why copy-and-paste each time when we can make into own functions!)

def get_KeyDatesTableID(caseItemID = 0):
  # this function will return the ID of the row on the KeyDates table - if it doesn't exist, a new row will be added and its ID will be returned
  
  if caseItemID == 0:
    MessageBox.Show("There was an error trying to get the KeyDates table ID for item with CaseItemID: {0}".format(caseItemID), "Error - Getting KeyDates Table ID")
    return 0

  # count items from KeyDates table matching passed ID
  kdCount = runSQL("[SQL: SELECT COUNT(ID) FROM Usr_Key_Dates WHERE EntityRef = '{0}' AND MatterNo = {1} AND TaskStepID = {2}]".format(_tikitEntity, _tikitMatter, caseItemID))

  # if nothing currently matching in Key Dates, create a new row and get id
  if int(kdCount) == 0:
    runSQL("INSERT INTO Usr_Key_Dates (EntityRef, MatterNo, TaskStepID) VALUES ('{0}', {1}, {2})".format(_tikitEntity, _tikitMatter, caseItemID))

  # now get id of matching item
  tmpID = runSQL("SELECT TOP(1) ID FROM Usr_Key_Dates WHERE EntityRef = '{0}' AND MatterNo = {1} AND TaskStepID = {2}".format(_tikitEntity, _tikitMatter, caseItemID))

  return tmpID

def get_CurrentMatterDepartment():
  # This function will lookup the current matters' Department and put into a field on XAML to save needing to 're-get' this from SQL all the time
  if lbl_CurrentDept.Content == None or lbl_CurrentDept.Content == 'EXAMPLE':
    deptSQL = """SELECT D.Code FROM Matters M LEFT OUTER JOIN CaseTypes CT ON M.CaseTypeRef = CT.Code 
                 LEFT OUTER JOIN CaseTypeGroups CTG ON CT.CaseTypeGroupRef = CTG.ID 
                 LEFT OUTER JOIN Departments D ON CTG.Name = D.Description 
                 WHERE M.EntityRef = '{0}' AND M.Number = {1}""".format(_tikitEntity, _tikitMatter)
    currDept = runSQL(deptSQL, True, "There was an error getting the Department for the current Matter", "Error: Matter Department...") 
    lbl_CurrentDept.Content = currDept
    return currDept
  else:
    return lbl_CurrentDept.Content


def get_CaseHistoryAgendaID(titleForError = 'Error: Case History Agenda...'):
  # we need to get the 'Case History' agenda ID for SQL (unless other departments state they want all, in which case we need to amend this (or add options to 'Dept defaults')
  if lbl_CHAgendaID.Content == None or lbl_CHAgendaID.Content == '':
    cHAg_SQL = """SELECT TOP(1) CMA.ItemID FROM Cm_Agendas CMA JOIN Cm_CaseItems CI ON CMA.ItemID = CI.ItemID 
                  WHERE EntityRef = '{0}' AND MatterNo = {1} 
                  AND CI.Description = 'Case History'""".format(_tikitEntity, _tikitMatter)
    cHAgID = runSQL(cHAg_SQL, True, "There was an error getting the Case History Agenda ID", titleForError)
    lbl_CHAgendaID.Content = cHAgID
    return cHAgID
  else:
    return lbl_CHAgendaID.Content


def get_currentStepID(agendaID, titleForError = 'Error: Current Step ID...'):
  # This function will lookup the current highest item ID for items within the passed 'agendaID'

  if agendaID == 0:
    return 0

  # check to see if there's any documents in the Case History agenda
  countOfCaseItems = int(runSQL("SELECT COUNT(ItemID) FROM Cm_CaseItems WHERE ParentID = {0}".format(agendaID)))
  if countOfCaseItems == 0:
    return 0
  else:
    #return runSQL("SELECT MAX(ItemID) FROM Cm_CaseItems WHERE ParentID = {0}".format(agendaID))
    # get ItemId of the item with the highest 'Order'
    return runSQL("SELECT TOP 1 ItemID FROM Cm_CaseItems WHERE ParentID = {0} AND ItemOrder = (SELECT MAX(ItemOrder) FROM Cm_CaseItems WHERE ParentID = {0})".format(agendaID))


def validate_DateDueDateTime(s, event):
  # This function removes any unwanted characters from the text box (only numbers and colon allowed)
  txt_DD_Time.Text = stripString(sourceString = txt_DD_Time.Text, leaveOnly = '1234567890:')
  return


def validate_DateDuration(s, event):
  # This function removes any unwanted characters from the text box (only numbers allowed)
  txt_DD_Duration.Text = stripString(sourceString = txt_DD_Duration.Text, leaveOnly = '1234567890')
  return

def validate_DateReminder(s, event):
  # This function removes any unwanted characters from the text box (only numbers allowed)
  txt_DD_ReminderQty.Text = stripString(sourceString = txt_DD_ReminderQty.Text, leaveOnly = '1234567890')
  return

def stripString(sourceString, leaveOnly):
  # This function has two parameters, first for defining the source string, and the second to specify the 'allowed' characters to return

  # set initial variables (force setting 'sourceString' to a string)
  sString = str(sourceString)
  nuString = ''

  # iterate over each character in source String
  for myChar in sString:
    # if character is in the 'leaveOnly' (allowedCharacters)
    if myChar in leaveOnly:
      # append character to our nuString variable
      nuString += myChar

  # finally return the final string
  return nuString


#class columnUpdates(object):
#  def __init__(self, myColNameAndValue):
#    self.ColNameAndValue = myColNameAndValue
#    return
#
#  def __getitem__(self, index):
#    if index == 'Item':
#      return self.ColNameAndValue
# not using the above now - was planning to use this for creating a list of columns to update (for 'cellEdit_Finished' functions)
# Idea was to use this to create a list of these 'objects' and then at end, iterate over and append a comma (to ensure we get 
# proper SQL, separating multiple columns updated)

def dg_KeyDates_cellEdit_Finished(s, event):
  # New - September 2024 - Louis wants to be able to edit in the DataGrid itself, rather than via the editable controls at bottom.
  # Therefore, need to splice-up 'addOrUpdate_KeyDate' function (line 1491) to ensure we both validate input and update applicable tables.
  # Note - shouldn't allow updating of 'Completed' items other than 'DateMissedNote'
  
  tmpColName = event.Column.Header # firstly, store current/active column into a variable
  tmpDebugTitle = "DEBUGGING - Editing in DataGrid..."
  
  # setup initial varialbles (SQL to update table, get ID for respective table, reset count of updates per table)
  uSQL_Diary = uSQL_CmS = uSQL_CmCI = uSQL_UsrKD = ""

  # set counts to zero
  dCount = cmsCount = cmciCount = usrKDCount = 0
  
  # get initial values from DataGrid
  mStatus = dg_KeyDates.SelectedItem['Grouping']
  diaryID = dg_KeyDates.SelectedItem['DAcode']
  cMsID = dg_KeyDates.SelectedItem['CaseStepID']
  uKDID = 0 if dg_KeyDates.SelectedItem['RowID'] is None else dg_KeyDates.SelectedItem['RowID']
  
  debugMessage(msgTitle = tmpDebugTitle, msgBody = "diaryID: {0}   - CaseManager ID: {1}   - KeyDates table ID: {2}".format(diaryID, cMsID, uKDID))

  # Conditionally add parts depending on column updated and whether value has changed
  if tmpColName == 'Description' and mStatus != '2) Complete':
    newName = dg_KeyDates.SelectedItem['Desc']

    if newName != txt_DD_Desc.Text:
      # make text SQL safe (replace single quotes with double)
      newName = sql_safe_string(stringToClean = newName)
      uSQL_Diary += "Description = '{0}' ".format(newName) 
      dCount += 1
      uSQL_CmCI += "Description = '{0}' ".format(newName) 
      cmciCount += 1
      uSQL_UsrKD += "Description = '{0}' ".format(newName) 
      usrKDCount += 1
    debugMessage(msgTitle = tmpDebugTitle, msgBody = "Diary SQL:\n{0}\n\nCaseItems SQL:\n{1}\n\nKeyDates SQL:\n{2}".format(uSQL_Diary, uSQL_CmCI, uSQL_UsrKD))


  if tmpColName == 'Location' and mStatus != '2) Complete':
    newLocation = dg_KeyDates.SelectedItem['Location']

    if newLocation != txt_DD_Location.Text:
      # make Locaiton text SQL safe (replace single quote with double)
      newLocation = sql_safe_string(stringToClean = newLocation)
      uSQL_Diary += "Location = '{0}' ".format(newLocation) 
      dCount += 1
      uSQL_UsrKD += "Location = '{0}' ".format(newLocation) 
      usrKDCount += 1
    debugMessage(msgTitle = tmpDebugTitle, msgBody = "Diary SQL:\n{0}\n\nKey Dates SQL:\n{1}".format(uSQL_Diary, uSQL_UsrKD))


  if tmpColName == 'Assigned To' and mStatus != '2) Complete':
    newAssignedTo = dg_KeyDates.SelectedItem['AssignedTo']
    #MessageBox.Show("newAssignedTo: {0}".format(newAssignedTo), "DEBUGGING - Editing in DataGrid")
    
    #if newAssignedTo != cbo_Date_AssignedTo.SelectedItem['Code']:
    # following our revelation with the 'Duration' drop-down (changing it in DG immediately updated 'combo box' in 'edit' area), we can't test
    # if value is different to combo box as they will alwats be the same value and code would never trigger. So just updating always.
    # We could perhaps add labels tp form to act as the 'static' (comparrison) values, but feels redundant as may just find that they update too

    # new check for in case Assigned To is missing
    if newAssignedTo is None or newAssignedTo == '':
      uSQL_Diary += "Username = NULL "
      dCount += 1
      uSQL_UsrKD += "AssignedTo = NULL "
      usrKDCount += 1
    else:
      uSQL_Diary += "Username = '{0}' ".format(newAssignedTo) 
      dCount += 1
      uSQL_UsrKD += "AssignedTo = '{0}' ".format(newAssignedTo) 
      usrKDCount += 1
    debugMessage(msgTitle = tmpDebugTitle, msgBody = "Diary SQL:\n{0}\n\n KeyDate SQL:\n{1}".format(uSQL_Diary, uSQL_UsrKD))


  if tmpColName == 'Date' and mStatus != '2) Complete':
    # form the 'new' date and time
    newDueDate = getSQLDate(dg_KeyDates.SelectedItem['Date'])
    newTime = getTimeFixed(dg_KeyDates.SelectedItem['Time'])
    actualDate = "{0} {1}".format(newDueDate, newTime)
    # get old date - and make sure we only grab first 10 characters
    oldDDate = getSQLDate(dp_DD_Date.SelectedDate)
    oldDDate = oldDDate[:10]
    
    if newDueDate != oldDDate:
      uSQL_Diary += "DateStamp = '{0}' ".format(actualDate) 
      dCount += 1
      uSQL_CmS += "StartDate = '{0}' ".format(actualDate) 
      cmsCount += 1
      uSQL_UsrKD += "Date = '{0}' ".format(actualDate) 
      usrKDCount += 1
      # NB we also update the StepCategory once date has been entered
      if mStatus[:1] == '0':
        uSQL_CmS += ", StepCategory = 'KeyD' " 
    debugMessage(msgTitle = tmpDebugTitle, msgBody = "Diary SQL:\n{0}\n\nCmSteps SQL:\n{1}\n\nKey Dates SQL:\n{2}".format(uSQL_Diary, uSQL_CmS, uSQL_UsrKD))
  
    # new version of code
    #newDueDate = dg_KeyDates.SelectedItem['Date']
    #newTime = dg_KeyDates.SelectedItem['Time']
    #if newDueDate is None:
    # I don't know I want to do this because code appears to be working fine as is, and I've now realised what the issue probably was...
    # looks like I didn't include 'getSQLDate()' around 'ReminderDate' (see Tasks update - line 1004), which meant I had to hanndle null
    # values separately... when I could've just updated 'getSQLDate()' accordingly (if needed)

  if tmpColName == 'Time' and mStatus != '2) Complete':
    newTime = getTimeFixed(dg_KeyDates.SelectedItem['Time'])
    oldTime = getTimeFixed(dg_KeyDates.SelectedItem['oldDueTime'])
    tmpDDate = getSQLDate(dp_DD_Date.SelectedDate)
    tmpDDate = tmpDDate[:10]
    actualDate = "{0} {1}".format(tmpDDate, newTime)
    if newTime != oldTime:
      uSQL_Diary += "DateStamp = '{0}' ".format(actualDate) 
      dCount += 1
      uSQL_CmS += "StartDate = '{0}' ".format(actualDate) 
      cmsCount += 1
      uSQL_UsrKD += "Date = '{0}' ".format(actualDate) 
      usrKDCount += 1
    debugMessage(msgTitle = tmpDebugTitle, msgBody = "Diary SQL:\n{0}\n\nCmSteps SQL:\n{1}\n\nKey Dates SQL:\n{2}".format(uSQL_Diary, uSQL_CmS, uSQL_UsrKD))


  # NEW Duration (combined both 'qty' and 'type' into one column as looks better)
  if tmpColName == 'Duration' and mStatus != '2) Complete':
    # append duration type first
    tmpDurTypeN = cbo_DD_DurType.SelectedItem['Code']
    uSQL_Diary += "DurationType = {0} ".format(tmpDurTypeN)
    dCount += 1
    uSQL_CmS += "DurationType = {0} ".format(tmpDurTypeN)
    cmsCount += 1
    uSQL_UsrKD += "DurationTypeN = {0} ".format(tmpDurTypeN) 
    usrKDCount += 1

    # now append duration (units)
    newDurQty = dg_KeyDates.SelectedItem['Duration']
    tmpDurQty = txt_DD_Duration.Text

    if newDurQty != tmpDurQty:
      uSQL_Diary += ", Duration = {0} ".format(newDurQty)
      dCount += 1
      uSQL_CmS += ", Duration = {0} ".format(newDurQty)
      cmsCount += 1
      uSQL_UsrKD += ", DurationQty = {0} ".format(newDurQty) 
      usrKDCount += 1
    debugMessage(msgTitle = tmpDebugTitle, msgBody = "Diary SQL:\n{0}\n\nCmSteps SQL:\n{1}\n\nKey Dates SQL:\n{2}".format(uSQL_Diary, uSQL_CmS, uSQL_UsrKD))

    # issue here as we're updating multiple columns and not adding a comma inbetween
    # 'simple fix' - I've moved DurationType to be first, and added comma to the Duration Quantity part as that's conditional anyway
    # Ideally, instead of a string that we add to (as we normally do), we could make a list, and at end, loop over said list to form the final string
    # adding in the applicable commas as necessary.  I have already written a small 'class' and working example so look to that if wanting to implement here
  if tmpColName == 'Reminder' and mStatus != '2) Complete':
    newRemQty = dg_KeyDates.SelectedItem['Reminder']
    tmpRemQty = txt_DD_ReminderQty.Text
    #newRemType = dg_KeyDates.SelectedItem['RemType']
    tmpRemTypeN = cbo_DD_RemType.SelectedItem['Code']

    # append 'Type' first, as we can't conditionaly test (because combo boxes updates same time as updating datagrid)
    uSQL_Diary += "ReminderType = {0} ".format(tmpRemTypeN)
    dCount += 1
    uSQL_CmS += "ReminderType = {0} ".format(tmpRemTypeN)
    cmsCount += 1
    uSQL_UsrKD += "ReminderTypeN = {0} ".format(tmpRemTypeN) 
    usrKDCount += 1

    if newRemQty != tmpRemQty:
      uSQL_Diary += ", Reminder = {0} ".format(newRemQty)
      dCount += 1
      uSQL_CmS += ", Reminder = {0} ".format(newRemQty)
      cmsCount += 1
      uSQL_UsrKD += ", ReminderQty = {0} ".format(newRemQty) 
      usrKDCount += 1   
    debugMessage(msgTitle = tmpDebugTitle, msgBody = "Diary SQL: {0}\nCmSteps SQL: {1}\nKey Dates SQL: {2}".format(uSQL_Diary, uSQL_CmS, uSQL_UsrKD))


  if tmpColName == 'Date Missed Notes':
    newDMN = dg_KeyDates.SelectedItem['DateMissedNotes']
    tmpDMN = txt_DateMissedNotes.Text
    if newDMN != tmpDMN:
      # make SQL safe text (replace single quote with double)
      newDMN = sql_safe_string(stringToClean = newDMN)
      uSQL_UsrKD += "DateMissedNotes = '{0}' ".format(newDMN)
      usrKDCount += 1
    debugMessage(msgTitle = tmpDebugTitle, msgBody = "Date Missed Notes SQL: {0}".format(uSQL_UsrKD))


  # if there's any update to Diary table field and we actually have a Diary table ID - do actual update
  if dCount > 0 and int(diaryID) > 0: 
    sqlToRun = """UPDATE Diary_Appointments SET {0}, UpdateThis = {1} 
                  WHERE Code = {2} AND EntityRef = '{3}' AND MatterNoRef = {4} """.format(uSQL_Diary, gUpdateThis, diaryID, _tikitEntity, _tikitMatter)
    #uSQL_Diary += ", UpdateThis = {3} WHERE Code = {0} AND EntityRef = '{1}' AND MatterNoRef = {2}".format(diaryID, _tikitEntity, _tikitMatter, gUpdateThis)
    if debugMessage(msgTitle = tmpDebugTitle, msgBody = "Diary Table SQL:\n{0}".format(sqlToRun)) == True:
      runSQL(sqlToRun)
  else:
    appendToDebugLog(textToAppend="Cannot update Diary_Appointments table, as record does not exist (no 'ID'/'Code') or no data to update in this table", inclTimeStamp=True)

  # if we have a Case Manager ID - do updates to 'Cm_Steps' and 'Cm_CaseItems' tables  
  if cMsID > 0:
    if cmsCount > 0:
      sqlToRun = "UPDATE Cm_Steps SET {0} WHERE ItemID = {1}".format(uSQL_CmS, cMsID)
      #uSQL_CmS += "WHERE ItemID = {0}".format(cMsID)
      if debugMessage(msgTitle = tmpDebugTitle, msgBody = "Cm_Steps Table SQL:\n{0}".format(sqlToRun)) == True:
        runSQL(sqlToRun)
    
    if cmciCount > 0:
      sqlToRun = "UPDATE Cm_CaseItems SET {0} WHERE ItemID = {1}".format(uSQL_CmCI, cMsID)
      #uSQL_CmCI += "WHERE ItemID = {0}".format(cMsID)
      if debugMessage(msgTitle = tmpDebugTitle, msgBody = "Cm_CaseItems Table SQL:\n{0}".format(sqlToRun)) == True:
        runSQL(sqlToRun)
  
  # if there have been any updates to KeyDate backup table fields, do actual updates
  if usrKDCount > 0:
    # if we dont yet have an ID for our KeyDate backup table, get one
    if int(uKDID) == 0:
      uKDID = get_KeyDatesTableID(caseItemID = cMsID)

    # do actual update
    if uKDID != 0:
      sqlToRun = "UPDATE Usr_Key_Dates SET {0} WHERE ID = {1}".format(uSQL_UsrKD, uKDID)
      #uSQL_UsrKD += "WHERE ID = {0}".format(uKDID)
      if debugMessage(msgTitle = tmpDebugTitle, msgBody = "Usr_Key_Dates Table SQL:\n{0}".format(sqlToRun)) == True:
        runSQL(sqlToRun)
    else:
      appendToDebugLog(textToAppend="Cannot update KeyDates backup table, as record does not exist (no 'ID'/'Code')", inclTimeStamp=True)


  # set update this for all attendees (so they get updates synced to outlook)
  if (dCount + cmsCount + cmciCount + usrKDCount > 0) and int(diaryID) > 0:
    countOfAttendeeItems = runSQL("SELECT COUNT(ID) FROM Diary_AppointmentsAttendees WHERE AppointmentRef = {0}".format(diaryID))
    if int(countOfAttendeeItems) > 0:
      runSQL("UPDATE Diary_AppointmentsAttendees SET UpdateThis = {0} WHERE AppointmentRef = {1}".format(gUpdateThis, diaryID))
  
  refresh_KeyDates(s, event)

  # select row again as we refreshed
  if dg_KeyDates.Items.Count > 0 and int(cMsID) > 0:
    xCount = -1
    for xRow in dg_KeyDates.Items:
      xCount += 1
      if xRow.iCaseStepID is not None:
        if xRow.iCaseStepID == cMsID:
          dg_KeyDates.SelectedIndex = xCount
          break
  return


def debugMessage(msgTitle = 'DEBUG MESSAGE', msgBody = '', askToConfirmSQL = False):
  # This function will print the SQL used to the debug text box and asks user if OK to run
  # It has two arguments: msgTitle - title to use for message box prompt (optional)
  # and 'msgBody' which is the text to print to the output and is also displayed in message box prompt

  endLine = "-----------------------------------------------------------------------------------------------------------"
  
  # Get the current date and time and format it
  current_datetime = datetime.now()
  formatted_datetime = current_datetime.strftime("%Y-%m-%d %H:%M:%S")
  
  # get current debug text into a variable so we can reformat with other details
  originalText = txt_DebugModeOutput.Text
  # start forming new text, eg
  #--------------------------------------------------------------------------
  #<2024-09-12 11:39:00>
  #> SELECT * FROM Matters WHERE FeeEarnerRef = 'MP'
  #newText = "{0}\n<{1}>\n> {2}\n".format(endLine, formatted_datetime, msgBody)
  #--------------------------------------------------------------------------
  # 2024-09-12 11:39:00> SELECT * FROM Matters WHERE FeeEarnerRef = 'MP'
  newText = "{0}\n{1}> {2}\n".format(endLine, formatted_datetime, msgBody)

  # if checkbox isn't checked (eg: current user not me or Louis, and not ticked box)
  if chk_DebugMode.IsChecked == False:
    textToAppend = ""
    returnVal = True
  else:
    # if elected to not prompt for confirmation to run SQL, just return true
    if askToConfirmSQL == False:
      textToAppend = "> askToConfirmSQL = False (eg will run SQL)"
      returnVal = True
    else:
      # ask if OK to run the SQL and return appropriate boolean answer for calling procedure 
      dMsgResult = MessageBox.Show(msgBody + "\n\nDo you want to continue?", msgTitle, MessageBoxButtons.YesNo)
      if dMsgResult == DialogResult.Yes: 
        textToAppend = "> askToConfirmSQL = True | Answered 'Yes'"
        returnVal = True
      else:
        textToAppend = "> askToConfirmSQL = True | Answered 'No'"
        returnVal = False

  # put text into debug log and return value to calling procedure
  txt_DebugModeOutput.Text = "{0}\n{1}\n{2}".format(newText, textToAppend, originalText)
  return returnVal


def appendToDebugLog(textToAppend = '', inclEndLine = False, inclTimeStamp = False):
  # This function will append passed text to the end of the 'debug' text box and put a line under each call
  endLine = "-----------------------------------------------------------------------------------------------------------"
  
  # Get the current date and time and format it
  current_datetime = datetime.now()
  formatted_datetime = current_datetime.strftime("%Y-%m-%d %H:%M:%S")
  
  # get current text into a variable
  tmpText = txt_DebugModeOutput.Text
  txt_DebugModeOutput.Text = "{0}\n<{1}>\n> {2}\n{3}".format(endLine, formatted_datetime, textToAppend, tmpText)
  
  # now append passed text
  if inclTimeStamp == True:
    txt_DebugModeOutput.Text = "<{0}>\n> {1}\n{2}".format(formatted_datetime, textToAppend, tmpText)
  else:
    txt_DebugModeOutput.Text = tmpText + "\n> {0}\n".format(textToAppend)
  
  if inclEndLine == True:
    txt_DebugModeOutput.Text = txt_DebugModeOutput.Text + "\n{0}\n".format(endLine)
  return


def find_DateAttendee(s, event):
  # This event fires whenever text is entered / deleted from the 'Search' box under the Attendees side-panel
  # Essentially just calls a 'refresh' to the list (which also searches entered text) whilst still keeping ticked users visible
  # NB: could add grouping to this list too (eg Included; not included/search results)
  
  if dg_KeyDates.SelectedIndex != -1:
    refresh_AttendeeList(usersToPreSelect = dg_KeyDates.SelectedItem['Attendees'])
  return

def dateAttendee_cellUpdated(s, event):
  # This function will update the 'Attendee' list for the selected Date item (eg: removing or adding selected participant)
  
  if dg_KeyDates.SelectedIndex == -1:
    return
    
  # get diary table ID and other variables
  #tmpCol = event.Column
  #tmpColName = tmpCol.Header
  tmpColName = event.Column.Header
  dairyTblID = dg_KeyDates.SelectedItem['DAcode']
  caseItemID = dg_KeyDates.SelectedItem['CaseStepID']
  inclStatus = dg_DD_Attendees.SelectedItem['Ticked']
  userCode = dg_DD_Attendees.SelectedItem['Code']
  
  if tmpColName == 'Incl.?':
    if inclStatus == False:
      # here we want to remove the selected user from the Attendees table
      runSQL("UPDATE Diary_AppointmentsAttendees SET DeleteThis = 1 WHERE AppointmentRef = {0} AND Username = '{1}'".format(dairyTblID, userCode))
      
    elif inclStatus == True:
      # here we want to add the selected user to the Attendees table
      aSQL = "INSERT INTO Diary_AppointmentsAttendees (AppointmentRef, StepItemID, Username, DeleteThis, UpdateThis) "
      aSQL += "VALUES({0}, {1}, '{2}', 0, {3})".format(dairyTblID, caseItemID, userCode, gUpdateThis)
      runSQL(aSQL)
  
  currDGIndex = dg_KeyDates.SelectedIndex
  refresh_KeyDates(s, event)
  dg_KeyDates.SelectedIndex = currDGIndex
  return


def date_MarkComplete(s, event):
  # This function will mark the selected date as complete and will refresh the DataGrid

  # if nothing selected, quit now
  if dg_KeyDates.SelectedIndex == -1:
    return

  # temp store date description
  tmpDesc = dg_KeyDates.SelectedItem['Desc']

  # double-check ok to continue
  confirmContinue = MessageBox.Show("Are you sure you want to mark the selected Date as 'Complete'?'\n\n{0}".format(tmpDesc), "Mark 'Complete' confirmation...", MessageBoxButtons.YesNo)
  if confirmContinue == DialogResult.No:
    return
  
  # now get other values needed for input
  dDesc = sql_safe_string(stringToClean = tmpDesc)
  dLocation = sql_safe_string(stringToClean = dg_KeyDates.SelectedItem['Location'])
  dAssignedTo = dg_KeyDates.SelectedItem['AssignedTo']
  dDate = getSQLDate(dg_KeyDates.SelectedItem['Date'])
  dDurationQty = dg_KeyDates.SelectedItem['Duration']
  dDurationTypeN = dg_KeyDates.SelectedItem['DurType']
  dRemindQty = dg_KeyDates.SelectedItem['Reminder']
  dRemindTypeN = dg_KeyDates.SelectedItem['RemType']
  dAttendees = dg_KeyDates.SelectedItem['Attendees']
  dDateMissedN = sql_safe_string(stringToClean = dg_KeyDates.SelectedItem['DateMissedNotes'])

  # get table IDs
  dTblID = dg_KeyDates.SelectedItem['DACode']
  caseItemID = dg_KeyDates.SelectedItem['CaseStepID']
  kdID = dg_KeyDates.SelectedItem['RowID']
  if kdID == None or kdID == 0:
    kdID = get_KeyDatesTableID(caseItemID)
  
  # Before Marking as complete, we should ensure that all info is copied to our KeyDates table for safe keeping
  # Point to note - if we create a 'Revert' button too, then we will need to re-create 'Diary_' table item but 
  # note we should also update the 'Code' in KeyDates (to get NEW id)
  updateKD_SQL = """UPDATE Usr_Key_Dates SET AssignedTo = '{0}', Date = '{1}', DurationQty = {2}, DurationTypeN = {3}, 
                    Description = '{4}', ReminderQty = {5}, ReminderTypeN = {6}, Attendees = '{7}', Location = '{8}', 
                    DateMissedNotes = '{9}', TaskOrDate = 'Date' WHERE ID = {10}""".format(dAssignedTo, dDate, dDurationQty, dDurationTypeN, 
                                                                                           dDesc, dRemindQty, dRemindTypeN, dAttendees, dLocation, dDateMissedN, kdID)
  
  if debugMessage(msgTitle = 'DEBUG MESSAGE - Testing Date backup', msgBody = updateKD_SQL) == True:
    runSQL(updateKD_SQL, True, "There was an error updating the KeyDates backup table", "ERROR - KeyDate Backup...")
        
  # now for actual code to 'complete'
  # According to SQL Profiler / Trace, following is what happens upon clicking 'Complete' on a 'Date'
  # UPDATE Cm_CaseItems SET CompletionDate = GETDATE()
  # UPDATE Cm_Steps SET DiaryDate = NULL
  # INSERT INTO Cm_Steps_ActionHistory (ItemID, UserID, StepActionID, Description) VALUES(id, _tikitUser, 12, NULL)
  # (FYI 12 = Completed - if need a button for revert, should use code 2 if updating history)
  # Interestingly, no updates are made to Diary_Appointments table as I would've expected - therefore
  # code as-is for now and test to see what happens - we can always manually add it in

  # run SQL as mentioned above
  runSQL("UPDATE Cm_CaseItems SET CompletionDate = GETDATE() WHERE ItemID = {0}".format(caseItemID))
  runSQL("UPDATE Cm_Steps SET DiaryDate = NULL WHERE ItemID = {0}".format(caseItemID))
  runSQL("INSERT INTO Cm_Steps_ActionHistory (ItemID, UserID, StepActionID, Description) VALUES ({0}, '{1}', {2}, NULL)".format(caseItemID, _tikitUser, 12))

  refresh_KeyDates(s, event)
  return

def task_MarkComplete(s, event):
  # This function will mark the selected task as complete and will refresh the DataGrid

  # if nothing selected, quit now
  if dg_KeyTasks.SelectedIndex == -1:
    return

  # temp store date description
  tmpDesc = dg_KeyTasks.SelectedItem['Desc']

  # double-check ok to continue
  confirmContinue = MessageBox.Show("Are you sure you want to mark the selected Date as 'Complete'?'\n\n{0}".format(tmpDesc), "Mark 'Complete' confirmation...", MessageBoxButtons.YesNo)
  if confirmContinue == DialogResult.No:
    return

  # now get other values needed for input
  tDesc = sql_safe_string(stringToClean = tmpDesc)
  # tStatus = dg_KeyTasks.SelectedItem['Status']
  tStatus = '2'
  tPriority = dg_KeyTasks.SelectedItem['Priority']
  # tPCComplete = dg_KeyTasks.SelectedItem['PercentComplete']
  tPCComplete = '100'
  tAssignedTo = dg_KeyTasks.SelectedItem['AssignedTo']
  tDate = getSQLDate(dg_KeyTasks.SelectedItem['Date'])
  tRemindTime = getTimeFixed(dg_KeyTasks.SelectedItem['ReminderTime'])
  tRemindDate = getSQLDate(dg_KeyTasks.SelectedItem['ReminderDate']) + " {0}.000".format(tRemindTime)
  tDateMissedN = sql_safe_string(stringToClean = dg_KeyTasks.SelectedItem['DateMissedNotes'])
  
  # get table IDs
  tTblID = dg_KeyTasks.SelectedItem['Code']
  caseItemID = dg_KeyTasks.SelectedItem['CaseStepID']
  kdID = dg_KeyTasks.SelectedItem['RowID']
  if kdID == None or kdID == 0:
    kdID = get_KeyDatesTableID(caseItemID)
    
  # Before Marking as complete, we should ensure that all info is copied to our KeyDates table for safe keeping
  # Point to note - if we create a 'Revert' button too, then we will need to re-create 'Diary_' table item but 
  # note we should also update the 'Code' in KeyDates (to get NEW id)
  updateKD_SQL = """UPDATE Usr_Key_Dates SET AssignedTo = '{0}', Date = '{1}', DurationQty = 1, DurationTypeN = 2, 
                    Description = '{2}', ReminderQty = 15, ReminderTypeN = 0, ReminderDate = '{3}', ReminderTime = '{4}', 
                    Status = {5}, Priority = {6}, PCComplete = {7}, DateMissedNotes = '{8}', TaskOrDate = 'Task' WHERE ID = {9}""".format(tAssignedTo, tDate, tDesc, tRemindDate, tRemindTime, 
                                                                                                                                          tStatus, tPriority, tPCComplete, tDateMissedN, kdID)

  if debugMessage(msgTitle = 'DEBUG MESSAGE - Testing Task backup', msgBody = updateKD_SQL) == True:
    runSQL(updateKD_SQL, True, "There was an error updating the KeyDates backup table", "ERROR - KeyDate Backup...")

  # now for actual code to 'complete'
  # According to SQL Profiler / Trace, following is what happens upon clicking 'Complete' on a 'Task'
  # UPDATE Cm_Steps SET DiaryDate = (date), SentDate = GETDATE(), Taker = _tikitUser, FileName = '', SAMTakeId = 0
  # INSERT INTO Cm_Steps_ActionHistory (ItemID, UserID, StepActionID, Description) VALUES (, _tikitUser, 4, NULL)  # NB: 4 = Processed
  # DELETE FROM Diary_Tasks WHERE Code = {0}
  # UPDATE Cm_CaseItems SET CompletionDate = GETDATE()
  # UPDATE Cm_Steps SET DiaryDate = NULL
  # INSERT INTO Cm_Steps_ActionHistory (ItemID, UserID, StepActionID, Description) VALUES(, , 12, NULL)
  
  # run SQL as mentioned above
  runSQL("UPDATE Cm_Steps SET DiaryDate = '{0}', SentDate = GETDATE(), Taker = '{1}', FileName = '', SAMTakeId = 0 WHERE ItemID = {2}".format(tDate, _tikitUser, caseItemID))
  runSQL("INSERT INTO Cm_Steps_ActionHistory (ItemID, UserID, StepActionID, Description) VALUES ({0}, '{1}', 4, NULL)".format(caseItemID, _tikitUser))
  # now we want to update Outlook with latest version of info, so set the 'UpdateThis' to 1 to trigger update
  runSQL("UPDATE Diary_Tasks SET UpdateThis = 1 WHERE Code = {0}".format(tTblID))
  # we then need to delete item from 'Diary_' table - however, not before a sync has happened 
  # (NB: we need the Tikit Exchange Connector as currently it doesn't have a chance to sync before deleting item from Diary_ table)
  runSQL("DELETE FROM Diary_Tasks WHERE Code = {0}".format(tTblID))
  runSQL("UPDATE Cm_CaseItems SET CompletionDate = GETDATE() WHERE ItemId = {0}".format(caseItemID))
  runSQL("UPDATE Cm_Steps SET DiaryDate = NULL WHERE ItemID = {0}".format(caseItemID))
  runSQL("INSERT INTO Cm_Steps_ActionHistory (ItemID, UserID, StepActionID, Description) VALUES({0}, '{1}', 12, NULL)".format(caseItemID, _tikitUser))
  
  refresh_KeyTasks(s, event)
  return


def deleteDate(s, event):
  # This function will delete the currently selected Date (on 'Appointments' tab) - after confirmation from user

  if dg_KeyDates.SelectedIndex != -1:
    myMessage = "Are you sure you want to remove the following Date?\n{0}".format(dg_KeyDates.SelectedItem['Desc'])
    result = MessageBox.Show(myMessage, 'Confirm deletion of Date...', MessageBoxButtons.YesNo)
  
    if result == DialogResult.Yes:
      # if has a CaseItem ID then we will want to remove from there too (as well as Diary_Tasks)
      # Come to think of it, users will not have ability to delete case items, therefore wondering how this would play out - would it error (we shall do anyway)
      tmpCaseStepID = 0 if dg_KeyDates.SelectedItem['CaseStepID'] == None else dg_KeyDates.SelectedItem['CaseStepID']
      tmpAgendaID = 0 if dg_KeyDates.SelectedItem['Agenda'] == None else dg_KeyDates.SelectedItem['Agenda']
      tmpKDid = dg_KeyDates.SelectedItem['RowID']

      if tmpCaseStepID > 0:
        # get the order ID as we'll need this later to update order of case items table
        tmpCIorder = runSQL("SELECT ItemOrder FROM Cm_CaseItems WHERE ItemID = {0}".format(tmpCaseStepID))
        # Update DeleteThis in Diary_Tasks (to delete Outlook item)
        runSQL("UPDATE Diary_Appointments SET DeleteThis = 1 WHERE CaseItemRef = {0}".format(tmpCaseStepID))
        # may need to add code here to check if any Attendees and delete from Diary_Attendees table too (if above doesn't auto delete for them)
        runSQL("UPDATE Diary_AppointmentsAttendees SET DeleteThis = 1 WHERE StepItemID = {0}".format(tmpCaseStepID))

        # Now form and run our DELETE FROM tables
        runSQL("DELETE FROM Cm_Steps WHERE ItemID = {0}".format(tmpCaseStepID))
        runSQL("DELETE FROM Cm_CaseItems WHERE ItemID = {0}".format(tmpCaseStepID))
        # for sake of tidying up database, ought to also delete from CaseActionHistory table too
        runSQL("DELETE FROM Cm_Steps_ActionHistory WHERE ItemID = {0}".format(tmpCaseStepID))

        if tmpAgendaID > 0:
          # Update ItemOrder - decrease all items with Order GREATER than selected item
          runSQL("UPDATE Cm_CaseItems SET ItemOrder = (ItemOrder - 1) WHERE ParentID = {0} AND ItemOrder > {1}".format(tmpAgendaID, tmpCIorder))

      if tmpKDid != None and tmpKDid > 0:
        # Delete the item from our Key Dates table (We do this regardless of whether Case item exists or not)
        if tmpCaseStepID > 0:
          runSQL("DELETE FROM Usr_Key_Dates WHERE TaskStepID = {0} AND ID = {1}".format(tmpCaseStepID, tmpKDid))
        else:
          runSQL("DELETE FROM Usr_Key_Dates WHERE ID = {0}".format(tmpKDid))

      # now refresh the datagrid to reflect updates
      refresh_KeyDates(s, event)
  return


def revertDate(s, event):
  # This function will revert (re-add) the selected 'Completed' Date
  # Add to CaseAction History table (with 'reverted' number (2))
  # Add to 'Diary Appointments' table (with applicable Case Item ID)
  # nullify 'completion date' in CaseItems table
  # saving all details onto our Key Dates table for backup
  
  if dg_KeyDates.SelectedIndex == -1:
   return

  if dg_KeyDates.SelectedItem['Grouping'] != '2) Complete':
    return

  # get variables needed for SQL
  caseItemID = dg_KeyDates.SelectedItem['CaseStepID']
  feREF = runSQL("SELECT FeeEarnerRef FROM Matters WHERE EntityRef = '{0}' AND Number = {1}".format(_tikitEntity, _tikitMatter))
  agendaID = dg_KeyDates.SelectedItem['Agenda']
  dDate = getSQLDate(dg_KeyDates.SelectedItem['Date'])
  dTime = getTimeFixed(dg_KeyDates.SelectedItem['Time'])
  # check to see if date is in the past (if so, get tomorrows date)
  newDueDate = getSQLDate(runSQL("SELECT CASE WHEN '{0}' < GETDATE() THEN DATEADD(day, 1, GETDATE()) ELSE '{0}' END".format(dDate)))
  fullDueDateAndTime = "{0} {1}".format(newDueDate, dTime)
  
  dDuration = dg_KeyDates.SelectedItem['Duration']
  dType = dg_KeyDates.SelectedItem['DurType']
  
  dReminder = dg_KeyDates.SelectedItem['Reminder']     #getSQLDate(dg_KeyDates.SelectedItem['Reminder'])
  dRemType = dg_KeyDates.SelectedItem['RemType']
  #newRemDate = getSQLDate(runSQL("SELECT CASE WHEN '{0}' < GETDATE() THEN DATEADD(day, 1, GETDATE()) ELSE '{0}' END".format(dReminder)))
  #fullReminderDateAndTime = "{0} {1}".format(newRemDate, dRemType)
  
  dDesc = getUniqueDescription(desiredDesc = dg_KeyDates.SelectedItem['Desc'], forUser = feREF, taskORdate = 'Date')
  dDesc = sql_safe_string(stringToClean = dDesc)
  dLocation = sql_safe_string(stringToClean = dg_KeyDates.SelectedItem['Location'])
  dDateMissedN = sql_safe_string(stringToClean = dg_KeyDates.SelectedItem['DateMissedNotes'])

  # get id for our Key Dates table
  kdID = dg_KeyDates.SelectedItem['RowID']
  if kdID == None or kdID == 0:
    kdID = get_KeyDatesTableID(caseItemID)
  
  # make a description for Cm Steps History table
  rDesc = "Key Date - Original details: Description: {0}; DueDate: {1}".format(dDesc, fullDueDateAndTime)

  # update Steps ActionHistory table to state current user is reverting this task
  runSQL("INSERT INTO Cm_Steps_ActionHistory (ItemID, UserID, StepActionID, Description) VALUES ({0}, '{1}', {2}, '{3}')".format(caseItemID, _tikitUser, 2, rDesc))

  # add new entry into Diary Tasks table (note: use FeeEarner of matter rather than current user when setting person assigned to)
  # NB: Reminder in this table is the number of mins before, and Type is the number representing 0, 1, 2, 3 (Minutes, Hours, Days, Weeks)
  dateSQL = """INSERT INTO Diary_Appointments(Username, DateStamp, Duration, DurationType, Reminder, ReminderType, 
                          [Description], EntityRef, MatterNoRef, AgendaRef, CaseItemRef, UserType, Location)
               VALUES('{0}', '{1}', {2}, {3}, {4}, {5}, '{6}', '{7}', {8}, {9}, {10}, 'A', '{11}')""".format(feREF, fullDueDateAndTime, dDuration, dType, 
                                                                dReminder, dRemType, dDesc, _tikitEntity, _tikitMatter, agendaID, caseItemID, dLocation)
  runSQL(dateSQL)

  # now we need to get the ID of the added item
  dateTblID = runSQL("SELECT Code FROM Diary_Appointments WHERE Username = '{0}' AND Description = '{1}' AND DateStamp = '{2}' AND EntityRef = '{3}' AND MatterNoRef = {4}".format(feREF, dDesc, fullDueDateAndTime, _tikitEntity, _tikitMatter))

  # if there's any Attendees, then we ought to re-add for them too, by adding to 'Attendees' table
  tmpAttendees = dg_KeyDates.SelectedItem['Attendees']
  if tmpAttendees != None and len(tmpAttendees) > 0 and len(dateTblID) > 0:
    # break up list of attendees (they're separated by a colon) and put into an iterable list, and sequentially add them to table

    # split passed argument into a list we can iterate over
    mySplitUsers = []
    mySplitUsers = tmpAttendees.split("; ")

    # if more than zero list items
    if len(mySplitUsers) > 0:
      # iterate over list
      for y in mySplitUsers:
        # get name to lookup (making SQL safe - replacing any apostrophes with double)
        nameToLookUp = y.replace("'", "''")
        # get user code from users table
        userCode = runSQL("SELECT Code FROM Users WHERE FullName = '{0}'".format(nameToLookUp))
        
        # if we got something we can add to table
        if len(userCode) > 0:
          # add to attendees table
          aSQL = """INSERT INTO Diary_AppointmentsAttendees (AppointmentRef, StepItemID, Username, DeleteThis, UpdateThis, ExternalEmail, ExternalEntRef) 
                    VALUES({0}, {1}, '{2}', 0, 0, '', '')""".format(dateTblID, caseItemID, userCode)
          runSQL(aSQL)


  # remove completion date on Case items table
  runSQL("UPDATE Cm_CaseItems SET CompletionDate = NULL WHERE ItemID = {0}".format(caseItemID))

  # update due date in Steps table (after checking in tables, it is the Task Due Date we want here, not the reminder)
  runSQL("UPDATE Cm_Steps SET SentDate = NULL, Taker = NULL WHERE ItemID = {0}".format(caseItemID))

  # update Key Dates table (are there any other fields we need to update here)
  updateKD_SQL = """UPDATE Usr_Key_Dates SET AssignedTo = '{0}', Date = '{1}', DurationQty = {2}, DurationTypeN = {3}, 
                    Description = '{4}', ReminderQty = {5}, ReminderTypeN = {6}, DateMissedNotes = '{7}', Location = '{8}',
                    TaskOrDate = 'Date', DairyTblID = {9} WHERE ID = {10}""".format(feREF, fullDueDateAndTime, dDuration, dType, dDesc, dReminder, dRemType, 
                                                                                    dDateMissedN, dLocation, dateTblID, kdID)

  if debugMessage(msgTitle = 'DEBUG MESSAGE - Testing Date Revert', msgBody = updateKD_SQL) == True:
    runSQL(updateKD_SQL, True, "There was an error updating the KeyDates table", "ERROR - KeyDate Revert...")

  # finally refresh the key tasks datagrid
  refresh_KeyDates(s, event)

  # it would be nice to select this item again, as it's going to move to 'outstanding' group
  tCount = -1
  for tRow in dg_KeyDates.Items:
    tCount += 1
    if tRow.iCaseStepID == caseItemID:
      dg_KeyDates.SelectedIndex = tCount
      break
  return


def sql_safe_string(stringToClean = None):
  # This function accepts one argument, and will test if 'none' and if so, returns an empty string, otherwise, it will
  # use the 'replace' function to replace single quote with double quotes (safe for SQL calls)
  
  if stringToClean == None:
    tmpReturn = ''
  else:
    tmpReturn = stringToClean.replace("'", "''")
  
  # technically, the above could be made into one line, though I prefer above for readability:
  #tmpReturn = '' if stringToClean == None else stringToClean.replace("'", "''")
  return tmpReturn


def getUniqueDescription(desiredDesc = '', forUser = '', taskORdate = ''):
  # This function will iterate over the items in the Diary_* table to look for any matches to current desired description, 
  # and returns next available 'description' to avoid duplicates.
  # Previously, we added a little code in main 'global_AddTask()' function to append a number if task with passed description
  #  already exists, but then make no further check to see if THAT name has already been used.  In our tests, we 
  #  successfully added 'New Task - Edit Here', and then the second time it appended a '1', but the third time results 
  #  in previous '1' item being deleted and a new '1' is added (eg: it won't allow adding duplicates, so deletes older item)

  # setup initial variables
  canExit = False
  tmpNumber = 0
  newDesc = desiredDesc.replace("'", "''")
  diaryTable = 'Diary_Tasks' if taskORdate == 'Task' else 'Diary_Appointments'

  # initiate loop
  while canExit == False:
    # increment number variable
    tmpNumber += 1

    # if this is the first iteration, don't append number to description, and test if it has any matches
    if tmpNumber == 1:
      countMatches = runSQL("""SELECT COUNT(Code) FROM {0} WHERE EntityRef = '{1}' AND MatterNoRef = {2} 
                               AND Username = '{3}' AND Description LIKE '{4}%'""".format(diaryTable, _tikitEntity, _tikitMatter, forUser, newDesc))
      # if no matches, set canExit variable to true to exit loop
      if int(countMatches) == 0:
        canExit = True
    else:
      # this is not the first iteration - so append current number to description and test for any matches to that
      newDesc = "{0} {1}".format(desiredDesc.replace("'", "''"), tmpNumber)
      countMatches = runSQL("""SELECT COUNT(Code) FROM {0} WHERE EntityRef = '{1}' AND MatterNoRef = {2} 
                               AND Username = '{3}' AND Description LIKE '{4}%'""".format(diaryTable, _tikitEntity, _tikitMatter, forUser, newDesc))
      # if there are no matches, set canExit to True to exit loop
      if int(countMatches) == 0:
        canExit = True

  # return current 'newDesc'
  return newDesc

###################################################
# New 'Time' code for combo boxes instead of allowing users to enter text
class hoursCbo(object):
  def __init__(self, myCode):
    self.iCode = myCode
    return
  
  def __getitem__(self, index):
    if index == 'Code':
      return self.iCode

def get_TimeHours(startHour = 7, endHour = 19):
  # This function will return a list of 'Hours' that will be used for combo bow within DataGrids for 'Time'
  # You can specify the 'startHour' and the 'endHour' - this is to help limit the options shown as in theory
  # no one wants to be reminded AFTER they've finished for the day.  Currently allowing 7am to 7pm, but may want to change
  mHour = []
  nEndHour = endHour + 1
  tmpH = ""
  
  # loop from the startHour to the endHour (+1 [as above] because by default it excludes the last number)
  for myH in range(startHour, nEndHour):
    # if only one character long
    if len(str(myH)) == 1:
      # prefix a leading zero
      tmpH = "0{0}".format(myH)
    else:
      # there are 2 digits, so use as is
      tmpH = "{0}".format(myH)
    # finally, add this item to our list
    mHour.append(hoursCbo(tmpH))
  # return list to calling procedure
  return mHour

class minsCbo(object):
  def __init__(self, myCode):
    self.iCode = myCode
    return
  
  def __getitem__(self, index):
    if index == 'Code':
      return self.iCode

def get_TimeMins(increment = 1):
  # This function will return a list of 'minutes' that can be used as an item source for a combo box within a DataGrid
  # It takes one argument 'increment' and will provide minutes from 0 to 60 at the increments specified
  mMin = []
  tmpMin = ""
  
  # if increment is one
  if increment == 1:
    # just loop from 0 to 60
    for myM in range(0, 60):
      # if length of number is one character long
      if len(str(myM)) == 1:
        # prefix with a leading zero
        tmpMin = "0{0}".format(myM)
      else:
        # number is two characters long, so use as is
        tmpMin = "{0}".format(myM)
      # finally, append this number to list
      mMin.append(minsCbo(tmpMin))
      
  else:
    # an increment greater than one was specified, so loop from 0 to 60 at the increment given
    for myM in range(0, 60, increment):
      # if length of number is one character
      if len(str(myM)) == 1:
        # prefix with a leading zero
        tmpMin = "0{0}".format(myM)
      else:
        # number is 2 character long, so use as is
        tmpMin = "{0}".format(myM)
      # finally add number to list
      mMin.append(minsCbo(tmpMin))

  # return list to calling procedure
  return mMin

#def OnPreviewKeyDown(s, event):
 #if str(event.Key) == "Delete":
  #  deleteDate(s, event)
   # refresh_KeyDates(s, event)

    


]]>
    </Init>
    <Loaded>
      <![CDATA[
 
# Globals

#Define controls that will be used in all of the code
# ControlName = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ControlName')
tc_Main = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tc_Main')


# DEBUG MODE - CONTROLS #
chk_DebugMode = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'chk_DebugMode')
stk_DebugMode = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'stk_DebugMode')
txt_DebugModeOutput = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'txt_DebugModeOutput')
lbl_CHAgendaID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_CHAgendaID')
lbl_CurrentDept = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_CurrentDept')

# NEW Version - Dates
# Tab Item
ti_Dates = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_Dates')
# Buttons
btn_AddAllDefaultDates = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_AddAllDefaultDates')
btn_AddAllDefaultDates.Click += date_AddAllDefaults
btn_AddNew_Date = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_AddNew_Date')
btn_AddNew_Date.Click += date_AddNew
btn_MarkAsComplete_Date = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_MarkAsComplete_Date')
btn_MarkAsComplete_Date.Click += date_MarkComplete
btn_RevertDate = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_RevertDate')
btn_RevertDate.Click += revertDate
btn_DeleteDate = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_DeleteDate')
btn_DeleteDate.Click += deleteDate
# DataGrid
dg_KeyDates = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_KeyDates')
dg_KeyDates.SelectionChanged += dg_KeyDates_SelectionChanged
dg_KeyDates.CellEditEnding += dg_KeyDates_cellEdit_Finished
#dg_KeyDates.SelectedCellsChanged += dg_KeyDates_cellSelectionChanged
dg_DD_Attendees = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_DD_Attendees')
dg_DD_Attendees.CellEditEnding += dateAttendee_cellUpdated

# Text boxes / Combo boxes / other controls
txt_DD_Desc = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'txt_DD_Desc')
dp_DD_Date = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dp_DD_Date')
txt_DD_Time = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'txt_DD_Time')
txt_DD_Time.LostFocus += validate_DateDueDateTime
txt_DD_Duration = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'txt_DD_Duration')
txt_DD_Duration.LostFocus += validate_DateDuration
cbo_DD_DurType = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cbo_DD_DurType')
txt_DD_ReminderQty = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'txt_DD_ReminderQty')
txt_DD_ReminderQty.LostFocus += validate_DateReminder
cbo_DD_RemType = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cbo_DD_RemType')
txt_AttendeeSearch = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'txt_AttendeeSearch')
txt_AttendeeSearch.TextChanged += find_DateAttendee
chk_DD_OnlyShowDeptUsers = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'chk_DD_OnlyShowDeptUsers')
chk_DD_OnlyShowDeptUsers.Click += populateAttendeesList
txt_DD_Location = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'txt_DD_Location')
txt_DateMissedNotes = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'txt_DateMissedNotes')
stk_DatesMissedNotes = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'stk_DatesMissedNotes')

lbl_DateMissedNotes = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_DateMissedNotes')
lbl_CaseStepID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_CaseStepID')
lbl_DAcode = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_DAcode')

cbo_Date_AssignedTo = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cbo_Date_AssignedTo')
btn_SetDateAssignee_MatterFE = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_SetDateAssignee_MatterFE')
btn_SetDateAssignee_MatterFE.Click += assignDate_toMatterFeeEarner
btn_SetDateAssignee_CurrUser = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_SetDateAssignee_CurrUser')
btn_SetDateAssignee_CurrUser.Click += addignDate_toCurrentUser

opt_DD_EditSelected = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'opt_DD_EditSelected')
opt_DD_EditSelected.Checked += dd_EditSelected_Click
opt_DD_AddNew = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'opt_DD_AddNew')
opt_DD_AddNew.Checked += dd_AddNew_Click

grd_MainDates = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'grd_MainDates')
col_DatesDG = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'col_DatesDG')
exp_DateAttendees = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'exp_DateAttendees')
exp_DateAttendees.Expanded += expand_DateAttendees
exp_DateAttendees.Collapsed += contract_DateAttendees


## New Defaults
ti_Defaults = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_Defaults')
defaults_label = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'defaults_label')
defaults_label.MouseLeftButtonDown += refresh_KDDefaults_List
# NO NO NO - this perpetually refreshes and doesn't allow anything to be selected
btn_AddDefaultsToDates = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_AddDefaultsToDates')
btn_AddDefaultsToDates.Click += addDefaults_ToDiaryDates
btn_AddDefaultsToTasks = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_AddDefaultsToTasks')
btn_AddDefaultsToTasks.Click += addDefaults_ToTasks
btn_defaults_Refresh = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_defaults_Refresh')
btn_defaults_Refresh.Click += refresh_KDDefaults_List
dg_DateDefaults = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_DateDefaults')
btn_DefaultsSelectAll = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_DefaultsSelectAll')
btn_DefaultsSelectAll.Click += defaultDiaryDates_SelectAllNone
tb_DefaultsSelectAll = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tb_DefaultsSelectAll')
lbl_NoDefaults = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_NoDefaults')
lbl_DeptDefaultType = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_DeptDefaultType')


# NEW Version - Tasks
ti_Tasks = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_Tasks')
btn_AddAllDefaultTasks = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_AddAllDefaultTasks')
btn_AddAllDefaultTasks.Click += task_AddAllDefaults
btn_AddNew_Task = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_AddNew_Task')
btn_AddNew_Task.Click += task_AddNew
btn_MarkAsComplete_Task = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_MarkAsComplete_Task')
btn_MarkAsComplete_Task.Click += task_MarkComplete
dg_KeyTasks = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_KeyTasks')
dg_KeyTasks.SelectionChanged += task_cellSelection_Changed
dg_KeyTasks.CellEditEnding += dg_KeyTasks_cellEdit_Finished
lbl_TaskRowID = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_TaskRowID')

opt_Task_EditSelected = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'opt_Task_EditSelected')
opt_Task_EditSelected.Checked += task_optEditSelected_Clicked
opt_Task_AddNew = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'opt_Task_AddNew')
opt_Task_AddNew.Checked += task_optAddNew_Clicked

cbo_Task_AssignedTo = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cbo_Task_AssignedTo')
btn_SetTaskAssignee_MatterFE = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_SetTaskAssignee_MatterFE')
btn_SetTaskAssignee_MatterFE.Click += task_SetToFeeEarner
btn_SetTaskAssignee_CurrUser = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_SetTaskAssignee_CurrUser')
btn_SetTaskAssignee_CurrUser.Click += task_SetToCurrentUser

dp_Task_DateDue = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dp_Task_DateDue')
dp_Task_DateReminder = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dp_Task_DateReminder')
txt_Task_TimeReminder = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'txt_Task_TimeReminder')
txt_TaskDescription = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'txt_TaskDescription')
cbo_Task_Status = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cbo_Task_Status')
cbo_Task_Priority = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cbo_Task_Priority')
txt_Task_PercentComplete = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'txt_Task_PercentComplete')
txt_DateMissedNotes_Task = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'txt_DateMissedNotes_Task')
stk_TaskMissedNotes = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'stk_TaskMissedNotes')

cbo_TaskPostpone = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cbo_TaskPostpone')
btn_taskPostpone = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_taskPostpone')
btn_taskPostpone.Click += task_PostponeNow
grp_PostponeOptions = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'grp_PostponeOptions')
#tSep1 = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'tSep1')
btn_DeleteTask = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_DeleteTask')
btn_DeleteTask.Click += deleteTask
btn_RevertTask = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_RevertTask')
btn_RevertTask.Click += revertTask

# Quick dirty testing of 'context menu' - doesn't appear to work like 'buttons' as we get error on load 'NoneType has no 'click' attribute'
#btn_MP_contextMenu_test = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_MP_contextMenu_test')
#btn_MP_contextMenu_test.Click += contextMenuTest_Button1Click
#btn_MP_contextMenu_test2 = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_MP_contextMenu_test2')
#btn_MP_contextMenu_test2.Click += contextMenuTest_Button2Click

# Auto run actions when this form loads...
myOnLoadEvent(_tikitSender, 'onLoad')



]]>
    </Loaded>
  </KeyDates>

</tfb>