/// <reference path="../App.js" />

(function () {
    "use strict";
    var cJob = "Jobs";
    var cTask = "Tasks";
    var first = true;

    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason)
    {
        $(document).ready(function () {
            app.initialize();

            var sender_email = getEmail(); 

            runApp(sender_email); // Main function

            // Event listener functions for clickable buttons.
            $("#uploadNote").click(uploadNote);
            $("#uploadTimesheet").click(uploadTimesheet);
            $("#uploadAttachment").click(uploadAttachment);
            $("#Jobs3").on('click', function ()
            {
                if(!first)
                document.getElementById("test").style.display = "none";
            });
            $(document).on('keydown', function (e)
            {
                if (e.keyCode == 65)
                {
                    var a = document.getElementById("Tasks2");
                    a.parentNode.removeChild(a);
                    //app.showNotification("blah");
                }
            });

            $("ul#Jobs").on('click', 'li', function ()
            {
                app.showNotification("hello");
                document.getElementById("test").style.display = "";
                printTasks($(this));
                
            });

            $("ul#Tasks").on('click', 'li', function ()
            {
                app.showNotification("hello");
                cTask = $(this).attr('data-value');
                
            });
        });
    };

    // Function to get the sender email.
    function getEmail()
    {
        var item = Office.context.mailbox.item.from;

        var email = item.emailAddress;

        return email;
    }

    function runApp(email)
    {
        var id = "Jobs";

        var staffID = getStaffID();

        // Prints jobs that are assigned to you.
        printJobs(staffID);

        makePretty(document.getElementById("selectJobs").id, id); // Makes the fancy looking job list.
    }

    // Uploads the attachment of the email if there is one.
    function uploadAttachment()
    {
        if (Office.context.mailbox.item.attachments == undefined)
        {
            app.showNotification("Sorry attachments are not supported by your Exchange server.");
        }
        else if (Office.context.mailbox.item.attachments.length == 0)
        {
            app.showNotification("Oops there are no attachments on this email.");
        }
        else
        {
            var apicall = "https://api.workflowmax.com/job.api/document?apiKey=14C10292983D48CE86E1AA1FE0F8DDFE&accountKey=8A39F28D022B4366975D6FCDB180C839";

            var documentXML = "<Document><Job>" + cJob + "</Job><Title>Document Title</Title><Text>Note for document</Text><FileName>test.txt</FileName><Content>" + string64 + "</Content></Document>";

            var xhr = new XMLHttpRequest();

            xhr.open('POST', apicall);

            xhr.send(documentXML);
        }
    }

    // Uploads the content of the email
    function uploadNote()
    {
        // Get the content of email. 
        var item = Office.context.mailbox.item.body.getAsync("text", callback);
    }

    function callback(asyncResult)
    {
        var notetext = asyncResult.value;

        var a = notetext.toString();

        //app.showNotification(notetext);

        var apicall = "https://api.workflowmax.com/job.api/note?apiKey=14C10292983D48CE86E1AA1FE0F8DDFE&accountKey=8A39F28D022B4366975D6FCDB180C839";

        var noteXML = "<Note><Job>" + cJob + "</Job><Title>Email content</Title><Text>" + notetext + "</Text></Note>"; // XML representing the note.

        var xhr = new XMLHttpRequest(); // Create a new XMLHTTPRequest

        xhr.open('POST', apicall); 

        xhr.send(noteXML); // Send the note to workflowmax via XMLHttpRequest
        
    }

    // Adds a timesheet entry.
    function uploadTimesheet()
    {
        var task = getTask(cJob);
    }

    function getTask(jobID)
    {
        var taskname = "Consulting - Email Processing";
        var foundtask = false;

       var apicall = "https://api.workflowmax.com/job.api/get/" + jobID + "?apiKey=14C10292983D48CE86E1AA1FE0F8DDFE&accountKey=8A39F28D022B4366975D6FCDB180C839";

        var jobDetails = getXML(apicall);

        var tasklist = jobDetails.getElementsByTagName("Task");

        

        for (var i = 0; i < tasklist.length; i++)
        {
            var thistask = tasklist[i].getElementsByTagName("Name")[0].childNodes[0].nodeValue;
            var taskID = tasklist[i].getElementsByTagName("ID")[0].childNodes[0].nodeValue;
            
            if(thistask == taskname)
            {
                updateTask(jobID, taskID);
                foundtask = true;
                break;
            }
        }

        // If the email processing task does not exist; create one and add 15 minutes to it.
        if (!foundtask)
        {
            createTask(jobID);
            var id = getTaskID(jobID);
            updateTask(jobID, id);
        }

        return tasklist;
    }

    function getStaffID()
    { 
        // Get email of the user.
        var email = Office.context.mailbox.userProfile.emailAddress;

        var apicall = "https://api.workflowmax.com/staff.api/list?apiKey=14C10292983D48CE86E1AA1FE0F8DDFE&accountKey=8A39F28D022B4366975D6FCDB180C839";

        // Get the list of all staff
        var staffdetails = getXML(apicall);

        var stafflist = staffdetails.getElementsByTagName("Staff");

        // Find the staff 
        for(var i = 0; i < stafflist.length; i++)
        {
            var tempEmail = stafflist[i].getElementsByTagName("Email")[0].childNodes[0].nodeValue;
            var tempID = stafflist[i].getElementsByTagName("ID")[0].childNodes[0].nodeValue;

            if(tempEmail == email)
            {
                return tempID; // Returns the staff ID if found.
            }
        }

        return null; // Return null if the staff is not found (this user is not a staff member).
    }

    // Function to add 15 minutes to the email processing task.
    function updateTask(jobID, taskID)
    {
        var noteText = document.getElementById("Note").value;

        var apicall = "https://api.workflowmax.com/time.api/add?apiKey=14C10292983D48CE86E1AA1FE0F8DDFE&accountKey=8A39F28D022B4366975D6FCDB180C839";
        var tsxml = "<Timesheet><Job>" + jobID + "</Job><Task>" + taskID + "</Task><Staff>" + getStaffID() + "</Staff><Date>" + getDate() + "</Date><Minutes>15</Minutes><Note>" + noteText + "</Note></Timesheet>";

        var xhr = new XMLHttpRequest();

        xhr.open('POST', apicall);

        xhr.send(tsxml);
    }

    function getTaskID(jobID)
    {
        var taskname = "Consulting - Email Processing";
        var apicall = "https://api.workflowmax.com/job.api/get/" + jobID + "?apiKey=14C10292983D48CE86E1AA1FE0F8DDFE&accountKey=8A39F28D022B4366975D6FCDB180C839";

        var jobDetails = getXML(apicall);

        var tasklist = jobDetails.getElementsByTagName("Task");

        for (var i = 0; i <= tasklist.length; i++)
        {
            //document.getElementById("Email").innerHTML = "<b>Email : Hello </b>";
            
            var thistask = tasklist[i].getElementsByTagName("Name")[0].childNodes[0].nodeValue;            
            var taskID = tasklist[i].getElementsByTagName("ID")[0].childNodes[0].nodeValue;
            if (taskID == null)
            {
                
            }

            if (thistask == taskname)
            {
                
                return taskID;
            }
        }
    }

    // Creates the email processing task.
    function createTask(jobID)
    {
        var taskID = "1772154";
        var label = "Email Processing";

        var apicall = "https://api.workflowmax.com/job.api/task?apiKey=14C10292983D48CE86E1AA1FE0F8DDFE&accountKey=8A39F28D022B4366975D6FCDB180C839";

         var taskXML = "<Task><Job>" + jobID + "</Job><TaskID>" + taskID + "</TaskID><Label>" + label + "</Label><EstimatedMinutes>300</EstimatedMinutes></Task>";

        var xhr = new XMLHttpRequest();

        xhr.open('POST', apicall);
        xhr.send(taskXML);
    }


    // Returns an xml document from given api call; used for GET requests.
    function getXML(list)
    {
        var thisXMLhttp = new XMLHttpRequest();
        thisXMLhttp.open("GET", list, false);
        thisXMLhttp.send();
        var thisXMLDoc = thisXMLhttp.responseXML;

        return thisXMLDoc;
    }


    // Prints all the jobs assigned to the user of the application.
    function printJobs(staffID)
    {
        var dropdown = document.getElementById("selectJobs");

        // Gets the list of jobs assigned to this staff member
        var jobList = "https://api.workflowmax.com/job.api/staff/" + staffID + "?apiKey=14C10292983D48CE86E1AA1FE0F8DDFE&accountKey=8A39F28D022B4366975D6FCDB180C839";

        var jobsXML = getXML(jobList);

        var numJobs = jobsXML.getElementsByTagName("Job");

        for (var i = 0; i < numJobs.length; i++)
        {
            var tempJobID = numJobs[i].getElementsByTagName("ID")[0].childNodes[0].nodeValue;
            var tempJobName = numJobs[i].getElementsByTagName("Name")[0].childNodes[0].nodeValue;
            //var tempClientID = numJobs[i].getElementsByTagName("ID")[1].childNodes[0].nodeValue;

            $('#selectJobs').append('<option>' + tempJobID + '-' + tempJobName + '</option>');
        }
    }

    function printTasks(job)
    {
        var name = job.attr('data-value');

        if (name != "Jobs")
        {
            // Get the job id part of the string (first 8 characters)
            cJob = name.substring(0, 7);

            if (first)
            {
                makeCircular(document.getElementById("circular").id, "circle"); // Create the first 'hamburger' icon next to job list.
            }
            else
            {
                $('#test').empty();
                $('#test').append('<select id="selectTasks" class="cs-select cs-skin-slide" hidden="hidden" disabled="disabled"><option selected>Tasks</option></select>' +
                                   "<select class='cs-select cs-skin-circular' id='circular2' hidden='hidden' disabled='disabled'" +
                                   "<option value='' disabled selected>Select an activity</option>" +
                                   "<option value='1'>&#57605;</option>");
            }

            //$('#selectTasks').append('<option selected>Tasks</option>'); // Add default selection.

            /*$("#Tasks3").text("Tasks");

            if (!first)
            {
                $("#Tasks").empty();
                $('#Tasks').append('<li data-option data-value="Tasks" class="cs-selected"><span>Tasks</span></li>');
                     
            }*/

            var id = "Tasks";

            var apicall = "https://api.workflowmax.com/job.api/get/" + cJob + "?apiKey=14C10292983D48CE86E1AA1FE0F8DDFE&accountKey=8A39F28D022B4366975D6FCDB180C839";

            var jobdetails = getXML(apicall);

            var numTasks = jobdetails.getElementsByTagName("Task");

            for (var i = 0; i < numTasks.length; i++)
            {
                var tempTaskName = numTasks[i].getElementsByTagName("Name")[0].childNodes[0].nodeValue;

                $('#selectTasks').append('<option>' + tempTaskName + '</option>'); // Append the current job's list of tasks.
            }
                makePretty(document.getElementById("selectTasks").id, id); // Create the fancy looking task list.
                makeCircular(document.getElementById("circular2").id, "circle2"); // Create the second 'hamburger' icon next to task list.
        }

        first = false;
    }

    // Evaluate which button was pressed.
    function processAction(val)
    {
        app.showNotification(cJob);
        switch (val)
        {
            case '1':
                //uploadAttachment();
                break;
            case '2':
                uploadNote();
                break;
            case '3':
                uploadTimesheet();
                break;
            default:
                break;
        }
    }

    // Function to return date in the form YYYYMMDD, to conform to WorkflowMax format.
    function getDate()
    {
        var date = new Date();

        var month = date.getMonth()+1;
        var day = date.getDate();

        if (date.getMonth() < 10)
        {
            month = "0" + month;
        }

        if (date.getDate() < 10)
        {
            day = "0" + day;
        }

        var datestring = "" + date.getFullYear() + month + day;

        return datestring;
    }

    // Function to make the standard list of items.
    function makePretty(id, ulid)
    {
        [].slice.call(document.querySelectorAll('#' + id)).forEach(function (el)
        {
            new SelectFx(el, ulid);
        });
    }

    // Function to make the 'hamburger' circular select icon
    function makeCircular(id, ulid)
    {
        [].slice.call(document.querySelectorAll('#' + id)).forEach(function (el)
        {
            new SelectFx(el, ulid, {
                stickyPlaceholder: true,
                onChange: function (val)
                {
                    processAction(val);
                }
            });
        });
    }

    /*Deprecated/unneeded functions*/
    
     /*
    function getID(passedXML, emailString)
    {
        var clientNum = passedXML.getElementsByTagName("Client"); 

        for (var i = 0; i < clientNum.length; i++)
        {
            try
            {
                var thisID = clientNum[i].getElementsByTagName("ID")[0].childNodes[0].nodeValue;
                var thisName = clientNum[i].getElementsByTagName("Name")[0].childNodes[0].nodeValue;
                var thisEmail = clientNum[i].getElementsByTagName("Email")[0].childNodes[0].nodeValue;

                if (thisEmail == emailString)
                {
                    document.getElementById("Company").innerHTML = "<b>Company Name : </b>" + thisName;

                    return thisID;
                }
                else
                {
                    document.getElementById("Company").innerHTML = "<b>Company Name : </b>This customer is not set up in Workflow Max"
                }
            }
            catch (err)
            {
                thisEmail = "null";
            }
        }
    }*/

    /*function selectJobs()
    {
        var jobs = document.getElementById("selectJobs").options.selectedIndex;
        var currentJob = document.getElementById("selectJobs").options[jobs].text;

        document.write(currentJob);

        if(currentJob != "Jobs")
        {
            return currentJob;
        }
        else
        {
            return null;
        }
    }*/
})();