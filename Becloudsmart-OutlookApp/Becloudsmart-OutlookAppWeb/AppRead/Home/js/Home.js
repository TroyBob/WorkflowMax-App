/// <reference path="../App.js" />

(function () {
    "use strict";
    var cJobID = "Jobs"; // Currently selected jobID.
    var cTaskID = ""; // Currently selected taskID.
    var cTaskName = "Tasks" // Currently selected task name.
    var first = true; // True when the user first selects a job.
    var staffID = ""; // StaffID of the user.
    var expanded = false;
    var selectedIcon = "home"; //ID of the selected icon.

    var apiKey = "14C10292983D48CE86E1AA1FE0F8DDFE";    // API key for calls to WorkflowMax API.
    var accountKey = "8A39F28D022B4366975D6FCDB180C839"; // Account key for calls to WorkflowMax API.

    var clients;
    var jobs;
    var tasks;

    var accordionToggles;


    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason)
    {

        $(document).ready(function () {
            app.initialize();

            //var sender_email = getEmail(); 

            runApp(); // Main function

            // Event listener functions.
            $(".icons").on('click', function ()
            {
                iconSelect($(this));
            });

            $("#keySubmitButton").on('click', updateKeys);
            $("#emailSubmitButton").on('click', function ()
            {
                
            });

            $("ul#Clients").on('click', 'li', displayClientInfo);
            $("ul#Jobs").on('click', 'li', displayJobInfo);
            /*$("#Jobs3").on('click', function ()
            {    
                if (!first)
                {
                    document.getElementById("test").style.display = "none"; // Hides the task list whilst selecting a job.
                    document.getElementById("note").style.display = "none"; // Hides the textbox whilst selecting a job.
                }
            });

            $("ul#Jobs").on('click', 'li', function ()
            {
                
                document.getElementById("test").style.display = ""; // Redisplays the task list after a job is selected.
                document.getElementById("note").style.display = ""; // Redisplays the textbox after a job is selected.
                printTasks($(this));   
            });*/
        });
    };

    // Main function.
    function runApp()
    {
        clients = new Array();
        jobs = new Array();
        accordionToggles = document.querySelectorAll(".js-accordionTrigger");

        for (var i = 0; i < accordionToggles.length; i++)
        {
            accordionToggles[i].addEventListener('click', switchAccordion, false);
        }

        init();
       
        displayHomeData();
    }

    function displayClientInfo(e)
    {
        var id = $(this).attr('data-value');
    }

    function displayJobInfo(e)
    {
        var id = $(this).attr('data-value');

        cJobID = jobs[id].id;

        fillTaskList();
    }

    function switchAccordion(e)
    {
        e.preventDefault();
        var thisAnswer = e.target.parentNode.nextElementSibling;
        var thisQuestion = e.target;

        if(thisAnswer.classList.contains('is-collapsed'))
        {
            setAccordionAria(thisQuestion, thisAnswer, 'true');
        }
        else
        {
            setAccordionAria(thisQuestion, thisAnswer, 'false');
        }
        thisQuestion.classList.toggle('is-collapsed');
        thisQuestion.classList.toggle('is-expanded');
        thisAnswer.classList.toggle('is-collapsed');
        thisAnswer.classList.toggle('is-expanded');

        thisAnswer.classList.toggle('animateIn');
    }

    function setAccordionAria(el1,el2, expanded)
    {
        switch(expanded)
        {
            case "true":
                el1.setAttribute('aria-expanded', 'true');
                el2.setAttribute('aria-hidden', 'false');
                break;
            case "false":
                el1.setAttribute('aria-expanded', 'false');
                el2.setAttribute('aria-hidden', 'true');
                break;
            default:
                break;
        }
    }

    function init()
    {
        staffID = getStaffID();
        fillClientList();
        fillJobList();
    }

    function updateKeys()
    {
        apiKey = $("#apikey").val();
        accountKey = $("#accountkey").val();
    }

    function iconSelect(icon)
    {
        if (icon.attr('id') == "arrow")
        {
            expandMenu();
        }
        else
        {
            if (!icon.hasClass("selected"))
            {
                $("#" + selectedIcon).removeClass("selected");
                $("#" + selectedIcon + "Data").hide();

                icon.addClass("selected");
                selectedIcon = icon.attr('id'); // Update the currently selected icon.
                
                $("#" + selectedIcon + "Data").show();

                switch (icon.attr('id'))
                {
                    case 'home':
                        displayHomeData();
                        break;
                    case 'client':
                        displayClients();
                        break;
                    case 'tickbox':
                        displayJobs();
                        break;
                    case 'settings':
                        displaySettings();
                        break;
                    case 'faq':
                        displayFAQ();
                        break;
                    case 'support':
                        displaySupport();
                        break;
                    default:
                        break;
                }
            }
        }
    }

    function displayHomeData()
    {
        $("#numClients").text(clients.length);
        $("#totalClients").text("Total Clients");

        $("#numJobs").text(jobs.length);
        $("#totalJobs").text("Total Jobs");
    }

    function displayClients()
    {

    }

    function displayJobs()
    {

    }

    function displaySettings()
    {

    }

    function displayFAQ()
    {

    }

    displaySupport()
    {
        var email = "troy@becloudsmart.com";
        var subject = "Support-BecloudsmartApp";
    }

    // Function to handle expansion of menu
    function expandMenu()
    {
        if (!expanded)
        {
            $(".left").width("8%");
            $(".right").width("92%");
            $("#arrow span").html("&#57481;");
            $(".icons").css("text-align", "left");
            document.getElementById("arrow").style.textAlign = "right";

            addText();

            expanded = true;
        }
        else
        {
            $(".left").width("3%");
            $(".right").width("97%");
            $("#arrow span").html("&#57482;");
            $(".icons").css("text-align", "center");
            document.getElementById("arrow").style.textAlign = "left";

            removeText();

            expanded = false;
        }
    }

    function removeText()
    {
        $("#home span").html("&#57377;");   
        $("#client span").html("&#57358;");
        $("#tickbox span").html("&#57408;");
        $("#settings span").html("&#57397;");
        $("#faq span").html("&#57350;");
        $("#support span").html("&#57398;");
    }

    function addText()
    {
        $("#home span").append(" Home");
        $("#client span").append(" Clients");
        $("#tickbox span").append(" Tasks");
        $("#settings span").append(" Settings");
        $("#faq span").append(" FAQ");
        $("#support span").append(" Support");
    }

    // Function to get the sender email.
    function getEmail()
    {
        var item = Office.context.mailbox.item.from;

        var email = item.emailAddress;

        return email;
    }

    function fillClientList()
    {
        
        var apicall = "https://api.workflowmax.com/client.api/list?apiKey=" + apiKey + "&accountKey=" + accountKey;
        
        var clientXML = getXML(apicall);

        var clientlist = clientXML.getElementsByTagName("Client");

        /*var temp = clientlist[2].getElementsByTagName("Email")[0];

        if (temp == null)
        {
            app.showNotification("Test!");
        }*/

        for (var i = 0; i < clientlist.length; i++)
        {
            var tempEmail = clientlist[i].getElementsByTagName("Email")[0];
            var tempPhone = clientlist[i].getElementsByTagName("Phone")[0];
            var tempWebsite = clientlist[i].getElementsByTagName("Website")[0];

            clients[i] = new Object();

            if (tempEmail != null)
            {
                clients[i].email = tempEmail.childNodes[0].nodeValue;
            }
            else
            {
                clients[i].email = null;
            }

            if (tempPhone != null)
            {
                //clients[i].phone = tempPhone.childNodes[0].nodeValue;
            }
            else
            {
                clients[i].phone = null;
            }

            if (tempWebsite != null)
            {
                //clients[i].website = tempWebsite.childNodes[0].nodeValue;
            }
            else
            {
                clients[i].website = null;
            }

            clients[i].id = clientlist[i].getElementsByTagName("ID")[0].childNodes[0].nodeValue;
            clients[i].name = clientlist[i].getElementsByTagName("Name")[0].childNodes[0].nodeValue;

            $("#clientList").append("<option value='" + i + "'>" + clients[i].name + "</option>");
        }
        makeJobList("clientList", "Clients");
    }

    function fillJobList()
    {
        var jobList = "https://api.workflowmax.com/job.api/staff/" + staffID + "?apiKey=" + apiKey + "&accountKey=" + accountKey;

        var jobsXML = getXML(jobList);

        var joblist = jobsXML.getElementsByTagName("Job");

        for (var i = 0; i < joblist.length; i++)
        {
            jobs[i] = new Object();
            
            jobs[i].id = joblist[i].getElementsByTagName("ID")[0].childNodes[0].nodeValue;
            jobs[i].name = joblist[i].getElementsByTagName("Name")[0].childNodes[0].nodeValue;

            $("#jobList").append("<option value='" + i + "'>" + jobs[i].id + "-" + jobs[i].name + "</option>");

        }

        makeJobList("jobList", "Jobs");
    }

    function fillTaskList()
    {

        $("#taskDiv").empty();
        $("#taskDiv").append('<label> Tasks: <select id="taskList" class="cs-select cs-skin-slide"></select> </label>'
                               + '<span class="icons">Hello</span>');
        
        var apicall = "https://api.workflowmax.com/job.api/get/" + cJobID + "?apiKey=" + apiKey + "&accountKey=" + accountKey;
        app.showNotification("test");

        var jobdetails = getXML(apicall);

        var numTasks = jobdetails.getElementsByTagName("Task");

        for (var i = 0; i < numTasks.length; i++)
        {
            var tempTaskID = numTasks[i].getElementsByTagName("ID")[0].childNodes[0].nodeValue;
            var tempTaskName = numTasks[i].getElementsByTagName("Name")[0].childNodes[0].nodeValue;

            $('#taskList').append('<option value="' + tempTaskID + '">' + tempTaskName + '</option>'); // Append the current job's list of tasks.
        }

        makeTaskList("taskList", "Tasks");

        $("#taskDiv").show();
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
            var apicall = "https://api.workflowmax.com/job.api/document?apiKey="+ apiKey + "&accountKey=" + accountKey;

            var documentXML = "<Document><Job>" + cJobID + "</Job><Title>Document Title</Title><Text>Note for document</Text><FileName>test.txt</FileName><Content>" + string64 + "</Content></Document>";

            var xhr = new XMLHttpRequest();

            xhr.open('POST', apicall);

            xhr.send(documentXML);
        }
    }

    // Uploads the content of the email
    function uploadNote()
    {
        // Get the content of email and then calls the 'callback' function.
        var item = Office.context.mailbox.item.body.getAsync("text", callback);
    }

    function callback(asyncResult)
    {
        // Get the actual text from the body of the email.
        var notetext = asyncResult.value;

        var apicall = "https://api.workflowmax.com/job.api/note?apiKey=" + apiKey + "&accountKey=" + accountKey;

        var noteXML = "<Note><Job>" + cJobID + "</Job><Title>Email content</Title><Text>" + notetext + "</Text></Note>"; // XML representing the note.

        var xhr = new XMLHttpRequest(); // Create a new XMLHTTPRequest

        xhr.open('POST', apicall, false);

        xhr.send(noteXML); // Send the note to workflowmax via XMLHttpRequest

        

        // Check HttpRequest status.
        if(xhr.status == 200)
        {
            // Successful status.
            app.showNotification("Uploading of email content successful! :)");
        }
        else if(xhr.status == 500)
        {
            // An error occurred.
            app.showNotification("Email content unsuccessfully uploaded - incompatible format. :(");
        }  
    }

    // Adds a timesheet entry.
    function uploadTimesheet()
    {
        if (cTaskName != "Task")
        {      
            var apicall = "https://api.workflowmax.com/time.api/add?apiKey=" + apiKey + "&accountKey=" + accountKey;
            //app.showNotification("test1");
            var tsxml = "<Timesheet><Job>" + cJobID + "</Job><Task>" + cTaskID + "</Task><Staff>" + staffID + "</Staff><Date>" + getDate() + "</Date><Minutes>" + $('#time :selected').text() + "</Minutes><Note>" + $('#timesheetNote').val() + "</Note></Timesheet>";
            
            var xhr = new XMLHttpRequest();

            xhr.open('POST', apicall, false);

            xhr.send(tsxml);

            // Check HttpRequest status.
            if(xhr.status == 200)
            {
                // Successful status.
                app.showNotification($('#time :selected').text() + " minutes added to " + cTaskName);
            }
            else if(xhr.status == 500)
            {
                // An error occurred.
                app.showNotification("Error in modifying timesheet of task: " + cTaskName);
            }
        }
    }

    function getStaffID()
    { 
        // Get email of the user.
        var email = Office.context.mailbox.userProfile.emailAddress;

        var apicall = "https://api.workflowmax.com/staff.api/list?apiKey=" + apiKey + "&accountKey=" + accountKey;

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
    function printJobs()
    {
        var dropdown = document.getElementById("selectJobs");

        // Gets the list of jobs assigned to this staff member
        var jobList = "https://api.workflowmax.com/job.api/staff/" + staffID + "?apiKey=" + apiKey + "&accountKey=" + accountKey;

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
            cJobID = name.substring(0, 7);

            if (first)
            {
                makeHamburgerIcon(document.getElementById("circular").id, "circle"); // Create the first 'hamburger' icon next to job list.
            }
            else
            {
                $('#test').empty();
                $('#test').append('<select id="selectTasks" class="cs-select cs-skin-slide" hidden="hidden" disabled="disabled"><option selected>Tasks</option></select>' +
                                   "<span class='icons'>&#57605;</span>");
                                
            }

            //$('#selectTasks').append('<option selected>Tasks</option>'); // Add default selection.

            /*$("#Tasks3").text("Tasks");

            if (!first)
            {
                $("#Tasks").empty();
                $('#Tasks').append('<li data-option data-value="Tasks" class="cs-selected"><span>Tasks</span></li>');
                     
            }*/

            var id = "Tasks";

            var apicall = "https://api.workflowmax.com/job.api/get/" + cJobID + "?apiKey=" + apikey + "&accountKey" + accountKey;

            var jobdetails = getXML(apicall);

            var numTasks = jobdetails.getElementsByTagName("Task");

            for (var i = 0; i < numTasks.length; i++)
            {
                var tempTaskID = numTasks[i].getElementsByTagName("ID")[0].childNodes[0].nodeValue;
                var tempTaskName = numTasks[i].getElementsByTagName("Name")[0].childNodes[0].nodeValue;

                $('#selectTasks').append('<option value="' + tempTaskID + '">' + tempTaskName + '</option>'); // Append the current job's list of tasks.
            }
                makeTaskList(document.getElementById("selectTasks").id, id); // Create the fancy looking task list.
                //makeHamburgerIcon(document.getElementById("circular2").id, "circle2"); // Create the second 'hamburger' icon next to task list.
        }


        $('#Tasks3').on('click', function ()
        {
            document.getElementById("note").style.display = "none";
        });

        $("ul#Tasks").on('click', 'li', function ()
        {
            cTaskName = $(this).find('span').text(); // Set the current task name.
            document.getElementById("note").style.display = ""; // Redisplay the textbox.
        });

        first = false;
    }

    // Evaluate which button was pressed.
    function processAction(val)
    {
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
    function makeJobList(id, ulid)
    {
        [].slice.call(document.querySelectorAll('#' + id)).forEach(function (el)
        {
            new SelectFx(el, ulid);
        });
    }

    function makeTaskList(id, ulid)
    {
        [].slice.call(document.querySelectorAll('#' + id)).forEach(function (el)
        {
            new SelectFx(el, ulid, {
                stickyPlaceholder: true,
                onChange: function (val)
                {
                    cTaskID = val;
                }
            });
        });
    }

    // Function to make the 'hamburger' circular select icon
    function makeHamburgerIcon(id, ulid)
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