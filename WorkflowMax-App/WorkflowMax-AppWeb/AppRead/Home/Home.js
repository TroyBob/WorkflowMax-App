/// <reference path="../App.js" />

(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason)
    {
        $(document).ready(function () {
            app.initialize();

            var sender_email = getEmail();

            runApp(sender_email);

            $("#uploadNote").click(uploadNote);
            $("#uploadTimesheet").click(uploadTimesheet);

        });
    };

    function getEmail()
    {
        var item = Office.context.mailbox.item.from;

        var email = item.emailAddress;

        return email;
    }

    function runApp(email)
    {
        var list = "https://api.workflowmax.com/client.api/list?apiKey=14C10292983D48CE86E1AA1FE0F8DDFE&accountKey=8A39F28D022B4366975D6FCDB180C839";

        var xmlDoc = getXML(list);

        var returnID = getID(xmlDoc, email);

        document.getElementById("Email").innerHTML = "<b>Email : </b>" + email;

        document.getElementById("Current").innerHTML = "<b>Company ID : </b>" + returnID;

        printJobs(returnID);    
    }

    function uploadNote()
    {
        var jobs = document.getElementById("Jobs").options.selectedIndex;
        var currentJob = document.getElementById("Jobs").options[jobs].text;
        var noteText = document.getElementById("Note").value;

        var apicall = "https://api.workflowmax.com/job.api/note?apiKey=14C10292983D48CE86E1AA1FE0F8DDFE&accountKey=8A39F28D022B4366975D6FCDB180C839";

        var note = "<Note><Job>" + currentJob + "</Job><Title>Email note</Title><Text>" + noteText + "</Text></Note>";

        var xhr = new XMLHttpRequest();

        xhr.open('POST', apicall);

        xhr.send(note);
    }

    function uploadTimesheet()
    {
        var jobs = document.getElementById("Jobs").options.selectedIndex;
        var currentJob = document.getElementById("Jobs").options[jobs].text;
        var noteText = document.getElementById("Note").value;

        var task = getTask(currentJob);

        var apicall = "https://api.workflowmax.com/time.api/add?apiKey=14C10292983D48CE86E1AA1FE0F8DDFE&accountKey=8A39F28D022B4366975D6FCDB180C839";
        var tsxml = "<Timesheet><Job>"+ currentJob + "</Job><Task>36465645</Task><Staff>304191</Staff><Date>20151125</Date><Minutes>15</Minutes><Note>" + noteText + "</Note></Timesheet>";

        var xhr = new XMLHttpRequest();

        xhr.open('POST', apicall);

        xhr.send(tsxml);





    }

    function getStaff()
    {

    }

    function getTask(jobID)
    {
        var apicall = "https://api.workflowmax.com/job.api/get/" + jobID + "?apiKey=14C10292983D48CE86E1AA1FE0F8DDFE&accountKey=8A39F28D022B4366975D6FCDB180C839";

        var jobDetails = getXML(apicall);

        var tasklist = jobDetails.getElementsByTagName("Task");

        var thistask = tasklist[0].getElementsByTagName("Name").childNodes[0].nodeValue;

        document.getElementById("Email").innerHTML = "<b>Email : boo! </b> ";

        
    }



    function getXML(list)
    {
        var thisXMLhttp = new XMLHttpRequest();
        thisXMLhttp.open("GET", list, false);
        thisXMLhttp.send();
        var thisXMLDoc = thisXMLhttp.responseXML;

        return thisXMLDoc;
    }

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
    }

    function printJobs(clientID)
    {
        var dropdown = "";
        var jobList = "https://api.workflowmax.com/job.api/client/" + clientID + "?apiKey=14C10292983D48CE86E1AA1FE0F8DDFE&accountKey=8A39F28D022B4366975D6FCDB180C839";

        var jobsXML = getXML(jobList);

        var numJobs = jobsXML.getElementsByTagName("Job");

        for(var i=0; i < numJobs.length; i++)
        {
            var currentJob = numJobs[i].getElementsByTagName("ID")[0].childNodes[0].nodeValue;

            dropdown += "<option value\"" + currentJob + "\">" + currentJob + "</option>";
        }

        document.getElementById("Jobs").innerHTML = dropdown;
       
    }

    function printTasks(jobID)
    {
        var dropdown = "";
    }

})();