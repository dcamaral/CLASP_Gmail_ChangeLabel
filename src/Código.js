function ChangeLabel() {
/* --------------------------------------------------------------------------------------------
This function allow to change Threads Labels queried by a range of date.
The user must previously create the label in your gmail and choose the star and finish date.
--------------------------------------------------------------------------------------------*/

// -------------------------------------------------------Step 0 - Setup the parameters of change
var MyLabel = "_Operação/Antigos/@ntigos - 2022" //define your Label Name
var DATE_INI = "2021/1/1" //YYYY/MM/DD
var DATE_FIN = "2022/12/15" //YYYY/MM/DD

//----------------------------------------------======----Step 1 - Collectiong Threads
//var FirstThread = GmailApp.getInboxThreads(0,15); //use this line if you want search all emails without queries.
QUERY_STRING = "label:inbox -is:starred after:" + DATE_INI + " before:" + DATE_FIN;
var FirstThread = GmailApp.search(QUERY_STRING , 0 , 400);

//-------------------------------------------------------Step 2 -  Printing Threads number
var THREADS_NUM = FirstThread.length;
Logger.log ("O número de threads é "+ THREADS_NUM);

//---------------------------------------------------------Step 3 - Creating Array with all messages from Threads
var messages = GmailApp.getMessagesForThreads(FirstThread);

//---------------------------------------------------------Step 4 - Listing messages
  var res = [];
  for (var i = 0; i < messages.length; i++) {
    for (var j = 0; j < messages[i].length; j++) {
      res.push(messages[i][j]);
      console.log("From: " + res[i].getFrom() + "/ Assunto: " + res[i].getSubject() + "/ Data: " + res[i].getDate())

    }
  };
  console.log("número de itens no res = " + res.length);

//---------------------------------------------------------Step 5 - Creating Label objects

  var label = GmailApp.getUserLabelByName(MyLabel);
  console.log ("O rótulo é: " + label.getName());

//---------------------------------------------------------Step 6 - Adding Label to the Threads
for (var i = 0; i < FirstThread.length; i++) {
  FirstThread[i].addLabel(label);
  FirstThread[i].moveToArchive();
    }
 
};
