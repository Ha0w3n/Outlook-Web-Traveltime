Office.onReady(() => {});

async function createTravel() {

 const item = Office.context.mailbox.item;

 const travelTo = parseInt(document.getElementById("travelTo").value);
 const travelFrom = parseInt(document.getElementById("travelFrom").value);

 const start = new Date(item.start);
 const end = new Date(item.end);

 if (travelTo > 0) {

   const travelStart = new Date(start.getTime() - travelTo*60000);

   createEvent(travelStart, start);
 }

 if (travelFrom > 0) {

   const travelEnd = new Date(end.getTime() + travelFrom*60000);

   createEvent(end, travelEnd);
 }

 document.getElementById("status").innerText = "Travel appointments created";
}

function createEvent(start, end) {

 const event = {
   subject: "Reistijd.",
   start: start,
   end: end,
   /**location: "Travel",*/
   Category: "Travel",
   ShowAs: "Out of Office",
   Reminder: "Off"
 };

 Office.context.mailbox.makeEwsRequestAsync(
   "",
   function(result){}
 );
}