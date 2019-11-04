
const id=document.form1Name.idName;
const name=document.form1Name.nameName;
const participants=document.form1Name.participantsName;
const length=document.form1Name.lengthName;
const dateS=document.form1Name.dateName;
var hour=document.form1Name.hourName;
const minute=document.form1Name.minuteName;

//creating a file.txt out of the folder
var fso=new ActiveXObject("Scripting.FileSystemObject");
var filename ="./freebusy.txt";
fso.CreateTextFile(filename);

const t= document.getElementById('form1');
t.addEventListener('submit',e=>{

 e.preventDefault();
 //Function to append the data in the file 
 append(id.value,name.value,dateS.value,hour.value,minute.value,length.value,participants.value);

 document.getElementById('form1').reset();
//display the confirmation box
 swal("Meeting Booked !!!", "You can book another meeting after clicking on ok", "success");

});
 
//Calculate the end date & appends the data in the file txt
const append= (idValue,idName,dateS,hour,minute,length,participants)=>{

   var remainderMinute=length % 60 ;
   var quotientHour= ~~(length/60);
   var minuteEnd=30;
   //appending the data
   var  file = fso.GetFile(filename);
   var streamAppend = file.OpenAsTextStream(8);

  if(remainderMinute==30 && minute==30){ 
   quotientHour+=1; minuteEnd="00";
   }
   hourEnd=parseInt(hour)+quotientHour;
   
   streamAppend.WriteLine("Id:"+idValue+"     Name:"+idName);
   streamAppend.WriteLine("Start:"+dateS+" "+hour+":"+minute+" End:"+dateS+" "+hourEnd+":"+minuteEnd+"  Participants:"+participants);
   streamAppend.Close();
}
