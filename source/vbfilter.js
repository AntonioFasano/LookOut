/*
 **  Antonio Fasano - VBA Filter for Doxygen **
 **              2013 - v. 11                  **
*/




//Open file
var fso = new ActiveXObject("Scripting.FileSystemObject"),
    ForReading = 1, 
    filepath = WScript.Arguments.Unnamed(0),
    file = fso.OpenTextFile(filepath, ForReading, false),
    content=file.ReadAll();
file.Close();


//Headache remover...  ie \r
content=content.replace(/\r/g, '');

//Get Class name
var classname=content.match(/^ *Attribute VB_Name.+/m);
classname=classname[0].replace(/(Attribute VB_Name = ")([^"]+)"/m, "$2" );



//Remove headers
content=content.replace(/^ *VERSION.+/, "" );
content=content.replace(/^ *Attribute.+/gm, "" );
content=content.replace(/^ *BEGIN(.|\n)+?END/m, "" );


//Remove _ break
content=content.replace(/_\n/gm, "" ); 
//WScript.StdOut.Write(content);
//WScript.Quit();


//Property Get (Make Public default)
content=content.replace( /^ *Property +Get/gm, "Public Property Get");
content=content.replace(
    /^ *(Public|Private)( +Property +Get +)(.+)(.|\n)+?^ *End +Property/gm, 
    "$1:$2$3;" );


//Property Set (Make Public default)
content=content.replace(/^ *Property +Let/gm, "Public Property Let");
content=content.replace(
    /^ *(Public|Private)( +Property +Let +)(.+)(.|\n)+?^ *End +Property/gm, 
    "$1:$2$3;" );


//Sub/Function  (Make Public  default)
content=content.replace(/^ *(Sub|Function)/gm, "Public $1");
content=content.replace(
    /^ *(Public|Private) +(Sub|Function) +(.+)(.|\n)+?^ *End +(Sub|Function)/gm, 
    "$1: $2 $3;" );

//Private/Public to lower case
content=content.replace(/^ *P(ublic|rivate):/gm, "p$1:" );


//Embed in class
content=content.replace(/^ *('\/\*\*(.|\n)+?\*\*\/)/m, "$1\nclass " + classname + " {\n\n" );
content=content + '\n}'



//Remove initial ' comments  
//From  '/** to /**
content=content.replace(/^ *'(\/\*\*)/gm, "$1");
//From  '* to *
content=content.replace(/^ *'(\*{1,2}\/)/gm, "$1");
//Leading ' insode c-like commants 
var re, c;
re=/^ *(\/\*\*)(.|\n)+?(\*\/)/gm 
content=content.replace (re, 
  function($0){ return($0.replace(/ *^'/gm, '')); }) 
 
//Initial ' changed to //. 
//Necessary for non-documenting comments set outside procedures 
content=content.replace(/^ *'/gm, "//");





//Taht's it
WScript.StdOut.Write(content);





function echo (estring){
    WScript.Echo(estring);
}

